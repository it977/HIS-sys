const SUPABASE_URL = "https://erueurkqzmtdefszqons.supabase.co";

const SUPABASE_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImVydWV1cmtxem10ZGVmc3pxb25zIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzMxOTA2OTksImV4cCI6MjA4ODc2NjY5OX0.uShip2ECxvFPfmDLx9-adHGXXTc3cazVdZpSF2tCFUw";

const supabaseClient = supabase.createClient(
    SUPABASE_URL,
    SUPABASE_ANON_KEY
);
console.log("Supabase Client:", supabaseClient);

let currentUser = null;
let masterDataStore = {};
let queueDataStore = [];
let dashRefreshInterval = null;
let reportRefreshInterval = null;
let chartInstances = {};
let html5QrCode = null;
let currentReportData = [];
let currentTriageData = [];
let systemSettings = { hospitalName: "", logoUrl: "", opdHeaderUrl: "", opdFooterUrl: "" };
let servicesDataStore = [];
let locationsDataStore = [];
let allPatientsList = [];
let vaccinesMasterList = [];
let activeOrgsList = [];
let drugsMasterList = [];
let labsMasterList = [];
let currentEMRLabs = [];
let currentEMRDrugs = [];

// ==========================================
// PARTIAL LOADER — fetch & inject HTML files
// ==========================================
async function loadPartials() {
    const views = [
        'dashboard', 'report', 'patients', 'triage', 'opd',
        'appointments', 'vaccines', 'vaccine_master', 'drugs',
        'labs', 'services', 'locations', 'users', 'orgs', 'settings', 'activity_log'
    ];
    const modals = [
        'triage-modal',
        'patient-modal',
        'appointment-qr-modal',
        'org-user-modal',
        'admin-modals',
        'vaccine-modals',
        'emr-modals'
    ];

    try {
        const [navbarHtml, ...rest] = await Promise.all([
            fetch('partials/navbar.html').then(r => r.text()),
            ...views.map(v => fetch(`partials/views/${v}.html`).then(r => r.text())),
            ...modals.map(m => fetch(`partials/modals/${m}.html`).then(r => r.text())),
            fetch('partials/print-areas.html').then(r => r.text())
        ]);

        document.getElementById('partial-navbar').innerHTML = navbarHtml;
        document.getElementById('partial-views').innerHTML =
            rest.slice(0, views.length).join('\n');
        document.getElementById('partial-modals').innerHTML =
            rest.slice(views.length, views.length + modals.length).join('\n');
        document.getElementById('partial-prints').innerHTML =
            rest[rest.length - 1];
    } catch (err) {
        console.error('loadPartials error:', err);
        // Fallback message so the user knows what went wrong
        document.body.innerHTML += `<div style="position:fixed;top:20px;left:50%;transform:translateX(-50%);background:#ef4444;color:#fff;padding:16px 24px;border-radius:8px;font-family:sans-serif;z-index:9999;">
          ⚠️ ບໍ່ສາມາດໂຫຼດ partials ໄດ້.<br>ກະລຸນາ run ຜ່ານ HTTP server (VS Code Live Server ຫຼື npx serve).<br><small>${err.message}</small></div>`;
    }
}

$(document).ready(async function () {
    // Load all HTML partials first, then init the app
    await loadPartials();

    $('#loading').hide();


    // 🌟 ຈັດການ Z-index Modal ທີ່ຊ້ອນກັນ
    $(document).on('show.bs.modal', '.modal', function () {
        var zIndex = 1040 + (10 * $('.modal:visible').length);
        $(this).css('z-index', zIndex);
        setTimeout(function () {
            $('.modal-backdrop').not('.modal-stack').css('z-index', zIndex - 1).addClass('modal-stack');
        }, 0);
    });

    $(document).on('hidden.bs.modal', '.modal', function () {
        if (document.activeElement) document.activeElement.blur();
        if ($('.modal:visible').length > 0) {
            setTimeout(function () { $('body').addClass('modal-open'); }, 100);
        }
    });

    if (typeof $.fn.modal !== 'undefined' && $.fn.modal.Constructor) {
        $.fn.modal.Constructor.prototype.enforceFocus = function () { };
    }

    if (typeof jQuery !== 'undefined' && $.fn.select2) {
        $('#emrService').select2({ dropdownParent: $('#emrModal'), placeholder: "-- ພິມຄົ້ນຫາບໍລິການ --", allowClear: true }).on('change', window.handleServiceSelectionChange);

        $('#p_district').select2({ dropdownParent: $('#patientModal'), placeholder: "-- ຄົ້ນຫາ ແລະ ເລືອກເມືອງ --", allowClear: false }).on('change', function () {
            let dist = $(this).val();
            let loc = locationsDataStore.find(l => l.district === dist);
            $('#p_province').val(loc ? loc.province : '');
        });

        $('#a_patient').select2({ dropdownParent: $('#apptModal'), placeholder: "-- ຄົ້ນຫາຄົນເຈັບ --", allowClear: true }).on('change', function () {
            let d = $(this).select2('data');
            if (d && d.length > 0 && d[0].id) {
                $('#a_target_id').val(d[0].id);
                let txt = d[0].text;
                $('#a_target_name').val(txt.includes(' - ') ? txt.split(' - ')[1] : txt);
            } else {
                $('#a_target_id').val('');
                $('#a_target_name').val('');
            }
        });

        $('#a_org').select2({ dropdownParent: $('#apptModal'), placeholder: "-- ຄົ້ນຫາອົງກອນ --", allowClear: true }).on('change', function () {
            let d = $(this).select2('data');
            if (d && d.length > 0 && d[0].id) {
                $('#a_target_id').val(d[0].id);
                let txt = d[0].text;
                $('#a_target_name').val(txt.includes(' - ') ? txt.split(' - ')[1] : txt);
            } else {
                $('#a_target_id').val('');
                $('#a_target_name').val('');
            }
        });

        $('#pv_patient').select2({ dropdownParent: $('#patientVacModal'), placeholder: "-- ຄົ້ນຫາຄົນເຈັບ --", allowClear: true }).on('change', function () {
            let d = $(this).select2('data');
            if (d && d.length > 0 && d[0].id) {
                $('#pv_patient_id').val(d[0].id);
                let txt = d[0].text;
                $('#pv_patient_name').val(txt.includes(' - ') ? txt.split(' - ')[1] : txt);
            } else {
                $('#pv_patient_id').val('');
                $('#pv_patient_name').val('');
            }
        });

        $('#emrAddDrugSelect').select2({ dropdownParent: $('#emrDrugModal'), placeholder: "-- ເລືອກຢາ --", allowClear: true });
    }

    $('#pv_vaccine, #pv_date').on('change', window.calculateNextVacDate);
    $('#pv_dose').on('input', window.calculateNextVacDate);

    $('[data-widget="pushmenu"]').on('click', function (e) {
        e.preventDefault();
        if ($(window).width() >= 992) {
            $('body').toggleClass('sidebar-collapse');
        } else {
            $('body').toggleClass('sidebar-open');
        }
    });

    $('.content-wrapper').on('click', function () {
        if ($(window).width() < 992 && $('body').hasClass('sidebar-open')) {
            $('body').removeClass('sidebar-open');
        }
    });

    // 🌟 ຈັດການ Dropdown ຂອງ TOP NAVBAR
    $(document).on('click', '.his-dropdown-toggle', function(e) {
        e.preventDefault();
        e.stopPropagation();
        
        let parent = $(this).closest('.his-dropdown');
        let wasOpen = parent.hasClass('open');
        
        // Close ALL dropdowns first
        $('.his-dropdown').removeClass('open');
        
        // Toggle current
        if (!wasOpen) { 
            parent.addClass('open');
            if ($(this).hasClass('his-bell-btn')) { window.renderNotifications(); }
        }
    });

    // Close on click outside OR item click
    $(document).on('click', function(e) {
        if (!$(e.target).closest('.his-dropdown-menu').length && !$(e.target).closest('.his-dropdown-toggle').length) {
            $('.his-dropdown').removeClass('open');
        }
    });

    $(document).on('click', '.his-dropdown-item', function() {
        $('.his-dropdown').removeClass('open');
    });

    // Mobile Hamburger
    $(document).on('click', '.his-hamburger', function(e) {
        e.preventDefault();
        $('#his-nav-items').toggleClass('open');
    });

    // Unified Nav Link Listener (supports both ID-based and Attribute-based navigation)
    $(document).on('click', '.his-nav-link, .his-dropdown-item, .nav-link', function(e) {
        let id = $(this).attr('id');
        if (id && id.startsWith('nav-')) {
            e.preventDefault();
            let view = id.replace('nav-', '');
            window.loadView(view);
        }
    });

    $('#triageForm input[name="v_bp"]').on('input', function () {
        let bp = $(this).val();
        if (bp.includes('/')) {
            let parts = bp.split('/');
            let sys = parseInt(parts[0]);
            let dia = parseInt(parts[1]);
            if (!isNaN(sys) && !isNaN(dia)) {
                if (sys >= 140 || dia >= 90) {
                    $(this).removeClass('border-success text-success border-warning text-dark bg-warning').addClass('border-danger text-danger bg-danger bg-opacity-10 fw-bold');
                } else if (sys <= 90 || dia <= 60) {
                    $(this).removeClass('border-success text-success border-danger text-danger bg-danger').addClass('border-warning text-dark bg-warning bg-opacity-10 fw-bold');
                } else {
                    $(this).removeClass('border-danger text-danger bg-danger border-warning text-dark bg-warning bg-opacity-10').addClass('border-success text-success fw-bold');
                }
            } else {
                $(this).removeClass('border-danger text-danger bg-danger border-warning text-dark bg-warning bg-opacity-10 border-success text-success fw-bold');
            }
        } else {
            $(this).removeClass('border-danger text-danger bg-danger border-warning text-dark bg-warning bg-opacity-10 border-success text-success fw-bold');
        }
    });
});

window.doLogin = async function () {
    let email = $('#loginEmail').val();
    let pass = $('#loginPass').val();

    if (!email || !pass) {
        Swal.fire('ແຈ້ງເຕືອນ', 'ກະລຸນາປ້ອນອີເມວ ແລະ ລະຫັດຜ່ານໃຫ້ຄົບ', 'warning');
        return;
    }

    window.toggleLoading(true);

    try {
        // 1. Authenticate with Supabase Auth (Secure)
        const { data: authData, error: authError } = await supabaseClient.auth.signInWithPassword({
            email: email,
            password: pass
        });

        if (authError || !authData.user) {
            window.toggleLoading(false);
            console.error("Auth Error:", authError);
            Swal.fire('ແຈ້ງເຕືອນ', 'ອີເມວ ຫຼື ລະຫັດຜ່ານບໍ່ຖືກຕ້ອງ', 'error');
            return;
        }

        // 2. Fetch User Profile from Custom Users Table for Roles/Permissions
        const { data, error } = await supabaseClient
            .from('Users')
            .select('*')
            .eq('Email', email)
            .single();

        window.toggleLoading(false);

        if (error || !data || data.Status !== 'active') {
            await supabaseClient.auth.signOut();
            Swal.fire('ແຈ້ງເຕືອນ', 'ບັນຊີຂອງທ່ານບໍ່ພົບໃນລະບົບ ຫຼື ຖືກປິດໃຊ້ງານ', 'error');
            return;
        }

        // ຖ້າຖືກຕ້ອງ ໃຫ້ເກັບຂໍ້ມູນລົງ currentUser
        currentUser = {
            id: data.ID,
            name: data.Name,
            role: data.Role,
            permissions: data.Permissions
        };

        // ສະແດງຂໍ້ຄວາມສຳເລັດ
        Swal.fire({
            title: 'ສຳເລັດ!',
            text: `ຍິນດີຕ້ອນຮັບ, ${currentUser.name}`,
            icon: 'success',
            timer: 2000,
            showConfirmButton: false
        });

        // ເຊື່ອງໜ້າ Login ແລະ ສະແດງໜ້າຫຼັກຊົ່ວຄາວ
        $('#login-section').hide();
        $('#app-content').show();
        $('#sidebarUserName').text(currentUser.name);

        console.log("Login ສຳເລັດ! ຂໍ້ມູນ User:", currentUser);

        window.logAction('Login', `ເຂົ້າສູ່ລະບົບ: ${currentUser.name} (${currentUser.role})`, 'Auth');

        window.initApp();

    } catch (err) {
        window.toggleLoading(false);
        console.error("System Error:", err);
        Swal.fire('ຂໍ້ຜິດພາດ', 'ການເຊື່ອມຕໍ່ມີບັນຫາ, ກະລຸນາກວດສອບອິນເຕີເນັດ!', 'error');
    }
};

window.calculateAgeForm = function () {
    let dobVal = $('#dobInput').val();
    if (dobVal) {
        let dob = new Date(dobVal);
        let age = Math.abs(new Date(Date.now() - dob.getTime()).getUTCFullYear() - 1970);
        $('#ageInput').val(age);
    }
};

window.handleServiceSelectionChange = function () {
    let selectedVals = $('#emrService').val() || [];
    let specs = [];
    let revs = [];
    selectedVals.forEach(val => {
        let srv = servicesDataStore.find(s => s.service === val);
        if (srv) {
            if (srv.specialist && srv.specialist !== "-") specs.push(srv.specialist);
            if (srv.revenue && srv.revenue !== "-") revs.push(srv.revenue);
        }
    });
    $('#emrSpecialist').val([...new Set(specs)].join(', '));
    $('#emrRevenue').val([...new Set(revs)].join(', '));
};

window.logout = async function () {
    window.toggleLoading(true);
    await supabaseClient.auth.signOut();
    window.logAction('Logout', 'ອອກຈາກລະບົບ', 'Auth');
    currentUser = null;
    window.toggleLoading(false);
    $('#app-content').hide();
    $('#login-section').show();
    clearInterval(dashRefreshInterval);
    clearInterval(reportRefreshInterval);
    if (window.closeQRScanner) window.closeQRScanner();
};

window.toggleLoading = function (s) {
    $('#loading').css('display', s ? 'block' : 'none');
};

window.getLocalStr = function (dObj) {
    return dObj.getFullYear() + '-' + String(dObj.getMonth() + 1).padStart(2, '0') + '-' + String(dObj.getDate()).padStart(2, '0');
};

window.initApp = function () {
    try {
        $('#login-section').hide();
        $('#app-content').show();
        $('#sidebarUserName').text(currentUser.name);
        $('.mnu-dashboard, .mnu-report, .mnu-patients, .mnu-triage, .mnu-opd, .mnu-orgs, .mnu-users, .mnu-settings, .mnu-services, .mnu-locations, .mnu-appointments, .mnu-vaccines, .mnu-vaccine_master, .mnu-drugs, .mnu-labs').hide();

        let perms = (currentUser.permissions || "").split(',');
        if (currentUser.role === 'admin' || perms.includes('all')) {
            $('.nav-item, .nav-header, .his-dropdown, .his-nav-link').show();
        } else {
            perms.forEach(p => {
                $('.mnu-' + p.trim()).show();
            });
            if (perms.includes('triage') || perms.includes('opd')) {
                $('.nav-header.mnu-opd').show();
            }
        }

        if (typeof ChartDataLabels !== 'undefined') Chart.register(ChartDataLabels);

        window.toggleLoading(true);
        Promise.all([
            supabaseClient.from('MasterData').select('ID,Category,Value').order('Category').then(({ data }) => {
                const map = {};
                (data || []).forEach(r => { if (!map[r.Category]) map[r.Category] = []; map[r.Category].push({ id: r.ID, value: r.Value }); });
                window.loadMasterDataGlobalCallback(map);
            }),
            supabaseClient.from('Service_Lists').select('*').order('ID').then(({ data }) => {
                servicesDataStore = (data || []).map(r => ({ id: r.ID, revenue: r.Revenue_Group, specialist: r.Mapped_Specialist, service: r.Services_List }));
                let so = '';
                servicesDataStore.forEach(s => { so += `<option value="${s.service}">${s.service}</option>`; });
                if (typeof jQuery !== 'undefined') $('#emrService').empty().append(so).trigger('change');
            }),
            supabaseClient.from('Locations').select('*').order('Province').then(({ data }) => {
                locationsDataStore = (data || []).map(r => ({ id: r.ID, district: r.District, province: r.Province }));
                let o = '<option value="">-- ຄົ້ນຫາ ແລະ ເລືອກເມືອງ --</option>';
                locationsDataStore.forEach(l => o += `<option value="${l.district}">${l.district}</option>`);
                if (typeof jQuery !== 'undefined') $('#p_district').html(o);
            }),
            supabaseClient.from('Organizations').select('*').eq('Status', 'Active').then(({ data }) => {
                let orgOptions = '<option value="">-- ເລືອກອົງກອນ --</option>';
                let seenOpts = new Set();
                (data || []).forEach(org => {
                    if (org.Org_Code && !seenOpts.has(org.Org_Code)) {
                        seenOpts.add(org.Org_Code);
                        orgOptions += `<option value="${org.Org_Code}">${org.Org_Code} - ${org.Org_Name}</option>`;
                    }
                });
                if (typeof jQuery !== 'undefined') {
                    $('#p_org_id').html(orgOptions).select2({ dropdownParent: $('#patientModal'), placeholder: "-- ເລືອກອົງກອນ --", allowClear: true });
                }
            }),
            new Promise((resolve) => { window.preloadDropdownDataCallback(resolve); })
        ]).then(() => {
            window.checkAlerts();
            window.toggleLoading(false);
            if (currentUser.role === 'admin' || perms.includes('dashboard') || perms.includes('all')) window.loadView('dashboard');
            else if (perms.includes('patients')) window.loadView('patients');
            else if (perms.includes('triage')) window.loadView('triage');
            else if (perms.includes('opd')) window.loadView('opd');
        });
    } catch (e) {
        console.error("InitApp Error:", e);
        Swal.fire('ແຈ້ງເຕືອນ', 'ເກີດຂໍ້ຜິດພາດໃນການໂຫຼດໜ້າຈໍ.', 'error');
    }
};

window.loadView = function (v) {
    if (typeof bootstrap !== 'undefined') {
        $('.modal.show').each(function () {
            let bsModal = bootstrap.Modal.getInstance(this);
            if (bsModal) bsModal.hide();
        });
    }
    $('.modal-backdrop').remove();
    $('body').removeClass('modal-open').css({ overflow: '', paddingRight: '' });

    // Handle Active States
    $('.nav-link, .his-nav-link, .his-dropdown-item, .his-dropdown-toggle').removeClass('active');
    
    var navEl = $('#nav-' + v);
    if (navEl.length) {
        navEl.addClass('active');
        let parentDropdown = navEl.closest('.his-dropdown');
        if (parentDropdown.length) {
            parentDropdown.find('.his-dropdown-toggle').addClass('active');
        }
    }

    // Switch Views
    let views = ['dashboard', 'report', 'patients', 'settings', 'orgs', 'triage', 'opd', 'users', 'services', 'locations', 'appointments', 'vaccines', 'vaccine_master', 'drugs', 'labs', 'activity_log'];
    views.forEach(n => {
        if (n === v) $('#view-' + n).show();
        else $('#view-' + n).hide();
    });

    // Reset intervals
    if (dashRefreshInterval) clearInterval(dashRefreshInterval);
    if (reportRefreshInterval) clearInterval(reportRefreshInterval);

    // Load Data
    if (v === 'patients') window.initPatientTable();
    if (v === 'orgs') window.loadOrgs();
    if (v === 'triage') window.loadTriageQueue();
    if (v === 'opd') window.loadQueue();
    if (v === 'users') window.loadUsers();
    if (v === 'services') window.loadServicesMasterView();
    if (v === 'locations') window.loadLocationsMasterView();
    if (v === 'appointments') window.loadAppointments();
    if (v === 'vaccines') window.loadPatientVaccines();
    if (v === 'vaccine_master') window.loadVaccineMaster();
    if (v === 'drugs') window.loadDrugsMaster();
    if (v === 'labs') window.loadLabsMaster();
    if (v === 'activity_log') window.loadActivityLog();

    if (v === 'settings') {
        supabaseClient.from('Settings').select('Key,Value').then(({ data }) => {
            let s = { hospitalName: 'HIS HOSPITAL', logoUrl: '', opdHeaderUrl: '', opdFooterUrl: '' };
            (data || []).forEach(r => {
                if (r.Key === 'HospitalName') s.hospitalName = r.Value || s.hospitalName;
                if (r.Key === 'LogoUrl') s.logoUrl = r.Value || '';
                if (r.Key === 'OpdHeaderUrl') s.opdHeaderUrl = r.Value || '';
                if (r.Key === 'OpdFooterUrl') s.opdFooterUrl = r.Value || '';
            });
            systemSettings = s;
            $('#setHospitalName').val(s.hospitalName);
            $('#setLogoUrl').val(s.logoUrl);
            $('#setOpdHeaderUrl').val(s.opdHeaderUrl);
            $('#setOpdFooterUrl').val(s.opdFooterUrl);
            window.loadMasterList();
        });
    }

    if (!systemSettings.hospitalName) {
        supabaseClient.from('Settings').select('Key,Value').then(({ data }) => {
            (data || []).forEach(r => { if (r.Key === 'HospitalName') systemSettings.hospitalName = r.Value; });
            window.setBrandName(systemSettings.hospitalName);
        });
    }

    if (v === 'dashboard') {
        window.setDashRange('today');
        dashRefreshInterval = setInterval(() => { window.fetchDashboardData(); window.checkAlerts(); }, 120000);
    }
    if (v === 'report') {
        window.setReportRange('today');
        reportRefreshInterval = setInterval(() => { window.fetchReportData(); window.checkAlerts(); }, 120000);
    }

    // Close menus
    $('.his-dropdown').removeClass('open');
    $('#his-nav-items').removeClass('open');
};

window.executePrint = function (containerId) {
    var targetContainer = document.getElementById(containerId);
    if (!targetContainer) return;

    // 1. ເຊື່ອງ Wrapper ຫຼັກຂອງລະບົບທັງໝົດ (Sidebar, Header, Main Content)
    var appWrapper = document.querySelector('.wrapper');
    if (appWrapper) appWrapper.style.display = 'none';

    // 2. ເຊື່ອງ Container Print ໂຕອື່ນໆ ທີ່ບໍ່ກ່ຽວຂ້ອງ
    document.querySelectorAll('.print-container').forEach(function (el) {
        el.style.display = 'none';
        el.classList.remove('print-active');
    });

    // 3. ເປີດສະແດງສະເພາະ Container ທີ່ຕ້ອງການພິມ
    targetContainer.style.display = 'block';
    targetContainer.classList.add('print-active');

    // 4. ກວດສອບຮູບພາບໃຫ້ໂຫຼດສຳເລັດກ່ອນພິມ
    requestAnimationFrame(function () {
        var images = Array.from(targetContainer.querySelectorAll('img')).filter(function (img) {
            if (!img.src || img.style.display === 'none') return false;
            var fullSrc = img.getAttribute('src');
            return fullSrc && fullSrc !== '';
        });

        function doPrintAction() {
            setTimeout(function () {
                window.print();

                // 5. ຫຼັງຈາກພິມແລ້ວ ຄືນຄ່າທຸກຢ່າງໃຫ້ເປັນປົກກະຕິ
                setTimeout(function () {
                    targetContainer.classList.remove('print-active');
                    targetContainer.style.display = 'none';
                    if (appWrapper) appWrapper.style.display = 'block'; // ເປີດລະບົບຄືນ
                }, 500);
            }, 500);
        }

        if (images.length === 0) {
            doPrintAction();
            return;
        }

        var loaded = 0;
        images.forEach(function (img) {
            if (img.complete && img.naturalHeight !== 0) {
                loaded++;
                if (loaded === images.length) doPrintAction();
            } else {
                img.onload = img.onerror = function () {
                    loaded++;
                    if (loaded === images.length) doPrintAction();
                };
            }
        });
    });
};

window.checkAlerts = async function () {
    const today = new Date().toISOString().split('T')[0];
    const futureDate = new Date();
    futureDate.setDate(futureDate.getDate() + 14);
    const futureDateStr = futureDate.toISOString().split('T')[0];
    const { data: appts } = await supabaseClient
        .from('Appointments')
        .select('Appt_ID,Patient_Name,Appt_Date,Appt_Time,Type,Status')
        .eq('Status', 'Pending');
    const alerts = [];
    (appts || []).forEach(r => {
        if (!r.Appt_Date) return;
        const aDate = r.Appt_Date.split('T')[0];
        const diffDays = Math.round((new Date(aDate) - new Date(today)) / 86400000);
        if (diffDays < 0 || (diffDays >= 0 && diffDays <= 14)) {
            alerts.push({ id: r.Appt_ID, patientName: r.Patient_Name, date: aDate, time: r.Appt_Time, type: r.Type, isOverdue: diffDays < 0, daysOut: diffDays });
        }
    });
    alerts.sort((a, b) => a.daysOut - b.daysOut);
    (() => {
        if (!alerts) return;
        let count = alerts.length;
        let badge = $('#bell-count');
        let header = $('#bell-header');
        let list = $('#bell-list');

        if (count > 0) {
            badge.text(count).show();
            header.html(`<span class="text-danger"><i class="fas fa-exclamation-circle"></i> ມີນັດໝາຍຕ້ອງຕິດຕາມ ${count} ລາຍການ</span>`);
            let html = '';
            alerts.forEach(a => {
                let dateColor = a.isOverdue ? "text-danger fw-bold" : (a.daysOut === 0 ? "text-info fw-bold" : "text-warning");
                let label = a.isOverdue ? "ກາຍກຳນົດແລ້ວ!" : (a.daysOut === 0 ? "ມື້ນີ້!" : (a.daysOut === 1 ? "ມື້ອື່ນ" : `ອີກ ${a.daysOut} ວັນ`));
                let iconBg = a.isOverdue ? "bg-danger text-white" : "bg-light text-dark";
                let textType = a.type === 'Vaccine' ? '<span class="text-success">ວັກຊີນ</span>' : 'ທົ່ວໄປ';

                html += `<a href="#" class="his-dropdown-item py-2 border-bottom border-secondary border-opacity-25" onclick="window.loadView('appointments'); return false;">
                            <div class="d-flex align-items-center w-100">
                                <div class="me-3">
                                    <div class="${iconBg} rounded-circle d-flex align-items-center justify-content-center" style="width:32px;height:32px; min-width:32px; font-size: 11px;">
                                        <i class="fas ${a.type === 'Vaccine' ? 'fa-syringe' : 'fa-calendar-check'}"></i>
                                    </div>
                                </div>
                                <div class="flex-grow-1 overflow-hidden" style="line-height: 1.2;">
                                    <h6 class="m-0 fw-bold text-white mb-1" style="font-size:12.5px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">${a.patientName}</h6>
                                    <div class="text-info opacity-100 mb-1" style="font-size:10.5px; font-weight: 500;">
                                        <i class="fas fa-tag me-1" style="font-size: 9px;"></i>${a.time} - ${textType}
                                    </div>
                                    <div class="${dateColor} opacity-100" style="font-size:10.5px; font-weight: 600;">
                                        <i class="far fa-clock me-1" style="font-size: 9px;"></i>${label} (${a.date})
                                    </div>
                                </div>
                            </div>
                         </a>`;
            });
            list.html(html);
        } else {
            badge.hide();
            header.text('ບໍ່ມີການແຈ້ງເຕືອນນັດໝາຍ');
            list.html('<div class="text-center py-3 text-muted small">ຍັງບໍ່ມີນັດໝາຍໃນໄລຍະ 14 ວັນນີ້</div>');
        }
    })();
};

window.renderNotifications = function () { window.checkAlerts(); };

window.preloadDropdownDataCallback = function (resolve) {
    let promises = [
        supabaseClient.from('Patients').select('Patient_ID,First_Name,Last_Name').order('Patient_ID', { ascending: false }).then(({ data }) => {
            allPatientsList = (data || []).map(p => ({ id: p.Patient_ID, fullname: `${p.First_Name || ''} ${p.Last_Name || ''}`.trim() }));
            let opts = '<option value="">-- ຄົ້ນຫາ ແລະ ເລືອກຄົນເຈັບ --</option>';
            allPatientsList.forEach(p => { opts += `<option value="${p.id}">${p.id} - ${p.fullname}</option>`; });
            if (typeof jQuery !== 'undefined') { $('#a_patient').html(opts).trigger('change'); $('#pv_patient').html(opts).trigger('change'); }
        }),
        supabaseClient.from('Organizations').select('Org_Code,Org_Name').eq('Status', 'Active').then(({ data }) => {
            activeOrgsList = [];
            let seenOpts = new Set();
            (data || []).forEach(r => {
                if (r.Org_Code && !seenOpts.has(r.Org_Code)) {
                    seenOpts.add(r.Org_Code);
                    activeOrgsList.push({ id: r.Org_Code, name: r.Org_Name });
                }
            });
            let opts = '<option value="">-- ຄົ້ນຫາ ແລະ ເລືອກອົງກອນ --</option>';
            activeOrgsList.forEach(o => { opts += `<option value="${o.id}">${o.id} - ${o.name}</option>`; });
            if (typeof jQuery !== 'undefined') { $('#a_org').html(opts).trigger('change'); }
        }),
        supabaseClient.from('Drugs_Master').select('Drug_ID,Drug_Name,Description').order('Drug_Name').then(({ data }) => {
            drugsMasterList = (data || []).map(r => ({ id: r.Drug_ID, name: r.Drug_Name, desc: r.Description || '' }));
            let o = '<option value="">-- ເລືອກຢາ --</option>';
            drugsMasterList.forEach(d => { o += `<option value="${d.name}">${d.name}${d.desc ? ' (' + d.desc + ')' : ''}</option>`; });
            if (typeof jQuery !== 'undefined') $('#emrAddDrugSelect').html(o).trigger('change');
        }),
        supabaseClient.from('Labs_Master').select('Lab_ID,Lab_Name,Description').order('Lab_Name').then(({ data }) => {
            labsMasterList = (data || []).map(r => ({ id: r.Lab_ID, name: r.Lab_Name, desc: r.Description || '' }));
            let h = '';
            labsMasterList.forEach((l, i) => {
                h += `<div class="form-check mb-2">
                        <input class="form-check-input lab-checkbox" type="checkbox" value="${l.name}" id="chkLab${i}">
                        <label class="form-check-label fw-bold text-primary" for="chkLab${i}">${l.name}</label>
                        <span class="text-muted small ms-1">- ${l.desc}</span>
                      </div>`;
            });
            if (document.getElementById('labCheckboxContainer')) document.getElementById('labCheckboxContainer').innerHTML = h;
        })
    ];
    Promise.all(promises).then(() => resolve());
}

window.preloadDropdownData = function () { window.preloadDropdownDataCallback(function () { }); };

window.setDashRange = function (type) {
    $('.btn-group .btn').removeClass('active btn-primary').addClass('btn-outline-primary');
    $('#btnDash' + type.charAt(0).toUpperCase() + type.slice(1)).addClass('active btn-primary').removeClass('btn-outline-primary');

    let start = new Date();
    let end = new Date();
    if (type === 'week') {
        let day = start.getDay() || 7;
        if (day !== 1) start.setDate(start.getDate() - (day - 1));
    } else if (type === 'month') {
        start.setDate(1);
    } else if (type === 'year') {
        start.setMonth(0, 1);
    }
    $('#dashStartDate').val(window.getLocalStr(start));
    $('#dashEndDate').val(window.getLocalStr(end));
    window.fetchDashboardData();
};

window.fetchDashboardData = async function () {
    let sDate = $('#dashStartDate').val();
    let eDate = $('#dashEndDate').val();
    if (!sDate || !eDate) return;

    let d = new Date();
    $('#dashRefreshTime').text(`ອັບເດດລ່າສຸດ: ${d.getHours().toString().padStart(2, '0')}:${d.getMinutes().toString().padStart(2, '0')}`);
    $('#dash-total, #dash-new, #dash-old, #dash-inscorp').html('<i class="fas fa-spinner fa-spin"></i>');

    try {
        const { data, error } = await supabaseClient
            .from('Visits')
            .select('*')
            .gte('Date', sDate + 'T00:00:00Z')
            .lte('Date', eDate + 'T23:59:59Z');

        if (error) {
            console.error('Dashboard Error:', error);
            return;
        }

        // Calculate isNew for each visit
        const pIds = [...new Set(data.map(v => v.Patient_ID).filter(id => !!id))];
        let firstVisitMap = {};
        if (pIds.length > 0) {
            const { data: allVisits, error: avError } = await supabaseClient.from('Visits')
                .select('Visit_ID, Patient_ID, Date')
                .in('Patient_ID', pIds)
                .order('Date', { ascending: true });
            
            if (!avError && allVisits) {
                allVisits.forEach(v => {
                    if (!firstVisitMap[v.Patient_ID]) {
                        firstVisitMap[v.Patient_ID] = v.Visit_ID;
                    }
                });
            }
        }

        const visitsWithIsNew = data.map(v => ({
            ...v,
            isNew: firstVisitMap[v.Patient_ID] === v.Visit_ID
        }));

        window.renderDashboardCharts(visitsWithIsNew);

    } catch (err) {
        console.error('System Error:', err);
    }
};

window.createChart = function (ctxId, type, labels, data, colors, isHorizontal = false) {
    if (chartInstances[ctxId]) chartInstances[ctxId].destroy();
    const el = document.getElementById(ctxId);
    if (!el) return;
    const ctx = el.getContext('2d');
    let options = {
        responsive: true, maintainAspectRatio: false,
        plugins: {
            legend: { position: 'bottom' },
            datalabels: {
                color: type === 'bar' ? '#334155' : '#ffffff',
                font: { weight: 'bold', size: 14 },
                anchor: type === 'bar' ? 'end' : 'center',
                align: type === 'bar' ? 'top' : 'center',
                formatter: (value) => { return value > 0 ? value : ''; }
            }
        }
    };
    if (type === 'bar') {
        options.plugins.legend.display = false;
        options.indexAxis = isHorizontal ? 'y' : 'x';
        options.scales = isHorizontal ?
            { x: { beginAtZero: true, ticks: { precision: 0 } }, y: { grid: { display: false } } } :
            { y: { beginAtZero: true, ticks: { precision: 0 } }, x: { grid: { display: false } } };
    }
    chartInstances[ctxId] = new Chart(ctx, { type: type, data: { labels: labels, datasets: [{ data: data, backgroundColor: colors, borderWidth: 1 }] }, options: options });
};

window.renderDashboardCharts = function (visits) {
    if (!visits) return;

    // 1. Stats
    let total = visits.length;
    let newPatients = 0;
    let oldPatients = 0;
    
    // We need to fetch all-time visits for these patients to see who is truly "New"
    // However, for dashboard, if we have the isNew flag pre-calculated it would be better.
    // Let's rely on a more efficient way or assume visits data passed here might have it.
    visits.forEach(v => {
        if (v.isNew) newPatients++;
        else oldPatients++;
    });

    let insCorp = visits.filter(v => v.Revenue_Group && v.Revenue_Group !== 'General Cash').length;

    $('#dash-total').text(total);
    $('#dash-new').text(newPatients);
    $('#dash-old').text(oldPatients);
    $('#dash-inscorp').text(insCorp);

    // 2. Helper
    const getTopN = (map, n = 10) => {
        let keys = Object.keys(map).sort((a, b) => map[b] - map[a]).slice(0, n);
        return { labels: keys, data: keys.map(k => map[k]) };
    };

    // 3. Process data
    let services = {}, revenue = {}, specialist = {}, gender = {}, deptType = {}, site = {}, opdGender = {}, timeSlot = {}, ageGroup = {}, province = {};

    visits.forEach(v => {
        let p = v.Patients || {};
        if (v.Services_List) v.Services_List.split(',').forEach(s => { let n = s.trim(); services[n] = (services[n] || 0) + 1; });
        if (v.Revenue_Group) revenue[v.Revenue_Group] = (revenue[v.Revenue_Group] || 0) + 1;
        if (v.Mapped_Specialist) specialist[v.Mapped_Specialist] = (specialist[v.Mapped_Specialist] || 0) + 1;
        
        let g = p.Gender || "ບໍ່ລະບຸ";
        gender[g] = (gender[g] || 0) + 1;
        if (v.Visit_Type === 'OPD') opdGender[g] = (opdGender[g] || 0) + 1;
        
        let age = parseInt(p.Age);
        if (!isNaN(age)) {
            let grp = age < 15 ? "0-14" : (age < 35 ? "15-34" : (age < 60 ? "35-59" : "60+"));
            ageGroup[grp] = (ageGroup[grp] || 0) + 1;
        }
        if (p.Province) province[p.Province] = (province[p.Province] || 0) + 1;
        
        if (v.Visit_Type) deptType[v.Visit_Type] = (deptType[v.Visit_Type] || 0) + 1;
        if (v.Site) site[v.Site] = (site[v.Site] || 0) + 1;
        
        if (v.Date) {
            let h = new Date(v.Date).getHours();
            let slot = h < 12 ? "08:00 - 12:00" : (h < 16 ? "12:00 - 16:00" : "16:00 - 20:00");
            timeSlot[slot] = (timeSlot[slot] || 0) + 1;
        }
    });

    const cB = '#0ea5e9', cP = '#f43f5e', cG = '#10b981', cY = '#f59e0b', cPu = '#8b5cf6', cO = '#fdba74';
    let topSvc = getTopN(services);
    let topProv = getTopN(province, 5);
    
    window.createChart('chartTopServices', 'bar', topSvc.labels, topSvc.data, [cB], false);
    window.createChart('chartRevenue', 'pie', Object.keys(revenue), Object.values(revenue), [cG, cB, cY, cP, cPu]);
    window.createChart('chartSpecialist', 'doughnut', Object.keys(specialist), Object.values(specialist), [cPu, cB, cO, cG, cP]);
    window.createChart('chartGender', 'doughnut', Object.keys(gender), Object.values(gender), [cB, cP]);
    window.createChart('chartDept', 'pie', Object.keys(deptType), Object.values(deptType), [cG, cY]);
    window.createChart('chartSite', 'pie', Object.keys(site), Object.values(site), [cPu, '#64748b']);
    window.createChart('chartOpdGender', 'doughnut', Object.keys(opdGender), Object.values(opdGender), [cB, cP]);
    window.createChart('chartTime', 'bar', Object.keys(timeSlot).sort(), Object.values(timeSlot), [cB, cY, '#3b82f6']);
    window.createChart('chartAge', 'bar', ["0-14", "15-34", "35-59", "60+"], ["0-14", "15-34", "35-59", "60+"].map(k => ageGroup[k] || 0), [cG, cB, cY, cP]);
    window.createChart('chartProvince', 'bar', topProv.labels, topProv.data, [cB], true);
};

window.exportDashboardPDF = function () {
    const element = document.getElementById('dashboardPrintArea');
    const opt = {
        margin: 0.5, filename: 'HIS_Dashboard.pdf', image: { type: 'jpeg', quality: 0.98 },
        html2canvas: { scale: 2 }, jsPDF: { unit: 'in', format: 'a4', orientation: 'landscape' }
    };
    Swal.fire({ title: 'ກຳລັງສ້າງ PDF...', didOpen: () => { Swal.showLoading() } });
    html2pdf().set(opt).from(element).save().then(() => Swal.close());
};

window.setReportRange = function (type) {
    $('.btn-group .btn').removeClass('active btn-primary').addClass('btn-outline-primary');
    $('#btnRep' + type.charAt(0).toUpperCase() + type.slice(1)).addClass('active btn-primary').removeClass('btn-outline-primary');

    let start = new Date();
    let end = new Date();
    if (type === 'week') {
        let day = start.getDay() || 7;
        if (day !== 1) start.setDate(start.getDate() - (day - 1));
    } else if (type === 'month') {
        start.setDate(1);
    } else if (type === 'year') {
        start.setMonth(0, 1);
    }
    $('#repStartDate').val(window.getLocalStr(start));
    $('#repEndDate').val(window.getLocalStr(end));
    window.fetchReportData();
};

window.fetchReportData = function () {
    let sDate = $('#repStartDate').val();
    let eDate = $('#repEndDate').val();
    if (!sDate || !eDate) return;
    let d = new Date();
    $('#repRefreshTime').text(`ອັບເດດ: ${d.getHours().toString().padStart(2, '0')}:${d.getMinutes().toString().padStart(2, '0')}`);
    if ($.fn.DataTable.isDataTable('#reportTable')) { $('#reportTable').DataTable().destroy(); }
    $('#reportTable tbody').html('<tr><td colspan="9" class="text-center py-4"><div class="spinner-border text-primary spinner-border-sm"></div> ກຳລັງໂຫຼດຂໍ້ມູນ...</td></tr>');
    window._fetchReportData(sDate, eDate);
};

window._fetchReportData = async function (sDate, eDate) {
    try {
        // 1. Fetch Visits
        const { data: visits, error: vError } = await supabaseClient.from('Visits')
            .select('*')
            .gte('Date', sDate + 'T00:00:00Z')
            .lte('Date', eDate + 'T23:59:59Z')
            .order('Date', { ascending: false });

        if (vError) throw vError;
        if (!visits || visits.length === 0) return window.renderReportPage([]);

        // 2. Fetch unique Patients involved
        const pIds = [...new Set(visits.map(v => v.Patient_ID).filter(id => !!id))];
        let pMap = {};
        
        if (pIds.length > 0) {
            const { data: patients, error: pError } = await supabaseClient.from('Patients')
                .select('Patient_ID, First_Name, Last_Name, Gender, Age')
                .in('Patient_ID', pIds);
            
            if (!pError && patients) {
                patients.forEach(p => pMap[p.Patient_ID] = p);
            }
        }

        // 3. Fetch all-time visits for these patients to determine isNew
        let firstVisitMap = {};
        if (pIds.length > 0) {
            const { data: allVisits, error: avError } = await supabaseClient.from('Visits')
                .select('Visit_ID, Patient_ID, Date')
                .in('Patient_ID', pIds)
                .order('Date', { ascending: true });
            
            if (!avError && allVisits) {
                allVisits.forEach(v => {
                    if (!firstVisitMap[v.Patient_ID]) {
                        firstVisitMap[v.Patient_ID] = v.Visit_ID;
                    }
                });
            }
        }

        // 4. Merge data
        const processed = visits.map(r => {
            let dObj = new Date(r.Date);
            let p = pMap[r.Patient_ID];
            return {
                ...r, // Keep original record for "View" detail
                date: dObj.toLocaleDateString('en-GB'),
                time: dObj.toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit' }),
                id: r.Patient_ID,
                name: r.Patient_Name || (p ? `${p.First_Name} ${p.Last_Name}` : '-'),
                gender: p?.Gender || '-',
                age: p?.Age || '-',
                type: r.Type || 'OPD',
                status: (firstVisitMap[r.Patient_ID] === r.Visit_ID) ? 'ໃໝ່' : 'ເກົ່າ',
                category: r.Category || 'GN'
            };
        });
        window.renderReportPage(processed);
    } catch (err) {
        console.error('Report Fetch Error:', err);
        Swal.fire('Error', 'ບໍ່ສາມາດໂຫຼດຂໍ້ມູນລາຍງານໄດ້: ' + err.message, 'error');
        window.renderReportPage([]);
    }
};

window.renderReportPage = function (res) {
    currentReportData = res || [];
    if ($.fn.DataTable.isDataTable('#reportTable')) $('#reportTable').DataTable().destroy();
    if (!res || res.length === 0) {
        $('#reportTable tbody').empty();
        $('#reportTable').DataTable({ responsive: true, language: { emptyTable: "ບໍ່ມີຂໍ້ມູນ", search: "ຄົ້ນຫາ:" } });
        return;
    }

    let h = '';
    res.forEach((r, i) => {
        let bs = r.status === 'ໃໝ່' ? '<span class="badge bg-success">ໃໝ່</span>' : '<span class="badge bg-secondary">ເກົ່າ</span>';
        let bt = r.type === 'IPD' ? '<span class="badge bg-warning text-dark">IPD</span>' : '<span class="badge bg-light text-dark border">OPD</span>';
        let bc = r.category === 'GN' ? 'ທົ່ວໄປ (GN)' : '<span class="text-danger fw-bold">ປະກັນ/ອົງກອນ</span>';
        let acts = `<button class="btn btn-sm btn-outline-primary fw-bold shadow-sm" onclick="window.viewReportDetail(${i})"><i class="fas fa-eye me-1"></i> View</button>`;
        h += `<tr>
                <td>${r.date}</td>
                <td>${r.time}</td>
                <td class="text-primary fw-bold">${r.id}</td>
                <td class="fw-bold">${r.name}</td>
                <td>${r.gender}</td>
                <td>${r.age}</td>
                <td>${bt}</td>
                <td>${bs}</td>
                <td>${bc}</td>
                <td class="text-center">${acts}</td>
              </tr>`;
    });
    $('#reportTable tbody').html(h);
    $('#reportTable').DataTable({ responsive: true, pageLength: 15, order: [[0, "desc"], [1, "desc"]], language: { search: "ຄົ້ນຫາ:", lengthMenu: "ສະແດງ _MENU_", info: "ສະແດງ _START_ ຫາ _END_ ຈາກ _TOTAL_ ລາຍການ", paginate: { previous: "ກ່ອນໜ້າ", next: "ຕໍ່ໄປ" } } });
};

window.exportReportExcel = function () {
    if (!currentReportData || currentReportData.length === 0) return Swal.fire('ແຈ້ງເຕືອນ', 'ບໍ່ມີຂໍ້ມູນສຳລັບ Export', 'warning');
    const exportArr = currentReportData.map(r => ({ "ວັນທີ": r.date, "ເວລາ": r.time, "ລະຫັດ": r.id, "ຊື່ ແລະ ນາມສະກຸນ": r.name, "ເພດ": r.gender, "ອາຍຸ": r.age, "ປະເພດ": r.type, "ສະຖານະ": r.status, "ໝວດໝູ່": r.category }));
    const ws = XLSX.utils.json_to_sheet(exportArr);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Report");
    XLSX.writeFile(wb, "HIS_Summary_Report.xlsx");
};

window.viewReportDetail = function (i) {
    const r = currentReportData[i];
    if (!r) return;

    let labList = "ບໍ່ມີການສັ່ງກວດ", drugList = "ບໍ່ມີການສັ່ງຢາ";
    try {
        let labs = r.Lab_Orders_JSON ? JSON.parse(r.Lab_Orders_JSON) : [];
        if (labs.length > 0) {
            labList = "<ul class='mb-0 text-start ps-3'>";
            labs.forEach(l => labList += `<li class="mb-1">${l.name}</li>`);
            labList += "</ul>";
        }
    } catch (e) { }

    try {
        let drugs = r.Prescription_JSON ? JSON.parse(r.Prescription_JSON) : [];
        if (drugs.length > 0) {
            drugList = "<ul class='mb-0 text-start ps-3'>";
            drugs.forEach(d => drugList += `<li class="mb-1"><b>${d.name}</b>: <span class="badge bg-secondary">${d.qty}</span> <div class="small text-muted">${d.usage}</div></li>`);
            drugList += "</ul>";
        }
    } catch (e) { }

    let html = `
        <div class="p-4">
            <!-- Patient Profile Header -->
            <div class="d-flex align-items-center mb-4 p-3 rounded" style="background: linear-gradient(135deg, #0ea5e9, #6366f1); color: white; box-shadow: 0 4px 15px rgba(14, 165, 233, 0.2);">
                <div class="rounded-circle bg-white text-primary d-flex align-items-center justify-content-center me-3" style="width: 55px; height: 55px; font-size: 24px;">
                    <i class="fas fa-user-circle"></i>
                </div>
                <div>
                    <h4 class="m-0 fw-bold">${r.name}</h4>
                    <span class="badge bg-light text-primary mt-1 shadow-sm">${r.id}</span>
                </div>
                <div class="ms-auto text-end">
                    <div class="small opacity-75"><i class="far fa-calendar-alt me-1"></i> ${r.date}</div>
                    <div class="fw-bold"><i class="far fa-clock me-1"></i> ${r.time}</div>
                </div>
            </div>

            <div class="row g-3">
                <!-- Vitals & Basic Info -->
                <div class="col-md-5">
                    <div class="card border-0 shadow-sm mb-3" style="border-radius: 12px; height: 100%;">
                        <div class="card-header bg-white border-0 pt-3 pb-0">
                            <h6 class="fw-bold text-muted small text-uppercase mb-0"><i class="fas fa-heartbeat text-danger me-2"></i>ຂໍ້ມູນຊີວະສັນຍານ</h6>
                        </div>
                        <div class="card-body">
                            <div class="d-flex flex-wrap gap-2 mb-3">
                                <div class="p-2 border rounded text-center" style="min-width: 70px; background: #fff;">
                                    <div class="small text-muted">BP</div>
                                    <div class="fw-bold text-primary">${r.BP || '-'}</div>
                                </div>
                                <div class="p-2 border rounded text-center" style="min-width: 70px; background: #fff;">
                                    <div class="small text-muted">Temp</div>
                                    <div class="fw-bold text-danger">${r.Temp || '-'} °C</div>
                                </div>
                                <div class="p-2 border rounded text-center" style="min-width: 70px; background: #fff;">
                                    <div class="small text-muted">Pulse</div>
                                    <div class="fw-bold text-warning">${r.Pulse || '-'}</div>
                                </div>
                                <div class="p-2 border rounded text-center" style="min-width: 70px; background: #fff;">
                                    <div class="small text-muted">Weight</div>
                                    <div class="fw-bold text-success">${r.Weight || '-'} kg</div>
                                </div>
                            </div>
                            <div class="border-top pt-3">
                                <p class="mb-2"><b>ອາການເບື້ອງຕົ້ນ:</b><br><span class="text-muted">${r.Symptoms || '-'}</span></p>
                                <p class="mb-0"><b>ແພດຜູ້ກວດ:</b><br><span class="text-primary fw-bold"><i class="fas fa-user-md me-1"></i>${r.Doctor_Name || '-'}</span></p>
                            </div>
                        </div>
                    </div>
                </div>

                <!-- Diagnosis & Treatment -->
                <div class="col-md-7">
                    <div class="card border-0 shadow-sm mb-3" style="border-radius: 12px;">
                        <div class="card-body">
                            <div class="mb-3">
                                <h6 class="fw-bold text-danger"><i class="fas fa-stethoscope me-2"></i>ການວິນິດໄສ (Diagnosis)</h6>
                                <div class="p-3 bg-light rounded border-start border-danger border-4 mt-2">
                                    <h5 class="m-0 fw-bold">${r.Diagnosis || 'ຍັງບໍ່ມີຂໍ້ມູນ'}</h5>
                                </div>
                            </div>
                            
                            <div class="row g-2">
                                <div class="col-6">
                                    <div class="p-2 border rounded h-100 bg-white">
                                        <h6 class="fw-bold text-primary small"><i class="fas fa-flask me-2"></i>ລາຍການແລັບ</h6>
                                        <div class="small mt-1">${labList}</div>
                                    </div>
                                </div>
                                <div class="col-6">
                                    <div class="p-2 border rounded h-100 bg-white">
                                        <h6 class="fw-bold text-success small"><i class="fas fa-pills me-2"></i>ລາຍການຢາ</h6>
                                        <div class="small mt-1">${drugList}</div>
                                    </div>
                                </div>
                            </div>

                            <div class="mt-3 p-3 rounded" style="background-color: #fef9c3; border: 1px dashed #f59e0b;">
                                <h6 class="fw-bold text-warning-emphasis small mb-1"><i class="fas fa-info-circle me-2"></i>ຄຳແນະນຳ/ຕິດຕາມ</h6>
                                <div class="small">${r.Advice || r.Follow_Up ? (r.Advice + ' ' + (r.Follow_Up || '')) : 'ບໍ່ມີຄຳແນະນຳເພີ່ມເຕີມ'}</div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    `;
    $('#reportDetailContent').html(html);
    $('#reportDetailModal').modal('show');
};

window.generateNextPatientID = async function () {
    try {
        const { data, error } = await supabaseClient
            .from('Patients')
            .select('Patient_ID');

        if (error) throw error;

        let maxNum = 0;
        if (data && data.length > 0) {
            data.forEach(row => {
                if (row.Patient_ID && row.Patient_ID.startsWith('CN')) {
                    let numPart = row.Patient_ID.replace(/[^0-9]/g, '');
                    let num = parseInt(numPart, 10);
                    if (!isNaN(num) && num > maxNum) {
                        maxNum = num;
                    }
                }
            });
        }
        
        return 'CN' + ('0000000' + (maxNum + 1)).slice(-7);
    } catch (err) {
        console.error("Error generating next CN:", err);
        return 'CN0000001'; // Fallback
    }
};

window.initPatientTable = async function () {
    if ($.fn.DataTable.isDataTable('#patientTable')) {
        $('#patientTable').DataTable().destroy();
    }
    $('#patientTable tbody').html('<tr><td colspan="10" class="text-center py-4"><div class="spinner-border text-primary spinner-border-sm"></div> ກຳລັງໂຫຼດ...</td></tr>');

    try {
        // ດຶງຂໍ້ມູນຈາກ Supabase ແທນ google.script.run
        const { data, error } = await supabaseClient
            .from('Patients')
            .select('*')
            .order('Patient_ID', { ascending: false }); // ລຽງລຳດັບຄົນເຈັບໃໝ່ລ່າສຸດຂຶ້ນກ່ອນ

        if (error) throw error;

        $('#patientTable tbody').empty();

        if (!data || data.length === 0) {
            $('#patientTable').DataTable({ responsive: true, language: { emptyTable: "ຍັງບໍ່ມີຂໍ້ມູນຄົນເຈັບ", search: "ຄົ້ນຫາ:" } });
            return;
        }

        let h = "";
        data.forEach(r => {
            // ໃຊ້ຊື່ Column ຕາມໃນ CSV ຂອງເຈົ້າ (First_Name, Last_Name, ແລະ ອື່ນໆ)
            let fullname = `${r.First_Name || ''} ${r.Last_Name || ''}`.trim();
            let safeName = fullname.replace(/'/g, "\\'").replace(/"/g, "&quot;");

            let acts = `<div class="d-flex gap-1 flex-nowrap justify-content-center">
                          <button class="btn btn-sm btn-warning text-dark shadow-sm fw-bold" onclick="window.sendToTriageFlow('${r.Patient_ID}', '${safeName}')"><i class="fas fa-share me-1"></i> Triage</button>
                          <button class="btn btn-sm btn-primary shadow-sm" title="ແກ້ໄຂ" onclick="window.editPatient('${r.Patient_ID}')"><i class="fas fa-edit"></i></button>
                          <button class="btn btn-sm btn-info text-white shadow-sm" title="ພິມບັດ QR" onclick="window.printQRCard('${r.Patient_ID}')"><i class="fas fa-qrcode"></i></button>
                          <button class="btn btn-sm btn-danger shadow-sm" title="ລຶບ" onclick="window.delPatient('${r.Patient_ID}')"><i class="fas fa-trash"></i></button>
                        </div>`;

            h += `<tr>
                    <td class="text-muted small">${r.Registration_Date || '-'}</td>
                    <td class="text-muted small">${r.Time || '-'}</td>
                    <td class="text-primary fw-bold">${r.Patient_ID || '-'}</td>
                    <td class="fw-bold">${fullname}</td>
                    <td>${r.Gender || '-'}</td>
                    <td>${r.Age || 0} ປີ</td>
                    <td class="text-muted">${r.Phone_Number || '-'}</td>
                    <td class="small text-muted">${r.District || ''} ${r.Province || ''}</td>
                    <td class="text-danger fw-bold small">${r.Drug_Allergy || '-'}</td>
                    <td>${acts}</td>
                  </tr>`;
        });

        $('#patientTable tbody').html(h);
        $('#patientTable').DataTable({
            responsive: true,
            pageLength: 10,
            order: [[0, "desc"], [1, "desc"]],
            language: {
                search: "ຄົ້ນຫາ:",
                lengthMenu: "ສະແດງ _MENU_",
                info: "ສະແດງ _START_ ຫາ _END_ ຈາກ _TOTAL_ ລາຍການ",
                paginate: { previous: "ກ່ອນໜ້າ", next: "ຕໍ່ໄປ" }
            }
        });

    } catch (err) {
        console.error("Error loading patients:", err);
        $('#patientTable tbody').html('<tr><td colspan="10" class="text-center text-danger py-4">ເກີດຂໍ້ຜິດພາດໃນການດຶງຂໍ້ມູນ</td></tr>');
    }
};

window.openNewPatientModal = function () {
    $('#patientForm')[0].reset();
    $('#p_action').val("new");
    $('#p_id').val("");
    $('#disp_p_id').val("ກຳລັງໂຫຼດ...");
    $('#p_org_id').val(null).trigger('change');
    $('#p_org_name, #p_discount_show').val('');
    $('#p_district').val(null).trigger('change');

    window.generateNextPatientID().then(id => {
        $('#disp_p_id').val(id);
    });
    let n = new Date();
    $('#p_date').val(window.getLocalStr(n));
    $('#p_time').val(n.toTimeString().split(' ')[0].substring(0, 5));

    if (document.activeElement) document.activeElement.blur();
    $('#patientModal').modal('show');
};

window.editPatient = async function (id) {
    Swal.fire({ title: 'ກຳລັງດຶງຂໍ້ມູນ...', didOpen: () => Swal.showLoading() });
    const { data, error } = await supabaseClient.from('Patients').select('*').eq('Patient_ID', id).single();
    Swal.close();
    if (error || !data) return;
    const d = {
        id: data.Patient_ID, title: data.Title || '', firstname: data.First_Name || '',
        lastname: data.Last_Name || '', gender: data.Gender || '', dob: data.Date_of_Birth || '',
        age: data.Age || '', nation: data.Nationality || '', job: data.Occupation || '',
        blood: data.Blood_Type || '', phone: data.Phone_Number || '', email: data.Email || '',
        address: data.Address || '', district: data.District || '', province: data.Province || '',
        org_id: data.Organization_ID || '', org_name: data.Name_Org || '',
        ins_company: data.Insurance_Company || '', ins_code: data.Insurance_Code || '',
        ins_name: data.Insured_Person_Name || '', allergy: data.Drug_Allergy || 'ບໍ່ມີ',
        disease: data.Underlying_Disease || 'ບໍ່ມີ', emer_name: data.Emergency_Name || '',
        emer_contact: data.Emergency_Contact || '', emer_relation: data.Emergency_Relation || '',
        channel: data.Channel || '', date: data.Registration_Date || '', time: data.Time || '', shift: data.Shift || ''
    };
    $('#patientModalTitle').html(`<i class="fas fa-user-edit text-warning me-2"></i>ແກ້ໄຂຂໍ້ມູນ: <span class="text-primary">${d.id}</span>`);
    $('#p_action').val('edit');
    $('#p_id').val(d.id);
    $('#disp_p_id').val(d.id);
    for (const [k, v] of Object.entries(d)) {
        if (k === 'org_id') continue;
        let el = document.getElementById('p_' + k) || document.getElementsByName('p_' + k)[0];
        if (el) el.value = v;
    }
    $('#p_org_id').val(d.org_id).trigger('change');
    $('#p_district').val(d.district).trigger('change');
    $('#p_province').val(d.province || '');
    window.fetchOrg();
    if (document.activeElement) document.activeElement.blur();
    $('#patientModal').modal('show');
};

window.submitPatientForm = async function (e) {
    if (e) e.preventDefault();
    Swal.fire({ title: 'ກຳລັງບັນທຶກ...', didOpen: () => Swal.showLoading() });
    const fd = {};
    new FormData($('#patientForm')[0]).forEach((v, k) => fd[k] = v);
    const isEdit = fd.p_action === 'edit' && fd.p_id;
    const age = parseInt(fd.p_age) || 0;
    const ageGroup = age <= 15 ? '0-15' : (age <= 35 ? '16-35' : (age <= 55 ? '36-55' : '55+'));
    let pId = fd.p_id;
    if (!isEdit) {
        pId = await window.generateNextPatientID();
    }
    const row = {
        Patient_ID: pId, Title: fd.p_title, First_Name: fd.p_firstname, Last_Name: fd.p_lastname,
        Gender: fd.p_gender, Date_of_Birth: fd.p_dob || null, Age: age,
        Nationality: fd.p_nation, Occupation: fd.p_job, Blood_Type: fd.p_blood,
        Phone_Number: fd.p_phone, Email: fd.p_email, Address: fd.p_address,
        District: fd.p_district, Province: fd.p_province,
        Organization_ID: fd.p_org_id, Name_Org: fd.p_org_name,
        Insurance_Company: fd.p_ins_company, Insurance_Code: fd.p_ins_code, Insured_Person_Name: fd.p_ins_name,
        Drug_Allergy: fd.p_allergy, Underlying_Disease: fd.p_disease,
        Emergency_Name: fd.p_emer_name, Emergency_Contact: fd.p_emer_contact, Emergency_Relation: fd.p_emer_relation,
        Channel: fd.p_channel, Registration_Date: fd.p_date || null, Time: fd.p_time,
        Shift: fd.p_shift, Age_Group: ageGroup
    };
    let result;
    if (isEdit) result = await supabaseClient.from('Patients').update(row).eq('Patient_ID', pId);
    else result = await supabaseClient.from('Patients').insert(row);
    Swal.close();
    if (result.error) { Swal.fire('ຜິດພາດ', result.error.message, 'error'); return; }
    $('#patientModal').modal('hide');
    window.logAction(isEdit ? 'Edit' : 'Add', `${isEdit ? 'ແກ້ໄຂ' : 'ເພີ່ມ'}ຄົນເຈັບ: ${pId} - ${row.First_Name} ${row.Last_Name}`, 'Patients');
    window.initPatientTable();
    window.preloadDropdownData();
    Swal.fire({
        title: 'ລົງທະບຽນສຳເລັດ!', text: 'ລະຫັດ: ' + pId, icon: 'success',
        showCancelButton: true, showDenyButton: true,
        confirmButtonText: 'ສົ່ງເຂົ້າ Triage', denyButtonText: 'ພິມບັດ QR', cancelButtonText: 'ປິດ',
        confirmButtonColor: '#f59e0b', denyButtonColor: '#0ea5e9'
    }).then((rr) => {
        if (rr.isConfirmed) window.sendToTriageFlow(pId, (fd.p_firstname || '') + ' ' + (fd.p_lastname || ''));
        else if (rr.isDenied) window.printQRCard(pId);
    });
};

window.delPatient = async function (id) {
    const result = await Swal.fire({ title: 'ລຶບ?', icon: 'warning', showCancelButton: true, confirmButtonColor: '#ef4444', confirmButtonText: 'ລຶບ' });
    if (result.isConfirmed) {
        await supabaseClient.from('Patients').delete().eq('Patient_ID', id);
        window.initPatientTable();
        window.preloadDropdownData();
    }
};

window.printQRCard = async function (id) {
    Swal.fire({ title: 'ກຳລັງສ້າງບັດ QR...', allowOutsideClick: false, didOpen: () => Swal.showLoading() });
    const { data, error } = await supabaseClient.from('Patients').select('*').eq('Patient_ID', id).single();
    if (error || !data) return Swal.fire('ຜິດພາດ', 'ບໍ່ພົບຂໍ້ມູນຄົນເຈັບ', 'error');
    const d = {
        id: data.Patient_ID, title: data.Title || '', firstname: data.First_Name || '',
        lastname: data.Last_Name || '', dob: data.Date_of_Birth || '', age: data.Age || '', phone: data.Phone_Number || '-'
    };
    $('#printID').text(d.id);
    $('#printName').text(`${d.title} ${d.firstname} ${d.lastname}`.trim());
    $('#printDob').text(`${d.dob} (${d.age} ປີ)`);
    $('#printPhone').text(d.phone);
    $('#qrcodeDisplay').empty();
    new QRCode(document.getElementById('qrcodeDisplay'), {
        text: d.id, width: 65, height: 65, colorDark: '#0f172a', colorLight: '#ffffff', correctLevel: QRCode.CorrectLevel.H
    });
    Swal.close();
    window.executePrint('print-area');
};

window.openQRScanner = function () {
    $('#qrScannerModal').modal('show');
    setTimeout(() => {
        if (!html5QrCode) html5QrCode = new Html5Qrcode("reader");
        html5QrCode.start({ facingMode: "environment" }, { fps: 10, qrbox: { width: 250, height: 250 }, aspectRatio: 1.0 }, window.onScanSuccess, () => { });
    }, 400);
};

window.closeQRScanner = function () {
    $('#qrScannerModal').modal('hide');
    if (html5QrCode) {
        try { html5QrCode.stop().then(() => html5QrCode.clear()).catch(() => html5QrCode.clear()); }
        catch (e) { html5QrCode.clear(); }
    }
};

window.onScanSuccess = function (t) {
    window.closeQRScanner();
    setTimeout(() => {
        $('#patientTable').DataTable().search(t).draw();
        Swal.fire({ title: 'ພົບລະຫັດ!', text: t, icon: 'success', timer: 1500, showConfirmButton: false });
    }, 500);
};

window.sendToTriageFlow = async function (id, n) {
    const result = await Swal.fire({
        title: 'ສົ່ງເຂົ້າ Triage?',
        text: `ສົ່ງ ${n} ໄປຈຸດຊັກປະຫວັດ?`,
        icon: 'question',
        showCancelButton: true,
        confirmButtonText: 'ສົ່ງເຂົ້າຄິວ',
        confirmButtonColor: '#f59e0b'
    });
    if (result.isConfirmed) {
        Swal.fire({ title: 'ກຳລັງສົ່ງເຂົ້າຄິວ...', didOpen: () => Swal.showLoading() });
        const { data: lastVisit } = await supabaseClient.from('Visits').select('Visit_ID').order('Visit_ID', { ascending: false }).limit(1);
        let lastNum = 0;
        if (lastVisit && lastVisit.length > 0) {
            const m = (lastVisit[0].Visit_ID || "").match(/\d+$/);
            if (m) lastNum = parseInt(m[0]);
        }
        const vId = "V" + new Date().getFullYear().toString().slice(-2) + "-" + ("0000" + (lastNum + 1)).slice(-4);
        const { error } = await supabaseClient.from('Visits').insert({
            Visit_ID: vId, Date: new Date().toISOString(),
            Patient_ID: id, Patient_Name: n, Status: 'Triage'
        });
        if (error) Swal.fire('ຜິດພາດ!', error.message, 'error');
        else {
            window.logAction('Add', 'Send to Triage: ' + n + ' (' + id + ')', 'Triage');
            Swal.fire('ສຳເລັດ!', 'ສົ່ງເຂົ້າ Triage ແລ້ວ', 'success');
        }
    }
};

window.submitTriageForm = function (e) {
    if (e) e.preventDefault();
    const fd = {};
    new FormData($('#triageForm')[0]).forEach((v, k) => fd[k] = v);
    let bp = fd.v_bp || "";
    let a = false;
    let m = "";

    if (bp.includes('/')) {
        let [s, d] = bp.split('/').map(Number);
        if (!isNaN(s) && !isNaN(d)) {
            if (s >= 140 || d >= 90) { a = true; m = `ຄວາມດັນສູງ (${bp})`; }
            else if (s <= 90 || d <= 60) { a = true; m = `ຄວາມດັນຕ່ຳ (${bp})`; }
        }
    }

    if (a) {
        Swal.fire({
            title: 'ແຈ້ງເຕືອນຄວາມດັນ!',
            html: `<h4 class="text-danger fw-bold mb-3">${m}</h4><p>ບັນທຶກຕໍ່ໄປແທ້ບໍ່?</p>`,
            icon: 'warning',
            showCancelButton: true,
            confirmButtonColor: '#ef4444',
            confirmButtonText: 'ບັນທຶກ'
        }).then(r => {
            if (r.isConfirmed) window.executeTriageSave(fd);
        });
    } else {
        window.executeTriageSave(fd);
    }
};

window.executeTriageSave = async function (fd) {
    Swal.fire({ title: 'ກຳລັງບັນທຶກ...', didOpen: () => Swal.showLoading() });
    const { error } = await supabaseClient.from('Visits').update({
        Status: 'Waiting OPD', BP: fd.v_bp, Temp: fd.v_temp,
        Weight: fd.v_weight, Height: fd.v_height,
        Pulse: fd.v_pulse, SpO2: fd.v_spo2,
        Department: fd.v_department, Symptoms: fd.v_symptoms
    }).eq('Visit_ID', fd.visitId);

    if (error) {
        Swal.fire('ຜິດພາດ!', error.message, 'error');
    } else {
        $('#triageModal').modal('hide');
        window.logAction('Save', 'Triage saved - Visit ' + fd.visitId, 'Triage');
        Swal.fire('ສຳເລັດ!', 'ສົ່ງເຂົ້າຫ້ອງກວດແລ້ວ', 'success');
        window.loadTriageQueue();
    }
};

window.openEMRLabModal = function () {
    $('.lab-checkbox').prop('checked', false);
    if (document.activeElement) document.activeElement.blur();
    $('#emrLabModal').modal('show');
};

window.addLabToEMRList = function () {
    currentEMRLabs = [];
    $('.lab-checkbox:checked').each(function () { currentEMRLabs.push({ name: $(this).val() }); });
    window.renderEMRLabTable();
    $('#emrLabModal').modal('hide');
};

window.renderEMRLabTable = function () {
    let h = '';
    if (currentEMRLabs.length === 0) {
        h = '<tr><td colspan="2" class="text-center text-muted small py-2">ຍັງບໍ່ມີລາຍການ</td></tr>';
    } else {
        currentEMRLabs.forEach((x, i) => {
            h += `<tr>
                    <td class="text-primary fw-bold">${x.name}</td>
                    <td class="text-center"><button type="button" class="btn btn-sm btn-danger py-0 px-2" onclick="window.removeEMRLab(${i})"><i class="fas fa-times"></i></button></td>
                  </tr>`;
        });
    }
    $('#emrLabTableBody').html(h);
};

window.removeEMRLab = function (i) { currentEMRLabs.splice(i, 1); window.renderEMRLabTable(); };

window.openEMRDrugModal = function () {
    $('#emrAddDrugQty').val(1);
    $('#emrAddDrugUnit,#emrAddDrugUsage').prop('selectedIndex', 0);
    $('#emrAddDrugSelect').val(null).trigger('change');
    if (document.activeElement) document.activeElement.blur();
    $('#emrDrugModal').modal('show');
};

window.addDrugToEMRList = function () {
    let d = $('#emrAddDrugSelect').select2('data');
    if (!d || d.length === 0 || !d[0].id) return Swal.fire('ແຈ້ງເຕືອນ', 'ກະລຸນາເລືອກຢາກ່ອນ', 'warning');
    currentEMRDrugs.push({
        name: d[0].text,
        qty: $('#emrAddDrugQty').val() + ' ' + $('#emrAddDrugUnit').val(),
        usage: $('#emrAddDrugUsage').val()
    });
    window.renderEMRDrugTable();
    $('#emrDrugModal').modal('hide');
};

window.renderEMRDrugTable = function () {
    let h = '';
    if (currentEMRDrugs.length === 0) {
        h = '<tr><td colspan="4" class="text-center text-muted small py-3"><i class="fas fa-pills mb-2 opacity-50 fa-2x"></i><br>ຍັງບໍ່ມີລາຍການຢາ</td></tr>';
    } else {
        currentEMRDrugs.forEach((x, i) => {
            h += `<tr>
                    <td class="p-2 border-bottom">
                        <div class="d-flex justify-content-between align-items-center">
                            <div class="d-flex flex-column">
                                <span class="fw-bold text-success mb-1" style="font-size: 13.5px;"><i class="fas fa-capsules me-1 opacity-75"></i>${x.name}</span>
                                <div class="d-flex align-items-center gap-2">
                                    <span class="badge bg-light text-dark border border-secondary border-opacity-25 px-2 py-1"><i class="fas fa-cubes me-1 opacity-50"></i>${x.qty}</span>
                                    <span class="small text-muted"><i class="fas fa-info-circle me-1 opacity-50"></i>${x.usage}</span>
                                </div>
                            </div>
                            <button type="button" class="btn btn-sm btn-outline-danger border-0 rounded-circle" style="width: 28px; height: 28px; padding: 0; display: flex; align-items: center; justify-content: center;" onclick="window.removeEMRDrug(${i})" title="ລຶບລາຍການ">
                                <i class="fas fa-times"></i>
                            </button>
                        </div>
                    </td>
                  </tr>`;
        });
    }
    $('#emrDrugTableBody').html(h);
};

window.removeEMRDrug = function (i) { currentEMRDrugs.splice(i, 1); window.renderEMRDrugTable(); };

window.calculateBMI = function () {
    let w = parseFloat($('#v_weight').val());
    let h = parseFloat($('#v_height').val());
    if (w > 0 && h > 0) {
        let bmi = (w / Math.pow(h / 100, 2)).toFixed(1);
        let status = bmi >= 25 ? " (ຕຸ້ຍ)" : (bmi < 18.5 ? " (ຈ່ອຍ)" : " (ປົກກະຕິ)");
        $('#v_bmi').val(bmi + status);
    } else {
        $('#v_bmi').val("");
    }
};

window._fetchTriageQueue = async function () {
    try {
        let today = new Date().toISOString().split('T')[0];
        // 1. Fetch Visits
        const { data: visits, error: vError } = await supabaseClient.from('Visits')
            .select('*')
            .gte('Date', today + 'T00:00:00Z')
            .lte('Date', today + 'T23:59:59Z')
            .order('Date', { ascending: true });

        if (vError) throw vError;
        if (!visits || visits.length === 0) return [];

        // 2. Fetch unique Patient IDs
        const pIds = [...new Set(visits.map(v => v.Patient_ID).filter(id => !!id))];
        let pMap = {};
        if (pIds.length > 0) {
            const { data: patients, error: pError } = await supabaseClient.from('Patients')
                .select('Patient_ID, Age')
                .in('Patient_ID', pIds);
            if (!pError && patients) {
                patients.forEach(p => pMap[p.Patient_ID] = p);
            }
        }
        
        // 3. Determine isNew
        let firstVisitMap = {};
        if (pIds.length > 0) {
            const { data: allVs } = await supabaseClient.from('Visits').select('Visit_ID, Patient_ID, Date').in('Patient_ID', pIds).order('Date', { ascending: true });
            (allVs || []).forEach(v => { if (!firstVisitMap[v.Patient_ID]) firstVisitMap[v.Patient_ID] = v.Visit_ID; });
        }

        // 4. Merge
        return visits.map((r, i) => {
            let dObj = new Date(r.Date);
            let p = pMap[r.Patient_ID];
            return {
                rowIdx: r.Visit_ID, visitId: r.Visit_ID,
                date: dObj.toLocaleDateString('en-GB'), time: dObj.toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit' }),
                patientId: r.Patient_ID, patientName: r.Patient_Name,
                status: r.Status, department: r.Department || 'OPD',
                isNew: (firstVisitMap[r.Patient_ID] === r.Visit_ID),
                age: p?.Age || 0,
                bp: r.BP, temp: r.Temp, weight: r.Weight, height: r.Height,
                bmi: r.BMI, pulse: r.Pulse, spo2: r.SpO2, symptoms: r.Symptoms
            };
        });
    } catch (err) {
        console.error('Triage Fetch Error:', err);
        return [];
    }
};

window._fetchOpdQueue = async function () {
    let today = new Date().toISOString().split('T')[0];
    const { data, error } = await supabaseClient.from('Visits').select('*')
        .gte('Date', today + 'T00:00:00Z')
        .lte('Date', today + 'T23:59:59Z')
        .order('Date', { ascending: true });

    if (error || !data) return [];
    const pIds = [...new Set(data.map(v => v.Patient_ID).filter(id => !!id))];
    let firstVisitMap = {};
    if (pIds.length > 0) {
        const { data: allVs } = await supabaseClient.from('Visits').select('Visit_ID, Patient_ID, Date').in('Patient_ID', pIds).order('Date', { ascending: true });
        (allVs || []).forEach(v => { if (!firstVisitMap[v.Patient_ID]) firstVisitMap[v.Patient_ID] = v.Visit_ID; });
    }

    return data.map((r, i) => {
        let dObj = new Date(r.Date);
        return {
            rowIdx: r.Visit_ID, visitId: r.Visit_ID,
            date: dObj.toLocaleDateString('en-GB'), time: dObj.toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit' }),
            patientId: r.Patient_ID, patientName: r.Patient_Name,
            status: r.Status, department: r.Department || 'OPD',
            isNew: (firstVisitMap[r.Patient_ID] === r.Visit_ID),
            dischargeStatus: r.Discharge_Status, doctor: r.Doctor_Name,
            symptoms: r.Symptoms, bp: r.BP, temp: r.Temp, weight: r.Weight, height: r.Height,
            pe: r.Physical_Exam, diagnosis: r.Diagnosis, advice: r.Advice, followup: r.Follow_Up,
            labOrdersStr: r.Lab_Orders_JSON, prescriptionStr: r.Prescription_JSON,
            site: r.Site, type: r.Visit_Type, services: r.Services_List
        };
    });
};

window.loadTriageQueue = async function () {
    if ($.fn.DataTable.isDataTable('#triageTable')) $('#triageTable').DataTable().destroy();
    $('#triageTableBody').html('<tr><td colspan="6" class="text-center py-4"><div class="spinner-border text-danger spinner-border-sm"></div> ກຳລັງໂຫຼດ...</td></tr>');
    const q = await window._fetchTriageQueue();
    currentTriageData = q || [];
    if ($.fn.DataTable.isDataTable('#triageTable')) $('#triageTable').DataTable().destroy();
    let h = '';
    if (q && q.length > 0) {
        q.forEach((r, i) => {
            let sb = r.status === 'Triage' ? '<span class="badge bg-warning text-dark"><i class="fas fa-hourglass-half"></i> ລໍຖ້າວັດແທກ</span>' : `<span class="badge bg-success"><i class="fas fa-check-circle"></i> ໄປ ${r.department} ແລ້ວ</span>`;
            let nb = r.isNew ? '<span class="badge bg-success ms-2">ໃໝ່</span>' : '<span class="badge bg-secondary ms-2">ເກົ່າ</span>';
            let btnHtml = `<button class="btn btn-sm btn-info text-white shadow-sm me-1" onclick="window.viewTriage(${i})"><i class="fas fa-eye"></i></button>`;
            if (r.status === 'Triage') {
                btnHtml += `<button class="btn btn-sm btn-danger fw-bold shadow-sm me-1" onclick="window.openTriage(${i})"><i class="fas fa-stethoscope"></i> ວັດແທກ</button>`;
            } else {
                btnHtml += `<button class="btn btn-sm btn-primary shadow-sm me-1" onclick="window.openTriage(${i})"><i class="fas fa-edit"></i></button>`;
            }
            btnHtml += `<button class="btn btn-sm btn-outline-danger shadow-sm me-1" onclick="window.deleteVisitFlow('${r.visitId}')"><i class="fas fa-trash"></i></button>
                        <button class="btn btn-sm btn-secondary text-white shadow-sm" onclick="window.printOPDCard('triage', ${i})"><i class="fas fa-file-medical"></i></button>`;
            h += `<tr>
                    <td class="text-muted">${r.date}</td>
                    <td class="fw-bold">${r.time}</td>
                    <td><div class="fw-bold text-primary">${r.patientName} ${nb}</div><div class="small text-muted">${r.patientId}</div></td>
                    <td><span class="badge bg-secondary rounded-pill">${r.age} ປີ</span></td>
                    <td>${sb}</td>
                    <td class="text-center"><div class="d-flex gap-1 justify-content-center">${btnHtml}</div></td>
                  </tr>`;
        });
    }
    $('#triageTableBody').html(h);
    $('#triageTable').DataTable({ responsive: true, pageLength: 10, language: { search: "ຄົ້ນຫາ:", emptyTable: "ບໍ່ມີຄິວ Triage" } });
};

window.viewTriage = function (i) {
    let r = currentTriageData[i];
    let isDone = r.status !== 'Triage';
    let statusBadge = isDone
        ? `<span class="badge bg-success fs-6 px-3 py-2"><i class="fas fa-check-circle me-1"></i> ສົ່ງຫ້ອງກວດແລ້ວ</span>`
        : `<span class="badge bg-warning text-dark fs-6 px-3 py-2"><i class="fas fa-hourglass-half me-1"></i> ລໍຖ້າວັດແທກ</span>`;
    let newBadge = r.isNew
        ? `<span class="badge bg-success ms-2">ຄົນເຈັບໃໝ່</span>`
        : `<span class="badge bg-secondary ms-2">ຄົນເຈັບເກົ່າ</span>`;

    const makeStat = (label, val, unit, color) =>
        `<div class="col-6 col-md-4 mb-3">
            <div class="p-3 border rounded text-center h-100" style="background:#f8fafc;">
                <div class="small text-muted fw-bold mb-1">${label}</div>
                <div class="fw-bold fs-5 text-${color}">${val || '<span class="text-muted">-</span>'}</div>
                ${unit ? `<div class="small text-muted">${unit}</div>` : ''}
            </div>
        </div>`;

    // Blood pressure color logic
    let bpColor = 'primary';
    if (r.bp && r.bp.includes('/')) {
        let parts = r.bp.split('/').map(Number);
        if (!isNaN(parts[0]) && !isNaN(parts[1])) {
            if (parts[0] >= 140 || parts[1] >= 90) bpColor = 'danger';
            else if (parts[0] <= 90 || parts[1] <= 60) bpColor = 'warning';
            else bpColor = 'success';
        }
    }

    let vitalsHtml = isDone ? `
        <div class="row g-2">
            ${makeStat('<i class="fas fa-heart me-1"></i> BP', r.bp, 'mmHg', bpColor)}
            ${makeStat('<i class="fas fa-thermometer-half me-1"></i> Temp', r.temp ? r.temp + ' °C' : null, null, parseFloat(r.temp) >= 37.5 ? 'danger' : 'info')}
            ${makeStat('<i class="fas fa-tint me-1"></i> Pulse', r.pulse, 'bpm', 'warning')}
            ${makeStat('<i class="fas fa-lungs me-1"></i> SpO2', r.spo2 ? r.spo2 + ' %' : null, null, parseFloat(r.spo2) < 95 ? 'danger' : 'success')}
            ${makeStat('<i class="fas fa-weight me-1"></i> Weight', r.weight, 'kg', 'secondary')}
            ${makeStat('<i class="fas fa-ruler-vertical me-1"></i> Height', r.height, 'cm', 'secondary')}
        </div>
        ${r.symptoms ? `<div class="alert alert-light border mt-2 py-2"><b><i class="fas fa-comment-medical text-danger me-1"></i> ອາການ (CC):</b> ${r.symptoms}</div>` : ''}
    ` : `<div class="text-center py-4">
        <i class="fas fa-hourglass-half fa-3x text-warning mb-3 d-block"></i>
        <p class="text-muted">ຍັງບໍ່ໄດ້ວັດ Vital Signs</p>
    </div>`;

    Swal.fire({
        title: `<span style="font-size:16px; font-weight:700;"><i class="fas fa-heartbeat text-danger me-2"></i>${r.patientName} ${newBadge}</span>`,
        html: `
            <div class="text-start">
                <div class="d-flex align-items-center justify-content-between mb-3 p-3 rounded"
                     style="background: linear-gradient(135deg, #1e293b, #334155); color:white; border-radius: 10px;">
                    <div>
                        <div class="fw-bold fs-6">${r.patientId}</div>
                        <div class="small opacity-75"><i class="fas fa-clock me-1"></i>${r.date} ${r.time}</div>
                    </div>
                    <div class="text-end">
                        ${statusBadge}
                        <div class="small mt-1 opacity-75"><i class="fas fa-door-open me-1"></i>${r.department || 'OPD'}</div>
                    </div>
                </div>
                ${vitalsHtml}
            </div>`,
        width: '600px',
        confirmButtonText: isDone ? '<i class="fas fa-edit me-1"></i> แก้ไข' : '<i class="fas fa-stethoscope me-1"></i> ວັດແທກຕອນນີ້',
        confirmButtonColor: isDone ? '#0ea5e9' : '#ef4444',
        showCancelButton: true,
        cancelButtonText: 'ປິດ',
        customClass: { popup: 'shadow-lg' }
    }).then(res => { if (res.isConfirmed) window.openTriage(i); });
};

window.openTriage = function (i) {
    let r = currentTriageData[i];
    $('#vRowIdx').val(r.rowIdx);
    $('#vPatientId').text(r.patientId);
    $('#vPatientName').text(r.patientName);
    $('#triageForm')[0].reset();
    $('input[name="v_bp"]').removeClass('border-danger text-danger bg-danger border-warning text-dark bg-warning bg-opacity-10 border-success text-success fw-bold');

    if (r.status !== 'Triage') {
        $('input[name="v_bp"]').val(r.bp || '').trigger('input');
        $('input[name="v_temp"]').val(r.temp || '');
        $('input[name="v_weight"]').val(r.weight || '');
        $('input[name="v_height"]').val(r.height || '');
        $('input[name="v_pulse"]').val(r.pulse || '');
        $('input[name="v_spo2"]').val(r.spo2 || '');
        $('textarea[name="v_symptoms"]').val(r.symptoms || '');
        $('select[name="v_department"]').val(r.department || '');
        window.calculateBMI();
    } else {
        $('#v_bmi').val('');
    }
    if (document.activeElement) document.activeElement.blur();
    $('#triageModal').modal('show');
};

window.deleteVisitFlow = async function (id) {
    const r = await Swal.fire({ title: 'ລຶບຄິວ?', icon: 'warning', showCancelButton: true, confirmButtonColor: '#ef4444', confirmButtonText: 'ລຶບ' });
    if (r.isConfirmed) {
        Swal.fire({ title: 'ກຳລັງລຶບ...', didOpen: () => Swal.showLoading() });
        await supabaseClient.from('Visits').delete().eq('Visit_ID', id);
        Swal.fire('ສຳເລັດ!', 'ລຶບແລ້ວ', 'success');
        window.loadTriageQueue();
        window.loadQueue();
    }
};

window.loadQueue = async function () {
    if ($.fn.DataTable.isDataTable('#queueTable')) $('#queueTable').DataTable().destroy();
    $('#queueTableBody').html('<tr><td colspan="6" class="text-center py-4"><div class="spinner-border text-info spinner-border-sm"></div> ກຳລັງໂຫຼດ...</td></tr>');

    const q = await window._fetchOpdQueue();
    queueDataStore = q || [];
    if ($.fn.DataTable.isDataTable('#queueTable')) $('#queueTable').DataTable().destroy();
    let h = '';
    if (q && q.length > 0) {
        q.forEach((r, i) => {
            let b = '', a = '';
            if (r.status === 'Waiting OPD') {
                b = '<span class="badge bg-warning text-dark"><i class="fas fa-user-clock"></i> ລໍຖ້າກວດ</span>';
                a = `<button class="btn btn-sm btn-info text-white fw-bold me-1" onclick="window.openEMR(${i})"><i class="fas fa-stethoscope"></i> ເປີດກວດ</button>
                     <button class="btn btn-sm btn-secondary text-white" onclick="window.printOPDCard('opd', ${i})"><i class="fas fa-file-medical"></i> ພິມ</button>`;
            } else if (r.status === 'Waiting Lab') {
                b = '<span class="badge bg-primary"><i class="fas fa-flask"></i> ລໍຖ້າຜົນແລັບ</span>';
                a = `<button class="btn btn-sm btn-primary text-white fw-bold me-1" onclick="window.openEMR(${i})"><i class="fas fa-edit"></i> ອ່ານຜົນແລັບ</button>
                     <button class="btn btn-sm btn-secondary text-white" onclick="window.printOPDCard('opd', ${i})"><i class="fas fa-file-medical"></i> ພິມ</button>`;
            } else {
                b = `<span class="badge bg-success"><i class="fas fa-check-circle"></i> ປິດຈົບ (${r.dischargeStatus || 'ກວດສຳເລັດ'})</span>`;
                a = `<button class="btn btn-sm btn-success text-white fw-bold me-1" onclick="window.viewEMR(${i})" title="ເບິ່ງລາຍລະອຽດການກວດ"><i class="fas fa-eye"></i></button>
                     <button class="btn btn-sm btn-primary text-white fw-bold me-1" onclick="window.openEMR(${i})" title="ແກ້ໄຂການກວດ"><i class="fas fa-edit"></i></button>
                     <button class="btn btn-sm btn-secondary text-white" onclick="window.printOPDCard('opd', ${i})"><i class="fas fa-print"></i></button>`;
            }
            let nb = r.isNew ? '<span class="badge bg-success ms-2">ໃໝ່</span>' : '<span class="badge bg-secondary ms-2">ເກົ່າ</span>';
            let dateTimeStr = (r.date ? r.date + ' ' : '') + r.time;
            h += `<tr>
                    <td class="text-muted small">${dateTimeStr}</td>
                    <td class="text-dark fw-bold">${r.patientId}</td>
                    <td><div class="fw-bold text-primary">${r.patientName} ${nb}</div></td>
                    <td><span class="badge bg-light text-dark border px-2 py-1">${r.department}</span></td>
                    <td>${b}</td>
                    <td class="text-center"><div class="d-flex gap-1 justify-content-center">${a}</div></td>
                  </tr>`;
        });
    }
    $('#queueTableBody').html(h);
    $('#queueTable').DataTable({ responsive: true, pageLength: 10, language: { search: "ຄົ້ນຫາ:", emptyTable: "ບໍ່ມີຄິວລໍຖ້າ" } });
};

window.viewEMR = function (i) {
    let q = queueDataStore[i];
    if (!q) return;

    let labList = "ບໍ່ມີການສັ່ງກວດ", drugList = "ບໍ່ມີການສັ່ງຢາ";
    try {
        let labs = q.labOrdersStr ? JSON.parse(q.labOrdersStr) : [];
        if (labs.length > 0) {
            labList = "<ul class='mb-0 text-start ps-3'>";
            labs.forEach(l => labList += `<li>${l.name}</li>`);
            labList += "</ul>";
        }
    } catch (e) { }

    try {
        let drugs = q.prescriptionStr ? JSON.parse(q.prescriptionStr) : [];
        if (drugs.length > 0) {
            drugList = "<ul class='mb-0 text-start ps-3'>";
            drugs.forEach(d => drugList += `<li><b>${d.name}</b>: <span class="badge bg-secondary">${d.qty}</span> (${d.usage})</li>`);
            drugList += "</ul>";
        }
    } catch (e) { }

    let htmlBody = `
        <div class="text-start" style="font-size: 14px; line-height: 1.6;">
            <div class="row border-bottom pb-2 mb-2">
                <div class="col-6"><b><i class="fas fa-user text-primary"></i> ຄົນເຈັບ:</b> ${q.patientName} <br><small class="text-muted">(${q.patientId})</small></div>
                <div class="col-6"><b><i class="far fa-clock text-info"></i> ເວລາ:</b> ${q.time}</div>
            </div>
            <p><b><i class="fas fa-user-md text-primary"></i> ແພດຜູ້ກວດ:</b> <span class="text-primary fw-bold">${q.doctor || '-'}</span></p>
            <p><b><i class="fas fa-comment-medical text-danger"></i> ອາການເບື້ອງຕົ້ນ (CC):</b><br> ${q.symptoms || '-'}</p>
            <div class="bg-light p-2 rounded mb-3 border">
                <b><i class="fas fa-heartbeat text-danger"></i> Vitals:</b> 
                BP: <span class="text-primary fw-bold">${q.bp || '-'}</span> | 
                Temp: <span class="text-danger fw-bold">${q.temp ? q.temp + ' °C' : '-'}</span> | 
                Wt: <span class="text-success fw-bold">${q.weight ? q.weight + ' kg' : '-'}</span>
            </div>
            <p><b><i class="fas fa-search text-dark"></i> ຜົນການກວດ (PE):</b><br> ${q.pe || '-'}</p>
            <p><b><i class="fas fa-stethoscope text-danger"></i> ການວິນິດໄສ (Dx):</b><br> <span class="text-danger fw-bold">${q.diagnosis || '-'}</span></p>
            <div class="row mt-3">
                <div class="col-md-6 mb-2">
                    <div class="border border-primary rounded p-2 h-100">
                        <b class="text-primary"><i class="fas fa-flask"></i> ລາຍການແລັບ:</b>
                        <div class="mt-1">${labList}</div>
                    </div>
                </div>
                <div class="col-md-6 mb-2">
                    <div class="border border-success rounded p-2 h-100">
                        <b class="text-success"><i class="fas fa-pills"></i> ລາຍການຢາ:</b>
                        <div class="mt-1">${drugList}</div>
                    </div>
                </div>
            </div>
            <p class="mt-3 mb-1"><b><i class="fas fa-comment-dots text-warning"></i> ຄຳແນະນຳ:</b> ${q.advice || '-'}</p>
            <p class="mb-1"><b><i class="fas fa-calendar-check text-info"></i> ນັດໝາຍ:</b> ${q.followup || '-'}</p>
            <p class="mb-0"><b><i class="fas fa-clipboard-check text-success"></i> ສະຖານະ:</b> <span class="badge bg-success">${q.dischargeStatus || '-'}</span></p>
        </div>
    `;

    Swal.fire({
        title: '<i class="fas fa-file-medical-alt text-primary"></i> ຂໍ້ມູນການກວດພະຍາດ',
        html: htmlBody,
        width: '700px',
        showCloseButton: true,
        focusConfirm: false,
        confirmButtonText: 'ປິດ',
        confirmButtonColor: '#64748b'
    });
};

window.handleSiteChange = function () {
    let site = $('#emrSite').val();
    if (!site) site = 'In-site';
    let typeSelect = $('#emrDeptType');
    typeSelect.empty();
    let options = [];
    if (site === 'In-site' || site === 'In-Site') {
        options = (masterDataStore['PatientType_InSite'] && masterDataStore['PatientType_InSite'].length > 0) ? masterDataStore['PatientType_InSite'].map(x => x.value) : ['OPD', 'IPD'];
    } else {
        options = (masterDataStore['PatientType_Onsite'] && masterDataStore['PatientType_Onsite'].length > 0) ? masterDataStore['PatientType_Onsite'].map(x => x.value) : ['Checkup Corporation', 'Individual First Aid', 'Corporation First Aid', 'HomeCare'];
    }
    let h = '';
    options.forEach(opt => h += `<option value="${opt}">${opt}</option>`);
    typeSelect.html(h);
};

window.openEMR = function (i) {
    let q = queueDataStore[i];
    $('#emrRowIdx').val(q.rowIdx);
    $('#emrPatientId').text(q.patientId);
    $('#emrPatientName').text(q.patientName);

    if (q.allergy && q.allergy.trim() !== "" && q.allergy !== "ບໍ່ມີ" && q.allergy !== "-") {
        $('#emrAllergy').text(q.allergy).removeClass('text-secondary').addClass('text-danger');
        $('#emrAllergyBox').removeClass('bg-light border-secondary').addClass('bg-danger bg-opacity-10 border-danger');
        $('#emrAllergyIcon').removeClass('text-secondary').addClass('text-danger');
    } else {
        $('#emrAllergy').text("ບໍ່ມີປະຫວັດການແພ້").removeClass('text-danger').addClass('text-secondary');
        $('#emrAllergyBox').removeClass('bg-danger bg-opacity-10 border-danger').addClass('bg-light border-secondary');
        $('#emrAllergyIcon').removeClass('text-danger').addClass('text-secondary');
    }

    $('#emrCC').text(q.symptoms || "-");
    $('#emrBp').removeClass('text-danger text-warning text-success fw-bold').addClass('text-dark');
    if (q.bp) {
        $('#emrBp').text(q.bp + " mmHg");
        let [s, d] = q.bp.split('/').map(Number);
        if (!isNaN(s) && !isNaN(d)) {
            if (s >= 140 || d >= 90) $('#emrBp').addClass('text-danger fw-bold');
            else if (s <= 90 || d <= 60) $('#emrBp').addClass('text-warning text-dark fw-bold');
            else $('#emrBp').addClass('text-success fw-bold');
        }
    } else {
        $('#emrBp').text("-");
    }

    $('#emrTemp').removeClass('text-danger text-success fw-bold').addClass('text-dark');
    if (q.temp) {
        $('#emrTemp').text(q.temp + " °C");
        let t = parseFloat(q.temp);
        if (!isNaN(t)) {
            if (t >= 37.5) $('#emrTemp').addClass('text-danger fw-bold');
            else $('#emrTemp').addClass('text-success fw-bold');
        }
    } else {
        $('#emrTemp').text("-");
    }

    $('#emrWeight').text(q.weight ? q.weight + " kg" : "-");
    $('#emrPE').val(q.pe || '');
    $('#emrDiagnosis').val(q.diagnosis || '');
    $('#emrAdvice').val(q.advice || '');
    $('#emrFollowup').val(q.followup || '');
    $('#emrDischargeStatus').val(q.dischargeStatus || '');
    $('#emrDoctor').val(q.doctor || (currentUser ? currentUser.name : ''));

    let os = '';
    let siteOptions = masterDataStore['Site'] ? masterDataStore['Site'].map(x => x.value) : ['In-site', 'Onsite'];
    siteOptions.forEach(x => os += `<option value="${x}">${x}</option>`);
    $('#emrSite').html(os).val(q.site || "In-site");
    window.handleSiteChange();

    if (q.type) {
        if (!Array.from($('#emrDeptType')[0].options).some(o => o.value === q.type)) {
            $('#emrDeptType').append(`<option value="${q.type}">${q.type}</option>`);
        }
        $('#emrDeptType').val(q.type);
    }

    $('#emrService').val(q.services ? q.services.split(',').map(x => x.trim()) : null).trigger('change');

    try { currentEMRLabs = q.labOrdersStr ? JSON.parse(q.labOrdersStr) : []; } catch (e) { currentEMRLabs = []; }
    try { currentEMRDrugs = q.prescriptionStr ? JSON.parse(q.prescriptionStr) : []; } catch (e) { currentEMRDrugs = []; }

    window.renderEMRLabTable();
    window.renderEMRDrugTable();

    if (document.activeElement) document.activeElement.blur();
    $('#emrModal').modal('show');
};

window.submitEMRForm = async function (e) {
    if (e) e.preventDefault();
    let ds = $('#emrDischargeStatus').val();
    if (!ds) return Swal.fire('ແຈ້ງເຕືອນ', 'ກະລຸນາເລືອກສະຖານະປິດຈົບການກວດ', 'warning');

    let dx = $('#emrDiagnosis').val();
    if (ds !== "ລໍຖ້າຜົນແລັບ (Waiting Lab)" && dx.trim() === "") {
        return Swal.fire('ແຈ້ງເຕືອນ', 'ກະລຸນາປ້ອນ ການວິນິດໄສ (Diagnosis) ກ່ອນທີ່ຈະປິດຈົບການກວດ!', 'warning');
    }

    let docName = $('#emrDoctor').val() || (currentUser ? currentUser.name : 'Doctor');
    let presJson = currentEMRDrugs.length > 0 ? JSON.stringify(currentEMRDrugs) : "";
    let labJson = currentEMRLabs.length > 0 ? JSON.stringify(currentEMRLabs) : "";

    Swal.fire({ title: 'ກຳລັງບັນທຶກ...', didOpen: () => Swal.showLoading() });

    let visitId = $('#emrRowIdx').val();
    let statusMap = {
        "ລໍຖ້າຜົນແລັບ (Waiting Lab)": "Waiting Lab",
        "ນອນຕິດຕາມ (Admit / IPD)": "Admit",
        "ສົ່ງຕໍ່ (Transfer)": "Transfer",
        "ກວດສຳເລັດ / ກັບບ້ານ": "Completed"
    };
    let mainStatus = statusMap[ds] || "Pharmacy";

    const { error: updateError } = await supabaseClient.from('Visits').update({
        Status: mainStatus, Symptoms: $('#emrCC').text(), Diagnosis: dx,
        Prescription_JSON: presJson, Doctor_Name: docName,
        Visit_Type: $('#emrDeptType').val() || 'OPD', Site: $('#emrSite').val() || 'In-site',
        Physical_Exam: $('#emrPE').val() || '', Advice: $('#emrAdvice').val() || '', Follow_Up: $('#emrFollowup').val() || '',
        Services_List: $('#emrService').val() ? $('#emrService').val().join(', ') : '',
        Mapped_Specialist: $('#emrSpecialist').val() || '', Revenue_Group: $('#emrRevenue').val() || '',
        Lab_Orders_JSON: labJson, Discharge_Status: ds || ''
    }).eq('Visit_ID', visitId);

    if (updateError) {
        Swal.fire('ຜິດພາດ!', updateError.message, 'error');
    } else {
        $('#emrModal').modal('hide');
        window.loadQueue();
        window.logAction('Save', 'EMR saved - Visit ' + visitId, 'OPD');
        Swal.fire({ title: 'ສຳເລັດ!', text: 'ບັນທຶກແລ້ວ', icon: 'success', timer: 1500, showConfirmButton: false });
    }
};

window.printOPDCard = async function (s, i) {
    let v = (s === 'triage') ? currentTriageData[i] : queueDataStore[i];
    if (!v) return;

    Swal.fire({ title: 'ກຳລັງສ້າງໃບ OPD...', didOpen: () => Swal.showLoading() });

    try {
        const { data: d, error } = await supabaseClient
            .from('Patients')
            .select('*')
            .eq('Patient_ID', v.patientId)
            .single();

        if (error || !d) {
            Swal.close();
            return Swal.fire('ຜິດພາດ', 'ບໍ່ພົບຂໍ້ມູນຄົນເຈັບໃນລະບົບ', 'error');
        }

        let h1 = document.getElementById('print-opd-header-1');
        let h2 = document.getElementById('print-opd-header-2');
        let f1 = document.getElementById('print-opd-footer-1');
        let f2 = document.getElementById('print-opd-footer-2');

        if (systemSettings.opdHeaderUrl) {
            if (h1) { h1.src = systemSettings.opdHeaderUrl; h1.style.display = 'block'; }
            if (h2) { h2.src = systemSettings.opdHeaderUrl; h2.style.display = 'block'; }
        } else {
            if (h1) { h1.src = ''; h1.setAttribute('src', ''); h1.style.display = 'none'; }
            if (h2) { h2.src = ''; h2.setAttribute('src', ''); h2.style.display = 'none'; }
        }

        if (systemSettings.opdFooterUrl) {
            if (f1) { f1.src = systemSettings.opdFooterUrl; f1.style.display = 'block'; }
            if (f2) { f2.src = systemSettings.opdFooterUrl; f2.style.display = 'block'; }
        } else {
            if (f1) { f1.src = ''; f1.setAttribute('src', ''); f1.style.display = 'none'; }
            if (f2) { f2.src = ''; f2.setAttribute('src', ''); f2.style.display = 'none'; }
        }

        let safeSetText = function (id, text) {
            let el = document.getElementById(id);
            if (el) el.innerText = text;
        };

        safeSetText('popd_cn', d.Patient_ID || "-");
        safeSetText('popd_orgid', d.Organization_ID || "-");
        safeSetText('popd_orgname', d.Name_Org || "-");
        safeSetText('popd_datetime', `${v.date || window.getLocalStr(new Date())} ${v.time || ""}`);
        safeSetText('popd_vid', v.visitId || "-");
        safeSetText('popd_dept', v.department || "-");
        safeSetText('popd_name', d.First_Name || "");
        safeSetText('popd_surname', d.Last_Name || "");
        safeSetText('popd_age', d.Age || "-");
        safeSetText('popd_dob', d.Date_of_Birth || "-");
        safeSetText('popd_gender', d.Gender || "-");
        safeSetText('popd_nation', d.Nationality || "-");
        safeSetText('popd_job', d.Occupation || "-");
        safeSetText('popd_village', d.Address || "-");
        safeSetText('popd_district', d.District || "-");
        safeSetText('popd_prov', d.Province || "-");
        safeSetText('popd_phone', d.Phone_Number || "-");
        safeSetText('popd_cc', v.symptoms || "-");
        safeSetText('popd_temp', v.temp ? v.temp + " °C" : "-");
        safeSetText('popd_bp', v.bp || "-");
        safeSetText('popd_pr', v.pulse || "-");
        safeSetText('popd_spo2', v.spo2 ? v.spo2 + " %" : "-");
        safeSetText('popd_w', v.weight ? v.weight + " kg" : "-");
        safeSetText('popd_h', v.height ? v.height + " cm" : "-");

        let bmiText = "-";
        if (v.weight && v.height) {
            let w = parseFloat(v.weight), h = parseFloat(v.height);
            if (w > 0 && h > 0) {
                let bmi = w / Math.pow(h / 100, 2);
                bmiText = bmi.toFixed(1) + (bmi >= 25 ? " (ຕຸ້ຍ)" : bmi < 18.5 ? " (ຈ່ອຍ)" : " (ປົກກະຕິ)");
            }
        }

        safeSetText('popd_bmi', bmiText);
        safeSetText('popd_allergy', d.Drug_Allergy || "ບໍ່ມີ");
        safeSetText('popd_disease', d.Underlying_Disease || "ບໍ່ມີ");

        Swal.close();
        window.executePrint('opd-print-area');

    } catch (err) {
        Swal.close();
        console.error("Print Error:", err);
        Swal.fire('ຂໍ້ຜິດພາດ', 'ເກີດບັນຫາໃນການຈັດກຽມເອກະສານ', 'error');
    }
};

window.openApptModal = function () {
    $('#apptForm')[0].reset();
    $('#typePatient').prop('checked', true);
    window.toggleApptCustomerType();
    $('#a_id').val('');
    $('#a_date').val(window.getLocalStr(new Date()));
    if (document.activeElement) document.activeElement.blur();
    $('#apptModal').modal('show');
};

window.toggleApptCustomerType = function () {
    let o = $('#typeOrg').is(':checked');
    if (o) {
        $('#boxApptPatient').hide();
        $('#boxApptOrg').show();
        $('#a_patient').val(null).trigger('change');
    } else {
        $('#boxApptPatient').show();
        $('#boxApptOrg').hide();
        $('#a_org').val(null).trigger('change');
    }
    $('#a_target_id, #a_target_name').val('');
};

window.openPatientVacModal = function () {
    $('#patientVacForm')[0].reset();
    $('#pv_patient').val(null).trigger('change');
    $('#pv_date').val(window.getLocalStr(new Date()));
    if (document.activeElement) document.activeElement.blur();
    $('#patientVacModal').modal('show');
};

window.delPatient = async function (id) {
    let r = await Swal.fire({ title: 'ລຶບ?', icon: 'warning', showCancelButton: true, confirmButtonText: 'ລຶບ' });
    if (r.isConfirmed) {
        const { error } = await supabaseClient.from('Patients').delete().eq('Patient_ID', id);
        if (error) {
            Swal.fire('Error', error.message, 'error');
        } else {
            window.initPatientTable();
            window.logAction('Delete', `ລຶບຄົນເຈັບ: ${id}`, 'Patients');
            Swal.fire('ສຳເລັດ!', 'ລຶບຂໍ້ມູນຄົນເຈັບແລ້ວ', 'success');
        }
    }
};

window.loadAppointments = async function () {
    if ($.fn.DataTable.isDataTable('#apptTable')) $('#apptTable').DataTable().destroy();
    $('#apptTable tbody').html('<tr><td colspan="8" class="text-center py-4"><div class="spinner-border text-warning spinner-border-sm"></div> ກຳລັງໂຫຼດ...</td></tr>');

    try {
        const { data: r, error } = await supabaseClient.from('Appointments').select('*').order('Appt_Date', { ascending: true });

        if (error) {
            console.error('Error:', error);
            Swal.fire('Error', error.message, 'error');
            return;
        }

        if ($.fn.DataTable.isDataTable('#apptTable')) $('#apptTable').DataTable().destroy();
        let h = '';
        if (r && r.length > 0) {
            r.forEach(a => {
                let bs = '', ac = '';
                if (a.Status === 'Completed') {
                    bs = '<span class="badge bg-success">ສຳເລັດ</span>';
                    ac = `<button class="btn btn-sm btn-outline-danger shadow-sm" onclick="window.deleteAppt('${a.Appt_ID}')"><i class="fas fa-trash"></i></button>`;
                } else if (a.Status === 'Cancelled') {
                    bs = '<span class="badge bg-secondary">ຍົກເລີກ</span>';
                    ac = `<button class="btn btn-sm btn-outline-danger shadow-sm" onclick="window.deleteAppt('${a.Appt_ID}')"><i class="fas fa-trash"></i></button>`;
                } else if (a.Status === 'Missed') {
                    bs = '<span class="badge bg-dark">ຂາດນັດ</span>';
                    ac = `<button class="btn btn-sm btn-outline-danger shadow-sm" onclick="window.deleteAppt('${a.Appt_ID}')"><i class="fas fa-trash"></i></button>`;
                } else if (a.Status === 'Overdue') {
                    bs = '<span class="badge bg-danger">ກາຍກຳນົດ!</span>';
                    ac = `<button class="btn btn-sm btn-success shadow-sm me-1" onclick="window.completeAppt('${a.Appt_ID}')"><i class="fas fa-check"></i></button><button class="btn btn-sm btn-dark shadow-sm me-1" onclick="window.missedAppt('${a.Appt_ID}')"><i class="fas fa-user-slash"></i></button><button class="btn btn-sm btn-secondary shadow-sm" onclick="window.cancelAppt('${a.Appt_ID}')"><i class="fas fa-times"></i></button>`;
                } else {
                    bs = '<span class="badge bg-warning text-dark">ລໍຖ້າ</span>';
                    ac = `<button class="btn btn-sm btn-success shadow-sm me-1" onclick="window.completeAppt('${a.Appt_ID}')"><i class="fas fa-check"></i></button><button class="btn btn-sm btn-dark shadow-sm me-1" onclick="window.missedAppt('${a.Appt_ID}')"><i class="fas fa-user-slash"></i></button><button class="btn btn-sm btn-secondary shadow-sm" onclick="window.cancelAppt('${a.Appt_ID}')"><i class="fas fa-times"></i></button>`;
                }
                let ty = a.Type === 'Vaccine' ? '<span class="badge border border-success text-success"><i class="fas fa-syringe"></i> ວັກຊີນ</span>' : '<span class="badge border border-primary text-primary"><i class="fas fa-stethoscope"></i> ທົ່ວໄປ</span>';
                let pt = (a.Patient_ID && a.Patient_ID.startsWith('ORG')) ? `<i class="fas fa-building text-success me-1"></i> ${a.Patient_Name}` : `<i class="fas fa-user text-primary me-1"></i> ${a.Patient_Name}`;
                let rowColor = a.Status === 'Overdue' ? 'text-danger' : 'text-primary';

                h += `<tr>
                        <td class="${rowColor} fw-bold">${a.Appt_Date || ''}</td>
                        <td>${a.Appt_Time || ''}</td>
                        <td>${ty}</td>
                        <td class="fw-bold">${pt}</td>
                        <td>${a.Reason || ''}</td>
                        <td class="text-muted small">${a.Doctor || ''}</td>
                        <td>${bs}</td>
                        <td class="text-center"><div class="d-flex justify-content-center">${ac}</div></td>
                      </tr>`;
            });
        }
        $('#apptTable tbody').html(h);
        $('#apptTable').DataTable({ responsive: true, pageLength: 10, order: [[0, "asc"]], language: { search: "ຄົ້ນຫາ:", emptyTable: "ຍັງບໍ່ມີຂໍ້ມູນ" } });
    } catch (err) {
        console.error('Error:', err);
        Swal.fire('Error', err.message, 'error');
    }
};

window.submitApptForm = async function (e) {
    if (e) e.preventDefault();
    if (!$('#a_target_id').val()) return Swal.fire('ແຈ້ງເຕືອນ', 'ເລືອກລູກຄ້າກ່ອນ', 'warning');
    Swal.fire({ title: 'ກຳລັງບັນທຶກ...', didOpen: () => Swal.showLoading() });

    let isEdit = $('#a_id').val() !== '';
    let row = {
        Target_ID: $('#a_target_id').val(), Target_Name: $('#a_target_name').val(),
        Appt_Date: $('#a_date').val(), Appt_Time: $('#a_time').val(),
        Type: $('#a_type').val(), Reason: $('#a_reason').val(), Doctor_Name: $('#a_doctor').val()
    };

    let res;
    if (isEdit) {
        res = await supabaseClient.from('Appointments').update(row).eq('Appt_ID', $('#a_id').val());
    } else {
        res = await supabaseClient.from('Appointments').insert(row);
    }

    if (res.error) {
        Swal.fire('ຜິດພາດ!', res.error.message, 'error');
    } else {
        $('#apptModal').modal('hide');
        window.loadAppointments();
        window.checkAlerts();
        Swal.fire('ສຳເລັດ!', 'ບັນທຶກແລ້ວ', 'success');
    }
};

window.deleteAppt = async function (id) {
    let r = await Swal.fire({ title: 'ລຶບ?', icon: 'warning', showCancelButton: true, confirmButtonText: 'ລຶບ' });
    if (r.isConfirmed) {
        await supabaseClient.from('Appointments').delete().eq('Appt_ID', id);
        window.loadAppointments();
    }
};

window.completeAppt = async function (id) {
    await supabaseClient.from('Appointments').update({ Status: 'Completed' }).eq('Appt_ID', id);
    window.loadAppointments();
    window.checkAlerts();
};

window.missedAppt = async function (id) {
    let r = await Swal.fire({ title: 'ຂາດນັດ?', icon: 'warning', showCancelButton: true, confirmButtonText: 'ຢືນຍັນ' });
    if (r.isConfirmed) {
        await supabaseClient.from('Appointments').update({ Status: 'Missed' }).eq('Appt_ID', id);
        window.loadAppointments();
        window.checkAlerts();
    }
};

window.cancelAppt = async function (id) {
    let r = await Swal.fire({ title: 'ຍົກເລີກ?', icon: 'warning', showCancelButton: true, confirmButtonText: 'ຍົກເລີກນັດ' });
    if (r.isConfirmed) {
        await supabaseClient.from('Appointments').update({ Status: 'Cancelled' }).eq('Appt_ID', id);
        window.loadAppointments();
        window.checkAlerts();
    }
};

window.generateIntervalInputs = function () {
    let d = parseInt($('#v_doses').val()) || 1;
    let v = ($('#v_interval_hidden').val() || "").split(',');
    let c = $('#intervalInputsContainer');
    c.empty();
    if (d <= 1) {
        c.hide();
        return;
    }
    c.css('display', 'flex').html('<p class="w-100 mb-2 text-info fw-bold small"><i class="fas fa-clock me-1"></i> ກຳນົດໄລຍະຫ່າງການສັກ (ມື້):</p>');
    for (let i = 1; i < d; i++) {
        c.append(`<div class="col-6 mb-2"><label class="form-label small text-muted mb-1">ເຂັມ ${i} -> ເຂັມ ${i + 1}</label><input type="number" class="form-control form-control-sm border-info interval-input" value="${v[i - 1] ? v[i - 1].trim() : '0'}" required min="0"></div>`);
    }
};

window.loadVaccineMaster = async function () {
    if ($.fn.DataTable.isDataTable('#vacMasterTable')) $('#vacMasterTable').DataTable().destroy();
    $('#vacMasterTable tbody').html('<tr><td colspan="6" class="text-center py-4"><div class="spinner-border text-primary spinner-border-sm"></div> ກຳລັງໂຫຼດ...</td></tr>');

    try {
        const { data: r, error } = await supabaseClient.from('Vaccines_Master').select('*');
        if (error) { console.error('Error:', error); Swal.fire('Error', error.message, 'error'); return; }

        vaccinesMasterList = r || [];
        if ($.fn.DataTable.isDataTable('#vacMasterTable')) $('#vacMasterTable').DataTable().destroy();

        let o = '<option value="">-- ເລືອກວັກຊີນ --</option>';
        if (r) r.forEach(v => o += `<option value="${v.Vaccine_Name}">${v.Vaccine_Name}</option>`);
        $('#pv_vaccine').html(o);

        let h = '';
        if (r && r.length > 0) {
            r.forEach(v => {
                let intv = (!v.Interval_Days || v.Interval_Days === "0") ? '<span class="text-muted">-</span>' : `<span class="text-info small">${v.Interval_Days.toString().split(',').join(' ມື້, ')} ມື້</span>`;
                h += `<tr>
                        <td class="text-center"><input type="checkbox" class="form-check-input bulk-check-vacMaster" value="${v.Vac_ID}"></td>
                        <td class="fw-bold text-primary">${v.Vaccine_Name}</td>
                        <td>${v.Disease}</td>
                        <td><span class="badge bg-secondary rounded-pill px-3">${v.Total_Doses} ໂດສ</span></td>
                        <td>${intv}</td>
                        <td class="text-center">
                            <button class="btn btn-sm btn-primary shadow-sm me-1" onclick="window.editVacMaster('${v.Vac_ID}','${v.Vaccine_Name}','${v.Disease}','${v.Total_Doses}','${v.Interval_Days}')"><i class="fas fa-edit"></i></button>
                            <button class="btn btn-sm btn-danger shadow-sm" onclick="window.delVacMaster('${v.Vac_ID}')"><i class="fas fa-trash"></i></button>
                        </td>
                      </tr>`;
            });
        }
        $('#vacMasterTable tbody').html(h);
        $('#vacMasterTable').DataTable({ responsive: true, pageLength: 10, language: { search: "ຄົ້ນຫາ:", emptyTable: "ບໍ່ມີຂໍ້ມູນ" } });
    } catch (err) {
        console.error('Error:', err);
        Swal.fire('Error', err.message, 'error');
    }
};

window.loadPatientVaccines = async function () {
    if ($.fn.DataTable.isDataTable('#patientVacTable')) $('#patientVacTable').DataTable().destroy();
    $('#patientVacTable tbody').html('<tr><td colspan="9" class="text-center py-4"><div class="spinner-border text-success spinner-border-sm"></div> ກຳລັງໂຫຼດ...</td></tr>');

    try {
        const { data: r, error } = await supabaseClient.from('Patient_Vaccines').select('*').order('Date_Given', { ascending: false });
        if (error) { console.error('Error:', error); Swal.fire('Error', error.message, 'error'); return; }

        if ($.fn.DataTable.isDataTable('#patientVacTable')) $('#patientVacTable').DataTable().destroy();
        let h = '';
        let tStr = window.getLocalStr(new Date());
        if (r && r.length > 0) {
            r.forEach(x => {
                let nextDue = x.Next_Appointment_Date || x.Next_Due_Date || "-";
                let nd = nextDue !== "-" ? `<span class="text-info fw-bold"><i class="far fa-calendar-alt"></i> ${nextDue}</span>` : '<span class="text-muted">-</span>';
                let vm = vaccinesMasterList.find(v => v.Vaccine_Name === x.Vaccine_Name || v.name === x.Vaccine_Name);
                let td = vm ? parseInt(vm.Total_Doses || vm.doses) : 1;
                let cd = parseInt(x.Dose_Number) || 1;
                let sb = "";

                if (cd >= td) {
                    sb = '<span class="badge bg-success">ສຳເລັດ (ຄົບໂດສ)</span>';
                } else {
                    if (nextDue < tStr && nextDue !== "-") {
                        sb = `<span class="badge bg-danger">ກາຍກຳນົດເຂັມ ${cd + 1}!</span>`;
                    } else {
                        sb = `<span class="badge bg-warning text-dark">ລໍຖ້າເຂັມ ${cd + 1}</span>`;
                    }
                }

                h += `<tr>
                        <td>${x.Date_Given || ''}</td>
                        <td class="fw-bold">${x.Patient_Name || ''}</td>
                        <td class="text-success fw-bold">${x.Vaccine_Name || ''}</td>
                        <td><span class="text-primary fw-bold">${x.Lot_Number || '-'}</span></td>
                        <td><span class="badge bg-primary rounded-pill">ໂດສທີ ${cd}/${td}</span></td>
                        <td class="text-muted small">${x.Given_By || ''}</td>
                        <td>${nd}</td>
                        <td class="text-center">${sb}</td>
                        <td class="text-center">
                            <button class="btn btn-sm btn-info text-white shadow-sm me-1" onclick="window.printVacCard('${x.Patient_ID}','${x.Patient_Name}','${x.Vaccine_Name}','${cd}/${td}','${x.Date_Given}','${nextDue}')"><i class="fas fa-print"></i></button>
                            <button class="btn btn-sm btn-outline-danger shadow-sm" onclick="window.delPatientVac('${x.Record_ID}')"><i class="fas fa-trash"></i></button>
                        </td>
                      </tr>`;
            });
        }
        $('#patientVacTable tbody').html(h);
        $('#patientVacTable').DataTable({ 
            responsive: true, 
            pageLength: 10, 
            order: [[0, "desc"], [1, "desc"]], 
            language: { 
                search: "ຄົ້ນຫາ:", 
                lengthMenu: "ສະແດງ _MENU_", 
                info: "ສະແດງ _START_ ຫາ _END_ ຈາກ _TOTAL_ ລາຍການ", 
                paginate: { previous: "ກ່ອນໜ້າ", next: "ຕໍ່ໄປ" },
                emptyTable: "ບໍ່ມີຂໍ້ມູນ"
            } 
        });
    } catch (err) {
        console.error('Error:', err);
        Swal.fire('Error', err.message, 'error');
    }
};

window.openVacMasterModal = function () {
    $('#vacMasterForm')[0].reset();
    $('#v_id, #v_interval_hidden').val('');
    window.generateIntervalInputs();
    $('#vacMasterModal').modal('show');
};

window.editVacMaster = function (id, n, d, ds, i) {
    $('#v_id').val(id);
    $('#v_name').val(n);
    $('#v_disease').val(d);
    $('#v_doses').val(ds);
    $('#v_interval_hidden').val(i || '');
    window.generateIntervalInputs();
    $('#vacMasterModal').modal('show');
};

window.delVacMaster = async function (id) {
    let r = await Swal.fire({ title: 'ລຶບ?', icon: 'warning', showCancelButton: true, confirmButtonText: 'ລຶບ' });
    if (r.isConfirmed) {
        const { error } = await supabaseClient.from('Vaccines_Master').delete().eq('Vac_ID', id);
        if (error) {
            Swal.fire('Error', error.message, 'error');
        } else {
            window.loadVaccineMaster();
            Swal.fire('ສຳເລັດ!', 'ລຶບວັກຊີນແລ້ວ', 'success');
        }
    }
};

window.submitVacMasterForm = async function (e) {
    if (e) e.preventDefault();
    Swal.fire({ title: 'ກຳລັງບັນທຶກ...', didOpen: () => Swal.showLoading() });
    let a = [];
    $('.interval-input').each(function () { a.push($(this).val()); });

    let isEdit = $('#v_id').val() !== '';
    let row = {
        Vaccine_Name: $('#v_name').val(), Disease_Target: $('#v_disease').val(),
        Total_Doses: parseInt($('#v_doses').val()) || 1, Dose_Interval: a.join(',')
    };

    if (isEdit) {
        const { error } = await supabaseClient.from('Vaccines_Master').update(row).eq('Vac_ID', $('#v_id').val());
        if (error) return Swal.fire('Error', error.message, 'error');
    } else {
        const { error } = await supabaseClient.from('Vaccines_Master').insert(row);
        if (error) return Swal.fire('Error', error.message, 'error');
    }

    $('#vacMasterModal').modal('hide');
    window.loadVaccineMaster();
    Swal.fire('ສຳເລັດ!', '', 'success');
};

window.calculateNextVacDate = function () {
    let vn = $('#pv_vaccine').val();
    let dg = $('#pv_date').val();
    let cd = parseInt($('#pv_dose').val()) || 1;
    if (!vn || !dg) return;

    let vm = vaccinesMasterList.find(x => x.name === vn);
    if (!vm) return;

    let md = parseInt(vm.doses) || 1;
    let iv = (vm.interval || "0").toString().split(',');

    if (cd >= md || iv.length === 0 || iv[0] === "0") {
        $('#pv_next_date').val('').prop('disabled', true);
        $('#pv_auto_appt').prop('checked', false).prop('disabled', true);
    } else {
        let ad = parseInt(iv[cd - 1]);
        if (isNaN(ad) || ad <= 0) {
            $('#pv_next_date').val('').prop('disabled', true);
        } else {
            $('#pv_next_date').prop('disabled', false);
            $('#pv_auto_appt').prop('disabled', false).prop('checked', true);
            let g = new Date(dg);
            g.setDate(g.getDate() + ad);
            $('#pv_next_date').val(window.getLocalStr(g));
        }
    }
};

window.printVacCard = function (id, n, vn, ds, dg, nd) {
    Swal.fire({ title: 'ກຳລັງສ້າງບັດ...', didOpen: () => Swal.showLoading() });
    $('#pVacHospName').text(systemSettings.hospitalName || "Clinic");
    $('#pVacCN').text(id);
    $('#pVacName').text(n);
    $('#pVacTitle').text(vn);
    $('#pVacDose').text(ds);
    $('#pVacGiven').text(dg);

    if (!nd || nd === "-") {
        $('#pVacNextDate').text("ສຳເລັດ (ຄົບໂດສ)").css('color', '#10b981');
    } else {
        $('#pVacNextDate').text(nd).css('color', '#dc2626');
    }
    Swal.close();
    window.executePrint('vac-print-area');
};



window.submitPatientVacForm = async function (e) {
    if (e) e.preventDefault();
    Swal.fire({ title: 'ກຳລັງບັນທຶກ...', didOpen: () => Swal.showLoading() });

    let pid = $('#pv_patient_id').val();
    let pname = $('#pv_patient_name').val();
    let vac = $('#pv_vaccine').val();
    let ndate = $('#pv_next_date').val();
    let autoAppt = $('#pv_auto_appt').is(':checked');

    const { error } = await supabaseClient.from('Patient_Vaccines').insert({
        Patient_ID: pid, Patient_Name: pname, Vaccine_Name: vac,
        Dose_Number: parseInt($('#pv_dose').val()) || 1, Lot_Number: $('#pv_lot').val(),
        Date_Given: $('#pv_date').val(), Next_Appointment_Date: ndate,
        Given_By: $('#pv_doctor').val()
    });

    if (error) {
        Swal.fire('ຜິດພາດ!', error.message, 'error');
        return;
    }

    if (autoAppt && ndate) {
        await supabaseClient.from('Appointments').insert({
            Target_ID: pid, Target_Name: pname, Type: 'Vaccine',
            Appt_Date: ndate, Appt_Time: '09:00',
            Reason: 'ນັດໝາຍວັກຊີນ: ' + vac, Doctor_Name: 'System', Status: 'Waiting'
        });
    }

    $('#patientVacModal').modal('hide');
    window.loadPatientVaccines();
    window.checkAlerts();
    window.loadAppointments();
    Swal.fire('ສຳເລັດ!', '', 'success');
};

window.delPatientVac = async function (id) {
    let r = await Swal.fire({ title: 'ລຶບປະຫວັດ?', icon: 'warning', showCancelButton: true, confirmButtonText: 'ລຶບ' });
    if (r.isConfirmed) {
        await supabaseClient.from('Patient_Vaccines').delete().eq('Record_ID', id);
        window.loadPatientVaccines();
    }
};

window.loadDrugsMaster = async function () {
    if ($.fn.DataTable.isDataTable('#drugTable')) $('#drugTable').DataTable().destroy();
    $('#drugTable tbody').html('<tr><td colspan="4" class="text-center py-4"><div class="spinner-border text-success spinner-border-sm"></div> ກຳລັງໂຫຼດ...</td></tr>');

    try {
        const { data: r, error } = await supabaseClient.from('Drugs_Master').select('*');

        // 🚨 ດັກຈັບ Error ຕາມ Antigravity ສັ່ງ
        if (error) {
            console.error('Error:', error);
            Swal.fire('Error', error.message, 'error');
            return;
        }

        if ($.fn.DataTable.isDataTable('#drugTable')) $('#drugTable').DataTable().destroy();
        let h = '';
        if (r && r.length > 0) {
            r.forEach(x => {
                // ອີງຕາມ Column ໃນ CSV: Drug_ID, Drug_Name, Description
                h += `<tr>
                        <td class="text-center"><input type="checkbox" class="form-check-input bulk-check-drugs" value="${x.Drug_ID}"></td>
                        <td class="fw-bold text-success">${x.Drug_Name}</td>
                        <td>${x.Description}</td>
                        <td class="text-center">
                            <button class="btn btn-sm btn-primary shadow-sm me-1" onclick="window.editDrugMaster('${x.Drug_ID}','${x.Drug_Name}','${x.Description}')"><i class="fas fa-edit"></i></button>
                            <button class="btn btn-sm btn-danger shadow-sm" onclick="window.delDrugMaster('${x.Drug_ID}')"><i class="fas fa-trash"></i></button>
                        </td>
                      </tr>`;
            });
        }
        $('#drugTable tbody').html(h);
        $('#drugTable').DataTable({ responsive: true, pageLength: 10, language: { search: "ຄົ້ນຫາ:", emptyTable: "ບໍ່ມີຂໍ້ມູນ" } });
    } catch (err) {
        console.error('System Error:', err);
        Swal.fire('Error', err.message, 'error');
    }
};

window.openDrugMasterModal = function () {
    $('#drugMasterForm')[0].reset();
    $('#dr_id').val('');
    $('#drugMasterModal').modal('show');
};

window.editDrugMaster = function (id, n, d) {
    $('#dr_id').val(id);
    $('#dr_name').val(n);
    $('#dr_desc').val(d);
    $('#drugMasterModal').modal('show');
};

window.delDrugMaster = async function (id) {
    let r = await Swal.fire({ title: 'ລຶບ?', icon: 'warning', showCancelButton: true, confirmButtonText: 'ລຶບ' });
    if (r.isConfirmed) {
        await supabaseClient.from('Drugs_Master').delete().eq('Drug_ID', id);
        window.loadDrugsMaster();
    }
};

window.submitDrugMasterForm = async function (e) {
    if (e) e.preventDefault();
    Swal.fire({ title: 'ກຳລັງບັນທຶກ...', didOpen: () => Swal.showLoading() });

    let isEdit = $('#dr_id').val() !== '';
    let row = {
        Drug_Name: $('#dr_name').val(), Description: $('#dr_desc').val()
    };

    if (isEdit) {
        await supabaseClient.from('Drugs_Master').update(row).eq('Drug_ID', $('#dr_id').val());
    } else {
        await supabaseClient.from('Drugs_Master').insert(row);
    }

    $('#drugMasterModal').modal('hide');
    window.loadDrugsMaster();
    window.preloadDropdownData();
    Swal.fire('ສຳເລັດ!', '', 'success');
};

window.loadLabsMaster = async function () {
    if ($.fn.DataTable.isDataTable('#labTable')) $('#labTable').DataTable().destroy();
    $('#labTable tbody').html('<tr><td colspan="4" class="text-center py-4"><div class="spinner-border text-primary spinner-border-sm"></div> ກຳລັງໂຫຼດ...</td></tr>');

    try {
        const { data: r, error } = await supabaseClient.from('Labs_Master').select('*');

        if (error) {
            console.error('Error:', error);
            Swal.fire('Error', error.message, 'error');
            return;
        }

        if ($.fn.DataTable.isDataTable('#labTable')) $('#labTable').DataTable().destroy();
        let h = '';
        if (r && r.length > 0) {
            r.forEach(x => {
                // ອີງຕາມ Column ໃນ CSV: Lab_ID, Lab_Name, Description
                h += `<tr>
                        <td class="text-center"><input type="checkbox" class="form-check-input bulk-check-labs" value="${x.Lab_ID}"></td>
                        <td class="fw-bold text-primary">${x.Lab_Name}</td>
                        <td>${x.Description}</td>
                        <td class="text-center">
                            <button class="btn btn-sm btn-primary shadow-sm me-1" onclick="window.editLabMaster('${x.Lab_ID}','${x.Lab_Name}','${x.Description}')"><i class="fas fa-edit"></i></button>
                            <button class="btn btn-sm btn-danger shadow-sm" onclick="window.delLabMaster('${x.Lab_ID}')"><i class="fas fa-trash"></i></button>
                        </td>
                      </tr>`;
            });
        }
        $('#labTable tbody').html(h);
        $('#labTable').DataTable({ responsive: true, pageLength: 10, language: { search: "ຄົ້ນຫາ:", emptyTable: "ບໍ່ມີຂໍ້ມູນ" } });
    } catch (err) {
        console.error('System Error:', err);
        Swal.fire('Error', err.message, 'error');
    }
};

window.openLabMasterModal = function () {
    $('#labMasterForm')[0].reset();
    $('#lb_id').val('');
    $('#labMasterModal').modal('show');
};

window.editLabMaster = function (id, n, d) {
    $('#lb_id').val(id);
    $('#lb_name').val(n);
    $('#lb_desc').val(d);
    $('#labMasterModal').modal('show');
};

window.delLabMaster = async function (id) {
    let r = await Swal.fire({ title: 'ລຶບ?', icon: 'warning', showCancelButton: true, confirmButtonText: 'ລຶບ' });
    if (r.isConfirmed) {
        await supabaseClient.from('Labs_Master').delete().eq('Lab_ID', id);
        window.loadLabsMaster();
    }
};

window.submitLabMasterForm = async function (e) {
    if (e) e.preventDefault();
    Swal.fire({ title: 'ກຳລັງບັນທຶກ...', didOpen: () => Swal.showLoading() });

    let isEdit = $('#lb_id').val() !== '';
    let row = {
        Lab_Name: $('#lb_name').val(), Description: $('#lb_desc').val()
    };

    if (isEdit) {
        await supabaseClient.from('Labs_Master').update(row).eq('Lab_ID', $('#lb_id').val());
    } else {
        await supabaseClient.from('Labs_Master').insert(row);
    }

    $('#labMasterModal').modal('hide');
    window.loadLabsMaster();
    window.preloadDropdownData();
    Swal.fire('ສຳເລັດ!', '', 'success');
};

window.loadUsers = async function () {
    if ($.fn.DataTable.isDataTable('#userTable')) $('#userTable').DataTable().destroy();
    $('#userTable tbody').html('<tr><td colspan="6" class="text-center py-4"><div class="spinner-border text-dark spinner-border-sm"></div> ກຳລັງໂຫຼດ...</td></tr>');

    try {
        const { data: u, error } = await supabaseClient.from('Users').select('*');

        // 🚨 ດັກຈັບ Error ຕາມ Antigravity
        if (error) {
            console.error('Error:', error);
            Swal.fire('Error', error.message, 'error');
            return;
        }

        if ($.fn.DataTable.isDataTable('#userTable')) $('#userTable').DataTable().destroy();
        let h = '';
        if (u && u.length > 0) {
            u.forEach(x => {
                let sb = x.Status === 'active' ? '<span class="badge bg-success rounded-pill px-3">ເປີດໃຊ້ງານ</span>' : '<span class="badge bg-danger rounded-pill px-3">ປິດໃຊ້ງານ</span>';
                h += `<tr>
                        <td class="text-center"><input type="checkbox" class="form-check-input bulk-check-users" value="${x.ID}"></td>
                        <td class="fw-bold">${x.Name}</td>
                        <td class="text-muted">${x.Email}</td>
                        <td><span class="badge ${x.Role === 'admin' ? 'bg-primary' : 'bg-secondary'} rounded-pill px-3 text-uppercase">${x.Role}</span></td>
                        <td>${sb}</td>
                        <td class="text-center">
                            <button class="btn btn-sm btn-primary shadow-sm me-1" onclick="window.openEditUserModal('${x.ID}','${x.Name}','${x.Email}','${x.Role}','${x.Permissions || ''}')"><i class="fas fa-edit"></i></button>
                            <button class="btn btn-sm btn-danger shadow-sm rounded" onclick="window.deleteUserRow('${x.ID}')"><i class="fas fa-trash"></i></button>
                        </td>
                      </tr>`;
            });
        }
        $('#userTable tbody').html(h);
        $('#userTable').DataTable({ responsive: true, pageLength: 10, language: { search: "ຄົ້ນຫາ:", emptyTable: "ບໍ່ມີຂໍ້ມູນ" } });
    } catch (err) {
        console.error('Error:', err);
        Swal.fire('Error', err.message, 'error');
    }
};

window.openAddUserModal = function () {
    $('#addUserForm')[0].reset();
    $('#u_id').val('');
    $('#userModalTitle').html('<i class="fas fa-user-plus me-2"></i>ເພີ່ມຜູ້ໃຊ້ລະບົບ');
    $('#u_pass').prop('required', true);
    window.togglePermissionsBox();
    $('#addUserModal').modal('show');
};

window.openEditUserModal = function (id, n, e, r, p) {
    $('#u_id').val(id);
    $('#u_name').val(n);
    $('#u_email').val(e);
    $('#u_role').val(r);
    $('#u_pass').prop('required', false);
    $('#userModalTitle').html('<i class="fas fa-user-edit me-2"></i>ແກ້ໄຂຜູ້ໃຊ້ລະບົບ');
    let pa = p.split(',');
    $('.permission-check').each(function () {
        $(this).prop('checked', pa.includes($(this).val()));
    });
    window.togglePermissionsBox();
    $('#addUserModal').modal('show');
};

window.togglePermissionsBox = function () {
    $('#permBox').css('display', $('#u_role').val() === 'admin' ? 'none' : 'block');
};

window.deleteUserRow = async function (id) {
    let r = await Swal.fire({ title: 'ລຶບ?', icon: 'warning', showCancelButton: true, confirmButtonColor: '#ef4444', confirmButtonText: 'ລຶບ' });
    if (r.isConfirmed) {
        const { error } = await supabaseClient.from('Users').delete().eq('ID', id);
        if (error) {
            Swal.fire('Error', error.message, 'error');
        } else {
            window.loadUsers();
            window.logAction('Delete', 'Delete User ID: ' + id, 'Users');
            Swal.fire('ສຳເລັດ!', 'ລຶບຜູ້ໃຊ້ແລ້ວ', 'success');
        }
    }
};

window.submitUserForm = async function (e) {
    if (e) e.preventDefault();
    let p = [];
    $('.permission-check:checked').each(function () { p.push($(this).val()); });

    let isEdit = $('#u_id').val() !== '';
    let email = $('#u_email').val();
    let password = $('#u_pass').val();
    let role = $('#u_role').val();
    let perms = role === 'admin' ? 'all' : p.join(',');

    let row = {
        Name: $('#u_name').val(), Email: email, Role: role, Permissions: perms, Status: 'active'
    };

    if (isEdit) {
        let { error } = await supabaseClient.from('Users').update(row).eq('ID', $('#u_id').val());
        if (error) { Swal.fire('Error', error.message, 'error'); return; }
        
        $('#addUserModal').modal('hide');
        window.logAction('Edit', 'Edit User: ' + $('#u_name').val() + ' (' + role + ')', 'Users');
        window.loadUsers();
        Swal.fire('ສຳເລັດ', 'ບັນທຶກແລ້ວ', 'success');
    } else {
        let { error } = await supabaseClient.from('Users').insert(row);
        if (error) { Swal.fire('Error', error.message, 'error'); return; }

        $('#addUserModal').modal('hide');
        window.logAction('Add', 'Add User: ' + $('#u_name').val() + ' (' + role + ')', 'Users');
        window.loadUsers();
        
        Swal.fire({
            title: 'ສຳເລັດ (ຂັ້ນຕອນທີ 1)',
            html: `ສ້າງຂໍ້ມູນຜູ້ໃຊ້ໃນລະບົບແລ້ວ.<br><br><b class="text-danger">ສຳຄັນ:</b> ທ່ານຕ້ອງໄປທີ່ໜ້າຈໍ <b>Supabase Dashboard > Authentication > Add User</b> ເພື່ອສ້າງລະຫັດຜ່ານສຳລັບອີເມວ <b>${email}</b> ນີ້ ກ່ອນທີ່ພະນັກງານຈະເຂົ້າສູ່ລະບົບໄດ້.`,
            icon: 'info'
        });
    }
};

window.fetchOrg = async function () {
    let c = $('#p_org_id').val();
    if (!c) {
        $('#p_org_name, #p_discount_show').val('');
        return;
    }
    const { data, error } = await supabaseClient.from('Organizations').select('*').eq('Org_Code', c).limit(1);
    if (data && data.length > 0) {
        $('#p_org_name').val(data[0].Org_Name);
        $('#p_discount_show').val(data[0].Discount || "ບໍ່ມີສ່ວນຫຼຸດ");
    } else {
        $('#p_org_name').val('❌ ບໍ່ພົບ');
        $('#p_discount_show').val('');
    }
};

window.loadMasterDataGlobalCallback = function (data) {
    masterDataStore = data || {};
    ['Department', 'Shift', 'Title', 'Gender', 'Nationality', 'Occupation', 'BloodType', 'InsCompany', 'Channel', 'Doctor', 'Site', 'PatientType_InSite', 'PatientType_Onsite', 'DrugUnit', 'DrugUsage'].forEach(c => {
        let o = '<option value="">-- ເລືອກ --</option>';
        if (masterDataStore[c]) {
            masterDataStore[c].forEach(i => o += `<option value="${i.value}">${i.value}</option>`);
        }
        $('.dyn-' + c).html(o);
    });
    if (document.getElementById('masterCategory')) window.loadMasterList();
};

window.loadMasterDataGlobal = async function () {
    try {
        const { data, error } = await supabaseClient.from('MasterData').select('*');
        if (error) {
            console.error('Error:', error);
            Swal.fire('Error', error.message, 'error');
            return;
        }

        let formattedMasterData = {};
        if (data) {
            data.forEach(item => {
                if (!formattedMasterData[item.Category]) formattedMasterData[item.Category] = [];
                formattedMasterData[item.Category].push({ id: item.ID, value: item.Value });
            });
        }
        window.loadMasterDataGlobalCallback(formattedMasterData);
    } catch (err) {
        console.error('Error:', err);
    }
};

window.loadMasterList = function () {
    let c = $('#masterCategory').val();
    if (!c) return;
    let h = '';
    if (masterDataStore[c]) {
        masterDataStore[c].forEach(i => {
            h += `<li class="list-group-item d-flex justify-content-between align-items-center border-0 border-bottom mb-1 bg-transparent">
                    <span class="fw-bold text-dark">${i.value}</span> 
                    <button class="btn btn-sm btn-outline-danger shadow-sm rounded-circle" style="width:30px;height:30px;padding:0;" onclick="window.delMaster(${i.id})"><i class="fas fa-trash"></i></button>
                  </li>`;
        });
    }
    $('#masterListUl').html(h);
};

window.addMaster = async function () {
    let c = $('#masterCategory').val();
    let v = $('#newMasterVal').val();
    if (!v) return;
    $('#newMasterVal').val('');

    let { error } = await supabaseClient.from('MasterData').insert({ Category: c, Value: v });
    if (!error) {
        window.loadMasterDataGlobal(); // Reloads and updates UI
    } else {
        Swal.fire('Error', error.message, 'error');
    }
};

window.delMaster = async function (id) {
    let r = await Swal.fire({ title: 'ລຶບ?', icon: 'warning', showCancelButton: true, confirmButtonText: 'ລຶບ' });
    if (r.isConfirmed) {
        const { error } = await supabaseClient.from('MasterData').delete().eq('ID', id);
        if (error) {
            Swal.fire('Error', error.message, 'error');
        } else {
            window.loadMasterDataGlobal();
        }
    }
};

window.loadOrgs = async function () {
    if ($.fn.DataTable.isDataTable('#orgTable')) $('#orgTable').DataTable().destroy();
    $('#orgTable tbody').html('<tr><td colspan="8" class="text-center py-4"><div class="spinner-border text-info spinner-border-sm"></div> ກຳລັງໂຫຼດ...</td></tr>');

    try {
        const { data: orgs, error } = await supabaseClient.from('Organizations').select('*');

        // 🚨 ດັກຈັບ Error ຕາມ Antigravity
        if (error) {
            console.error('Error:', error);
            Swal.fire('Error', error.message, 'error');
            return;
        }

        if ($.fn.DataTable.isDataTable('#orgTable')) $('#orgTable').DataTable().destroy();
        let h = '';
        if (orgs && orgs.length > 0) {
            orgs.forEach(o => {
                let st = o.Status === 'Active' ? `<span class="badge bg-success rounded-pill px-3 shadow-sm" style="cursor:pointer" onclick="window.toggleOrg('${o.Org_ID}','${o.Status}')"><i class="fas fa-check me-1"></i> Active</span>` : `<span class="badge bg-danger rounded-pill px-3 shadow-sm" style="cursor:pointer" onclick="window.toggleOrg('${o.Org_ID}','${o.Status}')"><i class="fas fa-times me-1"></i> Inactive</span>`;
                let sd = o.Discount ? String(o.Discount).replace(/[\r\n]+/g, " ").replace(/'/g, "\\'").replace(/"/g, "&quot;") : "";
                h += `<tr>
                        <td class="text-center"><input type="checkbox" class="form-check-input bulk-check-orgs" value="${o.Org_ID}"></td>
                        <td class="text-primary fw-bold">${o.Org_Code || '-'}</td>
                        <td class="text-muted">${o.Cus_ID_Ex || '-'}</td>
                        <td>${o.Name || '-'}</td>
                        <td class="fw-bold">${o.Org_Name || '-'}</td>
                        <td class="small text-danger" style="max-width: 250px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;" title="${sd || ''}">${sd || '-'}</td>
                        <td>${st}</td>
                        <td class="text-center"><button class="btn btn-sm btn-primary shadow-sm" onclick="window.editOrg('${o.Org_ID}','${o.Cus_ID_Ex}','${o.Name}','${o.Org_Name}','${o.Org_Code}','${sd}')"><i class="fas fa-edit"></i> ແກ້ໄຂ</button></td>
                      </tr>`;
            });
        }
        $('#orgTable tbody').html(h);
        $('#orgTable').DataTable({ responsive: true, pageLength: 10, language: { search: "ຄົ້ນຫາ:", emptyTable: "ບໍ່ມີຂໍ້ມູນ" } });
    } catch (err) {
        console.error('Error:', err);
        Swal.fire('Error', err.message, 'error');
    }
};

window.toggleOrg = async function (id, currentStatus) {
    let newStatus = currentStatus === 'Active' ? 'Inactive' : 'Active';
    await supabaseClient.from('Organizations').update({ Status: newStatus }).eq('Org_ID', id);
    window.loadOrgs();
};

window.submitOrgForm = async function (e) {
    if (e) e.preventDefault();

    let isEdit = $('#o_rowIdx').val() !== '';
    let row = {
        Cus_ID_Ex: $('#o_cusId').val(), Name: $('#o_name').val(),
        Org_Name: $('#o_orgName').val(), Org_Code: $('#o_orgCode').val(),
        Discount: $('#o_discount').val()
    };

    if (isEdit) {
        await supabaseClient.from('Organizations').update(row).eq('Org_ID', $('#o_rowIdx').val());
    } else {
        row.Status = 'Active';
        await supabaseClient.from('Organizations').insert(row);
    }

    $('#orgModal').modal('hide');
    window.loadOrgs();
    window.preloadDropdownData();
};

window.editOrg = function (r, c, n, on, oc, d) {
    $('#o_rowIdx').val(r);
    $('#o_cusId').val(c);
    $('#o_name').val(n);
    $('#o_orgName').val(on);
    $('#o_orgCode').val(oc);
    $('#o_discount').val(d);
    $('#orgModal').modal('show');
};

window.submitSettingsForm = async function (e) {
    if (e) e.preventDefault();
    Swal.fire({ title: 'ກຳລັງບັນທຶກ...', didOpen: () => Swal.showLoading() });

    // ດຶງຄ່າຈາກຟອມໜ້າເວັບ
    let ns = {
        HospitalName: $('#setHospitalName').val(),
        LogoUrl: $('#setLogoUrl').val(),
        OpdHeaderUrl: $('#setOpdHeaderUrl').val(),
        OpdFooterUrl: $('#setOpdFooterUrl').val()
    };

    // ປ່ຽນຮູບແບບໃຫ້ກາຍເປັນ Array ເພື່ອໃຊ້ບັນທຶກລົງ Table ແບບ Upsert
    let updates = Object.keys(ns).map(k => ({
        Key: k,
        Value: ns[k]
    }));

    try {
        const { error } = await supabaseClient
            .from('Settings')
            .upsert(updates, { onConflict: 'Key' });

        // 🚨 ເພີ່ມລະບົບແຈ້ງເຕືອນ Error ຕາມທີ່ Antigravity ແນະນຳ
        if (error) {
            console.error('Error:', error);
            Swal.fire('Error', error.message, 'error');
            return;
        }

        // ອັບເດດຄ່າໃນລະບົບ
        systemSettings = {
            hospitalName: ns.HospitalName,
            logoUrl: ns.LogoUrl,
            opdHeaderUrl: ns.OpdHeaderUrl,
            opdFooterUrl: ns.OpdFooterUrl
        };
        $('#sidebarBrandName').text(ns.HospitalName);
        window.logAction('Save', 'System Settings saved: ' + ns.HospitalName, 'Settings');

        Swal.fire('ສຳເລັດ!', 'ບັນທຶກການຕັ້ງຄ່າແລ້ວ', 'success');

    } catch (err) {
        console.error('System Error:', err);
        Swal.fire('Error', err.message, 'error');
    }
};

// ຟັງຊັນສຳລັບດຶງຄ່າຕັ້ງຄ່າມາສະແດງ
window.loadSettingsData = async function () {
    try {
        const { data, error } = await supabaseClient.from('Settings').select('*');
        if (error) {
            console.error('Error:', error);
            Swal.fire('Error', error.message, 'error');
            return;
        }

        let s = {};
        if (data) {
            data.forEach(row => {
                if (row.Key === 'HospitalName') s.hospitalName = row.Value;
                if (row.Key === 'LogoUrl') s.logoUrl = row.Value;
                if (row.Key === 'OpdHeaderUrl') s.opdHeaderUrl = row.Value;
                if (row.Key === 'OpdFooterUrl') s.opdFooterUrl = row.Value;
            });
        }

        systemSettings = s;
        $('#setHospitalName').val(s.hospitalName || "");
        $('#setLogoUrl').val(s.logoUrl || "");
        $('#setOpdHeaderUrl').val(s.opdHeaderUrl || "");
        $('#setOpdFooterUrl').val(s.opdFooterUrl || "");

        if (typeof window.loadMasterList === 'function') window.loadMasterList();
    } catch (err) {
        console.error('Error:', err);
    }
};

window.loadLocationsMasterView = async function () {
    if ($.fn.DataTable.isDataTable('#locationTable')) $('#locationTable').DataTable().destroy();
    $('#locationTable tbody').html('<tr><td colspan="4" class="text-center py-4"><div class="spinner-border text-info spinner-border-sm"></div> ກຳລັງໂຫຼດ...</td></tr>');

    try {
        const { data: r, error } = await supabaseClient.from('Locations').select('*');

        if (error) {
            console.error('Error:', error);
            Swal.fire('Error', error.message, 'error');
            return;
        }

        locationsDataStore = r || [];
        if ($.fn.DataTable.isDataTable('#locationTable')) $('#locationTable').DataTable().destroy();

        let o = '<option value="">-- ເລືອກເມືອງ --</option>';
        if (r) r.forEach(l => o += `<option value="${l.District}">${l.District}</option>`);
        $('#p_district').html(o);

        let h = '';
        if (r && r.length > 0) {
            r.forEach(l => {
                // ອີງຕາມ Column ໃນ CSV: ID, District, Province
                h += `<tr>
                        <td class="text-center"><input type="checkbox" class="form-check-input bulk-check-locations" value="${l.ID}"></td>
                        <td class="fw-bold text-primary">${l.District}</td>
                        <td><span class="badge bg-info text-dark">${l.Province}</span></td>
                        <td class="text-center">
                            <button class="btn btn-sm btn-primary shadow-sm me-1" onclick="window.editLocation('${l.ID}','${l.District}','${l.Province}')"><i class="fas fa-edit"></i></button>
                            <button class="btn btn-sm btn-danger shadow-sm" onclick="window.delLocation('${l.ID}')"><i class="fas fa-trash"></i></button>
                        </td>
                      </tr>`;
            });
        }
        $('#locationTable tbody').html(h);
        $('#locationTable').DataTable({ responsive: true, pageLength: 10, language: { search: "ຄົ້ນຫາ:", emptyTable: "ບໍ່ມີຂໍ້ມູນ" } });
    } catch (err) {
        console.error('System Error:', err);
        Swal.fire('Error', err.message, 'error');
    }
};

window.openAddLocationModal = function () {
    $('#locationForm')[0].reset();
    $('#l_id').val('');
    $('#locationModal').modal('show');
};

window.editLocation = function (id, d, p) {
    $('#l_id').val(id);
    $('#l_dist').val(d);
    $('#l_prov').val(p);
    $('#locationModal').modal('show');
};

window.delLocation = async function (id) {
    let r = await Swal.fire({ title: 'ລຶບ?', icon: 'warning', showCancelButton: true, confirmButtonText: 'ລຶບ' });
    if (r.isConfirmed) {
        await supabaseClient.from('Locations').delete().eq('ID', id);
        window.loadLocationsMasterView();
    }
};

window.submitLocationForm = async function (e) {
    if (e) e.preventDefault();
    Swal.fire({ title: 'ກຳລັງບັນທຶກ...', didOpen: () => Swal.showLoading() });

    try {
        let payload = {
            District: $('#l_dist').val(),
            Province: $('#l_prov').val()
        };
        let l_id = $('#l_id').val();

        if (l_id) {
            // ແກ້ໄຂຂໍ້ມູນເກົ່າ
            const { error } = await supabaseClient.from('Locations').update(payload).eq('ID', l_id);
            if (error) { console.error('Error:', error); Swal.fire('Error', error.message, 'error'); return; }
        } else {
            // ເພີ່ມຂໍ້ມູນໃໝ່ (ສ້າງ ID ອັດຕະໂນມັດ)
            payload.ID = 'LOC' + Date.now();
            const { error } = await supabaseClient.from('Locations').insert([payload]);
            if (error) { console.error('Error:', error); Swal.fire('Error', error.message, 'error'); return; }
        }

        $('#locationModal').modal('hide');
        window.loadLocationsMasterView();
        Swal.fire('ສຳເລັດ!', '', 'success');
    } catch (err) {
        console.error('Error:', err);
        Swal.fire('Error', err.message, 'error');
    }
};

window.delLocation = function (id) {
    Swal.fire({ title: 'ລຶບ?', icon: 'warning', showCancelButton: true, confirmButtonText: 'ລຶບ' }).then(async r => {
        if (r.isConfirmed) {
            try {
                const { error } = await supabaseClient.from('Locations').delete().eq('ID', id);
                if (error) { console.error('Error:', error); Swal.fire('Error', error.message, 'error'); return; }
                window.loadLocationsMasterView();
                Swal.fire('ລຶບແລ້ວ', '', 'success');
            } catch (err) {
                console.error('Error:', err);
                Swal.fire('Error', err.message, 'error');
            }
        }
    });
};

window.loadServicesMasterView = async function () {
    if ($.fn.DataTable.isDataTable('#serviceTable')) $('#serviceTable').DataTable().destroy();
    $('#serviceTable tbody').html('<tr><td colspan="5" class="text-center py-4"><div class="spinner-border text-info spinner-border-sm"></div> ກຳລັງໂຫຼດ...</td></tr>');

    try {
        const { data: r, error } = await supabaseClient.from('Service_Lists').select('*');
        if (error) { console.error('Error:', error); Swal.fire('Error', error.message, 'error'); return; }

        servicesDataStore = r || [];
        if ($.fn.DataTable.isDataTable('#serviceTable')) $('#serviceTable').DataTable().destroy();

        let o = '';
        if (r) r.forEach(s => { o += `<option value="${s.Services_List}">${s.Services_List}</option>`; });
        $('#emrService').empty().append(o).trigger('change');

        let h = '';
        if (r && r.length > 0) {
            r.forEach(s => {
                let safeService = (s.Services_List || "").replace(/'/g, "\\'");
                let safeSpec = (s.Mapped_Specialist || "").replace(/'/g, "\\'");
                let safeRev = (s.Revenue_Group || "").replace(/'/g, "\\'");

                h += `<tr>
                        <td class="text-center"><input type="checkbox" class="form-check-input bulk-check-services" value="${s.ID}"></td>
                        <td class="fw-bold text-primary">${s.Services_List}</td>
                        <td><span class="badge bg-info text-dark">${s.Mapped_Specialist || '-'}</span></td>
                        <td><span class="badge bg-success">${s.Revenue_Group || '-'}</span></td>
                        <td class="text-center">
                            <button class="btn btn-sm btn-primary shadow-sm me-1" onclick="window.editService('${s.ID}','${safeService}','${safeSpec}','${safeRev}')"><i class="fas fa-edit"></i></button>
                            <button class="btn btn-sm btn-danger shadow-sm" onclick="window.delService('${s.ID}')"><i class="fas fa-trash"></i></button>
                        </td>
                      </tr>`;
            });
        }
        $('#serviceTable tbody').html(h);
        $('#serviceTable').DataTable({ responsive: true, pageLength: 10, language: { search: "ຄົ້ນຫາ:", emptyTable: "ບໍ່ມີຂໍ້ມູນ" } });
    } catch (err) {
        console.error('System Error:', err);
        Swal.fire('Error', err.message, 'error');
    }
};

window.submitServiceForm = async function (e) {
    if (e) e.preventDefault();
    Swal.fire({ title: 'ກຳລັງບັນທຶກ...', didOpen: () => Swal.showLoading() });
    try {
        let payload = {
            Services_List: $('#s_serv').val(),
            Mapped_Specialist: $('#s_spec').val(),
            Revenue_Group: $('#s_rev').val()
        };
        let s_id = $('#s_id').val();

        if (s_id) {
            const { error } = await supabaseClient.from('Service_Lists').update(payload).eq('ID', s_id);
            if (error) { console.error('Error:', error); Swal.fire('Error', error.message, 'error'); return; }
        } else {
            payload.ID = 'SRV' + Date.now();
            const { error } = await supabaseClient.from('Service_Lists').insert([payload]);
            if (error) { console.error('Error:', error); Swal.fire('Error', error.message, 'error'); return; }
        }

        $('#serviceModal').modal('hide');
        window.loadServicesMasterView();
        Swal.fire('ສຳເລັດ!', '', 'success');
    } catch (err) {
        console.error('Error:', err);
        Swal.fire('Error', err.message, 'error');
    }
};

window.delService = function (id) {
    Swal.fire({ title: 'ລຶບ?', icon: 'warning', showCancelButton: true, confirmButtonText: 'ລຶບ' }).then(async r => {
        if (r.isConfirmed) {
            try {
                const { error } = await supabaseClient.from('Service_Lists').delete().eq('ID', id);
                if (error) { console.error('Error:', error); Swal.fire('Error', error.message, 'error'); return; }
                window.loadServicesMasterView();
                Swal.fire('ລຶບແລ້ວ', '', 'success');
            } catch (err) {
                console.error('Error:', err);
                Swal.fire('Error', err.message, 'error');
            }
        }
    });
};

window.openAddServiceModal = function () {
    $('#serviceForm')[0].reset();
    $('#s_id').val('');
    $('#serviceModal').modal('show');
};

window.editService = function (id, sv, sp, rv) {
    $('#s_id').val(id);
    $('#s_serv').val(sv);
    $('#s_spec').val(sp);
    $('#s_rev').val(rv);
    $('#serviceModal').modal('show');
};

window.setBrandName = function (name) {
    var el = document.getElementById('topnavBrandName');
    if (el && name) el.textContent = name;
};

window.handlePatientExcelUpload = function (e) {
    let file = e.target.files[0];
    if (!file) return;
    let reader = new FileReader();
    reader.onload = async function (evt) {
        Swal.fire({ title: 'ກຳລັງປະມວນຜົນ...', didOpen: () => Swal.showLoading() });
        let data = new Uint8Array(evt.target.result);
        let workbook = XLSX.read(data, { type: 'array', cellDates: true });
        let firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        let jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1, raw: true, defval: "" });

        for (let i = 1; i < jsonData.length; i++) {
            let row = jsonData[i];
            if (row[5] instanceof Date) {
                let d = row[5];
                row[5] = d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, '0') + "-" + String(d.getDate()).padStart(2, '0');
            }
            if (row[25] instanceof Date) {
                let d = row[25];
                row[25] = d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, '0') + "-" + String(d.getDate()).padStart(2, '0');
            }
            if (row[26] instanceof Date) {
                row[26] = String(row[26].getHours()).padStart(2, '0') + ":" + String(row[26].getMinutes()).padStart(2, '0');
            } else if (typeof row[26] === 'number') {
                let totalSeconds = Math.floor(row[26] * 86400);
                let hours = Math.floor(totalSeconds / 3600);
                let mins = Math.floor((totalSeconds % 3600) / 60);
                row[26] = String(hours).padStart(2, '0') + ":" + String(mins).padStart(2, '0');
            }
        }

        let insertData = [];
        for (let i = 1; i < jsonData.length; i++) {
            let row = jsonData[i];
            if (row.length < 1 || !row[0]) continue;
            // Map array columns to Supabase 'Patients' table columns
            // Assumes standard template structure - adjust map if template varies
            let parsedAge = parseInt(row[6]) || 0;
            let ageGroup = row[28] || '';
            if (parsedAge && !ageGroup) {
                ageGroup = parsedAge <= 15 ? '0-15' : (parsedAge <= 35 ? '16-35' : (parsedAge <= 55 ? '36-55' : '55+'));
            }

            // Map array columns to Supabase 'Patients' table columns accurately based on real Excel structure
            insertData.push({
                Patient_ID: row[0],
                Title: row[1] || '',
                First_Name: row[2] || '',
                Last_Name: row[3] || '',
                Gender: row[4] || '',
                Date_of_Birth: row[5] || null,
                Age: parsedAge,
                Nationality: row[7] || '',
                Occupation: row[8] || '',
                Blood_Type: row[9] || '',
                Phone_Number: row[10] || '',
                Email: row[11] || '',
                Address: row[12] || '',
                District: row[13] || '',
                Province: row[14] || '',
                Organization_ID: row[15] || '',
                Name_Org: row[16] || '',
                Insurance_Code: row[17] || '',
                Insured_Person_Name: row[18] || '',
                Drug_Allergy: row[19] || '',
                Underlying_Disease: row[20] || '',
                Emergency_Name: row[21] || '',
                Emergency_Contact: row[22] || '',
                Emergency_Relation: row[23] || '',
                Channel: row[24] || '',
                Registration_Date: row[25] || null,
                Time: row[26] || '',
                Shift: row[27] || '',
                Age_Group: ageGroup
            });
        }

        if (insertData.length > 0) {
            const { error } = await supabaseClient.from('Patients').insert(insertData);
            if (!error) {
                Swal.fire('ສຳເລັດ!', `ນຳເຂົ້າສຳເລັດ ${insertData.length} ລາຍການ`, 'success');
                window.initPatientTable();
                window.preloadDropdownData();
            } else {
                Swal.fire('ຜິດພາດ!', error.message, 'error');
            }
        } else {
            Swal.fire('ຜິດພາດ!', 'ບໍ່ພົບຂໍ້ມູນ', 'error');
        }
        $('#patientExcelInput').val("");
    };
    reader.readAsArrayBuffer(file);
};

window.handleLocationExcelUpload = function (e) {
    let file = e.target.files[0];
    if (!file) return;
    let reader = new FileReader();
    reader.onload = async function (evt) {
        Swal.fire({ title: 'ກຳລັງປະມວນຜົນ...', didOpen: () => Swal.showLoading() });
        let data = new Uint8Array(evt.target.result);
        let workbook = XLSX.read(data, { type: 'array' });
        let firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        let jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: "" });

        let insertData = [];
        for (let i = 1; i < jsonData.length; i++) {
            let row = jsonData[i];
            if (row.length < 1 || !row[0]) continue;
            insertData.push({ District: row[0], Province: row[1] || '' });
        }

        if (insertData.length > 0) {
            const { error } = await supabaseClient.from('Locations').insert(insertData);
            if (!error) {
                Swal.fire('ສຳເລັດ!', `ນຳເຂົ້າສຳເລັດ ${insertData.length} ລາຍການ`, 'success');
                window.loadLocationsMasterView();
            } else {
                Swal.fire('ຜິດພາດ!', error.message, 'error');
            }
        } else {
            Swal.fire('ຜິດພາດ!', 'ບໍ່ພົບຂໍ້ມູນ', 'error');
        }
        $('#locExcelInput').val("");
    };
    reader.readAsArrayBuffer(file);
};

window.handleExcelUpload = function (e) {
    let file = e.target.files[0];
    if (!file) return;
    let reader = new FileReader();
    reader.onload = async function (evt) {
        Swal.fire({ title: 'ກຳລັງປະມວນຜົນ...', didOpen: () => Swal.showLoading() });
        let data = new Uint8Array(evt.target.result);
        let workbook = XLSX.read(data, { type: 'array' });
        let firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        let jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: "" });

        let insertData = [];
        for (let i = 1; i < jsonData.length; i++) {
            let row = jsonData[i];
            if (row.length < 1 || !row[0]) continue;
            insertData.push({ Services_List: row[0], Mapped_Specialist: row[1] || '', Revenue_Group: row[2] || '' });
        }

        if (insertData.length > 0) {
            const { error } = await supabaseClient.from('Service_Lists').insert(insertData);
            if (!error) {
                Swal.fire('ສຳເລັດ!', `ນຳເຂົ້າສຳເລັດ ${insertData.length} ລາຍການ`, 'success');
                window.loadServicesMasterView();
            } else {
                Swal.fire('ຜິດພາດ!', error.message, 'error');
            }
        } else {
            Swal.fire('ຜິດພາດ!', 'ບໍ່ພົບຂໍ້ມູນ', 'error');
        }
        $('#excelFileInput').val("");
    };
    reader.readAsArrayBuffer(file);
};

window.handleOrgExcelUpload = function(e) {
    let file = e.target.files[0];
    if (!file) return;

    let reader = new FileReader();
    reader.onload = async function(ev) {
        try {
            let data = new Uint8Array(ev.target.result);
            let workbook = XLSX.read(data, {type: 'array'});
            let firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            let excelRows = XLSX.utils.sheet_to_json(firstSheet, {defval: ""});

            if (!excelRows || excelRows.length === 0) {
                Swal.fire('Error', 'ບໍ່ພົບຂໍ້ມູນໃນໄຟລ໌', 'error');
                return;
            }

            Swal.fire({ title: 'ກຳລັງອັບໂຫຼດ...', didOpen: () => { Swal.showLoading() } });

            let insertData = [];
            for (let row of excelRows) {
                let obj = {
                    Cus_ID_Ex: row['Cus_ID_Ex'] || '',
                    Name: row['Name'] || '',
                    Org_Name: row['Org_Name'] || '',
                    Org_Code: row['Org_Code'] || row['Org_ID'] || '',
                    Discount: row['Discount'] ? String(row['Discount']) : '',
                    Status: 'Active'
                };
                if (row['Org_ID']) obj.Org_ID = String(row['Org_ID']);
                insertData.push(obj);
            }

            const { error } = await supabaseClient.from('Organizations').insert(insertData);
            if (error) {
                Swal.fire('Error', error.message, 'error');
            } else {
                Swal.fire('ສຳເລັດ', 'ນຳເຂົ້າຂໍ້ມູນອົງກອນສຳເລັດແລ້ວ', 'success');
                window.loadOrgs();
                window.preloadDropdownData();
            }
        } catch (err) {
            console.error(err);
            Swal.fire('Error', 'ເກີດຂໍ້ຜິດພາດໃນການອ່ານໄຟລ໌ Excel', 'error');
        }
        $('#orgExcelInput').val(''); 
    };
    reader.readAsArrayBuffer(file);
};

window.handleDrugExcelUpload = function (e) {
    let f = e.target.files[0];
    if (!f) return;
    let r = new FileReader();
    r.onload = async function (ev) {
        Swal.fire({ title: 'ກຳລັງປະມວນຜົນ...', didOpen: () => Swal.showLoading() });
        let workbook = XLSX.read(new Uint8Array(ev.target.result), { type: 'array' });
        let firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        let jd = XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: "" });

        let insertData = [];
        for (let i = 1; i < jd.length; i++) {
            let row = jd[i];
            if (row.length < 1 || !row[0]) continue;
            insertData.push({ Drug_Name: row[0], Description: row[1] || '' });
        }

        if (insertData.length > 0) {
            const { error } = await supabaseClient.from('Drugs_Master').insert(insertData);
            if (!error) {
                Swal.fire('ສຳເລັດ!', `ນຳເຂົ້າ ${insertData.length} ລາຍການ`, 'success');
                window.loadDrugsMaster();
                window.preloadDropdownData();
            } else {
                Swal.fire('ຜິດພາດ!', error.message, 'error');
            }
        } else {
            Swal.fire('ແຈ້ງເຕືອນ!', 'ບໍ່ພົບຂໍ້ມູນທີ່ຈະນຳເຂົ້າ — ກວດສອບ format ຂອງ Excel', 'warning');
        }
        $('#drugExcelInput').val('');
    };
    r.readAsArrayBuffer(f);
};

window.handleLabExcelUpload = function (e) {
    let f = e.target.files[0];
    if (!f) return;
    let r = new FileReader();
    r.onload = async function (ev) {
        Swal.fire({ title: 'ກຳລັງປະມວນຜົນ...', didOpen: () => Swal.showLoading() });
        let workbook = XLSX.read(new Uint8Array(ev.target.result), { type: 'array' });
        let firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        let jd = XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: "" });

        let insertData = [];
        for (let i = 1; i < jd.length; i++) {
            let row = jd[i];
            if (row.length < 1 || !row[0]) continue;
            insertData.push({ Lab_Name: row[0], Description: row[1] || '' });
        }

        if (insertData.length > 0) {
            const { error } = await supabaseClient.from('Labs_Master').insert(insertData);
            if (!error) {
                Swal.fire('ສຳເລັດ!', `ນຳເຂົ້າ ${insertData.length} ລາຍການ`, 'success');
                window.loadLabsMaster();
                window.preloadDropdownData();
            } else {
                Swal.fire('ຜິດພາດ!', error.message, 'error');
            }
        } else {
            Swal.fire('ແຈ້ງເຕືອນ!', 'ບໍ່ພົບຂໍ້ມູນທີ່ຈະນຳເຂົ້າ — ກວດສອບ format ຂອງ Excel', 'warning');
        }
        $('#labExcelInput').val('');
    };
    r.readAsArrayBuffer(f);
};

// Bulk Delete Logic
window.toggleAllCheckboxes = function (masterCheckbox, type) {
    const isChecked = $(masterCheckbox).is(':checked');
    $(`.bulk-check-${type}`).prop('checked', !!isChecked);
};

window.bulkDelete = async function (type) {
    const checked = $(`.bulk-check-${type}:checked`);
    if (checked.length === 0) {
        Swal.fire('ແຈ້ງເຕືອນ', 'ກະລຸນາເລືອກລາຍການທີ່ຕ້ອງການລຶບກ່ອນ', 'warning');
        return;
    }

    const ids = [];
    checked.each(function () {
        ids.push($(this).val());
    });

    const config = {
        'vacMaster': { table: 'Vaccines_Master', col: 'Vac_ID', reload: window.loadVaccineMaster },
        'drugs': { table: 'Drugs_Master', col: 'Drug_ID', reload: window.loadDrugsMaster },
        'labs': { table: 'Labs_Master', col: 'Lab_ID', reload: window.loadLabsMaster },
        'users': { table: 'Users', col: 'ID', reload: window.loadUsers },
        'orgs': { table: 'Organizations', col: 'Org_ID', reload: window.loadOrgs },
        'locations': { table: 'Locations', col: 'ID', reload: window.loadLocationsMasterView },
        'services': { table: 'Service_Lists', col: 'ID', reload: window.loadServicesMasterView }
    };

    const cfg = config[type];
    if (!cfg) return;

    Swal.fire({
        title: 'ຍືນຍັນການລຶບ?',
        text: `ທ່ານແນ່ໃຈບໍ່ວ່າຕ້ອງການລຶບ ${ids.length} ລາຍການທີ່ເລືອກ?`,
        icon: 'warning',
        showCancelButton: true,
        confirmButtonColor: '#ef4444',
        confirmButtonText: 'ລຶບເລີຍ',
        cancelButtonText: 'ຍົກເລີກ'
    }).then(async (result) => {
        if (result.isConfirmed) {
            Swal.fire({ title: 'ກຳລັງລຶບ...', didOpen: () => Swal.showLoading() });
            try {
                const { error } = await supabaseClient.from(cfg.table).delete().in(cfg.col, ids);
                if (error) throw error;
                
                cfg.reload();
                Swal.fire('ສຳເລັດ', 'ລຶບລາຍການທີ່ເລືອກສຳເລັດແລ້ວ', 'success');
            } catch (err) {
                console.error(err);
                Swal.fire('ຂໍ້ຜິດພາດ', err.message, 'error');
            }
        }
    });
};

// ==========================================
// ACTIVITY LOG SYSTEM
// ==========================================

/**
 * logAction — fire-and-forget logger
 * @param {string} action  e.g. 'Login', 'Add', 'Edit', 'Delete', 'Save'
 * @param {string} details e.g. 'Patient P25-0001 (John Doe)'
 * @param {string} module  e.g. 'Patients', 'Triage', 'OPD'
 */
window.logAction = function (action, details, module) {
    try {
        const userId   = currentUser ? currentUser.id   : '-';
        const userName = currentUser ? currentUser.name : 'System';
        supabaseClient.from('activity_logs').insert({
            timestamp : new Date().toISOString(),
            user_id   : userId,
            user_name : userName,
            action    : action,
            details   : details || '',
            module    : module  || ''
        }).then(({ error }) => {
            if (error) console.warn('logAction error:', error.message);
        });
    } catch (e) {
        console.warn('logAction exception:', e);
    }
};

window.loadActivityLog = async function () {
    const today = window.getLocalStr(new Date());
    if (!$('#logStartDate').val()) $('#logStartDate').val(today);
    if (!$('#logEndDate').val())   $('#logEndDate').val(today);

    let sDate  = $('#logStartDate').val();
    let eDate  = $('#logEndDate').val();
    let mod    = $('#logModuleFilter').val();
    let user   = $('#logUserFilter').val().toLowerCase();
    let act    = $('#logActionFilter').val();

    if ($.fn.DataTable.isDataTable('#activityLogTable')) $('#activityLogTable').DataTable().destroy();
    $('#activityLogTableBody').html('<tr><td colspan="5" class="text-center py-4"><div class="spinner-border text-info spinner-border-sm"></div> ກຳລັງໂຫຼດ...</td></tr>');

    try {
        let query = supabaseClient
            .from('activity_logs')
            .select('*')
            .gte('timestamp', sDate + 'T00:00:00Z')
            .lte('timestamp', eDate + 'T23:59:59Z')
            .order('timestamp', { ascending: false })
            .limit(500);

        if (mod)  query = query.eq('module', mod);
        if (act)  query = query.ilike('action', '%' + act + '%');

        const { data: logs, error } = await query;
        if (error) throw error;

        let rows = (logs || []).filter(l => !user || (l.user_name || '').toLowerCase().includes(user));

        // Summary counts
        let cAdd = 0, cEdit = 0, cDel = 0;
        rows.forEach(l => {
            let a = (l.action || '').toLowerCase();
            if (a.includes('add') || a.includes('save') || a.includes('login')) cAdd++;
            else if (a.includes('edit')) cEdit++;
            else if (a.includes('delete')) cDel++;
        });
        $('#logCountTotal').text(rows.length);
        $('#logCountAdd').text(cAdd);
        $('#logCountEdit').text(cEdit);
        $('#logCountDelete').text(cDel);

        // Render rows
        const actionBadge = (a) => {
            let al = (a || '').toLowerCase();
            if (al === 'login')  return `<span class="badge bg-primary">${a}</span>`;
            if (al === 'logout') return `<span class="badge bg-secondary">${a}</span>`;
            if (al === 'add' || al === 'save') return `<span class="badge bg-success">${a}</span>`;
            if (al === 'edit')   return `<span class="badge bg-warning text-dark">${a}</span>`;
            if (al === 'delete') return `<span class="badge bg-danger">${a}</span>`;
            return `<span class="badge bg-info text-dark">${a}</span>`;
        };

        let h = '';
        rows.forEach(l => {
            let d = new Date(l.timestamp);
            let dateStr = d.toLocaleDateString('en-GB') + ' ' + d.toLocaleTimeString('en-GB', { hour:'2-digit', minute:'2-digit', second:'2-digit' });
            h += `<tr>
                <td class="text-muted small">${dateStr}</td>
                <td><span class="fw-bold">${l.user_name || '-'}</span><br><small class="text-muted">${l.user_id || ''}</small></td>
                <td>${actionBadge(l.action)}</td>
                <td class="small">${l.details || '-'}</td>
                <td><span class="badge bg-light text-dark border">${l.module || '-'}</span></td>
              </tr>`;
        });

        if (rows.length === 0) h = '<tr><td colspan="5" class="text-center py-4 text-muted"><i class="fas fa-inbox me-2"></i>ບໍ່ມີ Log ໃນຊ່ວງວັນທີນີ້</td></tr>';
        $('#activityLogTableBody').html(h);
        $('#activityLogTable').DataTable({
            responsive: true, pageLength: 25,
            language: { search: 'ຄົ້ນຫາ:', emptyTable: 'ບໍ່ມີຂໍ້ມູນ', lengthMenu: 'ສະແດງ _MENU_' }
        });
    } catch (err) {
        console.error('loadActivityLog error:', err);
        $('#activityLogTableBody').html(`<tr><td colspan="5" class="text-center text-danger py-4"><i class="fas fa-exclamation-triangle me-2"></i>${err.message}</td></tr>`);
    }
};

window.exportActivityLogCSV = function () {
    if (!$.fn.DataTable.isDataTable('#activityLogTable')) {
        Swal.fire('ແຈ້ງເຕືອນ', 'ກະລຸນາໂຫຼດ Log ກ່ອນ Export', 'warning');
        return;
    }
    let dt = $('#activityLogTable').DataTable();
    let rows = [['ວັນທີ / ເວລາ', 'ຜູ້ໃຊ້', 'User ID', 'ການກະທຳ', 'ລາຍລະອຽດ', 'Module']];
    dt.rows().data().each(function (r) {
        let cells = [];
        for (let i = 0; i < 5; i++) {
            let div = document.createElement('div');
            div.innerHTML = r[i] || '';
            cells.push(div.innerText.replace(/\n/g, ' ').trim());
        }
        rows.push(cells);
    });
    let csv = rows.map(r => r.map(c => '"' + (c + '').replace(/"/g, '""') + '"').join(',')).join('\n');
    let blob = new Blob(['\ufeff' + csv], { type: 'text/csv;charset=utf-8;' });
    let url = URL.createObjectURL(blob);
    let a = document.createElement('a');
    a.href = url; a.download = 'Activity_Log_' + window.getLocalStr(new Date()) + '.csv'; a.click();
    URL.revokeObjectURL(url);
};

