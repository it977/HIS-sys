// ============================================================
//  Code.gs — Supabase Backend (replaces Google Sheets)
//  ทุก function ยังคง signature เดิม เพื่อให้ main.js ที่ยัง
//  ใช้ google.script.run เรียกได้โดยไม่ต้องแก้
// ============================================================

var SUPABASE_URL = "https://erueurkqzmtdefszqons.supabase.co";
var SUPABASE_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImVydWV1cmtxem10ZGVmc3pxb25zIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzMxOTA2OTksImV4cCI6MjA4ODc2NjY5OX0.uShip2ECxvFPfmDLx9-adHGXXTc3cazVdZpSF2tCFUw";

// ─── Google Apps Script Web App Entry Points ─────────────────
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Hospital Information System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ─── Supabase REST API Helper ─────────────────────────────────
function sbFetch(method, table, params, body) {
  var url = SUPABASE_URL + "/rest/v1/" + table;
  if (params) url += "?" + params;
  var options = {
    method: method,
    headers: {
      "apikey": SUPABASE_ANON_KEY,
      "Authorization": "Bearer " + SUPABASE_ANON_KEY,
      "Content-Type": "application/json",
      "Prefer": method === "POST" ? "return=representation" : ""
    },
    muteHttpExceptions: true
  };
  if (body) options.payload = JSON.stringify(body);
  var res = UrlFetchApp.fetch(url, options);
  var code = res.getResponseCode();
  var text = res.getContentText();
  if (code >= 200 && code < 300) {
    return text ? JSON.parse(text) : [];
  }
  throw new Error("Supabase error " + code + ": " + text);
}

// ─── USERS ────────────────────────────────────────────────────
function loginUser(email, pass) {
  try {
    var rows = sbFetch("GET", "Users", "Email=eq." + encodeURIComponent(email) + "&select=*");
    if (!rows || rows.length === 0) return { success: false, message: "ອີເມວ ຫຼື ລະຫັດຜ່ານບໍ່ຖືກຕ້ອງ" };
    var u = rows[0];
    if (u.Status === "inactive") return { success: false, message: "ບັນຊີຖືກປິດໃຊ້ງານ" };
    if (u.Password !== pass) return { success: false, message: "ອີເມວ ຫຼື ລະຫັດຜ່ານບໍ່ຖືກຕ້ອງ" };
    // Update LastLogin
    sbFetch("PATCH", "Users", "ID=eq." + u.ID, { LastLogin: new Date().toISOString() });
    return { success: true, user: { id: u.ID, name: u.Name, email: u.Email, role: (u.Role || "").toLowerCase(), permissions: u.Permissions || "all" } };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function getAllUsers() {
  try {
    var rows = sbFetch("GET", "Users", "select=ID,Name,Email,Role,Permissions,Status&order=ID.asc");
    return rows.map(function(u) { return { id: u.ID, name: u.Name, email: u.Email, role: u.Role, permissions: u.Permissions, status: u.Status }; });
  } catch(e) { return []; }
}

function saveUser(f) {
  try {
    if (f.u_id) {
      var upd = { Name: f.u_name, Role: f.u_role, Permissions: f.u_permissions };
      if (f.u_pass) upd.Password = f.u_pass;
      upd.UpdatedAt = new Date().toISOString();
      sbFetch("PATCH", "Users", "ID=eq." + f.u_id, upd);
      return { success: true, message: "ແກ້ໄຂສຳເລັດ" };
    } else {
      // Check duplicate email
      var ex = sbFetch("GET", "Users", "Email=eq." + encodeURIComponent(f.u_email));
      if (ex && ex.length > 0) return { success: false, message: "ອີເມວນີ້ມີໃນລະບົບແລ້ວ" };
      sbFetch("POST", "Users", null, { ID: new Date().getTime().toString(), Name: f.u_name, Email: f.u_email, Password: f.u_pass, Role: (f.u_role || "").toLowerCase(), Permissions: f.u_permissions, Status: "active", UpdatedAt: new Date().toISOString() });
      return { success: true, message: "ເພີ່ມສຳເລັດ" };
    }
  } catch(e) { return { success: false, message: e.toString() }; }
}

function deleteUser(id) {
  try { sbFetch("DELETE", "Users", "ID=eq." + id); return { success: true }; }
  catch(e) { return { success: false }; }
}

// ─── SETTINGS ─────────────────────────────────────────────────
function getSettings() {
  try {
    var rows = sbFetch("GET", "Settings", "select=Key,Value");
    var set = { hospitalName: "HIS HOSPITAL", logoUrl: "", opdHeaderUrl: "", opdFooterUrl: "" };
    (rows || []).forEach(function(r) {
      if (r.Key === "HospitalName")  set.hospitalName  = r.Value || set.hospitalName;
      if (r.Key === "LogoUrl")       set.logoUrl       = r.Value || "";
      if (r.Key === "OpdHeaderUrl")  set.opdHeaderUrl  = r.Value || "";
      if (r.Key === "OpdFooterUrl")  set.opdFooterUrl  = r.Value || "";
    });
    return set;
  } catch(e) { return { hospitalName: "HIS HOSPITAL", logoUrl: "", opdHeaderUrl: "", opdFooterUrl: "" }; }
}

function saveSettings(f) {
  try {
    var pairs = [
      { Key: "HospitalName",  Value: f.hospitalName  || "" },
      { Key: "LogoUrl",       Value: f.logoUrl       || "" },
      { Key: "OpdHeaderUrl",  Value: f.opdHeaderUrl  || "" },
      { Key: "OpdFooterUrl",  Value: f.opdFooterUrl  || "" }
    ];
    pairs.forEach(function(p) {
      var ex = sbFetch("GET", "Settings", "Key=eq." + p.Key);
      if (ex && ex.length > 0) sbFetch("PATCH", "Settings", "Key=eq." + p.Key, { Value: p.Value });
      else sbFetch("POST", "Settings", null, p);
    });
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// ─── MASTER DATA ───────────────────────────────────────────────
function getMasterData() {
  try {
    var rows = sbFetch("GET", "MasterData", "select=ID,Category,Value&order=Category.asc");
    var r = {};
    (rows || []).forEach(function(row) {
      if (!r[row.Category]) r[row.Category] = [];
      r[row.Category].push({ id: row.ID, value: row.Value });
    });
    return r;
  } catch(e) { return {}; }
}

function addMasterData(category, value) {
  try {
    sbFetch("POST", "MasterData", null, { ID: new Date().getTime().toString(), Category: category, Value: value });
    return getMasterData();
  } catch(e) { return {}; }
}

function deleteMasterData(id) {
  try { sbFetch("DELETE", "MasterData", "ID=eq." + id); return getMasterData(); }
  catch(e) { return {}; }
}

// ─── PATIENTS ──────────────────────────────────────────────────
function getNextPatientID() {
  try {
    var rows = sbFetch("GET", "Patients", "select=Patient_ID&order=Patient_ID.desc&limit=1");
    if (!rows || rows.length === 0) return "CN0000001";
    var last = rows[0].Patient_ID || "CN0000000";
    var num = parseInt(last.replace("CN", ""), 10);
    if (isNaN(num)) return "CN0000001";
    return "CN" + ("0000000" + (num + 1)).slice(-7);
  } catch(e) { return "CN0000001"; }
}

function getAllPatients() {
  try {
    var rows = sbFetch("GET", "Patients", "select=*&order=Patient_ID.desc");
    return (rows || []).map(function(r) {
      var fullName = ((r.Title || "") + " " + (r.First_Name || "") + " " + (r.Last_Name || "")).trim();
      var addr = [r.District, r.Province].filter(Boolean).join(", ") || "-";
      return {
        id: r.Patient_ID, fullname: fullName, gender: r.Gender || "-",
        age: r.Age || "-", phone: r.Phone_Number || "-", address: addr,
        allergy: r.Drug_Allergy || "ບໍ່ມີ",
        regDate: (r.Registration_Date || "-").toString().split("T")[0],
        regTime: r.Time || "-"
      };
    });
  } catch(e) { return []; }
}

function savePatient(f) {
  try {
    var isEdit = f.p_action === "edit" && f.p_id;
    var pId = isEdit ? f.p_id : getNextPatientID();
    var age = parseInt(f.p_age) || 0;
    var ageGroup = age <= 15 ? "0-15" : (age <= 35 ? "16-35" : (age <= 55 ? "36-55" : "55+"));
    var row = {
      Patient_ID: pId, Title: f.p_title, First_Name: f.p_firstname, Last_Name: f.p_lastname,
      Gender: f.p_gender, Date_of_Birth: f.p_dob || null, Age: age,
      Nationality: f.p_nation, Occupation: f.p_job, Blood_Type: f.p_blood,
      Phone_Number: f.p_phone, Email: f.p_email, Address: f.p_address,
      District: f.p_district, Province: f.p_province,
      Organization_ID: f.p_org_id, Name_Org: f.p_org_name,
      Insurance_Company: f.p_ins_company, Insurance_Code: f.p_ins_code, Insured_Person_Name: f.p_ins_name,
      Drug_Allergy: f.p_allergy, Underlying_Disease: f.p_disease,
      Emergency_Name: f.p_emer_name, Emergency_Contact: f.p_emer_contact, Emergency_Relation: f.p_emer_relation,
      Channel: f.p_channel, Registration_Date: f.p_date || null, Time: f.p_time,
      Shift: f.p_shift, Age_Group: ageGroup
    };
    if (isEdit) sbFetch("PATCH", "Patients", "Patient_ID=eq." + pId, row);
    else sbFetch("POST", "Patients", null, row);
    return { success: true, id: pId };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function getPatientDetails(id) {
  try {
    var rows = sbFetch("GET", "Patients", "Patient_ID=eq." + encodeURIComponent(id) + "&select=*");
    if (!rows || rows.length === 0) return null;
    var r = rows[0];
    return {
      id: r.Patient_ID, title: r.Title || "", firstname: r.First_Name || "", lastname: r.Last_Name || "",
      gender: r.Gender || "", dob: r.Date_of_Birth || "", age: r.Age || "",
      nation: r.Nationality || "", job: r.Occupation || "", blood: r.Blood_Type || "",
      phone: r.Phone_Number || "", email: r.Email || "", address: r.Address || "",
      district: r.District || "", province: r.Province || "",
      org_id: r.Organization_ID || "", org_name: r.Name_Org || "",
      ins_company: r.Insurance_Company || "", ins_code: r.Insurance_Code || "", ins_name: r.Insured_Person_Name || "",
      allergy: r.Drug_Allergy || "ບໍ່ມີ", disease: r.Underlying_Disease || "ບໍ່ມີ",
      emer_name: r.Emergency_Name || "", emer_contact: r.Emergency_Contact || "", emer_relation: r.Emergency_Relation || "",
      channel: r.Channel || "", date: r.Registration_Date || "", time: r.Time || "", shift: r.Shift || ""
    };
  } catch(e) { return null; }
}

function deletePatient(id) {
  try { sbFetch("DELETE", "Patients", "Patient_ID=eq." + id); return { success: true }; }
  catch(e) { return { success: false }; }
}

function importPatientsFromExcel(dataArray) {
  try {
    var rows = [];
    for (var i = 0; i < dataArray.length; i++) {
      var r = dataArray[i];
      if (!r || r.length < 5) continue;
      var pId = r[0] ? r[0].toString().trim() : "";
      if (!pId || pId.toLowerCase() === "patient_id") continue;
      rows.push({
        Patient_ID: pId, Title: r[1] || "", First_Name: r[2] || "", Last_Name: r[3] || "",
        Gender: r[4] || "", Date_of_Birth: r[5] || null, Age: parseInt(r[6]) || 0,
        Nationality: r[7] || "", Occupation: r[8] || "", Blood_Type: r[9] || "",
        Phone_Number: r[10] || "", Email: r[11] || "", Address: r[12] || "",
        District: r[13] || "", Province: r[14] || "",
        Organization_ID: r[15] || "", Name_Org: r[16] || "",
        Insurance_Company: r[17] || "", Insurance_Code: r[18] || "", Insured_Person_Name: r[19] || "",
        Drug_Allergy: r[20] || "ບໍ່ມີ", Underlying_Disease: r[21] || "ບໍ່ມີ",
        Emergency_Name: r[22] || "", Emergency_Contact: r[23] || "", Emergency_Relation: r[24] || "",
        Channel: r[25] || "", Registration_Date: r[26] || null, Time: r[27] || "",
        Shift: r[28] || "", Age_Group: r[29] || ""
      });
    }
    if (rows.length > 0) sbFetch("POST", "Patients", null, rows);
    return { success: true, count: rows.length };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// ─── VISITS / TRIAGE ──────────────────────────────────────────
function sendToTriage(pId, pName) {
  try {
    var rows = sbFetch("GET", "Visits", "select=Visit_ID&order=Visit_ID.desc&limit=1");
    var lastNum = 0;
    if (rows && rows.length > 0) {
      var m = (rows[0].Visit_ID || "").match(/\d+$/);
      if (m) lastNum = parseInt(m[0]);
    }
    var vId = "V" + new Date().getFullYear().toString().slice(-2) + "-" + ("0000" + (lastNum + 1)).slice(-4);
    sbFetch("POST", "Visits", null, {
      Visit_ID: vId, Date: new Date().toISOString(),
      Patient_ID: pId, Patient_Name: pName, Status: "Triage"
    });
    return { success: true, message: "ສົ່ງສຳເລັດ" };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function getTriageQueue() {
  try {
    var visits = sbFetch("GET", "Visits", "Status=in.(Triage,Waiting OPD,Examining,Pharmacy,Completed,Admit,Transfer,Waiting Lab)&select=*&order=Date.desc");
    var patients = sbFetch("GET", "Patients", "select=Patient_ID,Age");
    var pAgeMap = {};
    (patients || []).forEach(function(p) { pAgeMap[p.Patient_ID] = p.Age; });
    var today = new Date().toISOString().split("T")[0];
    var firstVisitMap = {};
    (visits || []).forEach(function(v) {
      if (!firstVisitMap[v.Patient_ID]) firstVisitMap[v.Patient_ID] = v.Visit_ID;
    });
    return (visits || []).filter(function(v) {
      var vDate = (v.Date || "").split("T")[0];
      return v.Status === "Triage" || vDate === today;
    }).map(function(v) {
      var dt = v.Date || "";
      return {
        rowIdx: v.id || v.Visit_ID,
        visitId: v.Visit_ID, date: dt.split("T")[0],
        time: dt.split("T")[1] ? dt.split("T")[1].substring(0,5) : "-",
        patientId: v.Patient_ID, patientName: v.Patient_Name,
        age: pAgeMap[v.Patient_ID] || "-", status: v.Status,
        department: v.Department || "-",
        isNew: (firstVisitMap[v.Patient_ID] === v.Visit_ID),
        bp: v.BP, temp: v.Temp, weight: v.Weight, height: v.Height,
        bmi: v.BMI, pulse: v.Pulse, spo2: v.SpO2, symptoms: v.Symptoms
      };
    });
  } catch(e) { return []; }
}

function getTodayQueue() {
  try {
    var today = new Date().toISOString().split("T")[0];
    var visits = sbFetch("GET", "Visits", "Date=gte." + today + "T00:00:00&Date=lte." + today + "T23:59:59&Status=in.(Waiting OPD,Examining,Waiting Lab,Pharmacy,Completed,Admit,Transfer)&select=*&order=Date.asc");
    var patients = sbFetch("GET", "Patients", "select=Patient_ID,Drug_Allergy");
    var pAllergyMap = {};
    (patients || []).forEach(function(p) { pAllergyMap[p.Patient_ID] = p.Drug_Allergy || "ບໍ່ມີ"; });
    var allVisits = sbFetch("GET", "Visits", "select=Visit_ID,Patient_ID&order=Date.asc");
    var firstVisitMap = {};
    (allVisits || []).forEach(function(v) {
      if (!firstVisitMap[v.Patient_ID]) firstVisitMap[v.Patient_ID] = v.Visit_ID;
    });
    return (visits || []).map(function(v) {
      var dt = v.Date || "";
      return {
        rowIdx: v.id || v.Visit_ID,
        visitId: v.Visit_ID,
        date: dt.split("T")[0], time: dt.split("T")[1] ? dt.split("T")[1].substring(0,5) : "-",
        patientId: v.Patient_ID, patientName: v.Patient_Name, status: v.Status,
        bp: v.BP, temp: v.Temp, weight: v.Weight, height: v.Height,
        bmi: v.BMI, pulse: v.Pulse, spo2: v.SpO2,
        department: v.Department, symptoms: v.Symptoms, diagnosis: v.Diagnosis,
        prescriptionStr: v.Prescription_JSON || "[]",
        doctor: v.Doctor_Name, type: v.Visit_Type, site: v.Site,
        pe: v.Physical_Exam, advice: v.Advice, followup: v.Follow_Up,
        services: v.Services_List, specialist: v.Mapped_Specialist, revenue: v.Revenue_Group,
        labOrdersStr: v.Lab_Orders_JSON || "[]",
        dischargeStatus: v.Discharge_Status || "",
        allergy: pAllergyMap[v.Patient_ID] || "ບໍ່ມີ",
        isNew: (firstVisitMap[v.Patient_ID] === v.Visit_ID)
      };
    });
  } catch(e) { return []; }
}

function saveVitalSigns(f) {
  try {
    sbFetch("PATCH", "Visits", "Visit_ID=eq." + f.visitId, {
      Status: "Waiting OPD", BP: f.v_bp, Temp: f.v_temp,
      Weight: f.v_weight, Height: f.v_height, BMI: f.v_bmi,
      Pulse: f.v_pulse, SpO2: f.v_spo2,
      Department: f.v_department, Symptoms: f.v_symptoms
    });
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function saveMedicalRecord(visitId, symptoms, diagnosis, presJson, labJson, docName, vType, vSite, pe, advice, followup, services, specialist, revenue, ds) {
  try {
    var statusMap = {
      "ລໍຖ້າຜົນແລັບ (Waiting Lab)": "Waiting Lab",
      "ນອນຕິດຕາມ (Admit / IPD)": "Admit",
      "ສົ່ງຕໍ່ (Transfer)": "Transfer",
      "ກວດສຳເລັດ / ກັບບ້ານ": "Completed"
    };
    var mainStatus = statusMap[ds] || "Pharmacy";
    sbFetch("PATCH", "Visits", "Visit_ID=eq." + visitId, {
      Status: mainStatus, Symptoms: symptoms, Diagnosis: diagnosis,
      Prescription_JSON: presJson, Doctor_Name: docName,
      Visit_Type: vType || "OPD", Site: vSite || "In-site",
      Physical_Exam: pe || "", Advice: advice || "", Follow_Up: followup || "",
      Services_List: services || "", Mapped_Specialist: specialist || "",
      Revenue_Group: revenue || "", Lab_Orders_JSON: labJson || "[]",
      Discharge_Status: ds || ""
    });
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function deleteVisit(visitId) {
  try { sbFetch("DELETE", "Visits", "Visit_ID=eq." + visitId); return { success: true }; }
  catch(e) { return { success: false, message: e.toString() }; }
}

// ─── DASHBOARD & REPORT ────────────────────────────────────────
function getDashboardStats(startDateStr, endDateStr) {
  try {
    var visits = sbFetch("GET", "Visits", "Date=gte." + startDateStr + "T00:00:00&Date=lte." + endDateStr + "T23:59:59&select=*");
    var patients = sbFetch("GET", "Patients", "select=Patient_ID,Gender,Age,Organization_ID,Insurance_Code,Channel,Province");
    var pDict = {};
    (patients || []).forEach(function(p) { pDict[p.Patient_ID] = p; });
    var allVisits = sbFetch("GET", "Visits", "select=Visit_ID,Patient_ID&order=Date.asc");
    var firstVisitMap = {};
    (allVisits || []).forEach(function(v) {
      if (!firstVisitMap[v.Patient_ID]) firstVisitMap[v.Patient_ID] = v.Visit_ID;
    });
    var stats = {
      visitCount: 0, newPatients: 0, oldPatients: 0, gn: 0, insCorp: 0,
      timeSlot: { "08:00-16:00": 0, "16:00-21:00": 0, "21:00-08:00": 0 },
      gender: { "ຊາຍ": 0, "ຍິງ": 0 },
      site: { "In-site": 0, "Onsite": 0 },
      deptType: { "OPD": 0, "IPD": 0 },
      customerType: { "GN (Walk-in)": 0, "INS (Insurance)": 0, "Corporate": 0 },
      channel: {}, ageGroup: { "0-15": 0, "16-35": 0, "36-55": 0, "55+": 0 },
      opdGender: { "ຊາຍ": 0, "ຍິງ": 0 }, ipdGender: { "ຊາຍ": 0, "ຍິງ": 0 },
      provinces: {}, services: {}, specialist: {}, revenue: {}
    };
    (visits || []).forEach(function(v) {
      stats.visitCount++;
      var p = pDict[v.Patient_ID] || {};
      var gender = p.Gender || "";
      var age = parseInt(p.Age) || 0;
      var vType = v.Visit_Type || "OPD";
      var vSite = v.Site || "In-site";
      var isMale = (gender === "ຊາຍ" || gender === "Male");
      if (isMale) stats.gender["ຊາຍ"]++; else stats.gender["ຍິງ"]++;
      if (age <= 15) stats.ageGroup["0-15"]++;
      else if (age <= 35) stats.ageGroup["16-35"]++;
      else if (age <= 55) stats.ageGroup["36-55"]++;
      else stats.ageGroup["55+"]++;
      if (p.Organization_ID || p.Insurance_Code) {
        stats.insCorp++;
        if (p.Organization_ID) stats.customerType["Corporate"]++;
        else stats.customerType["INS (Insurance)"]++;
      } else { stats.gn++; stats.customerType["GN (Walk-in)"]++; }
      var ch = p.Channel || "ບໍ່ລະບຸ";
      stats.channel[ch] = (stats.channel[ch] || 0) + 1;
      var pv = p.Province || "ບໍ່ລະບຸ";
      stats.provinces[pv] = (stats.provinces[pv] || 0) + 1;
      if (firstVisitMap[v.Patient_ID] === v.Visit_ID) stats.newPatients++; else stats.oldPatients++;
      if (vType === "IPD") { stats.deptType["IPD"]++; if (isMale) stats.ipdGender["ຊາຍ"]++; else stats.ipdGender["ຍິງ"]++; }
      else { stats.deptType["OPD"]++; if (isMale) stats.opdGender["ຊາຍ"]++; else stats.opdGender["ຍິງ"]++; }
      if (vSite === "Onsite") stats.site["Onsite"]++; else stats.site["In-site"]++;
      var hour = v.Date ? parseInt(v.Date.split("T")[1]) : -1;
      if (hour >= 8 && hour < 16) stats.timeSlot["08:00-16:00"]++;
      else if (hour >= 16 && hour < 21) stats.timeSlot["16:00-21:00"]++;
      else if (hour >= 0) stats.timeSlot["21:00-08:00"]++;
      (v.Services_List || "").split(",").forEach(function(s) { var t = s.trim(); if (t) stats.services[t] = (stats.services[t] || 0) + 1; });
      (v.Mapped_Specialist || "").split(",").forEach(function(s) { var t = s.trim(); if (t) stats.specialist[t] = (stats.specialist[t] || 0) + 1; });
      (v.Revenue_Group || "").split(",").forEach(function(s) { var t = s.trim(); if (t) stats.revenue[t] = (stats.revenue[t] || 0) + 1; });
    });
    var sortedSrv = Object.keys(stats.services).map(function(k) { return [k, stats.services[k]]; }).sort(function(a,b){return b[1]-a[1];});
    stats.topServices = { labels: [], data: [] };
    sortedSrv.slice(0,10).forEach(function(x){ stats.topServices.labels.push(x[0]); stats.topServices.data.push(x[1]); });
    var sortedProv = Object.keys(stats.provinces).map(function(k){ return [k, stats.provinces[k]]; }).sort(function(a,b){return b[1]-a[1];});
    stats.topProvinces = { labels: [], data: [] };
    sortedProv.slice(0,5).forEach(function(x){ stats.topProvinces.labels.push(x[0]); stats.topProvinces.data.push(x[1]); });
    return stats;
  } catch(e) { return {}; }
}

function getReportData(startDateStr, endDateStr) {
  try {
    var visits = sbFetch("GET", "Visits", "Date=gte." + startDateStr + "T00:00:00&Date=lte." + endDateStr + "T23:59:59&select=*&order=Date.desc");
    var patients = sbFetch("GET", "Patients", "select=Patient_ID,Gender,Age,Organization_ID,Insurance_Code");
    var pDict = {};
    (patients || []).forEach(function(p) { pDict[p.Patient_ID] = p; });
    var allVisits = sbFetch("GET", "Visits", "select=Visit_ID,Patient_ID&order=Date.asc");
    var firstVisitMap = {};
    (allVisits || []).forEach(function(v) { if (!firstVisitMap[v.Patient_ID]) firstVisitMap[v.Patient_ID] = v.Visit_ID; });
    return (visits || []).map(function(v) {
      var p = pDict[v.Patient_ID] || {};
      var dt = (v.Date || "").split("T");
      return {
        date: dt[0] || "-", time: (dt[1] || "-").substring(0,5),
        id: v.Patient_ID, name: v.Patient_Name,
        gender: p.Gender || "-", age: p.Age || "-",
        type: v.Visit_Type || "OPD",
        status: (firstVisitMap[v.Patient_ID] === v.Visit_ID) ? "ໃໝ່" : "ເກົ່າ",
        category: (p.Organization_ID || p.Insurance_Code) ? "INS/Corp" : "GN"
      };
    });
  } catch(e) { return []; }
}

// ─── ORGANIZATIONS ─────────────────────────────────────────────
function getOrganizations() {
  try {
    var rows = sbFetch("GET", "Organizations", "select=*&order=Org_ID.asc");
    return (rows || []).map(function(r) {
      return { rowIdx: r.id || r.Org_ID, orgId: r.Org_ID, cusId: r.Cus_ID_Ex, name: r.Name, orgName: r.Org_Name, orgCode: r.Org_Code, discount: r.Discount, status: r.Status };
    });
  } catch(e) { return []; }
}

function getActiveOrgs() {
  try {
    var rows = sbFetch("GET", "Organizations", "Status=eq.Active&select=Org_Code,Org_Name");
    return (rows || []).map(function(r) { return { id: r.Org_Code, name: r.Org_Name }; });
  } catch(e) { return []; }
}

function saveOrganization(f) {
  try {
    if (f.orgId) {
      sbFetch("PATCH", "Organizations", "Org_ID=eq." + f.orgId, { Cus_ID_Ex: f.cusId, Name: f.name, Org_Name: f.orgName, Org_Code: f.orgCode, Discount: f.discount });
    } else {
      sbFetch("POST", "Organizations", null, { Org_ID: "ORG" + new Date().getTime().toString().slice(-6), Cus_ID_Ex: f.cusId, Name: f.name, Org_Name: f.orgName, Org_Code: f.orgCode, Discount: f.discount, Status: "Active" });
    }
    return { success: true };
  } catch(e) { return { success: false }; }
}

function toggleOrgStatus(orgId, currentStatus) {
  try {
    sbFetch("PATCH", "Organizations", "Org_ID=eq." + orgId, { Status: currentStatus === "Active" ? "Inactive" : "Active" });
    return { success: true };
  } catch(e) { return { success: false }; }
}

function fetchOrgNameById(orgCode) {
  try {
    var rows = sbFetch("GET", "Organizations", "Org_Code=eq." + encodeURIComponent(orgCode) + "&Status=eq.Active");
    if (!rows || rows.length === 0) return { success: false };
    return { success: true, orgName: rows[0].Org_Name, discount: rows[0].Discount };
  } catch(e) { return { success: false }; }
}

// ─── SERVICES MASTER ───────────────────────────────────────────
function getServicesMaster() {
  try {
    var rows = sbFetch("GET", "Service_Lists", "select=*&order=ID.asc");
    return (rows || []).map(function(r) { return { id: r.ID, revenue: r.Revenue_Group, specialist: r.Mapped_Specialist, service: r.Services_List }; });
  } catch(e) { return []; }
}

function saveServiceMaster(f) {
  try {
    if (f.s_id) sbFetch("PATCH", "Service_Lists", "ID=eq." + f.s_id, { Revenue_Group: f.s_rev, Mapped_Specialist: f.s_spec, Services_List: f.s_serv });
    else sbFetch("POST", "Service_Lists", null, { ID: "SRV" + new Date().getTime(), Revenue_Group: f.s_rev, Mapped_Specialist: f.s_spec, Services_List: f.s_serv });
    return getServicesMaster();
  } catch(e) { return []; }
}

function deleteServiceMaster(id) {
  try { sbFetch("DELETE", "Service_Lists", "ID=eq." + id); return getServicesMaster(); }
  catch(e) { return []; }
}

// ─── LOCATIONS MASTER ──────────────────────────────────────────
function getLocationsMaster() {
  try {
    var rows = sbFetch("GET", "Locations", "select=*&order=Province.asc");
    return (rows || []).map(function(r) { return { id: r.ID, district: r.District, province: r.Province }; });
  } catch(e) { return []; }
}

function saveLocationMaster(f) {
  try {
    if (f.l_id) sbFetch("PATCH", "Locations", "ID=eq." + f.l_id, { District: f.l_dist, Province: f.l_prov });
    else sbFetch("POST", "Locations", null, { ID: "LOC" + new Date().getTime(), District: f.l_dist, Province: f.l_prov });
    return getLocationsMaster();
  } catch(e) { return []; }
}

function deleteLocationMaster(id) {
  try { sbFetch("DELETE", "Locations", "ID=eq." + id); return getLocationsMaster(); }
  catch(e) { return []; }
}

// ─── APPOINTMENTS ──────────────────────────────────────────────
function getAppointments() {
  try {
    var rows = sbFetch("GET", "Appointments", "select=*&order=Appt_Date.asc");
    var today = new Date().toISOString().split("T")[0];
    return (rows || []).map(function(r) {
      var st = r.Status;
      if (st === "Pending" && r.Appt_Date && r.Appt_Date.split("T")[0] < today) st = "Overdue";
      return { id: r.Appt_ID, patientId: r.Patient_ID, patientName: r.Patient_Name, date: r.Appt_Date, time: r.Appt_Time, type: r.Type, reason: r.Reason, doctor: r.Doctor, status: st };
    });
  } catch(e) { return []; }
}

function saveAppointment(f) {
  try {
    if (f.a_id) {
      sbFetch("PATCH", "Appointments", "Appt_ID=eq." + f.a_id, { Patient_ID: f.a_patient_id, Patient_Name: f.a_patient_name, Appt_Date: f.a_date, Appt_Time: f.a_time, Type: f.a_type, Reason: f.a_reason, Doctor: f.a_doctor });
    } else {
      sbFetch("POST", "Appointments", null, { Appt_ID: "APT" + new Date().getTime(), Patient_ID: f.a_patient_id, Patient_Name: f.a_patient_name, Appt_Date: f.a_date, Appt_Time: f.a_time, Type: f.a_type, Reason: f.a_reason, Doctor: f.a_doctor, Status: "Pending" });
    }
    return { success: true };
  } catch(e) { return { success: false }; }
}

function deleteAppointment(id) {
  try { sbFetch("DELETE", "Appointments", "Appt_ID=eq." + id); return { success: true }; }
  catch(e) { return { success: false }; }
}

function updateAppointmentStatus(id, status) {
  try { sbFetch("PATCH", "Appointments", "Appt_ID=eq." + id, { Status: status }); return { success: true }; }
  catch(e) { return { success: false }; }
}

function getUpcomingAlerts() {
  try {
    var today = new Date().toISOString().split("T")[0];
    var rows = sbFetch("GET", "Appointments", "Status=eq.Pending&select=*");
    var alerts = [];
    (rows || []).forEach(function(r) {
      if (!r.Appt_Date) return;
      var aDate = r.Appt_Date.split("T")[0];
      var diffMs = new Date(aDate).getTime() - new Date(today).getTime();
      var diffDays = Math.round(diffMs / 86400000);
      if (diffDays < 0 || (diffDays >= 0 && diffDays <= 14)) {
        alerts.push({ id: r.Appt_ID, patientName: r.Patient_Name, date: aDate, time: r.Appt_Time, type: r.Type, isOverdue: diffDays < 0, daysOut: diffDays });
      }
    });
    alerts.sort(function(a,b){ return a.daysOut - b.daysOut; });
    return alerts;
  } catch(e) { return []; }
}

// ─── VACCINES MASTER ───────────────────────────────────────────
function getVaccinesMaster() {
  try {
    var rows = sbFetch("GET", "Vaccines_Master", "select=*&order=Vac_ID.asc");
    return (rows || []).map(function(r) { return { id: r.Vac_ID, name: r.Vaccine_Name, disease: r.Disease, doses: r.Total_Doses, interval: r.Interval_Days || 0 }; });
  } catch(e) { return []; }
}

function saveVaccineMaster(f) {
  try {
    if (f.v_id) sbFetch("PATCH", "Vaccines_Master", "Vac_ID=eq." + f.v_id, { Vaccine_Name: f.v_name, Disease: f.v_disease, Total_Doses: f.v_doses, Interval_Days: f.v_interval || 0 });
    else sbFetch("POST", "Vaccines_Master", null, { Vac_ID: "VAC" + new Date().getTime(), Vaccine_Name: f.v_name, Disease: f.v_disease, Total_Doses: f.v_doses, Interval_Days: f.v_interval || 0 });
    return { success: true };
  } catch(e) { return { success: false }; }
}

function deleteVaccineMaster(id) {
  try { sbFetch("DELETE", "Vaccines_Master", "Vac_ID=eq." + id); return { success: true }; }
  catch(e) { return { success: false }; }
}

function getPatientVaccines() {
  try {
    var rows = sbFetch("GET", "Patient_Vaccines", "select=*&order=Date_Given.desc");
    return (rows || []).map(function(r) {
      return { id: r.Record_ID, patientId: r.Patient_ID, patientName: r.Patient_Name, vaccineName: r.Vaccine_Name, dose: r.Dose_Number, lot: r.Lot_Number, dateGiven: r.Date_Given || "-", nextDue: r.Next_Due_Date || "-", nextDueRaw: r.Next_Due_Date || "", givenBy: r.Given_By };
    });
  } catch(e) { return []; }
}

function savePatientVaccine(f) {
  try {
    sbFetch("POST", "Patient_Vaccines", null, { Record_ID: "PV" + new Date().getTime(), Patient_ID: f.pv_patient_id, Patient_Name: f.pv_patient_name, Vaccine_Name: f.pv_vaccine, Dose_Number: f.pv_dose, Lot_Number: f.pv_lot, Date_Given: f.pv_date, Next_Due_Date: f.pv_next_date, Given_By: f.pv_doctor });
    if (f.pv_auto_appt === "yes" && f.pv_next_date) {
      sbFetch("POST", "Appointments", null, { Appt_ID: "APT" + new Date().getTime(), Patient_ID: f.pv_patient_id, Patient_Name: f.pv_patient_name, Appt_Date: f.pv_next_date, Appt_Time: "09:00", Type: "Vaccine", Reason: "ສັກວັກຊີນ " + f.pv_vaccine + " ເຂັມທີ " + (parseInt(f.pv_dose) + 1), Doctor: f.pv_doctor, Status: "Pending" });
    }
    return { success: true };
  } catch(e) { return { success: false }; }
}

function deletePatientVaccine(id) {
  try { sbFetch("DELETE", "Patient_Vaccines", "Record_ID=eq." + id); return { success: true }; }
  catch(e) { return { success: false }; }
}

// ─── DRUGS MASTER ──────────────────────────────────────────────
function getDrugsMaster() {
  try {
    var rows = sbFetch("GET", "Drugs_Master", "select=*&order=Drug_ID.asc");
    return (rows || []).map(function(r) { return { id: r.Drug_ID, name: r.Drug_Name, desc: r.Description }; });
  } catch(e) { return []; }
}

function saveDrugMaster(f) {
  try {
    if (f.dr_id) sbFetch("PATCH", "Drugs_Master", "Drug_ID=eq." + f.dr_id, { Drug_Name: f.dr_name, Description: f.dr_desc });
    else sbFetch("POST", "Drugs_Master", null, { Drug_ID: "DRG" + new Date().getTime(), Drug_Name: f.dr_name, Description: f.dr_desc });
    return getDrugsMaster();
  } catch(e) { return []; }
}

function deleteDrugMaster(id) {
  try { sbFetch("DELETE", "Drugs_Master", "Drug_ID=eq." + id); return getDrugsMaster(); }
  catch(e) { return []; }
}

function importDrugsFromExcel(dataArray) {
  try {
    var rows = dataArray.filter(function(r){ return r && r[0] && !/drug.?name|ຊື່ຢາ|^no$|^#$/i.test(r[0].toString()); })
      .map(function(r){ return { Drug_ID: "DRG" + new Date().getTime() + Math.random().toString(36).slice(-4), Drug_Name: r[0].toString().trim(), Description: r[1] ? r[1].toString().trim() : "" }; });
    if (rows.length > 0) sbFetch("POST", "Drugs_Master", null, rows);
    return { success: true, count: rows.length };
  } catch(e) { return { success: false }; }
}

// ─── LABS MASTER ───────────────────────────────────────────────
function getLabsMaster() {
  try {
    var rows = sbFetch("GET", "Labs_Master", "select=*&order=Lab_ID.asc");
    return (rows || []).map(function(r) { return { id: r.Lab_ID, name: r.Lab_Name, desc: r.Description }; });
  } catch(e) { return []; }
}

function saveLabMaster(f) {
  try {
    if (f.lb_id) sbFetch("PATCH", "Labs_Master", "Lab_ID=eq." + f.lb_id, { Lab_Name: f.lb_name, Description: f.lb_desc });
    else sbFetch("POST", "Labs_Master", null, { Lab_ID: "LAB" + new Date().getTime(), Lab_Name: f.lb_name, Description: f.lb_desc });
    return getLabsMaster();
  } catch(e) { return []; }
}

function deleteLabMaster(id) {
  try { sbFetch("DELETE", "Labs_Master", "Lab_ID=eq." + id); return getLabsMaster(); }
  catch(e) { return []; }
}

function importLabsFromExcel(dataArray) {
  try {
    var rows = dataArray.filter(function(r){ return r && r[0] && !/lab.?name|ຊື່ແລັບ|^no$|^#$/i.test(r[0].toString()); })
      .map(function(r){ return { Lab_ID: "LAB" + new Date().getTime() + Math.random().toString(36).slice(-4), Lab_Name: r[0].toString().trim(), Description: r[1] ? r[1].toString().trim() : "" }; });
    if (rows.length > 0) sbFetch("POST", "Labs_Master", null, rows);
    return { success: true, count: rows.length };
  } catch(e) { return { success: false }; }
}