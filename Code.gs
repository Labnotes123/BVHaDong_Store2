function doGet(e) {
  return ContentService.createTextOutput("API Kho Bệnh Viện đang chạy (Mode: TEXT/JSON)...").setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  var output = { success: false, msg: "Lỗi khởi tạo" };
  try {
    var jsonString = e.postData.contents;
    var params = JSON.parse(jsonString);
    var action = params.action;

    if (action === 'login') output = login(params.u, params.p);
    else if (action === 'getInitialData') output = getInitialData(params.isAdmin);
    else if (action === 'saveUser') output = saveUser(params.data);
    else if (action === 'deleteUser') output = deleteUser(params.user);
    else if (action === 'saveDMBatch') output = saveDMBatch(params.dataList);
    else if (action === 'updateDMSingle') output = updateDMSingle(params.data, params.oldKey);
    else if (action === 'deleteDM') output = deleteDM(params.key);
    else if (action === 'renameMachine') output = renameMachine(params.kho, params.oldName, params.newName);
    else if (action === 'processImport') output = processImport(params.dataList);
    else if (action === 'processExport') output = processExport(params.data);
    else if (action === 'getReport') output = getInventoryAndHistory(params.from, params.to);
    else if (action === 'changePassword') output = changePassword(params.user, params.oldPass, params.newPass);
    else if (action === 'saveTender') output = saveTender(params.data);
    else if (action === 'toggleTender') output = toggleTender(params.rowIndex, params.status);
    else if (action === 'getTenderData') output = getTenderData();
    else if (action === 'addMasterItem') output = addMasterItem(params.itemType, params.value);
    else if (action === 'deleteMasterItem') output = deleteMasterItem(params.itemType, params.value);
    else if (action === 'updateMasterItem') output = updateMasterItem(params.itemType, params.oldValue, params.newValue);
    else output = { success: false, msg: "Action not found: " + action };

  } catch (err) {
    output = { success: false, msg: "Lỗi Server Backend: " + err.toString() };
  }

  return ContentService.createTextOutput(JSON.stringify(output)).setMimeType(ContentService.MimeType.TEXT);
}

function login(u, p) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("TAI_KHOAN");
  var data = sheet.getRange(2, 1, sheet.getLastRow()-1, 5).getValues();
  for (var i = 0; i < data.length; i++) {
    var user = String(data[i][0]).trim().toLowerCase();
    var pass = String(data[i][1]).trim();
    var role = String(data[i][3]).trim();
    var status = String(data[i][4]).trim();
    if (user === String(u).trim().toLowerCase() && pass === String(p).trim()) {
      if (status.toLowerCase() === 'block') return { success: false, msg: "Tài khoản bị khóa!" };
      return { success: true, name: data[i][2], user: data[i][0], role: role };
    }
  }
  return { success: false, msg: "Sai tài khoản/mật khẩu!" };
}

function getInitialData(isAdmin) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetDM = ss.getSheetByName("DM");
  var dmRaw = sheetDM.getLastRow() > 1 ? sheetDM.getRange(2, 1, sheetDM.getLastRow()-1, 9).getValues() : [];
  var users = [];
  var tenderData = getTenderData();
  var masters = getMasterLists();
  if (isAdmin) {
    var sheetUser = ss.getSheetByName("TAI_KHOAN");
    if(sheetUser.getLastRow() > 1) users = sheetUser.getRange(2, 1, sheetUser.getLastRow()-1, 5).getValues();
  }
  return { dm: dmRaw, users: users, tenders: tenderData.tenders || [], masters: masters };
}

function saveUser(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); var sheet = ss.getSheetByName("TAI_KHOAN");
  var list = sheet.getRange(2, 1, sheet.getLastRow(), 1).getValues().flat();
  var rowIndex = -1;
  for(var i=0; i<list.length; i++){ if(String(list[i]).toLowerCase() === String(data.user).toLowerCase()){ rowIndex = i + 2; break; } }
  if (rowIndex > 0) { sheet.getRange(rowIndex, 2, 1, 4).setValues([[data.pass, data.name, data.role, data.status]]); return {success: true, msg: "Đã cập nhật!"}; }
  else { sheet.appendRow([data.user, data.pass, data.name, data.role, data.status]); return {success: true, msg: "Đã thêm mới!"}; }
}

function deleteUser(u) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); var sheet = ss.getSheetByName("TAI_KHOAN");
  var list = sheet.getRange(2, 1, sheet.getLastRow(), 1).getValues().flat();
  for(var i=0; i<list.length; i++){ if(String(list[i]).toLowerCase() === String(u).toLowerCase()){ sheet.deleteRow(i + 2); return {success: true, msg: "Đã xóa!"}; } }
  return {success: false, msg: "Không tìm thấy!"};
}

function saveDMBatch(dataList) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); var sheet = ss.getSheetByName("DM");
  var list = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow()-1, 3).getValues() : [];
  var count = 0;
  dataList.forEach(data => {
    var exists = false;
    for(var i=0; i<list.length; i++){ if(list[i][0] == data.kho && list[i][1] == data.may && list[i][2] == data.ten) { exists = true; break; } }
    if(!exists) { sheet.appendRow([data.kho, data.may, data.ten, data.loai, data.dv, data.minR, data.minY, data.hanR, data.hanY]); list.push([data.kho, data.may, data.ten]); count++; }
  });
  return {success: true, msg: "Đã thêm " + count + " dòng."};
}

function updateDMSingle(data, oldKey) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); var sheet = ss.getSheetByName("DM");
  var list = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow()-1, 3).getValues() : [];
  for(var i=0; i<list.length; i++) {
    var key = list[i][0] + "|" + list[i][1] + "|" + list[i][2];
    if (key === oldKey) {
      sheet.getRange(i + 2, 1, 1, 9).setValues([[data.kho, data.may, data.ten, data.loai, data.dv, data.minR, data.minY, data.hanR, data.hanY]]);
      return {success: true, msg: "Đã cập nhật!"};
    }
  }
  return {success: false, msg: "Lỗi: Không tìm thấy dòng cũ."};
}

function deleteDM(keyDel) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); var sheet = ss.getSheetByName("DM");
  var list = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow()-1, 3).getValues() : [];
  for(var i=0; i<list.length; i++) { if ((list[i][0]+"|"+list[i][1]+"|"+list[i][2]) === keyDel) { sheet.deleteRow(i + 2); return {success: true, msg: "Đã xóa!"}; } }
  return {success: false, msg: "Lỗi!"};
}

function renameMachine(kho, oldName, newName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); var sheet = ss.getSheetByName("DM");
  var data = sheet.getRange(2, 1, sheet.getLastRow()-1, 2).getValues(); var count = 0;
  for(var i=0; i<data.length; i++) { if(data[i][0] == kho && data[i][1] == oldName) { sheet.getRange(i + 2, 2).setValue(newName); count++; } }
  return {success: true, msg: "Đã cập nhật " + count + " dòng."};
}

function processImport(dataList) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("NHAP"); var ts = new Date();
  var tonKhoUpdates = [];
  var tenderWarnings = [];
  var tenderUsageMap = loadTenderUsage_();

  dataList.forEach(data => {
    var tenderKey = normalizeTenderKey_(data.kho, data.may, data.tenHC, data.hangSX, data.nhaCC, data.namThau);
    var tender = findTenderByKey_(tenderKey);
    if (!tender || tender.status !== "Active") {
      throw new Error("Hãng/NCC/Năm thầu không hợp lệ hoặc đang tắt.");
    }

    var useCalc = calcTenderUse_(data, tender);
    var useThau = useCalc.useQtyBoxes; // số hộp/đơn vị thầu cần dùng
    var currentUsedThau = tenderUsageMap[tenderKey] || 0;
    if (currentUsedThau + useThau - tender.soLuongThau > 0.0001) {
      throw new Error("Vượt quá số lượng thầu còn lại cho " + data.tenHC + " (" + tender.hang + " - " + tender.ncc + ")");
    }

    var lot2 = data.loaiNhap === 'R1R2' ? data.lotR2 : ""; var hsd2 = data.loaiNhap === 'R1R2' ? data.hsdR2 : ""; var sl2 = data.loaiNhap === 'R1R2' ? data.slR2 : "";
    sheet.appendRow([
      ts, new Date(data.ngayNhap), data.nguoi, data.kho, data.may, data.tenHC,
      data.lotR1, new Date(data.hsdR1), data.slR1,
      lot2, hsd2 ? new Date(hsd2) : "", sl2,
      data.donVi,
      data.hangSX, data.nhaCC, data.namThau, data.dvThau, toSafeNumber_(data.heSoQD), useThau, tenderKey
    ]);

    tenderUsageMap[tenderKey] = currentUsedThau + useThau;

    var remainPct = tender.soLuongThau > 0 ? (1 - (tenderUsageMap[tenderKey] / tender.soLuongThau)) : 1;
    if (remainPct <= 0.1) {
      tenderWarnings.push("Hãng " + tender.hang + " / NCC " + tender.ncc + " còn " + Math.max(0, tender.soLuongThau - tenderUsageMap[tenderKey]).toFixed(2) + " (đv thầu)");
    }
    if (useCalc.warnMsg) tenderWarnings.push(useCalc.warnMsg);

    tonKhoUpdates.push({
      kho: data.kho,
      may: data.may,
      ten: data.tenHC,
      part: 'R1',
      lot: data.lotR1,
      hsd: data.hsdR1,
      qty: toSafeNumber_(data.slR1),
      donVi: data.donVi,
      hang: data.hangSX,
      ncc: data.nhaCC
    });

    if (data.loaiNhap === 'R1R2' && data.lotR2) {
      tonKhoUpdates.push({
        kho: data.kho,
        may: data.may,
        ten: data.tenHC,
        part: 'R2',
        lot: data.lotR2,
        hsd: data.hsdR2,
        qty: toSafeNumber_(data.slR2),
        donVi: data.donVi,
        hang: data.hangSX,
        ncc: data.nhaCC
      });
    }
  });

  applyTonKhoUpdates_(tonKhoUpdates);
  var warnMsg = tenderWarnings.length ? (" | Cảnh báo: " + tenderWarnings.join("; ")) : "";
  return {success: true, msg: "Nhập kho thành công " + dataList.length + " phiếu!" + warnMsg};
}

function processExport(data) {
  var today = new Date(); today.setHours(0,0,0,0); var exportDate = new Date(data.ngayXuat); exportDate.setHours(0,0,0,0);
  if (exportDate < today && data.lyDo !== "Quên không xuất kho") return {success: false, msg: "Lỗi: Ngày quá khứ sai lý do."};
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("XUAT");
  sheet.appendRow([new Date(), new Date(data.ngayXuat), data.nguoi, data.kho, data.may, data.tenHC, data.lotR1, data.slR1, data.lotR2, data.slR2, data.lyDo, data.hangSX || "", data.nhaCC || ""]);

  var tonKhoUpdates = [];
  if (data.lotR1) {
    tonKhoUpdates.push({
      kho: data.kho,
      may: data.may,
      ten: data.tenHC,
      part: 'R1',
      lot: data.lotR1,
      hsd: "",
      qty: -toSafeNumber_(data.slR1),
      donVi: "",
      hang: data.hangSX,
      ncc: data.nhaCC
    });
  }
  if (data.lotR2) {
    tonKhoUpdates.push({
      kho: data.kho,
      may: data.may,
      ten: data.tenHC,
      part: 'R2',
      lot: data.lotR2,
      hsd: "",
      qty: -toSafeNumber_(data.slR2),
      donVi: "",
      hang: data.hangSX,
      ncc: data.nhaCC
    });
  }

  applyTonKhoUpdates_(tonKhoUpdates);
  return {success: true, msg: "Xuất kho thành công!"};
}

function changePassword(user, oldP, newP) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); var sheet = ss.getSheetByName("TAI_KHOAN");
  var data = sheet.getRange(2, 1, sheet.getLastRow()-1, 2).getValues();
  for(var i=0; i<data.length; i++){
    if(String(data[i][0]).toLowerCase() === String(user).toLowerCase()){
      if(String(data[i][1])===String(oldP)){
        sheet.getRange(i+2,2).setValue(newP);
        return {success: true, msg: "SUCCESS"};
      } else return {success: false, msg: "Mật khẩu cũ sai!"};
    }
  }
  return {success: false, msg: "Lỗi!"};
}

// === TENDER (QUẢN LÝ THẦU) ===
function getTenderSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("THAU");
  if (!sheet) sheet = ss.insertSheet("THAU");
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Kho", "May", "Ten_HC", "Hang_SX", "Nha_CC", "Nam_Thau", "Don_Vi_Thau", "Don_Vi_SD", "He_So_QD", "So_Luong_Thau", "Trang_Thai", "Ghi_Chu", "Created_At", "Updated_At", "R1_Per_Box", "R2_Per_Box"]);
  }
  return sheet;
}

function normalizeTenderKey_(kho, may, ten, hang, ncc, nam) {
  return [normalizeStr_(kho), normalizeStr_(may), normalizeStr_(ten), normalizeStr_(hang), normalizeStr_(ncc), normalizeStr_(nam)].join("|").toLowerCase();
}

function loadTenderUsage_() {
  var sheetNhap = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("NHAP");
  var usage = {};
  if (!sheetNhap || sheetNhap.getLastRow() <= 1) return usage;
  var data = sheetNhap.getRange(2, 1, sheetNhap.getLastRow()-1, sheetNhap.getLastColumn()).getValues();
  data.forEach(function(r) {
    var tenderKey = r[19]; // col 20 in NHAP (TenderKey)
    var usedThau = toSafeNumber_(r[18]); // col 19: SoLuongThauNhap (quy đổi về DV thầu)
    if (tenderKey) {
      if (!usage[tenderKey]) usage[tenderKey] = 0;
      usage[tenderKey] += usedThau;
    }
  });
  return usage;
}

function getTenderData() {
  var sheet = getTenderSheet_();
  var lastRow = sheet.getLastRow();
  var rows = lastRow > 1 ? sheet.getRange(2, 1, lastRow-1, sheet.getLastColumn()).getValues() : [];
  var usage = loadTenderUsage_();
  var out = [];
  rows.forEach(function(r, idx) {
    var kho = r[0], may = r[1], ten = r[2], hang = r[3], ncc = r[4];
    var nam = r[5]; var dvThau = r[6]; var dvSd = r[7];
    var heso = toSafeNumber_(r[8]);
    var slThau = toSafeNumber_(r[9]);
    var status = normalizeStr_(r[10]) || "Active";
    var note = r[11];
    var createdAt = r[12]; var updatedAt = r[13];
    var r1pb = toSafeNumber_(r[14]) || 1;
    var r2pb = toSafeNumber_(r[15]) || 1;
    var tenderKey = normalizeTenderKey_(kho, may, ten, hang, ncc, nam);
    var usedThau = usage[tenderKey] || 0;
    var remainThau = slThau - usedThau;
    var slQuyDoi = slThau * heso;
    var usedQuyDoi = usedThau * heso;
    var remainQuyDoi = slQuyDoi - usedQuyDoi;
    out.push({
      rowIndex: idx + 2,
      kho: kho, may: may, ten: ten,
      hang: hang, ncc: ncc, nam: nam,
      dvThau: dvThau, dvSd: dvSd, heso: heso,
      soLuongThau: slThau,
      soLuongQuyDoi: slQuyDoi,
      usedThau: usedThau,
      usedQuyDoi: usedQuyDoi,
      remainingThau: remainThau,
      remainingQuyDoi: remainQuyDoi,
      r1PerBox: r1pb,
      r2PerBox: r2pb,
      status: status,
      note: note,
      tenderKey: tenderKey,
      createdAt: createdAt,
      updatedAt: updatedAt
    });
  });
  return { success: true, tenders: out };
}

function saveTender(data) {
  var sheet = getTenderSheet_();
  var rows = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow()-1, sheet.getLastColumn()).getValues() : [];
  var key = normalizeTenderKey_(data.kho, data.may, data.ten, data.hang, data.ncc, data.nam);
  var now = new Date();

  // Check duplicate active tender
  for (var i = 0; i < rows.length; i++) {
    var existing = normalizeTenderKey_(rows[i][0], rows[i][1], rows[i][2], rows[i][3], rows[i][4], rows[i][5]);
    var status = normalizeStr_(rows[i][10]) || "Active";
    if (existing === key && status === "Active" && data.rowIndex !== i + 2) {
      return { success: false, msg: "Đã tồn tại thầu Active cho cùng Hãng/NCC/Năm!" };
    }
  }

  var rowValues = [
    data.kho, data.may, data.ten, data.hang, data.ncc,
    data.nam, data.dvThau, data.dvSd, toSafeNumber_(data.heso), toSafeNumber_(data.soLuongThau),
    normalizeStr_(data.status) || "Active",
    data.note || "",
    data.rowIndex ? rows[data.rowIndex - 2][12] || now : now, // Created_At (keep old if editing)
    now, // Updated_At
    toSafeNumber_(data.r1PerBox) || 1,
    toSafeNumber_(data.r2PerBox) || 1
  ];

  if (data.rowIndex && data.rowIndex > 1) {
    sheet.getRange(data.rowIndex, 1, 1, rowValues.length).setValues([rowValues]);
    return { success: true, msg: "Đã cập nhật thầu!" };
  }

  sheet.appendRow(rowValues);
  return { success: true, msg: "Đã thêm thầu mới!" };
}

function toggleTender(rowIndex, status) {
  var sheet = getTenderSheet_();
  if (!rowIndex || rowIndex < 2) return { success: false, msg: "rowIndex không hợp lệ" };
  sheet.getRange(rowIndex, 11).setValue(status);
  sheet.getRange(rowIndex, 14).setValue(new Date());
  return { success: true, msg: "Đã cập nhật trạng thái" };
}

function getMasterLists() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hangSheet = ss.getSheetByName("HANG_SX");
  if (!hangSheet) { hangSheet = ss.insertSheet("HANG_SX"); hangSheet.appendRow(["Hang_SX"]); }
  var nccSheet = ss.getSheetByName("NHA_CC");
  if (!nccSheet) { nccSheet = ss.insertSheet("NHA_CC"); nccSheet.appendRow(["Nha_CC"]); }
  var hang = hangSheet.getLastRow() > 1 ? hangSheet.getRange(2,1, hangSheet.getLastRow()-1,1).getValues().flat().filter(String) : [];
  var ncc = nccSheet.getLastRow() > 1 ? nccSheet.getRange(2,1, nccSheet.getLastRow()-1,1).getValues().flat().filter(String) : [];
  return {hang: hang, ncc: ncc};
}

function addMasterItem(itemType, value) {
  var val = normalizeStr_(value);
  if (!val) return {success:false, msg:"Giá trị trống"};
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = itemType === 'hang' ? ss.getSheetByName("HANG_SX") : ss.getSheetByName("NHA_CC");
  if (!sheet) {
    sheet = ss.insertSheet(itemType === 'hang' ? "HANG_SX" : "NHA_CC");
    sheet.appendRow([itemType === 'hang' ? "Hang_SX" : "Nha_CC"]);
  }
  var existing = sheet.getLastRow() > 1 ? sheet.getRange(2,1, sheet.getLastRow()-1,1).getValues().flat() : [];
  var exists = existing.some(function(x){ return normalizeStr_(x).toLowerCase() === val.toLowerCase(); });
  if (!exists) sheet.appendRow([val]);
  return {success:true, msg: exists ? "Đã tồn tại" : "Đã thêm"};
}

function deleteMasterItem(itemType, value) {
  var val = normalizeStr_(value);
  if (!val) return {success:false, msg:"Giá trị trống"};
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = itemType === 'hang' ? ss.getSheetByName("HANG_SX") : ss.getSheetByName("NHA_CC");
  if (!sheet || sheet.getLastRow() <= 1) return {success:false, msg:"Không tìm thấy"};
  var data = sheet.getRange(2,1, sheet.getLastRow()-1,1).getValues();
  for (var i=0;i<data.length;i++) {
    if (normalizeStr_(data[i][0]).toLowerCase() === val.toLowerCase()) {
      sheet.deleteRow(i+2);
      return {success:true, msg:"Đã xóa"};
    }
  }
  return {success:false, msg:"Không tìm thấy"};
}

function updateMasterItem(itemType, oldValue, newValue) {
  var oldVal = normalizeStr_(oldValue);
  var newVal = normalizeStr_(newValue);
  if (!oldVal || !newVal) return {success:false, msg:"Giá trị trống"};
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = itemType === 'hang' ? ss.getSheetByName("HANG_SX") : ss.getSheetByName("NHA_CC");
  if (!sheet || sheet.getLastRow() <= 1) return {success:false, msg:"Không tìm thấy"};
  var data = sheet.getRange(2,1, sheet.getLastRow()-1,1).getValues();
  for (var i=0;i<data.length;i++) {
    if (normalizeStr_(data[i][0]).toLowerCase() === oldVal.toLowerCase()) {
      sheet.getRange(i+2,1).setValue(newVal);
      return {success:true, msg:"Đã cập nhật"};
    }
  }
  return {success:false, msg:"Không tìm thấy"};
}

function findTenderByKey_(tenderKey) {
  var td = getTenderData();
  if (!td || !td.tenders) return null;
  for (var i = 0; i < td.tenders.length; i++) {
    if (td.tenders[i].tenderKey === tenderKey) return td.tenders[i];
  }
  return null;
}

function calcTenderUse_(data, tender) {
  var q1 = toSafeNumber_(data.slR1);
  var q2 = toSafeNumber_(data.slR2);
  var r1pb = toSafeNumber_(tender.r1PerBox) || 1;
  var r2pb = toSafeNumber_(tender.r2PerBox) || 1;
  var useBoxes = 0;
  var warn = "";

  if (data.loaiNhap === 'R1R2') {
    var boxR1 = r1pb > 0 ? (q1 / r1pb) : 0;
    var boxR2 = r2pb > 0 ? (q2 / r2pb) : 0;
    useBoxes = Math.max(boxR1, boxR2);
    if (boxR1 && boxR2) {
      var diffRatio = Math.abs(boxR1 - boxR2) / Math.max(boxR1, boxR2);
      if (diffRatio > 0.1) warn = "R1/R2 lệch nhiều, kiểm tra lại số lượng.";
    }
  } else {
    useBoxes = r1pb > 0 ? (q1 / r1pb) : q1;
  }

  return {
    useQtyBoxes: useBoxes,
    useQtyQD: useBoxes * (toSafeNumber_(tender.heso) || 0),
    warnMsg: warn
  };
}

function toSafeNumber_(v) {
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === 'number') return isNaN(v) ? 0 : v;
  var str = String(v).trim();
  if (!str) return 0;
  str = str.replace(/\s+/g, '').replace(/,/g, '.');
  var n = Number(str);
  return isNaN(n) ? 0 : n;
}

function normalizeStr_(s) {
  return String(s == null ? "" : s).trim();
}

function normalizeLotKey_(lot) {
  return normalizeStr_(lot).toUpperCase();
}

function parseDateSafe_(d) {
  if (d === null || d === undefined || d === "") return null;
  if (d instanceof Date) return isNaN(d.getTime()) ? null : new Date(d.getTime());

  var direct = new Date(d);
  if (!isNaN(direct.getTime())) return direct;

  var str = String(d).trim();
  var m = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?$/);
  if (m) {
    var day = Number(m[1]), mon = Number(m[2]) - 1, year = Number(m[3]);
    var hh = Number(m[4] || 0), mm = Number(m[5] || 0), ss = Number(m[6] || 0);
    var parsed = new Date(year, mon, day, hh, mm, ss);
    return isNaN(parsed.getTime()) ? null : parsed;
  }
  return null;
}

function getOrCreateTonKhoSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("TONKHO");
  if (!sheet) {
    sheet = ss.insertSheet("TONKHO");
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Kho", "May", "Ten_HC", "Loai", "Lot", "Han", "SL", "Don_Vi", "Updated_At", "Hang_SX", "Nha_CC"]);
  }
  return sheet;
}

function applyTonKhoUpdates_(updates) {
  if (!updates || updates.length === 0) return;

  var sheet = getOrCreateTonKhoSheet_();
  var rows = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow()-1, 11).getValues() : [];
  var map = {};

  function rowKey_(kho, may, ten, loai, lot, hang, ncc) {
    return normalizeStr_(kho) + "|" + normalizeStr_(may) + "|" + normalizeStr_(ten) + "|" + normalizeStr_(loai) + "|" + normalizeLotKey_(lot) + "|" + normalizeStr_(hang) + "|" + normalizeStr_(ncc);
  }

  for (var i = 0; i < rows.length; i++) {
    map[rowKey_(rows[i][0], rows[i][1], rows[i][2], rows[i][3], rows[i][4], rows[i][9] || "", rows[i][10] || "")] = { idx: i, row: rows[i] };
  }

  var now = new Date();
  var touched = {};
  var appendRows = [];

  updates.forEach(function(u) {
    var qty = toSafeNumber_(u.qty);
    var lot = normalizeStr_(u.lot);
    var hangVal = normalizeStr_(u.hang || u.hangSX || "");
    var nccVal = normalizeStr_(u.ncc || u.nhaCC || "");
    if (!lot || qty === 0) return;

    var key = rowKey_(u.kho, u.may, u.ten, u.part, lot, hangVal, nccVal);
    var found = map[key];

    if (found) {
      var row = found.row;
      row[6] = toSafeNumber_(row[6]) + qty;

      var hsdObj = parseDateSafe_(u.hsd);
      if (hsdObj) row[5] = hsdObj;
      if (normalizeStr_(u.donVi)) row[7] = normalizeStr_(u.donVi);
      row[8] = now;
      if (hangVal) row[9] = hangVal;
      if (nccVal) row[10] = nccVal;

      touched[found.idx] = row;
    } else {
      var hsdObj2 = parseDateSafe_(u.hsd);
      appendRows.push([
        normalizeStr_(u.kho),
        normalizeStr_(u.may),
        normalizeStr_(u.ten),
        normalizeStr_(u.part),
        lot,
        hsdObj2 || "",
        qty,
        normalizeStr_(u.donVi),
        now,
        hangVal,
        nccVal
      ]);
    }
  });

  Object.keys(touched).forEach(function(idxStr) {
    var idx = Number(idxStr);
    sheet.getRange(idx + 2, 1, 1, 11).setValues([touched[idx]]);
  });

  if (appendRows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, appendRows.length, 11).setValues(appendRows);
  }
}

function rebuildTonKhoSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetNhap = ss.getSheetByName("NHAP");
  var sheetXuat = ss.getSheetByName("XUAT");
  var tonSheet = getOrCreateTonKhoSheet_();

  var nhapData = (sheetNhap && sheetNhap.getLastRow() > 1) ? sheetNhap.getRange(2, 1, sheetNhap.getLastRow()-1, sheetNhap.getLastColumn()).getValues() : [];
  var xuatData = (sheetXuat && sheetXuat.getLastRow() > 1) ? sheetXuat.getRange(2, 1, sheetXuat.getLastRow()-1, sheetXuat.getLastColumn()).getValues() : [];

  var map = {};
  function key(kho, may, ten, loai, lot, hang, ncc) {
    return normalizeStr_(kho) + "|" + normalizeStr_(may) + "|" + normalizeStr_(ten) + "|" + normalizeStr_(loai) + "|" + normalizeLotKey_(lot) + "|" + normalizeStr_(hang) + "|" + normalizeStr_(ncc);
  }

  function add(kho, may, ten, loai, lot, hsd, qty, dv, hang, ncc) {
    var lotStr = normalizeStr_(lot);
    var q = toSafeNumber_(qty);
    if (!lotStr || q === 0) return;
    var k = key(kho, may, ten, loai, lotStr, hang, ncc);
    if (!map[k]) {
      var hsdObj = parseDateSafe_(hsd);
      map[k] = {
        kho: normalizeStr_(kho),
        may: normalizeStr_(may),
        ten: normalizeStr_(ten),
        loai: normalizeStr_(loai),
        lot: lotStr,
        han: hsdObj || "",
        sl: 0,
        dv: normalizeStr_(dv),
        hang: normalizeStr_(hang),
        ncc: normalizeStr_(ncc)
      };
    }
    map[k].sl += q;
    var hsdObj2 = parseDateSafe_(hsd);
    if (hsdObj2) map[k].han = hsdObj2;
    if (normalizeStr_(dv)) map[k].dv = normalizeStr_(dv);
    if (normalizeStr_(hang)) map[k].hang = normalizeStr_(hang);
    if (normalizeStr_(ncc)) map[k].ncc = normalizeStr_(ncc);
  }

  nhapData.forEach(function(r) {
    add(r[3], r[4], r[5], 'R1', r[6], r[7], r[8], r[12], r[13], r[14]);
    add(r[3], r[4], r[5], 'R2', r[9], r[10], r[11], r[12], r[13], r[14]);
  });
  xuatData.forEach(function(r) {
    add(r[3], r[4], r[5], 'R1', r[6], "", -toSafeNumber_(r[7]), "", r[11], r[12]);
    add(r[3], r[4], r[5], 'R2', r[8], "", -toSafeNumber_(r[9]), "", r[11], r[12]);
  });

  if (tonSheet.getLastRow() > 1) tonSheet.getRange(2, 1, tonSheet.getLastRow()-1, 11).clearContent();

  var out = [];
  var now = new Date();
  Object.keys(map).forEach(function(k) {
    var it = map[k];
    out.push([it.kho, it.may, it.ten, it.loai, it.lot, it.han, it.sl, it.dv, now, it.hang, it.ncc]);
  });

  if (out.length > 0) tonSheet.getRange(2, 1, out.length, 11).setValues(out);
}

function getInventoryAndHistory(fromStr, toStr) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetNhap = ss.getSheetByName("NHAP");
    var sheetXuat = ss.getSheetByName("XUAT");
    var tonSheet = getOrCreateTonKhoSheet_();

    var nhapData = (sheetNhap && sheetNhap.getLastRow() > 1) ? sheetNhap.getRange(2, 1, sheetNhap.getLastRow()-1, sheetNhap.getLastColumn()).getValues() : [];
    var xuatData = (sheetXuat && sheetXuat.getLastRow() > 1) ? sheetXuat.getRange(2, 1, sheetXuat.getLastRow()-1, sheetXuat.getLastColumn()).getValues() : [];
    if (tonSheet.getLastRow() <= 1) {
      rebuildTonKhoSheet_();
    }
    var tonData = tonSheet.getLastRow() > 1 ? tonSheet.getRange(2, 1, tonSheet.getLastRow()-1, 11).getValues() : [];

    var inventory = {};
    var timeline = [];

    function ensureItem(key) {
      if (!inventory[key]) inventory[key] = { R1: {}, R2: {} };
      return inventory[key];
    }

    function applyLotDelta(stockPart, lotRaw, hsdRaw, qtyDelta, hangRaw, nccRaw) {
      var lk = normalizeLotKey_(lotRaw) + "|" + normalizeStr_(hangRaw) + "|" + normalizeStr_(nccRaw);
      if (!lk) return;

      if (!stockPart[lk]) {
        var hsdObj = parseDateSafe_(hsdRaw);
        stockPart[lk] = {
          lot: normalizeStr_(lotRaw),
          hang: normalizeStr_(hangRaw),
          ncc: normalizeStr_(nccRaw),
          hsd: hsdObj ? formatDate(hsdObj) : "",
          rawHsd: hsdObj ? hsdObj.getTime() : 0,
          sl: 0
        };
      }

      if (!stockPart[lk].lot) stockPart[lk].lot = normalizeStr_(lotRaw);
      if (!stockPart[lk].rawHsd) {
        var hsdObj2 = parseDateSafe_(hsdRaw);
        if (hsdObj2) {
          stockPart[lk].rawHsd = hsdObj2.getTime();
          stockPart[lk].hsd = formatDate(hsdObj2);
        }
      }

      stockPart[lk].sl += qtyDelta;
    }

    tonData.forEach(function(r) {
      var kho = normalizeStr_(r[0]);
      var may = normalizeStr_(r[1]);
      var ten = normalizeStr_(r[2]);
      var loai = normalizeStr_(r[3]).toUpperCase();
      var lot = normalizeStr_(r[4]);
      var hsd = r[5];
      var sl = toSafeNumber_(r[6]);
      var hang = normalizeStr_(r[9]);
      var ncc = normalizeStr_(r[10]);
      if (!kho || !ten || !lot) return;

      var key = kho + "|" + may + "|" + ten;
      var item = ensureItem(key);
      if (loai === 'R2') applyLotDelta(item.R2, lot, hsd, sl, hang, ncc);
      else applyLotDelta(item.R1, lot, hsd, sl, hang, ncc);
    });

    nhapData.forEach(function(r) {
      var kho = normalizeStr_(r[3]);
      var may = normalizeStr_(r[4]);
      var ten = normalizeStr_(r[5]);
      if (!kho || !ten) return;

      var lotR1Raw = normalizeStr_(r[6]);
      var lotR2Raw = normalizeStr_(r[9]);
      var slR1 = toSafeNumber_(r[8]);
      var slR2 = toSafeNumber_(r[11]);
      var hangR = normalizeStr_(r[13]);
      var nccR = normalizeStr_(r[14]);

      var ioType = (lotR1Raw && lotR2Raw) ? 'Cả R1,R2' : (lotR2Raw ? 'R2' : 'R1');

      var ts = parseDateSafe_(r[0]) || new Date();
      var dDoc = parseDateSafe_(r[1]) || ts;

      timeline.push({
        type: 'NHẬP',
        action: 'NHẬP',
        sortTime: ts.getTime(),
        timestampStr: formatDateLong(ts),
        dateDocStr: formatDate(dDoc),
        user: r[2], kho: kho, may: may, ten: ten,
        r1_lot: lotR1Raw, r1_sl: slR1,
        r2_lot: lotR2Raw, r2_sl: slR2,
        ioType: ioType,
        hang: hangR,
        ncc: nccR,
        reason: "Nhập mới"
      });
    });

    xuatData.forEach(function(r) {
      var kho = normalizeStr_(r[3]);
      var may = normalizeStr_(r[4]);
      var ten = normalizeStr_(r[5]);
      if (!kho || !ten) return;

      var lotR1Raw = normalizeStr_(r[6]);
      var lotR2Raw = normalizeStr_(r[8]);
      var slR1 = toSafeNumber_(r[7]);
      var slR2 = toSafeNumber_(r[9]);
      var hangR = normalizeStr_(r[11]);
      var nccR = normalizeStr_(r[12]);

      var ioType = (lotR1Raw && lotR2Raw) ? 'Cả R1,R2' : (lotR2Raw ? 'R2' : 'R1');

      var ts = parseDateSafe_(r[0]) || new Date();
      var dDoc = parseDateSafe_(r[1]) || ts;

      timeline.push({
        type: 'XUẤT',
        action: 'XUẤT',
        sortTime: ts.getTime(),
        timestampStr: formatDateLong(ts),
        dateDocStr: formatDate(dDoc),
        user: r[2], kho: kho, may: may, ten: ten,
        r1_lot: lotR1Raw, r1_sl: slR1,
        r2_lot: lotR2Raw, r2_sl: slR2,
        ioType: ioType,
        hang: hangR,
        ncc: nccR,
        reason: r[10]
      });
    });

    var timelineAsc = timeline.slice().sort(function(a,b){ return a.sortTime - b.sortTime; });
    var runningTracker = {};

    var useFilter = (fromStr && toStr);
    var sDate = useFilter ? parseDateSafe_(fromStr) : null;
    var eDate = useFilter ? parseDateSafe_(toStr) : null;
    var sTs = sDate ? sDate.setHours(0,0,0,0) : 0;
    var eTs = eDate ? eDate.setHours(23,59,59,999) : 0;

    var result = [];

    timelineAsc.forEach(function(item) {
      var d = [];
      var base = item.kho + "|" + item.may + "|" + item.ten;
      var isNhap = (item.action === 'NHẬP');

      d.push(`<b>Loại:</b> ${item.ioType || 'R1'}`);

      if(item.r1_lot){
        var k1 = base + "|R1|" + normalizeLotKey_(item.r1_lot) + "|" + normalizeStr_(item.hang) + "|" + normalizeStr_(item.ncc);
        if(runningTracker[k1] == null) runningTracker[k1] = 0;
        runningTracker[k1] += isNhap ? item.r1_sl : -item.r1_sl;

        var style1 = isNhap ? 'text-success fw-bold' : 'text-danger fw-bold';
        var sign1 = isNhap ? '+' : '-';
        var vendor1 = (item.hang || item.ncc) ? ` <span class="text-muted">(${item.hang || ''}${item.hang && item.ncc ? ' / ' : ''}${item.ncc || ''})</span>` : '';
        d.push(`<b>R1:</b> ${item.r1_lot} <span class="${style1}">(${sign1}${item.r1_sl})</span> <span class="text-muted small">[Tồn: ${runningTracker[k1]}]</span>${vendor1}`);
      }

      if(item.r2_lot){
        var k2 = base + "|R2|" + normalizeLotKey_(item.r2_lot) + "|" + normalizeStr_(item.hang) + "|" + normalizeStr_(item.ncc);
        if(runningTracker[k2] == null) runningTracker[k2] = 0;
        runningTracker[k2] += isNhap ? item.r2_sl : -item.r2_sl;

        var style2 = isNhap ? 'text-success fw-bold' : 'text-danger fw-bold';
        var sign2 = isNhap ? '+' : '-';
        var vendor2 = (item.hang || item.ncc) ? ` <span class="text-muted">(${item.hang || ''}${item.hang && item.ncc ? ' / ' : ''}${item.ncc || ''})</span>` : '';
        d.push(`<b>R2:</b> ${item.r2_lot} <span class="${style2}">(${sign2}${item.r2_sl})</span> <span class="text-muted small">[Tồn: ${runningTracker[k2]}]</span>${vendor2}`);
      }

      item.detailHtml = d.join("<br>");

      if(!useFilter || (item.sortTime >= sTs && item.sortTime <= eTs)){
        if(d.length > 0 && item.kho) result.push(item);
      }
    });

    result.sort(function(a,b){ return b.sortTime - a.sortTime; });

    var stockOut = {};
    Object.keys(inventory).forEach(function(mainKey) {
      stockOut[mainKey] = { R1: {}, R2: {} };
      ['R1', 'R2'].forEach(function(part) {
        Object.keys(inventory[mainKey][part]).forEach(function(lk) {
          var obj = inventory[mainKey][part][lk];
          var lotLabel = obj.lot || lk;
          stockOut[mainKey][part][lk] = {
            lot: lotLabel,
            hsd: obj.hsd || "",
            rawHsd: obj.rawHsd || 0,
            sl: obj.sl,
            hang: obj.hang || "",
            ncc: obj.ncc || ""
          };
        });
      });
    });

    return { stock: stockOut, history: useFilter ? result : result.slice(0,50) };

  } catch (e) {
    return {stock:{}, history:[], error: "Lỗi tính toán Backend: " + e.toString()};
  }
}

function formatDate(d) {
  if (!d) return "";
  try {
    var date = (d instanceof Date) ? d : new Date(d);
    return Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yyyy");
  } catch(e){ return ""; }
}

function formatDateLong(d) {
  if (!d) return "";
  try {
    var date = (d instanceof Date) ? d : new Date(d);
    return Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
  } catch(e){ return ""; }
}
