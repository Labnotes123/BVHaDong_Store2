// API URL CỦA BẠN (Cập nhật link mới từ Backend)
const API_URL = "https://script.google.com/macros/s/AKfycbxpS5BhQv-ZseJQeA0opLOBNbE5JR3BVZ1B1MJJZKvFEiuOqJUJd6UTI9pbc0uLJMmj/exec";

var DB = { dm: [], tenders: [], stock: {}, history: [], users: [], currentUser: "", currentID: "", role: "" };
DB.masters = { hang: [], ncc: [] };
var rowCount = 0; var admRowCount = 0;

// SỬA 3: Thêm cờ đánh dấu trạng thái dữ liệu
var isStockLoaded = false; // Dùng cho tab Xuất (Load 1 lần đầu)
var isReportDirty = true;  // Dùng cho tab Báo cáo (True = cần reload, False = dùng cache)

async function callAPI(action, payload = {}) {
  payload.action = action;
  try {
      const response = await fetch(API_URL, { method: 'POST', body: JSON.stringify(payload) });
      const textData = await response.text();
      try { return JSON.parse(textData); }
      catch (e) { console.error("Lỗi JSON:", textData); return { success: false, msg: "Lỗi dữ liệu" }; }
  } catch (error) {
      console.error("API Error:", error); alert("Lỗi kết nối Server"); return { success: false, msg: error };
  }
}

document.getElementById("loginUser").addEventListener("keydown", function(e) { if (e.key === "Enter") { e.preventDefault(); document.getElementById("loginPass").focus(); } });
document.getElementById("loginPass").addEventListener("keydown", function(e) { if (e.key === "Enter") { e.preventDefault(); doLogin(); } });

function setTodayDate() {
  var d = new Date(), m = ''+(d.getMonth()+1), day = ''+d.getDate(), y = d.getFullYear();
  if(m.length<2) m='0'+m; if(day.length<2) day='0'+day;
  var todayStr = [y, m, day].join('-');
  document.getElementById('xuatNgay').value = todayStr;
  document.getElementById('importRowsContainer').innerHTML = ""; rowCount = 0; addImportRow(todayStr);
}

async function doLogin() {
  var u = document.getElementById('loginUser').value; var p = document.getElementById('loginPass').value;
  if(!u || !p) return document.getElementById('loginMsg').innerText = "Vui lòng nhập đầy đủ!";
  document.getElementById('loading').style.display = 'flex';
  const res = await callAPI('login', {u: u, p: p});
  document.getElementById('loading').style.display = 'none';
  if(res.success) { showApp(res.name, res.user, res.role); }
  else { document.getElementById('loginMsg').innerText = res.msg || "Lỗi đăng nhập"; }
}

function showApp(name, uid, role) {
  DB.currentUser = name; DB.currentID = uid; DB.role = role;
  document.getElementById('userDisplay').innerText = name + (role === 'Admin' ? ' (Admin)' : '');
  document.getElementById('nhapNguoi').value = name; document.getElementById('xuatNguoi').value = name;
  document.getElementById('navItemAdmin').style.display = (role === 'Admin') ? 'block' : 'none';
  document.getElementById('login-screen').style.display = 'none'; document.getElementById('app-screen').style.display = 'block';
  setTodayDate(); loadInitialData();
}

function doLogout() { window.location.reload(); }

async function loadInitialData() {
  var isAdmin = (DB.role === 'Admin');
  const data = await callAPI('getInitialData', {isAdmin: isAdmin});
  if(data) {
     DB.dm = data.dm || []; DB.users = data.users || []; DB.tenders = data.tenders || []; DB.masters = data.masters || {hang:[], ncc:[]};
     renderTenderView();
     renderMasterLists();
     if(isAdmin) { renderAdminTables(); renderTenderAdmin(); }
  }
}

function renderMasterLists() {
  var hangList = document.getElementById('listHang');
  var nccList = document.getElementById('listNCC');
  var hangDl = document.getElementById('dlHangMaster');
  var nccDl = document.getElementById('dlNccMaster');
  var hangArr = (DB.masters && DB.masters.hang) ? DB.masters.hang : [];
  var nccArr = (DB.masters && DB.masters.ncc) ? DB.masters.ncc : [];

  function buildList(arr, type) {
    if(!arr || arr.length === 0) return '<li class="list-group-item text-muted small">Chưa có dữ liệu</li>';
    return arr.map(v => `<li class="list-group-item d-flex justify-content-between align-items-center">${v}<button class="btn btn-sm btn-outline-danger" onclick="deleteMasterItem('${type}','${v.replace(/'/g, "\\'")}')"><i class="fas fa-trash"></i></button></li>`).join('');
  }

  if(hangList) hangList.innerHTML = buildList(hangArr, 'hang');
  if(nccList) nccList.innerHTML = buildList(nccArr, 'ncc');
  if(hangDl) hangDl.innerHTML = hangArr.map(v => `<option value="${v}">`).join('');
  if(nccDl) nccDl.innerHTML = nccArr.map(v => `<option value="${v}">`).join('');
}

async function saveMasterItem(type) {
  var inputId = type === 'hang' ? 'inputHang' : 'inputNCC';
  var input = document.getElementById(inputId);
  if(!input) return;
  var val = (input.value || '').trim();
  if(!val) return alert('Vui lòng nhập giá trị');

  document.getElementById('loading').style.display = 'flex';
  var res = await callAPI('addMasterItem', {itemType: type, value: val});
  document.getElementById('loading').style.display = 'none';
  if(!res || !res.success) return alert(res?.msg || 'Thêm thất bại');

  var targetArr = (DB.masters && DB.masters[type]) ? DB.masters[type] : [];
  var exists = targetArr.some(x => String(x).toLowerCase() === val.toLowerCase());
  if(!exists) targetArr.push(val);
  if(!DB.masters) DB.masters = {hang: [], ncc: []};
  DB.masters[type] = targetArr;
  input.value = '';
  renderMasterLists();
}

async function deleteMasterItem(type, value) {
  if(!value) return;
  if(!confirm('Xóa giá trị này?')) return;
  document.getElementById('loading').style.display = 'flex';
  var res = await callAPI('deleteMasterItem', {itemType: type, value: value});
  document.getElementById('loading').style.display = 'none';
  if(!res || !res.success) return alert(res?.msg || 'Xóa thất bại');
  if(!DB.masters) DB.masters = {hang: [], ncc: []};
  DB.masters[type] = (DB.masters[type] || []).filter(x => String(x).toLowerCase() !== String(value).toLowerCase());
  renderMasterLists();
}

function filterAdminTable(tableId, colIndex, query) {
  var filter = query.toUpperCase();
  var rows = document.getElementById(tableId).getElementsByTagName("tbody")[0].getElementsByTagName("tr");
  for (var i = 0; i < rows.length; i++) {
    var td = rows[i].getElementsByTagName("td")[colIndex];
    if (td) rows[i].style.display = (td.textContent || td.innerText).toUpperCase().indexOf(filter) > -1 ? "" : "none";
  }
}

// --- ADMIN FUNCTIONS ---
function renderAdminTables() { var htmlDM = DB.dm.map(d => { var key = `${d[0]}|${d[1]}|${d[2]}`; return `<tr><td>${d[0]}</td><td>${d[1]}</td><td class="fw-bold">${d[2]}</td><td>${d[3]}</td><td>${d[4]}</td><td class="text-danger text-center fw-bold">${d[5]}</td><td class="text-warning text-center fw-bold">${d[6]}</td><td class="text-danger text-center">${d[7]}</td><td class="text-warning text-center">${d[8]}</td><td><button class="btn btn-outline-primary btn-action btn-sm" onclick="editDM('${key}')"><i class="fas fa-edit"></i></button> <button class="btn btn-outline-danger btn-action btn-sm" onclick="delDM('${key}')"><i class="fas fa-trash"></i></button></td></tr>`; }).join(''); document.getElementById('bodyAdminDM').innerHTML = htmlDM; var htmlUser = DB.users.map(u => { var cls = u[4] === 'Active' ? 'text-success' : 'text-danger fw-bold'; return `<tr><td class="fw-bold">${u[0]}</td><td>***</td><td>${u[2]}</td><td>${u[3]}</td><td class="${cls}">${u[4]}</td><td><button class="btn btn-outline-warning btn-action btn-sm" onclick="editUser('${u[0]}')"><i class="fas fa-edit"></i></button> <button class="btn btn-outline-danger btn-action btn-sm" onclick="delUser('${u[0]}')"><i class="fas fa-trash"></i></button></td></tr>`; }).join(''); document.getElementById('bodyAdminUser').innerHTML = htmlUser; }
function openModalDM() { document.getElementById('admOldKey').value = ""; document.getElementById('titleAdmDM').innerText = "THÊM HÓA CHẤT (HÀNG LOẠT)"; document.getElementById('admRowsContainer').innerHTML = ""; admRowCount = 0; var uniqueUnits = [...new Set(DB.dm.map(r => r[4]))]; var dl = document.getElementById('listAdmUnits'); dl.innerHTML = ""; uniqueUnits.forEach(u => { if(u) dl.innerHTML += `<option value="${u}">`; }); changeAdmKho(); addAdmRow(); new bootstrap.Modal(document.getElementById('modalAdminDM')).show(); }
function editDM(key) { var item = DB.dm.find(d => `${d[0]}|${d[1]}|${d[2]}` === key); if (!item) return; document.getElementById('admOldKey').value = key; document.getElementById('titleAdmDM').innerText = "SỬA THÔNG TIN HÓA CHẤT"; document.getElementById('admKho').value = item[0]; document.getElementById('admRowsContainer').innerHTML = ""; admRowCount = 0; var idx = admRowCount++; var machineOptions = getMachineOptions(item[0]); var html = `<div class="import-row" id="admRow_${idx}"><div class="row g-2"><div class="col-md-2"><label class="small text-muted">Máy</label><input list="listMac_${idx}" class="form-control form-control-sm" id="admMay_${idx}" value="${item[1]}"><datalist id="listMac_${idx}">${machineOptions}</datalist></div><div class="col-md-3"><label class="small text-muted">Tên HC</label><input class="form-control form-control-sm" id="admTen_${idx}" value="${item[2]}"></div><div class="col-md-2"><label class="small text-muted">Loại</label><select class="form-select form-select-sm" id="admLoai_${idx}"><option value="R1R2" ${item[3]=='R1R2'?'selected':''}>R1R2</option><option value="Single" ${item[3]=='Single'?'selected':''}>Single</option></select></div><div class="col-md-1"><label class="small text-muted">ĐV</label><input list="listAdmUnits" class="form-control form-control-sm" id="admDV_${idx}" value="${item[4]}"></div><div class="col-md-1"><label class="small text-danger">SL Đỏ</label><input type="number" class="form-control form-control-sm" id="admMinR_${idx}" value="${item[5]}"></div><div class="col-md-1"><label class="small text-warning">SL Vàng</label><input type="number" class="form-control form-control-sm" id="admMinY_${idx}" value="${item[6]}"></div><div class="col-md-1"><label class="small text-danger">Date Đỏ</label><input type="number" class="form-control form-control-sm" id="admHanR_${idx}" value="${item[7]}"></div><div class="col-md-1"><label class="small text-warning">Date Vàng</label><input type="number" class="form-control form-control-sm" id="admHanY_${idx}" value="${item[8]}"></div></div></div>`; document.getElementById('admRowsContainer').innerHTML = html; new bootstrap.Modal(document.getElementById('modalAdminDM')).show(); }
function changeAdmKho() { var kho = document.getElementById('admKho').value; var machineOptions = getMachineOptions(kho); var datalists = document.querySelectorAll('#admRowsContainer datalist'); datalists.forEach(dl => { dl.innerHTML = machineOptions; }); }
function getMachineOptions(kho) { var machines = [...new Set(DB.dm.filter(r => String(r[0]).trim() == kho).map(r => r[1]))]; return machines.map(m => `<option value="${m}">`).join(''); }
function addAdmRow() { var kho = document.getElementById('admKho').value; var idx = admRowCount++; var machineOptions = getMachineOptions(kho); var html = `<div class="import-row position-relative" id="admRow_${idx}"><i class="fas fa-times btn-remove-row" onclick="document.getElementById('admRow_${idx}').remove()"></i><div class="row g-2"><div class="col-md-2"><label class="small text-muted">Máy</label><input list="listMac_${idx}" class="form-control form-control-sm" id="admMay_${idx}" placeholder="Chọn/Nhập máy"><datalist id="listMac_${idx}">${machineOptions}</datalist></div><div class="col-md-3"><label class="small text-muted">Tên HC</label><input class="form-control form-control-sm" id="admTen_${idx}" placeholder="Tên hóa chất"></div><div class="col-md-2"><label class="small text-muted">Loại</label><select class="form-select form-select-sm" id="admLoai_${idx}"><option value="R1R2">R1R2</option><option value="Single">Single</option></select></div><div class="col-md-1"><label class="small text-muted">ĐV</label><input list="listAdmUnits" class="form-control form-control-sm" id="admDV_${idx}"></div><div class="col-md-1"><label class="small text-danger">SL Đỏ</label><input type="number" class="form-control form-control-sm" id="admMinR_${idx}" value="0"></div><div class="col-md-1"><label class="small text-warning">SL Vàng</label><input type="number" class="form-control form-control-sm" id="admMinY_${idx}" value="0"></div><div class="col-md-1"><label class="small text-danger">Date Đỏ</label><input type="number" class="form-control form-control-sm" id="admHanR_${idx}" value="10"></div><div class="col-md-1"><label class="small text-warning">Date Vàng</label><input type="number" class="form-control form-control-sm" id="admHanY_${idx}" value="30"></div></div></div>`; document.getElementById('admRowsContainer').insertAdjacentHTML('beforeend', html); }
async function saveAdminDMBatch() { var kho = document.getElementById('admKho').value; var rows = document.querySelectorAll('#admRowsContainer .import-row'); if (rows.length === 0) return alert("Chưa có dòng nào!"); var dataList = []; var oldKey = document.getElementById('admOldKey').value; var uniqueUnits = [...new Set(DB.dm.map(r => r[4]))]; function validateUnit(uInput) { if (!uInput) return ""; var match = uniqueUnits.find(u => u.toLowerCase() === uInput.toLowerCase()); return match ? match : uInput; } document.getElementById('loading').style.display = 'flex'; if (oldKey) { var idx = rows[0].id.split('_')[1]; var singleData = { kho: kho, may: document.getElementById('admMay_' + idx).value, ten: document.getElementById('admTen_' + idx).value, loai: document.getElementById('admLoai_' + idx).value, dv: validateUnit(document.getElementById('admDV_' + idx).value), minR: document.getElementById('admMinR_' + idx).value, minY: document.getElementById('admMinY_' + idx).value, hanR: document.getElementById('admHanR_' + idx).value, hanY: document.getElementById('admHanY_' + idx).value }; if(!singleData.may || !singleData.ten) { document.getElementById('loading').style.display = 'none'; return alert("Vui lòng nhập đủ Máy và Tên!"); } await callAPI('updateDMSingle', {data: singleData, oldKey: oldKey}); } else { rows.forEach(row => { var idx = row.id.split('_')[1]; dataList.push({ kho: kho, may: document.getElementById('admMay_' + idx).value, ten: document.getElementById('admTen_' + idx).value, loai: document.getElementById('admLoai_' + idx).value, dv: validateUnit(document.getElementById('admDV_' + idx).value), minR: document.getElementById('admMinR_' + idx).value, minY: document.getElementById('admMinY_' + idx).value, hanR: document.getElementById('admHanR_' + idx).value, hanY: document.getElementById('admHanY_' + idx).value }); }); await callAPI('saveDMBatch', {dataList: dataList}); } bootstrap.Modal.getInstance(document.getElementById('modalAdminDM')).hide(); await loadInitialData(); document.getElementById('loading').style.display = 'none'; }
async function delDM(key) { if(!confirm("Xóa hóa chất này?")) return; document.getElementById('loading').style.display = 'flex'; await callAPI('deleteDM', {key: key}); await loadInitialData(); document.getElementById('loading').style.display = 'none'; }
function openModalManageMachine() { loadMachinesForManage(); document.getElementById('manNewName').value = ""; new bootstrap.Modal(document.getElementById('modalManageMachine')).show(); }
function loadMachinesForManage() { var kho = document.getElementById('manKho').value; var machines = [...new Set(DB.dm.filter(r => r[0] == kho).map(r => r[1]))]; var sel = document.getElementById('manOldName'); sel.innerHTML = ""; machines.forEach(m => { var opt = document.createElement('option'); opt.value = m; opt.text = m; sel.add(opt); }); }
async function submitRenameMachine() { var kho = document.getElementById('manKho').value; var oldName = document.getElementById('manOldName').value; var newName = document.getElementById('manNewName').value; if(!newName) return alert("Vui lòng nhập tên mới!"); document.getElementById('loading').style.display = 'flex'; await callAPI('renameMachine', {kho: kho, oldName: oldName, newName: newName}); bootstrap.Modal.getInstance(document.getElementById('modalManageMachine')).hide(); await loadInitialData(); document.getElementById('loading').style.display = 'none'; }
function openModalUser() { document.getElementById('admUserU').value = ""; document.getElementById('admUserU').readOnly = false; document.getElementById('admUserP').value = ""; document.getElementById('admUserN').value = ""; document.getElementById('admUserR').value = "User"; document.getElementById('admUserS').value = "Active"; new bootstrap.Modal(document.getElementById('modalAdminUser')).show(); }
function editUser(u) { var item = DB.users.find(x => x[0] == u); if (!item) return; document.getElementById('admUserU').value = item[0]; document.getElementById('admUserU').readOnly = true; document.getElementById('admUserP').value = item[1]; document.getElementById('admUserN').value = item[2]; document.getElementById('admUserR').value = item[3]; document.getElementById('admUserS').value = item[4]; new bootstrap.Modal(document.getElementById('modalAdminUser')).show(); }
async function saveAdminUser() { var data = { user: document.getElementById('admUserU').value, pass: document.getElementById('admUserP').value, name: document.getElementById('admUserN').value, role: document.getElementById('admUserR').value, status: document.getElementById('admUserS').value }; if(!data.user || !data.pass) return alert("Nhập đủ User/Pass!"); document.getElementById('loading').style.display = 'flex'; await callAPI('saveUser', {data: data}); bootstrap.Modal.getInstance(document.getElementById('modalAdminUser')).hide(); await loadInitialData(); document.getElementById('loading').style.display = 'none'; }
async function delUser(u) { if(!confirm("Xóa User [" + u + "]?")) return; document.getElementById('loading').style.display = 'flex'; await callAPI('deleteUser', {user: u}); await loadInitialData(); document.getElementById('loading').style.display = 'none'; }

// --- HÀM APP CŨ ---
function addImportRow(defaultDate) { if(!defaultDate) { var firstDate = document.querySelector('.row-date'); defaultDate = firstDate ? firstDate.value : new Date().toISOString().split('T')[0]; } var idx = rowCount++; var html = `<div class="import-row" id="row_${idx}"><i class="fas fa-times btn-remove-row" onclick="removeRow(${idx})"></i><div class="row g-3"><div class="col-md-3"><label class="form-label fw-bold small text-muted">Ngày nhập</label><input type="date" class="form-control row-date" id="date_${idx}" value="${defaultDate}" required></div><div class="col-md-9"><label class="form-label fw-bold small text-muted">Tên Hóa chất / Test</label><input class="form-control border-primary" list="listNhapHC" id="tenHC_${idx}" onchange="updateRowInput(${idx})" placeholder="Gõ để tìm kiếm..." required></div><input type="hidden" id="loai_${idx}"><input type="hidden" id="dv_${idx}"><div class="col-12" id="tenderArea_${idx}"></div><div class="col-12" id="inputArea_${idx}"></div></div></div>`; document.getElementById('importRowsContainer').insertAdjacentHTML('beforeend', html); }
function removeRow(idx) { var row = document.getElementById('row_' + idx); if(row) row.remove(); }
function updateRowInput(idx) {
  var tenHC = document.getElementById('tenHC_' + idx).value;
  var kho = document.getElementById('nhapKho').value;
  var may = document.getElementById('nhapMay').value;
  var item = DB.dm.find(r => r[0] == kho && r[1] == may && r[2] == tenHC);
  var area = document.getElementById('inputArea_' + idx);
  if(!item) { area.innerHTML = ''; document.getElementById('tenderArea_' + idx).innerHTML = ''; return; }
  var loai = item[3]; var dv = item[4];
  document.getElementById('loai_' + idx).value = loai;
  document.getElementById('dv_' + idx).value = dv;

  var html = (loai === 'R1R2') ? `<div class="row g-2"><div class="col-md-6 border-end"><div class="text-primary fw-bold small mb-1">Thành phần R1 / Chính</div><div class="row g-1"><div class="col-4"><input class="form-control form-control-sm" id="lotR1_${idx}" placeholder="Lot R1" required></div><div class="col-4"><input type="date" class="form-control form-control-sm" id="hsdR1_${idx}" required></div><div class="col-4"><input type="number" class="form-control form-control-sm" id="slR1_${idx}" placeholder="SL" required oninput="updateTenderQuotaHint(${idx})"></div></div></div><div class="col-md-6"><div class="text-danger fw-bold small mb-1">Thành phần R2</div><div class="row g-1"><div class="col-4"><input class="form-control form-control-sm" id="lotR2_${idx}" placeholder="Lot R2" required></div><div class="col-4"><input type="date" class="form-control form-control-sm" id="hsdR2_${idx}" required></div><div class="col-4"><input type="number" class="form-control form-control-sm" id="slR2_${idx}" placeholder="SL" required oninput="updateTenderQuotaHint(${idx})"></div></div></div><div class="col-12 text-end text-muted small fst-italic mt-1">Đơn vị: ${dv}</div></div>` : `<div class="row g-2"><div class="col-12"><div class="text-success fw-bold small mb-1">Chi tiết lô hạn</div><div class="row g-1"><div class="col-4"><input class="form-control form-control-sm" id="lotR1_${idx}" placeholder="Lot" required></div><div class="col-4"><input type="date" class="form-control form-control-sm" id="hsdR1_${idx}" required></div><div class="col-4"><input type="number" class="form-control form-control-sm" id="slR1_${idx}" placeholder="SL" required oninput="updateTenderQuotaHint(${idx})"></div></div></div><div class="col-12 text-end text-muted small fst-italic mt-1">Đơn vị: ${dv}</div></div>`;
  area.innerHTML = html;
  renderTenderSelector(idx, kho, may, tenHC);
}

function getTenderOptions(kho, may, ten) {
  return DB.tenders.filter(t => normalizeKeyPart(t.kho) === normalizeKeyPart(kho) && normalizeKeyPart(t.may) === normalizeKeyPart(may) && normalizeKeyPart(t.ten) === normalizeKeyPart(ten));
}

function renderTenderSelector(idx, kho, may, tenHC) {
  var box = document.getElementById('tenderArea_' + idx);
  var tenders = getTenderOptions(kho, may, tenHC).filter(t => t.status === 'Active' && t.remainingQuyDoi > 0);
  if(!box) return;
  if(tenders.length === 0) {
    box.innerHTML = `<div class="alert alert-warning shadow-sm d-flex align-items-center gap-2"><i class="fas fa-triangle-exclamation text-warning"></i><div><div class="fw-bold">Chưa có thầu Active hoặc đã hết số lượng</div><div class="small text-muted">Vui lòng liên hệ Admin để cấu hình thầu trước khi nhập kho.</div></div></div>`;
    return;
  }

  var options = tenders.map(t => {
    var remaining = Math.max(0, t.remainingQuyDoi || 0).toFixed(1);
    var ratioText = (t.r1PerBox || t.r2PerBox) ? ` | R1/hộp: ${t.r1PerBox || 1}, R2/hộp: ${t.r2PerBox || 1}` : '';
    return `<option value="${t.tenderKey}">${t.hang} | ${t.ncc} (Năm ${t.nam}) - Còn ${remaining} ${t.dvSd || ''}${ratioText}</option>`;
  }).join('');

  box.innerHTML = `<div class="card border-0 shadow-sm bg-white rounded p-3 mb-2">
      <div class="row g-2 align-items-end">
        <div class="col-md-4">
          <label class="small fw-bold text-muted">Chọn hãng/NCC (theo thầu)</label>
          <select class="form-select" id="selTender_${idx}" onchange="onTenderSelect(${idx})">${options}</select>
        </div>
        <div class="col-md-4">
          <label class="small fw-bold text-muted">Hãng sản xuất</label>
          <input class="form-control" id="hang_${idx}" readonly>
        </div>
        <div class="col-md-4">
          <label class="small fw-bold text-muted">Nhà cung cấp</label>
          <input class="form-control" id="ncc_${idx}" readonly>
        </div>
        <div class="col-md-3">
          <label class="small fw-bold text-muted">Năm thầu</label>
          <input class="form-control" id="namThau_${idx}" readonly>
        </div>
        <div class="col-md-3">
          <label class="small fw-bold text-muted">Đơn vị thầu</label>
          <input class="form-control" id="dvThau_${idx}" readonly>
        </div>
        <div class="col-md-3">
          <label class="small fw-bold text-muted">Hệ số quy đổi</label>
          <input class="form-control" id="heSo_${idx}" readonly>
        </div>
        <div class="col-md-3">
          <label class="small fw-bold text-muted">R1 / hộp</label>
          <input class="form-control" id="r1pb_${idx}" readonly>
        </div>
        <div class="col-md-3">
          <label class="small fw-bold text-muted">R2 / hộp</label>
          <input class="form-control" id="r2pb_${idx}" readonly>
        </div>
        <div class="col-md-3">
          <label class="small fw-bold text-muted">SL còn lại (quy đổi)</label>
          <div id="remainBadge_${idx}" class="badge bg-light text-dark w-100 p-2 border"></div>
        </div>
        <div class="col-12 mt-2" id="quotaHint_${idx}"></div>
      </div>
      <input type="hidden" id="tenderKey_${idx}">
    </div>`;

  onTenderSelect(idx);
}

function onTenderSelect(idx) {
  var sel = document.getElementById('selTender_' + idx);
  if(!sel) return;
  var key = sel.value;
  var tender = DB.tenders.find(t => t.tenderKey === key);
  document.getElementById('tenderKey_' + idx).value = key;
  document.getElementById('hang_' + idx).value = tender ? tender.hang : "";
  document.getElementById('ncc_' + idx).value = tender ? tender.ncc : "";
  document.getElementById('namThau_' + idx).value = tender ? tender.nam : "";
  document.getElementById('dvThau_' + idx).value = tender ? tender.dvThau : "";
  document.getElementById('heSo_' + idx).value = tender ? tender.heso : "";
  document.getElementById('r1pb_' + idx).value = tender ? (tender.r1PerBox || 1) : "";
  document.getElementById('r2pb_' + idx).value = tender ? (tender.r2PerBox || 1) : "";
  var remain = tender ? Math.max(0, tender.remainingQuyDoi || 0) : 0;
  var badge = document.getElementById('remainBadge_' + idx);
  if (badge) {
    var cls = remain <= (tender ? (tender.soLuongQuyDoi * 0.1) : 0) ? 'bg-danger text-white' : (remain <= (tender ? (tender.soLuongQuyDoi * 0.2) : 0) ? 'bg-warning text-dark' : 'bg-success');
    badge.className = 'badge ' + cls + ' w-100 p-2';
    badge.innerText = remain + ' ' + (tender && tender.dvSd ? tender.dvSd : '');
  }

  updateTenderQuotaHint(idx);
}

function calcTenderUsageClient(tender, loai, slR1, slR2) {
  var q1 = Number(slR1) || 0;
  var q2 = Number(slR2) || 0;
  var r1pb = Number(tender && tender.r1PerBox ? tender.r1PerBox : 1) || 1;
  var r2pb = Number(tender && tender.r2PerBox ? tender.r2PerBox : 1) || 1;
  var boxes = 0; var warn = "";

  if (loai === 'R1R2') {
    var boxR1 = r1pb > 0 ? (q1 / r1pb) : 0;
    var boxR2 = r2pb > 0 ? (q2 / r2pb) : 0;
    boxes = Math.max(boxR1, boxR2);
    if (boxR1 && boxR2) {
      var diff = Math.abs(boxR1 - boxR2) / Math.max(boxR1, boxR2);
      if (diff > 0.1) warn = "R1/R2 lệch, cân đối lại.";
    }
  } else {
    boxes = r1pb > 0 ? (q1 / r1pb) : q1;
  }

  var heso = Number(tender && tender.heso ? tender.heso : 0) || 0;
  return { useBoxes: boxes, useQD: boxes * heso, warn: warn };
}

function updateTenderQuotaHint(idx) {
  var hintBox = document.getElementById('quotaHint_' + idx);
  if(!hintBox) return;
  var tenderKeyEl = document.getElementById('tenderKey_' + idx);
  var tenderKey = tenderKeyEl ? tenderKeyEl.value : '';
  var tender = (DB.tenders || []).find(t => t.tenderKey === tenderKey);
  if(!tender) {
    hintBox.innerHTML = '<div class="text-muted small">Chọn thầu để xem hạn mức</div>';
    return;
  }

  var loai = document.getElementById('loai_' + idx) ? document.getElementById('loai_' + idx).value : 'Single';
  var sl1 = document.getElementById('slR1_' + idx) ? Number(document.getElementById('slR1_' + idx).value) || 0 : 0;
  var sl2El = document.getElementById('slR2_' + idx);
  var sl2 = sl2El ? Number(sl2El.value) || 0 : 0;

  var usage = calcTenderUsageClient(tender, loai, sl1, sl2);
  var remainBoxes = (Number(tender.soLuongThau) || 0) - (Number(tender.usedThau) || 0);
  var remainQD = (Number(tender.soLuongQuyDoi) || 0) - (Number(tender.usedQuyDoi) || 0);
  var afterBoxes = remainBoxes - usage.useBoxes;
  var afterQD = remainQD - usage.useQD;
  var cls = afterBoxes < 0 ? 'text-danger fw-bold' : (afterBoxes <= remainBoxes * 0.2 ? 'text-warning fw-bold' : 'text-success');
  var warnText = usage.warn ? ` <span class="text-danger">${usage.warn}</span>` : '';

  hintBox.innerHTML = `<div class="small ${cls}">Hạn mức còn: ${remainBoxes.toFixed(2)} ${tender.dvThau || 'ĐV thầu'} (~${remainQD.toFixed(1)} ${tender.dvSd || ''}). Sau nhập: ${Math.max(0, afterBoxes).toFixed(2)} ${tender.dvThau || ''} (~${Math.max(0, afterQD).toFixed(1)} ${tender.dvSd || ''}).${warnText}</div>`;
}

function onTabThauClick() {
  refreshTenderData();
}

async function refreshTenderData() {
  document.getElementById('loading').style.display = 'flex';
  const res = await callAPI('getTenderData');
  if(res && res.tenders) {
    DB.tenders = res.tenders;
    renderTenderView();
    if(DB.role === 'Admin') renderTenderAdmin();
  }
  document.getElementById('loading').style.display = 'none';
}

function renderTenderView() {
  var body = document.getElementById('tenderBody');
  if(!body) return;
  var addBtn = document.getElementById('btnAddTender');
  if(addBtn) addBtn.style.display = (DB.role === 'Admin') ? 'inline-flex' : 'none';

  if(!DB.tenders || DB.tenders.length === 0) {
    body.innerHTML = `<tr><td colspan="15" class="text-center text-muted py-4"><i class="fas fa-circle-info me-2"></i>Chưa có dữ liệu thầu</td></tr>`;
    return;
  }

  var rows = (DB.tenders || []).map(t => {
    var badgeCls = t.status === 'Active' ? 'badge bg-success' : 'badge bg-secondary';
    var remain = Math.max(0, t.remainingQuyDoi || 0).toFixed(1);
    var used = Math.max(0, t.usedQuyDoi || 0).toFixed(1);
    var actions = (DB.role === 'Admin') ? `<button class="btn btn-outline-primary btn-sm me-1" onclick="openTenderModal(${t.rowIndex})"><i class="fas fa-edit"></i></button><button class="btn btn-outline-${t.status==='Active'?'danger':'success'} btn-sm" onclick="toggleTender(${t.rowIndex}, '${t.status==='Active'?'Disabled':'Active'}')"><i class="fas ${t.status==='Active'?'fa-ban':'fa-rotate-right'}"></i></button>` : '';
    var warnCls = t.remainingQuyDoi <= (t.soLuongQuyDoi * 0.1) ? 'text-danger fw-bold' : (t.remainingQuyDoi <= (t.soLuongQuyDoi * 0.2) ? 'text-warning fw-bold' : '');
    return `<tr>
      <td>${t.kho}</td>
      <td>${t.may}</td>
      <td class="fw-bold">${t.ten}</td>
      <td>${t.hang}</td>
      <td>${t.ncc}</td>
      <td>${t.nam}</td>
      <td>${t.dvThau}</td>
      <td>${t.heso}</td>
      <td>${t.r1PerBox || 1}</td>
      <td>${t.r2PerBox || 1}</td>
      <td>${t.soLuongThau}</td>
      <td class="text-muted">${used}</td>
      <td class="${warnCls}">${remain}</td>
      <td><span class="${badgeCls}">${t.status}</span></td>
      <td class="text-end">${actions}</td>
    </tr>`;
  }).join('');

  body.innerHTML = rows;
}

function renderTenderAdmin() {
  var selKho = document.getElementById('tndKho');
  if(!selKho) return;
  selKho.innerHTML = '';
  var khoList = [...new Set(DB.dm.map(r => r[0]))];
  khoList.forEach(k => { var opt = document.createElement('option'); opt.value = k; opt.text = k; selKho.add(opt); });

  var machineDl = document.getElementById('listAdmMachines');
  var chemDl = document.getElementById('listAdmChem');
  if(machineDl) {
    machineDl.innerHTML = '';
    var machines = [...new Set(DB.dm.map(r => r[1]))];
    machines.forEach(m => machineDl.innerHTML += `<option value="${m}">`);
  }
  if(chemDl) {
    chemDl.innerHTML = '';
    var chems = [...new Set(DB.dm.map(r => r[2]))];
    chems.forEach(c => chemDl.innerHTML += `<option value="${c}">`);
  }
}

function openTenderModal(rowIndex) {
  renderTenderAdmin();
  var isEdit = !!rowIndex;
  document.getElementById('titleTender').innerText = isEdit ? 'SỬA THẦU' : 'THÊM THẦU';
  document.getElementById('tenderRowIndex').value = isEdit ? rowIndex : '';

  var data = isEdit ? (DB.tenders.find(t => t.rowIndex === rowIndex) || {}) : {};
  document.getElementById('tndKho').value = data.kho || '';
  document.getElementById('tndMay').value = data.may || '';
  document.getElementById('tndTen').value = data.ten || '';
  document.getElementById('tndHang').value = data.hang || '';
  document.getElementById('tndNcc').value = data.ncc || '';
  document.getElementById('tndNam').value = data.nam || '';
  document.getElementById('tndDvThau').value = data.dvThau || '';
  document.getElementById('tndDvSd').value = data.dvSd || '';
  document.getElementById('tndHeSo').value = data.heso || '';
  document.getElementById('tndR1PB').value = data.r1PerBox || 1;
  document.getElementById('tndR2PB').value = data.r2PerBox || 1;
  document.getElementById('tndSoLuong').value = data.soLuongThau || '';
  document.getElementById('tndNote').value = data.note || '';
  document.getElementById('tndStatus').value = data.status || 'Active';

  new bootstrap.Modal(document.getElementById('modalTender')).show();
}

async function submitTenderForm() {
  var data = {
    rowIndex: document.getElementById('tenderRowIndex').value ? Number(document.getElementById('tenderRowIndex').value) : null,
    kho: document.getElementById('tndKho').value,
    may: document.getElementById('tndMay').value,
    ten: document.getElementById('tndTen').value,
    hang: document.getElementById('tndHang').value,
    ncc: document.getElementById('tndNcc').value,
    nam: document.getElementById('tndNam').value,
    dvThau: document.getElementById('tndDvThau').value,
    dvSd: document.getElementById('tndDvSd').value,
    heso: document.getElementById('tndHeSo').value,
    soLuongThau: document.getElementById('tndSoLuong').value,
    r1PerBox: document.getElementById('tndR1PB').value,
    r2PerBox: document.getElementById('tndR2PB').value,
    note: document.getElementById('tndNote').value,
    status: document.getElementById('tndStatus').value
  };

  if(!data.kho || !data.ten || !data.hang || !data.ncc) return alert('Nhập đủ Kho/Tên/Hãng/NCC');
  if(!data.heso || Number(data.heso) <= 0) return alert('Hệ số quy đổi phải > 0');
  if(!data.soLuongThau || Number(data.soLuongThau) <= 0) return alert('Số lượng thầu phải > 0');
  if(!data.r1PerBox || Number(data.r1PerBox) <= 0) return alert('R1/hộp phải > 0');
  if(!data.r2PerBox || Number(data.r2PerBox) <= 0) return alert('R2/hộp phải > 0');

  document.getElementById('loading').style.display = 'flex';
  var res = await callAPI('saveTender', {data: data});
  document.getElementById('loading').style.display = 'none';
  if(!res || !res.success) return alert(res?.msg || 'Lưu thầu thất bại');
  bootstrap.Modal.getInstance(document.getElementById('modalTender')).hide();
  await refreshTenderData();
  alert(res.msg || 'Đã lưu thầu');
}

async function toggleTender(rowIndex, status) {
  document.getElementById('loading').style.display = 'flex';
  var res = await callAPI('toggleTender', {rowIndex: rowIndex, status: status});
  document.getElementById('loading').style.display = 'none';
  if(!res || !res.success) return alert(res?.msg || 'Không cập nhật được trạng thái');
  await refreshTenderData();
}
async function submitNhapBatch(e) {
  e.preventDefault();
  var kho = document.getElementById('nhapKho').value;
  var may = document.getElementById('nhapMay').value;
  var rows = document.querySelectorAll('.import-row');
  if(rows.length === 0) return alert("Chưa có dòng dữ liệu nào!");
  var dataList = [];
  var hasError = false;

  rows.forEach(row => {
    var idx = row.id.split('_')[1];
    var tenHC = document.getElementById('tenHC_' + idx).value;
    var exists = DB.dm.some(r => r[0] == kho && r[1] == may && r[2] == tenHC);
    if (!exists) { alert("Dòng hóa chất '" + tenHC + "' không hợp lệ hoặc sai Kho/Máy!"); hasError = true; return; }

    var tenderKeyEl = document.getElementById('tenderKey_' + idx);
    var tenderKey = tenderKeyEl ? tenderKeyEl.value : "";
    if (!tenderKey) { alert("Vui lòng chọn hãng/NCC thầu cho " + tenHC); hasError = true; return; }

    var itemData = {
      nguoi: DB.currentUser,
      kho: kho,
      may: may,
      tenHC: tenHC,
      ngayNhap: document.getElementById('date_' + idx).value,
      loaiNhap: document.getElementById('loai_' + idx).value,
      donVi: document.getElementById('dv_' + idx).value,
      lotR1: document.getElementById('lotR1_' + idx).value,
      hsdR1: document.getElementById('hsdR1_' + idx).value,
      slR1: document.getElementById('slR1_' + idx).value,
      lotR2: document.getElementById('lotR2_' + idx) ? document.getElementById('lotR2_' + idx).value : "",
      hsdR2: document.getElementById('hsdR2_' + idx) ? document.getElementById('hsdR2_' + idx).value : "",
      slR2: document.getElementById('slR2_' + idx) ? document.getElementById('slR2_' + idx).value : "",
      hangSX: document.getElementById('hang_' + idx).value,
      nhaCC: document.getElementById('ncc_' + idx).value,
      namThau: document.getElementById('namThau_' + idx).value,
      dvThau: document.getElementById('dvThau_' + idx).value,
      heSoQD: document.getElementById('heSo_' + idx).value,
      tenderKey: tenderKey
    };

    dataList.push(itemData);
  });

  if(hasError) return;
  document.getElementById('loading').style.display = 'flex';
  var res = await callAPI('processImport', {dataList: dataList});
  document.getElementById('loading').style.display = 'none';
  if(!res || !res.success) return alert(res?.msg || "Nhập kho thất bại!");
  alert(res.msg || "Nhập kho thành công!");
  await refreshTenderData();
  setTodayDate();
  isReportDirty = true;
  isStockLoaded = false;
}
async function refreshData() { document.getElementById('loading').style.display = 'flex'; await loadInitialData(); alert("Đã cập nhật!"); filterMachine('nhap'); document.getElementById('loading').style.display = 'none'; }
function filterMachine(type) { var kho = document.getElementById(type + 'Kho').value; var selMay = document.getElementById(type + 'May'); if(type === 'nhap') { document.getElementById('importRowsContainer').innerHTML = ""; rowCount = 0; addImportRow(); } else { document.getElementById('xuatTenHC').value = ""; } var listId = (type === 'nhap' ? 'listNhapHC' : 'listXuatHC'); document.getElementById(listId).innerHTML = ""; selMay.innerHTML = '<option value="">-- Chọn Máy --</option>'; if(!kho) return; var machines = [...new Set(DB.dm.filter(r => r[0] == kho).map(r => r[1]))]; machines.forEach(m => { var opt = document.createElement('option'); opt.value = m; opt.text = m===""?"Không dùng máy":m; selMay.add(opt); }); if(kho === 'Vi sinh' && (machines.length===0 || machines[0]==="")) filterChem(type); }
function filterChem(type) { var kho = document.getElementById(type + 'Kho').value; var may = document.getElementById(type + 'May').value; if(type === 'nhap') { document.getElementById('importRowsContainer').innerHTML = ""; rowCount = 0; addImportRow(); } else { document.getElementById('xuatTenHC').value = ""; } var listId = (type === 'nhap' ? 'listNhapHC' : 'listXuatHC'); var datalist = document.getElementById(listId); datalist.innerHTML = ""; var chems = DB.dm.filter(r => r[0] == kho && r[1] == may); chems.forEach(c => { var opt = document.createElement('option'); opt.value = c[2]; datalist.appendChild(opt); }); }
function checkBackdate() { var d = new Date(document.getElementById('xuatNgay').value).setHours(0,0,0,0); var t = new Date().setHours(0,0,0,0); var r = document.getElementById('xuatLyDo'); if(d < t) { r.value="Quên không xuất kho"; r.readOnly=true; } else { r.value=""; r.readOnly=false; } }

async function onTabXuatClick() {
    if (!isStockLoaded || Object.keys(DB.stock).length === 0) {
        await forceRefreshStock(false);
    }
}

async function forceRefreshStock(showAlert = true) {
    document.getElementById('loadingStockIndicator').style.display = 'block';
    if(showAlert) document.getElementById('loading').style.display = 'flex';
    var data = await callAPI('getReport', {from: "", to: ""});
    if (data && data.error) {
      document.getElementById('loadingStockIndicator').style.display = 'none';
      if(showAlert) document.getElementById('loading').style.display = 'none';
      return alert(data.error);
    }
    DB.stock = data.stock || {};
    isStockLoaded = true;
    document.getElementById('loadingStockIndicator').style.display = 'none';
    if(showAlert) {
        document.getElementById('loading').style.display = 'none';
        alert("Đã cập nhật dữ liệu tồn kho mới nhất!");
    }
    if(document.getElementById('xuatTenHC').value) {
        showStockDetails();
    }
}

function showStockDetails() {
   var area = document.getElementById('exportContainer');
   var tenHC = document.getElementById('xuatTenHC').value;
   if(!tenHC) {
       area.innerHTML = '<div class="text-center text-muted p-5 bg-light border rounded"><i class="fas fa-box-open fa-2x mb-2 text-secondary"></i><br>Chọn hóa chất để hiển thị danh sách Lot</div>';
       return;
   }
  var key = document.getElementById('xuatKho').value + "|" + document.getElementById('xuatMay').value + "|" + tenHC;
  var item = DB.stock[key] || findStockByKey(key);
   if(!item) {
       area.innerHTML = '<div class="text-danger text-center p-3 bg-light border rounded">Không có dữ liệu tồn kho (hoặc chưa nhập kho)</div>';
       return;
   }
   var htmlR1 = generateLotList('R1', item.R1);
   var htmlR2 = generateLotList('R2', item.R2);
   var colClass = htmlR2 ? "col-md-6" : "col-12";
   var displayR2 = htmlR2 ? "block" : "none";
   area.innerHTML = `<div class="row g-3"><div class="${colClass}"><div class="p-3 border rounded bg-white h-100 shadow-sm"><div class="fw-bold text-primary mb-3 border-bottom pb-2">Thành phần R1 / Chính</div>${htmlR1 || '<div class="text-muted small fst-italic">Hết hàng</div>'}<div id="fefo_alert_R1" class="alert-fefo"><i class="fas fa-exclamation-triangle me-1"></i> Cảnh báo: Có Lot khác hết hạn sớm hơn!</div></div></div><div class="${colClass}" style="display:${displayR2}"><div class="p-3 border rounded bg-white h-100 shadow-sm"><div class="fw-bold text-danger mb-3 border-bottom pb-2">Thành phần R2</div>${htmlR2 || '<div class="text-muted small fst-italic">Hết hàng</div>'}<div id="fefo_alert_R2" class="alert-fefo"><i class="fas fa-exclamation-triangle me-1"></i> Cảnh báo: Có Lot khác hết hạn sớm hơn!</div></div></div></div>`;
}

function normalizeKeyPart(value) {
  return String(value || '').trim().replace(/\s+/g, ' ').toLowerCase();
}

function normalizeStockKey(key) {
  var parts = String(key || '').split('|');
  return [normalizeKeyPart(parts[0]), normalizeKeyPart(parts[1]), normalizeKeyPart(parts[2])].join('|');
}

function findStockByKey(rawKey) {
  var target = normalizeStockKey(rawKey);
  for (var k in DB.stock) {
    if (normalizeStockKey(k) === target) return DB.stock[k];
  }
  return null;
}

async function onTabBaocaoClick() {
    if (isReportDirty || DB.history.length === 0) {
        await loadReport(false);
    } else {
        console.log("Dữ liệu báo cáo chưa thay đổi, dùng cache.");
    }
}

async function submitXuat(e) { e.preventDefault(); var radioR1 = document.querySelector('input[name="radio_R1"]:checked'); var lotR1 = "", slR1 = ""; if (radioR1) { lotR1 = radioR1.value; slR1 = document.getElementById(`qty_R1_${lotR1}`).value; } var radioR2 = document.querySelector('input[name="radio_R2"]:checked'); var lotR2 = "", slR2 = ""; if (radioR2) { lotR2 = radioR2.value; slR2 = document.getElementById(`qty_R2_${lotR2}`).value; } if (!lotR1 && !lotR2) return alert("Chưa chọn Lot để xuất!"); if ((lotR1 && !slR1) || (lotR2 && !slR2)) return alert("Chưa nhập số lượng xuất!"); document.getElementById('loading').style.display = 'flex'; var data = { ngayXuat: document.getElementById('xuatNgay').value, nguoi: DB.currentUser, kho: document.getElementById('xuatKho').value, may: document.getElementById('xuatMay').value, tenHC: document.getElementById('xuatTenHC').value, lotR1: lotR1, slR1: slR1, lotR2: lotR2, slR2: slR2, loaiChiTiet: (lotR1 && lotR2) ? "Cả 2" : (lotR1 ? "R1" : "R2"), lyDo: document.getElementById('xuatLyDo').value }; var res = await callAPI('processExport', {data: data}); document.getElementById('loading').style.display = 'none'; if(!res || !res.success) return alert(res?.msg || "Xuất kho thất bại!"); alert(res.msg || "Xuất kho thành công!");
    await forceRefreshStock(false);
    isReportDirty = true;
}

function openChangePassModal() { document.getElementById('cpUser').value = DB.currentID; document.getElementById('cpOldPass').value = ""; document.getElementById('cpNewPass').value = ""; document.getElementById('cpMsg').innerText = ""; var myModal = new bootstrap.Modal(document.getElementById('modalChangePass')); myModal.show(); }
async function doChangePass() { var u = document.getElementById('cpUser').value; var oldP = document.getElementById('cpOldPass').value; var newP = document.getElementById('cpNewPass').value; if(!oldP || !newP) return document.getElementById('cpMsg').innerText = "Vui lòng nhập đủ thông tin!"; document.getElementById('loading').style.display = 'flex'; var res = await callAPI('changePassword', {user: u, oldPass: oldP, newPass: newP}); document.getElementById('loading').style.display = 'none'; if(res && res.success) { alert("Đổi mật khẩu thành công! Vui lòng đăng nhập lại."); doLogout(); } else { document.getElementById('cpMsg').innerText = (res && res.msg) ? res.msg : "Đổi mật khẩu thất bại!"; } }
function filterTable(tableId, colIndex, query) { var filter = query.toUpperCase(); var rows = document.getElementById(tableId).getElementsByTagName("tbody")[0].getElementsByTagName("tr"); for (var i = 0; i < rows.length; i++) { var td = rows[i].getElementsByTagName("td")[colIndex]; if (td) rows[i].style.display = (td.textContent || td.innerText).toUpperCase().indexOf(filter) > -1 ? "" : "none"; } }

function filterTonKho() {
    var k = document.getElementById("f_kho").value.toUpperCase(); var t = document.getElementById("f_ten").value.toUpperCase();
    var c1 = document.getElementById("f_color_r1").value; var t1 = document.getElementById("f_text_r1").value.toUpperCase();
    var c2 = document.getElementById("f_color_r2").value; var t2 = document.getElementById("f_text_r2").value.toUpperCase();
    var rows = document.getElementById("tableTonKho").getElementsByTagName("tbody")[0].getElementsByTagName("tr");
    for(var i=0; i<rows.length; i++){
        var td0 = rows[i].cells[0].textContent.toUpperCase(); var td1 = rows[i].cells[1].textContent.toUpperCase();
        var cell1 = rows[i].cells[2]; var cell2 = rows[i].cells[3];
        var match0 = td0.indexOf(k) > -1; var match1 = td1.indexOf(t) > -1;
        var matchC1 = (c1==="") || cell1.classList.contains(c1); var matchT1 = cell1.textContent.toUpperCase().indexOf(t1) > -1;
        var matchC2 = (c2==="") || cell2.classList.contains(c2); var matchT2 = cell2.textContent.toUpperCase().indexOf(t2) > -1;
        if(match0 && match1 && matchC1 && matchT1 && matchC2 && matchT2) rows[i].style.display = ""; else rows[i].style.display = "none";
    }
}

function generateLotList(type, listObj) { var arr = []; for(var k in listObj) if(listObj[k].sl > 0) arr.push({lot: k, ...listObj[k]}); arr.sort((a,b) => a.rawHsd - b.rawHsd); var bestDate = arr.length > 0 ? arr[0].rawHsd : 0; return arr.map(i => { var isBest = i.rawHsd === bestDate; var cls = isBest ? "lot-best" : "lot-warn"; var icon = isBest ? "✅" : "⚠️"; var inputId = `qty_${type}_${i.lot}`; return `<div class="d-flex align-items-center justify-content-between lot-item ${cls}"><div class="flex-grow-1" onclick="document.getElementById('rad_${inputId}').checked=true; document.getElementById('${inputId}').focus(); checkFefo('${type}', ${i.rawHsd}, ${bestDate}, ${arr.length})"><input class="form-check-input me-2" type="radio" name="radio_${type}" id="rad_${inputId}" value="${i.lot}" onchange="checkFefo('${type}', ${i.rawHsd}, ${bestDate}, ${arr.length})"><label class="small cursor-pointer" for="rad_${inputId}">${icon} Lot: <b>${i.lot}</b><br><span class="text-muted ms-4">HSD: ${i.hsd} | Tồn: ${i.sl}</span></label></div><div><input type="number" id="${inputId}" class="form-control form-control-sm qty-input" placeholder="SL" min="1" max="${i.sl}" onfocus="document.getElementById('rad_${inputId}').checked=true; checkFefo('${type}', ${i.rawHsd}, ${bestDate}, ${arr.length})"></div></div>`; }).join(''); }
function checkFefo(type, cur, best, count) { var box = document.getElementById('fefo_alert_' + type); if(box) box.style.display = (count > 1 && cur > best) ? 'block' : 'none'; }
function applyDatePreset(val) { if (val === "50") { document.getElementById('histFrom').value = ""; document.getElementById('histTo').value = ""; } else { var days = parseInt(val); var to = new Date(); var from = new Date(); from.setDate(to.getDate() - days); document.getElementById('histTo').valueAsDate = to; document.getElementById('histFrom').valueAsDate = from; } loadReport(true); }

async function loadReport(forceReload) {
    var from = "", to = "";

    if (!forceReload && !isReportDirty && DB.history.length > 0) {
        return;
    }

    if (document.getElementById('histFrom').value && document.getElementById('histTo').value) {
        from = document.getElementById('histFrom').value;
        to = document.getElementById('histTo').value;
    }

    if (document.getElementById('histPreset').value === "50" && !from) { from = ""; to = ""; }

    document.getElementById('loading').style.display = 'flex';
    var data = await callAPI('getReport', {from: from, to: to});
    if (data && data.error) {
      document.getElementById('loading').style.display = 'none';
      return alert(data.error);
    }
    DB.stock = data.stock || {};
    DB.history = data.history || [];

    renderReport();
    renderHistory();

    isReportDirty = false;
    document.getElementById('loading').style.display = 'none';
}

function exportExcel(tableId, name) { var wb = XLSX.utils.table_to_book(document.getElementById(tableId), {sheet: "Sheet1"}); XLSX.writeFile(wb, name + "_" + new Date().toISOString().slice(0,10) + ".xlsx"); }

function exportPDF(elementId, name) {
  var element = document.getElementById(elementId);
  var originalOverflow = element.style.overflow; var originalMaxHeight = element.style.maxHeight; element.style.overflow = 'visible'; element.style.maxHeight = 'none';
  var clone = element.cloneNode(true);
  var thead = clone.querySelector('thead');
  if(thead) { thead.classList.remove('sticky-top'); thead.style.position = 'static'; }
  var inputs = clone.querySelectorAll('input, select, .no-print');
  inputs.forEach(i => i.remove());

  var tempContainer = document.createElement('div');
  tempContainer.style.position = 'absolute'; tempContainer.style.top = '-9999px'; tempContainer.style.width = '1300px';
  tempContainer.appendChild(clone);
  document.body.appendChild(tempContainer);

  var opt = { margin: 10, filename: name + "_" + new Date().toISOString().slice(0,10) + '.pdf', image: { type: 'jpeg', quality: 0.98 }, html2canvas: { scale: 2 }, jsPDF: { unit: 'mm', format: 'a4', orientation: 'landscape' } };

  html2pdf().set(opt).from(clone).save().then(function(){
      document.body.removeChild(tempContainer);
      element.style.overflow = originalOverflow; element.style.maxHeight = originalMaxHeight;
  });
}

function renderReport() {
  var html = "";
  var now = new Date().getTime();

  function safeParse(v) {
    var n = parseFloat(v);
    return isNaN(n) ? 0 : n;
  }

  function normalizeDmKey(kho, may, ten) {
    return [normalizeKeyPart(kho), normalizeKeyPart(may), normalizeKeyPart(ten)].join('|');
  }

  var dmByKey = {};
  DB.dm.forEach(function(dm) {
    dmByKey[normalizeDmKey(dm[0], dm[1], dm[2])] = dm;
  });

  var stockKeys = Object.keys(DB.stock || {});
  stockKeys.forEach(function(rawKey) {
    var item = DB.stock[rawKey] || { R1: {}, R2: {} };
    var parts = String(rawKey).split('|');
    var kho = parts[0] || '';
    var may = parts[1] || '';
    var ten = parts[2] || '';
    var dm = dmByKey[normalizeDmKey(kho, may, ten)] || null;

    var cfg = {
      minR: dm ? safeParse(dm[5]) : 0,
      minY: dm ? safeParse(dm[6]) : 0,
      dayR: dm ? safeParse(dm[7]) : 0
    };
    var isTwoPart = dm ? (dm[3] === 'R1R2') : (item.R2 && Object.keys(item.R2).length > 0);

    function buildCol(listObj) {
      var total = 0;
      var details = "";
      var colorClass = "";
      listObj = listObj || {};
      for (var lot in listObj) {
        var s = listObj[lot] || {};
        var qty = parseFloat(s.sl);
        if (!isNaN(qty) && qty > 0) {
          total += qty;
          var rawHsd = parseFloat(s.rawHsd) || 0;
          var daysLeft = rawHsd ? (rawHsd - now) / (1000 * 60 * 60 * 24) : 0;
          var dateStyle = (cfg.dayR > 0 && rawHsd && daysLeft <= cfg.dayR) ? "lot-expired" : "";
          var hsdText = s.hsd || "";
          var dayText = rawHsd ? ` (${Math.ceil(daysLeft)} ngày)` : "";
          details += `<div class="lot-detail">Lot: <b>${lot}</b> (SL: <b>${qty}</b>) | HSD: <span class="${dateStyle}">${hsdText}</span>${dayText}</div>`;
        }
      }
      if (cfg.minR > 0 && total <= cfg.minR) colorClass = "cell-critical";
      else if (cfg.minY > 0 && total <= cfg.minY) colorClass = "cell-warning";
      else if (total > 0) colorClass = "cell-ok";
      return { total: total, html: `<div><span class="total-badge">Tổng: ${total}</span></div>${details}`, cls: colorClass };
    }

    var c1 = buildCol(item.R1);
    var c2 = isTwoPart ? buildCol(item.R2) : { total: 0, html: '', cls: 'bg-light' };

    if (c1.total <= 0 && (!isTwoPart || c2.total <= 0)) return;

    html += `<tr><td><small class="text-muted">${kho}</small><br><b>${may}</b></td><td class="fw-bold text-primary">${ten}</td><td class="${c1.cls}">${c1.html}</td><td class="${c2.cls}">${c2.html}</td></tr>`;
  });

  document.getElementById('reportBody').innerHTML = html;
}

var sortDir = -1;
var sortCol = 'sortTime';

function sortHistory(col) {
    if(sortCol === col) sortDir *= -1;
    else { sortCol = col; sortDir = 1; }

    document.querySelectorAll('.sort-icon').forEach(i => i.classList.remove('active'));
    var icon = document.getElementById('icon_' + col);
    if(icon) {
        icon.classList.add('active');
        icon.className = sortDir === 1 ? 'fas fa-sort-up sort-icon active' : 'fas fa-sort-down sort-icon active';
    }

    DB.history.sort((a,b) => {
        if(a[col] < b[col]) return -1 * sortDir;
        if(a[col] > b[col]) return 1 * sortDir;
        return 0;
    });
    renderHistory();
}

function renderHistory() {
    var html = DB.history.map(h => {
        var badge = (h.action === 'NHẬP' || h.type === 'NHẬP') ? '<span class="badge bg-primary">NHẬP</span>' : '<span class="badge bg-danger">XUẤT</span>';
    var ioTypeRaw = (h.ioType || '').toUpperCase();
    var ioTypeBadge = '';
    if (ioTypeRaw.indexOf('R1,R2') > -1 || ioTypeRaw.indexOf('CẢ') > -1 || ioTypeRaw.indexOf('R1R2') > -1) ioTypeBadge = '<span class="badge io-badge io-both">R1+R2</span>';
    else if (ioTypeRaw === 'R2') ioTypeBadge = '<span class="badge io-badge io-r2">R2</span>';
    else if (ioTypeRaw === 'R1') ioTypeBadge = '<span class="badge io-badge io-r1">R1</span>';

        return `<tr>
            <td>${h.timestampStr}</td>
            <td>${h.dateDocStr}</td>
            <td>${h.user}</td>
      <td class="text-center">${badge}${ioTypeBadge ? '<br>' + ioTypeBadge : ''}</td>
            <td>${h.kho}<br>${h.may}<br><b>${h.ten}</b></td>
            <td class="small" style="line-height:1.6">${h.detailHtml || h.detail}</td>
            <td>${h.reason}</td>
        </tr>`;
    }).join('');
    document.getElementById('historyBody').innerHTML = html;
}

function exportHistoryExcel() {
    var headers = ["THỜI GIAN", "NGÀY NHẬP/XUẤT", "NGƯỜI THỰC HIỆN", "THAO TÁC", "KHO", "MÁY XÉT NGHIỆM", "TÊN HÓA CHẤT", "CHI TIẾT", "LÝ DO"];
    var data = [headers];
    DB.history.forEach(item => {
        var kho = item.kho || ""; var may = item.may || ""; var ten = item.ten || "";
        var rawDetail = item.detailHtml || "";
        var detailText = rawDetail.replace(/<br\s*\/?>/gi, "\n").replace(/<[^>]+>/g, "");
        var actionText = (item.action === 'NHẬP' || item.type === 'NHẬP') ? "NHẬP" : "XUẤT";
        if (item.ioType) actionText += " [" + item.ioType + "]";
        var row = [item.timestampStr, item.dateDocStr, item.user, actionText, kho, may, ten, detailText, item.reason];
        data.push(row);
    });
    var wb = XLSX.utils.book_new();
    var ws = XLSX.utils.aoa_to_sheet(data);
    ws['!cols'] = [{wch:20}, {wch:15}, {wch:20}, {wch:10}, {wch:15}, {wch:15}, {wch:25}, {wch:50}, {wch:20}];
    XLSX.utils.book_append_sheet(wb, ws, "LichSuHoatDong");
    XLSX.writeFile(wb, "Lich_Su_Hoat_Dong_" + new Date().toISOString().slice(0,10) + ".xlsx");
}

function createResizableTable(tableId) {
  const table = document.getElementById(tableId);
  if (!table) return;
  const cols = table.querySelectorAll('th');

  cols.forEach((col) => {
    const resizer = document.createElement('div');
    resizer.classList.add('resizer');
    col.appendChild(resizer);

    let x = 0; let w = 0;

    const mouseDownHandler = function(e) {
      x = e.clientX;
      const styles = window.getComputedStyle(col);
      w = parseInt(styles.width, 10);
      document.addEventListener('mousemove', mouseMoveHandler);
      document.addEventListener('mouseup', mouseUpHandler);
      resizer.classList.add('resizing');
    };

    const mouseMoveHandler = function(e) {
      const dx = e.clientX - x;
      col.style.width = `${w + dx}px`;
    };

    const mouseUpHandler = function() {
      document.removeEventListener('mousemove', mouseMoveHandler);
      document.removeEventListener('mouseup', mouseUpHandler);
      resizer.classList.remove('resizing');
    };

    resizer.addEventListener('mousedown', mouseDownHandler);
  });
}

window.onload = function() {
    createResizableTable('tableTonKho');
    createResizableTable('tableLichSu');
};
