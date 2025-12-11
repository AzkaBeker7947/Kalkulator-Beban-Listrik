// Util
const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => document.querySelectorAll(sel);
const toRp = (n) => {
if (!isFinite(n)) return "Rp 0";
return "Rp " + Number(n).toLocaleString("id-ID", {maximumFractionDigits:0});
};

// DOM
const nama = $('#nama'), watt = $('#watt'), jam = $('#jam'), jumlah = $('#jumlah');
const btnTambah = $('#btn-tambah'), btnReset = $('#btn-reset');
const tblBody = document.querySelector('#tbl tbody');
const selectAll = $('#selectAll');
const btnHapusSelected = $('#btn-hapus-selected');
const btnCalc = $('#btn-calc');
const hargaInput = $('#harga');
const daysPerMonthInput = $('#daysPerMonth');

const totalWhEl = $('#totalWh'), totalKwhEl = $('#totalkWh');
const costDailyEl = $('#costDaily'), costMonthlyEl = $('#costMonthly'), costYearlyEl = $('#costYearly');
const btnExport = $('#btn-export');

// state
let items = [];

// keyboard navigation
nama.addEventListener('keydown', (e) => { if (e.key === 'Enter') watt.focus(); });
watt.addEventListener('keydown', (e) => { if (e.key === 'Enter') jam.focus(); });
jam.addEventListener('keydown', (e) => { if (e.key === 'Enter') jumlah.focus(); });
jumlah.addEventListener('keydown', (e) => { if (e.key === 'Enter') addItem(); });

btnTambah.addEventListener('click', addItem);
btnReset.addEventListener('click', (e) => {
nama.value=''; watt.value=''; jam.value=''; jumlah.value='1'; nama.focus();
});

selectAll.addEventListener('change', (e) => {
const checked = e.target.checked;
document.querySelectorAll('.row-select').forEach(cb => cb.checked = checked);
});

btnHapusSelected.addEventListener('click', () => {
const checkedBoxes = Array.from(document.querySelectorAll('.row-select:checked'));
if (checkedBoxes.length === 0) return alert('Pilih baris yang ingin dihapus.');
if (!confirm('Hapus ' + checkedBoxes.length + ' item terpilih?')) return;
const ids = checkedBoxes.map(cb => cb.dataset.id);
items = items.filter(it => !ids.includes(String(it.id)));
renderTable();
// tidak otomatis menghitung
});

btnCalc.addEventListener('click', () => {
calcTotals();
window.scrollTo({top:0, behavior:'smooth'});
});

btnExport.addEventListener('click', exportExcel);

// helpers
function uid(){
return Date.now().toString(36) + Math.random().toString(36).slice(2,8);
}

function addItem(){
const n = nama.value.trim();
const w = parseFloat(watt.value);
const h = parseFloat(jam.value);
const q = parseInt(jumlah.value) || 1;

if (!n) return alert('Masukkan nama barang.');
if (!isFinite(w) || w < 0) return alert('Daya (Watt) tidak valid.');
if (!isFinite(h) || h < 0) return alert('Jam pemakaian tidak valid.');

const wh = w * h * q; // Wh per hari
const it = { id: uid(), nama: n, watt: Number(w), jam: Number(h), jumlah: Number(q), wh };
items.push(it);

nama.value=''; watt.value=''; jam.value=''; jumlah.value='1';
nama.focus();

renderTable();
// tidak otomatis menghitung
}

function renderTable(){
tblBody.innerHTML = '';
items.forEach((it, idx) => {
  const tr = document.createElement('tr');

  tr.innerHTML = `
    <td><input class="row-select" data-id="${it.id}" type="checkbox"></td>
    <td class="col-center">${idx+1}</td>
    <td class="editable" data-field="nama" data-id="${it.id}">${escapeHtml(it.nama)}</td>
    <td class="col-center editable" data-field="watt" data-id="${it.id}">${it.watt}</td>
    <td class="col-center editable" data-field="jam" data-id="${it.id}">${it.jam}</td>
    <td class="col-center editable" data-field="jumlah" data-id="${it.id}">${it.jumlah}</td>
    <td class="col-num">${(it.wh).toLocaleString('id-ID', {maximumFractionDigits:2})} Wh</td>
    <td class="col-center">
      <button class="ghost small btn-edit" data-id="${it.id}" title="Edit"><i class="fa fa-pen"></i></button>
      <button class="ghost small btn-delete" data-id="${it.id}" title="Delete"><i class="fa fa-trash"></i></button>
    </td>
  `;
  tblBody.appendChild(tr);
});

// attach events for edit/delete and double-click editing
tblBody.querySelectorAll('.btn-delete').forEach(btn => {
  btn.addEventListener('click', (e) => {
    const id = btn.dataset.id;
    if (!confirm('Hapus item ini?')) return;
    items = items.filter(it => it.id !== id);
    renderTable();
    // tidak otomatis menghitung
  });
});

tblBody.querySelectorAll('.editable').forEach(cell => {
  cell.addEventListener('dblclick', (e) => {
    startEdit(cell);
  });
});

tblBody.querySelectorAll('.btn-edit').forEach(btn => {
  btn.addEventListener('click', (e) => {
    const id = btn.dataset.id;
    const cell = tblBody.querySelector(`[data-id="${id}"][data-field="nama"]`);
    if (cell) startEdit(cell);
  });
});
}

function startEdit(cell){
const field = cell.dataset.field;
const id = cell.dataset.id;
const item = items.find(it => it.id === id);
if (!item) return;
const original = cell.textContent.trim();

const input = document.createElement('input');
input.type = (field === 'nama') ? 'text' : 'number';
input.value = original.replace(' Wh','');
input.style.width = '100%';
input.style.padding = '8px';
input.style.borderRadius = '8px';
cell.innerHTML = '';
cell.appendChild(input);
input.focus();
input.select();

input.addEventListener('keydown', (e) => {
  if (e.key === 'Enter') finishEdit();
  if (e.key === 'Escape') cancelEdit();
});

input.addEventListener('blur', finishEdit);

function finishEdit(){
  let val = input.value.trim();
  if (field === 'nama') {
    if (!val) { alert('Nama tidak boleh kosong'); input.focus(); return; }
    item.nama = val;
  } else {
    const num = Number(val);
    if (!isFinite(num) || num < 0) { alert('Nilai tidak valid'); input.focus(); return; }
    if (field === 'watt') item.watt = num;
    if (field === 'jam') item.jam = num;
    if (field === 'jumlah') item.jumlah = parseInt(num) || 1;
  }
  item.wh = item.watt * item.jam * item.jumlah;
  renderTable();
  // tidak otomatis menghitung
}

function cancelEdit(){
  cell.textContent = original;
}
}

function calcTotals(){
const totalWh = items.reduce((s,it) => s + (Number(it.wh) || 0), 0);
const totalKwh = totalWh/1000;
const harga = parseFloat(hargaInput.value) || 0;
const daysPerMonth = parseInt(daysPerMonthInput.value) || 30;

const costDaily = totalKwh * harga;
const costMonthly = costDaily * daysPerMonth;
const costYearly = costDaily * 365;

totalWhEl.textContent = (totalWh).toLocaleString('id-ID', {maximumFractionDigits:2}) + ' Wh';
totalKwhEl.textContent = totalKwh.toLocaleString('id-ID', {minimumFractionDigits:3, maximumFractionDigits:6}) + ' kWh';
costDailyEl.textContent = toRp(costDaily);
costMonthlyEl.textContent = toRp(costMonthly);
costYearlyEl.textContent = toRp(costYearly);
}

// Excel export using SheetJS with styling (header bg, borders, total bold)
function exportExcel(){
if (items.length === 0) return alert('Tidak ada data untuk diekspor.');

const header = ["No","Nama Barang","Watt (W)","Jam/hari","Jumlah","Wh/hari"];
const wsData = [header];
items.forEach((it, idx) => {
  wsData.push([idx+1, it.nama, it.watt, it.jam, it.jumlah, Number(it.wh.toFixed(2))]);
});

// compute totals
const totalWh = items.reduce((s,it) => s + it.wh, 0);
const totalKwh = totalWh/1000;
const harga = parseFloat(hargaInput.value) || 0;
const daysPerMonth = parseInt(daysPerMonthInput.value) || 30;
const costDaily = totalKwh*harga;
const costMonthly = costDaily * daysPerMonth;
const costYearly = costDaily * 365;

// totals block
wsData.push([]);
wsData.push(["", "TOTAL", "", "", "", ""]);
wsData.push(["", "Total Wh per hari", "", "", "", Number(totalWh.toFixed(2))]);
wsData.push(["", "Total kWh per hari", "", "", "", Number(totalKwh.toFixed(4))]);
wsData.push(["", "Harga per kWh (Rp)", "", "", "", Math.round(harga)]);
wsData.push(["", "Biaya Harian (Rp)", "", "", "", Math.round(costDaily)]);
wsData.push(["", "Biaya Bulanan (Rp)", "", "", "", Math.round(costMonthly)]);
wsData.push(["", "Biaya Tahunan (Rp)", "", "", "", Math.round(costYearly)]);

// Create worksheet + workbook
const ws = XLSX.utils.aoa_to_sheet(wsData);
const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, ws, "Beban Listrik");

// set column widths (auto-fit approx)
const headerLen = header.length;
const colWidths = [];
for (let c = 0; c < headerLen; c++){
  let maxLen = header[c] ? String(header[c]).length : 10;
  for (let r = 0; r < wsData.length; r++){
    const v = wsData[r][c] ? String(wsData[r][c]) : "";
    if (v.length > maxLen) maxLen = v.length;
  }
  colWidths.push({wch: Math.min(Math.max(maxLen + 4, 10), 40)});
}
ws['!cols'] = colWidths;

// Styling: header background, bold; borders for all data cells; bold for totals label rows.
// Note: Some spreadsheet apps may ignore certain style properties depending on renderer.
function cellAddress(r, c){
  // r and c are 0-based for wsData -> convert to A1
  const col = XLSX.utils.encode_col(c);
  const row = r + 1;
  return col + row;
}

// border style definition (thin light grey)
const border = {
  top: {style: "thin", color: {rgb: "FFBFBFBF"}},
  bottom: {style: "thin", color: {rgb: "FFBFBFBF"}},
  left: {style: "thin", color: {rgb: "FFBFBFBF"}},
  right: {style: "thin", color: {rgb: "FFBFBFBF"}}
};

// header style
const headerStyle = {
  font: {bold: true, color: {rgb: "FF111827"}},
  fill: {patternType: "solid", fgColor: {rgb: "FFF3F4F6"}}, // light grey
  alignment: {horizontal: "center", vertical: "center"},
  border
};

// normal cell style
const normalStyle = {
  font: {bold: false, color: {rgb: "FF111827"}},
  alignment: {horizontal: "left", vertical: "center"},
  border
};

// numeric right-align style
const numericStyle = Object.assign({}, normalStyle, {alignment: {horizontal: "right", vertical: "center"}});

// total label style (bold)
const totalLabelStyle = {
  font: {bold: true, color: {rgb: "FF111827"}},
  alignment: {horizontal: "left", vertical: "center"},
  border
};

// total value style (bold, right align)
const totalValueStyle = {
  font: {bold: true, color: {rgb: "FF111827"}},
  alignment: {horizontal: "right", vertical: "center"},
  border
};

// apply styles across wsData range
for (let r = 0; r < wsData.length; r++){
  for (let c = 0; c < headerLen; c++){
    const addr = cellAddress(r, c);
    if (!ws[addr]) continue; // skip empty cells
    // header row
    if (r === 0){
      ws[addr].s = headerStyle;
      // center numeric header
      if (c === 0 || c === 3 || c === 4 || c === 5) {
        ws[addr].s.alignment = {horizontal: "center", vertical: "center"};
      }
    } else {
      // totals section starts at row index where cell in column 1 equals "TOTAL" or "Total Wh per hari", etc.
      const cellVal = String(wsData[r][1] ?? "").toLowerCase();
      const isTotalsRow = (cellVal.includes('total') || cellVal.includes('biaya') || cellVal.includes('harga'));
      if (isTotalsRow){
        // make label column bold for totals rows (col B)
        if (c === 1){
          ws[addr].s = totalLabelStyle;
        } else if (c === headerLen - 1){
          ws[addr].s = totalValueStyle;
        } else {
          ws[addr].s = normalStyle;
        }
      } else {
        // normal rows: numeric columns right aligned
        if (c === 0 || c === 3 || c === 4) {
          // center small numeric cells (No, Jam/hari, Jumlah)
          ws[addr].s = Object.assign({}, normalStyle, {alignment:{horizontal:"center", vertical:"center"}});
        } else if (c === headerLen - 1) {
          // Wh/hari align right
          ws[addr].s = numericStyle;
        } else {
          ws[addr].s = normalStyle;
        }
      }
    }
  }
}

const fname = "beban_listrik_" + new Date().toISOString().slice(0,10) + ".xlsx";
try {
  XLSX.writeFile(wb, fname, {bookType: "xlsx", bookSST: false, cellStyles: true});
} catch (err) {
  XLSX.writeFile(wb, fname);
}

// small helpers
function escapeHtml(s){
return s.replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;');
}
