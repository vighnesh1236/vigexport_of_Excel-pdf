let headers = [];
let selectedRow = null;
let selectedColIndex = null;

window.onload = function () {
  loadSavedData();
};

document.getElementById('fileInput').addEventListener('change', function (e) {
  const file = e.target.files[0];
  const reader = new FileReader();
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    headers = jsonData[0];
    renderTable(jsonData);
  };
  reader.readAsArrayBuffer(file);
});

function renderTable(data) {
  let tableHTML = '<table id="excelTable"><thead><tr>';
  headers.forEach((header, i) => {
    tableHTML += `<th contenteditable="true" onclick="selectColumn(this, ${i})">${header}</th>`;
  });
  tableHTML += '</tr></thead><tbody>';
  for (let i = 1; i < data.length; i++) {
    tableHTML += `<tr onclick="selectRow(this)">`;
    for (let j = 0; j < headers.length; j++) {
      const value = data[i][j] !== undefined ? data[i][j] : '';
      tableHTML += `<td contenteditable="true">${value}</td>`;
    }
    tableHTML += '</tr>';
  }
  tableHTML += '</tbody></table>';
  document.getElementById('tableContainer').innerHTML = tableHTML;
}

function createEmptyTable() {
  headers = ['Column1', 'Column2'];
  const data = [headers, ['', '']];
  renderTable(data);
}

function addRow() {
  const table = document.getElementById('excelTable');
  if (!table) return alert("Please upload or create a table first.");
  const row = table.insertRow(-1);
  row.onclick = () => selectRow(row);
  headers.forEach(() => {
    const cell = row.insertCell();
    cell.contentEditable = true;
    cell.innerText = '';
  });
}

function addColumn() {
  addColumnAt(headers.length);
}

function addRowAt() {
  const index = parseInt(document.getElementById('rowIndex').value);
  const table = document.getElementById('excelTable');
  if (!table || isNaN(index)) return;

  const tbody = table.querySelector('tbody');
  const newRow = document.createElement('tr');
  newRow.onclick = () => selectRow(newRow);

  headers.forEach(() => {
    const td = document.createElement('td');
    td.contentEditable = true;
    td.innerText = '';
    newRow.appendChild(td);
  });

  if (index >= 0 && index <= tbody.rows.length) {
    tbody.insertBefore(newRow, tbody.rows[index]);
  } else {
    tbody.appendChild(newRow);
  }
}

function addColumnAt(index = null) {
  const table = document.getElementById('excelTable');
  if (!table) return;
  const colIndex = index !== null ? index : parseInt(document.getElementById('colIndex').value);
  if (isNaN(colIndex) || colIndex < 0) return;

  const newHeader = `Column${headers.length + 1}`;
  headers.splice(colIndex, 0, newHeader);

  const theadRow = table.querySelector('thead tr');
  const newTh = document.createElement('th');
  newTh.contentEditable = true;
  newTh.innerText = newHeader;
  newTh.onclick = () => selectColumn(newTh, colIndex);
  theadRow.insertBefore(newTh, theadRow.children[colIndex]);

  table.querySelectorAll('tbody tr').forEach(row => {
    const td = document.createElement('td');
    td.contentEditable = true;
    td.innerText = '';
    row.insertBefore(td, row.children[colIndex]);
  });
}

function deleteSelectedRow() {
  if (!selectedRow) return alert("Please select a row to delete.");
  selectedRow.remove();
  selectedRow = null;
}

function deleteSelectedColumn() {
  if (selectedColIndex === null) return alert("Please select a column to delete.");
  const table = document.getElementById('excelTable');
  headers.splice(selectedColIndex, 1);

  table.querySelectorAll('thead tr').forEach(tr => tr.deleteCell(selectedColIndex));
  table.querySelectorAll('tbody tr').forEach(tr => tr.deleteCell(selectedColIndex));
  selectedColIndex = null;
}

function selectRow(row) {
  const allRows = document.querySelectorAll('tbody tr');
  allRows.forEach(r => r.classList.remove('selected'));
  row.classList.add('selected');
  selectedRow = row;
}

function selectColumn(th, index) {
  const allHeaders = document.querySelectorAll('thead th');
  allHeaders.forEach(h => h.classList.remove('selected'));
  th.classList.add('selected');
  selectedColIndex = index;
}

function saveData() {
  const table = document.getElementById('excelTable');
  if (!table) return;

  const rows = [];
  const trs = table.querySelectorAll('tr');
  trs.forEach(tr => {
    const row = [];
    tr.querySelectorAll('th, td').forEach(cell => {
      row.push(cell.innerText.trim());
    });
    rows.push(row);
  });

  headers = rows[0];
  localStorage.setItem("excelTableData", JSON.stringify(rows));
  alert("Data saved permanently in your browser.");
}

function loadSavedData() {
  const saved = localStorage.getItem("excelTableData");
  if (saved) {
    const data = JSON.parse(saved);
    headers = data[0];
    renderTable(data);
  }
}

function downloadExcel() {
  const table = document.getElementById('excelTable');
  if (!table) return alert("No table found to save.");
  const rows = [];
  const trs = table.querySelectorAll('tr');
  trs.forEach(tr => {
    const row = [];
    tr.querySelectorAll('th, td').forEach(cell => {
      row.push(cell.innerText.trim());
    });
    rows.push(row);
  });

  const ws = XLSX.utils.aoa_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
  XLSX.writeFile(wb, 'table_data.xlsx');
}

function downloadPDF() {
  const table = document.getElementById('excelTable');
  if (!table) return alert("No table to export.");
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF();

  const rows = [];
  const trs = table.querySelectorAll('tbody tr');
  trs.forEach(tr => {
    const row = [];
    tr.querySelectorAll('td').forEach(cell => {
      row.push(cell.innerText.trim());
    });
    rows.push(row);
  });

  const headersRow = Array.from(table.querySelectorAll('thead th')).map(th => th.innerText.trim());

  doc.autoTable({
    head: [headersRow],
    body: rows
  });

  doc.save('table_data.pdf');
}
