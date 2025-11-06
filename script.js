document.getElementById('fileInput').addEventListener('change', handleFile, false);

function handleFile(e) {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = function (event) {
    try {
      const data = new Uint8Array(event.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });
      renderTable(json);
    } catch (err) {
      console.error('حدث خطأ أثناء قراءة الملف:', err);
      alert('حدث خطأ أثناء قراءة الملف. يرجى التحقق من الصيغة.');
    }
  };
  reader.readAsArrayBuffer(file);
}

function renderTable(data) {
  const tableContainer = document.getElementById('tableContainer');
  tableContainer.innerHTML = '';

  const table = document.createElement('table');
  table.className = 'min-w-full border border-gray-400 divide-y divide-gray-400 text-sm text-left';

  // Header
  const thead = document.createElement('thead');
  thead.className = 'bg-gray-700 text-white';
  const headerRow = document.createElement('tr');

  data[0].forEach(headerText => {
    const th = document.createElement('th');
    th.textContent = headerText;
    th.className = 'px-3 py-2 cursor-pointer select-none';
    th.addEventListener('click', () => sortTable(th.cellIndex));
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);
  table.appendChild(thead);

  // Body
  const tbody = document.createElement('tbody');
  const rows = data.slice(1);

  rows.forEach((rowData, rowIndex) => {
    const row = document.createElement('tr');

    // zebra striping
    row.className = rowIndex % 2 === 0 ? 'bg-white' : 'bg-gray-300';

    rowData.forEach((cellData, cellIndex) => {
      const td = document.createElement('td');
      td.className = 'px-3 py-1 border border-gray-400 text-sm';

      if (typeof cellData === 'number') {
        const isPercentage = data[0][cellIndex]?.toString().includes('%');
        let value = cellData.toFixed(2);
        if (isPercentage) value += '%';
        td.textContent = value;
        if (cellData < 0) td.classList.add('text-red-600');
      } else {
        td.textContent = cellData ?? '';
      }
      row.appendChild(td);
    });
    tbody.appendChild(row);
  });

  table.appendChild(tbody);
  tableContainer.appendChild(table);

  // تحديد آخر صف (مجاميع/متوسطات) وتغيير لونه
  const lastRow = tbody.lastElementChild;
  if (lastRow) {
    lastRow.className = 'bg-gray-800 text-white font-semibold';
  }
}

// البحث
document.getElementById('searchInput').addEventListener('input', function () {
  const filter = this.value.toLowerCase();
  const rows = document.querySelectorAll('tbody tr');
  rows.forEach(row => {
    const text = row.textContent.toLowerCase();
    row.style.display = text.includes(filter) ? '' : 'none';
  });
});

// الفرز
function sortTable(columnIndex) {
  const tbody = document.querySelector('tbody');
  const rows = Array.from(tbody.querySelectorAll('tr'));

  // استبعاد آخر صف من الفرز (المجاميع)
  const lastRow = rows.pop();

  const sortedRows = rows.sort((a, b) => {
    const aText = a.children[columnIndex].innerText.replace('%', '');
    const bText = b.children[columnIndex].innerText.replace('%', '');
    const aNum = parseFloat(aText);
    const bNum = parseFloat(bText);
    if (!isNaN(aNum) && !isNaN(bNum)) {
      return aNum - bNum;
    } else {
      return aText.localeCompare(bText);
    }
  });

  tbody.innerHTML = '';
  sortedRows.forEach(row => tbody.appendChild(row));
  tbody.appendChild(lastRow);
}

// عرض التاريخ الميلادي والرسالة
function updateDateAndMessage() {
  const dateElement = document.getElementById('date');
  const noteElement = document.getElementById('note');
  const now = new Date();

  const dateStr = now.toLocaleString('en-GB', {
    weekday: 'short',
    year: 'numeric',
    month: 'short',
    day: 'numeric',
    hour: '2-digit',
    minute: '2-digit',
  });

  dateElement.textContent = `Last update: ${dateStr}`;
  noteElement.textContent = 'Data automatically calculated based on uploaded Excel sheet.';
}

updateDateAndMessage();
