const excelUrl = "https://raw.githubusercontent.com/Emadk200/ek-stocks/main/watchliststoks.xlsx";

// تحميل الملف من GitHub
async function loadExcel() {
  const container = document.getElementById("tableContainer");
  container.innerHTML = "<p class='text-gray-600'>⏳ Loading data...</p>";

  try {
    const response = await fetch(excelUrl);
    if (!response.ok) throw new Error("File not found");

    const arrayBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: "array" });
    const sheet = workbook.SheetNames[0];
    const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheet], { defval: "" });
    renderTable(data);
    updateDate();
  } catch (err) {
    container.innerHTML = `<p class='text-red-500 font-semibold'>⚠️ Error loading Excel file: ${err.message}</p>`;
  }
}

// تنسيق الخلايا
function formatCell(value, header, isLastRow = false) {
  if (typeof value === "number") {
    const isPercent = /%|Dev|Change|growth/i.test(header);
    let display = value.toFixed(2);
    if (isPercent) display += "%";
    const color = isLastRow ? "text-white" : (value < 0 ? "text-red-600" : "text-gray-800");
    return `<td class="px-3 py-1 ${color}">${display}</td>`;
  }
  return `<td class="px-3 py-1 ${isLastRow ? 'text-white' : ''}">${value}</td>`;
}

// عرض الجدول
function renderTable(data) {
  const container = document.getElementById("tableContainer");
  if (!data || data.length === 0) {
    container.innerHTML = "<p>No data found in Excel file.</p>";
    return;
  }

  const headers = Object.keys(data[0]);
  const lastRowIndex = data.length - 1;

  // حفظ التخطيط الأساسي (zebra) بشكل دائم حسب الفهرس الأصلي
  const rowColors = data.map((_, i) =>
    i === lastRowIndex ? "bg-gray-800 text-white font-semibold" :
    i % 2 === 0 ? "bg-white" : "bg-gray-300"
  );

  let tableHTML = `
    <table id="dataTable" class="min-w-full border border-gray-300 rounded-lg overflow-hidden">
      <thead class="bg-gray-700 text-white select-none">
        <tr>${headers.map(h => `<th class="px-3 py-2 text-left cursor-pointer">${h}</th>`).join("")}</tr>
      </thead>
      <tbody>
        ${data.map((row, i) => `
          <tr class="${rowColors[i]}">
            ${headers.map(h => formatCell(row[h], h, i === lastRowIndex)).join("")}
          </tr>
        `).join("")}
      </tbody>
    </table>
  `;

  container.innerHTML = tableHTML;
  addSorting(rowColors);
  addSearch();
}

// البحث
function addSearch() {
  const searchInput = document.getElementById("searchInput");
  const rows = document.querySelectorAll("#dataTable tbody tr");
  searchInput.addEventListener("input", () => {
    const term = searchInput.value.toLowerCase();
    rows.forEach(row => {
      row.style.display = row.textContent.toLowerCase().includes(term) ? "" : "none";
    });
  });
}

// الفرز مع الحفاظ على zebra striping الأصلي
function addSorting(rowColors) {
  const table = document.getElementById("dataTable");
  const headers = table.querySelectorAll("th");
  headers.forEach((th, i) => {
    th.addEventListener("click", () => sortTable(table, i, rowColors));
  });
}

function sortTable(table, columnIndex, rowColors) {
  const tbody = table.querySelector("tbody");
  const rows = Array.from(tbody.rows);
  const lastRow = rows.pop(); // استبعاد السطر الأخير (المجاميع)

  const sorted = rows.sort((a, b) => {
    const A = a.cells[columnIndex].innerText.replace("%", "");
    const B = b.cells[columnIndex].innerText.replace("%", "");
    const numA = parseFloat(A);
    const numB = parseFloat(B);
    if (!isNaN(numA) && !isNaN(numB)) return numA - numB;
    return A.localeCompare(B);
  });

  // إعادة ترتيب الصفوف دون تغيير ألوانها الأصلية
  sorted.forEach((r, idx) => {
    r.className = rowColors[idx];
    tbody.appendChild(r);
  });
  lastRow.className = rowColors[rowColors.length - 1];
  tbody.appendChild(lastRow);
}

// التاريخ والملاحظة
function updateDate() {
  const now = new Date();
  const formatted = now.toLocaleString("en-US", {
    dateStyle: "full",
    timeStyle: "short"
  });
  document.getElementById("date").textContent = `Last updated: ${formatted}`;
  document.getElementById("note").textContent = "Data loaded automatically from GitHub (watchliststoks.xlsx). Click 🔄 to refresh.";
  document.getElementById("note").textContent = "Book Value For Q3";
  document.getElementById("note").textContent = "Graham Number: above it the stock is over priced";
  document.getElementById("note").textContent = "L.P /G.N ==> Last Price / Graham Number";
  document.getElementById("note").textContent = "Stock Yeild based on last price and the expected dividends for 2025";
}

// زر التحديث
document.getElementById("reloadBtn").addEventListener("click", loadExcel);

// تحميل أولي
loadExcel();
