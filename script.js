const excelUrl = "https://raw.githubusercontent.com/Emadk200/ek-stocks/main/watchliststoks.xlsx";

// تحميل الملف
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
function formatCell(value, header) {
  if (typeof value === "number") {
    const isPercent = /%|Dev|Change|growth/i.test(header);
    let display = value.toFixed(2);
    if (isPercent) display += "%";
    const color = value < 0 ? "text-red-600" : "text-gray-800";
    return `<td class="px-3 py-1 ${color}">${display}</td>`;
  }
  return `<td class="px-3 py-1">${value}</td>`;
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

  let tableHTML = `
    <table id="dataTable" class="min-w-full border border-gray-300 rounded-lg overflow-hidden">
      <thead class="bg-gray-700 text-white">
        <tr>${headers.map(h => `<th class="px-3 py-2 text-left cursor-pointer select-none">${h}</th>`).join("")}</tr>
      </thead>
      <tbody>
        ${data.map((row, i) => `
          <tr class="${i === lastRowIndex ? "bg-gray-800 text-white font-semibold" : (i % 2 === 0 ? "bg-white" : "bg-gray-300")}">
            ${headers.map(h => formatCell(row[h], h)).join("")}
          </tr>
        `).join("")}
      </tbody>
    </table>
  `;

  container.innerHTML = tableHTML;
  addSorting();
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

// الفرز
function addSorting() {
  const table = document.getElementById("dataTable");
  const headers = table.querySelectorAll("th");
  headers.forEach((th, i) => {
    th.addEventListener("click", () => sortTable(table, i));
  });
}

function sortTable(table, columnIndex) {
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

  sorted.forEach(r => tbody.appendChild(r));
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
}

// زر التحديث
document.getElementById("reloadBtn").addEventListener("click", loadExcel);

// تحميل أولي
loadExcel();
