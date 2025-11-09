const fileUrl = "https://emadk200.github.io/ek-stocks/watchliststoks.xlsx?v=" + Date.now();
const proxyURL = "https://bitter-frost-8e4d.emk200.workers.dev/";

// عرض التاريخ
document.getElementById("current-date").textContent =
  new Date().toLocaleString("en-US", {
    weekday: "long",
    year: "numeric",
    month: "long",
    day: "numeric",
    hour: "2-digit",
    minute: "2-digit",
  });

// الملاحظة الثابتة
document.getElementById("notes").innerHTML = `
- Dashboard updates Last Price live from ASE via proxy.
- Change & Change% are interpreted directly from your Excel file.
`;

// جلب الأسعار من البورصة عبر الـ Worker
async function fetchLastClosingPrices() {
  try {
    const res = await fetch(proxyURL);
    const html = await res.text();
    const parser = new DOMParser();
    const doc = parser.parseFromString(html, "text/html");

    const rows = doc.querySelectorAll("table tbody tr");
    const prices = {};

    rows.forEach(row => {
      const cells = row.querySelectorAll("td");
      if (cells.length > 7) {
        const symbol = cells[6]?.innerText.trim();
        const lastClosing = parseFloat(cells[7]?.innerText.trim());
        if (symbol && !isNaN(lastClosing)) {
          prices[symbol] = lastClosing;
        }
      }
    });

    // ✅ نضيف المتغير هنا
    window._lastPricesTest = prices;

    return prices;

  } catch (error) {
    console.error("Error fetching prices:", error);
    window._lastPricesTest = {}; // حتى لا يظهر undefined
    return {};
  }
}

// تحميل البيانات وعرضها
async function loadExcelData() {
  const loading = document.getElementById("loading-indicator");
  try {
    loading.classList.remove("hidden");

    const response = await fetch(fileUrl);
    const arrayBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(sheet, { defval: "" });

    const marketPrices = await fetchLastClosingPrices();

    data.forEach(row => {
      if (marketPrices[row.Symbol]) {
        row["Last Price"] = marketPrices[row.Symbol];
      }
    });

    renderTable(data);

  } finally {
    loading.classList.add("hidden");
  }
}

// عرض الجدول
function renderTable(data) {
  const table = document.getElementById("dataTable");
  const thead = table.querySelector("thead");
  const tbody = table.querySelector("tbody");

  thead.innerHTML = "";
  tbody.innerHTML = "";

  const columns = Object.keys(data[0]);

  // رأس الجدول
  const headerRow = document.createElement("tr");
  columns.forEach((col, index) => {
    const th = document.createElement("th");
    th.className = "px-3 py-2 bg-gray-800 text-white font-semibold cursor-pointer";
    th.textContent = col;
    th.addEventListener("click", () => sortTableByColumn(index));
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);

  // صفوف
  data.forEach((row, i) => {
    const tr = document.createElement("tr");
    const isTotalRow = row[columns[0]].toString().toLowerCase().includes("total");
    tr.className = isTotalRow ? "bg-gray-800 text-white font-bold" : (i % 2 === 0 ? "bg-gray-100" : "bg-gray-200");

    const changeColumn = Object.keys(row).find(c =>
      c.toLowerCase().includes("change") ||
      c.toLowerCase().includes("diff") ||
      c.toLowerCase().includes("dev")
    );

    const rawChange = changeColumn ? parseFloat(row[changeColumn]) : NaN;

    columns.forEach(col => {
      let value = row[col];
      const td = document.createElement("td");
      td.className = "px-3 py-1 border-b border-gray-300 text-sm";

      // تنسيق النسب %
      if (col.toLowerCase().includes("pct") || col.includes("%")) {
        value = Math.round(parseFloat(value) * 100) + "%";
      }

      // تنسيق Change ألوان
      if (col === changeColumn) {
        if (rawChange > 0) td.classList.add("text-green-600", "font-semibold");
        else if (rawChange < 0) td.classList.add("text-red-600", "font-semibold");
        else td.classList.add("text-gray-800");
      }

      // إضافة السهم بجانب Last Price
      if (col === "Last Price" && !isNaN(rawChange)) {
        let arrow = rawChange > 0 ? "↑" : rawChange < 0 ? "↓" : "-";
        let arrowClass = rawChange > 0 ? "text-green-600" : rawChange < 0 ? "text-red-600" : "text-gray-600";
        td.innerHTML = `<span>${value}</span> <span class="${arrowClass} ml-1">${arrow}</span>`;
        tr.appendChild(td);
        return;
      }

      td.textContent = value;
      tr.appendChild(td);
    });

    tbody.appendChild(tr);
  });
}

// بحث
document.getElementById("searchInput").addEventListener("keyup", function () {
  const filter = this.value.toLowerCase();
  const rows = document.querySelectorAll("#dataTable tbody tr");
  rows.forEach((row, index) => {
    if (index === rows.length - 1) return;
    row.style.display = row.textContent.toLowerCase().includes(filter) ? "" : "none";
  });
});

// فرز
function sortTableByColumn(colIndex) {
  const tbody = document.querySelector("#dataTable tbody");
  const rows = Array.from(tbody.querySelectorAll("tr"));
  const totalRow = rows.pop();
  const ascending = tbody.dataset.sortOrder !== "asc";
  tbody.dataset.sortOrder = ascending ? "asc" : "desc";

  rows.sort((a, b) => {
    const aVal = a.children[colIndex].innerText.replace(/[↑↓%-]/g, "").trim();
    const bVal = b.children[colIndex].innerText.replace(/[↑↓%-]/g, "").trim();
    const aNum = parseFloat(aVal), bNum = parseFloat(bVal);
    return (!isNaN(aNum) && !isNaN(bNum)) ? (ascending ? aNum - bNum : bNum - aNum) : (ascending ? aVal.localeCompare(bVal) : bVal.localeCompare(aVal));
  });

  rows.forEach((row, i) => {
    row.className = i % 2 === 0 ? "bg-gray-100" : "bg-gray-200";
    tbody.appendChild(row);
  });

  tbody.appendChild(totalRow);
}

document.getElementById("refresh-btn").addEventListener("click", loadExcelData);
loadExcelData();
