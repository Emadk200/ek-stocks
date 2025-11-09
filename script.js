const fileUrl = "https://emadk200.github.io/ek-stocks/watchliststoks.xlsx?v=" + Date.now();
const proxyURL = "https://bitter-frost-8e4d.emk200.workers.dev/";

// التاريخ
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
- Dashboard updates Last Price live from ASE.
- Change & Change% are interpreted directly from your Excel file.
`;

// جلب الأسعار من البورصة
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

    window._lastPricesTest = prices;
    return prices;

  } catch (err) {
    console.error("Error fetching prices:", err);
    return {};
  }
}

// تحميل ملف Excel ودمج الأسعار
async function loadExcelData() {
  document.getElementById("loading-indicator").classList.remove("hidden");

  const res = await fetch(fileUrl);
  const buf = await res.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array" });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const data = XLSX.utils.sheet_to_json(sheet, { defval: "" });

  const marketPrices = await fetchLastClosingPrices();

  data.forEach(row => {
    if (row.Symbol && marketPrices[row.Symbol]) {
      row["Last Price"] = marketPrices[row.Symbol];
    }
  });

  renderTable(data);
  document.getElementById("loading-indicator").classList.add("hidden");
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

  // الصفوف
  data.forEach((row, i) => {
    const isTotalRow = row.Symbol && row.Symbol.toLowerCase().includes("tot");
    const tr = document.createElement("tr");
    tr.className = isTotalRow ? "bg-gray-700 text-white font-bold" : (i % 2 === 0 ? "bg-gray-100" : "bg-gray-200");

    const changeVal = parseFloat(row["Change"]);
    const changePctVal = parseFloat(row["Change %"]) * 100;

    columns.forEach(col => {
      let value = row[col];
      const td = document.createElement("td");
      td.className = "px-3 py-1 border-b border-gray-300 text-sm";

      if (col === "Change" && !isNaN(changeVal)) {
        value = changeVal.toFixed(3);
        td.classList.add(changeVal > 0 ? "text-green-600" : changeVal < 0 ? "text-red-600" : "text-gray-800");
      }

      if (col === "Change %" && !isNaN(changePctVal)) {
        value = Math.round(changePctVal) + "%";
        td.classList.add(changePctVal > 0 ? "text-green-600" : changePctVal < 0 ? "text-red-600" : "text-gray-800");
      }

      if (col === "Last Price" && !isNaN(changeVal)) {
        const arrow = changeVal > 0 ? "↑" : changeVal < 0 ? "↓" : "-";
        const arrowColor = changeVal > 0 ? "text-green-600" : changeVal < 0 ? "text-red-600" : "text-gray-600";
        td.innerHTML = `${value} <span class="${arrowColor}">${arrow}</span>`;
      } else {
        td.textContent = value;
      }

      tr.appendChild(td);
    });

    tbody.appendChild(tr);
  });
}

// البحث
document.getElementById("searchInput").addEventListener("keyup", function () {
  const filter = this.value.toLowerCase();
  document.querySelectorAll("#dataTable tbody tr").forEach((row, i, rows) => {
    if (i === rows.length - 1) return;
    row.style.display = row.textContent.toLowerCase().includes(filter) ? "" : "none";
  });
});

// الفرز
function sortTableByColumn(colIndex) {
  const tbody = document.querySelector("#dataTable tbody");
  const rows = Array.from(tbody.querySelectorAll("tr"));
  const totalRow = rows.pop();
  const asc = tbody.dataset.sortOrder !== "asc";
  tbody.dataset.sortOrder = asc ? "asc" : "desc";

  rows.sort((a, b) => {
    const aText = a.children[colIndex].innerText.replace(/[↑↓%-]/g, "").trim();
    const bText = b.children[colIndex].innerText.replace(/[↑↓%-]/g, "").trim();
    const aNum = parseFloat(aText), bNum = parseFloat(bText);
    return (!isNaN(aNum) && !isNaN(bNum)) ? (asc ? aNum - bNum : bNum - aNum) : (asc ? aText.localeCompare(bText) : bText.localeCompare(aText));
  });

  rows.forEach((row, i) => {
    row.className = i % 2 === 0 ? "bg-gray-100" : "bg-gray-200";
    tbody.appendChild(row);
  });

  tbody.appendChild(totalRow);
}

document.getElementById("refresh-btn").addEventListener("click", loadExcelData);
loadExcelData();
