const EXCEL_URL = "https://emadk200.github.io/ek-stocks/watchliststocks.xlsx";
const ASE_DAILY_URL = "https://bitter-frost-8e4d.emk200.workers.dev/sites/default/files/daily-bulletin/en/Daily%20Bulletin.xlsx";

let jsonData = [];
let tableHeaders = [];

// ========================= LOAD WATCHLIST =========================
async function loadExcelData() {
  document.getElementById("loadingStatus").style.display = "block";

  const resp = await fetch(EXCEL_URL);
  const buf = await resp.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array" });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  let data = XLSX.utils.sheet_to_json(sheet, { defval: "" });

  data = data.filter(r => r.Symbol && r.Symbol.trim() !== "");

  let totalRow = data.find(r => r.Symbol === "TotAvg" || r.Symbol === "Tot/Avg");
  data = data.filter(r => r.Symbol !== "TotAvg" && r.Symbol !== "Tot/Avg");

  const asePrices = await fetchASEPrices();
  data.forEach(row => {
    const sym = row.Symbol.trim().toUpperCase();
    if (asePrices[sym] !== undefined) row["Last  Price"] = asePrices[sym];
  });

  data.forEach(row => formatRow(row));

  if (totalRow) {
    formatTotalsRow(totalRow);
    data.push(totalRow);
  }

  jsonData = data;
  renderTable(data);
  updateTimeStamp();
  document.getElementById("loadingStatus").style.display = "none";
}

// ========================= FETCH ASE PRICES =========================
async function fetchASEPrices() {
  try {
    const resp = await fetch(ASE_DAILY_URL);
    const buf = await resp.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

    const prices = {};
    rows.forEach(r => {
      const symbol = (r["__EMPTY_6"] || "").toString().trim().toUpperCase();
      const price = parseFloat(r["__EMPTY_8"]);
      if (symbol && !isNaN(price)) prices[symbol] = price;
    });
    return prices;
  } catch {
    return {};
  }
}

// ========================= FORMAT ROWS =========================
function formatRow(row) {
  const percentCols = ["Change %", "Cash Div% 24", "Average% 5Y", "Stock Yeild% 2024", "Average % 5Y"];
  const floatCols = ["Last  Price", "EPS 2024 Q3", "EPS 2025 Q3", "EPS Forecast 2025", "PE Ratio", "Fair Price Q3 2025", "PB Ratio", "P/E * P/B.V.", "Graham No", "L.P /G.N"];
  const intCols = ["ASE Code"];

  for (let key in row) {
    let val = row[key];

    if (intCols.includes(key)) {
      row[key] = parseInt(val) || "";
      continue;
    }

    if (percentCols.includes(key)) {
      let num = parseFloat(val);
      if (!isNaN(num)) {
        if (Math.abs(num) < 1) num = num * 100;
        row[key] = `${Math.round(num)}%`;
      }
      continue;
    }

    if (floatCols.includes(key)) {
      let num = parseFloat(val);
      if (!isNaN(num)) row[key] = num.toFixed(2);
      continue;
    }
  }
}

function formatTotalsRow(row) {
  for (let k in row) {
    if (!isNaN(parseFloat(row[k]))) row[k] = parseFloat(row[k]).toFixed(2);
  }
}

// ========================= RENDER =========================
function renderTable(data) {
  const table = document.getElementById("dataTable");
  const thead = table.querySelector("thead");
  const tbody = table.querySelector("tbody");

  thead.innerHTML = "";
  tbody.innerHTML = "";

  tableHeaders = Object.keys(data[0]);

  const tr = document.createElement("tr");
  tableHeaders.forEach(h => {
    const th = document.createElement("th");
    th.className = "px-3 py-2 bg-gray-800 text-white text-sm cursor-pointer";
    th.innerText = h;
    th.addEventListener("click", () => sortTable(h));
    tr.appendChild(th);
  });
  thead.appendChild(tr);

  data.forEach(row => {
  const tr = document.createElement("tr");

  // ✅ إذا كان صف Tot/Avg → أعطِ تنسيق خاص
  const isTotal = row["Symbol"] === "TotAvg" || row["Symbol"] === "Tot/Avg";
  if (isTotal) {
    tr.className = "bg-gray-800 text-white font-bold";
  }

  tableHeaders.forEach(key => {
    const td = document.createElement("td");

    td.className = isTotal
      ? "px-3 py-1 border-b border-gray-700 text-sm"
      : "px-3 py-1 border-b text-sm";

    td.innerText = row[key];
    tr.appendChild(td);
  });

  tbody.appendChild(tr);
});


  applyColorAndArrows();
}

// ========================= SORT (EXCLUDE TotAvg) =========================
function sortTable(column) {
  let totalRow = jsonData.find(r => r.Symbol === "TotAvg" || r.Symbol === "Tot/Avg");
  let list = jsonData.filter(r => r.Symbol !== "TotAvg" && r.Symbol !== "Tot/Avg");

  list.sort((a, b) => (a[column] > b[column] ? 1 : -1));

  if (totalRow) list.push(totalRow);
  jsonData = list;
  renderTable(jsonData);
}

// ========================= SEARCH =========================
document.getElementById("searchInput").addEventListener("input", e => {
  const txt = e.target.value.toLowerCase();
  renderTable(jsonData.filter(r => JSON.stringify(r).toLowerCase().includes(txt)));
});

// ========================= COLOR + ARROWS =========================
function applyColorAndArrows() {
  let tbody = document.querySelector("#dataTable tbody").rows;
  let c1 = tableHeaders.indexOf("Change");
  let c2 = tableHeaders.indexOf("Change %");

  for (let row of tbody) {
    if (row.cells[0].innerText.toLowerCase().includes("tot")) continue;

    if (c1 >= 0) {
      let td = row.cells[c1];
      let v = parseFloat(td.innerText);
      if (v > 0) td.innerHTML = `+${v.toFixed(2)} ↑`, td.style.color = "green";
      else if (v < 0) td.innerHTML = `${v.toFixed(2)} ↓`, td.style.color = "red";
      else td.style.color = "black";
    }

    if (c2 >= 0) {
      let td = row.cells[c2];
      let v = parseFloat(td.innerText);
      if (Math.abs(v) < 1) v = v * 100;
      let disp = `${Math.round(v)}%`;
      if (v > 0) td.innerHTML = `${disp} ↑`, td.style.color = "green";
      else if (v < 0) td.innerHTML = `${disp} ↓`, td.style.color = "red";
      else td.innerHTML = disp, td.style.color = "black";
    }
  }
}

// ========================= TIMESTAMP =========================
function updateTimeStamp() {
  document.getElementById("dateTime").innerText =
    new Date().toLocaleString("en-US", { dateStyle: "full", timeStyle: "short" });
}

// ========================= REFRESH BUTTON =========================
document.getElementById("refreshBtn").addEventListener("click", loadExcelData);

// ========================= START =========================
loadExcelData();
