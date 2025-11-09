const fileUrl = "https://emadk200.github.io/ek-stocks/watchliststoks.xlsx?v=" + Date.now();

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

// ملاحظة متعددة الأسطر (اكتب ملاحظتك هنا)
document.getElementById("notes").innerHTML = `
- This dashboard updates live prices from ASE website.
- Last closing price is automatically fetched.
- Data refresh happens when clicking "Update Data".
`;

// جلب أسعار الإغلاق الأخيرة من موقع البورصة
async function fetchLastClosingPrices() {
  try {
    const res = await fetch("https://www.ase.com.jo/en/bulletins/daily/new");
    const html = await res.text();
    const parser = new DOMParser();
    const doc = parser.parseFromString(html, "text/html");

    const rows = doc.querySelectorAll("table tbody tr");
    const prices = {};

    rows.forEach(row => {
      const cells = row.querySelectorAll("td");
      if (cells.length > 7) {
        const symbol = cells[6]?.innerText.trim(); // Symbol column
        const lastClosing = cells[7]?.innerText.trim(); // Last closing price column

        if (symbol && lastClosing && !isNaN(lastClosing)) {
          prices[symbol] = parseFloat(lastClosing);
        }
      }
    });

    return prices;

  } catch (error) {
    console.error("Error fetching prices:", error);
    return {};
  }
}

// تحميل ملف الإكسل + دمج الأسعار + عرض الجدول
async function loadExcelData() {
  const loading = document.getElementById("loading-indicator");
  try {
    loading.classList.remove("hidden");

    // قراءة ملف Excel من GitHub
    const response = await fetch(fileUrl);
    const arrayBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(sheet, { defval: "" });

    // جلب الأسعار اليومية من البورصة
    const marketPrices = await fetchLastClosingPrices();

    // دمج الأسعار في الجدول
    data.forEach(row => {
      if (marketPrices[row.Symbol]) {
          row["Last Price"] = marketPrices[row.Symbol];
      }
    });

    renderTable(data);

  } catch (error) {
    alert("حدث خطأ أثناء تحميل البيانات");
    console.error(error);
  } finally {
    loading.classList.add("hidden");
  }
}

// رسم الجدول
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
    th.innerHTML = `${col} <span class="sort-arrow text-gray-300"></span>`;
    th.addEventListener("click", () => sortTableByColumn(index));
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);

  // صفوف الجدول
  data.forEach((row, i) => {
    const tr = document.createElement("tr");
    tr.className = i % 2 === 0 ? "bg-gray-100" : "bg-gray-200";

    columns.forEach(col => {
      let value = row[col];
      const td = document.createElement("td");
      td.className = "px-3 py-1 border-b border-gray-300 text-sm";

      if (!isNaN(parseFloat(value)) && value !== "") {
        value = parseFloat(value).toFixed(2);
        if (value < 0) td.classList.add("text-red-600");
      }

      td.textContent = value;
      tr.appendChild(td);
    });

    // السطر الأخير للمجاميع
    if (row[columns[0]] && row[columns[0]].toString().toLowerCase().includes("total")) {
      tr.className = "bg-gray-800 text-white font-bold";
    }

    tbody.appendChild(tr);
  });
}

// البحث الفوري
document.getElementById("searchInput").addEventListener("keyup", function () {
  const filter = this.value.toLowerCase();
  const rows = document.querySelectorAll("#dataTable tbody tr");
  rows.forEach(row => {
    row.style.display = row.textContent.toLowerCase().includes(filter) ? "" : "none";
  });
});

// فرز الأعمدة + المحافظة على آخر سطر
function sortTableByColumn(colIndex) {
  const tbody = document.querySelector("#dataTable tbody");
  const rows = Array.from(tbody.querySelectorAll("tr"));
  const lastRow = rows.pop();

  const ascending = tbody.dataset.sortOrder !== "asc";
  tbody.dataset.sortOrder = ascending ? "asc" : "desc";

  rows.sort((a, b) => {
    const aVal = a.children[colIndex].textContent.trim();
    const bVal = b.children[colIndex].textContent.trim();
    const aNum = parseFloat(aVal);
    const bNum = parseFloat(bVal);
    if (!isNaN(aNum) && !isNaN(bNum)) return ascending ? aNum - bNum : bNum - aNum;
    return ascending ? aVal.localeCompare(bVal) : bVal.localeCompare(aVal);
  });

  rows.forEach((row, i) => {
    row.className = i % 2 === 0 ? "bg-gray-100" : "bg-gray-200";
    tbody.appendChild(row);
  });

  tbody.appendChild(lastRow);
}

// زر تحديث البيانات
document.getElementById("refresh-btn").addEventListener("click", loadExcelData);

// تحميل أول مرة
loadExcelData();
