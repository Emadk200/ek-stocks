const fileUrl = "https://emadk200.github.io/ek-stocks/watchliststoks.xlsx?v=" + Date.now();

// تحميل البيانات من GitHub
async function loadExcelData() {
  const loading = document.getElementById("loading-indicator");
  try {
    loading.classList.remove("hidden");
    const response = await fetch(fileUrl);
    const arrayBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: "array" });

    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    // ✅ هنا التغيير المهم
    const data = XLSX.utils.sheet_to_json(sheet, { defval: "" });

    renderTable(data);
  } catch (error) {
    console.error(error);
    alert("حدث خطأ أثناء تحميل البيانات من GitHub!");
  } finally {
    loading.classList.add("hidden");
  }
}

// عرض التاريخ الميلادي
document.getElementById("current-date").textContent =
  new Date().toLocaleString("en-US", {
    weekday: "long",
    year: "numeric",
    month: "long",
    day: "numeric",
    hour: "2-digit",
    minute: "2-digit",
  });

// زر التحديث
document.getElementById("refresh-btn").addEventListener("click", loadExcelData);

// البحث الفوري
document.getElementById("searchInput").addEventListener("keyup", function () {
  const filter = this.value.toLowerCase();
  const rows = document.querySelectorAll("#dataTable tbody tr");
  rows.forEach(row => {
    row.style.display = row.textContent.toLowerCase().includes(filter) ? "" : "none";
  });
});

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

  // صفوف
  data.forEach((row, i) => {
    const tr = document.createElement("tr");
    tr.className = i % 2 === 0 ? "bg-gray-100" : "bg-gray-200";

    columns.forEach(col => {
      let value = row[col];
      const td = document.createElement("td");
      td.className = "px-3 py-1 border-b border-gray-300";

      if (!isNaN(parseFloat(value)) && value !== "") {
        value = parseFloat(value).toFixed(2);
        if (value < 0) td.classList.add("text-red-600");
      }

      td.textContent = value;
      tr.appendChild(td);
    });

    // آخر سطر = مجاميع
    if (row[columns[0]] && row[columns[0]].toString().toLowerCase().includes("total")) {
      tr.className = "bg-gray-800 text-white font-semibold";
    }

    tbody.appendChild(tr);
  });
}

// فرز الأعمدة + الحفاظ على صف المجاميع
function sortTableByColumn(colIndex) {
  const tbody = document.querySelector("#dataTable tbody");
  const rows = Array.from(tbody.querySelectorAll("tr"));
  const lastRow = rows.pop();

  const ascending = tbody.dataset.sortOrder !== "asc";
  tbody.dataset.sortOrder = ascending ? "asc" : "desc";

  rows.sort((a, b) => {
    const aVal = a.children[colIndex].textContent.replace("%", "").trim();
    const bVal = b.children[colIndex].textContent.replace("%", "").trim();
    const aNum = parseFloat(aVal);
    const bNum = parseFloat(bVal);
    if (!isNaN(aNum) && !isNaN(bNum))
      return ascending ? aNum - bNum : bNum - aNum;
    return ascending ? aVal.localeCompare(bVal) : bVal.localeCompare(aVal);
  });

  rows.forEach((row, i) => {
    row.className = i % 2 === 0 ? "bg-gray-100" : "bg-gray-200";
    tbody.appendChild(row);
  });

  tbody.appendChild(lastRow);
}

// تشغيل أول تحميل
loadExcelData();
