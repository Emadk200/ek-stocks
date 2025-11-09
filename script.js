const fileUrl = "https://raw.githubusercontent.com/Emadk200/ek-stocks/main/watchliststoks.xlsx";

// تحميل البيانات من GitHub
async function loadExcelData() {
  try {
    document.getElementById("loading").classList.remove("hidden");
    const response = await fetch(fileUrl);
    const arrayBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: "array" });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    renderTable(json);
  } catch (error) {
    console.error("خطأ في تحميل الملف:", error);
    alert("حدث خطأ أثناء تحميل البيانات من GitHub!");
  } finally {
    document.getElementById("loading").classList.add("hidden");
  }
}

async function loadExcelData() {
  const loadingDiv = document.getElementById("loading-indicator");
  try {
    loadingDiv.classList.remove("hidden"); // إظهار مؤشر التحميل
    document.getElementById("loading").classList.remove("hidden");

    const response = await fetch(fileUrl);
    const arrayBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: "array" });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    renderTable(json);
  } catch (error) {
    console.error("خطأ في تحميل الملف:", error);
    alert("حدث خطأ أثناء تحميل البيانات من GitHub!");
  } finally {
    document.getElementById("loading").classList.add("hidden");
    loadingDiv.classList.add("hidden"); // إخفاء مؤشر التحميل بعد انتهاء العملية
  }
}


// عرض التاريخ الميلادي الحالي
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
    const text = row.textContent.toLowerCase();
    row.style.display = text.includes(filter) ? "" : "none";
  });
});

// إنشاء الجدول بتنسيقات متكاملة
function renderTable(data) {
  const table = document.getElementById("dataTable");
  const thead = table.querySelector("thead");
  const tbody = table.querySelector("tbody");

  thead.innerHTML = "";
  tbody.innerHTML = "";

  // رأس الجدول مع أسهم الفرز
  const headerRow = document.createElement("tr");
  data[0].forEach((header, index) => {
    const th = document.createElement("th");
    th.className = "px-3 py-2 border-b border-gray-300 cursor-pointer select-none";
    th.innerHTML = `${header} <span class='sort-arrow text-gray-400'></span>`;
    th.addEventListener("click", () => sortTableByColumn(table, index));
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);

  // البيانات
  data.slice(1).forEach((row, i) => {
    const tr = document.createElement("tr");

    // zebra striping الثابت
    tr.className = i % 2 === 0 ? "bg-gray-100" : "bg-gray-200";

    row.forEach((cell, j) => {
      const td = document.createElement("td");
      td.className = "px-3 py-1 border-b border-gray-300";

      if (typeof cell === "number") {
        td.textContent = cell.toFixed(2);
        if (data[0][j].includes("%")) td.textContent += " %";
        if (cell < 0) td.classList.add("text-red-600");
      } else {
        td.textContent = cell;
      }

      tr.appendChild(td);
    });

    // صف المجاميع/المتوسط
    if (row[0] && row[0].toString().toLowerCase().includes("total")) {
      tr.className = "bg-gray-800 text-white font-semibold";
      const tds = tr.querySelectorAll("td");
      tds.forEach(td => (td.style.color = "#fff"));
    }

    tbody.appendChild(tr);
  });
}

// الفرز مع الأسهم واستثناء صف المجاميع
function sortTableByColumn(table, columnIndex) {
  const tbody = table.querySelector("tbody");
  const rows = Array.from(tbody.querySelectorAll("tr"));
  const totalRow = rows.find(r => r.classList.contains("bg-gray-800"));
  const dataRows = rows.filter(r => r !== totalRow);

  const headerCells = table.querySelectorAll("th");
  headerCells.forEach(th => th.querySelector(".sort-arrow").textContent = "");

  let isAscending = table.dataset.sortCol === columnIndex.toString() && table.dataset.sortDir === "asc" ? false : true;
  table.dataset.sortCol = columnIndex;
  table.dataset.sortDir = isAscending ? "asc" : "desc";

  const arrow = headerCells[columnIndex].querySelector(".sort-arrow");
  arrow.textContent = isAscending ? "▲" : "▼";

  dataRows.sort((a, b) => {
    const aText = a.children[columnIndex].textContent.replace("%", "").trim();
    const bText = b.children[columnIndex].textContent.replace("%", "").trim();
    const aNum = parseFloat(aText);
    const bNum = parseFloat(bText);
    if (!isNaN(aNum) && !isNaN(bNum))
      return isAscending ? aNum - bNum : bNum - aNum;
    return isAscending
      ? aText.localeCompare(bText)
      : bText.localeCompare(aText);
  });

  // إعادة ترتيب مع الحفاظ على ألوان zebra
  dataRows.forEach((row, i) => {
    row.className = i % 2 === 0 ? "bg-gray-100" : "bg-gray-200";
    tbody.appendChild(row);
  });

  if (totalRow) tbody.appendChild(totalRow);
}

// تحميل مبدئي
loadExcelData();
