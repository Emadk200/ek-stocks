// ====================== إعدادات ======================

// اسم ملف الإكسل (يجب أن يكون في نفس مجلد index.html)
const EXCEL_FILE_URL = "watchliststocks.xlsx";

// عنصر الحالة (Loading…)
const loadingEl = document.getElementById("loadingStatus");

// ====================== تحميل ملف الإكسل ======================

async function loadExcelData() {
  loadingEl.style.display = "block";

let response;
try {
  response = await fetch(EXCEL_FILE_URL);
  if (!response.ok) throw new Error("HTTP " + response.status);
} catch(e) {
  console.error("❌ Error loading Excel file:", e);
  return;
}

  const arrayBuffer = await response.arrayBuffer();
  const workbook = XLSX.read(arrayBuffer, { type: "array" });

  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];

  // ✅ إجبار المكتبة على عدم إسقاط آخر الصفوف
  const range = XLSX.utils.decode_range(sheet['!ref']);
  range.e.r = range.e.r + 3;
  sheet['!ref'] = XLSX.utils.encode_range(range);

  // نقرأ القيم بصيغة مصفوفات
  const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

  // أول صف هو رؤوس الأعمدة
  const headers = raw.shift();

  // تحويل الصفوف → كائنات
  const data = raw.map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      obj[h] = row[i];
    });
    return obj;
  });

  window.allData = data; // للحفظ + البحث + فرز
  renderTable(data);
  loadingEl.style.display = "none";
}

// ====================== عرض الجدول ======================

function renderTable(data) {
  const table = document.getElementById("dataTable");
  const thead = table.querySelector("thead");
  const tbody = table.querySelector("tbody");

  thead.innerHTML = "";
  tbody.innerHTML = "";

  const columns = Object.keys(data[0]);

  // رأس الجدول + الفرز
  const trHead = document.createElement("tr");
  columns.forEach((col, idx) => {
    const th = document.createElement("th");
    th.textContent = col;
    th.className = "px-3 py-2 bg-gray-800 text-white font-semibold cursor-pointer";
    th.onclick = () => sortTableByColumn(idx);
    trHead.appendChild(th);
  });
  thead.appendChild(trHead);

  // ✅ فصل سطر المجاميع
  const totalRow = data.find(r =>
    Object.values(r).join("").replace(/[^a-z0-9]/gi, "").toLowerCase().includes("tot")
  );
  const rows = data.filter(r => r !== totalRow);

  rows.forEach((row, i) => {
    const tr = document.createElement("tr");
    tr.className = i % 2 === 0 ? "bg-gray-100" : "bg-gray-200";

    columns.forEach(col => {
      let val = row[col];
      const td = document.createElement("td");
      td.className = "px-3 py-1 border-b border-gray-300 text-sm";

      // نسب مئوية → بدون كسور + %
      if (col.toLowerCase().includes("%")) {
        const num = parseFloat(val);
        if (!isNaN(num)) val = Math.round(num * 100) + "%";
      }
      else if (!isNaN(parseFloat(val)) && val !== "") {
        if (col.toLowerCase().includes("code")) val = parseInt(val);
        else val = parseFloat(val).toFixed(2);
      }

      // سهم آخر سعر
      if (col === "Last  Price" && !isNaN(parseFloat(row["Change"]))) {
        const c = parseFloat(row["Change"]);
        const arrow = c > 0 ? "↑" : c < 0 ? "↓" : "-";
        const color = c > 0 ? "text-green-600" : c < 0 ? "text-red-600" : "text-gray-600";
        td.innerHTML = `${val} <span class="${color} ml-1">${arrow}</span>`;
      } else {
        td.textContent = val;
      }

      // تلوين Change
      if (col === "Change") {
        const num = parseFloat(val);
        if (num > 0) td.classList.add("text-green-600");
        else if (num < 0) td.classList.add("text-red-600");
      }

      tr.appendChild(td);
    });

    tbody.appendChild(tr);
  });

  // ✅ سطر المجاميع آخر الجدول
 if (totalRow) {
  const tr = document.createElement("tr");
  tr.className = "bg-gray-700 text-white font-bold";

  columns.forEach(col => {
    let val = totalRow[col];
    const td = document.createElement("td");
    td.className = "px-3 py-1 border-b border-gray-500 text-sm";

    // إذا العمود نسبة مئوية
    if (col.toLowerCase().includes("%")) {
      const num = parseFloat(val);
      if (!isNaN(num)) val = Math.round(num * 100) + "%";
    }
    // إذا العمود رقم عادي
    else if (!isNaN(parseFloat(val)) && val !== "") {
      val = parseFloat(val).toFixed(2);
    }

    td.textContent = val;
    tr.appendChild(td);
  });

  tbody.appendChild(tr);
}

}

// ====================== البحث ======================

document.getElementById("searchInput").addEventListener("input", function () {
  const q = this.value.toLowerCase();
  const filtered = window.allData.filter(r =>
    Object.values(r).join(" ").toLowerCase().includes(q)
  );
  renderTable(filtered);
});

// ====================== الفرز ======================

function sortTableByColumn(index) {
  const key = Object.keys(window.allData[0])[index];
  window.allData.sort((a, b) => (a[key] > b[key] ? 1 : -1));
  renderTable(window.allData);
}

// ====================== زر التحديث ======================

document.getElementById("refreshBtn").addEventListener("click", loadExcelData);

// Start
loadExcelData();
