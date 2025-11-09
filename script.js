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

// الملاحظة (تستطيع تعديل النص داخل هذا الحقل)
document.getElementById("notes").innerHTML = `
- This dashboard updates live prices from ASE website.
- Last Price is fetched automatically when pressing update.
- Other values are taken from your Excel sheet.
`;

// جلب الأسعار من موقع البورصة
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
        const symbol = cells[6]?.innerText.trim(); // Symbol Index
        const lastClosing = cells[7]?.innerText.trim(); // Last Closing Price Index

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

// تحميل sheet + دمج الأسعار + رسم الجدول
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

  } catch (error) {
    alert("حدث خطأ أثناء تحميل البيانات");
    console.error(error);
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

  const headerRow = document.createElement("tr");
  columns.forEach((col, index) => {
    const th = document.createElement("th");
    th.className = "px-3 py-2 bg-gray-800 text-white font-semibold cursor-pointer";
    th.innerHTML = `${col}`;
    th.addEventListener("click", () => sortTableByColumn(index));
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);

	data.forEach((row, i) => {
	  const tr = document.createElement("tr");

	  const isTotalRow =
	    row[Object.keys(row)[0]] &&
	    row[Object.keys(row)[0]].toString().toLowerCase().includes("total");

	  // صف المجاميع ثابت ولونه مميز
	  if (isTotalRow) {
	    tr.className = "bg-gray-800 text-white font-bold";
	  } else {
	    tr.className = i % 2 === 0 ? "bg-gray-100" : "bg-gray-200";
	  }

    columns.forEach(col => {
	const rawChange = parseFloat(row["Change"]);// ← نأخذ القيمة قبل التنسيق

      let value = row[col];
      const td = document.createElement("td");
      td.className = "px-3 py-1 border-b border-gray-300 text-sm";

      // لا ننسّق عمود ASE Code
      if (col !== "ASE Code" && !isNaN(parseFloat(value)) && value !== "") {
        value = parseFloat(value).toFixed(2);

        // السالب أحمر
	// تلوين خاص لعمودي Change و Change %
	if (col === "Change" || col === "Change %") {

	  const num = parseFloat(value);

	  if (!isNaN(num)) {
	    if (num > 0) {
	      td.classList.add("text-green-600", "font-semibold"); // أخضر
	    } else if (num < 0) {
	      td.classList.add("text-red-600", "font-semibold"); // أحمر
	    } else {
	      td.classList.add("text-gray-800"); // أسود للقيمة صفر
	    }
	  }
	}

        // العمود يحتوي نسبة مئوية → أضف % + نسبة مئوية → اضرب في 100 قبل الإضافة
	// نسبة مئوية → اضرب ×100 وعرض بدون كسور
	if (col.includes("%") || col.toLowerCase().includes("pct")) {
	  value = Math.round(parseFloat(value) * 100) + "%";
	}
      }

	// سهم الاتجاه في عمود Last Price بناءً على Change الخام
	if (col === "Last Price" && !isNaN(rawChange)) {
	  if (rawChange > 0) {
	    value = value + " ↑";
	    td.classList.add("text-green-600", "font-semibold");
	  } else if (rawChange < 0) {
	    value = value + " ↓";
	    td.classList.add("text-red-600", "font-semibold");
	  } else {
	    value = value + " -";
	    td.classList.add("text-gray-700");
	  }
	}
	}

	// ضع القيمة داخل الـ <td>
	td.textContent = value;
      tr.appendChild(td);
    });

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

// فرز الأعمدة
function sortTableByColumn(colIndex) {
  const tbody = document.querySelector("#dataTable tbody");
  const rows = Array.from(tbody.querySelectorAll("tr"));

  // ✅ افصل سطر المجاميع
  const totalRow = rows.pop();

  const ascending = tbody.dataset.sortOrder !== "asc";
  tbody.dataset.sortOrder = ascending ? "asc" : "desc";

  rows.sort((a, b) => {
    const aVal = a.children[colIndex].textContent.replace("%", "").trim();
    const bVal = b.children[colIndex].textContent.replace("%", "").trim();

    const aNum = parseFloat(aVal);
    const bNum = parseFloat(bVal);

    if (!isNaN(aNum) && !isNaN(bNum)) {
      return ascending ? aNum - bNum : bNum - aNum;
    }
    return ascending ? aVal.localeCompare(bVal) : bVal.localeCompare(aVal);
  });

  // ✅ أعد zebra striping بدون التأثير على سطر المجاميع
  rows.forEach((row, i) => {
    row.className = i % 2 === 0 ? "bg-gray-100" : "bg-gray-200";
    tbody.appendChild(row);
  });

  // ✅ أرجع سطر المجاميع في النهاية دائمًا
  tbody.appendChild(totalRow);
}

// زر التحديث
document.getElementById("refresh-btn").addEventListener("click", loadExcelData);

// تشغيل عند فتح الصفحة
loadExcelData();
