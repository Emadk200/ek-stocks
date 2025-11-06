document.addEventListener("DOMContentLoaded", () => {
  const fileInput = document.getElementById("fileInput");
  const tableContainer = document.getElementById("tableContainer");
  const searchInput = document.getElementById("searchInput");
  const dateTimeElement = document.getElementById("dateTime");
  const customMessage = document.getElementById("customMessage");
  const messageInput = document.getElementById("messageInput");
  const editMessageBtn = document.getElementById("editMessageBtn");

  let currentData = [];
  let summaryRow = null;
  let headers = [];
  let sortState = {};

  // ✅ عرض التاريخ بالميلادي (توقيت الرياض)
  function updateDateTime() {
    const now = new Date();
    const options = {
      weekday: "long",
      year: "numeric",
      month: "long",
      day: "numeric",
      hour: "2-digit",
      minute: "2-digit",
      hour12: true,
      timeZone: "Asia/Riyadh",
    };
    dateTimeElement.textContent = now.toLocaleString("en-GB", options);
  }
  updateDateTime();
  setInterval(updateDateTime, 60000);

  // ✅ رسالة قابلة للتعديل
  customMessage.textContent = "📅 آخر تحديث للأسعار محسوب من ملف التحليل المرفوع.";
  editMessageBtn.addEventListener("click", () => {
    messageInput.value = customMessage.textContent;
    messageInput.style.display = "inline-block";
    customMessage.style.display = "none";
    messageInput.focus();
  });
  messageInput.addEventListener("blur", () => {
    customMessage.textContent = messageInput.value || customMessage.textContent;
    messageInput.style.display = "none";
    customMessage.style.display = "inline";
  });

  // ✅ تحميل ملف Excel
  fileInput.addEventListener("change", (event) => {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(sheet, { defval: "" });

      if (jsonData.length > 1) {
        summaryRow = jsonData[jsonData.length - 1];
        currentData = jsonData.slice(0, -1);
      } else {
        currentData = jsonData;
        summaryRow = null;
      }

      headers = Object.keys(jsonData[0]);
      renderTable(currentData);
    };
    reader.readAsArrayBuffer(file);
  });

  // ✅ البحث
  searchInput.addEventListener("input", () => {
    const query = searchInput.value.toLowerCase();
    const filtered = currentData.filter((row) =>
      Object.values(row).some((val) =>
        String(val).toLowerCase().includes(query)
      )
    );
    renderTable(filtered);
  });

  // ✅ رسم الجدول مع السهم (▲▼)
  function renderTable(data) {
    if (!data.length) {
      tableContainer.innerHTML =
        "<p class='text-center text-gray-500'>الملف فارغ أو لا يحتوي بيانات.</p>";
      return;
    }

    const table = document.createElement("table");
    table.className =
      "min-w-full border border-gray-300 divide-y divide-gray-200 text-sm shadow-md rounded-xl overflow-hidden";

    // 🔹 الرأس
    const thead = document.createElement("thead");
    thead.className = "bg-blue-100 select-none";
    const headerRow = document.createElement("tr");

    headers.forEach((key) => {
      const th = document.createElement("th");
      th.className =
        "px-4 py-2 text-right font-semibold text-gray-800 cursor-pointer hover:bg-blue-200";

      const span = document.createElement("span");
      span.textContent = key;
      const arrow = document.createElement("span");
      arrow.className = "ml-1 text-gray-500";
      if (sortState[key] === "asc") arrow.textContent = "▲";
      else if (sortState[key] === "desc") arrow.textContent = "▼";
      else arrow.textContent = "";

      th.appendChild(span);
      th.appendChild(arrow);

      th.addEventListener("click", () => sortByColumn(key));
      headerRow.appendChild(th);
    });

    thead.appendChild(headerRow);

    // 🔹 الجسم
    const tbody = document.createElement("tbody");
    data.forEach((row) => {
      const tr = document.createElement("tr");
      tr.className = "hover:bg-gray-50";
      headers.forEach((key) => {
        const td = document.createElement("td");
        td.className = "px-4 py-2";
        const value = row[key];
        const isPercentageColumn = /%|Dev|Change|Return|Growth/i.test(key);

        if (!isNaN(value) && value !== "") {
          const num = parseFloat(value);
          let formatted = num.toFixed(2);
          if (isPercentageColumn) formatted += "%";
          td.textContent = formatted;
          if (num < 0) td.classList.add("text-red-500");
        } else td.textContent = value;

        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });

    // 🔹 صف المجاميع
    if (summaryRow) {
      const summaryTr = document.createElement("tr");
      summaryTr.className = "bg-gray-100 font-semibold";
      headers.forEach((key) => {
        const td = document.createElement("td");
        td.className = "px-4 py-2";
        const value = summaryRow[key];
        const isPercentageColumn = /%|Dev|Change|Return|Growth/i.test(key);

        if (!isNaN(value) && value !== "") {
          const num = parseFloat(value);
          let formatted = num.toFixed(2);
          if (isPercentageColumn) formatted += "%";
          td.textContent = formatted;
        } else td.textContent = value;

        summaryTr.appendChild(td);
      });
      tbody.appendChild(summaryTr);
    }

    table.appendChild(thead);
    table.appendChild(tbody);
    tableContainer.innerHTML = "";
    tableContainer.appendChild(table);
  }

  // ✅ وظيفة الفرز مع الأسهم واستثناء الصف الأخير
  function sortByColumn(column) {
    const direction = sortState[column] === "asc" ? "desc" : "asc";
    sortState = { [column]: direction }; // إعادة تعيين حالة الأسهم
    const sorted = [...currentData].sort((a, b) => {
      const aVal = a[column];
      const bVal = b[column];
      const aNum = parseFloat(aVal);
      const bNum = parseFloat(bVal);

      if (!isNaN(aNum) && !isNaN(bNum))
        return direction === "asc" ? aNum - bNum : bNum - aNum;
      return direction === "asc"
        ? String(aVal).localeCompare(String(bVal))
        : String(bVal).localeCompare(String(aVal));
    });

    renderTable(sorted);
  }
});
