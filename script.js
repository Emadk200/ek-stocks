document.addEventListener("DOMContentLoaded", () => {
  const fileInput = document.getElementById("fileInput");
  const tableContainer = document.getElementById("tableContainer");
  const searchInput = document.getElementById("searchInput");
  const dateTimeElement = document.getElementById("dateTime");
  const customMessage = document.getElementById("customMessage");
  const messageInput = document.getElementById("messageInput");
  const editMessageBtn = document.getElementById("editMessageBtn");

  let currentData = [];
  let sortState = {};

  // ✅ تحديث التاريخ الميلادي
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
      timeZone: "Asia/Riyadh", // ميلادي بتوقيت السعودية
    };
    dateTimeElement.textContent = now.toLocaleString("en-GB", options);
  }

  updateDateTime();
  setInterval(updateDateTime, 60000);

  // ✅ الرسالة المخصصة القابلة للتعديل
  customMessage.textContent = "📅 آخر تحديث للأسعار محسوب حسب ملف التحليل المرفوع.";
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

  // ✅ رفع ملف Excel
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
      currentData = jsonData;
      renderTable(jsonData);
    };
    reader.readAsArrayBuffer(file);
  });

  // ✅ فلترة البحث
  searchInput.addEventListener("input", () => {
    const query = searchInput.value.toLowerCase();
    const filtered = currentData.filter((row) =>
      Object.values(row).some((val) =>
        String(val).toLowerCase().includes(query)
      )
    );
    renderTable(filtered);
  });

  // ✅ عرض الجدول + الفرز + تنسيق الأرقام
  function renderTable(data) {
    if (!data.length) {
      tableContainer.innerHTML =
        "<p class='text-center text-gray-500'>الملف فارغ أو لا يحتوي على بيانات.</p>";
      return;
    }

    const table = document.createElement("table");
    table.className =
      "min-w-full border border-gray-300 divide-y divide-gray-200 text-sm";

    // رؤوس الأعمدة
    const thead = document.createElement("thead");
    thead.className = "bg-blue-100";
    const headerRow = document.createElement("tr");

    Object.keys(data[0]).forEach((key) => {
      const th = document.createElement("th");
      th.textContent = key;
      th.className =
        "px-4 py-2 text-right font-semibold text-gray-800 cursor-pointer hover:bg-blue-200";
      th.addEventListener("click", () => sortByColumn(key));
      headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);

    // بيانات الجدول
    const tbody = document.createElement("tbody");
    data.forEach((row) => {
      const tr = document.createElement("tr");
      tr.className = "hover:bg-gray-50";

      Object.values(row).forEach((value) => {
        const td = document.createElement("td");
        td.className = "px-4 py-2";

        if (!isNaN(value) && value !== "") {
          let num = parseFloat(value);
          const isPercentage =
            /%|نسبة|Dev|P\/E|P\/B|Return|Change|Growth/i.test(td.textContent);

          let formatted = num.toFixed(2);
          if (isPercentage) formatted += "%";
          td.textContent = formatted;

          if (num < 0) td.classList.add("text-red-500");
        } else {
          td.textContent = value;
        }

        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });

    table.appendChild(thead);
    table.appendChild(tbody);
    tableContainer.innerHTML = "";
    tableContainer.appendChild(table);
  }

  // ✅ وظيفة الفرز
  function sortByColumn(column) {
    const direction = sortState[column] === "asc" ? "desc" : "asc";
    sortState[column] = direction;

    currentData.sort((a, b) => {
      if (typeof a[column] === "number" && typeof b[column] === "number") {
        return direction === "asc" ? a[column] - b[column] : b[column] - a[column];
      } else {
        return direction === "asc"
          ? String(a[column]).localeCompare(String(b[column]))
          : String(b[column]).localeCompare(String(a[column]));
      }
    });

    renderTable(currentData);
  }
});
