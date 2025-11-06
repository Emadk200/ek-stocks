document.addEventListener("DOMContentLoaded", () => {
  const fileInput = document.getElementById("fileInput");
  const tableContainer = document.getElementById("tableContainer");
  const dateTimeElement = document.getElementById("dateTime");
  const customMessage = document.getElementById("customMessage");

  // ✅ تحديث التاريخ والوقت
  function updateDateTime() {
    const now = new Date();
    const options = {
      weekday: "long",
      year: "numeric",
      month: "long",
      day: "numeric",
      hour: "2-digit",
      minute: "2-digit",
    };
    dateTimeElement.textContent = now.toLocaleString("ar-SA", options);
  }

  updateDateTime();
  setInterval(updateDateTime, 60000);

  // ✅ الرسالة المخصصة
  customMessage.textContent = "📅 آخر تحديث للأسعار محسوب حسب ملف التحليل المرفوع.";

  // ✅ عند اختيار الملف
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

      renderTable(jsonData);
    };
    reader.readAsArrayBuffer(file);
  });

  // ✅ دالة عرض الجدول
  function renderTable(data) {
    if (!data.length) {
      tableContainer.innerHTML =
        "<p class='text-center text-gray-500'>الملف فارغ أو لا يحتوي على بيانات.</p>";
      return;
    }

    const table = document.createElement("table");
    table.className = "min-w-full border border-gray-300 divide-y divide-gray-200";

    const thead = document.createElement("thead");
    thead.className = "bg-blue-100";
    const headerRow = document.createElement("tr");

    Object.keys(data[0]).forEach((key) => {
      const th = document.createElement("th");
      th.textContent = key;
      th.className = "px-4 py-2 text-right text-sm font-semibold text-gray-800";
      headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);

    const tbody = document.createElement("tbody");

    data.forEach((row) => {
      const tr = document.createElement("tr");
      tr.className = "hover:bg-gray-50";

      Object.values(row).forEach((value) => {
        const td = document.createElement("td");
        td.className = "px-4 py-2 text-sm";

        if (typeof value === "number") {
          const isPercentage =
            /%|نسبة|Dev|P\/E|P\/B|Return|Change/i.test(value) ||
            /%|نسبة/i.test(td.textContent);

          let formatted = value.toFixed(2);
          if (isPercentage) formatted += "%";

          td.textContent = formatted;

          if (value < 0) td.classList.add("text-red-500");
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
});
