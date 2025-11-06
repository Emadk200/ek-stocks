document.addEventListener("DOMContentLoaded", () => {
  const githubFileURL =
    "https://emadk200.github.io/ek-stocks/watchliststoks.xlsx"; // المسار إلى ملف Excel

  const dateElem = document.getElementById("datetime");
  const table = document.getElementById("stockTable");

  // عرض التاريخ الميلادي الحالي
  const now = new Date();
  const formattedDate = now.toLocaleString("en-GB", {
    dateStyle: "full",
    timeStyle: "short",
  });
  dateElem.textContent = `Last updated: ${formattedDate}`;

  // تحميل ملف Excel من GitHub
  fetch(githubFileURL)
    .then((res) => res.arrayBuffer())
    .then((data) => {
      const workbook = XLSX.read(data, { type: "array" });
      const sheetName = workbook.SheetNames[0];
      const sheet = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: "" });

      if (sheet.length === 0) return;

      const columns = Object.keys(sheet[0]);

      // إنشاء رأس الجدول
      const thead = table.querySelector("thead");
      const headerRow = document.createElement("tr");
      columns.forEach((col) => {
        const th = document.createElement("th");
        th.textContent = col;
        th.className =
          "px-3 py-2 text-left font-semibold cursor-pointer select-none";
        th.addEventListener("click", () => sortTable(col));
        headerRow.appendChild(th);
      });
      thead.appendChild(headerRow);

      // إنشاء الصفوف
      const tbody = table.querySelector("tbody");
      sheet.forEach((rowData, rowIndex) => {
        const row = document.createElement("tr");

        // تلوين الصفوف بالتناوب
        if (rowIndex === sheet.length - 1) {
          row.className = "bg-gray-800 text-white font-semibold"; // السطر الأخير (مجاميع)
        } else {
          row.className = rowIndex % 2 === 0 ? "bg-white" : "bg-gray-300";
        }

        columns.forEach((col) => {
          const td = document.createElement("td");
          let value = rowData[col];

          // تنسيق الأرقام
          if (typeof value === "number") {
            if (col.includes("%")) {
              value = `${value.toFixed(2)}%`;
            } else {
              value = value.toFixed(2);
            }
          }

          // تلوين القيم السالبة بالأحمر
          if (parseFloat(rowData[col]) < 0) {
            td.classList.add("text-red-600");
          }

          td.classList.add("px-3", "py-2", "text-sm", "border-t", "border-gray-200");
          td.textContent = value;
          row.appendChild(td);
        });
        tbody.appendChild(row);
      });

      // وظيفة الفرز
      function sortTable(column) {
        const colIndex = columns.indexOf(column);
        const isNumeric = sheet.every((row) => !isNaN(parseFloat(row[column])));

        // فصل السطر الأخير (المجاميع)
        const lastRow = sheet.pop();

        sheet.sort((a, b) => {
          const valA = a[column];
          const valB = b[column];
          if (isNumeric) return valA - valB;
          return valA.toString().localeCompare(valB.toString());
        });

        // إعادة السطر الأخير
        sheet.push(lastRow);

        // إعادة رسم الجدول بعد الفرز
        tbody.innerHTML = "";
        sheet.forEach((rowData, rowIndex) => {
          const row = document.createElement("tr");
          if (rowIndex === sheet.length - 1) {
            row.className = "bg-gray-800 text-white font-semibold";
          } else {
            row.className = rowIndex % 2 === 0 ? "bg-white" : "bg-gray-300";
          }
          columns.forEach((col) => {
            const td = document.createElement("td");
            let value = rowData[col];
            if (typeof value === "number") {
              if (col.includes("%")) {
                value = `${value.toFixed(2)}%`;
              } else {
                value = value.toFixed(2);
              }
            }
            if (parseFloat(rowData[col]) < 0) {
              td.classList.add("text-red-600");
            }
            td.classList.add("px-3", "py-2", "text-sm", "border-t", "border-gray-200");
            td.textContent = value;
            row.appendChild(td);
          });
          tbody.appendChild(row);
        });
      }
    })
    .catch((err) => {
      console.error("Error reading Excel file:", err);
      alert("حدث خطأ أثناء قراءة الملف. يرجى التأكد من المسار.");
    });
});
