// script.js — عرض ملف Excel بخصائص تنسيق متقدمة
document.addEventListener('DOMContentLoaded', () => {
  const input = document.getElementById('fileInput');
  const tableContainer = document.getElementById('tableContainer');

  input.addEventListener('change', (event) => {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });

      console.log("📄 Sheets found:", workbook.SheetNames);
      const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
      const jsonData = XLSX.utils.sheet_to_json(firstSheet, { defval: "" });

      if (jsonData.length === 0) {
        tableContainer.innerHTML = "<p class='text-center text-gray-500'>الملف فارغ</p>";
        return;
      }

      const headers = Object.keys(jsonData[0]);
      let html = `<table class="min-w-full border-collapse border border-gray-300 text-sm">
                    <thead class="bg-gray-200">
                      <tr>${headers.map(h => `<th class="border border-gray-300 px-3 py-2">${h}</th>`).join('')}</tr>
                    </thead>
                    <tbody>`;

      jsonData.forEach(row => {
        html += "<tr>";
        headers.forEach(h => {
          let value = row[h];

          // تنسيق القيم الرقمية
          if (!isNaN(parseFloat(value)) && value !== "") {
            value = parseFloat(value).toFixed(2);

            // أعمدة النسب المئوية
            if (h.includes("%") || h.toLowerCase().includes("percent")) {
              value += "%";
            }

            // تلوين القيم السالبة أو الموجبة
            const numeric = parseFloat(value);
            const colorClass = numeric < 0 ? "text-red-600" : numeric > 0 ? "text-green-600" : "text-gray-800";
            html += `<td class="border border-gray-300 px-3 py-1 ${colorClass} text-right">${value}</td>`;
          } else {
            // نصوص أو قيم غير رقمية
            html += `<td class="border border-gray-300 px-3 py-1">${value}</td>`;
          }
        });
        html += "</tr>";
      });

      html += "</tbody></table>";
      tableContainer.innerHTML = html;
    };

    reader.readAsArrayBuffer(file);
  });
});
