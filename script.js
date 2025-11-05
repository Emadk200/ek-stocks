// script.js — read Excel file and display as table
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

      // قراءة أول شيت فقط
      const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
      const jsonData = XLSX.utils.sheet_to_json(firstSheet, { defval: "" });

      if (jsonData.length === 0) {
        tableContainer.innerHTML = "<p class='text-center text-gray-500'>الملف فارغ</p>";
        return;
      }

      // إنشاء الجدول
      const headers = Object.keys(jsonData[0]);
      let html = `<table class="min-w-full border-collapse border border-gray-300 text-sm">
                    <thead class="bg-gray-200">
                      <tr>${headers.map(h => `<th class="border border-gray-300 px-3 py-2">${h}</th>`).join('')}</tr>
                    </thead>
                    <tbody>
                      ${jsonData.map(row => `
                        <tr>
                          ${headers.map(h => `<td class="border border-gray-300 px-3 py-1">${row[h]}</td>`).join('')}
                        </tr>
                      `).join('')}
                    </tbody>
                  </table>`;

      tableContainer.innerHTML = html;
    };

    reader.readAsArrayBuffer(file);
  });
});
