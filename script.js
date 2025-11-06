// script.js — Excel Viewer with sorting and formatting
document.addEventListener('DOMContentLoaded', () => {
  const input = document.getElementById('fileInput');
  const tableContainer = document.getElementById('tableContainer');

  let jsonData = [];
  let headers = [];
  let sortState = {}; // لتتبع اتجاه الفرز لكل عمود

  input.addEventListener('change', (event) => {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });

      console.log("📄 Sheets found:", workbook.SheetNames);
      const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
      jsonData = XLSX.utils.sheet_to_json(firstSheet, { defval: "" });

      if (jsonData.length === 0) {
        tableContainer.innerHTML = "<p class='text-center text-gray-500'>الملف فارغ</p>";
        return;
      }

      headers = Object.keys(jsonData[0]);
      renderTable(jsonData);
    };

    reader.readAsArrayBuffer(file);
  });

  function renderTable(data) {
    let html = `<table class="min-w-full border-collapse border border-gray-300 text-sm">
                  <thead class="bg-gray-200">
                    <tr>
                      ${headers.map(h => `
                        <th class="border border-gray-300 px-3 py-2 cursor-pointer select-none hover:bg-gray-300" 
                            onclick="sortTable('${h}')">${h} ⬍</th>`).join('')}
                    </tr>
                  </thead>
                  <tbody>
                    ${data.map(row => {
                      return `<tr>${headers.map(h => {
                        let value = row[h];
                        if (!isNaN(parseFloat(value)) && value !== "") {
                          value = parseFloat(value).toFixed(2);

                          // إضافة علامة %
                          if (h.includes("%") || h.toLowerCase().includes("percent")) {
                            value += "%";
                          }

                          const numeric = parseFloat(value);
                          const colorClass =
                            numeric < 0 ? "text-red-600" :
                            numeric > 0 ? "text-green-600" :
                            "text-gray-800";
                          return `<td class="border border-gray-300 px-3 py-1 text-right ${colorClass}">${value}</td>`;
                        } else {
                          return `<td class="border border-gray-300 px-3 py-1">${value}</td>`;
                        }
                      }).join('')}</tr>`;
                    }).join('')}
                  </tbody>
                </table>`;

    tableContainer.innerHTML = html;
  }

  // ✅ دالة الفرز
  window.sortTable = (colName) => {
    const direction = sortState[colName] === "asc" ? "desc" : "asc";
    sortState[colName] = direction;

    jsonData.sort((a, b) => {
      const valA = parseFloat(a[colName]) || a[colName];
      const valB = parseFloat(b[colName]) || b[colName];

      if (typeof valA === "number" && typeof valB === "number") {
        return direction === "asc" ? valA - valB : valB - valA;
      } else {
        return direction === "asc"
          ? String(valA).localeCompare(String(valB))
          : String(valB).localeCompare(String(valA));
      }
    });

    renderTable(jsonData);
  };
});
