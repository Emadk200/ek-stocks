// تحميل ملف CSV تلقائيًا من المجلد
fetch("EKstoks.csv")
  .then(response => response.text())
  .then(csvText => {
    const data = Papa.parse(csvText, { header: true }).data;
    renderTable(data);
  });

function renderTable(data) {
  const tableHead = document.getElementById("table-head");
  const tableBody = document.getElementById("table-body");

  // توليد رؤوس الأعمدة
  const headers = Object.keys(data[0]);
  const extraHeaders = ["Dev_calc", "Dev%_calc", "P/E_calc", "P/BV_calc", "PE*PBV_calc", "LP_GN_calc"];
  tableHead.innerHTML = "<tr>" + headers.concat(extraHeaders).map(h => `<th>${h}</th>`).join("") + "</tr>";

  // تعبئة الصفوف
  data.forEach(row => {
    // تحويل النصوص إلى أرقام عند الحاجة
    const num = val => parseFloat(val) || 0;

    const Dev_calc = num(row["EPS 25/Q3"]) - num(row["EPS 24/Q3"]);
    const Dev_pct = (num(Dev_calc) / num(row["EPS 24/Q3"])) || 0;
    const PE_calc = num(row["Current Price"]) / num(row["EPS 2025"]);
    const PBV_calc = num(row["Current Price"]) / num(row["P/B.V"]);
    const PE_PBV_calc = PE_calc * PBV_calc;
    const LP_GN_calc = num(row["Current Price"]) / num(row["L.P/G.N"]);

    const allData = {
      ...row,
      Dev_calc: Dev_calc.toFixed(3),
      Dev%_calc: Dev_pct.toFixed(3),
      "P/E_calc": PE_calc.toFixed(3),
      "P/BV_calc": PBV_calc.toFixed(3),
      "PE*PBV_calc": PE_PBV_calc.toFixed(3),
      "LP_GN_calc": LP_GN_calc.toFixed(3)
    };

    const rowHTML = "<tr>" + Object.values(allData).map(v => `<td>${v}</td>`).join("") + "</tr>";
    tableBody.innerHTML += rowHTML;
  });
}
