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

  // الأعمدة الأصلية + المحسوبة
  const headers = Object.keys(data[0]);
  const extraHeaders = [
    "Dev_calc",
    "Dev%_calc",
    "P/E_calc",
    "P/BV_calc",
    "PE*PBV_calc",
    "LP_GN_calc"
  ];
  tableHead.innerHTML =
    "<tr>" +
    headers.concat(extraHeaders).map(h => `<th>${h}</th>`).join("") +
    "</tr>";

  data.forEach(row => {
    const num = val => parseFloat(val) || 0;

    // الحسابات
    const Dev_calc = num(row["EPS 25/Q3"]) - num(row["EPS 24/Q3"]);
    const Dev_pct = num(row["EPS 24/Q3"])
      ? (Dev_calc / num(row["EPS 24/Q3"])) * 100
      : 0;

    const PE_calc = num(row["EPS 2025"])
      ? num(row["Current Price"]) / num(row["EPS 2025"])
      : 0;

    const PBV_calc = num(row["P/B.V"])
      ? num(row["Current Price"]) / num(row["P/B.V"])
      : 0;

    const PE_PBV_calc = PE_calc * PBV_calc;

    const LP_GN_calc = num(row["L.P/G.N"])
      ? num(row["Current Price"]) / num(row["L.P/G.N"])
      : 0;

    const allData = {
      ...row,
      Dev_calc: Dev_calc.toFixed(3),
      Dev%_calc: Dev_pct.toFixed(2) + "%",
      "P/E_calc": PE_calc.toFixed(2),
      "P/BV_calc": PBV_calc.toFixed(2),
      "PE*PBV_calc": PE_PBV_calc.toFixed(2),
      "LP_GN_calc": LP_GN_calc.toFixed(2)
    };

    const rowHTML =
      "<tr>" +
      Object.values(allData)
        .map(v => `<td>${v}</td>`)
        .join("") +
      "</tr>";
    tableBody.innerHTML += rowHTML;
  });
}
