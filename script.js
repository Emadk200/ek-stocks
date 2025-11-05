// script.js — UPDATED: computes Excel formulas in the browser
// Requires PapaParse (already used in your site) and index.html to include it.

function computeRowComputedFields(row, headers) {
  // helper: get numeric value by Excel column letter (A..Z)
  const getVal = (colLetter) => {
    const idx = colLetter.charCodeAt(0) - 'A'.charCodeAt(0);
    const h = headers[idx] || null;
    const raw = h && row[h] !== undefined && row[h] !== '' ? String(row[h]).replace(/,/g,'').trim() : '';
    const v = raw === '' ? NaN : parseFloat(raw);
    return isNaN(v) ? NaN : v;
  };

  const out = {};

  // -- Generated from Excel formulas --
  // Column E (Dev.) : =D{r}-C{r}
  try {
    var val = getVal('D') - getVal('C');
    out['Dev.'] = isNaN(val) ? '' : val;
  } catch(e) { out['Dev.'] = ''; }

  // Column F (Dev. %) : =E{r}/C{r}
  try {
    var val = getVal('E') / getVal('C');
    out['Dev. %'] = isNaN(val) ? '' : val;
  } catch(e) { out['Dev. %'] = ''; }

  // Column J (P/E) : =B{r}/H{r}
  try {
    var val = getVal('B') / getVal('H');
    out['P/E'] = isNaN(val) ? '' : val;
  } catch(e) { out['P/E'] = ''; }

  // Column M (P/B.V) : =B{r}/L{r}
  try {
    var val = getVal('B') / getVal('L');
    out['P/B.V'] = isNaN(val) ? '' : val;
  } catch(e) { out['P/B.V'] = ''; }

  // Column O (P/E*P/B.V.) : =J{r}*M{r}
  try {
    var val = getVal('J') * getVal('M');
    out['P/E*P/B.V.'] = isNaN(val) ? '' : val;
  } catch(e) { out['P/E*P/B.V.'] = ''; }

  // Column Q (L.P/G.N) : =B{r}/P{r}
  try {
    var val = getVal('B') / getVal('P');
    out['L.P/G.N'] = isNaN(val) ? '' : val;
  } catch(e) { out['L.P/G.N'] = ''; }

  return out;
}


// --- Main: load CSV, compute, render table ---
fetch("EKstoks.csv")
  .then(resp => {
    if (!resp.ok) throw new Error("Failed to load EKstoks.csv: " + resp.status);
    return resp.text();
  })
  .then(csvText => {
    const parsed = Papa.parse(csvText, { header: true, skipEmptyLines: true });
    const data = parsed.data;
    if (!data || data.length === 0) {
      document.getElementById('table-body').innerHTML = '<tr><td colspan="20">No data found in EKstoks.csv</td></tr>';
      return;
    }

    // headers array (ordered) — we need it to map letters A,B,C... to headers
    const headers = Object.keys(data[0]);

    // compute additional fields for each row
    const computedNames = []; // collect names in order (use column headers from compute function)
    const rowsComputed = data.map(row => {
      const comp = computeRowComputedFields(row, headers);
      // collect computed keys (only once)
      for (const k of Object.keys(comp)) {
        if (!computedNames.includes(k)) computedNames.push(k);
      }
      return { original: row, computed: comp };
    });

    // render table header (original headers + computed headers)
    const tableHead = document.getElementById("table-head");
    const tableBody = document.getElementById("table-body");
    tableHead.innerHTML = "";
    tableBody.innerHTML = "";

    const allHeaders = headers.concat(computedNames);
    const trh = document.createElement('tr');
    allHeaders.forEach(h => {
      const th = document.createElement('th');
      th.textContent = h;
      trh.appendChild(th);
    });
    tableHead.appendChild(trh);

    // render rows
    rowsComputed.forEach(rc => {
      const tr = document.createElement('tr');
      // original columns
      headers.forEach(h => {
        const td = document.createElement('td');
        td.textContent = rc.original[h] !== undefined ? rc.original[h] : "";
        tr.appendChild(td);
      });
      // computed columns (format some nicely)
      computedNames.forEach(h => {
        const td = document.createElement('td');
        let v = rc.computed[h];
        // format percent-like Dev. % if header name includes %
        if (typeof v === 'number') {
          if (h.toLowerCase().includes('%')) {
            // Dev. % in Excel was fraction (e.g. -0.15) — show as percent
            td.textContent = (v * 100).toFixed(3) + "%";
          } else {
            td.textContent = Number(v).toFixed(4);
          }
        } else {
          td.textContent = v;
        }
        tr.appendChild(td);
      });
      tableBody.appendChild(tr);
    });

  })
  .catch(err => {
    console.error(err);
    document.getElementById('table-body').innerHTML = '<tr><td colspan="20">Error loading data — check console.</td></tr>';
  });
