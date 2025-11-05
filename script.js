// script.js — load and compute data directly from Excel (XLSX)
<script src="https://cdn.jsdelivr.net/npm/xlsx@0.19.3/dist/xlsx.full.min.js"></script>

function computeRowComputedFields(row, headers) {
  const getVal = (colLetter) => {
    const idx = colLetter.charCodeAt(0) - 'A'.charCodeAt(0);
    const h = headers[idx] || null;
    const raw = h && row[h] !== undefined && row[h] !== '' ? String(row[h]).replace(/,/g,'').trim() : '';
    const v = raw === '' ? NaN : parseFloat(raw);
    return isNaN(v) ? NaN : v;
  };

  const out = {};
  try { out['Dev.'] = getVal('D') - getVal('C'); } catch(e){ out['Dev.']=''; }
  try { out['Dev. %'] = getVal('E') / getVal('C'); } catch(e){ out['Dev. %']=''; }
  try { out['P/E'] = getVal('B') / getVal('H'); } catch(e){ out['P/E']=''; }
  try { out['P/B.V'] = getVal('B') / getVal('L'); } catch(e){ out['P/B.V']=''; }
  try { out['P/E*P/B.V.'] = getVal('J') * getVal('M'); } catch(e){ out['P/E*P/B.V.']=''; }
  try { out['L.P/G.N'] = getVal('B') / getVal('P'); } catch(e){ out['L.P/G.N']=''; }
  return out;
}

document.addEventListener('DOMContentLoaded', () => {
  const input = document.createElement('input');
  input.type = 'file';
  input.accept = '.xlsx';
  input.style.margin = '20px';
  document.body.insertBefore(input, document.body.firstChild);

  input.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const data = new Uint8Array(evt.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const firstSheet = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheet];
      const json = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

      const data = json;
      const headers = Object.keys(data[0]);
      const computedNames = [];
      const rowsComputed = data.map(row => {
        const comp = computeRowComputedFields(row, headers);
        for (const k of Object.keys(comp)) {
          if (!computedNames.includes(k)) computedNames.push(k);
        }
        return { original: row, computed: comp };
      });

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

      rowsComputed.forEach(rc => {
        const tr = document.createElement('tr');
        headers.forEach(h => {
          const td = document.createElement('td');
          td.textContent = rc.original[h] !== undefined ? rc.original[h] : "";
          tr.appendChild(td);
        });
        computedNames.forEach(h => {
          const td = document.createElement('td');
          let v = rc.computed[h];
          if (typeof v === 'number') {
            if (h.toLowerCase().includes('%')) {
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
    };
    reader.readAsArrayBuffer(file);
  });
});
