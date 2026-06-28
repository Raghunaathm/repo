const XLSX = require('xlsx');
const dayjs = require('dayjs');

const AGE_BUCKETS = [
  { name: '0-5', min: 0, max: 5 },
  { name: '6-10', min: 6, max: 10 },
  { name: '11-15', min: 11, max: 15 },
  { name: '>15', min: 16, max: Infinity }
];

function bucketForDays(days) {
  for (const b of AGE_BUCKETS) {
    if (days >= b.min && days <= b.max) return b.name;
  }
  return '>15';
}

async function processBuffer(buffer, options = {}) {
  if (!buffer) throw new Error('No data buffer provided to processBuffer');
  if (!Buffer.isBuffer(buffer) && buffer && buffer.buffer) {
    buffer = Buffer.from(buffer);
  }
  if (!Buffer.isBuffer(buffer)) throw new Error('processBuffer expects a Buffer or ArrayBuffer');

  if (buffer.length < 4 || buffer[0] !== 0x50 || buffer[1] !== 0x4B) {
    throw new Error('Not a .xlsx (ZIP) file — received a different file type (maybe .xls or CSV)');
  }

  let workbook;
  try {
    workbook = XLSX.read(buffer, { type: 'buffer', cellDates: true });
  } catch (err) {
    throw new Error('Invalid or corrupted Excel file: ' + err.message);
  }

  const sheetName = workbook.SheetNames[0];
  if (!sheetName) throw new Error('No worksheet found in workbook');
  const worksheet = workbook.Sheets[sheetName];

  // Sheet may have a decorative first row; scan row 1 and row 2 for real headers
  const rawRows = XLSX.utils.sheet_to_json(worksheet, { raw: false, defval: '', header: 1 });
  if (rawRows.length < 2) throw new Error('Worksheet has no data rows');

  // Find the header row: first row where >3 non-empty cells exist
  let headerRowIdx = 0;
  for (let i = 0; i < Math.min(rawRows.length, 5); i++) {
    if (rawRows[i].filter(v => String(v).trim()).length > 3) { headerRowIdx = i; break; }
  }

  const headerRow = rawRows[headerRowIdx].map(v => String(v).trim().toLowerCase());

  const col = (names) => {
    for (const n of names) {
      const idx = headerRow.findIndex(h => h.includes(n.toLowerCase()));
      if (idx !== -1) return idx;
    }
    return -1;
  };

  const openedColIdx = col(['log time', 'openeddate', 'opened date', 'opened', 'created']);
  const statusColIdx = col(['status']);
  const engineerColIdx = col(['assigned to engineer', 'engineer']);
  const classificationColIdx = col(['classification', 'classification type', 'ticket classification', 'ticket type', 'type', 'category', 'service category']);
  const workgroupNameColIdx = col(['workgroup name', 'workgroup']);
  const firstWorkgroupColIdx = col(['first workgroup name', 'first workgroup', 'workgroup', 'assignment group', 'team']);
  const slaMetColIdx = col(['response sla met', 'response sla', 'sla met', 'sla_met', 'sla met?', 'sla_met?']);
  const slaBreachColIdx = col(['sla breached', 'breach', 'breached']);

  const aggregated = {}; // engineer -> status -> bucket -> count
  const engineersSet = new Set();
  const statusesSet = new Set();
  const slaByDay = {};
  const slaBreachByTeam = {};
  const responseSlaByClassification = {};
  const classificationsSet = new Set();
  const responseMonthSet = new Set();

  for (let i = headerRowIdx + 1; i < rawRows.length; i++) {
    const row = rawRows[i];
    const openedRaw = openedColIdx >= 0 ? row[openedColIdx] : null;
    let openedDate = null;
    if (openedRaw instanceof Date) openedDate = dayjs(openedRaw);
    else if (openedRaw) openedDate = dayjs(String(openedRaw));

    // apply date filters if provided
    const start = options.startDate ? dayjs(options.startDate) : null;
    const end = options.endDate ? dayjs(options.endDate) : null;
    if ((start || end)) {
      if (!openedDate || !openedDate.isValid()) continue; // cannot determine date, skip
      if (start && openedDate.isBefore(start, 'day')) continue;
      if (end && openedDate.isAfter(end, 'day')) continue;
    }

    const now = dayjs();
    const days = openedDate && openedDate.isValid() ? now.diff(openedDate, 'day') : 0;
    const bucket = bucketForDays(days);

    const engineer = engineerColIdx >= 0 ? String(row[engineerColIdx] || '').trim() : '';
    const classification = classificationColIdx >= 0 ? String(row[classificationColIdx] || '').trim() : 'All';
    const workgroupName = workgroupNameColIdx >= 0 ? String(row[workgroupNameColIdx] || '').trim() : '';
    const responseSlaGroup = workgroupName || classification || 'All';
    const firstWorkgroup = firstWorkgroupColIdx >= 0 ? String(row[firstWorkgroupColIdx] || '').trim() : '';
    const status = statusColIdx >= 0 ? String(row[statusColIdx] || '').trim() : '';
    const slaMetRaw = slaMetColIdx >= 0 ? String(row[slaMetColIdx] || '').trim().toLowerCase() : '';
    const slaBreachRaw = slaBreachColIdx >= 0 ? String(row[slaBreachColIdx] || '').trim().toLowerCase() : '';

    const engKey = engineer || 'Unassigned';
    const statusKey = status || 'Unknown';

    let slaState = null;
    if (slaMetRaw) {
      if (['yes', 'true', '1', 'met'].includes(slaMetRaw)) slaState = 'met';
      else if (['no', 'false', '0', 'not met', 'not_met'].includes(slaMetRaw)) slaState = 'not met';
    } else if (slaBreachRaw) {
      if (['yes', 'true', '1', 'breached'].includes(slaBreachRaw)) slaState = 'not met';
      else if (['no', 'false', '0'].includes(slaBreachRaw)) slaState = 'met';
    }

    const dayKey = openedDate && openedDate.isValid() ? openedDate.format('YYYY-MM-DD') : 'Unknown';
    const monthKey = openedDate && openedDate.isValid() ? openedDate.format('YYYY-MM') : 'Unknown';

    engineersSet.add(engKey);
    classificationsSet.add(classification || 'All');
    statusesSet.add(statusKey);

    if (!slaByDay[dayKey]) slaByDay[dayKey] = { met: 0, notMet: 0, total: 0 };
    if (slaState === 'met') slaByDay[dayKey].met += 1;
    if (slaState === 'not met') slaByDay[dayKey].notMet += 1;
    slaByDay[dayKey].total += 1;

    const teamKey = firstWorkgroup || engKey;
    if (!slaBreachByTeam[teamKey]) slaBreachByTeam[teamKey] = { breached: 0, met: 0, total: 0 };
    if (slaState === 'met') slaBreachByTeam[teamKey].met += 1;
    if (slaState === 'not met') slaBreachByTeam[teamKey].breached += 1;
    slaBreachByTeam[teamKey].total += 1;

    if (slaState) {
      if (!responseSlaByClassification[responseSlaGroup]) responseSlaByClassification[responseSlaGroup] = {};
      if (!responseSlaByClassification[responseSlaGroup][monthKey]) responseSlaByClassification[responseSlaGroup][monthKey] = { met: 0, notMet: 0, total: 0 };
      const field = slaState === 'met' ? 'met' : 'notMet';
      responseSlaByClassification[responseSlaGroup][monthKey][field] += 1;
      responseSlaByClassification[responseSlaGroup][monthKey].total += 1;
      responseMonthSet.add(monthKey);
    }

    if (!aggregated[engKey]) aggregated[engKey] = {};
    if (!aggregated[engKey][statusKey]) aggregated[engKey][statusKey] = {};
    if (!aggregated[engKey][statusKey][bucket]) aggregated[engKey][statusKey][bucket] = 0;
    aggregated[engKey][statusKey][bucket] += 1;
  }

  return {
    aggregated,
    engineers: Array.from(engineersSet).sort(),
    statuses: Array.from(statusesSet).sort(),
    buckets: AGE_BUCKETS.map(b => b.name),
    slaByDay: Object.fromEntries(
      Object.entries(slaByDay).sort(([a], [b]) => a.localeCompare(b))
    ),
    slaBreachByTeam: Object.fromEntries(
      Object.entries(slaBreachByTeam).sort(([a], [b]) => a.localeCompare(b))
    ),
    responseSlaByClassification: Object.fromEntries(
      Object.entries(responseSlaByClassification).sort(([a], [b]) => a.localeCompare(b))
    ),
    responseSlaMonths: Array.from(responseMonthSet).sort()
  };
}

function renderHtmlTable(data) {
  const { aggregated, engineers, statuses, buckets, slaByDay, slaBreachByTeam, responseSlaByClassification, responseSlaMonths } = data;
  const visibleStatuses = (statuses || []).filter(status => !['cancelled', 'new'].includes(String(status).trim().toLowerCase()));
  let html = `<!doctype html><html><head><meta charset="utf-8"><title>Incident Report</title><style>body{font-family:Inter,Segoe UI,Arial,sans-serif;margin:24px;color:#0f172a;background:#f8fafc}h2,h3{margin:16px 0 10px;color:#0f172a}.toolbar{display:flex;justify-content:flex-end;margin:8px 0 16px}.toolbar button{background:#2563eb;color:#fff;border:none;border-radius:8px;padding:10px 14px;font-weight:600;cursor:pointer}.toolbar button:hover{background:#1d4ed8}.table-wrap{overflow-x:auto;background:#fff;border:1px solid #e2e8f0;border-radius:12px;padding:8px;box-shadow:0 8px 24px rgba(15,23,42,.06)}table{border-collapse:separate;border-spacing:0;width:100%;min-width:900px;table-layout:fixed;font-size:13px;background:#fff}th,td{border:1px solid #e2e8f0;padding:8px 10px;text-align:center;white-space:nowrap}th{background:linear-gradient(180deg,#f8fafc 0%,#eef2ff 100%);color:#334155;font-weight:700;position:sticky;top:0;z-index:1}tbody tr:nth-child(even) td{background:#fbfdff}tbody td:first-child{background:#f8fafc;text-align:left;font-weight:600;color:#111827}tbody td:not(:first-child){text-align:center}.totals-row td{font-weight:700;background:#eef2ff!important}</style></head><body>`;
  html += `<div class="toolbar"><button type="button" id="downloadCsvBtn">Download CSV</button></div>`;
  html += `<h2>Incident Report</h2>`;
  if (slaByDay && Object.keys(slaByDay).length) {
    html += `<h3>SLA by Day</h3>`;
    html += `<div class="table-wrap"><table class="report-table" data-title="SLA by Day">`;
    html += `<thead><tr><th>Date</th><th>Met</th><th>Not Met</th><th>Total</th></tr></thead><tbody>`;
    let metTotal = 0;
    let notMetTotal = 0;
    let dayTotal = 0;
    for (const [day, counts] of Object.entries(slaByDay)) {
      html += `<tr><td>${escapeHtml(day)}</td><td>${counts.met}</td><td>${counts.notMet}</td><td>${counts.total}</td></tr>`;
      metTotal += counts.met;
      notMetTotal += counts.notMet;
      dayTotal += counts.total;
    }
    html += `<tr class="totals-row"><td><strong>Total</strong></td><td><strong>${metTotal}</strong></td><td><strong>${notMetTotal}</strong></td><td><strong>${dayTotal}</strong></td></tr>`;
    html += `</tbody></table></div>`;
  }
  if (responseSlaByClassification && Object.keys(responseSlaByClassification).length) {
    const months = Array.from(new Set(responseSlaMonths)).sort();
    html += `<h3>Response SLA by Workgroup Name</h3>`;
    html += `<div class="table-wrap"><table class="report-table" data-title="Response SLA by Workgroup Name">`;
    html += `<thead><tr><th rowspan="2">Classification</th>`;
    for (const month of months) {
      html += `<th colspan="3">${escapeHtml(month)}</th>`;
    }
    html += `<th rowspan="2">Grand Total</th></tr>`;
    html += `<tr>`;
    for (let i = 0; i < months.length; i++) {
      html += `<th>Met</th><th>Not Met</th><th>Total</th>`;
    }
    html += `</tr></thead><tbody>`;

    let grandMet = 0;
    let grandNotMet = 0;
    let grandTotal = 0;

    for (const [classification, monthData] of Object.entries(responseSlaByClassification)) {
      html += `<tr><td>${escapeHtml(classification)}</td>`;
      let classificationTotal = 0;
      for (const month of months) {
        const counts = monthData[month] || { met: 0, notMet: 0, total: 0 };
        html += `<td>${counts.met}</td><td>${counts.notMet}</td><td>${counts.total}</td>`;
        grandMet += counts.met;
        grandNotMet += counts.notMet;
        grandTotal += counts.total;
        classificationTotal += counts.total;
      }
      html += `<td><strong>${classificationTotal}</strong></td></tr>`;
    }

    html += `<tr class="totals-row"><td><strong>Total</strong></td>`;
    for (const month of months) {
      const monthSum = months.reduce((sum, m) => sum, 0); // placeholder logic, values are already in grand totals below
      const met = Object.values(responseSlaByClassification).reduce((sum, monthData) => sum + ((monthData[month] && monthData[month].met) || 0), 0);
      const notMet = Object.values(responseSlaByClassification).reduce((sum, monthData) => sum + ((monthData[month] && monthData[month].notMet) || 0), 0);
      const total = Object.values(responseSlaByClassification).reduce((sum, monthData) => sum + ((monthData[month] && monthData[month].total) || 0), 0);
      html += `<td><strong>${met}</strong></td><td><strong>${notMet}</strong></td><td><strong>${total}</strong></td>`;
    }
    html += `<td><strong>${grandTotal}</strong></td></tr>`;
    html += `</tbody></table></div>`;
  }

  html += `<div class="table-wrap"><table class="report-table" data-title="Engineer by Status and Age Bucket">`;
  // Build two-row header: first row groups buckets under each status
  html += `<thead>`;
  html += `<tr><th rowspan="2">Engineer</th>`;
  if (visibleStatuses.length === 0) {
    // fallback: single status column with buckets
    html += `<th colspan="${buckets.length}">Status</th>`;
  } else {
    for (const status of visibleStatuses) {
      html += `<th colspan="${buckets.length}">${escapeHtml(status)}</th>`;
    }
  }
  // Total column for each engineer
  html += `<th rowspan="2">Total</th>`;
  html += `</tr>`;

  // second header row: bucket names repeated for each status
  html += `<tr>`;
  if (visibleStatuses.length === 0) {
    for (const b of buckets) html += `<th>${b}</th>`;
  } else {
    for (let i = 0; i < visibleStatuses.length; i++) {
      for (const b of buckets) html += `<th>${b}</th>`;
    }
  }
  html += `</tr>`;
  html += `</thead><tbody>`;

  // prepare column sums to render a final totals row
  const colCount = (visibleStatuses.length === 0) ? buckets.length : visibleStatuses.length * buckets.length;
  const colSums = new Array(colCount).fill(0);
  let grandTotal = 0;

  // Body: one row per engineer, columns are status × buckets
  for (const engineer of engineers) {
    html += `<tr><td>${escapeHtml(engineer)}</td>`;
    let rowTotal = 0;
    if (visibleStatuses.length === 0) {
      // no statuses found, sum across all statuses for each bucket
      for (let ci = 0; ci < buckets.length; ci++) {
        const b = buckets[ci];
        let sum = 0;
        const byStatus = aggregated[engineer] || {};
        for (const s of Object.keys(byStatus)) {
          sum += (byStatus[s] && byStatus[s][b]) || 0;
        }
        html += `<td>${sum}</td>`;
        colSums[ci] += sum;
        rowTotal += sum;
      }
    } else {
      let ci = 0;
      for (const status of visibleStatuses) {
        for (const b of buckets) {
          const val = (aggregated[engineer] && aggregated[engineer][status] && aggregated[engineer][status][b]) || 0;
          html += `<td>${val}</td>`;
          colSums[ci] += val;
          rowTotal += val;
          ci++;
        }
      }
    }
    html += `<td><strong>${rowTotal}</strong></td>`;
    grandTotal += rowTotal;
    html += `</tr>`;
  }

  // Totals row
  html += `<tr class="totals-row"><td><strong>Total</strong></td>`;
  for (let i = 0; i < colSums.length; i++) html += `<td><strong>${colSums[i]}</strong></td>`;
  html += `<td><strong>${grandTotal}</strong></td>`;
  html += `</tr>`;

  html += `</tbody></table></div>`;
  html += `<script>
      const btn = document.getElementById('downloadCsvBtn');
      if (btn) {
        btn.addEventListener('click', () => {
          const tables = Array.from(document.querySelectorAll('table.report-table'));
          const lines = [];
          const escapeCsv = (value) => '"' + String(value).replace(/"/g, '""') + '"';
          const toCsvRow = (values) => values.map(escapeCsv).join(',');

          tables.forEach((table, index) => {
            const title = table.getAttribute('data-title') || ('Table ' + (index + 1));
            lines.push(toCsvRow([title]));
            Array.from(table.querySelectorAll('tr')).forEach((row) => {
              const values = Array.from(row.querySelectorAll('th, td')).map(cell => cell.textContent.trim());
              if (values.length) lines.push(toCsvRow(values));
            });
            lines.push('');
          });

          const csv = lines.join('\n');
          const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
          const link = document.createElement('a');
          link.href = URL.createObjectURL(blob);
          link.download = 'incident-report.csv';
          document.body.appendChild(link);
          link.click();
          document.body.removeChild(link);
          URL.revokeObjectURL(link.href);
        });
      }
  </script></body></html>`;
  return html;
}

function escapeHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

module.exports = { processBuffer, renderHtmlTable };
