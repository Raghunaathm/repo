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

async function processBuffer(buffer) {
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

  const aggregated = {}; // engineer -> status -> bucket -> count
  const engineersSet = new Set();
  const statusesSet = new Set();

  for (let i = headerRowIdx + 1; i < rawRows.length; i++) {
    const row = rawRows[i];
    const openedRaw = openedColIdx >= 0 ? row[openedColIdx] : null;
    let openedDate = null;
    if (openedRaw instanceof Date) openedDate = dayjs(openedRaw);
    else if (openedRaw) openedDate = dayjs(String(openedRaw));

    const now = dayjs();
    const days = openedDate && openedDate.isValid() ? now.diff(openedDate, 'day') : 0;
    const bucket = bucketForDays(days);

    const engineer = engineerColIdx >= 0 ? String(row[engineerColIdx] || '').trim() : '';
    const status = statusColIdx >= 0 ? String(row[statusColIdx] || '').trim() : '';

    const engKey = engineer || 'Unassigned';
    const statusKey = status || 'Unknown';

    engineersSet.add(engKey);
    statusesSet.add(statusKey);

    if (!aggregated[engKey]) aggregated[engKey] = {};
    if (!aggregated[engKey][statusKey]) aggregated[engKey][statusKey] = {};
    if (!aggregated[engKey][statusKey][bucket]) aggregated[engKey][statusKey][bucket] = 0;
    aggregated[engKey][statusKey][bucket] += 1;
  }

  return {
    aggregated,
    engineers: Array.from(engineersSet).sort(),
    statuses: Array.from(statusesSet).sort(),
    buckets: AGE_BUCKETS.map(b => b.name)
  };
}

function renderHtmlTable(data) {
  const { aggregated, engineers, statuses, buckets } = data;
  let html = `<!doctype html><html><head><meta charset="utf-8"><title>Incident Report</title><style>table{border-collapse:collapse;width:100%}th,td{border:1px solid #ccc;padding:6px;text-align:center}th{background:#eee}</style></head><body>`;
  html += `<h2>Incident Report</h2>`;
  html += `<table>`;
  html += `<thead><tr><th>Engineer</th><th>Status</th>`;
  for (const b of buckets) html += `<th>${b}</th>`;
  html += `</tr></thead><tbody>`;

  for (const engineer of engineers) {
    const statusesForEngineer = Object.keys(aggregated[engineer] || {}).sort();
    if (statusesForEngineer.length === 0) {
      html += `<tr><td>${escapeHtml(engineer)}</td><td></td>`;
      for (const b of buckets) html += `<td>0</td>`;
      html += `</tr>`;
      continue;
    }
    for (const status of statusesForEngineer) {
      html += `<tr><td>${escapeHtml(engineer)}</td><td>${escapeHtml(status)}</td>`;
      for (const b of buckets) {
        const val = (aggregated[engineer] && aggregated[engineer][status] && aggregated[engineer][status][b]) || 0;
        html += `<td>${val}</td>`;
      }
      html += `</tr>`;
    }
  }

  html += `</tbody></table></body></html>`;
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
