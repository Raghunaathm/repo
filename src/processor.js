const ExcelJS = require('exceljs');
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
  // accept ArrayBuffer/Uint8Array as well as Buffer
  if (!Buffer.isBuffer(buffer) && buffer && buffer.buffer) {
    buffer = Buffer.from(buffer);
  }
  if (!Buffer.isBuffer(buffer)) throw new Error('processBuffer expects a Buffer or ArrayBuffer');

  const workbook = new ExcelJS.Workbook();
  try {
    // Quick check: .xlsx files are ZIP archives and start with 'PK' (0x50 0x4B)
    if (buffer.length < 4 || buffer[0] !== 0x50 || buffer[1] !== 0x4B) {
      throw new Error('Not a .xlsx (ZIP) file — received a different file type (maybe .xls or CSV)');
    }
    await workbook.xlsx.load(buffer);
  } catch (err) {
    throw new Error('Invalid or corrupted Excel file: ' + err.message);
  }
  const worksheet = workbook.worksheets[0];
  if (!worksheet) throw new Error('No worksheet found in workbook');

  // Read header row and map columns
  const headerRow = worksheet.getRow(1);
  const headers = {};
  headerRow.eachCell((cell, colNumber) => {
    const val = String(cell.value || '').trim().toLowerCase();
    headers[val] = colNumber;
  });

  const col = (names) => {
    for (const n of names) {
      const key = n.toLowerCase();
      if (headers[key]) return headers[key];
    }
    return null;
  };

  const ticketIdCol = col(['ticketid', 'ticket id', 'id']);
  const openedCol = col(['openeddate', 'opened date', 'opened', 'created']);
  const statusCol = col(['status']);
  const callerCol = col(['caller']);
  const engineerCol = col(['engineer']);

  const aggregated = {}; // engineer -> status -> bucket -> count
  const engineersSet = new Set();
  const statusesSet = new Set();

  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber === 1) return; // skip header
    const openedCell = openedCol ? row.getCell(openedCol).value : null;
    let openedDate = null;
    if (openedCell instanceof Date) openedDate = dayjs(openedCell);
    else if (openedCell && typeof openedCell === 'object' && openedCell.text) openedDate = dayjs(openedCell.text);
    else if (openedCell) openedDate = dayjs(String(openedCell));

    const now = dayjs();
    const days = openedDate && openedDate.isValid() ? now.diff(openedDate, 'day') : 0;
    const bucket = bucketForDays(days);

    const engineer = engineerCol ? String(row.getCell(engineerCol).text || row.getCell(engineerCol).value || '').trim() : '';
    const status = statusCol ? String(row.getCell(statusCol).text || row.getCell(statusCol).value || '').trim() : '';

    const engKey = engineer || 'Unassigned';
    const statusKey = status || 'Unknown';

    engineersSet.add(engKey);
    statusesSet.add(statusKey);

    if (!aggregated[engKey]) aggregated[engKey] = {};
    if (!aggregated[engKey][statusKey]) aggregated[engKey][statusKey] = {};
    if (!aggregated[engKey][statusKey][bucket]) aggregated[engKey][statusKey][bucket] = 0;
    aggregated[engKey][statusKey][bucket] += 1;
  });

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
