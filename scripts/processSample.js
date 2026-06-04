const fs = require('fs');
const path = require('path');
const { processBuffer, renderHtmlTable } = require('../src/processor');

async function main() {
  const args = process.argv.slice(2);
  if (args.length < 1) {
    console.log('Usage: node scripts/processSample.js <input.xlsx> [out.html]');
    process.exit(1);
  }
  const input = args[0];
  const out = args[1] || 'out.html';
  const startDate = args[2] || null;
  const endDate = args[3] || null;
  const slaDays = args[4] ? parseInt(args[4], 10) : null;
  const buffer = fs.readFileSync(input);
  const data = await processBuffer(buffer, { startDate, endDate, slaDays });
  const html = renderHtmlTable(data);
  fs.writeFileSync(out, html, 'utf8');
  console.log('Wrote', out);
}

main().catch(err => {
  console.error(err);
  process.exit(1);
});
