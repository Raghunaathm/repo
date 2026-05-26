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
  const buffer = fs.readFileSync(input);
  const data = await processBuffer(buffer);
  const html = renderHtmlTable(data);
  fs.writeFileSync(out, html, 'utf8');
  console.log('Wrote', out);
}

main().catch(err => {
  console.error(err);
  process.exit(1);
});
