## Plan: Incident Reporter Webapp

TL;DR: Build a small Node.js + Express app that accepts the daily `incident.xlsx` upload, parses the tickets (columns: TicketID, OpenedDate, Status, Caller, Engineer), computes ticket age in days, buckets ages into `0-5`, `6-10`, `11-15`, `>15`, aggregates counts grouped by `Engineer` × `Status` × `Age Category`, and renders the results as a simple HTML grid/table for viewing or download.

**Steps**
1. Initialize project: create `package.json` and install dependencies: `express`, `multer`, `dayjs`, `cors`. Add `start` and `dev` scripts.
2. Create server entry `src/server.js`: static serve `public/`, setup `POST /upload` to accept file upload (multer), call processing function, return rendered HTML table or JSON response.
3. Implement processing logic in `src/processor.js`: parse workbook buffer, normalize column names, compute age in days (today - `OpenedDate`), bucket into categories `0-5`, `6-10`, `11-15`, `>15`, aggregate counts by `Engineer` + `Status` + `Age Category`, and produce a data structure suitable for rendering as an HTML table. Export testable functions `processBuffer(buffer)` and `renderHtmlTable(aggregatedData)`.
4. Add minimal frontend `public/index.html`: upload form, progress indicator, and an HTML table view of results (with an option to download the table as an HTML file). Use plain HTML + fetch.
5. Add `scripts/processSample.js`: a small node runner to call `processBuffer` on the included `incident.xlsx` and write `out.html` or print the HTML table to stdout for local verification.
6. Documentation: create `README.md` with setup and run instructions and example CLI commands.
7. Verification: run the app, upload the sample file and confirm the rendered HTML table matches expected grouping and counts. Add a small unit test for the bucketing logic (optional).

**Relevant files**
- `package.json` — dependencies & scripts
- `src/server.js` — express app + upload endpoints
- `src/processor.js` — parsing, bucketing, aggregation, renderable data structure
- `public/index.html` — upload UI + table view and download option
- `scripts/processSample.js` — CLI runner to process sample file and emit HTML
- `README.md` — setup & usage

**Verification**
1. Run `npm install`.
2. Start app: `npm start` (or `npm run dev` for nodemon).
3. Open the upload UI in a browser, upload `incident.xlsx`, view the HTML table. Verify counts by engineer/status/age category.
4. Run `node scripts/processSample.js incident.xlsx out.html` (or print to console) and manually compare with your expected output.

**Decisions**
- Use `dayjs` for date math and a lightweight Excel parser (via `exceljs` or `xlsx`) to read the input workbook.
- Age buckets: `0-5`, `6-10`, `11-15`, `>15`.
- Report includes all statuses separately.
- Output format: HTML table/grid rendered on the web UI and optionally saved as an HTML file.
- No authentication (internal use).

**Further Considerations**
1. If input date formats vary, add parsing heuristics or let user select date format on the UI.
2. For very large files, switch to streaming parsing to reduce memory usage.
3. Optionally add basic logging and an admin page to review past uploads.
