# Incident Reporter

Simple Node.js app to upload `incident.xlsx`, aggregate tickets by engineer/status/age-bucket, and display results as an HTML table.

Setup

```bash
npm install
```

Run

```bash
npm start
# then open http://localhost:3000
```

Development

```bash
npm run dev
```

Process sample via CLI

```bash
node scripts/processSample.js incident.xlsx out.html
```

Notes

- Expects the input Excel file to have columns: `TicketID`, `OpenedDate`, `Status`, `Caller`, `Engineer` (case-insensitive).
- Age buckets: `0-5`, `6-10`, `11-15`, `>15` days.
