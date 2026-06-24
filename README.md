# Incident Reporter

Simple Node.js app to upload `incident.xlsx`, aggregate tickets by engineer/status/age-bucket, and display results as HTML tables.

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
- If present, SLA columns such as `SLA Met` or `SLA Breached` are used to build a day-wise summary of Met vs Not Met totals.
- If present, `First workgroup name` is used to build a team-wise SLA breach summary with breached and met counts.
- Age buckets: `0-5`, `6-10`, `11-15`, `>15` days.
