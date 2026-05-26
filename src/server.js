const path = require('path');
const express = require('express');
const multer = require('multer');
const cors = require('cors');

const { processBuffer, renderHtmlTable } = require('./processor');

const app = express();
app.use(cors());
app.use(express.static(path.join(__dirname, '..', 'public')));

const upload = multer({ storage: multer.memoryStorage() });

app.post('/upload', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).send('No file uploaded');
    const aggregated = await processBuffer(req.file.buffer);
    const html = renderHtmlTable(aggregated);
    res.send(html);
  } catch (err) {
    console.error(err);
    res.status(500).send('Processing error: ' + err.message);
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Incident Reporter listening on http://localhost:${PORT}`);
});
