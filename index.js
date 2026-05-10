const express = require('express');
const cors = require('cors');
const multer = require('multer');
const mammoth = require('mammoth');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');

const app = express();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 25 * 1024 * 1024 } });

app.use(cors());
app.use(express.json({ limit: '25mb' }));
app.use(express.urlencoded({ extended: true, limit: '25mb' }));

function safeName(name = 'arquivo.docx') {
  try {
    return Buffer.from(name, 'latin1').toString('utf8');
  } catch (e) {
    return name;
  }
}

function cleanText(text = '') {
  return String(text)
    .normalize('NFC')
    .replace(/\u0000/g, '')
    .replace(/[\u0001-\u0008\u000B\u000C\u000E-\u001F]/g, '')
    .replace(/\r/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .replace(/[ \t]{2,}/g, ' ')
    .trim();
}

app.get('/', (req, res) => {
  res.json({ status: 'ok', service: 'docx-service-t2r', routes: ['/health', '/extract-docx', '/render-t2r'] });
});

app.get('/health', (req, res) => {
  res.json({ status: 'ok', service: 'docx-service-t2r' });
});

app.post('/extract-docx', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'Nenhum arquivo enviado no campo file.' });

    const result = await mammoth.extractRawText({ buffer: req.file.buffer });
    const text = cleanText(result.value || '');

    res.json({
      fileName: safeName(req.file.originalname),
      text,
      length: text.length,
      warnings: result.messages || []
    });
  } catch (error) {
    res.status(500).json({ error: 'Erro ao extrair texto do DOCX.', details: error.message });
  }
});

app.post('/render-t2r', upload.single('template'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'Nenhum template enviado no campo template.' });

    let data = req.body.data || req.body;
    if (typeof data === 'string') data = JSON.parse(data);

    const zip = new PizZip(req.file.buffer);
    const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });

    doc.render(data);

    const buffer = doc.getZip().generate({ type: 'nodebuffer', compression: 'DEFLATE' });
    const outputName = data.nome_arquivo_saida || 'T2R_Gerado.docx';

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${outputName}"`);
    res.send(buffer);
  } catch (error) {
    res.status(500).json({ error: 'Erro ao gerar DOCX do T2R.', details: error.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, '0.0.0.0', () => {
  console.log(`docx-service-t2r rodando na porta ${PORT}`);
});
