import express from 'express';
import cors from 'cors';
import multer from 'multer';
import { promises as fs, createWriteStream, readFileSync } from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { PDFDocument, rgb } from 'pdf-lib';
import XLSX from 'xlsx';
import { parse } from 'csv-parse/sync';
import archiver from 'archiver';
import { nanoid } from 'nanoid';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const app = express();

const allowedOrigins = (process.env.CORS_ORIGIN || '')
  .split(',')
  .map((origin) => origin.trim())
  .filter(Boolean);

if (allowedOrigins.length) {
  app.use(cors({
    origin: (origin, callback) => {
      // Requests without Origin are typically server-to-server or same-origin navigations.
      if (!origin) return callback(null, true);
      if (allowedOrigins.includes(origin)) return callback(null, true);
      return callback(new Error('Origin not allowed by CORS'));
    }
  }));
}
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

const UPLOAD_DIR = path.join(__dirname, 'data', 'uploads');
const OUTPUT_DIR = path.join(__dirname, 'output');
const PROJECTS_DIR = path.join(__dirname, 'projects');
const TEMPLATES_DIR = path.join(__dirname, 'templates');

await fs.mkdir(UPLOAD_DIR, { recursive: true });
await fs.mkdir(OUTPUT_DIR, { recursive: true });
await fs.mkdir(PROJECTS_DIR, { recursive: true });
await fs.mkdir(TEMPLATES_DIR, { recursive: true });

const storage = multer.diskStorage({
  destination: async (req, file, cb) => {
    await fs.mkdir(UPLOAD_DIR, { recursive: true });
    cb(null, UPLOAD_DIR);
  },
  filename: (req, file, cb) => {
    const safe = file.originalname.replace(/[^a-zA-Z0-9_.-]/g, '_');
    cb(null, `${Date.now()}_${safe}`);
  }
});
const upload = multer({ storage });

// ---------- Helpers ----------

async function readJsonSafe(filePath) {
  try { return JSON.parse(await fs.readFile(filePath, 'utf8')); }
  catch { return null; }
}

function normalizeKey(key) {
  // Collapse every run of non-alphanumerics (slash, underscore, dash, dot,
  // whitespace) to a single space. Critically this keeps "Malpractice/Liability"
  // as two tokens instead of gluing it into "malpracticeliability".
  return String(key ?? '').trim().toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
}

// Repair common source-header defects (typos, glued words, abbreviations)
// BEFORE normalization so credentialing terms tokenize correctly.
function canonicalizeHeader(raw) {
  let s = ' ' + String(raw ?? '').toLowerCase() + ' ';
  s = s
    .replace(/certificatin/g, 'certification')          // typo: missing 'o'
    .replace(/certification\s*expiration/g, 'certification expiration')
    .replace(/certification\s*effective/g, 'certification effective')
    .replace(/([a-z])expiration/g, '$1 expiration')      // glued: ...CertificationExpiration
    .replace(/([a-z])effective/g, '$1 effective');
  return normalizeKey(s);
}

// Healthcare credentialing synonym expansion. Each canonical source key maps to
// alias phrases; the matcher scores every alias against each PDF field and keeps
// the best. Aliases carry the distinctive context tokens (e.g. "individual",
// "group", "board") so generic repeated PDF fields don't win by accident.
const HEADER_ALIASES = {
  'zip': ['zip', 'zip code', 'postal code'],
  'practice name': ['practice name', 'group corporate name as it appears on irs w9', 'group name', 'corporate name', 'organization name', 'facility name', 'business name'],
  'tax id': ['tax id', 'tax id number', 'tax identification number', 'federal tax id', 'tin', 'ein', 'employer identification number'],
  'malpractice liability': ['malpractice liability', 'professional liability', 'malpractice insurance', 'liability insurance', 'malpractice policy', 'liability policy number'],
  'individual npi': ['individual npi', 'individual national provider identifier', 'type 1 npi', 'provider npi', 'npi number'],
  'group npi': ['group npi', 'group national provider identifier', 'type 2 npi', 'organization npi', 'billing npi'],
  'caqh': ['caqh', 'caqh id', 'caqh provider id', 'caqh number'],
  'medicare provider number': ['medicare provider number', 'medicare number', 'medicare id', 'ptan', 'participating medicare provider'],
  'medicaid provider number': ['medicaid provider number', 'medicaid number', 'medicaid id', 'site specific medicaid number', 'tpi'],
  'board certification number': ['board certification number', 'board certificate number', 'certification number', 'certificate number'],
  'board certification effective date': ['board certification effective date', 'board effective date', 'certification effective date', 'board date effective'],
  'board certification expiration date': ['board certification expiration date', 'board expiration date', 'certification expiration date', 'board date expiration'],
};

function levenshtein(a, b) {
  const an = a.length, bn = b.length;
  const matrix = Array.from({ length: an + 1 }, (_, i) => {
    const row = new Array(bn + 1);
    row[0] = i;
    return row;
  });
  for (let j = 0; j <= bn; j++) matrix[0][j] = j;
  for (let i = 1; i <= an; i++) {
    for (let j = 1; j <= bn; j++) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      matrix[i][j] = Math.min(matrix[i - 1][j] + 1, matrix[i][j - 1] + 1, matrix[i - 1][j - 1] + cost);
    }
  }
  return matrix[an][bn];
}

function similarity(a, b) {
  const an = a.length, bn = b.length;
  if (!an && !bn) return 1;
  if (!an || !bn) return 0;
  const maxLen = Math.max(an, bn);
  return 1 - levenshtein(a, b) / maxLen;
}

function containsSimilarity(a, b) {
  const an = a.length, bn = b.length;
  if (!an || !bn) return 0;
  const wordsA = a.split(' ');
  const wordsB = b.split(' ');
  let best = 0;
  for (const wa of wordsA) {
    if (!wa) continue;
    for (const wb of wordsB) {
      if (!wb || wb.length < 2) continue;
      const s = similarity(wa, wb);
      if (s > best) best = s;
    }
  }
  return best;
}

function patternBoost(sourceNorm, targetNorm) {
  const patterns = [
    { regex: /^\d{3}-\d{2}-\d{4}$/, label: 'ssn' },
    { regex: /^\(\d{3}\)\s?\d{3}-\d{4}$/, label: 'phone' },
    { regex: /^\d{3}-\d{3}-\d{4}$/, label: 'phone' },
    { regex: /^[^\s@]+@[^\s@]+\.[^\s@]+$/, label: 'email' },
    { regex: /^\$?\d+(\.\d{2})$/, label: 'currency' },
    { regex: /^\d{2}\/\d{2}\/\d{4}$/, label: 'date' },
  ];
  for (const p of patterns) {
    if (p.regex.test(sourceNorm)) return 1;
  }
  return 0;
}

async function parsePdfFormFields(filePath) {
  const pdfBytes = await fs.readFile(filePath);
  const pdfDoc = await PDFDocument.load(pdfBytes);
  const form = pdfDoc.getForm();
  const fields = [];
  const seen = new Set();
  const add = (field, pageIndex) => {
    let name = 'Unnamed';
    let type = 'unknown';
    try { name = String(field.getName() || 'Unnamed'); } catch {}
    try {
      if (field.constructor?.name === 'PDFTextField') type = 'text';
      else if (field.constructor?.name === 'PDFCheckBox') type = 'checkbox';
      else if (field.constructor?.name === 'PDFRadioGroup') type = 'radio';
      else type = field.constructor?.name || 'unknown';
    } catch {}
    if (seen.has(name)) return;
    seen.add(name);
    fields.push({ name, type, page: pageIndex + 1 });
  };
  for (const field of form.getFields()) add(field, 0);
  for (let i = 0; i < pdfDoc.getPages().length; i++) {
    try {
      for (const field of form.getPageFields(i)) add(field, i);
    } catch {
      // some page field APIs throw on unusual PDFs
    }
  }
  return fields;
}

async function parseDataFile(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  if (ext === '.csv') {
    const content = await fs.readFile(filePath, 'utf8');
    const records = parse(content, { columns: true, skip_empty_lines: true });
    const headers = records.length ? Object.keys(records[0]) : [];
    return { headers, rows: records };
  }
  if (ext === '.json') {
    const content = await fs.readFile(filePath, 'utf8');
    const records = JSON.parse(content);
    if (!Array.isArray(records) || !records.length) return { headers: [], rows: [] };
    const headers = Object.keys(records[0]);
    return { headers, rows: records };
  }
  if (ext === '.xlsx' || ext === '.xls') {
    const workbook = XLSX.readFile(filePath, { cellDates: true });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    // raw:false + dateNF renders Excel date serials as mm/dd/yyyy instead of
    // leaking numbers like 39700. defval keeps blank cells as ''.
    const data = XLSX.utils.sheet_to_json(sheet, { defval: '', raw: false, dateNF: 'mm/dd/yyyy' });
    // Trim stray whitespace (e.g. Medicare number stored with trailing spaces).
    for (const row of data) {
      for (const k of Object.keys(row)) {
        if (typeof row[k] === 'string') row[k] = row[k].trim();
      }
    }
    const headers = data.length ? Object.keys(data[0]) : [];
    return { headers, rows: data };
  }
  throw new Error('Unsupported data file format');
}

function scoreMatch(sourceNorm, targetNorm) {
  const exact = sourceNorm === targetNorm;
  if (exact) return { score: 1, reason: 'exact' };

  // Length penalty: if target is much longer than source, it's probably a bad match
  // But allow single-word sources through — they'll be checked by word-level matching below
  const lenRatio = Math.min(sourceNorm.length, targetNorm.length) / Math.max(sourceNorm.length, targetNorm.length);
  const srcWords = sourceNorm.split(/[\s_]+/).filter(w => w.length > 1);
  const tgtWords = targetNorm.split(/[\s_]+/).filter(w => w.length > 1);
  if (lenRatio < 0.3 && srcWords.length > 1) return { score: lenRatio * 0.4, reason: 'weak' };

  const normSim = similarity(sourceNorm, targetNorm);
  const containSim = containsSimilarity(sourceNorm, targetNorm);

  // Word-level matching: check if ALL source words appear in target
  const exactWordMatches = srcWords.filter(sw => tgtWords.some(tw => tw === sw)).length;
  const wordMatchRatio = srcWords.length > 0 ? exactWordMatches / srcWords.length : 0;

  // Perfect word match: ALL source words found in target AND target isn't bloated
  if (wordMatchRatio === 1 && srcWords.length > 0 && lenRatio > 0.4) {
    return { score: 0.90 + 0.10 * lenRatio, reason: 'exact' };
  }

  // Reverse check: ALL target words found in source (e.g. target "FIRST" ⊂ source "first name")
  const reverseMatches = tgtWords.filter(tw => srcWords.some(sw => sw === tw)).length;
  const reverseRatio = tgtWords.length > 0 ? reverseMatches / tgtWords.length : 0;
  if (reverseRatio === 1 && tgtWords.length > 0 && lenRatio > 0.3) {
    return { score: 0.85 + 0.10 * lenRatio, reason: 'contains' };
  }

  // Single-word source appears as a leading/key word in a multi-word target
  // e.g. "State" matches "STATE OF REGISTRATION" — "state" is the first word
  if (srcWords.length === 1 && tgtWords.length > 0) {
    const sw = srcWords[0];
    if (tgtWords[0] === sw) {
      // Source word is the FIRST word of target — strong signal
      return { score: 0.80, reason: 'contains' };
    }
    if (tgtWords.includes(sw)) {
      // Source word appears somewhere in target
      return { score: 0.65, reason: 'suggested' };
    }
  }

  // Partial word match: penalize if key distinctive words are MISSING
  // e.g. "First Name" has words [first, name]; "LAST NAME" has [last, name]
  // "first" is missing from target → not a good match even though "name" matches
  if (wordMatchRatio < 1 && srcWords.length > 0) {
    // How many source words are NOT in target? That's a disqualifier.
    const missingWords = srcWords.filter(sw => !tgtWords.some(tw => tw === sw));
    const missingPenalty = missingWords.length / srcWords.length;
    // If more than half the source words are missing, cap the score low
    if (missingPenalty >= 0.5) {
      const capped = Math.min(normSim, containSim, 0.5) * (1 - missingPenalty * 0.5);
      return { score: capped, reason: 'weak' };
    }
  }

  const best = Math.max(normSim, containSim);
  const adjusted = best * (0.5 + 0.5 * lenRatio);

  const boost = patternBoost(sourceNorm, targetNorm);
  if (boost === 1) return { score: Math.max(adjusted, 0.8), reason: 'pattern' };
  if (adjusted > 0.88) return { score: adjusted, reason: 'fuzzy' };
  if (adjusted > 0.7 && containSim > 0.7 && lenRatio > 0.4) return { score: adjusted, reason: 'contains' };
  return { score: adjusted, reason: adjusted > 0.55 ? 'suggested' : 'weak' };
}

// Score a source header against a target field, expanding through credentialing
// aliases and requiring distinctive context tokens so generic repeated PDF
// fields (ADDRESS_2, EFFECTIVE DATE, etc.) don't win by accident.
function scoreWithAliases(canonNorm, targetNorm) {
  let best = { ...scoreMatch(canonNorm, targetNorm), alias: null };
  const aliases = HEADER_ALIASES[canonNorm];
  if (aliases) {
    for (const alias of aliases) {
      const r = scoreMatch(normalizeKey(alias), targetNorm);
      // Keep a STABLE reason enum ('alias') so the UI's automation whitelist
      // never silently drops these; expose the matched phrase separately for
      // display only. Never return dynamic reason strings the frontend can't
      // recognize.
      if (r.score > best.score) best = { score: r.score, reason: 'alias', alias };
    }
  }
  // Context-token gating for fields that share generic words across the form.
  // If the canonical key carries a distinctive qualifier, the target must too.
  const CONTEXT = {
    'individual npi': ['individual', 'type 1', 'national provider identifier'],
    'group npi': ['group', 'type 2', 'organization'],
    'license effective date': ['license'],
    'license expiration date': ['license'],
    'board certification effective date': ['board', 'certif'],
    'board certification expiration date': ['board', 'certif'],
  };
  const ctx = CONTEXT[canonNorm];
  if (ctx && best.score >= 0.55) {
    const hasCtx = ctx.some((t) => targetNorm.includes(t));
    if (!hasCtx) best = { score: best.score * 0.45, reason: 'weak:no-context' };
  }
  return best;
}

function smartSuggestMappings(sourceHeaders, targetFields) {
  const sourceNorm = sourceHeaders.map((h) => ({ original: h, norm: canonicalizeHeader(h) }));
  const targetNorm = targetFields.map((f) => ({ ...f, norm: normalizeKey(f.name) }));
  const suggestions = [];
  const usedTargets = new Set();
  // Pass 1: exact / strong
  for (const s of sourceNorm) {
    let best = null;
    for (const t of targetNorm) {
      if (usedTargets.has(t.name)) continue;
      const { score, reason, alias } = scoreWithAliases(s.norm, t.norm);
      if (!best || score > best.score) best = { source: s.original, target: t.name, score, reason, alias };
    }
    if (best && best.score >= 0.7) {
      suggestions.push(best);
      usedTargets.add(best.target);
    }
  }
  // Pass 2: remaining source headers with lower threshold
  for (const s of sourceNorm) {
    if (suggestions.some((x) => x.source === s.original)) continue;
    let best = null;
    for (const t of targetNorm) {
      if (usedTargets.has(t.name)) continue;
      const { score, reason, alias } = scoreWithAliases(s.norm, t.norm);
      if (!best || score > best.score) best = { source: s.original, target: t.name, score, reason, alias };
    }
    if (best && best.score >= 0.55) {
      suggestions.push(best);
      usedTargets.add(best.target);
    } else if (best) {
      suggestions.push({ source: s.original, target: best.target, score: best.score, reason: 'weak', needsReview: true });
    }
  }
  return suggestions;
}

function saveProject(project) {
  const id = project.id || nanoid();
  const filePath = path.join(PROJECTS_DIR, `${id}.json`);
  const payload = { ...project, id, updatedAt: new Date().toISOString(), createdAt: project.createdAt || new Date().toISOString() };
  fs.writeFile(filePath, JSON.stringify(payload, null, 2));
  return id;
}

function listProjects() {
  return fs.readdir(PROJECTS_DIR).then((files) =>
    Promise.all(
      files
        .filter((f) => f.endsWith('.json'))
        .map(async (f) => {
          const full = path.join(PROJECTS_DIR, f);
          const raw = await fs.readFile(full, 'utf8');
          return JSON.parse(raw);
        })
    )
  );
}

// ---------- Routes ----------

app.get('/api/health', (req, res) => res.json({ status: 'ok' }));

app.post('/api/upload/pdf', upload.single('pdf'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'PDF file is required' });
  try {
    const fields = await parsePdfFormFields(req.file.path);
    await fs.unlink(req.file.path).catch(() => undefined);
    res.json({ fields, filename: req.file.originalname });
  } catch (err) {
    await fs.unlink(req.file.path).catch(() => undefined);
    res.status(500).json({ error: 'Failed to parse PDF form fields', message: err.message });
  }
});

app.post('/api/upload/data', upload.single('data'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'Data file is required' });
  try {
    const { headers, rows } = await parseDataFile(req.file.path);
    await fs.unlink(req.file.path).catch(() => undefined);
    res.json({ headers, rowCount: rows.length, sample: rows.slice(0, 5), filename: req.file.originalname });
  } catch (err) {
    await fs.unlink(req.file.path).catch(() => undefined);
    res.status(500).json({ error: 'Failed to parse data file', message: err.message });
  }
});

app.post('/api/mapping/suggest', (req, res) => {
  try {
    const { sourceHeaders, pdfFields } = req.body;
    if (!sourceHeaders || !pdfFields) return res.status(400).json({ error: 'sourceHeaders and pdfFields are required' });
    const suggestions = smartSuggestMappings(sourceHeaders, pdfFields);
    const unmappedPdf = pdfFields.filter((f) => !suggestions.some((s) => s.target === f.name)).map((f) => ({ name: f.name, type: f.type, page: f.page }));
    res.json({ suggestions, unmappedPdf });
  } catch (err) {
    res.status(500).json({ error: 'Failed to compute mapping suggestions', message: err.message });
  }
});

app.post('/api/preview', upload.single('pdf'), async (req, res) => {
  try {
    const { mapping, rowIndex, data } = req.body;
    const parsedData = typeof data === 'string' ? JSON.parse(data) : data;
    const mappingObj = typeof mapping === 'string' ? JSON.parse(mapping) : mapping;
    if (!req.file) return res.status(400).json({ error: 'PDF template is required' });
    const pdfBytes = await fs.readFile(req.file.path);
    const pdfDoc = await PDFDocument.load(pdfBytes);
    const form = pdfDoc.getForm();
    const fillReport = [];
    const mappingEntries = Object.entries(mappingObj || state.mapping || {});
    for (const [pdfFieldName, sourceKey] of mappingEntries) {
      const value = String(parsedData?.[sourceKey] ?? '');
      try {
        let field;
        try {
          field = form.getField(pdfFieldName);
        } catch {
          field = (form.getFields() || []).find((candidate) => typeof candidate.getName === 'function' && candidate.getName().toLowerCase() === String(pdfFieldName).toLowerCase()) || null;
        }
        if (!field) throw new Error('field not found');
        let written = false;
        if (field.constructor?.name === 'PDFTextField') {
          if (value.trim()) { field.setText(value); written = true; }
        } else if (field.constructor?.name === 'PDFCheckBox') {
          if (value.trim()) { field.check(); written = true; }
        } else if (field.constructor?.name === 'PDFRadioGroup') {
          try { field.select(value); } catch { field.select(value); }
          written = true;
        }
        fillReport.push({ sourceKey, pdfFieldName, value, written });
      } catch (err) {
        fillReport.push({ sourceKey, pdfFieldName, value, written: false, error: err?.message || 'field not found' });
      }
    }
    const filled = await pdfDoc.save();
    const outPath = path.join(OUTPUT_DIR, `preview_${Date.now()}.pdf`);
    await fs.writeFile(outPath, filled);
    await fs.unlink(req.file.path).catch(() => undefined);
    res.json({ previewUrl: `/output/${path.basename(outPath)}`, rowIndex: Number(rowIndex) || 0, fillReport });
  } catch (err) {
    await fs.unlink(req.file.path).catch(() => undefined);
    res.status(500).json({ error: 'Failed to generate preview', message: err.message });
  }
});

app.post('/api/export', upload.single('pdf'), async (req, res) => {
  try {
    const { mapping, rows, filenameTemplate } = req.body;
    const dataRows = typeof rows === 'string' ? JSON.parse(rows) : rows;
    const template = typeof filenameTemplate === 'string' ? filenameTemplate.replace(/[^a-zA-Z0-9_\-\.\(\)\{\}]/g, '_') : 'output.pdf';
    const parsedMapping = typeof mapping === 'string' ? JSON.parse(mapping) : mapping;
    if (!req.file) return res.status(400).json({ error: 'PDF template is required' });
    const templateBytes = await fs.readFile(req.file.path);
    const zipName = `filled_${Date.now()}.zip`;
    const zipPath = path.join(OUTPUT_DIR, zipName);
    const output = createWriteStream(zipPath);
    const archive = archiver('zip', { zlib: { level: 9 } });
    archive.pipe(output);
    const good = [];
    const bad = [];
    for (let i = 0; i < dataRows.length; i++) {
      const record = dataRows[i];
      try {
        const pdfDoc = await PDFDocument.load(templateBytes);
        const form = pdfDoc.getForm();
        for (const [pdfFieldName, sourceKey] of Object.entries(parsedMapping || {})) {
          const value = record?.[sourceKey] ?? '';
          try {
            let field;
            try {
              field = form.getField(pdfFieldName);
            } catch {
              field = (form.getFields() || []).find(
                (c) => typeof c.getName === 'function' && c.getName().toLowerCase() === String(pdfFieldName).toLowerCase()
              ) || null;
            }
            if (!field) continue;
            if (field.constructor?.name === 'PDFTextField') {
              field.setText(String(value));
            } else if (field.constructor?.name === 'PDFCheckBox') {
              if (String(value).trim()) field.check();
            } else if (field.constructor?.name === 'PDFRadioGroup') {
              try { field.select(String(value)); } catch { /* skip */ }
            }
          } catch { /* skip missing */ }
        }
        const filled = await pdfDoc.save();
        const fileName = template
          .replace(/\{rowIndex\}/gi, String(i + 1))
          .replace(/\{([^}]+)\}/g, (_, k) => String(record?.[k] ?? 'unknown'))
          .replace(/[^a-zA-Z0-9_\-\.]/g, '_');
        archive.append(Buffer.from(filled), { name: fileName || `record_${i + 1}.pdf` });
        good.push({ row: i + 1, file: fileName });
      } catch (err) {
        bad.push({ row: i + 1, error: err.message });
      }
    }
    await archive.finalize();
    await new Promise((resolve, reject) => {
      output.on('close', resolve);
      output.on('error', reject);
    });
    await fs.unlink(req.file.path).catch(() => undefined);
    res.json({ zipUrl: `/output/${zipName}`, good: good.length, bad: bad.length, failed: bad });
  } catch (err) {
    await fs.unlink(req.file.path).catch(() => undefined);
    res.status(500).json({ error: 'Export failed', message: err.message });
  }
});

app.get('/output/:file', async (req, res) => {
  const filePath = path.join(OUTPUT_DIR, req.params.file);
  if (!(await fs.stat(filePath).catch(() => null))) return res.status(404).json({ error: 'Not found' });
  res.sendFile(filePath);
});

app.get('/api/projects', async (req, res) => {
  const projects = await listProjects();
  res.json({ projects });
});

app.post('/api/projects', async (req, res) => {
  const project = req.body || {};
  const id = saveProject(project);
  res.json({ id });
});

app.get('/api/projects/:id', async (req, res) => {
  const project = await readJsonSafe(path.join(PROJECTS_DIR, `${req.params.id}.json`));
  if (!project) return res.status(404).json({ error: 'Project not found' });
  res.json(project);
});

const port = Number(process.env.PORT || 3000);

app.listen(port, () => {
  console.log(`Bulk PDF app running on port ${port}`);
});
