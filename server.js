import express from 'express';
import cors from 'cors';
import multer from 'multer';
import { promises as fs, createWriteStream, readFileSync } from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { PDFDocument, rgb } from 'pdf-lib';
import XLSX from 'xlsx';
import ExcelJS from 'exceljs';
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
app.use(express.json({ limit: '10mb' }));
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

const FIELD_MAP_SHEET = '__PDF_FIELD_MAP';

// Detects an ACTUAL checkbox field whose name ends in a Yes/No suffix, e.g.
// "BOARD CERTIFIED Yes_44" -> { base: "BOARD CERTIFIED", answer: "yes" }.
// Only ever call this on PDFCheckBox fields — never on text fields that merely
// contain the words yes/no.
function parseYesNoCheckboxName(name) {
  const m = String(name).match(/^(.*?)(?:[\s_-]+)(Yes|No)(?:[_\s-]*\d+)?$/i);
  if (!m) return null;
  return { base: m[1].trim().replace(/\s+/g, ' '), answer: m[2].toLowerCase() };
}

function normalizeYesNo(value) {
  const v = String(value ?? '').trim().toLowerCase();
  if (['yes', 'y', 'true', '1', 'checked', 'check', 'x', '☑', '✓'].includes(v)) return 'yes';
  if (['no', 'n', 'false', '0', 'unchecked', 'off', '☐'].includes(v)) return 'no';
  return '';
}

function isTruthyCheckbox(value) {
  return ['yes', 'y', 'true', '1', 'checked', 'check', 'x', '☑', '✓'].includes(
    String(value ?? '').trim().toLowerCase()
  );
}

// Decides what to do with a checkbox given a cell value and the TARGET PDF
// field name. Field-awareness matters: a literal "No" mapped onto a field whose
// own name is the No-side ("... No_45") should CHECK that box, whereas "No" on a
// standalone logical checkbox means leave it unchecked. Empty => no action.
function checkboxAction(value, pdfFieldName = '') {
  const v = String(value ?? '').trim().toLowerCase();
  if (!v) return null;

  const target = parseYesNoCheckboxName(pdfFieldName);
  const literalAnswer = ['yes', 'y'].includes(v) ? 'yes' : (['no', 'n'].includes(v) ? 'no' : '');

  // Literal Yes/No aimed at one side of a Yes/No checkbox pair: check the box
  // only when the answer matches that box's own side.
  if (target && literalAnswer) {
    return literalAnswer === target.answer ? 'check' : 'uncheck';
  }

  if (['no', 'n', 'false', '0', 'unchecked', 'off', 'blank', 'none', '☐'].includes(v)) return 'uncheck';
  return 'check';
}

// Builds the template spec from a fillable PDF: a list of VISIBLE Excel columns
// plus the hidden mapping rows that let us expand each visible answer back into
// exact PDF field name(s) at import time. This is the contract that keeps the
// fill/export round-trip deterministic after collapsing checkbox pairs.
function trailingNumber(name) {
  const m = String(name).match(/_+\s*(\d+)\s*$/);
  return m ? Number(m[1]) : null;
}

function normalizeCheckboxBase(base) {
  return String(base || '')
    .toLowerCase()
    .replace(/[?.,:;()\[\]{}]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

// Pairs Yes/No checkboxes SAFELY even when the same base label repeats across
// sections. Base-only grouping is unsafe here (this PDF repeats labels like
// "BOARD CERTIFIED"), so we pair by same normalized base + nearest opposite
// answer using trailing-suffix proximity then field-order proximity. Ambiguous
// checkboxes are left standalone rather than mis-paired.
function detectYesNoPairs(fields) {
  const infos = fields
    .map((f, fieldIndex) => ({ f, fieldIndex, name: f.getName(), type: f.constructor?.name }))
    .filter((x) => x.type === 'PDFCheckBox')
    .map((x, cbIndex) => {
      const p = parseYesNoCheckboxName(x.name);
      return p ? { ...x, cbIndex, base: p.base, normBase: normalizeCheckboxBase(p.base), answer: p.answer, suffix: trailingNumber(x.name) } : null;
    })
    .filter(Boolean);

  const candidates = [];
  for (let i = 0; i < infos.length; i++) {
    for (let j = i + 1; j < infos.length; j++) {
      const a = infos[i], b = infos[j];
      if (a.normBase !== b.normBase || a.answer === b.answer) continue;
      const suffixDist = a.suffix != null && b.suffix != null ? Math.abs(a.suffix - b.suffix) : 9999;
      const orderDist = Math.abs(a.cbIndex - b.cbIndex);
      // Suffix proximity is the strong signal. When BOTH sides carry a trailing
      // number, require the suffixes to be close (prevents pairing repeated
      // labels across sections). Fall back to field-order proximity only when a
      // suffix is missing on one/both sides.
      const bothHaveSuffix = a.suffix != null && b.suffix != null;
      // Credentialing forms: a wrong checkbox is far worse than an extra
      // standalone column. When both sides carry a trailing number, demand
      // adjacency (dist <= 1). Only fall back to field-order proximity when a
      // suffix is missing. Anything looser stays standalone.
      const clearlyPaired = bothHaveSuffix ? suffixDist <= 1 : orderDist <= 3;
      if (clearlyPaired) {
        candidates.push({
          yes: a.answer === 'yes' ? a : b,
          no: a.answer === 'no' ? a : b,
          base: a.answer === 'yes' ? a.base : b.base,
          score: suffixDist * 10 + orderDist,
        });
      }
    }
  }
  // Greedy: best (lowest) scores first, each checkbox used at most once.
  candidates.sort((x, y) => x.score - y.score);
  const used = new Set();
  const pairs = [];
  for (const c of candidates) {
    if (used.has(c.yes.name) || used.has(c.no.name)) continue;
    used.add(c.yes.name);
    used.add(c.no.name);
    pairs.push({ base: c.base, yes: c.yes.name, no: c.no.name });
  }
  return { pairs, handledNames: used };
}

// Visual sort key from a field's earliest widget rectangle, so Excel columns
// follow the order fields appear on the FORM (page -> top-to-bottom ->
// left-to-right) rather than AcroForm internal field order.
function resolveWidgetPageIndex(widget, pageIndexByObj, pageIndexByRefKey) {
  try {
    // Prefer identifying the widget itself against the per-page annotation maps
    // (most reliable); only then fall back to pdf-lib's flaky widget.P().
    if (pageIndexByObj.has(widget)) return pageIndexByObj.get(widget);
    if (widget.dict && pageIndexByObj.has(widget.dict)) return pageIndexByObj.get(widget.dict);

    const widgetRef = widget.ref || widget.dict?.ref;
    if (widgetRef != null) {
      const byWidgetRef = pageIndexByRefKey.get(String(widgetRef));
      if (byWidgetRef != null) return byWidgetRef;
    }

    const pObj = widget.P?.();
    if (pObj != null) {
      if (pageIndexByObj.has(pObj)) return pageIndexByObj.get(pObj);
      if (pObj.ref != null) {
        const byPageRef = pageIndexByRefKey.get(String(pObj.ref));
        if (byPageRef != null) return byPageRef;
      }
      const direct = pageIndexByRefKey.get(String(pObj));
      if (direct != null) return direct;
    }
  } catch { /* ignore */ }
  return -1;
}

function visualKeyForField(field, fallbackIndex, pageIndexByObj, pageIndexByRefKey) {
  let best = null;
  try {
    const widgets = field.acroField?.getWidgets?.() || [];
    for (const w of widgets) {
      let rect;
      try { rect = w.getRectangle(); } catch { continue; }
      if (!rect) continue;
      const page = resolveWidgetPageIndex(w, pageIndexByObj, pageIndexByRefKey);
      // Without a resolved page we cannot trust cross-page rect ordering, so
      // skip this widget rather than mixing pages by raw Y/X.
      if (page < 0) continue;
      const top = rect.y + rect.height; // PDF origin is bottom-left; higher y = higher on page
      const cand = { page, top, x: rect.x, y: rect.y, fallback: fallbackIndex, resolved: true };
      if (!best || compareVisualKeys(cand, best) < 0) best = cand;
    }
  } catch { /* fall through to fallback */ }
  // No resolvable widget/page: sort AFTER positioned fields, preserving the
  // original PDF field order among such fallbacks.
  return best || { page: Number.MAX_SAFE_INTEGER, top: 0, x: 0, y: 0, fallback: fallbackIndex, resolved: false };
}

const ROW_TOLERANCE = 5; // PDF points; fields within this vertical band = same row
function compareVisualKeys(a, b) {
  if (a.page !== b.page) return a.page - b.page;
  if (Math.abs(a.top - b.top) > ROW_TOLERANCE) return b.top - a.top; // higher first
  if (a.x !== b.x) return a.x - b.x; // same row: left to right
  return a.fallback - b.fallback; // stable tie-break
}

function minVisualKey(a, b) {
  if (!a) return b;
  if (!b) return a;
  return compareVisualKeys(a, b) <= 0 ? a : b;
}

function buildTemplateSpec(pdfDoc, form) {
  const fields = form.getFields();

  // Map pages -> index via BOTH the page dict object and stringified refs.
  // widget.P() is not always populated depending on the PDF, so we ALSO map
  // every annotation ref on each page to that page index — the most robust way
  // to place a widget on its page.
  const pageIndexByObj = new Map();
  const pageIndexByRefKey = new Map();
  pdfDoc.getPages().forEach((p, i) => {
    if (p.node) pageIndexByObj.set(p.node, i);
    if (p.ref) pageIndexByRefKey.set(String(p.ref), i);
    try {
      const annots = p.node?.Annots?.();
      if (annots) {
        for (let a = 0; a < annots.size(); a++) {
          const annotRef = annots.get(a);
          pageIndexByRefKey.set(String(annotRef), i);
          const annotObj = pdfDoc.context.lookup(annotRef);
          if (annotObj) pageIndexByObj.set(annotObj, i);
        }
      }
    } catch { /* best effort */ }
  });

  // Precompute a visual key per field (by name) and original index.
  const keyByName = new Map();
  fields.forEach((f, i) => {
    const nm = f.getName();
    keyByName.set(nm, visualKeyForField(f, i, pageIndexByObj, pageIndexByRefKey));
  });

  // Safe proximity-aware Yes/No pairing (handles repeated base labels).
  const { pairs, handledNames } = detectYesNoPairs(fields);
  const pairByName = new Map();
  for (const p of pairs) { pairByName.set(p.yes, p); pairByName.set(p.no, p); }

  const usedHeaders = new Set();
  const uniqueHeader = (base) => {
    let h = base, i = 2;
    while (usedHeaders.has(h.toLowerCase())) h = `${base} (${i++})`;
    usedHeaders.add(h.toLowerCase());
    return h;
  };

  // Build provisional column specs (baseHeader + sortKey), WITHOUT unique
  // headers yet — headers must be assigned AFTER sorting so (2) suffixes follow
  // visual order.
  const pending = [];
  const emittedPairs = new Set();
  for (const f of fields) {
    const name = f.getName();
    const pair = pairByName.get(name);
    if (pair) {
      const key = `${pair.yes}||${pair.no}`;
      if (!emittedPairs.has(key)) {
        emittedPairs.add(key);
        pending.push({
          baseHeader: `${pair.base} (Yes/No)`,
          kind: 'yesno',
          yesFieldName: pair.yes,
          noFieldName: pair.no,
          // Collapsed pair sits at the earlier of its two widgets.
          sortKey: minVisualKey(keyByName.get(pair.yes), keyByName.get(pair.no)),
        });
      }
      continue;
    }
    if (f.constructor?.name === 'PDFCheckBox') {
      pending.push({ baseHeader: `${name} (Checked?)`, kind: 'checkbox', fieldName: name, sortKey: keyByName.get(name) });
    } else {
      pending.push({ baseHeader: name, kind: 'field', fieldName: name, sortKey: keyByName.get(name) });
    }
  }

  // Sort by visual position, then assign unique headers in that final order.
  pending.sort((a, b) => compareVisualKeys(a.sortKey, b.sortKey));

  const columns = pending.map((c) => {
    const base = { header: uniqueHeader(c.baseHeader), kind: c.kind };
    if (c.kind === 'yesno') { base.yesFieldName = c.yesFieldName; base.noFieldName = c.noFieldName; }
    else base.fieldName = c.fieldName;
    return base;
  });

  return { columns, pairCount: pairs.length, fieldCount: fields.length };
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
    const dataSheet = workbook.Sheets[workbook.SheetNames[0]];
    // Extract headers from row 1 INDEPENDENTLY of data rows, so a header-only
    // template (generated by /api/pdf/template-xlsx) still reports its columns
    // instead of looking empty.
    const headerRows = XLSX.utils.sheet_to_json(dataSheet, { header: 1, blankrows: false, raw: false, defval: '' });
    const headers = (headerRows[0] || []).map((h) => String(h ?? '').trim()).filter(Boolean);
    // raw:false + dateNF renders Excel date serials as mm/dd/yyyy instead of
    // leaking numbers like 39700. defval keeps blank cells as ''.
    const data = XLSX.utils.sheet_to_json(dataSheet, { defval: '', raw: false, dateNF: 'mm/dd/yyyy' });
    // Trim stray whitespace (e.g. Medicare number stored with trailing spaces).
    for (const row of data) {
      for (const k of Object.keys(row)) {
        if (typeof row[k] === 'string') row[k] = row[k].trim();
      }
    }

    // If this workbook was generated by our PDF-first flow, expand each visible
    // logical column back into exact PDF field name(s) so the downstream
    // mapping/fill logic sees real field names and stays deterministic.
    if (workbook.SheetNames.includes(FIELD_MAP_SHEET)) {
      const mapRows = XLSX.utils.sheet_to_json(workbook.Sheets[FIELD_MAP_SHEET], { defval: '' });
      const specByHeader = new Map();
      for (const r of mapRows) {
        const h = String(r.columnHeader ?? '').trim();
        if (h) specByHeader.set(h, r);
      }

      const expandedHeaders = [];
      const seen = new Set();
      const pushHeader = (name) => { if (name && !seen.has(name)) { seen.add(name); expandedHeaders.push(name); } };

      const expandRow = (row) => {
        const out = {};
        for (const [visibleHeader, spec] of specByHeader.entries()) {
          const raw = row[visibleHeader];
          const kind = String(spec.kind || 'field');
          if (kind === 'yesno') {
            const answer = normalizeYesNo(raw);
            // Field-neutral control markers: checkboxAction() maps 'checked' ->
            // check and 'unchecked' -> uncheck regardless of the target field's
            // own Yes/No side. Never emit 'Yes' here — field-aware checkboxAction
            // would misread a 'Yes' aimed at a No-side field as an uncheck.
            if (answer === 'yes') {
              out[spec.yesFieldName] = 'checked';
              out[spec.noFieldName] = 'unchecked';
            } else if (answer === 'no') {
              out[spec.yesFieldName] = 'unchecked';
              out[spec.noFieldName] = 'checked';
            } else {
              out[spec.yesFieldName] = '';
              out[spec.noFieldName] = '';
            }
          } else if (kind === 'checkbox') {
            out[spec.fieldName] = isTruthyCheckbox(raw) ? 'checked' : (String(raw ?? '').trim() ? 'unchecked' : '');
          } else {
            out[spec.fieldName] = raw ?? '';
          }
        }
        return out;
      };

      // Build expanded header order from the spec (data-independent).
      for (const spec of specByHeader.values()) {
        if (String(spec.kind) === 'yesno') { pushHeader(spec.yesFieldName); pushHeader(spec.noFieldName); }
        else pushHeader(spec.fieldName);
      }

      const expandedRows = data.map(expandRow);
      return { headers: expandedHeaders, rows: expandedRows };
    }

    return { headers: headers.length ? headers : (data.length ? Object.keys(data[0]) : []), rows: data };
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

// PDF-first flow: upload a fillable PDF, get back a blank .xlsx. Checkbox
// Yes/No pairs are condensed to ONE column each; a hidden __PDF_FIELD_MAP sheet
// records how each visible column expands back to exact PDF field name(s) so
// the fill/export round-trip stays deterministic.
app.post('/api/pdf/template-xlsx', upload.single('pdf'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'PDF file is required' });
  try {
    const bytes = await fs.readFile(req.file.path);
    const pdfDoc = await PDFDocument.load(bytes);
    const form = pdfDoc.getForm();
    const spec = buildTemplateSpec(pdfDoc, form);

    if (!spec.columns.length) {
      await fs.unlink(req.file.path).catch(() => undefined);
      return res.status(400).json({ error: 'No fillable fields found in this PDF.' });
    }
    if (spec.columns.length > 16384) {
      await fs.unlink(req.file.path).catch(() => undefined);
      return res.status(400).json({ error: `This PDF maps to ${spec.columns.length} columns, exceeding Excel's 16,384 column limit.` });
    }

    // Visible data sheet: one header row of logical column names.
    const headers = spec.columns.map((c) => c.header);

    const excelCol = (n) => {
      let s = '';
      while (n > 0) {
        const m = (n - 1) % 26;
        s = String.fromCharCode(65 + m) + s;
        n = Math.floor((n - 1) / 26);
      }
      return s;
    };

    const wb = new ExcelJS.Workbook();
    wb.creator = 'Bulk PDF App';
    wb.created = new Date();

    const ws = wb.addWorksheet('PDF Fields Template', {
      views: [{ state: 'frozen', ySplit: 1 }],
    });

    ws.addRow(headers);
    ws.getRow(1).font = { bold: true };
    ws.getRow(1).alignment = { vertical: 'middle', wrapText: true };

    headers.forEach((h, idx) => {
      ws.getColumn(idx + 1).width = Math.min(Math.max(String(h).length + 2, 12), 60);
    });

    ws.autoFilter = {
      from: { row: 1, column: 1 },
      to: { row: 1, column: headers.length },
    };

    // Checkbox-derived columns get a Yes/No dropdown (data validation) instead
    // of a raw text cell, so coordinators pick a value rather than type one.
    const yesNoValidation = {
      type: 'list',
      allowBlank: true,
      formulae: ['"Yes,No"'],
      showErrorMessage: true,
      errorStyle: 'stop',
      errorTitle: 'Invalid checkbox value',
      error: 'Choose Yes or No.',
      showInputMessage: true,
      promptTitle: 'Checkbox field',
      prompt: 'Choose Yes or No.',
    };

    for (let i = 0; i < spec.columns.length; i++) {
      const c = spec.columns[i];
      if (c.kind === 'yesno' || c.kind === 'checkbox') {
        const letter = excelCol(i + 1);
        ws.dataValidations.add(`${letter}2:${letter}1048576`, yesNoValidation);
      }
    }

    // Hidden metadata sheet: version + per-column expansion recipe.
    const mapRows = [['version', 'columnHeader', 'kind', 'fieldName', 'yesFieldName', 'noFieldName']];
    for (const c of spec.columns) {
      mapRows.push([1, c.header, c.kind, c.fieldName || '', c.yesFieldName || '', c.noFieldName || '']);
    }

    const wsMap = wb.addWorksheet(FIELD_MAP_SHEET);
    wsMap.addRows(mapRows);
    wsMap.state = 'hidden';
    wsMap.getRow(1).font = { bold: true };

    const buffer = Buffer.from(await wb.xlsx.writeBuffer());

    const base = path.basename(req.file.originalname, path.extname(req.file.originalname))
      .replace(/[^a-z0-9_-]+/gi, '_').slice(0, 80);
    await fs.unlink(req.file.path).catch(() => undefined);

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${base || 'pdf'}-fields-template.xlsx"`);
    res.setHeader('X-Field-Count', String(spec.fieldCount));
    res.setHeader('X-Column-Count', String(spec.columns.length));
    res.setHeader('X-Checkbox-Groups', String(spec.pairCount));
    res.send(buffer);
  } catch (err) {
    await fs.unlink(req.file.path).catch(() => undefined);
    res.status(500).json({ error: 'Failed to generate Excel template', message: err.message });
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
          const action = checkboxAction(value, pdfFieldName);
          if (action === 'check') { field.check(); written = true; }
          else if (action === 'uncheck') { field.uncheck(); written = true; }
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
              const action = checkboxAction(value, pdfFieldName);
              if (action === 'check') field.check();
              else if (action === 'uncheck') field.uncheck();
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
