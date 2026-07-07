import { PDFDocument, PDFName } from 'pdf-lib';
import fs from 'fs';

const filePath = '/home/rex/.hermes/cache/documents/doc_baacaa44782e_lhl234.pdf';
const pdfBytes = fs.readFileSync(filePath);
const pdfDoc = await PDFDocument.load(pdfBytes);

const pageCount = pdfDoc.getPages().length;
console.log('Pages:', pageCount);

const form = pdfDoc.getForm();
console.log('Form exists:', !!form);
console.log('Top-level fields:', form.getFields().length);

const total = { text: 0, checkbox: 0, radio: 0, unknown: 0 };
const samples = [];

function walkFields(fields, pageIndex) {
  for (const field of fields || []) {
    let type = 'unknown';
    let name = null;
    try { name = field.getName(); } catch {}
    try {
      if (field.constructor?.name === 'PDFTextField') type = 'text';
      else if (field.constructor?.name === 'PDFCheckBox') type = 'checkbox';
      else if (field.constructor?.name === 'PDFRadioGroup') type = 'radio';
      else type = field.constructor?.name || 'unknown';
    } catch {
      type = 'unknown';
    }
    total[type] = (total[type] || 0) + 1;
    if (samples.length < 30 && name) samples.push({ page: pageIndex + 1, name, type });
    try {
      if (field.getChildren && typeof field.getChildren === 'function') {
        walkFields(field.getChildren(), pageIndex);
      }
    } catch {}
  }
}

for (let i = 0; i < pageCount; i++) {
  try { walkFields(form.getPageFields(i), i); } catch (e) { /* page may not expose form fields */ }
}

console.log('Walked totals:', total);
console.log('Samples:', samples);

const acroForm = pdfDoc.catalog.get(PDFName.of('AcroForm'));
console.log('AcroForm present:', !!acroForm);
