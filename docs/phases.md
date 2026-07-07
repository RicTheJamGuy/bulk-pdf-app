# Bulk PDF App — Product Roadmap

## Current State: v1 MVP (Live)

| Capability | Status |
|---|---|
| PDF AcroForm field detection | ✅ Shipped |
| Data import (CSV, XLSX/XLS, JSON) | ✅ Shipped |
| 5-level smart mapping (exact → normalized → fuzzy → category → pattern) | ✅ Shipped |
| Manual mapping with live preview | ✅ Shipped |
| Fill report in Step 4 | ✅ Shipped |
| Batch export to ZIP with filename templating (`{rowIndex}`, `{field_name}`) | ✅ Shipped |
| Project save/load API | ✅ Shipped |
| Single-user, no auth | ✅ Shipped |
| Local filesystem storage (ephemeral) | ✅ Shipped |

Live repo: https://github.com/RicTheJamGuy/bulk-pdf-app

---

## Phase 2 — Reliability & Scale

**Goal:** Make the app safe to run in production for daily workflows.

| Feature | Notes |
|---|---|
| Persistent storage (SQLite or Postgres) | Replace `data/uploads/`, `output/`, `projects/` with DB-backed storage so exports survive restarts |
| Auth (single-user + optional team seats) | Session-based login; encrypted API tokens |
| Audit log | Track who uploaded what, when, and how many records were exported |
| Retry + error reporting | Per-row failure reasons in export output so Nikki can fix bad data without rerunning the whole batch |
| Queue / worker process | Offload long-running PDF generation to a background worker so the web UI stays responsive on 10k-row files |
| Health + metrics endpoint | `/api/health` already exists; add request latency, queue depth, storage usage |

---

## Phase 3 — Better Matching

**Goal:** Reduce manual mapping work for complex or inconsistent PDFs.

| Feature | Notes |
|---|---|
| AI-assisted mapping (Level 5) | LLM-backed column-to-field matching as an opt-in enhancement when deterministic scoring is weak |
| Mapping templates / reuse | Save and share mapping profiles across similar forms (e.g., all Texas credentialing revisions) |
| Field normalization rules | User-defined aliases (`DOB` → `DATE OF BIRTH`) to improve deterministic scores without touching source Excel |
| Fuzzy match tuning | Per-form confidence slider; auto-accept threshold instead of hardcoded thresholds |
| Duplicate-field grouping | When a PDF has 15 variants of `STATE OF REGISTRATION_2`, present them as a single mapped group |

---

## Phase 4 — Advanced PDF Output

**Goal:** Handle forms that aren't pure AcroForm text fields.

| Feature | Notes |
|---|---|
| Image / signature embedding | Fill signature blocks and photo fields from data |
| Checkbox + radio semantics | Better handling of Yes/No, Male/Female, and mutually exclusive radio groups |
| Barcode / QR generation | Generate and embed barcodes where the form expects machine-readable output |
| PDF/A compliance mode | Produce archival-grade output for regulated submissions |
| Watermark / overlay mode | Stamp filled forms with "COMPLETED" and date for tracking |
| Multi-page merge | Merge supplemental pages (attachments A, B, C) conditionally based on data |

---

## Phase 5 — Multi-User & API

**Goal:** Let teams and other tools use BPA as a service.

| Feature | Notes |
|---|---|
| REST API with API keys | External tools can POST `/api/v1/fill` and receive ZIP without the UI |
| Role-based access | Admin, mapper, viewer — control who can upload templates vs. run exports |
| Webhook notifications | POST to a callback when a batch export completes |
| Embeddable widget | Drop-in iframe or JS SDK for insurance portals and intake workflows |
| Usage quotas | Per-user/month limits on PDFs generated, storage used |

---

## Phase 6 — Embedded / Hosted Offering

**Goal:** Make this a hosted SaaS option for clients who don't want to self-host.

| Feature | Notes |
|---|---|
| Multi-tenancy | Isolated data per tenant, custom branding |
| Stripe billing | Usage-based or seat-based billing via Lemon Squeezy or Stripe |
| Managed PDF templates | Nikki can upload and version templates inside the platform instead of re-uploading per batch |
| Integrations | Google Sheets, Airtable, HubSpot as data sources in addition to file upload |
| SLA + support tiers | For production credentialing workflows at scale |

---

## Phase 7 — Platform & Extensibility

**Goal:** Make BPA a home for all form-related automation.

| Feature | Notes |
|---|---|
| Plugin system | Community-built parsers for IRS W-9, NPDB, DEA forms, etc. |
| CLI + Docker image | `docker run bulk-pdf-app fill --template form.pdf --data data.csv --out out.zip` |
| GitHub Actions / CI integration | Auto-fill and archive PDFs on PR merge for document-heavy workflows |
| Observability | OpenTelemetry traces, Sentry error tracking, structured JSON logging |

---

## Out of Scope (v2 and beyond)

- OCR for scanned/image-only PDFs (pdfplumber/pdf.js can extract text, but filling requires AcroForm fields)
- Real-time collaborative mapping (multiple cursors editing the same mapping table)
- Mobile-native app (responsive web UI covers tablets; native app is not planned)

---

## Contributing

From Phase 2 onward, implementation is tracked via GitHub Issues and PRs on https://github.com/RicTheJamGuy/bulk-pdf-app.

Branch naming:
- `feat/phase-N-short-description`
- `fix/phase-N-short-description`

PRs should reference the phase in the description, e.g.:

```markdown
Phase: 2 — Reliability & Scale
Related Issue: #NNN
```
