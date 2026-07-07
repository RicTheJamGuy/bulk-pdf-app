# Bulk PDF App

A lightweight web app to fill PDF form templates in bulk from CSV, Excel, or JSON data.

Upload a blank PDF form, map data columns to form fields, preview one record, and export a ZIP with one filled PDF per row.

## Highlights

- Parse PDF AcroForm fields (text, checkbox, radio)
- Import data from `.csv`, `.xlsx` / `.xls`, or `.json`
- Smart mapping suggestions with confidence levels
- Preview fill output for a selected row
- Batch export to ZIP with customizable filename template
- Save and load project mapping configurations via API

## Architecture

- Backend: Node.js + Express ([server.js](server.js))
- Frontend: single-page HTML/CSS/JS ([public/index.html](public/index.html))
- PDF engine: `pdf-lib`
- Data parsing: `csv-parse`, `xlsx`
- Export packaging: `archiver`

## Quick Start

### Prerequisites

- Node.js 18+
- npm

### Install

```bash
npm install
```

### Run (development)

```bash
npm run dev
```

### Run (production-like)

```bash
npm start
```

Open: [http://localhost:3000](http://localhost:3000)

## Deploy On Render (Free Tier)

This project is configured as a single Render Web Service using [render.yaml](render.yaml).

### 1) Push your repo

Push this repository to GitHub (or GitLab) with the current `master` branch.

### 2) Create service from Blueprint

In Render:

1. New -> Blueprint
2. Select this repository
3. Confirm the detected service from `render.yaml`
4. Create service

Render will run:

- Build command: `npm ci`
- Start command: `npm start`
- Health check: `/api/health`

### 3) Environment variables

- `NODE_ENV=production` is set in `render.yaml`.
- `PORT` is injected by Render automatically.
- `CORS_ORIGIN` is optional.
  - Leave empty when using the built-in UI served by this same app.
  - Set it only if you host the frontend on a different domain, for example:
    - `https://your-frontend.example.com`

Use [.env.example](.env.example) as reference.

### 4) Free-tier storage behavior

This app writes runtime files to:

- `data/uploads/`
- `output/`
- `projects/`

On Render free tier, the local filesystem is ephemeral. Generated files and saved project configs can be lost on restart/redeploy.

If you need persistence, move project/output storage to external storage (for example S3, database, or object storage).

## App Workflow

1. Upload your blank PDF template.
2. Upload your data file (`CSV`, `XLSX/XLS`, or `JSON`).
3. Generate mapping suggestions and adjust manually if needed.
4. Preview a row to validate output.
5. Export all rows as a ZIP file.

## Filename Templating

The export step supports placeholders in the filename pattern.

- `{rowIndex}` -> 1-based row number
- `{field_name}` -> value from each data row

Example:

```text
{first_name}_{last_name}_{rowIndex}.pdf
```

Unknown placeholders are replaced with `unknown`.

## API Endpoints

| Method | Endpoint | Description |
| --- | --- | --- |
| `GET` | `/api/health` | Health check |
| `POST` | `/api/upload/pdf` | Upload PDF and extract form fields |
| `POST` | `/api/upload/data` | Upload data file and parse headers/sample |
| `POST` | `/api/mapping/suggest` | Generate smart field mapping suggestions |
| `POST` | `/api/preview` | Fill one row and return preview PDF |
| `POST` | `/api/export` | Fill all rows and return ZIP download link |
| `GET` | `/api/projects` | List saved project configs |
| `POST` | `/api/projects` | Save a project config |
| `GET` | `/api/projects/:id` | Get a saved project config |
| `GET` | `/output/:file` | Download generated output file |

## Project Structure

```text
bulk-pdf-app/
  data/
    uploads/        # temporary upload storage
  output/           # generated previews and ZIP files
  projects/         # saved project JSON configs
  public/
    index.html      # web UI
  templates/        # reserved template folder
  server.js         # API + static server
  inspect.mjs       # local PDF inspection helper script
```

## Notes

> [!IMPORTANT]
> The `output/` and `projects/` folders are git-ignored by default. Exported ZIPs and saved project mappings are local-only unless you change `.gitignore`.
> [!NOTE]
> The UI includes an "AI suggest (optional)" button, but there is currently no backend endpoint wired for AI mapping.
> [!TIP]
> For better auto-mapping, keep column names close to PDF field names (for example: `first_name` -> `First Name`).

## Troubleshooting

- Upload fails with unsupported format:
  - Use `.csv`, `.json`, `.xlsx`, or `.xls` for data uploads.
- Preview/export fills fewer fields than expected:
  - Check mapping choices and ensure PDF fields are form fields (AcroForm), not static text.
- Download link returns 404:
  - Confirm the file exists in `output/` and the server is still running.

## Tech Stack

- `express`, `cors`, `multer`
- `pdf-lib`
- `csv-parse`, `xlsx`
- `archiver`
- `nanoid`
