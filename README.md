# AI Document Intelligence & Data Extraction Agent

A local-first, portfolio-grade Streamlit MVP for extracting structured suggestions from PDFs and images, reviewing them with a person, and maintaining a SQLite activity history. Deterministic extraction works without network access or an LLM.

**Live app:** https://documind-document-intelligence.streamlit.app/

## Portfolio fit

The adjacent Codex portfolio includes AI evaluation/security and governance applications, plus career and placement tools. Project 9 complements those by adding document ingestion, OCR, information extraction, human review, and structured exports. This project has its own code, database, and dependencies.

## Features

- PDF and common image uploads; 20 MB per file and 60 pages per PDF.
- Local text extraction from text PDFs and local Tesseract OCR for scanned pages and images.
- English and Hindi OCR options. Tesseract's `eng` and `hin` trained data are required for the corresponding choices.
- Heuristic classification: resume, certificate, invoice, form, or other.
- Starter field extraction for names, dates, amounts, qualifications and document-specific details; optional custom labels.
- Review flags point out missing or potentially inconsistent values. They are prompts to check source documents and do not assert fraud or authenticity.
- Editable human review, approval/needs-changes decisions, and SQLite audit events.
- CSV, Excel, and JSON exports.
- Two-way Excel workflow: export an editable review workbook, update structured fields and decisions in Microsoft Excel, then import the workbook to apply changes to matching SQLite records with an explicit confirmation step and audit event.
- **Ask your files:** session-only research-paper summaries and passage search with page citations, plus Excel worksheet previews, deterministic aggregates, and supporting rows. No external LLM is used by this feature.
- Optional OpenAI-compatible LLM adapter with configurable endpoint/model/key. It is off unless configured, a consent checkbox is selected, and the user presses the provider button. Never configure it for sensitive documents without an explicit organizational decision and review of provider data handling.
- A synthetic invoice PDF may be kept locally for testing; demo files, samples, local data, and work folders are excluded from Git uploads.

## Architecture

```text
Streamlit UI (app.py)
  ├── upload validation + PDF/image text pipeline (document_agent/extraction.py)
  │     ├── PyMuPDF text layer
  │     └── local Tesseract OCR fallback
  ├── deterministic classifier, field patterns, and review flags
  ├── optional explicitly invoked provider adapter (document_agent/providers.py)
  └── SQLite repository and audit events (document_agent/storage.py)
```

The app does not persist uploaded files. SQLite persists the original filename, SHA-256 digest, extracted values, flags, reviewer edits/status, timestamps, and audit events. A hash can still be linkable to a known source file; protect the database accordingly.

## Windows setup

Python 3.10 or newer is recommended.

1. Install [Tesseract OCR for Windows](https://github.com/UB-Mannheim/tesseract/wiki). During setup, include English and Hindi (`eng` and `hin`) language data. If Tesseract is not on `PATH`, set `TESSERACT_CMD` in your environment to its executable path; see note below.
2. Open PowerShell in this project folder:

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
streamlit run app.py
```

If PowerShell blocks activation, use `Set-ExecutionPolicy -Scope Process Bypass` for that terminal, then activate. (Or call `.venv\Scripts\python.exe -m pip ...` and `.venv\Scripts\streamlit.exe run app.py` without activation.)

On macOS install Tesseract and language data with the platform package manager (`tesseract`, `tesseract-lang`); on Debian/Ubuntu, `sudo apt install tesseract-ocr tesseract-ocr-eng tesseract-ocr-hin`. Then use the same Python virtual-environment steps with the platform activation command.

### Tesseract fallback and setup checks

OCR is performed with the local Tesseract executable through `pytesseract`. If not found, add Tesseract to `PATH`, restart the terminal, or configure `pytesseract.pytesseract.tesseract_cmd` in `app.py` from an environment setting such as `TESSERACT_CMD`. A missing language pack can be checked with `tesseract --list-langs`. If OCR cannot run, the app reports a local setup error; it does not silently upload pages to a cloud OCR service. Text-based PDF extraction still works without Tesseract.

## Optional LLM provider

Copy `.env.example` to `.env` and set `LLM_BASE_URL`, `LLM_API_KEY`, and `LLM_MODEL` for a provider that supports an OpenAI-compatible chat completions endpoint and JSON response format. The adapter accepts a configurable base URL rather than assuming a vendor. Provider compatibility varies. The app never sends document text on configuration alone: the user must explicitly tick the consent checkbox and press the request button. Provider output is only a suggestion and is not saved as approved data automatically. Avoid placing `.env` in source control.

## Run checks

```powershell
python -m pytest -q
```

Tests use synthetic input, a generated text PDF, and in-memory SQLite. They do not require a Tesseract installation or network API.

## Using the Excel workflow

1. Start the app with `streamlit run app.py` and open **Records & exports**.
2. Optionally filter the register, then select **Download Excel review workbook**. The workbook contains **Start here**, **Review**, and **Snapshot** sheets. Open it with Microsoft Excel.
3. On the **Review** sheet, edit the `field__...` columns with corrected values. Set `review_decision` to `pending`, `approved`, or `needs_changes`, and add context in `reviewer_note`. Keep `job_id` unchanged. Leave rows you do not want to update out of the sheet, or remove them from the Review sheet.
4. Save the workbook as `.xlsx`. Return to the same DocuMind database and upload it in **Bring reviewed rows back from Excel**.
5. Check the preview carefully, tick the confirmation box, and select **Apply Excel review updates**. Each imported row updates an existing job and creates an audit event. The importer rejects duplicate/unknown IDs and invalid decision values.
6. Open **Audit trail** to inspect imported updates. The **Export Excel snapshot** is a read-only-style data snapshot for reporting; use the dedicated **Excel review workbook** for the round trip.

The Excel file contains extracted personal data. Store and share it carefully, and verify every edit against the source document before approval. Import does not upload source files or create new jobs.

## Ask your files

Open **Ask your files** from the sidebar and choose **Research paper (PDF)** or **Spreadsheet (Excel)**. Files and extracted content stay in the active Streamlit session and are not saved in the SQLite document register. A server restart or new session clears them.

### Research papers

1. Upload a PDF (up to 20 MB / 200 pages), select OCR language, and choose **Read research paper**.
2. Use **Create paper summary** for an extractive summary built from labeled sections and key passages. Expand each result to see its page number.
3. Ask a question and choose **Search paper evidence**. The app ranks passages by word overlap and shows page references. Check the cited page and surrounding text; this starter retrieval does not synthesize a guaranteed answer or judge scientific validity.
4. Digital PDFs use embedded text. Scanned pages use local Tesseract. Install Tesseract and language data as described above; no cloud OCR fallback is used.

### Excel analysis

1. Upload an `.xlsx` workbook (up to 20 MB) and choose **Read spreadsheet**, or select **Try a synthetic sales example** to explore the feature using fictional data. The first row of each worksheet is treated as headers; the app reads up to 30 sheets.
2. Select the worksheet and inspect the preview before asking a question.
3. Ask a supported calculation such as **“Which month had the highest sales?”**, **“What is the total revenue?”**, or **“What is the average quantity?”**. The app shows the matched columns, calculation, and supporting rows when applicable.

Excel calculations are deterministic and local. They are useful for clean tabular data; they do not infer complex business definitions, handle every date grouping, or replace checking formulas, units, missing values, filters, and duplicate rows. For general questions, the app returns matching rows by simple word overlap and labels that as a search result. If it cannot find support, it says so.

## GitHub and Streamlit deployment

The recommended setup for keeping the source private is a **private GitHub repository** and a **private Streamlit Community Cloud app**. Invite only people who should use it. Streamlit Community Cloud provides private app access by invited email; private viewers do not get repository write access. Community Cloud currently limits an account to one private app. A public app URL can be used by anyone and is not appropriate when access should be restricted.

1. Create a private GitHub repository named `documind-document-intelligence`.
2. In PowerShell, open this project directory and run:

```powershell
git init
git add app.py document_agent .streamlit/config.toml requirements.txt packages.txt README.md .env.example .gitignore
git status --short
git commit -m "Build DocuMind document intelligence lab"
git branch -M main
git remote add origin https://github.com/sarajsinghal/documind-document-intelligence.git
git push -u origin main
```

3. Confirm `git status --short` shows no sample PDFs, `outputs/`, `samples/`, `data/`, `.venv/`, `.env`, or `work/` files before committing. `.gitignore` excludes these paths; the command above also stages only named application files.
4. Sign in at [Streamlit Community Cloud](https://share.streamlit.io/), connect GitHub, and select **Create app**.
5. Choose the private repository, `main` branch, and `app.py` entrypoint. In advanced settings, use Python 3.12. Streamlit installs Python dependencies from `requirements.txt` and Linux OCR packages from `packages.txt`.
6. In app settings, keep **Who can view this app** set to **Only specific people can view this app**. Add intended viewers by email. Keep repository access limited to you; anyone with repository write access can change the app's code.
7. After deployment, verify the app loads and try a synthetic spreadsheet example. Do not add provider credentials unless you intend to enable external LLM requests.

The app contains no generated-by-AI or ChatGPT watermark. Private access and a private repository reduce casual copying and prevent viewers from editing the source through the app, but no hosted interface can guarantee that its behavior, displayed text, or outputs cannot be copied or reverse-engineered. Restrict GitHub collaborators, keep the Streamlit app private, and do not share repository credentials.

Community Cloud's filesystem may be ephemeral, so SQLite data can disappear after a restart or rebuild. Treat the deployment as a demonstration: it has no login roles, encryption-at-rest management, malware scanning, retention controls, durable backup plan, or multi-tenant isolation. Do not upload sensitive files unless your privacy and hosting requirements permit it.
