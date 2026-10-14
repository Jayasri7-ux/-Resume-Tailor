# ResumeFit AI

A local-first React, TypeScript, Tailwind CSS, and FastAPI app for comparing a resume with a job description, reviewing every proposed change, and exporting an approved copy. It never modifies the original upload.

## Run locally on Windows

Open two PowerShell terminals in the project folder.

**API**

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r backend\requirements.txt
python -m uvicorn backend.main:app --host 127.0.0.1 --port 8000
```

**Web app**

Install Node.js 20 or newer, then run:

```powershell
npm install
npm run dev
```

Open the local URL printed by Vite (usually `http://127.0.0.1:5173`). The Vite development server proxies `/api` requests to FastAPI. The interactive API reference is available at `http://127.0.0.1:8000/docs`.

To run the production frontend, use `npm run build`, then start FastAPI; it will serve the generated `dist` directory when present.

The optional Google Docs helper in `gdocs_tailor.py` uses Google OAuth and Gemini libraries. Install `backend\requirements-dev.txt` to work on or run that helper.

## Analysis and AI mode

Without a configured provider, the app explicitly displays **Demo mode**. The match percentage is a transparent estimate of distinct, non-stopword terms shared with the job description, not an ATS score. Demo suggestions cite only source-resume text; no model-generated analysis is represented as AI. You can edit suggestions yourself and approve or reject each one.

For Gemini-generated suggestions, set `GEMINI_API_KEY` in the **backend process environment only**, then restart the API. Optionally set `GEMINI_MODEL`. The key is never sent to the browser or written to the workflow database. Provider errors are shown rather than silently substituting demo results. Independently check every AI proposal against your real experience before approving it.

When Gemini returns no evidence-supported edits, the review shows a single manual-edit option and clearly labels the workflow as demo/manual mode; unchanged AI proposals are omitted. Applying new instructions refreshes the suggestion queue and clears prior decisions so stale cards are not left behind. When a resume and job description have no matching terms, the review queue still offers existing resume text for manual editing. The unchanged fallback is not an optimization: the tailored score changes only after you edit a proposal, approve it, and generate the final resume. Add missing job-description terms only when they accurately describe your experience; unsupported keywords are not inserted automatically. “Gemini configured” means an API key is present, not that Google’s service has been reached successfully.

## Files, database, and exports

- Originals, approved copies, and the default SQLite history are stored under `.resume_fit_data`; that directory, local OAuth files, environment files, and dependency/build outputs are ignored by Git.
- To use a saved resume without choosing a file for each analysis, put the master DOCX at `.resume_fit_data/master_resume.docx` and restart the API. The backend loads its text into application state and the `workspace_resumes` database record, then copies the original DOCX into each workflow. An uploaded PDF or DOCX still overrides the saved master for that analysis.
- DOCX updates edit existing text nodes in place, retain character-level run styling, hyperlink detected web/email addresses, reject replacement text longer than its source paragraph, and refuse exports whose generated PDF exceeds one page. Without LibreOffice, the fallback PDF is reflowed and cannot prove the DOCX's rendered page count.
- Set `DATABASE_URL` to a SQLAlchemy PostgreSQL URL such as `postgresql+psycopg://user:password@localhost/resumefit` to use PostgreSQL for workflow metadata. Resume files remain local; production deployments need private, access-controlled file storage.
- DOCX exports edit existing paragraphs to preserve Word styles and most layout details. Review line wrapping and page breaks after any text changes. A DOCX-source PDF is exported through LibreOffice when installed; otherwise a clearly described, reflowed PDF is produced.
- Text-based PDFs can be analyzed. Scanned PDFs are rejected with a message because OCR is not configured. Editing or converting a PDF cannot preserve its exact formatting; the generated files use a basic text layout.
- Open both downloads and visually check pagination, alignment, blank pages, and hyperlinks before submitting. The in-app comparison is a text preview, not a page-rendering or overflow inspection tool.
- Deleting a resume in the app also deletes its saved outputs and tailoring history.

This starter runs as a **single-user local app**, bound to localhost. It does not implement user authentication. Do not expose it to a network or use it for multiple users without adding authentication, authorization, private file storage, HTTPS, and deployment-appropriate data-retention controls.

## Tests

```powershell
python -m pip install -r backend\requirements-dev.txt
python -m pytest
```
