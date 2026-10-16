# ResumeFit AI

A local-first React, TypeScript, Tailwind CSS, and FastAPI app for comparing a resume with a job description, reviewing every proposed change, and exporting an approved copy. It never modifies the original upload.

## Run locally on Windows

Open two PowerShell terminals in the project folder.

**API**

```powershell
# Python 3.10 to 3.12 is recommended (e.g. py -3.12 -m venv .venv-py312)
# Note: Python 3.15 lacks pre-built Windows wheels for pyyaml and requires MSVC C++ tools.
.\.venv-py312\Scripts\python.exe -m uvicorn backend.main:app --host 127.0.0.1 --port 8000
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

The backend uses the local Ollama API at `http://localhost:11434/api/generate` by default, with the `llama3.2` model. Install and start Ollama, then pull the configured model (for example, `ollama pull llama3.2`) before analyzing a resume. Optionally set `OLLAMA_API_URL`, `OLLAMA_MODEL`, or `OLLAMA_MAX_RESPONSE_TOKENS` in the backend environment. Each request returns up to four concise suggestions for the most relevant resume sections and is capped at 768 output tokens by default, with a 900-second server deadline. The resume text and job description are sent to that configured Ollama endpoint; this workflow does not call Google Gemini. Provider errors are shown rather than silently substituting demo results.

If `OLLAMA_API_URL` is explicitly empty, the app displays **Demo mode**. The match percentage is a transparent estimate of distinct job-description terms found in the resume, using exact terms and a small disclosed list of related-skill equivalents. It is not an ATS score; it does not guarantee a target score or hiring outcome. Keyword matching preserves short skills and compound technology terms such as `AI`, `C++`, `CI/CD`, and `.NET`. Demo suggestions cite only source-resume text; no model-generated analysis is represented as AI. You can edit suggestions yourself and approve or reject each one.

When Ollama returns no evidence-supported edits, the review shows a manual-edit option and clearly labels the workflow as demo/manual mode; unchanged AI proposals are omitted. Applying new instructions refreshes the suggestion queue and clears prior decisions so stale cards are not left behind. When a resume and job description have no matching terms, the review queue still offers existing resume text for manual editing. The unchanged fallback is not an optimization: the tailored score changes only after you edit a proposal, approve it, and generate the final resume. Add missing job-description terms only when they accurately describe your experience; unsupported keywords are not inserted automatically.

## Files, database, and exports

- Originals, approved copies, and the default SQLite history are stored under `.resume_fit_data`; that directory, local OAuth files, environment files, and dependency/build outputs are ignored by Git.
- To use a saved resume without choosing a file for each analysis, put the master DOCX at `.resume_fit_data/master_resume.docx` and restart the API. The backend loads its text into application state and the `workspace_resumes` database record, then copies the original DOCX into each workflow. An uploaded PDF or DOCX still overrides the saved master for that analysis.
- For DOCX uploads, both download routes use the stored original as the DOCX template, edit existing text nodes in place, retain character-level run styling, and reject replacement text longer than its source paragraph. Generated PDFs are limited to one page. When LibreOffice is unavailable, the ReportLab fallback is reflowed and cannot prove the DOCX's Word pagination.
- Set `DATABASE_URL` to a SQLAlchemy PostgreSQL URL such as `postgresql+psycopg://user:password@localhost/resumefit` to use PostgreSQL for workflow metadata. Resume files remain local; production deployments need private, access-controlled file storage.
- DOCX exports edit existing paragraphs to preserve Word styles and most layout details. Review line wrapping and page breaks after any text changes. A DOCX-source PDF is exported through LibreOffice when installed; otherwise a clearly described, reflowed PDF is produced.
- Text-based PDFs can be analyzed. Scanned PDFs are rejected with a message because OCR is not configured. Editing or converting a PDF cannot preserve its exact formatting; the generated files use a basic text layout.
- Open both downloads and visually check pagination, alignment, blank pages, and hyperlinks before submitting. The in-app comparison is a text preview, not a page-rendering or overflow inspection tool.
- Deleting a resume in the app also deletes its saved outputs and tailoring history.

This starter runs as a **single-user local app**, bound to localhost. It does not implement user authentication. Do not expose it to a network or use it for multiple users without adding authentication, authorization, private file storage, HTTPS, and deployment-appropriate data-retention controls.

## Tests

```powershell
# Run backend pytest suite
.\.venv-py312\Scripts\python.exe -m pytest

# Run frontend typecheck and build
npm run build
```
