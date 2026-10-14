import { CSSProperties, ChangeEvent, FormEvent, useCallback, useEffect, useState } from "react";
import {
  Activity, ArrowDownRight, ArrowRight, ArrowUpRight, Award, BriefcaseBusiness,
  Check, CheckCheck, ChevronDown, ChevronRight, CircleHelp, Clock3, CloudUpload,
  Download, Eye, FileBadge, FileCheck2, FileDown, FileText, Gauge, GraduationCap,
  History, LayoutDashboard, Lightbulb, LoaderCircle, LockKeyhole, Menu, MessageSquareText,
  Pencil, Plus, RefreshCw, RotateCcw, Send, Settings, ShieldCheck, Sparkles, Trash2,
  TriangleAlert, X, XCircle,
} from "lucide-react";

type Suggestion = {
  id: string;
  section: string;
  original_text: string;
  suggested_text: string;
  proposed_text: string;
  reason: string;
  requirement: string;
  expected_impact: string;
  status: "pending" | "approved" | "rejected";
  source: "demo" | "gemini";
};

type ScoreDetails = {
  score: number;
  matched_keywords: string[];
  missing_keywords: string[];
  matched_count: number;
  requirement_count: number;
  method: string;
};

type Workflow = {
  id: string;
  filename: string;
  extension: string;
  created_at: string;
  job_title: string;
  company_name: string;
  demo_mode: boolean;
  limitations: string[];
  analysis: ScoreDetails;
  suggestions: Suggestion[];
  suggestion_warnings: string[];
  instructions: string;
  finalized: boolean;
  final_score: number | null;
  final_score_details: ScoreDetails | null;
  layout_review: string[];
  final_text: string | null;
  original_text: string;
  download_urls: { original: string; docx: string; pdf: string } | null;
};

type HistoryEntry = {
  id: string;
  filename: string;
  job_title: string;
  company_name: string;
  created_at: string;
  score: number;
  finalized: boolean;
};

type Page = "dashboard" | "resumes" | "history" | "settings";
type Preview = "comparison" | "original" | "tailored";

const navigation: { id: Page; label: string; icon: typeof LayoutDashboard }[] = [
  { id: "dashboard", label: "Dashboard", icon: LayoutDashboard },
  { id: "resumes", label: "My Resumes", icon: FileText },
  { id: "history", label: "Tailoring history", icon: History },
  { id: "settings", label: "Settings", icon: Settings },
];

const sectionIcons: Record<string, typeof FileText> = {
  "Professional Summary": FileText,
  Skills: Sparkles,
  Experience: BriefcaseBusiness,
  Internships: GraduationCap,
  Projects: Activity,
  Education: GraduationCap,
  Certifications: FileBadge,
  Achievements: Award,
};

async function api<T>(path: string, options?: RequestInit): Promise<T> {
  const controller = new AbortController();
  const timeout = window.setTimeout(() => controller.abort(), 60000);
  try {
    const response = await fetch(`/api${path}`, { ...options, signal: controller.signal });
    if (!response.ok) {
      let message = `Request failed (${response.status}).`;
      try {
        const body = (await response.json()) as { detail?: string };
        if (body.detail) message = body.detail;
      } catch {
        message = `Request failed (${response.status}). Check that the API server is running.`;
      }
      throw new Error(message);
    }
    if (response.status === 204) return undefined as T;
    return (await response.json()) as T;
  } catch (reason) {
    if (controller.signal.aborted) {
      throw new Error("The request timed out. Check that the API server is responsive and try again.");
    }
    throw reason;
  } finally {
    window.clearTimeout(timeout);
  }
}

function formatDate(value: string) {
  return new Intl.DateTimeFormat("en", { month: "short", day: "numeric", year: "numeric" }).format(new Date(value));
}

function ScoreRing({ score, size = "large" }: { score: number; size?: "large" | "small" }) {
  const style = { "--score": `${score}%` } as CSSProperties;
  return (
    <div className={`score-ring ${size}`} style={style} role="img" aria-label={`Estimated match ${score} out of 100`}>
      <span className="score-ring-center"><strong>{score}</strong><small>/ 100</small></span>
    </div>
  );
}

function App() {
  const [page, setPage] = useState<Page>("dashboard");
  const [history, setHistory] = useState<HistoryEntry[]>([]);
  const [workflow, setWorkflow] = useState<Workflow | null>(null);
  const [activeWorkflowId, setActiveWorkflowId] = useState("");
  const [demoMode, setDemoMode] = useState(true);
  const [loading, setLoading] = useState(false);
  const [busyAction, setBusyAction] = useState("");
  const [error, setError] = useState("");
  const [notice, setNotice] = useState("");
  const [jobDescription, setJobDescription] = useState("");
  const [jobTitle, setJobTitle] = useState("");
  const [companyName, setCompanyName] = useState("");
  const [selectedFile, setSelectedFile] = useState<File | null>(null);
  const [instructions, setInstructions] = useState("");
  const [preview, setPreview] = useState<Preview>("comparison");
  const [showAllMissingKeywords, setShowAllMissingKeywords] = useState(false);
  const [mobileNavOpen, setMobileNavOpen] = useState(false);
  const [initialized, setInitialized] = useState(false);
  const [defaultResumeFilename, setDefaultResumeFilename] = useState<string | null>(null);

  const refreshHistory = useCallback(async () => {
    const [items, health] = await Promise.all([
      api<HistoryEntry[]>("/workflows"),
      api<{ mode: string; default_resume_filename: string | null }>("/health"),
    ]);
    setHistory(items);
    setDemoMode(health.mode === "demo");
    setDefaultResumeFilename(health.default_resume_filename);
  }, []);

  useEffect(() => {
    void refreshHistory()
      .catch((reason: Error) => setError(reason.message))
      .finally(() => setInitialized(true));
  }, [refreshHistory]);

  const loadWorkflow = useCallback(async (id: string) => {
    setError("");
    setLoading(true);
    try {
      const result = await api<Workflow>(`/workflows/${id}`);
      setWorkflow(result);
      setActiveWorkflowId(id);
      setInstructions(result.instructions);
      setPreview(result.finalized ? "comparison" : "original");
      setPage("dashboard");
    } catch (reason) {
      setError((reason as Error).message);
    } finally {
      setLoading(false);
    }
  }, []);

  async function handleAnalyze(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    setError("");
    setNotice("");
    const body = new FormData();
    if (selectedFile) body.set("file", selectedFile);
    body.set("job_description", jobDescription);
    body.set("job_title", jobTitle);
    body.set("company_name", companyName);
    setLoading(true);
    try {
      const result = await api<Workflow>("/analyze", { method: "POST", body });
      setWorkflow(result);
      setActiveWorkflowId(result.id);
      setInstructions("");
      setPreview("original");
      await refreshHistory();
      setNotice("Your original resume is saved separately. Review each suggested change below.");
    } catch (reason) {
      setError((reason as Error).message);
    } finally {
      setLoading(false);
    }
  }

  async function updateSuggestion(id: string, action: SuggestionAction, proposedText?: string) {
    if (!workflow) return;
    setBusyAction(id);
    setError("");
    try {
      const result = await api<Workflow>(`/workflows/${workflow.id}/suggestions/${id}`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ action, proposed_text: proposedText }),
      });
      setWorkflow(result);
      setNotice(action === "approve" ? "Change approved. It will appear in the final resume." : action === "reject" ? "Change rejected. Your original wording will be preserved." : "");
      if (result.finalized === false && workflow.finalized) await refreshHistory();
    } catch (reason) {
      setError((reason as Error).message);
    } finally {
      setBusyAction("");
    }
  }

  async function bulkSuggestions(action: "approve" | "reject") {
    if (!workflow) return;
    setBusyAction("bulk");
    setError("");
    try {
      const result = await api<Workflow>(`/workflows/${workflow.id}/suggestions/bulk`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ action }),
      });
      setWorkflow(result);
      setNotice(action === "approve" ? "All suggestions approved. Generate your tailored resume when ready." : "All suggestions rejected; the original wording will be preserved.");
    } catch (reason) {
      setError((reason as Error).message);
    } finally {
      setBusyAction("");
    }
  }

  async function applyInstructions(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (!workflow || !instructions.trim()) return;
    setError("");
    setBusyAction("instructions");
    try {
      const result = await api<Workflow>(`/workflows/${workflow.id}/instructions`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ instruction: instructions }),
      });
      setWorkflow(result);
      setNotice(result.demo_mode
        ? "Review queue refreshed. Demo mode cannot rewrite resume text; edit the manual option with verified wording, then approve it."
        : "Review queue refreshed. Check the new suggestions; previous decisions were cleared, and nothing was applied automatically.");
    } catch (reason) {
      setError((reason as Error).message);
    } finally {
      setBusyAction("");
    }
  }

  async function generateFinalResume() {
    if (!workflow) return;
    setError("");
    setBusyAction("finalize");
    try {
      const result = await api<Workflow>(`/workflows/${workflow.id}/finalize`, { method: "POST" });
      setWorkflow(result);
      setPreview("comparison");
      setNotice("Your approved resume is ready. The tailored score uses the same JD-coverage criteria as the original.");
      await refreshHistory();
    } catch (reason) {
      setError((reason as Error).message);
    } finally {
      setBusyAction("");
    }
  }

  async function deleteWorkflow(entry: HistoryEntry) {
    const confirmed = window.confirm(`Delete “${entry.filename}” and its tailored files and history? This cannot be undone.`);
    if (!confirmed) return;
    try {
      await api<void>(`/workflows/${entry.id}`, { method: "DELETE" });
      if (activeWorkflowId === entry.id) {
        setWorkflow(null);
        setActiveWorkflowId("");
      }
      await refreshHistory();
      setNotice("Resume, saved version, and tailoring history deleted.");
    } catch (reason) {
      setError((reason as Error).message);
    }
  }

  function choosePage(nextPage: Page) {
    setPage(nextPage);
    setMobileNavOpen(false);
    setError("");
  }

  const pendingCount = workflow?.suggestions.filter((item) => item.status === "pending").length ?? 0;
  const approvedCount = workflow?.suggestions.filter((item) => item.status === "approved").length ?? 0;
  const rejectedCount = workflow?.suggestions.filter((item) => item.status === "rejected").length ?? 0;

  return (
    <div className="app-shell">
      <header className="mobile-header">
        <button className="icon-button mobile-menu" onClick={() => setMobileNavOpen(!mobileNavOpen)} aria-label="Toggle navigation"><Menu size={19} /></button>
        <BrandMark />
        <span className="header-avatar">J</span>
      </header>
      {mobileNavOpen && <button className="mobile-nav-scrim" aria-label="Close menu" onClick={() => setMobileNavOpen(false)} />}
      <aside className={`sidebar ${mobileNavOpen ? "sidebar-open" : ""}`}>
        <BrandMark />
        <div className="workspace-caption">WORKSPACE</div>
        <nav aria-label="Main navigation">
          {navigation.map(({ id, label, icon: Icon }) => (
            <button key={id} className={`nav-item ${page === id ? "nav-item-active" : ""}`} onClick={() => choosePage(id)}>
              <Icon size={17} strokeWidth={1.8} />
              <span>{label}</span>
              {id === "history" && history.length > 0 && <span className="nav-count">{history.length}</span>}
            </button>
          ))}
        </nav>
        <div className="sidebar-bottom">
          <div className="privacy-card">
            <div className="privacy-icon"><ShieldCheck size={16} /></div>
            <div><strong>Your work, protected</strong><span>Files stay on this machine.</span></div>
          </div>
          <div className="user-profile">
            <span className="user-avatar">J</span>
            <div className="user-name"><strong>My workspace</strong><span>Personal account</span></div>
            <ChevronDown size={15} color="#99a7b2" />
          </div>
        </div>
      </aside>

      <main className="main-area">
        <header className="topbar">
          <div className="breadcrumb">Workspace <ChevronRight size={14} /> <span>{navigation.find((item) => item.id === page)?.label}</span></div>
          <div className="topbar-actions">
            <span className={`mode-chip ${demoMode ? "mode-demo" : "mode-ai"}`} title={demoMode ? "No Gemini API key is configured." : "A Gemini API key is configured; each request's status is shown in its review."}><span />{demoMode ? "Demo mode" : "Gemini configured"}</span>
            <button className="top-avatar" aria-label="My workspace">J</button>
          </div>
        </header>

        <div className="page-content">
          {(error || notice) && <div className={`toast ${error ? "toast-error" : "toast-success"}`} role={error ? "alert" : "status"}>
            {error ? <TriangleAlert size={17} /> : <CheckCircleIcon />}
            <span>{error || notice}</span>
            <button onClick={() => { setError(""); setNotice(""); }} aria-label="Dismiss message"><X size={16} /></button>
          </div>}

          {page === "dashboard" && (
            <>
              <section className="welcome-row">
                <div>
                  <div className="eyebrow"><span className="eyebrow-line" /> YOUR NEXT OPPORTUNITY STARTS HERE</div>
                  <h1>Make your experience <span>stand out.</span></h1>
                  <p>Align your resume to the role. Stay in control of every change.</p>
                </div>
                <div className="welcome-illustration" aria-hidden="true"><div className="illustration-orbit orbit-one" /><div className="illustration-orbit orbit-two" /><div className="illustration-document"><span /><span /><span /><span /><i><Check size={13} /></i></div><div className="illustration-spark spark-one">✦</div><div className="illustration-spark spark-two">✧</div></div>
              </section>

              <section className="step-strip" aria-label="Resume tailoring steps">
                {[["01", "Choose or use saved resume"], ["02", "Add the job details"], ["03", "Review & approve"]].map(([number, label], index) => (
                  <div key={number} className="step-item"><span className={`step-number ${index === 0 ? "step-current" : ""}`}>{number}</span><span>{label}</span>{index < 2 && <ArrowRight className="step-arrow" size={15} />}</div>
                ))}
              </section>

              {!workflow && (
                <form className="analysis-form" onSubmit={handleAnalyze}>
                  <div className="form-heading">
                    <div><span className="section-kicker">LET’S GET STARTED</span><h2>Tailor a resume</h2><p>Use your saved master resume or choose another file, then paste the job description.</p></div>
                    <span className="secure-note"><LockKeyhole size={13} /> Stored privately on this device</span>
                  </div>
                  {!demoMode && <div className="ai-notice"><ShieldCheck size={14} />AI mode sends resume text and the job description to Google Gemini for suggested changes. Continue only if you are comfortable sharing that content with the configured AI provider.</div>}
                  <div className="form-grid">
                    <div className="form-column">
                      <label className="field-label" htmlFor="resume-upload">Your original resume <span className="field-optional">Optional PDF or DOCX</span></label>
                      <FilePicker file={selectedFile} onFile={(file) => { setSelectedFile(file); setError(""); }} />
                      <p className="field-optional">{defaultResumeFilename ? `No file selected: using saved master resume (${defaultResumeFilename}).` : "No saved master resume is configured. Choose a PDF or DOCX to continue."}</p>
                      <div className="optional-grid">
                        <label className="input-wrap"><span className="field-label">Job title <span className="field-optional">Optional</span></span><input value={jobTitle} onChange={(event) => setJobTitle(event.target.value)} placeholder="e.g. Product Designer" maxLength={160} /></label>
                        <label className="input-wrap"><span className="field-label">Company <span className="field-optional">Optional</span></span><input value={companyName} onChange={(event) => setCompanyName(event.target.value)} placeholder="e.g. Acme, Inc." maxLength={160} /></label>
                      </div>
                    </div>
                    <div className="form-column">
                      <div className="field-label-row"><label className="field-label" htmlFor="job-description">Job description</label><span className="field-optional">Required</span></div>
                      <textarea id="job-description" className="jd-textarea" value={jobDescription} onChange={(event) => setJobDescription(event.target.value)} placeholder={"Paste the complete job description here...\n\nInclude responsibilities, required skills, and qualifications for a more useful comparison."} minLength={20} maxLength={40000} required />
                      <div className="textarea-footer"><span>{jobDescription.length} characters</span><span>We compare the resume against this role</span></div>
                    </div>
                  </div>
                  <div className="form-footer"><div className="accurate-note"><ShieldCheck size={15} />No skills or experience will be added without your review.</div><button className="button button-primary" type="submit" disabled={loading}>{loading ? <LoaderCircle className="spin" size={16} /> : <Sparkles size={16} />}{loading ? "Analyzing…" : "Analyze resume"}<ArrowRight size={15} /></button></div>
                </form>
              )}

              {workflow && (
                <>
                  <div className="workflow-heading">
                    <div><button className="back-link" onClick={() => { setWorkflow(null); setActiveWorkflowId(""); setError(""); }}><ArrowDownRight size={14} /> Start a new resume</button><h2>Tailoring review</h2><p>{workflow.job_title || "Job match"}{workflow.company_name ? ` at ${workflow.company_name}` : ""} <span className="heading-dot">·</span> {workflow.filename}</p></div>
                    <span className={`mode-chip ${workflow.demo_mode ? "mode-demo" : "mode-ai"}`}><span />{workflow.demo_mode ? "Transparent demo" : "AI suggestions"}</span>
                  </div>
                  {workflow.demo_mode && <div className="demo-banner"><Lightbulb size={17} /><div><strong>{workflow.suggestion_warnings.length ? "No AI edits were applied — manual review options are shown." : "Demo mode is on — no AI service is connected."}</strong><span>{workflow.suggestion_warnings.length ? `${workflow.suggestion_warnings.join(" ")} ` : ""}Your match is a transparent JD-keyword coverage estimate. Suggestions reference your existing text only; edit them yourself if you want to make changes. No information is invented or auto-applied.</span></div></div>}
                  {!workflow.demo_mode && <div className="ai-notice workflow-ai-notice"><ShieldCheck size={14} /><span>Gemini received your resume text and job description to generate these optional suggestions. Check every fact against the original before approving.{workflow.suggestion_warnings.length > 0 && <><br />{workflow.suggestion_warnings.join(" ")}</>}</span></div>}

                  <section className="results-grid">
                    <div className="card original-score-card">
                      <div className="card-label-row"><div><span className="mini-icon mini-blue"><Gauge size={16} /></span><span className="card-eyebrow">ORIGINAL RESUME</span></div><span className="muted-label">JD match estimate</span></div>
                      <div className="score-content"><ScoreRing score={workflow.analysis.score} /><div><strong className="score-heading">{workflow.analysis.score < 40 ? "Room to align" : workflow.analysis.score < 70 ? "Good foundation" : "Strong alignment"}</strong><p>{workflow.analysis.matched_count} of {workflow.analysis.requirement_count} distinct JD terms found in your resume.</p><span className="score-footnote">Not an ATS score or interview prediction</span></div></div>
                    </div>
                    {workflow.finalized && workflow.final_score !== null ? (
                      <div className="card tailored-score-card">
                        <div className="card-label-row"><div><span className="mini-icon mini-green"><ArrowUpRight size={17} /></span><span className="card-eyebrow">TAILORED RESUME</span></div><span className="muted-label">Same JD criteria</span></div>
                        <div className="score-content"><ScoreRing score={workflow.final_score} /><div><strong className="score-heading">{workflow.final_score >= workflow.analysis.score ? "Thoughtfully tailored" : "Verified match"}</strong><p>{workflow.final_score_details?.matched_count} of {workflow.final_score_details?.requirement_count} distinct JD terms found.</p><span className={`improvement-chip ${(workflow.final_score - workflow.analysis.score) >= 0 ? "" : "improvement-down"}`}>{(workflow.final_score - workflow.analysis.score) >= 0 ? <ArrowUpRight size={14} /> : <ArrowDownRight size={14} />}{(workflow.final_score - workflow.analysis.score) >= 0 ? "+" : ""}{workflow.final_score - workflow.analysis.score} pts <span>estimated change</span></span></div></div>
                      </div>
                    ) : (
                      <div className="card score-placeholder">
                        <div className="card-label-row"><div><span className="mini-icon mini-violet"><Sparkles size={15} /></span><span className="card-eyebrow">TAILORED RESUME</span></div><span className="muted-label">After your review</span></div>
                        <div className="placeholder-body"><span className="placeholder-lock"><LockKeyhole size={18} /></span><div><strong>Your next score comes after you approve.</strong><p>We only score the actual resume text after your approved edits are applied.</p></div></div>
                      </div>
                    )}
                  </section>

                  <section className="card insight-card">
                    <div className="insight-heading"><div><span className="mini-icon mini-blue"><Activity size={16} /></span><div><h3>What the job description is asking for</h3><p>{workflow.analysis.method}</p></div></div><span className="estimate-tag"><CircleHelp size={13} /> ESTIMATED MATCH</span></div>
                    <div className="keyword-groups">
                      <div><div className="keyword-title"><span className="keyword-dot match-dot" />Found in your resume <span className="keyword-count">{workflow.analysis.matched_keywords.length}</span></div><div className="keyword-list">{workflow.analysis.matched_keywords.slice(0, 20).map((word) => <span className="keyword keyword-found" key={word}>{word}</span>)}{workflow.analysis.matched_keywords.length === 0 && <span className="empty-copy">No matching terms found yet.</span>}</div></div>
                      <div><div className="keyword-title"><span className="keyword-dot missing-dot" />Not found or not explicit <span className="keyword-count">{workflow.analysis.missing_keywords.length}</span></div><div className="keyword-list">{(showAllMissingKeywords ? workflow.analysis.missing_keywords : workflow.analysis.missing_keywords.slice(0, 20)).map((word) => <span className="keyword keyword-missing" key={word}>{word}</span>)}{workflow.analysis.missing_keywords.length > 20 && <button className="text-action" onClick={() => setShowAllMissingKeywords(!showAllMissingKeywords)}>{showAllMissingKeywords ? "Show fewer terms" : `Show all ${workflow.analysis.missing_keywords.length} terms`}</button>}{workflow.analysis.missing_keywords.length === 0 && <span className="empty-copy">Every analyzed term appears in your resume.</span>}</div></div>
                    </div>
                    <div className="gap-note"><Lightbulb size={15} /><span>A missing keyword may be in your experience but expressed differently. Only add it if it accurately describes your background.</span></div>
                  </section>

                  <section className="review-section">
                    <div className="review-heading"><div><span className="section-kicker">NOTHING CHANGES WITHOUT YOU</span><h2>Review suggested changes <span className="suggestion-count">{workflow.suggestions.length}</span></h2><p>Approve, edit or reject every suggestion before building the final resume.</p></div><div className="bulk-actions"><button className="button button-secondary" onClick={() => void bulkSuggestions("reject")} disabled={busyAction !== ""}><XCircle size={15} />Reject all</button><button className="button button-approve-all" onClick={() => void bulkSuggestions("approve")} disabled={busyAction !== "" || pendingCount === 0}><CheckCheck size={15} />Approve all</button></div></div>
                    <div className="review-stats"><span><i className="status-dot dot-pending" />{pendingCount} to review</span><span><i className="status-dot dot-approved" />{approvedCount} approved</span><span><i className="status-dot dot-rejected" />{rejectedCount} rejected</span><span className="review-scope">Changes apply to your DOCX copy, never the original</span></div>
                    {workflow.suggestions.length === 0 ? <div className="card empty-suggestions"><div className="empty-icon"><MessageSquareText size={21} /></div><h3>No editable resume text was extracted</h3><p>There is no source paragraph to safely suggest or edit. Try a text-based PDF or DOCX with selectable text, or review the original document.</p></div> : (
                      <div className="suggestion-list">{workflow.suggestions.map((suggestion, index) => <SuggestionCard key={`${suggestion.id}-${index}`} suggestion={suggestion} busy={busyAction === suggestion.id} onAction={updateSuggestion} />)}</div>
                    )}
                  </section>

                  <section className="card instructions-card">
                    <div className="instructions-icon"><MessageSquareText size={17} /></div><div className="instructions-body"><div><span className="section-kicker">MAKE IT YOURS</span><h3>Have something specific in mind?</h3><p>Tell us what to focus on. Applying instructions refreshes the suggestion queue and clears previous decisions; nothing changes in your resume automatically.</p></div>
                      <form onSubmit={applyInstructions}><textarea value={instructions} onChange={(event) => setInstructions(event.target.value)} placeholder="e.g. Shorten my professional summary and highlight my verified SQL experience. Keep all facts unchanged." maxLength={2000} aria-label="Additional tailoring instructions" /><div className="instructions-footer"><span>{instructions.length} / 2,000</span><button type="submit" className="button button-primary button-small" disabled={busyAction === "instructions" || !instructions.trim()}>{busyAction === "instructions" ? <LoaderCircle className="spin" size={15} /> : <Send size={14} />}Apply my instructions</button></div></form>
                    </div>
                  </section>

                  {workflow.finalized && workflow.final_text ? (
                    <section className="card final-card">
                      <div className="final-heading"><div className="final-check"><Check size={18} /></div><div><span className="section-kicker">ALL SET</span><h2>Your tailored resume is ready</h2><p>Your unchanged original is safe. Only approved edits were applied.</p></div></div>
                      <div className="preview-tabs" role="tablist" aria-label="Resume preview options">{([["comparison", "Side by side"], ["original", "Original resume"], ["tailored", "Tailored resume"]] as const).map(([id, label]) => <button key={id} role="tab" aria-selected={preview === id} className={preview === id ? "preview-tab-active" : ""} onClick={() => setPreview(id)}>{label}</button>)}</div>
                      {preview === "comparison" ? <div className="comparison-grid"><ResumeTextPanel title="Original resume" text={workflow.original_text} badge="UNCHANGED" /><ResumeTextPanel title="Tailored resume" text={workflow.final_text} badge={`${approvedCount} APPROVED EDIT${approvedCount === 1 ? "" : "S"}`} highlighted /></div> : preview === "original" ? <ResumeTextPanel title="Original resume — unchanged" text={workflow.original_text} badge="ORIGINAL" /> : <ResumeTextPanel title="Tailored resume — approved changes only" text={workflow.final_text} badge={`${approvedCount} APPROVED EDIT${approvedCount === 1 ? "" : "S"}`} highlighted />}
                      <div className="final-score-compare"><div><span>Original</span><strong>{workflow.analysis.score}/100</strong></div><ArrowRight size={16} /><div><span>Tailored</span><strong>{workflow.final_score}/100</strong></div><div className={`final-score-delta ${(workflow.final_score! - workflow.analysis.score) >= 0 ? "" : "improvement-down"}`}>{(workflow.final_score! - workflow.analysis.score) >= 0 ? "+" : ""}{workflow.final_score! - workflow.analysis.score} pts</div></div>
                      <div className="remaining-gaps"><span className="remaining-icon"><Lightbulb size={15} /></span><div><strong>Still not explicit in the tailored text</strong><div className="keyword-list">{(showAllMissingKeywords ? workflow.final_score_details?.missing_keywords ?? [] : (workflow.final_score_details?.missing_keywords ?? []).slice(0, 14)).map((word) => <span className="keyword keyword-missing" key={word}>{word}</span>)}{(workflow.final_score_details?.missing_keywords.length ?? 0) > 14 && <button className="text-action" onClick={() => setShowAllMissingKeywords(!showAllMissingKeywords)}>{showAllMissingKeywords ? "Show fewer terms" : `Show all ${workflow.final_score_details?.missing_keywords.length} terms`}</button>}{!workflow.final_score_details?.missing_keywords.length && <span className="empty-copy">All analyzed JD terms are present. This still is not an ATS guarantee.</span>}</div></div></div>
                      {workflow.layout_review.map((check) => <div className="limitation-note" key={check}><TriangleAlert size={14} />{check}</div>)}
                      {workflow.limitations.map((limitation) => <div className="limitation-note" key={limitation}><CircleHelp size={14} />{limitation}</div>)}
                      <div className="download-row"><a className="button button-download" href={`/api/workflows/${workflow.id}/download/docx`} download><FileDown size={16} />Download DOCX</a><a className="button button-download-secondary" href={`/api/workflows/${workflow.id}/download/pdf`} download><Download size={16} />Download PDF</a><a className="original-link" href={`/api/workflows/${workflow.id}/original`} download><FileText size={15} />Download unchanged original</a><button className="return-edit" onClick={() => { setWorkflow({ ...workflow, finalized: false }); setNotice("Return to your suggestions and make another round of edits."); }}><Pencil size={14} />Return to editing</button></div>
                    </section>
                  ) : (
                    <div className="generate-bar"><div><div className="generate-icon"><FileCheck2 size={17} /></div><div><strong>Ready to build your tailored resume?</strong><span>{pendingCount > 0 ? `Resolve all ${pendingCount} pending suggestion${pendingCount === 1 ? "" : "s"} first.` : `${approvedCount} approved · ${rejectedCount} rejected · your original will remain unchanged.`}</span></div></div><button className="button button-primary" disabled={pendingCount > 0 || busyAction !== ""} onClick={() => void generateFinalResume()}>{busyAction === "finalize" ? <LoaderCircle className="spin" size={16} /> : <Sparkles size={16} />}{busyAction === "finalize" ? "Building…" : "Generate tailored resume"}<ArrowRight size={15} /></button></div>
                  )}
                </>
              )}
              {!workflow && history.length > 0 && <RecentActivity history={history.slice(0, 3)} onSelect={(id) => void loadWorkflow(id)} onHistory={() => choosePage("history")} />}
            </>
          )}

          {page === "resumes" && <CollectionPage title="My resumes" description="Keep your original and tailored versions organized in one place." items={history} onSelect={(id) => void loadWorkflow(id)} onDelete={(entry) => void deleteWorkflow(entry)} />}
          {page === "history" && <CollectionPage title="Tailoring history" description="Pick up where you left off with any job application." items={history} onSelect={(id) => void loadWorkflow(id)} onDelete={(entry) => void deleteWorkflow(entry)} historyView />}
          {page === "settings" && <SettingsPage demoMode={demoMode} count={history.length} onRefresh={() => void refreshHistory()} />}
          {!initialized && <div className="loading-screen"><LoaderCircle className="spin" size={20} />Connecting to your local workspace…</div>}
          {loading && workflow && <div className="working-overlay"><LoaderCircle className="spin" size={19} /><span>Opening your saved review…</span></div>}
        </div>
        <footer className="page-footer"><span>ResumeFit AI <i>·</i> Your story, your say.</span><span><LockKeyhole size={12} />Local workspace</span></footer>
      </main>
    </div>
  );
}

type SuggestionAction = "approve" | "reject" | "edit" | "undo";

function CheckCircleIcon() {
  return <span className="toast-check"><Check size={12} /></span>;
}

function BrandMark() {
  return <div className="brand"><span className="brand-symbol"><FileCheck2 size={20} strokeWidth={2.4} /></span><span>resume<span className="brand-strong">fit</span><span className="brand-ai">AI</span></span></div>;
}

function FilePicker({ file, onFile }: { file: File | null; onFile: (file: File | null) => void }) {
  const [dragging, setDragging] = useState(false);
  function selectFile(event: ChangeEvent<HTMLInputElement>) {
    const next = event.target.files?.[0];
    if (next) onFile(next);
  }
  return (
    <label className={`upload-zone ${dragging ? "upload-dragging" : ""} ${file ? "upload-selected" : ""}`} htmlFor="resume-upload" onDragOver={(event) => { event.preventDefault(); setDragging(true); }} onDragLeave={() => setDragging(false)} onDrop={(event) => { event.preventDefault(); setDragging(false); const next = event.dataTransfer.files[0]; if (next) onFile(next); }}>
      <input id="resume-upload" type="file" accept=".pdf,.docx,application/pdf,application/vnd.openxmlformats-officedocument.wordprocessingml.document" onChange={selectFile} />
      {file ? <><span className="upload-icon selected"><FileCheck2 size={20} /></span><div className="upload-copy"><strong>{file.name}</strong><span>{(file.size / 1024 / 1024).toFixed(2)} MB · <span className="file-ready">Ready to analyze</span></span></div><button className="remove-file" type="button" aria-label="Remove resume" onClick={(event) => { event.preventDefault(); onFile(null); }}><X size={16} /></button></> : <><span className="upload-icon"><CloudUpload size={20} /></span><div className="upload-copy"><strong>Drop your resume here, or <span className="browse-link">browse</span></strong><span>PDF or DOCX · Up to 10 MB</span></div><span className="upload-plus"><Plus size={16} /></span></>}
    </label>
  );
}

function SuggestionCard({ suggestion, busy, onAction }: { suggestion: Suggestion; busy: boolean; onAction: (id: string, action: SuggestionAction, text?: string) => Promise<void> }) {
  const [editing, setEditing] = useState(false);
  const [draft, setDraft] = useState(suggestion.proposed_text);
  const [showDifference, setShowDifference] = useState(false);
  const Icon = sectionIcons[suggestion.section] ?? FileText;
  const stateLabel = suggestion.status === "pending" ? "Awaiting your review" : suggestion.status === "approved" ? "Approved" : "Rejected";
  const hasChange = suggestion.proposed_text.trim() !== suggestion.original_text.trim();

  async function saveEdit() {
    await onAction(suggestion.id, "edit", draft);
    setEditing(false);
  }
  function toggleEdit() {
    setDraft(suggestion.proposed_text);
    setEditing(!editing);
  }

  return (
    <article className={`suggestion-card suggestion-${suggestion.status}`}>
      <div className="suggestion-topline"><div className="suggestion-section"><span className={`suggestion-icon ${suggestion.status === "approved" ? "suggestion-icon-approved" : suggestion.status === "rejected" ? "suggestion-icon-rejected" : ""}`}><Icon size={16} /></span><div><h3>{suggestion.section}</h3><span className={`suggestion-state state-${suggestion.status}`}>{suggestion.status === "approved" ? <Check size={11} /> : suggestion.status === "rejected" ? <X size={11} /> : <Clock3 size={11} />}{stateLabel}</span></div></div><button className={`source-chip ${suggestion.source === "gemini" ? "source-ai" : ""}`} title={suggestion.source === "demo" ? "Rule-based demo suggestion. No AI-generated rewrite." : "Generated using the server-configured Gemini API."}>{suggestion.source === "demo" ? <><Lightbulb size={12} />Evidence-led</> : <><Sparkles size={12} />AI suggestion</>}</button></div>
      <div className="suggestion-body">
        <div className="suggestion-reason"><span className="reason-bullet"><Lightbulb size={13} /></span><div><strong>Why this may help</strong><p>{suggestion.reason}</p></div></div>
        {suggestion.requirement && <div className="requirement-row"><span>JOB REQUIREMENT</span><span className="requirement-value">{suggestion.requirement}</span></div>}
        <div className={`change-grid ${showDifference ? "change-grid-diff" : ""}`}>
          <div className="change-panel change-original"><div className="change-panel-title"><span>Original content</span><span className="change-label-original">AS WRITTEN</span></div><div className="change-content">{suggestion.original_text || <span className="no-source">No matching text in this section. Add details only if they accurately reflect your experience.</span>}</div></div>
          <div className="change-arrow"><ArrowRight size={16} /></div>
          <div className={`change-panel change-proposed ${hasChange ? "change-edited" : ""}`}><div className="change-panel-title"><span>Proposed content</span><span className={`change-label-proposed ${!hasChange ? "label-unchanged" : ""}`}>{hasChange ? "SUGGESTED" : "NO CHANGE"}</span></div>{editing ? <textarea className="edit-textarea" value={draft} onChange={(event) => setDraft(event.target.value)} aria-label={`Edit ${suggestion.section} suggested text`} maxLength={10000} /> : <div className={`change-content ${showDifference && hasChange ? "diff-markup" : ""}`}>{showDifference && hasChange ? <><del>{suggestion.original_text}</del><ins>{suggestion.proposed_text}</ins></> : suggestion.proposed_text || <span className="no-source">No rewrite suggested. Edit this field to make a change yourself.</span>}</div>}</div>
        </div>
        <div className="impact-row"><span><Activity size={13} />Potential impact</span><p>{suggestion.expected_impact}</p></div>
      </div>
      <div className="suggestion-actions">{editing ? <><button className="button button-secondary action-cancel" disabled={busy} onClick={toggleEdit}><X size={14} />Cancel</button><button className="button button-primary action-save" disabled={busy || !draft.trim()} onClick={() => void saveEdit()}><Check size={14} />Save edit</button></> : <><button className="text-action" onClick={() => { setShowDifference(!showDifference); }}><Eye size={14} />{showDifference ? "Hide difference" : "View difference"}</button><div className="action-divider" /><button className="text-action" onClick={toggleEdit}><Pencil size={13} />Edit</button><button className="button button-secondary action-reject" disabled={busy || suggestion.status === "rejected"} onClick={() => void onAction(suggestion.id, "reject")}><X size={14} />Reject</button><button className="button button-approve" disabled={busy || suggestion.status === "approved"} onClick={() => void onAction(suggestion.id, "approve")}>{busy ? <LoaderCircle className="spin" size={14} /> : <Check size={14} />}Approve</button>{suggestion.status !== "pending" && <button className="undo-button" title="Undo decision and restore the original suggestion" onClick={() => void onAction(suggestion.id, "undo")}><RotateCcw size={14} />Undo</button>}</>}</div>
    </article>
  );
}

function ResumeTextPanel({ title, text, badge, highlighted = false }: { title: string; text: string; badge: string; highlighted?: boolean }) {
  return <div className={`resume-text-panel ${highlighted ? "resume-text-highlighted" : ""}`}><div className="resume-panel-head"><strong>{title}</strong><span>{badge}</span></div><pre>{text || "No text available for preview."}</pre></div>;
}

function RecentActivity({ history, onSelect, onHistory }: { history: HistoryEntry[]; onSelect: (id: string) => void; onHistory: () => void }) {
  return <section className="recent-section"><div className="recent-heading"><div><span className="section-kicker">PICK UP WHERE YOU LEFT OFF</span><h2>Recent tailoring</h2></div><button className="text-link" onClick={onHistory}>View history <ArrowRight size={14} /></button></div><div className="recent-grid">{history.map((entry) => <button className="recent-card" key={entry.id} onClick={() => onSelect(entry.id)}><span className="recent-file-icon"><FileText size={17} /></span><span className="recent-card-text"><strong>{entry.job_title || entry.filename}</strong><span>{entry.company_name || entry.filename}</span></span><span className="recent-score">{entry.score}% <span>match</span></span></button>)}</div></section>;
}

function CollectionPage({ title, description, items, onSelect, onDelete, historyView = false }: { title: string; description: string; items: HistoryEntry[]; onSelect: (id: string) => void; onDelete: (entry: HistoryEntry) => void; historyView?: boolean }) {
  const [query, setQuery] = useState("");
  const filtered = items.filter((entry) => `${entry.filename} ${entry.job_title} ${entry.company_name}`.toLowerCase().includes(query.toLowerCase()));
  return <div className="collection-page"><div className="collection-header"><div><span className="section-kicker">{historyView ? "YOUR APPLICATIONS" : "SAVED DOCUMENTS"}</span><h1>{title}</h1><p>{description}</p></div><span className="collection-total">{items.length} {items.length === 1 ? "item" : "items"}</span></div>
    <div className="collection-toolbar"><label className="search-box"><span className="search-label">⌕</span><input aria-label={`Search ${title}`} value={query} onChange={(event) => setQuery(event.target.value)} placeholder="Search by file, job title, or company..." /></label><span className="stored-tag"><LockKeyhole size={13} />Private to this device</span></div>
    {filtered.length === 0 ? <div className="card collection-empty"><span className="empty-icon"><FileText size={21} /></span><h3>{items.length === 0 ? "Your workspace is ready" : "No matching resumes found"}</h3><p>{items.length === 0 ? "Upload a resume on your dashboard to start tailoring for your next role." : "Try a different file name, title, or company."}</p></div> : <div className="card collection-table"><div className="collection-table-header"><span>RESUME / ROLE</span><span>CREATED</span><span>JD MATCH</span><span>STATUS</span><span /></div>{filtered.map((entry) => <div className="collection-row" key={entry.id}><button className="collection-file" onClick={() => onSelect(entry.id)}><span className="collection-file-icon"><FileText size={18} /></span><span><strong>{entry.job_title || entry.filename}</strong><span>{entry.company_name || entry.filename}</span></span></button><span className="collection-date">{formatDate(entry.created_at)}</span><span className="collection-match"><span>{entry.score}</span>/100</span><span className={`collection-status ${entry.finalized ? "status-ready" : ""}`}><i />{entry.finalized ? "Ready" : "In review"}</span><span className="collection-actions"><button title="Open tailoring review" aria-label={`Open ${entry.filename}`} onClick={() => onSelect(entry.id)}><ChevronRight size={17} /></button><button title="Delete resume and history" aria-label={`Delete ${entry.filename}`} onClick={() => onDelete(entry)}><Trash2 size={15} /></button></span></div>)}</div>}
  </div>;
}

function SettingsPage({ demoMode, count, onRefresh }: { demoMode: boolean; count: number; onRefresh: () => void }) {
  return <div className="settings-page"><div className="collection-header"><div><span className="section-kicker">YOUR WORKSPACE</span><h1>Settings</h1><p>Understand how your files, suggestions, and workspace are handled.</p></div></div><div className="settings-grid">  <section className="card settings-card"><div className="settings-title"><span className="settings-icon settings-teal"><Sparkles size={17} /></span><div><h3>Analysis & AI</h3><p>How your match estimate and resume suggestions are generated.</p></div></div><div className="settings-detail"><span>Current analysis mode</span><span className={`mode-chip ${demoMode ? "mode-demo" : "mode-ai"}`}><span />{demoMode ? "Transparent demo mode" : "Gemini configured"}</span></div><p className="settings-explainer">{demoMode ? "No GEMINI_API_KEY is configured for the backend, so matching uses visible keyword coverage and suggestions refer only to content already in your resume. To enable AI-generated suggestions, configure the key as a backend environment variable; it is never sent to this browser." : "A GEMINI_API_KEY is configured. The backend sends resume text and the job description to Google Gemini when you analyze or request suggestions; this label does not verify provider availability. Check each request result, compare every proposal to your source, and approve it only if accurate."}</p><div className="settings-note"><CircleHelp size={15} />Match percentages are estimates of distinct JD vocabulary found in resume text, not actual ATS scores.</div></section>
    <section className="card settings-card"><div className="settings-title"><span className="settings-icon settings-purple"><LockKeyhole size={17} /></span><div><h3>Privacy & storage</h3><p>Your personal files and saved tailoring sessions.</p></div></div><div className="privacy-stats"><div><strong>{count}</strong><span>saved {count === 1 ? "session" : "sessions"}</span></div><span className="privacy-divider" /><div><span className="local-chip"><ShieldCheck size={14} />Local device storage</span></div></div><p className="settings-explainer">By default, original resumes, tailored copies, and history are stored in the local .resume_fit_data folder and a local SQLite database. Set DATABASE_URL to PostgreSQL for deployment. This starter is designed for one local user; add authentication before exposing it to multiple users.</p><button className="settings-refresh" onClick={onRefresh}><RefreshCw size={14} />Refresh connection</button></section>
    <section className="card settings-card settings-docs"><div className="settings-title"><span className="settings-icon settings-blue"><FileBadge size={17} /></span><div><h3>Document compatibility</h3><p>Know what to expect when saving your resume.</p></div></div><div className="compat-row"><span className="compat-icon compat-docx"><FileText size={16} /></span><div><strong>DOCX</strong><span>Edits existing paragraphs to retain styles, layout and alignment. Review pagination after long text replacements.</span></div><span className="compat-tag">BEST PRESERVATION</span></div><div className="compat-row"><span className="compat-icon compat-pdf"><FileDown size={16} /></span><div><strong>PDF</strong><span>Scanned PDFs need OCR. PDF-source exports use text-only layout; DOCX-to-PDF fidelity depends on LibreOffice.</span></div><span className="compat-tag compat-limited">LAYOUT LIMITS</span></div><div className="settings-note"><TriangleAlert size={15} />Always open and visually check your downloaded files before sending an application.</div></section></div></div>;
}

export default App;
