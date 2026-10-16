import { CSSProperties, ChangeEvent, FormEvent, useCallback, useEffect, useRef, useState } from "react";
import { renderAsync } from "docx-preview";
import {
  Activity, ArrowDownRight, ArrowRight, ArrowUpRight, Award, BriefcaseBusiness,
  Check, CheckCheck, ChevronDown, ChevronRight, CircleHelp, Clock3, CloudUpload,
  Download, Eye, FileBadge, FileCheck2, FileDown, FileText, Gauge, GraduationCap,
  History, LayoutDashboard, Lightbulb, ListChecks, LoaderCircle, LockKeyhole, Menu, MessageSquareText,
  Pencil, Plus, RefreshCw, RotateCcw, Send, Settings, ShieldCheck, Sparkles, Trash2,
  TriangleAlert, UserCheck, X, XCircle,
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
  source: "demo" | "manual" | "ollama" | "gemini";
  paragraph_index?: number;
};

type ExportSuggestionState = Pick<Suggestion, "id" | "status" | "proposed_text">;

type CategoryScore = {
  category: string;
  weight: number;
  score: number;
  weighted_score: number;
  matched_count: number;
  total_count: number;
  matched_keywords: string[];
  missing_keywords: string[];
};

type EvidenceItem = {
  requirement: string;
  category: string;
  classification: "fully evidenced" | "partially evidenced" | "not evidenced" | "uncertain";
  snippets: string[];
  confidence: number;
};

type DiscoveredTool = {
  tool: string;
  domain: string;
  explanation: string;
  status: "matched" | "approved" | "related" | "skipped" | "pending";
};

type Tier1RoleMatch = {
  domain: string;
  matched_role: string;
  eligibility_status: string;
  eligible_clusters: string[];
  warning?: string;
};

type ScoreDetails = {
  score: number;
  matched_keywords: string[];
  missing_keywords: string[];
  matched_count: number;
  requirement_count: number;
  exact_matched_count: number;
  semantic_matched_count: number;
  semantic_matches: Record<string, string[]>;
  method: string;
  category_scores?: CategoryScore[];
  evidence_breakdown?: EvidenceItem[];
  career_stage?: string;
  tier1_role_match?: Tier1RoleMatch;
  discovered_tools?: DiscoveredTool[];
  missing_technologies?: DiscoveredTool[];
  core_domain?: string;
  core_domain_info?: {
    domain: string;
    detected: boolean;
    confidence: number;
    description: string;
    keywords_matched: string[];
  };
  confirmed_tools?: string[];
  skipped_tools?: string[];
};

type CandidateProfile = {
  id?: string;
  full_name: string;
  headline: string;
  contact?: string;
  email?: string;
  phone?: string;
  location?: string;
  career_stage?: string;
  eligible_job_titles?: string[];
  summary: string;
  skills: string[];
  categorized_skills?: Record<string, string[]>;
  experience: any[];
  education: any[];
  projects: any[];
  certifications: string[];
  updated_at?: string;
};

type JobStatus = "queued" | "processing" | "completed" | "failed";
type AnalysisJobResponse = {
  id: string;
  job_id: string;
  workflow_id?: string;
  status: JobStatus;
  progress: number;
  message?: string;
  error?: string | null;
  workflow?: Workflow;
};

type Workflow = {
  id: string;
  filename: string;
  extension: string;
  created_at: string;
  job_title: string;
  company_name: string;
  demo_mode: boolean;
  ai_provider?: "ollama" | "gemini" | null;
  limitations: string[];
  analysis: ScoreDetails;
  suggestions: Suggestion[];
  suggestion_warnings: string[];
  instructions: string;
  finalized: boolean;
  final_score: number | null;
  final_score_details: ScoreDetails | null;
  tailored_text?: string | null;
  tailored_score?: number | null;
  tailored_score_details?: ScoreDetails | null;
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

type Page = "dashboard" | "profile" | "resumes" | "history" | "settings";
type Preview = "comparison" | "original" | "tailored";

const navigation: { id: Page; label: string; icon: typeof LayoutDashboard }[] = [
  { id: "dashboard", label: "Dashboard", icon: LayoutDashboard },
  { id: "profile", label: "Candidate profile", icon: UserCheck },
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

async function api<T>(path: string, options?: RequestInit, timeoutMs = 300000): Promise<T> {
  const controller = new AbortController();
  const timeout = window.setTimeout(() => controller.abort(), timeoutMs);
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
      throw new Error("The request timed out. Check that the API server or local Ollama model is responsive and try again.");
    }
    throw reason;
  } finally {
    window.clearTimeout(timeout);
  }
}

async function downloadApiFile(path: string, body: unknown): Promise<{ blob: Blob; filename: string }> {
  const controller = new AbortController();
  const timeout = window.setTimeout(() => controller.abort(), 300000);
  try {
    const response = await fetch(`/api${path}`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
      signal: controller.signal,
    });
    if (!response.ok) {
      let message = `Download failed (${response.status}).`;
      try {
        const result = (await response.json()) as { detail?: string };
        if (result.detail) message = result.detail;
      } catch {
        message = `Download failed (${response.status}). Check that the API server is running.`;
      }
      throw new Error(message);
    }

    const disposition = response.headers.get("Content-Disposition") ?? "";
    const encodedFilename = disposition.match(/filename\*=UTF-8''([^;]+)/i)?.[1];
    const plainFilename = disposition.match(/filename="?([^";]+)"?/i)?.[1];
    const filename = encodedFilename
      ? decodeURIComponent(encodedFilename)
      : plainFilename ?? "tailored-resume";
    return { blob: await response.blob(), filename };
  } catch (reason) {
    if (controller.signal.aborted) {
      throw new Error("The download timed out. Check that the API server is responsive and try again.");
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
  const suggestionActionLock = useRef(false);
  const [activeWorkflowId, setActiveWorkflowId] = useState("");
  const [demoMode, setDemoMode] = useState(true);
  const [loading, setLoading] = useState(false);
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [analysisProgress, setAnalysisProgress] = useState(0);
  const [analysisMessage, setAnalysisMessage] = useState("");
  const [busyAction, setBusyAction] = useState("");
  const [aiRequestTimeoutMs, setAiRequestTimeoutMs] = useState(915000);
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
  const [candidateProfile, setCandidateProfile] = useState<CandidateProfile | null>(null);
  const [useProfileForTailoring, setUseProfileForTailoring] = useState(false);
  const [isPermissionModalOpen, setIsPermissionModalOpen] = useState(false);

  const refreshHistory = useCallback(async () => {
    const [items, health, profileRes] = await Promise.all([
      api<HistoryEntry[]>("/workflows"),
      api<{
        mode: string;
        default_resume_filename: string | null;
        ollama_timeout_seconds: number;
      }>("/health"),
      api<CandidateProfile>("/profile").catch(() => null),
    ]);
    setHistory(items);
    setDemoMode(health.mode === "demo");
    setAiRequestTimeoutMs((health.ollama_timeout_seconds + 30) * 1000);
    setDefaultResumeFilename(health.default_resume_filename);
    if (profileRes) setCandidateProfile(profileRes);
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
    if (isAnalyzing) return;
    setError("");
    setNotice("");
    setIsAnalyzing(true);
    setLoading(true);
    setAnalysisProgress(15);
    setAnalysisMessage("Starting background analysis job...");

    const body = new FormData();
    if (useProfileForTailoring) {
      body.set("use_profile", "true");
    } else if (selectedFile) {
      body.set("file", selectedFile);
    }
    body.set("job_description", jobDescription);
    body.set("job_title", jobTitle);
    body.set("company_name", companyName);
    body.set("background", "true");

    try {
      const jobResp = await api<AnalysisJobResponse>("/analyze", { method: "POST", body }, 30000);
      const jobId = jobResp.job_id || jobResp.id;

      if (jobResp.status === "completed" && ((jobResp as unknown as Workflow).analysis || jobResp.workflow)) {
        const wf = jobResp.workflow || (jobResp as unknown as Workflow);
        setWorkflow(wf);
        setActiveWorkflowId(wf.id);
        setInstructions("");
        setPreview("original");
        await refreshHistory();
        setNotice("Analysis complete! Review suggestions and weighted match breakdown below.");
        return;
      }

      setAnalysisProgress(jobResp.progress || 25);
      setAnalysisMessage(jobResp.message || "Analyzing resume against job requirements...");

      const startTime = Date.now();
      const maxWaitMs = Math.max(aiRequestTimeoutMs, 360000);
      let finished = false;

      while (Date.now() - startTime < maxWaitMs) {
        await new Promise((resolve) => setTimeout(resolve, 1200));
        try {
          const statusResp = await api<AnalysisJobResponse>(`/jobs/${jobId}`, undefined, 10000);
          if (statusResp.progress) {
            setAnalysisProgress(statusResp.progress);
          }
          if (statusResp.message) {
            setAnalysisMessage(statusResp.message);
          }

          if (statusResp.status === "completed") {
            let wf = statusResp.workflow;
            if (!wf) {
              wf = await api<Workflow>(`/workflows/${statusResp.workflow_id || jobId}`);
            }
            setWorkflow(wf);
            setActiveWorkflowId(wf.id);
            setInstructions(wf.instructions || "");
            setPreview("original");
            const pendingMissingCount =
              wf.analysis?.missing_technologies?.filter((t: DiscoveredTool) => t.status === "pending").length ||
              wf.analysis?.discovered_tools?.filter((t: DiscoveredTool) => t.status === "pending").length || 0;
            if (pendingMissingCount > 0) {
              setIsPermissionModalOpen(true);
            }
            await refreshHistory();
            setNotice("Analysis complete! Review each suggested change and weighted requirement evidence.");
            finished = true;
            break;
          } else if (statusResp.status === "failed") {
            throw new Error(statusResp.error || "Analysis job encountered an error.");
          }
        } catch (pollErr) {
          if (pollErr instanceof Error && pollErr.message.includes("Analysis job encountered an error")) {
            throw pollErr;
          }
        }
      }

      if (!finished) {
        throw new Error("Analysis job timed out while waiting for local AI response. Check Ollama status and try again.");
      }
    } catch (reason) {
      setError((reason as Error).message);
    } finally {
      setLoading(false);
      setIsAnalyzing(false);
      setAnalysisProgress(0);
      setAnalysisMessage("");
    }
  }

  async function updateSuggestion(id: string, action: SuggestionAction, proposedText?: string): Promise<boolean> {
    if (!workflow || suggestionActionLock.current) return false;
    suggestionActionLock.current = true;
    setBusyAction(id);
    setError("");
    try {
      const result = await api<Workflow>(`/workflows/${workflow.id}/suggestions/${id}`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ action, proposed_text: proposedText }),
      });
      setWorkflow(result);
      if (action === "approve") {
        setPreview("comparison");
        setNotice("Change approved. It now appears in the tailored resume preview.");
      } else if (action === "edit") {
        setNotice("Edit saved. Approve it when you are ready to apply it to the tailored resume.");
      } else {
        setNotice(action === "reject" ? "Change rejected. Your original wording will be preserved." : "");
      }
      if (result.finalized === false && workflow.finalized) await refreshHistory();
      return true;
    } catch (reason) {
      setError((reason as Error).message);
      return false;
    } finally {
      suggestionActionLock.current = false;
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
      if (action === "approve") setPreview("comparison");
      setNotice(action === "approve" ? "All suggestions approved. Generate your tailored resume when ready." : "All suggestions rejected; the original wording will be preserved.");
    } catch (reason) {
      setError((reason as Error).message);
    } finally {
      setBusyAction("");
    }
  }

  async function tailorSection(suggestion: Suggestion) {
    if (!workflow || busyAction !== "") return;
    setBusyAction(suggestion.id);
    setError("");
    setNotice(`Generating targeted AI rewrite for ${suggestion.section}...`);
    try {
      const result = await api<Workflow>(`/workflows/${workflow.id}/tailor-section`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          section: suggestion.section,
          paragraph_index: suggestion.paragraph_index,
        }),
      }, 60000);
      setWorkflow(result);
      setNotice(`Targeted AI rewrite for ${suggestion.section} completed and updated!`);
    } catch (reason) {
      setError((reason as Error).message);
    } finally {
      setBusyAction("");
    }
  }

  async function applyInstructions(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    const instruction = instructions.trim();
    if (!workflow || busyAction !== "" || !instruction) return;
    setError("");
    setNotice("");
    setBusyAction("instructions");
    try {
      const result = await api<Workflow>(`/workflows/${workflow.id}/instructions`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ instruction }),
      }, aiRequestTimeoutMs);
      setWorkflow(result);
      setNotice(result.demo_mode
        ? "Review queue refreshed. Demo mode cannot rewrite resume text; edit the manual option with verified wording, then approve it."
        : result.ai_provider === null
          ? "Ollama returned no safe rewrite. Your original is unchanged; edit the manual review card to continue."
          : "Review queue refreshed. Check the new suggestions; previous decisions were cleared, and nothing was applied automatically.");
    } catch (reason) {
      setError((reason as Error).message);
    } finally {
      setBusyAction("");
    }
  }

  async function downloadFinalResume(format: "docx" | "pdf") {
    if (!workflow || busyAction !== "") return;
    setError("");
    setNotice("");
    setBusyAction(`download-${format}`);
    try {
      const { blob, filename } = await downloadApiFile(
        `/workflows/${workflow.id}/download/${format}`,
        {
          suggestions: workflow.suggestions.map(({ id, status, proposed_text }): ExportSuggestionState => ({
            id,
            status,
            proposed_text,
          })),
        },
      );
      if (blob.size === 0) throw new Error("The API returned an empty resume file. Please try again.");
      const objectUrl = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = objectUrl;
      link.download = filename;
      document.body.appendChild(link);
      link.click();
      link.remove();
      window.setTimeout(() => URL.revokeObjectURL(objectUrl), 1000);

      setPreview("comparison");
      setNotice(`Your tailored ${format.toUpperCase()} download is ready.`);
    } catch (reason) {
      setError((reason as Error).message);
    } finally {
      setBusyAction("");
    }
  }

  async function generateFinalResume() {
    if (!workflow || busyAction !== "" || pendingCount > 0) return;
    setError("");
    setNotice("");
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
  const hasApprovedSuggestions = approvedCount > 0 || (workflow?.finalized ?? false);
  const liveScore = workflow ? (workflow.tailored_score ?? workflow.final_score ?? null) : null;
  const liveScoreDetails = workflow ? (workflow.tailored_score_details ?? workflow.final_score_details ?? null) : null;
  const scoreVariance = liveScore !== null && workflow ? liveScore - workflow.analysis.score : 0;
  const tailoredDisplay = workflow?.tailored_text ?? workflow?.final_text ?? "";

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
            <span className={`mode-chip ${demoMode ? "mode-demo" : "mode-ai"}`} title={demoMode ? "The local Ollama endpoint is not configured." : "Suggestions are generated by the configured local Ollama endpoint."}><span />{demoMode ? "Demo mode" : "Ollama configured"}</span>
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

              {!workflow && (selectedFile || defaultResumeFilename) && (
                <div className="single-preview-bounded">
                  <DocumentPreviewPanel
                    key={selectedFile ? `${selectedFile.name}-${selectedFile.size}-${selectedFile.lastModified}` : "saved-default-resume"}
                    title={selectedFile ? "Selected original resume" : "Saved original resume"}
                    badge={selectedFile ? "SELECTED FILE" : "SAVED DEFAULT"}
                    filename={selectedFile?.name ?? defaultResumeFilename ?? "resume.docx"}
                    sourceUrl={selectedFile ? "" : "/api/resume/default"}
                    sourceFile={selectedFile ?? undefined}
                  />
                </div>
              )}

              {!workflow && (
                <form className="analysis-form" onSubmit={handleAnalyze}>
                  <div className="form-heading">
                    <div><span className="section-kicker">LET’S GET STARTED</span><h2>Tailor a resume</h2><p>Use your saved master resume or choose another file, then paste the job description.</p></div>
                    <span className="secure-note"><LockKeyhole size={13} /> Stored privately on this device</span>
                  </div>
                  {!demoMode && <div className="ai-notice"><ShieldCheck size={14} />AI mode sends resume text and the job description to your local Ollama server for suggested changes. The configured endpoint is local; review your Ollama setup before use.</div>}
                  <div className="form-grid">
                    <div className="form-column">
                      <label className="field-label" htmlFor="resume-upload">Your original resume <span className="field-optional">Optional PDF or DOCX</span></label>
                      {candidateProfile?.full_name && (
                        <label className="use-profile-toggle">
                          <input
                            type="checkbox"
                            checked={useProfileForTailoring}
                            onChange={(event) => {
                              setUseProfileForTailoring(event.target.checked);
                              if (event.target.checked) setSelectedFile(null);
                            }}
                          />
                          <div>
                            <strong>Use saved candidate profile ({candidateProfile.full_name})</strong>
                            <span style={{ display: "block", color: "#6f827d", fontSize: "9px" }}>
                              Tailor directly from your saved profile without uploading a file.
                            </span>
                          </div>
                        </label>
                      )}
                      {useProfileForTailoring ? (
                        <div className="card" style={{ padding: "12px 14px", marginBottom: "15px", background: "#f8fbf9", borderColor: "#cde3dc" }}>
                          <div style={{ display: "flex", alignItems: "center", gap: "8px", color: "#258273", fontSize: "10px", fontWeight: "600" }}>
                            <UserCheck size={16} /> Using Candidate Profile: {candidateProfile?.full_name}
                          </div>
                          <p style={{ margin: "4px 0 0", fontSize: "9px", color: "#6d817b" }}>
                            {candidateProfile?.headline || "Verified candidate background details will be tailored to the job description."}
                          </p>
                        </div>
                      ) : (
                        <>
                          <FilePicker file={selectedFile} onFile={(file) => { setSelectedFile(file); setError(""); }} />
                          <p className="field-optional">{defaultResumeFilename ? `No file selected: using saved master resume (${defaultResumeFilename}).` : "No saved master resume is configured. Choose a PDF or DOCX to continue."}</p>
                        </>
                      )}
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
                  {isAnalyzing && <section className="analysis-progress" role="status" aria-live="polite" aria-label="Resume analysis progress">
                    <div className="analysis-progress-heading"><LoaderCircle className="spin" size={17} /><div><strong>{analysisMessage || "Analyzing your resume and job description"}</strong><span>Background analysis job is running locally with real-time progress.</span></div></div>
                    <div className="analysis-progress-bar">
                      <div className="analysis-progress-fill" style={{ width: `${Math.max(analysisProgress, 10)}%` }} />
                    </div>
                    <div className="analysis-progress-meta">
                      <span>{analysisProgress > 0 ? `${analysisProgress}% complete` : "Processing..."}</span>
                      <span>Keep this page open while your local AI analyzes requirements</span>
                    </div>
                  </section>}
                </form>
              )}

              {workflow && (
                <>
                  <div className="workflow-heading">
                    <div><button className="back-link" onClick={() => { setWorkflow(null); setActiveWorkflowId(""); setError(""); }}><ArrowDownRight size={14} /> Start a new resume</button><h2>Tailoring review</h2><p>{workflow.job_title || "Job match"}{workflow.company_name ? ` at ${workflow.company_name}` : ""} <span className="heading-dot">·</span> {workflow.filename}</p></div>
                    <span className={`mode-chip ${workflow.demo_mode || !workflow.ai_provider ? "mode-demo" : "mode-ai"}`}><span />{workflow.demo_mode ? "Transparent demo" : workflow.ai_provider ? "AI suggestions" : "Manual editing"}</span>
                  </div>
                  {workflow.demo_mode && <div className="demo-banner"><Lightbulb size={17} /><div><strong>{workflow.suggestion_warnings.length ? "No AI edits were applied — manual review options are shown." : "Demo mode is on — no AI service is connected."}</strong><span>{workflow.suggestion_warnings.length ? `${workflow.suggestion_warnings.join(" ")} ` : ""}Your match is a transparent JD-keyword coverage estimate. Suggestions reference your existing text only; edit them yourself if you want to make changes. No information is invented or auto-applied.</span></div></div>}
                  {!workflow.demo_mode && <div className="ai-notice workflow-ai-notice"><ShieldCheck size={14} /><span>{workflow.ai_provider === null ? "Ollama did not produce a safe rewrite for this request. Your original remains unchanged; edit the manual review card if you want to make a change." : workflow.ai_provider === "gemini" ? "These suggestions were generated by Google Gemini before switching to local Ollama." : "Your local Ollama server received the resume text and job description to generate these optional suggestions."} Check every fact against the original before approving.{workflow.suggestion_warnings.filter((warning) => !warning.startsWith("Ollama did not return a valid")).length > 0 && <><br />{workflow.suggestion_warnings.filter((warning) => !warning.startsWith("Ollama did not return a valid")).join(" ")}</>}</span></div>}

                  <RoleEligibilityBanner analysis={workflow.analysis} />

                  <section className="results-grid">
                    <div className="card original-score-card">
                      <div className="card-label-row"><div><span className="mini-icon mini-blue"><Gauge size={16} /></span><span className="card-eyebrow">ORIGINAL RESUME</span></div><span className="muted-label">JD match estimate</span></div>
                      <div className="score-content"><ScoreRing score={workflow.analysis.score} /><div><strong className="score-heading">{workflow.analysis.score < 40 ? "Room to align" : workflow.analysis.score < 70 ? "Good foundation" : "Strong alignment"}</strong><p>{workflow.analysis.matched_count} of {workflow.analysis.requirement_count} distinct JD terms found in your resume.</p><span className="score-footnote">Not an ATS score or interview prediction</span></div></div>
                    </div>
                    {hasApprovedSuggestions && liveScore !== null ? (
                      <div className="card tailored-score-card">
                        <div className="card-label-row"><div><span className="mini-icon mini-green"><ArrowUpRight size={17} /></span><span className="card-eyebrow">TAILORED RESUME</span></div><span className="muted-label">Live match estimate</span></div>
                        <div className="score-content"><ScoreRing score={liveScore} /><div><strong className="score-heading">{liveScore >= workflow.analysis.score ? "Thoughtfully tailored" : "Verified match"}</strong><p>{liveScoreDetails?.matched_count ?? workflow.analysis.matched_count} of {liveScoreDetails?.requirement_count ?? workflow.analysis.requirement_count} distinct JD terms found.</p><span className={`improvement-chip ${scoreVariance >= 0 ? "" : "improvement-down"}`}>{scoreVariance >= 0 ? <ArrowUpRight size={14} /> : <ArrowDownRight size={14} />}{scoreVariance >= 0 ? "+" : ""}{scoreVariance} pts <span>estimated change</span></span></div></div>
                      </div>
                    ) : (
                      <div className="card score-placeholder">
                        <div className="card-label-row"><div><span className="mini-icon mini-violet"><Sparkles size={15} /></span><span className="card-eyebrow">TAILORED RESUME</span></div><span className="muted-label">Live match estimate</span></div>
                        <div className="placeholder-body"><span className="placeholder-lock"><LockKeyhole size={18} /></span><div><strong>Your next score updates live as you approve.</strong><p>Approve or edit suggestions below to incorporate missing JD keywords and boost your score.</p></div></div>
                      </div>
                    )}
                  </section>

                  {(() => {
                    const pendingMissing = (workflow.analysis.missing_technologies || workflow.analysis.discovered_tools || [])
                      .filter((t) => t.status === "pending");
                    if (pendingMissing.length === 0) return null;
                    return (
                      <div style={{ marginBottom: "16px", padding: "14px 18px", background: "#fef8ee", border: "1px solid #f6e0b5", borderRadius: "10px", display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: "12px" }}>
                        <div style={{ display: "flex", alignItems: "center", gap: "10px" }}>
                          <span style={{ padding: "6px", background: "#faeccb", borderRadius: "50%", color: "#996a13", display: "inline-flex" }}>
                            <TriangleAlert size={18} />
                          </span>
                          <div>
                            <strong style={{ fontSize: "13px", color: "#744210" }}>
                              {pendingMissing.length} Missing Technical Requirements Require Your Permission (Option A)
                            </strong>
                            <p style={{ margin: "2px 0 0", fontSize: "11px", color: "#8c5b1b" }}>
                              Job description mentions {pendingMissing.slice(0, 4).map((t) => t.tool).join(", ")}{pendingMissing.length > 4 ? ` +${pendingMissing.length - 4} more` : ""}. Review explanations and grant permission to add to verified match or skip.
                            </p>
                          </div>
                        </div>
                        <button
                          type="button"
                          onClick={() => setIsPermissionModalOpen(true)}
                          style={{ padding: "8px 14px", fontSize: "11px", fontWeight: 600, background: "#996a13", color: "white", border: "none", borderRadius: "6px", cursor: "pointer", display: "flex", alignItems: "center", gap: "6px" }}
                        >
                          <Sparkles size={14} /> Review & Grant Permission
                        </button>
                      </div>
                    );
                  })()}

                  <DiscoveredToolsSection
                    workflowId={workflow.id}
                    analysis={workflow.analysis}
                    onToolUpdate={(updated) => {
                      setWorkflow(updated);
                    }}
                    onOpenModal={() => setIsPermissionModalOpen(true)}
                  />

                  <MissingTechPermissionModal
                    isOpen={isPermissionModalOpen}
                    onClose={() => setIsPermissionModalOpen(false)}
                    workflowId={workflow.id}
                    analysis={workflow.analysis}
                    onToolUpdate={(updated) => {
                      setWorkflow(updated);
                    }}
                  />

                  <section className="card insight-card">
                    <div className="insight-heading"><div><span className="mini-icon mini-blue"><Activity size={16} /></span><div><h3>What the job description is asking for</h3><p>{workflow.analysis.method}</p></div></div><span className="estimate-tag"><CircleHelp size={13} /> ESTIMATED MATCH</span></div>
                    <div className="keyword-groups">
                      <div><div className="keyword-title"><span className="keyword-dot match-dot" />Found in your resume <span className="keyword-count">{workflow.analysis.matched_keywords.length}</span></div><div className="keyword-list">{workflow.analysis.matched_keywords.slice(0, 20).map((word) => <span className="keyword keyword-found" key={word}>{word}</span>)}{workflow.analysis.matched_keywords.length === 0 && <span className="empty-copy">No matching terms found yet.</span>}</div></div>
                      <div><div className="keyword-title"><span className="keyword-dot missing-dot" />Not found or not explicit <span className="keyword-count">{workflow.analysis.missing_keywords.length}</span></div><div className="keyword-list">{(showAllMissingKeywords ? workflow.analysis.missing_keywords : workflow.analysis.missing_keywords.slice(0, 20)).map((word) => <span className="keyword keyword-missing" key={word}>{word}</span>)}{workflow.analysis.missing_keywords.length > 20 && <button className="text-action" onClick={() => setShowAllMissingKeywords(!showAllMissingKeywords)}>{showAllMissingKeywords ? "Show fewer terms" : `Show all ${workflow.analysis.missing_keywords.length} terms`}</button>}{workflow.analysis.missing_keywords.length === 0 && <span className="empty-copy">Every analyzed term appears in your resume.</span>}</div></div>
                    </div>
                    <div className="gap-note"><Lightbulb size={15} /><span>A missing keyword may be in your experience but expressed differently. Only add it if it accurately describes your background.</span></div>
                    <CategoryBreakdownSection analysis={workflow.analysis} title="Original resume weighted breakdown" />
                  </section>

                  <RequirementEvidenceSection analysis={workflow.analysis} title="Original requirement evidence classification" />

                  <section className="card final-card resume-preview-card">
                    <div className="card-label-row">
                      <div><span className="mini-icon mini-blue"><FileText size={16} /></span><span className="card-eyebrow">RESUME DOCUMENT PREVIEW</span></div>
                      <span className="muted-label">Original formatting retained</span>
                    </div>
                    <p className="document-preview-note">The original is previewed from your saved DOCX or PDF file and is never changed. Before final export, the tailored pane is a text-only review preview; the final DOCX keeps the original layout as far as possible.</p>
                    <div className="preview-tabs" role="tablist" aria-label="Resume preview options">{([["comparison", "Side by side"], ["original", "Original resume"], ["tailored", "Tailored resume"]] as const).map(([id, label]) => <button key={id} role="tab" aria-selected={preview === id} className={preview === id ? "preview-tab-active" : ""} onClick={() => setPreview(id)}>{label}</button>)}</div>
                    <div className="document-preview-wrapper-box">
                      {preview === "comparison" ? (
                        <div className="comparison-grid">
                          <DocumentPreviewPanel title="Original resume" badge="UNCHANGED" filename={workflow.filename} sourceUrl={`/api/workflows/${workflow.id}/original`} />
                          {workflow.finalized && workflow.download_urls ? <DocumentPreviewPanel title="Tailored resume" badge="APPROVED COPY" filename={`${workflow.filename.replace(/\.[^.]+$/, "")}-resumefit.docx`} sourceUrl={workflow.download_urls.docx} highlighted /> : <ResumeTextPanel title="Text-only review preview" text={tailoredDisplay || workflow.original_text} badge={`${approvedCount} APPROVED EDIT${approvedCount === 1 ? "" : "S"}`} highlighted />}
                        </div>
                      ) : preview === "original" ? (
                        <div className="single-preview-bounded">
                          <DocumentPreviewPanel title="Original resume — unchanged" badge="ORIGINAL" filename={workflow.filename} sourceUrl={`/api/workflows/${workflow.id}/original`} />
                        </div>
                      ) : workflow.finalized && workflow.download_urls ? (
                        <div className="single-preview-bounded">
                          <DocumentPreviewPanel title="Tailored resume — approved changes only" badge="TAILORED COPY" filename={`${workflow.filename.replace(/\.[^.]+$/, "")}-resumefit.docx`} sourceUrl={workflow.download_urls.docx} highlighted />
                        </div>
                      ) : (
                        <div className="single-preview-bounded">
                          <ResumeTextPanel title="Text-only review preview" text={tailoredDisplay || workflow.original_text} badge={`${approvedCount} APPROVED EDIT${approvedCount === 1 ? "" : "S"}`} highlighted />
                        </div>
                      )}
                    </div>
                  </section>

                  <section className="review-section">
                    <div className="review-heading"><div><span className="section-kicker">NOTHING CHANGES WITHOUT YOU</span><h2>Review suggested changes <span className="suggestion-count">{workflow.suggestions.length}</span></h2><p>Approve, edit or reject every suggestion before building the final resume.</p></div><div className="bulk-actions"><button className="button button-secondary" onClick={() => void bulkSuggestions("reject")} disabled={busyAction !== ""}><XCircle size={15} />Reject all</button><button className="button button-approve-all" onClick={() => void bulkSuggestions("approve")} disabled={busyAction !== "" || pendingCount === 0}><CheckCheck size={15} />Approve all</button></div></div>
                    <div className="review-stats"><span><i className="status-dot dot-pending" />{pendingCount} to review</span><span><i className="status-dot dot-approved" />{approvedCount} approved</span><span><i className="status-dot dot-rejected" />{rejectedCount} rejected</span><span className="review-scope">Changes apply to your DOCX copy, never the original</span></div>
                    {workflow.suggestions.length === 0 ? <div className="card empty-suggestions"><div className="empty-icon"><MessageSquareText size={21} /></div><h3>No editable resume text was extracted</h3><p>There is no source paragraph to safely suggest or edit. Try a text-based PDF or DOCX with selectable text, or review the original document.</p></div> : (
                      <div className="suggestion-list">{workflow.suggestions.map((suggestion, index) => <SuggestionCard key={`${suggestion.id}-${index}`} suggestion={suggestion} busy={busyAction !== ""} saving={busyAction === suggestion.id} onAction={updateSuggestion} onTailor={tailorSection} />)}</div>
                    )}
                  </section>

                  <section className="card instructions-card">
                    <div className="instructions-icon"><MessageSquareText size={17} /></div><div className="instructions-body"><div><span className="section-kicker">MAKE IT YOURS</span><h3>Have something specific in mind?</h3><p>Tell us what to focus on. Applying instructions refreshes the suggestion queue and clears previous decisions; nothing changes in your resume automatically.</p></div>
                      <form onSubmit={applyInstructions}><textarea value={instructions} onChange={(event) => setInstructions(event.target.value)} placeholder="e.g. Shorten my professional summary and highlight my verified SQL experience. Keep all facts unchanged." maxLength={2000} aria-label="Additional tailoring instructions" /><div className="instructions-footer"><span>{instructions.length} / 2,000</span><button type="submit" className="button button-primary button-small" disabled={busyAction !== "" || !instructions.trim()}>{busyAction === "instructions" ? <LoaderCircle className="spin" size={15} /> : <Send size={14} />}Apply my instructions</button></div></form>
                    </div>
                  </section>

                  {workflow.finalized && tailoredDisplay ? (
                    <section className="card final-card">
                      <div className="final-heading"><div className="final-check"><Check size={18} /></div><div><span className="section-kicker">ALL SET</span><h2>Your tailored resume is ready</h2><p>Your unchanged original is safe. Only approved edits were applied.</p></div></div>
                      <div className="final-score-compare"><div><span>Original</span><strong>{workflow.analysis.score}/100</strong></div><ArrowRight size={16} /><div><span>Tailored</span><strong>{liveScore}/100</strong></div><div className={`final-score-delta ${scoreVariance >= 0 ? "" : "improvement-down"}`}>{scoreVariance >= 0 ? "+" : ""}{scoreVariance} pts</div></div>
                      <div className="remaining-gaps"><span className="remaining-icon"><Lightbulb size={15} /></span><div><strong>Still not explicit in the tailored text</strong><div className="keyword-list">{(showAllMissingKeywords ? liveScoreDetails?.missing_keywords ?? [] : (liveScoreDetails?.missing_keywords ?? []).slice(0, 14)).map((word) => <span className="keyword keyword-missing" key={word}>{word}</span>)}{(liveScoreDetails?.missing_keywords.length ?? 0) > 14 && <button className="text-action" onClick={() => setShowAllMissingKeywords(!showAllMissingKeywords)}>{showAllMissingKeywords ? "Show fewer terms" : `Show all ${liveScoreDetails?.missing_keywords.length} terms`}</button>}{!liveScoreDetails?.missing_keywords.length && <span className="empty-copy">All analyzed JD terms are present. This still is not an ATS guarantee.</span>}</div></div></div>
                      {liveScoreDetails && (
                        <>
                          <CategoryBreakdownSection analysis={liveScoreDetails} title="Tailored resume weighted breakdown" />
                          <RequirementEvidenceSection analysis={liveScoreDetails} title="Tailored requirement evidence classification" />
                        </>
                      )}
                      {workflow.layout_review.map((check) => <div className="limitation-note" key={check}><TriangleAlert size={14} />{check}</div>)}
                      {workflow.limitations.map((limitation) => <div className="limitation-note" key={limitation}><CircleHelp size={14} />{limitation}</div>)}
                    </section>
                  ) : (
                    <div className="generate-bar"><div><div className="generate-icon"><FileCheck2 size={17} /></div><div><strong>Ready to build your tailored resume?</strong><span>{pendingCount > 0 ? `Resolve all ${pendingCount} pending suggestion${pendingCount === 1 ? "" : "s"} first.` : `${approvedCount} approved · ${rejectedCount} rejected · your original will remain unchanged.`}</span></div></div><button className="button button-primary" disabled={pendingCount > 0 || busyAction !== ""} onClick={() => void generateFinalResume()}>{busyAction === "finalize" ? <LoaderCircle className="spin" size={16} /> : <Sparkles size={16} />}{busyAction === "finalize" ? "Building…" : "Generate tailored resume"}<ArrowRight size={15} /></button></div>
                  )}
                  <div className="download-row" aria-label="Resume downloads">
                    <button className="button button-download" disabled={busyAction !== ""} onClick={() => void downloadFinalResume("docx")}>{busyAction === "download-docx" ? <LoaderCircle className="spin" size={16} /> : <FileDown size={16} />}Download DOCX</button>
                    <button className="button button-download-secondary" disabled={busyAction !== ""} onClick={() => void downloadFinalResume("pdf")}>{busyAction === "download-pdf" ? <LoaderCircle className="spin" size={16} /> : <Download size={16} />}Download PDF</button>
                    {workflow.download_urls && <a className="original-link" href={workflow.download_urls.original} download><FileText size={15} />Download unchanged original</a>}
                    {workflow.finalized && <button className="return-edit" onClick={() => { setWorkflow({ ...workflow, finalized: false }); setNotice("Return to your suggestions and make another round of edits."); }}><Pencil size={14} />Return to editing</button>}
                  </div>
                </>
              )}
              {!workflow && history.length > 0 && <RecentActivity history={history.slice(0, 3)} onSelect={(id) => void loadWorkflow(id)} onHistory={() => choosePage("history")} />}
            </>
          )}

          {page === "profile" && (
            <CandidateProfilePage
              profile={candidateProfile}
              onSave={async (updated) => {
                const saved = await api<CandidateProfile>("/profile", {
                  method: "PUT",
                  headers: { "Content-Type": "application/json" },
                  body: JSON.stringify(updated),
                });
                setCandidateProfile(saved);
                setNotice("Candidate profile saved securely. It will be remembered across sessions.");
              }}
              onExtract={async (file) => {
                const body = new FormData();
                if (file) body.set("file", file);
                const extracted = await api<CandidateProfile>("/profile/extract", {
                  method: "POST",
                  body,
                });
                setCandidateProfile(extracted);
                setNotice("Extracted profile details from your resume. Check each field and save.");
              }}
              onTailorWithProfile={() => {
                setUseProfileForTailoring(true);
                setSelectedFile(null);
                choosePage("dashboard");
                setNotice("Using your candidate profile as the source resume. Paste the job description to start tailoring.");
              }}
            />
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

function SuggestionCard({
  suggestion,
  busy,
  saving,
  onAction,
  onTailor,
}: {
  suggestion: Suggestion;
  busy: boolean;
  saving: boolean;
  onAction: (id: string, action: SuggestionAction, text?: string) => Promise<boolean>;
  onTailor?: (suggestion: Suggestion) => Promise<void>;
}) {
  const [editing, setEditing] = useState(false);
  const [draft, setDraft] = useState(suggestion.proposed_text);
  const [showDifference, setShowDifference] = useState(false);
  const Icon = sectionIcons[suggestion.section] ?? FileText;
  const stateLabel = suggestion.status === "pending" ? "Awaiting your review" : suggestion.status === "approved" ? "Approved" : "Rejected";
  const hasChange = suggestion.proposed_text.trim() !== suggestion.original_text.trim();

  async function saveEdit() {
    if (await onAction(suggestion.id, "edit", draft)) setEditing(false);
  }
  function toggleEdit() {
    setDraft(suggestion.proposed_text);
    setEditing(!editing);
  }

  return (
    <article className={`suggestion-card suggestion-${suggestion.status}`}>
      <div className="suggestion-topline"><div className="suggestion-section"><span className={`suggestion-icon ${suggestion.status === "approved" ? "suggestion-icon-approved" : suggestion.status === "rejected" ? "suggestion-icon-rejected" : ""}`}><Icon size={16} /></span><div><h3>{suggestion.section}</h3><span className={`suggestion-state state-${suggestion.status}`}>{suggestion.status === "approved" ? <Check size={11} /> : suggestion.status === "rejected" ? <X size={11} /> : <Clock3 size={11} />}{stateLabel}</span></div></div><button className={`source-chip ${suggestion.source === "ollama" || suggestion.source === "gemini" ? "source-ai" : ""}`} title={suggestion.source === "manual" ? "Manual editing option. No safe AI rewrite was returned." : suggestion.source === "demo" ? "Rule-based demo suggestion. No AI-generated rewrite." : suggestion.source === "gemini" ? "Generated using Google Gemini before switching to local Ollama." : "Generated using the server-configured local Ollama API."}>{suggestion.source === "manual" ? <><Pencil size={12} />Manual edit</> : suggestion.source === "demo" ? <><Lightbulb size={12} />Evidence-led</> : <><Sparkles size={12} />AI suggestion</>}</button></div>
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
      <div className="suggestion-actions">{editing ? <><button className="button button-secondary action-cancel" disabled={busy} onClick={toggleEdit}><X size={14} />Cancel</button><button className="button button-primary action-save" disabled={busy || !draft.trim()} onClick={() => void saveEdit()}>{saving ? <LoaderCircle className="spin" size={14} /> : <Check size={14} />}Save edit</button></> : <><button className="text-action" disabled={busy} onClick={() => { setShowDifference(!showDifference); }}><Eye size={14} />{showDifference ? "Hide difference" : "View difference"}</button><div className="action-divider" /><button className="text-action" disabled={busy} onClick={() => void onTailor?.(suggestion)} title="Run fast targeted AI rewrite for this section"><Sparkles size={13} />{saving ? "Tailoring..." : "AI Re-tailor"}</button><div className="action-divider" /><button className="text-action" disabled={busy} onClick={toggleEdit}><Pencil size={13} />Edit</button><button className="button button-secondary action-reject" disabled={busy || suggestion.status === "rejected"} onClick={() => void onAction(suggestion.id, "reject")}><X size={14} />Reject</button><button className="button button-approve" disabled={busy || suggestion.status === "approved"} onClick={() => void onAction(suggestion.id, "approve")}>{saving ? <LoaderCircle className="spin" size={14} /> : <Check size={14} />}Approve</button>{suggestion.status !== "pending" && <button className="undo-button" disabled={busy} title="Undo decision and restore the original suggestion" onClick={() => void onAction(suggestion.id, "undo")}><RotateCcw size={14} />Undo</button>}</>}</div>
    </article>
  );
}

function ResumeTextPanel({ title, text, badge, highlighted = false }: { title: string; text: string; badge: string; highlighted?: boolean }) {
  return (
    <div className={`resume-text-panel ${highlighted ? "resume-text-highlighted" : ""}`}>
      <div className="resume-panel-head"><strong>{title}</strong><span>{badge}</span></div>
      <div className="text-preview-box">
        <pre>{text || "No text available for preview."}</pre>
      </div>
    </div>
  );
}

function DocumentPreviewPanel({
  title,
  badge,
  filename,
  sourceUrl,
  sourceFile,
  highlighted = false,
}: {
  title: string;
  badge: string;
  filename: string;
  sourceUrl: string;
  sourceFile?: File;
  highlighted?: boolean;
}) {
  const containerRef = useRef<HTMLDivElement>(null);
  const documentRef = useRef<HTMLDivElement>(null);
  const styleRef = useRef<HTMLDivElement>(null);
  const [loadingPreview, setLoadingPreview] = useState(true);
  const [previewError, setPreviewError] = useState("");
  const [previewUrl, setPreviewUrl] = useState("");
  const [docDimensions, setDocDimensions] = useState<{ width: number; height: number }>({ width: 816, height: 1056 });
  const [scale, setScale] = useState<number>(0.5);
  const [pageCount, setPageCount] = useState<number>(1);
  const [currentPage, setCurrentPage] = useState<number>(0);
  const isPdf = filename.toLowerCase().endsWith(".pdf");

  const updateScale = useCallback(() => {
    if (!containerRef.current) return;
    const container = containerRef.current;
    const containerWidth = container.clientWidth;
    const containerHeight = container.clientHeight || 560;
    if (containerWidth <= 0 || containerHeight <= 0) return;

    const sections = (documentRef.current?.querySelectorAll("section.docx") ||
      documentRef.current?.querySelectorAll("section")) as NodeListOf<HTMLElement> | undefined;

    if (sections && sections.length > 0) {
      if (sections.length !== pageCount) {
        setPageCount(sections.length);
      }
      if (sections.length > 1) {
        sections.forEach((sec, idx) => {
          sec.style.display = idx === currentPage ? "block" : "none";
          sec.style.margin = "0";
        });
      }
    }

    const activeSection = (
      sections && sections.length > 0
        ? (sections[currentPage] || sections[0])
        : (documentRef.current?.firstElementChild as HTMLElement | null)
    );

    const naturalWidth = (activeSection && activeSection.offsetWidth > 0)
      ? activeSection.offsetWidth
      : 816;
    const naturalHeight = (activeSection && activeSection.offsetHeight > 0)
      ? Math.max(activeSection.offsetHeight, activeSection.scrollHeight, 1056)
      : 1056;

    setDocDimensions({ width: naturalWidth, height: naturalHeight });

    const paddingMargin = 20;
    const availW = Math.max(50, containerWidth - paddingMargin);
    const availH = Math.max(50, containerHeight - paddingMargin);

    const scaleW = availW / naturalWidth;
    const scaleH = availH / naturalHeight;
    const newScale = Math.min(scaleW, scaleH);

    setScale(newScale);
  }, [pageCount, currentPage]);

  useEffect(() => {
    if (!containerRef.current) return;
    const observer = new ResizeObserver(() => {
      requestAnimationFrame(updateScale);
    });
    observer.observe(containerRef.current);
    return () => {
      observer.disconnect();
    };
  }, [updateScale]);

  useEffect(() => {
    const controller = new AbortController();
    let cancelled = false;
    let objectUrl = "";
    setLoadingPreview(true);
    setPreviewError("");
    setPreviewUrl("");
    setCurrentPage(0);

    async function loadDocument() {
      try {
        let blob: Blob;
        if (sourceFile) {
          blob = sourceFile;
        } else {
          const response = await fetch(sourceUrl, { signal: controller.signal });
          if (!response.ok) {
            let message = `Could not load the original resume (${response.status}).`;
            try {
              const body = (await response.json()) as { detail?: string };
              if (body.detail) message = body.detail;
            } catch {
              message = `Could not load the original resume (${response.status}).`;
            }
            throw new Error(message);
          }
          blob = await response.blob();
        }
        if (cancelled) return;
        objectUrl = URL.createObjectURL(blob);
        setPreviewUrl(objectUrl);
        if (isPdf) {
          setLoadingPreview(false);
          return;
        }
        if (!documentRef.current || !styleRef.current) {
          throw new Error("The document preview area is unavailable. Download the original file to view it.");
        }
        await renderAsync(blob, documentRef.current, styleRef.current, {
          breakPages: true,
          renderHeaders: true,
          renderFooters: true,
        });
        if (!cancelled) {
          requestAnimationFrame(() => {
            updateScale();
            setLoadingPreview(false);
          });
          setTimeout(() => {
            if (!cancelled) updateScale();
          }, 80);
        }
      } catch (reason) {
        if (cancelled || controller.signal.aborted) return;
        setPreviewError(reason instanceof Error ? reason.message : String(reason));
        setLoadingPreview(false);
      }
    }

    void loadDocument();
    return () => {
      cancelled = true;
      controller.abort();
      if (objectUrl) URL.revokeObjectURL(objectUrl);
      documentRef.current?.replaceChildren();
      styleRef.current?.replaceChildren();
    };
  }, [sourceFile, sourceUrl, isPdf, updateScale]);

  return (
    <div className={`resume-text-panel document-preview-panel ${highlighted ? "resume-text-highlighted" : ""}`}>
      <div className="resume-panel-head">
        <strong>{title}</strong>
        <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
          {pageCount > 1 && (
            <div className="document-page-controls">
              <button
                type="button"
                className="page-nav-btn"
                onClick={() => setCurrentPage((p) => Math.max(0, p - 1))}
                disabled={currentPage === 0}
                title="Previous page"
              >
                ‹
              </button>
              <span className="page-nav-text">
                {currentPage + 1} / {pageCount}
              </span>
              <button
                type="button"
                className="page-nav-btn"
                onClick={() => setCurrentPage((p) => Math.min(pageCount - 1, p + 1))}
                disabled={currentPage === pageCount - 1}
                title="Next page"
              >
                ›
              </button>
            </div>
          )}
          <span>{badge}</span>
        </div>
      </div>
      {(sourceUrl || previewUrl) && (
        <a className="document-download-link" href={previewUrl || sourceUrl} download={filename}>
          <FileDown size={13} />
          {highlighted ? "Download tailored file" : "Download unchanged file"}
        </a>
      )}
      <div className="document-preview-box" ref={containerRef}>
        {loadingPreview && (
          <div className="document-preview-status">
            <LoaderCircle className="spin" size={17} />
            Loading original document preview…
          </div>
        )}
        {previewError && (
          <div className="document-preview-status document-preview-error" role="alert">
            {previewError} Download the original file to inspect its formatting.
          </div>
        )}
        {!previewError && isPdf && previewUrl && (
          <iframe
            className="document-preview-pdf"
            src={`${previewUrl}#view=Fit&toolbar=0&navpanes=0`}
            title={`${filename} original document preview`}
          />
        )}
        {!isPdf && (
          <div
            className="document-scaled-viewport"
            style={{
              width: `${Math.round(docDimensions.width * scale)}px`,
              height: `${Math.round(docDimensions.height * scale)}px`,
              opacity: loadingPreview ? 0 : 1,
              transition: "opacity 0.2s ease-in-out",
            }}
          >
            <div
              className="document-unscaled-canvas"
              style={{
                width: `${docDimensions.width}px`,
                height: `${docDimensions.height}px`,
                transform: `scale(${scale})`,
                transformOrigin: "top left",
              }}
            >
              <div ref={styleRef} />
              <div ref={documentRef} />
            </div>
          </div>
        )}
      </div>
    </div>
  );
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
  return <div className="settings-page"><div className="collection-header"><div><span className="section-kicker">YOUR WORKSPACE</span><h1>Settings</h1><p>Understand how your files, suggestions, and workspace are handled.</p></div></div><div className="settings-grid">  <section className="card settings-card"><div className="settings-title"><span className="settings-icon settings-teal"><Sparkles size={17} /></span><div><h3>Analysis & AI</h3><p>How your match estimate and resume suggestions are generated.</p></div></div><div className="settings-detail"><span>Current analysis mode</span><span className={`mode-chip ${demoMode ? "mode-demo" : "mode-ai"}`}><span />{demoMode ? "Transparent demo mode" : "Ollama configured"}</span></div><p className="settings-explainer">{demoMode ? "The local Ollama endpoint is not configured, so matching uses visible keyword coverage and suggestions refer only to content already in your resume. Set OLLAMA_API_URL for the backend to enable local AI-generated suggestions." : "The backend sends resume text and the job description to the configured Ollama API endpoint (localhost:11434 by default). No Google Gemini API key or external model provider is used by this workflow. Check each request result, compare every proposal to your source, and approve it only if accurate."}</p><div className="settings-note"><CircleHelp size={15} />Match percentages are estimates of distinct JD vocabulary found in resume text, not actual ATS scores.</div></section>
    <section className="card settings-card"><div className="settings-title"><span className="settings-icon settings-purple"><LockKeyhole size={17} /></span><div><h3>Privacy & storage</h3><p>Your personal files and saved tailoring sessions.</p></div></div><div className="privacy-stats"><div><strong>{count}</strong><span>saved {count === 1 ? "session" : "sessions"}</span></div><span className="privacy-divider" /><div><span className="local-chip"><ShieldCheck size={14} />Local device storage</span></div></div><p className="settings-explainer">By default, original resumes, tailored copies, and history are stored in the local .resume_fit_data folder and a local SQLite database. Set DATABASE_URL to PostgreSQL for deployment. This starter is designed for one local user; add authentication before exposing it to multiple users.</p><button className="settings-refresh" onClick={onRefresh}><RefreshCw size={14} />Refresh connection</button></section>
    <section className="card settings-card settings-docs"><div className="settings-title"><span className="settings-icon settings-blue"><FileBadge size={17} /></span><div><h3>Document compatibility</h3><p>Know what to expect when saving your resume.</p></div></div><div className="compat-row"><span className="compat-icon compat-docx"><FileText size={16} /></span><div><strong>DOCX</strong><span>Edits existing paragraphs to retain styles, layout and alignment. Review pagination after long text replacements.</span></div><span className="compat-tag">BEST PRESERVATION</span></div><div className="compat-row"><span className="compat-icon compat-pdf"><FileDown size={16} /></span><div><strong>PDF</strong><span>Scanned PDFs need OCR. PDF-source exports use text-only layout; DOCX-to-PDF fidelity depends on LibreOffice.</span></div><span className="compat-tag compat-limited">LAYOUT LIMITS</span></div><div className="settings-note"><TriangleAlert size={15} />Always open and visually check your downloaded files before sending an application.</div></section>
  </div></div>;
}

function CategoryBreakdownSection({ analysis, title = "Deterministic Weighted Category Scores" }: { analysis: ScoreDetails; title?: string }) {
  const categories = analysis.category_scores || [];
  if (categories.length === 0) return null;

  return (
    <div style={{ marginTop: "16px", paddingTop: "14px", borderTop: "1px solid #edf1f3" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "8px" }}>
        <strong style={{ fontSize: "11px", color: "#2d4452" }}>
          {title}
        </strong>
        <span style={{ fontSize: "8.5px", color: "#69827a" }}>
          Deterministic weights (100% total)
        </span>
      </div>
      <div className="categories-breakdown">
        {categories.map((cat) => (
          <div className="category-card" key={cat.category}>
            <div className="category-header">
              <span className="category-title">{cat.category}</span>
              <span className="category-weight">{cat.weight}% weight</span>
            </div>
            <div className="category-bar">
              <div className="category-fill" style={{ width: `${Math.min(100, Math.round(cat.score))}%` }} />
            </div>
            <div className="category-stats">
              <span>{Math.round(cat.score)}% match ({cat.weighted_score.toFixed(1)} / {cat.weight} pts)</span>
              <span>{cat.matched_count} of {cat.total_count} criteria evidenced</span>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

function RequirementEvidenceSection({ analysis, title = "Job Requirement Evidence Classification" }: { analysis: ScoreDetails; title?: string }) {
  const items = analysis.evidence_breakdown || [];
  const [filter, setFilter] = useState<string>("all");

  if (items.length === 0) return null;

  const filtered = filter === "all" ? items : items.filter((item) => item.classification === filter);

  return (
    <section className="card evidence-card">
      <div className="insight-heading">
        <div>
          <span className="mini-icon mini-green"><ListChecks size={16} /></span>
          <div>
            <h3>{title}</h3>
            <p>Every requirement classified based strictly on verifiable evidence in your resume.</p>
          </div>
        </div>
        <span className="estimate-tag"><ShieldCheck size={13} /> EVIDENCE-BASED</span>
      </div>

      <div className="evidence-filters">
        <button
          type="button"
          className={`evidence-filter-btn ${filter === "all" ? "active" : ""}`}
          onClick={() => setFilter("all")}
        >
          All ({items.length})
        </button>
        <button
          type="button"
          className={`evidence-filter-btn ${filter === "fully evidenced" ? "active" : ""}`}
          onClick={() => setFilter("fully evidenced")}
        >
          Fully Evidenced ({items.filter((i) => i.classification === "fully evidenced").length})
        </button>
        <button
          type="button"
          className={`evidence-filter-btn ${filter === "partially evidenced" ? "active" : ""}`}
          onClick={() => setFilter("partially evidenced")}
        >
          Partially Evidenced ({items.filter((i) => i.classification === "partially evidenced").length})
        </button>
        <button
          type="button"
          className={`evidence-filter-btn ${filter === "not evidenced" ? "active" : ""}`}
          onClick={() => setFilter("not evidenced")}
        >
          Not Evidenced ({items.filter((i) => i.classification === "not evidenced").length})
        </button>
        <button
          type="button"
          className={`evidence-filter-btn ${filter === "uncertain" ? "active" : ""}`}
          onClick={() => setFilter("uncertain")}
        >
          Uncertain ({items.filter((i) => i.classification === "uncertain").length})
        </button>
      </div>

      <div style={{ overflowX: "auto" }}>
        <table className="evidence-table">
          <thead>
            <tr>
              <th style={{ width: "35%" }}>Requirement</th>
              <th style={{ width: "20%" }}>Category</th>
              <th style={{ width: "20%" }}>Classification</th>
              <th style={{ width: "25%" }}>Resume Evidence</th>
            </tr>
          </thead>
          <tbody>
            {filtered.map((item, idx) => {
              const badgeClass =
                item.classification === "fully evidenced"
                  ? "badge-fully-evidenced"
                  : item.classification === "partially evidenced"
                  ? "badge-partially-evidenced"
                  : item.classification === "not evidenced"
                  ? "badge-not-evidenced"
                  : "badge-uncertain";

              return (
                <tr key={`${item.requirement}-${idx}`}>
                  <td>
                    <strong style={{ color: "#2f4654" }}>{item.requirement}</strong>
                  </td>
                  <td>
                    <span className="category-weight">{item.category}</span>
                  </td>
                  <td>
                    <span className={`badge-evidence ${badgeClass}`}>
                      {item.classification === "fully evidenced" && <Check size={11} />}
                      {item.classification === "partially evidenced" && <Sparkles size={11} />}
                      {item.classification === "not evidenced" && <X size={11} />}
                      {item.classification === "uncertain" && <CircleHelp size={11} />}
                      {item.classification.toUpperCase()}
                    </span>
                  </td>
                  <td>
                    {item.snippets && item.snippets.length > 0 ? (
                      <div className="evidence-snippets">
                        {item.snippets.map((snip, sIdx) => (
                          <div key={sIdx}>“{snip}”</div>
                        ))}
                      </div>
                    ) : (
                      <span style={{ color: "#98a3a8", fontStyle: "italic", fontSize: "8px" }}>
                        No direct evidence found in resume text.
                      </span>
                    )}
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </section>
  );
}

function RoleEligibilityBanner({ analysis }: { analysis: ScoreDetails }) {
  const roleMatch = analysis.tier1_role_match;
  const stage = analysis.career_stage || "Fresher / Entry-Level";
  const coreDomain = analysis.core_domain || roleMatch?.domain || "Full Stack & Web Development";

  return (
    <section className="card" style={{ marginBottom: "16px", padding: "14px 18px", background: "linear-gradient(135deg, #f7fafc 0%, #edf7f5 100%)", border: "1px solid #d4eae4" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: "10px" }}>
        <div>
          <div style={{ display: "flex", alignItems: "center", gap: "8px", marginBottom: "4px", flexWrap: "wrap" }}>
            <span style={{ fontSize: "9px", fontWeight: 700, letterSpacing: "0.5px", textTransform: "uppercase", color: "#196857" }}>
              TIER 1 ROLE ALIGNMENT
            </span>
            <span style={{ fontSize: "9px", padding: "2px 8px", borderRadius: "12px", background: "#dff4ee", color: "#166453", fontWeight: 600 }}>
              {roleMatch?.eligibility_status || "High Eligibility"}
            </span>
            <span style={{ fontSize: "9px", padding: "2px 8px", borderRadius: "12px", background: "#e8effc", color: "#2553a8", fontWeight: 600 }}>
              {stage}
            </span>
            <span style={{ fontSize: "9px", padding: "2px 8px", borderRadius: "12px", background: "#fef3c7", color: "#92400e", fontWeight: 600 }}>
              CORE DOMAIN: {coreDomain}
            </span>
          </div>
          <h3 style={{ margin: "2px 0 4px", fontSize: "16px", color: "#1a3848" }}>
            {roleMatch?.matched_role || "Full Stack Developer"} <span style={{ fontSize: "13px", fontWeight: 400, color: "#567180" }}>in {coreDomain}</span>
          </h3>
          <p style={{ margin: 0, fontSize: "11px", color: "#4f6874", lineHeight: 1.45 }}>
            Evaluating strictly against verified technical taxonomy. Unverified corporate buzzwords and filler phrases are completely excluded.
          </p>
        </div>
        {roleMatch?.eligible_clusters && (
          <div style={{ display: "flex", flexDirection: "column", alignItems: "flex-end", gap: "4px" }}>
            <span style={{ fontSize: "8.5px", color: "#6c8591" }}>Eligible Profile Domains</span>
            <div style={{ display: "flex", gap: "4px", flexWrap: "wrap", justifyContent: "flex-end" }}>
              {roleMatch.eligible_clusters.slice(0, 3).map((cl) => (
                <span key={cl} style={{ fontSize: "8px", padding: "1px 6px", borderRadius: "4px", background: "#ffffff", border: "1px solid #dbe6e4", color: "#365664" }}>
                  {cl}
                </span>
              ))}
            </div>
          </div>
        )}
      </div>

      {roleMatch?.warning && (
        <div style={{ marginTop: "12px", padding: "10px 14px", background: "#fff5f5", border: "1px solid #fed7d7", borderRadius: "8px", display: "flex", gap: "10px", alignItems: "flex-start", width: "100%" }}>
          <TriangleAlert size={18} color="#c53030" style={{ flexShrink: 0, marginTop: "2px" }} />
          <div>
            <strong style={{ fontSize: "12px", color: "#c53030", display: "block", marginBottom: "2px" }}>
              Specialized / Incompatible Domain Alert
            </strong>
            <span style={{ fontSize: "11px", color: "#9b2c2c", lineHeight: 1.45 }}>
              {roleMatch.warning}
            </span>
          </div>
        </div>
      )}
    </section>
  );
}

function DiscoveredToolsSection({
  workflowId,
  analysis,
  onToolUpdate,
  onOpenModal,
}: {
  workflowId: string;
  analysis: ScoreDetails;
  onToolUpdate: (updatedWorkflow: Workflow) => void;
  onOpenModal?: () => void;
}) {
  const tools = (analysis.missing_technologies && analysis.missing_technologies.length > 0)
    ? analysis.missing_technologies
    : (analysis.discovered_tools || []);
  const [loadingTool, setLoadingTool] = useState<string | null>(null);

  if (tools.length === 0) return null;

  async function handleConfirm(toolName: string, status: "approved" | "related" | "skipped") {
    setLoadingTool(toolName);
    try {
      const res = await api<Workflow>(`/workflows/${workflowId}/confirm-tool`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ tool: toolName, status }),
      });
      onToolUpdate(res);
    } finally {
      setLoadingTool(null);
    }
  }

  const pendingTools = tools.filter((t) => t.status === "pending");
  const confirmedTools = tools.filter((t) => t.status === "approved" || t.status === "related");

  return (
    <section className="card" style={{ marginBottom: "16px", padding: "16px 18px", border: "1px solid #e1e9ea" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "12px", flexWrap: "wrap", gap: "8px" }}>
        <div>
          <div style={{ display: "flex", alignItems: "center", gap: "7px" }}>
            <span className="mini-icon mini-teal"><Sparkles size={16} /></span>
            <h3 style={{ margin: 0, fontSize: "14px", color: "#1f3a48" }}>Discovered Job Tools & Missing Technologies</h3>
          </div>
          <p style={{ margin: "3px 0 0", fontSize: "11px", color: "#597381" }}>
            Novel or specialized technologies identified in this job description with plain-English explanations. Confirm which ones you know or have related experience with to boost your verified match.
          </p>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
          {onOpenModal && (
            <button
              type="button"
              onClick={onOpenModal}
              style={{ fontSize: "10.5px", padding: "4px 10px", borderRadius: "6px", background: "#0d9488", color: "white", fontWeight: 600, border: "none", cursor: "pointer", display: "flex", alignItems: "center", gap: "5px" }}
            >
              <Sparkles size={12} /> Permission Review Modal
            </button>
          )}
          <span style={{ fontSize: "10px", padding: "3px 9px", borderRadius: "12px", background: "#f0f4f6", color: "#486675", fontWeight: 600 }}>
            {confirmedTools.length} confirmed · {pendingTools.length} pending
          </span>
        </div>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(280px, 1fr))", gap: "10px" }}>
        {tools.map((item) => {
          const isPending = item.status === "pending";
          const isConfirmed = item.status === "approved" || item.status === "related";
          const isSkipped = item.status === "skipped";
          const isBusy = loadingTool === item.tool;

          return (
            <div
              key={item.tool}
              style={{
                border: isConfirmed ? "1px solid #a3dcce" : isSkipped ? "1px solid #e1e6e8" : "1px solid #dbe5e7",
                background: isConfirmed ? "#f4faf8" : isSkipped ? "#fafbfb" : "#ffffff",
                borderRadius: "8px",
                padding: "10px 12px",
                display: "flex",
                flexDirection: "column",
                justifyContent: "space-between",
                gap: "8px",
              }}
            >
              <div>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "4px" }}>
                  <strong style={{ fontSize: "12px", color: isSkipped ? "#7e929b" : "#1a3644", textTransform: "capitalize" }}>
                    {item.tool}
                  </strong>
                  <span style={{ fontSize: "8.5px", padding: "1px 6px", borderRadius: "4px", background: "#edf3f5", color: "#486674" }}>
                    {item.domain}
                  </span>
                </div>
                <p style={{ margin: 0, fontSize: "10px", color: isSkipped ? "#8b9ea6" : "#4a6472", lineHeight: 1.4 }}>
                  {item.explanation}
                </p>
              </div>

              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", paddingTop: "6px", borderTop: "1px solid #f0f4f5" }}>
                {isPending && (
                  <div style={{ display: "flex", gap: "6px", width: "100%" }}>
                    <button
                      type="button"
                      disabled={isBusy}
                      onClick={() => handleConfirm(item.tool, "approved")}
                      style={{ flex: 1, padding: "5px 6px", fontSize: "9.5px", fontWeight: 600, background: "#1c7b65", color: "white", border: "none", borderRadius: "5px", cursor: "pointer" }}
                    >
                      I know this
                    </button>
                    <button
                      type="button"
                      disabled={isBusy}
                      onClick={() => handleConfirm(item.tool, "related")}
                      style={{ flex: 1, padding: "5px 6px", fontSize: "9.5px", fontWeight: 600, background: "#e8effc", color: "#1d4ed8", border: "1px solid #bfdbfe", borderRadius: "5px", cursor: "pointer" }}
                    >
                      Related exp
                    </button>
                    <button
                      type="button"
                      disabled={isBusy}
                      onClick={() => handleConfirm(item.tool, "skipped")}
                      style={{ padding: "5px 8px", fontSize: "9.5px", background: "#f3f5f6", color: "#667f8c", border: "1px solid #e1e6e8", borderRadius: "5px", cursor: "pointer" }}
                    >
                      Skip
                    </button>
                  </div>
                )}
                {isConfirmed && (
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", width: "100%" }}>
                    <span style={{ fontSize: "9.5px", color: "#156b57", fontWeight: 600, display: "flex", alignItems: "center", gap: "4px" }}>
                      <Check size={12} /> Added to verified match
                    </span>
                    <button
                      type="button"
                      disabled={isBusy}
                      onClick={() => handleConfirm(item.tool, "skipped")}
                      style={{ fontSize: "8.5px", color: "#7a939f", background: "none", border: "none", textDecoration: "underline", cursor: "pointer" }}
                    >
                      Undo
                    </button>
                  </div>
                )}
                {isSkipped && (
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", width: "100%" }}>
                    <span style={{ fontSize: "9.5px", color: "#7b909a", fontStyle: "italic" }}>
                      Skipped (No score penalty)
                    </span>
                    <button
                      type="button"
                      disabled={isBusy}
                      onClick={() => handleConfirm(item.tool, "approved")}
                      style={{ fontSize: "8.5px", color: "#1c7b65", background: "none", border: "none", textDecoration: "underline", cursor: "pointer" }}
                    >
                      Add back
                    </button>
                  </div>
                )}
              </div>
            </div>
          );
        })}
      </div>
    </section>
  );
}

function MissingTechPermissionModal({
  isOpen,
  onClose,
  workflowId,
  analysis,
  onToolUpdate,
}: {
  isOpen: boolean;
  onClose: () => void;
  workflowId: string;
  analysis: ScoreDetails;
  onToolUpdate: (updatedWorkflow: Workflow) => void;
}) {
  const tools = (analysis.missing_technologies && analysis.missing_technologies.length > 0)
    ? analysis.missing_technologies
    : (analysis.discovered_tools || []);
  const [loadingTool, setLoadingTool] = useState<string | null>(null);
  const [batchLoading, setBatchLoading] = useState(false);

  if (!isOpen) return null;

  const pendingTools = tools.filter((t) => t.status === "pending");
  const confirmedTools = tools.filter((t) => t.status === "approved" || t.status === "related");
  const skippedTools = tools.filter((t) => t.status === "skipped");

  async function handleConfirm(toolName: string, status: "approved" | "related" | "skipped") {
    setLoadingTool(toolName);
    try {
      const res = await api<Workflow>(`/workflows/${workflowId}/confirm-tool`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ tool: toolName, status }),
      });
      onToolUpdate(res);
    } finally {
      setLoadingTool(null);
    }
  }

  async function handleBatch(status: "approved" | "skipped") {
    if (pendingTools.length === 0) return;
    setBatchLoading(true);
    try {
      let currentWorkflow: Workflow | null = null;
      for (const item of pendingTools) {
        currentWorkflow = await api<Workflow>(`/workflows/${workflowId}/confirm-tool`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ tool: item.tool, status }),
        });
      }
      if (currentWorkflow) {
        onToolUpdate(currentWorkflow);
      }
    } finally {
      setBatchLoading(false);
    }
  }

  return (
    <div
      style={{
        position: "fixed",
        top: 0,
        left: 0,
        right: 0,
        bottom: 0,
        backgroundColor: "rgba(15, 23, 42, 0.65)",
        backdropFilter: "blur(4px)",
        zIndex: 9999,
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        padding: "20px",
      }}
      onClick={onClose}
    >
      <div
        style={{
          background: "#ffffff",
          borderRadius: "14px",
          maxWidth: "760px",
          width: "100%",
          maxHeight: "88vh",
          display: "flex",
          flexDirection: "column",
          boxShadow: "0 25px 50px -12px rgba(0, 0, 0, 0.25)",
          border: "1px solid #e2e8f0",
          overflow: "hidden",
        }}
        onClick={(e) => e.stopPropagation()}
      >
        <div style={{ padding: "18px 22px", borderBottom: "1px solid #e2e8f0", display: "flex", justifyContent: "space-between", alignItems: "flex-start", background: "#f8fafc" }}>
          <div>
            <div style={{ display: "flex", alignItems: "center", gap: "8px", marginBottom: "4px" }}>
              <span style={{ fontSize: "10px", fontWeight: 700, letterSpacing: "0.5px", textTransform: "uppercase", color: "#0d9488", background: "#ccfbf1", padding: "2px 8px", borderRadius: "12px" }}>
                PERMISSION REQUIRED (OPTION A)
              </span>
              <span style={{ fontSize: "10px", color: "#64748b", fontWeight: 600 }}>
                {pendingTools.length} Pending Decision
              </span>
            </div>
            <h2 style={{ margin: "2px 0 4px", fontSize: "17px", color: "#0f172a", fontWeight: 700 }}>
              Review Missing Technologies in Job Description
            </h2>
            <p style={{ margin: 0, fontSize: "12px", color: "#475569", lineHeight: 1.45 }}>
              The Job Description requests technologies not found in your verified resume. Read each explanation below and grant permission to add only if you have relevant knowledge. No unapproved technologies will ever be hallucinated or added.
            </p>
          </div>
          <button
            type="button"
            onClick={onClose}
            style={{ background: "none", border: "none", cursor: "pointer", color: "#64748b", padding: "4px", borderRadius: "6px" }}
            aria-label="Close modal"
          >
            <X size={20} />
          </button>
        </div>

        {pendingTools.length > 0 && (
          <div style={{ padding: "10px 22px", background: "#f1f5f9", borderBottom: "1px solid #e2e8f0", display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: "8px" }}>
            <span style={{ fontSize: "11px", color: "#475569", fontWeight: 500 }}>
              Quick action for all {pendingTools.length} pending items:
            </span>
            <div style={{ display: "flex", gap: "8px" }}>
              <button
                type="button"
                disabled={batchLoading}
                onClick={() => handleBatch("approved")}
                style={{ padding: "5px 10px", fontSize: "11px", fontWeight: 600, background: "#0d9488", color: "white", border: "none", borderRadius: "6px", cursor: "pointer" }}
              >
                ✓ I Know All of These
              </button>
              <button
                type="button"
                disabled={batchLoading}
                onClick={() => handleBatch("skipped")}
                style={{ padding: "5px 10px", fontSize: "11px", fontWeight: 600, background: "#e2e8f0", color: "#475569", border: "none", borderRadius: "6px", cursor: "pointer" }}
              >
                ✕ Skip All (No Penalty)
              </button>
            </div>
          </div>
        )}

        <div style={{ padding: "18px 22px", overflowY: "auto", display: "flex", flexDirection: "column", gap: "12px", flex: 1 }}>
          {tools.length === 0 ? (
            <div style={{ textAlign: "center", padding: "30px 20px", color: "#64748b" }}>
              <CheckCircleIcon />
              <p style={{ marginTop: "10px", fontSize: "13px" }}>All technical requirements in the job description are already present in your verified resume!</p>
            </div>
          ) : (
            tools.map((item) => {
              const isPending = item.status === "pending";
              const isConfirmed = item.status === "approved" || item.status === "related";
              const isSkipped = item.status === "skipped";
              const isBusy = loadingTool === item.tool || batchLoading;

              return (
                <div
                  key={item.tool}
                  style={{
                    border: isConfirmed ? "1px solid #a7f3d0" : isSkipped ? "1px solid #e2e8f0" : "1px solid #cbd5e1",
                    background: isConfirmed ? "#f0fdf4" : isSkipped ? "#f8fafc" : "#ffffff",
                    borderRadius: "10px",
                    padding: "14px 16px",
                    display: "flex",
                    flexDirection: "column",
                    gap: "10px",
                  }}
                >
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: "10px" }}>
                    <div>
                      <div style={{ display: "flex", alignItems: "center", gap: "8px", marginBottom: "4px" }}>
                        <strong style={{ fontSize: "14px", color: isSkipped ? "#64748b" : "#0f172a", textTransform: "capitalize" }}>
                          {item.tool}
                        </strong>
                        <span style={{ fontSize: "9.5px", padding: "2px 7px", borderRadius: "4px", background: "#e0f2fe", color: "#0369a1", fontWeight: 600 }}>
                          {item.domain}
                        </span>
                        {isConfirmed && (
                          <span style={{ fontSize: "9.5px", padding: "2px 7px", borderRadius: "4px", background: "#dcfce7", color: "#15803d", fontWeight: 600, display: "flex", alignItems: "center", gap: "3px" }}>
                            <Check size={11} /> Approved
                          </span>
                        )}
                        {isSkipped && (
                          <span style={{ fontSize: "9.5px", padding: "2px 7px", borderRadius: "4px", background: "#f1f5f9", color: "#64748b", fontWeight: 600 }}>
                            Skipped
                          </span>
                        )}
                      </div>
                      <p style={{ margin: 0, fontSize: "11.5px", color: isSkipped ? "#94a3b8" : "#334155", lineHeight: 1.5 }}>
                        {item.explanation}
                      </p>
                    </div>
                  </div>

                  <div style={{ display: "flex", justifyContent: "flex-end", alignItems: "center", gap: "8px", paddingTop: "8px", borderTop: "1px solid #f1f5f9" }}>
                    {isPending ? (
                      <>
                        <button
                          type="button"
                          disabled={isBusy}
                          onClick={() => handleConfirm(item.tool, "approved")}
                          style={{ padding: "6px 12px", fontSize: "11px", fontWeight: 600, background: "#0d9488", color: "white", border: "none", borderRadius: "6px", cursor: "pointer" }}
                        >
                          ✓ I Know This / Add
                        </button>
                        <button
                          type="button"
                          disabled={isBusy}
                          onClick={() => handleConfirm(item.tool, "related")}
                          style={{ padding: "6px 12px", fontSize: "11px", fontWeight: 600, background: "#eff6ff", color: "#1d4ed8", border: "1px solid #bfdbfe", borderRadius: "6px", cursor: "pointer" }}
                        >
                          ~ Related Experience
                        </button>
                        <button
                          type="button"
                          disabled={isBusy}
                          onClick={() => handleConfirm(item.tool, "skipped")}
                          style={{ padding: "6px 12px", fontSize: "11px", background: "#f1f5f9", color: "#475569", border: "1px solid #cbd5e1", borderRadius: "6px", cursor: "pointer" }}
                        >
                          ✕ Do Not Add / Skip
                        </button>
                      </>
                    ) : (
                      <button
                        type="button"
                        disabled={isBusy}
                        onClick={() => handleConfirm(item.tool, isConfirmed ? "skipped" : "approved")}
                        style={{ fontSize: "10.5px", color: "#0d9488", background: "none", border: "none", textDecoration: "underline", cursor: "pointer" }}
                      >
                        Change to {isConfirmed ? "Skip" : "Approved"}
                      </button>
                    )}
                  </div>
                </div>
              );
            })
          )}
        </div>

        <div style={{ padding: "14px 22px", borderTop: "1px solid #e2e8f0", background: "#f8fafc", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <span style={{ fontSize: "11px", color: "#64748b" }}>
            {confirmedTools.length} confirmed · {skippedTools.length} skipped · {pendingTools.length} pending
          </span>
          <button
            type="button"
            onClick={onClose}
            style={{ padding: "8px 18px", fontSize: "12px", fontWeight: 600, background: "#0f172a", color: "white", border: "none", borderRadius: "6px", cursor: "pointer" }}
          >
            Done
          </button>
        </div>
      </div>
    </div>
  );
}

function CandidateProfilePage({
  profile,
  onSave,
  onExtract,
  onTailorWithProfile,
}: {
  profile: CandidateProfile | null;
  onSave: (updated: CandidateProfile) => Promise<void>;
  onExtract: (file?: File) => Promise<void>;
  onTailorWithProfile: () => void;
}) {
  const [formData, setFormData] = useState<CandidateProfile>({
    full_name: profile?.full_name || "",
    headline: profile?.headline || "",
    contact: profile?.contact || "",
    summary: profile?.summary || "",
    skills: profile?.skills || [],
    experience: profile?.experience || [],
    education: profile?.education || [],
    projects: profile?.projects || [],
    certifications: profile?.certifications || [],
  });
  const [skillsText, setSkillsText] = useState((profile?.skills || []).join(", "));
  const [experienceText, setExperienceText] = useState((profile?.experience || []).join("\n\n"));
  const [educationText, setEducationText] = useState((profile?.education || []).join("\n\n"));
  const [projectsText, setProjectsText] = useState((profile?.projects || []).join("\n\n"));
  const [certificationsText, setCertificationsText] = useState((profile?.certifications || []).join(", "));
  const [saving, setSaving] = useState(false);
  const [extracting, setExtracting] = useState(false);
  const [msg, setMsg] = useState("");

  useEffect(() => {
    if (profile) {
      setFormData({
        full_name: profile.full_name || "",
        headline: profile.headline || "",
        contact: profile.contact || "",
        summary: profile.summary || "",
        skills: profile.skills || [],
        experience: profile.experience || [],
        education: profile.education || [],
        projects: profile.projects || [],
        certifications: profile.certifications || [],
      });
      setSkillsText((profile.skills || []).join(", "));
      setExperienceText((profile.experience || []).join("\n\n"));
      setEducationText((profile.education || []).join("\n\n"));
      setProjectsText((profile.projects || []).join("\n\n"));
      setCertificationsText((profile.certifications || []).join(", "));
    }
  }, [profile]);

  async function handleSave(e: FormEvent) {
    e.preventDefault();
    setSaving(true);
    setMsg("");
    try {
      const skills = skillsText.split(/[,;\n]+/).map((s) => s.trim()).filter(Boolean);
      const experience = experienceText.split(/\n{2,}|\r\n{2,}/).map((s) => s.trim()).filter(Boolean);
      const education = educationText.split(/\n{2,}|\r\n{2,}/).map((s) => s.trim()).filter(Boolean);
      const projects = projectsText.split(/\n{2,}|\r\n{2,}/).map((s) => s.trim()).filter(Boolean);
      const certifications = certificationsText.split(/[,;\n]+/).map((s) => s.trim()).filter(Boolean);

      const updated: CandidateProfile = {
        ...formData,
        skills,
        experience,
        education,
        projects,
        certifications,
      };
      await onSave(updated);
      setMsg("Candidate profile saved successfully!");
    } catch (err) {
      setMsg((err as Error).message);
    } finally {
      setSaving(false);
    }
  }

  async function handleExtract(file?: File) {
    setExtracting(true);
    setMsg("");
    try {
      await onExtract(file);
      setMsg("Extracted profile from resume successfully!");
    } catch (err) {
      setMsg((err as Error).message);
    } finally {
      setExtracting(false);
    }
  }

  return (
    <div className="profile-page">
      <div className="collection-header">
        <div>
          <span className="section-kicker">PERSISTENT CANDIDATE PROFILE</span>
          <h1>Candidate Profile</h1>
          <p>Maintain your real, verified background details. Reusable across multiple tailoring sessions.</p>
        </div>
        <div style={{ display: "flex", gap: "8px" }}>
          {formData.full_name && (
            <button className="button button-primary" type="button" onClick={onTailorWithProfile}>
              <Sparkles size={14} /> Tailor with this profile
            </button>
          )}
        </div>
      </div>

      <div className="profile-extract-box">
        <div>
          <strong style={{ fontSize: "11px", color: "#25483f" }}>Populate from existing resume</strong>
          <p style={{ margin: "3px 0 0", fontSize: "9px", color: "#6a7f78" }}>
            Extract verified details from your saved master resume or upload a resume file without inventing details.
          </p>
        </div>
        <div style={{ display: "flex", gap: "8px", alignItems: "center" }}>
          <label className="button button-secondary button-small" style={{ cursor: "pointer" }}>
            <CloudUpload size={14} /> Upload resume to extract
            <input
              type="file"
              accept=".pdf,.docx"
              style={{ display: "none" }}
              onChange={(e) => {
                const file = e.target.files?.[0];
                if (file) void handleExtract(file);
              }}
            />
          </label>
          <button
            className="button button-secondary button-small"
            type="button"
            disabled={extracting}
            onClick={() => void handleExtract()}
          >
            {extracting ? <LoaderCircle className="spin" size={14} /> : <FileText size={14} />}
            Extract from master resume
          </button>
        </div>
      </div>

      {msg && (
        <div className="ai-notice" style={{ marginBottom: "14px" }}>
          <ShieldCheck size={14} /> {msg}
        </div>
      )}

      <form className="card" style={{ padding: "22px 24px" }} onSubmit={handleSave}>
        <div className="ai-notice" style={{ marginBottom: "16px" }}>
          <ShieldCheck size={14} />
          <strong>Ground truth guarantee:</strong> Profiles preserve your authentic experience. Never invent skills, companies, degrees, or achievements.
        </div>

        <div className="profile-grid">
          <div className="profile-field">
            <label>Full Name *</label>
            <input
              value={formData.full_name}
              onChange={(e) => setFormData({ ...formData, full_name: e.target.value })}
              placeholder="e.g. Jane Doe"
              required
            />
          </div>
          <div className="profile-field">
            <label>Professional Headline</label>
            <input
              value={formData.headline}
              onChange={(e) => setFormData({ ...formData, headline: e.target.value })}
              placeholder="e.g. Senior Software Engineer | Distributed Systems"
            />
          </div>
          <div className="profile-field profile-full-width">
            <label>Contact Information (Email, Phone, Location, Links)</label>
            <input
              value={formData.contact}
              onChange={(e) => setFormData({ ...formData, contact: e.target.value })}
              placeholder="e.g. jane@example.com | +1 (555) 019-2834 | San Francisco, CA | github.com/janedoe"
            />
          </div>
          <div className="profile-field profile-full-width">
            <label>Professional Summary</label>
            <textarea
              value={formData.summary}
              onChange={(e) => setFormData({ ...formData, summary: e.target.value })}
              placeholder="Concise overview of your verified experience and technical strengths..."
              rows={3}
            />
          </div>
          <div className="profile-field profile-full-width">
            <label>Core Skills (comma-separated)</label>
            <textarea
              value={skillsText}
              onChange={(e) => setSkillsText(e.target.value)}
              placeholder="Python, FastAPI, TypeScript, React, PostgreSQL, Docker, Kubernetes..."
              rows={2}
            />
          </div>
          <div className="profile-field profile-full-width">
            <label>Experience & Responsibilities (separate roles or paragraphs by blank lines)</label>
            <textarea
              value={experienceText}
              onChange={(e) => setExperienceText(e.target.value)}
              placeholder="Senior Engineer at Acme Corp (2021-Present)&#10;- Designed and deployed real-time data pipelines reducing latency by 40%&#10;- Mentored 5 junior engineers and led code reviews"
              rows={6}
            />
          </div>
          <div className="profile-field">
            <label>Education & Degrees</label>
            <textarea
              value={educationText}
              onChange={(e) => setEducationText(e.target.value)}
              placeholder="B.S. in Computer Science, University of California (2016-2020)"
              rows={3}
            />
          </div>
          <div className="profile-field">
            <label>Key Projects</label>
            <textarea
              value={projectsText}
              onChange={(e) => setProjectsText(e.target.value)}
              placeholder="Distributed Cache: Built high-throughput in-memory cache in Go handling 50k QPS."
              rows={3}
            />
          </div>
          <div className="profile-field profile-full-width">
            <label>Certifications & Licenses (comma-separated)</label>
            <input
              value={certificationsText}
              onChange={(e) => setCertificationsText(e.target.value)}
              placeholder="AWS Certified Solutions Architect, CKA, PMP"
            />
          </div>
        </div>

        <div className="profile-actions">
          <span style={{ fontSize: "9px", color: "#7a8a92" }}>
            Changes are saved to your local database and will persist between browser sessions.
          </span>
          <button className="button button-primary" type="submit" disabled={saving}>
            {saving ? <LoaderCircle className="spin" size={15} /> : <Check size={15} />}
            {saving ? "Saving..." : "Save Candidate Profile"}
          </button>
        </div>
      </form>
    </div>
  );
}

export default App;
