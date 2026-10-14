from io import BytesIO
import json

import fitz
import pytest
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from fastapi.testclient import TestClient

from backend import main


JOB_DESCRIPTION = (
    "Quality associate performs quality audits, communicates with stakeholders, "
    "uses SQL and Microsoft Excel, and meets productivity benchmarks."
)
client = TestClient(main.app)


@pytest.fixture
def isolated_workspace(tmp_path, monkeypatch):
    data_root = tmp_path / "files"
    data_root.mkdir()
    test_engine = main.create_engine(
        f"sqlite:///{tmp_path / 'workflows.sqlite'}",
        connect_args={"check_same_thread": False},
    )
    main.Base.metadata.create_all(test_engine)
    monkeypatch.setattr(main, "DATA_ROOT", data_root)
    monkeypatch.setattr(main, "engine", test_engine)
    monkeypatch.setattr(main.app.state, "master_resume", None)
    monkeypatch.setattr(main, "GEMINI_API_KEY", "")
    yield
    test_engine.dispose()


def upload_docx():
    doc = Document()
    doc.add_paragraph("Jordan Example", style="Title")
    doc.add_paragraph("Skills", style="Heading 1")
    skills = doc.add_paragraph("SQL, Excel, reporting and dashboard tools")
    skills.alignment = WD_ALIGN_PARAGRAPH.CENTER
    skills.runs[0].bold = True
    doc.add_paragraph("Experience", style="Heading 1")
    doc.add_paragraph("Performed quality audits and communicated with stakeholders.")
    content = BytesIO()
    doc.save(content)
    return content.getvalue()


def upload_pdf(scanned=False):
    doc = fitz.open()
    page = doc.new_page()
    if not scanned:
        page.insert_text((72, 72), "Jordan Example\nSkills\nSQL, Excel\nExperience")
        page.insert_text((72, 112), "Performed quality audits and communicated with stakeholders.")
    data = doc.tobytes()
    doc.close()
    return data


def analyze(filename, content, job_description=JOB_DESCRIPTION):
    response = client.post(
        "/api/analyze",
        files={"file": (filename, content)},
        data={"job_description": job_description, "job_title": "Quality Associate", "company_name": "Example"},
    )
    assert response.status_code == 200, response.text
    return response.json()


def resolve_all(workflow_id):
    response = client.post(
        f"/api/workflows/{workflow_id}/suggestions/bulk",
        json={"action": "approve"},
    )
    assert response.status_code == 200, response.text
    return response.json()


def remove_workflow(workflow_id):
    response = client.delete(f"/api/workflows/{workflow_id}")
    assert response.status_code == 204


def test_master_resume_is_loaded_and_used_without_upload(isolated_workspace):
    master_path = main.DATA_ROOT / "master_resume.docx"
    master_path.write_bytes(upload_docx())
    loaded = main.initialize_master_resume(master_path)
    assert loaded is not None
    assert loaded["paragraphs"]
    with main.Session(main.engine) as db:
        stored = db.get(main.WorkspaceResumeRecord, "default")
        assert stored is not None
        assert "Jordan Example" in stored.payload
    assert client.get("/api/health").json()["default_resume_available"] is True

    response = client.post(
        "/api/analyze",
        data={"job_description": JOB_DESCRIPTION, "job_title": "Quality Associate"},
    )
    assert response.status_code == 200, response.text
    workflow = response.json()
    assert workflow["filename"] == "master_resume.docx"
    original = client.get(f"/api/workflows/{workflow['id']}/original")
    assert original.status_code == 200
    assert "Jordan Example" in "\n".join(
        paragraph.text for paragraph in Document(BytesIO(original.content)).paragraphs
    )
    remove_workflow(workflow["id"])


def test_docx_approval_edit_export_and_delete(isolated_workspace):
    workflow = analyze("resume.docx", upload_docx())
    workflow_id = workflow["id"]
    assert workflow["demo_mode"] is True
    assert workflow["analysis"]["score"] < 100
    assert workflow["suggestions"]
    assert client.post(f"/api/workflows/{workflow_id}/finalize").status_code == 409
    assert client.get(f"/api/workflows/{workflow_id}/download/docx").status_code == 409
    skills = next(item for item in workflow["suggestions"] if item["section"] == "Skills")
    edited = client.patch(
        f"/api/workflows/{workflow_id}/suggestions/{skills['id']}",
        json={"action": "edit", "proposed_text": "SQL, Excel, productivity benchmarks"},
    )
    assert edited.status_code == 200, edited.text
    assert next(item for item in edited.json()["suggestions"] if item["id"] == skills["id"])["status"] == "pending"

    resolve_all(workflow_id)
    approve = client.patch(
        f"/api/workflows/{workflow_id}/suggestions/{skills['id']}",
        json={"action": "approve"},
    )
    assert approve.status_code == 200, approve.text
    finalized = client.post(f"/api/workflows/{workflow_id}/finalize")
    assert finalized.status_code == 200, finalized.text
    result = finalized.json()
    assert result["finalized"] is True
    assert result["final_score"] > result["analysis"]["score"]

    original = client.get(result["download_urls"]["original"])
    tailored_docx = client.get(result["download_urls"]["docx"])
    tailored_pdf = client.get(result["download_urls"]["pdf"])
    assert original.status_code == tailored_docx.status_code == tailored_pdf.status_code == 200
    source_doc = Document(BytesIO(original.content))
    tailored_doc = Document(BytesIO(tailored_docx.content))
    source_skills = next(p for p in source_doc.paragraphs if p.text.startswith("SQL, Excel,"))
    tailored_skills = next(p for p in tailored_doc.paragraphs if "productivity benchmarks" in p.text)
    assert tailored_skills.alignment == source_skills.alignment
    assert tailored_skills.runs[0].bold is True
    assert result["layout_review"]
    pdf = fitz.open(stream=tailored_pdf.content, filetype="pdf")
    assert "productivity benchmarks" in pdf[0].get_text()
    pdf.close()

    remove_workflow(workflow_id)
    assert client.get(f"/api/workflows/{workflow_id}").status_code == 404


def test_post_export_reconciles_suggestion_state_and_returns_sanitized_attachments(isolated_workspace):
    workflow = analyze("resume.docx", upload_docx())
    workflow_id = workflow["id"]
    pending_export = client.post(
        f"/api/workflows/{workflow_id}/download/docx",
        json={
            "suggestions": [
                {
                    "id": suggestion["id"],
                    "status": suggestion["status"],
                    "proposed_text": suggestion["proposed_text"],
                }
                for suggestion in workflow["suggestions"]
            ]
        },
    )
    assert pending_export.status_code == 409
    skills = next(item for item in workflow["suggestions"] if item["section"] == "Skills")
    edited_text = "• SQL, Excel, productivity benchmarks"
    updated = client.patch(
        f"/api/workflows/{workflow_id}/suggestions/{skills['id']}",
        json={"action": "approve", "proposed_text": edited_text},
    )
    assert updated.status_code == 200, updated.text

    for suggestion in updated.json()["suggestions"]:
        if suggestion["status"] == "pending":
            rejected = client.patch(
                f"/api/workflows/{workflow_id}/suggestions/{suggestion['id']}",
                json={"action": "reject"},
            )
            assert rejected.status_code == 200, rejected.text

    current = client.get(f"/api/workflows/{workflow_id}").json()
    export_state = {
        "suggestions": [
            {
                "id": suggestion["id"],
                "status": suggestion["status"],
                "proposed_text": suggestion["proposed_text"],
            }
            for suggestion in current["suggestions"]
        ]
    }
    docx_response = client.post(
        f"/api/workflows/{workflow_id}/download/docx",
        json=export_state,
    )
    assert docx_response.status_code == 200, docx_response.text
    assert "attachment" in docx_response.headers["content-disposition"]
    exported_docx = Document(BytesIO(docx_response.content))
    assert "- SQL, Excel, productivity benchmarks" in "\n".join(
        paragraph.text for paragraph in exported_docx.paragraphs
    )

    pdf_response = client.post(
        f"/api/workflows/{workflow_id}/download/pdf",
        json=export_state,
    )
    assert pdf_response.status_code == 200, pdf_response.text
    assert "attachment" in pdf_response.headers["content-disposition"]
    pdf = fitz.open(stream=pdf_response.content, filetype="pdf")
    assert "SQL, Excel, productivity benchmarks" in "\n".join(
        page.get_text() for page in pdf
    )
    pdf.close()
    remove_workflow(workflow_id)


def test_direct_download_route_builds_from_preview_text_in_memory(isolated_workspace):
    workflow = analyze("resume.docx", upload_docx())
    workflow_id = workflow["id"]
    blocked = client.post(
        "/api/download/pdf",
        json={"workflow_id": workflow_id, "text": workflow["final_text"]},
    )
    assert blocked.status_code == 409
    for suggestion in workflow["suggestions"]:
        response = client.patch(
            f"/api/workflows/{workflow_id}/suggestions/{suggestion['id']}",
            json={"action": "reject"},
        )
        assert response.status_code == 200, response.text

    current = client.get(f"/api/workflows/{workflow_id}").json()
    preview_text = current["final_text"]
    first_line, remaining_text = preview_text.split("\n", 1)
    submitted_text = f"**{first_line}**\n{remaining_text}"

    docx_response = client.post(
        "/api/download/docx",
        json={"workflow_id": workflow_id, "text": submitted_text},
    )
    assert docx_response.status_code == 200, docx_response.text
    assert "attachment" in docx_response.headers["content-disposition"]
    docx = Document(BytesIO(docx_response.content))
    docx_lines = [paragraph.text for paragraph in docx.paragraphs]
    assert docx_lines[0] == first_line
    assert "Skills" in docx_lines
    assert "Experience" in docx_lines

    pdf_response = client.post(
        "/api/download/pdf",
        json={"workflow_id": workflow_id, "text": submitted_text},
    )
    assert pdf_response.status_code == 200, pdf_response.text
    assert "attachment" in pdf_response.headers["content-disposition"]
    pdf = fitz.open(stream=pdf_response.content, filetype="pdf")
    pdf_text = "\n".join(page.get_text() for page in pdf)
    assert first_line in pdf_text
    assert "Skills" in pdf_text
    pdf.close()

    workflow_folder = main.DATA_ROOT / workflow_id
    assert not (workflow_folder / "tailored.docx").exists()
    assert not (workflow_folder / "tailored.pdf").exists()
    remove_workflow(workflow_id)


def test_docx_replacement_preserves_runs_and_embeds_styled_hyperlinks():
    doc = Document()
    paragraph = doc.add_paragraph()
    paragraph.add_run("Reach ").bold = True
    paragraph.add_run("https://linkedin.com/in/jaya").italic = True

    main.replace_docx_paragraph(
        paragraph,
        "Reach https://linkedin.com/in/jaya",
    )

    hyperlinks = paragraph._p.xpath("./w:hyperlink")
    assert len(hyperlinks) == 1
    hyperlink = hyperlinks[0]
    hyperlink_run = hyperlink.xpath(".//w:r")[0]
    hyperlink_properties = hyperlink_run.find(main.qn("w:rPr"))
    assert hyperlink_properties.find(main.qn("w:i")) is not None
    assert hyperlink_properties.find(main.qn("w:color")).get(main.qn("w:val")) == "0563C1"
    assert hyperlink_properties.find(main.qn("w:u")).get(main.qn("w:val")) == "single"
    relationship_id = hyperlink.get(main.qn("r:id"))
    assert paragraph.part.rels[relationship_id].target_ref == "https://linkedin.com/in/jaya"
    assert paragraph._p.xpath("./w:r")[0].find(main.qn("w:rPr")).find(main.qn("w:b")) is not None


def test_docx_export_rejects_replacements_over_the_existing_text_budget(isolated_workspace):
    workflow = analyze("resume.docx", upload_docx())
    workflow_id = workflow["id"]
    skills = next(item for item in workflow["suggestions"] if item["section"] == "Skills")
    edit = client.patch(
        f"/api/workflows/{workflow_id}/suggestions/{skills['id']}",
        json={"action": "edit", "proposed_text": "SQL, Excel, this replacement is much too long"},
    )
    assert edit.status_code == 200, edit.text
    approve = client.post(f"/api/workflows/{workflow_id}/suggestions/bulk", json={"action": "approve"})
    assert approve.status_code == 200, approve.text
    finalized = client.post(f"/api/workflows/{workflow_id}/finalize")
    assert finalized.status_code == 422
    assert "layout budget" in finalized.json()["detail"]
    remove_workflow(workflow_id)


def test_docx_export_refuses_a_multi_page_result(isolated_workspace, monkeypatch):
    workflow = analyze("resume.docx", upload_docx())
    workflow_id = workflow["id"]
    resolve_all(workflow_id)
    monkeypatch.setattr(
        main,
        "inspect_pdf_layout",
        lambda _path: {"page_count": 2, "blank_pages": [], "overflow_blocks": 0, "overlapping_blocks": 0},
    )

    finalized = client.post(f"/api/workflows/{workflow_id}/finalize")

    assert finalized.status_code == 422
    assert "one page" in finalized.json()["detail"]
    assert not (main.DATA_ROOT / workflow_id / "tailored.docx").exists()
    remove_workflow(workflow_id)


def test_pdf_analysis_export_and_scanned_pdf_error(isolated_workspace):
    workflow = analyze("resume.pdf", upload_pdf())
    workflow_id = workflow["id"]
    assert workflow["extension"] == ".pdf"
    assert workflow["analysis"]["score"] > 0
    skills = next(item for item in workflow["suggestions"] if item["section"] == "Skills")
    edited = client.patch(
        f"/api/workflows/{workflow_id}/suggestions/{skills['id']}",
        json={"action": "edit", "proposed_text": "SQL, Excel, Microsoft productivity"},
    )
    assert edited.status_code == 200, edited.text
    resolve_all(workflow_id)
    approved = client.patch(
        f"/api/workflows/{workflow_id}/suggestions/{skills['id']}",
        json={"action": "approve"},
    )
    assert approved.status_code == 200, approved.text
    finalized = client.post(f"/api/workflows/{workflow_id}/finalize")
    assert finalized.status_code == 200, finalized.text
    downloaded_docx = client.get(finalized.json()["download_urls"]["docx"])
    downloaded_pdf = client.get(finalized.json()["download_urls"]["pdf"])
    assert "Microsoft productivity" in "\n".join(
        paragraph.text for paragraph in Document(BytesIO(downloaded_docx.content)).paragraphs
    )
    pdf = fitz.open(stream=downloaded_pdf.content, filetype="pdf")
    assert "Microsoft productivity" in "\n".join(page.get_text() for page in pdf)
    pdf.close()
    remove_workflow(workflow_id)

    scanned = client.post(
        "/api/analyze",
        files={"file": ("scan.pdf", upload_pdf(scanned=True))},
        data={"job_description": JOB_DESCRIPTION},
    )
    assert scanned.status_code == 422
    assert "scanned" in scanned.json()["detail"].lower()


def test_rejected_changes_keep_original_and_instructions_return_to_review(isolated_workspace):
    workflow = analyze("resume.docx", upload_docx())
    workflow_id = workflow["id"]
    suggestion = workflow["suggestions"][0]
    response = client.patch(
        f"/api/workflows/{workflow_id}/suggestions/{suggestion['id']}",
        json={"action": "reject"},
    )
    assert response.status_code == 200, response.text
    undone = client.patch(
        f"/api/workflows/{workflow_id}/suggestions/{suggestion['id']}",
        json={"action": "undo"},
    )
    assert undone.status_code == 200, undone.text
    assert next(item for item in undone.json()["suggestions"] if item["id"] == suggestion["id"])["status"] == "pending"

    updated = client.post(
        f"/api/workflows/{workflow_id}/instructions",
        json={"instruction": "Please shorten my summary without changing any facts."},
    )
    assert updated.status_code == 200, updated.text
    assert updated.json()["demo_mode"] is True
    assert all(item["status"] == "pending" for item in updated.json()["suggestions"])
    suggestion = updated.json()["suggestions"][0]
    resolve_all(workflow_id)
    reject = client.patch(
        f"/api/workflows/{workflow_id}/suggestions/{suggestion['id']}",
        json={"action": "reject"},
    )
    assert reject.status_code == 200, reject.text
    finalized = client.post(f"/api/workflows/{workflow_id}/finalize")
    assert finalized.status_code == 200, finalized.text
    assert finalized.json()["final_text"] == finalized.json()["original_text"]
    assert finalized.json()["final_score"] == finalized.json()["analysis"]["score"]
    remove_workflow(workflow_id)


def test_unavailable_gemini_is_explicitly_labeled_demo(isolated_workspace, monkeypatch):
    monkeypatch.setattr(main, "GEMINI_API_KEY", "test-key")

    def unavailable(*_args):
        raise main.HTTPException(status_code=502, detail="Gemini is unavailable.")

    monkeypatch.setattr(main, "call_gemini_suggestions", unavailable)
    workflow = analyze("resume.docx", upload_docx())
    assert workflow["demo_mode"] is True
    assert workflow["suggestion_warnings"]
    assert "Gemini is unavailable" in workflow["suggestion_warnings"][0]
    assert all(suggestion["source"] == "demo" for suggestion in workflow["suggestions"])
    remove_workflow(workflow["id"])


def test_zero_overlap_still_offers_source_text_for_manual_editing(isolated_workspace):
    suggestions = main.make_demo_suggestions(
        ["Experience", "Coordinated vendor schedules, maintained records, and resolved delivery issues."],
        "Analytical capacity, critical thinking, attention to detail, and Excel tracking.",
    )

    assert len(suggestions) == 1
    suggestion = suggestions[0]
    assert suggestion["section"] == "Experience"
    assert suggestion["original_text"] == suggestion["proposed_text"]
    assert suggestion["requirement"] == ""
    assert suggestion["source"] == "demo"
    assert "manual editing" in suggestion["reason"]


def test_keyword_extraction_keeps_short_and_compound_job_terms():
    details = main.score_resume(
        "C# development and Node.js services",
        "AI, UX, QA, C#, C++, R, Go, CI/CD, .NET, Node.js, cloud-native architecture.",
    )

    assert {"ai", "ux", "qa", "c#", "c++", "r", "go", "ci/cd", ".net", "node.js"} <= set(
        details["matched_keywords"] + details["missing_keywords"]
    )
    assert {"c#", "node.js"} <= set(details["matched_keywords"])
    assert {"ai", "ux", "qa", "c++", "r", "go", "ci/cd", ".net"} <= set(
        details["missing_keywords"]
    )


def test_score_counts_curated_related_skill_equivalents_without_forcing_floor():
    details = main.score_resume(
        "Analysis capability, problem-solving, precision",
        "Analytical capacity, critical thinking, attention to detail.",
    )

    assert details["score"] == 100
    assert details["exact_matched_count"] == 0
    assert details["semantic_matched_count"] == 6
    assert details["missing_keywords"] == []
    assert details["semantic_matches"]["critical"] == ["problem-solving"]
    assert "not machine-understood meaning" in details["method"]

    unrelated = main.score_resume(
        "Managed schedules and maintained records",
        "Analytical capacity, critical thinking, attention to detail.",
    )
    assert unrelated["score"] == 0
    assert len(unrelated["missing_keywords"]) == unrelated["requirement_count"]


def test_empty_gemini_result_keeps_manual_editing_available_for_instructions(
    isolated_workspace, monkeypatch
):
    monkeypatch.setattr(main, "GEMINI_API_KEY", "test-key")
    monkeypatch.setattr(main, "call_gemini_suggestions", lambda *_args: ([], []))
    job_description = (
        "Analytical capacity, critical thinking, attention to detail, and Excel tracking."
    )
    content = upload_docx()
    document = Document(BytesIO(content))
    for paragraph in document.paragraphs:
        if paragraph.text == "SQL, Excel, reporting and dashboard tools":
            paragraph.text = "Coordinated vendor schedules, maintained records, and resolved delivery issues."
    rewritten = BytesIO()
    document.save(rewritten)

    workflow = analyze("resume.docx", rewritten.getvalue(), job_description)
    assert workflow["demo_mode"] is True
    assert len(workflow["suggestions"]) == 1
    assert workflow["suggestions"][0]["source"] == "demo"
    assert workflow["suggestion_warnings"]
    workflow_id = workflow["id"]

    updated = client.post(
        f"/api/workflows/{workflow_id}/instructions",
        json={"instruction": "Emphasize my verified experience without adding facts."},
    )
    assert updated.status_code == 200, updated.text
    assert updated.json()["demo_mode"] is True
    assert len(updated.json()["suggestions"]) == 1
    assert updated.json()["suggestions"][0]["status"] == "pending"
    assert updated.json()["suggestions"][0]["source"] == "demo"
    remove_workflow(workflow_id)


def test_applying_instructions_replaces_stale_queue_items(isolated_workspace, monkeypatch):
    workflow = analyze("resume.docx", upload_docx())
    workflow_id = workflow["id"]
    paragraphs = workflow["original_text"].splitlines()
    skills_index = paragraphs.index("SQL, Excel, reporting and dashboard tools")
    monkeypatch.setattr(main, "GEMINI_API_KEY", "test-key")
    monkeypatch.setattr(
        main,
        "call_gemini_suggestions",
        lambda *_args: ([
            {
                "id": main.suggestion_id("Skills"),
                "section": "Skills",
                "paragraph_index": skills_index,
                "original_text": paragraphs[skills_index],
                "suggested_text": "SQL and Excel reporting tools",
                "proposed_text": "SQL and Excel reporting tools",
                "reason": "A concise edit for review.",
                "requirement": "SQL, Excel",
                "expected_impact": "Keeps existing tools visible.",
                "status": "pending",
                "source": "gemini",
            }
        ], []),
    )

    updated = client.post(
        f"/api/workflows/{workflow_id}/instructions",
        json={"instruction": "Make my verified skills concise."},
    )

    assert updated.status_code == 200, updated.text
    assert updated.json()["demo_mode"] is False
    assert len(updated.json()["suggestions"]) == 1
    assert updated.json()["suggestions"][0]["section"] == "Skills"
    assert updated.json()["suggestions"][0]["source"] == "gemini"
    remove_workflow(workflow_id)


def test_empty_gemini_result_refreshes_to_one_manual_card(isolated_workspace, monkeypatch):
    workflow = analyze("resume.docx", upload_docx())
    workflow_id = workflow["id"]
    monkeypatch.setattr(main, "GEMINI_API_KEY", "test-key")
    monkeypatch.setattr(main, "call_gemini_suggestions", lambda *_args: ([], []))

    updated = client.post(
        f"/api/workflows/{workflow_id}/instructions",
        json={"instruction": "Emphasize my verified skills."},
    )

    assert updated.status_code == 200, updated.text
    result = updated.json()
    assert result["demo_mode"] is True
    assert len(result["suggestions"]) == 1
    assert result["suggestions"][0]["source"] == "demo"
    assert result["suggestions"][0]["original_text"] == result["suggestions"][0]["proposed_text"]
    assert result["suggestion_warnings"]
    remove_workflow(workflow_id)


def test_zero_overlap_manual_edit_updates_export_and_final_score(isolated_workspace):
    job_description = "analytical capacity economics macros mathematics"
    content = upload_docx()
    document = Document(BytesIO(content))
    for paragraph in document.paragraphs:
        if paragraph.text == "SQL, Excel, reporting and dashboard tools":
            paragraph.text = (
                "Reviewed weekly operational reports, coordinated detailed vendor schedules, "
                "maintained customer records, and documented team resolutions."
            )
        elif paragraph.text == "Performed quality audits and communicated with stakeholders.":
            paragraph.text = "Managed daily records, resolved service issues, and coordinated team schedules."
    rewritten = BytesIO()
    document.save(rewritten)

    workflow = analyze("resume.docx", rewritten.getvalue(), job_description)
    workflow_id = workflow["id"]
    assert workflow["analysis"]["score"] == 0
    assert workflow["suggestions"]
    suggestion = workflow["suggestions"][0]
    assert suggestion["proposed_text"] == suggestion["original_text"]

    proposed_text = (
        "Reviewed weekly economics reports, coordinated detailed vendor schedules, "
        "maintained customer records, and documented team resolutions."
    )
    assert len(proposed_text) <= len(suggestion["original_text"])
    edited = client.patch(
        f"/api/workflows/{workflow_id}/suggestions/{suggestion['id']}",
        json={"action": "edit", "proposed_text": proposed_text},
    )
    assert edited.status_code == 200, edited.text
    approved = client.patch(
        f"/api/workflows/{workflow_id}/suggestions/{suggestion['id']}",
        json={"action": "approve"},
    )
    assert approved.status_code == 200, approved.text
    for other in approved.json()["suggestions"]:
        if other["status"] == "pending":
            response = client.patch(
                f"/api/workflows/{workflow_id}/suggestions/{other['id']}",
                json={"action": "reject"},
            )
            assert response.status_code == 200, response.text

    finalized = client.post(f"/api/workflows/{workflow_id}/finalize")
    assert finalized.status_code == 200, finalized.text
    result = finalized.json()
    assert result["final_score"] > result["analysis"]["score"]
    assert "economics reports" in result["final_text"]
    assert "economics" in result["final_score_details"]["matched_keywords"]
    assert result["analysis"]["missing_keywords"] == [
        "analytical", "capacity", "economics", "macros", "mathematics"
    ]
    assert "economics" not in result["final_score_details"]["missing_keywords"]
    remove_workflow(workflow_id)


def test_gemini_404_identifies_the_configured_model(monkeypatch):
    def missing_model(*_args, **_kwargs):
        raise main.HTTPError("https://example.invalid", 404, "Not found", {}, None)

    monkeypatch.setattr(main, "urlopen", missing_model)
    monkeypatch.setattr(main, "GEMINI_MODEL", "missing-model")

    with pytest.raises(main.HTTPException) as exc_info:
        main.call_gemini_suggestions("resume text", "job description")

    assert exc_info.value.status_code == 502
    assert "missing-model" in exc_info.value.detail
    assert "GEMINI_MODEL" in exc_info.value.detail


def test_upload_validation(isolated_workspace):
    unsupported = client.post(
        "/api/analyze",
        files={"file": ("resume.txt", b"text")},
        data={"job_description": JOB_DESCRIPTION},
    )
    assert unsupported.status_code == 415

    invalid_docx = client.post(
        "/api/analyze",
        files={"file": ("resume.docx", b"PK invalid")},
        data={"job_description": JOB_DESCRIPTION},
    )
    assert invalid_docx.status_code == 422


def test_approved_section_replacement_and_live_dynamic_score(isolated_workspace):
    workflow = analyze("resume.docx", upload_docx())
    workflow_id = workflow["id"]
    baseline_score = workflow["analysis"]["score"]
    assert baseline_score < 100

    skills = next(item for item in workflow["suggestions"] if item["section"] == "Skills")
    # Commit approved edit with missing JD keywords ("productivity", "benchmarks").
    tailored_text_block = "SQL, Excel, productivity benchmarks."
    response = client.patch(
        f"/api/workflows/{workflow_id}/suggestions/{skills['id']}",
        json={"action": "approve", "proposed_text": tailored_text_block},
    )
    assert response.status_code == 200, response.text
    data = response.json()

    # 1. Check approved sections replacement in payload
    assert tailored_text_block in data["final_text"]
    assert tailored_text_block in data["tailored_text"]
    assert data["final_text"] != data["original_text"]

    # 2. Check dynamic live score calculation progressing past baseline with positive variance
    assert data["final_score"] is not None
    assert data["tailored_score"] is not None
    assert data["final_score"] > baseline_score
    point_variance = data["final_score"] - baseline_score
    assert point_variance > 0
    assert "productivity" in data["final_score_details"]["matched_keywords"]
    assert "benchmarks" in data["final_score_details"]["matched_keywords"]

    persisted = client.get(f"/api/workflows/{workflow_id}")
    assert persisted.status_code == 200, persisted.text
    assert tailored_text_block in persisted.json()["final_text"]
    assert persisted.json()["final_score"] == data["final_score"]
    remove_workflow(workflow_id)


def test_generative_rewrite_inverse_reconciliation_zero_overlap(monkeypatch):
    resume_text = (
        "Taylor Example\n"
        "Summary\n"
        "Customer service representative supporting inbound caller inquiries.\n"
        "Experience\n"
        "Answered telephone inquiries and filed customer support tickets."
    )
    job_description = (
        "Senior Cloud Engineer who builds Python microservices, manages Kubernetes clusters, "
        "and optimizes PostgreSQL database queries for high availability."
    )
    fake_gemini_payload = {
        "candidates": [{
            "content": {
                "parts": [{
                    "text": json.dumps({
                        "suggestions": [{
                            "section": "Experience",
                            "original_text": "Answered telephone inquiries and filed customer support tickets.",
                            "proposed_text": "Engineered Python services and managed Kubernetes clusters to automate support ticket workflows.",
                            "reason": "Naturally integrates Python and Kubernetes cloud engineering skills into verified workflow automation.",
                            "requirement": "Python, Kubernetes",
                            "expected_impact": "Directly targets core JD requirements without inventing false degrees.",
                        }]
                    })
                }]
            }
        }]
    }

    class MockResponse:
        def read(self):
            return json.dumps(fake_gemini_payload).encode("utf-8")
        def __enter__(self):
            return self
        def __exit__(self, *args):
            pass

    monkeypatch.setattr(main, "urlopen", lambda *args, **kwargs: MockResponse())
    monkeypatch.setattr(main, "GEMINI_API_KEY", "test-api-key")

    validated, warnings = main.call_gemini_suggestions(resume_text, job_description)
    assert len(validated) == 1
    suggestion = validated[0]
    assert suggestion["source"] == "gemini"
    assert suggestion["section"] == "Experience"
    assert "Python" in suggestion["proposed_text"]
    assert "Kubernetes" in suggestion["proposed_text"]
    assert suggestion["proposed_text"] != suggestion["original_text"]


def test_gemini_prompt_receives_every_missing_keyword(monkeypatch):
    keywords = [f"specialty{i}" for i in range(31)]
    captured_prompt = {}
    response_payload = {
        "candidates": [{
            "content": {
                "parts": [{"text": json.dumps({"suggestions": []})}]
            }
        }]
    }

    class MockResponse:
        def read(self):
            return json.dumps(response_payload).encode("utf-8")
        def __enter__(self):
            return self
        def __exit__(self, *_args):
            pass

    def capture_request(request, **_kwargs):
        body = json.loads(request.data)
        captured_prompt["text"] = body["contents"][0]["parts"][0]["text"]
        return MockResponse()

    monkeypatch.setattr(main, "urlopen", capture_request)
    monkeypatch.setattr(main, "GEMINI_API_KEY", "test-api-key")
    main.call_gemini_suggestions("entry level resume", " ".join(keywords))

    keyword_context = captured_prompt["text"].split("JOB DESCRIPTION:", 1)[0]
    assert all(keyword in keyword_context for keyword in keywords)


def test_end_to_end_data_stream_convergence_docx_and_pdf(isolated_workspace):
    workflow = analyze("resume.docx", upload_docx())
    workflow_id = workflow["id"]
    assert workflow["download_urls"] is None
    skills = next(item for item in workflow["suggestions"] if item["section"] == "Skills")
    tailored_text_block = "SQL, Excel, productivity benchmarks"

    # Commit approved replacement
    client.patch(
        f"/api/workflows/{workflow_id}/suggestions/{skills['id']}",
        json={"action": "approve", "proposed_text": tailored_text_block},
    )
    # Reject other pending to satisfy review requirement
    for other in workflow["suggestions"]:
        if other["id"] != skills["id"]:
            client.patch(
                f"/api/workflows/{workflow_id}/suggestions/{other['id']}",
                json={"action": "reject"},
            )

    finalized = client.post(f"/api/workflows/{workflow_id}/finalize")
    assert finalized.status_code == 200, finalized.text
    result = finalized.json()
    assert result["finalized"] is True
    assert result["download_urls"]["docx"].endswith("/download/docx")
    assert result["download_urls"]["pdf"].endswith("/download/pdf")

    # Check python-docx writer output
    tailored_docx_resp = client.get(result["download_urls"]["docx"])
    assert tailored_docx_resp.status_code == 200
    doc = Document(BytesIO(tailored_docx_resp.content))
    docx_text = "\n".join(p.text for p in doc.paragraphs)
    assert tailored_text_block in docx_text

    # Check ReportLab PDF builder output
    tailored_pdf_resp = client.get(result["download_urls"]["pdf"])
    assert tailored_pdf_resp.status_code == 200
    pdf = fitz.open(stream=tailored_pdf_resp.content, filetype="pdf")
    pdf_text = "\n".join(page.get_text() for page in pdf)
    pdf.close()
    assert "productivity benchmarks" in pdf_text
    remove_workflow(workflow_id)
