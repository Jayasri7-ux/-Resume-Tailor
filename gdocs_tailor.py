import os
from pathlib import Path

from google import genai
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ["https://www.googleapis.com/auth/documents", "https://www.googleapis.com/auth/drive"]
CREDENTIALS_PATH = Path(__file__).with_name("credentials.json")
TOKEN_PATH = Path(__file__).with_name("token.json")


def get_gdocs_service():
    creds = None
    if TOKEN_PATH.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN_PATH), SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(str(CREDENTIALS_PATH), SCOPES)
            creds = flow.run_local_server(port=0)
        TOKEN_PATH.write_text(creds.to_json(), encoding="utf-8")
    return build("docs", "v1", credentials=creds), build("drive", "v3", credentials=creds)


def find_working_document_id(drive_service, preferred_id):
    try:
        drive_service.files().get(fileId=preferred_id, fields="id").execute()
        return preferred_id
    except HttpError as exc:
        if exc.resp.status != 404:
            raise

    response = drive_service.files().list(
        q="mimeType='application/vnd.google-apps.document'",
        pageSize=5,
        orderBy="modifiedTime desc",
        fields="files(id, name)",
    ).execute()
    files = response.get("files", [])
    if not files:
        raise RuntimeError("No editable Google Docs were found in this Google Drive.")
    print(f"Using the newest Google Doc: '{files[0]['name']}'")
    return files[0]["id"]


def read_google_doc(docs_service, document_id):
    doc = docs_service.documents().get(documentId=document_id).execute()
    text = ""
    for element in doc.get("body", {}).get("content", []):
        if "paragraph" in element:
            for run in element["paragraph"].get("elements", []):
                text += run.get("textRun", {}).get("content", "")
    return text


def duplicate_google_doc(drive_service, original_id, new_title):
    copied_file = drive_service.files().copy(fileId=original_id, body={"name": new_title}).execute()
    return copied_file.get("id")


def get_ai_tailored_text(original_text, job_description):
    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        raise RuntimeError("Set GEMINI_API_KEY in the environment before using Gemini.")

    client = genai.Client(api_key=api_key)
    prompt = f"""
Suggest edits to the candidate's existing technical skills only. Use only skills explicitly
present in the candidate's resume. Never infer or add a skill just because it appears in the
job description. Return only factual lines that can be supported by the resume.

Candidate resume:
{original_text}

Job description:
{job_description}
"""
    response = client.models.generate_content(model="gemini-2.5-flash", contents=prompt)
    if not response.text:
        raise RuntimeError("Gemini returned no suggested changes.")
    return response.text


def update_google_doc(docs_service, doc_id, ai_instructions):
    doc = docs_service.documents().get(documentId=doc_id).execute()
    new_skills = [line.strip() for line in ai_instructions.splitlines() if ":" in line]
    requests = []

    for element in doc.get("body", {}).get("content", []):
        if "paragraph" not in element:
            continue
        paragraph = element["paragraph"]
        paragraph_text = "".join(
            run.get("textRun", {}).get("content", "")
            for run in paragraph.get("elements", [])
        ).strip()
        for new_line in new_skills:
            category = new_line.split(":", 1)[0].strip().casefold()
            if category and paragraph_text.casefold().startswith(f"{category}:".casefold()):
                requests.append({
                    "replaceAllText": {
                        "containsText": {"text": paragraph_text, "matchCase": True},
                        "replaceText": new_line,
                    }
                })

    if requests:
        docs_service.documents().batchUpdate(documentId=doc_id, body={"requests": requests}).execute()
        print("Approved technical-skill edits were saved to the copied Google Doc.")
    else:
        print("No matching technical-skill lines were found; the copied document is unchanged.")
