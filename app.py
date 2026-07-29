from __future__ import annotations

import json
import os
import secrets
import shutil
import threading
import uuid
from pathlib import Path
from typing import Any

import gspread
from dotenv import load_dotenv
from flask import (
    Flask,
    flash,
    redirect,
    render_template,
    request,
    session,
    url_for,
)
from google.auth.transport.requests import Request as GoogleRequest
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import Flow
from werkzeug.utils import secure_filename

from classifier_core import (
    DEFAULT_GOOGLE_SHEET_TITLE,
    create_google_sheet,
    merge_results_into_history,
    process_current_rows,
    read_history_csv,
    read_xlsx_rows,
    save_history_csv,
    save_json_backup,
)

load_dotenv()
os.environ["OAUTHLIB_INSECURE_TRANSPORT"] = os.getenv(
    "OAUTHLIB_INSECURE_TRANSPORT",
    "1",
)

os.environ["OAUTHLIB_RELAX_TOKEN_SCOPE"] = os.getenv(
    "OAUTHLIB_RELAX_TOKEN_SCOPE",
    "1",
)

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
UPLOAD_DIR = BASE_DIR / "uploads"

HISTORY_CSV = DATA_DIR / "image_results_history.csv"
JSON_BACKUP = DATA_DIR / "image_results_history_backup.json"
TOKEN_FILE = DATA_DIR / "google_token.json"
CLIENT_SECRET_FILE = BASE_DIR / "client_secret.json"

DATA_DIR.mkdir(parents=True, exist_ok=True)
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)

ALLOWED_GOOGLE_EMAIL = os.getenv("ALLOWED_GOOGLE_EMAIL", "").strip().lower()
MAX_UPLOAD_MB = int(os.getenv("MAX_UPLOAD_MB", "25"))
GOOGLE_REDIRECT_URI = os.getenv(
    "GOOGLE_REDIRECT_URI",
    "http://127.0.0.1:5000/oauth2callback",
)

SCOPES = [
    "openid",
    "https://www.googleapis.com/auth/userinfo.email",
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/spreadsheets",
]

app = Flask(__name__)
app.secret_key = os.getenv("FLASK_SECRET_KEY") or secrets.token_hex(32)
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_MB * 1024 * 1024

# One process at a time, so two clicks cannot corrupt the history CSV.
processing_lock = threading.Lock()


def is_xlsx(filename: str) -> bool:
    return filename.lower().endswith(".xlsx")


def load_saved_credentials() -> Credentials | None:
    if not TOKEN_FILE.exists():
        return None

    try:
        info = json.loads(TOKEN_FILE.read_text(encoding="utf-8"))
        credentials = Credentials.from_authorized_user_info(info, SCOPES)

        if credentials.expired and credentials.refresh_token:
            credentials.refresh(GoogleRequest())
            save_credentials(credentials)

        if not credentials.valid:
            return None

        return credentials

    except Exception:
        return None


def save_credentials(credentials: Credentials) -> None:
    TOKEN_FILE.write_text(credentials.to_json(), encoding="utf-8")


def google_client() -> gspread.Client:
    credentials = load_saved_credentials()

    if credentials is None:
        raise RuntimeError("Google account is not connected")

    return gspread.authorize(credentials)


def require_google_connection():
    if load_saved_credentials() is None:
        flash("יש להתחבר לחשבון Google לפני הפעלת הבדיקה.", "warning")
        return redirect(url_for("google_login"))

    return None


def fetch_connected_email(credentials: Credentials) -> str:
    import requests

    response = requests.get(
        "https://www.googleapis.com/oauth2/v2/userinfo",
        headers={"Authorization": f"Bearer {credentials.token}"},
        timeout=20,
    )
    response.raise_for_status()
    return str(response.json().get("email", "")).strip().lower()


@app.get("/")
def index():
    connected = load_saved_credentials() is not None
    return render_template(
        "index.html",
        connected=connected,
        allowed_email=ALLOWED_GOOGLE_EMAIL,
        history_exists=HISTORY_CSV.exists(),
    )


@app.get("/login/google")
def google_login():
    if not CLIENT_SECRET_FILE.exists():
        flash(
            "הקובץ client_secret.json חסר בתיקיית הפרויקט.",
            "danger",
        )
        return redirect(url_for("index"))

    # Create one PKCE verifier for this login attempt.
    # It must be reused in the callback.
    code_verifier = secrets.token_urlsafe(64)

    flow = Flow.from_client_secrets_file(
        str(CLIENT_SECRET_FILE),
        scopes=SCOPES,
        redirect_uri=GOOGLE_REDIRECT_URI,
        code_verifier=code_verifier,
    )

    authorization_url, state = flow.authorization_url(
        access_type="offline",
        include_granted_scopes="true",
        prompt="consent",
    )

    # Save both values in the Flask session so that the callback
    # can recreate the same OAuth flow.
    session["oauth_state"] = state
    session["oauth_code_verifier"] = code_verifier

    return redirect(authorization_url)

@app.get("/oauth2callback")
def oauth2callback():
    expected_state = session.get("oauth_state")
    code_verifier = session.get("oauth_code_verifier")

    if not expected_state:
        flash(
            "אימות Google נכשל: חסר OAuth state.",
            "danger",
        )
        return redirect(url_for("index"))

    if request.args.get("state") != expected_state:
        flash(
            "אימות Google נכשל: state לא תקין.",
            "danger",
        )
        return redirect(url_for("index"))

    if not code_verifier:
        flash(
            "אימות Google נכשל: חסר PKCE code verifier. "
            "יש להתחיל מחדש את ההתחברות.",
            "danger",
        )
        return redirect(url_for("index"))

    flow = Flow.from_client_secrets_file(
        str(CLIENT_SECRET_FILE),
        scopes=SCOPES,
        state=expected_state,
        redirect_uri=GOOGLE_REDIRECT_URI,
        code_verifier=code_verifier,
    )

    flow.fetch_token(
        authorization_response=request.url
    )

    credentials = flow.credentials
    email = fetch_connected_email(credentials)

    if (
        ALLOWED_GOOGLE_EMAIL
        and email != ALLOWED_GOOGLE_EMAIL
    ):
        session.pop("oauth_state", None)
        session.pop("oauth_code_verifier", None)

        flash(
            "חשבון Google זה אינו מורשה להשתמש במערכת.",
            "danger",
        )
        return redirect(url_for("index"))

    save_credentials(credentials)

    # These temporary OAuth values are no longer needed.
    session.pop("oauth_state", None)
    session.pop("oauth_code_verifier", None)

    flash(
        f"חשבון Google חובר בהצלחה: {email}",
        "success",
    )

    return redirect(url_for("index"))

@app.post("/run")
def run_classification():
    redirect_response = require_google_connection()
    if redirect_response is not None:
        return redirect_response

    uploaded_file = request.files.get("xlsx_file")

    if uploaded_file is None or not uploaded_file.filename:
        flash("לא נבחר קובץ XLSX.", "danger")
        return redirect(url_for("index"))

    if not is_xlsx(uploaded_file.filename):
        flash("ניתן להעלות קובץ XLSX בלבד.", "danger")
        return redirect(url_for("index"))

    safe_name = secure_filename(uploaded_file.filename) or "input.xlsx"
    run_id = uuid.uuid4().hex
    upload_path = UPLOAD_DIR / f"{run_id}_{safe_name}"
    uploaded_file.save(upload_path)

    try:
        with processing_lock:
            rows = read_xlsx_rows(upload_path)

            if not rows:
                raise ValueError("לא נמצאו כתובות תקינות בקובץ")

            current_results = process_current_rows(rows)
            existing_headers, existing_rows = read_history_csv(HISTORY_CSV)

            (
                merged_headers,
                merged_rows,
                current_run_column,
            ) = merge_results_into_history(
                existing_headers=existing_headers,
                existing_rows=existing_rows,
                current_results=current_results,
            )

            save_history_csv(
                HISTORY_CSV,
                merged_headers,
                merged_rows,
            )

            save_json_backup(
                JSON_BACKUP,
                merged_headers,
                merged_rows,
            )

            client = google_client()

            sheet_url = create_google_sheet(
                headers=merged_headers,
                rows=merged_rows,
                latest_result_column=current_run_column,
                sheet_title=DEFAULT_GOOGLE_SHEET_TITLE,
                client=client,
            )

        return render_template(
            "success.html",
            sheet_url=sheet_url,
            processed_count=len(current_results),
            total_history=len(merged_rows),
        )

    except Exception as error:
        app.logger.exception("Processing failed")
        flash(f"התהליך נכשל: {error}", "danger")
        return redirect(url_for("index"))

    finally:
        upload_path.unlink(missing_ok=True)


@app.post("/history/clear")
def clear_history():
    confirmation = request.form.get("confirmation", "").strip()

    if confirmation != "DELETE":
        flash("מחיקת ההיסטוריה בוטלה: האישור אינו תקין.", "warning")
        return redirect(url_for("index"))

    with processing_lock:
        HISTORY_CSV.unlink(missing_ok=True)
        JSON_BACKUP.unlink(missing_ok=True)

    flash("ההיסטוריה המקומית נמחקה. ההרצה הבאה תתחיל היסטוריה חדשה.", "success")
    return redirect(url_for("index"))


@app.post("/google/disconnect")
def disconnect_google():
    TOKEN_FILE.unlink(missing_ok=True)
    flash("חשבון Google נותק מהמערכת.", "success")
    return redirect(url_for("index"))


@app.errorhandler(413)
def file_too_large(_error):
    flash(f"הקובץ גדול מדי. הגודל המרבי הוא {MAX_UPLOAD_MB}MB.", "danger")
    return redirect(url_for("index"))


if __name__ == "__main__":
    app.run(
        host=os.getenv("FLASK_HOST", "127.0.0.1"),
        port=int(os.getenv("FLASK_PORT", "5000")),
        debug=os.getenv("FLASK_DEBUG", "false").lower() == "true",
    )
