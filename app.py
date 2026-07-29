from __future__ import annotations

import csv
import os
import secrets
import threading
import uuid
from pathlib import Path
from datetime import datetime
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

from dotenv import load_dotenv
from flask import (
    Flask,
    flash,
    redirect,
    render_template,
    request,
    send_file,
    url_for,
)
from werkzeug.utils import secure_filename

from classifier_core import (
    merge_results_into_history,
    process_current_rows,
    read_history_csv,
    read_xlsx_rows,
    save_history_csv,
    save_json_backup,
)


load_dotenv()

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
UPLOAD_DIR = BASE_DIR / "uploads"

HISTORY_CSV = DATA_DIR / "image_results_history.csv"
JSON_BACKUP = DATA_DIR / "image_results_history_backup.json"

DATA_DIR.mkdir(
    parents=True,
    exist_ok=True,
)

UPLOAD_DIR.mkdir(
    parents=True,
    exist_ok=True,
)

MAX_UPLOAD_MB = int(
    os.getenv("MAX_UPLOAD_MB", "25")
)

app = Flask(__name__)

app.secret_key = (
    os.getenv("FLASK_SECRET_KEY")
    or secrets.token_hex(32)
)

app.config["MAX_CONTENT_LENGTH"] = (
    MAX_UPLOAD_MB * 1024 * 1024
)

# Prevent two processes from writing to the CSV at once.
processing_lock = threading.Lock()


def create_history_xlsx(
    headers: list[str],
    rows: list[dict[str, str]],
) -> BytesIO:
    """
    Create a Windows-friendly XLSX workbook.
    """

    if not headers:
        raise ValueError(
            "לא נמצאו כותרות ליצירת קובץ Excel."
        )

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Results"

    worksheet.sheet_view.rightToLeft = True
    worksheet.freeze_panes = "D2"
    worksheet.sheet_properties.pageSetUpPr.fitToPage = True
    worksheet.page_setup.fitToWidth = 1
    worksheet.page_setup.fitToHeight = 0

    header_fill = PatternFill(
        fill_type="solid",
        fgColor="2E5480",
    )
    header_font = Font(
        color="FFFFFF",
        bold=True,
        size=11,
    )
    header_alignment = Alignment(
        horizontal="center",
        vertical="center",
        wrap_text=True,
    )

    body_alignment = Alignment(
        horizontal="right",
        vertical="center",
        wrap_text=True,
    )
    centered_alignment = Alignment(
        horizontal="center",
        vertical="center",
        wrap_text=True,
    )
    url_alignment = Alignment(
        horizontal="left",
        vertical="center",
        wrap_text=False,
        readingOrder=1,
    )

    thin_gray = Side(
        style="thin",
        color="D9E1F2",
    )
    cell_border = Border(
        left=thin_gray,
        right=thin_gray,
        top=thin_gray,
        bottom=thin_gray,
    )

    match_fill = PatternFill(
        fill_type="solid",
        fgColor="D9EAD3",
    )
    mismatch_fill = PatternFill(
        fill_type="solid",
        fgColor="F4CCCC",
    )
    error_fill = PatternFill(
        fill_type="solid",
        fgColor="FFF2CC",
    )

    safe_headers: list[str] = []
    used_headers: set[str] = set()

    for index, original_header in enumerate(
        headers,
        start=1,
    ):
        base_header = str(
            original_header or f"Column {index}"
        ).strip()

        if not base_header:
            base_header = f"Column {index}"

        unique_header = base_header
        duplicate_number = 2

        while unique_header in used_headers:
            unique_header = (
                f"{base_header} ({duplicate_number})"
            )
            duplicate_number += 1

        used_headers.add(unique_header)
        safe_headers.append(unique_header)

    for column_index, header in enumerate(
        safe_headers,
        start=1,
    ):
        cell = worksheet.cell(
            row=1,
            column=column_index,
            value=header,
        )
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_alignment
        cell.border = cell_border

    latest_header = headers[-1] if len(headers) > 3 else ""
    expected_header = headers[2] if len(headers) >= 3 else ""

    for row_index, row in enumerate(
        rows,
        start=2,
    ):
        expected_value = str(
            row.get(expected_header, "")
        ).strip()

        latest_value = str(
            row.get(latest_header, "")
        ).strip()

        row_fill = None

        if latest_header and latest_value:
            if latest_value.startswith("שגיאה"):
                row_fill = error_fill
            elif latest_value == expected_value:
                row_fill = match_fill
            else:
                row_fill = mismatch_fill

        for column_index, original_header in enumerate(
            headers,
            start=1,
        ):
            raw_value = row.get(
                original_header,
                "",
            )
            value = "" if raw_value is None else str(raw_value)

            cell = worksheet.cell(
                row=row_index,
                column=column_index,
                value=value,
            )
            cell.border = cell_border

            if row_fill is not None:
                cell.fill = row_fill

            if column_index == 1:
                cell.alignment = url_alignment

                if value.startswith(
                    ("http://", "https://")
                ):
                    cell.hyperlink = value
                    cell.style = "Hyperlink"
                    cell.border = cell_border

                    if row_fill is not None:
                        cell.fill = row_fill

            elif column_index >= 3:
                cell.alignment = centered_alignment

            else:
                cell.alignment = body_alignment

    data_last_row = len(rows) + 1
    column_count = len(headers)
    last_column_letter = get_column_letter(
        column_count
    )

    worksheet.row_dimensions[1].height = 34
    worksheet.auto_filter.ref = (
        f"A1:{last_column_letter}{data_last_row}"
    )

    preferred_widths = {
        1: 55,
        2: 48,
        3: 22,
    }

    for column_index in range(
        1,
        column_count + 1,
    ):
        column_letter = get_column_letter(
            column_index
        )
        width = preferred_widths.get(
            column_index,
            24,
        )
        worksheet.column_dimensions[
            column_letter
        ].width = width

    if rows:
        table = Table(
            displayName="ImageClassificationHistory",
            ref=(
                f"A1:{last_column_letter}"
                f"{data_last_row}"
            ),
        )

        table.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=False,
            showColumnStripes=False,
        )

        worksheet.add_table(table)

        for cell in worksheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = header_alignment
            cell.border = cell_border

    worksheet.print_title_rows = "1:1"
    worksheet.sheet_view.showGridLines = False

    output = BytesIO()
    workbook.save(output)
    output.seek(0)

    return output


def is_xlsx(filename: str) -> bool:
    return filename.lower().endswith(".xlsx")


def load_history_for_display() -> tuple[
    list[str],
    list[dict[str, str]],
]:
    """
    Load the existing CSV history for display in the browser.
    """
    try:
        return read_history_csv(HISTORY_CSV)

    except Exception as error:
        app.logger.exception(
            "Could not read history CSV"
        )

        flash(
            f"לא ניתן לקרוא את קובץ ההיסטוריה: {error}",
            "danger",
        )

        return [], []


# ============================================================
# Main page
# ============================================================

@app.get("/")
def index():
    headers, rows = load_history_for_display()

    latest_result_column = ""

    if len(headers) > 3:
        latest_result_column = headers[-1]

    return render_template(
        "index.html",
        headers=headers,
        rows=rows,
        latest_result_column=latest_result_column,
        history_exists=bool(rows),
    )


# ============================================================
# Run XLSX processing
# ============================================================

@app.post("/run")
def run_classification():
    uploaded_file = request.files.get(
        "xlsx_file"
    )

    if (
        uploaded_file is None
        or not uploaded_file.filename
    ):
        flash(
            "לא נבחר קובץ XLSX.",
            "danger",
        )

        return redirect(url_for("index"))

    if not is_xlsx(uploaded_file.filename):
        flash(
            "ניתן להעלות קובץ XLSX בלבד.",
            "danger",
        )

        return redirect(url_for("index"))

    safe_name = (
        secure_filename(uploaded_file.filename)
        or "input.xlsx"
    )

    run_id = uuid.uuid4().hex

    upload_path = (
        UPLOAD_DIR
        / f"{run_id}_{safe_name}"
    )

    uploaded_file.save(upload_path)

    try:
        with processing_lock:
            current_xlsx_rows = read_xlsx_rows(
                upload_path
            )

            if not current_xlsx_rows:
                raise ValueError(
                    "לא נמצאו כתובות תקינות בקובץ"
                )

            current_results = process_current_rows(
                current_xlsx_rows
            )

            (
                existing_headers,
                existing_rows,
            ) = read_history_csv(HISTORY_CSV)

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
                csv_path=HISTORY_CSV,
                headers=merged_headers,
                rows=merged_rows,
            )

            save_json_backup(
                output_path=JSON_BACKUP,
                headers=merged_headers,
                rows=merged_rows,
            )

        flash(
            (
                "הבדיקה הסתיימה בהצלחה. "
                f"עובדו {len(current_results)} כתובות. "
                f"עמודת התוצאה החדשה: "
                f"{current_run_column}"
            ),
            "success",
        )

        return redirect(url_for("index"))

    except Exception as error:
        app.logger.exception(
            "Processing failed"
        )

        flash(
            f"התהליך נכשל: {error}",
            "danger",
        )

        return redirect(url_for("index"))

    finally:
        upload_path.unlink(
            missing_ok=True
        )

# ============================================================
# Download formatted XLSX
# ============================================================

@app.get("/history/download")
def download_history():
    """
    Download the complete history as a formatted XLSX file.
    """

    if not HISTORY_CSV.exists():
        flash(
            "עדיין אין היסטוריה להורדה.",
            "warning",
        )
        return redirect(url_for("index"))

    try:
        with processing_lock:
            headers, rows = read_history_csv(
                HISTORY_CSV
            )

            if not headers or not rows:
                flash(
                    "קובץ ההיסטוריה עדיין ריק.",
                    "warning",
                )
                return redirect(url_for("index"))

            xlsx_file = create_history_xlsx(
                headers=headers,
                rows=rows,
            )

        timestamp = datetime.now().strftime(
            "%Y%m%d_%H%M%S"
        )
        filename = (
            f"image_results_{timestamp}.xlsx"
        )

        response = send_file(
            xlsx_file,
            as_attachment=True,
            download_name=filename,
            mimetype=(
                "application/vnd.openxmlformats-"
                "officedocument.spreadsheetml.sheet"
            ),
            max_age=0,
        )

        response.headers["Cache-Control"] = (
            "no-store, no-cache, must-revalidate, "
            "max-age=0"
        )
        response.headers["Pragma"] = "no-cache"
        response.headers["Expires"] = "0"

        return response

    except Exception as error:
        app.logger.exception(
            "Could not create XLSX history file"
        )

        flash(
            f"יצירת קובץ ה־XLSX נכשלה: {error}",
            "danger",
        )
        return redirect(url_for("index"))


# ============================================================
# Clear history
# ============================================================

@app.post("/history/clear")
def clear_history():
    confirmation = request.form.get(
        "confirmation",
        "",
    ).strip()

    if confirmation != "DELETE":
        flash(
            (
                "מחיקת ההיסטוריה בוטלה: "
                "האישור אינו תקין."
            ),
            "warning",
        )

        return redirect(url_for("index"))

    with processing_lock:
        HISTORY_CSV.unlink(
            missing_ok=True
        )

        JSON_BACKUP.unlink(
            missing_ok=True
        )

    flash(
        (
            "ההיסטוריה נמחקה. "
            "ההרצה הבאה תתחיל היסטוריה חדשה."
        ),
        "success",
    )

    return redirect(url_for("index"))


# ============================================================
# File too large
# ============================================================

@app.errorhandler(413)
def file_too_large(_error):
    flash(
        (
            "הקובץ גדול מדי. "
            f"הגודל המרבי הוא {MAX_UPLOAD_MB}MB."
        ),
        "danger",
    )

    return redirect(url_for("index"))


# ============================================================
# Start server
# ============================================================

if __name__ == "__main__":
    app.run(
        host=os.getenv(
            "FLASK_HOST",
            "0.0.0.0",
        ),
        port=int(
            os.getenv(
                "FLASK_PORT",
                "5000",
            )
        ),
        debug=(
            os.getenv(
                "FLASK_DEBUG",
                "false",
            ).lower()
            == "true"
        ),
    )