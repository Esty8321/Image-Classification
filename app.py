from __future__ import annotations
import os
import secrets
import threading
import uuid
from pathlib import Path
from datetime import datetime
from io import BytesIO
from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import (
    Alignment,
    Font,
    PatternFill,
)
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import (
    Table,
    TableStyleInfo,
)
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
)


load_dotenv()

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
UPLOAD_DIR = BASE_DIR / "uploads"

HISTORY_CSV = DATA_DIR / "image_results_history.csv"

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

def calculate_mismatch_summaries(
    headers: list[str],
    rows: list[dict[str, str]],
) -> dict[str, dict[str, object]]:
    summaries: dict[str, dict[str, object]] = {}

    if len(headers) <= 3:
        return summaries

    expected_header = headers[2]

    # All columns after the first three are dated result columns.
    for result_header in headers[3:]:
        checked_count = 0
        mismatch_count = 0

        for row in rows:
            expected_value = str(
                row.get(expected_header, "")
            ).strip()

            actual_value = str(
                row.get(result_header, "")
            ).strip()

            # A blank value means that this URL did not
            # participate in this historical run.
            if not actual_value:
                continue

            checked_count += 1

            if actual_value != expected_value:
                mismatch_count += 1

        percentage_value = (
            mismatch_count / checked_count
            if checked_count
            else None
        )

        summaries[result_header] = {
            "checked_count": checked_count,
            "mismatch_count": mismatch_count,
            "percentage_value": percentage_value,
            "display": (
                f"{percentage_value:.2%}"
                if percentage_value is not None
                else "—"
            ),
        }

    return summaries


def create_history_xlsx(
    headers: list[str],
    rows: list[dict[str, str]],
) -> BytesIO:

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Results"

    worksheet.sheet_view.rightToLeft = True

    # Freeze header row and first three columns.
    worksheet.freeze_panes = "D2"

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
        wrap_text=True,
        readingOrder=1,
    )

    # Special style for the percentage summary row.
    summary_fill = PatternFill(
        fill_type="solid",
        fgColor="DCEEFF",
    )

    summary_font = Font(
        color="003366",
        bold=True,
        size=11,
    )

    # ========================================================
    # Write headers
    # ========================================================

    for column_index, header in enumerate(
        headers,
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

    # ========================================================
    # Write normal history rows
    # ========================================================

    for row_index, row in enumerate(
        rows,
        start=2,
    ):
        for column_index, header in enumerate(
            headers,
            start=1,
        ):
            value = row.get(header, "")

            cell = worksheet.cell(
                row=row_index,
                column=column_index,
                value=value,
            )

            if column_index == 1:
                cell.alignment = url_alignment

            elif column_index >= 3:
                cell.alignment = centered_alignment

            else:
                cell.alignment = body_alignment

    # Row 1 is the header.
    # Data starts on row 2.
    data_last_row = len(rows) + 1

    # The percentage row is one row after the data.
    summary_row_number = data_last_row + 1

    column_count = len(headers)

    mismatch_summaries = calculate_mismatch_summaries(
        headers=headers,
        rows=rows,
    )

    # ========================================================
    # Write final percentage row
    # ========================================================

    for column_index, header in enumerate(
        headers,
        start=1,
    ):
        cell = worksheet.cell(
            row=summary_row_number,
            column=column_index,
        )

        if column_index == 2:
            cell.value = "אחוז אי־התאמה"

        elif column_index >= 4:
            summary = mismatch_summaries.get(header)

            if (
                summary
                and summary["percentage_value"] is not None
            ):
                # Store as a real numeric Excel percentage,
                # not as text.
                cell.value = summary["percentage_value"]
                cell.number_format = "0.00%"

            else:
                cell.value = "—"

        else:
            cell.value = ""

        cell.fill = summary_fill
        cell.font = summary_font
        cell.alignment = centered_alignment

    # ========================================================
    # Row heights and column widths
    # ========================================================

    worksheet.row_dimensions[1].height = 36

    worksheet.row_dimensions[
        summary_row_number
    ].height = 30

    if column_count >= 1:
        worksheet.column_dimensions["A"].width = 55

    if column_count >= 2:
        worksheet.column_dimensions["B"].width = 48

    if column_count >= 3:
        worksheet.column_dimensions["C"].width = 22

    for column_index in range(
        4,
        column_count + 1,
    ):
        column_letter = get_column_letter(
            column_index
        )

        worksheet.column_dimensions[
            column_letter
        ].width = 24

    # ========================================================
    # Native Excel table
    # ========================================================

    # The percentage row is intentionally outside the Excel
    # table, so sorting/filtering does not hide or mix it
    # with normal result rows.
    if headers and rows:
        last_column_letter = get_column_letter(
            column_count
        )

        table_reference = (
            f"A1:{last_column_letter}{data_last_row}"
        )

        table = Table(
            displayName="ImageClassificationHistory",
            ref=table_reference,
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

    # ========================================================
    # Conditional formatting for normal rows
    # ========================================================

    if rows and len(headers) > 3:
        latest_column_index = len(headers)

        latest_column_letter = get_column_letter(
            latest_column_index
        )

        data_range = (
            f"A2:"
            f"{get_column_letter(column_count)}"
            f"{data_last_row}"
        )

        green_fill = PatternFill(
            fill_type="solid",
            fgColor="D6EFD6",
        )

        red_fill = PatternFill(
            fill_type="solid",
            fgColor="F5D0D0",
        )

        yellow_fill = PatternFill(
            fill_type="solid",
            fgColor="FFEDAC",
        )

        green_formula = (
            f'AND('
            f'${latest_column_letter}2<>"",'
            f'LEFT(${latest_column_letter}2,5)'
            f'<>"שגיאה",'
            f'$C2=${latest_column_letter}2'
            f')'
        )

        worksheet.conditional_formatting.add(
            data_range,
            FormulaRule(
                formula=[green_formula],
                fill=green_fill,
            ),
        )

        red_formula = (
            f'AND('
            f'${latest_column_letter}2<>"",'
            f'LEFT(${latest_column_letter}2,5)'
            f'<>"שגיאה",'
            f'$C2<>${latest_column_letter}2'
            f')'
        )

        worksheet.conditional_formatting.add(
            data_range,
            FormulaRule(
                formula=[red_formula],
                fill=red_fill,
            ),
        )

        yellow_formula = (
            f'LEFT('
            f'${latest_column_letter}2,5'
            f')="שגיאה"'
        )

        worksheet.conditional_formatting.add(
            data_range,
            FormulaRule(
                formula=[yellow_formula],
                fill=yellow_fill,
            ),
        )

    # Filters cover only the normal data rows.
    if headers:
        last_column_letter = get_column_letter(
            column_count
        )

        worksheet.auto_filter.ref = (
            f"A1:{last_column_letter}{data_last_row}"
        )

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



@app.get("/")
def index():
    headers, rows = load_history_for_display()

    latest_result_column = ""

    if len(headers) > 3:
        latest_result_column = headers[-1]

    mismatch_summaries = calculate_mismatch_summaries(
        headers=headers,
        rows=rows,
    )

    return render_template(
        "index.html",
        headers=headers,
        rows=rows,
        latest_result_column=latest_result_column,
        result_columns=headers[3:],
        mismatch_summaries=mismatch_summaries,
        history_exists=bool(rows),
    )


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