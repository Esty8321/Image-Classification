from __future__ import annotations

import argparse
import csv
import hashlib
import json
import re
import struct
import time
from datetime import datetime
from pathlib import Path
from typing import Any
from urllib.parse import urlparse

import gspread
import requests
from gspread.exceptions import APIError
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry


# ============================================================
# Endpoint configuration
# ============================================================

SERVER_IP = "81.28.7.93"
SERVER_PORT = 80
ENDPOINT = "/gpu-server-api/predict-binary32-priority"

PRIORITY = 0

IMAGE_DOWNLOAD_TIMEOUT_SECONDS = 30
ENDPOINT_TIMEOUT_SECONDS = 60

# Maximum downloaded image size: 30 MB
MAX_IMAGE_SIZE_BYTES = 30 * 1024 * 1024


# ============================================================
# Google configuration
# ============================================================

GOOGLE_CREDENTIALS_FILE = Path("credentials.json")
GOOGLE_TOKEN_FILE = Path("token.json")

DEFAULT_GOOGLE_SHEET_TITLE = "Image Classification History"


# ============================================================
# XLSX columns
# ============================================================

URL_COLUMN = "B"
EXPECTED_STATUS_COLUMN = "E"
DESCRIPTION_COLUMN = "F"


# ============================================================
# CSV column names
# ============================================================

CSV_URL_COLUMN = "כתובת התמונה"
CSV_DESCRIPTION_COLUMN = "תיאור"
CSV_EXPECTED_COLUMN = "סטטוס צפוי"

ACTUAL_STATUS_PREFIX = "סטטוס בפועל "


# ============================================================
# HTTP session
# ============================================================

def create_http_session() -> requests.Session:
    """
    Create a requests session with retry support.
    """
    retry_strategy = Retry(
        total=3,
        connect=3,
        read=3,
        backoff_factor=1,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["GET", "POST"],
        raise_on_status=False,
    )

    adapter = HTTPAdapter(max_retries=retry_strategy)

    session = requests.Session()
    session.mount("http://", adapter)
    session.mount("https://", adapter)

    session.headers.update(
        {
            "User-Agent": (
                "Mozilla/5.0 (X11; Linux x86_64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/150.0.0.0 Safari/537.36"
            ),
            "Accept": (
                "image/avif,image/webp,image/apng,image/svg+xml,"
                "image/*,*/*;q=0.8"
            ),
        }
    )

    return session


# ============================================================
# General helpers
# ============================================================

def clean_cell_value(value: Any) -> str:
    if value is None:
        return ""

    return str(value).strip()


def extract_url_from_excel_value(value: Any) -> str:
    """
    Support regular URLs and Excel HYPERLINK formulas.

    Example:
        =HYPERLINK("https://example.com/image.jpg", "image")
    """
    text = clean_cell_value(value)

    if not text:
        return ""

    hyperlink_match = re.match(
        r'''^=HYPERLINK\(\s*["']([^"']+)["']''',
        text,
        flags=re.IGNORECASE,
    )

    if hyperlink_match:
        return hyperlink_match.group(1).strip()

    return text


def normalize_image_url(raw_url: str) -> str:
    """
    Add https:// when the URL does not contain a protocol.
    """
    url = raw_url.strip()

    if not url:
        raise ValueError("The image URL is empty")

    if url.startswith("//"):
        url = f"https:{url}"

    elif not re.match(r"^https?://", url, flags=re.IGNORECASE):
        url = f"https://{url}"

    parsed = urlparse(url)

    if parsed.scheme not in {"http", "https"}:
        raise ValueError(
            f"Unsupported URL protocol: {parsed.scheme}"
        )

    if not parsed.netloc:
        raise ValueError(f"Invalid image URL: {url}")

    return url


def create_url_key(url: str) -> str:
    normalized = normalize_image_url(url)
    parsed = urlparse(normalized)

    host = parsed.netloc.lower().strip()
    path = parsed.path or "/"

    if path != "/":
        path = path.rstrip("/")

    query = parsed.query.strip()

    key = f"{host}{path}"

    if query:
        key += f"?{query}"

    return key


def normalize_expected_status(value: Any) -> str:
    status = clean_cell_value(value)

    normalized_statuses = {
        "פתוח": "פתוח",
        "חסום": "חסום",
        "0": "פתוח",
        "-1": "חסום",
    }

    normalized = normalized_statuses.get(status.lower())

    if normalized:
        return normalized

    return status


def build_referer(image_url: str) -> str:
    parsed = urlparse(image_url)

    return f"{parsed.scheme}://{parsed.netloc}/"


def create_run_column_name(
    existing_headers: list[str],
) -> str:
    base_name = (
        ACTUAL_STATUS_PREFIX
        + datetime.now().strftime("%Y-%m-%d %H-%M-%S")
    )

    column_name = base_name
    counter = 2

    while column_name in existing_headers:
        column_name = f"{base_name} ({counter})"
        counter += 1

    return column_name


def find_first_data_row(worksheet) -> int | None:
    for row_number in range(1, worksheet.max_row + 1):
        raw_url = worksheet[f"{URL_COLUMN}{row_number}"].value
        extracted_url = extract_url_from_excel_value(raw_url)

        if extracted_url:
            return row_number

    return None


def read_xlsx_rows(xlsx_path: Path) -> list[dict[str, Any]]:
    """
    Read:
        Column B: image URL
        Column E: expected status
        Column F: description
    """
    if not xlsx_path.exists():
        raise FileNotFoundError(
            f"XLSX file not found: {xlsx_path.resolve()}"
        )

    if not xlsx_path.is_file():
        raise ValueError(
            f"The XLSX path is not a file: {xlsx_path.resolve()}"
        )

    try:
        workbook = load_workbook(
            filename=xlsx_path,
            read_only=True,
            data_only=False,
        )

    except InvalidFileException as error:
        raise ValueError(
            f"The file is not a valid XLSX file: {xlsx_path}"
        ) from error

    try:
        worksheet = workbook.active

        first_data_row = find_first_data_row(worksheet)

        if first_data_row is None:
            raise ValueError(
                f"No URL was found in column {URL_COLUMN}"
            )

        print(f"First data row: {first_data_row}")
        print(f"Last worksheet row: {worksheet.max_row}")

        rows: list[dict[str, Any]] = []
        urls_seen_in_current_file: set[str] = set()

        for row_number in range(
            first_data_row,
            worksheet.max_row + 1,
        ):
            raw_url = worksheet[f"{URL_COLUMN}{row_number}"].value

            raw_expected_status = worksheet[
                f"{EXPECTED_STATUS_COLUMN}{row_number}"
            ].value

            raw_description = worksheet[
                f"{DESCRIPTION_COLUMN}{row_number}"
            ].value

            extracted_url = extract_url_from_excel_value(raw_url)

            if not extracted_url:
                continue

            try:
                normalized_url = normalize_image_url(extracted_url)
                url_key = create_url_key(normalized_url)

            except Exception as error:
                print(
                    f"Skipping invalid URL in XLSX row "
                    f"{row_number}: {error}"
                )
                continue

            # Prevent processing the same URL twice in one XLSX file.
            if url_key in urls_seen_in_current_file:
                print(
                    f"Duplicate URL skipped in XLSX row "
                    f"{row_number}: {normalized_url}"
                )
                continue

            urls_seen_in_current_file.add(url_key)

            rows.append(
                {
                    "source_row": row_number,
                    "url": normalized_url,
                    "url_key": url_key,
                    "expected_status": normalize_expected_status(
                        raw_expected_status
                    ),
                    "description": clean_cell_value(
                        raw_description
                    ),
                }
            )

        return rows

    finally:
        workbook.close()


# ============================================================
# Image downloading
# ============================================================

def download_image(
    session: requests.Session,
    image_url: str,
) -> tuple[bytes, str]:
    """
    Download the image and return:
        image bytes, final URL after redirects
    """
    referer = build_referer(image_url)

    with session.get(
        image_url,
        headers={"Referer": referer},
        timeout=IMAGE_DOWNLOAD_TIMEOUT_SECONDS,
        allow_redirects=True,
        stream=True,
    ) as response:
        response.raise_for_status()

        content_type = (
            response.headers.get("Content-Type", "")
            .lower()
            .strip()
        )

        if (
            "text/html" in content_type
            or "application/xhtml" in content_type
        ):
            raise ValueError(
                "The URL returned HTML instead of an image. "
                f"Content-Type: {content_type}"
            )

        content_length = response.headers.get("Content-Length")

        if content_length:
            try:
                declared_size = int(content_length)

                if declared_size > MAX_IMAGE_SIZE_BYTES:
                    raise ValueError(
                        "The image is too large. "
                        f"Declared size: {declared_size} bytes"
                    )

            except ValueError as error:
                if "too large" in str(error):
                    raise

        chunks: list[bytes] = []
        total_size = 0

        for chunk in response.iter_content(
            chunk_size=64 * 1024
        ):
            if not chunk:
                continue

            total_size += len(chunk)

            if total_size > MAX_IMAGE_SIZE_BYTES:
                raise ValueError(
                    "The downloaded image exceeded the maximum "
                    f"allowed size of {MAX_IMAGE_SIZE_BYTES} bytes"
                )

            chunks.append(chunk)

        image_bytes = b"".join(chunks)

        if not image_bytes:
            raise ValueError("The downloaded image is empty")

        return image_bytes, response.url


# ============================================================
# Binary endpoint
# ============================================================

def null_terminated(value: str) -> bytes:
    return value.encode("utf-8") + b"\x00"


def create_request_body(
    image_bytes: bytes,
    page_url: str,
    referer: str,
    priority: int,
) -> tuple[bytes, str, str]:
    if not 0 <= priority <= 0xFFFFFFFF:
        raise ValueError(
            "priority must be between 0 and 4294967295"
        )

    if len(image_bytes) > 0xFFFFFFFF:
        raise ValueError(
            "The image is too large for a uint32 size field"
        )

    timestamp = str(int(time.time()))

    key_source = timestamp + "netspark_enativ_@~&^CSdfg544hjpsa"
    key = hashlib.md5(
        key_source.encode("utf-8")
    ).hexdigest()

    priority_bytes = struct.pack("!I", priority)
    image_size_bytes = struct.pack("!I", len(image_bytes))

    body = b"".join(
        [
            null_terminated(page_url),
            null_terminated(referer),
            null_terminated(key),
            null_terminated(timestamp),
            priority_bytes,
            image_size_bytes,
            image_bytes,
        ]
    )

    return body, timestamp, key


def extract_endpoint_status(result: Any) -> int:
    if not isinstance(result, dict):
        raise ValueError(
            "The endpoint response must be a JSON object"
        )

    if "status" not in result:
        raise ValueError(
            "The endpoint JSON does not contain a status field"
        )

    raw_status = result["status"]

    try:
        status = int(raw_status)

    except (TypeError, ValueError) as error:
        raise ValueError(
            f"Invalid endpoint status value: {raw_status!r}"
        ) from error

    if status not in {0, -1}:
        raise ValueError(
            f"Unexpected endpoint status: {status}. "
            "Expected 0 or -1."
        )

    return status


def endpoint_status_to_hebrew(status: int) -> str:
    if status == 0:
        return "פתוח"

    if status == -1:
        return "חסום"

    raise ValueError(f"Unsupported status: {status}")


def send_image_to_endpoint(
    session: requests.Session,
    image_bytes: bytes,
    image_url: str,
) -> int:
    referer = build_referer(image_url)

    body, timestamp, key = create_request_body(
        image_bytes=image_bytes,
        page_url=image_url,
        referer=referer,
        priority=PRIORITY,
    )

    request_url = (
        f"http://{SERVER_IP}:{SERVER_PORT}{ENDPOINT}"
    )

    response = session.post(
        request_url,
        data=body,
        headers={
            "Content-Type": "application/octet-stream",
            "Content-Length": str(len(body)),
        },
        timeout=ENDPOINT_TIMEOUT_SECONDS,
    )

    response.raise_for_status()

    if not response.content:
        raise ValueError(
            "The endpoint returned an empty response"
        )

    try:
        result = response.json()

    except requests.exceptions.JSONDecodeError as error:
        response_preview = response.text[:500]

        raise ValueError(
            "The endpoint response is not valid JSON. "
            f"Response: {response_preview}"
        ) from error

    status = extract_endpoint_status(result)

    print(
        f"    Endpoint timestamp={timestamp}, "
        f"key={key}, status={status}"
    )

    return status


# ============================================================
# Process the current XLSX
# ============================================================

def process_current_rows(
    rows: list[dict[str, Any]],
) -> dict[str, dict[str, str]]:
    """
    Process all rows from the current XLSX.

    Returns a dictionary by normalized URL key.
    """
    session = create_http_session()

    results: dict[str, dict[str, str]] = {}

    total = len(rows)

    for index, row in enumerate(rows, start=1):
        source_row = row["source_row"]
        image_url = row["url"]
        url_key = row["url_key"]
        description = row["description"]
        expected_status = row["expected_status"]

        print()
        print(
            f"[{index}/{total}] Processing XLSX row "
            f"{source_row}"
        )
        print(f"    URL: {image_url}")
        print(f"    Expected: {expected_status}")

        actual_status: str

        try:
            image_bytes, final_url = download_image(
                session=session,
                image_url=image_url,
            )

            print(f"    Final URL: {final_url}")
            print(
                f"    Downloaded size: "
                f"{len(image_bytes)} bytes"
            )

            endpoint_status = send_image_to_endpoint(
                session=session,
                image_bytes=image_bytes,
                image_url=final_url,
            )

            actual_status = endpoint_status_to_hebrew(
                endpoint_status
            )

            if expected_status == actual_status:
                print(
                    f"    Actual: {actual_status} — match"
                )
            else:
                print(
                    f"    Actual: {actual_status} — "
                    f"does not match"
                )

        except requests.exceptions.ConnectTimeout as error:
            actual_status = (
                f"שגיאה: חריגת זמן בחיבור — {error}"
            )
            print(f"    {actual_status}")

        except requests.exceptions.ReadTimeout as error:
            actual_status = (
                f"שגיאה: חריגת זמן בתגובה — {error}"
            )
            print(f"    {actual_status}")

        except requests.exceptions.HTTPError as error:
            status_code = (
                error.response.status_code
                if error.response is not None
                else "unknown"
            )

            actual_status = f"שגיאה: HTTP {status_code}"
            print(f"    {actual_status}: {error}")

        except requests.exceptions.RequestException as error:
            actual_status = (
                f"שגיאה: בקשת רשת נכשלה — {error}"
            )
            print(f"    {actual_status}")

        except Exception as error:
            actual_status = f"שגיאה: {error}"
            print(f"    {actual_status}")

        results[url_key] = {
            CSV_URL_COLUMN: image_url,
            CSV_DESCRIPTION_COLUMN: description,
            CSV_EXPECTED_COLUMN: expected_status,
            "actual_status": actual_status,
        }

    return results


# ============================================================
# CSV history
# ============================================================

def get_base_headers() -> list[str]:
    return [
        CSV_URL_COLUMN,
        CSV_DESCRIPTION_COLUMN,
        CSV_EXPECTED_COLUMN,
    ]


def read_history_csv(
    csv_path: Path,
) -> tuple[list[str], list[dict[str, str]]]:
    """
    Read the existing CSV history.

    When the file does not exist, return an empty history.
    """
    if not csv_path.exists():
        return get_base_headers(), []

    if not csv_path.is_file():
        raise ValueError(
            f"CSV history path is not a file: "
            f"{csv_path.resolve()}"
        )

    with csv_path.open(
        "r",
        encoding="utf-8-sig",
        newline="",
    ) as csv_file:
        reader = csv.DictReader(csv_file)

        if reader.fieldnames is None:
            return get_base_headers(), []

        headers = [
            clean_cell_value(header)
            for header in reader.fieldnames
            if clean_cell_value(header)
        ]

        required_headers = set(get_base_headers())
        missing_headers = required_headers - set(headers)

        if missing_headers:
            raise ValueError(
                "The history CSV is missing required columns: "
                + ", ".join(sorted(missing_headers))
            )

        rows: list[dict[str, str]] = []

        for raw_row in reader:
            row: dict[str, str] = {}

            for header in headers:
                row[header] = clean_cell_value(
                    raw_row.get(header, "")
                )

            if row.get(CSV_URL_COLUMN):
                rows.append(row)

        return headers, rows


def merge_results_into_history(
    existing_headers: list[str],
    existing_rows: list[dict[str, str]],
    current_results: dict[str, dict[str, str]],
) -> tuple[list[str], list[dict[str, str]], str]:
    """
    Merge the current run into the CSV history.

    Existing URLs remain on their original rows.
    New URLs are appended at the bottom.
    """
    headers = list(existing_headers)

    for base_header in get_base_headers():
        if base_header not in headers:
            headers.insert(
                get_base_headers().index(base_header),
                base_header,
            )

    current_run_column = create_run_column_name(headers)
    headers.append(current_run_column)

    history_by_key: dict[str, dict[str, str]] = {}
    ordered_keys: list[str] = []

    # Load existing rows while keeping the original order.
    for row in existing_rows:
        url = clean_cell_value(row.get(CSV_URL_COLUMN))

        if not url:
            continue

        try:
            url_key = create_url_key(url)

        except Exception:
            # Keep invalid historical URLs based on their raw text.
            url_key = f"raw::{url.lower()}"

        if url_key in history_by_key:
            print(
                "Duplicate URL found in history CSV and skipped: "
                f"{url}"
            )
            continue

        normalized_row = {
            header: clean_cell_value(row.get(header, ""))
            for header in headers
        }

        # The new run column is blank by default.
        normalized_row[current_run_column] = ""

        history_by_key[url_key] = normalized_row
        ordered_keys.append(url_key)

    # Update existing rows or append new rows.
    for url_key, current_result in current_results.items():
        if url_key in history_by_key:
            history_row = history_by_key[url_key]

            # Keep the same row, but update the current information.
            history_row[CSV_URL_COLUMN] = current_result[
                CSV_URL_COLUMN
            ]

            history_row[CSV_DESCRIPTION_COLUMN] = current_result[
                CSV_DESCRIPTION_COLUMN
            ]

            history_row[CSV_EXPECTED_COLUMN] = current_result[
                CSV_EXPECTED_COLUMN
            ]

            history_row[current_run_column] = current_result[
                "actual_status"
            ]

        else:
            new_row = {
                header: ""
                for header in headers
            }

            new_row[CSV_URL_COLUMN] = current_result[
                CSV_URL_COLUMN
            ]

            new_row[CSV_DESCRIPTION_COLUMN] = current_result[
                CSV_DESCRIPTION_COLUMN
            ]

            new_row[CSV_EXPECTED_COLUMN] = current_result[
                CSV_EXPECTED_COLUMN
            ]

            new_row[current_run_column] = current_result[
                "actual_status"
            ]

            history_by_key[url_key] = new_row
            ordered_keys.append(url_key)

    merged_rows = [
        history_by_key[url_key]
        for url_key in ordered_keys
    ]

    return headers, merged_rows, current_run_column


def save_history_csv(
    csv_path: Path,
    headers: list[str],
    rows: list[dict[str, str]],
) -> None:
    """
    Save the complete history to CSV.

    utf-8-sig is used so Hebrew text displays correctly in Excel.
    """
    csv_path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    # Write to a temporary file first, preventing damage to the
    # existing history if writing is interrupted.
    temporary_path = csv_path.with_suffix(
        csv_path.suffix + ".tmp"
    )

    with temporary_path.open(
        "w",
        encoding="utf-8-sig",
        newline="",
    ) as csv_file:
        writer = csv.DictWriter(
            csv_file,
            fieldnames=headers,
            extrasaction="ignore",
        )

        writer.writeheader()

        for row in rows:
            writer.writerow(
                {
                    header: clean_cell_value(
                        row.get(header, "")
                    )
                    for header in headers
                }
            )

    temporary_path.replace(csv_path)


# ============================================================
# Google Sheets
# ============================================================

def authenticate_google() -> gspread.Client:
    if not GOOGLE_CREDENTIALS_FILE.exists():
        raise FileNotFoundError(
            "\nGoogle credentials file was not found.\n"
            f"Expected file: "
            f"{GOOGLE_CREDENTIALS_FILE.resolve()}\n"
            "Download Desktop OAuth credentials and save the "
            "file as credentials.json."
        )

    return gspread.oauth(
        credentials_filename=str(
            GOOGLE_CREDENTIALS_FILE
        ),
        authorized_user_filename=str(
            GOOGLE_TOKEN_FILE
        ),
    )


def column_number_to_letter(
    column_number: int,
) -> str:
    """
    Convert:
        1 -> A
        26 -> Z
        27 -> AA
    """
    letters = ""

    while column_number > 0:
        column_number, remainder = divmod(
            column_number - 1,
            26,
        )

        letters = chr(65 + remainder) + letters

    return letters


def create_google_sheet(
    headers: list[str],
    rows: list[dict[str, str]],
    latest_result_column: str,
    sheet_title: str,
    client: gspread.Client | None = None,
) -> str:
    """
    Create a Google Sheet containing all historical results.
    """
    print()
    print("Connecting to Google Sheets...")

    if client is None:
        client = authenticate_google()

    timestamp = datetime.now().strftime(
        "%Y-%m-%d %H-%M-%S"
    )

    full_title = f"{sheet_title} - {timestamp}"

    spreadsheet = client.create(full_title)

    worksheet = spreadsheet.sheet1
    worksheet.update_title("Results")

    values = [headers]

    for row in rows:
        values.append(
            [
                clean_cell_value(row.get(header, ""))
                for header in headers
            ]
        )

    row_count = max(len(values), 2)
    column_count = len(headers)

    worksheet.resize(
        rows=row_count,
        cols=column_count,
    )

    last_column_letter = column_number_to_letter(
        column_count
    )

    worksheet.update(
        range_name=(
            f"A1:{last_column_letter}{len(values)}"
        ),
        values=values,
        value_input_option="USER_ENTERED",
    )

    latest_column_index = headers.index(
        latest_result_column
    )

    # Google Sheets indexes start at zero in formatting requests.
    latest_column_number_for_formula = (
        latest_column_index + 1
    )

    latest_column_letter = column_number_to_letter(
        latest_column_number_for_formula
    )

    requests_to_apply: list[dict[str, Any]] = [
        # Freeze the first row and the first three columns.
        {
            "updateSheetProperties": {
                "properties": {
                    "sheetId": worksheet.id,
                    "gridProperties": {
                        "frozenRowCount": 1,
                        "frozenColumnCount": 3,
                    },
                    "rightToLeft": True,
                },
                "fields": (
                    "gridProperties.frozenRowCount,"
                    "gridProperties.frozenColumnCount,"
                    "rightToLeft"
                ),
            }
        },

        # Header style.
        {
            "repeatCell": {
                "range": {
                    "sheetId": worksheet.id,
                    "startRowIndex": 0,
                    "endRowIndex": 1,
                    "startColumnIndex": 0,
                    "endColumnIndex": column_count,
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {
                            "red": 0.18,
                            "green": 0.33,
                            "blue": 0.55,
                        },
                        "textFormat": {
                            "foregroundColor": {
                                "red": 1,
                                "green": 1,
                                "blue": 1,
                            },
                            "bold": True,
                            "fontSize": 11,
                        },
                        "horizontalAlignment": "CENTER",
                        "verticalAlignment": "MIDDLE",
                        "wrapStrategy": "WRAP",
                    }
                },
                "fields": "userEnteredFormat",
            }
        },

        # Body formatting.
        {
            "repeatCell": {
                "range": {
                    "sheetId": worksheet.id,
                    "startRowIndex": 1,
                    "endRowIndex": row_count,
                    "startColumnIndex": 0,
                    "endColumnIndex": column_count,
                },
                "cell": {
                    "userEnteredFormat": {
                        "wrapStrategy": "WRAP",
                        "verticalAlignment": "MIDDLE",
                    }
                },
                "fields": (
                    "userEnteredFormat.wrapStrategy,"
                    "userEnteredFormat.verticalAlignment"
                ),
            }
        },

        # URL column displayed from left to right.
        {
            "repeatCell": {
                "range": {
                    "sheetId": worksheet.id,
                    "startRowIndex": 1,
                    "endRowIndex": row_count,
                    "startColumnIndex": 0,
                    "endColumnIndex": 1,
                },
                "cell": {
                    "userEnteredFormat": {
                        "textDirection": "LEFT_TO_RIGHT",
                        "horizontalAlignment": "LEFT",
                    }
                },
                "fields": (
                    "userEnteredFormat.textDirection,"
                    "userEnteredFormat.horizontalAlignment"
                ),
            }
        },

        # Center status columns.
        {
            "repeatCell": {
                "range": {
                    "sheetId": worksheet.id,
                    "startRowIndex": 1,
                    "endRowIndex": row_count,
                    "startColumnIndex": 2,
                    "endColumnIndex": column_count,
                },
                "cell": {
                    "userEnteredFormat": {
                        "horizontalAlignment": "CENTER",
                    }
                },
                "fields": (
                    "userEnteredFormat.horizontalAlignment"
                ),
            }
        },

        # URL width.
        {
            "updateDimensionProperties": {
                "range": {
                    "sheetId": worksheet.id,
                    "dimension": "COLUMNS",
                    "startIndex": 0,
                    "endIndex": 1,
                },
                "properties": {
                    "pixelSize": 380,
                },
                "fields": "pixelSize",
            }
        },

        # Description width.
        {
            "updateDimensionProperties": {
                "range": {
                    "sheetId": worksheet.id,
                    "dimension": "COLUMNS",
                    "startIndex": 1,
                    "endIndex": 2,
                },
                "properties": {
                    "pixelSize": 350,
                },
                "fields": "pixelSize",
            }
        },

        # Expected status and historical result widths.
        {
            "updateDimensionProperties": {
                "range": {
                    "sheetId": worksheet.id,
                    "dimension": "COLUMNS",
                    "startIndex": 2,
                    "endIndex": column_count,
                },
                "properties": {
                    "pixelSize": 175,
                },
                "fields": "pixelSize",
            }
        },

        # Header height.
        {
            "updateDimensionProperties": {
                "range": {
                    "sheetId": worksheet.id,
                    "dimension": "ROWS",
                    "startIndex": 0,
                    "endIndex": 1,
                },
                "properties": {
                    "pixelSize": 48,
                },
                "fields": "pixelSize",
            }
        },

        # Filter over all columns.
        {
            "setBasicFilter": {
                "filter": {
                    "range": {
                        "sheetId": worksheet.id,
                        "startRowIndex": 0,
                        "endRowIndex": row_count,
                        "startColumnIndex": 0,
                        "endColumnIndex": column_count,
                    }
                }
            }
        },
    ]

    if rows:
        # Green: latest result matches expected result.
        requests_to_apply.append(
            {
                "addConditionalFormatRule": {
                    "index": 0,
                    "rule": {
                        "ranges": [
                            {
                                "sheetId": worksheet.id,
                                "startRowIndex": 1,
                                "endRowIndex": row_count,
                                "startColumnIndex": 0,
                                "endColumnIndex": column_count,
                            }
                        ],
                        "booleanRule": {
                            "condition": {
                                "type": "CUSTOM_FORMULA",
                                "values": [
                                    {
                                        "userEnteredValue": (
                                            f'=AND('
                                            f'${latest_column_letter}2<>"",'
                                            f'LEFT(${latest_column_letter}2,5)'
                                            f'<>"שגיאה",'
                                            f'$C2=${latest_column_letter}2)'
                                        )
                                    }
                                ],
                            },
                            "format": {
                                "backgroundColor": {
                                    "red": 0.84,
                                    "green": 0.94,
                                    "blue": 0.84,
                                }
                            },
                        },
                    },
                }
            }
        )

        # Red: latest result does not match expected result.
        requests_to_apply.append(
            {
                "addConditionalFormatRule": {
                    "index": 1,
                    "rule": {
                        "ranges": [
                            {
                                "sheetId": worksheet.id,
                                "startRowIndex": 1,
                                "endRowIndex": row_count,
                                "startColumnIndex": 0,
                                "endColumnIndex": column_count,
                            }
                        ],
                        "booleanRule": {
                            "condition": {
                                "type": "CUSTOM_FORMULA",
                                "values": [
                                    {
                                        "userEnteredValue": (
                                            f'=AND('
                                            f'${latest_column_letter}2<>"",'
                                            f'LEFT(${latest_column_letter}2,5)'
                                            f'<>"שגיאה",'
                                            f'$C2<>${latest_column_letter}2)'
                                        )
                                    }
                                ],
                            },
                            "format": {
                                "backgroundColor": {
                                    "red": 0.96,
                                    "green": 0.82,
                                    "blue": 0.82,
                                }
                            },
                        },
                    },
                }
            }
        )

        # Yellow: latest result is an error.
        requests_to_apply.append(
            {
                "addConditionalFormatRule": {
                    "index": 2,
                    "rule": {
                        "ranges": [
                            {
                                "sheetId": worksheet.id,
                                "startRowIndex": 1,
                                "endRowIndex": row_count,
                                "startColumnIndex": 0,
                                "endColumnIndex": column_count,
                            }
                        ],
                        "booleanRule": {
                            "condition": {
                                "type": "CUSTOM_FORMULA",
                                "values": [
                                    {
                                        "userEnteredValue": (
                                            f'=LEFT('
                                            f'${latest_column_letter}2,5)'
                                            f'="שגיאה"'
                                        )
                                    }
                                ],
                            },
                            "format": {
                                "backgroundColor": {
                                    "red": 1,
                                    "green": 0.93,
                                    "blue": 0.68,
                                }
                            },
                        },
                    },
                }
            }
        )

    spreadsheet.batch_update(
        {"requests": requests_to_apply}
    )

    return spreadsheet.url


# ============================================================
# Optional JSON backup
# ============================================================

def save_json_backup(
    output_path: Path,
    headers: list[str],
    rows: list[dict[str, str]],
) -> None:
    output_path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    data = {
        "headers": headers,
        "rows": rows,
    }

    output_path.write_text(
        json.dumps(
            data,
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )


# ============================================================
# Command line arguments
# ============================================================

def parse_arguments() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Read image URLs from XLSX, classify the images, "
            "append a dated result column to a history CSV, "
            "and create a Google Sheet."
        )
    )

    parser.add_argument(
        "xlsx_file",
        type=Path,
        help="Path to the current XLSX input file",
    )

    parser.add_argument(
        "--history-csv",
        type=Path,
        default=Path("image_results_history.csv"),
        help=(
            "CSV file that stores all historical runs. "
            "Default: image_results_history.csv"
        ),
    )

    parser.add_argument(
        "--sheet-title",
        default=DEFAULT_GOOGLE_SHEET_TITLE,
        help=(
            "Base title for the Google Sheet. "
            f"Default: {DEFAULT_GOOGLE_SHEET_TITLE!r}"
        ),
    )

    parser.add_argument(
        "--json-backup",
        type=Path,
        default=Path("image_results_history_backup.json"),
        help=(
            "Optional JSON backup file. "
            "Default: image_results_history_backup.json"
        ),
    )

    parser.add_argument(
        "--skip-google-sheet",
        action="store_true",
        help=(
            "Update the CSV history without creating "
            "a Google Sheet"
        ),
    )

    return parser.parse_args()


# ============================================================
# Main
# ============================================================

def main() -> None:
    args = parse_arguments()

    print("=" * 70)
    print("Reading current XLSX file")
    print("=" * 70)
    print(f"XLSX: {args.xlsx_file.resolve()}")

    current_xlsx_rows = read_xlsx_rows(
        args.xlsx_file
    )

    if not current_xlsx_rows:
        raise ValueError(
            "The XLSX file does not contain valid image URLs"
        )

    print(
        f"Unique URLs in current XLSX: "
        f"{len(current_xlsx_rows)}"
    )

    print()
    print("=" * 70)
    print("Processing current images")
    print("=" * 70)

    current_results = process_current_rows(
        current_xlsx_rows
    )

    print()
    print("=" * 70)
    print("Reading existing CSV history")
    print("=" * 70)
    print(
        f"History CSV: "
        f"{args.history_csv.resolve()}"
    )

    existing_headers, existing_rows = read_history_csv(
        args.history_csv
    )

    print(
        f"Existing history rows: {len(existing_rows)}"
    )

    print()
    print("=" * 70)
    print("Merging current results into history")
    print("=" * 70)

    (
        merged_headers,
        merged_rows,
        current_run_column,
    ) = merge_results_into_history(
        existing_headers=existing_headers,
        existing_rows=existing_rows,
        current_results=current_results,
    )

    print(
        f"New result column: {current_run_column}"
    )

    print(
        f"Total rows after merge: {len(merged_rows)}"
    )

    save_history_csv(
        csv_path=args.history_csv,
        headers=merged_headers,
        rows=merged_rows,
    )

    print()
    print(
        f"CSV history saved successfully:\n"
        f"{args.history_csv.resolve()}"
    )

    save_json_backup(
        output_path=args.json_backup,
        headers=merged_headers,
        rows=merged_rows,
    )

    print(
        f"JSON backup saved:\n"
        f"{args.json_backup.resolve()}"
    )

    google_sheet_url = ""

    if not args.skip_google_sheet:
        try:
            google_sheet_url = create_google_sheet(
                headers=merged_headers,
                rows=merged_rows,
                latest_result_column=current_run_column,
                sheet_title=args.sheet_title,
            )

        except APIError as error:
            print()
            print(
                "The CSV history was saved, but Google Sheets "
                "returned an error."
            )

            print(
                "Verify that Google Sheets API and Google Drive "
                "API are enabled."
            )

            print(f"Google API error: {error}")

        except Exception as error:
            print()
            print(
                "The CSV history was saved, but the Google "
                "Sheet could not be created."
            )

            print(f"Google Sheet error: {error}")

    print()
    print("=" * 70)
    print("Finished")
    print("=" * 70)

    print(
        f"History CSV:\n"
        f"{args.history_csv.resolve()}"
    )

    print(
        f"Latest result column:\n"
        f"{current_run_column}"
    )

    print(
        f"URLs processed in this run: "
        f"{len(current_results)}"
    )

    print(
        f"Total URLs in history: "
        f"{len(merged_rows)}"
    )

    if google_sheet_url:
        print(
            f"Google Sheet URL:\n"
            f"{google_sheet_url}"
        )


if __name__ == "__main__":
    try:
        main()

    except KeyboardInterrupt:
        print()
        print("The process was stopped by the user.")

    except Exception as error:
        print()
        print("=" * 70)
        print("The process failed")
        print("=" * 70)
        print(f"Error: {error}")