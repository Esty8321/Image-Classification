from __future__ import annotations
import argparse
import concurrent.futures
import csv
import hashlib
import re
import struct
import time
from datetime import datetime
from pathlib import Path
from typing import Any, NamedTuple
from urllib.parse import urlparse

import requests
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry


# ============================================================
# Endpoint configuration
# ============================================================

DEV_SERVER_IP = "10.0.0.206"
DEV_SERVER_PORT = 80
# The DEV server exposes the non-priority variant of the endpoint
# and rejects requests that include the priority field.
DEV_ENDPOINT = "/gpu-server-api/predict-binary32"

PRODUCTION_SERVER_IP = "84.95.87.212"
PRODUCTION_SERVER_PORT = 80
PRODUCTION_ENDPOINT = "/gpu-server-api/predict-binary32-priority"


class ServerConfig(NamedTuple):
    name: str
    ip: str
    port: int
    endpoint: str
    supports_priority: bool


class ServerCheckResult(NamedTuple):
    status_text: str
    elapsed_ms: int


# Each image is checked against every server below, at the same
# time, so the results can be compared side by side in the history
# CSV. Each server has its own endpoint path and priority-field
# support, since DEV and PRODUCTION do not speak the same wire
# format.
SERVERS: list[ServerConfig] = [
    ServerConfig(
        name="DEV",
        ip=DEV_SERVER_IP,
        port=DEV_SERVER_PORT,
        endpoint=DEV_ENDPOINT,
        supports_priority=False,
    ),
    ServerConfig(
        name="PRODUCTION",
        ip=PRODUCTION_SERVER_IP,
        port=PRODUCTION_SERVER_PORT,
        endpoint=PRODUCTION_ENDPOINT,
        supports_priority=True,
    ),
]

PRIORITY = 0

IMAGE_DOWNLOAD_TIMEOUT_SECONDS = 30
ENDPOINT_TIMEOUT_SECONDS = 60

# Maximum downloaded image size: 30 MB
MAX_IMAGE_SIZE_BYTES = 30 * 1024 * 1024




# ============================================================
# XLSX columns
# ============================================================

URL_COLUMN = "B"
EXPECTED_STATUS_COLUMN = "E"
DESCRIPTION_COLUMN = "F"
CATEGORY_COLUMN = "W"


# ============================================================
# CSV column names
# ============================================================

CSV_URL_COLUMN = "כתובת התמונה"
CSV_DESCRIPTION_COLUMN = "תיאור"
CSV_CATEGORY_COLUMN = "קטגוריה"
CSV_EXPECTED_COLUMN = "סטטוס צפוי"

ACTUAL_STATUS_PREFIX = "סטטוס בפועל "



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



def clean_cell_value(value: Any) -> str:
    if value is None:
        return ""

    return str(value).strip()


def extract_url_from_excel_value(value: Any) -> str:
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


def create_run_column_names(
    existing_headers: list[str],
) -> dict[str, str]:
    """
    Create one result column name per server (DEV, PRODUCTION, ...),
    all sharing the same run timestamp so they land next to each
    other in the history CSV and can be compared directly.

    Returns:
        A dict mapping server name -> column name, in SERVERS order.
    """
    timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")

    counter = 2
    suffix = ""

    while True:
        candidate_names = {
            server.name: (
                f"{ACTUAL_STATUS_PREFIX}{server.name} "
                f"{timestamp}{suffix}"
            )
            for server in SERVERS
        }

        if not any(
            name in existing_headers
            for name in candidate_names.values()
        ):
            return candidate_names

        suffix = f" ({counter})"
        counter += 1


def get_server_name_from_column_header(header: str) -> str | None:
    """
    Return the server name (e.g. "DEV", "PRODUCTION") encoded in a
    history run-result column header created by
    create_run_column_names, or None when the header is not one of
    those columns.
    """
    if not header.startswith(ACTUAL_STATUS_PREFIX):
        return None

    remainder = header[len(ACTUAL_STATUS_PREFIX):]
    first_word = remainder.split(" ", 1)[0] if remainder else ""

    server_names = {server.name for server in SERVERS}

    if first_word in server_names:
        return first_word

    return None


def find_first_data_row(worksheet) -> int | None:
    """
    Find the first row that contains at least one real value.

    Returns:
        The row number when data is found.
        None when the worksheet is empty.
    """

    max_row = worksheet.max_row

    if max_row is None:
        return None

    try:
        max_row = int(max_row)
    except (TypeError, ValueError):
        return None

    if max_row < 1:
        return None

    max_column = worksheet.max_column

    if max_column is None:
        return None

    try:
        max_column = int(max_column)
    except (TypeError, ValueError):
        return None

    if max_column < 1:
        return None

    for row_number in range(
        1,
        max_row + 1,
    ):
        row_has_value = False

        for column_number in range(
            1,
            max_column + 1,
        ):
            value = worksheet.cell(
                row=row_number,
                column=column_number,
            ).value

            if value is None:
                continue

            if isinstance(value, str):
                value = value.strip()

            if value != "":
                row_has_value = True
                break

        if row_has_value:
            return row_number

    return None


def read_xlsx_rows(xlsx_path: Path) -> list[dict[str, Any]]:
    """
    Read:
        Column B: image URL
        Column D: expected status
        Column E: description
        Column F: category
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
                "גיליון ה־Excel שנבחר ריק או שאינו מכיל נתונים תקינים."
            )
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

            raw_category = worksheet[
                f"{CATEGORY_COLUMN}{row_number}"
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
                    "category": clean_cell_value(
                        raw_category
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
    priority: int | None,
) -> tuple[bytes, str, str]:
    """
    priority=None omits the priority field entirely, for servers
    (e.g. DEV) that do not accept it.
    """
    if (
        priority is not None
        and not 0 <= priority <= 0xFFFFFFFF
    ):
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

    image_size_bytes = struct.pack("!I", len(image_bytes))

    parts = [
        null_terminated(page_url),
        null_terminated(referer),
        null_terminated(key),
        null_terminated(timestamp),
    ]

    if priority is not None:
        parts.append(struct.pack("!I", priority))

    parts.append(image_size_bytes)
    parts.append(image_bytes)

    body = b"".join(parts)

    return body, timestamp, key

def extract_binary_endpoint_status(result: Any) -> int:
    """
    Extract status from the existing binary endpoint.
    """
    if not isinstance(result, dict):
        raise ValueError(
            "The binary endpoint response must be a JSON object"
        )

    if "status" not in result:
        raise ValueError(
            "The binary endpoint JSON does not contain a status field"
        )

    raw_status = result["status"]

    try:
        status = int(raw_status)

    except (TypeError, ValueError) as error:
        raise ValueError(
            f"Invalid binary endpoint status value: {raw_status!r}"
        ) from error

    if status not in {0, -1}:
        raise ValueError(
            f"Unexpected binary endpoint status: {status}. "
            "Expected 0 or -1."
        )

    return status


def endpoint_status_to_hebrew(status: int) -> str:
    if status == 0:
        return "פתוח"

    if status == -1:
        return "חסום"

    raise ValueError(f"Unsupported status: {status}")


def send_image_to_binary_endpoint(
    session: requests.Session,
    image_bytes: bytes,
    image_url: str,
    server_ip: str,
    server_port: int,
    endpoint: str,
    priority: int | None,
) -> int:
    referer = build_referer(image_url)

    body, timestamp, key = create_request_body(
        image_bytes=image_bytes,
        page_url=image_url,
        referer=referer,
        priority=priority,
    )

    request_url = (
        f"http://{server_ip}:{server_port}{endpoint}"
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
            "The binary endpoint returned an empty response"
        )

    try:
        result = response.json()

    except requests.exceptions.JSONDecodeError as error:
        response_preview = response.text[:500]

        raise ValueError(
            "The binary endpoint response is not valid JSON. "
            f"Response: {response_preview}"
        ) from error

    status = extract_binary_endpoint_status(result)

    print(
        f"    Binary endpoint timestamp={timestamp}, "
        f"key={key}, status={status}"
    )

    return status


def check_one_server(
    session: requests.Session,
    image_bytes: bytes,
    image_url: str,
    server: ServerConfig,
) -> ServerCheckResult:
    """
    Check the image against a single server and return its Hebrew
    status (or a Hebrew error message) together with how long the
    API call took, in milliseconds.
    """
    start_time = time.perf_counter()

    def elapsed_ms() -> int:
        return round((time.perf_counter() - start_time) * 1000)

    try:
        binary_status = send_image_to_binary_endpoint(
            session=session,
            image_bytes=image_bytes,
            image_url=image_url,
            server_ip=server.ip,
            server_port=server.port,
            endpoint=server.endpoint,
            priority=(
                PRIORITY if server.supports_priority else None
            ),
        )

        status_text = endpoint_status_to_hebrew(binary_status)
        duration_ms = elapsed_ms()

        print(
            f"    [{server.name}] binary={binary_status} "
            f"({status_text}) — {duration_ms} ms"
        )

        return ServerCheckResult(status_text, duration_ms)

    except requests.exceptions.ConnectTimeout as error:
        duration_ms = elapsed_ms()
        status_text = f"שגיאה: חריגת זמן בחיבור — {error}"
        print(f"    [{server.name}] {status_text} — {duration_ms} ms")
        return ServerCheckResult(status_text, duration_ms)

    except requests.exceptions.ReadTimeout as error:
        duration_ms = elapsed_ms()
        status_text = f"שגיאה: חריגת זמן בתגובה — {error}"
        print(f"    [{server.name}] {status_text} — {duration_ms} ms")
        return ServerCheckResult(status_text, duration_ms)

    except requests.exceptions.HTTPError as error:
        duration_ms = elapsed_ms()
        status_code = (
            error.response.status_code
            if error.response is not None
            else "unknown"
        )

        status_text = f"שגיאה: HTTP {status_code}"
        print(
            f"    [{server.name}] {status_text} — "
            f"{duration_ms} ms: {error}"
        )
        return ServerCheckResult(status_text, duration_ms)

    except requests.exceptions.RequestException as error:
        duration_ms = elapsed_ms()
        status_text = f"שגיאה: בקשת רשת נכשלה — {error}"
        print(f"    [{server.name}] {status_text} — {duration_ms} ms")
        return ServerCheckResult(status_text, duration_ms)

    except Exception as error:
        duration_ms = elapsed_ms()
        status_text = f"שגיאה: {error}"
        print(f"    [{server.name}] {status_text} — {duration_ms} ms")
        return ServerCheckResult(status_text, duration_ms)


def check_image_against_servers(
    session: requests.Session,
    image_bytes: bytes,
    image_url: str,
) -> dict[str, ServerCheckResult]:
    """
    Check the image against every server in SERVERS (DEV,
    PRODUCTION, ...) at the same time, so a slow or unresponsive
    server does not delay the others, and return the Hebrew status
    and response time (ms) per server name.
    """
    with concurrent.futures.ThreadPoolExecutor(
        max_workers=len(SERVERS)
    ) as executor:
        futures = {
            executor.submit(
                check_one_server,
                session=session,
                image_bytes=image_bytes,
                image_url=image_url,
                server=server,
            ): server.name
            for server in SERVERS
        }

        results: dict[str, ServerCheckResult] = {
            futures[future]: future.result()
            for future in concurrent.futures.as_completed(futures)
        }

    return results

def process_current_rows(
    rows: list[dict[str, Any]],
) -> tuple[dict[str, dict[str, str]], dict[str, int]]:
    """
    Process every row and return:
        results: per-URL status, keyed by url_key.
        total_elapsed_ms_by_server: sum of API response time (ms)
            across every image actually sent to each server.
    """
    session = create_http_session()

    results: dict[str, dict[str, str]] = {}
    total_elapsed_ms_by_server: dict[str, int] = {
        server.name: 0 for server in SERVERS
    }

    total = len(rows)

    for index, row in enumerate(rows, start=1):
        source_row = row["source_row"]
        image_url = row["url"]
        url_key = row["url_key"]
        description = row["description"]
        category = row["category"]
        expected_status = row["expected_status"]

        print()
        print(
            f"[{index}/{total}] Processing XLSX row "
            f"{source_row}"
        )
        print(f"    URL: {image_url}")
        print(f"    Expected: {expected_status}")

        actual_status_by_server: dict[str, str]

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

            server_results = check_image_against_servers(
                session=session,
                image_bytes=image_bytes,
                image_url=final_url,
            )

            actual_status_by_server = {
                server_name: result.status_text
                for server_name, result in server_results.items()
            }

            for server_name, result in server_results.items():
                total_elapsed_ms_by_server[server_name] += (
                    result.elapsed_ms
                )

                match_note = (
                    "match"
                    if expected_status == result.status_text
                    else "does not match"
                )
                print(
                    f"    [{server_name}] Actual: "
                    f"{result.status_text} "
                    f"({result.elapsed_ms} ms) — {match_note}"
                )

        except requests.exceptions.ConnectTimeout as error:
            download_error = (
                f"שגיאה: חריגת זמן בחיבור — {error}"
            )
            print(f"    {download_error}")
            actual_status_by_server = {
                server.name: download_error
                for server in SERVERS
            }

        except requests.exceptions.ReadTimeout as error:
            download_error = (
                f"שגיאה: חריגת זמן בתגובה — {error}"
            )
            print(f"    {download_error}")
            actual_status_by_server = {
                server.name: download_error
                for server in SERVERS
            }

        except requests.exceptions.HTTPError as error:
            status_code = (
                error.response.status_code
                if error.response is not None
                else "unknown"
            )

            download_error = f"שגיאה: HTTP {status_code}"
            print(f"    {download_error}: {error}")
            actual_status_by_server = {
                server.name: download_error
                for server in SERVERS
            }

        except requests.exceptions.RequestException as error:
            download_error = (
                f"שגיאה: בקשת רשת נכשלה — {error}"
            )
            print(f"    {download_error}")
            actual_status_by_server = {
                server.name: download_error
                for server in SERVERS
            }

        except Exception as error:
            download_error = f"שגיאה: {error}"
            print(f"    {download_error}")
            actual_status_by_server = {
                server.name: download_error
                for server in SERVERS
            }

        results[url_key] = {
            CSV_URL_COLUMN: image_url,
            CSV_DESCRIPTION_COLUMN: description,
            CSV_CATEGORY_COLUMN: category,
            CSV_EXPECTED_COLUMN: expected_status,
            "actual_status_by_server": actual_status_by_server,
        }

    return results, total_elapsed_ms_by_server


# ============================================================
# CSV history
# ============================================================

def get_base_headers() -> list[str]:
    return [
        CSV_URL_COLUMN,
        CSV_DESCRIPTION_COLUMN,
        CSV_CATEGORY_COLUMN,
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
) -> tuple[list[str], list[dict[str, str]], dict[str, str]]:

    headers = list(existing_headers)

    for base_header in get_base_headers():
        if base_header not in headers:
            headers.insert(
                get_base_headers().index(base_header),
                base_header,
            )

    current_run_columns = create_run_column_names(headers)

    for server in SERVERS:
        headers.append(current_run_columns[server.name])

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

        # The new run columns are blank by default.
        for server in SERVERS:
            normalized_row[current_run_columns[server.name]] = ""

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

            history_row[CSV_CATEGORY_COLUMN] = current_result[
                CSV_CATEGORY_COLUMN
            ]

            history_row[CSV_EXPECTED_COLUMN] = current_result[
                CSV_EXPECTED_COLUMN
            ]

            for server in SERVERS:
                history_row[
                    current_run_columns[server.name]
                ] = current_result["actual_status_by_server"][
                    server.name
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

            new_row[CSV_CATEGORY_COLUMN] = current_result[
                CSV_CATEGORY_COLUMN
            ]

            new_row[CSV_EXPECTED_COLUMN] = current_result[
                CSV_EXPECTED_COLUMN
            ]

            for server in SERVERS:
                new_row[
                    current_run_columns[server.name]
                ] = current_result["actual_status_by_server"][
                    server.name
                ]

            history_by_key[url_key] = new_row
            ordered_keys.append(url_key)

    merged_rows = [
        history_by_key[url_key]
        for url_key in ordered_keys
    ]

    return headers, merged_rows, current_run_columns


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

    current_results, total_elapsed_ms_by_server = (
        process_current_rows(
            current_xlsx_rows
        )
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
        current_run_columns,
    ) = merge_results_into_history(
        existing_headers=existing_headers,
        existing_rows=existing_rows,
        current_results=current_results,
    )

    for server_name, column_name in current_run_columns.items():
        print(
            f"New result column [{server_name}]: {column_name}"
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
    print()
    print("=" * 70)
    print("Finished")
    print("=" * 70)

    print(
        f"History CSV:\n"
        f"{args.history_csv.resolve()}"
    )

    for server_name, column_name in current_run_columns.items():
        print(
            f"Latest result column [{server_name}]:\n"
            f"{column_name}"
        )

    print(
        f"URLs processed in this run: "
        f"{len(current_results)}"
    )

    print(
        f"Total URLs in history: "
        f"{len(merged_rows)}"
    )

    print()
    print("Total API response time per server:")

    for server_name, total_ms in (
        total_elapsed_ms_by_server.items()
    ):
        print(f"  [{server_name}] {total_ms} ms")

  
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