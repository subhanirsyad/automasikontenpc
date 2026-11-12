"""
SC PC VIDEO - EXPLODED + BRAND UNIQUE + TAP/AFF PRODUCT ENRICHMENT
==================================================================

Input Google Sheets:
    1) RAW SC Video
    2) RAW TAP Short Video
    3) Satu atau lebih RAW Affiliate Order

Output Google Sheets:
    - GABUNGAN SC - ALL
    - GABUNGAN SC - RANGE
    - optional: <BRAND> - SC ALL
    - optional: <BRAND> - SC RANGE
    - BRAND SUMMARY
    - MISSING MAPPING

Flow:
    TAP RAW -> aggregate Video ID + Product ID
    AFF RAW -> aggregate Content ID + Product ID
    TAP + AFF -> MASTER MAPPING dengan key Video ID + Product ID
    SC RAW -> explode semua Product ID -> enrich metric per Product ID
    EXPLODED -> group Video ID + Brand -> FIRST PID per Brand + SUM metric unique PID

Rule metric commerce setelah explode:
    - Jika row SC hanya punya 1 Product ID:
        pertahankan GMV / Orders / Items dari SC original, walaupun pair ada di TAP/AFF.
        AOV dihitung ulang = SC GMV / SC Orders.
    - Jika row SC punya >1 Product ID:
        baru gunakan MASTER MAPPING Video ID + Product ID untuk mengisi:
            Creator video-attributed GMV
            Video-attributed orders
            Video-attributed items sold
            AOV = mapped GMV / mapped Orders

Prioritas mapping untuk row multi-PID:
    TAP jika pasangan Video ID + Product ID ada di TAP.
    Jika tidak ada di TAP tetapi ada di Affiliate Order, pakai Affiliate Order.

Kolom Affiliate Order tetap di-lookup untuk semua row (single maupun multi PID):
    - GMV Affiliate Order (Est. Commission Base)
    - Jumlah Order Affiliate

Product name + Brand:
    Product name ditulis ke output setelah Product ID.
    Brand PRIMARY diturunkan dari Product name menggunakan rule yang sama seperti TAP PC.
    Product name dicari dari exact pair Video ID + Product ID pada TAP/AFF,
    lalu fallback Product ID -> Product name global.
    Jika hasil Brand dari Product name = MISSING dan USE_PRODUCT_ID_MAPPING_FALLBACK=True,
    Brand diambil dari mapping Product ID data PC.

ALL/RANGE:
    ALL   = seluruh hasil SC setelah explode + enrichment.
    RANGE = ALL yang difilter berdasarkan SC Post date.

SC NOT IN TAP:
    Perbandingan keberadaan Video ID SC vs TAP memakai TAP.Post time pada periode
    METRIC_START_DATE s.d. METRIC_END_DATE yang sama dengan periode mapping.
    Jadi jika METRIC = 2026-10-01 s.d. 2026-10-06, Video ID TAP di luar
    1-6 Oktober tidak dianggap ada untuk pengecekan SC NOT IN TAP.

Mode:
    python clean_sc_video_enriched_gsheet.py
        -> DRY RUN

    python sc_pc_video_exploded_and_brand_unique.py --write
        -> tulis output

    python sc_pc_video_exploded_and_brand_unique.py --write --replace
        -> ganti tab output yang sudah ada

Optional override brand tabs:
    --brand-tabs
    --no-brand-tabs

Dependency:
    pip install google-api-python-client google-auth
"""

from __future__ import annotations

import math
import re
import sys
import time
from collections import Counter, defaultdict
from datetime import date, datetime
from pathlib import Path


# =============================================================================
# CONFIG - EDIT BAGIAN INI
# =============================================================================

BASE_DIR = Path(__file__).resolve().parent

# Script lama di project ini umumnya memakai credentials.json.
# credentials_sc.json tetap didukung sebagai fallback supaya kompatibel.
CREDENTIALS_CANDIDATES = [
    BASE_DIR / "credentials.json",
    BASE_DIR / "credentials_sc.json",
]
CREDENTIALS_FILE = next(
    (path for path in CREDENTIALS_CANDIDATES if path.exists()),
    CREDENTIALS_CANDIDATES[0],
)

# -----------------------------------------------------------------------------
# RAW SC VIDEO
# -----------------------------------------------------------------------------
SC_SPREADSHEET_ID = "1XCrfQWU5Db-qxF6tk0NKjTeZLq5UstyqXqZcqVyIC38"
SC_SHEET_GID = None
SC_SHEET_NAME = None

# -----------------------------------------------------------------------------
# RAW TAP SHORT VIDEO
# -----------------------------------------------------------------------------
TAP_SPREADSHEET_ID = "1yi2WIokaM56Lh1LKqb2NOHXCCCXbekLhEIr2eBEi0b8"
TAP_SHEET_GID = None
TAP_SHEET_NAME = "Custom report"

# -----------------------------------------------------------------------------
# RAW AFFILIATE ORDER
# Bisa satu atau lebih source.
# Jika sheet_name diisi, sheet_gid boleh None.
# Jika sheet_gid diisi, sheet_name boleh None.
# Jika keduanya None, script auto-detect tab berdasarkan required headers.
# -----------------------------------------------------------------------------
AFF_SOURCES = [
    {
        "spreadsheet_id": "1DnLIBPf9Ufgn6Q0-SYz__9m5-wwptki-g_9iGFWbkFg",
        "sheet_gid": None,
        "sheet_name": None,
        "label": "AFF 1",
    },
    # Contoh source kedua:
    # {
    #     "spreadsheet_id": "1DnLIBPf9Ufgn6Q0-SYz__9m5-wwptki-g_9iGFWbkFg",
    #     "sheet_gid": None,
    #     "sheet_name": None,
    #     "label": "AFF 2",
    # },
]

# -----------------------------------------------------------------------------
# OUTPUT
# -----------------------------------------------------------------------------
OUTPUT_SPREADSHEET_ID = "ISI_BACKEND_DISABLED_NO_DIRECT_WRITE"
# Optional Product ID -> Brand fallback (PC mapping).
# Product name tetap PRIMARY. Mapping Product ID hanya dipakai jika
# Product name tidak berhasil menghasilkan Brand PC.
USE_PRODUCT_ID_MAPPING_FALLBACK = True
MAPPING_SPREADSHEET_ID = "1O4dQVFYawbirEN5wpV0ygjNZtNFZPZH0tLX_Ov34fNA"
MAPPING_SHEET_GID = 437523491
MAPPING_PRODUCT_ID_HEADER = "Product ID TTS"
MAPPING_BRAND_HEADER = "BRAND FIX"

# -----------------------------------------------------------------------------
# DATE RANGE
# METRIC_* memfilter source mapping TAP.Post time dan AFF.Time Created.
# Isi None jika semua row source mapping ingin dipakai.
#
# RANGE_* hanya memfilter hasil SC untuk GABUNGAN SC - RANGE dan brand RANGE.
# ALL tetap berisi seluruh SC hasil explode + enrichment.
# -----------------------------------------------------------------------------
METRIC_START_DATE = "2026-10-07"       # contoh: "2026-10-01"
METRIC_END_DATE = "2026-10-07"         # contoh: "2026-10-07"
RANGE_START_DATE = "2026-10-07"     # contoh: "2026-10-01"
RANGE_END_DATE = "2026-10-07"          # contoh: "2026-10-06"

# -----------------------------------------------------------------------------
# OUTPUT OPTION
# Bisa diubah langsung di sini.
# Juga bisa dioverride saat run dengan --brand-tabs / --no-brand-tabs.
# -----------------------------------------------------------------------------
WRITE_BRAND_TABS = False

# -----------------------------------------------------------------------------
# GOOGLE SHEETS / SAFETY
# -----------------------------------------------------------------------------
SOURCE_SCAN_RANGE = "A1:AZ20"
WRITE_CHUNK_ROWS = 1500
MAX_CELLS = 20_000_000
MAX_RETRIES = 7
RETRYABLE_HTTP_STATUS = {429, 500, 502, 503, 504}

# Google Sheets API quota guard.
# Default per-user quota commonly separates READ and WRITE requests.
# 1.10s/request keeps each lane below ~55 requests/minute.
MIN_READ_INTERVAL_SECONDS = 1.10
MIN_WRITE_INTERVAL_SECONDS = 1.10

# Kalau 429 tetap terjadi (mis. script lain baru saja memakai quota yang sama),
# tunggu satu quota window penuh sebelum retry.
QUOTA_429_WAIT_SECONDS = 65.0

# -----------------------------------------------------------------------------
# OUTPUT TAB NAMES
# -----------------------------------------------------------------------------
SHEET_ALL = "GABUNGAN SC PC SV EXPLODED - ALL"
SHEET_RANGE = "GABUNGAN SC PC SV EXPLODED - RANGE"
SC_NOT_IN_TAP_ALL_SHEET = "SC PC SV EXPLODED NOT IN TAP - ALL"
SC_NOT_IN_TAP_RANGE_SHEET = "SC PC SV EXPLODED NOT IN TAP - RANGE"
BRAND_UNIQUE_ALL_SHEET = "GABUNGAN SC PC SV BRAND UNIQUE - ALL"
BRAND_UNIQUE_RANGE_SHEET = "GABUNGAN SC PC SV BRAND UNIQUE - RANGE"
BRAND_UNIQUE_NOT_IN_TAP_ALL_SHEET = "SC PC SV BRAND UNIQUE NOT IN TAP - ALL"
BRAND_UNIQUE_NOT_IN_TAP_RANGE_SHEET = "SC PC SV BRAND UNIQUE NOT IN TAP - RANGE"
SUMMARY_SHEET = "PC BRAND SUMMARY SV"
MISSING_SHEET = "PC SV MISSING MAPPING"
BRAND_ALL_SUFFIX = " - SC PC SV ALL"
BRAND_RANGE_SUFFIX = " - SC PC SV RANGE"

CLEANING_VERSION = "v8.1-quota-safe-model-a-tap-posttime-aff-timecreated"


# =============================================================================
# MODE
# =============================================================================

WRITE = "--write" in sys.argv
REPLACE = "--replace" in sys.argv

if REPLACE and not WRITE:
    raise SystemExit("--replace hanya boleh dipakai bersama --write")

if "--brand-tabs" in sys.argv and "--no-brand-tabs" in sys.argv:
    raise SystemExit("Pilih salah satu: --brand-tabs atau --no-brand-tabs")

if "--brand-tabs" in sys.argv:
    WRITE_BRAND_TABS_EFFECTIVE = True
elif "--no-brand-tabs" in sys.argv:
    WRITE_BRAND_TABS_EFFECTIVE = False
else:
    WRITE_BRAND_TABS_EFFECTIVE = WRITE_BRAND_TABS


# =============================================================================
# OUTPUT SCHEMA - PRODUCT NAME + BRAND INCLUDED
# =============================================================================

OUTPUT_COLUMNS = [
    "Video title",
    "Video ID",
    "Post date",
    "Video link",
    "Creator name",
    "Product ID",
    "Product name",
    "Brand",
    "Creator video-attributed GMV",
    "Video-attributed orders",
    "AOV",
    "Video-attributed items sold",
    "Refunds",
    "Items refunded",
    "Likes",
    "Comments",
    "Shares",
    "Video product impressions",
    "Video product clicks",
    "Completion rate",
    "Video views",
    "CTR",
    "Video GPM",
    "Engagement",
    "Avg. GMV per customer",
    "Est. commission",
    "GMV Affiliate Order (Est. Commission Base)",
    "Jumlah Order Affiliate",
]

SC_SOURCE_ALIASES = {
    "CTR": ["CTR", "Avg. daily video product CTR"],
    "Video GPM": ["Video GPM", "Avg. daily GPM"],
    "Engagement": ["Engagement", "Avg. daily engagement"],
}

SC_REQUIRED = [
    "Video title",
    "Video ID",
    "Post date",
    "Product ID",
    "Creator video-attributed GMV",
    "Video-attributed orders",
    "Video-attributed items sold",
]

TAP_REQUIRED = [
    "Video ID",
    "Product ID",
    "Post time",
    "Affiliate video-attributed GMV",
    "Creator video-attributed orders",
    "Creator-attributed items sold",
]

AFF_REQUIRED = [
    "Order ID",
    "Product ID",
    "Content Type",
    "Content ID",
    "Est. Commission Base",
    "Quantity",
]


# =============================================================================
# BRAND RULES - SAME PARENT BRAND LOGIC AS TAP PC
# =============================================================================

PC_BRAND_ORDER = [
    "CLOSE UP",
    "DOVE DEO",
    "DOVE SCL",
    "LIFEBUOY",
    "LUX",
    "PEPSODENT",
    "REXONA",
]
MISSING_BRAND = "MISSING"
BRAND_OUTPUT_ORDER = PC_BRAND_ORDER + [MISSING_BRAND]

MAPPING_BRAND_ALIASES = {
    "CLOSEUP": "CLOSE UP",
    "CLOSE-UP": "CLOSE UP",
    "REXONA MEN": "REXONA",
    "REXONA MAN": "REXONA",
    "DOVE DEODORANT": "DOVE DEO",
    "DOVE DEO": "DOVE DEO",
    "DOVE SCL": "DOVE SCL",
    "DOVE": "DOVE SCL",
}

BASE_BRAND_PATTERNS = {
    "CLOSE UP": re.compile(r"(?<![A-Z0-9])CLOSE\s*-?\s*UP(?![A-Z0-9])", re.I),
    "DOVE": re.compile(r"(?<![A-Z0-9])DOVE(?![A-Z0-9])", re.I),
    "LIFEBUOY": re.compile(r"(?<![A-Z0-9])LIFEBUOY(?![A-Z0-9])", re.I),
    "LUX": re.compile(r"(?<![A-Z0-9])LUX(?![A-Z0-9])", re.I),
    "PEPSODENT": re.compile(r"(?<![A-Z0-9])PEPSODENT(?![A-Z0-9])", re.I),
    "REXONA": re.compile(r"(?<![A-Z0-9])REXONA(?![A-Z0-9])", re.I),
}

DOVE_DEO_RE = re.compile(
    r"(?<![A-Z0-9])(?:DEO|DEODORANT)(?![A-Z0-9])",
    re.I,
)


# =============================================================================
# GENERAL HELPERS
# =============================================================================

def normalize_header(value: object) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(value or "").strip().lower())


def normalize_id(value: object) -> str:
    if value is None:
        return ""

    s = str(value).strip()
    if not s or s.lower() in {"nan", "none", "null", "--", "-"}:
        return ""

    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]

    if re.fullmatch(r"[+-]?\d+(?:\.\d+)?[eE][+-]?\d+", s):
        raise ValueError(
            f"ID dalam scientific notation tidak aman: {s!r}. "
            "Pastikan source Google Sheets menyimpan ID sebagai text / digit utuh."
        )

    return s


def split_product_ids(value: object) -> list[str]:
    s = str(value or "").strip()
    if not s:
        return []

    result = []
    seen = set()

    for piece in s.split(","):
        pid = normalize_id(piece)
        if pid and pid not in seen:
            result.append(pid)
            seen.add(pid)

    return result


def parse_number(value: object) -> float:
    if value is None:
        return 0.0
    if isinstance(value, bool):
        return 0.0
    if isinstance(value, (int, float)):
        if isinstance(value, float) and math.isnan(value):
            return 0.0
        return float(value)

    s = str(value).strip()
    if not s or s.upper() in {"-", "--", "N/A", "NA", "NULL", "NONE"}:
        return 0.0

    s = re.sub(r"^rp\s*", "", s, flags=re.I)
    s = s.replace(" ", "")

    negative = False
    if s.startswith("(") and s.endswith(")"):
        negative = True
        s = s[1:-1]

    if re.fullmatch(r"[+-]?\d{1,3}(?:\.\d{3})+", s):
        s = s.replace(".", "")
    elif re.fullmatch(r"[+-]?\d{1,3}(?:,\d{3})+", s):
        s = s.replace(",", "")
    elif re.fullmatch(r"[+-]?\d{1,3}(?:\.\d{3})*,\d+", s):
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", "")

    try:
        out = float(s)
    except ValueError as exc:
        raise ValueError(f"Angka tidak valid: {value!r}") from exc

    return -out if negative else out


def parse_int(value: object) -> int:
    return int(round(parse_number(value)))


def parse_percent(value: object) -> float:
    if value is None:
        return 0.0
    s = str(value).strip()
    if not s:
        return 0.0
    if s.endswith("%"):
        return parse_number(s[:-1]) / 100.0
    return parse_number(s)


def clean_numeric(value: float) -> int | float:
    if math.isclose(float(value), round(float(value)), rel_tol=0, abs_tol=1e-9):
        return int(round(float(value)))
    return round(float(value), 6)


def parse_iso_date(value: str | None) -> date | None:
    if value is None or str(value).strip() == "":
        return None
    return datetime.strptime(str(value).strip(), "%Y-%m-%d").date()


def parse_tap_date(value: object) -> date | None:
    s = str(value or "").strip()
    if not s or s in {"-", "--"}:
        return None

    for fmt in (
        "%Y-%m-%d",
        "%Y/%m/%d",
        "%m/%d/%Y",
        "%d/%m/%Y",
    ):
        try:
            return datetime.strptime(s[:10], fmt).date()
        except ValueError:
            pass

    raise ValueError(f"TAP Date tidak valid: {value!r}")


def parse_tap_post_time(value: object) -> datetime | None:
    """
    Parse TAP Short Video `Post time`.

    Format utama export TAP:
        M/D/YYYY HH:MM:SS
        MM/DD/YYYY HH:MM:SS

    Juga menerima variasi tanpa detik / date-only / ISO.
    """
    s = str(value or "").strip()
    if not s or s in {"-", "--"}:
        return None

    for fmt in (
        "%m/%d/%Y %H:%M:%S",
        "%m/%d/%Y %H:%M",
        "%m/%d/%Y",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%Y-%m-%d",
        "%Y/%m/%d %H:%M:%S",
        "%Y/%m/%d %H:%M",
        "%Y/%m/%d",
    ):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            pass

    raise ValueError(f"TAP Post time tidak valid: {value!r}")


def parse_aff_datetime(value: object) -> datetime | None:
    # Affiliate export yang diberikan memakai DD/MM/YYYY HH:MM:SS.
    s = str(value or "").strip()
    if not s or s in {"-", "--"}:
        return None

    for fmt in (
        "%d/%m/%Y %H:%M:%S",
        "%d/%m/%Y %H:%M",
        "%d/%m/%Y",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%Y-%m-%d",
    ):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            pass

    raise ValueError(f"AFF Time Created tidak valid: {value!r}")


def normalize_sc_post_date(value: object) -> str:
    # Seller Center Video memakai MM/DD/YYYY.
    s = str(value or "").strip()
    if not s or s in {"-", "--"}:
        return ""

    for fmt in (
        "%m/%d/%Y %H:%M:%S",
        "%m/%d/%Y %H:%M",
        "%m/%d/%Y",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%Y-%m-%d",
    ):
        try:
            dt = datetime.strptime(s, fmt)
            return dt.strftime("%Y-%m-%d %H:%M:%S")
        except ValueError:
            pass

    return s


def parse_output_post_date(value: object) -> date | None:
    s = str(value or "").strip()
    if not s or s in {"-", "--"}:
        return None

    for fmt in (
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%Y-%m-%d",
        "%m/%d/%Y %H:%M:%S",
        "%m/%d/%Y %H:%M",
        "%m/%d/%Y",
    ):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass

    return None


def in_period(d: date | None, start: date | None, end: date | None) -> bool:
    if d is None:
        return False if (start or end) else True
    if start and d < start:
        return False
    if end and d > end:
        return False
    return True


def resolve_header(header: list[str], wanted: str) -> str:
    target = normalize_header(wanted)
    for name in header:
        if normalize_header(name) == target:
            return name
    raise KeyError(wanted)


def resolve_first_header(header: list[str], candidates: list[str]) -> str | None:
    normalized = {normalize_header(name): name for name in header}
    for candidate in candidates:
        found = normalized.get(normalize_header(candidate))
        if found is not None:
            return found
    return None


def validate_headers(header: list[str], required: list[str], label: str) -> None:
    missing = []
    for wanted in required:
        try:
            resolve_header(header, wanted)
        except KeyError:
            missing.append(wanted)

    if missing:
        raise ValueError(
            f"{label}: kolom wajib tidak ditemukan: {', '.join(missing)}\n"
            f"Header tersedia: {header}"
        )


def column_letter(index_1_based: int) -> str:
    result = ""
    n = index_1_based
    while n:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result or "A"


def trim_header(row: list[object]) -> list[str]:
    values = [str(v or "").strip() for v in row]
    last = -1
    for i, value in enumerate(values):
        if value:
            last = i
    return values[: last + 1] if last >= 0 else []


def rows_to_dicts(header: list[str], rows: list[list[object]]) -> list[dict[str, object]]:
    result = []
    width = len(header)
    for raw in rows:
        padded = list(raw[:width]) + [""] * max(0, width - len(raw))
        result.append({header[i]: padded[i] for i in range(width)})
    return result


# =============================================================================
# GOOGLE SHEETS
# =============================================================================

def get_sheets_service():
    try:
        from google.oauth2 import service_account
        from googleapiclient.discovery import build
    except ImportError as exc:
        raise RuntimeError(
            "Google library belum terinstall. Jalankan:\n"
            "  pip install google-api-python-client google-auth"
        ) from exc

    scope = (
        "https://www.googleapis.com/auth/spreadsheets"
        if WRITE
        else "https://www.googleapis.com/auth/spreadsheets.readonly"
    )

    credentials = service_account.Credentials.from_service_account_file(
        str(CREDENTIALS_FILE),
        scopes=[scope],
    )

    return build(
        "sheets",
        "v4",
        credentials=credentials,
        cache_discovery=False,
    )


_SHEET_PROPERTIES_CACHE: dict[str, list[dict[str, object]]] = {}
_LAST_READ_REQUEST_MONOTONIC = 0.0
_LAST_WRITE_REQUEST_MONOTONIC = 0.0


def _request_lane(request) -> str:
    """Classify Google API request into read/write quota lane."""
    method = str(getattr(request, "method", "") or "").upper()
    return "READ" if method == "GET" else "WRITE"


def _pace_google_request(request) -> None:
    """Throttle reads/writes independently to stay below per-user quota."""
    global _LAST_READ_REQUEST_MONOTONIC
    global _LAST_WRITE_REQUEST_MONOTONIC

    lane = _request_lane(request)
    now = time.monotonic()

    if lane == "READ":
        last = _LAST_READ_REQUEST_MONOTONIC
        interval = MIN_READ_INTERVAL_SECONDS
    else:
        last = _LAST_WRITE_REQUEST_MONOTONIC
        interval = MIN_WRITE_INTERVAL_SECONDS

    if last:
        elapsed = now - last
        if elapsed < interval:
            time.sleep(interval - elapsed)

    now = time.monotonic()
    if lane == "READ":
        _LAST_READ_REQUEST_MONOTONIC = now
    else:
        _LAST_WRITE_REQUEST_MONOTONIC = now


def execute_with_retry(request):
    """Execute Google request with quota-aware pacing and retry."""
    for attempt in range(MAX_RETRIES):
        _pace_google_request(request)

        try:
            return request.execute()

        except Exception as exc:
            status = getattr(getattr(exc, "resp", None), "status", None)

            if status not in RETRYABLE_HTTP_STATUS:
                raise

            if attempt == MAX_RETRIES - 1:
                raise

            if status == 429:
                retry_after = None
                resp = getattr(exc, "resp", None)
                if resp is not None:
                    try:
                        retry_after = resp.get("retry-after")
                    except Exception:
                        retry_after = None

                try:
                    wait_seconds = (
                        max(float(retry_after), QUOTA_429_WAIT_SECONDS)
                        if retry_after
                        else QUOTA_429_WAIT_SECONDS
                    )
                except (TypeError, ValueError):
                    wait_seconds = QUOTA_429_WAIT_SECONDS

                print(
                    "Google API 429 (quota); "
                    f"menunggu {wait_seconds:.0f}s lalu retry "
                    f"{attempt + 1}/{MAX_RETRIES}...",
                    flush=True,
                )
            else:
                wait_seconds = min(2 ** attempt, 16)
                print(
                    f"Google API error {status}; "
                    f"retry {attempt + 1}/{MAX_RETRIES} "
                    f"dalam {wait_seconds}s...",
                    flush=True,
                )

            time.sleep(wait_seconds)


def get_sheet_properties(
    service,
    spreadsheet_id: str,
    *,
    refresh: bool = False,
):
    """Get spreadsheet metadata with an in-process cache."""
    if not refresh and spreadsheet_id in _SHEET_PROPERTIES_CACHE:
        return _SHEET_PROPERTIES_CACHE[spreadsheet_id]

    response = execute_with_retry(
        service.spreadsheets().get(
            spreadsheetId=spreadsheet_id,
            fields=(
                "sheets(properties("
                "sheetId,title,"
                "gridProperties(rowCount,columnCount)"
                "))"
            ),
        )
    )

    props = [
        sheet["properties"]
        for sheet in response.get("sheets", [])
    ]
    _SHEET_PROPERTIES_CACHE[spreadsheet_id] = props
    return props


def invalidate_sheet_properties_cache(spreadsheet_id: str) -> None:
    _SHEET_PROPERTIES_CACHE.pop(spreadsheet_id, None)


def detect_and_read_source(
    service,
    *,
    spreadsheet_id: str,
    sheet_gid: int | None,
    sheet_name: str | None,
    required_headers: list[str],
    label: str,
):
    """Detect source tab and read it with minimal Google READ requests.

    If sheet_name or sheet_gid points to exactly one tab:
        metadata = cached
        data     = one values.get request

    If tab must be auto-detected:
        small A1:AZ20 scan per candidate, then one full read for the match.
    """
    properties = get_sheet_properties(service, spreadsheet_id)
    if not properties:
        raise ValueError(f"{label}: spreadsheet tidak memiliki sheet.")

    titles = [p["title"] for p in properties]

    if sheet_name is not None:
        if sheet_name not in titles:
            raise ValueError(
                f'{label}: sheet_name "{sheet_name}" tidak ditemukan. '
                f"Available: {titles}"
            )
        candidates = [
            p for p in properties
            if p["title"] == sheet_name
        ]

    elif sheet_gid is not None:
        candidates = [
            p for p in properties
            if int(p["sheetId"]) == int(sheet_gid)
        ]
        if not candidates:
            available = [
                (p["sheetId"], p["title"])
                for p in properties
            ]
            raise ValueError(
                f"{label}: sheet_gid {sheet_gid} tidak ditemukan. "
                f"Available: {available}"
            )

    else:
        candidates = properties

    required_norm = {
        normalize_header(x)
        for x in required_headers
    }

    def find_header_row(rows):
        for row_number, raw_header in enumerate(rows[:20], start=1):
            header = trim_header(raw_header)
            if not header:
                continue
            normalized = {
                normalize_header(v)
                for v in header
                if v
            }
            if required_norm.issubset(normalized):
                return row_number, header
        return None, None

    # --------------------------------------------------------------
    # Known tab: one data request only.
    # Use grid row/column counts from cached metadata.
    # --------------------------------------------------------------
    if len(candidates) == 1 and (sheet_name is not None or sheet_gid is not None):
        prop = candidates[0]
        title = prop["title"]
        grid = prop.get("gridProperties", {}) or {}
        row_count = max(int(grid.get("rowCount", 1) or 1), 1)
        col_count = max(int(grid.get("columnCount", 1) or 1), 1)
        end_col = column_letter(col_count)

        values = execute_with_retry(
            service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=f"'{title}'!A1:{end_col}{row_count}",
                valueRenderOption="FORMATTED_VALUE",
            )
        ).get("values", [])

        if not values:
            raise ValueError(f'{label}: tab "{title}" kosong.')

        header_row_number, header = find_header_row(values)
        if header_row_number is None or header is None:
            raise ValueError(
                f'{label}: header wajib tidak ditemukan di tab "{title}". '
                f"Required: {required_headers}"
            )

        data_rows = rows_to_dicts(
            header,
            values[header_row_number:],
        )

        meta = {
            "spreadsheet_id": spreadsheet_id,
            "sheet_id": prop["sheetId"],
            "sheet_title": title,
            "header_row": header_row_number,
            "columns": len(header),
            "rows": len(data_rows),
        }
        return header, data_rows, meta

    # --------------------------------------------------------------
    # Auto-detect: cheap scan first, then one full read for match.
    # --------------------------------------------------------------
    for prop in candidates:
        title = prop["title"]

        scan = execute_with_retry(
            service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=f"'{title}'!{SOURCE_SCAN_RANGE}",
                valueRenderOption="FORMATTED_VALUE",
            )
        ).get("values", [])

        header_row_number, header = find_header_row(scan)
        if header_row_number is None or header is None:
            continue

        end_col = column_letter(len(header))
        grid = prop.get("gridProperties", {}) or {}
        row_count = max(int(grid.get("rowCount", 1) or 1), header_row_number)

        values = execute_with_retry(
            service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=(
                    f"'{title}'!"
                    f"A{header_row_number}:{end_col}{row_count}"
                ),
                valueRenderOption="FORMATTED_VALUE",
            )
        ).get("values", [])

        if not values:
            raise ValueError(f'{label}: tab "{title}" kosong.')

        final_header = trim_header(values[0])
        data_rows = rows_to_dicts(
            final_header,
            values[1:],
        )

        meta = {
            "spreadsheet_id": spreadsheet_id,
            "sheet_id": prop["sheetId"],
            "sheet_title": title,
            "header_row": header_row_number,
            "columns": len(final_header),
            "rows": len(data_rows),
        }
        return final_header, data_rows, meta

    raise ValueError(
        f"{label}: tidak menemukan header yang sesuai. "
        f"Required: {required_headers}"
    )


# =============================================================================
# PRODUCT-NAME BRAND CLASSIFICATION
# =============================================================================

def find_brand_mentions(product_name: object):
    text = str(product_name or "")
    mentions = []

    for base_brand, pattern in BASE_BRAND_PATTERNS.items():
        match = pattern.search(text)
        if match:
            mentions.append((match.start(), match.end(), base_brand, match.group(0)))

    mentions.sort(key=lambda item: (item[0], item[1], item[2]))
    return mentions


def extract_parent_brand_from_product_name(product_name: object) -> str:
    text = str(product_name or "")
    mentions = find_brand_mentions(text)

    if not mentions:
        return MISSING_BRAND

    _start, _end, parent_base, _matched = mentions[0]
    if parent_base != "DOVE":
        return parent_base

    next_brand_start = mentions[1][0] if len(mentions) > 1 else len(text)
    parent_segment = text[:next_brand_start]
    return "DOVE DEO" if DOVE_DEO_RE.search(parent_segment) else "DOVE SCL"


def canonicalize_mapping_brand(value: object) -> str:
    raw = str(value or "").strip().upper()
    if not raw:
        return ""

    canonical = MAPPING_BRAND_ALIASES.get(raw, raw)
    return canonical if canonical in PC_BRAND_ORDER else ""


def resolve_brand_with_pid_fallback(
    product_name: object,
    product_id: object,
    product_brand_mapping: dict[str, str] | None = None,
) -> tuple[str, str]:
    """Product name primary; Product ID mapping only if product-name Brand is MISSING."""
    brand = extract_parent_brand_from_product_name(product_name)

    if brand != MISSING_BRAND:
        return brand, "PRODUCT_NAME"

    if product_brand_mapping:
        pid = normalize_id(product_id)
        mapped_brand = canonicalize_mapping_brand(
            product_brand_mapping.get(pid, "") if pid else ""
        )
        if mapped_brand:
            return mapped_brand, "PRODUCT_ID_MAPPING_FALLBACK"

    return MISSING_BRAND, "MISSING"


def most_common_text(counter: Counter) -> str:
    return counter.most_common(1)[0][0] if counter else ""


# =============================================================================
# PRODUCT ID -> BRAND FALLBACK MAPPING
# =============================================================================

def get_sheet_title_by_gid(service, spreadsheet_id: str, gid: int) -> str:
    properties = get_sheet_properties(service, spreadsheet_id)
    for prop in properties:
        if int(prop["sheetId"]) == int(gid):
            return str(prop["title"])

    available = [(prop["sheetId"], prop["title"]) for prop in properties]
    raise ValueError(
        f"Mapping sheet GID {gid} tidak ditemukan. Available: {available}"
    )


def load_product_brand_mapping(
    service,
    spreadsheet_id: str = MAPPING_SPREADSHEET_ID,
    sheet_gid: int = MAPPING_SHEET_GID,
    pid_header: str = MAPPING_PRODUCT_ID_HEADER,
    brand_header: str = MAPPING_BRAND_HEADER,
) -> tuple[dict[str, str], dict[str, object]]:
    """Load Product ID -> canonical PC Brand untuk fallback."""
    title = get_sheet_title_by_gid(service, spreadsheet_id, sheet_gid)

    scan = execute_with_retry(
        service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=f"'{title}'!A1:AZ20",
            valueRenderOption="FORMATTED_VALUE",
        )
    ).get("values", [])

    pid_target = normalize_header(pid_header)
    brand_target = normalize_header(brand_header)
    header_row_number = None
    pid_index = None
    brand_index = None

    for row_number, row in enumerate(scan, start=1):
        normalized = [normalize_header(value) for value in row]
        current_pid_index = next(
            (i for i, value in enumerate(normalized) if value == pid_target),
            None,
        )
        current_brand_index = next(
            (i for i, value in enumerate(normalized) if value == brand_target),
            None,
        )

        if current_pid_index is not None and current_brand_index is not None:
            header_row_number = row_number
            pid_index = current_pid_index
            brand_index = current_brand_index
            break

    if header_row_number is None or pid_index is None or brand_index is None:
        raise ValueError(
            f'Mapping header tidak ditemukan di "{title}". Harus ada '
            f'"{pid_header}" dan "{brand_header}".'
        )

    end_col = column_letter(max(pid_index, brand_index) + 1)
    values = execute_with_retry(
        service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=f"'{title}'!A{header_row_number + 1}:{end_col}",
            valueRenderOption="FORMATTED_VALUE",
        )
    ).get("values", [])

    mapping: dict[str, str] = {}
    conflicts: dict[str, set[str]] = defaultdict(set)
    blank_pid = 0
    skipped_non_pc = 0
    duplicate_same = 0

    for row in values:
        pid = normalize_id(row[pid_index] if pid_index < len(row) else "")
        raw_brand = row[brand_index] if brand_index < len(row) else ""
        brand = canonicalize_mapping_brand(raw_brand)

        if not pid:
            blank_pid += 1
            continue
        if not brand:
            skipped_non_pc += 1
            continue

        if pid in mapping:
            if mapping[pid] == brand:
                duplicate_same += 1
                continue
            conflicts[pid].update({mapping[pid], brand})
            continue

        mapping[pid] = brand

    if conflicts:
        examples = "\n  ".join(
            f"{pid}: {sorted(brands)}"
            for pid, brands in list(conflicts.items())[:20]
        )
        raise ValueError("PRODUCT ID BRAND MAPPING CONFLICT:\n  " + examples)

    return mapping, {
        "sheet_title": title,
        "header_row": header_row_number,
        "mapping_count": len(mapping),
        "blank_pid": blank_pid,
        "skipped_non_pc": skipped_non_pc,
        "duplicate_same": duplicate_same,
    }


# =============================================================================
# MAPPING BUILDERS
# =============================================================================

def build_tap_mapping(
    header: list[str],
    rows: list[dict[str, object]],
    metric_start: date | None,
    metric_end: date | None,
):
    validate_headers(header, TAP_REQUIRED, "TAP")

    h_video = resolve_header(header, "Video ID")
    h_pid = resolve_header(header, "Product ID")
    h_name = resolve_first_header(header, ["Product name"])
    h_gmv = resolve_header(header, "Affiliate video-attributed GMV")
    h_orders = resolve_header(header, "Creator video-attributed orders")
    h_items = resolve_header(header, "Creator-attributed items sold")
    h_post_time = resolve_header(header, "Post time")

    mapping = defaultdict(
        lambda: {
            "gmv": 0.0,
            "orders": 0,
            "items": 0,
            "rows": 0,
            "product_names": Counter(),
        }
    )

    invalid_key = 0
    invalid_date = 0
    period_skipped = 0

    for row in rows:
        video_id = normalize_id(row.get(h_video, ""))
        product_id = normalize_id(row.get(h_pid, ""))

        if not video_id or not product_id:
            invalid_key += 1
            continue

        if metric_start or metric_end:
            try:
                post_dt = parse_tap_post_time(row.get(h_post_time, ""))
            except ValueError:
                invalid_date += 1
                continue

            post_date = post_dt.date() if post_dt else None
            if not in_period(post_date, metric_start, metric_end):
                period_skipped += 1
                continue

        key = (video_id, product_id)
        rec = mapping[key]
        rec["gmv"] += parse_number(row.get(h_gmv, ""))
        rec["orders"] += parse_int(row.get(h_orders, ""))
        rec["items"] += parse_int(row.get(h_items, ""))
        rec["rows"] += 1

        if h_name:
            product_name = str(row.get(h_name, "") or "").strip()
            if product_name:
                rec["product_names"][product_name] += 1

    finalized = {}
    for key, rec in mapping.items():
        finalized[key] = {
            "gmv": rec["gmv"],
            "orders": rec["orders"],
            "items": rec["items"],
            "rows": rec["rows"],
            "product_name": most_common_text(rec["product_names"]),
        }

    stats = {
        "raw_rows": len(rows),
        "pairs": len(finalized),
        "invalid_key": invalid_key,
        "invalid_date": invalid_date,
        "period_skipped": period_skipped,
    }
    return finalized, stats


def build_tap_video_id_set(
    header: list[str],
    rows: list[dict[str, object]],
    metric_start: date | None,
    metric_end: date | None,
):
    """
    Kumpulkan Video ID TAP untuk existence check SC NOT IN TAP.

    Rule final:
    - TAP difilter berdasarkan `Post time`, BUKAN kolom `Date`.
    - Periode memakai METRIC_START_DATE s.d. METRIC_END_DATE.
    - Compare ke SC hanya berdasarkan Video ID.
    - Jadi video TAP yang Post time-nya 1-6 Oktober dianggap ada walaupun
      kolom Date/report date pada raw TAP menunjukkan 7 Oktober.
    """
    h_video = resolve_header(header, "Video ID")
    h_post_time = resolve_header(header, "Post time")

    video_ids = set()
    stats = {
        "raw_rows": len(rows),
        "included_rows": 0,
        "blank_or_invalid_video_id": 0,
        "invalid_post_time": 0,
        "period_skipped": 0,
    }

    for row in rows:
        if metric_start or metric_end:
            try:
                post_dt = parse_tap_post_time(row.get(h_post_time, ""))
            except ValueError:
                stats["invalid_post_time"] += 1
                continue

            post_date = post_dt.date() if post_dt else None
            if not in_period(post_date, metric_start, metric_end):
                stats["period_skipped"] += 1
                continue

        video_id = normalize_id(row.get(h_video, ""))
        if not video_id:
            stats["blank_or_invalid_video_id"] += 1
            continue

        stats["included_rows"] += 1
        video_ids.add(video_id)

    return video_ids, stats


def build_aff_mapping(
    datasets: list[tuple[str, list[str], list[dict[str, object]]]],
    metric_start: date | None,
    metric_end: date | None,
):
    mapping = defaultdict(
        lambda: {
            "gmv": 0.0,
            "items": 0,
            "rows": 0,
            "product_names": Counter(),
        }
    )

    total_raw = 0
    non_video = 0
    invalid_key = 0
    invalid_date = 0
    period_skipped = 0
    duplicate_rows = 0
    seen_fingerprints = set()

    for label, header, rows in datasets:
        validate_headers(header, AFF_REQUIRED, label)

        h_order = resolve_header(header, "Order ID")
        h_pid = resolve_header(header, "Product ID")
        h_name = resolve_first_header(header, ["Product Name", "Product name"])
        h_type = resolve_header(header, "Content Type")
        h_content = resolve_header(header, "Content ID")
        h_gmv = resolve_header(header, "Est. Commission Base")
        h_qty = resolve_header(header, "Quantity")
        h_time = resolve_first_header(header, ["Time Created"])

        total_raw += len(rows)

        for row in rows:
            content_type = str(row.get(h_type, "") or "").strip().lower()
            if content_type != "video":
                non_video += 1
                continue

            video_id = normalize_id(row.get(h_content, ""))
            product_id = normalize_id(row.get(h_pid, ""))
            # Order ID tidak dipakai sebagai join key. Jumlah Order Affiliate
            # mengikuti jumlah analytical row AFF setelah exact-row dedupe.
            _order_id_raw = str(row.get(h_order, "") or "").strip()

            if not video_id or not product_id:
                invalid_key += 1
                continue

            if metric_start or metric_end:
                if h_time is None:
                    raise ValueError(
                        f"Filter METRIC diminta tetapi {label} tidak punya kolom Time Created."
                    )
                try:
                    dt = parse_aff_datetime(row.get(h_time, ""))
                except ValueError:
                    invalid_date += 1
                    continue

                d = dt.date() if dt else None
                if not in_period(d, metric_start, metric_end):
                    period_skipped += 1
                    continue

            fingerprint = tuple(str(row.get(name, "") or "").strip() for name in header)
            if fingerprint in seen_fingerprints:
                duplicate_rows += 1
                continue
            seen_fingerprints.add(fingerprint)

            key = (video_id, product_id)
            rec = mapping[key]
            rec["gmv"] += parse_number(row.get(h_gmv, ""))
            rec["items"] += parse_int(row.get(h_qty, ""))
            rec["rows"] += 1

            if h_name:
                product_name = str(row.get(h_name, "") or "").strip()
                if product_name:
                    rec["product_names"][product_name] += 1

    finalized = {}
    for key, rec in mapping.items():
        finalized[key] = {
            "gmv": rec["gmv"],
            "orders": rec["rows"],
            "items": rec["items"],
            "rows": rec["rows"],
            "product_name": most_common_text(rec["product_names"]),
        }

    stats = {
        "raw_rows": total_raw,
        "pairs": len(finalized),
        "non_video": non_video,
        "invalid_key": invalid_key,
        "invalid_date": invalid_date,
        "period_skipped": period_skipped,
        "duplicate_rows": duplicate_rows,
    }
    return finalized, stats


def build_master_mapping(
    tap_map: dict,
    aff_map: dict,
    product_brand_mapping: dict[str, str] | None = None,
):
    master = {}
    counts = {"BOTH": 0, "TAP_ONLY": 0, "AFF_ONLY": 0}

    for key in set(tap_map) | set(aff_map):
        tap = tap_map.get(key)
        aff = aff_map.get(key)

        if tap is not None and aff is not None:
            status = "BOTH"
        elif tap is not None:
            status = "TAP_ONLY"
        else:
            status = "AFF_ONLY"

        counts[status] += 1

        # Metric final: TAP prioritas, Affiliate sebagai fallback.
        chosen = tap if tap is not None else aff

        product_name = ""
        if tap is not None:
            product_name = str(tap.get("product_name", "") or "").strip()
        if not product_name and aff is not None:
            product_name = str(aff.get("product_name", "") or "").strip()

        resolved_brand, resolved_brand_source = resolve_brand_with_pid_fallback(
            product_name,
            key[1],
            product_brand_mapping,
        )

        master[key] = {
            "mapped_gmv": float(chosen["gmv"]),
            "mapped_orders": int(chosen["orders"]),
            "mapped_items": int(chosen["items"]),
            "metric_source": "TAP" if tap is not None else "AFF",
            "match_status": status,
            "product_name": product_name,
            "brand": resolved_brand,
            "brand_source": resolved_brand_source,
            "tap_gmv": float(tap["gmv"]) if tap is not None else 0.0,
            "tap_orders": int(tap["orders"]) if tap is not None else 0,
            "tap_items": int(tap["items"]) if tap is not None else 0,
            "aff_gmv": float(aff["gmv"]) if aff is not None else 0.0,
            "aff_orders": int(aff["orders"]) if aff is not None else 0,
            "aff_items": int(aff["items"]) if aff is not None else 0,
        }

    return master, counts


def build_pid_product_name_lookup(tap_map: dict, aff_map: dict) -> dict[str, str]:
    tap_names = defaultdict(Counter)
    aff_names = defaultdict(Counter)

    for (_video_id, pid), rec in tap_map.items():
        name = str(rec.get("product_name", "") or "").strip()
        if name:
            tap_names[pid][name] += 1

    for (_video_id, pid), rec in aff_map.items():
        name = str(rec.get("product_name", "") or "").strip()
        if name:
            aff_names[pid][name] += 1

    lookup = {}
    for pid in set(tap_names) | set(aff_names):
        if tap_names.get(pid):
            lookup[pid] = most_common_text(tap_names[pid])
        elif aff_names.get(pid):
            lookup[pid] = most_common_text(aff_names[pid])

    return lookup


def build_tap_brand_pair_set_from_mapping(
    tap_map: dict,
    pid_product_name_lookup: dict[str, str],
    product_brand_mapping: dict[str, str] | None = None,
) -> set[tuple[str, str]]:
    """Build existence key TAP pada grain Video ID + Brand.

    Hanya pair yang benar-benar ada di TAP map yang menjadi evidence existence.
    Product name boleh fallback via PID lookup hanya untuk menentukan Brand.
    """
    pairs: set[tuple[str, str]] = set()
    for (video_id, pid), rec in tap_map.items():
        video_id = normalize_id(video_id)
        if not video_id:
            continue
        product_name = str(rec.get("product_name", "") or "").strip()
        if not product_name:
            product_name = str(pid_product_name_lookup.get(pid, "") or "").strip()
        brand, _brand_source = resolve_brand_with_pid_fallback(
            product_name,
            pid,
            product_brand_mapping,
        )
        if brand not in BRAND_OUTPUT_ORDER:
            brand = MISSING_BRAND
        pairs.add((video_id, brand))
    return pairs


# =============================================================================
# SC ENRICHMENT
# =============================================================================

def find_sc_source_column(sc_header: list[str], output_name: str) -> str | None:
    if output_name in SC_SOURCE_ALIASES:
        return resolve_first_header(sc_header, SC_SOURCE_ALIASES[output_name])
    return resolve_first_header(sc_header, [output_name])


OPTIONAL_CONTENT_METRICS = ('Likes', 'Comments', 'Shares', 'Video product impressions', 'Video product clicks', 'Video views')


def metric_missing(value):
    return value is None or str(value).strip().upper() in {"", "-", "--", "N/A", "NA", "NULL", "NONE", "NAN"}


def clean_sc_value(output_name: str, value: object):
    if output_name in OPTIONAL_CONTENT_METRICS and metric_missing(value):
        return ""

    if output_name in {
        "Video-attributed orders",
        "Video-attributed items sold",
        "Items refunded",
        "Likes",
        "Comments",
        "Shares",
        "Video product impressions",
        "Video product clicks",
        "Video views",
    }:
        return parse_int(value)

    if output_name in {
        "Creator video-attributed GMV",
        "AOV",
        "Refunds",
        "Video GPM",
        "Avg. GMV per customer",
        "Est. commission",
    }:
        return clean_numeric(parse_number(value))

    if output_name in {"Completion rate", "CTR", "Engagement"}:
        return clean_numeric(parse_percent(value))

    if output_name == "Video ID":
        return normalize_id(value)

    if output_name == "Post date":
        return normalize_sc_post_date(value)

    return "" if value is None else str(value).strip()


def enrich_sc(
    header: list[str],
    rows: list[dict[str, object]],
    master_map: dict,
    aff_map: dict,
    pid_product_name_lookup: dict[str, str],
    product_brand_mapping: dict[str, str] | None = None,
):
    validate_headers(header, SC_REQUIRED, "SC")

    source_cols = {
        out_name: find_sc_source_column(header, out_name)
        for out_name in OUTPUT_COLUMNS
        if out_name not in {
            "Product name",
            "Brand",
            "GMV Affiliate Order (Est. Commission Base)",
            "Jumlah Order Affiliate",
        }
    }

    missing_template_sources = [
        out_name
        for out_name, source_name in source_cols.items()
        if source_name is None
        and out_name not in {
            "CTR",
            "Video GPM",
            "Engagement",
            "Video link",
            "Creator name",
            "Refunds",
            "Items refunded",
            "Likes",
            "Comments",
            "Shares",
            "Video product impressions",
            "Video product clicks",
            "Completion rate",
            "Video views",
            "Avg. GMV per customer",
            "Est. commission",
        }
    ]

    if missing_template_sources:
        raise ValueError(
            "SC tidak memiliki source untuk kolom output wajib: "
            + ", ".join(missing_template_sources)
        )

    h_video = resolve_header(header, "Video ID")
    h_pid = resolve_header(header, "Product ID")
    h_sc_name = resolve_first_header(header, ["Product name", "Product Name"])
    h_sc_gmv = resolve_header(header, "Creator video-attributed GMV")
    h_sc_orders = resolve_header(header, "Video-attributed orders")
    h_sc_items = resolve_header(header, "Video-attributed items sold")

    output_rows = []
    missing_details = defaultdict(
        lambda: {
            "count": 0,
            "example_title": "",
            "source_product_count": 0,
        }
    )

    stats = {
        "raw_rows": len(rows),
        "exploded_rows": 0,
        "multi_pid_source_rows": 0,
        "blank_pid_source_rows": 0,
        "multi_pid_mapped_tap": 0,
        "multi_pid_mapped_aff_only": 0,
        "single_pid_sc_preserved": 0,
        "unmapped_multi_pid": 0,
        "invalid_video_id": 0,
        "brand_from_product_name": 0,
        "brand_from_product_id_fallback": 0,
        "brand_missing": 0,
    }

    original_totals = {"gmv": 0.0, "orders": 0, "items": 0}
    enriched_totals = {"gmv": 0.0, "orders": 0, "items": 0}
    brand_counter = Counter()

    for row in rows:
        video_id = normalize_id(row.get(h_video, ""))
        if not video_id:
            stats["invalid_video_id"] += 1
            continue

        product_ids = split_product_ids(row.get(h_pid, ""))
        if not product_ids:
            stats["blank_pid_source_rows"] += 1
            continue

        if len(product_ids) > 1:
            stats["multi_pid_source_rows"] += 1

        sc_original_gmv = parse_number(row.get(h_sc_gmv, ""))
        sc_original_orders = parse_int(row.get(h_sc_orders, ""))
        sc_original_items = parse_int(row.get(h_sc_items, ""))

        original_totals["gmv"] += sc_original_gmv
        original_totals["orders"] += sc_original_orders
        original_totals["items"] += sc_original_items

        sc_product_name = (
            str(row.get(h_sc_name, "") or "").strip()
            if h_sc_name else ""
        )

        source_title_col = source_cols.get("Video title")
        example_title = str(row.get(source_title_col, "") or "").strip() if source_title_col else ""

        for pid in product_ids:
            key = (video_id, pid)
            master = master_map.get(key)
            aff = aff_map.get(key)

            out = {}
            for out_name in OUTPUT_COLUMNS:
                if out_name in {
                    "Product name",
                    "Brand",
                    "GMV Affiliate Order (Est. Commission Base)",
                    "Jumlah Order Affiliate",
                }:
                    continue

                if out_name == "Product ID":
                    out[out_name] = pid
                    continue

                source_name = source_cols.get(out_name)
                raw_value = row.get(source_name, "") if source_name else ""
                out[out_name] = clean_sc_value(out_name, raw_value)

            out["Video ID"] = video_id

            mapped_product_name = ""
            if master is not None:
                mapped_product_name = str(master.get("product_name", "") or "").strip()
            if not mapped_product_name:
                mapped_product_name = str(pid_product_name_lookup.get(pid, "") or "").strip()
            if not mapped_product_name:
                mapped_product_name = sc_product_name

            brand, brand_source = resolve_brand_with_pid_fallback(
                mapped_product_name,
                pid,
                product_brand_mapping,
            )
            if brand not in BRAND_OUTPUT_ORDER:
                brand = MISSING_BRAND
                brand_source = "MISSING"

            if brand_source == "PRODUCT_NAME":
                stats["brand_from_product_name"] += 1
            elif brand_source == "PRODUCT_ID_MAPPING_FALLBACK":
                stats["brand_from_product_id_fallback"] += 1
            else:
                stats["brand_missing"] += 1

            # Product name dan Brand sekarang menjadi kolom output resmi.
            # Product name: exact Video ID + Product ID TAP/AFF -> global PID lookup
            # -> Product name SC bila tersedia.
            # Brand: diturunkan dari Product name dengan rule parent-brand TAP PC.
            out["Product name"] = mapped_product_name
            out["Brand"] = brand
            out["_BRAND_SOURCE"] = brand_source

            # =============================================================
            # CORRECTED COMMERCE RULE
            # =============================================================
            # Single-PID row tidak butuh allocation. Metric SC original
            # sudah berada pada satu Product ID saja, jadi jangan overwrite
            # dengan TAP/AFF walaupun pair tersebut tersedia di master.
            # Ini mencegah drift seperti SC Orders 2,485 menjadi TAP 2,506.
            if len(product_ids) == 1:
                mapped_gmv = sc_original_gmv
                mapped_orders = sc_original_orders
                mapped_items = sc_original_items
                stats["single_pid_sc_preserved"] += 1

            # Multi-PID baru membutuhkan mapping per Video ID + Product ID.
            elif master is not None:
                mapped_gmv = float(master["mapped_gmv"])
                mapped_orders = int(master["mapped_orders"])
                mapped_items = int(master["mapped_items"])

                if master["metric_source"] == "TAP":
                    stats["multi_pid_mapped_tap"] += 1
                else:
                    stats["multi_pid_mapped_aff_only"] += 1

            else:
                # Multi-PID tetapi pair tidak ditemukan di master mapping.
                # Jangan copy metric level-video ke PID ini karena akan double count.
                mapped_gmv = 0.0
                mapped_orders = 0
                mapped_items = 0
                stats["unmapped_multi_pid"] += 1

                detail = missing_details[key]
                detail["count"] += 1
                detail["source_product_count"] = len(product_ids)
                if not detail["example_title"]:
                    detail["example_title"] = example_title

            mapped_aov = (mapped_gmv / mapped_orders) if mapped_orders else 0.0

            out["Creator video-attributed GMV"] = clean_numeric(mapped_gmv)
            out["Video-attributed orders"] = mapped_orders
            out["AOV"] = clean_numeric(mapped_aov)
            out["Video-attributed items sold"] = mapped_items

            out["GMV Affiliate Order (Est. Commission Base)"] = clean_numeric(
                float(aff["gmv"]) if aff is not None else 0.0
            )
            out["Jumlah Order Affiliate"] = int(aff["orders"]) if aff is not None else 0

            # Helper internal untuk split/diagnostic; Brand juga sudah ditulis sebagai kolom output.
            out["_BRAND"] = brand
            if len(product_ids) == 1:
                out["_MAPPING_SOURCE"] = "SC_SINGLE_PID_PRESERVED"
            elif master is not None:
                out["_MAPPING_SOURCE"] = master["metric_source"]
            else:
                out["_MAPPING_SOURCE"] = "MISSING"

            output_rows.append(out)
            stats["exploded_rows"] += 1
            brand_counter[brand] += 1

            enriched_totals["gmv"] += mapped_gmv
            enriched_totals["orders"] += mapped_orders
            enriched_totals["items"] += mapped_items

    return (
        output_rows,
        stats,
        original_totals,
        enriched_totals,
        brand_counter,
        missing_details,
    )


def filter_output_range(
    rows: list[dict[str, object]],
    range_start: date | None,
    range_end: date | None,
):
    if range_start is None and range_end is None:
        return list(rows), 0

    filtered = []
    invalid_date = 0

    for row in rows:
        d = parse_output_post_date(row.get("Post date", ""))
        if d is None:
            invalid_date += 1
            continue
        if in_period(d, range_start, range_end):
            filtered.append(row)

    return filtered, invalid_date


def build_brand_tables(rows: list[dict[str, object]]):
    grouped = {brand: [] for brand in BRAND_OUTPUT_ORDER}

    for row in rows:
        brand = str(row.get("_BRAND", MISSING_BRAND) or MISSING_BRAND)
        if brand not in grouped:
            brand = MISSING_BRAND
        grouped[brand].append(row)

    return grouped


# =============================================================================
# OUTPUT TABLE BUILDERS
# =============================================================================

def data_values(rows: list[dict[str, object]]) -> list[list[object]]:
    return [OUTPUT_COLUMNS] + [
        [row.get(col, "") for col in OUTPUT_COLUMNS]
        for row in rows
    ]


def build_brand_summary(all_rows, range_rows):
    all_counter = Counter(str(row.get("_BRAND", MISSING_BRAND)) for row in all_rows)
    range_counter = Counter(str(row.get("_BRAND", MISSING_BRAND)) for row in range_rows)

    values = [["Brand", "Rows ALL", "Rows RANGE", "GMV ALL", "GMV RANGE"]]

    for brand in BRAND_OUTPUT_ORDER:
        gmv_all = sum(
            parse_number(row.get("Creator video-attributed GMV", 0))
            for row in all_rows
            if row.get("_BRAND") == brand
        )
        gmv_range = sum(
            parse_number(row.get("Creator video-attributed GMV", 0))
            for row in range_rows
            if row.get("_BRAND") == brand
        )

        values.append([
            brand,
            all_counter.get(brand, 0),
            range_counter.get(brand, 0),
            clean_numeric(gmv_all),
            clean_numeric(gmv_range),
        ])

    return values


def build_missing_values(missing_details):
    values = [[
        "Video ID",
        "Product ID",
        "Occurrences",
        "Product Count in Source Row",
        "Example Video Title",
        "Reason",
    ]]

    for (video_id, product_id), detail in sorted(missing_details.items()):
        values.append([
            video_id,
            product_id,
            detail["count"],
            detail["source_product_count"],
            detail["example_title"],
            "Video ID + Product ID tidak ditemukan di master TAP + AFF",
        ])

    return values


def filter_sc_not_in_tap(
    rows: list[dict[str, object]],
    tap_video_brand_pairs: set[tuple[str, str]],
):
    """Return exploded SC rows yang Video ID + Brand tidak ada di TAP."""
    result = []
    for row in rows:
        video_id = normalize_id(row.get("Video ID", ""))
        brand = str(
            row.get("_BRAND", row.get("Brand", MISSING_BRAND))
            or MISSING_BRAND
        ).strip()
        if video_id and (video_id, brand) not in tap_video_brand_pairs:
            result.append(row)
    return result


def build_brand_unique_rows(
    rows: list[dict[str, object]],
) -> list[dict[str, object]]:
    """
    Ringkas exploded SC Video menjadi grain unik Video ID + Brand.

    - Product ID output = PID pertama yang muncul pada Video ID + Brand.
    - GMV / Orders / Items = SUM metric per Product ID unik pada brand tersebut.
    - PID yang sama tidak boleh menyumbang metric dua kali.
    - Product name mengikuti representative/first PID.
    - Metric video-level lain mengikuti row representative dan jangan dijumlahkan lintas Brand.
    """
    grouped: dict[tuple[str, str], dict[str, object]] = {}
    key_order: list[tuple[str, str]] = []

    for row in rows:
        video_id = normalize_id(row.get("Video ID", ""))
        if not video_id:
            continue

        brand = str(
            row.get("_BRAND", row.get("Brand", MISSING_BRAND))
            or MISSING_BRAND
        ).strip() or MISSING_BRAND
        key = (video_id, brand)
        pid = normalize_id(row.get("Product ID", ""))

        if key not in grouped:
            grouped[key] = {
                "row": dict(row),
                "seen_pids": set(),
                "gmv": 0.0,
                "orders": 0,
                "items": 0,
                "aff_gmv": 0.0,
                "aff_orders": 0,
            }
            key_order.append(key)

        rec = grouped[key]
        if pid and pid not in rec["seen_pids"]:
            rec["seen_pids"].add(pid)
            rec["gmv"] += parse_number(
                row.get("Creator video-attributed GMV", 0)
            )
            rec["orders"] += parse_int(
                row.get("Video-attributed orders", 0)
            )
            rec["items"] += parse_int(
                row.get("Video-attributed items sold", 0)
            )
            rec["aff_gmv"] += parse_number(
                row.get("GMV Affiliate Order (Est. Commission Base)", 0)
            )
            rec["aff_orders"] += parse_int(
                row.get("Jumlah Order Affiliate", 0)
            )

    result: list[dict[str, object]] = []
    for key in key_order:
        rec = grouped[key]
        out = rec["row"]
        gmv = float(rec["gmv"])
        orders = int(rec["orders"])
        items = int(rec["items"])

        out["Creator video-attributed GMV"] = clean_numeric(gmv)
        out["Video-attributed orders"] = orders
        out["Video-attributed items sold"] = items
        out["AOV"] = clean_numeric(gmv / orders if orders else 0.0)
        out["GMV Affiliate Order (Est. Commission Base)"] = clean_numeric(
            float(rec["aff_gmv"])
        )
        out["Jumlah Order Affiliate"] = int(rec["aff_orders"])
        out["_BRAND"] = key[1]
        out["_MAPPING_SOURCE"] = "BRAND_SUM_UNIQUE_PID"
        out["_BRAND_PID_COUNT"] = len(rec["seen_pids"])
        out["_BRAND_PRODUCT_IDS"] = ",".join(sorted(rec["seen_pids"]))
        result.append(out)

    seen_keys = set()
    for row in result:
        key = (
            normalize_id(row.get("Video ID", "")),
            str(row.get("_BRAND", MISSING_BRAND) or MISSING_BRAND),
        )
        if key in seen_keys:
            raise RuntimeError(f"Duplicate BRAND UNIQUE key terdeteksi: {key}")
        seen_keys.add(key)

    return result


def build_output_tables(
    all_rows,
    range_rows,
    sc_not_in_tap_all_rows,
    sc_not_in_tap_range_rows,
    missing_details,
    write_brand_tabs,
):
    tables = {
        SHEET_ALL: data_values(all_rows),
        SHEET_RANGE: data_values(range_rows),
        SC_NOT_IN_TAP_ALL_SHEET: data_values(sc_not_in_tap_all_rows),
        SC_NOT_IN_TAP_RANGE_SHEET: data_values(sc_not_in_tap_range_rows),
    }

    if write_brand_tabs:
        brand_all = build_brand_tables(all_rows)
        brand_range = build_brand_tables(range_rows)

        for brand in BRAND_OUTPUT_ORDER:
            tables[f"{brand}{BRAND_ALL_SUFFIX}"] = data_values(brand_all[brand])

        for brand in BRAND_OUTPUT_ORDER:
            tables[f"{brand}{BRAND_RANGE_SUFFIX}"] = data_values(brand_range[brand])

    tables[SUMMARY_SHEET] = build_brand_summary(all_rows, range_rows)
    tables[MISSING_SHEET] = build_missing_values(missing_details)

    return tables


def all_managed_titles():
    titles = {
        SHEET_ALL,
        SHEET_RANGE,
        SC_NOT_IN_TAP_ALL_SHEET,
        SC_NOT_IN_TAP_RANGE_SHEET,
        BRAND_UNIQUE_ALL_SHEET,
        BRAND_UNIQUE_RANGE_SHEET,
        BRAND_UNIQUE_NOT_IN_TAP_ALL_SHEET,
        BRAND_UNIQUE_NOT_IN_TAP_RANGE_SHEET,
        SUMMARY_SHEET,
        MISSING_SHEET,
    }

    for brand in BRAND_OUTPUT_ORDER:
        titles.add(f"{brand}{BRAND_ALL_SUFFIX}")
        titles.add(f"{brand}{BRAND_RANGE_SUFFIX}")

    return titles


# =============================================================================
# OUTPUT GOOGLE SHEETS
# =============================================================================

def get_output_grid_info(service):
    properties = get_sheet_properties(service, OUTPUT_SPREADSHEET_ID)

    info = {}
    total_cells = 0

    for prop in properties:
        grid = prop.get("gridProperties", {})
        rows = int(grid.get("rowCount", 0) or 0)
        cols = int(grid.get("columnCount", 0) or 0)
        cells = rows * cols

        info[prop["title"]] = {
            "sheetId": prop["sheetId"],
            "rows": rows,
            "cols": cols,
            "cells": cells,
        }
        total_cells += cells

    return info, total_cells


def cells_for_values(values):
    row_count = max(len(values), 2)
    col_count = max(max((len(row) for row in values), default=1), 1)
    return row_count * col_count


def validate_output_capacity(service, output_tables):
    info, existing_cells = get_output_grid_info(service)
    requested_titles = set(output_tables)
    managed_existing = set(info) & all_managed_titles()
    clashes = requested_titles & set(info)

    if clashes and not REPLACE:
        raise SystemExit(
            "Tab output sudah ada dan tidak akan ditimpa: "
            + ", ".join(sorted(clashes))
            + "\nJalankan dengan --write --replace jika memang mau mengganti output."
        )

    removed_titles = managed_existing if REPLACE else set()
    cells_removed = sum(info[t]["cells"] for t in removed_titles)
    new_cells = sum(cells_for_values(values) for values in output_tables.values())
    projected = existing_cells - cells_removed + new_cells

    print("\nValidasi kapasitas output:")
    print(f"  Existing cells       : {existing_cells:,}")
    if REPLACE:
        print(f"  Managed cells removed: -{cells_removed:,}")
    print(f"  New output cells     : {new_cells:,}")
    print(f"  Projected total      : {projected:,}")
    print(f"  Safety limit         : {MAX_CELLS:,}")

    if projected > MAX_CELLS:
        raise SystemExit(
            "Projected output melebihi MAX_CELLS. Gunakan output spreadsheet lain "
            "atau nonaktifkan brand tabs."
        )

    return info


def create_temp_sheet_if_needed(service, info, delete_titles):
    existing_titles = set(info)
    if not delete_titles or delete_titles != existing_titles:
        return None

    temp_title = "_TMP_AUTOMATION_DO_NOT_USE"
    suffix = 1
    while temp_title in existing_titles:
        suffix += 1
        temp_title = f"_TMP_AUTOMATION_DO_NOT_USE_{suffix}"

    response = execute_with_retry(
        service.spreadsheets().batchUpdate(
            spreadsheetId=OUTPUT_SPREADSHEET_ID,
            body={"requests": [{"addSheet": {"properties": {"title": temp_title}}}]},
        )
    )

    invalidate_sheet_properties_cache(OUTPUT_SPREADSHEET_ID)

    replies = response.get("replies", [])
    if replies:
        return replies[0]["addSheet"]["properties"]["sheetId"]
    return None


def recreate_output_tabs(service, info, output_tables):
    delete_titles = (set(info) & all_managed_titles()) if REPLACE else set()
    temp_sheet_id = create_temp_sheet_if_needed(service, info, delete_titles)

    requests = []

    for title in sorted(delete_titles):
        requests.append({"deleteSheet": {"sheetId": info[title]["sheetId"]}})

    for title, values in output_tables.items():
        row_count = max(len(values), 2)
        col_count = max(max((len(row) for row in values), default=1), 1)
        requests.append({
            "addSheet": {
                "properties": {
                    "title": title,
                    "gridProperties": {
                        "rowCount": row_count,
                        "columnCount": col_count,
                        "frozenRowCount": 1,
                    },
                }
            }
        })

    if requests:
        execute_with_retry(
            service.spreadsheets().batchUpdate(
                spreadsheetId=OUTPUT_SPREADSHEET_ID,
                body={"requests": requests},
            )
        )

    if temp_sheet_id is not None:
        execute_with_retry(
            service.spreadsheets().batchUpdate(
                spreadsheetId=OUTPUT_SPREADSHEET_ID,
                body={"requests": [{"deleteSheet": {"sheetId": temp_sheet_id}}]},
            )
        )

    invalidate_sheet_properties_cache(OUTPUT_SPREADSHEET_ID)


def write_values_chunked(service, title, values):
    for start in range(0, len(values), WRITE_CHUNK_ROWS):
        part = values[start : start + WRITE_CHUNK_ROWS]
        execute_with_retry(
            service.spreadsheets().values().update(
                spreadsheetId=OUTPUT_SPREADSHEET_ID,
                range=f"'{title}'!A{start + 1}",
                valueInputOption="RAW",
                body={"values": part},
            )
        )

    print(f'  "{title}" selesai ({max(len(values) - 1, 0):,} row)')


def write_output(service, output_tables, info):
    recreate_output_tabs(service, info, output_tables)

    for title, values in output_tables.items():
        write_values_chunked(service, title, values)


# =============================================================================
# DIAGNOSTICS
# =============================================================================

def pct_diff(mapped: float, original: float) -> float | None:
    if math.isclose(original, 0.0, abs_tol=1e-12):
        return None
    return (mapped - original) / original * 100.0


def fmt_num(value: float | int) -> str:
    if isinstance(value, int) or math.isclose(float(value), round(float(value)), abs_tol=1e-9):
        return f"{int(round(float(value))):,}"
    return f"{float(value):,.2f}"


def print_source_meta(label, meta):
    print(f"\n{label} source:")
    print(f"  Spreadsheet ID : {meta['spreadsheet_id']}")
    print(f"  Sheet          : {meta['sheet_title']}")
    print(f"  GID            : {meta['sheet_id']}")
    print(f"  Header row     : {meta['header_row']}")
    print(f"  Columns        : {meta['columns']}")
    print(f"  Rows           : {meta['rows']:,}")


def print_summary(
    tap_stats,
    aff_stats,
    master_counts,
    sc_stats,
    original_totals,
    enriched_totals,
    brand_counter,
    all_rows,
    range_rows,
    sc_not_in_tap_all_rows,
    sc_not_in_tap_range_rows,
    tap_video_ids,
    tap_existence_stats,
    invalid_range_dates,
    range_start,
    range_end,
    write_brand_tabs,
    missing_details,
):
    print("\n" + "=" * 78)
    print("SC VIDEO - EXPLODE + TAP/AFF PRODUCT MAPPING")
    print("=" * 78)

    print("\nTAP mapping:")
    print(f"  Raw row                  : {tap_stats['raw_rows']:,}")
    print(f"  Video ID + Product ID    : {tap_stats['pairs']:,}")
    print(f"  Blank/invalid key        : {tap_stats['invalid_key']:,}")
    print(f"  Invalid Post time        : {tap_stats['invalid_date']:,}")
    print(f"  Outside metric period    : {tap_stats['period_skipped']:,}")

    print("\nAffiliate mapping:")
    print(f"  Raw row                  : {aff_stats['raw_rows']:,}")
    print(f"  Non-Video skipped        : {aff_stats['non_video']:,}")
    print(f"  Video ID + Product ID    : {aff_stats['pairs']:,}")
    print(f"  Exact duplicate skipped  : {aff_stats['duplicate_rows']:,}")
    print(f"  Blank/invalid key        : {aff_stats['invalid_key']:,}")
    print(f"  Invalid date             : {aff_stats['invalid_date']:,}")
    print(f"  Outside metric period    : {aff_stats['period_skipped']:,}")

    print("\nMaster TAP + AFF mapping:")
    print(f"  BOTH                     : {master_counts['BOTH']:,}")
    print(f"  TAP only                 : {master_counts['TAP_ONLY']:,}")
    print(f"  AFF only                 : {master_counts['AFF_ONLY']:,}")
    print(f"  Total pair               : {sum(master_counts.values()):,}")

    print("\nSC explode + enrichment:")
    print(f"  Raw SC row               : {sc_stats['raw_rows']:,}")
    print(f"  Source row multi-PID     : {sc_stats['multi_pid_source_rows']:,}")
    print(f"  Exploded output row      : {sc_stats['exploded_rows']:,}")
    print(f"  Single-PID SC preserved  : {sc_stats['single_pid_sc_preserved']:,}")
    print(f"  Multi-PID mapped TAP     : {sc_stats['multi_pid_mapped_tap']:,}")
    print(f"  Multi-PID mapped AFF     : {sc_stats['multi_pid_mapped_aff_only']:,}")
    print(f"  Unmapped multi-PID row   : {sc_stats['unmapped_multi_pid']:,}")
    print(f"  Source row blank PID     : {sc_stats['blank_pid_source_rows']:,}")
    print(f"  Invalid Video ID         : {sc_stats['invalid_video_id']:,}")
    print(f"  Unique missing pair      : {len(missing_details):,}")

    print("\nBrand derivation (Product name -> Product ID fallback jika MISSING):")
    print(
        f"  Product name primary     : "
        f"{sc_stats['brand_from_product_name']:>10,} row"
    )
    print(
        f"  Product ID fallback      : "
        f"{sc_stats['brand_from_product_id_fallback']:>10,} row"
    )
    print(
        f"  Tetap MISSING            : "
        f"{sc_stats['brand_missing']:>10,} row"
    )
    for brand in BRAND_OUTPUT_ORDER:
        print(f"  {brand:<15} {brand_counter.get(brand, 0):>10,} row")

    print("\nALL / RANGE:")
    if range_start or range_end:
        start_text = range_start.isoformat() if range_start else "-inf"
        end_text = range_end.isoformat() if range_end else "+inf"
        print(f"  RANGE                    : {start_text} -> {end_text} (Post date)")
    else:
        print("  RANGE                    : same as ALL (RANGE_START/END = None)")

    print(f"  {SHEET_ALL:<25}: {len(all_rows):>10,} row")
    print(f"  {SHEET_RANGE:<25}: {len(range_rows):>10,} row")
    print(f"  Invalid Post date RANGE  : {invalid_range_dates:>10,}")
    print(f"  Brand tabs               : {'YES' if write_brand_tabs else 'NO'}")

    missing_all_video_ids = {
        normalize_id(row.get("Video ID", ""))
        for row in sc_not_in_tap_all_rows
        if normalize_id(row.get("Video ID", ""))
    }
    missing_range_video_ids = {
        normalize_id(row.get("Video ID", ""))
        for row in sc_not_in_tap_range_rows
        if normalize_id(row.get("Video ID", ""))
    }

    print("\nSC Video ID vs TAP Video ID (TAP POST TIME RANGE):")
    print(f"  TAP raw rows             : {tap_existence_stats['raw_rows']:>10,}")
    print(f"  TAP rows in period       : {tap_existence_stats['included_rows']:>10,}")
    print(f"  TAP outside period       : {tap_existence_stats['period_skipped']:>10,}")
    print(f"  TAP invalid Post time         : {tap_existence_stats['invalid_post_time']:>10,}")
    print(
        f"  TAP blank Video ID       : "
        f"{tap_existence_stats['blank_or_invalid_video_id']:>10,}"
    )
    print(f"  Unique TAP Video ID RANGE: {len(tap_video_ids):>10,}")
    print(f"  Missing SC Video ID ALL  : {len(missing_all_video_ids):>10,}")
    print(f"  Missing SC rows ALL      : {len(sc_not_in_tap_all_rows):>10,}")
    print(f"  Missing SC Video ID RANGE: {len(missing_range_video_ids):>10,}")
    print(f"  Missing SC rows RANGE    : {len(sc_not_in_tap_range_rows):>10,}")
    print("  Compare key              : Video ID saja")
    print("  TAP existence source     : TAP.Post time yang lolos METRIC_START_DATE/END_DATE")

    print("\nCommerce metric reconciliation:")
    labels = {"gmv": "GMV", "orders": "Orders", "items": "Items sold"}

    for key in ("gmv", "orders", "items"):
        original = original_totals[key]
        enriched = enriched_totals[key]
        delta = enriched - original
        error = pct_diff(float(enriched), float(original))
        error_text = "N/A" if error is None else f"{error:+.2f}%"

        print(
            f"  {labels[key]:<10} SC original {fmt_num(original):>16} | "
            f"mapped {fmt_num(enriched):>16} | "
            f"delta {fmt_num(delta):>16} | {error_text}"
        )

    print("\nOutput:")
    print(f"  Spreadsheet ID           : {OUTPUT_SPREADSHEET_ID}")
    print(f"  Template output          : {len(OUTPUT_COLUMNS)} kolom")
    print(f"  Mode                      : {'WRITE' if WRITE else 'DRY RUN'}")
    print(f"  Replace                   : {'YES' if REPLACE else 'NO'}")

    print("\nCATATAN:")
    print("  Product name dan Brand ditulis sebagai kolom output setelah Product ID.")
    print(
        "  Brand: Product name PRIMARY; jika MISSING dan fallback ON, "
        "gunakan Product ID mapping PC."
    )
    print("  Single-PID: metric commerce utama dipertahankan dari SC original.")
    print("  Multi-PID : metric commerce utama dialokasikan dari mapping TAP/AFF.")
    print(
        "  Likes/Comments/Shares/Views/CTR/Engagement dan metric video-level lain "
        "tetap berasal dari SC dan akan terulang pada setiap PID hasil explode."
    )
    print("  Jangan SUM metric video-level tersebut lintas Product ID.")
    print(
        "  SC NOT IN TAP berisi exploded row untuk Video ID + Brand SC yang tidak ditemukan "
        "di TAP.Post time pada periode METRIC_*; pengecekan hanya berdasarkan Video ID."
    )


# =============================================================================
# CONFIG VALIDATION
# =============================================================================

def is_placeholder(value: object) -> bool:
    text = str(value or "").strip().upper()
    return text.startswith("PASTE_") or text.startswith("ISI_") or not text


def validate_config():
    missing = []

    if is_placeholder(SC_SPREADSHEET_ID):
        missing.append("SC_SPREADSHEET_ID")
    if is_placeholder(TAP_SPREADSHEET_ID):
        missing.append("TAP_SPREADSHEET_ID")
    if is_placeholder(OUTPUT_SPREADSHEET_ID):
        missing.append("OUTPUT_SPREADSHEET_ID")

    if USE_PRODUCT_ID_MAPPING_FALLBACK:
        if is_placeholder(MAPPING_SPREADSHEET_ID):
            missing.append("MAPPING_SPREADSHEET_ID")
        if MAPPING_SHEET_GID is None:
            missing.append("MAPPING_SHEET_GID")

    if not AFF_SOURCES:
        missing.append("AFF_SOURCES")
    else:
        for i, source in enumerate(AFF_SOURCES, start=1):
            if is_placeholder(source.get("spreadsheet_id")):
                missing.append(f"AFF_SOURCES[{i}].spreadsheet_id")

    if missing:
        raise SystemExit(
            "Isi CONFIG spreadsheet ID dulu:\n  - " + "\n  - ".join(missing)
        )

    if not CREDENTIALS_FILE.exists():
        candidates = "\n  - ".join(str(path) for path in CREDENTIALS_CANDIDATES)
        raise SystemExit(
            "Credentials tidak ditemukan. Script mencoba:\n  - " + candidates
        )

    metric_start = parse_iso_date(METRIC_START_DATE)
    metric_end = parse_iso_date(METRIC_END_DATE)
    range_start = parse_iso_date(RANGE_START_DATE)
    range_end = parse_iso_date(RANGE_END_DATE)

    if metric_start and metric_end and metric_start > metric_end:
        raise SystemExit("METRIC_START_DATE tidak boleh lebih besar dari METRIC_END_DATE")

    if range_start and range_end and range_start > range_end:
        raise SystemExit("RANGE_START_DATE tidak boleh lebih besar dari RANGE_END_DATE")

    return metric_start, metric_end, range_start, range_end


# =============================================================================
# MAIN
# =============================================================================

def main():
    metric_start, metric_end, range_start, range_end = validate_config()
    service = get_sheets_service()

    print("=" * 78)
    print("SC VIDEO CLEAN + TAP/AFF ENRICHMENT - GOOGLE SHEETS")
    print("=" * 78)
    print(f"Mode             : {'WRITE' if WRITE else 'DRY RUN'}")
    print(f"Replace          : {'YES' if REPLACE else 'NO'}")
    print(f"Brand tabs       : {'YES' if WRITE_BRAND_TABS_EFFECTIVE else 'NO'}")
    print(f"Version          : {CLEANING_VERSION}")
    print(f"Read pacing      : {MIN_READ_INTERVAL_SECONDS:.2f}s/request")
    print(f"Write pacing     : {MIN_WRITE_INTERVAL_SECONDS:.2f}s/request")
    print(f"429 wait         : {QUOTA_429_WAIT_SECONDS:.0f}s")

    if metric_start or metric_end:
        print(
            "Metric period    : "
            f"{metric_start.isoformat() if metric_start else '-inf'} -> "
            f"{metric_end.isoformat() if metric_end else '+inf'}"
        )
    else:
        print("Metric period    : ALL TAP/AFF source rows")

    if range_start or range_end:
        print(
            "SC RANGE         : "
            f"{range_start.isoformat() if range_start else '-inf'} -> "
            f"{range_end.isoformat() if range_end else '+inf'}"
        )
    else:
        print("SC RANGE         : same as ALL")

    # -------------------------------------------------------------------------
    # Read SC
    # -------------------------------------------------------------------------
    sc_header, sc_rows, sc_meta = detect_and_read_source(
        service,
        spreadsheet_id=SC_SPREADSHEET_ID,
        sheet_gid=SC_SHEET_GID,
        sheet_name=SC_SHEET_NAME,
        required_headers=SC_REQUIRED,
        label="SC VIDEO",
    )
    print_source_meta("SC VIDEO", sc_meta)

    # -------------------------------------------------------------------------
    # Read TAP
    # -------------------------------------------------------------------------
    tap_header, tap_rows, tap_meta = detect_and_read_source(
        service,
        spreadsheet_id=TAP_SPREADSHEET_ID,
        sheet_gid=TAP_SHEET_GID,
        sheet_name=TAP_SHEET_NAME,
        required_headers=TAP_REQUIRED,
        label="TAP VIDEO",
    )
    print_source_meta("TAP VIDEO", tap_meta)

    # Video ID TAP untuk existence check SC NOT IN TAP memakai periode
    # METRIC_START_DATE / METRIC_END_DATE yang sama dengan mapping.
    tap_video_ids, tap_existence_stats = build_tap_video_id_set(
        tap_header,
        tap_rows,
        metric_start,
        metric_end,
    )

    # -------------------------------------------------------------------------
    # Read AFF source(s)
    # -------------------------------------------------------------------------
    aff_datasets = []
    for index, source in enumerate(AFF_SOURCES, start=1):
        label = str(source.get("label") or f"AFF {index}")
        header, rows, meta = detect_and_read_source(
            service,
            spreadsheet_id=source["spreadsheet_id"],
            sheet_gid=source.get("sheet_gid"),
            sheet_name=source.get("sheet_name"),
            required_headers=AFF_REQUIRED,
            label=label,
        )
        print_source_meta(label, meta)
        aff_datasets.append((label, header, rows))

    # -------------------------------------------------------------------------
    # Optional Product ID -> Brand fallback mapping
    # -------------------------------------------------------------------------
    product_brand_mapping = {}
    if USE_PRODUCT_ID_MAPPING_FALLBACK:
        product_brand_mapping, product_brand_mapping_stats = (
            load_product_brand_mapping(service)
        )
        print("\nProduct ID Brand fallback: ON")
        print(f"  Sheet             : {product_brand_mapping_stats['sheet_title']}")
        print(f"  Header row        : {product_brand_mapping_stats['header_row']}")
        print(f"  PC Product ID     : {product_brand_mapping_stats['mapping_count']:,}")
        print(f"  Duplicate same    : {product_brand_mapping_stats['duplicate_same']:,}")
        print(f"  Blank PID         : {product_brand_mapping_stats['blank_pid']:,}")
        print(f"  Non-PC skipped    : {product_brand_mapping_stats['skipped_non_pc']:,}")
    else:
        print("\nProduct ID Brand fallback: OFF")

    # -------------------------------------------------------------------------
    # Mapping
    # -------------------------------------------------------------------------
    print("\nMembangun TAP mapping...")
    tap_map, tap_stats = build_tap_mapping(
        tap_header,
        tap_rows,
        metric_start,
        metric_end,
    )

    print("Membangun Affiliate Order mapping...")
    aff_map, aff_stats = build_aff_mapping(
        aff_datasets,
        metric_start,
        metric_end,
    )

    print("Menggabungkan TAP + AFF menjadi master mapping...")
    master_map, master_counts = build_master_mapping(
        tap_map,
        aff_map,
        product_brand_mapping,
    )
    pid_product_name_lookup = build_pid_product_name_lookup(tap_map, aff_map)

    tap_video_brand_pairs = build_tap_brand_pair_set_from_mapping(
        tap_map,
        pid_product_name_lookup,
        product_brand_mapping,
    )

    # -------------------------------------------------------------------------
    # SC explode + enrichment
    # -------------------------------------------------------------------------
    print("Explode + enrich SC Video...")
    (
        all_rows,
        sc_stats,
        original_totals,
        enriched_totals,
        brand_counter,
        missing_details,
    ) = enrich_sc(
        sc_header,
        sc_rows,
        master_map,
        aff_map,
        pid_product_name_lookup,
        product_brand_mapping,
    )

    range_rows, invalid_range_dates = filter_output_range(
        all_rows,
        range_start,
        range_end,
    )

    # -------------------------------------------------------------------------
    # SC yang Video ID-nya tidak tersedia di TAP.Post time pada periode METRIC_*.
    # Compare hanya berdasarkan Video ID. Jika satu Video ID tidak ada di TAP
    # dalam periode tersebut, seluruh row PID hasil explode untuk video itu
    # ikut masuk sheet ini.
    # -------------------------------------------------------------------------
    sc_not_in_tap_all_rows = filter_sc_not_in_tap(
        all_rows,
        tap_video_brand_pairs,
    )
    sc_not_in_tap_range_rows = filter_sc_not_in_tap(
        range_rows,
        tap_video_brand_pairs,
    )

    # -------------------------------------------------------------------------
    # BRAND UNIQUE: first PID per Video ID + Brand.
    # NOT IN TAP memakai grain Video ID + Brand.
    # -------------------------------------------------------------------------
    brand_unique_all_rows = build_brand_unique_rows(all_rows)
    brand_unique_range_rows = build_brand_unique_rows(range_rows)
    brand_unique_not_in_tap_all_rows = build_brand_unique_rows(
        sc_not_in_tap_all_rows
    )
    brand_unique_not_in_tap_range_rows = build_brand_unique_rows(
        sc_not_in_tap_range_rows
    )

    # -------------------------------------------------------------------------
    # Build output tables
    # -------------------------------------------------------------------------
    output_tables = build_output_tables(
        all_rows,
        range_rows,
        sc_not_in_tap_all_rows,
        sc_not_in_tap_range_rows,
        missing_details,
        WRITE_BRAND_TABS_EFFECTIVE,
    )

    output_tables[BRAND_UNIQUE_ALL_SHEET] = data_values(brand_unique_all_rows)
    output_tables[BRAND_UNIQUE_RANGE_SHEET] = data_values(brand_unique_range_rows)
    output_tables[BRAND_UNIQUE_NOT_IN_TAP_ALL_SHEET] = data_values(
        brand_unique_not_in_tap_all_rows
    )
    output_tables[BRAND_UNIQUE_NOT_IN_TAP_RANGE_SHEET] = data_values(
        brand_unique_not_in_tap_range_rows
    )

    # Summary brand mengikuti grain final Video ID + Brand.
    output_tables[SUMMARY_SHEET] = build_brand_summary(
        brand_unique_all_rows,
        brand_unique_range_rows,
    )

    # Capacity / clash check dilakukan bahkan saat DRY RUN agar aman sebelum write.
    output_info = validate_output_capacity(service, output_tables)

    print("\nBRAND UNIQUE OUTPUT (FIRST PID per Video ID + Brand):")
    print(f"  {BRAND_UNIQUE_ALL_SHEET:<46}: {len(brand_unique_all_rows):>10,} row")
    print(f"  {BRAND_UNIQUE_RANGE_SHEET:<46}: {len(brand_unique_range_rows):>10,} row")
    print(f"  {BRAND_UNIQUE_NOT_IN_TAP_ALL_SHEET:<46}: {len(brand_unique_not_in_tap_all_rows):>10,} row")
    print(f"  {BRAND_UNIQUE_NOT_IN_TAP_RANGE_SHEET:<46}: {len(brand_unique_not_in_tap_range_rows):>10,} row")
    print("  NOT IN TAP key: Video ID + Brand.")
    print("  Brand GMV/Orders/Items: SUM unique Product ID per Video ID + Brand.")

    print_summary(
        tap_stats,
        aff_stats,
        master_counts,
        sc_stats,
        original_totals,
        enriched_totals,
        brand_counter,
        all_rows,
        range_rows,
        sc_not_in_tap_all_rows,
        sc_not_in_tap_range_rows,
        tap_video_ids,
        tap_existence_stats,
        invalid_range_dates,
        range_start,
        range_end,
        WRITE_BRAND_TABS_EFFECTIVE,
        missing_details,
    )

    if not WRITE:
        print("\nDRY RUN selesai. Tidak ada data ditulis.")
        print("Jika QC sudah sesuai:")
        print("  python sc_pc_video_exploded_and_brand_unique.py --write")
        print("Jika tab output lama memang mau diganti:")
        print("  python sc_pc_video_exploded_and_brand_unique.py --write --replace")
        return

    print("\nMenulis output ke Google Sheets...")
    write_output(service, output_tables, output_info)
    print("\nSELESAI.")


if __name__ == "__main__":
    raise SystemExit("Jalankan melalui Streamlit dengan perintah streamlit run app.py.")
