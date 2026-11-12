"""
SC PC LIVE / LIVESTREAM - EXPLODED + BRAND UNIQUE + TAP LS/AFF ENRICHMENT
============================================================================

Flow utama
----------
1. RAW SC LIVE dibaca dari Google Sheets.
2. Product ID SC yang comma-separated di-explode menjadi 1 row per Product ID.
3. Dari hasil exploded dibuat BRAND UNIQUE:
       1 row = LIVE ID + Brand
       Product ID = FIRST PID pada LIVE ID + Brand
       GMV / Orders / Items = SUM metric UNIQUE PID dalam LIVE ID + Brand.
4. RAW TAP Livestream dibuat mapping dengan key:
       Livestream room ID + Product ID
4. RAW Affiliate Order dibuat mapping dengan key:
       Content ID + Product ID
   hanya untuk Content Type = Livestream.
5. MASTER mapping = union TAP LS + AFF LS dengan prioritas metric:
       TAP jika exact pair tersedia, AFF sebagai fallback.
6. Model A:
   - Single Product ID:
       GMV / Orders / Items tetap dari SC original.
   - Multi Product ID:
       GMV / Orders / Items memakai direct mapping TAP/AFF per pair.
       Pair yang tidak ditemukan = 0 agar metric level-LIVE tidak terduplikasi.
7. AOV selalu dihitung ulang:
       mapped GMV / mapped Orders
8. Affiliate columns selalu di-lookup dari AFF:
       GMV Affiliate Order (Est. Commission Base)
       Jumlah Order Affiliate
9. Product name:
       exact LIVE ID + Product ID dari TAP/AFF
       -> fallback Product ID global dari TAP/AFF
       -> fallback Product name SC jika source suatu saat menyediakannya.
10. Brand diturunkan dari Product name memakai rule parent-brand TAP PC:
       CLOSE UP, DOVE DEO, DOVE SCL, LIFEBUOY, LUX, PEPSODENT, REXONA.
    Jika hasil Product name = MISSING dan USE_PRODUCT_ID_MAPPING_FALLBACK=True,
    Brand diambil dari mapping Product ID. Product ID tidak meng-override
    Brand yang sudah berhasil didapat dari Product name.
11. Periode:
    - TAP mapping metric      : berdasarkan START dari `LIVE time info`
    - AFF mapping metric      : berdasarkan `Time Created`
    - SC RANGE                : berdasarkan `LIVE start time`
    - SC LS NOT IN TAP       : compare LIVE ID SC vs Livestream room ID TAP
                               yang START `LIVE time info` berada di METRIC range.
12. Output utama:
    - GABUNGAN SC LS - ALL
    - GABUNGAN SC LS - RANGE
    - SC LS NOT IN TAP - ALL
    - SC LS NOT IN TAP - RANGE
    - MAPPING QC LS - ALL
    - MAPPING QC LS - RANGE
    - BRAND SUMMARY
    - MISSING MAPPING
    - optional per-brand ALL/RANGE tabs.

Run
---
Dry run:
    python sc_pc_live_exploded_and_brand_unique.py

Write:
    python sc_pc_live_exploded_and_brand_unique.py --write

Replace managed output tabs:
    python sc_pc_live_exploded_and_brand_unique.py --write --replace

Brand tabs:
    python sc_pc_live_exploded_and_brand_unique.py --write --brand-tabs
    python sc_pc_live_exploded_and_brand_unique.py --write --no-brand-tabs

Dependency
----------
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

CREDENTIALS_CANDIDATES = [
    BASE_DIR / "credentials.json",
    BASE_DIR / "credentials_sc.json",
]
CREDENTIALS_FILE = next(
    (path for path in CREDENTIALS_CANDIDATES if path.exists()),
    CREDENTIALS_CANDIDATES[0],
)

# RAW SC LIVE / Seller Center Livestream.
SC_LS_SPREADSHEET_ID = "1NJQ1PyWzfIy9XIY98Q4mjEcaXVuflKQqUqstlKZZSj4"
SC_LS_SHEET_GID = None
SC_LS_SHEET_NAME = None

# RAW TAP Livestream.
TAP_LS_SPREADSHEET_ID = "1tbwvLUak4V43lblk1fuetSrC36omccf8dyhByzTVdU0"
TAP_LS_SHEET_GID = None
TAP_LS_SHEET_NAME = "Custom report"

# AFF sama dengan flow SC Video sebelumnya.
AFF_SOURCES = [
    {
        "spreadsheet_id": "1DnLIBPf9Ufgn6Q0-SYz__9m5-wwptki-g_9iGFWbkFg",
        "sheet_gid": None,
        "sheet_name": None,
        "label": "AFF 1",
    },
    # {
    #     "spreadsheet_id": "1DnLIBPf9Ufgn6Q0-SYz__9m5-wwptki-g_9iGFWbkFg",
    #     "sheet_gid": None,
    #     "sheet_name": None,
    #     "label": "AFF 2",
    # },
]

OUTPUT_SPREADSHEET_ID = "ISI_BACKEND_DISABLED_NO_DIRECT_WRITE"
# Optional Product ID -> Brand fallback.
# Product name tetap PRIMARY. Mapping Product ID hanya dipakai ketika
# Product name tidak menghasilkan salah satu Brand PC dan berakhir MISSING.
USE_PRODUCT_ID_MAPPING_FALLBACK = True
MAPPING_SPREADSHEET_ID = "1O4dQVFYawbirEN5wpV0ygjNZtNFZPZH0tLX_Ov34fNA"
MAPPING_SHEET_GID = 437523491
MAPPING_PRODUCT_ID_HEADER = "Product ID TTS"
MAPPING_BRAND_HEADER = "BRAND FIX"

# Metric mapping period.
# TAP LS  -> START dari LIVE time info.
# AFF     -> Time Created.
METRIC_START_DATE = "2026-10-07"
METRIC_END_DATE = "2026-10-07"

# SC output RANGE -> LIVE start time.
RANGE_START_DATE = "2026-10-07"
RANGE_END_DATE = "2026-10-07"

WRITE_BRAND_TABS = False

SOURCE_SCAN_RANGE = "A1:AZ20"
WRITE_CHUNK_ROWS = 1500
MAX_CELLS = 20_000_000
MAX_RETRIES = 7
RETRYABLE_HTTP_STATUS = {429, 500, 502, 503, 504}

# Google Sheets API quota guard.
MIN_READ_INTERVAL_SECONDS = 1.10
MIN_WRITE_INTERVAL_SECONDS = 1.10
QUOTA_429_WAIT_SECONDS = 65.0

SHEET_ALL = "GABUNGAN SC PC LS EXPLODED - ALL"
SHEET_RANGE = "GABUNGAN SC PC LS EXPLODED - RANGE"
SC_NOT_IN_TAP_ALL_SHEET = "SC PC LS EXPLODED NOT IN TAP - ALL"
SC_NOT_IN_TAP_RANGE_SHEET = "SC PC LS EXPLODED NOT IN TAP - RANGE"
BRAND_UNIQUE_ALL_SHEET = "GABUNGAN SC PC LS BRAND UNIQUE - ALL"
BRAND_UNIQUE_RANGE_SHEET = "GABUNGAN SC PC LS BRAND UNIQUE - RANGE"
BRAND_UNIQUE_NOT_IN_TAP_ALL_SHEET = "SC PC LS BRAND UNIQUE NOT IN TAP - ALL"
BRAND_UNIQUE_NOT_IN_TAP_RANGE_SHEET = "SC PC LS BRAND UNIQUE NOT IN TAP - RANGE"
MAPPING_QC_ALL_SHEET = "MAPPING QC PC LS - ALL"
MAPPING_QC_RANGE_SHEET = "MAPPING QC PC LS - RANGE"
SUMMARY_SHEET = "PC BRAND SUMMARY LS"
MISSING_SHEET = "PC LS MISSING MAPPING"
BRAND_ALL_SUFFIX = " - SC PC LS ALL"
BRAND_RANGE_SUFFIX = " - SC PC LS RANGE"

CLEANING_VERSION = "sc-live-model-a-brand-pid-fallback-header-alias-quota-safe"


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
# OUTPUT SCHEMA
# =============================================================================

OUTPUT_COLUMNS = [
    "LIVE title",
    "LIVE ID",
    "LIVE start time",
    "LIVE end time",
    "Creator name",
    "Product ID",
    "Product name",
    "Brand",
    "Creator LIVE-attributed GMV",
    "LIVE-attributed items sold",
    "Refunds",
    "LIVE-attributed orders",
    "AOV",
    "Avg. viewing duration",
    "Likes",
    "Comments",
    "Shares",
    "LIVE product impressions",
    "Product clicks",
    "Impressions",
    "Avg. daily GPM",
    "Avg. daily products sold",
    "Avg. GMV per customer",
    "Avg. daily engagement",
    "Avg. daily LIVE product CTR",
    "Avg. daily unique viewers",
    "Avg. daily tap-through rate",
    "Est. commission",
    "GMV Affiliate Order (Est. Commission Base)",
    "Jumlah Order Affiliate",
]


# -----------------------------------------------------------------------------
# SC LIVE SOURCE HEADER ALIASES
# -----------------------------------------------------------------------------
# Output schema tetap memakai nama standar agar downstream konsisten.
# Seller Center export terbaru memakai beberapa nama kolom yang lebih pendek.
SC_SOURCE_ALIASES = {
    "Avg. daily GPM": [
        "Avg. daily GPM",
        "Show GPM",
    ],
    "Avg. daily products sold": [
        "Avg. daily products sold",
        "Products sold",
    ],
    "Avg. daily engagement": [
        "Avg. daily engagement",
        "Engagement",
    ],
    "Avg. daily LIVE product CTR": [
        "Avg. daily LIVE product CTR",
        "CTR",
    ],
    "Avg. daily unique viewers": [
        "Avg. daily unique viewers",
        "LIVE product viewers",
    ],
    "Avg. daily tap-through rate": [
        "Avg. daily tap-through rate",
        "Tap-through rate",
    ],
}

SC_REQUIRED = [
    "LIVE title",
    "LIVE ID",
    "LIVE start time",
    "Product ID",
    "Creator LIVE-attributed GMV",
    "LIVE-attributed items sold",
    "LIVE-attributed orders",
]

TAP_REQUIRED = [
    "Livestream room ID",
    "Product ID",
    "Product name",
    "LIVE time info",
    "Creator LIVE-attributed GMV",
    "Creator LIVE-attributed orders",
    "Creator-attributed items sold",
]

AFF_REQUIRED = [
    "Order ID",
    "Product ID",
    "Content Type",
    "Content ID",
    "Est. Commission Base",
    "Quantity",
    "Time Created",
]

INTEGER_COLUMNS = {
    "LIVE-attributed items sold",
    "LIVE-attributed orders",
    "Likes",
    "Comments",
    "Shares",
    "LIVE product impressions",
    "Product clicks",
    "Impressions",
    "Avg. daily products sold",
    "Avg. daily unique viewers",
}

NUMERIC_COLUMNS = {
    "Creator LIVE-attributed GMV",
    "Refunds",
    "AOV",
    "Avg. daily GPM",
    "Avg. GMV per customer",
    "Est. commission",
}

PERCENT_COLUMNS = {
    "Avg. daily engagement",
    "Avg. daily LIVE product CTR",
    "Avg. daily tap-through rate",
}


# =============================================================================
# BRAND RULES - SAME AS TAP PC LS/SV
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
            "Simpan ID di Google Sheets sebagai text / digit utuh."
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

    s = re.sub(r"^(?:RP|IDR)\s*", "", s, flags=re.I).replace(" ", "")
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
    elif "." in s and "," in s:
        if s.rfind(".") > s.rfind(","):
            s = s.replace(",", "")
        else:
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


def parse_sc_live_datetime(value: object) -> datetime | None:
    # SC LIVE export = MDY.
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
    ):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            pass

    raise ValueError(f"SC LIVE datetime tidak valid: {value!r}")


LIVE_TIME_RE = re.compile(
    r"^\s*"
    r"(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2})"
    r"\s*-\s*"
    r"(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2})"
    r"\s*$"
)


def parse_tap_live_time_info(value: object) -> tuple[datetime, datetime] | None:
    s = str(value or "").strip()
    if not s or s in {"-", "--"}:
        return None

    match = LIVE_TIME_RE.fullmatch(s)
    if not match:
        raise ValueError(f"TAP LIVE time info tidak valid: {value!r}")

    start_dt = datetime.strptime(match.group(1), "%Y-%m-%d %H:%M:%S")
    end_dt = datetime.strptime(match.group(2), "%Y-%m-%d %H:%M:%S")
    return start_dt, end_dt


def parse_aff_datetime(value: object) -> datetime | None:
    # AFF export = DD/MM/YYYY.
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


def normalize_sc_live_datetime(value: object) -> str:
    try:
        dt = parse_sc_live_datetime(value)
    except ValueError:
        return str(value or "").strip()
    return dt.strftime("%Y-%m-%d %H:%M:%S") if dt else ""


def parse_output_live_date(value: object) -> date | None:
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


def rows_to_dicts(
    header: list[str],
    rows: list[list[object]],
) -> list[dict[str, object]]:
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
    method = str(getattr(request, "method", "") or "").upper()
    return "READ" if method == "GET" else "WRITE"


def _pace_google_request(request) -> None:
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
    """Detect/read source with minimal Google read requests."""
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

    # Known tab: one full values.get only.
    if len(candidates) == 1 and (
        sheet_name is not None
        or sheet_gid is not None
    ):
        prop = candidates[0]
        title = prop["title"]
        grid = prop.get("gridProperties", {}) or {}
        row_count = max(
            int(grid.get("rowCount", 1) or 1),
            1,
        )
        col_count = max(
            int(grid.get("columnCount", 1) or 1),
            1,
        )
        end_col = column_letter(col_count)

        values = execute_with_retry(
            service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=f"'{title}'!A1:{end_col}{row_count}",
                valueRenderOption="FORMATTED_VALUE",
            )
        ).get("values", [])

        if not values:
            raise ValueError(
                f'{label}: tab "{title}" kosong.'
            )

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

    # Auto-detect: scan only until matching tab, then one full read.
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
        row_count = max(
            int(grid.get("rowCount", 1) or 1),
            header_row_number,
        )

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
            raise ValueError(
                f'{label}: tab "{title}" kosong.'
            )

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
# BRAND CLASSIFICATION
# =============================================================================

def find_brand_mentions(product_name: object):
    text = str(product_name or "")
    mentions = []

    for base_brand, pattern in BASE_BRAND_PATTERNS.items():
        match = pattern.search(text)
        if match:
            mentions.append(
                (match.start(), match.end(), base_brand, match.group(0))
            )

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

    return (
        "DOVE DEO"
        if DOVE_DEO_RE.search(parent_segment)
        else "DOVE SCL"
    )


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
    """Resolve Brand with Product name as primary and Product ID as fallback."""
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
    """Load Product ID -> canonical PC Brand for fallback only."""
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
# TAP LS MAPPING
# =============================================================================

def build_tap_ls_mapping(
    header: list[str],
    rows: list[dict[str, object]],
    metric_start: date | None,
    metric_end: date | None,
):
    validate_headers(header, TAP_REQUIRED, "TAP LS")

    h_live = resolve_header(header, "Livestream room ID")
    h_pid = resolve_header(header, "Product ID")
    h_name = resolve_header(header, "Product name")
    h_live_time = resolve_header(header, "LIVE time info")
    h_gmv = resolve_header(header, "Creator LIVE-attributed GMV")
    h_orders = resolve_header(header, "Creator LIVE-attributed orders")
    h_items = resolve_header(header, "Creator-attributed items sold")

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
    invalid_live_time = 0
    period_skipped = 0

    for row in rows:
        live_id = normalize_id(row.get(h_live, ""))
        product_id = normalize_id(row.get(h_pid, ""))

        if not live_id or not product_id:
            invalid_key += 1
            continue

        if metric_start or metric_end:
            try:
                live_period = parse_tap_live_time_info(row.get(h_live_time, ""))
            except ValueError:
                invalid_live_time += 1
                continue

            live_start = live_period[0].date() if live_period else None
            if not in_period(live_start, metric_start, metric_end):
                period_skipped += 1
                continue

        key = (live_id, product_id)
        rec = mapping[key]
        rec["gmv"] += parse_number(row.get(h_gmv, ""))
        rec["orders"] += parse_int(row.get(h_orders, ""))
        rec["items"] += parse_int(row.get(h_items, ""))
        rec["rows"] += 1

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
        "invalid_live_time": invalid_live_time,
        "period_skipped": period_skipped,
    }
    return finalized, stats


def build_tap_ls_live_id_set(
    header: list[str],
    rows: list[dict[str, object]],
    metric_start: date | None,
    metric_end: date | None,
):
    """
    Existence set untuk SC LS NOT IN TAP.

    TAP difilter berdasarkan START `LIVE time info`, bukan kolom Date.
    """
    h_live = resolve_header(header, "Livestream room ID")
    h_live_time = resolve_header(header, "LIVE time info")

    live_ids = set()
    stats = {
        "raw_rows": len(rows),
        "included_rows": 0,
        "blank_or_invalid_live_id": 0,
        "invalid_live_time": 0,
        "period_skipped": 0,
    }

    for row in rows:
        if metric_start or metric_end:
            try:
                live_period = parse_tap_live_time_info(row.get(h_live_time, ""))
            except ValueError:
                stats["invalid_live_time"] += 1
                continue

            live_start = live_period[0].date() if live_period else None
            if not in_period(live_start, metric_start, metric_end):
                stats["period_skipped"] += 1
                continue

        live_id = normalize_id(row.get(h_live, ""))
        if not live_id:
            stats["blank_or_invalid_live_id"] += 1
            continue

        stats["included_rows"] += 1
        live_ids.add(live_id)

    return live_ids, stats


# =============================================================================
# AFF LIVESTREAM MAPPING
# =============================================================================

def build_aff_ls_mapping(
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
    non_livestream = 0
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
        h_time = resolve_header(header, "Time Created")

        total_raw += len(rows)

        for row in rows:
            content_type = str(row.get(h_type, "") or "").strip().lower()
            if content_type != "livestream":
                non_livestream += 1
                continue

            live_id = normalize_id(row.get(h_content, ""))
            product_id = normalize_id(row.get(h_pid, ""))
            _order_id_raw = str(row.get(h_order, "") or "").strip()

            if not live_id or not product_id:
                invalid_key += 1
                continue

            if metric_start or metric_end:
                try:
                    dt = parse_aff_datetime(row.get(h_time, ""))
                except ValueError:
                    invalid_date += 1
                    continue

                d = dt.date() if dt else None
                if not in_period(d, metric_start, metric_end):
                    period_skipped += 1
                    continue

            fingerprint = tuple(
                str(row.get(name, "") or "").strip()
                for name in header
            )

            if fingerprint in seen_fingerprints:
                duplicate_rows += 1
                continue

            seen_fingerprints.add(fingerprint)

            key = (live_id, product_id)
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
        "non_livestream": non_livestream,
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

        # Model A: TAP metric prioritas, AFF fallback.
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


def build_pid_product_name_lookup(
    tap_map: dict,
    aff_map: dict,
) -> dict[str, str]:
    tap_names = defaultdict(Counter)
    aff_names = defaultdict(Counter)

    for (_live_id, pid), rec in tap_map.items():
        name = str(rec.get("product_name", "") or "").strip()
        if name:
            tap_names[pid][name] += 1

    for (_live_id, pid), rec in aff_map.items():
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
    """Build existence key TAP pada grain LIVE ID + Brand.

    Hanya pair yang benar-benar ada di TAP map yang menjadi evidence existence.
    Product name boleh fallback via PID lookup hanya untuk menentukan Brand.
    """
    pairs: set[tuple[str, str]] = set()
    for (live_id, pid), rec in tap_map.items():
        live_id = normalize_id(live_id)
        if not live_id:
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
        pairs.add((live_id, brand))
    return pairs


# =============================================================================
# SC LIVE ENRICHMENT
# =============================================================================

def find_sc_source_column(
    sc_header: list[str],
    output_name: str,
) -> str | None:
    candidates = SC_SOURCE_ALIASES.get(
        output_name,
        [output_name],
    )
    return resolve_first_header(
        sc_header,
        candidates,
    )


OPTIONAL_CONTENT_METRICS = ('Likes',)


def metric_missing(value):
    return value is None or str(value).strip().upper() in {"", "-", "--", "N/A", "NA", "NULL", "NONE", "NAN"}


def clean_sc_value(output_name: str, value: object):
    if output_name in OPTIONAL_CONTENT_METRICS and metric_missing(value):
        return ""

    if output_name in INTEGER_COLUMNS:
        return parse_int(value)

    if output_name in NUMERIC_COLUMNS:
        return clean_numeric(parse_number(value))

    if output_name in PERCENT_COLUMNS:
        return clean_numeric(parse_percent(value))

    if output_name == "LIVE ID":
        return normalize_id(value)

    if output_name in {"LIVE start time", "LIVE end time"}:
        return normalize_sc_live_datetime(value)

    return "" if value is None else str(value).strip()


def enrich_sc_live(
    header: list[str],
    rows: list[dict[str, object]],
    master_map: dict,
    aff_map: dict,
    pid_product_name_lookup: dict[str, str],
    product_brand_mapping: dict[str, str] | None = None,
):
    validate_headers(header, SC_REQUIRED, "SC LIVE")

    derived_columns = {
        "Product name",
        "Brand",
        "GMV Affiliate Order (Est. Commission Base)",
        "Jumlah Order Affiliate",
    }

    source_cols = {
        out_name: find_sc_source_column(header, out_name)
        for out_name in OUTPUT_COLUMNS
        if out_name not in derived_columns
    }

    required_output_sources = [
        name
        for name in OUTPUT_COLUMNS
        if name not in derived_columns
    ]

    missing_sources = [
        name
        for name in required_output_sources
        if source_cols.get(name) is None
    ]

    if missing_sources:
        alias_help = []
        for name in missing_sources:
            candidates = SC_SOURCE_ALIASES.get(name, [name])
            alias_help.append(
                f"{name} <- {candidates}"
            )

        raise ValueError(
            "SC LIVE tidak memiliki source untuk kolom output: "
            + ", ".join(missing_sources)
            + "\nAlias yang dicoba:\n  - "
            + "\n  - ".join(alias_help)
        )

    h_live = resolve_header(header, "LIVE ID")
    h_pid = resolve_header(header, "Product ID")
    h_sc_name = resolve_first_header(header, ["Product name", "Product Name"])
    h_sc_gmv = resolve_header(header, "Creator LIVE-attributed GMV")
    h_sc_orders = resolve_header(header, "LIVE-attributed orders")
    h_sc_items = resolve_header(header, "LIVE-attributed items sold")

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
        "invalid_live_id": 0,
        "brand_from_product_name": 0,
        "brand_from_product_id_fallback": 0,
        "brand_missing": 0,
    }

    original_totals = {
        "gmv": 0.0,
        "orders": 0,
        "items": 0,
    }

    enriched_totals = {
        "gmv": 0.0,
        "orders": 0,
        "items": 0,
    }

    brand_counter = Counter()

    for row in rows:
        live_id = normalize_id(row.get(h_live, ""))

        if not live_id:
            stats["invalid_live_id"] += 1
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
            if h_sc_name
            else ""
        )

        source_title_col = source_cols.get("LIVE title")
        example_title = (
            str(row.get(source_title_col, "") or "").strip()
            if source_title_col
            else ""
        )

        for pid in product_ids:
            key = (live_id, pid)
            master = master_map.get(key)
            aff = aff_map.get(key)

            out = {}

            for out_name in OUTPUT_COLUMNS:
                if out_name in derived_columns:
                    continue

                if out_name == "Product ID":
                    out[out_name] = pid
                    continue

                source_name = source_cols.get(out_name)
                raw_value = row.get(source_name, "") if source_name else ""
                out[out_name] = clean_sc_value(out_name, raw_value)

            out["LIVE ID"] = live_id

            mapped_product_name = ""

            if master is not None:
                mapped_product_name = str(
                    master.get("product_name", "") or ""
                ).strip()

            if not mapped_product_name:
                mapped_product_name = str(
                    pid_product_name_lookup.get(pid, "") or ""
                ).strip()

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

            out["Product name"] = mapped_product_name
            out["Brand"] = brand
            out["_BRAND_SOURCE"] = brand_source

            # -------------------------------------------------------------
            # MODEL A
            # -------------------------------------------------------------
            if len(product_ids) == 1:
                mapped_gmv = sc_original_gmv
                mapped_orders = sc_original_orders
                mapped_items = sc_original_items
                stats["single_pid_sc_preserved"] += 1

            elif master is not None:
                mapped_gmv = float(master["mapped_gmv"])
                mapped_orders = int(master["mapped_orders"])
                mapped_items = int(master["mapped_items"])

                if master["metric_source"] == "TAP":
                    stats["multi_pid_mapped_tap"] += 1
                else:
                    stats["multi_pid_mapped_aff_only"] += 1

            else:
                mapped_gmv = 0.0
                mapped_orders = 0
                mapped_items = 0
                stats["unmapped_multi_pid"] += 1

                detail = missing_details[key]
                detail["count"] += 1
                detail["source_product_count"] = len(product_ids)

                if not detail["example_title"]:
                    detail["example_title"] = example_title

            mapped_aov = (
                mapped_gmv / mapped_orders
                if mapped_orders
                else 0.0
            )

            out["Creator LIVE-attributed GMV"] = clean_numeric(
                mapped_gmv
            )
            out["LIVE-attributed orders"] = mapped_orders
            out["AOV"] = clean_numeric(mapped_aov)
            out["LIVE-attributed items sold"] = mapped_items

            out[
                "GMV Affiliate Order (Est. Commission Base)"
            ] = clean_numeric(
                float(aff["gmv"])
                if aff is not None
                else 0.0
            )

            out["Jumlah Order Affiliate"] = (
                int(aff["orders"])
                if aff is not None
                else 0
            )

            out["_BRAND"] = brand

            # Helper QC exact mapping TAP+AFF.
            # Ini sengaja dipisahkan dari _MAPPING_SOURCE karena single-PID
            # tetap preserve metric SC walaupun exact pair mungkin ada/tidak ada.
            if master is not None:
                out["_EXACT_MAPPING_STATUS"] = master["match_status"]
                out["_EXACT_MAPPING_SOURCE"] = master["metric_source"]
            else:
                out["_EXACT_MAPPING_STATUS"] = "MISSING"
                out["_EXACT_MAPPING_SOURCE"] = "MISSING"

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
        d = parse_output_live_date(
            row.get("LIVE start time", "")
        )

        if d is None:
            invalid_date += 1
            continue

        if in_period(d, range_start, range_end):
            filtered.append(row)

    return filtered, invalid_date


def filter_sc_ls_not_in_tap(
    rows: list[dict[str, object]],
    tap_live_brand_pairs: set[tuple[str, str]],
):
    """NOT IN TAP pada grain LIVE ID + Brand."""
    result = []

    for row in rows:
        live_id = normalize_id(row.get("LIVE ID", ""))
        brand = str(
            row.get("_BRAND", row.get("Brand", MISSING_BRAND))
            or MISSING_BRAND
        ).strip()

        if live_id and (live_id, brand) not in tap_live_brand_pairs:
            result.append(row)

    return result


def build_brand_unique_rows(
    rows: list[dict[str, object]],
) -> list[dict[str, object]]:
    """
    Ringkas exploded SC LIVE menjadi grain unik LIVE ID + Brand.

    - Product ID output = PID pertama yang muncul pada LIVE ID + Brand.
    - GMV / Orders / Items = SUM metric per Product ID unik pada brand tersebut.
    - PID yang sama tidak boleh menyumbang metric dua kali.
    - Product name mengikuti representative/first PID.
    - Metric LIVE-level lain mengikuti row representative dan jangan dijumlahkan lintas Brand.
    """
    grouped: dict[tuple[str, str], dict[str, object]] = {}
    key_order: list[tuple[str, str]] = []

    for row in rows:
        live_id = normalize_id(row.get("LIVE ID", ""))
        if not live_id:
            continue

        brand = str(
            row.get("_BRAND", row.get("Brand", MISSING_BRAND))
            or MISSING_BRAND
        ).strip() or MISSING_BRAND
        key = (live_id, brand)
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
                row.get("Creator LIVE-attributed GMV", 0)
            )
            rec["orders"] += parse_int(
                row.get("LIVE-attributed orders", 0)
            )
            rec["items"] += parse_int(
                row.get("LIVE-attributed items sold", 0)
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

        out["Creator LIVE-attributed GMV"] = clean_numeric(gmv)
        out["LIVE-attributed orders"] = orders
        out["LIVE-attributed items sold"] = items
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
            normalize_id(row.get("LIVE ID", "")),
            str(row.get("_BRAND", MISSING_BRAND) or MISSING_BRAND),
        )
        if key in seen_keys:
            raise RuntimeError(f"Duplicate BRAND UNIQUE key terdeteksi: {key}")
        seen_keys.add(key)

    return result


def build_brand_tables(
    rows: list[dict[str, object]],
):
    grouped = {
        brand: []
        for brand in BRAND_OUTPUT_ORDER
    }

    for row in rows:
        brand = str(
            row.get("_BRAND", MISSING_BRAND)
            or MISSING_BRAND
        )

        if brand not in grouped:
            brand = MISSING_BRAND

        grouped[brand].append(row)

    return grouped



# =============================================================================
# MAPPING QC PER LIVE ID
# =============================================================================

def build_mapping_qc_values(
    rows: list[dict[str, object]],
) -> list[list[object]]:
    """
    Ringkas coverage exact mapping TAP LS + AFF per LIVE ID.

    MAPPED berarti exact pair:
        LIVE ID + Product ID
    tersedia di MASTER mapping (TAP LS dan/atau AFF Livestream).

    Single-PID yang metric-nya preserve SC tetap dinilai mapped/unmapped
    berdasarkan keberadaan exact pair, bukan berdasarkan metric source final.
    """
    grouped = {}

    for row in rows:
        live_id = normalize_id(row.get("LIVE ID", ""))
        if not live_id:
            continue

        rec = grouped.setdefault(
            live_id,
            {
                "title": str(row.get("LIVE title", "") or "").strip(),
                "start_time": str(row.get("LIVE start time", "") or "").strip(),
                "creator": str(row.get("Creator name", "") or "").strip(),
                "all_pids": set(),
                "mapped_pids": set(),
                "unmapped_pids": set(),
                "both_pids": set(),
                "tap_only_pids": set(),
                "aff_only_pids": set(),
            },
        )

        # Isi metadata pertama yang nonblank.
        if not rec["title"]:
            rec["title"] = str(row.get("LIVE title", "") or "").strip()
        if not rec["start_time"]:
            rec["start_time"] = str(row.get("LIVE start time", "") or "").strip()
        if not rec["creator"]:
            rec["creator"] = str(row.get("Creator name", "") or "").strip()

        pid = normalize_id(row.get("Product ID", ""))
        if not pid:
            continue

        rec["all_pids"].add(pid)

        status = str(
            row.get("_EXACT_MAPPING_STATUS", "MISSING")
            or "MISSING"
        ).strip().upper()

        if status == "BOTH":
            rec["mapped_pids"].add(pid)
            rec["both_pids"].add(pid)
        elif status == "TAP_ONLY":
            rec["mapped_pids"].add(pid)
            rec["tap_only_pids"].add(pid)
        elif status == "AFF_ONLY":
            rec["mapped_pids"].add(pid)
            rec["aff_only_pids"].add(pid)
        else:
            rec["unmapped_pids"].add(pid)

    header = [
        "LIVE ID",
        "LIVE title",
        "LIVE start time",
        "Creator name",
        "PID Count",
        "Mapped PID Count",
        "Unmapped PID Count",
        "Mapping Coverage",
        "BOTH PID Count",
        "TAP Only PID Count",
        "AFF Only PID Count",
        "Mapped Product IDs",
        "Unmapped Product IDs",
        "BOTH Product IDs",
        "TAP Only Product IDs",
        "AFF Only Product IDs",
        "Mapping Status",
    ]

    body = []

    for live_id, rec in grouped.items():
        all_pids = sorted(rec["all_pids"])
        mapped_pids = sorted(rec["mapped_pids"])
        unmapped_pids = sorted(rec["unmapped_pids"])
        both_pids = sorted(rec["both_pids"])
        tap_only_pids = sorted(rec["tap_only_pids"])
        aff_only_pids = sorted(rec["aff_only_pids"])

        total_count = len(all_pids)
        mapped_count = len(mapped_pids)
        unmapped_count = len(unmapped_pids)

        coverage = (
            mapped_count / total_count
            if total_count
            else 0.0
        )

        if total_count == 0:
            mapping_status = "NO PID"
        elif mapped_count == total_count:
            mapping_status = "FULLY MAPPED"
        elif mapped_count == 0:
            mapping_status = "NO MAPPING"
        else:
            mapping_status = "PARTIAL MAPPING"

        body.append([
            live_id,
            rec["title"],
            rec["start_time"],
            rec["creator"],
            total_count,
            mapped_count,
            unmapped_count,
            clean_numeric(coverage),
            len(both_pids),
            len(tap_only_pids),
            len(aff_only_pids),
            ",".join(mapped_pids),
            ",".join(unmapped_pids),
            ",".join(both_pids),
            ",".join(tap_only_pids),
            ",".join(aff_only_pids),
            mapping_status,
        ])

    # Prioritaskan LS yang paling bermasalah di atas:
    # unmapped paling banyak -> PID total paling banyak -> LIVE ID.
    body.sort(
        key=lambda row: (
            -int(row[6]),
            -int(row[4]),
            str(row[0]),
        )
    )

    return [header] + body


def summarize_mapping_qc(
    rows: list[dict[str, object]],
):
    """Summary agregat coverage mapping exact pair untuk console QC."""
    values = build_mapping_qc_values(rows)
    body = values[1:]

    summary = {
        "live_ids": len(body),
        "fully_mapped": 0,
        "partial_mapping": 0,
        "no_mapping": 0,
        "no_pid": 0,
        "pid_total": 0,
        "pid_mapped": 0,
        "pid_unmapped": 0,
    }

    for row in body:
        summary["pid_total"] += int(row[4])
        summary["pid_mapped"] += int(row[5])
        summary["pid_unmapped"] += int(row[6])

        status = str(row[16])
        if status == "FULLY MAPPED":
            summary["fully_mapped"] += 1
        elif status == "PARTIAL MAPPING":
            summary["partial_mapping"] += 1
        elif status == "NO MAPPING":
            summary["no_mapping"] += 1
        else:
            summary["no_pid"] += 1

    summary["pid_coverage"] = (
        summary["pid_mapped"] / summary["pid_total"]
        if summary["pid_total"]
        else 0.0
    )

    return summary


# =============================================================================
# OUTPUT TABLE BUILDERS
# =============================================================================

def data_values(
    rows: list[dict[str, object]],
) -> list[list[object]]:
    return [OUTPUT_COLUMNS] + [
        [row.get(col, "") for col in OUTPUT_COLUMNS]
        for row in rows
    ]


def build_brand_summary(
    all_rows,
    range_rows,
):
    all_counter = Counter(
        str(row.get("_BRAND", MISSING_BRAND))
        for row in all_rows
    )

    range_counter = Counter(
        str(row.get("_BRAND", MISSING_BRAND))
        for row in range_rows
    )

    values = [[
        "Brand",
        "Rows ALL",
        "Rows RANGE",
        "GMV ALL",
        "GMV RANGE",
    ]]

    for brand in BRAND_OUTPUT_ORDER:
        gmv_all = sum(
            parse_number(
                row.get(
                    "Creator LIVE-attributed GMV",
                    0,
                )
            )
            for row in all_rows
            if row.get("_BRAND") == brand
        )

        gmv_range = sum(
            parse_number(
                row.get(
                    "Creator LIVE-attributed GMV",
                    0,
                )
            )
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


def build_missing_values(
    missing_details,
):
    values = [[
        "LIVE ID",
        "Product ID",
        "Occurrences",
        "Product Count in Source Row",
        "Example LIVE Title",
        "Reason",
    ]]

    for (
        live_id,
        product_id,
    ), detail in sorted(missing_details.items()):
        values.append([
            live_id,
            product_id,
            detail["count"],
            detail["source_product_count"],
            detail["example_title"],
            (
                "LIVE ID + Product ID tidak ditemukan "
                "di master TAP LS + AFF Livestream"
            ),
        ])

    return values


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
        SC_NOT_IN_TAP_ALL_SHEET: data_values(
            sc_not_in_tap_all_rows
        ),
        SC_NOT_IN_TAP_RANGE_SHEET: data_values(
            sc_not_in_tap_range_rows
        ),
        MAPPING_QC_ALL_SHEET: build_mapping_qc_values(
            all_rows
        ),
        MAPPING_QC_RANGE_SHEET: build_mapping_qc_values(
            range_rows
        ),
    }

    if write_brand_tabs:
        brand_all = build_brand_tables(all_rows)
        brand_range = build_brand_tables(range_rows)

        for brand in BRAND_OUTPUT_ORDER:
            tables[
                f"{brand}{BRAND_ALL_SUFFIX}"
            ] = data_values(
                brand_all[brand]
            )

        for brand in BRAND_OUTPUT_ORDER:
            tables[
                f"{brand}{BRAND_RANGE_SUFFIX}"
            ] = data_values(
                brand_range[brand]
            )

    tables[SUMMARY_SHEET] = build_brand_summary(
        all_rows,
        range_rows,
    )

    tables[MISSING_SHEET] = build_missing_values(
        missing_details
    )

    return tables


def all_managed_titles():
    titles = {
        SHEET_ALL,
        SHEET_RANGE,
        SC_NOT_IN_TAP_ALL_SHEET,
        SC_NOT_IN_TAP_RANGE_SHEET,
        MAPPING_QC_ALL_SHEET,
        MAPPING_QC_RANGE_SHEET,
        BRAND_UNIQUE_ALL_SHEET,
        BRAND_UNIQUE_RANGE_SHEET,
        BRAND_UNIQUE_NOT_IN_TAP_ALL_SHEET,
        BRAND_UNIQUE_NOT_IN_TAP_RANGE_SHEET,
        SUMMARY_SHEET,
        MISSING_SHEET,
    }

    for brand in BRAND_OUTPUT_ORDER:
        titles.add(
            f"{brand}{BRAND_ALL_SUFFIX}"
        )
        titles.add(
            f"{brand}{BRAND_RANGE_SUFFIX}"
        )

    return titles


# =============================================================================
# OUTPUT GOOGLE SHEETS
# =============================================================================

def get_output_grid_info(service):
    properties = get_sheet_properties(
        service,
        OUTPUT_SPREADSHEET_ID,
    )

    info = {}
    total_cells = 0

    for prop in properties:
        grid = prop.get(
            "gridProperties",
            {},
        )

        rows = int(
            grid.get("rowCount", 0)
            or 0
        )

        cols = int(
            grid.get("columnCount", 0)
            or 0
        )

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

    col_count = max(
        max(
            (len(row) for row in values),
            default=1,
        ),
        1,
    )

    return row_count * col_count


def validate_output_capacity(
    service,
    output_tables,
):
    info, existing_cells = get_output_grid_info(
        service
    )

    requested_titles = set(output_tables)

    managed_existing = (
        set(info)
        & all_managed_titles()
    )

    clashes = (
        requested_titles
        & set(info)
    )

    if clashes and not REPLACE:
        raise SystemExit(
            "Tab output sudah ada dan tidak akan ditimpa: "
            + ", ".join(sorted(clashes))
            + "\nJalankan dengan --write --replace "
              "jika memang mau mengganti output."
        )

    removed_titles = (
        managed_existing
        if REPLACE
        else set()
    )

    cells_removed = sum(
        info[t]["cells"]
        for t in removed_titles
    )

    new_cells = sum(
        cells_for_values(values)
        for values in output_tables.values()
    )

    projected = (
        existing_cells
        - cells_removed
        + new_cells
    )

    print("\nValidasi kapasitas output:")
    print(
        f"  Existing cells       : "
        f"{existing_cells:,}"
    )

    if REPLACE:
        print(
            f"  Managed cells removed: "
            f"-{cells_removed:,}"
        )

    print(
        f"  New output cells     : "
        f"{new_cells:,}"
    )

    print(
        f"  Projected total      : "
        f"{projected:,}"
    )

    print(
        f"  Safety limit         : "
        f"{MAX_CELLS:,}"
    )

    if projected > MAX_CELLS:
        raise SystemExit(
            "Projected output melebihi MAX_CELLS. "
            "Gunakan output spreadsheet lain "
            "atau nonaktifkan brand tabs."
        )

    return info


def create_temp_sheet_if_needed(
    service,
    info,
    delete_titles,
):
    existing_titles = set(info)

    if (
        not delete_titles
        or delete_titles != existing_titles
    ):
        return None

    temp_title = "_TMP_AUTOMATION_DO_NOT_USE"
    suffix = 1

    while temp_title in existing_titles:
        suffix += 1
        temp_title = (
            f"_TMP_AUTOMATION_DO_NOT_USE_"
            f"{suffix}"
        )

    response = execute_with_retry(
        service.spreadsheets().batchUpdate(
            spreadsheetId=OUTPUT_SPREADSHEET_ID,
            body={
                "requests": [
                    {
                        "addSheet": {
                            "properties": {
                                "title": temp_title
                            }
                        }
                    }
                ]
            },
        )
    )

    invalidate_sheet_properties_cache(OUTPUT_SPREADSHEET_ID)

    replies = response.get(
        "replies",
        [],
    )

    if replies:
        return replies[0][
            "addSheet"
        ][
            "properties"
        ][
            "sheetId"
        ]

    return None


def recreate_output_tabs(
    service,
    info,
    output_tables,
):
    delete_titles = (
        set(info)
        & all_managed_titles()
        if REPLACE
        else set()
    )

    temp_sheet_id = (
        create_temp_sheet_if_needed(
            service,
            info,
            delete_titles,
        )
    )

    requests = []

    for title in sorted(delete_titles):
        requests.append({
            "deleteSheet": {
                "sheetId": info[
                    title
                ][
                    "sheetId"
                ]
            }
        })

    for title, values in output_tables.items():
        row_count = max(
            len(values),
            2,
        )

        col_count = max(
            max(
                (
                    len(row)
                    for row in values
                ),
                default=1,
            ),
            1,
        )

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
                body={
                    "requests": [
                        {
                            "deleteSheet": {
                                "sheetId": temp_sheet_id
                            }
                        }
                    ]
                },
            )
        )

    invalidate_sheet_properties_cache(OUTPUT_SPREADSHEET_ID)


def write_values_chunked(
    service,
    title,
    values,
):
    for start in range(
        0,
        len(values),
        WRITE_CHUNK_ROWS,
    ):
        part = values[
            start:
            start + WRITE_CHUNK_ROWS
        ]

        execute_with_retry(
            service.spreadsheets().values().update(
                spreadsheetId=OUTPUT_SPREADSHEET_ID,
                range=(
                    f"'{title}'!"
                    f"A{start + 1}"
                ),
                valueInputOption="RAW",
                body={"values": part},
            )
        )

    print(
        f'  "{title}" selesai '
        f'({max(len(values) - 1, 0):,} row)'
    )


def write_output(
    service,
    output_tables,
    info,
):
    recreate_output_tabs(
        service,
        info,
        output_tables,
    )

    for title, values in output_tables.items():
        write_values_chunked(
            service,
            title,
            values,
        )


# =============================================================================
# DIAGNOSTICS
# =============================================================================

def pct_diff(
    mapped: float,
    original: float,
) -> float | None:
    if math.isclose(
        original,
        0.0,
        abs_tol=1e-12,
    ):
        return None

    return (
        (mapped - original)
        / original
        * 100.0
    )


def fmt_num(
    value: float | int,
) -> str:
    if (
        isinstance(value, int)
        or math.isclose(
            float(value),
            round(float(value)),
            abs_tol=1e-9,
        )
    ):
        return (
            f"{int(round(float(value))):,}"
        )

    return f"{float(value):,.2f}"


def print_source_meta(
    label,
    meta,
):
    print(f"\n{label} source:")
    print(
        f"  Spreadsheet ID : "
        f"{meta['spreadsheet_id']}"
    )
    print(
        f"  Sheet          : "
        f"{meta['sheet_title']}"
    )
    print(
        f"  GID            : "
        f"{meta['sheet_id']}"
    )
    print(
        f"  Header row     : "
        f"{meta['header_row']}"
    )
    print(
        f"  Columns        : "
        f"{meta['columns']}"
    )
    print(
        f"  Rows           : "
        f"{meta['rows']:,}"
    )


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
    tap_live_ids,
    tap_existence_stats,
    invalid_range_dates,
    range_start,
    range_end,
    write_brand_tabs,
    missing_details,
):
    print("\n" + "=" * 78)
    print(
        "SC LIVE - EXPLODE + TAP LS/AFF "
        "PRODUCT MAPPING"
    )
    print("=" * 78)

    print("\nTAP LS mapping:")
    print(
        f"  Raw row                  : "
        f"{tap_stats['raw_rows']:,}"
    )
    print(
        f"  LIVE ID + Product ID     : "
        f"{tap_stats['pairs']:,}"
    )
    print(
        f"  Blank/invalid key        : "
        f"{tap_stats['invalid_key']:,}"
    )
    print(
        f"  Invalid LIVE time info   : "
        f"{tap_stats['invalid_live_time']:,}"
    )
    print(
        f"  Outside LIVE start period: "
        f"{tap_stats['period_skipped']:,}"
    )

    print("\nAffiliate Livestream mapping:")
    print(
        f"  Raw row                  : "
        f"{aff_stats['raw_rows']:,}"
    )
    print(
        f"  Non-Livestream skipped   : "
        f"{aff_stats['non_livestream']:,}"
    )
    print(
        f"  LIVE ID + Product ID     : "
        f"{aff_stats['pairs']:,}"
    )
    print(
        f"  Exact duplicate skipped  : "
        f"{aff_stats['duplicate_rows']:,}"
    )
    print(
        f"  Blank/invalid key        : "
        f"{aff_stats['invalid_key']:,}"
    )
    print(
        f"  Invalid Time Created     : "
        f"{aff_stats['invalid_date']:,}"
    )
    print(
        f"  Outside Time Created     : "
        f"{aff_stats['period_skipped']:,}"
    )

    print("\nMaster TAP LS + AFF mapping:")
    print(
        f"  BOTH                     : "
        f"{master_counts['BOTH']:,}"
    )
    print(
        f"  TAP only                 : "
        f"{master_counts['TAP_ONLY']:,}"
    )
    print(
        f"  AFF only                 : "
        f"{master_counts['AFF_ONLY']:,}"
    )
    print(
        f"  Total pair               : "
        f"{sum(master_counts.values()):,}"
    )

    print("\nSC LIVE explode + enrichment:")
    print(
        f"  Raw SC row               : "
        f"{sc_stats['raw_rows']:,}"
    )
    print(
        f"  Source row multi-PID     : "
        f"{sc_stats['multi_pid_source_rows']:,}"
    )
    print(
        f"  Exploded output row      : "
        f"{sc_stats['exploded_rows']:,}"
    )
    print(
        f"  Single-PID SC preserved  : "
        f"{sc_stats['single_pid_sc_preserved']:,}"
    )
    print(
        f"  Multi-PID mapped TAP     : "
        f"{sc_stats['multi_pid_mapped_tap']:,}"
    )
    print(
        f"  Multi-PID mapped AFF     : "
        f"{sc_stats['multi_pid_mapped_aff_only']:,}"
    )
    print(
        f"  Unmapped multi-PID row   : "
        f"{sc_stats['unmapped_multi_pid']:,}"
    )
    print(
        f"  Source row blank PID     : "
        f"{sc_stats['blank_pid_source_rows']:,}"
    )
    print(
        f"  Invalid LIVE ID          : "
        f"{sc_stats['invalid_live_id']:,}"
    )
    print(
        f"  Unique missing pair      : "
        f"{len(missing_details):,}"
    )

    print(
        "\nBrand derivation "
        "(Product name -> Product ID fallback jika MISSING):"
    )
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
        print(
            f"  {brand:<15} "
            f"{brand_counter.get(brand, 0):>10,} row"
        )

    print("\nALL / RANGE:")

    if range_start or range_end:
        start_text = (
            range_start.isoformat()
            if range_start
            else "-inf"
        )

        end_text = (
            range_end.isoformat()
            if range_end
            else "+inf"
        )

        print(
            f"  RANGE                    : "
            f"{start_text} -> {end_text} "
            f"(SC LIVE start time)"
        )
    else:
        print(
            "  RANGE                    : "
            "same as ALL"
        )

    print(
        f"  {SHEET_ALL:<28}: "
        f"{len(all_rows):>10,} row"
    )

    print(
        f"  {SHEET_RANGE:<28}: "
        f"{len(range_rows):>10,} row"
    )

    print(
        f"  Invalid SC start RANGE   : "
        f"{invalid_range_dates:>10,}"
    )

    print(
        f"  Brand tabs               : "
        f"{'YES' if write_brand_tabs else 'NO'}"
    )

    missing_all_live_ids = {
        normalize_id(
            row.get("LIVE ID", "")
        )
        for row in sc_not_in_tap_all_rows
        if normalize_id(
            row.get("LIVE ID", "")
        )
    }

    missing_range_live_ids = {
        normalize_id(
            row.get("LIVE ID", "")
        )
        for row in sc_not_in_tap_range_rows
        if normalize_id(
            row.get("LIVE ID", "")
        )
    }

    print(
        "\nSC LIVE ID vs TAP LS "
        "(LIVE time info START range):"
    )

    print(
        f"  TAP raw rows             : "
        f"{tap_existence_stats['raw_rows']:>10,}"
    )

    print(
        f"  TAP rows start in period : "
        f"{tap_existence_stats['included_rows']:>10,}"
    )

    print(
        f"  TAP outside start period : "
        f"{tap_existence_stats['period_skipped']:>10,}"
    )

    print(
        f"  TAP invalid LIVE time    : "
        f"{tap_existence_stats['invalid_live_time']:>10,}"
    )

    print(
        f"  TAP blank LIVE ID        : "
        f"{tap_existence_stats['blank_or_invalid_live_id']:>10,}"
    )

    print(
        f"  Unique TAP LIVE ID RANGE : "
        f"{len(tap_live_ids):>10,}"
    )

    print(
        f"  Missing SC LIVE ID ALL   : "
        f"{len(missing_all_live_ids):>10,}"
    )

    print(
        f"  Missing SC rows ALL      : "
        f"{len(sc_not_in_tap_all_rows):>10,}"
    )

    print(
        f"  Missing SC LIVE ID RANGE : "
        f"{len(missing_range_live_ids):>10,}"
    )

    print(
        f"  Missing SC rows RANGE    : "
        f"{len(sc_not_in_tap_range_rows):>10,}"
    )

    print(
        "  Compare key              : "
        "LIVE ID vs Livestream room ID"
    )

    print(
        "  TAP existence basis      : "
        "START dari LIVE time info"
    )

    qc_all = summarize_mapping_qc(all_rows)
    qc_range = summarize_mapping_qc(range_rows)

    print("\nExact Product-ID Mapping QC (TAP LS + AFF):")
    print(
        f"  LIVE ID ALL              : "
        f"{qc_all['live_ids']:>10,}"
    )
    print(
        f"  Fully mapped ALL         : "
        f"{qc_all['fully_mapped']:>10,}"
    )
    print(
        f"  Partial mapping ALL      : "
        f"{qc_all['partial_mapping']:>10,}"
    )
    print(
        f"  No mapping ALL           : "
        f"{qc_all['no_mapping']:>10,}"
    )
    print(
        f"  PID mapped / total ALL   : "
        f"{qc_all['pid_mapped']:>10,} / "
        f"{qc_all['pid_total']:,} "
        f"({qc_all['pid_coverage']:.2%})"
    )
    print(
        f"  PID unmapped ALL         : "
        f"{qc_all['pid_unmapped']:>10,}"
    )
    print(
        f"  LIVE ID RANGE            : "
        f"{qc_range['live_ids']:>10,}"
    )
    print(
        f"  Fully mapped RANGE       : "
        f"{qc_range['fully_mapped']:>10,}"
    )
    print(
        f"  Partial mapping RANGE    : "
        f"{qc_range['partial_mapping']:>10,}"
    )
    print(
        f"  No mapping RANGE         : "
        f"{qc_range['no_mapping']:>10,}"
    )
    print(
        f"  PID mapped / total RANGE : "
        f"{qc_range['pid_mapped']:>10,} / "
        f"{qc_range['pid_total']:,} "
        f"({qc_range['pid_coverage']:.2%})"
    )
    print(
        f"  PID unmapped RANGE       : "
        f"{qc_range['pid_unmapped']:>10,}"
    )
    print(
        f"  Detail                   : "
        f"{MAPPING_QC_ALL_SHEET} / "
        f"{MAPPING_QC_RANGE_SHEET}"
    )

    print("\nCommerce metric reconciliation:")

    labels = {
        "gmv": "GMV",
        "orders": "Orders",
        "items": "Items sold",
    }

    for key in (
        "gmv",
        "orders",
        "items",
    ):
        original = original_totals[key]
        enriched = enriched_totals[key]
        delta = enriched - original

        error = pct_diff(
            float(enriched),
            float(original),
        )

        error_text = (
            "N/A"
            if error is None
            else f"{error:+.2f}%"
        )

        print(
            f"  {labels[key]:<10} "
            f"SC original {fmt_num(original):>16} | "
            f"mapped {fmt_num(enriched):>16} | "
            f"delta {fmt_num(delta):>16} | "
            f"{error_text}"
        )

    print("\nOutput:")
    print(
        f"  Spreadsheet ID           : "
        f"{OUTPUT_SPREADSHEET_ID}"
    )
    print(
        f"  Output columns           : "
        f"{len(OUTPUT_COLUMNS)}"
    )
    print(
        f"  Mode                     : "
        f"{'WRITE' if WRITE else 'DRY RUN'}"
    )
    print(
        f"  Replace                  : "
        f"{'YES' if REPLACE else 'NO'}"
    )

    print("\nCATATAN:")
    print(
        "  TAP metric mapping: START LIVE time info."
    )
    print(
        "  AFF metric mapping: Time Created."
    )
    print(
        "  SC RANGE: LIVE start time."
    )
    print(
        "  Brand: Product name PRIMARY; jika MISSING dan fallback ON, "
        "gunakan Product ID mapping."
    )
    print(
        "  SC LS NOT IN TAP: LIVE ID + Brand dibandingkan "
        "ke TAP Livestream room ID + Brand yang START live-nya "
        "masuk METRIC range."
    )
    print(
        "  Single-PID: metric commerce utama "
        "dipertahankan dari SC original."
    )
    print(
        "  Multi-PID: metric commerce utama "
        "direct mapping TAP/AFF (Model A)."
    )
    print(
        "  Metric level-LIVE seperti Likes/Comments/"
        "Shares/Impressions akan terulang pada setiap PID "
        "hasil explode; jangan di-SUM lintas Product ID."
    )


# =============================================================================
# CONFIG VALIDATION
# =============================================================================

def is_placeholder(value: object) -> bool:
    text = str(value or "").strip().upper()
    return (
        text.startswith("PASTE_")
        or text.startswith("ISI_")
        or not text
    )


def validate_config():
    missing = []

    if is_placeholder(SC_LS_SPREADSHEET_ID):
        missing.append(
            "SC_LS_SPREADSHEET_ID"
        )

    if is_placeholder(TAP_LS_SPREADSHEET_ID):
        missing.append(
            "TAP_LS_SPREADSHEET_ID"
        )

    if is_placeholder(OUTPUT_SPREADSHEET_ID):
        missing.append(
            "OUTPUT_SPREADSHEET_ID"
        )

    if USE_PRODUCT_ID_MAPPING_FALLBACK:
        if is_placeholder(MAPPING_SPREADSHEET_ID):
            missing.append("MAPPING_SPREADSHEET_ID")
        if MAPPING_SHEET_GID is None:
            missing.append("MAPPING_SHEET_GID")

    if not AFF_SOURCES:
        missing.append(
            "AFF_SOURCES"
        )
    else:
        for (
            i,
            source,
        ) in enumerate(
            AFF_SOURCES,
            start=1,
        ):
            if is_placeholder(
                source.get("spreadsheet_id")
            ):
                missing.append(
                    f"AFF_SOURCES[{i}]."
                    "spreadsheet_id"
                )

    if missing:
        raise SystemExit(
            "Isi CONFIG spreadsheet ID dulu:\n  - "
            + "\n  - ".join(missing)
        )

    if not CREDENTIALS_FILE.exists():
        candidates = "\n  - ".join(
            str(path)
            for path in CREDENTIALS_CANDIDATES
        )

        raise SystemExit(
            "Credentials tidak ditemukan. "
            "Script mencoba:\n  - "
            + candidates
        )

    metric_start = parse_iso_date(
        METRIC_START_DATE
    )

    metric_end = parse_iso_date(
        METRIC_END_DATE
    )

    range_start = parse_iso_date(
        RANGE_START_DATE
    )

    range_end = parse_iso_date(
        RANGE_END_DATE
    )

    if (
        metric_start
        and metric_end
        and metric_start > metric_end
    ):
        raise SystemExit(
            "METRIC_START_DATE tidak boleh "
            "lebih besar dari METRIC_END_DATE"
        )

    if (
        range_start
        and range_end
        and range_start > range_end
    ):
        raise SystemExit(
            "RANGE_START_DATE tidak boleh "
            "lebih besar dari RANGE_END_DATE"
        )

    return (
        metric_start,
        metric_end,
        range_start,
        range_end,
    )


# =============================================================================
# MAIN
# =============================================================================

def main():
    (
        metric_start,
        metric_end,
        range_start,
        range_end,
    ) = validate_config()

    service = get_sheets_service()

    print("=" * 78)
    print(
        "SC LIVE CLEAN + TAP LS/AFF "
        "ENRICHMENT - GOOGLE SHEETS"
    )
    print("=" * 78)

    print(
        f"Mode             : "
        f"{'WRITE' if WRITE else 'DRY RUN'}"
    )

    print(
        f"Replace          : "
        f"{'YES' if REPLACE else 'NO'}"
    )

    print(
        f"Brand tabs       : "
        f"{'YES' if WRITE_BRAND_TABS_EFFECTIVE else 'NO'}"
    )

    print(
        f"Version          : "
        f"{CLEANING_VERSION}"
    )

    if metric_start or metric_end:
        print(
            "Metric period    : "
            f"{metric_start.isoformat() if metric_start else '-inf'} "
            "-> "
            f"{metric_end.isoformat() if metric_end else '+inf'}"
        )

        print(
            "  TAP basis      : START LIVE time info"
        )

        print(
            "  AFF basis      : Time Created"
        )
    else:
        print(
            "Metric period    : "
            "ALL TAP/AFF source rows"
        )

    if range_start or range_end:
        print(
            "SC RANGE         : "
            f"{range_start.isoformat() if range_start else '-inf'} "
            "-> "
            f"{range_end.isoformat() if range_end else '+inf'} "
            "(LIVE start time)"
        )
    else:
        print(
            "SC RANGE         : same as ALL"
        )

    # -------------------------------------------------------------------------
    # Read SC LIVE
    # -------------------------------------------------------------------------
    (
        sc_header,
        sc_rows,
        sc_meta,
    ) = detect_and_read_source(
        service,
        spreadsheet_id=SC_LS_SPREADSHEET_ID,
        sheet_gid=SC_LS_SHEET_GID,
        sheet_name=SC_LS_SHEET_NAME,
        required_headers=SC_REQUIRED,
        label="SC LIVE",
    )

    print_source_meta(
        "SC LIVE",
        sc_meta,
    )

    # -------------------------------------------------------------------------
    # Read TAP LS
    # -------------------------------------------------------------------------
    (
        tap_header,
        tap_rows,
        tap_meta,
    ) = detect_and_read_source(
        service,
        spreadsheet_id=TAP_LS_SPREADSHEET_ID,
        sheet_gid=TAP_LS_SHEET_GID,
        sheet_name=TAP_LS_SHEET_NAME,
        required_headers=TAP_REQUIRED,
        label="TAP LS",
    )

    print_source_meta(
        "TAP LS",
        tap_meta,
    )

    (
        tap_live_ids,
        tap_existence_stats,
    ) = build_tap_ls_live_id_set(
        tap_header,
        tap_rows,
        metric_start,
        metric_end,
    )

    # -------------------------------------------------------------------------
    # Read AFF
    # -------------------------------------------------------------------------
    aff_datasets = []

    for (
        index,
        source,
    ) in enumerate(
        AFF_SOURCES,
        start=1,
    ):
        label = str(
            source.get("label")
            or f"AFF {index}"
        )

        (
            header,
            rows,
            meta,
        ) = detect_and_read_source(
            service,
            spreadsheet_id=source[
                "spreadsheet_id"
            ],
            sheet_gid=source.get(
                "sheet_gid"
            ),
            sheet_name=source.get(
                "sheet_name"
            ),
            required_headers=AFF_REQUIRED,
            label=label,
        )

        print_source_meta(
            label,
            meta,
        )

        aff_datasets.append(
            (
                label,
                header,
                rows,
            )
        )

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
    print(
        "\nMembangun TAP LS mapping..."
    )

    (
        tap_map,
        tap_stats,
    ) = build_tap_ls_mapping(
        tap_header,
        tap_rows,
        metric_start,
        metric_end,
    )

    print(
        "Membangun Affiliate "
        "Livestream mapping..."
    )

    (
        aff_map,
        aff_stats,
    ) = build_aff_ls_mapping(
        aff_datasets,
        metric_start,
        metric_end,
    )

    print(
        "Menggabungkan TAP LS + AFF "
        "menjadi master mapping..."
    )

    (
        master_map,
        master_counts,
    ) = build_master_mapping(
        tap_map,
        aff_map,
        product_brand_mapping,
    )

    pid_product_name_lookup = (
        build_pid_product_name_lookup(
            tap_map,
            aff_map,
        )
    )

    tap_live_brand_pairs = build_tap_brand_pair_set_from_mapping(
        tap_map,
        pid_product_name_lookup,
        product_brand_mapping,
    )

    # -------------------------------------------------------------------------
    # SC explode + enrichment
    # -------------------------------------------------------------------------
    print(
        "Explode + enrich SC LIVE..."
    )

    (
        all_rows,
        sc_stats,
        original_totals,
        enriched_totals,
        brand_counter,
        missing_details,
    ) = enrich_sc_live(
        sc_header,
        sc_rows,
        master_map,
        aff_map,
        pid_product_name_lookup,
        product_brand_mapping,
    )

    (
        range_rows,
        invalid_range_dates,
    ) = filter_output_range(
        all_rows,
        range_start,
        range_end,
    )

    # -------------------------------------------------------------------------
    # SC LS NOT IN TAP
    # -------------------------------------------------------------------------
    sc_not_in_tap_all_rows = (
        filter_sc_ls_not_in_tap(
            all_rows,
            tap_live_brand_pairs,
        )
    )

    sc_not_in_tap_range_rows = filter_sc_ls_not_in_tap(
        range_rows,
        tap_live_brand_pairs,
    )

    # BRAND UNIQUE: first PID per LIVE ID + Brand.
    # NOT IN TAP memakai grain LIVE ID + Brand.
    brand_unique_all_rows = build_brand_unique_rows(all_rows)
    brand_unique_range_rows = build_brand_unique_rows(range_rows)
    brand_unique_not_in_tap_all_rows = build_brand_unique_rows(
        sc_not_in_tap_all_rows
    )
    brand_unique_not_in_tap_range_rows = build_brand_unique_rows(
        sc_not_in_tap_range_rows
    )

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

    # Summary brand mengikuti grain final LIVE ID + Brand.
    output_tables[SUMMARY_SHEET] = build_brand_summary(
        brand_unique_all_rows,
        brand_unique_range_rows,
    )

    output_info = validate_output_capacity(
        service,
        output_tables,
    )

    print("\nBRAND UNIQUE OUTPUT (FIRST PID per LIVE ID + Brand):")
    print(f"  {BRAND_UNIQUE_ALL_SHEET:<46}: {len(brand_unique_all_rows):>10,} row")
    print(f"  {BRAND_UNIQUE_RANGE_SHEET:<46}: {len(brand_unique_range_rows):>10,} row")
    print(f"  {BRAND_UNIQUE_NOT_IN_TAP_ALL_SHEET:<46}: {len(brand_unique_not_in_tap_all_rows):>10,} row")
    print(f"  {BRAND_UNIQUE_NOT_IN_TAP_RANGE_SHEET:<46}: {len(brand_unique_not_in_tap_range_rows):>10,} row")
    print("  NOT IN TAP key: LIVE ID + Brand.")
    print("  Brand GMV/Orders/Items: SUM unique Product ID per LIVE ID + Brand.")

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
        tap_live_ids,
        tap_existence_stats,
        invalid_range_dates,
        range_start,
        range_end,
        WRITE_BRAND_TABS_EFFECTIVE,
        missing_details,
    )

    if not WRITE:
        print(
            "\nDRY RUN selesai. "
            "Tidak ada data ditulis."
        )

        print(
            "Jika QC sudah sesuai:"
        )

        print(
            "  python "
            "sc_live_enriched_notin_tap.py "
            "--write"
        )

        print(
            "Jika tab output lama memang "
            "mau diganti:"
        )

        print(
            "  python "
            "sc_live_enriched_notin_tap.py "
            "--write --replace"
        )

        return

    print(
        "\nMenulis output ke Google Sheets..."
    )

    write_output(
        service,
        output_tables,
        output_info,
    )

    print(
        "\nSELESAI."
    )


if __name__ == "__main__":
    raise SystemExit("Jalankan melalui Streamlit dengan perintah streamlit run app.py.")
