"""Shared utilities for TAP PC Livestream and Short Video processing."""

from __future__ import annotations

import hashlib
import json
import math
import re
import time
from collections import Counter, defaultdict
from datetime import date, datetime, timedelta

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


# =============================================================================
# Brand classification
# =============================================================================

PRODUCT_NAME_HEADER = "Product name"
SOURCE_PRODUCT_ID_HEADER = "Product ID"
OUTPUT_BRAND_HEADER = "Brand"
MISSING_BRAND = "MISSING"

PC_BRANDS = {
    "CLOSE UP",
    "DOVE DEO",
    "DOVE SCL",
    "LIFEBUOY",
    "LUX",
    "PEPSODENT",
    "REXONA",
}

BASE_BRAND_PATTERNS = {
    "CLOSE UP": re.compile(
        r"(?<![A-Z0-9])CLOSE\s*-?\s*UP(?![A-Z0-9])",
        re.I,
    ),
    "DOVE": re.compile(
        r"(?<![A-Z0-9])DOVE(?![A-Z0-9])",
        re.I,
    ),
    "LIFEBUOY": re.compile(
        r"(?<![A-Z0-9])LIFEBUOY(?![A-Z0-9])",
        re.I,
    ),
    "LUX": re.compile(
        r"(?<![A-Z0-9])LUX(?![A-Z0-9])",
        re.I,
    ),
    "PEPSODENT": re.compile(
        r"(?<![A-Z0-9])PEPSODENT(?![A-Z0-9])",
        re.I,
    ),
    "REXONA": re.compile(
        r"(?<![A-Z0-9])REXONA(?![A-Z0-9])",
        re.I,
    ),
}

# Dove Deo and Dove Deodorant are classified as DOVE DEO.
DOVE_DEO_RE = re.compile(
    r"(?<![A-Z0-9])(?:DEO|DEODORANT)(?![A-Z0-9])",
    re.I,
)

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


# =============================================================================
# Optional Product ID fallback mapping
# =============================================================================

MAPPING_SPREADSHEET_ID = "1O4dQVFYawbirEN5wpV0ygjNZtNFZPZH0tLX_Ov34fNA"
MAPPING_SHEET_GID = 437523491
MAPPING_PRODUCT_ID_HEADER = "Product ID TTS"
MAPPING_BRAND_HEADER = "BRAND FIX"


# =============================================================================
# Google Sheets policy
# =============================================================================

MAX_CELLS = 20_000_000
WRITE_CHUNK_ROWS = 1500
MAX_RETRIES = 5
RETRYABLE_HTTP_STATUS = {429, 500, 502, 503, 504}


# =============================================================================
# Type constants
# =============================================================================

TYPE_TEXT = "text"
TYPE_DATE = "date"
TYPE_DATETIME = "datetime"
TYPE_INTEGER = "integer"
TYPE_NUMBER = "number"
TYPE_CURRENCY = "currency"
TYPE_PERCENT = "percent"

NULL_NUMERIC_TOKENS = {
    "",
    "-",
    "--",
    "—",
    "N/A",
    "NA",
    "NULL",
    "NONE",
    "#N/A",
}


# =============================================================================
# Google Sheets client
# =============================================================================

def get_sheets_service(credentials_file, write=False):
    scope = (
        "https://www.googleapis.com/auth/spreadsheets"
        if write
        else "https://www.googleapis.com/auth/spreadsheets.readonly"
    )

    credentials = service_account.Credentials.from_service_account_file(
        str(credentials_file),
        scopes=[scope],
    )

    return build(
        "sheets",
        "v4",
        credentials=credentials,
        cache_discovery=False,
    )


def execute_with_retry(request, max_retries=MAX_RETRIES):
    for attempt in range(max_retries):
        try:
            return request.execute()

        except HttpError as exc:
            status = getattr(exc.resp, "status", None)

            if (
                status not in RETRYABLE_HTTP_STATUS
                or attempt == max_retries - 1
            ):
                raise

            wait_seconds = 2 ** attempt

            print(
                f"Google API error {status}; "
                f"retry {attempt + 1}/{max_retries} "
                f"in {wait_seconds}s..."
            )

            time.sleep(wait_seconds)


# =============================================================================
# General helpers
# =============================================================================

def cell(row, index):
    return row[index] if index < len(row) else ""


def normalize_header(value):
    return re.sub(
        r"[^a-z0-9]+",
        "",
        str(value or "").strip().lower(),
    )


def find_header_index(header, name):
    target = normalize_header(name)

    for index, value in enumerate(header):
        if normalize_header(value) == target:
            return index

    raise ValueError(
        f'Kolom "{name}" tidak ditemukan. Header: {header}'
    )


def column_letter(index_zero_based):
    number = index_zero_based + 1
    result = ""

    while number:
        number, remainder = divmod(number - 1, 26)
        result = chr(65 + remainder) + result

    return result


# =============================================================================
# ID handling
# =============================================================================

def normalize_id(value):
    if value is None:
        return ""

    if isinstance(value, bool):
        return str(value)

    if isinstance(value, int):
        return str(value)

    if isinstance(value, float):
        if not math.isfinite(value):
            return ""

        if abs(value) >= 1e15:
            raise ValueError(
                f"ID terbaca sebagai FLOAT besar ({value}). "
                "Format source ID sebagai Plain text."
            )

        if value.is_integer():
            return format(value, ".0f")

        return (
            format(value, "f")
            .rstrip("0")
            .rstrip(".")
        )

    result = str(value).strip()

    if not result:
        return ""

    if re.fullmatch(
        r"[+-]?\d+(?:\.\d+)?[eE][+-]?\d+",
        result,
    ):
        raise ValueError(
            f"ID tampil dalam scientific notation ({result}). "
            "Format source ID sebagai Plain text."
        )

    if re.fullmatch(r"\d+\.0", result):
        result = result[:-2]

    return result


# =============================================================================
# Numeric conversion
# =============================================================================

def _is_blank_numeric(value):
    if value is None:
        return True

    return (
        str(value).strip().upper()
        in NULL_NUMERIC_TOKENS
    )


def _split_compact_suffix(value):
    text = value.strip()

    match = re.fullmatch(
        r"(.+?)([KMB])",
        text,
        flags=re.I,
    )

    if not match:
        return text, 1.0

    multiplier = {
        "K": 1_000.0,
        "M": 1_000_000.0,
        "B": 1_000_000_000.0,
    }[match.group(2).upper()]

    return match.group(1), multiplier


def _normalize_decimal_string(value, currency_mode=False):
    text = str(value).strip()

    text = re.sub(
        r"^(?:RP|IDR)\s*",
        "",
        text,
        flags=re.I,
    )

    text = text.replace(" ", "")
    text, multiplier = _split_compact_suffix(text)

    negative = (
        text.startswith("(")
        and text.endswith(")")
    )

    if negative:
        text = text[1:-1]

    text = text.replace("%", "")

    if re.fullmatch(
        r"[+-]?\d+(?:\.\d+)?[eE][+-]?\d+",
        text,
    ):
        number = float(text) * multiplier
        return -number if negative else number

    if not re.fullmatch(
        r"[+-]?[\d.,]+",
        text,
    ):
        raise ValueError(
            f"nilai numerik tidak valid: {value!r}"
        )

    sign = ""

    if text[:1] in {"+", "-"}:
        sign = text[0]
        text = text[1:]

    dot_count = text.count(".")
    comma_count = text.count(",")

    if dot_count and comma_count:
        decimal_separator = (
            "."
            if text.rfind(".") > text.rfind(",")
            else ","
        )

        thousands_separator = (
            ","
            if decimal_separator == "."
            else "."
        )

        text = text.replace(
            thousands_separator,
            "",
        )

        text = text.replace(
            decimal_separator,
            ".",
        )

    elif dot_count or comma_count:
        separator = "." if dot_count else ","
        count = text.count(separator)
        groups = text.split(separator)

        if count > 1:
            if all(
                len(group) == 3
                for group in groups[1:]
            ):
                text = "".join(groups)
            else:
                text = (
                    "".join(groups[:-1])
                    + "."
                    + groups[-1]
                )

        else:
            left, right = groups

            if (
                currency_mode
                and len(right) == 3
            ):
                text = left + right
            else:
                text = left + "." + right

    number = float(sign + text) * multiplier

    if negative:
        number = -number

    return number


def parse_integer_value(value):
    if _is_blank_numeric(value):
        return ""

    if isinstance(value, bool):
        raise ValueError(
            f"boolean bukan integer metric: {value!r}"
        )

    if isinstance(value, int):
        return value

    if isinstance(value, float):
        if not math.isfinite(value):
            raise ValueError(
                f"nilai integer tidak finite: {value!r}"
            )

        if not value.is_integer():
            raise ValueError(
                f"nilai integer memiliki desimal: {value!r}"
            )

        return int(value)

    text = str(value).strip()

    if re.search(
        r"[KMB]$",
        text,
        flags=re.I,
    ):
        number = _normalize_decimal_string(
            text,
            currency_mode=False,
        )

    else:
        cleaned = re.sub(
            r"^(?:RP|IDR)\s*",
            "",
            text,
            flags=re.I,
        ).replace(" ", "")

        if re.fullmatch(
            r"[+-]?\d+(?:\.0+)?",
            cleaned,
        ):
            number = float(cleaned)

        elif re.fullmatch(
            r"[+-]?\d{1,3}(?:[.,]\d{3})+",
            cleaned,
        ):
            number = float(
                re.sub(r"[.,]", "", cleaned)
            )

        elif re.fullmatch(
            r"[+-]?\d+(?:\.\d+)?[eE][+-]?\d+",
            cleaned,
        ):
            number = float(cleaned)

        else:
            number = _normalize_decimal_string(
                cleaned,
                currency_mode=False,
            )

    if not float(number).is_integer():
        raise ValueError(
            f"nilai integer memiliki desimal: {value!r}"
        )

    return int(number)


def parse_number_value(value, currency_mode=False):
    if _is_blank_numeric(value):
        return ""

    if isinstance(value, bool):
        raise ValueError(
            f"boolean bukan numeric metric: {value!r}"
        )

    if isinstance(value, (int, float)):
        if (
            isinstance(value, float)
            and not math.isfinite(value)
        ):
            raise ValueError(
                f"nilai numerik tidak finite: {value!r}"
            )

        return value

    number = _normalize_decimal_string(
        value,
        currency_mode=currency_mode,
    )

    if float(number).is_integer():
        return int(number)

    return float(number)


def parse_percent_value(value):
    if _is_blank_numeric(value):
        return ""

    if isinstance(value, bool):
        raise ValueError(
            f"boolean bukan percentage metric: {value!r}"
        )

    if isinstance(value, (int, float)):
        return value

    text = str(value).strip()
    has_percent = text.endswith("%")

    if has_percent:
        text = text[:-1]

    number = parse_number_value(
        text,
        currency_mode=False,
    )

    if number == "":
        return ""

    return (
        float(number) / 100.0
        if has_percent
        else number
    )


def to_number(value):
    if _is_blank_numeric(value):
        return 0

    if isinstance(value, bool):
        return 0

    if isinstance(value, (int, float)):
        return value

    try:
        return parse_number_value(
            value,
            currency_mode=True,
        )

    except ValueError:
        return 0


def numbers_equal(left, right):
    return math.isclose(
        float(left),
        float(right),
        rel_tol=1e-9,
        abs_tol=1e-6,
    )


# =============================================================================
# Date and datetime conversion
# =============================================================================

def normalize_date_text(value):
    if (
        value is None
        or str(value).strip() in {"", "--", "-"}
    ):
        return ""

    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")

    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")

    if (
        isinstance(value, (int, float))
        and not isinstance(value, bool)
    ):
        converted = (
            datetime(1899, 12, 30)
            + timedelta(days=float(value))
        )
        return converted.strftime("%Y-%m-%d")

    text = str(value).strip()

    formats = (
        "%Y-%m-%d",
        "%Y/%m/%d",
        "%d/%m/%Y",
        "%m/%d/%Y",
        "%d-%m-%Y",
    )

    for fmt in formats:
        try:
            return datetime.strptime(
                text[:10],
                fmt,
            ).strftime("%Y-%m-%d")

        except ValueError:
            pass

    raise ValueError(
        f"tanggal tidak valid: {value!r}"
    )


def normalize_datetime_text(value):
    if (
        value is None
        or str(value).strip() in {"", "--", "-"}
    ):
        return ""

    if isinstance(value, datetime):
        return value.strftime(
            "%Y-%m-%d %H:%M:%S"
        )

    if isinstance(value, date):
        return datetime.combine(
            value,
            datetime.min.time(),
        ).strftime(
            "%Y-%m-%d %H:%M:%S"
        )

    text = str(value).strip()

    formats = (
    "%Y-%m-%d %H:%M:%S",
    "%Y-%m-%d %H:%M",
    "%Y-%m-%d",

    # TAP export: MM/DD/YYYY
    "%m/%d/%Y %H:%M:%S",
    "%m/%d/%Y %H:%M",
    "%m/%d/%Y",

    # Optional fallback
    "%d/%m/%Y %H:%M:%S",
    "%d/%m/%Y %H:%M",
    "%d/%m/%Y",
    )

    for fmt in formats:
        try:
            converted = datetime.strptime(
                text,
                fmt,
            )

            return converted.strftime(
                "%Y-%m-%d %H:%M:%S"
            )

        except ValueError:
            pass

    raise ValueError(
        f"datetime tidak valid: {value!r}"
    )


def infer_column_type(header_name):
    """Infer a conservative type from complete header tokens."""
    raw = str(header_name or "").strip().lower()
    tokens = re.findall(r"[a-z0-9]+", raw)
    token_set = set(tokens)
    compact = normalize_header(header_name)

    if token_set.intersection({"id", "code"}):
        return TYPE_TEXT

    if "duration" in token_set or "duration" in compact:
        return TYPE_TEXT

    if "date" in token_set:
        return TYPE_DATE

    if "time" in token_set:
        return TYPE_DATETIME

    if token_set.intersection(
        {
            "likes",
            "views",
            "orders",
            "items",
            "follower",
            "followers",
            "count",
            "clicks",
            "impressions",
            "comments",
            "shares",
        }
    ):
        return TYPE_INTEGER

    if token_set.intersection(
        {
            "gmv",
            "commission",
            "rpm",
            "aov",
            "gpm",
            "revenue",
            "amount",
            "price",
        }
    ):
        return TYPE_CURRENCY

    if token_set.intersection(
        {
            "ctr",
            "rate",
            "engagement",
            "percentage",
            "percent",
        }
    ):
        return TYPE_PERCENT

    return TYPE_TEXT


def coerce_rows_by_schema(
    header,
    rows,
    column_types=None,
    strict=True,
    skip_invalid_rows=False,
):
    """
    Convert source values to stable Python types.

    When skip_invalid_rows is enabled, a row with any conversion error is
    excluded from downstream processing and reported through the errors result.
    """
    configured = {
        normalize_header(name): data_type
        for name, data_type
        in (column_types or {}).items()
    }

    resolved = [
        configured.get(
            normalize_header(name),
            infer_column_type(name),
        )
        for name in header
    ]

    converted_rows = []
    errors = []

    for row_number, row in enumerate(
        rows,
        start=2,
    ):
        output_row = [
            cell(row, index)
            for index in range(len(header))
        ]

        row_errors = []

        for index, (
            name,
            data_type,
        ) in enumerate(
            zip(header, resolved)
        ):
            value = output_row[index]

            try:
                if data_type == TYPE_TEXT:
                    output_row[index] = (
                        ""
                        if value is None
                        else str(value).strip()
                    )

                elif data_type == TYPE_DATE:
                    output_row[index] = (
                        normalize_date_text(value)
                    )

                elif data_type == TYPE_DATETIME:
                    output_row[index] = (
                        normalize_datetime_text(value)
                    )

                elif data_type == TYPE_INTEGER:
                    output_row[index] = (
                        parse_integer_value(value)
                    )

                elif data_type == TYPE_NUMBER:
                    output_row[index] = (
                        parse_number_value(
                            value,
                            currency_mode=False,
                        )
                    )

                elif data_type == TYPE_CURRENCY:
                    output_row[index] = (
                        parse_number_value(
                            value,
                            currency_mode=True,
                        )
                    )

                elif data_type == TYPE_PERCENT:
                    output_row[index] = (
                        parse_percent_value(value)
                    )

                else:
                    raise ValueError(
                        f"tipe kolom tidak dikenal: "
                        f"{data_type!r}"
                    )

            except ValueError as exc:
                row_errors.append(
                    (
                        row_number,
                        name,
                        value,
                        str(exc),
                    )
                )

        if row_errors:
            errors.extend(row_errors)

            if skip_invalid_rows:
                continue

        converted_rows.append(
            output_row
        )

    if errors and strict:
        detail = "\n".join(
            f"  row {row_number}, "
            f"{name}: {value!r} "
            f"({reason})"
            for (
                row_number,
                name,
                value,
                reason,
            ) in errors[:20]
        )

        extra = (
            f"\n  ... {len(errors) - 20:,} error lainnya."
            if len(errors) > 20
            else ""
        )

        raise ValueError(
            "Konversi tipe data gagal. "
            "Periksa source atau COLUMN_TYPES:\n"
            + detail
            + extra
        )

    type_summary = Counter(resolved)

    return (
        converted_rows,
        dict(zip(header, resolved)),
        type_summary,
        errors,
    )


# =============================================================================
# Reporting dates
# =============================================================================

def parse_config_date(value, label):
    text = str(value or "").strip()

    if not text:
        raise ValueError(
            f"{label} wajib diisi YYYY-MM-DD."
        )

    try:
        return datetime.strptime(
            text,
            "%Y-%m-%d",
        ).date()

    except ValueError as exc:
        raise ValueError(
            f"{label} harus format YYYY-MM-DD. "
            f"Dapat: {value!r}"
        ) from exc


def parse_date_value(value):
    if (
        value is None
        or value == ""
    ):
        return None

    if isinstance(value, datetime):
        return value.date()

    if isinstance(value, date):
        return value

    text = str(value).strip()

    if not text:
        return None

    if re.match(
        r"^\d{4}-\d{2}-\d{2}",
        text,
    ):
        try:
            return datetime.strptime(
                text[:10],
                "%Y-%m-%d",
            ).date()

        except ValueError:
            pass

    for fmt in (
        "%d/%m/%Y",
        "%m/%d/%Y",
        "%d-%m-%Y",
    ):
        try:
            return datetime.strptime(
                text[:10],
                fmt,
            ).date()

        except ValueError:
            pass

    return None


def filter_groups_by_date_range(
    header,
    groups,
    date_column,
    start_date,
    end_date,
):
    start = parse_config_date(
        start_date,
        "RANGE_START_DATE",
    )

    end = parse_config_date(
        end_date,
        "RANGE_END_DATE",
    )

    if start > end:
        raise ValueError(
            "RANGE_START_DATE tidak boleh "
            "lebih besar dari RANGE_END_DATE."
        )

    date_index = find_header_index(
        header,
        date_column,
    )

    filtered = {}
    invalid_count = 0

    for key, row in groups.items():
        parsed_date = parse_date_value(
            cell(
                row,
                date_index,
            )
        )

        if parsed_date is None:
            invalid_count += 1
            continue

        if start <= parsed_date <= end:
            filtered[key] = row

    return (
        filtered,
        invalid_count,
    )


def filter_groups_by_date_ranges(
    header,
    groups,
    date_column,
    ranges,
):
    """Filter aggregated groups by the union of several inclusive ranges."""
    date_index = find_header_index(
        header,
        date_column,
    )

    normalized_ranges = []

    for start_date, end_date in ranges:
        if isinstance(start_date, str):
            start_date = parse_config_date(start_date, "RANGE_START_DATE")
        if isinstance(end_date, str):
            end_date = parse_config_date(end_date, "RANGE_END_DATE")

        if start_date > end_date:
            raise ValueError(
                "Tanggal awal cumulative RANGE tidak boleh lebih besar dari tanggal akhir."
            )

        normalized_ranges.append((start_date, end_date))

    normalized_ranges = merge_date_ranges(normalized_ranges)

    filtered = {}
    invalid_count = 0

    for key, row in groups.items():
        parsed_date = parse_date_value(
            cell(row, date_index)
        )

        if parsed_date is None:
            invalid_count += 1
            continue

        if any(
            start_date <= parsed_date <= end_date
            for start_date, end_date in normalized_ranges
        ):
            filtered[key] = row

    return filtered, invalid_count


# =============================================================================
# Source reader
# =============================================================================

def read_source(
    service,
    spreadsheet_id,
    source_sheet="Custom report",
    source_range="A1:AZ",
):
    response = execute_with_retry(
        service
        .spreadsheets()
        .values()
        .get(
            spreadsheetId=spreadsheet_id,
            range=f"'{source_sheet}'!{source_range}",
            valueRenderOption="UNFORMATTED_VALUE",
            dateTimeRenderOption="FORMATTED_STRING",
        )
    )

    values = response.get(
        "values",
        [],
    )

    if not values:
        raise ValueError(
            f'Tab source "{source_sheet}" '
            "kosong / tidak ditemukan."
        )

    return (
        values[0],
        values[1:],
    )


# =============================================================================
# Brand extraction
# =============================================================================

def find_brand_mentions(product_name):
    text = str(product_name or "")
    mentions = []

    for (
        base_brand,
        pattern,
    ) in BASE_BRAND_PATTERNS.items():
        match = pattern.search(text)

        if match:
            mentions.append(
                (
                    match.start(),
                    match.end(),
                    base_brand,
                    match.group(0),
                )
            )

    mentions.sort(
        key=lambda item: (
            item[0],
            item[1],
            item[2],
        )
    )

    return mentions


def extract_parent_brand_from_product_name(
    product_name,
):
    """
    Use the left-most recognized PC brand as the parent brand.

    Dove is classified as DOVE DEO when the parent-product segment contains
    either "Deo" or "Deodorant". Other Dove products are classified as
    DOVE SCL.
    """
    text = str(product_name or "")
    mentions = find_brand_mentions(text)

    if not mentions:
        return {
            "brand": "",
            "base_brand": "",
            "is_bundle": False,
            "match_count": 0,
            "matched_text": "",
            "match_position": None,
        }

    (
        parent_start,
        _parent_end,
        parent_base,
        matched_text,
    ) = mentions[0]

    brand = parent_base

    if parent_base == "DOVE":
        next_brand_start = (
            mentions[1][0]
            if len(mentions) > 1
            else len(text)
        )

        parent_segment = text[
            :next_brand_start
        ]

        brand = (
            "DOVE DEO"
            if DOVE_DEO_RE.search(
                parent_segment
            )
            else "DOVE SCL"
        )

    return {
        "brand": brand,
        "base_brand": parent_base,
        "is_bundle": len(mentions) > 1,
        "match_count": len(mentions),
        "matched_text": matched_text,
        "match_position": parent_start,
    }


def canonicalize_mapping_brand(value):
    raw = str(value or "").strip().upper()

    if not raw:
        return ""

    canonical = MAPPING_BRAND_ALIASES.get(
        raw,
        raw,
    )

    return (
        canonical
        if canonical in PC_BRANDS
        else ""
    )


def add_brand_from_product_name(
    header,
    rows,
    fallback_mapping=None,
    product_name_header=PRODUCT_NAME_HEADER,
    product_id_header=SOURCE_PRODUCT_ID_HEADER,
    brand_header=OUTPUT_BRAND_HEADER,
    missing_brand=MISSING_BRAND,
):
    fallback_mapping = fallback_mapping or {}

    name_index = find_header_index(
        header,
        product_name_header,
    )

    product_id_index = (
        find_header_index(
            header,
            product_id_header,
        )
        if fallback_mapping
        else None
    )

    try:
        brand_index = find_header_index(
            header,
            brand_header,
        )

        output_header = list(header)

    except ValueError:
        output_header = (
            list(header)
            + [brand_header]
        )

        brand_index = (
            len(output_header)
            - 1
        )

    output_rows = []
    status = Counter()
    missing_names = Counter()
    bundle_parent_counts = Counter()

    for row in rows:
        output_row = [
            cell(
                row,
                index,
            )
            for index in range(
                len(output_header)
            )
        ]

        product_name = str(
            cell(
                row,
                name_index,
            )
            or ""
        ).strip()

        result = (
            extract_parent_brand_from_product_name(
                product_name
            )
        )

        brand = result["brand"]

        if brand:
            if result["is_bundle"]:
                status[
                    "PRODUCT_NAME_BUNDLE_PARENT"
                ] += 1

                bundle_parent_counts[
                    brand
                ] += 1

            else:
                status[
                    "PRODUCT_NAME"
                ] += 1

            if brand == "DOVE DEO":
                status[
                    "DOVE_DEO"
                ] += 1

            elif brand == "DOVE SCL":
                status[
                    "DOVE_SCL"
                ] += 1

        elif fallback_mapping:
            product_id = normalize_id(
                cell(
                    row,
                    product_id_index,
                )
            )

            mapped_brand = (
                fallback_mapping.get(
                    product_id,
                    "",
                )
            )

            if mapped_brand:
                brand = mapped_brand

                status[
                    "PRODUCT_ID_MAPPING_FALLBACK"
                ] += 1

        if not brand:
            brand = missing_brand
            status[
                "MISSING"
            ] += 1

            missing_names[
                (
                    product_name
                    if product_name
                    else "<EMPTY_PRODUCT_NAME>"
                )
            ] += 1

        output_row[
            brand_index
        ] = brand

        output_rows.append(
            output_row
        )

    return (
        output_header,
        output_rows,
        status,
        missing_names,
        bundle_parent_counts,
    )


# =============================================================================
# Product ID fallback mapping
# =============================================================================

def get_sheet_title_by_gid(
    service,
    spreadsheet_id,
    gid,
):
    metadata = execute_with_retry(
        service
        .spreadsheets()
        .get(
            spreadsheetId=spreadsheet_id,
            fields=(
                "sheets("
                "properties("
                "sheetId,"
                "title"
                ")"
                ")"
            ),
        )
    )

    for sheet in metadata.get(
        "sheets",
        [],
    ):
        properties = sheet[
            "properties"
        ]

        if int(
            properties["sheetId"]
        ) == int(gid):
            return properties[
                "title"
            ]

    raise ValueError(
        f"GID {gid} tidak ditemukan "
        "di mapping spreadsheet."
    )


def detect_mapping_header(
    service,
    spreadsheet_id=MAPPING_SPREADSHEET_ID,
    sheet_gid=MAPPING_SHEET_GID,
    pid_header=MAPPING_PRODUCT_ID_HEADER,
    brand_header=MAPPING_BRAND_HEADER,
):
    title = get_sheet_title_by_gid(
        service,
        spreadsheet_id,
        sheet_gid,
    )

    response = execute_with_retry(
        service
        .spreadsheets()
        .values()
        .get(
            spreadsheetId=spreadsheet_id,
            range=f"'{title}'!A1:AZ20",
            valueRenderOption="FORMATTED_VALUE",
        )
    )

    product_id_target = normalize_header(
        pid_header
    )

    brand_target = normalize_header(
        brand_header
    )

    for row_number, row in enumerate(
        response.get(
            "values",
            [],
        ),
        start=1,
    ):
        normalized = [
            normalize_header(value)
            for value in row
        ]

        product_id_index = next(
            (
                index
                for index, value
                in enumerate(normalized)
                if value == product_id_target
            ),
            None,
        )

        brand_index = next(
            (
                index
                for index, value
                in enumerate(normalized)
                if value == brand_target
            ),
            None,
        )

        if (
            product_id_index is not None
            and brand_index is not None
        ):
            return (
                title,
                row_number,
                product_id_index,
                brand_index,
            )

    raise ValueError(
        "Header mapping tidak ditemukan. "
        f'Harus ada "{pid_header}" '
        f'dan "{brand_header}".'
    )


def load_product_brand_mapping(
    service,
    spreadsheet_id=MAPPING_SPREADSHEET_ID,
    sheet_gid=MAPPING_SHEET_GID,
    pid_header=MAPPING_PRODUCT_ID_HEADER,
    brand_header=MAPPING_BRAND_HEADER,
):
    (
        title,
        header_row,
        product_id_index,
        brand_index,
    ) = detect_mapping_header(
        service,
        spreadsheet_id,
        sheet_gid,
        pid_header,
        brand_header,
    )

    last_column = column_letter(
        max(
            product_id_index,
            brand_index,
        )
    )

    response = execute_with_retry(
        service
        .spreadsheets()
        .values()
        .get(
            spreadsheetId=spreadsheet_id,
            range=(
                f"'{title}'!"
                f"A{header_row + 1}:"
                f"{last_column}"
            ),
            valueRenderOption="FORMATTED_VALUE",
        )
    )

    mapping = {}
    conflicts = defaultdict(set)
    skipped_non_pc = 0

    for row in response.get(
        "values",
        [],
    ):
        product_id = normalize_id(
            cell(
                row,
                product_id_index,
            )
        )

        brand = canonicalize_mapping_brand(
            cell(
                row,
                brand_index,
            )
        )

        if not product_id:
            continue

        if not brand:
            skipped_non_pc += 1
            continue

        if (
            product_id in mapping
            and mapping[product_id] != brand
        ):
            conflicts[
                product_id
            ].update(
                {
                    mapping[product_id],
                    brand,
                }
            )

            continue

        mapping.setdefault(
            product_id,
            brand,
        )

    if conflicts:
        examples = [
            f"{product_id}: {sorted(brands)}"
            for (
                product_id,
                brands,
            ) in list(
                conflicts.items()
            )[:20]
        ]

        raise ValueError(
            "MASTER MAPPING CONFLICT:\n  "
            + "\n  ".join(examples)
        )

    return (
        mapping,
        {
            "sheet_title": title,
            "mapping_count": len(mapping),
            "skipped_non_pc": skipped_non_pc,
        },
    )


# =============================================================================
# Content-level aggregation
# =============================================================================

def aggregate_brand_content(
    header,
    rows,
    content_id_header,
    date_header,
    sum_columns,
    brand_header=OUTPUT_BRAND_HEADER,
    row_validator=None,
):
    """
    Aggregate source rows to the (Brand, Content ID) grain.

    Additive metrics are summed. Date keeps the earliest value. Other fields
    keep the first accepted row.
    """
    content_id_index = find_header_index(
        header,
        content_id_header,
    )

    brand_index = find_header_index(
        header,
        brand_header,
    )

    date_index = find_header_index(
        header,
        date_header,
    )

    sum_indexes = [
        find_header_index(
            header,
            column,
        )
        for column in sum_columns
    ]

    groups = {}
    accepted_rows = []
    skipped_blank = 0
    invalid_rows = []

    for row_number, row in enumerate(
        rows,
        start=2,
    ):
        content_id = normalize_id(
            cell(
                row,
                content_id_index,
            )
        )

        brand = str(
            cell(
                row,
                brand_index,
            )
            or ""
        ).strip().upper()

        if (
            not content_id
            or not brand
        ):
            skipped_blank += 1
            continue

        if row_validator is not None:
            (
                is_valid,
                reason,
            ) = row_validator(
                row,
                row_number,
                header,
            )

            if not is_valid:
                invalid_rows.append(
                    (
                        row_number,
                        reason,
                    )
                )
                continue

        accepted_rows.append(
            row
        )

        key = (
            brand,
            content_id,
        )

        if key not in groups:
            base = [
                cell(
                    row,
                    index,
                )
                for index in range(
                    len(header)
                )
            ]

            base[
                brand_index
            ] = brand

            base[
                content_id_index
            ] = content_id

            for index in sum_indexes:
                base[index] = 0

            groups[
                key
            ] = base

        base = groups[key]

        for index in sum_indexes:
            base[index] += to_number(
                cell(
                    row,
                    index,
                )
            )

        current_date = str(
            cell(
                row,
                date_index,
            )
            or ""
        ).strip()

        stored_date = str(
            cell(
                base,
                date_index,
            )
            or ""
        ).strip()

        if (
            current_date
            and (
                not stored_date
                or current_date < stored_date
            )
        ):
            base[
                date_index
            ] = current_date

    return (
        groups,
        accepted_rows,
        skipped_blank,
        invalid_rows,
    )


# =============================================================================
# Output builders
# =============================================================================

def build_brand_tabs(
    groups,
    i_gmv,
    tab_suffix,
):
    tabs = {}

    for (
        brand,
        _content_id,
    ), row in groups.items():
        jenis = (
            "GMV"
            if to_number(
                cell(
                    row,
                    i_gmv,
                )
            ) > 0
            else "NON GMV"
        )

        title = (
            f"{brand}"
            f"{tab_suffix}"
        )

        tabs.setdefault(
            title,
            [],
        ).append(
            list(row)
            + [jenis]
        )

    for title in tabs:
        tabs[
            title
        ].sort(
            key=lambda row: to_number(
                cell(
                    row,
                    i_gmv,
                )
            ),
            reverse=True,
        )

    return dict(
        sorted(
            tabs.items()
        )
    )


def build_gabungan(tabs):
    rows = []

    for tab_rows in tabs.values():
        rows.extend(
            tab_rows
        )

    return rows


def validate_brand_column(
    header,
    rows,
    brand_header=OUTPUT_BRAND_HEADER,
):
    brand_index = find_header_index(
        header,
        brand_header,
    )

    blank_count = sum(
        1
        for row in rows
        if not str(
            cell(
                row,
                brand_index,
            )
            or ""
        ).strip()
    )

    if blank_count:
        raise ValueError(
            f'Output memiliki {blank_count:,} row '
            f'dengan "{brand_header}" kosong.'
        )

    return brand_index


# =============================================================================
# Reconciliation
# =============================================================================

def reconcile_sum_columns(
    header,
    accepted_rows,
    groups,
    sum_columns,
):
    results = []

    for column in sum_columns:
        index = find_header_index(
            header,
            column,
        )

        source_total = sum(
            to_number(
                cell(
                    row,
                    index,
                )
            )
            for row in accepted_rows
        )

        aggregate_total = sum(
            to_number(
                cell(
                    row,
                    index,
                )
            )
            for row in groups.values()
        )

        results.append(
            {
                "column": column,
                "source": source_total,
                "aggregate": aggregate_total,
                "ok": numbers_equal(
                    source_total,
                    aggregate_total,
                ),
            }
        )

    return results


def reconcile_groups_vs_gabungan(
    header,
    groups,
    gabungan_rows,
    sum_columns,
):
    results = []

    for column in sum_columns:
        index = find_header_index(
            header,
            column,
        )

        group_total = sum(
            to_number(
                cell(
                    row,
                    index,
                )
            )
            for row in groups.values()
        )

        gabungan_total = sum(
            to_number(
                cell(
                    row,
                    index,
                )
            )
            for row in gabungan_rows
        )

        results.append(
            {
                "column": column,
                "groups": group_total,
                "gabungan": gabungan_total,
                "ok": numbers_equal(
                    group_total,
                    gabungan_total,
                ),
            }
        )

    return results


def metric_totals(
    header,
    rows,
    columns,
):
    totals = {}

    for column in columns:
        index = find_header_index(
            header,
            column,
        )

        totals[
            column
        ] = sum(
            to_number(
                cell(
                    row,
                    index,
                )
            )
            for row in rows
        )

    return totals


# =============================================================================
# Incremental monthly storage
# =============================================================================
# =============================================================================
# Incremental monthly storage
# =============================================================================

LOAD_LOG_SHEET = "_LOAD_LOG"

LOAD_LOG_HEADERS = [
    "Pipeline",
    "Load ID",
    "Source Spreadsheet ID",
    "Source Sheet",
    "Source Fingerprint",
    "Code Version",
    "Action",
    "Source Rows",
    "Stored ALL Rows",
    "Final ALL Rows",
    "Final RANGE Rows",
    "Range Start",
    "Range End",
    "Status",
    "Loaded At",
]

MASTER_META_HEADERS = [
    "_Pipeline",
    "_Load ID",
    "_Source Spreadsheet ID",
    "_Source Fingerprint",
    "_Code Version",
    "_Loaded At",
]


def table_fingerprint(header, rows):
    """Return a stable SHA-256 fingerprint for a source table."""
    digest = hashlib.sha256()

    def update_digest(row):
        payload = json.dumps(
            row,
            ensure_ascii=False,
            separators=(",", ":"),
            default=str,
        )
        digest.update(payload.encode("utf-8"))
        digest.update(b"\n")

    update_digest(header)

    for row in rows:
        update_digest(row)

    return digest.hexdigest()


def resolve_load_id(explicit_load_id, source_spreadsheet_id):
    """
    Resolve the persistent identity of one source pull.

    When LOAD_ID is blank, the source spreadsheet ID is used. If the same
    spreadsheet is overwritten for different pulls, configure a different
    LOAD_ID for each pull so historical contributions remain independent.
    """
    load_id = str(explicit_load_id or "").strip()

    if load_id:
        return load_id

    source_id = str(source_spreadsheet_id or "").strip()

    if not source_id:
        raise ValueError("Source Spreadsheet ID kosong.")

    return source_id


def get_spreadsheet_sheets(service, spreadsheet_id):
    """Return sheet metadata keyed by title."""
    response = execute_with_retry(
        service
        .spreadsheets()
        .get(
            spreadsheetId=spreadsheet_id,
            fields=(
                "sheets("
                "properties("
                "sheetId,"
                "title,"
                "gridProperties"
                ")"
                ")"
            ),
        )
    )

    result = {}

    for sheet in response.get("sheets", []):
        properties = sheet.get("properties", {})
        title = properties.get("title", "")

        if title:
            result[title] = properties

    return result


def read_table_if_exists(
    service,
    spreadsheet_id,
    sheet_title,
    source_range="A:AZ",
):
    """Read a table when the target sheet exists; otherwise return an empty table."""
    sheets = get_spreadsheet_sheets(
        service,
        spreadsheet_id,
    )

    if sheet_title not in sheets:
        return [], []

    response = execute_with_retry(
        service
        .spreadsheets()
        .values()
        .get(
            spreadsheetId=spreadsheet_id,
            range=f"'{sheet_title}'!{source_range}",
            valueRenderOption="UNFORMATTED_VALUE",
            dateTimeRenderOption="FORMATTED_STRING",
        )
    )

    values = response.get("values", [])

    if not values:
        return [], []

    return values[0], values[1:]


def validate_header_layout(actual_header, expected_header, label):
    """Require an existing persistent table to use the expected column layout."""
    if not actual_header:
        return

    actual = [normalize_header(value) for value in actual_header]
    expected = [normalize_header(value) for value in expected_header]

    if actual != expected:
        raise ValueError(
            f'Header {label} tidak sesuai schema v5.3. '
            "Gunakan output workbook baru atau hapus tab sistem versi lama "
            "sebelum menjalankan v5.3.\n"
            f"Expected: {expected_header}\n"
            f"Actual  : {actual_header}"
        )


def latest_load_record(
    log_header,
    log_rows,
    pipeline,
    load_id,
):
    """Return the latest successful load record for one pipeline and load ID."""
    if not log_header or not log_rows:
        return None

    required = [
        "Pipeline",
        "Load ID",
        "Status",
    ]

    indexes = {
        name: find_header_index(log_header, name)
        for name in required
    }

    matches = []

    for row in log_rows:
        row_pipeline = str(
            cell(row, indexes["Pipeline"]) or ""
        ).strip().upper()

        row_load_id = str(
            cell(row, indexes["Load ID"]) or ""
        ).strip()

        row_status = str(
            cell(row, indexes["Status"]) or ""
        ).strip().upper()

        if (
            row_pipeline == str(pipeline).strip().upper()
            and row_load_id == str(load_id).strip()
            and row_status == "SUCCESS"
        ):
            matches.append(row)

    if not matches:
        return None

    latest = matches[-1]

    return {
        name: cell(
            latest,
            find_header_index(log_header, name),
        )
        for name in log_header
    }


def determine_load_action(
    latest_record,
    source_spreadsheet_id,
    source_fingerprint,
    code_version,
    range_start=None,
    range_end=None,
    force=False,
):
    """Classify the current source pull as ADD, SKIP, or REPLACE.

    The configured Date Live/Post date range is part of the persistent load
    definition. Changing the range for an existing load ID therefore triggers
    a replacement so the cumulative RANGE output can be rebuilt correctly.
    """
    if latest_record is None:
        return "ADD", "load ID belum pernah dimuat"

    if force:
        return "REPLACE", "--force digunakan"

    same_source = (
        str(latest_record.get("Source Spreadsheet ID", "")).strip()
        == str(source_spreadsheet_id).strip()
    )

    same_fingerprint = (
        str(latest_record.get("Source Fingerprint", "")).strip()
        == str(source_fingerprint).strip()
    )

    same_version = (
        str(latest_record.get("Code Version", "")).strip()
        == str(code_version).strip()
    )

    same_range = True

    if range_start is not None or range_end is not None:
        same_range = (
            str(latest_record.get("Range Start", "")).strip()
            == str(range_start or "").strip()
            and str(latest_record.get("Range End", "")).strip()
            == str(range_end or "").strip()
        )

    if same_source and same_fingerprint and same_version and same_range:
        return (
            "SKIP",
            "load ID, source, fingerprint, versi code, dan range sama",
        )

    reasons = []

    if not same_source:
        reasons.append("source berubah")

    if not same_fingerprint:
        reasons.append("isi source berubah")

    if not same_version:
        reasons.append("versi code berubah")

    if not same_range:
        reasons.append("range berubah")

    return "REPLACE", ", ".join(reasons)


def latest_successful_load_records(
    log_header,
    log_rows,
    pipeline,
):
    """Return the latest successful log record for each load ID."""
    if not log_header or not log_rows:
        return {}

    required = [
        "Pipeline",
        "Load ID",
        "Status",
    ]

    indexes = {
        name: find_header_index(log_header, name)
        for name in required
    }

    target_pipeline = str(pipeline).strip().upper()
    records = {}

    for row in log_rows:
        row_pipeline = str(
            cell(row, indexes["Pipeline"]) or ""
        ).strip().upper()
        row_load_id = str(
            cell(row, indexes["Load ID"]) or ""
        ).strip()
        row_status = str(
            cell(row, indexes["Status"]) or ""
        ).strip().upper()

        if (
            row_pipeline != target_pipeline
            or not row_load_id
            or row_status != "SUCCESS"
        ):
            continue

        records[row_load_id] = {
            name: cell(
                row,
                find_header_index(log_header, name),
            )
            for name in log_header
        }

    return records


def merge_date_ranges(ranges):
    """Merge overlapping or adjacent date ranges."""
    if not ranges:
        return []

    normalized = sorted(ranges, key=lambda item: (item[0], item[1]))
    merged = [[normalized[0][0], normalized[0][1]]]

    for start_date, end_date in normalized[1:]:
        current = merged[-1]

        if start_date <= current[1] + timedelta(days=1):
            if end_date > current[1]:
                current[1] = end_date
        else:
            merged.append([start_date, end_date])

    return [tuple(item) for item in merged]


def cumulative_load_ranges(
    log_header,
    log_rows,
    pipeline,
    current_load_id,
    current_range_start,
    current_range_end,
):
    """Build cumulative Date Live/Post date ranges from successful loads.

    Each load ID contributes one active range. Historical records for the same
    load ID are superseded by its latest successful record. The current run is
    applied as an in-memory override before its SUCCESS log row is written.
    """
    records = latest_successful_load_records(
        log_header,
        log_rows,
        pipeline,
    )

    records[str(current_load_id).strip()] = {
        "Range Start": current_range_start,
        "Range End": current_range_end,
    }

    ranges = []

    for load_id, record in records.items():
        start_text = str(record.get("Range Start", "") or "").strip()
        end_text = str(record.get("Range End", "") or "").strip()

        if not start_text or not end_text:
            raise ValueError(
                f"Load ID {load_id!r} tidak memiliki Range Start/End yang valid."
            )

        start_date = parse_config_date(
            start_text,
            f"Range Start untuk load {load_id}",
        )
        end_date = parse_config_date(
            end_text,
            f"Range End untuk load {load_id}",
        )

        if start_date > end_date:
            raise ValueError(
                f"Range load {load_id!r} tidak valid: "
                f"{start_text} > {end_text}."
            )

        ranges.append((start_date, end_date))

    return merge_date_ranges(ranges)


def build_master_header(output_header):
    """Append persistent source-load metadata columns to the master schema."""
    return list(output_header) + list(MASTER_META_HEADERS)


def master_load_indexes(master_header):
    return {
        "pipeline": find_header_index(master_header, "_Pipeline"),
        "load_id": find_header_index(master_header, "_Load ID"),
    }


def attach_master_metadata(
    output_rows,
    pipeline,
    load_id,
    source_spreadsheet_id,
    source_fingerprint,
    code_version,
    loaded_at,
):
    """Add source-load metadata to processed ALL rows before persistent storage."""
    metadata = [
        pipeline,
        load_id,
        source_spreadsheet_id,
        source_fingerprint,
        code_version,
        loaded_at,
    ]

    return [
        list(row) + metadata
        for row in output_rows
    ]


def prepare_master_upsert(
    existing_header,
    existing_rows,
    expected_header,
    new_rows,
    pipeline,
    load_id,
):
    """
    Prepare an idempotent master update.

    A new load ID is appended. If the same load ID already exists, its stored
    contribution is replaced before the monthly result is rebuilt.
    """
    validate_header_layout(
        existing_header,
        expected_header,
        "master",
    )

    if not existing_header:
        return list(new_rows), "REPLACE"

    indexes = master_load_indexes(existing_header)
    target_pipeline = str(pipeline).strip().upper()
    target_load_id = str(load_id).strip()

    kept_rows = []
    matched = 0

    for row in existing_rows:
        is_target = (
            str(cell(row, indexes["pipeline"]) or "").strip().upper()
            == target_pipeline
            and str(cell(row, indexes["load_id"]) or "").strip()
            == target_load_id
        )

        if is_target:
            matched += 1
        else:
            kept_rows.append(row)

    if matched == 0:
        return list(existing_rows) + list(new_rows), "APPEND"

    return kept_rows + list(new_rows), "REPLACE"


def strip_master_metadata(master_header, master_rows, output_header):
    """Return the output portion of persistent master rows."""
    validate_header_layout(
        master_header,
        build_master_header(output_header),
        "master",
    )

    width = len(output_header)

    return [
        [cell(row, index) for index in range(width)]
        for row in master_rows
    ]


def validate_group_field_consistency(
    header,
    rows,
    key_headers,
    field_header,
    max_examples=20,
):
    """
    Validate that one content key does not carry conflicting reporting dates.

    The reporting field is expected to be stable for a content ID across all
    stored source pulls. Blank values are ignored when a non-blank value exists.
    """
    key_indexes = [
        find_header_index(header, name)
        for name in key_headers
    ]
    field_index = find_header_index(header, field_header)

    values_by_key = defaultdict(set)

    for row in rows:
        key = tuple(
            str(cell(row, index) or "").strip()
            for index in key_indexes
        )
        value = str(cell(row, field_index) or "").strip()

        if all(key) and value:
            values_by_key[key].add(value)

    conflicts = [
        (key, sorted(values))
        for key, values in values_by_key.items()
        if len(values) > 1
    ]

    if conflicts:
        examples = "\n".join(
            f"  {key}: {values}"
            for key, values in conflicts[:max_examples]
        )

        extra = (
            f"\n  ... {len(conflicts) - max_examples:,} conflict lainnya."
            if len(conflicts) > max_examples
            else ""
        )

        raise ValueError(
            f'Nilai "{field_header}" tidak konsisten untuk content key yang sama:\n'
            + examples
            + extra
        )

    return len(values_by_key)


# =============================================================================
# Spreadsheet capacity
# =============================================================================

def spreadsheet_grid_info(service, spreadsheet_id):
    sheets = get_spreadsheet_sheets(
        service,
        spreadsheet_id,
    )

    info = {}
    total_cells = 0

    for title, properties in sheets.items():
        grid = properties.get("gridProperties", {})
        row_count = grid.get("rowCount", 0)
        column_count = grid.get("columnCount", 0)
        cells = row_count * column_count
        total_cells += cells

        info[title] = {
            "sheetId": properties["sheetId"],
            "rowCount": row_count,
            "columnCount": column_count,
            "cells": cells,
        }

    return info, total_cells


def validate_target_capacity(
    service,
    spreadsheet_id,
    target_dimensions,
    label="OUTPUT",
    max_cells=MAX_CELLS,
):
    """Validate projected capacity when selected sheets are resized or created."""
    sheet_info, existing_cells = spreadsheet_grid_info(
        service,
        spreadsheet_id,
    )

    replaced_cells = sum(
        sheet_info.get(title, {}).get("cells", 0)
        for title in target_dimensions
    )

    target_cells = sum(
        max(int(rows), 2) * max(int(columns), 1)
        for rows, columns in target_dimensions.values()
    )

    projected_cells = (
        existing_cells
        - replaced_cells
        + target_cells
    )

    print(
        f"[{label}] sel sekarang {existing_cells:,} "
        f"- grid target lama {replaced_cells:,} "
        f"+ grid target baru {target_cells:,} "
        f"= {projected_cells:,} "
        f"(batas {max_cells:,})."
    )

    if projected_cells > max_cells:
        raise SystemExit(
            f"[{label}] Projected cell melebihi {max_cells:,}. "
            "Tidak ada output yang ditulis."
        )

    return existing_cells, target_cells, projected_cells


# =============================================================================
# Incremental writer
# =============================================================================

WRITE_BATCH_RANGES = 4


def _sheet_id(service, spreadsheet_id, sheet_title):
    sheets = get_spreadsheet_sheets(
        service,
        spreadsheet_id,
    )

    properties = sheets.get(sheet_title)

    return None if properties is None else properties["sheetId"]


def ensure_sheet_dimensions(
    service,
    spreadsheet_id,
    sheet_title,
    row_count,
    column_count,
    frozen_row_count=1,
):
    """Create a sheet when needed and set its grid to the requested dimensions."""
    row_count = max(int(row_count), 2)
    column_count = max(int(column_count), 1)

    sheet_id = _sheet_id(
        service,
        spreadsheet_id,
        sheet_title,
    )

    if sheet_id is None:
        response = execute_with_retry(
            service
            .spreadsheets()
            .batchUpdate(
                spreadsheetId=spreadsheet_id,
                body={
                    "requests": [
                        {
                            "addSheet": {
                                "properties": {
                                    "title": sheet_title,
                                    "gridProperties": {
                                        "rowCount": row_count,
                                        "columnCount": column_count,
                                        "frozenRowCount": frozen_row_count,
                                    },
                                }
                            }
                        }
                    ]
                },
            )
        )

        return response["replies"][0]["addSheet"]["properties"]["sheetId"]

    execute_with_retry(
        service
        .spreadsheets()
        .batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={
                "requests": [
                    {
                        "updateSheetProperties": {
                            "properties": {
                                "sheetId": sheet_id,
                                "gridProperties": {
                                    "rowCount": row_count,
                                    "columnCount": column_count,
                                    "frozenRowCount": frozen_row_count,
                                },
                            },
                            "fields": (
                                "gridProperties.rowCount,"
                                "gridProperties.columnCount,"
                                "gridProperties.frozenRowCount"
                            ),
                        }
                    }
                ]
            },
        )
    )

    return sheet_id


def clear_sheet_values(service, spreadsheet_id, sheet_title):
    execute_with_retry(
        service
        .spreadsheets()
        .values()
        .clear(
            spreadsheetId=spreadsheet_id,
            range=f"'{sheet_title}'",
            body={},
        )
    )


def _write_value_ranges(
    service,
    spreadsheet_id,
    value_ranges,
    batch_range_count=WRITE_BATCH_RANGES,
):
    """Write several non-overlapping ranges per Sheets API request."""
    if not value_ranges:
        return

    for start in range(
        0,
        len(value_ranges),
        batch_range_count,
    ):
        batch = value_ranges[
            start:
            start + batch_range_count
        ]

        execute_with_retry(
            service
            .spreadsheets()
            .values()
            .batchUpdate(
                spreadsheetId=spreadsheet_id,
                body={
                    "valueInputOption": "RAW",
                    "data": batch,
                },
            )
        )


def write_rows_at(
    service,
    spreadsheet_id,
    sheet_title,
    rows,
    start_row,
    chunk_rows=WRITE_CHUNK_ROWS,
):
    """Write rows from a known starting row using batched range updates."""
    if not rows:
        return

    value_ranges = []

    for offset in range(0, len(rows), chunk_rows):
        chunk = rows[offset:offset + chunk_rows]
        row_number = start_row + offset

        value_ranges.append(
            {
                "range": f"'{sheet_title}'!A{row_number}",
                "values": chunk,
            }
        )

    _write_value_ranges(
        service,
        spreadsheet_id,
        value_ranges,
    )


def replace_table(
    service,
    spreadsheet_id,
    sheet_title,
    header,
    rows,
):
    """Replace one sheet table while preserving the sheet name."""
    ensure_sheet_dimensions(
        service,
        spreadsheet_id,
        sheet_title,
        len(rows) + 1,
        len(header),
    )

    clear_sheet_values(
        service,
        spreadsheet_id,
        sheet_title,
    )

    write_rows_at(
        service,
        spreadsheet_id,
        sheet_title,
        [list(header)] + list(rows),
        start_row=1,
    )


def append_table_rows(
    service,
    spreadsheet_id,
    sheet_title,
    header,
    existing_row_count,
    new_rows,
):
    """Append rows to an existing persistent table without rewriting prior data."""
    if not new_rows:
        return

    sheet_id = _sheet_id(
        service,
        spreadsheet_id,
        sheet_title,
    )

    if sheet_id is None or existing_row_count == 0:
        replace_table(
            service,
            spreadsheet_id,
            sheet_title,
            header,
            new_rows,
        )
        return

    ensure_sheet_dimensions(
        service,
        spreadsheet_id,
        sheet_title,
        existing_row_count + len(new_rows) + 1,
        len(header),
    )

    write_rows_at(
        service,
        spreadsheet_id,
        sheet_title,
        new_rows,
        start_row=existing_row_count + 2,
    )


def upsert_master_table(
    service,
    spreadsheet_id,
    sheet_title,
    expected_header,
    existing_header,
    existing_rows,
    merged_rows,
    source_rows,
    write_mode,
):
    """Persist a prepared master update using append when safe and replace otherwise."""
    if write_mode == "APPEND" and existing_header:
        append_table_rows(
            service,
            spreadsheet_id,
            sheet_title,
            expected_header,
            len(existing_rows),
            source_rows,
        )
    else:
        replace_table(
            service,
            spreadsheet_id,
            sheet_title,
            expected_header,
            merged_rows,
        )


def append_load_log(
    service,
    spreadsheet_id,
    row,
    existing_header,
    existing_rows,
):
    """Append one successful run record to the persistent load log."""
    validate_header_layout(
        existing_header,
        LOAD_LOG_HEADERS,
        LOAD_LOG_SHEET,
    )

    append_table_rows(
        service,
        spreadsheet_id,
        LOAD_LOG_SHEET,
        LOAD_LOG_HEADERS,
        len(existing_rows),
        [row],
    )


# =============================================================================
# Monthly output rebuild
# =============================================================================

def build_combined_from_groups(
    groups,
    header,
    gmv_column,
):
    """Build a combined output using the same brand ordering as the legacy flow."""
    gmv_index = find_header_index(
        header,
        gmv_column,
    )

    brand_tabs = build_brand_tabs(
        groups,
        gmv_index,
        "",
    )

    return build_gabungan(brand_tabs)


def rebuild_monthly_output(
    master_header,
    master_rows,
    output_header,
    content_id_header,
    date_header,
    sum_columns,
    gmv_column,
    brand_header=OUTPUT_BRAND_HEADER,
    jenis_header="Jenis",
    consistent_fields=None,
):
    """
    Re-aggregate all stored source pulls to the monthly Brand + Content ID grain.

    RANGE filtering is intentionally not performed here. The caller must first
    complete this monthly aggregation and only then filter the resulting groups
    by Date Live or Post date.
    """
    if not output_header:
        raise ValueError("Output header kosong.")

    if normalize_header(output_header[-1]) != normalize_header(jenis_header):
        raise ValueError(
            f'Kolom terakhir output harus "{jenis_header}".'
        )

    stored_rows = strip_master_metadata(
        master_header,
        master_rows,
        output_header,
    )

    base_header = list(output_header[:-1])
    base_width = len(base_header)

    aggregation_rows = [
        [cell(row, index) for index in range(base_width)]
        for row in stored_rows
    ]

    for field_header in consistent_fields or []:
        validate_group_field_consistency(
            base_header,
            aggregation_rows,
            [brand_header, content_id_header],
            field_header,
        )

    (
        groups,
        accepted_rows,
        skipped_blank,
        invalid_rows,
    ) = aggregate_brand_content(
        base_header,
        aggregation_rows,
        content_id_header=content_id_header,
        date_header=date_header,
        sum_columns=sum_columns,
        brand_header=brand_header,
        row_validator=None,
    )

    if invalid_rows:
        raise ValueError(
            "Monthly rebuild menghasilkan row invalid yang tidak diharapkan."
        )

    combined = build_combined_from_groups(
        groups,
        base_header,
        gmv_column,
    )

    validate_brand_column(
        output_header,
        combined,
        brand_header,
    )

    checks = reconcile_sum_columns(
        base_header,
        accepted_rows,
        groups,
        sum_columns,
    )

    if not all(item["ok"] for item in checks):
        raise ValueError(
            "Monthly reconciliation gagal."
        )

    return {
        "groups": groups,
        "rows": combined,
        "accepted_rows": accepted_rows,
        "skipped_blank": skipped_blank,
        "checks": checks,
        "base_header": base_header,
    }
