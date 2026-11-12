"""TAP/SC staging, read-only preview and verified Google Sheets writes."""
from __future__ import annotations

import csv
import hashlib
import importlib.util
import json
import math
import os
import re
import sys
import time
from collections import Counter, defaultdict
from datetime import date, datetime, timedelta, timezone
from decimal import Decimal, InvalidOperation
from itertools import zip_longest
from pathlib import Path
from urllib.parse import parse_qs, urlparse

ROOT = Path(__file__).resolve().parent

NOVEMBER_ORIGINAL_ID = "1fWzScgmB8ufdz3t3V-HeRZD1Ras0KfKNutytNuL2eV8"

# Company workbooks are blocked by default. The November ID has an explicit opt-in.
BLOCKED_ORIGINAL_IDS = {
    NOVEMBER_ORIGINAL_ID,
    "1P7YjbAb6RlEhw5h47M9WU7KXpfVLgoJYna3nnlrUQvQ",
    "1vdTLs_mqMIaV0SuNtA5NY5pqAfUqgbsq441qh-TpR-k",
    "1MiYHTWLTuWBtTKHDHSqxTCtiWRlHK3xN5DcSid-pwgY",
}

PACE_SECONDS = 1.10
CHUNK_ROWS = 2_400             # max rows per atomic Sheets batchUpdate
LEGACY_CHUNK_ROWS = 600
# ============================================================================

MAX_BATCH_BYTES = 1_500_000
SCAN_ROWS = 10_000
PLAN_VERSION = 8
LEGACY_PLAN_VERSION = 5
# Earlier releases stage rows differently but use the same write rules for frozen data.
PRIOR_PIPELINE_SHA256 = {
    5: "871b4f56e4d454787f8b0eb86f7fe6af52bf52db725a0861277185113fe0e1ae",
    6: "f2b95708d7ef86ab9eac35367deb9c0ac23104b1266651e67e2945c99ced5b10",
    7: "c677ce263301dc1af67837bb1543715a9b6c76765852c958b4d24f7c73330168",
}
PRIOR_TAP_LS_SHA256 = "d22e01c6446fae89b2382f56b0108a4eceec061946b841a7d53ceb7d2a604a33"

BACKENDS = {
    "tap_pc_core": "tap_pc_core.py",
    "tap_sv": "split_brand_sv_combined.py",
    "tap_ls": "split_brand_live_combined.py",
    "sc_sv": "sc_pc_video_ebu_quota_safe.py",
    "sc_ls": "sc_pc_live_ebu_fixed_alias_quota_safe.py",
}
TABS = {"SV": "RAW SV", "LS": "RAW LS"}
FORMULA_FIELDS = {"SV": ("Date", "Type Video"), "LS": ("Date",)}
MANAGED_FORMULAS = {"SV": {"Date", "Type Video", "Video link"}, "LS": {"Date", "Duration"}}
BRANDS = {"PEPSODENT", "CLOSE UP", "LIFEBUOY", "DOVE SCL", "DOVE DEO", "LUX", "REXONA"}
# ONLY these columns can be written. Columns like Creator Check/formulas are NEVER touched.
INPUTS = {
    "SV": ["Source", "Post date", "Video title", "Video ID", "Video link", "Creator name", "Product ID",
           "Product name", "Brand", "Creator video-attributed GMV", "Video-attributed orders",
           "Video-attributed items sold", "Likes", "Comments", "Shares", "Video product impressions",
           "Video product clicks", "Video views"],
    "LS": ["Source", "Live Date", "Creator name", "Creator follower count", "Product ID", "Product name",
           "Brand", "Livestream name", "Livestream room ID", "Live Start", "Live End",
           "Creator LIVE-attributed GMV", "Creator-attributed items sold", "Creator LIVE-attributed orders",
           "Duration", "LIVE views", "LIVE likes"],
}
ALIAS = {
    "SV": {
        "Post date": ["Post date", "Post time"], "Video title": ["Video title", "Video name"],
        "Video ID": ["Video ID"], "Video link": ["Video link"], "Creator name": ["Creator name"],
        "Product ID": ["Product ID"], "Product name": ["Product name"], "Brand": ["Brand"],
        "Creator video-attributed GMV": ["Creator video-attributed GMV", "Affiliate video-attributed GMV"],
        "Video-attributed orders": ["Video-attributed orders", "Creator video-attributed orders", "Affiliate video orders"],
        "Video-attributed items sold": ["Video-attributed items sold", "Creator-attributed items sold"],
        "Likes": ["Likes", "Video likes"], "Comments": ["Comments", "Video comments"],
        "Shares": ["Shares", "Video shares"], "Video product impressions": ["Video product impressions"],
        "Video product clicks": ["Video product clicks"], "Video views": ["Video views"],
    },
    "LS": {
        "Live Date": ["Date Live", "Live Date", "LIVE start time", "Start Live"],
        "Creator name": ["Creator name"], "Creator follower count": ["Creator follower count"],
        "Product ID": ["Product ID"], "Product name": ["Product name"], "Brand": ["Brand"],
        "Livestream name": ["Livestream name", "LIVE title"],
        "Livestream room ID": ["Livestream room ID", "LIVE ID"],
        "Live Start": ["Start Live", "LIVE start time", "Live Start"],
        "Live End": ["End Live", "LIVE end time", "Live End"],
        "Creator LIVE-attributed GMV": ["Creator LIVE-attributed GMV"],
        "Creator-attributed items sold": ["Creator-attributed items sold", "LIVE-attributed items sold"],
        "Creator LIVE-attributed orders": ["Creator LIVE-attributed orders", "LIVE-attributed orders"],
        "Duration": ["Duration"], "LIVE views": ["LIVE views"], "LIVE likes": ["LIVE likes", "Likes"],
    },
}
NUMBERS = {"Creator follower count", "Creator video-attributed GMV", "Video-attributed orders",
           "Video-attributed items sold", "Likes", "Comments", "Shares", "Video product impressions",
           "Video product clicks", "Video views", "Creator LIVE-attributed GMV", "Creator-attributed items sold",
           "Creator LIVE-attributed orders", "LIVE views", "LIVE likes"}
# Source-specific ABSENT fields. Blank means unavailable, NOT numeric zero.
# Explicitly ignore those fields even if a backend accidentally emits a column.
SOURCE_UNAVAILABLE = {
    "SV": {"TAP": {"Comments", "Shares", "Video product impressions", "Video product clicks"}},
    "LS": {"SC": {"Duration", "LIVE views"}},
}
DATE_FIELDS = {"Post date", "Live Date"}
TIME_FIELDS = {"Live Start", "Live End"}
ID_FIELDS = {"Video ID", "Livestream room ID", "Product ID"}


class StopRun(RuntimeError):
    pass

# ============================================================================
# LOCAL CSV/XLSX INPUTS: the unchanged calculation backends consume these rows.
# No Google service is constructed or called by TAP/SC/AFF source processing.
# ============================================================================
def file_spec(item):
    if isinstance(item, str):
        return {'path': item, 'sheet': None}
    if isinstance(item, dict) and item.get('path'):
        return {'path': str(item['path']), 'sheet': item.get('sheet') or None}
    raise StopRun(f'Invalid INPUT_FILES entry {item!r}; expected path or dict(path, sheet)')


def local_path(spec):
    p = (ROOT / file_spec(spec)['path']).resolve()
    # Allow absolute paths or paths inside the project, but never allow input file to be a source of executable code.
    if p.suffix.lower() not in {'.csv', '.xlsx'}:
        raise StopRun(f'Only .csv or .xlsx supported: {p}')
    if not p.is_file():
        raise StopRun(f'Local input file missing: {p}')
    return p


def excel_cell(value, heading, label):
    if value is None:
        return ''
    if isinstance(value, datetime):
        if label.startswith('SC_'):
            return value.strftime('%m/%d/%Y %H:%M:%S')
        if label.startswith('AFF_'):
            return value.strftime('%d/%m/%Y %H:%M:%S')
        return value.strftime('%Y-%m-%d %H:%M:%S')
    if isinstance(value, date):
        if label.startswith('SC_'):
            return value.strftime('%m/%d/%Y')
        if label.startswith('AFF_'):
            return value.strftime('%d/%m/%Y')
        return value.isoformat()
    h = norm(heading)
    if isinstance(value, (int, float)) and (h.endswith('id') or 'productid' in h or h.endswith('code')):
        if abs(value) >= 1e15 or (isinstance(value, float) and not value.is_integer()):
            raise StopRun(f'{label}: Excel numeric identifier {heading}={value!r} is unsafe. '
                          'Export from source as CSV containing exact ID digits or format IDs as TEXT before export.')
        return str(int(value))
    if isinstance(value, float):
        return str(int(value)) if value.is_integer() else str(value)
    return str(value)


def read_local_file(label, entry, required, cache):
    spec = file_spec(entry)
    path = local_path(entry)
    cache_key = (str(path), spec['sheet'], label)
    required_norm = {norm(h) for h in required}
    cached = cache.get(cache_key)
    if cached and required_norm.issubset({norm(h) for h in cached[0]}):
        return cached

    def find_header(rows):
        for offset, line in enumerate(rows[:20]):
            header = [str(c).strip() for c in line]
            keys = {norm(c) for c in header if c}
            virtual_product_name = (label.startswith('TAP_') and norm('Product name') in required_norm
                                    and norm('Product name') not in keys
                                    and (required_norm - {norm('Product name')}).issubset(keys))
            if required_norm.issubset(keys) or virtual_product_name:
                if len(keys) != len([h for h in header if h]):
                    raise StopRun(f'{label}: duplicate normalized header in {path.name}')
                if virtual_product_name:
                    print(f'  WARNING {label}: Product name missing; using blank names (PID brand fallback if supplied)', flush=True)
                    header.append('Product name')
                result_rows = []
                for r in rows[offset + 1:]:
                    vals = list(r[:len(header)]) + [''] * max(0, len(header) - len(r))
                    if any(str(v).strip() for v in vals):
                        result_rows.append(vals)
                return header, result_rows, offset + 1
        return None

    result = None
    selected_sheet = ''
    if path.suffix.lower() == '.csv':
        if spec['sheet']:
            raise StopRun(f'{label}: a CSV has no sheets: {path.name}')
        data = None
        for encoding in ('utf-8-sig', 'cp1252'):
            try:
                with path.open('r', encoding=encoding, newline='') as fh:
                    sample = fh.read(16384)
                    fh.seek(0)
                    try:
                        delimiter = csv.Sniffer().sniff(sample, delimiters=',;\t').delimiter
                    except csv.Error:
                        delimiter = ','
                    data = [list(row) for row in csv.reader(fh, delimiter=delimiter)]
                break
            except UnicodeDecodeError:
                continue
        if data is None:
            raise StopRun(f'{label}: cannot decode CSV {path.name}')
        result = find_header(data)
        selected_sheet = '(CSV)'
    else:
        try:
            from openpyxl import load_workbook
        except ImportError as exc:
            raise StopRun('XLSX needs openpyxl. Install: pip install openpyxl') from exc
        wb = load_workbook(str(path), read_only=True, data_only=True)
        try:
            if spec['sheet'] and spec['sheet'] not in wb.sheetnames:
                raise StopRun(f'{label}: Excel sheet {spec["sheet"]!r} not found in {path.name}; found {wb.sheetnames}')
            candidates = [spec['sheet']] if spec['sheet'] else wb.sheetnames
            for sheet_name in candidates:
                ws = wb[sheet_name]
                # TikTok export XLSX sometimes advertises only A1:A1 although
                # the worksheet contains many rows/columns. With read_only=True,
                # openpyxl otherwise truncates every row to a single column.
                # Ignore unreliable worksheet dimension metadata and discover
                # real columns/rows while streaming. This NEVER edits the XLSX.
                ws.reset_dimensions()
                head_rows = list(ws.iter_rows(min_row=1, max_row=20, values_only=True))
                selection = find_header([[excel_cell(v, '', label) for v in r] for r in head_rows])
                if not selection:
                    continue
                header, _, row_index = selection
                rows = []
                for raw in ws.iter_rows(min_row=row_index + 1, values_only=True):
                    rec = [excel_cell(v, header[i] if i < len(header) else '', label)
                           for i, v in enumerate(raw[:len(header)])]
                    rec += [''] * max(0, len(header) - len(rec))
                    if any(x.strip() for x in rec):
                        rows.append(rec)
                result = (header, rows, row_index)
                selected_sheet = sheet_name
                break
        finally:
            wb.close()

    if result is None:
        raise StopRun(f'{label}: required headers not found in first 20 rows of {path.name}: {required}')
    header, rows, header_row = result
    out = (header, rows, {'path': str(path), 'sheet_title': selected_sheet,
                          'header_row': header_row, 'rows': len(rows), 'columns': len(header),
                          'sheet_id': 0, 'spreadsheet_id': 'LOCAL_' + label})
    cache[cache_key] = out
    print(f'  LOCAL {label}: {len(rows):,} rows in {path.name} [{selected_sheet}]', flush=True)
    return out


def mapping_local(cfg):
    configured = cfg['files'].get('PRODUCT_BRAND_MAPPING')
    if not configured:
        return {}, {}, {'sheet_title': 'DISABLED', 'header_row': 0, 'mapping_count': 0,
                         'duplicate_same': 0, 'blank_pid': 0, 'skipped_non_pc': 0}
    header, rows, meta = read_local_file('PRODUCT_BRAND_MAPPING', configured,
                                        ['Product ID TTS', 'BRAND FIX'], cfg['cache'])
    return parse_product_mapping(header, rows, meta)


def mapping_columns(header):
    columns = {norm(h): i for i, h in enumerate(header) if h}
    if not {'productidtts', 'brandfix'} <= columns.keys():
        raise StopRun('List Produk harus punya kolom Product ID TTS dan BRAND FIX.')
    names = tuple(columns[key] for key in ('pidname', 'productname', 'skuname') if key in columns)
    return columns['productidtts'], columns['brandfix'], names


def parse_product_mapping(header, rows, meta):
    pid_i, brand_i, name_columns = mapping_columns(header)
    # Backend's canonicalization keeps original brand alias rules.
    result, names = {}, {}
    stats = {'sheet_title': meta['sheet_title'], 'header_row': meta['header_row'],
             'mapping_count': 0, 'duplicate_same': 0, 'blank_pid': 0, 'skipped_non_pc': 0}
    aliases = {'PEPSODENT': 'PEPSODENT', 'CLOSE UP': 'CLOSE UP', 'CLOSEUP': 'CLOSE UP',
               'LIFEBUOY': 'LIFEBUOY', 'LUX': 'LUX', 'REXONA': 'REXONA',
               'REXONA MEN': 'REXONA', 'REXONA MAN': 'REXONA', 'CLOSE-UP': 'CLOSE UP', 'DOVE': 'DOVE SCL', 'DOVE SCL': 'DOVE SCL',
               'DOVE DEO': 'DOVE DEO', 'DOVE DEODORANT': 'DOVE DEO'}
    for row in rows:
        pid = exact_id(row[pid_i] if pid_i < len(row) else "")
        brand_value = row[brand_i] if brand_i < len(row) else ""
        brand = aliases.get(re.sub(r'\s+', ' ', str(brand_value or '').strip().upper()))
        if pid:
            for rank, name_i in enumerate(name_columns):
                name = str(row[name_i] or "").strip() if name_i < len(row) else ""
                if name and name.upper() not in {"-", "--", "N/A", "#N/A", "#REF!", "NULL"}:
                    if pid not in names or rank < names[pid][0]:
                        names[pid] = rank, name
                    break
        if not pid:
            stats['blank_pid'] += 1
        elif not brand:
            stats['skipped_non_pc'] += 1
        elif pid in result:
            if result[pid] != brand:
                raise StopRun(f'Product Brand Mapping has conflicting brands for Product ID {pid}')
            stats['duplicate_same'] += 1
        else:
            result[pid] = brand
    stats['mapping_count'] = len(result)
    return result, {pid: entry[1] for pid, entry in names.items()}, stats

def norm(v):
    return re.sub(r"[^a-z0-9]", "", str(v or "").lower())


def sid(value):
    s = str(value or "").strip()
    match = re.search(r"(?:https://docs\.google\.com/spreadsheets/d/)([a-zA-Z0-9_-]+)", s)
    if match:
        return match.group(1)
    if re.fullmatch(r"[a-zA-Z0-9_-]{20,}", s):
        return s
    if s == "":
        return ""
    raise StopRun(f"Invalid Google spreadsheet ID/URL: {value!r}")


def exact_id(value):
    """Return an exact ID as TEXT. Never approximate 16+ digit numeric cells.

    IDs provided as TEXT (CSV/XLSX strings/Google plain-text cells) retain all
    digits. The Sheets API may return numeric JSON values for old target cells;
    those values are not guaranteed to preserve the original 16+ digits.
    """
    if value is None or value == "":
        return ""
    if isinstance(value, bool):
        raise StopRun(f"Invalid boolean content ID {value!r}")
    if isinstance(value, (int, float)):
        if not math.isfinite(float(value)) or not float(value).is_integer():
            raise StopRun(f"Invalid numeric content ID {value!r}")
        if abs(value) >= 10**15:
            raise StopRun(f"Lossy numeric content ID {value!r}; exact digits cannot be recovered automatically")
        return str(int(value))
    s = str(value).strip().strip("'\"").replace("\u200b", "")
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    if re.search(r"[eE][+-]?\d+$", s) or ("." in s and not s.isdigit()):
        raise StopRun(f"Unsafe ID formatting {s!r}; original exact digits are required")
    if s and not re.fullmatch(r"\d+", s):
        raise StopRun(f"Non-decimal content ID {s!r}")
    return s


def _maybe_same_legacy_id(old, exact_digits):
    """Conservative collision test; never claim the old ID is exact.

    Google numeric IDs can be rounded (compare as IEEE float). Scientific text
    has limited significant digits (compare its displayed rounding interval).
    A positive result BLOCKS the write until the copied workbook is repaired.
    """
    if not exact_digits or not re.fullmatch(r"\d+", str(exact_digits)):
        return False
    try:
        if isinstance(old, (int, float)) and not isinstance(old, bool):
            return float(exact_digits) == float(old)
        token = str(old).strip().replace(" ", "")
        scientific = re.fullmatch(r"(\d+(?:\.(\d+))?)[eE]([+-]?\d+)", token)
        if scientific:
            d = Decimal(token)
            decimal_places = len(scientific.group(2) or "")
            unit = Decimal(10) ** (int(scientific.group(3)) - decimal_places)
            return abs(Decimal(str(exact_digits)) - d) <= unit / 2
        return str(exact_digits) == token
    except (InvalidOperation, ValueError, OverflowError):
        return True  # unknown legacy encoding: fail closed rather than guess


def standard_date(value, source=""):
    s = str(value or "").strip()
    if not s:
        return ""
    if re.fullmatch(r"\d{5}(?:\.0)?", s):
        return (date(1899, 12, 30) + timedelta(days=int(float(s)))).isoformat()
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}(?:[ T].*)?", s):
        try:
            return date.fromisoformat(s[:10]).isoformat()
        except ValueError:
            return ""
    # Deliberately accept both only when unambiguous, to prevent 05/06 confusion.
    p = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})", s)
    if p:
        a, b = int(p.group(1)), int(p.group(2))
        if a <= 12 and b <= 12 and source != "SC":
            raise StopRun(f"Ambiguous slash date {s!r}; use ISO YYYY-MM-DD")
        fmt = "%m/%d/%Y" if source == "SC" or a <= 12 else "%d/%m/%Y"
        try:
            return datetime.strptime(s[:len(p.group(0))], fmt).date().isoformat()
        except ValueError:
            return ""
    for fmt in ("%d-%m-%Y", "%d %b %Y"):
        try:
            return datetime.strptime(s, fmt).date().isoformat()
        except ValueError:
            pass
    return ""


def number(value):
    if value is None or str(value).strip().upper() in {"", "-", "--", "NULL", "N/A", "NA", "NONE", "NAN"}:
        return ""
    s = str(value).strip()
    try:
        d = Decimal(s.replace(",", ""))
        if not d.is_finite():
            raise InvalidOperation
        return format(d.normalize(), "f")
    except InvalidOperation as exc:
        raise StopRun(f"Invalid numeric field: {value!r}") from exc


def standard_datetime(value, source=""):
    if value is None or str(value).strip() == "":
        return ""
    if isinstance(value, (int, float)):
        # Sheet timestamps are floating serials; round to nearest second to
        # avoid binary float drift (e.g. 09:59:59 instead of 10:00:00).
        return (datetime(1899, 12, 30) + timedelta(seconds=round(float(value) * 86400))).strftime("%Y-%m-%d %H:%M:%S")
    s = str(value).strip()
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%dT%H:%M:%S",
                "%m/%d/%Y %H:%M:%S", "%m/%d/%Y %H:%M", "%d/%m/%Y %H:%M:%S"):
        try:
            # With ambiguous slashes SC is known to export US-format, and TAP is ISO.
            if source != "SC" and fmt.startswith("%m/") and re.match(r"^\d{1,2}/\d{1,2}/", s):
                a, b = map(int, s.split(" ")[0].split("/")[:2])
                if a <= 12 and b <= 12:
                    raise StopRun(f"Ambiguous datetime {s!r}; use ISO")
            return datetime.strptime(s, fmt).strftime("%Y-%m-%d %H:%M:%S")
        except ValueError:
            pass
    raise StopRun(f"Invalid datetime for {source}: {s!r}")


def clean_field(field, value, source=""):
    if field in ID_FIELDS:
        return exact_id(value)
    if field in DATE_FIELDS:
        return standard_date(value, source)
    if field in TIME_FIELDS:
        return standard_datetime(value, source)
    if field in NUMBERS:
        return number(value)
    if field == "Brand":
        return re.sub(r"\s+", " ", str(value or "").strip().upper())
    if field == "Creator name":
        return str(value or "").strip().lstrip("@").casefold()
    return re.sub(r"\s+", " ", str(value or "").strip())


def stage(kind, source, src):
    if source not in ("TAP", "SC"):
        raise StopRun(f"Invalid source: {source!r}")
    record = {"Source": source}
    present = {"Source"}
    header = {norm(h): h for h in src if norm(h)}
    if len(header) != sum(bool(norm(h)) for h in src):
        raise StopRun("Duplicate normalized OA header")
    absent = SOURCE_UNAVAILABLE.get(kind, {}).get(source, set())
    for field, candidates in ALIAS[kind].items():
        source_header = next((header[norm(h)] for h in candidates if norm(h) in header), None)
        if field in absent or source_header is None:
            record[field] = ""
            continue
        present.add(field)
        value = src[source_header]
        record[field] = clean_field(field, "" if value is None else str(value), source)
    record["_present"] = sorted(present)
    return record


def key(kind, row):
    return (row.get("Brand", ""), row.get("Video ID" if kind == "SV" else "Livestream room ID", ""))


def canonical(kind, row):
    return {k: clean_field(k, row.get(k, "")) for k in INPUTS[kind]}


def digest(obj):
    return hashlib.sha256(json.dumps(obj, ensure_ascii=False, sort_keys=True, separators=(",", ":")).encode()).hexdigest()


def is_valid(kind, row, cfg):
    if not all(key(kind, row)):
        return "MISSING_KEY"
    if row["Brand"] not in BRANDS:
        return "INVALID_BRAND"
    if not row.get("Creator name"):
        return "MISSING_CREATOR"
    d = row.get("Post date" if kind == "SV" else "Live Date", "")
    if not d:
        return "MISSING_DATE"
    if not cfg['start'] <= d <= cfg['end']:
        return "OUTSIDE_CUTOFF"
    return "OK"


def as_dict(header, rows):
    if len({norm(h) for h in header if h}) != len([h for h in header if h]):
        raise StopRun(f"Duplicate headers in OA output: {header}")
    return [{h: (r[i] if i < len(r) else "") for i, h in enumerate(header)} for r in rows]


def write_json(path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".tmp")
    with tmp.open("w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
        f.flush()
        os.fsync(f.fileno())
    tmp.replace(path)


def check_target_allowed(target, *, allow_november_original=False):
    if target == NOVEMBER_ORIGINAL_ID and allow_november_original is True:
        return
    if target == NOVEMBER_ORIGINAL_ID:
        raise StopRun(
            "Workbook asli November diblokir. Untuk mengaktifkannya, isi "
            "allow_november_original = true di bagian [pc_content] pada secrets.toml.")
    if target in BLOCKED_ORIGINAL_IDS:
        raise StopRun("Workbook asli ini diblokir. Gunakan salinan terpisah.")


def build_config(files, start, end, target="", *, allow_november_original=False):
    if start > end or (start.year, start.month) != (end.year, end.month):
        raise StopRun("Pilih rentang tanggal dalam satu bulan kalender.")
    for label in ("TAP_SV", "TAP_LS", "SC_SV", "SC_LS"):
        if not files.get(label):
            raise StopRun(f"Upload file {label}.")
    if not files.get("AFF"):
        raise StopRun("Upload minimal satu export AFF.")
    book = sid(target)
    check_target_allowed(book, allow_november_original=allow_november_original)
    return {"pc": book, "start": start.isoformat(), "end": end.isoformat(),
            "files": files, "cache": {}}


def check_write_armed(cfg):
    target = cfg.get("pc")
    if not target:
        raise StopRun("Target tulis kosong.")
    check_target_allowed(target, allow_november_original=cfg.get("allow_november_original"))
    if cfg.get("approved_copy") != target:
        raise StopRun("Konfirmasi target sebelum mengirim batch.")
    if target == NOVEMBER_ORIGINAL_ID and cfg.get("original_confirmed") is not True:
        raise StopRun("Konfirmasi upload ke workbook asli November terlebih dahulu.")


# ============================================================================
# READ-ONLY GOOGLE CLIENT + VALIDATION
# ============================================================================
class Google:
    def __init__(self, credential_info, readonly=True):
        from google.oauth2 import service_account
        from googleapiclient.discovery import build
        scope = "https://www.googleapis.com/auth/spreadsheets.readonly" if readonly else "https://www.googleapis.com/auth/spreadsheets"
        c = service_account.Credentials.from_service_account_info(credential_info, scopes=[scope])
        self.api = build("sheets", "v4", credentials=c, cache_discovery=False)
        self.last = 0.0
        self.readonly = readonly
        self.requests = 0

    def call(self, req, *, retry_write=False):
        from googleapiclient.errors import HttpError
        import http.client
        if self.readonly:
            # Sparse ranges can exceed httplib2's gzip expansion limit.
            req.headers['accept-encoding'] = 'identity'
        for n in range(6):
            wait = PACE_SECONDS - (time.monotonic() - self.last)
            if wait > 0:
                time.sleep(wait)
            self.last = time.monotonic()
            self.requests += 1
            try:
                return req.execute()
            except HttpError as exc:
                status = getattr(exc.resp, "status", None)
                if status not in {429, 500, 502, 503, 504} or n == 5 or (not self.readonly and not retry_write):
                    raise
                delay = 65 if status == 429 else min(2 ** (n + 1), 30)
                print(f'  Sheets read HTTP {status}: retry in {delay}s ({n + 1}/5)', flush=True)
                time.sleep(delay)
            except (OSError, http.client.HTTPException):
                if n == 5 or (not self.readonly and not retry_write):
                    raise
                delay = min(2 ** (n + 1), 30)
                print(f'  Sheets read connection error: retry in {delay}s ({n + 1}/5)', flush=True)
                time.sleep(delay)
        raise StopRun("Google API retries exhausted")

    def meta(self, book):
        return self.call(self.api.spreadsheets().get(spreadsheetId=book,
            fields="spreadsheetId,properties(title),sheets(properties(sheetId,title,gridProperties(rowCount,columnCount)))"))

    def values(self, book, rng, render="FORMATTED_VALUE"):
        return self.call(self.api.spreadsheets().values().get(spreadsheetId=book, range=rng,
            valueRenderOption=render, dateTimeRenderOption="SERIAL_NUMBER")).get("values", [])

    def ranges(self, book, ranges, render="FORMATTED_VALUE", dimension="ROWS"):
        return self.call(self.api.spreadsheets().values().batchGet(spreadsheetId=book,
            ranges=ranges, majorDimension=dimension, valueRenderOption=render,
            dateTimeRenderOption="SERIAL_NUMBER")).get("valueRanges", [])

    def grid(self, book, ranges, *, entered):
        values = 'effectiveValue,userEnteredValue' if entered else 'effectiveValue'
        response = self.call(self.api.spreadsheets().get(
            spreadsheetId=book, ranges=ranges,
            fields=f'sheets(properties(sheetId),data(startRow,startColumn,rowData(values({values}))))'))
        return response.get('sheets', [])


def quote(s):
    return "'" + s.replace("'", "''") + "'"


def col(n):
    result = ""
    while n:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result


def read_product_mapping(client, link):
    """Resolve an explicit Google Sheets tab and fetch its mapping via reads only."""
    url = urlparse(link.strip())
    match = re.fullmatch(r"/spreadsheets/d/([A-Za-z0-9_-]{20,})(?:/[^/]*)?/?", url.path)
    if url.scheme != "https" or url.netloc != "docs.google.com" or not match:
        raise StopRun("Masukkan link Google Sheets yang valid untuk List Produk.")
    book = match.group(1)
    gid = (parse_qs(url.fragment).get("gid") or parse_qs(url.query).get("gid") or [None])[0]
    if gid is not None and not gid.isdecimal():
        raise StopRun("GID tab List Produk harus berupa angka.")
    meta = client.meta(book)
    if meta.get("spreadsheetId") != book:
        raise StopRun("Google Sheets mengembalikan ID workbook mapping yang berbeda.")
    tabs = [entry["properties"] for entry in meta.get("sheets", [])]
    tab = next((t for t in tabs if (str(t["sheetId"]) == gid if gid is not None
                                    else t["title"] == "List Produk")), None)
    if tab is None:
        raise StopRun("Tab mapping tidak ditemukan. Gunakan link tab List Produk.")
    title = tab["title"]
    grid = tab["gridProperties"]
    header_rows = client.values(book, f"{quote(title)}!A1:{col(grid['columnCount'])}1",
                                "UNFORMATTED_VALUE")
    header = header_rows[0] if header_rows else []
    pid_i, brand_i, name_columns = mapping_columns(header)
    if not name_columns:
        raise StopRun("List Produk harus punya kolom PID Name, Product Name, atau SKU Name.")
    if grid["rowCount"] < 2:
        raise StopRun("List Produk kosong.")
    selected = (pid_i, brand_i, *name_columns)
    ranges = [f"{quote(title)}!{col(i + 1)}2:{col(i + 1)}{grid['rowCount']}"
              for i in selected]
    columns = client.ranges(book, ranges, "UNFORMATTED_VALUE", "COLUMNS")
    if len(columns) != len(ranges):
        raise StopRun("Pembacaan kolom List Produk tidak lengkap.")
    rows = zip_longest(*[(entry.get("values") or [[]])[0] for entry in columns], fillvalue="")
    return parse_product_mapping([header[i] for i in selected], rows,
                                 {"sheet_title": title, "header_row": 1})


def target_check(client, cfg):
    if not cfg["pc"]:
        raise StopRun("Target kosong.")
    check_target_allowed(cfg["pc"], allow_november_original=cfg.get("allow_november_original"))
    meta = client.meta(cfg["pc"])
    if meta.get("spreadsheetId") != cfg["pc"]:
        raise StopRun("Google Sheets mengembalikan ID workbook yang berbeda.")
    title = meta["properties"]["title"]
    tabs = {x["properties"]["title"] for x in meta.get("sheets", [])}
    if not set(TABS.values()).issubset(tabs):
        raise StopRun(f"Target {title!r} missing tabs: {set(TABS.values()) - tabs}")
    # Deliberately NO target-title checks: the trial copy may have any name.
    return meta


def schema(client, book, kind, meta=None):
    tab = TABS[kind]
    if meta is None:
        meta = client.meta(book)
    found = next((s["properties"] for s in meta["sheets"] if s["properties"]["title"] == tab), None)
    if not found:
        raise StopRun(f"Missing tab {tab}")
    cols = found.get("gridProperties", {}).get("columnCount", 0)
    rows = found.get("gridProperties", {}).get("rowCount", 0)
    template = client.values(book, f"{quote(tab)}!A1:{col(max(cols, 1))}1", "FORMULA")
    header = template[0] if template else []
    idx = {}
    for i, h in enumerate(header):
        if not str(h).strip():
            continue
        k = norm(h)
        if k in idx:
            raise StopRun(f"Duplicate target header in {tab}: {h!r}")
        idx[k] = i
    required = (*INPUTS[kind], *FORMULA_FIELDS[kind])
    if kind == "SV":
        required += ("Type Content",)
    missing = [h for h in required if norm(h) not in idx]
    if missing:
        raise StopRun(f"{tab}: missing target headers {missing}")
    # Guard template identity, not just input columns.
    if kind == "SV":
        expected = ["Product ID", "Product name", "Brand"]
        if [idx[norm(n)] for n in expected] != sorted(idx[norm(n)] for n in expected):
            raise StopRun("Raw SV columns Product ID -> Product name -> Brand are not in expected order")
    return {"tab": tab, "idx": idx, "rows": rows, "headers": header, "sheet_id": found["sheetId"]}


def existing_keys(client, book, sc, kind, possible_new_keys=None):
    """Read existing IDs as TEXT, never via Google formatted display.

    For older target rows that were already stored as large NUMERIC cells, the
    exact TikTok ID may be irreversibly rounded. Preserve them untouched with a
    synthetic internal key, and BLOCK only if a new exact ID could match one.
    This permits normal append when old unrepairable rows are unrelated, while
    preventing accidental duplication of a possibly matching content ID.
    """
    k1 = "Brand"
    k2 = "Video ID" if kind == "SV" else "Livestream room ID"
    cols = [col(sc["idx"][norm(k)] + 1) for k in (k1, k2)]
    result = defaultdict(list)
    count_legacy = 0
    examples = []
    new_by_brand = defaultdict(list)
    numeric_candidates = {}
    for brand, identifier in (possible_new_keys or ()):
        new_by_brand[brand].append(identifier)
    for first in range(2, sc["rows"] + 1, SCAN_ROWS):
        last = min(sc["rows"], first + SCAN_ROWS - 1)
        # Key fix: UNFORMATTED_VALUE yields original TEXT when a cell is text.
        # FORMATTED_VALUE can turn even valid IDs into '7.69714E+18'.
        vr = client.ranges(book, [f"{quote(sc['tab'])}!{c}{first}:{c}{last}" for c in cols],
                           "UNFORMATTED_VALUE", "COLUMNS")
        if len(vr) != len(cols):
            raise StopRun(f"{kind}: incomplete key scan at row {first}")
        pair = [(x.get("values") or [[]])[0] for x in vr]
        for offset in range(max((len(x) for x in pair), default=0)):
            brand = clean_field("Brand", pair[0][offset] if offset < len(pair[0]) else "")
            raw_id = pair[1][offset] if offset < len(pair[1]) else ""
            if not brand or raw_id in (None, ""):
                continue
            row_number = first + offset
            try:
                rid = exact_id(raw_id)
            except StopRun:
                # Do not fabricate an ID from scientific notation or a rounded
                # float. Compare against this run's new exact IDs conservatively.
                candidates = new_by_brand.get(brand, ())
                if type(raw_id) in (int, float):
                    if brand not in numeric_candidates:
                        by_number = defaultdict(list)
                        for candidate in candidates:
                            by_number[float(candidate)].append(candidate)
                        numeric_candidates[brand] = by_number
                    potential = numeric_candidates[brand].get(float(raw_id), ())
                else:
                    potential = [candidate for candidate in candidates
                                 if _maybe_same_legacy_id(raw_id, candidate)]
                if potential:
                    raise StopRun(
                        f"AMBIGUOUS_EXISTING_ID at {sc['tab']}!{cols[1]}{row_number}: "
                        f"stored={raw_id!r}; could be one of {potential[:3]}. "
                        "Cannot deduplicate safely. Restore this cell's EXACT ID "
                        "as plain text in your EXPERIMENT COPY, using the original source. "
                        "No Google write was attempted."
                    )
                # Keep the old row occupied for append placement checks. This
                # token NEVER matches a legitimate decimal content ID.
                rid = f"__LEGACY_NUMERIC_ID_ROW_{row_number}__"
                count_legacy += 1
                if len(examples) < 5:
                    examples.append(f"{sc['tab']}!{cols[1]}{row_number}={raw_id!r}")
            result[(brand, rid)].append(row_number)
    if count_legacy:
        print(f"  WARNING {kind}: {count_legacy} pre-existing IDs in the COPY were stored "
              "as numeric/scientific values; unrelated legacy rows were left "
              f"untouched. Examples: {', '.join(examples)}", flush=True)
    return result


def row_spans(row_numbers, *, max_gap=0):
    """Merge ordered row numbers into inclusive ranges."""
    spans = []
    for row in row_numbers:
        if spans and row <= spans[-1][1] + 1 + max_gap:
            spans[-1][1] = row
        else:
            spans.append([row, row])
    return spans


def input_grid(client, book, sc, kind, row_numbers, *, entered, include_formulas=False):
    """Fetch input values and formula metadata together, in bounded ranges."""
    fields = (*INPUTS[kind], *FORMULA_FIELDS[kind]) if include_formulas else INPUTS[kind]
    columns = {field: sc["idx"][norm(field)] for field in fields}
    left, right = min(columns.values()), max(columns.values()) + 1
    spans = row_spans(row_numbers, max_gap=1)
    wanted = set(row_numbers)
    tab = quote(sc['tab'])
    result = {}
    # Keep the GET URL short when row ranges are scattered.
    for i in range(0, len(spans), 35):
        subset = spans[i:i + 35]
        pending = {first - 1: last for first, last in subset}
        ranges = [f'{tab}!{col(left + 1)}{first}:{col(right)}{last}' for first, last in subset]
        for sheet in client.grid(book, ranges, entered=entered):
            if sheet.get("properties", {}).get("sheetId") != sc["sheet_id"]:
                raise StopRun(f"{kind}: unexpected sheet during row check")
            for block in sheet.get("data", []):
                first = block.get("startRow", 0)
                last = pending.pop(first, None)
                if last is None or block.get("startColumn", 0) != left:
                    raise StopRun(f"{kind}: unexpected grid range during row check")
                rows = block.get("rowData", [])
                if len(rows) > last - first:
                    raise StopRun(f"{kind}: grid response exceeded requested rows")
                for number, row in enumerate(rows, first + 1):
                    if number in wanted:
                        cells = row.get("values", [])
                        result[number] = {field: cells[ix - left] for field, ix in columns.items()
                                          if ix - left < len(cells) and cells[ix - left]}
        if pending:
            raise StopRun(f"{kind}: incomplete grid response during row check")
    return result


def last_used_input_row(client, book, sc, kind, *, min_row=2):
    """Search only pipeline input columns below the last known occupied key."""
    if not 2 <= min_row <= sc['rows'] + 1:
        raise StopRun(f'{kind}: invalid last-used scan start {min_row}')
    if min_row > sc['rows']:
        return min_row - 1
    fields = (*INPUTS[kind], *FORMULA_FIELDS[kind])
    letters = [col(sc['idx'][norm(f)] + 1) for f in fields]
    tab = quote(sc['tab'])
    for end in range(sc['rows'], min_row - 1, -SCAN_ROWS):
        first = max(min_row, end - SCAN_ROWS + 1)
        ranges = [f'{tab}!{c}{first}:{c}{end}' for c in letters]
        result = client.ranges(book, ranges, 'FORMULA', 'COLUMNS')
        if len(result) != len(ranges):
            raise StopRun(f'{kind}: incomplete append tail scan at row {first}')
        last = min_row - 1
        for field, item in zip(fields, result):
            cells = (item.get('values') or [[]])[0]
            for offset in range(len(cells) - 1, max(-1, last - first), -1):
                value = cells[offset]
                if field in MANAGED_FORMULAS[kind] and str(value).startswith("="):
                    continue
                if value is not None and str(value).strip():
                    last = first + offset
                    break
        if last >= first:
            return last
    return min_row - 1


def append_values(sc, kind, row_number, record):
    values = {field: value_for_sheet(field, record[field])
              for field in INPUTS[kind] if field in record.get("_present", ()) and record.get(field, "") != ""}

    def letter(field):
        return col(sc["idx"][norm(field)] + 1)

    date_field = "Post date" if kind == "SV" else "Live Date"
    date_cell = f"{letter(date_field)}{row_number}"
    values["Date"] = {"formulaValue": f'=TEXT({date_cell},"dd")'}
    if kind == "SV":
        content = f"{letter('Type Content')}{row_number}"
        values["Type Video"] = {
            "formulaValue": f'=IF(REGEXMATCH(LOWER(TO_TEXT({content})),"- kompetisi$"),'
                            f'"Kompetisi",IF(REGEXMATCH(LOWER(TO_TEXT({content})),"- event$"),'
                            '"Event","Mandatory"))'}
    if kind == "SV" and record["Source"] == "TAP":
        creator, video = (f"{letter(f)}{row_number}" for f in ("Creator name", "Video ID"))
        values["Video link"] = {
            "formulaValue": f'=IF(OR({creator}="",{video}=""),"","https://www.tiktok.com/@"&{creator}&"/video/"&{video})'}
    if kind == "LS" and record["Source"] == "SC" and record.get("Live Start") and record.get("Live End"):
        start, end = (f"{letter(f)}{row_number}" for f in ("Live Start", "Live End"))
        minutes = f"ROUND(ABS({start}-{end})*1440,0)"
        values["Duration"] = {"formulaValue": f'=INT({minutes}/60)&"h"&TEXT(MOD({minutes},60),"00")&"min"'}
    return values


def check_region(client, book, sc, kind, start_row, expected):
    """Permit template formulas in unused rows; verify committed rows exactly."""
    count = len(expected)
    if start_row < 2 or start_row + count - 1 > sc["rows"]:
        raise StopRun(f"{sc['tab']}: not enough allocated rows for {start_row}..{start_row + count - 1}")
    positions = range(start_row, start_row + count)
    grid = input_grid(client, book, sc, kind, positions, entered=True, include_formulas=True)
    all_match, has_literal = True, False
    for row_number, record in zip(positions, expected):
        cells = grid.get(row_number, {})
        target = append_values(sc, kind, row_number, record)
        for field in cells.keys() | target.keys():
            cell = cells.get(field, {})
            entered = cell.get("userEnteredValue", {})
            wanted = target.get(field, {})
            if "formulaValue" in entered:
                if field not in MANAGED_FORMULAS[kind]:
                    return "MISMATCH"
                if wanted and entered != wanted:
                    all_match = False
                continue
            if entered:
                has_literal = True
                if entered != wanted:
                    return "MISMATCH"
            elif wanted:
                if cell.get("effectiveValue"):
                    return "MISMATCH"
                all_match = False
    if all_match:
        return "MATCH"
    return "MISMATCH" if has_literal else "EMPTY"


def value_for_sheet(field, raw):
    if raw == "":
        return None
    if field in DATE_FIELDS:
        serial = (date.fromisoformat(raw) - date(1899, 12, 30)).days
        return {"numberValue": float(serial)}
    if field in TIME_FIELDS:
        try:
            dt = datetime.fromisoformat(raw)
        except ValueError:
            return {"stringValue": raw}
        seconds = dt.hour * 3600 + dt.minute * 60 + dt.second
        return {"numberValue": (dt.date() - date(1899, 12, 30)).days + seconds / 86400.0}
    if field in NUMBERS:
        v = Decimal(raw)
        if abs(v) > Decimal("9007199254740991"):
            raise StopRun(f"Numeric metric too large for Sheets precision: {field}={raw}")
        return {"numberValue": float(v)}
    return {"stringValue": str(raw)}  # never interpret creator/title as formulas


def atomic_append_requests(sc, kind, first, records):
    """Write only staged values and intended formulas in one atomic batch."""
    by_column = defaultdict(dict)
    for number, record in enumerate(records, first):
        for field, value in append_values(sc, kind, number, record).items():
            by_column[sc["idx"][norm(field)]][number] = value
    requests = column_write_requests(sc, by_column)
    for field in INPUTS[kind]:
        if field not in DATE_FIELDS | TIME_FIELDS:
            continue
        column = sc["idx"][norm(field)]
        date_only = field in DATE_FIELDS
        for start, end in row_spans(sorted(by_column[column])):
            requests.append({"repeatCell": {
                "range": {"sheetId": sc["sheet_id"], "startRowIndex": start - 1,
                          "endRowIndex": end, "startColumnIndex": column, "endColumnIndex": column + 1},
                "cell": {"userEnteredFormat": {"numberFormat": {
                    "type": "DATE" if date_only else "DATE_TIME",
                    "pattern": "yyyy-mm-dd" if date_only else "yyyy-mm-dd hh:mm:ss"}}},
                "fields": "userEnteredFormat.numberFormat"}})
    return requests


def load_engine(name, filename):
    path = ROOT / filename
    if not path.is_file():
        raise StopRun(f"Missing calculation backend: {filename}")
    spec = importlib.util.spec_from_file_location(name, path)
    mod = importlib.util.module_from_spec(spec)
    sys.modules[name] = mod
    orig = sys.argv[:]
    try:
        sys.argv = [str(path)]  # prevent backend --write from being inherited
        spec.loader.exec_module(mod)
    finally:
        sys.argv = orig
    if hasattr(mod, "WRITE"):
        mod.WRITE = False
    if hasattr(mod, "REPLACE"):
        mod.REPLACE = False
    return mod


def engines():
    if str(ROOT) not in sys.path:
        sys.path.insert(0, str(ROOT))
    result = {}
    result["tap_pc_core"] = load_engine("tap_pc_core", BACKENDS["tap_pc_core"])
    for name in ("tap_sv", "tap_ls", "sc_sv", "sc_ls"):
        result[name] = load_engine("simple_" + name, BACKENDS[name])
    return result


def process_tap(kind, eng, cfg):
    """Run the exact TAP core aggregation on local source records."""
    core = eng['tap_pc_core']
    mod = eng['tap_sv' if kind == 'SV' else 'tap_ls']
    label = 'TAP_' + kind
    required = [mod.COL_ID if kind == 'SV' else mod.COL_ROOM,
                mod.COL_DATE, mod.COL_POST_TIME if kind == 'SV' else mod.COL_TIME,
                *mod.SUM_COLS]
    header, data, _meta = read_local_file(label, cfg['files'][label], required, cfg['cache'])
    # Allow missing Product name when the export genuinely lacks that column:
    # only the OPTIONAL local PID->Brand mapping can classify the otherwise blank rows.
    if not any(norm(x) == norm('Product name') for x in header):
        print(f'  WARNING {label}: no Product name column; PID brand mapping is needed', flush=True)
        header = list(header) + ['Product name']
        data = [list(r) + [''] for r in data]
    mod.validate_required_columns(header)
    typed, _, _types, type_errors = core.coerce_rows_by_schema(
        header, data, column_types=mod.COLUMN_TYPES, strict=False, skip_invalid_rows=True)
    fallback, _, _mapping_stats = mapping_local(cfg)
    bh, br, _brandstat, _missing, _bundle = core.add_brand_from_product_name(
        header, typed, fallback_mapping=fallback, product_name_header=mod.COL_PRODUCT_NAME,
        product_id_header=mod.COL_PRODUCT_ID, brand_header=mod.COL_BRAND)
    validator = (mod.make_sv_row_validator(bh) if kind == 'SV'
                 else mod.make_live_row_validator(bh, allow_missing_end=True))
    groups, accepted, skipped, invalid = core.aggregate_brand_content(
        bh, br, content_id_header=mod.COL_ID if kind == 'SV' else mod.COL_ROOM,
        date_header=mod.COL_DATE, sum_columns=mod.SUM_COLS, brand_header=mod.COL_BRAND,
        row_validator=validator)
    if kind == 'SV':
        mod.add_post_date_column(groups, bh)
    else:
        mod.add_live_time_columns(groups, bh, accepted)
    base_header = bh + mod.EXTRA_HEADERS
    out_h = base_header + [mod.JENIS_HEADER]
    filtered, _ = core.filter_groups_by_date_ranges(
        base_header, groups, mod.FILTER_DATE_COLUMN, [(cfg['start'], cfg['end'])])
    range_rows = core.build_combined_from_groups(filtered, base_header, mod.GMV_COL)
    checks = core.reconcile_groups_vs_gabungan(base_header, filtered, range_rows, mod.SUM_COLS)
    if not all(x['ok'] for x in checks):
        raise StopRun(f'TAP {kind}: metric reconciliation failed: {checks}')
    # TAP sums use zero as an accumulator. Keep unavailable raw inputs from
    # becoming a zero or a partial metric in the upload.
    id_col = bh.index(mod.COL_ID if kind == "SV" else mod.COL_ROOM)
    brand_col = bh.index(mod.COL_BRAND)
    metric_cols = [bh.index(field) for field in mod.SUM_COLS]
    incomplete = defaultdict(set)
    for row in accepted:
        content_key = (row[brand_col], core.normalize_id(row[id_col]))
        if content_key in filtered:
            incomplete_fields = [i for i in metric_cols if row[i] in ("", None)]
            if incomplete_fields:
                incomplete[content_key].update(incomplete_fields)
    for row in range_rows:
        for i in incomplete.get((row[brand_col], core.normalize_id(row[id_col])), ()):
            row[i] = ""
    return (out_h, range_rows), {
        'raw': len(data), 'type_errors': len(type_errors), 'accepted': len(accepted),
        'invalid': len(invalid), 'skipped': skipped, 'range': len(range_rows)}


def process_sc(kind, eng, cfg, product_mapping=None):
    """Run unchanged SC calculation with a local-only source reader; never call Sheets for sources."""
    mod = eng['sc_sv' if kind == 'SV' else 'sc_ls']
    mod.CREDENTIALS_FILE = Path(__file__)  # backend validates existence; offline stub never reads it
    mod.WRITE = mod.REPLACE = False
    # Backend standalone output ID is deliberately disabled in bundled files.
    # Only the in-memory OA tables are captured; this dummy ID is NEVER used in an API call.
    mod.OUTPUT_SPREADSHEET_ID = 'LOCAL_NO_WRITE_ONLY'
    mod.WRITE_BRAND_TABS = mod.WRITE_BRAND_TABS_EFFECTIVE = False
    brands, _, stats = product_mapping if product_mapping is not None else mapping_local(cfg)
    mod.USE_PRODUCT_ID_MAPPING_FALLBACK = bool(brands)
    mod.MAPPING_SPREADSHEET_ID, mod.MAPPING_SHEET_GID = 'LOCAL_PRODUCT_MAPPING', 0
    mod.METRIC_START_DATE = mod.RANGE_START_DATE = cfg['start']
    mod.METRIC_END_DATE = mod.RANGE_END_DATE = cfg['end']
    if kind == 'SV':
        mod.SC_SPREADSHEET_ID, mod.SC_SHEET_NAME, mod.SC_SHEET_GID = 'LOCAL_SC_SV', 'LOCAL', None
        mod.TAP_SPREADSHEET_ID, mod.TAP_SHEET_NAME, mod.TAP_SHEET_GID = 'LOCAL_TAP_SV', 'LOCAL', None
    else:
        mod.SC_LS_SPREADSHEET_ID, mod.SC_LS_SHEET_NAME, mod.SC_LS_SHEET_GID = 'LOCAL_SC_LS', 'LOCAL', None
        mod.TAP_LS_SPREADSHEET_ID, mod.TAP_LS_SHEET_NAME, mod.TAP_LS_SHEET_GID = 'LOCAL_TAP_LS', 'LOCAL', None
    files = cfg['files']
    table_specs = {f'LOCAL_{name}': (name, files[name]) for name in ('TAP_' + kind, 'SC_' + kind)}
    mod.AFF_SOURCES = []
    for index, item in enumerate(files['AFF'], 1):
        ident = f'LOCAL_AFF_{index}'
        table_specs[ident] = ('AFF', item)
        mod.AFF_SOURCES.append({'spreadsheet_id': ident, 'sheet_gid': None,
                                'sheet_name': 'LOCAL', 'label': f'AFF {index}'})
    captured = {}

    def detect_local(_unused_service, *, spreadsheet_id, sheet_gid, sheet_name, required_headers, label):
        if spreadsheet_id not in table_specs:
            raise StopRun(f'SC {kind}: unexpected/remote source request: {spreadsheet_id}')
        table_label, entry = table_specs[spreadsheet_id]
        header, rows, meta = read_local_file(table_label, entry, required_headers, cfg['cache'])
        copied = dict(meta)
        copied['spreadsheet_id'] = spreadsheet_id
        copied['sheet_title'] = meta['sheet_title']
        return header, [dict(zip(header, row)) for row in rows], copied

    old = (mod.get_sheets_service, mod.detect_and_read_source,
           mod.load_product_brand_mapping, mod.validate_output_capacity)
    mod.get_sheets_service = lambda: object()  # offline sentinel: no Google client created
    mod.detect_and_read_source = detect_local
    mod.load_product_brand_mapping = lambda _svc, **_kw: (brands, stats)
    def capture_output(_svc, tables):
        name = mod.BRAND_UNIQUE_RANGE_SHEET
        captured[name] = tables[name]

    mod.validate_output_capacity = capture_output
    try:
        mod.main()
    except SystemExit as exc:
        raise StopRun(str(exc)) from exc
    finally:
        (mod.get_sheets_service, mod.detect_and_read_source,
         mod.load_product_brand_mapping, mod.validate_output_capacity) = old
    if not captured or mod.BRAND_UNIQUE_RANGE_SHEET not in captured:
        raise StopRun(f'SC {kind}: local OA output capture failed')
    return captured, mod.BRAND_UNIQUE_RANGE_SHEET

# ============================================================================
# PLAN / QC / WRITE / RESUME  (single runner, no external safety/write modules)
# ============================================================================
def assemble(cfg, engines_, progress):
    """Uses unchanged TAP/SC engines. Never publishes OA to Google Sheets."""
    staged = {'SV': [], 'LS': []}
    summary = {}
    product_mapping = cfg.get("sc_product_mapping") or mapping_local(cfg)
    _, product_names, _ = product_mapping
    for kind in ('SV', 'LS'):
        progress(f"{kind} · Membaca dan mengagregasi TAP")
        tap_range, counts = process_tap(kind, engines_, cfg)
        h, rows = tap_range
        tap_dicts = as_dict(h, rows)
        tap_keys = {(clean_field('Brand', r.get('Brand')), exact_id(r.get('Video ID' if kind == 'SV' else 'Livestream room ID')))
                    for r in tap_dicts}
        progress(f"{kind} · Mapping produk dan agregasi SC")
        sc_tables, sc_tab = process_sc(kind, engines_, cfg, product_mapping)
        sh, *sr = sc_tables[sc_tab]
        sc_dicts = []
        sc_in_tap = 0
        for r in as_dict(sh, sr):
            identifier = exact_id(r.get('Video ID' if kind == 'SV' else 'LIVE ID'))
            k = (clean_field('Brand', r.get('Brand')), identifier)
            if k in tap_keys:
                sc_in_tap += 1
            else:
                # Keep the legacy brand and GMV allocation. A catalog name
                # fills missing output metadata for the selected product only.
                name = str(r.get("Product name") or "").strip()
                if not name or name.casefold() == k[0].casefold():
                    pid = exact_id(r.get("Product ID"))
                    r["Product name"] = product_names.get(pid, name)
                sc_dicts.append(r)
        staged[kind] = [stage(kind, 'TAP', r) for r in tap_dicts] + [stage(kind, 'SC', r) for r in sc_dicts]
        summary[kind] = {'tap': counts, 'sc_brand_unique_not_in_tap': len(sc_dicts),
                         'sc_already_in_tap': sc_in_tap, 'staging': len(staged[kind])}
        print(f'[{kind}] Staging TAP={len(tap_dicts):,} SC={len(sc_dicts):,} (overlaps skipped={sc_in_tap:,})', flush=True)
    return staged, summary


def classify_candidates(client, cfg, staged, schemas):
    """Append only new keys and report invalid or conflicting source rows."""
    additions, results, key_indexes = {}, {}, {}
    for kind in ('SV', 'LS'):
        valid, counts, warnings, quarantined = {}, Counter(), [], []
        batch_conflicts = set()
        rows = staged[kind]
        for row in rows:
            status = is_valid(kind, row, cfg)
            if status != 'OK':
                counts[status] += 1
                # Outside-date rows are filtered, not data-quality failures.
                if status != 'OUTSIDE_CUTOFF':
                    quarantined.append({'reason': status, 'Brand': row.get('Brand',''),
                                        'Content ID': key(kind,row)[1], 'Source': row.get('Source',''),
                                        'detail': 'Incomplete or invalid source row'})
                continue
            k = key(kind, row)
            if k in valid:
                if canonical(kind, row) != canonical(kind, valid[k]):
                    batch_conflicts.add(k)
                    counts['CONFLICT_IN_BATCH'] += 1
                else:
                    counts['DUPLICATE_IN_BATCH'] += 1
                continue
            valid[k] = row
        for k in batch_conflicts:
            valid.pop(k, None)
            quarantined.append({'reason': 'CONFLICT_IN_BATCH', 'Brand': k[0], 'Content ID': k[1],
                                'Source': '', 'detail': 'Conflicting staged rows; skipped'})
        existing = existing_keys(client, cfg['pc'], schemas[kind], kind, set(valid))
        key_indexes[kind] = existing
        historical_duplicates = {k: positions for k, positions in existing.items() if len(positions) > 1}
        if historical_duplicates:
            warnings.append(f'EXISTING_DUPLICATE_KEYS={len(historical_duplicates)} (not auto-deleted)')
        fresh = []
        for k, row in valid.items():
            if k in existing:
                counts['SKIP_EXISTING'] += 1
                continue
            if kind == 'SV' and (not row.get('Product name') or
                    row['Product name'].casefold() == row['Brand'].casefold()):
                counts['SKIP_INVALID_PRODUCT_NAME'] += 1
                quarantined.append({'reason': 'INVALID_PRODUCT_NAME', 'Brand': k[0], 'Content ID': k[1],
                                    'Source': row.get('Source', ''), 'detail': 'Missing Product name or equals Brand'})
                continue
            fresh.append(row)
            counts['APPEND_NEW'] += 1
        invalid_data = sum(counts[x] for x in ('MISSING_KEY','INVALID_BRAND','MISSING_CREATOR','MISSING_DATE'))
        if invalid_data:
            warnings.append(f'SOURCE_INVALID_ROWS={invalid_data}/{len(rows)} (quarantined)')
        if counts['OUTSIDE_CUTOFF']:
            warnings.append(f'OUTSIDE_CUTOFF={counts["OUTSIDE_CUTOFF"]} (excluded, not invalid)')
        if batch_conflicts:
            warnings.append(f'CONFLICT_IN_BATCH_KEYS={len(batch_conflicts)} (quarantined)')
        additions[kind] = fresh
        results[kind] = {'warnings': warnings, 'quarantined': quarantined,
                         'new_rows': len(fresh), 'skipped_existing': counts['SKIP_EXISTING']}
    return additions, results, key_indexes


def code_digest(version=PLAN_VERSION):
    pipeline = Path(__file__).name
    files = {filename: hashlib.sha256((ROOT / filename).read_bytes()).hexdigest()
             for filename in (pipeline, "report_formulas.json", *BACKENDS.values())}
    if version in PRIOR_PIPELINE_SHA256:
        files[pipeline] = PRIOR_PIPELINE_SHA256[version]
    if version in (5, 6):
        files[BACKENDS["tap_ls"]] = PRIOR_TAP_LS_SHA256
    return digest(files)


def save_plan(path, plan):
    plain = dict(plan)
    plain.pop('integrity_sha256', None)
    plan['integrity_sha256'] = digest(plain)
    write_json(path, plan)


def read_plan(path):
    if not path.exists():
        raise StopRun('No frozen interrupted write plan. Use --write first; --resume is only for interrupted writes.')
    plan = json.loads(path.read_text(encoding='utf-8'))
    expected = plan.get('integrity_sha256')
    if not expected or digest({k: v for k, v in plan.items() if k != 'integrity_sha256'}) != expected:
        raise StopRun('Saved run plan checksum mismatched; refuse resume')
    return plan


def chunk_groups(plan, schemas):
    legacy = plan["version"] == LEGACY_PLAN_VERSION
    for kind in ("SV", "LS"):
        info = plan["kinds"][kind]
        rows = info["rows"]
        first = info["first_row"]
        schema = schemas[kind]
        if legacy:
            # Resume must reproduce the exact boundaries of the old upload.
            for offset in range(0, len(rows), LEGACY_CHUNK_ROWS):
                yield from legacy_bounded_chunks(
                    kind, first + offset, rows[offset:offset + LEGACY_CHUNK_ROWS], schema)
            continue
        offset = 0
        while offset < len(rows):
            batch, requests = fit_append_batch(
                kind, first + offset, rows[offset:offset + CHUNK_ROWS], schema)
            yield "APPEND", kind, first + offset, batch, requests
            offset += len(batch)


def fit_append_batch(kind, first, rows, schema):
    count = len(rows)
    while True:
        batch = rows[:count]
        requests = atomic_append_requests(schema, kind, first, batch)
        # json.dumps uses the same default ASCII escaping as the Sheets client.
        size = len(json.dumps({"requests": requests}))
        if size <= MAX_BATCH_BYTES:
            return batch, requests
        if count == 1:
            raise StopRun(f"{kind} row {first}: write payload exceeds {MAX_BATCH_BYTES:,} bytes")
        count = max(1, count * MAX_BATCH_BYTES // size)


def legacy_bounded_chunks(kind, first, rows, schema):
    requests = atomic_append_requests(schema, kind, first, rows)
    if len(json.dumps({"requests": requests})) <= MAX_BATCH_BYTES:
        yield "APPEND", kind, first, rows, requests
        return
    if len(rows) == 1:
        raise StopRun(f"{kind} row {first}: write payload exceeds {MAX_BATCH_BYTES:,} bytes")
    mid = len(rows) // 2
    yield from legacy_bounded_chunks(kind, first, rows[:mid], schema)
    yield from legacy_bounded_chunks(kind, first + mid, rows[mid:], schema)


def column_write_requests(sc, by_column):
    requests = []
    for column, cells in by_column.items():
        for first, last in row_spans(sorted(cells)):
            requests.append({'updateCells': {
                'start': {'sheetId': sc['sheet_id'], 'rowIndex': first - 1, 'columnIndex': column},
                'rows': [{'values': [{'userEnteredValue': cells[row]}]} for row in range(first, last + 1)],
                'fields': 'userEnteredValue'}})
    return requests


def verify_plan_keys(client, cfg, plan, schemas, *, committed=False):
    """Catch a frozen key unexpectedly moved, missing or duplicated."""
    for kind in ('SV', 'LS'):
        info = plan['kinds'][kind]
        candidate_keys = {key(kind, r) for r in info['rows']}
        if not candidate_keys:
            continue
        all_keys = existing_keys(client, cfg['pc'], schemas[kind], kind, candidate_keys)
        first = info['first_row']
        for j, r in enumerate(info['rows']):
            positions = all_keys.get(key(kind, r), [])
            if (committed or positions) and positions != [first + j]:
                raise StopRun(f'Unexpected or duplicated appended key in {kind}: {key(kind, r)} at {positions}')


def make_plan(client, cfg, schemas, additions, key_indexes):
    result = {'version': PLAN_VERSION, 'target_id': cfg['pc'], 'start': cfg['start'], 'end': cfg['end'],
              'script_sha256': code_digest(), 'created_at': datetime.now(timezone.utc).isoformat(),
              'state': 'PENDING', 'kinds': {}}
    for kind in ('SV', 'LS'):
        sc = schemas[kind]
        rows = additions[kind]
        existing = key_indexes[kind]
        last_key = max((r for rlist in existing.values() for r in rlist), default=1)
        used = last_used_input_row(client, cfg['pc'], sc, kind, min_row=last_key + 1) if rows else last_key
        # Orphan legacy rows may have content but no valid Brand+ID.
        # Append strictly AFTER the last occupied input row; never overwrite them.
        if used != last_key:
            print(f'  WARNING {kind}: last occupied input row={used}, last keyed row={last_key}; '
                  'preserving all legacy/orphan rows and appending after both', flush=True)
        first = used + 1
        required_rows = first + len(rows) - 1 if rows else sc['rows']
        result['kinds'][kind] = {'first_row': first, 'rows': rows,
                                'schema_hash': digest({'headers': sc['headers'], 'sheet_id': sc['sheet_id']}),
                                'expected_base_last_key': last_key, 'required_rows': required_rows}
    return result


def validate_frozen(cfg, plan, schemas):
    version = plan.get('version')
    if version != PLAN_VERSION and version not in PRIOR_PIPELINE_SHA256:
        raise StopRun('Versi rencana berbeda. Lanjutkan dengan versi app yang membuatnya.')
    if plan['target_id'] != cfg['pc'] or plan['start'] != cfg['start'] or plan['end'] != cfg['end']:
        raise StopRun('Resume target or dates changed vs frozen write plan')
    if plan['script_sha256'] != code_digest(version):
        raise StopRun('Runner script changed since original write. Restore exact code before --resume')
    for kind, sc in schemas.items():
        actual = digest({'headers': sc['headers'], 'sheet_id': sc['sheet_id']})
        if plan['kinds'][kind]['schema_hash'] != actual:
            raise StopRun(f'{kind} target schema changed; refuse resume')


def ensure_copy_row_capacity(reader, cfg, plan, schemas, writer):
    """Grow the approved target's RAW tabs to fit planned rows."""
    for kind in ('SV','LS'):
        sc = schemas[kind]
        needed = plan['kinds'][kind].get('required_rows', sc['rows'])
        if needed <= sc['rows']:
            continue
        growth = needed - sc['rows']
        # Recheck write authorization before expanding the target.
        check_write_armed(cfg)
        print(f'  Target capacity {kind}: allocate {growth} new blank rows', flush=True)
        writer.call(writer.api.spreadsheets().batchUpdate(
            spreadsheetId=cfg['pc'], body={'requests':[{'appendDimension': {
                'sheetId': sc['sheet_id'], 'dimension':'ROWS', 'length':growth}}]}),
            retry_write=False)
        refreshed = schema(reader, cfg['pc'], kind)
        if any(refreshed[k] != sc[k] for k in ("sheet_id", "headers")) or refreshed['rows'] < needed:
            raise StopRun(f'{kind} copy row expansion could not be verified')
        sc['rows'] = refreshed['rows']


def install_report_formulas(reader, writer, cfg, schemas, progress):
    """Seed empty report anchors on a compatible copy after the RAW append."""
    definitions = json.loads((ROOT / "report_formulas.json").read_text(encoding="utf-8"))
    meta = target_check(reader, cfg)
    tabs = {item["properties"]["title"]: item["properties"] for item in meta["sheets"]}
    required = ("CONFIG", "CALC SV", "CALC LS", "RINGKAS SV", "RINGKAS LS", "COUNTS")

    def matches(actual, expected):
        return all(norm(value) == norm(actual[i] if i < len(actual) else "")
                   for i, value in enumerate(expected))

    if not set(required).issubset(tabs) or any(
            not matches(schema["headers"], definitions[schema["tab"]]["headers"])
            for schema in schemas.values()):
        return

    names = list(required)
    if "CARA PAKAI" in tabs:
        names.append("CARA PAKAI")
    ranges = []
    for name in names:
        definition = definitions[name]
        rows = 7 if name == "CONFIG" else max(r for r, _, _ in definition["formulas"])
        columns = len(definition["headers"])
        grid = tabs[name]["gridProperties"]
        if rows > grid["rowCount"] or columns > grid["columnCount"]:
            return
        ranges.append(f"{quote(name)}!A1:{col(columns)}{rows}")
    progress("Memeriksa rumus rekap pada salinan")
    grids = reader.grid(cfg["pc"], ranges, entered=True)

    def index_grids(sheets):
        return {item["properties"]["sheetId"]: item["data"][0].get("rowData", [])
                for item in sheets if item.get("data")}

    def cell(rows, row, column):
        values = rows[row - 1].get("values", []) if row <= len(rows) else []
        return values[column] if column < len(values) else {}

    by_id = index_grids(grids)
    requests, written = [], []
    for name in names:
        sheet_id = tabs[name]["sheetId"]
        rows = by_id.get(sheet_id, [])
        header = [c.get("userEnteredValue", {}).get("stringValue", "")
                  for c in rows[0].get("values", [])] if rows else []
        definition = definitions[name]
        if not matches(header, definition["headers"]):
            if name == "CARA PAKAI":
                continue
            return
        if name == "CONFIG" and any(
                not cell(rows, row, column).get("effectiveValue")
                for row in range(2, 8) for column in (1, 2)):
            return
        by_column = defaultdict(dict)
        for row, column, formula in definition["formulas"]:
            current = cell(rows, row, column)
            if not current.get("userEnteredValue") and not current.get("effectiveValue"):
                value = {"formulaValue": formula}
                by_column[column][row] = value
                written.append((sheet_id, row, column, value))
        requests.extend(column_write_requests({"sheet_id": sheet_id}, by_column))
    if not requests:
        return
    check_write_armed(cfg)
    writer.call(writer.api.spreadsheets().batchUpdate(
        spreadsheetId=cfg["pc"], body={"requests": requests}), retry_write=False)
    verified = index_grids(reader.grid(cfg["pc"], ranges, entered=True))
    if any(cell(verified.get(sheet_id, []), row, column).get("userEnteredValue") != expected
           for sheet_id, row, column, expected in written):
        raise StopRun("Sebagian rumus rekap belum terverifikasi. Lanjutkan batch untuk memeriksa ulang.")
    print(f"  Rekap: {len(written)} rumus kosong diisi dan diverifikasi.", flush=True)


def perform_chunks(client, cfg, plan, path, creds, schemas, *, resume=False, progress=print):
    started = time.monotonic()
    reads = client.requests
    if not any(info['rows'] for info in plan['kinds'].values()):
        plan['state'] = 'COMPLETE'
        save_plan(path, plan)
        print('No new rows. No Google writes.')
        return
    progress('Memeriksa ulang key dan batch yang akan dikirim')
    verify_plan_keys(client, cfg, plan, schemas)
    check_write_armed(cfg)
    writer = Google(creds, readonly=False)
    ensure_copy_row_capacity(client, cfg, plan, schemas, writer)
    committed = set()
    batch_count = 0
    if resume:
        for action, kind, first, rows, _ in chunk_groups(plan, schemas):
            batch_count += 1
            state = check_region(client, cfg['pc'], schemas[kind], kind, first, rows)
            if state == 'MATCH':
                committed.add((action, kind, first))
            elif state != 'EMPTY':
                raise StopRun(f'Unsafe {action} {kind} row {first}: {state}; manual review required')
        print(f'Uploading {batch_count - len(committed):,}/{batch_count:,} batches '
              f'({len(committed):,} already committed) ...', flush=True)
    for number, (action, kind, first, rows, requests) in enumerate(chunk_groups(plan, schemas), 1):
        if (action, kind, first) in committed:
            continue
        sc = schemas[kind]
        # Check immediately before each atomic write. A failed write is never
        # retried blindly. Resume inspects its current state before continuing.
        now = check_region(client, cfg['pc'], sc, kind, first, rows)
        if now == 'MATCH':
            continue
        if now != 'EMPTY':
            raise StopRun(f'{action} {kind} row {first}: concurrent change detected ({now})')
        if not requests:
            raise StopRun('Internal error: empty write request set')
        batch_label = f"{number}/{batch_count}" if resume else str(number)
        progress(f'Batch {batch_label} · {kind} · {action} · {len(rows):,} baris')
        writer.call(writer.api.spreadsheets().batchUpdate(
            spreadsheetId=cfg['pc'], body={'requests': requests}), retry_write=False)
        after = check_region(client, cfg['pc'], sc, kind, first, rows)
        if after != 'MATCH':
            raise StopRun(f'WRITE VERIFY FAILED {action} {kind} row {first}: {after}. Use --resume for frozen plan.')
        print(f'  {kind}: appended {len(rows)} rows at {first} and verified', flush=True)
    verify_plan_keys(client, cfg, plan, schemas, committed=True)
    if cfg["pc"] != NOVEMBER_ORIGINAL_ID:
        install_report_formulas(client, writer, cfg, schemas, progress)
    plan['state'] = 'COMPLETE'
    plan['completed_at'] = datetime.now(timezone.utc).isoformat()
    save_plan(path, plan)
    print(f'SUCCESS: all rows verified in {time.monotonic() - started:.1f}s '
          f'({client.requests - reads:,} read, {writer.requests:,} write API attempts).', flush=True)


