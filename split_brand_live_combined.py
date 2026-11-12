"""Incremental monthly processing for TAP PC Livestream data."""

from __future__ import annotations

import re
import sys
from datetime import datetime
from pathlib import Path

from tap_pc_core import (
    LOAD_LOG_HEADERS,
    LOAD_LOG_SHEET,
    MAPPING_BRAND_HEADER,
    MAPPING_PRODUCT_ID_HEADER,
    MAPPING_SHEET_GID,
    MAPPING_SPREADSHEET_ID,
    MISSING_BRAND,
    OUTPUT_BRAND_HEADER,
    PC_BRANDS,
    PRODUCT_NAME_HEADER,
    SOURCE_PRODUCT_ID_HEADER,
    TYPE_CURRENCY,
    TYPE_DATE,
    TYPE_DATETIME,
    TYPE_INTEGER,
    TYPE_TEXT,
    add_brand_from_product_name,
    aggregate_brand_content,
    append_load_log,
    attach_master_metadata,
    build_brand_tabs,
    build_combined_from_groups,
    build_master_header,
    cell,
    coerce_rows_by_schema,
    cumulative_load_ranges,
    determine_load_action,
    filter_groups_by_date_ranges,
    find_header_index,
    get_sheets_service,
    latest_load_record,
    load_product_brand_mapping,
    metric_totals,
    normalize_id,
    parse_config_date,
    prepare_master_upsert,
    read_source,
    read_table_if_exists,
    rebuild_monthly_output,
    reconcile_groups_vs_gabungan,
    reconcile_sum_columns,
    replace_table,
    resolve_load_id,
    table_fingerprint,
    upsert_master_table,
    validate_brand_column,
    validate_header_layout,
    validate_target_capacity,
)


# =============================================================================
# Configuration
# =============================================================================

BASE_DIR = Path(__file__).resolve().parent
CREDENTIALS_FILE = BASE_DIR / "credentials.json"

PIPELINE = "LS"

# TRUE  = tulis/update tab per Brand (ALL + RANGE)
# FALSE = jangan tulis/update tab per Brand.
# GABUNGAN, persistent master, dan _LOAD_LOG tetap ditulis seperti biasa.
# Catatan: FALSE tidak menghapus tab Brand lama yang sudah ada.
WRITE_BRAND_TABS = False

# Toggle dimasukkan ke version identity supaya perubahan TRUE/FALSE pada
# Load ID yang sama tidak dianggap sebagai run identik / SKIP.
CODE_VERSION = (
    "5.4.0-brandtabs-on"
    if WRITE_BRAND_TABS
    else "5.4.0-brandtabs-off"
)

SOURCE_SPREADSHEET_ID = "1AapRa4hxn__NSTtcXQGD0qqE1KP8TuHxC1BJRXO-Bxw"
SOURCE_SHEET = "Custom report"
SOURCE_RANGE = "A1:AA"

# LS and SV should point to the same monthly output spreadsheet.
MONTHLY_OUTPUT_SPREADSHEET_ID = "ISI_BACKEND_DISABLED_NO_DIRECT_WRITE"
LOAD_ID = ""

RANGE_START_DATE = "2026-10-08"
RANGE_END_DATE = "2026-10-12"

USE_PRODUCT_ID_MAPPING_FALLBACK = True

MASTER_ALL_SHEET = "_MASTER_LS_ALL"

BRAND_ALL_SUFFIX = " - LS ALL"
BRAND_RANGE_SUFFIX = " - LS RANGE"

FINAL_ALL_SHEET = "GABUNGAN LS - ALL"
FINAL_RANGE_SHEET = "GABUNGAN LS - RANGE"

WRITE = "--write" in sys.argv
FORCE = "--force" in sys.argv


# =============================================================================
# Source columns and type policy
# =============================================================================

COL_PRODUCT_NAME = PRODUCT_NAME_HEADER
COL_PRODUCT_ID = SOURCE_PRODUCT_ID_HEADER
COL_BRAND = OUTPUT_BRAND_HEADER
COL_ROOM = "Livestream room ID"
COL_DATE = "Date"
COL_TIME = "LIVE time info"

COLUMN_TYPES = {
    "Date": TYPE_DATE,
    "Comparison date": TYPE_TEXT,
    "Campaign ID": TYPE_TEXT,
    "Campaign name": TYPE_TEXT,
    "Campaign duration": TYPE_TEXT,
    "Creator name": TYPE_TEXT,
    "Creator follower count": TYPE_INTEGER,
    "Product ID": TYPE_TEXT,
    "Product name": TYPE_TEXT,
    "brand": TYPE_TEXT,
    "subbrand": TYPE_TEXT,
    "Shop code": TYPE_TEXT,
    "Shop ID": TYPE_TEXT,
    "Shop name": TYPE_TEXT,
    "Livestream room ID": TYPE_TEXT,
    "Livestream name": TYPE_TEXT,
    "LIVE time info": TYPE_TEXT,
    "Level 1 category": TYPE_TEXT,
    "Level 2 category": TYPE_TEXT,
    "Creator LIVE-attributed GMV": TYPE_CURRENCY,
    "Creator LIVE-attributed orders": TYPE_INTEGER,
    "Affiliate LIVE orders": TYPE_INTEGER,
    "Duration": TYPE_TEXT,
    "LIVE views": TYPE_INTEGER,
    "LIVE likes": TYPE_INTEGER,
    "LIVE product RPM": TYPE_CURRENCY,
    "Estimated affiliate partner commission": TYPE_CURRENCY,
    "Actual affiliate partner commission": TYPE_CURRENCY,
    "Creator-attributed items sold": TYPE_INTEGER,
}

GMV_COL = "Creator LIVE-attributed GMV"

SUM_COLS = [
    GMV_COL,
    "Creator LIVE-attributed orders",
    "Affiliate LIVE orders",
    "Creator-attributed items sold",
]

EXTRA_HEADERS = [
    "Start Live",
    "End Live",
    "Date Live",
]

FILTER_DATE_COLUMN = "Date Live"
JENIS_HEADER = "Jenis"

LIVE_TIME_RE = re.compile(
    r"^\s*"
    r"(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})"
    r"(?:\s*-\s*"
    r"(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})?"
    r")?\s*$"
)


# =============================================================================
# Validation and derived fields
# =============================================================================

def validate_config():
    if str(MONTHLY_OUTPUT_SPREADSHEET_ID).startswith("ISI_"):
        raise SystemExit(
            "Isi MONTHLY_OUTPUT_SPREADSHEET_ID dengan ID spreadsheet output bulanan."
        )

    if SOURCE_SPREADSHEET_ID == MONTHLY_OUTPUT_SPREADSHEET_ID:
        raise SystemExit(
            "Source spreadsheet harus berbeda dari spreadsheet output bulanan."
        )

    range_start = parse_config_date(
        RANGE_START_DATE,
        "RANGE_START_DATE",
    )
    range_end = parse_config_date(
        RANGE_END_DATE,
        "RANGE_END_DATE",
    )

    if range_start > range_end:
        raise SystemExit(
            "RANGE_START_DATE tidak boleh lebih besar dari RANGE_END_DATE."
        )


def validate_required_columns(header):
    required = [
        COL_PRODUCT_NAME,
        COL_ROOM,
        COL_DATE,
        COL_TIME,
        *SUM_COLS,
    ]

    if USE_PRODUCT_ID_MAPPING_FALLBACK:
        required.append(COL_PRODUCT_ID)

    for name in required:
        find_header_index(header, name)


def make_live_row_validator(header, *, allow_missing_end=False):
    """Build a row validator with the LIVE time column index resolved once."""
    time_index = find_header_index(header, COL_TIME)

    def validate(row, row_number, _header):
        live_time = str(cell(row, time_index) or "").strip()

        match = LIVE_TIME_RE.fullmatch(live_time)
        if match and (match.group(2) or allow_missing_end):
            start, end = match.groups()
            try:
                datetime.fromisoformat(start)
                datetime.fromisoformat(end or start)
            except ValueError:
                pass
            else:
                return True, ""

        return False, f"LIVE time info tidak valid: {live_time!r}"

    return validate


def add_live_time_columns(groups, header, accepted_rows=()):
    """Derive Live Date from the start and keep an observed end from the same group."""
    time_index = find_header_index(header, COL_TIME)
    brand_index = find_header_index(header, COL_BRAND)
    room_index = find_header_index(header, COL_ROOM)
    observed_ends = {}
    for row in accepted_rows:
        match = LIVE_TIME_RE.fullmatch(str(cell(row, time_index) or "").strip())
        if match and match.group(2):
            brand = str(cell(row, brand_index) or "").strip().upper()
            room = normalize_id(cell(row, room_index))
            observed_ends.setdefault((brand, room, match.group(1)), match.group(2))

    invalid_count = 0
    for key, row in groups.items():
        live_time = str(cell(row, time_index) or "").strip()
        match = LIVE_TIME_RE.fullmatch(live_time)
        if match:
            start_live, end_live = match.groups()
            end_live = end_live or observed_ends.get((*key, start_live), "")
            row.extend([start_live, end_live, start_live[:10]])
        else:
            invalid_count += 1
            row.extend(["", "", ""])

    return invalid_count


# =============================================================================
# Diagnostics and output helpers
# =============================================================================

def print_type_diagnostics(type_summary, type_errors):
    print("\nData types:")

    for data_type, count in sorted(type_summary.items()):
        print(f"  {data_type:<12} {count:>3} kolom")

    if not type_errors:
        print("\nSource type validation: OK")
        return

    invalid_rows = sorted({item[0] for item in type_errors})

    print("\nInvalid source rows excluded during type conversion:")
    print(f"  Rows excluded     : {len(invalid_rows):,}")
    print(f"  Conversion errors : {len(type_errors):,}")

    for row_number, column, value, reason in type_errors[:20]:
        print(
            f"  Row {row_number}: {column} = {value!r} "
            f"({reason})"
        )

    if len(type_errors) > 20:
        print(
            f"  ... {len(type_errors) - 20:,} error lainnya tidak ditampilkan."
        )


def print_brand_diagnostics(status, missing_names, bundle_parent_counts):
    print("\nBrand derivation:")

    for key, count in sorted(status.items()):
        print(f"  {key:<30} {count:>10,}")

    if bundle_parent_counts:
        print("  Bundle parent by Brand:")
        for brand, count in sorted(bundle_parent_counts.items()):
            print(f"    {brand:<14} {count:>10,}")

    if missing_names:
        print(f"  Unique missing Product name: {len(missing_names):,}")
        print("  Top missing:")

        for name, count in missing_names.most_common(10):
            shown = name[:90] + ("..." if len(name) > 90 else "")
            print(f"    {count:>7,}x | {shown}")


def print_reconciliation(label, checks):
    print(f"\n{label}:")

    for check in checks:
        if "source" in check:
            print(
                f"  {check['column']}: source {check['source']:,} | "
                f"aggregate {check['aggregate']:,} | "
                f"{'OK' if check['ok'] else 'BEDA!'}"
            )
        else:
            print(
                f"  {check['column']}: groups {check['groups']:,} | "
                f"gabungan {check['gabungan']:,} | "
                f"{'OK' if check['ok'] else 'BEDA!'}"
            )

    if not all(check["ok"] for check in checks):
        raise ValueError(label + " ada yang BEDA.")


def ensure_monthly_brand_tabs(tabs, suffix):
    """Return all canonical monthly brand tabs, including empty tabs."""
    output = dict(tabs)

    for brand in sorted(set(PC_BRANDS) | {MISSING_BRAND}):
        output.setdefault(f"{brand}{suffix}", [])

    return dict(sorted(output.items()))


def build_load_log_row(
    load_id,
    source_fingerprint,
    action,
    source_rows,
    stored_all_rows,
    final_all_rows,
    final_range_rows,
    loaded_at,
):
    return [
        PIPELINE,
        load_id,
        SOURCE_SPREADSHEET_ID,
        SOURCE_SHEET,
        source_fingerprint,
        CODE_VERSION,
        action,
        source_rows,
        stored_all_rows,
        final_all_rows,
        final_range_rows,
        RANGE_START_DATE,
        RANGE_END_DATE,
        "SUCCESS",
        loaded_at,
    ]


# =============================================================================
# Main
# =============================================================================

def main():
    validate_config()

    service = get_sheets_service(
        CREDENTIALS_FILE,
        write=WRITE,
    )

    load_id = resolve_load_id(
        LOAD_ID,
        SOURCE_SPREADSHEET_ID,
    )

    print("=" * 78)
    print("TAP PC LIVESTREAM - INCREMENTAL MONTHLY PIPELINE")
    print("=" * 78)
    print(f'Mode       : {"WRITE" if WRITE else "DRY RUN"}')
    print(f"Pipeline   : {PIPELINE}")
    #print(f"Version    : {CODE_VERSION}")
    print(f"Load ID    : {load_id}")
    print(f"Brand tabs : {'ON' if WRITE_BRAND_TABS else 'OFF'}")
    print(f"Current range : {RANGE_START_DATE} -> {RANGE_END_DATE} (Date Live)")
    print(f"Force      : {'YES' if FORCE else 'NO'}")

    raw_header, raw_rows = read_source(
        service,
        SOURCE_SPREADSHEET_ID,
        SOURCE_SHEET,
        SOURCE_RANGE,
    )

    source_row_count = len(raw_rows)
    source_fingerprint = table_fingerprint(
        raw_header,
        raw_rows,
    )

    print(
        f'\nRAW "{SOURCE_SHEET}": '
        f"{source_row_count:,} row, {len(raw_header)} kolom."
    )
    print(f"Source fingerprint: {source_fingerprint[:16]}...")

    log_header, log_rows = read_table_if_exists(
        service,
        MONTHLY_OUTPUT_SPREADSHEET_ID,
        LOAD_LOG_SHEET,
    )

    validate_header_layout(
        log_header,
        LOAD_LOG_HEADERS,
        LOAD_LOG_SHEET,
    )

    latest_record = latest_load_record(
        log_header,
        log_rows,
        PIPELINE,
        load_id,
    )

    action, action_reason = determine_load_action(
        latest_record,
        SOURCE_SPREADSHEET_ID,
        source_fingerprint,
        CODE_VERSION,
        range_start=RANGE_START_DATE,
        range_end=RANGE_END_DATE,
        force=FORCE,
    )

    print(f"\nLoad action : {action}")
    print(f"Reason      : {action_reason}")

    if action == "SKIP":
        print(
            "\nSource pull ini sudah berhasil dimuat. "
            "Gunakan --force untuk memproses ulang load ID yang sama."
        )
        return

    validate_required_columns(raw_header)

    (
        typed_rows,
        _resolved_types,
        type_summary,
        type_errors,
    ) = coerce_rows_by_schema(
        raw_header,
        raw_rows,
        column_types=COLUMN_TYPES,
        strict=False,
        skip_invalid_rows=True,
    )

    print_type_diagnostics(
        type_summary,
        type_errors,
    )

    print(
        f"\nRows after type validation: "
        f"{len(typed_rows):,} of {source_row_count:,}"
    )

    fallback_mapping = {}

    if USE_PRODUCT_ID_MAPPING_FALLBACK:
        fallback_mapping, mapping_stats = load_product_brand_mapping(
            service,
            MAPPING_SPREADSHEET_ID,
            MAPPING_SHEET_GID,
            MAPPING_PRODUCT_ID_HEADER,
            MAPPING_BRAND_HEADER,
        )

        print("\nFallback mapping: ON")
        print(f"  Sheet          : {mapping_stats['sheet_title']}")
        print(f"  PC Product ID  : {mapping_stats['mapping_count']:,}")
        print(f"  Non-PC skipped : {mapping_stats['skipped_non_pc']:,}")
    else:
        print("\nFallback mapping: OFF")

    (
        header,
        rows,
        brand_status,
        missing_names,
        bundle_parent_counts,
    ) = add_brand_from_product_name(
        raw_header,
        typed_rows,
        fallback_mapping=fallback_mapping,
        product_name_header=COL_PRODUCT_NAME,
        product_id_header=COL_PRODUCT_ID,
        brand_header=COL_BRAND,
    )

    print_brand_diagnostics(
        brand_status,
        missing_names,
        bundle_parent_counts,
    )

    # The current source pull is aggregated without any RANGE filter.
    (
        source_groups,
        accepted_rows,
        skipped_blank,
        invalid_rows,
    ) = aggregate_brand_content(
        header,
        rows,
        content_id_header=COL_ROOM,
        date_header=COL_DATE,
        sum_columns=SUM_COLS,
        brand_header=COL_BRAND,
        row_validator=make_live_row_validator(header),
    )

    invalid_live_time = add_live_time_columns(
        source_groups,
        header,
    )

    base_output_header = header + EXTRA_HEADERS
    output_header = base_output_header + [JENIS_HEADER]

    source_all_rows = build_combined_from_groups(
        source_groups,
        base_output_header,
        GMV_COL,
    )

    validate_brand_column(
        output_header,
        source_all_rows,
        COL_BRAND,
    )

    print("\nSource-pull aggregation:")
    print(f"  Accepted source row       : {len(accepted_rows):,}")
    print(f"  Skipped blank Brand/Room  : {skipped_blank:,}")
    print(f"  Shifted / invalid row     : {len(invalid_rows):,}")
    print(f"  Unique Brand + Room ID    : {len(source_groups):,}")
    print(f"  Invalid LIVE time info    : {invalid_live_time:,}")
    print("  RANGE filter              : not applied at source-pull stage")

    print_reconciliation(
        "Source rows vs source-pull aggregate ALL",
        reconcile_sum_columns(
            header,
            accepted_rows,
            source_groups,
            SUM_COLS,
        ),
    )

    loaded_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    master_header = build_master_header(output_header)

    master_existing_header, master_existing_rows = read_table_if_exists(
        service,
        MONTHLY_OUTPUT_SPREADSHEET_ID,
        MASTER_ALL_SHEET,
    )

    validate_header_layout(
        master_existing_header,
        master_header,
        MASTER_ALL_SHEET,
    )

    source_master_rows = attach_master_metadata(
        source_all_rows,
        PIPELINE,
        load_id,
        SOURCE_SPREADSHEET_ID,
        source_fingerprint,
        CODE_VERSION,
        loaded_at,
    )

    merged_master_rows, master_write_mode = prepare_master_upsert(
        master_existing_header,
        master_existing_rows,
        master_header,
        source_master_rows,
        PIPELINE,
        load_id,
    )

    # Re-aggregate every stored source contribution before applying Date Live.
    monthly_all = rebuild_monthly_output(
        master_header,
        merged_master_rows,
        output_header,
        content_id_header=COL_ROOM,
        date_header=COL_DATE,
        sum_columns=SUM_COLS,
        gmv_column=GMV_COL,
        brand_header=COL_BRAND,
        jenis_header=JENIS_HEADER,
        consistent_fields=[FILTER_DATE_COLUMN],
    )

    cumulative_ranges = cumulative_load_ranges(
        log_header,
        log_rows,
        PIPELINE,
        load_id,
        RANGE_START_DATE,
        RANGE_END_DATE,
    )

    monthly_range_groups, invalid_monthly_dates = filter_groups_by_date_ranges(
        base_output_header,
        monthly_all["groups"],
        FILTER_DATE_COLUMN,
        cumulative_ranges,
    )

    monthly_range_rows = build_combined_from_groups(
        monthly_range_groups,
        base_output_header,
        GMV_COL,
    )

    validate_brand_column(
        output_header,
        monthly_all["rows"],
        COL_BRAND,
    )
    validate_brand_column(
        output_header,
        monthly_range_rows,
        COL_BRAND,
    )

    print_reconciliation(
        "Monthly groups RANGE vs GABUNGAN RANGE",
        reconcile_groups_vs_gabungan(
            base_output_header,
            monthly_range_groups,
            monthly_range_rows,
            SUM_COLS,
        ),
    )

    gmv_index = find_header_index(
        base_output_header,
        GMV_COL,
    )

    if WRITE_BRAND_TABS:
        brand_all_tabs = ensure_monthly_brand_tabs(
            build_brand_tabs(
                monthly_all["groups"],
                gmv_index,
                BRAND_ALL_SUFFIX,
            ),
            BRAND_ALL_SUFFIX,
        )

        brand_range_tabs = ensure_monthly_brand_tabs(
            build_brand_tabs(
                monthly_range_groups,
                gmv_index,
                BRAND_RANGE_SUFFIX,
            ),
            BRAND_RANGE_SUFFIX,
        )
    else:
        # Tidak dibutuhkan untuk GABUNGAN karena monthly_all/monthly_range_rows
        # sudah dibangun langsung dari aggregated groups.
        brand_all_tabs = {}
        brand_range_tabs = {}

    print("\nCumulative RANGE after this run:")
    for start_date, end_date in cumulative_ranges:
        print(f"  {start_date.isoformat()} -> {end_date.isoformat()}")

    print("\nMonthly output after this run:")
    print(f"  {FINAL_ALL_SHEET:<28} {len(monthly_all['rows']):>10,} row")
    print(f"  {FINAL_RANGE_SHEET:<28} {len(monthly_range_rows):>10,} row")
    print(f"  {MASTER_ALL_SHEET:<28} {len(merged_master_rows):>10,} row")
    print(
        f"  Write Brand tabs          : "
        f"{'YES' if WRITE_BRAND_TABS else 'NO':>10}"
    )
    print(f"  Brand tabs ALL            : {len(brand_all_tabs):>10,}")
    print(f"  Brand tabs RANGE          : {len(brand_range_tabs):>10,}")
    print(f"  Invalid/blank Date Live   : {invalid_monthly_dates:>10,}")

    print("\nMonthly metric total RANGE:")
    for name, total in metric_totals(
        output_header,
        monthly_range_rows,
        SUM_COLS,
    ).items():
        print(f"  {name}: {total:,}")

    target_dimensions = {
        MASTER_ALL_SHEET: (
            len(merged_master_rows) + 1,
            len(master_header),
        ),
        FINAL_ALL_SHEET: (
            len(monthly_all["rows"]) + 1,
            len(output_header),
        ),
        FINAL_RANGE_SHEET: (
            len(monthly_range_rows) + 1,
            len(output_header),
        ),
        LOAD_LOG_SHEET: (
            len(log_rows) + 2,
            len(LOAD_LOG_HEADERS),
        ),
    }

    if WRITE_BRAND_TABS:
        for title, rows_for_tab in brand_all_tabs.items():
            target_dimensions[title] = (
                len(rows_for_tab) + 1,
                len(output_header),
            )

        for title, rows_for_tab in brand_range_tabs.items():
            target_dimensions[title] = (
                len(rows_for_tab) + 1,
                len(output_header),
            )

    print("\nValidasi kapasitas:")
    validate_target_capacity(
        service,
        MONTHLY_OUTPUT_SPREADSHEET_ID,
        target_dimensions,
        label="MONTHLY LS OUTPUT",
    )

    if not WRITE:
        print("\nDRY RUN selesai. Tidak ada data ditulis.")
        print("Write mode: python split_brand_live_combined.py --write")
        return

    try:
        print("\nMenulis persistent master...")
        upsert_master_table(
            service,
            MONTHLY_OUTPUT_SPREADSHEET_ID,
            MASTER_ALL_SHEET,
            master_header,
            master_existing_header,
            master_existing_rows,
            merged_master_rows,
            source_master_rows,
            master_write_mode,
        )
        print(
            f'  "{MASTER_ALL_SHEET}" selesai '
            f"({master_write_mode}, {len(merged_master_rows):,} row)"
        )

        print("\nMenulis monthly final tabs...")
        replace_table(
            service,
            MONTHLY_OUTPUT_SPREADSHEET_ID,
            FINAL_ALL_SHEET,
            output_header,
            monthly_all["rows"],
        )
        print(f'  "{FINAL_ALL_SHEET}" selesai')

        replace_table(
            service,
            MONTHLY_OUTPUT_SPREADSHEET_ID,
            FINAL_RANGE_SHEET,
            output_header,
            monthly_range_rows,
        )
        print(f'  "{FINAL_RANGE_SHEET}" selesai')

        if WRITE_BRAND_TABS:
            print("\nMenulis monthly brand tabs...")
            for title, tab_rows in brand_all_tabs.items():
                replace_table(
                    service,
                    MONTHLY_OUTPUT_SPREADSHEET_ID,
                    title,
                    output_header,
                    tab_rows,
                )
                print(f'  "{title}" selesai ({len(tab_rows):,} row)')

            for title, tab_rows in brand_range_tabs.items():
                replace_table(
                    service,
                    MONTHLY_OUTPUT_SPREADSHEET_ID,
                    title,
                    output_header,
                    tab_rows,
                )
                print(f'  "{title}" selesai ({len(tab_rows):,} row)')
        else:
            print(
                "\nMonthly brand tabs: SKIP "
                "(WRITE_BRAND_TABS = False)."
            )
            print(
                "  Tab Brand lama yang sudah ada tidak dihapus / tidak di-update."
            )

        log_row = build_load_log_row(
            load_id,
            source_fingerprint,
            action,
            source_row_count,
            len(source_all_rows),
            len(monthly_all["rows"]),
            len(monthly_range_rows),
            loaded_at,
        )

        append_load_log(
            service,
            MONTHLY_OUTPUT_SPREADSHEET_ID,
            log_row,
            log_header,
            log_rows,
        )
        print(f'  "{LOAD_LOG_SHEET}" updated')

    except Exception:
        print(
            "\nWRITE GAGAL. _LOAD_LOG belum ditandai SUCCESS untuk run ini "
            "jika kegagalan terjadi sebelum langkah terakhir."
        )
        raise

    print("\nSELESAI.")


if __name__ == "__main__":
    raise SystemExit("Jalankan melalui Streamlit dengan perintah streamlit run app.py.")
