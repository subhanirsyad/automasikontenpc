"""Streamlit interface for the PC Content import."""

#.\.venv\Scripts\python.exe -m streamlit run app.py (Cara Buka Streamlit)
# format File - Tiktok Afiliate SV = TAP SV PC TANGGAL BULAN TAHUN (3 NOV 2026)
#             - Tiktok Afiliate LS = TAP LS PC TANGGAL BULAN TAHUN (3 NOV 2026)
#             - Seller Center SV   = SC SV PC TANGGAL BULAN TAHUN (3 NOV 2026)
#             - Seller Center LS   = SC LS PC TANGGAL BULAN TAHUN (3 NOV 2026)
#             - Affiliate Order    = AFF GMV PC TANGGAL BULAN TAHUN (3 NOV 2026)

from __future__ import annotations

import hmac
import json
import logging
import os
from datetime import datetime, timedelta, timezone
from pathlib import Path
from tempfile import TemporaryDirectory
from threading import Lock
from time import monotonic

import streamlit as st
from filelock import FileLock, Timeout
from googleapiclient.errors import HttpError

import pipeline as pc

INPUT_LABELS = ("TAP_SV", "SC_SV", "TAP_LS", "SC_LS")
COVERAGE_LABELS = ("TAP SV", "TAP LS", "SC SV", "SC LS")
UPLOAD_TYPES = ["xlsx", "csv"]
LIST_PRODUK_URL = "https://docs.google.com/spreadsheets/d/1O4dQVFYawbirEN5wpV0ygjNZtNFZPZH0tLX_Ov34fNA/edit#gid=437523491"


STYLE = """
<style>
:root {
  --pc-canvas: #f3f6fa;
  --pc-surface: #ffffff;
  --pc-ink: #10233b;
  --pc-muted: #53657a;
  --pc-border: #dbe4ee;
  --pc-accent: #0c6a62;
  --pc-tint: #eaf6f3;
  --pc-hero-bg: #0e2139;
  --pc-hero-text: #f5fbff;
  --pc-hero-muted: #c2d6e8;
}
@media (prefers-color-scheme: dark) {
  :root {
    --pc-canvas: #111a27;
    --pc-surface: #192538;
    --pc-ink: #eef5ff;
    --pc-muted: #aabdd2;
    --pc-border: #33465d;
    --pc-accent: #8be6d0;
    --pc-tint: #193d3b;
  }
}
[data-testid="stAppViewContainer"] { background: var(--pc-canvas); }
[data-testid="stHeader"] { background: transparent; }
[data-testid="stMainBlockContainer"] { max-width: 78rem; padding-top: 3.4rem; padding-bottom: 2rem; }
.pc-hero {
  position: relative;
  overflow: hidden;
  background: radial-gradient(ellipse at 90% 0%, #24496a 0%, #142d48 34%, var(--pc-hero-bg) 70%);
  border: 1px solid #29445e;
  border-radius: 1.1rem;
  padding: 1.65rem 2rem;
  color: var(--pc-hero-text);
}
.pc-hero::after {
  content: "";
  position: absolute;
  right: -2.5rem;
  top: -8rem;
  width: 21rem;
  height: 21rem;
  border: 1px solid #a5daca35;
  border-radius: 50%;
  box-shadow: 0 0 0 2.7rem #a5daca12, 0 0 0 5.4rem #a5daca0b;
  pointer-events: none;
}
.pc-brand {
  position: relative;
  z-index: 1;
  display: inline-flex;
  gap: .62rem;
  align-items: center;
  color: #c6f4e8;
  font-size: .72rem;
  font-weight: 750;
  letter-spacing: .14em;
}
.pc-brand-mark {
  display: inline-flex;
  width: 1.9rem;
  height: 1.9rem;
  border-radius: .54rem;
  align-items: center;
  justify-content: center;
  background: #95f0d7;
  color: #0e332e;
  font-size: .64rem;
  letter-spacing: .04em;
}
.pc-hero h1 {
  position: relative;
  z-index: 1;
  padding: 0;
  margin: .75rem 0 .4rem;
  color: var(--pc-hero-text);
  font-size: clamp(1.85rem, 3.9vw, 2.8rem);
  font-weight: 760;
  letter-spacing: -.045em;
  line-height: 1.12;
}
.pc-hero p {
  position: relative;
  z-index: 1;
  margin: 0;
  max-width: 36rem;
  color: var(--pc-hero-muted);
  font-size: 1rem;
}
.pc-mode {
  display: flex;
  gap: .7rem .9rem;
  align-items: center;
  flex-wrap: wrap;
  background: var(--pc-surface);
  border: 1px solid var(--pc-border);
  border-radius: .75rem;
  padding: .7rem .85rem;
  color: var(--pc-muted);
  font-size: .875rem;
}
.pc-mode span {
  flex: none;
  border-radius: 99px;
  padding: .24rem .65rem;
  background: var(--pc-tint);
  color: var(--pc-accent);
  font-weight: 700;
  font-size: .79rem;
}
.pc-step {
  display: flex;
  align-items: center;
  gap: .65rem;
  margin: 0 0 .36rem;
}
.pc-step span {
  border-radius: .45rem;
  background: var(--pc-tint);
  color: var(--pc-accent);
  padding: .32rem .48rem;
  font-weight: 750;
  font-size: .76rem;
  letter-spacing: .04em;
}
.pc-step h2 {
  color: var(--pc-ink);
  font-size: 1.22rem;
  line-height: 1.25;
  margin: 0;
  padding: 0;
}
.pc-section-note {
  margin: 0;
  color: var(--pc-muted);
  font-size: .86rem;
  line-height: 1.55;
}
.pc-source-label {
  margin: .2rem 0 -.22rem;
  color: var(--pc-ink);
  font-size: .91rem;
  font-weight: 720;
}
[data-testid="stForm"],
[data-testid="stMetric"] { background: var(--pc-surface); border-color: var(--pc-border); }
[data-testid="stForm"] { border-radius: 1rem; padding: 1.05rem; }
[data-testid="stMetric"] { border-radius: .75rem; }
[data-testid="stFileUploaderDropzone"] {
  padding: .58rem .65rem;
  border: 1px dashed var(--pc-border);
  border-radius: .6rem;
}
[data-testid="stFileUploaderDropzoneInstructions"] svg { width: 1.15rem; height: 1.15rem; }
[data-testid="stFileUploaderDropzoneInstructions"] small { line-height: 1.25; }
[data-testid="stWidgetLabel"] p { font-size: .88rem; font-weight: 650; }
[data-testid="stFormSubmitButton"] button:is([kind="primary"], [kind="primaryFormSubmit"]),
.st-key-pc_send button[kind="primary"] {
  background: #136c5f;
  border-color: #136c5f;
  color: #fff;
  font-weight: 680;
  min-height: 2.65rem;
}
[data-testid="stFormSubmitButton"] button:is([kind="primary"], [kind="primaryFormSubmit"]):hover,
.st-key-pc_send button[kind="primary"]:hover { background: #0c5349; border-color: #0c5349; }
.st-key-pc_send button[kind="primary"]:disabled {
  background: var(--pc-tint);
  color: var(--pc-muted);
  border-color: var(--pc-border);
}
[data-testid="stTabs"] [data-baseweb="tab"] { font-weight: 650; }
@media (max-width: 720px) {
  [data-testid="stMainBlockContainer"] { padding: 3rem .85rem 1.2rem; }
  .pc-hero { padding: 1.15rem 1.2rem; }
  .pc-hero::after { opacity: .42; }
  [data-testid="stForm"] { padding: .8rem; }
  [data-testid="stForm"] [data-testid="stHorizontalBlock"] { flex-wrap: wrap; }
  [data-testid="stForm"] [data-testid="stColumn"] { flex: 1 1 100%; width: 100%; min-width: 0; }
}
</style>
"""


def section(number, title, note=""):
    st.html(
        f'<div class="pc-step"><span>{number}</span><h2>{title}</h2></div>'
        + (f'<p class="pc-section-note">{note}</p>' if note else "")
    )


def page_header():
    st.html(STYLE)
    st.html("""
        <header class="pc-hero">
          <div class="pc-brand"><span class="pc-brand-mark" aria-hidden="true">WEI</span>
            AUTOMATE WEI</div>
          <h1>PC Content Import</h1>
          <p>Siapkan batch TAP, SC, dan AFF. Cek hasil SV dan LS sebelum mengirim ke Google Sheets.</p>
        </header>
    """)


def mode_banner(target, credentials, enabled):
    if not target:
        mode = "Preview lokal"
        detail = ("Mapping List Produk tersedia. Hasil tampil di app."
                  if credentials else "Hasil tampil di app. Mapping bisa dimasukkan lewat file.")
    elif enabled:
        mode = "Append ke workbook asli" if target == pc.NOVEMBER_ORIGINAL_ID else "Append ke salinan"
        detail = "Hanya menambah ID + Brand baru. Baris yang sudah ada dilewati."
    else:
        mode = "Mode baca"
        detail = "Preview memeriksa ID + Brand yang sudah ada di Google Sheets."
    st.html(f'<div class="pc-mode"><span>{mode}</span>{detail}</div>')


def clear_preview():
    st.session_state.pop("preview", None)
    st.session_state.pop("confirm_copy", None)


def clear_destination():
    clear_preview()
    st.session_state.pop("destination_period", None)
    st.session_state.pop("confirm_resume_original", None)


def upload_destination(credentials, *, allow_november_original=False):
    with st.container(border=True):
        link = st.text_input(
            "Link template PC Content November (tujuan upload)",
            key="destination_link", on_change=clear_destination,
            placeholder="https://docs.google.com/spreadsheets/d/...",
            help="Tempel link workbook tujuan yang berisi tab RAW SV dan RAW LS.")
        st.caption("Kosongkan untuk preview lokal. Link List Produk di bawah dipakai untuk mapping.")
    target = pc.sid(link)
    pc.check_target_allowed(target, allow_november_original=allow_november_original)
    if target and not credentials:
        raise pc.StopRun("Letakkan credentials.json di samping app.py untuk mengakses tujuan Google Sheets.")
    return target


def upload_form(credentials, *, destination_valid=True):
    with st.form("preview_form"):
        source_panel, config_panel = st.columns([2, 1], gap="large")
        with source_panel:
            section("01", "File sumber", "Masukkan empat export TAP/SC dan minimal satu export AFF Order.")
            files = {}
            for channel, suffix in (("Short Video", "SV"), ("Livestream", "LS")):
                st.html(f'<h3 class="pc-source-label">{channel}</h3>')
                for prefix, slot in zip(("TAP", "SC"), st.columns(2)):
                    label = f"{prefix}_{suffix}"
                    with slot:
                        files[label] = st.file_uploader(
                            label.replace("_", " "), type=UPLOAD_TYPES, key=label)
            st.html('<h3 class="pc-source-label">Affiliate</h3>')
            aff = st.file_uploader("AFF Order", type=UPLOAD_TYPES, accept_multiple_files=True,
                                   help="Satu atau beberapa export AFF untuk SV dan LS.")

        with config_panel:
            section("02", "Periode dan mapping", "Atur rentang tanggal dan referensi produk SC.")
            today = datetime.now(timezone(timedelta(hours=7))).date()
            dates = st.date_input(
                "Periode konten", value=(today - timedelta(days=min(6, today.day - 1)), today),
                format="DD/MM/YYYY")
            st.caption("Tanggal awal dan akhir harus dalam satu bulan kalender.")
            mapping_link = st.text_input(
                "Link Google Sheets List Produk", value=LIST_PRODUK_URL if credentials else "",
                placeholder="https://docs.google.com/spreadsheets/...",
                help="Mapping Product ID, nama produk, dan brand.")
            with st.expander("File mapping alternatif"):
                st.caption("Kosongkan link List Produk untuk memakai file lokal.")
                mapping = st.file_uploader(
                    "File mapping alternatif (opsional)", type=UPLOAD_TYPES,
                    help="Pakai export dari tab List Produk dalam format XLSX atau CSV.")
            st.space("small")
            prepare = st.form_submit_button("Proses dan preview", type="primary", width="stretch",
                                            disabled=not destination_valid)
            st.caption("Preview menampilkan hasil per tab sebelum konfirmasi upload.")
    return files, aff, dates, mapping_link, mapping, prepare


@st.cache_resource
def staging_lock():
    # The legacy backend loader temporarily changes sys.argv during import.
    return Lock()


def local_service_account():
    path = Path(__file__).resolve().with_name("credentials.json")
    try:
        with path.open(encoding="utf-8-sig") as file:
            credentials = json.load(file)
    except FileNotFoundError:
        return {}
    except (OSError, ValueError):
        raise pc.StopRun("credentials.json di samping app.py tidak bisa dibaca sebagai JSON.") from None
    fields = ("client_email", "private_key", "token_uri")
    if (not isinstance(credentials, dict) or credentials.get("type") != "service_account"
            or any(not isinstance(credentials.get(field), str) or not credentials[field].strip()
                   for field in fields)):
        raise pc.StopRun("credentials.json harus berupa key JSON untuk Google service account.")
    return credentials


def configuration():
    try:
        settings = dict(st.secrets.get("pc_content", {}))
        credentials = dict(st.secrets.get("gcp_service_account", {}))
    except FileNotFoundError as exc:
        if getattr(exc, "error_id", None) not in {None, "no-secrets-found"}:
            raise pc.StopRun(
                "Streamlit gagal membaca secrets.toml. Untuk mode lokal dengan credentials.json, "
                "cukup simpan [pc_content] dan enable_write di secrets.toml.") from None
        settings, credentials = {}, {}
    password = str(settings.get("app_password", ""))
    local_only = st.get_option("server.address") in {"127.0.0.1", "::1"}
    if local_only and not credentials:
        credentials = local_service_account()
    if credentials and len(password) < 16 and (password or not local_only):
        raise pc.StopRun(
            "Isi app_password minimal 16 karakter, atau jalankan Streamlit dengan "
            "--server.address 127.0.0.1 untuk akses lokal tanpa password.")
    return (password, credentials, settings.get("enable_write") is True,
            settings.get("allow_november_original") is True)


def require_password(password):
    if not password:
        return
    # Session state stores only a keyed fingerprint, not the password.
    fingerprint = hmac.digest(password.encode(), b"pc-content/session", "sha256")
    if st.session_state.get("authenticated") == fingerprint:
        if st.sidebar.button("Keluar"):
            st.session_state.clear()
            st.rerun()
        return
    _, login_column, _ = st.columns([1, 2, 1])
    with login_column, st.form("login"):
        st.subheader("Masuk ke PC Content Import")
        entered = st.text_input("Password aplikasi", type="password")
        if st.form_submit_button("Masuk", type="primary", width="stretch"):
            if hmac.compare_digest(entered.encode(), password.encode()):
                st.session_state["authenticated"] = fingerprint
                st.rerun()
            st.error("Password aplikasi salah.")
    st.stop()


def save_upload(folder, name, uploaded):
    suffix = Path(uploaded.name).suffix.lower()
    if suffix not in {".xlsx", ".csv"}:
        raise pc.StopRun(f"{name} harus berupa .xlsx atau .csv.")
    path = folder / (name + suffix)
    with path.open("wb") as stream:
        stream.write(uploaded.getbuffer())
    return str(path)


def read_target(credentials, cfg):
    reader = pc.Google(credentials, readonly=True)
    meta = pc.target_check(reader, cfg)
    schemas = {kind: pc.schema(reader, cfg["pc"], kind, meta) for kind in pc.TABS}
    return reader, meta["properties"]["title"], schemas



def destination_period(reader, target, title, schemas):
    groups = {label: {"dates": set(), "rows": 0, "unreadable": 0} for label in COVERAGE_LABELS}
    unknown_sources = 0
    for kind, schema in schemas.items():
        date_field, id_field = ("Post date", "Video ID") if kind == "SV" else ("Live Date", "Livestream room ID")
        letters = [pc.col(schema["idx"][pc.norm(field)] + 1) for field in ("Source", date_field, id_field)]
        tab = pc.quote(schema["tab"])
        for first in range(2, schema["rows"] + 1, pc.SCAN_ROWS):
            last = min(first + pc.SCAN_ROWS - 1, schema["rows"])
            ranges = [f"{tab}!{letter}{first}:{letter}{last}" for letter in letters]
            columns = reader.ranges(target, ranges, "UNFORMATTED_VALUE", "COLUMNS")
            if len(columns) != len(ranges):
                raise pc.StopRun(f"Pembacaan periode {schema['tab']} belum lengkap. Coba cek lagi.")
            sources, dates, ids = [(column.get("values") or [[]])[0] for column in columns]
            for offset, identifier in enumerate(ids):
                if identifier is None or not str(identifier).strip():
                    continue
                source = str(sources[offset]).strip().upper() if offset < len(sources) else ""
                group = groups.get(f"{source} {kind}")
                if group is None:
                    unknown_sources += 1
                    continue
                raw = dates[offset] if offset < len(dates) else ""
                try:
                    # Native Sheets dates are serial numbers. Drop time-of-day if present.
                    parsed = pc.standard_date(int(raw) if type(raw) in (float, int) else raw)
                except (pc.StopRun, ValueError, OverflowError):
                    parsed = ""
                if parsed:
                    group["dates"].add(parsed)
                    group["rows"] += 1
                else:
                    group["unreadable"] += 1
    return {"target": target, "title": title, "groups": groups, "unknown_sources": unknown_sources}


def check_destination_period(target, credentials, progress, *, allow_november_original=False):
    progress("Membaca tanggal dari RAW SV dan RAW LS")
    cfg = {"pc": target, "allow_november_original": allow_november_original}
    reader, title, schemas = read_target(credentials, cfg)
    return destination_period(reader, target, title, schemas)


def show_destination_period(snapshot, upload_period=None):
    st.markdown("**Periode yang sudah ada di tujuan**")
    st.caption(f'{snapshot["title"]} · RAW SV memakai Post date, RAW LS memakai Live Date.')
    rows = []
    for label, group in snapshot["groups"].items():
        days = group["dates"]
        row = {"Sumber": label, "Dari": min(days) if days else "—",
               "Sampai": max(days) if days else "—", "Baris bertanggal": group["rows"]}
        if upload_period:
            start, end = upload_period
            row["Hari sudah ada di periode upload"] = sum(start <= day <= end for day in days)
        rows.append(row)
    st.table(rows)
    unreadable = sum(group["unreadable"] for group in snapshot["groups"].values())
    if unreadable:
        st.caption(f"{unreadable:,} baris dengan tanggal kosong atau tidak terbaca dikecualikan dari rentang.")
    if snapshot["unknown_sources"]:
        st.caption(f'{snapshot["unknown_sources"]:,} baris dengan Source selain TAP atau SC dikecualikan.')

def build_preview(files, aff, mapping, dates, target, credentials, progress, mapping_link="",
                  *, allow_november_original=False):
    start = monotonic()
    if not all(files.values()) or not aff or len(dates) != 2:
        raise pc.StopRun("Isi tanggal awal dan akhir, empat export TAP/SC, dan minimal satu AFF.")
    mapping_link = mapping_link.strip()
    if mapping_link and mapping:
        raise pc.StopRun("Pilih link List Produk atau file mapping. Isi salah satu.")
    if mapping_link and not credentials:
        raise pc.StopRun("Letakkan credentials.json di samping app.py untuk membaca link List Produk.")
    with TemporaryDirectory(prefix="pc-content-") as tmp:
        folder = Path(tmp)
        paths = {label: save_upload(folder, label, upload) for label, upload in files.items()}
        paths["AFF"] = [save_upload(folder, f"AFF_{i}", file) for i, file in enumerate(aff, 1)]
        if mapping:
            paths["PRODUCT_BRAND_MAPPING"] = save_upload(folder, "PRODUCT_BRAND_MAPPING", mapping)
        cfg = pc.build_config(paths, dates[0], dates[1], target,
                              allow_november_original=allow_november_original)
        if mapping_link:
            progress("Membaca List Produk dari Google Sheets")
            cfg["sc_product_mapping"] = pc.read_product_mapping(
                pc.Google(credentials, readonly=True), mapping_link)
        with staging_lock():
            staged, summary = pc.assemble(cfg, pc.engines(), progress)
    cfg = {key: cfg[key] for key in ("pc", "start", "end")}
    cfg["allow_november_original"] = allow_november_original
    staging_seconds = monotonic() - start
    if not target:
        eligible = {kind: [row for row in rows if pc.is_valid(kind, row, cfg) == "OK"]
                    for kind, rows in staged.items()}
        return {"cfg": cfg, "summary": summary, "staging_seconds": staging_seconds,
                "samples": {kind: preview_rows(rows) for kind, rows in eligible.items()},
                "local_counts": {kind: len(rows) for kind, rows in eligible.items()},
                "plan": None}
    progress("Memeriksa ID dan Brand yang sudah ada")
    reader, title, schemas = read_target(credentials, cfg)
    additions, qc, indexes = pc.classify_candidates(reader, cfg, staged, schemas)
    plan = pc.make_plan(reader, cfg, schemas, additions, indexes)
    progress("Membaca periode yang sudah ada di tujuan")
    coverage = destination_period(reader, target, title, schemas)
    return {"cfg": cfg, "title": title, "summary": summary, "qc": qc, "plan": plan,
            "destination_period": coverage,
            "staging_seconds": staging_seconds, "qc_seconds": monotonic() - start - staging_seconds,
            "samples": {kind: preview_rows(rows) for kind, rows in additions.items()}}


def preview_rows(rows):
    return [{key: value for key, value in row.items() if key != "_present"} for row in rows[:30]]


def append_conflict_cells(reader, target, schema, kind, first, rows):
    positions = range(first, first + len(rows))
    grid = pc.input_grid(reader, target, schema, kind, positions,
                         entered=True, include_formulas=True)
    conflicts, missing = [], []
    has_literal = False
    for row, record in zip(positions, rows):
        actual = grid.get(row, {})
        expected = pc.append_values(schema, kind, row, record)
        columns = schema["idx"]
        for field in sorted(actual.keys() | expected.keys(), key=lambda field: columns[pc.norm(field)]):
            cell = actual.get(field, {})
            entered = cell.get("userEnteredValue", {})
            wanted = expected.get(field, {})
            formula = "formulaValue" in entered
            has_literal |= bool(entered) and not formula
            if entered == wanted or (formula and field in pc.MANAGED_FORMULAS[kind] and not wanted):
                continue
            address = f'{schema["tab"]}!{pc.col(columns[pc.norm(field)] + 1)}{row}'
            detail = {"Sel": address, "Kolom": field}
            if formula:
                detail["Perbedaan"] = "Rumus berbeda dari rencana"
                if field in pc.MANAGED_FORMULAS[kind]:
                    missing.append(detail)
                    continue
            elif entered:
                detail["Perbedaan"] = "Nilai yang sudah ada berbeda dari rencana"
            elif cell.get("effectiveValue"):
                detail["Perbedaan"] = "Sel berisi nilai turunan dari sheet"
            else:
                missing.append({**detail, "Perbedaan": "Sel kosong dalam batch yang sudah terisi sebagian"})
                continue
            conflicts.append(detail)
    return conflicts or (missing if has_literal else [])


def inspect_pending_append(path, target, credentials, progress, *, allow_november_original=False):
    with FileLock(str(path.parent / "write.lock"), timeout=0):
        plan = pc.read_plan(path)
        if plan["state"] != "PENDING":
            raise pc.StopRun("Tidak ada batch tertunda untuk diperiksa.")
        cfg = {"pc": target, "start": plan["start"], "end": plan["end"],
               "allow_november_original": allow_november_original}
        reader, _, schemas = read_target(credentials, cfg)
        pc.validate_frozen(cfg, plan, schemas)
        for action, kind, first, rows, _ in pc.chunk_groups(plan, schemas):
            if action != "APPEND":
                continue
            schema = schemas[kind]
            existing = rows[:max(0, schema["rows"] - first + 1)]
            if existing:
                progress(f'Memeriksa {schema["tab"]} mulai baris {first:,}')
                conflicts = append_conflict_cells(reader, target, schema, kind, first, existing)
                if conflicts:
                    return conflicts
    return []


def write_or_resume(path, target, credentials, enabled, progress, preview=None, *,
                    allow_november_original=False, original_confirmed=False):
    start = monotonic()
    if not enabled:
        raise pc.StopRun("Mode tulis belum diaktifkan di secrets.")
    if preview is not None and preview["cfg"]["pc"] != target:
        raise pc.StopRun("Tujuan sudah berubah. Buat preview baru sebelum mengirim batch.")
    path.parent.mkdir(parents=True, exist_ok=True)
    with FileLock(str(path.parent / "write.lock"), timeout=0):
        resume = preview is None
        saved = pc.read_plan(path) if path.exists() else None
        if resume:
            if not saved or saved["state"] != "PENDING":
                raise pc.StopRun("Tidak ada batch yang menunggu untuk dilanjutkan.")
            plan = saved
            cfg = {"pc": target, "start": plan["start"], "end": plan["end"]}
        else:
            if saved and saved["state"] == "PENDING":
                raise pc.StopRun("Lanjutkan batch yang tertunda sebelum mengirim batch baru.")
            cfg = dict(preview["cfg"])
            plan = preview["plan"]
        cfg.update(approved_copy=target, allow_november_original=allow_november_original,
                   original_confirmed=original_confirmed)
        pc.check_write_armed(cfg)
        reader, _, schemas = read_target(credentials, cfg)
        pc.validate_frozen(cfg, plan, schemas)
        if not resume:
            pc.save_plan(path, plan)
        pc.perform_chunks(reader, cfg, plan, path, credentials, schemas,
                          resume=resume, progress=progress)
    return monotonic() - start


def run_action(label, action):
    status = st.status(label, expanded=True)
    try:
        result = action(lambda message: status.update(label=message))
    except Timeout:
        message = "Batch lain sedang diproses. Tunggu sampai selesai."
    except (pc.StopRun, ValueError) as exc:
        message = str(exc)
    except HttpError as exc:
        message = f"Google Sheets HTTP {exc.resp.status}. Cek akses atau kuota."
    except Exception:
        logging.exception("PC Content operation failed")
        message = "Ada error saat memproses batch. Detail tersedia di log server."
    else:
        status.update(label="Selesai", state="complete", expanded=False)
        return True, result
    status.update(label="Proses berhenti", state="error")
    st.error(message)
    return False, message


def show_preview(preview):
    cfg = preview["cfg"]
    plan = preview["plan"]
    section("03", "Preview hasil")
    detail = f'{preview.get("title", "Lokal")} · {cfg["start"]} s.d. {cfg["end"]} · Proses file {preview["staging_seconds"]:.1f} detik'
    if plan:
        detail += f' · QC {preview["qc_seconds"]:.1f} detik'
    st.caption(detail)
    for kind, tab in zip(pc.TABS, st.tabs(list(pc.TABS.values()))):
        with tab:
            source = preview["summary"][kind]
            if plan:
                qc = preview["qc"][kind]
                metrics = (("Baris baru", qc["new_rows"]), ("Sudah ada", qc["skipped_existing"]),
                           ("Perlu review", len(qc["quarantined"])))
            else:
                metrics = (("Baris lolos cek awal", preview["local_counts"][kind]),
                           ("TAP terproses", source["tap"]["range"]),
                           ("SC terproses", source["sc_brand_unique_not_in_tap"]))
            for column, (label, count) in zip(st.columns(len(metrics)), metrics):
                column.metric(label, f"{count:,}".replace(",", "."), border=True)
            overlap = f'{source["sc_already_in_tap"]:,}'.replace(",", ".")
            st.caption(f"Overlap SC dengan TAP yang dilewati · {overlap}")

            if plan:
                if qc["quarantined"]:
                    with st.expander(f'Baris yang perlu review ({len(qc["quarantined"])})'):
                        st.dataframe(qc["quarantined"], hide_index=True, width="stretch")
                if qc["warnings"]:
                    st.caption(" · ".join(qc["warnings"]))

            samples = preview["samples"][kind]
            if samples:
                st.markdown("**Contoh baris baru**" if plan else "**Contoh hasil olah lokal**")
                st.caption(f"Menampilkan {len(samples)} baris sampel.")
                st.dataframe(samples, hide_index=True, width="stretch")
            elif not plan or not qc["new_rows"]:
                st.info("Belum ada baris baru untuk ditampilkan di tab ini.")


def main():
    st.set_page_config(page_title="PC Content Import", layout="wide")
    page_header()
    try:
        password, credentials, enabled, allow_november_original = configuration()
    except (pc.StopRun, ValueError) as exc:
        st.error(str(exc))
        st.stop()
    require_password(password)
    try:
        target = upload_destination(credentials, allow_november_original=allow_november_original)
    except (pc.StopRun, ValueError) as exc:
        st.error(str(exc))
        upload_form(credentials, destination_valid=False)
        return
    original = target == pc.NOVEMBER_ORIGINAL_ID
    if original:
        st.warning("Tujuan yang dipilih adalah workbook asli PC Content November.")
    if target:
        with st.container(border=True):
            if st.button("Cek periode di template"):
                st.session_state.pop("destination_period", None)
                ok, result = run_action("Memeriksa periode template", lambda progress: check_destination_period(
                    target, credentials, progress, allow_november_original=allow_november_original))
                if ok:
                    st.session_state["destination_period"] = result
            period_panel = st.container()
    mode_banner(target, credentials, enabled)
    elapsed = st.session_state.pop("upload_done", None)
    if elapsed is not None:
        st.success(f"Batch selesai dan sudah diverifikasi di target dalam {elapsed:.1f} detik.")
    message = st.session_state.pop("upload_failed", None)
    if message:
        st.error(message)

    state_dir = Path(os.environ.get("PC_CONTENT_RUN_DIR", ".pc_runs"))
    path = state_dir / target / "append_plan.json"
    pending = False
    if target and path.is_file():
        try:
            pending = pc.read_plan(path)["state"] == "PENDING"
        except (pc.StopRun, ValueError, OSError) as exc:
            st.error(f"Rencana batch tidak bisa dibaca. Cek log dan file rencana. {exc}")
            st.stop()
    if pending:
        st.warning("Ada batch yang belum selesai. Lanjutkan dari batch terakhir yang sudah diverifikasi.")
        if st.button("Cek konflik batch (read-only)"):
            ok, conflicts = run_action("Memeriksa batch tertunda", lambda progress: inspect_pending_append(
                path, target, credentials, progress, allow_november_original=allow_november_original))
            if ok and conflicts:
                st.caption(f"{len(conflicts):,} sel berbeda di batch pertama yang bentrok.")
                st.dataframe(conflicts[:20], hide_index=True, width="stretch")
            elif ok:
                st.info("Tidak ditemukan konflik append pada baris yang sudah tersedia.")
        resume_confirmed = False
        if original:
            resume_confirmed = st.checkbox("Saya konfirmasi lanjutkan batch ke workbook asli November.",
                                           key="confirm_resume_original", disabled=not enabled)
        if st.button("Lanjutkan batch", disabled=not enabled or (original and not resume_confirmed)):
            ok, elapsed = run_action("Memulai ulang batch", lambda progress: write_or_resume(
                path, target, credentials, enabled, progress, allow_november_original=allow_november_original,
                original_confirmed=resume_confirmed))
            if ok:
                clear_destination()
                st.session_state["upload_done"] = elapsed
                st.rerun()

    files, aff, dates, mapping_link, mapping, prepare = upload_form(credentials)

    if prepare:
        clear_preview()
        ok, result = run_action("Membaca file", lambda progress: build_preview(
            files, aff, dates=dates, mapping=mapping, target=target,
            credentials=credentials, progress=progress, mapping_link=mapping_link,
            allow_november_original=allow_november_original))
        if ok:
            st.session_state["preview"] = result
            if target:
                st.session_state["destination_period"] = result["destination_period"]

    preview = st.session_state.get("preview")
    snapshot = st.session_state.get("destination_period")
    if target and snapshot and snapshot["target"] == target:
        with period_panel:
            upload_period = None
            if preview and preview["cfg"]["pc"] == target:
                upload_period = (preview["cfg"]["start"], preview["cfg"]["end"])
            if upload_period:
                st.caption(f"Periode upload yang dipilih · {upload_period[0]} s.d. {upload_period[1]}")
            show_destination_period(snapshot, upload_period)
    if preview and preview["cfg"]["pc"] == target:
        show_preview(preview)
        plan = preview["plan"]
        if plan and any(info["rows"] for info in plan["kinds"].values()):
            with st.container(border=True, key="pc_send"):
                st.markdown("**Kirim batch ke Google Sheets**")
                st.caption(f'[Buka workbook tujuan](https://docs.google.com/spreadsheets/d/{target}/edit)')
                confirmation = (f'Saya konfirmasi append ke workbook asli {preview["title"]}.'
                                if original else f'Kirim hanya ke salinan {preview["title"]}.')
                confirmed = st.checkbox(confirmation, key="confirm_copy", disabled=not enabled or pending)
                if st.button("Kirim hasil preview", disabled=not enabled or pending or not confirmed):
                    ok, result = run_action("Menyiapkan upload", lambda progress: write_or_resume(
                        path, target, credentials, enabled, progress, preview=preview,
                        allow_november_original=allow_november_original, original_confirmed=original and confirmed))
                    if ok:
                        clear_destination()
                        st.session_state["upload_done"] = result
                    else:
                        st.session_state["upload_failed"] = result
                    st.rerun()


if __name__ == "__main__":
    main()
