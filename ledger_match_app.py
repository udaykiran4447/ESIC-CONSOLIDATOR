"""
Ledger Match — E-way Bill ⇄ Books Reconciliation
--------------------------------------------------
A Streamlit app that matches E-way Bill invoice values against Books,
voucher-wise, within a configurable tolerance.

Run with:
    pip install streamlit openpyxl pandas --break-system-packages
    streamlit run ledger_match_app.py
"""

import io
import re
import difflib
import hashlib
from copy import copy

import streamlit as st
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill

st.set_page_config(page_title="Ledger Match", page_icon="📒", layout="wide")

# ----------------------------------------------------------------------------
# Styling — ledger / register aesthetic
# ----------------------------------------------------------------------------
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&display=swap');
    html, body, [class*="css"] { font-family: -apple-system, "Segoe UI", sans-serif; }
    .ledger-title {
        font-family: Georgia, 'Times New Roman', serif;
        font-size: 40px;
        font-weight: 700;
        border-bottom: 3px double #16243B;
        padding-bottom: 10px;
        margin-bottom: 4px;
        color: #16243B;
    }
    .ledger-dek { color: #3A4A63; font-size: 14px; margin-bottom: 24px; }
    .section-label {
        font-family: 'IBM Plex Mono', monospace;
        font-size: 11px;
        letter-spacing: 2px;
        text-transform: uppercase;
        color: #9C7A2E;
        border-bottom: 1px solid #B9B096;
        padding-bottom: 6px;
        margin: 28px 0 12px 0;
    }
    .stamp-ok { color: #2E6B54; font-weight: 700; }
    .stamp-miss { color: #9E3A34; font-weight: 700; }
    div[data-testid="stMetricValue"] { font-family: 'IBM Plex Mono', monospace; }
    </style>
    <div class="ledger-title">Ledger Match</div>
    <div class="ledger-dek">Match E-way Bill invoice values against Books, voucher-wise — upload both sides, set a tolerance, and get a marked-up workbook back.</div>
    """,
    unsafe_allow_html=True,
)

# ----------------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------------

def norm(s):
    """Normalize a header string for loose matching."""
    if s is None:
        return ""
    return re.sub(r"[^a-z0-9]+", "", str(s).strip().lower())


BOOKS_ALIASES = {
    "voucher_no": ["voucherno", "voucherno.", "vouchernumber", "voucher"],
    "invoice_value": ["invoicevalue", "invoiceamount", "invoiceamt"],
    "party_name": ["partyname", "party", "vendorname", "vendor"],
}
EWB_ALIASES = {
    "invoice_value": [
        "invoicevalue", "asperewaybillinvoicevalue", "asperewaybill",
        "ewbvalue", "ewayvillvalue", "invoiceamount",
    ],
    "ewb_no": ["ewbno"],
    "doc_no": ["docno"],
    "party_name": ["fromtradername", "partyname", "tradername", "fromparty"],
}


def parse_amount(v):
    """Parse a cell value into a float, handling commas/currency/brackets. None if not parseable."""
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    if not s:
        return None
    neg = False
    if s.startswith("(") and s.endswith(")"):
        neg = True
        s = s[1:-1]
    s = re.sub(r"[₹$,\s]", "", s)
    try:
        n = float(s)
    except ValueError:
        return None
    return -n if neg else n


def cell_text(v):
    if v is None:
        return ""
    return str(v)


_CORP_SUFFIXES = re.compile(
    r"\b(PRIVATE|PVT|LIMITED|LTD|LLP|INC|CO|COMPANY|ENGG|ENGINEERING|"
    r"ENTERPRISES|WORKS|INDUSTRIES|TRADERS|CORP|CORPORATION)\b"
)


def normalize_name(s):
    """Uppercase, strip punctuation and common corporate suffixes, for fuzzy comparison."""
    if not s:
        return ""
    s = str(s).upper()
    s = re.sub(r"[^A-Z0-9 ]", " ", s)
    s = _CORP_SUFFIXES.sub(" ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def name_similarity(a, b):
    """0-1 similarity score between two party names, robust to minor spelling/formatting differences."""
    na, nb = normalize_name(a), normalize_name(b)
    if not na or not nb:
        return 0.0
    return difflib.SequenceMatcher(None, na, nb).ratio()


def pick_main_sheet(wb):
    """Pick the worksheet with the most rows."""
    return max(wb.worksheets, key=lambda ws: ws.max_row)


def detect_columns(ws, alias_map):
    """Return {key: column_index} for headers found in row 1, matched against alias_map."""
    found = {}
    header_row = next(ws.iter_rows(min_row=1, max_row=1))
    for cell in header_row:
        h = norm(cell.value)
        if not h:
            continue
        for key, aliases in alias_map.items():
            if key in found:
                continue
            if h in aliases:
                found[key] = cell.column
    return found


def get_or_create_col(ws, header_text, norm_key):
    """Find an existing column with this normalized header, or append a new one."""
    header_row = next(ws.iter_rows(min_row=1, max_row=1))
    for cell in header_row:
        if norm(cell.value) == norm_key:
            return cell.column
    col = ws.max_column + 1
    hc = ws.cell(row=1, column=col, value=header_text)
    hc.font = Font(bold=True, color="FFFFFFFF")
    hc.fill = PatternFill(start_color="FF16243B", end_color="FF16243B", fill_type="solid")
    return col


def load_workbook_from_upload(uploaded_file):
    # getvalue() is idempotent (unlike read(), which consumes the stream) — safe to call
    # again on later reruns without the file appearing empty.
    data = uploaded_file.getvalue()
    wb = openpyxl.load_workbook(io.BytesIO(data))
    ws = pick_main_sheet(wb)
    return wb, ws


def build_template(kind: str) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    if kind == "books":
        ws.title = "Books"
        headers = ["Voucher No", "Date", "Party Name", "Ledger", "Invoice Value", "Narration"]
        sample = ["19004", "2026-04-05", "Example Vendor Pvt Ltd", "Purchase Imports",
                  "8,60,366.44", "Sample row — replace with your data"]
    else:
        ws.title = "E way bills"
        headers = ["EWB No", "Doc No", "Doc Date", "Supply Type", "From Trader Name", "Invoice Value"]
        sample = ["4123456789012", "XIVBLR300310161", "2026-04-05", "Outward",
                  "Example Buyer Pvt Ltd", "8,60,366.44"]
    ws.append(headers)
    ws.append(sample)
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFFFF")
        cell.fill = PatternFill(start_color="FF16243B", end_color="FF16243B", fill_type="solid")
    for cell in ws[2]:
        cell.font = Font(italic=True, color="FF808080")
    for i in range(1, len(headers) + 1):
        ws.column_dimensions[get_column_letter(i)].width = 22
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def build_groups_and_ewb_rows(books_ws, books_cols, ewb_ws, ewb_cols, use_party):
    """Read both sheets into memory: Books voucher groups and E-way Bill rows."""
    groups = {}
    for row in books_ws.iter_rows(min_row=2):
        v_cell = row[books_cols["voucher_no"] - 1]
        vtext = cell_text(v_cell.value).strip()
        if not vtext:
            continue
        amt = parse_amount(row[books_cols["invoice_value"] - 1].value)
        if amt is None:
            continue
        g = groups.setdefault(vtext, {"rows": [], "sum": 0.0, "party": ""})
        g["rows"].append(v_cell.row)
        g["sum"] += amt
        if not g["party"] and "party_name" in books_cols:
            pname = cell_text(row[books_cols["party_name"] - 1].value).strip()
            if pname:
                g["party"] = pname
    for g in groups.values():
        g["sum"] = round(g["sum"], 2)

    ewb_rows = []
    for row in ewb_ws.iter_rows(min_row=2):
        cell = row[ewb_cols["invoice_value"] - 1]
        amt = parse_amount(cell.value)
        if amt is None:
            continue
        ref = ""
        if "ewb_no" in ewb_cols:
            ref = cell_text(row[ewb_cols["ewb_no"] - 1].value).strip()
        if not ref and "doc_no" in ewb_cols:
            ref = cell_text(row[ewb_cols["doc_no"] - 1].value).strip()
        if not ref:
            ref = f"Row {cell.row}"
        party = ""
        if "party_name" in ewb_cols:
            party = cell_text(row[ewb_cols["party_name"] - 1].value).strip()
        ewb_rows.append({"r": cell.row, "val": round(amt, 2), "ref": ref, "party": party})

    for g in groups.values():
        g["bkey"] = normalize_name(g["party"]) if use_party else ""
    for er in ewb_rows:
        er["ekey"] = normalize_name(er["party"]) if use_party else ""

    return groups, ewb_rows


def stage1_party_pairing(groups, ewb_rows, party_threshold):
    """Segregate Books/E-way Bills by party and auto-suggest pairings above `party_threshold`.

    Returns everything needed both to render an editable review table and, later, to run
    Stage 2 with whatever mapping the person confirms.
    """
    book_reps, ewb_reps = {}, {}
    vouchers_by_bkey = {}
    ewb_rows_by_ekey = {}

    for vn, g in groups.items():
        k = g["bkey"]
        if not k:
            continue
        vouchers_by_bkey.setdefault(k, []).append(vn)
        if k not in book_reps:
            book_reps[k] = g["party"]

    for er in ewb_rows:
        k = er["ekey"]
        if not k:
            continue
        ewb_rows_by_ekey.setdefault(k, []).append(er)
        if k not in ewb_reps:
            ewb_reps[k] = er["party"]

    candidates = []
    for bk, bname in book_reps.items():
        for ek, ename in ewb_reps.items():
            sim = name_similarity(bname, ename)
            if sim * 100 >= party_threshold:
                candidates.append((sim, bk, ek))
    candidates.sort(key=lambda c: -c[0])

    used_b, used_e = set(), set()
    auto_ekey_to_bkey, sim_by_ekey = {}, {}
    for sim, bk, ek in candidates:
        if bk in used_b or ek in used_e:
            continue
        auto_ekey_to_bkey[ek] = bk
        sim_by_ekey[ek] = sim
        used_b.add(bk)
        used_e.add(ek)

    return {
        "book_reps": book_reps,
        "ewb_reps": ewb_reps,
        "vouchers_by_bkey": vouchers_by_bkey,
        "ewb_rows_by_ekey": ewb_rows_by_ekey,
        "auto_ekey_to_bkey": auto_ekey_to_bkey,
        "sim_by_ekey": sim_by_ekey,
    }


def stage2_amount_match(groups, ewb_rows, use_party, ekey_to_bkeys, vouchers_by_bkey, tolerance, book_reps=None):
    """Sum Books by Voucher No. and greedy-match against E-way Bill Invoice Values within
    `tolerance`. When use_party, each E-way Bill row is only ever compared against the
    combined voucher list of the Books party (or parties — one E-way Bill party can map to
    several Books party names, e.g. multiple branches of the same vendor). Each voucher is
    used at most once.

    ekey_to_bkeys: dict mapping an E-way Bill party key -> list of Books party keys.
    """
    book_reps = book_reps or {}
    all_vns = list(groups.keys())

    def vouchers_for_bkeys(bkeys):
        vns = []
        for bk in bkeys:
            vns.extend(vouchers_by_bkey.get(bk, []))
        return vns

    pairs = []
    for er in ewb_rows:
        if use_party:
            bkeys = ekey_to_bkeys.get(er["ekey"], [])
            candidate_vns = vouchers_for_bkeys(bkeys)
        else:
            candidate_vns = all_vns
        for vn in candidate_vns:
            g = groups[vn]
            d = abs(g["sum"] - er["val"])
            if d <= tolerance:
                pairs.append({"d": d, "r": er["r"], "val": er["val"], "vn": vn,
                              "sum": g["sum"], "ref": er["ref"],
                              "books_party": g["party"], "ewb_party": er["party"]})
    pairs.sort(key=lambda p: p["d"])

    used_vn, used_row = set(), set()
    matches = []
    for p in pairs:
        if p["vn"] in used_vn or p["r"] in used_row:
            continue
        matches.append(p)
        used_vn.add(p["vn"])
        used_row.add(p["r"])

    matched_rows = {m["r"] for m in matches}
    unmatched_ewb = [er for er in ewb_rows if er["r"] not in matched_rows]

    unmatched_diag = []
    for er in unmatched_ewb:
        if use_party:
            ekey = er["ekey"]
            if not ekey:
                unmatched_diag.append({**er, "best": None, "party_status": "No party name on this row"})
                continue
            bkeys = ekey_to_bkeys.get(ekey, [])
            if not bkeys:
                unmatched_diag.append({**er, "best": None,
                                       "party_status": "No matching Books party confirmed"})
                continue
            best = None
            for vn in vouchers_for_bkeys(bkeys):
                d = abs(groups[vn]["sum"] - er["val"])
                if best is None or d < best["d"]:
                    best = {"d": d, "vn": vn, "sum": groups[vn]["sum"]}
            names = ", ".join(f"\u201c{book_reps.get(bk, bk)}\u201d" for bk in bkeys)
            unmatched_diag.append({**er, "best": best, "party_status": f"Party matched to Books {names}"})
        else:
            best = None
            for vn in all_vns:
                g = groups[vn]
                d = abs(g["sum"] - er["val"])
                if best is None or d < best["d"]:
                    best = {"d": d, "vn": vn, "sum": g["sum"]}
            unmatched_diag.append({**er, "best": best, "party_status": None})

    unused_vouchers = [{"vn": vn, "sum": g["sum"], "party": g["party"]}
                       for vn, g in groups.items() if vn not in used_vn]

    used_bkeys = {bk for bkeys in ekey_to_bkeys.values() for bk in bkeys} if use_party else set()
    unmatched_books_parties = [
        {"party": bname, "vouchers": len(vouchers_by_bkey.get(bk, [])),
         "sum": round(sum(groups[vn]["sum"] for vn in vouchers_by_bkey.get(bk, [])), 2)}
        for bk, bname in book_reps.items() if bk not in used_bkeys
    ]

    return {
        "groups": groups,
        "matches": matches,
        "unmatched_diag": unmatched_diag,
        "unused_vouchers": unused_vouchers,
        "total_ewb": len(ewb_rows),
        "use_party": use_party,
        "unmatched_books_parties": unmatched_books_parties,
    }


def workbook_to_bytes(wb) -> bytes:
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def fmt(n):
    return f"{n:,.2f}"


# ----------------------------------------------------------------------------
# 1. Templates
# ----------------------------------------------------------------------------
st.markdown('<div class="section-label">1 · Get the format right</div>', unsafe_allow_html=True)
c1, c2 = st.columns(2)
with c1:
    st.markdown("**Books template**")
    st.caption("One row per ledger line. Rows sharing a Voucher No. are summed together.")
    st.markdown(
        "- **Voucher No** — required\n"
        "- **Invoice Value** — required\n"
        "- Date, **Party Name**, Ledger, Narration — optional (Party Name needed for approx. name matching)"
    )
    st.download_button(
        "Download Books template",
        data=build_template("books"),
        file_name="Books - Template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
with c2:
    st.markdown("**E-way Bill template**")
    st.caption("One row per E-way bill. Its Invoice Value is what gets matched against Books.")
    st.markdown(
        "- **Invoice Value** — required\n"
        "- EWB No, Doc No, Doc Date, Supply Type, **From Trader Name** — optional (needed for approx. name matching)"
    )
    st.download_button(
        "Download E-way Bill template",
        data=build_template("ewb"),
        file_name="E way Bill - Template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

def file_signature(uploaded_file):
    """A content-based fingerprint for an uploaded file (hash of its actual bytes), used to
    detect when the person has swapped in a genuinely different file — independent of any
    Streamlit-internal file ID quirks, and independent of filename/size coincidences."""
    return hashlib.md5(uploaded_file.getvalue()).hexdigest()


def clear_downstream_state(clear_workbooks=False):
    """Wipe everything derived from the previous Books/E-way Bill files: party pairing,
    the review-table widgets, and any final results — so a fresh upload never shows stale
    data from a prior file."""
    keys = ["stage1", "result", "groups", "ewb_rows", "use_party_used", "tolerance_used"]
    if clear_workbooks:
        keys += ["books_wb", "books_ws_title", "books_cols",
                 "ewb_wb", "ewb_ws_title", "ewb_cols",
                 "books_file_sig", "ewb_file_sig"]
    for k in keys:
        st.session_state.pop(k, None)
    for k in list(st.session_state.keys()):
        if k.startswith("partymap_") or k.startswith("mapping_editor_"):
            del st.session_state[k]
    st.session_state["mapping_version"] = st.session_state.get("mapping_version", 0) + 1


# ----------------------------------------------------------------------------
# 2. Uploads
# ----------------------------------------------------------------------------
st.markdown('<div class="section-label">2 · Upload both sides</div>', unsafe_allow_html=True)

_, reset_col = st.columns([5, 1])
with reset_col:
    if st.button("🔄 Reset everything"):
        clear_downstream_state(clear_workbooks=True)
        st.rerun()

u1, u2 = st.columns(2)
with u1:
    books_file = st.file_uploader("Books (.xlsx)", type=["xlsx"], key="books_upload")
with u2:
    ewb_file = st.file_uploader("E-way Bills (.xlsx)", type=["xlsx"], key="ewb_upload")

books_ready = ewb_ready = False
books_cols = ewb_cols = None

if books_file is not None:
    sig = file_signature(books_file)
    if st.session_state.get("books_file_sig") != sig:
        clear_downstream_state()
        st.session_state["books_file_sig"] = sig

    books_wb, books_ws = load_workbook_from_upload(books_file)
    books_cols = detect_columns(books_ws, BOOKS_ALIASES)
    missing = [k for k in ("voucher_no", "invoice_value") if k not in books_cols]
    if missing:
        st.error(
            f"Books file: couldn't find {', '.join(missing)} column header in row 1. "
            "Use the Books template headers."
        )
    else:
        st.session_state["books_wb"] = books_wb
        st.session_state["books_ws_title"] = books_ws.title
        st.session_state["books_cols"] = books_cols
        books_ready = True
        party_note = " (Party Name detected)" if "party_name" in books_cols else " (no Party Name column found)"
        st.success(f"Books: {books_ws.max_row - 1} rows detected{party_note}")

if ewb_file is not None:
    sig = file_signature(ewb_file)
    if st.session_state.get("ewb_file_sig") != sig:
        clear_downstream_state()
        st.session_state["ewb_file_sig"] = sig

    ewb_wb, ewb_ws = load_workbook_from_upload(ewb_file)
    ewb_cols = detect_columns(ewb_ws, EWB_ALIASES)
    if "invoice_value" not in ewb_cols:
        st.error(
            "E-way Bill file: couldn't find an 'Invoice Value' column header in row 1. "
            "Use the E-way Bill template headers."
        )
    else:
        st.session_state["ewb_wb"] = ewb_wb
        st.session_state["ewb_ws_title"] = ewb_ws.title
        st.session_state["ewb_cols"] = ewb_cols
        ewb_ready = True
        party_note = " (Party Name detected)" if "party_name" in ewb_cols else " (no Party Name column found)"
        st.success(f"E-way Bills: {ewb_ws.max_row - 1} rows detected{party_note}")

# ----------------------------------------------------------------------------
# 3. Matching rules & run
# ----------------------------------------------------------------------------
# ----------------------------------------------------------------------------
# 3. Matching rules
# ----------------------------------------------------------------------------
st.markdown('<div class="section-label">3 · Set matching rules</div>', unsafe_allow_html=True)

party_cols_ok = bool(books_cols and "party_name" in books_cols and ewb_cols and "party_name" in ewb_cols)

use_party = st.checkbox(
    "Segregate by party first, then match amounts within each party (recommended for multi-party files)",
    value=party_cols_ok,
    disabled=not party_cols_ok,
    help="Groups Books by Party Name and E-way Bills by From Trader Name, suggests pairings for "
         "parties that look alike, and only then sums vouchers and matches amounts — separately per "
         "party. Prevents unrelated parties' invoices from being matched just because the amounts "
         "happen to be close. You'll get a chance to review and correct the suggested pairings.",
)
if not party_cols_ok:
    st.caption(
        "To enable party segregation, include a **Party Name** column in Books and a "
        "**From Trader Name** column in E-way Bills."
    )

party_threshold = 55.0
if use_party:
    party_threshold = st.slider(
        "Minimum name similarity to auto-suggest a pairing",
        min_value=0, max_value=100, value=55, step=5,
        help="Party names are compared after stripping punctuation and common suffixes "
             "(Pvt, Ltd, Enterprises, Engg, Works, etc.). This only affects the starting suggestions — "
             "you can override any of them by hand in the next step, either way.",
    )

t1, t2, t3 = st.columns([2, 1, 1])
with t1:
    preset = st.radio(
        "Treat amounts as matched when within",
        ["Exact (₹0)", "₹1", "₹100", "₹500", "₹1,000", "Custom"],
        index=1,
        horizontal=True,
    )
with t2:
    custom_tol = st.number_input("Custom ₹ tolerance", min_value=0.0, value=0.0, step=1.0,
                                  disabled=(preset != "Custom"))

preset_map = {"Exact (₹0)": 0.0, "₹1": 1.0, "₹100": 100.0, "₹500": 500.0, "₹1,000": 1000.0}
tolerance = custom_tol if preset == "Custom" else preset_map[preset]

with t3:
    st.write("")
    st.write("")
    find_clicked = st.button(
        "Segregate by party →" if use_party else "Run reconciliation →",
        type="primary", disabled=not (books_ready and ewb_ready),
    )

if find_clicked:
    groups, ewb_rows = build_groups_and_ewb_rows(
        st.session_state["books_wb"].worksheets[
            [ws.title for ws in st.session_state["books_wb"].worksheets].index(st.session_state["books_ws_title"])
        ],
        st.session_state["books_cols"],
        st.session_state["ewb_wb"].worksheets[
            [ws.title for ws in st.session_state["ewb_wb"].worksheets].index(st.session_state["ewb_ws_title"])
        ],
        st.session_state["ewb_cols"],
        use_party,
    )
    st.session_state["groups"] = groups
    st.session_state["ewb_rows"] = ewb_rows
    st.session_state["tolerance_used"] = tolerance
    st.session_state["use_party_used"] = use_party
    st.session_state.pop("result", None)

    if use_party:
        st.session_state["stage1"] = stage1_party_pairing(groups, ewb_rows, party_threshold)
        st.session_state["mapping_version"] = st.session_state.get("mapping_version", 0) + 1
    else:
        # No party segregation — go straight to amount matching across everything.
        result = stage2_amount_match(groups, ewb_rows, use_party=False, ekey_to_bkeys={},
                                      vouchers_by_bkey={}, tolerance=tolerance)
        st.session_state["result"] = result

# ----------------------------------------------------------------------------
# 4. Review & fix party matching (only when segregating by party)
# ----------------------------------------------------------------------------
if "stage1" in st.session_state and st.session_state.get("use_party_used"):
    st.markdown('<div class="section-label">4 · Review &amp; fix party matching</div>', unsafe_allow_html=True)
    st.caption(
        "Every E-way Bill party is listed below with its auto-suggested Books match (if any). "
        "Pick one or more Books party names for each — useful when the same vendor appears under "
        "several names in Books (different branches, spellings, etc.)."
    )

    s1 = st.session_state["stage1"]
    all_book_names = sorted(s1["book_reps"].values())
    name_to_bkey = {v: k for k, v in s1["book_reps"].items()}
    version = st.session_state["mapping_version"]

    ewb_keys_sorted = sorted(
        s1["ewb_reps"].keys(),
        key=lambda ek: (s1["auto_ekey_to_bkey"].get(ek) is not None, -s1["sim_by_ekey"].get(ek, 0)),
    )

    widget_keys = {}
    for ek in ewb_keys_sorted:
        rows_for = s1["ewb_rows_by_ekey"].get(ek, [])
        auto_bk = s1["auto_ekey_to_bkey"].get(ek)
        default = [s1["book_reps"][auto_bk]] if auto_bk else []
        sim_text = f"{s1['sim_by_ekey'][ek]*100:.0f}% match" if auto_bk else "no auto match"

        c1, c2 = st.columns([2, 3])
        with c1:
            st.markdown(f"**{s1['ewb_reps'][ek]}**")
            st.caption(f"{len(rows_for)} invoice(s) · {fmt(sum(r['val'] for r in rows_for))} · {sim_text}")
        with c2:
            key = f"partymap_{ek}_v{version}"
            widget_keys[ek] = key
            st.multiselect(
                "Books part(y/ies)", options=all_book_names, default=default,
                key=key, label_visibility="collapsed",
                placeholder="Select one or more Books parties, or leave empty for no match",
            )
        st.divider()

    if s1["book_reps"]:
        with st.expander("All Books parties (for reference)"):
            st.dataframe(
                [{"Party": bname, "Vouchers": len(s1["vouchers_by_bkey"].get(bk, []))}
                 for bk, bname in sorted(s1["book_reps"].items(), key=lambda kv: kv[1])],
                use_container_width=True, hide_index=True,
            )

    confirm_clicked = st.button("Confirm mapping & run amount matching →", type="primary")
    if confirm_clicked:
        final_ekey_to_bkeys = {}
        for ek, key in widget_keys.items():
            selected_names = st.session_state.get(key, [])
            bks = [name_to_bkey[n] for n in selected_names if n in name_to_bkey]
            if bks:
                final_ekey_to_bkeys[ek] = bks
        result = stage2_amount_match(
            st.session_state["groups"], st.session_state["ewb_rows"],
            use_party=True, ekey_to_bkeys=final_ekey_to_bkeys,
            vouchers_by_bkey=s1["vouchers_by_bkey"], tolerance=st.session_state["tolerance_used"],
            book_reps=s1["book_reps"],
        )
        # Keep the confirmed pairings around for display in Results
        result["party_pairs"] = [
            {"books_party": s1["book_reps"][bk], "ewb_party": s1["ewb_reps"][ek]}
            for ek, bks in final_ekey_to_bkeys.items() for bk in bks
        ]
        st.session_state["result"] = result

# ----------------------------------------------------------------------------
# 5. Results
# ----------------------------------------------------------------------------
if "result" in st.session_state:
    res = st.session_state["result"]
    st.markdown('<div class="section-label">5 · Results</div>', unsafe_allow_html=True)

    if res["use_party"]:
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Parties confirmed", len(res["party_pairs"]))
        m2.metric("Matched invoices", f"{len(res['matches'])} / {res['total_ewb']}")
        m3.metric("Unmatched E-way Bills", len(res["unmatched_diag"]))
        m4.metric("Unused Books vouchers", len(res["unused_vouchers"]))

        with st.expander(f"Confirmed party mapping ({len(res['party_pairs'])})"):
            st.dataframe(
                [{"Books party": p["books_party"], "E-way Bill party": p["ewb_party"]}
                 for p in sorted(res["party_pairs"], key=lambda p: p["books_party"])],
                use_container_width=True, hide_index=True,
            )
            if res["unmatched_books_parties"]:
                st.caption("Books parties left with no E-way Bill counterpart:")
                st.dataframe(
                    [{"Party": u["party"], "Vouchers": u["vouchers"], "Total": fmt(u["sum"])}
                     for u in res["unmatched_books_parties"]],
                    use_container_width=True, hide_index=True,
                )
    else:
        m1, m2, m3 = st.columns(3)
        m1.metric("Matched", f"{len(res['matches'])} / {res['total_ewb']}")
        m2.metric("Unmatched E-way Bills", len(res["unmatched_diag"]))
        m3.metric("Unused Books vouchers", len(res["unused_vouchers"]))

    with st.expander(f"Matched invoices ({len(res['matches'])})", expanded=True):
        rows = sorted(res["matches"], key=lambda m: m["r"])
        table_rows = []
        for m in rows:
            row = {
                "Voucher No": m["vn"],
                "EWB No": m["ref"],
                "Books sum": fmt(m["sum"]),
                "E-way Bill value": fmt(m["val"]),
                "Diff": fmt(abs(m["sum"] - m["val"])),
            }
            if res["use_party"]:
                row["Party"] = m["books_party"]
            table_rows.append(row)
        st.dataframe(table_rows, use_container_width=True, hide_index=True)

    with st.expander(f"Unmatched E-way Bills ({len(res['unmatched_diag'])})"):
        table_rows = []
        for u in res["unmatched_diag"]:
            row = {
                "Reference": u["ref"],
                "Invoice Value": fmt(u["val"]),
            }
            if res["use_party"]:
                row["Party"] = u["party"]
                row["Status"] = u["party_status"]
            row["Closest book voucher"] = u["best"]["vn"] if u["best"] else "—"
            row["Nearest diff"] = fmt(u["best"]["d"]) if u["best"] else "—"
            table_rows.append(row)
        st.dataframe(table_rows, use_container_width=True, hide_index=True)

    with st.expander(f"Unused Books vouchers ({len(res['unused_vouchers'])})"):
        table_rows = [
            {"Voucher No": u["vn"], "Books sum": fmt(u["sum"]),
             **({"Party": u["party"]} if res["use_party"] else {})}
            for u in res["unused_vouchers"]
        ]
        st.dataframe(table_rows, use_container_width=True, hide_index=True)

    # ---- Downloads ----
    d1, d2 = st.columns(2)

    with d1:
        books_wb = st.session_state["books_wb"]
        books_ws = books_wb.worksheets[
            [ws.title for ws in books_wb.worksheets].index(st.session_state["books_ws_title"])
        ]
        ewb_col = get_or_create_col(books_ws, "EWB No", "ewbno")
        # Reset the column first so a rerun at a different tolerance never leaves stale values
        for r in range(2, books_ws.max_row + 1):
            books_ws.cell(row=r, column=ewb_col).value = None
        for m in res["matches"]:
            for r in res["groups"][m["vn"]]["rows"]:
                books_ws.cell(row=r, column=ewb_col).value = m["ref"]
        st.download_button(
            "Download Books — with EWB No column",
            data=workbook_to_bytes(books_wb),
            file_name="Books - Matched.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    with d2:
        ewb_wb = st.session_state["ewb_wb"]
        ewb_ws = ewb_wb.worksheets[
            [ws.title for ws in ewb_wb.worksheets].index(st.session_state["ewb_ws_title"])
        ]
        c_books = get_or_create_col(ewb_ws, "As per Books", "asperbooks")
        c_diff = get_or_create_col(ewb_ws, "Diff", "diff")
        c_voucher = get_or_create_col(ewb_ws, "Matched Voucher No", "matchedvoucherno")
        by_row = {m["r"]: m for m in res["matches"]}
        for r in range(2, ewb_ws.max_row + 1):
            m = by_row.get(r)
            ewb_ws.cell(row=r, column=c_books).value = m["sum"] if m else None
            ewb_ws.cell(row=r, column=c_diff).value = round(m["val"] - m["sum"], 2) if m else None
            ewb_ws.cell(row=r, column=c_voucher).value = m["vn"] if m else None
        st.download_button(
            "Download E-way Bills — with matches",
            data=workbook_to_bytes(ewb_wb),
            file_name="E way Bills - Matched.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

st.markdown(
    """
    <div style="margin-top:40px; padding-top:14px; border-top:1px dotted #B9B096; font-size:11.5px; color:#3A4A63;">
    Reconciliation runs in two stages when party segregation is on: parties are paired up first
    (Books Party Name ⇄ E-way Bill From Trader Name, fuzzy-matched), then Books rows are summed
    per Voucher No. and matched against E-way Bill Invoice Values within your tolerance — separately
    within each paired party, so amounts are never compared across unrelated parties. Each voucher
    is used at most once.
    </div>
    """,
    unsafe_allow_html=True,
)
