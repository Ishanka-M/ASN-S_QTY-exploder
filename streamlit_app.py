"""
ASN S_QTY Exploder
------------------
S_QTY column එකේ අගයෙන් line ගාන හදනවා. S_QTY = 10 නම් line 10ක්,
හැම line එකකම S_QTY = 1. අනිත් හැම column එකකම details එහෙමම තියෙනවා.

GENERAL Format (ASN Master + ASN DETAIL) සහ
ATTRIBUTE EXTENTD format (Physical ASN) දෙකම support කරනවා.
"""

import io
import copy
import streamlit as st
from openpyxl import load_workbook

st.set_page_config(page_title="ASN S_QTY Exploder", page_icon="📦", layout="wide")

# ---------- helpers ----------

def find_col(ws, name):
    """Header row (row 1) එකේ column name එක හොයනවා. 1-based index එක return."""
    target = name.strip().upper()
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=1, column=c).value
        if v is not None and str(v).strip().upper() == target:
            return c
    return None


def is_blank_row(ws, r):
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=r, column=c).value
        if v not in (None, ""):
            return False
    return True


def to_int_qty(val):
    """S_QTY value එක line count එකකට හරවනවා. blank/<1 නම් 1."""
    if val in (None, ""):
        return 1
    try:
        n = int(float(val))
    except (ValueError, TypeError):
        return 1
    return n if n >= 1 else 1


def explode_sheet(ws, sqty_col, line_col=None, renumber_line=True):
    """
    Target sheet එකේ data rows explode කරනවා.
    - row එකක S_QTY = N නම් → ඒ row එක N පාරක්, හැම එකකම S_QTY = 1
    - renumber_line=True නම් ASN_LINE_NUMBER 1..N විදිහට renumber
    - අනිත් සියලුම column / cell styles එහෙමම copy වෙනවා
    Return: (original_data_rows, exploded_rows)
    """
    max_col = ws.max_column

    # 1) data rows කියවනවා (values + styles)
    src_rows = []
    for r in range(2, ws.max_row + 1):
        if is_blank_row(ws, r):
            continue
        values = [ws.cell(row=r, column=c).value for c in range(1, max_col + 1)]
        styles = [copy.copy(ws.cell(row=r, column=c)._style) for c in range(1, max_col + 1)]
        src_rows.append((values, styles))

    # 2) explode
    exploded = []
    for values, styles in src_rows:
        n = to_int_qty(values[sqty_col - 1])
        for _ in range(n):
            nv = list(values)
            nv[sqty_col - 1] = 1  # explode කරපු හැම line එකකම S_QTY = 1
            exploded.append((nv, styles))

    # 3) line number renumber (optional)
    if renumber_line and line_col is not None:
        for i, (nv, _styles) in enumerate(exploded, start=1):
            nv[line_col - 1] = i

    # 4) පරණ data rows මකනවා (header row 1 රකිනවා)
    if ws.max_row >= 2:
        ws.delete_rows(2, ws.max_row - 1)

    # 5) අලුත් rows ලියනවා (value + style)
    for ridx, (nv, styles) in enumerate(exploded, start=2):
        for cidx in range(1, max_col + 1):
            cell = ws.cell(row=ridx, column=cidx, value=nv[cidx - 1])
            cell._style = copy.copy(styles[cidx - 1])

    return len(src_rows), len(exploded)


# ---------- UI ----------

st.title("📦 ASN S_QTY Exploder")
st.caption(
    "S_QTY අගයෙන් line ගාන හදනවා · S_QTY 10 → line 10ක් · හැම line එකකම S_QTY = 1 · "
    "අනිත් columns details එහෙමම තියෙනවා"
)

uploaded = st.file_uploader("Excel file එකක් upload කරන්න (.xlsx)", type=["xlsx"])

if uploaded is None:
    st.info("⬆️ GENERAL Format හරි ATTRIBUTE EXTENTD format හරි Excel එකක් upload කරන්න.")
    st.stop()

# workbook එක memory එකේ load කරනවා
raw = uploaded.read()
wb = load_workbook(io.BytesIO(raw))

# S_QTY තියෙන sheets හොයනවා
sheet_info = {}
for name in wb.sheetnames:
    ws = wb[name]
    sheet_info[name] = {
        "sqty": find_col(ws, "S_QTY"),
        "line": find_col(ws, "ASN_LINE_NUMBER"),
        "rows": max(ws.max_row - 1, 0),
    }

sqty_sheets = [n for n, i in sheet_info.items() if i["sqty"]]

if not sqty_sheets:
    st.error("මේ file එකේ කිසිම sheet එකක `S_QTY` column එකක් නෑ. Header row එක row 1 ද කියලා බලන්න.")
    st.stop()

col1, col2 = st.columns([2, 1])
with col1:
    target_sheet = st.selectbox(
        "Explode කරන්න ඕන sheet එක",
        options=sqty_sheets,
        index=0,
        help="S_QTY column එක තියෙන sheets විතරයි මෙතන පෙන්නන්නේ.",
    )
with col2:
    renumber = st.checkbox(
        "ASN_LINE_NUMBER අලුතෙන් 1..N කරන්න",
        value=True,
        help="Explode වුණාම line number 1, 2, 3... විදිහට ආයෙ දාගන්නවා.",
    )

info = sheet_info[target_sheet]
ws_target = wb[target_sheet]

# preview: explode වුණාම කී rows එනවද කියලා count එකක්
preview_total = 0
for r in range(2, ws_target.max_row + 1):
    if is_blank_row(ws_target, r):
        continue
    preview_total += to_int_qty(ws_target.cell(row=r, column=info["sqty"]).value)

m1, m2, m3 = st.columns(3)
m1.metric("දැනට data rows", info["rows"])
m2.metric("Explode වුණාම rows", preview_total)
m3.metric("අලුතෙන් එකතු වන rows", preview_total - info["rows"])

if info["line"] is None and renumber:
    st.warning("මේ sheet එකේ `ASN_LINE_NUMBER` column එකක් නෑ — line renumber එක skip වෙනවා.")

st.divider()

if st.button("🚀 Explode & Generate", type="primary", use_container_width=True):
    n_src, n_out = explode_sheet(
        ws_target,
        sqty_col=info["sqty"],
        line_col=info["line"],
        renumber_line=renumber,
    )

    out_buf = io.BytesIO()
    wb.save(out_buf)
    out_buf.seek(0)

    st.success(f"✅ සාර්ථකයි · {n_src} rows → {n_out} rows ({target_sheet})")

    base = uploaded.name.rsplit(".", 1)[0]
    out_name = f"{base}_EXPLODED.xlsx"
    st.download_button(
        "⬇️ Download කරන්න",
        data=out_buf,
        file_name=out_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    # preview table
    st.caption("Output preview (මුල් rows 15)")
    headers = [ws_target.cell(row=1, column=c).value for c in range(1, ws_target.max_column + 1)]
    preview = []
    for r in range(2, min(ws_target.max_row, 16) + 1):
        preview.append([ws_target.cell(row=r, column=c).value for c in range(1, ws_target.max_column + 1)])
    st.dataframe({h or f"col{i}": [row[i] for row in preview] for i, h in enumerate(headers)},
                 use_container_width=True, height=360)
