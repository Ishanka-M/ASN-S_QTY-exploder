"""
ASN S_QTY Exploder
------------------
S_QTY column එකේ අගයෙන් line ගාන හදනවා. S_QTY = 10 නම් line 10ක්,
හැම line එකකම S_QTY = 1. අනිත් හැම column එකකම details එහෙමම තියෙනවා.

HU_ID generate කරන්න පුළුවන්:
  - Keep original  : තියෙන HU_ID එකම තියනවා
  - None           : HU_ID හිස්ව (blank)
  - Letters+Number : <Letters> + number (1,2,3.. line ගානට)
  - Item+Number    : <DISPLAY_ITEM_NUMBER> + number (1,2,3.. line ගානට)

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
    target = name.strip().upper()
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=1, column=c).value
        if v is not None and str(v).strip().upper() == target:
            return c
    return None


def is_blank_row(ws, r):
    for c in range(1, ws.max_column + 1):
        if ws.cell(row=r, column=c).value not in (None, ""):
            return False
    return True


def to_int_qty(val):
    if val in (None, ""):
        return 1
    try:
        n = int(float(val))
    except (ValueError, TypeError):
        return 1
    return n if n >= 1 else 1


def build_hu(prefix, sep, num, pad):
    numstr = str(num).zfill(pad) if pad and pad > 0 else str(num)
    return f"{prefix}{sep}{numstr}"


def explode_sheet(ws, sqty_col, line_col=None, renumber_line=True,
                  hu_col=None, item_col=None, hu_cfg=None):
    max_col = ws.max_column

    src_rows = []
    for r in range(2, ws.max_row + 1):
        if is_blank_row(ws, r):
            continue
        values = [ws.cell(row=r, column=c).value for c in range(1, max_col + 1)]
        styles = [copy.copy(ws.cell(row=r, column=c)._style) for c in range(1, max_col + 1)]
        src_rows.append((values, styles))

    exploded = []
    for values, styles in src_rows:
        n = to_int_qty(values[sqty_col - 1])
        for _ in range(n):
            nv = list(values)
            nv[sqty_col - 1] = 1
            exploded.append((nv, styles))

    if renumber_line and line_col is not None:
        for i, (nv, _s) in enumerate(exploded, start=1):
            nv[line_col - 1] = i

    if hu_col is not None and hu_cfg and hu_cfg.get("mode") != "keep":
        mode = hu_cfg["mode"]
        counter = int(hu_cfg.get("start", 1))
        for nv, _s in exploded:
            if mode == "none":
                nv[hu_col - 1] = None
            else:
                if mode == "item":
                    pv = nv[item_col - 1] if item_col else None
                    prefix = "" if pv is None else str(pv)
                else:
                    prefix = hu_cfg.get("letters", "")
                nv[hu_col - 1] = build_hu(prefix, hu_cfg.get("sep", ""), counter, hu_cfg.get("pad", 0))
                counter += 1

    if ws.max_row >= 2:
        ws.delete_rows(2, ws.max_row - 1)

    for ridx, (nv, styles) in enumerate(exploded, start=2):
        for cidx in range(1, max_col + 1):
            cell = ws.cell(row=ridx, column=cidx, value=nv[cidx - 1])
            cell._style = copy.copy(styles[cidx - 1])

    return len(src_rows), len(exploded)


# ---------- UI ----------

st.title("📦 ASN S_QTY Exploder")
st.caption("S_QTY අගයෙන් line ගාන හදනවා · S_QTY 10 → line 10ක් · හැම line එකකම S_QTY = 1 · අනිත් columns details එහෙමම")

uploaded = st.file_uploader("Excel file එකක් upload කරන්න (.xlsx)", type=["xlsx"])
if uploaded is None:
    st.info("⬆️ GENERAL Format හරි ATTRIBUTE EXTENTD format හරි Excel එකක් upload කරන්න.")
    st.stop()

raw = uploaded.read()
wb = load_workbook(io.BytesIO(raw))

sheet_info = {}
for name in wb.sheetnames:
    ws = wb[name]
    sheet_info[name] = {
        "sqty": find_col(ws, "S_QTY"),
        "line": find_col(ws, "ASN_LINE_NUMBER"),
        "hu": find_col(ws, "HU_ID"),
        "item": find_col(ws, "DISPLAY_ITEM_NUMBER"),
        "rows": max(ws.max_row - 1, 0),
    }

sqty_sheets = [n for n, i in sheet_info.items() if i["sqty"]]
if not sqty_sheets:
    st.error("මේ file එකේ කිසිම sheet එකක `S_QTY` column එකක් නෑ. Header row එක row 1 ද කියලා බලන්න.")
    st.stop()

col1, col2 = st.columns([2, 1])
with col1:
    target_sheet = st.selectbox("Explode කරන්න ඕන sheet එක", options=sqty_sheets, index=0)
with col2:
    renumber = st.checkbox("ASN_LINE_NUMBER අලුතෙන් 1..N", value=True)

info = sheet_info[target_sheet]
ws_target = wb[target_sheet]

# sheet name එක මොකක් වුණත් S_QTY column එකෙන් auto-detect වෙනවා
if len(sqty_sheets) == 1:
    st.caption(f"✓ `S_QTY` column එක හම්බුණේ sheet **{target_sheet}** එකේ (sheet name එක වෙනස් වුණත් auto-detect වෙනවා).")
else:
    st.caption(f"ℹ️ `S_QTY` column එක තියෙන sheets {len(sqty_sheets)}ක්: {', '.join(sqty_sheets)} — එකක් තෝරන්න.")

st.subheader("🏷️ HU_ID generate")
if info["hu"] is None:
    st.warning("මේ sheet එකේ `HU_ID` column එකක් නෑ — HU_ID generate skip වෙනවා.")
    hu_cfg = {"mode": "keep"}
else:
    hu_mode_label = st.radio(
        "HU_ID හදන විදිය",
        options=["Keep original (නොවෙනස්ව)", "None (හිස්ව)",
                 "Letters + Number (1,2,3..)", "DISPLAY_ITEM_NUMBER + Number (1,2,3..)"],
        index=0,
    )
    mode_map = {
        "Keep original (නොවෙනස්ව)": "keep",
        "None (හිස්ව)": "none",
        "Letters + Number (1,2,3..)": "letters",
        "DISPLAY_ITEM_NUMBER + Number (1,2,3..)": "item",
    }
    mode = mode_map[hu_mode_label]
    hu_cfg = {"mode": mode, "letters": "", "sep": "", "start": 1, "pad": 0}

    if mode in ("letters", "item"):
        c1, c2, c3, c4 = st.columns(4)
        if mode == "letters":
            with c1:
                hu_cfg["letters"] = st.text_input("Letters (prefix)", value="HU")
        else:
            with c1:
                st.text_input("Prefix", value="DISPLAY_ITEM_NUMBER", disabled=True,
                              help="හැම row එකකම DISPLAY_ITEM_NUMBER value එක prefix විදිහට යනවා")
        with c2:
            hu_cfg["start"] = st.number_input("Number පටන් ගන්න", value=1, step=1,
                                              help="මේකෙන් පටන් අරන් 1,2,3.. විදිහට line ගානට වැඩි වෙනවා")
        with c3:
            hu_cfg["sep"] = st.text_input("Separator (optional)", value="")
        with c4:
            hu_cfg["pad"] = st.number_input("Zero-pad ඉලක්කම් ගාන", value=0, min_value=0, step=1,
                                            help="උදා: 3 නම් 1 → 001. 0 නම් pad නෑ.")

        item_example = None
        if mode == "item" and info["item"]:
            item_example = ws_target.cell(row=2, column=info["item"]).value
        prev = []
        for k in range(3):
            pfx = hu_cfg["letters"] if mode == "letters" else ("" if item_example is None else str(item_example))
            prev.append(build_hu(pfx, hu_cfg["sep"], int(hu_cfg["start"]) + k, hu_cfg["pad"]))
        st.caption("Preview: " + "  ·  ".join(f"`{p}`" for p in prev) + "  · …")

preview_total = 0
for r in range(2, ws_target.max_row + 1):
    if is_blank_row(ws_target, r):
        continue
    preview_total += to_int_qty(ws_target.cell(row=r, column=info["sqty"]).value)

m1, m2, m3 = st.columns(3)
m1.metric("දැනට data rows", info["rows"])
m2.metric("Explode වුණාම rows", preview_total)
m3.metric("අලුතෙන් එකතු වන rows", preview_total - info["rows"])

st.divider()

if st.button("🚀 Explode & Generate", type="primary", use_container_width=True):
    n_src, n_out = explode_sheet(
        ws_target, sqty_col=info["sqty"], line_col=info["line"], renumber_line=renumber,
        hu_col=info["hu"], item_col=info["item"], hu_cfg=hu_cfg,
    )
    out_buf = io.BytesIO()
    wb.save(out_buf)
    out_buf.seek(0)
    st.success(f"✅ සාර්ථකයි · {n_src} rows → {n_out} rows ({target_sheet})")
    base = uploaded.name.rsplit(".", 1)[0]
    st.download_button("⬇️ Download කරන්න", data=out_buf, file_name=f"{base}_EXPLODED.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       use_container_width=True)
    st.caption("Output preview (මුල් rows 15)")
    headers = [ws_target.cell(row=1, column=c).value for c in range(1, ws_target.max_column + 1)]
    preview = []
    for r in range(2, min(ws_target.max_row, 16) + 1):
        preview.append([ws_target.cell(row=r, column=c).value for c in range(1, ws_target.max_column + 1)])
    st.dataframe({(h or f"col{i}"): [row[i] for row in preview] for i, h in enumerate(headers)},
                 use_container_width=True, height=360)
