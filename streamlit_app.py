"""
ASN S_QTY Exploder
------------------
S_QTY column එකේ අගයෙන් line ගාන හදනවා. S_QTY = 10 නම් line 10ක්,
හැම line එකකම S_QTY = 1. අනිත් හැම column එකකම details එහෙමම තියෙනවා.

Outputs:
  1) Exploded Excel (.xlsx)
  2) Summary PDF  — CLIENT_CODE + QR, DISPLAY_ASN_NUMBER + QR,
     සහ item-wise totals table

HU_ID generate: Keep original / None / Letters+Number / DISPLAY_ITEM_NUMBER+Number
Sheet name එක මොකක් වුණත් S_QTY column එකෙන් sheet එක auto-detect වෙනවා.
"""

import io
import copy
from functools import lru_cache
from collections import OrderedDict

import streamlit as st
from openpyxl import load_workbook
import qrcode
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                TableStyle, Image)
from reportlab.graphics.barcode import createBarcodeDrawing

st.set_page_config(page_title="ASN S_QTY Exploder", page_icon="📦", layout="wide")

SUMMARY_FIELDS = ["CLIENT_CODE", "DISPLAY_ASN_NUMBER", "DISPLAY_ITEM_NUMBER",
                  "LOT_NUMBER", "QUANTITY", "UOM", "S_QTY", "S_UOM",
                  "PACKAGE_TYPE", "GROSS_WEIGHT", "NET_WEIGHT"]

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


def to_num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def build_hu(prefix, sep, num, pad):
    numstr = str(num).zfill(pad) if pad and pad > 0 else str(num)
    return f"{prefix}{sep}{numstr}"


def explode_sheet(ws, sqty_col, line_col=None, renumber_line=True,
                  hu_col=None, item_col=None, hu_cfg=None):
    """
    Fast explode: iter_rows වලින් කියවලා, data rows delete කරලා, append වලින් ආයෙ ලියනවා.
    Header row 1 (styling එක්ක) රැකෙනවා. 'General' නොවන number formats ආයෙ apply වෙනවා.
    """
    max_col = ws.max_column

    # 'General' නොවන number formats විතරක් මතක තියාගන්නවා (reapply සඳහා)
    fmts = {}
    for c in range(1, max_col + 1):
        nf = ws.cell(row=2, column=c).number_format
        if nf and nf != "General":
            fmts[c] = nf

    # data rows fast read
    src_rows = [list(r) for r in ws.iter_rows(min_row=2, values_only=True)
                if any(v not in (None, "") for v in r)]

    # explode
    exploded = []
    for values in src_rows:
        n = to_int_qty(values[sqty_col - 1])
        base = list(values)
        base[sqty_col - 1] = 1
        for _ in range(n):
            exploded.append(list(base))

    if renumber_line and line_col is not None:
        for i, nv in enumerate(exploded, start=1):
            nv[line_col - 1] = i

    if hu_col is not None and hu_cfg and hu_cfg.get("mode") != "keep":
        mode = hu_cfg["mode"]
        counter = int(hu_cfg.get("start", 1))
        for nv in exploded:
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

    # පරණ data rows එකපාරටම මකනවා (header row 1 රැකෙනවා)
    if ws.max_row >= 2:
        ws.delete_rows(2, ws.max_row - 1)

    # fast write
    for nv in exploded:
        ws.append(nv)

    # number formats reapply (few cols විතරයි → fast)
    if fmts:
        for row in ws.iter_rows(min_row=2):
            for c, nf in fmts.items():
                row[c - 1].number_format = nf

    return len(src_rows), len(exploded)


# ---------- summary / PDF ----------

def fmt_num(x):
    if x == int(x):
        return f"{int(x)}"
    return f"{x:.3f}".rstrip("0").rstrip(".")


def summarize(ws, cols):
    """Exploded sheet එකෙන් item+lot wise totals + header values."""
    def first_val(ci):
        if not ci:
            return None
        for r in range(2, ws.max_row + 1):
            if not is_blank_row(ws, r):
                return ws.cell(row=r, column=ci).value
        return None

    client = first_val(cols.get("CLIENT_CODE"))
    asn = first_val(cols.get("DISPLAY_ASN_NUMBER"))

    groups = OrderedDict()
    for r in range(2, ws.max_row + 1):
        if is_blank_row(ws, r):
            continue
        item = ws.cell(row=r, column=cols["DISPLAY_ITEM_NUMBER"]).value if cols.get("DISPLAY_ITEM_NUMBER") else None
        lot = ws.cell(row=r, column=cols["LOT_NUMBER"]).value if cols.get("LOT_NUMBER") else None
        key = (item, lot)
        g = groups.setdefault(key, {"sum": {}, "has": {}, "uom": None, "suom": None,
                                    "pkg_num": 0.0, "pkg_has": False, "pkg_numeric": True, "pkg_txt": set()})
        for fld in ("QUANTITY", "S_QTY", "GROSS_WEIGHT", "NET_WEIGHT"):
            if cols.get(fld):
                x = to_num(ws.cell(row=r, column=cols[fld]).value)
                if x is not None:
                    g["sum"][fld] = g["sum"].get(fld, 0.0) + x
                    g["has"][fld] = True
        if cols.get("PACKAGE_TYPE"):
            pv = ws.cell(row=r, column=cols["PACKAGE_TYPE"]).value
            x = to_num(pv)
            if x is not None:
                g["pkg_num"] += x
                g["pkg_has"] = True
            elif pv not in (None, ""):
                g["pkg_numeric"] = False
                g["pkg_txt"].add(str(pv))
        if cols.get("UOM") and g["uom"] is None:
            g["uom"] = ws.cell(row=r, column=cols["UOM"]).value
        if cols.get("S_UOM") and g["suom"] is None:
            g["suom"] = ws.cell(row=r, column=cols["S_UOM"]).value
    return client, asn, groups


def qr_flowable(data, size_mm=22):
    qr = qrcode.QRCode(border=1, box_size=10)
    qr.add_data(str(data) if data not in (None, "") else "N/A")
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return Image(buf, width=size_mm * mm, height=size_mm * mm)


def make_summary_pdf(ws, cols, asn_title="ASN Summary"):
    client, asn, groups = summarize(ws, cols)

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, topMargin=14 * mm, bottomMargin=14 * mm,
                            leftMargin=12 * mm, rightMargin=12 * mm, title=asn_title)
    styles = getSampleStyleSheet()
    h = ParagraphStyle("h", parent=styles["Title"], fontSize=16, spaceAfter=6)
    lbl = ParagraphStyle("lbl", parent=styles["Normal"], fontSize=9, textColor=colors.grey)
    val = ParagraphStyle("val", parent=styles["Normal"], fontSize=13, leading=15)

    story = [Paragraph(asn_title, h), Spacer(1, 4)]

    head = Table([[
        [Paragraph("CLIENT_CODE", lbl),
         Paragraph(f"<b>{'' if client is None else client}</b>", val), Spacer(1, 3), qr_flowable(client)],
        [Paragraph("DISPLAY_ASN_NUMBER", lbl),
         Paragraph(f"<b>{'' if asn is None else asn}</b>", val), Spacer(1, 3), qr_flowable(asn)],
    ]], colWidths=[93 * mm, 93 * mm])
    head.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("BOX", (0, 0), (-1, -1), 0.5, colors.lightgrey),
        ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.lightgrey),
        ("LEFTPADDING", (0, 0), (-1, -1), 8), ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 8), ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
    ]))
    story += [head, Spacer(1, 10)]

    header = ["DISPLAY_ITEM_NUMBER", "LOT_NUMBER", "QUANTITY\n(Total)", "UOM",
              "S_QTY\n(Total)", "S_UOM", "PACKAGE_TYPE\n(Total)",
              "GROSS_WEIGHT\n(Total)", "NET_WEIGHT\n(Total)"]
    rows = [header]
    for (item, lot), g in groups.items():
        def cell(fld):
            return fmt_num(g["sum"][fld]) if g["has"].get(fld) else "-"
        if g["pkg_numeric"]:
            pkg = fmt_num(g["pkg_num"]) if g["pkg_has"] else "-"
        else:
            pkg = ", ".join(sorted(g["pkg_txt"])) or "-"
        rows.append([
            "" if item is None else str(item),
            "" if lot in (None, "") else str(lot),
            cell("QUANTITY"), "" if g["uom"] is None else str(g["uom"]),
            cell("S_QTY"), "" if g["suom"] is None else str(g["suom"]),
            pkg, cell("GROSS_WEIGHT"), cell("NET_WEIGHT"),
        ])

    t = Table(rows, repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1f3b57")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTSIZE", (0, 0), (-1, 0), 7.5), ("FONTSIZE", (0, 1), (-1, -1), 8),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#b8c4cf")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#eef2f6")]),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (2, 0), (-1, -1), "RIGHT"), ("ALIGN", (0, 0), (-1, 0), "CENTER"),
        ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    story.append(t)
    doc.build(story)
    buf.seek(0)
    return buf.getvalue(), len(groups)


# ---------- HU_ID labels (barcode + QR) ----------

LABEL_DETAIL_FIELDS = ["DISPLAY_ITEM_NUMBER", "DISPLAY_ASN_NUMBER", "LOT_NUMBER",
                       "QUANTITY", "SUPPLIER_DESC"]


@lru_cache(maxsize=8192)
def _qr_png(data):
    qr = qrcode.QRCode(border=1, box_size=10)
    qr.add_data(data)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    b = io.BytesIO()
    img.save(b, format="PNG")
    return b.getvalue()


def _qr_image(data, size_mm=18):
    return Image(io.BytesIO(_qr_png(str(data) if data not in (None, "") else "N/A")),
                 width=size_mm * mm, height=size_mm * mm)


def _barcode(value, width_mm=58, h_mm=11):
    return createBarcodeDrawing("Code128", value=str(value) if value not in (None, "") else "N/A",
                                barHeight=h_mm * mm, width=width_mm * mm, humanReadable=True, fontSize=6)


def make_labels_pdf(records, per_row=2):
    """records: [{'hu':..., 'details':[(label,val),...]}, ...]"""
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, topMargin=8 * mm, bottomMargin=8 * mm,
                            leftMargin=8 * mm, rightMargin=8 * mm, title="HU_ID Labels")
    ss = getSampleStyleSheet()
    big = ParagraphStyle("big", parent=ss["Normal"], fontSize=10, leading=12)
    det = ParagraphStyle("det", parent=ss["Normal"], fontSize=7.5, leading=9)

    col_w = (190.0 / per_row)
    qr_mm = 16 if per_row >= 3 else 18
    bc_mm = col_w - 24

    def cell(rec):
        left = [Paragraph(f"<b>{rec['hu']}</b>", big)]
        for k, v in rec["details"]:
            if v not in (None, ""):
                left.append(Paragraph(f"<font color='#667'>{k}:</font> {v}", det))
        top = Table([[left, _qr_image(rec["hu"], qr_mm)]], colWidths=[(col_w - qr_mm - 6) * mm, qr_mm * mm])
        top.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"),
                                 ("LEFTPADDING", (0, 0), (-1, -1), 2), ("RIGHTPADDING", (0, 0), (-1, -1), 2),
                                 ("TOPPADDING", (0, 0), (-1, -1), 1), ("BOTTOMPADDING", (0, 0), (-1, -1), 1)]))
        inner = Table([[top], [_barcode(rec["hu"], bc_mm)]], colWidths=[col_w * mm])
        inner.setStyle(TableStyle([("ALIGN", (0, 1), (0, 1), "CENTER"),
                                   ("TOPPADDING", (0, 0), (-1, -1), 1), ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
                                   ("LEFTPADDING", (0, 0), (-1, -1), 3)]))
        return inner

    cells = [cell(r) for r in records]
    rows = []
    for i in range(0, len(cells), per_row):
        row = cells[i:i + per_row]
        while len(row) < per_row:
            row.append("")
        rows.append(row)
    grid = Table(rows, colWidths=[col_w * mm] * per_row)
    grid.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#aab")),
                              ("VALIGN", (0, 0), (-1, -1), "TOP"),
                              ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 6)]))
    doc.build([grid])
    buf.seek(0)
    return buf.getvalue()


def build_label_records(ws, hu_col, detail_cols, limit):
    """Exploded sheet එකෙන් label records හදනවා (HU_ID + selected details)."""
    records = []
    truncated = False
    for r in range(2, ws.max_row + 1):
        if is_blank_row(ws, r):
            continue
        if len(records) >= limit:
            truncated = True
            break
        hu = ws.cell(row=r, column=hu_col).value if hu_col else None
        details = []
        for fld, ci in detail_cols:
            if ci:
                details.append((fld, ws.cell(row=r, column=ci).value))
        records.append({"hu": "" if hu is None else str(hu), "details": details})
    return records, truncated


# ---------- UI ----------

st.title("📦 ASN S_QTY Exploder")
st.caption("S_QTY අගයෙන් line ගාන හදනවා · S_QTY 10 → line 10ක් · හැම line එකකම S_QTY = 1 · Exploded Excel + Summary PDF")

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

if len(sqty_sheets) == 1:
    st.caption(f"✓ `S_QTY` column එක හම්බුණේ sheet **{target_sheet}** එකේ (sheet name එක වෙනස් වුණත් auto-detect වෙනවා).")
else:
    st.caption(f"ℹ️ `S_QTY` column එක තියෙන sheets {len(sqty_sheets)}ක්: {', '.join(sqty_sheets)} — එකක් තෝරන්න.")

# ---- HU_ID ----
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

# ---- Summary PDF ----
st.subheader("📄 Summary PDF")
make_pdf = st.checkbox("Summary PDF එකකුත් හදන්න", value=True)
sum_cols = {f: find_col(ws_target, f) for f in SUMMARY_FIELDS}
if make_pdf:
    missing = [f for f in SUMMARY_FIELDS if not sum_cols[f]]
    if missing:
        st.caption("⚠️ නැති columns (skip වෙනවා): " + ", ".join(missing))
    else:
        st.caption("CLIENT_CODE + QR · DISPLAY_ASN_NUMBER + QR · item-wise totals (QUANTITY, S_QTY, weights…)")

# ---- HU_ID Labels (barcode + QR) ----
st.subheader("🏷️ HU_ID Labels (barcode + QR)")
make_labels = st.checkbox("හැම HU_ID එකකටම barcode + QR label PDF එකක් හදන්න", value=False)
label_cfg = {"on": make_labels, "per_row": 2, "limit": 500}
if make_labels:
    if info["hu"] is None:
        st.warning("මේ sheet එකේ `HU_ID` column එකක් නෑ — labels හදන්න බෑ.")
        label_cfg["on"] = False
    else:
        lc1, lc2 = st.columns(2)
        with lc1:
            label_cfg["per_row"] = st.selectbox("Label එකක පේළියකට ගාන", options=[2, 3], index=0)
        with lc2:
            label_cfg["limit"] = st.number_input("උපරිම labels ගාන (speed)", value=500, min_value=1, step=100,
                                                  help="ලොකු ගානකට (>1000) PDF එක හදන්න තත්පර කීපයක් යනවා.")
        st.caption("හැම label එකකම: HU_ID · Code128 barcode · QR code · details (Item, ASN, Lot, Qty…)")

# ---- metrics ----
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

    base = uploaded.name.rsplit(".", 1)[0]

    pdf_bytes, ngrp, pdf_err = None, 0, None
    if make_pdf:
        try:
            pdf_bytes, ngrp = make_summary_pdf(ws_target, sum_cols)
        except Exception as e:
            pdf_err = str(e)

    labels_bytes, n_labels, labels_trunc, labels_err = None, 0, False, None
    if label_cfg["on"] and info["hu"]:
        try:
            detail_cols = [(f, find_col(ws_target, f)) for f in LABEL_DETAIL_FIELDS]
            recs, labels_trunc = build_label_records(ws_target, info["hu"], detail_cols, int(label_cfg["limit"]))
            n_labels = len(recs)
            labels_bytes = make_labels_pdf(recs, per_row=label_cfg["per_row"])
        except Exception as e:
            labels_err = str(e)

    headers = [ws_target.cell(row=1, column=c).value for c in range(1, ws_target.max_column + 1)]
    preview = []
    for r in range(2, min(ws_target.max_row, 16) + 1):
        preview.append([ws_target.cell(row=r, column=c).value for c in range(1, ws_target.max_column + 1)])

    # download buttons click එකකින් rerun වුණත් නැති නොවෙන්න session_state එකේ තියාගන්නවා
    st.session_state["result"] = {
        "base": base, "n_src": n_src, "n_out": n_out, "sheet": target_sheet,
        "excel": out_buf.getvalue(), "pdf": pdf_bytes, "ngrp": ngrp, "pdf_err": pdf_err,
        "labels": labels_bytes, "n_labels": n_labels, "labels_trunc": labels_trunc, "labels_err": labels_err,
        "headers": headers, "preview": preview,
    }

# ---- results (session_state එකෙන් — rerun වුණත් download buttons නැති වෙන්නේ නෑ) ----
res = st.session_state.get("result")
if res:
    st.success(f"✅ සාර්ථකයි · {res['n_src']} rows → {res['n_out']} rows ({res['sheet']})")

    d1, d2, d3 = st.columns(3)
    with d1:
        st.download_button("⬇️ Exploded Excel", data=res["excel"],
                           file_name=f"{res['base']}_EXPLODED.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True, key="dl_excel")
    with d2:
        if res["pdf"] is not None:
            st.download_button(f"⬇️ Summary PDF ({res['ngrp']} items)", data=res["pdf"],
                               file_name=f"{res['base']}_SUMMARY.pdf", mime="application/pdf",
                               use_container_width=True, key="dl_pdf")
        elif res["pdf_err"]:
            st.error(f"PDF error: {res['pdf_err']}")
    with d3:
        if res.get("labels") is not None:
            st.download_button(f"⬇️ HU_ID Labels ({res['n_labels']})", data=res["labels"],
                               file_name=f"{res['base']}_LABELS.pdf", mime="application/pdf",
                               use_container_width=True, key="dl_labels")
        elif res.get("labels_err"):
            st.error(f"Labels error: {res['labels_err']}")
    if res.get("labels_trunc"):
        st.caption(f"ℹ️ Labels {res['n_labels']}කට limit කළා. ඔක්කොම ඕන නම් 'උපරිම labels ගාන' වැඩි කරන්න.")

    st.caption("Output preview (මුල් rows 15)")
    st.dataframe({(hh or f"col{i}"): [row[i] for row in res["preview"]]
                  for i, hh in enumerate(res["headers"])},
                 use_container_width=True, height=360)
