# 📦 ASN S_QTY Exploder

S_QTY column එකේ අගයෙන් line ගාන හදන Streamlit tool එකක්.
**S_QTY = 10 → line 10ක්, හැම line එකකම S_QTY = 1, අනිත් columns details එහෙමම.**

Support කරන formats:
- **GENERAL Format** — `ASN Master` + `ASN DETAIL` (explode වෙන්නේ `ASN DETAIL`)
- **ATTRIBUTE EXTENTD format** — `Physical ASN`

S_QTY column එක තියෙන sheet එක auto-detect වෙනවා. **Sheet name එක වෙනස් වුණත් කමක් නෑ** — format එක එකම නම්, name එකෙන් නෙවෙයි S_QTY column එකෙන් sheet එක හොයාගන්නවා. Master වගේ අනිත් sheets නොවෙනස්ව output එකේ තියෙනවා.

## Local run (Windows)
`run.bat` double-click කරන්න. (Python 3.9+ ඕන.)

## Local run (manual)
```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

## GitHub + Streamlit Cloud deploy
1. මේ folder එක GitHub repo එකකට push කරන්න (streamlit_app.py + requirements.txt අනිවාර්යයි).
2. https://share.streamlit.io → New app → repo එක select → main file: `streamlit_app.py` → Deploy.

## HU_ID Labels — barcode + QR (අලුත්)
හැම HU_ID එකකටම (carton/line එකකට) **Code128 barcode + QR code + details** තියෙන label PDF එකක්:
- HU_ID, DISPLAY_ITEM_NUMBER, DISPLAY_ASN_NUMBER, LOT_NUMBER, QUANTITY, SUPPLIER_DESC
- පේළියකට labels 2 හරි 3 හරි තෝරගන්න පුළුවන්
- Speed එකට උපරිම labels ගානක් (default 500) set කරන්න පුළුවන් — ලොකු ගානකට (>1000) තත්පර කීපයක් යනවා

## Multiple users + speed
- Streamlit එක default එකෙන්ම **session per-user** — multiple users එකවර වැඩ කරනවා (state share වෙන්නේ නෑ).
- Explode logic එක `iter_rows` + `append` වලින් optimize කරලා (cell-by-cell නෙවෙයි) — 10,000 rows ~1s.
- QR images `lru_cache` කරලා — එකම value නැවත නැවත හදන්නේ නෑ.
- සැබෑ heavy concurrent load එකකට Streamlit Cloud resources වැඩි කරගන්න (හෝ dedicated server එකක).

## Summary PDF — per-ASN pages (update)
- **Page 1: ASN Overview** — සියලුම DISPLAY_ASN_NUMBER එකට: DISPLAY_ITEM_NUMBER, LOT_NUMBER, QUANTITY (Total), UOM (එකම ASN එකේ rows merge වෙනවා).
- **ඊට පස්සේ DISPLAY_ASN_NUMBER එකකට වෙන වෙනම page එකක්** — CLIENT_CODE + QR, DISPLAY_ASN_NUMBER + QR, සහ ඒ ASN එකේ item-wise totals (QUANTITY/S_QTY/PACKAGE_TYPE/GROSS/NET).

## Summary PDF (අලුත්)
Exploded data එකෙන් **summary PDF** එකක් හැදෙනවා:
- **CLIENT_CODE + QR Code**
- **DISPLAY_ASN_NUMBER + QR Code**
- Item-wise totals table: DISPLAY_ITEM_NUMBER, LOT_NUMBER, QUANTITY (Total), UOM, S_QTY (Total), S_UOM, PACKAGE_TYPE (Total), GROSS_WEIGHT (Total), NET_WEIGHT (Total)

Totals ගණනය වෙන්නේ explode වුණාට පස්සේ data එකෙන් — ඒ නිසා S_QTY (Total) = ඒ item එකේ මුළු carton/line ගාන. (numeric columns sum, text columns value පෙන්නනවා, value නැති numeric columns `-`.)

## HU_ID generate (අලුත්)
`HU_ID` column එක හදන විදි 4ක්:
- **Keep original** — තියෙන HU_ID එකම තියනවා
- **None** — HU_ID හිස්ව (blank)
- **Letters + Number** — `OKARW` + 142968, 142969… (line ගානට +1)
- **DISPLAY_ITEM_NUMBER + Number** — row එකේ item number + 001, 002… (separator + zero-pad options)

App එකේ live preview එකකින් generate වෙන HU_ID format එක කලින්ම බලාගන්න පුළුවන්.

## Logic
1. File එක upload කරනවා.
2. S_QTY තියෙන sheet එක තෝරනවා (auto).
3. හැම data row එකක්ම S_QTY පාරක් copy වෙනවා, හැම copy එකකම S_QTY = 1.
4. `ASN_LINE_NUMBER` 1..N විදිහට renumber (toggle එකෙන් off කරන්න පුළුවන්).
5. Output `.xlsx` එක download වෙනවා — headers සහ cell styles රැකෙනවා.

## Professional UI + animations
- Navy + amber theme, codes වලට monospace (JetBrains Mono) accent, Inter body font.
- Animated hero header (amber sheen), fade-in sections, hover lift buttons/metrics, card-style metrics.
- `prefers-reduced-motion` respect කරනවා. Theme `.streamlit/config.toml` එකෙන් set වෙනවා.
