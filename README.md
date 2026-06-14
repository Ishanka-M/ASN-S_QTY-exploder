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
