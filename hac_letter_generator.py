import streamlit as st
from docx import Document
from docx.text.paragraph import Paragraph
from docx.table import _Cell, Table
from docx.text.run import Run
from docx.shared import RGBColor
import pandas as pd
from io import BytesIO
import os
import re
import zipfile

st.set_page_config(page_title="HAC Letter Generator", layout="centered")
st.title("📄 Hybrid Asset Custody Letter Generator")

# -----------------------------
# Template registry
# -----------------------------
template_options = {
    "Letter of Appointment": "templates/Letter of Appointment.docx",
    "Letter of Instruction": "templates/Letter of Instruction.docx",
    "Letter of Set-Off": "templates/Letter of Set-Off.docx",
    "Letter of Affirmation (Quarterly)": "templates/Letter of Affirmation - Quarterly.docx",
    "Letter of Affirmation (Yearly)": "templates/Letter of Affirmation - Yearly.docx",
    "Letter of Welcome": "templates/Letter of Welcome.docx",
    "Letter of Agreement": "templates/Letter of Agreement.docx",
    "Letter of Acknowledgement": "templates/Letter of Acknowledgement.docx",
    "Letter of Dividend": "templates/Letter of Dividend.docx",
    "Letter of Lucky Draw": "templates/Letter of Lucky Draw.docx",
    "Trust Deed": "templates/Trust Deed.docx",
    "Commission Letter (7th)": "templates/Commission Letter (7th).docx",
    "HAC AGREEMENT": "templates/HAC Agreement.docx",
    "Dividend Letter": "templates/Letter of Dividend Payout.docx",
}

# -----------------------------
# Utilities
# -----------------------------

def norm(s: str) -> str:
    """Trim and collapse inner spaces into one space."""
    return re.sub(r"\s+", " ", s.strip()) if isinstance(s, str) else ""

def norm_multiline(s: str) -> str:
    """Preserve intended line breaks from Excel cells."""
    if s is None:
        return ""
    s = str(s)
    s = s.replace("\r\n", "\n").replace("\r", "\n").replace("\\n", "\n")
    lines = [ln.strip() for ln in s.split("\n")]
    return "\n".join(lines).strip()

def auto_multiline_address(s: str) -> str:
    """
    If address is a single line, split on commas into neat lines.
    Keep existing line breaks if already present.
    """
    if s is None:
        return ""
    s = str(s)
    if "\n" in s:
        s = s.replace("\r\n", "\n").replace("\r", "\n")
        return "\n".join(part.strip() for part in s.split("\n") if part.strip())
    parts = [p.strip() for p in s.split(",") if p.strip()]
    return "\n".join(parts)

def extract_placeholders(doc: Document) -> set:
    """Find all {{PLACEHOLDER}} in paragraphs and tables."""
    placeholder_pattern = re.compile(r"{{(.*?)}}")
    found = set()
    for para in doc.paragraphs:
        found.update(placeholder_pattern.findall(para.text))
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    found.update(placeholder_pattern.findall(para.text))
    return found

def extract_placeholders_from_docx_bytes(docx_bytes: bytes) -> set:
    """Scan all word/*.xml parts for placeholders, including headers/footers/shapes."""
    pattern = re.compile(r"{{(.*?)}}")
    found = set()
    with BytesIO(docx_bytes) as in_mem:
        with zipfile.ZipFile(in_mem, 'r') as zin:
            for item in zin.infolist():
                if item.filename.startswith('word/') and item.filename.endswith('.xml'):
                    try:
                        xml_text = zin.read(item.filename).decode('utf-8')
                    except UnicodeDecodeError:
                        continue
                    found.update(pattern.findall(xml_text))
    return found

def replace_in_docx_bytes(docx_bytes: bytes, mapping: dict, placeholders_upper_map: dict) -> bytes:
    """Low-level XML replace across word/*.xml parts to catch text in shapes, headers, and footers."""
    in_mem = BytesIO(docx_bytes)
    out_mem = BytesIO()
    with zipfile.ZipFile(in_mem, 'r') as zin, zipfile.ZipFile(out_mem, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data_bytes = zin.read(item.filename)
            if item.filename.startswith('word/') and item.filename.endswith('.xml'):
                try:
                    xml_text = data_bytes.decode('utf-8')
                except UnicodeDecodeError:
                    zout.writestr(item, data_bytes)
                    continue
                for orig, up in placeholders_upper_map.items():
                    ph = '{{' + orig + '}}'
                    val = mapping.get(up, '') or ''
                    if ph in xml_text:
                        xml_text = xml_text.replace("\t" + ph, ph).replace(ph + "\t", ph)
                        xml_text = xml_text.replace(ph, str(val))
                data_bytes = xml_text.encode('utf-8')
            zout.writestr(item, data_bytes)
    out_mem.seek(0)
    return out_mem.getvalue()

# ---------- Formatting helpers ----------

def _copy_run_format(src: Run, dst: Run) -> None:
    """Copy key inline styles from src to dst."""
    dst.bold = src.bold
    dst.italic = src.italic
    dst.underline = src.underline
    if src.font is not None:
        if src.font.size:
            dst.font.size = src.font.size
        if src.font.name:
            dst.font.name = src.font.name
        try:
            rgb = src.font.color.rgb
            if isinstance(rgb, RGBColor):
                dst.font.color.rgb = rgb
        except Exception:
            pass

def _walk_block_items(doc: Document):
    """Yield all paragraphs from the document, including those in tables."""
    for p in doc.paragraphs:
        yield p
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    yield p

def _to_2dp(value: str) -> str:
    """Sanitize a numeric-like string and format to two decimals."""
    s = str(value).strip()
    s = s.replace(",", "")
    s = re.sub(r"[^0-9.\-]", "", s)
    if s in {"", ".", "-", "-."}:
        return ""
    try:
        return f"{float(s):.2f}"
    except ValueError:
        return ""

# -------- Paragraph-level multiline replacer (removes tabs) --------

def _replace_multiline_paragraph(paragraph: Paragraph, placeholders_upper: dict, data_map: dict) -> bool:
    """
    If paragraph contains a placeholder whose value has newline(s), rebuild the entire paragraph:
    - remove tabs around the placeholder
    - write clean runs and insert hard line breaks
    Returns True if a replacement was made.
    """
    text = paragraph.text
    did = False
    for orig, up in placeholders_upper.items():
        ph = f"{{{{{orig}}}}}"
        value = str(data_map.get(up, ""))
        if ph in text and ("\n" in value):
            did = True
            text = text.replace("\t" + ph, ph).replace(ph + "\t", ph)
            text = text.replace("\t", " ")
            before, after = text.split(ph, 1)

            for r in list(paragraph.runs):
                r.text = ""
            paragraph.text = ""

            if before:
                paragraph.add_run(before)

            lines = value.split("\n")
            if lines:
                paragraph.add_run(lines[0])
                for line in lines[1:]:
                    paragraph.add_run().add_break()
                    paragraph.add_run(line)

            if after:
                paragraph.add_run(after)
            break
    return did

def _replace_simple_inrun(paragraph: Paragraph, placeholders_upper: dict, data_map: dict) -> None:
    """Simple in-run replacements for placeholders whose values are single-line."""
    for run in list(paragraph.runs):
        txt = run.text
        if "{{" not in txt:
            continue
        for orig, up in placeholders_upper.items():
            ph = f"{{{{{orig}}}}}"
            if ph in txt:
                val = str(data_map.get(up, ""))
                if "\n" not in val:
                    run.text = txt.replace(ph, val)

# -----------------------------
# UI — Excel upload
# -----------------------------
uploaded_excel = st.file_uploader("📈 Upload Master Excel File (All Letters)", type=["xlsx"], key="excel_upload")

# Session state for generated docs
if "generated_docs" not in st.session_state:
    st.session_state["generated_docs"] = []

if uploaded_excel:
    st.success("✅ File uploaded. Click the button to generate all letters.")
    debug = st.checkbox("🔍 Debug mode (show placeholders & mapping)")

    if st.button("Generate Letters"):
        try:
            df = pd.read_excel(uploaded_excel)
        except Exception as e:
            st.error(f"Failed to read Excel: {e}")
            st.stop()

        # Normalize headers to UPPER
        df.columns = [col.strip().upper() for col in df.columns]

        st.session_state["generated_docs"] = []
        ok_count = 0

        for index, row in df.iterrows():
            try:
                # Build raw data from row (no normalization yet)
                data = {}
                for k, v in row.items():
                    key = str(k).strip().upper()
                    val = "" if pd.isna(v) else str(v)
                    data[key] = val

                # Peek the letter type early for conditional address logic
                letter_type_raw = str(data.get("LETTER_TYPE", "")).strip()
                letter_type_upper = letter_type_raw.upper()

                # Normalize values, with conditional address handling
                for k in list(data.keys()):
                    if "ADDRESS" in k:
                        if "DIVIDEND" in letter_type_upper:
                            data[k] = auto_multiline_address(data[k])
                        else:
                            data[k] = norm_multiline(data[k])
                    else:
                        data[k] = norm(data[k])

                # --- Amount formatting fix (2 decimals, sanitize inputs) ---
                for field in ["AMOUNT", "DIVIDEND", "ACCUMULATED", "TRUST_CAPITAL"]:
                    if field in data and data[field]:
                        data[field] = _to_2dp(data[field])
                # --- End of amount fix ---

                # Final normalized types for downstream logic
                letter_type = norm(letter_type_raw)
                payout_type = norm(data.get("PAYOUT_TYPE", ""))

                if not letter_type:
                    st.warning(f"Row {index+1}: LETTER_TYPE is empty. Skipping.")
                    continue

                # Build template key safely
                template_key = f"{letter_type} ({payout_type})" if letter_type == "Letter of Affirmation" else letter_type
                template_path = template_options.get(template_key)

                if not template_path:
                    st.warning(f"Row {index+1}: Template key not found → '{template_key}'. Check spelling/casing/spaces.")
                    continue
                if not os.path.exists(template_path):
                    st.warning(f"Row {index+1}: Template path does not exist → {template_path}")
                    continue

                # Load Word template
                with open(template_path, "rb") as f:
                    template_bytes = f.read()
                doc = Document(BytesIO(template_bytes))

                # Detect placeholders in this template (original case)
                placeholders_orig = extract_placeholders(doc)
                placeholders_upper = {p: p.upper() for p in placeholders_orig}

                # Ensure every placeholder exists in data (default to empty string)
                for orig, up in placeholders_upper.items():
                    if up not in data:
                        data[up] = ""

                if debug:
                    xml_ph = extract_placeholders_from_docx_bytes(template_bytes)
                    sample = ", ".join(sorted(list(xml_ph))[:12])
                    if len(xml_ph) > 12:
                        sample += ", …"
                    st.info(
                        f"Row {index+1}: template='{template_key}'\n"
                        f"Path='{template_path}'\n"
                        f"Placeholders (docx XML)={len(xml_ph)} → {sample}"
                    )

                # -----------------------------
                # Pass A: Paragraph-level multiline replacer (kills tabs)
                # -----------------------------
                for para in _walk_block_items(doc):
                    if "{{" in para.text:
                        _replace_multiline_paragraph(para, placeholders_upper, data)

                # -----------------------------
                # Pass B: Simple single-line in-run replacements
                # -----------------------------
                for para in _walk_block_items(doc):
                    if "{{" in para.text:
                        _replace_simple_inrun(para, placeholders_upper, data)

                # Save docx first
                inter_buffer = BytesIO()
                doc.save(inter_buffer)
                inter_bytes = inter_buffer.getvalue()

                # -----------------------------
                # Pass C: Low-level XML replacement for headers/footers/shapes
                # -----------------------------
                final_bytes = replace_in_docx_bytes(inter_bytes, data, placeholders_upper)

                name_part = data.get("NAME", "Client")
                filename = f"{name_part} - {template_key}.docx"

                st.session_state["generated_docs"].append({
                    "filename": filename,
                    "buffer": final_bytes,
                })
                ok_count += 1

            except Exception as e:
                st.error(f"Row {index+1} failed: {e}")

        if ok_count == 0:
            st.warning("No files were generated. Check the warnings/errors above.")
        else:
            st.success(f"✅ Generated {ok_count} file(s). See download buttons below.")

# -----------------------------
# Downloads & housekeeping
# -----------------------------
if st.session_state.get("generated_docs"):
    for i, file in enumerate(st.session_state["generated_docs"]):
        st.download_button(
            label=f"🗕️ Download: {file['filename']}",
            data=file["buffer"],
            file_name=file["filename"],
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key=f"download_{i}",
        )

    if st.button("🗑️ Clear All"):
        st.session_state["generated_docs"] = []
        st.rerun()
else:
    st.warning("📌 Please upload the Excel file to begin.")
