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
    """Trim and collapse inner spaces. Newlines removed."""
    return re.sub(r"\s+", " ", s.strip()) if isinstance(s, str) else ""

def norm_multiline(s: str) -> str:
    """Preserve intended line breaks from Excel cells."""
    if s is None:
        return ""
    s = str(s)
    s = s.replace("\r\n", "\n").replace("\r", "\n").replace("\\n", "\n")
    # Trim spaces around each line but keep \n
    lines = [ln.strip() for ln in s.split("\n")]
    return "\n".join(lines).strip()

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

def _replace_placeholders_preserving_runs(paragraph: Paragraph, placeholders_upper: dict, data_map: dict) -> None:
    """
    Replace {{PLACEHOLDER}} even when split across runs.
    Keep formatting by applying the formatting of the run where the placeholder starts.
    Supports multiline values by inserting line breaks.
    """
    if not paragraph.runs:
        return

    runs = paragraph.runs
    concat = ""
    spans = []
    for idx, r in enumerate(runs):
        start = len(concat)
        concat += r.text
        spans.append((idx, start, len(concat)))

    def _find_span(pos: int):
        for idx, s, e in spans:
            if s <= pos < e:
                return idx, s, e
        return spans[-1][0], spans[-1][1], spans[-1][2]

    changed = True
    while changed:
        changed = False
        for orig, up in placeholders_upper.items():
            needle = "{{" + orig + "}}"
            pos = concat.find(needle)
            if pos == -1:
                continue

            changed = True
            start_pos = pos
            end_pos = pos + len(needle)

            run_start_idx, run_start_s, _ = _find_span(start_pos)
            run_end_idx, run_end_s, _ = _find_span(end_pos - 1)

            before_text = runs[run_start_idx].text[: start_pos - run_start_s]
            after_text = runs[run_end_idx].text[end_pos - run_end_s :]

            value = str(data_map.get(up, ""))

            # Prepare first run text
            runs[run_start_idx].text = before_text

            # Insert replacement with possible line breaks, preserving style
            first = True
            for i, line in enumerate(value.split("\n")):
                if first:
                    # append to the start run
                    runs[run_start_idx].text += line
                    first = False
                else:
                    r = paragraph.add_run()
                    _copy_run_format(runs[run_start_idx], r)
                    r.add_break()
                    r.add_text(line)

            # Clear fully covered middle runs
            for i in range(run_start_idx + 1, run_end_idx):
                runs[i].text = ""

            # Put trailing text back into the end run
            if run_end_idx != run_start_idx:
                runs[run_end_idx].text = after_text

            # Rebuild concat map
            concat = ""
            spans = []
            for idx, r in enumerate(paragraph.runs):
                s = len(concat)
                concat += r.text
                spans.append((idx, s, len(concat)))
            break

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
    # remove spaces, commas, currency and letters, keep digits, dot, minus
    s = s.replace(",", "")
    s = re.sub(r"[^0-9.\-]", "", s)
    if s == "" or s == "." or s == "-" or s == "-.":
        return ""
    try:
        return f"{float(s):.2f}"
    except ValueError:
        return ""

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
                # Build data from row as raw strings first
                data = {}
                for k, v in row.items():
                    key = str(k).strip().upper()
                    val = "" if pd.isna(v) else str(v)
                    data[key] = val

                # Preserve line breaks for any ADDRESS-like fields
                for k in list(data.keys()):
                    if "ADDRESS" in k:
                        data[k] = norm_multiline(data[k])
                    else:
                        data[k] = norm(data[k])

                # --- Amount formatting fix (2 decimals, sanitize inputs) ---
                for field in ["AMOUNT", "DIVIDEND", "ACCUMULATED", "TRUST_CAPITAL"]:
                    if field in data and data[field]:
                        data[field] = _to_2dp(data[field])
                # --- End of amount fix ---

                letter_type = norm(data.get("LETTER_TYPE", ""))
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
                # Pass 1: Simple in-run replacement with newline handling
                # -----------------------------
                for para in _walk_block_items(doc):
                    for run in list(para.runs):
                        text = run.text
                        if "{{" not in text:
                            continue
                        for orig, up in placeholders_upper.items():
                            ph = f"{{{{{orig}}}}}"
                            if ph in text:
                                value = str(data.get(up, ""))
                                if "\n" in value:
                                    parts = value.split("\n")
                                    # set first part on current run
                                    run.text = text.replace(ph, parts[0])
                                    # append subsequent lines with breaks using new runs that copy the style
                                    for line in parts[1:]:
                                        r = para.add_run()
                                        _copy_run_format(run, r)
                                        r.add_break()
                                        r.add_text(line)
                                else:
                                    run.text = text.replace(ph, value)

                # -----------------------------
                # Pass 2: Cross-run replacement preserving formatting
                # -----------------------------
                for para in _walk_block_items(doc):
                    if "{{" in para.text:
                        _replace_placeholders_preserving_runs(para, placeholders_upper, data)

                # Save docx first
                inter_buffer = BytesIO()
                doc.save(inter_buffer)
                inter_bytes = inter_buffer.getvalue()

                # -----------------------------
                # Pass 3: Low-level XML replacement for headers/footers/shapes
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
