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
st.title("📄 Hyrbdi Asset Custody Letter Generator")

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
}

# -----------------------------
# Utilities
# -----------------------------

def norm(s: str) -> str:
    """Trim + collapse inner whitespace."""
    return re.sub(r"\s+", " ", s.strip()) if isinstance(s, str) else ""

def extract_placeholders(doc: Document) -> set:
    """Scan a Word document (paragraphs + tables) for all {{PLACEHOLDER}} tags and return a set of keys (original case)."""
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
    """Scan all word/*.xml parts (including headers/footers/text boxes/shapes) for placeholders."""
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
    """Low-level XML replace across all word/*.xml parts to catch text in shapes, headers, and footers."""
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

# ---------- NEW: formatting helpers ----------

def _copy_run_format(src: Run, dst: Run) -> None:
    """Copy key inline styles from src to dst."""
    dst.bold = src.bold
    dst.italic = src.italic
    dst.underline = src.underline
    # font attributes
    if src.font is not None:
        if src.font.size:
            dst.font.size = src.font.size
        if src.font.name:
            dst.font.name = src.font.name
        # color
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
    """
    if not paragraph.runs:
        return

    # Build concatenated text with index map: [(run_idx, start_pos, end_pos)]
    runs = paragraph.runs
    concat = ""
    spans = []
    for idx, r in enumerate(runs):
        start = len(concat)
        concat += r.text
        spans.append((idx, start, len(concat)))

    # Make a quick lookup from absolute position to run index
    def _find_span(pos: int):
        for idx, s, e in spans:
            if s <= pos < e:
                return idx, s, e
        # end boundary case
        return spans[-1][0], spans[-1][1], spans[-1][2]

    changed = True
    # Repeat until no more placeholders remain in this paragraph
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

            run_start_idx, run_start_s, run_start_e = _find_span(start_pos)
            run_end_idx, run_end_s, run_end_e = _find_span(end_pos - 1)

            before_text = runs[run_start_idx].text[: start_pos - run_start_s]
            after_text = runs[run_end_idx].text[end_pos - run_end_s :]

            # Replacement value
            value = data_map.get(up, "")

            # Put combined text into the starting run, preserve its formatting
            runs[run_start_idx].text = before_text + value

            # Clear fully covered middle runs
            for i in range(run_start_idx + 1, run_end_idx):
                runs[i].text = ""

            # Put the trailing text into the end run
            if run_end_idx != run_start_idx:
                runs[run_end_idx].text = after_text
                # Ensure formatting of the replacement text matches the start run formatting.
                # The replacement sits inside run_start_idx, which already holds the right style.
            # If same run, the style is already preserved.

            # Copy style from the run where the placeholder started into run_start_idx explicitly
            _copy_run_format(runs[run_start_idx], runs[run_start_idx])

            # Rebuild concat and spans for the next loop
            concat = ""
            spans = []
            for idx, r in enumerate(runs):
                s = len(concat)
                concat += r.text
                spans.append((idx, s, len(concat)))
            # break to restart placeholder search from first placeholder each loop
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

        # Normalize headers (strip + UPPER to ensure consistent matching)
        df.columns = [col.strip().upper() for col in df.columns]

        st.session_state["generated_docs"] = []  # Clear previous outputs
        ok_count = 0

        for index, row in df.iterrows():
            try:
                # Normalize row values; blanks become empty strings and keys forced to UPPER
                data = {str(k).strip().upper(): ("" if pd.isna(v) else norm(str(v))) for k, v in row.items()}

                # --- Amount formatting fix (exactly 2 decimals, no thousands comma) ---
                if "AMOUNT" in data and data["AMOUNT"]:
                    try:
                        data["AMOUNT"] = f"{float(data['AMOUNT']):.2f}"
                    except ValueError:
                        pass
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
                # Pass 1: Run-level replacement (keeps formatting when fully inside a run)
                # -----------------------------
                for para in _walk_block_items(doc):
                    for run in para.runs:
                        text = run.text
                        if "{{" not in text:
                            continue
                        for orig, up in placeholders_upper.items():
                            ph = f"{{{{{orig}}}}}"
                            if ph in text:
                                value = data.get(up, "")
                                parts = value.split("\n")
                                if len(parts) > 1:
                                    text = text.replace(ph, parts[0])
                                    run.text = text
                                    for line in parts[1:]:
                                        run.add_break()
                                        run.add_text(line)
                                else:
                                    run.text = text.replace(ph, value)

                # -----------------------------
                # Pass 2: Cross-run replacement with formatting preservation
                #         (replaces placeholders split across runs)
                # -----------------------------
                for para in _walk_block_items(doc):
                    if "{{" in para.text:
                        _replace_placeholders_preserving_runs(para, placeholders_upper, data)

                # Save docx from python-docx first
                inter_buffer = BytesIO()
                doc.save(inter_buffer)
                inter_bytes = inter_buffer.getvalue()

                # -----------------------------
                # Pass 3: Low-level XML replacement
                # (handles placeholders in text boxes, headers/footers, etc.)
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
