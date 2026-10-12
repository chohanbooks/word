import io
import copy
import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

st.set_page_config(page_title="Word Formatting & Layout Replication Engine", page_icon="📄", layout="wide")

st.title("📄 Word Formatting & Layout Replication Engine")
st.write(
    "Upload a correctly formatted reference DOCX and a target DOCX. "
    "The app applies common styles and layout settings from the reference to the target."
)

st.warning(
    "Important: this is a rule-based first version, not a guaranteed pixel-perfect AI formatter. "
    "It matches body paragraphs mainly by position and can need review when document structures differ."
)

with st.expander("Formatting policy", expanded=True):
    st.markdown(
        """
        - **Document A** is used as the visual formatting reference.
        - **Document B** remains the source of the target document's content.
        - The app aims to preserve target text and data while copying common styles, paragraph formatting,
          page setup, and basic table formatting.
        - Header/footer layout is copied from the reference. Because headers and footers may contain
          document-specific details, choose the header/footer policy below.
        """
    )

reference_file = st.file_uploader(
    "1. Document A — correctly formatted reference (.docx)",
    type=["docx"],
    key="reference_doc",
)
target_file = st.file_uploader(
    "2. Document B — document to format (.docx)",
    type=["docx"],
    key="target_doc",
)

st.subheader("Header and footer policy")
header_footer_policy = st.radio(
    "How should header/footer content be handled?",
    [
        "Copy layout and formatting; preserve target-specific text where practical",
        "Copy the complete header/footer from the reference",
        "Do not modify headers or footers",
    ],
    index=0,
    help=(
        "The first option is safer when the documents have different company names, document codes, "
        "revision numbers, or dates. The second option replaces header/footer content with the reference content."
    ),
)

col1, col2 = st.columns(2)
with col1:
    apply_tables = st.checkbox("Format tables where structures correspond", value=True)
    copy_page_setup = st.checkbox("Copy page size, orientation, margins and header/footer distances", value=True)
with col2:
    copy_styles = st.checkbox("Copy common named style definitions", value=True)
    preserve_target_content = st.checkbox("Preserve target document content", value=True, disabled=True)

def copy_font(source_font, target_font):
    """Copy common font properties; do not alter text."""
    for attr in (
        "name", "size", "bold", "italic", "underline", "all_caps",
        "small_caps", "strike", "double_strike", "subscript", "superscript"
    ):
        try:
            value = getattr(source_font, attr)
            if value is not None:
                setattr(target_font, attr, value)
        except Exception:
            pass
    try:
        if source_font.color and source_font.color.rgb:
            target_font.color.rgb = source_font.color.rgb
    except Exception:
        pass

def copy_run_format(source_run, target_run):
    try:
        if source_run.style is not None:
            target_run.style = source_run.style
    except Exception:
        pass
    copy_font(source_run.font, target_run.font)

def copy_paragraph_format(source_para, target_para):
    try:
        if source_para.style is not None:
            target_para.style = source_para.style
    except Exception:
        pass

    sf = source_para.paragraph_format
    tf = target_para.paragraph_format
    for attr in (
        "alignment", "left_indent", "right_indent", "first_line_indent",
        "space_before", "space_after", "line_spacing", "keep_together",
        "keep_with_next", "page_break_before", "widow_control"
    ):
        try:
            value = getattr(sf, attr)
            if value is not None:
                setattr(tf, attr, value)
        except Exception:
            pass

    if source_para.runs and target_para.runs:
        # Use one-to-one matching when run counts align. Otherwise apply the
        # first reference run's character style as a fallback.
        if len(source_para.runs) == len(target_para.runs):
            for sr, tr in zip(source_para.runs, target_para.runs):
                copy_run_format(sr, tr)
        else:
            for tr in target_para.runs:
                copy_run_format(source_para.runs[0], tr)

def copy_style_definitions(reference_doc, target_doc):
    count = 0
    for source_style in reference_doc.styles:
        try:
            target_style = target_doc.styles[source_style.name]
        except Exception:
            continue
        try:
            copy_font(source_style.font, target_style.font)
        except Exception:
            pass
        try:
            sf, tf = source_style.paragraph_format, target_style.paragraph_format
            for attr in (
                "alignment", "left_indent", "right_indent", "first_line_indent",
                "space_before", "space_after", "line_spacing", "keep_with_next",
                "page_break_before", "widow_control"
            ):
                value = getattr(sf, attr)
                if value is not None:
                    setattr(tf, attr, value)
        except Exception:
            pass
        count += 1
    return count

def copy_page_settings(reference_doc, target_doc):
    if not reference_doc.sections:
        return 0
    count = 0
    for index, target_section in enumerate(target_doc.sections):
        source_section = reference_doc.sections[min(index, len(reference_doc.sections) - 1)]
        for attr in (
            "page_width", "page_height", "top_margin", "bottom_margin",
            "left_margin", "right_margin", "gutter", "header_distance",
            "footer_distance"
        ):
            try:
                value = getattr(source_section, attr)
                if value is not None:
                    setattr(target_section, attr, value)
            except Exception:
                pass
        try:
            target_section.orientation = source_section.orientation
        except Exception:
            pass
        try:
            target_section.different_first_page_header_footer = (
                source_section.different_first_page_header_footer
            )
        except Exception:
            pass
        count += 1
    return count

def copy_paragraphs_by_position(source_paragraphs, target_paragraphs):
    if not source_paragraphs:
        return 0
    for i, target_para in enumerate(target_paragraphs):
        source_para = source_paragraphs[min(i, len(source_paragraphs) - 1)]
        copy_paragraph_format(source_para, target_para)
    return len(target_paragraphs)

def copy_table_formatting(reference_doc, target_doc):
    processed = 0
    if not reference_doc.tables:
        return processed
    for ti, target_table in enumerate(target_doc.tables):
        source_table = reference_doc.tables[min(ti, len(reference_doc.tables) - 1)]
        try:
            target_table.alignment = source_table.alignment
            target_table.autofit = source_table.autofit
        except Exception:
            pass
        for ri, target_row in enumerate(target_table.rows):
            if not source_table.rows:
                continue
            source_row = source_table.rows[min(ri, len(source_table.rows) - 1)]
            for ci, target_cell in enumerate(target_row.cells):
                if not source_row.cells:
                    continue
                source_cell = source_row.cells[min(ci, len(source_row.cells) - 1)]
                for pi, target_para in enumerate(target_cell.paragraphs):
                    if source_cell.paragraphs:
                        source_para = source_cell.paragraphs[min(pi, len(source_cell.paragraphs) - 1)]
                        copy_paragraph_format(source_para, target_para)
                        processed += 1
    return processed

def copy_header_footer_content(source_part, target_part):
    """
    Copy the content body XML of a header/footer part. This preserves most
    paragraph/table/image-reference XML, but relationships to embedded images
    may not be portable between separate documents without relationship remapping.
    """
    source_element = source_part._element
    target_element = target_part._element
    for child in list(target_element):
        # Keep section properties if present; headers/footers usually contain p/tbl children.
        if child.tag.endswith("}p") or child.tag.endswith("}tbl"):
            target_element.remove(child)
    for child in list(source_element):
        if child.tag.endswith("}p") or child.tag.endswith("}tbl"):
            target_element.append(copy.deepcopy(child))

def copy_header_footer(reference_doc, target_doc, policy):
    if policy == "Do not modify headers or footers":
        return 0, "Headers and footers were left unchanged."

    processed = 0
    warnings = []
    for index, target_section in enumerate(target_doc.sections):
        source_section = reference_doc.sections[min(index, len(reference_doc.sections) - 1)]
        for kind in ("header", "footer", "first_page_header", "first_page_footer",
                     "even_page_header", "even_page_footer"):
            try:
                target_part = getattr(target_section, kind)
                source_part = getattr(source_section, kind)
            except Exception:
                continue

            if policy == "Copy the complete header/footer from the reference":
                try:
                    copy_header_footer_content(source_part, target_part)
                    processed += 1
                except Exception as exc:
                    warnings.append(f"{kind} in section {index + 1}: {exc}")
            else:
                # Safer partial option: copy paragraph formatting while retaining
                # target text. Tables and image layouts are not fully replicated here.
                try:
                    source_paras = source_part.paragraphs
                    target_paras = target_part.paragraphs
                    if source_paras and target_paras:
                        copy_paragraphs_by_position(source_paras, target_paras)
                        processed += 1
                except Exception as exc:
                    warnings.append(f"{kind} in section {index + 1}: {exc}")

    note = (
        "Header/footer content was copied from the reference. Embedded images may require "
        "relationship remapping and must be visually checked."
        if policy == "Copy the complete header/footer from the reference"
        else "Header/footer paragraph formatting was matched where possible; text and complex layouts may need manual review."
    )
    if warnings:
        note += " Some items could not be processed: " + "; ".join(warnings)
    return processed, note

if st.button("✨ Format Document B", type="primary", use_container_width=True):
    if not reference_file or not target_file:
        st.error("Please upload both DOCX files.")
    else:
        try:
            reference_doc = Document(io.BytesIO(reference_file.getvalue()))
            target_doc = Document(io.BytesIO(target_file.getvalue()))

            report = []
            if copy_page_setup:
                sections_done = copy_page_settings(reference_doc, target_doc)
                report.append(f"Page setup copied for {sections_done} target section(s).")

            if copy_styles:
                styles_done = copy_style_definitions(reference_doc, target_doc)
                report.append(f"Common named styles processed: {styles_done}.")

            paragraphs_done = copy_paragraphs_by_position(
                reference_doc.paragraphs, target_doc.paragraphs
            )
            report.append(f"Body paragraphs processed: {paragraphs_done}.")

            tables_done = 0
            if apply_tables:
                tables_done = copy_table_formatting(reference_doc, target_doc)
                report.append(f"Table paragraphs processed: {tables_done}.")

            header_footer_count, header_footer_note = copy_header_footer(
                reference_doc, target_doc, header_footer_policy
            )
            report.append(f"Header/footer parts processed: {header_footer_count}.")
            report.append(header_footer_note)

            output = io.BytesIO()
            target_doc.save(output)
            output.seek(0)

            st.success("Formatting pass completed. Review the result before using it.")
            st.subheader("Formatting report")
            for line in report:
                st.write(f"- {line}")

            st.download_button(
                "⬇️ Download formatted_document.docx",
                data=output.getvalue(),
                file_name="formatted_document.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )
            st.caption(
                "This version does not guarantee exact replication of every Word feature. "
                "Please inspect headers/footers, logos, page numbers, tables, section breaks, and pagination."
            )
        except Exception:
            st.error(
                "Could not process these files. Confirm that both are valid, unencrypted .docx files "
                "and try again."
            )
            st.exception(Exception("Processing failed. See the error details above."))

with st.expander("Master formatting specification"):
    st.markdown(
        """
        **Objective:** use Document A as the visual reference and Document B as the content source.

        **Inspect:** page size, orientation, margins, headers/footers, typography, paragraph spacing,
        heading hierarchy, lists, tables, borders, shading, graphics, and supported layout elements.

        **Preserve:** target text, meaning, data, order, tables, images, hyperlinks, and substantive content.
        Do not rewrite, summarize, translate, or invent content.

        **Validate:** save a valid DOCX, report changes, and flag elements that cannot be reliably matched.
        """
    )
