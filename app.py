import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
import io

def apply_huliot_text_style(shape, font_size, alignment, is_bold=True):
    """Aggressively forces the exact visual style onto a text box."""
    # Remove weird internal margins that cause jagged text
    shape.text_frame.margin_left = Inches(0)
    shape.text_frame.margin_top = Inches(0)
    
    for paragraph in shape.text_frame.paragraphs:
        paragraph.alignment = alignment
        for run in paragraph.runs:
            run.font.size = Pt(font_size)
            run.font.bold = is_bold
            # Force the exact Huliot White color
            run.font.color.rgb = RGBColor(255, 255, 255) 
            # Force a clean, standard font to overwrite stretched fonts
            run.font.name = 'Georgia' 

def format_front_page(slide):
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
            
        text = shape.text.lower()
        
        # 1. Header (Date/Time) -> Top Left
        if "date:-" in text and "time:-" in text:
            shape.left, shape.top = Inches(0.5), Inches(0.8)
            shape.width, shape.height = Inches(5.0), Inches(1.5)
            shape.text_frame.vertical_anchor = MSO_ANCHOR.TOP
            apply_huliot_text_style(shape, font_size=20, alignment=PP_ALIGN.LEFT)

        # 2. Main Title Bar -> Center Green Bar
        elif "site visit" in text or "shiv sai" in text:
            shape.left, shape.top = Inches(2.5), Inches(2.8)
            shape.width, shape.height = Inches(5.0), Inches(0.8)
            shape.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
            apply_huliot_text_style(shape, font_size=28, alignment=PP_ALIGN.CENTER)

        # 3. Details (Site Name, Members) -> Bottom Right/Center
        elif "site name:-" in text or "members present" in text:
            shape.left, shape.top = Inches(3.2), Inches(4.3)
            shape.width, shape.height = Inches(6.5), Inches(2.5)
            shape.text_frame.vertical_anchor = MSO_ANCHOR.TOP
            apply_huliot_text_style(shape, font_size=16, alignment=PP_ALIGN.LEFT)

def format_checklist_pages(slide):
    for shape in slide.shapes:
        if shape.has_text_frame:
            text = shape.text.lower()
            if "yes / no" in text or "yes/ no" in text:
                # Force checklist to match standard exactly
                shape.left, shape.top = Inches(1.5), Inches(2.0)
                shape.width = Inches(8.0)
                apply_huliot_text_style(shape, font_size=14, alignment=PP_ALIGN.LEFT)

def process_report(uploaded_file):
    prs = Presentation(uploaded_file)
    
    for i, slide in enumerate(prs.slides):
        if i == 0:
            format_front_page(slide)
        else:
            format_checklist_pages(slide)

    pptx_io = io.BytesIO()
    prs.save(pptx_io)
    return pptx_io

# --- Streamlit UI ---
st.set_page_config(page_title="Huliot Format Editor", layout="wide")
st.title("🏗️ Huliot India: Exact Format Fixer")

st.warning("This version aggressively overwrites the stretched text with the exact White Font, exact margins, and exact alignments from your standard sample.")

uploaded_file = st.file_uploader("Upload your Draft PPTX", type="pptx")

if uploaded_file:
    if st.button("Apply Exact Format"):
        with st.spinner("Applying exact formatting..."):
            fixed_pptx = process_report(uploaded_file)
            st.success("Format Applied!")
            st.download_button(
                label="📥 Download Fixed Report",
                data=fixed_pptx.getvalue(),
                file_name="Exact_Format_Report.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
