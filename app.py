import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
import io

def fix_front_page(slide):
    """Specifically targets and positions elements on the title slide."""
    
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
            
        text = shape.text_frame.text.lower()
        
        # 1. Header Block (Huliot India, Date, Time)
        # Moving it to the Top-Left
        if "date:-" in text and "time:-" in text:
            shape.left = Inches(0.5)
            shape.top = Inches(0.8)
            shape.width = Inches(5.0)
            
            # Standardize font size
            for paragraph in shape.text_frame.paragraphs:
                paragraph.font.size = Pt(20)
                paragraph.font.bold = True
                paragraph.alignment = PP_ALIGN.LEFT

        # 2. Main Title Bar (Site Visit / Shiv Sai Paradies)
        # Moving it to the Center Horizontal Bar
        elif "site visit" in text or "shiv sai" in text:
            shape.left = Inches(3.0)
            shape.top = Inches(3.2)
            shape.width = Inches(5.0)
            
            for paragraph in shape.text_frame.paragraphs:
                paragraph.font.size = Pt(28)
                paragraph.font.bold = True
                paragraph.alignment = PP_ALIGN.CENTER

        # 3. Details Bullet List (Site Name, Location, Members Present)
        # Moving it to the Bottom-Center/Right Area
        elif "location" in text and "members present" in text:
            shape.left = Inches(3.2)
            shape.top = Inches(4.2)
            shape.width = Inches(6.5)
            
            for paragraph in shape.text_frame.paragraphs:
                paragraph.font.size = Pt(14)
                paragraph.font.bold = True
                paragraph.alignment = PP_ALIGN.LEFT

def standardize_report(uploaded_file):
    prs = Presentation(uploaded_file)
    
    # Apply the strict layout fix ONLY to the first slide (index 0)
    if len(prs.slides) > 0:
        fix_front_page(prs.slides[0])
    
    # Save to buffer
    pptx_io = io.BytesIO()
    prs.save(pptx_io)
    return pptx_io

# --- Streamlit UI ---
st.set_page_config(page_title="Huliot Front Page Fixer", layout="wide")

st.title("🏗️ Huliot India: Report Formatter")
st.error("Targeting specific misalignments on the Title Page")

st.markdown("""
This updated script does not just format text; it physically relocates the **Header**, **Center Title**, and **Details List** back to their designated template coordinates.
""")

uploaded_file = st.file_uploader("Upload your 'Stretched' PPTX Report", type="pptx")

if uploaded_file:
    if st.button("Fix Layout"):
        with st.spinner("Snapping elements back to template grid..."):
            fixed_pptx = standardize_report(uploaded_file)
            
            st.success("Layout Fixed!")
            st.download_button(
                label="📥 Download Corrected Report",
                data=fixed_pptx.getvalue(),
                file_name="Fixed_Huliot_Report.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
