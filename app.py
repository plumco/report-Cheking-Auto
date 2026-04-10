import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
import io

def fix_front_page(slide):
    """Specifically targets and positions elements to match the Huliot Template."""
    
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
            
        text = shape.text_frame.text.lower()
        
        # 1. Header Block (Huliot India, Date, Time) -> Top Left
        if "date:-" in text and "time:-" in text:
            shape.left = Inches(0.4)
            shape.top = Inches(0.8)
            shape.width = Inches(5.0) 
            
            # Force left alignment
            for paragraph in shape.text_frame.paragraphs:
                paragraph.alignment = PP_ALIGN.LEFT

        # 2. Main Title Bar (Site Visit / Project Name) -> Center vertical, Middle horizontal
        # We check that it's NOT the details block by ensuring "location" isn't in it
        elif ("site visit" in text or "shiv sai" in text) and "location" not in text:
            shape.left = Inches(2.5)
            shape.top = Inches(2.8)
            shape.width = Inches(5.0)
            
            # Force center alignment for the title
            for paragraph in shape.text_frame.paragraphs:
                paragraph.alignment = PP_ALIGN.CENTER

        # 3. Details Bullet List -> Pushed Down and Right, but Left-Aligned text
        elif "site name:-" in text or "location" in text or "members present" in text:
            # Push the box down to Inches(4.2) so it doesn't hit the Title Bar
            shape.left = Inches(3.2)
            shape.top = Inches(4.0)
            shape.width = Inches(6.5) # Wide enough to stop text from wrapping awkwardly
            
            # Force left alignment for the bullet points
            for paragraph in shape.text_frame.paragraphs:
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
st.subheader("Targeting: Template Alignment Fixes")

st.markdown("""
This app reads the text boxes on your draft's first slide and snaps them into the exact coordinates of the standard Huliot visual template.
""")

uploaded_file = st.file_uploader("Upload your 'Stretched' PPTX Report", type="pptx")

if uploaded_file:
    if st.button("Format Slide to Template"):
        with st.spinner("Re-aligning text boxes to standard template..."):
            fixed_pptx = standardize_report(uploaded_file)
