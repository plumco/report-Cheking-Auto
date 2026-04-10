import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
import io

def format_title_page(shape, text):
    """Formats elements specifically on the Title Page."""
    if "time:-" in text:
        shape.left, shape.top = Inches(0.4), Inches(0.8)
        shape.width = Inches(5.0)
        shape.text_frame.vertical_anchor = MSO_ANCHOR.TOP
        for p in shape.text_frame.paragraphs: p.alignment = PP_ALIGN.LEFT
            
    elif "members present" in text:
        shape.left, shape.top = Inches(3.2), Inches(4.3)
        shape.width = Inches(6.5)
        shape.text_frame.vertical_anchor = MSO_ANCHOR.TOP
        for p in shape.text_frame.paragraphs: p.alignment = PP_ALIGN.LEFT
            
    elif "shiv sai" in text or "site visit" in text:
        shape.left, shape.top = Inches(2.5), Inches(2.8)
        shape.width = Inches(5.0)
        shape.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        for p in shape.text_frame.paragraphs: p.alignment = PP_ALIGN.CENTER

def format_checklist(shape):
    """Standardizes the Checklist slides (Yes/NO points)."""
    shape.left = Inches(1.5) # Align slightly indented from the left
    shape.width = Inches(8.0) # Give it plenty of width to avoid text wrapping
    for paragraph in shape.text_frame.paragraphs:
        paragraph.font.size = Pt(14) # Standard readable size
        paragraph.font.bold = True
        paragraph.alignment = PP_ALIGN.LEFT

def format_sign_off(shape):
    """Standardizes the final approval slide."""
    shape.left = Inches(1.0)
    shape.width = Inches(8.0)
    for paragraph in shape.text_frame.paragraphs:
        paragraph.font.size = Pt(16)
        paragraph.font.bold = True
        paragraph.alignment = PP_ALIGN.LEFT

def fix_stretched_images(shape):
    """Prevents images from stretching off the slide."""
    # Max width for an image on a standard slide is usually around 8.5 inches
    if shape.width > Inches(8.5):
        # Lock the aspect ratio and scale down
        ratio = shape.height / shape.width
        shape.width = Inches(8.0)
        shape.height = int(shape.width * ratio)
        shape.left = Inches(1.0) # Center-ish

def process_report(uploaded_file):
    prs = Presentation(uploaded_file)
    
    # Iterate through EVERY slide in the presentation
    for slide in prs.slides:
        for shape in slide.shapes:
            
            # 1. FIX STRETCHED IMAGES (Shape type 13 is Picture)
            if shape.shape_type == 13:
                fix_stretched_images(shape)
                continue

            # 2. FIX TEXT ELEMENTS
            if shape.has_text_frame:
                text = shape.text.lower()
                
                # Check if it's a checklist item
                if "yes / no" in text or "yes/ no" in text:
                    format_checklist(shape)
                
                # Check if it's the sign-off page
                elif "report prepared by" in text or "report check and approved" in text:
                    format_sign_off(shape)
                    
                # Check if it's the title page details
                elif "members present" in text or "time:-" in text:
                    format_title_page(shape, text)
                
                # Check for Installation Headers (e.g., "Explain - Installation")
                elif "explain -" in text:
                    shape.top = Inches(0.5) # Force headers to the top
                    for p in shape.text_frame.paragraphs:
                        p.font.size = Pt(20)
                        p.font.bold = True

    # Save to buffer
    pptx_io = io.BytesIO()
    prs.save(pptx_io)
    return pptx_io

# --- Streamlit UI ---
st.set_page_config(page_title="Huliot Full Report Editor", layout="wide")

st.title("🏗️ Huliot India: Full Report Formatter")
st.subheader("Phase 2: Document-Wide Standardization")

st.markdown("""
This tool now scans **every slide** in your uploaded draft. 
* It fixes Front Page alignment.
* It standardizes the font and spacing for the **Drainage Checklist**.
* It ensures **Installation Explanations** and **Sign-off pages** match the Huliot standard.
* It shrinks severely stretched tool photos to fit within the slide boundaries.
""")

uploaded_file = st.file_uploader("Upload your Draft PPTX Report", type="pptx")

if uploaded_file:
    if st.button("Apply Huliot Standard Format"):
        with st.spinner("Scanning and formatting entire document..."):
            fixed_pptx = process_report(uploaded_file)
            
            st.success("Document Standardized Successfully!")
            st.download_button(
                label="📥 Download Standardized Report",
                data=fixed_pptx.getvalue(),
                file_name="Huliot_Standardized_Report.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
