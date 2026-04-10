import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
import io

def standardize_report(uploaded_file):
    # Load the uploaded presentation
    prs = Presentation(uploaded_file)
    
    # Iterate through slides and shapes to apply standard formatting
    for slide in prs.slides:
        for shape in slide.shapes:
            
            # Example 1: Fix Text Formatting (e.g., Checklist alignment)
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    # Target the checklist points
                    if "Yes / NO" in paragraph.text or "Yes/ No" in paragraph.text:
                        paragraph.font.size = Pt(12)
                        paragraph.alignment = PP_ALIGN.LEFT
            
            # Example 2: Anchor/Resize Images (e.g., Logo)
            # Shape type 13 indicates a picture
            if shape.shape_type == 13: 
                # Note: You can add specific size/location logic here. 
                # This ensures an image doesn't exceed a certain width.
                if shape.width > Inches(6):
                    shape.width = Inches(5)

    # Save to a memory buffer so the user can download it directly
    pptx_io = io.BytesIO()
    prs.save(pptx_io)
    return pptx_io

# --- Streamlit UI Configuration ---
st.set_page_config(page_title="Huliot Report Editor", layout="wide")

st.title("🏗️ Huliot India: Report Formatting App")
st.subheader("Phase 1: Standard Format Correction")
st.info("Upload a draft Site Visit Report. This tool will auto-align the checklist and standardize image boundaries.")

# File Uploader
uploaded_file = st.file_uploader("Upload your 'Stretched' PPTX Report", type="pptx")

# Action Button
if uploaded_file:
    if st.button("Standardize Format"):
        with st.spinner("Aligning elements to standard format..."):
            fixed_pptx = standardize_report(uploaded_file)
            
            st.success("Format Standardized Successfully!")
            st.download_button(
                label="📥 Download Standardized Report",
                data=fixed_pptx.getvalue(),
                file_name="Standardized_Site_Visit_Report.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
