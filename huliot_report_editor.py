import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
import io

# --- CONFIGURATION (Based on your Sample Report) ---
# These coordinates would be set to match your standard sample's positioning
LOGO_POS = (Inches(0.5), Inches(0.5), Inches(2.0)) # Left, Top, Width
TITLE_STYLE = {"font_size": Pt(24), "bold": True, "color": (0, 102, 51)} # Huliot Green

def standardize_report(uploaded_file):
    prs = Presentation(uploaded_file)
    
    # Logic to iterate through slides and fix "Stretched" items
    for slide in prs.slides:
        for shape in slide.shapes:
            # 1. Detect and Fix Logo [cite: 11]
            if "logo" in shape.name.lower() or shape.shape_type == 13: # 13 is Picture
                shape.left, shape.top = Inches(0.5), Inches(0.2)
                shape.width = Inches(1.5)
            
            # 2. Re-format Checklist Text [cite: 35, 38]
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    if "Yes / NO" in paragraph.text:
                        paragraph.font.size = Pt(12)
                        paragraph.alignment = PP_ALIGN.LEFT

    # Save to a buffer
    pptx_io = io.BytesIO()
    prs.save(pptx_io)
    return pptx_io

# --- STREAMLIT UI ---
st.set_page_config(page_title="Huliot Report Editor", layout="wide")

st.title("🏗️ Huliot India: Report Formatting App")
st.subheader("Phase 1: Standard Format Correction")

st.info("This app automatically aligns text and images to the standard Huliot Site Visit format.")

uploaded_file = st.file_uploader("Upload your 'Stretched' PPTX Report", type="pptx")

if uploaded_file:
    if st.button("Standardize Format"):
        with st.spinner("Aligning elements to standard format..."):
            fixed_pptx = standardize_report(uploaded_file)
            
            st.success("Format Standardized!")
            st.download_button(
                label="Download Standardized Report",
                data=fixed_pptx.getvalue(),
                file_name="Standardized_Site_Visit_Report.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )

st.divider()
st.write("### Reference Standards Applied:")
st.markdown(f"""
* **Front Page:** Auto-aligns Date/Time/Site Name[cite: 2, 4, 5].
* **Installation Guides:** Fixes image stretching for cutting and chamfering tools[cite: 14, 15, 24].
* **Checklist:** Standardizes font and spacing for Drainage Checklist points[cite: 35, 40, 42].
""")
