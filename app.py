import streamlit as st
from pptx import Presentation
from pptx.util import Pt
import io

st.set_page_config(page_title="Report Formatter App", layout="centered")

st.title("📄 Site Visit Report Formatter")
st.write("Upload your raw Site Visit Report to automatically format it to match the Sample Report standards.")

# Step 1: File Uploaders
raw_report = st.file_uploader("1. Upload Raw Report (.pptx)", type="pptx")
sample_template = st.file_uploader("2. Upload Sample Report Template (.pptx) - Optional for advanced layouts", type="pptx")

# Target Font (Change this to match your sample report exactly, e.g., 'Arial' or 'Calibri')
TARGET_FONT = "Arial" 

def format_presentation(presentation):
    """
    Iterates through every slide and shape to apply the standard font.
    """
    for slide in presentation.slides:
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            
            # Update the font for every paragraph and run of text
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    run.font.name = TARGET_FONT
                    # You can also set a standard font size here if needed:
                    # run.font.size = Pt(12) 
                    
    return presentation

if st.button("Convert to Sample Format"):
    if raw_report is not None:
        with st.spinner("Formatting your report..."):
            try:
                # Load the raw presentation
                prs = Presentation(raw_report)
                
                # Apply the formatting (Font standardizing)
                formatted_prs = format_presentation(prs)
                
                # Save the formatted presentation to a memory buffer
                output = io.BytesIO()
                formatted_prs.save(output)
                output.seek(0)
                
                st.success("Report successfully formatted!")
                
                # Download button
                st.download_button(
                    label="⬇️ Download Formatted Report",
                    data=output,
                    file_name="Formatted_Site_Visit_Report.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
                
            except Exception as e:
                st.error(f"An error occurred: {e}")
    else:
        st.warning("Please upload a raw report to begin.")
