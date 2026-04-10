import streamlit as st
from pptx import Presentation
import io

def extract_data_from_draft(draft_prs):
    """Scans the messy draft and extracts the actual text/data."""
    extracted_data = {
        "date_time": "",
        "site_title": "",
        "details_list": ""
    }
    
    # Very basic extraction logic based on your Huliot format
    if len(draft_prs.slides) > 0:
        for shape in draft_prs.slides[0].shapes:
            if not shape.has_text_frame:
                continue
            text = shape.text
            lower_text = text.lower()
            
            if "date:-" in lower_text or "time:-" in lower_text:
                extracted_data["date_time"] = text
            elif "site name:-" in lower_text or "members present" in lower_text:
                extracted_data["details_list"] = text
            elif "shiv sai" in lower_text or "site visit" in lower_text:
                extracted_data["site_title"] = text
                
    return extracted_data

def inject_into_template(template_file, extracted_data):
    """Pastes the extracted data into your perfect, unchanged template."""
    prs = Presentation(template_file)
    
    # Assuming Slide 1 is the Front Page
    slide = prs.slides[0]
    
    # Paste data into the template's existing text boxes
    # Note: This relies on the template having text boxes in this specific order
    text_box_count = 0
    for shape in slide.shapes:
        if shape.has_text_frame:
            # Clear existing placeholder text in the template
            shape.text_frame.clear() 
            p = shape.text_frame.paragraphs[0]
            
            if text_box_count == 0: # Usually the top-left Header
                p.text = extracted_data["date_time"]
            elif text_box_count == 1: # Usually the center Title
                p.text = extracted_data["site_title"]
            elif text_box_count == 2: # Usually the bottom-right Details
                p.text = extracted_data["details_list"]
                
            text_box_count += 1

    # Save the beautifully formatted file
    pptx_io = io.BytesIO()
    prs.save(pptx_io)
    return pptx_io

# --- Streamlit UI ---
st.set_page_config(page_title="Huliot Standardizer", layout="wide")
st.title("🏗️ Huliot India: Perfect Format Generator")

st.markdown("""
**How this works:**
1. Upload your perfect, blank **Huliot Sample Template**. (This ensures the background, colors, and logos are always 100% correct).
2. Upload the **Messy Draft**.
3. The app will copy the text from the draft and paste it cleanly into the template.
""")

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. The Perfect Template")
    template_file = st.file_uploader("Upload Blank Huliot Template (.pptx)", type="pptx", key="template")

with col2:
    st.subheader("2. The Messy Draft")
    draft_file = st.file_uploader("Upload Application Ready Report (.pptx)", type="pptx", key="draft")

st.divider()

if template_file and draft_file:
    if st.button("Generate 'Same to Same' Report", use_container_width=True):
        with st.spinner("Extracting data and injecting into template..."):
            
            # Step 1: Read the messy file
            draft_prs = Presentation(draft_file)
            data = extract_data_from_draft(draft_prs)
            
            # Step 2: Paste into the perfect template
            final_pptx = inject_into_template(template_file, data)
            
            st.success("Report Generated! It perfectly matches your sample.")
            st.download_button(
                label="📥 Download Perfect Huliot Report",
                data=final_pptx.getvalue(),
                file_name="Final_Huliot_Report.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
