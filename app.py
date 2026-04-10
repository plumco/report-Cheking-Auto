import streamlit as st

# Configure the page to a wide layout to better match a presentation slide feel
st.set_page_config(page_title="Huliot Site Visit Report Editor", layout="wide")

# Inject Custom CSS to match fonts and positioning styling
st.markdown("""
    <style>
    /* Base font styling to match typical presentation fonts */
    html, body, [class*="css"]  {
        font-family: 'Segoe UI', Arial, sans-serif;
    }
    
    /* Huliot Green styling for headers */
    .huliot-header {
        color: #007A53; /* Adjust to exact Huliot Green */
        font-weight: bold;
        border-bottom: 2px solid #007A53;
        padding-bottom: 10px;
        margin-bottom: 20px;
    }
    
    .section-title {
        background-color: #007A53;
        color: white;
        padding: 10px;
        border-radius: 5px;
        margin-top: 20px;
        margin-bottom: 20px;
    }
    </style>
""", unsafe_allow_html=True)

# --- FRONT PAGE ---
st.markdown("<h1 class='huliot-header'>Huliot India - Site Visit Report</h1>", unsafe_allow_html=True)

# --- SITE DETAILS SECTION ---
st.markdown("<h3 class='section-title'>Site Information</h3>", unsafe_allow_html=True)
col1, col2 = st.columns(2)

with col1:
    report_date = st.date_input("Date")
    site_name = st.text_input("Site Name")
    
with col2:
    time_range = st.text_input("Time", value="00:00am to 00:00pm")
    location = st.text_input("Location")

# --- MEMBERS PRESENT ---
st.markdown("<h3 class='section-title'>Members Present During Site Visit / Mock up / Testing / Inspection</h3>", unsafe_allow_html=True)
col3, col4, col5 = st.columns(3)

with col3:
    huliot_rep = st.text_input("Huliot India – Mr.")
with col4:
    contractor = st.text_input("Contractor : Mr.")
with col5:
    plumber = st.text_input("Plumbers: - Mr.")

st.divider()

# --- EXPLANATION SECTIONS (Static text / Placeholders for photos) ---
st.markdown("<h3 class='section-title'>Installation Explanations & Methods</h3>", unsafe_allow_html=True)

st.subheader("1. Pipe Cutting & Chamfering")
st.info("Cut the Huliot pipe using suitable pipe cutter available in market or a fine-tooth saw that is suitably guided to guarantee a perpendicular cut. Chamfering and bevelling the Huliot pipe ends to an angle of roughly 15° to 30° using a suitable chamfered surface must be smooth to avoid damaging the socket when the pipe is inserted.")

st.subheader("2. Joining Method")
st.info("Join the pipe and fittings together by inserting the end/spigot into the socket to maximum socket depth. Ensure that the inside of the socket, the seal and spigot/end of the pipe piece to be inserted are perfectly clean. Lubricate the spigot/pipe/fitting end and rubber ring with the appropriate Huliot Lubricant only.")

st.subheader("3. Trap Inlet Opening")
st.info("Use 44 mm size hole saw cutter - remove burs if any with fine-file inside trap to avoid any blockages during operation.")

st.subheader("4. PRL Line Options")
st.info("Huliot Standard: Using 45 deg or 90 deg bend at pressure relief line.")

st.divider()

# --- DRAINAGE CHECKLIST ---
st.markdown("<h3 class='section-title'>Drainage Checklist Points (For Installation Process & Testing)</h3>", unsafe_allow_html=True)

checklist_options = ["Yes", "No", "Not Applicable"]

chk1 = st.radio("1. Pipe installation done as per drawing / site requirements / site in charge instruction approval", checklist_options, horizontal=True)
chk2 = st.radio("2. Pipe routine is as per drawing / as per site requirements changes / site in charge instruction approval", checklist_options, horizontal=True)
chk3 = st.radio("3. Pipe supports/ clamps provided with proper distance as per Huliot Table or as per site in charge instruction", checklist_options, horizontal=True)
chk4 = st.radio("4. Trap support provided tightly", checklist_options, horizontal=True)
chk5 = st.radio("5. Proper pipe slope maintains as per plumbing consultant or site requirements", checklist_options, horizontal=True)
chk6 = st.radio("6. Open pipe end closed with end cap or any other materials", checklist_options, horizontal=True)
chk7 = st.radio("7. Drainage - Water testing done in toilet with 500 mm or 1 meter height or as per site requirements with proper end cap", checklist_options, horizontal=True)
chk8 = st.radio("8. During water testing water leakages found", checklist_options, horizontal=True)
chk9 = st.radio("9. Leakages rectification done immediately", checklist_options, horizontal=True)
chk10 = st.radio("10. Inside Traps cement mortar clean before fixing Grating/zali / shower channel", checklist_options, horizontal=True)

st.divider()

# --- SIGN-OFF SECTION ---
st.markdown("<h3 class='section-title'>Report Sign-off</h3>", unsafe_allow_html=True)
st.caption("Record should be maintain for each and every toilet installation checklist & testing with site supervisor sign")

col6, col7 = st.columns(2)
with col6:
    prepared_by = st.text_input("Report Prepared by :")
with col7:
    approved_by = st.text_input("Report check and approved by :")

# --- EXPORT / SAVE BUTTON ---
if st.button("Save & Generate Report", type="primary"):
    st.success("Report data captured successfully! (Next step: integrate python-pptx to push this data into the actual PPTX template).")
