import streamlit as st
import datetime

# --- Custom CSS for Styling ---
# This replicates the fonts and dark green theme from the sample report
st.set_page_config(page_title="Huliot India Site Visit Report", layout="wide")
st.markdown("""
    <style>
    .huliot-header {
        background-color: #005643; /* Dark green from Huliot logo */
        color: white;
        padding: 20px;
        text-align: center;
        border-radius: 5px;
        font-family: sans-serif;
    }
    .slide-title {
        color: #005643;
        font-weight: bold;
        border-bottom: 2px solid #005643;
        padding-bottom: 10px;
    }
    .stRadio label {
        font-weight: bold;
    }
    </style>
""", unsafe_allow_html=True)

# --- Sidebar Navigation ---
st.sidebar.image("https://via.placeholder.com/300x100.png?text=Huliot+India+Logo", use_container_width=True) # Replace with actual logo path
st.sidebar.title("Report Navigation")
slide = st.sidebar.radio("Go to Slide:", [
    "1. Cover Page", 
    "2. Visit Details", 
    "3. Site Photo", 
    "4. Installation Instructions", 
    "5. Joining Method", 
    "6. Bracketing & Support", 
    "7. Tools (Cutters)", 
    "8. Trap Inlet Tools", 
    "9. Trap Inlet Method", 
    "10. Drainage Checklist", 
    "11. PRL Line Options", 
    "12. Sign-off"
])

# --- Slide 1: Cover Page ---
if slide == "1. Cover Page":
    st.markdown('<div class="huliot-header"><h1>Site Visit</h1></div>', unsafe_allow_html=True) #
    st.image("https://via.placeholder.com/1200x400.png?text=Huliot+India+Logo", use_container_width=True) #

# --- Slide 2: Visit Details ---
elif slide == "2. Visit Details":
    st.markdown('<h2 class="slide-title">Huliot India Site Visit Details</h2>', unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    with col1:
        st.date_input("Date:-", datetime.date(2026, 4, 10)) #
        st.time_input("Time:-") #
        st.text_input("Site Name:-") #
        st.text_input("Location :-") #
    
    with col2:
        st.markdown("**Members Present During Site visit/ Mock up/ Testing/ Inspection**") #
        st.text_input("Huliot India – Mr.") #
        st.text_input("Contractor : Mr.") #
        st.text_input("Plumbers: - Mr.") #

# --- Slide 3: Site Photo ---
elif slide == "3. Site Photo":
    st.markdown('<h2 class="slide-title">Site Photo</h2>', unsafe_allow_html=True) #
    st.write("**SH: 01**") #
    st.info("Please upload or insert site photos here.")
    st.file_uploader("Upload Site Photo", type=["jpg", "png", "jpeg"])

# --- Slide 4: Installation Instructions ---
elif slide == "4. Installation Instructions":
    st.markdown('<h2 class="slide-title">Explain - Installation instruction</h2>', unsafe_allow_html=True) #
    
    col1, col2 = st.columns(2)
    with col1:
        st.image("https://via.placeholder.com/400x300.png?text=Pipe+Cutter", caption="Suitable Pipe Cutter")
        st.write("Cut the Huliot pipe using suitable pipe cutter available in market or a fine-tooth saw that is suitably guided to guarantee a perpendicular cut.") #
    with col2:
        st.image("https://via.placeholder.com/400x300.png?text=Chamfering+Tool", caption="Chamfering Tool")
        st.write("Chamfering and bevelling the Huliot pipe ends to an angle of roughly 15° to 30° using a suitable chamfered surface must be smooth to avoid damaging the socket when the pipe is inserted.") #

# --- Slide 5: Joining Method ---
elif slide == "5. Joining Method":
    st.markdown('<h2 class="slide-title">Explain - Installation method</h2>', unsafe_allow_html=True) #
    
    st.write("1. Join the pipe and fittings together by inserting the end /spigot into the socket to maximum socket depth.") #
    st.write("2. Ensure that the inside of the socket, the seal and spigot/end of the pipe piece to be inserted are perfectly clean.") #
    st.write("3. Lubricate the spigot/pipe/fitting end and rubber ring with the appropriate Huliot Lubricant only.") #
    
    col1, col2 = st.columns(2)
    col1.image("https://via.placeholder.com/400x300.png?text=Inserting+Pipe", caption="Push-fit insertion")
    col2.image("https://via.placeholder.com/400x300.png?text=Lubricating", caption="Applying Huliot Lubricant")

# --- Slide 6: Bracketing & Support ---
elif slide == "6. Bracketing & Support":
    st.markdown('<h2 class="slide-title">Explain - Installation method during site visit.</h2>', unsafe_allow_html=True) #
    st.write("**Maximum Bracketing Intervals for push-fit socket system for Huliot Pipes (Ultra Silent & HT Pro)**")
    
    # Placeholder for the PDF/Table images provided in the slides
    st.image("https://via.placeholder.com/800x400.png?text=Table+1:+Bracketing+Intervals", use_container_width=True)
    st.image("https://via.placeholder.com/800x400.png?text=Huliot+Round+Rubberized+Clamp", use_container_width=True)

# --- Slide 7: Tools (Cutters) ---
elif slide == "7. Tools (Cutters)":
    st.markdown('<h2 class="slide-title">Proper Pipe Cutters & Chamfering Tools</h2>', unsafe_allow_html=True)
    st.write("Explain - to use proper any pipe cutter & any chamfering tools available in market for fast, easy installation without any error /mistake") #
    
    col1, col2 = st.columns(2)
    col1.image("https://via.placeholder.com/400x300.png?text=Blue+Pipe+Cutter+1")
    col2.image("https://via.placeholder.com/400x300.png?text=Orange+Pipe+Cutter")

# --- Slide 8: Trap Inlet Tools ---
elif slide == "8. Trap Inlet Tools":
    st.markdown('<h2 class="slide-title">Explain – Installation method – Hole saw cutter</h2>', unsafe_allow_html=True) #
    st.write("Hole saw cutter for trap inlet opening size 44mm.") #
    
    col1, col2 = st.columns(2)
    col1.image("https://via.placeholder.com/400x300.png?text=44mm+Hole+Saw")
    col2.image("https://via.placeholder.com/400x300.png?text=Drill+with+Hole+Saw")

# --- Slide 9: Trap Inlet Method ---
elif slide == "9. Trap Inlet Method":
    st.markdown('<h2 class="slide-title">Explain Installation method - Trap inlet opening</h2>', unsafe_allow_html=True)
