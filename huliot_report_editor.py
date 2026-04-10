"""
Huliot India – Site Visit Report Editor
Phase 1: Cover Page + Sign-off Page field editor
"""

import streamlit as st
import zipfile
import io
from datetime import date, time
from lxml import etree

# ─── Namespace ────────────────────────────────────────────────────────────────
NS_A = "http://purl.oclc.org/ooxml/drawingml/main"


# ─── Core patch functions ─────────────────────────────────────────────────────

def find_text_node(tree, prefix):
    """Return the first <a:t> node whose stripped text starts with prefix."""
    for t in tree.iter(f"{{{NS_A}}}t"):
        if t.text and t.text.strip().startswith(prefix):
            return t
    return None


def set_value(t_node, new_value):
    """
    Rewrite a text node keeping the label part (up to the separator)
    and replacing everything after it with new_value.
    """
    if t_node is None:
        return
    original = t_node.text or ""
    for sep in [":-", ": -", "–", ":"]:
        idx = original.find(sep)
        if idx != -1:
            t_node.text = original[:idx + len(sep)] + " " + new_value
            return
    t_node.text = original.rstrip() + " " + new_value


def patch_slide(xml_bytes, field_map):
    """Apply field_map = {text_prefix: new_value} to a slide XML bytes."""
    tree = etree.fromstring(xml_bytes)
    for prefix, new_val in field_map.items():
        if not new_val:
            continue
        t = find_text_node(tree, prefix)
        if t is not None:
            set_value(t, new_val)
    return etree.tostring(tree, xml_declaration=True, encoding="UTF-8", standalone=True)


def patch_pptx(pptx_bytes: bytes, slide1_map: dict, slide11_map: dict) -> bytes:
    """
    Read PPTX bytes, patch slide1 and slide11, return modified PPTX bytes.
    Works with Huliot template's purl.oclc.org namespace (older PowerPoint format).
    """
    input_buf  = io.BytesIO(pptx_bytes)
    output_buf = io.BytesIO()

    with zipfile.ZipFile(input_buf, "r") as zin:
        with zipfile.ZipFile(output_buf, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "ppt/slides/slide1.xml" and slide1_map:
                    data = patch_slide(data, slide1_map)
                elif item.filename == "ppt/slides/slide11.xml" and slide11_map:
                    data = patch_slide(data, slide11_map)
                zout.writestr(item, data)

    return output_buf.getvalue()


# ─── Page config ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Huliot Report Editor",
    page_icon="📋",
    layout="wide",
)

# ─── Custom CSS ──────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .huliot-header {
        background: linear-gradient(90deg, #00622F 0%, #009A4E 100%);
        padding: 18px 28px;
        border-radius: 10px;
        display: flex;
        align-items: center;
        gap: 18px;
        margin-bottom: 20px;
    }
    .huliot-header h1 { color:white!important; font-size:24px!important; margin:0!important; }
    .huliot-header p  { color:#b2f0cd!important; font-size:13px!important; margin:3px 0 0 0!important; }

    .section-card {
        background:#f4faf7;
        border:1px solid #c2e2d0;
        border-left:5px solid #00a94f;
        border-radius:8px;
        padding:14px 20px;
        margin-bottom:12px;
    }
    .section-card h3 { color:#00622F; margin:0; font-size:16px; }

    label { color:#1a3d2b!important; font-weight:500!important; }

    .stDownloadButton>button {
        background:#00622F!important; color:white!important;
        font-size:17px!important; border-radius:8px!important; height:52px!important;
    }
    footer { visibility:hidden; }
</style>
""", unsafe_allow_html=True)

# ─── Header ──────────────────────────────────────────────────────────────────
st.markdown("""
<div class="huliot-header">
    <div style="font-size:40px">📋</div>
    <div>
        <h1>Huliot India · Site Visit Report Editor</h1>
        <p>Phase 1 &nbsp;·&nbsp; Fills Cover Page & Sign-off Slide to standard format automatically</p>
    </div>
</div>
""", unsafe_allow_html=True)

# ─── Step 1 – Upload ─────────────────────────────────────────────────────────
st.markdown('<div class="section-card"><h3>① Upload PPTX Template</h3></div>', unsafe_allow_html=True)

uploaded = st.file_uploader(
    "Upload Huliot Site Visit Report template (.pptx)",
    type=["pptx"],
    label_visibility="collapsed",
)

if not uploaded:
    st.info("⬆️  Upload your Huliot PPTX template above to begin editing.")
    st.stop()

pptx_bytes = uploaded.read()
st.success(f"✅  **{uploaded.name}** loaded — {len(pptx_bytes)//1024} KB")

st.markdown("---")

# ─── Step 2 – Cover page fields ──────────────────────────────────────────────
st.markdown('<div class="section-card"><h3>② Cover Page Fields &nbsp; (Slide 1)</h3></div>', unsafe_allow_html=True)

col_l, col_r = st.columns(2, gap="large")

with col_l:
    site_name  = st.text_input("🏗️  Site Name",             placeholder="e.g. Lodha Palava Tower C")
    location   = st.text_input("📍  Location",               placeholder="e.g. Dombivli East, Thane")
    huliot_rep = st.text_input("👤  Huliot India Rep (Mr.)", placeholder="e.g. Umesh Kulkarni")

with col_r:
    visit_date = st.date_input("📅  Date of Visit", value=date.today())
    tc1, tc2 = st.columns(2)
    with tc1:
        time_from = st.time_input("⏰  Time From", value=time(10, 0))
    with tc2:
        time_to   = st.time_input("⏰  Time To",   value=time(17, 0))
    contractor = st.text_input("🏢  Contractor (Mr.)",       placeholder="e.g. Ramesh Patil")
    plumber    = st.text_input("🔧  Plumber (Mr.)",          placeholder="e.g. Suresh More")

st.markdown("---")

# ─── Step 3 – Sign-off page ──────────────────────────────────────────────────
st.markdown('<div class="section-card"><h3>③ Sign-off Slide &nbsp; (Last Slide)</h3></div>', unsafe_allow_html=True)

sc1, sc2 = st.columns(2, gap="large")

with sc1:
    signoff_site = st.text_input("📋  Site Visit label",          value=site_name or "",
                                 placeholder="e.g. Lodha Palava Tower C")
    prepared_by  = st.text_input("✍️  Report Prepared by",        placeholder="e.g. Umesh Kulkarni")

with sc2:
    approved_by  = st.text_input("✅  Checked & Approved by",     placeholder="e.g. Vishal Shah")

st.markdown("---")

# ─── Step 4 – Generate ───────────────────────────────────────────────────────
if st.button("🚀  Generate Updated Report", type="primary", use_container_width=True):

    date_str = visit_date.strftime("%d.%m.%Y")
    time_str = (
        time_from.strftime("%I:%M%p").lstrip("0").lower()
        + " to "
        + time_to.strftime("%I:%M%p").lstrip("0").lower()
    )

    slide1_map = {
        "Date:-":             date_str,
        "Time:-":             time_str,
        "Site Name:-":        site_name,
        "Location :-":        location,
        "Huliot India \u2013 Mr.": huliot_rep,   # – is U+2013
        "Contractor : Mr.":   contractor,
        "Plumbers: - Mr.":    plumber,
    }

    slide11_map = {
        "Site Visit :-":                   signoff_site,
        "Report Prepared by :-":           prepared_by,
        "Report check and approved by :-": approved_by,
    }

    with st.spinner("Patching template…"):
        try:
            result  = patch_pptx(pptx_bytes, slide1_map, slide11_map)
            safe    = (site_name or "Report").replace(" ", "_")
            out_name = f"Huliot_SiteVisit_{safe}_{date_str}.pptx"

            st.success("✅  Report generated! Click below to download.")

            st.download_button(
                label="⬇️  Download Updated Report (.pptx)",
                data=result,
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True,
            )

            # Summary of changes
            with st.expander("📄 Summary of Changes Applied", expanded=True):
                rows = [
                    ("📅 Date",          date_str),
                    ("⏰ Time",          time_str),
                    ("🏗️ Site Name",     site_name   or "—"),
                    ("📍 Location",      location    or "—"),
                    ("👤 Huliot Rep",    huliot_rep  or "—"),
                    ("🏢 Contractor",    contractor  or "—"),
                    ("🔧 Plumber",       plumber     or "—"),
                    ("📋 Sign-off Site", signoff_site or "—"),
                    ("✍️ Prepared by",   prepared_by or "—"),
                    ("✅ Approved by",   approved_by or "—"),
                ]
                half = len(rows) // 2
                ca, cb = st.columns(2)
                with ca:
                    for label, val in rows[:half]:
                        st.markdown(f"**{label}:** {val}")
                with cb:
                    for label, val in rows[half:]:
                        st.markdown(f"**{label}:** {val}")

        except Exception as e:
            st.error(f"❌  Error generating report: {e}")
            st.exception(e)

# ─── Footer ──────────────────────────────────────────────────────────────────
st.markdown("---")
st.caption("Huliot India · Site Visit Report Editor · Phase 1 · PVL Ltd")
