"""
Huliot India — Report Formatter & AI Comment Improver
=====================================================
• Fixes cover slide (banner, time, date, labels)
• Standardises fonts on ALL slides to Huliot format
• Standardises photo / drawing sizes on observation slides
• Uses Claude AI to rewrite observation comments in professional style
• You review + edit AI comments before the final file is built

Run:  streamlit run app.py
Deps: pip install -r requirements.txt
"""

import streamlit as st
import zipfile, re, io, time
from pathlib import Path
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.dml.color import RGBColor
import anthropic

# ══════════════════════════════════════════════════════
#  STANDARD FORMAT CONSTANTS  (Huliot India template)
# ══════════════════════════════════════════════════════
FONT_COVER_HEADING   = 36.0    # Cover title block: Huliot India / Date / Time
FONT_COVER_BODY      = 18.0    # Cover body: site name, members list
FONT_BANNER          = 28.0    # Green banner ("Site Visit")
FONT_EXPLAIN_HEAD    = 32.0    # White slides heading ("Explain - ...")
FONT_EXPLAIN_BODY    = 14.0    # White slides body text
FONT_OBS_COMMENT     = 14.0    # Observation slide comment text
FONT_CHECKLIST       = 13.0    # Drainage checklist text

GREEN_DARK  = RGBColor(0x1A, 0x5C, 0x2A)
WHITE       = RGBColor(0xFF, 0xFF, 0xFF)

# Standard photo layout on observation slides (13.33" × 7.50" slide)
PHOTO_TOP    = Inches(1.75)
PHOTO_HEIGHT = Inches(4.9)
PHOTO_LAYOUT = [
    {"left": Inches(0.20),  "width": Inches(4.10)},   # site photo 1
    {"left": Inches(4.50),  "width": Inches(4.10)},   # site photo 2
    {"left": Inches(8.80),  "width": Inches(4.30)},   # drawing / ref
]

# ══════════════════════════════════════════════════════
#  NAMESPACE FIX  (strict OOXML → transitional)
# ══════════════════════════════════════════════════════
NS_MAP = {
    "http://purl.oclc.org/ooxml/drawingml/main":
        "http://schemas.openxmlformats.org/drawingml/2006/main",
    "http://purl.oclc.org/ooxml/officeDocument/relationships":
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "http://purl.oclc.org/ooxml/presentationml/main":
        "http://schemas.openxmlformats.org/presentationml/2006/main",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument":
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/extendedProperties":
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties",
}

def fix_strict_ooxml(b: bytes) -> bytes:
    buf, out = io.BytesIO(b), io.BytesIO()
    with zipfile.ZipFile(buf, "r") as zin:
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename.endswith(".xml") or item.filename.endswith(".rels"):
                    try:
                        text = data.decode("utf-8")
                        for old, new in NS_MAP.items():
                            text = text.replace(old, new)
                        data = text.encode("utf-8")
                    except:
                        pass
                zout.writestr(item, data)
    out.seek(0)
    return out.read()

# ══════════════════════════════════════════════════════
#  HELPERS
# ══════════════════════════════════════════════════════
def shape_text(shape) -> str:
    if not shape.has_text_frame: return ""
    return " ".join("".join(r.text for r in p.runs)
                    for p in shape.text_frame.paragraphs).strip()

def set_fonts(shape, size_pt, bold=None, color=None):
    if not shape.has_text_frame: return
    for para in shape.text_frame.paragraphs:
        for run in para.runs:
            if run.text.strip():
                run.font.size = Pt(size_pt)
                if bold  is not None: run.font.bold = bold
                if color is not None: run.font.color.rgb = color

def fix_para(para, pattern, repl, flags=0):
    full = "".join(r.text for r in para.runs)
    m = re.search(pattern, full, flags)
    if not m: return False, "", ""
    new_full = re.sub(pattern, repl, full, flags=flags)
    if para.runs:
        para.runs[0].text = new_full
        for r in para.runs[1:]: r.text = ""
    return True, m.group(0), new_full

def is_obs_slide(slide, idx) -> bool:
    """True if slide is a site observation slide (has site photos, not an Explain template)."""
    pics = [s for s in slide.shapes if s.shape_type == 13]
    txt  = " ".join(shape_text(s) for s in slide.shapes)
    if not pics: return False
    if idx == 0: return False
    # Skip specific template slides (not just any mention of PRL)
    if any(kw in txt for kw in ["Explain","explain","Thank","thank"]): return False
    if "Checklist" in txt and "Drainage" in txt: return False
    if "PRL Option" in txt or "PRL Line" in txt: return False
    return True

def get_obs_text(slide) -> str:
    for shape in slide.shapes:
        if shape.has_text_frame:
            t = shape_text(shape)
            if len(t) > 20 and "u@" not in t: return t
    return ""

# ══════════════════════════════════════════════════════
#  COVER SLIDE FIX
# ══════════════════════════════════════════════════════
def fix_cover(slide, prs) -> list:
    changes = []
    SW, SH  = prs.slide_width, prs.slide_height

    for shape in slide.shapes:
        if not shape.has_text_frame: continue
        full = shape_text(shape)

        # Banner
        is_banner = (shape.width > SW*0.3 and shape.height < SH*0.15 and
                     SH*0.25 < (shape.top + shape.height/2) < SH*0.65 and
                     len(shape.text_frame.paragraphs) <= 2 and
                     full and "Huliot" not in full and len(full) < 80)
        if is_banner and "Site Visit" not in full:
            first = True
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    run.text = "Site Visit" if first else ""
                    first = False
            set_fonts(shape, FONT_BANNER, bold=True)
            changes.append(f'Banner → **"Site Visit"** (was "{full[:30]}")')

        # Time
        if re.search(r"\d+[.:]\d+\s*[Tt][oO]\s*\d+[.:]\d+", full):
            for para in shape.text_frame.paragraphs:
                ok, old, _ = fix_para(
                    para,
                    r"(\d{1,2})[.:](\d{2})\s*[Tt][oO]\s*(\d{1,2})[.:](\d{2})",
                    lambda m: f"{m.group(1)}:{m.group(2)}am to {m.group(3)}:{m.group(4)}pm")
                if ok:
                    changes.append(f'Time → correct am/pm format (was "{old}")')
            set_fonts(shape, FONT_COVER_HEADING, bold=True, color=WHITE)

        # Date spacing
        if re.search(r"\d+\s+\.\d+", full):
            for para in shape.text_frame.paragraphs:
                fix_para(para, r"(\d+)\s+\.(\d+)", r"\1.\2")
            changes.append("Date spacing corrected")

        # Client → Plumbers
        if re.search(r"\bClient\s*:", full, re.IGNORECASE):
            for para in shape.text_frame.paragraphs:
                fix_para(para, r"\bClient\s*:", "Plumbers :-", flags=re.IGNORECASE)
            changes.append('"Client :" → "Plumbers :-"')
            set_fonts(shape, FONT_COVER_BODY, bold=True, color=WHITE)

        # Members block font
        if "Members Present" in full or "Site Name" in full:
            set_fonts(shape, FONT_COVER_BODY, bold=True, color=WHITE)

    return changes

# ══════════════════════════════════════════════════════
#  OBSERVATION SLIDE — PHOTO STANDARDISE
# ══════════════════════════════════════════════════════
def standardise_photos(slide) -> list:
    changes = []
    pics = [s for s in slide.shapes if s.shape_type == 13]
    if not pics: return changes
    n = len(pics)

    if n == 1:
        layout = [{"left": Inches(1.5), "width": Inches(10.0)}]
    elif n == 2:
        layout = [
            {"left": Inches(0.2), "width": Inches(6.3)},
            {"left": Inches(6.7), "width": Inches(6.3)},
        ]
    else:
        layout = PHOTO_LAYOUT[:n]

    for i, pic in enumerate(pics[:3]):
        pos = layout[i]
        old = f'{pic.width/914400:.1f}"×{pic.height/914400:.1f}"'
        pic.top    = PHOTO_TOP
        pic.height = PHOTO_HEIGHT
        pic.left   = pos["left"]
        pic.width  = pos["width"]
        new = f'{pic.width/914400:.1f}"×{pic.height/914400:.1f}"'
        changes.append(f'Photo {i+1}: {old} → {new} standard size')
    return changes

# ══════════════════════════════════════════════════════
#  EXPLAIN / TEMPLATE SLIDES — FONT STANDARDISE
# ══════════════════════════════════════════════════════
def standardise_explain(slide, snum) -> list:
    changes = []
    for shape in slide.shapes:
        if not shape.has_text_frame: continue
        txt = shape_text(shape)
        if "u@" in txt: continue
        if "Explain" in txt or "Drainage Checklist" in txt:
            set_fonts(shape, FONT_EXPLAIN_HEAD, bold=True, color=GREEN_DARK)
            changes.append(f'Slide {snum}: Heading → {FONT_EXPLAIN_HEAD}pt bold green')
        elif len(txt) > 30:
            set_fonts(shape, FONT_EXPLAIN_BODY)
            changes.append(f'Slide {snum}: Body → {FONT_EXPLAIN_BODY}pt')
    return changes

# ══════════════════════════════════════════════════════
#  AI COMMENT REWRITER
# ══════════════════════════════════════════════════════
AI_SYSTEM = """You are a Senior Technical Manager at Huliot Pipes & Fittings Pvt. Ltd. India, 
specialising in Huliot PP single-stack drainage systems. 
Rewrite site observation comments in the exact Huliot India professional report style.

Rules:
- Begin: "It has been observed that..." OR "As observed on site..." OR "Shaft XX: ..."
- State observation clearly and technically (pipe, clamp, slope, fitting, trap, connection)
- For issues: "Kindly ensure..." or "Rectify the same at the earliest."
- For good work: acknowledge clearly e.g. "Vertical pipe installation done as per drawing."
- 2-3 sentences max. Concise and actionable.
- Reference EN 12056 / IS 1742 / Huliot table standards where relevant.
- Output ONLY the rewritten comment — no preamble, no explanation."""

def ai_improve(text: str, api_key: str) -> str:
    try:
        client   = anthropic.Anthropic(api_key=api_key)
        response = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=300,
            system=AI_SYSTEM,
            messages=[{"role": "user", "content":
                f"Rewrite this Huliot site observation in professional style:\n\n{text}"}]
        )
        return response.content[0].text.strip()
    except Exception as e:
        return f"[AI error: {e}]\n{text}"

# ══════════════════════════════════════════════════════
#  APPLY APPROVED COMMENT TO SLIDE
# ══════════════════════════════════════════════════════
def apply_comment(slide, new_text: str):
    for shape in slide.shapes:
        if not shape.has_text_frame: continue
        t = shape_text(shape)
        if len(t) > 20 and "u@" not in t:
            for para in shape.text_frame.paragraphs:
                runs = [r for r in para.runs if r.text.strip()]
                if runs:
                    runs[0].text = new_text
                    for r in runs[1:]: r.text = ""
                    # blank other paragraphs
                    for op in shape.text_frame.paragraphs:
                        if op is not para:
                            for r in op.runs: r.text = ""
                    set_fonts(shape, FONT_OBS_COMMENT)
                    return

# ══════════════════════════════════════════════════════
#  MAIN PIPELINE
# ══════════════════════════════════════════════════════
def build_report(pptx_bytes: bytes, approved_comments: dict = None):
    """
    Full pipeline. Returns (fixed_bytes, changes, obs_info_dict).
    approved_comments: {slide_idx: str} — if provided, writes into slides.
    """
    safe    = fix_strict_ooxml(pptx_bytes)
    prs     = Presentation(io.BytesIO(safe))
    changes = []
    obs     = {}   # {idx: {slide_num, original}}

    for idx, slide in enumerate(prs.slides):
        snum = idx + 1
        all_txt = " ".join(shape_text(s) for s in slide.shapes)

        if idx == 0:
            ch = fix_cover(slide, prs)
            changes += [f"Slide 1: {c}" for c in ch]

        elif is_obs_slide(slide, idx):
            orig = get_obs_text(slide)
            obs[idx] = {"slide_num": snum, "original": orig}
            ch  = standardise_photos(slide)
            changes += [f"Slide {snum}: {c}" for c in ch]
            # Font standardise for observation text
            for shape in slide.shapes:
                if shape.has_text_frame:
                    t = shape_text(shape)
                    if len(t) > 20 and "u@" not in t:
                        set_fonts(shape, FONT_OBS_COMMENT)
                        changes.append(f"Slide {snum}: Observation text → {FONT_OBS_COMMENT}pt")
            # Apply approved comment
            if approved_comments and idx in approved_comments:
                apply_comment(slide, approved_comments[idx])
                changes.append(f"Slide {snum}: AI comment applied ✅")

        else:
            if "Explain" in all_txt or "Drainage" in all_txt or "Checklist" in all_txt:
                ch = standardise_explain(slide, snum)
                changes += ch

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out.read(), changes, obs

# ══════════════════════════════════════════════════════
#  STREAMLIT UI
# ══════════════════════════════════════════════════════
def main():
    st.set_page_config(page_title="Huliot Report Formatter", page_icon="🏗️", layout="wide")

    st.markdown("""
    <div style="background:linear-gradient(90deg,#1B5E20,#388E3C);
                padding:22px 28px;border-radius:12px;margin-bottom:24px;">
        <div style="color:#A5D6A7;font-size:11px;font-weight:700;
                    letter-spacing:2px;text-transform:uppercase;margin-bottom:4px;">
            Huliot Pipes &amp; Fittings Pvt. Ltd. — Technical Manager West Zone
        </div>
        <div style="color:white;font-size:24px;font-weight:800;">
            🏗️ Report Formatter &amp; AI Comment Improver
        </div>
        <div style="color:#C8E6C9;font-size:13px;margin-top:4px;">
            Upload PPT → Auto-fix format + fonts + photo sizes → AI improves comments → Download
        </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Sidebar ──────────────────────────────────────────────────────────────
    with st.sidebar:
        st.markdown("### ⚙️ Settings")
        api_key = st.text_input(
            "Claude API Key", type="password", placeholder="sk-ant-...",
            help="Enter your Anthropic API key to enable AI comment improvement."
        )
        st.divider()
        st.markdown("### 📐 Standard Specs Applied")
        specs = [
            ("Cover heading",     f"{FONT_COVER_HEADING}pt bold white"),
            ("Cover body",        f"{FONT_COVER_BODY}pt bold white"),
            ("Banner",            f"{FONT_BANNER}pt bold — 'Site Visit'"),
            ("Explain heading",   f"{FONT_EXPLAIN_HEAD}pt bold dark green"),
            ("Explain body",      f"{FONT_EXPLAIN_BODY}pt regular"),
            ("Observation text",  f"{FONT_OBS_COMMENT}pt regular"),
            ("Photo 1 (site)",    '4.1" × 4.9"'),
            ("Photo 2 (site)",    '4.1" × 4.9"'),
            ("Drawing (ref)",     '4.3" × 4.9"'),
            ("Photos top margin", "1.75\" from top"),
        ]
        for k, v in specs:
            st.markdown(f"**{k}:** {v}")

    # ── Upload ───────────────────────────────────────────────────────────────
    st.markdown("### 📎 Step 1 — Upload Team Report (.pptx)")
    uploaded = st.file_uploader("pptx file", type=["pptx"], label_visibility="collapsed")

    if not uploaded:
        st.info("👆 Upload a .pptx report from your team to get started.")
        return

    st.success(f"✅ Loaded: **{uploaded.name}** ({uploaded.size/1024:.1f} KB)")
    pptx_bytes = uploaded.read()

    # ── Init session state ────────────────────────────────────────────────────
    if "file_name" not in st.session_state or st.session_state.file_name != uploaded.name:
        st.session_state.file_name         = uploaded.name
        st.session_state.step1_done        = False
        st.session_state.obs               = {}
        st.session_state.ai_comments       = {}
        st.session_state.approved          = {}
        st.session_state.changes           = []

    st.divider()

    # ── Step 2: Format Fix ───────────────────────────────────────────────────
    st.markdown("### 🔧 Step 2 — Auto Format Fix")
    st.caption("Fixes banner, time, date, labels, fonts, and photo sizes automatically.")

    if st.button("▶ Run Format Fix", type="primary", use_container_width=True):
        bar = st.progress(0, text="Fixing namespace...")
        for p, m in [(25,"Fixing cover slide..."),(55,"Standardising fonts..."),(80,"Resizing photos...")]:
            time.sleep(0.2); bar.progress(p, text=m)
        try:
            _, changes, obs = build_report(pptx_bytes, approved_comments=None)
        except Exception as e:
            bar.empty(); st.error(f"❌ Error: {e}"); st.exception(e); return
        bar.progress(100,"Done!"); time.sleep(0.3); bar.empty()
        st.session_state.step1_done = True
        st.session_state.changes    = changes
        st.session_state.obs        = obs
        st.session_state.ai_comments = {}
        st.session_state.approved    = {}

    if st.session_state.step1_done:
        changes = st.session_state.changes
        obs     = st.session_state.obs

        with st.expander(f"✅ {len(changes)} corrections applied — click to see details", expanded=False):
            for c in changes:
                st.markdown(f"• {c}")

        st.divider()

        # ── Step 3: AI Comments ──────────────────────────────────────────────
        st.markdown("### 🤖 Step 3 — Observation Comment Improvement")

        if not obs:
            st.info("No observation slides found — only template/instruction slides detected.")
        else:
            n_obs = len(obs)
            st.markdown(f"**{n_obs} observation slide(s) found.**")

            if not api_key:
                st.warning("⚠️ No API key entered — you can edit comments manually below.")
                for idx, info in obs.items():
                    snum = info["slide_num"]
                    st.markdown(f"**Slide {snum} — Edit Comment:**")
                    edited = st.text_area(
                        f"slide_{snum}_comment",
                        value=info["original"], height=110, label_visibility="collapsed",
                        key=f"manual_{idx}"
                    )
                    st.session_state.approved[idx] = edited

            else:
                col_a, col_b = st.columns([2, 1])
                with col_a:
                    if st.button("🤖 Generate AI Comments", use_container_width=True):
                        pb = st.progress(0, text="Starting AI...")
                        results = {}
                        for i, (idx, info) in enumerate(obs.items()):
                            pb.progress(int((i+1)/n_obs*100),
                                        text=f"Processing Slide {info['slide_num']}...")
                            results[idx] = ai_improve(info["original"], api_key)
                        pb.empty()
                        st.session_state.ai_comments = results
                        st.session_state.approved = {k: v for k, v in results.items()}
                        st.success("✅ AI comments ready — review below!")

                with col_b:
                    st.markdown("AI rewrites in Huliot professional style. You review before applying.")

                if st.session_state.ai_comments:
                    st.divider()
                    for idx, info in obs.items():
                        snum    = info["slide_num"]
                        ai_text = st.session_state.ai_comments.get(idx, info["original"])

                        st.markdown(f"#### 📋 Slide {snum}")
                        c1, c2 = st.columns(2)
                        with c1:
                            st.markdown("**📝 Team Original**")
                            st.info(info["original"] or "*(no text)*")
                        with c2:
                            st.markdown("**🤖 AI Improved**")
                            st.success(ai_text)

                        final = st.text_area(
                            f"✏️ Slide {snum} — Review & edit (this goes into the file)",
                            value=st.session_state.approved.get(idx, ai_text),
                            height=120, key=f"final_{idx}"
                        )
                        st.session_state.approved[idx] = final
                        st.divider()

        # ── Step 4: Download ─────────────────────────────────────────────────
        st.markdown("### ⬇️ Step 4 — Build & Download Final Report")

        if st.button("📦 Build Final File", type="primary", use_container_width=True):
            bar2 = st.progress(0, text="Applying all format fixes...")
            time.sleep(0.3)
            bar2.progress(50, text="Writing approved comments...")
            try:
                final_bytes, _, _ = build_report(
                    pptx_bytes,
                    approved_comments=st.session_state.approved if st.session_state.approved else None
                )
            except Exception as e:
                bar2.empty(); st.error(f"❌ Error: {e}"); return
            bar2.progress(100, text="Done!"); time.sleep(0.3); bar2.empty()

            out_name = f"FORMATTED_{Path(uploaded.name).stem}.pptx"
            st.download_button(
                label=f"⬇️  Download: {out_name}",
                data=final_bytes, file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True, type="primary"
            )
            st.caption(
                "✅ Ready. Open in PowerPoint → check slide 9 Checklist → fill Yes/NO → send to client."
            )

if __name__ == "__main__":
    main()
