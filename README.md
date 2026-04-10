# Huliot India — Report Formatter & AI Comment Improver

## What This App Does

| Feature | Details |
|---------|---------|
| **Cover Slide Fix** | Banner → "Site Visit", Time → am/pm format, Date spacing, "Client" → "Plumbers" |
| **Font Standardise** | All slides get correct Huliot font sizes (heading 32pt, body 14pt, etc.) |
| **Photo Standardise** | All observation slide photos set to standard size & position |
| **AI Comments** | Claude AI rewrites observation text in Huliot professional style |
| **Review Before Apply** | You see Original vs AI side-by-side, edit if needed, then download |

## Font Standards Applied

| Element | Size | Style |
|---------|------|-------|
| Cover heading | 36pt | Bold White |
| Cover body | 18pt | Bold White |
| Green banner | 28pt | Bold White |
| Explain heading | 32pt | Bold Dark Green |
| Explain body | 14pt | Regular |
| Observation comment | 14pt | Regular |

## Photo Standard Layout (per observation slide)

```
|  Photo 1 (site)  |  Photo 2 (site)  |  Drawing (ref)  |
|   4.1" × 4.9"   |   4.1" × 4.9"   |   4.3" × 4.9"  |
     Top: 1.75" from top of slide
```

## Setup (One Time Only)

```bash
pip install -r requirements.txt
```

## Run

```bash
streamlit run app.py
```

## How to Use

1. **Enter API Key** in the sidebar (optional — for AI comments)
2. **Upload** the team .pptx report
3. Click **Run Format Fix** — fixes cover, fonts, photo sizes
4. Click **Generate AI Comments** — AI rewrites observation text
5. **Review** each AI comment side-by-side, edit if needed
6. Click **Build Final File** → **Download**
7. Open in PowerPoint → fill Checklist Yes/NO → send to client

## AI Comment Style

The AI writes comments in the exact Huliot professional format:
- "It has been observed that..."
- "Kindly ensure clamp supports are fixed as per Huliot Table..."
- "Rectify the same at the earliest."

You always review and can edit before the final file is created.

---
Umesh Nikam | Ass. Technical Manager | Huliot Pipes & Fittings Pvt. Ltd.
