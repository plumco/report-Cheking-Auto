# 🏗️ Huliot India — Site Visit Report Formatter + AI Comment Fixer

## What This App Does

1. **Upload** any team PPTX report
2. **Auto-fixes** all format issues:
   - Banner text → "Site Visit"  
   - Time format → "10:30am to 01:30pm"
   - Title font → 32pt
   - Body text → 18–20pt
   - Images → max 5" wide, 4.5" tall
   - Label "Client" → "Plumbers :-"
   - Spelling errors ("PRL ines" → "PRL lines")
3. **AI rewrites** observation comments to professional standard
4. **You review** and edit comments in the app
5. **Download** the perfectly formatted PPTX — ready to send to client

---

## Setup (One Time — 2 Minutes)

### Requirements
- Python 3.9 or higher installed on your laptop
- Internet connection (for AI comment improvement)

### Steps

**Step 1 — Unzip this folder**

**Step 2 — Open Command Prompt in this folder**
Right-click inside the folder → "Open in Terminal" (Windows) or "New Terminal at Folder" (Mac)

**Step 3 — Install dependencies**
```
pip install -r requirements.txt
```

**Step 4 — Run the app**
```
streamlit run app.py
```

The app opens automatically in your browser at: http://localhost:8501

---

## Daily Use

1. Run: `streamlit run app.py`
2. Upload team PPTX
3. Click **Fix Format & Improve Comments**
4. Review AI comments (edit if needed)
5. Click **Generate Final File**
6. Download → open in PowerPoint → quick check → send to client

---

## What Gets Fixed Automatically

| Issue | Fixed |
|-------|-------|
| Green banner wrong text | ✅ Auto |
| Time format (dots/caps) | ✅ Auto |
| Date spacing error | ✅ Auto |
| "Client" label | ✅ Auto |
| Spelling errors | ✅ Auto |
| Font sizes (title/body) | ✅ Auto |
| Oversized photos | ✅ Auto |
| Observation comments | 🤖 AI + Your review |
| Checklist Yes/NO | ✏️ Manual in PowerPoint |

---

## Support
Mr. Umesh Nikam — Ass. Technical Manager, Huliot Pipes & Fittings Pvt. Ltd.
