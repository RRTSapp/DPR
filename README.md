# ☀️ Solar DPR Generator

A Flask web application that generates complete **Detailed Project Reports (DPR)** 
for grid-connected solar PV power projects. Fill in ~25 fields, click Generate — 
get a professional 30+ page Word document instantly.

---

## 📁 File Structure

```
dpr_generator/
├── app.py               ← Flask routes (main entry point)
├── dpr_generator.py     ← Word document builder (python-docx)
├── financial_model.py   ← 25-year cash flow + IRR + DSCR model
├── templates/
│   └── index.html       ← Single-page input form
├── requirements.txt     ← Python dependencies
├── render.yaml          ← Render deployment config
├── Procfile             ← For Railway / Heroku
└── README.md
```

---

## 🚀 Deploy to Render (Recommended — Free)

### Step 1: Push to GitHub
```bash
# Create a new GitHub repo, then:
git init
git add .
git commit -m "Initial commit"
git remote add origin https://github.com/YOUR_USERNAME/solar-dpr-generator.git
git push -u origin main
```

### Step 2: Deploy on Render
1. Go to **https://render.com** → Sign in / Sign up (free)
2. Click **"New +"** → **"Web Service"**
3. Connect your GitHub account → Select your repo
4. Render auto-detects `render.yaml` — settings are pre-filled:
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `gunicorn app:app --workers 2 --timeout 120`
5. Click **"Create Web Service"**
6. Wait ~2 minutes for first deploy
7. Your app is live at: `https://solar-dpr-generator.onrender.com`

> **Free tier note:** Render free services spin down after 15 min of inactivity
> and take ~30 sec to wake on first request. Upgrade to $7/month for always-on.

---

## 💻 Run Locally

```bash
# Clone / download the project
cd dpr_generator

# Create virtual environment
python -m venv venv
source venv/bin/activate        # Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Run development server
python app.py

# Open browser at:
# http://localhost:5000
```

---

## ✏️ Customising the Template

### Change DPR content / sections
Edit **`dpr_generator.py`** — each section is clearly labelled:
- `SECTION 1` → Company profile
- `SECTION 9` → System description (components)
- `SECTION 12` → Project costing table (cost split percentages)
- etc.

### Add more input fields
1. Add the HTML `<input>` in `templates/index.html`
2. Read it in `app.py` with `request.form.get("field_name")`
3. Pass it into `params` dict
4. Use it in `dpr_generator.py`

### Change fonts / colours in the Word doc
Edit the colour constants at the top of `dpr_generator.py`:
```python
C_NAVY   = RGBColor(0x1F, 0x38, 0x64)
C_BLUE   = RGBColor(0x2E, 0x74, 0xB5)
C_ORANGE = RGBColor(0xC5, 0x5A, 0x11)
```

### Add company logo to cover page
In `dpr_generator.py`, after the cover page title, add:
```python
from docx.shared import Inches
doc.add_picture("static/logo.png", width=Inches(2))
```
Then place your logo at `static/logo.png`.

---

## 📊 What's in the Generated DPR

| Section | Content |
|---|---|
| Cover Page | Project title, company, location, COD |
| Index | All 16 sections listed |
| Sec 1 | Company & group profile |
| Sec 2 | Proposal & project rationale |
| Sec 3 | Power sale concept (PPA) |
| Sec 4 | Solar resource & terminology |
| Sec 5 | Site details & meteorological data |
| Sec 6 | Project at a glance (summary table) |
| Sec 7 | Demand analysis & justification |
| Sec 8 | Benefits of solar PV |
| Sec 9 | System description (8 components) |
| Sec 10 | Bill of Quantity (BOQ) |
| Sec 11 | Project schedule |
| Sec 11A | Regulatory approvals & clearances |
| Sec 11B | O&M plan & risk matrix |
| Sec 12 | Project costing (cost table + funding) |
| Sec 13 | Financial assumptions + 25-year cash flow |
| Sec 14 | Debt repayment schedule & DSCR |
| Sec 15 | Summary of results |
| Sec 16 | Conclusion |
| Annexure A | PVSyst simulation & monthly generation |

---

## 🔧 Minimum Inputs Required

| Field | Example |
|---|---|
| Company Name | XYZ Private Limited |
| Capacity (AC) | 20 MW |
| Location | XX Village, yyy District, Tamil Nadu |
| Project Cost | ₹ 120 Crores |
| PPA Tariff | ₹ 4.80/unit |
| Debt % | 70% |
| Interest Rate | 9.0% p.a. |

All other fields have sensible defaults.

---

## 🆘 Troubleshooting

**App crashes on Render:**  
- Check Render logs (Dashboard → Logs tab)  
- Ensure `requirements.txt` has correct versions  
- Increase `--timeout` in start command if DPR takes >30 sec  

**Download doesn't start:**  
- Check browser pop-up blocker  
- Try a different browser  

**IRR/DSCR not showing:**  
- Click "Recalculate" button manually  
- Check browser console for errors  

---

## 📄 License
Free to use and modify for your projects.
