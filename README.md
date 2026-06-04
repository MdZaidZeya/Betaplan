# Beta Plan 2026 — Cloud Dashboard

A Flask web app that Rishab uploads `Beta_Plan.xlsx` to, and both Rishab and Jim
can see the live progress chart at a shared URL — no local server needed.

---

## How it works

| Who      | What they do                                                            |
|----------|-------------------------------------------------------------------------|
| Rishab   | Edits `Beta_Plan.xlsx` as normal, then opens the dashboard and uploads  |
| Jim      | Bookmarks the URL and refreshes any time to see the latest chart        |
| Everyone | The chart + metrics persist on the server — no one needs to be online   |

---

## Deploy to Render (free, ~10 min)

### Step 1 — Push to GitHub

1. Go to [github.com](https://github.com) → **New repository**
2. Name it `betaplan-dashboard`, set it to **Private** (recommended), click **Create**
3. Upload all files from this folder:
   - `app.py`
   - `requirements.txt`
   - `render.yaml`
   - `templates/index.html`

   Easiest way — drag-and-drop all files into the GitHub web UI after creating the repo.

### Step 2 — Create a Render account

Go to [render.com](https://render.com) and sign up with your GitHub account.

### Step 3 — Create a new Web Service

1. Dashboard → **New** → **Web Service**
2. Connect your GitHub account if prompted
3. Select the `betaplan-dashboard` repo → **Connect**
4. Fill in the settings:

   | Field           | Value                                                    |
   |-----------------|----------------------------------------------------------|
   | **Name**        | `betaplan-dashboard`                                     |
   | **Runtime**     | `Python 3`                                               |
   | **Build Command** | `pip install -r requirements.txt`                      |
   | **Start Command** | `gunicorn app:app --bind 0.0.0.0:$PORT --workers 1 --timeout 120` |

5. Under **Environment** → add:
   - Key: `DATA_DIR` → Value: `/var/data/betaplan`

6. Under **Disks** (scroll down) → **Add Disk**:
   - Name: `betaplan-data`
   - Mount Path: `/var/data/betaplan`
   - Size: `1 GB`
   *(This is what makes the chart persist — free tier includes 1 GB)*

7. Click **Create Web Service**

Render builds and deploys in ~3 minutes. You'll get a URL like:
```
https://betaplan-dashboard.onrender.com
```

### Step 4 — Share with Jim

Send Jim that URL on Jira. He bookmarks it and refreshes any time.

> **Note:** Free Render services spin down after 15 min of inactivity and take ~30 sec to wake up on first visit. That's fine for weekly updates. If you want it always-instant, upgrade to Render's $7/month plan.

---

## Weekly workflow (Rishab)

1. Edit `Beta_Plan.xlsx` as normal
2. Open `https://betaplan-dashboard.onrender.com`
3. Click **Update Sheet** tab
4. Drop in the updated xlsx
5. Done — Jim can see it immediately

---

## Local development

```bash
pip install -r requirements.txt
python app.py
# → http://localhost:5050
```

---

## File structure

```
betaplan-dashboard/
├── app.py                  # Flask backend (parse xlsx, generate chart, persist)
├── requirements.txt        # Python deps
├── render.yaml             # Render one-click deploy config
├── templates/
│   └── index.html          # Dashboard UI (viewer + uploader tabs)
└── README.md               # This file
```

---

## API endpoints

| Endpoint         | Method | Description                                      |
|------------------|--------|--------------------------------------------------|
| `/`              | GET    | Dashboard HTML (Jim's view + Rishab's upload tab)|
| `/upload`        | POST   | Accepts xlsx → parses → persists chart + summary |
| `/latest-chart`  | GET    | Serves the latest PNG (direct download link)     |
| `/latest-summary`| GET    | JSON summary of latest data                      |
| `/chart`         | POST   | Renders chart from posted rows (used inline)     |
