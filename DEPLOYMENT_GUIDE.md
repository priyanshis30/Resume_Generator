# ResuméCraft — Complete Deployment Guide

> Built by Priyanshi · Devanshi · Bhavya · Pahul · Purva  
> Mechanical & Automation Engineering · IGDTUW

---

## Project Structure

```
resumecraft/
├── frontend/
│   └── index.html            ← Complete app (hosted on GitHub Pages)
├── backend/
│   ├── server.js             ← Express API (hosted on Render)
│   ├── package.json
│   ├── render.yaml           ← One-click Render config
│   ├── .env.example          ← Copy to .env, fill in secrets
│   ├── .gitignore
│   └── data/                 ← Auto-created; holds feedback.xlsx
└── README.md
```

---

## API Endpoints (once deployed)

| Method | URL | Auth | What it does |
|--------|-----|------|-------------|
| GET  | `/` | None | Health check |
| POST | `/api/feedback` | None | Save one submission |
| GET  | `/api/stats` | None | Public JSON stats |
| GET  | `/api/admin?token=…` | Token | HTML dashboard |
| GET  | `/api/download-excel?token=…` | Token | Download .xlsx file |
| GET  | `/api/all-feedback?token=…` | Token | All rows as JSON |

---

## STEP 1 — Get a Gmail App Password

You need this so the backend can send you email alerts.

1. Open **myaccount.google.com**
2. Security → **2-Step Verification** (turn it ON if not already)
3. Search "App passwords" → click it
4. App name: type `ResuméCraft` → click **Generate**
5. **Copy the 16-character password** — you'll paste it in Step 3

> The email will be SENT FROM your Gmail, and delivered TO `mujhekyameintohbatakhoon@gmail.com`

---

## STEP 2 — Push to GitHub

Create **two separate GitHub repos**:

### Repo 1: Frontend
```bash
cd resumecraft/frontend
git init
git add index.html
git commit -m "Initial ResuméCraft frontend"
git branch -M main
git remote add origin https://github.com/YOUR-USERNAME/resumecraft-frontend.git
git push -u origin main
```

### Repo 2: Backend
```bash
cd ../backend
git init
git add .
git commit -m "Initial ResuméCraft backend"
git branch -M main
git remote add origin https://github.com/YOUR-USERNAME/resumecraft-backend.git
git push -u origin main
```

> ⚠️ `.gitignore` already excludes `.env` and `data/` — secrets and Excel data stay off GitHub.

---

## STEP 3 — Deploy Backend on Render (free)

1. Go to **https://render.com** → sign up (free, no card needed)
2. Click **New → Web Service**
3. Connect your **resumecraft-backend** GitHub repo
4. Render auto-detects Node. Confirm:
   - **Build Command**: `npm install`
   - **Start Command**: `node server.js`
5. Scroll to **Environment Variables** → click **Add Environment Variable** for each:

   | Key | Value |
   |-----|-------|
   | `GMAIL_USER` | your_gmail@gmail.com |
   | `GMAIL_APP_PASSWORD` | the 16-char app password from Step 1 |
   | `DOWNLOAD_TOKEN` | any secret string e.g. `igdtuw_resumecraft_2026` |
   | `NODE_ENV` | `production` |

6. Click **Create Web Service**
7. Wait ~2 min → Render gives you a URL like:  
   **`https://resumecraft-backend.onrender.com`** — copy this!

> 💡 Free tier sleeps after 15 min of inactivity. First request wakes it (~30s). That's fine.

---

## STEP 4 — Connect Frontend to Backend

Open `frontend/index.html` and find this line (near top of `<script>`):

```javascript
const BACKEND_URL = 'https://YOUR-BACKEND-URL.onrender.com';
```

Replace with your actual Render URL:

```javascript
const BACKEND_URL = 'https://resumecraft-backend.onrender.com';
```

Commit and push:
```bash
cd frontend
git add index.html
git commit -m "Connect frontend to backend"
git push
```

---

## STEP 5 — Deploy Frontend on GitHub Pages (free)

1. Open your **resumecraft-frontend** repo on GitHub
2. Go to **Settings → Pages**
3. Source: **Deploy from a branch**
4. Branch: `main` · Folder: `/ (root)`
5. Click **Save**
6. After ~1 minute, your site is live at:  
   **`https://YOUR-USERNAME.github.io/resumecraft-frontend/`**

---

## Using the Admin Dashboard

Once deployed, open:
```
https://resumecraft-backend.onrender.com/api/admin?token=igdtuw_resumecraft_2026
```

You'll see:
- 📊 Total submissions, average rating, placed count
- 📋 Full table of all feedback (newest first)
- ↓ Download Excel button
- {} View as JSON button

### Download the Excel sheet:
```
https://resumecraft-backend.onrender.com/api/download-excel?token=igdtuw_resumecraft_2026
```

### Email alerts:
Every submission sends a styled email to `mujhekyameintohbatakhoon@gmail.com` with name, email, stream, hired status, star rating, and message.

---

## Local Development

```bash
cd backend
cp .env.example .env      # Fill in GMAIL_USER, GMAIL_APP_PASSWORD, DOWNLOAD_TOKEN
npm install
npm run dev               # Starts on http://localhost:3001

# In index.html, temporarily change:
# const BACKEND_URL = 'http://localhost:3001';

# Open frontend/index.html in browser
```

Test it:
```bash
curl -X POST http://localhost:3001/api/feedback \
  -H "Content-Type: application/json" \
  -d '{"name":"Test User","email":"test@test.com","stream":"CS","hired":"Just exploring","rating":5,"message":"Works great!"}'
```

Then open: `http://localhost:3001/api/admin?token=YOUR_TOKEN`

---

## Troubleshooting

| Problem | Fix |
|---------|-----|
| Feedback not saving | Check Render logs → Dashboard → your service → Logs tab |
| Email not arriving | Verify GMAIL_APP_PASSWORD is the App Password (not your Gmail login). Check spam folder. |
| CORS error in console | Make sure BACKEND_URL in index.html matches your Render URL exactly (no trailing slash) |
| GitHub Pages 404 | `index.html` must be in root of `main` branch |
| Render deploy fails | Check Build logs — usually a missing package or Node version issue |
| `data/` folder missing | Server auto-creates it on first run — not a problem |

---

## What Gets Saved

Each submission stores in Excel with these columns:

| Timestamp | Name | Email | Stream | Hired Status | Rating (/5) | Message |
|-----------|------|-------|--------|--------------|-------------|---------|
| 14/6/2026, 3:45 PM | Aanya Sharma | aanya@... | CS | Yes! Got placed 🎉 | 5 | Amazing tool! |

Rows are zebra-striped, header is dark with gold border, rating cells are gold.

---

*Every student who dares to dream deserves a great resume. 💛*
