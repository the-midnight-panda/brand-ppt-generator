# Brand Intelligence PPT Generator

## Files
- `app.py` — Python backend (handles all research + Claude + PPT)
- `requirements.txt` — Python dependencies
- `Procfile` — Railway start command
- `index.html` — Frontend (host on Replit or anywhere)

## Deploy to Railway (5 minutes, free)

### Step 1 — Push to GitHub
1. Go to github.com → New repository → name it `brand-ppt-generator`
2. Upload these 4 files: app.py, requirements.txt, Procfile, index.html
3. Click "Commit changes"

### Step 2 — Deploy on Railway
1. Go to railway.app → Sign up free
2. Click "New Project" → "Deploy from GitHub repo"
3. Select your `brand-ppt-generator` repo
4. Railway auto-detects Python and deploys

### Step 3 — Add API Keys (Environment Variables)
In Railway dashboard → your project → Variables tab → Add:
- `CLAUDE_API_KEY` = your Claude API key (sk-ant-...)
- `SERPER_API_KEY` = your Serper API key
- `YOUTUBE_API_KEY` = your YouTube API key (AIzaSy...)

### Step 4 — Get your URL
Railway gives you a URL like: `https://brand-ppt-generator-production.railway.app`
Copy it.

### Step 5 — Update index.html
Open index.html, find this line:
```
const BACKEND_URL = "https://YOUR-APP-NAME.railway.app";
```
Replace with your actual Railway URL.

### Step 6 — Host index.html on Replit
Upload the updated index.html to your existing Replit project.
Done! Open your Replit URL, type any brand, get your PPT.

## What it does
- 35+ searches across Instagram, Facebook, YouTube, LinkedIn, Twitter, Amazon,
  Flipkart, Trustpilot, Reddit, Quora, Meta Ads, Google Ads, SEMrush, Inc42,
  YourStory, Startup India, MCA, Zauba, press coverage
- Real YouTube subscriber/view data via YouTube API
- Claude AI deep analysis (no token limits — runs on server)
- python-pptx builds 15-slide premium report server-side
- Returns .pptx download directly

## Why this never breaks
- No browser token limits
- No JSON parsing issues
- API keys safe on server
- PPT built server-side with python-pptx (not browser JS)
- Works for any brand, any size
