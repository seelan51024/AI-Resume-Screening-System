# AI-Based Resume Screening System

ML-powered HR tool — FastAPI backend + single-page frontend + SQLite database.

## Quick Start

```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Start (auto-trains models if needed, opens on http://localhost:8000)
python run.py
```

Open your browser → **http://localhost:8000**

## Project Structure

```
AI-Resume-Screener/
├── run.py               ← START HERE (auto-trains models, launches server)
├── main.py              ← FastAPI backend (API + SQLite + auth + email + export)
├── index.html           ← Frontend SPA (served by FastAPI at /)
├── requirements.txt     ← Python dependencies
├── README.md
├── ml/
│   └── train_model.py   ← ML training script (auto-run by run.py if needed)
└── models/              ← Trained model files
    ├── random_forest.pkl
    ├── decision_tree.pkl
    ├── label_encoder.pkl
    └── feature_importances.json
```

> **SQLite database** (`resume_screener.db`) is auto-created in the project root on first run.

## Default Credentials

| Role    | Email                   | Password    |
|---------|-------------------------|-------------|
| HR      | hr@company.com          | hr123456    |
| Manager | manager@company.com     | mgr123456   |

## API Endpoints

| Method | Path                        | Auth    | Description                     |
|--------|-----------------------------|---------|---------------------------------|
| GET    | `/`                         | —       | Serves the frontend (index.html)|
| GET    | `/health`                   | —       | Server + DB status              |
| POST   | `/auth/login`               | —       | JWT login                       |
| POST   | `/auth/register`            | —       | Register new user (admin key)   |
| GET    | `/auth/me`                  | JWT     | Current user info               |
| POST   | `/screen`                   | HR      | Screen single resume            |
| POST   | `/screen-batch`             | HR      | Screen multiple resumes         |
| GET    | `/rankings`                 | JWT     | Batch rankings from DB          |
| GET    | `/batch-sessions`           | JWT     | List screening sessions         |
| POST   | `/send-email`               | HR      | Send feedback email             |
| GET    | `/export/excel`             | JWT     | Download Excel rank list        |
| GET    | `/stats`                    | JWT     | Screening statistics            |
| DELETE | `/results`                  | Manager | Clear all records               |

## Database

SQLite file: `resume_screener.db` (auto-created, no setup needed)

Tables:
- `users` — HR/Manager accounts with hashed passwords
- `candidates` — All screened resume records with scores, labels, skills

## Email Config

Edit `main.py` → `EMAIL_CONFIG` section with your Gmail SMTP details.
Use a Gmail App Password (Settings → Security → App Passwords).

## Bug Fixes Applied

1. `requirements.txt` — added `email-validator` (needed by Pydantic `EmailStr`) + pinned `bcrypt==4.0.1` (passlib compatibility)
2. `clearAll()` — fixed inverted role check (managers clear records, not HR)
3. Batch screen — proper error JSON parsing (was showing generic error always)
4. Status dot — shows amber in demo mode, green in live mode
5. Email modal — XSS fix: names with quotes no longer break onclick handlers
6. Frontend API URL — now auto-detects same-origin (no hardcoded localhost needed)
7. Login timeout — increased from 3s → 8s (prevents false demo-mode on slow starts)
8. Static serving — `index.html` served directly by FastAPI at `/` (no separate web server)
