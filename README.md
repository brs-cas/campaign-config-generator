# Campaign Config Generator

Internal tool for the onsite merchandising team to generate pre-populated Excel spreadsheet templates for sale campaigns.

## Problem

Every sale campaign requires configuring 6–8 touchpoints across up to 5 markets (US, AU, SG, UK, CA) in two systems: **Storyblok** (CMS) and **Dynamic Yield** (personalization platform).

- **BAU sales** (single phase): ~2–3 hours of manual setup
- **Peak sales** (4 phases: Early Access → Regular → Upsize → Dash): 4x a BAU sale, since every touchpoint must be configured separately per phase

This tool generates a campaign config spreadsheet so the team has a single checklist of every touchpoint × phase × market combination, instead of doing it from memory.

## Project Structure

```
campaign-config-generator/
├── backend/
│   ├── main.py              # FastAPI app
│   ├── generator.py         # Spreadsheet generation logic
│   ├── config/
│   │   ├── touchpoints.json       # Configurable touchpoint list
│   │   └── market_defaults.json   # Market-specific defaults
│   └── requirements.txt
├── frontend/                # (coming soon)
├── templates/               # Saved campaign templates
├── .gitignore
└── README.md
```

## Setup

```bash
cd backend
pip install -r requirements.txt
uvicorn main:app --reload
```

## API Endpoints

| Method | Path | Description |
|--------|------|-------------|
| GET | `/health` | Health check |
| POST | `/generate` | Generate campaign config spreadsheet (.xlsx) |
| GET | `/templates` | List saved templates |
| GET | `/templates/{filename}` | Get a specific template |
| POST | `/templates` | Save a template |
| GET | `/config/touchpoints` | Get touchpoint definitions |
| GET | `/config/markets` | Get market defaults |

## Campaign Types

- **BAU**: Single phase, one set of content. Generates 8 rows per market.
- **Peak**: 4 phases (Early Access, Regular, Upsize, Dash), each with different content. Generates 32 rows per market.

## Markets

US, AU, SG, UK, CA — each with market-specific defaults for currency, timezone, sale page URL, promo copy, and notification bar styling.
