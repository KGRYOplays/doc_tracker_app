# AGENTS.md — Document Tracking System

## Stack
- **Backend**: Flask 2.3+ (Python 3.11), single `app.py` with all routes
- **Frontend**: Jinja2 SPA in `templates/index.html` — no JS frameworks (vanilla JS)
- **Database**: SQLite via `db_manager.py` (raw `sqlite3`, no ORM)
- **Sync**: Google Sheets via `gspread` + `google-auth`
- **Deploy**: Render (gunicorn), GitHub `master` branch

## File Structure
```
├── app.py                    # All Flask routes (~2400 lines)
├── db_manager.py             # DB connection, settings KV store, GSheets periodic pull
├── templates/
│   ├── index.html            # Main SPA (~3900 lines: HTML + CSS + JS)
│   └── scanner.html          # Scanner page (~1260 lines)
├── static/
│   ├── favicon.ico           # 32x32 favicon
│   ├── qr_generated/         # Generated barcodes/QRs (gitignored)
│   └── uploads/
│       ├── sdo_manila_logo.png
│       ├── system_seal.svg
│       └── fonts/            # Uploaded custom .ttf/.woff/.otf files
├── app.db                    # SQLite database
├── procfile                  # web: gunicorn app:app ...
├── requirements.txt
└── runtime.txt               # python-3.11
```

## Database Schema

### `code_lookup`
| Column | Type | Purpose |
|--------|------|---------|
| code | TEXT PK | Generated barcode/QR code string |
| doc_type | TEXT | Document type (e.g. "NOSA 2026") |
| date_generated | TEXT | ISO date generated |

### `routing_records`
| Column | Type | Purpose |
|--------|------|---------|
| id | INTEGER PK | Auto |
| code | TEXT | FK to code_lookup.code |
| employee | TEXT | Employee name |
| school_office | TEXT | Origin school/office |
| receiving_office_1..10 | TEXT | Scan offices in sequence |
| timestamp_1..10 | TEXT | Scan timestamps |
| status | TEXT | Current status (case-insensitive) |

### `settings`
| Column | Type | Purpose |
|--------|------|---------|
| key | TEXT PK | Setting name |
| value | TEXT | Setting value |

Key settings: `system_title`, `system_title_font`, `title_letter_spacing`, `title_bg_url`, `title_bg_opacity`, `title_glow_enabled`, `logo_url`, `gsheets_enabled`, `header_line_1/2/3`, `custom_font_url`, `custom_font_format`.

### `users`
| Column | Type | Purpose |
|--------|------|---------|
| id | INTEGER PK | Auto |
| username | TEXT UNIQUE | Login username |
| password | TEXT | SHA256 hash |
| email | TEXT | User email |
| role | TEXT | admin / user / guest / supervisor |
| school_office | TEXT | Assigned school |
| status | TEXT | approved / pending / rejected |
| requires_password_change | INT | 0 or 1 |

## API Routes

### Auth
- `POST /api/login` — authenticate, returns user session data
- `GET /api/check_auth` — session validity check
- `GET /api/get_user_profile` — current user info
- `POST /api/change_password` — update password

### Records (Scanner - POST)
- `POST /api/route_document` — record a scan against a code
- `POST /api/route_document_batch` — batch scan (one code, multiple employees)
- `POST /api/update_record_admin` — admin edit of a record

### Records (GET)
- `GET /api/get_records` — paginated records list
- `GET /api/get_all_codes` — paginated generated codes
- `GET /api/get_record/<code>` — single record detail

### Generator
- `POST /api/generate_barcode` — generate single code
- `POST /api/generate_batch` — batch code generation

### Dashboard
- `GET /api/dashboard_stats?count_mode=schools|rows&doc_type=...` — stat counts, status breakdown, doc_type breakdown
- `GET /api/recent_activity?count_mode=schools|rows&doc_type=...` — last 10 activity rows

### Admin
- `GET /api/get_users` — list all users
- `POST /api/approve_user` — approve user
- `POST /api/reject_user` — reject user
- `POST /api/delete_user` — remove user
- `POST /api/register` — sign up new user
- `GET /api/get_settings` — all branding/settings values
- `POST /api/save_settings` — save branding/settings
- `POST /api/upload_font` — upload custom font file
- `POST /api/remove_font` — remove custom font
- `POST /api/upload_title_bg` — upload title background image
- `POST /api/remove_title_bg` — remove title background
- `POST /api/upload_logo` — upload agency logo
- `POST /api/remove_logo` — remove agency logo
- `GET /api/gsheets_last_pull` — periodic pull status

### Google Sheets
- `POST /api/test_gsheets_connection` — test GSheets connection
- `POST /api/bulk_sync_sheets` — manual full sync
- `POST /api/export_local_db` — export DB to JSON
- `POST /api/import_local_db` — import DB from JSON

## Frontend Conventions

### Colors / Theme (CSS variables in `:root`)
- `--bg-primary`, `--surface`, `--text-main`, `--text-muted`, `--primary`
- `--glass-bg`, `--glass-border`, `--radius-md`, `--btn-transition`
- All via CSS custom properties, no preprocessor

### Layout
- `.container` — max 1800px centered wrapper
- `.glass-card` — frosted glass card for each section
- `.card-title-row` — icon + heading row
- `.branding-grid` — 2-column grid for branding settings
- `.grid-4` — 4-column grid for stat cards
- `.dashboard-filters` — filter bar with dropdown + count mode toggle

### Tabs (role-based)
| Role | Visible Tabs |
|------|-------------|
| admin | Dashboard, Generator, Records, All Codes, Admin Tools |
| user | Dashboard, Records |
| guest | Records only |
| supervisor | Dashboard, Records |

Default tab: Dashboard (admin/user), Records (guest).

### Dashboard
- Stat cards: Total Barcodes, Active Documents, Released, Pending Signature
- Donut chart: CSS conic-gradient (no library), centered count in hole
- Doc type breakdown: styled list with colored dots
- Recent activity table
- Global filter: doc_type dropdown + count mode toggle (Per School/Office ↔ Per Employee)

### Buttons
- `.btn` — no default width (auto-sizes to content)
- `.btn-primary` — accent background
- `.btn-secondary` — outline style
- `.btn-add-sheet` — green accent
- Inline `style="width:100%"` only on full-width buttons inside modals/login

### Forms
- Inputs/selects/textareas default to `max-width: 600px`
- `.form-group` — wraps label + input with 18px margin-bottom
- `.switch` is a custom CSS toggle for checkboxes

## Special Features

### System Branding
- Title rendered as per-letter `<span>`s for hover animation (lift, recolor, glow)
- Letter-spacing via `--title-letter-spacing` CSS variable
- Glow/shadow toggle via `.no-glow` class
- Watermark background on title bar via `--title-bg-url` + `--title-bg-opacity`
- Google Fonts dropdown (20 curated options) + custom uploaded font

### Status Auto-Release
When scanned receiving office matches `'RECORDS SERVICES'`, status is set to `'released'` automatically. All status matching uses `LOWER()` for case-insensitive comparison.

### Count Mode
- **Per School/Office**: `COUNT(DISTINCT school_office || '|' || doc_type)` — one per (school, doc_type) pair
- **Per Employee**: `COUNT(*)` — each record counts

### Google Sheets Sync
- Periodic pull every 5 minutes (configurable via settings)
- `db_manager._last_pull_info` tracks state: syncing / ok(idle) / error(msg)
- Credentials and sheet ID set via environment variables (`GSHEETS_CREDENTIALS`, `GSHEETS_ID`)

## Deployment

- **Platform**: Render
- **URL**: `https://dashboard.render.com/web/srv-d876ogf7f7vs73d982l0`
- **Deploy Hook**: `POST https://api.render.com/deploy/srv-d876ogf7f7vs73d982l0?key=_kpiu-SL8iQ`
- **Branch**: `master`
- **Procfile**: `web: gunicorn app:app --bind 0.0.0.0:$PORT --workers=2 --timeout=120`
- **Static files**: Flask auto-serves `/static/`, assets must be committed for Render deployment

## AGENTS / Skills

### For AI Coding Assistants

When working on this codebase:

1. **Read before edit** — always read the file first before making changes
2. **Match style** — use the same patterns (inline styles, no JS frameworks, `var` in JS, snake_case Python)
3. **No comments** — don't add explanatory comments unless the user asks
4. **Minimize changes** — prefer editing existing patterns over restructuring
5. **Test before deploy** — verify API responses work locally before committing
6. **Both templates** — changes to UI often need updates in both `index.html` and `scanner.html`
7. **CSS vars** — prefer CSS variables over hardcoded colors
8. **Status matching** — always use `LOWER()` for status comparisons
9. **Settings** — use `db_manager.get_setting(key)` / `save_setting(key, value)` for persistence
10. **Periodic pull** — if modifying GSheets sync, update `_last_pull_info` with thread lock
