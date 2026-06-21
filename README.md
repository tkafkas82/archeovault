# ArcheoVault

Archaeological warehouse inventory management — a single HTML file backed by Google Sheets.

Built for field teams that need a shared, mobile-friendly catalog of artifacts stored across shelves and crates, without setting up servers or databases.

## How It Works

```
index.html  ──HTTP──▶  Google Apps Script  ──▶  Google Sheets
(browser)              (serverless API)         (database)
```

One `index.html` file does everything. Google Sheets stores the data. Google Apps Script acts as the API layer. No backend to deploy or maintain.

## Features

- **Artifact catalog** — name, category, location (shelf/crate), description, photos
- **Hierarchical categories** — unlimited nesting (e.g. Κεραμικά > Αμφορείς > Τύπος Α)
- **20 stock illustrations** — built-in SVG icons for common artifact types (amphora, coin, statue, mosaic, etc.) — or use your own Google Drive photos
- **Search & filter** — by category (includes subcategories), location, or free text
- **Google Sign-In** — editor/read-only access control via an Editors list in the spreadsheet
- **Persistent login** — stay signed in across sessions (localStorage)
- **Export/Import** — full JSON backup and restore (editor-only)
- **PWA** — installable on mobile home screens, works like a native app
- **Responsive** — desktop, tablet, and mobile layouts
- **Greek UI** — interface in Greek, suitable for Hellenic archaeological projects
- **Single file** — all CSS, JS, SVGs, manifest, and service worker are inline. No build step, no dependencies.

## Quick Start

### 1. Create the Google Sheet

Create a new Google Sheet with three tabs:

**Κατηγορίες** (Categories)

| id | name | parentId |
|----|------|----------|

**Αντικείμενα** (Items)

| id | name | category | location | description | photos | createdAt | updatedAt |
|----|------|----------|----------|-------------|--------|-----------|-----------|

**Editors**

| email |
|-------|
| your@gmail.com |

> The app auto-creates these tabs if they don't exist.

### 2. Install the Apps Script

1. In Google Sheets: **Extensions > Apps Script**
2. Delete the default code, paste the contents of `apps-script.js`
3. **Deploy > New deployment > Web app**
   - Execute as: **Me**
   - Who has access: **Anyone**
4. Copy the deployment URL

### 3. Set Up Google Sign-In

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a project, configure the OAuth consent screen
3. Create an **OAuth 2.0 Client ID** (Web application)
4. Add your hosting domain to **Authorized JavaScript origins**
5. Copy the Client ID

### 4. Configure & Deploy

1. Open `index.html` in a text editor
2. Set the two constants at the top:
   ```javascript
   const API_URL = 'https://script.google.com/macros/s/.../exec';
   const GOOGLE_CLIENT_ID = '...apps.googleusercontent.com';
   ```
3. Host the file (Netlify, GitHub Pages, or any static host)

> Detailed step-by-step instructions in Greek: see [SETUP.md](SETUP.md)

## Access Control

| Role | Can do |
|------|--------|
| **Anyone** (no sign-in) | View items, search, filter |
| **Editor** (signed in + listed in Editors tab) | Add/edit/delete items and categories, export/import, manage subcategories |

Add editor emails in the **Editors** tab of the Google Sheet (one per row).

## Photos

Two options per item:

- **Stock illustrations** — pick from 20 built-in archaeological SVG icons (amphora, vase, coin, helmet, sword, mosaic, etc.)
- **Google Drive links** — paste a shareable Drive link; the app converts it to a thumbnail automatically

## File Structure

```
warehouse-stock/
├── index.html        # The entire application (single file)
├── apps-script.js    # Google Apps Script backend (paste into Script Editor)
├── SETUP.md          # Setup instructions in Greek
└── README.md         # This file
```

## Updating the Apps Script

After modifying `apps-script.js`:

1. Open Apps Script (Extensions > Apps Script in the Sheet)
2. Replace the code
3. **Deploy > Manage deployments > Edit (pencil) > Version: New version > Deploy**

The URL stays the same — no changes needed in `index.html`.

## Tech Stack

- Vanilla HTML/CSS/JS (no frameworks, no build tools)
- Google Sheets (database)
- Google Apps Script (serverless API)
- Google Identity Services (authentication)
- Inline PWA (manifest + service worker via blob URLs)
- Inline SVG illustrations

## License

MIT
