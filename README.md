*Original Repository at https://github.com/keatonthornock/Program*

# Sunday Service Program Website
Version: 1.0.0

This repository is intended to be used as a **GitHub template** for ward-specific Sunday service program websites.

## Quick setup (GitHub website only)

You can set up your site entirely in the browser—no terminal required.

1. In this repository, click **Use this template**.
2. Create your new repository.
3. In your new repository, open **`config.json`**.
4. Click the **pencil (Edit this file)** icon.
5. Replace `sheet_id` with your Google Sheet ID.
6. Update the tab `gid` values if needed (`admin_gid`, `agenda_gid`, etc.).
7. Commit the file changes.
8. Go to **Settings → Pages** and enable GitHub Pages (deploy from your main branch).

Once Pages is enabled, your site will use your committed `config.json` values.

## How to find your Google Sheet ID

Given a sheet URL like:

`https://docs.google.com/spreadsheets/d/1AbCdEfGhIjKlMnOpQrStUvWxYz1234567890/edit#gid=0`

The **sheet ID** is the part between `/d/` and `/edit`:

`1AbCdEfGhIjKlMnOpQrStUvWxYz1234567890`

Paste that value into `config.json` as `sheet_id`.

## Config file notes

- `config.json` is the main editable config file for template users.
- `config.example.json` is optional reference documentation for the expected keys.
- If `config.json` is missing or invalid JSON, the app will show a clear error.

## Hymn link lookup file

- `data/hymn-links.json` is an optional static lookup used for exact hymn deep links.
- Collections currently supported:
  - `hymns`
  - `childrens_songbook`
  - `hymns_for_home_and_church`
- Add entries with `id`, `title`, and `url` to make a hymn card resolve to an exact destination instead of a fallback guess.

## Setting up the backend Google Sheet

1. Create a copy of the template sheet into your Google Drive:  
   https://docs.google.com/spreadsheets/d/1I_Mj-ZoW57cR5PpBoMoBa2tVqpkfrlXOZd4WK99HsD4/edit
2. In Google Sheets, go to **File → Share → Publish to web**.
3. Publish the **Administrative** tab.
4. In published content settings, also publish:
   - Agenda
   - Announcements
   - Calendar
   - Ward Leadership
5. Do **not** publish:
   - CalendarConfig
   - Hymn Directory
6. In **File → Share → Share**, set General access to **Anyone with the link → Viewer**.

## Updating your site later

- This template now uses a manually edited, committed `config.json`.
- When new template versions are released, compare your README version and files with the latest template version.
- Copy in updated template files as needed, but avoid overwriting your own `config.json` values.

## Calendar set up (optional)

You can automatically pull events from the ward calendar.

1. Open the Google Sheet on desktop.
2. Click **Calendar Sync → Show Calendar Sheets**.
3. Authorize the script.
4. Go to https://www.churchofjesuschrist.org/calendar and copy a sync URL.
5. Paste the ICS link into `CalendarConfig!B1`.
6. Run **Calendar Sync → Sync Calendar Now**.

## Bonus: Install as a phone app

1. Open the website in your browser.
2. Select **Add to Home Screen**.
3. Choose **Web App**.
