*Original code Repository at https://github.com/keatonthornock/program*

*NOTE: if opening the example below on a desktop browser you can still view how it appears on mobile. Right click after opening the site → Inspect → Toggle Device Emulation (in the tool bar of dev tools).*

Example (live demo site): https://keatonthornock.github.io/program

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

## Setting up the backend Google Sheet

1. Create a copy of the template sheet into your Google Drive:  
   https://docs.google.com/spreadsheets/d/144P85pf3sp6_yxGTZIALg1nFahk8BgX-oxdM-LUPTV8/edit?gid=0#gid=0
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

## Calendar set up (optional)

You can automatically pull events from the ward calendar.

1. Open the Google Sheet on desktop.
2. Click **Calendar Sync → Show Calendar Sheets**.
3. Authorize the script.
4. Go to https://www.churchofjesuschrist.org/calendar and copy a sync URL.
5. Paste the ICS link into `CalendarConfig!B1`.
6. Run **Calendar Sync → Sync Calendar Now**.
7. For auto-syncing, click **Calendar Sync → Set Trigger: Every [x] Hour(s)**.
8. If you wish to stop this auto-syncing in the future, click **Calendar Sync → Stop Frequency Trigger**.

## Updating your site later if a new version releases

- This template uses a manually edited, committed `config.json`.
- When new template versions are released you have two options from simplest to most intensive:
1. **Recommended**: Delete your respository but not the Google Sheet Back-end. Copy the newest version as a new repository, and give it the same exact name you did before. This will result in your link being the same link it had prior, which means the QR code will still being functional (in case you had those printed off you needn't worry). Then all you will need to do is replace the single line in the `config.json` file again and publish the newly copied repository. This way you won't need to dig through every file to see what was changed.
2. Compare your README version and files with the latest template version. Then copy in updated template files as needed, but avoid overwriting your own `config.json` values.

## Bonus: Install as a phone app

1. Open the website in your browser.
2. Select **Add to Home Screen**.
3. Choose **Web App**.
