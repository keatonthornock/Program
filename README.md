*Original Repository at https://github.com/keatonthornock/Program*

# Sunday Service Program Website
This repository is intended to be used as a **GitHub template** for ward-specific Sunday service program websites.

Each new ward repo should keep its own backend settings in a local, ignored config file so future upstream updates are easier to merge.

---

## Getting started

1. On GitHub, click **Use this template** from this repository and create a new repository.
2. Clone your new repository:
   ```bash
   git clone https://github.com/<your-user-or-org>/<your-repo>.git
   cd <your-repo>
   ```
3. Create your local backend config file from the example:
   ```bash
   cp config.example.json config.local.json
   ```
4. Edit `config.local.json` and fill in your own Google Sheets backend values.
5. Run/publish the site as usual (for GitHub Pages, use your `main` branch in repository Settings → Pages).

---

## Config setup

- `config.example.json` is committed to the repository and documents the expected config schema.
- `config.local.json` is **for your ward-specific values** (sheet ID, tab gids, etc.).
- `config.local.json` is intentionally gitignored and should not be committed.

If the local config is missing, the app will fail with an explicit message telling you to copy the example file first.

---

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

---

## Pulling future updates from upstream

After creating your own repo from this template, add the original project as `upstream`:

```bash
git remote add upstream https://github.com/keatonthornock/Program.git
git fetch upstream
git merge upstream/main
```

Because your personal backend settings are in `config.local.json` (ignored by git), upstream merges are less likely to conflict with your local backend configuration.

---

## Calendar set up (optional)

You can automatically pull events from the ward calendar.

1. Open the Google Sheet on desktop.
2. Click **Calendar Sync → Show Calendar Sheets**.
3. Authorize the script.
4. Go to https://www.churchofjesuschrist.org/calendar and copy a sync URL.
5. Paste the ICS link into `CalendarConfig!B1`.
6. Run **Calendar Sync → Sync Calendar Now**.

---

## Potential issues

### Page stuck on "Loading"
- Refresh the page.
- If it still fails, verify that `config.local.json` exists and contains your actual backend values.

---

## Bonus: Install as a phone app

1. Open the website in your browser.
2. Select **Add to Home Screen**.
3. Choose **Web App**.
