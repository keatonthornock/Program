*Original Repository at https://github.com/keatonthornock/Program*

# Sunday Service Program Website
Create a simple Sunday Service program webpage for your ward. The administrator only needs to maintain a Google Sheet — the website will automatically display the information.

---

# Quick Setup Version - For The Common GitHub User
1. Copy the Google Sheet template, publish as webpage. For which tabs to make public, see point 4 and 5 below.
2. For calendar set up (optional but cool) see that section below.
3. Click **Use this template** on this repository.
4. Paste your sheet ID into `config.json`.
5. Enable GitHub Pages.

---

# Preparation Considerations

### URL Appearance
Your website will be hosted at:

https://[github-username].github.io/[repository-name]

If you want the URL to reflect your ward name, create the GitHub account using a ward-appropriate name.

Example:

https://pleasantgroveward.github.io/sunday-program

---

# Implementation Steps

## 1. Setting Up the Backend Google Sheet

1. Create a copy of the template sheet into your Google Drive:  
https://docs.google.com/spreadsheets/d/1I_Mj-ZoW57cR5PpBoMoBa2tVqpkfrlXOZd4WK99HsD4/edit

2. In the sheet go to:

File → Share → Publish to web

3. Change **Entire Document → Administrative** and click **Publish**.

4. Expand **Published content settings** and also publish these tabs:

- Agenda  
- Announcements  
- Calendar  
- Ward Leadership  

Do **NOT** publish:

- CalendarConfig  
- Hymn Directory

5. Now enable public viewing:

File → Share → Share

Set:

General access → **Anyone with the link → Viewer**

Admins can still be added with edit access.

---

## 2. Copying the GitHub Repository

1. Click **Use this template** at the top of this repository.
2. Select **Create a new repository**.

Settings:

- **Repository Name** → the name of your site  
- **Visibility** → **Public** (required for GitHub Pages)

Click **Create repository**.

---

## 3. Linking the GitHub Repository to the Google Sheet Backend

1. Open your new repository.
2. Open the file **config.json**.
3. Click the **edit (pencil) icon**.

Replace the `sheet_id` value with your Google Sheet ID.

Example sheet URL:

https://docs.google.com/spreadsheets/d/ABCDEFGHIJK123456/edit#gid=0

Your sheet ID is:

ABCDEFGHIJK123456

Paste this value inside `config.json`.

Do **not** change the other values.

Click **Commit changes**.

---

## 4. Publishing the Webpage

1. In your repository open:

Settings → Pages

2. Under **Branch**, select:

main

3. Click **Save**.

After about a minute GitHub will show:

Your site is live at:  
https://[username].github.io/[repository-name]

Your website is now live.

Any updates made in the Google Sheet will appear on the site after refreshing the page.

---

# Calendar Set Up (Optional)

You can automatically pull events from the ward calendar.

If you skip this step, the calendar tab will simply link to your ward website.

## Enable Calendar Sync

1. Open the Google Sheet on a **desktop**.
2. Click:

Calendar Sync → Show Calendar Sheets

3. Authorize the script when prompted.

## Get Your Ward Calendar Link

1. Go to:  
https://www.churchofjesuschrist.org/calendar

2. Click the **settings icon → Sync**.

3. Either:

- Copy the **Auto-synced calendars URL**, or  
- Create a custom calendar group and copy its URL.

## Add the Calendar Link

1. Open the **CalendarConfig** tab in the Google Sheet.
2. Paste the ICS link into **cell B1**.

You can also adjust:

- `LOOKAHEAD_DAYS`
- `INCLUDE_PAST_DAYS`

## Sync Events

Open:

Calendar Sync → Sync Calendar Now

Events will populate the **Calendar** tab.

Use the **checkbox column** to control which events appear on the website.

You can also enable automatic syncing from the **Calendar Sync** menu.

---

# Potential Issues

### Page stuck on "Loading"
Refresh the page.  
If it still fails, close and reopen the browser.

---

# Bonus: Install as a Phone App

You can add the site to your phone home screen so it behaves like an app.

1. Open the website in your browser.
2. Select **Add to Home Screen**.
3. Choose **Web App**.

An icon will appear on your home screen and the site will open like an app.
