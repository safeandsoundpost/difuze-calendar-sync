# Calendar Sync GUI - Quick Start Guide

## What This Does

This app syncs your Outlook calendar PDF to Google Calendar with a single click. 

**Important:** It **REPLACES** the events in your Google Calendar (clears old ones first, then imports new ones) - so you won't get duplicates!

## One-Time Setup

### 1. Install Python Packages

Open Terminal and run:
```bash
pip3 install pdfplumber google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client icalendar pytz
```

### 2. Set Up Google Calendar API

#### Get Your Credentials:

1. Go to [Google Cloud Console](https://console.cloud.google.com)
2. Create a new project
3. Enable **Google Calendar API**:
   - APIs & Services > Library > Search "Google Calendar API" > Enable
4. Create credentials:
   - APIs & Services > Credentials > Create Credentials > OAuth client ID
   - Configure consent screen if needed (External, add safeandsoundpost@gmail.com as test user)
   - Application type: **Desktop app**
   - Download the JSON file
   - **Rename it to `google_credentials.json`**
   - Put it in the same folder as `calendar_sync_gui.py`

## How to Use

### First Time:

1. **Export from Outlook:**
   - Open Outlook > Calendar > "Toronto Post"
   - Print > Detailed Agenda > Select next 2 months
   - Save as PDF on your Desktop with filename: `outlook_agenda.pdf`

2. **Run the app:**
   ```bash
   python3 calendar_sync_gui.py
   ```

3. **Configure:**
   - PDF Folder: Browse to your Desktop (or wherever you saved the PDF)
   - Google Calendar: Enter `DIFUZE - STUDIO CALENDAR`

4. **First sync:**
   - Click **SYNC NOW**
   - A browser will open asking you to sign in to Google
   - Sign in with **safeandsoundpost@gmail.com**
   - Grant permissions
   - The app will sync!

### Every Time After:

1. Export fresh PDF from Outlook (save as `outlook_agenda.pdf` in the same folder)
2. Open the app
3. Click **SYNC NOW**
4. Done!

## Important Notes

### About Replacing Events

**YES - it replaces old events!** The app:
1. Looks at the date range in your PDF
2. Deletes all events in that date range from your Google Calendar
3. Imports the fresh events from the PDF

This means **no duplicates** - each sync gives you a clean, up-to-date calendar.

### PDF File Name

The app looks for a file named exactly: `outlook_agenda.pdf`

If you want to use a different name, edit the script and change this line:
```python
pdf_path = os.path.join(self.pdf_folder.get(), 'outlook_agenda.pdf')
```

### Calendar Name

Make sure your Google Calendar is named exactly: `DIFUZE - STUDIO CALENDAR`

You can check this in Google Calendar > Settings > Find your calendar name

## Troubleshooting

### "PDF file not found"
- Make sure the PDF is in the folder you selected
- Make sure it's named `outlook_agenda.pdf`

### "Calendar not found"
- Check the exact name of your calendar in Google Calendar
- It's case-sensitive: `DIFUZE - STUDIO CALENDAR`

### "No events found"
- Make sure you printed as "Detailed Agenda" not monthly view
- Open the PDF and verify events are listed

### Google Authentication Issues
- Make sure `google_credentials.json` is in the same folder as the script
- Delete `token.pickle` and try again to re-authenticate

## Tips

- **Monthly workflow:** Export PDF from Outlook at the start of each month, sync with one click
- **The app remembers:** Your folder path and calendar name are saved for next time
- **Check the log:** The bottom section shows exactly what happened during sync
- **Desktop shortcut:** Create a shortcut to `calendar_sync_gui.py` for quick access

## Creating a Desktop Shortcut (Mac)

1. Open Script Editor
2. Paste this:
```applescript
do shell script "cd /path/to/your/script && python3 calendar_sync_gui.py"
```
3. Replace `/path/to/your/script` with the actual path
4. Save as Application on your Desktop
5. Name it "Calendar Sync"

Now you can just double-click the icon!
