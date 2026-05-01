#!/usr/bin/env python3
"""
Outlook PDF to Google Calendar Sync - GUI Application
Simple interface to sync Outlook calendar PDF to Google Calendar
"""

import os
import re
import pickle
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox
from datetime import datetime, timedelta
from threading import Thread
from icalendar import Calendar, Event as ICSEvent
import pytz
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import pdfplumber

# ===========================
# CONFIGURATION
# ===========================

GOOGLE_SCOPES = ['https://www.googleapis.com/auth/calendar']
LOCAL_TIMEZONE = pytz.timezone('America/Toronto')
CONFIG_FILE = 'sync_config.txt'

# ===========================
# PDF PARSING FUNCTIONS
# ===========================

def parse_outlook_pdf(pdf_path):
    """Parse Outlook PDF agenda and extract events"""
    events = []
    
    with pdfplumber.open(pdf_path) as pdf:
        full_text = ""
        for page in pdf.pages:
            full_text += page.extract_text() + "\n"
    
    lines = full_text.split('\n')
    current_date = None
    current_title = None
    
    for i, line in enumerate(lines):
        line = line.strip()
        if not line:
            continue
        
        # Skip header lines
        if 'Toronto Post' in line or re.match(r'[A-Z][a-z]+ \d{4}', line):
            continue
        
        # Try to match date headers (e.g., "Monday, February 2, 2026")
        date_match = re.match(r'([A-Za-z]+,\s+[A-Za-z]+\s+\d{1,2},\s+\d{4})', line)
        if date_match:
            current_date = date_match.group(1)
            current_title = None
            continue
        
        # Try to match time line (e.g., "Mon 2/2/2026 9:00 AM - 6:00 PM")
        time_match = re.match(r'[A-Za-z]{3}\s+\d{1,2}/\d{1,2}/\d{4}\s+(\d{1,2}:\d{2}\s*(?:AM|PM))\s*-\s*(\d{1,2}:\d{2}\s*(?:AM|PM))', line)
        if time_match and current_date and current_title:
            start_time = time_match.group(1)
            end_time = time_match.group(2)
            
            event = {
                'date': current_date,
                'start_time': start_time,
                'end_time': end_time,
                'title': current_title,
                'location': '',
                'description': ''
            }
            events.append(event)
            current_title = None
            continue
        
        # All-day events
        allday_match = re.match(r'[A-Za-z]{3}\s+\d{1,2}/\d{1,2}/\d{4}\s+\(All day\)', line)
        if allday_match and current_date and current_title:
            event = {
                'date': current_date,
                'start_time': '12:00 AM',
                'end_time': '11:59 PM',
                'title': current_title,
                'location': '',
                'description': 'All day event'
            }
            events.append(event)
            current_title = None
            continue
        
        # Date range events
        daterange_match = re.match(r'[A-Za-z]{3}\s+\d{1,2}/\d{1,2}/\d{4}\s+to\s+[A-Za-z]{3}\s+\d{1,2}/\d{1,2}/\d{4}', line)
        if daterange_match and current_date and current_title:
            event = {
                'date': current_date,
                'start_time': '12:00 AM',
                'end_time': '11:59 PM',
                'title': current_title + ' (Multi-day)',
                'location': '',
                'description': line
            }
            events.append(event)
            current_title = None
            continue
        
        # If we have a current_date but no current_title, check if this could be a title
        if current_date and not current_title:
            # Skip lines that are obviously not titles
            if (not line.startswith('Mon ') and not line.startswith('Tue ') and
                not line.startswith('Wed ') and not line.startswith('Thu ') and
                not line.startswith('Fri ') and not line.startswith('Sat ') and
                not line.startswith('Sun ') and
                not line.startswith('Location:') and not line.startswith('Where:') and
                len(line) > 3):
                
                # Look ahead to check if the IMMEDIATE next non-empty line is a time line
                # This prevents description lines from becoming titles
                has_time_line_ahead = False
                for j in range(i + 1, min(i + 3, len(lines))):  # Check next 2 lines
                    next_line = lines[j].strip()
                    if not next_line:  # Skip empty lines
                        continue
                    # If we hit a non-empty line, check if it's a time line
                    if re.match(r'[A-Za-z]{3}\s+\d{1,2}/\d{1,2}/\d{4}\s+\d{1,2}:\d{2}\s*(?:AM|PM)', next_line):
                        has_time_line_ahead = True
                    break  # Stop after first non-empty line
                
                if has_time_line_ahead:
                    current_title = line
    
    return events


def parse_datetime(date_str, time_str, timezone):
    """Convert date and time strings to datetime object"""
    date_parts = date_str.split(', ', 1)
    if len(date_parts) > 1:
        date_str = date_parts[1]
    
    dt = datetime.strptime(date_str, "%B %d, %Y")
    time_obj = datetime.strptime(time_str, "%I:%M %p")
    combined = dt.replace(hour=time_obj.hour, minute=time_obj.minute)
    
    return timezone.localize(combined)


# ===========================
# GOOGLE CALENDAR FUNCTIONS
# ===========================

def get_google_credentials():
    """Get Google Calendar API credentials"""
    # Get the directory where this script is located
    script_dir = os.path.dirname(os.path.abspath(__file__))
    token_path = os.path.join(script_dir, 'token.pickle')
    creds_path = os.path.join(script_dir, 'google_credentials.json')
    
    creds = None
    
    if os.path.exists(token_path):
        with open(token_path, 'rb') as token:
            creds = pickle.load(token)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                creds_path, GOOGLE_SCOPES)
            creds = flow.run_local_server(port=0)
        
        with open(token_path, 'wb') as token:
            pickle.dump(creds, token)
    
    return creds


def find_calendar_id(calendar_name):
    """Find Google Calendar ID by name"""
    creds = get_google_credentials()
    service = build('calendar', 'v3', credentials=creds)
    
    # Get all calendars
    calendars_result = service.calendarList().list().execute()
    calendars = calendars_result.get('items', [])
    
    for calendar in calendars:
        if calendar['summary'] == calendar_name:
            return calendar['id']
    
    raise ValueError(f"Calendar '{calendar_name}' not found. Available calendars: " + 
                     ", ".join([c['summary'] for c in calendars]))


def clear_calendar_events(calendar_id, start_date, end_date):
    """Clear all events in the calendar within the date range"""
    creds = get_google_credentials()
    service = build('calendar', 'v3', credentials=creds)
    
    # Get events in the date range
    events_result = service.events().list(
        calendarId=calendar_id,
        timeMin=start_date.isoformat(),
        timeMax=end_date.isoformat(),
        singleEvents=True,
        orderBy='startTime'
    ).execute()
    
    events = events_result.get('items', [])
    deleted_count = 0
    
    for event in events:
        try:
            service.events().delete(calendarId=calendar_id, eventId=event['id']).execute()
            deleted_count += 1
        except Exception as e:
            print(f"Error deleting event: {e}")
    
    return deleted_count


def import_events_to_google(events, calendar_id):
    """Import parsed events directly to Google Calendar"""
    creds = get_google_credentials()
    service = build('calendar', 'v3', credentials=creds)
    
    imported_count = 0
    errors = []
    
    for event in events:
        try:
            # Parse datetimes
            start_dt = parse_datetime(event['date'], event['start_time'], LOCAL_TIMEZONE)
            end_dt = parse_datetime(event['date'], event['end_time'], LOCAL_TIMEZONE)
            
            # Create Google Calendar event
            gcal_event = {
                'summary': event['title'],
                'start': {
                    'dateTime': start_dt.isoformat(),
                    'timeZone': 'America/Toronto',
                },
                'end': {
                    'dateTime': end_dt.isoformat(),
                    'timeZone': 'America/Toronto',
                }
            }
            
            if event.get('location'):
                gcal_event['location'] = event['location']
            
            if event.get('description'):
                gcal_event['description'] = event['description']
            
            # Insert event
            service.events().insert(calendarId=calendar_id, body=gcal_event).execute()
            imported_count += 1
            
        except Exception as e:
            errors.append(f"Error importing '{event['title']}': {str(e)}")
    
    return imported_count, errors


# ===========================
# GUI APPLICATION
# ===========================

class CalendarSyncApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Outlook to Google Calendar Sync")
        self.root.geometry("700x600")
        
        # Load saved config
        self.pdf_folder = tk.StringVar(value=self.load_config('pdf_folder', os.path.expanduser('~/Desktop')))
        self.calendar_name = tk.StringVar(value=self.load_config('calendar_name', 'DIFUZE - STUDIO CALENDAR'))
        
        self.create_widgets()
    
    def load_config(self, key, default=''):
        """Load saved configuration"""
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'r') as f:
                for line in f:
                    if line.startswith(f"{key}="):
                        return line.split('=', 1)[1].strip()
        return default
    
    def save_config(self):
        """Save configuration"""
        with open(CONFIG_FILE, 'w') as f:
            f.write(f"pdf_folder={self.pdf_folder.get()}\n")
            f.write(f"calendar_name={self.calendar_name.get()}\n")
    
    def create_widgets(self):
        # Main container
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="📅 Calendar Sync", font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=3, pady=10)
        
        # PDF Folder
        ttk.Label(main_frame, text="PDF Folder:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.pdf_folder, width=50).grid(row=1, column=1, sticky=(tk.W, tk.E), pady=5)
        ttk.Button(main_frame, text="Browse", command=self.browse_folder).grid(row=1, column=2, padx=5, pady=5)
        
        # Calendar Name
        ttk.Label(main_frame, text="Google Calendar:").grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.calendar_name, width=50).grid(row=2, column=1, sticky=(tk.W, tk.E), pady=5)
        
        # Info label
        info_text = "The script will process ALL PDF files in the folder and combine them"
        ttk.Label(main_frame, text=info_text, font=('Arial', 9), foreground='gray').grid(
            row=3, column=0, columnspan=3, pady=5)
        
        # Sync Button
        self.sync_button = ttk.Button(main_frame, text="🔄 SYNC NOW", command=self.start_sync, style='Accent.TButton')
        self.sync_button.grid(row=4, column=0, columnspan=3, pady=20, ipadx=40, ipady=10)
        
        # Configure button style
        style = ttk.Style()
        style.configure('Accent.TButton', font=('Arial', 12, 'bold'))
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        # Status label
        self.status_label = ttk.Label(main_frame, text="Ready to sync", foreground='green')
        self.status_label.grid(row=6, column=0, columnspan=3, pady=5)
        
        # Log area
        ttk.Label(main_frame, text="Log:").grid(row=7, column=0, sticky=tk.W, pady=(10, 0))
        
        self.log_text = scrolledtext.ScrolledText(main_frame, height=15, width=80, state='disabled')
        self.log_text.grid(row=8, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        
        # Configure grid weights for log area
        main_frame.rowconfigure(8, weight=1)
    
    def browse_folder(self):
        """Browse for PDF folder"""
        folder = filedialog.askdirectory(initialdir=self.pdf_folder.get())
        if folder:
            self.pdf_folder.set(folder)
            self.save_config()
    
    def log(self, message, level='info'):
        """Add message to log"""
        self.log_text.config(state='normal')
        
        timestamp = datetime.now().strftime("%H:%M:%S")
        
        if level == 'error':
            prefix = "❌"
            color = 'red'
        elif level == 'success':
            prefix = "✅"
            color = 'green'
        elif level == 'warning':
            prefix = "⚠️"
            color = 'orange'
        else:
            prefix = "ℹ️"
            color = 'black'
        
        self.log_text.insert(tk.END, f"[{timestamp}] {prefix} {message}\n")
        self.log_text.see(tk.END)
        self.log_text.config(state='disabled')
        self.root.update()
    
    def update_status(self, message, color='black'):
        """Update status label"""
        self.status_label.config(text=message, foreground=color)
        self.root.update()
    
    def start_sync(self):
        """Start sync in a separate thread"""
        # Save config
        self.save_config()
        
        # Start sync in background thread
        thread = Thread(target=self.sync_calendar)
        thread.daemon = True
        thread.start()
    
    def sync_calendar(self):
        """Main sync function"""
        try:
            # Disable sync button
            self.sync_button.config(state='disabled')
            self.progress.start()
            
            # Clear log
            self.log_text.config(state='normal')
            self.log_text.delete(1.0, tk.END)
            self.log_text.config(state='disabled')
            
            self.log("=" * 60)
            self.log("Starting calendar sync...")
            self.log("=" * 60)
            
            # Find PDF files in folder
            folder_path = self.pdf_folder.get()
            pdf_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.pdf')]
            
            if not pdf_files:
                self.log(f"No PDF files found in: {folder_path}", 'error')
                self.update_status("❌ No PDF files found", 'red')
                messagebox.showerror("Error", f"No PDF files found in:\n{folder_path}\n\nPlease save your Outlook agenda PDF in this folder.")
                return
            
            # Sort PDFs alphabetically for consistent processing
            pdf_files.sort()
            
            self.log(f"Found {len(pdf_files)} PDF file(s): {', '.join(pdf_files)}")
            self.update_status(f"📄 Parsing {len(pdf_files)} PDF(s)...", 'blue')
            
            # Parse all PDFs and combine events
            all_events = []
            for pdf_file in pdf_files:
                pdf_path = os.path.join(folder_path, pdf_file)
                self.log(f"Parsing {pdf_file}...")
                
                try:
                    events = parse_outlook_pdf(pdf_path)
                    all_events.extend(events)
                    self.log(f"  Found {len(events)} events in {pdf_file}", 'success')
                except Exception as e:
                    self.log(f"  Error parsing {pdf_file}: {str(e)}", 'error')
            
            events = all_events
            self.log(f"\nTotal events from all PDFs: {len(events)}", 'success')
            
            if not events:
                self.log("No events found in PDF", 'warning')
                self.update_status("⚠️ No events found", 'orange')
                messagebox.showwarning("Warning", "No events found in the PDF.\n\nMake sure you exported a 'Detailed Agenda' from Outlook.")
                return
            
            # Show sample events
            self.log("\nSample of parsed events:")
            for event in events[:3]:
                self.log(f"  • {event['date']} {event['start_time']}: {event['title']}")
            if len(events) > 3:
                self.log(f"  ... and {len(events) - 3} more")
            
            # Get date range - find actual min/max across all events
            all_datetimes = []
            for event in events:
                try:
                    start_dt = parse_datetime(event['date'], event['start_time'], LOCAL_TIMEZONE)
                    end_dt = parse_datetime(event['date'], event['end_time'], LOCAL_TIMEZONE)
                    all_datetimes.append(start_dt)
                    all_datetimes.append(end_dt)
                except:
                    pass
            
            if not all_datetimes:
                self.log("No valid dates found in events", 'error')
                self.update_status("❌ No valid dates", 'red')
                return
            
            first_event = min(all_datetimes)
            last_event = max(all_datetimes)
            
            self.log(f"\nDate range: {first_event.date()} to {last_event.date()}")
            
            # Find calendar
            self.update_status("🔍 Finding Google Calendar...", 'blue')
            self.log("\nFinding Google Calendar...")
            
            try:
                calendar_id = find_calendar_id(self.calendar_name.get())
                self.log(f"Found calendar: {self.calendar_name.get()}", 'success')
            except ValueError as e:
                self.log(str(e), 'error')
                self.update_status("❌ Calendar not found", 'red')
                messagebox.showerror("Error", str(e))
                return
            
            # Clear existing events
            self.update_status("🗑️ Clearing old events...", 'blue')
            self.log("\nClearing old events in date range...")
            
            deleted_count = clear_calendar_events(calendar_id, first_event, last_event + timedelta(days=1))
            self.log(f"Deleted {deleted_count} old events", 'success')
            
            # Import new events
            self.update_status("📤 Importing new events...", 'blue')
            self.log("\nImporting new events...")
            
            imported_count, errors = import_events_to_google(events, calendar_id)
            
            # Show results
            self.log("\n" + "=" * 60)
            self.log("SYNC COMPLETE", 'success')
            self.log("=" * 60)
            self.log(f"Events parsed: {len(events)}")
            self.log(f"Old events deleted: {deleted_count}")
            self.log(f"New events imported: {imported_count}", 'success')
            
            if errors:
                self.log(f"\nErrors: {len(errors)}", 'warning')
                for error in errors[:5]:
                    self.log(f"  {error}", 'error')
                if len(errors) > 5:
                    self.log(f"  ... and {len(errors) - 5} more errors", 'error')
            
            self.update_status(f"✅ Sync complete! Imported {imported_count} events", 'green')
            messagebox.showinfo("Success", f"Sync complete!\n\nImported {imported_count} events to '{self.calendar_name.get()}'")
            
        except Exception as e:
            self.log(f"\nFATAL ERROR: {str(e)}", 'error')
            self.update_status("❌ Sync failed", 'red')
            messagebox.showerror("Error", f"Sync failed:\n\n{str(e)}")
            
        finally:
            # Re-enable sync button
            self.sync_button.config(state='normal')
            self.progress.stop()


# ===========================
# MAIN
# ===========================

def main():
    root = tk.Tk()
    app = CalendarSyncApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
