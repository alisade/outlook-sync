#!/usr/bin/env python3
"""
Outlook Calendar HTML/MHTML to ICS Converter
Extracts calendar events from an Outlook HTML or MHTML export and creates an ICS file
compatible with macOS Calendar, or exports directly to Google Calendar.
"""

import argparse
import email
import os
import pickle
import re
import sys
from datetime import datetime, timedelta
from email import policy
from email.parser import BytesParser
from html.parser import HTMLParser
from pathlib import Path

# Load environment variables from .env file (optional)
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass  # dotenv not installed, continue without it

# Google Calendar API imports (optional)
try:
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build
    from googleapiclient.errors import HttpError
    GOOGLE_CALENDAR_AVAILABLE = True
except ImportError:
    GOOGLE_CALENDAR_AVAILABLE = False


def extract_html_from_mhtml(file_path):
    """Extract HTML content from MHTML file.
    
    MHTML (MIME HTML) is a web archive format that bundles HTML and resources.
    This function extracts the HTML part from the MHTML container.
    
    Args:
        file_path: Path to the MHTML file
        
    Returns:
        HTML content as string, or None if not MHTML format
    """
    try:
        with open(file_path, 'rb') as f:
            # Read first few bytes to check if it's MHTML
            first_line = f.readline().decode('utf-8', errors='ignore')
            f.seek(0)
            
            # Check for MHTML markers (MIME headers)
            if not any(marker in first_line for marker in ['From:', 'MIME-Version:', 'Content-Type: multipart']):
                # Not an MHTML file, return None
                return None
            
            # Parse as MIME message
            msg = BytesParser(policy=policy.default).parse(f)
            
            # Walk through all parts to find HTML content
            if msg.is_multipart():
                for part in msg.walk():
                    content_type = part.get_content_type()
                    if content_type == 'text/html':
                        # Found HTML part
                        return part.get_content()
            else:
                # Single part message
                if msg.get_content_type() == 'text/html':
                    return msg.get_content()
            
            return None
            
    except Exception as e:
        print(f"Warning: Failed to parse as MHTML: {e}")
        return None


class OutlookEventParser(HTMLParser):
    """Parser to extract calendar events from Outlook HTML."""
    
    def __init__(self):
        super().__init__()
        self.events = []
        self.processed_events = set()  # To avoid duplicates
        
    def handle_starttag(self, tag, attrs):
        """Extract event data from div tags with aria-label attributes."""
        if tag == "div":
            attrs_dict = dict(attrs)
            aria_label = attrs_dict.get("aria-label", "")
            role = attrs_dict.get("role", "")
            
            # Look for calendar event patterns in aria-label
            # Support both old format and new Outlook Web format with role="button"
            if ((" to " in aria_label and ("AM" in aria_label or "PM" in aria_label)) and 
                ("Monday" in aria_label or "Tuesday" in aria_label or 
                "Wednesday" in aria_label or "Thursday" in aria_label or 
                "Friday" in aria_label or "Saturday" in aria_label or "Sunday" in aria_label)):
                
                event_data = self.parse_event_label(aria_label)
                if event_data:
                    # Create a unique key to avoid duplicates (same event repeated in HTML)
                    event_key = f"{event_data['summary']}_{event_data['start'].isoformat()}"
                    if event_key not in self.processed_events:
                        self.processed_events.add(event_key)
                        self.events.append(event_data)
    
    def parse_event_label(self, label):
        """Parse event information from aria-label text."""
        try:
            # Pattern: "Event Name, H:MM AM to H:MM AM, Day, Month DD, YYYY, By Organizer, Status, Type"
            # Example: "CSVCS- Daily Stand Up, 7:30 AM to 8:00 AM, Monday, September 29, 2025, By Veronica Sanchez, Tentative, Recurring event"
            
            # Split by commas but be careful with commas in event names
            parts = label.split(", ")
            
            if len(parts) < 5:
                return None
            
            # Find the time pattern (H:MM AM/PM to H:MM AM/PM)
            time_pattern = re.compile(r'(\d{1,2}):(\d{2})\s*(AM|PM)\s*to\s*(\d{1,2}):(\d{2})\s*(AM|PM)', re.IGNORECASE)
            time_idx = None
            time_match = None
            
            for i, part in enumerate(parts):
                match = time_pattern.search(part)
                if match:
                    time_idx = i
                    time_match = match
                    break
            
            if time_idx is None or not time_match:
                return None
            
            # Event name is everything before the time
            event_name = ", ".join(parts[:time_idx])
            
            # Extract times with AM/PM
            start_hour, start_min, start_ampm, end_hour, end_min, end_ampm = time_match.groups()
            
            # Convert to 24-hour format
            start_hour_24 = int(start_hour)
            if start_ampm.upper() == 'PM' and start_hour_24 != 12:
                start_hour_24 += 12
            elif start_ampm.upper() == 'AM' and start_hour_24 == 12:
                start_hour_24 = 0
                
            end_hour_24 = int(end_hour)
            if end_ampm.upper() == 'PM' and end_hour_24 != 12:
                end_hour_24 += 12
            elif end_ampm.upper() == 'AM' and end_hour_24 == 12:
                end_hour_24 = 0
            
            # Find date (Day, Month DD, YYYY)
            # The day comes right after the time
            if time_idx + 3 >= len(parts):
                return None
                
            day_of_week = parts[time_idx + 1].strip()
            month_day = parts[time_idx + 2].strip()  # e.g., "September 29"
            year = parts[time_idx + 3].strip()
            
            # Parse the date
            date_str = f"{month_day}, {year}"
            date_obj = datetime.strptime(date_str, "%B %d, %Y")
            
            # Combine date with times
            start_time = date_obj.replace(hour=start_hour_24, minute=int(start_min))
            end_time = date_obj.replace(hour=end_hour_24, minute=int(end_min))
            
            # Extract location (e.g., "Microsoft Teams Meeting") - comes between date and organizer
            location = ""
            organizer = ""
            for i, part in enumerate(parts[time_idx + 4:], start=time_idx + 4):
                # Check for location indicators (before "By ")
                if part.startswith("By "):
                    organizer = part[3:].strip()
                    break
                # Common location patterns
                elif "Microsoft Teams Meeting" in part or "Teams" in part:
                    location = part.strip()
                elif "Meeting" in part and not part.startswith("By "):
                    # Could be a location
                    if not location:
                        location = part.strip()
            
            # Extract status and type
            status = ""
            event_type = ""
            for i in range(time_idx + 4, len(parts)):
                if parts[i] in ["Tentative", "Busy", "Free", "Out of Office"]:
                    status = parts[i]
                if "Recurring" in parts[i] or "Exception" in parts[i] or "Canceled" in parts[i]:
                    event_type = parts[i]
            
            # Check if event is canceled
            is_canceled = "Canceled:" in event_name or "Canceled" in event_type
            
            return {
                "summary": event_name,
                "start": start_time,
                "end": end_time,
                "organizer": organizer,
                "status": status,
                "event_type": event_type,
                "is_canceled": is_canceled,
                "location": location  # Add location field
            }
            
        except Exception as e:
            print(f"Warning: Could not parse event: {label[:100]}... Error: {e}", file=sys.stderr)
            return None


class ICSGenerator:
    """Generate ICS file from event data."""
    
    def __init__(self, email_domain="domain.com", timezone="America/Los_Angeles"):
        self.ics_lines = []
        self.email_domain = email_domain
        self.timezone = timezone
        
    def add_header(self):
        """Add ICS file header."""
        self.ics_lines.extend([
            "BEGIN:VCALENDAR",
            "VERSION:2.0",
            "PRODID:-//Outlook Calendar to ICS Converter//EN",
            "CALSCALE:GREGORIAN",
            "METHOD:PUBLISH",
            "X-WR-CALNAME:Outlook Calendar Export",
            f"X-WR-TIMEZONE:{self.timezone}",
        ])
    
    def add_event(self, event):
        """Add an event to the ICS file."""
        # Format datetimes for ICS (YYYYMMDDTHHmmss)
        start_dt = event["start"].strftime("%Y%m%dT%H%M%S")
        end_dt = event["end"].strftime("%Y%m%dT%H%M%S")
        now_dt = datetime.now().strftime("%Y%m%dT%H%M%SZ")
        
        # Generate a unique ID
        uid = f"{start_dt}-{hash(event['summary'])}"
        
        # Map status to ICS status
        status_map = {
            "Tentative": "TENTATIVE",
            "Busy": "CONFIRMED",
            "Free": "CONFIRMED",
            "Out of Office": "CONFIRMED"
        }
        ics_status = status_map.get(event["status"], "CONFIRMED")
        
        if event["is_canceled"]:
            ics_status = "CANCELLED"
        
        self.ics_lines.extend([
            "BEGIN:VEVENT",
            f"UID:{uid}",
            f"DTSTAMP:{now_dt}",
            f"DTSTART:{start_dt}",
            f"DTEND:{end_dt}",
            f"SUMMARY:{self.escape_text(event['summary'])}",
            f"STATUS:{ics_status}",
        ])
        
        if event["organizer"]:
            # Create email from organizer name (simplified - just use first word as username)
            organizer_name = event["organizer"].split()[0].lower() if event["organizer"] else "organizer"
            self.ics_lines.append(f"ORGANIZER;CN={self.escape_text(event['organizer'])}:mailto:{organizer_name}@{self.email_domain}")
        
        # Add location if available
        if event.get("location"):
            self.ics_lines.append(f"LOCATION:{self.escape_text(event['location'])}")
        
        if event["event_type"]:
            self.ics_lines.append(f"DESCRIPTION:{self.escape_text(event['event_type'])}")
        
        # Add transparency based on status
        transp = "TRANSPARENT" if event["status"] == "Free" else "OPAQUE"
        self.ics_lines.append(f"TRANSP:{transp}")
        
        self.ics_lines.append("END:VEVENT")
    
    def add_footer(self):
        """Add ICS file footer."""
        self.ics_lines.append("END:VCALENDAR")
    
    def escape_text(self, text):
        """Escape special characters in ICS text fields."""
        # Escape special characters according to RFC 5545
        text = text.replace("\\", "\\\\")
        text = text.replace(";", "\\;")
        text = text.replace(",", "\\,")
        text = text.replace("\n", "\\n")
        return text
    
    def generate(self, events):
        """Generate complete ICS file content."""
        self.add_header()
        
        for event in events:
            self.add_event(event)
        
        self.add_footer()
        
        return "\n".join(self.ics_lines)


class GoogleCalendarExporter:
    """Export events directly to Google Calendar."""
    
    # If modifying these scopes, delete the file token.pickle.
    SCOPES = ['https://www.googleapis.com/auth/calendar']
    
    def __init__(self, credentials_file='credentials.json', token_file='token.pickle', use_service_account=False):
        """Initialize Google Calendar exporter.
        
        Args:
            credentials_file: Path to credentials file (OAuth or service account JSON)
            token_file: Path to token file (only used for OAuth)
            use_service_account: If True, use service account authentication (static key)
        """
        self.credentials_file = credentials_file
        self.token_file = token_file
        self.use_service_account = use_service_account
        self.service = None
        
    def authenticate(self):
        """Authenticate with Google Calendar API."""
        if not GOOGLE_CALENDAR_AVAILABLE:
            raise ImportError(
                "Google Calendar API libraries not installed. "
                "Install them with: pip install -r requirements.txt"
            )
        
        if not os.path.exists(self.credentials_file):
            raise FileNotFoundError(
                f"Credentials file '{self.credentials_file}' not found. "
                "Please download it from Google Cloud Console. "
                "See README for instructions."
            )
        
        # Check if it's a service account by inspecting the file
        if self.use_service_account or self._is_service_account_file():
            # Use service account authentication (static key)
            creds = self._authenticate_service_account()
        else:
            # Use OAuth2 authentication (interactive)
            creds = self._authenticate_oauth()
        
        self.service = build('calendar', 'v3', credentials=creds)
        return True
    
    def _is_service_account_file(self):
        """Check if credentials file is a service account key."""
        try:
            import json
            with open(self.credentials_file, 'r') as f:
                data = json.load(f)
                return data.get('type') == 'service_account'
        except:
            return False
    
    def _authenticate_service_account(self):
        """Authenticate using service account (static key file)."""
        from google.oauth2 import service_account
        
        print("  Using service account authentication (static key)...")
        creds = service_account.Credentials.from_service_account_file(
            self.credentials_file,
            scopes=self.SCOPES
        )
        return creds
    
    def _authenticate_oauth(self):
        """Authenticate using OAuth2 (interactive browser flow)."""
        creds = None
        
        # The file token.pickle stores the user's access and refresh tokens
        if os.path.exists(self.token_file):
            with open(self.token_file, 'rb') as token:
                creds = pickle.load(token)
        
        # If there are no (valid) credentials available, let the user log in
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                print("  Using OAuth2 authentication (browser-based)...")
                flow = InstalledAppFlow.from_client_secrets_file(
                    self.credentials_file, self.SCOPES)
                creds = flow.run_local_server(port=0)
            
            # Save the credentials for the next run
            with open(self.token_file, 'wb') as token:
                pickle.dump(creds, token)
        
        return creds
    
    def create_calendar(self, calendar_name='Outlook Calendar Import', timezone='America/Los_Angeles'):
        """Create a new calendar or get existing one.
        
        Args:
            calendar_name: Name for the calendar
            timezone: IANA timezone string
        """
        try:
            # Check if calendar already exists
            calendar_list = self.service.calendarList().list().execute()
            for calendar in calendar_list.get('items', []):
                if calendar.get('summary') == calendar_name:
                    print(f"Using existing calendar: {calendar_name}")
                    return calendar['id']
            
            # Create new calendar
            calendar = {
                'summary': calendar_name,
                'timeZone': timezone
            }
            created_calendar = self.service.calendars().insert(body=calendar).execute()
            print(f"Created new calendar: {calendar_name} (timezone: {timezone})")
            return created_calendar['id']
            
        except HttpError as error:
            print(f"An error occurred: {error}")
            return None
    
    def _find_existing_event(self, event_data, calendar_id, timezone):
        """Find existing event in calendar that matches this event.
        
        Args:
            event_data: Event data dictionary
            calendar_id: Google Calendar ID
            timezone: IANA timezone string
            
        Returns:
            Existing event dict if found, None otherwise
        """
        try:
            # Search for events around the target time
            # Use a broader window to account for timezone differences
            from datetime import timedelta
            
            event_start = event_data['start']
            # Search from 1 day before to 1 day after the event
            # This broader window handles timezone conversion edge cases
            search_start = event_start - timedelta(days=1)
            search_end = event_start + timedelta(days=2)
            
            # Format times as UTC (which is what the API expects)
            # Since our times are naive (no timezone), we add 'Z' to treat them as UTC
            # Google will return events in that UTC range regardless of their stored timezone
            time_min = search_start.isoformat() + 'Z'
            time_max = search_end.isoformat() + 'Z'
            
            # Get all events in this time range
            # showDeleted=True includes cancelled events in the results
            events_result = self.service.events().list(
                calendarId=calendar_id,
                timeMin=time_min,
                timeMax=time_max,
                singleEvents=True,
                orderBy='startTime',
                showDeleted=True  # Include cancelled events
            ).execute()
            
            events = events_result.get('items', [])
            
            # Find exact match by title and start time (within 1 minute tolerance)
            # Skip cancelled events UNLESS we're looking for a cancelled event
            looking_for_cancelled = event_data.get('is_canceled', False)
            
            for event in events:
                # Skip cancelled events if we're not looking for a cancelled event
                # Cancelled events are deleted instances of recurring events that can't be updated
                if event.get('status') == 'cancelled' and not looking_for_cancelled:
                    continue
                
                # Check if title matches
                if event.get('summary') != event_data['summary']:
                    continue
                
                # Get existing event's start time
                existing_start_str = event.get('start', {}).get('dateTime', '')
                if not existing_start_str:
                    continue
                
                # Parse the existing start time (handle timezone)
                # Format: 2025-09-29T10:30:00-07:00 or 2025-09-29T10:30:00Z
                try:
                    # Extract just the date and time part (YYYY-MM-DDTHH:MM)
                    existing_start_prefix = existing_start_str[:16]  # "2025-09-29T10:30"
                    new_start_prefix = event_data['start'].strftime('%Y-%m-%dT%H:%M')
                    
                    if existing_start_prefix == new_start_prefix:
                        return event
                except:
                    continue
            
            return None
            
        except HttpError as error:
            # If search fails, return None to create new event
            return None
    
    def _events_are_different(self, existing_event, new_event_data, event_data, timezone):
        """Compare existing event with new event data to detect changes.
        
        Args:
            existing_event: Existing Google Calendar event
            new_event_data: New event data in Google Calendar format
            event_data: Original event data dictionary
            timezone: IANA timezone string
            
        Returns:
            True if events are different, False if identical
        """
        # Compare summary (required field)
        if existing_event.get('summary', '') != new_event_data.get('summary', ''):
            return True
        
        # Compare start/end times (only date and time, ignore timezone offset)
        # Format: "2025-09-29T10:30" (first 16 chars)
        existing_start = existing_event.get('start', {}).get('dateTime', '')[:16]
        new_start = event_data['start'].strftime('%Y-%m-%dT%H:%M')
        if existing_start != new_start:
            return True
        
        existing_end = existing_event.get('end', {}).get('dateTime', '')[:16]
        new_end = event_data['end'].strftime('%Y-%m-%dT%H:%M')
        if existing_end != new_end:
            return True
        
        # Compare description (normalize empty strings and None)
        existing_desc = (existing_event.get('description') or '').strip()
        new_desc = (new_event_data.get('description') or '').strip()
        if existing_desc != new_desc:
            return True
        
        # Compare status (normalize)
        existing_status = (existing_event.get('status') or 'confirmed').lower()
        new_status = (new_event_data.get('status') or 'confirmed').lower()
        if existing_status != new_status:
            return True
        
        # Compare transparency (normalize, default to opaque)
        existing_transp = (existing_event.get('transparency') or 'opaque').lower()
        new_transp = (new_event_data.get('transparency') or 'opaque').lower()
        if existing_transp != new_transp:
            return True
        
        # No differences found
        return False
    
    def export_event(self, event_data, calendar_id='primary', timezone='America/Los_Angeles', verbose=False):
        """Export a single event to Google Calendar (create or update if changed).
        
        Args:
            event_data: Event data dictionary
            calendar_id: Google Calendar ID
            timezone: IANA timezone string (e.g., 'America/New_York', 'America/Los_Angeles')
            verbose: If True, print debug information
            
        Returns:
            Tuple of (action, link) where action is 'created', 'updated', or 'skipped'
        """
        try:
            # Convert event data to Google Calendar format
            google_event = {
                'summary': event_data['summary'],
                'start': {
                    'dateTime': event_data['start'].isoformat(),
                    'timeZone': timezone,
                },
                'end': {
                    'dateTime': event_data['end'].isoformat(),
                    'timeZone': timezone,
                },
            }
            
            # Add organizer and Teams meeting link
            teams_link = os.getenv('TEAMS_MEETING_LINK', '')
            
            if event_data.get('organizer'):
                google_event['description'] = f"Organizer: {event_data['organizer']}"
                if event_data.get('event_type'):
                    google_event['description'] += f"\nType: {event_data['event_type']}"
            else:
                google_event['description'] = ""
            
            # Add location from parsed event or Teams meeting link
            if event_data.get('location'):
                google_event['location'] = event_data['location']
            elif teams_link:
                google_event['location'] = "Microsoft Teams Meeting"
            
            # Add Teams meeting link to description (if configured)
            if teams_link:
                if google_event['description']:
                    google_event['description'] += f"\n\nMicrosoft Teams Meeting:\n{teams_link}"
                else:
                    google_event['description'] = f"Microsoft Teams Meeting:\n{teams_link}"
            
            # Add status
            status_map = {
                'Tentative': 'tentative',
                'Busy': 'confirmed',
                'Free': 'confirmed',
                'Out of Office': 'confirmed'
            }
            google_event['status'] = status_map.get(event_data['status'], 'confirmed')
            
            if event_data['is_canceled']:
                google_event['status'] = 'cancelled'
            
            # Add transparency
            google_event['transparency'] = 'transparent' if event_data['status'] == 'Free' else 'opaque'
            
            # Check if event already exists
            if verbose:
                print(f"\n  Checking: {event_data['summary'][:50]}... @ {event_data['start'].strftime('%Y-%m-%d %H:%M')}")
            
            existing_event = self._find_existing_event(event_data, calendar_id, timezone)
            
            if existing_event:
                if verbose:
                    print(f"    Found existing event (ID: {existing_event['id'][:20]}...)")
                
                # Event exists, check if it's different
                is_different = self._events_are_different(existing_event, google_event, event_data, timezone)
                
                if is_different:
                    if verbose:
                        print(f"    → Event changed, updating...")
                        # Show what changed
                        if existing_event.get('summary', '') != google_event.get('summary', ''):
                            print(f"       Changed: title")
                        if existing_event.get('start', {}).get('dateTime', '')[:16] != event_data['start'].strftime('%Y-%m-%dT%H:%M'):
                            print(f"       Changed: start time")
                        if existing_event.get('end', {}).get('dateTime', '')[:16] != event_data['end'].strftime('%Y-%m-%dT%H:%M'):
                            print(f"       Changed: end time")
                        if (existing_event.get('description') or '').strip() != (google_event.get('description') or '').strip():
                            print(f"       Changed: description")
                            print(f"         Old: {(existing_event.get('description') or '')[:50]}")
                            print(f"         New: {(google_event.get('description') or '')[:50]}")
                        if (existing_event.get('status') or 'confirmed').lower() != (google_event.get('status') or 'confirmed').lower():
                            print(f"       Changed: status")
                            print(f"         Old: {existing_event.get('status') or 'confirmed'}")
                            print(f"         New: {google_event.get('status') or 'confirmed'}")
                        if (existing_event.get('transparency') or 'opaque').lower() != (google_event.get('transparency') or 'opaque').lower():
                            print(f"       Changed: transparency")
                    
                    # Update the existing event
                    updated_event = self.service.events().update(
                        calendarId=calendar_id,
                        eventId=existing_event['id'],
                        body=google_event
                    ).execute()
                    return ('updated', updated_event.get('htmlLink'))
                else:
                    if verbose:
                        print(f"    → Event unchanged, skipping")
                    # Event is identical, skip
                    return ('skipped', existing_event.get('htmlLink'))
            else:
                if verbose:
                    print(f"    → No existing event found, creating new...")
                # Event doesn't exist, create it
                event = self.service.events().insert(calendarId=calendar_id, body=google_event).execute()
                return ('created', event.get('htmlLink'))
            
        except HttpError as error:
            print(f"An error occurred: {error}", file=sys.stderr)
            return ('error', None)
    
    def export_events(self, events, calendar_name='Outlook Calendar Import', timezone='America/Los_Angeles'):
        """Export all events to Google Calendar.
        
        Args:
            events: List of event dictionaries
            calendar_name: Name for the calendar
            timezone: IANA timezone string for the events
        """
        if not self.service:
            print("Error: Not authenticated. Call authenticate() first.", file=sys.stderr)
            return False
        
        # Create or get calendar
        calendar_id = self.create_calendar(calendar_name, timezone)
        if not calendar_id:
            return False
        
        print(f"\nExporting {len(events)} events to Google Calendar (timezone: {timezone})...")
        success_count = 0
        
        for i, event in enumerate(events):
            link = self.export_event(event, calendar_id, timezone)
            if link:
                success_count += 1
                if (i + 1) % 10 == 0:
                    print(f"  Exported {i + 1}/{len(events)} events...")
        
        print(f"\n✓ Successfully exported {success_count}/{len(events)} events to Google Calendar!")
        print(f"  View your calendar at: https://calendar.google.com")
        return True


def merge_events(monthly_events, weekly_events):
    """Merge events from weekly calendar to enhance monthly calendar events.
    
    Args:
        monthly_events: List of events from monthly calendar
        weekly_events: List of events from weekly calendar
        
    Returns:
        Enhanced list of events with details from weekly calendar merged in
    """
    # Create a lookup dictionary for weekly events by (summary, start_time, end_time)
    weekly_lookup = {}
    for event in weekly_events:
        key = (
            event['summary'],
            event['start'].strftime('%Y-%m-%d %H:%M'),
            event['end'].strftime('%H:%M')
        )
        weekly_lookup[key] = event
    
    # Enhance monthly events with weekly calendar details
    enhanced_events = []
    for event in monthly_events:
        key = (
            event['summary'],
            event['start'].strftime('%Y-%m-%d %H:%M'),
            event['end'].strftime('%H:%M')
        )
        
        # Check if we have a matching event in weekly calendar
        if key in weekly_lookup:
            weekly_event = weekly_lookup[key]
            # Enhance with location if available in weekly view
            if weekly_event.get('location') and not event.get('location'):
                event['location'] = weekly_event['location']
            # Enhance with other details if needed
            # (weekly calendar might have more accurate organizer, status, etc.)
            if weekly_event.get('organizer') and not event.get('organizer'):
                event['organizer'] = weekly_event['organizer']
        
        enhanced_events.append(event)
    
    # Add any events from weekly calendar that weren't in monthly calendar
    monthly_keys = {
        (e['summary'], e['start'].strftime('%Y-%m-%d %H:%M'), e['end'].strftime('%H:%M'))
        for e in monthly_events
    }
    
    for event in weekly_events:
        key = (
            event['summary'],
            event['start'].strftime('%Y-%m-%d %H:%M'),
            event['end'].strftime('%H:%M')
        )
        if key not in monthly_keys:
            enhanced_events.append(event)
    
    return enhanced_events


def main():
    """Main function to convert Outlook HTML to ICS or Google Calendar."""
    # Parse command-line arguments
    parser = argparse.ArgumentParser(
        description='Extract calendar events from Outlook HTML and export to ICS or Google Calendar',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Export to ICS file (default)
  %(prog)s calendar.html
  
  # Export to your default Google Calendar (OAuth2 - browser login)
  %(prog)s calendar.html --google
  
  # Export with Eastern Time timezone (if events are in ET)
  %(prog)s calendar.html --google --timezone America/New_York
  
  # Export to a separate/new Google Calendar
  %(prog)s calendar.html --google --calendar-name "My Outlook Events"
  
  # Specify output file and domain
  %(prog)s calendar.html --output events.ics --domain domain.com
  
  # Advanced: Service Account (for automation, no browser)
  %(prog)s calendar.html --google --service-account \\
    --credentials service-account.json --calendar-id "your@email.com"
        """
    )
    
    parser.add_argument(
        'input_file',
        nargs='?',
        default=os.getenv('OUTLOOK_INPUT_FILE', 'calendar.mhtml'),
        help='Input HTML or MHTML file from Outlook calendar (default: from .env OUTLOOK_INPUT_FILE or calendar.mhtml)'
    )
    
    parser.add_argument(
        '-o', '--output',
        default=os.getenv('OUTLOOK_OUTPUT_FILE', 'outlook_calendar.ics'),
        help='Output ICS file path (default: from .env OUTLOOK_OUTPUT_FILE or outlook_calendar.ics)'
    )
    
    parser.add_argument(
        '-d', '--domain',
        default=os.getenv('OUTLOOK_EMAIL_DOMAIN', 'domain.com'),
        help='Email domain for organizer addresses (default: from .env OUTLOOK_EMAIL_DOMAIN or domain.com)'
    )
    
    parser.add_argument(
        '-g', '--google',
        action='store_true',
        help='Export to Google Calendar instead of ICS file'
    )
    
    parser.add_argument(
        '-c', '--calendar-name',
        default=os.getenv('GOOGLE_CALENDAR_NAME'),
        help='Name for a new/separate Google Calendar (default: from .env GOOGLE_CALENDAR_NAME or uses your primary calendar)'
    )
    
    parser.add_argument(
        '--credentials',
        default=os.getenv('GOOGLE_CREDENTIALS_FILE', 'credentials.json'),
        help='Path to Google Calendar API credentials file (default: from .env GOOGLE_CREDENTIALS_FILE or credentials.json)'
    )
    
    parser.add_argument(
        '--service-account',
        action='store_true',
        default=os.getenv('GOOGLE_USE_SERVICE_ACCOUNT', '').lower() in ('true', '1', 'yes'),
        help='Use service account authentication (static key) instead of OAuth2 (default: from .env GOOGLE_USE_SERVICE_ACCOUNT)'
    )
    
    parser.add_argument(
        '--calendar-id',
        default=os.getenv('GOOGLE_CALENDAR_ID'),
        help='Calendar ID to use (for service accounts, use email address of shared calendar) (default: from .env GOOGLE_CALENDAR_ID)'
    )
    
    parser.add_argument(
        '-tz', '--timezone',
        default=os.getenv('OUTLOOK_TIMEZONE', 'America/Los_Angeles'),
        help='IANA timezone for events (default: from .env OUTLOOK_TIMEZONE or America/Los_Angeles). Common: America/New_York (ET), America/Chicago (CT), America/Denver (MT)'
    )
    
    parser.add_argument(
        '-v', '--verbose',
        action='store_true',
        help='Enable verbose output for debugging (shows duplicate detection details)'
    )
    
    parser.add_argument(
        '--weekly-calendar',
        default=os.getenv('OUTLOOK_WEEKLY_FILE'),
        help='Optional weekly calendar HTML/MHTML file to enhance events with additional details (default: from .env OUTLOOK_WEEKLY_FILE)'
    )
    
    args = parser.parse_args()
    
    input_file = Path(args.input_file)
    output_file = Path(args.output)
    email_domain = args.domain
    
    # Check if input file exists
    if not input_file.exists():
        print(f"Error: Input file '{input_file}' not found.", file=sys.stderr)
        sys.exit(1)
    
    print(f"Reading calendar data from: {input_file}")
    
    # Try to read as MHTML first, then fall back to HTML
    html_content = None
    try:
        # First, try to parse as MHTML
        html_content = extract_html_from_mhtml(input_file)
        if html_content:
            print("Detected MHTML format, extracted HTML content")
        else:
            # Not MHTML, read as plain HTML
            with open(input_file, 'r', encoding='utf-8') as f:
                html_content = f.read()
            print("Reading as HTML format")
    except Exception as e:
        print(f"Error reading file: {e}", file=sys.stderr)
        sys.exit(1)
    
    if not html_content:
        print("Error: Could not extract HTML content from file", file=sys.stderr)
        sys.exit(1)
    
    # Parse events from monthly calendar
    print("Parsing events from monthly calendar...")
    parser = OutlookEventParser()
    parser.feed(html_content)
    
    monthly_events = parser.events
    print(f"Found {len(monthly_events)} events in monthly calendar")
    
    # Parse weekly calendar if provided
    weekly_events = []
    if args.weekly_calendar:
        weekly_file = Path(args.weekly_calendar)
        if weekly_file.exists():
            print(f"\nParsing events from weekly calendar: {weekly_file}")
            try:
                weekly_html = extract_html_from_mhtml(weekly_file)
                if not weekly_html:
                    with open(weekly_file, 'r', encoding='utf-8') as f:
                        weekly_html = f.read()
                
                weekly_parser = OutlookEventParser()
                weekly_parser.feed(weekly_html)
                weekly_events = weekly_parser.events
                print(f"Found {len(weekly_events)} events in weekly calendar")
                
                # Merge weekly events to enhance monthly events
                events = merge_events(monthly_events, weekly_events)
                print(f"Enhanced to {len(events)} total events (merged weekly calendar details)")
            except Exception as e:
                print(f"Warning: Could not parse weekly calendar: {e}", file=sys.stderr)
                print("Continuing with monthly calendar only...", file=sys.stderr)
                events = monthly_events
        else:
            print(f"Warning: Weekly calendar file '{weekly_file}' not found, using monthly calendar only", file=sys.stderr)
            events = monthly_events
    else:
        events = monthly_events
    
    if not events:
        print("Warning: No events found in the HTML file.", file=sys.stderr)
        sys.exit(1)
    
    # Display sample events
    print("\nSample events:")
    for i, event in enumerate(events[:5]):
        print(f"{i+1}. {event['summary']}")
        print(f"   {event['start'].strftime('%Y-%m-%d %H:%M')} - {event['end'].strftime('%H:%M')}")
        if event.get('location'):
            print(f"   Location: {event['location']}")
        print(f"   Organizer: {event['organizer']}")
        print(f"   Status: {event['status']}")
        if event['event_type']:
            print(f"   Type: {event['event_type']}")
        print()
    
    # Export events
    if args.google:
        # Export to Google Calendar
        if not GOOGLE_CALENDAR_AVAILABLE:
            print("\nError: Google Calendar API libraries not installed.", file=sys.stderr)
            print("Install them with: pip install -r requirements.txt", file=sys.stderr)
            sys.exit(1)
        
        print(f"\n{'='*60}")
        print("Exporting to Google Calendar")
        print(f"{'='*60}")
        
        try:
            exporter = GoogleCalendarExporter(
                credentials_file=args.credentials,
                use_service_account=args.service_account
            )
            print("\nAuthenticating with Google...")
            exporter.authenticate()
            print("✓ Authentication successful!")
            
            # Determine calendar ID
            if args.calendar_id:
                # Use specified calendar ID
                calendar_id = args.calendar_id
                print(f"Using specified calendar ID: {calendar_id}")
            elif args.calendar_name and args.calendar_name != 'primary':
                # Create or use existing calendar with custom name
                calendar_id = exporter.create_calendar(args.calendar_name, timezone=args.timezone)
                if not calendar_id:
                    print("Failed to create/access calendar.", file=sys.stderr)
                    sys.exit(1)
            else:
                # Use primary calendar (default)
                calendar_id = 'primary'
                print("Using your default Google Calendar")
            
            # Export events to the selected calendar
            created_count = 0
            updated_count = 0
            skipped_count = 0
            error_count = 0
            
            print(f"\nExporting {len(events)} events to Google Calendar (timezone: {args.timezone})...")
            print("  Checking for duplicates and changes...")
            
            for i, event in enumerate(events):
                action, link = exporter.export_event(event, calendar_id, timezone=args.timezone, verbose=args.verbose)
                
                if action == 'created':
                    created_count += 1
                elif action == 'updated':
                    updated_count += 1
                elif action == 'skipped':
                    skipped_count += 1
                elif action == 'error':
                    error_count += 1
                
                # Show progress every 10 events (skip if verbose mode is on)
                if not args.verbose and (i + 1) % 10 == 0:
                    print(f"  Processed {i + 1}/{len(events)} events... (created: {created_count}, updated: {updated_count}, skipped: {skipped_count})")
            
            success = (created_count + updated_count + skipped_count) > 0
            
            if success:
                print(f"\n{'='*60}")
                print("Export Complete!")
                print(f"{'='*60}")
                print(f"Created: {created_count} new events")
                print(f"Updated: {updated_count} changed events")
                print(f"Skipped: {skipped_count} unchanged events")
                if error_count > 0:
                    print(f"Errors: {error_count} events failed")
                print(f"\nTotal events processed: {len(events)}")
                if args.calendar_name:
                    print(f"Calendar name: {args.calendar_name}")
                print(f"View at: https://calendar.google.com")
            else:
                print("\nExport failed. Please check the errors above.", file=sys.stderr)
                sys.exit(1)
                
        except FileNotFoundError as e:
            print(f"\n{e}", file=sys.stderr)
            print("\nTo set up Google Calendar API:", file=sys.stderr)
            print("  1. Visit https://console.cloud.google.com", file=sys.stderr)
            print("  2. Create a new project", file=sys.stderr)
            print("  3. Enable Google Calendar API", file=sys.stderr)
            print("  4. Create OAuth 2.0 credentials", file=sys.stderr)
            print("  5. Download credentials.json", file=sys.stderr)
            print("\nSee README for detailed instructions.", file=sys.stderr)
            sys.exit(1)
        except Exception as e:
            print(f"\nError exporting to Google Calendar: {e}", file=sys.stderr)
            sys.exit(1)
    else:
        # Generate ICS file
        print(f"\nGenerating ICS file with domain: {email_domain}, timezone: {args.timezone}")
        generator = ICSGenerator(email_domain, args.timezone)
        ics_content = generator.generate(events)
        
        # Write ICS file
        try:
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(ics_content)
            print(f"\nSuccess! ICS file created: {output_file}")
            print(f"Total events exported: {len(events)}")
            print(f"\nYou can now import '{output_file}' into macOS Calendar:")
            print(f"  1. Double-click the ICS file, or")
            print(f"  2. Open Calendar.app and go to File > Import > Choose '{output_file}'")
        except Exception as e:
            print(f"Error writing ICS file: {e}", file=sys.stderr)
            sys.exit(1)


if __name__ == "__main__":
    main()

