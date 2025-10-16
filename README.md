# Outlook Calendar to ICS/Google Converter

This tool extracts calendar events from an Outlook HTML or MHTML export and converts them to ICS (iCalendar) format, or exports them directly to Google Calendar.

## Features

- ✅ Extracts event details from Outlook HTML or MHTML export
- ✅ **Two export methods:**
  - **ICS file** - Compatible with macOS Calendar, Outlook, and other calendar apps
  - **Google Calendar** - Direct export to your default calendar with OAuth2 (browser login)
- ✅ **Smart duplicate detection** - Safe to run multiple times, only updates changed events
- ✅ Preserves event information:
  - Event title/summary
  - Start and end times
  - Organizer information
  - Event status (Tentative, Busy, Free)
  - Recurring event indicators
  - Canceled event status
- ✅ No external dependencies for ICS export
- ✅ Simple OAuth2 authentication for Google Calendar
- ⚙️ Advanced: Service account support for automation

## Requirements

- Python 3.6 or higher

**For ICS Export:** No additional dependencies needed (but `.env` file support requires `python-dotenv`)

**For Google Calendar Export and .env support:** Install dependencies:
```bash
python3 -m venv .
source bin/activate  # On Windows: .\bin\activate
pip install -r requirements.txt
```

This installs:
- Google Calendar API libraries (for `--google` export)
- `python-dotenv` (for `.env` configuration file support)

## Getting Your Outlook Calendar Export

### HTML Format (Recommended)
1. Open Outlook Calendar (desktop or web)
2. Go to File > Save As (desktop) or use browser's "Save Page" feature
3. Save as "Web Page, Complete (*.html)" or "Web Page, HTML only"

### MHTML Format (Alternative)
1. Open Outlook Calendar in a web browser
2. Press `Ctrl+S` (Windows) or `Cmd+S` (Mac) to save
3. Choose "Webpage, Single File (*.mhtml)" format
4. Save the file

**Both formats are supported** - the script automatically detects whether you're providing an HTML or MHTML file and extracts the calendar data accordingly.

## Configuration

### Using a .env File (Recommended)

You can configure default settings using a `.env` file instead of passing command-line arguments every time:

1. **Copy the example file:**
   ```bash
   cp .env.example .env
   ```

2. **Edit `.env` with your settings:**
   ```bash
   # Input/Output Settings
   OUTLOOK_INPUT_FILE=calendar.mhtml
   OUTLOOK_OUTPUT_FILE=outlook_calendar.ics
   OUTLOOK_EMAIL_DOMAIN=yourdomain.com
   OUTLOOK_TIMEZONE=America/Los_Angeles
   
   # Teams Meeting Link (optional - adds meeting link to all events)
   TEAMS_MEETING_LINK=https://teams.microsoft.com/l/meetup-join/...
   
   # Google Calendar Settings
   GOOGLE_CREDENTIALS_FILE=credentials.json
   GOOGLE_CALENDAR_NAME=
   GOOGLE_CALENDAR_ID=
   GOOGLE_USE_SERVICE_ACCOUNT=false
   ```

3. **Run the script without arguments:**
   ```bash
   python3 outlook_to_ics.py --google
   ```

**Benefits:**
- ✅ No need to type long filenames or settings each time
- ✅ Keep sensitive settings (domain, Teams links) out of command history
- ✅ Easy to maintain multiple configurations (e.g., `.env.work`, `.env.personal`)
- ✅ Command-line arguments still work and override `.env` settings

**Note:** The `.env` file is automatically ignored by git to keep your settings private.

## Usage

### Export to ICS File (Default)

Basic usage - exports to `outlook_calendar.ics`:

```bash
python3 outlook_to_ics.py
```

Specify input file (HTML or MHTML):

```bash
python3 outlook_to_ics.py "calendar.html"
# or
python3 outlook_to_ics.py "calendar.mhtml"
```

Specify output file:

```bash
python3 outlook_to_ics.py "calendar.html" --output "my_events.ics"
```

Specify email domain:

```bash
python3 outlook_to_ics.py "calendar.html" --domain "yourdomain.com"
```

### Export to Google Calendar

Export directly to your **default Google Calendar** (OAuth2):

```bash
python3 outlook_to_ics.py "calendar.html" --google
```

**Smart Import:** The script automatically detects duplicate events and only creates or updates events when there are changes. This means:
- ✅ Safe to run multiple times
- ✅ No duplicate events created
- ✅ Only changed events are updated
- ✅ Unchanged events are skipped

**Debug duplicate detection:**
```bash
# Run with verbose flag to see duplicate detection in action
python3 outlook_to_ics.py "calendar.html" --google --verbose
```

**Specify timezone** (defaults to Pacific Time):

```bash
# Default: Pacific Time (PT) - America/Los_Angeles
python3 outlook_to_ics.py "calendar.html" --google

# Events are in Eastern Time (ET) - America/New_York
python3 outlook_to_ics.py "calendar.html" --google --timezone America/New_York

# Events are in Central Time (CT) - America/Chicago
python3 outlook_to_ics.py "calendar.html" --google --timezone America/Chicago

# Events are in Mountain Time (MT) - America/Denver
python3 outlook_to_ics.py "calendar.html" --google --timezone America/Denver
```

**Create a separate calendar** (optional):

```bash
# Create a new calendar instead of using your default calendar
python3 outlook_to_ics.py "calendar.html" --google --calendar-name "My Outlook Events"
```

For automation/scripts, see [Service Account Authentication](#advanced-service-account-authentication)

### View All Options

```bash
python3 outlook_to_ics.py --help
```

**Note:** The script defaults to `domain.com` as the email domain (from the Outlook realm). You can override this with `--domain` option.

### Make the Script Executable (Optional)

```bash
chmod +x outlook_to_ics.py
./outlook_to_ics.py
```

## How to Export Calendar from Outlook

1. Open Outlook Web (outlook.office365.us or outlook.office.com)
2. Navigate to your Calendar
3. In your browser, save the page as HTML:
   - Chrome/Edge: Press `Cmd+S` (Mac) or `Ctrl+S` (Windows)
   - Choose "Webpage, Complete" or "HTML Only"
   - Save the file

## Setting Up Google Calendar API (For Google Calendar Export)

**Quick Start:** Use OAuth2 authentication (browser-based, simple setup)

**Advanced:** Service Account authentication available for automation (see [Advanced Options](#advanced-service-account-authentication) below)

### Step 1: Create a Google Cloud Project

1. Go to [Google Cloud Console](https://console.cloud.google.com)
2. Click **"Select a project"** → **"New Project"**
3. Enter a project name (e.g., "Outlook Calendar Import")
4. Click **"Create"**

### Step 2: Enable Google Calendar API

1. In your project, go to **"APIs & Services"** → **"Library"**
2. Search for **"Google Calendar API"**
3. Click on it and press **"Enable"**

### Step 3: Create OAuth 2.0 Credentials

1. Go to **"APIs & Services"** → **"Credentials"**
2. Click **"Create Credentials"** → **"OAuth client ID"**
3. If prompted, configure the OAuth consent screen:
   - User Type: **"External"** (or "Internal" for workspace users)
   - App name: Enter your app name
   - User support email: Your email
   - Developer contact: Your email
   - Click **"Save and Continue"**
   - Scopes: Click **"Save and Continue"** (default is fine)
   - Test users: Add your email address
   - Click **"Save and Continue"**
4. Back to Create OAuth client ID:
   - Application type: **"Desktop app"**
   - Name: "Outlook Calendar Import Client"
   - Click **"Create"**
5. Click **"Download JSON"** and save as `credentials.json` in the same directory as the script

### Step 4: Install Required Python Packages

```bash
pip install -r requirements.txt
```

### Step 5: Run the Script

```bash
python3 outlook_to_ics.py "calendar.html" --google
```

On first run:
- A browser window will open
- Sign in with your Google account
- Click **"Allow"** to grant calendar access
- The script will save authentication token for future use (`token.pickle`)
- Subsequent runs won't require browser login!

## Importing to macOS Calendar

After generating the ICS file, you can import it to macOS Calendar in two ways:

### Method 1: Double-click
Simply double-click the generated `.ics` file, and Calendar will open and import the events.

### Method 2: Import Menu
1. Open Calendar.app
2. Go to `File` → `Import`
3. Select the generated `.ics` file
4. Choose which calendar to import to
5. Click `Import`

## Event Information Extracted

The script extracts the following information from each event:

- **Summary**: Event title/name
- **Start/End Time**: Date and time of the event
- **Organizer**: Person who created the event (with email address using @domain.com domain)
- **Status**: 
  - `TENTATIVE` for tentative events
  - `CONFIRMED` for busy/free events
  - `CANCELLED` for canceled events
- **Event Type**: Recurring, Exception to recurring, etc.
- **Transparency**: 
  - `OPAQUE` for busy events (blocks time)
  - `TRANSPARENT` for free events (doesn't block time)

## Output Example

```
Reading calendar data from: calendar.html
Parsing events...
Found 81 events

Sample events:
1. Product Operations Daily
   2025-09-29 10:30 - 11:00
   Organizer: Dama
   Status: Tentative
   Type: Recurring event

2. Meeting Rooms
   2025-09-30 06:00 - 13:30
   Organizer: Sam
   Status: Busy
   Type: Recurring event

Generating ICS file...

Success! ICS file created: outlook_calendar.ics
Total events exported: 81
```

## Troubleshooting

### No events found
- Make sure you saved the Outlook calendar page as HTML
- Ensure the HTML file contains the calendar view (not a settings or other page)
- Try saving the page in "month view" for best results

### Parsing errors
- Some events with unusual formatting may not parse correctly
- Check the console output for warning messages about specific events
- The script will continue processing other events even if some fail

### Import issues
- If Calendar doesn't recognize the file, make sure it has the `.ics` extension
- Try opening the ICS file in a text editor to verify it's not corrupted

## Technical Details

### ICS Format
The script generates ICS files following the iCalendar (RFC 5545) specification, ensuring compatibility with most calendar applications.

### Timezone
Events are exported with timezone information using the `--timezone` parameter (defaults to `America/Los_Angeles` - Pacific Time). 

**Important:** The Outlook HTML export doesn't include timezone information, so you need to specify the correct timezone that matches your Outlook calendar's original timezone.

Common IANA timezone identifiers:
- `America/Los_Angeles` - Pacific Time (PT) - **Default**
- `America/New_York` - Eastern Time (ET)
- `America/Chicago` - Central Time (CT)
- `America/Denver` - Mountain Time (MT)
- `Europe/London` - GMT/BST
- `Europe/Paris` - CET/CEST

**Example:** If your Outlook calendar shows events in Eastern Time, use:
```bash
python3 outlook_to_ics.py "calendar.html" --google --timezone America/New_York
```
This ensures events display at the correct time in Google Calendar regardless of your current timezone.

### Recurring Events
Note that recurring events from Outlook are exported as individual event instances, not as true recurring events with recurrence rules. Each occurrence becomes a separate event in the ICS file.

### Duplicate Detection & Smart Updates

When exporting to Google Calendar, the script intelligently handles existing events:

**How it works:**
1. Before creating an event, the script searches for existing events with matching:
   - Event title (summary)
   - Start date and time
2. If a match is found, it compares all event details:
   - Title, start/end times, description, status, transparency
3. Based on the comparison:
   - **Create**: Event doesn't exist → creates new event
   - **Update**: Event exists but details changed → updates existing event
   - **Skip**: Event exists and is identical → no action taken

**Benefits:**
- ✅ Run the script multiple times without creating duplicates
- ✅ Sync updated events from Outlook to Google Calendar
- ✅ Efficient - doesn't make unnecessary API calls
- ✅ Reports statistics: created, updated, and skipped counts

**Example output:**
```
Exporting 81 events to Google Calendar (timezone: America/Los_Angeles)...
  Checking for duplicates and changes...
  Processed 10/81 events... (created: 8, updated: 1, skipped: 1)
  ...

Export Complete!
Created: 75 new events
Updated: 3 changed events
Skipped: 3 unchanged events
Total events processed: 81
```

## Comparing Export Methods

| Feature | ICS File | Google Calendar |
|---------|----------|-----------------|
| Setup Required | None | Google Cloud API setup |
| Dependencies | None | `pip install -r requirements.txt` |
| Internet Required | No | Yes |
| Authentication | None | OAuth 2.0 (one-time browser login) |
| Import Process | Manual import to calendar app | Automatic |
| Portability | Works with any calendar app | Google Calendar only |
| Destination | Any calendar app | Default Google Calendar |
| Use Case | One-time import, multiple apps | Direct sync to Google |

**Recommendation:**
- Use **ICS export** for quick imports to macOS Calendar or other apps
- Use **Google Calendar export** for direct import to your default Google Calendar

### Google Calendar: Default vs Separate Calendar

By default, events are imported into your **primary Google Calendar** (the one you see when you open Google Calendar).

**When to use default calendar:**
- ✅ You want events mixed with your regular Google Calendar events
- ✅ Simplest option - no extra calendars to manage
- ✅ Events appear immediately in your main calendar

**When to create a separate calendar:**
- 📅 You want to keep Outlook events separate
- 🎨 You want different colors for these events
- 🗑️ You might want to delete all imported events later
- 👁️ You want to toggle visibility of Outlook events

```bash
# Use default calendar (recommended)
python3 outlook_to_ics.py "calendar.html" --google

# Create separate calendar
python3 outlook_to_ics.py "calendar.html" --google --calendar-name "Outlook Import"
```

## Advanced: Service Account Authentication

For automation, scripts, or CI/CD pipelines, you can use service account authentication instead of OAuth2. This provides a static JSON key file that doesn't require browser interaction.

### When to Use Service Accounts

- ✅ Automated scripts and cron jobs
- ✅ CI/CD pipelines
- ✅ Server-side applications
- ✅ No browser access available
- ⚠️ Requires additional setup (calendar sharing)

### Service Account Setup

#### Step 1: Create Service Account

1. Go to [Google Cloud Console](https://console.cloud.google.com) → Your Project
2. Go to **"APIs & Services"** → **"Credentials"**
3. Click **"Create Credentials"** → **"Service Account"**
4. Enter service account details:
   - Name: "Calendar Import Service"
   - Click **"Create and Continue"** → **"Continue"** → **"Done"**

#### Step 2: Create and Download Key

1. Click on the service account you just created
2. Go to the **"Keys"** tab
3. Click **"Add Key"** → **"Create new key"**
4. Choose **"JSON"** format
5. Click **"Create"**
6. Save the downloaded JSON file as `service-account.json`
   - **Important:** Keep this file secure!

#### Step 3: Share Your Calendar

1. Copy the service account email (e.g., `calendar-import@project.iam.gserviceaccount.com`)
2. Open [Google Calendar](https://calendar.google.com)
3. Click on your calendar → **"Settings and sharing"**
4. Under **"Share with specific people"**, click **"Add people"**
5. Paste the service account email
6. Set permissions to **"Make changes to events"**
7. Click **"Send"**
8. Note your **Calendar ID** (in settings under "Integrate calendar")

#### Step 4: Use Service Account

```bash
python3 outlook_to_ics.py "calendar.html" --google \
  --credentials service-account.json \
  --service-account \
  --calendar-id "your-email@gmail.com"
```

**No browser interaction required!**

### OAuth2 vs Service Account

| Feature | OAuth2 | Service Account |
|---------|--------|-----------------|
| Setup | Simpler | More complex |
| Browser | Required (first time) | Not required |
| Calendar Access | Automatic | Must share calendar |
| Use Case | Personal use | Automation/scripts |

**For most users:** Stick with OAuth2 (default)

## Files in This Repository

- `outlook_to_ics.py` - Main conversion script with dual export support
- `requirements.txt` - Python dependencies for Google Calendar integration
- `README.md` - This documentation file
- `calendar.html` - Sample Outlook HTML export (example)
- `outlook_calendar.ics` - Generated ICS file (example output)
- `credentials.json` - OAuth2 credentials (you create this for interactive auth)
- `service-account.json` - Service account key (you create this for static key auth)
- `service-account.json.template` - Template showing service account file structure
- `token.pickle` - Saved OAuth2 token (auto-generated after first login)
- `.gitignore` - Protects sensitive credential files

## License

This tool is provided as-is for personal use. Feel free to modify and distribute as needed.

## Contributing

If you find issues or have suggestions for improvements, please feel free to submit issues or pull requests.



