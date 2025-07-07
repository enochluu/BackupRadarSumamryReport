# Backup Radar Summary Report Generator

This script connects to the Backup Radar API and generates a formatted Excel report containing backup job statuses from the previous day (in AEST). It is designed to assist MSPs and IT teams in reviewing and triaging failed, warning, or pending backup jobs, with dedicated formatting for Microsoft 365 Acronis-related backups.

## Features

- Fetches scheduled backup jobs from the previous day (AEST)
- Filters for specific statuses: `Failure`, `Warning`, `No Result`, and `Pending`
- Separates regular jobs and Microsoft 365 Acronis jobs
- Groups jobs by client
- Generates a formatted Excel report with:
  - Styled headers and zebra striping for readability
  - Conditional formatting to highlight resolution status
  - Dropdown selection for resolved/unresolved
  - Manual entry fields for ticket numbers and technician notes
- Auto-sizing of columns for a clean layout

## Requirements

Install the required dependencies with:

```
pip install -r requirements.txt
```

Contents of `requirements.txt`:

```
requests
pytz
python-dotenv
openpyxl
```

## Environment Setup

Create a `.env` file in the root directory with your Backup Radar API key:

```
API_KEY=your_api_key_here
```

## Usage

Run the script:

```
python BackupRadarSummaryReport.py
```

If successful, an Excel file will be generated with the following naming format:

```
enhanced_backup_report_YYYY-MM-DD_AEST.xlsx
```

## Excel Output Structure

The output Excel report includes two sections:

### Regular Backups

All backup jobs excluding Microsoft 365 Acronis-specific jobs. Jobs are grouped by client and include the following columns:

- Server/Workload Affected
- Status
- Job Name
- Backup Method
- Resolved (✔ / ✘ dropdown)
- Ticket number
- Technician Notes

### Microsoft 365 Acronis Backups

Jobs containing any of the following keywords are categorised separately:

- OneDrive to Cloud storage
- Office 365 mailboxes to Cloud storage
- SharePoint sites to Cloud storage
- Microsoft 365 mailboxes to Cloud storage
- Microsoft Teams to Cloud storage

These jobs are presented in a separate section after the regular backups with the same formatting and input structure.
