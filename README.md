# VTU Diary Studio Chrome Extension

Chrome extension for VTU internship diary workflows.

## What It Does

- Export submitted diary entries using your active VTU session
- Generate PDF and DOC exports
- Bulk upload diary entries from JSON
- Keep logs for long-running operations

## Core Features

### Export Submitted Diaries

- Reads diary entries from VTU API endpoints through your current browser session
- Merges profile details from available sources
- Generates:
  - Internship_Diary.pdf
  - Internship_Diary.doc

### Bulk Upload JSON

- Accepts a single object or an array of objects
- Validates and normalizes data before upload
- Handles retries and row-level logging

Validation rules:
- Date: YYYY-MM-DD
- Hours: 1 to 24
- Description: required
- Learnings: optional
- Skills: optional

### Persistent Logs

- Supports long operations even if popup is closed
- Stores logs in local extension storage

## JSON Upload Schema

Supported input key variants:
- date / Date
- hours / Hours / hours_worked
- description / activity / work_description / workDescription / Work Description
- learnings / Learnings
- skills / Skills

Example:

```json
[
  {
    "Date": "2026-04-14",
    "Hours": 4,
    "Work Description": "Implemented API integration and validated payload flow.",
    "Learnings": "Understood retry strategy and schema normalization.",
    "Skills": ["python"]
  }
]
```

Sample file:
- [examples/2026-04-14.json](examples/2026-04-14.json)

## Project Files

- manifest.json: MV3 config, permissions, host access, service worker setup
- background.js: export/upload orchestration and runtime log handling
- popup.html: extension popup UI
- popup.js: UI behavior, messaging, and local profile persistence
- content.js: keeps workflow API-driven (no legacy UI automation)
- lib/jspdf.umd.min.js: PDF generation library

## Local Development

1. Open Chrome and go to chrome://extensions
2. Enable Developer mode
3. Click Load unpacked
4. Select this project folder
5. Keep an authenticated VTU session active before export/upload

## Packaging

1. Validate manifest details
2. Zip extension contents (without an extra parent folder layer)
3. Upload to Chrome Web Store Developer Dashboard

## Notes

- Export and upload depend on authenticated VTU cookies in your browser
- Operations are designed to continue while popup is closed
- Verify generated and uploaded data before final VTU submission

## Affiliate Disclosure

- The extension may show optional sponsored recommendations for student tools inside the extension popup only
- Affiliate links open only on explicit user click
- The extension does not auto-open, auto-redirect, or inject ads into VTU or other websites
- We may earn a commission from affiliate purchases at no extra cost to users
