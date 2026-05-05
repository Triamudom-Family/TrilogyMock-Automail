# TU90 Mock Automail

HTML email template for sending Dugga exam login credentials to students for **Mock x Triam The Trilogy for #TU90** by Triamudom Family.

## Overview

`mail.html` is a single-file email template used as the body of automated credential emails. It displays:

- Student username and password for the Dugga exam platform
- Exam open/close dates
- Buttons linking to the exam system and user manual
- Step-by-step login instructions
- Contact info (email & Line OA)

Placeholder variables (`{username}`, `{password}`) are replaced at send time via Google Apps Script.

## Usage

1. Open the Google Apps Script project that sends the mail.
2. The script reads `mail.html`, replaces `{username}` and `{password}` with each student's credentials, and sends via `GmailApp.sendEmail()` or `MailApp.sendEmail()` with `htmlBody`.

## Files

| File | Description |
|------|-------------|
| `mail.html` | Email template (Tailwind CSS, IBM Plex Sans Thai) |

## Exam Links

| | URL |
|---|---|
| Exam system | https://auth.dugga.com/ |
| User manual | Google Drive (see template) |

## License

AGPL-3.0 — Triamudom Family (`admin@triamudomfamily.org`)
