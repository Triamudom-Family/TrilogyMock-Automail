# TU90 Mock Automail

Automated Google Apps Script mailer that sends Dugga exam login credentials to students on form submission for **Mock x Triam The Trilogy for #TU90** by Triamudom Family.

## Files

| File | Description |
|------|-------------|
| `Code.gs` | Main Apps Script — trigger, credential lookup, email send |
| `inline.html` | HTML email template (`{username}`, `{password}` placeholders) |

## Functions

| Function | Description |
|---|---|
| `onFormSubmit(e)` | Triggered automatically on each form submission |
| `retryEmailByRow(row)` | Manually resend to a specific row (reuses stored credentials if available, does not increment counter) |
| `myFunction()` | Logs remaining daily email quota |

## Exam links

| | URL |
|---|---|
| Exam system | https://auth.dugga.com/ |
| User manual | Google Drive (see template) |

## License

AGPL-3.0 — Triamudom Family (`admin@triamudomfamily.org`)
