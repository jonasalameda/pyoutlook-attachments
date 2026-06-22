# Outlook Attachment Downloader

A small CLI tool that scans folders in a locally-running, signed-in Outlook
desktop profile and saves email attachments to disk.

## How it talks to Outlook

This uses COM automation (`pywin32`) against the Outlook desktop app —
the same approach as your earlier email-reading script. No Azure app
registration, no OAuth, no internet calls. Outlook just needs to be
installed and signed in to the account you want to read.

(This replaces `pyOutlook`, which can no longer reach a real mailbox:
it's built on the Outlook REST API v2.0, which Microsoft fully
decommissioned on March 31, 2024 — every call now returns HTTP 410 Gone.)

## Requirements

- Windows, with Outlook desktop installed and signed in
- Python 3.8+
- `pip install -r requirements.txt` (installs `pywin32`)

**License note:** `pywin32` is licensed under the Python Software
Foundation License — OSI-approved and permissive, commercial use is fine.
Everything else the script uses is the Python standard library.

## Quick start

```bash
pip install -r requirements.txt

# Not sure of your exact folder names/paths? List them first.
python outlook_attachment_downloader.py --list-folders

# Save every attachment from the Inbox into ./attachments
python outlook_attachment_downloader.py

# Last 30 days only
python outlook_attachment_downloader.py --since 2026-05-22

# Only PDFs and Excel files, from one sender, all in one flat folder
python outlook_attachment_downloader.py --sender alice@example.com --extensions pdf,xlsx --organize flat

# Multiple folders, including their subfolders — preview first
python outlook_attachment_downloader.py --folder Inbox --folder "Inbox/Invoices" --recursive --dry-run
```

## Options

| Flag | Description |
|---|---|
| `--folder` / `-f` | Outlook folder path, e.g. `Inbox` or `Inbox/Invoices`. Repeatable. Default: `Inbox`. |
| `--output` / `-o` | Output directory. Default: `./attachments`. |
| `--since YYYY-MM-DD` | Only emails received on/after this date. |
| `--until YYYY-MM-DD` | Only emails received on/before this date. |
| `--sender TEXT` | Substring match against sender name or email address. |
| `--subject-contains TEXT` | Substring match against the subject. |
| `--unread-only` | Only process unread emails. |
| `--extensions pdf,docx,...` | Only save attachments with these extensions. |
| `--max-emails N` | Stop after N emails with saved attachments. |
| `--include-inline` | Also save hidden/inline attachments (e.g. signature logos). Skipped by default. |
| `--organize` | `by-email` (default, one subfolder per email), `by-date`, or `flat`. |
| `--recursive` | Also scan subfolders of each `--folder`. |
| `--dry-run` | Show what would be saved without writing anything. |
| `--mark-read` | Mark matched emails as read afterward. |
| `--list-folders` | Print the full folder tree and exit. |

## Notes / gotchas

- **Inline images are attachments too.** Outlook stores signature logos
  and embedded images as hidden attachments. They're skipped by default;
  use `--include-inline` if you actually want them.
- **Filename collisions** are handled automatically — a duplicate name
  gets `(1)`, `(2)`, etc. appended.
- **Date filtering** is done in Python after pulling the items, not via
  an Outlook `Restrict()` query — simpler and avoids Outlook's
  locale-sensitive date-string format, at the cost of being slower on
  very large folders.
- **Running this on a server / scheduled task:** COM automation against
  Outlook generally needs an interactive desktop session for the signed-in
  profile (similar to the constraints you ran into with the serial-port
  permissions issue on the Smart Store project, just a different flavor
  of "needs the right session/context"). It's not a great fit for a
  headless service account. If you eventually need unattended/server-side
  access, that's exactly the case for switching to the Microsoft
  Graph + `O365` (Apache 2.0) path instead.
