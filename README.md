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

python outlook_attachment_downloader.py
```

Run it with no arguments and it stays open, connected to Outlook once,
prompting for commands until you tell it to stop:

```
Outlook Attachment Downloader -- interactive mode
Type a command using the same flags as the command line, e.g.:
  --sender alice@example.com --since 2026-06-01 --extensions pdf
Type "help" for the full list of options, "exit" or "quit" to leave.

outlook> --since 2026-05-22
Scanning "Inbox" (47 items)...
  saved: C:\...\attachments\20260524_0931_Bob_Invoice\invoice.pdf
...
outlook> --sender alice@example.com --extensions pdf,xlsx --organize flat
...
outlook> exit
Goodbye.
```

Every flag below works exactly the same way whether you type it after the
`outlook>` prompt or pass it on the command line in one shot:

```bash
# Single-shot mode still works too -- pass any flags and it runs once and exits
python outlook_attachment_downloader.py --since 2026-05-22

# Not sure of your exact folder names/paths? List them first.
python outlook_attachment_downloader.py --list-folders

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
| `--output` / `-o` | Output directory. Default: `./attachments`. Always resolved to an absolute path internally (see gotcha below). |
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
| `--no-mark-read` | Skip marking matched emails as read (marking read is the default whenever an attachment is actually saved). |
| `--list-folders` | Print the full folder tree and exit. |

## Notes / gotchas

- **Matched emails are marked read by default.** Whenever an attachment
  actually gets saved from an email, that email is set to read (`UnRead =
  False`) afterward -- harmless if it was already read. Use
  `--no-mark-read` if you want to leave read/unread state untouched.
- **Relative `--output` paths used to fail every save.** `Attachment.SaveAsFile()`
  is a COM call into the separate OUTLOOK.EXE process, which resolves a
  relative path against *its own* working directory, not this script's.
  That produced a 100%-failure pattern: Python's `mkdir()` happily creates
  the folder (in your script's cwd), but Outlook then looks for that same
  relative path somewhere else and reports `Cannot save the attachment.
  Path does not exist.` — even though the folder really does exist. Fixed
  by resolving `--output` to an absolute path before anything touches it.
  If you ever see that exact error again, an absolute path is the first
  thing to check.
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
  profile. It's not a great fit for a headless service account. If you
  eventually need unattended/server-side access, that's exactly the case
  for switching to the Microsoft Graph + `O365` (Apache 2.0) path instead.

## Building a standalone .exe (for users without Python)

Same approach as the PDF bookmark tool: PyInstaller bundles the script
and a Python runtime into one `.exe`, so recipients don't need Python
installed at all.

**I can't hand you the actual binary from here** — PyInstaller doesn't
cross-compile, so a Windows `.exe` has to be built on Windows. That's not
really an extra burden though, since this script already has to run on
Windows (it talks to the local Outlook COM interface), so you'll be
building it on the same machine where you've been testing it.

### Steps

1. On that Windows machine: `pip install pyinstaller`
2. Run `build_exe.bat` (included), or directly:
   ```
   pyinstaller --onefile --name outlook_attachment_downloader --hidden-import win32timezone outlook_attachment_downloader.py
   ```
3. Grab `dist\outlook_attachment_downloader.exe` — that single file is
   what you hand to other people.

### What the .exe does and doesn't remove

- **Removes:** the need for Python/pip to be installed. Double-clicking
  the `.exe` now drops you into the interactive prompt instead of running
  once with default settings and immediately closing the window.
- **Does NOT remove:** the need for Outlook desktop to be installed and
  signed in on whoever's machine runs it. This is still COM automation
  against a local, signed-in Outlook profile — each person running the
  `.exe` pulls attachments from *their own* signed-in mailbox, not a
  shared one. Centrally pulling attachments out of *other* people's
  mailboxes is a different job (the Microsoft Graph + `O365` path
  mentioned above), not something this `.exe` does.

### Heads-up: SmartScreen / antivirus

Unsigned PyInstaller `--onefile` executables get flagged by Windows
Defender SmartScreen or other antivirus as "unrecognized" fairly often —
a well-known false positive for this packaging style, not a sign of an
actual problem. Recipients may need to click "More info" → "Run anyway."
Code-signing the exe avoids this if you're distributing it widely.

## GUI version

`outlook_attachment_downloader_gui.py` is a windowed front-end for the
same tool. It imports and reuses the core logic from
`outlook_attachment_downloader.py` directly (folder traversal, filtering,
saving, stats) — both files need to be in the same folder.

**No new dependencies.** The GUI is built entirely with `tkinter`/`ttk`,
which ship with the Python standard library — the bindings are
PSF-licensed, the underlying Tcl/Tk is BSD-style licensed, both fully
permissive for commercial use. No third-party GUI toolkit was added,
since tkinter already covers everything needed (tabs, a folder tree, the
output-folder picker dialog) with no extra weight or license surface.

### Running it

```bash
python outlook_attachment_downloader_gui.py
```

- **Filters & Options tab** — every CLI filter (sender, subject, date
  range, extensions, max emails, unread-only, include-inline, dry-run,
  mark-as-read, organize mode) as entries/checkboxes/radio buttons.
- **Outlook Folder tab** — click **Load Folders** to connect to Outlook
  and show the top level of your folder tree. Folders load lazily: a
  folder's children are only fetched the first time you expand it, not
  all up front, so this stays fast even on large mailboxes. Ctrl/Shift-click
  to select more than one folder; check **Include subfolders** to also
  scan everything under your selection. Nothing selected → defaults to
  Inbox, same as the CLI.
- **Output folder** — shown at the bottom, defaults to the same
  `./attachments` the CLI uses. Click **Change...** to pick a different
  folder; otherwise it just uses the default.
- **Download Attachments** — runs on a background thread so the window
  doesn't freeze on large folders, and disables itself while running.
  Progress and the final summary stream into the **Activity Log** panel
  at the bottom, visible no matter which tab you're on.

Date fields use the same `YYYY-MM-DD` format as the CLI (no calendar
picker, by design — see the GUI section in the build notes above for why).

### Building a standalone .exe

Same idea as the CLI build, but with `build_gui_exe.bat` / the `--windowed`
flag instead, since this is a real GUI now and shouldn't pop a console
box alongside it:

```
pyinstaller --onefile --windowed --name outlook_attachment_downloader_gui --hidden-import win32timezone outlook_attachment_downloader_gui.py
```

Both `.py` files need to be in the folder when you build (PyInstaller
bundles the imported core module automatically), but the resulting
`.exe` is fully standalone — recipients only need that one file, plus
Outlook installed and signed in, same as always.
