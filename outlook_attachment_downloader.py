#!/usr/bin/env python3
"""
outlook_attachment_downloader.py

Scans one or more folders in a locally-running, signed-in Outlook desktop
profile (Windows) and saves attachments from matching emails to disk.

Talks to Outlook via COM automation (pywin32) -- no Azure app registration,
no OAuth, no internet calls. Outlook just needs to be installed and signed
into the account you want to read.

Requirements:
    Windows + Microsoft Outlook desktop, signed in
    pip install pywin32

License note: this script's only third-party dependency is pywin32, which
is licensed under the Python Software Foundation License (OSI-approved,
permissive -- commercial use is fine). Everything else used here is the
Python standard library.

Examples:
    # Everything received in the last 30 days, default Inbox
    python outlook_attachment_downloader.py --since 2026-05-22

    # Only PDFs/Excel files from one sender, all in one flat folder
    python outlook_attachment_downloader.py --sender alice@example.com --extensions pdf,xlsx --organize flat

    # Multiple folders, including their subfolders, preview only
    python outlook_attachment_downloader.py --folder Inbox --folder "Inbox/Invoices" --recursive --dry-run

    # Not sure of exact folder names/paths? List them first.
    python outlook_attachment_downloader.py --list-folders
"""

import argparse
import datetime
import re
import sys
from dataclasses import dataclass
from pathlib import Path

# Outlook "default folder" constants (OlDefaultFolders enum)
DEFAULT_FOLDERS = {
    "inbox": 6,
    "sent mail": 5,
    "sent items": 5,
    "drafts": 16,
    "deleted items": 3,
    "junk email": 23,
    "outbox": 4,
}

OL_MAIL_ITEM = 43  # MailItem.Class value
PR_ATTACH_HIDDEN = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"


@dataclass
class Stats:
    emails_scanned: int = 0
    emails_matched: int = 0
    attachments_saved: int = 0
    inline_skipped: int = 0
    extension_skipped: int = 0
    errors: int = 0


# --------------------------------------------------------------------------
# Outlook folder resolution
# --------------------------------------------------------------------------

def resolve_folder(namespace, path):
    """Resolve a path like 'Inbox', 'Inbox/Invoices', or 'Some Account/Inbox'
    to an Outlook Folder COM object."""
    parts = [p for p in re.split(r"[\\/]+", path.strip()) if p]
    if not parts:
        raise ValueError("empty folder path")

    first_lower = parts[0].lower()
    if first_lower in DEFAULT_FOLDERS:
        folder = namespace.GetDefaultFolder(DEFAULT_FOLDERS[first_lower])
        remaining = parts[1:]
    else:
        folder = _find_top_level(namespace, parts[0])
        if folder is None:
            raise ValueError(
                f'no top-level folder/account matches "{parts[0]}". '
                f"Run with --list-folders to see available names."
            )
        remaining = parts[1:]

    for name in remaining:
        match = None
        for sub in folder.Folders:
            if sub.Name.lower() == name.lower():
                match = sub
                break
        if match is None:
            raise ValueError(f'no subfolder "{name}" under "{folder.Name}"')
        folder = match

    return folder


def _find_top_level(namespace, name):
    """Look for a folder named `name` either as a mail account root, or one
    level under any account (covers e.g. a custom top-level folder)."""
    name_lower = name.lower()
    for store_folder in namespace.Folders:
        if store_folder.Name.lower() == name_lower:
            return store_folder
        for sub in store_folder.Folders:
            if sub.Name.lower() == name_lower:
                return sub
    return None


def iter_subfolders(folder):
    result = []
    for sub in folder.Folders:
        result.append(sub)
        result.extend(iter_subfolders(sub))
    return result


def print_folder_tree(namespace, max_depth=4):
    for store_folder in namespace.Folders:
        print(store_folder.Name)
        _print_tree(store_folder, depth=1, max_depth=max_depth)


def _print_tree(folder, depth, max_depth):
    if depth > max_depth:
        return
    for sub in folder.Folders:
        print("  " * depth + sub.Name)
        _print_tree(sub, depth + 1, max_depth)


# --------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------

def parse_date(s, end_of_day=False):
    dt = datetime.datetime.strptime(s, "%Y-%m-%d")
    if end_of_day:
        dt = dt.replace(hour=23, minute=59, second=59)
    return dt


def to_naive(dt):
    """Outlook can hand back timezone-aware or naive datetimes depending on
    environment; normalize to naive so comparisons never blow up."""
    if dt is None:
        return None
    if getattr(dt, "tzinfo", None) is not None:
        return dt.replace(tzinfo=None)
    return dt


def sanitize(text, max_len=80):
    text = re.sub(r'[\\/:*?"<>|\r\n\t]+', "_", text or "")
    text = re.sub(r"_+", "_", text).strip("_. ")
    return text[:max_len] or "untitled"


def unique_path(path):
    if not path.exists():
        return path
    stem, suffix = path.stem, path.suffix
    i = 1
    while True:
        candidate = path.with_name(f"{stem} ({i}){suffix}")
        if not candidate.exists():
            return candidate
        i += 1


def is_hidden_attachment(att):
    """Inline images (e.g. signature logos) show up as attachments too.
    Outlook marks them hidden via this MAPI property."""
    try:
        return bool(att.PropertyAccessor.GetProperty(PR_ATTACH_HIDDEN))
    except Exception:
        return False


def email_target_dir(item, output_root, organize):
    received = to_naive(getattr(item, "ReceivedTime", None))
    if organize == "flat":
        return output_root
    if organize == "by-date":
        date_str = received.strftime("%Y-%m-%d") if received else "unknown-date"
        return output_root / date_str
    # by-email (default)
    date_str = received.strftime("%Y%m%d_%H%M") if received else "unknown"
    name = sanitize(f"{date_str}_{item.SenderName}_{item.Subject}")
    return output_root / name


# --------------------------------------------------------------------------
# Core processing
# --------------------------------------------------------------------------

def save_attachments(item, attachments, output_root, args, extensions, stats):
    saved_any = False
    target_dir = email_target_dir(item, output_root, args.organize)

    for att in attachments:
        if not args.include_inline and is_hidden_attachment(att):
            stats.inline_skipped += 1
            continue

        filename = att.FileName or f"attachment_{att.Index}"
        ext = Path(filename).suffix.lower().lstrip(".")
        if extensions and ext not in extensions:
            stats.extension_skipped += 1
            continue

        safe_stem = sanitize(Path(filename).stem)
        safe_name = f"{safe_stem}.{ext}" if ext else safe_stem
        dest = target_dir / safe_name

        if args.dry_run:
            print(f"  [dry-run] would save: {dest}")
            stats.attachments_saved += 1
            saved_any = True
            continue

        target_dir.mkdir(parents=True, exist_ok=True)
        dest = unique_path(dest)
        try:
            att.SaveAsFile(str(dest))
            print(f"  saved: {dest}")
            stats.attachments_saved += 1
            saved_any = True
        except Exception as e:
            print(f'  FAILED to save "{filename}" from "{item.Subject}": {e}', file=sys.stderr)
            stats.errors += 1

    return saved_any


def process_folder(folder, output_root, args, since_dt, until_dt, extensions, stats):
    items = folder.Items
    try:
        items.Sort("[ReceivedTime]", True)  # newest first
    except Exception:
        pass

    print(f'Scanning "{folder.Name}" ({items.Count} items)...')

    # Snapshot into a plain list: mutating-while-iterating a live Outlook
    # collection (e.g. via --mark-read) can otherwise skip items.
    for item in list(items):
        if args.max_emails and stats.emails_matched >= args.max_emails:
            return
        stats.emails_scanned += 1

        if getattr(item, "Class", None) != OL_MAIL_ITEM:
            continue
        if args.unread_only and not item.UnRead:
            continue

        received = to_naive(getattr(item, "ReceivedTime", None))
        if since_dt and received and received < since_dt:
            continue
        if until_dt and received and received > until_dt:
            continue

        if args.sender:
            sender_text = f"{item.SenderName} {item.SenderEmailAddress}".lower()
            if args.sender.lower() not in sender_text:
                continue

        if args.subject_contains and args.subject_contains.lower() not in (item.Subject or "").lower():
            continue

        attachments = list(item.Attachments)
        if not attachments:
            continue

        saved_any = save_attachments(item, attachments, output_root, args, extensions, stats)
        if saved_any:
            stats.emails_matched += 1
            if args.mark_read and not args.dry_run:
                try:
                    item.UnRead = False
                except Exception:
                    pass


def print_summary(stats, dry_run):
    print()
    print("=" * 50)
    print("DRY RUN SUMMARY" if dry_run else "SUMMARY")
    print(f"Emails scanned:                {stats.emails_scanned}")
    print(f"Emails with saved attachments: {stats.emails_matched}")
    print(f"Attachments saved:             {stats.attachments_saved}")
    print(f"Inline/hidden skipped:         {stats.inline_skipped}")
    print(f"Extension-filtered:            {stats.extension_skipped}")
    print(f"Errors:                        {stats.errors}")


# --------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------

def build_parser():
    parser = argparse.ArgumentParser(
        description="Download attachments from emails in a local Outlook profile.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument(
        "--folder", "-f", action="append", default=None,
        help='Outlook folder path, e.g. "Inbox" or "Inbox/Invoices". '
             "Repeatable. Defaults to Inbox.",
    )
    parser.add_argument("--output", "-o", default="attachments", help="Output directory (default: ./attachments)")
    parser.add_argument("--since", default=None, help="YYYY-MM-DD, only emails received on/after this date")
    parser.add_argument("--until", default=None, help="YYYY-MM-DD, only emails received on/before this date")
    parser.add_argument("--sender", default=None, help="Substring match against sender name/email")
    parser.add_argument("--subject-contains", dest="subject_contains", default=None)
    parser.add_argument("--unread-only", action="store_true")
    parser.add_argument("--extensions", default=None, help="Comma separated, e.g. pdf,docx,xlsx")
    parser.add_argument("--max-emails", type=int, default=None, dest="max_emails")
    parser.add_argument(
        "--include-inline", action="store_true", dest="include_inline",
        help="Also save hidden/inline attachments (e.g. signature images). Skipped by default.",
    )
    parser.add_argument("--organize", choices=["flat", "by-email", "by-date"], default="by-email")
    parser.add_argument("--recursive", action="store_true", help="Also scan subfolders of each --folder")
    parser.add_argument("--dry-run", action="store_true", dest="dry_run")
    parser.add_argument("--mark-read", action="store_true", dest="mark_read", help="Mark matched emails as read")
    parser.add_argument(
        "--list-folders", action="store_true", dest="list_folders",
        help="Print the full folder tree for the signed-in profile(s) and exit",
    )
    return parser


def main():
    args = build_parser().parse_args()

    if sys.platform != "win32":
        print(
            "This tool requires Windows with Outlook desktop installed (it uses COM automation).",
            file=sys.stderr,
        )
        sys.exit(1)

    try:
        import win32com.client
    except ImportError:
        print("pywin32 is not installed. Run: pip install pywin32", file=sys.stderr)
        sys.exit(1)

    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")

    if args.list_folders:
        print_folder_tree(namespace)
        return

    since_dt = parse_date(args.since) if args.since else None
    until_dt = parse_date(args.until, end_of_day=True) if args.until else None
    extensions = (
        {e.strip().lower().lstrip(".") for e in args.extensions.split(",")}
        if args.extensions else None
    )

    folders_to_scan = []
    for path in (args.folder or ["Inbox"]):
        try:
            folder = resolve_folder(namespace, path)
        except ValueError as e:
            print(f'Skipping folder "{path}": {e}', file=sys.stderr)
            continue
        folders_to_scan.append(folder)
        if args.recursive:
            folders_to_scan.extend(iter_subfolders(folder))

    if not folders_to_scan:
        print("No valid folders to scan. Try --list-folders to see what's available.", file=sys.stderr)
        sys.exit(1)

    output_root = Path(args.output)
    if not args.dry_run:
        output_root.mkdir(parents=True, exist_ok=True)

    stats = Stats()
    for folder in folders_to_scan:
        if args.max_emails and stats.emails_matched >= args.max_emails:
            break
        process_folder(folder, output_root, args, since_dt, until_dt, extensions, stats)

    print_summary(stats, dry_run=args.dry_run)


if __name__ == "__main__":
    main()
