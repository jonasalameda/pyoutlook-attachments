#!/usr/bin/env python3
"""
outlook_attachment_downloader_gui.py

A GUI front-end for outlook_attachment_downloader.py. Same Outlook COM
automation underneath (no Azure, no OAuth, no network calls) -- this file
just adds a window around it. It imports and reuses the core logic
(folder traversal, filtering, saving, stats) from the CLI module instead
of duplicating it, so a bug fixed in one place is fixed in both.

Dependencies: NONE beyond what the CLI script already needs (pywin32).
The GUI itself is built entirely with tkinter/ttk, which ship with the
Python standard library -- the bindings are PSF-licensed and the
underlying Tcl/Tk is BSD-style licensed, both fully permissive for
commercial use. No third-party GUI toolkit was added, since tkinter
already covers everything this needed (tabs, a folder tree, the
output-folder picker dialog) with zero extra weight or license surface.

Requirements:
    Windows + Microsoft Outlook desktop, signed in
    pip install pywin32
    outlook_attachment_downloader.py in the same folder (imported, not copied)

Run:
    python outlook_attachment_downloader_gui.py
"""

import queue
import sys
import threading
from pathlib import Path
from types import SimpleNamespace

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from outlook_attachment_downloader import (
    Stats,
    iter_subfolders,
    parse_date,
    process_folder,
    print_summary,
    save_attachments,
    to_naive,
    OL_MAIL_ITEM,
)

DUMMY_TAG = "dummy-loading-node"


class QueueWriter:
    """A file-like object that pushes writes onto a thread-safe queue
    instead of stdout, so print() calls from the background worker thread
    can be drained into the log widget on the main thread."""

    def __init__(self, q):
        self.q = q

    def write(self, text):
        if text:
            self.q.put(("text", text))

    def flush(self):
        pass


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Outlook Attachment Downloader")
        self.root.geometry("780x680")
        self.root.minsize(640, 540)

        self._namespace = None  # connected lazily, on first use
        self.folder_map = {}    # tree item id -> Outlook Folder COM object
        self.log_queue = queue.Queue()
        self.is_running = False

        # Move target stored as (EntryID, StoreID, Name) when chosen
        self.move_target = None

        self._build_widgets()
        self.root.after(100, self._poll_log_queue)

    # ------------------------------------------------------------------
    # Outlook connection (lazy, cached)
    # ------------------------------------------------------------------

    def get_namespace(self):
        if self._namespace is not None:
            return self._namespace

        if sys.platform != "win32":
            messagebox.showerror(
                "Unsupported platform",
                "This requires Windows with Outlook desktop installed (it uses COM automation).",
            )
            return None

        try:
            import win32com.client
        except ImportError:
            messagebox.showerror(
                "Missing dependency",
                "pywin32 is not installed.\n\nRun: pip install pywin32",
            )
            return None

        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            self._namespace = outlook.GetNamespace("MAPI")
        except Exception as e:
            messagebox.showerror("Outlook connection failed", str(e))
            return None

        return self._namespace

    # ------------------------------------------------------------------
    # Widget construction
    # ------------------------------------------------------------------

    def _build_widgets(self):
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=3)
        self.root.rowconfigure(2, weight=2)

        notebook = ttk.Notebook(self.root)
        notebook.grid(row=0, column=0, sticky="nsew", padx=8, pady=(8, 4))

        filters_tab = ttk.Frame(notebook, padding=10)
        folder_tab = ttk.Frame(notebook, padding=10)
        notebook.add(filters_tab, text="Filters && Options")
        notebook.add(folder_tab, text="Outlook Folder")

        self._build_filters_tab(filters_tab)
        self._build_folder_tab(folder_tab)
        self._build_bottom_bar()
        self._build_log_panel()

    def _build_filters_tab(self, parent):
        # We'll put all filter controls inside a collapsible frame so the
        # user can hide/show them. A single toggle button controls the
        # visibility.
        parent.columnconfigure(0, weight=1)

        header = ttk.Frame(parent)
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(1, weight=1)
        self.filters_visible = tk.BooleanVar(value=True)
        self._filters_toggle_btn = ttk.Button(header, text="▼ Filters", width=12, command=self._toggle_filters)
        self._filters_toggle_btn.grid(row=0, column=0, sticky="w")
        ttk.Label(header, text="Show or hide filter options for a cleaner window.", foreground="#555").grid(row=0, column=1, sticky="w", padx=(8,0))

        # Content frame contains all the existing widgets. We'll toggle its
        # visibility by grid_remove()/grid().
        self.filters_content = ttk.Frame(parent)
        self.filters_content.grid(row=1, column=0, sticky="nsew", pady=(8,0))
        for col in (1, 3):
            self.filters_content.columnconfigure(col, weight=1)

        row = 0
        ttk.Label(self.filters_content, text="Sender contains:").grid(row=row, column=0, sticky="w", pady=4)
        self.sender_var = tk.StringVar()
        ttk.Entry(self.filters_content, textvariable=self.sender_var).grid(row=row, column=1, columnspan=3, sticky="ew", padx=(6, 0))

        row += 1
        ttk.Label(self.filters_content, text="Subject contains:").grid(row=row, column=0, sticky="w", pady=4)
        self.subject_var = tk.StringVar()
        ttk.Entry(self.filters_content, textvariable=self.subject_var).grid(row=row, column=1, columnspan=3, sticky="ew", padx=(6, 0))

        row += 1
        ttk.Label(self.filters_content, text="Since (YYYY-MM-DD):").grid(row=row, column=0, sticky="w", pady=4)
        self.since_var = tk.StringVar()
        ttk.Entry(self.filters_content, textvariable=self.since_var, width=16).grid(row=row, column=1, sticky="w", padx=(6, 0))

        ttk.Label(self.filters_content, text="Until (YYYY-MM-DD):").grid(row=row, column=2, sticky="w", padx=(12, 0))
        self.until_var = tk.StringVar()
        ttk.Entry(self.filters_content, textvariable=self.until_var, width=16).grid(row=row, column=3, sticky="w", padx=(6, 0))

        row += 1
        ttk.Label(self.filters_content, text="Extensions (comma separated):").grid(row=row, column=0, sticky="w", pady=4)
        self.extensions_var = tk.StringVar()
        ttk.Entry(self.filters_content, textvariable=self.extensions_var, width=24).grid(row=row, column=1, sticky="w", padx=(6, 0))

        ttk.Label(self.filters_content, text="Max emails (blank = no limit):").grid(row=row, column=2, sticky="w", padx=(12, 0))
        self.max_emails_var = tk.StringVar()
        ttk.Entry(self.filters_content, textvariable=self.max_emails_var, width=10).grid(row=row, column=3, sticky="w", padx=(6, 0))

        row += 1
        ttk.Separator(self.filters_content, orient="horizontal").grid(row=row, column=0, columnspan=4, sticky="ew", pady=10)

        row += 1
        self.unread_only_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(self.filters_content, text="Unread only", variable=self.unread_only_var).grid(row=row, column=0, sticky="w", pady=2)

        self.include_inline_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            self.filters_content, text="Include inline/hidden attachments (e.g. signature images)",
            variable=self.include_inline_var,
        ).grid(row=row, column=1, columnspan=3, sticky="w", pady=2)

        row += 1
        self.mark_read_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            self.filters_content, text="Mark matched emails as read after saving", variable=self.mark_read_var,
        ).grid(row=row, column=0, columnspan=2, sticky="w", pady=2)

        self.dry_run_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            self.filters_content, text="Dry run (preview only, save nothing)", variable=self.dry_run_var,
        ).grid(row=row, column=2, columnspan=2, sticky="w", pady=2)

        row += 1
        ttk.Separator(self.filters_content, orient="horizontal").grid(row=row, column=0, columnspan=4, sticky="ew", pady=10)

        row += 1
        ttk.Label(self.filters_content, text="Organize saved attachments:").grid(row=row, column=0, sticky="w", pady=2)
        self.organize_var = tk.StringVar(value="by-email")
        organize_frame = ttk.Frame(self.filters_content)
        organize_frame.grid(row=row, column=1, columnspan=3, sticky="w")
        ttk.Radiobutton(organize_frame, text="One folder per email", value="by-email", variable=self.organize_var).pack(side="left", padx=(0, 12))
        ttk.Radiobutton(organize_frame, text="One folder per date", value="by-date", variable=self.organize_var).pack(side="left", padx=(0, 12))
        ttk.Radiobutton(organize_frame, text="All in one flat folder", value="flat", variable=self.organize_var).pack(side="left")

        # New: move-processed-emails option with a "Choose..." button that
        # asks the user which Outlook folder to move matching emails into.
        row += 1
        ttk.Separator(self.filters_content, orient="horizontal").grid(row=row, column=0, columnspan=4, sticky="ew", pady=10)

        row += 1
        self.move_emails_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(self.filters_content, text="Move processed emails to another Outlook folder", variable=self.move_emails_var).grid(row=row, column=0, columnspan=2, sticky="w", pady=2)
        self.move_folder_label_var = tk.StringVar(value="(no folder chosen)")
        ttk.Label(self.filters_content, textvariable=self.move_folder_label_var, foreground="#555").grid(row=row, column=2, sticky="w")
        ttk.Button(self.filters_content, text="Choose...", command=self.on_choose_move_folder_clicked).grid(row=row, column=3, sticky="e")

    def _toggle_filters(self):
        if self.filters_visible.get():
            # hide
            self.filters_content.grid_remove()
            self.filters_visible.set(False)
            self._filters_toggle_btn.configure(text="▶ Filters")
        else:
            # show
            self.filters_content.grid()
            self.filters_visible.set(True)
            self._filters_toggle_btn.configure(text="▼ Filters")

    def _build_folder_tab(self, parent):
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)

        top = ttk.Frame(parent)
        top.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        ttk.Button(top, text="Load Folders", command=self.on_load_folders_clicked).pack(side="left")
        ttk.Label(
            top, text="  Click a folder's arrow to load its subfolders. Ctrl/Shift-click to select more than one.",
            foreground="#555",
        ).pack(side="left", padx=(8, 0))

        tree_frame = ttk.Frame(parent)
        tree_frame.grid(row=1, column=0, sticky="nsew")
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)

        self.tree = ttk.Treeview(tree_frame, selectmode="extended", show="tree")
        self.tree.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.bind("<<TreeviewOpen>>", self.on_tree_open)
        self.tree.bind("<<TreeviewSelect>>", self.on_tree_select)

        bottom = ttk.Frame(parent)
        bottom.grid(row=2, column=0, sticky="ew", pady=(6, 0))
        self.recursive_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(bottom, text="Include subfolders (recursive)", variable=self.recursive_var).pack(side="left")

        self.selection_label_var = tk.StringVar(value="No folder selected — Inbox will be used by default")
        ttk.Label(bottom, textvariable=self.selection_label_var, foreground="#555").pack(side="left", padx=(16, 0))

    def _build_bottom_bar(self):
        bar = ttk.Frame(self.root, padding=(8, 4))
        bar.grid(row=1, column=0, sticky="ew")
        bar.columnconfigure(1, weight=1)

        ttk.Label(bar, text="Output folder:").grid(row=0, column=0, sticky="w")
        self.output_var = tk.StringVar(value=str(Path("attachments").resolve()))
        ttk.Entry(bar, textvariable=self.output_var, state="readonly").grid(row=0, column=1, sticky="ew", padx=6)
        ttk.Button(bar, text="Change...", command=self.on_change_output_clicked).grid(row=0, column=2)

        self.download_button = ttk.Button(bar, text="Download Attachments", command=self.on_download_clicked)
        self.download_button.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(8, 0))

    def _build_log_panel(self):
        frame = ttk.LabelFrame(self.root, text="Activity Log", padding=6)
        frame.grid(row=2, column=0, sticky="nsew", padx=8, pady=(4, 8))
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)

        self.log_text = tk.Text(frame, height=10, wrap="word", state="disabled")
        self.log_text.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=self.log_text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=scrollbar.set)

    # ------------------------------------------------------------------
    # Folder tree (lazy-loaded: only fetches children when a node is
    # expanded, instead of walking the whole mailbox up front)
    # ------------------------------------------------------------------

    def on_load_folders_clicked(self):
        namespace = self.get_namespace()
        if namespace is None:
            return
        self.tree.delete(*self.tree.get_children())
        self.folder_map.clear()
        try:
            for store_folder in namespace.Folders:
                self._insert_folder_node("", store_folder)
        except Exception as e:
            messagebox.showerror("Could not load folders", str(e))

    def _insert_folder_node(self, parent_iid, folder_obj):
        iid = self.tree.insert(parent_iid, "end", text=folder_obj.Name)
        self.folder_map[iid] = folder_obj
        self._add_placeholder_if_needed(iid, folder_obj)
        return iid

    def _add_placeholder_if_needed(self, iid, folder_obj):
        try:
            has_children = folder_obj.Folders.Count > 0
        except Exception:
            has_children = False
        if has_children:
            self.tree.insert(iid, "end", text="Loading...", tags=(DUMMY_TAG,))

    def on_tree_open(self, event):
        iid = self.tree.focus()
        children = self.tree.get_children(iid)
        if len(children) != 1 or DUMMY_TAG not in self.tree.item(children[0], "tags"):
            return  # already loaded, or has no children

        self.tree.delete(children[0])
        folder_obj = self.folder_map.get(iid)
        if folder_obj is None:
            return
        try:
            subfolders = list(folder_obj.Folders)
        except Exception as e:
            self.tree.insert(iid, "end", text=f"(could not load: {e})")
            return
        for sub in subfolders:
            self._insert_folder_node(iid, sub)

    def on_tree_select(self, event):
        selected = self.tree.selection()
        if not selected:
            text = "No folder selected — Inbox will be used by default"
        elif len(selected) == 1:
            text = f"Selected: {self.tree.item(selected[0], 'text')}"
        else:
            text = f"{len(selected)} folders selected"
        self.selection_label_var.set(text)

    def get_selected_folders(self, namespace):
        # Return lightweight descriptors (EntryID, StoreID, Name) instead of COM objects.
        selected = [self.folder_map[iid] for iid in self.tree.selection() if iid in self.folder_map]
        if not selected:
            selected = [namespace.GetDefaultFolder(6)]  # Inbox

        if self.recursive_var.get():
            expanded = []
            for f in selected:
                expanded.append(f)
                expanded.extend(iter_subfolders(f))
            selected = expanded

        descriptors = []
        for f in selected:
            # EntryID and StoreID are stable identifiers we can use to re-open the folder in another thread
            entry_id = getattr(f, "EntryID", None)
            store_id = getattr(f, "StoreID", None)
            name = getattr(f, "Name", "?")
            descriptors.append((entry_id, store_id, name))
        return descriptors

    # ------------------------------------------------------------------
    # Output folder
    # ------------------------------------------------------------------

    def on_change_output_clicked(self):
        current = self.output_var.get()
        initial = current if Path(current).exists() else str(Path.home())
        chosen = filedialog.askdirectory(initialdir=initial, title="Choose output folder")
        if chosen:
            self.output_var.set(chosen)

    # ------------------------------------------------------------------
    # Move target selection
    # ------------------------------------------------------------------

    def on_choose_move_folder_clicked(self):
        namespace = self.get_namespace()
        if namespace is None:
            return
        try:
            # Use Outlook's folder picker UI to let the user choose a target folder.
            chosen = namespace.PickFolder()
            if chosen is None:
                return
            entry_id = getattr(chosen, "EntryID", None)
            store_id = getattr(chosen, "StoreID", None)
            name = getattr(chosen, "Name", "?")
            self.move_target = (entry_id, store_id, name)
            # Display a readable label for the chosen folder
            self.move_folder_label_var.set(f"Chosen: {name}")
        except Exception as e:
            messagebox.showerror("Could not pick folder", str(e))

    # ------------------------------------------------------------------
    # Download
    # ------------------------------------------------------------------

    def _parse_max_emails(self):
        raw = self.max_emails_var.get().strip()
        if not raw:
            return None
        try:
            value = int(raw)
            if value <= 0:
                raise ValueError
            return value
        except ValueError:
            raise ValueError("Max emails must be a positive whole number, or left blank.")

    def build_args(self, max_emails):
        ns = SimpleNamespace(
            sender=self.sender_var.get().strip() or None,
            subject_contains=self.subject_var.get().strip() or None,
            extensions=self.extensions_var.get().strip() or None,
            max_emails=max_emails,
            unread_only=self.unread_only_var.get(),
            include_inline=self.include_inline_var.get(),
            organize=self.organize_var.get(),
            dry_run=self.dry_run_var.get(),
            no_mark_read=not self.mark_read_var.get(),
        )
        # Include move options on the namespace so the worker can see them.
        ns.move_emails = self.move_emails_var.get()
        if self.move_target:
            ns.move_entry_id, ns.move_store_id, ns.move_folder_name = self.move_target
        else:
            ns.move_entry_id = ns.move_store_id = ns.move_folder_name = None
        return ns

    def on_download_clicked(self):
        if self.is_running:
            return

        namespace = self.get_namespace()
        if namespace is None:
            return

        try:
            since_dt = parse_date(self.since_var.get().strip()) if self.since_var.get().strip() else None
            until_dt = parse_date(self.until_var.get().strip(), end_of_day=True) if self.until_var.get().strip() else None
            max_emails = self._parse_max_emails()
        except ValueError as e:
            messagebox.showerror("Invalid input", str(e))
            return

        folders = self.get_selected_folders(namespace)

        extensions_raw = self.extensions_var.get().strip()
        extensions = {e.strip().lower().lstrip(".") for e in extensions_raw.split(",")} if extensions_raw else None

        args = self.build_args(max_emails)

        output_root = Path(self.output_var.get()).resolve()
        if not args.dry_run:
            try:
                output_root.mkdir(parents=True, exist_ok=True)
            except OSError as e:
                messagebox.showerror("Output folder error", str(e))
                return

        self._set_running(True)
        self._log_clear()
        thread = threading.Thread(
            target=self._download_worker,
            args=(folders, output_root, args, since_dt, until_dt, extensions),
            daemon=True,
        )
        thread.start()

    def _download_worker(self, folders, output_root, args, since_dt, until_dt, extensions):
        # folders is a list of (entry_id, store_id, name) tuples.
        try:
            import pythoncom
            import win32com.client
        except Exception as e:
            # pywin32 not available in this thread or import failed
            self.log_queue.put(("text", f"Worker thread import error: {e}\n"))
            self.log_queue.put(("done", None))
            return

        # Initialize COM on this thread
        pythoncom.CoInitialize()
        try:
            old_out, old_err = sys.stdout, sys.stderr
            sys.stdout = sys.stderr = QueueWriter(self.log_queue)
            try:
                # Create a separate Outlook Application/Namespace in this thread
                outlook = win32com.client.Dispatch("Outlook.Application")
                namespace = outlook.GetNamespace("MAPI")

                stats = Stats()
                # If the user asked to move emails, resolve the destination folder
                move_folder_obj = None
                if args.move_emails and args.move_entry_id:
                    try:
                        if args.move_store_id:
                            move_folder_obj = namespace.GetFolderFromID(args.move_entry_id, args.move_store_id)
                        else:
                            move_folder_obj = namespace.GetFolderFromID(args.move_entry_id)
                    except Exception as e:
                        print(f"Could not resolve move-to folder: {e}")
                        move_folder_obj = None

                # We'll re-implement the core per-folder loop here so we can
                # move individual MailItems after successful saves. This keeps
                # the heavy lifting (saving attachments) in the shared
                # save_attachments() function while avoiding modifying the
                # core module.
                for entry_id, store_id, name in folders:
                    if args.max_emails and stats.emails_matched >= args.max_emails:
                        break
                    try:
                        # Re-resolve the folder in this thread
                        if store_id:
                            folder = namespace.GetFolderFromID(entry_id, store_id)
                        else:
                            folder = namespace.GetFolderFromID(entry_id)
                    except Exception as e:
                        print(f'Error opening folder "{name}": {e}')
                        continue

                    try:
                        items = folder.Items
                        try:
                            items.Sort("[ReceivedTime]", True)
                        except Exception:
                            pass
                    except Exception as e:
                        print(f'Error accessing items in "{name}": {e}')
                        continue

                    print(f'Scanning "{name}" ({getattr(items, "Count", "?")} items)...')

                    for item in list(items):
                        if args.max_emails and stats.emails_matched >= args.max_emails:
                            break
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
                            if not args.dry_run and not args.no_mark_read:
                                try:
                                    item.UnRead = False
                                except Exception:
                                    pass
                            # If requested and we resolved a move destination,
                            # move the message to that folder.
                            if not args.dry_run and args.move_emails and move_folder_obj is not None:
                                try:
                                    item.Move(move_folder_obj)
                                    print(f'  moved message "{getattr(item, "Subject", "")}" to "{getattr(move_folder_obj, "Name", "?")}"')
                                except Exception as e:
                                    print(f'  FAILED to move message "{getattr(item, "Subject", "")}": {e}', file=sys.stderr)
                                    stats.errors += 1

                print_summary(stats, dry_run=args.dry_run)
            finally:
                sys.stdout, sys.stderr = old_out, old_err
        except Exception as e:
            # Any unexpected exception — print so it lands in the log widget
            self.log_queue.put(("text", f"Unexpected error: {e}\n"))
        finally:
            pythoncom.CoUninitialize()
            self.log_queue.put(("done", None))

    def _set_running(self, running):
        self.is_running = running
        self.download_button.configure(
            state="disabled" if running else "normal",
            text="Working..." if running else "Download Attachments",
        )

    def _log_clear(self):
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    def _poll_log_queue(self):
        try:
            while True:
                kind, payload = self.log_queue.get_nowait()
                if kind == "done":
                    self._set_running(False)
                else:
                    self.log_text.configure(state="normal")
                    self.log_text.insert("end", payload)
                    self.log_text.see("end")
                    self.log_text.configure(state="disabled")
        except queue.Empty:
            pass
        self.root.after(100, self._poll_log_queue)


def main():
    root = tk.Tk()
    if sys.platform != "win32":
        messagebox.showwarning(
            "Unsupported platform",
            "This requires Windows with Outlook desktop installed.\n\n"
            "The window will open, but any Outlook action will show an error.",
        )
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
