import os
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import List, Dict, Any, Optional

try:
    from PIL import Image, ImageTk
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

try:
    import customtkinter as ctk
    ctk.set_appearance_mode("Dark")
    ctk.set_default_color_theme("blue")
    HAS_CTK = True
except ImportError:
    HAS_CTK = False

from src.docx_parser import DocxParser
from src.file_matcher import FileMatcher
from src.file_copier import FileCopier


class ImageMatcherApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Word Image Matcher & Exporter")
        self.root.geometry("1100x750")
        self.root.minsize(900, 600)

        # Variables
        self.docx_path_var = tk.StringVar()
        self.source_dir_var = tk.StringVar()
        self.dest_dir_var = tk.StringVar()

        self.recursive_var = tk.BooleanVar(value=True)
        self.case_sensitive_var = tk.BooleanVar(value=False)
        self.extension_agnostic_var = tk.BooleanVar(value=True)
        self.image_only_var = tk.BooleanVar(value=True)
        self.overwrite_var = tk.BooleanVar(value=False)
        self.move_files_var = tk.BooleanVar(value=False)
        self.split_tokens_var = tk.BooleanVar(value=True)

        self.matched_items: List[Dict[str, Any]] = []
        self.missing_items: List[Dict[str, Any]] = []
        self.thumbnail_cache = {}

        self._build_ui()

    def _build_ui(self):
        # Apply dark mode style for Tkinter if CustomTkinter is missing
        style = ttk.Style()
        style.theme_use('clam')

        # Main Layout Frame
        main_container = ttk.Frame(self.root, padding="12")
        main_container.pack(fill=tk.BOTH, expand=True)

        # Header Title
        header_frame = ttk.Frame(main_container)
        header_frame.pack(fill=tk.X, pady=(0, 10))

        title_label = ttk.Label(
            header_frame,
            text="Word Document Image Matcher & Exporter",
            font=("Segoe UI", 16, "bold")
        )
        title_label.pack(anchor=tk.W)

        subtitle_label = ttk.Label(
            header_frame,
            text="Extract image names or IDs from Word file (.docx), find matching files in folder, and copy to destination.",
            font=("Segoe UI", 9)
        )
        subtitle_label.pack(anchor=tk.W)

        # File Selection Inputs Frame
        inputs_frame = ttk.LabelFrame(main_container, text=" File & Folder Inputs ", padding="10")
        inputs_frame.pack(fill=tk.X, pady=(0, 10))

        # 1. Word Document Path
        f1 = ttk.Frame(inputs_frame)
        f1.pack(fill=tk.X, pady=3)
        ttk.Label(f1, text="Word File (.docx):", width=18, anchor=tk.W).pack(side=tk.LEFT)
        ttk.Entry(f1, textvariable=self.docx_path_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Button(f1, text="Browse...", command=self.browse_docx).pack(side=tk.RIGHT)

        # 2. Source Directory Path
        f2 = ttk.Frame(inputs_frame)
        f2.pack(fill=tk.X, pady=3)
        ttk.Label(f2, text="Source Images Folder:", width=18, anchor=tk.W).pack(side=tk.LEFT)
        ttk.Entry(f2, textvariable=self.source_dir_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Button(f2, text="Browse...", command=self.browse_source).pack(side=tk.RIGHT)

        # 3. Output Directory Path
        f3 = ttk.Frame(inputs_frame)
        f3.pack(fill=tk.X, pady=3)
        ttk.Label(f3, text="Destination Folder:", width=18, anchor=tk.W).pack(side=tk.LEFT)
        ttk.Entry(f3, textvariable=self.dest_dir_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Button(f3, text="Browse...", command=self.browse_dest).pack(side=tk.RIGHT)

        # Options Row Frame
        options_frame = ttk.LabelFrame(main_container, text=" Matching & Copy Options ", padding="8")
        options_frame.pack(fill=tk.X, pady=(0, 10))

        opt_grid = ttk.Frame(options_frame)
        opt_grid.pack(fill=tk.X)

        ttk.Checkbutton(opt_grid, text="Search Subfolders Recursively", variable=self.recursive_var).grid(row=0, column=0, sticky=tk.W, padx=10, pady=2)
        ttk.Checkbutton(opt_grid, text="Case-Insensitive Search", variable=self.case_sensitive_var).grid(row=0, column=1, sticky=tk.W, padx=10, pady=2)
        ttk.Checkbutton(opt_grid, text="Extension-Agnostic Match (e.g. ID -> ID.png)", variable=self.extension_agnostic_var).grid(row=0, column=2, sticky=tk.W, padx=10, pady=2)

        ttk.Checkbutton(opt_grid, text="Image Files Only", variable=self.image_only_var).grid(row=1, column=0, sticky=tk.W, padx=10, pady=2)
        ttk.Checkbutton(opt_grid, text="Overwrite Existing Destination Files", variable=self.overwrite_var).grid(row=1, column=1, sticky=tk.W, padx=10, pady=2)
        ttk.Checkbutton(opt_grid, text="Move Files (instead of Copying)", variable=self.move_files_var).grid(row=1, column=2, sticky=tk.W, padx=10, pady=2)

        ttk.Checkbutton(opt_grid, text="Split Comma-Separated IDs", variable=self.split_tokens_var).grid(row=2, column=0, sticky=tk.W, padx=10, pady=2)

        # Action Buttons Row
        actions_frame = ttk.Frame(main_container)
        actions_frame.pack(fill=tk.X, pady=(0, 10))

        self.btn_search = ttk.Button(
            actions_frame,
            text="🔍 Step 1: Scan & Search Matched Files",
            command=self.start_search_thread
        )
        self.btn_search.pack(side=tk.LEFT, padx=(0, 10))

        self.btn_copy = ttk.Button(
            actions_frame,
            text="📁 Step 2: Export Selected Files to Destination",
            command=self.start_copy_thread,
            state=tk.DISABLED
        )
        self.btn_copy.pack(side=tk.LEFT)

        # Selection Toggles
        self.btn_select_all = ttk.Button(actions_frame, text="Select All", command=lambda: self.toggle_all_selections(True), state=tk.DISABLED)
        self.btn_select_all.pack(side=tk.RIGHT, padx=2)
        self.btn_deselect_all = ttk.Button(actions_frame, text="Deselect All", command=lambda: self.toggle_all_selections(False), state=tk.DISABLED)
        self.btn_deselect_all.pack(side=tk.RIGHT, padx=2)

        # Notebook / Tabs
        self.notebook = ttk.Notebook(main_container)
        self.notebook.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # Tab 1: Matched Preview Grid
        self.tab_matched = ttk.Frame(self.notebook, padding="5")
        self.notebook.add(self.tab_matched, text=" Matched Files (0) ")

        # Scrollable Canvas for Matched Items
        self.canvas = tk.Canvas(self.tab_matched, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self.tab_matched, orient=tk.VERTICAL, command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.bind('<Configure>', self._on_canvas_configure)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Tab 2: Missing Items List
        self.tab_missing = ttk.Frame(self.notebook, padding="5")
        self.notebook.add(self.tab_missing, text=" Missing / Unmatched IDs (0) ")

        missing_ctrl = ttk.Frame(self.tab_missing)
        missing_ctrl.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(missing_ctrl, text="The following IDs/names from Word were not found in the source folder:").pack(side=tk.LEFT)
        ttk.Button(missing_ctrl, text="Copy Missing List", command=self.copy_missing_to_clipboard).pack(side=tk.RIGHT)

        self.missing_tree = ttk.Treeview(self.tab_missing, columns=("RawText", "ID", "Location"), show="headings")
        self.missing_tree.heading("RawText", text="Word Text")
        self.missing_tree.heading("ID", text="Sanitized ID")
        self.missing_tree.heading("Location", text="Word Location")
        self.missing_tree.column("RawText", width=300)
        self.missing_tree.column("ID", width=200)
        self.missing_tree.column("Location", width=250)

        missing_scroll = ttk.Scrollbar(self.tab_missing, orient=tk.VERTICAL, command=self.missing_tree.yview)
        self.missing_tree.configure(yscrollcommand=missing_scroll.set)
        self.missing_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        missing_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # Tab 3: Activity Log
        self.tab_log = ttk.Frame(self.notebook, padding="5")
        self.notebook.add(self.tab_log, text=" Activity Log ")

        self.log_text = tk.Text(self.tab_log, wrap=tk.WORD, font=("Consolas", 9))
        log_scroll = ttk.Scrollbar(self.tab_log, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scroll.set)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        log_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # Status & Progress Footer
        footer_frame = ttk.Frame(main_container)
        footer_frame.pack(fill=tk.X)

        self.status_label = ttk.Label(footer_frame, text="Ready. Select Word file and folders to start.", font=("Segoe UI", 9, "italic"))
        self.status_label.pack(side=tk.LEFT)

        self.progress_bar = ttk.Progressbar(footer_frame, mode="determinate", length=250)
        self.progress_bar.pack(side=tk.RIGHT)

        self.log("Application initialized successfully.")

    def _on_canvas_configure(self, event):
        self.canvas.itemconfig(self.canvas_window, width=event.width)

    def log(self, message: str):
        def _log():
            self.log_text.insert(tk.END, f"{message}\n")
            self.log_text.see(tk.END)
            self.status_label.config(text=message)
        self.root.after(0, _log)

    def browse_docx(self):
        filename = filedialog.askopenfilename(
            title="Select Word Document",
            filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")]
        )
        if filename:
            self.docx_path_var.set(filename)
            self.log(f"Selected Word document: {filename}")

    def browse_source(self):
        folder = filedialog.askdirectory(title="Select Source Images Folder")
        if folder:
            self.source_dir_var.set(folder)
            self.log(f"Selected Source folder: {folder}")

    def browse_dest(self):
        folder = filedialog.askdirectory(title="Select Destination Folder")
        if folder:
            self.dest_dir_var.set(folder)
            self.log(f"Selected Destination folder: {folder}")

    def start_search_thread(self):
        docx_path = self.docx_path_var.get().strip()
        source_dir = self.source_dir_var.get().strip()

        if not docx_path or not os.path.exists(docx_path):
            messagebox.showerror("Invalid Input", "Please select a valid Word document (.docx).")
            return

        if not source_dir or not os.path.exists(source_dir):
            messagebox.showerror("Invalid Input", "Please select a valid Source folder.")
            return

        self.btn_search.config(state=tk.DISABLED)
        self.btn_copy.config(state=tk.DISABLED)
        self.progress_bar.config(mode="indeterminate")
        self.progress_bar.start(10)
        self.log("Starting Word extraction and folder scanning...")

        threading.Thread(target=self._run_search, args=(docx_path, source_dir), daemon=True).start()

    def _run_search(self, docx_path: str, source_dir: str):
        try:
            # 1. Parse docx
            extracted = DocxParser.extract_items(docx_path, extract_tokens_per_line=self.split_tokens_var.get())
            self.log(f"Extracted {len(extracted)} candidate IDs/names from Word document.")

            if not extracted:
                self.log("Warning: No IDs or filenames were found in the Word document.")

            # 2. Match against source directory
            result = FileMatcher.match_docx_items(
                extracted_items=extracted,
                source_folder=source_dir,
                recursive=self.recursive_var.get(),
                case_sensitive=self.case_sensitive_var.get(),
                extension_agnostic=self.extension_agnostic_var.get(),
                image_only=self.image_only_var.get()
            )

            self.matched_items = result['matched']
            self.missing_items = result['missing']

            self.log(f"Search complete: Found {len(self.matched_items)} matched files, {len(self.missing_items)} missing IDs.")
            self.root.after(0, self._update_results_ui)

        except Exception as e:
            self.log(f"Error during search: {str(e)}")
            self.root.after(0, lambda: messagebox.showerror("Search Error", str(e)))
        finally:
            self.root.after(0, self._stop_progress)

    def _stop_progress(self):
        self.progress_bar.stop()
        self.progress_bar.config(mode="determinate", value=0)
        self.btn_search.config(state=tk.NORMAL)

    def _update_results_ui(self):
        # Update notebook tab titles
        self.notebook.tab(self.tab_matched, text=f" Matched Files ({len(self.matched_items)}) ")
        self.notebook.tab(self.tab_missing, text=f" Missing / Unmatched IDs ({len(self.missing_items)}) ")

        # Clear scrollable frame
        for child in self.scrollable_frame.winfo_children():
            child.destroy()

        self.thumbnail_cache.clear()

        # Render Matched Items
        if not self.matched_items:
            empty_lbl = ttk.Label(self.scrollable_frame, text="No matching files found.", font=("Segoe UI", 10, "italic"), padding="20")
            empty_lbl.pack()
            self.btn_copy.config(state=tk.DISABLED)
            self.btn_select_all.config(state=tk.DISABLED)
            self.btn_deselect_all.config(state=tk.DISABLED)
        else:
            self.btn_copy.config(state=tk.NORMAL)
            self.btn_select_all.config(state=tk.NORMAL)
            self.btn_deselect_all.config(state=tk.NORMAL)

            for idx, item in enumerate(self.matched_items):
                card = ttk.Frame(self.scrollable_frame, padding="6", relief="solid", borderwidth=1)
                card.pack(fill=tk.X, pady=4, padx=4)

                # Checkbox
                var = tk.BooleanVar(value=item.get('selected', True))
                item['checkbox_var'] = var
                cb = ttk.Checkbutton(card, variable=var, command=lambda i=item: self._on_item_check(i))
                cb.pack(side=tk.LEFT, padx=(5, 10))

                # Image Thumbnail Preview
                img_label = ttk.Label(card, text="[No Preview]", width=12, anchor=tk.CENTER)
                img_label.pack(side=tk.LEFT, padx=(0, 10))

                # Load thumbnail asynchronously or lazily
                self._load_thumbnail(item['file_path'], img_label)

                # Info text
                info_frame = ttk.Frame(card)
                info_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

                title = ttk.Label(info_frame, text=item['file_name'], font=("Segoe UI", 10, "bold"))
                title.pack(anchor=tk.W)

                details = f"Word Match: '{item['raw']}' | Doc Location: {item['location']} | Type: {item['match_type']}"
                sub_lbl = ttk.Label(info_frame, text=details, font=("Segoe UI", 8))
                sub_lbl.pack(anchor=tk.W)

                path_lbl = ttk.Label(info_frame, text=f"Source Path: {item['file_path']}", font=("Segoe UI", 8, "italic"))
                path_lbl.pack(anchor=tk.W)

        # Render Missing Items Treeview
        for row in self.missing_tree.get_children():
            self.missing_tree.delete(row)

        for m in self.missing_items:
            self.missing_tree.insert("", tk.END, values=(m['raw'], m['id'], m['location']))

    def _on_item_check(self, item: Dict[str, Any]):
        item['selected'] = item['checkbox_var'].get()

    def toggle_all_selections(self, selected: bool):
        for item in self.matched_items:
            item['selected'] = selected
            if 'checkbox_var' in item:
                item['checkbox_var'].set(selected)

    def _load_thumbnail(self, file_path: str, label_widget: ttk.Label):
        if not HAS_PIL:
            return

        def _loader():
            try:
                ext = os.path.splitext(file_path)[1].lower()
                if ext in ('.jpg', '.jpeg', '.png', '.bmp', '.gif', '.webp'):
                    with Image.open(file_path) as img:
                        img.thumbnail((60, 60))
                        photo = ImageTk.PhotoImage(img)
                        self.thumbnail_cache[file_path] = photo
                        self.root.after(0, lambda: label_widget.config(image=photo, text=""))
            except Exception:
                pass

        threading.Thread(target=_loader, daemon=True).start()

    def copy_missing_to_clipboard(self):
        if not self.missing_items:
            messagebox.showinfo("Missing Items", "There are no missing items.")
            return

        lines = [f"{m['raw']}\t(ID: {m['id']})\t{m['location']}" for m in self.missing_items]
        text = "\n".join(lines)
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        messagebox.showinfo("Copied", f"Copied {len(self.missing_items)} missing items to clipboard.")

    def start_copy_thread(self):
        dest_dir = self.dest_dir_var.get().strip()
        if not dest_dir:
            dest_dir = filedialog.askdirectory(title="Select Destination Folder")
            if dest_dir:
                self.dest_dir_var.set(dest_dir)
            else:
                return

        selected_files = [item for item in self.matched_items if item.get('selected', True)]
        if not selected_files:
            messagebox.showwarning("No Files Selected", "Please select at least one matched file to export.")
            return

        self.btn_search.config(state=tk.DISABLED)
        self.btn_copy.config(state=tk.DISABLED)
        self.progress_bar.config(mode="determinate", value=0)

        self.log(f"Starting export of {len(selected_files)} files to {dest_dir}...")

        threading.Thread(target=self._run_copy, args=(selected_files, dest_dir), daemon=True).start()

    def _run_copy(self, selected_files: List[Dict[str, Any]], dest_dir: str):
        def progress_cb(current: int, total: int, msg: str):
            pct = (current / total) * 100
            self.root.after(0, lambda: self.progress_bar.config(value=pct))
            self.log(f"[{current}/{total}] {msg}")

        try:
            stats = FileCopier.process_files(
                selected_items=selected_files,
                dest_folder=dest_dir,
                missing_items=self.missing_items,
                overwrite=self.overwrite_var.get(),
                move_files=self.move_files_var.get(),
                progress_callback=progress_cb
            )

            msg = f"Export completed!\n\nSuccessfully exported: {stats['copied_count']} files"
            if stats['failed_count'] > 0:
                msg += f"\nFailed: {stats['failed_count']} files"
            if stats['report_path']:
                msg += f"\n\nCSV Summary Report saved to:\n{stats['report_path']}"

            self.log(f"Export finished. {stats['copied_count']} files saved to {dest_dir}")
            self.root.after(0, lambda: messagebox.showinfo("Export Successful", msg))

        except Exception as e:
            self.log(f"Export error: {str(e)}")
            self.root.after(0, lambda: messagebox.showerror("Export Error", str(e)))
        finally:
            self.root.after(0, lambda: self.btn_search.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.btn_copy.config(state=tk.NORMAL))


def main():
    root = tk.Tk()
    app = ImageMatcherApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
