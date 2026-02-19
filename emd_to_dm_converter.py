"""
EMD to DM3 Converter
====================
GUI application to convert Velox .emd files (TEM diffraction data)
to Gatan Digital Micrograph .dm3 format.

Uses rsciio (RosettaSciIO) for reading EMD files — the same low-level
reader used by the reference emd-converter by Dr. Tao Ma — and a custom
DM writer for outputting DM3/DM4 files with full calibration metadata.

Requirements:
    pip install rosettasciio

Usage:
    python emd_to_dm_converter.py
"""

import os
import sys
import threading
import traceback
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import numpy as np

# --- dependency checks ---
HAS_DEPS = True
DEP_ERROR = ""
try:
    from rsciio.emd import file_reader as emd_reader
    from dm_writer import write_dm
except ImportError as e:
    HAS_DEPS = False
    DEP_ERROR = str(e)


class EMDConverterApp:
    """GUI application for converting EMD files to DM3/DM4."""

    def __init__(self, root):
        self.root = root
        self.root.title("EMD \u2192 DM3 Converter (TEM Diffraction)")
        self.root.geometry("780x640")
        self.root.minsize(700, 560)
        self.root.configure(bg="#f0f0f0")

        self.files = []
        self.output_dir = tk.StringVar()
        self.output_format = tk.StringVar(value=".dm3")
        self.overwrite = tk.BooleanVar(value=False)
        self.is_converting = False

        self._build_ui()
        self._check_dependencies()

    # ------------------------------------------------------------------ UI
    def _build_ui(self):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TButton", padding=6)
        style.configure("Accent.TButton", foreground="white", background="#0078d4")
        style.configure("Header.TLabel", font=("Segoe UI", 11, "bold"))

        # ---------- Input files section ----------
        input_frame = ttk.LabelFrame(self.root, text="  Input EMD Files  ", padding=10)
        input_frame.pack(fill=tk.X, padx=12, pady=(12, 4))

        btn_row = ttk.Frame(input_frame)
        btn_row.pack(fill=tk.X)

        ttk.Button(btn_row, text="Add Files…", command=self._add_files).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btn_row, text="Add Folder…", command=self._add_folder).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btn_row, text="Remove Selected", command=self._remove_selected).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btn_row, text="Clear All", command=self._clear_files).pack(side=tk.LEFT)

        # File list with scrollbar
        list_frame = ttk.Frame(input_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(8, 0))

        self.file_tree = ttk.Treeview(
            list_frame,
            columns=("path", "size", "status"),
            show="headings",
            selectmode="extended",
            height=8,
        )
        self.file_tree.heading("path", text="File Path")
        self.file_tree.heading("size", text="Size")
        self.file_tree.heading("status", text="Status")
        self.file_tree.column("path", width=420, minwidth=200)
        self.file_tree.column("size", width=80, minwidth=60, anchor=tk.E)
        self.file_tree.column("status", width=120, minwidth=80, anchor=tk.CENTER)

        vsb = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.file_tree.yview)
        self.file_tree.configure(yscrollcommand=vsb.set)
        self.file_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        # ---------- Output settings section ----------
        output_frame = ttk.LabelFrame(self.root, text="  Output Settings  ", padding=10)
        output_frame.pack(fill=tk.X, padx=12, pady=4)

        # Output directory
        dir_row = ttk.Frame(output_frame)
        dir_row.pack(fill=tk.X, pady=(0, 6))

        ttk.Label(dir_row, text="Output folder:").pack(side=tk.LEFT, padx=(0, 6))
        ttk.Entry(dir_row, textvariable=self.output_dir, width=50).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        ttk.Button(dir_row, text="Browse…", command=self._browse_output).pack(side=tk.LEFT)

        # Format and overwrite
        opt_row = ttk.Frame(output_frame)
        opt_row.pack(fill=tk.X)

        ttk.Label(opt_row, text="Output format:").pack(side=tk.LEFT, padx=(0, 6))
        ttk.Label(opt_row, text=".dm3", font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=(0, 24))
        ttk.Checkbutton(opt_row, text="Overwrite existing files", variable=self.overwrite).pack(side=tk.LEFT)

        # ---------- Signal selection ----------
        sig_frame = ttk.LabelFrame(self.root, text="  Signal Selection (multi-signal EMD)  ", padding=10)
        sig_frame.pack(fill=tk.X, padx=12, pady=4)

        self.signal_mode = tk.StringVar(value="all")
        ttk.Radiobutton(sig_frame, text="Export all signals (each as separate file)",
                        variable=self.signal_mode, value="all").pack(anchor=tk.W)
        ttk.Radiobutton(sig_frame, text="Export first signal only",
                        variable=self.signal_mode, value="first").pack(anchor=tk.W)

        # ---------- Progress ----------
        prog_frame = ttk.Frame(self.root, padding=(12, 4))
        prog_frame.pack(fill=tk.X)

        self.progress = ttk.Progressbar(prog_frame, mode="determinate")
        self.progress.pack(fill=tk.X, pady=(0, 4))

        self.status_label = ttk.Label(prog_frame, text="Ready. Add EMD files to begin.", anchor=tk.W)
        self.status_label.pack(fill=tk.X)

        # ---------- Convert / Quit ----------
        btn_frame = ttk.Frame(self.root, padding=12)
        btn_frame.pack(fill=tk.X)

        self.convert_btn = ttk.Button(btn_frame, text="Convert", style="Accent.TButton",
                                      command=self._start_conversion)
        self.convert_btn.pack(side=tk.RIGHT, padx=(6, 0))
        ttk.Button(btn_frame, text="Quit", command=self.root.quit).pack(side=tk.RIGHT)

    # --------------------------------------------------------- dependency check
    def _check_dependencies(self):
        if not HAS_DEPS:
            messagebox.showwarning(
                "Missing dependency",
                f"Required packages are not installed.\n\n"
                f"Error: {DEP_ERROR}\n\n"
                f"Install with:\n"
                f"   pip install hyperspy[all]\n\n"
                f"The converter will not work without them."
            )
            self.convert_btn.configure(state=tk.DISABLED)

    # --------------------------------------------------------- file helpers
    @staticmethod
    def _human_size(nbytes):
        for unit in ("B", "KB", "MB", "GB"):
            if abs(nbytes) < 1024:
                return f"{nbytes:.1f} {unit}"
            nbytes /= 1024
        return f"{nbytes:.1f} TB"

    def _add_file_to_tree(self, path):
        """Add a single file path to the tree if not duplicate."""
        path = os.path.normpath(path)
        # skip duplicates
        for item in self.file_tree.get_children():
            if self.file_tree.item(item)["values"][0] == path:
                return
        size = self._human_size(os.path.getsize(path))
        self.file_tree.insert("", tk.END, values=(path, size, "Pending"))

    def _add_files(self):
        paths = filedialog.askopenfilenames(
            title="Select EMD files",
            filetypes=[("EMD files", "*.emd"), ("All files", "*.*")],
        )
        for p in paths:
            self._add_file_to_tree(p)
        self._update_status_count()

    def _add_folder(self):
        folder = filedialog.askdirectory(title="Select folder containing EMD files")
        if not folder:
            return
        count = 0
        for root, _, files in os.walk(folder):
            for f in files:
                if f.lower().endswith(".emd"):
                    self._add_file_to_tree(os.path.join(root, f))
                    count += 1
        if count == 0:
            messagebox.showinfo("No EMD files", "No .emd files found in the selected folder.")
        self._update_status_count()

    def _remove_selected(self):
        for item in self.file_tree.selection():
            self.file_tree.delete(item)
        self._update_status_count()

    def _clear_files(self):
        self.file_tree.delete(*self.file_tree.get_children())
        self._update_status_count()

    def _browse_output(self):
        d = filedialog.askdirectory(title="Select output folder")
        if d:
            self.output_dir.set(d)

    def _update_status_count(self):
        n = len(self.file_tree.get_children())
        self.status_label.config(text=f"{n} file(s) queued.")

    # --------------------------------------------------------- conversion
    def _set_item_status(self, item, status):
        vals = list(self.file_tree.item(item)["values"])
        vals[2] = status
        self.file_tree.item(item, values=vals)

    def _start_conversion(self):
        items = self.file_tree.get_children()
        if not items:
            messagebox.showinfo("Nothing to convert", "Please add EMD files first.")
            return

        out_dir = self.output_dir.get().strip()
        if not out_dir:
            # default: same directory as each input file
            out_dir = None
        elif not os.path.isdir(out_dir):
            try:
                os.makedirs(out_dir, exist_ok=True)
            except OSError as e:
                messagebox.showerror("Output folder error", str(e))
                return

        self.is_converting = True
        self.convert_btn.configure(state=tk.DISABLED)
        self.progress["maximum"] = len(items)
        self.progress["value"] = 0

        # Run in background thread to keep GUI responsive
        thread = threading.Thread(
            target=self._convert_worker,
            args=(items, out_dir),
            daemon=True,
        )
        thread.start()

    def _convert_worker(self, items, out_dir):
        fmt = self.output_format.get()       # ".dm3" or ".dm4"
        overwrite = self.overwrite.get()
        mode = self.signal_mode.get()
        success = 0
        fail = 0

        for idx, item in enumerate(items):
            filepath = self.file_tree.item(item)["values"][0]
            self.root.after(0, self.status_label.config,
                            {"text": f"Converting ({idx+1}/{len(items)}): {os.path.basename(filepath)}"})
            self.root.after(0, self._set_item_status, item, "Converting…")

            try:
                # --- Read EMD using rsciio (same approach as reference code) ---
                image_list = emd_reader(filepath, select_type='images')

                if not image_list or len(image_list) == 0:
                    self.root.after(0, self._set_item_status, item, "No images found")
                    fail += 1
                    self.root.after(0, self._advance_progress)
                    continue

                if mode == "first":
                    image_list = image_list[:1]

                base = os.path.splitext(os.path.basename(filepath))[0]
                target_dir = out_dir if out_dir else os.path.dirname(filepath)

                for img_idx, img_dict in enumerate(image_list):
                    data = img_dict['data']
                    axes = img_dict.get('axes', [])
                    metadata = img_dict.get('metadata', {})

                    # Get title from metadata (same as reference)
                    try:
                        title = metadata['General']['title']
                    except (KeyError, TypeError):
                        title = ''

                    # Extract calibration from axes list
                    scales, units, offsets = self._extract_calibration(axes)

                    # Handle 3D data stacks (DCFI / 4D-STEM style)
                    # Same logic as the reference converter
                    if data.ndim == 3:
                        stack_axes = list(axes)
                        if len(stack_axes) > 0:
                            stack_axes_2d = [a for a in stack_axes if a.get('index_in_array', -1) != 0]
                            for i, a in enumerate(stack_axes_2d):
                                a['index_in_array'] = i
                        else:
                            stack_axes_2d = []

                        scales_2d, units_2d, offsets_2d = self._extract_calibration(stack_axes_2d)
                        stack_num = data.shape[0]
                        for si in range(stack_num):
                            slice_data = data[si]
                            suffix = f"_{title}_{si}" if title else f"_{img_idx}_{si}"
                            out_name = f"{base}{suffix}{fmt}"
                            out_path = os.path.join(target_dir, out_name)

                            if os.path.exists(out_path) and not overwrite:
                                continue
                            write_dm(out_path, slice_data,
                                     pixel_scales=scales_2d,
                                     pixel_units=units_2d,
                                     pixel_offsets=offsets_2d,
                                     title=title or base,
                                     version=3)
                    else:
                        # Standard 2D image
                        if len(image_list) > 1:
                            suffix = f"_{title}" if title else f"_{img_idx}"
                            out_name = f"{base}{suffix}{fmt}"
                        else:
                            out_name = f"{base}{fmt}"

                        out_path = os.path.join(target_dir, out_name)
                        if os.path.exists(out_path) and not overwrite:
                            self.root.after(0, self._set_item_status, item, "Skipped (exists)")
                            continue
                        write_dm(out_path, data,
                                 pixel_scales=scales,
                                 pixel_units=units,
                                 pixel_offsets=offsets,
                                 title=title or base,
                                 version=3)

                self.root.after(0, self._set_item_status, item, "Done \u2713")
                success += 1

            except Exception as e:
                err_msg = str(e)
                print(f"Error converting {filepath}:\n{traceback.format_exc()}")
                self.root.after(0, self._set_item_status, item, "Error \u2717")
                self.root.after(0, lambda it=item, em=err_msg: self._attach_error(it, em))
                fail += 1

            self.root.after(0, self._advance_progress)

        summary = f"Conversion complete — {success} succeeded, {fail} failed."
        self.root.after(0, self.status_label.config, {"text": summary})
        self.root.after(0, lambda: self.convert_btn.configure(state=tk.NORMAL))
        self.root.after(0, lambda: messagebox.showinfo("Done", summary))
        self.is_converting = False

    @staticmethod
    def _extract_calibration(axes_list):
        """Extract (scales, units, offsets) tuples from rsciio axes list.

        Returns (y_scale, x_scale), (y_unit, x_unit), (y_offset, x_offset)
        matching the convention expected by write_dm.
        """
        scales  = [1.0, 1.0]
        units   = ["", ""]
        offsets = [0.0, 0.0]
        for ax in axes_list:
            idx = ax.get('index_in_array', None)
            if idx is not None and idx < 2:
                scales[idx]  = ax.get('scale', 1.0)
                units[idx]   = ax.get('units', '')
                offsets[idx] = ax.get('offset', 0.0)
        return tuple(scales), tuple(units), tuple(offsets)

    def _advance_progress(self):
        self.progress["value"] = self.progress["value"] + 1

    def _attach_error(self, item, msg):
        """Show error detail on double-click."""
        def _show(event, m=msg):
            messagebox.showerror("Conversion Error", m)
        self.file_tree.tag_bind(item, "<Double-1>", _show)


def main():
    root = tk.Tk()
    app = EMDConverterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
