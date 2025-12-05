# ---------------------------------------------------------
# 🧰 Excel/CSV Automation Tool - CustomTkinter Version
# © 2025 Amit Kadam
# Licensed under CC BY-NC 4.0 (Non-Commercial Use Only)
# Unauthorized selling or redistribution is prohibited.
# ---------------------------------------------------------
import os
import sys
import traceback
import threading
import time
from datetime import datetime

import pandas as pd
import matplotlib.pyplot as plt
from fpdf import FPDF

# Try to import CustomTkinter; if not installed, instruct installation and exit.
try:
    import customtkinter as ctk
    from tkinter import filedialog, messagebox
    ctk.set_appearance_mode("System")
    ctk.set_default_color_theme("blue")
except Exception as e:
    print("ERROR: customtkinter is not installed or failed to import.")
    print("Install it with: pip install customtkinter")
    raise e

# -------------------- LICENSE TEXT (embedded) --------------------
_LICENSE_TEXT = """
Creative Commons Attribution-NonCommercial 4.0 International License

You are free to:
- Share — copy and redistribute the material
- Adapt — remix, transform, and build upon the material

Under the following terms:
- Attribution — You must give appropriate credit.
- NonCommercial — You may not use the material for commercial purposes.
- No additional restrictions.

This license prohibits commercial resale or redistribution of this software.
"""
# -----------------------------------------------------------------

# -------------------- Utilities --------------------
def ensure_base_output_folder():
    base = os.getcwd()
    out = os.path.join(base, "output")
    os.makedirs(out, exist_ok=True)
    return out

def run_folder_name():
    return f"Data-Report-{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}"

def ts():
    return datetime.now().strftime("%Y%m%d_%H%M%S")

def load_file(path):
    """Load CSV or Excel file into a DataFrame."""
    if path.lower().endswith(".csv"):
        return pd.read_csv(path)
    else:
        return pd.read_excel(path)

def clean_df(df, progress_cb=None):
    """
    Clean a DataFrame:
    - Trim whitespace for object columns
    - Coerce numeric columns and fill NaNs with 0
    The progress_cb(percent:int, status:str) is called during processing if provided.
    """
    cols = list(df.columns)
    total = max(1, len(cols))
    for i, col in enumerate(cols, start=1):
        try:
            if pd.api.types.is_object_dtype(df[col]):
                df[col] = df[col].fillna("").astype(str).str.strip()
            elif pd.api.types.is_numeric_dtype(df[col]):
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            else:
                df[col] = df[col].fillna("")
        except Exception:
            df[col] = df[col].astype(str).fillna("").str.strip()

        if progress_cb:
            percent = int(i / total * 100)
            progress_cb(percent, f"Cleaning columns: {i}/{total}")
    return df

def merge_and_tag(files, progress_cb=None):
    """Read files, add Source_File column and combine. Progress callback receives load progress."""
    dfs = []
    read_files = []
    total = max(1, len(files))
    for idx, f in enumerate(files, start=1):
        try:
            df = load_file(f)
            df['Source_File'] = os.path.basename(f)
            dfs.append(df)
            read_files.append(os.path.basename(f))
        except Exception as e:
            print(f"Failed reading {f}: {e}")
        if progress_cb:
            pct = int(idx / total * 100)
            progress_cb(pct, f"Loaded {idx}/{total} files")
    if not dfs:
        return pd.DataFrame(), []
    combined = pd.concat(dfs, ignore_index=True)
    return combined, read_files

def remove_duplicates(df):
    """Drop duplicate rows, return new df and number removed."""
    before = len(df)
    df2 = df.drop_duplicates()
    removed = before - len(df2)
    return df2, removed

def generate_column_summary(df):
    """Create a simple per-column summary dictionary."""
    summary = {}
    for col in df.columns:
        try:
            if pd.api.types.is_numeric_dtype(df[col]):
                s = pd.to_numeric(df[col], errors='coerce').fillna(0)
                summary[col] = {
                    "type": "numeric",
                    "count": int(s.count()),
                    "sum": float(s.sum()),
                    "mean": float(s.mean()),
                    "min": float(s.min()),
                    "max": float(s.max())
                }
            else:
                s = df[col].astype(str).fillna("")
                most_common = s.mode().iloc[0] if not s.mode().empty else ""
                summary[col] = {
                    "type": "text",
                    "count": int(s.count()),
                    "unique": int(s.nunique()),
                    "most_common": most_common
                }
        except Exception as e:
            summary[col] = {"type": "unknown", "error": str(e)}
    return summary

# -------------------- Charts & PDF helpers --------------------
class PDFReport(FPDF):
    def header(self):
        self.set_font('Arial', 'B', 14)
        self.cell(0, 10, 'Data Processing Report', 0, 1, 'C')
        self.ln(2)

    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'Page {self.page_no()}', 0, 0, 'C')

def create_bar_chart_from_summary(summary_df, outpath):
    plt.figure(figsize=(6, 4))
    x = summary_df.iloc[:, 0].astype(str)
    y = pd.to_numeric(summary_df.iloc[:, 1], errors='coerce').fillna(0)
    plt.bar(x, y)
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    plt.savefig(outpath, dpi=150)
    plt.close()
    return outpath

def create_top_values_chart(df, column, outpath, top_n=10):
    s = df[column].astype(str).value_counts().iloc[:top_n]
    plt.figure(figsize=(6, 4))
    s.plot(kind='bar')
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    plt.savefig(outpath, dpi=150)
    plt.close()
    return outpath

def add_basic_stats(pdf, df, duplicates_removed, input_files):
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 8, "Dataset Summary", 0, 1)
    pdf.set_font("Arial", "", 10)
    pdf.cell(0, 6, f"Processed on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", 0, 1)
    pdf.cell(0, 6, f"Input files: {', '.join(input_files)}", 0, 1)
    pdf.cell(0, 6, f"Total rows (after processing): {len(df)}", 0, 1)
    pdf.cell(0, 6, f"Total columns: {len(df.columns)}", 0, 1)
    pdf.cell(0, 6, f"Duplicates removed: {duplicates_removed}", 0, 1)
    pdf.ln(4)

def add_column_summary(pdf, col_summary, max_cols=8):
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 8, "Column-wise Summary (sample)", 0, 1)
    pdf.set_font("Arial", "", 9)
    count = 0
    for col, stats in col_summary.items():
        if count >= max_cols:
            pdf.cell(0, 6, f"... and {len(col_summary)-max_cols} more columns", 0, 1)
            break
        pdf.set_font("Arial", "B", 10)
        pdf.cell(0, 6, f"{col}", 0, 1)
        pdf.set_font("Arial", "", 9)
        for k, v in stats.items():
            pdf.cell(0, 5, f"  {k}: {v}", 0, 1)
        pdf.ln(1)
        count += 1
    pdf.ln(4)

def add_table_from_df(pdf, df, max_rows=6):
    pdf.set_font("Arial", "B", 10)
    pdf.cell(0, 6, "Sample rows", 0, 1)
    pdf.set_font("Arial", "", 9)
    cols = list(df.columns)[:6]
    if not cols:
        pdf.cell(0, 6, "(No columns to display)", 0, 1)
        return
    col_width = pdf.w / (len(cols) + 1)
    row_h = 6
    for c in cols:
        pdf.cell(col_width, row_h, str(c)[:15], 1)
    pdf.ln(row_h)
    for i, (_, row) in enumerate(df.iterrows()):
        if i >= max_rows:
            break
        for c in cols:
            text = str(row[c])[:20]
            pdf.cell(col_width, row_h, text, 1)
        pdf.ln(row_h)
    pdf.ln(4)

def insert_chart_to_pdf(pdf, image_path, title=None):
    if title:
        pdf.set_font("Arial", "B", 11)
        pdf.cell(0, 6, title, 0, 1)
    try:
        pdf.image(image_path, x=15, w=pdf.w - 30)
        pdf.ln(6)
    except Exception as e:
        pdf.set_font("Arial", "", 10)
        pdf.cell(0, 5, f"Could not insert chart: {e}", 0, 1)

# -------------------- Main processing (with smart run folder and progress callbacks) --------------------
def process_and_report(files, progress_callback=None):
    """
    Processes provided files and stores all outputs inside a unique run folder.
    progress_callback(percent:int, status:str) will be called periodically if provided.
    """
    if not files:
        raise ValueError("No files provided.")
    base_output = ensure_base_output_folder()
    run_name = run_folder_name()
    output_folder = os.path.join(base_output, run_name)
    os.makedirs(output_folder, exist_ok=True)

    # Step 1: Load files
    if progress_callback:
        progress_callback(5, "Starting: loading files")
    combined, read_files = merge_and_tag(files, progress_cb=lambda p, s: progress_callback(5 + int(p * 0.1), s) if progress_callback else None)
    if combined.empty:
        raise ValueError("No readable data found in the provided files.")
    if progress_callback:
        progress_callback(25, "Files loaded; merging complete")

    # Step 2: Clean data
    if progress_callback:
        progress_callback(30, "Cleaning data")
    cleaned = clean_df(combined, progress_cb=lambda p, s: progress_callback(30 + int(p * 0.2), s) if progress_callback else None)
    if progress_callback:
        progress_callback(55, "Cleaning complete")

    # Step 3: Deduplicate
    if progress_callback:
        progress_callback(60, "Removing duplicates")
    cleaned, removed = remove_duplicates(cleaned)
    if progress_callback:
        progress_callback(70, f"Duplicates removed: {removed}")

    # Step 4: Save outputs
    timestamp = ts()
    merged_path = os.path.join(output_folder, f"merged_output_{timestamp}.xlsx")
    cleaned_path = os.path.join(output_folder, f"cleaned_output_{timestamp}.xlsx")
    if progress_callback:
        progress_callback(75, "Saving output files")
    cleaned.to_excel(merged_path, index=False)
    cleaned.to_excel(cleaned_path, index=False)

    # Step 5: Summaries & charts
    if progress_callback:
        progress_callback(80, "Generating summaries")
    col_summary = generate_column_summary(cleaned)

    if 'Source_File' in cleaned.columns:
        summary_df = cleaned.groupby('Source_File').size().reset_index(name='Rows')
    else:
        summary_df = pd.DataFrame({'All_Data': ['All'], 'Rows': [len(cleaned)]})

    chart_paths = []
    if progress_callback:
        progress_callback(85, "Creating chart: summary")
    chart1 = os.path.join(output_folder, f"chart_summary_{timestamp}.png")
    create_bar_chart_from_summary(summary_df, chart1)
    chart_paths.append(chart1)

    text_cols = [c for c in cleaned.columns if not pd.api.types.is_numeric_dtype(cleaned[c])]
    if text_cols:
        if progress_callback:
            progress_callback(88, f"Creating chart: top values ({text_cols[0]})")
        chart2 = os.path.join(output_folder, f"chart_top_values_{timestamp}.png")
        create_top_values_chart(cleaned, text_cols[0], chart2)
        chart_paths.append(chart2)

    # Step 6: PDF
    if progress_callback:
        progress_callback(92, "Building PDF report")
    pdf_path = os.path.join(output_folder, f"summary_report_{timestamp}.pdf")
    pdf = PDFReport()
    pdf.add_page()
    try:
        add_basic_stats(pdf, cleaned, removed, read_files)
        add_column_summary(pdf, col_summary, max_cols=8)
        add_table_from_df(pdf, cleaned, max_rows=6)
        for cp in chart_paths:
            insert_chart_to_pdf(pdf, cp)
    except Exception as e:
        pdf.add_page()
        pdf.set_font("Arial", "", 10)
        pdf.multi_cell(0, 6, f"Error building report sections: {e}\n\n{traceback.format_exc()}")
    pdf.output(pdf_path)
    if progress_callback:
        progress_callback(99, "Finalizing")

    return {
        "merged": merged_path,
        "cleaned": cleaned_path,
        "pdf": pdf_path,
        "charts": chart_paths,
        "output_folder": output_folder,
        "rows": len(cleaned),
        "duplicates_removed": removed,
        "read_files": read_files
    }

# -------------------- GUI App (CustomTkinter) with popup progress window --------------------
class ProgressPopup:
    def __init__(self, parent):
        self.top = ctk.CTkToplevel(parent)
        self.top.title("Processing... ⏳")
        self.top.geometry("420x140")
        self.top.resizable(False, False)
        self.top.grab_set()  # modal

        self.label = ctk.CTkLabel(self.top, text="Starting...", anchor="w")
        self.label.pack(fill="x", padx=12, pady=(12,6))

        self.progress = ctk.CTkProgressBar(self.top, width=380)
        self.progress.set(0)
        self.progress.pack(padx=12, pady=(0,6))

        self.percent_label = ctk.CTkLabel(self.top, text="0%", anchor="e")
        self.percent_label.pack(fill="x", padx=12, pady=(0,8))

    def update(self, percent, status):
        # percent 0..100
        try:
            self.progress.set(percent / 100.0)
            self.label.configure(text=status)
            self.percent_label.configure(text=f"{int(percent)}%")
            self.top.update_idletasks()
        except Exception:
            pass

    def close(self):
        try:
            self.top.grab_release()
            self.top.destroy()
        except Exception:
            pass

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel/CSV Automation Tool - Progress ✅")
        self.root.geometry("920x680")
        self.files = []
        self.last_run_folder = None
        self.output_base = ensure_base_output_folder()
        self._build_ui()

    def _build_ui(self):
        pad = 12
        main = ctk.CTkFrame(self.root, corner_radius=10)
        main.pack(fill="both", expand=True, padx=pad, pady=pad)

        title = ctk.CTkLabel(main, text="Excel/CSV Automation Tool", font=ctk.CTkFont(size=20, weight="bold"))
        title.pack(pady=(8, 6))

        top_frame = ctk.CTkFrame(main, fg_color="transparent")
        top_frame.pack(fill="x", padx=10, pady=(6, 4))

        self.btn_add = ctk.CTkButton(top_frame, text="Select Files", command=self.choose_files, width=140)
        self.btn_add.pack(side="left", padx=(0,10))

        self.btn_sample = ctk.CTkButton(top_frame, text="Load Sample", command=self.load_sample, width=120)
        self.btn_sample.pack(side="left", padx=(0,10))

        self.btn_clear = ctk.CTkButton(top_frame, text="Clear List", command=self.clear_list, width=120)
        self.btn_clear.pack(side="left", padx=(0,10))

        lbl_hint = ctk.CTkLabel(top_frame, text="Drag & drop removed. Use Select Files.", anchor="w")
        lbl_hint.pack(side="left", padx=(10,0))

        center = ctk.CTkFrame(main)
        center.pack(fill="both", expand=True, padx=10, pady=(6,10))

        left = ctk.CTkFrame(center, width=460)
        left.pack(side="left", fill="both", expand=False, padx=(0,10))
        self.textbox = ctk.CTkTextbox(left, width=460, height=360, state="disabled")
        self.textbox.pack(padx=6, pady=6)

        right = ctk.CTkFrame(center)
        right.pack(side="left", fill="both", expand=True)

        lbl_opts = ctk.CTkLabel(right, text="Options & Preview", font=ctk.CTkFont(size=14, weight="bold"))
        lbl_opts.pack(pady=(6,8))

        self.preview_box = ctk.CTkTextbox(right, height=220, state="disabled")
        self.preview_box.pack(fill="both", expand=True, padx=6, pady=6)

        bottom = ctk.CTkFrame(main, fg_color="transparent")
        bottom.pack(fill="x", padx=10, pady=(6,12))

        self.btn_process = ctk.CTkButton(bottom, text="Process & Generate Report", command=self.on_process_click, fg_color="#2196F3")
        self.btn_process.pack(side="left", padx=(0,10))

        self.btn_open = ctk.CTkButton(bottom, text="Open Last Run Folder", command=self.open_last_run_folder)
        self.btn_open.pack(side="left", padx=(6,10))

        self.btn_open_base = ctk.CTkButton(bottom, text="Open Output Base Folder", command=self.open_output_base_folder)
        self.btn_open_base.pack(side="left", padx=(6,10))

        self._refresh_listbox()

    # ---------- file handling ----------
    def choose_files(self):
        paths = filedialog.askopenfilenames(title="Select Excel or CSV Files",
                                            filetypes=[("Excel/CSV Files", "*.csv *.xlsx *.xls")])
        if paths:
            self.add_files(list(paths))

    def add_files(self, paths):
        new = []
        for p in paths:
            if p and os.path.exists(p) and p not in self.files:
                self.files.append(p)
                new.append(os.path.basename(p))
        self._refresh_listbox()
        if new:
            self.show_preview_for_file(self.files[-1])

    def clear_list(self):
        self.files = []
        self._refresh_listbox()
        self.clear_preview()

    def load_sample(self):
        # sample path present in your environment — change if needed
        sample = "/mnt/data/Practice-Data.xlsx"
        if os.path.exists(sample):
            self.add_files([sample])
            messagebox.showinfo("Sample loaded", f"Loaded sample: {sample}")
        else:
            messagebox.showwarning("Sample missing", f"No sample at {sample}")

    def _refresh_listbox(self):
        self.textbox.configure(state="normal")
        self.textbox.delete("0.0", "end")
        if not self.files:
            self.textbox.insert("0.0", "No files selected.\nUse 'Select Files'.\n")
        else:
            for p in self.files:
                self.textbox.insert("end", os.path.basename(p) + "\n")
        self.textbox.configure(state="disabled")

    # ---------- preview (fixed) ----------
    def show_preview_for_file(self, path):
        try:
            df = load_file(path)
            preview_df = df.iloc[:10, :8].copy()
            for col in preview_df.columns:
                preview_df[col] = preview_df[col].astype(str).apply(lambda x: (x[:80] + "...") if len(x) > 80 else x)
            preview_str = preview_df.to_string(index=False)
            self.preview_box.configure(state="normal")
            self.preview_box.delete("0.0", "end")
            self.preview_box.insert("0.0", preview_str)
            self.preview_box.configure(state="disabled")
        except Exception as e:
            self.preview_box.configure(state="normal")
            self.preview_box.delete("0.0", "end")
            self.preview_box.insert("0.0", f"Preview error:\n{e}")
            self.preview_box.configure(state="disabled")

    def clear_preview(self):
        self.preview_box.configure(state="normal")
        self.preview_box.delete("0.0", "end")
        self.preview_box.insert("0.0", "No preview available.")
        self.preview_box.configure(state="disabled")

    # ---------- processing with progress popup ----------
    def on_process_click(self):
        if not self.files:
            messagebox.showerror("No files", "Please select at least one file before processing.")
            return
        # disable UI buttons to prevent double-run
        self._set_ui_state(enabled=False)
        self.progress_popup = ProgressPopup(self.root)
        # start processing in background thread
        thread = threading.Thread(target=self._background_process, daemon=True)
        thread.start()

    def _background_process(self):
        def progress_cb(percent, status):
            # called from worker thread; schedule UI update on main thread
            self.root.after(0, lambda: self.progress_popup.update(percent, status))

        try:
            res = process_and_report(self.files, progress_callback=progress_cb)
            self.last_run_folder = res.get("output_folder")
            # final update to 100%
            self.root.after(0, lambda: self.progress_popup.update(100, "Completed ✅"))
            time.sleep(0.25)
            self.root.after(0, lambda: self.progress_popup.close())
            # re-enable UI
            self.root.after(0, lambda: self._set_ui_state(enabled=True))
            # show results
            msg = [
                "Processing complete!",
                f"Rows (after processing): {res['rows']}",
                f"Duplicates removed: {res['duplicates_removed']}",
                f"Saved (merged): {os.path.basename(res['merged'])}",
                f"Saved (report): {os.path.basename(res['pdf'])}",
                f"Run folder: {os.path.basename(self.last_run_folder) if self.last_run_folder else 'N/A'}"
            ]
            self.root.after(0, lambda: messagebox.showinfo("Done", "\n".join(msg)))
            # open folder
            if self.last_run_folder:
                self.root.after(500, lambda: self.open_folder(self.last_run_folder))
        except Exception as e:
            traceback.print_exc()
            self.root.after(0, lambda: self.progress_popup.update(100, f"Error: {e}"))
            time.sleep(0.5)
            self.root.after(0, lambda: self.progress_popup.close())
            self.root.after(0, lambda: self._set_ui_state(enabled=True))
            self.root.after(0, lambda: messagebox.showerror("Processing error", str(e)))

    def _set_ui_state(self, enabled=True):
        state = "normal" if enabled else "disabled"
        # enable/disable primary buttons
        try:
            self.btn_add.configure(state=state)
            self.btn_sample.configure(state=state)
            self.btn_clear.configure(state=state)
            self.btn_process.configure(state=state)
            self.btn_open.configure(state=state)
            self.btn_open_base.configure(state=state)
        except Exception:
            pass

    def open_folder(self, folder):
        try:
            if sys.platform.startswith('darwin'):
                os.system(f'open "{folder}"')
            elif os.name == 'nt':
                os.startfile(folder)
            elif os.name == 'posix':
                os.system(f'xdg-open "{folder}"')
        except Exception as e:
            messagebox.showerror("Could not open folder", str(e))

    def open_last_run_folder(self):
        folder = getattr(self, "last_run_folder", None)
        if folder and os.path.exists(folder):
            self.open_folder(folder)
        else:
            messagebox.showinfo("No run folder", "No previous run folder found. Run a process first.")

    def open_output_base_folder(self):
        if os.path.exists(self.output_base):
            self.open_folder(self.output_base)
        else:
            messagebox.showerror("Output folder missing", f"Expected output base at {self.output_base}")

# -------------------- Run --------------------
def main():
    root = ctk.CTk()
    app = App(root)
    root.mainloop()

if __name__ == "__main__":
    main()
