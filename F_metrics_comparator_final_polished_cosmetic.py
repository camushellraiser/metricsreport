
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
import os
import pandas as pd
import math
import xlsxwriter
from collections import defaultdict

original_files = []
reduced_files = []

def base_name(path):
    filename = os.path.basename(path)
    return filename.replace("_Reduced TM", "").replace("_Metrics.xlsx", "").replace(".xlsx", "").strip()

def extract_header_and_table(df):
    start_row = None
    for idx, row in df.iterrows():
        if str(row.iloc[0]).strip().lower() == "total count":
            start_row = idx - 1
            break
    if start_row is None:
        raise ValueError("No 'Total count' found.")
    full_table_data = []
    for idx in range(start_row, len(df)):
        row_data = df.iloc[idx, :7].tolist()
        full_table_data.append(row_data)
        if str(df.iloc[idx, 0]).strip() == "No matching":
            break
    table = pd.DataFrame(full_table_data, columns=[
        "Initial", "Segments", "Words", "Word %", "Characters", "Characters %", "Characters excluding spaces"
    ])
    table = table[~table['Initial'].isin(['Initial', 'Metric'])]
    return table

def process_all_and_save():
    if not original_files or not reduced_files:
        messagebox.showerror("Missing Files", "Please drop both original and reduced TM files.")
        return

    paired = {}
    for path in original_files:
        name = base_name(path)
        paired[name] = {'original': path, 'reduced': None}
    for path in reduced_files:
        name = base_name(path)
        if name in paired:
            paired[name]['reduced'] = path

    output_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if not output_path:
        return

    workbook = xlsxwriter.Workbook(output_path)
    worksheet = workbook.add_worksheet()
    row_cursor = 0

    header_format = workbook.add_format({'bold': True, 'font_color': 'black'})
    title_format = workbook.add_format({'bold': True, 'font_color': 'red'})
    center_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
    left_format = workbook.add_format({'align': 'left', 'valign': 'vcenter', 'border': 1})
    bold_format = workbook.add_format({'bold': True, 'border': 1})
    red_bold = workbook.add_format({'bold': True, 'font_color': 'red'})
    regular = workbook.add_format({'font_color': 'black'})
    header_bg = workbook.add_format({'bold': True, 'bg_color': '#D9D9D9', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True})

    delta_summary = {}

    for name, files in paired.items():
        if files['original'] and files['reduced']:
            df1 = pd.read_excel(files['original'], header=None)
            df2 = pd.read_excel(files['reduced'], header=None)
            t1 = extract_header_and_table(df1)
            t2 = extract_header_and_table(df2)

            worksheet.merge_range(row_cursor, 0, row_cursor, 6, "", center_format)
            worksheet.merge_range(row_cursor, 8, row_cursor, 14, "", center_format)
            worksheet.write_rich_string(row_cursor, 0,
                red_bold, "File Name 1: ",
                regular, os.path.basename(files['original']).replace(".xlsx", ""),
                red_bold, " Original Project",
                center_format)
            worksheet.write_rich_string(row_cursor, 8,
                red_bold, "File Name 2: ",
                regular, os.path.basename(files['reduced']).replace(".xlsx", ""),
                red_bold, " New Project",
                center_format)
            row_cursor += 1

            headers = ["Initial", "Segments", "Words", "Word %", "Characters", "Characters %", "Characters excluding spaces"]
            for idx, h in enumerate(headers):
                worksheet.write(row_cursor, idx, h, header_bg)
                worksheet.write(row_cursor, idx+8, h, header_bg)
                worksheet.write(row_cursor, idx+16, h + " Δ" if idx > 0 else "Metric", header_bg)
            worksheet.set_row(row_cursor, 30)
            worksheet.set_column(1, 1, 11)  # Segments
            worksheet.set_column(2, 2, 11)  # Words
            worksheet.set_column(3, 3, 10)  # Word %
            worksheet.set_column(4, 4, 13)  # Characters
            worksheet.set_column(5, 5, 15)  # Characters %
            worksheet.set_column(6, 6, 24)  # Characters excluding spaces
            worksheet.set_column(9, 9, 11)  # Segments
            worksheet.set_column(10, 10, 11)  # Words
            worksheet.set_column(11, 11, 10)  # Word %
            worksheet.set_column(12, 12, 13)  # Characters
            worksheet.set_column(13, 13, 15)  # Characters %
            worksheet.set_column(14, 14, 24)  # Characters excluding spaces
            worksheet.set_column(17, 17, 11)  # Segments Δ
            worksheet.set_column(18, 18, 11)  # Words Δ
            worksheet.set_column(19, 19, 10)  # Word % Δ
            worksheet.set_column(20, 20, 13)  # Characters Δ
            worksheet.set_column(21, 21, 15)  # Characters % Δ
            worksheet.set_column(22, 22, 24)  # Characters excluding spaces Δ
            row_cursor += 1

            merged = pd.merge(t1, t2, on="Initial", how="inner")

            for _, row in merged.iterrows():
                metric = row['Initial']
                for j, col in enumerate(headers):
                    val1 = row.get(col + "_x", row.get(col, ""))
                    val2 = row.get(col + "_y", row.get(col, ""))
                    val1_fmt = bold_format if metric in ["Translation memory matching", "Internal matching"] else (left_format if j == 0 else center_format)
                    val2_fmt = val1_fmt

                    worksheet.write(row_cursor, j, "" if pd.isna(val1) else val1, val1_fmt)
                    worksheet.write(row_cursor, j+8, "" if pd.isna(val2) else val2, val2_fmt)

                    if j == 0:
                        worksheet.write(row_cursor, j+16, metric, val1_fmt)
                    else:
                        try:
                            # If the values are the same, use the value instead of 0
                            if val1 == val2:
                                diff = float(val1)
                                fmt = workbook.add_format({
                                    "align": "center",
                                    "valign": "vcenter",
                                    "border": 1,
                                    "font_color": "black"
                                })
                                worksheet.write(row_cursor, j+16, str(int(diff)), fmt)
                            else:
                                diff = float(val2) - float(val1)
                                fmt = workbook.add_format({"align": "center", "valign": "vcenter", "border": 1, "font_color": "green" if diff > 0 else "red" if diff < 0 else "black"})
                                worksheet.write(row_cursor, j+16, f"+{int(diff)}" if diff > 0 else str(int(diff)), fmt)
                            if metric not in delta_summary:
                                delta_summary[metric] = defaultdict(float)
                            delta_summary[metric][j] += diff
                        except:
                            worksheet.write(row_cursor, j+16, "0", center_format)
                row_cursor += 1

            row_cursor += 1

    # Final combined summary
    
    # Final combined summary
    if delta_summary:
        worksheet.write(row_cursor, 0, "📊 Combined Delta Summary", red_bold)
        row_cursor += 1
        summary_headers = ["Metric", "Segments Δ", "Words Δ", "Word % Δ", "Characters Δ", "Characters % Δ", "Characters excluding spaces Δ"]
        for j, h in enumerate(summary_headers):
            worksheet.write(row_cursor, j, h, header_bg)
        worksheet.set_row(row_cursor, 30)
        row_cursor += 1
        for metric in delta_summary:
            is_total_count = metric.strip().lower() == "total count"
            row_format = workbook.add_format({"align": "center", "valign": "vcenter", "border": 1, "font_color": "black"}) if is_total_count else (bold_format if metric in ["Translation memory matching", "Internal matching"] else left_format)
            worksheet.write(row_cursor, 0, metric, row_format)
            for j in range(1, len(summary_headers)):
                total = delta_summary[metric].get(j, 0)
                if is_total_count:
                    fmt = workbook.add_format({
                        "align": "center",
                        "valign": "vcenter",
                        "border": 1,
                        "font_color": "black"
                    })
                    worksheet.write(row_cursor, j, str(int(total)), fmt)
                else:
                    fmt = workbook.add_format({
                        "align": "center", "valign": "vcenter", "border": 1,
                        "font_color": "green" if total > 0 else "red" if total < 0 else "black"
                    })
                    worksheet.write(row_cursor, j, f"+{int(total)}" if total > 0 else str(int(total)), fmt)
            row_cursor += 1

    workbook.close()
    messagebox.showinfo("✅ Done", f"Comparison report saved:{output_path}")

def handle_drop(event, target_list, label):
    paths = root.tk.splitlist(event.data)
    target_list.clear()
    label_text = []
    for p in paths:
        if p.endswith(".xlsx"):
            target_list.append(p)
            label_text.append(os.path.basename(p))
    label.config(text="\n".join(label_text) if label_text else "Drop .xlsx files here")
    if len(label_text) > 6:
        label.config(height=20)
    else:
        label.config(height=8)

def browse_files(target_list, label):
    files = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx")])
    for file in files:
        target_list.append(file)
    label_text = [os.path.basename(f) for f in target_list]
    label.config(text="\n".join(label_text) if label_text else "Drop .xlsx files here")
    if len(label_text) > 6:
        label.config(height=20)
    else:
        label.config(height=8)

def remove_files(target_list, label):
    target_list.clear()

def load_all_files():
    folder_selected = filedialog.askdirectory()
    if not folder_selected:
        return
    original_files.clear()
    reduced_files.clear()
    original_label_text = []
    reduced_label_text = []
    for file in os.listdir(folder_selected):
        full_path = os.path.join(folder_selected, file)
        if not file.lower().endswith(".xlsx"):
            continue
        if "_reduced tm_metrics" in file.lower():
            reduced_files.append(full_path)
            reduced_label_text.append(file)
        elif "_metrics" in file.lower() and "reduced tm" not in file.lower():
            original_files.append(full_path)
            original_label_text.append(file)
    original_label.config(text="\n".join(original_label_text) if original_label_text else "Drop .xlsx files here")
    reduced_label.config(text="\n".join(reduced_label_text) if reduced_label_text else "Drop .xlsx files here")
    original_label.config(height=20 if len(original_label_text) > 6 else 8)
    reduced_label.config(height=20 if len(reduced_label_text) > 6 else 8)


root = TkinterDnD.Tk()
root.title("Metrics Comparator 2.0")
root.geometry("850x600")
root.configure(bg="#F9F9F9")

tk.Label(root, text="Original Project Files", font=("Segoe UI", 12, "bold"), bg="#F9F9F9").pack(pady=(10, 0))
original_label = tk.Label(root, text="Drop .xlsx files here", relief="solid", width=80, height=8, bg="white", anchor="nw", justify="left")
original_label.pack(pady=5)
original_label.drop_target_register(DND_FILES)
original_label.dnd_bind("<<Drop>>", lambda e: handle_drop(e, original_files, original_label))

browse_original_btn = tk.Button(root, text="Browse", bg="#4CAF50", fg="white", command=lambda: browse_files(original_files, original_label))
browse_original_btn.pack(pady=5)

remove_original_btn = tk.Button(root, text="Remove", bg="red", fg="white", command=lambda: remove_files(original_files, original_label))
remove_original_btn.pack(pady=5)

tk.Label(root, text="Reduced TM Project Files", font=("Segoe UI", 12, "bold"), bg="#F9F9F9").pack(pady=(15, 0))
reduced_label = tk.Label(root, text="Drop .xlsx files here", relief="solid", width=80, height=8, bg="white", anchor="nw", justify="left")
reduced_label.pack(pady=5)
reduced_label.drop_target_register(DND_FILES)
reduced_label.dnd_bind("<<Drop>>", lambda e: handle_drop(e, reduced_files, reduced_label))

browse_reduced_btn = tk.Button(root, text="Browse", bg="#4CAF50", fg="white", command=lambda: browse_files(reduced_files, reduced_label))
browse_reduced_btn.pack(pady=5)

remove_reduced_btn = tk.Button(root, text="Remove", bg="red", fg="white", command=lambda: remove_files(reduced_files, reduced_label))
remove_reduced_btn.pack(pady=5)

frame = tk.Frame(root, bg="#F9F9F9")
frame.pack(pady=20)

load_all_btn = tk.Button(frame, text="📂 Load All", command=load_all_files, bg="#2196F3", fg="white", font=("Segoe UI", 10, "bold"), height=2, width=15)
load_all_btn.pack(side="left", padx=10)

def load_all_files():
    folder_selected = filedialog.askdirectory()
    if not folder_selected:
        return
    original_files.clear()
    reduced_files.clear()
    original_label_text = []
    reduced_label_text = []
    for file in os.listdir(folder_selected):
        full_path = os.path.join(folder_selected, file)
        if not file.lower().endswith(".xlsx"):
            continue
        if "_reduced tm_metrics" in file.lower():
            reduced_files.append(full_path)
            reduced_label_text.append(file)
        elif "_metrics" in file.lower() and "reduced tm" not in file.lower():
            original_files.append(full_path)
            original_label_text.append(file)
    original_label.config(text="\n".join(original_label_text) if original_label_text else "Drop .xlsx files here")
    reduced_label.config(text="\n".join(reduced_label_text) if reduced_label_text else "Drop .xlsx files here")
    original_label.config(height=20 if len(original_label_text) > 6 else 8)
    reduced_label.config(height=20 if len(reduced_label_text) > 6 else 8)

compare_btn = tk.Button(frame, text="Compare All and Export Report", command=process_all_and_save, bg="#4CAF50", fg="white", font=("Segoe UI", 11, "bold"), height=2, width=30)
compare_btn.pack(side="left", padx=10)


def load_all_files():
    folder_selected = filedialog.askdirectory()
    if not folder_selected:
        return
    original_files.clear()
    reduced_files.clear()
    original_label_text = []
    reduced_label_text = []

    for file in os.listdir(folder_selected):
        full_path = os.path.join(folder_selected, file)
        if not file.lower().endswith(".xlsx"):
            continue
        if "_reduced tm_metrics" in file.lower():
            reduced_files.append(full_path)
            reduced_label_text.append(file)
        elif "_metrics" in file.lower() and "reduced tm" not in file.lower():
            original_files.append(full_path)
            original_label_text.append(file)

    original_label.config(text="\n".join(original_label_text) if original_label_text else "Drop .xlsx files here")
    reduced_label.config(text="\n".join(reduced_label_text) if reduced_label_text else "Drop .xlsx files here")

    original_label.config(height=20 if len(original_label_text) > 6 else 8)
    reduced_label.config(height=20 if len(reduced_label_text) > 6 else 8)


root.mainloop()
