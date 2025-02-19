#!/usr/bin/env python3
import os
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from datetime import datetime

# -------------------------------
# Constants and Configuration
# -------------------------------
REF_HEADERS = ['Descr', 'OPC', 'SKU']
DATA_HEADERS = ['ADD', 'On Hand', 'Free ROD']

CATEGORY_SHEETS = {
    "FSV": {
        "ADD": "FSV_ADD",
        "On Hand": "FSV_OnHand",
        "Free ROD": "FSV_FreeROD"
    },
    "SF_PUCK": {
        "ADD": "SF_PUCK_ADD",
        "On Hand": "SF_PUCK_OnHand",
        "Free ROD": "SF_PUCK_FreeROD"
    }
}

# -------------------------------
# Formatting Functions
# -------------------------------
def format_add(value):
    """
    For the ADD column:
    - If the value is "0", "0.0" or empty, return 0.0001.
    - If the value is all digits and longer than 4 characters, treat the last 4 digits as fractional.
    - Otherwise, if fewer than 4 digits, zero-pad to 4 digits and return as fraction.
    """
    try:
        s = str(value).replace(",", "").strip()
        if s in ("", "0", "0.0"):
            return 0.0001
        if s.isdigit():
            if len(s) > 4:
                return float(s[:-4] + "." + s[-4:])
            else:
                return float("0." + s.zfill(4))
        return float(s)
    except (ValueError, TypeError):
        return 0.0

def format_free_rod(value):
    """
    For the Free ROD column:
    - If the value is "0", "0.0" or empty, return 0.0.
    - Otherwise, simply convert the value to float.
    """
    try:
        s = str(value).replace(",", "").strip()
        if s in ("", "0", "0.0"):
            return 0.0
        if s.isdigit():
            return float(s)
        return float(s)
    except (ValueError, TypeError):
        return 0.0

# -------------------------------
# Data Processing Functions
# -------------------------------
def process_input_file(file_path, file_index):
    """
    Reads an Excel file ensuring the reference headers are present.
    Tries the first row as header; if not found, tries the second row.
    Then extracts the reference and data columns, applying:
      - format_add to the ADD column,
      - format_free_rod to the Free ROD column.
    Renames data columns with a file index suffix and returns the combined DataFrame and the file's timestamp (DD/MM/YYYY).
    """
    df = pd.read_excel(file_path, header=0)
    if not all(ref in df.columns for ref in REF_HEADERS):
        df = pd.read_excel(file_path, header=1)
        if not all(ref in df.columns for ref in REF_HEADERS):
            raise ValueError(f"Reference headers {REF_HEADERS} not found in {file_path} (checked first and second rows).")
    
    df_ref = df[REF_HEADERS].copy()
    df_data = df[DATA_HEADERS].copy()
    
    # Apply formatting functions.
    df_data["ADD"] = df_data["ADD"].apply(format_add)
    df_data["Free ROD"] = df_data["Free ROD"].apply(format_free_rod)
    # For "On Hand", use a normal conversion.
    df_data["On Hand"] = df_data["On Hand"].apply(lambda x: float(str(x).replace(",", "").strip()) if str(x).replace(",", "").strip().isdigit() else x)
    
    # Rename data columns with a file index suffix.
    df_data = df_data.rename(columns=lambda x: f"{x}_{file_index}")
    
    df_combined = pd.concat([df_ref, df_data], axis=1)
    creation_time = os.path.getctime(file_path)
    timestamp = datetime.fromtimestamp(creation_time).strftime('%d/%m/%Y')
    return df_combined, timestamp

def merge_dataframes(existing_df, new_df):
    if existing_df is None or existing_df.empty:
        return new_df
    else:
        return pd.merge(existing_df, new_df, on=REF_HEADERS, how='outer')

def filter_category_data(df, category):
    if category == "FSV":
        return df[df["Descr"].str.contains('FSV', case=True, na=False)]
    elif category == "SF_PUCK":
        return df[df["Descr"].str.contains('SF|PUCK', case=True, na=False)]
    else:
        return pd.DataFrame(columns=df.columns)

def write_df_custom(writer, sheet_name, df, file_timestamps):
    wb = writer.book
    ws = wb.create_sheet(title=sheet_name)
    
    header1, header2 = [], []
    for col in df.columns:
        if col in REF_HEADERS:
            header1.append("")
            header2.append(col)
        else:
            if "_" in col:
                base, idx = col.rsplit("_", 1)
                ts = file_timestamps.get(idx, "")
                header1.append(ts)
                header2.append(base)
            else:
                header1.append("")
                header2.append(col)
    ws.append(header1)
    ws.append(header2)
    for row in df.itertuples(index=False, name=None):
        ws.append(list(row))

def extract_category_data(general_df, category, header):
    df_cat = filter_category_data(general_df, category)
    cols = REF_HEADERS + [col for col in general_df.columns if col.startswith(f"{header}_")]
    return df_cat[cols]

# -------------------------------
# Tkinter UI Class
# -------------------------------
class ExcelProcessorUI:
    def __init__(self, master):
        self.master = master
        master.title("Excel Data Processor")
        self.input_files = []
        self.output_file = ""
        
        self.select_input_button = tk.Button(master, text="Select Input Files", command=self.select_input_files)
        self.select_input_button.grid(row=0, column=0, padx=10, pady=10, sticky="w")
        
        self.input_files_label = tk.Label(master, text="No input files selected", justify="left")
        self.input_files_label.grid(row=0, column=1, padx=10, pady=10, sticky="w")
        
        self.select_output_button = tk.Button(master, text="Select Output File", command=self.select_output_file)
        self.select_output_button.grid(row=1, column=0, padx=10, pady=10, sticky="w")
        
        self.output_file_label = tk.Label(master, text="No output file selected", justify="left")
        self.output_file_label.grid(row=1, column=1, padx=10, pady=10, sticky="w")
        
        self.process_button = tk.Button(master, text="Process Data", command=self.process_data)
        self.process_button.grid(row=2, column=0, padx=10, pady=10, sticky="w")
        
        self.log_text = tk.Text(master, height=10, width=70)
        self.log_text.grid(row=3, column=0, columnspan=2, padx=10, pady=10)
    
    def log(self, message):
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
    
    def select_input_files(self):
        files = filedialog.askopenfilenames(title="Select Excel Input Files", filetypes=[("Excel Files", "*.xlsx")])
        if files:
            self.input_files = list(files)
            self.input_files_label.config(text="\n".join(self.input_files))
            self.log(f"Selected {len(self.input_files)} input file(s).")
    
    def select_output_file(self):
        file = filedialog.asksaveasfilename(title="Select Output Excel File", defaultextension=".xlsx",
                                              filetypes=[("Excel Files", "*.xlsx")])
        if file:
            self.output_file = file
            self.output_file_label.config(text=self.output_file)
            self.log(f"Output file set to: {self.output_file}")
    
    def process_data(self):
        if not self.input_files:
            messagebox.showerror("Error", "No input files selected!")
            return
        if not self.output_file:
            messagebox.showerror("Error", "No output file selected!")
            return
        
        pickle_path = self.output_file + ".pkl"
        ts_pickle_path = self.output_file + ".timestamps.pkl"
        
        if os.path.exists(pickle_path):
            try:
                general_df = pd.read_pickle(pickle_path)
                self.log("Loaded existing merged data.")
            except Exception as e:
                self.log(f"Error loading existing data: {e}")
                general_df = None
        else:
            general_df = None
        
        if os.path.exists(ts_pickle_path):
            try:
                file_timestamps = pd.read_pickle(ts_pickle_path)
            except Exception as e:
                self.log(f"Error loading timestamps: {e}")
                file_timestamps = {}
        else:
            file_timestamps = {}
        
        if file_timestamps:
            start_index = max(int(k) for k in file_timestamps.keys()) + 1
        else:
            start_index = 1
        
        for i, file in enumerate(self.input_files, start=start_index):
            self.log(f"Processing file: {file}")
            try:
                new_df, timestamp = process_input_file(file, i)
                file_timestamps[str(i)] = timestamp
                general_df = merge_dataframes(general_df, new_df)
            except Exception as e:
                self.log(f"Error processing {file}: {e}")
                continue
        
        try:
            pd.to_pickle(general_df, pickle_path)
            pd.to_pickle(file_timestamps, ts_pickle_path)
            self.log("Updated persistent merged data and timestamps.")
        except Exception as e:
            self.log(f"Error saving persistent data: {e}")
        
        if general_df is not None:
            general_fsv = general_df[general_df["Descr"].str.contains('FSV', case=True, na=False)]
            general_sf_puck = general_df[general_df["Descr"].str.contains('SF|PUCK', case=True, na=False)]
        else:
            general_fsv = None
            general_sf_puck = None
        
        category_dfs = {"FSV": {}, "SF_PUCK": {}}
        for header in DATA_HEADERS:
            category_dfs["FSV"][header] = extract_category_data(general_df, "FSV", header)
            category_dfs["SF_PUCK"][header] = extract_category_data(general_df, "SF_PUCK", header)
        
        try:
            with pd.ExcelWriter(self.output_file, engine='openpyxl') as writer:
                if general_df is not None:
                    write_df_custom(writer, "General", general_df, file_timestamps)
                if general_fsv is not None and not general_fsv.empty:
                    write_df_custom(writer, "General_FSV", general_fsv, file_timestamps)
                if general_sf_puck is not None and not general_sf_puck.empty:
                    write_df_custom(writer, "General_SF_PUCK", general_sf_puck, file_timestamps)
                for category in ["FSV", "SF_PUCK"]:
                    for header in DATA_HEADERS:
                        sheet_name = CATEGORY_SHEETS[category][header]
                        df_to_write = category_dfs[category][header]
                        if df_to_write is not None and not df_to_write.empty:
                            write_df_custom(writer, sheet_name, df_to_write, file_timestamps)
                        else:
                            empty_df = pd.DataFrame(columns=REF_HEADERS + [header])
                            write_df_custom(writer, sheet_name, empty_df, file_timestamps)
            self.log(f"Data consolidation complete. Output written to '{self.output_file}'.")
            messagebox.showinfo("Success", f"Data processed and output saved to {self.output_file}")
        except Exception as ex:
            self.log(f"An error occurred during writing output: {ex}")
            messagebox.showerror("Error", f"An error occurred: {ex}")

# -------------------------------
# Main Entry Point
# -------------------------------
if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelProcessorUI(root)
    root.mainloop()
