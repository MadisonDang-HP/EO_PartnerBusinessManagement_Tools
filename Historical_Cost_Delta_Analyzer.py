# Standard Library
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime

# Third-party
import pandas as pd
try:
    import pyxlsb
    XLSB_SUPPORT = True
except ImportError:
    XLSB_SUPPORT = False

# Local
from Spec_Comparator import get_first_price_for_spec


def choose_file():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls *.xlsb")])
    return file_path

def read_excel_file(file_path):
    """Read Excel file with appropriate engine based on file extension"""
    file_ext = os.path.splitext(file_path)[1].lower()
    
    if file_ext == '.xlsb':
        if not XLSB_SUPPORT:
            raise ImportError("pyxlsb is required to read .xlsb files. Install it with: pip install pyxlsb")
        return pd.ExcelFile(file_path, engine='pyxlsb')
    else:
        # Default engine for .xlsx and .xls files
        return pd.ExcelFile(file_path)

def find_priority_sheet(xl):
    # Prioritize sheets with all keywords
    keywords = ["doc kit", "sku", "summary", "for hp"]
    for sheet in xl.sheet_names:
        if all(k in sheet.lower() for k in keywords):
            return sheet
    # Fallback: first sheet
    return xl.sheet_names[0]

def find_col(cols, keywords):
    for col in cols:
        if any(k in col.lower() for k in keywords):
            return col
    return None

def get_bom_history_rows(bom_df, part_number):
    # Find the row with the part number, then all rows below until "HP CM - ALL OS - BTO"
    idx = bom_df[bom_df.apply(lambda row: part_number in row.values, axis=1)].index
    if len(idx) == 0:
        return None
    start = idx[0]
    rows = []
    for i in range(start, len(bom_df)):
        rows.append(bom_df.iloc[i])
        if "HP CM - ALL OS - BTO" in str(bom_df.iloc[i].values):
            break
    return pd.DataFrame(rows, columns=bom_df.columns)

def find_bom_sheet(xl):
    keywords = ["doc kit", "sku", "summary", "for hp"]
    for sheet in xl.sheet_names:
        if all(k in sheet.lower() for k in keywords):
            return sheet
    # Fallback: first sheet
    return xl.sheet_names[0]

def find_price_sheet(xl):
    if len(xl.sheet_names) > 1:
        return xl.sheet_names[1]
    return xl.sheet_names[0]

class HistoricalCostDeltaAnalyzer:
    def __init__(self, root):
        self.root = root
        self.root.title("Historical Cost Delta Analyzer")
        self.root.geometry("800x600")
        self.root.configure(bg='#f0f0f0')
        
        # Variables
        self.file_path = tk.StringVar()
        self.status_text = tk.StringVar(value="Ready to analyze file...")
        self.progress_var = tk.DoubleVar()
        
        self.create_widgets()
        
    def create_widgets(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="Historical Cost Delta Analyzer", 
                               font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # File selection section
        file_frame = ttk.LabelFrame(main_frame, text="File Selection", padding="10")
        file_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        file_frame.columnconfigure(1, weight=1)
        
        ttk.Label(file_frame, text="Excel File:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        
        self.file_entry = ttk.Entry(file_frame, textvariable=self.file_path, state='readonly')
        self.file_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        
        self.browse_button = ttk.Button(file_frame, text="Browse", command=self.browse_file)
        self.browse_button.grid(row=0, column=2)
        
        # Supported formats info
        formats_text = "Supported: .xlsx, .xls, .xlsb (Binary Worksheet)"
        if not XLSB_SUPPORT:
            formats_text += " - Note: Install 'pyxlsb' for .xlsb support"
        ttk.Label(file_frame, text=formats_text, font=('Arial', 8), foreground='gray').grid(
            row=1, column=0, columnspan=3, sticky=tk.W, pady=(5, 0))
        
        # Analysis options section
        options_frame = ttk.LabelFrame(main_frame, text="Analysis Options", padding="10")
        options_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        self.include_bom_var = tk.BooleanVar(value=True)
        self.include_spec_var = tk.BooleanVar(value=True)
        
        ttk.Checkbutton(options_frame, text="Include BOM Variances", 
                       variable=self.include_bom_var).grid(row=0, column=0, sticky=tk.W, pady=5)
        
        ttk.Checkbutton(options_frame, text="Include Spec Variances", 
                       variable=self.include_spec_var).grid(row=1, column=0, sticky=tk.W, pady=5)
        
        # Output options section
        output_frame = ttk.LabelFrame(main_frame, text="Output Options", padding="10")
        output_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        self.auto_open_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(output_frame, text="Auto-open result file", 
                       variable=self.auto_open_var).grid(row=0, column=0, sticky=tk.W, pady=5)
        
        # Control buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=4, column=0, columnspan=3, pady=20)
        
        self.analyze_button = ttk.Button(button_frame, text="Analyze File", 
                                        command=self.analyze_file, style='Accent.TButton')
        self.analyze_button.pack(side=tk.LEFT, padx=(0, 10))
        
        self.clear_button = ttk.Button(button_frame, text="Clear", command=self.clear_fields)
        self.clear_button.pack(side=tk.LEFT, padx=(0, 10))
        
        self.exit_button = ttk.Button(button_frame, text="Exit", command=self.root.quit)
        self.exit_button.pack(side=tk.LEFT)
        
        # Add install pyxlsb button if not available
        if not XLSB_SUPPORT:
            self.install_button = ttk.Button(button_frame, text="Install .xlsb Support", 
                                           command=self.install_xlsb_support)
            self.install_button.pack(side=tk.LEFT, padx=(10, 0))
        
        # Progress bar
        self.progress_bar = ttk.Progressbar(main_frame, variable=self.progress_var, 
                                           maximum=100, length=400)
        self.progress_bar.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 5))
        
        # Status label
        self.status_label = ttk.Label(main_frame, textvariable=self.status_text, 
                                     foreground='blue')
        self.status_label.grid(row=6, column=0, columnspan=3, pady=(0, 10))
        
        # Results text area
        results_frame = ttk.LabelFrame(main_frame, text="Analysis Results", padding="10")
        results_frame.grid(row=7, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        results_frame.columnconfigure(0, weight=1)
        results_frame.rowconfigure(0, weight=1)
        main_frame.rowconfigure(7, weight=1)
        
        self.results_text = tk.Text(results_frame, wrap=tk.WORD, height=15, width=80)
        scrollbar = ttk.Scrollbar(results_frame, orient=tk.VERTICAL, command=self.results_text.yview)
        self.results_text.configure(yscrollcommand=scrollbar.set)
        
        self.results_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
    def browse_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls *.xlsb"), ("All files", "*.*")]
        )
        if file_path:
            self.file_path.set(file_path)
            self.update_status(f"Selected: {os.path.basename(file_path)}")
            
            # Check for .xlsb support
            if file_path.lower().endswith('.xlsb') and not XLSB_SUPPORT:
                messagebox.showwarning(
                    "Missing Dependency", 
                    "To read .xlsb files, you need to install pyxlsb.\n\n"
                    "Run this command in your terminal:\n"
                    "pip install pyxlsb"
                )
            
    def clear_fields(self):
        self.file_path.set("")
        self.results_text.delete(1.0, tk.END)
        self.progress_var.set(0)
        self.update_status("Ready to analyze file...")
        
    def update_status(self, message):
        self.status_text.set(message)
        self.root.update_idletasks()
        
    def log_result(self, message):
        self.results_text.insert(tk.END, message + "\n")
        self.results_text.see(tk.END)
        self.root.update_idletasks()
        
    def install_xlsb_support(self):
        """Show instructions for installing pyxlsb"""
        message = (
            "To enable support for Excel Binary Worksheet (.xlsb) files:\n\n"
            "1. Open your command prompt or terminal\n"
            "2. Run the following command:\n\n"
            "   pip install pyxlsb\n\n"
            "3. Restart this application\n\n"
            "After installation, you'll be able to process .xlsb files."
        )
        messagebox.showinfo("Install .xlsb Support", message)
        
    def analyze_file(self):
        if not self.file_path.get():
            messagebox.showerror("Error", "Please select an Excel file first.")
            return
            
        try:
            self.analyze_button.config(state='disabled')
            self.progress_var.set(0)
            self.results_text.delete(1.0, tk.END)
            
            file_path = self.file_path.get()
            self.update_status("Loading Excel file...")
            self.progress_var.set(10)
            
            # Use the appropriate engine based on file type
            try:
                xl = read_excel_file(file_path)
                file_ext = os.path.splitext(file_path)[1].lower()
                self.log_result(f"📁 File type: {file_ext}")
            except ImportError as e:
                if "pyxlsb" in str(e):
                    self.log_result("❌ Cannot read .xlsb file: pyxlsb not installed")
                    messagebox.showerror("Missing Dependency", 
                                       "To read .xlsb files, you need to install pyxlsb.\n\n"
                                       "Run this command in your terminal:\n"
                                       "pip install pyxlsb")
                    return
                else:
                    raise e
            
            bom_sheet = find_bom_sheet(xl)
            price_sheet = find_price_sheet(xl)
            
            self.log_result(f"📁 Analyzing file: {os.path.basename(file_path)}")
            self.log_result(f"📊 BOM Sheet: {bom_sheet}")
            self.log_result(f"💰 Price Sheet: {price_sheet}")
            self.log_result("")
            
            self.update_status("Reading price sheet...")
            self.progress_var.set(20)
            
            df = xl.parse(price_sheet)
            df.columns = [str(c).strip() for c in df.columns]
            
            # Find columns
            price_col = find_col(df.columns, ["orderable", "price", "cost", "pricing"])
            volume_col = find_col(df.columns, ["volume", "qty", "quantity", "moq"])
            variance_col = find_col(df.columns, ["variance", "delta"])
            remark_col = find_col(df.columns, ["remark", "comment"])
            spec_col = find_col(df.columns, ["spec", "specification"])
            part_col = find_col(df.columns, ["hppart", "item", "module", "part no", "part number", "sku"])
            
            self.progress_var.set(30)
            
            # Check for required columns
            required_cols = {
                "Price": price_col,
                "Volume": volume_col,
                "Variance": variance_col,
                "Remark": remark_col,
                "Spec": spec_col,
                "Part": part_col,
            }
            
            missing = [k for k, v in required_cols.items() if v is None]
            if missing:
                self.log_result(f"❌ Missing columns: {', '.join(missing)}")
                self.log_result(f"Available columns: {', '.join(df.columns)}")
                messagebox.showerror("Error", f"Could not find required columns: {', '.join(missing)}")
                return
            
            self.log_result("✅ All required columns found:")
            for name, col in required_cols.items():
                self.log_result(f"   {name}: {col}")
            self.log_result("")
            
            # Load BOM data if needed
            bom_df = None
            if self.include_bom_var.get():
                self.update_status("Loading BOM sheet...")
                bom_df = xl.parse(bom_sheet) if bom_sheet else None
                
            self.progress_var.set(40)
            
            # Process variances
            self.update_status("Processing variances...")
            bom_variances = []
            spec_variances = []
            
            total_rows = len(df)
            processed = 0
            
            for idx, row in df.iterrows():
                variance = row[variance_col]
                if pd.isna(variance) or variance <= 0:
                    processed += 1
                    continue
                    
                remark = str(row[remark_col]).lower() if remark_col else ""
                part_number = row[part_col] if part_col else None
                
                # BOM Variance
                if self.include_bom_var.get() and ("volume" in remark or "bom" in remark):
                    if bom_df is not None and part_number is not None:
                        bom_rows = get_bom_history_rows(bom_df, part_number)
                        if bom_rows is not None:
                            for _, bom_row in bom_rows.iterrows():
                                combined = pd.concat([row, bom_row], axis=0)
                                combined["Source Sheet"] = bom_sheet
                                combined["Source File"] = os.path.basename(file_path)
                                bom_variances.append(combined)
                                
                # Spec Variance
                elif self.include_spec_var.get() and str(variance).strip() and not ("volume" in remark or "bom" in remark):
                    combined = row.copy()
                    combined["Source Sheet"] = price_sheet
                    combined["Source File"] = os.path.basename(file_path)
                    
                    # Find spec price
                    folder = os.path.dirname(os.path.dirname(os.path.dirname(file_path)))
                    spec_folder = os.path.join(folder, "SPEC PRICING FILES")
                    if os.path.isdir(spec_folder):
                        volume, price = get_first_price_for_spec(row[spec_col], spec_folder)
                        combined["Spec Price"] = price
                        combined["Spec Price Volume"] = volume
                    else:
                        combined["Spec Price"] = None
                        combined["Spec Price Volume"] = None
                        
                    spec_variances.append(combined)
                
                processed += 1
                progress = 40 + (processed / total_rows) * 40
                self.progress_var.set(progress)
                
                if processed % 10 == 0:  # Update every 10 rows
                    self.root.update_idletasks()
            
            self.progress_var.set(80)
            
            # Generate output
            self.update_status("Generating output file...")
            today = datetime.now().strftime("%Y-%m-%d")
            out_path = os.path.join(os.path.dirname(file_path), f"Historical Cost Delta Analyzer {today}.xlsx")
            
            sheets_created = []
            with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
                if bom_variances:
                    pd.DataFrame(bom_variances).to_excel(writer, sheet_name="BOM Variances", index=False)
                    sheets_created.append("BOM Variances")
                if spec_variances:
                    pd.DataFrame(spec_variances).to_excel(writer, sheet_name="Spec Variances", index=False)
                    sheets_created.append("Spec Variances")
            
            self.progress_var.set(100)
            
            # Display results
            self.log_result("📊 ANALYSIS COMPLETE!")
            self.log_result(f"📄 Output file: {os.path.basename(out_path)}")
            self.log_result(f"📋 Sheets created: {', '.join(sheets_created)}")
            self.log_result(f"🔍 BOM variances found: {len(bom_variances)}")
            self.log_result(f"🔍 Spec variances found: {len(spec_variances)}")
            
            self.update_status("Analysis complete!")
            
            # Auto-open file if requested
            if self.auto_open_var.get():
                os.startfile(out_path)
                
            messagebox.showinfo("Success", f"Analysis complete!\n\nOutput saved to:\n{out_path}")
            
        except ImportError as e:
            if "pyxlsb" in str(e):
                self.log_result("❌ Cannot read .xlsb file: pyxlsb not installed")
                messagebox.showerror("Missing Dependency", 
                                   "To read .xlsb files, you need to install pyxlsb.\n\n"
                                   "Run this command in your terminal:\n"
                                   "pip install pyxlsb")
            else:
                self.log_result(f"❌ Import Error: {str(e)}")
                messagebox.showerror("Import Error", f"Missing required library:\n\n{str(e)}")
        except Exception as e:
            self.log_result(f"❌ Error: {str(e)}")
            messagebox.showerror("Error", f"An error occurred during analysis:\n\n{str(e)}")
            
        finally:
            self.analyze_button.config(state='normal')

def main():
    root = tk.Tk()
    app = HistoricalCostDeltaAnalyzer(root)
    root.mainloop()

if __name__ == "__main__":
    main()