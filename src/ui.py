import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkcalendar import DateEntry
from datetime import datetime, timedelta
from report_generator import generate_excel_report
import sv_ttk

class MT5UI:
    def __init__(self, root):
        self.root = root
        self.root.title("MT5 Trade Report Generator")
        self.root.geometry("400x350")
        self.root.resizable(False, False)
        
        sv_ttk.set_theme("light")
        
        self.mainframe = ttk.Frame(root, padding="20")
        self.mainframe.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(self.mainframe, text="MT5 Trade Report Generator", 
                 font=('Helvetica', 14, 'bold')).pack(pady=(0, 20))
        
        date_frame = ttk.LabelFrame(self.mainframe, text="Select Date Range", padding=10)
        date_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(date_frame, text="Start Date:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.start_date = DateEntry(date_frame, width=20, background='darkblue', 
                                  foreground='white', borderwidth=2, date_pattern='y-mm-dd')
        self.start_date.grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(date_frame, text="End Date:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.end_date = DateEntry(date_frame, width=20, background='darkblue', 
                                foreground='white', borderwidth=2, date_pattern='y-mm-dd')
        self.end_date.grid(row=1, column=1, padx=5, pady=5)
        
        button_frame = ttk.Frame(self.mainframe)
        button_frame.pack(pady=20)
        
        self.export_btn = ttk.Button(button_frame, text="Generate Report", 
                                   command=self.generate_report, style='Accent.TButton')
        self.export_btn.pack(pady=10, ipadx=10, ipady=5)
        
        self.status = ttk.Label(self.mainframe, text="Ready", relief=tk.SUNKEN, 
                              anchor=tk.W, padding=(5, 2))
        self.status.pack(fill=tk.X, pady=(10, 0))
        
        self.set_default_dates()
    
    def set_default_dates(self):
        today = datetime.now()
        first_day_current = today.replace(day=1)

        self.start_date.set_date(first_day_current)
        self.end_date.set_date(today)
    
    def generate_report(self):
        # try:
            start = datetime.strptime(self.start_date.get(), "%Y-%m-%d")
            end = datetime.strptime(self.end_date.get(), "%Y-%m-%d")

            if end < start:
                raise ValueError("End date must be after start date")

            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx", 
                filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")],
                title="Save Report As",
                initialfile=f"MT5_Report_{start.date()}_to_{end.date()}.xlsx"
            )
            
            if not file_path:
                return

            self.status.config(text="Generating report... Please wait")
            self.root.update()
            
            generate_excel_report(start, end, file_path)
            self.status.config(text=f"Report saved successfully to: {file_path}")
            
            if messagebox.askyesno("Success", "Report generated successfully. Would you like to open the file now?"):
                import os
                os.startfile(file_path)
                
        # except Exception as e:
        #     messagebox.showerror("Error", f"Failed to generate report:\n{str(e)}")
        #     self.status.config(text="Error generating report")
