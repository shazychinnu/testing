import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
from tkinter.filedialog import asksaveasfilename

class ExcelApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Data Viewer")
        self.root.geometry("800x600")
        
        # Placeholder for loaded data
        self.df = None
        
        # Frame for the dropdown and data display
        self.frame = tk.Frame(self.root)
        self.frame.pack(fill=tk.BOTH, expand=True)
        
        # Dropdown for sheet selection
        self.sheet_label = tk.Label(self.frame, text="Select Sheet:")
        self.sheet_label.pack(side=tk.TOP, padx=10, pady=5)
        
        self.sheet_dropdown = ttk.Combobox(self.frame, state="readonly")
        self.sheet_dropdown.pack(side=tk.TOP, padx=10, pady=5)
        self.sheet_dropdown.bind("<<ComboboxSelected>>", self.load_sheet_data)
        
        # Treeview widget for displaying the data
        self.tree = ttk.Treeview(self.frame, show="headings")
        self.tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Scrollbars for Treeview
        self.x_scrollbar = ttk.Scrollbar(self.frame, orient="horizontal", command=self.tree.xview)
        self.x_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        self.y_scrollbar = ttk.Scrollbar(self.frame, orient="vertical", command=self.tree.yview)
        self.y_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.tree.configure(xscrollcommand=self.x_scrollbar.set, yscrollcommand=self.y_scrollbar.set)
        
        # Load the Excel file
        self.load_excel_file_button = tk.Button(self.root, text="Load Excel File", command=self.load_excel_file)
        self.load_excel_file_button.pack(side=tk.BOTTOM, padx=10, pady=10)
        
        # Button to download data
        self.download_button = tk.Button(self.root, text="Download Data", command=self.download_data)
        self.download_button.pack(side=tk.BOTTOM, padx=10, pady=10)
    
    def load_excel_file(self):
        """Open file dialog to load an Excel file and populate sheet dropdown"""
        file_path = tk.filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if not file_path:
            return

        # Load the Excel file and get sheet names
        self.df = pd.read_excel(file_path, sheet_name=None)  # Load all sheets
        sheet_names = self.df.keys()  # Extract sheet names

        # Update dropdown menu with sheet names
        self.sheet_dropdown['values'] = list(sheet_names)
        if sheet_names:
            self.sheet_dropdown.current(0)  # Select the first sheet by default
            self.load_sheet_data()  # Load the first sheet's data automatically

    def load_sheet_data(self, event=None):
        """Load the data from the selected sheet into the Treeview"""
        sheet_name = self.sheet_dropdown.get()
        if sheet_name:
            data = self.df[sheet_name]
            
            # Clear existing data in the treeview
            for i in self.tree.get_children():
                self.tree.delete(i)

            # Set columns and headings
            self.tree["columns"] = list(data.columns)
            for col in data.columns:
                self.tree.heading(col, text=col)
                self.tree.column(col, width=100, anchor="center")
            
            # Insert data into treeview
            for row in data.itertuples(index=False):
                self.tree.insert("", "end", values=row)
    
    def download_data(self):
        """Allow user to download the displayed data as an Excel file"""
        if not self.df:
            messagebox.showerror("Error", "No data to download. Please load an Excel file first.")
            return
        
        # Ask for file save location
        save_path = asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if save_path:
            selected_sheet = self.sheet_dropdown.get()
            data_to_download = self.df[selected_sheet]
            data_to_download.to_excel(save_path, index=False)
            messagebox.showinfo("Success", f"Data saved to {save_path}")
        
# Running the Tkinter Application
if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelApp(root)
    root.mainloop()
