import pandas as pd


class ExcelBackend:
    def __init__(self):
        self.df = None  # Placeholder for the Excel data

    def load_excel_file(self, file_path):
        """Load an Excel file and return sheet names."""
        if not file_path:
            return None

        self.df = pd.read_excel(file_path, sheet_name=None)  # Load all sheets
        return list(self.df.keys()) if self.df else None

    def get_sheet_data(self, sheet_name):
        """Retrieve data for a specific sheet."""
        if self.df and sheet_name in self.df:
            return self.df[sheet_name]
        return None

    def download_data(self, sheet_name, save_path):
        """Save the selected sheet's data to a file."""
        if self.df and sheet_name in self.df:
            data_to_save = self.df[sheet_name]
            data_to_save.to_excel(save_path, index=False)
            return True
        return False
==============================================================
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from sample import ExcelBackend  # Import backend class


class ExcelApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Data Viewer")
        self.root.geometry("800x600")

        # Backend instance
        self.backend = ExcelBackend()

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
        """Open file dialog to load an Excel file and populate sheet dropdown."""
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if not file_path:
            return

        # Load Excel file via backend and get sheet names
        sheet_names = self.backend.load_excel_file(file_path)

        if sheet_names:
            # Update dropdown menu with sheet names
            self.sheet_dropdown['values'] = sheet_names
            self.sheet_dropdown.current(0)  # Select the first sheet by default
            self.load_sheet_data()  # Load the first sheet's data automatically

    def load_sheet_data(self, event=None):
        """Load the data from the selected sheet into the Treeview."""
        sheet_name = self.sheet_dropdown.get()
        data = self.backend.get_sheet_data(sheet_name)

        if data is not None:
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
        """Allow user to download the displayed data as an Excel file."""
        sheet_name = self.sheet_dropdown.get()
        if not sheet_name:
            messagebox.showerror("Error", "No sheet selected. Please load an Excel file first.")
            return

        # Ask for file save location
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if save_path:
            success = self.backend.download_data(sheet_name, save_path)
            if success:
                messagebox.showinfo("Success", f"Data saved to {save_path}")
            else:
                messagebox.showerror("Error", "Failed to save data.")


# Running the Tkinter Application
if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelApp(root)
    root.mainloop()
