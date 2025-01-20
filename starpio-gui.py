import tkinter as tk
from tkinter import BooleanVar, filedialog, ttk
from threading import Thread

class Application(tk.Tk):
    def __init__(self):
        super().__init__()

        # Configure the main window
        self.title("Statpro Analyzing Tool")
        self.geometry("550x260")
        self.config(bg="#F5F5F5")  # Light mode by default
        self.message_label = None
        self.frame1_buttons = []

        # Dark mode variable
        self.dark_mode = BooleanVar(value=False)

        # Create and configure frames
        self.create_menu()
        self.create_frames()

        # Initialize the home screen
        self.show_home()

    def create_menu(self):
        # Add menu bar
        self.menu_bar = tk.Menu(self)
        self.config(menu=self.menu_bar)

        # Add Home option to menu bar
        self.menu_bar.add_command(label="Home", command=self.show_home)

        # Add Help option to menu bar
        self.menu_bar.add_command(label="Help", command=self.show_help)

    def create_frames(self):
        # Frame 1
        self.frame1 = tk.Frame(self, bg="#D6E4F0", relief="groove", borderwidth=2, width=150)
        self.frame1.grid(row=0, column=0, sticky="nsew")

        # Frame 2
        self.frame2 = tk.Frame(self, bg="#FFFFFF", relief="groove", borderwidth=2)
        self.frame2.grid(row=0, column=1, sticky="nsew")

        # Configure layout
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)
        self.columnconfigure(1, weight=3)

        # Add buttons to Frame 1
        self.securities = ["Security1", "Security2", "Security3", "Security4"]
        for i in self.securities:
            button = tk.Button(
                self.frame1, text=i, bg="#B3D4EA", fg="#000000", activebackground="#A9C4D5",
                activeforeground="#000000", command=lambda i=i: self.update_frame2(i), width=20,
                height=1
            )
            button.pack(pady=5)
            self.frame1_buttons.append(button)

        update_sec_master_button = tk.Button(
            self.frame1, text="Update security master", bg="#B3D4EA", fg="#000000", activebackground="#A9C4D5",
            activeforeground="#000000", command=self.update_security_master, width=20, height=1
        )
        update_sec_master_button.pack(pady=10)
        self.frame1_buttons.append(update_sec_master_button)

        update_unclassified_master_button = tk.Button(
            self.frame1, text="Update unclassified master", bg="#B3D4EA", fg="#000000", activebackground="#A9C4D5",
            activeforeground="#000000", command=self.update_unclassified_master, width=20, height=1
        )
        update_unclassified_master_button.pack(pady=10)
        self.frame1_buttons.append(update_unclassified_master_button)

    def toggle_frame1_buttons(self, state):
        """Enable or disable buttons in Frame 1."""
        for button in self.frame1_buttons:
            button.config(state=state)

    def show_home(self):
        self.clear_frame2()
        fg_color = "#E0E0E0" if self.dark_mode.get() else "#000000"
        bg_color = self.frame2.cget("bg")

        tk.Label(self.frame2, text="Welcome to the Home Screen!", bg=bg_color, fg=fg_color, font=("Arial", 14)).pack(pady=10)

    def show_help(self):
        self.clear_frame2()
        fg_color = "#E0E0E0" if self.dark_mode.get() else "#000000"
        bg_color = self.frame2.cget("bg")

        help_text = """
        Help Section:
        - Button 1: Upload files and submit them.
        - Button 2: Change settings such as password and profile.
        - Button 3: View reports or export data.
        - Button 4: Access help or contact support.

        For further assistance, email support@example.com.
        """
        tk.Label(self.frame2, text=help_text, bg=bg_color, fg=fg_color, justify="left", wraplength=400).pack(pady=10)

    def update_frame2(self, button_id):
        self.clear_frame2()
        fg_color = "#E0E0E0" if self.dark_mode.get() else "#000000"
        bg_color = self.frame2.cget("bg")

        tk.Label(self.frame2, text=f"{button_id.upper()}", bg=bg_color, fg=fg_color, 
                font=("Arial", 14),).pack(pady=10)

        input_frame = tk.Frame(self.frame2, bg=bg_color)
        input_frame.pack(pady=5)

        # Entry widget (Text box)
        file_entry = tk.Entry(input_frame, width=50)
        file_entry.grid(row=0, column=0, padx=(0, 10), pady=5)

        # Browse button
        def browse_file():
            file_path = filedialog.askopenfilename()
            if file_path:
                file_entry.delete(0, tk.END)
                file_entry.insert(0, file_path)

        tk.Button(input_frame, text="Browse", bg="#607D8B", fg="#FFFFFF", command=browse_file).grid(row=0, column=1, pady=5)

        # Submit button
        def submit_file():
            if file_entry.get():  # Check if the text box has content
                file_path = file_entry.get()
                self.save_file_path(file_path)
                file_entry.delete(0, tk.END)  # Clear the text box
            else:
                tk.Label(self.frame2, text="Please select a file before submitting.", bg=bg_color, 
                        fg="#FF0000",).pack(pady=5)

        tk.Button(self.frame2, text="Submit", bg="#4CAF50", fg="#FFFFFF", 
                command=lambda: Thread(target=submit_file).start(),).pack(pady=10)

        # Backend Simulation for Security Buttons
        if button_id == "Security1":
            self.backend_security1()
        elif button_id == "Security2":
            self.backend_security2()
        elif button_id == "Security3":
            self.backend_security3()
        elif button_id == "Security4":
            self.backend_security4()

    def backend_security1(self):
        # Simulate a backend task for Security1
        print("Processing Security1...")
        # self.run_with_loading_bar("Processing Security1 data", "Security1")

    def backend_security2(self):
        # Simulate a backend task for Security2
        print("Processing Security2...")
        # self.run_with_loading_bar("Processing Security2 data", "Security2")

    def backend_security3(self):
        # Simulate a backend task for Security3
        print("Processing Security3...")
        # self.run_with_loading_bar("Processing Security3 data", "Security3")

    def backend_security4(self):
        # Simulate a backend task for Security4
        print("Processing Security4...")
        # self.run_with_loading_bar("Processing Security4 data", "Security4")

    def update_security_master(self):
        self.run_with_loading_bar("Updating 'Security' master data", "Security")

    def update_unclassified_master(self):
        self.run_with_loading_bar("Updating 'Unclassified' master data", "Unclassified")

    def run_with_loading_bar(self, loading_message, success_message):
        """Display a loading bar and disable buttons in Frame 1 during the update."""
        self.clear_frame2()
        self.toggle_frame1_buttons(tk.DISABLED)

        bg_color = self.frame2.cget("bg")
        self.message_label = tk.Label(self.frame2, text=loading_message, bg=bg_color, fg="blue", font=("Arial", 10))
        self.message_label.pack(pady=10)

        progress = ttk.Progressbar(self.frame2, orient="horizontal", length=300, mode="indeterminate")
        progress.pack(pady=30)
        progress.start()

        # Simulate a delay and then finish the update
        self.after(2000, lambda: self.finish_update(progress, bg_color, success_message))

    def finish_update(self, progress, bg_color, success_message):
        progress.stop()
        progress.destroy()
        self.message_label.config(text="", fg="green")
        tk.Label(
            self.frame2,
            text=f"{success_message} Master data updated successfully!",
            bg=bg_color,
            fg="green",
            font=("Arial", 10),
        ).pack(pady=10)
        self.toggle_frame1_buttons(tk.NORMAL)

    def clear_frame2(self):
        for widget in self.frame2.winfo_children():
            widget.destroy()

    def save_file_path(self, file_path):
        save_path = filedialog.asksaveasfilename(
            initialfile="output.txt", defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        if save_path:
            with open(save_path, 'w') as file:
                file.write(file_path)

if __name__ == "__main__":
    app = Application()
    app.mainloop()
#######################################
    # db_module.py
import sqlite3
from typing import List, Tuple

def initialize_database():
    """Initialize the database with required tables."""
    conn = sqlite3.connect("master_data.db")
    cursor = conn.cursor()

    # Create security_master table
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS security_master (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            security_name TEXT NOT NULL,
            security_details TEXT
        )
        """
    )

    # Create unclassified_master table
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS unclassified_master (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            category_name TEXT NOT NULL,
            description TEXT
        )
        """
    )

    conn.commit()
    conn.close()

def update_security_master(data: List[Tuple[str, str]]):
    """Update the security_master table with new data.

    Args:
        data (List[Tuple[str, str]]): List of tuples containing security_name and security_details.
    """
    conn = sqlite3.connect("master_data.db")
    cursor = conn.cursor()

    # Insert or replace data into the table
    cursor.executemany(
        """
        INSERT INTO security_master (security_name, security_details)
        VALUES (?, ?)
        """,
        data
    )

    conn.commit()
    conn.close()

def update_unclassified_master(data: List[Tuple[str, str]]):
    """Update the unclassified_master table with new data.

    Args:
        data (List[Tuple[str, str]]): List of tuples containing category_name and description.
    """
    conn = sqlite3.connect("master_data.db")
    cursor = conn.cursor()

    # Insert or replace data into the table
    cursor.executemany(
        """
        INSERT INTO unclassified_master (category_name, description)
        VALUES (?, ?)
        """,
        data
    )

    conn.commit()
    conn.close()

def fetch_table_data(table_name: str) -> List[Tuple]:
    """Fetch all data from a specified table.

    Args:
        table_name (str): The name of the table to fetch data from.

    Returns:
        List[Tuple]: List of tuples containing the table data.
    """
    conn = sqlite3.connect("master_data.db")
    cursor = conn.cursor()

    cursor.execute(f"SELECT * FROM {table_name}")
    data = cursor.fetchall()

    conn.close()
    return data

# Initialize the database when this script is executed directly
if __name__ == "__main__":
    initialize_database()
