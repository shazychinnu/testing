import tkinter as tk
from tkinter import messagebox, filedialog
import threading
import os
from backend import perform_operation

class MacroToolApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Macro Automation Dashboard")
        self.root.geometry("750x480")
        self.root.resizable(False, False)

        self.light_theme = {
            "bg": "#F9FAFB",
            "fg": "#111827",
            "entry_bg": "#FFFFFF",
            "box_bg": "#F3F4F6",
            "box_fg": "#111827",
            "header_fg": "#111827",
            "success": "#22C55E",
            "error": "#EF4444",
            "running": "#FBBF24",
            "start_btn": "#4ADE80",
            "clear_btn": "#F87171",
            "refresh_btn": "#60A5FA",
            "theme_btn": "#FBBF24",
            "button_hover": "#4B5563"
        }

        self.dark_theme = {
            "bg": "#1E1E2E",
            "fg": "#E5E7EB",
            "entry_bg": "#2A2A3B",
            "box_bg": "#2D2D44",
            "box_fg": "#E5E7EB",
            "header_fg": "#FFFFFF",
            "success": "#10B981",
            "error": "#EF4444",
            "running": "#F59E0B",
            "start_btn": "#22C55E",
            "clear_btn": "#EF4444",
            "refresh_btn": "#3B82F6",
            "theme_btn": "#F59E0B",
            "button_hover": "#2563EB"
        }

        self.theme = self.dark_theme
        self.font = ("Segoe UI", 10)
        self.header_font = ("Segoe UI", 13, "bold")
        self.log_font = ("Consolas", 9)
        self.status_lines = []
        self.status_dict = {}
        self.directory = ""
        self.lock = threading.Lock()

        self.setup_ui()

    def setup_ui(self):
        self.main_frame = tk.Frame(self.root)
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        self.paned = tk.PanedWindow(self.main_frame, orient=tk.HORIZONTAL)
        self.paned.pack(fill=tk.BOTH, expand=True)

        self.left_panel = tk.Frame(self.paned)
        self.right_panel = tk.Frame(self.paned)
        self.paned.add(self.left_panel)
        self.paned.add(self.right_panel)

        self.build_left_panel()
        self.build_right_panel()
        self.apply_theme()

    def build_left_panel(self):
        self.title_label = tk.Label(self.left_panel, text="📊 Macro Tracker", font=self.header_font)
        self.title_label.pack(pady=(10, 5))

        login_frame = tk.Frame(self.left_panel)
        login_frame.pack(pady=5)
        self.login_frame = login_frame
        tk.Label(login_frame, text="Username", font=self.font).grid(row=0, column=0, padx=5)
        self.username = tk.Entry(login_frame, font=self.font)
        self.username.grid(row=0, column=1)
        tk.Label(login_frame, text="Password", font=self.font).grid(row=0, column=2, padx=5)
        self.password = tk.Entry(login_frame, show="*", font=self.font)
        self.password.grid(row=0, column=3)

        control_frame = tk.Frame(self.left_panel)
        control_frame.pack(pady=10)
        self.control_frame = control_frame
        self.start_btn = self.make_button(control_frame, "▶ Start", self.start_macros, "start_btn")
        self.start_btn.pack(side=tk.LEFT, padx=5)
        self.clear_btn = self.make_button(control_frame, "🧹 Clear", self.clear_logs, "clear_btn")
        self.clear_btn.pack(side=tk.LEFT, padx=5)
        self.refresh_btn = self.make_button(control_frame, "🔄 Refresh", self.refresh_status, "refresh_btn")
        self.refresh_btn.pack(side=tk.LEFT, padx=5)
        self.theme_btn = self.make_button(control_frame, "🌚 Theme", self.toggle_theme, "theme_btn")
        self.theme_btn.pack(side=tk.LEFT, padx=5)

        canvas = tk.Canvas(self.left_panel, highlightthickness=0)
        self.table_canvas = canvas
        scrollbar = tk.Scrollbar(self.left_panel, orient="vertical", command=canvas.yview)
        self.table_frame = tk.Frame(canvas)

        self.table_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.table_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(10, 0))
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.status_label = tk.Label(self.left_panel, font=self.font)
        self.status_label.pack(pady=(5, 10))

    def build_right_panel(self):
        self.log_title = tk.Label(self.right_panel, text="Macro Logs", font=self.header_font)
        self.log_title.pack(pady=20)

        log_frame = tk.Frame(self.right_panel)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.log_frame = log_frame
        scrollbar = tk.Scrollbar(log_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.log_view = tk.Text(log_frame, wrap="word", font=self.log_font, yscrollcommand=scrollbar.set)
        self.log_view.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.log_view.yview)
        self.log_view.config(state=tk.DISABLED)

    def apply_theme(self):
        widgets = [self.root, self.main_frame, self.left_panel, self.right_panel,
                   self.table_frame, self.login_frame, self.control_frame,
                   self.status_label, self.log_frame, self.table_canvas]
        for widget in widgets:
            widget.configure(bg=self.theme["bg"])

        for label in self.login_frame.winfo_children():
            if isinstance(label, tk.Label):
                label.configure(bg=self.theme["bg"], fg=self.theme["fg"])

        for entry in [self.username, self.password]:
            entry.configure(bg=self.theme["entry_bg"], fg=self.theme["fg"], insertbackground=self.theme["fg"])

        self.title_label.configure(bg=self.theme["bg"], fg=self.theme["header_fg"])
        self.log_title.configure(bg=self.theme["bg"], fg=self.theme["header_fg"])
        self.log_view.configure(bg=self.theme["box_bg"], fg=self.theme["box_fg"])

        self.start_btn.configure(bg=self.theme["start_btn"], activebackground=self.theme["button_hover"])
        self.clear_btn.configure(bg=self.theme["clear_btn"], activebackground=self.theme["button_hover"])
        self.refresh_btn.configure(bg=self.theme["refresh_btn"], activebackground=self.theme["button_hover"])
        self.theme_btn.configure(bg=self.theme["theme_btn"], activebackground=self.theme["button_hover"])

        self.render_table()

    def toggle_theme(self):
        self.theme = self.light_theme if self.theme == self.dark_theme else self.dark_theme
        self.apply_theme()

    def make_button(self, parent, text, command, theme_key):
        return tk.Button(parent, text=text, command=command, font=self.font,
                         bg=self.theme[theme_key], fg="#FFFFFF",
                         activebackground=self.theme["button_hover"], relief=tk.FLAT, bd=0, width=10)

    def start_macros(self):
        user = self.username.get()
        pwd = self.password.get()
        if not user or not pwd:
            messagebox.showerror("Error", "Username and Password required")
            return

        directory = filedialog.askdirectory()
        if not directory:
            return

        self.directory = directory
        macro_files = [f for f in os.listdir(directory) if f.lower().endswith((".xlsm", ".pdf", ".docx"))]
        if not macro_files:
            messagebox.showinfo("Info", "No .xlsm files found in the selected directory.")
            return

        display_names = [os.path.splitext(f)[0] for f in macro_files]
        self.status_dict = {name: "Pending" for name in display_names}
        self.status_lines = [[name, "Pending"] for name in display_names]
        self.render_table()
        self.clear_logs()
        self.start_btn.config(state=tk.DISABLED)

        for original_file, display_name in zip(macro_files, display_names):
            full_path = os.path.join(directory, original_file)
            threading.Thread(target=self.run_macro_thread, args=(display_name, full_path), daemon=True).start()

    def run_macro_thread(self, macro_name, file_path):
        self.update_status(macro_name, "Running...")
        perform_operation(file_path, lambda msg: self.add_log(macro_name, msg))
        self.update_status(macro_name, "Completed")
        if all(status == "Completed" for status in self.status_dict.values()):
            self.finish()

    def update_status(self, macro_name, status):
        with self.lock:
            self.status_dict[macro_name] = status
            self.status_lines = [[k, self.status_dict[k]] for k in self.status_dict]
        self.root.after(0, self.render_table)

    def render_table(self):
        for w in self.table_frame.winfo_children():
            w.destroy()
        for i, (macro, status) in enumerate(self.status_lines):
            tk.Label(self.table_frame, text=macro, width=30, anchor="w", font=self.font,
                     bg=self.theme["entry_bg"], fg=self.theme["fg"]).grid(row=i, column=0, padx=2, pady=2)
            color = self.theme["success"] if "Completed" in status else self.theme["running"] if "Running" in status else self.theme["fg"]
            tk.Label(self.table_frame, text=status, width=15, anchor="w", font=self.font,
                     bg=self.theme["entry_bg"], fg=color).grid(row=i, column=1, padx=2, pady=2)

    def add_log(self, macro, msg):
        self.root.after(0, lambda: self.show_log(macro, msg))

    def show_log(self, macro, msg):
        self.log_view.config(state=tk.NORMAL)
        self.log_view.insert(tk.END, f"{macro}: {msg}\n")
        self.log_view.see(tk.END)
        self.log_view.config(state=tk.DISABLED)

    def clear_logs(self):
        self.status_lines.clear()
        self.status_dict.clear()
        self.render_table()
        self.log_view.config(state=tk.NORMAL)
        self.log_view.delete("1.0", tk.END)
        self.log_view.config(state=tk.DISABLED)
        self.status_label.config(text="")

    def refresh_status(self):
        for key in self.status_dict:
            self.status_dict[key] = "Pending"
        self.status_lines = [[k, "Pending"] for k in self.status_dict]
        self.render_table()
        self.status_label.config(text="Table refreshed.")

    def finish(self):
        self.start_btn.config(state=tk.NORMAL)
        toast = tk.Toplevel(self.root)
        toast.overrideredirect(True)
        toast.config(bg="#ECFDF5")
        self.root.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - 120
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - 30
        toast.geometry(f"240x60+{x}+{y}")
        tk.Label(toast, text="✅ All macros completed!", bg="#ECFDF5", fg="#047857",
                 font=("Segoe UI", 10, "bold")).pack(expand=True)
        toast.after(3000, toast.destroy)

if __name__ == "__main__":
    root = tk.Tk()
    app = MacroToolApp(root)
    root.mainloop()
