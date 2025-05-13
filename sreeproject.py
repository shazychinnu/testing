import tkinter as tk
from tkinter import messagebox
import threading
import time
import os

def execute_macro(name):
    for i in range(5):
        yield f"{name} Step {i+1} completed"
        time.sleep(0.1)

LOG_DIR = "macro_logs"
os.makedirs(LOG_DIR, exist_ok=True)

def run_macros(macros, update_ui, finish_callback, log_callback):
    for i, macro in enumerate(macros):
        logs = [f"{macro} - Started"]
        update_ui(i, "Running", animate=True)
        for log in execute_macro(macro):
            logs.append(log)
            log_callback(macro, log)
        logs.append(f"{macro} - Completed")
        log_callback(macro, f"{macro} - Completed")
        with open(os.path.join(LOG_DIR, f"{macro}.log"), "w") as f:
            for entry in logs:
                f.write(entry + "\n")
        update_ui(i, "Completed", animate=False)
    finish_callback()

class MacroToolApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Macro Automation Dashboard")
        self.root.geometry("900x550")
        self.root.resizable(False, False)

        self.theme = {
            "bg": "#1E1E2F",
            "fg": "#FFFFFF",
            "entry_bg": "#2D2D44",
            "box_bg": "#2D2D44",
            "box_fg": "#FFFFFF",
            "success": "#10B981",   # Green
            "error": "#EF4444",     # Red
            "running": "#F59E0B",   # Amber
            "button_bg": "#3B82F6",  # Blue
            "button_hover": "#2563EB"
        }

        self.font = ("Segoe UI", 10)
        self.header_font = ("Segoe UI", 12, "bold")
        self.log_font = ("Consolas", 9)
        self.status_lines = []
        self.macro_logs = {}

        self.setup_ui()

    def setup_ui(self):
        self.main_frame = tk.Frame(self.root, bg=self.theme["bg"])
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        self.paned = tk.PanedWindow(self.main_frame, orient=tk.HORIZONTAL, bg=self.theme["bg"])
        self.paned.pack(fill=tk.BOTH, expand=True)

        self.left_panel = tk.Frame(self.paned, width=600, bg=self.theme["bg"])
        self.right_panel = tk.Frame(self.paned, width=300, bg=self.theme["bg"])
        self.paned.add(self.left_panel)
        self.paned.add(self.right_panel)

        self.build_left_panel()
        self.build_right_panel()

    def build_left_panel(self):
        tk.Label(self.left_panel, text="📊 Macro Tracker", font=self.header_font,
                 bg=self.theme["bg"], fg=self.theme["fg"]).pack(pady=(10, 5))

        login_frame = tk.Frame(self.left_panel, bg=self.theme["bg"])
        login_frame.pack(pady=5)
        tk.Label(login_frame, text="Username", bg=self.theme["bg"], fg=self.theme["fg"], font=self.font).grid(row=0, column=0, padx=5)
        self.username = tk.Entry(login_frame, bg=self.theme["entry_bg"], fg=self.theme["fg"],
                                 insertbackground=self.theme["fg"], font=self.font, relief=tk.FLAT)
        self.username.grid(row=0, column=1)
        tk.Label(login_frame, text="Password", bg=self.theme["bg"], fg=self.theme["fg"], font=self.font).grid(row=0, column=2, padx=5)
        self.password = tk.Entry(login_frame, show="*", bg=self.theme["entry_bg"], fg=self.theme["fg"],
                                 insertbackground=self.theme["fg"], font=self.font, relief=tk.FLAT)
        self.password.grid(row=0, column=3)

        control_frame = tk.Frame(self.left_panel, bg=self.theme["bg"])
        control_frame.pack(pady=10)

        self.start_btn = self.make_button(control_frame, "▶ Start", self.start_macros)
        self.start_btn.pack(side=tk.LEFT, padx=8)

        self.clear_btn = self.make_button(control_frame, "🧹 Clear", self.clear_logs)
        self.clear_btn.pack(side=tk.LEFT, padx=8)

        self.refresh_btn = self.make_button(control_frame, "🔄 Refresh", self.refresh_status)
        self.refresh_btn.pack(side=tk.LEFT, padx=8)

        canvas = tk.Canvas(self.left_panel, bg=self.theme["bg"], highlightthickness=0)
        scrollbar = tk.Scrollbar(self.left_panel, orient="vertical", command=canvas.yview)
        self.table_frame = tk.Frame(canvas, bg=self.theme["bg"])

        self.table_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.table_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(10, 0))
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.status_label = tk.Label(self.left_panel, text="", font=self.font,
                                     bg=self.theme["bg"], fg=self.theme["fg"])
        self.status_label.pack(pady=(5, 10))

    def build_right_panel(self):
        tk.Label(self.right_panel, text="Macro Logs", font=self.header_font,
                 bg=self.theme["bg"], fg=self.theme["fg"]).pack(pady=25)

        log_frame = tk.Frame(self.right_panel, bg=self.theme["bg"])
        log_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        scrollbar = tk.Scrollbar(log_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.log_view = tk.Text(log_frame, bg=self.theme["box_bg"], fg=self.theme["box_fg"],
                                insertbackground=self.theme["box_fg"], wrap="word", font=self.log_font,
                                yscrollcommand=scrollbar.set, relief=tk.FLAT)
        self.log_view.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.log_view.yview)
        self.log_view.config(state=tk.DISABLED)

    def make_button(self, parent, text, command):
        btn = tk.Button(parent, text=text, command=command,
                        font=self.font,
                        bg=self.theme["button_bg"],
                        fg="#FFFFFF",
                        activebackground=self.theme["button_hover"],
                        activeforeground="#FFFFFF",
                        relief=tk.FLAT,
                        bd=0,
                        width=10,
                        height=1,
                        cursor="hand2")
        return btn

    def start_macros(self):
        user = self.username.get()
        pwd = self.password.get()
        if not user or not pwd:
            messagebox.showerror("Error", "Username and Password required")
            return

        macros = [f"Macro_{i+1}.xlsm" for i in range(25)]
        self.status_lines = [[m, "Pending"] for m in macros]
        self.render_table()

        self.log_view.config(state=tk.NORMAL)
        self.log_view.delete("1.0", tk.END)
        self.log_view.config(state=tk.DISABLED)

        self.start_btn.config(state=tk.DISABLED)
        threading.Thread(target=run_macros, args=(macros, self.update_status, self.finish, self.add_log), daemon=True).start()

    def update_status(self, idx, status, animate=False):
        self.root.after(0, lambda: self._set_status(idx, status))
        if animate:
            def dots():
                for _ in range(3):
                    for dot in [".", "..", "..."]:
                        if "Running" not in self.status_lines[idx][1]:
                            return
                        self.status_lines[idx][1] = f"Running{dot}"
                        self.root.after(0, self.render_table)
                        time.sleep(0.3)
            threading.Thread(target=dots, daemon=True).start()

    def _set_status(self, idx, status):
        self.status_lines[idx][1] = status
        self.render_table()

    def render_table(self):
        for w in self.table_frame.winfo_children():
            w.destroy()
        for i, (macro, status) in enumerate(self.status_lines):
            tk.Label(self.table_frame, text=macro, width=25, anchor="w", font=self.font,
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
        self.render_table()
        self.log_view.config(state=tk.NORMAL)
        self.log_view.delete("1.0", tk.END)
        self.log_view.config(state=tk.DISABLED)
        self.status_label.config(text="")

    def refresh_status(self):
        for i in range(len(self.status_lines)):
            self.status_lines[i][1] = "Pending"
        self.render_table()
        self.status_label.config(text="Table refreshed.")

    def finish(self):
        # self.status_label.config(text="✅ All macros completed.")
        self.status_label.config(text="")
        self.start_btn.config(state=tk.NORMAL)

if __name__ == "__main__":
    root = tk.Tk()
    app = MacroToolApp(root)
    root.mainloop()
