import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import threading
from backend import perform_operation

# Function to browse directory
def browse_directory(username, password, log_area, path_label):
    if username.strip() == "" or username.strip().lower() == "username" or password.strip() == "" or password.strip().lower() == "password":
        messagebox.showerror("Error", "Username or Password cannot be empty!")
        return

    directory = filedialog.askdirectory()
    if directory:
        display_path(directory, path_label)
        threading.Thread(target=perform_operation, args=(directory, lambda msg: update_log(log_area, msg))).start()

# Function to update log widget
def update_log(log_widget, message):
    log_widget.insert(tk.END, message)
    log_widget.see(tk.END)

# Function to display path properly
def display_path(directory, path_label):
    formatted_path = "\n".join([directory[i:i+50] for i in range(0, len(directory), 50)])
    path_label.config(text=f"Selected Path:\n{formatted_path}")

# Main window setup
def main():
    window = tk.Tk()
    window.title("Login and Logs")
    window.geometry("700x450")
    window.config(bg="#0F172A")
    window.resizable(False, False)

    # Top frame for path display
    top_frame = tk.Frame(window, bg="#0F172A")
    top_frame.pack(side=tk.TOP, fill=tk.X)

    path_label = tk.Label(top_frame, text="", font=("Poppins", 10, "bold"), bg="#0F172A", fg="#22D3EE", wraplength=680, justify="left")
    path_label.pack(padx=10, pady=5)

    # Main frames
    main_frame = tk.Frame(window, bg="#0F172A")
    main_frame.pack(fill=tk.BOTH, expand=True)

    left_frame = tk.Frame(main_frame, bg="#0F172A", width=350, height=450)
    left_frame.pack(side=tk.LEFT, fill=tk.Y)

    right_frame = tk.Frame(main_frame, bg="#0F172A", width=350, height=450)
    right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

    # Title
    title_label = tk.Label(left_frame, text="MCH Credentials", font=("Poppins", 16, "bold"), bg="#0F172A", fg="#F8FAFC")
    title_label.pack(pady=20)

    # Username Entry
    username_entry = tk.Entry(left_frame, font=("Poppins", 10), bd=0, bg="#1E293B", fg="#F8FAFC", insertbackground="#F8FAFC", width=24)
    username_entry.insert(0, "Username")
    username_entry.pack(pady=8, ipady=6)

    # Password Entry
    password_entry = tk.Entry(left_frame, font=("Poppins", 10), bd=0, bg="#1E293B", fg="#F8FAFC", insertbackground="#F8FAFC", width=24, show="*")
    password_entry.insert(0, "Password")
    password_entry.pack(pady=8, ipady=6)

    # Login Button
    login_button = tk.Button(left_frame, text="LOGIN", font=("Poppins", 9, "bold"), bg="#38BDF8", fg="#0F172A", width=22, pady=6,
                              activebackground="#0EA5E9", activeforeground="#FFFFFF",
                              command=lambda: browse_directory(username_entry.get(), password_entry.get(), log_area, path_label))
    login_button.pack(pady=10)

    # Logs Area
    log_area = scrolledtext.ScrolledText(right_frame, font=("Consolas", 9), bg="#0F172A", fg="#F8FAFC")
    log_area.pack(padx=8, pady=8, fill=tk.BOTH, expand=True)

    window.mainloop()

if __name__ == "__main__":
    main()
