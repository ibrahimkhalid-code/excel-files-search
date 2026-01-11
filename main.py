import os
import threading
import tkinter as tk
from tkinter import messagebox

excel_files = []
ignore_folders = [
    "C:\\Windows",
    "C:\\Program Files",
    "C:\\Program Files (x86)",
    "C:\\Program Files (x86)",
    "D:\\$RECYCLE.BIN",
]


def search_drive(drive):
    global excel_files
    for current_root, dirs, files in os.walk(drive, topdown=True):

        dirs[:] = [
            d for d in dirs if os.path.join(current_root, d) not in ignore_folders
        ]
        for file in files:
            try:
                if file.lower().endswith((".xls", ".xlsx", ".doc", ".docx")):
                    full_path = os.path.join(current_root, file)
                    excel_files.append(full_path)

                    root.after(
                        0,
                        lambda p=full_path: excel_files_listbox.insert(tk.END, p),
                    )
            except PermissionError:
                pass


def start_search():
    drives_input = drives_entry.get()
    drives = [d.strip() for d in drives_input.split(",") if d.strip()]
    if not drives:
        messagebox.showerror("Error", "Enter at least one drive")
        return

    excel_files_listbox.delete(0, tk.END)
    threads = []
    for drive in drives:
        t = threading.Thread(target=search_drive, args=(drive,), daemon=True)
        t.start()
        threads.append(t)


def open_file(event):
    selected = excel_files_listbox.curselection()
    if selected:
        file_path = excel_files_listbox.get(selected)
        try:
            os.startfile(file_path)
        except Exception as e:
            messagebox.showerror("Error", f"Cannot open file:\n{e}")


root = tk.Tk()
root.title("Excel Files Finder")
root.geometry("800x500")

tk.Label(root, text="Enter drives (comma separated, e.g., C:\\, D:\\)").pack(pady=5)
drives_entry = tk.Entry(root, width=50)
drives_entry.pack(pady=5)
drives_entry.insert(0, "C:\\, D:\\")

tk.Button(root, text="Search Excel Files", command=start_search).pack(pady=10)

excel_files_listbox = tk.Listbox(root, width=100, height=20)
excel_files_listbox.pack(pady=10)
excel_files_listbox.bind("<Double-1>", open_file)

root.mainloop()
