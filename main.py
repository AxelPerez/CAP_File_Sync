import os
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import re

destination = r"\\your\destination\folder\goes\here"
audit_base_folder = "\\\\your\audit\base\folder\goes\here"

fiscal_year_to_audit_folder = {
    "2018": "1. FY18 AUDIT",
    "2019": "2. FY19 AUDIT",
    "2020": "3. FY20 AUDIT",
    "2021": "4. FY21 AUDIT",
    "2022": "5. FY22 AUDIT",
    "2023": "6. FY22 AUDIT"
}

def get_source_paths():
    source_paths = []

    while True:
        source_path = filedialog.askdirectory(title="Select source folder")
        if not source_path:
            break
        source_paths.append(source_path)

        if not messagebox.askyesno("Another folder", "Do you want to select another source folder?"):
            break

    return source_paths

def sync_folders(source_path, destination, text_to_append, progress_label):
   

    # Sync folders
    for foldername, _, filenames in os.walk(source_path):
        if "Desktop" not in foldername.split(os.sep):
            for filename in filenames:
                if filename.endswith('.xls') or filename.endswith('.xlsx'):
                    if "-FIN" in filename:
                        source_file_path = os.path.join(foldername, filename)
                        destination_folder = foldername.replace(source_path, destination)
                        destination_file_name = f"{os.path.splitext(filename)[0]} {text_to_append}{os.path.splitext(filename)[1]}"
                        destination_file_path = os.path.join(destination_folder, destination_file_name)

                        if not os.path.exists(destination_folder):
                            os.makedirs(destination_folder)

                        progress_label.config(text=f"{filename}")  # Update progress label
                        shutil.copy2(source_file_path, destination_file_path)

                        # Get fiscal year and unique ID
                        match = re.search(r"-(\d+)-(\d+)", filename)
                        if match:
                            fiscal_year = match.group(1)
                            unique_id = match.group(2)

                            audit_folder_name = fiscal_year_to_audit_folder.get(fiscal_year)
                            if audit_folder_name:
                                audit_folder = os.path.join(audit_base_folder, f"{audit_folder_name}\\1. Open")

                            target_folder = None

                            for folder in os.listdir(audit_folder):
                                folder_prefix = folder[:16]
                                if folder_prefix == f"WCF-FIN-{fiscal_year}-{unique_id}" or folder_prefix == f"GF-FIN-{fiscal_year}-{unique_id}":
                                    target_folder = os.path.join(audit_folder, folder)
                                    break

                            if target_folder:
                                cap_folder = os.path.join(target_folder, "CAP")
                                if not os.path.exists(cap_folder):
                                    os.makedirs(cap_folder)

                                specific_destination_path = os.path.join(cap_folder, destination_file_name)
                                shutil.copy2(source_file_path, specific_destination_path)

def show_message():
    messagebox.showinfo("Script completed", "New CAP Files have been saved.")

def get_text_to_append():
    text_to_append = simpledialog.askstring("Text to append", "Enter the text to append to the end of each CAP filename:")
    return text_to_append

def ask_to_run_script():
    result = messagebox.askyesno("Run the script", "Would you like to run the CAP File script?")
    return result

def run_script():
    if ask_to_run_script():
        source_paths = get_source_paths()
        text_to_append = get_text_to_append()

        # Create a message box that shows the script is running
        progress_window = tk.Toplevel()
        progress_window.title("Script Running")
        progress_window.geometry("400x100")
        progress_label = tk.Label(progress_window, text="Please wait while the script runs...")
        progress_label.pack()
        progress_window.update()
        progress_window.lift()

        for source_path in source_paths:
            sync_folders(source_path, destination, text_to_append, progress_label)

        # Destroy the progress window when the script is done running
        progress_window.destroy()

        show_message()

root = tk.Tk()
root.withdraw()

run_script()
