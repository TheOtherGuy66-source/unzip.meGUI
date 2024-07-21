import os
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import zipfile
import tarfile
import gzip
import bz2
import lzma
import shutil
import threading
import logging
import subprocess
import sys

# Ensure required packages are installed
required_packages = ['py7zr']

def install_and_import(package):
    try:
        __import__(package)
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])
        __import__(package)

for package in required_packages:
    install_and_import(package)

import py7zr

# Set up logging
logging.basicConfig(filename='unzip_me_gui.log', level=logging.INFO, format='%(asctime)s - %(message)s')

# Check and log missing packages
def check_and_log_packages():
    missing_packages = []
    for package in required_packages:
        try:
            __import__(package)
        except ImportError:
            missing_packages.append(package)
    
    if missing_packages:
        desktop_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
        log_file_path = os.path.join(desktop_path, "missing_packages_log.txt")
        with open(log_file_path, "w") as log_file:
            log_file.write("The following packages are missing:\n")
            for package in missing_packages:
                log_file.write(f"{package}\n")
        messagebox.showwarning("Missing Packages", f"Some required packages are missing. Please check the log file on your desktop: {log_file_path}")

# Function to create a directory if it does not exist
def ensure_dir(directory):
    if not os.path.exists(directory):
        os.makedirs(directory)

# Function to display status messages
def display_status(message, color, duration=10):
    status_output_label.config(text=message, fg=color)
    root.after(duration * 1000, clear_status)

# Function to clear status messages
def clear_status():
    status_output_label.config(text="")

# Function to log messages
def log_message(message):
    logging.info(message)
    print(message)

# Function to extract archive file to the specified output directory
def extract_archive(archive_type):
    def extract():
        file_path = filedialog.askopenfilename(filetypes=[("Archive Files", f"*.{archive_type}")])
        if not file_path:
            return
        
        output_dir = filedialog.askdirectory()
        if not output_dir:
            return
        
        display_status(f"Extracting {file_path} to {output_dir}", "yellow")
        log_message(f"Extracting {file_path} to {output_dir}")
        password = None

        if archive_type in ["zip", "7z"]:
            try:
                if archive_type == "zip":
                    with zipfile.ZipFile(file_path, 'r') as zip_ref:
                        if zip_ref.testzip() is None:
                            pass
                        else:
                            password = simpledialog.askstring("Password", "Enter password (if any):", show='*')
                elif archive_type == "7z":
                    with py7zr.SevenZipFile(file_path, 'r') as seven_z_ref:
                        if seven_z_ref.needs_password():
                            password = simpledialog.askstring("Password", "Enter password (if any):", show='*')
            except Exception as e:
                error_message = f"Error checking password requirement: {e}"
                display_status(error_message, "red")
                log_message(error_message)
                return
        
        try:
            if archive_type == "zip":
                with zipfile.ZipFile(file_path, 'r') as zip_ref:
                    if password:
                        zip_ref.setpassword(password.encode())
                    for file_info in zip_ref.infolist():
                        log_message(f"Extracting {file_info.filename}")
                        zip_ref.extract(file_info, output_dir)
                        log_message(f"Extracted {file_info.filename}")
            elif archive_type == "7z":
                try:
                    with py7zr.SevenZipFile(file_path, 'r', password=password) as seven_z_ref:
                        for name in seven_z_ref.getnames():
                            log_message(f"Extracting {name}")
                        seven_z_ref.extractall(output_dir)
                        log_message(f"Extracted all files")
                except Exception as e:
                    if "BCJ2 filter is not supported" in str(e):
                        error_message = f"BCJ2 filter is not supported by py7zr. Please use another tool to extract {file_path}."
                        display_status(error_message, "red")
                        log_message(error_message)
                    else:
                        error_message = f"Error extracting {file_path}: {e}"
                        display_status(error_message, "red")
                        log_message(error_message)
                    return
            elif archive_type == "tar":
                with tarfile.open(file_path, 'r') as tar_ref:
                    for member in tar_ref.getmembers():
                        log_message(f"Extracting {member.name}")
                    tar_ref.extractall(output_dir)
                    log_message(f"Extracted all files")
            elif archive_type == "gz":
                output_file_path = os.path.join(output_dir, os.path.basename(file_path).replace('.gz', ''))
                with gzip.open(file_path, 'rb') as gz_ref, open(output_file_path, 'wb') as out_f:
                    log_message(f"Extracting {file_path}")
                    shutil.copyfileobj(gz_ref, out_f)
                    log_message(f"Extracted {output_file_path}")
            elif archive_type == "bz2":
                output_file_path = os.path.join(output_dir, os.path.basename(file_path).replace('.bz2', ''))
                with bz2.open(file_path, 'rb') as bz2_ref, open(output_file_path, 'wb') as out_f:
                    log_message(f"Extracting {file_path}")
                    shutil.copyfileobj(bz2_ref, out_f)
                    log_message(f"Extracted {output_file_path}")
            elif archive_type == "xz":
                output_file_path = os.path.join(output_dir, os.path.basename(file_path).replace('.xz', ''))
                with lzma.open(file_path, 'rb') as xz_ref, open(output_file_path, 'wb') as out_f:
                    log_message(f"Extracting {file_path}")
                    shutil.copyfileobj(xz_ref, out_f)
                    log_message(f"Extracted {output_file_path}")
            display_status(f"Extracted: {file_path} to {output_dir}", "green")
            log_message(f"Extracted: {file_path} to {output_dir}")
        except FileNotFoundError:
            error_message = f"File not found: {file_path}"
            display_status(error_message, "red")
            log_message(error_message)
        except PermissionError:
            error_message = f"Permission denied: {file_path}"
            display_status(error_message, "red")
            log_message(error_message)
        except Exception as e:
            error_message = f"Error extracting {file_path}: {e}"
            display_status(error_message, "red")
            log_message(error_message)

    threading.Thread(target=extract).start()

# Function to create a zip archive from selected files or directories
def create_zip_archive():
    def create_zip():
        dir_path = filedialog.askdirectory(title="Select a directory to compress")
        if not dir_path:
            return

        archive_path = filedialog.asksaveasfilename(defaultextension=".zip", filetypes=[("ZIP files", "*.zip")])
        if not archive_path:
            return

        try:
            log_message(f"Creating {archive_path}")
            with zipfile.ZipFile(archive_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for root, dirs, files in os.walk(dir_path):
                    for file in files:
                        file_full_path = os.path.join(root, file)
                        arcname = os.path.relpath(file_full_path, dir_path)
                        zipf.write(file_full_path, arcname)
            display_status(f"Created: {archive_path}", "green")
            log_message(f"Created: {archive_path}")
        except Exception as e:
            error_message = f"Error creating archive: {e}"
            display_status(error_message, "red")
            log_message(error_message)

    threading.Thread(target=create_zip).start()

# Function to create a 7z archive from selected files or directories
def create_7z_archive():
    def create_7z():
        dir_path = filedialog.askdirectory(title="Select a directory to compress")
        if not dir_path:
            return

        archive_path = filedialog.asksaveasfilename(defaultextension=".7z", filetypes=[("7z files", "*.7z")])
        if not archive_path:
            return

        try:
            log_message(f"Creating {archive_path}")
            with py7zr.SevenZipFile(archive_path, 'w') as seven_z_ref:
                for root, dirs, files in os.walk(dir_path):
                    for file in files:
                        file_full_path = os.path.join(root, file)
                        arcname = os.path.relpath(file_full_path, dir_path)
                        seven_z_ref.write(file_full_path, arcname)
            display_status(f"Created: {archive_path}", "green")
            log_message(f"Created: {archive_path}")
        except Exception as e:
            error_message = f"Error creating archive: {e}"
            display_status(error_message, "red")
            log_message(error_message)

    threading.Thread(target=create_7z).start()

# Function to create a tar.gz archive from selected files or directories
def create_tar_gz_archive():
    def create_tar_gz():
        dir_path = filedialog.askdirectory(title="Select a directory to compress")
        if not dir_path:
            return

        archive_path = filedialog.asksaveasfilename(defaultextension=".tar.gz", filetypes=[("TAR.GZ files", "*.tar.gz")])
        if not archive_path:
            return

        try:
            log_message(f"Creating {archive_path}")
            with tarfile.open(archive_path, 'w:gz') as tarf:
                for root, dirs, files in os.walk(dir_path):
                    for file in files:
                        file_full_path = os.path.join(root, file)
                        arcname = os.path.relpath(file_full_path, dir_path)
                        tarf.add(file_full_path, arcname)
            display_status(f"Created: {archive_path}", "green")
            log_message(f"Created: {archive_path}")
        except Exception as e:
            error_message = f"Error creating archive: {e}"
            display_status(error_message, "red")
            log_message(error_message)

    threading.Thread(target=create_tar_gz).start()

# Main script
check_and_log_packages()

# Initialize the main application window
root = tk.Tk()
root.title("unzip.meGUI")
root.configure(bg="#1e1e1e")
font_style_label = ("Courier", 20, "bold")
font_style_button = ("Courier", 12, "bold")

# Status label
status_label = tk.Label(root, text="unzip.meGUI", bg="#1e1e1e", fg="red", font=("Courier", 31))
status_label.pack(pady=5)

# Button frame
button_frame = tk.Frame(root, bg="#1e1e1e")
button_frame.pack(pady=10)

# Status output frame
status_frame = tk.Frame(root, bg="#1e1e1e")
status_frame.pack(pady=10, fill=tk.X)
status_output_label = tk.Label(status_frame, text="", bg="#1e1e1e", fg="white", font=("Courier", 12))
status_output_label.pack(pady=5)

# Adding buttons for each archive format
archive_formats = ["zip", "7z", "tar", "gz", "bz2", "xz"]
for idx, fmt in enumerate(archive_formats):
    button = tk.Button(button_frame, text=f"Extract .{fmt}", command=lambda fmt=fmt: extract_archive(fmt), bg="#1e1e1e", fg="red", font=font_style_button)
    button.grid(row=idx//3, column=idx%3, padx=10, pady=10)
    button.bind("<Enter>", lambda e, b=button: b.config(relief=tk.SUNKEN))
    button.bind("<Leave>", lambda e, b=button: b.config(relief=tk.RAISED))

# Button to create zip archive
create_zip_button = tk.Button(button_frame, text="Create .zip", command=create_zip_archive, bg="#1e1e1e", fg="red", font=font_style_button)
create_zip_button.grid(row=(len(archive_formats)//3) + 1, column=0, padx=10, pady=10)

# Button to create 7z archive
create_7z_button = tk.Button(button_frame, text="Create .7z", command=create_7z_archive, bg="#1e1e1e", fg="red", font=font_style_button)
create_7z_button.grid(row=(len(archive_formats)//3) + 1, column=1, padx=10, pady=10)

# Button to create tar.gz archive
create_tar_gz_button = tk.Button(button_frame, text="Create .tar.gz", command=create_tar_gz_archive, bg="#1e1e1e", fg="red", font=font_style_button)
create_tar_gz_button.grid(row=(len(archive_formats)//3) + 1, column=2, padx=10, pady=10)

root.mainloop()
