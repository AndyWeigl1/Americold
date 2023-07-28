import os
import re
import shutil
import tkinter as tk
from tkinter import messagebox
from datetime import datetime

source_dir = r"C:\Users\Andy Weigl\OneDrive - Kodiak Cakes\Americold\Invoices"
base_destination_dir = r"C:\Users\Andy Weigl\Kodiak Cakes\Kodiak Cakes Team Site - Public\Vendors\Americold\Bills\Warehouse & Outbound Handling"

# Get a list of all files in the source directory
file_list = os.listdir(source_dir)

# Compile a regex pattern to extract the date from the file name
date_pattern = re.compile(r'(\d{1,2})\.(\d{1,2})\.(\d{2})')

# Iterate over each file, extract the date, and move it to the correct destination directory
for file_name in file_list:
    source_path = os.path.join(source_dir, file_name)

    # Extract the date from the file name
    match = date_pattern.search(file_name)
    if match:
        month, day, year = match.groups()
        year = "20" + year  # Convert 2-digit year to 4-digit year

        # Construct the destination path based on the date
        destination_dir = os.path.join(base_destination_dir, f"{year}\\{year}-{month.zfill(2)}")

        # Create the destination directory if it doesn't exist
        os.makedirs(destination_dir, exist_ok=True)

        # Move the file to the destination directory
        destination_path = os.path.join(destination_dir, file_name)
        shutil.move(source_path, destination_path)

messagebox.showinfo("File Move", "Files moved successfully!")