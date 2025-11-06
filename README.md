# Python_codes
# Daily tasks automations
# This script will open a shared file from the sharepoint
# search the name tab and add the entries in row daily

import os
from openpyxl import load_workbook, Workbook
from datetime import datetime

# 🔧 Configuration
shared_path = r'C:\Path\To\Teams\SharedFolder\your_file.xlsx'  # Update this path
your_name = 'YourName'  # Replace with your actual name

# 📂 Check if file exists
if not os.path.exists(shared_path):
    print(f"File not found at {shared_path}")
    exit()

# 📖 Load workbook
wb = load_workbook(shared_path)

# 🆕 Create sheet if it doesn't exist
if your_name not in wb.sheetnames:
    wb.create_sheet(title=your_name)

# 📄 Select your sheet
ws = wb[your_name]

# 📅 Add today's entry
today = datetime.today().strftime('%Y-%m-%d')
ws.append([today, f"Entry for {today}"])

# 💾 Save workbook
wb.save(shared_path)
print(f"Updated {shared_path} with today's entry.")
