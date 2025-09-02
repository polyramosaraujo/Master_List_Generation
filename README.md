# 📋 Automatic Master Checklist Generation

This script automates the creation of a master checklist spreadsheet for project files stored in Google Drive. It is triggered whenever a form is submitted and performs a series of steps to organize, log, and generate an updated report with the files found in the project folders.

## 🚀 How It Works

- The function is automatically triggered by an `onEdit` event (whenever the spreadsheet is edited, usually through a form submission).
- Reads the main spreadsheet and checks for new rows without a generated sheet link.
- For each new row:
  - Records the update date/time.
  - Checks the project folder in Drive.
  - Deletes the previous sheet (if it exists).
  - Creates a new sheet and moves it into the project folder.
  - Iterates through the subfolders (disciplines > formats > files).
  - Generates the content for the new sheet (file name, discipline, format, and direct Google Drive link).
  - Populates the new sheet with the generated content, formats it, and adds checkboxes for user verification.
  - Updates the main spreadsheet with the link to the newly created sheet and its creation date/time.

## 📅 Last Update

06/28/2024

## 👨‍💻 Author

Polyana Ramos Araújo
