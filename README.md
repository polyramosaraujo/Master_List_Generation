# 📋 Automatic scan of Google Drive folder and generation of files checklist

This script automates the creation of a project checklist spreadsheet that lists all files stored on Google Drive. It is triggered whenever a form is submitted and executes a series of steps to organize, record, and generate an updated report with the files found in the project folders.

## 🚀 How It Works

- The function is automatically triggered by an `onEdit` trigger (whenever the spreadsheet is edited, usually through a form).
- Reads the main spreadsheet and identifies any new rows without a generated spreadsheet link.
- For each new row:
   - Records the date/time of the update.
   - Checks the project folder on Google Drive.
   - Deletes the previous spreadsheet (if any).
   - Creates a new spreadsheet and moves it to the project folder.
   - Iterates through subfolders (disciplines > formats > files).
   - Generates the content for the new spreadsheet (File name, discipline, format and direct link to file on Google Drive).
   - Fills the new spreadsheet with the generated content, formats it and adds checkboxes for user verification.
   - Updates the main spreadsheet with the link to the newly created spreadsheet and the date/time of creation.


## 📅 Last Update

28/06/2024


## 👨‍💻 Author

Polyana Ramos Araújo
