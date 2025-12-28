FILE RENAMER PRO v9 - USER GUIDE
=================================

WHAT IS THIS?
-------------
A simple tool to rename multiple files at once using an Excel spreadsheet.


QUICK START
-----------
1. Run FileRenamerProV9.exe
2. Drag your folder into the "Target Folder" field
3. Click "🔍 Scan Folder & Create Template"
4. Open the Excel file that was created
5. Edit the column you want to use:
   - "Prefix" column = adds text before filename
   - "New_Filename" column = completely new name
6. Save the Excel file
7. In the app, click "Preview Changes"
8. If it looks good, click "Run Rename"


FEATURES
--------
✓ Drag & drop files and folders
✓ Automatic template creation from existing files
✓ Preview changes before applying
✓ Automatic backups
✓ Detailed log files
✓ Two rename modes: Prefix and Replace


IMPORTANT TIPS
--------------
- Always enable "Create backup" the first time
- Use "Preview Changes" before renaming
- Files are renamed in alphabetical order
- Excel Row 1 = First file alphabetically
- Keep the backup folder until you're sure everything is correct


NEED HELP?
----------
Click the "❓ Help" button in the app for detailed instructions.


SYSTEM REQUIREMENTS
-------------------
- Windows 7 or later
- No other software needed (includes everything)




VERSION
-------

V9 Critical Fixes Summary:
1. Excel File Lock Issue - SOLVED ✓

Removed caching system that held file references
Implemented read_excel_safe() - reads fresh every time
Added clear_excel_cache() called after every operation
Excel can now be opened/edited/re-run without restarting .exe

2. Incorrect File Matching - SOLVED ✓

Strict Column B matching - always uses Current_Filename to find files
No more fallback to index-based matching that caused wrong renames
Skips rows where Column B is empty or doesn't match existing files
Clear warnings in preview when files aren't found

3. Security Audit - COMPLETED ✓

100% local operation - no network connections anywhere
Added validate_local_path() to prevent path traversal attacks
All file operations confined to target folder
No external APIs, no data transmission
Configuration saved locally only

4. Additional Improvements:

Better error messages explaining Column B requirements
Instructions emphasize not changing Column B
Preview shows warnings for missing/unmatched files
Won't allow Run if files are missing (must fix Column B first)
Only Column B matching → only renames exact matches
Explicit skips logged for debugging

Build Command: 
FILE RENAMER PRO V9 - BUILD INSTRUCTIONS
========================================

STEP 1: Open Command Prompt or PowerShell
----------------------------------------
Windows Key + R, type "cmd" or "powershell", press Enter

STEP 2: Navigate to Your Project Folder
---------------------------------------
cd C:\Path\To\Your\FileRenamerV9\Folder
(Replace with your actual folder path)

STEP 3: Install Dependencies
----------------------------
Copy and paste this command:
python -m pip install -r requirements.txt

If you see "Successfully installed..." messages, proceed.

If pip fails due to permissions, try:
python -m pip install --user -r requirements.txt

STEP 4: Build the Executable
----------------------------
Copy and paste this command:

py -m PyInstaller --onefile --windowed --name FileRenamerProV9 --hidden-import=tkinterdnd2 --hidden-import=tkinter File-Org-V9.py --clean

Notes:
- Use "py" if "python" is not on PATH on Windows.
- The hidden-imports ensure tkinterdnd2 and tkinter are available to the bundle.

STEP 5: Find Your Application
-----------------------------
Open the "dist" folder in your project directory.
Your executable is: FileRenamerProV9.exe

TROUBLESHOOTING
---------------
- If Step 3 fails: use --user flag: python -m pip install --user -r requirements.txt
- If Step 4 fails with "python not recognized": use "py -m PyInstaller ..." instead of "python -m PyInstaller ..."
- If you get "tkinterdnd2 not found" error:
  python -m pip install tkinterdnd2 --force-reinstall

REBUILD (After Making Changes)
------------------------------
py -m PyInstaller --onefile --windowed --name FileRenamerProV9 --hidden-import=tkinterdnd2 --hidden-import=tkinter File-Org-V9.py --clean
