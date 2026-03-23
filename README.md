# 📁 Ultimate Drive Indexer for Google Sheets

A high-performance Google Apps Script that automatically syncs, indexes, and monitors Google Drive and Google Workspace Shared Drive files directly inside a Google Sheet. 

Built as a lightweight, faster, and free alternative to paid add-ons, this tool features smart upserts (delta syncing), native UI menus, and true-owner extraction for enterprise Shared Drives.

---

## ✨ Key Features

* **Universal Support:** Works seamlessly with standard "My Drive" folders and enterprise "Shared Drives".
* **True Owner Extraction:** Bypasses Google's Shared Drive limitations to extract the *actual* email address of the file creator/updater via the `lastModifyingUser` fallback.
* **Smart Upsert Engine:** Instead of wiping and rewriting the database, "Quick Sync" maps existing File IDs in `O(1)` time, only appending new files and updating modified ones.
* **Delta Tracking:** Automatically flags freshly added files with a `✅` in the "New" column, resetting on the next sync cycle.
* **Zero-Code Scheduling:** Set auto-refresh triggers (15m, 1hr, Daily) directly from a custom dropdown menu in the Sheets UI.
* **Rich Formatting:** Automatically applies zebra-striping, freezes headers, auto-resizes columns, and converts raw URLs into clickable rich-text hyperlinks.
* **Auto-Sorting:** Dynamically re-sorts the database by "Created At (Descending)" so the newest files are always at the top.

---

## 🚀 Installation & Setup

1. **Create a New Google Sheet:** Open a new or existing Google Sheet where you want the index to live.
2. **Open Apps Script:** In the top menu, go to `Extensions` > `Apps Script`.
3. **Paste the Code:** Delete any placeholder code (`function myFunction() {}`) and paste the provided `DriveIndexer.gs` code.
4. **Enable the Drive API (CRITICAL):**
   * In the left sidebar of the Apps Script editor, click the `+` icon next to **Services**.
   * Select **Drive API**, leave the version as `v3`, and click **Add**.
5. **Save and Reload:** Click the Save icon (💾) or press `Ctrl+S` / `Cmd+S`. Close the Apps Script tab and refresh your Google Sheet.

You will now see a custom **📁 Drive Index** menu next to "Help" in the Sheets toolbar!

---

## ⚙️ Configuration

Before running your first sync, you need to tell the script which drive or folder to monitor.

1. Click `📁 Drive Index` > `⚙️ Setup Config Sheet`.
2. A new tab named **Config** will be generated. Fill in the following:
   * **TARGET_DRIVE_ID:** Paste the ID of the folder or Shared Drive. *(You can find this in the URL: `drive.google.com/drive/folders/TARGET_DRIVE_ID`)*
   * **IS_SHARED_DRIVE:** Set to `TRUE` if the ID belongs to a Workspace Shared Drive, otherwise `FALSE`.
   * **INCLUDE_SUBFOLDERS:** Set to `TRUE` to recursively index all nested folders.
   * **TIMEZONE:** Defaults to your script's timezone (e.g., `Asia/Kolkata` or `America/New_York`).

---

## 📖 Usage Guide

Use the **📁 Drive Index** dropdown menu in Google Sheets to operate the script:

* **▶ Full Sync (All Files):** Wipes the current index clean and rebuilds the database from scratch. Recommended for the very first run or if data gets corrupted.
* **⚡ Quick Sync (New & Modified Only):** Scans the drive and only updates files that have changed, and appends brand new files. Much faster than a Full Sync.
* **⏰ Schedule Auto-Refresh:** Select an interval (15 mins, Hourly, 6 Hours, or Daily at 8 AM) to have the script run Quick Syncs automatically in the background. 
* **📋 View Sync Log:** Jumps to the auto-generated `Sync Log` tab to view a historical audit trail of all manual and automated syncs, including execution times and file counts.

---

## 📊 Data Dictionary (Columns Extracted)

The script automatically generates and populates the following columns:

| Column | Description |
| :--- | :--- |
| `#` | Auto-calculated row index sequence. |
| `✅ New` | Delta flag indicating files added since the last sync. |
| `File Name` | The name of the file or folder. |
| `File Type` | Human-readable format (e.g., Doc, Sheet, PDF, MP4). |
| `Folder Path` | Full breadcrumb directory path (`Marketing > Q1 > Assets`). |
| `Created At` | File creation timestamp (localized). |
| `Last Modified At` | Last modification timestamp (localized). |
| `Owner / Creator` | Email of the owner, or the last modifying user for Shared Drives. |
| `File Size` | Human-readable size format (KB, MB). |
| `Share Status` | Indicates if the file is Private or Shared. |
| `File Link` | Clickable rich-text hyperlink to open the file in Drive. |
| `Direct Download` | URL to directly download the file (bypassing the preview UI). |
| `Tags / Description` | Pulls data from the native Google Drive "Description" field. |
| `File ID` | Unique Google Drive ID (used by the script for O(1) Quick Sync mapping). |
| `Last Synced` | Timestamp of when this specific row was last updated by the script. |

---

## ⚠️ Troubleshooting & Limits

* **Execution Timeouts:** Google Apps Script has a hard 6-minute execution limit. The script includes a safety break at 8,000 files to prevent crashing. If your drive has >8,000 files, consider targeting specific subfolders instead of the root directory.
* **Missing Owner Emails:** If a user's email shows as their Name or "Unknown", your Google Workspace Administrator may have restricted internal directory API visibility. The script automatically falls back to Display Name in these scenarios.
* **Authorization Required:** The very first time you click a menu item, Google will prompt you to authorize the script. Click `Continue`, choose your account, click `Advanced`, and select `Go to Drive Indexer (unsafe)`. This is standard for custom Apps Scripts.
