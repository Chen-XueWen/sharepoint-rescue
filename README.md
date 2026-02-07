# Universal SharePoint Sequential Downloader

A robust JavaScript tool designed to download **all visible files** from a SharePoint/OneDrive page directly to your local file system.

## 💡 Motivation

1. **Local Backup Difficulty**: While Microsoft OneDrive makes uploading easy, performing a full local backup is often difficult and cumbersome.
2. **Time-Consuming Classic View**: In the classic SharePoint web interface, users are often forced to download files one by one, which is extremely time-consuming for large datasets.
3. **Unreliable "Download Folder"**: In the newer M365 interface, the "Download" folder button often fails or results in partial/corrupted downloads when dealing with large files.
4. **Automation**: This script works (especially with the classic view) to automate the tedious process of downloading files one by one, ensuring data integrity and saving time.

## 🚀 Key Features

- **Auto-Detection**: No need to configure filenames! The script scans the page and detects all files (PDFs, ZIPs, Videos, etc.).
- **Strict Sequential Downloading**: Downloads one file at a time. It waits for file A to be explicitly saved to disk before starting file B.
- **Data Integrity**: Uses the **File System Access API** to stream data directly to a local folder.
- **Resilient**: If a single file fails, it logs the error and continues to the next one.

## 📋 Prerequisites

- **Browser**: Google Chrome, Microsoft Edge, or Opera (Browsers that support File System Access API).
- **Access**: You must have read access to the SharePoint/OneDrive folder.

## 🛠️ Usage

1. **Prepare the Page**
    - Navigate to the SharePoint/OneDrive folder.
    - **Scroll Down**: Scroll to the bottom of the list to ensure **all files are loaded** in the browser view. (The script only sees what is currently loaded).

2. **Run the Script**
    - Open Developer Tools (`F12` -> `Console`).
    - Copy & Paste the code from `universal_sharepoint_downloader.js`.
    - Press **Enter**.

3. **Select Output Folder**
    - A browser prompt will appear.
    - Select (or create) the folder anywhere on your computer where you want to save the files.
    - Click **Allow** if asked for permission.

4. **Sit Back**
    - A status box in the top-right corner will show the progress.
    - You can see exactly which file is being verified and saved in the embedded logs.

## ⚙️ Configuration (Optional)

If you only want to download specific files, you can edit the top of the script:

```javascript
const CONFIG = {
    // Only download files containing this text (e.g. "Draft" or ".pdf")
    filterKeyword: "", 
    
    // Delay between downloads (in milliseconds)
    delayBetweenFiles: 1000
};
```
