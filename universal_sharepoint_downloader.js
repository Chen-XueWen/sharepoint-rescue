/**
 * Universal SharePoint/OneDrive Strict Sequential Downloader
 * 
 * DESCRIPTION:
 * This script scans the current SharePoint/OneDrive page for visible files and downloads them
 * one by one to a local folder you select.
 * 
 * KEY FEATURES:
 * 1. Auto-Detection: Scans the page for files (no need to configure patterns!).
 * 2. Sequential: Downloads one file at a time. It waits for the file to be fully written to disk before starting the next.
 * 3. Data Integrity: Uses the File System Access API to stream data directly to disk.
 * 
 * HOW TO USE:
 * 1. Scroll down to the bottom of your SharePoint/OneDrive page to ensure all files are loaded.
 * 2. Open Developer Tools (F12 or Right Click -> Inspect).
 * 3. Go to the "Console" tab.
 * 4. Paste this entire script and press Enter.
 * 5. Select a local folder when prompted.
 */

(async () => {
    // =================================================================================
    //  CONFIGURATION
    // =================================================================================
    const CONFIG = {
        // Optional: Only download files containing this text (leave empty to download all)
        filterKeyword: "",

        // Delay between downloads (in milliseconds)
        delayBetweenFiles: 1000,

        // SharePoint usually hides the full URL, but we can construct it or find it.
        // Leave null to attempt auto-detection.
        forcedBaseUrl: null
    };

    // =================================================================================
    //  UI & HELPER FUNCTIONS
    // =================================================================================
    const ui = {
        init: () => {
            const old = document.getElementById('seq-dl-ui');
            if (old) old.remove();

            const div = document.createElement('div');
            div.id = 'seq-dl-ui';
            Object.assign(div.style, {
                position: 'fixed', top: '10px', right: '10px', zIndex: '999999',
                backgroundColor: 'white', color: '#333', padding: '15px',
                border: '2px solid #00ea00', borderRadius: '4px',
                boxShadow: '0 4px 12px rgba(0,0,0,0.15)', fontFamily: 'Segoe UI, sans-serif',
                minWidth: '320px', maxWidth: '400px'
            });
            div.innerHTML = `
                <h3 style="margin:0 0 10px;font-size:16px;color:#0078d4">Universal Downloader</h3>
                <div id="seq-status" style="margin-bottom:5px;font-weight:600">Initializing...</div>
                <progress id="seq-prog" value="0" max="100" style="width:100%;height:8px"></progress>
                <div id="seq-details" style="margin-top:5px;font-size:12px;color:#666;word-break:break-all"></div>
                <div id="seq-logs" style="margin-top:8px;max-height:150px;overflow-y:auto;font-size:11px;background:#f5f5f5;padding:5px;border:1px solid #eee"></div>
            `;
            document.body.appendChild(div);
        },
        setStatus: (txt) => { const el = document.getElementById('seq-status'); if (el) el.innerText = txt; },
        setDetails: (txt) => { const el = document.getElementById('seq-details'); if (el) el.innerText = txt; },
        setProgress: (val, max) => {
            const el = document.getElementById('seq-prog');
            if (el) { el.value = val; el.max = max; }
        },
        log: (txt, isError = false) => {
            const el = document.getElementById('seq-logs');
            if (el) {
                const line = document.createElement('div');
                line.style.color = isError ? 'red' : 'black';
                line.innerText = `[${new Date().toLocaleTimeString()}] ${txt}`;
                el.appendChild(line);
                el.scrollTop = el.scrollHeight;
            }
            if (isError) console.error(txt); else console.log(txt);
        }
    };

    const sleep = (ms) => new Promise(r => setTimeout(r, ms));

    // =================================================================================
    //  FILE DETECTION LOGIC
    // =================================================================================
    const detectFiles = () => {
        ui.log("Scanning DOM for files...");
        const files = [];
        const seenUrls = new Set();

        // Strategy A: Find rows with role="row" and look for links inside
        // Most SharePoint lists use role="row"
        const rows = Array.from(document.querySelectorAll('[role="row"]'));

        // Helper: clean URL (remove query params)
        const cleanUrl = (u) => {
            try { return u.split('?')[0]; } catch (e) { return u; }
        };

        // Helper to extract best filename
        const getBestFilename = (a) => {
            try {
                const rawUrl = a.href;
                // Remove query params and hash
                const cleanUrl = rawUrl.split('?')[0].split('#')[0];
                const urlName = decodeURIComponent(cleanUrl.split('/').pop());

                // If URL filename looks valid/useful, prefer it (captures .part-aa, etc.)
                if (urlName && urlName.includes('.') && urlName.length > 2) {
                    return urlName;
                }
            } catch (e) { }
            return a.innerText || a.getAttribute('title') || "unknown_file";
        };

        rows.forEach(row => {
            // Find the primary link/name
            // Usually valid file links contain '/Documents/' or similar, and end with a file extension
            // We look for any anchor that doesn't look like a folder or system link
            const links = Array.from(row.querySelectorAll('a'));

            // Filter for meaningful links
            const fileLink = links.find(a => {
                const href = a.href || "";
                // Heuristics for a file link
                const isJS = href.startsWith('javascript:');
                const isAspx = href.includes('.aspx'); // usually a page/folder view, not a file
                const hasExt = href.split('/').pop().includes('.'); // rough check for extension
                return !isJS && !isAspx && hasExt;
            });

            if (fileLink) {
                const name = getBestFilename(fileLink);
                const url = fileLink.href;

                // Dedupe
                if (!seenUrls.has(url)) {
                    // Apply Filter
                    if (CONFIG.filterKeyword && !name.includes(CONFIG.filterKeyword)) return;

                    seenUrls.add(url);
                    files.push({ name: name.trim(), url: url });
                }
            }
        });

        // Strategy B: Fallback for Card View or other layouts
        // If Strategy A found nothing, try finding ALL interesting links
        if (files.length === 0) {
            ui.log("Row detection returned 0. Trying global link scan...");
            document.querySelectorAll('a').forEach(a => {
                const href = a.href || "";
                if (href.startsWith('javascript') || href.includes('.aspx')) return;

                // Check if it looks like a file (has extension)
                const segment = href.split('/').pop();
                if (segment.includes('.') && segment.length > 3) {
                    if (!seenUrls.has(href)) {
                        const name = getBestFilename(a);
                        if (CONFIG.filterKeyword && !name.includes(CONFIG.filterKeyword)) return;
                        seenUrls.add(href);
                        files.push({ name: name.trim(), url: href });
                    }
                }
            });
        }

        // =================================================================================
        //  filename DEDUPLICATION
        // =================================================================================
        // If multiple files have the exact same name (common in SharePoint "parts" or versioning),
        // we rename them to "Filename (1).ext", "Filename (2).ext", etc.
        const uniqueFiles = [];
        const usedNames = new Set();

        files.forEach(f => {
            let finalName = f.name;
            let counter = 1;

            // Separate extension for cleaner renaming: "file.tar.gz" -> "file (1).tar.gz"
            // We'll treat the *last* dot as the extension separator for simplicity, 
            // or if it ends in .tar.gz, we might want to preserve that. 
            // Standard 'path.parse' logic:
            let base = finalName;
            let ext = "";
            const lastDot = finalName.lastIndexOf('.');

            if (lastDot > 0) {
                base = finalName.substring(0, lastDot);
                ext = finalName.substring(lastDot);
            }

            while (usedNames.has(finalName)) {
                finalName = `${base} (${counter})${ext}`;
                counter++;
            }

            usedNames.add(finalName);
            f.name = finalName;
            uniqueFiles.push(f);
        });

        return uniqueFiles;
    };


    // =================================================================================
    //  MAIN LOGIC
    // =================================================================================
    try {
        ui.init();

        // 1. Browser Capability Check
        if (!window.showDirectoryPicker) {
            throw new Error("Your browser does not support the File System Access API. Please use Chrome, Edge, or Opera.");
        }

        // 2. Scan Files
        ui.setStatus("Scanning page...");
        const targetFiles = detectFiles();

        if (targetFiles.length === 0) {
            throw new Error("No files detected! Please scroll down to ensure files are visible, or check the page structure.");
        }

        ui.log(`Detected ${targetFiles.length} unique files.`);
        ui.log(`First file: ${targetFiles[0].name}`);

        // 3. User Directory Selection
        ui.setStatus("Waiting for folder selection...");
        const dirHandle = await window.showDirectoryPicker();
        ui.log("Target folder selected.");

        // 4. Download Loop
        ui.setStatus("Starting download...");
        let successCount = 0;
        let failCount = 0;

        for (let i = 0; i < targetFiles.length; i++) {
            const fileObj = targetFiles[i];
            const fileName = fileObj.name;
            // Append download=1 to force raw file download
            const separator = fileObj.url.includes('?') ? '&' : '?';
            const fileUrl = `${fileObj.url}${separator}download=1`;

            ui.setStatus(`Downloading ${i + 1}/${targetFiles.length}`);
            ui.setDetails(fileName);
            ui.setProgress(i, targetFiles.length);

            try {
                // Check if file exists (Optional: Skip if exists)
                // try { await dirHandle.getFileHandle(fileName); ui.log(`Skipping ${fileName} (Exists)`); continue; } catch(e){}

                ui.log(`Fetching: ${fileName}`);
                const response = await fetch(fileUrl);

                if (!response.ok) throw new Error(`HTTP ${response.status} - ${response.statusText}`);

                // We use streams to pipe directly to disk
                const fileHandle = await dirHandle.getFileHandle(fileName, { create: true });
                const writable = await fileHandle.createWritable();

                if (response.body) {
                    await response.body.pipeTo(writable);
                } else {
                    const blob = await response.blob();
                    await writable.write(blob);
                    await writable.close();
                }

                successCount++;
                ui.log(`Saved: ${fileName}`);

                if (CONFIG.delayBetweenFiles > 0) await sleep(CONFIG.delayBetweenFiles);

            } catch (err) {
                failCount++;
                ui.log(`ERROR on ${fileName}: ${err.message}`, true);
                // We do NOT stop, continue to next
            }
        }

        ui.setStatus("Job Complete!");
        ui.setDetails(`Success: ${successCount}, Failed: ${failCount}`);
        ui.setProgress(100, 100);
        ui.log("All operations finished.");

    } catch (err) {
        ui.setStatus("Error");
        ui.log(err.message, true);
        alert("Script Error: " + err.message);
    }
})();
