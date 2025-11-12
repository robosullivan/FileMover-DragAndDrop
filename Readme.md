# FileMover : Drag-and-Drop

###### Drag. Drop. Moved. ######  
A lightweight VBScript utility to move files with undo, logging, balloon tips, and tidy setup.

---

## ✨ Features

- **Drag-and-drop execution**: Drop files onto the FileMover script icon to move them instantly.
- **First-time setup wizard**: Choose a destination folder by browsing or typing a path (UNC supported).
- **Network-aware folder picker**: Detects local vs. UNC paths and starts the folder dialog in the right place.
- **Persistent destination config**: Stores your chosen folder in a config file for reuse.
- **Writable folder check**: Verifies destination permissions before moving files.
- **Filename collision detection, unique file renaming**: Prevents overwrites by appending `(1)`, `(2)`, etc.
- **Logging system**: Daily log files record user, action, source, destination, and status.
- **Undo last move**: If no files are dropped, you can undo the most recent successful move.
- **Balloon notifications**: Summarises each move with Windows balloon tips (single file or batch summary).
- **Optional hiding**: Config file and log folder can be hidden for a cleaner directory.
- **Config file stored alongside script**: Always recreated in the script’s folder if missing, ensuring predictable setup.
- **Safe folder skip**: Skips dropped folders, logs the skip, and avoids errors.
- **Cloud storage compatibility**: Works with synced folders from Google Drive and OneDrive (via their desktop apps).

---

## 🚀 Getting Started

1. **Download or copy the script** into a folder of your choice.
2. **Double-click the script** the first time to run setup:
   - Browse for a destination folder, or
   - Type a full path (UNC paths allowed).
3. **Drag files onto the script icon** to move them automatically.
4. **Check balloon tips** for a quick summary of what happened.
5. **Review logs** in the `logs` folder for detailed auditing.
6. **Undo last move** by double-clicking the script with no files dropped.

---

## 📂 Example Use Cases

- Organise photos by renaming a copy of the script to `Photos.vbs` and dragging images onto it.  
- Archive bills with another renamed copy, e.g. `My Bills.vbs`.  
- Maintain multiple movers in one directory, each renamed to reflect its destination (e.g. `Invoices.vbs`, `Reports.vbs`).  

---

## 🔄 Workflow Integration

FileMover isn’t just for personal organisation — it can streamline repetitive workplace workflows too.
 
By configuring multiple movers in one directory, each pointing to different destinations, you can stage files for downstream tasks with minimal effort.

- **Processed documents** can be moved into a controlled folder for manual review workflows.  
- **Team handoffs** become simpler when each mover represents a step in the process (e.g. “Reviewed”, “Archived”, “Ready for QA”).  
- **Network‑based workplaces and synced cloud drives** benefit from the UNC‑aware setup, ensuring paths are valid across shared environments.  
- **Built‑in logging** ensures every move is recorded with user, source, destination, and status — providing an audit trail for workplace processes.  

This makes FileMover a lightweight way to enforce consistency in manual workflows without complex automation.

---

## 🛡 Reliability Notes

- Config and logs are always stored alongside the script, ensuring predictable setup and behaviour.  
- If the config file is missing, FileMover automatically recreates it during setup.  
- Config and logs can be hidden but remain accessible if needed.  
- Logs are daily and user‑tagged, providing a clear audit trail of actions.  
- Undo is limited to the most recent successful move in today’s log.  
- Balloon tips provide immediate feedback, even for skipped folders or renamed files.  
- Works reliably with synced cloud folders (Google Drive, OneDrive) via their desktop apps, since they expose local paths for FileMover to use.
- On managed networks, administrators can set script and config files as undeletable, adding extra protection against accidental removal.

---

## 📜 License

MIT License.  
Use freely, but please retain attribution and notes in derivative works.
