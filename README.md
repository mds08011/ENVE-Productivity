# ENVE-Productivity

A collection of automation scripts and tools designed to streamline civil engineering project management workflows. These tools assist with file organization, email archiving, drawing management, and other repetitive tasks common in water/wastewater utility design.

## Repository Contents

### 1. Outlook Automation (VBA)
Located in: `/Email Saving`

**Goal:** Automate the process of saving project emails and attachments to network drives in a consistent format.

* **Function:** Creates a dated folder (`yyyy.mm.dd - Sender - Subject`) in your specified project directory.
* **Output:** Saves the email as `.msg`, the email body as `.pdf`, and extracts all attachments.
* **Mechanism:** Uses a VBA Macro triggered by a button in the Outlook ribbon. It prompts the user for the Project Name to ensure files go to the correct directory on the K: drive.

** Setup Instructions for Outlook:**
1.  **Enable Developer Tab:** Go to `File > Options > Customize Ribbon` and check **Developer**.
2.  **Add Reference:** In the VBA Editor (`Alt+F11`), go to `Tools > References` and check **Microsoft Word 16.0 Object Library** (required for PDF export).
3.  **Import Code:** Copy the code from the `/Email Saving` folder into a new Module.
4.  **Configure Path:** Update the `projectRootPath` variable in the code to match your company's network drive structure.

---

### 2. File Management Utilities (PowerShell)

These scripts are designed to be run directly on Windows. Right-click any `.ps1` file and select **"Run with PowerShell"**.

* **`/Clean Empty Folder`**
    * Recursively scans a directory and deletes any subfolders that contain no files. Useful for cleaning up vendor download structures.
* **`/Clean Model Folders`**
    * Targets hydraulic modeling folders (WaterGEMS, InfoWorks, etc.) and removes temporary simulation files (`.out`, `.log`, `.bak`), leaving only the essential model source files.
* **`/Create File List`**
    * Generates a clean text-based "Table of Contents" of every file in a folder and its subfolders. Ideal for creating transmittal logs for client submissions.
* **`/Project Folder Sizes`**
    * Scans a master project directory and reports the total disk usage of each project folder. Helps identify large projects for archiving.
* **`/Recently Modified Files`**
    * Scans a project folder to list files changed in the last *X* days. Useful for supervisors to quickly review recent work by the design team.
* **`/Rename Field Photos`**
    * Reads the EXIF "Date Taken" metadata from site photos and batch renames them (e.g., `2025-11-15_14-30_PumpStation.jpg`).
* **`/Sort Downloads`**
    * Moves loose PDF, DWG, and Excel files from your Downloads folder into a staging folder on the network drive for easier filing.

---

### 3. Document Conversion (PowerShell & Python)

* **`/Save to PDF`**
    * **Word to PDF:** A PowerShell script that batch converts all `.doc` and `.docx` files in a folder to PDF using the Microsoft Word engine.
    * **Merge PDFs:** A Python script (requires `pypdf`) that stitches multiple PDF files into a single document (e.g., combining a Specification package).

## How to Run
* **PowerShell:** Right-click the `.ps1` file > "Run with PowerShell".
* **Python:** Open a terminal in the script folder and run `python script_name.py`.
* **Outlook:** Add the macro to your Quick Access Toolbar for one-click archiving.

***

**Note:** These scripts are configured for a standard Windows environment. Check the variables at the top of each script (usually marked `USER: CONFIGURE THIS`) to adjust paths for your specific network environment.
