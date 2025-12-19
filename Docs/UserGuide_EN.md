# OSEM (Outlook Smart Event Manager) User Guide

## Introduction
OSEM (Outlook Smart Event Manager) is a powerful VSTO Add-in for Microsoft Outlook designed to streamline event-based email management. It transforms your inbox into a structured event tracking system, allowing you to group related emails, extract key information using LLM (Large Language Models) or Regex, and automate workflows with Python scripts.

**Key Features:**
*   **Event-Centric View:** Automatically groups related emails into "Events" based on subject patterns.
*   **Intelligent Extraction:** Use local LLMs (via Ollama) or Regex to extract structured data (Dashboard) from email bodies.
*   **Template System:** Define custom data structures (Dashboards) and email templates for different types of workflows (e.g., Logistics, Sales).
*   **Scripting Automation:** Run external Python scripts to process event data and attachments.
*   **Report Export:** Export event data, attachments, and files to local folders for archiving and reporting.
*   **Localization:** Fully supports English and Simplified Chinese (auto-detects system language).

---

## Installation & Setup
1.  **Prerequisites:**
    *   Microsoft Outlook (Desktop Version).
    *   .NET Framework 4.8 or later.
    *   (Optional) [Ollama](https://ollama.com/) for local LLM features.
    *   (Optional) Python environment for scripting features.
2.  **Installation:**
    *   **End User:** Run the `setup.exe` from the release package.
    *   **Developer:** Open `OSEMAddIn.sln` in Visual Studio, Build, and Run (F5).
    *   Open Outlook. You should see the "Event Manager" button in the Ribbon or a new Task Pane.
3.  **Uninstallation:**
    *   Go to Windows **Settings** > **Apps** > **Installed apps**.
    *   Search for "OSEM".
    *   Click the three dots menu and select **Uninstall**.

---

## Usage Tips
Since this add-in is deeply integrated into the Outlook environment, it is recommended to allow 1-2 minutes of buffer time shortly after launching Outlook (especially on Monday mornings when syncing large volumes of emails) for it to complete its initialization and data synchronization.

Although the add-in is optimized to yield to the main thread, Outlook's main UI thread can be sensitive under high load. To prevent temporary UI freezes caused by rendering blocks, please avoid rapid consecutive clicks or complex operations until the Outlook interface is fully responsive.

Additionally, please note that Outlook may experience a slight decrease in responsiveness compared to normal usage, depending on your hardware specifications and available memory.

---

## Main Interface: Event Manager
The **Event Manager** is the central hub for tracking your active workflows.

### 1. Event List
*   **Active Events:** Shows currently open events.
    *   **Columns:** Event Title, Last Updated, Info (Custom summary), Priority (Star icon).
    *   **Actions:**
        *   **Create Event (Drag & Drop):** Drag and drop a **single** email directly from Outlook into this list to instantly create a new event.
        *   **Double-click:** Opens the **Event Detail View**.
        *   **Right-click:** Options to Rename, Archive, or Delete the event.
        *   **Toggle Priority:** Click the star icon to mark important events.
*   **Completed Events:** Shows archived events. You can reopen them from here.

### 2. Toolbar Controls
*   **Search:** Filter events by title or content.
*   **Filter Template:** Show only events associated with a specific template.
*   **Refresh All:** Reloads all events from the monitored folders.
*   **Template Rules:** Configure auto-assignment of templates based on email participants.
*   **Template Editor:** Open the configuration center (Templates, Prompts, Scripts).
*   **Complete Selected:** Archive the selected event(s).
*   **Run Script:** Execute a global Python script on the selected event.
*   **Export:** Open the Export Options window to batch export event data.

---

## Event Detail View
Double-clicking an event opens the **Event Detail View**, where you process specific tasks.

### 1. Dashboard (Left Panel)
*   **Status:** Shows the current status of the event.
*   **Template:** Select the data structure (Template) for this event.
*   **Info Fields:** Key-value pairs displayed in the event details. These are fully editable.
    *   **LLM Extraction:** Click **"Run LLM Extraction"** to use AI to fill these fields automatically based on the selected email.
    *   **Regex Extraction:** Click **"Run Regex Extraction"** to use pre-defined patterns.
    *   **Copy:** One-click to copy all non-empty fields to the clipboard in JSON format, ready for pasting into other systems.

### 2. Email Pool (Middle Panel)
*   Displays all emails grouped under this event.
*   **Actions:**
    *   **Add Email (Drag & Drop):** Drag and drop **one or multiple** emails from Outlook into this area to instantly add them to the current event.
    *   **Refresh Emails:** Automatically retrieves and adds all related emails belonging to the same conversation, ensuring full context and access to all attachments.
    *   **Select & Preview:** Click an email to use it as the source for extraction. When the email preview pane is active, clicking an email in the pool will directly preview its content.
    *   **Remove:** Remove an unrelated email from this event.
    *   **Mark Read:** Mark selected emails as read.

### 3. Attachments & Files (Right Panel)
*   **Attachments:** Lists all attachments found in the event's emails.
*   **File Area:** A dedicated local folder for this event (independent for each event).
    *   **Interaction:** Supports **dragging and dropping files** directly from the "Attachments" list, Desktop, or File Explorer into this area.
    *   **Generate Folder:** Creates a local folder for the event.
    *   **Update to Folder:** Saves selected email attachments to this local folder.

---

## Report Export (Export Options)
Accessed via the "Export" button in the main interface. This feature allows you to batch export event data to your local file system.

*   **Selection Scope:**
    *   **Select Template:** Export only specific types of events (e.g., only "Logistics Orders").
    *   **Event Range:** Choose to export events within a specific date range.
*   **Export Content:**
    *   **Event Dashboard Data:** Export an Excel/CSV table containing all extracted fields (e.g., InvoiceNo, ETA).
    *   **Event Attachments:** Export raw attachments from emails.
    *   **Event Files:** Export local files from the "File Area".
*   **File Type Filter:** Export only specific file formats (e.g., only PDF or Excel).
*   **Folder Naming:** Customize the naming convention for export folders (e.g., `[Date]_[EventID]`).

---

## Configuration: Template & Config Editor
Access this via the "Template Editor" button in the main window.

### 1. Dashboard Templates
Define the structure of data you want to track.
*   **Fields:** Add keys (e.g., "InvoiceNo", "ETA") that you want to extract.
*   **Regex:** Optionally bind a Regex pattern to a field for rule-based extraction.
*   **Common Files:** Attach standard files (e.g., forms, checklists) to a template. When used with "Template Rules" for auto-assignment, these files are automatically copied to the event's "File Area", enabling automatic workflow file provisioning.

### 2. Prompt Management
Manage prompts sent to the LLM.
*   **Variables:** Use placeholders like `{{MAIL_BODY}}` or `{{DASHBOARD_JSON}}` to inject dynamic context.
*   **Association:** Link a prompt to a specific Template so it only appears for relevant events.

### 3. Script Management
Register Python scripts for automation.
*   **Global Scripts:** Can be run from the main list (e.g., "Export Daily Report").
*   **Event Scripts:** Run inside a specific event (e.g., "Extract PDF Data").
*   **Context:** Scripts receive a JSON file containing all event data (Emails, Attachments, Dashboard values).

### 4. LLM Settings
Configure your AI provider.
*   **Provider:** Supports **Ollama Local** and **HTTP API** (OpenAI compatible).
*   **Ollama Mode:** Automatically fetches the list of locally installed Ollama models for selection.
*   **API Mode:** Supports custom Endpoint and API Key, allowing connection to any OpenAI-compatible service.
*   **Scope:** Set global defaults or override settings for specific templates.

### 5. Monitored Folders
By default, OSEM monitors `Inbox` and `Sent Items`. You can add other Outlook folders here to include their emails in event grouping.

---

## Backup & Migration
Located in the **Export/Import** tab of the Template Editor.
*   **Export Backup:** Saves all your settings (Templates, Prompts, Scripts, Configs) and the Event List database to a `.zip` file.
*   **Import Backup:** Restores data from a backup. Supports "Merge Mode" to resolve conflicts between existing and imported data.

---

## Localization
The UI language automatically adapts to your Windows System Language.
*   **English:** Default.
*   **Simplified Chinese:** Activates automatically on Chinese systems.
