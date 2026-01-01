# OSEM: Outlook Smart Event Manager
> **Programmable, AI-Powered Event Management for Outlook.**
> *Transform your inbox into a structured data center with Local LLM extraction and Python automation.*

---

## Introduction

> **Imagine this:**
>
> You click an email, and it instantly registers in your Quest Log—just like in those classic RPGs we love. It sits there, tracking progress automatically. One glance tells you exactly what to do. You have every key item and piece of intel needed to complete the quest, ready to turn in.

Traditional Outlook feels like a game with broken UI/UX. You pick up an item, and it vanishes into a chaotic inventory. Unless an NPC prompts you again, or you're one of those hardcore players who plays with a walkthrough open or scribbles notes on scraps of paper, that item is lost.

It’s the worst kind of **"Fetch Quest"**: deliver five potatoes here, five sundaes there, all for a "100% Completion" achievement. You stumble upon an NPC in the deep forest demanding an item, forcing you to fast-travel back to town and trek all the way back—for a reward that doesn't even cover the teleportation fee.

Similarly, Outlook forces you to cache scattered emails, attachments, and task statuses in your own brain, constantly context-switching. This **Inbox-based** model is fundamentally anti-human. It silently increases your cognitive cost—especially when juggling twenty tickets that look identical except for a serial number. What you desperately need is a ToDoList packed with organized dossiers.

OSEM's code isn't complex; it does one simple but crucial thing: **Outsourcing memory to the machine.**

By restructuring fragmented emails into structured **"Events"**, OSEM lets Outlook remember "current quest progress" for you, so you don't have to rummage for "where that email is". Whether through local LLM extraction or Python automation, the goal is singular: **To liberate your brain from mechanical memory and retrieval, freeing it for decisions that truly require wisdom, or simply enabling you to enter a flow state to efficiently handle unavoidable tasks, saving that energy for life itself.**

---

## Core Features

### 1. Streamlined Workflow
*   **Drag & Drop Creation:** Create a tracking event by simply dragging an email into the Event Manager.
*   **Context Awareness:** Automatically aggregates all emails from the same conversation thread.
*   **Dedicated File Area:** Each event gets a mapped local folder (configurable repository path) for managing attachments and external files together.
*   **Smart Export:** Supports exporting both ongoing and archived events. Includes intelligent hash verification for template files—only modified files are exported to save space.

### 2. Python Scripting Interface
Extend Outlook's capabilities with standard Python scripts.
*   **Event-Level Automation:** Execute scripts on specific events to process attachments or validate data.
*   **System Integration (ERM Ready):** OSEM actually possesses the skeleton of an **ERM (Enterprise Resource Management)** system. You can even write scripts to extract key data from events and inject it directly into enterprise-grade systems (like **CargoWise One***, SAP) via API or XML, bridging the last mile from "Email" to "Core Business System".

    > *\* CargoWise One is a trademark of WiseTech Global. This project is not affiliated with WiseTech Global.*

### 3. Smart Extraction (AI & Regex)
*   **Local LLM Support:** Integrated with **Ollama**. Use models like Llama3 or DeepSeek locally to parse email content.
*   **Regex Support:** Use efficient Regular Expressions for precise extraction from structured documents.
*   **Data Privacy:** Your email data is processed locally and never sent to external cloud APIs (unless you configure it to).

---

## Workflow in Action

### 1. Define Template
Configure the fields you need to track (e.g., `OrderNo`, `ETA`, `Customer`) to fit your specific workflow.

### 2. Capture: Drag & Drop
Simply drag an email into OSEM to initialize an event. It automatically aggregates the conversation thread.
![Create Event](Docs/Images/Merge.gif)

### 3. Extract: AI Powered
Use "AI Extract" with local LLM to auto-fill the dashboard fields instantly.
![AI Extraction](Docs/Images/llm.gif)

### 4. Export & Automate
Run Python scripts to handle files or export data with smart hash verification.
![Smart Export](Docs/Images/Export.gif)

---

## Use Cases

| Scenario | Traditional Workflow | OSEM Workflow |
| :--- | :--- | :--- |
| **Logistics** | Search emails -> Find PDF -> Check ETA | **Dashboard View** shows ETA instantly; Attachments auto-archived. |
| **Sales** | Manually tracking client requirements | **AI Summarization** of client needs; Priority tagging. |
| **Operations** | Copy-pasting error codes to Excel | **Regex/AI Extraction** of error codes; Script auto-logs to system. |

---

## Performance & Best Practices
*   **Optimization:** OSEM is designed to yield to Outlook's main thread to prevent UI freezing.
*   **Startup Buffer:** When Outlook is syncing large volumes of mail (e.g., Monday mornings), it is recommended to allow 1-2 minutes for initialization before performing heavy batch operations.
*   **Responsiveness:** Please note that Outlook may experience a slight decrease in responsiveness compared to normal usage, depending on your hardware specifications and available memory.

---

## Getting Started

### Prerequisites
*   Windows 10/11
*   Outlook Desktop
*   .NET Framework 4.8+
*   (Optional) [Ollama](https://ollama.com/) for AI features
*   (Optional) Python 3.x for scripting

### Installation

#### Option 1: End User (Installer)
1.  Download the latest installer from the **Releases** page.
2.  Unzip and run `setup.exe`.
3.  Launch Outlook and click **"Event Manager"** in the ribbon.

#### Option 2: Developer (Build from Source)
1.  Clone this repository.
2.  Open `OSEMAddIn.sln` in Visual Studio 2022 (with "Office/SharePoint development" workload installed).
3.  Build the solution (Ctrl+Shift+B).
4.  Press **F5** to run and debug directly in Outlook.

### Uninstallation
1.  Go to Windows **Settings** > **Apps** > **Installed apps**.
2.  Search for "OSEM".
3.  Click the three dots menu and select **Uninstall**.

> 💡 **Documentation:** For detailed installation, configuration, and usage instructions, please refer to the [User Guide](Docs/UserGuide_EN.md).
>
> 🐍 **Scripting:** To extend functionality with Python scripts, please refer to the [Scripting Interface](Docs/ScriptingInterface_EN.md).

---

## Contributing
OSEM is built with C# (VSTO) and WPF. We welcome contributions to improve the core logic or add new script examples.

> **Special Note:**
> While the code in this project was generated by AI, **core functionalities have been rigorously reviewed and battle-tested by human developers** to ensure reliability in production environments.

---

## Support
If OSEM has truly helped you, giving you the time to finally finish that game you've been putting off, or if you simply find this "Bracer Notebook" useful, feel free to buy me a coffee.

<!-- 
Replace the links below with your own sponsorship pages:
1. Ko-fi: https://ko-fi.com/elysionlhant
2. Afdian: https://afdian.com/a/elysionlhant
-->
<div align="left">
  <!-- Ko-fi (International/PayPal) -->
  <a href="https://ko-fi.com/elysionlhant" target="_blank">
    <img src="https://storage.ko-fi.com/cdn/kofi2.png" alt="Buy Me a Coffee at ko-fi.com" height="50" >
  </a>
  
  <!-- Afdian (Chinese Users/WeChat/Alipay) -->
  <a href="https://afdian.com/a/elysionlhant" target="_blank" style="margin-left: 20px;">
    <img src="https://pic1.afdiancdn.com/static/img/logo/logo.png" alt="Afdian" height="50" >
  </a>
</div>

---

> "I hope the cognitive cost saved here allows you to return home and read a few pages of a book, to reserve attention for the people who truly matter, to write the lines you've been meaning to write, to simply daydream instead of being harvested by short videos, and to protect your possibilities from being stolen when you are most vulnerable.
>
> If this is achieved, then I have fulfilled my purpose."

