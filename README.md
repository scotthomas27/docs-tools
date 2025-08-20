# Writer's Multitool for Google Docs

This Google Apps Script project combines four powerful, separate utilities to enhance your writing and analysis workflow directly within Google Docs. It provides a "Hemingway editor" style interface, a comprehensive word counter, a custom word highlighter, and a highly configurable writing session logger, all accessible from a single "Multitool" menu.

---

## Features

This script adds a single "Multitool" menu to your Google Doc with the following functions:

### 1. Hemingway Tool
Launches a dedicated sidebar that provides real-time analysis of your document to help you craft clearer and more concise prose.
*   **Analysis:** Highlights hard-to-read sentences, adverbs, passive voice, and complex phrases.
*   **Interactive Suggestions:** Click inside a highlighted complex phrase and click the "Get Suggestion" button in the sidebar to see simpler alternatives.

### 2. Word Count Tool
Opens a separate sidebar that accurately counts the words from **every tab and subtab** within your Google Doc.
*   **Total Word Count:** A sum of all words in the document.
*   **Per-Tab Breakdown:** A list of each top-level tab and its individual word count.

### 3. Word Highlighter Tool
Launches a sidebar that allows you to find and apply custom styling to a list of words or phrases.
*   **Custom Rules:** Create a list of words and assign a unique highlight color to each.
*   **Advanced Styling:** Apply additional formatting like bold, italic, underline, and strikethrough.
*   **Scoped Application:** Highlight words in the entire document or only within your current text selection.
*   **Save & Load:** Save your list of rules and settings on a per-document basis.

### 4. Writing Tracker Tool
A highly configurable tool to log your writing sessions directly to a Google Sheet without leaving your document.
*   **Connects to Your Sheet:** A one-time configuration links the tool to your specific tracking spreadsheet.
*   **Fully Customizable Sidebar:** A setup assistant reads your spreadsheet's column headers and lets you build your own data-entry sidebar. You choose which columns to include, what type of input they use (text, number, dropdown, etc.), and their display order.
*   **Stateful Form:** Remembers your data if you close the sidebar mid-session and pre-fills "persistent" fields (like a bonus multiplier) with the last value from your sheet.
*   **Formula Safe:** Surgically writes data to the correct row in your sheet, preserving any existing formula columns.

## Installation Instructions

This setup allows the script to run directly from within a specific Google Document. Because the project is organized into multiple files for clarity, the installation requires a few steps.

### **Step 1: Open Your Google Doc and Access Apps Script**

1.  Open the Google Document where you want to use the tools.
2.  In the menu bar, go to **Extensions > Apps Script**. This will open the script editor.

### **Step 2: Create and Populate All Project Files**

The key is to create a new script or HTML file in the Apps Script editor for each corresponding file in this repository, then copy the contents over.

1.  In the Apps Script editor, you will see a default `Code.gs` file.
2.  For each file in this repository (e.g., `Tracker.gs`, `HemingwaySidebar.html`), click the **`+`** icon next to "Files" in the editor, choose "Script" (for `.gs` files) or "HTML" (for `.html` files), and give it the **exact same name**.
3.  Copy the code from the repository file and paste it into the corresponding file you just created in the editor, overwriting any default content.

You will need to create and populate the following files:
*   **Script Files (`.gs`):**
    *   `Code.gs`
    *   `Hemingway.gs`
    *   `Highlighter.gs`
    *   `Tracker.gs`
    *   `WordCount.gs`
*   **HTML Files (`.html`):**
    *   `ConfigSidebar.html`
    *   `HemingwaySidebar.html`
    *   `HighlighterSidebar.html`
    *   `TrackerSidebar.html`
    *   `TrackerTemplate.html`
    *   `WordCountSidebar.html`

### **Step 3: Create the `appsscript.json` Manifest File**

1.  In the Apps Script editor, click **Project Settings** (the gear icon ⚙️).
2.  Check the box that says **"Show 'appsscript.json' manifest file in editor"**.
3.  Return to the editor (`< >` icon). You should now see `appsscript.json` in your file list.
4.  Replace its contents with the code from the [appsscript.json](appsscript.json) file in this repository. This includes the necessary permissions to access Google Sheets.

### **Step 4: Save, Configure, and Use**

1.  Click the **"Save project"** icon (💾) in the toolbar.
2.  Go back to your Google Doc and **refresh the page**. A new **"Multitool"** menu should appear.
3.  **One-Time Tracker Setup:** Before using the tracker, you must configure it:
    *   Go to **Multitool > Writing Tracker Tool > Configure Tracker...**
    *   Follow the prompts to provide your Google Sheet URL and the exact sheet/tab name.
    *   Use the configuration sidebar to map your spreadsheet columns to the tool's field types.
4.  **Access the Tools:** Use the "Multitool" menu to launch any of the tools.

The first time you run the tracker, you will be prompted to grant the script permissions to access Google Sheets. This is a standard and necessary step.

## Credits and Attribution

*   The **Hemingway Tool** functionality is a modified version based on the original code from [SamWSoftware's Projects GitHub repository](https://github.com/SamWSoftware/Projects/tree/master/hemingway).
*   The **Total Word Count** functionality is based on a script authored by "Mr Shane" and shared in the [Google Docs Help community](https://support.google.com/docs/thread/301424533).
