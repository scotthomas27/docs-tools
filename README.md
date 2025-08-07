# Doc Tools for Google Docs

This Google Apps Script project combines two powerful utilities to enhance your writing and analysis workflow directly within Google Docs. It provides both a "Hemingway editor" style interface for real-time writing feedback and a comprehensive word counter that accurately calculates the total word count across documents with tabs and subtabs.

---

## Features

This script adds a single "Doc Tools" menu to your Google Doc with the following functions:

### 1. Hemingway Analyzer
Provides real-time analysis of your document to help you craft clearer and more concise prose. When you run the analyzer, a sidebar appears with insights on:
*   **Readability:** Calculates a grade level score based on the Automated Readability Index (ARI).
*   **Adverb Usage:** Highlights adverbs to help you use them sparingly.
*   **Passive Voice:** Flags instances of passive voice.
*   **Complex Phrases:** Identifies overly complex words and suggests simpler alternatives.
*   **Hard-to-Read Sentences:** Highlights sentences that are long or structurally complex.

### 2. Total Word Counter
Accurately counts the words from **every tab and subtab** within your Google Doc. This overcomes the limitation of the native Google Docs word counter, which only calculates words for the currently active tab. The result is displayed in a simple pop-up alert.

## Instructions for Installation

This setup allows the script to run directly from within a specific Google Document, providing a persistent "Doc Tools" menu for that document.

### **Step 1: Open Your Google Doc and Access Apps Script**

1.  Open the Google Document where you want to use the tools.
2.  In the menu bar, go to **Extensions > Apps Script**.
3.  This will open a new tab with the Google Apps Script editor.

### **Step 2: Create `Code.gs`**

1.  In the Apps Script editor, you should see a file named `Code.gs` by default. If it has any content, delete it.
2.  Replace the contents with the code from the [Code.gs](Code.gs) file in this repository. This file contains the combined logic for both tools.

### **Step 3: Create `SidebarHtml.html`**

1.  In the Apps Script editor, click the **`+`** icon next to "Files" and select **HTML**.
2.  Name the new file `SidebarHtml` (Apps Script will add the `.html` extension).
3.  Replace the default contents with the code from [SidebarHtml.html](SidebarHtml.html). This file is required for the Hemingway Analyzer's sidebar display.

### **Step 4: Create `appsscript.json` (Manifest File)**

1.  In the Apps Script editor, click **Project Settings** (the gear icon ⚙️ on the left sidebar).
2.  Check the box that says **"Show 'appsscript.json' manifest file in editor"**.
3.  Return to the editor (the `< >` icon). You should now see `appsscript.json` in your file list.
4.  Replace its contents with the code from [appsscript.json](appsscript.json).

### **Step 5: Save and Use**

1.  Click the **"Save project"** icon (💾) in the toolbar.
2.  Go back to your Google Doc and **refresh the page**.
3.  A new **"Doc Tools"** menu should now appear.
4.  Click the menu to access the functions:
    *   **Doc Tools > Analyze Document (Hemingway)**
    *   **Doc Tools > Count Words (All Tabs)**

The first time you run a function, you will be prompted to authorize the script. Follow the steps: select your Google account, click "Advanced," "Go to [Your Script Name] (unsafe)," and "Allow." This is standard for custom Apps Scripts.

## Credits and Attribution

*   The **Hemingway Helper** functionality is a modified version based on the original code from [SamWSoftware's Projects GitHub repository](https://github.com/SamWSoftware/Projects/tree/master/hemingway).
*   The **Total Word Count** functionality is based on a script authored by "Mr Shane" and shared in the [Google Docs Help community](https://support.google.com/docs/thread/301424533).
