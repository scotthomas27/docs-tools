# Doc Tools for Google Docs

This Google Apps Script project combines two powerful, separate utilities to enhance your writing and analysis workflow directly within Google Docs. It provides both a "Hemingway editor" style interface for writing feedback and a comprehensive word counter, each with its own dedicated sidebar.

---

## Features

This script adds a single "Doc Tools" menu to your Google Doc with the following functions:

### 1. Hemingway Analyzer
Launches a dedicated sidebar that provides real-time analysis of your document to help you craft clearer and more concise prose. When you run the analyzer, it highlights the text and the sidebar shows insights on:
*   **Readability:** A grade level score based on the Automated Readability Index (ARI).
*   **Adverb Usage:** Highlights adverbs to help you use them sparingly.
*   **Passive Voice:** Flags instances of passive voice.
*   **Complex Phrases:** Identifies overly complex words.
*   **Hard-to-Read Sentences:** Highlights sentences that are long or structurally complex.

### 2. Total Word Counter
Opens a separate sidebar that accurately counts the words from **every tab and subtab** within your Google Doc. This overcomes the limitation of the native Google Docs word counter. The sidebar displays:
*   **Total Word Count:** A sum of all words in the document.
*   **Per-Tab Breakdown:** A list of each top-level tab and its individual word count.

## Instructions for Installation

This setup allows the script to run directly from within a specific Google Document, providing a persistent "Doc Tools" menu for that document.

### **Step 1: Open Your Google Doc and Access Apps Script**

1.  Open the Google Document where you want to use the tools.
2.  In the menu bar, go to **Extensions > Apps Script**.
3.  This will open a new tab with the Google Apps Script editor.

### **Step 2: Create `Code.gs`**

1.  In the Apps Script editor, you should see a file named `Code.gs` by default. If it has any content, delete it.
2.  Replace the contents with the code from the [Code.gs](Code.gs) file in this repository.

### **Step 3: Create `WordCountSidebar.html`**

1.  In the Apps Script editor, click the **`+`** icon next to "Files" and select **HTML**.
2.  Name the new file `WordCountSidebar` (Apps Script will add the `.html` extension).
3.  Replace the default contents with the code from [WordCountSidebar.html](WordCountSidebar.html). This file is for the word count tool's sidebar.

### **Step 4: Create `HemingwaySidebar.html`**

1.  Click the **`+`** icon next to "Files" again and select **HTML**.
2.  Name this new file `HemingwaySidebar`.
3.  Replace the default contents with the code from [HemingwaySidebar.html](HemingwaySidebar.html). This file is for the writing analysis tool's sidebar.

### **Step 5: Create `appsscript.json` (Manifest File)**

1.  In the Apps Script editor, click **Project Settings** (the gear icon ⚙️ on the left sidebar).
2.  Check the box that says **"Show 'appsscript.json' manifest file in editor"**.
3.  Return to the editor (the `< >` icon). You should now see `appsscript.json` in your file list.
4.  Replace its contents with the code from [appsscript.json](appsscript.json).

### **Step 6: Save and Use**

1.  Click the **"Save project"** icon (💾) in the toolbar.
2.  Go back to your Google Doc and **refresh the page**.
3.  A new **"Doc Tools"** menu should now appear.
4.  Click the menu to access the functions:
    *   **Show Word Count:** Opens the sidebar with the total and per-tab word counts.
    *   **Show Hemingway Analysis:** Clears old highlights, analyzes the text, and opens the writing analysis sidebar.
    *   **Clear Analysis Highlights:** A convenience option to remove all highlights without running a new analysis.

The first time you run a function, you will be prompted to authorize the script. Follow the steps: select your Google account, click "Advanced," "Go to [Your Script Name] (unsafe)," and "Allow." This is standard for custom Apps Scripts.

## Credits and Attribution

*   The **Hemingway Helper** functionality is a modified version based on the original code from [SamWSoftware's Projects GitHub repository](https://github.com/SamWSoftware/Projects/tree/master/hemingway).
*   The **Total Word Count** functionality is based on a script authored by "Mr Shane" and shared in the [Google Docs Help community](https://support.google.com/docs/thread/301424533).
