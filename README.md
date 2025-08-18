# Doc Tools for Google Docs

This Google Apps Script project combines three powerful, separate utilities to enhance your writing and analysis workflow directly within Google Docs. It provides a "Hemingway editor" style interface, a comprehensive word counter, and a writing session logger, each with its own dedicated sidebar.

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

### 3. Writing Habit Tracker
Opens a sidebar that acts as a remote entry form for a Google Sheet, allowing you to log your writing sessions without leaving Google Docs.
*   **Connects to Your Sheet:** A one-time configuration links the tool to your specific tracking spreadsheet.
*   **Manual Data Entry:** Provides a full form to manually enter session data like date, start/stop times, XP, activity type, and notes.
*   **Smart Pre-filling:** Automatically fills in the current date and the document's title to speed up logging.
*   **Formula Safe:** Surgically writes data to your sheet, preserving any existing formula columns.

## Instructions for Installation

This setup allows the script to run directly from within a specific Google Document, providing a persistent "Doc Tools" menu.

### **Step 1: Open Your Google Doc and Access Apps Script**

1.  Open the Google Document where you want to use the tools.
2.  In the menu bar, go to **Extensions > Apps Script**.
3.  This will open a new tab with the Google Apps Script editor.

### **Step 2: Create `Code.gs`**

1.  In the Apps Script editor, you should see a file named `Code.gs` by default. If it has any content, delete it.
2.  Replace the contents with the code from the [Code.gs](Code.gs) file in this repository.

### **Step 3: Create the HTML Sidebar Files**

You will need to create three separate HTML files for each tool's sidebar.

1.  **Create `WordCountSidebar.html`:**
    *   In the Apps Script editor, click the **`+`** icon next to "Files" and select **HTML**.
    *   Name the new file `WordCountSidebar`.
    *   Replace its contents with the code from [WordCountSidebar.html](WordCountSidebar.html).

2.  **Create `HemingwaySidebar.html`:**
    *   Click the **`+`** icon again and select **HTML**.
    *   Name this file `HemingwaySidebar`.
    *   Replace its contents with the code from [HemingwaySidebar.html](HemingwaySidebar.html).

3.  **Create `TrackerSidebar.html`:**
    *   Click the **`+`** icon one more time and select **HTML**.
    *   Name this file `TrackerSidebar`.
    *   Replace its contents with the code from [TrackerSidebar.html](TrackerSidebar.html).

### **Step 4: Create `appsscript.json` (Manifest File)**

1.  In the Apps Script editor, click **Project Settings** (the gear icon ⚙️ on the left sidebar).
2.  Check the box that says **"Show 'appsscript.json' manifest file in editor"**.
3.  Return to the editor (the `< >` icon). You should now see `appsscript.json` in your file list.
4.  Replace its contents with the code from [appsscript.json](appsscript.json). This version includes the necessary permissions to access Google Sheets.

### **Step 5: Save, Configure, and Use**

1.  Click the **"Save project"** icon (💾) in the toolbar.
2.  Go back to your Google Doc and **refresh the page**.
3.  A new **"Doc Tools"** menu should now appear.
4.  **One-Time Tracker Setup:** Before using the tracker for the first time, you must configure it:
    *   Go to **Doc Tools > Writing Tracker > Configure Tracker...**
    *   Follow the two prompts to provide your Google Sheet URL and the exact name of the sheet/tab you want to write to.
5.  **Access the Tools:**
    *   **Total Word Count:** Opens the word count sidebar.
    *   **Hemingway Analysis:** Opens the writing analysis sidebar.
    *   **Clear Highlights:** Removes highlights from the Hemingway tool.
    *   **Writing Tracker > Log Writing Session:** Opens the sidebar to log your work.

The first time you run the tracker, you will be prompted to grant the script new permissions to access Google Sheets. This is a standard and necessary step.

## Credits and Attribution

*   The **Hemingway Helper** functionality is a modified version based on the original code from [SamWSoftware's Projects GitHub repository](https://github.com/SamWSoftware/Projects/tree/master/hemingway).
*   The **Total Word Count** functionality is based on a script authored by "Mr Shane" and shared in the [Google Docs Help community](https://support.google.com/docs/thread/301424533).
