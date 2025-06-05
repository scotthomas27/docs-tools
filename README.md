# Hemingway Helper for Google Docs

Hemingway Helper is a Google Apps Script-based "Hemingway editor interface" designed to enhance your writing within Google Docs. It provides real-time analysis of your document, offering insights into readability, adverb usage, passive voice, and complex words to help you craft clearer and more concise prose. This project is a modified version based on the original code sourced from [SamWSoftware's Projects GitHub repository](https://github.com/SamWSoftware/Projects/tree/master/hemingway).

---

## Instructions for Creating the Document-Specific Version of the App

This setup allows the "Hemingway Helper" to run directly from within a specific Google Document, providing a persistent interface for that document.

### **Step 1: Open Your Google Doc and Access Apps Script**

1.  Open the Google Document where you want to use the Hemingway Helper.
2.  In the Google Doc menu bar, go to **Extensions > Apps Script**.
3.  This will open a new tab with the Google Apps Script editor.

### **Step 2: Create `Code.gs`**

1.  In the Apps Script editor, you should see a file named `Code.gs` by default. If not, click **File > New > Script file**.
2.  Replace the default contents with [Code.gs](Code.gs).

### **Step 3: Create `SidebarHtml.html`**

1.  In the Apps Script editor, click **File > New > HTML file**.
2.  Name the new file `SidebarHtml`.
3.  Replace the default contents with [SidebarHtml.html](SidebarHtml.html).

### **Step 4: Create `appsscript.json` (Manifest File)**

1.  In the Apps Script editor, click **Project Settings** (the gear icon ⚙️ on the left sidebar).
2.  Scroll down and check the box that says **"Show 'appsscript.json' manifest file in editor"**.
3.  Go back to the editor (the code icon `< >` on the left sidebar). You should now see `appsscript.json` in your file list.
4.  Replace the default contents with [appsscript.json](appsscript.json).

### **Step 5: Save and Run**

1.  Click the **"Save project"** icon (💾) in the toolbar.
2.  Go back to your Google Doc.
3.  Refresh the Google Doc page.
4.  In the Google Doc, go to **Extensions > Hemingway Helper > Analyze Document**.

The first time you run it in a document, you will likely be prompted to authorize the script. Follow the authorization steps: select your Google account, click "Advanced," "Go to [Your Script Name] (unsafe)," and "Allow."

After authorization, the "Hemingway Helper" menu item should appear, and selecting "Analyze Document" should open the sidebar with the analysis.
