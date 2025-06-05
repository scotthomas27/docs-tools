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
2.  Replace the entire content of `Code.gs` with the code below:

    ```javascript
    // Helper function to calculate readability level
    function calculateLevel(letters, words, sentences) {
      if (words === 0 || sentences === 0) {
        return 0;
      }
      let level = Math.round(
        4.71 * (letters / words) + 0.5 * (words / sentences) - 21.43
      );
      return level <= 0 ? 0 : level;
    }

    // List of words that are not adverbs despite ending in "ly"
    function getLyWords() {
      return {
        actually: 1, additionally: 1, allegedly: 1, ally: 1, alternatively: 1,
        anomaly: 1, apply: 1, approximately: 1, ashely: 1, ashly: 1,
        assembly: 1, awfully: 1, baily: 1, belly: 1, bely: 1, billy: 1,
        bradly: 1, bristly: 1, bubbly: 1, bully: 1, burly: 1, busily: 1,
        calamity: 1, cavalry: 1, chilly: 1, comply: 1, contently: 1, costly: 1,
        courtly: 1, craftily: 1, crumbly: 1, cuddly: 1, daily: 1, deadly: 1,
        dearly: 1, deathly: 1, deny: 1, destiny: 1, daily: 1, disorderly: 1,
        display: 1, doily: 1, dolly: 1, doubly: 1, dragonfly: 1, dreary: 1,
        early: 1, easily: 1, elderly: 1, elly: 1, empty: 1, enmity: 1,
        especially: 1, exactly: 1, externally: 1, family: 1, folly: 1,
        friendly: 1, frilly: 1, ghastly: 1, godly: 1, grizzly: 1, grisly: 1,
        haily: 1, handy: 1, happily: 1, hardily: 1, heavenly: 1, hilly: 1,
        holy: 1, hourly: 1, humbly: 1, idly: 1, impurity: 1, infantry: 1,
        instantly: 1, insanity: 1, jely: 1, jolly: 1, kelly: 1, kindly: 1,
        lately: 1, leisurely: 1, letely: 1, likely: 1, lively: 1, lonely: 1,
        lovely: 1, lowly: 1, loyalty: 1, lucky: 1, manly: 1, marry: 1,
        masterly: 1, melancholy: 1, monthly: 1, multiply: 1, mystery: 1,
        nasty: 1, natively: 1, naturally: 1, nearby: 1, neatly: 1, needy: 1,
        only: 1, orderly: 1, oily: 1, paly: 1, panoply: 1, partly: 1,
        pastely: 1, pearly: 1, penalty: 1, plenty: 1, poily: 1, polity: 1,
        poorly: 1, portly: 1, possibly: 1, poultry: 1, prickly: 1, primary: 1,
        princely: 1, promptly: 1, properly: 1, purify: 1, quarry: 1, queenly: 1,
        quickly: 1, rarity: 1, readily: 1, really: 1, rely: 1, reply: 1,
        responsibly: 1, royalty: 1, rudely: 1, saintly: 1, scaly: 1, scilly: 1,
        seemly: 1, shabby: 1, shapely: 1, simply: 1, sickly: 1, silly: 1,
        singly: 1, skilly: 1, slippery: 1, slyly: 1, smelly: 1, sparkly: 1,
        spindly: 1, springly: 1, spryly: 1, stately: 1, steely: 1, sticky: 1,
        stilly: 1, straggly: 1, sultry: 1, surely: 1, tally: 1, timely: 1,
        timidly: 1, tiny: 1, truly: 1, ugly: 1, unruly: 1, usally: 1,
        usually: 1, vary: 1, vastly: 1, venally: 1, very: 1, villany: 1,
        virulently: 1, wily: 1, worldly: 1, woolly: 1, worldly: 1, yearly: 1,
        yelly: 1, zany: 1,
      };
    }

    // Function to analyze the document
    function runAnalysis() {
      let doc = DocumentApp.getActiveDocument();
      let body = doc.getBody();
      let text = body.getText();

      // Basic stats
      let sentences = text.split(/[.!?]+/).filter(Boolean).length;
      let words = text.split(/\s+/).filter(Boolean).length;
      let letters = text.replace(/[^a-zA-Z]/g, '').length;

      let ari = calculateLevel(letters, words, sentences);

      // Analyze adverbs
      let adverbs = [];
      let adverbPattern = /\b\w+ly\b/gi;
      let lyWords = getLyWords();
      let match;
      while ((match = adverbPattern.exec(text)) !== null) {
        if (!lyWords[match[0].toLowerCase()]) {
          adverbs.push(match[0]);
        }
      }

      // Analyze passive voice
      let passiveSentences = [];
      let passivePattern = /\b(am|is|are|was|were|be|being|been)\s+\w+ed\b/gi;
      let sentencesArray = text.split(/[.!?]+/);
      sentencesArray.forEach(sentence => {
        if (passivePattern.test(sentence)) {
          passiveSentences.push(sentence.trim());
        }
      });

      // Analyze complex words
      let complexWords = [];
      let wordPattern = /\b\w+\b/g;
      let complexCount = 0;
      while ((match = wordPattern.exec(text)) !== null) {
        let word = match[0].toLowerCase();
        if (word.length >= 7 && !isCommonWord(word)) { // A simplified check
          complexWords.push(word);
          complexCount++;
        }
      }

      // Very simplified common word check (you might want a more robust list)
      function isCommonWord(word) {
        const commonWords = new Set(['the', 'and', 'a', 'to', 'of', 'in', 'is', 'it', 'that', 'he', 'was', 'for', 'on', 'are', 'with', 'as', 'i', 'his', 'they', 'at', 'be', 'this', 'had', 'have', 'from', 'or', 'one', 'by', 'but', 'not', 'what', 'all', 'were', 'we', 'when', 'your', 'can', 'said', 'there', 'use', 'an', 'each', 'which', 'she', 'do', 'how', 'their', 'if', 'will', 'up', 'other', 'about', 'out', 'many', 'then', 'them', 'these', 'so', 'some', 'her', 'would', 'make', 'like', 'him', 'into', 'time', 'has', 'look', 'two', 'more', 'write', 'go', 'see', 'number', 'no', 'way', 'could', 'people', 'my', 'than', 'first', 'water', 'been', 'call', 'who', 'oil', 'sit', 'now', 'find', 'long', 'down', 'day', 'did', 'get', 'come', 'made', 'may', 'part']);
        return commonWords.has(word);
      }

      // Prepare HTML for sidebar
      let htmlOutput = HtmlService.createTemplateFromFile('SidebarHtml');
      htmlOutput.adverbs = adverbs;
      htmlOutput.passiveSentences = passiveSentences;
      htmlOutput.complexWords = complexWords;
      htmlOutput.ari = ari;
      htmlOutput.wordCount = words;
      htmlOutput.sentenceCount = sentences;

      let ui = HtmlService.createHtmlOutput(htmlOutput.evaluate())
        .setTitle('Hemingway Helper Analysis')
        .setWidth(300);
      DocumentApp.getUi().showSidebar(ui);
    }

    // Function that runs when the add-on is opened or installed
    function onOpen(e) {
      DocumentApp.getUi()
          .createMenu('Hemingway Helper')
          .addItem('Analyze Document', 'runAnalysis')
          .addToUi();
    }

    // Minimal onHomepage function required by manifest, does not need to do anything for this menu-driven add-on
    function onHomepage(e) {
      return HtmlService.createHtmlOutput('<p>Click "Hemingway Helper" > "Analyze Document" in the menu.</p>');
    }
    ```

### **Step 3: Create `SidebarHtml.html`**

1.  In the Apps Script editor, click **File > New > HTML file**.
2.  Name the new file `SidebarHtml`.
3.  Replace the entire content of `SidebarHtml.html` with the code below:

    ```html
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body { font-family: sans-serif; padding: 10px; }
          h3 { margin-top: 20px; }
          ul { list-style-type: none; padding: 0; }
          li { margin-bottom: 5px; }
          .section-title { font-weight: bold; margin-bottom: 5px; }
        </style>
      </head>
      <body>
        <h2>Hemingway Helper Analysis</h2>
        <p><strong>Word Count:</strong> <?= wordCount ?></p>
        <p><strong>Sentence Count:</strong> <?= sentenceCount ?></p>
        <p><strong>Readability Grade Level (ARI):</strong> <?= ari ?></p>

        <h3>Adverbs (<?= adverbs.length ?>)</h3>
        <? if (adverbs.length > 0) { ?>
          <ul>
            <? for (let i = 0; i < adverbs.length; i++) { ?>
              <li><?= adverbs[i] ?></li>
            <? } ?>
          </ul>
        <? } else { ?>
          <p>No adverbs found.</p>
        <? } ?>

        <h3>Passive Voice Sentences (<?= passiveSentences.length ?>)</h3>
        <? if (passiveSentences.length > 0) { ?>
          <ul>
            <? for (let i = 0; i < passiveSentences.length; i++) { ?>
              <li><?= passiveSentences[i] ?></li>
            <? } ?>
          </ul>
        <? } else { ?>
          <p>No passive voice found.</p>
        <? } ?>

        <h3>Complex Words (<?= complexWords.length ?>)</h3>
        <? if (complexWords.length > 0) { ?>
          <ul>
            <? for (let i = 0; i < complexWords.length; i++) { ?>
              <li><?= complexWords[i] ?></li>
            <? } ?>
          </ul>
        <? } else { ?>
          <p>No complex words found.</p>
        <? } ?>

      </body>
    </html>
    ```

### **Step 4: Create `appsscript.json` (Manifest File)**

1.  In the Apps Script editor, click **Project Settings** (the gear icon ⚙️ on the left sidebar).
2.  Scroll down and check the box that says **"Show 'appsscript.json' manifest file in editor"**.
3.  Go back to the editor (the code icon `< >` on the left sidebar). You should now see `appsscript.json` in your file list.
4.  Click on `appsscript.json` and replace its entire content with the code below:

    ```json
    {
      "timeZone": "Australia/Brisbane",
      "dependencies": {},
      "exceptionLogging": "STACKDRIVER",
      "runtimeVersion": "V8",
      "oauthScopes": [
        "[https://www.googleapis.com/auth/documents.currentonly](https://www.googleapis.com/auth/documents.currentonly)",
        "[https://www.googleapis.com/auth/script.container.ui](https://www.googleapis.com/auth/script.container.ui)"
      ],
      "addOns": {
        "common": {
          "name": "Hemingway Helper",
          "logoUrl": "[https://www.gstatic.com/images/branding/product/2x/apps_script_64dp.png](https://www.gstatic.com/images/branding/product/2x/apps_script_64dp.png)",
          "homepageTrigger": {
            "runFunction": "onHomepage",
            "enabled": false
          }
        },
        "docs": {
          "onFileScopeGrantedTrigger": {
            "runFunction": "onOpen"
          }
        }
      }
    }
    ```

### **Step 5: Save and Run `onOpen()`**

1.  Click the **"Save project"** icon (💾) in the toolbar.
2.  Go back to your Google Doc.
3.  Refresh the Google Doc page.
4.  In the Google Doc, go to **Extensions > Hemingway Helper > Analyze Document**.

The first time you run it in a document, you will likely be prompted to authorize the script. Follow the authorization steps: select your Google account, click "Advanced," "Go to [Your Script Name] (unsafe)," and "Allow."

After authorization, the "Hemingway Helper" menu item should appear under "Extensions", and selecting "Analyze Document" should open the sidebar with the analysis.
