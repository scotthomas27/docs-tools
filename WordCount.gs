// --- WORD COUNT FUNCTIONS ---

/**
 * Controller: Calculates word counts and displays them in their own dedicated sidebar.
 */
function showWordCountSidebar() {
  const wordCountData = getWordCountData(); // Get the data
  const htmlTemplate = HtmlService.createTemplateFromFile('WordCountSidebar.html'); // Use the new HTML file
  htmlTemplate.data = wordCountData; // Pass data to the template
  const htmlOutput = htmlTemplate.evaluate().setTitle('Word Count');
  DocumentApp.getUi().showSidebar(htmlOutput);
}

/**
 * Worker: Calculates total word count and a per-tab breakdown.
 * @returns {object} An object with totalWords and an array of tabDetails.
 */
function getWordCountData() {
  const doc = DocumentApp.getActiveDocument();
  const tabs = doc.getTabs();
  let totalWords = 0;
  const tabDetails = [];

  for (let i = 0; i < tabs.length; i++) {
    let tabWordCount = 0;
    const stack = [tabs[i]];
    const mainTabName = `Tab ${i + 1}`;

    while (stack.length > 0) {
      const currentTab = stack.pop();
      const documentTab = currentTab.asDocumentTab();
      const body = documentTab.getBody();
      const text = body.getText();
      if (text.trim()) {
          const words = text.trim().replace(/\s+/g, ' ').split(' ');
          tabWordCount += words.length;
      }
      const childTabs = currentTab.getChildTabs();
      for (let j = 0; j < childTabs.length; j++) {
        stack.push(childTabs[j]);
      }
    }
    totalWords += tabWordCount;
    tabDetails.push({ name: mainTabName, count: tabWordCount });
  }

  return { totalWords: totalWords, tabDetails: tabDetails };
}
