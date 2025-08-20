// --- WRITER'S HIGHLIGHTER FUNCTIONS ---

/**
 * Controller function to show the Highlighter sidebar.
 */
function showHighlighterSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('HighlighterSidebar.html')
      .setTitle("Word Highlighter");
  DocumentApp.getUi().showSidebar(html);
}

/**
 * Saves the user's highlighter rules and options.
 * @param {object} settings The settings object from the sidebar.
 */
function saveHighlighterSettings(settings) {
  PropertiesService.getDocumentProperties().setProperty('highlighterSettings', JSON.stringify(settings));
}

/**
 * Loads the user's saved highlighter rules and options.
 * @returns {object} The saved settings object.
 */
function loadHighlighterSettings() {
  const settingsString = PropertiesService.getDocumentProperties().getProperty('highlighterSettings');
  return settingsString ? JSON.parse(settingsString) : { rules: [{word: '', color: '#ffff00'}], options: {wholeWords: true} };
}

/**
 * Applies highlights and styles using a reliable manual search loop.
 * This avoids the problematic findText(pattern, from) loop.
 */
function applyHighlights(rules, options, selectionOnly) {
  const doc = DocumentApp.getActiveDocument();
  let elementsToSearch;

  // Determine the scope of our search
  if (selectionOnly) {
    const selection = doc.getSelection();
    if (!selection) {
      DocumentApp.getUi().alert('Please select some text to highlight.');
      return;
    }
    elementsToSearch = selection.getRangeElements().map(rangeEl => rangeEl.getElement());
  } else {
    elementsToSearch = [doc.getBody()];
  }
  
  // Define the styles to apply from the checkboxes
  const style = {};
  style[DocumentApp.Attribute.BOLD] = options.bold || null;
  style[DocumentApp.Attribute.ITALIC] = options.italic || null;
  style[DocumentApp.Attribute.UNDERLINE] = options.underline || null;
  style[DocumentApp.Attribute.STRIKETHROUGH] = options.strikethrough || null;

  // Iterate through each rule the user has defined
  rules.forEach(rule => {
    if (!rule.word) return; // Skip empty rules

    // Sanitize the user's word to be safe in a regular expression
    const sanitizedWord = rule.word.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const searchPattern = options.wholeWords ? '\\b' + sanitizedWord + '\\b' : sanitizedWord;
    const regExp = new RegExp(searchPattern, options.ignoreCase ? 'gi' : 'g');
    
    // Process each element in our search scope (the whole body or the selection)
    elementsToSearch.forEach(element => {
      // We can only search inside elements that contain text
      if (element.asText) {
        const textElement = element.asText();
        const text = textElement.getText();
        let match;
        
        // Use a standard, reliable JavaScript regex loop to find ALL matches
        while ((match = regExp.exec(text)) !== null) {
          const start = match.index;
          const end = start + match[0].length - 1;
          
          if (start <= end) {
            // Apply the background color and text styles for each match
            textElement.setBackgroundColor(start, end, rule.color);
            textElement.setAttributes(start, end, style);
          }
        }
      }
    });
  });
}
