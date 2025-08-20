/**
 * WRITER'S MULTITOOL - MAIN SCRIPT
 * This file contains the global definitions and the onOpen function
 * that builds the main menu for the application.
 * All other functionality is delegated to separate .gs files.
 */

// --- GLOBAL DECLARATIONS ---

// This object holds the analysis data for the Hemingway tool.
// It is kept global because the tool processes the document in parts.
let analysisData = {};

// This constant object holds the color codes for the Hemingway tool's highlights.
const HIGHLIGHT_COLORS = {
  hardSentence: '#FFFF00',    // Yellow
  veryHardSentence: '#FFC0CB',// Pinkish-red
  adverb: '#ADD8E6',         // Light Blue
  passive: '#90EE90',        // Light Green
  complex: '#E6E6FA',        // Lavender
  qualifier: '#ADD8E6'       // Light Blue (same as adverb)
};


// --- MENU CREATION ---

/**
 * Creates the main "Multitool" menu when the document is opened.
 */
function onOpen() {
  const ui = DocumentApp.getUi();
  ui.createMenu('Multitool')
    .addItem('Word Count Tool', 'showWordCountSidebar') // From WordCount.gs
    .addSeparator()
    .addItem('Hemingway Tool', 'showHemingwaySidebar') // From Hemingway.gs
    .addSeparator()
    .addItem("Word Highlighter Tool", 'showHighlighterSidebar') // From Highlighter.gs
    .addSeparator()
    .addSubMenu(ui.createMenu('Writing Tracker Tool')
      .addItem('Log Writing Session', 'showTrackerSidebar') // From Tracker.gs
      .addItem('Configure Tracker...', 'showTrackerConfigPrompt')) // From Tracker.gs
    .addToUi();
}
