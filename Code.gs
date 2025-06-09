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
    bradly: 1, bristly: 1, bubbly: 1, bully: 1, burly: 1, butterfly: 1,
    carly: 1, charly: 1, chilly: 1, comely: 1, completely: 1, comply: 1,
    consequently: 1, costly: 1, courtly: 1, crinkly: 1, crumbly: 1,
    cuddly: 1, curly: 1, currently: 1, daily: 1, dastardly: 1, deadly: 1,
    deathly: 1, definitely: 1, dilly: 1, disorderly: 1, doily: 1,
    dolly: 1, dragonfly: 1, early: 1, elderly: 1, elly: 1, emily: 1,
    especially: 1, exactly: 1, exclusively: 1, family: 1, finally: 1,
    firefly: 1, folly: 1, friendly: 1, frilly: 1, gadfly: 1, gangly: 1,
    generally: 1, ghastly: 1, giggly: 1, globally: 1, goodly: 1,
    gravelly: 1, grisly: 1, gully: 1, haily: 1, hally: 1, harly: 1,
    hardly: 1, heavenly: 1, hillbilly: 1, hilly: 1, holly: 1, holy: 1,
    homely: 1, homily: 1, horsefly: 1, hourly: 1, immediately: 1,
    instinctively: 1, imply: 1, italy: 1, jelly: 1, jiggly: 1, jilly: 1,
    jolly: 1, july: 1, karly: 1, kelly: 1, kindly: 1, lately: 1,
    likely: 1, lilly: 1, lily: 1, lively: 1, lolly: 1, lonely: 1,
    lovely: 1, lowly: 1, luckily: 1, mealy: 1, measly: 1, melancholy: 1,
    mentally: 1, molly: 1, monopoly: 1, monthly: 1, multiply: 1,
    nightly: 1, oily: 1, only: 1, orderly: 1, panoply: 1, particularly: 1,
    partly: 1, paully: 1, pearly: 1, pebbly: 1, polly: 1, potbelly: 1,
    presumably: 1, previously: 1, pualy: 1, quarterly: 1, rally: 1,
    rarely: 1, recently: 1, rely: 1, reply: 1, reportedly: 1, roughly: 1,
    sally: 1, scaly: 1, shapely: 1, shelly: 1, shirly: 1, shortly: 1,
    sickly: 1, silly: 1, sly: 1, smelly: 1, sparkly: 1, spindly: 1,
    spritely: 1, squiggly: 1, stately: 1, steely: 1, supply: 1,
    surly: 1, tally: 1, timely: 1, trolly: 1, ugly: 1, underbelly: 1,
    unfortunately: 1, unholy: 1, unlikely: 1, usually: 1, waverly: 1,
    weekly: 1, wholly: 1, willy: 1, wily: 1, wobbly: 1, wooly: 1,
    worldly: 1, wrinkly: 1, yearly: 1
  };
}

// List of complex words and their simpler alternatives (though we only use the keys for detection here)
function getComplexWords() {
  return {
    "a number of": ["many", "some"], abundance: ["enough", "plenty"],
    "accede to": ["allow", "agree to"], accelerate: ["speed up"],
    accentuate: ["stress"], accompany: ["go with", "with"],
    accomplish: ["do"], accorded: ["given"], accrue: ["add", "gain"],
    acquiesce: ["agree"], acquire: ["get"], additional: ["more", "extra"],
    "adjacent to": ["next to"], adjustment: ["change"],
    admissible: ["allowed", "accepted"], advantageous: ["helpful"],
    "adversely impact": ["hurt"], advise: ["tell"],
    aforementioned: ["remove"], aggregate: ["total", "add"],
    aircraft: ["plane"], "all of": ["all"], alleviate: ["ease", "reduce"],
    allocate: ["divide"], "along the lines of": ["like", "as in"],
    "already existing": ["existing"], alternatively: ["or"],
    ameliorate: ["improve", "help"], anticipate: ["expect"],
    apparent: ["clear", "plain"], appreciable: ["many"],
    "as a means of": ["to"], "as of yet": ["yet"], "as to": ["on", "about"],
    "as yet": ["yet"], ascertain: ["find out", "learn"],
    assistance: ["help"], "at this time": ["now"], attain: ["meet"],
    "attributable to": ["because"], authorize: ["allow", "let"],
    "because of the fact that": ["because"], belated: ["late"],
    "benefit from": ["enjoy"], bestow: ["give", "award"],
    "by virtue of": ["by", "under"], cease: ["stop"],
    "close proximity": ["near"], commence: ["begin or start"],
    "comply with": ["follow"], concerning: ["about", "on"],
    consequently: ["so"], consolidate: ["join", "merge"],
    constitutes: ["is", "forms", "makes up"], demonstrate: ["prove", "show"],
    depart: ["leave", "go"], designate: ["choose", "name"],
    discontinue: ["drop", "stop"], "due to the fact that": ["because", "since"],
    "each and every": ["each"], economical: ["cheap"],
    eliminate: ["cut", "drop", "end"], elucidate: ["explain"],
    employ: ["use"], endeavor: ["try"], enumerate: ["count"],
    equitable: ["fair"], equivalent: ["equal"], evaluate: ["test", "check"],
    evidenced: ["showed"], exclusively: ["only"], expedite: ["hurry"],
    expend: ["spend"], expiration: ["end"], facilitate: ["ease", "help"],
    "factual evidence": ["facts", "evidence"], feasible: ["workable"],
    finalize: ["complete", "finish"], "first and foremost": ["first"],
    "for the purpose of": ["to"], forfeit: ["lose", "give up"],
    formulate: ["plan"], "honest truth": ["truth"], however: ["but", "yet"],
    "if and when": ["if", "when"], impacted: ["affected", "harmed", "changed"],
    implement: ["install", "put in place", "tool"],
    "in a timely manner": ["on time"], "in accordance with": ["by", "under"],
    "in addition": ["also", "besides", "too"],
    "in all likelihood": ["probably"], "in an effort to": ["to"],
    "in between": ["between"], "in excess of": ["more than"],
    "in lieu of": ["instead"], "in light of the fact that": ["because"],
    "in many cases": ["often"], "in order to": ["to"],
    "in regard to": ["about", "concerning", "on"],
    "in some instances ": ["sometimes"], "in terms of": ["omit"],
    "in the near future": ["soon"], "in the process of": ["omit"],
    inception: ["start"], "incumbent upon": ["must"],
    indicate: ["say", "state", "or show"], indication: ["sign"],
    initiate: ["start"], "is applicable to": ["applies to"],
    "is authorized to": ["may"], "is responsible for": ["handles"],
    "it is essential": ["must", "need to"], literally: ["omit"],
    magnitude: ["size"], maximum: ["greatest", "largest", "most"],
    methodology: ["method"], minimize: ["cut"],
    minimum: ["least", "smallest", "small"], modify: ["change"],
    monitor: ["check", "watch", "track"], multiple: ["many"],
    necessitate: ["cause", "need"], nevertheless: ["still", "besides", "even so"],
    "not certain": ["uncertain"], "not many": ["few"], "not often": ["rarely"],
    "not unless": ["only if"], "not unlike": ["similar", "alike"],
    notwithstanding: ["in spite of", "still"],
    "null and void": ["use either null or void"], numerous: ["many"],
    objective: ["aim", "goal"], obligate: ["bind", "compel"],
    obtain: ["get"], "on the contrary": ["but", "so"],
    "on the other hand": ["omit", "but", "so"], "one particular": ["one"],
    optimum: ["best", "greatest", "most"], overall: ["omit"],
    "owing to the fact that": ["because", "since"], participate: ["take part"],
    particulars: ["details"], "pass away": ["die"],
    "pertaining to": ["about", "of", "on"],
    "point in time": ["time", "point", "moment", "now"], portion: ["part"],
    possess: ["have", "own"], preclude: ["prevent"], previously: ["before"],
    "prior to": ["before"], prioritize: ["rank", "focus on"],
    procure: ["buy", "get"], proficiency: ["skill"], "provided that": ["if"],
    purchase: ["buy", "sale"], "put simply": ["omit"],
    "readily apparent": ["clear"], "refer back": ["refer"],
    regarding: ["about", "of", "on"], relocate: ["move"],
    remainder: ["rest"], remuneration: ["payment"], require: ["must", "need"],
    requirement: ["need", "rule"], reside: ["live"], residence: ["house"],
    retain: ["keep"], satisfy: ["meet", "please"], shall: ["must", "will"],
    "should you wish": ["if you want"], "similar to": ["like"],
    solicit: ["ask for", "request"], "span across": ["span", "cross"],
    strategize: ["plan"], subsequent: ["later", "next", "after", "then"],
    substantial: ["large", "much"], "successfully complete": ["complete", "pass"],
    sufficient: ["enough"], terminate: ["end", "stop"],
    "the month of": ["omit"], therefore: ["thus", "so"],
    "this day and age": ["today"], "time period": ["time", "period"],
    "took advantage of": ["preyed on"], transmit: ["send"],
    transpire: ["happen"], "until such time as": ["until"],
    utilization: ["use"], utilize: ["use"], validate: ["confirm"],
    "various different": ["various", "different"], "whether or not": ["whether"],
    "with respect to": ["on", "about"],
    "with the exception of": ["except for"], witnessed: ["saw", "seen"]
  };
}

// List of qualifying phrases
function getQualifyingWords() {
  return {
    "i believe": 1, "i consider": 1, "i don't believe": 1,
    "i don't consider": 1, "i don't feel": 1, "i don't suggest": 1,
    "i don't think": 1, "i feel": 1, "i hope to": 1, "i might": 1,
    "i suggest": 1, "i think": 1, "i was wondering": 1, "i will try": 1,
    "i wonder": 1, "in my opinion": 1, "is kind of": 1, "is sort of": 1,
    just: 1, maybe: 1, perhaps: 1, possibly: 1, "we believe": 1,
    "we consider": 1, "we don't believe": 1, "we don't consider": 1,
    "we don't feel": 1, "we don't suggest": 1, "we don't think": 1,
    "we feel": 1, "we hope to": 1, "we might": 1, "we suggest": 1,
    "we think": 1, "we were wondering": 1, "we will try": 1, "we wonder": 1
  };
}
// Colors for highlighting
const HIGHLIGHT_COLORS = {
  hardSentence: '#FFFF00',    // Yellow
  veryHardSentence: '#FFC0CB',// Pinkish-red
  adverb: '#ADD8E6',         // Light Blue
  passive: '#90EE90',        // Light Green
  complex: '#E6E6FA',        // Lavender
  qualifier: '#ADD8E6'       // Light Blue (same as adverb)
};

// Function to run when the Google Doc is opened
function onOpen() {
  DocumentApp.getUi()
    .createMenu('Hemingway Helper')
    .addItem('Analyze Document', 'runAnalysis')
    .addToUi();
}

// Function to reset the analysis data
function resetAnalysisData() {
  analysisData = {
    paragraphs: 0,
    sentences: 0,
    words: 0,
    totalCharacters: 0,
    hardSentences: 0,
    veryHardSentences: 0,
    adverbs: 0,
    passiveVoice: 0,
    complexPhrases: 0, // Renamed from 'complex' for clarity
    qualifiers: 0
  };
}

// Main function to trigger the analysis
function runAnalysis() {
  resetAnalysisData();
  clearPreviousHighlights();

  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const paragraphs = body.getParagraphs();

  analysisData.paragraphs = paragraphs.filter(p => p.getText().trim() !== "").length;

  for (const paragraphElement of paragraphs) {
    const paragraphText = paragraphElement.getText();
    if (paragraphText.trim() !== "") {
      processParagraph(paragraphElement, paragraphText);
    }
  }
  showAnalysisResults();
}

// Function to clear previous highlights from the document
function clearPreviousHighlights() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const paragraphs = body.getParagraphs();

  for (const paragraph of paragraphs) {
    if (paragraph.getText().trim() !== "") {
      // Reset background color for the entire paragraph
      // Note: A more sophisticated approach might involve tracking only specific highlights
      // but for simplicity, we clear all background colors.
      paragraph.editAsText().setBackgroundColor(0, paragraph.getText().length - 1, null);
    }
  }
}

// Function to process each paragraph
function processParagraph(paragraphElement, paragraphText) {
  const sentencesInfo = getSentencesFromParagraphText(paragraphText);
  analysisData.sentences += sentencesInfo.length;

  sentencesInfo.forEach(sentenceInfo => {
    // sentenceInfo = { text: "Actual sentence.", startIndexInPara: 0 }
    let originalSentenceText = sentenceInfo.text;
    analysisData.totalCharacters += originalSentenceText.length;
    let cleanSentenceForStats = originalSentenceText.replace(/[^a-zA-Z0-9. ]/gi, "") + "."; // Ensure ends with period for splitting
    let wordsInSentence = cleanSentenceForStats.split(/\s+/).filter(w => w.length > 0);
    let wordCount = wordsInSentence.length;
    let letterCount = wordsInSentence.join("").length;

    analysisData.words += wordCount;

    // Sentence-level difficulty highlighting
    if (wordCount > 0) { // Avoid division by zero for calculateLevel
        let level = calculateLevel(letterCount, wordCount, 1); // 1 for one sentence
        let sentenceEndIndexInPara = sentenceInfo.startIndexInPara + originalSentenceText.length - 1;

        if (sentenceEndIndexInPara >= sentenceInfo.startIndexInPara) { // Ensure valid range
            if (wordCount >= 14) { // Only highlight if sentence is reasonably long
                if (level >= 10 && level < 14) {
                    analysisData.hardSentences++;
                    paragraphElement.asText().setBackgroundColor(sentenceInfo.startIndexInPara, sentenceEndIndexInPara, HIGHLIGHT_COLORS.hardSentence);
                } else if (level >= 14) {
                    analysisData.veryHardSentences++;
                    paragraphElement.asText().setBackgroundColor(sentenceInfo.startIndexInPara, sentenceEndIndexInPara, HIGHLIGHT_COLORS.veryHardSentence);
                }
            }
        }
    }

    // Find and highlight specific issues within the sentence
    let adverbIssues = findAdverbsInSentence(originalSentenceText);
    applyIssueFormatting(paragraphElement, sentenceInfo.startIndexInPara, adverbIssues, 'adverb');

    let passiveIssues = findPassiveInSentence(originalSentenceText);
    applyIssueFormatting(paragraphElement, sentenceInfo.startIndexInPara, passiveIssues, 'passive');

    let complexIssues = findComplexInSentence(originalSentenceText);
    applyIssueFormatting(paragraphElement, sentenceInfo.startIndexInPara, complexIssues, 'complex');

    let qualifierIssues = findQualifiersInSentence(originalSentenceText);
    applyIssueFormatting(paragraphElement, sentenceInfo.startIndexInPara, qualifierIssues, 'qualifier');
  });
}

// Function to split paragraph text into sentences and get their start indices
function getSentencesFromParagraphText(paragraphText) {
  const sentences = [];
  // Regex to split by common sentence terminators, keeping the terminator.
  // It looks for '.', '!', '?' followed by a space or end of string.
  const sentenceEndRegex = /([.!?])(\s+|$)/g;
  let lastIndex = 0;
  let match;

  while ((match = sentenceEndRegex.exec(paragraphText)) !== null) {
    let sentenceText = paragraphText.substring(lastIndex, match.index + match[1].length).trim();
    if (sentenceText) {
      sentences.push({ text: sentenceText, startIndexInPara: lastIndex });
    }
    lastIndex = match.index + match[0].length; // Move past the terminator and following space
     while(lastIndex < paragraphText.length && /\s/.test(paragraphText[lastIndex])) {
        lastIndex++; // Skip any additional leading spaces for the next sentence
    }
  }
  // Add any remaining text as the last sentence if it's not empty
  if (lastIndex < paragraphText.length) {
    let remainingText = paragraphText.substring(lastIndex).trim();
    if (remainingText) {
      sentences.push({ text: remainingText, startIndexInPara: lastIndex });
    }
  }
  return sentences.filter(s => s.text.length > 0);
}


// Function to find adverbs in a sentence
function findAdverbsInSentence(sentenceText) {
  const issues = [];
  const lyWordsToExclude = getLyWords();
  const words = sentenceText.split(/(\s+)/); // Split by space, keeping spaces to maintain indices

  let currentIndex = 0;
  words.forEach(wordOrSpace => {
    if (wordOrSpace.trim().length > 0) { // It's a word
      let cleanWord = wordOrSpace.replace(/[^a-zA-Z]/g, ""); // Clean punctuation for matching
      if (cleanWord.toLowerCase().endsWith("ly") && !lyWordsToExclude[cleanWord.toLowerCase()]) {
        issues.push({
          startIndexInSentence: currentIndex,
          endIndexInSentence: currentIndex + wordOrSpace.length - 1,
          type: 'adverb'
        });
        analysisData.adverbs++;
      }
    }
    currentIndex += wordOrSpace.length;
  });
  return issues;
}

// Function to find passive voice in a sentence
function findPassiveInSentence(sentenceText) {
  const issues = [];
  const preWords = ["is", "are", "was", "were", "be", "been", "being"];
  // A simple regex for words ending in 'ed' often found in passive voice.
  // This is a simplification and might have false positives/negatives.
  const edWordRegex = /\b(\w+ed)\b/gi;
  let match;

  // Split sentence into words and their original indices to handle punctuation
  const wordsWithIndices = [];
  let currentWord = "";
  let currentWordStartIndex = -1;
  for(let i = 0; i < sentenceText.length; i++) {
      const char = sentenceText[i];
      if (/\w/.test(char)) { // If it's a word character
          if(currentWordStartIndex === -1) currentWordStartIndex = i;
          currentWord += char;
      } else { // If it's a delimiter or end of word
          if (currentWord) {
              wordsWithIndices.push({text: currentWord, start: currentWordStartIndex, end: i - 1});
          }
          currentWord = "";
          currentWordStartIndex = -1;
      }
  }
  if (currentWord) { // Add last word if any
      wordsWithIndices.push({text: currentWord, start: currentWordStartIndex, end: sentenceText.length - 1});
  }


  for (let i = 0; i < wordsWithIndices.length; i++) {
    const currentWordInfo = wordsWithIndices[i];
    if (currentWordInfo.text.toLowerCase().endsWith("ed") && i > 0) {
      const prevWordInfo = wordsWithIndices[i-1];
      if (prevWordInfo && preWords.includes(prevWordInfo.text.toLowerCase())) {
         // Check if the 'ed' word is likely a verb (not an adjective like "talented person")
         // This is tricky. For simplicity, we'll assume it's passive here.
         // More advanced NLP would be needed for higher accuracy.
        issues.push({
          startIndexInSentence: prevWordInfo.start,
          endIndexInSentence: currentWordInfo.end,
          type: 'passive'
        });
        analysisData.passiveVoice++;
      }
    }
  }
  return issues;
}


// Function to find complex phrases in a sentence
function findComplexInSentence(sentenceText) {
  const issues = [];
  const complexPhraseMap = getComplexWords(); // We only need the keys
  const complexPhrases = Object.keys(complexPhraseMap);

  complexPhrases.forEach(phrase => {
    let searchStartIndex = 0;
    let foundIndex;
    while ((foundIndex = sentenceText.toLowerCase().indexOf(phrase.toLowerCase(), searchStartIndex)) !== -1) {
      // Basic check to avoid matching parts of words (e.g., "is" in "this")
      const before = sentenceText[foundIndex - 1];
      const after = sentenceText[foundIndex + phrase.length];
      const isWordBoundaryBefore = !before || /\W/.test(before);
      const isWordBoundaryAfter = !after || /\W/.test(after);

      if (isWordBoundaryBefore && isWordBoundaryAfter) {
        issues.push({
          startIndexInSentence: foundIndex,
          endIndexInSentence: foundIndex + phrase.length - 1,
          type: 'complex'
        });
        analysisData.complexPhrases++;
      }
      searchStartIndex = foundIndex + phrase.length; // Continue search after this match
    }
  });
  return issues;
}

// Function to find qualifying phrases in a sentence
function findQualifiersInSentence(sentenceText) {
  const issues = [];
  const qualifierMap = getQualifyingWords();
  const qualifiers = Object.keys(qualifierMap);

  qualifiers.forEach(qualifier => {
    let searchStartIndex = 0;
    let foundIndex;
    while ((foundIndex = sentenceText.toLowerCase().indexOf(qualifier.toLowerCase(), searchStartIndex)) !== -1) {
       const before = sentenceText[foundIndex - 1];
       const after = sentenceText[foundIndex + qualifier.length];
       const isWordBoundaryBefore = !before || /\W/.test(before);
       const isWordBoundaryAfter = !after || /\W/.test(after);

      if (isWordBoundaryBefore && isWordBoundaryAfter) {
        issues.push({
          startIndexInSentence: foundIndex,
          endIndexInSentence: foundIndex + qualifier.length - 1,
          type: 'qualifier'
        });
        analysisData.qualifiers++;
      }
      searchStartIndex = foundIndex + qualifier.length;
    }
  });
  return issues;
}

// Function to apply formatting for identified issues
function applyIssueFormatting(paragraphElement, sentenceOffsetInPara, issues, issueType) {
  const color = HIGHLIGHT_COLORS[issueType];
  if (!color) return;

  issues.forEach(issue => {
    let start = sentenceOffsetInPara + issue.startIndexInSentence;
    let end = sentenceOffsetInPara + issue.endIndexInSentence;

    // Ensure the highlight doesn't spill over sentence boundaries or paragraph boundaries
    // and that the range is valid.
    const paraTextLength = paragraphElement.getText().length;
    start = Math.max(0, start); // Cannot be less than 0
    end = Math.min(paraTextLength - 1, end); // Cannot exceed paragraph length

    if (end >= start) {
      try {
        // Check if there's already a sentence-level highlight (hard/very hard)
        // If so, we might choose not to overwrite it, or apply a combined style.
        // For simplicity, we'll overwrite. More complex logic could layer styles.
        // Also, ensure we don't re-highlight something already highlighted for the same reason by a broader category.
        paragraphElement.asText().setBackgroundColor(start, end, color);
      } catch (e) {
        console.error("Error applying formatting for " + issueType + ": " + e.toString() + " on text range " + start + "-" + end);
      }
    }
  });
}

// Function to show analysis results in a sidebar
function showAnalysisResults() {
  const ui = DocumentApp.getUi();
  // Pass a copy of analysisData to the HTML template
  const templateData = JSON.parse(JSON.stringify(analysisData));

  // Add suggestions to the templateData
  templateData.adverbSuggestion = `Try to use ${Math.round(templateData.paragraphs / 3) || 0} or fewer.`;
  templateData.passiveSuggestion = `Aim for ${Math.round(templateData.sentences / 5) || 0} or fewer.`;
  templateData.readabilityLevel = calculateOverallReadability(); // Calculate and add readability

  const htmlTemplate = HtmlService.createTemplateFromFile('SidebarHtml');
  htmlTemplate.data = templateData; // Pass data to the template
  const htmlOutput = htmlTemplate.evaluate().setTitle('Hemingway Analysis');
  ui.showSidebar(htmlOutput);
}

function calculateOverallReadability() {
  // Ensure we have enough data to calculate
  if (analysisData.words === 0 || analysisData.sentences === 0 || analysisData.totalCharacters === 0) {
    return "N/A - Insufficient Text";
  }

  // Automated Readability Index (ARI) Formula:
  // ARI = 4.71 * (characters/words) + 0.5 * (words/sentences) - 21.43
  // We use analysisData.totalCharacters, analysisData.words, and analysisData.sentences

  let score = (4.71 * (analysisData.totalCharacters / analysisData.words)) +
              (0.5 * (analysisData.words / analysisData.sentences)) - 21.43;

  // ARI scores are typically rounded to the nearest whole number to represent a grade level.
  // Clamp score to a minimum of 0 if it goes negative for very simple texts.
  let roundedScore = Math.max(0, Math.round(score));

  return `Readability Grade Level (ARI): ${roundedScore}`;
}

