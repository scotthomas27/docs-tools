// --- HEMINGWAY ANALYZER FUNCTIONS ---

/**
 * Controller: Runs the Hemingway analysis and displays it in its own sidebar.
 */
function showHemingwaySidebar() {
  // 1. Reset data and clear old highlights
  resetAnalysisData();
  clearPreviousHighlights();

  // 2. Perform Hemingway-style analysis
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

  // 3. Add suggestions to the data object
  analysisData.adverbSuggestion = `Aim for ${Math.round(analysisData.paragraphs / 3) || 0} or fewer.`;
  analysisData.passiveSuggestion = `Aim for ${Math.round(analysisData.sentences / 5) || 0} or fewer.`;
  analysisData.readabilityLevel = calculateOverallReadability();

  // 4. Show results in the Hemingway sidebar
  const htmlTemplate = HtmlService.createTemplateFromFile('HemingwaySidebar.html');
  htmlTemplate.data = analysisData;
  const htmlOutput = htmlTemplate.evaluate().setTitle('Hemingway Analysis');
  DocumentApp.getUi().showSidebar(htmlOutput);
}
/**
 * A helper function to clear highlights before a new analysis is run.
 * This is now internal to the Hemingway tool.
 */
function clearPreviousHighlights() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const paragraphs = body.getParagraphs();
  for (const paragraph of paragraphs) {
    if (paragraph.getText().trim() !== "") {
      // This is a "brute-force" clear that removes ALL background color
      // from the main body of the document.
      paragraph.editAsText().setBackgroundColor(0, paragraph.getText().length - 1, null);
    }
  }
}

/**
 * Resets the global analysis data object for a fresh run.
 */
function resetAnalysisData() {
  analysisData = {paragraphs:0,sentences:0,words:0,totalCharacters:0,hardSentences:0,veryHardSentences:0,adverbs:0,passiveVoice:0,complexPhrases:0,qualifiers:0};
}

/**
 * Processes a single paragraph for all Hemingway issues.
 */
function processParagraph(paragraphElement, paragraphText) {
  const sentencesInfo = getSentencesFromParagraphText(paragraphText);
  analysisData.sentences += sentencesInfo.length;
  sentencesInfo.forEach(sentenceInfo => {
    let originalSentenceText = sentenceInfo.text;
    analysisData.totalCharacters += originalSentenceText.length;
    let cleanSentenceForStats = originalSentenceText.replace(/[^a-zA-Z0-9. ]/gi, "") + ".";
    let wordsInSentence = cleanSentenceForStats.split(/\s+/).filter(w => w.length > 0);
    let wordCount = wordsInSentence.length;
    let letterCount = wordsInSentence.join("").length;
    analysisData.words += wordCount;
    if (wordCount > 0) {
      let level = calculateLevel(letterCount, wordCount, 1);
      let sentenceEndIndexInPara = sentenceInfo.startIndexInPara + originalSentenceText.length - 1;
      if (sentenceEndIndexInPara >= sentenceInfo.startIndexInPara) {
        if (wordCount >= 14) {
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

/**
 * Applies background color for a given issue type.
 */
function applyIssueFormatting(paragraphElement, sentenceOffsetInPara, issues, issueType) {
  const color = HIGHLIGHT_COLORS[issueType];
  if (!color) return;
  issues.forEach(issue => {
    let start = sentenceOffsetInPara + issue.startIndexInSentence;
    let end = sentenceOffsetInPara + issue.endIndexInSentence;
    const paraTextLength = paragraphElement.getText().length;
    start = Math.max(0, start);
    end = Math.min(paraTextLength - 1, end);
    if (end >= start) {
      try {
        paragraphElement.asText().setBackgroundColor(start, end, color);
      } catch (e) {
        console.error("Error applying formatting for " + issueType + ": " + e.toString() + " on text range " + start + "-" + end);
      }
    }
  });
}

/**
 * Utility to split a paragraph into sentences.
 */
function getSentencesFromParagraphText(paragraphText) {
  const sentences = [];
  const sentenceEndRegex = /([.!?])(\s+|$)/g;
  let lastIndex = 0;
  let match;
  while ((match = sentenceEndRegex.exec(paragraphText)) !== null) {
    let sentenceText = paragraphText.substring(lastIndex, match.index + match[1].length).trim();
    if (sentenceText) {
      sentences.push({ text: sentenceText, startIndexInPara: lastIndex });
    }
    lastIndex = match.index + match[0].length;
    while (lastIndex < paragraphText.length && /\s/.test(paragraphText[lastIndex])) {
      lastIndex++;
    }
  }
  if (lastIndex < paragraphText.length) {
    let remainingText = paragraphText.substring(lastIndex).trim();
    if (remainingText) {
      sentences.push({ text: remainingText, startIndexInPara: lastIndex });
    }
  }
  return sentences.filter(s => s.text.length > 0);
}

/**
 * Finds adverbs in a sentence.
 */
function findAdverbsInSentence(sentenceText) {
  const issues = [];
  const lyWordsToExclude = getLyWords();
  const words = sentenceText.split(/(\s+)/);
  let currentIndex = 0;
  words.forEach(wordOrSpace => {
    if (wordOrSpace.trim().length > 0) {
      let cleanWord = wordOrSpace.replace(/[^a-zA-Z]/g, "");
      if (cleanWord.toLowerCase().endsWith("ly") && !lyWordsToExclude[cleanWord.toLowerCase()]) {
        issues.push({ startIndexInSentence: currentIndex, endIndexInSentence: currentIndex + wordOrSpace.length - 1, type: 'adverb' });
        analysisData.adverbs++;
      }
    }
    currentIndex += wordOrSpace.length;
  });
  return issues;
}

/**
 * Finds passive voice in a sentence.
 */
function findPassiveInSentence(sentenceText) {
    const issues = [];
    const preWords = ["is", "are", "was", "were", "be", "been", "being"];
    const wordsWithIndices = [];
    let currentWord = "";
    let currentWordStartIndex = -1;
    for(let i = 0; i < sentenceText.length; i++) {
        const char = sentenceText[i];
        if (/\w/.test(char)) {
            if(currentWordStartIndex === -1) currentWordStartIndex = i;
            currentWord += char;
        } else {
            if (currentWord) {
                wordsWithIndices.push({text: currentWord, start: currentWordStartIndex, end: i - 1});
            }
            currentWord = "";
            currentWordStartIndex = -1;
        }
    }
    if (currentWord) {
        wordsWithIndices.push({text: currentWord, start: currentWordStartIndex, end: sentenceText.length - 1});
    }
    for (let i = 0; i < wordsWithIndices.length; i++) {
        const currentWordInfo = wordsWithIndices[i];
        if (currentWordInfo.text.toLowerCase().endsWith("ed") && i > 0) {
            const prevWordInfo = wordsWithIndices[i-1];
            if (prevWordInfo && preWords.includes(prevWordInfo.text.toLowerCase())) {
                issues.push({ startIndexInSentence: prevWordInfo.start, endIndexInSentence: currentWordInfo.end, type: 'passive' });
                analysisData.passiveVoice++;
            }
        }
    }
    return issues;
}

/**
 * Finds complex phrases in a sentence.
 */
function findComplexInSentence(sentenceText) {
  const issues = [];
  const complexPhrases = Object.keys(getComplexWords());
  complexPhrases.forEach(phrase => {
    let searchStartIndex = 0;
    let foundIndex;
    while ((foundIndex = sentenceText.toLowerCase().indexOf(phrase.toLowerCase(), searchStartIndex)) !== -1) {
      const before = sentenceText[foundIndex - 1];
      const after = sentenceText[foundIndex + phrase.length];
      const isWordBoundaryBefore = !before || /\W/.test(before);
      const isWordBoundaryAfter = !after || /\W/.test(after);
      if (isWordBoundaryBefore && isWordBoundaryAfter) {
        issues.push({ startIndexInSentence: foundIndex, endIndexInSentence: foundIndex + phrase.length - 1, type: 'complex' });
        analysisData.complexPhrases++;
      }
      searchStartIndex = foundIndex + phrase.length;
    }
  });
  return issues;
}

/**
 * Finds qualifying phrases in a sentence.
 */
function findQualifiersInSentence(sentenceText) {
  const issues = [];
  const qualifiers = Object.keys(getQualifyingWords());
  qualifiers.forEach(qualifier => {
    let searchStartIndex = 0;
    let foundIndex;
    while ((foundIndex = sentenceText.toLowerCase().indexOf(qualifier.toLowerCase(), searchStartIndex)) !== -1) {
      const before = sentenceText[foundIndex - 1];
      const after = sentenceText[foundIndex + qualifier.length];
      const isWordBoundaryBefore = !before || /\W/.test(before);
      const isWordBoundaryAfter = !after || /\W/.test(after);
      if (isWordBoundaryBefore && isWordBoundaryAfter) {
        issues.push({ startIndexInSentence: foundIndex, endIndexInSentence: foundIndex + qualifier.length - 1, type: 'qualifier' });
        analysisData.qualifiers++;
      }
      searchStartIndex = foundIndex + qualifier.length;
    }
  });
  return issues;
}


// --- DATA & READABILITY HELPERS ---

function calculateLevel(letters, words, sentences) {
  if (words === 0 || sentences === 0) return 0;
  let level = Math.round(4.71 * (letters / words) + 0.5 * (words / sentences) - 21.43);
  return level <= 0 ? 0 : level;
}

function calculateOverallReadability() {
  if (analysisData.words === 0 || analysisData.sentences === 0 || analysisData.totalCharacters === 0) {
    return "N/A - Insufficient Text";
  }
  let score = (4.71 * (analysisData.totalCharacters / analysisData.words)) +
              (0.5 * (analysisData.words / analysisData.sentences)) - 21.43;
  let roundedScore = Math.max(0, Math.round(score));
  return `Grade ${roundedScore} (ARI)`;
}

function getLyWords() {
  return {"actually":1,"additionally":1,"allegedly":1,"ally":1,"alternatively":1,"anomaly":1,"apply":1,"approximately":1,"ashely":1,"ashly":1,"assembly":1,"awfully":1,"baily":1,"belly":1,"bely":1,"billy":1,"bradly":1,"bristly":1,"bubbly":1,"bully":1,"burly":1,"butterfly":1,"carly":1,"charly":1,"chilly":1,"comely":1,"completely":1,"comply":1,"consequently":1,"costly":1,"courtly":1,"crinkly":1,"crumbly":1,"cuddly":1,"curly":1,"currently":1,"daily":1,"dastardly":1,"deadly":1,"deathly":1,"definitely":1,"dilly":1,"disorderly":1,"doily":1,"dolly":1,"dragonfly":1,"early":1,"elderly":1,"elly":1,"emily":1,"especially":1,"exactly":1,"exclusively":1,"family":1,"finally":1,"firefly":1,"folly":1,"friendly":1,"frilly":1,"gadfly":1,"gangly":1,"generally":1,"ghastly":1,"giggly":1,"globally":1,"goodly":1,"gravelly":1,"grisly":1,"gully":1,"haily":1,"hally":1,"harly":1,"hardly":1,"heavenly":1,"hillbilly":1,"hilly":1,"holly":1,"holy":1,"homely":1,"homily":1,"horsefly":1,"hourly":1,"immediately":1,"instinctively":1,"imply":1,"italy":1,"jelly":1,"jiggly":1,"jilly":1,"jolly":1,"july":1,"karly":1,"kelly":1,"kindly":1,"lately":1,"likely":1,"lilly":1,"lily":1,"lively":1,"lolly":1,"lonely":1,"lovely":1,"lowly":1,"luckily":1,"mealy":1,"measly":1,"melancholy":1,"mentally":1,"molly":1,"monopoly":1,"monthly":1,"multiply":1,"nightly":1,"oily":1,"only":1,"orderly":1,"panoply":1,"particularly":1,"partly":1,"paully":1,"pearly":1,"pebbly":1,"polly":1,"potbelly":1,"presumably":1,"previously":1,"pualy":1,"quarterly":1,"rally":1,"rarely":1,"recently":1,"rely":1,"reply":1,"reportedly":1,"roughly":1,"sally":1,"scaly":1,"shapely":1,"shelly":1,"shirly":1,"shortly":1,"sickly":1,"silly":1,"sly":1,"smelly":1,"sparkly":1,"spindly":1,"spritely":1,"squiggly":1,"stately":1,"steely":1,"supply":1,"surly":1,"tally":1,"timely":1,"trolly":1,"ugly":1,"underbelly":1,"unfortunately":1,"unholy":1,"unlikely":1,"usually":1,"waverly":1,"weekly":1,"wholly":1,"willy":1,"wily":1,"wobbly":1,"wooly":1,"worldly":1,"wrinkly":1,"yearly":1};
}

function getComplexWords() {
  return {"a number of":["many","some"],"abundance":["enough","plenty"],"accede to":["allow","agree to"],"accelerate":["speed up"],"accentuate":["stress"],"accompany":["go with","with"],"accomplish":["do"],"accorded":["given"],"accrue":["add","gain"],"acquiesce":["agree"],"acquire":["get"],"additional":["more","extra"],"adjacent to":["next to"],"adjustment":["change"],"admissible":["allowed","accepted"],"advantageous":["helpful"],"adversely impact":["hurt"],"advise":["tell"],"aforementioned":["remove"],"aggregate":["total","add"],"aircraft":["plane"],"all of":["all"],"alleviate":["ease","reduce"],"allocate":["divide"],"along the lines of":["like","as in"],"already existing":["existing"],"alternatively":["or"],"ameliorate":["improve","help"],"anticipate":["expect"],"apparent":["clear","plain"],"appreciable":["many"],"as a means of":["to"],"as of yet":["yet"],"as to":["on","about"],"as yet":["yet"],"ascertain":["find out","learn"],"assistance":["help"],"at this time":["now"],"attain":["meet"],"attributable to":["because"],"authorize":["allow","let"],"because of the fact that":["because"],"belated":["late"],"benefit from":["enjoy"],"bestow":["give","award"],"by virtue of":["by","under"],"cease":["stop"],"close proximity":["near"],"commence":["begin or start"],"comply with":["follow"],"concerning":["about","on"],"consequently":["so"],"consolidate":["join","merge"],"constitutes":["is","forms","makes up"],"demonstrate":["prove","show"],"depart":["leave","go"],"designate":["choose","name"],"discontinue":["drop","stop"],"due to the fact that":["because","since"],"each and every":["each"],"economical":["cheap"],"eliminate":["cut","drop","end"],"elucidate":["explain"],"employ":["use"],"endeavor":["try"],"enumerate":["count"],"equitable":["fair"],"equivalent":["equal"],"evaluate":["test","check"],"evidenced":["showed"],"exclusively":["only"],"expedite":["hurry"],"expend":["spend"],"expiration":["end"],"facilitate":["ease","help"],"factual evidence":["facts","evidence"],"feasible":["workable"],"finalize":["complete","finish"],"first and foremost":["first"],"for the purpose of":["to"],"forfeit":["lose","give up"],"formulate":["plan"],"honest truth":["truth"],"however":["but","yet"],"if and when":["if","when"],"impacted":["affected","harmed","changed"],"implement":["install","put in place","tool"],"in a timely manner":["on time"],"in accordance with":["by","under"],"in addition":["also","besides","too"],"in all likelihood":["probably"],"in an effort to":["to"],"in between":["between"],"in excess of":["more than"],"in lieu of":["instead"],"in light of the fact that":["because"],"in many cases":["often"],"in order to":["to"],"in regard to":["about","concerning","on"],"in some instances ":["sometimes"],"in terms of":["omit"],"in the near future":["soon"],"in the process of":["omit"],"inception":["start"],"incumbent upon":["must"],"indicate":["say","state","or show"],"indication":["sign"],"initiate":["start"],"is applicable to":["applies to"],"is authorized to":["may"],"is responsible for":["handles"],"it is essential":["must","need to"],"literally":["omit"],"magnitude":["size"],"maximum":["greatest","largest","most"],"methodology":["method"],"minimize":["cut"],"minimum":["least","smallest","small"],"modify":["change"],"monitor":["check","watch","track"],"multiple":["many"],"necessitate":["cause","need"],"nevertheless":["still","besides","even so"],"not certain":["uncertain"],"not many":["few"],"not often":["rarely"],"not unless":["only if"],"not unlike":["similar","alike"],"notwithstanding":["in spite of","still"],"null and void":["use either null or void"],"numerous":["many"],"objective":["aim","goal"],"obligate":["bind","compel"],"obtain":["get"],"on the contrary":["but","so"],"on the other hand":["omit","but","so"],"one particular":["one"],"optimum":["best","greatest","most"],"overall":["omit"],"owing to the fact that":["because","since"],"participate":["take part"],"particulars":["details"],"pass away":["die"],"pertaining to":["about","of","on"],"point in time":["time","point","moment","now"],"portion":["part"],"possess":["have","own"],"preclude":["prevent"],"previously":["before"],"prior to":["before"],"prioritize":["rank","focus on"],"procure":["buy","get"],"proficiency":["skill"],"provided that":["if"],"purchase":["buy","sale"],"put simply":["omit"],"readily apparent":["clear"],"refer back":["refer"],"regarding":["about","of","on"],"relocate":["move"],"remainder":["rest"],"remuneration":["payment"],"require":["must","need"],"requirement":["need","rule"],"reside":["live"],"residence":["house"],"retain":["keep"],"satisfy":["meet","please"],"shall":["must","will"],"should you wish":["if you want"],"similar to":["like"],"solicit":["ask for","request"],"span across":["span","cross"],"strategize":["plan"],"subsequent":["later","next","after","then"],"substantial":["large","much"],"successfully complete":["complete","pass"],"sufficient":["enough"],"terminate":["end","stop"],"the month of":["omit"],"therefore":["thus","so"],"this day and age":["today"],"time period":["time","period"],"took advantage of":["preyed on"],"transmit":["send"],"transpire":["happen"],"until such time as":["until"],"utilization":["use"],"utilize":["use"],"validate":["confirm"],"various different":["various","different"],"whether or not":["whether"],"with respect to":["on","about"],"with the exception of":["except for"],"witnessed":["saw","seen"]};
}

function getQualifyingWords() {
  return {"i believe":1,"i consider":1,"i don't believe":1,"i don't consider":1,"i don't feel":1,"i don't suggest":1,"i don't think":1,"i feel":1,"i hope to":1,"i might":1,"i suggest":1,"i think":1,"i was wondering":1,"i will try":1,"i wonder":1,"in my opinion":1,"is kind of":1,"is sort of":1,"just":1,"maybe":1,"perhaps":1,"possibly":1,"we believe":1,"we consider":1,"we don't believe":1,"we don't consider":1,"we don't feel":1,"we don't suggest":1,"we don't think":1,"we feel":1,"we hope to":1,"we might":1,"we suggest":1,"we think":1,"we were wondering":1,"we will try":1,"we wonder":1};
}

// --- NEW: INTERACTIVE SUGGESTION FUNCTIONS ---

/**
 * Gets the text of the word where the user's cursor is currently located.
 * @returns {string|null} The word at the cursor, or null if no text is selected.
 */
function getWordAtCursor() {
  const cursor = DocumentApp.getActiveDocument().getCursor();
  if (!cursor) {
    return null; // No cursor or selection.
  }

  // Find the word boundaries around the cursor position
  const surroundingText = cursor.getSurroundingText().getText();
  const cursorOffset = cursor.getSurroundingTextOffset();

  // Go backwards from cursor to find the start of the word
  let start = cursorOffset;
  while (start > 0 && /\w/.test(surroundingText.charAt(start - 1))) {
    start--;
  }

  // Go forwards from cursor to find the end of the word
  let end = cursorOffset;
  while (end < surroundingText.length && /\w/.test(surroundingText.charAt(end))) {
    end++;
  }
  
  const word = surroundingText.substring(start, end).trim();
  return word || null;
}

/**
 * Looks up a word in the complex words dictionary and returns suggestions.
 * @param {string} word The word to look up.
 * @returns {Array|null} An array of suggestion strings, or null if none are found.
 */
function getSuggestionForWord(word) {
  if (!word) return null;
  const complexWords = getComplexWords();
  // Perform a case-insensitive lookup
  const lowerCaseWord = word.toLowerCase();

  for (const key in complexWords) {
    if (key.toLowerCase() === lowerCaseWord) {
      return complexWords[key];
    }
  }
  return null; // No match found
}
