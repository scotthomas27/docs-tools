// --- CONTROLLERS: STATEFUL & CONFIGURABLE WRITING TRACKER ---

/**
 * Kicks off the advanced configuration process by showing the mapping sidebar.
 */
function showTrackerConfigPrompt() {
  const ui = DocumentApp.getUi();
  const properties = PropertiesService.getDocumentProperties();
  
  let sheetId = properties.getProperty('trackerSheetId');
  let sheetName = properties.getProperty('trackerSheetName');

  if (!sheetId) {
    const response = ui.prompt('Configure Tracker (Step 1 of 2)', 'Please paste the full URL of your Google Sheet tracker:', ui.ButtonSet.OK_CANCEL);
    if (response.getSelectedButton() !== ui.Button.OK) return;
    const url = response.getResponseText();
    try {
      sheetId = SpreadsheetApp.openByUrl(url).getId();
      properties.setProperty('trackerSheetId', sheetId);
    } catch (e) {
      ui.alert('Error', 'Could not open the spreadsheet from that URL. Please check the link and permissions.', ui.ButtonSet.OK);
      return;
    }
  }

  if (!sheetName) {
     const nameResponse = ui.prompt('Configure Tracker (Step 2 of 2)', 'Now, please enter the EXACT name of the sheet (the tab at the bottom) where your data is stored:', ui.ButtonSet.OK_CANCEL);
     if (nameResponse.getSelectedButton() !== ui.Button.OK || !nameResponse.getResponseText()) return;
     sheetName = nameResponse.getResponseText();
     properties.setProperty('trackerSheetName', sheetName);
  }
  
  const headers = getSheetHeaders(sheetId, sheetName);
  if (headers) {
    const htmlTemplate = HtmlService.createTemplateFromFile('ConfigSidebar.html');
    htmlTemplate.headers = headers;
    const htmlOutput = htmlTemplate.evaluate().setWidth(500).setHeight(450);
    ui.showModalDialog(htmlOutput, 'Configure Tracker Columns');
  }
}

/**
 * DYNAMIC V2.4 (Persistent Fields): Reads the config map, fetches the last value
 * for any "persistent" fields, and builds the custom sidebar.
 */
function showTrackerSidebar() {
  const properties = PropertiesService.getDocumentProperties();
  const configString = properties.getProperty('trackerConfigMap');
  const savedDataString = properties.getProperty('trackerFormData');
  const savedData = savedDataString ? JSON.parse(savedDataString) : {};

  if (!configString) {
    const html = HtmlService.createHtmlOutput(
      '<div style="padding: 12px; font-family: Arial, sans-serif;">' +
      '<p><b>Tracker is not configured.</b></p>' +
      '<p>Please run "Configure Tracker..." from the Writing Tracker menu to connect your Google Sheet.</p>' +
      '</div>'
    ).setTitle('Log Session');
    DocumentApp.getUi().showSidebar(html);
    return;
  }

  const configMap = JSON.parse(configString);
  let formHtml = '';
  const sheetId = properties.getProperty('trackerSheetId');
  const sheetName = properties.getProperty('trackerSheetName');

  // Fetch persistent values BEFORE building the form
  const persistentValues = {};
  configMap.forEach(item => {
    if (item.fieldType === 'persistent') {
      const lastValue = getLastValueFromColumn(sheetId, sheetName, item.columnName);
      if (lastValue !== null) {
        persistentValues[item.columnName] = lastValue;
      }
    }
  });

  const rawTemplate = HtmlService.createTemplateFromFile('TrackerTemplate').getRawContent();
  const templateMap = {
    'text':       rawTemplate.split('<!-- Text Input Snippet -->')[1].split('<!--')[0],
    'persistent': rawTemplate.split('<!-- Persistent Text Input Snippet -->')[1].split('<!--')[0],
    'date':       rawTemplate.split('<!-- Date Input Snippet -->')[1].split('<!--')[0],
    'time':       rawTemplate.split('<!-- Time Input Snippet -->')[1].split('<!--')[0],
    'numeric':    rawTemplate.split('<!-- Numeric Input Snippet -->')[1].split('<!--')[0],
    'checkbox':   rawTemplate.split('<!-- Checkbox Snippet -->')[1].split('<!--')[0],
    'title':      rawTemplate.split('<!-- Document Title Snippet -->')[1].split('<!--')[0],
    'dropdown':   rawTemplate.split('<!-- Dropdown Snippet -->')[1].split('<!--')[0]
  };

  configMap.forEach(item => {
    const id = item.columnName.replace(/[^a-zA-Z0-9]/g, '');
    let value = '';

    // Value Hierarchy:
    // 1. Prioritize data from the current session's memory.
    if (savedData[id] !== undefined) {
      value = savedData[id];
    } 
    // 2. If no session data, check for a fetched persistent value.
    else if (item.fieldType === 'persistent' && persistentValues[item.columnName] !== undefined) {
      value = persistentValues[item.columnName];
    } 
    // 3. Otherwise, use auto-fills and defaults.
    else {
      if (item.fieldType === 'date') value = new Date().toLocaleDateString('en-AU', { timeZone: "Australia/Brisbane" });
      if (item.fieldType === 'title') value = DocumentApp.getActiveDocument().getName();
      if (item.fieldType === 'numeric') value = 0;
    }
    
    let snippetTemplate = HtmlService.createTemplate(templateMap[item.fieldType]);
    snippetTemplate.id = id;
    snippetTemplate.label = item.columnName;
    snippetTemplate.value = value;
    if (item.fieldType === 'dropdown') {
      snippetTemplate.options = item.options || [];
    }
    formHtml += snippetTemplate.evaluate().getContent();
  });

  const finalHtml = buildFinalSidebarHtml(formHtml, configMap);
  const htmlOutput = HtmlService.createHtmlOutput(finalHtml).setTitle('Log Session');
  DocumentApp.getUi().showSidebar(htmlOutput);
}

/**
 * Saves the user's defined configuration map to document properties.
 * @param {Array} configMap An array of objects defining the user's choices.
 */
function saveTrackerConfiguration(configMap) {
  if (!configMap || !Array.isArray(configMap)) {
    throw new Error("Invalid configuration data received.");
  }
  const configString = JSON.stringify(configMap);
  PropertiesService.getDocumentProperties().setProperty('trackerConfigMap', configString);
}

/**
 * DYNAMIC V2.1 (Find-or-Append Logic Restored): Logs data based on the saved config map,
 * finding the correct date row before appending.
 */
function logSessionToSheet(formData) {
  const properties = PropertiesService.getDocumentProperties();
  const sheetId = properties.getProperty('trackerSheetId');
  const sheetName = properties.getProperty('trackerSheetName');
  const configString = properties.getProperty('trackerConfigMap');

  if (!sheetId || !sheetName || !configString) {
    throw new Error("Tracker is not configured.");
  }
  
  const sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);
  if (!sheet) { throw new Error(`Sheet named "${sheetName}" not found.`); }
  
  const configMap = JSON.parse(configString);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  let targetRow = 0;

  // --- NEW: Find-or-Append Logic ---
  // 1. Find which column is our "Date" column from the user's map.
  const dateConfig = configMap.find(item => item.fieldType === 'date');
  
  if (dateConfig) {
    const dateColumnIndex = headers.indexOf(dateConfig.columnName);
    const safeDateId = dateConfig.columnName.replace(/[^a-zA-Z0-9]/g, '');
    const formDate = formData[safeDateId];

    if (dateColumnIndex !== -1) {
      // 2. Search that specific column for the date.
      const dateValues = sheet.getRange(2, dateColumnIndex + 1, sheet.getLastRow() - 1, 1).getDisplayValues();
      for (let i = 0; i < dateValues.length; i++) {
        // Also check if ANY cell in that row has data. We look for Time Start (assuming it's column C for now as a fallback)
        // A better check might be needed if columns are very different.
        if (dateValues[i][0] === formDate && sheet.getRange(i + 2, 3).getValue() === '') {
          targetRow = i + 2;
          break;
        }
      }
    }
  }

  // Create a full-size empty array for the new row
  let newRowData = new Array(headers.length).fill('');
  
  // Place each piece of form data into the correct column based on the map
  configMap.forEach(item => {
    const columnIndex = headers.indexOf(item.columnName);
    if (columnIndex !== -1) {
      const safeId = item.columnName.replace(/[^a-zA-Z0-9]/g, '');
      newRowData[columnIndex] = formData[safeId];
    }
  });
  
  // 3. Decide whether to write to the found row or append a new one.
  if (targetRow > 0) {
    // We found a target row. We must do a "surgical write" to preserve formulas.
    configMap.forEach(item => {
      const columnIndex = headers.indexOf(item.columnName);
      if (columnIndex !== -1) {
        const safeId = item.columnName.replace(/[^a-zA-Z0-9]/g, '');
        // Write each value to its cell individually.
        sheet.getRange(targetRow, columnIndex + 1).setValue(formData[safeId]);
      }
    });
  } else {
    // No target row found, so it's safe to append the whole array.
    sheet.appendRow(newRowData);
  }
  
  clearTrackerState();
}

/**
 * Saves the current state of the tracker form to properties.
 * This is called by the sidebar's JavaScript.
 * @param {object} formData The data from the sidebar form fields.
 */
function saveTrackerState(formData) {
  PropertiesService.getDocumentProperties().setProperty('trackerFormData', JSON.stringify(formData));
}

/**
 * Clears the saved tracker form data.
 */
function clearTrackerState() {
  PropertiesService.getDocumentProperties().deleteProperty('trackerFormData');
}

/**
 * Reads the first row of a sheet to get its column headers.
 * @param {string} sheetId The ID of the spreadsheet.
 * @param {string} sheetName The name of the sheet/tab.
 * @returns {Array} An array of strings representing the column headers.
 */
function getSheetHeaders(sheetId, sheetName) {
  try {
    const sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);
    if (!sheet) {
      throw new Error(`Sheet named "${sheetName}" not found.`);
    }
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    return headers.filter(header => header !== "");
  } catch (e) {
    DocumentApp.getUi().alert('Error reading sheet', e.message, DocumentApp.getUi().ButtonSet.OK);
    return null;
  }
}

/**
 * Finds the last non-empty value in a specific column.
 * @param {string} sheetId The ID of the Google Sheet.
 * @param {string} sheetName The name of the target sheet/tab.
 * @param {string} columnName The header text of the column to search.
 * @returns {string|null} The last value found, or null if the column is empty or not found.
 */
function getLastValueFromColumn(sheetId, sheetName, columnName) {
  try {
    const sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);
    if (!sheet) return null;

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const columnIndex = headers.indexOf(columnName);

    if (columnIndex === -1) return null;

    const columnValues = sheet.getRange(2, columnIndex + 1, sheet.getLastRow() - 1, 1).getValues();
    const nonEmptyValues = columnValues.flat().filter(String);

    if (nonEmptyValues.length > 0) {
      return nonEmptyValues[nonEmptyValues.length - 1];
    } else {
      return null;
    }
  } catch (e) {
    console.error(`Error fetching last value for column "${columnName}": ${e.toString()}`);
    return null;
  }
}

/**
 * Helper function to build the final HTML for the dynamic sidebar.
 */
function buildFinalSidebarHtml(formHtml, configMap) {
  // Create a list of all form field IDs for our JavaScript functions
  const fieldIds = configMap.map(item => item.columnName.replace(/[^a-zA-Z0-9]/g, ''));

  return `
    <!DOCTYPE html>
    <html>
    <head>
        <base target="_top">
        <style>
            body { font-family: Arial, sans-serif; margin: 0; padding: 12px; font-size: 14px; background-color: #f8f9fa; }
            .card { border: 1px solid #ddd; border-radius: 8px; padding: 12px; background-color: #ffffff; }
            .form-group { margin-bottom: 12px; }
            label { display: block; font-weight: bold; margin-bottom: 4px; font-size: 0.9em; }
            input, select, textarea { width: 95%; padding: 8px; border: 1px solid #ccc; border-radius: 4px; }
            textarea { resize: vertical; min-height: 60px; }
            button { width: 100%; padding: 10px; font-size: 1em; color: white; border: none; border-radius: 4px; cursor: pointer; background-color: #007bff; }
            .status-message { text-align: center; padding: 8px; margin-top: 10px; border-radius: 4px; display: none; }
            .status-success { background-color: #d4edda; color: #155724; }
            .status-error { background-color: #f8d7da; color: #721c24; }
            .time-input-group { display: flex; align-items: center; }
            .time-input-group input { width: 75%; }
            .time-input-group .capture-btn { width: auto; margin-left: 8px; padding: 6px 10px; font-size: 0.9em; background-color: #6c757d; }
            #autosave-status { font-size: 0.8em; color: #6c757d; text-align: right; height: 1em; transition: opacity 0.5s; }
        </style>
    </head>
    <body>
        <div class="card">
            <form id="log-form">
                ${formHtml}
                <button type="button" id="logButton" onclick="logSession()">Log to Sheet</button>
                <div id="status" class="status-message"></div>
                <div id="autosave-status"></div>
            </form>
        </div>
    <script>
        window.onload = function() {
            document.getElementById('log-form').addEventListener('change', autoSaveChanges);
        };
        function captureTime(fieldId) {
            const now = new Date();
            const hours = String(now.getHours()).padStart(2, '0');
            const minutes = String(now.getMinutes()).padStart(2, '0');
            const seconds = String(now.getSeconds()).padStart(2, '0');
            const timeField = document.getElementById(fieldId);
            timeField.value = \`\${hours}:\${minutes}:\${seconds}\`;
            timeField.dispatchEvent(new Event('change'));
        }
        function getCurrentFormData() {
            const data = {};
            const fieldIds = ${JSON.stringify(fieldIds)};
            fieldIds.forEach(id => {
                const el = document.getElementById(id);
                if (!el) return;
                if (el.type === 'checkbox') {
                    data[id] = el.checked;
                } else {
                    data[id] = el.value;
                }
            });
            return data;
        }
        function autoSaveChanges() {
            const statusEl = document.getElementById('autosave-status');
            statusEl.textContent = 'Saving...';
            statusEl.style.opacity = '1';
            const formData = getCurrentFormData();
            google.script.run
              .withSuccessHandler(() => setTimeout(() => statusEl.style.opacity = '0', 1000))
              .saveTrackerState(formData);
        }
        function logSession() {
            const btn = document.getElementById('logButton');
            const statusEl = document.getElementById('status');
            btn.disabled = true;
            btn.textContent = 'Logging...';
            const formData = getCurrentFormData();
            google.script.run
              // --- THIS IS THE REVERTED CODE ---
              .withSuccessHandler(function() {
                statusEl.textContent = 'Logged successfully!';
                statusEl.className = 'status-message status-success';
                statusEl.style.display = 'block';
                document.getElementById('log-form').reset();
                // This is a bit tricky since the form is dynamic, so we just clear it
                // and let the user re-open for a fresh pre-filled form.

                setTimeout(function() {
                  btn.disabled = false;
                  btn.textContent = 'Log to Sheet';
                  statusEl.style.display = 'none';
                  // To see the new pre-filled form, the user should re-open the sidebar.
                  // We can't easily repopulate dynamically here without reloading.
                }, 2500);
              })
              // --- END OF REVERTED CODE ---
              .withFailureHandler(err => {
                  alert('Error: ' + err.message);
                  btn.disabled = false;
                  btn.textContent = 'Log to Sheet';
              })
              .logSessionToSheet(formData);
        }
    </script>
    </body>
    </html>
  `;
}
/**
 * NEW: Disconnects the tracker by deleting all relevant document properties.
 */
function disconnectAndResetTracker() {
  const properties = PropertiesService.getDocumentProperties();
  properties.deleteProperty('trackerSheetId');
  properties.deleteProperty('trackerSheetName');
  properties.deleteProperty('trackerConfigMap');
  properties.deleteProperty('trackerFormData');
}
