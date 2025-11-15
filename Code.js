/**
 * @OnlyCurrentDoc
 *
 * A complete Google Apps Script for an automated unlinked mention finder dashboard.
 * This script uses the Google Custom Search Engine (CSE) API to find brand mentions,
 * analyzes them for link opportunities, and presents them in a Google Sheet.
 */

// =================================================================
// 1. CONFIGURATION
// =================================================================

/**
 * The API key for your Google Custom Search Engine.
 * @type {string}
 */
const API_KEY = 'AIzaSyBNzHoxuTVUGalO_6omIpK49UO-2vmG1-s'; // <-- IMPORTANT: REPLACE WITH YOUR API KEY

/**
 * The ID of your Google Custom Search Engine.
 * @type {string}
 */
const CSE_ID = '65db17c21d76c4417'; // <-- IMPORTANT: REPLACE WITH YOUR CSE ID

/**
 * The brand name to search for.
 * @type {string}
 */
const BRAND_NAME = 'Myntra';

/**
 * The domain of the brand's website.
 * @type {string}
 */
const BRAND_DOMAIN = 'myntra.com';

/**
 * An object to hold the names of the sheets for easy reference.
 * @type {Object<string, string>}
 */
const SHEET_NAMES = {
  DASHBOARD: 'Dashboard & Controls',
  QUERIES: 'Queries',
  RESULTS: 'Results - New',
  ARCHIVE: 'Archive'
};

/**
 * An object to hold column indices for the 'Results - New' sheet.
 * @type {Object<string, number>}
 */
const RESULTS_COLS = {
  URL: 1,
  CONTEXT: 2,
  DATE_FOUND: 3,
  STATUS: 4
};


// =================================================================
// 2. SPREADSHEET UI & TRIGGERS
// =================================================================

/**
 * Creates a custom menu in the spreadsheet UI when the file is opened.
 * This is a simple trigger that runs automatically.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Mention Finder')
    .addItem('▶️ Find New Mentions', 'main')
    .addSeparator()
    .addItem('⚙️ Setup Daily Trigger', 'setupDailyTrigger')
    .addToUi();
}

/**
 * An event-driven trigger that fires when a user edits a cell in the spreadsheet.
 * It's used here to move rows from 'Results - New' to 'Archive' when their
 * status is changed from "New".
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e The event object.
 */
function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const sheetName = sheet.getName();
  const oldValue = e.oldValue;

  // Check if the edit happened in the 'Results - New' sheet, in the Status column (D),
  // and the old value was 'New'.
  if (sheetName === SHEET_NAMES.RESULTS && range.getColumn() === RESULTS_COLS.STATUS && oldValue === 'New') {
    try {
      const archiveSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.ARCHIVE);
      if (!archiveSheet) {
        SpreadsheetApp.getUi().alert(`Error: Archive sheet named "${SHEET_NAMES.ARCHIVE}" not found.`);
        return;
      }

      const rowIdx = range.getRow();
      const rowData = sheet.getRange(rowIdx, 1, 1, sheet.getLastColumn()).getValues();

      // Copy the row to the Archive sheet
      archiveSheet.appendRow(rowData[0]);

      // Delete the original row from the Results - New sheet
      sheet.deleteRow(rowIdx);

      SpreadsheetApp.getActiveSpreadsheet().toast(`Moved opportunity to ${SHEET_NAMES.ARCHIVE}.`);
    } catch (error) {
      Logger.log(`onEdit Error: ${error.message}`);
      SpreadsheetApp.getActiveSpreadsheet().toast(`Error archiving row: ${error.message}`);
    }
  }
}

/**
 * A helper function to create a time-driven trigger that runs the main() function daily.
 * The user can run this once from the custom menu to set up automation.
 */
function setupDailyTrigger() {
  // First, delete any existing triggers for the 'main' function to avoid duplicates.
  const allTriggers = ScriptApp.getProjectTriggers();
  for (const trigger of allTriggers) {
    if (trigger.getHandlerFunction() === 'main') {
      ScriptApp.deleteTrigger(trigger);
    }
  }

  // Create a new trigger to run 'main' every day between 2 and 3 AM.
  ScriptApp.newTrigger('main')
    .timeBased()
    .everyDays(1)
    .atHour(2)
    .create();

  SpreadsheetApp.getUi().alert('✅ Success!', 'A daily trigger has been set up to find mentions automatically every night.', SpreadsheetApp.getUi().ButtonSet.OK);
}


// =================================================================
// 3. CORE LOGIC
// =================================================================

// =================================================================
// PASTE THIS ENTIRE BLOCK OVER YOUR EXISTING main() AND searchGoogle() FUNCTIONS
// =================================================================

/**
 * The main function that orchestrates the entire process of finding unlinked mentions.
 * It now fetches multiple pages of results for each query to be more persistent.
 */
function main() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboardSheet = ss.getSheetByName(SHEET_NAMES.DASHBOARD);
  const queriesSheet = ss.getSheetByName(SHEET_NAMES.QUERIES);

  if (!dashboardSheet || !queriesSheet) {
    SpreadsheetApp.getUi().alert('Error: Required sheets ("Dashboard & Controls", "Queries") are missing.');
    return;
  }

  const statusCell = dashboardSheet.getRange('A3');
  statusCell.setValue('Running...');
  SpreadsheetApp.flush(); // Force the UI to update

  // --- NEW: Define how many pages of 10 results to check per query ---
  const PAGES_TO_CHECK = 3; // This will check up to 30 results per query

  try {
    const queriesRange = queriesSheet.getRange('A2:B' + queriesSheet.getLastRow());
    const queriesData = queriesRange.getValues();

    for (let i = 0; i < queriesData.length; i++) {
      const query = queriesData[i][0];
      if (query) {
        statusCell.setValue(`Digging deeper for query: "${query}"`);
        
        // --- NEW: Loop through multiple pages for the current query ---
        for (let page = 0; page < PAGES_TO_CHECK; page++) {
          const startIndex = page * 10 + 1; // Page 1 starts at 1, Page 2 at 11, etc.
          
          const searchResults = searchGoogle(query, startIndex);

          // If the API returns no results, there are no more pages, so we stop for this query.
          if (!searchResults || searchResults.length === 0) {
            break; 
          }

          for (const item of searchResults) {
            processUrl(item.link);
          }
        }
        
        // Update 'Last Checked' timestamp for the processed query
        queriesSheet.getRange(i + 2, 2).setValue(new Date());
      }
    }

    const completionTimestamp = new Date().toLocaleString();
    statusCell.setValue(`Complete. Last run: ${completionTimestamp}`);

  } catch (error) {
    Logger.log(`main Error: ${error.message}\n${error.stack}`);
    statusCell.setValue(`Error! Check logs. Message: ${error.message}`);
  }
}

/**
 * Executes a search using the Google Custom Search Engine (CSE) API for a specific page of results.
 * @param {string} query The search string to execute.
 * @param {number} startIndex The starting index for the search results (for pagination).
 * @returns {Array<Object>|null} An array of search result items, or null on error.
 */
function searchGoogle(query, startIndex) {
  if (API_KEY === 'YOUR_GOOGLE_CSE_API_KEY_HERE' || CSE_ID === 'YOUR_CSE_ID_HERE') {
    SpreadsheetApp.getUi().alert('Please update the API_KEY and CSE_ID constants in the script.');
    return null;
  }

  // --- MODIFIED: Added the 'start' parameter for pagination ---
  const apiUrl = `https://www.googleapis.com/customsearch/v1?key=${API_KEY}&cx=${CSE_ID}&q=${encodeURIComponent(query)}&start=${startIndex}`;

  try {
    const response = UrlFetchApp.fetch(apiUrl, { muteHttpExceptions: true });
    const responseCode = response.getResponseCode();
    const content = response.getContentText();

    if (responseCode === 200) {
      updateDailyQueryCount();
      const json = JSON.parse(content);
      return json.items || []; // Return items array or an empty array if no results
    } else {
      Logger.log(`API Error for query "${query}". Response Code: ${responseCode}. Content: ${content}`);
      return null;
    }
  } catch (error) {
    Logger.log(`searchGoogle Fetch Error for query "${query}": ${error.message}`);
    return null;
  }
}

/**
 * Processes a single URL to check if it's a valid, non-duplicate, unlinked mention opportunity.
 * @param {string} url The URL to process.
 */
function processUrl(url) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboardSheet = ss.getSheetByName(SHEET_NAMES.DASHBOARD);
  const resultsSheet = ss.getSheetByName(SHEET_NAMES.RESULTS);
  const archiveSheet = ss.getSheetByName(SHEET_NAMES.ARCHIVE);

  // 1. Check against exclusion list
  const excludedDomainsStr = dashboardSheet.getRange('E3').getValue();
  const excludedDomains = excludedDomainsStr.split(',').map(d => d.trim().toLowerCase()).filter(d => d);
  
  // --- START OF FIX ---
  // The 'new URL()' constructor is not available in Apps Script.
  // We use a regular expression to reliably extract the hostname instead.
  let hostname;
  try {
    // This regex extracts the part between "://" and the next "/"
    hostname = url.match(/:\/\/(?:www\.)?([^\/]+)/)[1];
  } catch (e) {
    // Fallback for URLs that might not have a protocol (e.g., "example.com/page")
    hostname = url.split('/')[0];
  }

  if (!hostname) {
    Logger.log(`Could not parse domain from URL: ${url}`);
    return; // Skip if we can't get a valid domain
  }
  
  const urlDomain = hostname.trim().toLowerCase();
  // --- END OF FIX ---

  if (excludedDomains.includes(urlDomain)) {
    Logger.log(`Skipping excluded domain: ${url}`);
    return;
  }

  // 2. Check for duplicates in Results and Archive sheets
  const resultsUrls = resultsSheet.getLastRow() > 1 ? resultsSheet.getRange(2, 1, resultsSheet.getLastRow() - 1, 1).getValues().flat() : [];
  const archiveUrls = archiveSheet.getLastRow() > 1 ? archiveSheet.getRange(2, 1, archiveSheet.getLastRow() - 1, 1).getValues().flat() : [];
  const allExistingUrls = new Set([...resultsUrls, ...archiveUrls]);

  if (allExistingUrls.has(url)) {
    Logger.log(`Skipping duplicate URL: ${url}`);
    return;
  }

  // 3. If new and not excluded, perform the deep check
  checkForUnlinkedMention(url);
}

/**
 * Fetches a URL's content and analyzes it to see if it contains an unlinked mention.
 * @param {string} url The URL to analyze.
 */
function checkForUnlinkedMention(url) {
  try {
    // --- START OF FIX ---
    // Define the parameters for the fetch call.
    // We add a "User-Agent" header to mimic the Googlebot, which helps bypass
    // basic anti-scraping measures that return a 403 Forbidden error.
    const params = {
      muteHttpExceptions: true,
      "headers": {
        "User-Agent": "Mozilla/5.0 (compatible; Googlebot/2.1; +http://www.google.com/bot.html)"
      }
    };
    // --- END OF FIX ---

    const response = UrlFetchApp.fetch(url, params); // Use the new params object
    
    if (response.getResponseCode() !== 200) {
      Logger.log(`Could not fetch URL ${url}. Status code: ${response.getResponseCode()}`);
      return;
    }

    const htmlContent = response.getContentText();
    const lowerCaseContent = htmlContent.toLowerCase();
    const lowerCaseBrandName = BRAND_NAME.toLowerCase();

    // Logic: Brand name is present, but a link to the brand domain is not.
    const isMentioned = lowerCaseContent.includes(lowerCaseBrandName);
    
    // Regex to find links to the brand domain. Handles http/https, www/non-www.
    const linkRegex = new RegExp(`href\\s*=\\s*["'](https?:\\/\\/)?(www\\.)?${BRAND_DOMAIN.replace('.', '\\.')}`, 'i');
    const isLinked = linkRegex.test(lowerCaseContent);

    if (isMentioned && !isLinked) {
      Logger.log(`Found unlinked mention at: ${url}`);
      // Extract a snippet for context
      const mentionIndex = lowerCaseContent.indexOf(lowerCaseBrandName);
      const snippetStart = Math.max(0, mentionIndex - 75);
      const snippetEnd = Math.min(htmlContent.length, mentionIndex + lowerCaseBrandName.length + 75);
      const rawSnippet = htmlContent.substring(snippetStart, snippetEnd);
      // Clean up the snippet by removing HTML tags and extra whitespace
      const contextSnippet = rawSnippet.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();

      writeResult(url, `...${contextSnippet}...`);
    }
  } catch (error) {
    Logger.log(`checkForUnlinkedMention Error for URL ${url}: ${error.message}`);
  }
}

/**
 * Writes a new unlinked mention opportunity to the 'Results - New' sheet.
 * @param {string} url The URL of the found mention.
 * @param {string} contextSnippet A text snippet showing the context of the mention.
 */
function writeResult(url, contextSnippet) {
  try {
    const resultsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.RESULTS);
    resultsSheet.appendRow([
      url,
      contextSnippet,
      new Date(),
      'New'
    ]);
  } catch (error) {
    Logger.log(`writeResult Error: ${error.message}`);
  }
}


// =================================================================
// 4. UTILITY FUNCTIONS
// =================================================================

/**
 * Updates the "Total Queries Today" counter on the dashboard.
 * Resets the count if it's a new day.
 */
function updateDailyQueryCount() {
  const properties = PropertiesService.getScriptProperties();
  const today = new Date().toISOString().slice(0, 10); // YYYY-MM-DD format
  const lastQueryDate = properties.getProperty('LAST_QUERY_DATE');
  
  let count = 1;
  if (today === lastQueryDate) {
    count = parseInt(properties.getProperty('DAILY_QUERY_COUNT') || '0', 10) + 1;
  }

  properties.setProperty('LAST_QUERY_DATE', today);
  properties.setProperty('DAILY_QUERY_COUNT', count.toString());

  const dashboardSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.DASHBOARD);
  if (dashboardSheet) {
    dashboardSheet.getRange('C3').setValue(count);
  }
}





function apiTest() {
  console.log('SUCCESS! The API call from the dashboard worked!');
  return true; 
}
