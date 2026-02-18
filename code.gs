// ===============================================================
//                          CONFIGURATION
// ===============================================================
// (recipientEmail has been REMOVED from here and moved to PropertiesService)
const dataSheetName = "Data";
const userSheetName = "Users"; // Headers: Username, PasswordHash, Salt, Role, Status, Team, AllowedTypes
const teamSheetName = "Teams"; // NEW: Headers: TeamName
const translationSheetName = "Translations";
const profanitySheetName = "ProfanityWords"; // Header: Word
const templateSheetName = "Templates"; // Headers: Title, Message
const auditLogSheetName = "AuditLog"; 
const uploadFolderId = "1Nqxi63gJtXPCCpe3NB5Y4X0qQ_a1C89A"; // Replace with actual Folder ID

// ===============================================================
//                  SECURITY HELPER FUNCTIONS
// ===============================================================

/**
 * Retrieves session data associated with a token.
 * This is the core function for Multi-Session support.
 * @param {string} token The client-side session token.
 * @returns {object|null} The session object {role, username, allowedTypes} or null if invalid.
 */
function getSessionData(token) {
  if (!token) return null;
  const userCache = CacheService.getUserCache();
  const sessionKey = 'admin_session_' + token;
  const sessionJson = userCache.get(sessionKey);
  
  if (!sessionJson) return null;
  
  try {
    return JSON.parse(sessionJson);
  } catch (e) {
    Logger.log("Error parsing session data: " + e);
    return null;
  }
}

/**
 * Checks if a user is authenticated (token is valid), regardless of role.
 * Also extends the session expiry.
 * @param {string} token The client-side session token.
 * @returns {boolean} True if the token is valid, false otherwise.
 */
function isUserAuthenticated(token) {
  if (!token) return false;
  
  const userCache = CacheService.getUserCache();
  const sessionKey = 'admin_session_' + token;
  const sessionJson = userCache.get(sessionKey);

  // Check if session exists
  if (!sessionJson) {
    return false;
  }

  // Extend session expiry
  userCache.put(sessionKey, sessionJson, 3600); // Extend for 1 hour
  return true;
}

/**
 * Checks if an admin user is authenticated (token valid AND role is 'Admin').
 * Relies on isUserAuthenticated to extend the session.
 * @param {string} token The client-side session token.
 * @returns {boolean} True if the token is valid AND role is Admin, false otherwise.
 */
function isUserAdmin(token) {
  // First, check/extend validity
  if (!isUserAuthenticated(token)) {
    return false;
  }

  const session = getSessionData(token);
  if (!session || session.role !== 'Admin') {
    Logger.log(`Authorization failed. Role is '${session ? session.role : 'N/A'}' (Expected 'Admin').`);
    return false;
  }

  return true;
}

/**
 * Logs out the user by removing their specific session key.
 * @param {string} token The client-side session token to invalidate.
 * @returns {object} Success status.
 */
function logout(token) {
  try {
    if (token) {
      const userCache = CacheService.getUserCache();
      const sessionKey = 'admin_session_' + token;
      userCache.remove(sessionKey);
      Logger.log('User logged out, session cleared: ' + sessionKey);
    }
    return { success: true };
  } catch (e) {
    Logger.log(`Error during logout: ${e}`);
    return { success: true }; 
  }
}

// ===============================================================
//                  SYSTEM AUDIT LOG
// ===============================================================

/**
 * Internal helper function to log administrative actions to the 'AuditLog' sheet.
 * This function MUST NOT throw errors, as it's a non-critical logging task.
 * @param {string} username The username (from session) performing the action.
 * @param {string} action A short description of the action (e.g., "Add User").
 * @param {string} details Specific details of the action.
 */
function _logAdminAction(username, action, details) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let logSheet = ss.getSheetByName(auditLogSheetName);
    const expectedHeader = ["Timestamp", "Username", "Action", "Details"];

    // 1. Check if sheet exists
    if (!logSheet) {
      logSheet = ss.insertSheet(auditLogSheetName);
      logSheet.appendRow(expectedHeader);
    } 
    // 2. Check if header is correct
    else if (logSheet.getLastRow() < 1) {
      logSheet.appendRow(expectedHeader);
    } 
    else {
      const header = logSheet.getRange(1, 1, 1, 4).getValues()[0];
      if (JSON.stringify(header) !== JSON.stringify(expectedHeader)) {
        logSheet.getRange(1, 1, 1, 4).setValues([expectedHeader]);
      }
    }

    // 3. Append the new log entry
    logSheet.appendRow([
      new Date(), // Timestamp
      username || "(Unknown User)", // Username
      action || "Unknown Action", // Action
      details || "" // Details
    ]);

  } catch (e) {
    Logger.log(`CRITICAL: Failed to write to AuditLog. Error: ${e}`);
  }
}


// ===============================================================
//                  PROFANITY FILTER FUNCTIONS
// ===============================================================

function getProfanityList() {
  const cache = CacheService.getScriptCache();
  const cachedListKey = 'profanityList_v3'; 
  const cachedList = cache.get(cachedListKey);
  if (cachedList) {
    try {
      return JSON.parse(cachedList);
    } catch (e) {
      cache.remove(cachedListKey);
    }
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(profanitySheetName);

    if (!sheet) {
      sheet = ss.insertSheet(profanitySheetName);
      sheet.appendRow(['Word']);
      cache.put(cachedListKey, JSON.stringify([]), 600);
      return [];
    }

    if (sheet.getLastRow() >= 1) {
       const header = sheet.getRange(1, 1).getValue();
       if (header !== 'Word') {
          sheet.getRange(1, 1).setValue('Word');
       }
    } else {
       sheet.appendRow(['Word']);
       cache.put(cachedListKey, JSON.stringify([]), 600);
       return [];
    }

    if (sheet.getLastRow() < 2) {
      cache.put(cachedListKey, JSON.stringify([]), 600);
      return [];
    }

    const words = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1)
                      .getValues()
                      .flat()
                      .map(word => String(word).trim().toLowerCase())
                      .filter(word => word);

    cache.put(cachedListKey, JSON.stringify(words), 600);
    return words;

  } catch (error) {
    Logger.log(`Error in getProfanityList: ${error}`);
    return [];
  }
}

function saveProfanityList(token, words) {
  if (!isUserAuthenticated(token)) { 
    return { success: false, error: "Authentication failed. Please log in again." };
  }
  if (!Array.isArray(words)) {
      return { success: false, error: "Invalid data format. Expected an array of words." };
  }

  try {
    const oldWords = getProfanityList(); 

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(profanitySheetName);

    if (!sheet) {
      sheet = ss.insertSheet(profanitySheetName);
      sheet.appendRow(['Word']);
    } else if (sheet.getLastRow() < 1 || sheet.getRange(1, 1).getValue() !== 'Word') {
       sheet.clearContents();
       sheet.appendRow(['Word']);
    }

    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).clearContent();
    }

    const uniqueWords = [...new Set(words.map(w => String(w).trim()).filter(w => w))];
    const wordsToInsert = uniqueWords.map(word => [word]); 

    if (wordsToInsert.length > 0) {
      sheet.getRange(2, 1, wordsToInsert.length, 1).setValues(wordsToInsert);
    }

    CacheService.getScriptCache().remove('profanityList_v3');

    // Logging
    const session = getSessionData(token);
    const username = session ? session.username : 'Unknown';
    
    const newWordsSet = new Set(uniqueWords);
    const oldWordsSet = new Set(oldWords);
    const addedWords = uniqueWords.filter(word => !oldWordsSet.has(word));
    const removedWords = oldWords.filter(word => !newWordsSet.has(word));

    let logDetails = "";
    if (addedWords.length > 0) logDetails += `Added: [${addedWords.join(', ')}]`;
    if (removedWords.length > 0) {
      if (logDetails) logDetails += "; ";
      logDetails += `Removed: [${removedWords.join(', ')}]`;
    }
    if (!logDetails) logDetails = "No changes detected.";
    
    _logAdminAction(username, "Save Profanity List", logDetails);

    return { success: true };
  } catch (error) {
    Logger.log(`Error in saveProfanityList: ${error}`);
    return { success: false, error: error.message };
  }
}

function containsProfanity(text) {
    if (!text) return false;
    const lowerCaseText = String(text).toLowerCase();
    const profanityList = getProfanityList(); 
    return profanityList.some(word => lowerCaseText.includes(word));
}

// ===============================================================
//                  TEMPLATE FUNCTIONS (CANNED RESPONSES)
// ===============================================================

function getTemplates(token) {
  if (!isUserAuthenticated(token)) {
      return { error: "Authentication failed. Please log in again." };
  }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(templateSheetName);

    if (!sheet) {
        // Create if not exists
        sheet = ss.insertSheet(templateSheetName);
        sheet.appendRow(['Title', 'Message']); // Header
        return { success: true, data: [] };
    }

    if (sheet.getLastRow() < 2) {
        return { success: true, data: [] };
    }

    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    const templates = data.map(row => ({
        title: row[0],
        message: row[1]
    })).filter(t => t.title); // Filter out empty titles

    return { success: true, data: templates };
  } catch (e) {
    Logger.log(`Error in getTemplates: ${e}`);
    return { error: e.message };
  }
}

function saveTemplate(token, template) {
    if (!isUserAdmin(token)) {
        return { success: false, error: "Authentication failed. Admin role required." };
    }
    if (!template || !template.title || !template.message) {
        return { success: false, error: "Title and Message are required." };
    }

    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        let sheet = ss.getSheetByName(templateSheetName);
        if (!sheet) {
            sheet = ss.insertSheet(templateSheetName);
            sheet.appendRow(['Title', 'Message']);
        }

        const data = sheet.getDataRange().getValues();
        // data[0] is header
        let rowIndex = -1;

        // If editing (originalTitle provided), find by originalTitle
        if (template.originalTitle) {
            for (let i = 1; i < data.length; i++) {
                if (String(data[i][0]) === String(template.originalTitle)) {
                    rowIndex = i + 1;
                    break;
                }
            }
        }
        
        // Check for duplicates if renaming or creating new
        if (!template.originalTitle || template.originalTitle !== template.title) {
            const duplicate = data.some((row, i) => i > 0 && String(row[0]) === String(template.title));
            if (duplicate) {
                return { success: false, error: `Template name "${template.title}" already exists.` };
            }
        }

        if (rowIndex !== -1) {
            // Update
            sheet.getRange(rowIndex, 1, 1, 2).setValues([[template.title, template.message]]);
        } else {
            // Append
            sheet.appendRow([template.title, template.message]);
        }

        // Log action
        const session = getSessionData(token);
        _logAdminAction(session.username, "Save Template", `Title: "${template.title}"`);

        return { success: true };

    } catch (e) {
        Logger.log(`Error in saveTemplate: ${e}`);
        return { success: false, error: e.message };
    }
}

function deleteTemplate(token, title) {
    if (!isUserAdmin(token)) {
        return { success: false, error: "Authentication failed. Admin role required." };
    }
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName(templateSheetName);
        if (!sheet) return { success: false, error: "Sheet not found" };

        const data = sheet.getDataRange().getValues();
        let rowIndex = -1;

        for (let i = 1; i < data.length; i++) {
            if (String(data[i][0]) === String(title)) {
                rowIndex = i + 1;
                break;
            }
        }

        if (rowIndex !== -1) {
            sheet.deleteRow(rowIndex);
             // Log action
            const session = getSessionData(token);
            _logAdminAction(session.username, "Delete Template", `Title: "${title}"`);
            return { success: true };
        } else {
            return { success: false, error: "Template not found" };
        }

    } catch (e) {
        Logger.log(`Error in deleteTemplate: ${e}`);
        return { success: false, error: e.message };
    }
}

// ===============================================================
//                  FORM CONFIGURATION FUNCTIONS
// ===============================================================

function getFormConfig() {
  const props = PropertiesService.getScriptProperties();
  return {
    enableComplaint: props.getProperty('ENABLE_TYPE_COMPLAINT') !== 'false', 
    enableSuggestion: props.getProperty('ENABLE_TYPE_SUGGESTION') !== 'false', 
    enableReport: props.getProperty('ENABLE_TYPE_REPORT') !== 'false' 
  };
}

function saveFormConfig(token, config) {
  if (!isUserAdmin(token)) {
    return { success: false, error: "Authentication failed. Admin role required." };
  }
  try {
    const props = PropertiesService.getScriptProperties();
    props.setProperty('ENABLE_TYPE_COMPLAINT', String(config.enableComplaint));
    props.setProperty('ENABLE_TYPE_SUGGESTION', String(config.enableSuggestion));
    props.setProperty('ENABLE_TYPE_REPORT', String(config.enableReport));
    
    const session = getSessionData(token);
    _logAdminAction(session.username, "Update Form Config", `Updated request type toggles.`);
    
    return { success: true };
  } catch (e) {
    Logger.log("Error saving form config: " + e);
    return { success: false, error: e.message };
  }
}

// ===============================================================
//                  TEAM MANAGEMENT FUNCTIONS (NEW)
// ===============================================================

/**
 * Retrieves all teams from the master 'Teams' sheet.
 * @param {string} token
 * @returns {object} {success: boolean, data: string[]}
 */
function getTeams(token) {
    if (!isUserAuthenticated(token)) return { error: "Authentication failed." };
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        let sheet = ss.getSheetByName(teamSheetName);
        if (!sheet) {
            sheet = ss.insertSheet(teamSheetName);
            sheet.appendRow(['TeamName']);
            return { success: true, data: [] };
        }
        if (sheet.getLastRow() < 2) return { success: true, data: [] };

        const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat();
        const teams = data.map(t => String(t).trim()).filter(t => t);
        return { success: true, data: [...new Set(teams)].sort() }; // Unique & Sorted
    } catch (e) {
        Logger.log("Error getting teams: " + e);
        return { success: false, error: e.message };
    }
}

/**
 * Adds a new team to the master 'Teams' sheet.
 * @param {string} token
 * @param {string} teamName
 */
function saveTeam(token, teamName) {
    if (!isUserAdmin(token)) {
        return { success: false, error: "Authentication failed. Admin role required." };
    }
    const name = String(teamName || '').trim();
    if (!name) return { success: false, error: "Team name cannot be empty." };

    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        let sheet = ss.getSheetByName(teamSheetName);
        if (!sheet) {
            sheet = ss.insertSheet(teamSheetName);
            sheet.appendRow(['TeamName']);
        }

        // Check duplicates
        const existingTeams = getTeams(token).data;
        if (existingTeams.includes(name)) {
            return { success: false, error: `Team "${name}" already exists.` };
        }

        sheet.appendRow([name]);

        const session = getSessionData(token);
        _logAdminAction(session.username, "Add Team", `Team: "${name}"`);

        return { success: true };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

/**
 * Deletes a team from the master 'Teams' sheet.
 * @param {string} token
 * @param {string} teamName
 */
function deleteTeam(token, teamName) {
   if (!isUserAdmin(token)) {
       return { success: false, error: "Authentication failed. Admin role required." };
   }
   const name = String(teamName || '').trim();

   try {
       const ss = SpreadsheetApp.getActiveSpreadsheet();
       const sheet = ss.getSheetByName(teamSheetName);
       if (!sheet) return { success: false, error: "Sheet not found" };

       const data = sheet.getDataRange().getValues();
       let rowIndex = -1;
       
       // Start from 1 to skip header
       for (let i = 1; i < data.length; i++) {
           if (String(data[i][0]).trim() === name) {
               rowIndex = i + 1;
               break;
           }
       }

       if (rowIndex !== -1) {
           sheet.deleteRow(rowIndex);
           const session = getSessionData(token);
           _logAdminAction(session.username, "Delete Team", `Team: "${name}"`);
           return { success: true };
       } else {
           return { success: false, error: "Team not found." };
       }
   } catch (e) {
       return { success: false, error: e.message };
   }
}

// ===============================================================
//                        WEB APP ROUTING & TEMPLATING
// ===============================================================
function doGet(e) {
  const lang = e.parameter.lang || 'th'; // Default language
  let translations = {}; // Initialize as empty object
  try {
    translations = getTranslations(lang); // Get translations (including consentText now)
    // Add default titles if specific keys are missing, preventing template errors
    translations.mainTitle = translations.mainTitle || 'Complaint & Suggestion Form';
    translations.adminTitle = translations.adminTitle || 'Admin Dashboard';
    translations.consentText = translations.consentText || 'Default Consent Text - Please configure in Translations sheet.'; // Default if missing
    translations.profanityWarning = translations.profanityWarning || 'กรุณาใช้ถ้อยคำสุภาพ'; // Use profanityWarning
    translations.fileSizeError = translations.fileSizeError || 'ขนาดไฟล์ต้องไม่เกิน 5MB';
    translations.fileReadError = translations.fileReadError || 'เกิดข้อผิดพลาดในการอ่านไฟล์';
    
    translations.checkStatusFail = translations.checkStatusFail || 'เลขที่คำร้อง หรือ รหัสลับ ไม่ถูกต้อง';
    translations.checkStatusBtn = translations.checkStatusBtn || 'ตรวจสอบสถานะ';

  } catch (err) {
      Logger.log(`ERROR fetching translations in doGet: ${err}`);
      // Use fallback defaults if getTranslations fails entirely
      translations = {
          mainTitle: 'Complaint & Suggestion Form (Error)',
          adminTitle: 'Admin Dashboard (Error)',
          consentText: 'Error loading consent text.',
          profanityWarning: 'กรุณาใช้ถ้อยคำสุภาพ (Error)',
          fileSizeError: 'ขนาดไฟล์ต้องไม่เกิน 5MB (Error)',
          fileReadError: 'เกิดข้อผิดพลาดในการอ่านไฟล์ (Error)',
          checkStatusFail: 'เลขที่คำร้อง หรือ รหัสลับ ไม่ถูกต้อง (Error)',
          checkStatusBtn: 'ตรวจสอบสถานะ (Error)'
      };
  }
  const webAppUrl = ScriptApp.getService().getUrl();
  let logoUrl = null; // Initialize logoUrl
  try {
      logoUrl = getLogo(); // Fetch logo URL
  } catch(err) {
      Logger.log(`ERROR fetching logo in doGet: ${err}`);
  }


  if (e.parameter.page === 'admin') {
    // Serve admin page
    let template = HtmlService.createTemplateFromFile('admin.html'); // Ensure .html extension
    template.webAppUrl = webAppUrl; // Pass URL for linking back
    return template.evaluate()
      .setTitle(translations.adminTitle) // Use translation (or default) for title
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  } else {
    // Serve index page (default)
    let template = HtmlService.createTemplateFromFile('index.html'); // Ensure .html extension
    template.logoUrl = logoUrl; // Pass logo URL (might be null)
    template.t = translations; // Pass all translations (incl. consentText, potentially defaults)
    template.webAppUrl = webAppUrl; // Pass URL for language switching
    template.currentLang = lang; // Pass current language for highlighting flag
    
    // NEW: Inject Form Configuration
    template.formConfig = getFormConfig();
    
    return template.evaluate()
      .setTitle(translations.mainTitle) // Use translation (or default) for title
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  }
}

// ===============================================================
//                    TRANSLATION FUNCTIONS
// ===============================================================

/**
 * Retrieves all translations for a given language from the sheet.
 * Handles missing sheet/language. Caches results.
 * @param {string} lang Language code (e.g., 'th', 'jp', 'my').
 * @returns {object} Key-value pairs of translations. Includes 'consentText'.
 */
function getTranslations(lang) {
  const cache = CacheService.getScriptCache();
  const cacheKey = `translations_${lang}_v2`;
  const cachedTranslations = cache.get(cacheKey);
  if (cachedTranslations) {
    try {
      return JSON.parse(cachedTranslations);
    } catch (e) {
      Logger.log(`Error parsing cached translations for ${lang}: ${e}. Fetching fresh.`);
      cache.remove(cacheKey);
    }
  }

  const translations = {};
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(translationSheetName);
    if (!sheet || sheet.getLastRow() < 2) { // Need header + at least one key
        Logger.log(`Translation sheet '${translationSheetName}' not found or empty.`);
        cache.put(cacheKey, JSON.stringify({}), 300); // Cache empty object for 5 mins
        return {};
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).toLowerCase().trim()); // Read headers, trim whitespace
    const langIndex = headers.indexOf(lang.toLowerCase());
    const keyIndex = headers.indexOf('key'); // Find 'key' column index

     if (keyIndex === -1) {
        Logger.log(`'Key' column not found in translation sheet headers: ${headers}`);
        cache.put(cacheKey, JSON.stringify({}), 300);
        return {};
    }

    if (langIndex === -1) {
        Logger.log(`Language column '${lang}' not found in translation sheet headers: ${headers}`);
        cache.put(cacheKey, JSON.stringify({}), 300); // Cache empty object
        return {};
    }

    // Iterate through data rows (starting from index 1 to skip header)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const key = row[keyIndex] ? String(row[keyIndex]).trim() : null; // Get key, trim
      const text = (langIndex < row.length) ? row[langIndex] : ''; // Get text safely
      if (key) { // Only add if key is not empty
        translations[key] = text || ''; // Use empty string if translation is missing or null
      }
    }

    cache.put(cacheKey, JSON.stringify(translations), 3600); // Cache for 1 hour
    return translations;

  } catch (error) {
    Logger.log(`Error in getTranslations for ${lang}: ${error}`);
    return {}; // Return empty object on error
  }
}


/**
 * Retrieves all translation data (including headers) for the admin view. Requires Admin role.
 * @param {string} token The admin session token.
 * @returns {object|Array<Array<string>>} 2D array of data or {error: string}.
 */
function getAllTranslationsForAdmin(token) {
  if (!isUserAuthenticated(token)) { 
    Logger.log('getAllTranslationsForAdmin failed: Authentication error.');
    return { error: "Authentication failed. Please log in again." };
  }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(translationSheetName);
    // Return empty array if sheet not found, client handles this better
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    // Return empty array if sheet is empty
    return data.length === 0 ? [] : data;
  } catch (error) {
    Logger.log('Error in getAllTranslationsForAdmin: ' + error);
    return { error: error.message }; // Return error object
  }
}

/**
 * Saves all translation data back to the sheet. Requires Admin role.
 * Overwrites the entire sheet.
 * @param {string} token The admin session token.
 * @param {Array<Array<string>>} data The 2D array of translation data (including headers).
 * @returns {object} {success: boolean, error?: string}
 */
function saveAllTranslations(token, data) {
  if (!isUserAuthenticated(token)) { 
      return { success: false, error: "Authentication failed. Please log in again." };
  }
  // Basic data validation
  if (!Array.isArray(data) || data.length === 0 || !Array.isArray(data[0]) || data[0].length === 0) {
      return { success: false, error: "Invalid or empty data provided for saving translations." };
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(translationSheetName);

    if (!sheet) {
        // If sheet doesn't exist, create it. This case might indicate an issue.
        sheet = ss.insertSheet(translationSheetName);
        Logger.log(`Created missing sheet during save: ${translationSheetName}`);
    }

    // Clear the entire sheet before writing new data
    sheet.clearContents();
    // Write data ensuring it fits the sheet dimensions
    const numRows = data.length;
    const numCols = data[0].length; // Assume all rows have same length as header
    // Resize sheet if necessary BEFORE writing
    if (sheet.getMaxRows() < numRows) sheet.insertRowsAfter(sheet.getMaxRows(), numRows - sheet.getMaxRows());
    if (sheet.getMaxColumns() < numCols) sheet.insertColumnsAfter(sheet.getMaxColumns(), numCols - sheet.getMaxColumns());
    // Get range matching data size
    sheet.getRange(1, 1, numRows, numCols).setValues(data);


    // Clear relevant caches after saving
    const languages = ['th', 'jp', 'my']; // Add other languages if needed
    const cache = CacheService.getScriptCache();
    // MODIFIED: Updated cache version to v2
    languages.forEach(lang => cache.remove(`translations_${lang}_v2`));

    // Clear data cache as well, since translations might affect new entries
    cache.remove('all_data_headers_v2');
    Logger.log('Cleared headers cache due to translation update.');


    Logger.log(`Saved translations (${numRows} rows, ${numCols} cols). Cleared caches for languages: ${languages.join(', ')}`);
    
    // --- START: AUDIT LOG ---
    const session = getSessionData(token);
    _logAdminAction(session.username, "Save Translations", `Saved ${numRows} rows and ${numCols} columns of translations.`);
    // --- END: AUDIT LOG ---
    
    return { success: true };

  } catch (error) {
    Logger.log(`Error in saveAllTranslations: ${error}`);
    return { success: false, error: error.message };
  }
}

// ===============================================================
//                    PROPERTIES SERVICE FUNCTIONS (Logo only)
// ===============================================================

/**
 * Saves the logo URL to Script Properties. Requires Admin role.
 * @param {string} token The admin session token.
 * @param {string} url The URL of the logo image.
 * @returns {object} {success: boolean, error?: string}
 */
function saveLogo(token, url) {
  if (!isUserAuthenticated(token)) { 
      return { success: false, error: "Authentication failed. Please log in again." };
  }
  try {
      const trimmedUrl = String(url || '').trim(); // Trim and handle null/undefined
      // Basic URL validation (optional but recommended)
      if (trimmedUrl && !(trimmedUrl.startsWith('http://') || trimmedUrl.startsWith('https://'))) {
        // Allow empty URL to remove logo
        return { success: false, error: "Invalid URL format. Must start with http:// or https://" };
      }
      PropertiesService.getScriptProperties().setProperty('LOGO_URL', trimmedUrl); // Save trimmed or empty string
      Logger.log(`Logo URL saved: ${trimmedUrl || '(empty)'}`);
      
      // --- START: AUDIT LOG ---
      const session = getSessionData(token);
      _logAdminAction(session.username, "Save Logo", `Set URL to: ${trimmedUrl || '(empty)'}`);
      // --- END: AUDIT LOG ---
      
      return { success: true };
  } catch (error) {
      Logger.log(`Error saving logo URL: ${error}`);
      return { success: false, error: error.message };
  }
}

/**
 * Retrieves the logo URL from Script Properties. Publicly accessible.
 * @returns {string|null} The logo URL or null if not set or error.
 */
function getLogo() {
  try {
    return PropertiesService.getScriptProperties().getProperty('LOGO_URL');
  } catch (error) {
    Logger.log(`Error getting logo URL: ${error}`);
    return null; // Return null on error
  }
}

// ===== START: NEW CORE SETTINGS FUNCTIONS (v10) =====
/**
 * Retrieves core system settings (like email) for the admin panel.
 * REQUIRES ADMIN ROLE.
 * @param {string} token The admin session token.
 * @returns {object} {success: boolean, settings: {email}} or {success: false, error: string}
 */
function getCoreSettings(token) {
  if (!isUserAdmin(token)) { // Must be Admin
      return { success: false, error: "Authentication failed. Requires Admin role." };
  }
  try {
      const props = PropertiesService.getScriptProperties();
      const settings = {
          email: props.getProperty('RECIPIENT_EMAIL') || '' // Get email, default to empty string
      };
      return { success: true, settings: settings };
  } catch (error) {
      Logger.log(`Error in getCoreSettings: ${error}`);
      return { success: false, error: error.message };
  }
}

/**
 * Saves core system settings (like email) from the admin panel.
 * REQUIRES ADMIN ROLE.
 * @param {string} token The admin session token.
 * @param {object} settings The settings object. e.g., { email: 'a@b.com,c@d.com' }
 * @returns {object} {success: boolean, error?: string}
 */
function saveCoreSettings(token, settings) {
  if (!isUserAdmin(token)) { // Must be Admin
      return { success: false, error: "Authentication failed. Requires Admin role." };
  }
  try {
      if (!settings) {
          throw new Error("Invalid settings object.");
      }
      
      const props = PropertiesService.getScriptProperties();
      
      // Validate and save Email
      const trimmedEmail = String(settings.email || '').trim();
      
      // Simple validation: allow empty string, or must contain '@'
      if (trimmedEmail && !trimmedEmail.includes('@')) {
          throw new Error("Invalid Recipient Email format. Must contain @ or be empty.");
      }
      
      // Get old value for logging
      const oldEmail = props.getProperty('RECIPIENT_EMAIL') || '';
      
      // Save the new value
      props.setProperty('RECIPIENT_EMAIL', trimmedEmail);
      Logger.log(`Recipient Email saved: ${trimmedEmail || '(empty)'}`);
      
      // --- START: AUDIT LOG ---
      const session = getSessionData(token);
      _logAdminAction(session.username, "Save Core Settings", `Changed Recipient Email from "${oldEmail}" to "${trimmedEmail || '(empty)'}"`);
      // --- END: AUDIT LOG ---
      
      return { success: true };
  } catch (error) {
      Logger.log(`Error saving core settings: ${error}`);
      return { success: false, error: error.message };
  }
}
// ===== END: NEW CORE SETTINGS FUNCTIONS (v10) =====


// ===============================================================
//                        USER FORM SUBMISSION
// ===============================================================

/**
 * Handles submission from the index.html form. Performs validation, saves data, uploads file, sends email.
 * Modified to support multiple files.
 * @param {object} formData Form data (type, topic, details, location).
 * @param {Array|object|null} fileObjects Array of file objects or single file object or null.
 * @returns {object} {success: boolean, message: string, ticketId?: string, accessKey?: string}
 */
function submitForm(formData, fileObjects) {
  // --- START: LockService Implementation ---
  const lock = LockService.getScriptLock();
  try {
    // Wait up to 30 seconds for the lock to become available.
    lock.waitLock(30000); // 30 seconds
    Logger.log("Lock acquired for submitForm.");

    // --- Original function logic starts here ---
    try {
      // --- Server-Side Validation ---
      if (!formData || !formData.type || !formData.topic || !formData.details || !formData.location) {
          return { success: false, message: 'Missing required form fields.' };
      }
      
      // ===== NEW: FORM CONFIG VALIDATION =====
      const config = getFormConfig();
      if (formData.type === 'คำร้องเรียน' && !config.enableComplaint) {
          return { success: false, message: 'ระบบปิดรับคำร้องเรียนชั่วคราว (Complaints are temporarily disabled).' };
      }
      if (formData.type === 'ข้อเสนอแนะ' && !config.enableSuggestion) {
          return { success: false, message: 'ระบบปิดรับข้อเสนอแนะชั่วคราว (Suggestions are temporarily disabled).' };
      }
      if (formData.type === 'แจ้งปัญหา' && !config.enableReport) {
          return { success: false, message: 'ระบบปิดรับแจ้งปัญหาชั่วคราว (Issue reporting is temporarily disabled).' };
      }
      // ======================================

      // Check all relevant fields for profanity
      if (containsProfanity(formData.topic) || containsProfanity(formData.details) || containsProfanity(formData.location)) {
          const userLocale = Session.getActiveUserLocale().split('_')[0] || 'th';
          const translations = getTranslations(userLocale);
          const profanityErrorMessage = translations.profanityWarning || 'Please use polite language (Server Validation).';
          return { success: false, message: profanityErrorMessage };
      }

      // Normalize fileObjects to an array or empty array
      let filesToUpload = [];
      if (fileObjects) {
          if (Array.isArray(fileObjects)) {
              filesToUpload = fileObjects;
          } else {
              filesToUpload = [fileObjects];
          }
      }

      // File Validation & Upload Check
      for (const fileObj of filesToUpload) {
          if (fileObj && fileObj.base64) {
              const approxSizeInBytes = fileObj.base64.length * 0.75;
              if (approxSizeInBytes > 5 * 1024 * 1024) { // 5MB Limit per file
                  const translations = getTranslations('th');
                  const fileSizeErrorMessage = translations.fileSizeError || 'File size must not exceed 5MB (Server Validation).';
                  return { success: false, message: fileSizeErrorMessage };
              }
          }
      }
      
      if (filesToUpload.length > 0 && (!uploadFolderId || uploadFolderId === "YOUR_FOLDER_ID_HERE")) {
           Logger.log("File upload skipped: uploadFolderId not configured.");
           filesToUpload = []; // Clear files to prevent upload attempts
      }
      // --- End Validation ---

      // Get Sheet and Headers (within the lock)
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let dataSheet = ss.getSheetByName(dataSheetName);
      
      // Default headers
      let headers = ["เลขที่", "วันที่", "ประเภท", "หัวข้อ", "รายละเอียด", "สถานที่", "สถานะ", "ไฟล์แนบ", "Admin Comments", "Access Key", "Public Comment", "Assigned To"];

      if (!dataSheet) {
        dataSheet = ss.insertSheet(dataSheetName);
        dataSheet.appendRow(headers);
        Logger.log(`Created data sheet: ${dataSheetName}`);
      } else if (dataSheet.getLastRow() < 1) {
        dataSheet.appendRow(headers);
        Logger.log(`Added headers to empty data sheet: ${dataSheetName}`);
      } else {
        // Check for missing headers and add them if needed
        const lastCol = dataSheet.getLastColumn();
        headers = dataSheet.getRange(1, 1, 1, lastCol).getValues()[0];
        
        const checkAndAddHeader = (headerName) => {
            if (!headers.includes(headerName)) {
                const newCol = headers.length + 1;
                dataSheet.getRange(1, newCol).setValue(headerName);
                headers.push(headerName);
                Logger.log(`Added missing '${headerName}' header.`);
            }
        };

        checkAndAddHeader("Admin Comments");
        checkAndAddHeader("Access Key");
        checkAndAddHeader("Public Comment");
        checkAndAddHeader("Assigned To");
        // NEW: Check for Phone header
        checkAndAddHeader("Phone"); 
        checkAndAddHeader("Employee ID"); 
      }

      // Generate Ticket ID (within the lock)
      const timestamp = new Date();
      const currentPrefix = 'REQ-' + Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "yyMM"); 

      // --- CRITICAL PART within lock ---
      const lastRow = dataSheet.getLastRow(); 
      
      let newCounter = 1; 

      if (lastRow > 1) {
          const ticketIdColIndex = headers.indexOf("เลขที่"); 
          
          if (ticketIdColIndex !== -1) {
              const lastTicketId = dataSheet.getRange(lastRow, ticketIdColIndex + 1).getValue(); 
              
              const match = String(lastTicketId).match(/^REQ-(\d{4})-?(\d{3})$/);

              if (match) {
                  const lastPrefix = "REQ-" + match[1]; 
                  const lastCounter = parseInt(match[2], 10); 
                  
                  if (currentPrefix === lastPrefix) {
                      newCounter = lastCounter + 1;
                  } else {
                      newCounter = 1; 
                  }
              }
          }
      }

      const formattedCounter = ('00' + newCounter).slice(-3); 
      const ticketId = currentPrefix + formattedCounter; 
      
      const accessKey = Utilities.getUuid().substring(0, 8);

      // Upload Files (Loop)
      let fileUrls = [];
      if (filesToUpload.length > 0) {
          try {
              const mainFolder = DriveApp.getFolderById(uploadFolderId);
              let targetFolder;
              const subFolders = mainFolder.getFoldersByName(ticketId); 
              if (subFolders.hasNext()) {
                  targetFolder = subFolders.next();
                  Logger.log(`Using existing subfolder for ${ticketId}`);
              } else {
                  targetFolder = mainFolder.createFolder(ticketId); 
                  Logger.log(`Created new subfolder for ${ticketId}`);
              }

              for (const fileObj of filesToUpload) {
                  if (fileObj.base64) {
                      const decoded = Utilities.base64Decode(fileObj.base64, Utilities.Charset.UTF_8);
                      const blob = Utilities.newBlob(decoded, fileObj.mimeType, fileObj.fileName);
                      const uniqueFileName = `${Date.now()}_${fileObj.fileName}`;
                      const newFile = targetFolder.createFile(blob.setName(uniqueFileName));
                      fileUrls.push(newFile.getUrl());
                      Logger.log(`File uploaded: ${newFile.getUrl()}`);
                  }
              }
          } catch (uploadError) {
              Logger.log(`File upload failed for ${ticketId}: ${uploadError}. Proceeding with partial or no files.`);
          }
      }
      
      // Join URLs with comma for storage
      const fileUrlString = fileUrls.join(',');

      // Prepare and Append Data Row (within the lock)
      const headerMap = {};
      headers.forEach((h, i) => { if(h) headerMap[String(h).trim()] = i; });
      const newRowData = Array(headers.length).fill('');
      
      // Helper to set data if header exists
      const setVal = (key, val) => { if (headerMap[key] !== undefined) newRowData[headerMap[key]] = val; };

      setVal("เลขที่", ticketId);
      setVal("วันที่", timestamp);
      setVal("ประเภท", formData.type);
      setVal("หัวข้อ", formData.topic);
      setVal("รายละเอียด", formData.details);
      setVal("สถานที่", formData.location);
      setVal("สถานะ", 'ยังไม่ดำเนินการ');
      setVal("ไฟล์แนบ", fileUrlString);
      setVal("Admin Comments", '');
      setVal("Access Key", "'" + accessKey); // Force string
      setVal("Public Comment", '');
      setVal("Assigned To", '');
      // NEW: Set Phone value
      // Use 'Phone' or 'เบอร์โทรศัพท์' depending on what is in the sheet. 
      // Code above adds "Phone" if missing, so we use "Phone". 
      // If user manually added "Phone", this works.
      setVal("Phone", "'" + (formData.phone || '')); // Force string to keep leading zero
      setVal("Employee ID", "'" + (formData.employeeId || '')); // Add this

      // --- CRITICAL PART within lock ---
      dataSheet.appendRow(newRowData);
      Logger.log(`Data appended for Ticket ID: ${ticketId}`);
      // --- END CRITICAL PART ---

      // Send Email Notification
      const recipientEmail = PropertiesService.getScriptProperties().getProperty('RECIPIENT_EMAIL');
      
      if (recipientEmail && recipientEmail.includes('@')) {
          try {
              // Generate HTML links for all files
              let fileHtml = '';
              if (fileUrls.length > 0) {
                  fileHtml = '<p><b>ไฟล์แนบ:</b><br>';
                  fileUrls.forEach((url, index) => {
                      fileHtml += `<a href="${url}" target="_blank" rel="noopener noreferrer">คลิกเพื่อดูไฟล์ที่ ${index + 1}</a><br>`;
                  });
                  fileHtml += '</p>';
              }
              
              const adminDashboardUrl = ScriptApp.getService().getUrl() + '?page=admin';
              
              const formattedTimestamp = timestamp.toLocaleString('th-TH', { timeZone: Session.getScriptTimeZone(), dateStyle: 'full', timeStyle: 'medium'});
              const detailsHtml = formData.details.replace(/\n/g, '<br>');
              
              // Add phone to email body if exists
              const phoneHtml = formData.phone ? `<tr><td style="padding: 8px 0; font-size: 16px; color: #555555; font-weight: bold; width: 100px;">เบอร์โทร:</td><td style="padding: 8px 0; font-size: 16px; color: #111111;">${formData.phone}</td></tr>` : '';
              
              // Add Employee ID HTML
              const empIdHtml = formData.employeeId ? `<tr><td style="padding: 8px 0; font-size: 16px; color: #555555; font-weight: bold; width: 100px;">รหัสพนักงาน:</td><td style="padding: 8px 0; font-size: 16px; color: #111111;">${formData.employeeId}</td></tr>` : '';

              const htmlBody = `
                <html lang="th">
                <head>
                  <meta charset="UTF-8">
                  <style>
                    body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, 'Noto Sans', sans-serif; line-height: 1.6; }
                    .container { width: 90%; max-width: 600px; margin: 0 auto; border-radius: 8px; overflow: hidden; }
                    .header { background-color: #4f46e5; padding: 20px; text-align: center; }
                    .header h1 { color: #ffffff; margin: 0; font-size: 24px; }
                    .content { background-color: #ffffff; padding: 30px; }
                    .content-table { width: 100%; border-collapse: collapse; }
                    .content-table td { padding: 8px 0; font-size: 16px; }
                    .content-table td:first-child { color: #555555; font-weight: bold; width: 100px; }
                    .details-box { background-color: #f9fafb; border: 1px solid #e5e7eb; border-radius: 8px; padding: 16px; margin-top: 16px; }
                    .details-box p { margin: 0; white-space: pre-wrap; }
                    .footer { padding: 30px; text-align: center; }
                    .button { display: inline-block; background-color: #4f46e5; color: #ffffff; text-decoration: none; padding: 12px 24px; border-radius: 6px; font-weight: bold; }
                  </style>
                </head>
                <body style="background-color: #f4f4f4; padding: 20px; margin: 0; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, 'Noto Sans', sans-serif;">
                  <table class="container" role="presentation" border="0" cellpadding="0" cellspacing="0" style="width: 90%; max-width: 600px; margin: 0 auto; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 12px rgba(0,0,0,0.05);">
                    <tr>
                      <td class="header" style="background-color: #4f46e5; padding: 24px; text-align: center;">
                        <h1 style="color: #ffffff; margin: 0; font-size: 24px; font-weight: 600;">มีเรื่องใหม่เข้ามาในระบบ</h1>
                      </td>
                    </tr>
                    <tr>
                      <td class="content" style="background-color: #ffffff; padding: 30px;">
                        <table class="content-table" role="presentation" border="0" cellpadding="0" cellspacing="0" style="width: 100%; border-collapse: collapse;">
                          <tr><td style="padding: 8px 0; font-size: 16px; color: #555555; font-weight: bold; width: 100px;">เลขที่:</td><td style="padding: 8px 0; font-size: 16px; color: #111111;">${ticketId}</td></tr>
                          <tr><td style="padding: 8px 0; font-size: 16px; color: #555555; font-weight: bold; width: 100px;">ประเภท:</td><td style="padding: 8px 0; font-size: 16px; color: #111111;">${formData.type}</td></tr>
                          <tr><td style="padding: 8px 0; font-size: 16px; color: #555555; font-weight: bold; width: 100px;">หัวข้อ:</td><td style="padding: 8px 0; font-size: 16px; color: #111111;">${formData.topic}</td></tr>
                          ${empIdHtml}
                          ${phoneHtml}
                          <tr><td style="padding: 8px 0; font-size: 16px; color: #555555; font-weight: bold; width: 100px;">สถานที่:</td><td style="padding: 8px 0; font-size: 16px; color: #111111;">${formData.location}</td></tr>
                          <tr><td style="padding: 8px 0; font-size: 16px; color: #555555; font-weight: bold; width: 100px;">เวลา:</td><td style="padding: 8px 0; font-size: 16px; color: #111111;">${formattedTimestamp}</td></tr>
                        </table>
                        <div class="details-box" style="background-color: #f9fafb; border: 1px solid #e5e7eb; border-radius: 8px; padding: 16px; margin-top: 20px;">
                          <h4 style="margin: 0 0 10px 0; color: #555555; font-size: 16px; font-weight: bold;">รายละเอียด:</h4>
                          <p style="margin: 0; white-space: pre-wrap; color: #111111; font-size: 16px;">${detailsHtml}</p>
                        </div>
                        ${fileHtml}
                      </td>
                    </tr>
                    <tr>
                      <td class="footer" style="background-color: #ffffff; padding: 30px; padding-top: 10px; text-align: center; border-top: 1px solid #e5e7eb;">
                        <a href="${adminDashboardUrl}" class="button" style="display: inline-block; background-color: #4f46e5; color: #ffffff; text-decoration: none; padding: 12px 24px; border-radius: 6px; font-weight: bold; font-size: 16px;">
                          เปิด Dashboard เพื่อจัดการ
                        </a>
                      </td>
                    </tr>
                  </table>
                </body>
                </html>
              `;

              MailApp.sendEmail({
                  to: recipientEmail, 
                  subject: `[เรื่องใหม่ ${ticketId}] ${formData.type}: ${formData.topic}`,
                  htmlBody: htmlBody 
              });
              Logger.log(`Email notification sent for ${ticketId} to ${recipientEmail}`);
          } catch (emailError) {
              Logger.log(`Failed to send email notification for ${ticketId}: ${emailError}`);
          }
      } else {
           Logger.log(`Email notification skipped: RECIPIENT_EMAIL is not configured in PropertiesService or is invalid.`);
      }

      // Cleared headers cache due to new submission.
      const cache = CacheService.getScriptCache();
      cache.remove('all_data_headers_v2');
      Logger.log('Cleared headers cache due to new submission.');

      const successTranslations = getTranslations(Session.getActiveUserLocale().split('_')[0] || 'th');
      const successMessageText = successTranslations.successMessage || 'ข้อมูลของคุณถูกส่งเรียบร้อยแล้ว';

      return { success: true, message: successMessageText, ticketId: ticketId, accessKey: accessKey };

    } catch (innerError) {
      Logger.log(`Critical error inside submitForm lock: ${innerError} Stack: ${innerError.stack}`);
      return { success: false, message: 'เกิดข้อผิดพลาดร้ายแรงขณะบันทึกข้อมูล กรุณาลองใหม่อีกครั้ง' };
    }
    // --- Original function logic ends here ---

  } catch (lockError) {
    Logger.log(`Could not obtain lock within 30 seconds for submitForm: ${lockError}`);
    return { success: false, message: 'เซิร์ฟเวอร์กำลังประมวลผลคำขออื่นอยู่ กรุณาลองใหม่อีกครั้งในภายหลัง' };
  } finally {
    if (lock) {
      lock.releaseLock();
      Logger.log("Lock released for submitForm.");
    }
  }
}

// ===============================================================
//          ADMIN DATA FUNCTIONS (WITH CACHING)
// ===============================================================

// ===== START: NEW (checkLogin function RESTORED) =====
/**
 * Checks login credentials, returns role and session token if valid. Includes detailed logging.
 * @param {object} credentials {username, password}.
 * @returns {object|null} {role, token} or null if login fails.
 */
function checkLogin(credentials) {
  // --- START: Detailed Logging ---
  Logger.log(`checkLogin called for user: ${credentials ? credentials.username : 'undefined'}`);
  try {
    // Input validation
    if (!credentials || !credentials.username || !credentials.password) {
        Logger.log("Login attempt with missing credentials.");
        return null;
    }
    // REVIEW: Changed to trim AND convert to lower case immediately.
    const inputUsername = String(credentials.username).trim().toLowerCase();
    Logger.log(`Attempting login for trimmed username: '${inputUsername}'`);


    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const userSheet = ss.getSheetByName(userSheetName);
    const expectedHeader = ['Username', 'PasswordHash', 'Salt', 'Role', 'Status', 'Team', 'AllowedTypes'];

    if (!userSheet) {
        Logger.log(`Login failed: User sheet '${userSheetName}' not found.`);
        // Consider creating the sheet or returning a specific error if needed
        return null; // Stop execution if sheet not found
    }

    // Ensure header row exists and matches expected format
    if (userSheet.getLastRow() < 1) { // Sheet is completely empty
       userSheet.appendRow(expectedHeader);
       Logger.log(`Created header for empty sheet: ${userSheetName}`);
       return null; // No users exist
    }

    const header = userSheet.getRange(1, 1, 1, 7).getValues()[0];
    if (JSON.stringify(header) !== JSON.stringify(expectedHeader)) {
        if (userSheet.getLastRow() === 1) {
             userSheet.getRange(1, 1, 1, 7).setValues([expectedHeader]);
             Logger.log(`Fixed header for sheet: ${userSheetName}`);
             return null; // No users exist yet
        } else {
             Logger.log(`CRITICAL: Sheet '${userSheetName}' has incorrect headers. Expected: ${expectedHeader}, Found: ${header}`);
             // Return null or throw error depending on desired behavior
             return null; // Prevent login if header is wrong with existing data
        }
    }

    if (userSheet.getLastRow() < 2) {
      Logger.log(`Login attempt failed: No users found in sheet '${userSheetName}'.`);
      return null; // Only header row, no users to check
    }

    // Get all user data (Username, Hash, Salt, Role, Status, Team, AllowedTypes)
    const users = userSheet.getRange(2, 1, userSheet.getLastRow() - 1, 7).getValues();
    Logger.log(`Found ${users.length} user record(s) to check.`);

    let userFound = false; // Flag to check if user was found at all

    for (let i = 0; i < users.length; i++) {
      const userData = users[i];
      // Check if row has enough columns and a username
      if (userData.length < 6 || !userData[0]) {
          Logger.log(`Skipping row ${i+2}: incomplete data or missing username.`);
          continue;
      }

      const usernameFromSheet = String(userData[0]).trim(); // Trim sheet username
      const hashFromSheet = userData[1];
      const saltFromSheet = userData[2];
      const roleFromSheet = userData[3];
      const statusFromSheet = userData[4]; // Status
      const teamFromSheet = userData[5]; // Team
      const allowedTypesFromSheet = userData[6]; // Allowed Types (Col G)

      // Check User Status FIRST
      if (usernameFromSheet.toLowerCase() === inputUsername) {
         if (statusFromSheet && String(statusFromSheet).trim().toLowerCase() === 'inactive') {
             Logger.log(`Login denied for '${inputUsername}': User status is Inactive.`);
             // MODIFIED: Return specific error object instead of null
             return { error: 'ACCOUNT_INACTIVE' }; 
         }
      }

      // REVIEW: Changed comparison to be case-insensitive.
      // We now compare the sheet's lowercase username to our already-lowercase inputUsername.
      if (usernameFromSheet.toLowerCase() === inputUsername) {
        userFound = true; // Mark user as found in the sheet
        Logger.log(`Username match for '${inputUsername}' at row ${i+2}. Checking password...`);

        // Ensure salt and hash are not empty before proceeding
        if (!saltFromSheet || !hashFromSheet) {
          Logger.log(`WARNING: User ${inputUsername} at row ${i+2} has missing salt or hash. Cannot verify password for this entry.`);
          continue; // Skip password check for this user entry, might be another entry later
        }

        try {
            // Compute hash from input password + sheet salt
            const combined = credentials.password + saltFromSheet;
            const hashFromInput = Utilities.base64Encode(Utilities.computeDigest(
                // REVIEW: Fixed typo from SHA_26 to SHA_26. This was the cause of the login failure.
                Utilities.DigestAlgorithm.SHA_256,
                combined,
                Utilities.Charset.UTF_8
            ));

            // Compare computed hash with stored hash
            if (hashFromInput === hashFromSheet) {
              Logger.log(`Password match successful for '${inputUsername}'.`);
              // Password is correct, create session token
              const token = Utilities.getUuid();
              const userCache = CacheService.getUserCache();

              // --- NEW MULTI-SESSION LOGIC ---
              // Store user-specific data in cache with a UNIQUE key based on the token
              const validRole = (roleFromSheet && ['Admin', 'User'].includes(roleFromSheet)) ? roleFromSheet : 'User';
              const sessionData = {
                  role: validRole,
                  username: usernameFromSheet,
                  allowedTypes: allowedTypesFromSheet || ''
              };
              
              // Store session data with token-specific key
              const sessionKey = 'admin_session_' + token;
              userCache.put(sessionKey, JSON.stringify(sessionData), 3600); // 1 hour expiry

              Logger.log(`Login successful. Session key: ${sessionKey}`);

              // Return role, token, AND username to the client
              return { role: validRole, token: token, username: usernameFromSheet };
            } else {
               Logger.log(`Password mismatch for user: '${inputUsername}' at row ${i+2}.`);
            }
        } catch (hashError) {
             Logger.log(`Error during password hash comparison for user ${inputUsername} at row ${i+2}: ${hashError}`);
             continue;
        }
      } 
    } 

    // After checking all users
    if (userFound) {
      // If user was found but password never matched any entry for that username
      Logger.log(`Login failed: User '${inputUsername}' was found, but password did not match.`);
    } else {
      // If username was never found in the loop
      Logger.log(`Login failed: User '${inputUsername}' not found in the sheet.`);
    }
    return null; // User not found or password incorrect after checking all possibilities

  } catch (error) {
    // Log any critical error during the process (e.g., sheet access issues)
    Logger.log(`CRITICAL error in checkLogin function: ${error} Stack: ${error.stack}`);
    // Don't expose internal error details to client, just return null
    return null; // Return null on any unexpected error to prevent login
  }
}
// ===== END: NEW (checkLogin function RESTORED) =====


/**
 * Retrieves data for the admin dashboard.
 * @param {string} token The session token.
 * @param {object} options Contains pagination, filter settings.
 * @returns {object} Dashboard data or {error: string}.
 */
function getPaginatedData(token, options) {
  // 1. Auth Check
  if (!isUserAuthenticated(token)) {
      return { error: "Authentication failed. Please log in again." };
  }
  
  // --- START: FORCE REFRESH CHECK ---
  const session = getSessionData(token);
  const username = session ? session.username : null;
  
  if (username) {
      const scriptCache = CacheService.getScriptCache();
      const refreshFlag = scriptCache.get('force_refresh_' + username.toLowerCase());
      
      if (refreshFlag) {
          try {
              const ss = SpreadsheetApp.getActiveSpreadsheet();
              const userSheet = ss.getSheetByName(userSheetName);
              if (userSheet) {
                  // Find user row to get updated permissions
                  // We scan to find the user.
                  // Optimization: We could use TextFinder if needed, but loop is fine for admin panel scale.
                  const data = userSheet.getDataRange().getValues(); // Read all
                  // Assume Row 1 is header
                  for (let i = 1; i < data.length; i++) {
                      if (String(data[i][0]).toLowerCase() === username.toLowerCase()) {
                          const newAllowedTypes = data[i][6] || ''; // Col G is index 6
                          
                          // Update session object
                          session.allowedTypes = String(newAllowedTypes);
                          const userCache = CacheService.getUserCache();
                          userCache.put('admin_session_' + token, JSON.stringify(session), 3600);
                          
                          scriptCache.remove('force_refresh_' + username.toLowerCase()); // Clear flag
                          Logger.log(`Refreshed session permissions for ${username}: ${newAllowedTypes}`);
                          break;
                      }
                  }
              }
          } catch (e) {
              Logger.log("Error refreshing user session: " + e);
          }
      }
  }
  // --- END: FORCE REFRESH CHECK ---
  
  const opt = options || {};
  const cache = CacheService.getScriptCache();
  const headersCacheKey = 'all_data_headers_v2'; // Cache ONLY headers

  let headers;
  let dataRows;

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dataSheet = ss.getSheetByName(dataSheetName);

    // --- PART 1: HEADERS (Use Cache) ---
    const cachedHeaders = cache.get(headersCacheKey);
    if (cachedHeaders && !opt.forceRefresh) {
       headers = JSON.parse(cachedHeaders);
    } else {
       // If no cache, fetch headers from sheet
       if (!dataSheet || dataSheet.getLastRow() < 1) {
          // Fallback headers if sheet is empty
          headers = ["เลขที่", "วันที่", "ประเภท", "หัวข้อ", "รายละเอียด", "สถานที่", "สถานะ", "ไฟล์แนบ", "Admin Comments", "Access Key", "Public Comment", "Assigned To"];
       } else {
          const lastCol = dataSheet.getLastColumn();
          headers = dataSheet.getRange(1, 1, 1, lastCol).getValues()[0].filter(h => h);
       }
       // Save Headers to Cache (Safe, small size)
       cache.put(headersCacheKey, JSON.stringify(headers), 21600); // Cache for 6 hours
    }

    // --- PART 2: DATA ROWS (Fetch Fresh - NO CACHE to avoid 100KB limit) ---
    if (!dataSheet || dataSheet.getLastRow() < 2) {
       dataRows = [];
    } else {
       // Fetch all data rows fresh from the sheet
       dataRows = dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, headers.length).getValues();
    }

  } catch (e) {
    Logger.log(`Error reading Sheet: ${e}`);
    return { error: "Server error reading Sheet: " + e.message };
  }

  // --- PART 3: PROCESSING (Filtering & Pagination) ---
  const itemsPerPage = 30;
  const page = parseInt(opt.page, 10) || 1;
  const filters = opt.filters || {};
  const filterType = filters.type || 'ทั้งหมด';
  const filterStatus = filters.status || 'ทั้งหมด';
  const filterAssignee = (filters.assignee !== undefined) ? filters.assignee : 'ทั้งหมด';
  const searchTerm = (opt.searchTerm || '').toLowerCase();

  // --- NEW: Apply "Allowed Types" Filter ---
  // Use allowedTypes from session object (retrieved earlier)
  const allowedTypesRaw = session ? session.allowedTypes : '';
  const typeIndex = headers.indexOf("ประเภท");

  if (allowedTypesRaw && typeIndex > -1) {
      const allowedList = allowedTypesRaw.split(',').map(t => t.trim().toLowerCase());
      // Only filter if 'all' is NOT in the list (case-insensitive check)
      if (!allowedList.includes('all') && !allowedList.includes('ทั้งหมด')) {
          dataRows = dataRows.filter(row => {
              const rowType = String(row[typeIndex] || '').toLowerCase();
              return allowedList.includes(rowType);
          });
      }
  }


  try {
    const initialStatusCounts = {'ยังไม่ดำเนินการ': 0, 'กำลังดำเนินการ': 0, 'เสร็จสิ้น': 0, 'ยกเลิก': 0};

    // Prepare variables for Phase 1 (SLA & Assignee)
    const slaSummary = { critical: 0, warning: 0 };
    const assigneeStatsMap = {}; 
    const now = new Date();
    
    // ===== START: NEW LOGIC (Unassigned Queue) =====
    const unassignedQueue = [];
    // ===============================================

    if (!dataRows || dataRows.length === 0) {
       return { 
        summary: { total: 0, complaints: 0, suggestions: 0, reportIssues: 0, statusCounts: {...initialStatusCounts} },
        filteredStatusCounts: {...initialStatusCounts},
        slaSummary: slaSummary, // Empty
        assigneePerformance: [], // Empty
        unassignedQueue: [], // Empty
        records: [], 
        pagination: { currentPage: 1, totalPages: 0, totalFilteredItems: 0, itemsPerPage: itemsPerPage },
        headers: headers 
      };
    }

    // Map Indices
    // typeIndex already defined above
    const statusIndex = headers.indexOf("สถานะ");
    const dateIndex = headers.indexOf("วันที่");
    const assignIndex = headers.indexOf("Assigned To");
    const topicIndex = headers.indexOf("หัวข้อ");
    const ticketIdIndex = headers.indexOf("เลขที่");

    let complaints = 0;
    let suggestions = 0;
    let reportIssues = 0;
    const grandTotalStatusCounts = {...initialStatusCounts};

    dataRows.forEach(row => {
      if (typeIndex > -1 && row.length > typeIndex) {
        if (row[typeIndex] === 'คำร้องเรียน') complaints++;
        else if (row[typeIndex] === 'ข้อเสนอแนะ') suggestions++;
        else if (row[typeIndex] === 'แจ้งปัญหา') reportIssues++;
      }
      
      const status = (statusIndex > -1 && row.length > statusIndex) ? row[statusIndex] : '';
      if (grandTotalStatusCounts.hasOwnProperty(status)) {
        grandTotalStatusCounts[status]++;
      }

      // --- PHASE 1 LOGIC START (SLA & Assignee Stats) ---
      const rawAssignee = (assignIndex > -1 && row.length > assignIndex) ? row[assignIndex] : '';
      const assignee = rawAssignee ? String(rawAssignee).trim() : 'Unassigned';
      
      // Initialize Assignee in Map if new
      if (!assigneeStatsMap[assignee]) {
        assigneeStatsMap[assignee] = { name: assignee, active: 0, completed: 0, cancelled: 0, total: 0 };
      }
      
      assigneeStatsMap[assignee].total++;
      
      // Calculate SLA for this ticket
      let isOverdue = false;
      if (dateIndex > -1 && row.length > dateIndex) {
          let ticketDate = row[dateIndex];
          // Ensure it's a date object (it should be from Sheet, but safety first)
          if (!(ticketDate instanceof Date) && typeof ticketDate === 'string') {
            // Try parsing if string (rare case in hybrid mode but possible)
              const parts = ticketDate.split(/[/\s:]/);
              if(parts.length >= 3) {
                  ticketDate = new Date(parts[2], parts[1]-1, parts[0]); // Simple fallback
              } else {
                  ticketDate = new Date(ticketDate);
              }
          }

          if (ticketDate instanceof Date && !isNaN(ticketDate.getTime())) {
              const diffMs = now - ticketDate;
              const diffHours = diffMs / (1000 * 60 * 60);
              
              // Logic matching frontend
              if (status === 'ยังไม่ดำเนินการ' && diffHours > 24) {
                  slaSummary.critical++;
                  isOverdue = true;
              } else if (status === 'กำลังดำเนินการ' && diffHours > 72) {
                  slaSummary.warning++;
                  isOverdue = true;
              }
          }
      }

      // Update Assignee Stats
      if (status === 'ยังไม่ดำเนินการ' || status === 'กำลังดำเนินการ') {
          assigneeStatsMap[assignee].active++;
      } else if (status === 'เสร็จสิ้น') {
          assigneeStatsMap[assignee].completed++;
      } else if (status === 'ยกเลิก') { // <--- NEW LOGIC: Count Cancelled
          assigneeStatsMap[assignee].cancelled++;
      }
      
      // --- PHASE 1 LOGIC END ---

      // ===== START: NEW LOGIC (Populate Unassigned Queue) =====
      // Condition: Status is Active (New/InProgress) AND Assignee is 'Unassigned'
      if ((status === 'ยังไม่ดำเนินการ' || status === 'กำลังดำเนินการ') && assignee === 'Unassigned') {
          // Limit queue size to top 20 to keep payload light
          if (unassignedQueue.length < 20) {
              // Create a display-friendly copy of the row (format date)
              // Fix: Convert ALL Date objects in the row to strings to prevent serialization errors
              let displayRow = row.map(cell => {
                  if (cell instanceof Date) {
                      return Utilities.formatDate(cell, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
                  }
                  return cell;
              });
              unassignedQueue.push(displayRow);
          }
      }
      // ===== END: NEW LOGIC =====
    });

    const grandTotalSummary = {
      total: dataRows.length,
      complaints: complaints,
      suggestions: suggestions,
      reportIssues: reportIssues,
      statusCounts: grandTotalStatusCounts
    };
    
    // Extract Recent Activity (Raw top 5 before filtering)
    // Note: dataRows is usually oldest -> newest (top -> bottom), so reverse for recent
    const recentRaw = [...dataRows].reverse().slice(0, 5);
    const recentRecords = recentRaw.map(row => {
      const displayRow = [...row];
      // *** START: FIX FOR DATE OBJECT SERIALIZATION ERROR ***
      // Loop through all items in the row and convert any Date objects to strings
      return displayRow.map(cell => {
          if (cell instanceof Date) {
              return Utilities.formatDate(cell, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
          }
          return cell;
      });
      // *** END: FIX ***
    });

    // Convert Assignee Map to Sorted Array
    const assigneePerformance = Object.values(assigneeStatsMap).sort((a, b) => {
        if (b.active !== a.active) return b.active - a.active;
        return b.total - a.total;
    });

    // --- FILTERING ---
    let filteredData = dataRows;

    // 1. Type Filter
    if (typeIndex > -1 && filterType !== 'ทั้งหมด') {
      filteredData = filteredData.filter(row => row[typeIndex] === filterType);
    }
    
    // 2. Assignee Filter
    if (assignIndex > -1 && filterAssignee !== 'ทั้งหมด') {
      filteredData = filteredData.filter(row => {
        const assignee = row[assignIndex] || '';
        return assignee === filterAssignee;
      });
    }

    // 3. Date Filter
    const dateRangeKey = filters.dateRangeKey || 'all';
    let filterStartDate = null;
    let filterEndDate = null;

    if (dateRangeKey === 'custom' && filters.dateStart && filters.dateEnd) {
        try {
            filterStartDate = new Date(filters.dateStart + 'T00:00:00'); 
            let tempEnd = new Date(filters.dateEnd + 'T00:00:00');
            tempEnd.setDate(tempEnd.getDate() + 1); 
            filterEndDate = tempEnd;
        } catch(e) {}
    } else if (dateRangeKey !== 'all') {
        const getStartOfDay = (d) => { let x = new Date(d); x.setHours(0,0,0,0); return x; };
        switch (dateRangeKey) {
            case 'today':
                filterStartDate = getStartOfDay(now);
                filterEndDate = new Date(filterStartDate); filterEndDate.setDate(filterEndDate.getDate() + 1);
                break;
            case 'this_week':
                let day = now.getDay(); let diff = now.getDate() - day + (day == 0 ? -6 : 1);
                filterStartDate = getStartOfDay(new Date(now.setDate(diff)));
                filterEndDate = new Date(filterStartDate); filterEndDate.setDate(filterEndDate.getDate() + 7);
                break;
            case 'this_month':
                filterStartDate = new Date(now.getFullYear(), now.getMonth(), 1);
                filterEndDate = new Date(now.getFullYear(), now.getMonth() + 1, 1);
                break;
             case 'last_month':
                filterStartDate = new Date(now.getFullYear(), now.getMonth() - 1, 1);
                filterEndDate = new Date(now.getFullYear(), now.getMonth(), 1);
                break;
        }
    }
    
    if (filterStartDate && filterEndDate && dateIndex > -1) {
        filteredData = filteredData.filter(row => {
            const ticketDate = row[dateIndex];
            if (ticketDate instanceof Date) {
                return ticketDate >= filterStartDate && ticketDate < filterEndDate;
            }
            return false;
        });
    }

    // Calculate Filtered Status Counts (for cards)
    const filteredStatusCounts = {...initialStatusCounts};
    if (statusIndex > -1) {
      filteredData.forEach(row => {
          const status = row[statusIndex];
          if (filteredStatusCounts.hasOwnProperty(status)) {
            filteredStatusCounts[status]++;
          }
      });
    }

    // 4. Status Filter
    if (statusIndex > -1 && filterStatus !== 'ทั้งหมด') {
      filteredData = filteredData.filter(row => row[statusIndex] === filterStatus);
    }

    // 5. Search Term Filter
    if (searchTerm) {
      filteredData = filteredData.filter(row => {
        return row.some((cell, i) => {
          if (i === dateIndex && cell instanceof Date) {
             const formatted = Utilities.formatDate(cell, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
             return formatted.toLowerCase().includes(searchTerm);
          }
          return String(cell).toLowerCase().includes(searchTerm);
        });
      });
    }

    // Sort (Newest first)
    filteredData.reverse();

    // Pagination
    const totalFilteredItems = filteredData.length;
    const totalPages = Math.ceil(totalFilteredItems / itemsPerPage);
    const validPage = Math.max(1, Math.min(page, totalPages || 1));
    const startIndex = (validPage - 1) * itemsPerPage;
    const endIndex = startIndex + itemsPerPage;
    const paginatedData = filteredData.slice(startIndex, startIndex + itemsPerPage);

    // Format for Display
    const recordsForDisplay = paginatedData.map(row => {
      const displayRow = [...row];
      return displayRow.map(cell => {
          if (cell instanceof Date) {
              return Utilities.formatDate(cell, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
          }
          return cell;
      });
    });

    return {
      summary: grandTotalSummary,
      filteredStatusCounts: filteredStatusCounts,
      records: recordsForDisplay,
      recentRecords: recentRecords,
      slaSummary: slaSummary, 
      assigneePerformance: assigneePerformance, 
      unassignedQueue: unassignedQueue,
      pagination: {
        currentPage: validPage,
        totalPages: totalPages,
        totalFilteredItems: totalFilteredItems,
        itemsPerPage: itemsPerPage
      },
      headers: headers
    };

  } catch (error) {
    Logger.log(`Error processing data: ${error}`);
    return { error: "Server error processing data: " + error.message };
  }
}

// ===== START: NEW (Get Audit Log) =====
/**
 * Retrieves the last 200 entries from the AuditLog.
 * REQUIRES ADMIN ROLE.
 * @param {string} token The admin session token.
 * @returns {object} {success: boolean, data?: string[][], headers?: string[]} or {success: false, error: string}
 */
function getAuditLog(token) {
  // 1. Auth Check: Must be Admin
  if (!isUserAdmin(token)) {
    return { success: false, error: "Authentication failed. Requires Admin role." };
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName(auditLogSheetName);
    const expectedHeader = ["Timestamp", "Username", "Action", "Details"];

    // 2. Check sheet existence and headers
    if (!logSheet) {
      return { success: true, data: [], headers: expectedHeader };
    }
    if (logSheet.getLastRow() < 1) {
      return { success: true, data: [], headers: expectedHeader };
    }

    const headers = logSheet.getRange(1, 1, 1, logSheet.getLastColumn()).getValues()[0];
    
    if (logSheet.getLastRow() < 2) {
      return { success: true, data: [], headers: headers };
    }

    const dataRows = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, headers.length).getValues();

    const processedData = dataRows.map(row => {
      let formattedDate = row[0];
      if (formattedDate instanceof Date) {
        formattedDate = Utilities.formatDate(formattedDate, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
      }
      
      // Return a new array with the formatted date
      return [
        String(formattedDate), // Timestamp
        String(row[1] || ''),    // Username
        String(row[2] || ''),    // Action
        String(row[3] || '')     // Details
      ];
    });

    // 7. Reverse and slice to get the last 200
    processedData.reverse(); // Newest first
    const recentLogs = processedData.slice(0, 200); // Get only the top 200

    Logger.log(`getAuditLog: Returning ${recentLogs.length} most recent log entries.`);
    return { success: true, data: recentLogs, headers: headers };

  } catch (e) {
    Logger.log(`Error in getAuditLog: ${e}`);
    return { success: false, error: "Server error reading AuditLog: " + e.message };
  }
}
// ===== END: NEW (Get Audit Log) =====

// ===============================================================
//                        ADMIN SETTINGS FUNCTIONS (USERS)
// ===============================================================

/**
 * Retrieves a list of users (username, role, status, team, allowedTypes). Requires Admin role.
 * @param {string} token The admin session token.
 * @returns {object[]|object} Array of {username, role, status, team, allowedTypes} or {error: string}.
 */
function getUsers(token) {
  if (!isUserAdmin(token)) {
      return { error: "Authentication failed. Please log in again." };
  }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const userSheet = ss.getSheetByName(userSheetName);
    const defaultReturn = []; // Return empty array if sheet issues
    const expectedHeader = ['Username', 'PasswordHash', 'Salt', 'Role', 'Status', 'Team', 'AllowedTypes'];

    if (!userSheet || userSheet.getLastRow() < 1) {
        Logger.log(`User sheet '${userSheetName}' not found or completely empty.`);
        return defaultReturn;
    }

    // Check header presence and content
    const headerRange = userSheet.getRange(1, 1, 1, 7); // Check A1:G1
    const header = headerRange.getValues()[0];
    if (JSON.stringify(header) !== JSON.stringify(expectedHeader)) {
         Logger.log(`User sheet header mismatch. Expected: ${expectedHeader}, Found: ${header}. Attempting to fix or proceed.`);
         if (userSheet.getLastRow() === 1) {
             headerRange.setValues([expectedHeader]);
             Logger.log(`Header corrected for sheet: ${userSheetName}`);
             return defaultReturn; // No users yet
         } else {
             Logger.log(`CRITICAL WARNING: User sheet '${userSheetName}' has data but incorrect headers.`);
             // Throw error to prevent potential data corruption on add/delete/update
             throw new Error(`User sheet '${userSheetName}' has incorrect headers.`);
         }
    }

    if (userSheet.getLastRow() < 2) return defaultReturn; // Header exists, but no users

    // Read Username (col 1), Role (col 4), Status (col 5), Team (col 6), AllowedTypes (col 7)
    const usersData = userSheet.getRange(2, 1, userSheet.getLastRow() - 1, 7).getValues(); // Read A:G from row 2
    return usersData.map(row => ({
       username: row[0] || '', // Ensure username is a string
       // Validate Role, default to 'User'
       role: (row[3] && ['Admin', 'User'].includes(row[3])) ? row[3] : 'User',
       status: (row[4]) ? row[4] : 'Active', // Default status to Active if empty
       team: (row[5]) ? row[5] : '', // Default team to empty string
       allowedTypes: (row[6]) ? row[6] : '' // Default Allowed Types to empty
     })).filter(user => user.username); // Filter out rows with empty usernames

  } catch (e) {
    Logger.log(`Error in getUsers: ${e}`);
    return {error: `Error loading users: ${e.message}`}; // Return error object
  }
}

// ===== START: NEW FUNCTION (Assign To) =====
/**
 * Retrieves a list of all ACTIVE TEAMS for the "Assign To" dropdown.
 * Requires authenticated user (any role).
 * @param {string} token The user session token.
 * @returns {object[]|object} Array of strings {teams} or {error: string}.
 */
function getAssignableUsers(token) {
  // Use isUserAuthenticated (not Admin) so 'User' roles can also be assigned tasks if needed.
  if (!isUserAuthenticated(token)) { 
      return { error: "Authentication failed. Please log in again." };
  }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    // CHANGED: Source from 'Teams' sheet
    let sheet = ss.getSheetByName(teamSheetName);
    if (!sheet) return [];
    if (sheet.getLastRow() < 2) return [];

    const teams = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat();
    return [...new Set(teams.map(t => String(t).trim()).filter(t => t))].sort();

  } catch (e) {
    Logger.log(`Error in getAssignableUsers: ${e}`);
    return {error: `Error loading assignable teams: ${e.message}`}; // Return error object
  }
}
// ===== END: NEW FUNCTION (Assign To) =====

// ===== START: NEW FUNCTION (Get Unique Teams for Datalist) =====
/**
 * Retrieves a list of all unique teams from the Users sheet.
 * Requires authenticated user.
 * @param {string} token The user session token.
 * @returns {string[]} Sorted unique team names.
 */
function getAllUniqueTeams(token) {
  if (!isUserAuthenticated(token)) {
      return { error: "Authentication failed." };
  }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    // CHANGED: Source from 'Teams' sheet
    let sheet = ss.getSheetByName(teamSheetName);
    if (!sheet) return [];
    if (sheet.getLastRow() < 2) return [];

    const teams = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat();
    return [...new Set(teams.map(t => String(t).trim()).filter(t => t))].sort();
  } catch (e) {
    Logger.log("Error getting teams: " + e);
    return [];
  }
}
// ===== END: NEW FUNCTION (Get Unique Teams for Datalist) =====


/**
 * Adds a new user with hashed password, role, Active status, and Team. Requires Admin role.
 * @param {string} token The admin session token.
 * @param {object} user User details {username, password, role, team, allowedTypes}.
 * @returns {object[]|object} Updated user list or {error: string}.
 */
function addUser(token, user) {
  if (!isUserAdmin(token)) {
      return { error: "Authentication failed. Please log in again." };
  }
  try {
    // Validate input user object
    if (!user || !user.username || !user.password) {
        throw new Error("Username and password are required.");
    }
    const username = String(user.username).trim();
    if (!username) {
        throw new Error("Username cannot be empty.");
    }
    if (!user.role || (user.role !== 'Admin' && user.role !== 'User')) {
        throw new Error("Valid Role ('Admin' or 'User') is required.");
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let userSheet = ss.getSheetByName(userSheetName);
    const expectedHeader = ['Username', 'PasswordHash', 'Salt', 'Role', 'Status', 'Team', 'AllowedTypes'];

    // Setup sheet and header if needed
    if (!userSheet) {
      userSheet = ss.insertSheet(userSheetName);
      userSheet.appendRow(expectedHeader);
      Logger.log(`Created sheet: ${userSheetName}`);
    } else if (userSheet.getLastRow() < 1) {
      userSheet.appendRow(expectedHeader);
      Logger.log(`Added header to empty sheet: ${userSheetName}`);
    } else {
       // Validate existing header
       const header = userSheet.getRange(1, 1, 1, 7).getValues()[0];
       if (JSON.stringify(header) !== JSON.stringify(expectedHeader)) {
           throw new Error(`Sheet '${userSheetName}' has incorrect headers. Expected: ${expectedHeader}`);
       }
    }

    // Check if username already exists (case-insensitive)
    if (userSheet.getLastRow() > 1) {
      const usernames = userSheet.getRange(2, 1, userSheet.getLastRow() - 1, 1).getValues()
                                 .flat().map(u => String(u).trim().toLowerCase()); // Trim and lower case existing usernames
      if (usernames.includes(username.toLowerCase())) { // Compare with lower case input
        return { error: `Username '${username}' already exists.` };
      }
    }

    // Hash the password
    const salt = Utilities.base64Encode(Utilities.getUuid());
    const hash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, user.password + salt, Utilities.Charset.UTF_8));

    // Append new user data including role, 'Active' status, team, and allowedTypes
    const allowedTypes = user.allowedTypes || '';
    userSheet.appendRow([username, hash, salt, user.role, 'Active', user.team || '', allowedTypes]);
    Logger.log(`Added user: ${username} with role: ${user.role}, team: ${user.team}, allowedTypes: ${allowedTypes}`);

    // --- START: AUDIT LOG ---
    const session = getSessionData(token);
    _logAdminAction(session.username, "Add User", `Username: "${username}", Role: "${user.role}", Team: "${user.team}"`);
    // --- END: AUDIT LOG ---

    // Return the updated list of users (including roles)
    return getUsers(token); // Re-fetch the list to send back

  } catch(e) {
    Logger.log(`Error in addUser: ${e}`);
    return {error: `Error adding user: ${e.message}`}; // Return error object
  }
}

/**
 * Updates the Team of a specific user. Requires Admin role.
 * @param {string} token The admin session token.
 * @param {string} username The username whose team needs changing.
 * @param {string} newTeam The new team name.
 * @returns {object} {success: boolean, error?: string}
 */
function updateUserTeam(token, username, newTeam) {
    if (!isUserAdmin(token)) {
        return { success: false, error: "Authentication failed. Please log in again." };
    }
    try {
        // Validate inputs
        if (!username || typeof username !== 'string' || !username.trim()) {
            throw new Error("Valid username is required.");
        }
        const trimmedUsername = username.trim();
        const teamValue = newTeam ? String(newTeam).trim() : '';

        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const userSheet = ss.getSheetByName(userSheetName);
        if (!userSheet || userSheet.getLastRow() <= 1) { 
            throw new Error(`User sheet '${userSheetName}' not found or is empty.`);
        }

        // Check header
        const header = userSheet.getRange(1, 1, 1, 6).getValues()[0];
        const expectedHeader = ['Username', 'PasswordHash', 'Salt', 'Role', 'Status', 'Team'];
        // Note: Not strictly validating full header here to allow partial updates if sheet grows
        const usernameColumnNumber = 1; // A
        const teamColumnNumber = 6;   // F

        // Find the user row
        const lastRow = userSheet.getLastRow();
        const usernames = userSheet.getRange(2, usernameColumnNumber, lastRow - 1, 1)
                                   .getValues()
                                   .flat()
                                   .map(u => String(u).trim().toLowerCase()); 
        
        const arrayIndex = usernames.indexOf(trimmedUsername.toLowerCase());

        if (arrayIndex === -1) {
            throw new Error(`User '${trimmedUsername}' not found`);
        }

        const rowIndex = arrayIndex + 2;

        // Update Team
        userSheet.getRange(rowIndex, teamColumnNumber).setValue(teamValue);
        Logger.log(`Team updated for user: ${trimmedUsername} to ${teamValue} in row ${rowIndex}`);

        // --- START: AUDIT LOG ---
        const session = getSessionData(token);
        _logAdminAction(session.username, "Update User Team", `Set team for "${trimmedUsername}" to "${teamValue}"`);
        // --- END: AUDIT LOG ---

        return { success: true }; 

    } catch (error) {
        Logger.log(`Error in updateUserTeam: ${error}`);
        return { success: false, error: `Error updating team: ${error.message}` };
    }
}

/**
 * Updates the Allowed Types of a specific user. Requires Admin role.
 * @param {string} token The admin session token.
 * @param {string} username The username whose allowed types need changing.
 * @param {string} allowedTypes The new allowed types string (e.g., "Type1,Type2").
 * @returns {object} {success: boolean, error?: string}
 */
function updateUserAllowedTypes(token, username, allowedTypes) {
    if (!isUserAdmin(token)) {
        return { success: false, error: "Authentication failed. Please log in again." };
    }
    try {
        // Validate inputs
        if (!username || typeof username !== 'string' || !username.trim()) {
            throw new Error("Valid username is required.");
        }
        const trimmedUsername = username.trim();
        const typesValue = allowedTypes ? String(allowedTypes).trim() : '';

        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const userSheet = ss.getSheetByName(userSheetName);
        if (!userSheet || userSheet.getLastRow() <= 1) { 
            throw new Error(`User sheet '${userSheetName}' not found or is empty.`);
        }

        const usernameColumnNumber = 1; // A
        const allowedTypesColumnNumber = 7;   // G

        // Find the user row
        const lastRow = userSheet.getLastRow();
        const usernames = userSheet.getRange(2, usernameColumnNumber, lastRow - 1, 1)
                                   .getValues()
                                   .flat()
                                   .map(u => String(u).trim().toLowerCase()); 
        
        const arrayIndex = usernames.indexOf(trimmedUsername.toLowerCase());

        if (arrayIndex === -1) {
            throw new Error(`User '${trimmedUsername}' not found`);
        }

        const rowIndex = arrayIndex + 2;

        // Update Allowed Types
        userSheet.getRange(rowIndex, allowedTypesColumnNumber).setValue(typesValue);
        Logger.log(`Allowed Types updated for user: ${trimmedUsername} to ${typesValue} in row ${rowIndex}`);

        // NEW: Set Force Refresh Flag
        CacheService.getScriptCache().put('force_refresh_' + trimmedUsername.toLowerCase(), 'true', 21600); // 6 hours

        // --- START: AUDIT LOG ---
        const session = getSessionData(token);
        _logAdminAction(session.username, "Update User Permissions", `Set allowed types for "${trimmedUsername}" to "${typesValue}"`);
        // --- END: AUDIT LOG ---

        return { success: true }; 

    } catch (error) {
        Logger.log(`Error in updateUserAllowedTypes: ${error}`);
        return { success: false, error: `Error updating allowed types: ${error.message}` };
    }
}

/**
 * Updates the Status of a specific user (Active/Inactive). Requires Admin role.
 * @param {string} token The admin session token.
 * @param {string} username The username whose status needs changing.
 * @param {string} newStatus The new status ('Active' or 'Inactive').
 * @returns {object} {success: boolean, error?: string}
 */
function updateUserStatus(token, username, newStatus) {
    if (!isUserAdmin(token)) {
        return { success: false, error: "Authentication failed. Please log in again." };
    }
    try {
        // Validate inputs
        if (!username || typeof username !== 'string' || !username.trim()) {
            throw new Error("Valid username is required.");
        }
        if (!newStatus || (newStatus !== 'Active' && newStatus !== 'Inactive')) {
            throw new Error("Invalid status. Must be 'Active' or 'Inactive'.");
        }
        const trimmedUsername = username.trim();

        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const userSheet = ss.getSheetByName(userSheetName);
        if (!userSheet || userSheet.getLastRow() <= 1) { 
            throw new Error(`User sheet '${userSheetName}' not found or is empty.`);
        }

        // Check header
        const header = userSheet.getRange(1, 1, 1, 6).getValues()[0];
        const expectedHeader = ['Username', 'PasswordHash', 'Salt', 'Role', 'Status', 'Team'];
        // Not strictly validating full header here to allow growth
        const usernameColumnNumber = 1; // A
        const statusColumnNumber = 5;   // E

        // Find the user row
        const lastRow = userSheet.getLastRow();
        const usernames = userSheet.getRange(2, usernameColumnNumber, lastRow - 1, 1)
                                   .getValues()
                                   .flat()
                                   .map(u => String(u).trim().toLowerCase()); 
        
        const arrayIndex = usernames.indexOf(trimmedUsername.toLowerCase());

        if (arrayIndex === -1) {
            throw new Error(`User '${trimmedUsername}' not found`);
        }

        const rowIndex = arrayIndex + 2;

        // Update Status
        userSheet.getRange(rowIndex, statusColumnNumber).setValue(newStatus);
        Logger.log(`Status updated for user: ${trimmedUsername} to ${newStatus} in row ${rowIndex}`);

        // --- START: AUDIT LOG ---
        const session = getSessionData(token);
        _logAdminAction(session.username, "Update User Status", `Set status for "${trimmedUsername}" to "${newStatus}"`);
        // --- END: AUDIT LOG ---

        return { success: true }; 

    } catch (error) {
        Logger.log(`Error in updateUserStatus: ${error}`);
        return { success: false, error: `Error updating status: ${error.message}` };
    }
}

/**
 * Deletes a user. Requires Admin role.
 * @param {string} token The admin session token.
 * @param {string} username The username to delete.
 * @returns {object[]|object} Updated user list or {error: string}.
 */
function deleteUser(token, username) {
  if (!isUserAdmin(token)) {
      return { error: "Authentication failed. Please log in again." };
  }
  try {
     if (!username || typeof username !== 'string' || !username.trim()) {
        throw new Error("Valid username is required to delete.");
     }
     const trimmedUsername = username.trim();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const userSheet = ss.getSheetByName(userSheetName);
    // If sheet doesn't exist or has only header, user doesn't exist. Return current list.
    if (!userSheet || userSheet.getLastRow() <= 1) {
        Logger.log(`User sheet not found or empty during delete for ${trimmedUsername}.`);
        return getUsers(token); // Return potentially empty list or error from getUsers
    }

    // It's safer to read data and find row index first, then delete by index.
    const usernames = userSheet.getRange(2, 1, userSheet.getLastRow() - 1, 1).getValues().flat();
    let rowsToDelete = [];
    usernames.forEach((name, index) => {
        // Compare trimmed names (case-sensitive, adjust if needed)
        if (String(name).trim() === trimmedUsername) {
            rowsToDelete.push(index + 2); // +2 because index is 0-based and data starts row 2
        }
    });

    if (rowsToDelete.length === 0) {
       Logger.log(`User '${trimmedUsername}' not found for deletion.`);
       // Optionally return error: return { error: `User '${trimmedUsername}' not found.` };
    } else {
        // --- START: AUDIT LOG ---
        const session = getSessionData(token);
        const adminUsername = session.username;
        // --- END: AUDIT LOG ---
        
        // Delete rows from bottom up to avoid index shifting issues
        rowsToDelete.sort((a, b) => b - a).forEach(rowIndex => {
            userSheet.deleteRow(rowIndex);
            Logger.log(`Deleted user: ${trimmedUsername} from sheet row ${rowIndex}`);
            
            // --- START: AUDIT LOG ---
            _logAdminAction(adminUsername, "Delete User", `Username: "${trimmedUsername}"`);
            // --- END: AUDIT LOG ---
        });
    }

    return getUsers(token); // Return updated list

  } catch(e) {
    Logger.log(`Error in deleteUser: ${e}`);
    return {error: `Error deleting user: ${e.message}`};
  }
}

/**
 * Changes a user's password. Requires Admin role OR Matching Username.
 * @param {string} token The admin session token.
 * @param {string} username The username whose password needs changing.
 * @param {string} newPassword The new password.
 * @returns {object} {success: boolean, message: string}
 */
function changePassword(token, username, newPassword) {
    // 1. Check if user is authenticated at all (User OR Admin)
    if (!isUserAuthenticated(token)) {
        return { success: false, message: "Authentication failed. Please log in again." };
    }

    try {
        if (!username || typeof username !== 'string' || !username.trim() || !newPassword) {
            throw new Error("Username and new password are required.");
        }
         // Optional: Add minimum password length check
        if (newPassword.length < 6) {
           throw new Error("Password must be at least 6 characters long.");
        }

        const trimmedUsername = username.trim();
        
        // 2. Get Current Session Info
        const session = getSessionData(token);
        const currentRole = session.role;
        const currentUser = session.username;

        // 3. Permission Check
        // If not Admin AND attempting to change someone else's password -> Deny
        if (currentRole !== 'Admin' && String(currentUser).toLowerCase() !== trimmedUsername.toLowerCase()) {
             return { success: false, message: "Permission denied. You can only change your own password." };
        }

        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const userSheet = ss.getSheetByName(userSheetName);
        if (!userSheet || userSheet.getLastRow() <= 1) { // Need header + data
            throw new Error(`User sheet '${userSheetName}' not found or is empty.`);
        }

        // Check header and find column numbers
        const header = userSheet.getRange(1, 1, 1, 6).getValues()[0];
        const expectedHeader = ['Username', 'PasswordHash', 'Salt', 'Role', 'Status', 'Team'];
        // Not validating full header strictly here
        const usernameColumnNumber = 1; // A
        const hashColumnNumber = 2;       // B
        const saltColumnNumber = 3;       // C

        // --- START REFACTOR ---
        // Find the user row using indexOf (faster than TextFinder)
        const lastRow = userSheet.getLastRow();
        const usernames = userSheet.getRange(2, usernameColumnNumber, lastRow - 1, 1)
                                   .getValues()
                                   .flat()
                                   .map(u => String(u).trim().toLowerCase()); // Trim and normalize to lowercase
        
        // Find the index in the array (case-insensitive)
        const arrayIndex = usernames.indexOf(trimmedUsername.toLowerCase());

        if (arrayIndex === -1) {
            throw new Error(`User '${trimmedUsername}' not found`);
        }
        
        // Calculate the actual row number in the sheet
        // +2 because array is 0-based and we started data from row 2
        const rowIndex = arrayIndex + 2;
        // --- END REFACTOR ---

        // Create new hash and salt
        const newSalt = Utilities.base64Encode(Utilities.getUuid());
        const newHash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, newPassword + newSalt, Utilities.Charset.UTF_8));

        // Update hash and salt in the found row
        userSheet.getRange(rowIndex, hashColumnNumber).setValue(newHash);
        userSheet.getRange(rowIndex, saltColumnNumber).setValue(newSalt);
        Logger.log(`Password changed for user: ${trimmedUsername} in row ${rowIndex}`);

        // --- START: AUDIT LOG ---
        _logAdminAction(currentUser, "Change Password", `Password changed for user: "${trimmedUsername}"`);
        // --- END: AUDIT LOG ---

        return { success: true, message: "Password changed successfully" };

    } catch (error) {
        Logger.log(`Error in changePassword: ${error}`);
        return { success: false, message: `Error changing password: ${error.message}` };
    }
}

/**
 * Updates the Role of a specific user. Requires Admin role. Uses TextFinder for lookup.
 * @param {string} token The admin session token.
 * @param {string} username The username whose role needs changing.
 * @param {string} newRole The new role ('Admin' or 'User').
 * @returns {object} {success: boolean, error?: string} or updated user list {username, role}[]
 */
function updateUserRole(token, username, newRole) {
    if (!isUserAdmin(token)) {
        return { success: false, error: "Authentication failed. Please log in again." };
    }
    try {
        // Validate inputs
        if (!username || typeof username !== 'string' || !username.trim()) {
            throw new Error("Valid username is required.");
        }
        if (!newRole || (newRole !== 'Admin' && newRole !== 'User')) {
            throw new Error("Invalid role specified. Must be 'Admin' or 'User'.");
        }
        const trimmedUsername = username.trim();

        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const userSheet = ss.getSheetByName(userSheetName);
        if (!userSheet || userSheet.getLastRow() <= 1) { // Need header + data
            throw new Error(`User sheet '${userSheetName}' not found or is empty.`);
        }

        // Check header and find column numbers
        const header = userSheet.getRange(1, 1, 1, 6).getValues()[0];
        // Not validating full header strictly here
        const usernameColumnNumber = 1; // A
        const roleColumnNumber = 4;       // D

        // --- START REFACTOR ---
        // Find the user row using indexOf (faster than TextFinder)
        const lastRow = userSheet.getLastRow();
        const usernames = userSheet.getRange(2, usernameColumnNumber, lastRow - 1, 1)
                                   .getValues()
                                   .flat()
                                   .map(u => String(u).trim().toLowerCase()); // Trim and normalize to lowercase
        
        // Find the index in the array (case-insensitive)
        const arrayIndex = usernames.indexOf(trimmedUsername.toLowerCase());

        if (arrayIndex === -1) {
            throw new Error(`User '${trimmedUsername}' not found`);
        }

        // Calculate the actual row number in the sheet
        const rowIndex = arrayIndex + 2;
        // --- END REFACTOR ---

        // Update the Role in the found row
        userSheet.getRange(rowIndex, roleColumnNumber).setValue(newRole);
        Logger.log(`Role updated for user: ${trimmedUsername} to ${newRole} in row ${rowIndex}`);

        // --- START: AUDIT LOG ---
        const session = getSessionData(token);
        _logAdminAction(session.username, "Update User Role", `Set role for "${trimmedUsername}" to "${newRole}"`);
        // --- END: AUDIT LOG ---

        // Return success and optionally the updated user list
        // Returning the list might be useful for the client to refresh its display
        // return getUsers(token); // Option 1: Return updated list
        return { success: true }; // Option 2: Just return success status

    } catch (error) {
        Logger.log(`Error in updateUserRole: ${error}`);
        return { success: false, error: `Error updating role: ${error.message}` };
    }
}


// ===== START: NEW/MODIFIED COMBINED FUNCTION =====
/**
 * Saves status, internal admin comment, public comment, AND assigned user.
 * Logs status changes and assignment changes to the comment history.
 * Requires authenticated user.
 * @param {string} token The user session token.
 * @param {string} ticketId The ID of the ticket to update.
 * @param {string} newStatus The new status value (Thai).
 * @param {string} newCommentText The new *internal* comment text to add (can be empty).
 * @param {string} publicCommentText The new *public* resolution comment (only provided on close).
 * @param {string} assignedToUsername The new username to assign (can be empty for unassigned).
 * @returns {object} {success: boolean, newComments: string} or {success: false, error: string}
 */
function saveStatusAndComment(token, ticketId, newStatus, newCommentText, publicCommentText, assignedToUsername, adminFiles) {
    // 1. Auth Check (Check before lock for performance)
    if (!isUserAuthenticated(token)) {
        return { success: false, error: "Authentication failed. Please log in again." };
    }
    
    // 2. Validation
    const validStatuses = ['ยังไม่ดำเนินการ', 'กำลังดำเนินการ', 'เสร็จสิ้น', 'ยกเลิก'];
    if (!ticketId || !newStatus || !validStatuses.includes(newStatus)) {
        return { success: false, error: "Invalid Ticket ID or Status." };
    }
    
    // --- START: LockService Implementation ---
    const lock = LockService.getScriptLock();

    try {
        // Wait up to 30 seconds for the lock to become available.
        lock.waitLock(30000); // 30 seconds
        Logger.log(`Lock acquired for saveStatusAndComment ticket: ${ticketId}`);

        // 3. Get User and Timestamp (needed for all logs)
        const session = getSessionData(token);
        if (!session) {
             throw new Error("User session error. Please log in again.");
        }
        const username = session.username;

        const timestamp = new Date().toLocaleString('th-TH', { 
            timeZone: Session.getScriptTimeZone(),
            year: 'numeric', month: '2-digit', day: '2-digit', 
            hour: '2-digit', minute: '2-digit'
        });
        // Create the log prefix once
        const logPrefix = `[${timestamp} - ${username}]:`;

        // 4. Find Sheet and Columns
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const dataSheet = ss.getSheetByName(dataSheetName);
        if (!dataSheet || dataSheet.getLastRow() < 2) { // Need header + data
             throw new Error(`Sheet '${dataSheetName}' not found or empty.`);
        }
        
        // Use the cached headers if available, otherwise fetch
        const cache = CacheService.getScriptCache();
        let headers;
        const cachedHeaders = cache.get('all_data_headers_v2');
        if (cachedHeaders) {
          headers = JSON.parse(cachedHeaders);
        } else {
          headers = dataSheet.getRange(1, 1, 1, dataSheet.getLastColumn()).getValues()[0];
          // We could cache them here, but getPaginatedData will do it soon anyway
        }
        
        const ticketIdColumnIndex = headers.indexOf("เลขที่"); // 0-based
        if (ticketIdColumnIndex === -1) throw new Error("Column 'เลขที่' not found.");
        
        const statusColumnIndex = headers.indexOf("สถานะ"); // 0-based
        if (statusColumnIndex === -1) throw new Error("Column 'สถานะ' not found.");
        
        const commentColumnIndex = headers.indexOf("Admin Comments"); // 0-based
        if (commentColumnIndex === -1) throw new Error("Column 'Admin Comments' not found.");
        
        const publicCommentIndex = headers.indexOf("Public Comment"); // 0-based
        if (publicCommentIndex === -1) throw new Error("Column 'Public Comment' not found.");
        
        // ===== START: MODIFICATION (Find Assigned To Column) =====
        const assignedToIndex = headers.indexOf("Assigned To"); // 0-based
        if (assignedToIndex === -1) throw new Error("Column 'Assigned To' not found.");
        // ===== END: MODIFICATION =====

        // ===== START: MODIFICATION (Optimized Row Finder) =====
        // 5. Find Row (using the *Optimized* TextFinder method)
        const ticketIdColumnNumber = ticketIdColumnIndex + 1; // 1-based (e.g., Col 1 for A)
        
        // Create a TextFinder for the specific Ticket ID column (e.g., Column A, starting row 2)
        const ticketColumnRange = dataSheet.getRange(2, ticketIdColumnNumber, dataSheet.getLastRow() - 1, 1);
        const textFinder = ticketColumnRange.createTextFinder(ticketId)
                                          .matchEntireCell(true) // Match the whole cell
                                          .matchCase(true);      // Case-sensitive match

        const foundCell = textFinder.findNext(); // Find the first match

        if (!foundCell) {
            // If TextFinder didn't find it
            throw new Error("Ticket ID '" + ticketId + "' not found (using TextFinder).");
        }

        const foundRow = foundCell.getRow(); // Get the exact row number
        // ===== END: MODIFICATION (Optimized Row Finder) =====
        
        // 6. Get Cells and Existing Values
        const statusCell = dataSheet.getRange(foundRow, statusColumnIndex + 1);
        const commentCell = dataSheet.getRange(foundRow, commentColumnIndex + 1);
        const publicCommentCell = dataSheet.getRange(foundRow, publicCommentIndex + 1);
        
        // ===== START: MODIFICATION (Get Assigned To Cell) =====
        const assignedToCell = dataSheet.getRange(foundRow, assignedToIndex + 1);
        // ===== END: MODIFICATION =====
        
        const oldStatus = statusCell.getValue() ? String(statusCell.getValue()) : '';
        const existingComments = commentCell.getValue() ? String(commentCell.getValue()) : '';
        
        // ===== START: MODIFICATION (Get old Assigned To value) =====
        const oldAssignedTo = assignedToCell.getValue() ? String(assignedToCell.getValue()) : '';
        // ===== END: MODIFICATION =====
        
        // ===== START: MODIFICATION (Get Old Public Comment for Archive) =====
        const oldPublicComment = publicCommentCell.getValue() ? String(publicCommentCell.getValue()) : '';
        // ===== END: MODIFICATION =====

        // ===== NEW: Get Admin Attachments Cell & Value =====
        // Check for Admin Attachments Column (Assumed Col 13 / Index 12)
        let adminFileIndex = headers.findIndex(h => h.toLowerCase().includes('admin attachment') || h.includes('รูปปิดงาน'));
        if (adminFileIndex === -1 && headers.length >= 13) {
             adminFileIndex = 12; // Fallback to 13th column (index 12) if header not found but data exists
        }
        
        let adminFilesCell = null;
        let oldAdminFiles = '';
        if (adminFileIndex !== -1) {
            adminFilesCell = dataSheet.getRange(foundRow, adminFileIndex + 1);
            oldAdminFiles = adminFilesCell.getValue() ? String(adminFilesCell.getValue()) : '';
        }
        // ===================================================

        let newLogEntries = []; // An array to hold all new log lines

        // 7. --- Status Change Logic ---
        // Check if the status has actually changed
        if (oldStatus !== newStatus) {
            
            // ===== START: MODIFICATION (Hybrid Thai/English Status Log) =====
            // Helper map to translate Thai status to English
            const statusMap = {
              'ยังไม่ดำเนินการ': 'New',
              'กำลังดำเนินการ': 'In Progress',
              'เสร็จสิ้น': 'Completed',
              'ยกเลิก': 'Cancelled',
              '': '(Blank)' // Handle empty/blank status
            };

            // Translate old and new status, fallback to original if not in map
            const oldStatusEng = statusMap[oldStatus] || oldStatus || '(Blank)';
            const newStatusEng = statusMap[newStatus] || newStatus;

            // Create a log entry for the status change (Thai text, English statuses)
            // Added ⚙️ icon
            const statusLogEntry = `${logPrefix} ⚙️ เปลี่ยนสถานะจาก '${oldStatusEng}' เป็น '${newStatusEng}'`;
            // ===== END: MODIFICATION (Hybrid Thai/English Status Log) =====
            
            newLogEntries.push(statusLogEntry);
            
            // Now, actually update the status cell
            statusCell.setValue(newStatus);
            Logger.log(`Status updated for ${ticketId} to ${newStatus}`);
        }

        // ===== START: NEW RE-OPEN LOGIC (Archive History) =====
        const isOldStatusClosed = (oldStatus === 'เสร็จสิ้น' || oldStatus === 'ยกเลิก');
        const isNewStatusActive = (newStatus === 'ยังไม่ดำเนินการ' || newStatus === 'กำลังดำเนินการ');

        if (isOldStatusClosed && isNewStatusActive) {
            // Archive old public comment if exists
            if (oldPublicComment) {
                 const historyLogEntry = `${logPrefix} 🗂️ [History] ข้อความปิดงานรอบที่แล้ว: "${oldPublicComment}"`;
                 newLogEntries.push(historyLogEntry);
            }
            // Clear public comment
            publicCommentCell.setValue('');
            
            // NEW: Archive and Clear Old Admin Attachments
            if (oldAdminFiles && adminFilesCell) {
                 const historyFileEntry = `${logPrefix} 🗂️ [History] รูปภาพปิดงานรอบที่แล้ว: ${oldAdminFiles}`;
                 newLogEntries.push(historyFileEntry);
                 adminFilesCell.setValue(''); // Clear the cell
                 Logger.log(`Cleared Admin Attachments for Ticket ID ${ticketId} due to re-open.`);
            }
        }
        // ===== END: NEW RE-OPEN LOGIC =====

        // 8. --- New Internal Comment Logic (Smart Tagging) ---
        // Check if the user actually typed a new *internal* comment
        const trimmedComment = newCommentText ? newCommentText.trim() : '';
        if (trimmedComment) {
            // Create a log entry for the new comment
            // Added 📝 [Note]: prefix
            const commentLogEntry = `${logPrefix} 📝 [Note]: ${trimmedComment}`;
            newLogEntries.push(commentLogEntry);
            Logger.log(`Internal comment added for Ticket ID ${ticketId} by ${username}`);
        }

        // ===== START: MODIFICATION (Public Comment Logic) =====
        // 9. --- New Public Comment Logic ---
        const trimmedPublicComment = publicCommentText ? publicCommentText.trim() : '';
        
        // Only update if it's NOT a re-open clear action (which passes empty string usually)
        // If trimmedPublicComment has value, it means user is Closing or Updating public note explicitly.
        if (trimmedPublicComment) {
          publicCommentCell.setValue(trimmedPublicComment);
          Logger.log(`Public comment added for Ticket ID ${ticketId}`);
          
          // Added 📢 icon
          const publicLogEntry = `${logPrefix} 📢 (ได้บันทึกคอมเมนต์สรุปถึงผู้ใช้แล้ว): "${trimmedPublicComment}"`;
          newLogEntries.push(publicLogEntry);
        }
        // ===== END: MODIFICATION =====
        
        // ===== START: NEW (Assigned To Logic) =====
        // 10. --- Assign To Logic ---
        // Check if the assigned user has changed. `assignedToUsername` can be "" (for Unassigned)
        const newAssignedTo = (assignedToUsername === null || assignedToUsername === undefined) ? '' : String(assignedToUsername).trim();
        
        if (oldAssignedTo !== newAssignedTo) {
            const oldAssignee = oldAssignedTo || '(ยังไม่ระบุหน่วยงาน)';
            const newAssignee = newAssignedTo || '(ยังไม่ระบุหน่วยงาน)';
            
            // Create a log entry
            // Added ⚙️ icon
            const assignLogEntry = `${logPrefix} ⚙️ เปลี่ยนหน่วยงานที่รับผิดชอบจาก '${oldAssignee}' เป็น '${newAssignee}'`;
            newLogEntries.push(assignLogEntry);
            
            // Update the cell
            assignedToCell.setValue(newAssignedTo);
            Logger.log(`Assignment updated for ${ticketId} to ${newAssignee}`);
        }
        // ===== END: NEW (Assigned To Logic) =====

        // ===== START: NEW (Admin Files Logic) =====
        if (adminFiles && adminFiles.length > 0) {
            try {
                const mainFolder = DriveApp.getFolderById(uploadFolderId);
                let targetFolder;
                const subFolders = mainFolder.getFoldersByName(ticketId);
                if (subFolders.hasNext()) {
                    targetFolder = subFolders.next();
                } else {
                    targetFolder = mainFolder.createFolder(ticketId);
                }

                let newAdminFileUrls = [];
                for (const fileObj of adminFiles) {
                    if (fileObj.base64) {
                        const decoded = Utilities.base64Decode(fileObj.base64, Utilities.Charset.UTF_8);
                        const blob = Utilities.newBlob(decoded, fileObj.mimeType, "ADMIN_" + fileObj.fileName);
                        const newFile = targetFolder.createFile(blob);
                        newAdminFileUrls.push(newFile.getUrl());
                    }
                }
                
                if (newAdminFileUrls.length > 0) {
                    // We already retrieved adminFilesCell and oldAdminFiles above
                    if (adminFilesCell) {
                         // Refresh old value in case it was cleared by re-open logic (though unlikely to re-open AND upload in same transaction, good to be safe)
                         // Actually, if re-open cleared it, we start fresh. If not re-open, we append.
                         // But wait, re-open logic runs BEFORE this. So if re-opened, cell is empty.
                         // We need to re-read the value? No, we set it to '' in memory but didn't flush?
                         // GAS operations are batch-like but setValue happens immediately in script object.
                         // However, to be safe, let's use the logic:
                         
                         let currentValToAppend = "";
                         if (isOldStatusClosed && isNewStatusActive) {
                             currentValToAppend = ""; // It was cleared
                         } else {
                             currentValToAppend = oldAdminFiles; // Keep existing
                         }
                         
                         const updatedAdminFiles = currentValToAppend ? currentValToAppend + "," + newAdminFileUrls.join(",") : newAdminFileUrls.join(",");
                         adminFilesCell.setValue(updatedAdminFiles);
                         
                         // Added ⚙️ icon
                         const fileLogEntry = `${logPrefix} ⚙️ อัปโหลดรูปภาพเพิ่มเติม (Admin) จำนวน ${newAdminFileUrls.length} รูป`;
                         newLogEntries.push(fileLogEntry);
                    }
                }

            } catch (e) {
                 Logger.log("Admin file upload error: " + e.toString());
                 // Optionally add a log entry about failure but don't stop the whole process
            }
        }
        // ===== END: NEW (Admin Files Logic) =====

        
        // 11. --- Combine and Save Internal Logs ---
        let finalCommentString = existingComments; // Start with what's already there
        
        // Only update the comment cell if there are new entries to add
        if (newLogEntries.length > 0) {
            // Join all new entries (could be 1, 2, or 3 entries) with double newlines
            const combinedNewLogs = newLogEntries.join('\n\n');
            
            // Append new logs to existing comments
            finalCommentString = existingComments ? `${existingComments}\n\n${combinedNewLogs}` : combinedNewLogs;
            
            // Save the final, updated string back to the comment cell
            commentCell.setValue(finalCommentString);
        }

        // ===== START: UPDATE LAST UPDATED (Column O) =====
        // Update Last Updated (Column O / 15)
        dataSheet.getRange(foundRow, 15).setValue(new Date());
        // ===== END: UPDATE LAST UPDATED (Column O) =====

        // 12. Return Success
        // Always return the finalCommentString (even if it wasn't modified)
        // so the client UI stays in sync.
        
        // ===== START: CACHING MODIFICATION =====
        // An update happened, so the cache is now invalid.
        // const cache = CacheService.getScriptCache(); // Already defined above
        cache.remove('all_data_headers_v2');
        Logger.log('Cleared headers cache due to status/comment/assignment update.');
        // ===== END: CACHING MODIFICATION =====

        return { success: true, newComments: finalCommentString };

    } catch (error) {
        Logger.log(`Error in saveStatusAndComment: ${error} Stack: ${error.stack}`);
        return { success: false, error: "Error saving changes: " + error.message };
    } finally {
        // --- RELEASE LOCK ---
        if (lock) {
            lock.releaseLock();
            Logger.log("Lock released for saveStatusAndComment.");
        }
    }
}

/**
 * ฟังก์ชันสำหรับสร้างหน้า Print View (A4)
 * รับ Token และ Ticket ID (รับ Token มาแต่ไม่ได้ใช้ เพื่อให้ตรงกับหน้าบ้านที่ส่งมา 2 ตัวแปร)
 */
function generatePrintView(token, ticketId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(dataSheetName); // Use constant defined at top
  
  if (!sheet) {
    throw new Error("ไม่พบ Sheet ฐานข้อมูล (" + dataSheetName + ")");
  }

  // ดึงข้อมูลทั้งหมด (ใช้ getDisplayValues เพื่อให้วันที่แสดงเหมือนใน Sheet)
  const data = sheet.getDataRange().getDisplayValues();
  const headers = data[0];
  
  // Map หัวตาราง (Index)
  const idx = {
    id: headers.indexOf('เลขที่'),
    date: headers.indexOf('วันที่'),
    type: headers.indexOf('ประเภท'),
    topic: headers.indexOf('หัวข้อ'),
    details: headers.indexOf('รายละเอียด'),
    status: headers.indexOf('สถานะ'),
    location: headers.indexOf('สถานที่'),
    requester: headers.indexOf('Assigned To'), // Map requester to 'Assigned To' or change column name
    file: headers.findIndex(h => h.toLowerCase().includes('ไฟล์แนบ') || h.toLowerCase().includes('attachment')),
    comments: headers.indexOf('Admin Comments')
  };

  // ค้นหาแถวที่ตรงกับ Ticket ID
  const row = data.find(r => r[idx.id] === ticketId);
  
  if (!row) {
    throw new Error("ไม่พบข้อมูล Ticket ID: " + ticketId);
  }

  // เตรียมข้อมูลใส่ Object
  const printData = {
    ticketId: row[idx.id],
    date: row[idx.date],
    type: row[idx.type],
    topic: row[idx.topic],
    details: row[idx.details] || '-',
    status: row[idx.status],
    location: row[idx.location] || '-',
    requester: (idx.requester !== -1 && row[idx.requester]) ? row[idx.requester] : 'ไม่ระบุ (Public User)',
    fileUrl: (idx.file !== -1) ? row[idx.file] : '',
    comments: (idx.comments !== -1) ? row[idx.comments] : '',
    // ดึงโลโก้ (ถ้ามีฟังก์ชัน getLogo หรือใส่ URL ตรงๆ ก็ได้)
    logoUrl: (typeof getLogo === 'function') ? getLogo() : '' 
  };

  // เรียกไฟล์ Template และส่งข้อมูลไป
  const template = HtmlService.createTemplateFromFile('print_template');
  template.data = printData; // ส่งตัวแปร data ไปให้ html
  
  return template.evaluate()
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL) // อนุญาตให้แสดงผลได้
      .getContent();
}

/**
 * Retrieves raw data for CSV export based on current filters.
 * @param {string} token The admin session token.
 * @param {object} filters Filter settings { type, status, assignee, dateRangeKey, dateStart, dateEnd }.
 * @returns {object} {success: boolean, data: Array<Array<any>>} or {success: false, error: string}
 */
function getExportData(token, filters) {
  // 1. Auth Check
  if (!isUserAuthenticated(token)) {
      return { success: false, error: "Authentication failed. Please log in again." };
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dataSheet = ss.getSheetByName(dataSheetName);
    
    if (!dataSheet || dataSheet.getLastRow() < 1) {
       return { success: true, data: [] };
    }

    // Get all data
    const range = dataSheet.getDataRange();
    const values = range.getValues();
    let headers = values[0];
    let dataRows = values.slice(1);

    // Indices
    const typeIndex = headers.indexOf("ประเภท");
    const statusIndex = headers.indexOf("สถานะ");
    const dateIndex = headers.indexOf("วันที่");
    const assignIndex = headers.indexOf("Assigned To");

    // Filter Logic (Copied and adapted from getPaginatedData to ensure consistency)
    const filterType = filters.type || 'ทั้งหมด';
    const filterStatus = filters.status || 'ทั้งหมด';
    const filterAssignee = (filters.assignee !== undefined) ? filters.assignee : 'ทั้งหมด';
    
    // --- Date Filter Setup ---
    const dateRangeKey = filters.dateRangeKey || 'all';
    let filterStartDate = null;
    let filterEndDate = null;
    const now = new Date();

    if (dateRangeKey === 'custom' && filters.dateStart && filters.dateEnd) {
        filterStartDate = new Date(filters.dateStart + 'T00:00:00');
        let tempEnd = new Date(filters.dateEnd + 'T00:00:00');
        tempEnd.setDate(tempEnd.getDate() + 1);
        filterEndDate = tempEnd;
    } else if (dateRangeKey !== 'all') {
        const getStartOfDay = (d) => { let x = new Date(d); x.setHours(0,0,0,0); return x; };
        switch (dateRangeKey) {
            case 'today':
                filterStartDate = getStartOfDay(now);
                filterEndDate = new Date(filterStartDate); filterEndDate.setDate(filterEndDate.getDate() + 1);
                break;
            case 'this_week':
                let day = now.getDay(); let diff = now.getDate() - day + (day == 0 ? -6 : 1);
                filterStartDate = getStartOfDay(new Date(now.setDate(diff)));
                filterEndDate = new Date(filterStartDate); filterEndDate.setDate(filterEndDate.getDate() + 7);
                break;
            case 'this_month':
                filterStartDate = new Date(now.getFullYear(), now.getMonth(), 1);
                filterEndDate = new Date(now.getFullYear(), now.getMonth() + 1, 1);
                break;
             case 'last_month':
                filterStartDate = new Date(now.getFullYear(), now.getMonth() - 1, 1);
                filterEndDate = new Date(now.getFullYear(), now.getMonth(), 1);
                break;
        }
    }

    // Apply Filters
    let filteredData = dataRows.filter(row => {
        // 1. Type
        if (filterType !== 'ทั้งหมด' && typeIndex > -1 && row[typeIndex] !== filterType) return false;
        
        // 2. Status
        if (filterStatus !== 'ทั้งหมด' && statusIndex > -1 && row[statusIndex] !== filterStatus) return false;
        
        // 3. Assignee
        if (filterAssignee !== 'ทั้งหมด' && assignIndex > -1) {
             const rowAssignee = row[assignIndex] || '';
             if (rowAssignee !== filterAssignee) return false;
        }

        // 4. Date
        if (filterStartDate && filterEndDate && dateIndex > -1) {
            const rowDate = row[dateIndex];
            if (rowDate instanceof Date) {
                if (rowDate < filterStartDate || rowDate >= filterEndDate) return false;
            } else {
                return false; // Invalid date in row
            }
        }

        return true;
    });

    // Format Dates for CSV (Optional, but good for raw data readability)
    filteredData = filteredData.map(row => {
        if (dateIndex > -1 && row[dateIndex] instanceof Date) {
            row[dateIndex] = Utilities.formatDate(row[dateIndex], Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
        }
        
        // *** START: FIX FOR DATE OBJECT SERIALIZATION ERROR IN EXPORT ***
        // Ensure all Date objects in the row are converted to strings
        // This is crucial because google.script.run cannot return Date objects within arrays in some contexts,
        // leading to "Cannot read properties of null (reading 'summary')" or similar errors on the client side.
        return row.map(cell => {
            if (cell instanceof Date) {
                return Utilities.formatDate(cell, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
            }
            return cell;
        });
        // *** END: FIX ***
    });

    // Combine Header + Filtered Rows
    const result = [headers, ...filteredData];
    
    return { success: true, data: result };

  } catch (e) {
    Logger.log("Error in getExportData: " + e.message);
    return { success: false, error: e.message };
  }
}

// ===== START: NEW FUNCTION (Check Ticket Status - Multilingual Support) =====
/**
 * Publicly callable function for users to check their ticket status.
 * Requires Access Key and Language code. (Modified to search by Access Key only)
 * @param {string} accessKey The secret Access Key.
 * @param {string} lang Language code ('th', 'jp', 'my').
 * @returns {object} {success, data, message}
 */
function getTicketStatus(accessKey, lang) {
  // Default to 'th' if no lang provided
  const currentLang = lang || 'th';
  const t = getTranslations(currentLang); // Fetch translations for error messages

  try {
    const trimmedAccessKey = accessKey ? accessKey.trim() : '';

    if (!trimmedAccessKey) {
      return { success: false, message: t.checkStatusFail || 'กรุณากรอกรหัสลับ (Please enter Access Key)' };
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dataSheet = ss.getSheetByName(dataSheetName);
    if (!dataSheet || dataSheet.getLastRow() < 2) {
      return { success: false, message: t.checkStatusFail || 'ไม่พบข้อมูลคำร้อง' };
    }

    // Get Headers
    const headers = dataSheet.getRange(1, 1, 1, dataSheet.getLastColumn()).getValues()[0];
    const ticketIdIndex = headers.indexOf("เลขที่");
    const accessKeyIndex = headers.indexOf("Access Key");
    const dateIndex = headers.indexOf("วันที่");
    const typeIndex = headers.indexOf("ประเภท");
    const statusIndex = headers.indexOf("สถานะ");
    const publicCommentIndex = headers.indexOf("Public Comment");
    
    // NEW: Find Admin Files Column (Index 12 / Col M usually, but search header to be safe)
    let adminFileIndex = headers.findIndex(h => h.toLowerCase().includes('admin attachment') || h.includes('รูปปิดงาน'));
    if (adminFileIndex === -1 && headers.length >= 13) {
         adminFileIndex = 12; // Fallback to 13th column (index 12) if header not found but data exists
    }

    if (ticketIdIndex === -1 || accessKeyIndex === -1 || statusIndex === -1 || dateIndex === -1 || typeIndex === -1 || publicCommentIndex === -1) {
      return { success: false, message: 'System Error: Columns missing.' };
    }

    // Get all data
    const dataRows = dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, headers.length).getValues();
    
    for (const row of dataRows) {
      // MODIFIED: Check ONLY Access Key
      // Note: We use String() to ensure type safety in comparison
      if (String(row[accessKeyIndex]).trim() === trimmedAccessKey) {
        
        // Found match
        let formattedDate = row[dateIndex]; // Default to created date

        // Try to find Last Updated (Column O / Index 14)
        // We assume Column O is index 14. Check if it exists and is valid.
        if (row.length > 14 && row[14] && row[14] instanceof Date) {
            formattedDate = Utilities.formatDate(row[14], Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
        } else if (formattedDate instanceof Date) {
            formattedDate = Utilities.formatDate(formattedDate, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
        }
        
        // --- Type Translation Logic ---
        const rawType = row[typeIndex] || '';
        const multilingualType = {
          'คำร้องเรียน': 'type_complaint',
          'ข้อเสนอแนะ': 'type_suggestion'
        };
        // Get translation key first
        const typeKey = multilingualType[rawType];
        // Then fetch translation, fallback to rawType
        const userFriendlyType = (typeKey && t[typeKey]) ? t[typeKey] : rawType;


        // --- Status Translation Logic ---
        const rawStatus = row[statusIndex] || 'ยังไม่ดำเนินการ';
        const multilingualStatus = {
          'ยังไม่ดำเนินการ': 'status_pending',
          'กำลังดำเนินการ': 'status_inprogress',
          'เสร็จสิ้น': 'status_completed',
          'ยกเลิก': 'status_cancelled'
        };
        
        // Get status key, then fetch translation
        const statusKey = multilingualStatus[rawStatus];
        const userFriendlyStatus = (statusKey && t[statusKey]) ? t[statusKey] : rawStatus;

        // Handle Public Comment
        const publicComment = row[publicCommentIndex] || '';
        
        // NEW: Handle Admin Files
        const adminFiles = (adminFileIndex !== -1) ? row[adminFileIndex] : '';

        return { 
          success: true, 
          data: {
            ticketId: row[ticketIdIndex], // Still return Ticket ID for display
            date: String(formattedDate),
            type: userFriendlyType,
            status: userFriendlyStatus,
            publicComment: publicComment,
            rawStatus: rawStatus,
            adminFiles: adminFiles 
          } 
        };
      }
    }

    // Not found
    return { success: false, message: t.checkStatusFail || 'รหัสลับไม่ถูกต้อง หรือไม่พบข้อมูล' };

  } catch (e) {
    Logger.log(`Error in getTicketStatus: ${e}`);
    return { success: false, message: 'Error: ' + e.message };
  }
}
// ===== END: NEW FUNCTION (Check Ticket Status) =====

// ===============================================================
//                  TEAM MANAGEMENT FUNCTIONS (NEW)
// ===============================================================

/**
 * Retrieves all teams from the master 'Teams' sheet.
 * @param {string} token
 * @returns {object} {success: boolean, data: string[]}
 */
function getTeams(token) {
    if (!isUserAuthenticated(token)) return { error: "Authentication failed." };
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        let sheet = ss.getSheetByName(teamSheetName);
        if (!sheet) {
            sheet = ss.insertSheet(teamSheetName);
            sheet.appendRow(['TeamName']);
            return { success: true, data: [] };
        }
        if (sheet.getLastRow() < 2) return { success: true, data: [] };

        const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat();
        const teams = data.map(t => String(t).trim()).filter(t => t);
        return { success: true, data: [...new Set(teams)].sort() }; // Unique & Sorted
    } catch (e) {
        Logger.log("Error getting teams: " + e);
        return { success: false, error: e.message };
    }
}

/**
 * Adds a new team to the master 'Teams' sheet.
 * @param {string} token
 * @param {string} teamName
 */
function saveTeam(token, teamName) {
    if (!isUserAdmin(token)) {
        return { success: false, error: "Authentication failed. Admin role required." };
    }
    const name = String(teamName || '').trim();
    if (!name) return { success: false, error: "Team name cannot be empty." };

    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        let sheet = ss.getSheetByName(teamSheetName);
        if (!sheet) {
            sheet = ss.insertSheet(teamSheetName);
            sheet.appendRow(['TeamName']);
        }

        // Check duplicates
        const existingTeams = getTeams(token).data;
        if (existingTeams.includes(name)) {
            return { success: false, error: `Team "${name}" already exists.` };
        }

        sheet.appendRow([name]);

        const session = getSessionData(token);
        _logAdminAction(session.username, "Add Team", `Team: "${name}"`);

        return { success: true };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

/**
 * Deletes a team from the master 'Teams' sheet.
 * @param {string} token
 * @param {string} teamName
 */
function deleteTeam(token, teamName) {
   if (!isUserAdmin(token)) {
       return { success: false, error: "Authentication failed. Admin role required." };
   }
   const name = String(teamName || '').trim();

   try {
       const ss = SpreadsheetApp.getActiveSpreadsheet();
       const sheet = ss.getSheetByName(teamSheetName);
       if (!sheet) return { success: false, error: "Sheet not found" };

       const data = sheet.getDataRange().getValues();
       let rowIndex = -1;
       
       // Start from 1 to skip header
       for (let i = 1; i < data.length; i++) {
           if (String(data[i][0]).trim() === name) {
               rowIndex = i + 1;
               break;
           }
       }

       if (rowIndex !== -1) {
           sheet.deleteRow(rowIndex);
           const session = getSessionData(token);
           _logAdminAction(session.username, "Delete Team", `Team: "${name}"`);
           return { success: true };
       } else {
           return { success: false, error: "Team not found." };
       }
   } catch (e) {
       return { success: false, error: e.message };
   }
}