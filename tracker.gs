const CONFIG_AUDIT = {
  logColors: {
    SYSTEM_INIT: '#D1ECF1',
    CELL_EDIT: '#ff9223',
    CELL_ADD: '#93b474',
    CELL_DELETE: '#ff0000',
    FORMULA_EDIT: '#FFECB3',
    SHEET_RENAME: '#d44b4b',
    SHEET_DELETION: '#970000',
    SHEET_CREATION: '#93b474',
    SHEET_PROTECTED: '#FFCDD2',
    SHEET_UNPROTECTED: '#C8E6C9',
    PROTECTION_MODIFIED: '#FFE082',
    ERROR: '#FF6B6B'
  },
  maxChangeLogRows: 10000,
  periodicInterval: 5
};

//MAIN SETUP
function SETUP_ENHANCED_MULTIUSER_TRACKER() {
  try {
    console.log('🚀 Setting up MULTI-USER tracker...');
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = initializeLogSheet(ss);
    
    PropertiesService.getScriptProperties().setProperties({
      'USE_LOCAL_SHEET': 'true',
      'LOCAL_SHEET_NAME': 'ChangeLog',
      'TRACK_ALL': 'true',
      'MULTIUSER_MODE': 'true'
    });
    
    setupInstallableTriggers();
    initializeSheetTracking();
    initializeStructureTracking();
    setupMultiUserPolling();
    
    console.log('✅ MULTI-USER Setup complete!');
    console.log('📊 Changes from ALL USERS will be logged to: ' + ss.getName() + ' → ChangeLog sheet');
    
    logChangeLocal({
      timestamp: new Date(),
      sheetName: 'SYSTEM',
      cellAddress: 'INITIALIZATION',
      oldValue: '',
      newValue: 'Multi-user tracker initialized',
      user: getUserEmail({}),
      changeType: 'SYSTEM_INIT',
      formula: '',
      details: 'Installable triggers created for all users',
      userType: 'OWNER'
    });
    
    return 'SUCCESS: Multi-user tracker active! ALL users\' changes will be tracked.';
    
  } catch (error) {
    console.error('Setup failed:', error);
    return 'ERROR: ' + error.toString();
  }
}

//LOG SHEET INITIALIZATION
function initializeLogSheet(ss) {


  let logSheet = ss.getSheetByName('ChangeLog');
  
  if (!logSheet) {
    logSheet = ss.insertSheet('ChangeLog');
    console.log('✅ Created ChangeLog sheet');
  }
  
  if (logSheet.getLastRow() === 0) {
    const headers = ['Timestamp', 'Sheet Name', 'Cell Address', 'Old Value', 'New Value', 
                     'User', 'Change Type', 'Formula', 'Details', 'User Type'];
    
    logSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    logSheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#4285F4')
      .setFontColor('#FFFFFF');
    logSheet.setFrozenRows(1);
  }
  if (logSheet.getConditionalFormatRules().length === 0) {
    console.log('Setting up conditional formatting rules for ChangeLog...');
    const range = logSheet.getRange('G:G'); // Apply to the whole "Change Type" column
    const rules = [];
    
    for (const type in CONFIG_AUDIT.logColors) {
      rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(type)
        .setBackground(CONFIG_AUDIT.logColors[type])
        .setRanges([range])
        .build());
    }
    
    // Add rule for external users
    const userTypeRange = logSheet.getRange('J:J');
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextDoesNotContain('OWNER')
      .whenTextDoesNotContain('SYSTEM')
      .whenTextIsNotEmpty()
      .setBackground('#FFE5B4')
      .setRanges([userTypeRange])
      .build());

    logSheet.setConditionalFormatRules(rules);
  }
  return logSheet;
  
}

// ============ TRIGGER SETUP ============
function setupInstallableTriggers() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const triggerFunctions = ['onEditInstallable', 'onChangeInstallable', 'onOpenInstallable', 'periodicSheetCheck'];
    
    // Remove existing triggers
    ScriptApp.getProjectTriggers().forEach(trigger => {
      if (triggerFunctions.includes(trigger.getHandlerFunction())) {
        ScriptApp.deleteTrigger(trigger);
      }
    });
    
    // Create new triggers
    ScriptApp.newTrigger('onEditInstallable').forSpreadsheet(ss).onEdit().create();
    ScriptApp.newTrigger('onChangeInstallable').forSpreadsheet(ss).onChange().create();
    ScriptApp.newTrigger('onOpenInstallable').forSpreadsheet(ss).onOpen().create();
    ScriptApp.newTrigger('periodicSheetCheck').timeBased().everyMinutes(CONFIG_AUDIT.periodicInterval).create();
    
    console.log('✅ Installable triggers created');
  } catch (error) {
    console.error('Failed to setup installable triggers:', error);
    throw error;
  }
}

// ============ INITIALIZATION ============
function initializeSheetTracking() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const currentSheets = ss.getSheets()
    .filter(s => s.getName() !== 'ChangeLog')
    .map(s => {
      const protectionInfo = getSheetProtectionInfo(s);
      return {
        name: s.getName(),
        id: s.getSheetId(),
        lastModified: Date.now(),
        isProtected: protectionInfo.isProtected,
        protectionDescription: protectionInfo.description,
        editors: protectionInfo.editors || []
      };
    });
  
  const props = PropertiesService.getScriptProperties();
  props.setProperty('knownSheets', JSON.stringify(currentSheets));
  props.setProperty('lastSheetCheck', Date.now().toString());
  
  console.log('🔄 Sheet tracking initialized with', currentSheets.length, 'sheets');
}

function initializeStructureTracking() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const structures = {};
  
  ss.getSheets()
    .filter(s => s.getName() !== 'ChangeLog')
    .forEach(sheet => {
      structures[sheet.getSheetId()] = {
        name: sheet.getName(),
        maxRows: sheet.getMaxRows(),
        maxColumns: sheet.getMaxColumns(),
        lastRow: sheet.getLastRow(),
        lastColumn: sheet.getLastColumn()
      };
    });
  
  PropertiesService.getScriptProperties().setProperty('sheetStructures', JSON.stringify(structures));
  console.log('📐 Structure tracking initialized');
}

function setupMultiUserPolling() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const snapshot = {};
  
  ss.getSheets()
    .filter(s => s.getName() !== 'ChangeLog')
    .forEach(sheet => {
      const data = sheet.getDataRange().getValues();
      snapshot[sheet.getName()] = JSON.stringify(data);
    });
  
  PropertiesService.getScriptProperties().setProperty('dataSnapshot', JSON.stringify(snapshot));
  console.log('📸 Initial data snapshot created');
}

// ============ PROTECTION INFO ============
function getSheetProtectionInfo(sheet) {
  try {
    const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    
    if (protections.length === 0) {
      return { isProtected: false, description: '', editors: [], canCurrentUserEdit: true };
    }
    
    const protection = protections[0];
    return {
      isProtected: true,
      description: protection.getDescription() || 'No description',
      editors: protection.getEditors().map(e => e.getEmail()),
      canCurrentUserEdit: protection.canEdit(),
      rangeProtections: sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).length
    };
  } catch (error) {
    console.error('Error getting protection info:', error);
    return { isProtected: true, description: 'Error getting info', editors: [], canCurrentUserEdit: false };
  }
}

// ============ USER DETECTION ============
function getUserEmail(e) {
  try {
    // Try event user
    if (e && e.user) {
      const eventUser = e.user.getEmail ? e.user.getEmail() : e.user.toString();
      if (eventUser && eventUser !== '') return eventUser;
    }
    
    // Try active user
    const activeUser = Session.getActiveUser().getEmail();
    if (activeUser && activeUser !== '') return activeUser;
    
    // Try effective user
    const effectiveUser = Session.getEffectiveUser().getEmail();
    if (effectiveUser && effectiveUser !== '') return effectiveUser;
    
    // Try temporary key
    const tempKey = Session.getTemporaryActiveUserKey();
    if (tempKey && tempKey !== '') return 'TempUser_' + tempKey.substring(0, 8);
    
    // Fallback
    return 'Unidentified_' + getContextClues(e);
    
  } catch (error) {
    console.error('User detection error:', error);
    return 'Error_' + Date.now().toString().substring(7);
  }
}

function detectUserType(e, userEmail) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const file = DriveApp.getFileById(ss.getId());
    
    const owner = file.getOwner();
    if (owner && owner.getEmail() === userEmail) return 'OWNER';
    
    const editors = file.getEditors();
    if (editors.some(editor => editor.getEmail() === userEmail)) return 'EDITOR';
    
    const viewers = file.getViewers();
    if (viewers.some(viewer => viewer.getEmail() === userEmail)) return 'VIEWER_WITH_COMMENT';
    
    if (userEmail.startsWith('External_')) return 'EXTERNAL_FINGERPRINTED';
    if (userEmail.startsWith('TempUser_')) return 'TEMPORARY';
    if (userEmail.startsWith('Unidentified_')) return 'UNIDENTIFIED';
    
    return 'EXTERNAL_USER';
    
  } catch (error) {
    console.error('User type detection error:', error);
    return 'TYPE_DETECTION_ERROR';
  }
}

function detectExternalUser(e, range, oldValue, newValue) {
  let email = getUserEmail(e);
  let type = 'UNKNOWN';
  let fingerprint = '';
  
  if (!email || email === 'Unknown User' || email.startsWith('Unidentified_') || !email.includes('@')) {
    fingerprint = createUserFingerprint(e, range, oldValue, newValue);
    
    const patternMatch = matchUserPattern(fingerprint);
    if (patternMatch && !email.includes('@')) {
      email = patternMatch;
      type = 'EXTERNAL_PATTERN_MATCH';
    } else {
      email = 'External_' + fingerprint;
      type = 'EXTERNAL_FINGERPRINTED';
      storeExternalFingerprint(fingerprint, range, newValue);
    }
  } else {
    type = detectUserType(e, email);
  }
  
  return { email, type, fingerprint };
}

function createUserFingerprint(e, range, oldValue, newValue) {
  const components = [];
  
  try {
    const hour = new Date().getHours();
    components.push('T' + Math.floor(hour / 6));
    
    const col = range.getColumn();
    const row = range.getRow();
    components.push('L' + (col % 3) + '' + (row % 3));
    
    if (oldValue && !newValue) {
      components.push('CLR');
    } else if (typeof newValue === 'number') {
      components.push('NUM');
    } else if (newValue.length > 50) {
      components.push('LNG');
    } else {
      components.push('TXT');
    }
    
    components.push('S' + (range.getSheet().getIndex() % 5));
    components.push('M' + Math.floor(new Date().getMilliseconds() / 100));
    
  } catch (err) {
    console.log('Fingerprinting error:', err);
  }
  
  return components.join('_');
}

function matchUserPattern(fingerprint) {
  try {
    const props = PropertiesService.getScriptProperties();
    const patterns = JSON.parse(props.getProperty('userPatterns') || '{}');
    
    for (let storedFingerprint in patterns) {
      const similarity = calculateSimilarity(fingerprint, storedFingerprint);
      if (similarity > 0.7) {
        return patterns[storedFingerprint].user;
      }
    }
  } catch (err) {
    console.log('Pattern matching error:', err);
  }
  return null;
}

function calculateSimilarity(fp1, fp2) {
  const parts1 = fp1.split('_');
  const parts2 = fp2.split('_');
  let matches = 0;
  
  for (let i = 0; i < Math.min(parts1.length, parts2.length); i++) {
    if (parts1[i] === parts2[i]) matches++;
  }
  
  return matches / Math.max(parts1.length, parts2.length);
}

function storeExternalFingerprint(fingerprint, range, value) {
  try {
    const props = PropertiesService.getScriptProperties();
    const patterns = JSON.parse(props.getProperty('userPatterns') || '{}');
    
    if (!patterns[fingerprint]) {
      patterns[fingerprint] = {
        user: 'External_' + fingerprint,
        firstSeen: new Date().toISOString(),
        editCount: 0,
        cells: []
      };
    }
    
    patterns[fingerprint].lastSeen = new Date().toISOString();
    patterns[fingerprint].editCount++;
    patterns[fingerprint].cells.push(range.getA1Notation());
    
    if (patterns[fingerprint].cells.length > 10) {
      patterns[fingerprint].cells.shift();
    }
    
    props.setProperty('userPatterns', JSON.stringify(patterns));
  } catch (err) {
    console.error('Failed to store fingerprint:', err);
  }
}

function getContextClues(e) {
  const clues = [];
  
  try {
    if (e) {
      if (e.authMode) clues.push(e.authMode.toString().replace('ScriptApp.AuthMode.', ''));
      if (e.source) clues.push('Src:' + e.source.getName().substring(0, 5));
      if (e.range) clues.push('Rng:' + e.range.getA1Notation());
    }
    clues.push(Date.now().toString().substring(7));
  } catch (err) {
    clues.push('NoContext');
  }
  
  return clues.join('_');
}

// ============ CHANGE DETECTION ============
function checkSheetChanges() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const currentSheets = ss.getSheets()
      .filter(s => s.getName() !== 'ChangeLog')
      .map(s => {
        const protectionInfo = getSheetProtectionInfo(s);
        return {
          name: s.getName(),
          id: s.getSheetId(),
          lastModified: Date.now(),
          isProtected: protectionInfo.isProtected,
          protectionDescription: protectionInfo.description,
          editors: protectionInfo.editors || []
        };
      });
    
    const props = PropertiesService.getScriptProperties();
    const storedSheets = JSON.parse(props.getProperty('knownSheets') || '[]');
    
    if (storedSheets.length === 0) {
      props.setProperty('knownSheets', JSON.stringify(currentSheets));
      return 'Sheet tracking initialized';
    }
    
    const storedIds = storedSheets.map(s => s.id);
    const currentIds = currentSheets.map(s => s.id);
    let changesFound = 0;
    
    // New sheets
    const newSheets = currentSheets.filter(s => !storedIds.includes(s.id));
    newSheets.forEach(sheet => {
      logChangeLocal({
        timestamp: new Date(),
        sheetName: sheet.name,
        cellAddress: 'ENTIRE_SHEET',
        oldValue: '[NONE]',
        newValue: 'SHEET_CREATED',
        user: getUserEmail({}),
        changeType: 'SHEET_CREATION',
        formula: '',
        details: sheet.isProtected ? 'Created with protection: ' + sheet.protectionDescription : 'Created without protection',
        userType: 'SYSTEM'
      });
      changesFound++;
    });
    
    // Deleted sheets
    const deletedSheets = storedSheets.filter(s => !currentIds.includes(s.id));
    deletedSheets.forEach(sheet => {
      logChangeLocal({
        timestamp: new Date(),
        sheetName: sheet.name,
        cellAddress: 'ENTIRE_SHEET',
        oldValue: 'SHEET_EXISTED',
        newValue: '[DELETED]',
        user: getUserEmail({}),
        changeType: 'SHEET_DELETION',
        formula: '',
        details: sheet.isProtected ? 'Was protected' : 'Was not protected',
        userType: 'SYSTEM'
      });
      changesFound++;
    });
    
    // Sheet changes (renames, protection)
    currentSheets.forEach(current => {
      const stored = storedSheets.find(s => s.id === current.id);
      if (!stored) return;
      
      // Rename
      if (stored.name !== current.name) {
        logChangeLocal({
          timestamp: new Date(),
          sheetName: current.name,
          cellAddress: 'SHEET_NAME',
          oldValue: stored.name,
          newValue: current.name,
          user: getUserEmail({}),
          changeType: 'SHEET_RENAME',
          formula: '',
          details: 'Renamed from: ' + stored.name,
          userType: 'SYSTEM'
        });
        changesFound++;
      }
      
      // Protection changes
      if (stored.isProtected !== current.isProtected) {
        logChangeLocal({
          timestamp: new Date(),
          sheetName: current.name,
          cellAddress: 'PROTECTION',
          oldValue: stored.isProtected ? 'PROTECTED' : 'UNPROTECTED',
          newValue: current.isProtected ? 'PROTECTED' : 'UNPROTECTED',
          user: getUserEmail({}),
          changeType: current.isProtected ? 'SHEET_PROTECTED' : 'SHEET_UNPROTECTED',
          formula: '',
          details: current.isProtected ? 
            'Protection added: ' + current.protectionDescription + ' | Editors: ' + current.editors.join(', ') :
            'Protection removed (was: ' + stored.protectionDescription + ')',
          userType: 'PROTECTION_CHANGE'
        });
        changesFound++;
      }
      
      // Protection settings modified
      if (current.isProtected && stored.isProtected) {
        const currentEditors = (current.editors || []).sort().join(',');
        const storedEditors = (stored.editors || []).sort().join(',');
        
        if (currentEditors !== storedEditors || current.protectionDescription !== stored.protectionDescription) {
          logChangeLocal({
            timestamp: new Date(),
            sheetName: current.name,
            cellAddress: 'PROTECTION_SETTINGS',
            oldValue: stored.protectionDescription + ' | Editors: ' + storedEditors,
            newValue: current.protectionDescription + ' | Editors: ' + currentEditors,
            user: getUserEmail({}),
            changeType: 'PROTECTION_MODIFIED',
            formula: '',
            details: 'Protection settings changed',
            userType: 'PROTECTION_CHANGE'
          });
          changesFound++;
        }
      }
    });
    
    props.setProperties({
      'knownSheets': JSON.stringify(currentSheets),
      'lastSheetCheck': Date.now().toString()
    });
    
    if (changesFound > 0) SpreadsheetApp.flush();
    return changesFound > 0 ? 'Found and logged ' + changesFound + ' changes' : 'No changes detected';
    
  } catch (error) {
    console.error('Sheet change check error:', error);
    return 'ERROR: ' + error.toString();
  }
}

function checkStructureChanges() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const props = PropertiesService.getScriptProperties();
    const storedStructures = JSON.parse(props.getProperty('sheetStructures') || '{}');
    const currentStructures = {};
    let changesFound = 0;
    
    ss.getSheets()
      .filter(s => s.getName() !== 'ChangeLog')
      .forEach(sheet => {
        const sheetId = sheet.getSheetId();
        const current = {
          name: sheet.getName(),
          maxRows: sheet.getMaxRows(),
          maxColumns: sheet.getMaxColumns(),
          lastRow: sheet.getLastRow(),
          lastColumn: sheet.getLastColumn()
        };
        
        currentStructures[sheetId] = current;
        const stored = storedStructures[sheetId];
        
        if (stored) {
          // Row changes
          if (current.maxRows !== stored.maxRows) {
            const rowDiff = current.maxRows - stored.maxRows;
            logChangeLocal({
              timestamp: new Date(),
              sheetName: current.name,
              cellAddress: rowDiff > 0 ? 
                'ROWS_' + (stored.maxRows + 1) + ':' + current.maxRows : 
                'ROWS_' + (current.maxRows + 1) + ':' + stored.maxRows,
              oldValue: stored.maxRows + ' rows',
              newValue: current.maxRows + ' rows',
              user: getUserEmail({}),
              changeType: rowDiff > 0 ? 'ROWS_ADDED' : 'ROWS_DELETED',
              formula: '',
              details: Math.abs(rowDiff) + ' row(s) ' + (rowDiff > 0 ? 'added' : 'deleted'),
              userType: 'SYSTEM'
            });
            changesFound++;
          }
          
          // Column changes
          if (current.maxColumns !== stored.maxColumns) {
            const colDiff = current.maxColumns - stored.maxColumns;
            logChangeLocal({
              timestamp: new Date(),
              sheetName: current.name,
              cellAddress: colDiff > 0 ? 
                'COLS_' + columnToLetter(stored.maxColumns + 1) + ':' + columnToLetter(current.maxColumns) : 
                'COLS_' + columnToLetter(current.maxColumns + 1) + ':' + columnToLetter(stored.maxColumns),
              oldValue: stored.maxColumns + ' columns',
              newValue: current.maxColumns + ' columns',
              user: getUserEmail({}),
              changeType: colDiff > 0 ? 'COLUMNS_ADDED' : 'COLUMNS_DELETED',
              formula: '',
              details: Math.abs(colDiff) + ' column(s) ' + (colDiff > 0 ? 'added' : 'deleted'),
              userType: 'SYSTEM'
            });
            changesFound++;
          }
        }
      });
    
    props.setProperty('sheetStructures', JSON.stringify(currentStructures));
    return changesFound;
    
  } catch (error) {
    console.error('Structure change check error:', error);
    return 0;
  }
}

function columnToLetter(column) {
  let temp = '';
  let letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

function logChangeLocal(data) {
  try {
    const lock = LockService.getScriptLock();
    lock.waitLock(10000); // Wait up to 10s for other processes
    
    const cache = CacheService.getScriptCache();
    const logCacheKey = 'auditLogCache';
    let logQueue = JSON.parse(cache.get(logCacheKey) || '[]');
    
    const rowData = [
      data.timestamp.toISOString(), // Store as string for JSON
      data.sheetName || '',
      data.cellAddress || '',
      data.oldValue || '',
      data.newValue || '',
      data.user || 'Unknown',
      data.changeType || 'UNKNOWN',
      data.formula || '',
      data.details || '',
      data.userType || 'UNKNOWN'
    ];
    
    logQueue.push(rowData);
    cache.put(logCacheKey, JSON.stringify(logQueue), 600); // Store for 10 min
    
    lock.releaseLock();
    console.log('✅ Queued log:', data.changeType);
    return true;
    
  } catch (error) {
    console.error('Logging failed:', error);
    return false;
  }
}

// ============ TRIGGER HANDLERS ============
function onEditInstallable(e) {
  try {
    if (!e || !e.range) return;
    
    const range = e.range;
    const sheet = range.getSheet();
    if (sheet.getName() === 'ChangeLog') return;
    
    const userInfo = detectExternalUser(e, range, e.oldValue, e.value);
    
    // Check if e.value is defined. If not, it's a multi-cell paste.
    if (e.value !== undefined) {
      // SINGLE-CELL EDIT (Your existing logic)
      const newValue = e.value || '';
      const oldValue = e.oldValue || '';
      const formula = range.getFormula();
      const isFormula = formula.startsWith('=');
      
      let changeType = 'CELL_EDIT';
      if (isFormula) {
        changeType = 'FORMULA_EDIT';
      } else if (oldValue && !newValue) {
        changeType = 'CELL_DELETE';
      } else if (!oldValue && newValue) {
        changeType = 'CELL_ADD';
      }
      
      logChangeLocal({
        timestamp: new Date(),
        sheetName: sheet.getName(),
        cellAddress: range.getA1Notation(),
        oldValue: oldValue,
        newValue: newValue,
        user: userInfo.email,
        changeType: changeType,
        formula: formula,
        details: (isFormula ? 'Formula: ' + formula : '') + ' | FP:' + userInfo.fingerprint,
        userType: userInfo.type
      });
    
    } else {
      // MULTI-CELL EDIT (e.g., paste, clear range)
      // We log this as 'CELL_EDIT' instead of 'BULK_CHANGE_DETECTED'
      logChangeLocal({
        timestamp: new Date(),
        sheetName: sheet.getName(),
        cellAddress: range.getA1Notation(),
        oldValue: '[MULTI_CELL]',
        newValue: '[MULTI_CELL]',
        user: userInfo.email,
      	  changeType: 'CELL_EDIT', // Using CELL_EDIT here
      	  formula: '',
      	  details: 'Paste or multi-cell operation in ' + range.getA1Notation() + ' | FP:' + userInfo.fingerprint,
      	  userType: userInfo.type
      });
    }
    
  } catch (error) {
    console.error('Edit tracking error:', error);
  }
}

function onChangeInstallable(e) {
  try {
    if (!e) return;
    
    const userEmail = getUserEmail(e);
    const changeType = e.changeType;
    
    if (['INSERT_ROW', 'REMOVE_ROW', 'INSERT_COLUMN', 'REMOVE_COLUMN'].includes(changeType)) {
      logChangeLocal({
        timestamp: new Date(),
        sheetName: 'DETECTED',
        cellAddress: changeType.includes('ROW') ? 'ROWS' : 'COLUMNS',
        oldValue: '',
        newValue: changeType,
        user: userEmail,
        changeType: changeType,
        formula: '',
        details: changeType + ' detected',
        userType: detectUserType(e, userEmail)
      });
    } else if (changeType === 'OTHER') {
      checkSheetChanges();
      checkStructureChanges();
    }
    
  } catch (error) {
    console.error('onChange error:', error);
  }
}

function onOpenInstallable() {
  try {
    const userEmail = getUserEmail({});
    
    logChangeLocal({
      timestamp: new Date(),
      sheetName: 'SYSTEM',
      cellAddress: 'SPREADSHEET_OPEN',
      oldValue: '',
      newValue: 'Spreadsheet opened',
      user: userEmail,
      changeType: 'SYSTEM_ACCESS',
      formula: '',
      details: 'User accessed spreadsheet',
      userType: detectUserType({}, userEmail)
    });
    
    setTimeout(() => {
      checkSheetChanges();
      checkStructureChanges();
    }, 2000);
    
  } catch (error) {
    console.error('onOpen error:', error);
  }
}

function periodicSheetCheck() {
  try {
    console.log('⏰ Running periodic check...');
    
    // 1. Flush the log cache (from Suggestion #1)
    flushLogs(); 
    
    // 2. Check for sheet/structure changes (these are lightweight)
    const lock = LockService.getScriptLock();
    if (lock.tryLock(10000)) { // Don't run if another process is changing structures
      try {
        checkSheetChanges();
        checkStructureChanges();
      } finally {
        lock.releaseLock();
      }
    } else {
      console.log('Periodic check skipped structure scan, lock busy.');
    }
    
  } catch (error) {
    console.error('Periodic check error:', error);
  }
}

// ============ UTILITY FUNCTIONS ============
function cleanupAllTriggers() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    let deletedCount = 0;
    
    triggers.forEach(trigger => {
      ScriptApp.deleteTrigger(trigger);
      deletedCount++;
    });
    
    console.log('Cleaned up', deletedCount, 'triggers');
    return 'Cleaned up ' + deletedCount + ' triggers';
    
  } catch (error) {
    console.error('Error cleaning triggers:', error);
    return 'ERROR: ' + error.toString();
  }
}

function getMultiUserStatus() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName('ChangeLog');
    const triggers = ScriptApp.getProjectTriggers();
    const props = PropertiesService.getScriptProperties();
    
    const installableTriggers = triggers.filter(t => 
      ['onEditInstallable', 'onChangeInstallable', 'periodicSheetCheck'].includes(t.getHandlerFunction())
    );
    
    const fileId = ss.getId();
    const file = DriveApp.getFileById(fileId);
    
    const status = {
      spreadsheetName: ss.getName(),
      spreadsheetUrl: ss.getUrl(),
      owner: ss.getOwner() ? ss.getOwner().getEmail() : 'Unknown',
      editors: ss.getEditors().map(e => e.getEmail()),
      viewers: ss.getViewers().map(v => v.getEmail()),
      sharingAccess: file.getSharingAccess().toString(),
      sharingPermission: file.getSharingPermission().toString(),
      changeLogExists: !!logSheet,
      totalLogEntries: logSheet ? logSheet.getLastRow() - 1 : 0,
      multiUserTracking: installableTriggers.length > 0 ? 'ACTIVE' : 'INACTIVE',
      installableTriggers: installableTriggers.length,
      totalTriggers: triggers.length,
      trackingMode: 'ADVANCED_MULTI-USER'
    };
    
    console.log('Status:', status);
    return status;
    
  } catch (error) {
    console.error('Error getting status:', error);
    return { error: error.toString() };
  }
}

function analyzeUserDetection() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName('ChangeLog');
    
    if (!logSheet || logSheet.getLastRow() <= 1) {
      return 'No log data to analyze';
    }
    
    const userData = logSheet.getRange(2, 6, logSheet.getLastRow() - 1, 5).getValues();
    
    const analysis = {
      totalEntries: userData.length,
      identifiedUsers: {},
      unidentifiedCount: 0,
      anonymousCount: 0,
      userTypes: {}
    };
    
    userData.forEach(row => {
      const user = row[0];
      const userType = row[4];
      
      if (user) {
        if (user.includes('@')) {
          analysis.identifiedUsers[user] = (analysis.identifiedUsers[user] || 0) + 1;
        } else if (user.startsWith('External_')) {
          analysis.anonymousCount++;
        } else if (user.startsWith('Unidentified_') || user === 'Unknown User') {
          analysis.unidentifiedCount++;
        }
      }
      
      if (userType) {
        analysis.userTypes[userType] = (analysis.userTypes[userType] || 0) + 1;
      }
    });
    
    analysis.identificationRate = ((analysis.totalEntries - analysis.unidentifiedCount - analysis.anonymousCount) / analysis.totalEntries * 100).toFixed(2) + '%';
    
    console.log('User Detection Analysis:', analysis);
    return analysis;
    
  } catch (error) {
    console.error('Error analyzing user detection:', error);
    return { error: error.toString() };
  }
}

function assignFingerprintToUser(fingerprint, userEmail) {
  try {
    const props = PropertiesService.getScriptProperties();
    const patterns = JSON.parse(props.getProperty('userPatterns') || '{}');
    
    if (patterns[fingerprint]) {
      patterns[fingerprint].user = userEmail;
      patterns[fingerprint].manuallyIdentified = true;
      props.setProperty('userPatterns', JSON.stringify(patterns));
      
      console.log('Fingerprint', fingerprint, 'assigned to', userEmail);
      updateHistoricalLogs(fingerprint, userEmail);
      
      return 'Successfully assigned ' + fingerprint + ' to ' + userEmail;
    } else {
      return 'Fingerprint not found: ' + fingerprint;
    }
    
  } catch (error) {
    console.error('Assignment error:', error);
    return { error: error.toString() };
  }
}

function updateHistoricalLogs(fingerprint, userEmail) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName('ChangeLog');
    
    if (!logSheet) return;
    
    const data = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 10).getValues();
    
    for (let i = 0; i < data.length; i++) {
      const details = data[i][8];
      if (details && details.includes('FP:' + fingerprint)) {
        logSheet.getRange(i + 2, 6).setValue(userEmail + ' (identified)');
        logSheet.getRange(i + 2, 10).setValue('EXTERNAL_IDENTIFIED');
      }
    }
    
    console.log('Updated historical logs for', userEmail);
    
  } catch (error) {
    console.error('Historical update error:', error);
  }
}

// ============ DASHBOARD FUNCTIONS ============
function createUserDashboard() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let dashboard = ss.getSheetByName('UserDashboard');
    
    if (!dashboard) {
      dashboard = ss.insertSheet('UserDashboard');
    }
    
    dashboard.clear();
    
    const headers = [
      ['EXTERNAL USER TRACKING DASHBOARD'],
      [''],
      ['Fingerprint', 'First Seen', 'Last Seen', 'Edit Count', 'Likely User', 'Confidence']
    ];
    
    dashboard.getRange(1, 1, headers.length, 6).setValues(headers);
    dashboard.getRange(1, 1, 1, 6).merge().setFontSize(16).setFontWeight('bold').setBackground('#4285F4').setFontColor('#FFFFFF');
    dashboard.getRange(3, 1, 1, 6).setFontWeight('bold').setBackground('#E8E8E8');
    
    const props = PropertiesService.getScriptProperties();
    const patterns = JSON.parse(props.getProperty('userPatterns') || '{}');
    
    let row = 4;
    for (let fingerprint in patterns) {
      const pattern = patterns[fingerprint];
      
      let confidence = 'Low';
      if (pattern.editCount > 10) confidence = 'Medium';
      if (pattern.editCount > 50) confidence = 'High';
      
      dashboard.getRange(row, 1, 1, 6).setValues([[
        fingerprint,
        pattern.firstSeen || '',
        pattern.lastSeen || '',
        pattern.editCount || 0,
        pattern.user,
        confidence
      ]]);
      
      const colors = { 'Low': '#FFE5E5', 'Medium': '#FFF4E5', 'High': '#E5FFE5' };
      dashboard.getRange(row, 6).setBackground(colors[confidence]);
      
      row++;
    }
    
    dashboard.autoResizeColumns(1, 6);
    
    console.log('User Dashboard created');
    return 'Dashboard created - check "UserDashboard" sheet';
    
  } catch (error) {
    console.error('Dashboard creation error:', error);
    return { error: error.toString() };
  }
}

function createProtectionDashboard() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let dashboard = ss.getSheetByName('ProtectionDashboard');
    
    if (!dashboard) {
      dashboard = ss.insertSheet('ProtectionDashboard');
    }
    
    dashboard.clear();
    
    const headers = [
      ['SHEET PROTECTION MONITORING DASHBOARD'],
      ['Last Updated: ' + new Date().toLocaleString()],
      [''],
      ['Sheet Name', 'Protection Status', 'Description', 'Can You Edit?', 'Authorized Editors', 'Range Protections']
    ];
    
    dashboard.getRange(1, 1, headers.length, 6).setValues(headers);
    dashboard.getRange(1, 1, 1, 6).merge().setFontSize(16).setFontWeight('bold').setBackground('#EA4335').setFontColor('#FFFFFF');
    dashboard.getRange(2, 1, 1, 6).merge().setFontStyle('italic');
    dashboard.getRange(4, 1, 1, 6).setFontWeight('bold').setBackground('#E8E8E8');
    
    const sheets = ss.getSheets();
    let row = 5;
    
    sheets.forEach(sheet => {
      const protectionInfo = getSheetProtectionInfo(sheet);
      
      dashboard.getRange(row, 1, 1, 6).setValues([[
        sheet.getName(),
        protectionInfo.isProtected ? 'PROTECTED' : 'UNPROTECTED',
        protectionInfo.description || 'None',
        protectionInfo.canCurrentUserEdit ? 'Yes' : 'No',
        protectionInfo.editors.join(', ') || 'Owner only',
        protectionInfo.rangeProtections > 0 ? protectionInfo.rangeProtections + ' ranges' : 'None'
      ]]);
      
      dashboard.getRange(row, 2).setBackground(protectionInfo.isProtected ? '#FFCDD2' : '#C8E6C9');
      row++;
    });
    
    dashboard.autoResizeColumns(1, 6);
    
    console.log('Protection Dashboard created');
    return 'Protection dashboard created - check "ProtectionDashboard" sheet';
    
  } catch (error) {
    console.error('Dashboard creation error:', error);
    return { error: error.toString() };
  }
}

// ============ DIAGNOSTIC FUNCTIONS ============
function DIAGNOSE_TRACKER() {
  console.log('Running diagnostic...');
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const props = PropertiesService.getScriptProperties();
    const triggers = ScriptApp.getProjectTriggers();
    
    console.log('\nSPREADSHEET INFO:');
    console.log('  Name:', ss.getName());
    console.log('  ID:', ss.getId());
    console.log('  Active user:', getUserEmail({}));
    
    console.log('\nSCRIPT PROPERTIES:');
    const allProps = props.getProperties();
    Object.keys(allProps).forEach(key => {
      const value = allProps[key];
      console.log('  ' + key + ':', value.length > 100 ? value.substring(0, 100) + '...' : value);
    });
    
    console.log('\nINSTALLED TRIGGERS:');
    if (triggers.length === 0) {
      console.log('  NO TRIGGERS FOUND! Run SETUP_ENHANCED_MULTIUSER_TRACKER() first');
    } else {
      triggers.forEach((trigger, index) => {
        console.log('  ' + (index + 1) + '. ' + trigger.getHandlerFunction() + ' - ' + trigger.getEventType());
      });
    }
    
    console.log('\nCHANGELOG SHEET:');
    const logSheet = ss.getSheetByName('ChangeLog');
    if (!logSheet) {
      console.log('  ChangeLog sheet not found!');
    } else {
      console.log('  Found -', logSheet.getLastRow(), 'rows');
    }
    
    console.log('\nDiagnostic complete!');
    
    return 'Diagnostic complete - check console logs';
    
  } catch (error) {
    console.error('Diagnostic failed:', error);
    return 'ERROR: ' + error.toString();
  }
}

function MANUAL_TEST_LOG() {
  console.log('Creating test log entry...');
  
  try {
    logChangeLocal({
      timestamp: new Date(),
      sheetName: 'TEST',
      cellAddress: 'A1',
      oldValue: 'test_old',
      newValue: 'test_new',
      user: getUserEmail({}),
      changeType: 'CELL_EDIT',
      formula: '',
      details: 'Manual test entry',
      userType: 'TESTER'
    });
    
    console.log('Test log created');
    console.log('Check ChangeLog sheet for the entry');
    
    return 'Test log created successfully!';
  } catch (error) {
    console.error('Test failed:', error);
    return 'ERROR: ' + error.toString();
  }
}

function flushLogs() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) { // Don't run if another flush is happening
    console.log('Log flush skipped, lock busy.');
    return; 
  }
  
  try {
    const cache = CacheService.getScriptCache();
    const logCacheKey = 'auditLogCache';
    const logQueueJson = cache.get(logCacheKey);
    
    if (!logQueueJson) return; // Nothing to log
    
    const logQueue = JSON.parse(logQueueJson);
    if (logQueue.length === 0) return;
    
    // Clear cache *before* writing to prevent duplicate logs on error
    cache.remove(logCacheKey);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let logSheet = ss.getSheetByName('ChangeLog');
    if (!logSheet) logSheet = initializeLogSheet(ss);
    
    const startRow = logSheet.getLastRow() + 1;
    const numRows = logQueue.length;
    const numCols = logQueue[0].length;
    
    logSheet.getRange(startRow, 1, numRows, numCols).setValues(logQueue);
    
    // Format timestamp column for new rows
    logSheet.getRange(startRow, 1, numRows, 1).setNumberFormat('yyyy-MM-dd HH:mm:ss');
    
    console.log('Log flush complete:', numRows, 'entries written.');
    
    // Auto-cleanup old logs (now efficient)
    const totalRows = logSheet.getLastRow();
    if (totalRows > CONFIG_AUDIT.maxChangeLogRows) {
      const extraRows = totalRows - CONFIG_AUDIT.maxChangeLogRows;
      logSheet.deleteRows(2, extraRows); // Delete from row 2 (after header)
      console.log('🧹 Auto-cleaned ' + extraRows + ' old log rows');
    }
    
  } catch (error) {
    console.error('Log flush error:', error);
  } finally {
    lock.releaseLock();
  }
}
