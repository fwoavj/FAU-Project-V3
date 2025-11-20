// ============================================================================
// MAIN SUBMISSION AND UI FUNCTIONS
// ============================================================================

/**
 * Main entry point for form submission
 */
function savePersonnelData(formData) {
  try {
    Logger.log('📝 Starting personnel data submission');
    Logger.log('Form data received: ' + JSON.stringify(formData, null, 2));

    // Validate form data
    const validation = validateForm_savePersonnel(formData); // Using our uniquely named validator
    if (!validation.isValid) {
      const errorMessages = validation.errors.map(e => `${e.field}: ${e.message}`).join('\n');
      return {
        success: false,
        message: '❌ Validation failed:\n\n' + errorMessages
      };
    }
    
    if (validation.warnings.length > 0) {
      const warningMessages = validation.warnings.map(w => `${w.field}: ${w.message}`).join('\n');
      Logger.log('⚠️ Warnings:\n' + warningMessages);
    }
    
    // Check for duplicates
    const dupCheck = checkDuplicatePrincipal(formData.fullName, formData.dateOfBirth);
    if (dupCheck.isDuplicate) {
      return {
        success: false,
        message: '❌ Duplicate detected:\n\n' + dupCheck.message
      };
    }
    
    // Process file uploads
    Logger.log('Processing file uploads...');
    formData = processAllFiles(formData); // This is in FileUpload.gs
    
    // Get spreadsheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.PERSONNEL_TRACKING);
    if (!sheet) {
      throw new Error('Sheet "Personnel" not found');
    }

    // Prepare data arrays
    const pIdx = SYSTEM_CONFIG.COLUMNS.PRINCIPAL._INDICES;
    const dIdx = SYSTEM_CONFIG.COLUMNS.DEPENDENT._INDICES;
    const sIdx = SYSTEM_CONFIG.COLUMNS.STAFF._INDICES;

    // Create principal base data - UPDATED TO 64 COLUMNS
    const principalBaseData = new Array(64).fill('');
    principalBaseData[pIdx.POST_STATION] = formData.postStation || '';
    principalBaseData[pIdx.FULL_NAME] = formData.fullName || '';
    principalBaseData[pIdx.RANK] = formData.rank || '';
    principalBaseData[pIdx.DESIGNATION] = formData.designation || '';
    principalBaseData[pIdx.DATE_OF_BIRTH] = formData.dateOfBirth ? parseDate_ddMMyyyy(formData.dateOfBirth) : '';
    principalBaseData[pIdx.AGE] = calculateAgeFromDateString(formData.dateOfBirth);
    principalBaseData[pIdx.SEX] = formData.sex || '';
    principalBaseData[pIdx.ASSUMPTION_DATE] = formData.assumptionDate ? parseDate_ddMMyyyy(formData.assumptionDate) : '';
    principalBaseData[pIdx.PASSPORT_NUMBER] = formData.principalPassport || '';
    principalBaseData[pIdx.PASSPORT_EXPIRATION] = formData.principalPassportExp ? parseDate_ddMMyyyy(formData.principalPassportExp) : '';
    principalBaseData[pIdx.PASSPORT_URL] = formData.principalPassportUrl || '';
    principalBaseData[pIdx.VISA_NUMBER] = formData.principalVisaNumber || '';
    principalBaseData[pIdx.VISA_EXPIRATION] = formData.principalVisaExp ? parseDate_ddMMyyyy(formData.principalVisaExp) : '';
    principalBaseData[pIdx.DIPLOMATIC_ID] = formData.principalDipId || '';
    principalBaseData[pIdx.DIPLOMATIC_ID_EXP] = formData.principalDipIdExp ? parseDate_ddMMyyyy(formData.principalDipIdExp) : '';
    principalBaseData[pIdx.DEPARTURE_DATE] = formData.principalDepartureDate ? parseDate_ddMMyyyy(formData.principalDepartureDate) : '';
    principalBaseData[pIdx.SOLO_PARENT] = formData.soloParent || 'No';
    principalBaseData[pIdx.SOLO_PARENT_URL] = formData.soloParentUrl || '';
    principalBaseData[pIdx.EXTENDED] = formData.extended || ''; 
    principalBaseData[pIdx.CURRENT_DEPARTURE_DATE] = formData.currentDepartureDate ? parseDate_ddMMyyyy(formData.currentDepartureDate) : ''; // Col 20
    principalBaseData[pIdx.EXTENSION_DETAILS] = formData.extensionDetails || ''; // Col 21

    // Create dependent data arrays
    const dependentDataList = (formData.dependents || []).map(dep => {
      const depRowArray = new Array(64).fill(''); 

      let depFullName = dep.lastName + ', ' + dep.firstName;
      if (dep.middleName) depFullName += ' ' + dep.middleName;
      if (dep.suffix) depFullName += ' ' + dep.suffix;

      depRowArray[dIdx.FULL_NAME] = depFullName;
      depRowArray[dIdx.RELATIONSHIP] = dep.relationship || '';
      depRowArray[dIdx.SEX] = dep.sex || '';
      depRowArray[dIdx.DATE_OF_BIRTH] = dep.dateOfBirth ? parseDate_ddMMyyyy(dep.dateOfBirth) : '';
      depRowArray[dIdx.AGE] = dep.age || calculateAgeFromDateString(dep.dateOfBirth);
      
      if (dep.dateOfBirth) {
        const depDOBDate = parseDate_ddMMyyyy(dep.dateOfBirth);
        if (depDOBDate) {
          const turns18 = new Date(depDOBDate);
          turns18.setFullYear(turns18.getFullYear() + 18);
          depRowArray[dIdx.TURNS_18_DATE] = turns18;
        }
      }
      
      depRowArray[dIdx.AT_POST] = dep.atPost || 'Yes';
      depRowArray[dIdx.NOTICE_OF_ARRIVAL] = dep.noticeOfArrivalDate ? parseDate_ddMMyyyy(dep.noticeOfArrivalDate) : '';
      depRowArray[dIdx.FAMILY_ALLOWANCE] = dep.receivesFamilyAllowance || 'No';
      depRowArray[dIdx.PASSPORT_NUMBER] = dep.passport || '';
      depRowArray[dIdx.PASSPORT_EXPIRATION] = dep.passportExp ? parseDate_ddMMyyyy(dep.passportExp) : '';
      depRowArray[dIdx.PASSPORT_URL] = dep.passportUrl || '';
      depRowArray[dIdx.VISA_NUMBER] = dep.visaNumber || '';
      depRowArray[dIdx.VISA_EXPIRATION] = dep.visaExp ? parseDate_ddMMyyyy(dep.visaExp) : '';
      depRowArray[dIdx.DIPLOMATIC_ID] = dep.dipId || '';
      depRowArray[dIdx.DIPLOMATIC_ID_EXP] = dep.dipIdExp ? parseDate_ddMMyyyy(dep.dipIdExp) : '';
      depRowArray[dIdx.DEPARTURE_DATE] = dep.departureDate ? parseDate_ddMMyyyy(dep.departureDate) : '';
      depRowArray[dIdx.PWD_STATUS] = dep.pwdStatus || 'No';
      depRowArray[dIdx.PWD_URL] = dep.pwdUrl || '';
      depRowArray[dIdx.APPROVAL_FAX_URL] = dep.approvalFaxUrl || ''; 
      depRowArray[dIdx.EXTENDED] = dep.extended || '';
      depRowArray[dIdx.CURRENT_DEPARTURE_DATE] = dep.currentDepartureDate ? parseDate_ddMMyyyy(dep.currentDepartureDate) : '';
      depRowArray[dIdx.EXTENSION_DETAILS] = dep.extensionDetails || ''; 
      
      return depRowArray;
    });

    // Create staff data arrays
    const staffDataList = (formData.privateStaff || []).map(staff => {
      const staffRowArray = new Array(64).fill(''); 

      let staffFullName = staff.lastName + ', ' + staff.firstName;
      if (staff.middleName) staffFullName += ' ' + staff.middleName;
      if (staff.suffix) staffFullName += ' ' + staff.suffix;

      staffRowArray[sIdx.FULL_NAME] = staffFullName; 
      staffRowArray[sIdx.SEX] = staff.sex || '';
      staffRowArray[sIdx.DATE_OF_BIRTH] = staff.dateOfBirth ? parseDate_ddMMyyyy(staff.dateOfBirth) : '';
      staffRowArray[sIdx.AGE] = staff.age || calculateAgeFromDateString(staff.dateOfBirth);
      staffRowArray[sIdx.AT_POST] = staff.atPost || 'Yes';
      staffRowArray[sIdx.ARRIVAL_DATE] = staff.arrivalDate ? parseDate_ddMMyyyy(staff.arrivalDate) : '';
      staffRowArray[sIdx.PASSPORT_NUMBER] = staff.passport || '';
      staffRowArray[sIdx.PASSPORT_EXPIRATION] = staff.passportExp ? parseDate_ddMMyyyy(staff.passportExp) : '';
      staffRowArray[sIdx.PASSPORT_URL] = staff.passportUrl || '';
      staffRowArray[sIdx.VISA_NUMBER] = staff.visaNumber || '';
      staffRowArray[sIdx.VISA_EXPIRATION] = staff.visaExp ? parseDate_ddMMyyyy(staff.visaExp) : '';
      staffRowArray[sIdx.DIPLOMATIC_ID] = staff.dipId || '';
      staffRowArray[sIdx.DIPLOMATIC_ID_EXP] = staff.dipIdExp ? parseDate_ddMMyyyy(staff.dipIdExp) : '';
      staffRowArray[sIdx.DEPARTURE_DATE] = staff.departureDate ? parseDate_ddMMyyyy(staff.departureDate) : '';
      staffRowArray[sIdx.PWD_STATUS] = staff.pwdStatus || 'No';
      staffRowArray[sIdx.PWD_URL] = staff.pwdUrl || '';
      staffRowArray[sIdx.EMERGENCY_CONTACT] = staff.emergencyContact || '';
      staffRowArray[sIdx.EXTENDED] = staff.extended || '';
      staffRowArray[sIdx.CURRENT_DEPARTURE_DATE] = staff.currentDepartureDate ? parseDate_ddMMyyyy(staff.currentDepartureDate) : '';
      staffRowArray[sIdx.EXTENSION_DETAILS] = staff.extensionDetails || '';
      
      return staffRowArray;
    });
    
    // Combine and write rows
    const rowsToWrite = [];
    const maxRows = Math.max(1, dependentDataList.length, staffDataList.length);
    const principalTargetRow = sheet.getLastRow() + 1;

    for (let i = 0; i < maxRows; i++) {
      const finalRow = [...principalBaseData];
      const depData = dependentDataList[i];
      const staffData = staffDataList[i];

      if (depData) {
        for (let j = dIdx.FULL_NAME; j <= dIdx.EXTENSION_DETAILS; j++) {
          if (depData[j] !== undefined) {
            finalRow[j] = depData[j];
          }
        }
      }
      
      if (staffData) {
        for (let j = sIdx.FULL_NAME; j <= sIdx.EXTENSION_DETAILS; j++) {
          if (staffData[j] !== undefined) {
            finalRow[j] = staffData[j];
          }
        }
      }

      rowsToWrite.push(finalRow);
    }

    // Write all rows at once
    if (rowsToWrite.length > 0) {
      sheet.getRange(principalTargetRow, 1, rowsToWrite.length, 64).setValues(rowsToWrite);
      Logger.log(`✅ ${rowsToWrite.length} row(s) for ${formData.fullName} recorded starting row ${principalTargetRow}`);

      // ================= AUDIT LOG INTEGRATION =================
      try {
        if (typeof logChangeLocal === 'function') {
          logChangeLocal({
            timestamp: new Date(),
            sheetName: SYSTEM_CONFIG.SHEETS.PERSONNEL_TRACKING,
            cellAddress: 'ROW_' + principalTargetRow + ':' + (principalTargetRow + rowsToWrite.length - 1),
            oldValue: '',
            newValue: 'New Record Created',
            user: Session.getActiveUser().getEmail(),
            changeType: 'UI_ENTRY_SUBMISSION',
            formula: '',
            details: `Added Principal: ${formData.fullName} with ${rowsToWrite.length - 1} dependents/staff`,
            userType: 'EDITOR'
          });
          flushLogs(); // Force write to ChangeLog
        }
      } catch (e) {
        console.error("Logging failed for savePersonnelData: " + e);
      }
      // =========================================================
    }

    SpreadsheetApp.flush();
    return {
      success: true,
      message: `✅ Personnel data recorded successfully starting row ${principalTargetRow}`,
      row: principalTargetRow
    };
  } catch (error) {
    Logger.log('❌ Error in savePersonnelData: ' + error);
    Logger.log('Error Stack: ' + error.stack);
    return {
      success: false,
      message: 'Failed to save data: ' + error.message
    };
  }
}

/**
 * SAVES updated data from the UpdateDetailsForm.
 * This function finds the existing rows and overwrites them.
 */
function saveUpdatedPersonnelData(formData) {
  try {
    Logger.log('📝 Starting personnel data update for: ' + formData.fullName);
    
    // 1. Validation
    const validation = validateForm_savePersonnel(formData);
    if (!validation.isValid) {
      const errorMessages = validation.errors.map(e => `${e.field}: ${e.message}`).join('\n');
      return {
        success: false,
        message: '❌ Validation failed:\n\n' + errorMessages
      };
    }
    
    // 2. Process File Uploads
    Logger.log('Processing file uploads...');
    formData = processAllFiles(formData); // From FileUpload.gs
    
    // 3. Get Sheet and Find Rows
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.PERSONNEL_TRACKING);
    if (!sheet) {
      throw new Error('Sheet "Personnel" not found');
    }

    // Find all existing rows for this principal
    const principalRows = findPrincipalRows(sheet, formData.originalFullName || formData.fullName);
    if (!principalRows || principalRows.length === 0) {
      throw new Error(`Could not find existing record for principal: ${formData.fullName}`);
    }
    
    const principalStartRow = principalRows[0]; 

    // 4. Prepare Data Arrays
    const pIdx = SYSTEM_CONFIG.COLUMNS.PRINCIPAL._INDICES;
    const dIdx = SYSTEM_CONFIG.COLUMNS.DEPENDENT._INDICES;
    const sIdx = SYSTEM_CONFIG.COLUMNS.STAFF._INDICES;

    // Create principal base data (64 columns)
    const principalBaseData = new Array(64).fill('');
    principalBaseData[pIdx.POST_STATION] = formData.postStation || '';
    principalBaseData[pIdx.FULL_NAME] = formData.fullName || '';
    principalBaseData[pIdx.RANK] = formData.rank || '';
    principalBaseData[pIdx.DESIGNATION] = formData.designation || '';
    principalBaseData[pIdx.DATE_OF_BIRTH] = formData.dateOfBirth ? parseDate_ddMMyyyy(formData.dateOfBirth) : '';
    principalBaseData[pIdx.AGE] = calculateAgeFromDateString(formData.dateOfBirth);
    principalBaseData[pIdx.SEX] = formData.sex || '';
    principalBaseData[pIdx.ASSUMPTION_DATE] = formData.assumptionDate ? parseDate_ddMMyyyy(formData.assumptionDate) : '';
    principalBaseData[pIdx.PASSPORT_NUMBER] = formData.principalPassport || '';
    principalBaseData[pIdx.PASSPORT_EXPIRATION] = formData.principalPassportExp ? parseDate_ddMMyyyy(formData.principalPassportExp) : '';
    principalBaseData[pIdx.PASSPORT_URL] = formData.principalPassportUrl || ''; 
    principalBaseData[pIdx.VISA_NUMBER] = formData.principalVisaNumber || '';
    principalBaseData[pIdx.VISA_EXPIRATION] = formData.principalVisaExp ? parseDate_ddMMyyyy(formData.principalVisaExp) : '';
    principalBaseData[pIdx.DIPLOMATIC_ID] = formData.principalDipId || '';
    principalBaseData[pIdx.DIPLOMATIC_ID_EXP] = formData.principalDipIdExp ? parseDate_ddMMyyyy(formData.principalDipIdExp) : '';
    principalBaseData[pIdx.DEPARTURE_DATE] = formData.principalDepartureDate ? parseDate_ddMMyyyy(formData.principalDepartureDate) : '';
    principalBaseData[pIdx.SOLO_PARENT] = formData.soloParent || 'No';
    principalBaseData[pIdx.SOLO_PARENT_URL] = formData.soloParentUrl || '';
    principalBaseData[pIdx.EXTENDED] = formData.extended || '';
    principalBaseData[pIdx.CURRENT_DEPARTURE_DATE] = formData.currentDepartureDate ? parseDate_ddMMyyyy(formData.currentDepartureDate) : '';
    principalBaseData[pIdx.EXTENSION_DETAILS] = formData.extensionDetails || '';

    // Create dependent data arrays
    const dependentDataList = (formData.dependents || []).map(dep => {
      const depRowArray = new Array(64).fill(''); 
      let depFullName = dep.lastName + ', ' + dep.firstName;
      if (dep.middleName) depFullName += ' ' + dep.middleName;
      if (dep.suffix) depFullName += ' ' + dep.suffix;
      
      depRowArray[dIdx.FULL_NAME] = depFullName;
      depRowArray[dIdx.RELATIONSHIP] = dep.relationship || '';
      depRowArray[dIdx.SEX] = dep.sex || '';
      depRowArray[dIdx.DATE_OF_BIRTH] = dep.dateOfBirth ? parseDate_ddMMyyyy(dep.dateOfBirth) : '';
      depRowArray[dIdx.AGE] = dep.age || calculateAgeFromDateString(dep.dateOfBirth);
      
      if (dep.dateOfBirth) {
        const depDOBDate = parseDate_ddMMyyyy(dep.dateOfBirth);
        if (depDOBDate) {
          const turns18 = new Date(depDOBDate);
          turns18.setFullYear(turns18.getFullYear() + 18);
          depRowArray[dIdx.TURNS_18_DATE] = turns18;
        }
      }
      
      depRowArray[dIdx.AT_POST] = dep.atPost || 'Yes';
      depRowArray[dIdx.NOTICE_OF_ARRIVAL] = dep.noticeOfArrivalDate ? parseDate_ddMMyyyy(dep.noticeOfArrivalDate) : '';
      depRowArray[dIdx.FAMILY_ALLOWANCE] = dep.receivesFamilyAllowance || 'No';
      depRowArray[dIdx.PASSPORT_NUMBER] = dep.passport || '';
      depRowArray[dIdx.PASSPORT_EXPIRATION] = dep.passportExp ? parseDate_ddMMyyyy(dep.passportExp) : '';
      depRowArray[dIdx.PASSPORT_URL] = dep.passportUrl || '';
      depRowArray[dIdx.VISA_NUMBER] = dep.visaNumber || '';
      depRowArray[dIdx.VISA_EXPIRATION] = dep.visaExp ? parseDate_ddMMyyyy(dep.visaExp) : '';
      depRowArray[dIdx.DIPLOMATIC_ID] = dep.dipId || '';
      depRowArray[dIdx.DIPLOMATIC_ID_EXP] = dep.dipIdExp ? parseDate_ddMMyyyy(dep.dipIdExp) : '';
      depRowArray[dIdx.DEPARTURE_DATE] = dep.departureDate ? parseDate_ddMMyyyy(dep.departureDate) : '';
      depRowArray[dIdx.PWD_STATUS] = dep.pwdStatus || 'No';
      depRowArray[dIdx.PWD_URL] = dep.pwdUrl || '';
      depRowArray[dIdx.APPROVAL_FAX_URL] = dep.approvalFaxUrl || '';
      depRowArray[dIdx.EXTENDED] = dep.extended || '';
      depRowArray[dIdx.CURRENT_DEPARTURE_DATE] = dep.currentDepartureDate ? parseDate_ddMMyyyy(dep.currentDepartureDate) : '';
      depRowArray[dIdx.EXTENSION_DETAILS] = dep.extensionDetails || '';
      return depRowArray;
    });

    // Create staff data arrays
    const staffDataList = (formData.privateStaff || []).map(staff => {
      const staffRowArray = new Array(64).fill('');
      let staffFullName = staff.lastName + ', ' + staff.firstName;
      if (staff.middleName) staffFullName += ' ' + staff.middleName;
      if (staff.suffix) staffFullName += ' ' + staff.suffix;
      
      staffRowArray[sIdx.FULL_NAME] = staffFullName;
      staffRowArray[sIdx.SEX] = staff.sex || '';
      staffRowArray[sIdx.DATE_OF_BIRTH] = staff.dateOfBirth ? parseDate_ddMMyyyy(staff.dateOfBirth) : '';
      staffRowArray[sIdx.AGE] = staff.age || calculateAgeFromDateString(staff.dateOfBirth);
      staffRowArray[sIdx.AT_POST] = staff.atPost || 'Yes';
      staffRowArray[sIdx.ARRIVAL_DATE] = staff.arrivalDate ? parseDate_ddMMyyyy(staff.arrivalDate) : '';
      staffRowArray[sIdx.PASSPORT_NUMBER] = staff.passport || '';
      staffRowArray[sIdx.PASSPORT_EXPIRATION] = staff.passportExp ? parseDate_ddMMyyyy(staff.passportExp) : '';
      staffRowArray[sIdx.PASSPORT_URL] = staff.passportUrl || '';
      staffRowArray[sIdx.VISA_NUMBER] = staff.visaNumber || '';
      staffRowArray[sIdx.VISA_EXPIRATION] = staff.visaExp ? parseDate_ddMMyyyy(staff.visaExp) : '';
      staffRowArray[sIdx.DIPLOMATIC_ID] = staff.dipId || '';
      staffRowArray[sIdx.DIPLOMATIC_ID_EXP] = staff.dipIdExp ? parseDate_ddMMyyyy(staff.dipIdExp) : '';
      staffRowArray[sIdx.DEPARTURE_DATE] = staff.departureDate ? parseDate_ddMMyyyy(staff.departureDate) : '';
      staffRowArray[sIdx.PWD_STATUS] = staff.pwdStatus || 'No';
      staffRowArray[sIdx.PWD_URL] = staff.pwdUrl || '';
      staffRowArray[sIdx.EMERGENCY_CONTACT] = staff.emergencyContact || '';
      staffRowArray[sIdx.EXTENDED] = staff.extended || '';
      staffRowArray[sIdx.CURRENT_DEPARTURE_DATE] = staff.currentDepartureDate ? parseDate_ddMMyyyy(staff.currentDepartureDate) : '';
      staffRowArray[sIdx.EXTENSION_DETAILS] = staff.extensionDetails || '';
      return staffRowArray;
    });
    
    // 5. Combine and Write Rows
    const rowsToWrite = [];
    const maxRows = Math.max(1, dependentDataList.length, staffDataList.length);

    for (let i = 0; i < maxRows; i++) {
      const finalRow = [...principalBaseData];
      const depData = dependentDataList[i];
      const staffData = staffDataList[i];

      if (depData) {
        for (let j = dIdx.FULL_NAME; j <= dIdx.EXTENSION_DETAILS; j++) {
          if (depData[j] !== undefined) finalRow[j] = depData[j];
        }
      }
      
      if (staffData) {
        for (let j = sIdx.FULL_NAME; j <= sIdx.EXTENSION_DETAILS; j++) {
          if (staffData[j] !== undefined) finalRow[j] = staffData[j];
        }
      }
      rowsToWrite.push(finalRow);
    }

    // 6. Overwrite Existing Data
    
    // Clear old rows first, in case the new data has fewer rows
    if (rowsToWrite.length < principalRows.length) {
      const startClear = principalStartRow + rowsToWrite.length;
      const numToClear = principalRows.length - rowsToWrite.length;
      sheet.getRange(startClear, 1, numToClear, 64).clearContent();
      Logger.log(`Cleared ${numToClear} old row(s)`);
    }
    
    // Write the new data
    if (rowsToWrite.length > 0) {
      sheet.getRange(principalStartRow, 1, rowsToWrite.length, 64).setValues(rowsToWrite);
      Logger.log(`✅ ${rowsToWrite.length} row(s) for ${formData.fullName} UPDATED starting row ${principalStartRow}`);

      // ================= AUDIT LOG INTEGRATION =================
      try {
        if (typeof logChangeLocal === 'function') {
          logChangeLocal({
            timestamp: new Date(),
            sheetName: SYSTEM_CONFIG.SHEETS.PERSONNEL_TRACKING,
            cellAddress: 'ROW_' + principalStartRow + ':' + (principalStartRow + rowsToWrite.length - 1),
            oldValue: 'Previous Data',
            newValue: 'Updated Data',
            user: Session.getActiveUser().getEmail(),
            changeType: 'UI_UPDATE_SUBMISSION',
            formula: '',
            details: `Updated details for Principal: ${formData.fullName}`,
            userType: 'EDITOR'
          });
          flushLogs();
        }
      } catch (e) {
        console.error("Logging failed for saveUpdatedPersonnelData: " + e);
      }
      // =========================================================
    }

    clearDataCache();
    SpreadsheetApp.flush();
    return {
      success: true,
      message: `✅ Personnel data for ${formData.fullName} updated successfully.`,
      row: principalStartRow
    };
    
  } catch (error) {
    Logger.log('❌ Error in saveUpdatedPersonnelData: ' + error);
    Logger.log('Error Stack: ' + error.stack);
    return {
      success: false,
      message: 'Failed to save updates: ' + error.message
    };
  }
}

/**
 * Checks for a duplicate principal in the tracking sheet.
 */
function checkDuplicatePrincipal(fullName, dateOfBirth) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.PERSONNEL_TRACKING);
    if (!sheet) return { isDuplicate: false };
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 3) return { isDuplicate: false };

    // Get column numbers from CONFIG
    const nameCol = SYSTEM_CONFIG.COLUMNS.PRINCIPAL.FULL_NAME;
    const dobCol = SYSTEM_CONFIG.COLUMNS.PRINCIPAL.DATE_OF_BIRTH;
    
    const names = sheet.getRange(3, nameCol, lastRow - 2, 1).getValues();
    const dobs = sheet.getRange(3, dobCol, lastRow - 2, 1).getValues();
    
    // Use parseDate_ddMMyyyy from this file and convert to a comparable string
    const dobString = dateOfBirth ? parseDate_ddMMyyyy(dateOfBirth).toDateString() : '';

    for (let i = 0; i < names.length; i++) {
      if (names[i][0] === fullName) {
        // Only compare dates if the name matches
        const sheetDob = dobs[i][0] ? new Date(dobs[i][0]).toDateString() : '';
        if (sheetDob === dobString) {
          return {
            isDuplicate: true,
            message: `Principal "${fullName}" with DOB "${dobString}" already exists in row ${i + 3}.`
          };
        }
      }
    }
    return { isDuplicate: false };
  } catch (error) {
    Logger.log("Error checking duplicate: " + error);
    // Don't block submission if check fails, just log it.
    return { isDuplicate: false, message: "Error checking for duplicates: " + error.message }; 
  }
}

/**
 * Validates the flat formData object from savePersonnelData
 */
function validateForm_savePersonnel(formData) {
  const errors = [];
  const warnings = [];
  const config = SYSTEM_CONFIG.VALIDATION;

  function addError(field, message) {
    errors.push({ field, message });
  }

  // --- Principal Validation ---
  if (!formData.fullName) addError('Principal', 'Full Name is required. (Did you select a principal?)');
  if (!formData.rank) addError('Principal', 'Rank is required.');
  if (!formData.sex) addError('Principal', 'Sex is required.');
  if (!formData.assumptionDate) addError('Principal', 'Date of Assumption is required.');
  if (!formData.principalPassport) addError('Principal', 'Passport No. is required.');
  if (!formData.principalPassportExp) addError('Principal', 'Passport Expiration is required.');

  // --- Dependents Validation ---
  if (formData.dependents && formData.dependents.length > 0) {
    formData.dependents.forEach((dep, index) => {
      const depNum = index + 1;
      if (!dep.lastName || !dep.firstName) addError(`Dependent ${depNum}`, 'First and Last Name are required.');
      if (!dep.relationship) addError(`Dependent ${depNum}`, 'Relationship is required.');
      if (!dep.sex) addError(`Dependent ${depNum}`, 'Sex is required.');
      if (!dep.dateOfBirth) addError(`Dependent ${depNum}`, 'Date of Birth is required.');
      if (!dep.atPost) addError(`Dependent ${depNum}`, '"At Post" status is required.');
      if (!dep.receivesFamilyAllowance) addError(`Dependent ${depNum}`, '"Receives Family Allowance" status is required.');
      if (!dep.passport) addError(`Dependent ${depNum}`, 'Passport No. is required.');
      if (!dep.passportExp) addError(`Dependent ${depNum}`, 'Passport Expiration is required.');
      if (dep.pwdStatus === "") addError(`Dependent ${depNum}`, 'PWD Status is required.'); 
    });
  }

  // --- Private Staff Validation ---
  if (formData.privateStaff && formData.privateStaff.length > 0) {
    formData.privateStaff.forEach((staff, index) => {
      const staffNum = index + 1;
      if (!staff.lastName || !staff.firstName) addError(`Staff ${staffNum}`, 'First and Last Name are required.');
      if (!staff.sex) addError(`Staff ${staffNum}`, 'Sex is required.');
      if (!staff.dateOfBirth) {
        addError(`Staff ${staffNum}`, 'Date of Birth is required.');
      } else {
        // Check age requirement
        const age = calculateAgeFromDateString(staff.dateOfBirth); 
        if (age < config.MIN_AGE_STAFF) {
          addError(`Staff ${staffNum}`, `Staff member must be at least ${config.MIN_AGE_STAFF} years old. Current age is ${age}.`);
        }
      }
      if (!staff.atPost) addError(`Staff ${staffNum}`, '"At Post" status is required.');
      if (!staff.passport) addError(`Staff ${staffNum}`, 'Passport No. is required.');
      if (!staff.passportExp) addError(`Staff ${staffNum}`, 'Passport Expiration is required.');
      if (staff.pwdStatus === "") addError(`Staff ${staffNum}`, 'PWD Status is required.'); 
    });
  }

  return {
    isValid: errors.length === 0,
    errors: errors,
    warnings: warnings
  };
}

/**
 * Calculate age from date string
 */
function calculateAgeFromDateString(dateString) {
  if (!dateString) return '';
  try {
    const birthDate = new Date(dateString);
    if (isNaN(birthDate.getTime())) return '';
    
    const today = new Date();
    let age = today.getFullYear() - birthDate.getFullYear();
    const monthDiff = today.getMonth() - birthDate.getMonth();
    if (monthDiff < 0 || (monthDiff === 0 && today.getDate() < birthDate.getDate())) {
      age--;
    }
    
    return age >= 0 ? age : '';
  } catch (error) {
    return '';
  }
}

/**
 * Parse date in dd/mm/yyyy format OR yyyy-mm-dd
 */
function parseDate_ddMMyyyy(dateString) {
  if (!dateString || typeof dateString !== 'string') {
    return null;
  }

  // Try yyyy-mm-dd first (standard HTML date)
  if (dateString.includes('-')) {
    const parts = dateString.split('-');
    if (parts.length === 3 && parts[0].length === 4) {
      const date = new Date(dateString);
      if (!isNaN(date.getTime())) {
        return date;
      }
    }
  }

  // Try dd/mm/yyyy
  const parts = dateString.split(/[\/\-]/);
  if (parts.length === 3) {
    const day = parseInt(parts[0], 10);
    const month = parseInt(parts[1], 10); 
    const year = parseInt(parts[2], 10);
    if (!isNaN(day) && !isNaN(month) && !isNaN(year) && year > 1900 && month >= 1 && month <= 12) {
       return new Date(year, month - 1, day);
    }
  }

  // Fallback attempt
  const fallbackDate = new Date(dateString);
  if (!isNaN(fallbackDate.getTime())) {
     Logger.log("Warning: parseDate_ddMMyyyy used fallback for input: " + dateString);
    return fallbackDate;
  }

  Logger.log("Error: parseDate_ddMMyyyy failed to parse input: " + dateString);
  return null;
}

// ============================================================================
// SHEET CREATION FUNCTIONS - UPDATED FOR 64 COLUMNS
// ============================================================================

/**
 * Creates the "Personnel" sheet with 64 columns (A-BL)
 */
function createTrackingSheet(spreadsheet = null) {
  try {
    const ss = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
    const trackingSheet = ss.insertSheet('Personnel');

    // Ensure we have exactly 64 columns
    const maxColumns = trackingSheet.getMaxColumns();
    if (maxColumns < 64) {
      trackingSheet.insertColumnsAfter(maxColumns, 64 - maxColumns);
    } else if (maxColumns > 64) {
      trackingSheet.deleteColumns(65, maxColumns - 64);
    }

    // Group headers (Row 1)
    trackingSheet.getRange(1, 1, 1, 21).merge().setValue('Principals')
      .setFontSize(12).setFontWeight('bold').setBackground('#a4c2f4')
      .setFontColor('black').setHorizontalAlignment('center');
    trackingSheet.getRange(1, 22, 1, 23).merge().setValue('Dependents')
      .setFontSize(12).setFontWeight('bold').setBackground('#93CCEA')
      .setFontColor('black').setHorizontalAlignment('center');
    trackingSheet.getRange(1, 45, 1, 20).merge().setValue('Private Staff')
      .setFontSize(12).setFontWeight('bold').setBackground('#f4cccc')
      .setFontColor('black').setHorizontalAlignment('center');
    
    // Column Headers (Row 2)
    const headers = [
      // Principal (Cols 1-21)
      'Post/Station', 'Principal Full Name', 'Rank', 'Designation', 'Date of Birth', 'Age', 'Sex',
      'Assumption Date', 'Principal Passport', 'Passport Expiration', 'Passport URL', 'Visa Number', 
      'Visa Expiration', 'Diplomatic/Consular ID', 'ID Expiration', 'Departure Date', 'Solo Parent', 
      'Solo Parent URL', 'Extended?', 'New Departure Date', 'Extension Details',
      
      // Dependent (Cols 22-44)
      'Dependent Name', 'Relationship', 'Sex', 'Date of Birth', 'Age', 'Turns 18 Date', 'At Post',
      'Notice of Arrival Date', 'Receives Family Allowance', 'Dependent Passport', 'Dependent Passport Exp', 
      'Passport URL', 'Visa Number', 'Visa Expiration', 'Diplomatic/Consular ID', 'ID Expiration', 
      'Departure Date', 'PWD Status', 'PWD URL', 'Approval Fax URL', 'Extended?', 'New Departure Date', 
      'Extension Details',
            
      // Staff (Cols 45-64)
      'Staff Name', 'Sex', 'Date of Birth', 'Age', 'At Post', 'Arrival Date',
      'Staff Passport', 'Staff Passport Exp', 'Passport URL', 'Visa Number', 'Visa Expiration', 
      'Diplomatic/Consular ID', 'ID Expiration', 'Departure Date', 'PWD Status', 'PWD URL', 
      'Emergency Contact', 'Extended?', 'New Departure Date', 'Extension Details'
    ];
    
    // Write all 64 headers to Row 2
    trackingSheet.getRange(2, 1, 1, 64).setValues([headers]);
  
    // Format Header Row 2
    trackingSheet.getRange(2, 1, 1, 64).setBackground('#a4c2f4').setFontColor('black').setFontWeight('bold')
      .setWrap(true).setHorizontalAlignment('center');
    
    trackingSheet.getRange(2, 22, 1, 23).setBackground('#93ccea'); // Dependent headers
    trackingSheet.getRange(2, 45, 1, 20).setBackground('#f4cccc'); // Staff headers
  
    // Column Widths
    const columnWidths = [
      // Principal (1-21)
      200, 150, 180, 150, 120, 60, 80, 120, 120, 120, 250, 120, 120, 120, 120, 120, 100, 250, 100, 140, 300,
      
      // Dependent (22-44)
      200, 120, 80, 120, 60, 120, 80, 120, 120, 120, 120, 250, 120, 120, 120, 120, 120, 80, 250, 250, 100, 140, 300,
            
      // Staff (45-64)
      200, 80, 120, 60, 80, 120, 120, 120, 250, 120, 120, 120, 120, 120, 80, 250, 150, 100, 140, 300
    ];
    
    // Set all 64 column widths
    columnWidths.forEach((width, index) => {
      trackingSheet.setColumnWidth(index + 1, width);
    });
    
    // Data Validations
    const extendedValidation = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Yes', 'No'])
      .setAllowInvalid(false)
      .build();
    
    trackingSheet.getRange('S3:S').setDataValidation(extendedValidation); // Principal Extended
    trackingSheet.getRange('AP3:AP').setDataValidation(extendedValidation); // Dependent Extended
    trackingSheet.getRange('BJ3:BJ').setDataValidation(extendedValidation); // Staff Extended (NOW COL 62/BJ)
    
    // Date Formats
    trackingSheet.getRange('T3:T').setNumberFormat('yyyy-mm-dd');
    trackingSheet.getRange('AQ3:AQ').setNumberFormat('yyyy-mm-dd');
    trackingSheet.getRange('BK3:BK').setNumberFormat('yyyy-mm-dd'); // Staff New Departure (NOW COL 63/BK)
    
    // Freeze Panes
    trackingSheet.setFrozenRows(2);
    trackingSheet.setFrozenColumns(2);
    
    console.log('✅ Personnel sheet created successfully with 64 columns');
    return trackingSheet;
    
  } catch (error) {
    console.error('❌ Error creating tracking sheet:', error);
    throw error;
  }
}

function createPrincipalsSheet(spreadsheet = null) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const principalsSheet = ss.insertSheet('Principals List');

    const headers = ['Post/Station', 'Full Name', 'Date of Birth'];
    principalsSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    const headerRange = principalsSheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#4a86e8').setFontColor('white').setFontWeight('bold')
      .setHorizontalAlignment('center');

    principalsSheet.setColumnWidth(1, 200);
    principalsSheet.setColumnWidth(2, 200);
    principalsSheet.setColumnWidth(3, 120);
    principalsSheet.setFrozenRows(1);
    
    console.log('✅ Principals List sheet created successfully');
    return principalsSheet;
    
  } catch (error) {
    console.error('❌ Error creating principals sheet:', error);
    throw error;
  }
}

/**
 * Creates the "Allowance Removal Log" sheet if it doesn't exist
 */
function createAllowanceLogSheet(spreadsheet = null) {
  const ss = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.ALLOWANCE_LOG);
  
  if (!sheet) {
    sheet = ss.insertSheet(SYSTEM_CONFIG.SHEETS.ALLOWANCE_LOG);
    const headers = ['Date Logged', 'Principal Name', 'Dependent Name', 'Relationship', 'Date of Birth', 'Age', 'Status'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#e6b8af'); // Light red header
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 150); // Date
    sheet.setColumnWidth(2, 200); // Principal
    sheet.setColumnWidth(3, 200); // Dependent
    console.log('✅ Created Allowance Removal Log sheet');
  }
  return sheet;
}

// ============================================================================
// UI FUNCTIONS
// ============================================================================

function showEntryForm() {
  const html = HtmlService.createHtmlOutputFromFile('EntryForm')
    .setWidth(SYSTEM_CONFIG.UI.MODAL_WIDTH)
    .setHeight(800)
    .setTitle('Personnel Entry Form');
  SpreadsheetApp.getUi().showModalDialog(html, 'Personnel Entry Form');
}

function showUpdateForm() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('UpdateForm')
      .setWidth(900)
      .setHeight(800)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    SpreadsheetApp.getUi().showModalDialog(html, '✏️ Update Personnel');
  } catch (error) {
    console.error('Error showing update form:', error);
    SpreadsheetApp.getUi().alert('Error opening form: ' + error.toString());
  }
}

function showTrackingSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let trackingSheet = ss.getSheetByName('Personnel');
  
    if (!trackingSheet) {
      trackingSheet = createTrackingSheet(ss);
    }
  
    trackingSheet.activate();
  
    if (trackingSheet.getLastRow() <= 2) {
      SpreadsheetApp.getUi().alert('📋 No personnel records found. Use "Open Entry Form" to add data.');
    }
  } catch (error) {
    console.error('Error showing tracking sheet:', error);
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

function showPrincipalsSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let principalsSheet = ss.getSheetByName('Principals List');
    if (!principalsSheet) {
      principalsSheet = createPrincipalsSheet(ss);
    }
  
    principalsSheet.activate();
  } catch (error) {
    console.error('Error showing principals sheet:', error);
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

function openAttendanceTracker() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('attendanceForm')
      .setWidth(1200)
      .setHeight(800)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    SpreadsheetApp.getUi().showModalDialog(html, '📊 Attendance Tracker');
  } catch (error) {
    console.error('Error showing attendance tracker:', error);
    SpreadsheetApp.getUi().alert('Error opening attendance tracker: ' + error.toString());
  }
}

function openArchiveViewer() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('archive')
      .setWidth(1200)
      .setHeight(800)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    SpreadsheetApp.getUi().showModalDialog(html, '📦 Archived Withdrawals');
  } catch (error) {
    console.error('Error showing archive viewer:', error);
    SpreadsheetApp.getUi().alert('Error opening archive: ' + error.toString());
  }
}

/**
 * NEW: Shows the new form for updating all details of existing personnel.
 */
function showUpdateDetailsForm() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('UpdateDetailsForm')
      .setWidth(1200)
      .setHeight(800)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    SpreadsheetApp.getUi().showModalDialog(html, '✏️ Update Personnel Details');
  } catch (error) {
    console.error('Error showing update details form:', error);
    SpreadsheetApp.getUi().alert('Error opening form: ' + error.toString());
  }
}

function updatePrincipal() {
  SpreadsheetApp.getUi().alert('Use "Manage Principals" to view and edit the Principals List sheet directly.');
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Menu')
    .addItem('Entry Form', 'showEntryForm')
    .addItem('Status of Residency', 'openAttendanceTracker')
    .addItem('Update Personnel Details', 'showUpdateDetailsForm')
    .addItem('Extension Form', 'showUpdateForm')
    .addSeparator()
    .addItem('Run Manual Alerts', 'showStartupAlerts')
    .addToUi();
  
  // This will run when the sheet is opened
  showStartupAlerts();
}

// ============================================================================
// ALERTS FUNCTIONS
// ============================================================================

/**
 * Main function to run all startup alerts.
 * Combines alerts into one pop-up.
 */
function showStartupAlerts() {
  let alertMessages = [];
  
  // 1. Run Allowance Check (Auto-log 18+ dependents)
  const newlyLoggedAdults = checkAndLogAdultDependents();
  if (newlyLoggedAdults.length > 0) {
    alertMessages.push('🚫 DEPENDENTS FLAGGED FOR ALLOWANCE REMOVAL (18+):\n(Added to "Allowance Removal Log")\n- ' + newlyLoggedAdults.join('\n- '));
  }
  
  // 2. Get Attendance Alerts
  const attendanceAlerts = getUpcomingAttendanceAlerts();
  if (attendanceAlerts.length > 0) {
    alertMessages.push('⏳ ATTENDANCE DUE (IN 3 DAYS OR LESS):\n' + attendanceAlerts.join('\n'));
  }

  // 3. Get Birthday Alerts (Just for info, "Turned 18 Today")
  const birthdayAlerts = getTurning18TodayAlerts();
  if (birthdayAlerts.length > 0) {
    alertMessages.push('🎂 TURNED 18 TODAY:\n' + birthdayAlerts.join('\n'));
  }
  
  // 4. Show combined alert
  if (alertMessages.length > 0) {
    SpreadsheetApp.getUi().alert('🔔 System Alerts', alertMessages.join('\n\n________________________\n\n'), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Helper function to find attendance alerts (3 days or less remaining)
 */
function getUpcomingAttendanceAlerts() {
  try {
    const data = loadAttendanceData(); 
    const alerts = [];
    
    data.forEach(row => {
      if (row.windowInfo && row.windowInfo.isOpen === true && row.windowInfo.daysRemaining <= 3) {
        const daysText = row.windowInfo.daysRemaining === 1 ? 'day' : 'days';
        alerts.push(`- ${row.name} (${row.windowInfo.daysRemaining} ${daysText} left)`);
      }
    });
    
    return alerts;
    
  } catch (e) {
    Logger.log("Error getting attendance alerts: " + e);
    return [];
  }
}

/**
 * Helper function to find 18th birthday alerts for TODAY
 */
function getTurning18TodayAlerts() {
  try {
    const dependents = getTurning18Today(); 
    const alerts = [];
    
    dependents.forEach(dep => {
      alerts.push(`- ${dep.dependentName} (Principal: ${dep.principalName})`);
    });
    
    return alerts;

  } catch (e) {
    Logger.log("Error getting birthday alerts: " + e);
    return [];
  }
}

/**
 * Scans for dependents who have reached their "Turns 18 Date".
 * Logs them if not already logged.
 */
function checkAndLogAdultDependents() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let logSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.ALLOWANCE_LOG);
    
    if (!logSheet) {
        logSheet = createAllowanceLogSheet(ss);
    }
    
    const lastRow = logSheet.getLastRow();
    const loggedPairs = new Set();
    if (lastRow > 1) {
      const data = logSheet.getRange(2, 2, lastRow - 1, 2).getValues();
      data.forEach(row => {
        if (row[0] && row[1]) {
          loggedPairs.add(`${row[0]}_${row[1]}`);
        }
      });
    }
    
    const allPersonnel = getAllPersonnelDataCached();
    const newlyLogged = [];
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    
    allPersonnel.forEach(person => {
      person.dependents.forEach(dep => {
        if (dep.turns18Date) {
          const turns18Date = new Date(dep.turns18Date);
          turns18Date.setHours(0, 0, 0, 0);

          if (turns18Date <= today) {
            const uniqueKey = `${person.principal.fullName}_${dep.fullName}`;
            
            if (!loggedPairs.has(uniqueKey)) {
              const currentAge = calculateAgeFromDateString(dep.dateOfBirth);

              logSheet.appendRow([
                new Date(),
                person.principal.fullName,
                dep.fullName,
                dep.relationship,
                formatDateForClient(dep.dateOfBirth),
                currentAge,
                'Flagged for Removal'
              ]);
              
              newlyLogged.push(`${dep.fullName} (Turned 18 on ${formatDateForClient(dep.turns18Date)})`);
              loggedPairs.add(uniqueKey);
            }
          }
        }
      });
    });
    
    return newlyLogged;
    
  } catch (e) {
    Logger.log('Error in checkAndLogAdultDependents: ' + e);
    return [];
  }
}

// ============================================================================
// PRINCIPALS LIST FUNCTIONS (for dropdowns)
// ============================================================================

function getPrincipalsList() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.PRINCIPALS_LIST);
    if (!sheet) return [];

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    
    const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues(); 

    const principals = data
      .filter(row => row[1] && row[1].toString().trim() !== "")
      .map(row => {
        let isoDob = "";
        const rawDob = row[2];
        if (rawDob instanceof Date) {
          isoDob = Utilities.formatDate(rawDob, Session.getScriptTimeZone(), "yyyy-MM-dd");
        } else if (typeof rawDob === "string" && rawDob.trim() !== "") {
          const parsed = parseDate_ddMMyyyy(rawDob);
          if(parsed) {
            isoDob = Utilities.formatDate(parsed, Session.getScriptTimeZone(), "yyyy-MM-dd");
          }
        }
        return {
          postStation: (row[0] || "").toString().trim(),
          fullName: (row[1] || "").toString().trim(),
          dateOfBirth: isoDob
        };
      });
    
    return principals;
  } catch (err) {
    Logger.log("❌ getPrincipalsList error: " + err);
    return [];
  }
}

function getAllPrincipals() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.PERSONNEL_TRACKING);
    if (!sheet) return [];
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 3) return [];

    const data = sheet.getRange(3, 1, lastRow - 2, 2).getValues();
    const principals = [];
    const seenPrincipals = new Set();
    
    for (let i = 0; i < data.length; i++) {
      const post = data[i][0];
      const principalName = data[i][1];
      
      if (principalName && principalName !== '' && !seenPrincipals.has(principalName)) {
        seenPrincipals.add(principalName);
        principals.push({
          fullName: principalName,
          post: post || ''
        });
      }
    }
    return principals;
  } catch (error) {
    console.error('Error getting principals:', error);
    return [];
  }
}

// ============================================================================
// DATA ACCESS HELPERS (Duplicates of DATA_ACCESS.gs for safety/convenience)
// ============================================================================

function findPrincipalRows(sheet, principalName) {
  const nameColumn = SYSTEM_CONFIG.COLUMNS.PRINCIPAL.FULL_NAME;
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];
  
  const data = sheet.getRange(3, nameColumn, lastRow - 2, 1).getValues();
  const matchingRows = [];
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === principalName) {
      matchingRows.push(i + 3);
    }
  }
  return matchingRows;
}

function findPrincipalRow(sheet, principalName) {
  const rows = findPrincipalRows(sheet, principalName);
  return rows.length > 0 ? rows[0] : null;
}
