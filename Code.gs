function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('Process Excellence — Team Hub')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1.0, user-scalable=no');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getESTNow() {
  const estString = Utilities.formatDate(new Date(), "America/New_York", "yyyy-MM-dd'T'HH:mm:ss");
  return new Date(estString);
}

function parseInputDateToSheetToken(dateString) {
  if (!dateString) return null;
  const dateParts = dateString.split('-');
  return {
    monthIndex: parseInt(dateParts[1], 10) - 1,
    token: parseInt(dateParts[2], 10) + "-" + ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"][parseInt(dateParts[1], 10) - 1]
  };
}

function getMonthConfigByIndex(monthIndex) {
  const blocks = { 
    4: { startRow: 158, monthLabel: "May" }, 
    5: { startRow: 171, monthLabel: "Jun" },
    6: { startRow: 184, monthLabel: "Jul" },
    7: { startRow: 197, monthLabel: "Aug" },
    8: { startRow: 210, monthLabel: "Sep" }
  };
  return blocks[monthIndex] || blocks[5];
}

function fetchMonthLeaveData(monthIndex) {
  if (isDemoUser()) return { success: true, calendarLeaveMap: {}, monthIndex: monthIndex };

  try {
    const trackerSheetId = "1w1qRTNg53VUw9Z7lik_tFx5m05ssZwTyn1-9ZvT2hhA";
    const sheetAtt = SpreadsheetApp.openById(trackerSheetId).getSheetByName("1A_Attendance Tracker");
    const config = getMonthConfigByIndex(monthIndex);
    if (!config) throw new Error("Month not configured in backend.");

    const fullDataMatrix = sheetAtt.getRange(config.startRow, 2, 13, 32).getValues();
    const headerDays = fullDataMatrix[0];
    const calendarLeaveMap = {};
    
    const stringifiedHeaders = headerDays.map(day => {
      if (day instanceof Date) return day.getDate() + "-" + ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"][day.getMonth()];
      return String(day).trim();
    });
    
    for (let i = 2; i < fullDataMatrix.length; i++) {
      const name = fullDataMatrix[i][0];
      if (!name || String(name).trim() === "") continue;
      
      for (let c = 1; c < stringifiedHeaders.length; c++) {
        const dayToken = stringifiedHeaders[c];
        const statusValue = String(fullDataMatrix[i][c] || "").toUpperCase().trim();
        if (statusValue.startsWith('VL') || statusValue.startsWith('HDVL') || statusValue.startsWith('BL')) {
          if (!calendarLeaveMap[dayToken]) calendarLeaveMap[dayToken] = [];
          calendarLeaveMap[dayToken].push({ name: name, type: statusValue, rowIndex: config.startRow + i });
        }
      }
    }
    return { success: true, calendarLeaveMap: calendarLeaveMap, monthIndex: monthIndex };
  } catch(err) {
    return { success: false, message: err.message };
  }
}

function isFuzzyNameMatch(nameA, nameB) {
  const normalize = (str) => String(str).toLowerCase().replace(/[^a-z]/g, "");
  if (normalize(nameA) === normalize(nameB)) return true;

  const clean = (str) => String(str).toLowerCase().replace(/[^a-z0-9\s]/g, " ").split(/\s+/).filter(word => word.length > 2);
  const wordsA = clean(nameA);
  const wordsB = clean(nameB);

  if (wordsA.length === 0 || wordsB.length === 0) return false;

  let matchCount = 0;
  wordsA.forEach(wa => {
    if (wordsB.includes(wa)) matchCount++;
  });

  return matchCount >= 2 || (matchCount === 1 && (wordsA.length === 1 || wordsB.length === 1));
}


function isDemoUser() {
  const email = Session.getActiveUser().getEmail().toLowerCase();
  return email === "" || !email.endsWith("@openlane.com");
}

/**
 * LOGS NATIVE USER RECOGNITION AND SESSION TELEMETRY DATA
 */
function logUserSessionTelemetry(actionTypeContext) {
  try {
    const targetUserEmail = Session.getActiveUser().getEmail() || "anonymous@openlane.com";
    if (isDemoUser()) return { success: true, verifiedUser: targetUserEmail + " (DEMO MODE)" };

    const logSheetId = "1NNAr_1Z9vgynr0GtKqCuPY8UwS7MqkaW0uiaU5P64ig";
    const logSheet = SpreadsheetApp.openById(logSheetId).getSheetByName("Site_sessionLog");
    const timestampNow = new Date();
    const estFormatted = Utilities.formatDate(timestampNow, "America/New_York", "yyyy-MM-dd HH:mm:ss") + " EST";
    
    logSheet.appendRow([estFormatted, targetUserEmail, actionTypeContext]);
    return { success: true, verifiedUser: targetUserEmail };
  } catch(err) {
    console.error("Session telemetry loop dropped: " + err.message);
    return { success: false, error: err.message };
  }
}

/**
 * GATHERS ATTENDANCE, DEMOGRAPHICS AND RESOLVES ACTIVE USER EMAIL CHANNELS
 */
function getTodayAttendanceData() {
  const authMetrics = logUserSessionTelemetry("Application Consolidated Console Initialization Launch Pass");
  const estNow = getESTNow();
  const currentMonthNum = estNow.getMonth();
  const config = getMonthConfigByIndex(currentMonthNum);
  const currentDayStr = estNow.getDate() + "-" + config.monthLabel;
  const todayDateString = estNow.toLocaleDateString('en-US', { weekday: 'long', month: 'long', day: 'numeric', year: 'numeric', timeZone: 'America/New_York' }) + " (EST)";
  const inputDefault = estNow.getFullYear() + "-" + String(estNow.getMonth() + 1).padStart(2, '0') + "-" + String(estNow.getDate()).padStart(2, '0');

  if (isDemoUser()) {
    return {
      roster: [
        {name: "Myra Perry", initials: "MP", rawStatus: "P", rowIndex: 5},
        {name: "Mildred Irish Lopez", initials: "MI", rawStatus: "HDVL 1st ", rowIndex: 6}
      ],
      calendarLeaveMap: { [currentDayStr]: [{name: "Mildred Irish Lopez", type: "HDVL 1st ", rowIndex: 6}] },
      demographicsMap: {},
      currentUserVerifiedEmail: authMetrics.verifiedUser,
      dateString: todayDateString + " [DEMO MODE]",
      targetDayStr: currentDayStr,
      currentYear: estNow.getFullYear(),
      currentMonthNum: estNow.getMonth(),
      inputDefaultDate: inputDefault
    };
  }

  const trackerSheetId = "1w1qRTNg53VUw9Z7lik_tFx5m05ssZwTyn1-9ZvT2hhA";
  const ss = SpreadsheetApp.openById(trackerSheetId);
  const sheetAtt = ss.getSheetByName("1A_Attendance Tracker");
  
  const roster = [];
  const calendarLeaveMap = {};
  
  const blocks = { 
    4: { startRow: 158, monthLabel: "May" }, 
    5: { startRow: 171, monthLabel: "Jun" },
    6: { startRow: 184, monthLabel: "Jul" },
    7: { startRow: 197, monthLabel: "Aug" },
    8: { startRow: 210, monthLabel: "Sep" }
  };

  const monthsToProcess = [currentMonthNum];
  let nextMonth = currentMonthNum + 1;
  if (nextMonth > 11) nextMonth = 0; 
  if (blocks[nextMonth]) monthsToProcess.push(nextMonth);

  monthsToProcess.forEach(mIndex => {
    const conf = blocks[mIndex];
    if (!conf) return;
    
    const fullDataRange = sheetAtt.getRange(conf.startRow, 2, 13, 32);
    const fullDataMatrix = fullDataRange.getValues();
    const fullNotesMatrix = fullDataRange.getNotes();
    const headerDays = fullDataMatrix[0];
    
    const stringifiedHeaders = headerDays.map(day => {
      if (day instanceof Date) return day.getDate() + "-" + ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"][day.getMonth()];
      return String(day).trim();
    });

    if (mIndex === currentMonthNum) {
      let todayColIndex = stringifiedHeaders.findIndex(day => day.toLowerCase() === currentDayStr.toLowerCase());
      if (todayColIndex === -1) {
        const dayNum = currentDayStr.split('-')[0];
        const monthStr = currentDayStr.split('-')[1].toLowerCase().substring(0, 3);
        todayColIndex = stringifiedHeaders.findIndex(day => {
          const str = day.toLowerCase();
          return str.includes(dayNum + "-") && str.includes(monthStr);
        });
      }
      if (todayColIndex === -1) todayColIndex = 1; 
      
      for (let i = 2; i < fullDataMatrix.length; i++) {
        const name = fullDataMatrix[i][0];
        if (!name || String(name).trim() === "") continue;
        
        let todayStatus = String(fullDataMatrix[i][todayColIndex] || "").trim();
        let cellNote = String(fullNotesMatrix[i][todayColIndex] || "").trim();

        // Only process note if status is validly present
        if (todayStatus.toUpperCase() === "HDVL" && cellNote.includes("Segment: ")) {
          todayStatus = cellNote.replace("Segment: ", "").trim();
        }

        roster.push({
          name: name,
          initials: name.split(' ').map(n => n[0]).join('').substring(0, 2).toUpperCase(),
          rawStatus: todayStatus || "",
          rowIndex: conf.startRow + i
        });
      }
    }

    for (let i = 2; i < fullDataMatrix.length; i++) {
      const name = fullDataMatrix[i][0];
      if (!name || String(name).trim() === "") continue;
      
      for (let c = 1; c < stringifiedHeaders.length; c++) {
        const dayToken = stringifiedHeaders[c];
        let statusValue = String(fullDataMatrix[i][c] || "").trim();
        let cellNote = String(fullNotesMatrix[i][c] || "").trim();

        // STRICT CHECK: Skip genuinely empty cells
        if (!statusValue || statusValue === "") continue;

        let upperStatus = statusValue.toUpperCase();

        if (upperStatus === "HDVL" && cellNote.includes("Segment: ")) {
          statusValue = cellNote.replace("Segment: ", "").trim();
          upperStatus = statusValue.toUpperCase();
        }

        // Only register if the cell explicitly contains a leave code
        if (upperStatus.startsWith('VL') || upperStatus.startsWith('HDVL') || upperStatus.startsWith('HD VL') || upperStatus.startsWith('BL')) {
          if (!calendarLeaveMap[dayToken]) calendarLeaveMap[dayToken] = [];
          calendarLeaveMap[dayToken].push({ name: name, type: statusValue, rowIndex: conf.startRow + i });
        }
      }
    }
  });
  
  const sheetDemo = ss.getSheetByName("1_Demographics");
  const rawDemoData = sheetDemo.getRange("B3:P100").getValues(); 
  const rawHeaders = rawDemoData[0]; 
  const demographicsMap = {};
  const demoEntries = [];
  
  for (let r = 2; r < rawDemoData.length; r++) {
    const demoName = String(rawDemoData[r][1]).trim();
    if (!demoName || demoName === "" || demoName === "undefined" || demoName === "Name") continue;
    
    const profileDetails = [];
    for (let c = 0; c < rawHeaders.length; c++) {
      let headerName = String(rawHeaders[c]).trim();
      let rawVal = rawDemoData[r][c];
      if (rawVal instanceof Date) rawVal = rawVal.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
      if (!headerName || headerName === "Sub" || headerName === "undefined") continue;
      profileDetails.push({ label: headerName, value: rawVal !== "" ? String(rawVal).trim() : "Not Listed" });
    }
    demoEntries.push({ sheetName: demoName, details: profileDetails });
  }
  
  roster.forEach(member => {
    const matchedEntry = demoEntries.find(entry => isFuzzyNameMatch(member.name, entry.sheetName));
    demographicsMap[member.name] = matchedEntry ? matchedEntry.details : null;
  });
  
  return {
    roster: roster,
    calendarLeaveMap: calendarLeaveMap,
    demographicsMap: demographicsMap,
    currentUserVerifiedEmail: authMetrics.verifiedUser,
    dateString: todayDateString,
    targetDayStr: currentDayStr,
    currentYear: estNow.getFullYear(),
    currentMonthNum: estNow.getMonth(),
    inputDefaultDate: inputDefault
  };
}
/**
 * Updates attendance tracker status in Google Sheet with Data Validation Fallback.
 */
/**
 * Updates attendance tracker status in Google Sheet with exact data validation compliance.
 */
function updateStatus(frontendRowIndex, rawStatusCode, specifiedInputDate, employeeName) {
  let targetDateToken = "";
  let monthIndex = getESTNow().getMonth();
  
  if (specifiedInputDate) {
    const parsed = parseInputDateToSheetToken(specifiedInputDate);
    monthIndex = parsed.monthIndex;
    targetDateToken = parsed.token;
  } else {
    const config = getMonthConfigByIndex(monthIndex);
    targetDateToken = getESTNow().getDate() + "-" + config.monthLabel;
  }

  if (isDemoUser()) {
    return { success: true, savedValue: rawStatusCode, savedDate: targetDateToken + " (Prototype)" };
  }

  const sheetId = "1w1qRTNg53VUw9Z7lik_tFx5m05ssZwTyn1-9ZvT2hhA";
  const sheet = SpreadsheetApp.openById(sheetId).getSheetByName("1A_Attendance Tracker");
  const config = getMonthConfigByIndex(monthIndex);
  
  let finalRowIndex = frontendRowIndex;
  if (employeeName) {
    const nameRange = sheet.getRange(config.startRow, 2, 15, 1).getValues(); 
    for (let r = 0; r < nameRange.length; r++) {
      const sheetName = String(nameRange[r][0]).trim();
      if (sheetName && isFuzzyNameMatch(sheetName, employeeName)) {
        finalRowIndex = config.startRow + r;
        break;
      }
    }
  }

  const headerDays = sheet.getRange(config.startRow, 2, 1, 32).getValues()[0];
  const stringifiedHeaders = headerDays.map(day => {
    if (day instanceof Date) return day.getDate() + "-" + ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"][day.getMonth()];
    return String(day).trim();
  });
  
  let targetColIndex = stringifiedHeaders.findIndex(day => day.toLowerCase() === targetDateToken.toLowerCase());
  
  if (targetColIndex === -1) {
    const dayNum = targetDateToken.split('-')[0];
    const monthStr = targetDateToken.split('-')[1].toLowerCase().substring(0, 3); 
    targetColIndex = stringifiedHeaders.findIndex(day => {
      const str = day.toLowerCase();
      return str.includes(dayNum + "-") && str.includes(monthStr);
    });
  }

  if (targetColIndex === -1) {
    throw new Error(`Target date out of bounds. Searched for [${targetDateToken}] on Row ${config.startRow}.`);
  }
  
  const targetCell = sheet.getRange(finalRowIndex, 2 + targetColIndex);
  
  // Directly set the raw status string matching Google Sheet Data Validation
  targetCell.setValue(rawStatusCode);

  return { success: true, savedValue: rawStatusCode, savedDate: targetDateToken };
}
/**
 * Processes Multi-Date Leave Filings while automatically excluding Weekends (Sat/Sun).
 */
function updateStatusBatch(frontendRowIndex, rawStatusCode, startDateStr, endDateStr, employeeName) {
  if (!startDateStr) throw new Error("Start date is required.");
  
  // If no end date provided, treat it as a single-day filing
  const endInputStr = endDateStr && endDateStr.trim() !== "" ? endDateStr : startDateStr;
  
  const startParts = startDateStr.split('-');
  const endParts = endInputStr.split('-');
  
  let currDate = new Date(parseInt(startParts[0],10), parseInt(startParts[1],10)-1, parseInt(startParts[2],10));
  const endDate = new Date(parseInt(endParts[0],10), parseInt(endParts[1],10)-1, parseInt(endParts[2],10));
  
  if (currDate > endDate) {
    throw new Error("Start Date cannot be after End Date.");
  }

  const updatedDates = [];

  while (currDate <= endDate) {
    const dayOfWeek = currDate.getDay();
    
    // Automatically skip Saturday (6) and Sunday (0)
    if (dayOfWeek !== 0 && dayOfWeek !== 6) {
      const year = currDate.getFullYear();
      const month = String(currDate.getMonth() + 1).padStart(2, '0');
      const day = String(currDate.getDate()).padStart(2, '0');
      const formattedInputDate = `${year}-${month}-${day}`;
      
      updateStatus(frontendRowIndex, rawStatusCode, formattedInputDate, employeeName);
      updatedDates.push(formattedInputDate);
    }
    
    // Increment to the next calendar day
    currDate.setDate(currDate.getDate() + 1);
  }

  if (updatedDates.length === 0) {
    throw new Error("Selected date range only contains weekend rest days. No leave was filed.");
  }

  return { success: true, count: updatedDates.length, savedValue: rawStatusCode };
}

/**
 * INTELLIGENT HANDOFF PARSER LOGIC MATRIX
 */
function analyzeTaskHandoffMatrix(employeeResourceName, requestedTargetDateString, leaveCacheMap) {
  try {
    if (!employeeResourceName || !requestedTargetDateString) return [];
    
    const dataSheetId = "1NNAr_1Z9vgynr0GtKqCuPY8UwS7MqkaW0uiaU5P64ig";
    const handoffSheet = SpreadsheetApp.openById(dataSheetId).getSheetByName("Handoffs");
    const rawData = handoffSheet.getRange("A2:D200").getValues(); 
    
    const dateParts = requestedTargetDateString.split('-');
    const targetDateObj = new Date(parseInt(dateParts[0],10), parseInt(dateParts[1],10)-1, parseInt(dateParts[2],10));
    const dayOfWeekLabels = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
    const targetDayNameString = dayOfWeekLabels[targetDateObj.getDay()];
    const cacheDayToken = targetDateObj.getDate() + "-" + ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"][targetDateObj.getMonth()];
    
    // Grab exact names of everyone on leave today
    const leavesOnTargetDate = leaveCacheMap[cacheDayToken] || [];
    const agentsOnLeave = leavesOnTargetDate.map(l => l.name); 
    const handoffsMatchedArray = [];
    
    for (let i = 0; i < rawData.length; i++) {
      let currentResponsible = String(rawData[i][0]).trim();
      let taskName = String(rawData[i][1]).trim();
      let backupPOCRaw = String(rawData[i][2]).trim();
      let frequency = String(rawData[i][3]).trim().toLowerCase();
      
      if (!taskName) continue;
      
      let isMainPOC = isFuzzyNameMatch(currentResponsible, employeeResourceName);
      
      // Split comma-separated backups into a real array to check them individually
      let backupListArray = backupPOCRaw.split(/[\n,]+/).map(p => p.trim()).filter(p => p !== "");
      let isBackupPOC = backupListArray.some(b => isFuzzyNameMatch(b, employeeResourceName));
      
      if (!isMainPOC && !isBackupPOC) continue;
      
      let isHandoffRequired = false;
      
      if (isMainPOC) {
        if (frequency.includes("daily") || frequency === "") {
          isHandoffRequired = true;
        } else {
          const variants = { "monday": ["mon"], "tuesday": ["tue"], "wednesday": ["wed"], "thursday": ["thu"], "friday": ["fri"] }[targetDayNameString.toLowerCase()] || [targetDayNameString.toLowerCase()];
          if (variants.some(v => frequency.includes(v))) isHandoffRequired = true;
        }
      } else if (isBackupPOC && frequency.includes("back up")) {
        let mainPOCOnLeave = agentsOnLeave.some(leaveName => isFuzzyNameMatch(leaveName, currentResponsible));
        if (mainPOCOnLeave) isHandoffRequired = true;
      }

      if (isHandoffRequired) {
        let requiresReassignment = false;
        let finalPOCString = backupPOCRaw;

        if (isMainPOC) {
          let activeBackups = [];
          
          // Loop through every backup and filter out the ones on leave
          backupListArray.forEach(b => {
             let isOut = agentsOnLeave.some(leaveName => isFuzzyNameMatch(leaveName, b));
             if (!isOut) activeBackups.push(b); // Keep them if they aren't on leave
          });

          if (activeBackups.length === 0) {
             // Everyone designated is on leave -> Trigger Modal
             requiresReassignment = true;
             finalPOCString = backupPOCRaw; 
          } else {
             // At least one person is here -> Only add them to the email draft
             finalPOCString = activeBackups.join(", "); 
          }
        } else if (isBackupPOC) {
          // If the resource is the Backup, check if the Main POC is also out
          let mainPOCOnLeave = agentsOnLeave.some(leaveName => isFuzzyNameMatch(leaveName, currentResponsible));
          if (mainPOCOnLeave) requiresReassignment = true;
        }

        handoffsMatchedArray.push({
          task: taskName,
          poc: finalPOCString !== "" ? finalPOCString : "Unassigned / Team Backup Support",
          frequency: frequency || "Daily",
          targetDayMatched: targetDayNameString,
          requiresReassignment: requiresReassignment,
          conflictDate: cacheDayToken
        });
      }
    }
    return handoffsMatchedArray;
  } catch(err) {
    throw new Error("Task handoff calculation pipeline failure: " + err.message);
  }
}

/**
 * Deletes a filed leave from the tracker.
 */
function deleteOfficialLeave(frontendRowIndex, dateToken, employeeName) {
  if (isDemoUser()) {
    return { success: true, clearedDate: dateToken + " (Prototype)" };
  }

  const sheetId = "1w1qRTNg53VUw9Z7lik_tFx5m05ssZwTyn1-9ZvT2hhA";
  const sheet = SpreadsheetApp.openById(sheetId).getSheetByName("1A_Attendance Tracker");
  const tokenParts = dateToken.split('-');
  const monthLabel = tokenParts[1];
  const monthsShorthand = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const monthIndex = monthsShorthand.findIndex(m => m.toLowerCase() === monthLabel.toLowerCase().substring(0,3));
  
  const config = getMonthConfigByIndex(monthIndex);
  
  let finalRowIndex = frontendRowIndex;
  if (employeeName) {
    const nameRange = sheet.getRange(config.startRow, 2, 15, 1).getValues(); 
    for (let r = 0; r < nameRange.length; r++) {
      const sheetName = String(nameRange[r][0]).trim();
      if (sheetName && isFuzzyNameMatch(sheetName, employeeName)) {
        finalRowIndex = config.startRow + r;
        break;
      }
    }
  }

  const headerDays = sheet.getRange(config.startRow, 2, 1, 32).getValues()[0];
  
  const stringifiedHeaders = headerDays.map(day => {
    if (day instanceof Date) return day.getDate() + "-" + ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"][day.getMonth()];
    return String(day).trim();
  });
  
  let targetColIndex = stringifiedHeaders.findIndex(day => day.toLowerCase() === dateToken.toLowerCase());
  
  if (targetColIndex === -1) {
    const dayNum = dateToken.split('-')[0];
    const monthStr = monthLabel.toLowerCase().substring(0, 3); 
    targetColIndex = stringifiedHeaders.findIndex(day => {
      const str = day.toLowerCase();
      return str.includes(dayNum + "-") && str.includes(monthStr);
    });
  }

  if (targetColIndex === -1) {
    throw new Error(`Target date out of bounds. Searched for [${dateToken}] on Row ${config.startRow}. Found headers: ${stringifiedHeaders.slice(1, 5).join(", ")}...`);
  }
  
  sheet.getRange(finalRowIndex, 2 + targetColIndex).setValue("");
  return { success: true, clearedDate: dateToken };
}

function fetchHandoffDataForUI(employeeName, dateStringsJSON, leaveCacheJSON) {
  try {
    const dateStringsArray = JSON.parse(dateStringsJSON);
    const leaveCacheMap = JSON.parse(leaveCacheJSON);
    let combinedTasks = [];
    
    dateStringsArray.forEach(dateStr => {
      const tasksForDay = analyzeTaskHandoffMatrix(employeeName, dateStr, leaveCacheMap);
      combinedTasks = combinedTasks.concat(tasksForDay);
    });
    
    const uniqueTasksMap = new Map();
    combinedTasks.forEach(taskObj => {
      if (uniqueTasksMap.has(taskObj.task)) {
         if (taskObj.requiresReassignment) uniqueTasksMap.get(taskObj.task).requiresReassignment = true;
      } else {
        uniqueTasksMap.set(taskObj.task, taskObj);
      }
    });
    
    return {
      success: true,
      data: Array.from(uniqueTasksMap.values()),
      signature: employeeName + "\nProcess Excellence Operations Central"
    };
  } catch (err) {
    return { success: false, data: [], signature: "" };
  }
}

function createEditedGmailDraft(employeeName, customSubj, customBody) {
  try {
    const formattedHtmlBody = "<div style='font-family: Arial, sans-serif; font-size: 13px; color: #222;'>" + 
                              customBody.replace(/\n/g, '<br>') + 
                              "</div>";
    
    const draftNode = GmailApp.createDraft("", customSubj, customBody, { htmlBody: formattedHtmlBody });
    
    return {
      success: true,
      draftUrl: "https://mail.google.com/mail/u/0/#drafts?compose=" + draftNode.getId()
    };
  } catch(err) {
    return { success: false, message: err.message };
  }
}

function authenticateUser() {
  try {
    const email = Session.getActiveUser().getEmail().toLowerCase();
    const isAuthorized = email === "" || email.endsWith("@openlane.com") || email.endsWith("@gmail.com");
    
    if (isAuthorized) {
      logUserSessionTelemetry("System Login: Authorized access granted.");
    } else {
      logUserSessionTelemetry("System Login: Blocked unauthorized domain attempt (" + email + ")");
    }
    
    return {
      authorized: isAuthorized,
      emailAttempt: email || "Anonymous Test Account"
    };
    
  } catch (err) {
    throw new Error("Authentication Engine Error: " + err.message);
  }
}


function submitUserFeedback(text) {
  if (isDemoUser()) {
    return { success: true };
  }
  return { success: true };
}

function getTaskAllocationData() {
  try {
    const dataSheetId = "1NNAr_1Z9vgynr0GtKqCuPY8UwS7MqkaW0uiaU5P64ig";
    const ss = SpreadsheetApp.openById(dataSheetId);
    
    // --- DIRECTORY PARSING ---
    const directorySheet = ss.getSheetByName("Project Directory");
    const dirData = directorySheet.getRange("A2:C200").getValues();
    let trainedAgentsClean = {};

    for (let i = 0; i < dirData.length; i++) {
      let mainProject = String(dirData[i][0]).trim();
      let subCat = String(dirData[i][1]).trim();
      let agentsRaw = String(dirData[i][2]).trim();

      if (!mainProject || !agentsRaw) continue;
      if (!trainedAgentsClean[mainProject]) trainedAgentsClean[mainProject] = [];

      let agentList = agentsRaw.split(/[\n,]+/).map(p => p.trim()).filter(p => p !== "");
      agentList.forEach(a => {
        let displayAgent = a;
        if (subCat && mainProject !== "BEACON ASSOCIATION" && mainProject !== "STORAGE FEE") {
          displayAgent = `${a} (${subCat})`;
        }
        if (!trainedAgentsClean[mainProject].includes(displayAgent)) {
           trainedAgentsClean[mainProject].push(displayAgent);
        }
      });
    }

    // --- STAFF AVERAGE HOURS PARSING (Columns F-K on Handoffs Tab) ---
    const handoffSheet = ss.getSheetByName("Handoffs");
    const hoursData = handoffSheet.getRange("F2:K100").getValues();
    const days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Weekly Avg"];
    const structuredChartData = { "Monday": [], "Tuesday": [], "Wednesday": [], "Thursday": [], "Friday": [], "Weekly Avg": [] };
    
    for (let i = 0; i < hoursData.length; i++) {
      let person = String(hoursData[i][0]).trim();
      if (!person) continue;

      let weeklyTotal = 0;
      let hasTasks = false;
      let dailyVals = [
        parseFloat(hoursData[i][1]) || 0,
        parseFloat(hoursData[i][2]) || 0,
        parseFloat(hoursData[i][3]) || 0,
        parseFloat(hoursData[i][4]) || 0,
        parseFloat(hoursData[i][5]) || 0 
      ];

      for(let d = 0; d < 5; d++) {
        let val = dailyVals[d];
        weeklyTotal += val;
        if (val > 0) hasTasks = true;
        structuredChartData[days[d]].push({ name: person, hours: parseFloat(val.toFixed(2)) });
      }

      if (hasTasks) {
        structuredChartData["Weekly Avg"].push({ name: person, hours: parseFloat((weeklyTotal / 5).toFixed(2)) });
      }
    }

    days.forEach(d => structuredChartData[d].sort((a, b) => b.hours - a.hours));

    return { success: true, chartData: structuredChartData, trainedAgents: trainedAgentsClean };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

// --- START SOP LIBRARY INTEGRATION ---
function getSOPDataForWebApp() {
  try {
    const sheetId = "1NNAr_1Z9vgynr0GtKqCuPY8UwS7MqkaW0uiaU5P64ig"; 
    const sheet = SpreadsheetApp.openById(sheetId).getSheetByName("SOPs");
    
    if (!sheet) throw new Error("SOPs tab not found in spreadsheet 1NNAr_1Z9vgynr0GtKqCuPY8UwS7MqkaW0uiaU5P64ig.");

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, data: [] }; 

    const rawData = sheet.getRange(2, 1, lastRow - 1, 8).getDisplayValues();
    const formattedData = [];

    for (let i = 0; i < rawData.length; i++) {
      const row = rawData[i];
      const title = row[0];
      const fileUrl = row[1];
      
      if (!title || !fileUrl) continue;

      let docId = "";
      try {
        const match = fileUrl.match(/\/d\/(.*?)\//);
        docId = match ? match[1] : "";
      } catch(e) {}

      let parsedTopics = [];
      if (row[6] && String(row[6]).trim() !== "") {
        try {
          parsedTopics = JSON.parse(row[6]);
        } catch(e) {
          parsedTopics = [row[6]]; 
        }
      }

      formattedData.push({
        title: title,
        url: fileUrl,
        author: row[2] || "Unknown",
        createdDate: row[3] || "N/A",
        lastUpdatedBy: row[4] || "N/A",
        effectiveDate: row[5] || "N/A",
        topics: Array.isArray(parsedTopics) ? parsedTopics : [],
        searchIndex: row[7] || "",
        previewIframe: docId ? `https://docs.google.com/document/d/${docId}/preview` : fileUrl
      });
    }
    
    return { success: true, data: formattedData };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function runFullSync() {
  const folderId = "1Y8WWr-96ys8Q18H8s7H78IHzdR3WYpio";
  const sheetId = "1NNAr_1Z9vgynr0GtKqCuPY8UwS7MqkaW0uiaU5P64ig"; 
  const sheetName = "SOPs";

  const sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);
  if (!sheet) throw new Error("SOPs tab not found in the target spreadsheet.");

  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFilesByType(MimeType.GOOGLE_DOCS); 
  
  const updatedDataMatrix = [];

  while (files.hasNext()) {
    const file = files.next();
    const docId = file.getId();
    
    const title = file.getName();
    const fileUrl = file.getUrl();
    const createdDate = file.getDateCreated();
    const lastUpdated = file.getLastUpdated();
    
    const owner = file.getOwner() ? file.getOwner().getName() : "Unknown";

    let topics = [];
    let fullTextIndex = "";

    try {
      const doc = DocumentApp.openById(docId);
      const body = doc.getBody();
      const text = body.getText();
      
      const paragraphs = body.getParagraphs();
      paragraphs.forEach(p => {
        const heading = p.getHeading();
        if (heading === DocumentApp.ParagraphHeading.HEADING1 || heading === DocumentApp.ParagraphHeading.HEADING2) {
          const txt = p.getText().trim();
          if (txt) topics.push(txt);
        }
      });
      
      fullTextIndex = text.toLowerCase().replace(/\s+/g, ' ').trim();
      
    } catch (e) {
      console.warn("Failed to scrape document body: " + fileUrl + " | Error: " + e.message);
    }

    updatedDataMatrix.push([
      title, 
      fileUrl, 
      owner, 
      Utilities.formatDate(createdDate, Session.getScriptTimeZone(), "yyyy-MM-dd"),
      Utilities.formatDate(lastUpdated, Session.getScriptTimeZone(), "yyyy-MM-dd"),
      "N/A", 
      JSON.stringify(topics.slice(0, 10)),
      fullTextIndex 
    ]);
  }

  if (updatedDataMatrix.length > 0) {
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, 8).clearContent();
    }
    sheet.getRange(2, 1, updatedDataMatrix.length, 8).setValues(updatedDataMatrix);
  }
}

function getKPIDataForWebApp() {
  try {
    const sheetId = "1NNAr_1Z9vgynr0GtKqCuPY8UwS7MqkaW0uiaU5P64ig"; 
    const ss = SpreadsheetApp.openById(sheetId);
    let sheet = ss.getSheetByName("KPIs");
    
    if (!sheet) {
      sheet = ss.insertSheet("KPIs");
      sheet.appendRow(["Title", "URL", "Author", "Created Date", "Last Updated", "Effective Date", "Topics", "Search Index"]);
      runKPISync(); 
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, data: [] };

    const rawData = sheet.getRange(2, 1, lastRow - 1, 8).getDisplayValues();
    const formattedData = [];

    for (let i = 0; i < rawData.length; i++) {
      const row = rawData[i];
      const title = row[0];
      const fileUrl = row[1];
      
      if (!title || !fileUrl) continue;

      let docId = "";
      try {
        const match = fileUrl.match(/\/d\/(.*?)\//) || fileUrl.match(/id=(.*?)(&|$)/);
        docId = match ? match[1] : "";
      } catch(e) {}

      let parsedTopics = [];
      if (row[6] && String(row[6]).trim() !== "") {
        try {
          parsedTopics = JSON.parse(row[6]);
        } catch(e) {
          parsedTopics = [row[6]];
        }
      }

      let previewIframe = fileUrl;
      if (docId) {
        if (fileUrl.includes("document")) {
          previewIframe = `https://docs.google.com/document/d/${docId}/preview`;
        } else if (fileUrl.includes("spreadsheets")) {
          previewIframe = `https://docs.google.com/spreadsheets/d/${docId}/preview`;
        } else {
          previewIframe = `https://drive.google.com/file/d/${docId}/preview`;
        }
      }

      formattedData.push({
        title: title,
        url: fileUrl,
        author: row[2] || "Unknown",
        createdDate: row[3] || "N/A",
        lastUpdatedBy: row[4] || "N/A",
        effectiveDate: row[5] || "N/A",
        topics: Array.isArray(parsedTopics) ? parsedTopics : [],
        searchIndex: row[7] || "",
        previewIframe: previewIframe
      });
    }
    
    return { success: true, data: formattedData };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function runKPISync() {
  const folderId = "13Dwz2eHDb9MvZJCvtLBkEuh4QD3pNASx";
  const databaseSheetId = "1NNAr_1Z9vgynr0GtKqCuPY8UwS7MqkaW0uiaU5P64ig"; 
  const sheetName = "KPIs";

  const ss = SpreadsheetApp.openById(databaseSheetId);
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(["Title", "URL", "Author", "Created Date", "Last Updated", "Effective Date", "Topics", "Search Index"]);
  }

  const searchQuery = `'${folderId}' in parents and trashed = false`;
  const files = DriveApp.searchFiles(searchQuery); 
  const updatedDataMatrix = [];

  while (files.hasNext()) {
    const file = files.next();
    const docId = file.getId();
    const title = file.getName();
    const fileUrl = file.getUrl();
    const createdDate = file.getDateCreated();
    const lastUpdated = file.getLastUpdated();
    
    let owner = "Team Drive";
    try {
      if (file.getOwner()) {
        owner = file.getOwner().getName() || file.getOwner().getEmail();
      }
    } catch(e) {
      owner = "Openlane Team";
    }

    let topics = [];
    let fullTextIndex = title.toLowerCase();

    if (file.getMimeType() === MimeType.GOOGLE_DOCS) {
      try {
        const doc = DocumentApp.openById(docId);
        const body = doc.getBody();
        fullTextIndex += " " + body.getText().toLowerCase().replace(/\s+/g, ' ').trim();
        
        body.getParagraphs().forEach(p => {
          const heading = p.getHeading();
          if (heading === DocumentApp.ParagraphHeading.HEADING1 || heading === DocumentApp.ParagraphHeading.HEADING2) {
            const txt = p.getText().trim();
            if (txt) topics.push(txt);
          }
        });
      } catch (e) {
        console.warn("Could not parse doc text: " + fileUrl);
      }
    }

    updatedDataMatrix.push([
      title, 
      fileUrl, 
      owner, 
      Utilities.formatDate(createdDate, Session.getScriptTimeZone(), "yyyy-MM-dd"),
      Utilities.formatDate(lastUpdated, Session.getScriptTimeZone(), "yyyy-MM-dd"),
      "N/A", 
      JSON.stringify(topics.slice(0, 10)),
      fullTextIndex 
    ]);
  }

  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 8).clearContent();
  }

  if (updatedDataMatrix.length > 0) {
    sheet.getRange(2, 1, updatedDataMatrix.length, 8).setValues(updatedDataMatrix);
  }
}
