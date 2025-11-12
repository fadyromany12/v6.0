
// === CONFIGURATION ===
const SPREADSHEET_ID = "1FotLFASWuFinDnvpyLTsyO51OpJeKWtuG31VFje3Oik"; // Only the ID
const SICK_NOTE_FOLDER_ID = "1Wu_eoEQ3FmfrzOdAwJkqMu4sPucLRu_0";
const SHEET_NAMES = {
  adherence: "Adherence Tracker",
  database: "Data Base",
  schedule: "Schedules",
  logs: "Logs",
  otherCodes: "Other Codes",
  leaveRequests: "Leave Requests", 
  coaching_OLD: "Coaching", // Renamed old sheet
  coachingSessions: "CoachingSessions", // NEW
  coachingScores: "CoachingScores", // NEW
  coachingTemplates: "CoachingTemplates", // <-- *** ADD THIS LINE ***
  pendingRegistrations: "PendingRegistrations",
  movementRequests: "MovementRequests",
  announcements: "Announcements",
  roleRequests: "Role Requests"
};
// --- Break Time Configuration (in seconds) ---
const PLANNED_BREAK_SECONDS = 15 * 60; // 15 minutes
const PLANNED_LUNCH_SECONDS = 30 * 60; // 30 minutes

// --- Shift Cutoff Hour (e.g., 7 = 7 AM) ---
const SHIFT_CUTOFF_HOUR = 7; 

// ================= WEB APP ENTRY =================
function doGet() {
  return HtmlService.createTemplateFromFile('index') // <-- 1. Change this
    .evaluate() // <-- 2. Add this
    .setTitle('KOMPASS (Konecta Operations, Management & Personnel Self-Service)')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
// ================= WEB APP APIs (UPDATED) =================

// === Web App API for Punching ===
function webPunch(action, targetUserName, adminTimestamp) { 
  try {
    const puncherEmail = Session.getActiveUser().getEmail().toLowerCase();
    const resultMessage = punch(action, targetUserName, puncherEmail, adminTimestamp); // <-- Renamed to 'resultMessage'
    
    // --- START: NEW STATUS LOGIC ---
    // Get the user's new status after the punch
    const userEmail = Session.getActiveUser().getEmail().toLowerCase(); // This is puncher
    
    const dbSheet = getOrCreateSheet(getSpreadsheet(), SHEET_NAMES.database);
    const userData = getUserDataFromDb(dbSheet);
    
    const targetEmail = (targetUserName === userData.emailToName[userEmail]) ? userEmail : userData.nameToEmail[targetUserName];
    
    const timeZone = Session.getScriptTimeZone();
    const now = adminTimestamp ? new Date(adminTimestamp) : new Date();
    const shiftDate = getShiftDate(now, SHIFT_CUTOFF_HOUR);
    const formattedDate = Utilities.formatDate(shiftDate, timeZone, "MM/dd/yyyy");
        
    const newStatus = getLatestPunchStatus(targetEmail, targetUserName, shiftDate, formattedDate);
    // --- END: NEW STATUS LOGIC ---

    return { message: resultMessage, newStatus: newStatus }; // Return an object

  } catch (err) {
    // Return error in the same object format
    return { message: "Error: " + err.message, newStatus: null };
  }
}

// REPLACE this function
function webSubmitScheduleRange(userEmail, userName, startDateStr, endDateStr, startTime, endTime, leaveType, shiftEndDate) { // <-- ADD shiftEndDate
  try {
    const puncherEmail = Session.getActiveUser().getEmail().toLowerCase();
    // *** MODIFIED: Pass all 8 fields to submitScheduleRange ***
    const result = submitScheduleRange(puncherEmail, userEmail, userName, startDateStr, endDateStr, startTime, endTime, leaveType, shiftEndDate);
    return result;
  } catch (err) {
    return "Error: " + err.message;
  }
}

// === Web App APIs for Leave Requests ===
function webSubmitLeaveRequest(requestObject, targetUserEmail) { // Now accepts optional target user
  try {
    const submitterEmail = Session.getActiveUser().getEmail().toLowerCase();
    return submitLeaveRequest(submitterEmail, requestObject, targetUserEmail);
  } catch (err) {
    return "Error: " + err.message;
  }
}

function webGetMyRequests_V2() {
  try {
    const userEmail = Session.getActiveUser().getEmail().toLowerCase();
    return getMyRequests(userEmail); 
  } catch (err) {
    Logger.log("Error in webGetMyRequests_V2: " + err.message);
    throw new Error(err.message); 
  }
}

function webGetAdminLeaveRequests(filter) {
  try {
    const adminEmail = Session.getActiveUser().getEmail().toLowerCase();
    return getAdminLeaveRequests(adminEmail, filter);
  } catch (err) {
    Logger.log("webGetAdminLeaveRequests Error: " + err.message);
    return { error: err.message };
  }
}

function webApproveDenyRequest(requestID, newStatus, reason) {
  try {
    const adminEmail = Session.getActiveUser().getEmail().toLowerCase();
    return approveDenyRequest(adminEmail, requestID, newStatus, reason);
  } catch (err) {
    return "Error: " + err.message;
  }
}

// === Web App API for History ===
function webGetAdherenceRange(userNames, startDateStr, endDateStr) {
  try {
    const userEmail = Session.getActiveUser().getEmail().toLowerCase();
    return getAdherenceRange(userEmail, userNames, startDateStr, endDateStr);
  } catch (err) {
    return { error: "Error: " + err.message };
  }
}

// === Web App API for My Schedule ===
function webGetMySchedule() {
  try {
    const userEmail = Session.getActiveUser().getEmail().toLowerCase();
    return getMySchedule(userEmail);
  } catch (err) {
    return { error: "Error: " + err.message };
  }
}

// === Web App API for Admin Tools ===
function webAdjustLeaveBalance(userEmail, leaveType, amount, reason) {
  try {
    const adminEmail = Session.getActiveUser().getEmail().toLowerCase();
    return adjustLeaveBalance(adminEmail, userEmail, leaveType, amount, reason);
  } catch (err) {
    return "Error: " + err.message;
  }
}

function webImportScheduleCSV(csvData) {
  try {
    const adminEmail = Session.getActiveUser().getEmail().toLowerCase();
    return importScheduleCSV(adminEmail, csvData);
  } catch (err) {
    return "Error: " + err.message;
  }
}

// === Web App API for Dashboard ===
function webGetDashboardData(userEmails, date) { 
  try {
    const adminEmail = Session.getActiveUser().getEmail().toLowerCase();
    return getDashboardData(adminEmail, userEmails, date);
  } catch (err) {
    Logger.log("webGetDashboardData Error: " + err.message);
    throw new Error(err.message);
  }
}

// --- MODIFIED: "My Team" Functions ---
function webSaveMyTeam(userEmails) {
  try {
    const adminEmail = Session.getActiveUser().getEmail().toLowerCase();
    return saveMyTeam(adminEmail, userEmails);
  } catch (err) {
    return "Error: " + err.message;
  }
}

function webGetMyTeam() {
  try {
    const adminEmail = Session.getActiveUser().getEmail().toLowerCase();
    return getMyTeam(adminEmail);
  } catch (err) {
    return "Error: " + err.message;
  }
}

// --- Web App API for Reporting Line (MODIFIED) ---
function webSubmitMovementRequest(userToMoveEmail, newSupervisorEmail) {
  try {
    const requestedByEmail = Session.getActiveUser().getEmail().toLowerCase();
    const ss = getSpreadsheet();
    const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
    const userData = getUserDataFromDb(dbSheet);

    // --- Validation ---
    const adminRole = userData.emailToRole[requestedByEmail] || 'agent';
    if (adminRole !== 'admin' && adminRole !== 'superadmin') {
      throw new Error("Permission denied. Only admins can submit movement requests.");
    }
    
    const userToMoveName = userData.emailToName[userToMoveEmail];
    if (!userToMoveName) {
      throw new Error(`User to move (${userToMoveEmail}) not found.`);
    }
    
    const newSupervisorName = userData.emailToName[newSupervisorEmail];
    if (!newSupervisorName) {
      throw new Error(`Receiving supervisor (${newSupervisorEmail}) not found.`);
    }

    const fromSupervisorEmail = userData.emailToSupervisor[userToMoveEmail];
    if (!fromSupervisorEmail) {
      throw new Error(`User ${userToMoveName} has no current supervisor assigned.`);
    }

    if (fromSupervisorEmail === newSupervisorEmail) {
      throw new Error(`${userToMoveName} already reports to ${newSupervisorName}.`);
    }

    // --- Create Request ---
    const moveSheet = getOrCreateSheet(ss, SHEET_NAMES.movementRequests);
    const movementID = `MOV-${new Date().getTime()}`;

    moveSheet.appendRow([
      movementID,
      "Pending",
      userToMoveEmail,
      userToMoveName,
      fromSupervisorEmail,
      newSupervisorEmail,
      new Date(), // RequestTimestamp
      "", // ActionTimestamp
      "", // ActionByEmail
      requestedByEmail // RequestedByEmail
    ]);

    SpreadsheetApp.flush();
    return `Movement request submitted for ${userToMoveName} to be moved to ${newSupervisorName}.`;

  } catch (err) {
    Logger.log("webSubmitMovementRequest Error: " + err.message);
    return "Error: " + err.message;
  }
}
/**
 * NEW: Fetches pending movement requests for the admin or their subordinates.
 */
function webGetPendingMovements() {
  try {
    const adminEmail = Session.getActiveUser().getEmail().toLowerCase();
    const ss = getSpreadsheet();
    const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
    const userData = getUserDataFromDb(dbSheet);

    // *** ADD THIS LINE TO FIX THE ERROR ***
    const adminRole = userData.emailToRole[adminEmail] || 'agent';
    
    // Get all subordinates (direct and indirect)
    const mySubordinateEmails = new Set(webGetAllSubordinateEmails(adminEmail));
    const moveSheet = getOrCreateSheet(ss, SHEET_NAMES.movementRequests);
    const data = moveSheet.getDataRange().getValues();
    const results = [];

    // Get headers
    const headers = data[0];
    const statusIndex = headers.indexOf("Status");
    const toSupervisorIndex = headers.indexOf("ToSupervisorEmail");
    
    if (statusIndex === -1 || toSupervisorIndex === -1) {
      throw new Error("MovementRequests sheet is missing required columns.");
    }

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const status = row[statusIndex];
      const toSupervisorEmail = (row[toSupervisorIndex] || "").toLowerCase();

      if (status === 'Pending') {
        let canView = false;
        
        // --- NEW VIEWING LOGIC ---
        if (adminRole === 'superadmin') {
          // Superadmin can see ALL pending requests
          canView = true;
        } else if (toSupervisorEmail === adminEmail || mySubordinateEmails.has(toSupervisorEmail)) {
          // Admin can only see requests for themselves or their subordinates
          canView = true;
        }
        // --- END NEW LOGIC ---

        if (canView) {
          results.push({
            movementID: row[headers.indexOf("MovementID")],
            userToMoveName: row[headers.indexOf("UserToMoveName")],
            fromSupervisorName: userData.emailToName[row[headers.indexOf("FromSupervisorEmail")]] || "Unknown",
            
  toSupervisorName: userData.emailToName[row[headers.indexOf("ToSupervisorEmail")]] || "Unknown",
            requestedDate: convertDateToString(new Date(row[headers.indexOf("RequestTimestamp")])),
            requestedByName: userData.emailToName[row[headers.indexOf("RequestedByEmail")]] || "Unknown"
          });
}
      }
    }
    return results;
  } catch (e) {
    Logger.log("webGetPendingMovements Error: " + e.message);
    return { error: e.message };
  }
}

/**
 * NEW: Approves or denies a movement request.
 */
function webApproveDenyMovement(movementID, newStatus) {
  try {
    const adminEmail = Session.getActiveUser().getEmail().toLowerCase();
    const ss = getSpreadsheet();
    const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
    const userData = getUserDataFromDb(dbSheet);
    const moveSheet = getOrCreateSheet(ss, SHEET_NAMES.movementRequests);
    const data = moveSheet.getDataRange().getValues();
    
    // Get headers
    const headers = data[0];
    const idIndex = headers.indexOf("MovementID");
    const statusIndex = headers.indexOf("Status");
    const toSupervisorIndex = headers.indexOf("ToSupervisorEmail");
    const userToMoveIndex = headers.indexOf("UserToMoveEmail");
    const actionTimeIndex = headers.indexOf("ActionTimestamp");
    const actionByIndex = headers.indexOf("ActionByEmail");

    let rowToUpdate = -1;
    let requestDetails = {};

    for (let i = 1; i < data.length; i++) {
      if (data[i][idIndex] === movementID) {
        rowToUpdate = i + 1; // 1-based index
        requestDetails = {
          status: data[i][statusIndex],
          toSupervisorEmail: (data[i][toSupervisorIndex] || "").toLowerCase(),
          userToMoveEmail: (data[i][userToMoveIndex] || "").toLowerCase()
        };
        break;
      }
    }

    if (rowToUpdate === -1) {
      throw new Error("Movement request not found.");
    }
    if (requestDetails.status !== 'Pending') {
      throw new Error(`This request has already been ${requestDetails.status}.`);
    }

    // --- MODIFIED: Security Check ---
    // An admin can action a request if it's FOR them, or FOR a supervisor in their hierarchy.
    
    // Get all subordinates (direct and indirect)
    const mySubordinateEmails = new Set(webGetAllSubordinateEmails(adminEmail));
    
    const isReceivingSupervisor = (requestDetails.toSupervisorEmail === adminEmail);
    // Check if the request is for someone who reports to the admin
    const isSupervisorOfReceiver = mySubordinateEmails.has(requestDetails.toSupervisorEmail);

    if (!isReceivingSupervisor && !isSupervisorOfReceiver) {
      // This check covers all roles. 
      // An Admin/Superadmin can only approve for their own hierarchy (as you requested: "for a only not for b").
      throw new Error("Permission denied. You can only approve requests for yourself or for supervisors in your reporting line.");
    }
    // --- END MODIFICATION ---
    // All checks passed, update the status
    moveSheet.getRange(rowToUpdate, statusIndex + 1).setValue(newStatus);
    moveSheet.getRange(rowToUpdate, actionTimeIndex + 1).setValue(new Date());
    moveSheet.getRange(rowToUpdate, actionByIndex + 1).setValue(adminEmail);

    if (newStatus === 'Approved') {
      // Find the user in the Data Base
      const userDBRow = userData.emailToRow[requestDetails.userToMoveEmail];
      if (!userDBRow) {
        throw new Error(`Could not find user ${requestDetails.userToMoveEmail} in Data Base to update.`);
      }
      // Update their supervisor (Column G = 7)
      dbSheet.getRange(userDBRow, 7).setValue(requestDetails.toSupervisorEmail);

      // Log the change
      const logsSheet = getOrCreateSheet(ss, SHEET_NAMES.logs);
      logsSheet.appendRow([
        new Date(), 
        userData.emailToName[requestDetails.userToMoveEmail] || "Unknown User", 
        adminEmail, 
        "Reporting Line Change Approved", 
        `MovementID: ${movementID}`
      ]);
    }
    
    SpreadsheetApp.flush();
    return { success: true, message: `Request has been ${newStatus}.` };

  } catch (e) {
    Logger.log("webApproveDenyMovement Error: " + e.message);
    return { error: e.message };
  }
}

/**
 * NEW: Fetches the movement history for a selected user.
 */
function webGetMovementHistory(selectedUserEmail) {
  try {
    const adminEmail = Session.getActiveUser().getEmail().toLowerCase();
    const ss = getSpreadsheet();
    const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
    const userData = getUserDataFromDb(dbSheet);
    
    // Security check: Is this admin allowed to see this user's history?
    const adminRole = userData.emailToRole[adminEmail];
    const mySubordinateEmails = new Set(webGetAllSubordinateEmails(adminEmail));

    if (adminRole !== 'superadmin' && !mySubordinateEmails.has(selectedUserEmail)) {
      throw new Error("Permission denied. You can only view the history of users in your reporting line.");
    }
    
    const moveSheet = getOrCreateSheet(ss, SHEET_NAMES.movementRequests);
    const data = moveSheet.getDataRange().getValues();
    const headers = data[0];
    const results = [];

    // Find rows where the user was the one being moved
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const userToMoveEmail = (row[headers.indexOf("UserToMoveEmail")] || "").toLowerCase();
      
      if (userToMoveEmail === selectedUserEmail) {
        results.push({
          status: row[headers.indexOf("Status")],
          requestDate: convertDateToString(new Date(row[headers.indexOf("RequestTimestamp")])),
          actionDate: convertDateToString(new Date(row[headers.indexOf("ActionTimestamp")])),
          fromSupervisorName: userData.emailToName[row[headers.indexOf("FromSupervisorEmail")]] || "N/A",
          toSupervisorName: userData.emailToName[row[headers.indexOf("ToSupervisorEmail")]] || "N/A",
          actionByName: userData.emailToName[row[headers.indexOf("ActionByEmail")]] || "N/A",
          requestedByName: userData.emailToName[row[headers.indexOf("RequestedByEmail")]] || "N/A"
        });
      }
    }
    
    // Sort by request date, newest first
    results.sort((a, b) => new Date(b.requestDate) - new Date(a.requestDate));
    return results;

  } catch (e) {
    Logger.log("webGetMovementHistory Error: " + e.message);
    return { error: e.message };
  }
}

// ==========================================================
// === NEW/REPLACED COACHING FUNCTIONS (START) ===
// ==========================================================

/**
 * (REPLACED)
 * Saves a new coaching session and its detailed scores.
 * Matches the new frontend form.
 */
function webSubmitCoaching(sessionObject) {
  try {
    const ss = getSpreadsheet();
    const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
    const userData = getUserDataFromDb(dbSheet);
    const sessionSheet = getOrCreateSheet(ss, SHEET_NAMES.coachingSessions);
    const scoreSheet = getOrCreateSheet(ss, SHEET_NAMES.coachingScores);
    
    const coachEmail = Session.getActiveUser().getEmail().toLowerCase();
    const coachName = userData.emailToName[coachEmail] || coachEmail;
    
    // Simple validation
    if (!sessionObject.agentEmail || !sessionObject.sessionDate) {
      throw new Error("Agent and Session Date are required.");
    }

    const agentName = userData.emailToName[sessionObject.agentEmail.toLowerCase()];
    if (!agentName) {
      throw new Error(`Could not find agent with email ${sessionObject.agentEmail}.`);
    }

    const sessionID = `CS-${new Date().getTime()}`; // Simple unique ID
    const sessionDate = new Date(sessionObject.sessionDate + 'T00:00:00');
    // *** NEW: Handle FollowUpDate ***
    const followUpDate = sessionObject.followUpDate ? new Date(sessionObject.followUpDate + 'T00:00:00') : null;
    const followUpStatus = followUpDate ? "Pending" : ""; // Set to pending if date exists

    // 1. Log the main session
    sessionSheet.appendRow([
      sessionID,
      sessionObject.agentEmail,
      agentName,
      coachEmail,
      coachName,
      sessionDate,
      sessionObject.weekNumber,
      sessionObject.overallScore,
      sessionObject.followUpComment,
      new Date(), // Timestamp of submission
      followUpDate || "", // *** NEW: Add follow-up date ***
      followUpStatus  // *** NEW: Add follow-up status ***
    ]);

    // 2. Log the individual scores
    const scoresToLog = [];
    if (sessionObject.scores && Array.isArray(sessionObject.scores)) {
      sessionObject.scores.forEach(score => {
        scoresToLog.push([
          sessionID,
          score.category,
          score.criteria,
          score.score,
          score.comment
        ]);
      });
    }

    if (scoresToLog.length > 0) {
      scoreSheet.getRange(scoreSheet.getLastRow() + 1, 1, scoresToLog.length, 5).setValues(scoresToLog);
    }
    
    return `Coaching session for ${agentName} saved successfully.`;

  } catch (err) {
    Logger.log("webSubmitCoaching Error: " + err.message);
    return "Error: " + err.message;
  }
}

/**
 * (REPLACED)
 * Gets coaching history for the logged-in user or their team.
 * Reads from the new CoachingSessions sheet.
 */
function webGetCoachingHistory(filter) { // filter is unused for now, but good practice
  try {
    const userEmail = Session.getActiveUser().getEmail().toLowerCase();
    const ss = getSpreadsheet();
    const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
    const userData = getUserDataFromDb(dbSheet);
    const role = userData.emailToRole[userEmail] || 'agent';
    const sheet = getOrCreateSheet(ss, SHEET_NAMES.coachingSessions);

    // Get all data as objects
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const allData = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    
    const allSessions = allData.map(row => {
      let obj = {};
      headers.forEach((header, index) => {
        obj[header] = row[index];
      });
      return obj;
    });

    const results = [];
    
    // Get a list of users this person manages (if they are a manager)
    let myTeamEmails = new Set();
    if (role === 'admin' || role === 'superadmin') {
      // Use the hierarchy-aware function
      const myTeamList = webGetAllSubordinateEmails(userEmail);
      myTeamList.forEach(email => myTeamEmails.add(email.toLowerCase()));
    }

    for (let i = allSessions.length - 1; i >= 0; i--) {
      const session = allSessions[i];
      if (!session || !session.AgentEmail) continue; // Skip empty/invalid rows

      const agentEmail = session.AgentEmail.toLowerCase();

      let canView = false;
      
      // *** MODIFIED LOGIC HERE ***
      if (agentEmail === userEmail) {
        // Anyone can see their own coaching
        canView = true;
      } else if (role === 'admin' && myTeamEmails.has(agentEmail)) {
        // An admin can see their team's
        canView = true;
      } else if (role === 'superadmin') {
        // Superadmin can see all (team members + their own, which is covered above)
        canView = true;
      }
      // *** END MODIFIED LOGIC ***

      if (canView) {
        results.push({
          sessionID: session.SessionID,
          agentName: session.AgentName,
          coachName: session.CoachName,
          sessionDate: convertDateToString(new Date(session.SessionDate)),
          weekNumber: session.WeekNumber,
          overallScore: session.OverallScore,
          followUpComment: session.FollowUpComment,
          followUpDate: convertDateToString(new Date(session.FollowUpDate)),
          followUpStatus: session.FollowUpStatus,
          agentAcknowledgementTimestamp: convertDateToString(new Date(session.AgentAcknowledgementTimestamp))
        });
      }
    }
    return results;

  } catch (err) {
    Logger.log("webGetCoachingHistory Error: " + err.message);
    return { error: err.message };
  }
}

/**
 * NEW: Fetches the details for a single coaching session.
 * (MODIFIED: Renamed to webGetCoachingSessionDetails to be callable)
 * (MODIFIED 2: Added date-to-string conversion to fix null return)
 * (MODIFIED 3: Added AgentAcknowledgementTimestamp conversion)
 */
function webGetCoachingSessionDetails(sessionID) {
  try {
    const ss = getSpreadsheet();
    const sessionSheet = getOrCreateSheet(ss, SHEET_NAMES.coachingSessions);
    const scoreSheet = getOrCreateSheet(ss, SHEET_NAMES.coachingScores);

    const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
    const userData = getUserDataFromDb(dbSheet);

    // 1. Get Session Summary
    const sessionHeaders = sessionSheet.getRange(1, 1, 1, sessionSheet.getLastColumn()).getValues()[0];
    const sessionData = sessionSheet.getDataRange().getValues();
    let sessionSummary = null;

    for (let i = 1; i < sessionData.length; i++) {
      if (sessionData[i][0] === sessionID) {
        sessionSummary = {};
        sessionHeaders.forEach((header, index) => {
          sessionSummary[header] = sessionData[i][index];
        });
        break;
      }
    }

    if (!sessionSummary) {
      throw new Error("Session not found.");
    }

    // 2. Get Session Scores
    const scoreHeaders = scoreSheet.getRange(1, 1, 1, scoreSheet.getLastColumn()).getValues()[0];
    const scoreData = scoreSheet.getDataRange().getValues();
    const sessionScores = [];

    for (let i = 1; i < scoreData.length; i++) {
      if (scoreData[i][0] === sessionID) {
        let scoreObj = {};
        scoreHeaders.forEach((header, index) => {
          scoreObj[header] = scoreData[i][index];
        });
        sessionScores.push(scoreObj);
      }
    }
    
    sessionSummary.CoachName = userData.emailToName[sessionSummary.CoachEmail] || sessionSummary.CoachName;
    
    // *** Convert Date objects to Strings before returning ***
    sessionSummary.SessionDate = convertDateToString(new Date(sessionSummary.SessionDate));
    sessionSummary.SubmissionTimestamp = convertDateToString(new Date(sessionSummary.SubmissionTimestamp));
    sessionSummary.FollowUpDate = convertDateToString(new Date(sessionSummary.FollowUpDate));
    // *** NEW: Convert the new column ***
    sessionSummary.AgentAcknowledgementTimestamp = convertDateToString(new Date(sessionSummary.AgentAcknowledgementTimestamp));
    // *** END NEW SECTION ***

    return {
      summary: sessionSummary,
      scores: sessionScores
    };

  } catch (err) {
    Logger.log("webGetCoachingSessionDetails Error: " + err.message);
    return { error: err.message };
  }
}

/**
 * NEW: Updates the follow-up status for a coaching session.
 */
function webUpdateFollowUpStatus(sessionID, newStatus, newDateStr) {
  try {
    const adminEmail = Session.getActiveUser().getEmail().toLowerCase();
    const ss = getSpreadsheet();
    
    // Check permission
    const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
    const userData = getUserDataFromDb(dbSheet);
    const adminRole = userData.emailToRole[adminEmail] || 'agent';

    if (adminRole !== 'admin' && adminRole !== 'superadmin') {
      throw new Error("Permission denied. Only managers can update follow-up status.");
    }
    
    const sessionSheet = getOrCreateSheet(ss, SHEET_NAMES.coachingSessions);
    const sessionData = sessionSheet.getDataRange().getValues();
    const sessionHeaders = sessionData[0];
    
    // Find the column indexes
    const statusColIndex = sessionHeaders.indexOf("FollowUpStatus");
    const dateColIndex = sessionHeaders.indexOf("FollowUpDate");
    
    if (statusColIndex === -1 || dateColIndex === -1) {
      throw new Error("Could not find 'FollowUpStatus' or 'FollowUpDate' columns in CoachingSessions sheet.");
    }

    // Find the row
    let sessionRow = -1;
    for (let i = 1; i < sessionData.length; i++) {
      if (sessionData[i][0] === sessionID) {
        sessionRow = i + 1; // 1-based index
        break;
      }
    }

    if (sessionRow === -1) {
      throw new Error("Session not found.");
    }

    // Prepare new values
    let newFollowUpDate = null;
    if (newDateStr) {
      newFollowUpDate = new Date(newDateStr + 'T00:00:00');
    } else {
      // If marking completed, use today's date
      newFollowUpDate = new Date();
    }
    
    // Update the sheet
    sessionSheet.getRange(sessionRow, statusColIndex + 1).setValue(newStatus);
    sessionSheet.getRange(sessionRow, dateColIndex + 1).setValue(newFollowUpDate);

    SpreadsheetApp.flush(); // Ensure changes are saved

    return { success: true, message: `Status updated to ${newStatus}.` };

  } catch (err) {
    Logger.log("webUpdateFollowUpStatus Error: " + err.message);
    return { error: err.message };
  }
}

/**
 * NEW: Allows an agent to acknowledge their coaching session.
 */
function webSubmitCoachingAcknowledgement(sessionID) {
  try {
    const userEmail = Session.getActiveUser().getEmail().toLowerCase();
    const ss = getSpreadsheet();
    const sessionSheet = getOrCreateSheet(ss, SHEET_NAMES.coachingSessions);

    // *** MODIFIED: Explicitly read headers ***
    const sessionHeaders = sessionSheet.getRange(1, 1, 1, sessionSheet.getLastColumn()).getValues()[0];
    // Get data rows separately, skipping header
    const sessionData = sessionSheet.getRange(2, 1, sessionSheet.getLastRow() - 1, sessionSheet.getLastColumn()).getValues();

    // Find the column indexes
    const ackColIndex = sessionHeaders.indexOf("AgentAcknowledgementTimestamp");
    const agentEmailColIndex = sessionHeaders.indexOf("AgentEmail");
    if (ackColIndex === -1 || agentEmailColIndex === -1) {
      throw new Error("Could not find 'AgentAcknowledgementTimestamp' or 'AgentEmail' columns in CoachingSessions sheet.");
    }

    // Find the row
    let sessionRow = -1;
    let agentEmailOnRow = null;
    let currentAckStatus = null;

    // *** MODIFIED: Loop starts at 0 and row index is i + 2 ***
    for (let i = 0; i < sessionData.length; i++) {
      if (sessionData[i][0] === sessionID) {
        sessionRow = i + 2; // Data starts from row 2
        agentEmailOnRow = sessionData[i][agentEmailColIndex].toLowerCase();
        currentAckStatus = sessionData[i][ackColIndex];
        break;
      }
    }

    if (sessionRow === -1) {
      throw new Error("Session not found.");
    }
    
    // Security Check: Is this the correct agent?
    if (agentEmailOnRow !== userEmail) {
      throw new Error("Permission denied. You can only acknowledge your own coaching sessions.");
    }
    
    // Check if already acknowledged
    if (currentAckStatus) {
      return { success: false, message: "This session has already been acknowledged." };
    }
    
    // Update the sheet
    sessionSheet.getRange(sessionRow, ackColIndex + 1).setValue(new Date());

    SpreadsheetApp.flush(); // Ensure changes are saved

    return { success: true, message: "Coaching session acknowledged successfully." };

  } catch (err) {
    Logger.log("webSubmitCoachingAcknowledgement Error: " + err.message);
    return { error: err.message };
  }
}


/**
 * NEW: Gets a list of unique, active template names.
 */
function webGetActiveTemplates() {
  try {
    const ss = getSpreadsheet();
    const templateSheet = getOrCreateSheet(ss, SHEET_NAMES.coachingTemplates);
    const data = templateSheet.getRange(2, 1, templateSheet.getLastRow() - 1, 4).getValues();
    
    const templateNames = new Set();
    
    data.forEach(row => {
      const templateName = row[0];
      const status = row[3];
      if (templateName && status === 'Active') {
        templateNames.add(templateName);
      }
    });
    
    return Array.from(templateNames).sort();
    
  } catch (err) {
    Logger.log("webGetActiveTemplates Error: " + err.message);
    return { error: err.message };
  }
}

/**
 * NEW: Gets all criteria for a specific template name.
 */
function webGetTemplateCriteria(templateName) {
  try {
    const ss = getSpreadsheet();
    const templateSheet = getOrCreateSheet(ss, SHEET_NAMES.coachingTemplates);
    const data = templateSheet.getRange(2, 1, templateSheet.getLastRow() - 1, 4).getValues();
    
    const categories = {}; // Use an object to group criteria by category
    
    data.forEach(row => {
      const name = row[0];
      const category = row[1];
      const criteria = row[2];
      const status = row[3];
      
      if (name === templateName && status === 'Active' && category && criteria) {
        if (!categories[category]) {
          categories[category] = [];
        }
        categories[category].push(criteria);
      }
    });
    
    // Convert from object to the array structure the frontend expects
    const results = Object.keys(categories).map(categoryName => {
      return {
        category: categoryName,
        criteria: categories[categoryName]
      };
    });
    
    return results;
    
  } catch (err) {
    Logger.log("webGetTemplateCriteria Error: " + err.message);
    return { error: err.message };
  }
}

// ==========================================================
// === NEW/REPLACED COACHING FUNCTIONS (END) ===
// ==========================================================

// [START] MODIFICATION 8: Add webSaveNewTemplate function
/**
 * NEW: Saves a new coaching template from the admin tab.
 */
function webSaveNewTemplate(templateName, categories) {
  try {
    const adminEmail = Session.getActiveUser().getEmail().toLowerCase();
    const ss = getSpreadsheet();
    
    // Check permission
    const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
    const userData = getUserDataFromDb(dbSheet);
    const adminRole = userData.emailToRole[adminEmail] || 'agent';

    if (adminRole !== 'admin' && adminRole !== 'superadmin') {
      throw new Error("Permission denied. Only managers can create templates.");
    }
    
    // Validation
    if (!templateName) {
      throw new Error("Template Name is required.");
    }
    if (!categories || categories.length === 0) {
      throw new Error("At least one category is required.");
    }

    const templateSheet = getOrCreateSheet(ss, SHEET_NAMES.coachingTemplates);
    
    // Check if template name already exists
    const templateNames = templateSheet.getRange(2, 1, templateSheet.getLastRow() - 1, 1).getValues();
    const
      lowerTemplateName = templateName.toLowerCase();
    for (let i = 0; i < templateNames.length; i++) {
      if (templateNames[i][0] && templateNames[i][0].toLowerCase() === lowerTemplateName) {
        throw new Error(`A template with the name '${templateName}' already exists.`);
      }
    }

    const rowsToAppend = [];
    categories.forEach(category => {
      if (category.criteria && category.criteria.length > 0) {
        category.criteria.forEach(criterion => {
          rowsToAppend.push([
            templateName,
            category.name,
            criterion,
            'Active' // Default to Active
          ]);
        });
      }
    });

    if (rowsToAppend.length === 0) {
      throw new Error("No criteria were found to save.");
    }
    
    // Write all new rows at once
    templateSheet.getRange(templateSheet.getLastRow() + 1, 1, rowsToAppend.length, 4).setValues(rowsToAppend);
    
    SpreadsheetApp.flush();
    return `Template '${templateName}' saved successfully with ${rowsToAppend.length} criteria.`;

  } catch (err) {
    Logger.log("webSaveNewTemplate Error: " + err.message);
    return "Error: " + err.message;
  }
}
// [END] MODIFICATION 8

// === NEW: Web App API for Manager Hierarchy ===
function webGetManagerHierarchy() {
  try {
    const managerEmail = Session.getActiveUser().getEmail().toLowerCase();
    const ss = getSpreadsheet();
    const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
    const userData = getUserDataFromDb(dbSheet);
    
    const managerRole = userData.emailToRole[managerEmail] || 'agent';
    if (managerRole === 'agent') {
      return { error: "Permission denied. Only managers can view the hierarchy." };
    }
    
    // --- Step 1: Build the direct reporting map (Supervisor -> [Subordinates]) ---
    const reportsMap = {};
    const userEmailMap = {}; // Map email -> {name, role}

    userData.userList.forEach(user => {
      userEmailMap[user.email] = { name: user.name, role: user.role };
      const supervisorEmail = user.supervisor;
      
      if (supervisorEmail) {
        if (!reportsMap[supervisorEmail]) {
          reportsMap[supervisorEmail] = [];
        }
        reportsMap[supervisorEmail].push(user.email);
      }
    });

    // --- Step 2: Recursive function to build the tree (Hierarchy) ---
    // MODIFIED: Added `visited` Set to track users in the current path.
    function buildHierarchy(currentEmail, depth = 0, visited = new Set()) {
      const user = userEmailMap[currentEmail];
      
      // If the email doesn't map to a user, it's likely a blank entry in the DB, so return null
      if (!user) return null; 
      
      // CRITICAL CHECK: Detect circular reference
      if (visited.has(currentEmail)) {
        Logger.log(`Circular reference detected at user: ${currentEmail}`);
        return {
          email: currentEmail,
          name: user.name,
          role: user.role,
          subordinates: [],
          circularError: true
        };
      }
      
      // Add current user to visited set for this path
      const newVisited = new Set(visited).add(currentEmail);


      const subordinates = reportsMap[currentEmail] || [];
      
      // Separate managers/admins from agents
      const adminSubordinates = subordinates
        .filter(email => userData.emailToRole[email] === 'admin' || userData.emailToRole[email] === 'superadmin')
        .map(email => buildHierarchy(email, depth + 1, newVisited))
        .filter(s => s !== null); // Build sub-teams for managers

      const agentSubordinates = subordinates
        .filter(email => userData.emailToRole[email] === 'agent')
        .map(email => ({
          email: email,
          name: userEmailMap[email].name,
          role: userEmailMap[email].role,
          subordinates: [] // Agents have no subordinates
        }));
        
      // Combine and sort: Managers first, then Agents, then alphabetically
      const combinedSubordinates = [...adminSubordinates, ...agentSubordinates];
      
      combinedSubordinates.sort((a, b) => {
          // Sort by role (manager/admin first)
          const aIsManager = a.role !== 'agent';
          const bIsManager = b.role !== 'agent';
          
          if (aIsManager && !bIsManager) return -1;
          if (!aIsManager && bIsManager) return 1;
          
          // Then sort by name
          return a.name.localeCompare(b.name);
      });


      return {
        email: currentEmail,
        name: user.name,
        role: user.role,
        subordinates: combinedSubordinates,
        depth: depth
      };
    }

    // Start building the hierarchy from the manager's email
    const hierarchy = buildHierarchy(managerEmail);
    
    // Check if the root node returned a circular error
    if (hierarchy && hierarchy.circularError) {
        throw new Error("Critical Error: Circular reporting line detected at the top level.");
    }

    return hierarchy;

  } catch (err) {
    Logger.log("webGetManagerHierarchy Error: " + err.message);
    throw new Error(err.message);
  }
}

// === NEW: Web App API to get all reports (flat list) ===
function webGetAllSubordinateEmails(managerEmail) {
    try {
        const ss = getSpreadsheet();
        const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
        const userData = getUserDataFromDb(dbSheet);
        
        const managerRole = userData.emailToRole[managerEmail] || 'agent';
        if (managerRole === 'agent') {
            throw new Error("Permission denied.");
        }
        
        // --- Build the direct reporting map ---
        const reportsMap = {};
        userData.userList.forEach(user => {
            const supervisorEmail = user.supervisor;
            if (supervisorEmail) {
                if (!reportsMap[supervisorEmail]) {
                    reportsMap[supervisorEmail] = [];
                }
                reportsMap[supervisorEmail].push(user.email);
            }
        });
        
        const allSubordinates = new Set();
        const queue = [managerEmail];
        
        // Use a set to track users we've already processed (including the manager him/herself)
        const processed = new Set();
        
        while (queue.length > 0) {
            const currentEmail = queue.shift();
            
            // Check for processing loop (shouldn't happen in BFS, but safe check)
            if (processed.has(currentEmail)) continue;
            processed.add(currentEmail);

            const directReports = reportsMap[currentEmail] || [];
            
            directReports.forEach(reportEmail => {
                if (!allSubordinates.has(reportEmail)) {
                    allSubordinates.add(reportEmail);
                    // If the report is a manager, add them to the queue to find their reports
                    if (userData.emailToRole[reportEmail] !== 'agent') {
                        queue.push(reportEmail); // <-- FIX: Was 'push(reportEmail)'
                    }
                }
            
        });
        }
        
        // Return all subordinates *plus* the manager
        allSubordinates.add(managerEmail);
        return Array.from(allSubordinates);

    } catch (err) {
        Logger.log("webGetAllSubordinateEmails Error: " + err.message);
        return [];
    }
}
// --- END OF WEB APP API SECTION ---


// Get user info for front-end display
function getUserInfo() { 
  try {
    const userEmail = Session.getActiveUser().getEmail().toLowerCase();
    const ss = getSpreadsheet();
    const timeZone = Session.getScriptTimeZone(); 
    const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
    
    let userData = getUserDataFromDb(dbSheet); 
    let isNewUser = false; 

    const KONECTA_DOMAIN = "@konecta.com"; 
    if (!userData.emailToName[userEmail] && userEmail.endsWith(KONECTA_DOMAIN)) {
      Logger.log(`New user detected: ${userEmail}. Auto-registering.`);
      isNewUser = true;

      const nameParts = userEmail.split('@')[0].split('.');
      const firstName = nameParts[0] ? nameParts[0].charAt(0).toUpperCase() + nameParts[0].slice(1) : '';
      const lastName = nameParts[1] ? nameParts[1].charAt(0).toUpperCase() + nameParts[1].slice(1) : '';
      const newName = [firstName, lastName].join(' ').trim();
      
      dbSheet.appendRow([
        newName || userEmail, // User Name
        userEmail,  // Email
        'agent',    // Role
        0,          // Annual Balance
        0,          // Sick Balance
        0,          // Casual Balance
        "",         // SupervisorEmail (BLANK)
        "Pending"   // *** NEW: AccountStatus set to Pending ***
      ]);
      SpreadsheetApp.flush(); 
      userData = getUserDataFromDb(dbSheet);
    }
    
    // *** NEW: Get the user's account status ***
    const accountStatus = userData.emailToAccountStatus[userEmail] || 'Pending';
    const userName = userData.emailToName[userEmail] || ""; 
    const role = userData.emailToRole[userEmail] || 'agent'; 

    // --- START: NEW STATUS LOGIC ---
    let currentStatus = null;
    if (accountStatus === 'Active') {
      const now = new Date();
      const shiftDate = getShiftDate(now, SHIFT_CUTOFF_HOUR);
      const formattedDate = Utilities.formatDate(shiftDate, timeZone, "MM/dd/yyyy");
      currentStatus = getLatestPunchStatus(userEmail, userName, shiftDate, formattedDate);
    }
    // --- END: NEW STATUS LOGIC ---

    let allUsers = []; 
    let allAdmins = [];
    // *** MODIFIED: Send admin list to new users so they can pick a supervisor ***
    if (role === 'admin' || role === 'superadmin' || isNewUser || accountStatus === 'Pending') { 
      allUsers = userData.userList;
    }
    
    allAdmins = userData.userList.filter(u => u.role === 'admin' || u.role === 'superadmin');
    
    const myBalances = userData.emailToBalances[userEmail] ||
    { annual: 0, sick: 0, casual: 0 };

    // *** ADD THIS BLOCK ***
let hasPendingRoleRequests = false;
if (role === 'superadmin') {
  const reqSheet = getOrCreateSheet(ss, SHEET_NAMES.roleRequests);
const data = reqSheet.getDataRange().getValues();
  const statusIndex = data[0].indexOf("Status");
  for (let i = 1; i < data.length; i++) {
    if (data[i][statusIndex] === 'Pending') {
      hasPendingRoleRequests = true;
break;
    }
  }
}


// *** END BLOCK ***

    return {
      name: userName, 
      email: userEmail,
      role: role,
      allUsers: allUsers,
      allAdmins: allAdmins,
      myBalances: myBalances,
      isNewUser: isNewUser, // This flag is still useful
      accountStatus: accountStatus, // *** NEW: Send status to frontend ***
      hasPendingRoleRequests: hasPendingRoleRequests, // *** ADD THIS LINE ***
      currentStatus: currentStatus // <-- ADD THIS LINE
    };
} catch (e) {
    throw new Error("Failed in getUserInfo: " + e.message);
  }
}

// REPLACE this function in your code.gs file
// ================= PUNCH MAIN FUNCTION =================
function punch(action, targetUserName, puncherEmail, adminTimestamp) { 
  const ss = getSpreadsheet();
  const adherenceSheet = getOrCreateSheet(ss, SHEET_NAMES.adherence);
  const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
  const scheduleSheet = getOrCreateSheet(ss, SHEET_NAMES.schedule);
  const logsSheet = getOrCreateSheet(ss, SHEET_NAMES.logs);
  const otherCodesSheet = getOrCreateSheet(ss, SHEET_NAMES.otherCodes);
  const timeZone = Session.getScriptTimeZone(); 

  // === 1. GET ALL USER DATA ===
  const userData = getUserDataFromDb(dbSheet);
  // === 2. IDENTIFY PUNCHER & TARGET ===
  const puncherRole = userData.emailToRole[puncherEmail] || 'agent';
  const puncherIsAdmin = (puncherRole === 'admin' || puncherRole === 'superadmin');
  
  const userName = targetUserName; 
  const userEmail = userData.nameToEmail[userName];
  if (!puncherIsAdmin && puncherEmail !== userEmail) { 
    throw new Error("Permission denied. You can only submit punches for yourself.");
  }
  const isAdmin = puncherIsAdmin; 
  
  // === 3. VALIDATE TARGET USER ===
  if (!userEmail) { 
     throw new Error(`User "${userName}" not found in Data Base.`);
  }
  if (!userName && !puncherIsAdmin) { 
    throw new Error("Your email is not registered in the Data Base sheet. Contact your supervisor.");
  }
  
  const nowTimestamp = adminTimestamp ? new Date(adminTimestamp) : new Date();
  const shiftDate = getShiftDate(new Date(nowTimestamp), SHIFT_CUTOFF_HOUR);
  const formattedDate = Utilities.formatDate(shiftDate, timeZone, "MM/dd/yyyy");

  // === 4. HANDLE "OTHER CODES" ===
  const otherCodeActions = ["Meeting", "Personal", "Coaching"];
  for (const code of otherCodeActions) {
    if (action.startsWith(code)) {
      const resultMsg = logOtherCode(
        otherCodesSheet, userName, action, nowTimestamp, 
        isAdmin && (puncherEmail !== userEmail || adminTimestamp) ? puncherEmail : null 
      );
      logsSheet.appendRow([new Date(), userName, userEmail, action, nowTimestamp]); 
      return resultMsg;
    }
  }

  // === 5. PROCEED WITH ADHERENCE PUNCH ===
  const scheduleData = scheduleSheet.getDataRange().getValues();
  // *** MODIFIED for 7-column layout ***
  let schName, schStartDate, schStartTime, schEndDate, schEndTime, schLeave, schEmail;
  let shiftStartStr = "", shiftEndStr = "", leaveType = "";
  let shiftStartDateObj = null, shiftEndDateObj = null;
  let foundSchedule = false;

  for (let i = 1; i < scheduleData.length; i++) {
    // Read all 7 columns
    [schName, schStartDate, schStartTime, schEndDate, schEndTime, schLeave, schEmail] = scheduleData[i];
    const dateObj = parseDate(schStartDate); // Use StartDate (Col B) for matching
    if (!dateObj || isNaN(dateObj.getTime())) continue;
    const dateStr = Utilities.formatDate(dateObj, timeZone, "MM/dd/yyyy");
    
    // Check Email (Col G, index 6)
    if (schEmail && schEmail.toLowerCase() === userEmail && dateStr === formattedDate) { 
      
      // Get Start Time (Col C, index 2)
      if (schStartTime instanceof Date) {
        shiftStartStr = Utilities.formatDate(schStartTime, timeZone, "HH:mm:ss");
      } else {
        shiftStartStr = (schStartTime || "").toString();
      }
      
      // Get End Time (Col E, index 4)
      if (schEndTime instanceof Date) {
        shiftEndStr = Utilities.formatDate(schEndTime, timeZone, "HH:mm:ss");
      } else {
        shiftEndStr = (schEndTime || "").toString();
      }
          
      // Get Leave Type (Col F, index 5)
      leaveType = (schLeave || "").toString().trim();
      // *** NEW: Build full shift start/end Date objects ***
      if (shiftStartStr) {
        shiftStartDateObj = createDateTime(new Date(schStartDate), shiftStartStr);
      }
      if (shiftEndStr) {
        // Use schEndDate (Col D) if it exists, otherwise use schStartDate (Col B)
        const baseEndDate = schEndDate ? new Date(schEndDate) : new Date(schStartDate);
        shiftEndDateObj = createDateTime(baseEndDate, shiftEndStr);
        
        // Handle overnight *only if* no explicit EndDate was given
        if (shiftEndDateObj && !schEndDate && shiftEndDateObj <= shiftStartDateObj) {
          shiftEndDateObj.setDate(shiftEndDateObj.getDate() + 1);
        }
      }
      // *** END NEW ***
      
      foundSchedule = true;
      break;
    }
  }

  // *** MODIFIED for Request 3: Handle "Day Off" ***
  // No schedule row was found for this user on this date
  if (!foundSchedule) { 
    throw new Error(`Today is a scheduled Day Off. No punches are required.`);
  }

  // *** MODIFIED: Logic for "Present" vs. "Leave" ***
  // If LeaveType is empty, it's "Day Off"
  if (leaveType === "") {
    throw new Error(`Today is a scheduled Day Off. No punches are required.`);
  }
  // If StartTime is empty, but LeaveType is not, it's leave
  if (!shiftStartStr && leaveType) {
    const row = findOrCreateRow(adherenceSheet, userName, shiftDate, formattedDate);
    adherenceSheet.getRange(row, 14).setValue(leaveType);
    if (leaveType.toLowerCase() === "absent") {
      adherenceSheet.getRange(row, 20).setValue("Yes");
    }
    return `${userName}: Leave type "${leaveType}" recorded. No further punches needed.`;
  }
  // If we are here, LeaveType is "Present"
  const row = findOrCreateRow(adherenceSheet, userName, shiftDate, formattedDate); 
  adherenceSheet.getRange(row, 14).setValue("Present");
  // --- End of Schedule Read Logic ---

  const columns = {
    "Login": 3, "First Break In": 4, "First Break Out": 5, "Lunch In": 6, 
    "Lunch Out": 7, "Last Break In": 8, "Last Break Out": 9, "Logout": 10
  };
  const col = columns[action];
  if (!col) throw new Error("Invalid action: " + action);

  // --- START: MODIFICATION FOR REQUEST 1 (Prevent Double "In" Punch) ---
  const isActionIn = (action === "Login" || action === "First Break In" || action === "Lunch In" || action === "Last Break In");
  const existingValue = adherenceSheet.getRange(row, col).getValue();
  
  if (isActionIn && existingValue) {
      throw new Error(`Error: "${action}" has already been punched today.`);
  }
  // --- END: MODIFICATION FOR REQUEST 1 ---

  const currentPunches = adherenceSheet.getRange(row, 3, 1, 8).getValues()[0];
  const punches = {
    login: currentPunches[0], firstBreakIn: currentPunches[1], firstBreakOut: currentPunches[2],
    lunchIn: currentPunches[3], lunchOut: currentPunches[4], lastBreakIn: currentPunches[5],
    lastBreakOut: currentPunches[6], logout: currentPunches[7]
  };
  
  // This block now only checks for sequential errors for non-admins
  if (!isAdmin) {
    if (action !== "Login" && !punches.login) {
      throw new Error("You must 'Login' first.");
    }
    const sequentialErrors = {
      "First Break Out": { required: punches.firstBreakIn, msg: "You must punch 'First Break In' first." },
      "Lunch Out":       { required: punches.lunchIn,     msg: "YouS must punch 'Lunch In' first." },
      "Last Break Out":  { required: punches.lastBreakIn,   msg: "YouS must punch 'Last Break In' first." }
    };
    if (sequentialErrors[action] && !sequentialErrors[action].required) {
      throw new Error(sequentialErrors[action].msg);
    }
    // Double-punch check for "Out" actions ("In" actions are checked above for all users)
    if (!isActionIn && existingValue) {
      throw new Error(`"${action}" already punched today.`);
    }
  }

  if (isAdmin && (puncherEmail !== userEmail || adminTimestamp)) { 
    adherenceSheet.getRange(row, 15).setValue("Yes");
    adherenceSheet.getRange(row, 21).setValue(puncherEmail);
  }

  // === SAVE PUNCH ===
  adherenceSheet.getRange(row, col).setValue(nowTimestamp);
  logsSheet.appendRow([new Date(), userName, userEmail, action, nowTimestamp]);

  // --- START: MODIFICATION FOR REQUEST 2 (Fix Exceeding Bug) ---
  // Replaced the buggy actionKey logic with this switch statement.
  // This correctly updates the local 'punches' object with the new Date object,
  // which is required for the exceed calculations to run immediately.
  switch(action) {
    case "Login": punches.login = nowTimestamp; break;
    case "First Break In": punches.firstBreakIn = nowTimestamp; break;
    case "First Break Out": punches.firstBreakOut = nowTimestamp; break;
    case "Lunch In": punches.lunchIn = nowTimestamp; break;
    case "Lunch Out": punches.lunchOut = nowTimestamp; break;
    case "Last Break In": punches.lastBreakIn = nowTimestamp; break;
    case "Last Break Out": punches.lastBreakOut = nowTimestamp; break;
    case "Logout": punches.logout = nowTimestamp; break;
  }
  // --- END: MODIFICATION FOR REQUEST 2 ---

  // === DATE-AWARE SHIFT METRICS (Now uses objects from step 5) ===
  
  if (!shiftStartDateObj) {
    throw new Error(`Could not parse Shift Start Time ("${shiftStartStr}"). Please check the schedule.`);
  }
  
  if (action === "Login" || punches.login) {
    const loginTime = (action === "Login") ? nowTimestamp : punches.login;
    const diff = timeDiffInSeconds(shiftStartDateObj, loginTime);
    
    const timeFormat = "HH:mm";
    const scheduledTime = Utilities.formatDate(shiftStartDateObj, timeZone, timeFormat);
    const punchTime = Utilities.formatDate(loginTime, timeZone, timeFormat);

    if (action === "Login") {
      if (diff > (4 * 60 * 60)) {
        throw new Error(`Login is over 4 hours late (Shift: ${scheduledTime}, Punch: ${punchTime}). Please check your shift schedule or contact your manager.`);
      }
      if (diff < -(2 * 60 * 60)) {
        throw new Error(`Login is over 2 hours early (Shift: ${scheduledTime}, Punch: ${punchTime}). Please check your shift schedule or contact your manager.`);
      }
    }
    adherenceSheet.getRange(row, 11).setValue(diff > 0 ? diff : 0);
  }

  if (action === "Logout" || punches.logout) {
    if (!shiftEndDateObj) {
      throw new Error(`Could not parse Shift End Time ("${shiftEndStr}"). Please check the schedule.`);
    }
    const logoutTime = (action === "Logout") ? nowTimestamp : punches.logout;
    const diff = timeDiffInSeconds(shiftEndDateObj, logoutTime);
    if (diff > 0) {
      adherenceSheet.getRange(row, 12).setValue(diff);
      adherenceSheet.getRange(row, 13).setValue(0);
    } else {
      adherenceSheet.getRange(row, 12).setValue(0);
      adherenceSheet.getRange(row, 13).setValue(Math.abs(diff));
    }
  }

  // === BREAK EXCEED CALCULATIONS ===
  let exceedMsg = "No";
  let duration = 0;
  let diff = 0;
  try {
    duration = timeDiffInSeconds(punches.firstBreakIn, punches.firstBreakOut);
    diff = duration - PLANNED_BREAK_SECONDS;
    if (diff > 0 && duration > 0) exceedMsg = diff; else exceedMsg = "No";
    adherenceSheet.getRange(row, 17).setValue(exceedMsg);
    
    duration = timeDiffInSeconds(punches.lunchIn, punches.lunchOut);
    diff = duration - PLANNED_LUNCH_SECONDS;
    if (diff > 0 && duration > 0) exceedMsg = diff; else exceedMsg = "No";
    adherenceSheet.getRange(row, 18).setValue(exceedMsg);

    duration = timeDiffInSeconds(punches.lastBreakIn, punches.lastBreakOut);
    diff = duration - PLANNED_BREAK_SECONDS;
    if (diff > 0 && duration > 0) exceedMsg = diff; else exceedMsg = "No";
    adherenceSheet.getRange(row, 19).setValue(exceedMsg);
  } catch (e) {
    logsSheet.appendRow([new Date(), userName, userEmail, "Break Exceed Error", e.message]);
  }

  return `${userName}: ${action} recorded at ${Utilities.formatDate(nowTimestamp, timeZone, "HH:mm:ss")}`;
}


// REPLACE this function
// ================= SCHEDULE RANGE SUBMIT FUNCTION =================
function submitScheduleRange(puncherEmail, userEmail, userName, startDateStr, endDateStr, startTime, endTime, leaveType) {
  const ss = getSpreadsheet();
const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
  const userData = getUserDataFromDb(dbSheet);
  const puncherRole = userData.emailToRole[puncherEmail] || 'agent';
  const timeZone = Session.getScriptTimeZone();
if (puncherRole !== 'admin' && puncherRole !== 'superadmin') {
    throw new Error("Permission denied. Only admins can submit schedules.");
}
  
  const scheduleSheet = getOrCreateSheet(ss, SHEET_NAMES.schedule);
  const scheduleData = scheduleSheet.getDataRange().getValues();
  const logsSheet = getOrCreateSheet(ss, SHEET_NAMES.logs);
const userScheduleMap = {};
  for (let i = 1; i < scheduleData.length; i++) {
    // *** MODIFIED: Read Email from Col G (index 6) ***
    const rowEmail = scheduleData[i][6];
// *** MODIFIED: Read Date from Col B (index 1) ***
    const rowDateRaw = scheduleData[i][1];
if (rowEmail && rowDateRaw && rowEmail.toLowerCase() === userEmail) {
      const rowDate = new Date(rowDateRaw);
const rowDateStr = Utilities.formatDate(rowDate, timeZone, "MM/dd/yyyy");
      userScheduleMap[rowDateStr] = i + 1;
}
  }
  
  const startDate = new Date(startDateStr);
  const endDate = new Date(endDateStr);
let currentDate = new Date(startDate);
  let daysProcessed = 0;
  let daysUpdated = 0;
  let daysCreated = 0;
const oneDayInMs = 24 * 60 * 60 * 1000;
  
  currentDate = new Date(currentDate.valueOf() + currentDate.getTimezoneOffset() * 60000);
const finalDate = new Date(endDate.valueOf() + endDate.getTimezoneOffset() * 60000);
  
  while (currentDate <= finalDate) {
    const currentDateStr = Utilities.formatDate(currentDate, timeZone, "MM/dd/yyyy");
// *** NEW: Auto-calculate shift end date for overnight shifts ***
    let shiftEndDate = new Date(currentDate);
// Start with the same date
    if (startTime && endTime) {
      const startDateTime = createDateTime(currentDate, startTime);
const endDateTime = createDateTime(currentDate, endTime);
      if (endDateTime <= startDateTime) {
        shiftEndDate.setDate(shiftEndDate.getDate() + 1);
// It's the next day
      }
    }
    // *** END NEW ***

    const result = updateOrAddSingleSchedule(
      scheduleSheet, userScheduleMap, logsSheet,
      userEmail, userName, 
      currentDate, // This is StartDate (Col B)
      shiftEndDate, // *** NEW: This is EndDate (Col D) ***
      currentDateStr, 
      startTime, endTime, leaveType, puncherEmail
    );
if (result === "UPDATED") daysUpdated++;
    if (result === "CREATED") daysCreated++;
    
    daysProcessed++;
    currentDate.setTime(currentDate.getTime() + oneDayInMs);
}
  
  if (daysProcessed === 0) {
    throw new Error("No dates were processed. Check date range.");
}
  
  return `Schedule submission complete for ${userName}. Days processed: ${daysProcessed} (Updated: ${daysUpdated}, Created: ${daysCreated}).`;
}

// REPLACE this function
// (Helper for above)
function updateOrAddSingleSchedule(scheduleSheet, userScheduleMap, logsSheet, userEmail, userName, shiftStartDate, shiftEndDate, targetDateStr, startTime, endTime, leaveType, puncherEmail) {
  
  const existingRow = userScheduleMap[targetDateStr];
  let startTimeObj = startTime ? new Date(`1899-12-30T${startTime}`) : "";
  let endTimeObj = endTime ? new Date(`1899-12-30T${endTime}`) : "";
  
  // *** MODIFIED: Set EndDate (Col D) to null if not "Present" ***
  let endDateObj = (leaveType === 'Present' && endTimeObj) ? shiftEndDate : "";

  // *** MODIFIED: Write to 7 columns ***
  const rowData = [[
    userName,     // A
    shiftStartDate, // B
    startTimeObj,   // C
    endDateObj,     // D
    endTimeObj,     // E
    leaveType,      // F
    userEmail       // G
  ]];

  if (existingRow) {
    // *** MODIFIED: Write 7 columns ***
    scheduleSheet.getRange(existingRow, 1, 1, 7).setValues(rowData);
    logsSheet.appendRow([new Date(), userName, puncherEmail, "Schedule UPDATE", `Set to: ${leaveType}, ${startTime}-${endTime}`]);
    return "UPDATED";
  } else {
    // *** MODIFIED: Append 7 columns ***
    scheduleSheet.appendRow(rowData[0]);
    logsSheet.appendRow([new Date(), userName, puncherEmail, "Schedule CREATE", `Set to: ${leaveType}, ${startTime}-${endTime}`]);
    return "CREATED";
  }
}


// ================= HELPER FUNCTIONS =================

function getShiftDate(dateObj, cutoffHour) {
  if (dateObj.getHours() < cutoffHour) {
    dateObj.setDate(dateObj.getDate() - 1);
  }
  return dateObj;
}

function createDateTime(dateObj, timeStr) {
  if (!timeStr) return null;
  const parts = timeStr.split(':');
  if (parts.length < 2) return null;
  
  const [hours, minutes, seconds] = parts.map(Number);
  if (isNaN(hours) || isNaN(minutes)) return null; 

  const newDate = new Date(dateObj);
  newDate.setHours(hours, minutes, seconds || 0, 0);
  return newDate;
}

// REPLACE this function
function getUserDataFromDb(dbSheet) { 
  const dbData = dbSheet.getDataRange().getValues();
  const nameToEmail = {};
  const emailToName = {};
  const emailToRole = {}; 
  const emailToBalances = {}; 
  const emailToRow = {};
  const emailToSupervisor = {}; 
  const emailToAccountStatus = {};
  const emailToHiringDate = {}; // *** NEW ***
  const userList = [];
  
  for (let i = 1; i < dbData.length; i++) {
    // --- START: Fix 1 (Catch invalid data) ---
    try {
    // --- END: Fix 1 ---

      let [name, email, role, annual, sick, casual, supervisor, accountStatus, hiringDate] = dbData[i];
      
      if (name && email) {
        const cleanName = name.toString().trim();
        const cleanEmail = email.toString().trim().toLowerCase();
        const userRole = (role || 'agent').toString().trim().toLowerCase(); 
        const supervisorEmail = (supervisor || "").toString().trim().toLowerCase(); 
        const userAccountStatus = (accountStatus || "Pending").toString().trim();
        
        const hiringDateObj = parseDate(hiringDate);
        
        nameToEmail[cleanName] = cleanEmail;
        emailToName[cleanEmail] = cleanName;
        emailToRole[cleanEmail] = userRole; 
        emailToRow[cleanEmail] = i + 1; 
        emailToSupervisor[cleanEmail] = supervisorEmail; 
        emailToAccountStatus[cleanEmail] = userAccountStatus;

        // --- START: Fix 2 (Convert Date to String) ---
        const hiringDateStr = convertDateToString(hiringDateObj);
        emailToHiringDate[cleanEmail] = hiringDateStr;
        // --- END: Fix 2 ---
        
        emailToBalances[cleanEmail] = {
          annual: parseFloat(annual) || 0,
          sick: parseFloat(sick) || 0,
          casual: parseFloat(casual) || 0
        };
        
        userList.push({ 
          name: cleanName, 
          email: cleanEmail, 
          role: userRole,
          balances: emailToBalances[cleanEmail],
          supervisor: supervisorEmail,
          accountStatus: userAccountStatus,
          hiringDate: hiringDateStr // --- Fix 2: Use the string version ---
        });
      }

    // --- START: Fix 1 (Catch invalid data) ---
    } catch (e) {
      // If one row fails (e.g., bad date), log it and continue to the next row
      Logger.log(`Failed to process row ${i + 1} in Data Base. Error: ${e.message}`);
    }
    // --- END: Fix 1 ---
  }
  
  userList.sort((a, b) => a.name.localeCompare(b.name)); 
  
  return { 
    nameToEmail, emailToName, emailToRole, emailToBalances, 
    emailToRow, emailToSupervisor, emailToAccountStatus, 
    emailToHiringDate, userList 
  };
}



/**
 * NEW: Finds the latest adherence or other code punch for a user 
 * on a given shift date and determines their current logical status.
 */
function getLatestPunchStatus(userEmail, userName, shiftDate, formattedDate) {
  const ss = getSpreadsheet();
  const adherenceSheet = getOrCreateSheet(ss, SHEET_NAMES.adherence);
  const otherCodesSheet = getOrCreateSheet(ss, SHEET_NAMES.otherCodes);
  
  let lastAdherencePunch = null;
  let lastAdherenceTime = new Date(0);
  let lastOtherPunch = null;
  let lastOtherTime = new Date(0);

  // 1. Check Adherence Tracker
  const adherenceData = adherenceSheet.getDataRange().getValues();
  for (let i = adherenceData.length - 1; i > 0; i--) {
    const row = adherenceData[i];
    if (row[1] === userName && Utilities.formatDate(new Date(row[0]), Session.getScriptTimeZone(), "MM/dd/yyyy") === formattedDate) {
      // Found the user's row for today. Now find their last punch.
      // Columns: Login (2), B1-In(3), B1-Out(4), L-In(5), L-Out(6), B2-In(7), B2-Out(8), Logout(9)
      const punches = [
        { name: "Login", time: row[2] },
        { name: "First Break In", time: row[3] },
        { name: "First Break Out", time: row[4] },
        { name: "Lunch In", time: row[5] },
        { name: "Lunch Out", time: row[6] },
        { name: "Last Break In", time: row[7] },
        { name: "Last Break Out", time: row[8] },
        { name: "Logout", time: row[9] }
      ];

      for (const punch of punches) {
        if (punch.time instanceof Date && punch.time > lastAdherenceTime) {
          lastAdherenceTime = punch.time;
          lastAdherencePunch = punch.name;
        }
      }
      break; // Found the right row, no need to search further
    }
  }

  // 2. Check Other Codes
  const otherCodesData = otherCodesSheet.getDataRange().getValues();
  for (let i = otherCodesData.length - 1; i > 0; i--) {
    const row = otherCodesData[i];
    const rowShiftDate = getShiftDate(new Date(row[0]), SHIFT_CUTOFF_HOUR);
    if (row[1] === userName && Utilities.formatDate(rowShiftDate, Session.getScriptTimeZone(), "MM/dd/yyyy") === formattedDate) {
      // Check both "In" (col 3) and "Out" (col 4)
      const timeIn = row[3];
      const timeOut = row[4];
      const code = row[2];

      if (timeIn instanceof Date && timeIn > lastOtherTime) {
        lastOtherTime = timeIn;
        lastOtherPunch = `${code} In`;
      }
      if (timeOut instanceof Date && timeOut > lastOtherTime) {
        lastOtherTime = timeOut;
        lastOtherPunch = `${code} Out`;
      }
    }
  }

  // 3. Compare and determine final status
  let lastPunchName = null;
  let lastPunchTime = null;

  if (lastAdherenceTime > lastOtherTime) {
    lastPunchName = lastAdherencePunch;
    lastPunchTime = lastAdherenceTime;
  } else {
    lastPunchName = lastOtherPunch;
    lastPunchTime = lastOtherTime;
  }

  if (!lastPunchName) {
    return null; // No punches found
  }

  // 4. Determine logical *current* status
  let currentStatus = "Logged Out";
  if (lastPunchName.endsWith(" In")) {
    // "Login", "First Break In", "Lunch In", "Coaching In", etc.
    currentStatus = lastPunchName.replace(" In", "");
    if (currentStatus === "Login") {
       currentStatus = "Logged In";
    } else {
       currentStatus = `On ${currentStatus}`;
    }
  } else if (lastPunchName.endsWith(" Out") && lastPunchName !== "Logout") {
    // "First Break Out", "Lunch Out", "Coaching Out", etc.
    currentStatus = "Logged In"; // After a break/meeting ends, status is Logged In
  } else if (lastPunchName === "Logout") {
    currentStatus = "Logged Out";
  }

  return {
    status: currentStatus,
    time: convertDateToString(lastPunchTime) // Use existing helper
  };
}

// REPLACE this function in your code.gs file
function logOtherCode(sheet, userName, action, nowTimestamp, adminEmail) { 
  const [code, type] = action.split(" ");
  const data = sheet.getDataRange().getValues();
  const timeZone = Session.getScriptTimeZone();
  
  const shiftDate = getShiftDate(new Date(nowTimestamp), SHIFT_CUTOFF_HOUR);
  const dateStr = Utilities.formatDate(shiftDate, timeZone, "MM/dd/yyyy");

  if (type === "In") {
    
    // --- START: MODIFICATION FOR REQUEST 1 (Prevent Double "In" Punch for Other Codes) ---
    // This check applies to EVERYONE, including admins using the main punch button.
    let alreadyPunchedIn = false;
    for (let i = data.length - 1; i > 0; i--) {
        const [rowDateRaw, rowName, rowCode, rowIn] = data[i];
        if (!rowDateRaw || !rowName || !rowIn) continue; // Skip rows without an "In" punch
        
        const rowShiftDate = getShiftDate(new Date(rowDateRaw), SHIFT_CUTOFF_HOUR);
        const rowDateStr = Utilities.formatDate(rowShiftDate, timeZone, "MM/dd/yyyy");

        if (rowName === userName && rowDateStr === dateStr && rowCode === code) {
            // Found an "In" punch for this code, user, and date.
            alreadyPunchedIn = true;
            break;
        }
    }
    if (alreadyPunchedIn) {
        throw new Error(`Error: "${action}" has already been punched today.`);
    }
    // --- END: MODIFICATION FOR REQUEST 1 ---

    if (adminEmail) { 
       sheet.appendRow([nowTimestamp, userName, code, nowTimestamp, "", "", adminEmail]);
       return `${userName}: ${action} recorded at ${Utilities.formatDate(nowTimestamp, timeZone, "HH:mm:ss")}.`;
    }
    
    // This loop now only checks for sequential errors (In without Out) for non-admins
    for (let i = data.length - 1; i > 0; i--) {
      const [rowDateRaw, rowName, rowCode, rowIn, rowOut] = data[i];
      if (!rowDateRaw || !rowName) continue;
      
      const rowShiftDate = getShiftDate(new Date(rowDateRaw), SHIFT_CUTOFF_HOUR);
      const rowDateStr = Utilities.formatDate(rowShiftDate, timeZone, "MM/dd/yyyy");
      if (rowName === userName && rowDateStr === dateStr && rowCode === code && rowIn && !rowOut) { 
        throw new Error(`You must punch "${code} Out" before punching "In" again.`);
      }
    }
    sheet.appendRow([nowTimestamp, userName, code, nowTimestamp, "", "", adminEmail || ""]);

  } else if (type === "Out") {
    let matchingInPunch = null;
    let matchingInRow = -1;
    for (let i = data.length - 1; i > 0; i--) {
      const [rowDateRaw, rowName, rowCode, rowIn, rowOut] = data[i];
      if (!rowDateRaw || !rowName || !rowIn) continue;
      
      const rowShiftDate = getShiftDate(new Date(rowDateRaw), SHIFT_CUTOFF_HOUR);
      const rowDateStr = Utilities.formatDate(rowShiftDate, timeZone, "MM/dd/yyyy");
      if (rowName === userName && rowDateStr === dateStr && rowCode === code && rowIn && !rowOut) { 
        matchingInPunch = rowIn; // This is a Date object
        matchingInRow = i + 1;
        break;
      }
    }
    
    if (matchingInPunch) {
      const duration = timeDiffInSeconds(matchingInPunch, nowTimestamp);
      sheet.getRange(matchingInRow, 5).setValue(nowTimestamp);
      sheet.getRange(matchingInRow, 6).setValue(duration);
      if (adminEmail) {
        sheet.getRange(matchingInRow, 7).setValue(adminEmail);
      }
      return `${userName}: ${action} recorded. Duration: ${Math.round(duration/60)} mins.`;
    } else {
      if (adminEmail) { 
        sheet.appendRow([nowTimestamp, userName, code, "", nowTimestamp, 0, adminEmail]);
        return `${userName}: ${action} (Out) recorded without matching In.`;
      }
      throw new Error(`You must punch "${code} In" first.`);
    }
  }
  return `${userName}: ${action} recorded at ${Utilities.formatDate(nowTimestamp, timeZone, "HH:mm:ss")}.`; 
}

// (No Change)
function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

// (No Change)
function findOrCreateRow(sheet, userName, shiftDate, formattedDate) { 
  const data = sheet.getDataRange().getValues();
  const timeZone = Session.getScriptTimeZone();
  let row = -1;
  for (let i = 1; i < data.length; i++) {
    const rowDate = new Date(data[i][0]);
    const rowUser = data[i][1]; 
    if (
      rowUser && 
      rowUser.toString().toLowerCase() === userName.toLowerCase() && 
      Utilities.formatDate(rowDate, timeZone, "MM/dd/yyyy") === formattedDate
    ) {
      row = i + 1;
      break;
    }
  }

  if (row === -1) {
    row = sheet.getLastRow() + 1;
    sheet.getRange(row, 1).setValue(shiftDate);
    sheet.getRange(row, 2).setValue(userName); 
  }
  return row;
}

// REPLACE this function
function getOrCreateSheet(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (name === SHEET_NAMES.database) {
      // MODIFIED: Added "HiringDate" as the 9th column (index I)
      sheet.getRange("A1:I1").setValues([["User Name", "Email", "Role", "Annual Balance", "Sick Balance", "Casual Balance", "SupervisorEmail", "AccountStatus", "HiringDate"]]);
      sheet.getRange("I:I").setNumberFormat("yyyy-mm-dd"); // Format the hiring date column
    } else if (name === SHEET_NAMES.schedule) {
      // *** REVERTED TO 7-COLUMN LAYOUT TO MATCH YOUR DATA ***
      sheet.getRange("A1:G1").setValues([["Name", "StartDate", "ShiftStartTime", "EndDate", "ShiftEndTime", "LeaveType", "agent email"]]);
      // *** NEW: Format all date/time columns ***
      sheet.getRange("B:B").setNumberFormat("mm/dd/yyyy");
      sheet.getRange("C:C").setNumberFormat("hh:mm");
      sheet.getRange("D:D").setNumberFormat("mm/dd/yyyy");
      sheet.getRange("E:E").setNumberFormat("hh:mm");
    } else if (name === SHEET_NAMES.adherence) {
      sheet.getRange("A1:U1").setValues([[ 
        "Date", "User Name", "Login", "First Break In", "First Break Out", "Lunch In", "Lunch Out", 
        "Last Break In", "Last Break Out", "Logout", "Tardy (Seconds)", "Overtime (Seconds)", "Early Leave (Seconds)",
        "Leave Type", "Admin Audit", "—", "1st Break Exceed", "Lunch Exceed", "Last Break Exceed", "Absent", "Admin Code"
      ]]);
    } else if (name === SHEET_NAMES.logs) {
      sheet.getRange("A1:E1").setValues([["Timestamp", "User Name", "Email", "Action", "Time"]]);
    } else if (name === SHEET_NAMES.otherCodes) { 
      sheet.getRange("A1:G1").setValues([["Date", "User Name", "Code", "Time In", "Time Out", "Duration (Seconds)", "Admin Audit (Email)"]]);
    } else if (name === SHEET_NAMES.leaveRequests) { 
      sheet.getRange("A1:N1").setValues([[ // <-- MODIFIED TO N1
        "RequestID", "Status", "RequestedByEmail", "RequestedByName", 
        "LeaveType", "StartDate", "EndDate", "TotalDays", "Reason", 
        "ActionDate", "ActionBy", "SupervisorEmail", "ActionReason",
        "SickNoteURL" // <-- ADDED THIS
      ]]);
      sheet.getRange("F:G").setNumberFormat("mm/dd/yyyy");
      sheet.getRange("J:J").setNumberFormat("mm/dd/yyyy");
    } else if (name === SHEET_NAMES.coaching_OLD) {
      sheet.getRange("A1:R1").setValues([[
        "SessionID", "Status", "AgentEmail", "AgentName", "CoachEmail", "CoachName",
        "WeekNumber", "SessionDate", "AreaOfConcern", "RootCause", "CoachingTopic",
        "ActionsTaken", "AgentFeedback", "FollowUpPlan", "NextReviewDate",
        "QA_ID", "QA_Score", "LoggedByAdmin"
      ]]);
      sheet.getRange("H:H").setNumberFormat("mm/dd/yyyy");
      sheet.getRange("O:O").setNumberFormat("mm/dd/yyyy");
    } 
    else if (name === SHEET_NAMES.coachingSessions) { 
      sheet.getRange("A1:M1").setValues([[ 
        "SessionID", "AgentEmail", "AgentName", "CoachEmail", "CoachName",
        "SessionDate", "WeekNumber", "OverallScore", "FollowUpComment", "SubmissionTimestamp",
        "FollowUpDate", "FollowUpStatus", "AgentAcknowledgementTimestamp"
      ]]);
      sheet.getRange("F:F").setNumberFormat("mm/dd/yyyy");
      sheet.getRange("J:J").setNumberFormat("mm/dd/yyyy hh:mm:ss");
      sheet.getRange("K:K").setNumberFormat("mm/dd/yyyy");
      sheet.getRange("M:M").setNumberFormat("mm/dd/yyyy hh:mm:ss");
    } else if (name === SHEET_NAMES.coachingScores) { 
      sheet.getRange("A1:E1").setValues([[
        "SessionID", "Category", "Criteria", "Score", "Comment"
      ]]);
    } 
    else if (name === SHEET_NAMES.coachingTemplates) {
      sheet.getRange("A1:D1").setValues([[
        "TemplateName", "Category", "Criteria", "Status"
      ]]);
      sheet.getRange("F1").setValue("Manually add your template criteria here, or use the 'Coaching Admin' tab. Use 'Active' in the Status column to make them appear.");
      sheet.appendRow(["Default", "Greeting", "Agent greeted customer", "Active"]);
      sheet.appendRow(["Default", "Greeting", "Agent confirmed name", "Active"]);
      sheet.appendRow(["Default", "Sales Process", "Agent offered solution", "Active"]);
      sheet.appendRow(["Default", "Sales Process", "Agent handled objections", "Active"]);
      sheet.appendRow(["Default", "Closing", "Agent confirmed next steps", "Active"]);
    }
    else if (name === SHEET_NAMES.pendingRegistrations) {
      sheet.getRange("A1:F1").setValues([[
        "RequestID", "UserEmail", "UserName", "SelectedSupervisorEmail", "Status", "RequestTimestamp"
      ]]);
      sheet.getRange("F:F").setNumberFormat("mm/dd/yyyy hh:mm:ss");
    }
    else if (name === SHEET_NAMES.movementRequests) {
      sheet.getRange("A1:J1").setValues([[
        "MovementID", "Status", "UserToMoveEmail", "UserToMoveName", 
        "FromSupervisorEmail", "ToSupervisorEmail", 
        "RequestTimestamp", "ActionTimestamp", "ActionByEmail", "RequestedByEmail"
      ]]);
      sheet.getRange("G:H").setNumberFormat("mm/dd/yyyy hh:mm:ss");
    }
    else if (name === SHEET_NAMES.announcements) {
      sheet.getRange("A1:E1").setValues([[
        "AnnouncementID", "Content", "Status", "CreatedByEmail", "Timestamp"
      ]]);
      sheet.getRange("E:E").setNumberFormat("mm/dd/yyyy hh:mm:ss");
    }
    else if (name === SHEET_NAMES.roleRequests) {
      sheet.getRange("A1:J1").setValues([[
        "RequestID", "UserEmail", "UserName", "CurrentRole", "RequestedRole", "Justification", 
        "RequestTimestamp", "Status", "ActionByEmail", "ActionTimestamp"
      ]]);
      sheet.getRange("G:G").setNumberFormat("mm/dd/yyyy hh:mm:ss");
      sheet.getRange("J:J").setNumberFormat("mm/dd/yyyy hh:mm:ss");
    }
  }
  if (name === SHEET_NAMES.adherence) {
    sheet.getRange("C:J").setNumberFormat("hh:mm:ss");
  }
  if (name === SHEET_NAMES.otherCodes) {
    sheet.getRange("D:E").setNumberFormat("hh:mm:ss");
  }
  return sheet;
}


// (No Change)
function timeDiffInSeconds(start, end) {
  if (!start || !end || !(start instanceof Date) || !(end instanceof Date)) {
    return 0;
  }
  return Math.round((end.getTime() - start.getTime()) / 1000);
}


// ================= DAILY AUTO-LOG FUNCTION =================
function dailyLeaveSweeper() {
  const ss = getSpreadsheet();
  const scheduleSheet = getOrCreateSheet(ss, SHEET_NAMES.schedule);
  const adherenceSheet = getOrCreateSheet(ss, SHEET_NAMES.adherence);
  const logsSheet = getOrCreateSheet(ss, SHEET_NAMES.logs);
  const timeZone = Session.getScriptTimeZone();
  // 1. Define the 7-day lookback period
  const lookbackDays = 7;
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const endDate = new Date(today); // Today
  endDate.setDate(endDate.getDate() - 1); // End date is yesterday

  const startDate = new Date(endDate);
  startDate.setDate(startDate.getDate() - (lookbackDays - 1)); // Start date is 7 days ago

  const startDateStr = Utilities.formatDate(startDate, timeZone, "MM/dd/yyyy");
  const endDateStr = Utilities.formatDate(endDate, timeZone, "MM/dd/yyyy");

  Logger.log(`Starting dailyLeaveSweeper for date range: ${startDateStr} to ${endDateStr}`);
  // 2. Get all Adherence rows for the past 7 days and create a lookup Set
  const allAdherence = adherenceSheet.getDataRange().getValues();
  const adherenceLookup = new Set();
  for (let i = 1; i < allAdherence.length; i++) {
    try {
      const rowDate = new Date(allAdherence[i][0]);
      if (rowDate >= startDate && rowDate <= endDate) {
        const rowDateStr = Utilities.formatDate(rowDate, timeZone, "MM/dd/yyyy");
        const userName = allAdherence[i][1].toString().trim().toLowerCase();
        adherenceLookup.add(`${userName}:${rowDateStr}`);
      }
    } catch (e) {
      Logger.log(`Skipping adherence row ${i+1}: ${e.message}`);
    }
  }
  Logger.log(`Found ${adherenceLookup.size} existing adherence records in the date range.`);
  // 3. Get all Schedules and loop through them
  const allSchedules = scheduleSheet.getDataRange().getValues();
  let missedLogs = 0;
  for (let i = 1; i < allSchedules.length; i++) {
    try {
      // *** THIS LINE IS THE FIX ***
      // It now correctly reads all 7 columns, matching your sheet structure.
      const [schName, schDate, schStart, schEndDate, schEndTime, schLeave, schEmail] = allSchedules[i];
      // *** END OF FIX ***

      const leaveType = (schLeave || "").toString().trim(); // schLeave is now correctly column F (index 5)

      // This logic is now correct because schLeave and schEmail are from the right columns
      if (leaveType === "" || !schName || !schEmail) {
        continue;
      }

     const schDateObj = parseDate(schDate);

      if (schDateObj && schDateObj >= startDate && schDateObj <= endDate) {
        const schDateStr = Utilities.formatDate(schDateObj, timeZone, "MM/dd/yyyy");
        const userName = schName.toString().trim();
        const userNameLower = userName.toLowerCase();

        const lookupKey = `${userNameLower}:${schDateStr}`;
        // 4. Check if this user is *already* in the Adherence sheet
        if (adherenceLookup.has(lookupKey)) {
          continue; // We found them, so skip
        }

        // 5. We found a missed user!
        Logger.log(`Found missed user: ${userName} for ${schDateStr}. Logging: ${leaveType}`);

        const row = findOrCreateRow(adherenceSheet, userName, schDateObj, schDateStr);
        // *** MODIFIED for Request 3: Mark "Present" as "Absent" ***
        if (leaveType.toLowerCase() === "present") {
          adherenceSheet.getRange(row, 14).setValue("Absent"); // Set Leave Type to Absent
          adherenceSheet.getRange(row, 20).setValue("Yes"); // Set Absent flag to Yes (Col T)
          logsSheet.appendRow([new Date(), userName, schEmail, "Auto-Log Absent", "User was 'Present' but did not punch in."]);
        } else {
          adherenceSheet.getRange(row, 14).setValue(leaveType); // Log Sick, Annual, etc.
          if (leaveType.toLowerCase() === "absent") {
            adherenceSheet.getRange(row, 20).setValue("Yes"); // Set Absent flag (Col T)
          }
          logsSheet.appendRow([new Date(), userName, schEmail, "Auto-Log Leave", leaveType]);
        }

        missedLogs++;
        adherenceLookup.add(lookupKey); // Add to lookup so we don't process again
      }
    } catch (e) {
      Logger.log(`Skipping schedule row ${i+1}: ${e.message}`);
    }
  }

  Logger.log(`dailyLeaveSweeper finished. Logged ${missedLogs} missed users.`);
}

// ================= LEAVE REQUEST FUNCTIONS =================

// (Helper - No Change)
function convertDateToString(dateObj) {
  if (dateObj instanceof Date && !isNaN(dateObj)) {
    return dateObj.toISOString(); // "2025-11-06T18:30:00.000Z"
  }
  return null; // Return null if it's not a valid date
}

// (No Change)
function getMyRequests(userEmail) {
  const ss = getSpreadsheet();
  const reqSheet = getOrCreateSheet(ss, SHEET_NAMES.leaveRequests);
  const allData = reqSheet.getDataRange().getValues();
  const timeZone = Session.getScriptTimeZone();
  const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
  const userData = getUserDataFromDb(dbSheet);
  const myRequests = [];
  Logger.log(`getMyRequests: Searching for email: ${userEmail}`);

  for (let i = allData.length - 1; i > 0; i--) { 
    const row = allData[i];
    if (String(row[2] || "").trim().toLowerCase() === userEmail) {
      try { 
        const startDate = new Date(row[5]);
        const endDate = new Date(row[6]);
        const requestedDateNum = Number(row[0].split('_')[1]);
        Logger.log(`getMyRequests: Found match for ${userEmail} at row ${i+1}`);


        if (isNaN(startDate.getTime()) || isNaN(endDate.getTime()) || isNaN(requestedDateNum)) {
          Logger.log(`Skipping Row ${i+1}. It contains invalid date data.`);
          continue; 
        }

        const supervisorEmail = row[11];
        myRequests.push({
          requestID: row[0],
          status: row[1],
          leaveType: row[4],
          startDate: convertDateToString(startDate),
          endDate: convertDateToString(endDate),
          totalDays: row[7],
          reason: row[8],
          requestedDate: convertDateToString(new Date(requestedDateNum)),
          supervisorName: userData.emailToName[supervisorEmail] || supervisorEmail,
          actionDate: convertDateToString(new Date(row[9])), // ActionDate
          actionBy: userData.emailToName[row[10]] || row[10], // ActionBy
          actionByReason: row[12] || "",
          sickNoteUrl: row[13] || "" // <-- 🛑 ADD THIS LINE
        });
      } catch (e) {
        Logger.log(`CRITICAL ERROR processing row ${i+1} for getMyRequests. Error: ${e.message}`);
        Logger.log(`getMyRequests for ${userEmail}: Found ${myRequests.length} total requests.`);
      }
    }
  }
  Logger.log(`getMyRequests: Returning ${myRequests.length} requests for ${userEmail}`);
  return myRequests;
}

function getAdminLeaveRequests(adminEmail, filter) {
  const ss = getSpreadsheet();
  const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
  const userData = getUserDataFromDb(dbSheet);
  const adminRole = userData.emailToRole[adminEmail] || 'agent';

  if (adminRole !== 'admin' && adminRole !== 'superadmin') {
    return { error: "Permission Denied." };
  }

  // Get my subordinates
  const mySubordinateEmails = new Set(webGetAllSubordinateEmails(adminEmail));

  const reqSheet = getOrCreateSheet(ss, SHEET_NAMES.leaveRequests);
  const allData = reqSheet.getDataRange().getValues();
  const results = [];

  const filterStatus = filter.status.toLowerCase();
  const filterUser = filter.userEmail;

  for (let i = 1; i < allData.length; i++) { // Loop from top to bottom
    const row = allData[i];
    if (!row[0]) continue; // Skip empty rows

    try {
      const requestStatus = (row[1] || "").toString().trim().toLowerCase();
      const requesterEmail = (row[2] || "").toString().trim().toLowerCase();
      const supervisorEmail = (row[11] || "").toString().trim().toLowerCase();

      // --- FILTERING LOGIC ---
      // 1. Filter by Status
      if (filterStatus !== 'all' && requestStatus !== filterStatus) {
        continue;
      }

      // 2. Filter by User
      let userMatch = false;
      if (filterUser === 'ALL_USERS') {
        if (adminRole === 'superadmin') userMatch = true;
      } else if (filterUser === 'ALL_SUBORDINATES') {
        if (mySubordinateEmails.has(requesterEmail)) userMatch = true;
      } else {
        if (requesterEmail === filterUser) userMatch = true;
      }

      if (!userMatch) continue;

      // 3. Security Filter: Admin can only see their subs, Superadmin can see all
      if (adminRole === 'admin' && !mySubordinateEmails.has(requesterEmail)) {
         continue; // This user is not in the admin's hierarchy
      }

      // --- END FILTERS ---

      const startDate = new Date(row[5]);
      const endDate = new Date(row[6]);
      const requestedDateNum = Number(row[0].split('_')[1]);

      if (isNaN(startDate.getTime()) || isNaN(endDate.getTime()) || isNaN(requestedDateNum)) {
        Logger.log(`Skipping Row ${i+1}. It contains invalid date data.`);
        continue; 
      }

      const requesterBalance = userData.emailToBalances[requesterEmail];

      results.push({
        requestID: row[0],
        status: row[1],
        requestedByName: row[3],
        leaveType: row[4],
        startDate: convertDateToString(startDate),
        endDate: convertDateToString(endDate),
        totalDays: row[7],
        reason: row[8], // User's reason
        requestedDate: convertDateToString(new Date(requestedDateNum)),
        actionDate: convertDateToString(new Date(row[9])),
        actionBy: userData.emailToName[row[10]] || row[10],
        supervisorName: userData.emailToName[row[11]] || row[11],
        actionByReason: row[12] || "", // NEW
        requesterBalance: requesterBalance,
        sickNoteUrl: row[13] || "" // <-- 
      });
    } catch (e) {
       Logger.log(`Failed to process row ${i+1} for getAdminLeaveRequests. Error: ${e.message}`);
    }
  }
  return results;
}

// REPLACE this function
function submitLeaveRequest(submitterEmail, request, targetUserEmail) {
  const ss = getSpreadsheet();
  const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
  const userData = getUserDataFromDb(dbSheet);
  
  const isSelfRequest = !targetUserEmail;
  const requestEmail = isSelfRequest ? submitterEmail : targetUserEmail;
  const requestName = userData.emailToName[requestEmail];
  
  if (!requestName) {
    throw new Error(`Could not find user ${requestEmail} in the Data Base.`);
  }
  
  const supervisorEmail = userData.emailToSupervisor[requestEmail];
  if (!supervisorEmail) {
     throw new Error(`Cannot submit request. User ${requestName} does not have a supervisor assigned in the Data Base.`);
  }
  
  const reqSheet = getOrCreateSheet(ss, SHEET_NAMES.leaveRequests);
  const startDate = new Date(request.startDate + 'T00:00:00');
  let endDate;
  if (request.endDate) {
    endDate = new Date(request.endDate + 'T00:00:00');
  } else {
    endDate = startDate; 
  }

  const ONE_DAY_MS = 24 * 60 * 60 * 1000;
  const totalDays = Math.round((endDate.getTime() - startDate.getTime()) / ONE_DAY_MS) + 1;
  if (totalDays < 1) {
    throw new Error("Invalid date range.");
  }
  
  const balanceKey = request.leaveType.toLowerCase(); 
  const userBalances = userData.emailToBalances[requestEmail];
  if (!userBalances || userBalances[balanceKey] === undefined) {
    throw new Error(`Could not find balance information for user ${requestName}. Check the Data Base sheet.`);
  }

  if (userBalances[balanceKey] < totalDays) {
    throw new Error(`Insufficient balance for ${requestName}. User has ${userBalances[balanceKey]} ${request.leaveType} days, but are requesting ${totalDays}.`);
  }

  // --- START NEW FILE UPLOAD LOGIC ---
  let sickNoteUrl = ""; // Default to empty string

  if (request.fileInfo) {
    // A file was included in the request object
    try {
      const folder = DriveApp.getFolderById(SICK_NOTE_FOLDER_ID);
      const fileData = Utilities.base64Decode(request.fileInfo.data);
      const blob = Utilities.newBlob(fileData, request.fileInfo.type, request.fileInfo.name);
      
      // Create a unique name
      const newFileName = `${requestName}_${new Date().toISOString()}_${request.fileInfo.name}`;
      const newFile = folder.createFile(blob).setName(newFileName);
      
      sickNoteUrl = newFile.getUrl();
    } catch (e) {
      Logger.log("File Upload Error: " + e.message);
      throw new Error("Failed to upload sick note file. Please try again.");
    }
  }

  // Backend mandatory check
  if (balanceKey === 'sick' && !sickNoteUrl) {
    throw new Error("A PDF sick note is mandatory for sick leave.");
  }
  // --- END NEW FILE UPLOAD LOGIC ---

  const requestID = `req_${new Date().getTime()}`;
  
  reqSheet.appendRow([
    requestID,
    "Pending",
    requestEmail,
    requestName,
    request.leaveType,
    startDate, 
    endDate,   
    totalDays,
    request.reason,
    "", // ActionDate
    "", // ActionBy
    supervisorEmail,
    "", // ActionReason
    sickNoteUrl // <-- ADDED NEW COLUMN
  ]);
  
  SpreadsheetApp.flush(); 
  
  return `Leave request submitted successfully for ${requestName}.`;
}

function approveDenyRequest(adminEmail, requestID, newStatus, reason) {
  const ss = getSpreadsheet(); 
  const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database); 
  const userData = getUserDataFromDb(dbSheet); 
  const adminRole = userData.emailToRole[adminEmail] || 'agent'; 
  const adminName = userData.emailToName[adminEmail] || adminEmail; 
  
  // --- START MODIFICATION ---
  // Get the admin's full hierarchy, regardless of role (admin/superadmin)
  const mySubordinateEmails = new Set(webGetAllSubordinateEmails(adminEmail)); 

  if (adminRole === 'agent') { // <-- Changed from original check
    throw new Error("Permission denied. Only admins can take this action."); 
  }
  // --- END MODIFICATION ---
  
  const reqSheet = getOrCreateSheet(ss, SHEET_NAMES.leaveRequests); 
  const allData = reqSheet.getDataRange().getValues(); 
  
  for (let i = 1; i < allData.length; i++) { 
    if (allData[i][0] === requestID) { 
      const row = allData[i]; 
      const status = row[1]; 
      
      if (status !== 'Pending') { 
        throw new Error(`This request has already been ${status.toLowerCase()}.`); 
      }
      
      const supervisorEmail = (row[11] || "").toLowerCase(); // Get the email of the supervisor assigned to the request 
      
      // --- START MODIFICATION ---
      // NEW HIERARCHY CHECK:
      // The approver must be the assigned supervisor OR a manager of the assigned supervisor.
      if (supervisorEmail !== adminEmail.toLowerCase() && !mySubordinateEmails.has(supervisorEmail)) { 
        throw new Error("Permission denied. You can only approve requests assigned to you or to a supervisor in your reporting line."); 
      }
      // --- END MODIFICATION ---
      
      const reqEmail = row[2]; 
      const reqName = row[3]; 
      const reqLeaveType = row[4]; 
      const reqStartDate = new Date(row[5]); 
      const reqEndDate = new Date(row[6]); 
      const reqStartDateStr = Utilities.formatDate(reqStartDate, Session.getScriptTimeZone(), "yyyy-MM-dd"); 
      const reqEndDateStr = Utilities.formatDate(reqEndDate, Session.getScriptTimeZone(), "yyyy-MM-dd"); 
      const totalDays = row[7]; 
      let scheduleResult = ""; 
      
      if (newStatus === 'Approved') { 
      
        const balanceKey = reqLeaveType.toLowerCase(); 
        const balanceCol = { annual: 4, sick: 5, casual: 6 }[balanceKey]; 
        
        if (!balanceCol) { 
          throw new Error(`Unknown leave type: ${reqLeaveType}. Balance cannot be deducted.`); 
        }
        
        const userRow = userData.emailToRow[reqEmail]; 
        
        if (!userRow) { 
          throw new Error(`Could not find user ${reqName} in Data Base to deduct balance.`); 
        }
        
        const balanceRange = dbSheet.getRange(userRow, balanceCol); 
        const currentBalance = parseFloat(balanceRange.getValue()) || 0; 
        
        if (currentBalance < totalDays) { 
          throw new Error(`Cannot approve. User only has ${currentBalance} ${reqLeaveType} days, but request is for ${totalDays}.`); 
        }
        
        balanceRange.setValue(currentBalance - totalDays); 
        
        scheduleResult = submitScheduleRange( 
          adminEmail, reqEmail, reqName, 
          reqStartDateStr, reqEndDateStr, 
          "", "", reqLeaveType 
        );
        
        reqSheet.getRange(i + 1, 2).setValue(newStatus); 
        reqSheet.getRange(i + 1, 10).setValue(new Date()); 
        reqSheet.getRange(i + 1, 11).setValue(adminEmail); 
        reqSheet.getRange(i + 1, 13).setValue(reason || ""); 
        
      } else {
        reqSheet.getRange(i + 1, 2).setValue(newStatus); 
        reqSheet.getRange(i + 1, 10).setValue(new Date()); 
        reqSheet.getRange(i + 1, 11).setValue(adminEmail); 
        reqSheet.getRange(i + 1, 13).setValue(reason || ""); 
      }

      if (newStatus === 'Approved') { 
        return `Request approved. ${scheduleResult}`; 
      } else {
        return "Request has been denied."; 
      }
    }
  }
  
  throw new Error("Could not find the request ID."); 
}

// ================= NEW/MODIFIED FUNCTIONS =================

function getAdherenceRange(adminEmail, userNames, startDateStr, endDateStr) {
  const ss = getSpreadsheet();
  const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
  const userData = getUserDataFromDb(dbSheet);
  const adminRole = userData.emailToRole[adminEmail] || 'agent';
  const timeZone = Session.getScriptTimeZone();
  let targetUserNames = [];
  // Security Check: If user is an agent, force userNames to be only them
  if (adminRole === 'agent') {
    const selfName = userData.emailToName[adminEmail];
    if (!selfName) throw new Error("Your user account was not found.");
    targetUserNames = [selfName];
  } else {
    targetUserNames = userNames; // Admin can view the list they provided
  }

  const targetUserSet = new Set(targetUserNames.map(name => name.toLowerCase()));
  const startDate = new Date(startDateStr);
  const endDate = new Date(endDateStr);
  startDate.setHours(0, 0, 0, 0);
  endDate.setHours(23, 59, 59, 999);
  const results = [];

  // *** NEW for Request 3: Get Schedule Data ***
  const scheduleSheet = getOrCreateSheet(ss, SHEET_NAMES.schedule);
  const scheduleData = scheduleSheet.getDataRange().getValues();
  const scheduleMap = {}; // Key: "username:mm/dd/yyyy", Value: "LeaveType"

  for (let i = 1; i < scheduleData.length; i++) {
    const schName = (scheduleData[i][0] || "").toLowerCase();
    // Check against the lowercase name set
    if (targetUserSet.has(schName)) {
      try {
        const schDate = parseDate(scheduleData[i][1]);
        if (schDate >= startDate && schDate <= endDate) {
          const schDateStr = Utilities.formatDate(schDate, timeZone, "MM/dd/yyyy");
          // *** THIS LINE IS THE FIX ***
          // It now reads from index 5 (Column F) instead of 4 (Column E).
          const leaveType = scheduleData[i][5] || "Present";
          // *** END OF FIX ***
          scheduleMap[`${schName}:${schDateStr}`] = leaveType;
        }
      } catch (e) { /* ignore invalid schedule dates */ }
    }
  }
  // *** END NEW ***

  // 1. Get Adherence Data
  const adherenceSheet = getOrCreateSheet(ss, SHEET_NAMES.adherence);
  const adherenceData = adherenceSheet.getDataRange().getValues();
  const resultsLookup = new Set(); // *** NEW: To track found records ***

  for (let i = 1; i < adherenceData.length; i++) {
    const row = adherenceData[i];
    const rowUser = (row[1] || "").toString().trim().toLowerCase();

    if (targetUserSet.has(rowUser)) {
      try {
        const rowDate = new Date(row[0]);
        if (rowDate >= startDate && rowDate <= endDate) {
          results.push({
            date: convertDateToString(row[0]),
            userName: row[1],
            login: convertDateToString(row[2]),
            firstBreakIn: convertDateToString(row[3]),
            firstBreakOut: convertDateToString(row[4]),
            lunchIn: convertDateToString(row[5]),
            lunchOut: convertDateToString(row[6]),
            lastBreakIn: convertDateToString(row[7]),
            lastBreakOut: convertDateToString(row[8]),
            logout: convertDateToString(row[9]),
            tardy: row[10] || 0,
            overtime: row[11] || 0,
            earlyLeave: row[12] || 0,
            leaveType: row[13],
            firstBreakExceed: row[16] || 0,
            lunchExceed: row[17] || 0,
            lastBreakExceed: row[18] || 0,
            otherCodes: [] // This property was unused, but keeping for consistency
          });
          // *** NEW: Add to lookup ***
          const rDateStr = Utilities.formatDate(rowDate, timeZone, "MM/dd/yyyy");
          resultsLookup.add(`${rowUser}:${rDateStr}`);
        }
      } catch (e) {
        Logger.log(`Skipping adherence row ${i+1}. Invalid date. Error: ${e.message}`);
      }
    }
  }

  // *** NEW for Request 3: Fill in missing days ***
  let currentDate = new Date(startDate);
  const oneDayInMs = 24 * 60 * 60 * 1000;
  while (currentDate <= endDate) {
    const currentDateStr = Utilities.formatDate(currentDate, timeZone, "MM/dd/yyyy");
    for (const userName of targetUserNames) {
      const userNameLower = userName.toLowerCase();
      const adherenceKey = `${userNameLower}:${currentDateStr}`;
      // If this user/day is NOT in the adherence sheet
      if (!resultsLookup.has(adherenceKey)) {
        const scheduleKey = `${userNameLower}:${currentDateStr}`;
        const leaveType = scheduleMap[scheduleKey]; // Get schedule (Present, Sick, etc)

        let finalLeaveType = "Day Off"; // Default if no schedule
        if (leaveType) {
          // If scheduled "Present" but no record, they were "Absent"
          finalLeaveType = (leaveType.toLowerCase() === "present") ? "Absent" : leaveType;
        }

        // Add a stub record
        results.push({
          date: convertDateToString(currentDate),
          userName: userName,
          login: null, firstBreakIn: null, firstBreakOut: null, lunchIn: null,
          lunchOut: null, lastBreakIn: null, lastBreakOut: null, logout: null,
          tardy: 0, overtime: 0, earlyLeave: 0,
          leaveType: finalLeaveType,
          firstBreakExceed: 0, lunchExceed: 0, lastBreakExceed: 0,
          otherCodes: []
        });
      }
    }
    currentDate.setTime(currentDate.getTime() + oneDayInMs);
  }
  // *** END NEW ***

  // Sort by date, then by user name
  results.sort((a, b) => {
    if (a.date < b.date) return -1;
    if (a.date > b.date) return 1;
    if (a.userName < b.userName) return -1;
    if (a.userName > b.userName) return 1;
    return 0;
  });
  // This check is no longer needed as we fill stubs
  // if (results.length === 0) { ... }

  return results;
}


// REPLACE this function
function getMySchedule(userEmail) {
  const ss = getSpreadsheet();
  const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
  const userData = getUserDataFromDb(dbSheet);
  const userRole = userData.emailToRole[userEmail] || 'agent';

  const targetEmails = new Set();
  if (userRole === 'agent') {
    targetEmails.add(userEmail);
  } else {
    const subEmails = webGetAllSubordinateEmails(userEmail);
    subEmails.forEach(email => targetEmails.add(email.toLowerCase()));
  }

  const scheduleSheet = getOrCreateSheet(ss, SHEET_NAMES.schedule);
  const scheduleData = scheduleSheet.getDataRange().getValues();
  const timeZone = Session.getScriptTimeZone();
  
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  const nextSevenDays = new Date(today);
  nextSevenDays.setDate(today.getDate() + 7);

  const mySchedule = [];
  for (let i = 1; i < scheduleData.length; i++) {
    const row = scheduleData[i];
    // *** MODIFIED: Read Email from Col G (index 6) ***
    const schEmail = (row[6] || "").toString().trim().toLowerCase(); 
    
    if (targetEmails.has(schEmail)) {
      try {
        // *** MODIFIED: Read Date from Col B (index 1) ***
        const schDate = parseDate(row[1]);
        if (schDate >= today && schDate < nextSevenDays) { 
          
          // *** MODIFIED: Read times/leave from Col C, E, F ***
          let startTime = row[2]; // Col C
          let endTime = row[4];   // Col E
          let leaveType = row[5] || ""; // Col F

          // *** MODIFIED for Request 3: Handle "Day Off" ***
          if (leaveType === "" && !startTime) {
            leaveType = "Day Off";
          } else if (leaveType === "" && startTime) {
            leaveType = "Present"; // Default if times exist but no type
          }
          // *** END MODIFICATION ***
          
          if (startTime instanceof Date) {
            startTime = Utilities.formatDate(startTime, timeZone, "HH:mm");
          }
          if (endTime instanceof Date) {
            endTime = Utilities.formatDate(endTime, timeZone, "HH:mm");
          }
          
          mySchedule.push({
            userName: userData.emailToName[schEmail] || schEmail,
            date: convertDateToString(schDate),
            leaveType: leaveType,
            startTime: startTime,
            endTime: endTime
          });
        }
      } catch(e) {
        Logger.log(`Skipping schedule row ${i+1}. Invalid date. Error: ${e.message}`);
      }
    }
  }
  
  mySchedule.sort((a, b) => {
    const dateA = new Date(a.date);
    const dateB = new Date(b.date);
    if (dateA < dateB) return -1;
    if (dateA > dateB) return 1;
    return a.userName.localeCompare(b.userName);
  });
  return mySchedule;
}


// (No Change)
function adjustLeaveBalance(adminEmail, userEmail, leaveType, amount, reason) {
  const ss = getSpreadsheet();
  const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
  const userData = getUserDataFromDb(dbSheet);
  
  const adminRole = userData.emailToRole[adminEmail] || 'agent';
  if (adminRole !== 'admin' && adminRole !== 'superadmin') {
    throw new Error("Permission denied. Only admins can adjust balances.");
  }
  
  const balanceKey = leaveType.toLowerCase();
  const balanceCol = { annual: 4, sick: 5, casual: 6 }[balanceKey];
  if (!balanceCol) {
    throw new Error(`Unknown leave type: ${leaveType}.`);
  }
  
  const userRow = userData.emailToRow[userEmail];
  const userName = userData.emailToName[userEmail];
  if (!userRow) {
    throw new Error(`Could not find user ${userName} in Data Base.`);
  }
  
  const balanceRange = dbSheet.getRange(userRow, balanceCol);
  const currentBalance = parseFloat(balanceRange.getValue()) || 0;
  const newBalance = currentBalance + amount;
  
  balanceRange.setValue(newBalance);
  
  // Log the adjustment
  const logsSheet = getOrCreateSheet(ss, SHEET_NAMES.logs);
  logsSheet.appendRow([
    new Date(), 
    userName, 
    adminEmail, 
    "Balance Adjustment", 
    `Admin: ${adminEmail} | User: ${userName} | Type: ${leaveType} | Amount: ${amount} | Reason: ${reason} | Old: ${currentBalance} | New: ${newBalance}`
  ]);
  
  return `Successfully adjusted ${userName}'s ${leaveType} balance from ${currentBalance} to ${newBalance}.`;
}

// REPLACE this function
function importScheduleCSV(adminEmail, csvData) {
  const ss = getSpreadsheet();
  const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
  const userData = getUserDataFromDb(dbSheet);
  const adminRole = userData.emailToRole[adminEmail] || 'agent';
  if (adminRole !== 'admin' && adminRole !== 'superadmin') {
    throw new Error("Permission denied. Only admins can import schedules.");
  }
  
  const scheduleSheet = getOrCreateSheet(ss, SHEET_NAMES.schedule);
  const scheduleData = scheduleSheet.getDataRange().getValues();
  const logsSheet = getOrCreateSheet(ss, SHEET_NAMES.logs);
  const timeZone = Session.getScriptTimeZone(); // *** ADDED: Get timezone ***
  
  // Build a map of existing schedules
  const userScheduleMap = {};
  for (let i = 1; i < scheduleData.length; i++) {
    const rowEmail = scheduleData[i][6];
    const rowDateRaw = scheduleData[i][1]; 
    if (rowEmail && rowDateRaw) {
      const email = rowEmail.toLowerCase();
      if (!userScheduleMap[email]) {
        userScheduleMap[email] = {};
      }
      const rowDate = new Date(rowDateRaw);
      const rowDateStr = Utilities.formatDate(rowDate, timeZone, "MM/dd/yyyy");
      userScheduleMap[email][rowDateStr] = i + 1;
    }
  }
  
  let daysUpdated = 0;
  let daysCreated = 0;
  let errors = 0;
  let errorLog = [];

  for (const row of csvData) {
    try {
      // *** MODIFIED: Read new 7-column CSV headers ***
      const userName = row.Name;
      const userEmail = (row['agent email'] || "").toLowerCase();
      
      // *** MODIFIED: Use new parsers ***
      const targetStartDate = parseDate(row.StartDate);
      let startTime = parseCsvTime(row.ShiftStartTime, timeZone);
      const targetEndDate = parseDate(row.EndDate);
      let endTime = parseCsvTime(row.ShiftEndTime, timeZone);
      // *** END MODIFICATION ***
      
      let leaveType = row.LeaveType || "Present";
      
      if (!userName || !userEmail) {
        throw new Error("Missing required field (Name or agent email).");
      }
      
      // *** MODIFIED: Check parsed date ***
      if (!targetStartDate || isNaN(targetStartDate.getTime())) {
        throw new Error(`Invalid or missing StartDate: ${row.StartDate}.`);
      }
      
      // *** NEW: Format startDateStr from parsed date ***
      const startDateStr = Utilities.formatDate(targetStartDate, timeZone, "MM/dd/yyyy");

      // If leave type is not Present, clear times
      if (leaveType.toLowerCase() !== "present") {
        startTime = "";
        endTime = "";
      }

      // *** MODIFIED: Handle parsed EndDate ***
      let finalEndDate;
      if (leaveType.toLowerCase() === "present" && targetEndDate && !isNaN(targetEndDate.getTime())) {
        finalEndDate = targetEndDate; // Use the valid, parsed EndDate
      } else {
        finalEndDate = new Date(targetStartDate); // Default to StartDate
      }
      // *** END MODIFICATION ***

      const emailMap = userScheduleMap[userEmail] || {};
      const result = updateOrAddSingleSchedule(
        scheduleSheet, emailMap, logsSheet,
        userEmail, userName,
        targetStartDate, // shiftStartDate (Col B) - Now a Date object
        finalEndDate,    // shiftEndDate (Col D) - Now a Date object
        startDateStr,    // targetDateStr (for lookup) - Now a formatted string
        startTime, endTime, leaveType, adminEmail // Times are now HH:mm:ss strings
      );
      
      if (result === "UPDATED") daysUpdated++;
      if (result === "CREATED") daysCreated++;

    } catch (e) {
      errors++;
      errorLog.push(`Row ${row.Name}/${row.StartDate}: ${e.message}`);
    }
  }

  if (errors > 0) {
    return `Error: Import complete with ${errors} errors. (Created: ${daysCreated}, Updated: ${daysUpdated}). Errors: ${errorLog.join(' | ')}`;
  }
  
  return `Import successful. Records Created: ${daysCreated}, Records Updated: ${daysUpdated}.`;
}

// REPLACE this function
function getDashboardData(adminEmail, userEmails, date) {
  const ss = getSpreadsheet();
  const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
  const userData = getUserDataFromDb(dbSheet);
  const adminRole = userData.emailToRole[adminEmail] || 'agent';
  if (adminRole !== 'admin' && adminRole !== 'superadmin') {
    throw new Error("Permission denied.");
  }
  
  const timeZone = Session.getScriptTimeZone();
  const targetDate = new Date(date);
  const targetDateStr = Utilities.formatDate(targetDate, timeZone, "MM/dd/yyyy");
  const targetUserSet = new Set(userEmails.map(e => e.toLowerCase()));
  const userStatusMap = {};
  
  const totalAdherenceMetrics = {
    totalTardy: 0, totalEarlyLeave: 0, totalOvertime: 0,
    totalBreakExceed: 0, totalLunchExceed: 0
  };
  
  const userMetricsMap = {}; 
  userEmails.forEach(email => {
    const lEmail = email.toLowerCase();
    const name = userData.emailToName[lEmail] || lEmail;
    // *** MODIFIED for Request 3: Default is now "Day Off" ***
    userStatusMap[lEmail] = "Day Off"; 
    userMetricsMap[name] = {
      name: name, tardy: 0, earlyLeave: 0, overtime: 0,
      breakExceed: 0, lunchExceed: 0
    };
  });

  const usersScheduledToday = new Set(); 

  // 1. Get Today's Schedule
  const scheduleSheet = getOrCreateSheet(ss, SHEET_NAMES.schedule);
  const scheduleData = scheduleSheet.getDataRange().getValues();
  
  for (let i = 1; i < scheduleData.length; i++) {
    const row = scheduleData[i];
    // *** MODIFIED: Read Email from Col G (index 6) ***
    const schEmail = (row[6] || "").toLowerCase();
    
    if (!targetUserSet.has(schEmail)) continue;
    
    // *** MODIFIED: Read Date from Col B (index 1) ***
    const schDate = new Date(row[1]); 
    const schDateStr = Utilities.formatDate(schDate, timeZone, "MM/dd/yyyy");
    
    if (schDateStr === targetDateStr) {
      // *** MODIFIED: Read LeaveType from Col F (index 5) ***
      const leaveType = (row[5] || "").toString().trim().toLowerCase();
      // *** MODIFIED: Read StartTime from Col C (index 2) ***
      const startTime = row[2]; 

      // *** MODIFIED for Request 3: Handle "Day Off" (empty type, empty time) ***
      if (leaveType === "" && !startTime) {
        userStatusMap[schEmail] = "Day Off";
      } else if (leaveType === "present" || (leaveType === "" && startTime)) {
        usersScheduledToday.add(schEmail);
        userStatusMap[schEmail] = "Pending Login";
      } else if (leaveType === "absent") {
        userStatusMap[schEmail] = "Absent";
      } else {
        userStatusMap[schEmail] = "On Leave";
      }
    }
  }
  
  // 2. Get Today's Adherence
  const adherenceSheet = getOrCreateSheet(ss, SHEET_NAMES.adherence);
  const adherenceData = adherenceSheet.getDataRange().getValues();
  
  const otherCodesSheet = getOrCreateSheet(ss, SHEET_NAMES.otherCodes);
  const otherCodesData = otherCodesSheet.getDataRange().getValues();
  const userLastOtherCode = {}; 
  
  for (let i = otherCodesData.length - 1; i > 0; i--) { 
    const row = otherCodesData[i];
    const rowDate = new Date(row[0]);
    const rowShiftDate = getShiftDate(rowDate, SHIFT_CUTOFF_HOUR);
    const rowDateStr = Utilities.formatDate(rowShiftDate, timeZone, "MM/dd/yyyy");
    
    if (rowDateStr === targetDateStr) {
      const userName = row[1];
      const userEmail = userData.nameToEmail[userName];
      
      if (userEmail && targetUserSet.has(userEmail.toLowerCase())) {
        if (!userLastOtherCode[userEmail.toLowerCase()]) { 
          const [code, type] = (row[2] || "").split(" ");
          userLastOtherCode[userEmail.toLowerCase()] = { code: code, type: type };
        }
      }
    }
  }
  
  for (let i = 1; i < adherenceData.length; i++) {
    const row = adherenceData[i];
    const rowDate = new Date(row[0]);
    const rowDateStr = Utilities.formatDate(rowDate, timeZone, "MM/dd/yyyy");
    
    if (rowDateStr === targetDateStr) { 
      const userName = row[1];
      const userEmail = userData.nameToEmail[userName];
      
      if (userEmail && targetUserSet.has(userEmail.toLowerCase())) {
        const lEmail = userEmail.toLowerCase();
        
        if (usersScheduledToday.has(lEmail)) {
          const login = row[2], b1_in = row[3], b1_out = row[4], l_in = row[5],
                l_out = row[6], b2_in = row[7], b2_out = row[8], logout = row[9];
          
          let agentStatus = "Pending Login";
          if (login && !logout) {
            agentStatus = "Logged In";
            const lastOther = userLastOtherCode[lEmail];
            
            if (lastOther && lastOther.type === 'In') {
              agentStatus = "On Break/Other";
            } else {
              if (b1_in && !b1_out) agentStatus = "On Break/Other";
              if (l_in && !l_out) agentStatus = "On Break/Other";
              if (b2_in && !b2_out) agentStatus = "On Break/Other";
            }
          } else if (login && logout) {
            agentStatus = "Logged Out";
          }
          
          userStatusMap[lEmail] = agentStatus;
          usersScheduledToday.delete(lEmail);
        }
        
        // 3. Sum Adherence Metrics
        const tardy = parseFloat(row[10]) || 0;
        const earlyLeave = parseFloat(row[12]) || 0;
        const overtime = parseFloat(row[11]) || 0;
        const breakExceed = (parseFloat(row[16]) || 0) + (parseFloat(row[18]) || 0);
        const lunchExceed = parseFloat(row[17]) || 0;

        totalAdherenceMetrics.totalTardy += tardy;
        totalAdherenceMetrics.totalEarlyLeave += earlyLeave;
        totalAdherenceMetrics.totalOvertime += overtime;
        totalAdherenceMetrics.totalBreakExceed += breakExceed;
        totalAdherenceMetrics.totalLunchExceed += lunchExceed;
        
        if (userMetricsMap[userName]) {
          userMetricsMap[userName].tardy += tardy;
          userMetricsMap[userName].earlyLeave += earlyLeave;
          userMetricsMap[userName].overtime += overtime;
          userMetricsMap[userName].breakExceed += breakExceed;
          userMetricsMap[userName].lunchExceed += lunchExceed;
        }
      }
    }
  }
  
  // 4. Get Pending Leave Requests
  const reqSheet = getOrCreateSheet(ss, SHEET_NAMES.leaveRequests);
  const reqData = reqSheet.getDataRange().getValues();
  const pendingRequests = [];
  
  for (let i = 1; i < reqData.length; i++) {
    const row = reqData[i];
    const reqEmail = (row[2] || "").toLowerCase();
    if (row[1] && row[1].toString().trim().toLowerCase() === 'pending' && targetUserSet.has(reqEmail)) {
      try {
        pendingRequests.push({
          name: row[3], type: row[4], 
          startDate: convertDateToString(new Date(row[5])), days: row[7] 
        });
      } catch (e) {
        Logger.log(`Failed to parse pending request row ${i+1}. Error: ${e.message}`);
      }
    }
  }
  
  const agentStatusList = [];
  for (const email of targetUserSet) {
      const name = userData.emailToName[email] || email;
      // *** MODIFIED for Request 3: Fallback is now "Day Off" ***
      const status = userStatusMap[email] || "Day Off";
      agentStatusList.push({ name: name, status: status });
  }
  agentStatusList.sort((a, b) => a.name.localeCompare(b.name));
  
  const individualAdherenceMetrics = Object.values(userMetricsMap);
  
  return {
    agentStatusList: agentStatusList,
    totalAdherenceMetrics: totalAdherenceMetrics,
    individualAdherenceMetrics: individualAdherenceMetrics,
    pendingRequests: pendingRequests
  };
}

// --- NEW: "My Team" Helper Functions ---
function saveMyTeam(adminEmail, userEmails) {
  try {
    // Uses Google Apps Script's built-in User Properties for saving user-specific settings.
    const userProperties = PropertiesService.getUserProperties();
    userProperties.setProperty('myTeam', JSON.stringify(userEmails));
    return "Successfully saved 'My Team' preference.";
  } catch (e) {
    throw new Error("Failed to save team preferences: " + e.message);
  }
}

function getMyTeam(adminEmail) {
  try {
    const userProperties = PropertiesService.getUserProperties();
    // Getting properties implicitly forces the Google auth dialog if needed.
    const properties = userProperties.getProperties(); 
    const myTeam = properties['myTeam'];
    return myTeam ? JSON.parse(myTeam) : [];
  } catch (e) {
    Logger.log("Failed to load team preferences: " + e.message);
    // Throwing an error here would break the dashboard's initial load. 
    // We return an empty array instead, and let the front-end handle the fallback.
   return [];
  }
}

// --- NEW: Reporting Line Function ---
function updateReportingLine(adminEmail, userEmail, newSupervisorEmail) {
  const ss = getSpreadsheet();
  const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
  const userData = getUserDataFromDb(dbSheet);
  
  const adminRole = userData.emailToRole[adminEmail] || 'agent';
  if (adminRole !== 'admin' && adminRole !== 'superadmin') {
    throw new Error("Permission denied. Only admins can change reporting lines.");
  }
  
  const userName = userData.emailToName[userEmail];
  const newSupervisorName = userData.emailToName[newSupervisorEmail];
  if (!userName) throw new Error(`Could not find user: ${userEmail}`);
  if (!newSupervisorName) throw new Error(`Could not find new supervisor: ${newSupervisorEmail}`);

  const userRow = userData.emailToRow[userEmail];
  const currentUserSupervisor = userData.emailToSupervisor[userEmail];

  // Check for auto-approval
  let canAutoApprove = false;
  if (adminRole === 'superadmin') {
    canAutoApprove = true;
  } else if (adminRole === 'admin') {
    // Check if both the user's current supervisor AND the new supervisor report to this admin
    const currentSupervisorManager = userData.emailToSupervisor[currentUserSupervisor];
    const newSupervisorManager = userData.emailToSupervisor[newSupervisorEmail];
    
    if (currentSupervisorManager === adminEmail && newSupervisorManager === adminEmail) {
      canAutoApprove = true;
    }
  }

  if (!canAutoApprove) {
    // This is where we will build Phase 2 (requesting the change)
    // For now, we will just show a permission error.
    throw new Error("Permission Denied: You do not have authority to approve this change. (This will become a request in Phase 2).");
  }

  // --- Auto-Approval Logic ---
  // Update the SupervisorEmail column (Column G = 7)
  dbSheet.getRange(userRow, 7).setValue(newSupervisorEmail);
  
  // Log the change
  const logsSheet = getOrCreateSheet(ss, SHEET_NAMES.logs);
  logsSheet.appendRow([
    new Date(), 
    userName, 
    adminEmail, 
    "Reporting Line Change", 
    `User: ${userName} moved to Supervisor: ${newSupervisorName} by ${adminEmail}`
  ]);
  
  return `${userName} has been successfully reassigned to ${newSupervisorName}.`;
}

// [START] MODIFICATION 2: Replace _ONE_TIME_FIX_TEMPLATE
/**
 * ==========================================================
 * ONE-TIME MIGRATION SCRIPT
 * Run this function ONCE from the Apps Script editor
 * to build your main template in the Google Sheet.
 * It will only run if the sheet is empty.
 * ==========================================================
 */
function _SETUP_DEFAULT_TEMPLATE() {
  // This data is from your original hard-coded 'qualityCategories' variable
  const qualityCategories = [
    { 
      category: "Greeting & Opening",
      criteria: [
        "Agent greeted the customer professionally and introduced themselves appropriately",
        "Agent confirmed the customer’s name & purpose of the call/chat"
      ]
    },
    {
      category: "Communication Skills & Understanding Needs",
      criteria: [
        "Agent conversed actively without interrupting",
        "Agent asked relevant questions to understand customer needs",
        "Agent acknowledged customer concerns appropriately",
        "Language was clear, understandable, and free of jargon",
        "Agent applied correct hold etiquettes",
        "Tone was confident, professional and engaging"
      ]
    },
    {
      category: "Product Knowledge & providing solution",
      criteria: [
        "Agent demonstrated strong knowledge of Lenovo products/services",
        "Agent offered the right solution based on customer's needs",
        "Agent was able to handle objections confidently & Highlighted Lenovo's competitive advantage"
      ]
    },
    {
      category: "Tools usage and Chat/ Call Logging",
      criteria: [
        "Agent applied correct disposition.",
        "Agent logged the chat with all relevant details in Dynamics 365 B2C"
      ]
    },
    {
      category: "Sales Closing & Call to Action",
      criteria: [
        "Agent clearly stated pricing, offers, and benefits.",
        "Agent confirmed next steps (e.g., sending a quote, scheduling a follow-up)"
      ]
    },
    {
      category: "Process Compliance",
      criteria: [
        "Agent followed Lenovo's sales process & compliance guidelines. OR Agent transfered the chat to the approriate que when applicable"
      ]
    },
    {
      category: "Wrap-Up & Closing",
      criteria: [
        "Agent confirmed if the customer’s query was fully addressed",
        "Agent ended the chat approprietly.",
        "Follow-up commitment created (if applicable)"
      ]
    }
  ];

  const ss = getSpreadsheet();
  const templateSheet = getOrCreateSheet(ss, SHEET_NAMES.coachingTemplates);
  
  // *** NEW CHECK: Only run if the sheet is empty (has 1 row - the header) ***
  if (templateSheet.getLastRow() > 1) {
    Logger.log("CoachingTemplates sheet is not empty. Skipping default template setup.");
    return;
  }
  
  // 1. Clear all old data (but not the header row)
  templateSheet.getRange(2, 1, templateSheet.getLastRow(), 4).clearContent();
  
  const newRows = [];
  const templateName = "Main Template"; // This will be the name in the dropdown
  
  // 2. Build the new rows
  qualityCategories.forEach(cat => {
    cat.criteria.forEach(crit => {
      newRows.push([
        templateName,
        cat.category,
        crit,
        "Active"
      ]);
    });
  });
  
  // 3. Write all new rows at once
  if (newRows.length > 0) {
    templateSheet.getRange(2, 1, newRows.length, 4).setValues(newRows);
  }
  
  Logger.log(`Migration complete! Added ${newRows.length} criteria for 'Main Template'.`);
}
// [END] MODIFICATION 2


/**
 * REPLACES webSetMySupervisor.
 * New user submits their chosen supervisor. This logs it for approval.
 */
function webSubmitSupervisorSelection(supervisorEmail) {
  try {
    const userEmail = Session.getActiveUser().getEmail().toLowerCase();
    const ss = getSpreadsheet();
    const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
    const userData = getUserDataFromDb(dbSheet);
    const regSheet = getOrCreateSheet(ss, SHEET_NAMES.pendingRegistrations);

    const userName = userData.emailToName[userEmail];
    if (!userName) {
      throw new Error("Your user account could not be found.");
    }
    
    // Validate the selection
    const supervisorName = userData.emailToName[supervisorEmail];
    if (!supervisorName) {
      throw new Error("The selected supervisor is not a valid user.");
    }

    // Check for an existing request
    const existing = regSheet.getDataRange().getValues();
    let existingRequestID = null;
    let existingStatus = null;
    for (let i = 1; i < existing.length; i++) {
      if (existing[i][1] === userEmail) {
        existingRequestID = existing[i][0];
        existingStatus = existing[i][4];
        break;
      }
    }

    if (existingStatus === 'Pending') {
      throw new Error("You already have a pending submission.");
    }

    // Log the new request (or update the old 'Denied' one)
    const requestID = existingRequestID || `reg_${new Date().getTime()}`;
    regSheet.appendRow([
      requestID,
      userEmail,
      userName,
      supervisorEmail,
      "Pending",
      new Date()
    ]);

    return "Submission received! Your account is pending final approval by an administrator. Please check back later.";

  } catch (err) {
    Logger.log("webSubmitSupervisorSelection Error: " + err.message);
    return "Error: " + err.message;
  }
}

/**
 * For the pending user to check their own status.
 */
function webGetMyRegistrationStatus() {
  try {
    const userEmail = Session.getActiveUser().getEmail().toLowerCase();
    const regSheet = getOrCreateSheet(getSpreadsheet(), SHEET_NAMES.pendingRegistrations);
    const data = regSheet.getDataRange().getValues();

    for (let i = data.length - 1; i > 0; i--) { // Check newest first
      if (data[i][1] === userEmail) {
        return { status: data[i][4], supervisor: data[i][3] }; // Returns { status: "Pending" } or { status: "Denied" }
      }
    }
    return { status: "New" }; // No submission found
  } catch (e) {
    return { error: e.message };
  }
}

/**
 * For the Superadmin/Admin to load the approval queue.
 * MODIFIED: Allows Admins to see requests for their subordinates.
 */
function webGetPendingRegistrations() {
  try {
    const adminEmail = Session.getActiveUser().getEmail().toLowerCase();
    const ss = getSpreadsheet();
    const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
    const userData = getUserDataFromDb(dbSheet);
    
    // --- MODIFIED PERMISSION CHECK ---
    const adminRole = userData.emailToRole[adminEmail] || 'agent';
    if (adminRole !== 'admin' && adminRole !== 'superadmin') {
      throw new Error("Permission denied. Only admins and superadmins can approve new users.");
    }
    // --- END MODIFICATION ---

    // --- NEW: Get Admin's team for filtering ---
    let myTeamEmails = new Set();
    if (adminRole === 'admin') {
      // webGetAllSubordinateEmails includes the admin's own email
      myTeamEmails = new Set(webGetAllSubordinateEmails(adminEmail));
    }
    // --- END NEW ---

    const regSheet = getOrCreateSheet(ss, SHEET_NAMES.pendingRegistrations);
    const data = regSheet.getDataRange().getValues();
    const pending = [];
    
    // --- MODIFIED LOOP with filtering ---
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const status = row[4];
      const selectedSupervisorEmail = (row[3] || "").toLowerCase();

      if (status === "Pending") {
        
        let canView = false;
        if (adminRole === 'superadmin') {
          // Superadmin can see all
          canView = true;
        } else if (adminRole === 'admin' && myTeamEmails.has(selectedSupervisorEmail)) {
          // Admin can see if the selected supervisor is in their hierarchy
          canView = true;
        }
        
        if (canView) {
          pending.push({
            requestID: row[0],
            userEmail: row[1],
            userName: row[2],
            selectedSupervisorEmail: selectedSupervisorEmail, // Use cleaned variable
            supervisorName: userData.emailToName[selectedSupervisorEmail] || 'Unknown Supervisor',
            timestamp: convertDateToString(row[5])
          });
        }
      }
    }
    // --- END MODIFICATION ---

    return pending.sort((a,b) => new Date(b.timestamp) - new Date(a.timestamp)); // Newest first
  } catch (e) {
    return { error: e.message };
}
}

// REPLACE this function
/**
 * For the Superadmin/Admin to action the request.
 * MODIFIED: Now accepts hiringDateStr
 */
function webApproveDenyRegistration(requestID, userEmail, selectedSupervisorEmail, newStatus, hiringDateStr) {
  try {
    const adminEmail = Session.getActiveUser().getEmail().toLowerCase();
    const ss = getSpreadsheet(); 
    const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database); 
    const userData = getUserDataFromDb(dbSheet);
    
    const adminRole = userData.emailToRole[adminEmail] || 'agent';
    if (adminRole === 'agent') { 
      throw new Error("Permission denied.");
    }
    
    const myTeamEmails = new Set(webGetAllSubordinateEmails(adminEmail));
    if (!myTeamEmails.has(selectedSupervisorEmail.toLowerCase())) { 
      throw new Error("Permission denied. You can only Approve or Deny registrations for users who have selected a supervisor in your reporting line.");
    }

    // 1. Update the PendingRegistrations sheet
    const regSheet = getOrCreateSheet(ss, SHEET_NAMES.pendingRegistrations);
    const regData = regSheet.getDataRange().getValues(); 
    let regRow = -1; 
    for (let i = 1; i < regData.length; i++) { 
      if (regData[i][0] === requestID && regData[i][4] === 'Pending') { 
        regRow = i + 1;
        break; 
      }
    }
    if (regRow === -1) { 
      throw new Error("Could not find the registration request.");
    }
    regSheet.getRange(regRow, 5).setValue(newStatus); // Set Status (Column E) 
    regSheet.getRange(regRow, 1, 1, regSheet.getLastColumn()).setBackground("#f4f4f4");

    // 2. If Approved, update the Data Base
    if (newStatus === 'Approved') {
      // *** NEW: Validate hiring date ***
      if (!hiringDateStr || isNaN(new Date(hiringDateStr).getTime())) {
        throw new Error("A valid Hiring Date is required to approve a user.");
      }
      const hiringDate = new Date(hiringDateStr);
      // *** END NEW ***

      const userDBRow = userData.emailToRow[userEmail];
      if (!userDBRow) { 
        throw new Error(`Could not find user ${userEmail} in Data Base to approve.`);
      }
      dbSheet.getRange(userDBRow, 7).setValue(selectedSupervisorEmail); // Set SupervisorEmail (Col G) 
      dbSheet.getRange(userDBRow, 8).setValue("Active"); // Set AccountStatus (Col H) 
      dbSheet.getRange(userDBRow, 9).setValue(hiringDate); // *** NEW: Set HiringDate (Col I) ***
    }
    
    SpreadsheetApp.flush(); 
    return { success: true, message: `User registration ${newStatus.toLowerCase()}.` };
  } catch (e) {
    Logger.log(`webApproveDenyRegistration Error: ${e.message}`); 
    return { error: e.message };
  }
}
// --- ADD TO THE END OF code.gs ---

// ==========================================================
// === ANNOUNCEMENTS MODULE ===
// ==========================================================

/**
 * Fetches only active announcements for all users.
 */
function webGetAnnouncements() {
  try {
    const ss = getSpreadsheet();
    const sheet = getOrCreateSheet(ss, SHEET_NAMES.announcements);
    const data = sheet.getDataRange().getValues();
    const announcements = [];
    
    // Loop backwards to get newest first
    for (let i = data.length - 1; i > 0; i--) {
      const row = data[i];
      const status = row[2];
      
      if (status === 'Active') {
        announcements.push({
          id: row[0],
          content: row[1]
        });
      }
    }
    return announcements;
    
  } catch (e) {
    Logger.log("webGetAnnouncements Error: " + e.message);
    return []; // Return empty array on error
  }
}

/**
 * Fetches all announcements for the admin panel.
 * Only Superadmins can access this.
 */
function webGetAnnouncements_Admin() {
  try {
    const adminEmail = Session.getActiveUser().getEmail().toLowerCase();
    const ss = getSpreadsheet();
    const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
    const userData = getUserDataFromDb(dbSheet);
    
    if (userData.emailToRole[adminEmail] !== 'superadmin') {
      throw new Error("Permission denied. Only superadmins can manage announcements.");
    }

    const sheet = getOrCreateSheet(ss, SHEET_NAMES.announcements);
    const data = sheet.getDataRange().getValues();
    const results = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      results.push({
        id: row[0],
        content: row[1],
        status: row[2],
        createdBy: row[3],
        timestamp: convertDateToString(new Date(row[4]))
      });
    }
    
    return results;

  } catch (e) {
    Logger.log("webGetAnnouncements_Admin Error: " + e.message);
    return { error: e.message };
  }
}

/**
 * Saves (creates or updates) an announcement.
 * Only Superadmins can access this.
 */
function webSaveAnnouncement(announcementObject) {
  try {
    const adminEmail = Session.getActiveUser().getEmail().toLowerCase();
    const ss = getSpreadsheet();
    const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
    const userData = getUserDataFromDb(dbSheet);
    
    if (userData.emailToRole[adminEmail] !== 'superadmin') {
      throw new Error("Permission denied. Only superadmins can save announcements.");
    }

    const sheet = getOrCreateSheet(ss, SHEET_NAMES.announcements);
    const { id, content, status } = announcementObject;

    if (!content) {
      throw new Error("Content cannot be empty.");
    }

    if (id) {
      // --- Update Existing ---
      const data = sheet.getDataRange().getValues();
      let rowFound = -1;
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === id) {
          rowFound = i + 1;
          break;
        }
      }
      
      if (rowFound === -1) {
        throw new Error("Announcement ID not found. Could not update.");
      }
      
      sheet.getRange(rowFound, 2).setValue(content);
      sheet.getRange(rowFound, 3).setValue(status);
      
    } else {
      // --- Create New ---
      const newID = `ann-${new Date().getTime()}`;
      sheet.appendRow([
        newID,
        content,
        status,
        adminEmail,
        new Date()
      ]);
    }
    
    SpreadsheetApp.flush();
    return { success: true };

  } catch (e) {
    Logger.log("webSaveAnnouncement Error: " + e.message);
    return { error: e.message };
  }
}

/**
 * Deletes an announcement.
 * Only Superadmins can access this.
 */
function webDeleteAnnouncement(announcementID) {
  try {
    const adminEmail = Session.getActiveUser().getEmail().toLowerCase();
    const ss = getSpreadsheet();
    const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
    const userData = getUserDataFromDb(dbSheet);
    
    if (userData.emailToRole[adminEmail] !== 'superadmin') {
      throw new Error("Permission denied. Only superadmins can delete announcements.");
    }

    const sheet = getOrCreateSheet(ss, SHEET_NAMES.announcements);
    const data = sheet.getDataRange().getValues();
    let rowFound = -1;

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === announcementID) {
        rowFound = i + 1;
        break;
      }
    }

    if (rowFound > -1) {
      sheet.deleteRow(rowFound);
      SpreadsheetApp.flush();
      return { success: true };
    } else {
      throw new Error("Announcement not found.");
    }

  } catch (e) {
    Logger.log("webDeleteAnnouncement Error: " + e.message);
    return { error: e.message };
  }
}

/**
 * NEW: Logs a request from a user to upgrade their role.
 */
function webRequestAdminAccess(justification, requestedRole) {
  try {
    const userEmail = Session.getActiveUser().getEmail().toLowerCase();
    const ss = getSpreadsheet();
    const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
    const userData = getUserDataFromDb(dbSheet);
    
    const userName = userData.emailToName[userEmail];
    const currentRole = userData.emailToRole[userEmail] || 'agent';

    if (!userName) {
      throw new Error("Your user account could not be found.");
    }
    if (currentRole === 'superadmin') {
      throw new Error("You are already a Superadmin.");
    }
    if (currentRole === 'admin' && requestedRole === 'admin') {
      throw new Error("You are already an Admin.");
    }
    if (currentRole === 'agent' && requestedRole === 'superadmin') {
      throw new Error("You must be an Admin to request Superadmin access.");
    }

    const reqSheet = getOrCreateSheet(ss, SHEET_NAMES.roleRequests);
    const requestID = `ROLE-${new Date().getTime()}`;

    // ...
reqSheet.appendRow([
  requestID,
  userEmail,
  userName,
  currentRole,
  requestedRole,
  justification,
  new Date(),
  "Pending", // *** ADD "Pending" STATUS ***
  "",        // ActionByEmail
  ""         // ActionTimestamp
]);

    return "Your role upgrade request has been submitted for review.";

  } catch (e) {
    Logger.log("webRequestAdminAccess Error: " + e.message);
    return "Error: " + e.message;
  }
}

/**
 * Fetches pending role requests. Superadmin only.
 */
function webGetRoleRequests() {
  try {
    const adminEmail = Session.getActiveUser().getEmail().toLowerCase();
    const ss = getSpreadsheet();
    const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
    const userData = getUserDataFromDb(dbSheet);

    if (userData.emailToRole[adminEmail] !== 'superadmin') {
      throw new Error("Permission denied. Only superadmins can view role requests.");
    }

    const reqSheet = getOrCreateSheet(ss, SHEET_NAMES.roleRequests);
    const data = reqSheet.getDataRange().getValues();
    const headers = data[0];
    const results = [];
    
    // Find column indexes
    const statusIndex = headers.indexOf("Status");
    const idIndex = headers.indexOf("RequestID");
    const emailIndex = headers.indexOf("UserEmail");
    const nameIndex = headers.indexOf("UserName");
    const currentIndex = headers.indexOf("CurrentRole");
    const requestedIndex = headers.indexOf("RequestedRole");
    const justifyIndex = headers.indexOf("Justification");
    const timeIndex = headers.indexOf("RequestTimestamp");

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[statusIndex] === 'Pending') {
        results.push({
          requestID: row[idIndex],
          userEmail: row[emailIndex],
          userName: row[nameIndex],
          currentRole: row[currentIndex],
          requestedRole: row[requestedIndex],
          justification: row[justifyIndex],
          timestamp: convertDateToString(new Date(row[timeIndex]))
        });
      }
    }
    return results.sort((a,b) => new Date(b.timestamp) - new Date(a.timestamp)); // Newest first
  } catch (e) {
    Logger.log("webGetRoleRequests Error: " + e.message);
    return { error: e.message };
  }
}

/**
 * Approves or denies a role request. Superadmin only.
 */
function webApproveDenyRoleRequest(requestID, newStatus) {
  try {
    const adminEmail = Session.getActiveUser().getEmail().toLowerCase();
    const ss = getSpreadsheet();
    const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
    const userData = getUserDataFromDb(dbSheet);

    if (userData.emailToRole[adminEmail] !== 'superadmin') {
      throw new Error("Permission denied. Only superadmins can action role requests.");
    }

    const reqSheet = getOrCreateSheet(ss, SHEET_NAMES.roleRequests);
    const data = reqSheet.getDataRange().getValues();
    const headers = data[0];

    // Find columns
    const idIndex = headers.indexOf("RequestID");
    const statusIndex = headers.indexOf("Status");
    const emailIndex = headers.indexOf("UserEmail");
    const requestedIndex = headers.indexOf("RequestedRole");
    const actionByIndex = headers.indexOf("ActionByEmail");
    const actionTimeIndex = headers.indexOf("ActionTimestamp");
    
    let rowToUpdate = -1;
    let requestDetails = {};

    for (let i = 1; i < data.length; i++) {
      if (data[i][idIndex] === requestID) {
        rowToUpdate = i + 1; // 1-based index
        requestDetails = {
          status: data[i][statusIndex],
          userEmail: data[i][emailIndex],
          newRole: data[i][requestedIndex]
        };
        break;
      }
    }

    if (rowToUpdate === -1) throw new Error("Request ID not found.");
    if (requestDetails.status !== 'Pending') throw new Error(`This request has already been ${requestDetails.status}.`);

    // 1. Update the Role Request sheet
    reqSheet.getRange(rowToUpdate, statusIndex + 1).setValue(newStatus);
    reqSheet.getRange(rowToUpdate, actionByIndex + 1).setValue(adminEmail);
    reqSheet.getRange(rowToUpdate, actionTimeIndex + 1).setValue(new Date());

    // 2. If Approved, update the Data Base
    if (newStatus === 'Approved') {
      const userDBRow = userData.emailToRow[requestDetails.userEmail];
      if (!userDBRow) {
        throw new Error(`Could not find user ${requestDetails.userEmail} in Data Base to update role.`);
      }
      // Find Role column (Column C = 3)
      dbSheet.getRange(userDBRow, 3).setValue(requestDetails.newRole);
    }
    
    SpreadsheetApp.flush();
    return { success: true, message: `Request has been ${newStatus}.` };
  } catch (e) {
    Logger.log("webApproveDenyRoleRequest Error: " + e.message);
    return { error: e.message };
  }
}

// ADD this new function to the end of your code.gs file
/**
 * Calculates and adds leave balances monthly based on hiring date.
 * This function should be run on a monthly time-based trigger.
 */
function monthlyLeaveAccrual() {
  const ss = getSpreadsheet();
  const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
  const logsSheet = getOrCreateSheet(ss, SHEET_NAMES.logs);
  const userData = getUserDataFromDb(dbSheet);
  const today = new Date();
  
  Logger.log("Starting monthlyLeaveAccrual trigger...");

  for (const user of userData.userList) {
    try {
      const hiringDate = userData.emailToHiringDate[user.email];
      
      // Skip if no hiring date or account is not active
      if (!hiringDate || user.accountStatus !== 'Active') {
        continue;
      }

      // Calculate years of service
      const yearsOfService = (today.getTime() - hiringDate.getTime()) / (1000 * 60 * 60 * 24 * 365.25);
      
      let annualDaysPerYear;
      if (yearsOfService >= 10) {
        annualDaysPerYear = 30;
      } else if (yearsOfService >= 1) {
        annualDaysPerYear = 21;
      } else {
        annualDaysPerYear = 15;
      }

      const monthlyAccrual = annualDaysPerYear / 12;
      
      const userRow = userData.emailToRow[user.email];
      if (!userRow) continue; // Should not happen, but safe check
      
      // Get Annual Balance range (Column D = 4)
      const balanceRange = dbSheet.getRange(userRow, 4); 
      const currentBalance = parseFloat(balanceRange.getValue()) || 0;
      const newBalance = currentBalance + monthlyAccrual;
      
      balanceRange.setValue(newBalance);
      
      logsSheet.appendRow([
        new Date(), 
        user.name, 
        'SYSTEM', 
        'Monthly Accrual', 
        `Added ${monthlyAccrual.toFixed(2)} days (Rate: ${annualDaysPerYear}/yr). New Balance: ${newBalance.toFixed(2)}`
      ]);

    } catch (e) {
      Logger.log(`Failed to process accrual for ${user.name}: ${e.message}`);
    }
  }
  Logger.log("Finished monthlyLeaveAccrual trigger.");
}

/**
 * REPLACED: Robustly parses a date from CSV, handling strings, numbers, and Date objects.
 */
function parseDate(dateInput) {
  if (!dateInput) return null;
  if (dateInput instanceof Date) return dateInput; // Already a date

  try {
    // Check if it's a serial number (e.g., 45576)
    if (typeof dateInput === 'number' && dateInput > 1) {
      // Google Sheets/Excel serial date (days since Dec 30, 1899)
      // Use UTC to avoid timezone issues during calculation.
      const baseDate = new Date(Date.UTC(1899, 11, 30)); // 1899-12-30 UTC
      baseDate.setUTCDate(baseDate.getUTCDate() + dateInput);
      if (!isNaN(baseDate.getTime())) return baseDate;
    }
    
    // Check for MM/dd/yyyy format (common in US CSVs)
    if (typeof dateInput === 'string' && /^\d{1,2}\/\d{1,2}\/\d{4}$/.test(dateInput)) {
      const parts = dateInput.split('/');
      // new Date(year, monthIndex, day)
      const newDate = new Date(parts[2], parts[0] - 1, parts[1]);
      if (!isNaN(newDate.getTime())) return newDate;
    }

    // Try standard parsing for ISO (yyyy-MM-dd) or other recognizable formats
    const newDate = new Date(dateInput);
    if (!isNaN(newDate.getTime())) return newDate;

    return null; // Invalid date
  } catch(e) {
    return null;
  }
}

/**
 * NEW: Robustly parses a time from CSV, handling strings and serial numbers (fractions).
 * Returns a string in HH:mm:ss format.
 */
function parseCsvTime(timeInput, timeZone) {
  if (timeInput === null || timeInput === undefined || timeInput === "") return ""; // Allow empty time

  try {
    // Check if it's a serial number (e.g., 0.5 for 12:00 PM)
    if (typeof timeInput === 'number' && timeInput >= 0 && timeInput <= 1) { // 1.0 is 24:00, which is 00:00
      // Handle edge case 1.0 = 00:00:00
      if (timeInput === 1) return "00:00:00"; 
      
      const totalSeconds = Math.round(timeInput * 86400);
      const hours = Math.floor(totalSeconds / 3600);
      const minutes = Math.floor((totalSeconds % 3600) / 60);
      const seconds = totalSeconds % 60;
      
      const hh = String(hours).padStart(2, '0');
      const mm = String(minutes).padStart(2, '0');
      const ss = String(seconds).padStart(2, '0');
      
      return `${hh}:${mm}:${ss}`;
    }

    // Check if it's a string (e.g., "12:00" or "12:00:00" or "12:00 PM")
    if (typeof timeInput === 'string') {
      // Try parsing as a date (handles "12:00 PM", "12:00", "12:00:00")
      const dateFromTime = new Date('1970-01-01 ' + timeInput);
      if (!isNaN(dateFromTime.getTime())) {
          return Utilities.formatDate(dateFromTime, timeZone, "HH:mm:ss");
      }
    }
    
    // Check if it's a full Date object (e.g., from a formatted cell)
    if (timeInput instanceof Date) {
      return Utilities.formatDate(timeInput, timeZone, "HH:mm:ss");
    }
    
    return ""; // Could not parse
  } catch(e) {
    Logger.log(`parseCsvTime Error for input "${timeInput}": ${e.message}`);
    return ""; // Return empty on error
  }
}
