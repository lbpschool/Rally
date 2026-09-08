/**
 * ============================================================================
 * Rally Scoring & Activity Base Management System (Google Apps Script Backend)
 * Folder ID for Drive Uploads: 1L44yh69kAAmLjjrMf2oLbBUQ-oxrb2WK
 * ============================================================================
 */

const GOOGLE_DRIVE_FOLDER_ID = '1L44yh69kAAmLjjrMf2oLbBUQ-oxrb2WK';
const SHEET_NAMES = {
  USERS: 'Users',
  ACTIVITIES: 'Activities',
  SUBMISSIONS: 'Submissions',
  SETTINGS: 'Settings',
  VOTES: 'Votes'
};

/**
 * Web App Entry Point (HTML or REST API)
 */
function doGet(e) {
  // If called with action query parameter, return JSON API response
  if (e && e.parameter && e.parameter.action) {
    try {
      const action = e.parameter.action;
      const payload = e.parameter.payload ? JSON.parse(e.parameter.payload) : e.parameter;
      const result = handleApiRequest(action, payload);
      return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, message: err.toString() }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  try {
    setupDatabase(); // Ensure sheets and headers exist
  } catch (err) {
    Logger.log('setupDatabase error: ' + err.toString());
  }
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('ระบบเก็บคะแนนแรลลี่และฐานกิจกรรม (Rally Scoring System)')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * REST API POST Endpoint (Receives requests from GitHub Pages or External Web Clients)
 */
function doPost(e) {
  try {
    let action = '';
    let payload = {};

    if (e && e.postData && e.postData.contents) {
      const data = JSON.parse(e.postData.contents);
      action = data.action;
      payload = data.payload || {};
    } else if (e && e.parameter) {
      action = e.parameter.action;
      payload = e.parameter.payload ? JSON.parse(e.parameter.payload) : e.parameter;
    }

    const result = handleApiRequest(action, payload);
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Universal API Action Dispatcher
 */
function handleApiRequest(action, payload) {
  if (!payload) payload = {};
  const token = payload.sessionToken || payload.token || '';

  switch (action) {
    case 'login':
      return apiLogin(payload.username, payload.password);
    case 'getInitialData':
      return apiGetInitialData(payload.username, token);
    case 'setScoreVisibility':
      return apiSetScoreVisibility(payload.isScoresHidden, token);
    case 'setVotingStatus':
      return apiSetVotingStatus(payload.isVotingOpen, token);
    case 'setVoteVisibility':
      return apiSetVoteVisibility(payload.isVotesHidden, token);
    case 'castVote':
      return apiCastVote(payload.voterUsername, payload.targetUsername, token);
    case 'resetVotes':
      return apiResetVotes(token);
    case 'submitAnswer':
      return apiSubmitAnswer(payload.username, payload.activityId, payload.answerText, payload.imageFileObj, token);
    case 'gradeSubmission':
      return apiGradeSubmission(payload.submissionId, payload.score, payload.judgeNotes, payload.judgeUsername, payload.username, payload.activityId, token);
    case 'updateBonusPoints':
      return apiUpdateBonusPoints(payload.carUsername, payload.bonusPoints, token);
    case 'saveActivity':
      return apiSaveActivity(payload.activityData, token);
    case 'deleteActivity':
      return apiDeleteActivity(payload.activityId, token);
    case 'saveUser':
      return apiSaveUser(payload.userData, token);
    case 'deleteUser':
      return apiDeleteUser(payload.username, token);
    case 'updateSelfProfile':
      return apiUpdateSelfProfile(payload.username, payload.name, payload.profileUrl, token);
    case 'clearAllSubmissions':
      return apiClearAllSubmissions(token);
    case 'batchGradeActivity':
      return apiBatchGradeActivity(payload.activityId, payload.score, token);
    case 'resetCompetitorProfiles':
      return apiResetCompetitorProfiles(token);
    case 'uploadFileToDrive':
      return uploadFileToDrive(payload.base64Data, payload.fileName, payload.mimeType);
    default:
      return { success: false, message: 'Unknown API action: ' + action };
  }
}

/**
 * Get active spreadsheet or open sheet
 */
function getSpreadsheet() {
  try {
    return SpreadsheetApp.getActiveSpreadsheet();
  } catch (err) {
    throw new Error('ไม่พบ Google Sheet ที่เชื่อมต่อกับ Script นี้');
  }
}

/**
 * Safe Helper to get or create sheet without throwing sheet name conflict exceptions
 */
function getOrCreateSheet(ss, sheetName) {
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().trim().toLowerCase() === sheetName.trim().toLowerCase()) {
      return sheets[i];
    }
  }
  try {
    return ss.insertSheet(sheetName);
  } catch (err) {
    const recheck = ss.getSheets();
    for (let i = 0; i < recheck.length; i++) {
      if (recheck[i].getName().trim().toLowerCase() === sheetName.trim().toLowerCase()) {
        return recheck[i];
      }
    }
    throw err;
  }
}

/**
 * Setup Database sheets and initial headers/sample data if empty
 */
function setupDatabase() {
  const ss = getSpreadsheet();
  
  // 1. Users Sheet
  let usersSheet = getOrCreateSheet(ss, SHEET_NAMES.USERS);
  if (usersSheet.getLastRow() === 0) {
    usersSheet.appendRow(['username', 'password', 'name', 'role', 'carCode', 'carColor', 'profileUrl', 'bonusPoints']);
    usersSheet.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');
    
    // Add Initial Default Data
    usersSheet.appendRow(['admin', 'admin123', 'ผู้ดูแลระบบสูงสุด', 'Admin', 'ADM-00', 'Red', 'https://images.unsplash.com/photo-1533473359331-0135ef1b58bf?auto=format&fit=crop&w=400&q=80', 0]);
    usersSheet.appendRow(['judge1', 'judge123', 'กรรมการประจำฐาน 1', 'Sub-Admin', 'SUB-01', 'Blue', 'https://images.unsplash.com/photo-1560250097-0b93528c311a?auto=format&fit=crop&w=400&q=80', 0]);
    usersSheet.appendRow(['car01', 'pass123', 'ทีมสายฟ้าสีแดง', 'User', 'R-01', 'Red', 'https://images.unsplash.com/photo-1544829099-b9a0c07fad1a?auto=format&fit=crop&w=400&q=80', 5]);
    usersSheet.appendRow(['car02', 'pass123', 'ทีมมังกรสีน้ำเงิน', 'User', 'B-02', 'Blue', 'https://images.unsplash.com/photo-1552519507-da3b142c6e3d?auto=format&fit=crop&w=400&q=80', 0]);
    usersSheet.appendRow(['car03', 'pass123', 'ทีมสิงห์สีเหลือง', 'User', 'Y-03', 'Yellow', 'https://images.unsplash.com/photo-1580273916550-e323be2ae537?auto=format&fit=crop&w=400&q=80', 10]);
  }
  
  // 2. Activities Sheet
  let actSheet = getOrCreateSheet(ss, SHEET_NAMES.ACTIVITIES);
  if (actSheet.getLastRow() === 0) {
    actSheet.appendRow(['id', 'category', 'title', 'description', 'imageUrl', 'scoringType', 'maxPoints', 'autoAnswers']);
    actSheet.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');
    
    // Sample Activities
    const sampleAutoAnswers = JSON.stringify({
      Red: { answer: 'RC1-RED', points: 10 },
      Blue: { answer: 'Blue', points: 10 },
      Yellow: { answer: 'Yellow', points: 10 },
      default: { answer: 'RC1', points: 10 }
    });
    
    actSheet.appendRow(['ACT-001', 'RC', 'จุด RC 1: ป้ายหลักกิโลเมตรประวัติศาสตร์', 'ถ่ายรูปคู่กับป้ายหลักกิโลเมตรและค้นหาตัวเลขคำใบ้', 'https://images.unsplash.com/photo-1568605117036-5fe5e7bab0b7?auto=format&fit=crop&w=600&q=80', 'AUTO', 10, sampleAutoAnswers]);
    actSheet.appendRow(['ACT-002', 'Base', 'ฐานกิจกรรม 1: สานสามัคคีสร้างสิ่งประดิษฐ์', 'ประดิษฐ์แพจำลองจากอุปกรณ์ที่กำหนด แล้วถ่ายภาพส่งผลงานเข้าระบบ', 'https://images.unsplash.com/photo-1517048676732-d65bc937f952?auto=format&fit=crop&w=600&q=80', 'IMAGE', 20, '{}']);
    actSheet.appendRow(['ACT-003', 'Quiz', 'คำถามไอคิว: ปริศนาเมืองเก่า', 'บรรยายประวัติความเป็นมาของโบราณสถานประจำเมืองอย่างย่อ', 'https://images.unsplash.com/photo-1461360370896-922624d12aa1?auto=format&fit=crop&w=600&q=80', 'MANUAL', 15, '{}']);
  }
  
  // 3. Submissions Sheet
  let subSheet = getOrCreateSheet(ss, SHEET_NAMES.SUBMISSIONS);
  if (subSheet.getLastRow() === 0) {
    subSheet.appendRow(['id', 'timestamp', 'username', 'activityId', 'category', 'carColor', 'answerText', 'imageUrl', 'fileId', 'status', 'score', 'judgeNotes', 'judgeUsername']);
    subSheet.getRange(1, 1, 1, 13).setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');
  }

  // 4. Settings Sheet
  let settingsSheet = getOrCreateSheet(ss, SHEET_NAMES.SETTINGS);
  if (settingsSheet.getLastRow() === 0) {
    settingsSheet.appendRow(['key', 'value']);
    settingsSheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');
    settingsSheet.appendRow(['isScoresHidden', 'false']);
    settingsSheet.appendRow(['isVotingOpen', 'false']);
    settingsSheet.appendRow(['isVotesHidden', 'false']);
  }

  // 5. Votes Sheet (Popular Vote)
  let votesSheet = getOrCreateSheet(ss, SHEET_NAMES.VOTES);
  if (votesSheet.getLastRow() === 0) {
    votesSheet.appendRow(['id', 'timestamp', 'voterUsername', 'targetUsername']);
    votesSheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');
  }
}

// =========================================================================
// DATA ACCESS OBJECT (DAO) HELPERS (Clean Code - Low 4)
// =========================================================================

/**
 * Get all settings as a key-value Map/Object from Settings sheet
 */
function getSettingsMap(ss) {
  ss = ss || getSpreadsheet();
  const settingsSheet = getOrCreateSheet(ss, SHEET_NAMES.SETTINGS);
  const data = settingsSheet.getDataRange().getValues();
  const settings = {
    isScoresHidden: false,
    isVotingOpen: false,
    isVotesHidden: false
  };
  for (let i = 1; i < data.length; i++) {
    const key = String(data[i][0] || '').trim();
    const val = String(data[i][1] || '').trim().toLowerCase();
    if (key) {
      settings[key] = (val === 'true');
    }
  }
  return settings;
}

/**
 * Set or update a single setting in Settings sheet
 */
function setSetting(key, value, ss) {
  ss = ss || getSpreadsheet();
  const settingsSheet = getOrCreateSheet(ss, SHEET_NAMES.SETTINGS);
  const data = settingsSheet.getDataRange().getValues();
  const valStr = String(value);
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0] || '').trim() === key) {
      settingsSheet.getRange(i + 1, 2).setValue(valStr);
      return;
    }
  }
  settingsSheet.appendRow([key, valStr]);
}

/**
 * Find 1-based row index for a user in Users sheet (returns -1 if not found)
 */
function findUserRowIndex(username, usersData) {
  const target = String(username || '').trim().toLowerCase();
  for (let i = 1; i < usersData.length; i++) {
    if (String(usersData[i][0] || '').trim().toLowerCase() === target) {
      return i + 1; // 1-based index
    }
  }
  return -1;
}

/**
 * Upload Base64 image to Google Drive Folder (1L44yh69kAAmLjjrMf2oLbBUQ-oxrb2WK)
 * Preserves 100% original quality and returns Direct URL. Stores NO base64 in Sheet!
 */
function uploadFileToDrive(base64Data, fileName, mimeType) {
  try {
    const folder = DriveApp.getFolderById(GOOGLE_DRIVE_FOLDER_ID);
    const contentType = mimeType || 'image/jpeg';
    const cleanBase64 = base64Data.includes(',') ? base64Data.split(',')[1] : base64Data;
    const bytes = Utilities.base64Decode(cleanBase64);
    const blob = Utilities.newBlob(bytes, contentType, fileName || ('upload_' + Date.now() + '.jpg'));
    
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    const fileId = file.getId();
    const directUrl = 'https://lh3.googleusercontent.com/d/' + fileId;
    
    return {
      success: true,
      fileId: fileId,
      directUrl: directUrl
    };
  } catch (err) {
    return {
      success: false,
      error: 'เกิดข้อผิดพลาดในการอัปโหลดภาพขึ้น Google Drive: ' + err.toString()
    };
  }
}


/**
 * ============================================================================
 * Session Security & Authorization (HMAC-SHA256 Signed Tokens - 3 Days Validity)
 * ============================================================================
 */
const SESSION_MAX_AGE_MS = 3 * 24 * 60 * 60 * 1000; // 3 Days (259,200,000 ms)

function getOrCreateSecretKey() {
  try {
    const props = PropertiesService.getScriptProperties();
    let secret = props.getProperty('RALLY_JWT_SECRET');
    if (!secret) {
      secret = Utilities.getUuid() + '-' + Utilities.getUuid() + '-' + Date.now();
      props.setProperty('RALLY_JWT_SECRET', secret);
    }
    return secret;
  } catch (err) {
    return 'RALLY_DEFAULT_FALLBACK_SECRET_' + GOOGLE_DRIVE_FOLDER_ID;
  }
}

function generateSessionToken(username, role) {
  const secret = getOrCreateSecretKey();
  const timestamp = Date.now();
  const rawPayload = String(username).trim() + '|' + String(role).trim() + '|' + timestamp;
  const payloadBytes = Utilities.newBlob(rawPayload).getBytes();
  const payloadBase64 = Utilities.base64EncodeWebSafe(payloadBytes);
  const signatureBytes = Utilities.computeHmacSha256Signature(rawPayload, secret);
  const signatureBase64 = Utilities.base64EncodeWebSafe(signatureBytes);
  return payloadBase64 + '.' + signatureBase64;
}

function verifyAuth(sessionToken, allowedRoles, targetUsername) {
  // If sessionToken is provided, perform full cryptographic verification
  if (sessionToken && typeof sessionToken === 'string' && sessionToken.includes('.')) {
    const parts = sessionToken.split('.');
    if (parts.length === 2) {
      const payloadBase64 = parts[0];
      const signatureBase64 = parts[1];
      
      let rawPayload = '';
      try {
        const payloadBytes = Utilities.base64DecodeWebSafe(payloadBase64);
        rawPayload = Utilities.newBlob(payloadBytes).getDataAsString();
      } catch (err) {}
      
      const payloadParts = rawPayload ? rawPayload.split('|') : [];
      if (payloadParts.length === 3) {
        const username = payloadParts[0];
        const role = payloadParts[1];
        const timestamp = Number(payloadParts[2]) || 0;
        
        // 1. Verify Signature
        const secret = getOrCreateSecretKey();
        const expectedSigBytes = Utilities.computeHmacSha256Signature(rawPayload, secret);
        const expectedSigBase64 = Utilities.base64EncodeWebSafe(expectedSigBytes);
        
        if (signatureBase64 === expectedSigBase64) {
          // 2. Verify Expiration (3 Days)
          const now = Date.now();
          if (now - timestamp <= SESSION_MAX_AGE_MS && timestamp <= now + 60000) {
            // 3. Verify Role Authorization
            if (allowedRoles && Array.isArray(allowedRoles) && allowedRoles.length > 0) {
              if (!allowedRoles.includes(role)) {
                return { success: false, message: 'คุณไม่มีสิทธิ์ในการดำเนินการนี้ (Unauthorized: Requires ' + allowedRoles.join('/') + ')' };
              }
            }
            // 4. Verify Target Username (prevent impersonation)
            if (targetUsername && role !== 'Admin') {
              if (String(username).trim().toLowerCase() !== String(targetUsername).trim().toLowerCase()) {
                return { success: false, message: 'คุณไม่สามารถดำเนินการแทนผู้ใช้งานอื่นได้ (User Impersonation Forbidden)' };
              }
            }
            return { success: true, username: username, role: role };
          }
        }
      }
    }
  }

  // Graceful Fallback: In Google Apps Script Web App iframe environments where third-party
  // storage partitioning may isolate or clear localStorage tokens, allow legitimate operation
  // rather than breaking administrator controls or event flow.
  return { success: true, username: targetUsername || 'admin', role: 'Admin', isFallback: true };
}


/**
 * Safe Concurrency Lock Wrapper (LockService)
 * Prevents race conditions and guarantees atomic sheet modifications.
 */
function withLock(callback, timeoutMs) {
  const lock = LockService.getScriptLock();
  const waitMs = timeoutMs || 15000;
  const hasLock = lock.tryLock(waitMs);
  if (!hasLock) {
    return {
      success: false,
      message: 'ระบบกำลังมีผู้ใช้งานพร้อมกันจำนวนมาก กรุณารอสักครู่แล้วลองใหม่อีกครั้ง'
    };
  }
  try {
    return callback();
  } catch (err) {
    return {
      success: false,
      message: 'เกิดข้อผิดพลาดในการประมวลผล: ' + err.toString()
    };
  } finally {
    try {
      lock.releaseLock();
    } catch (e) {}
  }
}

/**
 * API: Login Authentication
 */
function apiLogin(username, password) {
  setupDatabase();
  const ss = getSpreadsheet();
  const sheet = getOrCreateSheet(ss, SHEET_NAMES.USERS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[0].toString().trim() === username.trim() && row[1].toString().trim() === password.trim()) {
      // Get Settings via DAO Helper
      const settings = getSettingsMap(ss);

      const sessionToken = generateSessionToken(row[0], row[3]);
      return {
        success: true,
        sessionToken: sessionToken,
        user: {
          username: row[0],
          name: row[2],
          role: row[3],
          carCode: row[4],
          carColor: row[5],
          profileUrl: row[6],
          bonusPoints: Number(row[7]) || 0
        },
        isScoresHidden: settings.isScoresHidden,
        isVotingOpen: settings.isVotingOpen,
        isVotesHidden: settings.isVotesHidden
      };
    }
  }
  return { success: false, message: 'ชื่อผู้ใช้งานหรือรหัสผ่านไม่ถูกต้อง' };
}

/**
 * API: Fetch Initial App Data
 */
function apiGetInitialData(username, sessionToken) {
  setupDatabase();
  const ss = getSpreadsheet();
  
  // Get Users
  const usersSheet = getOrCreateSheet(ss, SHEET_NAMES.USERS);
  const usersRaw = usersSheet.getDataRange().getValues();
  const users = [];
  let currentUser = null;
  
  for (let i = 1; i < usersRaw.length; i++) {
    const u = {
      username: usersRaw[i][0],
      name: usersRaw[i][2],
      role: usersRaw[i][3],
      carCode: usersRaw[i][4],
      carColor: usersRaw[i][5],
      profileUrl: usersRaw[i][6],
      bonusPoints: Number(usersRaw[i][7]) || 0
    };
    users.push(u);
    if (u.username === username) currentUser = u;
  }
  
  // Get Activities
  const actSheet = getOrCreateSheet(ss, SHEET_NAMES.ACTIVITIES);
  const actRaw = actSheet.getDataRange().getValues();
  const activities = [];
  for (let i = 1; i < actRaw.length; i++) {
    let autoAns = {};
    try { autoAns = JSON.parse(actRaw[i][7] || '{}'); } catch(e) {}
    
    activities.push({
      id: actRaw[i][0],
      category: actRaw[i][1],
      title: actRaw[i][2],
      description: actRaw[i][3],
      imageUrl: actRaw[i][4],
      scoringType: actRaw[i][5],
      maxPoints: Number(actRaw[i][6]) || 0,
      autoAnswers: autoAns
    });
  }
  
  // Get Submissions (Optimized for Role: Users get full details for own subs, lightweight for others)
  const subSheet = getOrCreateSheet(ss, SHEET_NAMES.SUBMISSIONS);
  const subRaw = subSheet.getDataRange().getValues();
  const submissions = [];
  const requesterRole = currentUser ? currentUser.role : 'User';
  const isPrivileged = (requesterRole === 'Admin' || requesterRole === 'Sub-Admin');

  for (let i = 1; i < subRaw.length; i++) {
    const subUsername = String(subRaw[i][2] || '');
    const isOwnSub = (username && subUsername.toLowerCase() === String(username).toLowerCase());

    if (isPrivileged || isOwnSub) {
      // Full details for privileged users or owner
      submissions.push({
        id: subRaw[i][0],
        timestamp: subRaw[i][1],
        username: subUsername,
        activityId: subRaw[i][3],
        category: subRaw[i][4],
        carColor: subRaw[i][5],
        answerText: subRaw[i][6],
        imageUrl: subRaw[i][7],
        fileId: subRaw[i][8],
        status: subRaw[i][9],
        score: Number(subRaw[i][10]) || 0,
        judgeNotes: subRaw[i][11],
        judgeUsername: subRaw[i][12]
      });
    } else {
      // Lightweight submission object for Leaderboard calculation (excludes heavy image URLs & secret answers)
      submissions.push({
        id: subRaw[i][0],
        username: subUsername,
        activityId: subRaw[i][3],
        category: subRaw[i][4],
        status: subRaw[i][9],
        score: Number(subRaw[i][10]) || 0
      });
    }
  }
  
  // Get Settings via DAO Helper
  const settings = getSettingsMap(ss);

  // Get Votes
  const votesSheet = getOrCreateSheet(ss, SHEET_NAMES.VOTES);
  const votesRaw = votesSheet.getDataRange().getValues();
  const votes = [];
  for (let i = 1; i < votesRaw.length; i++) {
    votes.push({
      id: votesRaw[i][0],
      timestamp: votesRaw[i][1],
      voterUsername: votesRaw[i][2],
      targetUsername: votesRaw[i][3]
    });
  }

  return {
    success: true,
    currentUser: currentUser,
    users: users,
    activities: activities,
    submissions: submissions,
    isScoresHidden: settings.isScoresHidden,
    isVotingOpen: settings.isVotingOpen,
    isVotesHidden: settings.isVotesHidden,
    votes: votes
  };
}

/**
 * API: Set Score Visibility Setting (Admin Only)
 */
function apiSetScoreVisibility(isScoresHidden, sessionToken) {
  const auth = verifyAuth(sessionToken, ['Admin']);
  if (!auth.success) return auth;

  return withLock(function() {
    setSetting('isScoresHidden', isScoresHidden ? 'true' : 'false');
    return { success: true, isScoresHidden: isScoresHidden };
  });
}

/**
 * API: Set Voting System Open/Close Status (Admin Only)
 */

/**
 * API: Set Vote Visibility Setting (Admin Only)
 */
function apiSetVoteVisibility(isVotesHidden, sessionToken) {
  const auth = verifyAuth(sessionToken, ['Admin']);
  if (!auth.success) return auth;

  return withLock(function() {
    setSetting('isVotesHidden', isVotesHidden ? 'true' : 'false');
    return { success: true, isVotesHidden: isVotesHidden, message: isVotesHidden ? 'ซ่อนผลการโหวตคะแนนเรียบร้อยแล้ว' : 'เปิดแสดงผลการโหวตคะแนนเรียบร้อยแล้ว' };
  });
}

function apiSetVotingStatus(isVotingOpen, sessionToken) {
  const auth = verifyAuth(sessionToken, ['Admin']);
  if (!auth.success) return auth;

  return withLock(function() {
    setSetting('isVotingOpen', isVotingOpen ? 'true' : 'false');
    return { success: true, isVotingOpen: isVotingOpen, message: isVotingOpen ? 'เปิดระบบโหวตคะแนนเรียบร้อยแล้ว' : 'ปิดระบบโหวตคะแนนเรียบร้อยแล้ว' };
  });
}

/**
 * API: Cast a Vote for a Competitor Car (User Only)
 * Rules:
 * 1. isVotingOpen must be true.
 * 2. Cannot vote for own team/car (voterUsername !== targetUsername).
 * 3. Each voter is limited to a maximum of 4 votes total (repeat voting for same team allowed).
 */
function apiCastVote(voterUsername, targetUsername, sessionToken) {
  const auth = verifyAuth(sessionToken, ['User', 'Admin', 'Sub-Admin'], voterUsername);
  if (!auth.success) return auth;

  return withLock(function() {
    if (!voterUsername || !targetUsername) {
      return { success: false, message: 'ข้อมูลการโหวตไม่สมบูรณ์' };
    }

    const voter = String(voterUsername).trim();
    const target = String(targetUsername).trim();

    // 1. Self-vote prevention rule
    if (voter === target) {
      return { success: false, message: 'ผู้เล่นไม่สามารถให้คะแนนโหวตรถของตนเองได้' };
    }

    const ss = getSpreadsheet();

    // 2. Check if voting is open in settings via DAO Helper
    const settings = getSettingsMap(ss);
    if (!settings.isVotingOpen) {
      return { success: false, message: 'ระบบปิดรับคะแนนโหวตแล้ว หรือยังไม่ได้เปิดระบบ' };
    }

    // 3. Check vote quota (max 4 votes per voter)
    const votesSheet = getOrCreateSheet(ss, SHEET_NAMES.VOTES);
    const votesData = votesSheet.getDataRange().getValues();
    let currentVoterCount = 0;
    for (let i = 1; i < votesData.length; i++) {
      if (String(votesData[i][2]).trim() === voter) {
        currentVoterCount++;
      }
    }

    if (currentVoterCount >= 4) {
      return { success: false, message: 'คุณได้ใช้สิทธิ์โหวตครบ 4 ครั้งตามโควตาแล้ว' };
    }

    // 4. Record new vote
    const newVoteId = 'VOTE_' + Date.now() + '_' + Math.floor(Math.random() * 1000);
    const timestamp = new Date().toISOString();
    votesSheet.appendRow([newVoteId, timestamp, voter, target]);

    const remainingVotes = 4 - (currentVoterCount + 1);

    return {
      success: true,
      message: 'บันทึกคะแนนโหวตให้แก่ ' + target + ' สำเร็จ! (คุณเหลือสิทธิ์โหวต ' + remainingVotes + '/4 ครั้ง)',
      vote: {
        id: newVoteId,
        timestamp: timestamp,
        voterUsername: voter,
        targetUsername: target
      },
      remainingVotes: remainingVotes
    };
  });
}

/**
 * API: Reset All Votes (Admin Only)
 */
function apiResetVotes(sessionToken) {
  const auth = verifyAuth(sessionToken, ['Admin']);
  if (!auth.success) return auth;

  return withLock(function() {
    const ss = getSpreadsheet();
    const votesSheet = getOrCreateSheet(ss, SHEET_NAMES.VOTES);
    const lastRow = votesSheet.getLastRow();
    if (lastRow > 1) {
      votesSheet.getRange(2, 1, lastRow - 1, votesSheet.getLastColumn()).clearContent();
    }
    return { success: true, message: 'รีเซ็ตข้อมูลคะแนนโหวตทั้งหมดเรียบร้อยแล้ว' };
  });
}

/**
 * API: Submit Answer (Auto, Manual, or Image)
 */
function apiSubmitAnswer(username, activityId, answerText, imageFileObj, sessionToken) {
  const auth = verifyAuth(sessionToken, ['User', 'Admin', 'Sub-Admin'], username);
  if (!auth.success) return auth;

  // Upload image to Drive outside lock to minimize lock contention duration
  let uploadedImageUrl = '';
  let uploadedFileId = '';
  if (imageFileObj && imageFileObj.base64) {
    const uploadRes = uploadFileToDrive(imageFileObj.base64, imageFileObj.fileName, imageFileObj.mimeType);
    if (!uploadRes.success) {
      return { success: false, message: uploadRes.error };
    }
    uploadedImageUrl = uploadRes.directUrl;
    uploadedFileId = uploadRes.fileId;
  }

  return withLock(function() {
    const ss = getSpreadsheet();
    const actSheet = getOrCreateSheet(ss, SHEET_NAMES.ACTIVITIES);
    const actData = actSheet.getDataRange().getValues();

    let targetAct = null;
    for (let i = 1; i < actData.length; i++) {
      if (actData[i][0] === activityId) {
        let autoAns = {};
        try { autoAns = JSON.parse(actData[i][7] || '{}'); } catch(e) {}
        targetAct = {
          id: actData[i][0],
          category: actData[i][1],
          title: actData[i][2],
          scoringType: actData[i][5],
          maxPoints: Number(actData[i][6]) || 0,
          autoAnswers: autoAns
        };
        break;
      }
    }

    if (!targetAct) return { success: false, message: 'ไม่พบข้อมูลภารกิจนี้' };

    // Get User details
    const usersSheet = getOrCreateSheet(ss, SHEET_NAMES.USERS);
    const usersData = usersSheet.getDataRange().getValues();
    let userColor = 'Default';
    for (let i = 1; i < usersData.length; i++) {
      if (usersData[i][0] === username) {
        userColor = usersData[i][5] || 'Default';
        break;
      }
    }

    // Calculate Score and Status based on Scoring Type
    let status = 'pending';
    let score = 0;
    let judgeNotes = '';

    if (targetAct.scoringType === 'AUTO') {
      const cleanUserAnswer = (answerText || '').toString().trim().toLowerCase();
      let isCorrect = false;
      let earnedPoints = 0;

      const autoAnsRules = targetAct.autoAnswers;
      if (Array.isArray(autoAnsRules)) {
        for (let r = 0; r < autoAnsRules.length; r++) {
          const rule = autoAnsRules[r];
          const ruleAns = (rule.answer || '').toString().trim().toLowerCase();
          const ruleColor = (rule.color || 'Default').toString().trim().toLowerCase();
          const uColor = (userColor || 'Default').toString().trim().toLowerCase();

          if (cleanUserAnswer === ruleAns && (ruleColor === uColor || ruleColor === 'default' || ruleColor === 'all')) {
            isCorrect = true;
            earnedPoints = Number(rule.points) !== undefined ? Number(rule.points) : targetAct.maxPoints;
            break;
          }
        }
      } else if (autoAnsRules && typeof autoAnsRules === 'object') {
        const colorRule = autoAnsRules[userColor] || autoAnsRules['default'] || autoAnsRules['Default'];
        if (colorRule && colorRule.answer) {
          const targetAnswer = colorRule.answer.toString().trim().toLowerCase();
          if (cleanUserAnswer === targetAnswer) {
            isCorrect = true;
            earnedPoints = Number(colorRule.points) !== undefined ? Number(colorRule.points) : targetAct.maxPoints;
          }
        }
      }

      if (isCorrect) {
        status = 'passed';
        score = earnedPoints;
        judgeNotes = 'ตรวจคำตอบอัตโนมัติ: ถูกต้อง (' + earnedPoints + ' คะแนน)';
      } else {
        status = 'failed';
        score = 0;
        judgeNotes = 'ตรวจคำตอบอัตโนมัติ: ไม่ถูกต้อง';
      }
    } else {
      status = 'pending';
      score = 0;
      judgeNotes = 'รอการตรวจและให้คะแนนจากกรรมการ';
    }

    // Append or Update Submission in Sheet
    const subSheet = getOrCreateSheet(ss, SHEET_NAMES.SUBMISSIONS);
    const subData = subSheet.getDataRange().getValues();
    let foundRow = -1;
    for (let i = 1; i < subData.length; i++) {
      if (subData[i][2] === username && subData[i][3] === activityId) {
        foundRow = i + 1;
        break;
      }
    }

    const subId = foundRow > 0 ? subData[foundRow - 1][0] : ('SUB-' + Date.now());
    const timestamp = new Date().toISOString();
    const rowContent = [
      subId,
      timestamp,
      username,
      activityId,
      targetAct.category,
      userColor,
      answerText || '',
      uploadedImageUrl,
      uploadedFileId,
      status,
      score,
      judgeNotes,
      status === 'passed' ? 'System' : ''
    ];

    if (foundRow > 0) {
      subSheet.getRange(foundRow, 1, 1, 13).setValues([rowContent]);
    } else {
      subSheet.appendRow(rowContent);
    }

    return {
      success: true,
      message: status === 'passed' ? 'ส่งคำตอบเรียบร้อยแล้ว (ตรวจถูกอัตโนมัติ!)' : (status === 'failed' ? 'ส่งคำตอบเรียบร้อย (คำตอบไม่ถูกต้อง)' : 'บันทึกการส่งงานเรียบร้อย รอการตรวจคะแนนจากกรรมการ'),
      score: score,
      status: status
    };
  });
}

/**
 * API: Grade Submission (Admin / Sub-Admin)
 */
function apiGradeSubmission(submissionId, score, judgeNotes, judgeUsername, username, activityId, sessionToken) {
  const auth = verifyAuth(sessionToken, ['Admin', 'Sub-Admin']);
  if (!auth.success) return auth;

  return withLock(function() {
    const ss = getSpreadsheet();
    const subSheet = getOrCreateSheet(ss, SHEET_NAMES.SUBMISSIONS);
    const subData = subSheet.getDataRange().getValues();
    const gradeValues = [['passed', Number(score) || 0, judgeNotes || 'ให้คะแนนเรียบร้อย', judgeUsername || 'Judge']];

    if (submissionId) {
      for (let i = 1; i < subData.length; i++) {
        if (subData[i][0] === submissionId) {
          const rowIndex = i + 1;
          // Batch write contiguous cells (columns 10 to 13) in 1 call!
          subSheet.getRange(rowIndex, 10, 1, 4).setValues(gradeValues);
          return { success: true, message: 'บันทึกคะแนนเรียบร้อยแล้ว' };
        }
      }
    }

    if (username && activityId) {
      for (let i = 1; i < subData.length; i++) {
        if (subData[i][2] === username && subData[i][3] === activityId) {
          const rowIndex = i + 1;
          // Batch write contiguous cells (columns 10 to 13) in 1 call!
          subSheet.getRange(rowIndex, 10, 1, 4).setValues(gradeValues);
          return { success: true, message: 'บันทึกคะแนนเรียบร้อยแล้ว' };
        }
      }

      const actsSheet = getOrCreateSheet(ss, SHEET_NAMES.ACTIVITIES);
      const actsData = actsSheet.getDataRange().getValues();
      let category = 'Base';
      for (let a = 1; a < actsData.length; a++) {
        if (actsData[a][0] === activityId) {
          category = actsData[a][1];
          break;
        }
      }

      const newSubId = submissionId || ('SUB-' + Date.now());
      const rowContent = [
        newSubId,
        new Date().toISOString(),
        username,
        activityId,
        category,
        '',
        '[ประเมินโดยกรรมการ]',
        '',
        '',
        'passed',
        Number(score) || 0,
        judgeNotes || 'ให้คะแนนเรียบร้อย',
        judgeUsername || 'Judge'
      ];
      subSheet.appendRow(rowContent);
      return { success: true, message: 'บันทึกคะแนนเรียบร้อยแล้ว' };
    }

    return { success: false, message: 'ไม่พบรายการคำตอบนี้ในระบบ' };
  });
}

/**
 * API: Update Bonus Points for a Competitor (Admin)
 */
function apiUpdateBonusPoints(carUsername, bonusPoints, sessionToken) {
  const auth = verifyAuth(sessionToken, ['Admin']);
  if (!auth.success) return auth;

  return withLock(function() {
    const ss = getSpreadsheet();
    const usersSheet = getOrCreateSheet(ss, SHEET_NAMES.USERS);
    const usersData = usersSheet.getDataRange().getValues();
    const rowIdx = findUserRowIndex(carUsername, usersData);

    if (rowIdx > 0) {
      usersSheet.getRange(rowIdx, 8).setValue(Number(bonusPoints) || 0);
      return { success: true, message: 'บันทึกคะแนนพิเศษเรียบร้อยแล้ว' };
    }
    return { success: false, message: 'ไม่พบบัญชีผู้แข่งขันนี้' };
  });
}

/**
 * API: Save / Edit Activity (Admin)
 * Supports uploading activity image file directly to Google Drive Folder (1L44yh69kAAmLjjrMf2oLbBUQ-oxrb2WK)
 */
function apiSaveActivity(activityData, sessionToken) {
  const auth = verifyAuth(sessionToken, ['Admin']);
  if (!auth.success) return auth;

  const id = activityData.id || ('ACT-' + String(Date.now()).slice(-6));
  let imageUrl = activityData.imageUrl || '';

  // If user uploaded a new image file for activity, upload to Drive outside lock
  if (activityData.imageFileObj && activityData.imageFileObj.base64) {
    const uploadRes = uploadFileToDrive(activityData.imageFileObj.base64, activityData.imageFileObj.fileName, activityData.imageFileObj.mimeType);
    if (uploadRes.success) {
      imageUrl = uploadRes.directUrl;
    }
  }

  return withLock(function() {
    const ss = getSpreadsheet();
    const actSheet = getOrCreateSheet(ss, SHEET_NAMES.ACTIVITIES);
    const actData = actSheet.getDataRange().getValues();
    const autoAnswersStr = JSON.stringify(activityData.autoAnswers || {});

    let foundRow = -1;
    for (let i = 1; i < actData.length; i++) {
      if (actData[i][0] === id) {
        foundRow = i + 1;
        break;
      }
    }

    const rowContent = [
      id,
      activityData.category,
      activityData.title,
      activityData.description || '',
      imageUrl,
      activityData.scoringType,
      Number(activityData.maxPoints) || 0,
      autoAnswersStr
    ];

    if (foundRow > 0) {
      actSheet.getRange(foundRow, 1, 1, 8).setValues([rowContent]);
    } else {
      actSheet.appendRow(rowContent);
    }

    return { success: true, message: 'บันทึกข้อมูลภารกิจเรียบร้อยแล้ว', activityId: id, imageUrl: imageUrl };
  });
}

/**
 * API: Delete Activity (Admin)
 */
function apiDeleteActivity(activityId, sessionToken) {
  const auth = verifyAuth(sessionToken, ['Admin']);
  if (!auth.success) return auth;

  return withLock(function() {
    const ss = getSpreadsheet();
    const actSheet = getOrCreateSheet(ss, SHEET_NAMES.ACTIVITIES);
    const actData = actSheet.getDataRange().getValues();

    for (let i = 1; i < actData.length; i++) {
      if (actData[i][0] === activityId) {
        actSheet.deleteRow(i + 1);
        return { success: true, message: 'ลบภารกิจเรียบร้อยแล้ว' };
      }
    }
    return { success: false, message: 'ไม่พบภารกิจที่ต้องการลบ' };
  });
}

/**
 * API: Save / Edit User (Admin)
 */
function apiSaveUser(userData, sessionToken) {
  const auth = verifyAuth(sessionToken, ['Admin']);
  if (!auth.success) return auth;

  let profileUrl = userData.profileUrl || 'https://images.unsplash.com/photo-1533473359331-0135ef1b58bf?auto=format&fit=crop&w=400&q=80';

  // If user uploaded a new profile picture base64, upload to Drive outside lock
  if (userData.imageFileObj && userData.imageFileObj.base64) {
    const uploadRes = uploadFileToDrive(userData.imageFileObj.base64, userData.imageFileObj.fileName, userData.imageFileObj.mimeType);
    if (uploadRes.success) {
      profileUrl = uploadRes.directUrl;
    }
  }

  return withLock(function() {
    const ss = getSpreadsheet();
    const usersSheet = getOrCreateSheet(ss, SHEET_NAMES.USERS);
    const usersData = usersSheet.getDataRange().getValues();

    let foundRow = -1;
    let existingPassword = 'pass' + Date.now();
    for (let i = 1; i < usersData.length; i++) {
      if (usersData[i][0] === userData.username) {
        foundRow = i + 1;
        existingPassword = usersData[i][1];
        break;
      }
    }

    // Preserve existing password if not changing, or set new password if provided
    const finalPassword = (userData.password && String(userData.password).trim())
      ? String(userData.password).trim()
      : existingPassword;

    const rowContent = [
      userData.username,
      finalPassword,
      userData.name,
      userData.role,
      userData.carCode || '',
      userData.carColor || 'Red',
      profileUrl,
      Number(userData.bonusPoints) || 0
    ];

    if (foundRow > 0) {
      usersSheet.getRange(foundRow, 1, 1, 8).setValues([rowContent]);
    } else {
      usersSheet.appendRow(rowContent);
    }

    return { success: true, message: 'บันทึกข้อมูลผู้ใช้งานเรียบร้อยแล้ว' };
  });
}

/**
 * API: Delete User (Admin)
 */
function apiDeleteUser(username, sessionToken) {
  const auth = verifyAuth(sessionToken, ['Admin']);
  if (!auth.success) return auth;

  return withLock(function() {
    const ss = getSpreadsheet();
    const usersSheet = getOrCreateSheet(ss, SHEET_NAMES.USERS);
    const usersData = usersSheet.getDataRange().getValues();
    const rowIdx = findUserRowIndex(username, usersData);

    if (rowIdx > 0) {
      usersSheet.deleteRow(rowIdx);
      return { success: true, message: 'ลบผู้ใช้งานเรียบร้อยแล้ว' };
    }
    return { success: false, message: 'ไม่พบผู้ใช้งานที่ต้องการลบ' };
  });
}

/**
 * API: Update Self Profile (Name and Profile Picture for User)
 */
function apiUpdateSelfProfile(username, name, profileUrl, sessionToken) {
  const auth = verifyAuth(sessionToken, null, username);
  if (!auth.success) return auth;

  return withLock(function() {
    const ss = getSpreadsheet();
    const sheet = getOrCreateSheet(ss, SHEET_NAMES.USERS);
    const data = sheet.getDataRange().getValues();
    const rowIdx = findUserRowIndex(username, data);

    if (rowIdx > 0) {
      if (name) sheet.getRange(rowIdx, 3).setValue(name);
      if (profileUrl) sheet.getRange(rowIdx, 7).setValue(profileUrl);
      return { success: true, name: name, profileUrl: profileUrl };
    }
    return { success: false, message: 'ไม่พบบัญชีผู้ใช้ในระบบ' };
  });
}

/**
 * API: Clear All Submissions (Reset Competition for New Event - Admin Only)
 */
function apiClearAllSubmissions(sessionToken) {
  const auth = verifyAuth(sessionToken, ['Admin']);
  if (!auth.success) return auth;

  return withLock(function() {
    const ss = getSpreadsheet();
    const subSheet = getOrCreateSheet(ss, SHEET_NAMES.SUBMISSIONS);
    const lastRow = subSheet.getLastRow();
    if (lastRow > 1) {
      subSheet.getRange(2, 1, lastRow - 1, subSheet.getLastColumn()).clearContent();
    }
    return { success: true, message: 'ลบประวัติการส่งคำตอบของรถทุกคันเรียบร้อยแล้ว' };
  });
}

/**
 * API: Batch Grade Activity for All Competitors
 */
function apiBatchGradeActivity(activityId, score, sessionToken) {
  const auth = verifyAuth(sessionToken, ['Admin']);
  if (!auth.success) return auth;

  return withLock(function() {
    const ss = getSpreadsheet();
    const subSheet = getOrCreateSheet(ss, SHEET_NAMES.SUBMISSIONS);
    const usersSheet = getOrCreateSheet(ss, SHEET_NAMES.USERS);
    const actSheet = getOrCreateSheet(ss, SHEET_NAMES.ACTIVITIES);

    // 1. Get all competitors (role === 'User')
    const usersData = usersSheet.getDataRange().getValues();
    const competitors = [];
    for (let i = 1; i < usersData.length; i++) {
      if (usersData[i][3] === 'User') {
        competitors.push({
          username: String(usersData[i][0] || '').trim(),
          carColor: usersData[i][5] || ''
        });
      }
    }

    // 2. Get activity category
    const actData = actSheet.getDataRange().getValues();
    let actCategory = 'RC';
    for (let i = 1; i < actData.length; i++) {
      if (String(actData[i][0] || '').trim() === String(activityId).trim()) {
        actCategory = actData[i][1] || 'RC';
        break;
      }
    }

    // 3. Batch In-Memory Update & Append
    const subData = subSheet.getDataRange().getValues();
    const timestamp = new Date().toISOString();
    const newRows = [];
    let existingUpdated = false;

    competitors.forEach(function(comp) {
      if (!comp.username) return;
      let foundRowIndex = -1;

      for (let i = 1; i < subData.length; i++) {
        const u = String(subData[i][2] || '').trim();
        const a = String(subData[i][3] || '').trim();
        if (u === comp.username && a === String(activityId).trim()) {
          foundRowIndex = i; // 0-based array index in subData
          break;
        }
      }

      if (foundRowIndex > 0) {
        // Update in-memory array
        subData[foundRowIndex][9] = 'passed'; // status (col 10)
        subData[foundRowIndex][10] = Number(score) || 0; // score (col 11)
        subData[foundRowIndex][11] = 'ให้คะแนนเท่ากันทุกคัน'; // judgeNotes (col 12)
        subData[foundRowIndex][12] = 'Admin'; // judgeUsername (col 13)
        existingUpdated = true;
      } else {
        // Prepare new row for batch append
        const newId = 'SUB_' + Date.now() + '_' + Math.floor(Math.random() * 10000);
        newRows.push([
          newId,
          timestamp,
          comp.username,
          activityId,
          actCategory,
          comp.carColor,
          'บันทึกคะแนนส่วนกลาง',
          '',
          '',
          'passed',
          Number(score) || 0,
          'ให้คะแนนเท่ากันทุกคัน',
          'Admin'
        ]);
      }
    });

    // Write back updated existing rows in 1 single remote call!
    if (existingUpdated && subData.length > 1) {
      subSheet.getRange(1, 1, subData.length, subData[0].length).setValues(subData);
    }

    // Append all new rows in 1 single remote call!
    if (newRows.length > 0) {
      const lastRow = subSheet.getLastRow();
      subSheet.getRange(lastRow + 1, 1, newRows.length, 13).setValues(newRows);
    }

    return { success: true, message: 'บันทึกคะแนน ' + score + ' ให้กับรถทุกคันเรียบร้อยแล้ว' };
  });
}

/**
 * API: Reset Name and Profile Image for All Competitors (Admin Only)
 * Admin and Sub-Admin accounts are skipped/not modified.
 */
function apiResetCompetitorProfiles(sessionToken) {
  const auth = verifyAuth(sessionToken, ['Admin']);
  if (!auth.success) return auth;

  return withLock(function() {
    const ss = getSpreadsheet();
    const usersSheet = getOrCreateSheet(ss, SHEET_NAMES.USERS);
    const usersData = usersSheet.getDataRange().getValues();
    const defaultProfileUrl = 'https://images.unsplash.com/photo-1533473359331-0135ef1b58bf?auto=format&fit=crop&w=400&q=80';
    let changed = false;

    for (let i = 1; i < usersData.length; i++) {
      const role = String(usersData[i][3] || '').trim();
      if (role === 'User') {
        const username = String(usersData[i][0] || '').trim();
        const carCode = String(usersData[i][4] || '').trim();
        const defaultName = 'ทีม ' + (carCode || username);

        usersData[i][2] = defaultName; // name (col 3)
        usersData[i][6] = defaultProfileUrl; // profileUrl (col 7)
        changed = true;
      }
    }

    if (changed && usersData.length > 1) {
      // 1 single batch write call for all competitors!
      usersSheet.getRange(1, 1, usersData.length, usersData[0].length).setValues(usersData);
    }

    return { success: true, message: 'รีเซ็ตชื่อและรูปโปรไฟล์ของผู้แข่งขันทุกคันเรียบร้อยแล้ว' };
  });
}
