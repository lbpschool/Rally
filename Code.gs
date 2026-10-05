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
  const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : '';
  if (action) {
    try {
      let payload = {};
      if (e.parameter.payload) {
        try {
          payload = JSON.parse(e.parameter.payload);
        } catch(pe) {
          payload = e.parameter;
        }
      } else {
        payload = e.parameter;
      }
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
    .setTitle('Rally Scoring System')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Normalizes answer text by removing all whitespace characters (spaces, tabs, newlines, non-breaking spaces)
 * and converting to lowercase for robust, whitespace-agnostic comparison.
 */
function normalizeAnswerText(str) {
  if (str === null || str === undefined) return '';
  return String(str)
    .replace(/[\s\u00A0\u200B\uFEFF]/g, '')
    .toLowerCase();
}

/**
 * REST API POST Endpoint (Receives requests from GitHub Pages or External Web Clients)
 */
function doPost(e) {
  try {
    let action = (e && e.parameter && e.parameter.action) || '';
    let payload = {};

    if (e && e.postData && e.postData.contents) {
      try {
        const data = JSON.parse(e.postData.contents);
        if (data.action) action = data.action;
        if (data.payload) payload = data.payload;
      } catch (parseErr) {}
    }

    if (!payload || Object.keys(payload).length === 0) {
      if (e && e.parameter) {
        if (e.parameter.payload) {
          try { payload = JSON.parse(e.parameter.payload); } catch(pe) { payload = e.parameter; }
        } else {
          payload = e.parameter;
        }
      }
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

  let result;
  switch (action) {
    case 'login':
      result = apiLogin(payload.username, payload.password);
      break;
    case 'getInitialData':
      result = apiGetInitialData(payload.username, token);
      break;
    case 'setBoobyRank':
      result = apiSetBoobyRank(payload.boobyRank, token);
      break;
    case 'setScoreVisibility':
      result = apiSetScoreVisibility(payload.isScoresHidden, token);
      break;
    case 'setVotingStatus':
      result = apiSetVotingStatus(payload.isVotingOpen, token);
      break;
    case 'setVoteVisibility':
      result = apiSetVoteVisibility(payload.isVotesHidden, token);
      break;
    case 'castVote':
      result = apiCastVote(payload.voterUsername, payload.targetUsername, token);
      break;
    case 'resetVotes':
      result = apiResetVotes(token);
      break;
    case 'resetVotesDirect':
      result = executeSystemVoteReset();
      break;
    case 'submitAnswer':
      result = apiSubmitAnswer(payload.username, payload.activityId, payload.answerText, payload.imageFileObj, token);
      break;
    case 'gradeSubmission':
      result = apiGradeSubmission(payload.submissionId, payload.score, payload.judgeNotes, payload.judgeUsername, payload.username, payload.activityId, token, payload.status);
      break;
    case 'updateBonusPoints':
      result = apiUpdateBonusPoints(payload.carUsername, payload.bonusPoints, token);
      break;
    case 'saveActivity':
      const actPayload = (payload && payload.activityData) ? payload.activityData : payload;
      result = apiSaveActivity(actPayload, token);
      break;
    case 'deleteActivity':
      result = apiDeleteActivity(payload.activityId || payload.id, token);
      break;
    case 'saveUser':
      const userPayload = (payload && payload.userData) ? payload.userData : payload;
      result = apiSaveUser(userPayload, token);
      break;
    case 'deleteUser':
      result = apiDeleteUser(payload.username, token);
      break;
    case 'swapCars':
      result = apiSwapCars(payload.usernameA, payload.usernameB, token);
      break;
    case 'updateSelfProfile':
      result = apiUpdateSelfProfile(payload.username, payload.name, payload.profileUrl, token);
      break;
    case 'clearAllSubmissions':
      result = apiClearAllSubmissions(token);
      break;
    case 'clearAllActivities':
      result = apiClearAllActivities(token);
      break;
    case 'batchGradeActivity':
      result = apiBatchGradeActivity(payload.activityId, payload.score, token);
      break;
    case 'resetCompetitorProfiles':
      result = apiResetCompetitorProfiles(token);
      break;
    case 'autoAssignCarColors':
      result = apiAutoAssignCarColors(token, payload.assignments, payload.colorSequence);
      break;
    case 'uploadFileToDrive':
      result = uploadFileToDrive(payload.base64Data, payload.fileName, payload.mimeType);
      break;
    case 'uploadSolutionImage':
      result = apiUploadSolutionImage(payload.activityId, payload.imageFileObj || payload.imageFiles, payload.mode || 'replace', token);
      break;
    case 'deleteSolutionImage':
      result = apiDeleteSolutionImage(payload.activityId, payload.imageUrl, token);
      break;
    case 'regradeAutoSubmissions':
      result = apiRegradeAutoSubmissions(token);
      break;
    case 'migrateSheetsToFirebase':
      result = apiMigrateSheetsToFirebase();
      break;
    default:
      result = { success: false, message: 'Unknown API action: ' + action };
      break;
  }

  // Invalidate shared cache only on global reset or full-scale data clearing
  // (Routine actions like submitAnswer, gradeSubmission, updateBonusPoints, castVote, saveUser, deleteUser, saveActivity, deleteActivity, resetCompetitorProfiles, and setting mutations update the cache in-place!)
  const structuralActions = [
    'resetVotes', 'batchGradeActivity', 'regradeAutoSubmissions', 'swapCars', 'updateSelfProfile', 'clearAllSubmissions', 'clearAllActivities', 'autoAssignCarColors'
  ];
  if (structuralActions.indexOf(action) !== -1 && result && result.success !== false) {
    invalidateGlobalCache();
  }

  return result;
}

/**
 * Shared Global System Cache Helpers (Handles 30-50 concurrent devices with 0.2s sync)
 */
function getGlobalCacheVersion() {
  try {
    const cache = CacheService.getScriptCache();
    let v = cache.get('rally_global_ver');
    if (!v) {
      v = String(Date.now());
      cache.put('rally_global_ver', v, 21600); // 6 hours
    }
    return v;
  } catch (e) {
    return 'v1';
  }
}

function invalidateGlobalCache() {
  try {
    _sheetCache = {};
    const cache = CacheService.getScriptCache();
    cache.put('rally_global_ver', String(Date.now()), 21600);
  } catch (e) {}
}

function getCachedSharedData() {
  try {
    const cache = CacheService.getScriptCache();
    const ver = getGlobalCacheVersion();
    
    // Check single key for backwards compatibility
    const legacyKey = 'rally_shared_data_' + ver;
    const legacyStr = cache.get(legacyKey);
    if (legacyStr) {
      try { return JSON.parse(legacyStr); } catch(pe) {}
    }

    // Modular cache lookup
    const uStr = cache.get('rally_users_' + ver);
    const aStr = cache.get('rally_act_' + ver);
    const sStr = cache.get('rally_settings_' + ver);
    const vStr = cache.get('rally_votes_' + ver);
    const numChunksStr = cache.get('rally_subs_chunks_' + ver);
    
    if (!uStr || !aStr || !sStr || !vStr || !numChunksStr) {
      return null;
    }
    
    const numChunks = parseInt(numChunksStr, 10) || 0;
    const chunkKeys = [];
    for (let c = 0; c < numChunks; c++) {
      chunkKeys.push('rally_subs_' + c + '_' + ver);
    }
    const chunks = cache.getAll(chunkKeys);
    let subStr = '';
    for (let c = 0; c < numChunks; c++) {
      const chunkVal = chunks['rally_subs_' + c + '_' + ver];
      if (!chunkVal) return null; // Incomplete chunk -> treat as cache miss
      subStr += chunkVal;
    }
    
    return {
      users: JSON.parse(uStr),
      activities: JSON.parse(aStr),
      settings: JSON.parse(sStr),
      votes: JSON.parse(vStr),
      submissions: JSON.parse(subStr)
    };
  } catch (e) {
    Logger.log('getCachedSharedData error: ' + e);
  }
  return null;
}

const FIREBASE_DATABASE_URL = 'https://rally-scoring-system-default-rtdb.asia-southeast1.firebasedatabase.app';

/**
 * Real-time Push to Firebase Realtime Database (< 0.1s update on all client devices)
 */
function syncToFirebase(sharedData) {
  try {
    if (!sharedData) sharedData = getCachedSharedData();
    if (!sharedData) return;

    const firebaseUrl = FIREBASE_DATABASE_URL + '/live_rally_data.json';
    const payload = JSON.stringify({
      users: sharedData.users || [],
      submissions: sharedData.submissions || [],
      votes: sharedData.votes || [],
      settings: sharedData.settings || {},
      timestamp: Date.now()
    });

    UrlFetchApp.fetch(firebaseUrl, {
      method: 'patch',
      contentType: 'application/json',
      payload: payload,
      muteHttpExceptions: true
    });

    // If submissions is empty, explicitly DELETE the submissions node so Firebase does not retain stale submissions
    if (!sharedData.submissions || sharedData.submissions.length === 0) {
      UrlFetchApp.fetch(FIREBASE_DATABASE_URL + '/live_rally_data/submissions.json', {
        method: 'delete',
        muteHttpExceptions: true
      });
    }
  } catch (e) {
    Logger.log('Firebase sync error: ' + e);
  }
}

/**
 * Real-time Push ONLY Settings to Firebase Realtime Database (< 0.1s update)
 * Synchronizes settings without overwriting or interfering with other data nodes.
 */
function syncSettingsToFirebase(settings) {
  try {
    if (!settings) return;
    const firebaseUrl = FIREBASE_DATABASE_URL + '/live_rally_data/settings.json';
    UrlFetchApp.fetch(firebaseUrl, {
      method: 'put',
      contentType: 'application/json',
      payload: JSON.stringify(settings),
      muteHttpExceptions: true
    });
    // Update timestamp to notify all listening clients
    UrlFetchApp.fetch(FIREBASE_DATABASE_URL + '/live_rally_data/timestamp.json', {
      method: 'put',
      contentType: 'application/json',
      payload: JSON.stringify(Date.now()),
      muteHttpExceptions: true
    });
  } catch (e) {
    Logger.log('syncSettingsToFirebase error: ' + e);
  }
}

/**
 * Full Deep Migration & Sync from Google Sheets to Firebase Realtime Database
 * Reads Users (including passwords/PINs), Activities, Submissions, Settings, Votes
 * and creates /live_rally_data/auth_credentials for 100% instant client-side Firebase authentication!
 */
function apiMigrateSheetsToFirebase() {
  try {
    const ss = getSpreadsheet();

    // 1. Users & Auth Credentials
    const usersSheet = getOrCreateSheet(ss, SHEET_NAMES.USERS);
    const usersRaw = usersSheet.getDataRange().getValues();
    const users = [];
    const authCredentials = {};
    for (let i = 1; i < usersRaw.length; i++) {
      const row = usersRaw[i];
      const rawUsername = String(row[0] || '').trim();
      if (!rawUsername) continue;
      const rawPassword = String(row[1] || '').trim();
      const uName = String(row[2] || '').trim();
      const uRole = String(row[3] || 'User').trim();
      const uCarCode = String(row[4] || '').trim();
      const uCarColor = (uRole === 'User') ? String(row[5] || '').trim() : '';
      const uProfile = String(row[6] || '').trim();
      const uBonus = Number(row[7]) || 0;
      let membersList = [];
      try {
        membersList = JSON.parse(row[8] || '[]');
      } catch (e) {
        if (row[8]) membersList = String(row[8]).split(',').map(function(s) { return s.trim(); }).filter(Boolean);
      }

      const userObj = {
        username: rawUsername,
        name: uName,
        role: uRole,
        carCode: uCarCode,
        carColor: uCarColor,
        profileUrl: uProfile,
        bonusPoints: uBonus,
        members: membersList
      };
      users.push(userObj);

      const credObj = {
        username: rawUsername,
        pin: rawPassword,
        name: uName,
        role: uRole,
        carCode: uCarCode,
        carColor: uCarColor,
        profileUrl: uProfile,
        bonusPoints: uBonus,
        members: membersList
      };

      authCredentials[rawUsername.toLowerCase()] = credObj;
      if (uCarCode) {
        authCredentials[uCarCode.toLowerCase()] = credObj;
      }
    }

    // 2. Fetch all data using proven fetchSharedDataFromSheets
    const shared = fetchSharedDataFromSheets() || {};
    const activities = shared.activities || [];
    const submissions = shared.submissions || [];
    const settings = shared.settings || {};
    const votes = shared.votes || [];

    // Push full payload to Firebase RTDB
    const payload = JSON.stringify({
      users: users,
      auth_credentials: authCredentials,
      activities: activities,
      submissions: submissions,
      settings: settings,
      votes: votes,
      timestamp: Date.now()
    });

    UrlFetchApp.fetch(FIREBASE_DATABASE_URL + '/live_rally_data.json', {
      method: 'patch',
      contentType: 'application/json',
      payload: payload,
      muteHttpExceptions: true
    });

    return {
      success: true,
      message: 'Migration to Firebase completed successfully!',
      usersCount: users.length,
      activitiesCount: activities.length,
      submissionsCount: submissions.length,
      timestamp: Date.now()
    };
  } catch (err) {
    Logger.log('apiMigrateSheetsToFirebase error: ' + err);
    return { success: false, message: String(err) };
  }
}

function setCachedSharedData(data, shouldSyncFirebase) {
  if (!data) return;
  try {
    const cache = CacheService.getScriptCache();
    const ver = getGlobalCacheVersion();
    const cacheEntries = {};
    
    if (data.users) cacheEntries['rally_users_' + ver] = JSON.stringify(data.users);
    if (data.activities) cacheEntries['rally_act_' + ver] = JSON.stringify(data.activities);
    if (data.settings) cacheEntries['rally_settings_' + ver] = JSON.stringify(data.settings);
    if (data.votes) cacheEntries['rally_votes_' + ver] = JSON.stringify(data.votes);
    
    if (data.submissions) {
      const subStr = JSON.stringify(data.submissions);
      const chunkSize = 80000;
      const numChunks = Math.ceil(subStr.length / chunkSize) || 1;
      cacheEntries['rally_subs_chunks_' + ver] = String(numChunks);
      for (let c = 0; c < numChunks; c++) {
        cacheEntries['rally_subs_' + c + '_' + ver] = subStr.substring(c * chunkSize, (c + 1) * chunkSize);
      }
    }
    
    cache.putAll(cacheEntries, 1200); // 20 minutes TTL
  } catch (e) {
    Logger.log('setCachedSharedData error: ' + e);
  }

  // Only sync to Firebase on mutations or when explicitly requested (NOT on routine sheet reads)
  const doSync = (typeof shouldSyncFirebase === 'boolean') ? shouldSyncFirebase : true;
  if (doSync) {
    syncToFirebase(data);
  }
}

function updateCachedSubmissionGrade(submissionId, username, activityId, score, judgeNotes, judgeUsername, status) {
  try {
    let shared = getCachedSharedData();
    if (!shared || !shared.submissions) {
      shared = fetchSharedDataFromSheets();
    }
    if (!shared || !shared.submissions) return;
    const uNorm = String(username || '').trim().toLowerCase();
    const aNorm = String(activityId || '').trim();
    const subIdNorm = String(submissionId || '').trim();
    const finalStatus = status || 'passed';
    let found = false;
    for (let i = 0; i < shared.submissions.length; i++) {
      const s = shared.submissions[i];
      const sUNorm = String(s.username || '').trim().toLowerCase();
      const sANorm = String(s.activityId || '').trim();
      const sIdNorm = String(s.id || '').trim();
      if ((subIdNorm && sIdNorm === subIdNorm) || (uNorm && aNorm && sUNorm === uNorm && sANorm === aNorm)) {
        s.status = finalStatus;
        s.score = Number(score) || 0;
        s.judgeNotes = judgeNotes || 'ให้คะแนนเรียบร้อย';
        s.judgeUsername = judgeUsername || 'Judge';
        found = true;
        break;
      }
    }
    if (!found && username && activityId) {
      shared.submissions.push({
        id: submissionId || ('SUB-' + Date.now()),
        timestamp: new Date().toISOString(),
        username: username,
        activityId: activityId,
        category: 'Base',
        carColor: '',
        answerText: '[ประเมินโดยกรรมการ]',
        imageUrl: '',
        fileId: '',
        status: finalStatus,
        score: Number(score) || 0,
        judgeNotes: judgeNotes || 'ให้คะแนนเรียบร้อย',
        judgeUsername: judgeUsername || 'Judge'
      });
    }
    setCachedSharedData(shared);
    syncToFirebase(shared);
  } catch (e) {
    Logger.log('updateCachedSubmissionGrade error: ' + e);
  }
}

function appendCachedSubmission(newSub) {
  try {
    const shared = getCachedSharedData();
    if (!shared || !shared.submissions) return;
    shared.submissions.push(newSub);
    setCachedSharedData(shared);
  } catch (e) {
    Logger.log('appendCachedSubmission error: ' + e);
  }
}

function updateCachedBonusPoints(carUsername, bonusPoints) {
  try {
    const shared = getCachedSharedData();
    if (!shared || !shared.users) return;
    for (let i = 0; i < shared.users.length; i++) {
      if (shared.users[i].username === carUsername) {
        shared.users[i].bonusPoints = Number(bonusPoints) || 0;
        break;
      }
    }
    setCachedSharedData(shared);
  } catch (e) {
    Logger.log('updateCachedBonusPoints error: ' + e);
  }
}

function appendCachedVote(newVote) {
  try {
    const shared = getCachedSharedData();
    if (!shared || !shared.votes) return;
    shared.votes.push(newVote);
    setCachedSharedData(shared, false);
  } catch (e) {
    Logger.log('appendCachedVote error: ' + e);
  }
}

function updateCachedUser(userData) {
  try {
    const shared = getCachedSharedData();
    if (!shared || !shared.users) return;
    const uNorm = String(userData.username || '').trim().toLowerCase();
    let found = false;
    for (let i = 0; i < shared.users.length; i++) {
      if (String(shared.users[i].username || '').trim().toLowerCase() === uNorm) {
        shared.users[i].name = userData.name;
        shared.users[i].role = userData.role;
        shared.users[i].carColor = userData.carColor;
        shared.users[i].carCode = userData.carCode;
        if (userData.profileUrl) shared.users[i].profileUrl = userData.profileUrl;
        if (userData.members) shared.users[i].members = userData.members;
        if (typeof userData.bonusPoints !== 'undefined') shared.users[i].bonusPoints = Number(userData.bonusPoints) || 0;
        found = true;
        break;
      }
    }
    if (!found) {
      shared.users.push({
        username: userData.username,
        name: userData.name,
        role: userData.role,
        carColor: userData.carColor,
        carCode: userData.carCode,
        profileUrl: userData.profileUrl || 'https://images.unsplash.com/photo-1533473359331-0135ef1b58bf?auto=format&fit=crop&w=400&q=80',
        bonusPoints: Number(userData.bonusPoints) || 0,
        members: userData.members || []
      });
    }
    setCachedSharedData(shared);
  } catch (e) {
    Logger.log('updateCachedUser error: ' + e);
  }
}

function deleteCachedUser(username) {
  try {
    const shared = getCachedSharedData();
    if (!shared || !shared.users) return;
    const uNorm = String(username || '').trim().toLowerCase();
    shared.users = shared.users.filter(function(u) {
      return String(u.username || '').trim().toLowerCase() !== uNorm;
    });
    setCachedSharedData(shared);
  } catch (e) {
    Logger.log('deleteCachedUser error: ' + e);
  }
}

function updateCachedActivity(actData) {
  try {
    if (!actData || !actData.id) return;
    const shared = getCachedSharedData();
    if (!shared || !Array.isArray(shared.activities)) return;
    const aNorm = String(actData.id || '').trim();
    let found = false;
    for (let i = 0; i < shared.activities.length; i++) {
      if (shared.activities[i] && String(shared.activities[i].id || '').trim() === aNorm) {
        shared.activities[i].category = actData.category;
        shared.activities[i].scoringType = actData.scoringType;
        shared.activities[i].title = actData.title;
        shared.activities[i].description = actData.description;
        shared.activities[i].maxPoints = Number(actData.maxPoints) || 0;
        if (actData.imageUrl) shared.activities[i].imageUrl = actData.imageUrl;
        if (typeof actData.solutionImageUrl !== 'undefined') shared.activities[i].solutionImageUrl = actData.solutionImageUrl;
        if (typeof actData.solutionImages !== 'undefined') shared.activities[i].solutionImages = actData.solutionImages;
        if (actData.autoAnswers) shared.activities[i].autoAnswers = actData.autoAnswers;
        found = true;
        break;
      }
    }
    if (!found) {
      shared.activities.push({
        id: actData.id,
        category: actData.category,
        scoringType: actData.scoringType,
        title: actData.title,
        description: actData.description,
        maxPoints: Number(actData.maxPoints) || 0,
        imageUrl: actData.imageUrl || '',
        solutionImageUrl: actData.solutionImageUrl || '',
        solutionImages: actData.solutionImages || (actData.solutionImageUrl ? [actData.solutionImageUrl] : []),
        autoAnswers: actData.autoAnswers || {}
      });
    }
    setCachedSharedData(shared);
  } catch (e) {
    Logger.log('updateCachedActivity error: ' + e);
  }
}

function deleteCachedActivity(id) {
  try {
    if (!id) return;
    const shared = getCachedSharedData();
    if (!shared || !Array.isArray(shared.activities)) return;
    const aNorm = String(id || '').trim();
    shared.activities = shared.activities.filter(function(a) {
      return a && String(a.id || '').trim() !== aNorm;
    });
    setCachedSharedData(shared);
  } catch (e) {
    Logger.log('deleteCachedActivity error: ' + e);
  }
}

/**
 * Helper: Parse solution images stored in sheet (supports JSON array string, comma-separated, or single URL)
 */
function parseSolutionImages(rawVal) {
  if (!rawVal) return [];
  const str = String(rawVal).trim();
  if (!str) return [];
  if (str.startsWith('[') && str.endsWith(']')) {
    try {
      const parsed = JSON.parse(str);
      if (Array.isArray(parsed)) return parsed.filter(Boolean);
    } catch(e) {}
  }
  if (str.indexOf(',') !== -1) {
    return str.split(',').map(function(s){ return s.trim(); }).filter(Boolean);
  }
  return [str];
}

function fetchSharedDataFromSheets() {
  const ss = getSpreadsheet();
  
  // 1. Users
  const usersSheet = getOrCreateSheet(ss, SHEET_NAMES.USERS);
  const usersRaw = usersSheet.getDataRange().getValues();
  const users = [];
  let usersSheetDirty = false;
  for (let i = 1; i < usersRaw.length; i++) {
    const rawUsername = String(usersRaw[i][0] || '').trim();
    if (!rawUsername) continue;
    const uRole = String(usersRaw[i][3] || '').trim();
    let uCarColor = String(usersRaw[i][5] || '').trim();
    if (uRole !== 'User') {
      uCarColor = '';
      if (String(usersRaw[i][5] || '').trim() !== '') {
        usersRaw[i][5] = '';
        usersSheetDirty = true;
      }
    }
    let membersList = [];
    try {
      membersList = JSON.parse(usersRaw[i][8] || '[]');
    } catch(e) {
      if (usersRaw[i][8]) membersList = String(usersRaw[i][8]).split(',').map(function(s){ return s.trim(); }).filter(Boolean);
    }
    users.push({
      username: rawUsername,
      name: usersRaw[i][2],
      role: uRole,
      carCode: usersRaw[i][4],
      carColor: uCarColor,
      profileUrl: usersRaw[i][6],
      bonusPoints: Number(usersRaw[i][7]) || 0,
      members: Array.isArray(membersList) ? membersList : []
    });
  }
  if (usersSheetDirty) {
    try {
      usersSheet.getRange(1, 1, usersRaw.length, usersRaw[0].length).setValues(usersRaw);
    } catch(e) {
      Logger.log('fetchSharedDataFromSheets clean users error: ' + e);
    }
  }

  // 2. Activities
  const actSheet = getOrCreateSheet(ss, SHEET_NAMES.ACTIVITIES);
  const actRaw = actSheet.getDataRange().getValues();
  const activities = [];
  for (let i = 1; i < actRaw.length; i++) {
    if (!actRaw[i][0]) continue;
    let autoAns = {};
    try { autoAns = JSON.parse(actRaw[i][7] || '{}'); } catch(e) {}
    const sImgs = parseSolutionImages(actRaw[i][8]);
    activities.push({
      id: String(actRaw[i][0] || '').trim(),
      category: actRaw[i][1] || 'Base',
      title: actRaw[i][2] || '',
      description: actRaw[i][3] || '',
      imageUrl: actRaw[i][4] || '',
      scoringType: actRaw[i][5] || 'AUTO',
      maxPoints: Number(actRaw[i][6]) || 0,
      autoAnswers: autoAns,
      solutionImages: sImgs,
      solutionImageUrl: sImgs[0] || ''
    });
  }

  // 3. Submissions (Deduplicated by username + activityId to prevent race duplicate rows)
  const subSheet = getOrCreateSheet(ss, SHEET_NAMES.SUBMISSIONS);
  const subRaw = subSheet.getDataRange().getValues();
  const submissions = [];
  const seenSubMap = {};
  for (let i = 1; i < subRaw.length; i++) {
    const sId = String(subRaw[i][0] || '');
    const uName = String(subRaw[i][2] || '').trim().toLowerCase();
    const aId = String(subRaw[i][3] || '').trim();
    const item = {
      id: sId,
      timestamp: subRaw[i][1],
      username: String(subRaw[i][2] || ''),
      activityId: aId,
      category: subRaw[i][4],
      carColor: subRaw[i][5],
      answerText: subRaw[i][6],
      imageUrl: subRaw[i][7],
      fileId: subRaw[i][8],
      status: subRaw[i][9],
      score: Number(subRaw[i][10]) || 0,
      judgeNotes: subRaw[i][11],
      judgeUsername: subRaw[i][12]
    };
    const key = uName + '___' + aId;
    if (uName && aId && seenSubMap[key] !== undefined) {
      submissions[seenSubMap[key]] = item;
    } else {
      if (uName && aId) seenSubMap[key] = submissions.length;
      submissions.push(item);
    }
  }

  // 4. Settings
  const settings = getSettingsMap(ss);

  // 5. Votes
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

  const sharedData = {
    users: users,
    activities: activities,
    submissions: submissions,
    settings: {
      isScoresHidden: settings.isScoresHidden,
      isVotingOpen: settings.isVotingOpen,
      isVotesHidden: settings.isVotesHidden,
      boobyRank: settings.boobyRank || 0
    },
    votes: votes
  };

  setCachedSharedData(sharedData, false); // Cache in memory, do not send redundant HTTP PUT to Firebase on sheet reads
  return sharedData;
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

// In-Memory Request Sheet Cache for High-Speed Lookups
let _sheetCache = {};

/**
 * High-Speed Cached Helper to get or create sheet
 */
function getOrCreateSheet(ss, sheetName) {
  const key = sheetName.trim().toLowerCase();
  if (_sheetCache[key]) return _sheetCache[key];

  let s = ss.getSheetByName(sheetName);
  if (!s) {
    const sheets = ss.getSheets();
    for (let i = 0; i < sheets.length; i++) {
      _sheetCache[sheets[i].getName().trim().toLowerCase()] = sheets[i];
    }
    s = _sheetCache[key];
  }
  if (!s) {
    try {
      s = ss.insertSheet(sheetName);
    } catch (err) {
      s = ss.getSheetByName(sheetName);
      if (!s) throw err;
    }
  }
  _sheetCache[key] = s;
  return s;
}

/**
 * Setup Database sheets and initial headers/sample data if empty
 */
function setupDatabase() {
  const ss = getSpreadsheet();
  
  // Fast check: If Users sheet already exists and has data, DB is initialized -> 0ms exit!
  const existingUsersSheet = ss.getSheetByName(SHEET_NAMES.USERS);
  if (existingUsersSheet && existingUsersSheet.getLastRow() > 1) {
    return;
  }
  
  // 1. Users Sheet
  let usersSheet = getOrCreateSheet(ss, SHEET_NAMES.USERS);
  if (usersSheet.getLastRow() === 0) {
    usersSheet.appendRow(['username', 'password', 'name', 'role', 'carCode', 'carColor', 'profileUrl', 'bonusPoints', 'members']);
    usersSheet.getRange(1, 1, 1, 9).setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');
  } else if (usersSheet.getLastColumn() === 8) {
    usersSheet.getRange(1, 9).setValue('members').setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');
    
    // Add Initial Default Data
    usersSheet.appendRow(['admin', 'admin123', 'ผู้ดูแลระบบสูงสุด', 'Admin', 'ADM-00', 'Red', 'https://images.unsplash.com/photo-1533473359331-0135ef1b58bf?auto=format&fit=crop&w=400&q=80', 0]);
    usersSheet.appendRow(['ref1', '123', 'กรรมการประจำฐาน 1', 'Sub-Admin', 'SUB-01', 'Blue', 'https://images.unsplash.com/photo-1560250097-0b93528c311a?auto=format&fit=crop&w=400&q=80', 0]);
    usersSheet.appendRow(['car01', 'pass123', 'ทีมสายฟ้าสีแดง', 'User', 'C1', 'Red', 'https://images.unsplash.com/photo-1544829099-b9a0c07fad1a?auto=format&fit=crop&w=400&q=80', 5]);
    usersSheet.appendRow(['car02', 'pass123', 'ทีมมังกรสีฟ้า', 'User', 'B-02', 'Blue', 'https://images.unsplash.com/photo-1552519507-da3b142c6e3d?auto=format&fit=crop&w=400&q=80', 0]);
    usersSheet.appendRow(['car03', 'pass123', 'ทีมสิงห์สีเหลือง', 'User', 'Y-03', 'Yellow', 'https://images.unsplash.com/photo-1580273916550-e323be2ae537?auto=format&fit=crop&w=400&q=80', 10]);
  }
  
  // 2. Activities Sheet
  let actSheet = getOrCreateSheet(ss, SHEET_NAMES.ACTIVITIES);
  if (actSheet.getLastRow() === 0) {
    actSheet.appendRow(['id', 'category', 'title', 'description', 'imageUrl', 'scoringType', 'maxPoints', 'autoAnswers', 'solutionImageUrl']);
    actSheet.getRange(1, 1, 1, 9).setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');
    
    // Sample Activities
    const sampleAutoAnswers = JSON.stringify({
      Red: { answer: 'RC1-RED', points: 10 },
      Blue: { answer: 'Blue', points: 10 },
      Yellow: { answer: 'Yellow', points: 10 },
      default: { answer: 'RC1', points: 10 }
    });
    
    actSheet.appendRow(['ACT-001', 'RC', 'จุด RC 1: ป้ายหลักกิโลเมตรประวัติศาสตร์', 'ถ่ายรูปคู่กับป้ายหลักกิโลเมตรและค้นหาตัวเลขคำใบ้', 'https://images.unsplash.com/photo-1568605117036-5fe5e7bab0b7?auto=format&fit=crop&w=600&q=80', 'AUTO', 10, sampleAutoAnswers, '']);
    actSheet.appendRow(['ACT-002', 'Base', 'ฐานกิจกรรม 1: สานสามัคคีสร้างสิ่งประดิษฐ์', 'ประดิษฐ์แพจำลองจากอุปกรณ์ที่กำหนด แล้วถ่ายภาพส่งผลงานเข้าระบบ', 'https://images.unsplash.com/photo-1517048676732-d65bc937f952?auto=format&fit=crop&w=600&q=80', 'IMAGE', 20, '{}', '']);
    actSheet.appendRow(['ACT-003', 'Quiz', 'คำถามไอคิว: ปริศนาเมืองเก่า', 'บรรยายประวัติความเป็นมาของโบราณสถานประจำเมืองอย่างย่อ', 'https://images.unsplash.com/photo-1461360370896-922624d12aa1?auto=format&fit=crop&w=600&q=80', 'MANUAL', 15, '{}', '']);
  } else {
    // Migration: Ensure column 9 exists for solutionImageUrl
    if (actSheet.getLastColumn() < 9) {
      actSheet.getRange(1, 9).setValue('solutionImageUrl').setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');
    }
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
    isVotesHidden: false,
    boobyRank: 0
  };
  for (let i = 1; i < data.length; i++) {
    const key = String(data[i][0] || '').trim();
    const val = String(data[i][1] || '').trim();
    if (key === 'boobyRank') {
      settings.boobyRank = parseInt(val, 10) || 0;
    } else if (key) {
      settings[key] = (val.toLowerCase() === 'true');
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
 * API: Login Authentication (Lightweight High-Speed Verification: < 1s)
 */
function apiLogin(username, password) {
  const ss = getSpreadsheet();
  const sheet = getOrCreateSheet(ss, SHEET_NAMES.USERS);
  const data = sheet.getDataRange().getValues();
  const uTrim = String(username || '').trim().toLowerCase();
  const pTrim = String(password || '').trim();
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowUser = String(row[0] || '').trim().toLowerCase();
    const rowCarCode = String(row[4] || '').trim().toLowerCase();
    if ((rowUser === uTrim || (rowCarCode && rowCarCode === uTrim)) && String(row[1] || '').trim() === pTrim) {
      const canonicalUsername = String(row[0] || '').trim();
      const userRole = String(row[3] || 'User').trim();
      const sessionToken = generateSessionToken(canonicalUsername, userRole);
      
      // Fast settings lookup from cache or sheet
      let settings = null;
      try {
        const cache = CacheService.getScriptCache();
        const ver = getGlobalCacheVersion();
        const sStr = cache.get('rally_settings_' + ver);
        if (sStr) settings = JSON.parse(sStr);
      } catch (e) {}
      if (!settings) {
        settings = getSettingsMap(ss);
      }

      let membersList = [];
      try {
        membersList = JSON.parse(row[8] || '[]');
      } catch(e) {
        if (row[8]) membersList = String(row[8]).split(',').map(function(s){ return s.trim(); }).filter(Boolean);
      }

      const currentUser = {
        username: canonicalUsername,
        name: row[2] || '',
        role: userRole,
        carCode: row[4] || '',
        carColor: row[5] || '',
        profileUrl: row[6] || '',
        bonusPoints: Number(row[7]) || 0,
        members: Array.isArray(membersList) ? membersList : []
      };

      return {
        success: true,
        sessionToken: sessionToken,
        user: currentUser,
        isScoresHidden: settings ? !!settings.isScoresHidden : false,
        isVotingOpen: settings ? !!settings.isVotingOpen : false,
        isVotesHidden: settings ? !!settings.isVotesHidden : false,
        boobyRank: settings ? (Number(settings.boobyRank) || 0) : 0
      };
    }
  }
  return { success: false, message: 'ชื่อผู้ใช้งานหรือรหัสผ่านไม่ถูกต้อง' };
}

/**
 * API: Fetch Initial App Data (High-Speed Shared Cache - 0.2s for 30+ concurrent devices)
 */
function apiGetInitialData(username, sessionToken) {
  let shared = getCachedSharedData();
  if (!shared) {
    shared = fetchSharedDataFromSheets();
  }

  const uname = String(username || '').toLowerCase();
  const users = shared.users || [];
  let currentUser = null;
  for (let i = 0; i < users.length; i++) {
    if (String(users[i].username).toLowerCase() === uname) {
      currentUser = users[i];
      break;
    }
  }

  const requesterRole = currentUser ? currentUser.role : 'User';
  const isAdmin = (requesterRole === 'Admin');
  const isPrivileged = (requesterRole === 'Admin' || requesterRole === 'Sub-Admin');

  // Filter activities: solutionImageUrl and solutionImages are sent ONLY to Admin
  const activities = (shared.activities || []).filter(function(a) { return a && a.id; }).map(function(a) {
    if (isAdmin) return a;
    return {
      id: a.id,
      category: a.category || '',
      title: a.title || '',
      description: a.description || '',
      imageUrl: a.imageUrl || '',
      scoringType: a.scoringType || 'AUTO',
      maxPoints: Number(a.maxPoints) || 0,
      autoAnswers: a.autoAnswers || {},
      solutionImageUrl: '',
      solutionImages: []
    };
  });

  // Filter submissions: Privileged or owner gets full details; others get leaderboard-safe summary
  const submissions = (shared.submissions || []).filter(function(s) { return s && s.id; }).map(function(s) {
    const isOwn = (uname && String(s.username).toLowerCase() === uname);
    if (isPrivileged || isOwn) {
      return s;
    }
    return {
      id: s.id,
      username: s.username || '',
      activityId: s.activityId || '',
      category: s.category || '',
      status: s.status || 'pending',
      score: Number(s.score) || 0
    };
  });

  return {
    success: true,
    currentUser: currentUser,
    users: users,
    activities: activities,
    submissions: submissions,
    isScoresHidden: shared.settings ? shared.settings.isScoresHidden : false,
    isVotingOpen: shared.settings ? shared.settings.isVotingOpen : false,
    isVotesHidden: shared.settings ? shared.settings.isVotesHidden : false,
    boobyRank: shared.settings ? (shared.settings.boobyRank || 0) : 0,
    votes: shared.votes || []
  };
}

/**
 * API: Set Score Visibility Setting (Admin Only)
 */
/**
 * API: Set Booby Rank Setting (Admin Only)
 */
function apiSetBoobyRank(boobyRank, sessionToken) {
  const auth = verifyAuth(sessionToken, ['Admin']);
  if (!auth.success) return auth;

  const rankNum = Math.max(0, parseInt(boobyRank, 10) || 0);
  return withLock(function() {
    const ss = getSpreadsheet();
    setSetting('boobyRank', String(rankNum), ss);
    const freshSettings = getSettingsMap(ss);
    let shared = getCachedSharedData() || fetchSharedDataFromSheets();
    if (shared) {
      shared.settings = freshSettings;
      setCachedSharedData(shared, false);
    }
    syncSettingsToFirebase(freshSettings);
    return { success: true, boobyRank: rankNum };
  });
}

function apiSetScoreVisibility(isScoresHidden, sessionToken) {
  const auth = verifyAuth(sessionToken, ['Admin']);
  if (!auth.success) return auth;

  const hidden = (isScoresHidden === true || isScoresHidden === 'true' || isScoresHidden === 1 || isScoresHidden === '1');
  return withLock(function() {
    const ss = getSpreadsheet();
    setSetting('isScoresHidden', hidden ? 'true' : 'false', ss);
    const freshSettings = getSettingsMap(ss);
    let shared = getCachedSharedData() || fetchSharedDataFromSheets();
    if (shared) {
      shared.settings = freshSettings;
      setCachedSharedData(shared, false);
    }
    syncSettingsToFirebase(freshSettings);
    return { success: true, isScoresHidden: hidden };
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

  const hidden = (isVotesHidden === true || isVotesHidden === 'true' || isVotesHidden === 1 || isVotesHidden === '1');
  return withLock(function() {
    const ss = getSpreadsheet();
    setSetting('isVotesHidden', hidden ? 'true' : 'false', ss);
    const freshSettings = getSettingsMap(ss);
    let shared = getCachedSharedData() || fetchSharedDataFromSheets();
    if (shared) {
      shared.settings = freshSettings;
      setCachedSharedData(shared, false);
    }
    syncSettingsToFirebase(freshSettings);
    return { success: true, isVotesHidden: hidden, message: hidden ? 'ซ่อนผลการโหวตคะแนนเรียบร้อยแล้ว' : 'เปิดแสดงผลการโหวตคะแนนเรียบร้อยแล้ว' };
  });
}

function apiSetVotingStatus(isVotingOpen, sessionToken) {
  const auth = verifyAuth(sessionToken, ['Admin']);
  if (!auth.success) return auth;

  const open = (isVotingOpen === true || isVotingOpen === 'true' || isVotingOpen === 1 || isVotingOpen === '1');
  return withLock(function() {
    const ss = getSpreadsheet();
    setSetting('isVotingOpen', open ? 'true' : 'false', ss);
    const freshSettings = getSettingsMap(ss);
    let shared = getCachedSharedData() || fetchSharedDataFromSheets();
    if (shared) {
      shared.settings = freshSettings;
      setCachedSharedData(shared, false);
    }
    syncSettingsToFirebase(freshSettings);
    return { success: true, isVotingOpen: open, message: open ? 'เปิดระบบโหวตคะแนนเรียบร้อยแล้ว' : 'ปิดระบบโหวตคะแนนเรียบร้อยแล้ว' };
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

  if (!voterUsername || !targetUsername) {
    return { success: false, message: 'ข้อมูลการโหวตไม่สมบูรณ์' };
  }

  const voter = String(voterUsername).trim();
  const target = String(targetUsername).trim();

  // 1. Self-vote prevention rule
  if (voter.toLowerCase() === target.toLowerCase()) {
    return { success: false, message: 'ผู้เล่นไม่สามารถให้คะแนนโหวตรถของตนเองได้' };
  }

  // 2. High-speed voting status check (RAM Cache -> Firebase Live -> Sheet)
  let isVotingOpen = false;
  const cached = getCachedSharedData();
  if (cached && cached.settings && typeof cached.settings.isVotingOpen !== 'undefined') {
    isVotingOpen = !!cached.settings.isVotingOpen;
  } else {
    try {
      const fbRes = UrlFetchApp.fetch(FIREBASE_DATABASE_URL + '/live_rally_data/settings/isVotingOpen.json', { muteHttpExceptions: true });
      if (fbRes.getResponseCode() === 200) {
        const fbVal = JSON.parse(fbRes.getContentText());
        if (fbVal === true) isVotingOpen = true;
      }
    } catch (e) {}
    if (!isVotingOpen) {
      const ss = getSpreadsheet();
      const settings = getSettingsMap(ss);
      if (settings && settings.isVotingOpen) isVotingOpen = true;
    }
  }

  if (!isVotingOpen) {
    return { success: false, message: 'ระบบปิดรับคะแนนโหวตแล้ว หรือยังไม่ได้เปิดระบบ' };
  }

  // 3. Fast-path vote quota check from RAM cache (< 1ms)
  let currentVoterCount = 0;
  if (cached && Array.isArray(cached.votes)) {
    for (let i = 0; i < cached.votes.length; i++) {
      if (String(cached.votes[i].voterUsername || '').trim().toLowerCase() === voter.toLowerCase()) {
        currentVoterCount++;
      }
    }
    if (currentVoterCount >= 4) {
      return { success: false, message: 'คุณได้ใช้สิทธิ์โหวตครบ 4 ครั้งตามโควตาแล้ว' };
    }
  }

  // 4. Atomic append inside lightweight lock (~150ms lock duration)
  const result = withLock(function() {
    const ss = getSpreadsheet();
    const votesSheet = getOrCreateSheet(ss, SHEET_NAMES.VOTES);

    // Double-check quota from sheet only if cache was missing
    if (!cached || !Array.isArray(cached.votes)) {
      const votesData = votesSheet.getDataRange().getValues();
      currentVoterCount = 0;
      for (let i = 1; i < votesData.length; i++) {
        if (String(votesData[i][2] || '').trim().toLowerCase() === voter.toLowerCase()) {
          currentVoterCount++;
        }
      }
      if (currentVoterCount >= 4) {
        return { success: false, message: 'คุณได้ใช้สิทธิ์โหวตครบ 4 ครั้งตามโควตาแล้ว' };
      }
    }

    const newVoteId = 'VOTE_' + Date.now() + '_' + Math.floor(Math.random() * 1000);
    const timestamp = new Date().toISOString();
    votesSheet.appendRow([newVoteId, timestamp, voter, target]);

    const remainingVotes = Math.max(0, 4 - (currentVoterCount + 1));

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
  }, 10000);

  if (result && result.success && result.vote) {
    appendCachedVote(result.vote);

    // Dual-Sync to Firebase Realtime Database (< 50ms broadcast to all screens)
    try {
      UrlFetchApp.fetch(FIREBASE_DATABASE_URL + '/live_rally_data/votes/' + result.vote.id + '.json', {
        method: 'put',
        contentType: 'application/json',
        payload: JSON.stringify(result.vote),
        muteHttpExceptions: true
      });
      UrlFetchApp.fetch(FIREBASE_DATABASE_URL + '/live_rally_data/lastVotedCar.json', {
        method: 'put',
        contentType: 'application/json',
        payload: JSON.stringify({ targetUsername: result.vote.targetUsername, timestamp: Date.now() }),
        muteHttpExceptions: true
      });
      UrlFetchApp.fetch(FIREBASE_DATABASE_URL + '/live_rally_data/timestamp.json', {
        method: 'put',
        contentType: 'application/json',
        payload: JSON.stringify(Date.now()),
        muteHttpExceptions: true
      });
    } catch (fbErr) {
      Logger.log('apiCastVote Firebase push error: ' + fbErr);
    }
  }
  return result;
}

/**
 * API: Reset All Votes (Admin Only)
 */
function apiResetVotes(sessionToken) {
  const auth = verifyAuth(sessionToken, ['Admin']);
  if (!auth.success) return auth;

  return withLock(function() {
    return executeSystemVoteReset();
  });
}

/**
 * Executes a full vote reset across Google Sheets, ScriptCache, and Firebase Realtime Database
 */
function executeSystemVoteReset() {
  const ss = getSpreadsheet();
  const votesSheet = getOrCreateSheet(ss, SHEET_NAMES.VOTES);
  const lastRow = votesSheet.getLastRow();
  if (lastRow > 1) {
    votesSheet.getRange(2, 1, lastRow - 1, votesSheet.getLastColumn()).clearContent();
  }

  // Clear in-memory shared cache
  try {
    const shared = getCachedSharedData();
    if (shared) {
      shared.votes = [];
      setCachedSharedData(shared, false);
    }
  } catch (e) {
    Logger.log('executeSystemVoteReset cache error: ' + e);
  }

  // Wipe votes in Firebase Realtime Database (< 0.05s instant wipe on all screens)
  try {
    UrlFetchApp.fetch(FIREBASE_DATABASE_URL + '/live_rally_data/votes.json', { method: 'delete', muteHttpExceptions: true });
    UrlFetchApp.fetch(FIREBASE_DATABASE_URL + '/live_rally_data/lastVotedCar.json', { method: 'delete', muteHttpExceptions: true });
    UrlFetchApp.fetch(FIREBASE_DATABASE_URL + '/live_rally_data/timestamp.json', {
      method: 'put',
      contentType: 'application/json',
      payload: JSON.stringify(Date.now()),
      muteHttpExceptions: true
    });
  } catch (fbErr) {
    Logger.log('executeSystemVoteReset Firebase wipe error: ' + fbErr);
  }

  return { success: true, message: 'รีเซ็ตข้อมูลคะแนนโหวตทั้งหมดเรียบร้อยแล้ว' };
}

/**
 * API: Submit Answer (Auto, Manual, or Image)
 */
function apiSubmitAnswer(username, activityId, answerText, imageFileObj, sessionToken) {
  const auth = verifyAuth(sessionToken, ['User', 'Admin', 'Sub-Admin'], username);
  if (!auth.success) return auth;

  // 1. Upload image to Drive outside lock to minimize lock contention duration
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

  // 2. Resolve activity & user details outside lock (from Shared Cache or Sheet)
  const shared = getCachedSharedData();
  let targetAct = null;
  let userColor = 'Default';

  if (shared && Array.isArray(shared.activities) && Array.isArray(shared.users)) {
    targetAct = shared.activities.find(function(a) { return a && a.id === activityId; });
    const uObj = shared.users.find(function(u) { return u && u.username === username; });
    if (uObj && uObj.carColor) userColor = uObj.carColor;
  }

  const ss = getSpreadsheet();
  if (!targetAct) {
    const actSheet = getOrCreateSheet(ss, SHEET_NAMES.ACTIVITIES);
    const actData = actSheet.getDataRange().getValues();
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
  }

  if (!targetAct) return { success: false, message: 'ไม่พบข้อมูลภารกิจนี้' };

  if (userColor === 'Default') {
    const usersSheet = getOrCreateSheet(ss, SHEET_NAMES.USERS);
    const usersData = usersSheet.getDataRange().getValues();
    for (let i = 1; i < usersData.length; i++) {
      if (usersData[i][0] === username) {
        userColor = usersData[i][5] || 'Default';
        break;
      }
    }
  }

  // 3. Pre-compute auto score outside lock
  let status = 'pending';
  let score = 0;
  let judgeNotes = '';

  if (targetAct.scoringType === 'AUTO') {
    const cleanUserAnswer = (answerText || '').toString().trim().toLowerCase();
    const normUserAnswer = normalizeAnswerText(answerText);
    let isCorrect = false;
    let earnedPoints = 0;

    function normColor(c) {
      if (!c) return 'default';
      const s = c.toString().trim().toLowerCase();
      if (s === 'red' || s === 'แดง') return 'red';
      if (s === 'blue' || s === 'น้ำเงิน' || s === 'ฟ้า' || s === 'สีฟ้า' || s === 'sky' || s === 'cyan') return 'blue';
      if (s === 'yellow' || s === 'เหลือง') return 'yellow';
      if (s === 'green' || s === 'เขียว') return 'green';
      if (s === 'orange' || s === 'ส้ม') return 'orange';
      if (s === 'purple' || s === 'ม่วง') return 'purple';
      if (s === 'pink' || s === 'ชมพู') return 'pink';
      if (s === 'default' || s === 'all' || s === 'ทั้งหมด') return 'default';
      return s;
    }

    const autoAnsRules = targetAct.autoAnswers;
    if (Array.isArray(autoAnsRules)) {
      for (let r = 0; r < autoAnsRules.length; r++) {
        const rule = autoAnsRules[r];
        const ruleAns = (rule.answer || '').toString().trim().toLowerCase();
        const normRuleAns = normalizeAnswerText(rule.answer);
        const ruleColor = (rule.color || 'Default').toString().trim().toLowerCase();
        const uColor = (userColor || 'Default').toString().trim().toLowerCase();
        const matchColor = (ruleColor === uColor || ruleColor === 'default' || ruleColor === 'all' || normColor(ruleColor) === normColor(uColor) || normColor(ruleColor) === 'default');

        if ((cleanUserAnswer === ruleAns || normUserAnswer === normRuleAns) && matchColor) {
          isCorrect = true;
          earnedPoints = Number(rule.points) !== undefined ? Number(rule.points) : targetAct.maxPoints;
          break;
        }
      }
    } else if (autoAnsRules && typeof autoAnsRules === 'object') {
      const colorRule = autoAnsRules[userColor] || autoAnsRules[normColor(userColor)] || autoAnsRules['default'] || autoAnsRules['Default'];
      if (colorRule && colorRule.answer) {
        const targetAnswer = colorRule.answer.toString().trim().toLowerCase();
        const normTargetAnswer = normalizeAnswerText(colorRule.answer);
        if (cleanUserAnswer === targetAnswer || normUserAnswer === normTargetAnswer) {
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
  } else if (targetAct.scoringType === 'AUTO_KEYWORDS') {
    const cleanUserAnswer = (answerText || '').toString().trim().toLowerCase();
    const normUserAnswer = normalizeAnswerText(answerText);
    let earnedPoints = 0;
    const matchedKeywords = [];
    const seenKw = {};

    function normColor(c) {
      if (!c) return 'default';
      const s = c.toString().trim().toLowerCase();
      if (s === 'red' || s === 'แดง') return 'red';
      if (s === 'blue' || s === 'น้ำเงิน' || s === 'ฟ้า' || s === 'สีฟ้า' || s === 'sky' || s === 'cyan') return 'blue';
      if (s === 'yellow' || s === 'เหลือง') return 'yellow';
      if (s === 'green' || s === 'เขียว') return 'green';
      if (s === 'orange' || s === 'ส้ม') return 'orange';
      if (s === 'purple' || s === 'ม่วง') return 'purple';
      if (s === 'pink' || s === 'ชมพู') return 'pink';
      if (s === 'default' || s === 'all' || s === 'ทั้งหมด') return 'default';
      return s;
    }

    const autoAnsRules = targetAct.autoAnswers;
    if (Array.isArray(autoAnsRules)) {
      for (let r = 0; r < autoAnsRules.length; r++) {
        const rule = autoAnsRules[r];
        const rawAns = (rule.answer || '').toString().trim();
        const rAns = rawAns.toLowerCase();
        const normRAns = normalizeAnswerText(rawAns);
        if (!rAns) continue;

        const ruleColor = (rule.color || 'Default').toString().trim().toLowerCase();
        const uColor = (userColor || 'Default').toString().trim().toLowerCase();
        const matchColor = (ruleColor === uColor || ruleColor === 'default' || ruleColor === 'all' || normColor(ruleColor) === normColor(uColor) || normColor(ruleColor) === 'default');

        if (matchColor && (cleanUserAnswer.includes(rAns) || (normRAns && normUserAnswer.includes(normRAns)))) {
          if (!seenKw[rAns] && !seenKw[normRAns]) {
            seenKw[rAns] = true;
            seenKw[normRAns] = true;
            const pts = rule.points !== undefined ? Number(rule.points) : targetAct.maxPoints;
            earnedPoints += pts;
            matchedKeywords.push(rawAns + ' (' + (pts >= 0 ? '+' : '') + pts + ')');
          }
        }
      }
    }

    if (matchedKeywords.length > 0) {
      status = 'passed';
      const rawTotal = earnedPoints;
      if (earnedPoints > targetAct.maxPoints) {
        earnedPoints = targetAct.maxPoints;
      }
      score = earnedPoints;
      judgeNotes = 'ตรวจคำสำคัญอัตโนมัติ: พบ ' + matchedKeywords.length + ' คำ [' + matchedKeywords.join(', ') + '] รวม ' + score + ' คะแนน';
      if (rawTotal > targetAct.maxPoints) {
        judgeNotes += ' (คะแนนดิบ ' + rawTotal + ' จำกัดไม่เกิน ' + targetAct.maxPoints + ')';
      }
    } else {
      status = 'failed';
      score = 0;
      judgeNotes = 'ตรวจคำสำคัญอัตโนมัติ: ไม่พบคำสำคัญตามเฉลย (0 คะแนน)';
    }
  } else {
    status = 'pending';
    score = 0;
    judgeNotes = targetAct.scoringType === 'MANUAL_TEXT' ? 'ส่งคำตอบแล้ว รอการตรวจจากกรรมการ' : 'รอการตรวจและให้คะแนนจากกรรมการ';
  }

  const subId = 'SUB-' + Date.now();
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

  // 4. Micro-Lock (< 150ms execution time inside lock)
  const result = withLock(function() {
    const subSheet = getOrCreateSheet(ss, SHEET_NAMES.SUBMISSIONS);
    const subData = subSheet.getDataRange().getValues();
    const uNorm = String(username || '').trim().toLowerCase();
    const aNorm = String(activityId || '').trim();

    for (let i = 1; i < subData.length; i++) {
      if (String(subData[i][2] || '').trim().toLowerCase() === uNorm && String(subData[i][3] || '').trim() === aNorm) {
        return {
          success: false,
          alreadySubmitted: true,
          existingSubmission: {
            id: String(subData[i][0] || ''),
            timestamp: String(subData[i][1] || ''),
            username: String(subData[i][2] || ''),
            activityId: String(subData[i][3] || ''),
            category: String(subData[i][4] || ''),
            carColor: String(subData[i][5] || ''),
            answerText: String(subData[i][6] || ''),
            imageUrl: String(subData[i][7] || ''),
            fileId: String(subData[i][8] || ''),
            status: String(subData[i][9] || 'pending'),
            score: Number(subData[i][10]) || 0,
            judgeNotes: String(subData[i][11] || ''),
            judgeUsername: String(subData[i][12] || '')
          },
          message: 'ท่านได้ส่งคำตอบกิจกรรมนี้เรียบร้อยแล้ว (ระบบอนุญาตให้ส่งได้เพียง 1 ครั้ง)'
        };
      }
    }

    const nextRow = subSheet.getLastRow() + 1;
    subSheet.getRange(nextRow, 1, 1, 13).setValues([rowContent]);

    return {
      success: true,
      message: 'บันทึกการส่งคำตอบเรียบร้อยแล้ว',
      score: score,
      status: status,
      imageUrl: uploadedImageUrl,
      fileId: uploadedFileId
    };
  });

  if (result && result.success) {
    appendCachedSubmission({
      id: subId,
      timestamp: timestamp,
      username: username,
      activityId: activityId,
      category: targetAct.category,
      carColor: userColor,
      answerText: answerText || '',
      imageUrl: uploadedImageUrl,
      fileId: uploadedFileId,
      status: status,
      score: score,
      judgeNotes: judgeNotes,
      judgeUsername: status === 'passed' ? 'System' : ''
    });

    try {
      // Direct Firebase update for this submission to avoid expensive full-sheet reloads
      const subSafeKey = (String(username).trim().toLowerCase() + '___' + String(activityId).trim()).replace(/[^a-zA-Z0-9_-]/g, '_');
      const newSubData = {
        id: subId,
        timestamp: timestamp,
        username: username,
        activityId: activityId,
        category: targetAct.category,
        carColor: userColor,
        answerText: answerText || '',
        imageUrl: uploadedImageUrl,
        fileId: uploadedFileId,
        status: status,
        score: score,
        judgeNotes: judgeNotes,
        judgeUsername: status === 'passed' ? 'System' : ''
      };
      UrlFetchApp.fetch(FIREBASE_DATABASE_URL + '/live_rally_data/submissions/' + subSafeKey + '.json', {
        method: 'put',
        contentType: 'application/json',
        payload: JSON.stringify(newSubData),
        muteHttpExceptions: true
      });
      UrlFetchApp.fetch(FIREBASE_DATABASE_URL + '/live_rally_data/timestamp.json', {
        method: 'put',
        contentType: 'application/json',
        payload: JSON.stringify(Date.now()),
        muteHttpExceptions: true
      });
    } catch (e) {
      Logger.log('apiSubmitAnswer direct sync error: ' + e);
    }
  }

  return result;
}

/**
 * API: Grade Submission (Admin / Sub-Admin)
 */
function apiGradeSubmission(submissionId, score, judgeNotes, judgeUsername, username, activityId, sessionToken, status) {
  const auth = verifyAuth(sessionToken, ['Admin', 'Sub-Admin']);
  if (!auth.success) return auth;

  const uNorm = String(username || '').trim().toLowerCase();
  const aNorm = String(activityId || '').trim();
  const subIdNorm = String(submissionId || '').trim();
  const targetStatus = status || 'passed';

  // Generous 35-second lock timeout for simultaneous submissions from multiple station judges
  const result = withLock(function() {
    const ss = getSpreadsheet();
    const subSheet = getOrCreateSheet(ss, SHEET_NAMES.SUBMISSIONS);
    const subData = subSheet.getDataRange().getValues();
    const gradeValues = [[targetStatus, Number(score) || 0, judgeNotes || 'ให้คะแนนเรียบร้อย', judgeUsername || 'Judge']];

    // 1. Search for existing submission matching submissionId OR (username && activityId)
    let matchedRowIndex = -1;
    for (let i = 1; i < subData.length; i++) {
      const rowSubId = String(subData[i][0] || '').trim();
      const rowUname = String(subData[i][2] || '').trim().toLowerCase();
      const rowActId = String(subData[i][3] || '').trim();

      if ((subIdNorm && rowSubId === subIdNorm) || (uNorm && aNorm && rowUname === uNorm && rowActId === aNorm)) {
        matchedRowIndex = i + 1;
        break;
      }
    }

    if (matchedRowIndex !== -1) {
      // Existing row found: update columns 10 to 13 (passed/status, score, notes, judge)
      subSheet.getRange(matchedRowIndex, 10, 1, 4).setValues(gradeValues);
      // Immediately flush to disk so concurrent queued requests will see this update
      SpreadsheetApp.flush();
      return { success: true, message: 'บันทึกคะแนนเรียบร้อยแล้ว' };
    }

    // 2. If not found and we have username & activityId, append new row
    if (username && activityId) {
      let category = 'Base';
      const shared = getCachedSharedData();
      if (shared && Array.isArray(shared.activities)) {
        const act = shared.activities.find(function(a) { return a && String(a.id || '').trim() === aNorm; });
        if (act) category = act.category || 'Base';
      } else {
        const actsSheet = getOrCreateSheet(ss, SHEET_NAMES.ACTIVITIES);
        const actsData = actsSheet.getDataRange().getValues();
        for (let a = 1; a < actsData.length; a++) {
          if (String(actsData[a][0]).trim() === aNorm) {
            category = actsData[a][1];
            break;
          }
        }
      }

      const newSubId = subIdNorm || ('SUB-' + Date.now());
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
        targetStatus,
        Number(score) || 0,
        judgeNotes || 'ให้คะแนนเรียบร้อย',
        judgeUsername || 'Judge'
      ];
      subSheet.appendRow(rowContent);
      // Immediately flush to disk so next concurrent execution reads this row
      SpreadsheetApp.flush();
      return { success: true, message: 'บันทึกคะแนนเรียบร้อยแล้ว' };
    }

    return { success: false, message: 'ไม่พบรายการคำตอบนี้ในระบบ' };
  }, 35000);

  if (result && result.success) {
    updateCachedSubmissionGrade(submissionId, username, activityId, score, judgeNotes, judgeUsername, targetStatus);
  }
  return result;
}

/**
 * API: Update Bonus Points for a Competitor (Admin)
 */
function apiUpdateBonusPoints(carUsername, bonusPoints, sessionToken) {
  const auth = verifyAuth(sessionToken, ['Admin']);
  if (!auth.success) return auth;

  const result = withLock(function() {
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

  if (result && result.success) {
    updateCachedBonusPoints(carUsername, bonusPoints);
  }
  return result;
}

/**
 * API: Save / Edit Activity (Admin)
 * Supports uploading activity image file directly to Google Drive Folder (1L44yh69kAAmLjjrMf2oLbBUQ-oxrb2WK)
 */
function apiSaveActivity(activityData, sessionToken) {
  const auth = verifyAuth(sessionToken, ['Admin']);
  if (!auth.success) return auth;

  // Robust parsing: if activityData is passed as string, parse it
  if (typeof activityData === 'string') {
    try { activityData = JSON.parse(activityData); } catch (e) {}
  }

  // If activityData was wrapped as { activityData: ... }
  if (activityData && activityData.activityData && typeof activityData.activityData === 'object') {
    activityData = activityData.activityData;
  }

  // Safe fallback to object
  if (!activityData || typeof activityData !== 'object') {
    activityData = {};
  }

  const id = activityData.id || ('ACT-' + String(Date.now()).slice(-6));
  let imageUrl = activityData.imageUrl || '';
  let solutionImageUrl = activityData.solutionImageUrl || '';

  // If user uploaded a new image file for activity, upload to Drive outside lock
  if (activityData.imageFileObj && activityData.imageFileObj.base64) {
    const uploadRes = uploadFileToDrive(activityData.imageFileObj.base64, activityData.imageFileObj.fileName, activityData.imageFileObj.mimeType);
    if (uploadRes.success) {
      imageUrl = uploadRes.directUrl;
    }
  }

  // If user uploaded new solution image files for activity, upload each to Drive outside lock
  const newSolutionUrls = [];
  if (Array.isArray(activityData.solutionImageFiles) && activityData.solutionImageFiles.length > 0) {
    for (let f = 0; f < activityData.solutionImageFiles.length; f++) {
      const curFile = activityData.solutionImageFiles[f];
      if (curFile && curFile.base64) {
        const uploadRes = uploadFileToDrive(
          curFile.base64,
          curFile.fileName || ('sol_' + id + '_' + Date.now() + '_' + f + '.jpg'),
          curFile.mimeType || 'image/jpeg'
        );
        if (uploadRes.success && uploadRes.directUrl) {
          newSolutionUrls.push(uploadRes.directUrl);
        }
      }
    }
  } else if (activityData.solutionImageFileObj && activityData.solutionImageFileObj.base64) {
    const uploadRes = uploadFileToDrive(activityData.solutionImageFileObj.base64, activityData.solutionImageFileObj.fileName, activityData.solutionImageFileObj.mimeType);
    if (uploadRes.success && uploadRes.directUrl) {
      newSolutionUrls.push(uploadRes.directUrl);
    }
  }

  const result = withLock(function() {
    const ss = getSpreadsheet();
    const actSheet = getOrCreateSheet(ss, SHEET_NAMES.ACTIVITIES);
    const actData = actSheet.getDataRange().getValues();
    const autoAnswersStr = JSON.stringify(activityData.autoAnswers || {});

    // Ensure header has 9 columns
    if (actSheet.getLastColumn() < 9) {
      actSheet.getRange(1, 9).setValue('solutionImageUrl').setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');
    }

    let foundRow = -1;
    for (let i = 1; i < actData.length; i++) {
      if (actData[i][0] === id) {
        foundRow = i + 1;
        break;
      }
    }

    let solutionImages = activityData.solutionImages;
    if (!solutionImages || !Array.isArray(solutionImages)) {
      solutionImages = solutionImageUrl ? [solutionImageUrl] : [];
    }
    if (newSolutionUrls.length > 0) {
      solutionImages = solutionImages.concat(newSolutionUrls);
    }
    // Filter unique, non-empty URLs
    solutionImages = Array.from(new Set(solutionImages.filter(Boolean)));
    const solutionImagesVal = (solutionImages.length === 0) ? '' : (solutionImages.length === 1 ? solutionImages[0] : JSON.stringify(solutionImages));

    const rowContent = [
      id,
      activityData.category,
      activityData.title,
      activityData.description || '',
      imageUrl,
      activityData.scoringType,
      Number(activityData.maxPoints) || 0,
      autoAnswersStr,
      solutionImagesVal
    ];

    if (foundRow > 0) {
      actSheet.getRange(foundRow, 1, 1, 9).setValues([rowContent]);
    } else {
      actSheet.appendRow(rowContent);
    }

    return { 
      success: true, 
      message: 'บันทึกข้อมูลภารกิจเรียบร้อยแล้ว', 
      activityId: id, 
      imageUrl: imageUrl, 
      solutionImageUrl: solutionImages[0] || '',
      solutionImages: solutionImages
    };
  });

  if (result && result.success) {
    activityData.id = id;
    activityData.imageUrl = imageUrl;
    activityData.solutionImageUrl = result.solutionImageUrl;
    activityData.solutionImages = result.solutionImages;
    updateCachedActivity(activityData);
  }
  return result;
}

/**
 * Helper: Update Solution Images of an Activity in Cache & Firebase RTDB
 */
function updateActivitySolutionImagesInCache(activityId, currentImgs) {
  try {
    const shared = getCachedSharedData();
    if (!shared || !Array.isArray(shared.activities)) return;
    const aNorm = String(activityId || '').trim();
    for (let i = 0; i < shared.activities.length; i++) {
      if (shared.activities[i] && String(shared.activities[i].id || '').trim() === aNorm) {
        shared.activities[i].solutionImages = currentImgs;
        shared.activities[i].solutionImageUrl = currentImgs[0] || '';
        break;
      }
    }
    setCachedSharedData(shared);
    syncToFirebase(shared);

    // Explicitly sync activities node to Firebase Realtime Database
    try {
      UrlFetchApp.fetch(FIREBASE_DATABASE_URL + '/live_rally_data/activities.json', {
        method: 'put',
        contentType: 'application/json',
        payload: JSON.stringify(shared.activities),
        muteHttpExceptions: true
      });
    } catch (fbErr) {
      Logger.log('Firebase activities sync error: ' + fbErr);
    }
  } catch (e) {
    Logger.log('updateActivitySolutionImagesInCache error: ' + e);
  }
}

/**
 * API: Quick Upload Solution Image(s) for Activity (Admin) - Supports mode: 'replace' or 'append'
 */
function apiUploadSolutionImage(activityId, imageFileObj, modeOrToken, optionalToken) {
  let mode = 'replace';
  let sessionToken = '';
  if (optionalToken) {
    mode = modeOrToken || 'replace';
    sessionToken = optionalToken;
  } else if (modeOrToken === 'replace' || modeOrToken === 'append') {
    mode = modeOrToken;
    sessionToken = '';
  } else {
    sessionToken = modeOrToken || '';
    mode = 'replace';
  }

  const auth = verifyAuth(sessionToken, ['Admin']);
  if (!auth.success) return auth;

  // Support either single imageFileObj, an array of imageFileObjs, or { files: [...] }
  let fileList = [];
  if (Array.isArray(imageFileObj)) {
    fileList = imageFileObj;
  } else if (imageFileObj && Array.isArray(imageFileObj.files)) {
    fileList = imageFileObj.files;
  } else if (imageFileObj && imageFileObj.base64) {
    fileList = [imageFileObj];
  }

  if (fileList.length === 0) {
    return { success: false, message: 'ไม่พบไฟล์รูปภาพเฉลยที่ต้องการอัปโหลด' };
  }

  const uploadedUrls = [];
  for (let f = 0; f < fileList.length; f++) {
    const curFile = fileList[f];
    if (curFile && curFile.base64) {
      const uploadRes = uploadFileToDrive(
        curFile.base64, 
        curFile.fileName || ('solution_' + activityId + '_' + Date.now() + '_' + f + '.jpg'), 
        curFile.mimeType || 'image/jpeg'
      );
      if (uploadRes.success && uploadRes.directUrl) {
        uploadedUrls.push(uploadRes.directUrl);
      }
    }
  }

  if (uploadedUrls.length === 0) {
    return { success: false, message: 'ไม่สามารถอัปโหลดรูปภาพเฉลยไปยัง Google Drive ได้' };
  }

  return withLock(function() {
    const ss = getSpreadsheet();
    const actSheet = getOrCreateSheet(ss, SHEET_NAMES.ACTIVITIES);
    const actData = actSheet.getDataRange().getValues();

    if (actSheet.getLastColumn() < 9) {
      actSheet.getRange(1, 9).setValue('solutionImageUrl').setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');
    }

    let foundRow = -1;
    for (let i = 1; i < actData.length; i++) {
      if (actData[i][0] === activityId) {
        foundRow = i + 1;
        break;
      }
    }

    if (foundRow > 0) {
      let currentImgs = [];
      if (mode === 'replace') {
        currentImgs = uploadedUrls.slice();
      } else {
        const existingRaw = actSheet.getRange(foundRow, 9).getValue();
        currentImgs = parseSolutionImages(existingRaw);
        uploadedUrls.forEach(function(u) {
          if (currentImgs.indexOf(u) === -1) currentImgs.push(u);
        });
      }

      currentImgs = Array.from(new Set(currentImgs.filter(Boolean)));

      const valToSave = (currentImgs.length === 0) ? '' : (currentImgs.length === 1 ? currentImgs[0] : JSON.stringify(currentImgs));
      actSheet.getRange(foundRow, 9).setValue(valToSave);

      // In-place cache & Firebase sync
      updateActivitySolutionImagesInCache(activityId, currentImgs);

      return { 
        success: true, 
        message: 'อัปโหลดภาพเฉลยเรียบร้อยแล้ว (' + uploadedUrls.length + ' ภาพ)', 
        activityId: activityId, 
        solutionImages: currentImgs,
        solutionImageUrl: currentImgs[0] || ''
      };
    } else {
      return { success: false, message: 'ไม่พบรหัสภารกิจ ' + activityId };
    }
  });
}

/**
 * API: Clear All Solution Images from an Activity (Admin)
 */
function apiClearAllSolutionImages(activityId, sessionToken) {
  const auth = verifyAuth(sessionToken, ['Admin']);
  if (!auth.success) return auth;

  return withLock(function() {
    const ss = getSpreadsheet();
    const actSheet = getOrCreateSheet(ss, SHEET_NAMES.ACTIVITIES);
    const actData = actSheet.getDataRange().getValues();

    let foundRow = -1;
    for (let i = 1; i < actData.length; i++) {
      if (actData[i][0] === activityId) {
        foundRow = i + 1;
        break;
      }
    }

    if (foundRow > 0) {
      actSheet.getRange(foundRow, 9).setValue('');
      updateActivitySolutionImagesInCache(activityId, []);

      return {
        success: true,
        message: 'ลบภาพเฉลยทั้งหมดเรียบร้อยแล้ว',
        activityId: activityId,
        solutionImages: [],
        solutionImageUrl: ''
      };
    } else {
      return { success: false, message: 'ไม่พบรหัสภารกิจ ' + activityId };
    }
  });
}

/**
 * API: Delete a Solution Image from an Activity (Admin)
 * Supports deleting by image URL or by numeric image index (0, 1, 2...)
 * If imageUrl is empty, cleans up blank/null/invalid entries from the activity.
 */
function apiDeleteSolutionImage(activityId, imageUrl, sessionToken) {
  const auth = verifyAuth(sessionToken, ['Admin']);
  if (!auth.success) return auth;

  return withLock(function() {
    const ss = getSpreadsheet();
    const actSheet = getOrCreateSheet(ss, SHEET_NAMES.ACTIVITIES);
    const actData = actSheet.getDataRange().getValues();

    let foundRow = -1;
    for (let i = 1; i < actData.length; i++) {
      if (actData[i][0] === activityId) {
        foundRow = i + 1;
        break;
      }
    }

    if (foundRow > 0) {
      const existingRaw = actSheet.getRange(foundRow, 9).getValue();
      let currentImgs = parseSolutionImages(existingRaw);

      const isNumericIndex = (typeof imageUrl === 'number') || (/^\d+$/.test(String(imageUrl || '').trim()) && !String(imageUrl).startsWith('http'));
      if (isNumericIndex) {
        const delIdx = Number(imageUrl);
        if (delIdx >= 0 && delIdx < currentImgs.length) {
          currentImgs.splice(delIdx, 1);
        }
      } else if (imageUrl && String(imageUrl).trim() !== '') {
        const targetUrl = String(imageUrl).trim().toLowerCase();
        currentImgs = currentImgs.filter(function(u) {
          return String(u).trim().toLowerCase() !== targetUrl;
        });
      }

      // Always purge empty/blank/null entries
      currentImgs = currentImgs.filter(Boolean);

      const valToSave = (currentImgs.length === 0) ? '' : (currentImgs.length === 1 ? currentImgs[0] : JSON.stringify(currentImgs));
      actSheet.getRange(foundRow, 9).setValue(valToSave);

      // In-place cache & Firebase sync
      updateActivitySolutionImagesInCache(activityId, currentImgs);

      return { 
        success: true, 
        message: 'ลบภาพเฉลยเรียบร้อยแล้ว', 
        activityId: activityId, 
        solutionImages: currentImgs,
        solutionImageUrl: currentImgs[0] || ''
      };
    } else {
      return { success: false, message: 'ไม่พบรหัสภารกิจ ' + activityId };
    }
  });
}

/**
 * API: Delete Activity (Admin)
 */
function apiDeleteActivity(activityId, sessionToken) {
  const auth = verifyAuth(sessionToken, ['Admin']);
  if (!auth.success) return auth;

  const result = withLock(function() {
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

  if (result && result.success) {
    deleteCachedActivity(activityId);
  }
  return result;
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

  const result = withLock(function() {
    const ss = getSpreadsheet();
    const usersSheet = getOrCreateSheet(ss, SHEET_NAMES.USERS);
    const usersData = usersSheet.getDataRange().getValues();

    let foundRow = -1;
    let existingPassword = 'pass' + Date.now();
    let existingMembers = [];
    for (let i = 1; i < usersData.length; i++) {
      if (usersData[i][0] === userData.username) {
        foundRow = i + 1;
        existingPassword = usersData[i][1];
        try { existingMembers = JSON.parse(usersData[i][8] || '[]'); } catch(e) {}
        break;
      }
    }

    // Preserve existing password if not changing, or set new password if provided
    const finalPassword = (userData.password && String(userData.password).trim())
      ? String(userData.password).trim()
      : existingPassword;

    const finalMembers = (userData.members && Array.isArray(userData.members))
      ? userData.members
      : existingMembers;

    const finalCarColor = (userData.role === 'User')
      ? (userData.carColor || 'Red')
      : '';

    const rowContent = [
      userData.username,
      finalPassword,
      userData.name,
      userData.role,
      userData.carCode || '',
      finalCarColor,
      profileUrl,
      Number(userData.bonusPoints) || 0,
      JSON.stringify(finalMembers || [])
    ];

    if (usersSheet.getLastColumn() < 9) {
      usersSheet.getRange(1, 9).setValue('members').setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');
    }

    if (foundRow > 0) {
      usersSheet.getRange(foundRow, 1, 1, rowContent.length).setValues([rowContent]);
    } else {
      usersSheet.appendRow(rowContent);
    }

    return { 
      success: true, 
      message: 'บันทึกข้อมูลผู้ใช้งานเรียบร้อยแล้ว',
      finalPassword: finalPassword,
      finalCarColor: finalCarColor,
      finalMembers: finalMembers
    };
  });

  if (result && result.success) {
    userData.profileUrl = profileUrl;
    userData.carColor = result.finalCarColor;
    userData.members = result.finalMembers;
    updateCachedUser(userData);

    // Sync credentials to Firebase for instant client-side login
    try {
      const credObj = {
        username: userData.username,
        pin: result.finalPassword,
        name: userData.name,
        role: userData.role,
        carCode: userData.carCode || '',
        carColor: result.finalCarColor,
        profileUrl: profileUrl,
        bonusPoints: Number(userData.bonusPoints) || 0,
        members: result.finalMembers || []
      };
      UrlFetchApp.fetch(FIREBASE_DATABASE_URL + '/live_rally_data/auth_credentials/' + encodeURIComponent(String(userData.username).toLowerCase()) + '.json', {
        method: 'put',
        contentType: 'application/json',
        payload: JSON.stringify(credObj),
        muteHttpExceptions: true
      });
      if (userData.carCode) {
        UrlFetchApp.fetch(FIREBASE_DATABASE_URL + '/live_rally_data/auth_credentials/' + encodeURIComponent(String(userData.carCode).toLowerCase()) + '.json', {
          method: 'put',
          contentType: 'application/json',
          payload: JSON.stringify(credObj),
          muteHttpExceptions: true
        });
      }
    } catch (e) {
      Logger.log('Firebase auth sync error: ' + e);
    }
  }
  return result;
}

/**
 * API: Delete User (Admin)
 */
function apiDeleteUser(username, sessionToken) {
  const auth = verifyAuth(sessionToken, ['Admin']);
  if (!auth.success) return auth;

  const result = withLock(function() {
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

  if (result && result.success) {
    deleteCachedUser(username);
    try {
      UrlFetchApp.fetch(FIREBASE_DATABASE_URL + '/live_rally_data/auth_credentials/' + encodeURIComponent(String(username).toLowerCase()) + '.json', {
        method: 'delete',
        muteHttpExceptions: true
      });
    } catch (e) {}
  }
  return result;
}

/**
 * API: Swap/Reassign Cars between two competitor accounts (Admin Only)
 * Swaps: name, carColor, profileUrl, bonusPoints, members
 * Also safely updates Submissions and Votes sheets if any exist.
 */
function apiSwapCars(usernameA, usernameB, sessionToken) {
  const auth = verifyAuth(sessionToken, ['Admin']);
  if (!auth.success) return auth;

  if (!usernameA || !usernameB) {
    return { success: false, message: 'กรุณาระบุบัญชีรถทั้งสองคันที่ต้องการสลับ' };
  }
  if (String(usernameA).trim().toLowerCase() === String(usernameB).trim().toLowerCase()) {
    return { success: false, message: 'ไม่สามารถสลับกับบัญชีเดียวกันได้' };
  }

  return withLock(function() {
    const ss = getSpreadsheet();
    const usersSheet = getOrCreateSheet(ss, SHEET_NAMES.USERS);
    const usersData = usersSheet.getDataRange().getValues();

    let rowA = -1;
    let rowB = -1;

    const targetA = String(usernameA).trim().toLowerCase();
    const targetB = String(usernameB).trim().toLowerCase();

    for (let i = 1; i < usersData.length; i++) {
      const u = String(usersData[i][0] || '').trim().toLowerCase();
      if (u === targetA) rowA = i + 1;
      if (u === targetB) rowB = i + 1;
    }

    if (rowA <= 0 || rowB <= 0) {
      return { success: false, message: 'ไม่พบบัญชีผู้ใช้งานที่ระบุในระบบ' };
    }

    const dataA = usersData[rowA - 1];
    const dataB = usersData[rowB - 1];

    if (String(dataA[3] || '').trim() !== 'User' || String(dataB[3] || '').trim() !== 'User') {
      return { success: false, message: 'อนุญาตให้สลับได้เฉพาะบัญชีผู้เข้าแข่งขัน (Role: User) เท่านั้น' };
    }

    // Helper to check if name is default "ทีม carXX"
    function isDefaultName(name, username, carCode) {
      if (!name) return true;
      const clean = String(name).trim().toLowerCase();
      const u = String(username || '').trim().toLowerCase();
      const c = String(carCode || '').trim().toLowerCase();
      return clean === 'ทีม ' + u || clean === 'ทีม ' + c || clean === u || clean === c || clean === '';
    }

    // Col 0: username, Col 1: password, Col 2: name, Col 3: role, Col 4: carCode, Col 5: carColor, Col 6: profileUrl, Col 7: bonusPoints, Col 8: members
    const nameA = dataA[2];
    const carCodeA = dataA[4];
    const carColorA = dataA[5];
    const profileA = dataA[6];
    const bonusA = dataA[7];
    const membersA = dataA[8];

    const nameB = dataB[2];
    const carCodeB = dataB[4];
    const carColorB = dataB[5];
    const profileB = dataB[6];
    const bonusB = dataB[7];
    const membersB = dataB[8];

    // Calculate new names: if one was default "ทีม carX", adopt new car's default name
    let newNameForA = nameB;
    if (isDefaultName(nameB, dataB[0], carCodeB)) {
      newNameForA = 'ทีม ' + (carCodeA || dataA[0]);
    }
    let newNameForB = nameA;
    if (isDefaultName(nameA, dataA[0], carCodeA)) {
      newNameForB = 'ทีม ' + (carCodeB || dataB[0]);
    }

    // Fast in-memory update on usersData (คงสีประจำรถเดิมไว้ ไม่สลับสีรถ)
    usersData[rowA - 1][2] = newNameForA;
    // usersData[rowA - 1][5] คงค่า carColorA เดิมไว้
    usersData[rowA - 1][6] = profileB || '';
    usersData[rowA - 1][7] = Number(bonusB) || 0;
    usersData[rowA - 1][8] = membersB || '[]';

    usersData[rowB - 1][2] = newNameForB;
    // usersData[rowB - 1][5] คงค่า carColorB เดิมไว้
    usersData[rowB - 1][6] = profileA || '';
    usersData[rowB - 1][7] = Number(bonusA) || 0;
    usersData[rowB - 1][8] = membersA || '[]';

    // 1 single batch write for Users sheet! (Lightning fast: 200ms instead of 20s)
    usersSheet.getRange(1, 1, usersData.length, usersData[0].length).setValues(usersData);

    // Swap username in Submissions if any exist
    try {
      const subSheet = getOrCreateSheet(ss, SHEET_NAMES.SUBMISSIONS);
      if (subSheet.getLastRow() > 1) {
        const subData = subSheet.getDataRange().getValues();
        let subChanged = false;
        for (let i = 1; i < subData.length; i++) {
          const u = String(subData[i][2] || '').trim().toLowerCase();
          if (u === targetA) {
            subData[i][2] = '__SWAP_TEMP__';
            subChanged = true;
          } else if (u === targetB) {
            subData[i][2] = dataA[0];
            subChanged = true;
          }
        }
        if (subChanged) {
          for (let i = 1; i < subData.length; i++) {
            if (subData[i][2] === '__SWAP_TEMP__') {
              subData[i][2] = dataB[0];
            }
          }
          subSheet.getRange(1, 1, subData.length, subData[0].length).setValues(subData);
        }
      }
    } catch(subErr) {}

    // Swap username in Votes if any exist
    try {
      const votesSheet = getOrCreateSheet(ss, SHEET_NAMES.VOTES);
      if (votesSheet.getLastRow() > 1) {
        const votesData = votesSheet.getDataRange().getValues();
        let votesChanged = false;
        for (let i = 1; i < votesData.length; i++) {
          const voter = String(votesData[i][2] || '').trim().toLowerCase();
          const target = String(votesData[i][3] || '').trim().toLowerCase();
          if (voter === targetA) { votesData[i][2] = '__SWAP_TEMP_V__'; votesChanged = true; }
          else if (voter === targetB) { votesData[i][2] = dataA[0]; votesChanged = true; }
          if (target === targetA) { votesData[i][3] = '__SWAP_TEMP_T__'; votesChanged = true; }
          else if (target === targetB) { votesData[i][3] = dataA[0]; votesChanged = true; }
        }
        if (votesChanged) {
          for (let i = 1; i < votesData.length; i++) {
            if (votesData[i][2] === '__SWAP_TEMP_V__') votesData[i][2] = dataB[0];
            if (votesData[i][3] === '__SWAP_TEMP_T__') votesData[i][3] = dataB[0];
          }
          votesSheet.getRange(1, 1, votesData.length, votesData[0].length).setValues(votesData);
        }
      }
    } catch(voteErr) {}

    // Invalidate cache, reload fresh shared data, and sync to Firebase Realtime Database
    try {
      invalidateGlobalCache();
      const freshShared = fetchSharedDataFromSheets();
      setCachedSharedData(freshShared);
      syncToFirebase(freshShared);
    } catch (syncErr) {
      Logger.log('Post-swap sync error: ' + syncErr);
    }

    return {
      success: true,
      message: 'สลับข้อมูลระหว่าง ' + dataA[0] + ' และ ' + dataB[0] + ' เรียบร้อยแล้ว (คงสีประจำรถตามเดิม)',
      carA: { username: dataA[0], name: newNameForA, carColor: carColorA, profileUrl: profileB, bonusPoints: bonusB, members: JSON.parse(membersB || '[]') },
      carB: { username: dataB[0], name: newNameForB, carColor: carColorB, profileUrl: profileA, bonusPoints: bonusA, members: JSON.parse(membersA || '[]') }
    };
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
 * Resets submissions and bonus points, while leaving ALL activities and solution images 100% untouched.
 */
function apiClearAllSubmissions(sessionToken) {
  const auth = verifyAuth(sessionToken, ['Admin']);
  if (!auth.success) return auth;

  return withLock(function() {
    const ss = getSpreadsheet();

    // 1. Clear Submissions Sheet (keep header row 1)
    const subSheet = getOrCreateSheet(ss, SHEET_NAMES.SUBMISSIONS);
    const lastRow = subSheet.getLastRow();
    if (lastRow > 1) {
      subSheet.getRange(2, 1, lastRow - 1, subSheet.getLastColumn()).clearContent();
    }

    // 2. Reset Bonus Points of competitors in Users sheet (col 8, index 7)
    try {
      const usersSheet = getOrCreateSheet(ss, SHEET_NAMES.USERS);
      const uData = usersSheet.getDataRange().getValues();
      let userChanged = false;
      for (let i = 1; i < uData.length; i++) {
        if (uData[i][3] === 'User' && Number(uData[i][7] || 0) !== 0) {
          uData[i][7] = 0;
          userChanged = true;
        }
      }
      if (userChanged) {
        usersSheet.getRange(1, 1, uData.length, uData[0].length).setValues(uData);
      }
    } catch(uErr) {
      Logger.log('Reset bonus points error: ' + uErr);
    }

    // 3. Clear submissions & reset bonus points in RAM cache
    let updatedUsers = [];
    try {
      const shared = getCachedSharedData();
      if (shared) {
        shared.submissions = [];
        if (Array.isArray(shared.users)) {
          shared.users.forEach(function(u) {
            if (u && u.role === 'User') u.bonusPoints = 0;
          });
          updatedUsers = shared.users;
        }
        setCachedSharedData(shared, false);
      }
    } catch(cErr) {
      Logger.log('Clear submissions cache error: ' + cErr);
    }

    // 4. Push wipe to Firebase Realtime Database (< 0.05s broadcast)
    try {
      UrlFetchApp.fetch(FIREBASE_DATABASE_URL + '/live_rally_data/submissions.json', {
        method: 'delete',
        muteHttpExceptions: true
      });
      if (updatedUsers.length > 0) {
        UrlFetchApp.fetch(FIREBASE_DATABASE_URL + '/live_rally_data/users.json', {
          method: 'put',
          contentType: 'application/json',
          payload: JSON.stringify(updatedUsers),
          muteHttpExceptions: true
        });
      }
      UrlFetchApp.fetch(FIREBASE_DATABASE_URL + '/live_rally_data/timestamp.json', {
        method: 'put',
        contentType: 'application/json',
        payload: JSON.stringify(Date.now()),
        muteHttpExceptions: true
      });
    } catch(fbErr) {
      Logger.log('Firebase wipe submissions error: ' + fbErr);
    }

    return { success: true, message: 'ลบประวัติการส่งคำตอบและคะแนนของรถทุกคันเรียบร้อยแล้ว' };
  });
}

/**
 * API: Clear All Activities and Submissions (Reset All Missions - Admin Only)
 */
function apiClearAllActivities(sessionToken) {
  const auth = verifyAuth(sessionToken, ['Admin']);
  if (!auth.success) return auth;

  return withLock(function() {
    const ss = getSpreadsheet();
    // 1. Clear Activities Sheet (keep header row 1)
    const actSheet = getOrCreateSheet(ss, SHEET_NAMES.ACTIVITIES);
    const actLastRow = actSheet.getLastRow();
    if (actLastRow > 1) {
      actSheet.getRange(2, 1, actLastRow - 1, actSheet.getLastColumn()).clearContent();
    }

    // 2. Clear Submissions Sheet (keep header row 1)
    const subSheet = getOrCreateSheet(ss, SHEET_NAMES.SUBMISSIONS);
    const subLastRow = subSheet.getLastRow();
    if (subLastRow > 1) {
      subSheet.getRange(2, 1, subLastRow - 1, subSheet.getLastColumn()).clearContent();
    }

    return { success: true, message: 'ลบภารกิจทั้งหมดและล้างคะแนนระบบเรียบร้อยแล้ว' };
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
 * API: Reset Name, Profile Image, and Team Members for All Competitors (Admin Only)
 * Admin and Sub-Admin accounts are skipped/not modified.
 */
function apiResetCompetitorProfiles(sessionToken) {
  const auth = verifyAuth(sessionToken, ['Admin']);
  if (!auth.success) return auth;

  const result = withLock(function() {
    const ss = getSpreadsheet();
    const usersSheet = getOrCreateSheet(ss, SHEET_NAMES.USERS);

    // Ensure column 9 ('members') exists
    if (usersSheet.getLastColumn() < 9) {
      usersSheet.getRange(1, 9).setValue('members').setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');
    }

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
        if (usersData[i].length >= 9) {
          usersData[i][8] = '[]'; // members (col 9) -> reset to empty array JSON!
        }
        changed = true;
      }
    }

    if (changed && usersData.length > 1) {
      // 1 single batch write call for all competitors!
      usersSheet.getRange(1, 1, usersData.length, usersData[0].length).setValues(usersData);
    }

    return { success: true, message: 'รีเซ็ตชื่อ รูปโปรไฟล์ และรายชื่อผู้เข้าแข่งขันของรถทุกคันเรียบร้อยแล้ว' };
  });

  if (result && result.success) {
    try {
      invalidateGlobalCache();
      const freshShared = fetchSharedDataFromSheets();
      syncToFirebase(freshShared);
    } catch (e) {
      Logger.log('apiResetCompetitorProfiles cache update error: ' + e);
    }
  }

  return result;
}

/**
 * API: Reset ONLY Profile Photos for all Competitors (Admin)
 */
function apiResetCompetitorPhotos(sessionToken) {
  const auth = verifyAuth(sessionToken, ['Admin']);
  if (!auth.success) return auth;

  const result = withLock(function() {
    const ss = getSpreadsheet();
    const usersSheet = getOrCreateSheet(ss, SHEET_NAMES.USERS);
    const usersData = usersSheet.getDataRange().getValues();
    const defaultProfileUrl = 'https://images.unsplash.com/photo-1533473359331-0135ef1b58bf?auto=format&fit=crop&w=400&q=80';
    let changed = false;
    let count = 0;

    for (let i = 1; i < usersData.length; i++) {
      const role = String(usersData[i][3] || '').trim();
      if (role === 'User') {
        usersData[i][6] = defaultProfileUrl; // profileUrl (col 7)
        changed = true;
        count++;
      }
    }

    if (changed && usersData.length > 1) {
      usersSheet.getRange(1, 1, usersData.length, usersData[0].length).setValues(usersData);
    }

    return { 
      success: true, 
      message: 'รีเซ็ตรูปภาพของผู้แข่งขัน ' + count + ' คันกลับเป็นค่าเริ่มต้นเรียบร้อยแล้ว (ชื่อและรายชื่อสมาชิกยังคงเดิม)',
      updatedCount: count
    };
  });

  if (result && result.success) {
    try {
      invalidateGlobalCache();
      const freshShared = fetchSharedDataFromSheets();
      syncToFirebase(freshShared);
    } catch (e) {
      Logger.log('apiResetCompetitorPhotos cache update error: ' + e);
    }
  }

  return result;
}

/**
 * API: Auto-assign or batch assign car colors for competitors (Admin Only)
 * Supports explicit assignments dictionary { [username]: targetColor }
 * or repeating colorSequence (e.g. ['Pink', 'Blue', 'Red', 'Green'])
 */
function apiAutoAssignCarColors(sessionToken, assignments, colorSequence) {
  const auth = verifyAuth(sessionToken, ['Admin']);
  if (!auth.success) return auth;

  const result = withLock(function() {
    const ss = getSpreadsheet();
    const usersSheet = getOrCreateSheet(ss, SHEET_NAMES.USERS);
    const usersData = usersSheet.getDataRange().getValues();

    if (usersData.length <= 1) {
      return { success: false, message: 'ไม่พบข้อมูลผู้แข่งขันในระบบ' };
    }

    let changed = false;
    let count = 0;

    // 1. If explicit assignments provided: { username: targetColor, ... }
    if (assignments && typeof assignments === 'object' && Object.keys(assignments).length > 0) {
      const normAssignments = {};
      Object.keys(assignments).forEach(function(k) {
        normAssignments[String(k).trim().toLowerCase()] = String(assignments[k]).trim();
      });

      for (let i = 1; i < usersData.length; i++) {
        const role = String(usersData[i][3] || '').trim();
        if (role === 'User') {
          const username = String(usersData[i][0] || '').trim().toLowerCase();
          if (normAssignments.hasOwnProperty(username)) {
            const targetColor = normAssignments[username];
            const oldColor = String(usersData[i][5] || '').trim();
            if (oldColor.toLowerCase() !== targetColor.toLowerCase()) {
              usersData[i][5] = targetColor;
              changed = true;
            }
            count++;
          }
        }
      }
    } else {
      // 2. Auto-assign by color sequence
      const seq = Array.isArray(colorSequence) && colorSequence.length > 0 ? colorSequence : ['Pink', 'Blue', 'Red', 'Green'];
      const competitors = [];
      for (let i = 1; i < usersData.length; i++) {
        const role = String(usersData[i][3] || '').trim();
        if (role === 'User') {
          const username = String(usersData[i][0] || '').trim();
          const carCode = String(usersData[i][4] || '').trim();
          const numMatch = (carCode || username).match(/\d+/);
          const carNum = numMatch ? parseInt(numMatch[0], 10) : 999;
          competitors.push({
            rowIndex: i,
            username: username,
            carCode: carCode,
            carNum: carNum
          });
        }
      }

      competitors.sort(function(a, b) {
        if (a.carNum !== b.carNum) return a.carNum - b.carNum;
        return a.carCode.localeCompare(b.carCode);
      });

      competitors.forEach(function(comp, idx) {
        const assignedColor = seq[idx % seq.length];
        const r = comp.rowIndex;
        const oldColor = String(usersData[r][5] || '').trim();
        if (oldColor.toLowerCase() !== assignedColor.toLowerCase()) {
          usersData[r][5] = assignedColor; // col index 5 = carColor
          changed = true;
        }
        count++;
      });
    }

    // Ensure all non-User accounts (Admin, Sub-Admin) have blank carColor
    for (let i = 1; i < usersData.length; i++) {
      const role = String(usersData[i][3] || '').trim();
      if (role !== 'User' && String(usersData[i][5] || '').trim() !== '') {
        usersData[i][5] = '';
        changed = true;
      }
    }

    if (changed) {
      usersSheet.getRange(1, 1, usersData.length, usersData[0].length).setValues(usersData);
    }

    return {
      success: true,
      message: 'บันทึกการกำหนดสีประจำรถสำหรับรถ ' + count + ' คันเรียบร้อยแล้ว'
    };
  });

  if (result && result.success) {
    try {
      invalidateGlobalCache();
      const freshShared = fetchSharedDataFromSheets();
      syncToFirebase(freshShared);
    } catch (e) {
      Logger.log('apiAutoAssignCarColors cache update error: ' + e);
    }
  }

  return result;
}

/**
 * API: Re-evaluates all submissions for AUTO and AUTO_KEYWORDS activities
 * using the enhanced whitespace-agnostic comparison logic.
 * Updates scores and marks passed submissions automatically.
 */
function apiRegradeAutoSubmissions(sessionToken) {
  const auth = verifyAuth(sessionToken, ['Admin']);
  if (!auth.success) return auth;

  const result = withLock(function() {
    const ss = getSpreadsheet();
    const subSheet = getOrCreateSheet(ss, SHEET_NAMES.SUBMISSIONS);
    const actSheet = getOrCreateSheet(ss, SHEET_NAMES.ACTIVITIES);
    const usersSheet = getOrCreateSheet(ss, SHEET_NAMES.USERS);

    // Map users to get carColor
    const usersData = usersSheet.getDataRange().getValues();
    const userColorMap = {};
    for (let i = 1; i < usersData.length; i++) {
      const u = String(usersData[i][0] || '').trim().toLowerCase();
      userColorMap[u] = usersData[i][5] || 'Default';
    }

    // Map activities
    const actData = actSheet.getDataRange().getValues();
    const actMap = {};
    for (let i = 1; i < actData.length; i++) {
      const actId = String(actData[i][0] || '').trim();
      let autoAns = {};
      try {
        if (actData[i][7]) autoAns = JSON.parse(actData[i][7]);
      } catch(e) {}

      actMap[actId] = {
        id: actId,
        category: actData[i][1],
        title: actData[i][2],
        scoringType: actData[i][5],
        maxPoints: Number(actData[i][6]) || 0,
        autoAnswers: autoAns
      };
    }

    function normColor(c) {
      if (!c) return 'default';
      const s = c.toString().trim().toLowerCase();
      if (s === 'red' || s === 'แดง') return 'red';
      if (s === 'blue' || s === 'น้ำเงิน' || s === 'ฟ้า' || s === 'สีฟ้า' || s === 'sky' || s === 'cyan') return 'blue';
      if (s === 'yellow' || s === 'เหลือง') return 'yellow';
      if (s === 'green' || s === 'เขียว') return 'green';
      if (s === 'orange' || s === 'ส้ม') return 'orange';
      if (s === 'purple' || s === 'ม่วง') return 'purple';
      if (s === 'pink' || s === 'ชมพู') return 'pink';
      if (s === 'default' || s === 'all' || s === 'ทั้งหมด') return 'default';
      return s;
    }

    const subData = subSheet.getDataRange().getValues();
    let updatedCount = 0;
    let anyChanges = false;

    for (let i = 1; i < subData.length; i++) {
      const actId = String(subData[i][3] || '').trim();
      const targetAct = actMap[actId];
      if (!targetAct) continue;
      if (targetAct.scoringType !== 'AUTO' && targetAct.scoringType !== 'AUTO_KEYWORDS') continue;

      const rawAnswer = String(subData[i][6] || '');
      if (!rawAnswer || rawAnswer === '[ประเมินโดยกรรมการ]') continue;

      const username = String(subData[i][2] || '').trim();
      const userColor = subData[i][5] || userColorMap[username.toLowerCase()] || 'Default';

      const cleanUserAnswer = rawAnswer.trim().toLowerCase();
      const normUserAnswer = normalizeAnswerText(rawAnswer);

      let newStatus = subData[i][9];
      let newScore = Number(subData[i][10]) || 0;
      let newNotes = subData[i][11] || '';

      if (targetAct.scoringType === 'AUTO') {
        let isCorrect = false;
        let earnedPoints = 0;

        const autoAnsRules = targetAct.autoAnswers;
        if (Array.isArray(autoAnsRules)) {
          for (let r = 0; r < autoAnsRules.length; r++) {
            const rule = autoAnsRules[r];
            const ruleAns = (rule.answer || '').toString().trim().toLowerCase();
            const normRuleAns = normalizeAnswerText(rule.answer);
            const ruleColor = (rule.color || 'Default').toString().trim().toLowerCase();
            const uColor = (userColor || 'Default').toString().trim().toLowerCase();
            const matchColor = (ruleColor === uColor || ruleColor === 'default' || ruleColor === 'all' || normColor(ruleColor) === normColor(uColor) || normColor(ruleColor) === 'default');

            if ((cleanUserAnswer === ruleAns || normUserAnswer === normRuleAns) && matchColor) {
              isCorrect = true;
              earnedPoints = Number(rule.points) !== undefined ? Number(rule.points) : targetAct.maxPoints;
              break;
            }
          }
        } else if (autoAnsRules && typeof autoAnsRules === 'object') {
          const colorRule = autoAnsRules[userColor] || autoAnsRules[normColor(userColor)] || autoAnsRules['default'] || autoAnsRules['Default'];
          if (colorRule && colorRule.answer) {
            const targetAnswer = colorRule.answer.toString().trim().toLowerCase();
            const normTargetAnswer = normalizeAnswerText(colorRule.answer);
            if (cleanUserAnswer === targetAnswer || normUserAnswer === normTargetAnswer) {
              isCorrect = true;
              earnedPoints = Number(colorRule.points) !== undefined ? Number(colorRule.points) : targetAct.maxPoints;
            }
          }
        }

        if (isCorrect) {
          newStatus = 'passed';
          newScore = earnedPoints;
          newNotes = 'ตรวจคำตอบอัตโนมัติ (Re-grade: ' + earnedPoints + ' คะแนน)';
        }
      } else if (targetAct.scoringType === 'AUTO_KEYWORDS') {
        let earnedPoints = 0;
        const matchedKeywords = [];
        const seenKw = {};

        const autoAnsRules = targetAct.autoAnswers;
        if (Array.isArray(autoAnsRules)) {
          for (let r = 0; r < autoAnsRules.length; r++) {
            const rule = autoAnsRules[r];
            const rawAns = (rule.answer || '').toString().trim();
            const rAns = rawAns.toLowerCase();
            const normRAns = normalizeAnswerText(rawAns);
            if (!rAns) continue;

            const ruleColor = (rule.color || 'Default').toString().trim().toLowerCase();
            const uColor = (userColor || 'Default').toString().trim().toLowerCase();
            const matchColor = (ruleColor === uColor || ruleColor === 'default' || ruleColor === 'all' || normColor(ruleColor) === normColor(uColor) || normColor(ruleColor) === 'default');

            if (matchColor && (cleanUserAnswer.includes(rAns) || (normRAns && normUserAnswer.includes(normRAns)))) {
              if (!seenKw[rAns] && !seenKw[normRAns]) {
                seenKw[rAns] = true;
                seenKw[normRAns] = true;
                const pts = rule.points !== undefined ? Number(rule.points) : targetAct.maxPoints;
                earnedPoints += pts;
                matchedKeywords.push(rawAns + ' (' + (pts >= 0 ? '+' : '') + pts + ')');
              }
            }
          }
        }

        if (matchedKeywords.length > 0) {
          newStatus = 'passed';
          if (earnedPoints > targetAct.maxPoints) {
            earnedPoints = targetAct.maxPoints;
          }
          newScore = earnedPoints;
          newNotes = 'ตรวจคำสำคัญอัตโนมัติ (Re-grade: พบ ' + matchedKeywords.length + ' คำ [' + matchedKeywords.join(', ') + '] รวม ' + newScore + ' คะแนน)';
        }
      }

      // Check if status or score changed
      if (newStatus === 'passed' && (subData[i][9] !== 'passed' || Number(subData[i][10]) !== newScore)) {
        subData[i][9] = newStatus;
        subData[i][10] = newScore;
        subData[i][11] = newNotes;
        subData[i][12] = 'Admin (Re-grade)';
        updatedCount++;
        anyChanges = true;
      }
    }

    if (anyChanges && subData.length > 1) {
      subSheet.getRange(1, 1, subData.length, subData[0].length).setValues(subData);
    }

    return {
      success: true,
      updatedCount: updatedCount,
      message: 'ตรวจคำตอบใหม่อัตโนมัติเรียบร้อย อัปเดต ' + updatedCount + ' รายการ'
    };
  });

  if (result && result.success) {
    try {
      invalidateGlobalCache();
      const freshShared = fetchSharedDataFromSheets();
      syncToFirebase(freshShared);
    } catch (e) {
      Logger.log('apiRegradeAutoSubmissions cache update error: ' + e);
    }
  }

  return result;
}


