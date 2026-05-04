// ════════════════════════════════════════════════
// PROJECT SOLOMON · Google Apps Script v7
// Execute as: Me  |  Access: Anyone
// ════════════════════════════════════════════════

var SS = SpreadsheetApp.getActiveSpreadsheet();
var ANTHROPIC_API_KEY = 'YOUR_ANTHROPIC_API_KEY_HERE';

function resp(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  var p = e.parameter || {};
  return handleAction(p);
}

function doPost(e) {
  var p = {};
  try { p = JSON.parse(e.postData.contents); } catch(x) { p = e.parameter || {}; }
  return handleAction(p);
}

function handleAction(p) {
  var result;
  try {
    var a = p.action || '';
    if      (a === 'getStudent')       result = getStudent(p.email);
    else if (a === 'getStudents')      result = getAllStudents();
    else if (a === 'getQuizzes')       result = getQuizzes(p.section);
    else if (a === 'getQuizById')      result = getQuizById(p.id);
    else if (a === 'getScores')        result = getAllScores();
    else if (a === 'getStudentScores') result = getStudentScores(p.email);
    else if (a === 'addStudent')       result = addStudent(p);
    else if (a === 'updateStudent')    result = updateStudent(p);
    else if (a === 'addQuiz')          result = addQuiz(p);
    else if (a === 'updateQuiz')       result = updateQuiz(p);
    else if (a === 'addScore')         result = addScore(p);
    else if (a === 'fixSheets')        result = fixSheets();
    else if (a === 'saveLesson')       result = saveLesson(p);
    else if (a === 'getLesson')        result = getLesson(p.quizId);
    else if (a === 'addQuizFull')      result = addQuizFull(p);
    else result = { ok: true, status: 'Solomon API v7 Online' };
  } catch(err) {
    result = { ok: false, error: err.message };
  }
  return resp(result);
}

// ══════════════════════════════════════════════
// LESSON DATA — stored in dedicated LessonData sheet
// ══════════════════════════════════════════════

function getLessonSheet() {
  var sh = SS.getSheetByName('LessonData');
  if (!sh) {
    sh = SS.insertSheet('LessonData');
    sh.appendRow(['quizId', 'lessonData', 'savedAt']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function saveLesson(p) {
  var quizId = String(p.quizId || '').trim();
  var lessonData = p.lessonData || '';
  var chunk = p.chunk !== undefined ? parseInt(p.chunk) : -1;
  var totalChunks = p.totalChunks !== undefined ? parseInt(p.totalChunks) : 1;

  if (!quizId) return { ok: false, error: 'quizId required' };
  if (!lessonData) return { ok: false, error: 'lessonData empty' };

  var sh = getLessonSheet();
  var vals = sh.getDataRange().getValues();

  if (chunk === -1 || totalChunks === 1) {
    // Single chunk — save directly
    for (var i = 1; i < vals.length; i++) {
      if (String(vals[i][0]).trim() === quizId) {
        sh.getRange(i + 1, 2).setValue(lessonData);
        sh.getRange(i + 1, 3).setValue(new Date().toISOString());
        return { ok: true, updated: true };
      }
    }
    sh.appendRow([quizId, lessonData, new Date().toISOString()]);
    return { ok: true, inserted: true };
  }

  // Multi-chunk: save to a temp column, assemble when all chunks arrive
  // Use a ChunkCache sheet to store chunks temporarily
  var cacheSheet = SS.getSheetByName('ChunkCache');
  if (!cacheSheet) {
    cacheSheet = SS.insertSheet('ChunkCache');
    cacheSheet.appendRow(['key', 'data', 'ts']);
    cacheSheet.setFrozenRows(1);
  }

  // Save this chunk with key = quizId_chunkN
  var chunkKey = quizId + '_chunk' + chunk;
  var cacheVals = cacheSheet.getDataRange().getValues();
  var found = false;
  for (var ci = 1; ci < cacheVals.length; ci++) {
    if (String(cacheVals[ci][0]) === chunkKey) {
      cacheSheet.getRange(ci + 1, 2).setValue(lessonData);
      found = true; break;
    }
  }
  if (!found) cacheSheet.appendRow([chunkKey, lessonData, new Date().toISOString()]);

  // Check if all chunks are present
  var assembled = '';
  var allPresent = true;
  cacheVals = cacheSheet.getDataRange().getValues(); // refresh
  for (var n = 0; n < totalChunks; n++) {
    var key = quizId + '_chunk' + n;
    var found2 = false;
    for (var ci2 = 1; ci2 < cacheVals.length; ci2++) {
      if (String(cacheVals[ci2][0]) === key) {
        assembled += String(cacheVals[ci2][1]);
        found2 = true; break;
      }
    }
    if (!found2) { allPresent = false; break; }
  }

  if (allPresent) {
    // All chunks received — save assembled lesson
    var shVals = sh.getDataRange().getValues();
    var existsRow = -1;
    for (var si = 1; si < shVals.length; si++) {
      if (String(shVals[si][0]).trim() === quizId) { existsRow = si + 1; break; }
    }
    if (existsRow > 0) {
      sh.getRange(existsRow, 2).setValue(assembled);
      sh.getRange(existsRow, 3).setValue(new Date().toISOString());
    } else {
      sh.appendRow([quizId, assembled, new Date().toISOString()]);
    }
    // Clean up chunk cache
    for (var n2 = 0; n2 < totalChunks; n2++) {
      var key2 = quizId + '_chunk' + n2;
      var cacheVals2 = cacheSheet.getDataRange().getValues();
      for (var ci3 = 1; ci3 < cacheVals2.length; ci3++) {
        if (String(cacheVals2[ci3][0]) === key2) {
          cacheSheet.deleteRow(ci3 + 1); break;
        }
      }
    }
    return { ok: true, assembled: true, size: assembled.length };
  }

  return { ok: true, chunk_saved: chunk, waiting_for_more: true };
}

function getLesson(quizId) {
  quizId = String(quizId || '').trim();
  if (!quizId) return { ok: false, error: 'quizId required' };
  
  try {
    var sh = getLessonSheet();
    var vals = sh.getDataRange().getValues();
    for (var i = 1; i < vals.length; i++) {
      if (String(vals[i][0]).trim() === quizId) {
        var ld = String(vals[i][1] || '');
        if (ld.length > 5) return { ok: true, lessonData: ld };
      }
    }
    return { ok: false, error: 'not_found' };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ══════════════════════════════
// SHEET HELPERS
// ══════════════════════════════
function getSheetData(sheetName) {
  var sh = SS.getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 1) return { sh: sh, headers: [], rows: [] };
  var all = sh.getDataRange().getValues();
  var headers = all[0].map(function(h) { return String(h).trim().toLowerCase(); });
  var rows = all.slice(1);
  return { sh: sh, headers: headers, rows: rows };
}
function col(headers, name) { return headers.indexOf(name.toLowerCase()); }
function getCell(row, headers, name) { var i = col(headers, name); return i >= 0 ? row[i] : ''; }

// ══════════════════════════════
// STUDENTS
// ══════════════════════════════
function getStudentSheet() {
  var sh = SS.getSheetByName('Students');
  if (!sh) {
    sh = SS.insertSheet('Students');
    sh.appendRow(['email','name','section','avatar','status','xp','lootBags','createdAt']);
    sh.setFrozenRows(1);
    return sh;
  }
  // Fix header if needed
  if (sh.getLastRow() >= 1) {
    var first = String(sh.getRange(1,1).getValue()||'').toLowerCase().trim();
    if (first !== 'email') {
      if (first === 'id') sh.deleteColumn(1);
      sh.insertRowBefore(1);
      sh.getRange(1,1,1,8).setValues([['email','name','section','avatar','status','xp','lootBags','createdAt']]);
      sh.setFrozenRows(1);
    }
  }
  return sh;
}

function addStudent(d) {
  var sh = getStudentSheet();
  var email = String(d.email||'').toLowerCase().trim();
  if (!email) return { ok: false, error: 'email_required' };
  var data = getSheetData('Students');
  for (var i = 0; i < data.rows.length; i++) {
    if (String(getCell(data.rows[i], data.headers, 'email')||'').toLowerCase().trim() === email)
      return { ok: false, error: 'duplicate' };
  }
  sh.appendRow([email, String(d.name||''), String(d.section||''), String(d.avatar||''),
    'pending', 0, 0, String(d.createdAt||new Date().toISOString())]);
  return { ok: true };
}

function updateStudent(d) {
  var sh = getStudentSheet();
  var data = getSheetData('Students');
  var email = String(d.email||'').toLowerCase().trim();
  for (var i = 0; i < data.rows.length; i++) {
    if (String(getCell(data.rows[i], data.headers, 'email')||'').toLowerCase().trim() === email) {
      var r = i + 2;
      var sc = col(data.headers,'status'), xc = col(data.headers,'xp'), lc = col(data.headers,'lootbags');
      if (d.status !== undefined && sc >= 0) sh.getRange(r, sc+1).setValue(String(d.status));
      if (d.xp !== undefined && xc >= 0) sh.getRange(r, xc+1).setValue(Number(d.xp)||0);
      if (d.lootBags !== undefined && lc >= 0) sh.getRange(r, lc+1).setValue(Number(d.lootBags)||0);
      return { ok: true };
    }
  }
  return { ok: false, error: 'not_found' };
}

function getStudent(email) {
  getStudentSheet();
  var data = getSheetData('Students');
  email = String(email||'').toLowerCase().trim();
  for (var i = 0; i < data.rows.length; i++) {
    if (String(getCell(data.rows[i], data.headers, 'email')||'').toLowerCase().trim() === email)
      return { ok: true, student: buildStudent(data.rows[i], data.headers) };
  }
  return { ok: false, error: 'not_found' };
}

function getAllStudents() {
  getStudentSheet();
  var data = getSheetData('Students');
  var out = [];
  for (var i = 0; i < data.rows.length; i++) {
    var e = String(getCell(data.rows[i], data.headers, 'email')||'').trim();
    if (e && e.toLowerCase() !== 'email') out.push(buildStudent(data.rows[i], data.headers));
  }
  return { ok: true, students: out };
}

function buildStudent(row, headers) {
  return {
    email:     String(getCell(row, headers, 'email')||''),
    name:      String(getCell(row, headers, 'name')||''),
    section:   String(getCell(row, headers, 'section')||''),
    avatar:    String(getCell(row, headers, 'avatar')||''),
    status:    String(getCell(row, headers, 'status')||'pending'),
    xp:        Number(getCell(row, headers, 'xp'))||0,
    lootBags:  Number(getCell(row, headers, 'lootbags')||getCell(row, headers, 'lootBags'))||0,
    createdAt: String(getCell(row, headers, 'createdat')||'')
  };
}

// ══════════════════════════════
// QUIZZES
// ══════════════════════════════
function getQuizSheet() {
  var sh = SS.getSheetByName('Quizzes');
  if (!sh) {
    sh = SS.insertSheet('Quizzes');
    sh.appendRow(['id','title','subject','time','sections','locked','category','questions','createdAt']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function getQDataSheet() {
  var sh = SS.getSheetByName('QData');
  if (!sh) { sh = SS.insertSheet('QData'); sh.appendRow(['quizId','questions']); sh.setFrozenRows(1); }
  return sh;
}

function saveQuestions(quizId, questionsJson) {
  var sh = getQDataSheet();
  var data = getSheetData('QData');
  for (var i = 0; i < data.rows.length; i++) {
    if (String(getCell(data.rows[i], data.headers, 'quizid')) === String(quizId)) {
      sh.getRange(i+2, 2).setValue(questionsJson); return;
    }
  }
  sh.appendRow([String(quizId), questionsJson]);
}

function loadQuestions(quizId) {
  var data = getSheetData('QData');
  for (var i = 0; i < data.rows.length; i++) {
    if (String(getCell(data.rows[i], data.headers, 'quizid')) === String(quizId))
      return String(getCell(data.rows[i], data.headers, 'questions')||'[]');
  }
  return null;
}

function addQuiz(d) {
  var sh = getQuizSheet();
  var data = getSheetData('Quizzes');
  var id = String(d.id || Date.now());
  for (var i = 0; i < data.rows.length; i++) {
    if (String(getCell(data.rows[i], data.headers, 'id')) === id) return { ok: false, error: 'duplicate' };
  }
  var questionsJson = String(d.questions||'[]');
  saveQuestions(id, questionsJson);
  sh.appendRow([id, String(d.title||''), String(d.subject||''), Number(d.time)||15,
    String(d.sections||''), false, String(d.category||'practice'),
    'stored_in_QData', String(d.createdAt||new Date().toISOString())]);
  return { ok: true };
}

function addQuizFull(d) {
  // Like addQuiz but saves the full questions JSON (including embedded lessonData)
  var sh = getQuizSheet();
  var data = getSheetData('Quizzes');
  var id = String(d.id || Date.now());
  // Update if exists, insert if not
  for (var i = 0; i < data.rows.length; i++) {
    if (String(getCell(data.rows[i], data.headers, 'id')) === id) {
      // Already exists — just update questions in QData
      saveQuestions(id, String(d.questions||'[]'));
      return { ok: true, updated: true };
    }
  }
  // New quiz
  var questionsJson = String(d.questions||'[]');
  saveQuestions(id, questionsJson);
  sh.appendRow([id, String(d.title||''), String(d.subject||''), Number(d.time)||30,
    String(d.sections||''), false, String(d.category||'adventure'),
    'stored_in_QData', String(d.createdAt||new Date().toISOString())]);
  return { ok: true, inserted: true };
}

function updateQuiz(d) {
  var data = getSheetData('Quizzes');
  for (var i = 0; i < data.rows.length; i++) {
    if (String(getCell(data.rows[i], data.headers, 'id')) === String(d.id)) {
      var r = i+2, sh = data.sh;
      var lc = col(data.headers,'locked'), cc = col(data.headers,'category');
      if (d.locked !== undefined && lc >= 0) sh.getRange(r, lc+1).setValue(d.locked);
      if (d.category !== undefined && cc >= 0) sh.getRange(r, cc+1).setValue(d.category);
      return { ok: true };
    }
  }
  return { ok: false, error: 'not_found' };
}

function getQuizzes(section) {
  getQuizSheet();
  var data = getSheetData('Quizzes');
  var out = [];
  for (var i = 0; i < data.rows.length; i++) {
    var row = data.rows[i];
    if (!getCell(row, data.headers, 'id')) continue;
    var locked = getCell(row, data.headers, 'locked');
    if (locked === true || String(locked).toUpperCase() === 'TRUE') continue;
    var secs = String(getCell(row, data.headers, 'sections')||'').split(',').map(function(s){return s.trim();}).filter(Boolean);
    if (section && secs.length && secs.indexOf(section) < 0) continue;
    out.push(buildQuiz(row, data.headers, false)); // don't include lessonData in list
  }
  return { ok: true, quizzes: out };
}

function getQuizById(id) {
  getQuizSheet();
  var data = getSheetData('Quizzes');
  for (var i = 0; i < data.rows.length; i++) {
    if (String(getCell(data.rows[i], data.headers, 'id')) === String(id))
      return { ok: true, quiz: buildQuiz(data.rows[i], data.headers, true) }; // include lessonData
  }
  return { ok: false, error: 'not_found' };
}

function buildQuiz(row, headers, includeLessonData) {
  var id = String(getCell(row, headers, 'id')||'');
  var questionsJson = loadQuestions(id) || '[]';
  var secs = String(getCell(row, headers, 'sections')||'').split(',').map(function(s){return s.trim();}).filter(Boolean);
  var locked = getCell(row, headers, 'locked');
  var quiz = {
    id:        id,
    title:     String(getCell(row, headers, 'title')||''),
    subject:   String(getCell(row, headers, 'subject')||''),
    time:      Number(getCell(row, headers, 'time'))||15,
    sections:  secs,
    locked:    locked === true || String(locked).toUpperCase() === 'TRUE',
    category:  String(getCell(row, headers, 'category')||'practice'),
    questions: questionsJson,
    createdAt: String(getCell(row, headers, 'createdat')||'')
  };
  // Extract lessonData from questions field if it's an adventure lesson
  if (includeLessonData) {
    try {
      var qParsed = JSON.parse(questionsJson);
      if (Array.isArray(qParsed) && qParsed[0] && qParsed[0].question === '__ADVENTURE_LESSON__') {
        quiz.lessonData = qParsed[0].lessonData || '';
      }
    } catch(e) {}
    // Also try dedicated LessonData sheet as backup
    if (!quiz.lessonData) {
      var ld = getLesson(id);
      if (ld.ok) quiz.lessonData = ld.lessonData;
    }
  }
  return quiz;
}

// ══════════════════════════════
// SCORES
// ══════════════════════════════
function getScoreSheet() {
  var sh = SS.getSheetByName('Scores');
  if (!sh) {
    sh = SS.insertSheet('Scores');
    sh.appendRow(['email','name','section','quizId','quizTitle','category','score','total','percent','xp','tabSwitches','timeTaken','createdAt']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function addScore(d) {
  var sh = getScoreSheet();
  var score = Number(d.score)||0, total = Number(d.total)||1;
  var pct = Math.round((score/total)*100), xp = Math.round((score/total)*50);
  sh.appendRow([String(d.email||'').toLowerCase(), String(d.name||''), String(d.section||''),
    String(d.quizId||''), String(d.quizTitle||''), String(d.category||'practice'),
    score, total, pct, xp, Number(d.tabSwitches)||0, String(d.timeTaken||''),
    String(d.createdAt||new Date().toISOString())]);
  try {
    var s = getStudent(d.email);
    if (s.ok && s.student)
      updateStudent({ email:d.email, xp:(s.student.xp||0)+xp, lootBags:(s.student.lootBags||0)+(score===total?1:0) });
  } catch(e) {}
  return { ok: true };
}

function getAllScores() {
  getScoreSheet();
  var data = getSheetData('Scores');
  var out = [];
  for (var i = 0; i < data.rows.length; i++) {
    if (getCell(data.rows[i], data.headers, 'email')) out.push(buildScore(data.rows[i], data.headers));
  }
  return { ok: true, scores: out };
}

function getStudentScores(email) {
  getScoreSheet();
  var data = getSheetData('Scores');
  email = String(email||'').toLowerCase().trim();
  var out = [];
  for (var i = 0; i < data.rows.length; i++) {
    if (String(getCell(data.rows[i], data.headers, 'email')||'').toLowerCase() === email)
      out.push(buildScore(data.rows[i], data.headers));
  }
  return { ok: true, scores: out };
}

function buildScore(row, headers) {
  return {
    email:       String(getCell(row, headers, 'email')||''),
    name:        String(getCell(row, headers, 'name')||''),
    section:     String(getCell(row, headers, 'section')||''),
    quizId:      String(getCell(row, headers, 'quizid')||''),
    quizTitle:   String(getCell(row, headers, 'quiztitle')||''),
    category:    String(getCell(row, headers, 'category')||'practice'),
    score:       Number(getCell(row, headers, 'score'))||0,
    total:       Number(getCell(row, headers, 'total'))||0,
    xp:          Number(getCell(row, headers, 'xp'))||0,
    tabSwitches: Number(getCell(row, headers, 'tabswitches'))||0,
    createdAt:   String(getCell(row, headers, 'createdat')||'')
  };
}

// ══════════════════════════════
// FIX SHEETS
// ══════════════════════════════
function fixSheets() {
  var stuSh = SS.getSheetByName('Students');
  if (stuSh && stuSh.getLastRow() >= 1) {
    var first = String(stuSh.getRange(1,1).getValue()||'').toLowerCase().trim();
    if (first === 'id') stuSh.deleteColumn(1);
    var h1 = String(stuSh.getRange(1,1).getValue()||'').toLowerCase().trim();
    if (h1 !== 'email') {
      stuSh.insertRowBefore(1);
      stuSh.getRange(1,1,1,8).setValues([['email','name','section','avatar','status','xp','lootBags','createdAt']]);
      stuSh.setFrozenRows(1);
    }
  }
  getQuizSheet(); getScoreSheet(); getQDataSheet(); getLessonSheet();
  return { ok: true, message: 'All sheets ready including LessonData.' };
}
