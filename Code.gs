// ════════════════════════════════════════════════
// PROJECT SOLOMON · Google Apps Script v8
// Execute as: Me  |  Access: Anyone
// ════════════════════════════════════════════════

var SS = SpreadsheetApp.getActiveSpreadsheet();

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
    else result = { ok: true, status: 'Solomon API v8 Online' };
  } catch(err) {
    result = { ok: false, error: err.message };
  }
  return resp(result);
}

// ══════════════════════════════════════════════
// LESSON DATA — saved to Drive, ID stored in LessonData sheet
// ══════════════════════════════════════════════

function getLessonFolder() {
  var folders = DriveApp.getFoldersByName('SolomonLessons');
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder('SolomonLessons');
}

function saveLesson(p) {
  var quizId = String(p.quizId || '').trim();
  var lessonData = p.lessonData || '';
  if (!quizId) return { ok: false, error: 'quizId required' };
  if (!lessonData || lessonData.length < 5) return { ok: false, error: 'lessonData empty, size: ' + lessonData.length };

  try {
    var folder = getLessonFolder();
    var fileName = 'lesson_' + quizId + '.json';
    
    // Delete existing file if present
    var existing = folder.getFilesByName(fileName);
    while (existing.hasNext()) existing.next().setTrashed(true);
    
    // Create new file with lesson data
    var file = folder.createFile(fileName, lessonData, 'application/json');
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var fileId = file.getId();
    
    // Store the file ID in the LessonData sheet
    var sh = getLessonSheet();
    var vals = sh.getDataRange().getValues();
    var found = false;
    for (var i = 1; i < vals.length; i++) {
      if (String(vals[i][0]).trim() === quizId) {
        sh.getRange(i + 1, 2).setValue(fileId);
        sh.getRange(i + 1, 3).setValue(new Date().toISOString());
        found = true; break;
      }
    }
    if (!found) sh.appendRow([quizId, fileId, new Date().toISOString()]);
    
    return { ok: true, fileId: fileId, size: lessonData.length };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

function getLesson(quizId) {
  quizId = String(quizId || '').trim();
  if (!quizId) return { ok: false, error: 'quizId required' };
  
  try {
    var sh = getLessonSheet();
    var vals = sh.getDataRange().getValues();
    for (var i = 1; i < vals.length; i++) {
      if (String(vals[i][0]).trim() === quizId) {
        var fileId = String(vals[i][1] || '').trim();
        if (!fileId) return { ok: false, error: 'no_file_id' };
        
        // Read the lesson JSON from Drive
        var file = DriveApp.getFileById(fileId);
        var content = file.getBlob().getDataAsString();
        return { ok: true, lessonData: content };
      }
    }
    return { ok: false, error: 'not_found' };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

function getLessonSheet() {
  var sh = SS.getSheetByName('LessonData');
  if (!sh) {
    sh = SS.insertSheet('LessonData');
    sh.appendRow(['quizId', 'driveFileId', 'savedAt']);
    sh.setFrozenRows(1);
  }
  return sh;
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
  return '[]';
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
    out.push(buildQuiz(row, data.headers));
  }
  return { ok: true, quizzes: out };
}

function getQuizById(id) {
  getQuizSheet();
  var data = getSheetData('Quizzes');
  for (var i = 0; i < data.rows.length; i++) {
    if (String(getCell(data.rows[i], data.headers, 'id')) === String(id))
      return { ok: true, quiz: buildQuiz(data.rows[i], data.headers) };
  }
  return { ok: false, error: 'not_found' };
}

function buildQuiz(row, headers) {
  var id = String(getCell(row, headers, 'id')||'');
  var questionsJson = loadQuestions(id);
  var secs = String(getCell(row, headers, 'sections')||'').split(',').map(function(s){return s.trim();}).filter(Boolean);
  var locked = getCell(row, headers, 'locked');
  return {
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
  return { ok: true, message: 'All sheets ready.' };
}
