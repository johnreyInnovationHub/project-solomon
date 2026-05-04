// ════════════════════════════════════════════════
// PROJECT SOLOMON · Google Apps Script v5
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
    if      (a === 'getStudent')        result = getStudent(p.email);
    else if (a === 'getStudents')       result = getAllStudents();
    else if (a === 'getQuizzes')        result = getQuizzes(p.section);
    else if (a === 'getQuizById')       result = getQuizById(p.id);
    else if (a === 'getScores')         result = getAllScores();
    else if (a === 'getStudentScores')  result = getStudentScores(p.email);
    else if (a === 'addStudent')        result = addStudent(p);
    else if (a === 'updateStudent')     result = updateStudent(p);
    else if (a === 'addQuiz')           result = addQuiz(p);
    else if (a === 'updateQuiz')        result = updateQuiz(p);
    else if (a === 'addScore')          result = addScore(p);
    else if (a === 'fixSheets')         result = fixSheets();
    else result = { ok:true, status:'Solomon API v5 Online' };
  } catch(err) {
    result = { ok:false, error: err.message + ' | ' + err.stack };
  }
  return resp(result);
}

// ══════════════════════════════
// SHEET HELPERS
// ══════════════════════════════
function getSheetData(sheetName) {
  var sh = SS.getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 1) return { sh:sh, headers:[], rows:[] };
  var all = sh.getDataRange().getValues();
  var headers = all[0].map(function(h){ return String(h).trim().toLowerCase(); });
  var rows = all.slice(1);
  return { sh:sh, headers:headers, rows:rows };
}

function col(headers, name) { return headers.indexOf(name.toLowerCase()); }

function getCell(row, headers, name) {
  var i = col(headers, name);
  return i >= 0 ? row[i] : '';
}

// ══════════════════════════════
// STUDENTS
// Sheet columns: email | name | section | avatar | status | xp | lootBags | createdAt
// (email is the primary key — NO separate id column needed)
// ══════════════════════════════
function ensureStudentHeaders(sh) {
  // Always make sure row 1 is the correct header row
  var headers = ['email','name','section','avatar','status','xp','lootBags','createdAt'];
  var firstRow = sh.getRange(1, 1, 1, sh.getLastColumn() || headers.length).getValues()[0];
  var firstCell = String(firstRow[0]||'').toLowerCase().trim();
  
  // If first cell is not 'email', fix it
  if (firstCell !== 'email') {
    // Check if it looks like an email address (data in row 1) — insert header row
    if (firstCell.indexOf('@') >= 0 || firstCell === 'id' || firstCell === '') {
      sh.insertRowBefore(1);
    }
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  }
}

function getStudentSheet() {
  var sh = SS.getSheetByName('Students');
  if (!sh) {
    sh = SS.insertSheet('Students');
    sh.appendRow(['email','name','section','avatar','status','xp','lootBags','createdAt']);
    sh.setFrozenRows(1);
    return sh;
  }
  ensureStudentHeaders(sh);
  return sh;
}

function addStudent(d) {
  var sh = getStudentSheet();
  var data = getSheetData('Students');
  var email = String(d.email||'').toLowerCase().trim();
  if (!email) return { ok:false, error:'email_required' };

  // Check for duplicate email
  for (var i = 0; i < data.rows.length; i++) {
    var rowEmail = String(getCell(data.rows[i], data.headers, 'email')||'').toLowerCase().trim();
    if (rowEmail === email) return { ok:false, error:'duplicate' };
  }

  sh.appendRow([
    email,
    String(d.name||''),
    String(d.section||''),
    String(d.avatar||''),
    'pending',
    0, 0,
    String(d.createdAt || new Date().toISOString())
  ]);
  return { ok:true };
}

function updateStudent(d) {
  var sh = getStudentSheet();
  var data = getSheetData('Students');
  var email = String(d.email||'').toLowerCase().trim();
  
  for (var i = 0; i < data.rows.length; i++) {
    var rowEmail = String(getCell(data.rows[i], data.headers, 'email')||'').toLowerCase().trim();
    if (rowEmail === email) {
      var rowNum = i + 2; // +2: row 1 is header, array is 0-indexed
      var sc = col(data.headers, 'status');
      var xc = col(data.headers, 'xp');
      var lc = col(data.headers, 'lootbags');
      
      if (d.status    !== undefined && sc >= 0) sh.getRange(rowNum, sc+1).setValue(String(d.status));
      if (d.xp        !== undefined && xc >= 0) sh.getRange(rowNum, xc+1).setValue(Number(d.xp)||0);
      if (d.lootBags  !== undefined && lc >= 0) sh.getRange(rowNum, lc+1).setValue(Number(d.lootBags)||0);
      return { ok:true };
    }
  }
  return { ok:false, error:'not_found' };
}

function getStudent(email) {
  getStudentSheet(); // ensure headers are correct
  var data = getSheetData('Students');
  email = String(email||'').toLowerCase().trim();
  
  for (var i = 0; i < data.rows.length; i++) {
    var rowEmail = String(getCell(data.rows[i], data.headers, 'email')||'').toLowerCase().trim();
    if (rowEmail === email) {
      return { ok:true, student: buildStudent(data.rows[i], data.headers) };
    }
  }
  return { ok:false, error:'not_found' };
}

function getAllStudents() {
  getStudentSheet();
  var data = getSheetData('Students');
  var out = [];
  for (var i = 0; i < data.rows.length; i++) {
    var e = String(getCell(data.rows[i], data.headers, 'email')||'').trim();
    if (e && e.toLowerCase() !== 'email') {
      out.push(buildStudent(data.rows[i], data.headers));
    }
  }
  return { ok:true, students:out };
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
    return sh;
  }
  // Ensure header row exists
  if (sh.getLastRow() < 1) {
    sh.appendRow(['id','title','subject','time','sections','locked','category','questions','createdAt']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function getQDataSheet() {
  var sh = SS.getSheetByName('QData');
  if (!sh) {
    sh = SS.insertSheet('QData');
    sh.appendRow(['quizId','questions']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function saveQuestions(quizId, questionsJson) {
  var sh = getQDataSheet();
  var data = getSheetData('QData');
  for (var i = 0; i < data.rows.length; i++) {
    if (String(getCell(data.rows[i], data.headers, 'quizid')) === String(quizId)) {
      sh.getRange(i+2, 2).setValue(questionsJson);
      return;
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
    if (String(getCell(data.rows[i], data.headers, 'id')) === id)
      return { ok:false, error:'duplicate' };
  }
  var questionsJson = String(d.questions||'[]');
  saveQuestions(id, questionsJson);
  sh.appendRow([
    id, String(d.title||''), String(d.subject||''),
    Number(d.time)||15, String(d.sections||''), false,
    String(d.category||'practice'), 'stored_in_QData',
    String(d.createdAt||new Date().toISOString())
  ]);
  return { ok:true };
}

function updateQuiz(d) {
  var data = getSheetData('Quizzes');
  for (var i = 0; i < data.rows.length; i++) {
    if (String(getCell(data.rows[i], data.headers, 'id')) === String(d.id)) {
      var r = i+2, sh = data.sh;
      var lc = col(data.headers,'locked'), cc = col(data.headers,'category');
      if (d.locked   !== undefined && lc >= 0) sh.getRange(r, lc+1).setValue(d.locked);
      if (d.category !== undefined && cc >= 0) sh.getRange(r, cc+1).setValue(d.category);
      return { ok:true };
    }
  }
  return { ok:false, error:'not_found' };
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
    var secs = String(getCell(row, data.headers, 'sections')||'')
      .split(',').map(function(s){return s.trim();}).filter(Boolean);
    if (section && secs.length && secs.indexOf(section) < 0) continue;
    out.push(buildQuiz(row, data.headers));
  }
  return { ok:true, quizzes:out };
}

function getQuizById(id) {
  getQuizSheet();
  var data = getSheetData('Quizzes');
  for (var i = 0; i < data.rows.length; i++) {
    if (String(getCell(data.rows[i], data.headers, 'id')) === String(id))
      return { ok:true, quiz: buildQuiz(data.rows[i], data.headers) };
  }
  return { ok:false, error:'not_found' };
}

function buildQuiz(row, headers) {
  var id = String(getCell(row, headers, 'id')||'');
  var questionsJson = loadQuestions(id);
  if (!questionsJson) {
    var inlineQ = String(getCell(row, headers, 'questions')||'[]');
    questionsJson = (inlineQ && inlineQ !== 'stored_in_QData') ? inlineQ : '[]';
  }
  var secs = String(getCell(row, headers, 'sections')||'')
    .split(',').map(function(s){return s.trim();}).filter(Boolean);
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
    return sh;
  }
  if (sh.getLastRow() < 1) {
    sh.appendRow(['email','name','section','quizId','quizTitle','category','score','total','percent','xp','tabSwitches','timeTaken','createdAt']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function addScore(d) {
  var sh = getScoreSheet();
  var score = Number(d.score)||0, total = Number(d.total)||1;
  var pct = Math.round((score/total)*100), xp = Math.round((score/total)*50);
  sh.appendRow([
    String(d.email||'').toLowerCase(), String(d.name||''), String(d.section||''),
    String(d.quizId||''), String(d.quizTitle||''), String(d.category||'practice'),
    score, total, pct, xp, Number(d.tabSwitches)||0, String(d.timeTaken||''),
    String(d.createdAt||new Date().toISOString())
  ]);
  try {
    var s = getStudent(d.email);
    if (s.ok && s.student) {
      updateStudent({ email:d.email, xp:(s.student.xp||0)+xp,
                      lootBags:(s.student.lootBags||0)+(score===total?1:0) });
    }
  } catch(e) {}
  return { ok:true };
}

function getAllScores() {
  getScoreSheet();
  var data = getSheetData('Scores');
  var out = [];
  for (var i = 0; i < data.rows.length; i++) {
    if (getCell(data.rows[i], data.headers, 'email'))
      out.push(buildScore(data.rows[i], data.headers));
  }
  return { ok:true, scores:out };
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
  return { ok:true, scores:out };
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
    percent:     Number(getCell(row, headers, 'percent'))||0,
    xp:          Number(getCell(row, headers, 'xp'))||0,
    tabSwitches: Number(getCell(row, headers, 'tabswitches'))||0,
    timeTaken:   String(getCell(row, headers, 'timetaken')||''),
    createdAt:   String(getCell(row, headers, 'createdat')||'')
  };
}

// ══════════════════════════════
// FIX SHEETS — repairs headers and column order
// ══════════════════════════════
function fixSheets() {
  // Fix Students sheet — ensure email is column A
  var stuSh = SS.getSheetByName('Students');
  if (stuSh && stuSh.getLastRow() >= 1) {
    var firstCell = String(stuSh.getRange(1,1).getValue()||'').toLowerCase().trim();
    if (firstCell === 'id') {
      // Old format had 'id' as first column — delete it
      stuSh.deleteColumn(1);
    }
    // Now ensure header row is correct
    var h1 = String(stuSh.getRange(1,1).getValue()||'').toLowerCase().trim();
    if (h1 !== 'email') {
      stuSh.insertRowBefore(1);
      stuSh.getRange(1,1,1,8).setValues([['email','name','section','avatar','status','xp','lootBags','createdAt']]);
      stuSh.setFrozenRows(1);
    }
  }
  getQuizSheet();
  getScoreSheet();
  getQDataSheet();
  return { ok:true, message:'Sheets fixed! Students sheet now uses email as primary key.' };
}
