// ════════════════════════════════════════════════
// PROJECT SOLOMON · Google Apps Script v4
// Execute as: Me  |  Access: Anyone
// ════════════════════════════════════════════════

var SS = SpreadsheetApp.getActiveSpreadsheet();

function resp(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ══════════ BOTH GET AND POST handle all actions ══════════
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
    else result = { ok:true, status:'Solomon API v4 Online' };
  } catch(err) {
    result = { ok:false, error:err.message };
  }
  return resp(result);
}

// ══════════════════════════════
// SHEET HELPERS — reads by header name, not column index
// ══════════════════════════════
function getSheetData(sheetName) {
  var sh = SS.getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 1) return { sh:sh, headers:[], rows:[] };
  var all = sh.getDataRange().getValues();
  var headers = all[0].map(function(h){ return String(h).trim().toLowerCase(); });
  var rows = all.slice(1);
  return { sh:sh, headers:headers, rows:rows };
}

function col(headers, name) {
  var idx = headers.indexOf(name.toLowerCase());
  return idx; // -1 if not found
}

function getCell(row, headers, name) {
  var i = col(headers, name);
  return i >= 0 ? row[i] : '';
}

// ══════════════════════════════
// STUDENTS
// Expected headers: email, name, section, avatar, status, xp, lootbags, createdat
// ══════════════════════════════
function getStudentSheet() {
  var sh = SS.getSheetByName('Students');
  if (!sh) {
    sh = SS.insertSheet('Students');
    sh.appendRow(['email','name','section','avatar','status','xp','lootBags','createdAt']);
    sh.setFrozenRows(1);
    return sh;
  }
  var lastCol = sh.getLastColumn();
  var lastRow = sh.getLastRow();
  // Sheet is completely empty — just append headers
  if (lastCol < 1 || lastRow < 1) {
    sh.appendRow(['email','name','section','avatar','status','xp','lootBags','createdAt']);
    sh.setFrozenRows(1);
    return sh;
  }
  // Check if first row looks like data (starts with email address or timestamp)
  var firstCell = String(sh.getRange(1,1).getValue()||'');
  if (!firstCell || firstCell.indexOf('@') > 0 || firstCell.match(/^\d{13}/)) {
    sh.insertRowBefore(1);
    sh.getRange(1,1,1,8).setValues([['email','name','section','avatar','status','xp','lootBags','createdAt']]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function addStudent(d) {
  var sh = getStudentSheet();
  var data = getSheetData('Students');
  var email = String(d.email||'').toLowerCase().trim();
  if (!email) return { ok:false, error:'email_required' };
  for (var i=0;i<data.rows.length;i++) {
    if (String(getCell(data.rows[i],data.headers,'email')||'').toLowerCase()===email)
      return { ok:false, error:'duplicate' };
  }
  sh.appendRow([email, d.name||'', d.section||'', d.avatar||'', 'pending', 0, 0,
                d.createdAt||new Date().toISOString()]);
  return { ok:true };
}

function updateStudent(d) {
  var data = getSheetData('Students');
  var email = String(d.email||'').toLowerCase().trim();
  for (var i=0;i<data.rows.length;i++) {
    if (String(getCell(data.rows[i],data.headers,'email')||'').toLowerCase()===email) {
      var r = i+2; // +2 because row 1 is header, and array is 0-indexed
      var sh = data.sh;
      var sc = col(data.headers,'status'), xc = col(data.headers,'xp'), lc = col(data.headers,'lootbags');
      if (d.status   !== undefined && sc>=0) sh.getRange(r,sc+1).setValue(d.status);
      if (d.xp       !== undefined && xc>=0) sh.getRange(r,xc+1).setValue(Number(d.xp)||0);
      if (d.lootBags !== undefined && lc>=0) sh.getRange(r,lc+1).setValue(Number(d.lootBags)||0);
      return { ok:true };
    }
  }
  return { ok:false, error:'not_found' };
}

function getStudent(email) {
  var data = getSheetData('Students');
  email = String(email||'').toLowerCase().trim();
  for (var i=0;i<data.rows.length;i++) {
    if (String(getCell(data.rows[i],data.headers,'email')||'').toLowerCase()===email)
      return { ok:true, student:buildStudent(data.rows[i],data.headers) };
  }
  return { ok:false, error:'not_found' };
}

function getAllStudents() {
  var data = getSheetData('Students');
  var out = [];
  for (var i=0;i<data.rows.length;i++) {
    if (getCell(data.rows[i],data.headers,'email'))
      out.push(buildStudent(data.rows[i],data.headers));
  }
  return { ok:true, students:out };
}

function buildStudent(row, headers) {
  return {
    email:     String(getCell(row,headers,'email')||''),
    name:      String(getCell(row,headers,'name')||''),
    section:   String(getCell(row,headers,'section')||''),
    avatar:    String(getCell(row,headers,'avatar')||''),
    status:    String(getCell(row,headers,'status')||'pending'),
    xp:        Number(getCell(row,headers,'xp'))||0,
    lootBags:  Number(getCell(row,headers,'lootbags'))||0,
    createdAt: String(getCell(row,headers,'createdat')||'')
  };
}

// ══════════════════════════════
// QUIZZES
// Expected headers: id, title, subject, time, sections, locked, category, questions, createdAt
// ══════════════════════════════
function getQuizSheet() {
  var sh = SS.getSheetByName('Quizzes');
  if (!sh) {
    sh = SS.insertSheet('Quizzes');
    sh.appendRow(['id','title','subject','time','sections','locked','category','questions','createdAt']);
    sh.setFrozenRows(1);
    return sh;
  }
  var lastCol = sh.getLastColumn();
  var lastRow = sh.getLastRow();
  if (lastCol < 1 || lastRow < 1) {
    sh.appendRow(['id','title','subject','time','sections','locked','category','questions','createdAt']);
    sh.setFrozenRows(1);
    return sh;
  }
  var firstCell = String(sh.getRange(1,1).getValue()||'');
  if (!firstCell || firstCell.match(/^\d{13}/)) {
    sh.insertRowBefore(1);
    sh.getRange(1,1,1,9).setValues([['id','title','subject','time','sections','locked','category','questions','createdAt']]);
    sh.setFrozenRows(1);
  }
  return sh;
}

// ── QData sheet: stores questions JSON separately (no URL length limit) ──
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
  // Update if exists
  for (var i=0;i<data.rows.length;i++) {
    if (String(getCell(data.rows[i],data.headers,'quizid'))===String(quizId)) {
      sh.getRange(i+2, 2).setValue(questionsJson);
      return;
    }
  }
  sh.appendRow([String(quizId), questionsJson]);
}

function loadQuestions(quizId) {
  var data = getSheetData('QData');
  for (var i=0;i<data.rows.length;i++) {
    if (String(getCell(data.rows[i],data.headers,'quizid'))===String(quizId))
      return String(getCell(data.rows[i],data.headers,'questions')||'[]');
  }
  return null; // not found in QData
}

function addQuiz(d) {
  var sh = getQuizSheet();
  var data = getSheetData('Quizzes');
  var id = String(d.id||Date.now());
  // Check duplicate
  for (var i=0;i<data.rows.length;i++) {
    if (String(getCell(data.rows[i],data.headers,'id'))===id)
      return { ok:false, error:'duplicate' };
  }
  // Save questions to QData sheet (no size limit)
  var questionsJson = String(d.questions||'[]');
  saveQuestions(id, questionsJson);
  // Save quiz metadata to Quizzes sheet (questions column left as marker)
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
  for (var i=0;i<data.rows.length;i++) {
    if (String(getCell(data.rows[i],data.headers,'id'))===String(d.id)) {
      var r = i+2, sh = data.sh;
      var lc = col(data.headers,'locked'), cc = col(data.headers,'category');
      if (d.locked   !== undefined && lc>=0) sh.getRange(r,lc+1).setValue(d.locked);
      if (d.category !== undefined && cc>=0) sh.getRange(r,cc+1).setValue(d.category);
      return { ok:true };
    }
  }
  return { ok:false, error:'not_found' };
}

function getQuizzes(section) {
  getQuizSheet(); // ensure headers fixed
  var data = getSheetData('Quizzes');
  var out = [];
  for (var i=0;i<data.rows.length;i++) {
    var row = data.rows[i];
    if (!getCell(row,data.headers,'id')) continue;
    var locked = getCell(row,data.headers,'locked');
    if (locked===true||String(locked).toUpperCase()==='TRUE') continue;
    var secs = String(getCell(row,data.headers,'sections')||'').split(',').map(function(s){return s.trim();}).filter(Boolean);
    if (section && secs.length && secs.indexOf(section)<0) continue;
    out.push(buildQuiz(row,data.headers));
  }
  return { ok:true, quizzes:out };
}

function getQuizById(id) {
  getQuizSheet(); // ensure headers fixed
  var data = getSheetData('Quizzes');
  for (var i=0;i<data.rows.length;i++) {
    if (String(getCell(data.rows[i],data.headers,'id'))===String(id))
      return { ok:true, quiz:buildQuiz(data.rows[i],data.headers) };
  }
  return { ok:false, error:'not_found' };
}

function buildQuiz(row, headers) {
  var secs = String(getCell(row,headers,'sections')||'').split(',').map(function(s){return s.trim();}).filter(Boolean);
  var locked = getCell(row,headers,'locked');
  var id = String(getCell(row,headers,'id')||'');
  // Load questions from QData sheet (handles large JSONs)
  var questionsJson = loadQuestions(id);
  if (!questionsJson) {
    // Fallback: questions might be in the Quizzes sheet directly (old format)
    var inlineQ = String(getCell(row,headers,'questions')||'[]');
    questionsJson = (inlineQ && inlineQ !== 'stored_in_QData') ? inlineQ : '[]';
  }
  return {
    id:        id,
    title:     String(getCell(row,headers,'title')||''),
    subject:   String(getCell(row,headers,'subject')||''),
    time:      Number(getCell(row,headers,'time'))||15,
    sections:  secs,
    locked:    locked===true||String(locked).toUpperCase()==='TRUE',
    category:  String(getCell(row,headers,'category')||'practice'),
    questions: questionsJson,
    createdAt: String(getCell(row,headers,'createdat')||'')
  };
}

// ══════════════════════════════
// SCORES
// Expected headers: email, name, section, quizid, quiztitle, category, score, total, percent, xp, tabswitches, timetaken, createdat
// ══════════════════════════════
function getScoreSheet() {
  var sh = SS.getSheetByName('Scores');
  if (!sh) {
    sh = SS.insertSheet('Scores');
    sh.appendRow(['email','name','section','quizId','quizTitle','category','score','total','percent','xp','tabSwitches','timeTaken','createdAt']);
    sh.setFrozenRows(1);
    return sh;
  }
  var lastCol = sh.getLastColumn();
  var lastRow = sh.getLastRow();
  if (lastCol < 1 || lastRow < 1) {
    sh.appendRow(['email','name','section','quizId','quizTitle','category','score','total','percent','xp','tabSwitches','timeTaken','createdAt']);
    sh.setFrozenRows(1);
    return sh;
  }
  var firstCell = String(sh.getRange(1,1).getValue()||'');
  if (!firstCell || firstCell.indexOf('@') > 0) {
    sh.insertRowBefore(1);
    sh.getRange(1,1,1,13).setValues([['email','name','section','quizId','quizTitle','category','score','total','percent','xp','tabSwitches','timeTaken','createdAt']]);
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
  for (var i=0;i<data.rows.length;i++) {
    if (getCell(data.rows[i],data.headers,'email')) out.push(buildScore(data.rows[i],data.headers));
  }
  return { ok:true, scores:out };
}

function getStudentScores(email) {
  getScoreSheet();
  var data = getSheetData('Scores');
  email = String(email||'').toLowerCase().trim();
  var out = [];
  for (var i=0;i<data.rows.length;i++) {
    if (String(getCell(data.rows[i],data.headers,'email')||'').toLowerCase()===email)
      out.push(buildScore(data.rows[i],data.headers));
  }
  return { ok:true, scores:out };
}

function buildScore(row, headers) {
  return {
    email:       String(getCell(row,headers,'email')||''),
    name:        String(getCell(row,headers,'name')||''),
    section:     String(getCell(row,headers,'section')||''),
    quizId:      String(getCell(row,headers,'quizid')||''),
    quizTitle:   String(getCell(row,headers,'quiztitle')||''),
    category:    String(getCell(row,headers,'category')||'practice'),
    score:       Number(getCell(row,headers,'score'))||0,
    total:       Number(getCell(row,headers,'total'))||0,
    percent:     Number(getCell(row,headers,'percent'))||0,
    xp:          Number(getCell(row,headers,'xp'))||0,
    tabSwitches: Number(getCell(row,headers,'tabswitches'))||0,
    timeTaken:   String(getCell(row,headers,'timetaken')||''),
    createdAt:   String(getCell(row,headers,'createdat')||'')
  };
}

// ══════════════════════════════
// FIX SHEETS — call once to repair headers
// ══════════════════════════════
function fixSheets() {
  getStudentSheet();
  getQuizSheet();
  getScoreSheet();
  getQDataSheet();
  return { ok:true, message:'All sheet headers verified and fixed.' };
}
