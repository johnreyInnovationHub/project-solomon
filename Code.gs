// ════════════════════════════════════════════════
// PROJECT SOLOMON · Google Apps Script v6
// Execute as: Me  |  Access: Anyone
// ════════════════════════════════════════════════

var SS = SpreadsheetApp.getActiveSpreadsheet();

// ══ PASTE YOUR ANTHROPIC API KEY HERE ══
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
    else if (a === 'generateAdventure') result = generateAdventure(p);
    else if (a === 'saveLessonData')    result = saveLessonData(p);
    else if (a === 'getLessonData')     result = getLessonData(p.quizId);
    else result = { ok:true, status:'Solomon API v6 Online' };
  } catch(err) {
    result = { ok:false, error: err.message + ' | ' + err.stack };
  }
  return resp(result);
}

// ══════════════════════════════
// AI ADVENTURE GENERATOR
// Uses Claude API server-side (no CORS issues!)
// ══════════════════════════════
function generateAdventure(p) {
  var apiKey = ANTHROPIC_API_KEY;
  if (!apiKey || apiKey === 'YOUR_ANTHROPIC_API_KEY_HERE') {
    return { ok: false, error: 'Anthropic API key not configured in Code.gs' };
  }

  var pdfBase64 = p.pdfBase64 || '';
  var title = p.title || 'Lesson';
  var subject = p.subject || '';

  if (!pdfBase64) {
    return { ok: false, error: 'No PDF data provided' };
  }

  var prompt = 'You are an educational game designer. Read this lesson PDF and convert it into an interactive adventure lesson for a gamified learning app called Project Solomon (Adventure Time themed).\n\nGenerate a JSON object with this EXACT structure (no markdown, just raw JSON):\n{\n  "title": "lesson title",\n  "subject": "subject area",\n  "scenes": [\n    {\n      "type": "chapter",\n      "icon": "relevant emoji",\n      "chapterNum": 1,\n      "title": "chapter title",\n      "subtitle": "engaging subtitle"\n    },\n    {\n      "type": "dialogue",\n      "character": "finn|jake|bmo",\n      "text": "character says something engaging about the topic, using <em>highlighted terms</em> and <strong>important points</strong>"\n    },\n    {\n      "type": "content",\n      "title": "section title with emoji",\n      "body": "explanation text",\n      "formula": "optional formula or key equation (empty string if none)"\n    },\n    {\n      "type": "two_col",\n      "leftTitle": "left column title",\n      "leftColor": "exo|endo|blue|purple",\n      "leftItems": ["bullet 1","bullet 2","bullet 3"],\n      "leftFormula": "optional formula",\n      "rightTitle": "right column title",\n      "rightColor": "exo|endo|blue|purple",\n      "rightItems": ["bullet 1","bullet 2","bullet 3"],\n      "rightFormula": "optional formula"\n    },\n    {\n      "type": "highlight",\n      "style": "cyan|gold|green",\n      "icon": "emoji",\n      "label": "label text",\n      "body": "highlight content"\n    },\n    {\n      "type": "checkpoint",\n      "cpNum": 1,\n      "title": "Checkpoint Title",\n      "questions": [\n        {\n          "q": "question text",\n          "options": ["option A","option B","option C","option D"],\n          "correct": 0,\n          "feedback_right": "why correct answer is right",\n          "feedback_wrong": "explanation of correct answer"\n        }\n      ]\n    }\n  ]\n}\n\nRules:\n- Create 6-10 scenes total mixing chapters, dialogues, content, and at least 2 checkpoints\n- Each checkpoint should have 3-4 questions based on the actual lesson content\n- Characters: Finn (sword emoji, enthusiastic student), Jake (dog emoji, wise explainer), BMO (game emoji, gives tips/mnemonics)\n- Make dialogues engaging and educational, referencing actual content from the PDF\n- For two_col: use "exo" for red/warm styling, "endo" for blue/cool styling\n- Questions must be based on the ACTUAL content in this PDF\n- Keep all text educational but fun\n- Return ONLY the JSON object, no other text';

  var payload = {
    model: 'claude-opus-4-6',
    max_tokens: 8000,
    messages: [{
      role: 'user',
      content: [
        {
          type: 'document',
          source: {
            type: 'base64',
            media_type: 'application/pdf',
            data: pdfBase64
          }
        },
        {
          type: 'text',
          text: prompt
        }
      ]
    }]
  };

  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    var response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', options);
    var code = response.getResponseCode();
    var text = response.getContentText();
    var data = JSON.parse(text);

    if (code !== 200) {
      return { ok: false, error: 'Claude API error: ' + (data.error ? data.error.message : text.substring(0, 200)) };
    }

    // Extract the text content
    var content = data.content && data.content[0] && data.content[0].text ? data.content[0].text : '';
    
    // Parse JSON from response
    var lessonData;
    try {
      var cleaned = content.replace(/```json|```/g, '').trim();
      lessonData = JSON.parse(cleaned);
    } catch(e) {
      return { ok: false, error: 'Could not parse AI response: ' + e.message };
    }

    return { ok: true, lessonData: lessonData };

  } catch(err) {
    return { ok: false, error: 'Network error: ' + err.message };
  }
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
// ══════════════════════════════
function ensureStudentHeaders(sh) {
  var headers = ['email','name','section','avatar','status','xp','lootBags','createdAt'];
  var firstRow = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), headers.length)).getValues()[0];
  var firstCell = String(firstRow[0]||'').toLowerCase().trim();
  if (firstCell !== 'email') {
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
  for (var i = 0; i < data.rows.length; i++) {
    var rowEmail = String(getCell(data.rows[i], data.headers, 'email')||'').toLowerCase().trim();
    if (rowEmail === email) return { ok:false, error:'duplicate' };
  }
  sh.appendRow([email, String(d.name||''), String(d.section||''), String(d.avatar||''), 'pending', 0, 0, String(d.createdAt || new Date().toISOString())]);
  return { ok:true };
}

function updateStudent(d) {
  var sh = getStudentSheet();
  var data = getSheetData('Students');
  var email = String(d.email||'').toLowerCase().trim();
  for (var i = 0; i < data.rows.length; i++) {
    var rowEmail = String(getCell(data.rows[i], data.headers, 'email')||'').toLowerCase().trim();
    if (rowEmail === email) {
      var rowNum = i + 2;
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
  getStudentSheet();
  var data = getSheetData('Students');
  email = String(email||'').toLowerCase().trim();
  for (var i = 0; i < data.rows.length; i++) {
    var rowEmail = String(getCell(data.rows[i], data.headers, 'email')||'').toLowerCase().trim();
    if (rowEmail === email) return { ok:true, student: buildStudent(data.rows[i], data.headers) };
  }
  return { ok:false, error:'not_found' };
}

function getAllStudents() {
  getStudentSheet();
  var data = getSheetData('Students');
  var out = [];
  for (var i = 0; i < data.rows.length; i++) {
    var e = String(getCell(data.rows[i], data.headers, 'email')||'').trim();
    if (e && e.toLowerCase() !== 'email') out.push(buildStudent(data.rows[i], data.headers));
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
  if (sh.getLastRow() < 1) {
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
    if (String(getCell(data.rows[i], data.headers, 'id')) === id) return { ok:false, error:'duplicate' };
  }
  var questionsJson = String(d.questions||'[]');
  saveQuestions(id, questionsJson);
  // Also save lessonData if present (for AI-generated adventures)
  if (d.lessonData) {
    getLessonDataSheet();
    var ldData = getSheetData('LessonData');
    var ldSh = SS.getSheetByName('LessonData');
    ldSh.appendRow([id, String(d.lessonData)]);
  }
  sh.appendRow([id, String(d.title||''), String(d.subject||''), Number(d.time)||15, String(d.sections||''), false, String(d.category||'practice'), 'stored_in_QData', String(d.createdAt||new Date().toISOString())]);
  return { ok:true };
}

function getLessonDataSheet() {
  var sh = SS.getSheetByName('LessonData');
  if (!sh) { sh = SS.insertSheet('LessonData'); sh.appendRow(['quizId','lessonData']); sh.setFrozenRows(1); }
  return sh;
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
    var secs = String(getCell(row, data.headers, 'sections')||'').split(',').map(function(s){return s.trim();}).filter(Boolean);
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
  // Load lessonData if available (from LessonData sheet)
  var lessonData = null;
  try {
    var ldSh = SS.getSheetByName('LessonData');
    if (ldSh && ldSh.getLastRow() > 1) {
      var ldData = getSheetData('LessonData');
      for (var i = 0; i < ldData.rows.length; i++) {
        var ldId = String(getCell(ldData.rows[i], ldData.headers, 'quizid')||'').trim();
        if (ldId === id) {
          var rawLD = String(getCell(ldData.rows[i], ldData.headers, 'lessondata')||'');
          lessonData = rawLD.length > 10 ? rawLD : null;
          break;
        }
      }
    }
  } catch(e) {}
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
    lessonData: lessonData,
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
  sh.appendRow([String(d.email||'').toLowerCase(), String(d.name||''), String(d.section||''), String(d.quizId||''), String(d.quizTitle||''), String(d.category||'practice'), score, total, pct, xp, Number(d.tabSwitches)||0, String(d.timeTaken||''), String(d.createdAt||new Date().toISOString())]);
  try {
    var s = getStudent(d.email);
    if (s.ok && s.student) {
      updateStudent({ email:d.email, xp:(s.student.xp||0)+xp, lootBags:(s.student.lootBags||0)+(score===total?1:0) });
    }
  } catch(e) {}
  return { ok:true };
}

function getAllScores() {
  getScoreSheet();
  var data = getSheetData('Scores');
  var out = [];
  for (var i = 0; i < data.rows.length; i++) {
    if (getCell(data.rows[i], data.headers, 'email')) out.push(buildScore(data.rows[i], data.headers));
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
// GET LESSON DATA
// ══════════════════════════════
function getLessonData(quizId) {
  quizId = String(quizId || '').trim();
  if (!quizId) return { ok: false, error: 'quizId required' };
  try {
    var ldSh = SS.getSheetByName('LessonData');
    if (!ldSh) return { ok: false, error: 'LessonData sheet not found' };
    var ldData = getSheetData('LessonData');
    for (var i = 0; i < ldData.rows.length; i++) {
      var rowId = String(getCell(ldData.rows[i], ldData.headers, 'quizid') || '').trim();
      if (rowId === quizId) {
        var ld = String(getCell(ldData.rows[i], ldData.headers, 'lessondata') || '');
        if (ld.length > 10) return { ok: true, lessonData: ld };
      }
    }
    return { ok: false, error: 'not_found' };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ══════════════════════════════
// SAVE LESSON DATA
// Saves adventure lesson JSON to LessonData sheet
// Called via GET (metadata) then POST (lessonData)
// ══════════════════════════════
function saveLessonData(p) {
  var quizId = String(p.quizId || '');
  if (!quizId) return { ok: false, error: 'quizId required' };

  // If lessonData is provided (POST call), save it
  if (p.lessonData) {
    var ldSh = getLessonDataSheet();
    var ldData = getSheetData('LessonData');
    // Check if row already exists
    for (var i = 0; i < ldData.rows.length; i++) {
      if (String(getCell(ldData.rows[i], ldData.headers, 'quizid')) === quizId) {
        // Update existing
        ldSh.getRange(i + 2, 2).setValue(String(p.lessonData));
        return { ok: true, updated: true };
      }
    }
    // Insert new
    ldSh.appendRow([quizId, String(p.lessonData)]);
    return { ok: true, inserted: true };
  }

  // GET call — save quiz metadata to Quizzes sheet and confirm
  var qSh = getQuizSheet();
  var qData = getSheetData('Quizzes');
  // Check if quiz already exists
  for (var i = 0; i < qData.rows.length; i++) {
    if (String(getCell(qData.rows[i], qData.headers, 'id')) === quizId) {
      return { ok: true, exists: true };
    }
  }
  // Add to Quizzes sheet
  qSh.appendRow([
    quizId,
    String(p.title || ''),
    String(p.subject || ''),
    30,
    String(p.sections || ''),
    false,
    'adventure',
    'stored_in_LessonData',
    String(p.createdAt || new Date().toISOString())
  ]);
  return { ok: true, created: true };
}

// ══════════════════════════════
// FIX SHEETS
// ══════════════════════════════
function fixSheets() {
  var stuSh = SS.getSheetByName('Students');
  if (stuSh && stuSh.getLastRow() >= 1) {
    var firstCell = String(stuSh.getRange(1,1).getValue()||'').toLowerCase().trim();
    if (firstCell === 'id') stuSh.deleteColumn(1);
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
  getLessonDataSheet();
  return { ok:true, message:'All sheets fixed and ready.' };
}
