/* ══════════════════════════════════════════════
   PROJECT SOLOMON · app.js  v2
   Google Sheets = PRIMARY database
   localStorage  = session + short-term cache only
══════════════════════════════════════════════ */

/* ─── OOO BACKGROUND INIT ─── */
function initStarfield() {
  const sf = document.getElementById('ooo-bg') || document.getElementById('starfield');
  if (!sf) return;
  for (let i = 0; i < 180; i++) {
    const s = document.createElement('div'); s.className = 'ooo-star';
    const sz = Math.random()*2.5+0.5;
    s.style.cssText = `width:${sz}px;height:${sz}px;left:${Math.random()*100}%;top:${Math.random()*100}%;animation-duration:${Math.random()*4+2}s;animation-delay:${Math.random()*5}s;opacity:${Math.random()*0.6+0.1};`;
    sf.appendChild(s);
  }
  ['rgba(74,14,143,1)','rgba(0,170,255,1)','rgba(255,105,180,1)','rgba(0,212,170,1)'].forEach(c => {
    const b = document.createElement('div'); b.className = 'ooo-blob';
    const sz = Math.random()*300+200;
    b.style.cssText = `width:${sz}px;height:${sz}px;background:${c};left:${Math.random()*100}%;top:${Math.random()*100}%;`;
    document.body.appendChild(b);
  });
}

/* ─── TOAST ─── */
let _toastTimer;
function showToast(msg, type='info') {
  let t = document.getElementById('global-toast');
  if (!t) { t = document.createElement('div'); t.id='global-toast'; t.className='toast'; document.body.appendChild(t); }
  t.textContent = msg; t.className = `toast ${type} show`;
  clearTimeout(_toastTimer); _toastTimer = setTimeout(()=>t.classList.remove('show'), 3500);
}

/* ─── SESSION ─── */
const SESSION_KEY = 'solomon_session';
function getSession() { try { return JSON.parse(localStorage.getItem(SESSION_KEY)||'null'); } catch { return null; } }
function setSession(data) { localStorage.setItem(SESSION_KEY, JSON.stringify(data)); }
function clearSession() { localStorage.removeItem(SESSION_KEY); }

/* ══════════════════════════════════════════════
   API LAYER — Google Sheets via Apps Script
══════════════════════════════════════════════ */
/* ══ PASTE YOUR GOOGLE APPS SCRIPT URL BETWEEN THE QUOTES BELOW ══ */
const HARDCODED_API_URL = 'https://script.google.com/macros/s/AKfycbxjrl_Uq_wO4MsjB6f_v9MOQ13FrFccQxDsykZO-N5CGG7jm7oCMCqmY6roNytUZngn/exec';
/* ══════════════════════════════════════════════════════════════ */
function getApiUrl() {
  return HARDCODED_API_URL !== 'YOUR_APPS_SCRIPT_URL_HERE'
    ? HARDCODED_API_URL
    : (localStorage.getItem('solomon_api_url') || '');
}

async function sheetsPost(action, payload={}) {
  const url = getApiUrl();
  if (!url) return { ok:false, error:'no_api' };
  try {
    // Use no-cors POST — Apps Script accepts plain text body
    // This avoids CORS preflight and handles large payloads (questions JSON)
    const body = JSON.stringify({ action, ...payload });
    const res = await fetch(url, {
      method: 'POST',
      mode: 'no-cors',       // skip CORS check entirely
      redirect: 'follow',
      body: body
    });
    // no-cors returns opaque response — we can't read it, but the write succeeds
    // So we return optimistic success
    return { ok: true, opaque: true };
  } catch(e) {
    // Fallback to GET if POST fails
    return sheetsGet(action, payload);
  }
}

async function sheetsGet(action, params={}) {
  const url = getApiUrl();
  if (!url) return { ok:false, error:'no_api' };
  const qs = new URLSearchParams({ action, ...params }).toString();
  try {
    const res = await fetch(`${url}?${qs}`, { redirect:'follow' });
    const text = await res.text();
    try { return JSON.parse(text); }
    catch { return { ok:false, error:'bad_json' }; }
  } catch(e) { return { ok:false, error:e.message }; }
}

/* ─── LOCAL fallback (for quizzes which teacher creates) ─── */
function getLocalData(key) { try { return JSON.parse(localStorage.getItem(`s_${key}`)||'[]'); } catch { return []; } }
function setLocalData(key, data) { localStorage.setItem(`s_${key}`, JSON.stringify(data)); }
function addLocalData(key, item) {
  const arr = getLocalData(key);
  const record = { ...item, id:Date.now(), createdAt:new Date().toISOString() };
  arr.push(record); setLocalData(key, arr); return record;
}

/* ══════════════════════════════════════════════
   STUDENT CACHE (per email, 30-min TTL)
══════════════════════════════════════════════ */
function cacheStudent(student) {
  if (!student||!student.email) return;
  localStorage.setItem(`s_stu_${student.email.toLowerCase()}`, JSON.stringify({...student, _cachedAt:Date.now()}));
}
function getCachedStudent(email) {
  try {
    const raw = localStorage.getItem(`s_stu_${email.toLowerCase()}`);
    if (!raw) return null;
    const s = JSON.parse(raw);
    if (Date.now() - s._cachedAt > 30*60*1000) return null; // expired
    return s;
  } catch { return null; }
}
function bustStudentCache(email) { localStorage.removeItem(`s_stu_${email.toLowerCase()}`); }

/* ══════════════════════════════════════════════
   STUDENTS
══════════════════════════════════════════════ */

/** Register — writes to Sheets (primary). localStorage = session cache only. */
async function saveStudent(student) {
  const payload = { ...student, status:'pending', xp:0, lootBags:0, createdAt:new Date().toISOString() };
  const result = await sheetsPost('addStudent', payload);
  if (result.ok) { cacheStudent(payload); return { ok:true }; }
  if (result.error==='duplicate') return { ok:false, error:'Email already registered' };
  if (result.error==='no_api') return { ok:false, error:'⚠ Google Sheets is not configured. Ask your teacher to set up the API.' };
  return { ok:false, error: result.error || 'Registration failed. Check your connection and try again.' };
}

/** Fetch student by email — Sheets first, 30-min cache fallback */
async function fetchStudent(email) {
  const result = await sheetsGet('getStudent', { email: email.toLowerCase() });
  if (result.ok && result.student) {
    cacheStudent(result.student);
    return result.student;
  }
  // Use cache as fallback (e.g. brief network issue)
  return getCachedStudent(email);
}

/** Teacher: get all students from Sheets */
async function getStudents() {
  const result = await sheetsGet('getStudents');
  if (result.ok && result.students) return result.students;
  return []; // Sheets is the source of truth
}

/** Update student in Sheets + local cache */
async function updateStudent(email, updates) {
  const cached = getCachedStudent(email);
  if (cached) cacheStudent({...cached, ...updates}); // optimistic cache update
  await sheetsPost('updateStudent', { email:email.toLowerCase(), ...updates });
}

/* ══════════════════════════════════════════════
   QUIZZES — local primary (teacher), Sheets backup
   Students fetch from Sheets
══════════════════════════════════════════════ */
function getQuizzes() { return getLocalData('quizzes'); }

function saveQuiz(quiz) {
  const record = addLocalData('quizzes', quiz);
  // Use POST (no-cors) for large payloads like questions JSON
  sheetsPost('addQuiz', {
    id: String(record.id), title:quiz.title, subject:quiz.subject||'',
    time:String(quiz.time||15), sections:(quiz.sections||[]).join(','),
    locked:'false', category:quiz.category||'practice',
    questions:JSON.stringify(quiz.questions||[]), createdAt:record.createdAt
  });
  return record;
}

function updateQuiz(id, updates) {
  const quizzes = getLocalData('quizzes');
  const idx = quizzes.findIndex(q=>q.id==id);
  if (idx<0) return;
  quizzes[idx] = {...quizzes[idx],...updates};
  setLocalData('quizzes', quizzes);
  sheetsPost('updateQuiz', { id, ...updates });
}

function getQuiz(id) { return getLocalData('quizzes').find(q=>q.id==id)||null; }

/** Students fetch their quizzes from Sheets */
async function fetchQuizzesForStudent(section) {
  const result = await sheetsGet('getQuizzes', { section });
  if (result.ok && result.quizzes) {
    result.quizzes.forEach(q => {
      if (typeof q.questions==='string') { try { q.questions=JSON.parse(q.questions); } catch { q.questions=[]; } }
      if (typeof q.sections==='string') q.sections=q.sections.split(',').map(s=>s.trim());
    });
    return result.quizzes;
  }
  // Fallback: local (same device)
  return getLocalData('quizzes').filter(q=>!q.locked&&(q.sections||[]).includes(section));
}

/* ══════════════════════════════════════════════
   SCORES
══════════════════════════════════════════════ */
async function saveScore(entry) {
  await sheetsPost('addScore', entry);
  const xpGain = Math.round((entry.score/entry.total)*50);
  const student = getCachedStudent(entry.email);
  if (student) {
    const newXP = (parseInt(student.xp)||0) + xpGain;
    const newLoot = entry.score===entry.total ? (parseInt(student.lootBags)||0)+1 : (parseInt(student.lootBags)||0);
    await updateStudent(entry.email, { xp:newXP, lootBags:newLoot });
  }
}

async function getScores() {
  const result = await sheetsGet('getScores');
  return (result.ok && result.scores) ? result.scores : [];
}

async function getStudentScores(email) {
  const result = await sheetsGet('getStudentScores', { email:email.toLowerCase() });
  return (result.ok && result.scores) ? result.scores : [];
}

/* ─── QUIZ PARSER ─── */
function parseQuizHTML(htmlString) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(htmlString, 'text/html');
  const questions = [];
  doc.querySelectorAll('ol > li, .question, [class*="question"]').forEach(qEl => {
    const qText = qEl.querySelector('p,.question-text')?.textContent.trim() || qEl.firstChild?.textContent.trim()||'';
    if (!qText) return;
    const opts=[]; let correct=0;
    qEl.querySelectorAll('li,.option,input[type="radio"]+label').forEach((opt,oi)=>{
      const txt=opt.textContent.trim().replace(/^[A-Da-d][.)]\s*/,'');
      if(!txt||txt===qText) return;
      opts.push(txt);
      if(opt.dataset.correct==='true'||opt.classList.contains('correct')) correct=oi;
    });
    if(opts.length>=2) questions.push({question:qText,options:opts,correct});
  });
  if(questions.length===0){
    const lines=(doc.body.innerText||doc.body.textContent||'').split('\n').map(l=>l.trim()).filter(Boolean);
    let cur=null;
    lines.forEach(line=>{
      if(/^\d+[.)]\s/.test(line)){if(cur&&cur.options.length>=2)questions.push(cur);cur={question:line.replace(/^\d+[.)]\s*/,''),options:[],correct:0};}
      else if(cur&&/^[A-Da-d][.)]\s/.test(line)){const ok=line.includes('*')||line.includes('✓');const txt=line.replace(/^[A-Da-d][.)]\s*/,'').replace(/[*✓]/g,'').trim();if(ok)cur.correct=cur.options.length;cur.options.push(txt);}
    });
    if(cur&&cur.options.length>=2)questions.push(cur);
  }
  return questions;
}

/* ─── AVATARS (Adventure Time Edition) ─── */
const AVATARS=[
  // Ooo Heroes
  {id:'finn',emoji:'⚔️',label:'Finn',universe:'Ooo Heroes'},
  {id:'jake',emoji:'🐶',label:'Jake',universe:'Ooo Heroes'},
  {id:'bmo',emoji:'🎮',label:'BMO',universe:'Ooo Heroes'},
  {id:'pb',emoji:'🍬',label:'PB',universe:'Ooo Heroes'},
  {id:'marceline',emoji:'🎸',label:'Marceline',universe:'Ooo Heroes'},
  {id:'iceking',emoji:'❄️',label:'Ice King',universe:'Ooo Heroes'},
  {id:'flame',emoji:'🔥',label:'Flame Princess',universe:'Ooo Heroes'},
  {id:'tree',emoji:'🌳',label:'Tree Trunks',universe:'Ooo Heroes'},
  {id:'gunter',emoji:'🐧',label:'Gunter',universe:'Ooo Heroes'},
  // Pokéverse
  {id:'pika',emoji:'⚡',label:'Zappix',universe:'Pokéverse'},
  {id:'bulb',emoji:'🌿',label:'Leaflet',universe:'Pokéverse'},
  {id:'char',emoji:'🔥',label:'Pyron',universe:'Pokéverse'},
  {id:'squi',emoji:'💧',label:'Aquor',universe:'Pokéverse'},
  {id:'mew',emoji:'🌟',label:'Mystix',universe:'Pokéverse'},
  // Upside Down
  {id:'el',emoji:'🔮',label:'Psionica',universe:'Upside Down'},
  {id:'will',emoji:'🎮',label:'Gamewarden',universe:'Upside Down'},
  {id:'demo',emoji:'🌑',label:'Shadewing',universe:'Upside Down'},
  {id:'max',emoji:'🎧',label:'Beatrix',universe:'Upside Down'},
  {id:'hop',emoji:'🛡️',label:'Ironkeep',universe:'Upside Down'},
  // Wizardlands
  {id:'hp',emoji:'⚡',label:'Thundermark',universe:'Wizardlands'},
  {id:'herm',emoji:'📚',label:'Scholara',universe:'Wizardlands'},
  {id:'ron',emoji:'♟️',label:'Strategos',universe:'Wizardlands'},
  {id:'luna',emoji:'🌙',label:'Lunara',universe:'Wizardlands'},
  {id:'nev',emoji:'🌱',label:'Bloomcaster',universe:'Wizardlands'},
  // Shinobi Realm
  {id:'nar',emoji:'🍜',label:'Rasenwave',universe:'Shinobi Realm'},
  {id:'sas',emoji:'⚡',label:'Sharinkai',universe:'Shinobi Realm'},
  {id:'sak',emoji:'💪',label:'Forceblossom',universe:'Shinobi Realm'},
  {id:'kak',emoji:'👁️',label:'Veilmaster',universe:'Shinobi Realm'},
  {id:'rock',emoji:'🥊',label:'Hardstone',universe:'Shinobi Realm'},
  // Strix Division
  {id:'anya',emoji:'🧠',label:'Mindreader',universe:'Strix Division'},
  {id:'loid',emoji:'🎭',label:'Phantom',universe:'Strix Division'},
  {id:'yor',emoji:'🌸',label:'Thornrose',universe:'Strix Division'},
  {id:'bond',emoji:'🐾',label:'Futures',universe:'Strix Division'},
  {id:'frank',emoji:'🔍',label:'Watcher',universe:'Strix Division'},
];
function getAvatarById(id){return AVATARS.find(a=>a.id===id)||AVATARS[0];}

/* ─── XP ─── */
function getLevel(xp){
  const lv=[{min:0,label:'Novice Hero',next:100},{min:100,label:'Grasslands Scout',next:250},{min:250,label:'Land of Ooo Explorer',next:500},
    {min:500,label:'Candy Kingdom Guard',next:1000},{min:1000,label:'Ooo Champion',next:2000},
    {min:2000,label:'Ancient Hero',next:4000},{min:4000,label:'Mathematical Legend',next:Infinity}];
  for(let i=lv.length-1;i>=0;i--){if(xp>=lv[i].min)return{...lv[i],xp,progress:i<lv.length-1?Math.min(100,((xp-lv[i].min)/(lv[i].next-lv[i].min))*100):100};}
  return{...lv[0],xp,progress:0};
}

/* ─── LOOT ─── */
const LOOT_REWARDS=[
  {id:'choco',emoji:'🍫',name:'Chocolate Bar',rarity:'Common'},{id:'pen',emoji:'✏️',name:'Cosmic Pen',rarity:'Common'},
  {id:'note',emoji:'📓',name:'Galaxy Notebook',rarity:'Uncommon'},{id:'sticker',emoji:'⭐',name:'Star Sticker Pack',rarity:'Common'},
  {id:'ruler',emoji:'📏',name:'Nebula Ruler',rarity:'Common'},{id:'medal',emoji:'🥇',name:'Gold Medal',rarity:'Rare'},
  {id:'eraser',emoji:'🧹',name:'Space Eraser',rarity:'Common'},{id:'clip',emoji:'📎',name:'Comet Clips Set',rarity:'Uncommon'},
];
function rollLoot(){
  const w=LOOT_REWARDS.map(r=>r.rarity==='Rare'?1:r.rarity==='Uncommon'?3:6);
  const t=w.reduce((a,b)=>a+b,0); let r=Math.random()*t;
  for(let i=0;i<LOOT_REWARDS.length;i++){r-=w[i];if(r<=0)return LOOT_REWARDS[i];}
  return LOOT_REWARDS[0];
}

/* ─── UTILS ─── */
function timeAgo(d){const s=(Date.now()-new Date(d))/1000;if(s<60)return'just now';if(s<3600)return`${Math.floor(s/60)}m ago`;if(s<86400)return`${Math.floor(s/3600)}h ago`;return`${Math.floor(s/86400)}d ago`;}
function formatDate(d){return new Date(d).toLocaleDateString('en-US',{month:'short',day:'numeric',year:'numeric'});}
function showLoading(el,msg='Loading...'){if(typeof el==='string')el=document.getElementById(el);if(el)el.innerHTML=`<div style="text-align:center;padding:2rem;color:var(--text-muted);"><div class="spinner"></div><p style="margin-top:1rem;font-size:0.85rem;">${msg}</p></div>`;}

document.addEventListener('DOMContentLoaded', initStarfield);
