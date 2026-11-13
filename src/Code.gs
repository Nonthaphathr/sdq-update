/******************************************************
 * ระบบประเมิน SDQ และ EQ โรงเรียนกระดุมทองวิทยา
 * Code.gs  (Google Apps Script)
 * พัฒนาระบบโดย นางสาวกานต์มณี แก้วพวง (งานจัดระบบการบริหารและพัฒนาองค์กร)
 *           และ นายนนทพัทธ์ หิรัญเรือง (งานพัฒนาและใช้สื่อเทคโนโลยีเพื่อการศึกษา)
 * ฐานข้อมูล: Google Sheets (ชีตละ collection: id | json | updated)
 ******************************************************/
const APP_NAME = 'ระบบประเมิน SDQ และ EQ โรงเรียนกระดุมทองวิทยา';
const COLLS = ['Users','Students','Assessments','Cares','Settings'];
const SESSION_HOURS = 5;
const SALT = 'sdqeq|';
const LOGO_URL = 'https://nonthaphathr.github.io/uploadimg/logoschool.png';

/* ---------------- Web entry ---------------- */
function doGet(e){
  ensureSetup_();
  const q = (e && e.parameter) || {};
  const t = HtmlService.createTemplateFromFile('index');
  const boot = { settings: getSettings_(), url: ScriptApp.getService().getUrl(),
                 pub: String(q.p||'').replace(/[^a-z]/g,''),
                 m: String(q.m||'').replace(/[^a-z]/g,'').slice(0,2), u: String(q.u||'').toLowerCase().replace(/[^a-z0-9_.]/g,'').slice(0,40),
                 sid: String(q.s||'').replace(/[^A-Za-z0-9_-]/g,'').slice(0,40) };
  t.boot = JSON.stringify(boot).replace(/</g,'\\u003c');
  return t.evaluate()
    .setTitle(APP_NAME)
    .addMetaTag('viewport','width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ---------------- Sheet helpers ---------------- */
function ss_(){ return SpreadsheetApp.getActive(); }
function sh_(name){
  const s = ss_().getSheetByName(name);
  if (s) return s;
  const n = ss_().insertSheet(name);
  n.getRange(1,1,1,3).setValues([['id','json','updated']]).setFontWeight('bold');
  n.getRange('A:A').setNumberFormat('@');          /* รหัสนักเรียนขึ้นต้นด้วย 0 ไม่หาย */
  n.setFrozenRows(1);
  return n;
}
function safeParse_(s){ try{ return JSON.parse(s||'{}'); }catch(e){ return {}; } }
function readColl_(name){
  const v = sh_(name).getDataRange().getValues(), out = {};
  for (let i=1;i<v.length;i++){ if(!v[i][0]) continue; out[String(v[i][0])] = safeParse_(v[i][1]); }
  return out;
}
/* ops: [{op:'set'|'merge'|'delete', id, data}] — เขียนทีเดียวทั้งชีต */
function writeColl_(name, ops){
  const s = sh_(name), v = s.getDataRange().getValues(), idx = {};
  for (let i=1;i<v.length;i++) if (v[i][0]) idx[String(v[i][0])] = i;
  const now = new Date();
  ops.forEach(o=>{
    const id = String(o.id), i = idx[id];
    if (o.op==='delete'){ if(i!=null){ v[i] = ['','','']; delete idx[id]; } return; }
    const old = (i!=null && o.op==='merge') ? safeParse_(v[i][1]) : {};
    const json = JSON.stringify(Object.assign(old, o.data||{}));
    if (json.length > 49000) throw new Error('ข้อมูลใหญ่เกิน 50,000 ตัวอักษร/เซลล์');
    if (i!=null) v[i] = [id, json, now]; else { idx[id] = v.length; v.push([id, json, now]); }
  });
  const rows = v.slice(1).filter(r=>r[0]), last = s.getLastRow();
  if (last>1) s.getRange(2,1,last-1,3).clearContent();
  if (rows.length) s.getRange(2,1,rows.length,3).setValues(rows);
}
function pick_(obj, f){ const o = {}; Object.keys(obj).forEach(k=>{ if (f(obj[k],k)) o[k]=obj[k]; }); return o; }
function hash_(s){
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, SALT+String(s))
    .map(b=>('0'+(b&255).toString(16)).slice(-2)).join('');
}
function withLock_(fn){ const l = LockService.getScriptLock(); l.waitLock(30000); try{ return fn(); } finally { l.releaseLock(); } }

/* ---------------- Setup / settings ---------------- */
function defaultSettings_(){
  return {
    academic_year:'2569', term:'2',
    school_info:{ name:'โรงเรียนกระดุมทองวิทยา', affiliation:'องค์การบริหารส่วนจังหวัดศรีสะเกษ', logo:LOGO_URL,
                  director:'นางสุนิสา ใจใหญ่', director_pos:'ผู้อำนวยการโรงเรียนกระดุมทองวิทยา',
                  deputy:'นางสุคันทร บาอินทร์', deputy_pos:'รองผู้อำนวยการโรงเรียน',
                  counselor:'นางสาวกานต์มณี แก้วพวง', counselor_pos:'หัวหน้างานระบบดูแลช่วยเหลือนักเรียน' },
    credits:{ dev1:'นางสาวกานต์มณี แก้วพวง', dev1_role:'งานจัดระบบการบริหารและพัฒนาองค์กร',
              dev2:'นายนนทพัทธ์ หิรัญเรือง', dev2_role:'งานพัฒนาและใช้สื่อเทคโนโลยีเพื่อการศึกษา' },
    public_open:{ student:true, parent:true, eq:true },
    public_verify:false,      /* true = ลิงก์นักเรียน/ผู้ปกครองต้องกรอกวันเกิดยืนยันด้วย */
    memo:{ org:'โรงเรียนกระดุมทองวิทยา อำเภอกันทรารมย์ จังหวัดศรีสะเกษ', no:'ศก 51008.08/' },
    pwa_url:'',
    eq_norms:null            /* {good:{lo,hi}, smart:{lo,hi}, happy:{lo,hi}, total:{lo,hi}} — แอดมินตั้งเองในหน้าตั้งค่า */
  };
}
function getSettings_(){
  const s = readColl_('Settings').main || {}, d = defaultSettings_();
  return Object.assign(d, s, { school_info:Object.assign(d.school_info, s.school_info||{}),
    credits:Object.assign(d.credits, s.credits||{}), public_open:Object.assign(d.public_open, s.public_open||{}), memo:Object.assign(d.memo, s.memo||{}) });
}
function ensureSetup_(){
  const p = PropertiesService.getScriptProperties();
  if (p.getProperty('SETUP_DONE')){
    const us = readColl_('Users');
    if (us.admin && us.admin.name==='ผู้ดูแลระบบ') writeColl_('Users', [{op:'merge', id:'admin', data:{ name:'นางสาวกานต์มณี แก้วพวง', position:'งานจัดระบบการบริหารและพัฒนาองค์กร' }}]);
    if (Object.keys(us).some(k=>us[k].role==='Admin')) return;
    p.deleteProperty('SETUP_DONE');
  }
  COLLS.forEach(sh_);
  const users = readColl_('Users');
  if (!Object.keys(users).some(k=>users[k].role==='Admin')){
    writeColl_('Users', [{op:'set', id:'admin', data:{ username:'admin', pass:hash_('admin1234'), name:'นางสาวกานต์มณี แก้วพวง',
      position:'งานจัดระบบการบริหารและพัฒนาองค์กร', role:'Admin', status:'active', classes:[] }}]);
  }
  if (!readColl_('Settings').main) writeColl_('Settings', [{op:'set', id:'main', data:defaultSettings_()}]);
  p.setProperty('SETUP_DONE','1');
}
function round_(){ const s = getSettings_(); return s.academic_year+'-'+s.term; }

/* ---------------- Auth (session token) ---------------- */
function apiLogin(username, password, mode){
  ensureSetup_();
  const un = String(username||'').trim().toLowerCase();
  const cache = CacheService.getScriptCache(), fk = 'lf_'+un, fails = +(cache.get(fk)||0);
  if (fails >= 8) throw new Error('ลองผิดหลายครั้ง กรุณารอ 5 นาทีแล้วลองใหม่');
  const u = Object.values(readColl_('Users')).find(x=>String(x.username).toLowerCase()===un);
  if (!u || u.pass !== hash_(password)){ cache.put(fk, String(fails+1), 300); throw new Error('ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง'); }
  if (u.status !== 'active') throw new Error('บัญชีนี้ถูกระงับการใช้งาน');
  if (mode==='admin' && u.role!=='Admin') throw new Error('บัญชีนี้ไม่ใช่ผู้ดูแลระบบ — กรุณากลับไปเลือกโหมด “ครูประจำชั้น / ผู้บริหาร”');
  if (mode==='teacher' && u.role==='Admin') throw new Error('บัญชีผู้ดูแลระบบ — กรุณากลับไปเลือกโหมด “แอดมิน”');
  cache.remove(fk);
  const token = Utilities.getUuid();
  cache.put('S_'+token, u.username, SESSION_HOURS*3600);
  return JSON.stringify({ ok:true, token, user:publicUser_(u), data:loadFor_(u) });
}
function apiResume(token){ const u = me_(token); return JSON.stringify({ ok:true, user:publicUser_(u), data:loadFor_(u) }); }
function apiLogout(token){ CacheService.getScriptCache().remove('S_'+token); return '{"ok":true}'; }
function me_(token){
  const name = CacheService.getScriptCache().get('S_'+token);
  if (!name) throw new Error('AUTH_EXPIRED: เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่');
  const u = Object.values(readColl_('Users')).find(x=>x.username===name);
  if (!u || u.status!=='active') throw new Error('AUTH_EXPIRED: บัญชีไม่พร้อมใช้งาน');
  return u;
}
function publicUser_(u){ const c = Object.assign({}, u); delete c.pass; delete c.initpw; return c; }
function adminUser_(u){ const c = Object.assign({}, u); delete c.pass; return c; }   /* แอดมินเห็นรหัสเริ่มต้น (initpw) เพื่อคัดลอกส่งครู — ล้างทิ้งเมื่อเจ้าของบัญชีเปลี่ยนรหัสเอง */
function apiChangePassword(token, json){
  const u = me_(token), p = JSON.parse(json);
  if (u.pass !== hash_(p.old)) throw new Error('รหัสผ่านเดิมไม่ถูกต้อง');
  if (String(p.new||'').length < 6) throw new Error('รหัสผ่านใหม่ต้องมีอย่างน้อย 6 ตัวอักษร');
  withLock_(()=>writeColl_('Users', [{op:'merge', id:u.username, data:{pass:hash_(p.new), initpw:''}}]));
  return '{"ok":true}';
}

/* ---------------- RBAC: Admin = ทุกอย่าง | Viewer(ผู้บริหาร) = ดูทั้งโรงเรียน | Teacher = เฉพาะห้องที่ดูแล ---------------- */
function loadFor_(u){
  const users = readColl_('Users'), st = readColl_('Students'), as = readColl_('Assessments'), ca = readColl_('Cares');
  let S = st, A = as, C = ca, U = {};
  if (u.role==='Teacher'){
    const cl = u.classes||[];
    S = pick_(st, x=>cl.indexOf(x.cls)>=0);
    A = pick_(as, x=>S[x.sid]); C = pick_(ca, (x,k)=>S[k]);
    U[u.username] = publicUser_(u);
  } else if (u.role==='Viewer'){ U[u.username] = publicUser_(u); }
  else Object.keys(users).forEach(k=>{ U[k] = adminUser_(users[k]); });
  const advisors = Object.keys(users).map(k=>users[k]).filter(x=>x.role==='Teacher' && (x.status||'active')==='active')
    .map(x=>({ name:x.name, position:x.position||'', classes:x.classes||[] }));   /* ชื่อ-ห้อง ของครูที่ปรึกษา ใช้ลงนามในเอกสาร (ไม่มีรหัสผ่าน) */
  return { settings:getSettings_(), Users:U, Students:S, Assessments:A, Cares:C, advisors:advisors };
}
function authorize_(u, o, students){
  if (u.role==='Admin'){
    if (o.c==='Users' && o.op==='delete' && o.id===u.username) throw new Error('ลบบัญชีที่กำลังใช้งานอยู่ไม่ได้');
    return;
  }
  if (u.role!=='Teacher') throw new Error('ไม่มีสิทธิ์ทำรายการนี้');
  const cl = u.classes||[];
  if (o.c==='Students'){
    const old = students[o.id], cls = (o.data&&o.data.cls) || (old&&old.cls);
    if (!cls || cl.indexOf(cls)<0 || (old && cl.indexOf(old.cls)<0)) throw new Error('จัดการได้เฉพาะนักเรียนในห้องที่ท่านดูแล');
    return;
  }
  if (o.c==='Assessments' || o.c==='Cares'){
    const sid = o.c==='Cares' ? o.id : ((o.data&&o.data.sid) || String(o.id).split('_')[0]);
    const s = students[sid];
    if (!s || cl.indexOf(s.cls)<0) throw new Error('ไม่มีสิทธิ์ในนักเรียนคนนี้');
    return;
  }
  throw new Error('ไม่มีสิทธิ์ทำรายการนี้');
}
/* ops: [{c:'Students'|'Assessments'|..., op, id, data}] */
function apiCommit(token, opsJson){
  const u = me_(token), ops = JSON.parse(opsJson), students = readColl_('Students'), by = {};
  if (u.role==='Viewer') throw new Error('บัญชีผู้บริหารดูข้อมูลได้อย่างเดียว');
  ops.forEach(o=>{
    if (COLLS.indexOf(o.c)<0) throw new Error('ไม่รู้จักชุดข้อมูล '+o.c);
    authorize_(u, o, students);
    if (o.c==='Users' && o.data && o.data.newpass){
      if (String(o.data.newpass).length<6) throw new Error('รหัสผ่านต้องมีอย่างน้อย 6 ตัวอักษร');
      o.data.pass = hash_(o.data.newpass); if (o.id!==u.username) o.data.initpw = String(o.data.newpass); delete o.data.newpass;
    }
    if (o.c==='Users' && o.data) delete o.data.newpass;
    if (o.c==='Assessments' && o.data){ o.data.by = u.username; o.data.at = new Date().toISOString(); }
    (by[o.c] = by[o.c]||[]).push(o);
  });
  withLock_(()=>Object.keys(by).forEach(c=>writeColl_(c, by[c])));
  return '{"ok":true}';
}
/* นำเข้านักเรียน: rows = [{id,no,name,cls,sex,birth,parent}] */
function apiImportStudents(token, json){
  const u = me_(token), rows = JSON.parse(json), cl = u.classes||[];
  if (u.role!=='Admin' && u.role!=='Teacher') throw new Error('ไม่มีสิทธิ์นำเข้า');
  const ops = [], skipped = [];
  rows.forEach(r=>{
    const id = String(r.id||'').trim(), name = String(r.name||'').trim(), cls = String(r.cls||'').trim();
    if (!id || !name || !cls){ skipped.push(id||name||'(ว่าง)'); return; }
    if (u.role==='Teacher' && cl.indexOf(cls)<0){ skipped.push(id+' (ไม่ใช่ห้องที่ท่านดูแล)'); return; }
    ops.push({op:'merge', id, data:{ id, no:String(r.no||''), name, cls, sex:r.sex==='ญ'?'ญ':(r.sex==='ช'?'ช':''),
      birth:String(r.birth||''), parent:String(r.parent||'') }});
  });
  withLock_(()=>writeColl_('Students', ops));
  return JSON.stringify({ ok:true, added:ops.length, skipped });
}
/* นำเข้าบัญชีผู้ใช้ (แอดมินเท่านั้น): rows = [{username,name,position,role,classes,password}] */
function apiImportUsers(token, json){
  const u = me_(token), rows = JSON.parse(json);
  if (u.role!=='Admin') throw new Error('เฉพาะแอดมิน');
  const ops = [], skipped = [], ROLES = ['Admin','Teacher','Viewer'];
  rows.forEach(r=>{
    const un = String(r.username||'').trim().toLowerCase().replace(/[^a-z0-9_.]/g,'');
    if (!un || !r.name){ skipped.push(r.username||r.name||'(ว่าง)'); return; }
    const d = { username:un, name:String(r.name).trim(), position:String(r.position||''),
      role:ROLES.indexOf(r.role)>=0 ? r.role : 'Teacher', status:'active',
      classes:String(r.classes||'').split(/[,;，]/).map(x=>x.trim()).filter(String) };
    if (r.password){ d.pass = hash_(String(r.password)); d.initpw = String(r.password); }
    else if (!readColl_('Users')[un]){ d.pass = hash_(un); d.initpw = un; }
    ops.push({op:'merge', id:un, data:d});
  });
  withLock_(()=>writeColl_('Users', ops));
  return JSON.stringify({ ok:true, added:ops.length, skipped });
}


/* ออกบัญชีครูประจำชั้น (ต่อห้อง) / ผู้บริหาร (แอดมินเท่านั้น)
   json = {reset:bool, items:[{kind:'hr',cls:'ม.1'} | {kind:'exec',username:'exec1',name:'ผู้อำนวยการ',position:'...'}]}
   คืน [{username,name,role,classes,password,isNew}] — password = รหัสเริ่มต้น (ว่างถ้าเจ้าของเปลี่ยนรหัสเองแล้ว) */
function pwGen_(n){ const A='abcdefghjkmnpqrstuvwxyz23456789'; let o=''; for (let i=0;i<n;i++) o+=A.charAt(Math.floor(Math.random()*A.length)); return o; }
function clsUser_(cls){
  const m = String(cls).replace(/\s+/g,'').replace(/^ม\.?/,'m').replace(/[\/]/g,'_').replace(/[^a-z0-9_]/g,'');
  return 'hr_' + (m||('c'+Math.floor(Math.random()*1000)));
}
function apiIssueAccounts(token, json){
  const u = me_(token), p = JSON.parse(json), users = readColl_('Users');
  if (u.role!=='Admin') throw new Error('เฉพาะแอดมิน');
  const ops = [], out = [], used = {};
  (p.items||[]).forEach(it=>{
    let un, d;
    if (it.kind==='hr'){
      const cls = String(it.cls||'').trim(); if (!cls) return;
      un = clsUser_(cls);
      d = { username:un, name:'ครูประจำชั้น '+cls, position:'ครูที่ปรึกษา '+cls, role:'Teacher', status:'active', classes:[cls] };
    } else {
      un = String(it.username||'').toLowerCase().replace(/[^a-z0-9_.]/g,''); if (!un) return;
      d = { username:un, name:String(it.name||'ผู้บริหาร').trim(), position:String(it.position||'ผู้บริหาร').trim(), role:'Viewer', status:'active', classes:[] };
    }
    if (used[un]) return; used[un] = 1;
    const ex = users[un];
    if (ex && ex.role==='Admin') return;
    if (ex && !p.reset){ out.push({ username:un, name:ex.name, role:ex.role, classes:ex.classes||[], password:ex.initpw||'', isNew:false }); return; }
    const pw = pwGen_(6);
    if (ex){ delete d.name; delete d.position; }       /* ออกรหัสใหม่ทับ: คงชื่อ/ตำแหน่งที่แอดมินแก้ไว้ */
    d.pass = hash_(pw); d.initpw = pw;
    ops.push({ op:'merge', id:un, data:d });
    out.push({ username:un, name:(ex&&ex.name)||d.name, role:d.role, classes:d.classes, password:pw, isNew:!ex });
  });
  if (ops.length) withLock_(()=>writeColl_('Users', ops));
  return JSON.stringify({ ok:true, accounts:out });
}

/* ---------------- ลิงก์สาธารณะ: นักเรียน/ผู้ปกครองกรอกเอง (ต้องใช้รหัสนักเรียน + วันเกิดยืนยัน) ---------------- */
const PUB_FORMS = { s:{key:'student', n:25}, p:{key:'parent', n:25}, eq:{key:'eq', n:52} };
function pubCheck_(p){
  const f = PUB_FORMS[p.form];
  const st = getSettings_();
  if (!f || !st.public_open[f.key]) throw new Error('ลิงก์นี้ปิดรับการประเมินอยู่');
  const id = String(p.id||'').trim(), c = CacheService.getScriptCache(), k = 'pub_'+id, n = +(c.get(k)||0);
  if (n >= 12) throw new Error('ลองหลายครั้งเกินไป กรุณารอ 10 นาที');
  c.put(k, String(n+1), 600);
  const s = readColl_('Students')[id];
  if (!s) throw new Error('ไม่พบข้อมูลนักเรียน');
  if (st.public_verify && s.birth && String(p.birth||'').trim()!==s.birth) throw new Error('วันเดือนปีเกิดไม่ตรงกับในระบบ');
  return { s, f, id };
}
/* รายชื่อสำหรับหน้า "เลือกห้อง → เลือกชื่อ" (ไม่ส่งวันเกิด/ข้อมูลอื่นออกไป) */
function apiPublicRoster(json){
  const p = JSON.parse(json), f = PUB_FORMS[p.form], st = getSettings_();
  if (!f || !st.public_open[f.key]) throw new Error('ลิงก์นี้ปิดรับการประเมินอยู่');
  const S = readColl_('Students'), A = readColl_('Assessments'), rd = round_(), items = [];
  Object.keys(S).forEach(id=>{ const s = S[id]; if (!s || !s.name || !s.cls) return;
    const a = A[id+'_'+rd+'_'+p.form];
    items.push({ id:id, no:s.no||'', name:s.name, cls:s.cls, done:!!a, locked:!!(a && a.source!=='link') }); });
  return JSON.stringify({ ok:true, verify:!!st.public_verify, items:items });
}
function apiPublicLookup(json){
  const p = JSON.parse(json), r = pubCheck_(p), key = r.id+'_'+round_()+'_'+p.form;
  const ex = readColl_('Assessments')[key];
  return JSON.stringify({ ok:true, name:r.s.name, cls:r.s.cls, no:r.s.no, done:!!ex, locked:!!(ex && ex.source!=='link') });
}
function apiPublicSubmit(json){
  const p = JSON.parse(json), r = pubCheck_(p), key = r.id+'_'+round_()+'_'+p.form;
  const a = p.a;
  if (!Array.isArray(a) || a.length!==r.f.n || a.some(x=>x!==0&&x!==1&&x!==2&&x!==3)) throw new Error('กรอกข้อมูลไม่ครบหรือไม่ถูกต้อง');
  const ex = readColl_('Assessments')[key];
  if (ex && ex.source!=='link') throw new Error('ครูบันทึกผลการประเมินชุดนี้ไว้แล้ว');
  const imp = p.imp || null, d = { sid:r.id, round:round_(), form:p.form, a, source:'link',
    date:new Date().toISOString().slice(0,10), by:'link', at:new Date().toISOString() };
  if (p.form!=='eq'){
    d.imp = imp ? { q1:+imp.q1, dur:+imp.dur, dis:+imp.dis, ifr:(imp.ifr||[]).map(Number), burden:+imp.burden } : null;
    d.signer = { role:String((p.signer&&p.signer.role)||''), name:String((p.signer&&p.signer.name)||'').slice(0,80) };
  }
  withLock_(()=>writeColl_('Assessments', [{op:'set', id:key, data:d}]));
  return '{"ok":true}';
}

/* ---------------- ศูนย์อัปเดตระบบ ---------------- */
const APP_VERSION = '1.0.0';
const UPD_BASE = ['https:','','nonthaphathr.github.io','sdq-update'].join('/') + '/';
function updFetch_(p){
  const r = UrlFetchApp.fetch(UPD_BASE + p + '?t=' + Date.now(), { muteHttpExceptions:true });
  if (r.getResponseCode()!==200) throw new Error('เปิดศูนย์อัปเดตไม่ได้ (' + r.getResponseCode() + ')');
  return r.getContentText();
}
function updNewer_(a, b){
  const x = String(a||'0').split('.').map(Number), y = String(b||'0').split('.').map(Number);
  for (let i=0;i<Math.max(x.length,y.length);i++){ const p=x[i]||0, q=y[i]||0; if (p!==q) return p>q; }
  return false;
}
function updEditorUrl_(){ return ['https:','','script.google.com','home','projects',ScriptApp.getScriptId(),'edit'].join('/'); }
function apiUpdateInfo(token, clientVer){
  const u = me_(token); if (u.role!=='Admin') throw new Error('เฉพาะผู้ดูแลระบบ');
  let owner = ''; try{ owner = SpreadsheetApp.getActive().getOwner().getEmail(); }catch(e){}
  const out = { ok:true, current:APP_VERSION, client:String(clientVer||''), center:UPD_BASE, editor:updEditorUrl_(), owner:owner, latest:'', date:'', notes:[], newer:false, warn:'' };
  try{
    const v = JSON.parse(updFetch_('version.json'));
    out.latest = v.version||''; out.date = v.date||''; out.notes = v.notes||[]; out.files = v.files||[];
    out.newer = updNewer_(out.latest, APP_VERSION);
  }catch(e){ out.warn = String(e.message||e); }
  if (clientVer && clientVer!==APP_VERSION) out.warn = (out.warn?out.warn+' · ':'') + 'เวอร์ชัน Code.gs ('+APP_VERSION+') ไม่ตรงกับ index.html ('+clientVer+') — วางโค้ดให้ครบทั้ง 2 ไฟล์ แล้วสร้างเวอร์ชันใหม่';
  return JSON.stringify(out);
}
function showUpdateHelp(){
  let msg = 'เวอร์ชันที่ใช้อยู่: ' + APP_VERSION;
  try{ const v = JSON.parse(updFetch_('version.json')); msg += '\nเวอร์ชันล่าสุด: ' + v.version + (updNewer_(v.version, APP_VERSION)?'  ← มีรุ่นใหม่':'  (เป็นรุ่นล่าสุดแล้ว)'); }catch(e){ msg += '\n(ตรวจรุ่นล่าสุดไม่ได้: ' + e.message + ')'; }
  msg += '\n\nวิธีอัปเดต: เปิดเว็บแอป › เมนู “อัปเดตระบบ” (ผู้ดูแลระบบ) หรือเปิด ' + UPD_BASE;
  SpreadsheetApp.getUi().alert('อัปเดตระบบ SDQ / EQ', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

/* ---------------- เมนูในชีต (ติดตั้งแบบ Template) ---------------- */
function onOpen(){
  SpreadsheetApp.getUi().createMenu('🌿 ระบบ SDQ / EQ')
    .addItem('① ติดตั้งระบบ (ครั้งแรก)', 'setupWizard')
    .addItem('② วิธีเผยแพร่เป็นเว็บแอป', 'showDeployHelp')
    .addItem('③ เปิด / คัดลอกลิงก์เว็บแอป', 'openWebApp')
    .addSeparator()
    .addItem('ตรวจเวอร์ชัน / วิธีอัปเดต', 'showUpdateHelp')
    .addItem('รีเซ็ตรหัสผ่านแอดมิน', 'resetAdminPassword')
    .addToUi();
}
function setupWizard(){
  const ui = SpreadsheetApp.getUi();
  const has = Object.keys(readColl_('Students')).length + Object.keys(readColl_('Assessments')).length;
  if (has && ui.alert('ติดตั้งระบบใหม่?','ไฟล์นี้มีข้อมูลนักเรียน/ผลประเมินอยู่แล้ว การติดตั้งจะล้างข้อมูลทั้งหมดและเริ่มใหม่ ดำเนินการต่อหรือไม่', ui.ButtonSet.YES_NO)!==ui.Button.YES) return;
  const ask = function(t, m, def){ const r = ui.prompt(t, m+(def?'\n(เว้นว่าง = '+def+')':''), ui.ButtonSet.OK_CANCEL);
    if (r.getSelectedButton()!==ui.Button.OK) throw new Error('CANCEL'); return String(r.getResponseText()||'').trim() || def || ''; };
  try{
    const org  = ask('ขั้นที่ 1/4 · ชื่อโรงเรียน', 'ชื่อที่แสดงบนหัวระบบและเอกสารพิมพ์', 'โรงเรียนกระดุมทองวิทยา');
    const aff  = ask('ขั้นที่ 2/4 · สังกัด', 'เช่น องค์การบริหารส่วนจังหวัด…', 'องค์การบริหารส่วนจังหวัดศรีสะเกษ');
    const user = ask('ขั้นที่ 3/4 · ชื่อผู้ใช้แอดมิน', 'ภาษาอังกฤษ ไม่เว้นวรรค', 'admin').toLowerCase().replace(/[^a-z0-9_.]/g,'') || 'admin';
    let pass = ''; while (pass.length < 6){ pass = ask('ขั้นที่ 4/4 · รหัสผ่านแอดมิน', 'อย่างน้อย 6 ตัวอักษร (จดไว้ให้ดี)', ''); if (pass.length<6) ui.alert('รหัสผ่านสั้นเกินไป'); }
    withLock_(function(){
      COLLS.forEach(function(n){ const s = sh_(n), r = s.getLastRow(); if (r>1) s.getRange(2,1,r-1,3).clearContent(); });
      const st = defaultSettings_(); st.school_info.name = org; st.school_info.affiliation = aff;
      writeColl_('Settings', [{op:'set', id:'main', data:st}]);
      writeColl_('Users', [{op:'set', id:user, data:{ username:user, pass:hash_(pass), name:'นางสาวกานต์มณี แก้วพวง', position:'งานจัดระบบการบริหารและพัฒนาองค์กร', role:'Admin', status:'active', classes:[] }}]);
      PropertiesService.getScriptProperties().setProperty('SETUP_DONE','1');
    });
    ui.alert('ติดตั้งเรียบร้อย ✅', 'แอดมิน: '+user+'\n\nขั้นต่อไป: เมนู 🌿 ระบบ SDQ / EQ › ② วิธีเผยแพร่เป็นเว็บแอป', ui.ButtonSet.OK);
  }catch(e){ if (String(e.message)!=='CANCEL') throw e; }
}
function resetAdminPassword(){
  const ui = SpreadsheetApp.getUi(), us = readColl_('Users'), admins = Object.keys(us).filter(function(k){ return us[k].role==='Admin'; });
  if (!admins.length){ ui.alert('ไม่พบบัญชีแอดมิน — ใช้เมนู ① ติดตั้งระบบ'); return; }
  const r = ui.prompt('รีเซ็ตรหัสผ่าน', 'ชื่อผู้ใช้แอดมิน ('+admins.join(', ')+')', ui.ButtonSet.OK_CANCEL); if (r.getSelectedButton()!==ui.Button.OK) return;
  const id = String(r.getResponseText()).trim(); if (admins.indexOf(id)<0){ ui.alert('ไม่พบผู้ใช้นี้'); return; }
  const p = ui.prompt('รหัสผ่านใหม่', 'อย่างน้อย 6 ตัวอักษร', ui.ButtonSet.OK_CANCEL); if (p.getSelectedButton()!==ui.Button.OK) return;
  const pw = String(p.getResponseText()); if (pw.length<6){ ui.alert('รหัสผ่านสั้นเกินไป'); return; }
  writeColl_('Users', [{op:'merge', id:id, data:{pass:hash_(pw), status:'active'}}]); ui.alert('เปลี่ยนรหัสผ่านแล้ว');
}
function showDeployHelp(){
  const h = '<div style="font-family:Sarabun,Tahoma,sans-serif;font-size:14px;line-height:1.7;color:#12332E">'
   + '<p style="margin-top:0">ทำครั้งเดียว ประมาณ 1 นาที</p><ol style="padding-left:20px">'
   + '<li>เมนู <b>ส่วนขยาย › Apps Script</b></li>'
   + '<li>มุมขวาบน กด <b>การทำให้ใช้งานได้ › การทำให้ใช้งานได้รายการใหม่</b></li>'
   + '<li>กดรูปเฟือง ⚙ เลือก <b>เว็บแอป</b></li>'
   + '<li>เรียกใช้ในฐานะ: <b>ฉัน</b> · ผู้ที่มีสิทธิ์เข้าถึง: <b>ทุกคน</b></li>'
   + '<li>กด <b>ทำให้ใช้งานได้</b> แล้วอนุญาตสิทธิ์ (ถ้ามีคำเตือน กด ขั้นสูง › ไปที่โครงการ)</li>'
   + '<li>กลับมาที่ชีตนี้ เมนู <b>🌿 ระบบ SDQ / EQ › ③ เปิด / คัดลอกลิงก์เว็บแอป</b></li></ol>'
   + '<p style="background:#E8F8F4;padding:8px 10px;border-radius:8px">เมื่อแก้โค้ดภายหลัง ให้ <b>จัดการการทำให้ใช้งานได้ › ✏️ › เวอร์ชันใหม่</b> เพื่อคงลิงก์เดิม</p></div>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(h).setWidth(460).setHeight(400), 'เผยแพร่เป็นเว็บแอป');
}
function openWebApp(){
  let url = ''; try{ url = ScriptApp.getService().getUrl() || ''; }catch(e){}
  const ui = SpreadsheetApp.getUi();
  if (!url){ ui.alert('ยังไม่ได้เผยแพร่เป็นเว็บแอป','ทำตามเมนู ② ก่อน', ui.ButtonSet.OK); return; }
  const h = '<div style="font-family:Sarabun,Tahoma,sans-serif;font-size:14px"><p>ลิงก์เว็บแอปของท่าน</p>'
   + '<input id="u" value="'+url+'" style="width:100%;padding:8px;font-size:13px" readonly onclick="this.select()">'
   + '<p style="margin-top:14px"><a href="'+url+'" target="_blank" style="background:#12A58C;color:#fff;padding:9px 16px;border-radius:8px;text-decoration:none">เปิดระบบ</a> '
   + '<button onclick="var i=document.getElementById(\'u\');i.select();document.execCommand(\'copy\');this.textContent=\'คัดลอกแล้ว ✓\'" style="padding:8px 14px;border-radius:8px;border:1px solid #ccc;background:#fff">คัดลอกลิงก์</button></p>'
   + '<p style="color:#64748b;font-size:12px">เข้าระบบด้วยบัญชีแอดมินที่ตั้งไว้ แล้วนำเข้านักเรียน/บัญชีครูที่เมนู นักเรียน และ ผู้ใช้งาน</p></div>';
  ui.showModalDialog(HtmlService.createHtmlOutput(h).setWidth(460).setHeight(220), 'เว็บแอประบบ SDQ / EQ');
}
/* กดรันครั้งแรกใน Apps Script เพื่อขอสิทธิ์ */
function authorizeOnce(){
  ensureSetup_();
  Logger.log('พร้อมใช้งาน — ผู้ใช้เริ่มต้น admin / admin1234 (เปลี่ยนรหัสทันที)');
}
