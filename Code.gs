const CFG = Object.freeze({
  TZ: 'Asia/Seoul',
  NOTIFY_EMAIL: 'caysbighand9@naver.com',
  SHEET: '상담신청',
  STATS: '월별통계',
  TIMES: ['08:00~08:30', '13:00~13:30', '16:00~16:30'],
  CATEGORIES: ['진학', '취업', '진로탐색', '학과/전공', '진로미결정', '기타'],
  STATUSES: ['대기', '승인', '보류', '완료', '반려'],
  HEADERS: ['접수시각','상담일','상담시간','상담일시키','학년','반','번호','학생이름','학번','상담분야','상담내용','상태','상담결과','완료일시','신청ID','학생전화번호']
});

function onOpen() {
  SpreadsheetApp.getUi().createMenu('진로상담 관리')
    .addItem('최초 설정', 'setup')
    .addItem('관리자 비밀번호 변경', 'changeAdminPin')
    .addItem('전체 기능 테스트', 'runAllTests')
    .addItem('월별 통계 새로 계산', 'rebuildMonthlyStats')
    .addToUi();
}

function setup() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone(CFG.TZ);
  let sh = ss.getSheetByName(CFG.SHEET);
  if (!sh) sh = ss.insertSheet(CFG.SHEET);
  sh.getRange(1,1,1,CFG.HEADERS.length).setValues([CFG.HEADERS]);
  sh.setFrozenRows(1);
  sh.getRange(1,1,1,CFG.HEADERS.length).setFontWeight('bold').setBackground('#173b6c').setFontColor('#fff');
  sh.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  sh.getRange('N:N').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  sh.autoResizeColumns(1, CFG.HEADERS.length);
  ensureStatsSheet_();
  MailApp.getRemainingDailyQuota();
  if (!PropertiesService.getScriptProperties().getProperty('ADMIN_PIN_HASH')) setPinFromPrompt_('관리자 비밀번호 설정');
  SpreadsheetApp.getUi().alert('설정 완료', '학생 화면과 교사 화면을 사용할 준비가 되었습니다. 전체 기능 테스트를 실행하세요.', SpreadsheetApp.getUi().ButtonSet.OK);
}

function changeAdminPin() { setPinFromPrompt_('관리자 비밀번호 변경'); }

function setPinFromPrompt_(title) {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt(title, '교사 화면에서 사용할 숫자 6자리를 입력하세요.', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const pin = r.getResponseText().trim();
  if (!/^\d{6}$/.test(pin)) return ui.alert('비밀번호는 숫자 6자리여야 합니다.');
  PropertiesService.getScriptProperties().setProperty('ADMIN_PIN_HASH', hash_(pin));
  ui.alert('관리자 비밀번호가 저장되었습니다.');
}

function doGet(e) {
  const admin = e && e.parameter && e.parameter.view === 'admin';
  return HtmlService.createHtmlOutputFromFile(admin ? 'Admin' : 'Student')
    .setTitle(admin ? '진로상담 관리' : '나의 진로상담 신청')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getStudentInit() {
  return { times: CFG.TIMES, categories: CFG.CATEGORIES, today: today_() };
}

function getAvailableTimes(dateText) {
  validateDate_(dateText);
  const booked = new Set(readRows_().filter(r => r.status !== '반려' && r.date === dateText).map(r => r.time));
  return CFG.TIMES.map(time => ({ time, available: !booked.has(time) }));
}

function submitCounseling(form) {
  const v = validateApplication_(form);
  if (!v.ok) return { ok:false, message:v.message };
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return { ok:false, message:'신청자가 많습니다. 잠시 후 다시 시도해 주세요.' }; }
  let id = '';
  let response;
  try {
    const rows = readRows_();
    if (rows.some(r => r.date === v.date && r.time === v.time && r.status !== '반려')) {
      return { ok:false, code:'SLOT_TAKEN', message:'먼저 신청한 사람이 있어 해당 시간은 선택할 수 없습니다. 다른 시간을 선택하세요.', times:getAvailableTimes(v.date) };
    }
    const now = new Date();
    id = Utilities.getUuid();
    sheet_().appendRow([now,v.date,v.time,v.date+'|'+v.time,v.grade,v.classNo,v.number,v.name,v.studentId,v.category,v.content,'대기','', '',id,v.phone]);
    SpreadsheetApp.flush();
    rebuildMonthlyStats_();
    response = { ok:true, message:`${v.name} 학생의 상담 신청이 완료되었습니다.`, id:id.slice(0,8).toUpperCase() };
  } finally { lock.releaseLock(); }
  notifyTeacher_(v, id);
  return response;
}

function notifyTeacher_(v, id) {
  try {
    const subject = `[진로상담 신청] ${v.name} 학생 · ${v.date} ${v.time}`;
    const body = [
      '새로운 진로상담 신청이 접수되었습니다.', '',
      `학생: ${v.name} (${v.studentId})`,
      `학생 전화번호: ${v.phone}`,
      `상담 분야: ${v.category}`,
      `희망 일시: ${v.date} ${v.time}`,
      `상담 내용: ${v.content}`,
      `접수번호: ${id.slice(0,8).toUpperCase()}`, '',
      '진로상담 교사용 관리 화면에서 신청 내용을 확인해 주세요.'
    ].join('\n');
    MailApp.sendEmail({ to:CFG.NOTIFY_EMAIL, subject, body, name:'진로상담 신청 알림' });
  } catch (e) {
    console.error('교사 알림 메일 발송 실패: ' + e.message);
  }
}

function getMyApplications(studentId, name) {
  const v = validateLookup_(studentId, name);
  if (!v.ok) return { ok:false, message:v.message, rows:[] };
  const rows = readRows_()
    .filter(r => r.studentId === v.studentId && r.name === v.name)
    .sort((a,b) => (b.date+' '+b.time).localeCompare(a.date+' '+a.time))
    .map(r => ({ date:r.date, time:r.time, category:r.category, status:r.status, id:r.id.slice(0,8).toUpperCase() }));
  return rows.length
    ? { ok:true, message:`총 ${rows.length}건의 상담 신청 내역이 있습니다.`, rows }
    : { ok:false, message:'입력한 학번과 이름에 해당하는 신청 내역이 없습니다.', rows:[] };
}

function validateLookup_(studentId, name) {
  studentId=String(studentId||'').trim(); name=String(name||'').trim();
  if (!/^\d{5}$/.test(studentId)) return bad_('학번 5자리를 입력하세요. 예: 1학년 1반 1번 → 10101');
  const grade=Number(studentId.slice(0,1)), classNo=Number(studentId.slice(1,3)), number=Number(studentId.slice(3,5));
  if (![1,2,3].includes(grade)||classNo<1||classNo>10||number<1||number>25) return bad_('학번을 확인하세요. 예: 1학년 1반 1번 → 10101');
  if (name.length<2||name.length>20) return bad_('신청할 때 입력한 이름을 정확히 입력하세요.');
  return {ok:true,studentId,name};
}

function validateApplication_(f) {
  f = f || {};
  const grade=Number(f.grade), classNo=Number(f.classNo), number=Number(f.number);
  const name=String(f.name||'').trim(), phoneDigits=String(f.phone||'').replace(/\D/g,''), category=String(f.category||'').trim();
  const date=String(f.date||'').trim(), time=String(f.time||'').trim(), content=String(f.content||'').trim();
  if (![1,2,3].includes(grade)) return bad_('학년을 선택하세요.');
  if (!Number.isInteger(classNo)||classNo<1||classNo>10) return bad_('반을 선택하세요.');
  if (!Number.isInteger(number)||number<1||number>25) return bad_('번호를 선택하세요.');
  if (name.length<2||name.length>20) return bad_('이름을 2~20자로 입력하세요.');
  if (!/^010\d{8}$/.test(phoneDigits)) return bad_('학생 전화번호를 010-1234-5678 형식으로 입력하세요.');
  if (!CFG.CATEGORIES.includes(category)) return bad_('상담 분야를 선택하세요.');
  try { validateDate_(date); } catch(e) { return bad_(e.message); }
  if (!CFG.TIMES.includes(time)) return bad_('상담 시간을 선택하세요.');
  if (content.length<2||content.length>1000) return bad_('상담받고 싶은 내용을 2~1000자로 입력하세요.');
  const phone=phoneDigits.replace(/(\d{3})(\d{4})(\d{4})/,'$1-$2-$3');
  return {ok:true,grade,classNo,number,name,phone,category,date,time,content,studentId:String(grade)+String(classNo).padStart(2,'0')+String(number).padStart(2,'0')};
}

function validateDate_(dateText) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dateText)) throw new Error('상담일을 선택하세요.');
  if (dateText < today_()) throw new Error('오늘보다 이전 날짜는 선택할 수 없습니다.');
}

function adminLogin(pin) {
  const saved = PropertiesService.getScriptProperties().getProperty('ADMIN_PIN_HASH');
  if (!saved) return {ok:false,message:'관리자 비밀번호가 설정되지 않았습니다. 스프레드시트에서 최초 설정을 실행하세요.'};
  if (hash_(String(pin||'')) !== saved) return {ok:false,message:'비밀번호가 올바르지 않습니다.'};
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('ADMIN_'+token,'1',21600);
  return {ok:true,token};
}

function getAdminData(token) {
  requireAdmin_(token);
  const rows=readRows_();
  const today=today_(), week=weekRange_();
  return {
    rows,
    stats: {
      today:rows.filter(r=>r.date===today && r.status!=='반려').length,
      week:rows.filter(r=>r.date>=week.start&&r.date<=week.end&&r.status!=='반려').length,
      pending:rows.filter(r=>r.status==='대기').length,
      completed:rows.filter(r=>r.status==='완료').length
    },
    categories:CFG.CATEGORIES,
    statuses:CFG.STATUSES,
    monthly:buildMonthlyStats_(rows)
  };
}

function updateCounseling(token, payload) {
  requireAdmin_(token);
  payload=payload||{};
  if (!CFG.STATUSES.includes(payload.status)) throw new Error('올바르지 않은 상태입니다.');
  const result=String(payload.result||'').trim();
  if (result.length>3000) throw new Error('상담 결과는 3000자 이내로 입력하세요.');
  const lock=LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const sh=sheet_(), values=sh.getDataRange().getValues();
    const idx=values.findIndex((r,i)=>i>0&&String(r[14])===String(payload.id));
    if (idx<1) throw new Error('해당 신청을 찾을 수 없습니다.');
    const currentStatus=String(values[idx][11]||'대기');
    if (payload.expectedStatus && currentStatus!==payload.expectedStatus) throw new Error('다른 화면에서 상태가 변경되었습니다. 새로고침 후 다시 시도하세요.');
    sh.getRange(idx+1,12,1,3).setValues([[payload.status,result,payload.status==='완료'?new Date():'']]);
    SpreadsheetApp.flush(); rebuildMonthlyStats_();
    return {ok:true,message:'상담 내용이 저장되었습니다.'};
  } finally { lock.releaseLock(); }
}

function readRows_() {
  const sh=sheet_(); if (sh.getLastRow()<2) return [];
  return sh.getRange(2,1,sh.getLastRow()-1,CFG.HEADERS.length).getValues().map((r,i)=>({
    row:i+2, created:fmtDateTime_(r[0]), date:fmtDate_(r[1]), time:String(r[2]||''),
    grade:Number(r[4]), classNo:Number(r[5]), number:Number(r[6]), name:String(r[7]||''), studentId:String(r[8]||''),
    category:String(r[9]||''), content:String(r[10]||''), status:String(r[11]||'대기'), result:String(r[12]||''), completed:fmtDateTime_(r[13]), id:String(r[14]||''), phone:String(r[15]||'')
  }));
}

function rebuildMonthlyStats() { rebuildMonthlyStats_(); SpreadsheetApp.getUi().alert('월별 통계를 새로 계산했습니다.'); }
function rebuildMonthlyStats_() {
  const sh=ensureStatsSheet_(), stats=buildMonthlyStats_(readRows_());
  sh.clearContents();
  const headers=['월','전체','진학','취업','진로탐색','학과/전공','진로미결정','기타','대기','승인','완료','반려'];
  sh.getRange(1,1,1,headers.length).setValues([headers]);
  if (stats.length) sh.getRange(2,1,stats.length,headers.length).setValues(stats.map(s=>headers.map(h=>s[h]||0)));
  sh.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#173b6c').setFontColor('#fff'); sh.setFrozenRows(1); sh.autoResizeColumns(1,headers.length);
}

function buildMonthlyStats_(rows) {
  const map={}; rows.forEach(r=>{ const m=r.date.slice(0,7); if(!m)return; if(!map[m])map[m]={'월':m,'전체':0,'진학':0,'취업':0,'진로탐색':0,'학과/전공':0,'진로미결정':0,'기타':0,'대기':0,'승인':0,'보류':0,'완료':0,'반려':0}; map[m]['전체']++; if(map[m][r.category]!==undefined)map[m][r.category]++; if(map[m][r.status]!==undefined)map[m][r.status]++; }); return Object.keys(map).sort().reverse().map(k=>map[k]);
}

function runAllTests() {
  const tests=[], a=(n,c)=>tests.push([n,!!c]);
  const v=validateApplication_({grade:1,classNo:10,number:25,name:'테스트학생',phone:'010-1234-5678',category:'진학',date:today_(),time:CFG.TIMES[0],content:'진학 상담을 받고 싶습니다.'});
  a('학번 생성',v.ok&&v.studentId==='11025');
  a('학생 전화번호 검증',v.ok&&v.phone==='010-1234-5678'&&!validateApplication_({grade:1,classNo:1,number:1,name:'학생',phone:'010-123',category:'진학',date:today_(),time:CFG.TIMES[0],content:'상담 내용'}).ok);
  a('학년 범위',!validateApplication_({grade:4,classNo:1,number:1,name:'학생',category:'진학',date:today_(),time:CFG.TIMES[0],content:'상담 내용'}).ok);
  a('반 범위',!validateApplication_({grade:1,classNo:11,number:1,name:'학생',category:'진학',date:today_(),time:CFG.TIMES[0],content:'상담 내용'}).ok);
  a('번호 범위',!validateApplication_({grade:1,classNo:1,number:26,name:'학생',category:'진학',date:today_(),time:CFG.TIMES[0],content:'상담 내용'}).ok);
  a('상담시간 3개',CFG.TIMES.length===3&&CFG.TIMES.includes('13:00~13:30'));
  a('상담분야 6개',CFG.CATEGORIES.length===6&&!CFG.CATEGORIES.includes('직업')&&!CFG.CATEGORIES.includes('개인상담'));
  const st=buildMonthlyStats_([{date:'2026-09-01',category:'진학',status:'완료'},{date:'2026-09-02',category:'취업',status:'대기'}]);
  a('월별 통계',st[0]['전체']===2&&st[0]['진학']===1&&st[0]['완료']===1);
  try { a('Student HTML 연결',HtmlService.createHtmlOutputFromFile('Student').getContent().includes('applicationForm')); } catch(e){a('Student HTML 연결',false);}
  try { a('Admin HTML 연결',HtmlService.createHtmlOutputFromFile('Admin').getContent().includes('adminTable')); } catch(e){a('Admin HTML 연결',false);}
  try { a('교사 승인 버튼 연결',HtmlService.createHtmlOutputFromFile('Admin').getContent().includes('approve-btn')); } catch(e){a('교사 승인 버튼 연결',false);}
  try { const adminHtml=HtmlService.createHtmlOutputFromFile('Admin').getContent(); a('신규 신청 PC 알림 연결',adminHtml.includes('notificationButton')&&adminHtml.includes('detectNewApplications')); } catch(e){a('신규 신청 PC 알림 연결',false);}
  try { a('교사 로그인 제목 연결',HtmlService.createHtmlOutputFromFile('Admin').getContent().includes('천안여자상업고등학교 진로상담실')); } catch(e){a('교사 로그인 제목 연결',false);}
  const lookup=validateLookup_('10101','테스트학생');
  a('신청내역 조회 입력검증',lookup.ok&&lookup.studentId==='10101');
  a('교사 알림 이메일 설정',/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(CFG.NOTIFY_EMAIL));
  const failed=tests.filter(t=>!t[1]); SpreadsheetApp.getUi().alert(failed.length?'테스트 실패':'전체 테스트 통과',tests.map(t=>(t[1]?'✅ ':'❌ ')+t[0]).join('\n')+`\n\n총 ${tests.length}개 중 ${tests.length-failed.length}개 통과`,SpreadsheetApp.getUi().ButtonSet.OK); return tests;
}

function sheet_(){const s=SpreadsheetApp.getActive().getSheetByName(CFG.SHEET);if(!s)throw new Error('먼저 최초 설정을 실행하세요.');return s;}
function ensureStatsSheet_(){const ss=SpreadsheetApp.getActive();return ss.getSheetByName(CFG.STATS)||ss.insertSheet(CFG.STATS);}
function bad_(m){return{ok:false,message:m};}
function today_(){return Utilities.formatDate(new Date(),CFG.TZ,'yyyy-MM-dd');}
function fmtDate_(v){if(!v)return'';if(Object.prototype.toString.call(v)==='[object Date]')return Utilities.formatDate(v,CFG.TZ,'yyyy-MM-dd');return String(v).slice(0,10);}
function fmtDateTime_(v){if(!v)return'';if(Object.prototype.toString.call(v)==='[object Date]')return Utilities.formatDate(v,CFG.TZ,'yyyy-MM-dd HH:mm:ss');return String(v);}
function weekRange_(){const d=new Date(),day=Number(Utilities.formatDate(d,CFG.TZ,'u'));const start=new Date(d);start.setDate(d.getDate()-day+1);const end=new Date(start);end.setDate(start.getDate()+6);return{start:Utilities.formatDate(start,CFG.TZ,'yyyy-MM-dd'),end:Utilities.formatDate(end,CFG.TZ,'yyyy-MM-dd')};}
function hash_(s){return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,s).map(b=>(b+256)%256).map(b=>b.toString(16).padStart(2,'0')).join('');}
function requireAdmin_(token){if(!token||CacheService.getScriptCache().get('ADMIN_'+token)!=='1')throw new Error('관리자 인증이 만료되었습니다. 다시 로그인하세요.');CacheService.getScriptCache().put('ADMIN_'+token,'1',21600);}
