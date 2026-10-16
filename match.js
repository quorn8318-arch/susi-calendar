/* 수시 달력 만들기 — 엑셀 읽기와 대학 일정 자동 매칭
   모든 처리는 브라우저 안에서만 이루어지고, 어떤 자료도 서버로 보내지 않습니다. */
(function(){
'use strict';

const WD = ['일','월','화','수','목','금','토'];
const pad = n => String(n).padStart(2,'0');
const KIND = {'1단계합격':'first','면접':'면접','논술':'논술','실기':'실기','최종합격':'final'};
const SKIP_REF = /원서접수|추천명단|서류마감|수능/;

/* ---------- 공통 ---------- */
function sheetRows(wb, name){
  const ws = wb.Sheets[name || wb.SheetNames[0]];
  return ws ? XLSX.utils.sheet_to_json(ws, {header:1, raw:false, defval:''}).map(r=>r.map(c=>String(c??'').normalize('NFKC').trim())) : [];
}
const clean = s => String(s||'').replace(/[\s()\[\]{}·•,./~\-:_'"「」【】<>]/g,'').replace(/전형|수시|모집/g,'');
function bigrams(s){ const out=new Map(); for(let i=0;i<s.length-1;i++){ const b=s.slice(i,i+2); out.set(b,(out.get(b)||0)+1); } return out; }
function dice(a,b){
  if(!a || !b) return 0; if(a===b) return 1; if(a.length<2 || b.length<2) return 0;
  const A=bigrams(a), B=bigrams(b); let inter=0;
  for(const [k,v] of A) if(B.has(k)) inter += Math.min(v,B.get(k));
  return 2*inter/(a.length-1 + b.length-1);
}
function isoOf(y,m,d){ return `${y}-${pad(m)}-${pad(d)}`; }
function wdOf(iso){ const [y,m,d]=iso.split('-').map(Number); return WD[new Date(y,m-1,d).getDay()]; }
// "2026-10-31", "2026.10.31.", "2026.10.31. 15:00", 엑셀 날짜 숫자 → {date,time}
function parseDate(s){
  s = String(s||'').trim(); if(!s) return null;
  let m = s.match(/(20\d\d)[-./년\s]+(\d{1,2})[-./월\s]+(\d{1,2})/);
  if(m){ const t = s.slice(m.index+m[0].length).match(/(\d{1,2}):(\d{2})/); return {date:isoOf(+m[1],+m[2],+m[3]), time:t?`${pad(+t[1])}:${t[2]}`:''}; }
  if(/^\d{5}(\.\d+)?$/.test(s)){ const d=new Date(Math.round((+s-25569)*864e5)); return {date:isoOf(d.getUTCFullYear(),d.getUTCMonth()+1,d.getUTCDate()), time:''}; }
  return null;
}

/* ---------- 대학 일정 목록 ---------- */
// 시트 '일정목록': 날짜 요일 구분 대학 전형 시각 비고 / 시트 '전형일정': 전문대 수시 일정표
function loadSchedule(wb){
  const names = wb.SheetNames;
  const evName = names.find(n=>/일정목록/.test(n)) || names[0];
  const jcName = names.find(n=>/전형일정/.test(n));
  const ev = []; const byU = {};
  const head = sheetRows(wb, evName);
  let hi = head.findIndex(r=>r.includes('날짜') && r.includes('대학'));
  if(hi<0) throw new Error('대학 일정 목록에서 「날짜·구분·대학·전형」 머리글을 찾지 못했습니다.');
  const H = head[hi], col = k => H.indexOf(k);
  const cD=col('날짜'), cK=col('구분'), cU=col('대학'), cT=col('전형'), cTm=col('시각'), cN=col('비고');
  for(const r of head.slice(hi+1)){
    const p = parseDate(r[cD]); if(!p || !r[cU]) continue;
    const e = {date:p.date, kind:r[cK]||'', u:r[cU], t:r[cT]||'', time:(r[cTm]||'').match(/\d{1,2}:\d{2}/)?.[0]||'', note:cN>=0?r[cN]||'':''};
    ev.push(e); (byU[e.u] ||= []).push(e);
  }
  const years = ev.map(e=>+e.date.slice(0,4)).filter(Boolean).sort();
  const Y = years[0] || new Date().getFullYear();
  const jc = {};
  if(jcName){
    const R = sheetRows(wb, jcName);
    for(const r of R){
      const u=r[3], mode=r[4]; if(!u || !/수시/.test(mode||'')) continue;
      (jc[u] ||= {})[/2차/.test(mode)?'2':'1'] = { iv:[r[6],r[7]], pr:[r[8],r[9]], ann:r[10], reg:r[11], add:r[12], apply:r[5] };
    }
  }
  return { ev, byU, jc, Y, keys4:new Set(Object.keys(byU)), keysJ:new Set(Object.keys(jc)) };
}
// 전문대 일정표의 "11/3", "10/16~17" → ISO
function md(s, Y){
  const m = String(s||'').match(/(\d{1,2})\s*\/\s*(\d{1,2})/); if(!m) return '';
  const mo=+m[1]; return isoOf(mo>=3?Y:Y+1, mo, +m[2]);
}

/* ---------- 대학 이름 맞추기 ---------- */
const U_ALIAS = {
  '한국과학기술원':'KAIST','포항공과대학교':'POSTECH','울산과학기술원':'UNIST','대구경북과학기술원':'DGIST','광주과학기술원':'GIST',
  '한국에너지공과대학교':'KENTECH','한국예술종합학교':'한예종','한국외국어대학교':'한국외대','서울과학기술대학교':'서울과기대',
  '한국기술교육대학교':'한국기술교대','한국체육대학교':'한국체대','동국대학교경주캠퍼스':'동국대(WISE)','동국대학교WISE캠퍼스':'동국대(WISE)',
  '고려대학교세종캠퍼스':'고려대(세종)','건국대학교글로컬캠퍼스':'건국대(글로컬)','연세대학교미래캠퍼스':'연세대(미래)',
  '한양대학교ERICA캠퍼스':'한양대(ERICA)','한양대학교(ERICA)':'한양대(ERICA)','상명대학교천안캠퍼스':'상명대(천안)',
  '을지대학교성남캠퍼스':'을지대','을지대학교의정부캠퍼스':'을지대','을지대학교대전캠퍼스':'을지대','국립목포대학교':'국립목포대(미확정)'
};
function uniVariants(raw){
  const s = String(raw||'').replace(/\s/g,'');
  const out = [];
  const add = x => { if(x && !out.includes(x)) out.push(x); };
  if(U_ALIAS[s]) add(U_ALIAS[s]);
  let base = s, campus = '';
  const cm = s.match(/^(.*?)(?:\(([^)]+)\)|([가-힣A-Za-z]+)캠퍼스)$/);
  if(cm && (cm[2] || cm[3])){ base = cm[1]; campus = (cm[2]||cm[3]).replace(/캠퍼스$/,''); }
  const shorts = b => {
    const r = [];
    r.push(b.replace(/학교$/,''));                   // 가천대학교 → 가천대
    r.push(b.replace(/대학교$/,'대').replace(/대학$/,'대'));
    r.push(b.replace(/여자대학교$/,'여대'));
    r.push(b.replace(/교육대학교$/,'교대'));
    r.push(b.replace(/공과대학교$/,'공대'));
    r.push(b.replace(/과학기술대학교$/,'과기대'));
    r.push(b.replace(/전문대학$/,'전문대'));
    r.push(b.replace(/공업전문대학$/,'공업전문대'));
    r.push(b);
    return r;
  };
  for(const b of [base, base.replace(/^국립/,'')]){
    for(const x of shorts(b)){
      if(campus){ add(`${x}(${campus})`); }
      add(x);
    }
  }
  return out;
}
function findUniv(raw, S){
  const vs = uniVariants(raw);
  for(const v of vs){ if(S.keys4.has(v)) return {keys:[v], jc:false}; if(S.keysJ.has(v)) return {keys:[v], jc:true}; }
  // 캠퍼스가 여러 개인 대학(예: 강원대 → 강원대(춘천삼척), 강원대(강릉원주))
  for(const v of vs){
    const multi = [...S.keys4].filter(k=>k.startsWith(v+'('));
    if(multi.length) return {keys:multi, jc:false};
    const mj = [...S.keysJ].filter(k=>k.startsWith(v+'('));
    if(mj.length) return {keys:mj, jc:true};
  }
  return {keys:[], jc:false};
}

/* ---------- 전형·모집단위 나누기 ---------- */
const MAJOR_END = /(학과|학부|과|전공|계열|대학|학군|자율전공|단|부)$/;
function splitUnit(unit){
  const parts = String(unit||'').split(/\s*[-()\[\]_]\s*|\s+/).map(x=>x.trim()).filter(Boolean)
    .filter(x=>!/^수시(모집)?$/.test(x) && !/^[A-Z]{0,3}\d{2,}[A-Z]?$/.test(x));
  const track=[], major=[];
  for(const p of parts){
    const q = p.replace(/^수시(?=[가-힣])/,'');
    if(/전형|전체|우수자|추천|인재|균형|일반|특별|기회|농어촌|논술|실기|면접|서류|교과|종합|고른|특기|대상자|출신|차$/.test(q) && !/(학과|학부)$/.test(q)) track.push(q);
    else if(MAJOR_END.test(q)) major.push(q);
    else track.push(q);
  }
  return { track: track.join(' '), major: major.join(' ') };
}
// 모집단위가 '사범대학 화학교육과'처럼 여러 단어면 마지막(가장 좁은) 단위를 쓴다
function majorStem(m){ const w = String(m||'').trim().split(/\s+/).pop()||''; return clean(w).replace(/^\(?주\)?|\(?야\)?$/g,'').replace(/(학과|학부|전공|계열|과)$/,''); }

/* ---------- 일정 점수 ---------- */
// 계열 추정(논술 인문/자연 구분 등에 사용)
const NAT = /공학|공과|기계|전기|전자|컴퓨터|소프트|SW|AI|인공지능|데이터|정보|반도체|간호|의예|의학|약학|치의|한의|수의|방사선|물리치료|작업치료|임상|보건|화학|생명|바이오|수학|물리|과학|건축|토목|환경|에너지|식품|농|산림|자동차|항공|모빌리티|시스템|융합공|자연/;
const HUM = /경영|경제|국어|국문|영어|영문|어문|문학|사회|행정|법|교육|심리|미디어|언론|관광|무역|회계|세무|복지|역사|사학|철학|정치|외교|글로벌|인문|상경|일본|중국|문화|신학|광고|커뮤니케이션|아동|유아/;
function fieldOf(major, given){
  if(/자연|공학|의약/.test(given||'')) return '자연'; if(/인문|사회/.test(given||'')) return '인문';
  if(/인문/.test(major)) return '인문'; if(/자연/.test(major)) return '자연';
  if(NAT.test(major)) return '자연'; if(HUM.test(major)) return '인문'; return '';
}
const ARTS = /태권도|체육|스포츠|무용|댄스|연기|공연|뮤지컬|실용음악|음악|성악|기악|작곡|국악|보컬|모델|미술|회화|조소|디자인|만화|애니|영상|연극|e스포츠|경호|실기/;
const CAT_RE = {논술:/논술/, 실기:/실기|실적|특기/, 교과:/교과/, 종합:/종합/};
function scoreEvent(e, trackC, majorS, cat, field){
  const T = e.t + (e.note && !/^\d/.test(e.note) ? ' '+e.note : '');
  const BARE = /^(학생부)?(교과|종합|논술|실기|실적)(위주)?(전체|모든|공통)?(\d차)?$/;
  // 칸 나누기: 쉼표·가운뎃점·빗금·괄호·띄어쓰기 단위. 종류 이름만 있는 칸('학생부종합')은 이름 맞추기에서 뺀다
  const parts = [...new Set(T.split(/[·,/()\[\]]|및|\s+/).map(clean).filter(p=>p && !BARE.test(p)))];
  const whole = clean(T);
  const excl = /제외|외\s*전체|외\s*모든|이외/.test(T);
  const exclText = excl ? (T.match(/\(([^)]*제외)\)|([^\s,]+)\s*제외|([^\s,]+)\s*외\s/)||[''])[0] : '';
  let s = 0;
  if(trackC){
    for(const p0 of parts){
      const p = p0.replace(/(?<=[가-힣])(I{1,3}|IV|V|[1-9])$/,'');   // '광운참빛인재I' → '광운참빛인재'
      if(p.length>=2 && (trackC.includes(p) || p.includes(trackC))) s = Math.max(s, Math.min(p.length,trackC.length)>=3 ? 1 : .6);
      s = Math.max(s, dice(trackC, p)*.95);
    }
    if(trackC.length>=3 && whole.includes(trackC)) s = Math.max(s,1);
  }
  const trackHit = s;
  if(/전체|모든\s*전형|전\s*전형|다수|공통/.test(T)) s = Math.max(s, .45);
  if(/수능최저|수능일/.test(e.t)) s = Math.max(s, .4);
  // 이 전형 종류를 이름으로 부르는 일정(예: '학생부교과', '논술전형')
  let adj = 0;
  for(const [k,re] of Object.entries(CAT_RE)){
    const has = re.test(T), mine = k===cat;
    if(!has) continue;
    const isExcluded = exclText && re.test(exclText);
    // 종류 이름만 있는 칸(예: '학생부교과', '학생부종합 전체', '논술전형')
    const bare = bareCat(T, k);
    if(mine && isExcluded) adj -= .8;
    else if(mine) s = Math.max(s, bare ? .7 : .45 + .1);
    else if(trackHit>=.6) continue;   // 우리 전형 이름이 분명히 들어 있으면 다른 종류가 섞여 있어도 깎지 않는다
    else if(!isExcluded && (k==='논술' || k==='실기') && !/교과|종합|전체|모든/.test(T)) adj -= .6;
    else if(!isExcluded && (k==='교과'||k==='종합') && (cat==='교과'||cat==='종합') && !CAT_RE[cat].test(T)) adj -= .3;  // 교과 일정인데 우리는 종합(또는 반대)
  }
  s += adj;
  // 예체능 일정(음악대학·체육특기 등)은 예체능 학과가 아니면 깎는다
  if(/음악|미술|체육|예술|예체능|연기|무용|디자인/.test(T) && !ARTS.test(majorS) && trackHit<.9) s -= .4;
  // '면접형' 일정인데 우리는 '서류형'·'추천형'처럼 유형이 다르면 깎는다
  const FORM = /(서류|면접|추천|활동|탐구|학업|교과)형/;
  const fT = (T.match(FORM)||[])[1], fK = (trackC.match(FORM)||[])[1];
  if(fT && fK && fT!==fK) s -= .6;
  const majorList = /(학과|학부|전공|의예|약학|간호|치의|한의|수의)/.test(e.t);
  if(majorS && majorS.length>=2 && whole.includes(majorS)) s += majorList ? .65 : .25;
  else if(majorList && majorS && !/전체|제외/.test(e.t)) s -= .35;  // 다른 학과 일정
  // 수능최저 적용 여부를 모를 때: 교과·논술은 적용(늦은 발표), 종합은 미적용(이른 발표) 쪽을 조금 더 높게
  if(/미적용|수능일\s*이전|수능\s*전/.test(T)) s += cat==='종합' ? .08 : 0;
  else if(/적용|수능일\s*이후|수능\s*후/.test(T)) s += (cat==='교과'||cat==='논술') ? .08 : 0;
  // 인문·자연 구분
  if(field){
    const other = field==='자연' ? '인문' : '자연';
    if(new RegExp(field).test(T)) s += .25; else if(new RegExp(other).test(T)) s -= .3;
  }
  return s;
}
// '학생부교과'처럼 종류 이름만 쓴 일정인가? '학생부교과(평생학습자)'처럼 괄호로 세부 전형을 단 것은 아니다
function bareCat(T, k){
  const word = k==='실기' ? '(?:실기|실적)' : k;
  const rx = new RegExp(`(?:학생부)?${word}(?:전형|위주)?(\\s*\\(([^)]*)\\))?`, 'g');
  let m; while((m = rx.exec(T))){
    if(!m[1] || /^\s*(전체|모든|공통|\d차|\d단계)/.test(m[2])) return true;
  }
  return false;
}
const SYN = [[/지균/g,'지역균형'],[/기균/g,'기회균형'],[/학추/g,'학교장추천'],[/교과일반/g,'일반']];
function catOf(track, major, given){
  const t = (given||'') + ' ' + track;
  if(/논술/.test(t)) return '논술';
  if(/실기|실적|특기/.test(t)) return '실기';
  if(/교과/.test(t)) return '교과';
  if(/종합/.test(t)) return '종합';
  // 이름으로 짐작(확실하지 않으므로 결과는 '미정'으로 표시됨)
  if(/인재|자기추천|서류형|면접형|활동우수|잠재|바람개비|학교생활우수|SSU|KU|DKU|창의|리더/.test(t)) return '종합';
  if(/학교장추천|고교추천|지역균형|교과우수|학생부우수|교과성적|일반학생|일반전형|일반고|학업우수|성적우수|^\s*일반\s*$/.test(t)) return '교과';
  return '';
}
// "10.31~11.1", "10.3~10.4 및 10.10~10.11", "12.5~12.7" → 범위 끝(첫 범위만)
function rangeEnd(e){
  const m = (e.note||'').match(/(\d{1,2})\.(\d{1,2})\s*~\s*(?:(\d{1,2})\.)?(\d{1,2})/);
  if(!m) return '';
  const [y,mo,d] = e.date.split('-').map(Number);
  if(+m[1]!==mo || +m[2]!==d) return '';
  let em = m[3]?+m[3]:mo, ed=+m[4]; const ey = em<mo ? y+1 : y;
  const end = isoOf(ey,em,ed); return end>e.date ? end : '';
}
const fmtMD = iso => { const [,m,d]=iso.split('-').map(Number); return `${m}.${d}(${wdOf(iso)})`; };
function slotFrom(e, unc){
  const s = {raw:fmtMD(e.date)+(e.time?' '+e.time:''), date:e.date};
  if(e.time) s.time = e.time;
  const end = rangeEnd(e);
  if(end){ s.end = end; s.unc = true; }
  if(unc) s.unc = true;
  return s;
}
const candOf = e => ({date:e.date, time:e.time||'', end:rangeEnd(e)||'', label:`${fmtMD(e.date)}${e.time?' '+e.time:''} ${e.t}${e.note?' — '+e.note:''}`});

function pick(events, kinds, trackC, majorS, cat, field, th, fallbackAll, after){
  const list = events.filter(e=>kinds.includes(e.kind) && (!after || e.date>=after)).map(e=>({e, s:scoreEvent(e,trackC,majorS,cat,field)})).sort((a,b)=>b.s-a.s || (a.e.date<b.e.date?-1:1));
  if(!list.length) return null;
  let good = list.filter(x=>x.s>=th);
  let fromFallback = false;
  if(!good.length){
    if(!fallbackAll) return null;
    good = list.filter(x=>x.s>-.1); fromFallback = true;
    if(!good.length) return null;
  }
  const best = good[0];
  const rivals = good.filter(x=>x.e.date!==best.e.date && x.s>=best.s-.15);
  const dates = new Set(good.map(x=>x.e.date));
  const unc = rivals.length>0 || (fromFallback && dates.size>1) || (best.s<.9 && dates.size>1);
  const cands = (unc ? good : []).slice(0,6).map(x=>candOf(x.e));
  return { e:best.e, unc, cands, score:best.s };
}

/* ---------- 학급 엑셀 읽기 ---------- */
// 지원하는 형식 1: NEIS 「반별제공현황진행상황」 (반·번호·성명·대학·평가단위명·수험번호)
// 지원하는 형식 2: 「수시 지원 상황 및 결과」 (학번·이름·대학·전형명·모집단위·전형일자 …)
function readClassFile(wb){
  for(const name of wb.SheetNames){
    const R = sheetRows(wb, name);
    for(let i=0;i<Math.min(R.length,12);i++){
      const r = R[i].map(x=>x.replace(/\s/g,''));
      if(r.includes('성명') && r.includes('대학') && r.includes('평가단위명')) return {fmt:'neis', H:r, rows:R.slice(i+1), sheet:name};
      if(r.includes('대학') && r.includes('모집단위') && (r.includes('이름')||r.includes('성명'))) return {fmt:'result', H:r, rows:R.slice(i+1), sheet:name};
    }
  }
  return null;
}

function buildApps(wb, S){
  const cf = readClassFile(wb);
  if(!cf) throw new Error('학생 지원 현황 머리글(성명·대학·평가단위명 또는 이름·대학·모집단위)을 찾지 못했습니다. NEIS에서 내려받은 「반별제공현황진행상황」 엑셀인지 확인해 주세요.');
  const H = cf.H; const c = (...ks) => { for(const k of ks){ const i=H.indexOf(k); if(i>=0) return i; } return -1; };
  const raw = [];
  if(cf.fmt==='neis'){
    const iG=c('계열/학년/학과','학년'), iC=c('반'), iN=c('번호'), iName=c('성명'), iU=c('대학'), iM=c('모집시기'), iUnit=c('평가단위명'), iE=c('수험번호'), iSt=c('진행상태');
    for(const r of cf.rows){
      if(!r[iName] || !r[iU]) continue;
      if(iM>=0 && r[iM] && !/수시/.test(r[iM])) continue;
      const g = (r[iG]||'').match(/\d+/)?.[0]||'', cl=(r[iC]||'').match(/\d+/)?.[0]||'', no=(r[iN]||'').match(/\d+/)?.[0]||'';
      const sp = splitUnit(r[iUnit]);
      raw.push({ grade:g, cls:cl, no:+no||99, sid: g&&cl&&no ? g+pad(cl)+pad(no) : '', name:r[iName], univ:r[iU], season:r[iM]||'수시',
        unit:r[iUnit]||'', track:sp.track, major:sp.major, eno:r[iE]||'', status:iSt>=0?r[iSt]:'' });
    }
  } else {
    const g = k => c(k);
    const iS=g('학번'), iName=c('이름','성명'), iR=g('지역'), iU=g('대학'), iSe=c('지원시기','모집시기'), iT=g('전형명'), iF=g('계열'), iMj=g('모집단위'),
      iK=g('전형종류'), iX=g('전형일자'), iP=g('1단계발표일'), iFi=g('최종발표일'), iCs=g('최저학력기준'), iQ=g('모집인원'), iCa=g('전형분류'), iMe=g('전형방법'), iE=g('수험번호');
    for(const r of cf.rows){
      if(!r[iName] || !r[iU]) continue;
      if(iSe>=0 && r[iSe] && !/수시/.test(r[iSe])) continue;
      const sid = (r[iS]||'').replace(/\D/g,'');
      const no = sid.length>=4 ? +sid.slice(-2) : 99;
      raw.push({ grade:sid.slice(0,1), cls:sid.length>=5?String(+sid.slice(1,3)):'', no, sid, name:r[iName], univ:r[iU], season:r[iSe]||'수시',
        unit:[r[iT],r[iMj]].filter(Boolean).join(' · '), track:r[iT]||'', major:r[iMj]||'', region:iR>=0?r[iR]:'', field:iF>=0?r[iF]:'',
        kindGiven:iK>=0?r[iK]:'', examGiven:iX>=0?r[iX]:'', firstGiven:iP>=0?r[iP]:'', finalGiven:iFi>=0?r[iFi]:'',
        csat:iCs>=0?r[iCs]:'', quota:iQ>=0?r[iQ]:'', cat:iCa>=0?r[iCa]:'', method:iMe>=0?r[iMe].replace(/\r/g,''):'', eno:iE>=0?r[iE]:'' });
    }
  }
  // 같은 학생·대학·수험번호가 두 번 있으면(재요청) 하나만 남긴다
  const seen = new Set(); const list = raw.filter(x=>{ const k=[x.name,x.univ,x.unit,x.eno].join('|'); if(seen.has(k)) return false; seen.add(k); return true; });
  list.sort((a,b)=>a.no-b.no || (a.sid<b.sid?-1:a.sid>b.sid?1:0));

  const apps = []; const refs = {}; const stats = {total:0, found:0, unc:0, noUniv:0, univMiss:new Set()};
  list.forEach((x, i)=>{
    const U = findUniv(x.univ, S);
    const a = { sid:x.sid, name:x.name, region:x.region||'', univ:x.univ, season:x.season, track:x.track, field:x.field||'', major:x.major, unit:cf.fmt==='neis'?x.unit:'',
      r1:'', fin:'', wait:'', reg:'', note:'', kind:'', csat:x.csat||'', quota:x.quota||'', cat:x.cat||'', method:x.method||'',
      id:'a'+String(i+1).padStart(3,'0'), eno:x.eno, no:x.no, sk:x.sid||x.name, jc:U.jc, exam:null, first:null, final:null, ref:'', uk:U.keys[0]||'', cands:{} };
    stats.total++;
    const given = x.fmt==='result' || cf.fmt==='result';
    const trackN = SYN.reduce((s,[re,to])=>s.replace(re,to), x.track);
    const trackC = clean(trackN), majorS = majorStem(x.major);
    const cat = catOf(trackN, x.major, x.cat);
    const field = fieldOf(x.major + ' ' + trackN, x.field);
    const refNotes = [];

    if(given){
      // 엑셀에 이미 날짜가 있으면 그대로 쓴다
      const ex = parseDate(x.examGiven), fi = parseDate(x.firstGiven), fn = parseDate(x.finalGiven);
      if(fn) a.final = {raw:x.finalGiven, date:fn.date, ...(fn.time?{time:fn.time}:{})};
      if(fi) a.first = {raw:x.firstGiven, date:fi.date, ...(fi.time?{time:fi.time}:{})};
      if(ex && x.kindGiven){ a.kind = x.kindGiven; a.exam = {raw:x.examGiven, date:ex.date};
        const r2 = x.examGiven.match(/~\s*(?:(\d{1,2})\.)?(\d{1,2})\.?/); if(r2){ const [y,m]=ex.date.split('-').map(Number); const em=r2[1]?+r2[1]:m; const end=isoOf(em<m?y+1:y,em,+r2[2]); if(end>ex.date){ a.exam.end=end; a.exam.unc=true; } } }
    }

    if(!U.keys.length){
      stats.noUniv++; stats.univMiss.add(x.univ);
      if(!a.final && !a.exam) a.ref = '일정 목록에 이 대학이 없어 날짜를 직접 넣어야 합니다.';
    } else if(U.jc){
      const J = S.jc[U.keys[0]]; const r = (/2차/.test(x.unit) && J['2']) ? J['2'] : (J['1']||J['2']);
      if(r){
        const ann = md(r.ann, S.Y);
        if(!a.final && ann) a.final = {raw:r.ann, date:ann};
        const iv = r.iv.filter(Boolean).join('~'), pr = r.pr.filter(Boolean).join('~');
        refs[U.keys[0]] = {jc:{iv, pr, ann:r.ann||'', reg:r.reg||'', add:r.add||''}};
        const setRange = (kind, rg, label) => {
          const s0 = md(rg[0]||rg[1], S.Y), s1 = md(rg[1], S.Y);
          if(!s0) return;
          a.kind = kind; a.exam = {raw:`${label} 기간 ${rg.filter(Boolean).join('~')}`, date:s0, unc:true}; if(s1 && s1>s0) a.exam.end = s1;
          refNotes.push(`학과 ${label} 날짜는 학교 일정표 기간 안에서 확인해 주세요`);
        };
        const JC_PR = /태권도|체육|스포츠지도|무용|댄스|연기|공연|뮤지컬|실용음악|보컬|모델|경호/;
        if(!a.exam && pr && /실기/.test(x.unit) || !a.exam && pr && JC_PR.test(x.major+x.track) && !/면접/.test(x.unit)) setRange('실기', r.pr, '실기');
        else if(!a.exam && iv && /면접/.test(x.unit)) setRange('면접', r.iv, '면접');
        // 전형일 후보: 학교 일정표의 면접·실기 기간
        const jcC = [];
        if(iv) jcC.push({date:md(r.iv[0]||r.iv[1],S.Y), end:md(r.iv[1],S.Y), time:'', kind:'면접', label:`면접 기간 ${iv}`});
        if(pr) jcC.push({date:md(r.pr[0]||r.pr[1],S.Y), end:md(r.pr[1],S.Y), time:'', kind:'실기', label:`실기 기간 ${pr}`});
        if(jcC.length) a.cands.exam = jcC.map(c=>({...c, end:c.end>c.date?c.end:''}));
      }
    } else {
      // 4년제: 캠퍼스가 여러 개면 일정이 더 잘 맞는 쪽을 고른다
      let bestKey = U.keys[0], bestSc = -9;
      if(U.keys.length>1){
        for(const k of U.keys){ const sc = Math.max(...S.byU[k].map(e=>scoreEvent(e,trackC,majorS,cat,field))); if(sc>bestSc){ bestSc=sc; bestKey=k; } }
        refNotes.push(`캠퍼스 확인 필요: ${U.keys.join(', ')} 중 ${bestKey} 일정으로 넣었습니다`);
      }
      a.uk = bestKey;
      const E = S.byU[bestKey];
      refs[bestKey] = {ev: E.filter(e=>!SKIP_REF.test(e.kind)).map(e=>[e.date, wdOf(e.date), e.kind, e.t, e.time, e.note])};
      if(!a.exam && !(/서류/.test(trackN) && cat!=='논술' && cat!=='실기')){   // 서류형은 면접이 없다
        const kinds = cat==='논술' ? ['논술'] : cat==='실기' ? ['실기','면접'] : ARTS.test(x.major+x.track) ? ['면접','실기'] : ['면접'];
        const p = pick(E, kinds, trackC, majorS, cat, field, .6, cat==='논술' || cat==='실기');
        if(p){ a.kind = p.e.kind; a.exam = slotFrom(p.e, p.unc); if(p.cands.length) a.cands.exam = p.cands; }
      }
      if(!a.first){
        const p = pick(E, ['1단계합격'], trackC, majorS, cat, field, .6, false);
        if(p && (a.exam || p.score>=.9) && (!a.exam || p.e.date<=a.exam.date)){ a.first = slotFrom(p.e, p.unc); delete a.first.end; if(p.cands.length) a.cands.first = p.cands; }
      }
      if(!a.final){
        // 최종 발표는 전형일(또는 1단계 발표) 뒤에 있다
        const after = a.exam ? (a.exam.end || a.exam.date) : a.first ? a.first.date : '';
        const p = pick(E, ['최종합격'], trackC, majorS, cat, field, .6, true, after) || (after && pick(E, ['최종합격'], trackC, majorS, cat, field, .6, true));
        if(p){ a.final = slotFrom(p.e, p.unc); delete a.final.end; if(p.cands.length) a.cands.final = p.cands; }
      }
    }
    if(a.cands.final) refNotes.push('최종 발표일이 수능최저 적용 여부·전형에 따라 다를 수 있어 후보 중에서 골라 주세요');
    if(a.cands.exam && !a.jc) refNotes.push('전형일 후보가 여러 개입니다');
    if(refNotes.length) a.ref = [a.ref, ...refNotes].filter(Boolean).join(' · ');
    if(!Object.keys(a.cands).length) delete a.cands;
    if(a.final || a.exam || a.first) stats.found++;
    if([a.exam,a.first,a.final].some(s=>s && s.unc)) stats.unc++;
    apps.push(a);
  });
  // 학급 정보
  const g0 = list.find(x=>x.grade)||{}, c0 = list.find(x=>x.cls)||{};
  const classes = [...new Set(list.map(x=>x.cls).filter(Boolean))];
  return { apps, refs, fmt:cf.fmt, grade:g0.grade||'', cls:classes.length===1?classes[0]:(c0.cls||''), classes, stats:{...stats, univMiss:[...stats.univMiss]} };
}

window.SusiMatch = { loadSchedule, buildApps, readClassFile, uniVariants, findUniv, splitUnit, scoreEvent, parseDate };
})();
