// 채용공고 수집기 — Google Sheets(DB) + Apps Script(수집·웹앱)
const CATEGORIES = [
  // '분석'만으로는 성분분석·회계분석 등까지 걸려서, 데이터 단서를 요구하고 실험·회계·설비 쪽은 제외 (단, 제목에 '데이터'가 있으면 유지)
  { name: '데이터 분석', kw: '데이터 분석', kws: ['데이터 분석', '데이터 애널리스트', 'BI 분석'], re: /데이터|data|analy|애널리스트|\bBI\b|통계|빅데이터|SQL|태블로|tableau|그로스|분석가/i,
    ex: /성분|화학|시험|실험|임상|기기분석|재료|광물|토양|수질|대기|환경|세포|유전체|병리|약물|회계|세무|감사|감정평가|구조해석|전기|기계|설비|생산|품질관리|(?<!온)라인|건축|토목|영업|세일즈|sales|디자이너|MD|강사/i,
    exAlways: /국비|KDT|교육생|수강생|훈련생/, // 채용이 아닌 교육과정 모집
    exUnless: /데이터|data|SQL|애널리스트|analyst/i },
  // '자동화'만으로는 공장·설비 자동화까지 걸려서, AI·업무 자동화 단서가 있어야 포함하고 제조·설비 단서가 있으면 제외
  { name: 'AI 자동화', kw: 'AI 자동화', kws: ['AI 자동화', 'RPA', '업무 자동화'], re: /AI|인공지능|RPA|LLM|GPT|생성형|에이전트|agent|automation|업무\s*자동화|프로세스\s*자동화|하이퍼\s*오토메이션|노코드|n8n|make\.com|zapier/i,
    ex: /자동차|기계|설비|(?<!온)라인|생산(?!성)|공장|제조|PLC|제어|전기|전장|금형|용접|현장|로봇\s*(조작|운영)|반도체\s*장비|정비|시공|배관/ },
  { name: 'AX 컨설턴트', kw: 'AX 컨설턴트', kws: ['AX 컨설턴트', 'AI 전환', 'DX 컨설팅'], re: /AX|컨설턴트|컨설팅|consult|전환/i },
];
const SHEET = 'jobs_v3'; // 스키마(경력 열)가 바뀌어 새 시트 사용 (기존 시트는 그대로 둠)
const HEADERS = ['category', 'source', 'id', 'title', 'company', 'stack', 'location', 'contract', 'career', 'deadline', 'link', 'collected_at', 'salary', 'last_seen']; // 뒤의 두 열은 나중에 추가돼 맨 끝
const UA = { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/130 Safari/537.36' };
const CONTRACT_RE = /정규직|계약직|인턴|프리랜서|파견직|아르바이트|위촉직/;

const strip = s => String(s || '').replace(/<[^>]+>/g, '').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>')
  .replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/\s+/g, ' ').trim();
// 목록에 계약형태가 없는 사이트는 제목에서 추정, 없으면 '미표기'
const contractOf = s => ((String(s).match(CONTRACT_RE) || [])[0]) || (/계약/.test(s) ? '계약직' : '미표기');

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET) || ss.insertSheet(SHEET);
  if (sh.getLastRow() === 0) sh.appendRow(HEADERS);
  else if (sh.getLastColumn() < HEADERS.length) sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]); // 기존 시트에 새 열 헤더 추가
  return sh;
}

function get_(url, charset) {
  for (let i = 0; ; i++) {
    try { return UrlFetchApp.fetch(url, { headers: UA }).getContentText(charset || 'UTF-8'); }
    catch (e) { if (i >= 1) throw e; Utilities.sleep(3000); } // 일시적 차단·오류는 3초 뒤 한 번 더
  }
}

// 연봉 태그: '연봉 3,300만원~', '월급 360~400만원', '평균연봉 상위 10%', '연봉상승률 10% 이상', 제목 속 '연봉 6000만원'·'연봉인상' 등
const SALARY_RE = /(평균연봉|연봉상승률|연봉\s*인상률?|연봉|월급|시급)(?:[^,|\]\[()·]|,(?=\d)){0,24}/g; // 쉼표는 3,300처럼 숫자 앞일 때만 포함
const salaryOf = (...texts) => [...new Set(texts.join(' | ').match(SALARY_RE) || [])].map(s => s.trim().replace(/~\s*99,?990만원/, '만원~')) // 잡코리아의 상한 없음 표기
  .filter(s => /\d|인상|상위/.test(s) && !/협의|회사내규|면접 후/.test(s)).join(', ');

// 각 fetcher는 [source, id, title, company, stack, location, contract, career, deadline, link, salary] 를 반환
function fetchWanted_(c) {
  const d = JSON.parse(get_('https://www.wanted.co.kr/api/chaos/search/v1/results?tab=position&country=kr&limit=50&query=' + encodeURIComponent(c.kw)));
  return ((d.positions && d.positions.data) || []).map(j => ['원티드', 'w' + j.id, j.position, j.company.name, '',
    (j.address && j.address.location) || '', contractOf(j.position), j.is_newbie ? '신입·경력' : '경력', j.due_time ? j.due_time.slice(0, 10) : '상시',
    'https://www.wanted.co.kr/wd/' + j.id, salaryOf(j.position)]);
}

function fetchJumpit_(c) {
  const d = JSON.parse(get_('https://jumpit-api.saramin.co.kr/api/positions?page=1&keyword=' + encodeURIComponent(c.kw)));
  return d.result.positions.map(p => ['점핏', 'j' + p.id, strip(p.title), p.companyName, (p.techStacks || []).join(', '),
    (p.locations || []).join(', '), contractOf(p.title),
    p.newcomer && !p.maxCareer ? '신입' : (p.minCareer ? '경력 ' : '신입·경력 ') + p.minCareer + '~' + p.maxCareer + '년', p.alwaysOpen ? '상시' : String(p.closedAt).slice(0, 10),
    'https://jumpit.saramin.co.kr/position/' + p.id, salaryOf(p.title)]);
}

function fetchSaramin_(c) {
  const html = get_('https://www.saramin.co.kr/zf_user/search/recruit?recruitPageCount=50&searchword=' + encodeURIComponent(c.kw));
  // shortcut: 정규식 HTML 파싱이라 사람인 마크업이 바뀌면 깨짐 — 그때 정규식만 갱신
  return html.split('class="item_recruit"').slice(1).map(b => {
    const id = (b.match(/rec_idx=(\d+)/) || [])[1];
    const title = (b.match(/class="job_tit">[\s\S]*?title="([^"]*)"/) || [])[1];
    const corp = (b.match(/class="corp_name">[\s\S]*?>([\s\S]*?)<\/a>/) || [])[1];
    const date = (b.match(/class="date">([^<]*)</) || [])[1];
    const cond = [...((b.match(/class="job_condition">([\s\S]*?)<\/div>/) || [])[1] || '').matchAll(/<span>([\s\S]*?)<\/span>/g)].map(m => strip(m[1]));
    const sector = (b.match(/class="job_sector">([\s\S]*?)<\/div>/) || [])[1];
    const badges = [...b.matchAll(/class="badge[^"]*"[^>]*>([\s\S]*?)<\/span>/g)].map(m => strip(m[1]));
    if (!id || !title) return null;
    return ['사람인', 's' + id, strip(title), strip(corp), strip(sector).slice(0, 120), cond[0] || '',
      contractOf(cond.find(x => CONTRACT_RE.test(x)) || strip(title)), cond.find(x => /신입|경력/.test(x)) || '미표기', strip(date), // '정규직·계약직' 같은 복합값은 첫 형태로
      'https://www.saramin.co.kr/zf_user/jobs/relay/view?rec_idx=' + id, salaryOf(strip(title), ...badges)];
  }).filter(Boolean);
}

function fetchJobkorea_(c) {
  const html = get_('https://www.jobkorea.co.kr/Search/?stext=' + encodeURIComponent(c.kw));
  // shortcut: 정규식 HTML 파싱(Tailwind 클래스명 기준)이라 잡코리아 마크업이 바뀌면 깨짐
  const seen = {}, due = {};
  html.replace(/\\"/g, '"').split(/"id":"(?=\d+","legacyJobNo")/).slice(1).forEach(seg => {
    const e = seg.match(/"applicationPeriod":\{"start":"[^"]*","end":"(\d{4}-\d{2}-\d{2})/);
    if (e) due[seg.match(/^\d+/)[0]] = +e[1].slice(0, 4) > new Date().getFullYear() + 3 ? '상시' : e[1]; // 9999·2070년은 상시 표기
  });
  return html.split('data-sentry-component="CardJob"').slice(1).map(b => {
    const id = (b.match(/GI_Read\/(\d+)/) || [])[1];
    const title = strip((b.match(/text-typo-b1-18[^"]*">([^<]*)</) || [])[1]);
    const corp = strip((b.match(/text-typo-b2-16">([^<]*)</) || [])[1]);
    const chips = [...b.matchAll(/GrayChip[\s\S]*?text-typo-b4-14">([^<]*)</g)].map(m => strip(m[1]));
    if (!id || !title || seen[id]) return null;
    seen[id] = 1;
    const isPay = x => /연봉|월급|시급|주급|일급/.test(x);
    return ['잡코리아', 'k' + id, title, corp, chips.slice(1).filter(x => !isPay(x)).join(', '), chips[0] || '',
      contractOf(strip(b)), strip(([...b.matchAll(/text-typo-c1-13">([^<]*)</g)].map(m => m[1]).find(x => /신입|경력/.test(x))) || '미표기'), due[id] || '공고 참조', 'https://www.jobkorea.co.kr/Recruit/GI_Read/' + id,
      salaryOf(title, ...chips.filter(isPay))];
  }).filter(Boolean);
}

function fetchIncruit_(c) {
  const html = get_('https://search.incruit.com/list/search.asp?col=job&kw=' + encodeURIComponent(c.kw), 'EUC-KR');
  // shortcut: 정규식 HTML 파싱이라 인크루트 마크업이 바뀌면 깨짐
  return html.split('class="c_row').slice(1).map(b => {
    const id = (b.match(/jobno="(\d+)"/) || [])[1];
    const corp = strip((b.match(/class="cpname"[^>]*>([\s\S]*?)<\/a>/) || [])[1]);
    const mid = (b.match(/class="cell_mid">([\s\S]*?)<div class="cell_last"/) || [])[1] || '';
    const title = strip((mid.match(/class="cl_top"><a [^>]*>([\s\S]*?)<\/a>/) || [])[1]);
    const md = [...((mid.match(/class="cl_md">([\s\S]*?)<\/div>/) || [])[1] || '').matchAll(/<span>([\s\S]*?)<\/span>/g)].map(m => strip(m[1]));
    const sector = [...((mid.match(/class="cl_btm">([\s\S]*?)<\/div>/) || [])[1] || '').matchAll(/<span>([\s\S]*?)<\/span>/g)].map(m => strip(m[1]).replace(/,$/, ''));
    const last = (b.match(/class="cell_last"><div class="cl_btm"><span>([\s\S]*?)<\/span>/) || [])[1];
    if (!id || !title) return null;
    return ['인크루트', 'i' + id, title, corp, sector.join(', '), md[0] || '',
      contractOf(md.find(x => CONTRACT_RE.test(x)) || title), md.find(x => /신입|경력/.test(x)) || '미표기', strip(last) || '공고 참조',
      'https://job.incruit.com/jobdb_info/jobpost.asp?job=' + id, salaryOf(title, ...md)];
  }).filter(Boolean);
}

/** 직무 탭 포함 여부: 포함 단서(re)가 있고, 제목에 제외 단서(ex)가 없거나 예외(exUnless)가 있을 때 */
function inCategory_(c, title, stack) {
  if (!c.re.test(title + ' ' + stack) || (c.exAlways && c.exAlways.test(title))) return false;
  return !(c.ex && c.ex.test(title) && !(c.exUnless && c.exUnless.test(title)));
}

/** 수집: 직무×사이트별로 받아 새 공고(직무+id 기준)만 시트에 추가. 트리거로 매일 실행. */
function collectJobs() {
  const sh = getSheet_();
  const all = sh.getDataRange().getValues().slice(1);
  const seen = new Set(all.map(r => r[0] + '|' + r[2]));
  const rowOf = {}; all.forEach((r, i) => rowOf[r[0] + '|' + r[2]] = i);
  // 기존 공고도 다시 보이면 연봉·마감일·마지막 확인일을 갱신 (사라진 공고를 숨기는 기준이 됨)
  const COL = n => HEADERS.indexOf(n);
  const salaries = all.map(r => [r[COL('salary')] || '']);
  const deadlines = all.map(r => [r[COL('deadline')]]);
  const lastSeen = all.map(r => [r[COL('last_seen')] || r[COL('collected_at')]]);
  const now = new Date();
  const log = [];
  const stat = {}; // 사이트별 { raw: 받은 공고 수, added: 새로 추가, errors: [] }
  let rows = [];
  CATEGORIES.forEach(c => (c.kws || [c.kw]).forEach(kw => {
    SOURCES.forEach(([name, fn]) => {
      const st = stat[name] = stat[name] || { raw: 0, added: 0, errors: [] };
      try {
        const fetched = fn(Object.assign({}, c, { kw: kw }));
        st.raw += fetched.length;
        const got = fetched.filter(x => inCategory_(c, x[2], x[4]));
        got.forEach(x => { const k = c.name + '|' + x[1]; if (!(k in rowOf)) return; const i = rowOf[k];
          lastSeen[i] = [now]; if (x[10]) salaries[i] = [x[10]]; if (dueOf_(x[8], now)) deadlines[i] = [x[8]]; });
        const r = got.filter(x => !seen.has(c.name + '|' + x[1]));
        r.forEach(x => seen.add(c.name + '|' + x[1]));
        rows = rows.concat(r.map(x => [c.name].concat(x.slice(0, 10), [now, x[10] || '', now])));
        st.added += r.length;
      } catch (e) {
        st.errors.push(c.name + '/' + kw + ': ' + e.message.slice(0, 80)); // 한 곳이 막혀도 나머지는 저장
      }
    });
  }));
  SOURCES.forEach(([n]) => log.push(n + ' 받음 ' + stat[n].raw + ' · 추가 ' + stat[n].added + (stat[n].errors.length ? ' · 실패 ' + stat[n].errors.length : '')));
  if (all.length) [['salary', salaries], ['deadline', deadlines], ['last_seen', lastSeen]]
    .forEach(([n, v]) => sh.getRange(2, COL(n) + 1, v.length, 1).setValues(v));
  if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, HEADERS.length).setValues(rows);
  let ctErr = '';
  try { log.push('기업형태 ' + updateCompanyTypes_() + '곳'); } catch (e) { ctErr = e.message.slice(0, 120); log.push('기업형태 실패: ' + ctErr); }
  try { log.push('합쳐진 공고 ' + updateMergedSheet() + '그룹'); } catch (e) { log.push('합쳐진 공고 탭 실패: ' + e.message.slice(0, 80)); }
  recordRun_(now, stat, ctErr, rows.length);
  Logger.log(log.join(' | '));
  return log.join(' | ');
}

const SOURCES = [['원티드', fetchWanted_], ['점핏', fetchJumpit_], ['사람인', fetchSaramin_], ['잡코리아', fetchJobkorea_], ['인크루트', fetchIncruit_]];

/** 사이트 상태 판정: 전부 실패면 '실패', 받은 공고가 0건이면 '0건'(사이트 구조 변경 의심), 일부 실패면 '일부 실패' */
function sourceState_(st) {
  if (st.errors.length && !st.raw) return '실패';
  if (!st.raw) return '0건';
  return st.errors.length ? '일부 실패' : '정상';
}

/** 수집 기록을 'runs' 시트에 남기고, 문제가 있으면 스크립트 소유자에게 메일 */
function recordRun_(when, stat, ctErr, added) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('runs') || ss.insertSheet('runs');
  if (sh.getLastRow() === 0) sh.appendRow(['ran_at', 'added', 'status_json', 'problems']);
  const status = {};
  SOURCES.forEach(([n]) => status[n] = { state: sourceState_(stat[n]), raw: stat[n].raw, added: stat[n].added });
  const problems = SOURCES.filter(([n]) => status[n].state !== '정상')
    .map(([n]) => n + ' ' + status[n].state + (stat[n].errors[0] ? ' (' + stat[n].errors[0] + ')' : ''));
  if (ctErr) problems.push('기업형태 표 갱신 실패 (' + ctErr + ')');
  sh.appendRow([when, added, JSON.stringify(status), problems.join('\n')]);
  if (problems.length) {
    const to = Session.getEffectiveUser().getEmail();
    if (to) MailApp.sendEmail(to, '[채용 보드] 수집 문제 ' + problems.length + '건',
      problems.join('\n') + '\n\n0건은 사이트 화면 구조가 바뀌었을 가능성이 커요. 수집 기록: runs 시트');
  }
}

/** 'merged' 시트: 화면에서 한 카드로 합쳐진 공고를 사이트별 원래 제목과 함께 기록 (편집기에서 직접 실행해도 됨) */
function updateMergedSheet() {
  const groups = getJobs().filter(j => j.members && j.members.length > 1);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('merged') || ss.insertSheet('merged');
  const rows = [['group', 'category', '대표 제목', '대표 회사', '사이트', '사이트별 제목', '사이트별 회사', 'link', '대표 여부']];
  groups.forEach((g, i) => g.members.forEach((m, k) => rows.push([i + 1, g.category, g.title, g.company, m.source, m.title, m.company, m.link,
    k === g.rep ? '대표' : ''])));
  sh.clear();
  sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, rows[0].length).setFontWeight('bold').setBackground('#e7ecfb');
  return groups.length;
}

/** 화면 상단에 보여줄 마지막 수집 상태 */
function lastRun_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('runs');
  if (!sh || sh.getLastRow() < 2) return null;
  const r = sh.getRange(sh.getLastRow(), 1, 1, 3).getValues()[0];
  return { at: Utilities.formatDate(r[0], 'Asia/Seoul', 'M/d HH:mm'), status: JSON.parse(r[2] || '{}') };
}

// 기업 형태: 채용 목록엔 없어서 사람인의 기업형태 검색 필터로 '회사명 → 형태' 표를 만들고, 다른 사이트 공고에도 회사명으로 붙임
const COMPANY_TYPES = [['대기업', 'scale001'], ['중견기업', 'scale003'], ['중소기업', 'scale004'], ['스타트업', 'scale005'],
  ['외국계', 'foreign'], ['공기업·공공기관', 'public']];
const normName = s => String(s).replace(/\(주\)|㈜|주식회사|\(유\)|유한회사|\(재\)|\(사\)|\s/g, '').toLowerCase();

function updateCompanyTypes_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('company_types') || ss.insertSheet('company_types');
  const map = {};
  sh.getDataRange().getValues().slice(1).forEach(r => { if (r[0]) map[r[0]] = r[1]; }); // 이전 결과 유지
  CATEGORIES.forEach(c => COMPANY_TYPES.forEach(([label, code]) => {
    const html = get_('https://www.saramin.co.kr/zf_user/search/recruit?recruitPageCount=100&searchword=' +
      encodeURIComponent(c.kw) + '&company_type%5B%5D=' + code);
    html.split('class="item_recruit"').slice(1).forEach(b => {
      const corp = strip((b.match(/class="corp_name">[\s\S]*?>([\s\S]*?)<\/a>/) || [])[1]);
      if (corp) map[normName(corp)] = label;
    });
  }));
  const rows = Object.keys(map).map(k => [k, map[k]]);
  sh.clear();
  sh.getRange(1, 1, rows.length + 1, 2).setValues([['company_key', 'type']].concat(rows));
  return rows.length;
}

/** 매일 새벽 6~7시 자동 수집 트리거 설치 (한 번만 실행). Google 트리거는 지정한 시각부터 1시간 안의 임의 시점에 실행됨 */
function installTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('collectJobs').timeBased().everyDays(1).atHour(6).create();
}

function getJobs() {
  const range = getSheet_().getDataRange();
  const v = range.getDisplayValues();
  const raw = range.getValues(); // 첫 수집 시각(ms)을 NEW 배지에 사용
  const h = v.shift();
  const ct = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('company_types');
  const types = {};
  if (ct) ct.getDataRange().getValues().slice(1).forEach(r => types[r[0]] = r[1]);
  const today = ymd_(new Date());
  const staleBefore = ymd_(new Date(Date.now() - STALE_DAYS * 864e5));
  const jobs = v.map((r, ri) => {
    const o = Object.fromEntries(h.map((k, i) => [k, r[i]]));
    const ca = raw[ri + 1][h.indexOf('collected_at')];
    o.ts = ca instanceof Date ? ca.getTime() : 0;
    o.title = strip(o.title); // 예전에 저장된 점핏 제목의 <span> 태그 정리
    o.companyType = types[normName(o.company)] || '미분류';
    o.due = dueOf_(o.deadline, o.collected_at);
    o.seen = dueOf_(o.last_seen || o.collected_at, null) || today;
    return o;
  }).filter(o => { const c = CATEGORIES.find(x => x.name === o.category);
    return !c || inCategory_(c, o.title, o.stack); })
    // 마감일이 지났거나, 마감일 없이 STALE_DAYS일 넘게 검색에 안 보인 공고는 숨김 (시트에서 지우지는 않음)
    .filter(o => o.due ? o.due >= today : o.seen >= staleBefore)
    .reverse();
  return mergeDuplicates_(jobs);
}

const STALE_DAYS = 7;
const ymd_ = d => Utilities.formatDate(d, 'Asia/Seoul', 'yyyy-MM-dd');

/** 마감일 원문을 'yyyy-MM-dd'로. 상시·채용시·공고 참조 등 날짜가 없으면 ''.
 *  예: '2026-11-25', '2026. 11. 25', '~ 11/23(월)', '~12.11 (금)', '23시 마감'(=수집일), '내일마감' */
function dueOf_(raw, collected) {
  if (raw instanceof Date) return ymd_(raw);
  const s = String(raw || '').trim();
  let m = s.match(/(\d{4})\s*[-.\/년]\s*(\d{1,2})\s*[-.\/월]\s*(\d{1,2})/);
  if (m) return +m[1] > new Date().getFullYear() + 3 ? '' : m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
  if (collected === null) return '';
  const base = collected instanceof Date ? collected : new Date((dueOf_(collected, null) || ymd_(new Date())) + 'T12:00:00+09:00');
  m = s.match(/(\d{1,2})\s*[\/.]\s*(\d{1,2})/);
  if (m && +m[1] >= 1 && +m[1] <= 12) {
    const by = +ymd_(base).slice(0, 4), bm = +ymd_(base).slice(5, 7);
    const y = +m[1] < bm - 2 ? by + 1 : by; // 11월에 수집한 '01/05'는 다음 해
    return y + '-' + ('0' + m[1]).slice(-2) + '-' + ('0' + m[2]).slice(-2);
  }
  if (/내일\s*마감/.test(s)) return ymd_(new Date(base.getTime() + 864e5));
  if (/오늘\s*마감|\d+\s*시\s*마감/.test(s)) return ymd_(base);
  return '';
}

/** 같은 직무 탭 안에서 사이트만 다른 같은 공고를 하나로 합침
 *  조건: ① 서로 다른 사이트 ② 회사명이 같음(법인 표기·괄호 속 영문/브랜드·지점명 무시, 한쪽이 다른 쪽을 포함해도 인정)
 *        ③ 제목이 같거나, 유사도 0.75 이상이면서 신입/경력·계약직·기수·연차 같은 구분 단어가 충돌하지 않음 */
const normTitle = t => String(t).replace(/\[[^\]]*\]/g, '').replace(/[^0-9a-z가-힣()]/gi, '').toLowerCase();
const coKey_ = s => String(s).replace(/\(주\)|㈜|주식회사|\(유\)|유한회사|\(재\)|\(사\)|농업회사법인|사단법인|재단법인/g, '')
  .replace(/\([^)]*\)/g, '').replace(/\S*(사옥|지점|본점)/g, '').replace(/\s/g, '').toLowerCase();
const titleKey_ = t => strip(t).replace(/\[[^\]]*\]/g, '').replace(/채용\s*공고|채용|모집|공고|입니다|합니다/g, '')
  .replace(/[^0-9a-z가-힣]/gi, '').toLowerCase();
// 구분 단어: 서로 다르면 다른 공고로 봄 (예: 계약직 vs 경력직, 2기 vs 3기, 7년 이상 vs 2~7년, 인턴 vs 주니어)
const DIST_RE = /신입|경력|인턴|주니어|시니어|계약직|정규직|전환형|체험형|\d+\s*~\s*\d+\s*년|\d+\s*(년|기|차)|spanish|arabic|english|japanese|chinese/gi;
const distOf_ = t => new Set((strip(t).match(DIST_RE) || []).map(x => x.replace(/\s/g, '').toLowerCase()));
function bigrams_(s) { const a = new Set(); for (let i = 0; i < s.length - 1; i++) a.add(s.slice(i, i + 2)); return a; }
function similarity_(a, b) {
  const A = bigrams_(a), B = bigrams_(b); if (!A.size || !B.size) return 0;
  let n = 0; A.forEach(x => { if (B.has(x)) n++; }); return 2 * n / (A.size + B.size);
}
// 제목 단어 중 상대 제목에 없는 단어: 양쪽 모두에 있으면 직무가 다른 것으로 봄 (예: 전략컨설팅 vs 인사조직컨설팅)
const wordsOf_ = t => strip(t).replace(/\[[^\]]*\]/g, ' ').replace(/채용\s*공고|채용|모집|공고|입니다|합니다/g, ' ')
  .toLowerCase().split(/[^0-9a-z가-힣]+/).filter(w => w.length >= 2);
const onlyIn_ = (words, otherKey) => words.filter(w => !otherKey.includes(w));
function sameCompany_(a, b) {
  if (a === b) return true;
  const [s, l] = a.length < b.length ? [a, b] : [b, a];
  return s.length >= 3 && l.includes(s);
}
function samePosting_(a, b) {
  if (a.category !== b.category) return false;
  if (a.sources.some(x => b.sources.some(y => y.source === x.source))) return false; // 같은 사이트끼리는 합치지 않음
  if (!sameCompany_(a._co, b._co)) return false;
  if (a._tk === b._tk) return true;
  if (similarity_(a._tk, b._tk) < 0.75) return false;
  if (onlyIn_(a._words, b._tk).length && onlyIn_(b._words, a._tk).length) return false;
  const da = a._dist, db = b._dist;
  if (da.size && db.size && ([...da].some(x => !db.has(x)) || [...db].some(x => !da.has(x)))) return false;
  return true;
}
/** 대표 제목 점수: 단어 수가 많을수록, 같으면 글자 수가 길수록 (정보가 많은 제목) */
const infoScore_ = t => wordsOf_(t).length * 1000 + strip(t).length;
const FRAME_OK = ['사람인', '잡코리아', '인크루트', '점핏', '원티드']; // 분할 뷰에 먼저 띄울 순서 (원티드는 삽입 차단)
function mergeDuplicates_(jobs) {
  const out = [];
  jobs.forEach(j => {
    j.sources = j.sources || [{ source: j.source, link: j.link }];
    j.members = j.members || [{ source: j.source, title: j.title, company: j.company, link: j.link, stack: j.stack, location: j.location }];
    j._co = coKey_(j.company); j._tk = titleKey_(j.title); j._dist = distOf_(j.title); j._words = wordsOf_(j.title);
    const m = out.find(x => samePosting_(x, j));
    if (!m) { out.push(j); return; }
    j.sources.forEach(x => m.sources.push(x));
    j.members.forEach(x => m.members.push(x));
    if (!m.due || (j.due && j.due < m.due)) { m.due = j.due || m.due; }
    if (m.career === '미표기' && j.career !== '미표기') m.career = j.career;
    if (m.contract === '미표기' && j.contract !== '미표기') m.contract = j.contract;
    if (!m.location && j.location) m.location = j.location;
    if (j.salary) m.salary = [...new Set((m.salary ? m.salary.split(', ') : []).concat(j.salary.split(', ')))].join(', ');
    if (j.stack && !m.stack) m.stack = j.stack;
    if (m.companyType === '미분류' && j.companyType && j.companyType !== '미분류') m.companyType = j.companyType;
  });
  out.forEach(j => {
    if (j.members.length > 1) {
      const best = j.members.slice().sort((a, b) => infoScore_(b.title) - infoScore_(a.title))[0];
      j.rep = j.members.indexOf(best); // 'merged' 시트에서 대표 행 표시용
      j.title = best.title; j.company = best.company;
      const longest = k => j.members.map(x => x[k] || '').sort((a, b) => b.length - a.length)[0];
      j.stack = longest('stack'); j.location = longest('location');
    }
    // 분할 뷰는 화면 삽입이 되는 사이트를 먼저 연다
    j.sources.sort((a, b) => FRAME_OK.indexOf(a.source) - FRAME_OK.indexOf(b.source));
    j.source = j.sources[0].source; j.link = j.sources[0].link; delete j._co; delete j._tk; delete j._dist; delete j._words; });
  return out;
}

/** 웹앱: ?format=json 이면 JSON API, 아니면 보드 화면 */
function doGet(e) {
  if (e && e.parameter.format === 'json') {
    return ContentService.createTextOutput(JSON.stringify(getJobs())).setMimeType(ContentService.MimeType.JSON);
  }
  const t = HtmlService.createTemplate(PAGE);
  t.data = JSON.stringify(getJobs());
  t.cats = JSON.stringify(CATEGORIES.map(c => c.name));
  t.run = JSON.stringify(lastRun_());
  return t.evaluate().setTitle('채용 보드').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

const PAGE = `<!doctype html><html><head><meta charset="utf-8">
<link rel="stylesheet" href="https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/variable/pretendardvariable-dynamic-subset.min.css">
<style>
*{box-sizing:border-box}html,body{height:100%}
body{font-family:'Pretendard Variable',Pretendard,system-ui,'Malgun Gothic',sans-serif;margin:0;background:#f4f6fa;color:#1b2333;display:flex;flex-direction:column;letter-spacing:-.01em}
header{background:linear-gradient(120deg,#1e3a8a,#3b5bdb);color:#fff;padding:18px 20px 14px}
h1{margin:0;font-size:21px;font-weight:800;letter-spacing:-.02em}.sub{font-size:12.5px;opacity:.75;margin-top:3px}
.tabs{display:inline-flex;gap:4px;margin-top:14px;padding:4px;background:rgba(255,255,255,.14);border-radius:12px;overflow-x:auto;max-width:100%}
.tabs button{border:0;border-radius:9px;padding:8px 16px;background:transparent;color:rgba(255,255,255,.85);font:inherit;font-size:14.5px;font-weight:600;cursor:pointer;white-space:nowrap;display:flex;align-items:center;gap:6px;transition:background .15s,color .15s}
.tabs button:hover{background:rgba(255,255,255,.12);color:#fff}
.tabs button.on{background:#fff;color:#1e3a8a;box-shadow:0 2px 6px rgba(0,0,0,.12)}
.tabs .n{font-size:12px;font-weight:700;padding:1px 7px;border-radius:99px;background:rgba(255,255,255,.2)}
.tabs button.on .n{background:#e7ecfb;color:#3b5bdb}
main{flex:1;display:flex;min-height:0}
#left{flex:1 1 auto;overflow-y:auto;min-width:0}
.bar{display:flex;flex-wrap:wrap;align-items:center;gap:6px;padding:9px 16px;position:sticky;top:0;background:rgba(244,246,250,.92);backdrop-filter:blur(6px);z-index:1;border-bottom:1px solid #e4e8f0}
input,select,button{font:inherit;font-size:13.5px;height:32px;padding:0 10px;border:1px solid #d9dfea;border-radius:8px;background:#fff;color:#1b2333}
select{cursor:pointer}input:focus,select:focus{outline:2px solid #c7d2fe;border-color:#3b5bdb}
.bar select.set{border-color:#3b5bdb;background:#eef2ff;color:#1e3a8a;font-weight:600}
#q{flex:1;min-width:160px;padding-left:30px;background:#fff url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='14' height='14' fill='none' stroke='%238a94a6' stroke-width='2'%3E%3Ccircle cx='6' cy='6' r='4.5'/%3E%3Cpath d='M9.5 9.5 13 13'/%3E%3C/svg%3E") no-repeat 10px center}
.yrs{display:flex;align-items:center;gap:4px;height:32px;padding:0 8px;border:1px solid #d9dfea;border-radius:8px;background:#fff;font-size:13px;color:#5b6475}
.yrs.set{border-color:#3b5bdb;background:#eef2ff}
.yrs input{width:40px;height:24px;padding:0 4px;text-align:center;border-radius:6px;-moz-appearance:textfield}
.yrs input::-webkit-outer-spin-button,.yrs input::-webkit-inner-spin-button{-webkit-appearance:none;margin:0}
.ms{position:relative}
.ms>button{display:flex;align-items:center;gap:6px;cursor:pointer;white-space:nowrap}
.ms>button::after{content:'';width:6px;height:6px;border:solid #8a94a6;border-width:0 1.5px 1.5px 0;transform:rotate(45deg) translateY(-2px);margin-left:2px}
.ms.set>button{border-color:#3b5bdb;background:#eef2ff;color:#1e3a8a;font-weight:600}
.ms .b{font-size:11px;font-weight:700;background:#3b5bdb;color:#fff;border-radius:99px;padding:0 6px;line-height:16px}
.ms .pop{display:none;position:absolute;top:36px;left:0;z-index:10;min-width:170px;max-height:300px;overflow-y:auto;background:#fff;
  border:1px solid #e1e6ef;border-radius:10px;box-shadow:0 8px 24px rgba(20,30,60,.14);padding:6px}
.ms.open .pop{display:block}
.ms label{display:flex;align-items:center;gap:8px;padding:6px 8px;border-radius:6px;font-size:13.5px;cursor:pointer;white-space:nowrap}
.ms label:hover{background:#f2f5fb}.ms input{width:15px;height:15px;margin:0;accent-color:#3b5bdb}
.ms .cnt{margin-left:auto;color:#9aa4b8;font-size:12px}
.ms .clr{display:block;width:100%;margin-top:4px;height:28px;border:0;border-top:1px solid #eef1f6;border-radius:0;color:#3b5bdb;cursor:pointer;background:none}
#bm{cursor:pointer}#bm.on{background:#fff7d6;border-color:#f5b301;color:#8a6100;font-weight:600}
#reset{border:0;background:none;color:#3b5bdb;cursor:pointer;padding:0 6px}
#sort{cursor:pointer}
.chip.dd{background:#eef1f6;color:#3d4657;font-weight:700}.chip.dd.soon{background:#ffe4e4;color:#c92a2a}.chip.dd.today{background:#c92a2a;color:#fff}
.chip.new{background:#3b5bdb;color:#fff;font-weight:700}
.runst{display:inline-flex;flex-wrap:wrap;gap:6px;margin-left:8px}.runst span{font-size:11.5px;padding:1px 8px;border-radius:99px;background:rgba(255,255,255,.16)}
.runst span.bad{background:#ffd8d8;color:#a61e1e;font-weight:600}
.chip.more{background:#fff;border:1px solid #d9dfea;color:#5b6475}
#valt{display:flex;gap:6px;white-space:nowrap}#valt a{font-size:12.5px;padding:2px 8px;border:1px solid #d9dfea;border-radius:99px}
#count{margin-left:auto;color:#5b6475;font-size:13px;font-weight:600}
.grid{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:12px;padding:0 16px 24px}
body.split .grid{grid-template-columns:repeat(2,minmax(0,1fr))}
@media (max-width:900px){.grid{grid-template-columns:repeat(2,minmax(0,1fr))}}
@media (max-width:600px){.grid,body.split .grid{grid-template-columns:1fr}}
.card{background:#fff;border-radius:10px;padding:14px;box-shadow:0 1px 3px rgba(0,0,0,.08);position:relative;cursor:pointer;border:2px solid transparent}
.card:hover,.card.sel{border-color:#1f3a8a}.card h3{margin:6px 28px 4px 0;font-size:15px}
.chip{display:inline-block;font-size:12px;padding:2px 8px;border-radius:99px;background:#e8eefc;color:#1f3a8a;margin:0 4px 2px 0}
.chip.ct{background:#eef7ee;color:#256029}.chip.co{background:#f3eefc;color:#5b3aa8}.chip.sal{background:#fff4d6;color:#8a5a00;font-weight:600}.chip.cr{background:#fdf1e6;color:#8a4b12}
.meta{font-size:13px;color:#555;margin-top:6px}.stack{font-size:12px;color:#777;margin-top:6px}
.star{position:absolute;top:10px;right:10px;border:0;background:none;font-size:20px;cursor:pointer;padding:0;color:#bbb}.star.on{color:#f5b301}
#view{display:none;flex:0 0 var(--vw,50%);min-width:280px;max-width:80%;flex-direction:column;background:#fff}
body.split #view{display:flex}
#gutter{display:none;flex:0 0 8px;cursor:col-resize;background:#e4e8f0;position:relative;touch-action:none}
#gutter::after{content:'';position:absolute;left:3px;top:50%;width:2px;height:36px;margin-top:-18px;border-radius:2px;background:#9aa4b8}
#gutter:hover,#gutter.drag{background:#c7d2fe}
body.split #gutter{display:block}
body.dragging,body.dragging *{cursor:col-resize!important;user-select:none}
body.dragging #frame{pointer-events:none}
#vbar{display:flex;align-items:center;gap:8px;padding:8px 12px;border-bottom:1px solid #e3e6eb;font-size:14px}
#vtitle{flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-weight:600}
#vbar a{color:#1f3a8a;text-decoration:none;white-space:nowrap}#vclose{border:0;font-size:20px;padding:2px 8px;cursor:pointer}
#frame{flex:1;border:0;width:100%}
#blocked{display:none;padding:40px 24px;text-align:center;color:#555}
.go{display:inline-block;margin-top:14px;padding:12px 20px;border-radius:8px;background:#1f3a8a;color:#fff;text-decoration:none;font-weight:600}
@media (max-width:700px){body.split #left,body.split #gutter{display:none}body.split #view{flex:1;max-width:none}}
</style></head><body>
<header><h1>채용 보드</h1><div class="sub">원티드 · 점핏 · 사람인 · 잡코리아 · 인크루트 자동 수집 · 매일 06~07시<span class="runst" id="runst"></span></div>
<div class="tabs" id="tabs" role="tablist"></div></header>
<main><div id="left">
<div class="bar"><input id="q" placeholder="회사·직무·기술스택 검색 (예: SQL, MVP)" aria-label="키워드 검색">
<div class="ms" id="src"></div><div class="ms" id="ctype"></div><div class="ms" id="region"></div><div class="ms" id="contract"></div><div class="ms" id="career"></div>
<label class="yrs" id="yrs">연차 <input id="ymin" type="number" min="0" max="30" inputmode="numeric" aria-label="최소 연차">~<input id="ymax" type="number" min="0" max="30" inputmode="numeric" aria-label="최대 연차">년</label>
<button id="bm" aria-pressed="false">★ 북마크</button><button id="reset">초기화</button><span id="count"></span>
<select id="sort" aria-label="정렬"><option value="new">최신 수집순</option><option value="due">마감 임박순</option><option value="sal">연봉 정보 우선</option></select></div>
<div class="grid" id="list"></div></div>
<div id="gutter" role="separator" aria-orientation="vertical" aria-label="분할 뷰 크기 조절" tabindex="0"></div>
<section id="view" aria-label="공고 원문"><div id="vbar"><span id="vtitle"></span><span id="valt"></span><a id="vopen" target="_blank" rel="noopener">새 탭 ↗</a>
<button id="vclose" aria-label="분할 뷰 닫기">×</button></div>
<iframe id="frame" title="공고 원문" sandbox="allow-scripts allow-same-origin allow-popups allow-forms" referrerpolicy="no-referrer"></iframe>
<div id="blocked">이 사이트(원티드)는 다른 페이지 안에서 열리는 것을 막고 있어요.<br><a id="bopen" class="go" target="_blank" rel="noopener">원문 공고 새 탭에서 보기 ↗</a></div>
</section></main>
<script>
const JOBS = <?!= data ?>, CATS = <?!= cats ?>, RUN = <?!= run ?>;
// NEW 배지: 지난 방문(30분 이상 전) 이후 처음 수집된 공고. 첫 방문이면 최근 24시간
const PREV = (() => { try { const last = +localStorage.getItem('lastVisit') || 0, prev = +localStorage.getItem('prevVisit') || 0;
  const now = Date.now(); if (now - last > 30 * 60e3) { localStorage.setItem('prevVisit', last); localStorage.setItem('lastVisit', now); return last || now - 864e5; }
  return prev || now - 864e5; } catch (e) { return Date.now() - 864e5; } })();
const isNew = j => j.ts && j.ts > PREV;
const NO_FRAME = ['원티드']; // X-Frame-Options로 삽입 차단
let cat = CATS[0], onlyBm = false, selId = null;
const $ = id => document.getElementById(id);
const bm = new Set(JSON.parse((()=>{try{return localStorage.getItem('bm')}catch(e){}})()||'[]'));
const esc = s => String(s).replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const key = j => j.category + '|' + j.id;
const regionOf = j => (String(j.location).trim().split(/[\\s,·\\/]+/)[0] || '미표기').replace(/특별시|광역시|특별자치시|특별자치도|도$/,'') || '미표기';
// 사이트마다 표기가 달라('경력3~5년', '신입·경력1년↑', '경력무관') 필터용으로 묶음
const careerOf = j => { const c = String(j.career||''); return /무관/.test(c) ? '경력무관' : /신입/.test(c) && /경력/.test(c) ? '신입·경력' :
  /신입/.test(c) ? '신입' : /경력|\\d년/.test(c) ? '경력' : '미표기'; };
// 연차 범위 [min,max]: '경력3~5년'→[3,5], '경력3년↑'→[3,99], '신입'→[0,0], '신입·경력1년↑'→[0,99], '경력무관'→[0,99], 미표기→null
const rangeOf = j => { const c = String(j.career||'');
  if (/무관/.test(c)) return [0,99];
  const m = c.match(/(\\d+)\\s*~\\s*(\\d+)/); if (m) return [/신입/.test(c) ? 0 : +m[1], +m[2]];
  const u = c.match(/(\\d+)\\s*년/); const lo = /신입/.test(c) ? 0 : u ? +u[1] : 1;
  if (u || /경력/.test(c)) return [lo, 99];
  return /신입/.test(c) ? [0,0] : null; };
function yearsOk(j, a, b){ if (a==='' && b==='') return true; const r = rangeOf(j); if (!r) return false;
  const lo = a==='' ? 0 : +a, hi = b==='' ? 99 : +b; return r[0] <= hi && r[1] >= lo; }
// 다중 선택 필터: 각 필터는 선택값 Set, 비어 있으면 '전체'
const MS = { src: ['사이트', j => (j.sources || [{source: j.source}]).map(x => x.source)], ctype: ['기업형태', j => j.companyType || '미분류'], region: ['지역', regionOf],
  contract: ['계약형태', j => j.contract || '미표기'], career: ['경력', careerOf] };
const SEL = Object.fromEntries(Object.keys(MS).map(k => [k, new Set()]));
const ORDER = { ctype: ['대기업','중견기업','중소기업','스타트업','외국계','공기업·공공기관','미분류'], career: ['신입','신입·경력','경력','경력무관','미표기'],
  src: ['원티드','점핏','사람인','잡코리아','인크루트'] };
function fillMulti(id, inCat){
  const [label, fn] = MS[id], cnt = {};
  inCat.forEach(j => [].concat(fn(j)).forEach(v => cnt[v] = (cnt[v]||0) + 1));
  const vals = ORDER[id] ? ORDER[id].filter(v => cnt[v]) : Object.keys(cnt).sort();
  [...SEL[id]].forEach(v => { if (!cnt[v]) SEL[id].delete(v); });
  const n = SEL[id].size, open = $(id).classList.contains('open');
  $(id).innerHTML = '<button type="button" aria-haspopup="true" aria-expanded="'+open+'">'+label+(n?' <span class="b">'+n+'</span>':'')+'</button>'+
    '<div class="pop" role="group" aria-label="'+label+'">'+vals.map(v => '<label><input type="checkbox" data-v="'+esc(v)+'"'+(SEL[id].has(v)?' checked':'')+'>'+
    esc(v)+'<span class="cnt">'+cnt[v]+'</span></label>').join('')+(n?'<button type="button" class="clr">선택 해제</button>':'')+'</div>';
  $(id).classList.toggle('set', n > 0);
}
function renderTabs(){
  $('tabs').innerHTML = CATS.map(c => '<button role="tab" aria-selected="'+(c===cat)+'" class="'+(c===cat?'on':'')+'" data-cat="'+esc(c)+'">'+esc(c)+'<span class="n">'+JOBS.filter(j=>j.category===c).length+'</span></button>').join('');
  const inCat = JOBS.filter(j => j.category === cat);
  Object.keys(MS).forEach(id => fillMulti(id, inCat));
}
function render(){
  const q = $('q').value.toLowerCase(), ya = $('ymin').value, yb = $('ymax').value;
  $('yrs').classList.toggle('set', ya!=='' || yb!=='');
  const list = JOBS.filter(j => j.category===cat && Object.keys(MS).every(id => !SEL[id].size || [].concat(MS[id][1](j)).some(v => SEL[id].has(v))) &&
    yearsOk(j, ya, yb) && (!onlyBm || bm.has(key(j))) &&
    (j.title+' '+j.company+' '+j.stack+' '+j.location+' '+(j.salary||'')).toLowerCase().includes(q));
  if ($('sort').value === 'sal') list.sort((a, b) => (b.salary ? 1 : 0) - (a.salary ? 1 : 0));
  if ($('sort').value === 'due') list.sort((a, b) => (a.due || '9999') < (b.due || '9999') ? -1 : (a.due || '9999') > (b.due || '9999') ? 1 : 0);
  $('count').textContent = list.length + '건';
  $('list').innerHTML = list.map(j => '<div class="card'+(key(j)===selId?' sel':'')+'" data-k="'+esc(key(j))+'" tabindex="0">'+
    '<button class="star'+(bm.has(key(j))?' on':'')+'" data-star="'+esc(key(j))+'" aria-label="북마크">★</button>'+
    (isNew(j)?'<span class="chip new">NEW</span>':'')+ddChip(j)+(j.sources||[{source:j.source}]).map(x=>'<span class="chip">'+esc(x.source)+'</span>').join('')+(j.companyType&&j.companyType!=='미분류'?'<span class="chip co">'+esc(j.companyType)+'</span>':'')+
    '<span class="chip ct">'+esc(j.contract||'미표기')+'</span><span class="chip cr">'+esc(j.career||'경력 미표기')+'</span>'+
    (j.salary?String(j.salary).split(', ').map(x=>'<span class="chip sal">'+esc(x)+'</span>').join(''):'')+
    '<h3>'+esc(j.title)+'</h3><div class="meta">'+esc(j.company)+' · '+esc(j.location)+' · 마감 '+esc(j.due || j.deadline)+'</div>'+
    (j.stack?'<div class="stack">'+esc(j.stack)+'</div>':'')+'</div>').join('');
}
document.querySelector('.bar').addEventListener('click', e => {
  const ms = e.target.closest('.ms'); if (!ms) return;
  if (e.target.matches('.ms>button')) { const was = ms.classList.contains('open');
    document.querySelectorAll('.ms.open').forEach(x => x.classList.remove('open')); ms.classList.toggle('open', !was);
    ms.firstChild.setAttribute('aria-expanded', !was); return; }
  if (e.target.matches('.clr')) { SEL[ms.id].clear(); renderTabs(); render(); }
});
document.querySelector('.bar').addEventListener('change', e => {
  const ms = e.target.closest('.ms'); if (!ms) return;
  const v = e.target.dataset.v; e.target.checked ? SEL[ms.id].add(v) : SEL[ms.id].delete(v);
  renderTabs(); render();
});
document.addEventListener('click', e => { if (!e.target.closest('.ms')) document.querySelectorAll('.ms.open').forEach(x => x.classList.remove('open')); });
const TODAY = new Intl.DateTimeFormat('en-CA', {timeZone: 'Asia/Seoul'}).format(new Date());
function ddChip(j){ if (!j.due) return '';
  const d = Math.round((new Date(j.due) - new Date(TODAY)) / 864e5);
  return '<span class="chip dd'+(d===0?' today':d<=3?' soon':'')+'" title="마감 '+esc(j.due)+'">'+(d===0?'오늘 마감':'D-'+d)+'</span>'; }
function openSplit(k){
  const j = JOBS.find(x => key(x) === k); if(!j) return; selId = k;
  const blocked = NO_FRAME.includes(j.source);
  $('vtitle').textContent = j.company + ' · ' + j.title; $('vopen').href = $('bopen').href = j.link;
  $('valt').innerHTML = (j.sources||[]).slice(1).map(x => '<a href="'+esc(x.link)+'" target="_blank" rel="noopener">'+esc(x.source)+' ↗</a>').join('');
  $('frame').style.display = blocked ? 'none' : 'block'; $('blocked').style.display = blocked ? 'block' : 'none';
  $('frame').src = blocked ? 'about:blank' : j.link;
  document.body.classList.add('split'); render();
}
function closeSplit(){ document.body.classList.remove('split'); $('frame').src = 'about:blank'; selId = null; render(); }
$('tabs').onclick = e => { const c = e.target.dataset.cat; if(!c) return; cat = c; renderTabs(); render(); };
$('list').onclick = e => { const st = e.target.dataset.star;
  if (st) { bm.has(st)?bm.delete(st):bm.add(st); try{localStorage.setItem('bm', JSON.stringify([...bm]))}catch(e){} render(); return; }
  const card = e.target.closest('.card'); if (card) openSplit(card.dataset.k); };
$('list').onkeydown = e => { const card = e.target.closest('.card'); if (card && e.key==='Enter') openSplit(card.dataset.k); };
$('vclose').onclick = closeSplit;
document.onkeydown = e => { if (e.key!=='Escape') return; const o = document.querySelector('.ms.open'); if (o) { o.classList.remove('open'); return; } if (document.body.classList.contains('split')) closeSplit(); };
$('bm').onclick = () => { onlyBm=!onlyBm; $('bm').classList.toggle('on', onlyBm); $('bm').setAttribute('aria-pressed', onlyBm); render(); };
$('reset').onclick = () => { ['q','ymin','ymax'].forEach(id => $(id).value = ''); Object.values(SEL).forEach(x => x.clear()); renderTabs();
  onlyBm = false; $('bm').classList.remove('on'); render(); };
$('q').oninput = $('ymin').oninput = $('ymax').oninput = render;
$('sort').onchange = render;
// 분할 뷰 크기 조절: 드래그(또는 ←/→ 키)로 오른쪽 폭(%) 변경, 브라우저에 기억
const setVw = p => { p = Math.min(80, Math.max(20, p)); document.body.style.setProperty('--vw', p + '%'); try{localStorage.setItem('vw', p)}catch(e){} };
try{ const v = +localStorage.getItem('vw'); if (v) setVw(v); }catch(e){}
$('gutter').onpointerdown = e => { e.preventDefault(); $('gutter').setPointerCapture(e.pointerId);
  document.body.classList.add('dragging'); $('gutter').classList.add('drag'); };
$('gutter').onpointermove = e => { if (!$('gutter').classList.contains('drag')) return;
  const w = document.querySelector('main').getBoundingClientRect(); setVw((w.right - e.clientX) / w.width * 100); };
$('gutter').onpointerup = $('gutter').onpointercancel = () => { document.body.classList.remove('dragging'); $('gutter').classList.remove('drag'); };
$('gutter').onkeydown = e => { const cur = parseFloat(getComputedStyle(document.body).getPropertyValue('--vw')) || 50;
  if (e.key==='ArrowLeft') setVw(cur + 5); if (e.key==='ArrowRight') setVw(cur - 5); };
if (RUN) $('runst').innerHTML = '<span>마지막 수집 '+esc(RUN.at)+'</span>' + Object.keys(RUN.status).map(n => { const x = RUN.status[n];
  return x.state === '정상' ? '' : '<span class="bad" title="받은 공고 '+x.raw+'건">'+esc(n)+' '+esc(x.state)+'</span>'; }).join('');
renderTabs(); render();
</script></body></html>`;
