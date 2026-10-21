/* 설비고장 원인 진단 도우미 — 순수 로직 (화면·저장소와 무관)
   브라우저에서는 window.FixLogic, node 에서는 require('./logic.js') 로 씁니다.
   기준 문서: docs/01_프로젝트_기획서.md (3장 데이터, 5장 기능, 8장 1단계)

   흐름: 정비 사례(손으로 등록 · 정비 내역서 엑셀 가져오기)
        → similar(증상 글자 2개 묶음 TF-IDF + 증상 태그 + 같은 설비) → diagnose(원인별 근거 사례 수·비율 · 먼저 확인할 것)
        → guideDraft(원인별 자체 정비 가이드 초안) · costReport(외부 수리 비용·자체 처리 비율)
        → buildDiagnosePrompt · buildGuidePrompt (AI 반자동) */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.FixLogic = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var SCHEMA_VERSION = 1;
  var REPAIR = ['외부', '자체'];

  /* ── 1. 기본 도우미 ─────────────────────────────────────────── */
  function str(v) { return v == null ? '' : String(v).trim(); }
  function pad(n) { return (n < 10 ? '0' : '') + n; }
  function ymd(y, m, d) { return y + '-' + pad(m) + '-' + pad(d); }
  function validYmd(y, m, d) {
    if (!(y >= 1990 && y <= 2100 && m >= 1 && m <= 12 && d >= 1 && d <= 31)) return false;
    var t = new Date(Date.UTC(y, m - 1, d));
    return t.getUTCMonth() === m - 1 && t.getUTCDate() === d;
  }
  /* 날짜 → 'YYYY-MM-DD'. Date · 엑셀 일련번호 · 2026-03-05 · 2026.3.5 · 2026/03/05 · 26.3.5 · 20260305 */
  function toISO(v) {
    if (v == null || v === '') return '';
    if (v instanceof Date) {
      if (isNaN(v.getTime())) return '';
      // SheetJS 는 날짜를 그 지역 자정으로 줍니다. 지역 날짜로 읽습니다.
      return ymd(v.getFullYear(), v.getMonth() + 1, v.getDate());
    }
    if (typeof v === 'number' && isFinite(v)) {
      if (v > 20000 && v < 80000) { // 엑셀 일련번호
        var t = new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 86400000);
        return ymd(t.getUTCFullYear(), t.getUTCMonth() + 1, t.getUTCDate());
      }
      if (v >= 19900101 && v <= 21001231) v = String(v); else return '';
    }
    var s = str(v).replace(/\s+/g, '');
    var m = /^(\d{4})[-./년](\d{1,2})[-./월](\d{1,2})/.exec(s);
    if (!m) { m = /^(\d{2})[-./](\d{1,2})[-./](\d{1,2})$/.exec(s); if (m) m[1] = String(2000 + Number(m[1])); }
    if (!m) { m = /^(\d{4})(\d{2})(\d{2})$/.exec(s); }
    if (!m) return '';
    var y = Number(m[1]), mo = Number(m[2]), d = Number(m[3]);
    return validYmd(y, mo, d) ? ymd(y, mo, d) : '';
  }
  function monthOf(date) { return date ? String(date).slice(0, 7) : ''; }
  /* 금액·수량 → 숫자. 「1,250,000원」 「₩680,000」 「15만원」 「1.2백만」 */
  function toNum(v) {
    if (v == null || v === '') return null;
    if (typeof v === 'number') return isFinite(v) ? v : null;
    var s = String(v).replace(/[\s,₩원]/g, '');
    if (!s || s === '-') return null;
    var man = /^(-?\d+(?:\.\d+)?)만$/.exec(s);
    if (man) return Math.round(Number(man[1]) * 10000);
    var mm = /^(\d+)만(\d+)$/.exec(s);
    if (mm) return Number(mm[1]) * 10000 + Number(mm[2]);
    var n = Number(s);
    return isFinite(n) ? n : null;
  }
  /* 소요시간 → 시간(소수). 「2」 「1.5h」 「1시간 30분」 「90분」 「2시간」 */
  function toHours(v) {
    if (v == null || v === '') return null;
    if (typeof v === 'number') return isFinite(v) && v >= 0 ? v : null;
    var s = String(v).replace(/\s+/g, '').toLowerCase();
    var hm = /^(\d+(?:\.\d+)?)(?:시간|hr|h)(?:(\d+)(?:분|m|min))?$/.exec(s);
    if (hm) return Number(hm[1]) + (hm[2] ? Number(hm[2]) / 60 : 0);
    var mi = /^(\d+)(?:분|m|min)$/.exec(s);
    if (mi) return Math.round(Number(mi[1]) / 60 * 100) / 100;
    var n = Number(s);
    return isFinite(n) && n >= 0 ? n : null;
  }
  /* 느슨한 날짜 — 수리이력 보고서에 흔한 「2025-03 2주차」 「2025-03~04」 「2026-01 (보고서 기준)」 「3월 2주차」(+연도 칸).
     { date: 'YYYY-MM-DD', prec: 'day'|'week'|'month' }. 주차는 그 달의 (주-1)*7+1 일로, 월은 1일로 둡니다(달 집계용) */
  var MISSING_RE = /^(미기재|미확인|미상|모름|없음|해당\s?없음|-|—|\?|n\/a|na)$/i;
  function looseDate(v, year) {
    if (v == null || v === '') return { date: '', prec: '' };
    var iso = toISO(v);
    if (iso) return { date: iso, prec: 'day' };
    var s = str(v);
    if (MISSING_RE.test(s)) return { date: '', prec: '' };
    var m = /(\d{4})\s*[-./년]\s*(\d{1,2})\s*(?:월)?(?:\s*~\s*\d{1,2}\s*월?)?(?:\s*\(?\s*(\d)\s*(?:~\s*\d\s*)?주\s*차?)?/.exec(s);
    var y, mo, w;
    if (m) { y = Number(m[1]); mo = Number(m[2]); w = m[3] ? Number(m[3]) : 0; }
    else {
      m = /(?:^|\D)(\d{1,2})\s*월(?:\s*(\d)\s*(?:~\s*\d\s*)?주\s*차?)?/.exec(s);
      var yy = Number(year);
      if (!m || !(yy >= 1990 && yy <= 2100)) return { date: '', prec: '' };
      y = yy; mo = Number(m[1]); w = m[2] ? Number(m[2]) : 0;
    }
    if (!(mo >= 1 && mo <= 12 && y >= 1990 && y <= 2100)) return { date: '', prec: '' };
    if (w >= 1 && w <= 5) return { date: ymd(y, mo, Math.min(28, (w - 1) * 7 + 1)), prec: 'week' };
    return { date: ymd(y, mo, 1), prec: 'month' };
  }
  var PREC_TEXT = { day: '일', week: '주차', month: '월' };
  /* 비용 칸 → { value, kind }. kind: won(금액) · free(무상, 0) · missing(미기재 — 0 이 아님) · estimate(견적 — 실제 지출과 따로)
     · unit(「40,000원/개」 개당 단가 — 수량을 모름) · unread(글은 있는데 금액이 없음).
     「견적 2,000,000원 + 운반비 300,000원」 처럼 금액이 여럿이면 더합니다. */
  function parseCost(v) {
    if (v == null || v === '') return { value: null, kind: 'missing' };
    if (typeof v === 'number') return isFinite(v) && v >= 0 ? { value: v, kind: 'won' } : { value: null, kind: 'unread' };
    var s = str(v);
    if (!s || MISSING_RE.test(s) || /^(금액\s?)?(미기재|미확인)/.test(s)) return { value: null, kind: 'missing' };
    if (/^(무상|무료|0원|보증|무상\s?(수리|교체|처리))/.test(s.replace(/\s+/g, ''))) return { value: 0, kind: 'free' };
    var plain = toNum(s);
    var est = /견적/.test(s), unit = /\/\s*(개|ea|대|조|set|세트)|개당|ea당|단가/i.test(s);
    var sum = 0, n = 0, re = /(\d{1,3}(?:,\d{3})+|\d+(?:\.\d+)?)\s*(만|천)?\s*원/g, m;
    while ((m = re.exec(s))) { var x = Number(m[1].replace(/,/g, '')) * (m[2] === '만' ? 10000 : m[2] === '천' ? 1000 : 1); sum += x; n++; }
    var value = n ? Math.round(sum) : plain;
    if (value == null || !isFinite(value) || value < 0) return { value: null, kind: 'unread' };
    return { value: value, kind: est ? 'estimate' : unit ? 'unit' : 'won' };
  }
  var COST_KIND_TEXT = { won: '금액', free: '무상', missing: '미기재', estimate: '견적', unit: '개당 단가', unread: '못 읽음' };
  /* 실제 지출로 셀 금액 — 견적은 빼고(따로 셈), 미기재는 0 이 아니라 「모름」 */
  function spent(c) { return c && c.costKind !== 'estimate' && c.cost != null && isFinite(c.cost) ? Number(c.cost) : 0; }
  /* 수리 업체 칸 → { vendor, by }. 「미기재」 → 모름, 「자체 조치」 → 자체, 그 밖의 이름(「…(점검 예정)」 「전기기사」 「A/S」) → 외부 */
  function parseVendor(v) {
    var s = str(v);
    if (!s || MISSING_RE.test(s)) return { vendor: '', by: '' };
    var core = s.replace(/\s*[(（][^)）]*[)）]\s*/g, ' ').replace(/(업체명?\s*)?(미기재|미확인)/g, ' ').replace(/\s+/g, ' ').trim();
    if (/자체|사내|직접|내부\s?조치|팀\s?조치/.test(s)) return { vendor: core || '자체', by: '자체' };
    if (!core || MISSING_RE.test(core) || /^[/·,\s-]+$/.test(core)) return { vendor: '', by: '' };
    return { vendor: core, by: '외부' };
  }
  /* 처리 상태 — 처리일자·비고·조치 글에서. 완료 · 진행 중 · ''(모름) */
  function caseStatus(text) {
    var s = str(text).replace(/완료\s?(여부|일)\s?(미확인|미기재|확인\s?필요)?/g, ' ');
    if (!s) return '';
    var done = /완료|정상\s?(작동|가동|화)?|해결|개선|제거|없음|교체함|교체했|충전\s?후|주입|부착|체결\s?후/.test(s);
    var pending = /예정|진행\s?중|진행$|수리\s?진행|교체\s?진행|가능일|조율|요청|확인\s?중|재의뢰|수배\s?중|필요로\s?보고|교체\s?필요|수리\s?필요|구매\s?필요|점검\s?중|대기/.test(s);
    if (done) return '완료';
    if (pending) return '진행 중';
    if (/교체$|교체\s*[,.;]|수리$|설치$/.test(s)) return '완료';
    return '';
  }
  /* 처리일자 칸에 날짜와 조치가 함께 적힌 경우(「2026-02 (보고서 기준) 온도 조절계 교체 완료」) 날짜 부분을 떼어 조치만 */
  function actionOf(text) {
    var s = str(text);
    if (!s) return '';
    s = s.replace(/^\s*\d{4}\s*[-./]\s*\d{1,2}(?:\s*[-./]\s*\d{1,2})?(?:\s*~\s*\d{1,2})?(?:\s*\d\s*(?:~\s*\d\s*)?주\s*차)?\s*/, '')
      .replace(/^\s*\(?\s*보고서\s?기준\s*\)?\s*/, '').replace(/^중\s+/, '').trim();
    if (!s || MISSING_RE.test(s)) return '';
    return s;
  }
  function round(v, d) { var p = Math.pow(10, d || 0); return Math.round(v * p) / p; }
  function fmtWon(n) { return n == null || !isFinite(n) ? '-' : Math.round(n).toLocaleString('ko-KR') + '원'; }
  function fmtPct(r) { return r == null || !isFinite(r) ? '-' : Math.round(r * 100) + '%'; }
  /* 비교용으로 공백·기호를 지우고 소문자로 */
  function squash(s) { return str(s).toLowerCase().replace(/[\s_\-()（）\[\]·.,:;/'"~!?<>{}|+*=#@&%^`]/g, ''); }

  /* ── 2. 외부·자체, 단계, 부품 ───────────────────────────────── */
  /* 「외부업체」 「A/S」 「외주」 → 외부, 「자체」 「사내」 「직접」 「팀」 → 자체 */
  function normRepairBy(v, cost) {
    var s = squash(v);
    if (s) {
      if (/자체|사내|직접|팀|내부|본인/.test(s)) return '자체';
      if (/외부|업체|외주|as|a\/s|협력|출장|메이커|제조사|서비스/.test(s)) return '외부';
    }
    return '';
  }
  /* 조치 내용 한 칸 → 단계 목록. 줄바꿈 · 「1) 2)」 「①②」 · 「→」 · 「 / 」 · 「;」 로 나눕니다 */
  function splitSteps(v) {
    if (Array.isArray(v)) return v.map(str).filter(Boolean);
    var s = str(v);
    if (!s) return [];
    var parts;
    if (/\n/.test(s)) parts = s.split(/\r?\n/);
    else {
      // 「85dB→76dB」 처럼 숫자 사이 화살표, 「2.5일」 같은 소수는 단계 구분이 아닙니다
      s = s.replace(/(\d[a-zA-Z%℃]*)\s*(→|->)\s*(?=\d)/g, '$1\u21E2')
        .replace(/(^|\s)(\d{1,2}[.)](?!\d)|[①-⑳])\s*/g, '\n').replace(/\s*(→|->|=>|;|\s\/\s)\s*/g, '\n');
      parts = s.split('\n').map(function (p) { return p.replace(/\u21E2/g, '→'); });
    }
    return parts.map(function (p) { return str(p).replace(/^(\d{1,2}[.)](?!\d)|[①-⑳]|[-*•·])\s*/, ''); }).filter(Boolean);
  }
  function splitList(v) {
    if (Array.isArray(v)) return v.map(str).filter(Boolean);
    return str(v).split(/\s*(?:[,/;·\n]|\s및\s)\s*/).map(str).filter(Boolean);
  }

  /* ── 3. 증상 태그 ─────────────────────────────────────────────
     증상 글에서 자동으로 뽑고, 사용자가 켜고 끌 수 있습니다. 같은 태그가 있으면 유사도에 더해집니다. */
  var TAGS = [
    { tag: '온도 안 오름', re: /온도.{0,8}(안\s?올라|안\s?오르|안\s?오름|못\s?올라|낮|미달|저하|떨어|상승하지|상승\s?불량|상승이\s?안)|승온.{0,3}(불량|안|지연|느|이 안)|히팅.{0,2}(안|불량)/ },
    { tag: '온도 표시 불안정', re: /온도.{0,8}(들쭉|튀|불안정|오락가락|흔들|왔다갔다|0\s?℃|0도)|표시\s?온도.{0,6}(다름|낮|이상)/ },
    { tag: '과열', re: /과열|너무\s?뜨|뜨거|온도.{0,4}(높|올라감|초과|상승으로|상승해)|열\s?남|열이\s?남|고온으로/ },
    { tag: '설비 정지', re: /멈춤|멈춰|멈췄|멈추|정지|섰|꺼짐|꺼졌|다운|작동\s?불가|미작동|구동\s?불가|작동\s?안|생산\s?불가|작업\s?중단|가동\s?중지/ },
    { tag: '알람', re: /알람|경보|에러|alarm|오류\s?표시/ },
    { tag: '이상 소음', re: /소음|소리|딱딱|끼익|웅웅|덜컹|굉음|이음/ },
    { tag: '진동', re: /진동|떨림/ },
    { tag: '누유', re: /누유|기름|오일.{0,4}(샘|새|묻|자국)|새어|새고|샌다/ },
    { tag: '누수·냉각 불량', re: /누수|냉각.{0,4}(불량|안\s?됨|약|이상)|냉매|프레온|가스\s?충전/ },
    { tag: '압력 저하', re: /압력.{0,4}(저하|떨어|낮|늦게|안\s?올라)|형체력.{0,3}약|출력값?.{0,3}저하/ },
    { tag: '차단기·과부하', re: /차단기|트립|써멀|서멀|과전류|과부하|두꺼비집|쇼트|합선/ },
    { tag: '탄내·연기', re: /탄\s?냄새|타는\s?냄새|탄내|연기|스파크|소손/ },
    { tag: '회전 이상', re: /회전.{0,6}(멈|무거|안|이상)|안\s?돌|스크류.{0,4}(멈|무거)|벨트.{0,6}(멈|미끄|헛)/ },
    { tag: '토출·성형 불량', re: /미성형|덜\s?채|안\s?나와|안\s?나옴|잘\s?안\s?나|토출|압출량|흐름.{0,4}(고르지|불균일)|분사.{0,3}불균일|변형|편차|불균일|불량\s?(다수|발생|재발)/ },
    { tag: '파손·마모', re: /파손|부러|깨짐|깨졌|마모|닳|노후|부식|녹\s|녹이|풀림|헐거/ },
    { tag: '전원·배터리', re: /배터리|방전|전원.{0,3}(안|불량|꺼)/ }
  ];
  function tagList() { return TAGS.map(function (t) { return t.tag; }); }
  function extractTags(text) {
    var s = str(text).toLowerCase();
    if (!s) return [];
    return TAGS.filter(function (t) { return t.re.test(s); }).map(function (t) { return t.tag; });
  }
  function uniq(list) {
    var seen = {}, out = [];
    (list || []).forEach(function (x) { x = str(x); if (x && !seen[x]) { seen[x] = 1; out.push(x); } });
    return out;
  }

  /* ── 4. 원인 묶기 ────────────────────────────────────────────
     같은 원인을 다르게 적은 것(히터 끊김 · 히터밴드 단선)을 대표 이름으로 묶습니다.
     괄호 안 설명(「히터 단선 (2존 히터 밴드)」)은 묶을 때 빼고 봅니다. 순서가 중요합니다 — 「센서 단선」이 히터로 가지 않게 센서를 먼저 봅니다. */
  /* weak: true 는 「파손」 「노후」 같은 넓은 말 묶음 — 다른 묶음에 하나도 안 걸릴 때만 씁니다.
     2차(실제 수리이력 반영): 원인 칸이 따로 없고 「릴레이 접점 불량으로 공급 멈춤」 처럼 증상 칸에 원인이 섞여 있어,
     부위(부품) 낱말 묶음을 늘렸습니다. 순서가 중요합니다 — 서보 → 제어 → 센서 → 히터 → 냉각 → 유압 → 모터 과부하 → 감속기 → 베어링 → 모터 … */
  var CAUSE_GROUPS = [
    { name: '서보 드라이버·서보 이상', cats: ['전기'], part: '서보·구동', syn: ['서보드라이버', '서보드라이브', '서보모터', '서보', '출력보드', 'servo'] },
    { name: '인버터 이상', cats: ['전기'], part: '전기·제어', syn: ['인버터', 'inverter', 'vfd'] },
    { name: '제어 기판·계기 고장', cats: ['전기'], part: '전기·제어', syn: ['기판', '계기판', '프로그램', '컨트롤러', 'plc', '수신부', '리모컨', '전압계', '조절계', '온도계', '축에러', '에러발생', '디지털계기', '표시기'] },
    { name: '타이머 고장', cats: ['전기'], part: '전기·제어', syn: ['타이머', 'timer'] },
    { name: '릴레이 불량', cats: ['전기'], part: '전기·제어', syn: ['릴레이', 'relay', '마그네트', '전자접촉기'] },
    { name: '위치·동작 센서 불량', cats: ['전기'], part: '센서', syn: ['절단센서', '작동센서', '근접센서', '광센서', '리밋', '리미트', '감지센서', '위치센서', '에어센서'] },
    { name: '온도 센서 불량', cats: ['전기'], part: '센서', syn: ['온도센서', '센서불량', '센서단선', '센서접촉', '열전대', '써모커플', '서모커플', 'thermocouple', 'tc불량', 'tc단선', '측온'] },
    { name: '히터 단선', cats: ['전기', '고온'], part: '히터', syn: ['히터단선', '히터끊', '히터밴드단선', '히터소손', '히터불량', '히터고장', '히터선단선', '히터기', '히터연결', '히터선'] },
    { name: '냉각 계통 이상(냉매·누유)', cats: ['냉매', '회전체'], part: '냉각', syn: ['냉각가스', '가스라인', '냉각라인', '냉각기누유', '냉각불량', '프레온', '냉매', '콤프레서', '콤프레셔', '컴프레서', '컴프레샤', '압축기'] },
    { name: '유압 누유', cats: ['유압'], part: '유압', syn: ['유압누유', '유압유누유', '오일누유', '호스누유', '피팅누유', '실누유', '씰누유', '오링', '유압호스', '유압피팅', '누유'] },
    { name: '모터 과부하', cats: ['전기', '회전체'], part: '모터·구동', syn: ['모터과부하', '과부하', '모터과열', '오버로드', 'overload', '써멀트립', '서멀트립'] },
    { name: '감속기 고장', cats: ['회전체'], part: '모터·구동', syn: ['감속기', '기어박스', '기어'] },
    { name: '베어링·롤러 마모', cats: ['회전체'], part: '모터·구동', syn: ['베어링', '롤러', '구리스', '그리스'] },
    { name: '모터 고장·소손', cats: ['전기', '회전체'], part: '모터·구동', syn: ['모터쇼트', '모터소손', '모터고장', '모터이상', '모터미작동', '트랜스모터', '팬모터', '모터'] },
    { name: '벨트 마모·노후', cats: ['회전체'], part: '모터·구동', syn: ['벨트', 'belt'] },
    { name: '펌프 고장·누수', cats: ['전기', '회전체'], part: '펌프·배관', syn: ['펌프', '양수', '누수'] },
    { name: '스크류·바렐 손상', cats: ['고온', '회전체'], part: '스크류·바렐', syn: ['스크류', '스크루', '바렐', '배럴손상', '실린더헤드'] },
    { name: '호퍼·투입구 막힘', cats: ['고온'], part: '스크류·바렐', syn: ['호퍼', '투입구'] },
    { name: '금형 막힘', cats: ['고온'], part: '금형·노즐', syn: ['금형막힘', '노즐막힘', '게이트막힘', '다이막힘', '스크린막힘', '런너막힘', '막힘'] },
    { name: '호스·배관 노후(녹·이물)', cats: [], part: '펌프·배관', syn: ['물호스', '호스내부', '배관', '녹', '이물질'] },
    { name: '배터리 방전·수명', cats: ['전기'], part: '전기·제어', syn: ['배터리', '방전', 'battery'] },
    { name: '볼트 풀림·파손', cats: [], part: '구조·체결', syn: ['볼트', '체결', '너트'] },
    { name: '축·샤프트 파손', cats: ['회전체'], part: '구조·체결', syn: ['축파손', '샤프트', '밀어내는축'] },
    { name: '설정값(압력·조건) 이상', cats: [], part: '운전 조건', syn: ['압력수치', '압력조정', '조건조정', '설정값조정', '설정변경'] },
    { name: '작업 환경(고온·먼지)', cats: [], part: '운전 조건', syn: ['현장고온', '주변온도', '실내온도', '여름', '하절기', '더위', '에어컨', '냉풍기', '환기', '고온'] },
    { name: '소모 부품·구조물 파손', cats: [], part: '구조·체결', weak: true, syn: ['커버', '받침대', '합판', '상판', '걸이', '조절기', '파손', '마모', '노후', '부러', '깨짐'] }
  ];
  var UNKNOWN_CAUSE = '원인 미상';
  function causeCore(text) { return str(text).replace(/\s*[(（\[][^)）\]]*[)）\]]\s*/g, ' ').trim(); }
  function groupHit(core, g) { for (var j = 0; j < g.syn.length; j++) if (core.indexOf(g.syn[j]) >= 0) return g.syn[j]; return ''; }
  function canonCause(text) {
    var full = squash(text), core = squash(causeCore(text)) || full;
    if (!core) return '';
    for (var k = 0; k < CAUSE_GROUPS.length; k++) if (squash(CAUSE_GROUPS[k].name) === full) return CAUSE_GROUPS[k].name;
    var weak = null;
    for (var i = 0; i < CAUSE_GROUPS.length; i++) {
      var g = CAUSE_GROUPS[i];
      if (!groupHit(core, g)) continue;
      if (!g.weak) return g.name;
      if (!weak) weak = g;
    }
    if (weak) return weak.name;
    return causeCore(text) || str(text);
  }
  /* 원인 칸이 없을 때 — 증상·조치·설비 글에서 원인 묶음을 짐작합니다(「추정」 표시).
     ① 「A(으)로 B」 의 A(원인 쪽)를 먼저 보고 ② 증상 전체 ③ 조치 ④ 설비 이름 순.
     한 글에 여러 부위가 나오면 고장 낱말(소손·단선·고장·불량 …)에 가장 가까이 붙은 부위를 고릅니다
     (「타이머 교체 중 합선, 구동 모터 소손」 → 모터). 아무것도 안 걸리면 「원인 미상」. */
  var FAIL_RE = /소손|단선|파손|고장|불량|마모|누유|누수|쇼트|방전|풀림|부식|막힘|노후|이상|에러|알람|미작동|불가|저하|끊/g;
  function bestGroup(text) {
    var core = squash(text);
    if (!core) return null;
    var fails = [], m; FAIL_RE.lastIndex = 0;
    while ((m = FAIL_RE.exec(core))) fails.push(m.index);
    var best = null, bestD = Infinity, weak = null;
    CAUSE_GROUPS.forEach(function (g, gi) {
      var hit = groupHit(core, g); if (!hit) return;
      if (g.weak) { if (!weak) weak = g; return; }
      // 그 묶음의 모든 낱말·모든 자리 중 고장 낱말에 가장 가까운 거리
      var d = 999, pure = new RegExp('^(' + FAIL_RE.source + ')$');
      g.syn.forEach(function (w) {
        for (var at = core.indexOf(w); at >= 0; at = core.indexOf(w, at + 1)) {
          var end = at + w.length, dd = 999;
          fails.forEach(function (f) { if (f >= at && Math.max(0, f - end) < dd) dd = Math.max(0, f - end); });
          // 「누유」 처럼 부위 없이 고장 낱말만인 것은 부위가 붙은 것(「냉각가스 배관 … 누유」)보다 뒤로
          if (pure.test(w)) dd = 6;
          if (dd < d) d = dd;
        }
      });
      if (d <= 2) d = 0; // 두 글자 안쪽은 같은 거리로 보고 앞 묶음(더 구체적인 부위)을 고릅니다 — 「냉각가스 배관 누유」 → 냉각
      if (d < bestD) { best = g; bestD = d; }
    });
    return best || weak;
  }
  function causeClause(text) { var m = /^(.{2,}?)(?:으로|로|로 인해|때문에)\s/.exec(str(text)); return m ? m[1] : ''; }
  function inferCause(symptom, action, equipment) {
    var texts = [causeClause(symptom), symptom, action, equipment];
    for (var i = 0; i < texts.length; i++) {
      if (!str(texts[i])) continue;
      var g = bestGroup(texts[i]);
      if (g && !(g.weak && i === 3)) return g.name;
    }
    return UNKNOWN_CAUSE;
  }
  function isUnknownCause(name) { return !str(name) || /미상|모름|불명|확인\s?중/.test(str(name)); }
  /* 부위 분류 — 원인 묶음에 붙은 부위(전기·제어 · 센서 · 히터 · 냉각 · 유압 · 모터·구동 · 펌프·배관 · 스크류·바렐 · 금형·노즐 · 구조·체결 · 운전 조건) */
  function partOf(causeName) { var g = causeGroupOf(canonCause(causeName)); return g ? g.part : ''; }
  /* 설비 이름 → 설비 종류(가족). 「압출기 1호기」 「압출 1호기」 「압출기」 를 한 종류로 봅니다. 순서 중요(냉각기 → 압출기 …) */
  var EQUIP_FAMILIES = [
    ['환풍기', /환풍|환기팬|배기팬/], ['냉각기', /냉각|칠러|chiller/i], ['배합기', /배합|믹서|혼합기/], ['열처리기', /열처리|소둔/],
    ['건조기', /건조/], ['금형 온도 조절기', /온조|온도\s?조절기/], ['압출기', /압출/], ['사출기', /사출/], ['스프링기', /스프링/],
    ['선재공급기', /선재|공급기|피더/], ['분쇄기', /분쇄|파쇄/], ['리프트·호이스트', /리프트|호이스트|크레인|지게차/],
    ['펌프', /펌프|양수/], ['대차', /대차|카트/], ['컨베이어', /컨베이어/], ['절단기', /절단|커터|컷터/], ['프레스', /프레스/], ['공압기', /컴프레서|공압/]
  ];
  function equipFamily(name) {
    var s = str(name);
    if (!s) return '';
    for (var i = 0; i < EQUIP_FAMILIES.length; i++) if (EQUIP_FAMILIES[i][1].test(s)) return EQUIP_FAMILIES[i][0];
    return s.replace(/\s*\d+\s*호기.*$/, '').replace(/\s*[(（][^)）]*[)）]/g, '').trim() || s;
  }
  function equipUnit(name) { var m = /(\d+)\s*호기/.exec(str(name)); return m ? m[1] + '호기' : ''; }
  function causeKey(text) { return squash(canonCause(text)); }
  function causeGroupOf(name) {
    var k = squash(name);
    for (var i = 0; i < CAUSE_GROUPS.length; i++) if (squash(CAUSE_GROUPS[i].name) === k) return CAUSE_GROUPS[i];
    return null;
  }

  /* ── 5. 낱말 쪼개기와 유사도 ────────────────────────────────
     한국어는 조사가 붙어 「온도가」 「온도까지」 가 다른 낱말이 됩니다. 그래서
       ① 끝의 조사를 떼어 낸 낱말(w:) ② 그 낱말의 글자 2개 묶음(b:)
     을 함께 씁니다. 글자 2개 묶음은 「올라가고」 「올라감」 처럼 끝이 달라도 「올라」 로 겹칩니다. */
  var JOSA = ['에서는', '에서도', '으로는', '까지는', '에서', '으로', '까지', '부터', '하고', '이고', '하여', '해서', '되고', '했음', '했고', '함', '됨',
    '이', '가', '은', '는', '을', '를', '에', '도', '고', '로', '와', '과', '의', '서', '만'];
  function stem(w) {
    for (var i = 0; i < JOSA.length; i++) {
      var j = JOSA[i];
      if (w.length > j.length + 1 && w.slice(-j.length) === j) return w.slice(0, -j.length);
    }
    return w;
  }
  /* 같은 뜻 다른 말 — 묶음의 대표 낱말을 덧붙여 「작동 불가」 로 찾아도 「멈춤」 사례가 걸리게 합니다(사례·질문 양쪽에 똑같이) */
  var SYNONYMS = [
    ['멈춤', /멈춤|멈춰|멈췄|정지|작동\s?불가|미작동|구동\s?불가|작동\s?안|생산\s?불가|작업\s?중단|가동\s?중지|꺼짐|다운/],
    ['소음', /소음|소리|이음|굉음|덜컹|웅웅/],
    ['샘', /누유|누수|새어|새고|샌다|샘|기름이?\s?고/],
    ['고장', /고장|불량|이상|에러|알람|오류|안\s?됨/],
    ['파손', /파손|부러|깨짐|깨졌|부서/],
    ['마모', /마모|닳|노후|헐거|낡/],
    ['온도낮음', /온도.{0,8}(안\s?올라|안\s?오르|못\s?올라|낮|미달|상승하지|상승\s?불량)|승온/],
    ['과열', /과열|온도\s?상승으로|뜨거|고온/],
    ['냉각불량', /냉각.{0,4}(불량|안|약|이상)|냉매|프레온|누수/],
    ['전원', /배터리|방전|전원/]
  ];
  function expandText(text) {
    var t = str(text);
    if (!t) return '';
    var add = [];
    SYNONYMS.forEach(function (g) { if (g[1].test(t)) add.push(g[0]); });
    return add.length ? t + ' ' + add.join(' ') : t;
  }
  function words(text) {
    return str(text).toLowerCase().replace(/[^0-9a-z가-힣℃\s]/g, ' ').split(/\s+/).filter(Boolean);
  }
  function tokens(text) {
    var out = [];
    words(text).forEach(function (w0) {
      var w = stem(w0);
      if (w.length >= 2) out.push('w:' + w);
      if (/[가-힣]/.test(w)) for (var i = 0; i + 1 < w.length; i++) out.push('b:' + w.slice(i, i + 2));
    });
    return out;
  }
  function addTokens(tf, list, weight) {
    list.forEach(function (t) { tf[t] = (tf[t] || 0) + weight; });
  }
  /* 사례 한 건의 글: 증상(무게 1) + 태그(1) + 원인(0.5). 원인 낱말도 조금 넣어 「히터 쪽이 이상」 같은 글에도 걸리게 합니다 */
  function caseTf(c) {
    var tf = {};
    addTokens(tf, tokens(expandText(c.symptom)), 1);
    addTokens(tf, tokens((c.tags || []).join(' ')), 1);
    addTokens(tf, tokens(canonCause(c.cause)), 0.5);
    return tf;
  }
  function buildIndex(cases) {
    var df = {}, docs = [];
    (cases || []).forEach(function (c) {
      var tf = caseTf(c);
      Object.keys(tf).forEach(function (t) { df[t] = (df[t] || 0) + 1; });
      docs.push({ c: c, tf: tf });
    });
    var N = docs.length, idf = {};
    Object.keys(df).forEach(function (t) { idf[t] = Math.log((N + 1) / (df[t] + 1)) + 1; });
    docs.forEach(function (d) {
      var v = {}, n2 = 0;
      Object.keys(d.tf).forEach(function (t) { var w = (1 + Math.log(d.tf[t])) * idf[t]; v[t] = w; n2 += w * w; });
      d.vec = v; d.norm = Math.sqrt(n2);
    });
    return { N: N, idf: idf, docs: docs };
  }
  function queryVec(index, text) {
    var tf = {}; addTokens(tf, tokens(expandText(text)), 1);
    var v = {}, n2 = 0;
    Object.keys(tf).forEach(function (t) {
      if (!index.idf[t]) return; // 사례에 한 번도 없는 낱말은 비교에 쓰지 않습니다
      var w = (1 + Math.log(tf[t])) * index.idf[t]; v[t] = w; n2 += w * w;
    });
    return { vec: v, norm: Math.sqrt(n2) };
  }
  function cosine(a, an, b, bn) {
    if (!an || !bn) return 0;
    var s = 0;
    Object.keys(a).forEach(function (t) { if (b[t]) s += a[t] * b[t]; });
    return s / (an * bn);
  }
  function jaccard(a, b) {
    if (!a.length || !b.length) return 0;
    var A = {}, inter = 0, uni = {};
    a.forEach(function (x) { A[x] = 1; uni[x] = 1; });
    b.forEach(function (x) { if (A[x]) inter++; uni[x] = 1; });
    return inter / Object.keys(uni).length;
  }

  var DEFAULT_OPTS = { topN: 8, minScore: 0.08, tagWeight: 0.35, sameEquipBonus: 0.08, sameTypeBonus: 0.03, partBonus: 0.4 };
  function cleanOpts(o) {
    var r = {}; o = o || {};
    Object.keys(DEFAULT_OPTS).forEach(function (k) { var v = Number(o[k]); r[k] = o[k] != null && isFinite(v) ? v : DEFAULT_OPTS[k]; });
    r.topN = Math.max(1, Math.min(30, Math.round(r.topN)));
    return r;
  }
  /* 비슷한 과거 사례. query = { text, tags[], equipment, equipmentType }, equipTypes = { 설비이름: 종류 } */
  function similar(cases, query, opts, equipTypes) {
    var o = cleanOpts(opts);
    query = query || {};
    var index = buildIndex(cases);
    var q = queryVec(index, str(query.text) + ' ' + (query.tags || []).join(' '));
    var qTags = uniq(query.tags && query.tags.length ? query.tags : extractTags(query.text));
    // 2차: 증상 글에 부위 낱말(서보·타이머·벨트·냉각 …)이 있으면 그 원인 묶음 사례에 가산점
    var qCore = squash(query.text), qGroups = {};
    if (qCore) CAUSE_GROUPS.forEach(function (g) {
      if (g.weak) return;
      var hit = groupHit(qCore, g);
      if (hit && !new RegExp('^(' + FAIL_RE.source + ')$').test(hit)) qGroups[g.name] = 1;
    });
    var qEq = str(query.equipment), qType = str(query.equipmentType) || (equipTypes && equipTypes[qEq]) || (qEq ? equipFamily(qEq) : '');
    var out = [];
    index.docs.forEach(function (d) {
      var cos = cosine(q.vec, q.norm, d.vec, d.norm);
      var tagScore = jaccard(qTags, d.c.tags || []);
      var base = qTags.length ? cos * (1 - o.tagWeight) + tagScore * o.tagWeight : cos;
      if (base <= 0) return;
      var sameEq = !!qEq && d.c.equipment === qEq;
      var cType = (equipTypes && equipTypes[d.c.equipment]) || (d.c.equipment ? equipFamily(d.c.equipment) : '');
      var sameType = !sameEq && !!qType && cType === qType;
      var samePart = !!qGroups[canonCause(d.c.cause)];
      var score = base + (sameEq ? o.sameEquipBonus : 0) + (sameType ? o.sameTypeBonus : 0) + (samePart && base > 0.02 ? o.partBonus : 0);
      if (score < o.minScore) return;
      var hits = [];
      Object.keys(q.vec).forEach(function (t) { if (t.slice(0, 2) === 'w:' && d.vec[t]) hits.push(t.slice(2)); });
      out.push({ c: d.c, score: round(score, 4), cos: round(cos, 4), tagScore: round(tagScore, 4), sameEquipment: sameEq, sameType: sameType, samePart: samePart, hits: hits,
        sharedTags: qTags.filter(function (t) { return (d.c.tags || []).indexOf(t) >= 0; }) });
    });
    out.sort(function (a, b) { return b.score - a.score || String(b.c.date).localeCompare(String(a.c.date)); });
    return out.slice(0, o.topN);
  }

  /* ── 6. 단계 나누기: 안전 · 초보 가능 · 담당자·전문가 ─────────
     「먼저 확인할 것」 은 안전 단계(전원 차단·잔압 빼기)를 뺀 첫 확인 단계입니다.
     전기 측정·배선, 유압 라인 분리·교체, 고온부 분해, 회전체 분해는 초보자가 하지 않고 담당자·외부를 부르는 단계로 표시합니다. */
  var SAFETY_STEP = /전원\s?차단|loto|잠금|조작\s?금지|잔압|압력.{0,3}(빼|제거|해제)|설비\s?정지\s?후|정지·전원|정지\s?후\s?전원|비상\s?정지/i;
  var EXPERT_STEP = /조절계|전기기사|업체|a\/s|출장|타이머|릴레이|서보|기판|계기판|인버터|트랜스|콤프레|컴프레|냉매|프레온|가스\s?충전|감속기|모터.{0,3}(교체|수리)|가스\s?라인|담당|전기\s?교육|입회|전문가|측정|테스터|멀티미터|메거|절연|배선|결선|전류|전압|인버터|(히터|모터|베어링|호스|실|씰|펌프|커플링|열전대|센서).{0,4}교체|교체.{0,3}(히터|모터|베어링|호스)|분해|라인.{0,3}분리|호스.{0,3}분리|어큐뮬|릴리프|용접|스크류.{0,3}(인출|빼)/;
  /* 2차: 수리이력의 「조치」 칸에는 일 순서가 아니라 진행 기록(「A/S 방문 예정」 「교체 필요로 보고」)이 섞입니다.
     이런 줄은 「기록(note)」 으로 보고 「먼저 확인할 것」·가이드 확인 순서에서 뺍니다(사례 보기에는 그대로) */
  var NOTE_STEP = /예정|보고됨|로\s?보고|보고\s?(함|됨)?$|기재|기록|진행$|진행\s?중|요청|조율|확인\s?중|재의뢰|수배|방문|견적|품의|가능일|결과\s?확인|부터\s?생산|대응$|^이후/;
  function stepKind(step) {
    var s = str(step);
    if (SAFETY_STEP.test(s)) return 'safety';
    if (NOTE_STEP.test(s)) return 'note';
    if (EXPERT_STEP.test(s)) return 'expert';
    return 'basic';
  }
  /* 「존별 온도 표시 확인 — 2존만 낮음」 처럼 그때의 결과가 붙은 단계는 결과를 떼고 모읍니다 */
  function stepCore(step) { return str(step).replace(/\s+[—–-]\s[^()]*$/, ''); }
  function stepKey(step) { return squash(stepCore(step).replace(/\([^)]*\)/g, '')); }
  function firstCheck(c) {
    var st = splitSteps(c.steps);
    for (var i = 0; i < st.length; i++) { var k = stepKind(st[i]); if (k !== 'safety' && k !== 'note') return st[i]; }
    return '';
  }

  /* ── 7. 진단 — 비슷한 사례로 원인 순위 매기기 ───────────────
     원인마다 비슷한 사례들의 유사도를 더해 비율을 냅니다(투표). 확정 진단이 아니라 「과거 사례에서 이 증상일 때 무엇이었나」 입니다. */
  function diagnose(cases, query, opts, equipTypes) {
    var sims = similar(cases, query, opts, equipTypes);
    var groups = {}, order = [], total = 0;
    sims.forEach(function (s) {
      var k = causeKey(s.c.cause);
      if (!k || isUnknownCause(s.c.cause)) return; // 원인 미상 사례는 비슷한 사례 표에만 보이고 원인 순위에는 넣지 않습니다
      if (!groups[k]) { groups[k] = { key: k, cause: canonCause(s.c.cause), weight: 0, items: [] }; order.push(k); }
      groups[k].weight += s.score; groups[k].items.push(s); total += s.score;
    });
    var all = cases || [];
    var causes = order.map(function (k) {
      var g = groups[k];
      // 먼저 확인할 것 — 근거 사례들의 첫 확인 단계를 모아 자주 나온 순으로
      var fc = {}, fOrder = [];
      g.items.forEach(function (s) {
        var f = firstCheck(s.c); if (!f) return;
        var fk = stepKey(f);
        if (!fc[fk]) { fc[fk] = { text: stepCore(f), count: 0, weight: 0 }; fOrder.push(fk); }
        fc[fk].count++; fc[fk].weight += s.score;
      });
      var checks = fOrder.map(function (x) { return fc[x]; }).sort(function (a, b) { return b.count - a.count || b.weight - a.weight; }).slice(0, 3)
        .map(function (x) { return { text: x.text, count: x.count, kind: stepKind(x.text) }; });
      var same = all.filter(function (c) { return causeKey(c.cause) === k; });
      var ext = same.filter(function (c) { return c.repairBy === '외부'; });
      var extKnown = ext.filter(hasSpent), extCost = extKnown.reduce(function (s, c) { return s + spent(c); }, 0);
      return {
        key: k, cause: g.cause, share: total ? g.weight / total : 0, weight: round(g.weight, 4), count: g.items.length,
        best: g.items[0].score, caseIds: g.items.map(function (s) { return s.c.id; }), checks: checks,
        totalCases: same.length, extCount: ext.length, selfCount: same.filter(function (c) { return c.repairBy === '자체'; }).length, avgExtCost: extKnown.length ? Math.round(extCost / extKnown.length) : null,
        level: g.items.length >= 3 && g.items[0].score >= 0.3 ? 'strong' : g.items.length >= 2 ? 'medium' : 'weak'
      };
    }).sort(function (a, b) { return b.weight - a.weight; });
    return { sims: sims, causes: causes, tags: uniq(query && query.tags && query.tags.length ? query.tags : extractTags(query && query.text)) };
  }
  var LEVEL_TEXT = { strong: '근거 많음', medium: '근거 보통', weak: '근거 적음' };

  /* ── 8. 안전 문구(표준) ────────────────────────────────────── */
  var SAFETY = {
    common: [
      '작업 전 설비를 멈추고 주 전원을 내린 뒤, 차단기에 잠금장치와 「작업 중 — 조작 금지」 표지를 답니다(LOTO). 열쇠는 작업하는 본인이 가집니다.',
      '전원을 내린 뒤 실제로 꺼졌는지(표시등 꺼짐·조작 버튼 무반응) 확인하고 시작합니다.',
      '혼자 판단이 서지 않으면 멈추고 정비 담당자에게 먼저 묻습니다.'
    ],
    '전기': '배선·단자를 만지거나 테스터로 전압·저항·전류를 재는 일은 전기 교육을 받은 담당자만 합니다. 초보자는 겉모양(변색·탄 자국·빠진 커넥터)만 눈으로 확인합니다.',
    '고온': '히터·실린더·노즐·금형은 전원을 내린 뒤에도 한동안 뜨겁습니다. 내열 장갑·보안경을 쓰고, 식기 전에는 분해하지 않습니다. 퍼징처럼 히터를 켠 채 하는 일은 노즐 정면에 서지 않습니다(뜨거운 수지가 튐).',
    '유압': '유압 라인은 설비를 멈춰도 압력이 남아 있습니다. 잔압을 빼기 전에는 호스·피팅을 풀지 않고, 새는 곳을 맨손으로 더듬어 찾지 않습니다(고압 분사 부상).',
    '회전체': '모터·스크류·벨트 같은 회전부는 완전히 멈춘 것을 확인하고 덮개를 엽니다. 장갑·소매가 말려 들어가지 않게 합니다.',
    '냉매': '냉각기의 냉매(가스) 라인과 압축기에는 압력이 걸려 있습니다. 배관을 풀거나 가스를 채우는 일은 업체가 합니다. 초보자는 누유 자국·팬에 낀 이물·필터 먼지만 눈으로 봅니다.'
  };
  var CALL_EXPERT = {
    common: [
      '탄 냄새·연기·스파크가 나거나 차단기가 반복해서 떨어질 때 — 다시 켜지 말고 바로 부릅니다.',
      '아래 「확인 순서」 에서 「담당자·전문가」 표시가 붙은 단계에 이르렀을 때 — 초보자는 여기서 멈춥니다.',
      '확인 순서를 다 해도 원인이 보이지 않거나, 조치한 뒤 같은 증상이 다시 날 때.'
    ],
    '전기': '전압·저항 측정, 배선 연결·교체가 필요할 때.',
    '고온': '뜨거운 상태에서 분해해야 하거나 노즐·스크류를 빼야 할 때.',
    '유압': '호스·실·펌프를 교체하거나 유압 라인을 분리해야 할 때.',
    '회전체': '베어링·커플링에서 소음이 나거나 축을 분해해야 할 때.',
    '냉매': '냉매(가스) 충전·배관 교체·압축기 교체가 필요할 때.'
  };
  /* 원인 이름에 정해진 분류가 없으면 단계·원인 글에서 짐작합니다 */
  function hazardCats(causeName, texts) {
    var g = causeGroupOf(causeName);
    if (g) return g.cats.slice();
    var s = (str(causeName) + ' ' + (texts || []).join(' ')).toLowerCase();
    var cats = [];
    if (/전기|전원|배선|단자|히터|모터|센서|차단기|인버터|전류|전압/.test(s)) cats.push('전기');
    if (/히터|노즐|실린더|배럴|금형|다이|고온|뜨거/.test(s)) cats.push('고온');
    if (/유압|호스|피팅|오일|압력|펌프/.test(s)) cats.push('유압');
    if (/모터|스크류|베어링|벨트|회전|커플링|감속기/.test(s)) cats.push('회전체');
    if (/냉매|프레온|콤프레|컴프레|냉각\s?가스/.test(s)) cats.push('냉매');
    return cats;
  }

  /* ── 9. 자체 정비 가이드북 — 원인별 카드 초안 ─────────────── */
  var TOOL_WORDS = ['테스터', '멀티미터', '클램프미터', '육각렌치', '스패너', '몽키스패너', '토크렌치', '드라이버', '내열 장갑', '보안경', '손전등', '걸레', '에어건', '황동 브러시', '퍼징재', '오일 받이', '그리스', '흡착포'];
  function countList(items) {
    var m = {}, order = [];
    items.forEach(function (x) { var k = squash(x); if (!k) return; if (!m[k]) { m[k] = { name: str(x), count: 0 }; order.push(k); } m[k].count++; });
    return order.map(function (k) { return m[k]; }).sort(function (a, b) { return b.count - a.count; });
  }
  function causeList(cases) {
    var m = {}, order = [];
    (cases || []).forEach(function (c) {
      var k = causeKey(c.cause); if (!k) return;
      if (!m[k]) { m[k] = { key: k, cause: canonCause(c.cause), count: 0 }; order.push(k); }
      m[k].count++;
    });
    return order.map(function (k) { return m[k]; }).sort(function (a, b) { return b.count - a.count || a.cause.localeCompare(b.cause); });
  }
  function guideDraft(cases, key) {
    var list = (cases || []).filter(function (c) { return causeKey(c.cause) === key; });
    if (!list.length) return null;
    var name = canonCause(list[0].cause);
    // 단계 모으기 — 사례 안에서의 상대 위치 평균으로 순서를 정합니다
    var st = {}, sOrder = [], actions = [];
    list.forEach(function (c) {
      var s = splitSteps(c.steps).filter(function (x) { var k = stepKind(x); return k !== 'safety' && k !== 'note'; });
      if (s.length < 2) { s.forEach(function (x) { actions.push(stepCore(x)); }); return; } // 한 줄 조치는 순서가 없으니 「지난 조치 기록」 으로
      s.forEach(function (x, i) {
        var k = stepKey(x); if (!k) return;
        if (!st[k]) { st[k] = { text: stepCore(x), count: 0, pos: 0 }; sOrder.push(k); }
        st[k].count++; st[k].pos += s.length > 1 ? i / (s.length - 1) : 0;
      });
    });
    var steps = sOrder.map(function (k) { var x = st[k]; return { text: x.text, count: x.count, pos: x.pos / x.count, kind: stepKind(x.text) }; })
      .sort(function (a, b) { return a.pos - b.pos || b.count - a.count; });
    // 순서는 실제로 한 순서(위치 평균) 그대로 둡니다. 초보자가 멈출 곳은 화면에서 첫 「담당자·전문가」 단계 앞에 긋습니다(stopIndex)
    steps = steps.slice(0, 8);
    var allSteps = [];
    list.forEach(function (c) { allSteps = allSteps.concat(splitSteps(c.steps)); });
    var tools = countList(list.reduce(function (a, c) { return a.concat(splitList(c.tools)); }, []).concat(
      TOOL_WORDS.filter(function (t) { return allSteps.some(function (s) { return s.indexOf(t) >= 0; }); }))).map(function (x) { return x.name; });
    var cats = hazardCats(name, allSteps);
    if (cats.indexOf('고온') >= 0) { if (tools.indexOf('내열 장갑') < 0) tools.push('내열 장갑'); }
    if (tools.indexOf('보안경') < 0) tools.push('보안경');
    var ext = list.filter(function (c) { return c.repairBy === '외부'; }), self = list.filter(function (c) { return c.repairBy === '자체'; });
    var extKnown = ext.filter(hasSpent), extCost = extKnown.reduce(function (s, c) { return s + spent(c); }, 0);
    var extParts = countList(ext.reduce(function (a, c) { return a.concat(splitList(c.parts)); }, []));
    var callExpert = CALL_EXPERT.common.slice();
    cats.forEach(function (k) { if (CALL_EXPERT[k]) callExpert.push(CALL_EXPERT[k]); });
    if (ext.length) callExpert.push('과거 이 원인 ' + list.length + '건 중 외부 수리 ' + ext.length + '건' +
      (extParts.length ? ' — 주로 ' + extParts.slice(0, 2).map(function (p) { return p.name; }).join('·') + ' 교체 때 외부를 불렀습니다.' : '.'));
    var safety = SAFETY.common.slice();
    cats.forEach(function (k) { if (SAFETY[k]) safety.push(SAFETY[k]); });
    var symptoms = list.slice().sort(function (a, b) { return String(b.date).localeCompare(String(a.date)); })
      .map(function (c) { return str(c.symptom); }).filter(Boolean);
    return {
      key: key, cause: name, count: list.length, cats: cats,
      equipments: countList(list.map(function (c) { return c.equipment; })),
      tags: countList(list.reduce(function (a, c) { return a.concat(c.tags || []); }, [])).slice(0, 5),
      symptoms: uniq(symptoms).slice(0, 3),
      steps: steps,
      actions: countList(actions).map(function (x) { return { text: x.name, count: x.count, kind: stepKind(x.name) }; }).slice(0, 8),
      parts: countList(list.reduce(function (a, c) { return a.concat(splitList(c.parts)); }, [])),
      tools: uniq(tools),
      safety: safety, callExpert: callExpert,
      extCount: ext.length, selfCount: self.length, extCost: extCost, avgExtCost: extKnown.length ? Math.round(extCost / extKnown.length) : null,
      caseIds: list.map(function (c) { return c.id; })
    };
  }
  /* 화면에 보일 카드 = 초안 + 사용자가 고친 칸(edit). 고친 칸은 글 그대로(줄 = 항목) */
  var GUIDE_EDIT_FIELDS = ['symptoms', 'steps', 'tools', 'parts', 'safety', 'callExpert', 'note'];
  function linesOf(v) { return Array.isArray(v) ? v.map(str).filter(Boolean) : str(v).split(/\r?\n/).map(function (x) { return str(x).replace(/^[-*•·]\s*/, ''); }).filter(Boolean); }
  function guideCard(draft, edit) {
    if (!draft) return null;
    var e = edit || {};
    var card = {
      key: draft.key, cause: draft.cause, count: draft.count, cats: draft.cats, equipments: draft.equipments, tags: draft.tags,
      extCount: draft.extCount, selfCount: draft.selfCount, extCost: draft.extCost, avgExtCost: draft.avgExtCost, caseIds: draft.caseIds,
      symptoms: draft.symptoms, steps: draft.steps.map(function (s) { return { text: s.text, kind: s.kind, count: s.count }; }), actions: draft.actions || [],
      tools: draft.tools, parts: draft.parts.map(function (p) { return p.name; }), safety: draft.safety, callExpert: draft.callExpert, note: '',
      edited: [], updatedAt: e.updatedAt || ''
    };
    GUIDE_EDIT_FIELDS.forEach(function (f) {
      if (e[f] == null) return;
      card.edited.push(f);
      if (f === 'note') { card.note = str(e.note); return; }
      var lines = linesOf(e[f]);
      card[f] = f === 'steps' ? lines.map(function (t) { return { text: t, kind: stepKind(t), count: null }; }) : lines;
    });
    card.stopIndex = stopIndex(card.steps);
    // 안전 문구는 고쳐도 표준 LOTO 문장은 빠지지 않게 맨 앞에 남깁니다
    if (card.edited.indexOf('safety') >= 0 && !card.safety.some(function (s) { return /LOTO|잠금/.test(s); })) card.safety.unshift(SAFETY.common[0]);
    return card;
  }
  /* 초보자가 멈출 곳 — 첫 「담당자·전문가」 단계의 번호(없으면 -1) */
  function stopIndex(steps) { for (var i = 0; i < (steps || []).length; i++) if (steps[i].kind === 'expert') return i; return -1; }
  function guideToText(card) {
    if (!card) return '';
    var out = ['[자체 정비 가이드] ' + card.cause, ''];
    out.push('■ 이런 증상일 때'); card.symptoms.forEach(function (s) { out.push('- ' + s); });
    out.push('', '■ 안전 먼저'); card.safety.forEach(function (s) { out.push('- ' + s); });
    out.push('', '■ 확인 순서');
    card.steps.forEach(function (s, i) {
      if (i === card.stopIndex) out.push('--- 초보자는 여기까지. 다음 단계부터는 정비 담당자·외부 전문가가 합니다 ---');
      out.push((i + 1) + '. ' + s.text + (s.kind === 'expert' ? '  [담당자·전문가]' : ''));
    });
    if (!card.steps.length) out.push('(조치가 두 단계 이상 적힌 사례가 아직 없어 순서를 만들지 못했습니다)');
    if ((card.actions || []).length) { out.push('', '■ 지난 조치 기록(한 줄로 적힌 것)'); card.actions.forEach(function (a) { out.push('- ' + a.text + (a.count > 1 ? ' (' + a.count + '건)' : '') + (a.kind === 'expert' ? '  [담당자·전문가]' : '')); }); }
    out.push('', '■ 필요 공구: ' + (card.tools.join(', ') || '-'));
    out.push('■ 자주 바꾼 부품: ' + (card.parts.join(', ') || '-'));
    out.push('', '■ 이럴 때는 멈추고 전문가를 부릅니다'); card.callExpert.forEach(function (s) { out.push('- ' + s); });
    if (card.note) out.push('', '■ 메모', card.note);
    out.push('', '※ 과거 정비 사례 ' + card.count + '건에서 만든 초안입니다. 설비 매뉴얼과 정비 담당자의 확인을 받아 쓰세요.');
    return out.join('\n');
  }

  /* ── 10. 비용 현황 ─────────────────────────────────────────── */
  function inPeriod(c, from, to) { var m = monthOf(c.date); return (!from || m >= from) && (!to || m <= to); }
  function monthRange(from, to) {
    var out = [];
    if (!from || !to || from > to) return out;
    var y = Number(from.slice(0, 4)), m = Number(from.slice(5, 7));
    for (var i = 0; i < 240; i++) {
      var k = y + '-' + pad(m); out.push(k);
      if (k >= to) break;
      m++; if (m > 12) { m = 1; y++; }
    }
    return out;
  }
  /* 금액이 실제 지출로 적힌 사례(금액·무상·개당 단가). 미기재·견적·못 읽음은 빠집니다 */
  function hasSpent(c) { return c && c.cost != null && (c.costKind === 'won' || c.costKind === 'free' || c.costKind === 'unit' || !c.costKind); }
  /* skip = { 사례 id: true } — 중복 의심(같은 설비·같은 금액이 이어 보고된 건)의 뒤 건 금액을 한 번만 세기 위함 */
  function stat(list, skip) {
    skip = skip || {};
    function sum(arr) { return arr.reduce(function (s, c) { return s + (skip[c.id] ? 0 : spent(c)); }, 0); }
    var ext = list.filter(function (c) { return c.repairBy === '외부'; });
    var self = list.filter(function (c) { return c.repairBy === '자체'; });
    var unk = list.filter(function (c) { return c.repairBy !== '외부' && c.repairBy !== '자체'; });
    var known = ext.length + self.length;
    return {
      total: list.length, extCount: ext.length, selfCount: self.length, unknownCount: unk.length,
      extCost: sum(ext), selfCost: sum(self), unknownCost: sum(unk),
      allCost: sum(list),
      costKnown: list.filter(hasSpent).length,
      estCost: list.reduce(function (s, c) { return s + (c.costKind === 'estimate' && c.cost ? Number(c.cost) : 0); }, 0),
      estCount: list.filter(function (c) { return c.costKind === 'estimate'; }).length,
      pending: list.filter(function (c) { return c.status === '진행 중'; }).length,
      hours: round(list.reduce(function (s, c) { return s + (Number(c.hours) || 0); }, 0), 1),
      selfRatio: known ? self.length / known : null
    };
  }
  /* 중복·연속 보고 의심 — 같은 설비(종류+호기)에서 45일 안에 같은 원인 묶음이거나 증상 글이 많이 겹치는 두 건.
     같은 금액(0 보다 큼)까지 같으면 「비용이 두 번 셈될 수 있음」(sameCost). 뒤 건이 later */
  function bigramSet(t) { var s = squash(t), o = {}; for (var i = 0; i + 1 < s.length; i++) o[s.slice(i, i + 2)] = 1; return o; }
  function overlap(a, b) { var ka = Object.keys(a), kb = Object.keys(b); if (!ka.length || !kb.length) return 0; var n = 0; ka.forEach(function (k) { if (b[k]) n++; }); return n / Math.min(ka.length, kb.length); }
  function dayNo(date) { var t = Date.parse(date + 'T00:00:00Z'); return isFinite(t) ? Math.round(t / 86400000) : null; }
  function dupSuspects(cases) {
    var list = (cases || []).filter(function (c) { return c.date && c.equipment; }).slice()
      .sort(function (a, b) { return String(a.date).localeCompare(String(b.date)); });
    var out = [];
    for (var i = 0; i < list.length; i++) {
      for (var j = i + 1; j < list.length; j++) {
        var a = list[i], b = list[j];
        var gap = dayNo(b.date) - dayNo(a.date);
        if (gap > 45) break;
        if (equipFamily(a.equipment) !== equipFamily(b.equipment) || equipUnit(a.equipment) !== equipUnit(b.equipment)) continue;
        var sameCause = !isUnknownCause(a.cause) && causeKey(a.cause) === causeKey(b.cause);
        var sim = overlap(bigramSet(a.symptom), bigramSet(b.symptom));
        if (!sameCause && sim < 0.6) continue;
        var sameCost = a.cost != null && a.cost > 0 && a.cost === b.cost && a.costKind === b.costKind;
        out.push({ first: a.id, later: b.id, gap: gap, sameCause: sameCause, sim: round(sim, 2), sameCost: sameCost, cost: sameCost ? a.cost : null });
      }
    }
    return out;
  }
  function costReport(cases, from, to, opts) {
    opts = opts || {};
    var dated = (cases || []).filter(function (c) { return c.date; });
    var months = dated.map(function (c) { return monthOf(c.date); }).sort();
    var f = from || months[0] || '', t = to || months[months.length - 1] || '';
    var list = dated.filter(function (c) { return inPeriod(c, f, t); });
    var dups = dupSuspects(list), skip = {};
    if (opts.dedupe !== false) dups.forEach(function (d) { if (d.sameCost) skip[d.later] = true; });
    var monthly = monthRange(f, t).map(function (m) {
      var s = stat(list.filter(function (c) { return monthOf(c.date) === m; }), skip); s.month = m; return s;
    });
    function by(getKey, getLabel) {
      var g = {}, order = [];
      list.forEach(function (c) { var k = getKey(c); if (!k) return; if (!g[k]) { g[k] = { label: getLabel(c), items: [] }; order.push(k); } g[k].items.push(c); });
      return order.map(function (k) { var s = stat(g[k].items, skip); s.key = k; s.label = g[k].label; return s; })
        .sort(function (a, b) { return b.allCost - a.allCost || b.extCost - a.extCost || b.total - a.total; });
    }
    var byEquipment = by(function (c) { return str(c.equipment) || '(설비 미기재)'; }, function (c) { return str(c.equipment) || '(설비 미기재)'; });
    var byFamily = by(function (c) { return equipFamily(c.equipment) || '(설비 미기재)'; }, function (c) { return equipFamily(c.equipment) || '(설비 미기재)'; });
    var byCause = by(function (c) { return causeKey(c.cause); }, function (c) { return canonCause(c.cause); });
    var byPart = by(function (c) { return partOf(c.cause) || '(부위 모름)'; }, function (c) { return partOf(c.cause) || '(부위 모름)'; });
    var byVendor = by(function (c) { return c.repairBy === '자체' ? '(자체)' : str(c.vendor) || '(업체 미기재)'; }, function (c) { return c.repairBy === '자체' ? '(자체)' : str(c.vendor) || '(업체 미기재)'; });
    // 자체 전환 후보 — 자체로 고친 적이 있는데 외부도 부른 원인. 외부 비용이 큰 순
    var candidates = byCause.filter(function (r) { return r.selfCount > 0 && r.extCount > 0; }).sort(function (a, b) { return b.extCost - a.extCost; });
    // 앞·뒤 절반 비교 — 자체 처리 비율이 늘었는지
    var half = Math.floor(monthly.length / 2), early = null, late = null;
    if (monthly.length >= 2) {
      early = stat(list.filter(function (c) { return monthOf(c.date) < monthly[half].month; }), skip);
      late = stat(list.filter(function (c) { return monthOf(c.date) >= monthly[half].month; }), skip);
    }
    var dupCost = dups.reduce(function (s, d) { return s + (d.sameCost ? d.cost : 0); }, 0);
    return { from: f, to: t, total: stat(list, skip), monthly: monthly, byEquipment: byEquipment, byFamily: byFamily, byCause: byCause, byPart: byPart, byVendor: byVendor,
      candidates: candidates, early: early, late: late, splitMonth: monthly.length >= 2 ? monthly[half].month : '', undated: (cases || []).length - dated.length,
      dups: dups, dupCost: dupCost, deduped: opts.dedupe !== false };
  }

  /* ── 11. AI 프롬프트 (반자동: 복사 → 회사가 허용한 AI → 답 붙여 넣기) ─── */
  function caseLine(c, i) {
    var st = splitSteps(c.steps);
    return '[사례 ' + (i + 1) + '] ' + (c.date || '날짜 없음') + ' · ' + (c.equipment || '설비 미기재') + '\n' +
      '  증상: ' + (str(c.symptom) || '-') + ((c.tags || []).length ? ' (태그: ' + c.tags.join(', ') + ')' : '') + '\n' +
      '  원인: ' + (str(c.cause) || '-') + '\n' +
      '  조치: ' + (st.length ? st.map(function (s, j) { return (j + 1) + ') ' + s; }).join(' ') : '-') + '\n' +
      '  교체 부품: ' + (splitList(c.parts).join(', ') || '없음') + ' · 수리: ' + (c.repairBy || '-') +
      (c.repairBy === '외부' && c.cost ? ' (' + fmtWon(c.cost) + ')' : '');
  }
  function buildDiagnosePrompt(input, result) {
    input = input || {}; result = result || { sims: [], causes: [] };
    var eq = input.equipment ? input.equipment + (input.equipmentType ? ' (' + input.equipmentType + ')' : '') : '지정 안 함';
    var L = [];
    L.push('부자재생산팀 설비 고장 원인을 함께 짚어 줘. 아래 「현재 증상」 과 「과거 정비 사례」, 「매뉴얼 발췌」 만 근거로 답해줘.');
    L.push('읽는 사람은 근속이 짧은 현장 작업자야. 쉬운 말로, 안전을 먼저 챙겨줘.');
    L.push('');
    L.push('## 현재 증상');
    L.push('- 설비: ' + eq);
    L.push('- 증상: ' + (str(input.text) || '-'));
    if ((input.tags || []).length) L.push('- 증상 태그: ' + input.tags.join(', '));
    L.push('');
    L.push('## 도구가 과거 사례로 계산한 가능 원인 (참고용 추정)');
    if (!result.causes.length) L.push('- 비슷한 과거 사례가 없어 계산하지 못했어.');
    result.causes.slice(0, 5).forEach(function (c, i) {
      L.push((i + 1) + '. ' + c.cause + ' — 비슷한 사례 ' + c.count + '건, 비중 ' + fmtPct(c.share) + (c.checks.length ? ', 그때 먼저 본 것: ' + c.checks.map(function (x) { return x.text; }).join(' / ') : ''));
    });
    L.push('');
    L.push('## 비슷한 과거 정비 사례 (유사도 순)');
    if (!result.sims.length) L.push('- 없음');
    result.sims.forEach(function (s, i) { L.push(caseLine(s.c, i)); });
    L.push('');
    L.push('## 매뉴얼 발췌');
    L.push(str(input.manual) || '- 없음');
    L.push('');
    L.push('## 답하는 방법');
    L.push('1. 가능 원인 상위 3가지 — 원인마다 근거가 된 사례 번호([사례 N])를 붙여줘. 사례·매뉴얼에 없는 내용은 끝에 「(추정)」 이라고 써줘.');
    L.push('2. 먼저 확인할 것 — 초보자가 전원 차단·잠금(LOTO) 뒤 눈으로 보거나 손으로 안전하게 할 수 있는 것만, 순서대로 3~5개.');
    L.push('3. 여기서 멈추고 전문가를 부를 때 — 전기 측정·배선, 유압 라인 분리·교체, 뜨거운 부위 분해는 초보자에게 시키지 말고 이 항목에 넣어줘.');
    L.push('4. 원인을 좁히려면 현장에서 더 알아볼 것 — 질문 2~3개.');
    L.push('확정 진단처럼 단정하지 말고, 「과거 사례로 볼 때 가능성이 높다」 는 말투로 써줘.');
    return L.join('\n');
  }
  function buildGuidePrompt(card, cases) {
    var L = [];
    L.push('「' + card.cause + '」 고장에 대한 초보자용 자체 정비 가이드 초안을 다듬어줘. 아래 초안과 과거 정비 사례만 근거로 써줘.');
    L.push('읽는 사람은 근속이 짧은 현장 작업자야. 한 단계에 한 가지 일만, 쉬운 말로 써줘.');
    L.push('');
    L.push('## 지금 초안');
    L.push(guideToText(card));
    L.push('');
    L.push('## 근거 사례');
    (cases || []).forEach(function (c, i) { L.push(caseLine(c, i)); });
    L.push('');
    L.push('## 지켜야 할 것');
    L.push('- 맨 앞은 늘 「전원 차단·잠금(LOTO)」 이어야 해.');
    L.push('- 전기 측정·배선, 유압 라인 분리·교체, 뜨거운 부위 분해는 「담당자·전문가」 단계로 따로 표시하고, 초보자는 그 앞에서 멈추게 써줘.');
    L.push('- 사례에 없는 내용은 「(추정)」 이라고 표시해줘.');
    L.push('- 답은 「■ 이런 증상일 때 / ■ 안전 먼저 / ■ 확인 순서 / ■ 필요 공구 / ■ 이럴 때는 멈추고 전문가를 부릅니다」 머리를 그대로 써줘.');
    return L.join('\n');
  }
  /* AI 가 같은 머리(■ …)로 답하면 카드 칸으로 나눠 넣습니다 */
  function parseGuideAnswer(text) {
    var map = { '이런 증상일 때': 'symptoms', '안전 먼저': 'safety', '확인 순서': 'steps', '필요 공구': 'tools', '이럴 때는 멈추고 전문가를 부릅니다': 'callExpert', '자주 바꾼 부품': 'parts', '메모': 'note' };
    var out = {}, cur = null;
    str(text).split(/\r?\n/).forEach(function (line) {
      var m = /^\s*(?:■|#+|\*\*)\s*([^:：*]+?)\s*(?:\*\*)?\s*(?:[:：]\s*(.*))?$/.exec(line);
      if (m) {
        var hit = Object.keys(map).filter(function (k) { return squash(m[1]).indexOf(squash(k)) === 0; })[0];
        if (hit) { cur = map[hit]; out[cur] = out[cur] || []; if (m[2]) out[cur] = out[cur].concat(cur === 'tools' || cur === 'parts' ? splitList(m[2]) : [m[2]]); return; }
        if (/^\s*■/.test(line)) { cur = null; return; } // 모르는 머리(「지난 조치 기록」 등)는 칸에 넣지 않습니다
      }
      if (/^\(조치가 두 단계/.test(str(line))) return;
      if (!cur) return;
      if (/^-{2,}.*초보자는 여기까지/.test(str(line))) return;
      var t = str(line).replace(/^(\d{1,2}[.)]|[-*•·])\s*/, '').replace(/\s*\[담당자·전문가[^\]]*\]\s*$/, '');
      if (t) out[cur].push(t);
    });
    Object.keys(out).forEach(function (k) { out[k] = k === 'note' ? out[k].join('\n') : out[k].join('\n'); if (!out[k]) delete out[k]; });
    return out;
  }

  /* ── 12. 정비 사례 한 건 ───────────────────────────────────── */
  var STATUS = ['완료', '진행 중'];
  function cleanCase(p, seq) {
    p = p || {};
    var d = p.datePrec && p.date && toISO(p.date) ? { date: toISO(p.date), prec: p.datePrec } : looseDate(p.date, p.year);
    var cost = p.costKind && COST_KIND_TEXT[p.costKind] ? { value: toNum(p.cost), kind: p.costKind } : parseCost(p.cost);
    var c = {
      id: str(p.id) || ('c' + (seq != null ? seq : Date.now().toString(36))),
      date: d.date,
      datePrec: d.date ? d.prec || 'day' : '',
      equipment: str(p.equipment),
      symptom: str(p.symptom),
      tags: uniq(Array.isArray(p.tags) ? p.tags : splitList(p.tags)),
      cause: str(p.cause),
      causeInferred: !!p.causeInferred,
      steps: splitSteps(p.steps),
      parts: splitList(p.parts),
      tools: splitList(p.tools),
      repairBy: REPAIR.indexOf(str(p.repairBy)) >= 0 ? str(p.repairBy) : normRepairBy(p.repairBy),
      vendor: str(p.vendor),
      cost: cost.value,
      costKind: cost.kind,
      hours: toHours(p.hours),
      status: STATUS.indexOf(str(p.status)) >= 0 ? str(p.status) : caseStatus(p.status),
      media: str(p.media),
      memo: str(p.memo),
      ref: str(p.ref),
      photos: Array.isArray(p.photos) ? p.photos.filter(function (x) { return x && x.id; }) : [],
      source: str(p.source) || '직접 입력',
      createdAt: str(p.createdAt)
    };
    if (c.cost != null && c.cost < 0) { c.cost = null; c.costKind = 'unread'; }
    if (c.cost == null && c.costKind === 'won') c.costKind = 'missing';
    return c;
  }
  function validateCase(c) {
    var e = {};
    if (!str(c.symptom)) e.symptom = '증상을 적어 주세요.';
    if (!str(c.cause)) e.cause = '원인을 적어 주세요. 모르면 「원인 미상」 이라고 적어 두고 나중에 고칩니다.';
    if (c.date && !looseDate(c.date).date) e.date = '날짜를 읽지 못했습니다.';
    return e;
  }
  function dupKey(c) { return [c.date, squash(c.equipment), squash(c.symptom), squash(c.cause)].join('|'); }

  /* ── 13. 정비 내역서 엑셀 가져오기 — 열 맞추기 ────────────── */
  /* 2차: 실제 수리이력 양식(연도 · 설비명 · 일자 · 증상 · 처리일자 · 수리업체 · 비용 · 비고 · 출처)을 그대로 읽도록 처리일자 · 수리 업체 · 처리 상태 · 연도 · 출처 칸을 더했습니다.
     원인 칸은 이제 꼭 필요하지 않습니다 — 없으면 증상·조치에서 짐작하고 「추정」 으로 표시합니다. */
  var FIELDS = [
    { id: 'date', label: '일자', req: false, aliases: ['일자', '날짜', '수리일', '수리일자', '정비일', '정비일자', '발생일', '발생일자', '작업일', '작업일자', '고장일', '고장일자', '발생시기', '고장시기'] },
    { id: 'equipment', label: '설비', req: false, aliases: ['설비', '설비명', '장비', '장비명', '기계', '기계명', '호기', '설비번호', '대상설비'] },
    { id: 'symptom', label: '증상', req: true, aliases: ['증상', '고장내용', '고장현상', '현상', '이상내용', '이상현상', '의뢰내용', '고장증상', '불량내용', '고장사항'] },
    { id: 'tags', label: '증상 태그', req: false, aliases: ['증상태그', '태그', '증상분류', '고장분류', '고장유형'] },
    { id: 'cause', label: '원인', req: false, aliases: ['원인', '고장원인', '점검결과', '점검결과원인', '진단', '진단결과', '발생원인'] },
    { id: 'steps', label: '조치 단계', req: false, aliases: ['조치', '조치내용', '조치사항', '조치순서', '수리내용', '작업내용', '처리내용', '정비내용', '수리내역'] },
    { id: 'parts', label: '교체 부품', req: false, aliases: ['교체부품', '부품', '사용부품', '자재', '사용자재', '교체품'] },
    { id: 'tools', label: '필요 공구', req: false, aliases: ['공구', '사용공구', '필요공구'] },
    { id: 'repairBy', label: '외부/자체', req: false, aliases: ['외부자체', '수리구분', '구분', '수리주체', '처리구분', '정비구분', '외주여부'] },
    { id: 'vendor', label: '수리 업체', req: false, aliases: ['수리업체', '업체', '업체명', '외주업체', '협력업체', '수리처', '거래처', '정비업체', '시공업체'] },
    { id: 'cost', label: '비용(원)', req: false, aliases: ['비용', '금액', '수리비', '수리비용', '청구금액', '청구금액원', '합계', '합계금액', '견적', '비용원'] },
    { id: 'costKind', label: '비용 구분(견적/실지출)', req: false, aliases: ['견적실지출', '비용구분', '견적여부', '지출구분', '금액구분'] },
    { id: 'hours', label: '소요시간(시간)', req: false, aliases: ['소요시간', '작업시간', '수리시간', '정비시간', '시간', '소요시간h', '작업시간h', '멈춘시간', '정지시간', '가동중지시간'] },
    { id: 'done', label: '처리일자·처리 내용', req: false, aliases: ['처리일자', '처리일', '완료일', '완료일자', '조치일', '조치일자', '처리결과', '처리현황', '진행상황', '처리'] },
    { id: 'status', label: '처리 상태', req: false, aliases: ['처리상태', '상태', '진행상태', '완료여부', '진행여부'] },
    { id: 'media', label: '사진·영상 위치', req: false, aliases: ['사진', '영상', '동영상', '첨부', '사진위치', '사진파일명', '파일'] },
    { id: 'memo', label: '비고', req: false, aliases: ['비고', '메모', '특이사항', '참고'] },
    { id: 'year', label: '연도', req: false, aliases: ['연도', '년도', '해당연도'] },
    { id: 'ref', label: '출처(보고서)', req: false, aliases: ['출처', '근거', '보고서', '근거자료', '출처보고서'] }
  ];
  function matchField(header) {
    var h = squash(header).replace(/\d+$/, '');
    if (!h) return '';
    for (var i = 0; i < FIELDS.length; i++) if (FIELDS[i].aliases.some(function (a) { return a === h; })) return FIELDS[i].id;
    for (i = 0; i < FIELDS.length; i++) if (FIELDS[i].aliases.some(function (a) { return a.length >= 2 && h.indexOf(a) === 0; })) return FIELDS[i].id;
    for (i = 0; i < FIELDS.length; i++) if (FIELDS[i].aliases.some(function (a) { return a.length >= 3 && h.indexOf(a) >= 0; })) return FIELDS[i].id;
    return '';
  }
  /* 머리행 찾기 — 위 15줄 안에서 별칭에 맞는 칸이 가장 많은 줄(2칸 이상). 제목 줄이 있어도 됩니다 */
  function detectHeaderRow(aoa) {
    var best = -1, bestN = 1;
    for (var r = 0; r < Math.min(15, (aoa || []).length); r++) {
      var row = aoa[r] || [], n = 0, seen = {};
      row.forEach(function (v) { var f = matchField(v); if (f && !seen[f]) { seen[f] = 1; n++; } });
      if (n > bestN) { best = r; bestN = n; }
    }
    return best < 0 ? 0 : best;
  }
  function headersOf(aoa, hr) {
    var row = (aoa && aoa[hr]) || [], w = 0;
    (aoa || []).slice(hr, hr + 30).forEach(function (r) { if (r && r.length > w) w = r.length; });
    var out = [];
    for (var i = 0; i < w; i++) out.push(str(row[i]) || ('(' + colName(i) + '열)'));
    return out;
  }
  function colName(i) { var s = ''; i++; while (i > 0) { var m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; }
  /* { 항목 id: 열 번호 }. saved 는 지난번 연결({ 항목 id: 머리 글 }) */
  function autoMap(headers, saved) {
    var map = {}, used = {};
    if (saved) Object.keys(saved).forEach(function (f) {
      var i = headers.map(squash).indexOf(squash(saved[f]));
      if (i >= 0 && !used[i]) { map[f] = i; used[i] = 1; }
    });
    headers.forEach(function (h, i) {
      if (used[i]) return;
      var f = matchField(h);
      if (f && map[f] == null) { map[f] = i; used[i] = 1; }
    });
    return map;
  }
  function mappingToSaved(headers, map) { var s = {}; Object.keys(map).forEach(function (f) { if (map[f] != null && headers[map[f]] != null) s[f] = headers[map[f]]; }); return s; }
  function missingRequired(map) { return FIELDS.filter(function (f) { return f.req && map[f.id] == null; }).map(function (f) { return f.label; }); }
  /* 병합 셀 풀기 — 엑셀에서 세로로 합친 칸(같은 설비 여러 줄 등)은 왼쪽 위 칸에만 값이 있습니다. 합친 범위 전체에 그 값을 채웁니다.
     merges = SheetJS 의 ws['!merges'] ([{ s:{r,c}, e:{r,c} }]). 원본 aoa 는 건드리지 않습니다 */
  function unmerge(aoa, merges) {
    var out = (aoa || []).map(function (r) { return (r || []).slice(); });
    (merges || []).forEach(function (m) {
      if (!m || !m.s || !m.e) return;
      var v = out[m.s.r] ? out[m.s.r][m.s.c] : null;
      if (v == null || v === '') return;
      for (var r = m.s.r; r <= m.e.r; r++) {
        if (!out[r]) out[r] = [];
        for (var c = m.s.c; c <= m.e.c; c++) if (out[r][c] == null || out[r][c] === '') out[r][c] = v;
      }
    });
    return out;
  }
  /* 같은 머리 모양의 시트 묶기 — 2025·2026 처럼 시트가 나뉘어 있어도 머리행이 같은 항목으로 맞춰지면 함께 가져옵니다.
     sheets = [{ name, aoa }] → [{ name, hr, headers, map, n }] 중 증상 칸이 잡힌 것만 */
  function sheetPlans(sheets, saved) {
    return (sheets || []).map(function (sh) {
      var hr = detectHeaderRow(sh.aoa), headers = headersOf(sh.aoa, hr), map = autoMap(headers, saved);
      return { name: sh.name, hr: hr, headers: headers, map: map, n: Object.keys(map).length, rows: Math.max(0, (sh.aoa || []).length - hr - 1) };
    }).filter(function (x) { return x.map.symptom != null && x.n >= 3; });
  }
  /* 행 → 사례. 증상·원인이 모두 빈 줄은 건너뜁니다(합계 줄·빈 줄).
     2차: 조치 칸이 없으면 처리일자 칸의 조치 글(날짜를 뗀 것) → 비고 순으로 조치로 읽고, 원인 칸이 없거나 비면 증상·조치에서 짐작합니다(추정 표시).
     외부/자체 칸이 없으면 수리 업체 칸으로(이름 → 외부, 「자체 조치」 → 자체, 「미기재」 → 모름). 일자가 없으면 처리일자에서. */
  function applyMapping(aoa, hr, map, opts) {
    opts = opts || {};
    var rows = [], issues = [], seq = opts.seqStart || 1, sheet = opts.sheet ? opts.sheet + ' ' : '';
    for (var r = hr + 1; r < (aoa || []).length; r++) {
      var row = aoa[r] || [];
      var get = function (f) { return map[f] == null ? null : row[map[f]]; };
      var any = row.some(function (v) { return str(v) !== ''; });
      if (!any) continue;
      var symptom = str(get('symptom')), cause = str(get('cause'));
      if (!symptom && !cause) { issues.push({ row: r + 1, sheet: opts.sheet, msg: sheet + '증상·원인이 비어 건너뜀' }); continue; }
      var rawDate = get('date'), rawCost = get('cost'), rawHours = get('hours'), rawBy = get('repairBy'), rawDone = get('done'), year = get('year');
      var doneAction = actionOf(rawDone), memo = str(get('memo'));
      var steps = get('steps'), stepsFrom = 'steps';
      if (!str(steps)) {
        if (doneAction && !/^(완료|교체\s?완료|수리\s?완료)$/.test(doneAction)) { steps = doneAction; stepsFrom = 'done'; }
        else if (memo && map.steps == null) { steps = memo; stepsFrom = 'memo'; }
        else stepsFrom = '';
      }
      var vend = parseVendor(get('vendor'));
      var by = str(rawBy) ? rawBy : vend.by;
      if (!str(rawBy) && !by) {
        var hint = [str(rawDone), memo].join(' ');
        if (/출장|외부\s?업체|a\/s|전기기사|외주|업체\s?(방문|이동)|전문\s?수리/i.test(hint)) by = '외부';
        else if (/자체\s?(교체|수리|조치|처리)|팀\s?(교체|수리)/.test(hint)) by = '자체';
      }
      var d = looseDate(rawDate, year);
      if (!d.date && str(rawDone)) d = looseDate(rawDone, year);
      var inferred = false;
      if (!cause || isUnknownCause(cause) && !str(get('cause'))) {
        cause = inferCause(symptom, str(steps) || [doneAction, memo].join(' '), str(get('equipment')));
        inferred = true;
      }
      var c = cleanCase({
        id: 'c' + (opts.idPrefix || 'x') + (seq++),
        date: d.date, datePrec: d.prec, equipment: get('equipment'), symptom: symptom,
        tags: uniq(splitList(get('tags')).concat(extractTags(symptom))),
        cause: cause, causeInferred: inferred, steps: steps, parts: get('parts'), tools: get('tools'),
        repairBy: by, vendor: vend.vendor, cost: rawCost, hours: rawHours,
        status: str(get('status')) || [str(rawDone), memo, stepsFrom === 'steps' ? str(steps) : ''].join(' '),
        media: get('media'), memo: stepsFrom === 'memo' ? '' : memo, ref: get('ref'), source: opts.source || '엑셀'
      });
      if (stepsFrom === 'memo') c.memo = '';
      var kindTxt = str(get('costKind'));
      if (kindTxt && c.cost != null) {
        if (/견적/.test(kindTxt)) c.costKind = 'estimate';
        else if (/무상/.test(kindTxt)) { c.costKind = 'free'; c.cost = 0; }
        else if (/개당|단가/.test(kindTxt)) c.costKind = 'unit';
        else if (/실|지출|확정/.test(kindTxt) && c.costKind === 'estimate') c.costKind = 'won';
      }
      if (str(rawDate) && !c.date) issues.push({ row: r + 1, sheet: opts.sheet, msg: sheet + '날짜를 읽지 못함: ' + str(rawDate) });
      if (str(rawCost) && c.costKind === 'unread') issues.push({ row: r + 1, sheet: opts.sheet, msg: sheet + '금액을 읽지 못함: ' + str(rawCost) });
      if (str(rawHours) && c.hours == null) issues.push({ row: r + 1, sheet: opts.sheet, msg: sheet + '소요시간을 읽지 못함: ' + str(rawHours) });
      if (str(rawBy) && !c.repairBy) issues.push({ row: r + 1, sheet: opts.sheet, msg: sheet + '외부/자체를 알아보지 못함: ' + str(rawBy) });
      if (!symptom) issues.push({ row: r + 1, sheet: opts.sheet, msg: sheet + '증상이 비어 있음' });
      if (!str(get('cause')) && map.cause != null) issues.push({ row: r + 1, sheet: opts.sheet, msg: sheet + '원인이 비어 있어 증상·조치에서 짐작함: ' + c.cause });
      rows.push(c);
    }
    return { rows: rows, issues: issues };
  }
  function mergeCases(existing, incoming) {
    var keys = {}; (existing || []).forEach(function (c) { keys[dupKey(c)] = 1; });
    var added = [], dup = 0;
    (incoming || []).forEach(function (c) { var k = dupKey(c); if (keys[k]) { dup++; return; } keys[k] = 1; added.push(c); });
    return { cases: (existing || []).concat(added), added: added.length, dup: dup };
  }

  /* ── 13-2. 자료 점검 — 「이렇게 올리면 되나요?」 에 답하는 화면용 ───
     칸별로 몇 건이 채워졌는지, 어느 기능이 지금 자료로 되는지, 어떤 칸을 더 채우면 진단·가이드·비용이 좋아지는지. */
  var CHECK_FIELDS = [
    { id: 'date', label: '일자', need: '권장', why: '월별 비용·같은 고장 반복 간격을 봅니다. 「2025-03 2주차」 도 읽지만 날짜(일)까지 있으면 중복 보고를 더 잘 가려냅니다.' },
    { id: 'equipment', label: '설비명', need: '필수', why: '같은 설비·같은 종류 사례를 앞으로 올립니다. 「압출기 1호기」 처럼 종류 + 호기로 적으면 됩니다.' },
    { id: 'symptom', label: '증상', need: '필수', why: '진단 검색의 재료입니다. 「무엇이 · 어떻게」(「냉각기 누수, 바닥에 물 고임」) 로 적습니다.' },
    { id: 'cause', label: '원인(점검 결과)', need: '필수', why: '진단이 내놓는 답입니다. 지금은 증상·조치 글에서 짐작하는데, 업체가 말한 원인을 따로 적으면 그대로 씁니다.' },
    { id: 'steps', label: '조치 순서', need: '권장', why: '가이드북의 「확인 순서」 가 됩니다. 「1) 전원 차단 2) … 확인 3) … 교체」 처럼 단계로 적으면 초보자가 멈출 곳까지 나옵니다.' },
    { id: 'parts', label: '교체 부품', need: '권장', why: '가이드북의 「자주 바꾼 부품」, 예비 부품 준비에 씁니다.' },
    { id: 'repairBy', label: '외부/자체 구분', need: '필수', why: '자체 처리 비율·외부 비용의 기준입니다. 업체 칸이 「미기재」 면 구분을 모릅니다.' },
    { id: 'cost', label: '비용(실제 지출)', need: '권장', why: '외부 수리 비용 현황입니다. 견적과 실제 지출은 따로, 「개당」 이면 수량도 함께 적습니다.' },
    { id: 'hours', label: '멈춘 시간·소요시간', need: '선택', why: '「반나절 라인 중단」 같은 손실을 숫자로 모읍니다.' },
    { id: 'status', label: '처리 상태', need: '선택', why: '완료 · 진행 중을 나눠 아직 안 끝난 고장을 따로 봅니다.' },
    { id: 'photos', label: '증상 사진', need: '선택', why: '사례에 붙이면 진단 결과에서 바로 봅니다(줄여서 이 브라우저에만 저장).' }
  ];
  function filled(c, id) {
    if (id === 'date') return !!c.date;
    if (id === 'cause') return !isUnknownCause(c.cause) && !c.causeInferred;
    if (id === 'steps') return splitSteps(c.steps).length > 0;
    if (id === 'parts' || id === 'tools') return splitList(c[id]).length > 0;
    if (id === 'repairBy') return c.repairBy === '외부' || c.repairBy === '자체';
    if (id === 'cost') return hasSpent(c);
    if (id === 'hours') return c.hours != null;
    if (id === 'status') return !!c.status;
    if (id === 'photos') return (c.photos || []).length > 0;
    return !!str(c[id]);
  }
  function dataCheck(cases) {
    var list = cases || [], n = list.length;
    function cnt(fn) { return list.filter(fn).length; }
    var fields = CHECK_FIELDS.map(function (f) {
      var k = cnt(function (c) { return filled(c, f.id); });
      var r = { id: f.id, label: f.label, need: f.need, why: f.why, filled: k, rate: n ? k / n : 0, detail: '' };
      if (f.id === 'date') { var dd = cnt(function (c) { return c.datePrec === 'day'; }), wk = cnt(function (c) { return c.datePrec === 'week'; }), mo = cnt(function (c) { return c.datePrec === 'month'; }); r.detail = '일까지 ' + dd + ' · 주차 ' + wk + ' · 월만 ' + mo; }
      if (f.id === 'cause') { var inf = cnt(function (c) { return c.causeInferred && !isUnknownCause(c.cause); }), unk = cnt(function (c) { return isUnknownCause(c.cause); }); r.detail = '직접 적음 ' + k + ' · 증상·조치에서 짐작 ' + inf + ' · 미상 ' + unk; r.inferred = inf; r.unknown = unk; }
      if (f.id === 'steps') { var two = cnt(function (c) { return splitSteps(c.steps).length >= 2; }); r.detail = '한 줄 ' + (k - two) + ' · 두 단계 이상 ' + two; r.multi = two; }
      if (f.id === 'repairBy') { r.detail = '외부 ' + cnt(function (c) { return c.repairBy === '외부'; }) + ' · 자체 ' + cnt(function (c) { return c.repairBy === '자체'; }) + ' · 모름 ' + (n - k); }
      if (f.id === 'cost') {
        var by = {}; list.forEach(function (c) { var kk = c.costKind || 'missing'; by[kk] = (by[kk] || 0) + 1; });
        r.detail = Object.keys(COST_KIND_TEXT).filter(function (x) { return by[x]; }).map(function (x) { return COST_KIND_TEXT[x] + ' ' + by[x]; }).join(' · ');
      }
      if (f.id === 'status') r.detail = '완료 ' + cnt(function (c) { return c.status === '완료'; }) + ' · 진행 중 ' + cnt(function (c) { return c.status === '진행 중'; }) + ' · 모름 ' + (n - k);
      return r;
    });
    var F = {}; fields.forEach(function (f) { F[f.id] = f; });
    // 기능별로 지금 되는지
    var causes = causeList(list).filter(function (x) { return !isUnknownCause(x.cause); });
    var guideReady = causes.filter(function (x) { return list.filter(function (c) { return causeKey(c.cause) === x.key && splitSteps(c.steps).length >= 2; }).length >= 1; }).length;
    var known = cnt(function (c) { return c.symptom && !isUnknownCause(c.cause); });
    var features = [
      { name: '① 증상 진단', ok: known >= 10, level: known >= 30 ? '좋음' : known >= 10 ? '됨' : '부족',
        text: '증상과 원인(직접 또는 짐작)이 있는 사례 ' + known + '건 · 원인 묶음 ' + causes.length + '가지' + (F.cause.inferred ? ' (그중 짐작 ' + F.cause.inferred + '건 — 원인 칸을 채우면 정확해집니다)' : '') },
      { name: '③ 자체 정비 가이드북', ok: guideReady >= 3, level: guideReady >= 5 && n && F.steps.multi / n >= 0.3 ? '좋음' : guideReady >= 3 ? '됨' : '부족',
        text: '조치가 두 단계 이상 적힌 원인 ' + guideReady + '가지 · 한 줄 조치 ' + (F.steps.filled - F.steps.multi) + '건은 확인 순서가 한 줄짜리 카드가 됩니다' },
      { name: '④ 비용 현황', ok: F.cost.filled >= 5 && F.repairBy.rate >= 0.5, level: F.cost.rate >= 0.6 && F.repairBy.rate >= 0.8 ? '좋음' : F.cost.filled >= 5 ? '일부' : '부족',
        text: '금액이 적힌 사례 ' + F.cost.filled + '건(' + fmtPct(F.cost.rate) + ') · 외부/자체를 아는 사례 ' + F.repairBy.filled + '건(' + fmtPct(F.repairBy.rate) + ')' }
    ];
    // 더 채우면 좋은 칸 — 필수·권장 중 채움률이 낮은 순
    var order = { '필수': 0, '권장': 1, '선택': 2 };
    var todo = fields.filter(function (f) { return f.rate < 0.8 && f.need !== '선택'; })
      .sort(function (a, b) { return order[a.need] - order[b.need] || a.rate - b.rate; });
    var dups = dupSuspects(list);
    return {
      n: n, fields: fields, features: features, todo: todo, dups: dups,
      dupSameCost: dups.filter(function (d) { return d.sameCost; }),
      pending: list.filter(function (c) { return c.status === '진행 중'; }).map(function (c) { return c.id; }),
      estimates: list.filter(function (c) { return c.costKind === 'estimate'; }).map(function (c) { return c.id; }),
      unitPrices: list.filter(function (c) { return c.costKind === 'unit'; }).map(function (c) { return c.id; }),
      families: countList(list.map(function (c) { return equipFamily(c.equipment); }).filter(Boolean)),
      parts: countList(list.map(function (c) { return partOf(c.cause); }).filter(Boolean))
    };
  }
  /* 권장 양식 — 지금 양식(연도·설비명·일자·증상·처리일자·수리업체·비용·비고·출처)에 칸 4개(원인·조치 순서·교체 부품·외부/자체)를 더한 모양.
     예시 두 줄은 지어낸 값입니다 */
  var RECOMMENDED_HEAD = ['일자', '설비명', '증상', '원인(점검 결과)', '조치 순서', '교체 부품', '외부/자체', '수리 업체', '비용(원)', '견적/실지출', '멈춘 시간(시간)', '처리 상태', '사진 파일명', '비고'];
  var RECOMMENDED_EXAMPLE = [
    ['2026-03-12', '사출기 1호기', '실린더 온도가 설정값까지 안 올라가고 알람 후 멈춤', '히터 밴드 단선(2존)', '1) 전원 차단·잠금 2) 존별 온도 표시 확인 — 2존만 낮음 3) 히터 단자 변색 확인 4) 히터 저항 측정(업체) 5) 히터 밴드 교체', '히터 밴드 1EA', '외부', 'A정비(가상)', 680000, '실지출', 4, '완료', '사출1_0312_히터.jpg', '(가상 예시)'],
    ['2026-04-03', '냉각기 2호기', '냉각이 약하고 배관 연결부에 기름 자국', '냉매 배관 누유', '1) 전원 차단 2) 누유 자국 위치 사진 3) 업체 점검 요청 4) 배관 교체·냉매 충전(업체)', '냉매 배관 1식', '외부', 'B냉동(가상)', 1200000, '견적', 8, '진행 중', '냉각2_0403_누유.jpg', '(가상 예시)']
  ];
  /* 지금 사례를 권장 양식으로 — 빈 칸을 채워 다시 가져오라는 용도. 짐작한 원인은 원인 칸을 비우고 맨 끝 「참고」 칸에만 둡니다(다시 가져올 때 직접 적은 원인으로 오인하지 않게) */
  function toRecommendedAoa(cases) {
    var head = RECOMMENDED_HEAD.concat(['도구가 짐작한 원인(참고)']);
    return [head].concat((cases || []).map(function (c) {
      var d = c.datePrec === 'month' ? String(c.date).slice(0, 7) : c.datePrec === 'week' ? String(c.date).slice(0, 7) + ' ' + (Math.floor((Number(String(c.date).slice(8, 10)) - 1) / 7) + 1) + '주차' : c.date;
      return [d || '', c.equipment, c.symptom, c.causeInferred || isUnknownCause(c.cause) ? '' : c.cause,
        splitSteps(c.steps).map(function (x, i) { return (i + 1) + ') ' + x; }).join(' '), splitList(c.parts).join(', '), c.repairBy || '', c.vendor || '',
        c.cost == null ? '' : c.cost, c.costKind === 'estimate' ? '견적' : c.cost != null ? '실지출' : '', c.hours == null ? '' : c.hours, c.status || '', c.media || '', c.memo || '',
        c.causeInferred ? canonCause(c.cause) : ''];
    }));
  }
  function recommendedAoa() { return [RECOMMENDED_HEAD].concat(RECOMMENDED_EXAMPLE.map(function (r) { return r.slice(); })); }

  /* ── 14. 내보내기 표 ───────────────────────────────────────── */
  var EXPORT_HEAD = ['일자', '설비', '증상', '증상 태그', '원인', '원인 묶음', '조치 단계', '교체 부품', '필요 공구', '외부/자체', '비용(원)', '소요시간(시간)', '사진 수', '사진·영상 위치', '비고',
    '일자 정밀도', '설비 종류', '부위', '원인 짐작 여부', '수리 업체', '비용 구분', '처리 상태', '출처'];
  function casesToAoa(cases) {
    return [EXPORT_HEAD].concat((cases || []).map(function (c) {
      return [c.date, c.equipment, c.symptom, (c.tags || []).join(', '), c.cause, canonCause(c.cause),
        splitSteps(c.steps).map(function (s, i) { return (i + 1) + ') ' + s; }).join('\n'), splitList(c.parts).join(', '), splitList(c.tools).join(', '),
        c.repairBy, c.cost == null ? '' : c.cost, c.hours == null ? '' : c.hours, (c.photos || []).length, c.media, c.memo,
        PREC_TEXT[c.datePrec] || '', equipFamily(c.equipment), partOf(c.cause), c.causeInferred ? '짐작' : '', c.vendor || '', COST_KIND_TEXT[c.costKind] || '', c.status || '', c.ref || ''];
    }));
  }
  function templateAoa() { return [FIELDS.map(function (f) { return f.label; })]; }

  /* ── 15. 저장 형식·백업 ────────────────────────────────────── */
  function emptyDb() { return { v: SCHEMA_VERSION, cases: [], equipments: [], guides: {}, mapping: null, opts: cleanOpts(), seq: 0, nextChecks: {}, _sample: false }; }
  function cleanEquipment(e) {
    e = e || {};
    return { name: str(e.name), type: str(e.type), model: str(e.model), maker: str(e.maker), manual: str(e.manual), memo: str(e.memo) };
  }
  function restoreDb(p) {
    var db = emptyDb();
    if (!p || typeof p !== 'object') return db;
    if (Array.isArray(p.cases)) db.cases = p.cases.map(function (c) { return cleanCase(c); });
    if (Array.isArray(p.equipments)) db.equipments = p.equipments.map(cleanEquipment).filter(function (e) { return e.name; });
    if (p.guides && typeof p.guides === 'object') db.guides = p.guides;
    if (p.mapping && typeof p.mapping === 'object') db.mapping = p.mapping;
    db.opts = cleanOpts(p.opts);
    if (typeof p.seq === 'number') db.seq = p.seq;
    db.cases.forEach(function (c) { var n = Number(String(c.id).replace(/^\D+/, '')); if (isFinite(n) && n > db.seq) db.seq = n; });
    if (p.nextChecks && typeof p.nextChecks === 'object') db.nextChecks = p.nextChecks;
    if (p._sample) db._sample = true;
    return db;
  }
  /* 사례에 나온 설비 이름이 설비 목록에 없으면 이름만 넣어 둡니다 */
  function syncEquipments(db) {
    var have = {}; db.equipments.forEach(function (e) { have[e.name] = 1; });
    var added = 0;
    db.cases.forEach(function (c) { if (c.equipment && !have[c.equipment]) { have[c.equipment] = 1; db.equipments.push(cleanEquipment({ name: c.equipment, type: equipFamily(c.equipment) })); added++; } });
    return added;
  }
  function equipTypes(db) { var m = {}; (db.equipments || []).forEach(function (e) { m[e.name] = e.type; }); return m; }
  function buildBackup(db, photos, when) {
    return { app: 'data10-06', kind: 'backup', v: SCHEMA_VERSION, exportedAt: when || new Date().toISOString(), db: db, photos: photos || [] };
  }
  function parseBackup(textOrObj) {
    var o = textOrObj;
    if (typeof o === 'string') { try { o = JSON.parse(o); } catch (e) { throw new Error('JSON 으로 읽지 못했습니다.'); } }
    if (!o || o.app !== 'data10-06' || !o.db) throw new Error('이 도구(data10-06)의 백업 파일이 아닙니다.');
    return { db: restoreDb(o.db), photos: Array.isArray(o.photos) ? o.photos.filter(function (p) { return p && p.id && p.data; }) : [], exportedAt: str(o.exportedAt) };
  }

  /* ── 16. 사진 ──────────────────────────────────────────────── */
  var PHOTO_MAX = 1280, THUMB_MAX = 240, PHOTO_LIMIT = 10;
  function fitSize(w, h, max) {
    if (!(w > 0 && h > 0)) return { w: max, h: max };
    var s = Math.min(1, max / Math.max(w, h));
    return { w: Math.max(1, Math.round(w * s)), h: Math.max(1, Math.round(h * s)) };
  }
  function allPhotoIds(cases) { var out = []; (cases || []).forEach(function (c) { (c.photos || []).forEach(function (p) { out.push(p.id); }); }); return out; }
  function photoBytes(cases) { var n = 0; (cases || []).forEach(function (c) { (c.photos || []).forEach(function (p) { n += Number(p.bytes) || 0; }); }); return n; }
  function orphanPhotoIds(cases, storedIds, keep) {
    var used = {}; allPhotoIds(cases).concat(keep || []).forEach(function (id) { used[id] = 1; });
    return (storedIds || []).filter(function (id) { return !used[id]; });
  }
  /* 저장 공간 경고 단계 — 사진 합계(바이트)와 브라우저가 알려 준 남은 공간으로 */
  function storageLevel(photoBytesTotal, quota) {
    var mb = photoBytesTotal / 1048576;
    if (quota && quota.quota > 0 && quota.usage / quota.quota > 0.8) return 'bad';
    if (mb >= 200) return 'bad';
    if (mb >= 50) return 'warn';
    return 'ok';
  }

  return {
    SCHEMA_VERSION: SCHEMA_VERSION, REPAIR: REPAIR, TAGS: TAGS, CAUSE_GROUPS: CAUSE_GROUPS, FIELDS: FIELDS, DEFAULT_OPTS: DEFAULT_OPTS,
    SAFETY: SAFETY, CALL_EXPERT: CALL_EXPERT, LEVEL_TEXT: LEVEL_TEXT, EXPORT_HEAD: EXPORT_HEAD, GUIDE_EDIT_FIELDS: GUIDE_EDIT_FIELDS,
    PHOTO_MAX: PHOTO_MAX, THUMB_MAX: THUMB_MAX, PHOTO_LIMIT: PHOTO_LIMIT,
    str: str, squash: squash, toISO: toISO, toNum: toNum, toHours: toHours, monthOf: monthOf, round: round, fmtWon: fmtWon, fmtPct: fmtPct, uniq: uniq,
    normRepairBy: normRepairBy, splitSteps: splitSteps, splitList: splitList, tagList: tagList, extractTags: extractTags,
    canonCause: canonCause, causeKey: causeKey, causeList: causeList, hazardCats: hazardCats,
    stem: stem, tokens: tokens, buildIndex: buildIndex, similar: similar, cleanOpts: cleanOpts, diagnose: diagnose,
    stepKind: stepKind, stepCore: stepCore, stopIndex: stopIndex, firstCheck: firstCheck, guideDraft: guideDraft, guideCard: guideCard, guideToText: guideToText, linesOf: linesOf,
    monthRange: monthRange, costReport: costReport,
    buildDiagnosePrompt: buildDiagnosePrompt, buildGuidePrompt: buildGuidePrompt, parseGuideAnswer: parseGuideAnswer,
    cleanCase: cleanCase, validateCase: validateCase, dupKey: dupKey,
    matchField: matchField, detectHeaderRow: detectHeaderRow, headersOf: headersOf, colName: colName, autoMap: autoMap, mappingToSaved: mappingToSaved,
    missingRequired: missingRequired, applyMapping: applyMapping, mergeCases: mergeCases, casesToAoa: casesToAoa, templateAoa: templateAoa,
    emptyDb: emptyDb, restoreDb: restoreDb, cleanEquipment: cleanEquipment, syncEquipments: syncEquipments, equipTypes: equipTypes,
    buildBackup: buildBackup, parseBackup: parseBackup,
    fitSize: fitSize, allPhotoIds: allPhotoIds, photoBytes: photoBytes, orphanPhotoIds: orphanPhotoIds, storageLevel: storageLevel,
    // 2차(실제 수리이력 반영)
    PREC_TEXT: PREC_TEXT, COST_KIND_TEXT: COST_KIND_TEXT, UNKNOWN_CAUSE: UNKNOWN_CAUSE, CHECK_FIELDS: CHECK_FIELDS, RECOMMENDED_HEAD: RECOMMENDED_HEAD, SYNONYMS: SYNONYMS,
    looseDate: looseDate, parseCost: parseCost, spent: spent, hasSpent: hasSpent, parseVendor: parseVendor, caseStatus: caseStatus, actionOf: actionOf,
    inferCause: inferCause, isUnknownCause: isUnknownCause, partOf: partOf, equipFamily: equipFamily, equipUnit: equipUnit, expandText: expandText,
    unmerge: unmerge, sheetPlans: sheetPlans, dupSuspects: dupSuspects, dataCheck: dataCheck, recommendedAoa: recommendedAoa, toRecommendedAoa: toRecommendedAoa
  };
});
