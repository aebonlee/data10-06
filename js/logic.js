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
      s = s.replace(/(^|\s)(\d{1,2}[.)]|[①-⑳])\s*/g, '\n').replace(/\s*(→|->|=>|;|\s\/\s)\s*/g, '\n');
      parts = s.split('\n');
    }
    return parts.map(function (p) { return str(p).replace(/^(\d{1,2}[.)]|[①-⑳]|[-*•·])\s*/, ''); }).filter(Boolean);
  }
  function splitList(v) {
    if (Array.isArray(v)) return v.map(str).filter(Boolean);
    return str(v).split(/\s*(?:[,/;·\n]|\s및\s)\s*/).map(str).filter(Boolean);
  }

  /* ── 3. 증상 태그 ─────────────────────────────────────────────
     증상 글에서 자동으로 뽑고, 사용자가 켜고 끌 수 있습니다. 같은 태그가 있으면 유사도에 더해집니다. */
  var TAGS = [
    { tag: '온도 안 오름', re: /온도.{0,8}(안\s?올라|안\s?오르|안\s?오름|못\s?올라|낮|미달|저하|떨어)|승온.{0,3}(불량|안|지연|느|이 안)|히팅.{0,2}(안|불량)/ },
    { tag: '온도 표시 불안정', re: /온도.{0,8}(들쭉|튀|불안정|오락가락|흔들|왔다갔다|0\s?℃|0도)|표시\s?온도.{0,6}(다름|낮|이상)/ },
    { tag: '과열', re: /과열|너무\s?뜨|뜨거|온도.{0,4}(높|올라감|초과)|열\s?남|열이\s?남/ },
    { tag: '설비 정지', re: /멈춤|멈춰|멈췄|멈추|정지|섰|꺼짐|꺼졌|다운/ },
    { tag: '알람', re: /알람|경보|에러|alarm|오류\s?표시/ },
    { tag: '이상 소음', re: /소음|소리|딱딱|끼익|웅웅|덜컹|굉음/ },
    { tag: '진동', re: /진동|떨림/ },
    { tag: '누유', re: /누유|기름|오일.{0,4}(샘|새|묻|자국)|새어|새고|샌다/ },
    { tag: '압력 저하', re: /압력.{0,4}(저하|떨어|낮|늦게|안\s?올라)|형체력.{0,3}약/ },
    { tag: '차단기·과부하', re: /차단기|트립|써멀|서멀|과전류|과부하|두꺼비집/ },
    { tag: '탄내·연기', re: /탄\s?냄새|타는\s?냄새|탄내|연기|스파크/ },
    { tag: '회전 이상', re: /회전.{0,6}(멈|무거|안|이상)|안\s?돌|스크류.{0,4}(멈|무거)/ },
    { tag: '토출·성형 불량', re: /미성형|덜\s?채|안\s?나와|안\s?나옴|잘\s?안\s?나|토출|압출량|쇼트|흐름.{0,4}(고르지|불균일)/ }
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
  var CAUSE_GROUPS = [
    { name: '온도 센서 불량', cats: ['전기'], syn: ['온도센서', '센서불량', '센서단선', '센서접촉', '열전대', '써모커플', '서모커플', 'thermocouple', 'tc불량', 'tc단선', '측온'] },
    { name: '히터 단선', cats: ['전기', '고온'], syn: ['히터단선', '히터끊', '히터밴드단선', '히터소손', '히터불량', '히터고장', '히터선단선'] },
    { name: '유압 누유', cats: ['유압'], syn: ['유압누유', '유압유누유', '오일누유', '호스누유', '피팅누유', '실누유', '씰누유', '오링', '유압호스', '유압피팅', '누유'] },
    { name: '모터 과부하', cats: ['전기', '회전체'], syn: ['모터과부하', '과부하', '모터과열', '오버로드', 'overload', '써멀트립', '서멀트립', '모터베어링'] },
    { name: '금형 막힘', cats: ['고온'], syn: ['금형막힘', '노즐막힘', '게이트막힘', '다이막힘', '스크린막힘', '런너막힘', '막힘'] }
  ];
  function causeCore(text) { return str(text).replace(/\s*[(（\[][^)）\]]*[)）\]]\s*/g, ' ').trim(); }
  function canonCause(text) {
    var core = squash(causeCore(text)) || squash(text);
    if (!core) return '';
    for (var i = 0; i < CAUSE_GROUPS.length; i++) {
      var g = CAUSE_GROUPS[i];
      for (var j = 0; j < g.syn.length; j++) if (core.indexOf(g.syn[j]) >= 0) return g.name;
    }
    return causeCore(text) || str(text);
  }
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
    addTokens(tf, tokens(c.symptom), 1);
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
    var tf = {}; addTokens(tf, tokens(text), 1);
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

  var DEFAULT_OPTS = { topN: 8, minScore: 0.08, tagWeight: 0.35, sameEquipBonus: 0.08, sameTypeBonus: 0.03 };
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
    var qEq = str(query.equipment), qType = str(query.equipmentType) || (equipTypes && equipTypes[qEq]) || '';
    var out = [];
    index.docs.forEach(function (d) {
      var cos = cosine(q.vec, q.norm, d.vec, d.norm);
      var tagScore = jaccard(qTags, d.c.tags || []);
      var base = qTags.length ? cos * (1 - o.tagWeight) + tagScore * o.tagWeight : cos;
      if (base <= 0) return;
      var sameEq = !!qEq && d.c.equipment === qEq;
      var cType = (equipTypes && equipTypes[d.c.equipment]) || '';
      var sameType = !sameEq && !!qType && cType === qType;
      var score = base + (sameEq ? o.sameEquipBonus : 0) + (sameType ? o.sameTypeBonus : 0);
      if (score < o.minScore) return;
      var hits = [];
      Object.keys(q.vec).forEach(function (t) { if (t.slice(0, 2) === 'w:' && d.vec[t]) hits.push(t.slice(2)); });
      out.push({ c: d.c, score: round(score, 4), cos: round(cos, 4), tagScore: round(tagScore, 4), sameEquipment: sameEq, sameType: sameType, hits: hits,
        sharedTags: qTags.filter(function (t) { return (d.c.tags || []).indexOf(t) >= 0; }) });
    });
    out.sort(function (a, b) { return b.score - a.score || String(b.c.date).localeCompare(String(a.c.date)); });
    return out.slice(0, o.topN);
  }

  /* ── 6. 단계 나누기: 안전 · 초보 가능 · 담당자·전문가 ─────────
     「먼저 확인할 것」 은 안전 단계(전원 차단·잔압 빼기)를 뺀 첫 확인 단계입니다.
     전기 측정·배선, 유압 라인 분리·교체, 고온부 분해, 회전체 분해는 초보자가 하지 않고 담당자·외부를 부르는 단계로 표시합니다. */
  var SAFETY_STEP = /전원\s?차단|loto|잠금|조작\s?금지|잔압|압력.{0,3}(빼|제거|해제)|설비\s?정지\s?후|정지·전원|정지\s?후\s?전원|비상\s?정지/i;
  var EXPERT_STEP = /담당|전기\s?교육|입회|전문가|측정|테스터|멀티미터|메거|절연|배선|결선|전류|전압|인버터|(히터|모터|베어링|호스|실|씰|펌프|커플링|열전대|센서).{0,4}교체|교체.{0,3}(히터|모터|베어링|호스)|분해|라인.{0,3}분리|호스.{0,3}분리|어큐뮬|릴리프|용접|스크류.{0,3}(인출|빼)/;
  function stepKind(step) {
    var s = str(step);
    if (SAFETY_STEP.test(s)) return 'safety';
    if (EXPERT_STEP.test(s)) return 'expert';
    return 'basic';
  }
  /* 「존별 온도 표시 확인 — 2존만 낮음」 처럼 그때의 결과가 붙은 단계는 결과를 떼고 모읍니다 */
  function stepCore(step) { return str(step).replace(/\s+[—–-]\s[^()]*$/, ''); }
  function stepKey(step) { return squash(stepCore(step).replace(/\([^)]*\)/g, '')); }
  function firstCheck(c) {
    var st = splitSteps(c.steps);
    for (var i = 0; i < st.length; i++) if (stepKind(st[i]) !== 'safety') return st[i];
    return '';
  }

  /* ── 7. 진단 — 비슷한 사례로 원인 순위 매기기 ───────────────
     원인마다 비슷한 사례들의 유사도를 더해 비율을 냅니다(투표). 확정 진단이 아니라 「과거 사례에서 이 증상일 때 무엇이었나」 입니다. */
  function diagnose(cases, query, opts, equipTypes) {
    var sims = similar(cases, query, opts, equipTypes);
    var groups = {}, order = [], total = 0;
    sims.forEach(function (s) {
      var k = causeKey(s.c.cause);
      if (!k) return;
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
      var extCost = ext.reduce(function (s, c) { return s + (Number(c.cost) || 0); }, 0);
      return {
        key: k, cause: g.cause, share: total ? g.weight / total : 0, weight: round(g.weight, 4), count: g.items.length,
        best: g.items[0].score, caseIds: g.items.map(function (s) { return s.c.id; }), checks: checks,
        totalCases: same.length, extCount: ext.length, selfCount: same.length - ext.length, avgExtCost: ext.length ? Math.round(extCost / ext.length) : null,
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
    '회전체': '모터·스크류·벨트 같은 회전부는 완전히 멈춘 것을 확인하고 덮개를 엽니다. 장갑·소매가 말려 들어가지 않게 합니다.'
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
    '회전체': '베어링·커플링에서 소음이 나거나 축을 분해해야 할 때.'
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
    var st = {}, sOrder = [];
    list.forEach(function (c) {
      var s = splitSteps(c.steps).filter(function (x) { return stepKind(x) !== 'safety'; });
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
    var extCost = ext.reduce(function (s, c) { return s + (Number(c.cost) || 0); }, 0);
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
      parts: countList(list.reduce(function (a, c) { return a.concat(splitList(c.parts)); }, [])),
      tools: uniq(tools),
      safety: safety, callExpert: callExpert,
      extCount: ext.length, selfCount: self.length, extCost: extCost, avgExtCost: ext.length ? Math.round(extCost / ext.length) : null,
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
      symptoms: draft.symptoms, steps: draft.steps.map(function (s) { return { text: s.text, kind: s.kind, count: s.count }; }),
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
  function stat(list) {
    var ext = list.filter(function (c) { return c.repairBy === '외부'; });
    var self = list.filter(function (c) { return c.repairBy === '자체'; });
    var known = ext.length + self.length;
    return {
      total: list.length, extCount: ext.length, selfCount: self.length,
      extCost: ext.reduce(function (s, c) { return s + (Number(c.cost) || 0); }, 0),
      selfCost: self.reduce(function (s, c) { return s + (Number(c.cost) || 0); }, 0),
      hours: round(list.reduce(function (s, c) { return s + (Number(c.hours) || 0); }, 0), 1),
      selfRatio: known ? self.length / known : null
    };
  }
  function costReport(cases, from, to) {
    var dated = (cases || []).filter(function (c) { return c.date; });
    var months = dated.map(function (c) { return monthOf(c.date); }).sort();
    var f = from || months[0] || '', t = to || months[months.length - 1] || '';
    var list = dated.filter(function (c) { return inPeriod(c, f, t); });
    var monthly = monthRange(f, t).map(function (m) {
      var s = stat(list.filter(function (c) { return monthOf(c.date) === m; })); s.month = m; return s;
    });
    function by(getKey, getLabel) {
      var g = {}, order = [];
      list.forEach(function (c) { var k = getKey(c); if (!k) return; if (!g[k]) { g[k] = { label: getLabel(c), items: [] }; order.push(k); } g[k].items.push(c); });
      return order.map(function (k) { var s = stat(g[k].items); s.key = k; s.label = g[k].label; return s; })
        .sort(function (a, b) { return b.extCost - a.extCost || b.total - a.total; });
    }
    var byEquipment = by(function (c) { return str(c.equipment) || '(설비 미기재)'; }, function (c) { return str(c.equipment) || '(설비 미기재)'; });
    var byCause = by(function (c) { return causeKey(c.cause); }, function (c) { return canonCause(c.cause); });
    // 자체 전환 후보 — 자체로 고친 적이 있는데 외부도 부른 원인. 외부 비용이 큰 순
    var candidates = byCause.filter(function (r) { return r.selfCount > 0 && r.extCount > 0; });
    // 앞·뒤 절반 비교 — 자체 처리 비율이 늘었는지
    var half = Math.floor(monthly.length / 2), early = null, late = null;
    if (monthly.length >= 2) {
      early = stat(list.filter(function (c) { return monthOf(c.date) < monthly[half].month; }));
      late = stat(list.filter(function (c) { return monthOf(c.date) >= monthly[half].month; }));
    }
    return { from: f, to: t, total: stat(list), monthly: monthly, byEquipment: byEquipment, byCause: byCause, candidates: candidates,
      early: early, late: late, splitMonth: monthly.length >= 2 ? monthly[half].month : '', undated: (cases || []).length - dated.length };
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
      }
      if (!cur) return;
      if (/^-{2,}.*초보자는 여기까지/.test(str(line))) return;
      var t = str(line).replace(/^(\d{1,2}[.)]|[-*•·])\s*/, '').replace(/\s*\[담당자·전문가[^\]]*\]\s*$/, '');
      if (t) out[cur].push(t);
    });
    Object.keys(out).forEach(function (k) { out[k] = k === 'note' ? out[k].join('\n') : out[k].join('\n'); if (!out[k]) delete out[k]; });
    return out;
  }

  /* ── 12. 정비 사례 한 건 ───────────────────────────────────── */
  function cleanCase(p, seq) {
    p = p || {};
    var c = {
      id: str(p.id) || ('c' + (seq != null ? seq : Date.now().toString(36))),
      date: toISO(p.date),
      equipment: str(p.equipment),
      symptom: str(p.symptom),
      tags: uniq(Array.isArray(p.tags) ? p.tags : splitList(p.tags)),
      cause: str(p.cause),
      steps: splitSteps(p.steps),
      parts: splitList(p.parts),
      tools: splitList(p.tools),
      repairBy: REPAIR.indexOf(str(p.repairBy)) >= 0 ? str(p.repairBy) : normRepairBy(p.repairBy),
      cost: toNum(p.cost),
      hours: toHours(p.hours),
      media: str(p.media),
      memo: str(p.memo),
      photos: Array.isArray(p.photos) ? p.photos.filter(function (x) { return x && x.id; }) : [],
      source: str(p.source) || '직접 입력',
      createdAt: str(p.createdAt)
    };
    if (c.cost != null && c.cost < 0) c.cost = null;
    return c;
  }
  function validateCase(c) {
    var e = {};
    if (!str(c.symptom)) e.symptom = '증상을 적어 주세요.';
    if (!str(c.cause)) e.cause = '원인을 적어 주세요. 모르면 「원인 미상」 이라고 적어 두고 나중에 고칩니다.';
    if (c.date && !toISO(c.date)) e.date = '날짜를 읽지 못했습니다.';
    return e;
  }
  function dupKey(c) { return [c.date, squash(c.equipment), squash(c.symptom), squash(c.cause)].join('|'); }

  /* ── 13. 정비 내역서 엑셀 가져오기 — 열 맞추기 ────────────── */
  var FIELDS = [
    { id: 'date', label: '일자', req: false, aliases: ['일자', '날짜', '수리일', '수리일자', '정비일', '정비일자', '발생일', '발생일자', '작업일', '작업일자', '고장일', '고장일자'] },
    { id: 'equipment', label: '설비', req: false, aliases: ['설비', '설비명', '장비', '장비명', '기계', '기계명', '호기', '설비번호'] },
    { id: 'symptom', label: '증상', req: true, aliases: ['증상', '고장내용', '고장현상', '현상', '이상내용', '이상현상', '의뢰내용', '고장증상', '불량내용'] },
    { id: 'tags', label: '증상 태그', req: false, aliases: ['증상태그', '태그', '증상분류', '고장분류', '고장유형'] },
    { id: 'cause', label: '원인', req: true, aliases: ['원인', '고장원인', '점검결과', '점검결과원인', '진단', '진단결과', '발생원인'] },
    { id: 'steps', label: '조치 단계', req: false, aliases: ['조치', '조치내용', '조치사항', '수리내용', '작업내용', '처리내용', '정비내용', '수리내역'] },
    { id: 'parts', label: '교체 부품', req: false, aliases: ['교체부품', '부품', '사용부품', '자재', '사용자재', '교체품'] },
    { id: 'tools', label: '필요 공구', req: false, aliases: ['공구', '사용공구', '필요공구'] },
    { id: 'repairBy', label: '외부/자체', req: false, aliases: ['외부자체', '수리구분', '구분', '수리주체', '처리구분', '정비구분', '외주여부'] },
    { id: 'cost', label: '비용(원)', req: false, aliases: ['비용', '금액', '수리비', '수리비용', '청구금액', '청구금액원', '합계', '합계금액', '견적', '비용원'] },
    { id: 'hours', label: '소요시간(시간)', req: false, aliases: ['소요시간', '작업시간', '수리시간', '정비시간', '시간', '소요시간h', '작업시간h'] },
    { id: 'media', label: '사진·영상 위치', req: false, aliases: ['사진', '영상', '동영상', '첨부', '사진위치', '파일'] },
    { id: 'memo', label: '비고', req: false, aliases: ['비고', '메모', '특이사항', '참고'] }
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
  /* 행 → 사례. 증상·원인이 모두 빈 줄은 건너뜁니다(합계 줄·빈 줄) */
  function applyMapping(aoa, hr, map, opts) {
    opts = opts || {};
    var rows = [], issues = [], seq = opts.seqStart || 1;
    for (var r = hr + 1; r < (aoa || []).length; r++) {
      var row = aoa[r] || [];
      function get(f) { return map[f] == null ? null : row[map[f]]; }
      var any = row.some(function (v) { return str(v) !== ''; });
      if (!any) continue;
      var symptom = str(get('symptom')), cause = str(get('cause'));
      if (!symptom && !cause) { issues.push({ row: r + 1, msg: '증상·원인이 비어 건너뜀' }); continue; }
      var rawDate = get('date'), rawCost = get('cost'), rawHours = get('hours'), rawBy = get('repairBy');
      var c = cleanCase({
        id: 'c' + (opts.idPrefix || 'x') + (seq++),
        date: rawDate, equipment: get('equipment'), symptom: symptom,
        tags: uniq(splitList(get('tags')).concat(extractTags(symptom))),
        cause: cause, steps: get('steps'), parts: get('parts'), tools: get('tools'),
        repairBy: rawBy, cost: rawCost, hours: rawHours, media: get('media'), memo: get('memo'), source: opts.source || '엑셀'
      });
      if (str(rawDate) && !c.date) issues.push({ row: r + 1, msg: '날짜를 읽지 못함: ' + str(rawDate) });
      if (str(rawCost) && c.cost == null) issues.push({ row: r + 1, msg: '금액을 읽지 못함: ' + str(rawCost) });
      if (str(rawHours) && c.hours == null) issues.push({ row: r + 1, msg: '소요시간을 읽지 못함: ' + str(rawHours) });
      if (str(rawBy) && !c.repairBy) issues.push({ row: r + 1, msg: '외부/자체를 알아보지 못함: ' + str(rawBy) });
      if (!symptom) issues.push({ row: r + 1, msg: '증상이 비어 있음' });
      if (!cause) issues.push({ row: r + 1, msg: '원인이 비어 있음' });
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

  /* ── 14. 내보내기 표 ───────────────────────────────────────── */
  var EXPORT_HEAD = ['일자', '설비', '증상', '증상 태그', '원인', '원인 묶음', '조치 단계', '교체 부품', '필요 공구', '외부/자체', '비용(원)', '소요시간(시간)', '사진 수', '사진·영상 위치', '비고'];
  function casesToAoa(cases) {
    return [EXPORT_HEAD].concat((cases || []).map(function (c) {
      return [c.date, c.equipment, c.symptom, (c.tags || []).join(', '), c.cause, canonCause(c.cause),
        splitSteps(c.steps).map(function (s, i) { return (i + 1) + ') ' + s; }).join('\n'), splitList(c.parts).join(', '), splitList(c.tools).join(', '),
        c.repairBy, c.cost == null ? '' : c.cost, c.hours == null ? '' : c.hours, (c.photos || []).length, c.media, c.memo];
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
    db.cases.forEach(function (c) { if (c.equipment && !have[c.equipment]) { have[c.equipment] = 1; db.equipments.push(cleanEquipment({ name: c.equipment })); added++; } });
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
    fitSize: fitSize, allPhotoIds: allPhotoIds, photoBytes: photoBytes, orphanPhotoIds: orphanPhotoIds, storageLevel: storageLevel
  };
});
