// 실행: node test/logic.test.mjs   (의존성 없음 — vendor 의 SheetJS 만 씀)
// 실제 정비 내역서는 아직 받지 않았습니다. 가상 사례 30건과 같은 모양의 가상 내역서로 검사합니다.
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
const require = createRequire(import.meta.url);
const L = require('../js/logic.js');
const Sample = require('../js/sample-data.js');
const E = require('../js/ai-endpoint.js');
const XLSX = require('../vendor/xlsx.full.min.js');

let passed = 0;
function test(name, fn) {
  try { fn(); passed++; console.log('  ok  ' + name); }
  catch (e) { console.error('  FAIL ' + name + '\n       ' + e.message); process.exitCode = 1; }
}
const db = Sample.db();
const T = L.equipTypes(db);
const diag = (text, extra = {}) => L.diagnose(db.cases, Object.assign({ text }, extra), null, T);

console.log('1. 값 읽기');
test('날짜 — 여러 모양', () => {
  assert.equal(L.toISO('2026-03-05'), '2026-03-05');
  assert.equal(L.toISO('2026.3.5'), '2026-03-05');
  assert.equal(L.toISO('2026/03/05'), '2026-03-05');
  assert.equal(L.toISO('26.3.5'), '2026-03-05');
  assert.equal(L.toISO(20260305), '2026-03-05');
  assert.equal(L.toISO(46086), '2026-03-05'); // 엑셀 일련번호
  assert.equal(L.toISO(new Date(2026, 2, 5)), '2026-03-05');
  assert.equal(L.toISO('2026-02-30'), '');
  assert.equal(L.toISO('어제'), '');
});
test('금액 — 쉼표·원·만원', () => {
  assert.equal(L.toNum('1,250,000원'), 1250000);
  assert.equal(L.toNum('₩680,000'), 680000);
  assert.equal(L.toNum('15만원'), 150000);
  assert.equal(L.toNum('1.5만'), 15000);
  assert.equal(L.toNum('3만5000'), 35000);
  assert.equal(L.toNum('-'), null);
  assert.equal(L.toNum('견적 대기'), null);
});
test('소요시간 — 시간·분', () => {
  assert.equal(L.toHours('1시간 30분'), 1.5);
  assert.equal(L.toHours('90분'), 1.5);
  assert.equal(L.toHours('2h'), 2);
  assert.equal(L.toHours(2.5), 2.5);
  assert.equal(L.toHours('반나절'), null);
});
test('외부/자체 알아보기', () => {
  assert.equal(L.normRepairBy('외부업체'), '외부');
  assert.equal(L.normRepairBy('A/S'), '외부');
  assert.equal(L.normRepairBy('자체(팀)'), '자체');
  assert.equal(L.normRepairBy('사내 정비'), '자체');
  assert.equal(L.normRepairBy('?'), '');
});
test('조치 한 칸 → 단계', () => {
  assert.deepEqual(L.splitSteps('1) 전원 차단 2) 센서 확인 3) 교체'), ['전원 차단', '센서 확인', '교체']);
  assert.deepEqual(L.splitSteps('①전원 차단 ②센서 확인'), ['전원 차단', '센서 확인']);
  assert.deepEqual(L.splitSteps('전원 차단 → 센서 확인 → 교체'), ['전원 차단', '센서 확인', '교체']);
  assert.deepEqual(L.splitSteps('전원 차단\n2. 센서 확인'), ['전원 차단', '센서 확인']);
  assert.deepEqual(L.splitList('히터 밴드, 오링 / 유압유 20L'), ['히터 밴드', '오링', '유압유 20L']);
});

console.log('2. 증상 태그·원인 묶기');
test('증상 글에서 태그 — 수강생 예시 문장', () => {
  assert.deepEqual(L.extractTags('온도가 안 올라가고 기계가 멈춤'), ['온도 안 오름', '설비 정지']);
  assert.ok(L.extractTags('기계 아래 바닥에 기름이 고임').includes('누유'));
  assert.ok(L.extractTags('모터에서 탄 냄새가 나고 차단기가 떨어짐').includes('탄내·연기'));
  assert.ok(L.extractTags('모터에서 탄 냄새가 나고 차단기가 떨어짐').includes('차단기·과부하'));
  assert.ok(L.extractTags('제품이 덜 채워짐').includes('토출·성형 불량'));
  assert.deepEqual(L.extractTags(''), []);
});
test('원인 묶기 — 다르게 적은 같은 원인', () => {
  assert.equal(L.canonCause('히터 끊김'), '히터 단선');
  assert.equal(L.canonCause('히터밴드 단선'), '히터 단선');
  assert.equal(L.canonCause('히터 단선 (실린더 2존 히터 밴드)'), '히터 단선');
  assert.equal(L.canonCause('열전대 단선'), '온도 센서 불량', '센서 단선이 히터로 가면 안 됨');
  assert.equal(L.canonCause('TC 단선'), '온도 센서 불량');
  assert.equal(L.canonCause('유압 피팅 풀림 누유'), '유압 누유');
  assert.equal(L.canonCause('실린더 씰 누유'), '유압 누유');
  assert.equal(L.canonCause('노즐 막힘 (이물)'), '금형 막힘');
  assert.equal(L.canonCause('써멀 트립'), '모터 과부하');
  assert.equal(L.canonCause('SSR 고장 (릴레이)'), 'SSR 고장', '모르는 원인은 괄호만 떼고 그대로');
  assert.equal(L.causeKey('히터 끊김'), L.causeKey('히터단선'));
});

console.log('3. 가상 샘플');
test('설비 5대 · 사례 30건 · 원인 5가지(7·6·6·6·5)', () => {
  assert.equal(db.equipments.length, 5);
  assert.deepEqual(db.equipments.map(e => e.name), ['설비 A', '설비 B', '설비 C', '설비 D', '설비 E']);
  assert.equal(db.cases.length, 30);
  const counts = Object.fromEntries(L.causeList(db.cases).map(c => [c.cause, c.count]));
  assert.deepEqual(counts, { '히터 단선': 7, '온도 센서 불량': 6, '유압 누유': 6, '모터 과부하': 6, '금형 막힘': 5 });
});
test('사례가 모두 유효 · id 겹침 없음 · 날짜·구분·비용 채움', () => {
  assert.equal(new Set(db.cases.map(c => c.id)).size, 30);
  db.cases.forEach(c => {
    assert.deepEqual(L.validateCase(c), {}, c.id);
    assert.ok(c.date && c.repairBy && c.cost != null && c.hours != null && c.tags.length, c.id + ' 빈 칸');
    assert.ok(db.equipments.some(e => e.name === c.equipment), c.id + ' 설비');
  });
});
test('회사 이름 없음 — 가상 표기만', () => {
  const txt = JSON.stringify(db) + JSON.stringify(Sample.receiptAoa());
  assert.ok(!/\uD55C\uC6B8|주식회사|\(주\)/.test(txt)); // 제출 원문의 회사 이름(유니코드로 적어 리포 검색에 걸리지 않게)
  assert.ok(db.equipments.every(e => /^설비 [A-E]$/.test(e.name) && /가상/.test(e.model)));
});

console.log('4. 증상 진단');
test('「온도가 안 올라가고 기계가 멈춤」 → 히터 단선 · 온도 센서 불량이 위', () => {
  const r = diag('온도가 안 올라가고 기계가 멈춤');
  assert.deepEqual(r.causes.slice(0, 2).map(c => c.cause).sort(), ['온도 센서 불량', '히터 단선']);
  const sensor = r.causes.find(c => c.cause === '온도 센서 불량');
  assert.equal(sensor.checks[0].text, '센서 커넥터 연결 상태 확인', '센서는 연결부터');
  const sum = r.causes.reduce((s, c) => s + c.share, 0);
  assert.ok(Math.abs(sum - 1) < 1e-9);
  assert.ok(r.sims.length <= L.DEFAULT_OPTS.topN);
  r.sims.forEach((s, i) => { if (i) assert.ok(r.sims[i - 1].score >= s.score); });
});
test('먼저 확인할 것에는 안전 단계(전원 차단·잔압)가 아니라 첫 확인 단계', () => {
  for (const q of ['온도가 안 올라가고 기계가 멈춤', '바닥에 기름이 고임', '차단기가 떨어지고 모터에서 탄 냄새']) {
    diag(q).causes.forEach(c => c.checks.forEach(x => assert.notEqual(L.stepKind(x.text), 'safety', q + ': ' + x.text)));
  }
});
test('다른 증상들 — 맞는 원인이 1위', () => {
  assert.equal(diag('기계 아래 바닥에 기름이 고임').causes[0].cause, '유압 누유');
  assert.equal(diag('차단기가 떨어지고 모터에서 탄 냄새').causes[0].cause, '모터 과부하');
  assert.equal(diag('제품이 덜 채워지고 노즐에서 수지가 안 나옴').causes[0].cause, '금형 막힘');
  assert.equal(diag('온도 표시가 왔다갔다 함').causes[0].cause, '온도 센서 불량');
});
test('조사가 달라도 찾음 — 「온도까지」 「올라감」', () => {
  assert.equal(L.stem('온도까지'), '온도');
  const r = diag('설정 온도까지 안 올라감');
  assert.ok(r.sims.length > 0);
  assert.ok(['히터 단선', '온도 센서 불량'].includes(r.causes[0].cause));
});
test('같은 설비를 고르면 그 설비 사례가 앞으로', () => {
  const a = diag('온도가 안 올라감');
  const b = diag('온도가 안 올라감', { equipment: '설비 D', equipmentType: '원료 건조기' });
  const rankD = r => r.sims.findIndex(s => s.c.equipment === '설비 D');
  assert.ok(rankD(b) >= 0 && (rankD(a) < 0 || rankD(b) <= rankD(a)));
  assert.ok(b.sims.find(s => s.c.equipment === '설비 D').sameEquipment);
});
test('태그만 골라도 진단', () => {
  const r = diag('', { tags: ['누유'] });
  assert.equal(r.causes[0].cause, '유압 누유');
});
test('비슷한 사례가 없으면 빈 결과(지어내지 않음)', () => {
  const r = diag('qwerty zzz');
  assert.equal(r.sims.length, 0);
  assert.equal(r.causes.length, 0);
  assert.equal(L.diagnose([], { text: '온도' }).causes.length, 0);
});
test('근거 수준 — 사례 1건이면 「근거 적음」', () => {
  const one = [db.cases[0]];
  const r = L.diagnose(one, { text: '실린더 온도가 안 올라감' });
  assert.equal(r.causes[0].level, 'weak');
  assert.equal(diag('온도가 안 올라가고 기계가 멈춤').causes[0].level, 'strong');
});

console.log('5. 단계 나누기·가이드북');
test('단계 종류 — 안전 · 초보 가능 · 담당자·전문가', () => {
  assert.equal(L.stepKind('설비 정지 후 전원 차단·잠금(LOTO)'), 'safety');
  assert.equal(L.stepKind('설비 정지·전원 차단 후 유압 잔압 빼기'), 'safety');
  assert.equal(L.stepKind('센서 커넥터 연결 상태 확인'), 'basic');
  assert.equal(L.stepKind('냉각팬 청소'), 'basic');
  assert.equal(L.stepKind('테스터로 히터 저항 측정'), 'expert');
  assert.equal(L.stepKind('유압 호스 교체'), 'expert');
  assert.equal(L.stepKind('노즐 분해 청소(고온 — 정비 담당)'), 'expert');
  assert.equal(L.firstCheck(db.cases[7]), '센서 커넥터 연결 상태 확인 — 반쯤 빠져 있음');
  assert.equal(L.stepCore('센서 커넥터 연결 상태 확인 — 반쯤 빠져 있음'), '센서 커넥터 연결 상태 확인');
  assert.equal(L.stepCore('노즐 분해 청소(고온 — 정비 담당)'), '노즐 분해 청소(고온 — 정비 담당)', '괄호 안의 — 는 자르지 않음');
});
test('히터 단선 카드 — 확인 순서·멈출 곳·안전·외부 기준', () => {
  const key = L.causeKey('히터 단선');
  const card = L.guideCard(L.guideDraft(db.cases, key));
  assert.equal(card.count, 7); assert.equal(card.extCount, 4); assert.equal(card.selfCount, 3);
  assert.equal(card.steps[0].text, '존별 온도 표시 확인');
  assert.ok(card.steps.every(s => s.kind !== 'safety'), '안전 단계는 「안전 먼저」 칸으로');
  assert.ok(card.stopIndex > 0);
  assert.equal(card.steps[card.stopIndex].kind, 'expert');
  assert.ok(card.steps.slice(0, card.stopIndex).every(s => s.kind === 'basic'));
  assert.ok(/LOTO/.test(card.safety[0]));
  assert.ok(card.cats.includes('전기') && card.cats.includes('고온'));
  assert.ok(card.safety.some(s => /전기 교육/.test(s)) && card.safety.some(s => /뜨겁/.test(s)));
  assert.ok(card.callExpert.some(s => /외부 수리 4건/.test(s)));
  assert.ok(card.tools.includes('테스터') && card.tools.includes('내열 장갑') && card.tools.includes('보안경'));
  assert.ok(card.parts.includes('히터 밴드'));
});
test('모든 원인 카드 — LOTO·전문가 기준·멈출 곳', () => {
  L.causeList(db.cases).forEach(x => {
    const c = L.guideCard(L.guideDraft(db.cases, x.key));
    assert.ok(/LOTO/.test(c.safety[0]), x.cause);
    assert.ok(c.callExpert.length >= 4, x.cause);
    assert.ok(c.stopIndex >= 1, x.cause + ' 초보 단계가 하나 이상 있어야 함');
    const t = L.guideToText(c);
    assert.ok(t.includes('초보자는 여기까지') && t.includes('■ 안전 먼저'), x.cause);
  });
});
test('카드 고치기 — 고친 칸만 덮고, LOTO 문장은 지켜짐', () => {
  const key = L.causeKey('유압 누유');
  const draft = L.guideDraft(db.cases, key);
  const c = L.guideCard(draft, { steps: '누유 위치 사진 찍기\n유압 호스 교체', safety: '장갑 착용' });
  assert.deepEqual(c.steps.map(s => s.text), ['누유 위치 사진 찍기', '유압 호스 교체']);
  assert.equal(c.stopIndex, 1);
  assert.ok(/LOTO/.test(c.safety[0]) && c.safety.includes('장갑 착용'));
  assert.deepEqual(c.symptoms, draft.symptoms);
  assert.deepEqual(c.edited.sort(), ['safety', 'steps']);
});
test('AI 답 → 카드 칸 (가이드 글을 그대로 붙여도 같은 단계)', () => {
  const c = L.guideCard(L.guideDraft(db.cases, L.causeKey('모터 과부하')));
  const p = L.parseGuideAnswer(L.guideToText(c));
  assert.deepEqual(L.linesOf(p.steps), c.steps.map(s => s.text));
  assert.deepEqual(L.linesOf(p.safety), c.safety);
  assert.deepEqual(L.linesOf(p.tools), c.tools);
  assert.deepEqual(L.parseGuideAnswer('그냥 문장입니다'), {});
});

console.log('6. 비용 현황');
test('합계·월별·자체 비율', () => {
  const r = L.costReport(db.cases);
  const ext = db.cases.filter(c => c.repairBy === '외부');
  assert.equal(r.total.extCount, 16);
  assert.equal(r.total.selfCount, 14);
  assert.equal(r.total.extCost, ext.reduce((s, c) => s + c.cost, 0));
  assert.deepEqual(r.monthly.map(m => m.month), ['2026-03', '2026-04', '2026-05', '2026-06', '2026-07', '2026-08', '2026-09']);
  assert.equal(r.monthly.reduce((s, m) => s + m.total, 0), 30);
  assert.equal(r.monthly.reduce((s, m) => s + m.extCost, 0), r.total.extCost);
  assert.ok(r.early.selfRatio < r.late.selfRatio, '뒤로 갈수록 자체 처리가 늘어남');
  assert.equal(r.byEquipment.reduce((s, e) => s + e.total, 0), 30);
  assert.equal(r.byCause.length, 5);
  assert.ok(r.candidates.length >= 1 && r.candidates.every(c => c.selfCount > 0 && c.extCount > 0));
});
test('기간 거르기 · 빈 달도 0으로 · 일자 없는 사례 따로 셈', () => {
  const r = L.costReport(db.cases, '2026-09', '2026-09');
  assert.equal(r.total.total, 5); assert.equal(r.total.extCost, 0); assert.equal(r.total.selfRatio, 1);
  const gap = L.costReport([L.cleanCase({ date: '2026-01-10', symptom: 'a', cause: 'b', repairBy: '외부', cost: 1 }), L.cleanCase({ date: '2026-03-10', symptom: 'a', cause: 'b', repairBy: '자체' }), L.cleanCase({ symptom: 'a', cause: 'b' })]);
  assert.deepEqual(gap.monthly.map(m => m.total), [1, 0, 1]);
  assert.equal(gap.undated, 1);
  assert.deepEqual(L.monthRange('2025-11', '2026-02'), ['2025-11', '2025-12', '2026-01', '2026-02']);
});

console.log('7. 정비 내역서 엑셀 가져오기');
function roundTrip(aoa) {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), '정비내역');
  const back = XLSX.read(XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' }), { type: 'buffer', cellDates: true });
  return XLSX.utils.sheet_to_json(back.Sheets['정비내역'], { header: 1, raw: true, defval: null });
}
test('제목 줄 아래 머리행 찾기 · 다른 열 이름 자동 연결', () => {
  const aoa = roundTrip(Sample.receiptAoa());
  const hr = L.detectHeaderRow(aoa);
  assert.equal(hr, 2);
  const map = L.autoMap(L.headersOf(aoa, hr), null);
  assert.deepEqual(map, { date: 0, equipment: 1, symptom: 2, cause: 3, steps: 4, parts: 5, repairBy: 6, cost: 7, hours: 8, memo: 9 });
  assert.deepEqual(L.missingRequired(map), []);
  assert.deepEqual(L.missingRequired({ date: 0 }), ['증상', '원인']);
});
test('내역서 30건 → 사례 30건 (단계·금액·시간·구분이 원본과 같음, 합계 줄은 건너뜀)', () => {
  const aoa = roundTrip(Sample.receiptAoa());
  const hr = L.detectHeaderRow(aoa);
  const res = L.applyMapping(aoa, hr, L.autoMap(L.headersOf(aoa, hr), null), { idPrefix: 't' });
  assert.equal(res.rows.length, 30);
  assert.equal(res.issues.length, 1);
  assert.ok(/건너뜀/.test(res.issues[0].msg));
  res.rows.forEach((c, i) => {
    const o = db.cases[i];
    assert.equal(c.date, o.date, i + ' 날짜');
    assert.deepEqual(c.steps, o.steps, i + ' 단계');
    assert.equal(c.cost, o.cost, i + ' 금액');
    assert.equal(c.hours, o.hours, i + ' 시간');
    assert.equal(c.repairBy, o.repairBy, i + ' 구분');
    assert.equal(L.canonCause(c.cause), L.canonCause(o.cause), i + ' 원인');
    assert.deepEqual(c.parts, o.parts, i + ' 부품');
    assert.deepEqual(c.tags, o.tags, i + ' 태그(자동)');
  });
});
test('같은 파일을 두 번 넣어도 늘지 않음 · 저장한 열 연결 재사용', () => {
  const aoa = roundTrip(Sample.receiptAoa());
  const hr = L.detectHeaderRow(aoa), headers = L.headersOf(aoa, hr);
  const rows = L.applyMapping(aoa, hr, L.autoMap(headers, null)).rows;
  const m1 = L.mergeCases([], rows);
  const m2 = L.mergeCases(m1.cases, rows);
  assert.equal(m1.added, 30); assert.equal(m2.added, 0); assert.equal(m2.dup, 30);
  // 이름이 별칭에 없는 열도 한 번 연결해 두면 다음에 자동으로
  const saved = L.mappingToSaved(['처리일', '호기명', '증세', '판정'], { date: 0, equipment: 1, symptom: 2, cause: 3 });
  assert.deepEqual(L.autoMap(['판정', '증세', '처리일', '호기명'], saved), { date: 2, equipment: 3, symptom: 1, cause: 0 });
});
test('못 읽은 값은 줄 번호와 함께 알림', () => {
  const aoa = [['일자', '설비', '증상', '원인', '비용', '구분', '소요시간'],
    ['어제', '설비 A', '온도 낮음', '히터 단선', '견적 대기', '모름', '반나절'],
    ['2026-05-01', '설비 B', '', '', '', '', '']];
  const res = L.applyMapping(aoa, 0, L.autoMap(aoa[0], null));
  const msgs = res.issues.map(x => x.row + ':' + x.msg).join(' | ');
  assert.ok(/2:날짜를 읽지 못함/.test(msgs) && /2:금액을 읽지 못함/.test(msgs) && /2:외부\/자체/.test(msgs) && /2:소요시간/.test(msgs), msgs);
  assert.ok(/3:증상·원인이 비어 건너뜀/.test(msgs));
  assert.equal(res.rows.length, 1);
});
test('내보내기 표 — 머리·행 수·단계 번호', () => {
  const aoa = L.casesToAoa(db.cases);
  assert.deepEqual(aoa[0], L.EXPORT_HEAD);
  assert.equal(aoa.length, 31);
  assert.ok(/^1\) /.test(aoa[1][6]));
});

console.log('8. AI 프롬프트');
test('진단 프롬프트 — 증상·사례·원인 순위·매뉴얼·안전 지시', () => {
  const r = diag('온도가 안 올라가고 기계가 멈춤', { equipment: '설비 A', equipmentType: '사출기' });
  const p = L.buildDiagnosePrompt({ text: '온도가 안 올라가고 기계가 멈춤', tags: r.tags, equipment: '설비 A', equipmentType: '사출기', manual: '[가상 매뉴얼] 열전대 확인' }, r);
  assert.ok(p.includes('설비 A (사출기)'));
  assert.ok(p.includes('[사례 1]') && p.includes('[사례 ' + r.sims.length + ']'));
  assert.ok(p.includes('히터 단선') && p.includes('[가상 매뉴얼] 열전대 확인'));
  assert.ok(/LOTO/.test(p) && /전기 측정·배선/.test(p) && /추정/.test(p));
  assert.ok(!/sk-|Bearer/.test(p));
  assert.ok(/해줘|써줘|붙여줘/.test(p), '요청형 말투');
  const empty = L.buildDiagnosePrompt({ text: 'qwerty' }, diag('qwerty'));
  assert.ok(empty.includes('비슷한 과거 사례가 없어'));
});
test('가이드 프롬프트 — 초안·근거 사례·멈춤 규칙', () => {
  const c = L.guideCard(L.guideDraft(db.cases, L.causeKey('금형 막힘')));
  const p = L.buildGuidePrompt(c, db.cases.filter(x => L.causeKey(x.cause) === c.key));
  assert.ok(p.includes('「금형 막힘」') && p.includes('[사례 5]') && p.includes('담당자·전문가'));
});

console.log('9. 저장·백업·사진');
test('저장 형식 되살리기 — 깨진 값은 빈 DB', () => {
  const back = L.restoreDb(JSON.parse(JSON.stringify(db)));
  assert.equal(back.cases.length, 30); assert.equal(back.equipments.length, 5); assert.equal(back.seq, 30);
  assert.equal(L.restoreDb(null).cases.length, 0);
  assert.equal(L.restoreDb({ cases: 'x', equipments: [{}] }).equipments.length, 0);
  assert.deepEqual(L.restoreDb({ opts: { topN: 999, minScore: 'a' } }).opts.topN, 30);
});
test('백업 JSON — 이 도구 것만 받음', () => {
  const bk = L.buildBackup(db, [{ id: 'ph1', data: 'AAAA' }, { id: 'ph2' }], '2026-10-20T00:00:00Z');
  const p = L.parseBackup(JSON.stringify(bk));
  assert.equal(p.db.cases.length, 30); assert.equal(p.photos.length, 1); assert.equal(p.exportedAt, '2026-10-20T00:00:00Z');
  assert.throws(() => L.parseBackup('{"app":"data09-17","db":{}}'), /백업 파일이 아닙니다/);
  assert.throws(() => L.parseBackup('{'), /JSON/);
});
test('사진 — 줄이기 크기·저장 공간 경고·남은 사진 정리', () => {
  assert.deepEqual(L.fitSize(4000, 3000, 1280), { w: 1280, h: 960 });
  assert.deepEqual(L.fitSize(800, 600, 1280), { w: 800, h: 600 });
  assert.equal(L.storageLevel(10 * 1048576), 'ok');
  assert.equal(L.storageLevel(60 * 1048576), 'warn');
  assert.equal(L.storageLevel(250 * 1048576), 'bad');
  assert.equal(L.storageLevel(1, { usage: 90, quota: 100 }), 'bad');
  const cs = [L.cleanCase({ symptom: 'a', cause: 'b', photos: [{ id: 'p1', bytes: 100 }, { id: 'p2', bytes: 50 }] })];
  assert.equal(L.photoBytes(cs), 150);
  assert.deepEqual(L.orphanPhotoIds(cs, ['p1', 'p2', 'p3', 'p4'], ['p4']), ['p3']);
});
test('설비 목록 맞추기 — 사례에만 있는 설비 이름을 더함', () => {
  const d = L.emptyDb();
  d.cases = [L.cleanCase({ equipment: '설비 Z', symptom: 'a', cause: 'b' })];
  assert.equal(L.syncEquipments(d), 1);
  assert.equal(L.syncEquipments(d), 0);
  assert.equal(d.equipments[0].name, '설비 Z');
});

console.log('10. AI 연결(OpenAI 호환)');
test('주소 정리·키 없으면 인증 머리글 없음·응답 읽기', () => {
  assert.equal(E.normalizeBaseUrl('https://x.example/v1/chat/completions/'), 'https://x.example/v1');
  const req = E.buildChatRequest({ baseUrl: 'https://x.example/v1', model: 'm' }, E.promptMessages('hi'));
  assert.equal(req.url, 'https://x.example/v1/chat/completions');
  assert.ok(!('Authorization' in req.init.headers));
  assert.equal(E.parseChatResponse({ choices: [{ message: { content: '답' } }] }), '답');
  assert.equal(E.STORE_KEY, 'data10-06.ai');
  assert.equal(E.toStored({ baseUrl: 'https://x/v1', model: 'm', apiKey: 'k', remember: false }).apiKey, '');
});

console.log('\n' + passed + '개 통과' + (process.exitCode ? ' · 실패 있음' : ''));
