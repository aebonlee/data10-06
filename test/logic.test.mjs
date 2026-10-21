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
  assert.deepEqual(L.missingRequired({ date: 0 }), ['증상'], '2차: 원인 칸은 없어도 됨(증상·조치에서 짐작)');
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

console.log('11. 2차 — 실제 수리이력 양식(가상 값으로 같은 모양)');
const LOG = Sample.repairLogAoa();
const logRes = () => { const hr = L.detectHeaderRow(LOG); return L.applyMapping(LOG, hr, L.autoMap(L.headersOf(LOG, hr), null), { idPrefix: 'g' }); };
test('느슨한 날짜 — 주차 · 월 · 보고서 기준 · 물결 · 연도 칸', () => {
  assert.deepEqual(L.looseDate('2025-03 2주차'), { date: '2025-03-08', prec: 'week' });
  assert.deepEqual(L.looseDate('2025-04 3~4주차'), { date: '2025-04-15', prec: 'week' });
  assert.deepEqual(L.looseDate('2025-03~04'), { date: '2025-03-01', prec: 'month' });
  assert.deepEqual(L.looseDate('2026-01 (보고서 기준)'), { date: '2026-01-01', prec: 'month' });
  assert.deepEqual(L.looseDate('2026-09-07 (보고서 기준)'), { date: '2026-09-07', prec: 'day' });
  assert.deepEqual(L.looseDate('3월 2주차', 2025), { date: '2025-03-08', prec: 'week' });
  assert.deepEqual(L.looseDate('3월 2주차'), { date: '', prec: '' }, '연도가 없으면 지어내지 않음');
  assert.deepEqual(L.looseDate('미확인'), { date: '', prec: '' });
  assert.deepEqual(L.looseDate('2025-13 1주차'), { date: '', prec: '' });
});
test('비용 칸 — 미기재는 0 이 아님 · 무상 · 개당 · 견적 · 여러 금액 더하기', () => {
  assert.deepEqual(L.parseCost('미기재'), { value: null, kind: 'missing' });
  assert.deepEqual(L.parseCost('무상'), { value: 0, kind: 'free' });
  assert.deepEqual(L.parseCost('350,000원'), { value: 350000, kind: 'won' });
  assert.deepEqual(L.parseCost('40,000원/개'), { value: 40000, kind: 'unit' });
  assert.deepEqual(L.parseCost('보수 견적 2,000,000원 + 운반비 300,000원(견적 기준, 실제 지출 확인 필요)'), { value: 2300000, kind: 'estimate' });
  assert.deepEqual(L.parseCost('중고 부품 900,000원'), { value: 900000, kind: 'won' });
  assert.deepEqual(L.parseCost('15만원'), { value: 150000, kind: 'won' });
  assert.deepEqual(L.parseCost(250000), { value: 250000, kind: 'won' });
  assert.equal(L.parseCost('나중에').kind, 'unread');
  assert.equal(L.spent({ cost: 100, costKind: 'estimate' }), 0, '견적은 지출 합계에서 뺌');
});
test('수리 업체 → 외부/자체 · 처리 상태 · 처리일자에서 조치만', () => {
  assert.deepEqual(L.parseVendor('미기재'), { vendor: '', by: '' });
  assert.deepEqual(L.parseVendor('자체 조치'), { vendor: '자체 조치', by: '자체' });
  assert.deepEqual(L.parseVendor('A정비(점검 예정)'), { vendor: 'A정비', by: '외부' });
  assert.deepEqual(L.parseVendor('전기기사(업체명 미기재)'), { vendor: '전기기사', by: '외부' });
  assert.deepEqual(L.parseVendor('A/S 업체 미기재'), { vendor: 'A/S', by: '외부' });
  assert.equal(L.caseStatus('타이머 교체 완료'), '완료');
  assert.equal(L.caseStatus('교체 필요로 보고'), '진행 중');
  assert.equal(L.caseStatus('다음 주 점검 예정으로 기재, 완료 여부 미확인'), '진행 중', '「완료 여부 미확인」 은 완료가 아님');
  assert.equal(L.caseStatus('2026-03 중 교체 예정; 완료일 미기재'), '진행 중');
  assert.equal(L.caseStatus('이동식 냉풍기 설치 후 멈춤 없음'), '완료');
  assert.equal(L.caseStatus(''), '');
  assert.equal(L.actionOf('2026-02 (보고서 기준) 온도 조절계 교체 완료'), '온도 조절계 교체 완료');
  assert.equal(L.actionOf('2025-03 3주차'), '');
  assert.equal(L.actionOf('2026-03-03 헤드 우선 교체 예정 / 전체 교체 예정'), '헤드 우선 교체 예정 / 전체 교체 예정');
});
test('조치 나누기 — 소수(2.5일)·숫자 사이 화살표(85dB→76dB)는 자르지 않음', () => {
  assert.deepEqual(L.splitSteps('교체; 2.5일 가동 중지'), ['교체', '2.5일 가동 중지']);
  assert.deepEqual(L.splitSteps('감속기 교체 완료, 소음 85dB→76dB'), ['감속기 교체 완료, 소음 85dB→76dB']);
  assert.deepEqual(L.splitSteps('전원 차단 → 센서 확인'), ['전원 차단', '센서 확인']);
});
test('원인 짐작 — 「A로 B」 의 A · 고장 낱말에 가까운 부위 · 넓은 말은 마지막', () => {
  const cases = [
    ['릴레이 접점 불량으로 공급 멈춤', '', '릴레이 불량'],
    ['타이머 교체 중 합선, 구동 모터 소손', '', '모터 고장·소손'],
    ['냉각가스 배관 누유', '', '냉각 계통 이상(냉매·누유)'],
    ['냉각기 누유 발생', '', '냉각 계통 이상(냉매·누유)'],
    ['호퍼 입구 이물 막힘으로 온도 상승, 설비 정지', '', '호퍼·투입구 막힘'],
    ['모터를 바꾼 뒤에도 서보 에러; 서보드라이버는 점검해도 이상 없음', '', '서보 드라이버·서보 이상'],
    ['돌릴 때 소리가 커짐', '모터 베어링에 그리스 주입, 소음 개선', '베어링·롤러 마모'],
    ['스크류 헤드 파손', '', '스크류·바렐 손상'],
    ['걸이 부품 파손', '', '소모 부품·구조물 파손'],
    ['여름철 고온으로 작동 멈춤', '', '작업 환경(고온·먼지)'],
    ['2존 온도 상승 불량', '온도 조절계 교체 완료', '제어 기판·계기 고장'],
    ['수리 필요(견적 요청함)', '', '원인 미상'],
    ['누수 발생', '', '펌프 고장·누수'],
    ['절단 센서 불량', '', '위치·동작 센서 불량']
  ];
  cases.forEach(([sym, act, want]) => assert.equal(L.inferCause(sym, act, ''), want, sym));
  // 1단계 원인 묶기는 그대로(센서 단선이 히터로 가지 않음 등)
  assert.equal(L.canonCause('열전대 단선'), '온도 센서 불량');
  assert.equal(L.canonCause('냉각 계통 이상(냉매·누유)'), '냉각 계통 이상(냉매·누유)', '묶음 이름은 자기 자신으로');
});
test('설비 종류 — 「압출기 1호기」 「압출 1호기」 「압출기」 는 한 종류, 냉각기가 압출기보다 먼저', () => {
  assert.equal(L.equipFamily('압출기 1호기'), '압출기');
  assert.equal(L.equipFamily('압출 1호기'), '압출기');
  assert.equal(L.equipFamily('압출 냉각기 1호기'), '냉각기');
  assert.equal(L.equipFamily('사출 1호기(구형)'), '사출기');
  assert.equal(L.equipFamily('스프링 환풍기'), '환풍기');
  assert.equal(L.equipFamily('설비 A'), '설비 A');
  assert.equal(L.equipUnit('사출기 2호기'), '2호기');
});
test('지금 양식 그대로 읽기 — 머리행 · 열 자동 연결 · 원인 없이 26건', () => {
  const hr = L.detectHeaderRow(LOG);
  assert.equal(hr, 0);
  const map = L.autoMap(L.headersOf(LOG, hr), null);
  assert.deepEqual(map, { year: 0, equipment: 1, date: 2, symptom: 3, done: 4, vendor: 5, cost: 6, memo: 7, ref: 8 });
  assert.deepEqual(L.missingRequired(map), []);
  const res = logRes();
  assert.equal(res.rows.length, Sample.LOG_ROWS.length);
  assert.deepEqual(res.issues, []);
  assert.ok(res.rows.every(c => c.causeInferred && c.cause), '원인은 모두 짐작');
  assert.equal(res.rows.filter(c => L.isUnknownCause(c.cause)).length, 0);
});
test('조치는 비고(앞 해) → 처리일자의 조치 글(뒤 해)에서, 일자는 처리일자에서도', () => {
  const r = logRes().rows;
  const by = sym => r.find(c => c.symptom === sym);
  assert.deepEqual(by('타이머 고장').steps, ['타이머 교체 완료']);
  assert.equal(by('타이머 고장').memo, '', '비고를 조치로 옮겼으면 비고는 비움');
  assert.deepEqual(by('2존 온도 상승 불량').steps, ['온도 조절계 교체 완료']);
  assert.deepEqual(by('스크류 헤드 파손').steps, ['헤드 우선 교체 예정', '전체 스크류 교체 예정']);
  assert.equal(by('펌프 쪽 누수').datePrec, 'month');
  assert.equal(by('호퍼 입구 이물 막힘으로 온도 상승, 설비 정지').date, '2026-08-19');
  assert.equal(by('벨트 마모로 속도 불균일').date, '2025-06-08');
  const precs = r.reduce((m, c) => (m[c.datePrec] = (m[c.datePrec] || 0) + 1, m), {});
  assert.deepEqual(precs, { month: 9, week: 14, day: 3 });
});
test('외부/자체 · 비용 구분 · 처리 상태 개수', () => {
  const r = logRes().rows;
  const cnt = f => r.reduce((m, c) => (m[f(c) || '-'] = (m[f(c) || '-'] || 0) + 1, m), {});
  assert.deepEqual(cnt(c => c.repairBy), { '외부': 9, '-': 15, '자체': 2 });
  assert.deepEqual(cnt(c => c.costKind), { missing: 15, won: 8, unit: 1, free: 1, estimate: 1 });
  assert.deepEqual(cnt(c => c.status), { '완료': 18, '진행 중': 7, '-': 1 });
  assert.equal(r.find(c => c.symptom === '모터 쇼트로 정지').repairBy, '외부', '「전기기사」 는 외부');
  assert.equal(r.find(c => c.symptom === '펌프 쪽 누수').repairBy, '외부');
});
test('병합 셀 풀기 · 같은 머리 모양 시트 함께 · 요약 시트는 빼기', () => {
  const aoa = [['설비명', '증상', '비용'], ['압출기 1호기', '벨트 마모', '1원'], [null, '타이머 고장', null]];
  const u = L.unmerge(aoa, [{ s: { r: 1, c: 0 }, e: { r: 2, c: 0 } }]);
  assert.equal(u[2][0], '압출기 1호기');
  assert.equal(aoa[2][0], null, '원본은 그대로');
  const plans = L.sheetPlans([{ name: '2025', aoa: LOG }, { name: '2026', aoa: LOG }, { name: '요약', aoa: Sample.repairLogSummaryAoa() }]);
  assert.deepEqual(plans.map(p => p.name), ['2025', '2026']);
});
test('중복·연속 보고 의심 — 같은 금액 쌍은 비용을 한 번만 셈(끄면 두 번)', () => {
  const r = logRes().rows;
  const d = L.dupSuspects(r);
  const same = d.filter(x => x.sameCost);
  assert.equal(same.length, 1);
  assert.equal(same[0].cost, 1500000);
  const on = L.costReport(r), off = L.costReport(r, '', '', { dedupe: false });
  assert.equal(off.total.allCost - on.total.allCost, 1500000);
  assert.equal(on.dupCost, 1500000);
});
test('비용 현황 — 구분 모름 · 견적을 따로, 종류·부위·업체별', () => {
  const rep = L.costReport(logRes().rows);
  const T = rep.total;
  assert.equal(T.total, 26);
  assert.equal(T.unknownCount, 15);
  assert.equal(T.estCount, 1); assert.equal(T.estCost, 2000000);
  assert.equal(T.allCost, T.extCost + T.selfCost + T.unknownCost);
  assert.equal(T.allCost, 180000 + 25000 + 60000 + 1200000 + 45000 + 250000 + 1500000 + 0 + 900000, '적힌 실지출만, 같은 금액 중복 한 번');
  assert.ok(rep.byFamily.some(x => x.label === '냉각기' && x.total === 5), '「냉각기 1호기 순환펌프」 도 냉각기');
  assert.ok(rep.byPart.some(x => x.label === '냉각'));
  assert.ok(rep.byVendor.some(x => x.label === '(업체 미기재)'));
});
test('진단 — 부위 낱말 가산 · 같은 뜻 말 · 원인 미상은 순위에서 뺌', () => {
  const r = logRes().rows;
  assert.equal(L.diagnose(r, { text: '냉각기에서 기름이 샘' }).causes[0].cause, '냉각 계통 이상(냉매·누유)');
  assert.equal(L.diagnose(r, { text: '서보 알람 뜨고 멈춤' }).causes[0].cause, '서보 드라이버·서보 이상');
  assert.equal(L.diagnose(r, { text: '타이머가 안 됨' }).causes[0].cause, '타이머 고장');
  assert.ok(L.expandText('작동 불가').includes('멈춤'));
  // 증상에 적은 부위(서보)가 원인과 같으면 가산 — 없으면 「알람·멈춤」 이 겹친 다른 원인이 1위가 됨
  const toy = [L.cleanCase({ id: 'a', symptom: '서보드라이버 고장', cause: '서보드라이버 고장' }), L.cleanCase({ id: 'b', symptom: '알람 뜨고 멈춤', cause: '제어 기판 고장' }), L.cleanCase({ id: 'c', symptom: '벨트 마모', cause: '벨트 마모' })];
  assert.equal(L.diagnose(toy, { text: '서보 알람 뜨고 멈춤' }).causes[0].cause, '서보 드라이버·서보 이상');
  assert.equal(L.diagnose(toy, { text: '서보 알람 뜨고 멈춤' }, { partBonus: 0 }).causes[0].cause, '제어 기판·계기 고장');
  const unk = r.concat([L.cleanCase({ id: 'u1', symptom: '벨트 쪽 이상', cause: '원인 미상' })]);
  const dg = L.diagnose(unk, { text: '벨트 쪽 이상' });
  assert.ok(dg.sims.some(x => x.c.id === 'u1'), '비슷한 사례 표에는 보임');
  assert.ok(!dg.causes.some(x => L.isUnknownCause(x.cause)), '원인 순위에는 없음');
});
test('먼저 확인할 것·가이드 — 「예정·보고」 같은 진행 기록은 빼고, 한 줄 조치는 지난 조치 기록으로', () => {
  assert.equal(L.stepKind('A/S 방문 예정'), 'note');
  assert.equal(L.stepKind('교체 필요로 보고'), 'note');
  assert.equal(L.firstCheck({ steps: ['A/S 방문 예정', '서보드라이버 교체'] }), '서보드라이버 교체');
  const r = logRes().rows;
  const key = L.causeKey('타이머 고장');
  const card = L.guideCard(L.guideDraft(r, key));
  assert.equal(card.steps.length, 0, '두 단계 이상 적힌 사례가 없으면 순서를 지어내지 않음');
  assert.deepEqual(card.actions.map(a => a.text), ['타이머 교체 완료']);
  const txt = L.guideToText(card);
  assert.ok(/지난 조치 기록/.test(txt));
  assert.ok(!L.parseGuideAnswer(txt).steps, '되읽어도 지난 조치가 확인 순서로 들어가지 않음');
});
test('자료 점검 — 칸별 채움 · 기능별 상태 · 더 채울 칸(원인 먼저)', () => {
  const chk = L.dataCheck(logRes().rows);
  const F = Object.fromEntries(chk.fields.map(f => [f.id, f]));
  assert.equal(chk.n, 26);
  assert.equal(F.symptom.filled, 26);
  assert.equal(F.cause.filled, 0); assert.equal(F.cause.inferred, 26);
  assert.equal(F.repairBy.filled, 11);
  assert.equal(F.cost.filled, 10, '금액 8 + 무상 1 + 개당 1 (견적·미기재 제외)');
  assert.equal(chk.todo[0].id, 'cause');
  assert.deepEqual(chk.todo.map(f => f.id).sort(), ['cause', 'cost', 'parts', 'repairBy'].sort());
  assert.equal(chk.features[0].level, '됨');
  assert.equal(chk.features[2].level, '일부');
  assert.equal(chk.dupSameCost.length, 1);
  assert.equal(chk.estimates.length, 1); assert.equal(chk.unitPrices.length, 1);
  // 1단계 예시(칸이 다 찬 30건)는 필수 칸이 다 차 있음(자체 수리 건은 교체 부품이 없어 「교체 부품」 만 남음)
  const full = L.dataCheck(db.cases);
  assert.deepEqual(full.todo.map(f => f.id), ['parts']);
  assert.equal(full.features[0].level, '좋음');
});
test('권장 양식 — 예시 2줄 다시 읽기 · 내 자료 내보내기(짐작한 원인은 참고 칸에만)', () => {
  const rec = L.recommendedAoa();
  const m = L.autoMap(L.headersOf(rec, 0), null);
  assert.equal(Object.keys(m).length, L.RECOMMENDED_HEAD.length, '권장 양식 칸은 모두 자동 연결');
  const rr = L.applyMapping(rec, 0, m).rows;
  assert.equal(rr[0].steps.length, 5); assert.equal(rr[0].causeInferred, false);
  assert.equal(rr[1].costKind, 'estimate'); assert.equal(rr[1].status, '진행 중');
  const out = L.toRecommendedAoa(logRes().rows);
  assert.equal(out.length, 27);
  assert.equal(out[1][3], '', '짐작한 원인은 원인 칸에 쓰지 않음');
  assert.ok(out[1][out[0].length - 1], '참고 칸에 짐작한 원인');
  const back = L.applyMapping(out, 0, L.autoMap(L.headersOf(out, 0), null)).rows;
  assert.equal(back.length, 26);
  assert.ok(back.every(c => c.causeInferred), '다시 가져와도 짐작은 짐작');
  assert.equal(back.find(c => c.symptom === '벨트 마모로 속도 불균일').datePrec, 'week', '주차 표기가 유지됨');
});
test('사례 저장·되살리기 — 2차 칸(업체·상태·비용 구분·일자 정밀도·짐작)이 남음', () => {
  const r = logRes().rows;
  const d2 = L.restoreDb(JSON.parse(JSON.stringify({ cases: r })));
  ['vendor', 'status', 'costKind', 'datePrec', 'causeInferred', 'cost', 'date', 'ref'].forEach(k =>
    assert.deepEqual(d2.cases.map(c => c[k]), r.map(c => c[k]), k));
});
test('가상 수리이력 — 실제 이름 없이 지어낸 표기', () => {
  const txt = JSON.stringify(Sample.LOG_ROWS);
  assert.ok(/가상/.test(txt));
  assert.ok(Sample.LOG_ROWS.every(r => /^(미기재|자체 조치|[A-D][가-힣]+(\(점검 예정\))?|전기기사\(업체명 미기재\)|A\/S 업체 미기재)$/.test(r[5])), '업체는 A~D 가짜 이름·미기재·자체만');
});

console.log('\n' + passed + '개 통과' + (process.exitCode ? ' · 실패 있음' : ''));
