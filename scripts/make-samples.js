// 가상 샘플 파일 만들기: node scripts/make-samples.js
// js/sample-data.js 의 가상 정비 사례 30건을 「정비 내역서」 모양의 .xlsx 로 씁니다(전부 지어낸 값).
//  · 정비내역 시트 — 제목 줄 · 회사마다 다른 열 이름(수리일자·장비명·고장 내용 …) · 조치는 「1) 2)」 한 칸 · 금액은 글자 · 합계 줄
//  · 설비목록 시트 — 설비 A~E 와 종류(가상)
// 날짜 칸은 엑셀 날짜로 넣고, 방금 쓴 파일을 앱과 같은 방식으로 다시 읽어 원본과 같은지 확인합니다.
const fs = require('fs');
const path = require('path');
const XLSX = require('../vendor/xlsx.full.min.js');
const L = require('../js/logic.js');
const Sample = require('../js/sample-data.js');

const out = path.join(__dirname, '..', 'samples');
fs.mkdirSync(out, { recursive: true });
const file = path.join(out, '예시데이터_정비내역서.xlsx');

const aoa = Sample.receiptAoa().map((row, i) => i < 3 ? row : row.map((v, c) => (c === 0 && typeof v === 'string' && /^\d{4}-\d\d-\d\d$/.test(v)) ? new Date(v + 'T00:00:00') : v));
const wb = XLSX.utils.book_new();
const ws = XLSX.utils.aoa_to_sheet(aoa, { cellDates: true, dateNF: 'yyyy-mm-dd' });
ws['!cols'] = [12, 10, 40, 28, 70, 22, 10, 14, 12, 30].map(w => ({ wch: w }));
ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 9 } }];
XLSX.utils.book_append_sheet(wb, ws, '정비내역');
const eq = [['설비', '종류', '모델', '제조사', '비고']].concat(Sample.EQUIPMENTS.map(e => [e.name, e.type, e.model, '(가상)', e.memo]));
XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(eq), '설비목록');
fs.writeFileSync(file, XLSX.write(wb, { bookType: 'xlsx', type: 'buffer', cellDates: true }));

// 다시 읽어 확인
const back = XLSX.read(fs.readFileSync(file), { type: 'buffer', cellDates: true });
const rows = XLSX.utils.sheet_to_json(back.Sheets['정비내역'], { header: 1, raw: true, defval: null });
const hr = L.detectHeaderRow(rows);
const map = L.autoMap(L.headersOf(rows, hr), null);
const res = L.applyMapping(rows, hr, map);
const orig = Sample.cases();
if (res.rows.length !== orig.length) throw new Error('행 수가 다릅니다: ' + res.rows.length);
res.rows.forEach((c, i) => {
  const o = orig[i];
  ['date', 'cost', 'hours', 'repairBy'].forEach(k => { if (c[k] !== o[k]) throw new Error((i + 1) + '행 ' + k + ': ' + c[k] + ' ≠ ' + o[k]); });
  if (JSON.stringify(c.steps) !== JSON.stringify(o.steps)) throw new Error((i + 1) + '행 단계가 다릅니다');
});
console.log(path.relative(process.cwd(), file), '— 머리행', hr + 1, '· 사례', res.rows.length, '건 · 확인할 줄', res.issues.length, '(합계 줄)');

// ── 2차: 실제 수리이력과 같은 모양의 가상 파일 + 권장 양식 ──────────────
//  · 예시데이터_수리이력_지금양식.xlsx — 「통합 수리이력」(연도·설비명·일자·증상·처리일자·수리업체·비용·비고·출처) + 「요약」(첫 줄 병합 제목)
//  · 수리이력_권장양식.xlsx — 지금 양식에 원인·조치 순서·교체 부품·외부/자체 등을 더한 머리 + 지어낸 예시 2줄
{
  const file2 = path.join(out, '예시데이터_수리이력_지금양식.xlsx');
  const wb2 = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(Sample.repairLogAoa());
  ws['!cols'] = [6, 22, 22, 40, 44, 22, 30, 40, 22].map(w => ({ wch: w }));
  XLSX.utils.book_append_sheet(wb2, ws, '통합 수리이력');
  const sum = XLSX.utils.aoa_to_sheet(Sample.repairLogSummaryAoa());
  sum['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 3 } }];
  XLSX.utils.book_append_sheet(wb2, sum, '요약');
  fs.writeFileSync(file2, XLSX.write(wb2, { bookType: 'xlsx', type: 'buffer' }));
  const back2 = XLSX.read(fs.readFileSync(file2), { type: 'buffer', cellDates: true });
  const sheets = back2.SheetNames.map(n => ({ name: n, aoa: L.unmerge(XLSX.utils.sheet_to_json(back2.Sheets[n], { header: 1, raw: true, defval: null }), back2.Sheets[n]['!merges']) }));
  const plans = L.sheetPlans(sheets, null);
  if (plans.length !== 1 || plans[0].name !== '통합 수리이력') throw new Error('수리이력 시트를 고르지 못했습니다: ' + plans.map(p => p.name));
  const res2 = L.applyMapping(sheets[0].aoa, plans[0].hr, plans[0].map);
  if (res2.rows.length !== Sample.LOG_ROWS.length) throw new Error('수리이력 행 수가 다릅니다: ' + res2.rows.length);
  console.log(path.relative(process.cwd(), file2), '— 사례', res2.rows.length, '건 · 원인 짐작', res2.rows.filter(c => c.causeInferred).length, '건');

  const file3 = path.join(out, '수리이력_권장양식.xlsx');
  const wb3 = XLSX.utils.book_new();
  const ws3 = XLSX.utils.aoa_to_sheet(L.recommendedAoa());
  ws3['!cols'] = L.RECOMMENDED_HEAD.map(h => ({ wch: Math.max(10, h.length * 2 + 4) }));
  XLSX.utils.book_append_sheet(wb3, ws3, '수리이력(권장 양식)');
  fs.writeFileSync(file3, XLSX.write(wb3, { bookType: 'xlsx', type: 'buffer' }));
  const a3 = XLSX.utils.sheet_to_json(XLSX.read(fs.readFileSync(file3), { type: 'buffer' }).Sheets['수리이력(권장 양식)'], { header: 1, raw: true, defval: null });
  const m3 = L.autoMap(L.headersOf(a3, 0), null);
  const r3 = L.applyMapping(a3, 0, m3);
  if (r3.rows.length !== 2 || r3.rows.some(c => c.causeInferred) || r3.rows[1].costKind !== 'estimate') throw new Error('권장 양식을 다시 읽지 못했습니다');
  console.log(path.relative(process.cwd(), file3), '— 머리', L.RECOMMENDED_HEAD.length, '칸 · 다시 읽어 연결', Object.keys(m3).length, '칸');
}
