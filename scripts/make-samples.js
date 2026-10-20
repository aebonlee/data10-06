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
