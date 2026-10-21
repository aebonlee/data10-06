// 헤드리스 브라우저로 주요 흐름 점검 (선택 — playwright-core 가 있는 PC 에서만)
//   PLAYWRIGHT_CORE=/경로/node_modules/playwright-core node scripts/check-browser.mjs
// file:// 로 index.html 을 열어 ① 예시 시작 → 8개 화면(1280·390px 가로 넘침) → 진단
// ② 「예시 수리이력(지금 양식 모양)」 가져오기 → 미리 보기 26건 · 자료 점검 · 비용 현황(구분 모름·중복 의심) · 진단(서보) 순서로 누르고 JS 오류를 잽니다.
// FILE=엑셀경로 를 주면 그 파일을 가져와 읽히는지·화면이 그려지는지만 봅니다(내 PC 의 실제 자료 점검용 — 결과에 값은 찍지 않고 건수만).
// BREAK=1 이면 일부러 넓은 상자를 넣어 넘침 검사가 잡는지 봅니다.
import { createRequire } from 'node:module';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
const require = createRequire(import.meta.url);
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const { chromium } = require(process.env.PLAYWRIGHT_CORE || 'playwright-core');
const url = 'file://' + path.join(root, 'index.html');
const VIEWS = ['diagnose', 'cases', 'guide', 'cost', 'import', 'check', 'settings', 'next'];

const errors = [], problems = [];
const browser = await chromium.launch();
const ctx = await browser.newContext({ acceptDownloads: true, viewport: { width: 1280, height: 900 } });
const page = await ctx.newPage();
page.on('pageerror', e => errors.push('pageerror: ' + e.message));
page.on('console', m => { if (m.type() === 'error') errors.push('console: ' + m.text()); });
page.on('dialog', d => d.accept());
function ok(cond, msg) { if (!cond) { problems.push(msg); console.log('  FAIL ' + msg); } else console.log('  ok  ' + msg); }
async function go(hash) { await page.evaluate(h => { location.hash = h; }, hash); await page.waitForTimeout(250); }
async function overflow() {
  return page.evaluate(() => {
    const W = document.documentElement.clientWidth, bad = [];
    document.querySelectorAll('main *').forEach(el => {
      const r = el.getBoundingClientRect(); if (!r.width || !r.height) return;
      let p = el.parentElement, inScroll = false;
      while (p) { const o = getComputedStyle(p).overflowX; if (o === 'auto' || o === 'scroll') { inScroll = true; break; } p = p.parentElement; }
      if (!inScroll && r.right > W + 1) bad.push(el.tagName + '.' + el.className);
    });
    return bad.slice(0, 3);
  });
}
async function sweep(tag) {
  for (const id of VIEWS) {
    for (const w of [1280, 390]) {
      await page.setViewportSize({ width: w, height: 900 });
      await go('#/' + id);
      const txt = await page.textContent('main');
      const bad = await overflow();
      ok(!/화면을 그리지 못했습니다/.test(txt) && !bad.length, tag + ' ' + id + ' ' + w + 'px 그려짐·가로 넘침 없음' + (bad.length ? ' — ' + bad.join(', ') : ''));
    }
  }
  await page.setViewportSize({ width: 1280, height: 900 });
}
async function diagnose(text) {
  await go('#/diagnose');
  await page.fill('textarea[aria-label="증상"]', text);
  await page.locator('button', { hasText: '원인 짚어 보기' }).click();
  await page.waitForTimeout(300);
  return page.locator('#diagResult .cause-name').allTextContents();
}

await page.goto(url);
if (process.env.FILE) {
  await go('#/import');
  await page.setInputFiles('input[type=file][accept=".xlsx,.xls,.xlsm,.csv"]', process.env.FILE);
  await page.waitForTimeout(1500);
  const pv = await page.textContent('main');
  const m = /읽은 사례 (\d+)건/.exec(pv);
  ok(!!m && Number(m[1]) > 0, '실제 파일 읽힘 — 사례 ' + (m ? m[1] : 0) + '건');
  ok(/이 파일의 자료 점검/.test(pv), '미리 보기에 자료 점검 요약');
  await page.locator('button', { hasText: /^사례 \d+건 더하기/ }).click();
  await page.waitForTimeout(600);
  ok(/칸별 채움/.test(await page.textContent('main')), '가져온 뒤 자료 점검 화면으로');
  await sweep('실제');
  const top = await diagnose('냉각기에서 누유');
  ok(top.length > 0, '실제 자료로 진단 결과 ' + top.length + '개 원인');
} else {
  if (process.env.BREAK) await page.addStyleTag({ content: '.page-head { min-width: 900px; }' });
  await page.locator('button', { hasText: '예시 데이터로 시작' }).first().click();
  await page.waitForTimeout(500);
  ok(await page.locator('#sampleBanner').isVisible(), '예시 띠 보임');
  const top = await diagnose('온도가 안 올라가고 기계가 멈춤');
  ok(top[0] === '히터 단선' && top.includes('온도 센서 불량'), '예시 진단: 히터 단선 · 온도 센서 불량');
  await sweep('예시');

  // 2차 — 지금 양식 모양의 가상 수리이력 가져오기
  await go('#/import');
  await page.locator('button', { hasText: '예시 수리이력(지금 양식 모양)으로 시험' }).click();
  await page.waitForTimeout(600);
  const pv = await page.textContent('main');
  ok(/읽은 사례 26건/.test(pv), '수리이력 미리 보기 26건');
  ok(/짐작/.test(pv), '원인 「짐작」 표시');
  ok(/2025-03 2주차/.test(pv) || /2025-03/.test(pv), '주차·월 일자 표기');
  ok(/이 파일의 자료 점검/.test(pv) && /더 채우면 좋은 칸/.test(pv), '미리 보기에 자료 점검 요약');
  ok(!/확인할 줄 \d+개/.test(pv), '못 읽은 줄 없음');
  await page.locator('button', { hasText: /^사례 26건 더하기/ }).click();
  await page.waitForTimeout(600);
  const chk = await page.textContent('main');
  ok(/칸별 채움/.test(chk) && /원인\(점검 결과\)/.test(chk), '자료 점검 화면 — 칸별 채움');
  ok(/같은 금액 1쌍/.test(chk), '자료 점검 — 같은 금액 중복 의심 1쌍');
  ok(!(await page.locator('#sampleBanner').isVisible()), '예시 띠 사라짐(예시 지우고 가져옴)');
  ok((await page.textContent('#dataChip')).includes('26건'), '사례 26건');
  await go('#/cost');
  const cost = await page.textContent('main');
  ok(/구분 모름/.test(cost) && /견적\(실지출 아님\)/.test(cost) && /한 번만 셈/.test(cost), '비용 현황 — 구분 모름 · 견적 · 중복 한 번만');
  ok(/설비 종류별/.test(cost) && /부위별/.test(cost) && /수리 업체별/.test(cost), '비용 현황 — 종류·부위·업체별 표');
  const top2 = await diagnose('서보 알람 뜨고 멈춤');
  ok(top2[0] === '서보 드라이버·서보 이상', '수리이력 진단: 서보 1위 (' + top2[0] + ')');
  await go('#/guide');
  ok(/지난 조치 기록/.test(await page.textContent('main')), '가이드북 — 지난 조치 기록');
  await sweep('수리이력');
  // 사례 고치기 — 원인을 고치면 짐작 표시가 사라짐
  await go('#/cases');
  await page.locator('button', { hasText: '고치기' }).first().click();
  await page.waitForTimeout(200);
  ok(await page.locator('dialog select').count() >= 3, '고치기 창 — 비용 구분·처리 상태 고르기');
  await page.locator('dialog button', { hasText: '취소' }).click();
}
ok(!errors.length, 'JS 오류 없음' + (errors.length ? ' — ' + errors.slice(0, 3).join(' | ') : ''));
await browser.close();
console.log(problems.length ? '\n' + problems.length + '개 문제' : '\n모두 통과');
process.exitCode = problems.length ? 1 : 0;
