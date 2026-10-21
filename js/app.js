/* 설비고장 원인 진단 도우미 — 화면
   계산은 js/logic.js(FixLogic), 그래프는 js/charts.js, 글 저장은 js/store.js(localStorage), 사진은 js/photos.js(IndexedDB),
   AI 연결은 js/ai-endpoint.js·ai-panel.js(data09-25 에서 옮김). */
(function () {
  'use strict';
  var L = window.FixLogic, S = window.FixStore, P = window.FixPhotos, C = window.FixCharts, AIP = window.AIPanel, SMP = window.FixSample;
  var db = S.loadDb() || L.emptyDb();
  var main = document.getElementById('main');
  var ui = {
    diag: { equipment: '', text: '', tags: [], tagTouched: {}, result: null, manual: '', manualFor: null, answer: '' },
    cases: { equipment: '', cause: '', by: '', q: '', from: '', to: '' },
    imp: null,
    cost: { from: '', to: '', dedupe: true }
  };

  /* ── 작은 도구 ───────────────────────────── */
  function h(tag, attrs) {
    var e = document.createElement(tag);
    if (attrs) Object.keys(attrs).forEach(function (k) {
      var v = attrs[k];
      if (v == null || v === false) return;
      if (k === 'class') e.className = v;
      else if (k.slice(0, 2) === 'on' && typeof v === 'function') e.addEventListener(k.slice(2), v);
      else if (v === true) e.setAttribute(k, '');
      else e.setAttribute(k, v);
    });
    for (var i = 2; i < arguments.length; i++) add(e, arguments[i]);
    return e;
  }
  function add(e, c) {
    if (c == null || c === false) return;
    if (Array.isArray(c)) { c.forEach(function (x) { add(e, x); }); return; }
    e.appendChild(typeof c === 'object' ? c : document.createTextNode(String(c)));
  }
  function fld(label, control, hint, cls) {
    return h('label', { class: 'field' + (cls ? ' ' + cls : '') }, h('span', null, label), control, hint ? h('small', { class: 'hint' }, hint) : null);
  }
  function sel(options, value, attrs) {
    var s = h('select', attrs || null);
    options.forEach(function (o) {
      var v = Array.isArray(o) ? o[0] : o, t = Array.isArray(o) ? o[1] : o;
      var op = h('option', { value: v }, t); if (String(v) === String(value == null ? '' : value)) op.selected = true; s.appendChild(op);
    });
    return s;
  }
  var toastTimer = null;
  function toast(msg, err) {
    var t = document.getElementById('toast');
    t.textContent = msg; t.className = 'toast' + (err ? ' error' : ''); t.hidden = false;
    clearTimeout(toastTimer); toastTimer = setTimeout(function () { t.hidden = true; }, 3800);
  }
  function save() {
    var ok = S.saveDb(db);
    var b = document.getElementById('storeBanner');
    b.hidden = ok;
    b.textContent = ok ? '' : '브라우저 저장소에 저장하지 못했습니다(막혀 있거나 가득 참). 이 창을 닫으면 사라지니 「설정 → 백업 내려받기」 로 받아 두세요.';
    topbar();
    return ok;
  }
  function download(name, blob) {
    var a = h('a', { href: URL.createObjectURL(blob), download: name });
    document.body.appendChild(a); a.click(); setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 500);
  }
  function today() { var d = new Date(); return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0'); }
  function copyText(text, what) {
    function fallback() {
      var ta = h('textarea', { style: 'position:fixed;left:-9999px' }); ta.value = text; document.body.appendChild(ta); ta.select();
      try { document.execCommand('copy'); toast((what || '글') + '을 복사했습니다'); } catch (e) { toast('복사하지 못했습니다. 글을 직접 골라 복사해 주세요', true); }
      ta.remove();
    }
    if (navigator.clipboard && navigator.clipboard.writeText) navigator.clipboard.writeText(text).then(function () { toast((what || '글') + '을 복사했습니다'); }, fallback);
    else fallback();
  }
  function writeXlsx(name, sheets) {
    var wb = XLSX.utils.book_new();
    Object.keys(sheets).forEach(function (n) { XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(sheets[n]), n.slice(0, 31)); });
    var out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    download(name, new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
  }
  function pageHead(stage, title, lead, right) {
    return h('div', { class: 'page-head' }, h('div', { class: 'titles' }, h('div', { class: 'stage' }, stage), h('h1', null, title), lead ? h('p', { class: 'lead' }, lead) : null), right || null);
  }
  function tile(k, v, sub) { return h('div', { class: 'tile' }, h('div', { class: 'k' }, k), h('div', { class: 'v' }, v), sub ? h('div', { class: 'note' }, sub) : null); }
  function table(head, rows, opts) {
    opts = opts || {};
    var t = h('table', { class: 'list' + (opts.cls ? ' ' + opts.cls : '') });
    var tr = h('tr');
    head.forEach(function (x, i) { tr.appendChild(h('th', { class: opts.num && opts.num[i] ? 'num' : null, scope: 'col' }, x)); });
    t.appendChild(h('thead', null, tr));
    var tb = h('tbody');
    rows.forEach(function (r) {
      var row = h('tr', { class: r.cls || null });
      (r.cells || r).forEach(function (x, i) { row.appendChild(h('td', { class: opts.num && opts.num[i] ? 'num' : null }, x)); });
      tb.appendChild(row);
    });
    t.appendChild(tb);
    return h('div', { class: 'table-wrap' }, t);
  }
  function equipTypes() { return L.equipTypes(db); }
  function equipOf(name) { return db.equipments.filter(function (e) { return e.name === name; })[0] || null; }
  function caseById(id) { return db.cases.filter(function (c) { return c.id === id; })[0] || null; }
  function nextId() { db.seq = (db.seq || 0) + 1; return 'c' + db.seq; }
  /* 2차 — 실제 수리이력 표기: 「2025-03 2주차」 · 비용 미기재/견적/개당 · 원인 짐작 · 처리 상태 */
  function dateText(c) {
    if (!c.date) return '날짜 없음';
    if (c.datePrec === 'month') return c.date.slice(0, 7);
    if (c.datePrec === 'week') return c.date.slice(0, 7) + ' ' + (Math.floor((Number(c.date.slice(8, 10)) - 1) / 7) + 1) + '주차';
    return c.date;
  }
  function costText(c) {
    var k = c.costKind || (c.cost == null ? 'missing' : 'won');
    if (k === 'missing') return '미기재';
    if (k === 'free') return '무상';
    if (k === 'unread') return '못 읽음';
    if (k === 'estimate') return '견적 ' + L.fmtWon(c.cost);
    if (k === 'unit') return L.fmtWon(c.cost) + '/개';
    return L.fmtWon(c.cost);
  }
  function causeCell(c) {
    return h('span', null, L.canonCause(c.cause), c.causeInferred ? h('span', { class: 'kind meta', title: '원인 칸이 없어 증상·조치 글에서 짐작했습니다' }, '짐작') : null,
      !c.causeInferred && L.canonCause(c.cause) !== c.cause ? h('small', { class: 'note block' }, c.cause) : null);
  }
  function statusBadge(c) { return c.status ? h('span', { class: 'kind ' + (c.status === '완료' ? 'ok' : 'obs') }, c.status) : null; }
  function byBadge(c) { return c.repairBy ? h('span', { class: 'kind ' + (c.repairBy === '외부' ? 'ext' : 'self') }, c.repairBy) : h('span', { class: 'kind meta' }, '구분 없음'); }
  function safetyStrip() {
    return h('div', { class: 'safety-strip', role: 'note' }, h('strong', null, '안전 먼저 '), L.SAFETY.common[0]);
  }

  /* ── 사진 (data09-17 방식: 줄인 사본을 IndexedDB 에, 사례에는 { id, name, w, h, bytes, memo } 만) ── */
  var thumbUrls = {}, pendingPhotos = {}, photoStoreWarned = false;
  function thumbUrl(id) {
    if (thumbUrls[id]) return Promise.resolve(thumbUrls[id]);
    return P.get(id).then(function (rec) {
      if (!rec) return '';
      thumbUrls[id] = URL.createObjectURL(rec.thumb || rec.blob);
      return thumbUrls[id];
    }, function () { return ''; });
  }
  function forgetThumbs(ids) { ids.forEach(function (id) { if (thumbUrls[id]) { URL.revokeObjectURL(thumbUrls[id]); delete thumbUrls[id]; } }); }
  function dropPhotos(ids) { if (!ids.length) return Promise.resolve(); forgetThumbs(ids); return P.del(ids).catch(function () {}); }
  function gcPhotos() {
    return P.keys().then(function (ks) { return dropPhotos(L.orphanPhotoIds(db.cases, ks || [], Object.keys(pendingPhotos))); }).catch(function () {});
  }
  function clearPhotos() { forgetThumbs(Object.keys(thumbUrls)); return P.clear().catch(function () {}); }
  function thumbImg(id, alt) {
    var img = h('img', { class: 'thumb', alt: alt || '고장 사진', loading: 'lazy', width: '56', height: '56' });
    thumbUrl(id).then(function (u) { if (u) img.src = u; else img.classList.add('thumb-missing'); });
    return img;
  }
  function caseTitle(c) { return dateText(c) + ' · ' + (c.equipment || '설비 미기재') + ' · ' + L.canonCause(c.cause); }
  function photoCell(c) {
    var ps = c.photos || [];
    if (!ps.length) return h('span', { class: 'note' }, '-');
    return h('button', { type: 'button', class: 'thumb-btn', 'aria-label': '사진 ' + ps.length + '장 크게 보기', onclick: function (e) { e.stopPropagation(); openLightbox(ps, 0, caseTitle(c)); } },
      thumbImg(ps[0].id, caseTitle(c) + ' 사진 1'), ps.length > 1 ? h('span', { class: 'thumb-count' }, '+' + (ps.length - 1)) : null);
  }
  var lb = { list: [], i: 0, url: '', title: '' };
  function openLightbox(list, i, title) {
    lb.list = list || []; lb.title = title || ''; lb.i = Math.max(0, Math.min(i || 0, lb.list.length - 1));
    if (!lb.list.length) return;
    showLightbox();
    var d = document.getElementById('lightbox'); if (!d.open) d.showModal();
  }
  function showLightbox() {
    var p = lb.list[lb.i], img = document.getElementById('lbImg');
    document.getElementById('lbTitle').textContent = lb.title || '고장 사진';
    document.getElementById('lbCap').textContent = (lb.i + 1) + ' / ' + lb.list.length + (p.memo ? ' · ' + p.memo : '') + (p.name ? ' · ' + p.name : '');
    document.getElementById('lbPrev').disabled = lb.i === 0;
    document.getElementById('lbNext').disabled = lb.i >= lb.list.length - 1;
    img.removeAttribute('src'); img.alt = '고장 사진 ' + (lb.i + 1);
    P.get(p.id).then(function (rec) {
      if (lb.url) URL.revokeObjectURL(lb.url);
      lb.url = rec ? URL.createObjectURL(rec.blob) : '';
      if (lb.url) img.src = lb.url; else img.alt = '사진을 찾지 못했습니다';
    });
  }
  (function () {
    var d = document.getElementById('lightbox');
    document.getElementById('lbPrev').addEventListener('click', function () { if (lb.i > 0) { lb.i--; showLightbox(); } });
    document.getElementById('lbNext').addEventListener('click', function () { if (lb.i < lb.list.length - 1) { lb.i++; showLightbox(); } });
    document.getElementById('lbClose').addEventListener('click', function () { d.close(); });
    d.addEventListener('click', function (e) { if (e.target === d) d.close(); });
    d.addEventListener('close', function () { if (lb.url) { URL.revokeObjectURL(lb.url); lb.url = ''; } });
  })();
  /* 저장 공간 — 사진 합계와 브라우저가 알려 주는 사용량으로 경고 */
  function storageNote(box) {
    var bytes = L.photoBytes(db.cases), n = L.allPhotoIds(db.cases).length;
    function draw(q) {
      var lv = L.storageLevel(bytes, q);
      box.textContent = '';
      var msg = '사진 ' + n + '장 · 약 ' + L.round(bytes / 1048576, 1) + 'MB' + (q && q.quota ? ' (이 사이트가 쓰는 공간 ' + L.round(q.usage / 1048576, 1) + 'MB / 브라우저 허용 약 ' + Math.round(q.quota / 1048576) + 'MB)' : '') +
        ' · 글 ' + S.sizeKb() + 'KB(localStorage, 보통 5MB 안팎까지).';
      add(box, h('p', { class: lv === 'ok' ? 'note' : 'alert ' + (lv === 'bad' ? 'warn' : 'info') }, msg,
        lv === 'bad' ? ' 저장 공간이 거의 찼습니다. 「설정 → 백업 내려받기」 로 받아 두고, 오래된 사진은 빼 주세요.' :
          lv === 'warn' ? ' 사진이 많이 쌓였습니다. 브라우저 데이터를 지우면 함께 사라지니 백업을 받아 두세요.' : ''));
    }
    draw(null);
    if (navigator.storage && navigator.storage.estimate) navigator.storage.estimate().then(draw, function () {});
    return box;
  }

  /* ── 대화상자 ───────────────────────────── */
  var dlg = document.getElementById('dlg'), onDialogClose = null;
  function openDialog(title, content, actions, wide) {
    document.getElementById('dlgTitle').textContent = title;
    var c = document.getElementById('dlgContent'); c.textContent = ''; add(c, content);
    var a = document.getElementById('dlgActions'); a.textContent = ''; add(a, actions);
    dlg.classList.toggle('wide', !!wide);
    if (!dlg.open) dlg.showModal();
  }
  function closeDialog() { if (dlg.open) dlg.close(); }
  dlg.addEventListener('close', function () { var f = onDialogClose; onDialogClose = null; if (f) f(); });

  /* ── 메뉴·상단 ───────────────────────────── */
  var ROUTES = [
    { id: 'diagnose', no: '①', t: '증상 진단', s: '증상 → 가능 원인 · 먼저 확인할 것' },
    { id: 'cases', no: '②', t: '정비 사례', s: '등록 · 사진 · 엑셀 내보내기' },
    { id: 'guide', no: '③', t: '자체 정비 가이드북', s: '원인별 카드 · 인쇄' },
    { id: 'cost', no: '④', t: '비용 현황', s: '외부 수리 비용 · 자체 처리 비율' },
    { id: 'import', no: '', t: '내역서 가져오기', s: '수리이력·내역서 엑셀 → 사례' },
    { id: 'check', no: '', t: '자료 점검', s: '칸별 채움 · 더 채울 칸 · 권장 양식' },
    { id: 'settings', no: '', t: '설정', s: '설비 · AI 연결 · 백업' },
    { id: 'next', no: '', t: '다음 단계', s: '보내 주실 자료' }
  ];
  function route() {
    var m = /^#\/([a-z]+)(?:\?(.*))?$/.exec(location.hash || '');
    var id = m ? m[1] : 'diagnose';
    if (!ROUTES.some(function (r) { return r.id === id; })) id = 'diagnose';
    var q = {};
    if (m && m[2]) m[2].split('&').forEach(function (kv) { var p = kv.split('='); q[decodeURIComponent(p[0])] = decodeURIComponent(p[1] || ''); });
    return { id: id, q: q };
  }
  function nav() {
    var cur = route().id, n = document.getElementById('nav'); n.textContent = '';
    ROUTES.forEach(function (r) {
      n.appendChild(h('a', { href: '#/' + r.id, 'aria-current': r.id === cur ? 'page' : null }, h('span', { class: 'no' }, r.no || '·'), h('span', { class: 't' }, r.t), h('span', { class: 's' }, r.s)));
    });
  }
  function topbar() {
    var chip = document.getElementById('dataChip'); chip.textContent = '';
    add(chip, ['정비 사례 ' + db.cases.length + '건', h('span', { class: 'dim' }, '설비 ' + db.equipments.length + '대 · 원인 ' + L.causeList(db.cases).length + '가지')]);
    document.getElementById('storeNote').textContent = S.available() ? '이 브라우저에 저장됨 · ' + S.sizeKb() + 'KB' : '저장 안 됨 — 백업을 받아 두세요';
    var b = document.getElementById('sampleBanner');
    b.hidden = !db._sample; b.textContent = '';
    if (db._sample) add(b, ['예시 데이터(가상)를 쓰는 중입니다 — 설비 A~E 와 정비 사례 30건은 모두 지어낸 값입니다. ',
      h('button', { type: 'button', class: 'linklike', onclick: clearAll }, '예시 지우고 내 자료로 시작')]);
  }
  function emptyNotice(what) {
    return h('section', { class: 'card empty' },
      h('p', null, what || '아직 정비 사례가 없습니다.'),
      h('p', null, '정비 내역서 엑셀을 가져오거나, 사례를 직접 등록하거나, 가상 예시 30건으로 먼저 둘러봐 주세요.'),
      h('div', { class: 'btn-row', style: 'justify-content:center' },
        h('button', { type: 'button', class: 'btn btn-primary', onclick: loadSample }, '예시 데이터로 시작'),
        h('a', { class: 'btn', href: '#/import' }, '내역서 가져오기'),
        h('button', { type: 'button', class: 'btn', onclick: function () { editCase(null); } }, '사례 직접 등록')));
  }
  function loadSample() {
    function go() {
      clearPhotos().then(function () {
        db = SMP.db(); save(); closeDialog(); toast('가상 예시(설비 5대 · 정비 사례 30건)를 넣었습니다');
        ui.diag = { equipment: '', text: '', tags: [], tagTouched: {}, result: null, manual: '', manualFor: null, answer: '' };
        render();
      });
    }
    if (!db.cases.length) { go(); return; }
    openDialog('예시 데이터 불러오기', h('p', null, '지금 있는 정비 사례 ' + db.cases.length + '건' + (L.allPhotoIds(db.cases).length ? '·사진 ' + L.allPhotoIds(db.cases).length + '장' : '') + '과 설비 목록·가이드 수정 내용을 지우고 예시로 바꿉니다. 먼저 「설정 → 백업 내려받기」 로 받아 두세요.'), [
      h('button', { type: 'button', class: 'btn', onclick: closeDialog }, '취소'),
      h('button', { type: 'button', class: 'btn btn-danger', onclick: go }, '예시로 바꾸기')]);
  }
  function clearAll() {
    openDialog('모두 지우기', h('p', null, '이 브라우저에 저장된 정비 사례 ' + db.cases.length + '건' + (L.allPhotoIds(db.cases).length ? '·사진 ' + L.allPhotoIds(db.cases).length + '장' : '') + '·설비 목록·가이드 수정 내용을 지웁니다. AI 연결 설정은 남깁니다. 되돌릴 수 없습니다.'), [
      h('button', { type: 'button', class: 'btn', onclick: closeDialog }, '취소'),
      h('button', { type: 'button', class: 'btn btn-danger', onclick: function () {
        var keepMap = db.mapping, keepOpts = db.opts;
        db = L.emptyDb(); db.mapping = keepMap; db.opts = keepOpts;
        save(); closeDialog(); clearPhotos(); ui.diag.result = null; toast('모두 지웠습니다'); render();
      } }, '지우기')]);
  }

  /* ── ① 증상 진단 ─────────────────────────── */
  function viewDiagnose() {
    var d = ui.diag;
    var wrap = h('div');
    add(wrap, pageHead('① 증상 진단', '증상을 넣으면 원인을 먼저 짚어 봅니다',
      '과거 정비 사례에서 비슷한 증상을 찾아 「그때 무엇이 원인이었고, 무엇부터 확인했는지」 를 보여 줍니다. 확정 진단이 아니라 과거 사례 기반 추정입니다.'));
    if (!db.cases.length) { add(wrap, emptyNotice('아직 정비 사례가 없어 비교할 과거 사례가 없습니다.')); return wrap; }

    var eqSel = sel([['', '설비 지정 안 함(모든 설비에서 찾기)']].concat(db.equipments.map(function (e) { return [e.name, e.name + (e.type ? ' · ' + e.type : '')]; })), d.equipment, { 'aria-label': '설비' });
    var ta = h('textarea', { rows: '3', placeholder: '예: 온도가 안 올라가고 기계가 멈춤', 'aria-label': '증상' }); ta.value = d.text;
    var chips = h('div', { class: 'chips', role: 'group', 'aria-label': '증상 태그' });
    function autoTags() {
      var auto = L.extractTags(ta.value);
      var out = [];
      L.tagList().forEach(function (t) {
        var touched = d.tagTouched[t];
        if (touched === true || (touched == null && auto.indexOf(t) >= 0)) out.push(t);
      });
      d.tags = out;
    }
    function drawChips() {
      chips.textContent = '';
      L.tagList().forEach(function (t) {
        var on = d.tags.indexOf(t) >= 0;
        chips.appendChild(h('button', { type: 'button', class: 'chip', 'aria-pressed': on ? 'true' : 'false', onclick: function () {
          d.tagTouched[t] = !on; autoTags(); drawChips();
        } }, t));
      });
    }
    ta.addEventListener('input', function () { d.text = ta.value; autoTags(); drawChips(); });
    eqSel.addEventListener('change', function () { d.equipment = eqSel.value; });
    autoTags(); drawChips();
    function run() {
      d.text = ta.value; d.equipment = eqSel.value;
      if (!L.str(d.text) && !d.tags.length) { toast('증상을 적거나 태그를 골라 주세요', true); ta.focus(); return; }
      var eq = equipOf(d.equipment);
      d.result = L.diagnose(db.cases, { text: d.text, tags: d.tags, equipment: d.equipment, equipmentType: eq ? eq.type : '' }, db.opts, equipTypes());
      d.answer = '';
      if (d.manualFor !== d.equipment) { d.manual = eq ? eq.manual : ''; d.manualFor = d.equipment; }
      render();
      setTimeout(function () { var r = document.getElementById('diagResult'); if (r) r.scrollIntoView({ behavior: 'smooth', block: 'start' }); }, 30);
    }
    var examples = ['온도가 안 올라가고 기계가 멈춤', '기계 아래 바닥에 기름이 고임', '차단기가 떨어지고 모터에서 탄 냄새', '제품이 덜 채워지고 노즐에서 수지가 안 나옴'];
    // 내 자료일 때는 자주 난 원인 묶음의 최근 증상으로 예시 버튼을 만듭니다
    if (!db._sample && db.cases.length) examples = L.causeList(db.cases).filter(function (x) { return !L.isUnknownCause(x.cause); }).slice(0, 4).map(function (x) {
      var c = db.cases.filter(function (y) { return L.causeKey(y.cause) === x.key && y.symptom; }).sort(function (a, b) { return String(b.date).localeCompare(String(a.date)); })[0];
      return c ? c.symptom.slice(0, 30) : '';
    }).filter(Boolean);
    add(wrap, h('section', { class: 'card' },
      h('h2', null, '지금 증상'),
      h('div', { class: 'form-grid diag-form' },
        fld('설비', eqSel, '고르면 같은 설비 사례에 조금 더 무게를 둡니다'),
        fld('증상 (보이는 그대로 한두 문장)', ta, '소리·냄새·알람 글자·몇 번째 존인지도 적으면 더 잘 찾습니다', 'span-2')),
      h('div', { class: 'field' }, h('span', null, '증상 태그'), chips, h('small', { class: 'hint' }, '글에서 자동으로 켜집니다. 눌러서 켜고 끌 수 있습니다.')),
      h('div', { class: 'btn-row', style: 'margin-top:12px' },
        h('button', { type: 'button', class: 'btn btn-primary btn-big', onclick: run }, '원인 짚어 보기'),
        h('span', { class: 'note ex-label' }, '예:'),
        examples.map(function (x) { return h('button', { type: 'button', class: 'btn btn-sm', onclick: function () { ta.value = x; d.text = x; d.tagTouched = {}; autoTags(); drawChips(); run(); } }, x); }))));

    if (d.result) add(wrap, diagResult(d));
    return wrap;
  }
  function diagResult(d) {
    var r = d.result, box = h('div', { id: 'diagResult' });
    add(box, safetyStrip());
    add(box, h('p', { class: 'alert info' }, '아래는 과거 정비 사례 ' + db.cases.length + '건 가운데 비슷한 ' + r.sims.length + '건으로 계산한 추정입니다. 확정 진단이 아니며, 사례가 쌓일수록 정확해집니다.'));
    if (!r.causes.length) {
      add(box, h('section', { class: 'card' }, h('h2', null, '비슷한 과거 사례가 없습니다'),
        h('p', null, '증상을 다른 말로 적어 보거나(예: 「기름」 「누유」 「오일」), 태그를 골라 보세요. 처음 보는 고장이면 정비 담당자·외부 업체를 부르고, 고친 뒤 「정비 사례」 에 등록해 두면 다음부터 찾을 수 있습니다.'),
        h('button', { type: 'button', class: 'btn', onclick: function () { editCase(null, { equipment: d.equipment, symptom: d.text, tags: d.tags, date: today() }); } }, '이 증상을 새 사례로 등록')));
      add(box, aiCard(d));
      return box;
    }
    var list = h('div', { class: 'an-list' });
    r.causes.forEach(function (c, i) {
      list.appendChild(h('article', { class: 'an cause lv-' + c.level },
        h('div', { class: 'an-top' },
          h('span', { class: 'rank' }, (i + 1) + '위'),
          h('h3', { class: 'cause-name' }, c.cause),
          h('span', { class: 'kind lv ' + c.level }, L.LEVEL_TEXT[c.level])),
        h('div', { class: 'share' }, h('div', { class: 'share-bar', role: 'img', 'aria-label': '비중 ' + L.fmtPct(c.share) }, h('i', { style: 'width:' + Math.round(c.share * 100) + '%' })), h('b', null, L.fmtPct(c.share))),
        h('p', { class: 'note' }, '비슷한 사례 ' + c.count + '건이 이 원인이었습니다 · 이 원인 전체 ' + c.totalCases + '건(외부 ' + c.extCount + ' · 자체 ' + c.selfCount +
          (c.totalCases - c.extCount - c.selfCount ? ' · 구분 모름 ' + (c.totalCases - c.extCount - c.selfCount) : '') + ')' + (L.partOf(c.cause) ? ' · 부위: ' + L.partOf(c.cause) : '') +
          (c.avgExtCost != null ? ' · 외부 수리 평균 ' + L.fmtWon(c.avgExtCost) : '')),
        c.checks.length ? h('div', { class: 'checks' }, h('strong', null, '먼저 확인할 것'),
          h('ol', null, c.checks.map(function (x) {
            return h('li', null, x.text, h('span', { class: 'note' }, ' (그때 사례 ' + x.count + '건에서 먼저 본 것)'),
              x.kind === 'expert' ? h('span', { class: 'kind danger' }, '담당자·전문가') : null);
          }))) : null,
        h('div', { class: 'btn-row' }, h('a', { class: 'btn btn-sm', href: '#/guide?k=' + encodeURIComponent(c.key) }, '이 원인의 정비 가이드 보기'))));
    });
    add(box, h('section', { class: 'card' }, h('h2', null, '가능 원인 (과거 사례 기반 추정)'), list));

    var rows = r.sims.map(function (s, i) {
      var c = s.c;
      return { cells: [
        String(i + 1), Math.min(100, Math.round(s.score * 100)) + '%', dateText(c), c.equipment || '-',
        h('span', null, c.symptom, s.sharedTags.length ? h('span', { class: 'note' }, ' · 같은 태그: ' + s.sharedTags.join(', ')) : null),
        causeCell(c), L.stepCore(L.firstCheck(c)) || L.splitSteps(c.steps).join(' / ') || '-', h('span', null, byBadge(c), ' ', costText(c), statusBadge(c)), photoCell(c),
        h('button', { type: 'button', class: 'btn btn-sm', onclick: function () { showCase(c); } }, '자세히')
      ] };
    });
    add(box, h('section', { class: 'card' }, h('h2', null, '비슷한 과거 정비 사례'),
      h('p', { class: 'note' }, '유사도 = 증상 글의 겹치는 글자 묶음(TF-IDF, 「멈춤·작동 불가」 같은 같은 뜻 말 포함) ' + (r.tags.length ? '+ 같은 증상 태그 ' : '') + '+ 같은 설비·같은 종류 가산점 + 증상에 적은 부위(서보·벨트·냉각 …)가 원인과 같으면 가산점. 「짐작」 은 원인 칸이 없어 증상·조치 글에서 짐작한 원인입니다.'),
      table(['#', '유사도', '일자', '설비', '증상', '원인', '먼저 본 것', '수리', '사진', ''], rows, { num: [true, true], cls: 'sims' })));
    add(box, aiCard(d));
    add(box, h('div', { class: 'btn-row' },
      h('button', { type: 'button', class: 'btn', onclick: function () { editCase(null, { equipment: d.equipment, symptom: d.text, tags: d.tags, date: today() }); } }, '고친 뒤 이 증상을 새 정비 사례로 등록')));
    return box;
  }
  function aiCard(d) {
    var eq = equipOf(d.equipment);
    var man = h('textarea', { rows: '4', 'aria-label': '매뉴얼 발췌', placeholder: '설비 매뉴얼의 고장 대응 표·알람 설명을 붙여 넣으면 AI 가 함께 봅니다(선택)' }); man.value = d.manual || '';
    var pr = h('textarea', { rows: '10', readonly: true, class: 'mono', 'aria-label': 'AI 프롬프트' });
    function prompt() { return L.buildDiagnosePrompt({ text: d.text, tags: d.tags, equipment: d.equipment, equipmentType: eq ? eq.type : '', manual: man.value }, d.result); }
    pr.value = prompt();
    man.addEventListener('input', function () { d.manual = man.value; pr.value = prompt(); });
    var ans = h('textarea', { rows: '8', 'aria-label': 'AI 답', placeholder: 'AI 의 답을 여기에 붙여 넣으면 이 화면에 남겨 둡니다(이 창을 닫으면 사라집니다)' }); ans.value = d.answer || '';
    ans.addEventListener('input', function () { d.answer = ans.value; });
    return h('section', { class: 'card' },
      h('h2', null, 'AI 에게 더 자세히 묻기 (반자동)'),
      h('p', { class: 'note' }, '증상 + 위의 비슷한 사례 ' + d.result.sims.length + '건 + 매뉴얼 발췌를 묶은 프롬프트입니다. 복사해서 회사가 허용한 AI 에 붙여 넣고, 답을 아래에 붙여 넣어 주세요. 사례에 회사·업체 이름이 있다면 지우고 보내세요.'),
      fld('매뉴얼 발췌' + (eq && eq.manual ? ' (설정의 「' + eq.name + '」 매뉴얼 메모에서 가져옴)' : ''), man),
      fld('프롬프트', pr),
      h('div', { class: 'btn-row', style: 'margin:8px 0 14px' },
        h('button', { type: 'button', class: 'btn btn-primary', onclick: function () { copyText(pr.value, '프롬프트'); } }, '프롬프트 복사'),
        AIP.sendButton('설정한 AI 서버로 보내기', function () { return pr.value; }, function (t) { ans.value = t; d.answer = t; toast('AI 답을 받았습니다'); }, { toast: toast, settingsHref: '#/settings' })),
      fld('AI 답', ans, '답도 추정입니다. 전기·유압·고온 작업이 나오면 초보자는 멈추고 담당자를 부릅니다.'));
  }

  /* ── 사례 자세히 ───────────────────────────── */
  function showCase(c) {
    var st = L.splitSteps(c.steps);
    var dl = h('dl', { class: 'summary' },
      h('dt', null, '일자'), h('dd', null, dateText(c), c.datePrec && c.datePrec !== 'day' ? h('span', { class: 'note' }, ' (' + L.PREC_TEXT[c.datePrec] + '까지 적힘)') : null),
      h('dt', null, '설비'), h('dd', null, c.equipment || '-', L.equipFamily(c.equipment) && L.equipFamily(c.equipment) !== c.equipment ? h('span', { class: 'note' }, ' (종류: ' + L.equipFamily(c.equipment) + ')') : null),
      h('dt', null, '증상'), h('dd', null, c.symptom || '-'),
      h('dt', null, '증상 태그'), h('dd', null, (c.tags || []).join(', ') || '-'),
      h('dt', null, '원인'), h('dd', null, c.cause || '-', c.causeInferred ? h('span', { class: 'kind meta' }, '증상·조치에서 짐작') : null,
        !c.causeInferred && L.canonCause(c.cause) !== c.cause ? h('span', { class: 'note' }, ' (묶음: ' + L.canonCause(c.cause) + ')') : null, L.partOf(c.cause) ? h('span', { class: 'note' }, ' · 부위: ' + L.partOf(c.cause)) : null),
      h('dt', null, '조치 단계'), h('dd', null, st.length ? h('ol', { class: 'steps' }, st.map(function (s) { var k = L.stepKind(s); return h('li', null, s, k === 'safety' ? h('span', { class: 'kind ok' }, '안전') : k === 'expert' ? h('span', { class: 'kind danger' }, '담당자·전문가') : k === 'note' ? h('span', { class: 'kind meta' }, '진행 기록') : null); })) : '-'),
      h('dt', null, '교체 부품'), h('dd', null, L.splitList(c.parts).join(', ') || '-'),
      h('dt', null, '필요 공구'), h('dd', null, L.splitList(c.tools).join(', ') || '-'),
      h('dt', null, '수리'), h('dd', null, byBadge(c), c.vendor && c.repairBy !== '자체' ? ' ' + c.vendor : '', ' · ', costText(c), c.hours != null ? ' · ' + c.hours + '시간' : ''),
      h('dt', null, '처리 상태'), h('dd', null, statusBadge(c) || '-'),
      h('dt', null, '사진·영상 위치'), h('dd', null, c.media || '-'),
      h('dt', null, '비고'), h('dd', null, c.memo || '-'),
      h('dt', null, '출처'), h('dd', null, [c.ref, c.source].filter(Boolean).join(' · ') || '-'));
    var pics = (c.photos || []).length ? h('ul', { class: 'photo-grid' }, c.photos.map(function (p, i) {
      return h('li', { class: 'photo-item' }, h('button', { type: 'button', class: 'thumb-btn big', 'aria-label': '사진 ' + (i + 1) + ' 크게 보기', onclick: function () { openLightbox(c.photos, i, caseTitle(c)); } }, thumbImg(p.id, '사진 ' + (i + 1))), p.memo ? h('small', null, p.memo) : null);
    })) : null;
    openDialog(caseTitle(c), [dl, pics], [
      h('button', { type: 'button', class: 'btn', onclick: function () { closeDialog(); editCase(c); } }, '고치기'),
      h('button', { type: 'button', class: 'btn btn-primary', onclick: closeDialog }, '닫기')], true);
  }

  /* ── 사례 추가·고치기 ───────────────────────── */
  function editCase(c, preset) {
    var isNew = !c;
    var src = c || L.cleanCase(Object.assign({ date: today(), source: '직접 입력' }, preset || {}));
    var inputs = {
      date: h('input', { type: 'date', value: src.date || '' }),
      equipment: h('input', { value: src.equipment || '', list: 'eqList', autocomplete: 'off', placeholder: '예: 설비 A' }),
      symptom: h('textarea', { rows: '3', placeholder: '예: 온도가 안 올라가고 기계가 멈춤' }),
      cause: h('input', { value: src.cause || '', list: 'causeList', autocomplete: 'off', placeholder: '예: 히터 단선 (2존 히터 밴드)' }),
      steps: h('textarea', { rows: '6', placeholder: '한 줄에 한 단계\n설비 정지 후 전원 차단·잠금(LOTO)\n센서 커넥터 연결 상태 확인\n…' }),
      parts: h('input', { value: L.splitList(src.parts).join(', '), placeholder: '쉼표로 나눔. 예: 히터 밴드, 오링' }),
      tools: h('input', { value: L.splitList(src.tools).join(', '), placeholder: '예: 테스터, 육각렌치' }),
      repairBy: sel([['', '고르기'], ['외부', '외부(업체 수리)'], ['자체', '자체(팀에서 수리)']], src.repairBy),
      vendor: h('input', { value: src.vendor || '', placeholder: '예: 업체 이름 · 전기기사 · 미기재' }),
      cost: h('input', { type: 'number', min: '0', step: '1000', value: src.cost == null ? '' : src.cost, inputmode: 'numeric' }),
      costKind: sel([['won', '실제 지출'], ['estimate', '견적(실지출 아님)'], ['unit', '개당 단가'], ['free', '무상'], ['missing', '미기재']], src.costKind || (src.cost == null ? 'missing' : 'won')),
      status: sel([['', '모름'], ['완료', '완료'], ['진행 중', '진행 중']], src.status || ''),
      hours: h('input', { type: 'number', min: '0', step: '0.5', value: src.hours == null ? '' : src.hours, inputmode: 'decimal' }),
      media: h('input', { value: src.media || '', placeholder: '예: 공유폴더/정비사진/2026-09 설비A 소음.mp4' }),
      memo: h('textarea', { rows: '2' })
    };
    inputs.symptom.value = src.symptom || ''; inputs.steps.value = L.splitSteps(src.steps).join('\n'); inputs.memo.value = src.memo || '';
    var tags = (src.tags || []).slice();
    var chips = h('div', { class: 'chips', role: 'group', 'aria-label': '증상 태그' });
    function drawChips() {
      chips.textContent = '';
      L.tagList().concat(tags.filter(function (t) { return L.tagList().indexOf(t) < 0; })).forEach(function (t) {
        var on = tags.indexOf(t) >= 0;
        chips.appendChild(h('button', { type: 'button', class: 'chip', 'aria-pressed': on ? 'true' : 'false', onclick: function () {
          if (on) tags.splice(tags.indexOf(t), 1); else tags.push(t); drawChips();
        } }, t));
      });
    }
    drawChips();
    var autoBtn = h('button', { type: 'button', class: 'btn btn-sm', onclick: function () { tags = L.uniq(tags.concat(L.extractTags(inputs.symptom.value))); drawChips(); } }, '증상 글에서 태그 뽑기');
    var dl1 = h('datalist', { id: 'eqList' }, db.equipments.map(function (e) { return h('option', { value: e.name }); }));
    var dl2 = h('datalist', { id: 'causeList' }, L.causeList(db.cases).map(function (x) { return h('option', { value: x.cause }); }));

    // 사진 — 창 안에서는 work 만 바꾸고 「저장」 때 사례에 붙입니다. 저장하지 않고 닫으면 이번에 넣은 사진을 지웁니다.
    var orig = (src.photos || []).map(function (p) { return Object.assign({}, p); });
    var work = orig.map(function (p) { return Object.assign({}, p); });
    var added = [], saved = false;
    var photoBox = h('div', { class: 'photo-edit' }), photoMsg = h('p', { class: 'note', role: 'status', 'aria-live': 'polite' }), storeBox = h('div');
    function addPhotoFiles(fileList) {
      var files = Array.prototype.slice.call(fileList || []).filter(function (f) { return /^image\//.test(f.type) || /\.(jpe?g|png|webp|gif|bmp|heic|heif)$/i.test(f.name); });
      var skippedVideo = Array.prototype.slice.call(fileList || []).filter(function (f) { return /^video\//.test(f.type); }).length;
      if (skippedVideo) toast('동영상은 저장하지 않습니다. 파일 이름·보관 위치를 「사진·영상 위치」 칸에 적어 주세요', true);
      if (!files.length) return;
      var room = L.PHOTO_LIMIT - work.length;
      if (room <= 0) { toast('한 건에 사진은 ' + L.PHOTO_LIMIT + '장까지 붙일 수 있습니다', true); return; }
      if (files.length > room) toast('한 건에 ' + L.PHOTO_LIMIT + '장까지라 ' + room + '장만 넣습니다', true);
      files = files.slice(0, room);
      photoMsg.textContent = '사진 ' + files.length + '장을 줄여서 저장하는 중입니다…';
      var fails = [];
      files.reduce(function (chain, f) {
        return chain.then(function () {
          return P.shrink(f).then(function (x) {
            var id = P.newId();
            pendingPhotos[id] = true; added.push(id);
            return P.put({ id: id, caseId: src.id || '', name: f.name || '', w: x.w, h: x.h, bytes: x.bytes, memo: '', at: Date.now(), blob: x.blob, thumb: x.thumb })
              .then(function () { work.push({ id: id, name: f.name || '', w: x.w, h: x.h, bytes: x.bytes, memo: '' }); drawPhotos(); });
          }).catch(function (e) { fails.push((f.name || '사진') + ': ' + (e && e.message ? e.message : '저장하지 못했습니다')); });
        });
      }, Promise.resolve()).then(function () {
        photoMsg.textContent = fails.length ? '넣지 못한 사진 — ' + fails.join(' / ') : '';
        if (!P.available() && !photoStoreWarned) { photoStoreWarned = true; toast('이 브라우저에서는 사진 저장소(IndexedDB)를 쓸 수 없어, 창을 닫으면 사진이 사라집니다', true); }
        if (navigator.storage && navigator.storage.persist) navigator.storage.persist().catch(function () {});
        storageNote(storeBox);
      });
    }
    function rotatePhoto(i) {
      var p = work[i];
      photoMsg.textContent = '사진을 돌리는 중입니다…';
      P.get(p.id).then(function (rec) {
        if (!rec) throw new Error('사진을 찾지 못했습니다');
        return P.shrink(rec.blob, 90).then(function (x) {
          var id = P.newId();
          pendingPhotos[id] = true; added.push(id);
          return P.put({ id: id, caseId: rec.caseId, name: p.name, w: x.w, h: x.h, bytes: x.bytes, memo: p.memo, at: Date.now(), blob: x.blob, thumb: x.thumb }).then(function () {
            work[i] = Object.assign({}, p, { id: id, w: x.w, h: x.h, bytes: x.bytes }); photoMsg.textContent = ''; drawPhotos();
          });
        });
      }).catch(function (e) { photoMsg.textContent = e.message || '돌리지 못했습니다'; });
    }
    function drawPhotos() {
      photoBox.textContent = '';
      var grid = h('ul', { class: 'photo-grid' });
      work.forEach(function (p, i) {
        var memo = h('input', { type: 'text', value: p.memo || '', placeholder: '설명(선택)', 'aria-label': '사진 ' + (i + 1) + ' 설명', maxlength: '100' });
        memo.addEventListener('input', function () { p.memo = memo.value; });
        grid.appendChild(h('li', { class: 'photo-item' },
          h('button', { type: 'button', class: 'thumb-btn big', 'aria-label': '사진 ' + (i + 1) + ' 크게 보기', onclick: function () { openLightbox(work, i, isNew ? '새 사례' : caseTitle(src)); } }, thumbImg(p.id, '사진 ' + (i + 1))),
          memo,
          h('div', { class: 'photo-tools' },
            h('button', { type: 'button', class: 'btn btn-sm', disabled: i === 0, onclick: function () { var t = work[i - 1]; work[i - 1] = work[i]; work[i] = t; drawPhotos(); } }, '앞으로'),
            h('button', { type: 'button', class: 'btn btn-sm', onclick: function () { rotatePhoto(i); } }, '돌리기'),
            h('button', { type: 'button', class: 'btn btn-sm btn-danger', onclick: function () { work.splice(i, 1); drawPhotos(); } }, '빼기'))));
      });
      var pick = h('input', { type: 'file', accept: 'image/*', multiple: true, 'aria-label': '사진 파일 고르기', onchange: function (e) { addPhotoFiles(e.target.files); e.target.value = ''; } });
      var cam = h('input', { type: 'file', accept: 'image/*', capture: 'environment', 'aria-label': '카메라로 찍기', onchange: function (e) { addPhotoFiles(e.target.files); e.target.value = ''; } });
      add(photoBox, [
        work.length ? grid : h('p', { class: 'note' }, '붙인 사진이 없습니다.'),
        h('div', { class: 'btn-row' }, h('label', { class: 'btn file-btn' }, '사진 고르기', pick), h('label', { class: 'btn file-btn' }, '카메라로 찍기', cam)),
        photoMsg,
        h('p', { class: 'note' }, '한 건에 ' + L.PHOTO_LIMIT + '장까지. 긴 변 ' + L.PHOTO_MAX + 'px 로 줄인 사본만 이 브라우저(IndexedDB)에 저장합니다. 원본·동영상은 회사 공유폴더에 두고 위치만 적어 주세요.'),
        storeBox
      ]);
    }
    drawPhotos(); storageNote(storeBox);
    onDialogClose = function () {
      added.forEach(function (id) { delete pendingPhotos[id]; });
      if (!saved) dropPhotos(added);
    };
    var errBox = h('div');
    var form = h('div', { class: 'form-grid' },
      fld('일자', inputs.date), fld('설비', inputs.equipment, '새 이름이면 설비 목록에 이름만 더해집니다'), fld('외부/자체', inputs.repairBy),
      fld('증상 *', inputs.symptom, '현장 작업자가 본 그대로', 'span-all'),
      h('div', { class: 'field span-all' }, h('span', null, '증상 태그'), chips, autoBtn),
      fld('원인 *', inputs.cause, src.causeInferred ? '증상·조치 글에서 짐작한 원인입니다. 업체가 말한 원인으로 고치면 「짐작」 표시가 사라집니다' : '같은 원인을 다르게 적어도 「히터 단선」 처럼 묶어서 셉니다', 'span-all'),
      fld('조치 단계', inputs.steps, '한 줄에 한 단계. 첫 줄은 전원 차단·잠금(LOTO). 「— 결과」 를 붙이면 그때 본 결과로 남깁니다(예: 센서 커넥터 확인 — 반쯤 빠짐)', 'span-all'),
      fld('교체 부품', inputs.parts), fld('필요 공구', inputs.tools), fld('수리 업체', inputs.vendor),
      fld('비용(원)', inputs.cost, '외부는 청구 금액, 자체는 부품 값'), fld('비용 구분', inputs.costKind, '견적은 비용 합계에서 따로 셉니다'), fld('처리 상태', inputs.status),
      fld('소요시간(시간)', inputs.hours), fld('사진·영상 위치', inputs.media, '동영상은 저장하지 않고 위치만', 'span-2'),
      fld('비고', inputs.memo, null, 'span-all'),
      h('div', { class: 'field span-all' }, h('span', null, '사진'), photoBox));
    function collect() {
      var o = {};
      Object.keys(inputs).forEach(function (k) { o[k] = inputs[k].value; });
      o.tags = tags; o.id = src.id; o.source = src.source; o.ref = src.ref; o.createdAt = src.createdAt || new Date().toISOString();
      // 주차·월까지만 적힌 일자는 고치지 않았으면 그 정밀도를 지킵니다. 원인을 고치면 「짐작」 표시를 뗍니다
      o.datePrec = o.date && o.date === src.date ? src.datePrec || 'day' : 'day';
      o.causeInferred = !!src.causeInferred && L.str(o.cause) === L.str(src.cause);
      if (L.str(o.cost) !== '' && (o.costKind === 'missing' || o.costKind === 'free' && Number(o.cost) > 0)) o.costKind = 'won';
      if (L.str(o.cost) === '' && o.costKind !== 'free') o.costKind = 'missing';
      if (o.costKind === 'free') o.cost = 0;
      o.photos = work.map(function (p) { return { id: p.id, name: p.name || '', w: p.w, h: p.h, bytes: p.bytes, memo: L.str(p.memo) }; });
      return L.cleanCase(o);
    }
    openDialog(isNew ? '정비 사례 등록' : '정비 사례 고치기', [errBox, form, dl1, dl2], [
      isNew ? null : h('button', { type: 'button', class: 'btn btn-danger', onclick: function () {
        db.cases = db.cases.filter(function (x) { return x.id !== src.id; });
        saved = true; save(); closeDialog(); dropPhotos(orig.map(function (p) { return p.id; }));
        toast('지웠습니다' + (orig.length ? '(사진 ' + orig.length + '장 포함)' : '')); render();
      } }, '삭제'),
      h('button', { type: 'button', class: 'btn', onclick: closeDialog }, '취소'),
      h('button', { type: 'button', class: 'btn btn-primary', onclick: function () {
        var nc = collect(), e = L.validateCase(nc);
        errBox.textContent = '';
        if (Object.keys(e).length) { add(errBox, h('div', { class: 'alert warn' }, Object.keys(e).map(function (k) { return h('div', null, e[k]); }))); errBox.scrollIntoView({ block: 'nearest' }); return; }
        if (isNew) { nc.id = nextId(); db.cases.push(nc); }
        else db.cases = db.cases.map(function (x) { return x.id === src.id ? nc : x; });
        L.syncEquipments(db);
        saved = true; save(); closeDialog();
        var keep = {}; work.forEach(function (p) { keep[p.id] = true; });
        dropPhotos(orig.filter(function (p) { return !keep[p.id]; }).map(function (p) { return p.id; }));
        work.forEach(function (p) { P.get(p.id).then(function (rec) { if (rec && (rec.memo !== L.str(p.memo) || rec.caseId !== nc.id)) { rec.memo = L.str(p.memo); rec.caseId = nc.id; P.put(rec); } }); });
        toast((isNew ? '등록했습니다' : '고쳤습니다') + (work.length ? '(사진 ' + work.length + '장)' : ''));
        render();
      } }, '저장')], true);
  }

  /* ── ② 정비 사례 ─────────────────────────── */
  function filteredCases() {
    var f = ui.cases, q = L.squash(f.q);
    return db.cases.filter(function (c) {
      if (f.equipment && c.equipment !== f.equipment) return false;
      if (f.cause && L.causeKey(c.cause) !== f.cause) return false;
      if (f.by && (c.repairBy || '-') !== f.by) return false;
      if (f.status && (c.status || '-') !== f.status) return false;
      if (f.from && (!c.date || c.date < f.from)) return false;
      if (f.to && (!c.date || c.date > f.to)) return false;
      if (q && L.squash([c.symptom, c.cause, L.canonCause(c.cause), L.splitSteps(c.steps).join(' '), L.splitList(c.parts).join(' '), c.memo, c.equipment, c.vendor].join(' ')).indexOf(q) < 0) return false;
      return true;
    }).sort(function (a, b) { return String(b.date).localeCompare(String(a.date)); });
  }
  function viewCases() {
    var wrap = h('div');
    add(wrap, pageHead('② 정비 사례', '정비 사례', '외부 업체가 고치고 간 내역, 팀에서 직접 고친 내역을 한 건씩 쌓습니다. 진단·가이드북·비용 현황이 모두 이 사례에서 나옵니다.',
      h('div', { class: 'btn-row' },
        h('button', { type: 'button', class: 'btn btn-primary', onclick: function () { editCase(null); } }, '사례 등록'),
        h('a', { class: 'btn', href: '#/import' }, '내역서 가져오기'))));
    if (!db.cases.length) { add(wrap, emptyNotice()); return wrap; }
    var f = ui.cases;
    function onF(k) { return function (e) { f[k] = e.target.value; render(); }; }
    var qIn = h('input', { type: 'search', value: f.q, placeholder: '낱말(증상·원인·조치·부품)', 'aria-label': '낱말 찾기' });
    qIn.addEventListener('change', onF('q'));
    add(wrap, h('section', { class: 'card filters-card' }, h('div', { class: 'filters' },
      fld('설비', sel([['', '전체']].concat(db.equipments.map(function (e) { return e.name; })), f.equipment, { onchange: onF('equipment') })),
      fld('원인(묶음)', sel([['', '전체']].concat(L.causeList(db.cases).map(function (x) { return [x.key, x.cause + ' (' + x.count + ')']; })), f.cause, { onchange: onF('cause') })),
      fld('외부/자체', sel([['', '전체'], '외부', '자체', ['-', '구분 없음']], f.by, { onchange: onF('by') })),
      fld('처리 상태', sel([['', '전체'], '완료', '진행 중', ['-', '모름']], f.status || '', { onchange: onF('status') })),
      fld('시작일', h('input', { type: 'date', value: f.from, onchange: onF('from') })),
      fld('끝일', h('input', { type: 'date', value: f.to, onchange: onF('to') })),
      fld('낱말', qIn))));
    var list = filteredCases();
    var rows = list.map(function (c) {
      var st = L.splitSteps(c.steps);
      return { cells: [
        photoCell(c), dateText(c), c.equipment || '-', c.symptom,
        causeCell(c),
        st.length ? h('span', null, L.stepCore(L.firstCheck(c)) || st[0], h('small', { class: 'note block' }, '단계 ' + st.length + '개')) : '-',
        L.splitList(c.parts).join(', ') || '-', h('span', null, byBadge(c), c.vendor && c.repairBy !== '자체' ? h('small', { class: 'note block' }, c.vendor) : null), costText(c), statusBadge(c) || '-',
        h('span', { class: 'btn-row nowrap' },
          h('button', { type: 'button', class: 'btn btn-sm', onclick: function () { showCase(c); } }, '보기'),
          h('button', { type: 'button', class: 'btn btn-sm', onclick: function () { editCase(c); } }, '고치기'))
      ] };
    });
    var extCost = list.filter(function (c) { return c.repairBy === '외부'; }).reduce(function (s, c) { return s + L.spent(c); }, 0);
    var allCost = list.reduce(function (s, c) { return s + L.spent(c); }, 0);
    add(wrap, h('section', { class: 'card' },
      h('div', { class: 'list-head' }, h('h2', null, list.length + '건' + (list.length !== db.cases.length ? ' (전체 ' + db.cases.length + '건 중)' : '')),
        h('span', { class: 'note' }, '외부 수리 비용 ' + L.fmtWon(extCost) + ' · 적힌 비용 전체 ' + L.fmtWon(allCost) + '(견적·미기재 제외)'),
        h('div', { class: 'btn-row' },
          h('button', { type: 'button', class: 'btn btn-sm', onclick: function () { writeXlsx('정비사례_' + today() + '.xlsx', { '정비 사례': L.casesToAoa(list) }); } }, '엑셀 내보내기'),
          h('button', { type: 'button', class: 'btn btn-sm', onclick: function () { writeXlsx('정비사례_빈양식.xlsx', { '정비 사례': L.templateAoa() }); } }, '빈 양식 받기'))),
      table(['사진', '일자', '설비', '증상', '원인', '먼저 본 것', '교체 부품', '수리', '비용', '상태', ''], rows, { num: [false, false, false, false, false, false, false, false, true, false], cls: 'cases' }),
      storageNote(h('div', { style: 'margin-top:10px' }))));
    return wrap;
  }

  /* ── 내역서 가져오기 ───────────────────────── */
  function readWorkbook(file) {
    return new Promise(function (resolve, reject) {
      var fr = new FileReader();
      fr.onload = function () {
        try { resolve(XLSX.read(new Uint8Array(fr.result), { type: 'array', cellDates: true })); }
        catch (e) { reject(new Error(file.name + ': 엑셀로 읽지 못했습니다(' + e.message + ')')); }
      };
      fr.onerror = function () { reject(new Error(file.name + ': 파일을 읽지 못했습니다.')); };
      fr.readAsArrayBuffer(file);
    });
  }
  function wbSheets(wb) {
    // 병합 셀(세로로 합친 설비명 등)은 합친 범위 전체에 값을 채워 읽습니다
    return wb.SheetNames.map(function (n) { return { name: n, aoa: L.unmerge(XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: null }), wb.Sheets[n]['!merges']) }; });
  }
  function startImport(fileName, wb) {
    var sheets = wbSheets(wb);
    // 머리행이 가장 잘 맞는 시트를 먼저 고릅니다. 머리 모양이 같은 시트가 여럿이면(2025 · 2026 처럼) 함께 가져오기를 켭니다
    var plans = L.sheetPlans(sheets, db.mapping);
    var best = sheets[0], bestN = -1;
    sheets.forEach(function (s) {
      var hr = L.detectHeaderRow(s.aoa), m = L.autoMap(L.headersOf(s.aoa, hr), null), n = Object.keys(m).length;
      if (n > bestN) { best = s; bestN = n; }
    });
    ui.imp = { fileName: fileName, sheets: sheets, sheet: best ? best.name : '', plans: plans, multi: plans.length > 1 };
    pickSheet(ui.imp.sheet);
    render();
  }
  function pickSheet(name) {
    var imp = ui.imp, sh = imp.sheets.filter(function (s) { return s.name === name; })[0] || imp.sheets[0];
    imp.sheet = sh ? sh.name : ''; imp.aoa = sh ? sh.aoa : [];
    imp.hr = L.detectHeaderRow(imp.aoa);
    remap();
  }
  function remap() {
    var imp = ui.imp;
    imp.headers = L.headersOf(imp.aoa, imp.hr);
    imp.map = L.autoMap(imp.headers, db.mapping);
  }
  /* 지금 고른 시트(열 맞추기 그대로) + 「함께 가져오기」 를 켰으면 머리 모양이 같은 다른 시트(자동 연결) */
  function importRows() {
    var imp = ui.imp, pre = 'i' + Date.now().toString(36) + '_', out = { rows: [], issues: [], sheets: [imp.sheet] };
    var r0 = L.applyMapping(imp.aoa, imp.hr, imp.map, { idPrefix: pre + '0_', source: '엑셀: ' + imp.fileName + ' · ' + imp.sheet, sheet: imp.multi ? imp.sheet : '' });
    out.rows = r0.rows; out.issues = r0.issues;
    if (imp.multi) (imp.plans || []).forEach(function (p, i) {
      if (p.name === imp.sheet) return;
      var sh = imp.sheets.filter(function (x) { return x.name === p.name; })[0];
      var r = L.applyMapping(sh.aoa, p.hr, p.map, { idPrefix: pre + (i + 1) + '_', source: '엑셀: ' + imp.fileName + ' · ' + p.name, sheet: p.name });
      out.rows = out.rows.concat(r.rows); out.issues = out.issues.concat(r.issues); out.sheets.push(p.name);
    });
    return out;
  }
  function sampleWorkbook(kind) {
    var wb = XLSX.utils.book_new();
    if (kind === 'log') {
      var ws = XLSX.utils.aoa_to_sheet(SMP.repairLogAoa());
      XLSX.utils.book_append_sheet(wb, ws, '통합 수리이력');
      var ws2 = XLSX.utils.aoa_to_sheet(SMP.repairLogSummaryAoa()); ws2['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 3 } }];
      XLSX.utils.book_append_sheet(wb, ws2, '요약');
    } else if (kind === 'recommended') {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(L.recommendedAoa()), '수리이력(권장 양식)');
    } else XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(SMP.receiptAoa()), '정비내역');
    return wb;
  }
  function viewImport() {
    var wrap = h('div');
    add(wrap, pageHead('가져오기', '정비 내역서·수리이력 엑셀 가져오기', '외부 업체 정비 내역서·설비 고장 수리이력을 엑셀로 정리해 두었다면 한 번에 사례로 넣습니다. 열 이름이 달라도 「열 맞추기」 에서 연결합니다. 파일은 이 브라우저 안에서만 읽습니다.'));
    var input = h('input', { type: 'file', accept: '.xlsx,.xls,.xlsm,.csv', 'aria-label': '엑셀 파일 고르기', onchange: function (e) {
      var f = e.target.files[0]; e.target.value = '';
      if (!f) return;
      readWorkbook(f).then(function (wb) { startImport(f.name, wb); }, function (err) { toast(err.message, true); });
    } });
    function tryWb(name, wb) { startImport(name, XLSX.read(XLSX.write(wb, { bookType: 'xlsx', type: 'array' }), { type: 'array', cellDates: true })); }
    add(wrap, h('section', { class: 'card' },
      h('div', { class: 'drop' },
        h('label', { class: 'btn btn-primary' }, '엑셀 파일 고르기', input),
        h('button', { type: 'button', class: 'btn', onclick: function () { tryWb('예시데이터_수리이력_지금양식.xlsx (가상)', sampleWorkbook('log')); } }, '예시 수리이력(지금 양식 모양)으로 시험'),
        h('button', { type: 'button', class: 'btn', onclick: function () { tryWb('예시데이터_정비내역서.xlsx (가상)', sampleWorkbook('receipt')); } }, '예시 정비 내역서(칸이 다 찬 것)로 시험'),
        h('button', { type: 'button', class: 'btn', onclick: function () { writeXlsx('수리이력_권장양식.xlsx', { '수리이력(권장 양식)': L.recommendedAoa() }); } }, '권장 양식 받기')),
      h('ul', { class: 'note' },
        h('li', null, '한 줄 = 고장·정비 1건. 위쪽에 제목 줄이 있어도 머리행(일자·설비명·증상·처리일자·수리업체·비용…)을 찾아 읽습니다. 병합된 칸은 범위 전체에 같은 값으로 채워 읽고, 머리 모양이 같은 시트가 여럿이면 함께 가져옵니다.'),
        h('li', null, '원인 칸이 없어도 됩니다 — 「릴레이 접촉 불량으로 … 파손」 처럼 증상에 섞인 원인이나 조치 글에서 짐작하고 「짐작」 표시를 붙입니다. 조치 칸이 없으면 처리일자 칸의 조치 글, 그다음 비고를 조치로 읽습니다.'),
        h('li', null, '일자는 「2025-03 2주차」 「2026-01 (보고서 기준)」 도 읽습니다. 비용은 「미기재」(모름) · 「무상」(0원) · 「40,000원/개」(개당) · 「견적 …원」(실지출과 따로) 를 나눠 셉니다. 업체 칸이 「미기재」 면 외부/자체를 모름으로 둡니다.'),
        h('li', null, '종이 내역서·영수증뿐이라면: 「권장 양식 받기」 에 5~10건만 옮겨 적어도 시작할 수 있습니다.'))));
    var imp = ui.imp;
    if (!imp) return wrap;
    var sheetSel = sel(imp.sheets.map(function (s) { return s.name; }), imp.sheet, { onchange: function (e) { pickSheet(e.target.value); render(); } });
    var hrIn = h('input', { type: 'number', min: '1', max: String(Math.max(1, imp.aoa.length)), value: String(imp.hr + 1), onchange: function (e) {
      var v = Math.max(1, Math.min(imp.aoa.length, Number(e.target.value) || 1)); imp.hr = v - 1; remap(); render();
    } });
    var others = (imp.plans || []).filter(function (p) { return p.name !== imp.sheet; });
    var multiCb = h('input', { type: 'checkbox' }); multiCb.checked = !!imp.multi;
    multiCb.addEventListener('change', function () { imp.multi = multiCb.checked; render(); });
    var grid = h('div', { class: 'map-grid' });
    L.FIELDS.forEach(function (f) {
      var s = sel([['', '— 연결 안 함 —']].concat(imp.headers.map(function (x, i) { return [String(i), L.colName(i) + '열 · ' + x]; })), imp.map[f.id] == null ? '' : String(imp.map[f.id]), { 'aria-label': f.label + ' 열' });
      s.addEventListener('change', function () { if (s.value === '') delete imp.map[f.id]; else imp.map[f.id] = Number(s.value); render(); });
      var ex = imp.map[f.id] != null ? imp.aoa.slice(imp.hr + 1, imp.hr + 4).map(function (r) { return L.str(r && r[imp.map[f.id]]); }).filter(Boolean).join(' / ') : '';
      grid.appendChild(h('div', { class: 'map-row' + (f.req && imp.map[f.id] == null ? ' missing' : '') },
        h('span', { class: 'col' }, f.label, f.req ? h('span', { class: 'req' }, ' 꼭 필요') : null), s, h('span', { class: 'ex', title: ex }, ex || ' ')));
    });
    var missing = L.missingRequired(imp.map);
    var res = missing.length ? null : importRows();
    var merged = res ? L.mergeCases(db.cases, res.rows) : null;
    add(wrap, h('section', { class: 'card' },
      h('h2', null, '열 맞추기 — ' + imp.fileName),
      h('div', { class: 'form-grid' }, fld('시트', sheetSel), fld('머리행(줄 번호)', hrIn, '자동으로 찾았습니다. 틀리면 고쳐 주세요')),
      others.length ? h('label', { class: 'opt', style: 'margin-top:10px' }, multiCb, h('span', null, '머리 모양이 같은 시트 ' + others.length + '개(' + others.map(function (p) { return p.name; }).join(', ') + ')도 함께 가져오기')) : null,
      h('p', { class: 'note', style: 'margin-top:10px' }, '왼쪽 항목마다 내 파일의 어느 열인지 고릅니다. 연결은 저장되어 다음 파일에서 자동으로 맞춰집니다.'),
      grid,
      missing.length ? h('p', { class: 'alert warn' }, '꼭 필요한 항목이 연결되지 않았습니다: ' + missing.join(', ')) : null));
    if (res) {
      var pv = res.rows.slice(0, 6).map(function (c) {
        return [dateText(c), c.equipment || '-', c.symptom, causeCell(c), L.splitSteps(c.steps).join(' / ') || '-',
          h('span', null, byBadge(c), c.vendor && c.repairBy !== '자체' ? h('small', { class: 'note block' }, c.vendor) : null), costText(c), statusBadge(c) || '-'];
      });
      var chk = L.dataCheck(res.rows);
      add(wrap, h('section', { class: 'card' },
        h('h2', null, '미리 보기 — 읽은 사례 ' + res.rows.length + '건' + (res.sheets.length > 1 ? ' (시트 ' + res.sheets.length + '개)' : '')),
        table(['일자', '설비', '증상', '원인(짐작 표시)', '조치', '수리', '비용', '상태'], pv, { cls: 'sims' }),
        res.issues.length ? h('details', { class: 'warn-list' }, h('summary', null, '확인할 줄 ' + res.issues.length + '개'),
          h('ul', null, res.issues.map(function (x) { return h('li', null, x.row + '행: ' + x.msg); }))) : null,
        h('h3', null, '이 파일의 자료 점검'), checkSummary(chk, true),
        h('p', null, '이미 있는 사례와 같은 건(일자·설비·증상·원인이 같음) ' + merged.dup + '건은 건너뛰고, 새 사례 ' + merged.added + '건을 더합니다.'),
        h('div', { class: 'btn-row' },
          h('button', { type: 'button', class: 'btn btn-primary', disabled: !merged.added, onclick: function () {
            var wasSample = db._sample;
            var base = wasSample ? [] : db.cases;
            if (wasSample) { db = L.emptyDb(); clearPhotos(); }
            var m2 = L.mergeCases(base, res.rows);
            var addedRows = m2.cases.slice(base.length);
            addedRows.forEach(function (c) { c.id = nextId(); });
            db.cases = m2.cases;
            db.mapping = L.mappingToSaved(imp.headers, imp.map);
            var eqAdded = L.syncEquipments(db);
            save(); ui.imp = null; ui.diag.result = null;
            toast('사례 ' + m2.added + '건을 더했습니다' + (eqAdded ? ' · 새 설비 ' + eqAdded + '대' : '') + (wasSample ? ' (예시 사례는 지웠습니다)' : ''));
            location.hash = '#/check';
          } }, '사례 ' + merged.added + '건 더하기' + (db._sample ? '(예시 지우고)' : '')),
          h('button', { type: 'button', class: 'btn', onclick: function () { ui.imp = null; render(); } }, '취소'))));
    }
    return wrap;
  }

  /* ── 자료 점검 — 「이렇게 올리면 되나요?」 ─────────── */
  function bar(rate) { return h('span', { class: 'fill-bar', role: 'img', 'aria-label': L.fmtPct(rate) }, h('i', { style: 'width:' + Math.round(rate * 100) + '%' })); }
  function checkSummary(chk, compact) {
    var box = h('div', { class: 'check-sum' });
    add(box, h('div', { class: 'feat-list' }, chk.features.map(function (f) {
      return h('div', { class: 'feat ' + (f.level === '좋음' ? 'good' : f.level === '부족' ? 'bad' : 'mid') }, h('strong', null, f.name), h('span', { class: 'kind ' + (f.level === '좋음' ? 'ok' : f.level === '부족' ? 'danger' : 'obs') }, f.level), h('p', { class: 'note' }, f.text));
    })));
    if (chk.todo.length) add(box, h('p', null, h('strong', null, '더 채우면 좋은 칸: '), chk.todo.map(function (f) { return f.label + ' ' + L.fmtPct(f.rate); }).join(' · ')));
    if (!compact) return box;
    add(box, h('p', { class: 'note' }, '칸별 채움률·중복 의심·권장 양식은 가져온 뒤 「자료 점검」 화면에서 자세히 봅니다.'));
    return box;
  }
  function viewCheck() {
    var wrap = h('div');
    add(wrap, pageHead('자료 점검', '이렇게 올리면 되나요? — 칸별 점검', '지금 들어 있는 사례로 어느 칸이 얼마나 채워졌는지, 어떤 기능이 지금 자료로 되는지, 어떤 칸을 더 채우면 진단·가이드북·비용 현황이 좋아지는지 보여 줍니다.'));
    if (!db.cases.length) { add(wrap, emptyNotice()); return wrap; }
    var chk = L.dataCheck(db.cases);
    var F = {}; chk.fields.forEach(function (f) { F[f.id] = f; });
    var answer = [];
    answer.push('네, 지금 양식 그대로 읽힙니다. 사례 ' + chk.n + '건으로 증상 진단은 바로 됩니다' + (F.cause.inferred ? '(원인 ' + F.cause.inferred + '건은 증상·조치 글에서 짐작).' : '.'));
    if (chk.todo.length) answer.push('더 좋아지려면 ' + chk.todo.slice(0, 4).map(function (f) { return '「' + f.label + '」'; }).join(' · ') + ' 칸을 채워 주세요. 새 칸을 만들기 어렵다면 지금 칸에 「원인: … / 조치: 1) … 2) …」 처럼 적어도 읽습니다.');
    if (db._sample) answer = ['지금은 예시 데이터(가상)입니다. 「내역서 가져오기」 로 실제 수리이력을 넣으면 그 자료로 점검합니다.'];
    add(wrap, h('section', { class: 'card answer' }, h('h2', null, '한 줄 답'), answer.map(function (t) { return h('p', null, t); }), checkSummary(chk, false)));
    add(wrap, h('section', { class: 'card' }, h('h2', null, '칸별 채움'),
      table(['칸', '구분', '채움', '', '내용', '왜 필요한가'], chk.fields.map(function (f) {
        return [f.label, h('span', { class: 'kind ' + (f.need === '필수' ? 'danger' : f.need === '권장' ? 'obs' : 'meta') }, f.need), f.filled + '/' + chk.n, bar(f.rate), f.detail || '-', h('span', { class: 'note' }, f.why)];
      }), { num: [false, false, true], cls: 'check-table' })));
    var byId = {}; db.cases.forEach(function (c) { byId[c.id] = c; });
    function caseLink(id) { var c = byId[id]; return c ? h('button', { type: 'button', class: 'linklike', onclick: function () { showCase(c); } }, dateText(c) + ' ' + (c.equipment || '') + ' — ' + c.symptom) : id; }
    add(wrap, h('section', { class: 'card' }, h('h2', null, '확인할 줄'),
      h('h3', null, '같은 고장이 이어 보고된 것 같은 줄 ' + chk.dups.length + '쌍' + (chk.dupSameCost.length ? ' · 그중 같은 금액 ' + chk.dupSameCost.length + '쌍' : '')),
      h('p', { class: 'note' }, '같은 설비에서 45일 안에 같은 원인(또는 거의 같은 증상)으로 적힌 두 줄입니다. 주간 보고가 이어진 것이면 한 건으로 합치거나 비고에 「앞 건과 같은 고장」 이라고 적어 주세요. 같은 금액이 두 번 적힌 쌍은 비용 현황에서 한 번만 셉니다.'),
      chk.dups.length ? h('ul', { class: 'dup-list' }, chk.dups.map(function (d) {
        return h('li', null, caseLink(d.first), h('span', { class: 'note' }, ' → ' + d.gap + '일 뒤 '), caseLink(d.later), d.sameCost ? h('span', { class: 'kind obs' }, '같은 금액 ' + L.fmtWon(d.cost)) : null);
      })) : h('p', null, '없습니다.'),
      h('h3', null, '아직 끝나지 않은 고장 ' + chk.pending.length + '건'),
      chk.pending.length ? h('ul', null, chk.pending.map(function (id) { return h('li', null, caseLink(id)); })) : h('p', null, '없습니다.'),
      (chk.estimates.length || chk.unitPrices.length) ? h('div', null, h('h3', null, '금액을 다시 볼 줄'),
        h('ul', null, chk.estimates.map(function (id) { return h('li', null, caseLink(id), h('span', { class: 'kind meta' }, '견적 — 실제 지출 금액을 적어 주세요')); })
          .concat(chk.unitPrices.map(function (id) { return h('li', null, caseLink(id), h('span', { class: 'kind meta' }, '개당 단가 — 수량을 함께 적어 주세요')); })))) : null));
    add(wrap, h('section', { class: 'card' }, h('h2', null, '설비 종류 · 부위'),
      h('div', { class: 'guide-grid' },
        h('div', null, h('h3', null, '설비 종류별 건수'), h('p', null, chk.families.map(function (x) { return x.name + ' ' + x.count; }).join(' · ') || '-'),
          h('p', { class: 'note' }, '「압출기 1호기」 「압출 1호기」 「압출기」 를 한 종류로 묶습니다. 같은 종류 사례는 진단에서 조금 앞으로 옵니다.')),
        h('div', null, h('h3', null, '부위별 건수(원인 묶음 기준)'), h('p', null, chk.parts.map(function (x) { return x.name + ' ' + x.count; }).join(' · ') || '-')))));
    var rec = L.recommendedAoa();
    add(wrap, h('section', { class: 'card' }, h('h2', null, '권장 양식 — 꼭 더할 칸은 넷'),
      h('p', null, '지금 양식(연도 · 설비명 · 일자 · 증상 · 처리일자 · 수리업체 · 비용 · 비고 · 출처)도 그대로 읽힙니다. 새로 정리하신다면 아래 모양을 권합니다 — 꼭 더할 칸은 「원인(점검 결과)」 「조치 순서」 「교체 부품」 「외부/자체」 넷이고, 나머지는 있으면 좋은 칸입니다. 아래 두 줄은 지어낸 예시입니다.'),
      table(rec[0], rec.slice(1).map(function (r) { return r.map(function (v) { return typeof v === 'number' ? v.toLocaleString('ko-KR') : String(v); }); }), { cls: 'sims' }),
      h('ul', { class: 'note' },
        h('li', null, '원인은 업체가 말한 「점검 결과」 그대로(「서보드라이버 고장」). 모르면 「원인 미상」.'),
        h('li', null, '조치는 한 칸에 「1) … 2) … 3) …」 로 — 그때 무엇부터 봤는지가 가이드북의 확인 순서가 됩니다.'),
        h('li', null, '비용은 실제 지출만, 견적이면 「견적/실지출」 칸에 견적이라고. 개당 금액이면 수량도 함께.'),
        h('li', null, '증상 사진은 「사진 파일명」 칸에 파일 이름만 적고, 사진은 「정비 사례 → 고치기 → 사진」 에서 붙이면 줄여서 저장됩니다.')),
      h('div', { class: 'btn-row' },
        h('button', { type: 'button', class: 'btn btn-primary', onclick: function () { writeXlsx('수리이력_권장양식.xlsx', { '수리이력(권장 양식)': L.recommendedAoa() }); } }, '권장 양식 받기(예시 2줄)'),
        h('button', { type: 'button', class: 'btn', onclick: function () { writeXlsx('수리이력_권장양식_내자료_' + today() + '.xlsx', { '수리이력': L.toRecommendedAoa(db.cases) }); } }, '지금 자료를 권장 양식으로 내보내기(빈 칸 채우기용)'))));
    return wrap;
  }

  /* ── ③ 자체 정비 가이드북 ───────────────────── */
  function guideCards() {
    return L.causeList(db.cases).filter(function (x) { return !/미상|모름|불명|확인중/.test(x.cause); }).map(function (x) {
      return L.guideCard(L.guideDraft(db.cases, x.key), db.guides[x.key]);
    }).filter(Boolean);
  }
  function viewGuide() {
    var wrap = h('div');
    add(wrap, pageHead('③ 자체 정비 가이드북', '팀 전용 자체 정비 가이드북',
      '정비 사례를 원인별로 모아 초보자용 카드를 자동으로 만듭니다. 외부 기사가 고치고 간 순서가 쌓일수록 카드가 자랍니다. 정비 담당자가 확인하고 고쳐서 씁니다.',
      db.cases.length ? h('div', { class: 'btn-row no-print' }, h('button', { type: 'button', class: 'btn', onclick: function () { printCards(null); } }, '전체 인쇄')) : null));
    if (!db.cases.length) { add(wrap, emptyNotice()); return wrap; }
    var cards = guideCards();
    var want = route().q.k;
    add(wrap, h('nav', { class: 'toc no-print', 'aria-label': '원인 목록' }, cards.map(function (c) {
      return h('a', { class: 'chip' + (want === c.key ? ' on' : ''), href: '#/guide?k=' + encodeURIComponent(c.key) }, c.cause + ' · ' + c.count + '건' + (c.edited.length ? ' · 수정됨' : ''));
    })));
    add(wrap, h('p', { class: 'alert info no-print' }, '이 카드는 과거 사례에서 만든 초안입니다. 「담당자·전문가」 표시가 붙은 단계(전기 측정·배선, 유압 라인 분리·교체, 고온부 분해)부터는 초보자가 하지 않고 멈춥니다.'));
    cards.forEach(function (c) { add(wrap, guideArticle(c)); });
    if (want) setTimeout(function () { var el = document.getElementById('g-' + want); if (el) el.scrollIntoView({ block: 'start' }); }, 30);
    return wrap;
  }
  function guideArticle(c) {
    var ol = h('ol', { class: 'guide-steps' });
    c.steps.forEach(function (s, i) {
      if (i === c.stopIndex) ol.appendChild(h('li', { class: 'stop', 'aria-hidden': 'false' }, '초보자는 여기까지 — 다음 단계부터는 정비 담당자·외부 전문가가 합니다'));
      ol.appendChild(h('li', { class: 'step ' + s.kind + (c.stopIndex >= 0 && i > c.stopIndex ? ' after' : '') }, s.text,
        s.kind === 'expert' ? h('span', { class: 'kind danger' }, '담당자·전문가') : null,
        s.count ? h('span', { class: 'note' }, ' · 사례 ' + s.count + '건') : null));
    });
    var cases = c.caseIds.map(caseById).filter(Boolean);
    return h('article', { class: 'card guide-card', id: 'g-' + c.key },
      h('header', { class: 'guide-head' },
        h('div', null, h('div', { class: 'stage' }, '자체 정비 가이드'), h('h2', null, c.cause),
          h('p', { class: 'note' }, '사례 ' + c.count + '건 · 외부 ' + c.extCount + ' · 자체 ' + c.selfCount + (c.count - c.extCount - c.selfCount ? ' · 구분 모름 ' + (c.count - c.extCount - c.selfCount) : '') + (L.partOf(c.cause) ? ' · 부위: ' + L.partOf(c.cause) : '') + (c.avgExtCost != null ? ' · 외부 수리 평균 ' + L.fmtWon(c.avgExtCost) : '') +
            ' · 설비: ' + c.equipments.map(function (e) { return e.name + '(' + e.count + ')'; }).join(', '))),
        h('div', { class: 'badges' }, h('span', { class: 'kind ' + (c.edited.length ? 'obs' : 'meta') }, c.edited.length ? '수정됨' : '자동 초안'),
          c.count < 2 ? h('span', { class: 'kind obs' }, '사례 1건 — 얕은 초안') : null,
          c.cats.map(function (k) { return h('span', { class: 'kind haz' }, k + ' 위험'); }))),
      h('div', { class: 'guide-grid' },
        h('section', null, h('h3', null, '이런 증상일 때'), h('ul', null, c.symptoms.map(function (s) { return h('li', null, s); })),
          c.tags.length ? h('p', { class: 'note' }, '자주 붙은 태그: ' + c.tags.map(function (t) { return t.name; }).join(', ')) : null),
        h('section', { class: 'safety-box' }, h('h3', null, '안전 먼저'), h('ul', null, c.safety.map(function (s) { return h('li', null, s); })))),
      h('section', null, h('h3', null, '확인 순서'), c.steps.length ? ol : h('p', { class: 'note' }, '조치가 두 단계 이상(「1) … 2) …」) 적힌 사례가 아직 없어 순서를 만들지 못했습니다. 아래 「지난 조치 기록」 을 보고, 다음 정비 때 무엇부터 봤는지 순서대로 적어 주세요.'),
        (c.actions || []).length ? h('div', { class: 'past-actions' }, h('h4', null, '지난 조치 기록(한 줄로 적힌 것)'), h('ul', null, c.actions.map(function (a) {
          return h('li', null, a.text, a.count > 1 ? h('span', { class: 'note' }, ' · ' + a.count + '건') : null, a.kind === 'expert' ? h('span', { class: 'kind danger' }, '담당자·전문가') : null);
        }))) : null),
      h('div', { class: 'guide-grid' },
        h('section', null, h('h3', null, '필요 공구'), h('p', null, c.tools.join(', ') || '-'), h('h3', null, '자주 바꾼 부품'), h('p', null, c.parts.join(', ') || '-')),
        h('section', { class: 'call-box' }, h('h3', null, '이럴 때는 멈추고 전문가를 부릅니다'), h('ul', null, c.callExpert.map(function (s) { return h('li', null, s); })))),
      c.note ? h('section', null, h('h3', null, '메모'), h('p', { class: 'pre' }, c.note)) : null,
      h('p', { class: 'note print-foot' }, '과거 정비 사례 ' + c.count + '건에서 만든 초안' + (c.edited.length ? '을 고친 것' : '') + '입니다. 설비 매뉴얼과 정비 담당자의 확인을 받아 쓰세요.'),
      h('div', { class: 'btn-row no-print' },
        h('button', { type: 'button', class: 'btn btn-sm', onclick: function () { editGuide(c); } }, '고치기'),
        c.edited.length ? h('button', { type: 'button', class: 'btn btn-sm', onclick: function () { delete db.guides[c.key]; save(); toast('자동 초안으로 되돌렸습니다'); render(); } }, '초안으로 되돌리기') : null,
        h('button', { type: 'button', class: 'btn btn-sm', onclick: function () { printCards(c.key); } }, '이 카드만 인쇄'),
        h('button', { type: 'button', class: 'btn btn-sm', onclick: function () { copyText(L.guideToText(c), '가이드 글'); } }, '글로 복사'),
        h('button', { type: 'button', class: 'btn btn-sm', onclick: function () { guideAi(c, cases); } }, 'AI 로 다듬기'),
        h('a', { class: 'btn btn-sm', href: '#/cases' , onclick: function () { ui.cases = { equipment: '', cause: c.key, by: '', q: '', from: '', to: '' }; } }, '근거 사례 ' + c.count + '건 보기')));
  }
  function printCards(key) {
    document.body.classList.toggle('print-one', !!key);
    document.querySelectorAll('.guide-card').forEach(function (el) { el.classList.toggle('print-target', !key || el.id === 'g-' + key); });
    setTimeout(function () { window.print(); }, 50);
  }
  window.addEventListener('afterprint', function () { document.body.classList.remove('print-one'); });
  function editGuide(c, preset) {
    var e = Object.assign({}, db.guides[c.key] || {}, preset || {});
    function val(f) {
      if (e[f] != null) return f === 'note' ? L.str(e[f]) : L.linesOf(e[f]).join('\n');
      if (f === 'steps') return c.steps.map(function (s) { return s.text; }).join('\n');
      if (f === 'note') return c.note || '';
      return (c[f] || []).join('\n');
    }
    var labels = { symptoms: '이런 증상일 때', steps: '확인 순서', tools: '필요 공구', parts: '자주 바꾼 부품', safety: '안전 먼저', callExpert: '이럴 때는 멈추고 전문가를 부릅니다', note: '메모' };
    var hints = { steps: '한 줄에 한 단계. 측정·배선·분해·교체 같은 낱말이 들어간 단계는 자동으로 「담당자·전문가」 가 됩니다.', safety: '표준 전원 차단·잠금(LOTO) 문장은 지워도 맨 앞에 다시 들어갑니다.' };
    var tas = {};
    var form = h('div', { class: 'form-grid' });
    L.GUIDE_EDIT_FIELDS.forEach(function (f) {
      var ta = h('textarea', { rows: f === 'steps' || f === 'safety' || f === 'callExpert' ? '7' : '3' }); ta.value = val(f); tas[f] = ta;
      form.appendChild(fld(labels[f] + ' (한 줄 = 한 항목)', ta, hints[f], 'span-all'));
    });
    openDialog('가이드 고치기 — ' + c.cause, [h('p', { class: 'note' }, '고친 칸만 저장됩니다. 고치지 않은 칸은 사례가 늘면 자동으로 바뀝니다.'), form], [
      h('button', { type: 'button', class: 'btn', onclick: closeDialog }, '취소'),
      h('button', { type: 'button', class: 'btn btn-primary', onclick: function () {
        var out = {}, draft = L.guideCard(L.guideDraft(db.cases, c.key));
        L.GUIDE_EDIT_FIELDS.forEach(function (f) {
          var v = tas[f].value.replace(/\s+$/, '');
          var dv = f === 'steps' ? draft.steps.map(function (s) { return s.text; }).join('\n') : f === 'note' ? '' : (draft[f] || []).join('\n');
          if (v !== dv) out[f] = v;
        });
        out.updatedAt = new Date().toISOString();
        if (Object.keys(out).length > 1) db.guides[c.key] = out; else delete db.guides[c.key];
        save(); closeDialog(); toast('가이드를 저장했습니다'); render();
      } }, '저장')], true);
  }
  function guideAi(c, cases) {
    var pr = h('textarea', { rows: '10', readonly: true, class: 'mono' }); pr.value = L.buildGuidePrompt(c, cases);
    var ans = h('textarea', { rows: '8', placeholder: 'AI 답을 붙여 넣고 「답을 카드에 넣기」 를 누르면, ■ 머리별로 나눠 고치기 창에 채웁니다' });
    openDialog('AI 로 가이드 다듬기 — ' + c.cause, [
      h('p', { class: 'note' }, '초안과 근거 사례를 묶은 프롬프트입니다. 복사해서 회사가 허용한 AI 에 붙여 넣고, 답을 아래에 붙여 넣어 주세요.'),
      fld('프롬프트', pr),
      h('div', { class: 'btn-row', style: 'margin:8px 0' },
        h('button', { type: 'button', class: 'btn btn-primary', onclick: function () { copyText(pr.value, '프롬프트'); } }, '프롬프트 복사'),
        AIP.sendButton('설정한 AI 서버로 보내기', function () { return pr.value; }, function (t) { ans.value = t; toast('AI 답을 받았습니다'); }, { toast: toast, settingsHref: '#/settings' })),
      fld('AI 답', ans)
    ], [
      h('button', { type: 'button', class: 'btn', onclick: closeDialog }, '닫기'),
      h('button', { type: 'button', class: 'btn btn-primary', onclick: function () {
        var parsed = L.parseGuideAnswer(ans.value);
        if (!Object.keys(parsed).length) { toast('「■ 확인 순서」 같은 머리를 찾지 못했습니다. 답을 직접 고치기 창에 옮겨 주세요', true); return; }
        closeDialog(); setTimeout(function () { editGuide(c, parsed); toast('AI 답을 칸별로 넣었습니다. 확인 후 저장해 주세요'); }, 30);
      } }, '답을 카드에 넣기')], true);
  }

  /* ── ④ 비용 현황 ─────────────────────────── */
  function viewCost() {
    var wrap = h('div');
    add(wrap, pageHead('④ 비용 현황', '외부 수리 비용과 자체 처리 비율', '외부 수리를 몇 번, 얼마나 불렀는지와 팀이 직접 고친 비율이 어떻게 바뀌는지 봅니다. 「외부/자체」 와 「비용」 칸이 채워진 사례로 계산합니다.'));
    if (!db.cases.length) { add(wrap, emptyNotice()); return wrap; }
    var all = L.costReport(db.cases, '', '', { dedupe: ui.cost.dedupe });
    var months = L.monthRange(all.from, all.to);
    var f = ui.cost;
    var from = f.from && months.indexOf(f.from) >= 0 ? f.from : all.from, to = f.to && months.indexOf(f.to) >= 0 ? f.to : all.to;
    if (from > to) { var t = from; from = to; to = t; }
    var rep = L.costReport(db.cases, from, to, { dedupe: f.dedupe });
    var dedupeCb = h('input', { type: 'checkbox' }); dedupeCb.checked = f.dedupe !== false;
    dedupeCb.addEventListener('change', function () { f.dedupe = dedupeCb.checked; render(); });
    add(wrap, h('section', { class: 'card filters-card no-print' }, h('div', { class: 'filters' },
      fld('시작 달', sel(months, from, { onchange: function (e) { f.from = e.target.value; render(); } })),
      fld('끝 달', sel(months, to, { onchange: function (e) { f.to = e.target.value; render(); } })),
      h('div', { class: 'btn-row' },
        h('button', { type: 'button', class: 'btn', onclick: function () { exportCost(rep); } }, '엑셀 내보내기'),
        h('button', { type: 'button', class: 'btn', onclick: function () { window.print(); } }, '인쇄'))),
      rep.dups.some(function (d) { return d.sameCost; }) || f.dedupe === false ? h('label', { class: 'opt', style: 'margin-top:8px' }, dedupeCb, h('span', null, '같은 설비·같은 금액으로 이어 보고된 줄은 금액을 한 번만 셈')) : null));
    var T = rep.total;
    add(wrap, h('div', { class: 'tiles' },
      tile('적힌 수리 비용 합계', L.fmtWon(T.allCost), '금액이 적힌 ' + T.costKnown + '건 · ' + from + ' ~ ' + to),
      tile('외부 수리 비용', L.fmtWon(T.extCost), '외부 ' + T.extCount + '건 · 구분 모름 비용 ' + L.fmtWon(T.unknownCost)),
      tile('자체 처리 비율', L.fmtPct(T.selfRatio), '자체 ' + T.selfCount + '건 · 구분 모름 ' + T.unknownCount + '건'),
      tile('견적(실지출 아님)', T.estCount ? L.fmtWon(T.estCost) : '-', T.estCount ? '견적 ' + T.estCount + '건 — 합계에서 뺌' : '진행 중 ' + T.pending + '건')));
    var gaps = [];
    if (T.total - T.costKnown - T.estCount > 0) gaps.push('비용이 「미기재」 인 사례가 ' + (T.total - T.costKnown - T.estCount) + '건이라 합계는 적힌 금액만입니다(0원으로 치지 않음).');
    if (T.unknownCount) gaps.push('수리 업체·외부/자체를 모르는 사례 ' + T.unknownCount + '건은 자체 처리 비율에서 빠집니다.');
    if (rep.dups.length) gaps.push('같은 고장이 이어 보고된 것 같은 줄 ' + rep.dups.length + '쌍' + (rep.dupCost ? '(같은 금액 ' + L.fmtWon(rep.dupCost) + (rep.deduped ? ' — 한 번만 셈' : ' — 두 번 셈 중') + ')' : '') + '. 「자료 점검」 에서 확인해 주세요.');
    if (gaps.length) add(wrap, h('div', { class: 'alert warn' }, gaps.map(function (g) { return h('div', null, g); }), h('a', { href: '#/check' }, '자료 점검 보기')));
    if (rep.early && rep.late && rep.early.total && rep.late.total) {
      add(wrap, h('p', { class: 'alert info' }, '앞 기간(' + from + ' ~ ' + rep.monthly[Math.floor(rep.monthly.length / 2) - 1].month + ') 자체 처리 ' + L.fmtPct(rep.early.selfRatio) + ' · 외부 비용 ' + L.fmtWon(rep.early.extCost) +
        ' → 뒤 기간(' + rep.splitMonth + ' ~ ' + to + ') 자체 처리 ' + L.fmtPct(rep.late.selfRatio) + ' · 외부 비용 ' + L.fmtWon(rep.late.extCost) + '.'));
    }
    if (rep.undated || db.cases.some(function (c) { return !c.repairBy; })) add(wrap, h('p', { class: 'note' },
      (rep.undated ? '일자가 없는 사례 ' + rep.undated + '건은 빠졌습니다. ' : '') + '외부/자체가 비어 있는 사례는 비율 계산에서 빠집니다.'));
    var labels = rep.monthly.map(function (m) { return m.month.slice(2).replace('-', '.'); });
    add(wrap, h('div', { class: 'charts' },
      C.figure('월별 수리 비용(만원, 적힌 금액)', { labels: labels, series: [
        { name: '적힌 비용 전체', values: rep.monthly.map(function (m) { return L.round(m.allCost / 10000, 1); }), color: 'var(--series-1)' },
        { name: '그중 외부', values: rep.monthly.map(function (m) { return L.round(m.extCost / 10000, 1); }), color: 'var(--series-2)' }], unit: '만원' }),
      C.figure('월별 정비 건수', { labels: labels, series: [
        { name: '외부', values: rep.monthly.map(function (m) { return m.extCount; }), color: 'var(--series-2)' },
        { name: '자체', values: rep.monthly.map(function (m) { return m.selfCount; }), color: 'var(--series-3)' },
        { name: '구분 모름', values: rep.monthly.map(function (m) { return m.unknownCount; }), color: 'var(--muted)' }], unit: '건' }),
      C.figure('월별 자체 처리 비율', { labels: labels, series: [{ name: '자체 처리 비율', values: rep.monthly.map(function (m) { return m.selfRatio; }), color: 'var(--series-3)' }], unit: '%', type: 'line', minMax: 1 })));
    add(wrap, h('section', { class: 'card' }, h('h2', null, '월별'),
      table(['달', '전체', '외부', '자체', '구분 모름', '자체 비율', '적힌 비용', '외부 비용', '자체 부품비'], rep.monthly.map(function (m) {
        return [m.month, m.total, m.extCount, m.selfCount, m.unknownCount, L.fmtPct(m.selfRatio), L.fmtWon(m.allCost), L.fmtWon(m.extCost), L.fmtWon(m.selfCost)];
      }), { num: [false, true, true, true, true, true, true, true, true] })));
    function groupTable(title, list, note) {
      return h('section', { class: 'card' }, h('h2', null, title), note ? h('p', { class: 'note' }, note) : null,
        table(['구분', '전체', '외부', '자체', '구분 모름', '적힌 비용', '외부 비용', '진행 중'], list.map(function (r) {
          return [r.label, r.total, r.extCount, r.selfCount, r.unknownCount, L.fmtWon(r.allCost), L.fmtWon(r.extCost), r.pending];
        }), { num: [false, true, true, true, true, true, true, true] }));
    }
    add(wrap, groupTable('설비 종류별', rep.byFamily, '「압출기 1호기」 「압출 1호기」 「압출기」 를 한 종류로 묶었습니다. 호기별은 아래 「설비별」.'));
    add(wrap, groupTable('부위별', rep.byPart, '원인 묶음의 부위(전기·제어 · 냉각 · 모터·구동 …)로 묶었습니다.'));
    add(wrap, groupTable('수리 업체별', rep.byVendor, '업체 이름은 이 브라우저 안에만 있습니다.'));
    add(wrap, h('section', { class: 'card' }, h('h2', null, '설비별'),
      table(['설비', '전체', '외부', '자체', '자체 비율', '적힌 비용', '외부 비용'], rep.byEquipment.map(function (r) {
        var e = equipOf(r.label);
        return [r.label + (e && e.type ? ' · ' + e.type : ''), r.total, r.extCount, r.selfCount, L.fmtPct(r.selfRatio), L.fmtWon(r.allCost), L.fmtWon(r.extCost)];
      }), { num: [false, true, true, true, true, true, true] })));
    add(wrap, h('section', { class: 'card' }, h('h2', null, '원인별'),
      table(['원인', '전체', '외부', '자체', '자체 비율', '적힌 비용', ''], rep.byCause.map(function (r) {
        return [r.label, r.total, r.extCount, r.selfCount, L.fmtPct(r.selfRatio), L.fmtWon(r.allCost),
          h('a', { href: '#/guide?k=' + encodeURIComponent(r.key), class: 'no-print' }, '가이드')];
      }), { num: [false, true, true, true, true, true] })));
    add(wrap, h('section', { class: 'card' }, h('h2', null, '자체 처리 전환 후보'),
      h('p', { class: 'note' }, '팀에서 직접 고친 적이 있는데 외부도 부른 원인입니다. 가이드 카드를 다듬어 공유하면 다음번엔 자체로 처리할 수 있을지 검토해 볼 만합니다(전기·유압·고온 작업은 교육받은 담당자만).'),
      rep.candidates.length ? h('ol', null, rep.candidates.map(function (r) {
        return h('li', null, h('a', { href: '#/guide?k=' + encodeURIComponent(r.key) }, r.label), ' — 외부 ' + r.extCount + '건 · ' + L.fmtWon(r.extCost) + ' / 자체 ' + r.selfCount + '건');
      })) : h('p', null, '아직 없습니다.')));
    return wrap;
  }
  function exportCost(rep) {
    writeXlsx('정비비용현황_' + rep.from + '_' + rep.to + '.xlsx', {
      '월별': [['달', '전체', '외부', '자체', '자체 비율(%)', '외부 비용(원)', '자체 부품비(원)', '정비 시간']].concat(rep.monthly.map(function (m) { return [m.month, m.total, m.extCount, m.selfCount, m.selfRatio == null ? '' : Math.round(m.selfRatio * 100), m.extCost, m.selfCost, m.hours]; })),
      '설비별': [['설비', '전체', '외부', '자체', '자체 비율(%)', '외부 비용(원)', '정비 시간']].concat(rep.byEquipment.map(function (r) { return [r.label, r.total, r.extCount, r.selfCount, r.selfRatio == null ? '' : Math.round(r.selfRatio * 100), r.extCost, r.hours]; })),
      '원인별': [['원인', '전체', '외부', '자체', '자체 비율(%)', '외부 비용(원)', '정비 시간']].concat(rep.byCause.map(function (r) { return [r.label, r.total, r.extCount, r.selfCount, r.selfRatio == null ? '' : Math.round(r.selfRatio * 100), r.extCost, r.hours]; })),
      '종류·부위·업체별': [['나눔', '구분', '전체', '외부', '자체', '구분 모름', '적힌 비용(원)', '외부 비용(원)', '견적(원)', '진행 중']].concat(
        [['설비 종류', rep.byFamily], ['부위', rep.byPart], ['수리 업체', rep.byVendor]].reduce(function (a, g) {
          return a.concat(g[1].map(function (r) { return [g[0], r.label, r.total, r.extCount, r.selfCount, r.unknownCount, r.allCost, r.extCost, r.estCost, r.pending]; }));
        }, []))
    });
  }

  /* ── 설정 ───────────────────────────────── */
  function viewSettings() {
    var wrap = h('div');
    add(wrap, pageHead('설정', '설정', '설비 목록과 매뉴얼 메모, 진단 기준, AI 연결, 백업을 관리합니다.'));
    add(wrap, h('section', { class: 'card' }, h('h2', null, '설비 목록'),
      h('p', { class: 'note' }, '설비 종류가 같으면(예: 사출기끼리) 진단 때 조금 더 무게를 둡니다. 「매뉴얼 메모」 에 고장 대응 표·알람 설명을 붙여 두면 AI 프롬프트에 함께 들어갑니다.'),
      db.equipments.length ? table(['설비', '종류', '모델', '제조사', '매뉴얼 메모', '사례', ''], db.equipments.map(function (e, i) {
        var n = db.cases.filter(function (c) { return c.equipment === e.name; }).length;
        return [e.name, e.type || '-', e.model || '-', e.maker || '-', e.manual ? e.manual.length + '자' : '-', n,
          h('span', { class: 'btn-row nowrap' }, h('button', { type: 'button', class: 'btn btn-sm', onclick: function () { editEquipment(i); } }, '고치기'),
            h('button', { type: 'button', class: 'btn btn-sm btn-danger', disabled: n > 0, title: n ? '이 설비의 사례가 있어 뺄 수 없습니다' : null, onclick: function () { db.equipments.splice(i, 1); save(); render(); } }, '빼기'))];
      }), { num: [false, false, false, false, false, true] }) : h('p', { class: 'note' }, '아직 없습니다.'),
      h('div', { class: 'btn-row', style: 'margin-top:10px' }, h('button', { type: 'button', class: 'btn', onclick: function () { editEquipment(-1); } }, '설비 더하기'))));
    var o = db.opts;
    function num(k, step, min, max) { var i = h('input', { type: 'number', step: step, min: min, max: max, value: String(o[k]) }); i.addEventListener('change', function () { o[k] = Number(i.value); db.opts = L.cleanOpts(o); save(); ui.diag.result = null; }); return i; }
    add(wrap, h('section', { class: 'card' }, h('h2', null, '진단 기준'),
      h('div', { class: 'form-grid' },
        fld('비슷한 사례 몇 건까지', num('topN', '1', '1', '30'), '기본 ' + L.DEFAULT_OPTS.topN),
        fld('최소 유사도(0~1)', num('minScore', '0.01', '0', '1'), '이보다 낮으면 빼기. 기본 ' + L.DEFAULT_OPTS.minScore),
        fld('태그 무게(0~1)', num('tagWeight', '0.05', '0', '1'), '증상 태그가 같을 때의 비중. 기본 ' + L.DEFAULT_OPTS.tagWeight),
        fld('같은 설비 가산점', num('sameEquipBonus', '0.01', '0', '0.5'), '기본 ' + L.DEFAULT_OPTS.sameEquipBonus),
        fld('같은 종류 설비 가산점', num('sameTypeBonus', '0.01', '0', '0.5'), '기본 ' + L.DEFAULT_OPTS.sameTypeBonus),
        fld('같은 부위 가산점', num('partBonus', '0.05', '0', '1'), '증상에 적은 부위(서보·벨트·냉각 …)가 원인과 같을 때. 기본 ' + L.DEFAULT_OPTS.partBonus))));
    add(wrap, AIP.settingsCard({ toast: toast }));
    var restoreIn = h('input', { type: 'file', accept: '.json,application/json', 'aria-label': '백업 파일 고르기', onchange: function (e) {
      var f = e.target.files[0]; e.target.value = ''; if (!f) return;
      var rd = new FileReader();
      rd.onload = function () { try { confirmRestore(L.parseBackup(String(rd.result))); } catch (err) { toast(err.message, true); } };
      rd.readAsText(f);
    } });
    add(wrap, h('section', { class: 'card' }, h('h2', null, '데이터'),
      storageNote(h('div')),
      h('div', { class: 'btn-row' },
        h('button', { type: 'button', class: 'btn btn-primary', onclick: exportBackup }, '백업 내려받기(JSON, 사진 포함)'),
        h('label', { class: 'btn' }, '백업 되살리기', restoreIn),
        h('button', { type: 'button', class: 'btn', onclick: loadSample }, '예시 데이터 불러오기'),
        h('button', { type: 'button', class: 'btn btn-danger', onclick: clearAll }, '모두 지우기')),
      h('p', { class: 'note' }, '정비 사례는 이 브라우저에만 있습니다. 다른 PC 로 옮기거나 브라우저 데이터를 지우기 전에 백업을 받아 두세요.')));
    return wrap;
  }
  function editEquipment(i) {
    var e = i >= 0 ? db.equipments[i] : L.cleanEquipment({});
    var ins = { name: h('input', { value: e.name }), type: h('input', { value: e.type, list: 'typeList', placeholder: '예: 사출기 · 압출기 · 원료 건조기' }), model: h('input', { value: e.model }), maker: h('input', { value: e.maker }),
      manual: h('textarea', { rows: '6', placeholder: '매뉴얼의 고장 대응 표·알람 설명을 붙여 넣기' }), memo: h('textarea', { rows: '2' }) };
    ins.manual.value = e.manual; ins.memo.value = e.memo;
    var types = L.uniq(db.equipments.map(function (x) { return x.type; }).concat(['사출기', '압출기', '원료 건조기', '금형 온도 조절기']));
    openDialog(i >= 0 ? '설비 고치기' : '설비 더하기', [h('div', { class: 'form-grid' },
      fld('설비 이름 *', ins.name, i >= 0 ? '이름을 바꾸면 이 설비의 사례도 함께 바뀝니다' : null), fld('종류', ins.type), fld('모델', ins.model), fld('제조사', ins.maker),
      fld('매뉴얼 메모', ins.manual, null, 'span-all'), fld('비고', ins.memo, null, 'span-all')), h('datalist', { id: 'typeList' }, types.map(function (t) { return h('option', { value: t }); }))], [
      h('button', { type: 'button', class: 'btn', onclick: closeDialog }, '취소'),
      h('button', { type: 'button', class: 'btn btn-primary', onclick: function () {
        var ne = L.cleanEquipment({ name: ins.name.value, type: ins.type.value, model: ins.model.value, maker: ins.maker.value, manual: ins.manual.value, memo: ins.memo.value });
        if (!ne.name) { toast('설비 이름을 적어 주세요', true); return; }
        if (db.equipments.some(function (x, j) { return j !== i && x.name === ne.name; })) { toast('같은 이름의 설비가 있습니다', true); return; }
        if (i >= 0) {
          var old = db.equipments[i].name;
          if (old !== ne.name) db.cases.forEach(function (c) { if (c.equipment === old) c.equipment = ne.name; });
          db.equipments[i] = ne;
        } else db.equipments.push(ne);
        save(); closeDialog(); render();
      } }, '저장')], true);
  }
  function exportBackup() {
    var ids = L.allPhotoIds(db.cases), owner = {};
    db.cases.forEach(function (c) { (c.photos || []).forEach(function (p) { owner[p.id] = c.id; }); });
    toast('백업을 만드는 중입니다');
    Promise.all(ids.map(function (id) {
      return P.get(id).then(function (rec) {
        if (!rec) return null;
        return Promise.all([P.blobToB64(rec.blob), P.blobToB64(rec.thumb)]).then(function (b) {
          return { id: id, caseId: owner[id], name: rec.name || '', type: rec.blob.type || 'image/jpeg', w: rec.w, h: rec.h, bytes: rec.bytes, memo: rec.memo || '', data: b[0], thumb: b[1] };
        });
      });
    })).then(function (photos) {
      var out = L.buildBackup(db, photos.filter(Boolean));
      download('정비사례_백업_' + today() + '.json', new Blob([JSON.stringify(out)], { type: 'application/json' }));
      toast('백업을 내려받았습니다(사례 ' + db.cases.length + '건 · 사진 ' + out.photos.length + '장)');
    }, function () { toast('백업을 만들지 못했습니다', true); });
  }
  function confirmRestore(bk) {
    openDialog('백업 되살리기', [
      h('p', null, '백업(' + (bk.exportedAt ? bk.exportedAt.slice(0, 10) + ' 만듦, ' : '') + '사례 ' + bk.db.cases.length + '건 · 사진 ' + bk.photos.length + '장)으로 바꿉니다.'),
      h('p', null, '지금 이 브라우저에 있는 사례 ' + db.cases.length + '건과 설비 목록·가이드 수정 내용은 지워집니다.')], [
      h('button', { type: 'button', class: 'btn', onclick: closeDialog }, '취소'),
      h('button', { type: 'button', class: 'btn btn-danger', onclick: function () {
        var recs = bk.photos.map(function (x) {
          return { id: x.id, caseId: x.caseId || '', name: x.name || '', w: x.w, h: x.h, bytes: x.bytes, memo: x.memo || '', at: Date.now(),
            blob: P.b64ToBlob(x.data, x.type), thumb: x.thumb ? P.b64ToBlob(x.thumb, 'image/jpeg') : P.b64ToBlob(x.data, x.type) };
        });
        clearPhotos().then(function () { return P.putMany(recs); }).then(function () {
          db = bk.db; save(); closeDialog(); ui.diag.result = null;
          toast('백업을 되살렸습니다(사례 ' + db.cases.length + '건 · 사진 ' + recs.length + '장)'); render();
        }, function () { toast('사진을 저장하지 못했습니다. 저장 공간을 확인해 주세요', true); });
      } }, '바꾸기')]);
  }

  /* ── 다음 단계 ─────────────────────────────── */
  var NEXT = [
    { k: 'equip', t: '대상 설비 목록', d: '설비명·모델·제조사(사진 한 장씩이면 충분합니다). 「설정 → 설비 목록」 에 바로 넣어도 됩니다.' },
    { k: 'manual', t: '설비 매뉴얼 PDF', d: '특히 「알람 목록」 「고장 대응(트러블슈팅) 표」 쪽. 설비마다 「매뉴얼 메모」 에 붙이면 AI 프롬프트에 들어갑니다.' },
    { k: 'receipt', t: '설비 고장 수리이력(엑셀)', d: '2025~2026 수리이력 60건을 받아 2차에 반영했습니다. 「자료 점검」 화면의 「더 채우면 좋은 칸」(원인 · 조치 순서 · 외부/자체 · 비용)을 채워 주시면 진단·가이드북·비용 현황이 좋아집니다.' },
    { k: 'photo', t: '고장 사진·영상과 증상 메모', d: '찍은 날·설비·「무엇이 이상했는지」 한 줄. 사진은 사례에 붙이고, 영상은 보관 위치만 적습니다.' },
    { k: 'top', t: '자주 나는 고장 상위 3~5가지', d: '예시에 넣은 히터 단선·센서·유압 누유·모터 과부하·금형 막힘이 실제와 맞는지, 다른 것이 있는지.' },
    { k: 'safety', t: '팀에서 직접 해도 되는 작업의 선', d: '전기·유압·고온 작업 가운데 사내 규정상 팀원이 할 수 있는 것과 반드시 외부를 부르는 것.' }
  ];
  function viewNext() {
    var wrap = h('div');
    db.nextChecks = db.nextChecks || {};
    add(wrap, pageHead('다음 단계', '실제 자료로 넘어가려면', '1단계는 가상 사례로 만들었고, 2차에서 보내 주신 수리이력 양식을 그대로 읽게 고쳤습니다. 아래 자료가 더 오면 진단·가이드를 맞춥니다.'));
    add(wrap, h('section', { class: 'card' }, h('h2', null, '보내 주실 것'),
      h('ul', { class: 'checklist' }, NEXT.map(function (x) {
        var cb = h('input', { type: 'checkbox' }); cb.checked = !!db.nextChecks[x.k];
        cb.addEventListener('change', function () { db.nextChecks[x.k] = cb.checked; save(); });
        return h('li', null, h('label', { class: 'opt' }, cb, h('span', null, h('span', { class: 't' }, x.t), h('span', { class: 's' }, x.d))));
      }))));
    add(wrap, h('section', { class: 'card' }, h('h2', null, '2단계에서 만드는 것'),
      h('ol', null,
        h('li', null, '수리이력의 빈 칸(원인 · 조치 순서 · 외부/자체 · 비용) 채우기 — 「자료 점검 → 지금 자료를 권장 양식으로 내보내기」 로 받은 표에 적어 다시 가져오기'),
        h('li', null, '증상 사진 한꺼번에 붙이기 — 「사진 파일명」 칸과 사진 폴더를 맞춰 사례마다 자동으로 붙이기(지금은 사례마다 「고치기 → 사진」)'),
        h('li', null, '매뉴얼 PDF 에서 알람·고장 대응 표를 뽑아 진단 근거로 함께 보여 주기(문서 검색, NotebookLM·Dify 부터)'),
        h('li', null, '팀 공유 — 여러 사람이 같은 사례집을 쓰도록 공유 저장소(사내 환경 확인 후)'),
        h('li', null, '사진으로 부위·증상 설명 초안 만들기(사진이 모이면)'))));
    return wrap;
  }

  /* ── 그리기 ─────────────────────────────── */
  var VIEWS = { diagnose: viewDiagnose, cases: viewCases, import: viewImport, check: viewCheck, guide: viewGuide, cost: viewCost, settings: viewSettings, next: viewNext };
  var lastRoute = '';
  function render() {
    var r = route();
    nav(); topbar();
    var y = window.scrollY;
    main.textContent = '';
    add(main, VIEWS[r.id]());
    if (lastRoute !== location.hash) { window.scrollTo(0, 0); lastRoute = location.hash; } else window.scrollTo(0, y);
  }
  window.addEventListener('hashchange', function () { render(); main.focus({ preventScroll: true }); });
  if (!S.available()) save();
  render();
  gcPhotos();
})();
