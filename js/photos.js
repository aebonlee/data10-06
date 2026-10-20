/* 고장 사진 저장소 — IndexedDB (data09-17 의 photos.js 를 옮겨 옴)
 *
 * 사진은 localStorage(보통 5MB 안팎)에 넣으면 몇 장 만에 가득 차므로 IndexedDB 에 둡니다. 정비 사례 글은 localStorage.
 * 저장하는 것은 줄인 사본뿐입니다: 긴 변 1280px JPEG + 목록용 작은 그림(긴 변 240px).
 *   { id, caseId, name, type, w, h, bytes, memo, at, blob, thumb }
 * IndexedDB 를 쓸 수 없는 브라우저(사생활 보호 창 등)에서는 메모리에만 두고, 화면에 알립니다.
 */
(function (root) {
  'use strict';
  var DB_NAME = 'data10-06.photos', STORE = 'photos';
  var memory = {};            // IndexedDB 를 못 쓸 때
  var dbp = null, ok = true;
  function open() {
    if (dbp) return dbp;
    dbp = new Promise(function (resolve) {
      try {
        if (!root.indexedDB) { ok = false; resolve(null); return; }
        var rq = root.indexedDB.open(DB_NAME, 1);
        rq.onupgradeneeded = function () { var d = rq.result; if (!d.objectStoreNames.contains(STORE)) d.createObjectStore(STORE, { keyPath: 'id' }); };
        rq.onsuccess = function () { resolve(rq.result); };
        rq.onerror = function () { ok = false; resolve(null); };
        rq.onblocked = function () { ok = false; resolve(null); };
      } catch (e) { ok = false; resolve(null); }
    });
    return dbp;
  }
  function tx(mode, fn) {
    return open().then(function (d) {
      if (!d) return fn(null);
      return new Promise(function (resolve, reject) {
        var t = d.transaction(STORE, mode), st = t.objectStore(STORE), out;
        Promise.resolve(fn(st)).then(function (v) { out = v; });
        t.oncomplete = function () { resolve(out && out.__req ? out.__req.result : out); };
        t.onerror = function () { reject(t.error); };
        t.onabort = function () { reject(t.error || new Error('저장하지 못했습니다')); };
      });
    });
  }
  function put(rec) { return tx('readwrite', function (st) { if (!st) { memory[rec.id] = rec; return; } st.put(rec); }); }
  function putMany(recs) { return tx('readwrite', function (st) { recs.forEach(function (r) { if (!st) memory[r.id] = r; else st.put(r); }); }); }
  function get(id) { return tx('readonly', function (st) { if (!st) return memory[id] || null; return { __req: st.get(id) }; }); }
  function del(ids) { return tx('readwrite', function (st) { (ids || []).forEach(function (id) { if (!st) delete memory[id]; else st.delete(id); }); }); }
  function keys() { return tx('readonly', function (st) { if (!st) return Object.keys(memory); return { __req: st.getAllKeys() }; }); }
  function all() { return tx('readonly', function (st) { if (!st) return Object.keys(memory).map(function (k) { return memory[k]; }); return { __req: st.getAll() }; }); }
  function clear() { return tx('readwrite', function (st) { if (!st) { memory = {}; return; } st.clear(); }); }

  // ── 그림 줄이기 ─────────────────────────────────────────────
  function loadImage(blob) {
    return new Promise(function (resolve, reject) {
      var url = URL.createObjectURL(blob), img = new Image();
      img.onload = function () { resolve({ img: img, url: url }); };
      img.onerror = function () { URL.revokeObjectURL(url); reject(new Error('이 사진 형식은 읽지 못했습니다(HEIC 등). JPG·PNG 로 바꿔 올려 주세요')); };
      img.src = url;
    });
  }
  function toBlob(canvas, q) {
    return new Promise(function (resolve, reject) {
      if (canvas.toBlob) canvas.toBlob(function (b) { b ? resolve(b) : reject(new Error('사진을 줄이지 못했습니다')); }, 'image/jpeg', q);
      else reject(new Error('이 브라우저는 사진 줄이기를 지원하지 않습니다'));
    });
  }
  // 긴 변을 max 로 줄이고(휴대폰 사진의 세로·가로 방향은 브라우저가 맞춰 그립니다) rotate(0·90·180·270)만큼 돌립니다.
  function draw(img, max, rotate) {
    var L = root.FixLogic;
    var sw = img.naturalWidth || img.width, sh = img.naturalHeight || img.height;
    var s = L.fitSize(sw, sh, max);
    var turn = rotate === 90 || rotate === 270;
    var c = document.createElement('canvas');
    c.width = turn ? s.h : s.w; c.height = turn ? s.w : s.h;
    var g = c.getContext('2d');
    g.fillStyle = '#fff'; g.fillRect(0, 0, c.width, c.height); // PNG 투명 부분은 흰색으로
    g.translate(c.width / 2, c.height / 2);
    if (rotate) g.rotate(rotate * Math.PI / 180);
    g.drawImage(img, -s.w / 2, -s.h / 2, s.w, s.h);
    return c;
  }
  function shrink(blob, rotate) {
    var L = root.FixLogic;
    return loadImage(blob).then(function (x) {
      var big = draw(x.img, L.PHOTO_MAX, rotate || 0), small = draw(x.img, L.THUMB_MAX, rotate || 0);
      URL.revokeObjectURL(x.url);
      return Promise.all([toBlob(big, 0.85), toBlob(small, 0.8)]).then(function (b) {
        return { blob: b[0], thumb: b[1], w: big.width, h: big.height, bytes: b[0].size };
      });
    });
  }

  // ── 백업용 base64 ───────────────────────────────────────────
  function blobToB64(blob) {
    return new Promise(function (resolve, reject) {
      if (!blob) { resolve(''); return; }
      var rd = new FileReader();
      rd.onload = function () { resolve(String(rd.result).replace(/^data:[^,]*,/, '')); };
      rd.onerror = function () { reject(rd.error); };
      rd.readAsDataURL(blob);
    });
  }
  function b64ToBlob(b64, type) {
    var bin = atob(b64 || ''), u = new Uint8Array(bin.length);
    for (var i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i);
    return new Blob([u], { type: type || 'image/jpeg' });
  }
  function newId() { return 'ph' + Date.now().toString(36) + Math.random().toString(36).slice(2, 8); }

  root.FixPhotos = {
    put: put, putMany: putMany, get: get, del: del, keys: keys, all: all, clear: clear,
    shrink: shrink, blobToB64: blobToB64, b64ToBlob: b64ToBlob, newId: newId,
    ready: function () { return open().then(function () { return ok; }); },
    available: function () { return ok; }
  };
})(window);
