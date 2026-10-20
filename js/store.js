/* 브라우저 저장소 — 정비 사례·설비·가이드 수정은 localStorage(data10-06.db), 사진은 IndexedDB(js/photos.js).
   막혀 있거나 가득 차면 이번 창의 메모리로만 동작하고, 화면 위에 알립니다. */
(function (root) {
  'use strict';
  var KEY_DB = 'data10-06.db';
  var memory = {};
  var ok = true;
  var lastError = '';
  function get(k) {
    try { var v = root.localStorage.getItem(k); return v == null && memory[k] != null ? memory[k] : v; }
    catch (e) { ok = false; return memory[k] == null ? null : memory[k]; }
  }
  function set(k, v) {
    memory[k] = v;
    try { root.localStorage.setItem(k, v); lastError = ''; return true; }
    catch (e) {
      ok = false;
      lastError = (e && (e.name === 'QuotaExceededError' || e.code === 22)) ? 'quota' : 'blocked';
      return false;
    }
  }
  function del(k) {
    delete memory[k];
    try { root.localStorage.removeItem(k); } catch (e) { ok = false; }
  }
  root.FixStore = {
    loadDb: function () {
      var raw = get(KEY_DB);
      if (!raw) return null;
      try { return root.FixLogic.restoreDb(JSON.parse(raw)); } catch (e) { return null; }
    },
    /* true = 브라우저에 저장됨, false = 이번 창의 메모리에만 있음 */
    saveDb: function (db) { return set(KEY_DB, JSON.stringify(db)); },
    clearDb: function () { del(KEY_DB); },
    lastError: function () { return lastError; },
    sizeKb: function () { var v = get(KEY_DB); return v ? Math.round(v.length * 2 / 1024) : 0; },
    available: function () { get(KEY_DB); return ok; }
  };
})(window);
