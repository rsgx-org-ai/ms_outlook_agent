/* eslint @typescript-eslint/no-unused-vars: ["warn", { "caughtErrors": "none" }] -- ES5-style catch (e) kept for older Outlook webviews */
/* RSGx Ideas — Outlook task pane (prototype).
 * Plain browser JS, no build step. Entries are stored in Office roaming
 * settings (per mailbox, follows the user across Outlook clients). When the
 * page is opened outside Outlook it falls back to localStorage so it can be
 * previewed in a normal browser. */
(function () {
  'use strict';

  var STORAGE_KEY = 'rsgxIdeas.entries';
  // Hard stop on the number of entries. The real limit is Outlook's 32 KB
  // roaming-settings budget per add-in, which is roughly 8 long notes or many
  // short ones; a save that exceeds it fails and is reported in the status line.
  var MAX_ENTRIES = 200;
  var HUB_FUNNEL_URL = 'https://ai.rsgx.com/submit-request';
  // The funnel seeds its description box from ?problem= (see the hub's
  // app/submit-request/page.tsx). Keep this under the hub's own 4000-char cap
  // and well inside the ~2000-char URL length that older Outlook webviews and
  // proxies handle reliably.
  var MAX_SEED_CHARS = 1800;

  var $ = function (id) { return document.getElementById(id); };
  var form = $('entry-form');
  var textEl = $('entry-text');
  var statusEl = $('status');
  var listEl = $('entries');
  var emptyEl = $('empty');
  var countEl = $('count');
  var contextRow = $('context-row');
  var contextLine = $('context-line');
  var includeContext = $('include-context');
  var template = $('entry-template');

  var storage = null;     // { load(): entry[], save(entries, cb) }
  var context = null;     // { subject, from } of the current email, if any
  var entries = [];
  var statusTimer = null;

  /* ── Storage backends ─────────────────────────────────────────────── */

  function roamingStorage(settings) {
    return {
      label: 'Saved to your mailbox settings',
      load: function () {
        try { return JSON.parse(settings.get(STORAGE_KEY) || '[]'); } catch (e) { return []; }
      },
      save: function (list, cb) {
        settings.set(STORAGE_KEY, JSON.stringify(list));
        settings.saveAsync(function (result) {
          cb(result.status === Office.AsyncResultStatus.Succeeded ? null : (result.error && result.error.message) || 'Could not save.');
        });
      }
    };
  }

  function localStorageBackend() {
    return {
      label: 'Saved in this browser (preview mode)',
      load: function () {
        try { return JSON.parse(window.localStorage.getItem(STORAGE_KEY) || '[]'); } catch (e) { return []; }
      },
      save: function (list, cb) {
        try { window.localStorage.setItem(STORAGE_KEY, JSON.stringify(list)); cb(null); } catch (e) { cb('Could not save in this browser.'); }
      }
    };
  }

  /* ── Email context (subject + sender) ─────────────────────────────── */

  function readMailContext(item, done) {
    if (!item) return done(null);
    var result = { subject: '', from: '' };
    var pending = 2;
    function finish() { if (--pending === 0) done(result.subject || result.from ? result : null); }

    // Read mode exposes plain values; compose mode exposes *.getAsync.
    if (item.subject && typeof item.subject.getAsync === 'function') {
      item.subject.getAsync(function (r) { if (r.status === 'succeeded') result.subject = r.value || ''; finish(); });
    } else {
      result.subject = typeof item.subject === 'string' ? item.subject : '';
      finish();
    }

    var from = item.from || item.sender;
    if (from && typeof from.getAsync === 'function') {
      from.getAsync(function (r) { if (r.status === 'succeeded' && r.value) result.from = formatAddress(r.value); finish(); });
    } else {
      result.from = from ? formatAddress(from) : '';
      finish();
    }
  }

  function formatAddress(a) {
    if (!a) return '';
    if (a.displayName && a.emailAddress) return a.displayName + ' <' + a.emailAddress + '>';
    return a.displayName || a.emailAddress || '';
  }

  function contextText(ctx) {
    if (!ctx) return '';
    var parts = [];
    if (ctx.subject) parts.push('Email: ' + ctx.subject);
    if (ctx.from) parts.push('From: ' + ctx.from);
    return parts.join(' · ');
  }

  /* ── Rendering ────────────────────────────────────────────────────── */

  function formatTime(iso) {
    var d = new Date(iso);
    if (isNaN(d)) return '';
    var today = new Date();
    var sameDay = d.toDateString() === today.toDateString();
    var time = d.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' });
    if (sameDay) return 'Today ' + time;
    return d.toLocaleDateString([], { day: 'numeric', month: 'short' }) + ' ' + time;
  }

  function render() {
    listEl.textContent = '';
    var sorted = entries.slice().sort(function (a, b) { return (b.createdAt || '').localeCompare(a.createdAt || ''); });
    sorted.forEach(function (entry) {
      var node = template.content.firstElementChild.cloneNode(true);
      node.dataset.id = entry.id;
      node.dataset.type = entry.type;
      node.querySelector('[data-role="type"]').textContent = entry.type === 'note' ? 'Note' : 'Idea';
      var t = node.querySelector('[data-role="time"]');
      t.textContent = formatTime(entry.createdAt);
      t.setAttribute('datetime', entry.createdAt || '');
      node.querySelector('[data-role="text"]').textContent = entry.text;
      var ctx = node.querySelector('[data-role="context"]');
      var ctxText = contextText(entry.context);
      if (ctxText) { ctx.textContent = ctxText; ctx.hidden = false; }
      listEl.appendChild(node);
    });
    emptyEl.hidden = entries.length > 0;
    countEl.textContent = entries.length ? '(' + entries.length + ')' : '';
  }

  function setStatus(msg, isError) {
    statusEl.textContent = msg;
    statusEl.classList.toggle('error', !!isError);
    clearTimeout(statusTimer);
    if (msg && !isError) statusTimer = setTimeout(function () { statusEl.textContent = ''; }, 3000);
  }

  /* ── Actions ──────────────────────────────────────────────────────── */

  function persist(next, onOk, okMessage) {
    storage.save(next, function (err) {
      if (err) { setStatus(err, true); return; }
      entries = next;
      render();
      if (onOk) onOk();
      if (okMessage) setStatus(okMessage);
    });
  }

  function onSave(ev) {
    ev.preventDefault();
    var text = textEl.value.trim();
    if (!text) { setStatus('Write something first.', true); textEl.focus(); return; }
    if (entries.length >= MAX_ENTRIES) { setStatus('Limit of ' + MAX_ENTRIES + ' entries reached — delete some first.', true); return; }
    var type = (form.querySelector('input[name="type"]:checked') || {}).value || 'idea';
    var entry = {
      id: String(Date.now()) + Math.random().toString(36).slice(2, 7),
      type: type,
      text: text,
      context: (context && includeContext.checked) ? context : null,
      createdAt: new Date().toISOString()
    };
    persist([entry].concat(entries), function () { textEl.value = ''; textEl.focus(); }, 'Saved.');
  }

  function onListClick(ev) {
    var btn = ev.target.closest('button[data-action]');
    if (!btn) return;
    var li = btn.closest('.entry');
    var entry = entries.filter(function (e) { return e.id === li.dataset.id; })[0];
    if (!entry) return;
    var action = btn.dataset.action;

    if (action === 'delete') {
      persist(entries.filter(function (e) { return e.id !== entry.id; }), null, 'Deleted.');
    } else if (action === 'copy') {
      copyText(entryClipboardText(entry), function (ok) { setStatus(ok ? 'Copied.' : 'Could not copy — select the text instead.', !ok); });
    } else if (action === 'send') {
      // The funnel reads ?problem= and prefills its description box, so the
      // entry arrives in the form. The text is still copied as a fallback for
      // the rare case where a proxy strips or truncates the query string.
      var seed = entryClipboardText(entry);
      copyText(seed, function () {
        setStatus('Opening the funnel with your idea filled in.');
        openExternal(funnelUrlFor(seed));
      });
    }
  }

  function entryClipboardText(entry) {
    var ctx = contextText(entry.context);
    return entry.text + (ctx ? '\n\n(' + ctx + ')' : '');
  }

  // /submit-request?problem=…&source=outlook — `source` lets the hub show the
  // user why the box is already filled in. encodeURIComponent handles the
  // newlines and parentheses the context line adds.
  function funnelUrlFor(text) {
    var seed = String(text || '').slice(0, MAX_SEED_CHARS);
    return HUB_FUNNEL_URL + '?problem=' + encodeURIComponent(seed) + '&source=outlook';
  }

  // Outlook's webview may block window.open; Office.context.ui.openBrowserWindow
  // hands the URL to the system browser when the host supports it.
  function openExternal(url) {
    var ui = window.Office && Office.context && Office.context.ui;
    if (ui && typeof ui.openBrowserWindow === 'function') {
      try { ui.openBrowserWindow(url); return; } catch (e) { /* fall through */ }
    }
    window.open(url, '_blank', 'noopener');
  }

  function copyText(text, cb) {
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).then(function () { cb(true); }, function () { cb(legacyCopy(text)); });
    } else {
      cb(legacyCopy(text));
    }
  }

  function legacyCopy(text) {
    var ta = document.createElement('textarea');
    ta.value = text;
    ta.setAttribute('readonly', '');
    ta.style.position = 'fixed';
    ta.style.opacity = '0';
    document.body.appendChild(ta);
    ta.select();
    var ok = false;
    try { ok = document.execCommand('copy'); } catch (e) { ok = false; }
    document.body.removeChild(ta);
    return ok;
  }

  /* ── Boot ─────────────────────────────────────────────────────────── */

  function start(backend, mailItem) {
    storage = backend;
    $('storage-mode').textContent = backend.label;
    var loaded = storage.load();
    entries = Array.isArray(loaded) ? loaded.filter(function (e) { return e && typeof e === 'object'; }) : [];
    render();

    readMailContext(mailItem, function (ctx) {
      context = ctx;
      if (ctx) {
        contextLine.textContent = contextText(ctx);
        contextRow.hidden = false;
      }
    });

    form.addEventListener('submit', onSave);
    listEl.addEventListener('click', onListClick);
    // Hub links in the header/footer: same system-browser path as "Send to AI Hub".
    document.addEventListener('click', function (ev) {
      var a = ev.target.closest && ev.target.closest('a[target="_blank"][href]');
      if (a) { ev.preventDefault(); openExternal(a.href); }
    });
    textEl.addEventListener('keydown', function (ev) {
      if ((ev.ctrlKey || ev.metaKey) && ev.key === 'Enter') { ev.preventDefault(); if (form.requestSubmit) { form.requestSubmit(); } else { onSave(ev); } }
    });
    textEl.focus();
  }

  function bootInOffice() {
    var ctx = window.Office && Office.context;
    var mailbox = ctx && ctx.mailbox;
    var settings = ctx && ctx.roamingSettings;
    start(settings ? roamingStorage(settings) : localStorageBackend(), mailbox ? mailbox.item : null);
  }

  var booted = false;
  function bootOnce(fn) { if (!booted) { booted = true; fn(); } }

  if (window.Office && typeof Office.onReady === 'function') {
    Office.onReady(function (info) {
      bootOnce(function () { if (info && info.host) bootInOffice(); else start(localStorageBackend(), null); });
    });
    // Belt and braces: if Office.onReady never resolves (page opened as a plain
    // URL in an older browser) fall back to preview mode.
    setTimeout(function () { bootOnce(function () { start(localStorageBackend(), null); }); }, 5000);
  } else {
    // office.js did not load (offline, blocked CDN, plain browser preview).
    document.addEventListener('DOMContentLoaded', function () { bootOnce(function () { start(localStorageBackend(), null); }); });
    if (document.readyState !== 'loading') bootOnce(function () { start(localStorageBackend(), null); });
  }
})();
