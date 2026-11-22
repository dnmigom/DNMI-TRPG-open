/*
 * app.js — 단미 TRPG 시트 보관함
 * - 캐릭터:   비공개 저장소의 characters/*.json (캐릭터 하나에 룰별 시트 여러 장)
 * - 인장:     비공개 저장소의 images/<캐릭터id>/ (또는 바깥 이미지 주소)
 * - 룰별 틀:  비공개 저장소의 templates/*.json
 * - 공유:     룰 시트 1장 + 인장 + 틀을 묶어 비밀 Gist로 올림 (공유 중지 = Gist 삭제)
 */
(function () {
  'use strict';
  const E = window.SheetEngine;
  const API = 'https://api.github.com';
  const CONFIG_KEY = 'trpg-sheet-config';
  const app = document.getElementById('app');
  const nav = document.getElementById('nav');

  // ---------------- 작은 도구들 ----------------
  function h(tag, attrs, ...children) {
    const el = document.createElement(tag);
    for (const [k, v] of Object.entries(attrs || {})) {
      if (v == null || v === false) continue;
      if (k === 'class') el.className = v;
      else if (k === 'value') el.value = v;
      else if (k.startsWith('on')) el.addEventListener(k.slice(2), v);
      else if (k in el && typeof v !== 'string') el[k] = v;
      else el.setAttribute(k, v === true ? '' : v);
    }
    for (const c of children.flat(Infinity)) if (c != null && c !== false) el.append(c instanceof Node ? c : String(c));
    return el;
  }
  function toast(msg, isError) {
    const t = document.getElementById('toast');
    t.textContent = msg;
    t.className = 'show' + (isError ? ' error' : '');
    clearTimeout(toast.timer);
    toast.timer = setTimeout(() => (t.className = ''), isError ? 6000 : 2500);
  }
  // 긴 글 칸: 처음엔 한 줄 높이, 글이 길어지면 그만큼만 늘어남
  function grow(el) {
    if (!el.isConnected) return;
    el.style.height = 'auto';
    const b = el.offsetHeight - el.clientHeight; // 테두리 두께
    el.style.height = (el.scrollHeight + b) + 'px';
  }
  function autoGrow(el) {
    el.rows = 1;
    el.classList.add('auto-grow');
    requestAnimationFrame(() => grow(el));
    return el;
  }
  document.addEventListener('input', e => { if (e.target.classList && e.target.classList.contains('auto-grow')) grow(e.target); });
  let growTimer;
  window.addEventListener('resize', () => {
    clearTimeout(growTimer);
    growTimer = setTimeout(() => document.querySelectorAll('textarea.auto-grow').forEach(grow), 150);
  });

  function show(...nodes) { applyTheme(null); app.replaceChildren(...nodes.flat(Infinity).filter(n => n != null && n !== false)); window.scrollTo(0, 0); }
  function loading(text) { show(h('p', { class: 'muted center' }, text || '불러오는 중…')); }
  function failed(err) {
    show(h('div', { class: 'card' }, h('h2', {}, '문제가 생겼어요'), h('p', {}, err.message || String(err)),
      h('p', {}, h('a', { href: '#/' }, '목록으로'), ' · ', h('a', { href: '#/setup' }, '설정 확인'))));
  }
  const newId = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
  const isSafeUrl = u => /^https:\/\//i.test(u || '');

  function b64encode(str) {
    const bytes = new TextEncoder().encode(str);
    let bin = '';
    for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
    return btoa(bin);
  }
  function b64decode(b64) {
    const bin = atob(b64.replace(/\s/g, ''));
    return new TextDecoder().decode(Uint8Array.from(bin, c => c.charCodeAt(0)));
  }

  // ---------------- 설정 (이 브라우저에만 저장) ----------------
  function loadConfig() {
    try { return JSON.parse(localStorage.getItem(CONFIG_KEY)) || {}; } catch { return {}; }
  }
  function saveConfig(c) {
    try { localStorage.setItem(CONFIG_KEY, JSON.stringify(c)); return true; } catch { return false; }
  }
  let config = loadConfig();

  // 화면 폭 (이 브라우저에만 저장): 비우면 기본 1080px, 아니면 창 너비의 %
  const VIEW_KEY = 'trpg-sheet-view';
  function loadView() { try { return JSON.parse(localStorage.getItem(VIEW_KEY)) || {}; } catch { return {}; } }
  function applyView(v) {
    const pct = Number(v && v.width);
    if (pct >= 30 && pct <= 100) document.documentElement.style.setProperty('--page-max', pct + 'vw');
    else document.documentElement.style.removeProperty('--page-max');
    // 글씨 크기: 10% ~ 150% (모든 글씨가 rem 단위라 한꺼번에 바뀜)
    const fs = Number(v && v.font);
    if (fs >= 10 && fs <= 150 && fs !== 100) document.documentElement.style.fontSize = fs + '%';
    else document.documentElement.style.removeProperty('font-size');
  }
  let view = loadView();
  applyView(view);

  // ---------------- GitHub API ----------------
  function explain(status, data) {
    const m = (data && data.message) || '';
    if (status === 401) return '토큰이 올바르지 않거나 만료됐어요. 설정에서 토큰을 다시 넣어 주세요.';
    if (status === 403 && /rate limit/i.test(m)) return 'GitHub 요청 한도를 넘었어요. 잠시 뒤 다시 시도해 주세요.';
    if (status === 403) return '토큰 권한이 부족해요. Contents(읽기·쓰기)와 Gists(읽기·쓰기) 권한을 확인해 주세요.';
    if (status === 404) return '찾을 수 없어요. 저장소 이름이나 토큰의 저장소 접근 범위를 확인해 주세요.';
    if (status === 409 || status === 422) return '다른 곳에서 먼저 수정된 것 같아요. 새로고침한 뒤 다시 저장해 주세요.';
    return `GitHub 오류 (${status}) ${m}`;
  }
  async function gh(path, { method = 'GET', body, auth = true } = {}) {
    const headers = { Accept: 'application/vnd.github+json', 'X-GitHub-Api-Version': '2022-11-28' };
    if (auth) headers.Authorization = `Bearer ${config.token}`;
    if (body) headers['Content-Type'] = 'application/json';
    let res;
    try {
      res = await fetch(API + path, { method, headers, body: body ? JSON.stringify(body) : undefined, cache: 'no-store' });
    } catch {
      throw new Error('인터넷 연결을 확인해 주세요.');
    }
    if (res.status === 204) return null;
    const data = await res.json().catch(() => null);
    if (!res.ok) { const e = new Error(explain(res.status, data)); e.status = res.status; throw e; }
    return data;
  }
  const repoBase = () => `/repos/${encodeURIComponent(config.owner)}/${encodeURIComponent(config.repo)}`;
  const contentsPath = p => `${repoBase()}/contents/${p.split('/').map(encodeURIComponent).join('/')}`;

  // 방금 저장한 파일 기억하기: GitHub은 저장 직후 잠깐(몇 초~1분) 예전 내용을 돌려줄 때가 있어서,
  // 이 브라우저에서 3분 안에 저장한 파일은 GitHub이 예전 것을 주더라도 방금 저장한 내용을 씀 (탭끼리도 공유)
  const RECENT_KEY = 'trpg-recent-writes', RECENT_MS = 3 * 60 * 1000;
  const memRecent = new Map();
  function recentAll() {
    let all = {};
    try { all = JSON.parse(localStorage.getItem(RECENT_KEY)) || {}; } catch {}
    for (const [k, v] of memRecent) if (!all[k] || all[k].at < v.at) all[k] = v;
    const now = Date.now();
    for (const k of Object.keys(all)) if (now - all[k].at > RECENT_MS) delete all[k];
    return all;
  }
  function rememberWrite(path, entry) {
    entry.at = Date.now();
    memRecent.set(path, entry);
    try {
      const all = recentAll();
      all[path] = entry;
      localStorage.setItem(RECENT_KEY, JSON.stringify(all));
    } catch { /* 저장 공간이 모자라면 이 탭에서만 기억 */ }
  }
  const recentOf = path => recentAll()[path] || null;

  async function readJson(path) {
    const rec = recentOf(path);
    if (rec && rec.deleted) { const e = new Error(`${path} 파일이 없어요.`); e.status = 404; throw e; }
    const d = await gh(contentsPath(path));
    if (rec && rec.sha && d.sha !== rec.sha) return { data: JSON.parse(JSON.stringify(rec.data)), sha: rec.sha };
    let text;
    if (d.content) text = b64decode(d.content);
    else { // 1MB가 넘는 파일은 content가 비어 있어서 raw로 다시 받음
      const r = await gh(`${repoBase()}/git/blobs/${d.sha}`);
      text = b64decode(r.content);
    }
    try { return { data: JSON.parse(text), sha: d.sha }; }
    catch (e) { throw new Error(`${path} 파일의 JSON 형식이 잘못됐어요: ${e.message}`); }
  }
  async function writeJson(path, obj, sha, message) {
    const body = { message, content: b64encode(JSON.stringify(obj, null, 2) + '\n') };
    if (sha) body.sha = sha;
    const d = await gh(contentsPath(path), { method: 'PUT', body });
    rememberWrite(path, { sha: d.content.sha, data: obj });
    return d.content.sha;
  }
  async function deleteFile(path, sha, message) {
    await gh(contentsPath(path), { method: 'DELETE', body: { message, sha } });
    rememberWrite(path, { deleted: true });
  }
  async function listDir(path) {
    let list;
    try { const d = await gh(contentsPath(path)); list = Array.isArray(d) ? d : []; }
    catch (e) { if (e.status === 404) list = []; else throw e; }
    // 방금 만든 파일은 목록에 아직 없을 수 있고, 방금 지운 파일은 아직 남아 있을 수 있어서 맞춰 줌
    const all = recentAll(), prefix = path + '/';
    list = list.filter(f => !(all[f.path] && all[f.path].deleted));
    for (const [p, v] of Object.entries(all)) {
      if (v.deleted || !p.startsWith(prefix) || p.slice(prefix.length).includes('/')) continue;
      if (!list.some(f => f.path === p)) list.push({ type: 'file', name: p.slice(prefix.length), path: p, sha: v.sha });
    }
    return list;
  }

  // ---------------- 템플릿 ----------------
  let templateCache = null;
  async function loadTemplates(force) {
    if (templateCache && !force) return templateCache;
    const files = (await listDir('templates')).filter(f => f.type === 'file' && f.name.endsWith('.json'));
    const entries = await Promise.all(files.map(async f => {
      const key = f.name.replace(/\.json$/, '');
      try { return [key, (await readJson(f.path)).data]; }
      catch (e) { return [key, { rule: `${key} (오류)`, error: e.message, sections: [] }]; }
    }));
    templateCache = Object.fromEntries(entries);
    return templateCache;
  }

  // ---------------- 데이터 보관함 ----------------
  let datasetCache = null;
  async function loadDatasets(force) {
    if (datasetCache && !force) return datasetCache;
    const files = (await listDir('datasets')).filter(f => f.type === 'file' && f.name.endsWith('.json'));
    const entries = await Promise.all(files.map(async f => {
      const key = f.name.replace(/\.json$/, '');
      try { return [key, (await readJson(f.path)).data]; }
      catch (e) { return [key, { name: `${key} (오류)`, error: e.message, fields: [], entries: [] }]; }
    }));
    datasetCache = Object.fromEntries(entries);
    return datasetCache;
  }

  // 룰 틀에 나중에 추가된 칸은 처음 값으로 채움 (이미 값이 있으면 그대로)
  function fillDefaults(template, sheet) {
    sheet.values = sheet.values || {};
    sheet.lists = sheet.lists || {};
    for (const s of template.sections || []) {
      if (s.type === 'grid') {
        sheet.grids = sheet.grids || {};
        if (!sheet.grids[s.id]) sheet.grids[s.id] = { acquired: [], blackout: null, disabled: [] };
      } else if (s.type === 'list') {
        if (sheet.lists[s.id] === undefined) sheet.lists[s.id] = (s.defaultRows || []).map(r => ({ ...r }));
      } else for (const f of s.fields || []) for (const x of [f, ...(f.parts || [])]) {
        if (x.formula !== undefined || sheet.values[x.id] !== undefined) continue;
        if (x.default !== undefined) sheet.values[x.id] = x.default;
      }
    }
    return sheet;
  }

  // 룰 틀이나 데이터를 고쳤을 때, 그걸 쓰는 공유본을 모두 새로 올림
  async function refreshShares({ templateKey, datasetKey }) {
    const [templates, chars] = await Promise.all([loadTemplates(true), loadCharacters()]);
    if (datasetKey) await loadDatasets(true);
    let ok = 0, fail = 0;
    for (const c of chars) for (const k of Object.keys(c.sheets || {})) {
      const sh = c.sheets[k], t = templates[sheetTpl(k, sh)];
      if (!sh.share || !sh.share.gistId || !t) continue;
      const hit = (templateKey && sheetTpl(k, sh) === templateKey) ||
        (datasetKey && (t.sections || []).some(s => E.sourceKeys(s).includes(datasetKey)));
      if (!hit) continue;
      try { await gh(`/gists/${sh.share.gistId}`, { method: 'PATCH', body: await gistBody(c, k, t) }); ok++; }
      catch { fail++; }
    }
    return { ok, fail };
  }

  function blankSheet(template) {
    const values = {}, lists = {};
    for (const s of template.sections || []) {
      if (s.type === 'list') lists[s.id] = (s.defaultRows || []).map(r => ({ ...r }));
      else for (const f of s.fields || []) {
        if (f.formula !== undefined) continue;
        if (f.default !== undefined) values[f.id] = f.default;
        else if (f.type === 'checkbox') values[f.id] = false;
      }
    }
    const now = new Date().toISOString();
    return { values, lists, share: null, created: now, updated: now };
  }

  // ---------------- 상태 & 화면 이동 ----------------
  const state = { dirty: false, hash: location.hash.slice(1) || '/' };
  function setDirty(v) {
    state.dirty = v;
    const btn = document.getElementById('save-btn');
    if (btn) { btn.disabled = !v; btn.textContent = v ? '저장' : '저장됨'; }
  }
  window.addEventListener('beforeunload', e => { if (state.dirty) { e.preventDefault(); e.returnValue = ''; } });
  window.addEventListener('hashchange', () => {
    if (state.dirty && !confirm('저장하지 않은 변경 사항이 있어요. 이동할까요?')) {
      history.replaceState(null, '', '#' + state.hash);
      return;
    }
    setDirty(false);
    route();
  });
  const go = hash => { location.hash = hash; };

  function route() {
    const share = new URLSearchParams(location.search).get('share');
    if (share) { nav.hidden = true; return showShared(share); }
    nav.hidden = false;
    state.hash = location.hash.slice(1) || '/';
    if (!config.token || !config.owner || !config.repo) return showSetup();
    const hash = state.hash;
    if (hash === '/setup') return showSetup();
    const m = hash.match(/^\/c\/([^/]+)(?:\/(.+))?$/);
    if (m) return showCharacter(decodeURIComponent(m[1]), m[2] ? decodeURIComponent(m[2]) : null);
    for (const r of routes) if (r.match(hash)) return r.show(hash);
    return showHome();
  }
  const routes = [];


  // ---------------- 설정 화면 ----------------
  function showSetup() {
    const owner = h('input', { value: config.owner || '', placeholder: '예: my-github-id', autocomplete: 'off' });
    const repo = h('input', { value: config.repo || '', placeholder: '예: trpg-data', autocomplete: 'off' });
    const token = h('input', { type: 'password', value: config.token || '', placeholder: 'github_pat_…', autocomplete: 'off' });
    const status = h('p', { class: 'muted' });
    async function connect(e) {
      e.preventDefault();
      const next = { owner: owner.value.trim(), repo: repo.value.trim(), token: token.value.trim() };
      if (!next.owner || !next.repo || !next.token) { status.textContent = '세 칸을 모두 채워 주세요.'; return; }
      status.textContent = '연결 확인 중…';
      const prev = config; config = next;
      try {
        const r = await gh(repoBase());
        if (!r.private) status.textContent = '⚠ 이 저장소는 공개(public) 상태예요. 시트가 누구에게나 보일 수 있으니 비공개로 바꿔 주세요.';
        await listDir('templates');
        if (!saveConfig(next)) toast('이 브라우저에 설정을 저장하지 못했어요. 새로고침하면 다시 입력해야 해요.', true);
        templateCache = null;
        toast('연결됐어요!');
        if (r.private) go('/');
      } catch (err) {
        config = prev;
        status.textContent = err.message;
      }
    }
    function logout() {
      if (!confirm('이 기기에 저장된 토큰을 지울까요?')) return;
      try { localStorage.removeItem(CONFIG_KEY); } catch {}
      config = {}; templateCache = null;
      showSetup();
    }
    // 보기 설정: 화면 폭
    const pctLabel = h('strong', {}, view.width ? `${view.width}%` : '기본');
    const range = h('input', { type: 'range', min: 40, max: 100, step: 5, value: view.width || 70, oninput: e => {
      view = { ...view, width: Number(e.target.value) };
      pctLabel.textContent = `${view.width}%`;
      applyView(view);
      try { localStorage.setItem(VIEW_KEY, JSON.stringify(view)); } catch {}
    } });
    const viewCard = h('div', { class: 'card narrow view-card' },
      h('h2', {}, '보기 설정'),
      h('p', { class: 'muted' }, '시트와 편집기의 가로 폭을 창 너비의 몇 %로 쓸지 정해요. 이 브라우저에만 저장돼요.'),
      h('div', { class: 'row' }, h('span', {}, '화면 폭'), range, pctLabel,
        h('button', { type: 'button', onclick: () => {
          view = { ...view }; delete view.width;
          try { localStorage.setItem(VIEW_KEY, JSON.stringify(view)); } catch {}
          applyView(view); pctLabel.textContent = '기본'; range.value = 70;
        } }, '기본으로')));
    // 글씨 크기
    const saveView = () => { try { localStorage.setItem(VIEW_KEY, JSON.stringify(view)); } catch {} };
    const fontNow = () => Number(view.font) || 100;
    const fontLabel = h('strong', { class: 'font-pct' }, `${fontNow()}%`);
    const setFont = pct => {
      pct = Math.max(10, Math.min(150, Math.round(pct / 5) * 5));
      view = { ...view, font: pct };
      if (pct === 100) delete view.font;
      fontLabel.textContent = `${pct}%`; fontRange.value = pct;
      applyView(view); saveView();
    };
    const fontRange = h('input', { type: 'range', min: 10, max: 150, step: 5, value: fontNow(), 'aria-label': '글씨 크기', oninput: e => setFont(Number(e.target.value)) });
    const fontCard = h('div', { class: 'card narrow view-card' },
      h('h2', {}, '글씨 크기'),
      h('p', { class: 'muted' }, '사이트 전체 글씨 크기를 10%부터 150%까지 바꿔요. 이 브라우저에만 저장돼요.'),
      h('div', { class: 'row' }, h('span', {}, '글씨'),
        h('button', { type: 'button', class: 'small-btn', 'aria-label': '작게', onclick: () => setFont(fontNow() - 5) }, '−'),
        fontRange,
        h('button', { type: 'button', class: 'small-btn', 'aria-label': '크게', onclick: () => setFont(fontNow() + 5) }, '+'),
        fontLabel,
        h('button', { type: 'button', onclick: () => setFont(100) }, '기본으로')),
      h('p', { class: 'font-sample' }, '가나다 ABC 123 — 이 정도 크기로 보여요.'));
    show(fontCard, viewCard, h('form', { class: 'card narrow', onsubmit: connect },
      h('h2', {}, '연결 설정'),
      h('p', { class: 'muted' }, '시트가 저장될 비공개 저장소와 토큰을 입력해 주세요. 이 정보는 지금 쓰는 브라우저에만 저장돼요.'),
      h('label', { class: 'field' }, h('span', {}, 'GitHub 사용자명'), owner),
      h('label', { class: 'field' }, h('span', {}, '데이터 저장소 이름'), repo),
      h('label', { class: 'field' }, h('span', {}, '토큰 (Fine-grained)'), token),
      status,
      h('div', { class: 'row' },
        h('button', { type: 'submit', class: 'primary' }, '연결'),
        config.token ? h('button', { type: 'button', onclick: logout }, '이 기기에서 로그아웃') : null)));
  }


  // ---------------- 캐릭터 색 (먼셀 10색 + 흰색·회색·검은색) ----------------
  const MUNSELL = [
    ['R', '빨강', '#E03A3E'], ['YR', '주황', '#F28A1E'], ['Y', '노랑', '#F5CB1D'], ['GY', '연두', '#96C43B'],
    ['G', '초록', '#2E9E57'], ['BG', '청록', '#1C9E9B'], ['SK', '하늘', '#6EC1EA'], ['B', '파랑', '#2E86D1'], ['PB', '남색', '#3B4AAE'],
    ['P', '보라', '#8A4FB8'], ['RP', '자주', '#C33F80'], ['PK', '분홍', '#F29BBF'], ['BR', '갈색', '#8A5A3B'],
    ['WH', '흰색', '#F4F3EE'], ['GR', '회색', '#9A9AA0'], ['BK', '검은색', '#222226'],
    // 투톤 이상: 점은 무지개색으로 표시
    ['MX', '투톤 이상', 'conic-gradient(#E03A3E, #F28A1E, #F5CB1D, #2E9E57, #2E86D1, #8A4FB8, #E03A3E)'],
  ];
  const COLOR = Object.fromEntries(MUNSELL.map(([code, name, hex]) => [code, { code, name, hex }]));
  const COLOR_SLOTS = [['hair', '머리'], ['eye1', '눈 1'], ['eye2', '눈 2']];
  // 예전 "무채색(N)"은 회색으로 바꿔서 읽음
  function normalizeColors(c) {
    if (c && c.colors) for (const [k] of COLOR_SLOTS) if (c.colors[k] === 'N') c.colors[k] = 'GR';
    return c;
  }
  // 흰색·회색·검은색은 색상(hue)이 없어서 강조색·배경색을 따로 정함
  const NEUTRAL = {
    // 투톤 이상: 한 가지 색이 없어서 강조색은 보라 계열, 배경은 옅은 라벤더
    MX: { light: { accent: 'hsl(265 45% 48%)', on: '#fff', bg: 'hsl(265 40% 95%)', field: 'hsl(265 30% 98%)', line: 'hsl(265 20% 87%)' },
          dark: { accent: 'hsl(265 60% 74%)', on: '#111', bg: 'hsl(265 18% 12%)', field: 'hsl(265 15% 15%)', line: 'hsl(265 12% 25%)' } },
    WH: { light: { accent: 'hsl(40 4% 46%)', on: '#fff', bg: 'hsl(45 30% 98%)', field: '#ffffff', line: 'hsl(45 12% 88%)' },
          dark: { accent: 'hsl(40 6% 88%)', on: '#111', bg: 'hsl(40 3% 18%)', field: 'hsl(40 3% 21%)', line: 'hsl(40 3% 31%)' } },
    GR: { light: { accent: 'hsl(240 3% 40%)', on: '#fff', bg: 'hsl(240 3% 92%)', field: 'hsl(240 3% 97%)', line: 'hsl(240 3% 84%)' },
          dark: { accent: 'hsl(240 3% 64%)', on: '#111', bg: 'hsl(240 3% 13%)', field: 'hsl(240 3% 16%)', line: 'hsl(240 3% 26%)' } },
    BK: { light: { accent: 'hsl(240 6% 14%)', on: '#fff', bg: 'hsl(240 4% 86%)', field: 'hsl(240 4% 94%)', line: 'hsl(240 4% 76%)' },
          dark: { accent: 'hsl(240 4% 72%)', on: '#111', bg: 'hsl(240 6% 6%)', field: 'hsl(240 6% 9%)', line: 'hsl(240 5% 20%)' } },
  };
  function hexToHsl(hex) {
    const n = parseInt(hex.slice(1), 16);
    const r = (n >> 16 & 255) / 255, g = (n >> 8 & 255) / 255, b = (n & 255) / 255;
    const max = Math.max(r, g, b), min = Math.min(r, g, b), l = (max + min) / 2;
    let hue = 0, sat = 0;
    if (max !== min) {
      const d = max - min;
      sat = l > 0.5 ? d / (2 - max - min) : d / (max + min);
      hue = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
      hue *= 60;
    }
    return [Math.round(hue), Math.round(sat * 100), Math.round(l * 100)];
  }
  const isDark = () => window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches;
  // 강조색: 글씨가 읽히도록 밝기를 맞춤 / 배경색: 같은 색의 파스텔 톤
  function accentOf(code) {
    const c = COLOR[code];
    if (!c) return null;
    if (NEUTRAL[code]) { const n = NEUTRAL[code][isDark() ? 'dark' : 'light']; return { accent: n.accent, on: n.on }; }
    const [hh, ss, ll] = hexToHsl(c.hex);
    return isDark()
      ? { accent: `hsl(${hh} ${Math.min(ss, 70)}% ${Math.max(ll, 62)}%)`, on: '#111' }
      : { accent: `hsl(${hh} ${ss}% ${Math.min(ll, 42)}%)`, on: '#fff' };
  }
  function pastelOf(code) {
    const c = COLOR[code];
    if (!c) return null;
    if (NEUTRAL[code]) { const n = NEUTRAL[code][isDark() ? 'dark' : 'light']; return { bg: n.bg, field: n.field, line: n.line }; }
    const [hh, ss] = hexToHsl(c.hex);
    const s2 = Math.min(ss, 75);
    return isDark()
      ? { bg: `hsl(${hh} ${s2 * 0.35}% 12%)`, field: `hsl(${hh} ${s2 * 0.3}% 15%)`, line: `hsl(${hh} ${s2 * 0.3}% 24%)` }
      : { bg: `hsl(${hh} ${s2}% 93%)`, field: `hsl(${hh} ${s2 * 0.6}% 97%)`, line: `hsl(${hh} ${s2 * 0.5}% 86%)` };
  }
  // 캐릭터의 꾸미기 설정 → { accent, bg, field, line }
  function themeOf(colors, theme) {
    colors = colors || {}; theme = theme || {};
    const out = {};
    const a = theme.accent && colors[theme.accent] && accentOf(colors[theme.accent]);
    if (a) { out.accent = a.accent; out.on = a.on; }
    const p = theme.background && colors[theme.background] && pastelOf(colors[theme.background]);
    if (p) Object.assign(out, p);
    return out;
  }
  function applyTheme(t) {
    const st = document.documentElement.style;
    for (const v of ['--accent', '--on-accent', '--bg', '--field', '--line']) st.removeProperty(v);
    if (!t) return;
    if (t.accent) { st.setProperty('--accent', t.accent); st.setProperty('--on-accent', t.on); }
    if (t.bg) { st.setProperty('--bg', t.bg); st.setProperty('--field', t.field); st.setProperty('--line', t.line); }
  }
  const dot = (code, title) => {
    const c = COLOR[code];
    const el = h('span', { class: 'color-dot' + (c ? '' : ' none'), title: c ? `${title}: ${c.name}` : `${title}: 없음` });
    if (c) el.style.background = c.hex;
    return el;
  };
  const colorDots = colors => h('span', { class: 'color-dots' },
    COLOR_SLOTS.filter(([k]) => colors && colors[k]).map(([k, label]) => dot(colors[k], label)));

  // ---------------- 캐릭터 ----------------
  // characters/<id>.json = { id, name, images: { head, bust, half, full }, sheets: { <룰키>: { values, lists, share } } }
  // 이미지 값: "repo:images/…"(데이터 저장소에 올린 파일) 또는 "https://…"(바깥 주소)
  const IMAGE_KINDS = [['head', '두상'], ['bust', '흉상'], ['half', '반신'], ['full', '전신']];
  const charPath = id => `characters/${id}.json`;
  // 캐릭터의 시트: 키 → 시트. 같은 룰로 여러 장이면 sheet.template에 템플릿 키 (예전 시트는 키 = 템플릿 키)
  const sheetTpl = (k, sh) => (sh && sh.template) || k;
  const tplKeysOf = c => [...new Set(Object.entries(c.sheets || {}).map(([k, sh]) => sheetTpl(k, sh)))];
  const sheetsOfTpl = (c, t) => Object.keys(c.sheets || {}).filter(k => sheetTpl(k, c.sheets[k]) === t)
    .sort((a, b) => String(c.sheets[a].created || '').localeCompare(String(c.sheets[b].created || '')));

  async function loadCharacters() {
    const files = (await listDir('characters')).filter(f => f.name.endsWith('.json'));
    const list = await Promise.all(files.map(f => readJson(f.path).then(r => normalizeColors(r.data))
      .catch(() => ({ id: f.name.replace(/\.json$/, ''), name: `${f.name} (읽기 오류)`, sheets: {}, images: {}, broken: true }))));
    return list.sort((a, b) => String(b.updated || '').localeCompare(String(a.updated || '')));
  }
  async function charactersUsing(key) {
    return (await loadCharacters()).filter(c => tplKeysOf(c).includes(key));
  }
  const firstImageKind = images => (IMAGE_KINDS.find(([k]) => images && images[k]) || [])[0] || null;

  // 이미지 표시용 주소 만들기 (저장소 이미지는 받아서 임시 주소로)
  const MIME = { webp: 'image/webp', png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif' };
  const blobCache = new Map();
  async function repoImageBlob(path) {
    if (blobCache.has(path)) return blobCache.get(path);
    const p = (async () => {
      const res = await fetch(API + contentsPath(path), {
        headers: { Authorization: `Bearer ${config.token}`, Accept: 'application/vnd.github.raw+json' }, cache: 'no-store',
      });
      if (!res.ok) throw new Error(explain(res.status, null));
      const ext = (path.split('.').pop() || '').toLowerCase();
      return new Blob([await res.arrayBuffer()], { type: MIME[ext] || 'application/octet-stream' });
    })();
    blobCache.set(path, p);
    p.catch(() => blobCache.delete(path));
    return p;
  }
  const urlCache = new Map();
  async function imageSrc(ref) {
    if (!ref) return null;
    if (ref.startsWith('repo:')) {
      const path = ref.slice(5);
      if (!urlCache.has(path)) urlCache.set(path, URL.createObjectURL(await repoImageBlob(path)));
      return urlCache.get(path);
    }
    return /^(https:|data:image\/)/i.test(ref) ? ref : null;
  }
  const blobToBase64 = blob => new Promise((ok, no) => {
    const r = new FileReader();
    r.onload = () => ok(String(r.result).split(',')[1]);
    r.onerror = () => no(r.error);
    r.readAsDataURL(blob);
  });
  async function shareableImage(ref) {
    if (!ref) return null;
    if (!ref.startsWith('repo:')) return ref;
    const blob = await repoImageBlob(ref.slice(5));
    return `data:${blob.type};base64,${await blobToBase64(blob)}`;
  }

  // 올린 이미지는 너무 크면 줄여서 저장 (가장 긴 쪽 1600px)
  async function prepareImage(file) {
    if (!file.type.startsWith('image/')) throw new Error('이미지 파일만 올릴 수 있어요.');
    if (file.type === 'image/gif') {
      if (file.size > 4e6) throw new Error('GIF는 4MB 이하만 올릴 수 있어요.');
      return { blob: file, ext: 'gif' };
    }
    const bmp = await createImageBitmap(file);
    const scale = Math.min(1, 1600 / Math.max(bmp.width, bmp.height));
    const canvas = document.createElement('canvas');
    canvas.width = Math.round(bmp.width * scale);
    canvas.height = Math.round(bmp.height * scale);
    canvas.getContext('2d').drawImage(bmp, 0, 0, canvas.width, canvas.height);
    const toBlob = (type, q) => new Promise(ok => canvas.toBlob(ok, type, q));
    let blob = await toBlob('image/webp', 0.9), ext = 'webp';
    if (!blob || blob.type !== 'image/webp') { blob = await toBlob('image/png'); ext = 'png'; }
    const origExt = { 'image/png': 'png', 'image/jpeg': 'jpg', 'image/webp': 'webp' }[file.type];
    if (scale === 1 && origExt && file.size <= blob.size) return { blob: file, ext: origExt };
    if (blob.size > 8e6) throw new Error('이미지가 너무 커요. 더 작은 파일로 올려 주세요.');
    return { blob, ext };
  }

  // ---------------- 홈: 캐릭터 목록 ----------------
  const homeFilter = { q: '', rules: new Set(), hair: new Set(), eye: new Set() }; // 화면을 오가도 필터 유지
  async function showHome() {
    loading();
    try {
      const [templates, chars, oldSheets, order] = await Promise.all([
        loadTemplates(), loadCharacters(), listDir('sheets').then(f => f.filter(x => x.name.endsWith('.json'))), loadOrder(),
      ]);
      const prefs = listPrefs('characters', 'custom');
      const SORTS = [['custom', '직접 배치'], ['name', '이름 가나다순'], ['rule', '룰 순서'], ['updated', '최근 수정순'], ['created', '만든 순']];
      if (!SORTS.some(x => x[0] === prefs.sort)) prefs.sort = 'custom';
      const tplRank = new Map(applyOrder(Object.keys(templates), order.templates).map((k, i) => [k, i]));
      const ruleRank = c => Math.min(...tplKeysOf(c).map(k => tplRank.has(k) ? tplRank.get(k) : 9999), 99999);
      const customIds = () => applyOrder(chars.map(c => c.id), order.characters, true);
      function sorted(list) {
        const pos = new Map(customIds().map((k, i) => [k, i]));
        const base = [...list].sort((a, b) => pos.get(a.id) - pos.get(b.id)); // 같은 값끼리는 직접 배치한 순서대로
        const cmp = {
          name: (a, b) => byName(a.name, b.name),
          rule: (a, b) => ruleRank(a) - ruleRank(b),
          updated: (a, b) => String(b.updated || '').localeCompare(String(a.updated || '')),
          created: (a, b) => String(a.created || '').localeCompare(String(b.created || '')),
        }[prefs.sort];
        return cmp ? base.sort(cmp) : base;
      }
      async function dropped(visibleKeys) {
        const full = prefs.sort === 'custom' ? customIds() : sorted(chars).map(c => c.id);
        const next = mergeVisible(full, visibleKeys);
        order.characters = next;
        if (prefs.sort !== 'custom') { prefs.sort = 'custom'; saveListPrefs('characters', prefs); toast('정렬을 "직접 배치"로 바꿨어요.'); }
        draw();
        try { await saveOrder('characters', next); } catch (e) { toast('순서를 저장하지 못했어요: ' + e.message, true); }
      }
      const search = h('input', { type: 'search', placeholder: '캐릭터 검색', class: 'search', value: homeFilter.q });
      const grid = h('div', { class: 'char-list' });
      const filterBar = h('div', { class: 'filter-bar' });
      const f = homeFilter;
      const eyesOf = c => [c.colors && c.colors.eye1, c.colors && c.colors.eye2].filter(Boolean);
      // 같은 묶음 안에서는 "하나라도", 묶음끼리는 "모두" 맞는 캐릭터
      const passes = (c, skip) =>
        (skip === 'q' || !f.q || String(c.name || '').toLowerCase().includes(f.q.toLowerCase())) &&
        (skip === 'rules' || !f.rules.size || [...f.rules].some(k => tplKeysOf(c).includes(k))) &&
        (skip === 'hair' || !f.hair.size || (c.colors && f.hair.has(c.colors.hair))) &&
        (skip === 'eye' || !f.eye.size || eyesOf(c).some(e => f.eye.has(e)));
      function drawFilters() {
        const ruleKeys = Object.keys(templates).filter(k => chars.some(c => tplKeysOf(c).includes(k)));
        const count = (group, test) => chars.filter(c => passes(c, group) && test(c)).length;
        const toggle = (set, v) => () => { if (set.has(v)) set.delete(v); else set.add(v); draw(); };
        const swatchRow = (group, set, test) => MUNSELL.map(([code, name, hex]) => {
          const n = count(group, c => test(c, code));
          if (!n && !set.has(code)) return null;
          const b = h('button', { type: 'button', class: 'swatch-chip' + (set.has(code) ? ' on' : ''), title: name, onclick: toggle(set, code) },
            h('span', { class: 'color-dot' }), name, h('small', {}, n));
          b.firstChild.style.background = hex;
          return b;
        });
        const groups = [
          ruleKeys.length ? ['룰', ruleKeys.map(k => {
            const t = templates[k];
            const b = h('button', { type: 'button', class: 'filter-chip' + (f.rules.has(k) ? ' on' : ''), onclick: toggle(f.rules, k) },
              t.rule || k, h('small', {}, count('rules', c => tplKeysOf(c).includes(k))));
            b.style.setProperty('--accent', t.accent || '');
            return b;
          })] : null,
          ['머리색', swatchRow('hair', f.hair, (c, code) => c.colors && c.colors.hair === code)],
          ['눈 색', swatchRow('eye', f.eye, (c, code) => eyesOf(c).includes(code))],
        ].filter(g => g && g[1].some(Boolean));
        const active = f.rules.size + f.hair.size + f.eye.size;
        filterBar.replaceChildren(...listTools(prefs, SORTS, () => { saveListPrefs('characters', prefs); draw(); }), ...groups.map(([label, items]) => h('div', { class: 'filter-group' }, h('span', { class: 'filter-label' }, label), items)),
          ...(active ? [h('button', { type: 'button', class: 'ghost small-btn', onclick: () => { f.rules.clear(); f.hair.clear(); f.eye.clear(); draw(); } }, '필터 지우기')] : []));
      }
      function draw() {
        f.q = search.value.trim();
        const list = sorted(chars.filter(c => passes(c)));
        drawFilters();
        grid.className = `char-list size-${prefs.size}`;
        grid.replaceChildren(...(list.length ? list.map(charCard) : [h('p', { class: 'muted' },
          chars.length ? '조건에 맞는 캐릭터가 없어요.' : '아직 캐릭터가 없어요. "새 캐릭터"로 첫 캐릭터를 만들어 보세요.')]));
        sortableCards(grid, dropped);
      }
      function charCard(c) {
        const img = h('img', { alt: '', loading: 'lazy' });
        const thumb = h('div', { class: 'char-thumb' }, img);
        const th = themeOf(c.colors, c.theme);
        if (th.bg) thumb.style.background = th.bg;
        const kind = firstImageKind(c.images);
        if (kind) imageSrc(c.images[kind]).then(src => { if (src) img.src = src; else thumb.classList.add('empty'); }).catch(() => thumb.classList.add('empty'));
        else thumb.classList.add('empty');
        const keys = tplKeysOf(c);
        return h('a', { class: 'char-card', 'data-key': c.id, href: `#/c/${encodeURIComponent(c.id)}` }, thumb,
          h('div', { class: 'char-info' },
            h('div', { class: 'char-title' }, h('strong', {}, c.name || '(이름 없음)'), colorDots(c.colors)),
            h('div', { class: 'rule-chips' }, keys.length ? keys.map(k => {
              const t = templates[k] || {};
              const n = sheetsOfTpl(c, k).length;
              const chip = h('span', { class: 'rule-chip' }, t.rule || k, n > 1 ? ` · ${n}` : '');
              chip.style.setProperty('--accent', t.accent || '');
              return chip;
            }) : h('span', { class: 'muted small' }, '시트 없음')),
            Object.values(c.sheets || {}).some(sh => sh.share && sh.share.gistId) ? h('span', { class: 'badge' }, '공유 중') : null));
      }
      search.addEventListener('input', draw);
      draw();

      const nameInput = h('input', { placeholder: '새 캐릭터 이름', required: true });
      const newForm = h('form', { class: 'card new-char', hidden: true, onsubmit: async e => {
        e.preventDefault();
        const now = new Date().toISOString();
        const c = { id: newId(), name: nameInput.value.trim() || '이름 없는 캐릭터', images: {}, sheets: {}, created: now, updated: now };
        try { await writeJson(charPath(c.id), c, null, `캐릭터 생성: ${c.name}`); go(`/c/${c.id}`); }
        catch (err) { toast(err.message, true); }
      } }, nameInput, h('button', { type: 'submit', class: 'primary' }, '만들기'));

      show(
        h('div', { class: 'toolbar' }, h('h1', {}, '내 캐릭터'), search,
          h('button', { class: 'primary', onclick: () => { newForm.hidden = false; nameInput.focus(); } }, '+ 새 캐릭터')),
        oldSheets.length ? migrationBanner(oldSheets) : null,
        newForm, filterBar, grid);
    } catch (e) { failed(e); }
  }

  // 예전 방식(sheets/*.json, 시트 하나 = 캐릭터 하나)을 캐릭터 구조로 옮기기. 이름이 같으면 한 캐릭터로 합침.
  function migrationBanner(files) {
    const btn = h('button', { class: 'primary', onclick: async () => {
      btn.disabled = true; btn.textContent = '옮기는 중…';
      try {
        const old = await Promise.all(files.map(f => readJson(f.path).then(r => ({ ...r, path: f.path }))));
        const existing = await loadCharacters();
        const byName = new Map(existing.map(c => [c.name, { c, sha: undefined, isNew: false }]));
        for (const { data: s } of old) {
          let entry = byName.get(s.name);
          if (!entry || (entry.c.sheets && entry.c.sheets[s.template])) {
            entry = { c: { id: s.id || newId(), name: s.name, images: {}, sheets: {}, created: s.created, updated: s.updated }, isNew: true };
            if (!byName.has(s.name)) byName.set(s.name, entry); else byName.set(s.name + '#' + s.id, entry);
          }
          entry.c.sheets = entry.c.sheets || {};
          entry.c.sheets[s.template] = { values: s.values || {}, lists: s.lists || {}, share: s.share || null, created: s.created, updated: s.updated };
          entry.touched = true;
        }
        for (const entry of byName.values()) {
          if (!entry.touched) continue;
          let sha = null;
          if (!entry.isNew) sha = (await readJson(charPath(entry.c.id))).sha;
          await writeJson(charPath(entry.c.id), entry.c, sha, `예전 시트 옮기기: ${entry.c.name}`);
        }
        for (const o of old) await deleteFile(o.path, o.sha, '예전 시트 파일 정리');
        toast('모두 옮겼어요!');
        showHome();
      } catch (e) { toast(e.message, true); btn.disabled = false; btn.textContent = '다시 시도'; }
    } }, '캐릭터로 옮기기');
    return h('div', { class: 'card notice row' },
      h('span', { class: 'grow' }, `예전 방식으로 저장된 시트가 ${files.length}개 있어요. 캐릭터 목록으로 옮길까요? 이름이 같은 시트는 한 캐릭터로 합쳐져요. 공유 링크는 그대로 유지돼요.`), btn);
  }

  // ---------------- 공유 ----------------
  function shareUrl(gistId) { return `${location.origin}${location.pathname}?share=${gistId}`; }
  async function snapshot(char, key, template) {
    const images = {};
    for (const [k] of IMAGE_KINDS) {
      if (!char.images || !char.images[k]) continue;
      try { images[k] = await shareableImage(char.images[k]); } catch { /* 이미지를 못 읽으면 빼고 공유 */ }
    }
    const sh = JSON.parse(JSON.stringify(char.sheets[key]));
    fillDefaults(template, sh);
    E.syncLinkedRows(template, sh, await loadDatasets());
    return JSON.stringify({
      format: 'trpg-sheet-share/2', sharedAt: new Date().toISOString(), template,
      character: { name: char.name, images, colors: char.colors || {}, theme: char.theme || {} }, sheet: { values: sh.values, lists: sh.lists, grids: sh.grids || {} },
    });
  }
  async function gistBody(char, key, template) {
    return { description: 'TRPG 캐릭터 시트 공유본', files: { 'sheet.json': { content: await snapshot(char, key, template) } } };
  }

  // 코코포리아용 데이터를 클립보드에 복사. 실패하면 직접 복사할 수 있게 글상자를 보여 줌.
  async function copyCcfolia(template, sheet, externalUrl, anchor) {
    let out;
    try { out = E.toCcfolia(template, sheet, { externalUrl }); }
    catch (e) { return toast('코코포리아 데이터를 만들지 못했어요: ' + e.message, true); }
    const text = JSON.stringify(out.clipboard);
    const old = document.getElementById('ccfolia-fallback');
    if (old) old.remove();
    try {
      await navigator.clipboard.writeText(text);
      toast(out.warnings.length
        ? `복사했어요. 다만 확인이 필요한 곳이 있어요: ${out.warnings[0]}`
        : '복사했어요! 코코포리아 방 화면에서 붙여넣기(Ctrl+V) 하세요.', out.warnings.length > 0);
    } catch {
      const ta = h('textarea', { rows: 4, readOnly: true, value: text, onfocus: e => e.target.select() });
      anchor.before(h('div', { id: 'ccfolia-fallback', class: 'card notice' },
        h('p', {}, '자동 복사가 막혀 있어요. 아래 글을 전부 선택해서 복사한 뒤 코코포리아에 붙여넣어 주세요.'), ta));
      ta.focus();
    }
  }

  // ---------------- 인장(이미지) 칸: 왼쪽에 따로 ----------------
  // edit: { pending: {kind: {blob, ext, url}}, onChange } 를 주면 편집 가능
  function renderPortrait(images, edit) {
    images = images || {};
    const has = k => !!(images[k] || (edit && edit.pending[k]));
    const viewOnly = !edit || edit.viewOnly; // 읽기 모드: 있는 이미지만, 고치는 버튼 없음
    const kinds = viewOnly ? IMAGE_KINDS.filter(([k]) => has(k)) : IMAGE_KINDS;
    const panel = h('aside', { class: 'portrait-panel' });
    if (!kinds.length) { panel.classList.add('no-image'); return panel; }
    let current = (kinds.find(([k]) => has(k)) || kinds[0])[0];
    const tabs = h('div', { class: 'portrait-tabs', role: 'tablist' });
    const frame = h('div', { class: 'portrait-frame' });
    const controls = h('div', { class: 'portrait-controls' });

    async function srcOf(k) {
      if (edit && edit.pending[k]) return edit.pending[k].url;
      return imageSrc(images[k]);
    }
    function drawTabs() {
      tabs.replaceChildren(...(kinds.length > 1 || !viewOnly ? kinds.map(([k, label]) => h('button', {
        type: 'button', role: 'tab', class: 'portrait-tab' + (k === current ? ' active' : '') + (has(k) ? '' : ' empty'),
        'aria-selected': String(k === current), onclick: () => { current = k; drawAll(); },
      }, label)) : []));
    }
    async function drawFrame() {
      const k = current;
      if (!has(k)) {
        frame.replaceChildren(h('div', { class: 'portrait-empty' }, edit ? `${IMAGE_KINDS.find(x => x[0] === k)[1]} 이미지가 없어요` : ''));
        return;
      }
      frame.replaceChildren(h('div', { class: 'portrait-empty' }, '불러오는 중…'));
      try {
        const src = await srcOf(k);
        if (k !== current) return;
        frame.replaceChildren(src ? h('img', { src, alt: IMAGE_KINDS.find(x => x[0] === k)[1] }) : h('div', { class: 'portrait-empty' }, '이미지를 표시할 수 없어요'));
      } catch (e) {
        if (k === current) frame.replaceChildren(h('div', { class: 'portrait-empty' }, '이미지를 불러오지 못했어요'));
      }
    }
    function drawControls() {
      if (viewOnly) return;
      const k = current;
      const file = h('input', { type: 'file', accept: 'image/*', hidden: true, onchange: async e => {
        const f = e.target.files[0];
        if (!f) return;
        try {
          const { blob, ext } = await prepareImage(f);
          if (edit.pending[k]) URL.revokeObjectURL(edit.pending[k].url);
          edit.pending[k] = { blob, ext, url: URL.createObjectURL(blob) };
          edit.onChange(); drawAll();
        } catch (err) { toast(err.message, true); }
      } });
      const urlBox = h('div', { class: 'row tight', hidden: true },
        h('input', { type: 'url', placeholder: 'https://… 이미지 주소', value: images[k] && !images[k].startsWith('repo:') ? images[k] : '' }),
        h('button', { type: 'button', onclick: e => {
          const v = e.target.previousSibling.value.trim();
          if (!/^https:\/\//i.test(v)) return toast('https:// 로 시작하는 주소를 넣어 주세요.', true);
          if (edit.pending[k]) { URL.revokeObjectURL(edit.pending[k].url); delete edit.pending[k]; }
          images[k] = v; edit.onChange(); drawAll();
        } }, '적용'));
      controls.replaceChildren(file,
        h('div', { class: 'row tight' },
          h('button', { type: 'button', onclick: () => file.click() }, has(k) ? '파일 바꾸기' : '파일 올리기'),
          h('button', { type: 'button', onclick: () => { urlBox.hidden = !urlBox.hidden; } }, '주소로 넣기'),
          has(k) ? h('button', { type: 'button', class: 'danger ghost', onclick: () => {
            if (edit.pending[k]) { URL.revokeObjectURL(edit.pending[k].url); delete edit.pending[k]; }
            images[k] = null; edit.onChange(); drawAll();
          } }, '지우기') : null),
        urlBox,
        h('p', { class: 'muted small' }, '저장을 눌러야 반영돼요. 큰 이미지는 자동으로 줄여서 저장해요.'));
    }
    function drawAll() { drawTabs(); drawFrame(); drawControls(); }
    drawAll();
    panel.append(tabs, frame, controls);
    return panel;
  }

  // ---------------- 캐릭터 화면 ----------------
  async function showCharacter(id, wantKey) {
    loading();
    let char, sha, templates, datasets;
    try {
      [{ data: char, sha }, templates, datasets] = await Promise.all([readJson(charPath(id)), loadTemplates(true), loadDatasets(true)]);
    } catch (e) { return failed(e); }
    char.images = char.images || {};
    char.sheets = char.sheets || {};
    char.colors = char.colors || {};
    normalizeColors(char);
    char.theme = char.theme || {};
    const pending = {};        // 아직 저장 안 한 새 이미지
    const replaced = new Set(); // 저장할 때 지울 예전 저장소 이미지
    const originalImages = { ...char.images };
    const markDirty = () => setDirty(true);

    // 읽기 / 편집 모드 (마지막에 고른 모드를 이 브라우저에 기억)
    let editing = false;
    try { editing = localStorage.getItem('trpg-sheet-mode') === 'edit'; } catch {}
    function setMode(edit) {
      editing = edit;
      try { localStorage.setItem('trpg-sheet-mode', edit ? 'edit' : 'read'); } catch {}
      pickingSlot = null;
      drawHead(); drawColorBar(); drawPortrait(); drawTabs(); drawSheet();
    }

    const sheetKeys = () => Object.keys(char.sheets)
      .sort((a, b) => String(char.sheets[a].created || '').localeCompare(String(char.sheets[b].created || '')));
    let currentKey = sheetKeys().includes(wantKey) ? wantKey : sheetKeys()[0] || null;

    // ----- 저장 -----
    async function save(btn) {
      if (btn) btn.disabled = true;
      try {
        for (const [k, p] of Object.entries(pending)) {
          const path = `images/${char.id}/${k}-${Date.now().toString(36)}.${p.ext}`;
          await gh(contentsPath(path), { method: 'PUT', body: { message: `인장 올리기: ${char.name}`, content: await blobToBase64(p.blob) } });
          blobCache.set(path, Promise.resolve(p.blob));
          urlCache.set(path, p.url);
          char.images[k] = 'repo:' + path;
          delete pending[k];
        }
        for (const [k] of IMAGE_KINDS) {
          const before = originalImages[k];
          if (before && before.startsWith('repo:') && before !== char.images[k]) replaced.add(before.slice(5));
          if (!char.images[k]) delete char.images[k];
        }
        char.updated = new Date().toISOString();
        sha = await writeJson(charPath(char.id), char, sha, `캐릭터 저장: ${char.name}`);
        Object.assign(originalImages, char.images);
        for (const k of Object.keys(originalImages)) if (!char.images[k]) delete originalImages[k];
        if (replaced.size) {
          const listing = await listDir(`images/${char.id}`);
          for (const f of listing) if (replaced.has(f.path)) { try { await deleteFile(f.path, f.sha, '예전 인장 정리'); } catch {} }
          replaced.clear();
        }
        let shareFail = null;
        for (const k of sheetKeys()) {
          const s = char.sheets[k];
          const t = templates[sheetTpl(k, s)];
          if (!s.share || !s.share.gistId || !t) continue;
          try { await gh(`/gists/${s.share.gistId}`, { method: 'PATCH', body: await gistBody(char, k, t) }); }
          catch (e) { shareFail = e; }
        }
        setDirty(false);
        drawPortrait();
        if (shareFail) toast('저장은 됐지만 공유본 갱신에 실패했어요: ' + shareFail.message, true);
        else toast('저장했어요.');
      } catch (e) { toast(e.message, true); if (btn) btn.disabled = false; }
    }

    async function removeCharacter(btn) {
      if (!confirm(`"${char.name}" 캐릭터와 모든 룰 시트, 인장을 삭제할까요? 되돌릴 수 없어요.`)) return;
      btn.disabled = true;
      try {
        for (const k of sheetKeys()) {
          const g = char.sheets[k].share && char.sheets[k].share.gistId;
          if (g) { try { await gh(`/gists/${g}`, { method: 'DELETE' }); } catch (e) { if (e.status !== 404) throw e; } }
        }
        for (const f of await listDir(`images/${char.id}`)) { try { await deleteFile(f.path, f.sha, '인장 삭제'); } catch {} }
        await deleteFile(charPath(char.id), sha, `캐릭터 삭제: ${char.name}`);
        setDirty(false);
        toast('삭제했어요.');
        go('/');
      } catch (e) { toast(e.message, true); btn.disabled = false; }
    }

    // ----- 화면 -----
    const portraitSlot = h('div', { class: 'portrait-slot' });
    const ruleTabs = h('div', { class: 'rule-tabs' });
    const sheetArea = h('div', { class: 'sheet-area' });
    const sheetTools = h('div', { class: 'sheet-tools row' }); // 공유 · 코코포리아 버튼 (읽기/편집 버튼 왼쪽)
    const layout = h('div', { class: 'char-layout' });
    function drawPortrait() {
      // 편집할 때는 인장 칸을 맨 위에 가로로 놓아서 시트가 화면 너비를 다 쓰게
      layout.classList.toggle('editing', editing);
      portraitSlot.replaceChildren(renderPortrait(char.images, { pending, onChange: markDirty, viewOnly: !editing }));
    }

    function selectSheet(key) {
      currentKey = key;
      history.replaceState(null, '', `#/c/${encodeURIComponent(char.id)}${key ? '/' + encodeURIComponent(key) : ''}`);
      state.hash = location.hash.slice(1);
      drawTabs();
      drawSheet();
    }

    // 룰 버튼 "(룰 이름) · n개" → 누르면 아래에 그 룰로 쓴 시트들이 나열됨
    let openRule = currentKey ? sheetTpl(currentKey, char.sheets[currentKey]) : null, listOpen = true;
    const lastOfRule = {};
    const titleOf = k => {
      const sh = char.sheets[k];
      if (sh && sh.title && String(sh.title).trim()) return sh.title;
      const t = sheetTpl(k, sh), list = sheetsOfTpl(char, t);
      return `${(templates[t] || {}).rule || t} ${list.indexOf(k) + 1}`;
    };
    function addSheet(t) {
      const k = char.sheets[t] ? `${t}~${Date.now().toString(36)}` : t;
      const sh = blankSheet(templates[t]);
      sh.template = t;
      char.sheets[k] = sh;
      openRule = t; listOpen = true;
      markDirty();
      selectSheet(k);
    }
    function drawTabs() {
      if (currentKey) { openRule = sheetTpl(currentKey, char.sheets[currentKey]); lastOfRule[openRule] = currentKey; }
      const choices = Object.keys(templates).filter(k => !templates[k].error);
      const picker = h('select', {}, choices.map(k => h('option', { value: k }, templates[k].rule || k)));
      const addBox = h('div', { class: 'add-sheet', hidden: true },
        choices.length ? [picker, h('button', { type: 'button', class: 'primary', onclick: () => addSheet(picker.value) }, '추가')]
          : h('span', { class: 'muted small' }, '추가할 수 있는 템플릿이 없어요.'),
        h('span', { class: 'muted small' }, '이미 있는 룰도 고를 수 있어요. 같은 룰로 시트를 여러 장 만들 수 있어요.'),
        h('a', { href: '#/template/new', class: 'small-link' }, '+ 새 시트 템플릿 만들기'));
      const accent = t => themeOf(char.colors, char.theme).accent || (templates[t] || {}).accent || '';
      const rules = [...new Set(sheetKeys().map(k => sheetTpl(k, char.sheets[k])))];
      const ruleRow = h('div', { class: 'rule-row' },
        ...rules.map(t => {
          const list = sheetsOfTpl(char, t);
          const shared = list.some(k => char.sheets[k].share && char.sheets[k].share.gistId);
          const active = t === openRule;
          const tab = h('button', { type: 'button', class: 'rule-tab' + (active ? ' active' : ''), 'aria-expanded': String(active && listOpen),
            onclick: () => {
              if (active) { listOpen = !listOpen; drawTabs(); return; }
              openRule = t; listOpen = true;
              selectSheet(list.includes(lastOfRule[t]) ? lastOfRule[t] : list[0]);
            } },
            `${(templates[t] || {}).rule || `${t} (틀 없음)`} · ${list.length}개`,
            shared ? h('span', { class: 'dot', title: '공유 중인 시트가 있어요' }) : null,
            h('span', { class: 'caret' }, active && listOpen ? '▴' : '▾'));
          tab.style.setProperty('--accent', accent(t));
          return tab;
        }),
        ...(editing ? [h('button', { type: 'button', class: 'rule-tab add', onclick: () => { addBox.hidden = !addBox.hidden; } }, '+ 룰 시트 추가')] : []));
      let sheetRow = null;
      if (openRule && listOpen) {
        sheetRow = h('div', { class: 'sheet-row-tabs' },
          sheetsOfTpl(char, openRule).map(k => {
            const sh = char.sheets[k];
            return h('button', { type: 'button', class: 'sheet-tab' + (k === currentKey ? ' active' : ''), onclick: () => selectSheet(k) },
              titleOf(k), sh.share && sh.share.gistId ? h('span', { class: 'dot', title: '공유 중' }) : null);
          }),
          editing && templates[openRule] ? h('button', { type: 'button', class: 'sheet-tab add', onclick: () => addSheet(openRule) }, '+ 이 룰로 시트 하나 더') : null);
        sheetRow.style.setProperty('--accent', accent(openRule));
      }
      ruleTabs.replaceChildren(ruleRow, ...(sheetRow ? [sheetRow] : []), ...(editing ? [addBox] : []));
    }

    function drawSheet() {
      const key = currentKey;
      sheetTools.replaceChildren(); // 룰 시트를 다 그린 뒤에 버튼을 채움
      if (!key) {
        sheetArea.replaceChildren(h('div', { class: 'card empty-sheet' },
          h('p', {}, '아직 이 캐릭터의 룰 시트가 없어요.'),
          editing ? h('p', { class: 'muted' }, '위의 "+ 룰 시트 추가"에서 룰을 골라 주세요.')
            : h('button', { class: 'primary', onclick: () => setMode(true) }, '편집 모드로 바꿔서 룰 시트 추가하기')));
        return;
      }
      const sheet = char.sheets[key];
      const tplKey = sheetTpl(key, sheet);
      const template = templates[tplKey];
      if (!template) {
        sheetArea.replaceChildren(h('div', { class: 'card notice' }, `이 시트의 템플릿(templates/${tplKey}.json)을 찾을 수 없어요.`));
        return;
      }
      const shareBox = h('span', { class: 'share-tools row tight' });
      const busy = (b, on) => { if (b) b.disabled = on; };
      async function startShare(b) {
        busy(b, true);
        try {
          const g = await gh('/gists', { method: 'POST', body: { ...(await gistBody(char, key, template)), public: false } });
          sheet.share = { gistId: g.id, since: new Date().toISOString() };
          await save();
          drawShare(); drawTabs();
        } catch (e) { toast(e.message, true); busy(b, false); }
      }
      async function stopShare(b) {
        if (!confirm('공유를 중지할까요? 이미 보낸 링크는 더 이상 열리지 않아요.')) return;
        busy(b, true);
        try {
          try { await gh(`/gists/${sheet.share.gistId}`, { method: 'DELETE' }); } catch (e) { if (e.status !== 404) throw e; }
          sheet.share = null;
          await save();
          drawShare(); drawTabs();
        } catch (e) { toast(e.message, true); busy(b, false); }
      }
      function drawShare() {
        if (sheet.share && sheet.share.gistId) {
          const url = shareUrl(sheet.share.gistId);
          shareBox.replaceChildren(h('span', { class: 'badge', title: url }, '공유 중'),
            h('button', { title: url, onclick: async () => {
              try { await navigator.clipboard.writeText(url); toast('링크를 복사했어요.'); }
              catch { window.prompt('이 링크를 복사해서 보내 주세요.', url); }
            } }, '링크 복사'),
            h('button', { class: 'danger', onclick: e => stopShare(e.target) }, '공유 중지'));
        } else {
          shareBox.replaceChildren(
            h('button', { title: '지금은 나만 볼 수 있어요. 누르면 읽기 전용 링크를 만들어요.', onclick: e => startShare(e.target) }, '공유 허용'));
        }
      }
      async function removeSheet(b) {
        if (!confirm(`"${titleOf(key)}" 시트를 이 캐릭터에서 삭제할까요? 저장하면 되돌릴 수 없어요.`)) return;
        busy(b, true);
        try {
          if (sheet.share && sheet.share.gistId) {
            try { await gh(`/gists/${sheet.share.gistId}`, { method: 'DELETE' }); } catch (e) { if (e.status !== 404) throw e; }
          }
          delete char.sheets[key];
          await save();
          selectSheet(sheetsOfTpl(char, tplKey)[0] || sheetKeys()[0] || null);
        } catch (e) { toast(e.message, true); busy(b, false); }
      }
      drawShare();
      sheetTools.replaceChildren(shareBox,
        h('button', { class: 'ccfolia-btn', title: '이 캐릭터를 코코포리아에 붙여넣을 수 있게 복사해요',
          onclick: () => copyCcfolia(template, { name: char.name, values: sheet.values, lists: sheet.lists, grids: sheet.grids },
            sheet.share && sheet.share.gistId ? shareUrl(sheet.share.gistId) : undefined, sheetBody) }, '코코포리아로 복사'));
      // 시트 이름 (같은 룰 시트가 여러 장일 때 구분용)
      const titleInput = h('input', { class: 'sheet-title-input', value: sheet.title || '', placeholder: titleOf(key), 'aria-label': '시트 이름',
        oninput: e => {
          if (e.target.value.trim()) sheet.title = e.target.value; else delete sheet.title;
          markDirty(); drawTabs();
        } });
      sheetArea.replaceChildren(...[
        !editing ? null : h('div', { class: 'sheet-actions row' },
          h('label', { class: 'mini-field sheet-title-field' }, h('span', {}, '시트 이름'), titleInput),
          h('a', { class: 'button-link', href: `#/template/${encodeURIComponent(tplKey)}` }, '템플릿 편집'),
          h('button', { class: 'danger ghost', onclick: e => removeSheet(e.target) }, '이 시트 삭제')),
        sheetBody = renderSheet(template, fillDefaults(template, sheet),
          { readonly: !editing, onChange: markDirty, accent: themeOf(char.colors, char.theme).accent, datasets })].filter(Boolean));
      sheetArea.classList.toggle('reading', !editing);
    }
    let sheetBody = null;

    // ----- 색 고르기 & 시트 꾸미기 -----
    const colorBar = h('div', { class: 'color-bar' });
    const picker = h('div', { class: 'swatch-pop', hidden: true });
    let pickingSlot = null;
    function applyCharTheme() {
      applyTheme(themeOf(char.colors, char.theme));
      drawTabs(); drawSheet();
    }
    function drawColorBar() {
      colorBar.hidden = !editing;
      if (!editing) { picker.hidden = true; return; }
      const slotBtn = ([k, label]) => h('button', { type: 'button', class: 'color-slot' + (pickingSlot === k ? ' open' : ''),
        onclick: () => { pickingSlot = pickingSlot === k ? null : k; drawColorBar(); } },
        dot(char.colors[k], label), h('span', {}, label), h('small', {}, COLOR[char.colors[k]] ? COLOR[char.colors[k]].name : '없음'));
      const themePick = (prop, label, noneLabel) => h('div', { class: 'theme-pick' },
        h('span', { class: 'filter-label' }, label),
        [[null, noneLabel], ...COLOR_SLOTS].map(([k, name]) => {
          const usable = !k || char.colors[k];
          const b = h('button', { type: 'button', disabled: !usable, title: usable ? '' : '먼저 이 색을 골라 주세요',
            class: 'seg' + ((char.theme[prop] || null) === k ? ' on' : ''),
            onclick: () => { if (k) char.theme[prop] = k; else delete char.theme[prop]; markDirty(); drawColorBar(); applyCharTheme(); } },
            k ? dot(char.colors[k], name) : null, name);
          return b;
        }));
      colorBar.replaceChildren(
        h('div', { class: 'color-slots' }, COLOR_SLOTS.map(slotBtn)),
        h('div', { class: 'theme-picks' }, themePick('accent', '강조색', '룰 기본'), themePick('background', '배경색', '기본')));
      if (pickingSlot) {
        const label = COLOR_SLOTS.find(x => x[0] === pickingSlot)[1];
        picker.replaceChildren(h('span', { class: 'filter-label' }, `${label} 색`),
          ...[[null, '없음', null], ...MUNSELL].map(([code, name, hex]) => {
            const b = h('button', { type: 'button', class: 'swatch-chip' + ((char.colors[pickingSlot] || null) === code ? ' on' : ''), title: name,
              onclick: () => {
                if (code) char.colors[pickingSlot] = code; else delete char.colors[pickingSlot];
                for (const prop of ['accent', 'background']) if (char.theme[prop] && !char.colors[char.theme[prop]]) delete char.theme[prop];
                pickingSlot = null; markDirty(); drawColorBar(); applyCharTheme();
              } }, h('span', { class: 'color-dot' + (hex ? '' : ' none') }), name);
            if (hex) b.firstChild.style.background = hex;
            return b;
          }));
        picker.hidden = false;
      } else picker.hidden = true;
    }

    const head = h('div', { class: 'sheet-head' });
    function drawHead() {
      const modeSwitch = h('div', { class: 'mode-switch', role: 'group', 'aria-label': '보기 방식' },
        h('button', { type: 'button', class: editing ? '' : 'on', 'aria-pressed': String(!editing), onclick: () => editing && setMode(false) }, '읽기'),
        h('button', { type: 'button', class: editing ? 'on' : '', 'aria-pressed': String(editing), onclick: () => !editing && setMode(true) }, '편집'));
      head.replaceChildren(
        h('div', { class: 'title-row' },
          h('a', { class: 'back-link', href: '#/', title: '캐릭터 목록으로' }, '← 캐릭터 목록'),
          editing
            ? h('input', { class: 'sheet-name', value: char.name || '', 'aria-label': '캐릭터 이름', oninput: e => { char.name = e.target.value; markDirty(); } })
            : h('h1', { class: 'sheet-name' }, char.name || '(이름 없음)', colorDots(char.colors))),
        h('div', { class: 'row' }, sheetTools, modeSwitch,
          editing || state.dirty
            ? h('button', { id: 'save-btn', class: 'primary', disabled: !state.dirty, onclick: e => save(e.target) }, state.dirty ? '저장' : '저장됨')
            : null,
          editing ? h('button', { class: 'save-list-btn', onclick: async e => {
            if (state.dirty) { await save(e.target); if (state.dirty) return; }
            go('/');
          } }, '저장하고 목록으로') : null,
          editing ? h('button', { class: 'danger ghost', onclick: e => removeCharacter(e.target) }, '캐릭터 삭제') : null));
    }
    drawHead();
    show(
      head,
      colorBar, picker,
      (layout.replaceChildren(portraitSlot, h('div', { class: 'char-main' }, ruleTabs, sheetArea)), layout));
    drawPortrait();
    drawColorBar();
    applyTheme(themeOf(char.colors, char.theme));
    selectSheet(currentKey);

    // 다른 탭에서 룰 틀·데이터를 고치고 돌아오면 새 룰 틀로 다시 그림 (입력해 둔 값은 그대로)
    let checking = false, lastCheck = Date.now();
    const sig = () => JSON.stringify(Object.keys(char.sheets).map(k => [templates[sheetTpl(k, char.sheets[k])], datasets]));
    async function recheck() {
      if (!document.body.contains(head)) {
        document.removeEventListener('visibilitychange', onVisible); window.removeEventListener('focus', onVisible);
        return;
      }
      if (checking || document.hidden || Date.now() - lastCheck < 1500) return;
      checking = true; lastCheck = Date.now();
      try {
        const before = sig();
        const [t2, d2] = await Promise.all([loadTemplates(true), loadDatasets(true)]);
        templates = t2; datasets = d2;
        if (sig() !== before && document.body.contains(head)) {
          drawTabs(); drawSheet();
          toast('템플릿이 바뀌어서 새 내용으로 다시 그렸어요.');
        }
      } catch { /* 연결이 잠깐 안 되면 다음에 다시 */ }
      checking = false;
    }
    const onVisible = () => { recheck(); };
    document.addEventListener('visibilitychange', onVisible);
    window.addEventListener('focus', onVisible);

    const keyHandler = e => {
      if (!document.body.contains(head)) return document.removeEventListener('keydown', keyHandler);
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') { e.preventDefault(); if (state.dirty) save(document.getElementById('save-btn')); }
    };
    document.addEventListener('keydown', keyHandler);
  }

  // ---------------- 공유된 시트 (읽기 전용) ----------------
  async function showShared(gistId) {
    if (!/^[0-9a-f]{8,}$/i.test(gistId)) return failed(new Error('공유 링크가 올바르지 않아요.'));
    loading('공유된 시트를 불러오는 중…');
    try {
      let g;
      try { g = await gh(`/gists/${gistId}`, { auth: false }); }
      catch (e) { if (e.status === 404) throw new Error('공유가 중지됐거나 없는 링크예요.'); throw e; }
      const file = g.files && g.files['sheet.json'];
      if (!file) throw new Error('시트 공유본이 아니에요.');
      let text = file.content;
      if (file.truncated) text = await (await fetch(file.raw_url, { cache: 'no-store' })).text();
      const snap = JSON.parse(text);
      let name, images, sheet;
      let colors = {}, theme = {};
      if (snap.format === 'trpg-sheet-share/2') {
        name = snap.character.name; sheet = snap.sheet;
        images = snap.character.images || {}; colors = normalizeColors({ colors: snap.character.colors || {} }).colors; theme = snap.character.theme || {};
      }
      else if (snap.format === 'trpg-sheet-share/1') { name = snap.sheet.name; images = {}; sheet = snap.sheet; }
      else throw new Error('지원하지 않는 공유본 형식이에요.');
      document.title = `${name} — 공유된 시트`;
      const note = h('p', { class: 'muted share-note' }, `읽기 전용 공유본 · ${new Date(snap.sharedAt).toLocaleString('ko-KR')} 기준`);
      show(
        h('div', { class: 'sheet-head' },
          h('h1', { class: 'sheet-name' }, name, colorDots(colors)),
          h('span', { class: 'rule-tag' }, (snap.template && snap.template.rule) || ''),
          h('button', { class: 'ccfolia-btn', onclick: () => copyCcfolia(snap.template, { name, values: sheet.values, lists: sheet.lists, grids: sheet.grids }, location.href, note) }, '코코포리아로 복사')),
        note,
        h('div', { class: 'char-layout' }, renderPortrait(images, null),
          h('div', { class: 'char-main' }, renderSheet(snap.template, sheet, { readonly: true, accent: themeOf(colors, theme).accent }))));
      applyTheme(themeOf(colors, theme));
    } catch (e) { failed(e); }
  }


  // ---------------- 자유 배치(격자) ----------------
  // template.canvas = { cols: 12 } 이면 칸 묶음(카드)마다 pos {x, y, w, h}, 칸 묶음 안의 칸마다 pos, 칸 묶음의 가로 칸 수 s.cols
  const CANVAS_COLS = 12;
  const isCanvas = t => !!(t && t.canvas);
  const posOf = (o, cols) => {
    const p = o.pos || {};
    const w = Math.max(1, Math.min(cols, Math.floor(p.w) || 1));
    return { x: Math.max(0, Math.min(cols - w, Math.floor(p.x) || 0)), y: Math.max(0, Math.floor(p.y) || 0), w, h: Math.max(1, Math.floor(p.h) || 1) };
  };
  const overlaps = (a, b) => a.x < b.x + b.w && b.x < a.x + a.w && a.y < b.y + b.h && b.y < a.y + a.h;
  // 겹치지 않게 위로 붙여서 정리. fixed(옮기는 중인 것)는 그 자리에 두고 나머지가 비켜 감
  function settle(items, cols, fixed) {
    const placed = fixed ? [fixed] : [];
    const rest = items.filter(it => it !== fixed).sort((a, b) => a.y - b.y || a.x - b.x);
    for (const it of rest) {
      it.w = Math.max(1, Math.min(cols, it.w)); it.x = Math.max(0, Math.min(cols - it.w, it.x));
      let y = 0;
      while (placed.some(p => overlaps(p, { ...it, y }))) y++;
      it.y = y; placed.push(it);
    }
    return items;
  }
  // 빈 자리 중 가장 위·왼쪽
  function freeSpot(items, cols, w, h) {
    w = Math.min(w, cols);
    for (let y = 0; ; y++) for (let x = 0; x + w <= cols; x++) {
      const cand = { x, y, w, h };
      if (!items.some(p => overlaps(p, cand))) return cand;
    }
  }
  // 예전 방식(한 줄에 n개 · 한 줄에 칸 n개) 템플릿을 격자 배치로 바꿈 (모양은 최대한 비슷하게)
  function toCanvas(t) {
    if (isCanvas(t)) return t;
    const cols = CANVAS_COLS;
    t.canvas = { cols };
    const spec = (t.cardRows || []).map(n => Math.floor(Number(n))).filter(n => n >= 1);
    const rest = Math.floor(Number(t.cardsPerRow)) || 2;
    let y = 0, row = [], rowIdx = 0;
    const flush = () => {
      if (!row.length) return;
      const n = row.length, w = Math.floor(cols / n);
      row.forEach((s, i) => { s.pos = { x: i * w, y, w: i === n - 1 ? cols - i * w : w, h: 1 }; });
      y++; row = []; rowIdx++;
    };
    const want = () => spec[rowIdx] || rest;
    for (const s of t.sections || []) {
      if (s.type === 'grid' || (s.wide && !spec.length)) { flush(); s.pos = { x: 0, y, w: cols, h: 1 }; y++; rowIdx++; continue; }
      row.push(s);
      if (row.length >= Math.min(want(), 6)) flush();
    }
    flush();
    for (const s of t.sections || []) {
      if (s.type === 'list' || s.type === 'grid') continue;
      const perRow = Math.floor(Number(s.perRow));
      const cw = s.pos.w;
      s.cols = s.layout === 'stack' ? 1 : perRow >= 2 ? Math.min(perRow, 12) : cw >= 12 ? 6 : cw >= 6 ? 3 : 2;
      let x = 0, fy = 0;
      for (const f of s.fields || []) {
        const w = f.wide ? s.cols : (f.parts || []).length ? Math.min(2, s.cols) : 1;
        if (x + w > s.cols) { x = 0; fy++; }
        f.pos = { x, y: fy, w, h: 1 };
        x += w;
        if (f.wide) { x = 0; fy++; }
      }
      delete s.perRow; delete s.layout;
      for (const f of s.fields || []) delete f.wide;
    }
    for (const s of t.sections || []) delete s.wide;
    delete t.cardRows; delete t.cardsPerRow;
    return t;
  }
  function placeStyle(el, p, cols) {
    el.style.setProperty('--gc', `${p.x + 1} / span ${p.w}`);
    el.style.setProperty('--gr', `${p.y + 1} / span ${p.h}`);
    el.style.setProperty('--gcm', `auto / span ${Math.min(p.w, Math.min(cols, 2))}`);
    el.style.setProperty('--o', String(p.y * 100 + p.x));
  }

  // ---------------- 시트 그리기 (모든 룰 공통) ----------------
  const openFolds = new WeakSet(); // 펼쳐 둔 "숨긴 칸"은 다시 그려도 펼친 채로
  const openSections = new WeakSet(); // 펼쳐 둔 "숨긴 묶음"도 마찬가지
  const openToggles = new WeakSet(); // 펼쳐 둔 "접어 둔 카드"
  function renderSheet(template, sheet, { readonly, onChange, accent, datasets, exactLayout }) {
    sheet.values = sheet.values || {};
    sheet.lists = sheet.lists || {};
    if (!readonly) fillDefaults(template, sheet);
    if (datasets) E.syncLinkedRows(template, sheet, datasets); // 데이터 보관함이 바뀌었으면 연결된 칸을 새 값으로
    const updaters = [];
    const changed = () => { if (onChange) onChange(); recompute(); };

    function recompute() {
      const r = E.makeResolver(template, sheet);
      for (const u of updaters) u(r);
    }
    function computedCell(fn) {
      const out = h('output', { class: 'computed' });
      updaters.push(r => {
        try { out.textContent = E.format(fn(r)); out.classList.remove('err'); out.title = ''; }
        catch (e) { out.textContent = '⚠'; out.classList.add('err'); out.title = e.message; }
      });
      return out;
    }

    function input(type, value, set, rows) {
      if (type === 'checkbox') {
        return h('input', { type: 'checkbox', checked: !!value, disabled: readonly,
          onchange: e => { set(e.target.checked); changed(); } });
      }
      if (type === 'textarea') {
        return autoGrow(h('textarea', { value: value ?? '', readOnly: readonly, placeholder: readonly ? '—' : null,
          oninput: e => { set(e.target.value); changed(); } }));
      }
      if (type === 'number') {
        return h('input', { type: 'number', inputmode: 'decimal', value: value ?? '', readOnly: readonly, placeholder: readonly ? '—' : null,
          oninput: e => { set(e.target.value === '' ? '' : Number(e.target.value)); changed(); } });
      }
      // 글자 칸: 한 줄짜리지만 폭이 모자라면 줄을 바꿔서 다 보이게 (Enter로는 줄을 넣지 않음)
      return autoGrow(h('textarea', { class: 'line-text', value: value ?? '', readOnly: readonly, placeholder: readonly ? '—' : null,
        onkeydown: e => { if (e.key === 'Enter' && !e.isComposing) e.preventDefault(); },
        oninput: e => { const v = e.target.value.replace(/\n/g, ' '); if (v !== e.target.value) e.target.value = v; set(v); changed(); } }));
    }

    function renderField(f) {
      const set = v => { sheet.values[f.id] = v; };
      const val = sheet.values[f.id];
      let control;
      if (f.formula) control = computedCell(r => r.get(f.id));
      else if (f.type === 'select' && readonly) control = h('div', { class: 'read-text' }, val ? String(val) : '—');
      else if (f.type === 'select') {
        control = h('select', { disabled: readonly, onchange: e => { set(e.target.value); changed(); } },
          h('option', { value: '' }, '—'),
          (f.options || []).map(o => h('option', { value: o, selected: o === val }, o)));
      } else if (f.type === 'image') {
        const img = h('img', { class: 'portrait', alt: f.label || '', hidden: !isSafeUrl(val), src: isSafeUrl(val) ? val : null });
        const urlInput = readonly ? null : h('input', { type: 'url', value: val ?? '', placeholder: 'https://…',
          oninput: e => { set(e.target.value); img.hidden = !isSafeUrl(e.target.value); if (!img.hidden) img.src = e.target.value; changed(); } });
        control = h('div', { class: 'image-field' }, img, urlInput);
      } else {
        control = input(f.type || 'text', val, set);
        if (f.max != null && f.type === 'number') {
          const maxOut = computedCell(r => r.evaluate(f.max));
          control = h('div', { class: 'with-max' }, control, h('span', { class: 'slash' }, '/'), maxOut);
        }
      }
      // 한 칸에 숫자 여러 개: [첫 숫자] / [붙인 숫자] / …  (붙인 숫자도 직접 입력 또는 수식)
      const parts = (f.type === 'number' || f.formula !== undefined) ? (f.parts || []) : [];
      if (parts.length) {
        const sep = () => h('span', { class: 'part-sep' }, f.sep || '/');
        const cell = (x, ctrl, sub) => h('div', { class: 'part', 'data-ref': x.id }, ctrl, sub ? h('small', {}, sub) : null);
        control = h('div', { class: 'multi-num' },
          cell(f, control, f.mainLabel),
          parts.flatMap(p => [sep(), cell(p,
            p.formula !== undefined ? computedCell(r => r.get(p.id))
              : input('number', sheet.values[p.id], v => { sheet.values[p.id] = v; }),
            p.label)]));
      }
      const cls = ['field', f.formula ? 'is-computed' : '', f.type === 'checkbox' ? 'is-check' : '', f.wide ? 'wide' : '', parts.length ? 'has-parts' : ''].join(' ');
      return h('label', { class: cls, 'data-ref': f.id }, h('span', {}, f.label || f.id), control,
        f.hint ? h('small', { class: 'muted' }, f.hint) : null);
    }

    function renderList(s) {
      const rows = (sheet.lists[s.id] = sheet.lists[s.id] || []);
      const cols = s.columns || [];
      const tbody = h('tbody');
      const link = E.linkInfo(s, datasets);
      const listId = link ? 'dl-' + Math.random().toString(36).slice(2, 9) : null;
      function cell(c, row, i) {
        if (c.formula) return computedCell(r => r.cell(s.id, i, c.id));
        if (link && c === link.keyCol && !readonly) {
          // 이름을 다 적고 확정할 때(목록에서 고르기·Enter·다른 곳 누르기)만 데이터를 불러옴
          // → 한글 조합 중이거나 "힐"처럼 다른 이름의 앞부분일 때 미리 채워지지 않음
          const el = h('input', { type: 'text', value: row[c.id] ?? '', list: listId, placeholder: `${link.keyField.label || '이름'} 입력` });
          const wasLinked = !!link.find(row[c.id]);
          let before = row[c.id] ?? '';
          const commit = () => {
            const entry = link.find(el.value);
            if (String(el.value).trim() !== String(before).trim()) delete row._edited; // 다른 데이터를 불러오면 직접 고친 것은 초기화
            before = el.value;
            if (entry) E.fillRowFromEntry(s, row, entry, link.keyCol);
            else if (wasLinked) for (const col of cols) if (col.from && col !== link.keyCol) delete row[col.id]; // 다른 이름으로 바꾸면 예전 내용은 비움
            if (entry || wasLinked) { if (onChange) onChange(); draw(); }
          };
          el.addEventListener('input', e => {
            row[c.id] = el.value;
            if (onChange) onChange();
            recompute();
            if (e.inputType === 'insertReplacementText') commit(); // 목록에서 골랐을 때
          });
          el.addEventListener('change', commit);
          return el;
        }
        const entry = link && c.from && c !== link.keyCol ? link.find(row[link.keyCol.id]) : null;
        const edited = Array.isArray(row._edited) && row._edited.includes(c.id);
        if (entry && !readonly && link.editable(row[link.keyCol.id], c.from)) {
          // 데이터에서 불러왔지만 시트에서 고칠 수 있는 칸: 고치면 데이터 값 대신 이 값을 씀 (↺로 데이터 값으로 되돌리기)
          const reset = h('button', { type: 'button', class: 'cell-reset', title: '데이터 값으로 되돌리기', hidden: !edited,
            onclick: () => {
              row[c.id] = entry[c.from] ?? '';
              row._edited = (row._edited || []).filter(x => x !== c.id);
              if (!row._edited.length) delete row._edited;
              if (onChange) onChange(); draw();
            } }, '↺');
          const ed = plainCell(c, row);
          const mark = () => {
            const same = String(row[c.id] ?? '') === String(entry[c.from] ?? '');
            const list = new Set(Array.isArray(row._edited) ? row._edited : []);
            if (same) list.delete(c.id); else list.add(c.id);
            if (list.size) row._edited = [...list]; else delete row._edited;
            reset.hidden = same;
            wrap.classList.toggle('is-edited', !same);
          };
          ed.addEventListener('input', mark); ed.addEventListener('change', mark);
          const wrap = h('div', { class: 'linked-edit' + (edited ? ' is-edited' : ''), title: '데이터에서 불러온 값이에요. 이 시트에서만 고칠 수 있어요.' }, ed, reset);
          return wrap;
        }
        if (entry || (readonly && (c.type === 'textarea' || (c.type || 'text') === 'text'))) {
          if (entry && edited) return h('div', { class: 'linked-cell is-linked is-edited', title: '데이터 값을 이 시트에서 고친 값' }, String(row[c.id] ?? ''));
          return h('div', { class: 'linked-cell' + (link && c.from ? ' is-linked' : ''), title: link && c.from ? '데이터에서 불러온 값' : '' }, String(row[c.id] ?? ''));
        }
        return plainCell(c, row);
      }
      function plainCell(c, row) {
        if (c.type === 'select') {
          // 선택 목록 열: 고르기만 하면 됨 (선택지에 없는 예전 값도 그대로 보여 줌)
          const cur = String(row[c.id] ?? ''), opts = c.options || [];
          if (readonly) return h('div', { class: 'linked-cell' }, cur || '—');
          return h('select', { class: 'cell-select', onchange: e => { row[c.id] = e.target.value; changed(); } },
            h('option', { value: '' }, '—'), opts.map(o => h('option', { value: o, selected: o === cur }, o)),
            cur && !opts.includes(cur) ? h('option', { value: cur, selected: true }, cur) : null);
        }
        if ((c.type || 'text') === 'text') {
          // 글자 열: 길어지면 칸 안에서 줄을 바꿔서 다 보이게 (Enter로는 줄을 넣지 않음)
          return autoGrow(h('textarea', { class: 'cell-text', value: row[c.id] ?? '',
            onkeydown: e => { if (e.key === 'Enter' && !e.isComposing) e.preventDefault(); },
            oninput: e => { const v = e.target.value.replace(/\n/g, ' '); if (v !== e.target.value) e.target.value = v; row[c.id] = v; changed(); } }));
        }
        return input(c.type || 'text', row[c.id], v => { row[c.id] = v; }, 2);
      }
      function draw() {
        updaters.length = baseUpdaters; // 표를 다시 그릴 때 이전 줄의 계산기를 정리
        listUpdaters.forEach(fn => fn());
        recompute();
      }
      function drawRows() {
        tbody.replaceChildren(...rows.map((row, i) => h('tr', {},
          cols.map(c => h('td', { 'data-label': c.label || c.id, 'data-ref': c.id, 'data-list': s.id },
            cell(c, row, i))),
          readonly ? null : h('td', { class: 'row-del' },
            h('button', { type: 'button', class: 'ghost', title: '이 줄 삭제',
              onclick: () => { rows.splice(i, 1); if (onChange) onChange(); draw(); } }, '×')))));
        if (!rows.length && !(link && !readonly)) tbody.append(h('tr', {}, h('td', { colspan: cols.length + 1, class: 'muted center' }, '비어 있어요')));
      }
      listUpdaters.push(drawRows);
      drawRows();
      // 데이터와 연결된 표: 맨 아래 칸에 이름을 적으면 새 줄이 생기면서 나머지 열이 채워짐
      let quick = null;
      if (link && !readonly) {
        const qi = h('input', { type: 'text', list: listId, class: 'quick-add-input',
          placeholder: `${link.keyField.label || '이름'}을 적거나 골라서 새 줄 추가` });
        const add = () => {
          const v = qi.value.trim();
          if (!v) return;
          const row = { [link.keyCol.id]: v };
          const entry = link.find(v);
          if (entry) E.fillRowFromEntry(s, row, entry, link.keyCol);
          rows.push(row);
          qi.value = '';
          if (onChange) onChange();
          draw();
          qi.focus();
        };
        qi.addEventListener('input', e => { if (e.inputType === 'insertReplacementText') add(); });
        qi.addEventListener('change', add);
        qi.addEventListener('keydown', e => { if (e.key === 'Enter' && !e.isComposing) { e.preventDefault(); add(); } });
        quick = h('tfoot', {}, h('tr', { class: 'quick-add' }, h('td', { colspan: cols.length + 1 }, qi)));
      }
      return h('div', { class: 'list-wrap' },
        link && !readonly ? h('datalist', { id: listId }, [...link.map.keys()].map(k => h('option', { value: k }))) : null,
        h('table', { class: 'list' },
          h('thead', {}, h('tr', {}, cols.map(c => h('th', {}, c.label || c.id)), readonly ? null : h('th'))),
          tbody, quick),
        readonly ? null : h('button', { type: 'button', class: 'add-row',
          onclick: () => { rows.push({}); if (onChange) onChange(); draw(); } }, '+ 줄 추가'));
    }

    // 특기표(격자): 습득 체크 → 가까운 습득 칸에서부터 거리만큼 달성치가 올라감
    function renderGrid(s) {
      sheet.grids = sheet.grids || {};
      const st = sheet.grids[s.id] = sheet.grids[s.id] || { acquired: [], blackout: null, disabled: [] };
      st.acquired = st.acquired || []; st.disabled = st.disabled || [];
      const cols = s.columns || [], rows = s.rows || [];
      const box = h('div', { class: 'skill-grid' });
      const names = s.columnNames || [];
      const colName = c => [cols[c] || String(c + 1), names[c]].filter(Boolean).join(' ');
      const edited = () => { if (onChange) onChange(); recompute(); draw(); };
      function draw() {
        const vals = E.computeGrid(s, E.gridState(s, sheet));
        const acq = new Set(st.acquired);
        const dis = new Set(st.disabled.map(Number));
        const black = st.blackout === null || st.blackout === undefined ? null : Number(st.blackout);
        // 열 c 왼쪽의 빈 줄(c-1과 c 사이)이 검게 칠해졌는지
        const gapBlack = c => black !== null && c > 0 && (black === c || black === c - 1);
        // 읽기·공유 화면: 지금 상태를 한 줄로 / 편집: 열 머리글 양쪽의 작은 버튼으로 바꿈
        const sideOf = c => (cols.length === 1 ? '' : c === 0 ? '오른쪽' : c === cols.length - 1 ? '왼쪽' : '양쪽');
        const controls = readonly && (black !== null || dis.size)
          ? h('p', { class: 'muted small grid-note' },
            [black !== null ? `무효화한 빈 줄: ${colName(black)} ${sideOf(black)}`.trim() : null,
              dis.size ? `습득 안 한 것으로 취급: ${[...dis].sort((a, b) => a - b).map(colName).join(', ')}` : null].filter(Boolean).join(' · '))
          : null;
        const colHead = c => {
          const label = h('span', { class: 'col-label' }, h('b', {}, cols[c] || String(c + 1)),
            names[c] ? h('span', { class: 'col-name' }, names[c]) : null);
          if (readonly) return label;
          return h('div', { class: 'col-head' },
            h('button', { type: 'button', class: 'gtog gap-tog' + (black === c ? ' on' : ''),
              title: black === c ? '빈 줄 무효화 풀기' : `${colName(c)} ${sideOf(c)} 빈 줄 무효화 (검게 칠하기)`,
              onclick: () => { st.blackout = black === c ? null : c; edited(); } }, '◼'),
            label,
            h('button', { type: 'button', class: 'gtog off-tog' + (dis.has(c) ? ' on' : ''),
              title: dis.has(c) ? '미습득 취급 풀기' : `${colName(c)}의 습득을 모두 없는 것으로 계산하기`,
              onclick: () => { st.disabled = dis.has(c) ? st.disabled.filter(x => Number(x) !== c) : [...st.disabled, c]; edited(); } }, '⊘'));
        };
        const table = h('table', { class: 'grid-table' },
          h('thead', {}, h('tr', {}, h('th', { class: 'corner' }),
            cols.map((label, c) => [h('th', { class: 'gap' + (gapBlack(c) ? ' black' : '') }),
              h('th', { class: 'col-h' + (dis.has(c) ? ' col-off' : '') }, colHead(c))]))),
          h('tbody', {}, rows.map((rl, r) => h('tr', {}, h('th', { class: 'row-h' }, rl),
            cols.map((_, c) => {
              const key = `${r}-${c}`, name = ((s.cells || [])[r] || [])[c] || '';
              const on = acq.has(key);
              const td = h('td', { class: 'cell' + (on ? ' learned' : '') + (dis.has(c) ? ' col-off' : ''), },
                h('div', { class: 'cell-in' },
                  h('input', { type: 'checkbox', checked: on, disabled: readonly, 'aria-label': name || `${rl} ${colName(c)}`,
                    onclick: e => e.stopPropagation(),
                    onchange: e => { st.acquired = e.target.checked ? [...st.acquired, key] : st.acquired.filter(k => k !== key); edited(); } }),
                  h('span', { class: 'name', title: name }, name),
                  h('span', { class: 'val' }, String(vals[r][c]))));
              if (!readonly) td.addEventListener('click', () => td.querySelector('input').click());
              return [h('td', { class: 'gap' + (gapBlack(c) ? ' black' : '') }), td];
            })))));
        box.replaceChildren(...[controls, h('div', { class: 'grid-wrap' }, table)].filter(Boolean));
      }
      draw();
      return box;
    }

    // 표는 다시 그려질 수 있으므로, 고정 칸 계산기와 표 계산기를 나눠서 관리
    const listUpdaters = [];
    const fieldSections = [], listSections = [], gridSections = [];
    const root = h('div', { class: 'sheet' });
    if (accent || template.accent) root.style.setProperty('--accent', accent || template.accent);
    // 한 줄에 놓을 칸 묶음 수 (룰 틀마다). 화면이 좁으면 2개, 휴대폰은 1개까지 줄어듦
    const canvas = isCanvas(template);
    const cardsPerRow = canvas ? 0 : Math.floor(Number(template.cardsPerRow));
    const cardRows = canvas ? [] : (template.cardRows || []).map(n => Math.floor(Number(n))).filter(n => n >= 1);
    if (canvas) { root.classList.add('canvas'); root.style.setProperty('--cols', CANVAS_COLS); }
    if (cardsPerRow >= 1 && !cardRows.length) {
      root.classList.add('fixed-cards');
      root.style.setProperty('--cards', cardsPerRow);
      root.style.setProperty('--cards-md', Math.min(cardsPerRow, 2));
    }

    // 숨긴 칸 보기: 카드 안 격자의 가로 칸 수를 따라 차례대로 채움 (칸 너비는 정한 너비 그대로)
    function foldedGrid(s, folded) {
      const cols = Math.max(1, Math.min(12, Math.floor(s.cols) || 4));
      const grid = h('div', { class: 'fields canvas folded-grid' });
      grid.style.setProperty('--cols', cols);
      grid.style.setProperty('--cols-m', Math.min(cols, 2));
      for (const f of folded) {
        const el = renderField(f);
        const w = Math.min(posOf(f, cols).w, cols);
        el.style.setProperty('--gc', `auto / span ${w}`);
        el.style.setProperty('--gr', 'auto');
        el.style.setProperty('--gcm', `auto / span ${Math.min(w, Math.min(cols, 2))}`);
        grid.append(el);
      }
      return grid;
    }
    const cardBySection = (template.sections || []).map((s, si) => {
      // 숨긴 묶음: 읽기·공유 화면에서는 아예 빼고, 편집할 때는 접힌 카드로 (눌러서 펼침)
      if (s.folded && readonly) return null;
      const card = h('section', { class: 'card' + (!canvas && (s.wide || s.type === 'grid') ? ' wide' : '') + (s.folded ? ' folded-card' : '') + (s.type === 'grid' ? ' grid-card' : ''), 'data-sec': si });
      if (canvas) placeStyle(card, posOf(s, CANVAS_COLS), CANVAS_COLS);
      let target = card;
      if (s.folded) {
        target = h('details', { class: 'folded-section', open: openSections.has(s),
          ontoggle: e => { if (e.target.open) openSections.add(s); else openSections.delete(s); } },
          h('summary', {}, h('h2', {}, s.title || ''), h('span', { class: 'muted small' }, '숨긴 묶음')));
        card.append(target);
      } else if (s.toggle) {
        // 접어 둔 카드: 읽기·공유 화면에서도 제목을 눌러 펼쳐 볼 수 있음 (템플릿 편집기에서는 펼친 채로)
        target = h('details', { class: 'toggle-section', open: exactLayout || openToggles.has(s),
          ontoggle: e => { if (exactLayout) return; if (e.target.open) openToggles.add(s); else openToggles.delete(s); } },
          h('summary', {}, h('h2', {}, s.title || ''), h('span', { class: 'toggle-hint muted small' }, '펼치기')));
        card.classList.add('toggle-card');
        card.append(target);
      } else card.append(h('h2', {}, s.title || ''));
      if (s.type === 'grid') gridSections.push([target, s]);
      else if (s.type === 'list') listSections.push([target, s]);
      else fieldSections.push([target, s]);
      return card;
    });
    const cards = cardBySection.filter(Boolean);
    for (const [card, s] of fieldSections) {
      // 칸 배치: 한 줄에 하나씩(stack) / 정한 칸 수(perRow) / 자동
      const perRow = Math.floor(Number(s.perRow));
      const fixed = s.layout !== 'stack' && perRow >= 2;
      const layout = 'fields ' + (s.layout === 'stack' ? 'stack' : 'grid') + (fixed ? ' fixed' : '');
      const box = kids => {
        const el = h('div', { class: layout }, kids);
        if (fixed) { el.style.setProperty('--per-row', perRow); el.style.setProperty('--per-row-sm', Math.min(perRow, 2)); }
        return el;
      };
      const shown = (s.fields || []).filter(f => !f.folded), folded = (s.fields || []).filter(f => f.folded);
      if (canvas) {
        // 격자 배치: 칸마다 정한 자리에 (빈 묶음도 편집할 때는 칸을 끌어다 놓을 수 있게 빈 격자를 둠)
        const cols = Math.max(1, Math.min(12, Math.floor(s.cols) || 4));
        if (shown.length || !readonly) {
          const grid = h('div', { class: 'fields canvas' + (shown.length ? '' : ' empty') });
          grid.style.setProperty('--cols', cols);
          grid.style.setProperty('--cols-m', Math.min(cols, 2));
          for (const f of shown) { const el = renderField(f); placeStyle(el, posOf(f, cols), cols); grid.append(el); }
          card.append(grid);
        }
      } else if (shown.length) card.append(box(shown.map(renderField)));
      // 접힌 칸: 편집할 때만 "숨긴 칸 보기"로 펼쳐서 입력, 읽기·공유 화면에서는 안 보임
      if (folded.length && !readonly) {
        const det = h('details', { class: 'folded-fields', open: openFolds.has(s),
          ontoggle: e => { if (e.target.open) openFolds.add(s); else openFolds.delete(s); } },
          h('summary', {}, `숨긴 칸 ${folded.length}개 보기`), canvas ? foldedGrid(s, folded) : box(folded.map(renderField)));
        card.append(det);
      }
      if (readonly && !shown.length) card.hidden = true;
    }
    const baseUpdaters = updaters.length;
    for (const [card, s] of listSections) card.append(renderList(s));
    for (const [card, s] of gridSections) card.append(renderGrid(s));

    if (cardRows.length) {
      // 줄마다 정한 개수만큼: 룰 틀의 칸 묶음 순서대로 1번째 줄에 n개, 2번째 줄에 m개 … 나머지는 "나머지 줄" 설정대로
      root.classList.add('row-mode');
      const rowOf = n => {
        const row = h('div', { class: 'sheet-row' });
        row.style.setProperty('--cards', n);
        row.style.setProperty('--cards-md', Math.min(n, 2));
        return row;
      };
      let k = 0;
      const take = n => { const part = cardBySection.slice(k, k + n).filter(Boolean); k += n; return part; };
      for (const n of cardRows) {
        if (k >= cardBySection.length) break;
        const part = take(n);
        // 숨긴 묶음이 빠져서 모자라면 남은 카드가 그 줄을 나눠 씀
        if (part.length) { const row = rowOf(Math.min(n, part.length)); row.append(...part); root.append(row); }
      }
      while (k < cardBySection.length) {
        if (cardsPerRow >= 1) {
          const part = take(cardsPerRow);
          if (part.length) { const row = rowOf(Math.min(cardsPerRow, part.length)); row.append(...part); root.append(row); }
        } else {
          const rest = cardBySection.slice(k).filter(Boolean); k = cardBySection.length;
          if (rest.length) root.append(h('div', { class: 'sheet-row auto' }, rest));
        }
      }
    } else {
      if (canvas && !exactLayout) {
        // 시트 화면: 숨긴 카드·빈 카드가 빠진 자리(편집 중에는 접힌 카드의 남는 높이)를 남은 카드가 위로 올라와 채움
        const shownCards = cardBySection.map((card, i) => [card, (template.sections || [])[i]]).filter(([card]) => card && !card.hidden);
        const items = shownCards.map(([card, s]) => { const q = posOf(s, CANVAS_COLS); return { card, ...q, h: s.folded || s.toggle ? 1 : q.h }; });
        settle(items, CANVAS_COLS);
        for (const it of items) placeStyle(it.card, it, CANVAS_COLS);
      }
      root.append(...cards);
    }
    recompute();
    return root;
  }

  // ---------------- 목록 화면 공통: 카드 크기 · 정렬 · 꾹 눌러서 직접 배치 ----------------
  // 카드 크기·정렬 방식은 이 브라우저에, 직접 배치한 순서는 데이터 저장소(meta/order.json)에 저장 → 다른 기기에서도 같은 순서
  const LIST_KEY = 'trpg-list-view';
  const CARD_SIZES = [['s', '작게'], ['m', '보통'], ['l', '크게']];
  function listPrefs(kind, defSort) {
    let all = {};
    try { all = JSON.parse(localStorage.getItem(LIST_KEY)) || {}; } catch {}
    return { size: 'm', sort: defSort || 'custom', ...(all[kind] || {}) };
  }
  function saveListPrefs(kind, p) {
    try {
      const all = JSON.parse(localStorage.getItem(LIST_KEY) || '{}') || {};
      all[kind] = p; localStorage.setItem(LIST_KEY, JSON.stringify(all));
    } catch {}
  }
  const ORDER_PATH = 'meta/order.json';
  let orderState = { data: {}, sha: null };
  async function loadOrder() {
    try { orderState = await readJson(ORDER_PATH); if (!orderState.data || typeof orderState.data !== 'object') orderState.data = {}; }
    catch (e) { if (e.status !== 404) console.warn(e); orderState = { data: {}, sha: null }; }
    return orderState.data;
  }
  async function saveOrder(kind, keys) {
    for (let i = 0; i < 2; i++) {
      try {
        const data = { ...orderState.data, [kind]: keys };
        orderState = { data, sha: await writeJson(ORDER_PATH, data, orderState.sha, `순서 바꾸기: ${kind}`) };
        return;
      } catch (e) {
        if (i === 0 && (e.status === 409 || e.status === 422)) { // 다른 기기에서 먼저 고쳤으면 새로 읽고 한 번 더
          try { orderState = await readJson(ORDER_PATH); } catch { orderState = { data: {}, sha: null }; }
          continue;
        }
        throw e;
      }
    }
  }
  // 저장된 순서대로 정렬. 순서에 없는 새 항목은 맨 앞(newFirst) 또는 맨 뒤
  function applyOrder(keys, saved, newFirst) {
    const pos = new Map((saved || []).map((k, i) => [k, i]));
    const known = keys.filter(k => pos.has(k)).sort((a, b) => pos.get(a) - pos.get(b));
    const fresh = keys.filter(k => !pos.has(k));
    return newFirst ? [...fresh, ...known] : [...known, ...fresh];
  }
  // 필터로 일부만 보일 때: 보이는 항목들의 자리만 새 순서로 바꿈
  function mergeVisible(full, visible) {
    const set = new Set(visible); let i = 0;
    return full.map(k => (set.has(k) ? visible[i++] : k));
  }
  const byName = (a, b) => String(a || '').localeCompare(String(b || ''), 'ko', { numeric: true });

  // 정렬 · 카드 크기 고르는 줄 (필터 막대 안에 들어감)
  function listTools(prefs, sorts, onChange) {
    const chip = (on, label, click) => h('button', { type: 'button', class: 'filter-chip' + (on ? ' on' : ''), onclick: click }, label);
    return [
      h('div', { class: 'filter-group' }, h('span', { class: 'filter-label' }, '정렬'),
        sorts.map(([id, label]) => chip(prefs.sort === id, label, () => { prefs.sort = id; onChange(); }))),
      h('div', { class: 'filter-group' }, h('span', { class: 'filter-label' }, '카드 크기'),
        CARD_SIZES.map(([id, label]) => chip(prefs.size === id, label, () => { prefs.size = id; onChange(); }))),
    ];
  }

  // 카드를 꾹 눌러서(0.4초) 끌어다 놓기. 카드마다 data-key가 있어야 함. onDrop(새 순서의 키 배열)
  function sortableCards(container, onDrop) {
    container.classList.add('sortable-cards');
    for (const card of container.children) if (card.dataset && card.dataset.key) armCard(card);
    function armCard(card) {
      card.addEventListener('dragstart', e => e.preventDefault()); // 링크·그림 기본 끌기 막기
      card.addEventListener('contextmenu', e => { if (card._pressing || document.body.classList.contains('dragging')) e.preventDefault(); });
      card.addEventListener('pointerdown', e => {
        if (e.button !== 0 || e.target.closest('button, input, select, textarea')) return;
        const sx = e.clientX, sy = e.clientY;
        card._pressing = true;
        let timer = setTimeout(() => { timer = null; done(); beginDrag(card, sx, sy); }, 400);
        const move = ev => { if (Math.hypot(ev.clientX - sx, ev.clientY - sy) > 8) done(); };
        function done() {
          if (timer) clearTimeout(timer);
          timer = null; card._pressing = false;
          window.removeEventListener('pointermove', move);
          window.removeEventListener('pointerup', done);
          window.removeEventListener('pointercancel', done);
        }
        window.addEventListener('pointermove', move);
        window.addEventListener('pointerup', done);
        window.addEventListener('pointercancel', done);
      });
    }
    function beginDrag(card, x0, y0) {
      const before = [...container.children].map(c => c.dataset.key).filter(Boolean);
      const rect = card.getBoundingClientRect();
      const ghost = card.cloneNode(true);
      ghost.classList.add('drag-ghost', 'card-ghost');
      Object.assign(ghost.style, { width: rect.width + 'px', height: rect.height + 'px', left: rect.left + 'px', top: rect.top + 'px' });
      document.body.append(ghost);
      card.classList.add('drag-source');
      document.body.classList.add('dragging');
      if (navigator.vibrate) try { navigator.vibrate(15); } catch {}
      const dx = x0 - rect.left, dy = y0 - rect.top;
      let lastX = x0, lastY = y0, raf = 0;
      const blockTouch = ev => ev.preventDefault(); // 끄는 동안 화면이 같이 스크롤되지 않게
      document.addEventListener('touchmove', blockTouch, { passive: false });
      function place() {
        ghost.style.left = (lastX - dx) + 'px';
        ghost.style.top = (lastY - dy) + 'px';
        const el = document.elementFromPoint(lastX, lastY);
        const over = el && el.closest('[data-key]');
        if (!over || over === card || over.parentElement !== container) return;
        const r = over.getBoundingClientRect();
        const oneCol = getComputedStyle(container).gridTemplateColumns.split(' ').filter(Boolean).length <= 1;
        const first = oneCol ? lastY < r.top + r.height / 2 : lastX < r.left + r.width / 2;
        const ref = first ? over : over.nextSibling;
        if (ref !== card && card.nextSibling !== ref) container.insertBefore(card, ref);
        else if (ref === null && container.lastChild !== card) container.append(card);
      }
      function edgeScroll() {
        const edge = 70, wh = window.innerHeight;
        const speed = lastY < edge ? -Math.ceil((edge - lastY) / 4) : lastY > wh - edge ? Math.ceil((lastY - (wh - edge)) / 4) : 0;
        if (speed) { window.scrollBy(0, speed); place(); }
        raf = requestAnimationFrame(edgeScroll);
      }
      raf = requestAnimationFrame(edgeScroll);
      const onMove = ev => { lastX = ev.clientX; lastY = ev.clientY; place(); };
      function end() {
        cancelAnimationFrame(raf);
        window.removeEventListener('pointermove', onMove);
        window.removeEventListener('pointerup', end);
        window.removeEventListener('pointercancel', end);
        document.removeEventListener('touchmove', blockTouch);
        ghost.remove();
        card.classList.remove('drag-source');
        document.body.classList.remove('dragging');
        // 끌고 난 뒤 바로 따라오는 클릭(카드 열기)은 무시
        const eat = ev => { ev.preventDefault(); ev.stopPropagation(); };
        window.addEventListener('click', eat, true);
        setTimeout(() => window.removeEventListener('click', eat, true), 350);
        const after = [...container.children].map(c => c.dataset.key).filter(Boolean);
        if (after.join('\u0000') !== before.join('\u0000')) onDrop(after);
      }
      window.addEventListener('pointermove', onMove);
      window.addEventListener('pointerup', end);
      window.addEventListener('pointercancel', end);
    }
    return container;
  }

  // ---------------- 끌어서 순서 바꾸기 (룰 틀 편집기·데이터 보관함 공통) ----------------
  // dropList(요소, 배열, 묶음키): 이 요소 안의 .drag-item들이 배열의 순서를 나타냄. 묶음키가 같은 목록끼리만 옮길 수 있음
  // draggable(요소, 객체, { onMoved, canMove, ghost }): ⠿ 손잡이 또는 빈 곳을 꾹 눌러 끌기
  const dragHandle = () => h('span', { class: 'drag-handle', title: '꾹 누른 채로 끌어서 옮기기' }, '⠿');
  function dropList(el, arr, group) {
    el.classList.add('drop-list');
    el._arr = arr; el._group = group;
    return el;
  }
  function draggable(item, obj, opts) {
    opts = opts || {};
    item.classList.add('drag-item');
    item._obj = obj;
    item.addEventListener('pointerdown', e => {
      if (e.button !== 0 || e._dragClaimed) return;
      const onHandle = e.target.closest('.drag-handle');
      if (onHandle && onHandle.closest('.drag-item') !== item) return; // 안쪽 항목의 손잡이
      if (!onHandle && (e.pointerType === 'touch' || e.target.closest('input, textarea, select, button, a, summary, label, .formula-wrap, .formula-helper, .tpl-chips'))) return;
      e._dragClaimed = true;
      const sx = e.clientX, sy = e.clientY;
      let timer = setTimeout(() => { timer = null; stop(); startDrag(item, obj, opts, sx, sy); }, onHandle ? 150 : 350);
      const move = ev => { if (Math.hypot(ev.clientX - sx, ev.clientY - sy) > 6) stop(); };
      function stop() {
        if (timer) clearTimeout(timer);
        timer = null;
        window.removeEventListener('pointermove', move);
        window.removeEventListener('pointerup', stop);
        window.removeEventListener('pointercancel', stop);
      }
      window.addEventListener('pointermove', move);
      window.addEventListener('pointerup', stop);
      window.addEventListener('pointercancel', stop);
    });
    return item;
  }
  function startDrag(item, obj, opts, x0, y0) {
    const src = item.parentElement && item.parentElement.closest('.drop-list');
    if (!src) return;
    const group = src._group;
    const shape = (opts.ghost && opts.ghost()) || item;
    const rect = item.getBoundingClientRect();
    const ghost = shape.cloneNode(true);
    ghost.classList.add('drag-ghost');
    const gw = shape === item ? rect.width : Math.min(rect.width, 420);
    Object.assign(ghost.style, { width: gw + 'px', left: rect.left + 'px', top: rect.top + 'px' });
    document.body.append(ghost);
    item.classList.add('drag-source');
    document.body.classList.add('dragging');
    const marker = h('div', { class: 'drop-marker' });
    let target = null, lastX = x0, lastY = y0, raf = 0;
    const dx = Math.min(x0 - rect.left, gw - 10), dy = Math.min(y0 - rect.top, 30);
    function findList(el) {
      let l = el && el.closest('.drop-list');
      while (l && l._group !== group) l = l.parentElement && l.parentElement.closest('.drop-list');
      return l;
    }
    function update() {
      ghost.style.left = (lastX - dx) + 'px';
      ghost.style.top = (lastY - dy) + 'px';
      const list = findList(document.elementFromPoint(lastX, lastY));
      if (!list) { marker.remove(); target = null; return; }
      const items = [...list.children].filter(c => c.classList.contains('drag-item') && c !== item);
      let idx = items.findIndex(c => { const r = c.getBoundingClientRect(); return lastY < r.top + r.height / 2; });
      if (idx < 0) idx = items.length;
      if (idx < items.length) list.insertBefore(marker, items[idx]); else list.append(marker);
      target = { list, idx };
    }
    // 화면 위·아래 끝에 대고 있으면 계속 스크롤
    function edgeScroll() {
      const edge = 70, wh = window.innerHeight;
      const speed = lastY < edge ? -Math.ceil((edge - lastY) / 4) : lastY > wh - edge ? Math.ceil((lastY - (wh - edge)) / 4) : 0;
      if (speed) { window.scrollBy(0, speed); update(); }
      raf = requestAnimationFrame(edgeScroll);
    }
    raf = requestAnimationFrame(edgeScroll);
    const onMove = ev => { ev.preventDefault(); lastX = ev.clientX; lastY = ev.clientY; update(); };
    const onScroll = () => update();
    function onUp() {
      cancelAnimationFrame(raf);
      window.removeEventListener('scroll', onScroll);
      window.removeEventListener('pointermove', onMove);
      window.removeEventListener('pointerup', onUp);
      window.removeEventListener('pointercancel', onUp);
      ghost.remove(); marker.remove();
      item.classList.remove('drag-source');
      document.body.classList.remove('dragging');
      if (!target) return;
      const from = src._arr, to = target.list._arr;
      const oldIdx = from.indexOf(obj);
      if (oldIdx < 0) return;
      if (from === to && oldIdx === target.idx) return;
      if (opts.canMove && !opts.canMove({ from, to, oldIdx, newIdx: target.idx })) return;
      from.splice(oldIdx, 1);
      to.splice(Math.min(target.idx, to.length), 0, obj);
      if (opts.onMoved) opts.onMoved({ from, to, oldIdx, newIdx: target.idx });
    }
    window.addEventListener('pointermove', onMove, { passive: false });
    window.addEventListener('pointerup', onUp);
    window.addEventListener('pointercancel', onUp);
    window.addEventListener('scroll', onScroll, { passive: true });
    update();
  }

  // 다른 파일(editor.js)이 쓸 수 있도록 내부 도구를 내보냄
  window.SheetApp = {
    canvas: { COLS: CANVAS_COLS, isCanvas, posOf, overlaps, settle, freeSpot, toCanvas, placeStyle },
    h, toast, show, loading, failed, go, setDirty, state, routes,
    readJson, writeJson, deleteFile, listDir, loadTemplates,
    resetTemplates: () => { templateCache = null; },
    renderSheet, blankSheet, charactersUsing, loadDatasets, refreshShares, autoGrow,
    dragHandle, dropList, draggable,
    listPrefs, saveListPrefs, loadOrder, saveOrder, applyOrder, mergeVisible, byName, listTools, sortableCards,
    resetDatasets: () => { datasetCache = null; },
    fileKey: rule => String(rule || '').trim().replace(/[\\/:*?"<>|#%&\s]+/g, '-').replace(/^-+|-+$/g, '') || 'data',
  };
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', route);
  else route();
})();
