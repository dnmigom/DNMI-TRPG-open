/*
 * editor.js — 룰 틀(템플릿) 편집기
 * - 사이트 안에서 룰 틀을 만들고, 고치고, 복제하고, 삭제합니다.
 * - 수식은 엑셀처럼: 버튼으로 함수·칸을 넣거나, 수식 입력 중 미리보기의 칸을 클릭해서 넣습니다.
 */
(function () {
  'use strict';
  const A = window.SheetApp, E = window.SheetEngine, h = A.h;
  const CV = A.canvas;

  const KINDS = [['text', '글자'], ['number', '숫자'], ['textarea', '긴 글'], ['checkbox', '체크'],
    ['select', '선택 목록'], ['image', '그림 주소'], ['formula', '자동 계산']];
  const COL_KINDS = [['text', '글자'], ['number', '숫자'], ['textarea', '긴 글'], ['checkbox', '체크'], ['select', '선택 목록'], ['formula', '자동 계산']];
  const FUNCS = [
    ['내림', '소수점 아래를 버려요. 내림(7 / 2) → 3'],
    ['올림', '소수점 아래를 올려요. 올림(7 / 2) → 4'],
    ['반올림', '반올림해요. 반올림(7 / 2) → 4'],
    ['최대', '가장 큰 값. 최대(근력, 민첩)'],
    ['최소', '가장 작은 값. 최소(근력, 민첩)'],
    ['합계', '모두 더해요. 합계(근력, 민첩, 정신)'],
    ['범위', '범위(값, 최소, 최대) 사이로 묶어요'],
    ['만약', '만약(조건, 맞을 때, 틀릴 때). 예: 만약(부상, -2, 0)'],
    ['들어있음', '들어있음(값, "글자") → 그 글자가 들어 있으면 1. 예: 만약(들어있음(이름, "불"), "*", "&")'],
    ['잇기', '잇기(a, b, …) → 글자를 이어 붙여요. 예: 잇기("*", 이름)'],
    ['목록잇기', '목록잇기("표", "열", ", ") → 표의 한 열을 줄마다 이어서 나열'],
  ];
  const OPS = ['+', '-', '*', '/', '(', ')', '>', '>=', '<', '<=', '='];
  const OP_LABEL = { '*': '×', '/': '÷', '>=': '≥', '<=': '≤' };
  const OP_TIP = { '*': '곱하기 (수식에는 * 로 들어가요)', '/': '나누기 (수식에는 / 로 들어가요)', '>': '크면 1, 아니면 0', '=': '같으면 1, 아니면 0' };

  const clone = o => JSON.parse(JSON.stringify(o));
  const esc = s => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const validId = id => /^[\p{L}_][\p{L}\p{N}_]*$/u.test(id || '');
  function idFrom(label, taken) {
    let base = String(label || '').replace(/[^\p{L}\p{N}_]/gu, '') || '칸';
    if (/^\p{N}/u.test(base)) base = '_' + base;
    let id = base, n = 2;
    while (taken.has(id)) id = base + n++;
    return id;
  }
  const fileKey = rule => String(rule || '').trim().replace(/[\\/:*?"<>|#%&\s]+/g, '-').replace(/^-+|-+$/g, '') || 'rule';
  const kindOf = f => (f.formula !== undefined ? 'formula' : (f.type || 'text'));

  // 수식 안의 이름 바꾸기 (칸 이름이 바뀌면 그 칸을 쓰는 수식도 같이 바뀜)
  const replaceIdent = (src, oldId, newId) => String(src).replace(
    new RegExp(`(?<![\\p{L}\\p{N}_"'])${esc(oldId)}(?![\\p{L}\\p{N}_"'(])`, 'gu'), newId);

  A.routes.push({ match: hash => hash === '/templates', show: showTemplateList });
  A.routes.push({ match: hash => hash.startsWith('/template/'), show: hash => showEditor(decodeURIComponent(hash.slice(10))) });


  // ================= 룰 틀 목록 =================
  async function showTemplateList() {
    A.loading();
    try {
      const [templates, order] = await Promise.all([A.loadTemplates(true), A.loadOrder()]);
      const prefs = A.listPrefs('templates', 'custom');
      const SORTS = [['custom', '직접 배치'], ['name', '이름 가나다순']];
      if (!SORTS.some(x => x[0] === prefs.sort)) prefs.sort = 'custom';
      const customKeys = () => A.applyOrder(Object.keys(templates), order.templates);
      const sortedKeys = () => prefs.sort === 'name'
        ? [...customKeys()].sort((a, b) => A.byName(templates[a].rule || a, templates[b].rule || b)) : customKeys();
      const bar = h('div', { class: 'filter-bar' });
      const list = h('div', { class: 'sheet-list' });
      async function dropped(keys) {
        order.templates = keys;
        if (prefs.sort !== 'custom') { prefs.sort = 'custom'; A.saveListPrefs('templates', prefs); A.toast('정렬을 "직접 배치"로 바꿨어요.'); }
        draw();
        try { await A.saveOrder('templates', keys); } catch (e) { A.toast('순서를 저장하지 못했어요: ' + e.message, true); }
      }
      function card(k) {
        const t = templates[k];
        const c = h('div', { class: 'sheet-card', 'data-key': k },
          h('strong', {}, t.rule || k),
          h('span', { class: 'muted' }, t.error || t.description || ''),
          h('div', { class: 'row' },
            h('a', { class: 'button-link', href: `#/template/${encodeURIComponent(k)}` }, '편집'),
            h('a', { class: 'button-link', href: `#/template/new?from=${encodeURIComponent(k)}` }, '복제')));
        c.style.setProperty('--accent', t.accent || '');
        return c;
      }
      function draw() {
        bar.replaceChildren(...A.listTools(prefs, SORTS, () => { A.saveListPrefs('templates', prefs); draw(); }));
        const keys = sortedKeys();
        list.className = `sheet-list size-${prefs.size}`;
        list.replaceChildren(...(keys.length ? keys.map(card) : [h('p', { class: 'muted' }, '아직 템플릿이 없어요.')]));
        A.sortableCards(list, dropped);
      }
      draw();
      A.show(
        h('div', { class: 'toolbar' }, h('h1', {}, '시트 템플릿'),
          h('button', { class: 'primary', onclick: () => A.go('/template/new') }, '+ 새 시트 템플릿')),
        h('p', { class: 'muted' }, '룰마다 시트에 들어갈 칸과 계산식을 정해 두는 곳이에요. 비슷한 룰은 "복제"해서 고치면 편해요. 카드를 꾹 누르면 순서를 바꿀 수 있어요.'),
        bar, list);
    } catch (e) { A.failed(e); }
  }

  // ================= 편집기 =================
  async function showEditor(param) {
    A.loading();
    let templates, draft, key = null, sha = null, isNew = false;
    let datasets = {}, users = [];
    try {
      [templates, datasets] = await Promise.all([A.loadTemplates(true), A.loadDatasets(true)]);
      if (param !== 'new' && !param.startsWith('new?')) {
        users = (await A.charactersUsing(param)).flatMap(c => Object.entries(c.sheets || {})
          .filter(([k, sh]) => ((sh && sh.template) || k) === param)
          .map(([k, sh]) => ({ name: sh.title ? `${c.name}(${sh.title})` : c.name, sheet: sh || {} })));
      }
    } catch (e) { return A.failed(e); }

    if (param === 'new' || param.startsWith('new?')) {
      isNew = true;
      const from = new URLSearchParams(param.split('?')[1] || '').get('from');
      if (from && templates[from] && !templates[from].error) {
        draft = clone(templates[from]);
        draft.rule = (draft.rule || from) + ' 사본';
      } else {
        draft = {
          rule: '', description: '', accent: '#7c5cff',
          sections: [{ title: '능력치', layout: 'grid', fields: [{ id: '근력', label: '근력', type: 'number', default: 10 }] }],
        };
      }
    } else {
      key = param;
      try { const r = await A.readJson(`templates/${key}.json`); draft = r.data; sha = r.sha; }
      catch (e) { return A.failed(e); }
    }
    draft.sections = draft.sections || [];
    for (const s of draft.sections) {
      if (s.type === 'list') { s.columns = s.columns || []; } else if (s.type !== 'grid') s.fields = s.fields || [];
    }
    CV.toCanvas(draft); // 예전 배치 방식 템플릿은 격자 배치로 바꿔서 엶 (저장해야 반영)

    // 새로 만든 칸은 "이름"을 고치면 수식 이름도 따라 바뀜. 이미 시트가 있을 수 있는 칸은 고정.
    const fresh = new WeakSet(), manual = new WeakSet();
    if (isNew) for (const s of draft.sections) { fresh.add(s); for (const x of s.fields || s.columns || []) fresh.add(x); }

    const preview = { values: {}, lists: {} };
    let focusTarget = null;

    // ---------- 캐릭터 시트에 값이 들어 있는지 (있으면 지우기·이름 바꾸기 막음) ----------
    const hasValue = v => v !== undefined && v !== null && v !== '' && v !== false;
    const namesOf = list => list.slice(0, 3).map(u => u.name).join(', ') + (list.length > 3 ? ` 외 ${list.length - 3}명` : '');
    const fieldUsers = id => users.filter(u => hasValue((u.sheet.values || {})[id]));
    const listUsers = id => users.filter(u => ((u.sheet.lists || {})[id] || []).some(r => Object.values(r).some(hasValue)));
    const colUsers = (lid, cid) => users.filter(u => ((u.sheet.lists || {})[lid] || []).some(r => hasValue(r[cid])));
    function guard(found, what) {
      if (!found.length) return true;
      A.toast(`${what}에는 ${namesOf(found)} 캐릭터의 값이 들어 있어서 바꿀 수 없어요. 캐릭터 시트에서 값을 지우고 저장하면 바꿀 수 있어요.`, true);
      return false;
    }
    const fieldGuard = f => fresh.has(f) || guard([...new Set([f, ...(f.parts || [])].flatMap(x => fieldUsers(x.id)))], `"${labelOf(f)}" 칸`);
    const colGuard = (l, c) => fresh.has(c) || fresh.has(l) || guard(colUsers(l.id, c.id), `"${l.title}" 표의 "${c.label || c.id}" 열`);
    const listGuard = l => fresh.has(l) || guard(listUsers(l.id), `"${l.title || l.id}" 표`);

    // ---------- 이름 목록 / 이름 바꾸기 ----------
    // 칸 + 칸에 붙은 추가 숫자(parts)까지 모두
    const parentOf = new WeakMap();
    const allFields = () => draft.sections.filter(s => s.type !== 'list' && s.type !== 'grid').flatMap(s => s.fields.flatMap(f => {
      for (const p of f.parts || []) parentOf.set(p, f);
      return [f, ...(f.parts || [])];
    }));
    const labelOf = x => (parentOf.has(x) ? `${parentOf.get(x).label || parentOf.get(x).id} ${x.label || ''}`.trim() : (x.label || x.id));
    const allLists = () => draft.sections.filter(s => s.type === 'list');
    const allGrids = () => draft.sections.filter(s => s.type === 'grid');
    function eachFormula(fn) {
      for (const f of allFields()) {
        if (f.formula !== undefined) f.formula = fn(f.formula, null);
        if (f.max !== undefined && typeof f.max === 'string') f.max = fn(f.max, null);
      }
      for (const l of allLists()) for (const c of l.columns) if (c.formula !== undefined) c.formula = fn(c.formula, l);
      const cc = draft.ccfolia;
      if (cc) {
        const inText = (t, list) => E.mapBraces(t, e => (e.trim().startsWith('#') ? `{${e}}` : `{${fn(e, list)}}`));
        for (const st of cc.status || []) {
          if (st.value != null) st.value = fn(String(st.value), null);
          if (st.max != null) st.max = fn(String(st.max), null);
        }
        if (cc.initiative != null) cc.initiative = fn(String(cc.initiative), null);
        for (const p of cc.params || []) p.value = inText(p.value, null);
        cc.commands = inText(cc.commands, null);
        cc.memo = inText(cc.memo, null);
        if (cc.color) cc.color = inText(cc.color, null);
        for (const lc of cc.listCommands || []) lc.line = inText(lc.line, allLists().find(l => l.id === lc.list) || null);
      }
    }
    // 이름이 바뀌면 화면에 떠 있는 수식 입력칸도 새 내용으로 맞춤
    function syncFormulaInputs() {
      for (const inp of editorCol.querySelectorAll('.formula-input, .tpl-text')) {
        const [obj, prop] = inp._bind || [];
        if (obj && (obj[prop] ?? '') !== inp.value) inp.value = obj[prop] ?? '';
      }
    }
    function renameField(f, newId) {
      const old = f.id;
      if (old === newId) return;
      f.id = newId;
      if (old) eachFormula((src, list) => (list && list.columns.some(c => c.id === old) ? src : replaceIdent(src, old, newId)));
      if (old in preview.values) { preview.values[newId] = preview.values[old]; delete preview.values[old]; }
      syncFormulaInputs();
    }
    function renameGrid(g, newId) {
      const old = g.id;
      if (old === newId) return;
      g.id = newId;
      if (old) eachFormula(src => src.replace(new RegExp(`(["'])${esc(old)}\\1`, 'g'), `$1${newId}$1`));
      if (preview.grids && old in preview.grids) { preview.grids[newId] = preview.grids[old]; delete preview.grids[old]; }
    }
    function renameList(l, newId) {
      const old = l.id;
      if (old === newId) return;
      l.id = newId;
      for (const lc of (draft.ccfolia && draft.ccfolia.listCommands) || []) if (lc.list === old) lc.list = newId;
      if (old) eachFormula(src => src.replace(new RegExp(`(["'])${esc(old)}\\1`, 'g'), `$1${newId}$1`));
      if (old in preview.lists) { preview.lists[newId] = preview.lists[old]; delete preview.lists[old]; }
      syncFormulaInputs();
    }
    function renameColumn(l, c, newId) {
      const old = c.id;
      if (old === newId) return;
      c.id = newId;
      if (!old) return;
      const totalRe = new RegExp(`((?:목록합계|total)\\s*\\(\\s*(["'])${esc(l.id)}\\2\\s*,\\s*)(["'])${esc(old)}\\3`, 'g');
      eachFormula((src, list) => {
        let out = src.replace(totalRe, `$1$3${newId}$3`);
        if (list === l) out = replaceIdent(out, old, newId);
        return out;
      });
      for (const row of [...(preview.lists[l.id] || []), ...(l.defaultRows || [])]) {
        if (old in row) { row[newId] = row[old]; delete row[old]; }
      }
      syncFormulaInputs();
    }

    // ---------- 변경 처리 ----------
    function changed() { A.setDirty(true); normalizeLayout(); refreshPreview(); }

    // ---------- 격자 배치 정리: 자리가 없는 것은 빈 곳에, 겹친 것은 비켜서, 모두 위로 붙임 ----------
    const isFieldSec = x => x.type !== 'list' && x.type !== 'grid';
    const colsOf = x => Math.max(1, Math.min(12, Math.floor(x.cols) || 4));
    function normPos(o, cols) { o.pos = CV.posOf(o, cols); return o.pos; }
    function normalizeLayout() {
      const C = CV.COLS;
      const secs = draft.sections;
      const placed = secs.filter(x => x.pos).map(x => normPos(x, C));
      for (const x of secs) if (!x.pos) { x.pos = CV.freeSpot(placed, C, x.type === 'grid' ? C : 6, x.type === 'grid' ? 3 : 2); placed.push(x.pos); }
      CV.settle(secs.map(x => x.pos), C);
      for (const x of secs) {
        if (!isFieldSec(x)) continue;
        x.cols = colsOf(x);
        const shown = x.fields.filter(f => !f.folded);
        const fp = shown.filter(f => f.pos).map(f => normPos(f, x.cols));
        for (const f of shown) if (!f.pos) { f.pos = CV.freeSpot(fp, x.cols, (f.parts || []).length ? 2 : 1, 1); fp.push(f.pos); }
        CV.settle(shown.map(f => f.pos), x.cols);
        // 칸 순서(탭 이동·좁은 화면·코코포리아 순서)도 배치 순서대로
        const order = f => (f.folded || !f.pos ? 1e9 : f.pos.y * 100 + f.pos.x);
        x.fields.sort((a, b) => order(a) - order(b));
      }
      secs.sort((a, b) => (a.pos.y * 100 + a.pos.x) - (b.pos.y * 100 + b.pos.x));
    }
    function seedPreview() {
      for (const f of allFields()) if (f.default !== undefined && !(f.id in preview.values)) preview.values[f.id] = f.default;
      for (const l of allLists()) if (!preview.lists[l.id]) preview.lists[l.id] = clone(l.defaultRows || []);
    }
    let previewTimer;
    function refreshPreview(now) {
      clearTimeout(previewTimer);
      const run = () => {
        seedPreview();
        try {
          previewBody.replaceChildren(A.renderSheet(draft, preview, { readonly: false, onChange: () => activeInput && updateHelper(), datasets, exactLayout: true }));
          decorate();
        } catch (e) {
          previewBody.replaceChildren(h('p', { class: 'muted' }, '미리보기를 그리지 못했어요: ' + e.message));
        }
        if (activeInput) updateHelper();
        updateCcPreview();
      };
      if (now) run(); else previewTimer = setTimeout(run, 200);
    }

    // ---------- 작은 부품 ----------
    const iconBtn = (text, title, fn, cls) => h('button', { type: 'button', class: 'icon-btn ' + (cls || ''), title, onclick: fn }, text);
    // 접어 두기: 숨기기와 달리 읽기·공유 화면에서도 제목은 보이고, 눌러서 펼쳐 볼 수 있음
    const toggleBtn = x => iconBtn(x.toggle ? '▸' : '▾', x.toggle
      ? '접어 둔 카드예요 — 시트에서 제목만 보이고 눌러서 펼쳐요. 눌러서 늘 펼쳐 두기'
      : '접어 두기 (시트에서 제목만 보이고, 읽기 모드에서도 눌러서 펼쳐 볼 수 있어요)', () => {
      if (x.toggle) delete x.toggle; else x.toggle = true; changed(); drawEditor();
    }, 'toggle-btn' + (x.toggle ? ' on' : ''));
    // 끌어서 순서 바꾸기 (⠿ 손잡이 또는 빈 곳을 꾹 눌러서)
    const dragHandle = A.dragHandle;
    const moved = () => { changed(); drawEditor(); };
    const drag = (el, obj, extra) => A.draggable(el, obj, { onMoved: moved, ...(extra || {}) });
    function kindSelect(kinds, current, onPick) {
      return h('select', { class: 'kind', onchange: e => onPick(e.target.value) },
        kinds.map(([v, l]) => h('option', { value: v, selected: v === current }, l)));
    }
    function applyKind(x, kind) {
      if (kind === 'formula') {
        delete x.type; delete x.default; delete x.max; delete x.options;
        if (x.formula === undefined) x.formula = '';
      } else {
        delete x.formula;
        x.type = kind;
        if (kind !== 'select') delete x.options; else x.options = x.options || ['선택지1', '선택지2'];
        if (kind !== 'number') delete x.max;
        if (kind === 'checkbox') x.default = x.default === true; else if (typeof x.default === 'boolean') delete x.default;
      }
    }
    function labelInput(x, onRename, placeholder) {
      const idOut = { el: null };
      const input = h('input', {
        class: 'label-input', value: x.label ?? x.title ?? '', placeholder,
        oninput: e => {
          if ('title' in x && !('label' in x)) x.title = e.target.value; else x.label = e.target.value;
          if (fresh.has(x) && !manual.has(x)) { onRename(e.target.value); if (idOut.el) idOut.el.value = x.id; }
          changed();
        },
      });
      if (focusTarget === x) requestAnimationFrame(() => { input.focus(); input.select(); focusTarget = null; });
      return [input, idOut];
    }
    function idEditor(x, idOut, label, apply, taken, canChange) {
      const input = h('input', { value: x.id || '', class: 'id-input', spellcheck: false });
      idOut.el = input;
      input.addEventListener('change', () => {
        const v = input.value.trim();
        if (!validId(v)) { A.toast('수식 이름은 띄어쓰기·기호 없이 글자, 숫자, _만 쓸 수 있고 숫자로 시작할 수 없어요.', true); input.value = x.id; return; }
        if (v !== x.id && taken().has(v)) { A.toast(`'${v}'는 이미 쓰고 있는 이름이에요.`, true); input.value = x.id; return; }
        if (v !== x.id && canChange && !canChange()) { input.value = x.id; return; }
        manual.add(x);
        apply(v);
        changed();
      });
      return h('label', { class: 'mini-field' }, h('span', {}, label), input);
    }

    // ---------- 수식 입력칸 + 도우미 ----------
    let activeInput = null;
    const helper = h('div', { class: 'formula-helper' });
    helper.addEventListener('mousedown', e => { if (e.target.closest('button')) e.preventDefault(); });

    function formulaInput(obj, prop, ctx, placeholder) {
      const input = h('input', {
        class: 'formula-input', value: obj[prop] ?? '', placeholder: placeholder || '예: 내림((근력 - 10) / 2)',
        spellcheck: false, autocomplete: 'off',
      });
      input._ctx = ctx;
      input._bind = [obj, prop];
      input.addEventListener('input', () => {
        const v = input.value;
        if ((prop === 'max' || ctx.optional) && !v.trim()) delete obj[prop]; else obj[prop] = v;
        changed();
        updateHelper();
      });
      input.addEventListener('focus', () => {
        activeInput = input;
        input.closest('.formula-wrap').after(helper);
        buildHelper();
        previewBody.classList.add('picking');
      });
      input.addEventListener('blur', () => {
        if (activeInput === input) { activeInput = null; helper.remove(); previewBody.classList.remove('picking'); }
      });
      return h('div', { class: 'formula-wrap' }, h('span', { class: 'fx' }, 'fx'), input);
    }

    function insert(text, caretBack) {
      const inp = activeInput;
      if (!inp) return;
      const s = inp.selectionStart ?? inp.value.length, e = inp.selectionEnd ?? s;
      inp.setRangeText(text, s, e, 'end');
      if (caretBack) { const p = inp.selectionEnd - caretBack; inp.setSelectionRange(p, p); }
      inp.dispatchEvent(new Event('input'));
    }
    function insertFunc(name) {
      const inp = activeInput;
      if (!inp) return;
      const s = inp.selectionStart ?? inp.value.length, e = inp.selectionEnd ?? s;
      const sel = inp.value.slice(s, e);
      if (sel) insert(`${name}(${sel})`); else insert(`${name}()`, 1);
    }
    const chip = (text, title, fn, cls) => h('button', { type: 'button', class: 'chip ' + (cls || ''), title, onclick: fn }, text);
    const group = (title, chips) => (chips.length ? h('div', { class: 'fx-group' }, h('span', { class: 'fx-title' }, title), chips) : null);

    const helperStatus = h('div', { class: 'fx-status' });
    function buildHelper() {
      if (!activeInput) return;
      const ctx = activeInput._ctx;
      const fields = allFields().filter(f => f.id && f !== ctx.field);
      const listChips = allLists().flatMap(l => [
        chip(`${l.title || l.id} 줄 수`, `개수("${l.id}")`, () => insert(`개수("${l.id}")`), 'list'),
        // 글자·선택 목록 열: 그 열에 특정 이름이 있는지 (있으면 1) → 만약(…, 2, 0) 처럼 써요
        ...l.columns.filter(c => c.id && c.formula === undefined && ['text', 'select', undefined].includes(c.type))
          .map(c => chip(`${l.title || l.id}: ${c.label || c.id}에 이름이 있으면`, `만약(포함("${l.id}", "${c.id}", "이름"), 2, 0) — 그 열에 이름이 있으면 2, 없으면 0`,
            () => insert(`만약(포함("${l.id}", "${c.id}", ""), 2, 0)`, 9), 'list')),
        ...l.columns.filter(c => c.id && (c.type === 'number' || c.formula !== undefined) && !(ctx.list === l && c === ctx.col))
          .map(c => chip(`${l.title || l.id}: ${c.label || c.id} 합계`, `목록합계("${l.id}", "${c.id}")`,
            () => insert(`목록합계("${l.id}", "${c.id}")`), 'list')),
      ]);
      helper.replaceChildren(...[
        helperStatus,
        group('함수', FUNCS.map(([n, d]) => chip(`${n}( )`, d, () => insertFunc(n), 'fn'))),
        group('기호', OPS.map(o => chip(OP_LABEL[o] || o, OP_TIP[o] || '', () => insert(o === '(' || o === ')' ? o : ` ${o} `), 'op'))),
        ctx.list ? group('같은 줄', ctx.list.columns.filter(c => c.id && c !== ctx.col)
          .map(c => chip(c.label || c.id, `수식 이름: ${c.id}`, () => insert(c.id), 'ref'))) : null,
        group('칸', fields.map(f => chip(labelOf(f), `수식 이름: ${f.id}`, () => insert(f.id), 'ref'))),
        group('표', listChips),
        group('특기표', allGrids().flatMap(g => [
          chip(`${g.title || g.id}: 달성치`, `달성치("${g.id}", "특기 이름") — 그 특기의 달성치`, () => insert(`달성치("${g.id}", "")`, 2), 'list'),
          chip(`${g.title || g.id}: 습득`, `습득("${g.id}", "특기 이름") — 습득했으면 1, 아니면 0`, () => insert(`습득("${g.id}", "")`, 2), 'list'),
        ])),
        h('p', { class: 'fx-tip' }, '수식을 쓰는 중에 미리보기의 칸을 클릭해도 그 칸이 들어가요. 체크 칸은 체크하면 1, 아니면 0이에요.'),
      ].filter(Boolean));
      updateHelper();
    }
    function updateHelper() {
      if (!activeInput) return;
      const ctx = activeInput._ctx;
      const src = activeInput.value.trim();
      helperStatus.className = 'fx-status';
      if (!src) { helperStatus.textContent = ctx.emptyText || (ctx.optional ? '비워 두면 최대값을 표시하지 않아요.' : '수식을 입력해 주세요.'); return; }
      try {
        E.parse(src);
        seedPreview();
        const r = E.makeResolver(draft, preview);
        let v;
        if (ctx.list) {
          if (!(preview.lists[ctx.list.id] || []).length) {
            helperStatus.textContent = '✓ 수식 형식은 맞아요. (표에 줄이 생기면 계산 결과가 보여요)';
            helperStatus.classList.add('ok');
            return;
          }
          v = r.cell(ctx.list.id, 0, ctx.col.id);
        } else v = ctx.expr || ctx.prop === 'max' ? r.evaluate(src) : r.get(ctx.field.id);
        helperStatus.textContent = `✓ 미리보기 값: ${E.format(v)}` + (ctx.list ? ' (첫 줄 기준)' : '');
        helperStatus.classList.add('ok');
        // 비교(=, !=)에서 칸 이름을 따옴표로 감싸면 칸 값이 아니라 글자 그대로 비교돼서 늘 틀리게 나옴 → 알려 줌
        const ids = new Set([...allFields().map(f => f.id), ...(ctx.list ? ctx.list.columns.map(c => c.id) : [])]);
        const quoted = [];
        (function walk(n) {
          if (!n || typeof n !== 'object') return;
          if (n.k === 'bin' && (n.op === '==' || n.op === '!=')) for (const side of [n.a, n.b]) if (side.k === 'str' && ids.has(side.v)) quoted.push(side.v);
          for (const x of [n.a, n.b, ...(n.args || [])]) walk(x);
        })(E.parse(src));
        if (quoted.length) {
          helperStatus.textContent += ` — ⚠ "${quoted[0]}"는 따옴표 안이라 칸 값이 아니라 글자 "${quoted[0]}" 그대로 비교돼요. 칸 값을 쓰려면 따옴표 없이 ${quoted[0]}라고 써 주세요.`;
          helperStatus.classList.remove('ok'); helperStatus.classList.add('err');
        }
      } catch (e) {
        helperStatus.textContent = '⚠ ' + e.message;
        helperStatus.classList.add('err');
      }
    }

    // ---------- 칸 한 줄 편집 ----------
    function fieldItem(f, s, i) {
      const kind = kindOf(f);
      const takenIds = () => new Set(allFields().filter(x => x !== f).map(x => x.id));
      const [label, idOut] = labelInput(f, v => renameField(f, idFrom(v, takenIds())), '칸 이름');
      const extra = [];
      if (kind === 'formula') extra.push(formulaInput(f, 'formula', { field: f, prop: 'formula' }));
      if (kind === 'number') {
        extra.push(h('div', { class: 'row tight' },
          h('label', { class: 'mini-field' }, h('span', {}, '처음 값'),
            h('input', { type: 'number', value: f.default ?? '', oninput: e => {
              if (e.target.value === '') delete f.default; else f.default = Number(e.target.value);
              delete preview.values[f.id]; changed();
            } })),
          h('div', { class: 'mini-field grow' }, h('span', {}, '최대값 표시 (선택, 수식 가능)'),
            formulaInput(f, 'max', { field: f, prop: 'max', optional: true }, '예: 최대체력'))));
      }
      if (kind === 'number' || kind === 'formula') extra.push(partsEditor(f));
      if (kind === 'text' || kind === 'textarea') {
        extra.push(h('label', { class: 'mini-field' }, h('span', {}, '처음 값'),
          h('input', { value: f.default ?? '', oninput: e => { if (e.target.value) f.default = e.target.value; else delete f.default; changed(); } })));
      }
      if (kind === 'checkbox') {
        extra.push(h('label', { class: 'mini-check' }, h('input', { type: 'checkbox', checked: !!f.default,
          onchange: e => { f.default = e.target.checked; delete preview.values[f.id]; changed(); } }), '처음부터 체크'));
      }
      if (kind === 'select') {
        extra.push(h('label', { class: 'mini-field' }, h('span', {}, '선택지 (쉼표로 구분)'),
          h('input', { value: (f.options || []).join(', '), oninput: e => {
            f.options = e.target.value.split(',').map(x => x.trim()).filter(Boolean); changed();
          } })));
      }
      const item = h('div', { class: 'tpl-item' + (kind === 'formula' ? ' is-formula' : '') + (f.folded ? ' is-folded' : '') },
        h('div', { class: 'item-head' }, label,
          kindSelect(KINDS, kind, v => {
            if (v === 'formula' && !fieldGuard(f)) return drawEditor();
            if (v !== 'number' && v !== 'formula' && (f.parts || []).length) {
              if (!(f.parts || []).every(p => partGuard(p))) return drawEditor();
              if (!confirm('글자·체크 같은 칸에는 숫자를 붙일 수 없어서, 붙여 둔 숫자가 지워져요. 바꿀까요?')) return drawEditor();
              delete f.parts; delete f.sep; delete f.mainLabel;
            }
            applyKind(f, v); changed(); drawEditor();
          }),
          iconBtn(f.folded ? '🙈' : '👁', f.folded ? '접어 둔 칸이에요 (시트에서 평소엔 숨김). 눌러서 다시 보이게' : '시트에서 접어 두기 (평소엔 숨김)', () => {
            if (f.folded) { delete f.folded; delete f.pos; } else f.folded = true; changed(); drawEditor();
          }, 'fold-btn' + (f.folded ? ' on' : '')),
          iconBtn('⧉', '이 칸 복제', () => duplicateField(s, s.fields.indexOf(f))),
          iconBtn('×', '칸 삭제', () => { if (!fieldGuard(f)) return; s.fields.splice(s.fields.indexOf(f), 1); changed(); drawEditor(); }, 'danger')),
        f.folded ? h('span', { class: 'fold-tag' }, '🙈 접힌 칸 — 시트에서는 "숨긴 칸 보기"를 눌러야 보이고, 읽기 모드·공유 화면에서는 안 보여요. 수식에서는 그대로 쓸 수 있어요.') : null,
        extra,
        h('details', { class: 'more' }, h('summary', {}, '더보기'),
          h('div', { class: 'row tight' },
            idEditor(f, idOut, '수식에서 부르는 이름', v => renameField(f, v), takenIds, () => fieldGuard(f)),
            h('label', { class: 'mini-field grow' }, h('span', {}, '도움말 (칸 아래 작은 글씨)'),
              h('input', { value: f.hint || '', oninput: e => { if (e.target.value) f.hint = e.target.value; else delete f.hint; changed(); } })))));
      return item;
    }

    // ---------- 칸 복제 (붙은 숫자까지, 수식 안의 자기 이름도 새 이름으로) ----------
    function duplicateField(s, i) {
      const src = s.fields[i];
      const copy = clone(src);
      const taken = new Set(allFields().map(x => x.id));
      copy.label = (src.label || src.id) + ' 사본';
      const map = new Map();
      copy.id = idFrom(copy.label, taken); taken.add(copy.id); map.set(src.id, copy.id);
      for (const p of copy.parts || []) {
        const nid = idFrom(copy.label + (p.label || '숫자'), taken); taken.add(nid); map.set(p.id, nid); p.id = nid;
      }
      const remap = f => { let out = f; for (const [o, n] of map) out = replaceIdent(out, o, n); return out; };
      for (const x of [copy, ...(copy.parts || [])]) {
        if (x.formula !== undefined) x.formula = remap(x.formula);
        if (typeof x.max === 'string') x.max = remap(x.max);
        fresh.add(x);
      }
      delete copy.pos;
      s.fields.splice(i + 1, 0, copy);
      focusTarget = copy; changed();
      if (sel && sel.kind === 'field') sel = { kind: 'field', s, f: copy };
      drawEditor();
    }

    // ---------- 한 칸에 숫자 여러 개 (예: 보통 / 절반 / 1/5) ----------
    const PART_KINDS = [['number', '숫자'], ['formula', '자동 계산']];
    const partGuard = p => fresh.has(p) || guard(fieldUsers(p.id), `"${labelOf(p)}" 숫자`);
    function partsEditor(f) {
      const parts = f.parts || [];
      const takenIds = p => () => new Set(allFields().filter(x => x !== p).map(x => x.id));
      const rows = parts.map((p, j) => {
        const pk = p.formula !== undefined ? 'formula' : 'number';
        const idOut = { el: null };
        const nameIn = h('input', { class: 'label-input', value: p.label || '', placeholder: '이름 (예: 절반)', oninput: e => {
          p.label = e.target.value;
          if (fresh.has(p) && !manual.has(p)) { renameField(p, idFrom((f.label || f.id) + e.target.value, takenIds(p)())); if (idOut.el) idOut.el.value = p.id; }
          changed();
        } });
        if (focusTarget === p) requestAnimationFrame(() => { nameIn.focus(); nameIn.select(); focusTarget = null; });
        return drag(h('div', { class: 'part-item' },
          h('div', { class: 'item-head' }, dragHandle(), h('span', { class: 'part-sep' }, f.sep || '/'), nameIn,
            kindSelect(PART_KINDS, pk, v => {
              if (v === 'formula' && !partGuard(p)) return drawEditor();
              applyKind(p, v); changed(); drawEditor();
            }),
            iconBtn('×', '이 숫자 빼기', () => {
              if (!partGuard(p)) return;
              parts.splice(j, 1);
              if (!parts.length) { delete f.parts; delete f.sep; delete f.mainLabel; }
              changed(); drawEditor();
            }, 'danger')),
          pk === 'formula'
            ? formulaInput(p, 'formula', { field: p, prop: 'formula' }, `예: 내림(${f.id} / 2)`)
            : h('label', { class: 'mini-field' }, h('span', {}, '처음 값'),
              h('input', { type: 'number', value: p.default ?? '', oninput: e => {
                if (e.target.value === '') delete p.default; else p.default = Number(e.target.value);
                delete preview.values[p.id]; changed();
              } })),
          h('details', { class: 'more' }, h('summary', {}, '더보기'),
            idEditor(p, idOut, '수식에서 부르는 이름', v => renameField(p, v), takenIds(p), () => partGuard(p)))), p);
      });
      return h('div', { class: 'parts-box' },
        parts.length ? h('div', { class: 'row tight' },
          h('label', { class: 'mini-field' }, h('span', {}, '첫 숫자 이름 (선택)'),
            h('input', { value: f.mainLabel || '', placeholder: '예: 보통', oninput: e => { if (e.target.value) f.mainLabel = e.target.value; else delete f.mainLabel; changed(); } })),
          h('label', { class: 'mini-field' }, h('span', {}, '숫자 사이 기호'),
            h('input', { value: f.sep || '/', style: 'width:64px', oninput: e => {
              if (e.target.value && e.target.value !== '/') f.sep = e.target.value; else delete f.sep;
              changed();
              for (const el of e.target.closest('.parts-box').querySelectorAll('.part-sep')) el.textContent = f.sep || '/';
            } }))) : null,
        parts.length ? A.dropList(h('div', { class: 'part-list' }, rows), parts, f) : null,
        h('button', { type: 'button', class: 'add-row small-btn', onclick: () => {
          const p = { label: parts.length ? '' : '절반', formula: parts.length ? '' : `내림(${f.id} / 2)` };
          p.id = idFrom((f.label || f.id) + (p.label || '숫자'), new Set(allFields().map(x => x.id)));
          fresh.add(p);
          f.parts = [...parts, p];
          focusTarget = p; changed(); drawEditor();
        } }, parts.length ? '+ 숫자 하나 더' : '+ 이 칸에 숫자 더 붙이기 (예: 보통 / 절반 / 1/5)'));
    }

    // ---------- 표의 열 한 줄 편집 ----------
    function columnItem(c, l, i) {
      const kind = kindOf(c);
      const takenIds = () => new Set(l.columns.filter(x => x !== c).map(x => x.id));
      const [label, idOut] = labelInput(c, v => renameColumn(l, c, idFrom(v, takenIds())), '열 이름');
      const ds = l.source && datasets[l.source.dataset];
      const fromField = c.from && ds && (ds.fields || []).find(x => x.id === c.from);
      return drag(h('div', { class: 'tpl-item' + (kind === 'formula' ? ' is-formula' : '') + (fromField ? ' is-linked' : '') },
        fromField ? h('span', { class: 'link-tag' }, (ds.fields[0].id === c.from ? '🔎 찾기 이름 · ' : '🔗 ') +
          `${ds.name}${(l.source.also || []).length ? ` 외 ${(l.source.also || []).length}개` : ''}의 "${fromField.label || fromField.id}"`) : null,
        h('div', { class: 'item-head' }, dragHandle(), label,
          kindSelect(COL_KINDS, kind, v => { if (v === 'formula' && !colGuard(l, c)) return drawEditor(); applyKind(c, v); changed(); drawEditor(); }),
          iconBtn('⧉', '이 열 복제', () => {
            const copy = clone(c);
            delete copy.from;
            copy.label = (c.label || c.id) + ' 사본';
            copy.id = idFrom(copy.label, new Set(l.columns.map(x => x.id)));
            fresh.add(copy); l.columns.splice(i + 1, 0, copy); focusTarget = copy; changed(); drawEditor();
          }),
          iconBtn('×', '열 삭제', () => { if (!colGuard(l, c)) return; l.columns.splice(i, 1); changed(); drawEditor(); }, 'danger')),
        kind === 'formula' ? formulaInput(c, 'formula', { list: l, col: c, prop: 'formula' }, '예: 기본 + 투자') : null,
        kind === 'select' ? h('label', { class: 'mini-field' }, h('span', {}, '선택지 (쉼표로 구분)'),
          h('input', { value: (c.options || []).join(', '), placeholder: '예: 주동작, 보조동작, 상시', oninput: e => {
            c.options = e.target.value.split(',').map(x => x.trim()).filter(Boolean); changed();
          } })) : null,
        h('details', { class: 'more' }, h('summary', {}, '더보기'),
          idEditor(c, idOut, '수식에서 부르는 이름', v => renameColumn(l, c, v), takenIds, () => colGuard(l, c)))), c);
    }

    function defaultRowsEditor(l) {
      l.defaultRows = l.defaultRows || [];
      const cols = l.columns.filter(c => c.formula === undefined && c.id);
      const sync = () => { preview.lists[l.id] = clone(l.defaultRows); };
      const link = E.linkInfo(l, datasets);
      const dlId = link ? 'tdl-' + Math.random().toString(36).slice(2, 9) : null;
      const cellInput = (row, c) => {
        if (link && c === link.keyCol) {
          // 확정할 때(목록에서 고르기·Enter·다른 곳 누르기)만 불러옴 — 한글 조합 중에 끊기지 않게
          const el = h('input', { value: row[c.id] ?? '', list: dlId, placeholder: `${link.keyField.label || '이름'} 입력` });
          const commit = () => {
            const entry = link.find(el.value);
            if (entry) { E.fillRowFromEntry(l, row, entry, link.keyCol); sync(); changed(); drawEditor(); }
          };
          el.addEventListener('input', e => {
            const v = el.value;
            if (v === '') delete row[c.id]; else row[c.id] = v;
            sync(); changed();
            if (e.inputType === 'insertReplacementText') commit();
          });
          el.addEventListener('change', commit);
          return el;
        }
        if (link && c.from && link.find(row[link.keyCol.id])) {
          return h('div', { class: 'linked-cell is-linked', title: '데이터에서 불러온 값' }, String(row[c.id] ?? ''));
        }
        if (c.type === 'checkbox') return h('input', { type: 'checkbox', checked: !!row[c.id], onchange: e => { row[c.id] = e.target.checked; sync(); changed(); } });
        if (c.type === 'select') {
          const cur = String(row[c.id] ?? ''), opts = c.options || [];
          return h('select', { onchange: e => { if (e.target.value) row[c.id] = e.target.value; else delete row[c.id]; sync(); changed(); } },
            h('option', { value: '' }, '—'), opts.map(o => h('option', { value: o, selected: o === cur }, o)),
            cur && !opts.includes(cur) ? h('option', { value: cur, selected: true }, cur) : null);
        }
        return h('input', { type: c.type === 'number' ? 'number' : 'text', value: row[c.id] ?? '', oninput: e => {
          const v = e.target.value;
          if (v === '') delete row[c.id]; else row[c.id] = c.type === 'number' ? Number(v) : v;
          sync(); changed();
        } });
      };
      return h('div', { class: 'default-rows' },
        link ? h('datalist', { id: dlId }, [...link.map.keys()].map(k => h('option', { value: k }))) : null,
        h('div', { class: 'sub-title' }, '처음부터 들어 있을 줄 ', h('span', { class: 'muted' }, link
          ? `(선택 — 비워 둬도 돼요. 캐릭터 시트의 이 표에서 ${link.keyField.label || '이름'}을 적으면 체크한 항목이 자동으로 채워져요. 모든 캐릭터가 처음부터 가져야 하는 것만 여기에 넣어요)`
          : '(예: 기본 기능 목록. 비워 둬도 돼요)')),
        l.defaultRows.length ? h('div', { class: 'list-wrap' }, h('table', { class: 'list mini' },
          h('thead', {}, h('tr', {}, cols.map(c => h('th', {}, c.label || c.id)), h('th'))),
          h('tbody', {}, l.defaultRows.map((row, i) => h('tr', {}, cols.map(c => h('td', {}, cellInput(row, c))),
            h('td', {}, iconBtn('×', '줄 삭제', () => { l.defaultRows.splice(i, 1); sync(); changed(); drawEditor(); }, 'danger'))))))) : null,
        h('button', { type: 'button', class: 'add-row', onclick: () => { l.defaultRows.push({}); sync(); changed(); drawEditor(); } }, '+ 처음 줄 추가'));
    }

    // ---------- 표 ↔ 데이터 보관함 연결 ----------
    function dataLinkPanel(l) {
      const keys = Object.keys(datasets).filter(k => !datasets[k].error);
      const cur = l.source && l.source.dataset;
      const ds = cur && datasets[cur];
      const ruleLabel = k => { const r = datasets[k].rule; return r ? ((templates[r] && templates[r].rule) || r) : '공용'; };
      const sorted = keys.sort((a, b) => (datasets[b].rule === key) - (datasets[a].rule === key) || String(datasets[a].name).localeCompare(String(datasets[b].name)));
      const select = h('select', { onchange: e => setSource(e.target.value) },
        h('option', { value: '' }, '연결 안 함'),
        sorted.map(k => h('option', { value: k, selected: k === cur }, `${datasets[k].name || k} (${ruleLabel(k)})`)));

      function setSource(k) {
        const linked = l.columns.filter(c => c.from);
        if (!k) {
          for (const c of linked) delete c.from;
          delete l.source; changed(); drawEditor(); return;
        }
        // 다른 데이터로 바꾸면: 예전 연결 열 중 값이 없는 건 지우고, 값이 있는 건 일반 열로 남김
        for (const c of linked) {
          if (c === l.columns.find(x => x.from === (ds && ds.fields[0] && ds.fields[0].id))) { delete c.from; continue; }
          if (fresh.has(c) || !colUsers(l.id, c.id).length) l.columns.splice(l.columns.indexOf(c), 1); else delete c.from;
        }
        // 함께 쓰던 데이터 중 새 첫 데이터와 구성이 같은 것만 남김
        const keepAlso = ((l.source && l.source.also) || []).filter(x => x !== k && datasets[x] && E.sameSchema(datasets[k], datasets[x]));
        l.source = keepAlso.length ? { dataset: k, also: keepAlso } : { dataset: k };
        const nds = datasets[k];
        const keyField = nds.fields[0];
        const first = l.columns.find(c => c.formula === undefined && (c.type || 'text') === 'text');
        if (first) first.from = keyField.id;
        else {
          const c = { id: idFrom(keyField.label || keyField.id, new Set(l.columns.map(x => x.id))), label: keyField.label || keyField.id, type: 'text', from: keyField.id };
          fresh.add(c); l.columns.unshift(c);
        }
        changed(); drawEditor();
      }
      function toggleField(fld, on) {
        if (on) {
          const type = fld.type === 'number' ? 'number' : fld.type === 'textarea' ? 'textarea' : fld.type === 'select' ? 'select' : 'text';
          const c = { id: idFrom(fld.label || fld.id, new Set(l.columns.map(x => x.id))), label: fld.label || fld.id, type, from: fld.id };
          if (type === 'select') c.options = [...(fld.options || [])];
          fresh.add(c); l.columns.push(c);
          for (const row of l.defaultRows || []) { const e = E.linkInfo(l, datasets); const en = e && e.find(row[e.keyCol.id]); if (en) E.fillRowFromEntry(l, row, en, e.keyCol); }
        } else {
          const c = l.columns.find(x => x.from === fld.id);
          if (c && !colGuard(l, c)) return drawEditor();
          if (c) l.columns.splice(l.columns.indexOf(c), 1);
        }
        preview.lists[l.id] = clone(l.defaultRows || []);
        changed(); drawEditor();
      }

      if (!keys.length) {
        return h('div', { class: 'data-link muted small' }, '위쪽 "데이터" 메뉴에서 데이터를 만들면 이 표에 연결해서 이름만 입력해도 내용이 채워지게 할 수 있어요. ',
          h('a', { href: '#/data/new' }, '데이터 만들기'));
      }
      // ----- 함께 불러올 데이터 (항목 구성이 첫 데이터와 같은 것만) -----
      const also = (l.source && l.source.also) || [];
      const setAlso = list => { if (list.length) l.source.also = list; else delete l.source.also; changed(); drawEditor(); };
      const describe = k => `${datasets[k].name || k} (${ruleLabel(k)})`;
      const whyNot = k => {
        const d = E.schemaDiff(ds, datasets[k]);
        if (d.keyDiff) return '찾기 이름(첫 항목)이 달라요';
        const parts = [];
        if (d.onlyA.length) parts.push(`없는 항목: ${d.onlyA.join(', ')}`);
        if (d.onlyB.length) parts.push(`더 있는 항목: ${d.onlyB.join(', ')}`);
        return parts.join(' / ') || '항목 구성이 달라요';
      };
      let alsoBox = null;
      if (ds) {
        const candidates = sorted.filter(k => k !== cur && !also.includes(k));
        const addSel = h('select', { onchange: e => { if (e.target.value) setAlso([...also, e.target.value]); } },
          h('option', { value: '' }, '+ 같은 구성의 데이터 더 연결…'),
          candidates.map(k => {
            const ok = E.sameSchema(ds, datasets[k]);
            return h('option', { value: k, disabled: !ok }, ok ? describe(k) : `${describe(k)} — ${whyNot(k)}`);
          }));
        // 같은 이름이 여러 데이터에 있으면 앞쪽 데이터 내용을 씀
        const link = E.linkInfo(l, datasets);
        const kf = ds.fields[0].id;
        let dup = 0;
        if (link) for (const k of link.keys.slice(1)) for (const e2 of (datasets[k] && datasets[k].entries) || []) {
          const nm = String(e2[kf] ?? '').trim();
          if (nm && link.from.get(nm) !== k) dup++;
        }
        alsoBox = h('div', { class: 'also-box' },
          h('div', { class: 'sub-title' }, '함께 불러올 데이터 ', h('span', { class: 'muted' }, '(항목 구성이 위 데이터와 똑같은 것만 고를 수 있어요)')),
          also.length ? h('div', { class: 'also-list' }, also.map((k, idx) => h('span', { class: 'also-chip' + (datasets[k] && E.sameSchema(ds, datasets[k]) ? '' : ' bad'),
            title: datasets[k] && !E.sameSchema(ds, datasets[k]) ? `구성이 달라져서 지금은 불러오지 않아요: ${whyNot(k)}` : '' },
            datasets[k] ? describe(k) : `${k} (없음)`,
            h('button', { type: 'button', class: 'ghost', title: '연결 빼기', onclick: () => setAlso(also.filter((_, j) => j !== idx)) }, '×')))) : null,
          candidates.length ? addSel : null,
          dup ? h('p', { class: 'muted small' }, `같은 ${ds.fields[0].label || '이름'}이 여러 데이터에 ${dup}개 있어요. 이럴 땐 위쪽(먼저 연결한) 데이터의 내용을 불러와요.`) : null);
      }
      return h('div', { class: 'data-link' },
        h('label', { class: 'mini-field' }, h('span', {}, '데이터 연결'), select),
        alsoBox,
        ds ? [
          h('div', { class: 'sub-title' }, '불러올 항목 ', h('span', { class: 'muted' }, '(체크한 항목이 열로 생겨요)')),
          h('div', { class: 'field-checks' }, (ds.fields || []).map((fld, idx) => h('label', { class: 'mini-check' },
            h('input', { type: 'checkbox', checked: idx === 0 || l.columns.some(c => c.from === fld.id), disabled: idx === 0,
              onchange: e => toggleField(fld, e.target.checked) }),
            fld.label || fld.id, idx === 0 ? h('small', { class: 'muted' }, ' (찾기 이름)') : null))),
          h('p', { class: 'muted small' }, `시트에서 "${(l.columns.find(c => c.from === ds.fields[0].id) || {}).label || ''}" 열에 ${ds.fields[0].label || '이름'}을 입력하면 나머지가 자동으로 채워져요. "데이터"에서 내용을 고치면 시트에도 바로 반영돼요.`),
        ] : null);
    }

    // ---------- 묶음(카드) 하나 ----------
    // ---------- 특기표(격자) 편집 ----------
    const gridUsers = id => users.filter(u => (((u.sheet.grids || {})[id] || {}).acquired || []).length);
    const gridGuard = (g, test, what) => {
      if (fresh.has(g)) return true;
      const found = users.filter(u => (((u.sheet.grids || {})[g.id] || {}).acquired || []).some(test || (() => true)));
      return guard(found, what || `"${g.title || g.id}" 특기표`);
    };
    function gridCard(g, i) {
      g.columns = g.columns || []; g.rows = g.rows || []; g.cells = g.cells || [];
      const takenGrids = () => new Set(allGrids().filter(x => x !== g).map(x => x.id));
      const idOut = { el: null };
      const titleInput = h('input', { class: 'section-title', value: g.title || '', placeholder: '특기표 이름 (예: 특기)',
        oninput: e => {
          g.title = e.target.value;
          if (fresh.has(g) && !manual.has(g)) { renameGrid(g, idFrom(e.target.value, takenGrids())); if (idOut.el) idOut.el.value = g.id; }
          changed();
        } });
      if (focusTarget === g) requestAnimationFrame(() => { titleInput.focus(); titleInput.select(); focusTarget = null; });
      const head = h('div', { class: 'section-head' }, h('span', { class: 'section-kind' }, '특기표'), titleInput,
        iconBtn(g.folded ? '🙈' : '👁', g.folded ? '숨긴 묶음이에요. 눌러서 다시 보이게' : '이 특기표를 시트에서 숨기기', () => {
          if (g.folded) delete g.folded; else g.folded = true; changed(); drawEditor();
        }, 'fold-btn' + (g.folded ? ' on' : '')),
        toggleBtn(g),
        iconBtn('×', '삭제', () => {
          if (!gridGuard(g)) return;
          if (!confirm(`"${g.title || g.id}" 특기표를 지울까요?`)) return;
          draft.sections.splice(draft.sections.indexOf(g), 1); changed(); drawEditor();
        }, 'danger'));

      const numIn = (prop, label, def, hint) => h('label', { class: 'mini-field' }, h('span', { title: hint || '' }, label),
        h('input', { type: 'number', value: g[prop] ?? def, style: 'width:90px', oninput: e => {
          if (e.target.value === '' || Number(e.target.value) === def) delete g[prop]; else g[prop] = Number(e.target.value);
          changed();
        } }));
      const R = g.rows.length, C = g.columns.length;
      const setCell = (r, c, v) => {
        while (g.cells.length <= r) g.cells.push([]);
        const row = g.cells[r];
        while (row.length <= c) row.push('');
        row[c] = v;
      };
      g.columnNames = g.columnNames || [];
      const labelsEditor = (arr, kind) => h('div', { class: 'grid-labels' },
        h('span', { class: 'filter-label' }, kind === 'col' ? '열 (번호·이름)' : '줄 이름'),
        arr.map((v, k) => {
          const num = h('input', { value: v, style: 'width:' + (kind === 'col' ? '40px' : '56px'), title: kind === 'col' ? '열 번호' : '줄 이름',
            oninput: e => { arr[k] = e.target.value; changed(); syncHeads(); } });
          if (kind !== 'col') return num;
          return h('span', { class: 'grid-col-pair' }, num,
            h('input', { value: g.columnNames[k] || '', placeholder: '이름', style: 'width:84px', title: '열 이름 (선택)',
              oninput: e => { g.columnNames[k] = e.target.value; changed(); syncHeads(); } }));
        }),
        h('button', { type: 'button', class: 'small-btn', onclick: () => {
          arr.push(kind === 'col' ? String(arr.length + 1) : String(arr.length + 2));
          if (kind === 'col') g.columnNames.length = arr.length;
          changed(); drawEditor();
        } }, '+'),
        arr.length > 1 ? h('button', { type: 'button', class: 'small-btn', title: '마지막 것을 빼기', onclick: () => {
          const last = arr.length - 1;
          const test = k => { const [r, c] = String(k).split('-').map(Number); return kind === 'col' ? c === last : r === last; };
          if (!gridGuard(g, test, `"${g.title || g.id}"의 마지막 ${kind === 'col' ? '열' : '줄'}`)) return;
          arr.pop();
          if (kind === 'col') { g.cells.forEach(row => row.splice(last)); g.columnNames.splice(last); } else g.cells.splice(last);
          changed(); drawEditor();
        } }, '−') : null);
      // 특기 이름 칸들
      const headCells = [];
      const syncHeads = () => headCells.forEach(f => f());
      const nameTable = h('table', { class: 'grid-table grid-names' },
        h('thead', {}, h('tr', {}, h('th', { class: 'corner' }), g.columns.map((_, c) => {
          const headText = () => [g.columns[c], g.columnNames[c]].filter(Boolean).join(' ');
          const th = h('th', { class: 'col-h' }, headText());
          headCells.push(() => { th.textContent = headText(); });
          return th;
        }))),
        h('tbody', {}, g.rows.map((_, r) => {
          const th = h('th', { class: 'row-h' }, g.rows[r]);
          headCells.push(() => { th.textContent = g.rows[r]; });
          return h('tr', {}, th, g.columns.map((__, c) => h('td', {},
            h('input', { value: ((g.cells[r] || [])[c]) || '', placeholder: '특기 이름', oninput: e => { setCell(r, c, e.target.value); changed(); } }))));
        })));
      const paste = h('textarea', { rows: 4, placeholder: '엑셀·구글 시트에서 특기 이름 표를 복사해서 붙여넣어요. (한 줄 = 한 줄, 칸은 탭으로 구분)' });
      const pasteBox = h('div', { class: 'paste-box', hidden: true }, paste, h('div', { class: 'row tight' },
        h('button', { type: 'button', class: 'primary', onclick: () => {
          const lines = paste.value.replace(/\r/g, '').split('\n').filter(l => l.trim());
          lines.slice(0, R).forEach((line, r) => line.split('\t').slice(0, C).forEach((v, c) => setCell(r, c, v.trim())));
          changed(); drawEditor();
        } }, '채우기'),
        h('button', { type: 'button', onclick: () => { pasteBox.hidden = true; } }, '취소')));

      return drag(h('section', { class: 'card tpl-section is-grid' + (g.folded ? ' is-folded' : '') }, head,
        h('p', { class: 'muted small', style: 'margin:0' },
          '체크한(습득한) 칸은 "습득 칸 달성치", 거기서 위아래·옆으로 멀어질수록 달성치가 올라가요. 열 사이의 빈 줄도 한 칸으로 세고, 시트에서 고른 열의 양쪽 빈 줄은 검게 칠해 무효화할 수 있어요.'),
        h('div', { class: 'row tight' },
          numIn('base', '습득 칸 달성치', 5), numIn('empty', '습득 없을 때', 12), numIn('cap', '최대 (0=제한 없음)', 12),
          numIn('rowCost', '위아래 한 칸', 1), numIn('colCost', '옆 열 한 칸', 1), numIn('gapCost', '빈 줄 한 칸', 1)),
        labelsEditor(g.columns, 'col'),
        labelsEditor(g.rows, 'row'),
        h('div', { class: 'row tight' }, h('div', { class: 'sub-title grow' }, '특기 이름'),
          h('button', { type: 'button', class: 'small-btn', onclick: () => { pasteBox.hidden = !pasteBox.hidden; } }, '표 붙여넣기')),
        pasteBox,
        h('div', { class: 'grid-wrap' }, nameTable),
        h('details', { class: 'more' }, h('summary', {}, '더보기'),
          idEditor(g, idOut, '수식에서 부르는 특기표 이름', v => renameGrid(g, v), takenGrids, () => gridGuard(g)),
          h('p', { class: 'muted small' }, `다른 칸의 수식에서 달성치("${g.id}", "특기 이름")으로 그 특기의 달성치를, 습득("${g.id}", "특기 이름")으로 습득 여부(1/0)를 쓸 수 있어요.`))),
        g, { ghost: () => head });
    }

    function sectionCard(s, i) {
      if (s.type === 'grid') return gridCard(s, i);
      const isList = s.type === 'list';
      const takenLists = () => new Set(allLists().filter(x => x !== s).map(x => x.id));
      const titleInput = h('input', { class: 'section-title', value: s.title || '', placeholder: isList ? '표 이름 (예: 기능)' : '묶음 이름 (예: 능력치)',
        oninput: e => {
          s.title = e.target.value;
          if (isList && fresh.has(s) && !manual.has(s)) { renameList(s, idFrom(e.target.value, takenLists())); if (listIdOut.el) listIdOut.el.value = s.id; }
          changed();
        } });
      if (focusTarget === s) requestAnimationFrame(() => { titleInput.focus(); titleInput.select(); focusTarget = null; });
      const listIdOut = { el: null };

      const head = h('div', { class: 'section-head' },
        h('span', { class: 'section-kind' }, isList ? '표' : '칸 묶음'),
        titleInput,
        iconBtn(s.folded ? '🙈' : '👁', s.folded ? '숨긴 묶음이에요. 눌러서 다시 보이게' : '이 묶음 전체를 시트에서 숨기기', () => {
          if (s.folded) delete s.folded; else s.folded = true; changed(); drawEditor();
        }, 'fold-btn' + (s.folded ? ' on' : '')),
        toggleBtn(s),
        isList ? null : iconBtn('⇲', '다른 칸 묶음과 합치기', () => {
          if (mergeOpen.has(s)) mergeOpen.delete(s); else mergeOpen.add(s);
          drawEditor();
        }, 'merge-btn' + (mergeOpen.has(s) ? ' on' : '')),
        iconBtn('×', '삭제', () => {
          if (isList ? !listGuard(s) : !s.fields.every(f => fieldGuard(f))) return;
          const n = isList ? s.columns.length : s.fields.length;
          if (n && !confirm(`"${s.title || '이름 없음'}"과 그 안의 ${n}개 항목을 지울까요?`)) return;
          draft.sections.splice(draft.sections.indexOf(s), 1); changed(); drawEditor();
        }, 'danger'));

      const opts = h('div', { class: 'row tight section-opts' },
        isList ? null : h('label', { class: 'mini-field' }, h('span', {}, '카드 안 격자의 가로 칸 수'),
          h('input', { type: 'number', min: 1, max: 12, step: 1, style: 'width:120px', value: colsOf(s),
            title: '카드 안을 가로 몇 칸으로 나눌지. 칸은 이 격자에 맞춰 옮기고 늘려요.',
            onchange: e => {
              const n = Math.floor(Number(e.target.value));
              if (n >= 1) { s.cols = Math.min(n, 12); changed(); }
            } })),
        h('span', { class: 'muted small' }, '카드 위치·크기는 시트 화면에서 끌어서 바꿔요.'),
        isList ? idEditor(s, listIdOut, '수식에서 부르는 표 이름', v => renameList(s, v), takenLists, () => listGuard(s)) : null);

      let body;
      if (isList) {
        body = [
          dataLinkPanel(s),
          h('div', { class: 'sub-title' }, '열'),
          A.dropList(h('div', { class: 'col-list' }, s.columns.map((c, j) => columnItem(c, s, j))), s.columns, s),
          h('button', { type: 'button', class: 'add-row', onclick: () => {
            const c = { label: '새 열', type: 'number' };
            c.id = idFrom('새열', new Set(s.columns.map(x => x.id)));
            fresh.add(c); s.columns.push(c); focusTarget = c; changed(); drawEditor();
          } }, '+ 열 추가'),
          defaultRowsEditor(s),
        ];
      } else {
        body = [
          h('div', { class: 'field-list' }, s.fields.map((f, j) => fieldItem(f, s, j)),
            s.fields.length ? null : h('p', { class: 'muted small empty-drop' }, '칸이 없어요. 시트 화면에서 다른 카드의 칸을 이 카드로 끌어다 놓을 수도 있어요.')),
          h('button', { type: 'button', class: 'add-row', onclick: () => {
            const f = { label: '새 칸', type: 'number' };
            f.id = idFrom('새칸', new Set(allFields().map(x => x.id)));
            fresh.add(f); s.fields.push(f); focusTarget = f; changed(); drawEditor();
          } }, '+ 칸 추가'),
        ];
      }
      return (h('section', { class: 'card tpl-section' + (isList ? ' is-list' : '') + (s.folded ? ' is-folded' : '') }, head,
        !isList && mergeOpen.has(s) ? mergePanel(s) : null,
        s.folded ? h('span', { class: 'fold-tag' }, '🙈 숨긴 묶음 — 시트에서는 접힌 채로 보이고(눌러서 펼치기), 읽기 모드·공유 화면에서는 안 보여요.') : null,
        !s.folded && s.toggle ? h('span', { class: 'fold-tag' }, '▸ 접어 둔 카드 — 시트에서는 제목만 보이고, 읽기 모드·공유 화면에서도 눌러서 펼쳐 볼 수 있어요.') : null,
        opts, body));
    }

    // ---------- 칸 묶음 합치기: 이 묶음의 칸을 모두 다른 묶음으로 옮기고, 빈 묶음은 없앰 ----------
    const mergeOpen = new WeakSet();
    function mergePanel(s) {
      const targets = draft.sections.filter(x => x !== s && x.type !== 'list' && x.type !== 'grid');
      if (!targets.length) {
        return h('div', { class: 'merge-panel' }, h('span', { class: 'muted small' }, '합칠 다른 칸 묶음이 없어요.'),
          h('button', { type: 'button', class: 'small-btn', onclick: () => { mergeOpen.delete(s); drawEditor(); } }, '닫기'));
      }
      const target = h('select', {}, targets.map(x => h('option', { value: String(draft.sections.indexOf(x)) }, x.title || '(이름 없음)')));
      const where = h('select', {}, h('option', { value: 'end' }, '맨 뒤에'), h('option', { value: 'start' }, '맨 앞에'));
      return h('div', { class: 'merge-panel' },
        h('span', {}, `이 묶음의 칸 ${s.fields.length}개를`), target, h('span', {}, '묶음의'), where,
        h('span', {}, '넣고, 이 묶음은 없애요.'),
        h('div', { class: 'row tight' },
          h('button', { type: 'button', class: 'primary small-btn', onclick: () => {
            const to = draft.sections[Number(target.value)];
            if (!to) return;
            const moving = s.fields.slice();
            // 숨긴 묶음에서 옮기는 칸은 계속 숨겨지도록 접힌 칸으로
            if (s.folded && !to.folded) moving.forEach(f => { f.folded = true; });
            to.fields = where.value === 'start' ? [...moving, ...to.fields] : [...to.fields, ...moving];
            draft.sections.splice(draft.sections.indexOf(s), 1);
            mergeOpen.delete(s);
            changed(); drawEditor();
            A.toast(`"${s.title || '이름 없음'}"의 칸 ${moving.length}개를 "${to.title || '이름 없음'}"에 합쳤어요. 마음에 안 들면 저장하기 전에 페이지를 나가면 돼요.`);
          } }, '합치기'),
          h('button', { type: 'button', class: 'small-btn', onclick: () => { mergeOpen.delete(s); drawEditor(); } }, '취소')));
    }

    // ---------- 글 틀 입력칸 ({칸이름}을 넣는 글) ----------
    function textTemplate(obj, prop, opt) {
      opt = opt || {};
      const el = h(opt.single ? 'input' : 'textarea', {
        class: 'tpl-text', rows: opt.rows || 3, value: obj[prop] ?? '', placeholder: opt.placeholder || '',
        spellcheck: false, autocomplete: 'off',
        oninput: e => { obj[prop] = e.target.value; changed(); },
      });
      el._bind = [obj, prop];
      const chips = h('div', { class: 'fx-group tpl-chips', hidden: true });
      chips.addEventListener('mousedown', e => e.preventDefault());
      const ins = t => {
        const s = el.selectionStart ?? el.value.length, e = el.selectionEnd ?? s;
        el.setRangeText(t, s, e, 'end');
        el.dispatchEvent(new Event('input'));
      };
      const list = () => (opt.list ? opt.list() : null);
      el.addEventListener('focus', () => {
        const l = list();
        chips.replaceChildren(...[
          h('span', { class: 'fx-title' }, '넣기'),
          chip('캐릭터 이름', '{#이름}', () => ins('{#이름}'), 'fn'),
          l ? l.columns.filter(c => c.id).map(c => chip(`${c.label || c.id} (같은 줄)`, `{${c.id}}`, () => ins(`{${c.id}}`), 'ref')) : [],
          allFields().filter(f => f.id).map(f => chip(labelOf(f), `{${f.id}}`, () => ins(`{${f.id}}`), l ? '' : 'ref')),
          h('span', { class: 'fx-tip' }, '{ } 안에는 수식도 쓸 수 있어요. 예: {근력 * 5}'),
        ].flat());
        chips.hidden = false;
      });
      el.addEventListener('blur', () => { chips.hidden = true; });
      return h('div', { class: 'tpl-text-wrap' }, el, chips);
    }

    // ---------- 코코포리아 연동 카드 ----------
    let ccPreview = null;
    function updateCcPreview() {
      if (!ccPreview) return;
      try {
        seedPreview();
        const out = E.toCcfolia(draft, { ...preview, name: '미리보기 캐릭터' });
        ccPreview.textContent = JSON.stringify(out.clipboard.data, null, 2);
        ccWarn.textContent = out.warnings.length ? '⚠ ' + out.warnings.join(' / ') : '';
      } catch (e) { ccPreview.textContent = '⚠ ' + e.message; }
    }
    const ccWarn = h('p', { class: 'fx-status err' });

    function ccfoliaCard() {
      const cc = draft.ccfolia;
      if (!cc) {
        ccPreview = null;
        return h('section', { class: 'card tpl-section ccfolia-card' },
          h('div', { class: 'section-head' }, h('span', { class: 'section-kind' }, '연동'), h('strong', {}, '코코포리아 연동')),
          h('p', { class: 'muted small' }, '지금은 자동 설정이에요. 최대값이 있는 숫자 칸은 상태(막대)로, 나머지 숫자·자동 계산 칸은 파라미터로 보내요. 채팅 팔레트나 메모까지 정하고 싶으면 직접 설정해 주세요.'),
          h('button', { type: 'button', onclick: () => { draft.ccfolia = E.autoCcfolia(draft); changed(); drawEditor(); } }, '직접 설정하기 (자동 설정에서 시작)'));
      }
      cc.status = cc.status || []; cc.params = cc.params || []; cc.listCommands = cc.listCommands || [];
      const lists = allLists();

      const statusRows = cc.status.map((st, i) => drag(h('div', { class: 'tpl-item' },
        h('div', { class: 'item-head' }, dragHandle(),
          h('input', { class: 'label-input', value: st.label || '', placeholder: '이름 (예: HP)', oninput: e => { st.label = e.target.value; changed(); } }),
          iconBtn('×', '삭제', () => { cc.status.splice(i, 1); changed(); drawEditor(); }, 'danger')),
        h('div', { class: 'row tight' },
          h('div', { class: 'mini-field grow' }, h('span', {}, '현재 값'), formulaInput(st, 'value', { expr: true }, '예: 체력')),
          h('div', { class: 'mini-field grow' }, h('span', {}, '최대 값'), formulaInput(st, 'max', { expr: true, optional: true, emptyText: '비워 두면 최대값 0으로 보내요.' }, '예: 최대체력')))), st));

      const paramRows = cc.params.map((p, i) => drag(h('div', { class: 'tpl-item' },
        h('div', { class: 'item-head' }, dragHandle(),
          h('input', { class: 'label-input', value: p.label || '', placeholder: '이름 (예: 근력)', oninput: e => { p.label = e.target.value; changed(); } }),
          iconBtn('×', '삭제', () => { cc.params.splice(i, 1); changed(); drawEditor(); }, 'danger')),
        textTemplate(p, 'value', { single: true, placeholder: '예: {근력}' })), p));

      const listCmdRows = cc.listCommands.map((lc, i) => h('div', { class: 'tpl-item' },
        h('div', { class: 'item-head' },
          h('select', { class: 'label-input', onchange: e => { lc.list = e.target.value; changed(); } },
            lists.map(l => h('option', { value: l.id, selected: l.id === lc.list }, `"${l.title || l.id}" 표의 줄마다`))),
          iconBtn('×', '삭제', () => { cc.listCommands.splice(i, 1); changed(); drawEditor(); }, 'danger')),
        textTemplate(lc, 'line', { single: true, placeholder: '예: CC<={합계} 【{이름}】', list: () => lists.find(l => l.id === lc.list) })));

      ccPreview = h('pre', { class: 'ccfolia-preview' });
      const sub = (title, desc) => h('div', { class: 'sub-title' }, title, desc ? h('span', { class: 'muted' }, ' ' + desc) : null);
      const card = h('section', { class: 'card tpl-section ccfolia-card' },
        h('div', { class: 'section-head' }, h('span', { class: 'section-kind' }, '연동'), h('strong', { class: 'grow' }, '코코포리아 연동'),
          h('button', { type: 'button', class: 'small-btn', onclick: () => {
            if (!confirm('지금 설정을 지우고 자동 설정으로 다시 채울까요?')) return;
            draft.ccfolia = E.autoCcfolia(draft); changed(); drawEditor();
          } }, '자동으로 다시 채우기'),
          h('button', { type: 'button', class: 'small-btn', onclick: () => {
            if (!confirm('직접 정한 연동 설정을 지우고 자동 설정으로 돌아갈까요?')) return;
            delete draft.ccfolia; changed(); drawEditor();
          } }, '자동 설정으로')),
        sub('상태', '(HP·MP처럼 막대로 보이는 값. 수식)'), A.dropList(h('div', { class: 'row-list' }, statusRows), cc.status, cc.status),
        h('button', { type: 'button', class: 'add-row', onclick: () => { cc.status.push({ label: '', value: '', max: '' }); changed(); drawEditor(); } }, '+ 상태 추가'),
        sub('파라미터', '(능력치처럼 "이름: 값"으로 보이는 것. {칸이름}으로 값 넣기)'), A.dropList(h('div', { class: 'row-list' }, paramRows), cc.params, cc.params),
        h('button', { type: 'button', class: 'add-row', onclick: () => { cc.params.push({ label: '', value: '' }); changed(); drawEditor(); } }, '+ 파라미터 추가'),
        sub('이니셔티브', '(행동 순서. 수식, 비워도 돼요)'),
        formulaInput(cc, 'initiative', { expr: true, optional: true, emptyText: '비워 두면 0으로 보내요.' }, '예: 민첩'),
        sub('채팅 팔레트', '(한 줄에 하나씩)'),
        textTemplate(cc, 'commands', { rows: 5, placeholder: '예:\n1d100<={근력 * 5} 근력 판정\n1d6+{근력보정} 대미지' }),
        lists.length ? [
          sub('표 줄마다 반복할 채팅 팔레트', '(기능 표의 줄 수만큼 자동으로 늘어나요)'), listCmdRows,
          h('button', { type: 'button', class: 'add-row', onclick: () => { cc.listCommands.push({ list: lists[0].id, line: '' }); changed(); drawEditor(); } }, '+ 반복 줄 추가'),
        ] : null,
        sub('메모'),
        textTemplate(cc, 'memo', { rows: 3, placeholder: '예: {직업} / {배경}' }),
        sub('채팅 글자색', '(#ff8800처럼 직접 적거나 {색상코드}처럼 칸에서 불러오기. 비우면 보내지 않아요)'),
        h('div', { class: 'row tight cc-color-row' },
          textTemplate(cc, 'color', { single: true, placeholder: '예: {색상코드} 또는 #ff8800' }),
          h('input', { type: 'color', title: '색을 골라서 넣기', value: /^#[0-9a-f]{6}$/i.test(cc.color || '') ? cc.color : '#7c5cff',
            oninput: e => { cc.color = e.target.value; changed(); }, onchange: () => drawEditor() })),
        h('details', { class: 'more' }, h('summary', {}, '보내질 데이터 미리보기 (미리보기 칸의 값 기준)'), ccWarn, ccPreview));
      return card;
    }

    // ---------- 줄마다 칸 묶음 수 ----------
    // ---------- 편집기 전체 그리기: 큰 시트 화면 + 옆에서 열리는 설정 창 ----------
    const editorCol = h('div', { class: 'tpl-editor' });
    const previewBody = h('div', { class: 'tpl-preview-body' });
    const fileLabel = h('span', { class: 'muted small' });
    const updateFileLabel = () => { fileLabel.textContent = `저장 파일: templates/${isNew ? fileKey(draft.rule) : key}.json`; };
    let sel = null; // { kind: 'meta' } | { kind: 'sec', s } | { kind: 'field', s, f }
    let mode = 'arrange'; // arrange: 배치 편집 / values: 값 넣어 보기
    const drawerTitle = h('strong', {});
    const drawer = h('aside', { class: 'cv-drawer', hidden: true },
      h('div', { class: 'cv-drawer-head' }, drawerTitle,
        h('button', { type: 'button', class: 'icon-btn', title: '설정 창 닫기', onclick: () => select(null) }, '×')),
      editorCol);
    function select(next) { sel = next; drawEditor(); }
    function findField(id) {
      for (const x of draft.sections) if (isFieldSec(x)) { const f = x.fields.find(y => y.id === id); if (f) return { s: x, f }; }
      return null;
    }

    function addSection(kind) {
      let x;
      if (kind === 'list') {
        x = { title: '새 표', type: 'list', id: idFrom('새표', new Set(allLists().map(y => y.id))), columns: [{ id: '이름', label: '이름', type: 'text' }], defaultRows: [] };
        fresh.add(x.columns[0]);
      } else if (kind === 'grid') {
        const rows = ['2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12'], columns = ['1', '2', '3', '4', '5', '6'];
        x = { title: '특기', type: 'grid', id: idFrom('특기', new Set(allGrids().map(y => y.id))), columns, rows, cells: rows.map(() => columns.map(() => '')) };
      } else x = { title: '새 묶음', cols: 3, fields: [] };
      const others = draft.sections.map(y => CV.posOf(y, CV.COLS));
      x.pos = CV.freeSpot(others, CV.COLS, kind === 'grid' || kind === 'list' ? CV.COLS : 6, 2);
      fresh.add(x); draft.sections.push(x); focusTarget = x;
      sel = { kind: 'sec', s: x };
      changed(); refreshPreview(true); drawEditor();
    }

    function metaCard() {
      return h('section', { class: 'card tpl-section' },
        h('div', { class: 'row tight' },
          h('label', { class: 'mini-field grow' }, h('span', {}, '룰 이름'),
            h('input', { value: draft.rule || '', placeholder: '예: 크툴루의 부름 7판', oninput: e => { draft.rule = e.target.value; updateFileLabel(); changed(); } })),
          h('label', { class: 'mini-field' }, h('span', {}, '강조색'),
            h('input', { type: 'color', value: /^#[0-9a-f]{6}$/i.test(draft.accent || '') ? draft.accent : '#7c5cff',
              oninput: e => { draft.accent = e.target.value; root.style.setProperty('--accent', draft.accent); changed(); } }))),
        h('label', { class: 'mini-field grow' }, h('span', {}, '설명 (룰 고를 때 보이는 한 줄)'),
          h('input', { value: draft.description || '', oninput: e => { draft.description = e.target.value; changed(); } })),
        fileLabel);
    }

    // ---------- 다른 템플릿(또는 이 템플릿)의 카드를 복사해 오기 ----------
    const importState = { src: null, chosen: new Set() };
    function sourceOf(k) {
      if (k === '__self') return draft;
      const t = templates[k];
      return t && !t.error ? CV.toCanvas(clone(t)) : null;
    }
    const kindName = x => (x.type === 'list' ? '표' : x.type === 'grid' ? '특기표' : '칸 묶음');
    function importPanel() {
      const keys = Object.keys(templates).filter(k => k !== key && !templates[k].error);
      if (!importState.src || (importState.src !== '__self' && !keys.includes(importState.src))) importState.src = keys[0] || '__self';
      const src = sourceOf(importState.src);
      const pick = h('select', { onchange: e => { importState.src = e.target.value; importState.chosen = new Set(); drawEditor(); } },
        keys.map(k => h('option', { value: k, selected: k === importState.src }, templates[k].rule || k)),
        h('option', { value: '__self', selected: importState.src === '__self' }, '이 템플릿 (카드 복제)'));
      const secs = (src && src.sections) || [];
      const items = secs.map((x, i) => {
        const n = x.type === 'list' ? (x.columns || []).length : x.type === 'grid' ? (x.rows || []).length : (x.fields || []).length;
        const unit = x.type === 'list' ? '열' : x.type === 'grid' ? '줄' : '칸';
        return h('label', { class: 'import-item' + (importState.chosen.has(i) ? ' on' : '') },
          h('input', { type: 'checkbox', checked: importState.chosen.has(i), onchange: e => {
            if (e.target.checked) importState.chosen.add(i); else importState.chosen.delete(i);
            e.target.closest('.import-item').classList.toggle('on', e.target.checked);
            go.disabled = !importState.chosen.size; go.textContent = `선택한 카드 ${importState.chosen.size}개 가져오기`;
          } }),
          h('span', { class: 'section-kind' }, kindName(x)),
          h('strong', {}, x.title || '(이름 없음)'),
          h('span', { class: 'muted small' }, `${n}${unit}${x.folded ? ' · 숨긴 묶음' : ''}`));
      });
      const go = h('button', { type: 'button', class: 'primary', disabled: !importState.chosen.size,
        onclick: () => importCards(src, [...importState.chosen].sort((a, b) => a - b)) }, `선택한 카드 ${importState.chosen.size}개 가져오기`);
      return h('section', { class: 'card tpl-section import-panel' },
        h('p', { class: 'muted small' }, '고른 카드를 칸·수식까지 통째로 복사해서 이 템플릿 맨 아래에 붙여요. 원래 템플릿은 바뀌지 않아요. 이름이 겹치는 칸은 이름 뒤에 숫자를 붙이고, 가져온 수식도 그 이름으로 맞춰요.'),
        h('label', { class: 'mini-field' }, h('span', {}, '가져올 템플릿'), pick),
        secs.length ? h('div', { class: 'import-list' }, items) : h('p', { class: 'muted' }, '카드가 없어요.'),
        h('div', { class: 'row tight' },
          secs.length ? h('button', { type: 'button', class: 'small-btn', onclick: () => {
            importState.chosen = importState.chosen.size === secs.length ? new Set() : new Set(secs.map((_, i) => i)); drawEditor();
          } }, importState.chosen.size === secs.length ? '모두 풀기' : '모두 고르기') : null,
          go));
    }
    function importCards(src, idxs) {
      if (!src || !idxs.length) return;
      const secs = idxs.map(i => clone(src.sections[i]));
      if (src === draft) for (const x of secs) x.title = (x.title || '') + ' 사본';
      // 이 템플릿에서 이미 쓰는 이름(칸·붙은 숫자·표·특기표)과 겹치면 새 이름으로
      const taken = new Set([...allFields().map(x => x.id), ...allLists().map(x => x.id), ...allGrids().map(x => x.id)]);
      const fmap = new Map(), tmap = new Map();
      const fresh1 = (id, map) => {
        if (!taken.has(id)) { taken.add(id); return id; }
        const n = idFrom(id, taken); taken.add(n); map.set(id, n); return n;
      };
      for (const x of secs) {
        if (x.type === 'list' || x.type === 'grid') x.id = fresh1(x.id, tmap);
        else for (const f of x.fields || []) { f.id = fresh1(f.id, fmap); for (const p of f.parts || []) p.id = fresh1(p.id, fmap); }
      }
      // 가져온 수식 안의 이름도 새 이름으로 (두 단계로 바꿔서 a→b, b→c가 엉키지 않게)
      const all = [...fmap, ...tmap];
      const remap = (str, localCols) => {
        if (typeof str !== 'string' || !all.length) return str;
        let out = str;
        all.forEach(([o], i) => {
          if (localCols && localCols.has(o) && fmap.has(o)) return; // 표 안에서는 같은 이름의 열이 먼저
          out = replaceIdent(out, o, `__가져옴${i}__`);
          if (tmap.has(o)) out = out.split(`"${o}"`).join(`"__가져옴${i}__"`).split(`'${o}'`).join(`'__가져옴${i}__'`);
        });
        all.forEach(([, n], i) => { out = out.split(`__가져옴${i}__`).join(n); });
        return out;
      };
      for (const x of secs) {
        if (x.type === 'list') {
          const cols = new Set((x.columns || []).map(c => c.id));
          for (const c of x.columns || []) if (c.formula !== undefined) c.formula = remap(c.formula, cols);
        } else if (x.type !== 'grid') {
          for (const f of x.fields || []) for (const y of [f, ...(f.parts || [])]) {
            if (y.formula !== undefined) y.formula = remap(y.formula);
            if (typeof y.max === 'string') y.max = remap(y.max);
          }
        }
      }
      // 자리: 지금 카드들 아래에, 가져온 카드끼리의 배치는 그대로
      const bottom = Math.max(0, ...draft.sections.map(x => { const q = CV.posOf(x, CV.COLS); return q.y + q.h; }));
      const top = Math.min(...secs.map(x => CV.posOf(x, CV.COLS).y));
      for (const x of secs) {
        const q = CV.posOf(x, CV.COLS);
        x.pos = { ...q, y: q.y - top + bottom };
        fresh.add(x);
        for (const y of [...(x.fields || []), ...(x.columns || [])]) { fresh.add(y); for (const p of y.parts || []) fresh.add(p); }
        draft.sections.push(x);
      }
      importState.chosen = new Set();
      changed(); refreshPreview(true);
      // 가져온 수식이 이 템플릿에 없는 칸을 쓰는지 확인
      const broken = [];
      try {
        seedPreview();
        const r = E.makeResolver(draft, preview);
        for (const x of secs) if (isFieldSec(x)) for (const f of x.fields) for (const y of [f, ...(f.parts || [])]) {
          if (y.formula === undefined && typeof y.max !== 'string') continue;
          try { if (y.formula !== undefined) r.get(y.id); if (typeof y.max === 'string') r.evaluate(y.max); }
          catch (e) { broken.push(labelOf(y)); }
        }
        for (const x of secs) if (x.type === 'list' && (preview.lists[x.id] || []).length) for (const c of x.columns || []) {
          if (c.formula === undefined) continue;
          try { r.cell(x.id, 0, c.id); } catch (e) { broken.push(`${x.title || x.id}의 ${c.label || c.id} 열`); }
        }
      } catch {}
      const renamed = fmap.size + tmap.size;
      A.toast(`카드 ${secs.length}개를 가져왔어요.` + (renamed ? ` 이름이 겹친 ${renamed}개는 새 이름을 붙였어요.` : '')
        + (broken.length ? ` ⚠ 수식에서 이 템플릿에 없는 칸을 쓰는 칸: ${broken.slice(0, 4).join(', ')}${broken.length > 4 ? ' 외' : ''} — 눌러서 수식을 고쳐 주세요.` : ''), !!broken.length);
      sel = { kind: 'sec', s: secs[0] };
      drawEditor();
      const el = previewBody.querySelector(`.sheet > .card[data-sec="${draft.sections.indexOf(secs[0])}"]`);
      if (el) el.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }

    function drawEditor() {
      if (activeInput) { activeInput = null; helper.remove(); previewBody.classList.remove('picking'); }
      // 지워진 것을 가리키고 있으면 정리
      if (sel && sel.kind === 'field') {
        const hit = findField(sel.f.id);
        sel = hit && hit.f === sel.f ? { kind: 'field', ...hit } : (draft.sections.includes(sel.s) ? { kind: 'sec', s: sel.s } : null);
      }
      if (sel && sel.kind === 'sec' && !draft.sections.includes(sel.s)) sel = null;
      drawer.hidden = !sel;
      root.classList.toggle('drawer-open', !!sel);
      fitToolbar();
      markSelected();
      if (!sel) { editorCol.replaceChildren(); return; }
      if (sel.kind === 'import') {
        drawerTitle.textContent = '다른 템플릿에서 카드 가져오기';
        editorCol.replaceChildren(importPanel());
      } else if (sel.kind === 'meta') {
        drawerTitle.textContent = '템플릿 설정';
        editorCol.replaceChildren(metaCard(), ccfoliaCard());
        updateFileLabel(); updateCcPreview();
      } else if (sel.kind === 'sec') {
        const x = sel.s;
        drawerTitle.textContent = x.type === 'list' ? '표 설정' : x.type === 'grid' ? '특기표 설정' : '칸 묶음 설정';
        editorCol.replaceChildren(sectionCard(x, draft.sections.indexOf(x)));
      } else {
        const { s: x, f } = sel;
        drawerTitle.textContent = '칸 설정';
        editorCol.replaceChildren(
          h('button', { type: 'button', class: 'small-btn back-sec', onclick: () => select({ kind: 'sec', s: x }) }, `← "${x.title || '이름 없는 묶음'}" 묶음 전체 보기`),
          h('div', { class: 'card tpl-section solo-field' }, fieldItem(f, x, x.fields.indexOf(f))));
      }
    }
    function markSelected() {
      for (const el of previewBody.querySelectorAll('.cv-sel')) el.classList.remove('cv-sel');
      if (!sel) return;
      if (sel.kind === 'sec') {
        const el = previewBody.querySelector(`.sheet > .card[data-sec="${draft.sections.indexOf(sel.s)}"]`);
        if (el) el.classList.add('cv-sel');
      } else if (sel.kind === 'field') {
        const el = [...previewBody.querySelectorAll('.field[data-ref]')].find(e => e.dataset.ref === sel.f.id);
        if (el) el.classList.add('cv-sel');
      }
    }

    // ---------- 시트 화면에서 끌어서 옮기기·크기 바꾸기 ----------
    function gridMetrics(g) {
      const cs = getComputedStyle(g), r = g.getBoundingClientRect();
      const nums = v => String(v).split(/\s+/).map(parseFloat).filter(n => !Number.isNaN(n));
      const m = /minmax\(\s*([\d.]+)px/.exec(cs.gridAutoRows || '');
      return {
        cols: nums(cs.gridTemplateColumns), rows: nums(cs.gridTemplateRows),
        cg: parseFloat(cs.columnGap) || 0, rg: parseFloat(cs.rowGap) || 0,
        left: r.left + (parseFloat(cs.paddingLeft) || 0) + (parseFloat(cs.borderLeftWidth) || 0),
        top: r.top + (parseFloat(cs.paddingTop) || 0) + (parseFloat(cs.borderTopWidth) || 0),
        minRow: m ? parseFloat(m[1]) : 40,
      };
    }
    function trackIndex(sizes, gap, pos, extra) {
      if (pos < 0) return 0;
      let acc = 0;
      for (let i = 0; i < sizes.length; i++) {
        if (pos < acc + sizes[i] + gap / 2) return i;
        acc += sizes[i] + gap;
      }
      return extra ? sizes.length + Math.floor((pos - acc) / (extra + gap)) : sizes.length - 1;
    }
    const cellAt = (m, x, y) => ({ x: trackIndex(m.cols, m.cg, x - m.left, 0), y: trackIndex(m.rows, m.rg, y - m.top, m.minRow) });
    const setPlace = (el, p, cols) => CV.placeStyle(el, p, cols);

    function decorate() {
      const sheetEl = previewBody.querySelector('.sheet');
      previewBody.classList.toggle('cv-arrange', mode === 'arrange');
      if (!sheetEl) return;
      sheetEl.classList.add('cv-force');
      if (mode === 'arrange') {
        for (const card of sheetEl.querySelectorAll(':scope > .card[data-sec]')) {
          const x = draft.sections[Number(card.dataset.sec)];
          if (!x) continue;
          card._obj = x;
          card.append(h('div', { class: 'cv-bar', title: '끌어서 옮기기 · 눌러서 설정' }, '⠿'),
            h('div', { class: 'cv-size', title: '끌어서 크기 바꾸기' }));
          card.addEventListener('pointerdown', e => press(e, card, 'card'));
        }
        for (const fe of sheetEl.querySelectorAll('.fields.canvas > .field[data-ref]')) {
          const hit = findField(fe.dataset.ref);
          if (!hit) continue;
          fe._obj = hit.f; fe._sec = hit.s;
          fe.append(h('div', { class: 'cv-size small', title: '끌어서 크기 바꾸기' }));
          fe.addEventListener('pointerdown', e => press(e, fe, 'field'));
        }
        for (const fe of sheetEl.querySelectorAll('.folded-fields .field[data-ref]')) {
          const hit = findField(fe.dataset.ref);
          if (hit) fe.addEventListener('pointerdown', e => { if (activeInput || e.button !== 0) return; e.stopPropagation(); select({ kind: 'field', ...hit }); });
        }
      }
      markSelected();
    }

    function press(e, el, kind) {
      if (activeInput || e.button !== 0) return; // 수식 입력 중이면 칸 이름 넣기가 먼저
      e.stopPropagation();
      const sizing = !!e.target.closest('.cv-size');
      const handle = sizing || !!e.target.closest('.cv-bar');
      const touch = e.pointerType === 'touch';
      const sx = e.clientX, sy = e.clientY;
      let timer = null, done = false;
      const immediate = !touch || handle;
      if (!immediate) timer = setTimeout(() => { timer = null; finish(); beginDrag(el, kind, sizing, sx, sy); }, 350);
      function move(ev) {
        if (Math.hypot(ev.clientX - sx, ev.clientY - sy) <= 6) return;
        if (immediate) { finish(); beginDrag(el, kind, sizing, sx, sy, ev); } else finish(); // 터치로 그냥 움직이면 스크롤
      }
      function up() {
        const tap = !done;
        finish();
        if (tap) {
          if (kind === 'card') select({ kind: 'sec', s: el._obj });
          else select({ kind: 'field', s: el._sec, f: el._obj });
        }
      }
      function finish() {
        done = true;
        if (timer) clearTimeout(timer);
        window.removeEventListener('pointermove', move);
        window.removeEventListener('pointerup', up);
        window.removeEventListener('pointercancel', finish);
      }
      window.addEventListener('pointermove', move);
      window.addEventListener('pointerup', up);
      window.addEventListener('pointercancel', finish);
    }

    function beginDrag(el, kind, sizing, x0, y0, firstEv) {
      const sheetEl = previewBody.querySelector('.sheet');
      const isField = kind === 'field';
      const obj = el._obj, fromSec = isField ? el._sec : null;
      const srcGrid = el.parentElement;
      const cols0 = isField ? colsOf(fromSec) : CV.COLS;
      const orig = CV.posOf(obj, cols0);
      const start = cellAt(gridMetrics(srcGrid), x0, y0);
      const ph = h('div', { class: 'cv-ph' + (isField ? ' small' : '') });
      el.classList.add('cv-dragging');
      document.body.classList.add('dragging');
      if (!sizing) el.style.pointerEvents = 'none';
      let target = null, shownGrid = null, lastX = x0, lastY = y0, raf = 0;
      const blockTouch = ev => ev.preventDefault();
      document.addEventListener('touchmove', blockTouch, { passive: false });

      const itemsOf = g => (g === sheetEl
        ? [...g.querySelectorAll(':scope > .card[data-sec]')]
        : [...g.querySelectorAll(':scope > .field[data-ref]')]).filter(c => c !== el && c._obj);
      const colsFor = g => (g === sheetEl ? CV.COLS : colsOf(g.closest('.card')._obj));
      function restore(g) { if (g) for (const c of itemsOf(g)) setPlace(c, CV.posOf(c._obj, colsFor(g)), colsFor(g)); }
      function layoutPreview(g, p) {
        const cols = colsFor(g);
        const items = itemsOf(g).map(c => ({ c, ...CV.posOf(c._obj, cols) }));
        const fixed = { ...p };
        CV.settle([fixed, ...items], cols, fixed);
        for (const it of items) setPlace(it.c, it, cols);
        return items;
      }
      function place() {
        if (!sizing) el.style.transform = `translate(${lastX - x0}px, ${lastY - y0}px)`;
        let g = srcGrid;
        if (isField && !sizing) {
          const under = document.elementFromPoint(lastX, lastY);
          const card = under && under.closest('.sheet > .card[data-sec]');
          const fg = card && card.querySelector(':scope > .fields.canvas');
          g = fg && card._obj && isFieldSec(card._obj) ? fg : null;
        }
        if (g !== shownGrid) { restore(shownGrid); shownGrid = g; }
        if (!g) { ph.remove(); target = null; return; }
        const cols = colsFor(g);
        const c = cellAt(gridMetrics(g), lastX, lastY);
        let p;
        if (sizing) p = { x: orig.x, y: orig.y, w: c.x - orig.x + 1, h: c.y - orig.y + 1 };
        else if (g === srcGrid) p = { w: orig.w, h: orig.h, x: orig.x + c.x - start.x, y: orig.y + c.y - start.y };
        else p = { w: Math.min(orig.w, cols), h: orig.h, x: c.x, y: c.y };
        p.w = Math.max(1, Math.min(cols, p.w)); p.h = Math.max(1, Math.min(40, p.h));
        p.x = Math.max(0, Math.min(cols - p.w, p.x)); p.y = Math.max(0, p.y);
        if (ph.parentElement !== g) g.append(ph);
        setPlace(ph, p, cols);
        if (sizing) setPlace(el, p, cols);
        layoutPreview(g, p);
        target = { g, p, cols };
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
        window.removeEventListener('pointercancel', cancel);
        document.removeEventListener('touchmove', blockTouch);
        document.body.classList.remove('dragging');
        ph.remove();
        if (!target) { refreshPreview(true); return; }
        // 확정: 옮긴 것은 그 자리에, 나머지는 비켜서 위로 붙임
        const { g, p, cols } = target;
        const toSec = g === sheetEl ? null : g.closest('.card')._obj;
        if (isField && toSec !== fromSec) {
          fromSec.fields.splice(fromSec.fields.indexOf(obj), 1);
          toSec.fields.push(obj);
          if (sel && sel.kind === 'field' && sel.f === obj) sel = { kind: 'field', s: toSec, f: obj };
        }
        const siblings = itemsOf(g).map(c => c._obj);
        const fixed = { ...p };
        const items = siblings.map(o => ({ o, ...CV.posOf(o, cols) }));
        CV.settle([fixed, ...items], cols, fixed);
        for (const it of items) it.o.pos = { x: it.x, y: it.y, w: it.w, h: it.h };
        obj.pos = { x: p.x, y: p.y, w: p.w, h: p.h };
        changed(); refreshPreview(true);
        if (sel) drawEditor();
      }
      function cancel() { target = null; end(); }
      window.addEventListener('pointermove', onMove);
      window.addEventListener('pointerup', end);
      window.addEventListener('pointercancel', cancel);
      if (firstEv) onMove(firstEv); else place();
    }

    // ---------- 점검 & 저장 ----------
    function validate() {
      const errors = [], warnings = [];
      if (!String(draft.rule || '').trim()) errors.push('룰 이름을 입력해 주세요.');
      const seen = new Set();
      for (const f of allFields()) {
        const name = f.label || f.id || '(이름 없음)';
        if (!validId(f.id)) errors.push(`"${name}" 칸의 수식 이름이 올바르지 않아요. (더보기에서 확인)`);
        else if (seen.has(f.id)) errors.push(`수식 이름 '${f.id}'가 두 번 쓰였어요.`);
        seen.add(f.id);
        if (f.type === 'select' && !(f.options || []).length) warnings.push(`"${name}": 선택지가 비어 있어요.`);
      }
      const seenL = new Set();
      for (const l of allLists()) {
        if (!validId(l.id)) errors.push(`"${l.title}" 표의 수식 이름이 올바르지 않아요.`);
        else if (seenL.has(l.id)) errors.push(`표 이름 '${l.id}'가 두 번 쓰였어요.`);
        seenL.add(l.id);
        const seenC = new Set();
        for (const c of l.columns) {
          if (!validId(c.id)) errors.push(`"${l.title}" 표의 "${c.label}" 열 이름이 올바르지 않아요.`);
          else if (seenC.has(c.id)) errors.push(`"${l.title}" 표에 '${c.id}' 열이 두 번 있어요.`);
          seenC.add(c.id);
        }
      }
      if (errors.length) return { errors, warnings };
      seedPreview();
      const r = E.makeResolver(draft, preview);
      for (const f of allFields()) {
        const name = f.label || f.id;
        try {
          if (f.formula !== undefined) { if (!f.formula.trim()) throw new Error('수식이 비어 있어요'); r.get(f.id); }
          if (typeof f.max === 'string' && f.max.trim()) r.evaluate(f.max);
        } catch (e) { warnings.push(`"${name}": ${e.message}`); }
      }
      for (const l of allLists()) for (const c of l.columns) {
        if (c.formula === undefined) continue;
        try {
          if (!c.formula.trim()) throw new Error('수식이 비어 있어요');
          E.parse(c.formula);
          if ((preview.lists[l.id] || []).length) r.cell(l.id, 0, c.id);
        } catch (e) { warnings.push(`"${l.title}" 표의 "${c.label}": ${e.message}`); }
      }
      if (draft.ccfolia) {
        try { for (const w of E.toCcfolia(draft, { ...preview, name: '' }).warnings) warnings.push('코코포리아 ' + w); }
        catch (e) { warnings.push('코코포리아: ' + e.message); }
      }
      return { errors, warnings };
    }
    function cleaned() {
      const t = clone(draft);
      t.rule = String(t.rule).trim();
      if (!t.description) delete t.description;
      if (!(t.cardRows || []).length) delete t.cardRows;
      for (const s of t.sections) {
        if (!s.wide) delete s.wide;
        for (const x of [...(s.fields || []), ...(s.columns || [])]) {
          if (x.formula !== undefined) { delete x.type; delete x.default; delete x.max; delete x.options; }
          if (x.max !== undefined && !String(x.max).trim()) delete x.max;
          if (!x.wide) delete x.wide;
          for (const pt of x.parts || []) {
            if (pt.formula !== undefined) { delete pt.type; delete pt.default; } else pt.type = 'number';
            if (!pt.label) delete pt.label;
          }
          if (x.parts && !x.parts.length) delete x.parts;
        }
        if (s.type === 'grid') {
          s.columnNames = Array.from({ length: (s.columns || []).length }, (_, k) => (s.columnNames || [])[k] || '');
          if (!s.columnNames.some(Boolean)) delete s.columnNames;
        }
        if (s.type === 'list') {
          s.defaultRows = (s.defaultRows || []).filter(row => Object.keys(row).length);
          if (!s.defaultRows.length) delete s.defaultRows;
        }
      }
      return t;
    }
    async function save(btn, goList) {
      const { errors, warnings } = validate();
      if (errors.length) return A.toast(errors[0], true);
      if (warnings.length && !confirm('확인이 필요한 칸이 있어요.\n\n' + warnings.slice(0, 6).join('\n') +
        (warnings.length > 6 ? `\n… 외 ${warnings.length - 6}개` : '') + '\n\n그래도 저장할까요?')) return;
      let saveKey = key;
      if (isNew) {
        saveKey = fileKey(draft.rule);
        if (templates[saveKey]) return A.toast('같은 이름의 템플릿이 이미 있어요. 룰 이름을 바꿔 주세요.', true);
      }
      btn.disabled = true;
      try {
        sha = await A.writeJson(`templates/${saveKey}.json`, cleaned(), sha, `템플릿 ${isNew ? '생성' : '수정'}: ${draft.rule}`);
        A.resetTemplates();
        A.setDirty(false);
        A.toast('템플릿을 저장했어요. 이 틀을 쓰는 캐릭터 시트에 바로 반영돼요.');
        if (!isNew) {
          A.refreshShares({ templateKey: saveKey }).then(({ ok, fail }) => {
            if (fail) A.toast(`공유본 ${fail}개를 새로 올리지 못했어요. 해당 캐릭터를 열어 저장하면 다시 올라가요.`, true);
            else if (ok) A.toast(`템플릿을 저장했어요. 공유 중인 시트 ${ok}개도 새로 반영했어요.`);
          }).catch(() => {});
        }
        if (goList) A.go('/templates');
        else if (isNew) A.go(`/template/${encodeURIComponent(saveKey)}`);
      } catch (e) { A.toast(e.message, true); btn.disabled = false; }
    }
    async function remove(btn) {
      btn.disabled = true;
      try {
        const using = await A.charactersUsing(key);
        if (using.length) {
          A.toast(`이 룰 시트를 가진 캐릭터가 ${using.length}명 있어서 삭제할 수 없어요. (${using.slice(0, 3).map(s => s.name).join(', ')}${using.length > 3 ? ' …' : ''})`, true);
          btn.disabled = false; return;
        }
        if (!confirm(`"${draft.rule}" 템플릿을 삭제할까요?`)) { btn.disabled = false; return; }
        await A.deleteFile(`templates/${key}.json`, sha, `템플릿 삭제: ${draft.rule}`);
        A.resetTemplates(); A.setDirty(false);
        A.toast('삭제했어요.');
        A.go('/templates');
      } catch (e) { A.toast(e.message, true); btn.disabled = false; }
    }

    // ---------- 화면 배치 ----------
    const saveBtn = h('button', { id: 'save-btn', class: 'primary', onclick: e => save(e.target) }, '저장');
    const modeBtns = h('div', { class: 'mode-switch', role: 'group', 'aria-label': '편집 방식' });
    function drawMode() {
      modeBtns.replaceChildren(
        h('button', { type: 'button', class: mode === 'arrange' ? 'on' : '', title: '카드와 칸을 끌어서 옮기고, 오른쪽 아래 모서리를 끌어서 크기를 바꿔요. 누르면 설정 창이 열려요.',
          onclick: () => { mode = 'arrange'; drawMode(); refreshPreview(true); } }, '배치 편집'),
        h('button', { type: 'button', class: mode === 'values' ? 'on' : '', title: '실제 시트처럼 값을 넣어 보며 계산을 확인해요. 넣은 값은 저장되지 않아요.',
          onclick: () => { mode = 'values'; drawMode(); refreshPreview(true); } }, '값 넣어 보기'));
    }
    drawMode();
    const toolbar = h('div', { class: 'cv-toolbar' },
      saveBtn,
      h('button', { class: 'save-list-btn', onclick: e => save(e.target, true) }, '저장하고 목록으로'),
      h('span', { class: 'cv-sep' }),
      modeBtns,
      h('button', { type: 'button', onclick: () => addSection('fields') }, '+ 칸 묶음'),
      h('button', { type: 'button', onclick: () => addSection('list'), title: '기능·소지품처럼 줄이 늘어나는 표' }, '+ 표'),
      h('button', { type: 'button', onclick: () => addSection('grid'), title: '격자 모양 특기표' }, '+ 특기표'),
      h('button', { type: 'button', onclick: () => select({ kind: 'import' }), title: '다른 템플릿의 카드를 통째로 복사해 와요' }, '⤵ 카드 가져오기'),
      h('button', { type: 'button', onclick: () => select({ kind: 'meta' }) }, '⚙ 템플릿 설정 · 코코포리아'),
      h('span', { class: 'muted small cv-tip' }, '카드·칸을 끌어서 옮기고 ◢ 모서리로 크기를 바꿔요. 휴대폰은 꾹 눌러서 옮겨요.'));
    const stage = h('div', { class: 'cv-stage' }, previewBody);
    // 설정 창이 열려 있으면 위쪽 도구 줄이 창에 가려지지 않게 오른쪽을 비움
    function fitToolbar() {
      toolbar.style.marginRight = '';
      if (drawer.hidden || !toolbar.isConnected) return;
      const over = toolbar.getBoundingClientRect().right - drawer.getBoundingClientRect().left + 10;
      if (over > 0) toolbar.style.marginRight = Math.min(over, toolbar.getBoundingClientRect().width - 160) + 'px';
    }
    const onResize = () => { if (!toolbar.isConnected) window.removeEventListener('resize', onResize); else fitToolbar(); };
    window.addEventListener('resize', onResize);
    const root = h('div', { class: 'tpl-layout canvas-editor' }, toolbar, stage, drawer);
    root.style.setProperty('--accent', draft.accent || '');
    stage.addEventListener('pointerdown', e => {
      if (activeInput || mode !== 'arrange' || e.target.closest('.card')) return;
      if (sel && sel.kind !== 'meta' && sel.kind !== 'import') select(null); // 빈 곳을 누르면 설정 창 닫기
    });

    // 엑셀처럼: 수식 입력 중에 미리보기의 칸을 클릭하면 그 칸 이름이 들어감
    previewBody.addEventListener('mousedown', e => {
      if (!activeInput) return;
      const cell = e.target.closest('[data-ref]');
      if (!cell) return;
      e.preventDefault();
      const ref = cell.dataset.ref, list = cell.dataset.list, ctx = activeInput._ctx;
      if (list) insert(ctx.list && ctx.list.id === list ? ref : `목록합계("${list}", "${ref}")`);
      else insert(ref);
    }, true);

    A.show(
      h('div', { class: 'sheet-head' },
        h('div', { class: 'title-row' },
          h('a', { class: 'back-link', href: '#/templates', title: '시트 템플릿 목록으로' }, '← 시트 템플릿'),
          h('h1', { class: 'page-title' }, isNew ? '새 템플릿' : '템플릿 편집')),
        h('div', { class: 'row' },
          isNew ? null : h('a', { class: 'button-link', href: `#/template/new?from=${encodeURIComponent(key)}` }, '복제'),
          isNew ? null : h('button', { class: 'danger ghost', onclick: e => remove(e.target) }, '삭제'))),
      root);
    A.setDirty(false);
    if (isNew) { saveBtn.disabled = false; saveBtn.textContent = '저장'; sel = { kind: 'meta' }; }
    normalizeLayout();
    drawEditor();
    refreshPreview(true);
  }
})();
