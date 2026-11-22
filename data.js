/*
 * data.js — 데이터 보관함
 * 룰마다 여러 시트에서 같이 쓰는 데이터(스킬, 아이템, 주문 등)를 모아 두는 곳.
 * datasets/<키>.json = { name, rule, fields: [{ id, label, type }], entries: [{ <id>: 값 }] }
 * fields[0](첫 번째 항목)이 "찾기 이름"이 되고, 룰 틀의 표에서 이 이름으로 데이터를 불러옴.
 */
(function () {
  'use strict';
  const A = window.SheetApp, E = window.SheetEngine, h = A.h;
  const TYPES = [['text', '글자'], ['number', '숫자'], ['textarea', '긴 글'], ['select', '선택 목록']];
  const clone = o => JSON.parse(JSON.stringify(o));
  const validId = id => /^[\p{L}_][\p{L}\p{N}_]*$/u.test(id || '');
  function idFrom(label, taken) {
    let base = String(label || '').replace(/[^\p{L}\p{N}_]/gu, '') || '항목';
    if (/^\p{N}/u.test(base)) base = '_' + base;
    let id = base, n = 2;
    while (taken.has(id)) id = base + n++;
    return id;
  }

  A.routes.push({ match: hash => hash === '/data', show: showDataList });
  A.routes.push({ match: hash => hash.startsWith('/data/'), show: hash => showDataEditor(decodeURIComponent(hash.slice(6))) });

  // 이 데이터를 연결해서 쓰는 룰 틀들 → [{ key, rule, columns: [{ from, label }] }]
  function templatesUsing(templates, dsKey) {
    return Object.entries(templates).flatMap(([k, t]) => (t.sections || [])
      .filter(s => s.type === 'list' && E.sourceKeys(s).includes(dsKey))
      .map(s => ({ key: k, rule: t.rule || k, table: s.title || s.id, from: (s.columns || []).filter(c => c.from).map(c => c.from),
        partners: E.sourceKeys(s).filter(x => x !== dsKey) })));
  }

  // ================= 목록 =================
  async function showDataList() {
    A.loading();
    try {
      const [datasets, templates, order] = await Promise.all([A.loadDatasets(true), A.loadTemplates(), A.loadOrder()]);
      const prefs = A.listPrefs('datasets', 'rule');
      const SORTS = [['custom', '직접 배치'], ['name', '이름 가나다순'], ['rule', '룰 순서'], ['count', '항목 많은 순']];
      if (!SORTS.some(x => x[0] === prefs.sort)) prefs.sort = 'rule';
      const tplOrder = A.applyOrder(Object.keys(templates), order.templates);
      const customKeys = () => A.applyOrder(Object.keys(datasets), order.datasets);
      const groupOf = d => (d.rule ? ((templates[d.rule] && templates[d.rule].rule) || d.rule) : '공용 (룰 없음)');
      const groupRank = d => { const i = d.rule ? tplOrder.indexOf(d.rule) : -1; return d.rule ? (i < 0 ? 9998 : i) : 9999; };
      function sortedKeys() {
        const keys = customKeys();
        if (prefs.sort === 'name') return keys.sort((a, b) => A.byName(datasets[a].name || a, datasets[b].name || b));
        if (prefs.sort === 'count') return keys.sort((a, b) => (datasets[b].entries || []).length - (datasets[a].entries || []).length);
        if (prefs.sort === 'rule') return keys.sort((a, b) => groupRank(datasets[a]) - groupRank(datasets[b]) || A.byName(groupOf(datasets[a]), groupOf(datasets[b])));
        return keys;
      }
      const bar = h('div', { class: 'filter-bar' });
      const body = h('div', {});
      let lists = [];
      async function dropped() {
        // 룰 순서로 묶어 볼 때는 묶음 안에서만 옮길 수 있고, 묶음 순서는 그대로
        const visible = lists.flatMap(l => [...l.children].map(c => c.dataset.key).filter(Boolean));
        const next = prefs.sort === 'custom' || prefs.sort === 'rule' ? A.mergeVisible(customKeys(), visible) : visible;
        order.datasets = next;
        if (prefs.sort !== 'custom' && prefs.sort !== 'rule') { prefs.sort = 'custom'; A.saveListPrefs('datasets', prefs); A.toast('정렬을 "직접 배치"로 바꿨어요.'); }
        draw();
        try { await A.saveOrder('datasets', next); } catch (e) { A.toast('순서를 저장하지 못했어요: ' + e.message, true); }
      }
      function card(k) {
        const d = datasets[k];
        const c = h('a', { class: 'sheet-card', 'data-key': k, href: `#/data/${encodeURIComponent(k)}` },
          h('strong', {}, d.name || k),
          h('span', { class: 'muted' }, d.error || `${prefs.sort === 'rule' ? '' : groupOf(d) + ' · '}${(d.entries || []).length}개 · ${(d.fields || []).map(f => f.label || f.id).join(' / ')}`));
        const t = d.rule && templates[d.rule];
        c.style.setProperty('--accent', (t && t.accent) || '');
        return c;
      }
      const listEl = keys => {
        const el = h('div', { class: `sheet-list size-${prefs.size}` }, keys.map(card));
        lists.push(el);
        return A.sortableCards(el, dropped);
      };
      function draw() {
        bar.replaceChildren(...A.listTools(prefs, SORTS, () => { A.saveListPrefs('datasets', prefs); draw(); }));
        const keys = sortedKeys();
        lists = [];
        if (!keys.length) return body.replaceChildren(h('p', { class: 'muted' }, '아직 데이터가 없어요.'));
        if (prefs.sort === 'rule') {
          const groups = new Map();
          for (const k of keys) { const g = groupOf(datasets[k]); if (!groups.has(g)) groups.set(g, []); groups.get(g).push(k); }
          body.replaceChildren(...[...groups.entries()].map(([g, ks]) => h('section', { class: 'data-group' }, h('h2', {}, g), listEl(ks))));
        } else body.replaceChildren(listEl(keys));
      }
      draw();
      A.show(
        h('div', { class: 'toolbar' }, h('h1', {}, '데이터'),
          h('button', { class: 'primary', onclick: () => A.go('/data/new') }, '+ 새 데이터')),
        h('p', { class: 'muted' }, '스킬·아이템·주문처럼 여러 캐릭터가 같이 쓰는 데이터를 룰별로 모아 두는 곳이에요. 템플릿의 표에 연결하면 이름만 입력해도 나머지 내용이 채워져요. 카드를 꾹 누르면 순서를 바꿀 수 있어요.'),
        bar, body);
    } catch (e) { A.failed(e); }
  }

  // ================= 편집 =================
  async function showDataEditor(param) {
    A.loading();
    let datasets, templates, draft, key = null, sha = null;
    const isNew = param === 'new';
    try {
      [datasets, templates] = await Promise.all([A.loadDatasets(true), A.loadTemplates(true)]);
      if (isNew) {
        draft = { name: '', rule: '', fields: [
          { id: '이름', label: '이름', type: 'text' },
          { id: '효과', label: '효과', type: 'textarea' },
        ], entries: [] };
      } else {
        key = param;
        const r = await A.readJson(`datasets/${key}.json`);
        draft = r.data; sha = r.sha;
      }
    } catch (e) { return A.failed(e); }
    draft.fields = draft.fields || [];
    draft.entries = draft.entries || [];
    const usedBy = isNew ? [] : templatesUsing(templates, key);
    const fresh = new WeakSet(), manual = new WeakSet();
    if (isNew) draft.fields.forEach(f => fresh.add(f));
    const changed = () => A.setDirty(true);

    const fieldUse = f => usedBy.filter(u => u.from.includes(f.id));
    function fieldGuard(f, what) {
      const u = fieldUse(f);
      if (!u.length) return true;
      A.toast(`"${f.label || f.id}" 항목은 ${u.map(x => `${x.rule}의 "${x.table}" 표`).join(', ')}에 연결돼 있어서 ${what} 수 없어요. 템플릿에서 연결을 먼저 풀어 주세요.`, true);
      return false;
    }
    function renameField(f, newId) {
      const old = f.id;
      if (old === newId) return;
      for (const e of draft.entries) if (old in e) { e[newId] = e[old]; delete e[old]; }
      f.id = newId;
    }

    // ---------- 선택 목록 항목 ----------
    // 이미 적힌 값들 (선택지로 바꿀 때 그대로 쓸 수 있게)
    const distinctValues = f => [...new Set(draft.entries.map(e => String(e[f.id] ?? '').trim()).filter(Boolean))];
    function optionsEditor(f) {
      f.options = f.options || [];
      const missing = () => distinctValues(f).filter(v => !f.options.includes(v));
      const note = h('span', { class: 'muted small' });
      const addBtn = h('button', { type: 'button', class: 'small-btn', onclick: () => {
        f.options = [...f.options, ...missing()];
        input.value = f.options.join(', ');
        changed(); drawTable(); updateNote();
      } }, '적힌 값을 선택지에 추가');
      const updateNote = () => {
        const m = missing();
        note.textContent = m.length ? `선택지에 없는 값 ${m.length}개가 적혀 있어요: ${m.slice(0, 4).join(', ')}${m.length > 4 ? ' …' : ''}` : '';
        addBtn.hidden = !m.length;
      };
      let timer;
      const input = h('input', { value: f.options.join(', '), placeholder: '예: 주동작, 보조동작, 상시', oninput: e => {
        f.options = e.target.value.split(',').map(x => x.trim()).filter(Boolean);
        changed(); updateNote();
        clearTimeout(timer); timer = setTimeout(drawTable, 300);
      } });
      updateNote();
      return h('div', { class: 'options-editor' },
        h('label', { class: 'mini-field' }, h('span', {}, '선택지 (쉼표로 구분)'), input),
        h('div', { class: 'row tight' }, note, addBtn));
    }

    // ---------- 항목 구성 ----------
    const schemaCard = h('section', { class: 'card tpl-section' });
    function drawSchema() {
      const taken = f => new Set(draft.fields.filter(x => x !== f).map(x => x.id));
      schemaCard.replaceChildren(
        h('div', { class: 'section-head' }, h('span', { class: 'section-kind' }, '항목'), h('strong', { class: 'grow' }, '항목 구성')),
        h('p', { class: 'muted small', style: 'margin:0' }, '데이터마다 들어갈 칸을 정해요. 예: 이름 / 사용 횟수 / 효과 / 해설, 또는 이름 / 개요'),
        A.dropList(h('div', { class: 'row-list' }, draft.fields.map((f, i) => {
          const idInput = h('input', { class: 'id-input', value: f.id, spellcheck: false });
          idInput.addEventListener('change', () => {
            const v = idInput.value.trim();
            if (!validId(v)) { A.toast('이름은 띄어쓰기·기호 없이 글자, 숫자, _만 쓸 수 있어요.', true); idInput.value = f.id; return; }
            if (v !== f.id && taken(f).has(v)) { A.toast(`'${v}'는 이미 있는 이름이에요.`, true); idInput.value = f.id; return; }
            if (v !== f.id && !fieldGuard(f, '이름을 바꿀')) { idInput.value = f.id; return; }
            manual.add(f); renameField(f, v); changed(); drawTable();
          });
          return A.draggable(h('div', { class: 'tpl-item' + (i === 0 ? ' is-formula' : '') },
            i === 0 ? h('span', { class: 'link-tag' }, '🔎 찾기 이름 — 템플릿에서 이 항목으로 데이터를 찾아요') : null,
            h('div', { class: 'item-head' }, A.dragHandle(),
              h('input', { class: 'label-input', value: f.label || '', placeholder: '항목 이름', oninput: e => {
                f.label = e.target.value;
                if (fresh.has(f) && !manual.has(f)) { renameField(f, idFrom(e.target.value, taken(f))); idInput.value = f.id; }
                changed(); drawTableHead();
              } }),
              h('select', { class: 'kind', onchange: e => {
                f.type = e.target.value;
                if (f.type === 'select') { if (!f.options || !f.options.length) f.options = distinctValues(f); } else delete f.options;
                changed(); drawSchema(); drawTable();
              } },
                TYPES.map(([v, l]) => h('option', { value: v, selected: (f.type || 'text') === v }, l))),
              h('button', { type: 'button', class: 'icon-btn danger', title: '항목 삭제', onclick: () => {
                if (draft.fields.length === 1) return A.toast('항목은 하나 이상 있어야 해요.', true);
                if (!fieldGuard(f, '지울')) return;
                const filled = draft.entries.filter(e => e[f.id] !== undefined && e[f.id] !== '').length;
                if (filled && !confirm(`"${f.label}" 항목에 적힌 값 ${filled}개도 같이 지워져요. 지울까요?`)) return;
                for (const e of draft.entries) delete e[f.id];
                draft.fields.splice(i, 1); changed(); drawSchema(); drawTable();
              } }, '×')),
            f.type === 'select' ? optionsEditor(f) : null,
            h('details', { class: 'more' }, h('summary', {}, '더보기'),
              h('label', { class: 'mini-field' }, h('span', {}, '저장용 이름 (템플릿에서 연결할 때 쓰여요)'), idInput))), f, {
            // 룰 틀에 연결된 데이터는 찾기 이름(첫 번째 항목)을 바꿀 수 없음
            canMove: ({ oldIdx, newIdx }) => {
              if ((oldIdx === 0 || newIdx === 0) && usedBy.length) { A.toast('템플릿에 연결된 데이터라 찾기 이름(첫 번째 항목)은 바꿀 수 없어요.', true); return false; }
              return true;
            },
            onMoved: () => { changed(); drawSchema(); drawTable(); },
          });
        })), draft.fields, 'ds-fields'),
        h('button', { type: 'button', class: 'add-row', onclick: () => {
          const f = { label: '새 항목', type: 'text' };
          f.id = idFrom('새항목', new Set(draft.fields.map(x => x.id)));
          fresh.add(f); draft.fields.push(f); changed(); drawSchema(); drawTable();
        } }, '+ 항목 추가'));
    }

    // ---------- 데이터 표 ----------
    const search = h('input', { type: 'search', placeholder: '찾기', class: 'search' });
    const countLabel = h('span', { class: 'muted small' });
    const thead = h('thead');
    const tbody = h('tbody');
    const pasteBox = h('div', { class: 'paste-box', hidden: true });
    function drawTableHead() {
      thead.replaceChildren(h('tr', {}, draft.fields.map(f => h('th', {}, f.label || f.id)), h('th')));
    }
    function rowEl(e) {
      const tr = h('tr', {},
        draft.fields.map((f, i) => {
          const t = f.type || 'text';
          const set = v => { if (v === '' || v == null) delete e[f.id]; else e[f.id] = v; changed(); };
          const cur = String(e[f.id] ?? '');
          const opts = f.options || [];
          const el = t === 'select'
            ? h('select', { class: i === 0 ? 'key-input' : '', onchange: ev => set(ev.target.value) },
              h('option', { value: '' }, '—'),
              opts.map(o => h('option', { value: o, selected: o === cur }, o)),
              cur && !opts.includes(cur) ? h('option', { value: cur, selected: true }, `${cur} (선택지에 없음)`) : null)
            : t === 'textarea'
            ? A.autoGrow(h('textarea', { value: e[f.id] ?? '', oninput: ev => set(ev.target.value) }))
            : h('input', { type: t === 'number' ? 'number' : 'text', value: e[f.id] ?? '',
              class: i === 0 ? 'key-input' : '',
              oninput: ev => set(t === 'number' && ev.target.value !== '' ? Number(ev.target.value) : ev.target.value) });
          return h('td', { 'data-label': f.label || f.id }, el);
        }),
        h('td', { class: 'row-del' }, h('button', { type: 'button', class: 'ghost', title: '이 줄 삭제', onclick: () => {
          draft.entries.splice(draft.entries.indexOf(e), 1); tr.remove(); changed(); applySearch();
        } }, '×')));
      tr._entry = e;
      return tr;
    }
    function drawTable() {
      drawTableHead();
      tbody.replaceChildren(...draft.entries.map(rowEl));
      applySearch();
    }
    function applySearch() {
      const q = search.value.trim().toLowerCase();
      let shown = 0;
      for (const tr of tbody.children) {
        const hit = !q || draft.fields.some(f => String(tr._entry[f.id] ?? '').toLowerCase().includes(q));
        tr.hidden = !hit;
        if (hit) shown++;
      }
      countLabel.textContent = q ? `${shown} / ${draft.entries.length}개` : `${draft.entries.length}개`;
    }
    search.addEventListener('input', applySearch);

    // 엑셀·구글 시트에서 복사한 표(탭으로 구분)나 쉼표로 구분된 글을 한꺼번에 추가
    function parseTable(text) {
      const lines = text.replace(/\r/g, '').split('\n').filter(l => l.trim());
      const tab = lines.some(l => l.includes('\t'));
      const splitCsv = line => {
        const out = []; let cur = '', q = false;
        for (let i = 0; i < line.length; i++) {
          const ch = line[i];
          if (q) { if (ch === '"' && line[i + 1] === '"') { cur += '"'; i++; } else if (ch === '"') q = false; else cur += ch; }
          else if (ch === '"') q = true; else if (ch === ',') { out.push(cur); cur = ''; } else cur += ch;
        }
        out.push(cur);
        return out;
      };
      return lines.map(l => (tab ? l.split('\t') : splitCsv(l)).map(x => x.trim()));
    }
    function drawPasteBox() {
      const ta = h('textarea', { rows: 6, placeholder: `한 줄에 하나씩, 칸은 탭(엑셀에서 복사) 또는 쉼표로 구분해요.\n순서: ${draft.fields.map(f => f.label || f.id).join(' / ')}` });
      pasteBox.replaceChildren(ta, h('div', { class: 'row tight' },
        h('button', { type: 'button', class: 'primary', onclick: () => {
          let rows = parseTable(ta.value);
          if (!rows.length) return;
          const labels = draft.fields.map(f => (f.label || f.id).trim());
          if (rows[0].every((x, i) => !labels[i] || x === labels[i] || x === draft.fields[i].id)) rows = rows.slice(1); // 첫 줄이 항목 이름이면 건너뜀
          const existing = new Map(draft.entries.map(e => [String(e[draft.fields[0].id] ?? '').trim(), e]));
          let added = 0, updated = 0;
          for (const r of rows) {
            const e = {};
            draft.fields.forEach((f, i) => {
              const v = r[i];
              if (v === undefined || v === '') return;
              e[f.id] = f.type === 'number' && !Number.isNaN(Number(v)) ? Number(v) : v.replace(/\\n/g, '\n');
            });
            if (!Object.keys(e).length) continue;
            const k = String(e[draft.fields[0].id] ?? '').trim();
            if (k && existing.has(k)) { Object.assign(existing.get(k), e); updated++; }
            else { draft.entries.push(e); if (k) existing.set(k, e); added++; }
          }
          changed(); drawTable(); pasteBox.hidden = true;
          A.toast(`${added}개 추가${updated ? `, 같은 이름 ${updated}개는 새 내용으로 바꿨어요` : '했어요'}.`);
        } }, '추가하기'),
        h('button', { type: 'button', onclick: () => { pasteBox.hidden = true; } }, '취소'),
        h('span', { class: 'muted small' }, '같은 이름이 이미 있으면 새 내용으로 바뀌어요.')));
    }

    // ---------- 저장 / 삭제 ----------
    function validate() {
      if (!String(draft.name || '').trim()) return '데이터 이름을 입력해 주세요. (예: 스킬)';
      if (!draft.fields.length) return '항목이 하나 이상 있어야 해요.';
      const seen = new Set();
      for (const f of draft.fields) {
        if (!validId(f.id)) return `"${f.label}" 항목의 저장용 이름이 올바르지 않아요.`;
        if (seen.has(f.id)) return `저장용 이름 '${f.id}'가 두 번 쓰였어요.`;
        seen.add(f.id);
      }
      return null;
    }
    async function save(btn, goList) {
      const err = validate();
      if (err) return A.toast(err, true);
      const kf = draft.fields[0].id;
      draft.entries = draft.entries.filter(e => Object.keys(e).length);
      const names = draft.entries.map(e => String(e[kf] ?? '').trim());
      const blank = names.filter(n => !n).length;
      const dup = names.filter((n, i) => n && names.indexOf(n) !== i);
      if ((blank || dup.length) && !confirm(
        (blank ? `${draft.fields[0].label || '이름'}이 비어 있는 줄이 ${blank}개 있어요. 이런 줄은 템플릿에서 불러올 수 없어요.\n` : '') +
        (dup.length ? `같은 ${draft.fields[0].label || '이름'}이 여러 번 있어요: ${[...new Set(dup)].slice(0, 5).join(', ')}. 맨 위의 것만 불러와져요.\n` : '') +
        '그래도 저장할까요?')) return;
      // 같은 표에 함께 연결된 데이터와 항목 구성이 달라지면 알려 줌
      const partnerKeys = [...new Set(usedBy.flatMap(u => u.partners || []))];
      const off = partnerKeys.filter(k => datasets[k] && !E.sameSchema(draft, datasets[k]));
      if (off.length && !confirm(`이 데이터는 ${off.map(k => `"${datasets[k].name || k}"`).join(', ')}와(과) 같은 표에 함께 연결돼 있는데, 항목 구성이 서로 달라져요.\n` +
        '구성이 다른 데이터는 그 표에서 불러오지 않게 돼요. 함께 쓰는 데이터도 똑같이 고쳐 주세요.\n\n그래도 저장할까요?')) return;
      let saveKey = key;
      if (isNew) {
        saveKey = A.fileKey(`${draft.rule || '공용'}-${draft.name}`);
        if (datasets[saveKey]) return A.toast('같은 룰에 같은 이름의 데이터가 이미 있어요.', true);
      }
      btn.disabled = true;
      try {
        const out = clone(draft);
        out.name = String(out.name).trim();
        if (!out.rule) delete out.rule;
        sha = await A.writeJson(`datasets/${saveKey}.json`, out, sha, `데이터 ${isNew ? '생성' : '수정'}: ${out.name}`);
        A.resetDatasets();
        A.setDirty(false);
        A.toast('저장했어요. 이 데이터를 연결한 시트에 바로 반영돼요.');
        if (!isNew) A.refreshShares({ datasetKey: saveKey }).then(({ ok, fail }) => {
          if (fail) A.toast(`공유본 ${fail}개를 새로 올리지 못했어요. 해당 캐릭터를 열어 저장하면 다시 올라가요.`, true);
          else if (ok) A.toast(`저장했어요. 공유 중인 시트 ${ok}개도 새로 반영했어요.`);
        }).catch(() => {});
        if (goList) A.go('/data');
        else if (isNew) A.go(`/data/${encodeURIComponent(saveKey)}`);
      } catch (e) { A.toast(e.message, true); btn.disabled = false; }
    }
    async function remove(btn) {
      if (usedBy.length) return A.toast(`${usedBy.map(u => `${u.rule}의 "${u.table}" 표`).join(', ')}에 연결돼 있어서 지울 수 없어요. 템플릿에서 연결을 먼저 풀어 주세요.`, true);
      if (!confirm(`"${draft.name}" 데이터를 삭제할까요? 되돌릴 수 없어요.`)) return;
      btn.disabled = true;
      try {
        await A.deleteFile(`datasets/${key}.json`, sha, `데이터 삭제: ${draft.name}`);
        A.resetDatasets(); A.setDirty(false); A.toast('삭제했어요.'); A.go('/data');
      } catch (e) { A.toast(e.message, true); btn.disabled = false; }
    }

    // ---------- 화면 ----------
    const ruleSelect = h('select', { onchange: e => { draft.rule = e.target.value; changed(); updateFile(); } },
      h('option', { value: '' }, '공용 (룰 없음)'),
      Object.entries(templates).filter(([, t]) => !t.error).map(([k, t]) => h('option', { value: k, selected: draft.rule === k }, t.rule || k)));
    const fileLabel = h('span', { class: 'muted small' });
    const updateFile = () => { fileLabel.textContent = `저장 파일: datasets/${isNew ? A.fileKey(`${draft.rule || '공용'}-${draft.name}`) : key}.json`; };
    updateFile();
    const saveBtn = h('button', { id: 'save-btn', class: 'primary', onclick: e => save(e.target) }, '저장');

    A.show(
      h('div', { class: 'sheet-head' },
        h('div', { class: 'title-row' },
          h('a', { class: 'back-link', href: '#/data', title: '데이터 목록으로' }, '← 데이터'),
          h('h1', { class: 'page-title' }, isNew ? '새 데이터' : (draft.name || key))),
        h('div', { class: 'row' }, saveBtn,
          h('button', { class: 'save-list-btn', onclick: e => save(e.target, true) }, '저장하고 목록으로'),
          isNew ? null : h('button', { class: 'danger ghost', onclick: e => remove(e.target) }, '삭제'))),
      h('div', { class: 'data-layout' },
        h('div', { class: 'tpl-editor' },
          h('section', { class: 'card tpl-section' },
            h('div', { class: 'row tight' },
              h('label', { class: 'mini-field grow' }, h('span', {}, '데이터 이름'),
                h('input', { value: draft.name || '', placeholder: '예: 스킬, 아이템, 주문', oninput: e => { draft.name = e.target.value; changed(); updateFile(); } })),
              h('label', { class: 'mini-field' }, h('span', {}, '룰'), ruleSelect)),
            fileLabel,
            usedBy.length ? h('p', { class: 'muted small' }, '연결된 템플릿: ' + usedBy.map(u => `${u.rule}의 "${u.table}" 표` +
              (u.partners && u.partners.length ? ` (함께: ${u.partners.map(k => (datasets[k] && datasets[k].name) || k).join(', ')})` : '')).join(', ')) : null),
          schemaCard),
        h('section', { class: 'card data-table-card' },
          h('div', { class: 'toolbar tight' }, h('strong', {}, '데이터'), countLabel, search,
            h('button', { type: 'button', onclick: () => { drawPasteBox(); pasteBox.hidden = !pasteBox.hidden; } }, '표 붙여넣기'),
            h('button', { type: 'button', class: 'primary', onclick: () => {
              const e = {}; draft.entries.push(e); const tr = rowEl(e); tbody.append(tr); changed(); applySearch();
              search.value = ''; applySearch(); tr.querySelector('input, textarea').focus();
            } }, '+ 줄 추가')),
          pasteBox,
          h('div', { class: 'list-wrap' }, h('table', { class: 'list data-table' }, thead, tbody)))));
    A.setDirty(false);
    if (isNew) { saveBtn.disabled = false; saveBtn.textContent = '저장'; }
    drawSchema();
    drawTable();
  }
})();
