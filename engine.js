/*
 * engine.js — 시트 계산 엔진
 * 템플릿의 "formula"를 안전하게 계산합니다. (자바스크립트 eval을 쓰지 않음)
 * 허용: 숫자, 칸 id, + - * / %, 괄호, 비교(> < >= <= == != =), &&, ||, !, 정해진 함수
 */
(function (global) {
  'use strict';

  // ---------- 1. 글자를 토막으로 나누기 ----------
  function tokenize(src) {
    const tokens = [];
    let i = 0;
    const isNum = c => /[0-9.]/.test(c);
    const isIdStart = c => /[\p{L}_]/u.test(c);
    const isId = c => /[\p{L}\p{N}_]/u.test(c);
    while (i < src.length) {
      const c = src[i];
      if (/\s/.test(c)) { i++; continue; }
      if (isNum(c)) {
        let j = i; while (j < src.length && isNum(src[j])) j++;
        const v = Number(src.slice(i, j));
        if (Number.isNaN(v)) throw new Error(`숫자를 읽을 수 없어요: ${src.slice(i, j)}`);
        tokens.push({ t: 'num', v }); i = j; continue;
      }
      if (isIdStart(c)) {
        let j = i; while (j < src.length && isId(src[j])) j++;
        tokens.push({ t: 'id', v: src.slice(i, j) }); i = j; continue;
      }
      if (c === '"' || c === "'") {
        let j = i + 1; while (j < src.length && src[j] !== c) j++;
        if (j >= src.length) throw new Error('따옴표가 닫히지 않았어요');
        tokens.push({ t: 'str', v: src.slice(i + 1, j) }); i = j + 1; continue;
      }
      const two = src.slice(i, i + 2);
      if (['>=', '<=', '==', '!=', '&&', '||'].includes(two)) { tokens.push({ t: 'op', v: two }); i += 2; continue; }
      if (c === '=') { tokens.push({ t: 'op', v: '==' }); i++; continue; } // "=" 하나도 "같다"로 취급
      if ('+-*/%(),<>!'.includes(c)) { tokens.push({ t: 'op', v: c }); i++; continue; }
      throw new Error(`수식에 쓸 수 없는 문자가 있어요: ${c}`);
    }
    return tokens;
  }

  // ---------- 2. 토막을 계산 구조로 바꾸기 ----------
  function parse(src) {
    const tk = tokenize(String(src));
    let p = 0;
    const peek = () => tk[p];
    const isOp = v => tk[p] && tk[p].t === 'op' && tk[p].v === v;
    const expect = v => { if (!isOp(v)) throw new Error(`'${v}'가 있어야 할 자리예요`); p++; };

    function binary(next, ops) {
      return function () {
        let a = next();
        while (tk[p] && tk[p].t === 'op' && ops.includes(tk[p].v)) {
          const op = tk[p++].v;
          a = { k: 'bin', op, a, b: next() };
        }
        return a;
      };
    }
    function unary() {
      if (isOp('-') || isOp('+') || isOp('!')) { const op = tk[p++].v; return { k: 'un', op, a: unary() }; }
      return primary();
    }
    function primary() {
      const t = peek();
      if (!t) throw new Error('수식이 중간에 끝났어요');
      if (t.t === 'num') { p++; return { k: 'num', v: t.v }; }
      if (t.t === 'str') { p++; return { k: 'str', v: t.v }; }
      if (t.t === 'id') {
        p++;
        if (isOp('(')) {
          p++;
          const args = [];
          if (!isOp(')')) { args.push(or()); while (isOp(',')) { p++; args.push(or()); } }
          expect(')');
          return { k: 'call', name: t.v, args };
        }
        return { k: 'ref', id: t.v };
      }
      if (isOp('(')) { p++; const e = or(); expect(')'); return e; }
      throw new Error(`'${t.v}'는 이 자리에 올 수 없어요`);
    }
    const mul = binary(unary, ['*', '/', '%']);
    const add = binary(mul, ['+', '-']);
    const cmp = binary(add, ['>', '<', '>=', '<=', '==', '!=']);
    const and = binary(cmp, ['&&']);
    const or = binary(and, ['||']);

    const ast = or();
    if (p < tk.length) throw new Error(`'${tk[p].v}' 부분을 이해할 수 없어요`);
    return ast;
  }

  // ---------- 3. 함수 목록 (영어 / 한국어 둘 다 사용 가능) ----------
  const num = v => (typeof v === 'number' ? v : Number(v) || 0);
  const sameText = (a, b) => String(a ?? '').trim() === String(b ?? '').trim();
  const FUNCS = {
    floor: (ev, a) => Math.floor(num(ev(a[0]))),
    ceil: (ev, a) => Math.ceil(num(ev(a[0]))),
    round: (ev, a) => { const d = a[1] ? num(ev(a[1])) : 0; const m = 10 ** d; return Math.round(num(ev(a[0])) * m) / m; },
    trunc: (ev, a) => Math.trunc(num(ev(a[0]))),
    abs: (ev, a) => Math.abs(num(ev(a[0]))),
    min: (ev, a) => Math.min(...a.map(x => num(ev(x)))),
    max: (ev, a) => Math.max(...a.map(x => num(ev(x)))),
    sum: (ev, a) => a.reduce((s, x) => s + num(ev(x)), 0),
    clamp: (ev, a) => Math.min(Math.max(num(ev(a[0])), num(ev(a[1]))), num(ev(a[2]))),
    if: (ev, a) => (ev(a[0]) ? ev(a[1]) : (a.length > 2 ? ev(a[2]) : 0)),
    // 목록(표) 함수: total("목록id", "열id"), count("목록id")
    total: (ev, a, ctx) => ctx.column(String(ev(a[0])), String(ev(a[1]))).reduce((s, v) => s + num(v), 0),
    count: (ev, a, ctx) => ctx.count(String(ev(a[0]))),
    // 표에서 글자 찾기: has("표id", "열id", "A") → 그 열에 A가 있으면 1, countif(…) → A가 몇 줄인지,
    // cell("표id", 줄번호, "열id") → 그 칸의 값 (줄번호는 1부터)
    has: (ev, a, ctx) => (ctx.raw(String(ev(a[0])), String(ev(a[1]))).some(v => sameText(v, ev(a[2]))) ? 1 : 0),
    countif: (ev, a, ctx) => ctx.raw(String(ev(a[0])), String(ev(a[1]))).filter(v => sameText(v, ev(a[2]))).length,
    cell: (ev, a, ctx) => { const col = ctx.raw(String(ev(a[0])), String(ev(a[2]))); const v = col[num(ev(a[1])) - 1]; return v === undefined || v === null ? '' : v; },
    // 특기표: target("표id", "특기 이름") → 달성치, learned("표id", "특기 이름") → 습득했으면 1
    target: (ev, a, ctx) => ctx.grid(String(ev(a[0])), String(ev(a[1]))).value,
    learned: (ev, a, ctx) => (ctx.grid(String(ev(a[0])), String(ev(a[1]))).learned ? 1 : 0),
  };
  const ALIAS = {
    내림: 'floor', 올림: 'ceil', 반올림: 'round', 버림: 'trunc', 절댓값: 'abs',
    최소: 'min', 최대: 'max', 합: 'sum', 합계: 'sum', 범위: 'clamp', 만약: 'if',
    목록합계: 'total', 개수: 'count', 달성치: 'target', 습득: 'learned',
    포함: 'has', 몇개: 'countif', 표칸: 'cell',
  };

  function evalAst(node, ctx) {
    const ev = n => evalAst(n, ctx);
    switch (node.k) {
      case 'num': case 'str': return node.v;
      case 'ref': return ctx.get(node.id);
      case 'un': {
        const v = ev(node.a);
        return node.op === '-' ? -num(v) : node.op === '!' ? (v ? 0 : 1) : num(v);
      }
      case 'bin': {
        if (node.op === '&&') return ev(node.a) && ev(node.b) ? 1 : 0;
        if (node.op === '||') return ev(node.a) || ev(node.b) ? 1 : 0;
        if (node.op === '==' || node.op === '!=') {
          // 글자끼리도 비교 가능: 이름 = "A"
          const x = ev(node.a), y = ev(node.b);
          const eq = (typeof x === 'string' || typeof y === 'string')
            ? String(x ?? '').trim() === String(y ?? '').trim()
            : num(x) === num(y);
          return (node.op === '==') === eq ? 1 : 0;
        }
        const a = num(ev(node.a)), b = num(ev(node.b));
        switch (node.op) {
          case '+': return a + b;
          case '-': return a - b;
          case '*': return a * b;
          case '/': if (b === 0) throw new Error('0으로 나눌 수 없어요'); return a / b;
          case '%': if (b === 0) throw new Error('0으로 나눌 수 없어요'); return a % b;
          case '>': return a > b ? 1 : 0;
          case '<': return a < b ? 1 : 0;
          case '>=': return a >= b ? 1 : 0;
          case '<=': return a <= b ? 1 : 0;
          case '==': return a === b ? 1 : 0;
          case '!=': return a !== b ? 1 : 0;
        }
        break;
      }
      case 'call': {
        const key = ALIAS[node.name] || node.name.toLowerCase();
        const fn = FUNCS[key];
        if (!fn) throw new Error(`'${node.name}'라는 함수는 없어요`);
        return fn(ev, node.args, ctx);
      }
    }
    throw new Error('알 수 없는 수식이에요');
  }

  const astCache = new Map();
  function evaluate(src, ctx) {
    let ast = astCache.get(src);
    if (!ast) { ast = parse(src); astCache.set(src, ast); }
    return evalAst(ast, ctx);
  }

  // ---------- 3-1. 특기표(격자) ----------
  // section: { type: 'grid', id, columns: [열 이름], rows: [줄 이름], cells: [[특기 이름]] (줄 × 열),
  //            base(습득한 칸의 달성치, 기본 5), empty(아무것도 없을 때, 기본 12), cap(최대, 기본 12, 0이면 제한 없음),
  //            rowCost(위아래 한 칸, 기본 1), colCost(옆 열로 한 칸, 기본 1), gapCost(열 사이 빈 줄, 기본 1) }
  // state:   { acquired: ["줄-열", …], blackout: 열 번호 | null, disabled: [열 번호, …] }
  //   blackout: 고른 열의 양쪽 빈 줄을 검게 칠해 무효화 (맨 왼쪽·맨 오른쪽 열은 한쪽만)
  //   disabled: 그 열의 습득을 모두 없는 것으로 취급
  function gridNum(v, d) { const n = Number(v); return v === undefined || v === null || v === '' || Number.isNaN(n) ? d : n; }
  function gridState(section, sheet) {
    const st = (sheet && sheet.grids && sheet.grids[section.id]) || {};
    return { acquired: st.acquired || [], blackout: st.blackout ?? null, disabled: st.disabled || [] };
  }
  function computeGrid(section, state) {
    const R = (section.rows || []).length, C = (section.columns || []).length;
    const base = gridNum(section.base, 5), empty = gridNum(section.empty, 12), cap = gridNum(section.cap, 12);
    const rowCost = gridNum(section.rowCost, 1), colCost = gridNum(section.colCost, 1), gapCost = gridNum(section.gapCost, 1);
    const disabled = new Set((state.disabled || []).map(Number));
    // 열 c와 c+1 사이를 지나는 비용
    const black = state.blackout === null || state.blackout === undefined ? -1 : Number(state.blackout);
    const cross = c => colCost + ((black === c || black === c + 1) ? 0 : gapCost);
    const prefix = [0];
    for (let c = 0; c < C - 1; c++) prefix.push(prefix[c] + cross(c));
    const sources = [];
    for (const key of state.acquired || []) {
      const [r, c] = String(key).split('-').map(Number);
      if (r >= 0 && r < R && c >= 0 && c < C && !disabled.has(c)) sources.push([r, c]);
    }
    const out = [];
    for (let r = 0; r < R; r++) {
      const row = [];
      for (let c = 0; c < C; c++) {
        if (!sources.length) { row.push(empty); continue; }
        let best = Infinity;
        for (const [sr, sc] of sources) {
          const d = Math.abs(r - sr) * rowCost + Math.abs(prefix[c] - prefix[sc]);
          if (d < best) best = d;
        }
        const v = base + best;
        row.push(cap > 0 ? Math.min(cap, v) : v);
      }
      out.push(row);
    }
    return out;
  }
  function gridFind(section, name) {
    name = String(name ?? '').trim();
    const cells = section.cells || [];
    for (let r = 0; r < cells.length; r++) for (let c = 0; c < (cells[r] || []).length; c++) {
      if (String(cells[r][c] ?? '').trim() === name && name) return [r, c];
    }
    return null;
  }

  // ---------- 4. 시트 하나의 값 계산기 ----------
  function toNumber(v) {
    if (v === true) return 1;
    if (v === false || v === '' || v == null) return 0;
    const n = Number(v);
    return Number.isNaN(n) ? 0 : n;
  }

  const TEXTY = ['text', 'textarea', 'select'];
  function makeResolver(template, sheet) {
    const fields = {}, lists = {}, grids = {};
    for (const s of template.sections || []) {
      if (s.type === 'grid') { grids[s.id] = s; continue; }
      if (s.type === 'list') lists[s.id] = s;
      else for (const f of s.fields || []) {
        fields[f.id] = f;
        for (const p of f.parts || []) fields[p.id] = p; // 한 칸에 붙은 추가 숫자도 각자 이름으로 부를 수 있음
      }
    }
    const values = sheet.values || {};
    const rowsOf = id => (sheet.lists && sheet.lists[id]) || [];
    const cache = new Map();
    const visiting = new Set();

    function guarded(key, fn) {
      if (visiting.has(key)) throw new Error('수식끼리 서로를 계속 참조하고 있어요 (순환 참조)');
      visiting.add(key);
      try { return fn(); } finally { visiting.delete(key); }
    }

    const ctx = {
      get(id) {
        if (cache.has(id)) return cache.get(id);
        const f = fields[id];
        if (!f) throw new Error(`'${id}'라는 칸이 없어요`);
        const v = f.formula ? guarded(id, () => evaluate(f.formula, ctx))
          : TEXTY.includes(f.type) ? String(values[id] ?? '') // 글자 칸은 글자 그대로 (계산할 땐 숫자로 바뀜)
          : toNumber(values[id]);
        cache.set(id, v);
        return v;
      },
      column(listId, colId) {
        const s = lists[listId];
        if (!s) throw new Error(`'${listId}'라는 목록이 없어요`);
        return rowsOf(listId).map((row, i) => rowCtx(s, row, i).get(colId));
      },
      count(listId) {
        if (!lists[listId]) throw new Error(`'${listId}'라는 목록이 없어요`);
        return rowsOf(listId).length;
      },
      raw(listId, colId) {
        const s = lists[listId];
        if (!s) throw new Error(`'${listId}'라는 목록이 없어요`);
        const c = (s.columns || []).find(x => x.id === colId);
        if (!c) throw new Error(`'${listId}' 표에 '${colId}' 열이 없어요`);
        return rowsOf(listId).map((row, i) => (c.formula !== undefined ? rowCtx(s, row, i).get(colId) : row[colId]));
      },
      grid(gridId, name) {
        const g = grids[gridId];
        if (!g) throw new Error(`'${gridId}'라는 특기표가 없어요`);
        const pos = gridFind(g, name);
        if (!pos) throw new Error(`특기표에 '${name}'이(가) 없어요`);
        const key = `${gridId}#table`;
        if (!cache.has(key)) cache.set(key, computeGrid(g, gridState(g, sheet)));
        const st = gridState(g, sheet);
        return { value: cache.get(key)[pos[0]][pos[1]], learned: st.acquired.includes(pos.join('-')) && !st.disabled.map(Number).includes(pos[1]) };
      },
    };

    // 표의 한 줄 안에서는 같은 줄의 열 id를 먼저 찾고, 없으면 시트 전체 칸을 찾음
    function rowCtx(s, row, i) {
      const cols = {};
      for (const c of s.columns || []) cols[c.id] = c;
      const rc = {
        get(id) {
          const c = cols[id];
          if (!c) return ctx.get(id);
          if (c.formula) return guarded(`${s.id}#${i}.${id}`, () => evaluate(c.formula, rc));
          return TEXTY.includes(c.type || 'text') && c.type !== undefined ? String(row[id] ?? '') : toNumber(row[id]);
        },
        column: ctx.column,
        count: ctx.count,
        raw: ctx.raw,
        grid: ctx.grid,
      };
      return rc;
    }

    return {
      get: ctx.get,
      evaluate: src => evaluate(String(src), ctx),
      cell: (listId, i, colId) => rowCtx(lists[listId], rowsOf(listId)[i] || {}, i).get(colId),
      rowEvaluate: (listId, i, src) => evaluate(String(src), rowCtx(lists[listId], rowsOf(listId)[i] || {}, i)),
    };
  }

  // ---------- 5. 코코포리아 내보내기 ----------
  // 글 틀 안의 {칸이름} 또는 {수식}을 값으로 바꿈. {#이름}은 캐릭터 이름.
  const TEXT_TYPES = ['text', 'textarea', 'select', 'image'];
  function fillText(text, r, template, sheet, row, warn) {
    const fields = {};
    for (const s of template.sections || []) for (const f of s.fields || []) { fields[f.id] = f; for (const p of f.parts || []) fields[p.id] = p; }
    const cols = {};
    if (row) {
      const l = (template.sections || []).find(s => s.type === 'list' && s.id === row.listId);
      for (const c of (l && l.columns) || []) cols[c.id] = c;
    }
    return String(text || '').replace(/\{([^{}]+)\}/g, (m, expr) => {
      expr = expr.trim();
      if (expr === '#이름') return sheet.name || '';
      if (row && cols[expr] && cols[expr].formula === undefined && TEXT_TYPES.includes(cols[expr].type || 'text')) return String(row.data[expr] ?? '');
      if (!(row && cols[expr]) && fields[expr] && fields[expr].formula === undefined && TEXT_TYPES.includes(fields[expr].type || 'text')) return String((sheet.values || {})[expr] ?? '');
      try { return format(row ? r.rowEvaluate(row.listId, row.i, expr) : r.evaluate(expr)); }
      catch (e) { warn(`{${expr}}: ${e.message}`); return '?'; }
    });
  }

  // 연동 설정이 없는 룰 틀은 자동으로 추측: 최대값이 있는 숫자 칸 → 상태, 나머지 숫자·계산 칸 → 파라미터
  function autoCcfolia(template) {
    const fields = (template.sections || []).filter(s => s.type !== 'list').flatMap(s => s.fields || []);
    const maxRefs = new Set(fields.filter(f => typeof f.max === 'string').map(f => f.max.trim()));
    return {
      status: fields.filter(f => f.type === 'number' && f.max != null)
        .map(f => ({ label: f.label || f.id, value: f.id, max: String(f.max) })),
      params: fields.filter(f => (f.type === 'number' && f.max == null) || f.formula !== undefined)
        .filter(f => !maxRefs.has(f.id))
        .map(f => ({ label: f.label || f.id, value: [f, ...(f.parts || [])].map(x => `{${x.id}}`).join(f.sep || '/') })),
      initiative: '',
      commands: '',
      listCommands: [],
      memo: '',
    };
  }

  function toCcfolia(template, sheet, opts) {
    opts = opts || {};
    const cfg = template.ccfolia || autoCcfolia(template);
    const r = makeResolver(template, sheet);
    const warnings = [];
    const num = (src, label) => {
      if (src == null || String(src).trim() === '') return 0;
      try { const v = Number(r.evaluate(String(src))); return Number.isFinite(v) ? Math.round(v * 100) / 100 : 0; }
      catch (e) { warnings.push(`${label}: ${e.message}`); return 0; }
    };
    const text = (t, label, row) => fillText(t, r, template, sheet, row, m => warnings.push(`${label} ${m}`));

    let commands = text(cfg.commands, '채팅 팔레트');
    for (const lc of cfg.listCommands || []) {
      const rows = (sheet.lists && sheet.lists[lc.list]) || [];
      const lines = rows.map((data, i) => (Object.keys(data).length ? text(lc.line, '채팅 팔레트', { listId: lc.list, i, data }) : ''))
        .filter(line => line.trim());
      if (lines.length) commands += (commands && !commands.endsWith('\n') ? '\n' : '') + lines.join('\n');
    }
    const data = {
      name: sheet.name || '',
      memo: text(cfg.memo, '메모'),
      initiative: num(cfg.initiative, '이니셔티브'),
      status: (cfg.status || []).map(s => ({ label: s.label || '', value: num(s.value, s.label), max: num(s.max, s.label) })),
      params: (cfg.params || []).map(p => ({ label: p.label || '', value: text(p.value, p.label) })),
      commands,
    };
    if (cfg.color) data.color = cfg.color;
    if (opts.externalUrl) data.externalUrl = opts.externalUrl;
    return { clipboard: { kind: 'character', data }, warnings };
  }

  function format(v) {
    if (typeof v !== 'number') return String(v);
    if (!Number.isFinite(v)) return '—';
    return String(Math.round(v * 100) / 100);
  }

  // ---------- 6. 데이터 보관함 연결 ----------
  // 표(list)에 "source": { "dataset": "데이터키" }가 있고, 열에 "from": "데이터 항목 id"가 있으면
  // 찾기 이름 열(데이터의 첫 항목과 연결된 열)에 적힌 이름으로 데이터를 찾아 나머지 연결된 열을 채움
  // 한 표에 데이터를 여러 개 연결할 수 있음: "source": { "dataset": "첫 데이터", "also": ["함께 쓸 데이터", …] }
  // 함께 쓰는 데이터는 항목 구성(저장용 이름들)이 첫 데이터와 같아야 함. 같은 이름이 겹치면 앞쪽 데이터가 우선.
  function sourceKeys(section) {
    const src = section && section.source;
    if (!src || !src.dataset) return [];
    return [src.dataset, ...(src.also || []).filter(k => k && k !== src.dataset)];
  }
  function sameSchema(a, b) {
    const fa = (a && a.fields) || [], fb = (b && b.fields) || [];
    if (!fa.length || !fb.length || fa[0].id !== fb[0].id || fa.length !== fb.length) return false;
    const ids = new Set(fa.map(f => f.id));
    return fb.every(f => ids.has(f.id));
  }
  function schemaDiff(a, b) {
    const ia = new Set(((a && a.fields) || []).map(f => f.id)), ib = new Set(((b && b.fields) || []).map(f => f.id));
    return { onlyA: [...ia].filter(x => !ib.has(x)), onlyB: [...ib].filter(x => !ia.has(x)),
      keyDiff: ((a && a.fields) || [])[0]?.id !== ((b && b.fields) || [])[0]?.id };
  }
  function linkInfo(section, datasets) {
    if (!section || section.type !== 'list') return null;
    const keys = sourceKeys(section);
    if (!keys.length) return null;
    const first = datasets && datasets[keys[0]];
    const keyField = first && (first.fields || [])[0];
    if (!keyField) return null;
    const keyCol = (section.columns || []).find(c => c.from === keyField.id);
    if (!keyCol) return null;
    const map = new Map(), from = new Map();
    for (const k of keys) {
      const ds = datasets[k];
      if (!ds || (k !== keys[0] && !sameSchema(first, ds))) continue; // 구성이 다른 데이터는 건너뜀
      for (const e of ds.entries || []) {
        const name = String(e[keyField.id] ?? '').trim();
        if (name && !map.has(name)) { map.set(name, e); from.set(name, k); }
      }
    }
    return { ds: first, keys, keyCol, keyField, map, from, find: name => map.get(String(name ?? '').trim()) || null };
  }
  function fillRowFromEntry(section, row, entry, keyCol) {
    let changed = false;
    for (const c of section.columns || []) {
      if (!c.from || c === keyCol) continue;
      const v = entry[c.from] ?? '';
      if (row[c.id] !== v) { row[c.id] = v; changed = true; }
    }
    return changed;
  }
  function syncLinkedRows(template, sheet, datasets) {
    let changed = false;
    for (const s of template.sections || []) {
      const link = linkInfo(s, datasets);
      if (!link) continue;
      for (const row of (sheet.lists && sheet.lists[s.id]) || []) {
        const e = link.find(row[link.keyCol.id]);
        if (e && fillRowFromEntry(s, row, e, link.keyCol)) changed = true;
      }
    }
    return changed;
  }

  const api = { parse, evaluate, makeResolver, format, toNumber, toCcfolia, autoCcfolia, computeGrid, gridState, gridFind,
    linkInfo, fillRowFromEntry, syncLinkedRows, sourceKeys, sameSchema, schemaDiff };
  global.SheetEngine = api;
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
})(typeof window !== 'undefined' ? window : globalThis);
