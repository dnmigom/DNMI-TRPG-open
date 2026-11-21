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
  };
  const ALIAS = {
    내림: 'floor', 올림: 'ceil', 반올림: 'round', 버림: 'trunc', 절댓값: 'abs',
    최소: 'min', 최대: 'max', 합: 'sum', 합계: 'sum', 범위: 'clamp', 만약: 'if',
    목록합계: 'total', 개수: 'count',
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

  // ---------- 4. 시트 하나의 값 계산기 ----------
  function toNumber(v) {
    if (v === true) return 1;
    if (v === false || v === '' || v == null) return 0;
    const n = Number(v);
    return Number.isNaN(n) ? 0 : n;
  }

  function makeResolver(template, sheet) {
    const fields = {}, lists = {};
    for (const s of template.sections || []) {
      if (s.type === 'list') lists[s.id] = s;
      else for (const f of s.fields || []) fields[f.id] = f;
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
        const v = f.formula ? guarded(id, () => evaluate(f.formula, ctx)) : toNumber(values[id]);
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
          return toNumber(row[id]);
        },
        column: ctx.column,
        count: ctx.count,
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
    for (const s of template.sections || []) for (const f of s.fields || []) fields[f.id] = f;
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
        .map(f => ({ label: f.label || f.id, value: `{${f.id}}` })),
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
  function datasetIndex(ds) {
    const keyField = ds && (ds.fields || [])[0];
    if (!keyField) return null;
    const map = new Map();
    for (const e of ds.entries || []) {
      const k = String(e[keyField.id] ?? '').trim();
      if (k && !map.has(k)) map.set(k, e);
    }
    return { keyField, map };
  }
  function linkInfo(section, datasets) {
    if (!section || section.type !== 'list' || !section.source || !section.source.dataset) return null;
    const ds = datasets && datasets[section.source.dataset];
    const idx = datasetIndex(ds);
    if (!idx) return null;
    const keyCol = (section.columns || []).find(c => c.from === idx.keyField.id);
    if (!keyCol) return null;
    return { ds, keyCol, keyField: idx.keyField, map: idx.map, find: name => idx.map.get(String(name ?? '').trim()) || null };
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

  const api = { parse, evaluate, makeResolver, format, toNumber, toCcfolia, autoCcfolia,
    linkInfo, fillRowFromEntry, syncLinkedRows };
  global.SheetEngine = api;
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
})(typeof window !== 'undefined' ? window : globalThis);
