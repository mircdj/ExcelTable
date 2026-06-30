/**
 * Formula engine — opt-in with `formulas: true`.
 *
 * A cell whose raw value is a string starting with `=` is a formula. The
 * grid shows the computed value, the editor shows the formula, undo/redo
 * and dirty tracking keep working unchanged (the stored value IS the
 * formula string).
 *
 * Reference semantics (documented, deliberate): `A1` addresses **data
 * coordinates** — column letter = position in the grid's column list,
 * row number = 1-based position in the master data order. Sorting and
 * filtering change only what you see, never what a reference means; row
 * insert/delete triggers a full recalc (references are positional and are
 * not rewritten — Excel-style relative adjustment is out of scope).
 *
 * Syntax (Italian Excel conventions, since locale default is it-IT):
 *  - argument separator `;` — decimals accept both `,` and `.`
 *  - operators: `+ - * / ^ &` (concat), comparisons `= <> < > <= >=`
 *  - ranges `A1:B10`, absolute `$A$1` accepted (the `$` is ignored)
 *  - functions, IT and EN names: SOMMA/SUM, MEDIA/AVERAGE, MIN, MAX,
 *    CONTA/COUNT, CONTA.VALORI/COUNTA, SE/IF, E/AND, O/OR, NON/NOT,
 *    ARROTONDA/ROUND, ASS/ABS, CONCATENA/CONCAT, LUNGHEZZA/LEN,
 *    MAIUSC/UPPER, MINUSC/LOWER, OGGI/TODAY
 *  - errors: #NOME? #RIF! #DIV/0! #CIRC! #ERRORE!
 */

/* ================================================================== */
/* Values & errors                                                     */
/* ================================================================== */

export class FormulaError {
  constructor(public readonly code: string) {}
  toString(): string {
    return this.code;
  }
}
export type FVal = number | string | boolean | null | FormulaError;

const ERR = {
  NAME: new FormulaError('#NOME?'),
  REF: new FormulaError('#RIF!'),
  DIV0: new FormulaError('#DIV/0!'),
  CYCLE: new FormulaError('#CIRC!'),
  VALUE: new FormulaError('#ERRORE!'),
};

/* ================================================================== */
/* Tokenizer                                                           */
/* ================================================================== */

type Tok =
  | { t: 'num'; v: number }
  | { t: 'str'; v: string }
  | { t: 'ref'; col: number; row: number }
  | { t: 'name'; v: string }
  | { t: 'op'; v: string }
  | { t: '('; v: '(' }
  | { t: ')'; v: ')' }
  | { t: ';'; v: ';' }
  | { t: ':'; v: ':' };

export function colLetterToIndex(s: string): number {
  let n = 0;
  for (const ch of s) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

function tokenize(src: string): Tok[] {
  const toks: Tok[] = [];
  let i = 0;
  const n = src.length;
  while (i < n) {
    const c = src[i];
    if (c === ' ' || c === '\t') {
      i++;
      continue;
    }
    if (c === '"') {
      let j = i + 1;
      let out = '';
      while (j < n) {
        if (src[j] === '"' && src[j + 1] === '"') {
          out += '"';
          j += 2;
        } else if (src[j] === '"') break;
        else out += src[j++];
      }
      toks.push({ t: 'str', v: out });
      i = j + 1;
      continue;
    }
    // number: digits with , or . decimal (',' only when not followed by another number context — we
    // resolve the Italian ambiguity by reserving ';' as the only argument separator)
    if (/\d/.test(c)) {
      let j = i;
      while (j < n && /\d/.test(src[j])) j++;
      if ((src[j] === '.' || src[j] === ',') && /\d/.test(src[j + 1] ?? '')) {
        j++;
        while (j < n && /\d/.test(src[j])) j++;
      }
      toks.push({ t: 'num', v: Number(src.slice(i, j).replace(',', '.')) });
      i = j;
      continue;
    }
    // ref or function name (refs may carry $)
    if (/[A-Za-z$]/.test(c)) {
      let j = i;
      while (j < n && /[A-Za-z0-9_.$]/.test(src[j])) j++;
      const word = src.slice(i, j);
      const ref = /^\$?([A-Z]+)\$?(\d+)$/i.exec(word);
      if (ref) toks.push({ t: 'ref', col: colLetterToIndex(ref[1].toUpperCase()), row: Number(ref[2]) - 1 });
      else toks.push({ t: 'name', v: word.toUpperCase() });
      i = j;
      continue;
    }
    if (c === '(') { toks.push({ t: '(', v: '(' }); i++; continue; }
    if (c === ')') { toks.push({ t: ')', v: ')' }); i++; continue; }
    if (c === ';') { toks.push({ t: ';', v: ';' }); i++; continue; }
    if (c === ':') { toks.push({ t: ':', v: ':' }); i++; continue; }
    if (c === '<' && src[i + 1] === '=') { toks.push({ t: 'op', v: '<=' }); i += 2; continue; }
    if (c === '>' && src[i + 1] === '=') { toks.push({ t: 'op', v: '>=' }); i += 2; continue; }
    if (c === '<' && src[i + 1] === '>') { toks.push({ t: 'op', v: '<>' }); i += 2; continue; }
    if ('+-*/^&=<>%'.includes(c)) { toks.push({ t: 'op', v: c }); i++; continue; }
    throw ERR.VALUE;
  }
  return toks;
}

/* ================================================================== */
/* Parser (recursive descent) → AST                                    */
/* ================================================================== */

export type Ast =
  | { t: 'num'; v: number }
  | { t: 'str'; v: string }
  | { t: 'ref'; col: number; row: number }
  | { t: 'range'; c1: number; r1: number; c2: number; r2: number }
  | { t: 'bin'; op: string; l: Ast; r: Ast }
  | { t: 'un'; op: string; e: Ast }
  | { t: 'call'; name: string; args: Ast[] };

class Parser {
  private p = 0;
  constructor(private toks: Tok[]) {}
  private peek(): Tok | undefined {
    return this.toks[this.p];
  }
  private next(): Tok | undefined {
    return this.toks[this.p++];
  }
  private expect(t: string): void {
    const tok = this.next();
    if (!tok || tok.t !== t) throw ERR.VALUE;
  }

  parse(): Ast {
    const e = this.comparison();
    if (this.p !== this.toks.length) throw ERR.VALUE;
    return e;
  }

  private comparison(): Ast {
    let l = this.concat();
    while (this.peek()?.t === 'op' && ['=', '<>', '<', '>', '<=', '>='].includes((this.peek() as { v: string }).v)) {
      const op = (this.next() as { v: string }).v;
      l = { t: 'bin', op, l, r: this.concat() };
    }
    return l;
  }
  private concat(): Ast {
    let l = this.additive();
    while (this.peek()?.t === 'op' && (this.peek() as { v: string }).v === '&') {
      this.next();
      l = { t: 'bin', op: '&', l, r: this.additive() };
    }
    return l;
  }
  private additive(): Ast {
    let l = this.multiplicative();
    while (this.peek()?.t === 'op' && ['+', '-'].includes((this.peek() as { v: string }).v)) {
      const op = (this.next() as { v: string }).v;
      l = { t: 'bin', op, l, r: this.multiplicative() };
    }
    return l;
  }
  private multiplicative(): Ast {
    let l = this.power();
    while (this.peek()?.t === 'op' && ['*', '/'].includes((this.peek() as { v: string }).v)) {
      const op = (this.next() as { v: string }).v;
      l = { t: 'bin', op, l, r: this.power() };
    }
    return l;
  }
  private power(): Ast {
    const l = this.unary();
    if (this.peek()?.t === 'op' && (this.peek() as { v: string }).v === '^') {
      this.next();
      return { t: 'bin', op: '^', l, r: this.power() }; // right-assoc
    }
    return l;
  }
  private unary(): Ast {
    const tok = this.peek();
    if (tok?.t === 'op' && (tok.v === '-' || tok.v === '+')) {
      this.next();
      return { t: 'un', op: tok.v, e: this.unary() };
    }
    return this.primary();
  }
  private primary(): Ast {
    const tok = this.next();
    if (!tok) throw ERR.VALUE;
    if (tok.t === 'num') return { t: 'num', v: tok.v };
    if (tok.t === 'str') return { t: 'str', v: tok.v };
    if (tok.t === '(') {
      const e = this.comparison();
      this.expect(')');
      return e;
    }
    if (tok.t === 'ref') {
      if (this.peek()?.t === ':') {
        this.next();
        const end = this.next();
        if (end?.t !== 'ref') throw ERR.VALUE;
        return {
          t: 'range',
          c1: Math.min(tok.col, end.col),
          r1: Math.min(tok.row, end.row),
          c2: Math.max(tok.col, end.col),
          r2: Math.max(tok.row, end.row),
        };
      }
      return { t: 'ref', col: tok.col, row: tok.row };
    }
    if (tok.t === 'name') {
      if (this.peek()?.t === '(') {
        this.next();
        const args: Ast[] = [];
        if (this.peek()?.t !== ')') {
          args.push(this.comparison());
          while (this.peek()?.t === ';') {
            this.next();
            args.push(this.comparison());
          }
        }
        this.expect(')');
        return { t: 'call', name: tok.v, args };
      }
      if (tok.v === 'VERO' || tok.v === 'TRUE') return { t: 'num', v: 1 };
      if (tok.v === 'FALSO' || tok.v === 'FALSE') return { t: 'num', v: 0 };
      throw ERR.NAME;
    }
    throw ERR.VALUE;
  }
}

export function parseFormula(src: string): Ast {
  return new Parser(tokenize(src.startsWith('=') ? src.slice(1) : src)).parse();
}

/** All cell references (expanded from ranges) an AST depends on. */
export function collectRefs(ast: Ast, out: { col: number; row: number }[] = []): { col: number; row: number }[] {
  switch (ast.t) {
    case 'ref':
      out.push({ col: ast.col, row: ast.row });
      break;
    case 'range':
      for (let r = ast.r1; r <= ast.r2; r++)
        for (let c = ast.c1; c <= ast.c2; c++) out.push({ col: c, row: r });
      break;
    case 'bin':
      collectRefs(ast.l, out);
      collectRefs(ast.r, out);
      break;
    case 'un':
      collectRefs(ast.e, out);
      break;
    case 'call':
      for (const a of ast.args) collectRefs(a, out);
      break;
  }
  return out;
}

/* ================================================================== */
/* Evaluator                                                           */
/* ================================================================== */

export interface EvalContext {
  /** Computed value at data coordinates; null when out of range. */
  cell(col: number, row: number): FVal;
}

const toNum = (v: FVal): number | FormulaError => {
  if (v instanceof FormulaError) return v;
  if (v === null || v === '') return 0;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (typeof v === 'number') return v;
  const n = Number(String(v).replace(',', '.'));
  return Number.isFinite(n) ? n : ERR.VALUE;
};
const toStr = (v: FVal): string =>
  v instanceof FormulaError ? v.code : v === null ? '' : typeof v === 'boolean' ? (v ? 'VERO' : 'FALSO') : String(v);
const truthy = (v: FVal): boolean | FormulaError => {
  if (v instanceof FormulaError) return v;
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v !== 0;
  if (v === null || v === '') return false;
  return true;
};

/** Flatten args: ranges/refs become their values; used by aggregate functions. */
function flatten(args: Ast[], ctx: EvalContext): FVal[] | FormulaError {
  const out: FVal[] = [];
  for (const a of args) {
    if (a.t === 'range') {
      for (let r = a.r1; r <= a.r2; r++)
        for (let c = a.c1; c <= a.c2; c++) {
          const v = ctx.cell(c, r);
          if (v instanceof FormulaError) return v;
          out.push(v);
        }
    } else {
      const v = evaluate(a, ctx);
      if (v instanceof FormulaError) return v;
      out.push(v);
    }
  }
  return out;
}

const numbersOf = (vals: FVal[]): number[] =>
  vals.filter((v): v is number => typeof v === 'number' && Number.isFinite(v));

type Fn = (args: Ast[], ctx: EvalContext) => FVal;

const FUNCTIONS: Record<string, Fn> = {
  SUM: (a, ctx) => {
    const vals = flatten(a, ctx);
    return vals instanceof FormulaError ? vals : numbersOf(vals).reduce((s, v) => s + v, 0);
  },
  AVERAGE: (a, ctx) => {
    const vals = flatten(a, ctx);
    if (vals instanceof FormulaError) return vals;
    const nums = numbersOf(vals);
    return nums.length ? nums.reduce((s, v) => s + v, 0) / nums.length : ERR.DIV0;
  },
  MIN: (a, ctx) => {
    const vals = flatten(a, ctx);
    if (vals instanceof FormulaError) return vals;
    const nums = numbersOf(vals);
    return nums.length ? Math.min(...nums) : 0;
  },
  MAX: (a, ctx) => {
    const vals = flatten(a, ctx);
    if (vals instanceof FormulaError) return vals;
    const nums = numbersOf(vals);
    return nums.length ? Math.max(...nums) : 0;
  },
  COUNT: (a, ctx) => {
    const vals = flatten(a, ctx);
    return vals instanceof FormulaError ? vals : numbersOf(vals).length;
  },
  COUNTA: (a, ctx) => {
    const vals = flatten(a, ctx);
    return vals instanceof FormulaError ? vals : vals.filter((v) => v !== null && v !== '').length;
  },
  IF: (a, ctx) => {
    if (a.length < 2 || a.length > 3) return ERR.VALUE;
    const cond = truthy(evaluate(a[0], ctx));
    if (cond instanceof FormulaError) return cond;
    return cond ? evaluate(a[1], ctx) : a[2] ? evaluate(a[2], ctx) : false;
  },
  AND: (a, ctx) => {
    for (const arg of a) {
      const v = truthy(evaluate(arg, ctx));
      if (v instanceof FormulaError) return v;
      if (!v) return false;
    }
    return true;
  },
  OR: (a, ctx) => {
    for (const arg of a) {
      const v = truthy(evaluate(arg, ctx));
      if (v instanceof FormulaError) return v;
      if (v) return true;
    }
    return false;
  },
  NOT: (a, ctx) => {
    const v = truthy(evaluate(a[0], ctx));
    return v instanceof FormulaError ? v : !v;
  },
  ROUND: (a, ctx) => {
    const v = toNum(evaluate(a[0], ctx));
    if (v instanceof FormulaError) return v;
    const d = a[1] ? toNum(evaluate(a[1], ctx)) : 0;
    if (d instanceof FormulaError) return d;
    const f = Math.pow(10, d);
    return Math.round(v * f) / f;
  },
  ABS: (a, ctx) => {
    const v = toNum(evaluate(a[0], ctx));
    return v instanceof FormulaError ? v : Math.abs(v);
  },
  CONCAT: (a, ctx) => {
    const vals = flatten(a, ctx);
    return vals instanceof FormulaError ? vals : vals.map(toStr).join('');
  },
  LEN: (a, ctx) => toStr(evaluate(a[0], ctx)).length,
  UPPER: (a, ctx) => toStr(evaluate(a[0], ctx)).toUpperCase(),
  LOWER: (a, ctx) => toStr(evaluate(a[0], ctx)).toLowerCase(),
  TODAY: () => new Date().toISOString().slice(0, 10),
};

// Italian aliases
const ALIASES: Record<string, string> = {
  SOMMA: 'SUM', MEDIA: 'AVERAGE', CONTA: 'COUNT', 'CONTA.VALORI': 'COUNTA',
  SE: 'IF', E: 'AND', O: 'OR', NON: 'NOT', ARROTONDA: 'ROUND', ASS: 'ABS',
  CONCATENA: 'CONCAT', LUNGHEZZA: 'LEN', MAIUSC: 'UPPER', MINUSC: 'LOWER', OGGI: 'TODAY',
};

export function evaluate(ast: Ast, ctx: EvalContext): FVal {
  switch (ast.t) {
    case 'num':
      return ast.v;
    case 'str':
      return ast.v;
    case 'ref':
      return ctx.cell(ast.col, ast.row);
    case 'range':
      return ERR.VALUE; // a bare range is only valid inside aggregate functions
    case 'un': {
      const v = toNum(evaluate(ast.e, ctx));
      return v instanceof FormulaError ? v : ast.op === '-' ? -v : v;
    }
    case 'bin': {
      const { op } = ast;
      if (op === '&') return toStr(evaluate(ast.l, ctx)) + toStr(evaluate(ast.r, ctx));
      if (['=', '<>', '<', '>', '<=', '>='].includes(op)) {
        const l = evaluate(ast.l, ctx);
        const r = evaluate(ast.r, ctx);
        if (l instanceof FormulaError) return l;
        if (r instanceof FormulaError) return r;
        const [a, b] =
          typeof l === 'number' || typeof r === 'number'
            ? [toNum(l), toNum(r)]
            : [toStr(l).toLowerCase(), toStr(r).toLowerCase()];
        if (a instanceof FormulaError) return a;
        if (b instanceof FormulaError) return b;
        switch (op) {
          case '=': return a === b;
          case '<>': return a !== b;
          case '<': return a < b;
          case '>': return a > b;
          case '<=': return a <= b;
          case '>=': return a >= b;
        }
      }
      const l = toNum(evaluate(ast.l, ctx));
      if (l instanceof FormulaError) return l;
      const r = toNum(evaluate(ast.r, ctx));
      if (r instanceof FormulaError) return r;
      switch (op) {
        case '+': return l + r;
        case '-': return l - r;
        case '*': return l * r;
        case '/': return r === 0 ? ERR.DIV0 : l / r;
        case '^': return Math.pow(l, r);
        default: return ERR.VALUE;
      }
    }
    case 'call': {
      const fn = FUNCTIONS[ALIASES[ast.name] ?? ast.name];
      if (!fn) return ERR.NAME;
      try {
        return fn(ast.args, ctx);
      } catch (e) {
        return e instanceof FormulaError ? e : ERR.VALUE;
      }
    }
  }
}

/* ================================================================== */
/* Grid-facing engine: cache + cycle detection                         */
/* ================================================================== */

export const isFormula = (v: unknown): v is string => typeof v === 'string' && v.startsWith('=');

/**
 * Per-grid computation cache. `raw(col,row)` supplies stored values;
 * computed results are memoized until `invalidate()` (the grid calls it
 * on every data mutation). Cycles resolve to #CIRC!.
 */
export class FormulaEngine {
  private cache = new Map<string, FVal>();
  private computing = new Set<string>();
  private asts = new Map<string, Ast | FormulaError>();

  constructor(private raw: (col: number, row: number) => unknown) {}

  invalidate(): void {
    this.cache.clear();
    this.asts.clear();
  }

  /** Computed value at data coordinates. Non-formula cells pass through. */
  cell = (col: number, row: number): FVal => {
    const key = col + ':' + row;
    if (this.cache.has(key)) return this.cache.get(key)!;
    if (this.computing.has(key)) return ERR.CYCLE;
    const rawVal = this.raw(col, row);
    if (rawVal === undefined) return null; // out of range → empty (Excel-like)
    let result: FVal;
    if (isFormula(rawVal)) {
      this.computing.add(key);
      try {
        let ast = this.asts.get(rawVal);
        if (!ast) {
          try {
            ast = parseFormula(rawVal);
          } catch (e) {
            ast = e instanceof FormulaError ? e : ERR.VALUE;
          }
          this.asts.set(rawVal, ast);
        }
        result = ast instanceof FormulaError ? ast : evaluate(ast, { cell: this.cell });
      } finally {
        this.computing.delete(key);
      }
    } else {
      result =
        rawVal === null || rawVal === ''
          ? null
          : typeof rawVal === 'number' || typeof rawVal === 'boolean'
            ? rawVal
            : String(rawVal);
    }
    this.cache.set(key, result);
    return result;
  };
}
