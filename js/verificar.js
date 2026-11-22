let __fxReg = false;
function registerExtra(HF) {
  if (__fxReg) return;
  __fxReg = true;
  const { FunctionPlugin, FunctionArgumentType: T, SimpleRangeValue, CellError, ErrorType, ArraySize } = HF;
  const roundAway = (x, d) => {
    const p = Math.pow(10, d);
    const y = Math.abs(x) * p;
    return (Math.sign(x) * Math.round(y + 1e-9 * Math.max(1, y))) / p;
  };
  class Extra extends FunctionPlugin {
    fixed(ast, state) {
      return this.runFunction(ast.args, state, this.metadata('FIXED'), (n, d, noc) => {
        d = Math.trunc(d);
        if (d > 127) return new CellError(ErrorType.VALUE);
        const v = roundAway(n, d);
        let [ip, fp] = Math.abs(v).toFixed(Math.max(0, d)).split('.');
        if (!noc) ip = ip.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
        return (v < 0 ? '-' : '') + ip + (fp ? '.' + fp : '');
      });
    }
    replaceb(ast, state) {
      return this.runFunction(ast.args, state, this.metadata('REPLACEB'), (t, st, n, nw) => {
        st = Math.trunc(st);
        n = Math.trunc(n);
        if (st < 1 || n < 0) return new CellError(ErrorType.VALUE);
        return t.slice(0, st - 1) + nw + t.slice(st - 1 + n);
      });
    }
    udf() {
      return new CellError(ErrorType.NAME);
    }
    linest(ast, state) {
      try {
        return this.runFunction(ast.args, state, this.metadata('LINEST'), (Y, X, cst, stats) => {
          if (stats) return new CellError(ErrorType.NAME);
          const y = Y.rawData(),
            yh = Y.height(),
            yw = Y.width();
          let ys;
          if (yw === 1) ys = y.map(r => r[0]);
          else if (yh === 1) ys = y[0].slice();
          else return new CellError(ErrorType.VALUE);
          const n = ys.length;
          let rows;
          if (X === undefined || X === null) rows = ys.map((_, i) => [i + 1]);
          else {
            const x = X.rawData();
            if (yw === 1) {
              if (x.length !== n) return new CellError(ErrorType.REF);
              rows = x.map(r => r.slice());
            } else {
              if (x[0].length !== n) return new CellError(ErrorType.REF);
              rows = x[0].map((_, j) => x.map(r => r[j]));
            }
          }
          if (rows.some(r => r.some(v => typeof v !== 'number')) || ys.some(v => typeof v !== 'number'))
            return new CellError(ErrorType.VALUE);
          const k = rows[0].length;
          const A = rows.map(r => (cst ? r.concat([1]) : r.slice()));
          const m = A[0].length;
          const M = [];
          for (let i = 0; i < m; i++) {
            M.push([]);
            for (let j = 0; j < m; j++) {
              let s = 0;
              for (let t = 0; t < n; t++) s += A[t][i] * A[t][j];
              M[i].push(s);
            }
            let s = 0;
            for (let t = 0; t < n; t++) s += A[t][i] * ys[t];
            M[i].push(s);
          }
          for (let c = 0; c < m; c++) {
            let p = c;
            for (let r = c + 1; r < m; r++) if (Math.abs(M[r][c]) > Math.abs(M[p][c])) p = r;
            if (Math.abs(M[p][c]) < 1e-300) return new CellError(ErrorType.NUM);
            [M[c], M[p]] = [M[p], M[c]];
            for (let r = 0; r < m; r++) {
              if (r === c) continue;
              const f = M[r][c] / M[c][c];
              for (let j = c; j <= m; j++) M[r][j] -= f * M[c][j];
            }
          }
          const b = M.map((r, i) => r[m] / r[i]);
          return SimpleRangeValue.onlyValues([
            b
              .slice(0, k)
              .reverse()
              .concat([cst ? b[k] : 0]),
          ]);
        });
      } catch (e) {
        return new CellError(ErrorType.NAME);
      }
    }
    linestSize(ast, state) {
      if (ast.args.length < 2) return new ArraySize(2, 1);
      const st2 = new state.constructor(state.formulaAddress, true);
      const xs = this.arraySizeForAst(ast.args[1], st2);
      const ys = this.arraySizeForAst(ast.args[0], st2);
      return new ArraySize((ys.width === 1 ? xs.width : xs.height) + 1, 1);
    }
  }
  Extra.implementedFunctions = {
    FIXED: {
      method: 'fixed',
      parameters: [
        { argumentType: T.NUMBER },
        { argumentType: T.NUMBER, defaultValue: 2 },
        { argumentType: T.BOOLEAN, defaultValue: false },
      ],
    },
    REPLACEB: {
      method: 'replaceb',
      parameters: [
        { argumentType: T.STRING },
        { argumentType: T.NUMBER },
        { argumentType: T.NUMBER },
        { argumentType: T.STRING },
      ],
    },
    XLUDFNAME: { method: 'udf', parameters: [{ argumentType: T.ANY }], repeatLastArgs: 1 },
    LINEST: {
      method: 'linest',
      sizeOfResultArrayMethod: 'linestSize',
      vectorizationForbidden: true,
      enableArrayArithmeticForArguments: true,
      parameters: [
        { argumentType: T.RANGE },
        { argumentType: T.RANGE, optionalArg: true },
        { argumentType: T.BOOLEAN, defaultValue: true },
        { argumentType: T.BOOLEAN, defaultValue: false },
      ],
    },
  };
  const tr = { FIXED: 'FIXED', REPLACEB: 'REPLACEB', XLUDFNAME: 'XLUDFNAME', LINEST: 'LINEST' };
  HF.HyperFormula.registerFunctionPlugin(Extra, { enGB: tr, enUS: tr });
}

function verifyWorkbook(WB, XLSX, HF) {
  try {
    registerExtra(HF);
  } catch (e) {}
  const U = XLSX.utils;
  const VOLATIL = /\b(TODAY|NOW|RAND|RANDBETWEEN|RANDARRAY|INDIRECT|OFFSET|CELL|INFO)\s*\(/i;
  const ERR = /^#(REF!|N\/A|VALUE!|DIV\/0!|NAME\?|NUM!|NULL!|SPILL!|CALC!)$/;
  const SEP = /[.,\s  ]/;
  const clean = f =>
    f.replace(/_xlfn\._xlws\.|_xlfn\.|_xlws\./g, '').replace(/_xludf\.[A-Za-z0-9_.]+\(/g, 'XLUDFNAME(');
  const topOf = F => F.split(':')[0].replace(/\$/g, '');
  const sheets = {};
  WB.SheetNames.forEach(sn => {
    const ws = WB.Sheets[sn];
    const data = [];
    if (ws) {
      let mr = -1,
        mc = -1;
      const cells = [];
      for (const k in ws) {
        if (k[0] === '!') continue;
        const a = U.decode_cell(k);
        cells.push([a.r, a.c, ws[k]]);
        if (a.r > mr) mr = a.r;
        if (a.c > mc) mc = a.c;
      }
      for (let r = 0; r <= mr; r++) data.push(new Array(mc + 1).fill(null));
      cells.forEach(([r, c, cl]) => {
        if (!cl) return;
        if (cl.f) {
          const f = clean(cl.f);
          data[r][c] = cl.F && topOf(cl.F) === U.encode_cell({ r, c }) ? '=ARRAYFORMULA(' + f + ')' : '=' + f;
        } else if (cl.F) data[r][c] = null;
        else if (cl.t === 'n') data[r][c] = cl.v;
        else if (cl.t === 'b') data[r][c] = !!cl.v;
        else if (cl.t === 'e') data[r][c] = cl.w || null;
        else if (cl.v === '') data[r][c] = '=""';
        else if (cl.v !== undefined && cl.v !== null) data[r][c] = "'" + String(cl.v);
      });
    }
    sheets[sn] = data;
  });
  let hf;
  try {
    hf = HF.HyperFormula.buildFromSheets(sheets, {
      licenseKey: 'gpl-v3',
      maxRows: 1048576,
      maxColumns: 16384,
      leapYear1900: false,
      nullDate: { year: 1899, month: 12, day: 30 },
      smartRounding: false,
      useArrayArithmetic: false,
      caseSensitive: false,
      useWildcards: true,
      matchWholeCell: true,
      evaluateNullToZero: false,
      precisionEpsilon: 1e-13,
    });
  } catch (e) {
    return { error: e.message, res: {} };
  }
  try {
    ((WB.Workbook && WB.Workbook.Names) || []).forEach(n => {
      try {
        const scope = n.Sheet != null ? hf.getSheetId(WB.SheetNames[n.Sheet]) : undefined;
        hf.addNamedExpression(n.Name, '=' + clean(n.Ref), scope);
      } catch (e) {}
    });
  } catch (e) {}
  const out = {};
  const same = (s, c) => {
    if (c === null || c === undefined) c = typeof s === 'string' ? '' : 0;
    if (typeof s === 'string' && ERR.test(s))
      return typeof c === 'string' && c.toUpperCase() === s.toUpperCase();
    if (typeof s === 'number') {
      if (typeof c === 'boolean') c = c ? 1 : 0;
      if (typeof c !== 'number') return false;
      return Math.abs(s - c) <= 1e-9 * Math.max(1, Math.abs(s), Math.abs(c));
    }
    if (typeof s === 'boolean')
      return (typeof c === 'boolean' && c === s) || (typeof c === 'number' && (c !== 0) === s);
    if (typeof c === 'number' && s !== '' && Number(s) === c) return true;
    s = String(s);
    c = String(c);
    if (s === c) return true;
    if (s.length !== c.length) return false;
    for (let i = 0; i < s.length; i++) if (s[i] !== c[i] && !(SEP.test(s[i]) && SEP.test(c[i]))) return false;
    return true;
  };
  WB.SheetNames.forEach(sn => {
    const ws = WB.Sheets[sn];
    const sid = hf.getSheetId(sn);
    if (!ws || sid === undefined) return;
    for (const k in ws) {
      if (k[0] === '!') continue;
      const cl = ws[k];
      if (!cl || (!cl.f && !cl.F)) continue;
      const key = sn + '!' + k;
      let f = cl.f;
      if (!f && cl.F) {
        const t = ws[topOf(cl.F)];
        f = (t && t.f) || '';
      }
      f = clean(f || '');
      if (cl.v === undefined || cl.v === null) {
        out[key] = { s: 'nv', why: 'sin resultado guardado' };
        continue;
      }
      if (/\[\d+\]/.test(f)) {
        out[key] = { s: 'nv', why: 'vínculo a otro libro' };
        continue;
      }
      const a = U.decode_cell(k);
      let v;
      try {
        v = hf.getCellValue({ sheet: sid, row: a.r, col: a.c });
      } catch (e) {
        out[key] = { s: 'nv', why: 'no se pudo calcular' };
        continue;
      }
      const saved = cl.t === 'e' ? cl.w || '' : cl.v;
      let comp = v;
      if (v && typeof v === 'object' && v.type !== undefined) {
        if ((v.type === 'ERROR' || v.type === 'NAME') && saved !== '#NAME?') {
          out[key] = { s: 'nv', why: 'función no soportada' };
          continue;
        }
        comp = v.value;
      }
      if (VOLATIL.test(f)) {
        out[key] = { s: 'nv', why: 'cambia al recalcular' };
        continue;
      }
      out[key] = { s: same(saved, comp) ? 'ok' : 'bad', calc: comp };
    }
  });
  WB.SheetNames.forEach(sn => {
    const ws = WB.Sheets[sn];
    if (!ws) return;
    for (const k in ws) {
      if (k[0] === '!') continue;
      const cl = ws[k];
      if (!cl || cl.f || !cl.F) continue;
      const A = out[sn + '!' + topOf(cl.F)],
        me = out[sn + '!' + k];
      if (A && A.s === 'nv' && me && me.s !== 'nv') out[sn + '!' + k] = { s: 'nv', why: A.why };
    }
  });
  try {
    hf.destroy();
  } catch (e) {}
  return { res: out };
}
