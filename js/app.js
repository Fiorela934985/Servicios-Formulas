(function () {
  'use strict';
  const $ = id => document.getElementById(id);
  const U = XLSX.utils;
  const PAGE = 1000;
  const ERR = /^#(REF!|N\/A|VALUE!|DIV\/0!|NAME\?|NUM!|NULL!|SPILL!|CALC!)$/;
  let WB = null,
    ALL = [],
    ROWS = [],
    VIEW = [],
    SHOWN = 0,
    VER = null,
    CNT = {},
    filt = 'all',
    sheetFilt = 'all',
    sep = ';';
  const picked = new Set();
  try {
    sep = localStorage.getItem('fx_sep') || ';';
  } catch (e) {}

  const ES = {
    SUM: 'SUMA',
    SUMIF: 'SUMAR.SI',
    SUMIFS: 'SUMAR.SI.CONJUNTO',
    SUMPRODUCT: 'SUMAPRODUCTO',
    SUMSQ: 'SUMA.CUADRADOS',
    AVERAGE: 'PROMEDIO',
    AVERAGEA: 'PROMEDIOA',
    AVERAGEIF: 'PROMEDIO.SI',
    AVERAGEIFS: 'PROMEDIO.SI.CONJUNTO',
    COUNT: 'CONTAR',
    COUNTA: 'CONTARA',
    COUNTIF: 'CONTAR.SI',
    COUNTIFS: 'CONTAR.SI.CONJUNTO',
    COUNTBLANK: 'CONTAR.BLANCO',
    MAXIFS: 'MAX.SI.CONJUNTO',
    MINIFS: 'MIN.SI.CONJUNTO',
    IF: 'SI',
    IFS: 'SI.CONJUNTO',
    IFERROR: 'SI.ERROR',
    IFNA: 'SI.ND',
    AND: 'Y',
    OR: 'O',
    NOT: 'NO',
    XOR: 'XO',
    SWITCH: 'CAMBIAR',
    TRUE: 'VERDADERO',
    FALSE: 'FALSO',
    VLOOKUP: 'BUSCARV',
    HLOOKUP: 'BUSCARH',
    LOOKUP: 'BUSCAR',
    XLOOKUP: 'BUSCARX',
    MATCH: 'COINCIDIR',
    XMATCH: 'COINCIDIRX',
    INDEX: 'INDICE',
    INDIRECT: 'INDIRECTO',
    OFFSET: 'DESREF',
    ROW: 'FILA',
    ROWS: 'FILAS',
    COLUMN: 'COLUMNA',
    COLUMNS: 'COLUMNAS',
    CHOOSE: 'ELEGIR',
    ADDRESS: 'DIRECCION',
    HYPERLINK: 'HIPERVINCULO',
    TRANSPOSE: 'TRANSPONER',
    FILTER: 'FILTRAR',
    SORT: 'ORDENAR',
    SORTBY: 'ORDENARPOR',
    UNIQUE: 'UNICOS',
    SEQUENCE: 'SECUENCIA',
    FORMULATEXT: 'FORMULATEXTO',
    ROUND: 'REDONDEAR',
    ROUNDUP: 'REDONDEAR.MAS',
    ROUNDDOWN: 'REDONDEAR.MENOS',
    MROUND: 'REDOND.MULT',
    INT: 'ENTERO',
    TRUNC: 'TRUNCAR',
    SQRT: 'RAIZ',
    POWER: 'POTENCIA',
    MOD: 'RESIDUO',
    CEILING: 'MULTIPLO.SUPERIOR',
    'CEILING.MATH': 'MULTIPLO.SUPERIOR.MAT',
    FLOOR: 'MULTIPLO.INFERIOR',
    'FLOOR.MATH': 'MULTIPLO.INFERIOR.MAT',
    SIGN: 'SIGNO',
    PRODUCT: 'PRODUCTO',
    QUOTIENT: 'COCIENTE',
    RAND: 'ALEATORIO',
    RANDBETWEEN: 'ALEATORIO.ENTRE',
    DEGREES: 'GRADOS',
    RADIANS: 'RADIANES',
    SIN: 'SENO',
    ASIN: 'ASENO',
    EVEN: 'REDONDEA.PAR',
    ODD: 'REDONDEA.IMPAR',
    ISEVEN: 'ES.PAR',
    ISODD: 'ES.IMPAR',
    SUBTOTAL: 'SUBTOTALES',
    AGGREGATE: 'AGREGAR',
    COMBIN: 'COMBINAT',
    STDEV: 'DESVEST',
    'STDEV.S': 'DESVEST.M',
    'STDEV.P': 'DESVEST.P',
    STDEVA: 'DESVESTA',
    STDEVP: 'DESVESTP',
    MEDIAN: 'MEDIANA',
    MODE: 'MODA',
    'MODE.SNGL': 'MODA.UNO',
    LARGE: 'K.ESIMO.MAYOR',
    SMALL: 'K.ESIMO.MENOR',
    RANK: 'JERARQUIA',
    'RANK.EQ': 'JERARQUIA.EQV',
    SLOPE: 'PENDIENTE',
    INTERCEPT: 'INTERSECCION.EJE',
    FORECAST: 'PRONOSTICO',
    'FORECAST.LINEAR': 'PRONOSTICO.LINEAL',
    LINEST: 'ESTIMACION.LINEAL',
    LOGEST: 'ESTIMACION.LOGARITMICA',
    TREND: 'TENDENCIA',
    GROWTH: 'CRECIMIENTO',
    RSQ: 'COEFICIENTE.R2',
    CORREL: 'COEF.DE.CORREL',
    STEYX: 'ERROR.TIPICO.XY',
    AVEDEV: 'DESVPROM',
    DEVSQ: 'DESVIA2',
    PERCENTILE: 'PERCENTIL',
    QUARTILE: 'CUARTIL',
    TINV: 'DISTR.T.INV',
    'T.INV': 'INV.T',
    'T.INV.2T': 'INV.T.2C',
    'T.DIST.2T': 'DISTR.T.2C',
    'NORM.S.INV': 'INV.NORM.ESTAND',
    'NORM.INV': 'INV.NORM',
    'NORM.DIST': 'DISTR.NORM.N',
    NORMSINV: 'DISTR.NORM.ESTAND.INV',
    CHIINV: 'PRUEBA.CHI.INV',
    CONFIDENCE: 'INTERVALO.CONFIANZA',
    MINVERSE: 'MINVERSA',
    FREQUENCY: 'FRECUENCIA',
    CONCATENATE: 'CONCATENAR',
    TEXTJOIN: 'UNIRCADENAS',
    TEXT: 'TEXTO',
    LEFT: 'IZQUIERDA',
    RIGHT: 'DERECHA',
    MID: 'EXTRAE',
    LEN: 'LARGO',
    FIND: 'ENCONTRAR',
    SEARCH: 'HALLAR',
    SUBSTITUTE: 'SUSTITUIR',
    REPLACE: 'REEMPLAZAR',
    REPLACEB: 'REEMPLAZARB',
    TRIM: 'ESPACIOS',
    CLEAN: 'LIMPIAR',
    UPPER: 'MAYUSC',
    LOWER: 'MINUSC',
    PROPER: 'NOMPROPIO',
    VALUE: 'VALOR',
    REPT: 'REPETIR',
    CHAR: 'CARACTER',
    CODE: 'CODIGO',
    EXACT: 'IGUAL',
    FIXED: 'DECIMAL',
    DOLLAR: 'MONEDA',
    NUMBERVALUE: 'VALOR.NUMERO',
    TEXTBEFORE: 'TEXTOANTES',
    TEXTAFTER: 'TEXTODESPUES',
    TEXTSPLIT: 'DIVIDIRTEXTO',
    ISBLANK: 'ESBLANCO',
    ISERROR: 'ESERROR',
    ISERR: 'ESERR',
    ISNUMBER: 'ESNUMERO',
    ISTEXT: 'ESTEXTO',
    ISNA: 'ESNOD',
    ISLOGICAL: 'ESLOGICO',
    ISFORMULA: 'ESFORMULA',
    NA: 'NOD',
    CELL: 'CELDA',
    TODAY: 'HOY',
    NOW: 'AHORA',
    DATE: 'FECHA',
    TIME: 'NSHORA',
    YEAR: 'AÑO',
    MONTH: 'MES',
    DAY: 'DIA',
    HOUR: 'HORA',
    MINUTE: 'MINUTO',
    SECOND: 'SEGUNDO',
    WEEKDAY: 'DIASEM',
    WEEKNUM: 'NUM.DE.SEMANA',
    EDATE: 'FECHA.MES',
    EOMONTH: 'FIN.MES',
    DATEDIF: 'SIFECHA',
    NETWORKDAYS: 'DIAS.LAB',
    WORKDAY: 'DIA.LAB',
    DAYS: 'DIAS',
    DATEVALUE: 'FECHANUMERO',
    PMT: 'PAGO',
    NPV: 'VNA',
    IRR: 'TIR',
    FV: 'VF',
    PV: 'VA',
    RATE: 'TASA',
  };

  function toES(f) {
    let o = '',
      i = 0,
      br = 0;
    const n = f.length;
    while (i < n) {
      const ch = f[i];
      if (ch === '"' || ch === "'") {
        let j = i + 1;
        while (j < n) {
          if (f[j] === ch) {
            if (f[j + 1] === ch) {
              j += 2;
              continue;
            }
            break;
          }
          j++;
        }
        o += f.slice(i, j + 1);
        i = j + 1;
        continue;
      }
      if (ch === '[') {
        let d = 0,
          j = i;
        for (; j < n; j++) {
          if (f[j] === '[') d++;
          else if (f[j] === ']') {
            d--;
            if (!d) break;
          }
        }
        o += f.slice(i, j + 1);
        i = j + 1;
        continue;
      }
      if (ch === '{') {
        br++;
        o += ch;
        i++;
        continue;
      }
      if (ch === '}') {
        br--;
        o += ch;
        i++;
        continue;
      }
      if (ch === ',' && br === 0) {
        o += sep;
        i++;
        continue;
      }
      if (/[A-Za-z_]/.test(ch) && (i === 0 || !/[A-Za-z0-9_.$]/.test(f[i - 1]))) {
        const w = /^[A-Za-z_][A-Za-z0-9_.]*/.exec(f.slice(i))[0],
          nx = f[i + w.length],
          W = w.toUpperCase();
        if (nx === '(' && ES[W]) {
          o += ES[W];
          i += w.length;
          continue;
        }
        if ((W === 'TRUE' || W === 'FALSE') && nx !== '(' && nx !== '!') {
          o += ES[W];
          i += w.length;
          continue;
        }
        o += w;
        i += w.length;
        continue;
      }
      o += ch;
      i++;
    }
    return o;
  }
  const clean = f => f.replace(/_xlfn\._xlws\.|_xlfn\.|_xlws\./g, '');
  function shown(c) {
    if (!c || c.v === undefined || c.v === null) return '';
    if (c.t === 'b') return c.v ? 'VERDADERO' : 'FALSO';
    if (c.w !== undefined && c.w !== null) return String(c.w);
    return String(c.v);
  }
  function isErr(c) {
    return c.t === 'e' || ERR.test(String(c.v));
  }

  const inp = $('file');
  const drop = $('drop');
  $('pick').onclick = () => inp.click();
  drop.onclick = () => inp.click();
  inp.onchange = () => {
    if (inp.files[0]) load(inp.files[0]);
    inp.value = '';
  };
  ['dragenter', 'dragover'].forEach(e => document.addEventListener(e, ev => ev.preventDefault()));
  document.addEventListener('drop', ev => {
    ev.preventDefault();
    const f = ev.dataTransfer && ev.dataTransfer.files[0];
    if (f) load(f);
  });

  function load(file) {
    const msg = $('loadMsg');
    msg.textContent = 'Leyendo ' + file.name + '…';
    msg.classList.remove('hidden');
    $('step2').classList.add('hidden');
    $('step3').classList.add('hidden');
    const r = new FileReader();
    r.onerror = () => {
      msg.textContent = 'No se pudo leer el archivo.';
    };
    r.onload = () =>
      setTimeout(() => {
        try {
          WB = XLSX.read(new Uint8Array(r.result), {
            type: 'array',
            cellFormula: true,
            cellNF: true,
            cellText: true,
            cellDates: false,
          });
        } catch (e) {
          msg.textContent =
            'No se pudo leer el archivo. Si tiene contraseña, quítasela y vuelve a intentarlo.';
          WB = null;
          return;
        }
        msg.classList.add('hidden');
        $('hero').classList.add('hidden');
        $('filebar').classList.remove('hidden');
        $('fname').textContent = file.name;
        VER = null;
        filt = 'all';
        sheetFilt = 'all';
        index();
        picked.clear();
        WB.SheetNames.forEach(s => {
          if (CNT[s]) picked.add(s);
        });
        renderSheets();
        $('fmeta').textContent =
          size(file.size) + ' · ' + WB.SheetNames.length + ' hojas · ' + ALL.length + ' fórmulas';
        $('step2').classList.remove('hidden');
      }, 20);
    r.readAsArrayBuffer(file);
  }

  function index() {
    ALL = [];
    CNT = {};
    WB.SheetNames.forEach(sn => {
      const ws = WB.Sheets[sn];
      if (!ws) return;
      const l = [];
      for (const k in ws) {
        if (k[0] === '!') continue;
        const c = ws[k];
        if (c && (c.f || c.F)) {
          const a = U.decode_cell(k);
          l.push([a.r, a.c, k, c]);
        }
      }
      l.sort((x, y) => x[0] - y[0] || x[1] - y[1]);
      l.forEach(([, , k, c]) => {
        let f = c.f;
        if (!f) {
          const t = ws[c.F.split(':')[0].replace(/\$/g, '')];
          f = t && t.f ? t.f : '';
        }
        ALL.push({ sheet: sn, addr: k, cell: c, f: clean(f || '') });
      });
      CNT[sn] = l.length;
    });
  }

  function renderSheets() {
    const names = WB.SheetNames,
      w = names.filter(s => CNT[s]).length;
    const lbl = {
      all: 'Todas <b>' + names.length + '</b>',
      with: 'Con fórmulas <b>' + w + '</b>',
      without: 'Sin fórmulas <b>' + (names.length - w) + '</b>',
    };
    [...$('sheetTabs').children].forEach(b => {
      b.innerHTML = lbl[b.dataset.k];
      b.classList.toggle('on', b.dataset.k === sheetFilt);
    });
    const vis = names.filter(s => sheetFilt === 'all' || (sheetFilt === 'with' ? CNT[s] : !CNT[s]));
    $('chips').innerHTML = vis.length
      ? vis
          .map(
            s =>
              '<button class="chip' +
              (picked.has(s) ? ' on' : '') +
              '" data-s="' +
              esc(s) +
              '">' +
              esc(s) +
              ' <i>(' +
              (CNT[s] || 0) +
              ')</i></button>',
          )
          .join('')
      : '<span class="hint">No hay hojas en este grupo.</span>';
    [...$('chips').querySelectorAll('.chip')].forEach(
      b =>
        (b.onclick = () => {
          const s = b.dataset.s;
          if (picked.has(s)) picked.delete(s);
          else picked.add(s);
          b.classList.toggle('on', picked.has(s));
          selInfo();
        }),
    );
    selInfo();
  }
  [...$('sheetTabs').children].forEach(
    b =>
      (b.onclick = () => {
        sheetFilt = b.dataset.k;
        renderSheets();
      }),
  );
  function selInfo() {
    const n = picked.size,
      f = [...picked].reduce((a, s) => a + (CNT[s] || 0), 0);
    $('selInfo').textContent = n
      ? n + ' hoja' + (n > 1 ? 's' : '') + ' seleccionada' + (n > 1 ? 's' : '') + ' · ' + f + ' fórmulas'
      : 'Selecciona al menos una hoja';
    $('run').disabled = !n;
  }

  $('run').onclick = () => {
    ROWS = ALL.filter(r => picked.has(r.sheet));
    filt = 'all';
    if (VER) {
      show();
      return;
    }
    $('run').disabled = true;
    $('run').textContent = 'Verificando…';
    setTimeout(() => {
      try {
        VER = verifyWorkbook(WB, XLSX, HyperFormula).res || {};
      } catch (e) {
        VER = {};
      }
      $('run').disabled = false;
      $('run').textContent = 'Extraer';
      show();
    }, 30);
  };

  function status(r) {
    const c = r.cell;
    if (isErr(c)) return { s: 'bad', why: 'la fórmula da ' + (c.t === 'e' ? c.w : c.v) };
    return VER[r.sheet + '!' + r.addr] || { s: 'nv', why: 'no se pudo verificar' };
  }
  const LBL = { ok: 'Correcto', bad: 'Incorrecto', nv: 'No verificable' };
  function calcTxt(v) {
    if (v === null || v === undefined) return '(vacío)';
    if (typeof v === 'boolean') return v ? 'VERDADERO' : 'FALSO';
    if (typeof v === 'number') return String(+v.toPrecision(12));
    return String(v);
  }
  function detail(st) {
    if (st.s === 'bad') return st.why || 'debería dar ' + calcTxt(st.calc);
    if (st.s === 'nv') return st.why || '';
    return '';
  }
  function verFull(st) {
    const d = detail(st);
    return LBL[st.s] + (d ? ' (' + d + ')' : '');
  }

  function show() {
    const n = { all: ROWS.length, ok: 0, bad: 0, nv: 0 };
    ROWS.forEach(r => n[status(r).s]++);
    const items = [
      ['all', 'Todas'],
      ['ok', 'Correctas'],
      ['bad', 'Incorrectas'],
      ['nv', 'No verificables'],
    ];
    $('counts').innerHTML = items
      .filter(([k]) => k === 'all' || n[k])
      .map(
        ([k, l]) =>
          '<button class="cnt' +
          (filt === k ? ' on' : '') +
          '" data-k="' +
          k +
          '"><span class="d ' +
          k +
          '"></span>' +
          l +
          ' <b>' +
          n[k] +
          '</b></button>',
      )
      .join('');
    [...$('counts').children].forEach(
      b =>
        (b.onclick = () => {
          filt = b.dataset.k;
          [...$('counts').children].forEach(x => x.classList.toggle('on', x === b));
          apply();
        }),
    );
    $('step3').classList.remove('hidden');
    setSep();
    apply();
    $('step3').scrollIntoView({ behavior: 'smooth', block: 'start' });
  }
  function setSep() {
    [...$('sep').children].forEach(b => b.classList.toggle('on', b.dataset.v === sep));
  }
  [...$('sep').children].forEach(
    b =>
      (b.onclick = () => {
        sep = b.dataset.v;
        try {
          localStorage.setItem('fx_sep', sep);
        } catch (e) {}
        setSep();
        apply();
      }),
  );

  function fx(r) {
    return r.f ? '=' + toES(r.f) : '';
  }
  function apply() {
    VIEW = filt === 'all' ? ROWS : ROWS.filter(r => status(r).s === filt);
    $('tb').innerHTML = '';
    SHOWN = 0;
    more();
    $('copy').textContent = 'Copiar tabla (' + VIEW.length + ')';
    $('copy').disabled = !VIEW.length;
  }
  function more() {
    $('tb').insertAdjacentHTML(
      'beforeend',
      VIEW.slice(SHOWN, SHOWN + PAGE)
        .map(r => {
          const st = status(r),
            d = detail(st);
          return (
            '<tr><td class="h">' +
            esc(r.sheet) +
            '</td><td class="c">' +
            esc(r.addr) +
            '</td><td class="fx">' +
            esc(fx(r)) +
            '</td><td class="r' +
            (isErr(r.cell) ? ' err' : '') +
            '">' +
            esc(shown(r.cell)) +
            '</td><td><span class="bd ' +
            st.s +
            '">' +
            LBL[st.s] +
            '</span>' +
            (d ? '<span class="why">' + esc(d) + '</span>' : '') +
            '</td></tr>'
          );
        })
        .join(''),
    );
    SHOWN = Math.min(SHOWN + PAGE, VIEW.length);
    $('moreBox').classList.toggle('hidden', SHOWN >= VIEW.length);
  }
  $('more').onclick = more;

  $('copy').onclick = async () => {
    const H = ['Hoja', 'Celda', 'Fórmula', 'Resultado', 'Verificación'];
    const rows = VIEW.map(r => {
      const st = status(r);
      return { v: [r.sheet, r.addr, fx(r), shown(r.cell), verFull(st)], s: st.s, e: isErr(r.cell) };
    });
    const tsv = [H]
      .concat(rows.map(x => x.v))
      .map(a =>
        a
          .map(v => {
            v = String(v);
            return /[\t\n\r"]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v;
          })
          .join('\t'),
      )
      .join('\r\n');
    const base =
      "border:1px solid #808080;padding:3px 6px;font-family:Arial;font-size:10pt;vertical-align:top;mso-number-format:'\\@';";
    const col = { ok: '#1e7b3a', bad: '#c8102e', nv: '#9a5b00' };
    const html =
      '<meta charset="utf-8"><table style="border-collapse:collapse;font-family:Arial;font-size:10pt">' +
      '<tr>' +
      H.map(
        h => '<th style="' + base + 'background:#f2f2f2;font-weight:bold;text-align:left">' + h + '</th>',
      ).join('') +
      '</tr>' +
      rows
        .map(
          x =>
            '<tr>' +
            x.v
              .map((v, i) => {
                let st = base;
                if (i === 1) st += 'font-weight:bold;';
                if (i === 3 && x.e) st += 'color:#c8102e;';
                if (i === 4) st += 'color:' + col[x.s] + ';font-weight:bold;';
                return (
                  '<td style="' +
                  st +
                  '">' +
                  esc(String(v)).replace(/\n/g, '<br style="mso-data-placement:same-cell">') +
                  '</td>'
                );
              })
              .join('') +
            '</tr>',
        )
        .join('') +
      '</table>';
    try {
      if (window.ClipboardItem && navigator.clipboard && navigator.clipboard.write)
        await navigator.clipboard.write([
          new ClipboardItem({
            'text/plain': new Blob([tsv], { type: 'text/plain' }),
            'text/html': new Blob([html], { type: 'text/html' }),
          }),
        ]);
      else await navigator.clipboard.writeText(tsv);
      toast('Tabla copiada: ' + rows.length + ' filas');
    } catch (e) {
      if (copyHtmlFallback(html)) {
        toast('Tabla copiada: ' + rows.length + ' filas');
        return;
      }
      $('mdTxt').value = tsv;
      $('md').classList.remove('hidden');
      setTimeout(() => {
        $('mdTxt').focus();
        $('mdTxt').select();
      }, 50);
    }
  };
  function copyHtmlFallback(html) {
    try {
      const d = document.createElement('div');
      d.style.cssText = 'position:fixed;left:-9999px;top:0';
      d.innerHTML = html;
      document.body.appendChild(d);
      const rg = document.createRange();
      rg.selectNodeContents(d);
      const sel = getSelection();
      sel.removeAllRanges();
      sel.addRange(rg);
      const ok = document.execCommand('copy');
      sel.removeAllRanges();
      d.remove();
      return ok;
    } catch (e) {
      return false;
    }
  }
  $('mdClose').onclick = () => $('md').classList.add('hidden');
  $('mdSel').onclick = () => {
    $('mdTxt').focus();
    $('mdTxt').select();
  };

  function size(b) {
    return b > 1048576 ? (b / 1048576).toFixed(1) + ' MB' : Math.max(1, Math.round(b / 1024)) + ' KB';
  }
  function toast(t) {
    const e = $('toast');
    e.textContent = t;
    e.classList.add('show');
    clearTimeout(e._t);
    e._t = setTimeout(() => e.classList.remove('show'), 2500);
  }
  function esc(s) {
    return String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]);
  }
})();
