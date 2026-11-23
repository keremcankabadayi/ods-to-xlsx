/*
 * ODS (OpenDocument Spreadsheet) → XLSX dönüştürücü.
 *
 * ODS paketindeki content.xml / styles.xml / settings.xml dosyalarını okuyup ExcelJS ile xlsx üretir.
 * Aktarılanlar: değerler, formüller (ODF → Excel sözdizimi), hücre stilleri (yazı tipi, dolgu, kenarlık,
 * hizalama, sayı biçimi, koruma), birleştirilmiş hücreler, sütun genişlikleri, satır yükseklikleri,
 * gizli satır/sütunlar, satır grupları, veri doğrulama (açılır listeler), köprüler, notlar, resimler
 * (png/jpeg/gif), tanımlı adlar, yazdırma alanları, sayfa yapısı, sayfa koruması, sekme renkleri,
 * gizli sayfalar, bölmeleri dondurma, yakınlaştırma.
 * Aktarılmayanlar: makrolar, grafikler, çizim nesneleri, koşullu biçimlendirme, üstbilgi/altbilgi.
 *
 * Bağımlılıklar: ExcelJS, DOMParser (tarayıcıda yerleşik; Node'da @xmldom/xmldom ile verilebilir).
 */
(function (root) {
  'use strict';

  // ---------------------------------------------------------------- yardımcılar

  function elements(node) {
    const out = [];
    for (let n = node.firstChild; n; n = n.nextSibling) if (n.nodeType === 1) out.push(n);
    return out;
  }

  function child(node, name) {
    for (let n = node.firstChild; n; n = n.nextSibling) if (n.nodeType === 1 && n.nodeName === name) return n;
    return null;
  }

  function descendants(node, name) {
    return Array.from(node.getElementsByTagName(name));
  }

  function attr(node, name) {
    return node.hasAttribute(name) ? node.getAttribute(name) : null;
  }

  function attrMap(node) {
    const out = {};
    if (!node) return out;
    for (let i = 0; i < node.attributes.length; i++) {
      const a = node.attributes[i];
      out[a.name] = a.value;
    }
    return out;
  }

  // Uzunluğu punto (pt) cinsine çevirir.
  function toPt(v) {
    if (v == null) return null;
    const m = /^(-?[\d.]+)\s*(pt|cm|mm|in|inch|pc|px)?$/.exec(String(v).trim());
    if (!m) return null;
    const n = parseFloat(m[1]);
    switch (m[2]) {
      case 'cm': return n * 72 / 2.54;
      case 'mm': return n * 72 / 25.4;
      case 'in': case 'inch': return n * 72;
      case 'pc': return n * 12;
      case 'px': return n * 0.75;
      default: return n;
    }
  }

  function argb(color) {
    if (!color || color === 'transparent') return null;
    const m = /^#([0-9a-f]{6})$/i.exec(color.trim());
    return m ? 'FF' + m[1].toUpperCase() : null;
  }

  function colName(c) { // 1 tabanlı
    let s = '';
    while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); }
    return s;
  }

  function colIndex(letters) {
    let n = 0;
    for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n;
  }

  function quoteSheet(name) {
    if (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name) && !/^(R|C|RC)\d*$/i.test(name)) return name;
    return "'" + name.replace(/'/g, "''") + "'";
  }

  // Boşlukla ayrılmış aralık listesini böler; tırnaklı sayfa adlarındaki boşluklara dokunmaz.
  function splitRanges(s) {
    const out = [];
    let cur = '', q = false;
    for (const ch of s) {
      if (ch === "'") q = !q;
      if (ch === ' ' && !q) { if (cur) out.push(cur); cur = ''; continue; }
      cur += ch;
    }
    if (cur) out.push(cur);
    return out;
  }

  // "$A$1" biçimine zorlar (tanımlı adlar ve yazdırma alanları için)
  function absolute(cells) {
    return cells.split(':').map((p) => p.replace(/^\$?([A-Z]+)\$?(\d+)$/i, '$$$1$$$2')).join(':');
  }

  // ---------------------------------------------------------------- formüller

  // Excel 2010 sonrası işlevler dosyada "_xlfn." önekiyle saklanır.
  const XLFN = new Set(('AGGREGATE BETA.DIST BETA.INV BINOM.DIST BINOM.INV BINOM.DIST.RANGE CEILING.MATH CEILING.PRECISE ' +
    'CHISQ.DIST CHISQ.DIST.RT CHISQ.INV CHISQ.INV.RT CHISQ.TEST CONCAT CONFIDENCE.NORM CONFIDENCE.T COVARIANCE.P ' +
    'COVARIANCE.S DAYS ERF.PRECISE ERFC.PRECISE EXPON.DIST F.DIST F.DIST.RT F.INV F.INV.RT F.TEST FILTER FLOOR.MATH ' +
    'FLOOR.PRECISE FORECAST.LINEAR FORMULATEXT GAMMA GAMMA.DIST GAMMA.INV GAMMALN.PRECISE HYPGEOM.DIST IFNA IFS ' +
    'ISFORMULA ISOWEEKNUM LET LOGNORM.DIST LOGNORM.INV MAXIFS MINIFS MODE.MULT MODE.SNGL NEGBINOM.DIST ' +
    'NETWORKDAYS.INTL NORM.DIST NORM.INV NORM.S.DIST NORM.S.INV NUMBERVALUE PERCENTILE.EXC PERCENTILE.INC ' +
    'PERCENTRANK.EXC PERCENTRANK.INC POISSON.DIST QUARTILE.EXC QUARTILE.INC RANK.AVG RANK.EQ SEQUENCE SHEET SHEETS ' +
    'SORT SORTBY STDEV.P STDEV.S SWITCH T.DIST T.DIST.2T T.DIST.RT T.INV T.INV.2T T.TEST TEXTJOIN UNIQUE VAR.P VAR.S ' +
    'WEIBULL.DIST WORKDAY.INTL XLOOKUP XMATCH XOR Z.TEST DECIMAL BASE ARABIC BITAND BITOR BITXOR BITLSHIFT BITRSHIFT ' +
    'COT COTH CSC CSCH SEC SECH ACOT ACOTH COMBINA PERMUTATIONA PHI GAUSS SKEW.P RRI PDURATION UNICHAR UNICODE ' +
    'ENCODEURL FILTERXML WEBSERVICE IMCOSH IMCOT IMCSC IMCSCH IMSEC IMSECH IMSINH IMTAN').split(' '));

  const RENAME = { 'CEILING.XCL': 'CEILING', 'FLOOR.XCL': 'FLOOR', 'ERRORTYPE': 'ERROR.TYPE', 'MULTIPLE.OPERATIONS': 'TABLE' };

  function functionName(name) {
    let n = name.toUpperCase().replace(/^(COM\.MICROSOFT\.|ORG\.OPENOFFICE\.|LEGACY\.|_XLFN\.)/, '');
    if (RENAME[n]) n = RENAME[n];
    return XLFN.has(n) ? '_xlfn.' + n : n;
  }

  // "[$Sayfa.$A$1:.$B$2]" köşeli parantez içini Excel referansına çevirir.
  function translateRef(inner) {
    const parts = [];
    let cur = '', q = false;
    for (let i = 0; i < inner.length; i++) {
      const ch = inner[i];
      if (ch === "'") { q = !q; cur += ch; continue; }
      if (ch === ':' && !q) { parts.push(cur); cur = ''; continue; }
      cur += ch;
    }
    parts.push(cur);

    const parsed = parts.map((p) => {
      let dot = -1; q = false;
      for (let i = 0; i < p.length; i++) {
        if (p[i] === "'") q = !q;
        else if (p[i] === '.' && !q) dot = i;
      }
      let sheet = null, cell = p;
      if (dot >= 0) {
        sheet = p.slice(0, dot).replace(/^\$/, '');
        cell = p.slice(dot + 1);
        if (sheet.startsWith("'")) sheet = sheet.slice(1, -1).replace(/''/g, "'");
        if (sheet === '') sheet = null;
      }
      return { sheet, cell };
    });

    if (parsed.some((p) => /#REF!/i.test(p.cell) || (p.sheet && /#REF!/i.test(p.sheet)))) return '#REF!';
    if (parsed.some((p) => p.sheet && /^file:|#\$/.test(p.sheet))) return '#REF!'; // dış dosya bağlantısı
    const cells = parsed.map((p) => p.cell).join(':');
    const s1 = parsed[0].sheet;
    const s2 = parsed.length > 1 ? parsed[parsed.length - 1].sheet : null;
    if (s1 && s2 && s2 !== s1) return quoteSheet(s1 + ':' + s2) + '!' + cells;
    if (s1) return quoteSheet(s1) + '!' + cells;
    return cells;
  }

  // "of:=SUM([.A1:.A5];1)" → "SUM(A1:A5,1)"
  function translateFormula(f) {
    if (f == null) return null;
    let src = f;
    const pm = /^([a-z]+):=?/i.exec(src);
    if (pm) {
      if (pm[1].toLowerCase() === 'msoxl') return src.slice(pm[0].length);
      src = src.slice(pm[0].length);
    } else if (src.startsWith('=')) {
      src = src.slice(1);
    }
    let out = '';
    let brace = 0;
    for (let i = 0; i < src.length; i++) {
      const ch = src[i];
      if (ch === '"') {
        let j = i + 1;
        while (j < src.length) {
          if (src[j] === '"') { if (src[j + 1] === '"') { j += 2; continue; } break; }
          j++;
        }
        out += src.slice(i, j + 1);
        i = j;
        continue;
      }
      if (ch === '[') {
        let j = i + 1, q = false;
        while (j < src.length && (src[j] !== ']' || q)) { if (src[j] === "'") q = !q; j++; }
        out += translateRef(src.slice(i + 1, j));
        i = j;
        continue;
      }
      if (ch === '{') { brace++; out += ch; continue; }
      if (ch === '}') { brace--; out += ch; continue; }
      if (ch === ';') { out += ','; continue; }
      if (ch === '|' && brace > 0) { out += ';'; continue; }
      if (ch === '~') { out += ','; continue; }
      if (ch === '!') { out += ' '; continue; }
      if (/[A-Za-z_À-￿]/.test(ch)) {
        let j = i;
        while (j < src.length && /[A-Za-z0-9_.À-￿]/.test(src[j])) j++;
        const word = src.slice(i, j);
        let k = j;
        while (src[k] === ' ') k++;
        out += src[k] === '(' ? functionName(word) : word;
        i = j - 1;
        continue;
      }
      out += ch;
    }
    return out;
  }

  // ---------------------------------------------------------------- sayı biçimleri

  const SAFE_LITERAL = /^[ .,\-/:()$+]*$/;

  function literal(text, allowPercent) {
    if (!text) return '';
    if (SAFE_LITERAL.test(text)) return text;
    if (allowPercent && text === '%') return '%';
    return '"' + text.replace(/"/g, '\\"') + '"';
  }

  const LCID = { tr: '41F', 'tr-TR': '41F', en: '409', 'en-US': '409', 'en-GB': '809', de: '407', 'de-DE': '407', fr: '40C', 'fr-FR': '40C', es: '40A', it: '410', ru: '419', ar: '401', nl: '413', pt: '816', 'pt-BR': '416' };

  const NAMED_COLORS = { FF000000: 'Black', FFFFFFFF: 'White', FFFF0000: 'Red', FF00FF00: 'Green', FF0000FF: 'Blue', FFFFFF00: 'Yellow', FFFF00FF: 'Magenta', FF00FFFF: 'Cyan' };

  function numberPart(n) {
    const dec = attr(n, 'number:decimal-places');
    const minDec = attr(n, 'number:min-decimal-places');
    const minInt = parseInt(attr(n, 'number:min-integer-digits') || '0', 10);
    let s = minInt > 0 ? '0'.repeat(minInt) : '#';
    if (attr(n, 'number:grouping') === 'true') s = '#,' + '#'.repeat(Math.max(0, 3 - minInt)) + '0'.repeat(Math.min(3, Math.max(minInt, 1)));
    const d = dec == null ? 0 : parseInt(dec, 10);
    const md = minDec == null ? d : parseInt(minDec, 10);
    if (d > 0) s += '.' + '0'.repeat(md) + '#'.repeat(Math.max(0, d - md));
    const factor = parseFloat(attr(n, 'number:display-factor') || '1');
    if (factor > 1) s += ','.repeat(Math.round(Math.log10(factor) / 3));
    return s;
  }

  function formatOf(styleEl, numStyles, depth) {
    const kind = styleEl.nodeName;
    if (kind === 'number:boolean-style') return 'General';
    let fmt = '';
    let color = '';
    const isPercent = kind === 'number:percentage-style';
    for (const n of elements(styleEl)) {
      switch (n.nodeName) {
        case 'number:number':
          if (kind === 'number:number-style' && attr(n, 'number:decimal-places') == null &&
              attr(n, 'number:grouping') == null && attr(n, 'number:min-decimal-places') == null) {
            fmt += 'General';
          } else {
            fmt += numberPart(n);
          }
          break;
        case 'number:scientific-number': {
          const d = parseInt(attr(n, 'number:decimal-places') || '0', 10);
          const minInt = parseInt(attr(n, 'number:min-integer-digits') || '1', 10);
          const e = parseInt(attr(n, 'number:min-exponent-digits') || '2', 10);
          fmt += '0'.repeat(Math.max(1, minInt)) + (d ? '.' + '0'.repeat(d) : '') + 'E+' + '0'.repeat(e);
          break;
        }
        case 'number:fraction': {
          const minInt = parseInt(attr(n, 'number:min-integer-digits') || '0', 10);
          const num = parseInt(attr(n, 'number:min-numerator-digits') || '1', 10);
          const den = attr(n, 'number:denominator-value');
          const denDigits = parseInt(attr(n, 'number:min-denominator-digits') || '1', 10);
          fmt += (minInt ? '0'.repeat(minInt) : '#') + ' ' + '?'.repeat(num) + '/' + (den || '?'.repeat(denDigits));
          break;
        }
        case 'number:currency-symbol': fmt += literal(n.textContent); break;
        case 'number:text': fmt += literal(n.textContent, isPercent); break;
        case 'number:text-content': fmt += '@'; break;
        case 'number:fill-character': fmt += '*' + (n.textContent || ' '); break;
        case 'number:day': fmt += attr(n, 'number:style') === 'long' ? 'dd' : 'd'; break;
        case 'number:month': {
          const textual = attr(n, 'number:textual') === 'true';
          const long = attr(n, 'number:style') === 'long';
          fmt += textual ? (long ? 'mmmm' : 'mmm') : (long ? 'mm' : 'm');
          break;
        }
        case 'number:year': fmt += attr(n, 'number:style') === 'long' ? 'yyyy' : 'yy'; break;
        case 'number:day-of-week': fmt += attr(n, 'number:style') === 'long' ? 'dddd' : 'ddd'; break;
        case 'number:hours': {
          const h = attr(n, 'number:style') === 'long' ? 'hh' : 'h';
          fmt += attr(styleEl, 'number:truncate-on-overflow') === 'false' ? '[' + h + ']' : h;
          break;
        }
        case 'number:minutes': fmt += attr(n, 'number:style') === 'long' ? 'mm' : 'm'; break;
        case 'number:seconds': {
          fmt += attr(n, 'number:style') === 'long' ? 'ss' : 's';
          const d = parseInt(attr(n, 'number:decimal-places') || '0', 10);
          if (d) fmt += '.' + '0'.repeat(d);
          break;
        }
        case 'number:am-pm': fmt += 'AM/PM'; break;
        case 'style:text-properties': {
          const c = argb(attr(n, 'fo:color'));
          if (c && NAMED_COLORS[c]) color = '[' + NAMED_COLORS[c] + ']';
          break;
        }
        default: break;
      }
    }
    if (!fmt) fmt = 'General';
    if (/date-style$/.test(kind) && /mmm|ddd/.test(fmt)) {
      const lcid = LCID[(attr(styleEl, 'number:language') || '') + '-' + (attr(styleEl, 'number:country') || '')] ||
        LCID[attr(styleEl, 'number:language') || ''];
      if (lcid) fmt = '[$-' + lcid + ']' + fmt;
    }
    let own = color + fmt;

    // style:map → koşullu bölümler
    const maps = elements(styleEl).filter((n) => n.nodeName === 'style:map');
    if (maps.length && depth < 3) {
      const sections = maps.map((m) => {
        const cond = (attr(m, 'style:condition') || '').replace(/^value\(\)/, '').trim();
        const target = numStyles.get(attr(m, 'style:apply-style-name'));
        return { cond, fmt: target ? formatOf(target, numStyles, depth + 1) : 'General' };
      });
      // LibreOffice metin bölümü eklemek için "her zaman doğru" koşul kullanır (ör. value()<=1.79E+308)
      for (const sec of sections) if (/^(<=\s*1\.79769|>=\s*-1\.79769)/.test(sec.cond)) sec.cond = '';
      if (sections.every((sec) => sec.cond === '')) return sections.map((sec) => sec.fmt).concat(own).join(';');
      const conds = sections.map((s) => s.cond).join('|');
      if (conds === '>=0' || conds === '>0') own = sections[0].fmt + ';' + own;
      else if (conds === '>0|<0') own = sections[0].fmt + ';' + sections[1].fmt + ';' + own;
      else own = sections.slice(0, 2).map((s) => '[' + s.cond + ']' + s.fmt).join(';') + ';' + own;
    }
    return own;
  }

  // ---------------------------------------------------------------- stiller

  const BORDER_SIDES = { top: 'fo:border-top', bottom: 'fo:border-bottom', left: 'fo:border-left', right: 'fo:border-right' };

  function parseBorder(v) {
    if (!v || v === 'none' || v === 'hidden') return null;
    let width = 0.75, style = 'solid', color = 'FF000000';
    for (const p of v.trim().split(/\s+/)) {
      if (/^-?[\d.]/.test(p)) width = toPt(p) || 0;
      else if (/^#/.test(p)) color = argb(p) || color;
      else style = p;
    }
    if (width <= 0 || style === 'none' || style === 'hidden') return null;
    let s;
    if (style === 'double' || style === 'double-thin') s = 'double';
    else if (style === 'dotted') s = width >= 1.5 ? 'mediumDashDot' : 'dotted';
    else if (style === 'dashed' || style === 'long-dash') s = width >= 1.5 ? 'mediumDashed' : width < 0.7 ? 'dotted' : 'dashed';
    else if (style === 'fine-dashed') s = width >= 1.5 ? 'mediumDashDot' : 'dashed';
    else if (style === 'dash-dot') s = width >= 1.5 ? 'mediumDashDot' : width < 0.7 ? 'dashed' : 'dashDot';
    else if (style === 'dash-dot-dot') s = width >= 1.5 ? 'mediumDashDotDot' : width < 0.7 ? 'dashed' : 'dashDotDot';
    // düz çizgi: LibreOffice'in aralıkları — <0.7pt ince çizgi (hair), <1.5 thin, <2.5 medium, üstü thick
    else if (width < 0.7) s = 'hair';
    else if (width < 1.5) s = 'thin';
    else if (width < 2.5) s = 'medium';
    else s = 'thick';
    return { style: s, color: { argb: color } };
  }

  const PROP_KEYS = {
    'style:table-cell-properties': 'cell',
    'style:paragraph-properties': 'para',
    'style:text-properties': 'text',
    'style:table-column-properties': 'col',
    'style:table-row-properties': 'row',
    'style:table-properties': 'table',
  };
  const NUMBER_STYLE = /^number:(number|currency|percentage|date|time|boolean|text)-style$/;

  function buildStyles(docs) {
    const fonts = new Map();
    const numStyles = new Map();
    const raw = new Map();
    const defaults = new Map();

    function readProps(el) {
      const props = {};
      for (const p of elements(el)) if (PROP_KEYS[p.nodeName]) props[PROP_KEYS[p.nodeName]] = attrMap(p);
      return props;
    }

    for (const doc of docs) {
      if (!doc) continue;
      for (const ff of descendants(doc, 'style:font-face')) {
        fonts.set(attr(ff, 'style:name'), (attr(ff, 'svg:font-family') || attr(ff, 'style:name') || '').replace(/^['"]|['"]$/g, ''));
      }
      for (const container of ['office:styles', 'office:automatic-styles']) {
        for (const c of descendants(doc, container)) {
          for (const el of elements(c)) {
            if (NUMBER_STYLE.test(el.nodeName)) numStyles.set(attr(el, 'style:name'), el);
            else if (el.nodeName === 'style:style') {
              // Stil adları aileye özgüdür ("Default" hem hücre hem çizim stili olabilir)
              raw.set(attr(el, 'style:family') + ':' + attr(el, 'style:name'), {
                family: attr(el, 'style:family'),
                parent: attr(el, 'style:parent-style-name'),
                dataStyle: attr(el, 'style:data-style-name'),
                master: attr(el, 'style:master-page-name'),
                props: readProps(el),
              });
            } else if (el.nodeName === 'style:default-style') {
              defaults.set(attr(el, 'style:family'), readProps(el));
            }
          }
        }
      }
    }

    const resolved = new Map();
    function resolve(name, family) {
      const key = family + ':' + name;
      if (resolved.has(key)) return resolved.get(key);
      const chain = [];
      let cur = name, guard = 0;
      while (cur && raw.has(family + ':' + cur) && guard++ < 20) {
        const st = raw.get(family + ':' + cur);
        chain.unshift(st);
        cur = st.parent;
      }
      const out = { cell: {}, para: {}, text: {}, col: {}, row: {}, table: {}, dataStyle: null, master: null };
      const base = defaults.get(family) || {};
      for (const k of Object.keys(base)) Object.assign(out[k] || (out[k] = {}), base[k]);
      // Hücre stilleri her zaman "Default" stilinden türer
      const def = raw.get('table-cell:Default');
      if (family === 'table-cell' && def && chain[0] !== def) chain.unshift(def);
      for (const s of chain) {
        const tp = s.props.text;
        if (tp && tp['fo:color'] && !('style:use-window-font-color' in tp)) delete out.text['style:use-window-font-color'];
        for (const k of Object.keys(s.props)) Object.assign(out[k] || (out[k] = {}), s.props[k]);
        if (s.dataStyle) out.dataStyle = s.dataStyle;
        if (s.master) out.master = s.master;
      }
      resolved.set(key, out);
      return out;
    }

    const formatCache = new Map();
    function numFmt(dataStyle) {
      if (!dataStyle) return null;
      if (!formatCache.has(dataStyle)) {
        const el = numStyles.get(dataStyle);
        formatCache.set(dataStyle, el ? formatOf(el, numStyles, 0) : null);
      }
      return formatCache.get(dataStyle);
    }

    const excelCache = new Map();
    function excelStyle(name) {
      const key = name || 'Default';
      if (excelCache.has(key)) return excelCache.get(key);
      const r = resolve(key, 'table-cell');
      const st = {};

      const t = r.text;
      const font = {};
      const fname = t['style:font-name'] ? fonts.get(t['style:font-name']) || t['style:font-name'] : t['fo:font-family'];
      if (fname) font.name = fname.replace(/^['"]|['"]$/g, '');
      const size = toPt(t['fo:font-size']);
      if (size) font.size = Math.round(size * 10) / 10;
      const weight = t['fo:font-weight'];
      if (weight === 'bold' || parseInt(weight, 10) >= 600) font.bold = true;
      if (t['fo:font-style'] === 'italic' || t['fo:font-style'] === 'oblique') font.italic = true;
      if (t['style:text-underline-style'] && t['style:text-underline-style'] !== 'none') {
        font.underline = t['style:text-underline-type'] === 'double' ? 'double' : true;
      }
      if (t['style:text-line-through-style'] && t['style:text-line-through-style'] !== 'none') font.strike = true;
      if (t['style:use-window-font-color'] !== 'true') {
        const c = argb(t['fo:color']);
        if (c) font.color = { argb: c };
      }
      if (/^super/.test(t['style:text-position'] || '')) font.vertAlign = 'superscript';
      if (/^sub/.test(t['style:text-position'] || '')) font.vertAlign = 'subscript';
      if (Object.keys(font).length) st.font = font;

      const c = r.cell;
      const bg = argb(c['fo:background-color']);
      if (bg) st.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg }, bgColor: { indexed: 64 } };

      const border = {};
      for (const side of Object.keys(BORDER_SIDES)) {
        const b = parseBorder(c[BORDER_SIDES[side]] || c['fo:border']);
        if (b) border[side] = b;
      }
      const up = parseBorder(c['style:diagonal-bl-tr']);
      const down = parseBorder(c['style:diagonal-tl-br']);
      if (up || down) border.diagonal = Object.assign({ up: !!up, down: !!down }, up || down);
      if (Object.keys(border).length) st.border = border;

      const al = {};
      const ta = r.para['fo:text-align'];
      if (c['style:text-align-source'] !== 'value-type' && ta) {
        al.horizontal = { start: 'left', left: 'left', end: 'right', right: 'right', center: 'center', justify: 'justify' }[ta];
      }
      if (c['style:repeat-content'] === 'true') al.horizontal = 'fill';
      const va = c['style:vertical-align'];
      if (va && va !== 'automatic') al.vertical = { top: 'top', middle: 'middle', bottom: 'bottom' }[va];
      if (c['fo:wrap-option'] === 'wrap') al.wrapText = true;
      if (c['style:shrink-to-fit'] === 'true') al.shrinkToFit = true;
      const rot = parseFloat(c['style:rotation-angle'] || '0');
      if (c['style:direction'] === 'ttb') al.textRotation = 'vertical';
      else if (rot) {
        const a = ((rot % 360) + 360) % 360;
        al.textRotation = a <= 90 ? a : a >= 270 ? a - 360 : a < 180 ? 90 : -90;
      }
      const indent = toPt(r.para['fo:margin-left']);
      if (indent && indent > 0 && al.horizontal !== 'center') al.indent = Math.round(indent / 10);
      if (Object.keys(al).length) st.alignment = al;

      const prot = c['style:cell-protect'];
      if (prot) st.protection = { locked: /protected/.test(prot), hidden: /hidden/.test(prot) };

      const nf = numFmt(r.dataStyle);
      if (nf && nf !== 'General') st.numFmt = nf;

      excelCache.set(key, st);
      return st;
    }

    return { resolve, excelStyle };
  }

  // ---------------------------------------------------------------- hücre içeriği

  function cellText(cell) {
    const paras = [];
    let link = null;
    function walk(node, acc) {
      for (let n = node.firstChild; n; n = n.nextSibling) {
        if (n.nodeType === 3) { acc.push(n.nodeValue); continue; }
        if (n.nodeType !== 1) continue;
        switch (n.nodeName) {
          case 'text:s': acc.push(' '.repeat(parseInt(attr(n, 'text:c') || '1', 10))); break;
          case 'text:tab': acc.push('\t'); break;
          case 'text:line-break': acc.push('\n'); break;
          case 'text:a': if (!link) link = attr(n, 'xlink:href'); walk(n, acc); break;
          case 'office:annotation': case 'text:note': break;
          default: walk(n, acc);
        }
      }
    }
    for (const p of elements(cell)) {
      if (p.nodeName !== 'text:p' && p.nodeName !== 'text:h') continue;
      const acc = [];
      walk(p, acc);
      paras.push(acc.join(''));
    }
    return { text: paras.length ? paras.join('\n') : null, link };
  }

  function annotationText(cell) {
    const a = child(cell, 'office:annotation');
    if (!a) return null;
    return elements(a).filter((n) => n.nodeName === 'text:p').map((p) => p.textContent).join('\n');
  }

  const EPOCH = Date.UTC(1899, 11, 30);

  // Tarihleri Excel seri sayısına çevirir (saat dilimi kaymalarını önlemek için Date nesnesi kullanmadan).
  function dateSerial(v) {
    const m = /^(-?\d{4,})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2})(?::(\d{2})(\.\d+)?)?)?/.exec(v);
    if (!m) return null;
    const days = (Date.UTC(+m[1], +m[2] - 1, +m[3]) - EPOCH) / 86400000;
    const secs = (+(m[4] || 0)) * 3600 + (+(m[5] || 0)) * 60 + (+(m[6] || 0)) + parseFloat(m[7] || '0');
    return days + secs / 86400;
  }

  function durationDays(v) {
    const m = /^(-)?P(?:(\d+)D)?(?:T(?:(\d+)H)?(?:(\d+)M)?(?:([\d.]+)S)?)?$/.exec(v);
    if (!m) return null;
    const d = (+(m[2] || 0)) + ((+(m[3] || 0)) * 3600 + (+(m[4] || 0)) * 60 + parseFloat(m[5] || '0')) / 86400;
    return m[1] ? -d : d;
  }

  // LibreOffice hata metinleri arayüz diline göre yerelleşir; Türkçe ve İngilizce karşılıklar.
  const ERRORS = {
    '#N/A': '#N/A', '#YOK': '#N/A', '#VALUE!': '#VALUE!', '#DEĞER!': '#VALUE!', '#REF!': '#REF!', '#BAŞV!': '#REF!',
    '#DIV/0!': '#DIV/0!', '#SAYI/0!': '#DIV/0!', '#NUM!': '#NUM!', '#SAYI!': '#NUM!', '#NAME?': '#NAME?', '#AD?': '#NAME?',
    '#NULL!': '#NULL!', '#BOŞ!': '#NULL!',
  };

  function cellValue(cell) {
    const type = attr(cell, 'office:value-type') || attr(cell, 'calcext:value-type');
    const { text, link } = cellText(cell);
    let value;
    switch (type) {
      case 'float': case 'percentage': case 'currency':
        value = parseFloat(attr(cell, 'office:value'));
        if (Number.isNaN(value)) value = null;
        break;
      case 'date': value = dateSerial(attr(cell, 'office:date-value') || ''); break;
      case 'time': value = durationDays(attr(cell, 'office:time-value') || ''); break;
      case 'boolean': {
        const b = attr(cell, 'office:boolean-value');
        value = b === 'true' || b === '1';
        break;
      }
      case 'string': {
        const sv = attr(cell, 'office:string-value');
        value = sv != null ? sv : (text || '');
        break;
      }
      default: value = text;
    }
    if (attr(cell, 'calcext:value-type') === 'error') value = { error: ERRORS[(text || '').trim()] || '#VALUE!' };
    return { value, link };
  }

  // ---------------------------------------------------------------- veri doğrulama

  function splitArgs(s) {
    const out = [];
    let cur = '', depth = 0, q = false;
    for (let i = 0; i < s.length; i++) {
      const ch = s[i];
      if (ch === '"') { q = !q; cur += ch; continue; }
      if (!q && (ch === '(' || ch === '[')) depth++;
      if (!q && (ch === ')' || ch === ']')) depth--;
      if (!q && depth === 0 && ch === ';') { out.push(cur); cur = ''; continue; }
      cur += ch;
    }
    if (cur.trim() !== '') out.push(cur);
    return out.map((x) => x.trim());
  }

  const OPERATORS = { '<=': 'lessThanOrEqual', '>=': 'greaterThanOrEqual', '!=': 'notEqual', '<>': 'notEqual', '=': 'equal', '<': 'lessThan', '>': 'greaterThan' };

  function parseValidation(el) {
    const cond = (attr(el, 'table:condition') || '').replace(/^[a-z]+:/i, '');
    const v = { allowBlank: attr(el, 'table:allow-empty-cell') !== 'false' };
    const help = child(el, 'table:help-message');
    if (help) {
      v.showInputMessage = attr(help, 'table:display') === 'true';
      if (attr(help, 'table:title')) v.promptTitle = attr(help, 'table:title');
      const msg = elements(help).filter((n) => n.nodeName === 'text:p').map((p) => p.textContent).join('\n');
      if (msg) v.prompt = msg;
    }
    const err = child(el, 'table:error-message');
    if (err) {
      v.showErrorMessage = attr(err, 'table:display') === 'true';
      v.errorStyle = { warning: 'warning', information: 'information' }[attr(err, 'table:message-type')] || 'stop';
      if (attr(err, 'table:title')) v.errorTitle = attr(err, 'table:title');
      const msg = elements(err).filter((n) => n.nodeName === 'text:p').map((p) => p.textContent).join('\n');
      if (msg) v.error = msg;
    }

    const fx = (s) => translateFormula('of:=' + s);
    let m;
    if ((m = /^cell-content-is-in-list\(([\s\S]*)\)$/.exec(cond))) {
      const inner = m[1].trim();
      const args = splitArgs(inner);
      v.type = 'list';
      if (args.length && args.every((a) => /^".*"$/.test(a))) {
        v.formulae = ['"' + args.map((a) => a.slice(1, -1).replace(/""/g, '"')).join(',') + '"'];
      } else {
        v.formulae = [fx(inner)];
      }
      return v;
    }

    const typeMap = {
      'cell-content-is-whole-number()': 'whole',
      'cell-content-is-decimal-number()': 'decimal',
      'cell-content-is-date()': 'date',
      'cell-content-is-time()': 'time',
    };
    let rest = cond;
    for (const k of Object.keys(typeMap)) {
      if (rest.startsWith(k)) { v.type = typeMap[k]; rest = rest.slice(k.length).replace(/^\s*and\s*/, ''); break; }
    }
    if (/^cell-content-text-length/.test(rest)) {
      v.type = 'textLength';
      rest = rest.replace(/^cell-content-text-length-is-/, 'cell-content-is-').replace(/^cell-content-text-length\(\)/, 'cell-content()');
    }
    if ((m = /^is-true-formula\(([\s\S]*)\)$/.exec(rest))) {
      v.type = 'custom';
      v.formulae = [fx(m[1])];
      return v;
    }
    if ((m = /^cell-content-is-(not-)?between\(([\s\S]*)\)$/.exec(rest))) {
      v.operator = m[1] ? 'notBetween' : 'between';
      v.formulae = splitArgs(m[2]).map(fx);
    } else if ((m = /^cell-content\(\)\s*(<=|>=|!=|<>|=|<|>)\s*([\s\S]+)$/.exec(rest))) {
      v.operator = OPERATORS[m[1]];
      v.formulae = [fx(m[2])];
    } else if (v.type && !rest) {
      v.operator = 'greaterThanOrEqual';
      v.formulae = [v.type === 'date' || v.type === 'time' ? '0' : '-1E+307'];
    } else {
      return null;
    }
    if (!v.type) v.type = 'decimal';
    return v;
  }

  // ---------------------------------------------------------------- sayfa yapısı ve görünüm

  function pageSetup(p) {
    p = p || {};
    const out = {};
    const w = toPt(p['fo:page-width']), h = toPt(p['fo:page-height']);
    if (w && h) {
      const a = Math.min(w, h) * 25.4 / 72, b = Math.max(w, h) * 25.4 / 72;
      const near = (x, y) => Math.abs(a - x) < 2 && Math.abs(b - y) < 2;
      if (near(210, 297)) out.paperSize = 9;
      else if (near(297, 420)) out.paperSize = 8;
      else if (near(148, 210)) out.paperSize = 11;
      else if (near(215.9, 279.4)) out.paperSize = 1;
      else if (near(215.9, 355.6)) out.paperSize = 5;
    }
    if (p['style:print-orientation']) out.orientation = p['style:print-orientation'] === 'landscape' ? 'landscape' : 'portrait';
    const margins = {};
    for (const [k, a] of [['top', 'fo:margin-top'], ['bottom', 'fo:margin-bottom'], ['left', 'fo:margin-left'], ['right', 'fo:margin-right']]) {
      const v = toPt(p[a]);
      if (v != null) margins[k] = Math.round(v / 72 * 10000) / 10000;
    }
    if (Object.keys(margins).length) out.margins = Object.assign({ left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 }, margins);
    if (p['style:scale-to-X'] || p['style:scale-to-Y']) {
      out.fitToPage = true;
      out.fitToWidth = parseInt(p['style:scale-to-X'] || '0', 10);
      out.fitToHeight = parseInt(p['style:scale-to-Y'] || '0', 10);
    } else if (p['style:scale-to-pages']) {
      out.fitToPage = true;
      out.fitToWidth = 1;
      out.fitToHeight = parseInt(p['style:scale-to-pages'], 10);
    } else if (p['style:scale-to']) {
      out.scale = Math.round(parseFloat(p['style:scale-to']));
    }
    if (p['style:print-page-order'] === 'ltr') out.pageOrder = 'overThenDown';
    const centering = p['style:table-centering'];
    if (centering === 'both' || centering === 'horizontal') out.horizontalCentered = true;
    if (centering === 'both' || centering === 'vertical') out.verticalCentered = true;
    if (p['style:print']) out.showGridLines = /grid/.test(p['style:print']);
    return out;
  }

  function readSettings(doc) {
    const out = { sheets: new Map(), activeTable: null, showGrid: true };
    if (!doc) return out;
    const item = (parent, name) => {
      for (const n of elements(parent)) if (n.nodeName === 'config:config-item' && attr(n, 'config:name') === name) return n.textContent;
      return null;
    };
    const views = descendants(doc, 'config:config-item-map-indexed').find((v) => attr(v, 'config:name') === 'Views');
    const entry = views && elements(views)[0];
    if (!entry) return out;
    out.activeTable = item(entry, 'ActiveTable');
    if (item(entry, 'ShowGrid') != null) out.showGrid = item(entry, 'ShowGrid') === 'true';
    const tablesMap = elements(entry).find((n) => n.nodeName === 'config:config-item-map-named' && attr(n, 'config:name') === 'Tables');
    if (tablesMap) {
      for (const t of elements(tablesMap)) {
        const get = (k) => item(t, k);
        out.sheets.set(attr(t, 'config:name'), {
          hSplitMode: get('HorizontalSplitMode'), vSplitMode: get('VerticalSplitMode'),
          hSplit: get('HorizontalSplitPosition'), vSplit: get('VerticalSplitPosition'),
          posRight: get('PositionRight'), posBottom: get('PositionBottom'),
          zoom: get('ZoomValue'), showGrid: get('ShowGrid'),
          cursorX: get('CursorPositionX'), cursorY: get('CursorPositionY'),
        });
      }
    }
    return out;
  }

  // ---------------------------------------------------------------- dönüştürme

  // Excel sütun genişliği birimi: varsayılan yazı tipindeki "0" karakterinin genişliği.
  // LibreOffice'in xlsx çıktısıyla ölçülerek bulundu (varsayılan 10pt yazı tipi için).
  const PT_PER_CHAR = 5.55;
  const MAX_ROWS = 1048576, MAX_COLS = 16384;
  const DEFAULT_ROW_PT = 12.8, DEFAULT_COL_PT = 64.25;

  function base64(bytes) {
    let s = '';
    for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
    return btoa(s);
  }

  function imageExt(path, mime) {
    const m = /\.(png|jpe?g|gif)$/i.exec(path || '');
    if (m) return m[1].toLowerCase().replace('jpg', 'jpeg');
    if (/png/.test(mime || '')) return 'png';
    if (/jpe?g/.test(mime || '')) return 'jpeg';
    if (/gif/.test(mime || '')) return 'gif';
    return null;
  }

  // Stil nesneleri önbellekten paylaşıldığı için değişiklikler kopya üzerinde yapılır (yine önbellekli).
  const derived = new WeakMap();
  function variant(style, key, make) {
    let m = derived.get(style);
    if (!m) derived.set(style, (m = new Map()));
    if (!m.has(key)) m.set(key, make());
    return m.get(key);
  }

  function withWrap(style) {
    if (style.alignment && style.alignment.wrapText) return style;
    return variant(style, 'wrap', () => Object.assign({}, style, { alignment: Object.assign({}, style.alignment, { wrapText: true }) }));
  }

  function withBorder(style, master, edges) {
    return variant(style, 'b' + edges + JSON.stringify(master), () => {
      const b = {};
      if (edges & 1 && master.top) b.top = master.top;
      if (edges & 2 && master.bottom) b.bottom = master.bottom;
      if (edges & 4 && master.left) b.left = master.left;
      if (edges & 8 && master.right) b.right = master.right;
      if (master.diagonal) b.diagonal = master.diagonal;
      const out = Object.assign({}, style);
      if (Object.keys(b).length) out.border = b; else delete out.border;
      return out;
    });
  }

  // "Sayfa.B4" / "'Sayfa adı'.$B$4" → { col, row } (1 tabanlı)
  function parseCellAddress(a) {
    const m = /\$?([A-Z]+)\$?(\d+)$/i.exec(a || '');
    return m ? { col: colIndex(m[1]), row: parseInt(m[2], 10) } : null;
  }

  const num = (v, d) => { const n = parseInt(v, 10); return Number.isFinite(n) ? n : d; };
  const isCell = (n) => n.nodeName === 'table:table-cell' || n.nodeName === 'table:covered-table-cell';

  // input: ODS paketinin dosyaları (Map yol → Uint8Array) veya .fods XML metni
  // Dönüş: { data: xlsx baytları, sheets: sayfa sayısı }
  async function convert(input, opts) {
    const options = opts || {};
    const parser = new (options.DOMParser || root.DOMParser)();
    const parse = (data) => (data ? parser.parseFromString(typeof data === 'string' ? data : new TextDecoder().decode(data), 'application/xml') : null);

    const files = input instanceof Map ? input : null;
    let contentDoc, stylesDoc, settingsDoc;
    if (files) {
      contentDoc = parse(files.get('content.xml'));
      stylesDoc = parse(files.get('styles.xml'));
      settingsDoc = parse(files.get('settings.xml'));
    } else {
      contentDoc = stylesDoc = settingsDoc = parse(input);
    }
    const spreadsheet = contentDoc && descendants(contentDoc, 'office:spreadsheet')[0];
    if (!spreadsheet) throw new Error('Dosyada tablo bulunamadı');

    const styles = buildStyles(stylesDoc === contentDoc ? [contentDoc] : [stylesDoc, contentDoc]);
    const settings = readSettings(settingsDoc);

    const masters = new Map(), layouts = new Map();
    for (const doc of new Set([stylesDoc, contentDoc])) {
      if (!doc) continue;
      for (const pl of descendants(doc, 'style:page-layout')) layouts.set(attr(pl, 'style:name'), attrMap(child(pl, 'style:page-layout-properties')));
      for (const mp of descendants(doc, 'style:master-page')) masters.set(attr(mp, 'style:name'), attr(mp, 'style:page-layout-name'));
    }

    const ExcelJS = options.ExcelJS || root.ExcelJS;
    const wb = new ExcelJS.Workbook();
    wb.calcProperties.fullCalcOnLoad = true; // Excel açılışta tüm formülleri yeniden hesaplar

    const tables = elements(spreadsheet).filter((n) => n.nodeName === 'table:table');
    const validations = new Map();
    for (const v of descendants(spreadsheet, 'table:content-validation')) {
      const parsed = parseValidation(v);
      if (parsed) validations.set(attr(v, 'table:name'), parsed);
    }

    const definedNames = [];
    const sheetIndex = new Map();
    const usedNames = new Set();

    for (const table of tables) {
      const name = attr(table, 'table:name');
      let wsName = name.replace(/[\\/?*[\]:]/g, '_').slice(0, 31);
      for (let i = 2; usedNames.has(wsName.toLowerCase()); i++) wsName = wsName.slice(0, 28) + '_' + i;
      usedNames.add(wsName.toLowerCase());
      sheetIndex.set(name, wb.worksheets.length);
      const ws = wb.addWorksheet(wsName);
      const tStyle = styles.resolve(attr(table, 'table:style-name'), 'table');
      if (tStyle.table['table:display'] === 'false') ws.state = 'hidden';
      const tab = argb(tStyle.table['table:tab-color'] || tStyle.table['tableooo:tab-color']);
      if (tab) ws.properties.tabColor = { argb: tab };

      // ---- sütun tanımları
      const cols = [];
      (function walkCols(node, level) {
        for (const n of elements(node)) {
          if (n.nodeName === 'table:table-column') {
            const rep = Math.min(num(attr(n, 'table:number-columns-repeated'), 1), MAX_COLS - cols.length);
            const cs = styles.resolve(attr(n, 'table:style-name'), 'table-column');
            const info = {
              widthPt: toPt(cs.col['style:column-width']),
              hidden: /^(collapse|filter)$/.test(attr(n, 'table:visibility') || ''),
              defaultStyle: attr(n, 'table:default-cell-style-name') || 'Default',
              outline: level,
              repeat: rep,
            };
            for (let i = 0; i < rep; i++) cols.push(info);
          } else if (/^table:table-(header-columns|column-group|columns)$/.test(n.nodeName)) {
            walkCols(n, level + (n.nodeName === 'table:table-column-group' ? 1 : 0));
          }
        }
      })(table, 0);
      const colDefault = (i) => (cols[i] ? cols[i].defaultStyle : 'Default');

      // ---- satır öğeleri
      const rowEls = [];
      (function walkRows(node, level) {
        for (const n of elements(node)) {
          if (n.nodeName === 'table:table-row') rowEls.push({ el: n, level });
          else if (/^table:table-(header-rows|row-group|rows)$/.test(n.nodeName)) {
            walkRows(n, level + (n.nodeName === 'table:table-row-group' ? 1 : 0));
          }
        }
      })(table, 0);

      // İçerik sınırları: dosyaların sonundaki milyonlarca boş tekrar satır/sütunu yazmamak için.
      let maxRow = 0, maxCol = 0;
      {
        let r = 0;
        for (const { el } of rowEls) {
          const rrep = num(attr(el, 'table:number-rows-repeated'), 1);
          const rowDefault = attr(el, 'table:default-cell-style-name');
          let c = 0, rowHas = false;
          for (const cell of elements(el)) {
            if (!isCell(cell)) continue;
            const crep = num(attr(cell, 'table:number-columns-repeated'), 1);
            const content = cell.firstChild || attr(cell, 'office:value-type') || attr(cell, 'table:formula') ||
              attr(cell, 'table:content-validation-name') || attr(cell, 'table:number-columns-spanned');
            const sname = attr(cell, 'table:style-name') || rowDefault;
            const styled = crep <= 256 && sname && sname !== colDefault(c);
            if (content || styled) {
              rowHas = true;
              const span = num(attr(cell, 'table:number-columns-spanned'), 1);
              maxCol = Math.max(maxCol, Math.min(c + Math.max(crep, span), MAX_COLS));
              maxRow = Math.max(maxRow, r + num(attr(cell, 'table:number-rows-spanned'), 1));
            }
            c += crep;
          }
          if (rowHas) maxRow = Math.max(maxRow, r + (rrep <= 1000 ? rrep : 1));
          r += rrep;
          if (r >= MAX_ROWS) break;
        }
      }

      // ---- sütunlar
      const lastCol = cols[cols.length - 1];
      const colWidth = (pt) => Math.round(pt / PT_PER_CHAR * 100) / 100;
      if (lastCol && lastCol.widthPt) ws.properties.defaultColWidth = colWidth(lastCol.widthPt);
      for (let i = 0; i < cols.length; i++) {
        const info = cols[i];
        if (i >= maxCol && info === lastCol && info.repeat > 64) break; // sondaki boş sütun yığını
        const col = ws.getColumn(i + 1);
        if (info.widthPt) col.width = colWidth(info.widthPt);
        if (info.hidden) col.hidden = true;
        if (info.outline) col.outlineLevel = Math.min(info.outline, 7);
        if (info.defaultStyle !== 'Default') col.style = styles.excelStyle(info.defaultStyle);
      }

      // ---- satırlar ve hücreler
      const merges = [];
      const frames = [];
      const rowHeights = [];
      let lastRowHeight = null;
      let r = 0;
      for (const { el, level } of rowEls) {
        if (r >= maxRow) break;
        const rrep = num(attr(el, 'table:number-rows-repeated'), 1);
        const rs = styles.resolve(attr(el, 'table:style-name'), 'table-row');
        const heightPt = toPt(rs.row['style:row-height']);
        const hidden = /^(collapse|filter)$/.test(attr(el, 'table:visibility') || '');
        const rowDefault = attr(el, 'table:default-cell-style-name');
        lastRowHeight = heightPt || lastRowHeight;
        const cells = elements(el).filter(isCell);
        const rowCount = Math.min(rrep, maxRow - r);

        for (let k = 0; k < rowCount; k++) {
          const rowNo = r + k + 1;
          rowHeights[rowNo] = hidden ? 0 : (heightPt || DEFAULT_ROW_PT);
          const row = ws.getRow(rowNo);
          if (heightPt) row.height = Math.round(heightPt * 100) / 100;
          if (hidden) row.hidden = true;
          if (level) row.outlineLevel = Math.min(level, 7);

          let c = 0;
          for (const cell of cells) {
            if (c >= maxCol) break;
            const crep = num(attr(cell, 'table:number-columns-repeated'), 1);
            const covered = cell.nodeName === 'table:covered-table-cell';
            const formula = attr(cell, 'table:formula');
            const vname = attr(cell, 'table:content-validation-name');
            const span = num(attr(cell, 'table:number-columns-spanned'), 1);
            const rspan = num(attr(cell, 'table:number-rows-spanned'), 1);
            const content = !covered && (formula || attr(cell, 'office:value-type') || child(cell, 'text:p')) ? cellValue(cell) : null;
            const note = annotationText(cell);
            if (k === 0) for (const f of elements(cell)) if (f.nodeName === 'draw:frame') frames.push({ frame: f, row: r, col: c });

            const count = Math.min(crep, maxCol - c);
            for (let j = 0; j < count; j++) {
              const colNo = c + j + 1;
              const styleName = attr(cell, 'table:style-name') || rowDefault || colDefault(c + j);
              const empty = !formula && (!content || content.value === null || content.value === '');
              if (empty && !vname && !note && styleName === colDefault(c + j) && span === 1 && rspan === 1) continue;
              const xc = row.getCell(colNo);
              xc.style = styles.excelStyle(styleName);
              if (content) {
                const v = content.value;
                if (formula) {
                  const fv = { formula: translateFormula(formula) };
                  if (v !== null && v !== undefined) fv.result = v;
                  const mc = num(attr(cell, 'table:number-matrix-columns-spanned'), 0);
                  const mr = num(attr(cell, 'table:number-matrix-rows-spanned'), 0);
                  if (mc > 1 || mr > 1) {
                    fv.shareType = 'array';
                    fv.ref = colName(colNo) + rowNo + ':' + colName(colNo + Math.max(mc, 1) - 1) + (rowNo + Math.max(mr, 1) - 1);
                  }
                  xc.value = fv;
                } else if (content.link && typeof v === 'string') {
                  xc.value = { text: v, hyperlink: content.link };
                } else if (v !== null && v !== '') {
                  xc.value = v;
                }
                // LibreOffice, satır sonu içeren metinlerde metni kaydırmayı açar
                if (typeof v === 'string' && v.includes('\n')) xc.style = withWrap(xc.style);
              }
              if (note) xc.note = note;
              if (vname && validations.has(vname)) xc.dataValidation = validations.get(vname);
              if (!covered && j === 0 && (span > 1 || rspan > 1)) merges.push([rowNo, colNo, rowNo + rspan - 1, colNo + span - 1]);
            }
            c += crep;
          }
        }
        r += rrep;
      }
      if (lastRowHeight) ws.properties.defaultRowHeight = Math.round(lastRowHeight * 100) / 100;

      // Birleştirme: ODS'de kenarlık ana hücrenin stilinde tüm alana uygulanır; Excel'de ise kenardaki
      // hücrelerden çizilir. LibreOffice gibi ana hücrenin kenarlığını yalnızca dış kenarlara dağıtıyoruz;
      // hücrelerin diğer stil özellikleri (yazı tipi vb.) kendi stillerinden gelir.
      for (const m of merges) {
        const masterBorder = ws.getCell(m[0], m[1]).style.border || {};
        const saved = [];
        for (let rr = m[0]; rr <= m[2]; rr++) for (let cc = m[1]; cc <= m[3]; cc++) {
          const cl = ws.getCell(rr, cc);
          const edges = (rr === m[0] ? 1 : 0) | (rr === m[2] ? 2 : 0) | (cc === m[1] ? 4 : 0) | (cc === m[3] ? 8 : 0);
          saved.push([cl, withBorder(cl.style, masterBorder, edges)]);
        }
        try { ws.mergeCellsWithoutStyle(m[0], m[1], m[2], m[3]); } catch (e) { continue; }
        for (const [cl, st] of saved) cl.style = st;
      }

      // ---- resimler
      // Sayfaya bağlı çerçeveler table:shapes içinde (köprülüyse draw:a, grupluysa draw:g altında) ve
      // konumları sayfanın sol üstüne göredir; hücreye bağlı olanlarda konum o hücreye göredir.
      (function collect(node) {
        for (const f of elements(node)) {
          if (f.nodeName === 'draw:frame') frames.push({ frame: f, row: 0, col: 0 });
          else if (f.nodeName === 'draw:a' || f.nodeName === 'draw:g') collect(f);
        }
      })(child(table, 'table:shapes') || table.ownerDocument.createElement('x'));
      if (files && frames.length) {
        const colPt = (i) => (cols[i] ? (cols[i].hidden ? 0 : cols[i].widthPt || DEFAULT_COL_PT) : DEFAULT_COL_PT);
        const rowPt = (i) => (rowHeights[i + 1] != null ? rowHeights[i + 1] : (lastRowHeight || DEFAULT_ROW_PT));
        const EMU = 12700; // 1pt
        // (c0, r0) hücresinden (xPt, yPt) uzaklıktaki noktanın hücre + ofset karşılığı
        const anchor = (c0, r0, xPt, yPt) => {
          let c = c0, x = Math.max(0, xPt);
          while (c < MAX_COLS - 1 && x >= colPt(c) && (colPt(c) > 0 || c < c0 + 1000)) { x -= colPt(c); c++; }
          let rr = r0, y = Math.max(0, yPt);
          while (rr < MAX_ROWS - 1 && y >= rowPt(rr) && rr < r0 + 1048576) { y -= rowPt(rr); rr++; }
          return { nativeCol: c, nativeColOff: Math.round(x * EMU), nativeRow: rr, nativeRowOff: Math.round(y * EMU) };
        };
        const imageIds = new Map();
        for (const { frame, row: r0, col: c0 } of frames) {
          let chosen = null, ext = null;
          for (const im of elements(frame)) {
            if (im.nodeName !== 'draw:image') continue;
            const href = attr(im, 'xlink:href');
            const e = imageExt(href, attr(im, 'draw:mime-type'));
            if (e && files.has(href)) { chosen = href; ext = e; break; }
          }
          const w = toPt(attr(frame, 'svg:width')), h = toPt(attr(frame, 'svg:height'));
          if (!chosen || !w || !h) continue;
          if (!imageIds.has(chosen)) imageIds.set(chosen, wb.addImage({ base64: base64(files.get(chosen)), extension: ext }));
          const x = toPt(attr(frame, 'svg:x')) || 0, y = toPt(attr(frame, 'svg:y')) || 0;
          const end = parseCellAddress(attr(frame, 'table:end-cell-address'));
          const br = end
            ? anchor(end.col - 1, end.row - 1, toPt(attr(frame, 'table:end-x')) || 0, toPt(attr(frame, 'table:end-y')) || 0)
            : anchor(c0, r0, x + w, y + h);
          ws.addImage(imageIds.get(chosen), { tl: anchor(c0, r0, x, y), br, editAs: end ? 'oneCell' : 'absolute' });
        }
      }

      // ---- sayfa yapısı, yazdırma alanı, koruma, görünüm
      Object.assign(ws.pageSetup, pageSetup(layouts.get(masters.get(tStyle.master || 'Default'))));

      const pr = attr(table, 'table:print-ranges');
      if (pr) {
        const ranges = splitRanges(pr).map((a) => translateRef(a)).filter((x) => !/#REF!/.test(x))
          .map((ref) => quoteSheet(wsName) + '!' + absolute(ref.replace(/^.*!/, '')));
        if (ranges.length) definedNames.push({ name: '_xlnm.Print_Area', ranges, localSheetId: sheetIndex.get(name) });
      }

      if (attr(table, 'table:protected') === 'true') {
        // ODS parola özeti Excel'e taşınamaz; LibreOffice gibi parolasız koruma yazıyoruz.
        const p = child(table, 'loext:table-protection');
        const opt = { objects: true, scenarios: true };
        if (p && attr(p, 'loext:select-protected-cells') === 'false') opt.selectLockedCells = false;
        if (p && attr(p, 'loext:select-unprotected-cells') === 'false') opt.selectUnlockedCells = false;
        await ws.protect('', opt);
      }

      const st = settings.sheets.get(name) || {};
      const view = {};
      const grid = st.showGrid != null ? st.showGrid === 'true' : settings.showGrid;
      if (grid === false) view.showGridLines = false;
      if (st.zoom && st.zoom !== '100') view.zoomScale = num(st.zoom, 100);
      const xs = st.hSplitMode === '2' ? num(st.hSplit, 0) : 0;
      const ys = st.vSplitMode === '2' ? num(st.vSplit, 0) : 0;
      if (xs || ys) {
        Object.assign(view, { state: 'frozen', xSplit: xs, ySplit: ys });
        view.topLeftCell = colName(Math.max(num(st.posRight, 0), xs) + 1) + (Math.max(num(st.posBottom, 0), ys) + 1);
      }
      if (st.cursorX != null && st.cursorY != null) view.activeCell = colName(num(st.cursorX, 0) + 1) + (num(st.cursorY, 0) + 1);
      if (settings.activeTable === name) view.tabSelected = true;
      ws.views = [view];
    }

    // ---- otomatik filtreler
    for (const dr of descendants(spreadsheet, 'table:database-range')) {
      if (attr(dr, 'table:display-filter-buttons') !== 'true') continue;
      const ref = translateRef(attr(dr, 'table:target-range-address') || '');
      const m = /^(?:'((?:[^']|'')+)'|([^!]+))!(.+)$/.exec(ref);
      if (!m) continue;
      const i = sheetIndex.get((m[1] || m[2]).replace(/''/g, "'"));
      if (i == null) continue;
      const ws = wb.worksheets[i];
      ws.autoFilter = m[3].replace(/\$/g, '');
      definedNames.push({ name: '_xlnm._FilterDatabase', ranges: [quoteSheet(ws.name) + '!' + absolute(m[3])], localSheetId: i });
    }

    // ---- tanımlı adlar
    for (const nr of descendants(spreadsheet, 'table:named-range')) {
      const ref = translateRef(attr(nr, 'table:cell-range-address') || '');
      if (!/!/.test(ref) || /#REF!/.test(ref)) continue;
      const entry = { name: attr(nr, 'table:name'), ranges: [ref.replace(/!(.*)$/, (s, cells) => '!' + absolute(cells))] };
      const owner = nr.parentNode && nr.parentNode.parentNode;
      if (owner && owner.nodeName === 'table:table') entry.localSheetId = sheetIndex.get(attr(owner, 'table:name'));
      definedNames.push(entry);
    }
    for (const ne of descendants(spreadsheet, 'table:named-expression')) {
      const expr = translateFormula(attr(ne, 'table:expression') || '');
      if (expr && !/#REF!/.test(expr)) definedNames.push({ name: attr(ne, 'table:name'), ranges: [expr] });
    }
    // ExcelJS'in ad modeli sayfaya özel adları ve çoklu yazdırma alanlarını desteklemediği için modeli doğrudan veriyoruz.
    wb._definedNames = { model: definedNames };

    const active = tables.findIndex((t) => attr(t, 'table:name') === settings.activeTable);
    if (active >= 0) wb.views = [{ activeTab: active, firstSheet: 0, visibility: 'visible' }];

    const xlsx = new Uint8Array(await wb.xlsx.writeBuffer());
    return { data: await setDefaultFont(xlsx, styles.excelStyle('Default').font), sheets: wb.worksheets.length };
  }

  // ---------------------------------------------------------------- xlsx son düzeltme

  // ExcelJS, styles.xml'deki varsayılan yazı tipini (fontId 0) Calibri 11 olarak sabit yazar. Stil
  // atanmamış hücreler ve Excel'de sonradan yazılanlar bu yazı tipini kullandığından onu belgenin
  // varsayılanıyla değiştiriyoruz. Diğer zip girdileri sıkıştırılmış hâlleriyle aynen kopyalanır.
  const CRC_TABLE = (() => {
    const t = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      t[n] = c >>> 0;
    }
    return t;
  })();

  function crc32(bytes) {
    let c = 0xffffffff;
    for (let i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xff] ^ (c >>> 8);
    return (c ^ 0xffffffff) >>> 0;
  }

  async function inflate(bytes) {
    const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
    return new Uint8Array(await new Response(stream).arrayBuffer());
  }

  async function deflate(bytes) {
    const stream = new Blob([bytes]).stream().pipeThrough(new CompressionStream('deflate-raw'));
    return new Uint8Array(await new Response(stream).arrayBuffer());
  }

  const xmlEscape = (v) => String(v).replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;');

  async function setDefaultFont(zip, font) {
    if (!font || !font.name) return zip;
    const dv = new DataView(zip.buffer, zip.byteOffset, zip.byteLength);
    let eocd = -1;
    for (let i = zip.length - 22; i >= 0; i--) if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    if (eocd < 0) return zip;
    const count = dv.getUint16(eocd + 10, true);
    let p = dv.getUint32(eocd + 16, true);
    const enc = new TextEncoder(), dec = new TextDecoder();
    const entries = [];
    for (let n = 0; n < count; n++) {
      const e = {
        method: dv.getUint16(p + 10, true), time: dv.getUint16(p + 12, true), date: dv.getUint16(p + 14, true),
        crc: dv.getUint32(p + 16, true), csize: dv.getUint32(p + 20, true), usize: dv.getUint32(p + 24, true),
        nameBytes: zip.slice(p + 46, p + 46 + dv.getUint16(p + 28, true)),
      };
      const local = dv.getUint32(p + 42, true);
      const start = local + 30 + dv.getUint16(local + 26, true) + dv.getUint16(local + 28, true);
      e.data = zip.subarray(start, start + e.csize);
      e.name = dec.decode(e.nameBytes);
      p += 46 + dv.getUint16(p + 28, true) + dv.getUint16(p + 30, true) + dv.getUint16(p + 32, true);
      entries.push(e);
    }
    const st = entries.find((e) => e.name === 'xl/styles.xml');
    if (!st) return zip;
    const xml = dec.decode(st.method === 8 ? await inflate(st.data) : st.data);
    const fontXml = '<font>' + (font.bold ? '<b/>' : '') + (font.italic ? '<i/>' : '') +
      '<sz val="' + (font.size || 10) + '"/>' + (font.color ? '<color rgb="' + font.color.argb + '"/>' : '') +
      '<name val="' + xmlEscape(font.name) + '"/><family val="2"/></font>';
    const patched = xml.replace(/(<fonts\b[^>]*>)\s*<font>[\s\S]*?<\/font>/, '$1' + fontXml);
    if (patched === xml) return zip;
    const plain = enc.encode(patched);
    st.crc = crc32(plain);
    st.usize = plain.length;
    st.data = await deflate(plain);
    st.method = 8;
    st.csize = st.data.length;

    let size = 22;
    for (const e of entries) size += 30 + 46 + 2 * e.nameBytes.length + e.data.length;
    const out = new Uint8Array(size);
    const o = new DataView(out.buffer);
    let at = 0;
    const offsets = [];
    for (const e of entries) {
      offsets.push(at);
      o.setUint32(at, 0x04034b50, true); o.setUint16(at + 4, 20, true); o.setUint16(at + 6, 0x0800, true);
      o.setUint16(at + 8, e.method, true); o.setUint16(at + 10, e.time, true); o.setUint16(at + 12, e.date, true);
      o.setUint32(at + 14, e.crc, true); o.setUint32(at + 18, e.csize, true); o.setUint32(at + 22, e.usize, true);
      o.setUint16(at + 26, e.nameBytes.length, true); o.setUint16(at + 28, 0, true);
      out.set(e.nameBytes, at + 30);
      out.set(e.data, at + 30 + e.nameBytes.length);
      at += 30 + e.nameBytes.length + e.data.length;
    }
    const cdStart = at;
    entries.forEach((e, i) => {
      o.setUint32(at, 0x02014b50, true); o.setUint16(at + 4, 20, true); o.setUint16(at + 6, 20, true);
      o.setUint16(at + 8, 0x0800, true); o.setUint16(at + 10, e.method, true); o.setUint16(at + 12, e.time, true);
      o.setUint16(at + 14, e.date, true); o.setUint32(at + 16, e.crc, true); o.setUint32(at + 20, e.csize, true);
      o.setUint32(at + 24, e.usize, true); o.setUint16(at + 28, e.nameBytes.length, true);
      o.setUint32(at + 42, offsets[i], true);
      out.set(e.nameBytes, at + 46);
      at += 46 + e.nameBytes.length;
    });
    o.setUint32(at, 0x06054b50, true);
    o.setUint16(at + 8, entries.length, true); o.setUint16(at + 10, entries.length, true);
    o.setUint32(at + 12, at - cdStart, true); o.setUint32(at + 16, cdStart, true);
    return out.subarray(0, at + 22);
  }

  root.OdsToXlsx = { convert, translateFormula };
})(typeof globalThis !== 'undefined' ? globalThis : this);
