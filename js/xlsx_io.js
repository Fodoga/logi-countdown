/* xlsx_io.js — 零依赖 浏览器端 Excel 读取/写入 + CSV 解析/导出
 * 不依赖 SheetJS / 网络。xlsx 本质是 ZIP(deflate) 包 XML，
 * 用浏览器原生 DecompressionStream('deflate-raw') 解压、手写 ZIP(存储) 写回。
 * 导出时对“订单号/快递单号/单号”等标识列强制文本格式(s="1", numFmtId=49)，
 * 防止 Excel 把长/前导零数字重新转成数值。
 */
(function (global) {
  "use strict";

  /* ---------- CRC32 ---------- */
  const CRC_TABLE = (function () {
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

  /* ---------- ZIP 解析（读 xlsx） ---------- */
  function parseZip(ab) {
    const dv = new DataView(ab);
    const bytes = new Uint8Array(ab);
    let eocd = -1;
    const minPos = Math.max(0, ab.byteLength - 22 - 65535);
    for (let i = ab.byteLength - 22; i >= minPos; i--) {
      if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    }
    if (eocd < 0) throw new Error("不是有效的 ZIP / xlsx 文件");
    const cdOffset = dv.getUint32(eocd + 16, true);
    const numEntries = dv.getUint16(eocd + 10, true);
    const map = new Map();
    let p = cdOffset;
    for (let n = 0; n < numEntries; n++) {
      if (dv.getUint32(p, true) !== 0x02014b50) throw new Error("ZIP 目录损坏");
      const method = dv.getUint16(p + 10, true);
      const csize = dv.getUint32(p + 20, true);
      const nl = dv.getUint16(p + 28, true);
      const el = dv.getUint16(p + 30, true);
      const cl = dv.getUint16(p + 32, true);
      const name = new TextDecoder().decode(bytes.subarray(p + 46, p + 46 + nl));
      const localOffset = dv.getUint32(p + 42, true);
      map.set(name, { offset: localOffset, comp: method, csize });
      p += 46 + nl + el + cl;
    }
    return map;
  }

  function readEntry(ab, entry) {
    const dv = new DataView(ab);
    const bytes = new Uint8Array(ab);
    const p = entry.offset;
    const nl = dv.getUint16(p + 26, true);
    const el = dv.getUint16(p + 28, true);
    const dataStart = p + 30 + nl + el;
    return bytes.subarray(dataStart, dataStart + entry.csize);
  }

  /* ---------- 纯 JS inflate (RFC1951 raw deflate)，零依赖、跨环境一致 ---------- */
  function inflateRawJS(input) {
    const data = input instanceof Uint8Array ? input : new Uint8Array(input);
    const bytes = data;
    let pos = 0, bitPos = 0;
    function readBit() {
      const bit = (bytes[pos] >> bitPos) & 1;
      bitPos++;
      if (bitPos === 8) { pos++; bitPos = 0; }
      return bit;
    }
    function readBits(n) {
      let val = 0;
      for (let i = 0; i < n; i++) val |= readBit() << i;
      return val;
    }
    function readBytes(n) {
      const out = bytes.subarray(pos, pos + n);
      pos += n; bitPos = 0;
      return out;
    }
    // 构建 Huffman 树
    function makeHuffman(lengths) {
      const maxLen = Math.max.apply(null, lengths) || 1;
      const count = new Array(maxLen + 1).fill(0);
      const symbols = [];
      lengths.forEach((l, i) => { if (l) count[l]++; });
      const first = new Array(maxLen + 1).fill(0);
      let code = 0;
      for (let l = 1; l <= maxLen; l++) { first[l] = code; code += count[l]; code <<= 1; }
      const next = first.slice();
      const tree = new Array(lengths.length);
      lengths.forEach((l, i) => {
        if (l) { tree[i] = { code: first[l] + (next[l] - first[l]), len: l }; next[l]++; }
      });
      return { maxLen, first, count, tree };
    }
    function decodeSymbol(h) {
      let code = 0, len = 0;
      while (len <= h.maxLen) {
        code = (code << 1) | readBit(); len++;
        const target = code;
        for (let i = 0; i < h.tree.length; i++) {
          const t = h.tree[i];
          if (t && t.len === len && t.code === target) return i;
        }
      }
      throw new Error("inflate: 无效的 Huffman 编码");
    }
    const out = [];
    const LENGTH_BASE = [3,4,5,6,7,8,9,10,11,13,15,17,19,23,27,31,35,43,51,59,67,83,99,115,131,163,195,227,258];
    const LENGTH_EXTRA = [0,0,0,0,0,0,0,0,1,1,1,1,2,2,2,2,3,3,3,3,4,4,4,4,5,5,5,5,0];
    const DIST_BASE = [1,2,3,4,5,7,9,13,17,25,33,49,65,97,129,193,257,385,513,769,1025,1537,2049,3073,4097,6145,8193,12289,16385,24577];
    const DIST_EXTRA = [0,0,0,0,1,1,2,2,3,3,4,4,5,5,6,6,7,7,8,8,9,9,10,10,11,11,12,12,13,13];
    while (true) {
      const btype = readBits(3); // BFINAL(1)+BTYPE(2)
      const bfinal = btype & 1;
      const btype2 = btype >> 1;
      if (btype2 === 0) {
        // 存储块
        bitPos = 0; pos++; // 对齐到字节
        const len = bytes[pos] | (bytes[pos + 1] << 8);
        pos += 4;
        const chunk = readBytes(len);
        for (let i = 0; i < chunk.length; i++) out.push(chunk[i]);
      } else if (btype2 === 1) {
        // 固定 Huffman
        const litLengths = new Array(288);
        for (let i = 0; i < 288; i++) litLengths[i] = i < 144 ? 8 : i < 256 ? 9 : i < 280 ? 7 : 8;
        const distLengths = new Array(30).fill(5);
        const hl = makeHuffman(litLengths), hd = makeHuffman(distLengths);
        while (true) {
          const sym = decodeSymbol(hl);
          if (sym === 256) break;
          if (sym < 256) out.push(sym);
          else {
            const idx = sym - 257;
            const len = LENGTH_BASE[idx] + readBits(LENGTH_EXTRA[idx]);
            const dsym = decodeSymbol(hd);
            const dist = DIST_BASE[dsym] + readBits(DIST_EXTRA[dsym]);
            for (let i = 0; i < len; i++) out.push(out[out.length - dist]);
          }
        }
      } else if (btype2 === 2) {
        // 动态 Huffman
        const hlit = readBits(5) + 257;
        const hdist = readBits(5) + 1;
        const hclen = readBits(4) + 4;
        const clOrder = [16,17,18,0,8,7,9,6,10,5,11,4,12,3,13,2,14,1,15];
        const clLengths = new Array(19).fill(0);
        for (let i = 0; i < hclen; i++) clLengths[clOrder[i]] = readBits(3);
        const hc = makeHuffman(clLengths);
        const lengths = [];
        while (lengths.length < hlit + hdist) {
          const sym = decodeSymbol(hc);
          if (sym < 16) lengths.push(sym);
          else if (sym === 16) { const rep = readBits(2) + 3; const prev = lengths[lengths.length - 1]; for (let i = 0; i < rep; i++) lengths.push(prev); }
          else if (sym === 17) { const rep = readBits(3) + 3; for (let i = 0; i < rep; i++) lengths.push(0); }
          else if (sym === 18) { const rep = readBits(7) + 11; for (let i = 0; i < rep; i++) lengths.push(0); }
        }
        const litLengths = lengths.slice(0, hlit);
        const distLengths = lengths.slice(hlit);
        const hl = makeHuffman(litLengths), hd = makeHuffman(distLengths);
        while (true) {
          const sym = decodeSymbol(hl);
          if (sym === 256) break;
          if (sym < 256) out.push(sym);
          else {
            const idx = sym - 257;
            const len = LENGTH_BASE[idx] + readBits(LENGTH_EXTRA[idx]);
            const dsym = decodeSymbol(hd);
            const dist = DIST_BASE[dsym] + readBits(DIST_EXTRA[dsym]);
            for (let i = 0; i < len; i++) out.push(out[out.length - dist]);
          }
        }
      } else {
        throw new Error("inflate: 无效的块类型");
      }
      if (bfinal) break;
    }
    return new Uint8Array(out);
  }

  async function inflateRaw(raw) {
    const buf = raw instanceof Uint8Array ? raw : new Uint8Array(await raw.arrayBuffer());
    return inflateRawJS(buf); // 纯 JS 实现，不 fallback（避免 Node 端 DecompressionStream 段错误）
  }

  async function getEntryText(ab, map, name) {
    const e = map.get(name);
    if (!e) return null;
    let data = readEntry(ab, e);
    if (e.comp === 8) data = await inflateRaw(data);
    return new TextDecoder("utf-8").decode(data);
  }

  function textContent(el) {
    let s = "";
    for (const c of el.childNodes) {
      if (c.nodeType === 3) s += c.nodeValue;
      else if (c.nodeType === 1) s += textContent(c);
    }
    return s;
  }

  // 解码 XML 实体（&amp; &lt; &gt; &quot; &apos; 及 &#13; 等数值实体）
  function decodeXml(s) {
    if (!s) return "";
    return s
      .replace(/&lt;/g, "<").replace(/&gt;/g, ">")
      .replace(/&quot;/g, '"').replace(/&apos;/g, "'")
      .replace(/&#(\d+);/g, (_, d) => String.fromCharCode(parseInt(d, 10)))
      .replace(/&amp;/g, "&");
  }

  function refToCol(ref) {
    const m = ref.match(/^([A-Za-z]+)/);
    if (!m) return 0;
    let n = 0;
    for (const ch of m[1].toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n - 1;
  }

  function parseShared(xmlText) {
    const doc = new DOMParser().parseFromString(xmlText, "application/xml");
    const sis = doc.getElementsByTagName("si");
    const arr = [];
    for (let i = 0; i < sis.length; i++) arr.push(decodeXml(textContent(sis[i])));
    return arr;
  }

  function parseSheet(xmlText, shared) {
    const doc = new DOMParser().parseFromString(xmlText, "application/xml");
    if (doc.getElementsByTagName("parsererror").length) throw new Error("工作表 XML 解析失败");
    const rowEls = doc.getElementsByTagName("row");
    const grid = [];
    for (let i = 0; i < rowEls.length; i++) {
      const cellEls = rowEls[i].getElementsByTagName("c");
      const obj = {};
      for (let j = 0; j < cellEls.length; j++) {
        const cell = cellEls[j];
        const ref = cell.getAttribute("r");
        if (!ref) continue; // 无 ref 的单元格跳过（标准 Excel 均带 r）
        const col = refToCol(ref);
        const t = cell.getAttribute("t");
        let val = "";
        const isEls = cell.getElementsByTagName("is");
        if (isEls.length) {
          val = decodeXml(textContent(isEls[0]));
        } else {
          const vEls = cell.getElementsByTagName("v");
          if (vEls.length) {
            const v = vEls[0].textContent;
            if (t === "s") val = shared[parseInt(v, 10)] || "";
            else val = v;
          }
        }
        obj[col] = val;
      }
      grid.push(obj);
    }
    if (!grid.length) return [];
    const header = grid[0];
    const result = [];
    for (let i = 1; i < grid.length; i++) {
      const o = {};
      for (const k in header) {
        o[header[k]] = grid[i][k] != null ? grid[i][k] : "";
      }
      result.push(o); // 保留所有数据行（含空行），由引擎层决定是否过滤
    }
    return result;
  }

  async function readXlsxFile(file) {
    const ab = await file.arrayBuffer();
    const map = parseZip(ab);
    const sharedXml = await getEntryText(ab, map, "xl/sharedStrings.xml");
    const shared = sharedXml ? parseShared(sharedXml) : [];

    let sheetTarget = null;
    const wbXml = await getEntryText(ab, map, "xl/workbook.xml");
    if (wbXml) {
      const doc = new DOMParser().parseFromString(wbXml, "application/xml");
      const sheetEls = doc.getElementsByTagName("sheet");
      if (sheetEls.length) {
        const rid = sheetEls[0].getAttribute("r:id") || sheetEls[0].getAttribute("r:id");
        const relsXml = await getEntryText(ab, map, "xl/_rels/workbook.xml.rels");
        if (relsXml) {
          const rdoc = new DOMParser().parseFromString(relsXml, "application/xml");
          const rels = rdoc.getElementsByTagName("Relationship");
          for (const r of rels) {
            if (r.getAttribute("Id") === rid) { sheetTarget = r.getAttribute("Target"); break; }
          }
        }
        if (!sheetTarget && rid) sheetTarget = "worksheets/sheet1.xml";
      }
    }
    if (!sheetTarget) {
      for (const name of map.keys()) {
        if (/xl\/worksheets\/sheet\d+\.xml$/.test(name)) { sheetTarget = name; break; }
      }
    }
    if (!sheetTarget) throw new Error("未找到工作表");
    if (sheetTarget.startsWith("/")) sheetTarget = sheetTarget.slice(1);
    if (!/^xl\//.test(sheetTarget)) sheetTarget = "xl/" + sheetTarget;
    const sheetXml = await getEntryText(ab, map, sheetTarget);
    if (!sheetXml) throw new Error("工作表内容为空");
    return parseSheet(sheetXml, shared);
  }

  /* ---------- CSV 解析 ---------- */
  function parseCSV(text) {
    if (text.charCodeAt(0) === 0xfeff) text = text.slice(1);
    // 探测分隔符
    const firstLine = (text.split(/\r?\n/)[0] || "");
    const sep = firstLine.indexOf(",") !== -1 ? "," : (firstLine.indexOf(";") !== -1 ? ";" : ",");
    const rows = [];
    let field = "", row = [], inQ = false;
    for (let i = 0; i < text.length; i++) {
      const c = text[i];
      if (inQ) {
        if (c === '"') { if (text[i + 1] === '"') { field += '"'; i++; } else inQ = false; }
        else field += c;
      } else {
        if (c === '"') inQ = true;
        else if (c === sep) { row.push(field); field = ""; }
        else if (c === "\r") { /* ignore */ }
        else if (c === "\n") { row.push(field); rows.push(row); row = []; field = ""; }
        else field += c;
      }
    }
    if (field !== "" || row.length) { row.push(field); rows.push(row); }
    if (!rows.length) return [];
    const header = rows[0].map((h, idx) => (h || "").trim() || ("列" + (idx + 1)));
    const out = [];
    for (let i = 1; i < rows.length; i++) {
      const o = {}; let empty = true;
      for (let j = 0; j < header.length; j++) {
        const v = rows[i][j] != null ? rows[i][j] : "";
        if (v !== "") empty = false;
        o[header[j]] = v;
      }
      if (!empty) out.push(o);
    }
    return out;
  }

  async function readFileAuto(file) {
    const name = (file && file.name ? file.name : "").toLowerCase();
    if (name.endsWith(".csv")) return parseCSV(await file.text());
    if (name.endsWith(".xlsx") || name.endsWith(".xlsm")) return await readXlsxFile(file);
    // 兜底：先试 xlsx，再试 csv
    try { return await readXlsxFile(file); } catch (e) { /* fall through */ }
    return parseCSV(await file.text());
  }

  /* ---------- 写 xlsx ---------- */
  const STYLES_XML =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>' +
    '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>' +
    '<borders count="1"><border/></borders>' +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    '<cellXfs count="2">' +
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
    '<xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +
    '</cellXfs></styleSheet>';

  const WORKBOOK_XML =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>';

  const WB_RELS =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
    '</Relationships>';

  const ROOT_RELS =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
    '</Relationships>';

  const CONTENT_TYPES =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
    '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
    '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
    '</Types>';

  function colLetter(i) {
    let s = ""; i++;
    while (i > 0) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); }
    return s;
  }
  function xmlEscape(s) {
    s = String(s == null ? "" : s);
    return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  }

  function buildSheet(rows, columns, idCols) {
    let out = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>';
    out += '<row r="1">';
    columns.forEach((c, ci) => {
      const ref = colLetter(ci) + "1";
      out += '<c r="' + ref + '" s="1" t="inlineStr"><is><t xml:space="preserve">' + xmlEscape(c) + '</t></is></c>';
    });
    out += '</row>';
    rows.forEach((r, ri) => {
      const rn = ri + 2;
      out += '<row r="' + rn + '">';
      columns.forEach((c, ci) => {
        const ref = colLetter(ci) + rn;
        let v = r[c]; if (v == null) v = ""; v = String(v);
        if (idCols.has(c)) {
          out += '<c r="' + ref + '" s="1" t="inlineStr"><is><t xml:space="preserve">' + xmlEscape(v) + '</t></is></c>';
        } else if (/^-?\d+(\.\d+)?$/.test(v)) {
          out += '<c r="' + ref + '"><v>' + v + '</v></c>';
        } else {
          out += '<c r="' + ref + '" t="inlineStr"><is><t xml:space="preserve">' + xmlEscape(v) + '</t></is></c>';
        }
      });
      out += '</row>';
    });
    out += '</sheetData></worksheet>';
    return out;
  }

  function localHeader(nameBytes, crc, size) {
    const buf = new Uint8Array(30 + nameBytes.length);
    const dv = new DataView(buf.buffer);
    dv.setUint32(0, 0x04034b50, true);
    dv.setUint16(4, 20, true); dv.setUint16(6, 0, true); dv.setUint16(8, 0, true);
    dv.setUint16(10, 0, true); dv.setUint16(12, 0, true);
    dv.setUint32(14, crc, true); dv.setUint32(18, size, true); dv.setUint32(22, size, true);
    dv.setUint16(26, nameBytes.length, true); dv.setUint16(28, 0, true);
    buf.set(nameBytes, 30);
    return buf;
  }
  function centralHeader(nameBytes, crc, size, localOffset) {
    const buf = new Uint8Array(46 + nameBytes.length);
    const dv = new DataView(buf.buffer);
    dv.setUint32(0, 0x02014b50, true);
    dv.setUint16(4, 20, true); dv.setUint16(6, 20, true); dv.setUint16(8, 0, true);
    dv.setUint16(10, 0, true); dv.setUint16(12, 0, true); dv.setUint16(14, 0, true);
    dv.setUint32(16, crc, true); dv.setUint32(20, size, true); dv.setUint32(24, size, true);
    dv.setUint16(28, nameBytes.length, true); dv.setUint16(30, 0, true); dv.setUint16(32, 0, true);
    dv.setUint32(34, 0, true); dv.setUint16(38, 0, true); dv.setUint32(40, 0, true);
    dv.setUint32(42, localOffset, true);
    buf.set(nameBytes, 46);
    return buf;
  }
  function eocd(cdSize, cdOffset, count) {
    const buf = new Uint8Array(22);
    const dv = new DataView(buf.buffer);
    dv.setUint32(0, 0x06054b50, true);
    dv.setUint16(4, 0, true); dv.setUint16(6, 0, true);
    dv.setUint16(8, count, true); dv.setUint16(10, count, true);
    dv.setUint32(12, cdSize, true); dv.setUint32(16, cdOffset, true); dv.setUint16(20, 0, true);
    return buf;
  }

  function makeZip(files) {
    const enc = new TextEncoder();
    const parts = [];
    const central = [];
    let offset = 0;
    for (const [name, text] of files) {
      const nameBytes = enc.encode(name);
      const data = enc.encode(text);
      const crc = crc32(data);
      const lh = localHeader(nameBytes, crc, data.length);
      parts.push(lh); parts.push(data);
      const localOffset = offset;
      offset += lh.length + data.length;
      central.push(centralHeader(nameBytes, crc, data.length, localOffset));
    }
    const cdStart = offset;
    let cdSize = 0;
    central.forEach((ch) => { parts.push(ch); cdSize += ch.length; });
    offset += cdSize;
    parts.push(eocd(cdSize, cdStart, files.length));
    let total = 0; parts.forEach((p) => (total += p.length));
    const out = new Uint8Array(total);
    let pos = 0; parts.forEach((p) => { out.set(p, pos); pos += p.length; });
    return new Blob([out], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  }

  function toXLSX(rows, columns) {
    const idCols = new Set(columns.filter((c) => /订单号|快递单号|单号|订单编号|运单|物流单/.test(c)));
    const sheet = buildSheet(rows, columns, idCols);
    const files = [
      ["[Content_Types].xml", CONTENT_TYPES],
      ["_rels/.rels", ROOT_RELS],
      ["xl/workbook.xml", WORKBOOK_XML],
      ["xl/_rels/workbook.xml.rels", WB_RELS],
      ["xl/styles.xml", STYLES_XML],
      ["xl/worksheets/sheet1.xml", sheet],
    ];
    return makeZip(files);
  }

  function toCSV(rows, columns) {
    const esc = (v) => {
      v = v == null ? "" : String(v);
      if (/[",\n\r]/.test(v)) return '"' + v.replace(/"/g, '""') + '"';
      return v;
    };
    const lines = [columns.map(esc).join(",")];
    rows.forEach((r) => lines.push(columns.map((c) => esc(r[c])).join(",")));
    return "﻿" + lines.join("\r\n");
  }

  const XlsxIO = { readXlsxFile, parseCSV, readFileAuto, toXLSX, toCSV, parseSheet };
  if (typeof module !== "undefined" && module.exports) module.exports = XlsxIO;
  else global.QXlsxIO = XlsxIO;
})(typeof window !== "undefined" ? window : globalThis);
