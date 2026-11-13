// ===== Cœur : calculs + génération du fichier Excel d'origine complété =====
const pad2 = n => String(n).padStart(2, "0");
function toMin(t) { if (!t) return null; const m = /^(\d{1,2}):(\d{2})$/.exec(t); return m ? (+m[1]) * 60 + (+m[2]) : null; }
function dayHours(s, e) { const a = toMin(s), b = toMin(e); if (a == null || b == null) return null; return (((b - a) % 1440) + 1440) % 1440 / 60; }
function fmtH(h) { if (h == null || h === "") return "—"; const t = Math.round(h * 60); return pad2(Math.floor(t / 60)) + ":" + pad2(t % 60); }
function fmtDec(h) { return (h == null || h === "") ? "" : h.toFixed(2).replace(".", ","); }
// Prénoms : ceux de l'Excel, remplaçables par config.js (prenoms: [...]) — l'en-tête de l'Excel suit
META.people.forEach(p => { if (p.orig == null) p.orig = p.name; });
{ const pr = (typeof window !== "undefined" && window.POINTEUSE_CONFIG && window.POINTEUSE_CONFIG.prenoms) || [];
  META.people.forEach((p, i) => { if (pr[i] && String(pr[i]).trim()) p.name = String(pr[i]).trim(); }); }
// Clé stable par colonne Excel (B, F, J) : changer un prénom ne perd aucun pointage
const recKey = (iso, pi) => iso + "_" + META.people[pi].start;
const personOfKey = k => META.people.findIndex(p => k.endsWith("_" + p.start));

// Évalue les formules du tableau (mêmes formules que dans Excel) → {"D6": 7.5, ...}
function computeValues(recs) {
  const v = {};
  for (const row of META.rows) {
    if (row.kind !== "day") continue;
    META.people.forEach((p, pi) => {
      const r = recs[recKey(row.iso, pi)] || {};
      const h = dayHours(r.start, r.end);
      v[p.total + row.r] = h == null ? "" : h;
    });
  }
  const refVal = ref => { const x = v[ref]; return typeof x === "number" ? x : 0; };
  for (const row of META.rows) {
    if (row.kind !== "total") continue;
    META.people.forEach((p, pi) => {
      // formule de la colonne D, transposée à la colonne de la personne
      const f = row.f.replace(/\bD(\d+)/g, (_, n) => p.total + n);
      let m;
      if ((m = /^=SUM\((.*)\)$/.exec(f))) v[p.total + row.r] = m[1].split(",").reduce((s, ref) => s + refVal(ref.trim()), 0);
      else v[p.total + row.r] = f.slice(1).split("+").reduce((s, ref) => s + refVal(ref.trim()), 0);
    });
  }
  return v;
}

const xmlEsc = s => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

function setCell(xml, ref, build) {
  const re = new RegExp('<c r="' + ref + '"([^>]*?)(?:/>|>([\\s\\S]*?)</c>)');
  const m = re.exec(xml);
  if (!m) throw new Error("Cellule introuvable : " + ref);
  const attrs = m[1].replace(/\s*\bt="[^"]*"/, "");
  return xml.slice(0, m.index) + build(attrs, m[2] || "") + xml.slice(m.index + m[0].length);
}

function buildSheet(recs, addStr) {
  let xml = META.parts["xl/worksheets/sheet1.xml"];
  META.people.forEach(p => {
    if (p.name !== p.orig) xml = setCell(xml, p.start + "4", a => '<c r="' + p.start + '4"' + a + ' t="s"><v>' + addStr(p.name) + "</v></c>");
  });
  for (const row of META.rows) {
    if (row.kind !== "day") continue;
    META.people.forEach((p, pi) => {
      const r = recs[recKey(row.iso, pi)] || {};
      for (const [col, val] of [[p.start, r.start], [p.end, r.end]]) {
        const mn = toMin(val);
        xml = setCell(xml, col + row.r, a => mn == null ? '<c r="' + col + row.r + '"' + a + "/>" : '<c r="' + col + row.r + '"' + a + "><v>" + (mn / 1440) + "</v></c>");
      }
      const sg = (r.sign || "").trim();
      xml = setCell(xml, p.sign + row.r, a => sg ? '<c r="' + p.sign + row.r + '"' + a + ' t="s"><v>' + addStr(sg) + "</v></c>" : '<c r="' + p.sign + row.r + '"' + a + "/>");
    });
  }
  // valeurs en cache des formules (lisibles même dans un aperçu qui ne recalcule pas)
  const vals = computeValues(recs);
  for (const ref in vals) {
    xml = setCell(xml, ref, (a, inner) => {
      const fm = /<f\b[^>]*?(?:\/>|>[\s\S]*?<\/f>)/.exec(inner); if (!fm) throw new Error("pas de formule " + ref + " " + inner); const f = fm[0];
      const x = vals[ref];
      return x === "" ? '<c r="' + ref + '"' + a + ' t="str">' + f + "<v></v></c>" : '<c r="' + ref + '"' + a + ">" + f + "<v>" + (x / 24) + "</v></c>"; // valeur en jours Excel (format [h]:mm)
    });
  }
  return xml;
}

// --- ZIP (méthode « stored », sans dépendance) ---
const CRC_T = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
function crc32(b) { let c = 0xFFFFFFFF; for (let i = 0; i < b.length; i++) c = CRC_T[(c ^ b[i]) & 255] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; }
function makeZip(files) {
  const enc = new TextEncoder(), chunks = [], central = []; let off = 0;
  const d = new Date(), dt = ((d.getHours() << 11) | (d.getMinutes() << 5) | (d.getSeconds() >> 1)) & 0xFFFF, dd = (((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate()) & 0xFFFF;
  for (const [name, text] of files) {
    const nb = enc.encode(name), data = enc.encode(text), crc = crc32(data);
    const h = new DataView(new ArrayBuffer(30));
    h.setUint32(0, 0x04034b50, true); h.setUint16(4, 20, true); h.setUint16(6, 0x0800, true); h.setUint16(8, 0, true);
    h.setUint16(10, dt, true); h.setUint16(12, dd, true); h.setUint32(14, crc, true); h.setUint32(18, data.length, true); h.setUint32(22, data.length, true);
    h.setUint16(26, nb.length, true); h.setUint16(28, 0, true);
    chunks.push(new Uint8Array(h.buffer), nb, data);
    const c = new DataView(new ArrayBuffer(46));
    c.setUint32(0, 0x02014b50, true); c.setUint16(4, 20, true); c.setUint16(6, 20, true); c.setUint16(8, 0x0800, true); c.setUint16(10, 0, true);
    c.setUint16(12, dt, true); c.setUint16(14, dd, true); c.setUint32(16, crc, true); c.setUint32(20, data.length, true); c.setUint32(24, data.length, true);
    c.setUint16(28, nb.length, true); c.setUint32(42, off, true);
    central.push(new Uint8Array(c.buffer), nb);
    off += 30 + nb.length + data.length;
  }
  const cs = central.reduce((s, a) => s + a.length, 0);
  const e = new DataView(new ArrayBuffer(22));
  e.setUint32(0, 0x06054b50, true); e.setUint16(8, files.length, true); e.setUint16(10, files.length, true); e.setUint32(12, cs, true); e.setUint32(16, off, true);
  const all = [...chunks, ...central, new Uint8Array(e.buffer)], out = new Uint8Array(off + cs + 22); let p = 0;
  for (const a of all) { out.set(a, p); p += a.length; }
  return out;
}

function buildXlsx(recs) {
  const parts = Object.assign({}, META.parts);
  // signatures ajoutées à la table des chaînes partagées (format le plus compatible : Excel, Numbers, aperçus)
  let sst = parts["xl/sharedStrings.xml"];
  const uniq = +(/uniqueCount="(\d+)"/.exec(sst) || [0, (sst.match(/<si>/g) || []).length])[1];
  const added = [], idx = {};
  const addStr = t => { if (!(t in idx)) { idx[t] = uniq + added.length; added.push(t); } return idx[t]; };
  parts["xl/worksheets/sheet1.xml"] = buildSheet(recs, addStr);
  let used = META.people.filter(p => p.name !== p.orig).length; for (const row of META.rows) if (row.kind === "day") META.people.forEach((p, pi) => { if ((recs[recKey(row.iso, pi)] || {}).sign) used++; });
  sst = sst.replace(/<\/sst>\s*$/, added.map(t => '<si><t xml:space="preserve">' + xmlEsc(t) + "</t></si>").join("") + "</sst>")
    .replace(/uniqueCount="\d+"/, 'uniqueCount="' + (uniq + added.length) + '"')
    .replace(/(<sst[^>]*?\s)count="(\d+)"/, (m, a, n) => a + 'count="' + (+n + used) + '"');
  parts["xl/sharedStrings.xml"] = sst;
  parts["xl/workbook.xml"] = parts["xl/workbook.xml"].replace(/<calcPr([^/]*)\/>/, (m, a) => /fullCalcOnLoad/.test(a) ? m : "<calcPr" + a + ' fullCalcOnLoad="1"/>');
  return makeZip(META.order.map(n => [n, parts[n]]));
}

// ===== Export du détail quotidien d'une personne (classeur simple, une feuille) =====
function buildDetailXlsx(recs, pi) {
  const p = META.people[pi];
  const esc = s => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  const S = (ref, txt, st) => '<c r="' + ref + '" t="inlineStr"' + (st ? ' s="' + st + '"' : "") + "><is><t>" + esc(txt) + "</t></is></c>";
  const N = (ref, v, st) => '<c r="' + ref + '" s="' + st + '"><v>' + v + "</v></c>";
  const rows = []; let n = 0;
  const row = cells => { n++; rows.push('<row r="' + n + '">' + cells.map(f => f(n)).join("") + "</row>"); };
  row([r => S("A" + r, "Heures de " + p.name + " — " + (META.title || "Planning"), 3)]);
  row([]);
  row([r => S("A" + r, "Date", 3), r => S("B" + r, "Arrivée", 3), r => S("C" + r, "Départ", 3), r => S("D" + r, "Durée", 3), r => S("E" + r, "Signature", 3)]);
  let week = 0, total = 0;
  for (const rw of META.rows) {
    if (rw.kind === "day") {
      const x = recs[recKey(rw.iso, pi)] || {}, h = dayHours(x.start, x.end), a = toMin(x.start), b = toMin(x.end);
      if (h != null) { week += h; total += h; }
      row([r => S("A" + r, rw.label),
           r => a == null ? "" : N("B" + r, a / 1440, 1),
           r => b == null ? "" : N("C" + r, b / 1440, 1),
           r => h == null ? "" : N("D" + r, h / 24, 2),
           r => x.sign ? S("E" + r, x.sign) : ""]);
    } else if (/SEMAINE/.test(rw.label)) {
      row([r => S("A" + r, rw.label.replace("TOTAL ", "Total "), 3), r => N("D" + r, week / 24, 4)]);
      week = 0;
    }
  }
  row([]);
  row([r => S("A" + r, "Total général", 3), r => N("D" + r, total / 24, 4)]);
  const ns = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
  const files = [
    ["[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'],
    ["_rels/.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'],
    ["xl/workbook.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook ' + ns + ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Mes heures" sheetId="1" r:id="rId1"/></sheets></workbook>'],
    ["xl/_rels/workbook.xml.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'],
    ["xl/styles.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet ' + ns + '><numFmts count="2"><numFmt numFmtId="164" formatCode="hh:mm"/><numFmt numFmtId="165" formatCode="[hh]:mm"/></numFmts><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="5"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="165" fontId="1" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'],
    ["xl/worksheets/sheet1.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet ' + ns + '><cols><col min="1" max="1" width="30" customWidth="1"/><col min="2" max="4" width="11" customWidth="1"/><col min="5" max="5" width="26" customWidth="1"/></cols><sheetData>' + rows.join("") + "</sheetData></worksheet>"]
  ];
  return makeZip(files);
}
