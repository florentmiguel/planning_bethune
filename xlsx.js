// ===== Planning : structure de la période, calculs, génération du fichier Excel =====
const pad2 = n => String(n).padStart(2, "0");
const MOIS = ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre"];
const JOURS = ["dimanche", "lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi"];
function toMin(t) { if (!t) return null; const m = /^(\d{1,2}):(\d{2})$/.exec(t); return m ? (+m[1]) * 60 + (+m[2]) : null; }
function dayHours(s, e) { const a = toMin(s), b = toMin(e); if (a == null || b == null) return null; return (((b - a) % 1440) + 1440) % 1440 / 60; }
function fmtH(h) { if (h == null || h === "") return "—"; const t = Math.round(h * 60); return pad2(Math.floor(t / 60)) + ":" + pad2(t % 60); }
const isoOf = d => d.getFullYear() + "-" + pad2(d.getMonth() + 1) + "-" + pad2(d.getDate());
const dateOf = iso => { const [y, m, d] = iso.split("-").map(Number); return new Date(y, m - 1, d); };
const recKey = (iso, uid) => iso + "_" + uid;
function isoWeek(d) {
  const t = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  const day = t.getUTCDay() || 7; t.setUTCDate(t.getUTCDate() + 4 - day);
  const y0 = new Date(Date.UTC(t.getUTCFullYear(), 0, 1));
  return Math.ceil(((t - y0) / 86400000 + 1) / 7);
}
const dayLabel = d => JOURS[d.getDay()] + " " + d.getDate() + " " + MOIS[d.getMonth()] + " " + d.getFullYear();
const longDate = iso => { const d = dateOf(iso); return d.getDate() + " " + MOIS[d.getMonth()] + " " + d.getFullYear(); };

// Lignes du planning : jours, total de semaine après chaque dimanche, total de mois après le dernier jour du mois
function buildRows(debut, fin) {
  const rows = [], end = dateOf(fin); let week = [], month = [];
  for (let d = dateOf(debut); d <= end; d = new Date(d.getFullYear(), d.getMonth(), d.getDate() + 1)) {
    const iso = isoOf(d), next = new Date(d.getFullYear(), d.getMonth(), d.getDate() + 1), last = next > end;
    rows.push({ kind: "day", iso, label: dayLabel(d) }); week.push(iso); month.push(iso);
    if (d.getDay() === 0 || last) { rows.push({ kind: "week", label: "TOTAL SEMAINE " + isoWeek(d), days: week }); week = []; }
    if (next.getMonth() !== d.getMonth() || last) { rows.push({ kind: "month", label: "TOTAL " + MOIS[d.getMonth()].toUpperCase() + " " + d.getFullYear(), days: month }); month = []; }
  }
  rows.push({ kind: "total", label: "TOTAL GÉNÉRAL" });
  return rows;
}
// Total en heures d'une personne sur une liste de jours
function sumHours(recs, uid, days) { let s = 0; for (const iso of days) { const r = recs[recKey(iso, uid)]; const h = r && dayHours(r.start, r.end); if (h != null) s += h; } return s; }

// --- XML / ZIP ---
const xmlEsc = s => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
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
const colName = n => { let s = ""; n++; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }; // 0 -> A

// Styles (mêmes couleurs que le planning d'origine)
const STYLES = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
  '<numFmts count="2"><numFmt numFmtId="164" formatCode="hh:mm"/><numFmt numFmtId="165" formatCode="[hh]:mm"/></numFmts>' +
  '<fonts count="5"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="15"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font><font><i/><sz val="10"/><color rgb="FF595959"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>' +
  '<fills count="9"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
  ["17324D", "334E68", "DCEBFF", "E3F4D9", "FFF0CE", "00B0F0", "FFC000"].map(c => '<fill><patternFill patternType="solid"><fgColor rgb="FF' + c + '"/><bgColor indexed="64"/></patternFill></fill>').join("") + "</fills>" +
  '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color rgb="FFBFBFBF"/></left><right style="thin"><color rgb="FFBFBFBF"/></right><top style="thin"><color rgb="FFBFBFBF"/></top><bottom style="thin"><color rgb="FFBFBFBF"/></bottom><diagonal/></border></borders>' +
  '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="22">' +
  [ // [numFmt, font, fill, border, align]
    [0, 0, 0, 0, ""], [0, 1, 2, 0, "left"], [0, 2, 0, 0, "left"], [0, 3, 3, 1, "center"], [0, 3, 3, 1, "center"], [0, 0, 6, 1, ""],
    [164, 0, 4, 1, "center"], [165, 0, 4, 1, "center"], [0, 0, 4, 1, ""], [164, 0, 5, 1, "center"], [165, 0, 5, 1, "center"], [0, 0, 5, 1, ""],
    [0, 4, 7, 1, ""], [165, 4, 7, 1, "center"], [0, 4, 7, 1, ""], [0, 4, 8, 1, ""], [165, 4, 8, 1, "center"], [0, 4, 8, 1, ""],
    [0, 3, 2, 1, ""], [165, 3, 2, 1, "center"], [0, 3, 2, 1, ""], [0, 3, 3, 1, ""]
  ].map(([n, f, l, b, a]) => '<xf numFmtId="' + n + '" fontId="' + f + '" fillId="' + l + '" borderId="' + b + '" xfId="0"' + (n ? ' applyNumberFormat="1"' : "") + ' applyFont="1" applyFill="1" applyBorder="1"' + (a ? ' applyAlignment="1"><alignment horizontal="' + a + '" vertical="center"/></xf>' : "/>")).join("") +
  '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';
const ST = { title: 1, note: 2, name: 3, hdr: 4, date: 5, time: [6, 9], dur: [7, 10], sign: [8, 11], weekL: 12, weekT: 13, weekF: 14, monthL: 15, monthT: 16, monthF: 17, totL: 18, totT: 19, totF: 20, corner: 21 };

// Classeur complet : une ligne par jour, un bloc Début / Fin / Durée / Signature par personne
function buildPlanningXlsx(recs, periode, persons) {
  const rows = buildRows(periode.debut, periode.fin), nc = 1 + 4 * persons.length, last = colName(nc - 1);
  const out = [], merges = ["A1:" + last + "1", "A2:" + last + "2"];
  const S = (ref, txt, st) => '<c r="' + ref + '" s="' + st + '" t="inlineStr"><is><t xml:space="preserve">' + xmlEsc(txt) + "</t></is></c>";
  const E = (ref, st) => '<c r="' + ref + '" s="' + st + '"/>';
  const N = (ref, v, st) => '<c r="' + ref + '" s="' + st + '"><v>' + v + "</v></c>";
  const F = (ref, f, v, st) => v === "" || v == null ? '<c r="' + ref + '" s="' + st + '" t="str"><f>' + xmlEsc(f) + "</f><v></v></c>" : '<c r="' + ref + '" s="' + st + '"><f>' + xmlEsc(f) + "</f><v>" + v + "</v></c>";
  const R = (r, cells, ht) => out.push('<row r="' + r + '"' + (ht ? ' ht="' + ht + '" customHeight="1"' : "") + ">" + cells.join("") + "</row>");
  R(1, [S("A1", "PLANNING " + periode.libelle.toUpperCase() + " — DU " + longDate(periode.debut).toUpperCase() + " AU " + longDate(periode.fin).toUpperCase(), ST.title)], 30);
  R(2, [S("A2", "Début et fin au format 08:00. Durées et totaux en heures:minutes (ex. 07:30).", ST.note)]);
  const r3 = [E("A3", ST.corner)], r4 = [S("A4", "Date", ST.hdr)];
  persons.forEach((p, k) => {
    const c0 = 1 + 4 * k; r3.push(S(colName(c0) + "3", p.nom, ST.name)); for (let j = 1; j < 4; j++) r3.push(E(colName(c0 + j) + "3", ST.name));
    merges.push(colName(c0) + "3:" + colName(c0 + 3) + "3");
    ["Début", "Fin", "Durée", "Signature"].forEach((h, j) => r4.push(S(colName(c0 + j) + "4", h, ST.hdr)));
  });
  R(3, r3, 22); R(4, r4);
  // numéro de ligne Excel de chaque jour, pour les formules de total
  let r = 5; const rowOf = {}; for (const x of rows) { if (x.kind === "day") rowOf[x.iso] = r; r++; }
  const segs = days => { const ns = days.map(d => rowOf[d]).sort((a, b) => a - b), out = []; let s = ns[0], p = ns[0]; for (const n of ns.slice(1)) { if (n === p + 1) { p = n; continue; } out.push([s, p]); s = p = n; } if (ns.length) out.push([s, p]); return out; };
  const monthRows = []; r = 5;
  for (const x of rows) {
    const cells = [];
    if (x.kind === "day") {
      cells.push(S("A" + r, x.label, ST.date));
      persons.forEach((p, k) => {
        const c0 = 1 + 4 * k, alt = k % 2, rec = recs[recKey(x.iso, p.id)] || {}, a = toMin(rec.start), b = toMin(rec.end), h = dayHours(rec.start, rec.end);
        const cs = colName(c0) + r, ce = colName(c0 + 1) + r;
        cells.push(a == null ? E(cs, ST.time[alt]) : N(cs, a / 1440, ST.time[alt]));
        cells.push(b == null ? E(ce, ST.time[alt]) : N(ce, b / 1440, ST.time[alt]));
        cells.push(F(colName(c0 + 2) + r, 'IF(OR(' + cs + '="",' + ce + '=""),"",MOD(' + ce + "-" + cs + ",1))", h == null ? "" : h / 24, ST.dur[alt]));
        cells.push(rec.sign ? S(colName(c0 + 3) + r, rec.sign, ST.sign[alt]) : E(colName(c0 + 3) + r, ST.sign[alt]));
      });
    } else {
      const [L, T, Fi] = x.kind === "week" ? [ST.weekL, ST.weekT, ST.weekF] : x.kind === "month" ? [ST.monthL, ST.monthT, ST.monthF] : [ST.totL, ST.totT, ST.totF];
      cells.push(S("A" + r, x.label, L));
      if (x.kind === "month") monthRows.push(r);
      persons.forEach((p, k) => {
        const c0 = 1 + 4 * k, ct = colName(c0 + 2);
        let f, v;
        if (x.kind === "total") { f = "SUM(" + monthRows.map(m => ct + m).join(",") + ")"; v = sumHours(recs, p.id, rows.filter(y => y.kind === "day").map(y => y.iso)); }
        else { f = "SUM(" + segs(x.days).map(([s, e]) => s === e ? ct + s : ct + s + ":" + ct + e).join(",") + ")"; v = sumHours(recs, p.id, x.days); }
        cells.push(E(colName(c0) + r, Fi), E(colName(c0 + 1) + r, Fi), F(ct + r, f, v / 24, T), E(colName(c0 + 3) + r, Fi));
      });
    }
    R(r, cells, x.kind === "day" ? 0 : 20); r++;
  }
  const cols = '<cols><col min="1" max="1" width="27" customWidth="1"/>' + persons.map((p, k) => { const c = 2 + 4 * k; return '<col min="' + c + '" max="' + (c + 2) + '" width="9" customWidth="1"/><col min="' + (c + 3) + '" max="' + (c + 3) + '" width="34" customWidth="1"/>'; }).join("") + "</cols>";
  const ns = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"', nsr = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
  const sheet = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet ' + ns + " " + nsr + '><sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>' +
    '<sheetViews><sheetView workbookViewId="0"><pane xSplit="1" ySplit="4" topLeftCell="B5" activePane="bottomRight" state="frozen"/></sheetView></sheetViews>' +
    '<sheetFormatPr defaultRowHeight="15"/>' + cols + "<sheetData>" + out.join("") + "</sheetData>" +
    '<mergeCells count="' + merges.length + '">' + merges.map(m => '<mergeCell ref="' + m + '"/>').join("") + "</mergeCells>" +
    '<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/><pageSetup paperSize="9" orientation="landscape" fitToWidth="1" fitToHeight="0"/></worksheet>';
  return makeZip([
    ["[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'],
    ["_rels/.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'],
    ["xl/workbook.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook ' + ns + " " + nsr + '><sheets><sheet name="Planning" sheetId="1" r:id="rId1"/></sheets><definedNames><definedName name="_xlnm.Print_Titles" localSheetId="0">Planning!$1:$4</definedName></definedNames><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>'],
    ["xl/_rels/workbook.xml.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'],
    ["xl/styles.xml", STYLES],
    ["xl/worksheets/sheet1.xml", sheet]
  ]);
}
