/* Listy płac → Excel: rdzeń (parsowanie PDF i dopasowanie nazwisk).
 * Działa w przeglądarce (window.ListyPlac) i w Node (module.exports).
 * Wymaga pdf.js (pdfjsLib) przekazanego do parsePdf().
 */
(function (root) {
  "use strict";

  var NUM_RE = /-?\d{1,3}(?: \d{3})*,\d{2}/g;
  var NUM_FULL = /^-?\d{1,3}(?: \d{3})*,\d{2}$/;

  function toNum(s) { return parseFloat(s.replace(/\s/g, "").replace(",", ".")); }
  function numsIn(t) { var m = t.match(NUM_RE); return m ? m.map(toNum) : []; }
  function sum(a) { return a.reduce(function (x, y) { return x + y; }, 0); }
  function r2(x) { return Math.round(x * 100) / 100; }

  function norm(s) {
    return s.replace(/ł/g, "l").replace(/Ł/g, "L").normalize("NFKD")
      .replace(/[̀-ͯ]/g, "").replace(/\s+/g, " ").trim().toLowerCase();
  }
  function titlePl(s) {
    return s.trim().split(/\s+/).map(function (w) {
      return w.split("-").map(function (p) {
        return p.charAt(0).toLocaleUpperCase("pl") + p.slice(1).toLocaleLowerCase("pl");
      }).join("-");
    }).join(" ");
  }

  // Ratcliff/Obershelp – jak difflib.SequenceMatcher.ratio()
  function ratio(a, b) {
    function m(a, b) {
      if (!a.length || !b.length) return 0;
      var best = 0, ai = 0, bi = 0;
      for (var i = 0; i < a.length; i++) for (var j = 0; j < b.length; j++) {
        var k = 0; while (i + k < a.length && j + k < b.length && a[i + k] === b[j + k]) k++;
        if (k > best) { best = k; ai = i; bi = j; }
      }
      if (!best) return 0;
      return best + m(a.slice(0, ai), b.slice(0, bi)) + m(a.slice(ai + best), b.slice(bi + best));
    }
    var t = a.length + b.length;
    return t ? 2 * m(a, b) / t : 1;
  }

  // ---------------------------------------------------------------- tekst z PDF
  async function loadItems(pdfjsLib, data) {
    var doc = await pdfjsLib.getDocument({ data: data, isEvalSupported: false }).promise;
    var items = [], first = "";
    for (var p = 1; p <= doc.numPages; p++) {
      var pg = await doc.getPage(p);
      var vp = pg.getViewport({ scale: 1 });
      var tc = await pg.getTextContent();
      tc.items.forEach(function (it) {
        if (!it.str || !it.str.trim()) return;
        var x0 = it.transform[4];
        items.push({ x0: x0, x1: x0 + it.width, y: (p - 1) * 10000 + vp.height - it.transform[5], s: it.str });
      });
      if (p === 1) first = tc.items.map(function (i) { return i.str; }).join(" ");
    }
    items.sort(function (a, b) { return a.y - b.y || a.x0 - b.x0; });
    return { items: items, first: first };
  }

  function lineText(items, xmin, xmax, y, tol) {
    tol = tol == null ? 2.5 : tol;
    var sel = items.filter(function (i) { return i.x0 >= xmin - 0.5 && i.x0 < xmax && Math.abs(i.y - y) <= tol; })
      .sort(function (a, b) { return a.x0 - b.x0; });
    var out = "", last = null;
    sel.forEach(function (i) { if (last !== null && i.x0 - last > 1.2) out += " "; out += i.s; last = i.x1; });
    return out.trim();
  }

  function Rec(nazwisko, imie) {
    return { nazwisko: nazwisko.trim(), imie: imie.trim(), pdfName: titlePl(nazwisko + " " + imie),
      l4: 0, ppkPrac: 0, ppkFirma: 0, doWyplaty: 0, uwagi: [] };
  }

  // ---------------------------------------------------------------- Comarch Optima
  function parseComarch(items) {
    var hb = {}, sumY = null;
    items.forEach(function (i) {
      var p = Math.floor(i.y / 10000);
      if (i.s.trim() === "OPODAT.") hb[p] = Math.max(hb[p] || 0, i.y + 4);
      if (i.s.trim() === "PODSUMOWANIE:") sumY = i.y;
    });
    var endY = sumY == null ? Infinity : sumY;
    function inData(y) { var p = Math.floor(y / 10000); return y > (hb[p] || p * 10000 + 240) && y < endY; }

    var lps = items.filter(function (i) { return i.x1 < 36 && /^\d+$/.test(i.s.trim()) && inData(i.y); });
    var recs = lps.map(function (lp, k) {
      var b0 = lp.y - 4, b1 = k + 1 < lps.length ? lps[k + 1].y - 4 : endY;
      var bi = items.filter(function (i) { return i.y >= b0 && i.y < b1 && inData(i.y); });
      var nazwisko = lineText(bi, 36, 140, lp.y, 2).split("[")[0].trim();
      var imie = "";
      var cand = bi.filter(function (i) { return i.x0 >= 36 && i.x0 < 140 && i.y > lp.y + 4; });
      for (var c = 0; c < cand.length; c++) {
        var t = lineText(bi, 36, 140, cand[c].y, 1);
        if (t && !/\d/.test(t) && !/[\[\]]/.test(t)) { imie = t; break; }
      }
      var r = Rec(nazwisko, imie);
      bi.forEach(function (i) {
        var s = i.s.trim();
        if (s === "RAZEM:" && i.x0 > 470) { var v = numsIn(lineText(bi, 505, 570, i.y)); if (v.length) r.doWyplaty = v[0]; }
        else if (/^PPK\s*[PD]$/.test(s) && i.x0 > 305 && i.x0 < 322) r.ppkPrac += sum(numsIn(lineText(bi, 326, 368, i.y, 2.2)));
        else if (/^PPK\s*[PD]$/.test(s) && i.x0 > 366 && i.x0 < 380) r.ppkFirma += sum(numsIn(lineText(bi, 384, 422, i.y, 2.2)));
        else if (i.x0 >= 140 && i.x0 < 232 && /chorob/i.test(s)) r.l4 += sum(numsIn(lineText(bi, 225, 312, i.y, 2)));
      });
      return r;
    });

    var tot = null;
    if (sumY != null) {
      var si = items.filter(function (i) { return i.y >= sumY; });
      tot = { doWyplaty: 0, ppkPrac: 0, ppkFirma: 0 };
      si.forEach(function (i) {
        var s = i.s.trim();
        if (s === "RAZEM:") { var v = numsIn(lineText(si, 505, 570, i.y)); if (v.length) tot.doWyplaty = v[0]; }
        else if (/^PPK\s*[PD]$/.test(s) && i.x0 > 305 && i.x0 < 322) tot.ppkPrac += sum(numsIn(lineText(si, 326, 368, i.y, 2.2)));
        else if (/^PPK\s*[PD]$/.test(s) && i.x0 > 366 && i.x0 < 380) tot.ppkFirma += sum(numsIn(lineText(si, 384, 422, i.y, 2.2)));
      });
    }
    return { recs: recs, tot: tot };
  }

  // ---------------------------------------------------------------- Gratyfikant
  function parseGratyfikant(items) {
    var lefts = null;
    var pod = items.find(function (i) { return i.s.trim() === "Podstawa"; });
    if (pod) {
      var row = items.filter(function (i) { return Math.abs(i.y - pod.y) < 1.5; }).sort(function (a, b) { return a.x0 - b.x0; });
      if (row.length >= 8) lefts = row.slice(0, 8).map(function (i) { return i.x0; });
    }
    if (!lefts) lefts = [73.2, 155.8, 238.3, 320.9, 403.4, 486.0, 568.5, 651.1];
    var data = items.find(function (i) { return /^DATA/.test(i.s.trim()); });
    var bounds = lefts.concat([data ? data.x0 : lefts[7] + 85]);
    function colOf(x1) { for (var k = 0; k < 8; k++) if (x1 > bounds[k] && x1 <= bounds[k + 1] + 2) return k; return -1; }

    function grid(y0, first) {
      var g = {};
      items.forEach(function (i) {
        var s = i.s.trim();
        if (!NUM_FULL.test(s)) return;
        var d = i.y - y0 - first;
        if (d < -5 || d > 72) return;
        var r = Math.round(d / 11), k = colOf(i.x1);
        if (r >= 0 && r <= 6 && k >= 0 && !((r + "," + k) in g)) g[r + "," + k] = toNum(s);
      });
      return g;
    }
    function v(g, r, k) { return g[r + "," + k] || 0; }

    var recs = [];
    items.forEach(function (i) {
      if (/^\d+$/.test(i.s.trim()) && i.x1 < 60 && i.x0 > 25) {
        var name = items.filter(function (j) { return Math.abs(j.y - i.y) < 2 && j.x0 > 60 && j.x0 < bounds[1] + 40; })
          .sort(function (a, b) { return a.x0 - b.x0; }).map(function (j) { return j.s.trim(); }).join(" ");
        if (!name || /\d/.test(name)) return;
        var parts = name.split(/\s+/), nameItem = items.find(function (j) { return Math.abs(j.y - i.y) < 2 && j.x0 > 60; });
        var r = Rec(parts[0], parts.slice(1).join(" "));
        var g = grid(nameItem.y, 14);
        r.doWyplaty = v(g, 2, 7); r.ppkPrac = v(g, 3, 7); r.ppkFirma = v(g, 4, 7); r.l4 = v(g, 6, 0);
        var zas = v(g, 3, 6);
        if (Math.abs(zas) > 0.004) r.uwagi.push("Zasiłki (pozostałe) = " + zas.toFixed(2) + " – nie wliczone do L4, sprawdź ręcznie");
        if (!("2,7" in g)) r.uwagi.push("nie odczytano 'Do wypłaty'");
        recs.push(r);
      }
    });
    var tot = null;
    var raz = items.find(function (i) { return i.s.trim() === "Razem:"; });
    if (raz) { var g = grid(raz.y, 0.5); tot = { doWyplaty: v(g, 2, 7), ppkPrac: v(g, 3, 7), ppkFirma: v(g, 4, 7) }; }
    return { recs: recs, tot: tot };
  }

  var MIES = ["stycz", "lut", "marz", "kwie", "maj", "czerw", "lip", "sierp", "wrze", "pazdz", "listop", "grud"];
  var MIES_PL = ["styczeń", "luty", "marzec", "kwiecień", "maj", "czerwiec", "lipiec", "sierpień", "wrzesień",
    "październik", "listopad", "grudzień"];
  /** "2026-07" albo null */
  function detectPeriod(t) {
    var m = t.match(/Za okres:?\s*(\d{4})[.\-\/](\d{2})/i);
    if (m) return m[1] + "-" + m[2];
    m = norm(t).match(/miesiac\S*\s+([a-z]+)\s+(\d{4})/);
    if (m) for (var i = 0; i < 12; i++) if (m[1].indexOf(MIES[i]) === 0) return m[2] + "-" + ("0" + (i + 1)).slice(-2);
    return null;
  }
  function periodLabel(p) { if (!p) return "?"; var a = p.split("-"); return MIES_PL[+a[1] - 1] + " " + a[0]; }

  // ---------------------------------------------------------------- PDF → wynik
  async function parsePdf(pdfjsLib, data, fileName) {
    var L = await loadItems(pdfjsLib, data);
    var res, fmt;
    if (/\[AKRONIM\]|SZCZEG|Comarch/.test(L.first)) { res = parseComarch(L.items); fmt = "Comarch Optima"; }
    else if (/Nazwisko i imi|Gratyfikant/.test(L.first)) { res = parseGratyfikant(L.items); fmt = "Gratyfikant"; }
    else throw new Error("Nieznany format listy płac: " + fileName);
    var firstLine = L.items.length ? lineText(L.items, 0, 400, L.items[0].y, 2) : "";
    var checks = [], ok = !!res.tot;
    [["doWyplaty", "Do wypłaty"], ["ppkPrac", "PPK pracownika"], ["ppkFirma", "PPK pracodawcy"]].forEach(function (k) {
      var s = r2(sum(res.recs.map(function (r) { return r[k[0]]; })));
      if (!res.tot) return;
      var t = r2(res.tot[k[0]]), good = Math.abs(s - t) < 0.011;
      if (!good) ok = false;
      checks.push({ label: k[1], read: s, total: t, ok: good });
    });
    return { file: fileName, format: fmt, firma: firstLine, okres: detectPeriod(L.first), recs: res.recs,
      checks: checks, ok: ok, hasTotals: !!res.tot };
  }

  // ---------------------------------------------------------------- dopasowanie nazwisk
  function matchName(rec, names) {
    var byNorm = {};
    names.forEach(function (n) { if (!(norm(n) in byNorm)) byNorm[norm(n)] = n; });
    var key = norm(rec.pdfName);
    if (key in byNorm) return { name: byNorm[key], note: "" };
    var rev = norm(rec.imie + " " + rec.nazwisko);
    if (rev in byNorm) return { name: byNorm[rev], note: "" };
    var sur = norm(rec.nazwisko), first = norm(rec.imie);
    var same = names.filter(function (n) { return norm(n).split(" ")[0] === sur; });
    var best = null, score = 0;
    same.forEach(function (n) { var sc = ratio(norm(n).split(" ").slice(1).join(" "), first); if (sc > score) { best = n; score = sc; } });
    if (best && (score >= 0.75 || (same.length === 1 && (score >= 0.5 || first.slice(0, 2) === norm(best).split(" ")[1].slice(0, 2)))))
      return { name: best, note: "przybliżone: PDF „" + rec.pdfName + "” → „" + best + "”", fuzzy: true };
    var bk = null, bs = 0;
    Object.keys(byNorm).forEach(function (k) { var sc = ratio(k, key); if (sc > bs) { bs = sc; bk = k; } });
    if (bk && bs >= 0.88) return { name: byNorm[bk], note: "przybliżone: PDF „" + rec.pdfName + "” → „" + byNorm[bk] + "”", fuzzy: true };
    return { name: null, note: "„" + rec.pdfName + "” – brak w arkuszach, wpisano nazwisko z PDF", missing: true };
  }

  /** parsed: wyniki parsePdf; names: lista nazw z Excela; etat: {nazwa: etat} z arkusza xls */
  function build(parsed, names, etat) {
    var agg = {}, order = [], notes = [];
    parsed.forEach(function (pf) {
      pf.recs.forEach(function (r) {
        var m = matchName(r, names), nm = m.name || r.pdfName;
        if (m.note) notes.push({ type: m.missing ? "err" : "warn", text: m.note, file: pf.file });
        r.uwagi.forEach(function (u) { notes.push({ type: "warn", text: nm + ": " + u, file: pf.file }); });
        if (agg[nm]) {
          notes.push({ type: "warn", text: nm + " jest na kilku listach – zsumowano", file: pf.file });
          agg[nm].l4 += r.l4; agg[nm].ppkPrac += r.ppkPrac; agg[nm].ppkFirma += r.ppkFirma; agg[nm].doWyplaty += r.doWyplaty;
        } else {
          agg[nm] = { name: nm, l4: r.l4, ppkPrac: r.ppkPrac, ppkFirma: r.ppkFirma, doWyplaty: r.doWyplaty,
            file: pf.file, fuzzy: !!m.fuzzy, missing: !!m.missing };
          order.push(nm);
        }
      });
    });
    var rows = order.map(function (n) { return agg[n]; })
      .sort(function (a, b) { return norm(a.name) < norm(b.name) ? -1 : norm(a.name) > norm(b.name) ? 1 : 0; });
    var missing = Object.keys(etat || {}).filter(function (n) {
      var e = etat[n];
      return e != null && e !== "" && e !== " " && e !== "BRAK" && !(typeof e === "string" && e.charAt(0) === "#") && !agg[n];
    });
    return { rows: rows, notes: notes, missing: missing };
  }

  // ---------------------------------------------------------------- wklejka z rozliczenia → xls
  function parseAmount(s) {
    s = String(s || "").replace(/[\s\u00a0]/g, "");
    if (!s) return NaN;
    if (/,/.test(s)) s = s.replace(/\./g, "").replace(",", ".");
    return /^-?\d+(\.\d+)?$/.test(s) ? parseFloat(s) : NaN;
  }
  /** Tekst rozdzielany tabulatorami: kol. 3 = Dział, 4 = Pracownik, 18 = E2 (1-based). */
  function parsePaste(text, opts) {
    opts = opts || {};
    var cD = (opts.colDzial || 3) - 1, cP = (opts.colName || 4) - 1, cE = (opts.colE2 || 18) - 1;
    var rows = [], zero = [], bad = [], dup = [], seen = {};
    String(text || "").split(/\r?\n/).forEach(function (line, i) {
      if (!line.trim()) return;
      var f = line.split("\t");
      if (f.length < 4) return;                      // nagłówki typu "Kredyty"
      var name = (f[cP] || "").trim().replace(/\s+/g, " "), dz = (f[cD] || "").trim(), e2 = parseAmount(f[cE]);
      if (!name || isNaN(e2)) { bad.push({ line: i + 1, text: line.slice(0, 60) }); return; }
      var k = norm(name);
      if (seen[k]) { dup.push(name); return; }
      seen[k] = 1;
      var r = { dzial: dz, name: name, e2: r2(e2) };
      if (Math.abs(e2) < 0.005 && opts.skipZero !== false) zero.push(r); else rows.push(r);
    });
    return { rows: rows, zero: zero, bad: bad, dup: dup };
  }

  var api = { parsePdf: parsePdf, parsePaste: parsePaste, detectPeriod: detectPeriod, periodLabel: periodLabel, build: build, matchName: matchName, norm: norm, ratio: ratio, r2: r2 };
  if (typeof module !== "undefined" && module.exports) module.exports = api; else root.ListyPlac = api;
})(this);
