/* Listy płac → Excel – logika panelu dodatku */
(function () {
  "use strict";
  var VERSION = "1.1";
  var TABLE = "zlistplac";
  var NAME_TABLES = ["XLS", "Pracownicy_Etat"];
  var SETTING_PERIOD = "llp.okres";
  var LP = window.ListyPlac;

  var state = { inExcel: false, files: [], names: [], etat: {}, built: null, undo: null, busy: false, period: null };
  var $ = function (id) { return document.getElementById(id); };

  if (window.pdfjsLib) pdfjsLib.GlobalWorkerOptions.workerSrc = "vendor/pdf.worker.min.js";
  $("ver").textContent = "v" + VERSION;

  function fmt(n) { return (Math.round(n * 100) / 100).toLocaleString("pl-PL", { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }
  function esc(s) { return String(s).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; }); }

  function banner(type, html, action) {
    var b = $("banner");
    if (!type) { b.className = "banner"; b.innerHTML = ""; return; }
    b.className = "banner show " + type;
    b.innerHTML = html;
    if (action) {
      var btn = document.createElement("button"); btn.className = "btn"; btn.textContent = action.label;
      btn.onclick = action.run; b.appendChild(btn);
    }
  }
  function setStatus(html, bad) { var s = $("status"); s.innerHTML = html; s.className = "status" + (bad ? " bad" : ""); }

  // ------------------------------------------------------------------ Excel
  function excelRun(fn) {
    return Excel.run(fn).catch(function (e) {
      console.error(e);
      var msg = (e && e.debugInfo && e.debugInfo.message) || (e && e.message) || String(e);
      banner("err", "Excel zgłosił błąd: " + esc(msg));
      throw e;
    });
  }

  function readWorkbook() {
    return excelRun(async function (ctx) {
      var tables = ctx.workbook.tables; tables.load("items/name");
      await ctx.sync();
      var have = tables.items.map(function (t) { return t.name; });
      var res = { hasTarget: have.indexOf(TABLE) >= 0, hasXls: have.indexOf("XLS") >= 0, names: [], etat: {}, peNames: [],
        rowsInTable: 0, xlsFilled: 0, xlsEmpty: 0 };
      var loads = [];
      NAME_TABLES.forEach(function (tn) {
        if (have.indexOf(tn) < 0) return;
        var t = tables.getItem(tn), h = t.getHeaderRowRange(), b = t.getDataBodyRange();
        h.load("values"); b.load("values"); loads.push({ tn: tn, h: h, b: b });
      });
      var tb = null;
      if (res.hasTarget) { tb = tables.getItem(TABLE).getDataBodyRange(); tb.load("values"); }
      await ctx.sync();
      loads.forEach(function (L) {
        var hdr = L.h.values[0], ip = hdr.indexOf("Pracownik"), ie = hdr.indexOf("Etat");
        if (ip < 0) return;
        L.b.values.forEach(function (r) {
          var n = typeof r[ip] === "string" ? r[ip].trim() : "";
          if (!n) return;
          if (res.names.indexOf(n) < 0) res.names.push(n);
          if (L.tn === "XLS" && ie >= 0) res.etat[n] = r[ie];
          if (L.tn === "Pracownicy_Etat") res.peNames.push(n);
        });
      });
      loads.forEach(function (L) {
        if (L.tn !== "XLS") return;
        var hdr = L.h.values[0], ip = hdr.indexOf("Pracownik"), id = hdr.indexOf("Dział"), ix = hdr.indexOf("E2");
        L.b.values.forEach(function (r) {
          if (String(r[ip] || "").trim()) res.xlsFilled++;
          else if (!String(r[id] || "").trim() && (r[ix] === "" || r[ix] == null)) res.xlsEmpty++;
        });
      });
      if (tb) res.rowsInTable = tb.values.filter(function (r) { return String(r[0] || "").trim(); }).length;
      return res;
    });
  }

  /** Zapisuje całe ciało tabeli (wiersze po 5 kolumn). Zwraca poprzednią zawartość. */
  function writeBody(rows) {
    return excelRun(async function (ctx) {
      var t = ctx.workbook.tables.getItem(TABLE);
      var body = t.getDataBodyRange(), hdr = t.getHeaderRowRange();
      body.load("values,rowCount"); hdr.load("address");
      await ctx.sync();
      var prev = body.values.map(function (r) { return r.slice(); });
      var data = rows.length ? rows : [["", "", "", "", ""]];
      var n = data.length, cur = body.rowCount;
      body.clear(Excel.ClearApplyTo.contents);
      if (n !== cur) {
        if (Office.context.requirements.isSetSupported("ExcelApi", "1.13")) {
          t.resize(hdr.getResizedRange(n, 0));
        } else if (n > cur) {
          var blank = []; for (var i = 0; i < n - cur; i++) blank.push(["", "", "", "", ""]);
          t.rows.add(null, blank);
        } else {
          for (var j = cur - 1; j >= n; j--) t.rows.getItemAt(j).delete();
        }
        await ctx.sync();
      }
      var nb = t.getDataBodyRange();
      nb.values = data;
      nb.getColumn(1).getResizedRange(0, 3).numberFormat = data.map(function () { return ["#,##0.00", "#,##0.00", "#,##0.00", "#,##0.00"]; });
      await ctx.sync();
      return prev;
    });
  }

  function readBody() {
    return excelRun(async function (ctx) {
      var b = ctx.workbook.tables.getItem(TABLE).getDataBodyRange(); b.load("values");
      await ctx.sync(); return b.values;
    });
  }

  function selectName(name, tableName) {
    if (!state.inExcel) return;
    excelRun(async function (ctx) {
      var t = ctx.workbook.tables.getItem(tableName), b = t.getDataBodyRange(), h = t.getHeaderRowRange();
      b.load("values"); h.load("values"); await ctx.sync();
      var ip = Math.max(0, h.values[0].indexOf("Pracownik"));
      for (var i = 0; i < b.values.length; i++) {
        if (String(b.values[i][ip]).trim() === name) { t.worksheet.activate(); b.getRow(i).select(); break; }
      }
      await ctx.sync();
    });
  }

  function getSetting(k) { try { return Office.context.document.settings.get(k); } catch (e) { return null; } }
  function setSetting(k, v) {
    try { Office.context.document.settings.set(k, v); Office.context.document.settings.saveAsync(function () {}); } catch (e) {}
  }

  // ------------------------------------------------------------------ import
  function readFile(f) {
    return new Promise(function (res, rej) {
      var r = new FileReader();
      r.onload = function () { res(new Uint8Array(r.result)); };
      r.onerror = function () { rej(r.error); };
      r.readAsArrayBuffer(f);
    });
  }

  async function handleFiles(list) {
    var pdfs = Array.prototype.filter.call(list || [], function (f) { return /\.pdf$/i.test(f.name) || f.type === "application/pdf"; });
    if (!pdfs.length) { banner("warn", "To nie wygląda na PDF. Upuść listę płac w formacie PDF."); return; }
    if (state.busy) return;
    state.busy = true; banner(null);
    $("drop").classList.add("busy");
    $("dropBig").textContent = "Czytam " + (pdfs.length > 1 ? pdfs.length + " pliki…" : pdfs[0].name + "…");
    var errors = [];
    for (var i = 0; i < pdfs.length; i++) {
      try {
        var data = await readFile(pdfs[i]);
        var parsed = await LP.parsePdf(pdfjsLib, data, pdfs[i].name);
        if (!parsed.recs.length) throw new Error("nie znalazłem pracowników");
        state.files = state.files.filter(function (x) { return x.file !== parsed.file; });
        state.files.push(parsed);
      } catch (e) { console.error(e); errors.push(pdfs[i].name + ": " + (e.message || e)); }
    }
    try { await apply(); }
    catch (e) { console.error(e); }
    finally {
      state.busy = false; $("drop").classList.remove("busy");
      $("dropBig").textContent = state.files.length ? "Upuść kolejną listę" : "Upuść listy płac (PDF)";
      $("dropSmall").textContent = "albo kliknij, żeby wybrać — można kilka naraz";
      if (errors.length) banner("err", "Nie udało się odczytać:<br>" + errors.map(esc).join("<br>"));
    }
  }

  function sessionPeriod() {
    var ps = state.files.map(function (f) { return f.okres; }).filter(Boolean);
    var uniq = ps.filter(function (p, i) { return ps.indexOf(p) === i; }).sort();
    return { period: uniq.length ? uniq[uniq.length - 1] : null, mixed: uniq.length > 1, all: uniq };
  }

  async function apply() {
    state.built = LP.build(state.files, state.names, state.etat);
    var sp = sessionPeriod();
    state.period = sp.period;
    if (sp.mixed) banner("warn", "Listy są z różnych okresów: " + sp.all.map(LP.periodLabel).join(", ") + ". Sprawdź, czy to właściwe pliki.");

    if (state.inExcel) {
      var current = await readBody();
      var stored = getSetting(SETTING_PERIOD);
      var existing = current.filter(function (r) { return String(r[0] || "").trim(); });
      var newPeriod = sp.period && stored && stored !== sp.period && existing.length;
      var base = newPeriod ? [] : existing;
      var map = {}, order = [];
      base.forEach(function (r) { var k = LP.norm(String(r[0])); if (!map[k]) order.push(k); map[k] = r.slice(0, 5); });
      state.built.rows.forEach(function (r) {
        var k = LP.norm(r.name);
        if (!map[k]) order.push(k);
        map[k] = [r.name, LP.r2(r.l4), LP.r2(r.ppkPrac), LP.r2(r.ppkFirma), LP.r2(r.doWyplaty)];
      });
      var out = order.sort().map(function (k) { return map[k]; });
      var prev = await writeBody(out);
      state.undo = { kind: "pdf", rows: prev, period: stored };
      $("btnUndo").disabled = false;
      if (sp.period) setSetting(SETTING_PERIOD, sp.period);
      state.tableRows = out;
      if (newPeriod) banner("info", "Nowy okres: <b>" + esc(LP.periodLabel(sp.period)) + "</b> — usunąłem dane z " +
        esc(LP.periodLabel(stored)) + " (" + existing.length + " wierszy).", { label: "Cofnij", run: undo });
      setStatus("Tabela <b>" + TABLE + "</b>: " + out.length + " osób · " + state.names.length + " pracowników w arkuszach");
    }
    render();
  }

  async function undo() {
    if (!state.undo) return;
    var u = state.undo; state.undo = null;
    if (u.kind === "xls") return undoXls(u);
    await writeBody(u.rows);
    if (u.period) setSetting(SETTING_PERIOD, u.period);
    state.files = [];
    $("btnUndo").disabled = true;
    banner("info", "Przywrócono poprzednią zawartość tabeli.");
    state.built = null; render();
    var wb = await readWorkbook(); setStatus("Tabela <b>" + TABLE + "</b>: " + wb.rowsInTable + " osób");
  }

  var clearArmed = null;
  async function clearTable() {
    if (!clearArmed) {
      $("btnClear").textContent = "Na pewno? Kliknij ponownie";
      clearArmed = setTimeout(function () { clearArmed = null; $("btnClear").textContent = "Wyczyść tabelę"; }, 3500);
      return;
    }
    clearTimeout(clearArmed); clearArmed = null; $("btnClear").textContent = "Wyczyść tabelę";
    if (!state.inExcel) return;
    var prev = await writeBody([]);
    state.undo = { kind: "pdf", rows: prev, period: getSetting(SETTING_PERIOD) };
    $("btnUndo").disabled = false;
    state.files = []; state.built = null; render();
    setStatus("Tabela <b>" + TABLE + "</b> wyczyszczona.");
    banner("info", "Tabela wyczyszczona.", { label: "Cofnij", run: undo });
  }

  function firmName(t) {
    var u = t.toUpperCase();
    if (/VISUALQ/.test(u)) return "VisualQ s.c.";
    if (/LIGHT MEDIA/.test(u)) return "Light Media s.c.";
    if (/LIGHT GROUP/.test(u)) return "Light Group s.c.";
    if (/LIGHT (SPÓŁKA CYWILNA|S\.?C\.?)/.test(u)) return "Light s.c.";
    return t.replace(/\s+Data wydruku.*$/, "").split(/\s+/).slice(0, 3).join(" ");
  }

  // ------------------------------------------------------------------ render
  function render() {
    var F = state.files;
    $("filesSec").classList.toggle("hidden", !F.length);
    $("statsSec").classList.toggle("hidden", !F.length);
    $("issuesSec").classList.toggle("hidden", !F.length);
    if (!F.length) return;
    $("period").textContent = state.period ? LP.periodLabel(state.period) : "";
    $("files").innerHTML = "";
    F.forEach(function (f) {
      var d = document.createElement("div");
      d.className = "file" + (!f.hasTotals ? " nototal" : !f.ok ? " bad" : "");
      var check = f.hasTotals ? (f.ok ? "sumy zgodne z podsumowaniem ✓" : "") : "brak podsumowania – nie sprawdzono sum";
      var firm = firmName(f.firma || "");
      d.innerHTML = '<span class="dot"></span><div class="fname" title="' + esc(f.file) + '">' + esc(firm || f.file) + "</div>" +
        '<button class="x" title="Usuń z listy">×</button>' +
        '<div class="fmeta">' + f.recs.length + " os. · " + esc(f.format) + (check ? " · " + check : "") + "</div>";
      f.checks.filter(function (c) { return !c.ok; }).forEach(function (c) {
        var e = document.createElement("div"); e.className = "chk";
        e.textContent = c.label + ": odczytano " + fmt(c.read) + ", na liście " + fmt(c.total);
        d.appendChild(e);
      });
      d.querySelector(".x").onclick = function () {
        state.files = state.files.filter(function (x) { return x !== f; });
        banner("info", "Usunięto „" + esc(f.file) + "” z panelu. Dane tej listy zostały w tabeli – użyj „Cofnij” albo „Wyczyść tabelę”, jeśli trzeba.");
        state.built = LP.build(state.files, state.names, state.etat); render();
      };
      $("files").appendChild(d);
    });

    var rows = (state.built && state.built.rows) || [];
    var s = function (k) { return rows.reduce(function (a, r) { return a + r[k]; }, 0); };
    $("sDw").textContent = fmt(s("doWyplaty")) + " zł";
    $("sL4").textContent = fmt(s("l4"));
    $("sPpkP").textContent = fmt(s("ppkPrac"));
    $("sPpkF").textContent = fmt(s("ppkFirma"));
    $("rowsInfo").textContent = rows.length + " osób z " + F.length + (F.length === 1 ? " listy" : " list");

    var box = $("issues"); box.innerHTML = "";
    var notes = (state.built && state.built.notes) || [], seen = {};
    function add(cls, icon, html, name, table) {
      var d = document.createElement("div"); d.className = "issue " + cls;
      d.innerHTML = '<span class="ic">' + icon + "</span><span>" + html + "</span>";
      if (name) { var a = d.querySelector("a"); if (a) a.onclick = function () { selectName(name, table); }; }
      box.appendChild(d);
    }
    notes.forEach(function (n) {
      if (seen[n.text]) return; seen[n.text] = 1;
      var row = rows.find(function (r) { return n.text.indexOf("„" + r.name + "”") >= 0 || n.text.indexOf(r.name + ":") === 0 || n.text.indexOf(r.name + " jest") === 0; });
      var html = esc(n.text);
      if (row) html = html.replace(esc(row.name), "<a>" + esc(row.name) + "</a>");
      add(n.type === "err" ? "err" : "warn", n.type === "err" ? "!" : "?", html, row && row.name, TABLE);
    });
    var miss = (state.built && state.built.missing) || [];
    if (miss.length && state.inExcel) {
      var det = document.createElement("details"); det.className = "issue miss group";
      det.innerHTML = '<summary><span class="ic">–</span><span>' + miss.length +
        (miss.length === 1 ? " osoba" : miss.length < 5 ? " osoby" : " osób") +
        " z etatem w arkuszu xls nie ma na wczytanych listach (VLOOKUP pokaże #N/A)</span></summary>";
      var ul = document.createElement("div"); ul.className = "misslist";
      miss.forEach(function (m) {
        var a = document.createElement("a"); a.textContent = m; a.onclick = function () { selectName(m, "XLS"); }; ul.appendChild(a);
      });
      det.appendChild(ul); box.appendChild(det);
    }
    var cnt = box.children.length;
    $("issuesCount").textContent = cnt ? cnt : "";
    if (!cnt) box.innerHTML = '<div class="ok-line">Wszystko dopasowane, nic do sprawdzenia ✓</div>';
  }

  // ------------------------------------------------------------------ zakładka: zestawienie → xls
  var XLS = "XLS";
  function xStatus(wb) {
    if (!wb.hasXls) { $("xstatus").innerHTML = "<b>W tym skoroszycie nie ma tabeli „XLS”.</b>"; $("xstatus").className = "status bad"; return; }
    $("xstatus").className = "status";
    $("xstatus").innerHTML = "Tabela <b>XLS</b>: " + wb.xlsFilled + " osób · " + wb.xlsEmpty + " pustych wierszy";
  }

  function fillXls(rows) {
    return excelRun(async function (ctx) {
      var t = ctx.workbook.tables.getItem(XLS), h = t.getHeaderRowRange(), b = t.getDataBodyRange();
      h.load("values"); b.load("values,formulas,rowCount");
      await ctx.sync();
      var hdr = h.values[0], iD = hdr.indexOf("Dział"), iP = hdr.indexOf("Pracownik"), iE = hdr.indexOf("E2");
      if (iD < 0 || iP < 0 || iE < 0) throw new Error("w tabeli XLS brakuje kolumny Dział, Pracownik albo E2");
      var vals = b.values, frm = b.formulas, n = b.rowCount;
      var cols = [iD, iP, iE];
      var snap = frm.map(function (r) { return cols.map(function (c) { return r[c]; }); });
      var next = snap.map(function (r) { return r.slice(); });
      var byName = {}, empty = [];
      vals.forEach(function (r, i) {
        var nm = String(r[iP] || "").trim();
        if (nm) byName[LP.norm(nm)] = i;
        else if (!String(r[iD] || "").trim() && (r[iE] === "" || r[iE] == null)) empty.push(i);
      });
      var out = { filled: [], appended: [], already: [], updated: [], conflicts: [] };
      var toAppend = [];
      rows.forEach(function (r) {
        var k = LP.norm(r.name);
        if (k in byName) {
          var i = byName[k], cur = vals[i][iE];
          if (cur === "" || cur == null) {
            next[i][2] = r.e2; if (!String(vals[i][iD] || "").trim()) next[i][0] = r.dzial;
            out.updated.push(r);
          } else if (typeof cur === "number" && Math.abs(cur - r.e2) > 0.005) {
            out.conflicts.push({ name: vals[i][iP], cur: cur, e2: r.e2 });
          } else out.already.push(r);
        } else if (empty.length) {
          var j = empty.shift(); next[j] = [r.dzial, r.name, r.e2]; byName[k] = j; out.filled.push(r);
        } else { toAppend.push(r); out.appended.push(r); }
      });
      if (out.filled.length || out.updated.length) {
        cols.forEach(function (c, j) { b.getColumn(c).values = next.map(function (r) { return [r[j]]; }); });
      }
      if (toAppend.length) {
        var tpl = frm[0];
        t.rows.add(null, toAppend.map(function (r) {
          return tpl.map(function (f, c) {
            if (c === iD) return r.dzial; if (c === iP) return r.name; if (c === iE) return r.e2;
            return typeof f === "string" && f.charAt(0) === "=" ? f : "";
          });
        }));
      }
      await ctx.sync();
      out.snap = snap; out.n = n; out.cols = cols;
      return out;
    });
  }

  async function undoXls(u) {
    await excelRun(async function (ctx) {
      var t = ctx.workbook.tables.getItem(XLS), b = t.getDataBodyRange();
      for (var i = u.n + u.appended - 1; i >= u.n; i--) t.rows.getItemAt(i).delete();
      await ctx.sync();
      u.cols.forEach(function (c, j) { b.getColumn(c).values = u.snap.map(function (r) { return [r[j]]; }); });
      await ctx.sync();
    });
    $("btnUndo").disabled = true;
    $("xResSec").classList.add("hidden"); $("xIssuesSec").classList.add("hidden");
    banner("info", "Przywrócono poprzednią zawartość zakładki xls.");
    await refreshWorkbook();
  }

  async function refreshWorkbook() {
    var wb = await readWorkbook();
    state.names = wb.names; state.etat = wb.etat; state.peNames = wb.peNames;
    xStatus(wb);
    setStatus("Tabela <b>" + TABLE + "</b>: " + wb.rowsInTable + " osób · " + wb.names.length + " pracowników w arkuszach");
    return wb;
  }

  var pasteBusy = false;
  async function handlePaste() {
    var text = $("paste").value;
    if (!text.trim() || pasteBusy) return;
    var p = LP.parsePaste(text, { skipZero: $("skipZero").checked });
    banner(null);
    if (!p.rows.length) {
      banner("warn", "Nie rozpoznałem żadnych wierszy. Oczekuję danych rozdzielonych tabulatorami (Dział w 3., pracownik w 4., kwota w 18. kolumnie).");
      return;
    }
    if (!state.inExcel) { renderXls(p, null); return; }
    pasteBusy = true; $("pasteBox").classList.add("busy");
    try {
      var res = await fillXls(p.rows);
      state.undo = { kind: "xls", snap: res.snap, n: res.n, cols: res.cols, appended: res.appended.length };
      $("btnUndo").disabled = false;
      await refreshWorkbook();
      renderXls(p, res);
      $("paste").value = "";
    } catch (e) {
      console.error(e);
      banner("err", "Nie udało się wpisać danych: " + esc(e.message || e));
    } finally { pasteBusy = false; $("pasteBox").classList.remove("busy"); }
  }

  function renderXls(p, res) {
    var box = $("xIssues"); box.innerHTML = "";
    function add(cls, icon, html, name) {
      var d = document.createElement("div"); d.className = "issue " + cls;
      d.innerHTML = '<span class="ic">' + icon + "</span><span>" + html + "</span>";
      var a = d.querySelector("a"); if (a && name) a.onclick = function () { selectName(name, XLS); };
      box.appendChild(d);
    }
    function group(cls, icon, title, names) {
      var det = document.createElement("details"); det.className = "issue " + cls + " group";
      det.innerHTML = '<summary><span class="ic">' + icon + "</span><span>" + title + "</span></summary>";
      var ul = document.createElement("div"); ul.className = "misslist";
      names.forEach(function (m) { var a = document.createElement("a"); a.textContent = m; a.onclick = function () { selectName(m, XLS); }; ul.appendChild(a); });
      det.appendChild(ul); box.appendChild(det);
    }
    var written = res ? res.filled.concat(res.appended, res.updated) : p.rows;
    $("xResSec").classList.remove("hidden");
    $("xNew").textContent = res ? String(res.filled.length + res.appended.length) : String(p.rows.length);
    $("xSum").textContent = fmt(written.reduce(function (a, r) { return a + r.e2; }, 0));
    $("xInfo").textContent = p.rows.length + " rozpoznanych" + (p.zero.length ? " + " + p.zero.length + " z zerem" : "");

    if (res) {
      res.conflicts.forEach(function (c) {
        add("warn", "?", "<a>" + esc(c.name) + "</a>: w tabeli " + fmt(c.cur) + ", we wklejce " + fmt(c.e2) + " – nie nadpisano", c.name);
      });
      var pe = {}; (state.peNames || []).forEach(function (n) { pe[LP.norm(n)] = 1; });
      written.forEach(function (r) {
        if ((state.peNames || []).length && !pe[LP.norm(r.name)])
          add("err", "!", "<a>" + esc(r.name) + "</a> – brak w „Przypisanie do firm”, formuły w tym wierszu pokażą #N/A", r.name);
      });
      if (res.updated.length) group("warn", "+", "Uzupełniono E2 u " + res.updated.length + " os. już obecnych w tabeli", res.updated.map(function (r) { return r.name; }));
      if (res.appended.length) add("warn", "+", "Zabrakło pustych wierszy – dodano " + res.appended.length + " do tabeli XLS (z formułami)");
      if (res.already.length) group("miss", "=", res.already.length + " os. już było w tabeli z tą samą kwotą – bez zmian", res.already.map(function (r) { return r.name; }));
    }
    if (p.zero.length) group("miss", "0", "Pominięto " + p.zero.length + " os. z kwotą 0,00", p.zero.map(function (r) { return r.name; }));
    if (p.dup.length) add("warn", "?", "Powtórzone w wklejce (wzięto pierwsze wystąpienie): " + esc(p.dup.join(", ")));
    if (p.bad.length) add("warn", "?", "Nie rozpoznano " + p.bad.length + " wierszy, np. wiersz " + p.bad[0].line + ": „" + esc(p.bad[0].text) + "”");

    var cnt = box.children.length;
    $("xIssuesSec").classList.remove("hidden");
    $("xIssuesCount").textContent = cnt || "";
    if (!cnt) box.innerHTML = '<div class="ok-line">Wszystko wpisane, nic do sprawdzenia ✓</div>';
    if (!res) banner("info", "Tryb podglądu – poza Excelem nic nie zostało zapisane.");
  }

  $("paste").addEventListener("paste", function () { setTimeout(handlePaste, 0); });
  $("paste").addEventListener("drop", function () { setTimeout(handlePaste, 0); });
  $("btnShowXls").onclick = function () {
    if (!state.inExcel) return;
    excelRun(async function (ctx) { var t = ctx.workbook.tables.getItem(XLS); t.worksheet.activate(); t.getRange().select(); await ctx.sync(); });
  };

  // ------------------------------------------------------------------ zakładki
  function showTab(which) {
    var pdf = which !== "xls";
    $("viewPdf").classList.toggle("hidden", !pdf); $("viewXls").classList.toggle("hidden", pdf);
    $("tabPdf").setAttribute("aria-selected", String(pdf)); $("tabXls").setAttribute("aria-selected", String(!pdf));
    try { localStorage.setItem("llp.tab", pdf ? "pdf" : "xls"); } catch (e) {}
    if (!pdf) setTimeout(function () { $("paste").focus(); }, 0);
  }
  $("tabPdf").onclick = function () { showTab("pdf"); };
  $("tabXls").onclick = function () { showTab("xls"); };
  try { if (localStorage.getItem("llp.tab") === "xls") showTab("xls"); } catch (e) {}

  // ------------------------------------------------------------------ UI
  var drop = $("drop"), input = $("file");
  input.addEventListener("change", function () { handleFiles(input.files); input.value = ""; });
  drop.addEventListener("keydown", function (e) { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); input.click(); } });
  var depth = 0;
  document.addEventListener("dragenter", function (e) { if (e.target && e.target.id === "paste") return; e.preventDefault(); depth++; drop.classList.add("over"); });
  document.addEventListener("dragleave", function () { if (--depth <= 0) { depth = 0; drop.classList.remove("over"); } });
  document.addEventListener("dragover", function (e) { if (e.target && e.target.id === "paste") return; e.preventDefault(); if (e.dataTransfer) e.dataTransfer.dropEffect = "copy"; });
  document.addEventListener("drop", function (e) {
    if (e.target && e.target.id === "paste") { depth = 0; drop.classList.remove("over"); return; }
    e.preventDefault(); depth = 0; drop.classList.remove("over");
    if (e.dataTransfer && e.dataTransfer.files) handleFiles(e.dataTransfer.files);
  });
  $("btnUndo").onclick = undo;
  $("btnClear").onclick = clearTable;
  $("btnShow").onclick = function () {
    if (!state.inExcel) return;
    excelRun(async function (ctx) {
      var t = ctx.workbook.tables.getItem(TABLE); t.worksheet.activate(); t.getRange().select(); await ctx.sync();
    });
  };

  async function init(info) {
    state.inExcel = !!(info && info.host === Office.HostType.Excel);
    if (!state.inExcel) {
      setStatus("Tryb podglądu (poza Excelem) — dane nie będą zapisywane.");
      $("xstatus").textContent = "Tryb podglądu (poza Excelem) — dane nie będą zapisywane.";
      ["btnUndo", "btnShow", "btnClear", "btnShowXls"].forEach(function (id) { $(id).disabled = true; });
      return;
    }
    if (!Office.context.requirements.isSetSupported("ExcelApi", "1.4")) {
      setStatus("<b>Ta wersja Excela jest za stara dla dodatku.</b>", true); return;
    }
    try {
      var wb = await readWorkbook();
      state.names = wb.names; state.etat = wb.etat; state.peNames = wb.peNames;
      xStatus(wb);
      if (!wb.hasTarget) {
        setStatus("<b>W tym skoroszycie nie ma tabeli „" + TABLE + "”.</b> Otwórz zestawienie wypłat.", true);
        drop.classList.add("busy");
        return;
      }
      setStatus("Tabela <b>" + TABLE + "</b>: " + wb.rowsInTable + " osób · " + wb.names.length + " pracowników w arkuszach");
    } catch (e) { setStatus("<b>Nie udało się odczytać skoroszytu.</b>", true); }
  }

  if (window.Office && Office.onReady) Office.onReady(init);
  else init(null);
})();
