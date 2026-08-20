/* =====================================================================
   Breathe Villa Resolution Report
   Every figure on this page is reproducible against the original export.
   The Excel formulas rendered under each metric reference the RAW sheet
   (header on row 3, data in rows 4:1180), not this dashboard's own data.
   ===================================================================== */
(function () {
  "use strict";

  var DATA = window.TICKET_DATA;
  var META = DATA.meta;
  var TICKETS = DATA.tickets;

  /* --- geometry of the source workbook, used to build every formula --- */
  var XL = {
    first: 4,
    last: 3 + TICKETS.length,
    col: { created: "B", category: "D", flat: "G", status: "J", closed: "L", esc: "N", rtat: "R" },
    mins: "W",       // helper column the reader adds
    cutoff: "Z$1",   // cell holding the open-ticket cutoff
    uniq: "AA",      // unique villa list
    avg: "AB"        // per-villa average
  };
  XL.uniqLast = 3 + META.flats;

  var ESC = { "Level-1": "Urgent", "NA": "Usual" };
  var ESC_KEYS = ["Level-1", "NA"];

  /* ========================= formatting ========================= */
  function dhm(m) {
    if (m == null || isNaN(m)) return "—";
    m = Math.round(m);
    var d = Math.floor(m / 1440), h = Math.floor((m % 1440) / 60), mi = m % 60;
    if (d > 0) return d + "d " + h + "h";
    if (h > 0) return h + "h " + mi + "m";
    return mi + "m";
  }
  function days(m) { return m == null ? "—" : (m / 1440).toFixed(1); }
  function n0(x) { return Math.round(x).toLocaleString("en-US"); }
  function esc(s) {
    return String(s).replace(/[&<>"']/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c];
    });
  }
  function dt(iso) {
    if (!iso) return "—";
    return new Date(iso).toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" });
  }
  function num(x) { return +x.replace(/\D/g, "") || 0; }
  /* Category values carry stray whitespace in the export ("Gardening "). Excel
     criteria must use the raw value, so only the label is trimmed. */
  function catLabel(c) { return c.trim(); }
  function hasPadding(c) { return c !== c.trim(); }

  /* ========================= aggregation ========================= */
  function stats(list) {
    if (!list.length) return { n: 0, mean: null, median: null, min: null, max: null };
    var v = list.map(function (t) { return t.mins; }).sort(function (a, b) { return a - b; });
    var sum = v.reduce(function (a, b) { return a + b; }, 0);
    var mid = Math.floor(v.length / 2);
    return {
      n: v.length,
      mean: sum / v.length,
      median: v.length % 2 ? v[mid] : (v[mid - 1] + v[mid]) / 2,
      min: v[0], max: v[v.length - 1]
    };
  }

  var FLATS = Array.from(new Set(TICKETS.map(function (t) { return t.flat; })))
    .sort(function (a, b) { return num(a) - num(b); });
  var CATEGORIES = Array.from(new Set(TICKETS.map(function (t) { return t.category; }))).sort();

  var VILLA = {};
  FLATS.forEach(function (f) { VILLA[f] = { flat: f, tickets: [], open: 0, byCat: {} }; });
  TICKETS.forEach(function (t) {
    var v = VILLA[t.flat];
    v.tickets.push(t);
    if (t.open) v.open++;
  });
  FLATS.forEach(function (f) {
    var v = VILLA[f];
    v.all = stats(v.tickets);
    ESC_KEYS.forEach(function (k) {
      v[k] = stats(v.tickets.filter(function (t) { return t.esc === k; }));
    });
    CATEGORIES.forEach(function (c) {
      var s = stats(v.tickets.filter(function (t) { return t.category === c; }));
      if (s.n) v.byCat[c] = s;
    });
    v.tickets.sort(function (a, b) { return a.created < b.created ? 1 : -1; });
  });

  var SOC = { all: stats(TICKETS), byCat: {} };
  ESC_KEYS.forEach(function (k) {
    SOC[k] = stats(TICKETS.filter(function (t) { return t.esc === k; }));
  });
  CATEGORIES.forEach(function (c) {
    SOC.byCat[c] = stats(TICKETS.filter(function (t) { return t.category === c; }));
  });

  /* ranking: 1 = fastest average */
  var RANKED = FLATS.slice().sort(function (a, b) { return VILLA[a].all.mean - VILLA[b].all.mean; });
  RANKED.forEach(function (f, i) { VILLA[f].rank = i + 1; });
  var MEDIAN_VILLA = RANKED[Math.floor(RANKED.length / 2)];
  var FASTEST = RANKED[0], SLOWEST = RANKED[RANKED.length - 1];

  function standing(f) {
    var me = VILLA[f].all.mean, longer = 0, shorter = 0, same = 0;
    FLATS.forEach(function (o) {
      if (o === f) return;
      var m = VILLA[o].all.mean;
      if (m > me) longer++; else if (m < me) shorter++; else same++;
    });
    return {
      longer: longer, shorter: shorter, same: same,
      rank: VILLA[f].rank, total: FLATS.length,
      pctSlower: FLATS.length > 1 ? (longer / (FLATS.length - 1)) * 100 : 0
    };
  }

  /* ==================== Excel formula builders ==================== */
  function R(c) { return "$" + c + "$" + XL.first + ":$" + c + "$" + XL.last; }
  var MR = R(XL.mins);
  var AVGR = "$" + XL.avg + "$" + XL.first + ":$" + XL.avg + "$" + XL.uniqLast;
  var C = XL.col, N1 = XL.first;

  var F = {
    cutoff: "=DATE(2026,8,20)+TIME(9,30,0)",

    minutes:
      '=IF($' + C.closed + N1 + '="-",\n' +
      '  ROUND(($' + XL.cutoff + '-(DATE(2000+VALUE(MID($' + C.created + N1 + ',7,2)),VALUE(MID($' +
      C.created + N1 + ',4,2)),VALUE(LEFT($' + C.created + N1 + ',2)))+TIMEVALUE(MID($' +
      C.created + N1 + ',11,8))))*1440,0),\n' +
      '  VALUE(LEFT($' + C.rtat + N1 + ',FIND(" ",$' + C.rtat + N1 + ')-1))*1440\n' +
      '   +VALUE(MID($' + C.rtat + N1 + ',FIND("): ",$' + C.rtat + N1 + ')+3,FIND(" hour",$' +
      C.rtat + N1 + ')-FIND("): ",$' + C.rtat + N1 + ')-3))*60\n' +
      '   +VALUE(MID($' + C.rtat + N1 + ',FIND("hour(s): ",$' + C.rtat + N1 + ')+9,FIND(" minute",$' +
      C.rtat + N1 + ')-FIND("hour(s): ",$' + C.rtat + N1 + ')-9)))',

    villaEsc: function (f, k) {
      return "=AVERAGEIFS(" + MR + "," + R(C.flat) + ',"' + f + '",' + R(C.esc) + ',"' + k + '")';
    },
    villaEscCount: function (f, k) {
      return "=COUNTIFS(" + R(C.flat) + ',"' + f + '",' + R(C.esc) + ',"' + k + '")';
    },
    villaEscMedian: function (f, k) {
      return "=MEDIAN(IF((" + R(C.flat) + '="' + f + '")*(' + R(C.esc) + '="' + k + '"),' + MR + "))";
    },
    socEsc: function (k) { return "=AVERAGEIFS(" + MR + "," + R(C.esc) + ',"' + k + '")'; },
    villaAll: function (f) { return "=AVERAGEIFS(" + MR + "," + R(C.flat) + ',"' + f + '")'; },
    socAll: "=AVERAGE(" + MR + ")",
    uniqList: "=UNIQUE(" + R(C.flat) + ")",
    perVilla: "=AVERAGEIFS(" + MR + "," + R(C.flat) + ",$" + XL.uniq + N1 + ")",
    longer: function (f) {
      return "=COUNTIF(" + AVGR + ',">"&AVERAGEIFS(' + MR + "," + R(C.flat) + ',"' + f + '"))';
    },
    shorter: function (f) {
      return "=COUNTIF(" + AVGR + ',"<"&AVERAGEIFS(' + MR + "," + R(C.flat) + ',"' + f + '"))';
    },
    rank: function (f) {
      return "=RANK(AVERAGEIFS(" + MR + "," + R(C.flat) + ',"' + f + '"),' + AVGR + ",1)";
    },
    cat: function (f, c) {
      return "=AVERAGEIFS(" + MR + "," + R(C.flat) + ',"' + f + '",' + R(C.category) + ',"' + c + '")';
    },
    socCat: function (c) { return "=AVERAGEIFS(" + MR + "," + R(C.category) + ',"' + c + '")'; },
    count: function (f) { return "=COUNTIF(" + R(C.flat) + ',"' + f + '")'; },
    openCount: function (f) {
      return "=COUNTIFS(" + R(C.flat) + ',"' + f + '",' + R(C.closed) + ',"-")';
    },
    display: "=INT(" + XL.mins + N1 + '/1440)&"d "&INT(MOD(' + XL.mins + N1 + ',1440)/60)&"h"'
  };

  /* ==================== view helpers ==================== */
  var OPEN_FX = {};   // persists which formula drawers are open across re-render

  function formula(label, cell, code) {
    return '<div class="formula">' +
      '<div class="formula-lab">' + esc(label) +
      (cell ? '<span class="cell">' + esc(cell) + "</span>" : "") + "</div>" +
      "<pre>" + esc(code) + "</pre>" +
      '<button class="copy" type="button" data-copy="' + esc(code) + '">Copy</button>' +
      "</div>";
  }

  function fx(id, note, blocks) {
    var on = !!OPEN_FX[id];
    return '<div class="fx-bar">' +
      '<button class="fx-btn" type="button" data-fx="' + id + '" aria-expanded="' + on + '" aria-controls="' + id + '">' +
      '<span class="fx">fx</span><span>Show the Excel formula</span><span class="car">▶</span>' +
      "</button></div>" +
      '<div class="fx-body' + (on ? " on" : "") + '" id="' + id + '">' +
      (note ? '<p class="fx-note">' + note + "</p>" : "") + blocks + "</div>";
  }

  function chip(mine, base) {
    if (mine == null || base == null || !base) return "";
    var d = ((mine - base) / base) * 100;
    var r = Math.round(Math.abs(d));
    if (r < 1) return '<span class="chip flat"><span class="g">=</span>Same as society average</span>';
    var slower = d > 0;
    return '<span class="chip ' + (slower ? "bad" : "good") + '">' +
      '<span class="g">' + (slower ? "▲" : "▼") + "</span>" +
      r + "% " + (slower ? "slower" : "faster") + " than society</span>";
  }

  function bars(rows) {
    var max = Math.max.apply(null, rows.map(function (r) { return r.v || 0; })) || 1;
    return '<div class="cmp">' + rows.map(function (r) {
      return '<div class="cmp-row"><span class="lbl">' + esc(r.label) + "</span>" +
        '<span class="track"><span class="fill ' + r.cls + '" style="width:' +
        (r.v == null ? 0 : Math.max(1.5, (r.v / max) * 100)) + '%"></span></span>' +
        '<span class="val">' + dhm(r.v) + "</span></div>";
    }).join("") + "</div>";
  }

  /* ==================== state ==================== */
  var S = {
    flat: null,
    peer: null,
    sortKey: "rank",
    sortDir: 1,
    q: ""
  };

  /* ==================== sections ==================== */

  function secIdentity(f) {
    var v = VILLA[f];
    var first = v.tickets[v.tickets.length - 1], last = v.tickets[0];
    return '<div class="ident">' +
      '<div class="ident-name"><span class="vn">' + esc(f) + "</span>" +
      '<span class="vs">' + esc(META.society) + " · tickets from " + dt(first.created) + " to " + dt(last.created) + "</span></div>" +
      '<div class="minis">' +
      '<div class="mini"><b>' + v.all.n + "</b><span>Tickets</span></div>" +
      '<div class="mini"><b>' + (v.all.n - v.open) + "</b><span>Closed</span></div>" +
      '<div class="mini"><b>' + v.open + "</b><span>Still open</span></div>" +
      '<div class="mini"><b>' + dhm(v.all.mean) + "</b><span>Avg resolution</span></div>" +
      '<div class="mini"><b>#' + v.rank + "</b><span>of " + FLATS.length + " villas</span></div>" +
      "</div></div>";
  }

  function escCard(f, k) {
    var v = VILLA[f], s = v[k], soc = SOC[k], label = ESC[k];
    var id = "fx-esc-" + k.replace(/\W/g, "");
    var body;

    if (!s.n) {
      body = '<div class="empty"><b>No ' + label.toLowerCase() + " tickets for " + esc(f) + ".</b><br>" +
        "Across the whole society there " + (soc.n === 1 ? "is " : "are ") + soc.n + " " +
        label.toLowerCase() + " ticket" + (soc.n === 1 ? "" : "s") +
        ", averaging <b>" + dhm(soc.mean) + "</b>.</div>";
    } else {
      body =
        '<div class="hero"><span class="v">' + days(s.mean) + '</span><span class="u">days average · ' +
        dhm(s.mean) + "</span></div>" +
        chip(s.mean, soc.mean) +
        '<div class="card-sub"><span>Median <b>' + dhm(s.median) + "</b></span>" +
        "<span>Fastest <b>" + dhm(s.min) + "</b></span>" +
        "<span>Slowest <b>" + dhm(s.max) + "</b></span>" +
        "<span>Tickets <b>" + s.n + "</b></span></div>" +
        bars([
          { label: esc(f), v: s.mean, cls: "me" },
          { label: "Society", v: soc.mean, cls: "soc" }
        ]);
    }

    return '<div class="card"><div class="card-top">' +
      '<div class="card-tag"><span class="dot ' + (k === "Level-1" ? "urgent" : "usual") + '"></span>' +
      "<b>" + label + "</b><code>Escalation Level = " + esc(k) + "</code></div>" +
      body + "</div>" +
      fx(id + "-" + f,
        "Add the <code>" + XL.mins + "</code> helper column first (see <b>How every number is built</b> at the foot of the page), then:",
        formula("Average, " + label.toLowerCase() + ", " + f, null, F.villaEsc(f, k)) +
        formula("Ticket count", null, F.villaEscCount(f, k)) +
        formula("Median (array formula — Ctrl+Shift+Enter on older Excel)", null, F.villaEscMedian(f, k)) +
        formula("Society average for comparison", null, F.socEsc(k))) +
      "</div>";
  }

  function secEscalation(f) {
    return '<div class="panel"><div class="panel-h"><div>' +
      '<h2>Average resolution time by escalation level</h2>' +
      "<p>How long tickets from " + esc(f) + " take to close, split by how the help desk classified them. " +
      "Lower is better.</p></div></div>" +
      '<div class="panel-b"><div class="cards">' +
      ESC_KEYS.map(function (k) { return escCard(f, k); }).join("") +
      "</div></div></div>";
  }

  /* ---- distribution of all villa averages, with this villa marked ---- */
  function histogram(f) {
    var W = 900, H = 210, PL = 44, PR = 16, PT = 14, PB = 34;
    var BIN = 2, MAXD = 30;                       // 2-day bins to 30d, then overflow
    var nb = MAXD / BIN + 1;
    var binsArr = new Array(nb).fill(0);
    var mineBin = 0;
    FLATS.forEach(function (o) {
      var d = VILLA[o].all.mean / 1440;
      var i = d >= MAXD ? nb - 1 : Math.floor(d / BIN);
      binsArr[i]++;
      if (o === f) mineBin = i;
    });
    var maxC = Math.max.apply(null, binsArr);
    var iw = (W - PL - PR) / nb, ih = H - PT - PB;

    var g = "";
    for (var t = 0; t <= 2; t++) {
      var y = PT + ih - (ih * t) / 2, val = Math.round((maxC * t) / 2);
      g += '<line class="grid-l" x1="' + PL + '" y1="' + y + '" x2="' + (W - PR) + '" y2="' + y + '"/>' +
        '<text class="axis-l" x="' + (PL - 8) + '" y="' + (y + 3.5) + '" text-anchor="end">' + val + "</text>";
    }

    var b = binsArr.map(function (c, i) {
      var h = maxC ? (c / maxC) * ih : 0;
      var x = PL + i * iw + 2, y = PT + ih - h;
      var mine = i === mineBin;
      return '<rect x="' + x + '" y="' + y + '" width="' + (iw - 4) + '" height="' + Math.max(h, c ? 2 : 0) +
        '" rx="3" fill="' + (mine ? "var(--accent)" : "var(--neutral-mark)") +
        '" opacity="' + (mine ? 1 : 0.55) + '"><title>' +
        (i === nb - 1 ? "30+ days" : i * BIN + "–" + (i + 1) * BIN + " days") +
        ": " + c + " villa" + (c === 1 ? "" : "s") + "</title></rect>";
    }).join("");

    var ax = "";
    for (var i = 0; i < nb; i += 2) {
      ax += '<text class="axis-l" x="' + (PL + i * iw + iw / 2) + '" y="' + (H - PB + 16) +
        '" text-anchor="middle">' + (i === nb - 1 ? "30+" : i * BIN) + "</text>";
    }
    ax += '<text class="axis-l" x="' + (PL + (W - PL - PR) / 2) + '" y="' + (H - 3) +
      '" text-anchor="middle">Average resolution time (days)</text>';

    var mx = PL + mineBin * iw + iw / 2;
    var mark = '<line x1="' + mx + '" y1="' + (PT - 6) + '" x2="' + mx + '" y2="' + (PT + ih) +
      '" stroke="var(--accent)" stroke-width="2" stroke-dasharray="3 3"/>' +
      '<text class="mark-lab" x="' + (mx + (mineBin > nb - 4 ? -7 : 7)) + '" y="' + (PT - 1) +
      '" text-anchor="' + (mineBin > nb - 4 ? "end" : "start") + '">' + esc(f) + "</text>";

    return '<div class="legend"><span><i style="background:var(--accent)"></i>' + esc(f) + " (your villa)</span>" +
      '<span><i style="background:var(--neutral-mark);opacity:.55"></i>Other villas</span></div>' +
      '<div class="chart-wrap"><svg class="chart" viewBox="0 0 ' + W + " " + H +
      '" role="img" aria-label="Distribution of average resolution time across all villas">' +
      g + b + mark + ax + "</svg></div>";
  }

  function secStanding(f) {
    var st = standing(f), v = VILLA[f];
    return '<div class="panel"><div class="panel-h"><div>' +
      "<h2>How " + esc(f) + " compares with every other villa</h2>" +
      "<p>All " + FLATS.length + " villas ranked by their own average resolution time across every ticket they raised.</p>" +
      "</div></div>" +
      '<div class="panel-b">' +
      '<div class="cards" style="margin-bottom:22px">' +
      '<div class="card"><div class="card-top">' +
      '<div class="card-tag"><span class="dot usual"></span><b>Villas slower than yours</b></div>' +
      '<div class="hero"><span class="v">' + st.longer + '</span><span class="u">of ' + (st.total - 1) +
      " other villas take longer</span></div>" +
      '<div class="card-sub"><span>That puts you ahead of <b>' + Math.round(st.pctSlower) + "%</b> of villas</span></div>" +
      "</div></div>" +
      '<div class="card"><div class="card-top">' +
      '<div class="card-tag"><span class="dot urgent"></span><b>Villas faster than yours</b></div>' +
      '<div class="hero"><span class="v">' + st.shorter + '</span><span class="u">of ' + (st.total - 1) +
      " other villas resolve quicker</span></div>" +
      '<div class="card-sub"><span>Overall rank <b>#' + st.rank + "</b> of " + st.total +
      " · average <b>" + dhm(v.all.mean) + "</b></span></div>" +
      "</div></div>" +
      "</div>" +
      histogram(f) +
      "</div>" +
      fx("fx-standing-" + f,
        "First build a per-villa average table: put <code>" + F.uniqList + "</code> in <code>" + XL.uniq + XL.first +
        "</code> to list every villa, and <code>" + esc(F.perVilla) + "</code> in <code>" + XL.avg + XL.first +
        "</code> filled down to row " + XL.uniqLast + ". Then:",
        formula("Villas that take LONGER than " + f, null, F.longer(f)) +
        formula("Villas that are FASTER than " + f, null, F.shorter(f)) +
        formula("Rank of " + f + " (1 = fastest)", null, F.rank(f)) +
        formula("This villa's own average", null, F.villaAll(f))) +
      "</div>";
  }

  /* ---- head to head against any other villa ---- */
  function secHeadToHead(f) {
    var peer = S.peer && S.peer !== f ? S.peer : (RANKED[0] === f ? RANKED[1] : RANKED[0]);
    var a = VILLA[f], b = VILLA[peer];

    function row(label, av, bv, lower) {
      var better = av == null || bv == null ? "" :
        (av === bv ? "" : (av < bv) === !!lower ? "a" : "b");
      return "<tr><td>" + esc(label) + "</td>" +
        '<td class="n"' + (better === "a" ? ' style="color:var(--good);font-weight:600"' : "") + ">" +
        (typeof av === "number" && av > 100 ? dhm(av) : av == null ? "—" : av) + "</td>" +
        '<td class="n"' + (better === "b" ? ' style="color:var(--good);font-weight:600"' : "") + ">" +
        (typeof bv === "number" && bv > 100 ? dhm(bv) : bv == null ? "—" : bv) + "</td></tr>";
    }

    var opts = FLATS.filter(function (x) { return x !== f; }).map(function (x) {
      return '<option value="' + esc(x) + '"' + (x === peer ? " selected" : "") + ">" +
        esc(x) + " — " + dhm(VILLA[x].all.mean) + " avg</option>";
    }).join("");

    return '<div class="panel"><div class="panel-h"><div>' +
      "<h2>Head to head</h2>" +
      "<p>Compare " + esc(f) + " directly against any other villa in the society.</p></div>" +
      '<div style="margin-left:auto"><label class="eyebrow" for="peerSel" style="display:block;margin-bottom:4px">Compare with</label>' +
      '<select id="peerSel" style="border:1.5px solid var(--line);border-radius:8px;background:var(--panel-2);' +
      'color:var(--ink);padding:7px 10px;font:inherit;font-size:13px;font-family:\'IBM Plex Mono\',monospace">' +
      opts + "</select></div></div>" +
      '<div class="panel-b"><div class="tbl-scroll" style="max-height:none"><table>' +
      "<thead><tr><th>Measure</th><th>" + esc(f) + "</th><th>" + esc(peer) + "</th></tr></thead><tbody>" +
      row("Average resolution", a.all.mean, b.all.mean, true) +
      row("Median resolution", a.all.median, b.all.median, true) +
      row("Urgent average", a["Level-1"].mean, b["Level-1"].mean, true) +
      row("Usual average", a["NA"].mean, b["NA"].mean, true) +
      row("Fastest ticket", a.all.min, b.all.min, true) +
      row("Slowest ticket", a.all.max, b.all.max, true) +
      row("Tickets raised", a.all.n, b.all.n, false) +
      row("Still open", a.open, b.open, true) +
      row("Rank (1 = fastest)", a.rank, b.rank, true) +
      "</tbody></table></div></div>" +
      fx("fx-h2h-" + f + "-" + peer,
        "Swap the villa name in the criteria to compare any two villas.",
        formula("Average for " + f, null, F.villaAll(f)) +
        formula("Average for " + peer, null, F.villaAll(peer)) +
        formula("Urgent average for " + f, null, F.villaEsc(f, "Level-1")) +
        formula("Tickets raised by " + f, null, F.count(f)) +
        formula("Still-open tickets for " + f, null, F.openCount(f))) +
      "</div>";
  }

  /* ---- category breakdown, villa vs society ---- */
  function secCategories(f) {
    var v = VILLA[f];
    var cats = Object.keys(v.byCat).sort(function (a, b) { return v.byCat[b].mean - v.byCat[a].mean; });
    if (!cats.length) return "";
    var max = Math.max.apply(null, cats.map(function (c) {
      return Math.max(v.byCat[c].mean, SOC.byCat[c].mean);
    }));

    var rows = cats.map(function (c) {
      var mine = v.byCat[c], soc = SOC.byCat[c];
      return '<div style="margin-bottom:16px">' +
        '<div style="display:flex;justify-content:space-between;align-items:baseline;margin-bottom:6px;gap:12px;flex-wrap:wrap">' +
        '<b style="font-size:13px">' + esc(catLabel(c)) + '</b>' +
        '<span style="font-size:11.5px;color:var(--ink-3)">' + mine.n + " ticket" + (mine.n === 1 ? "" : "s") +
        " · " + esc(f) + " <b class=\"num\" style=\"color:var(--ink)\">" + dhm(mine.mean) +
        "</b> vs society <b class=\"num\" style=\"color:var(--ink)\">" + dhm(soc.mean) + "</b></span></div>" +
        bars([
          { label: esc(f), v: mine.mean, cls: "me" },
          { label: "Society", v: soc.mean, cls: "soc" }
        ]).replace('class="cmp"', 'class="cmp" style="gap:5px"') +
        "</div>";
    }).join("");

    return '<div class="panel"><div class="panel-h"><div>' +
      "<h2>Which kinds of problems take longest</h2>" +
      "<p>" + esc(f) + "&rsquo;s average resolution time per category, against the society average for the same category.</p>" +
      "</div></div><div class=\"panel-b\">" + rows + "</div>" +
      fx("fx-cat-" + f,
        "One pair of formulas per category. " + (cats.some(hasPadding)
          ? "Note that <b>" + cats.filter(hasPadding).map(function (c) { return esc(catLabel(c)); }).join(", ") +
            "</b> is stored in the sheet with a trailing space, so the criteria below keep it &mdash; " +
            "retyping the name without it makes Excel match nothing."
          : "The criteria match the category text exactly as the sheet stores it."),
        cats.map(function (c) {
          return formula(catLabel(c) + " — " + f, null, F.cat(f, c)) +
            formula(catLabel(c) + " — society", null, F.socCat(c));
        }).join("")) +
      "</div>";
  }

  /* ---- every villa, sortable ---- */
  var COLS = [
    { k: "rank", t: "#", n: true },
    { k: "flat", t: "Villa", n: false },
    { k: "n", t: "Tickets", n: true },
    { k: "open", t: "Open", n: true },
    { k: "urgent", t: "Urgent avg", n: true },
    { k: "usual", t: "Usual avg", n: true },
    { k: "mean", t: "Overall avg", n: true },
    { k: "median", t: "Median", n: true }
  ];

  function rowVals(f) {
    var v = VILLA[f];
    return {
      rank: v.rank, flat: f, n: v.all.n, open: v.open,
      urgent: v["Level-1"].mean, usual: v["NA"].mean,
      mean: v.all.mean, median: v.all.median
    };
  }

  function secTable(f) {
    var q = S.q.trim().toLowerCase();
    var list = FLATS.map(rowVals).filter(function (r) {
      return !q || r.flat.toLowerCase().indexOf(q) >= 0;
    });
    list.sort(function (a, b) {
      var x = a[S.sortKey], y = b[S.sortKey];
      if (x == null) return 1;
      if (y == null) return -1;
      if (S.sortKey === "flat") return (num(x) - num(y)) * S.sortDir;
      return (x - y) * S.sortDir;
    });

    var head = COLS.map(function (c) {
      var on = S.sortKey === c.k;
      return '<th class="s' + (on ? " on" : "") + '" data-sort="' + c.k + '" scope="col">' + esc(c.t) +
        '<span class="ar">' + (on ? (S.sortDir > 0 ? "▲" : "▼") : "▲") + "</span></th>";
    }).join("");

    var body = list.map(function (r) {
      return '<tr class="' + (r.flat === f ? "me" : "") + '" data-goto="' + esc(r.flat) + '" style="cursor:pointer">' +
        '<td class="n">' + r.rank + "</td>" +
        '<td class="n">' + esc(r.flat) + "</td>" +
        '<td class="n">' + r.n + "</td>" +
        '<td class="n">' + (r.open || "—") + "</td>" +
        '<td class="n">' + dhm(r.urgent) + "</td>" +
        '<td class="n">' + dhm(r.usual) + "</td>" +
        '<td class="n">' + dhm(r.mean) + "</td>" +
        '<td class="n">' + dhm(r.median) + "</td></tr>";
    }).join("");

    return '<div class="panel"><div class="panel-h"><div>' +
      "<h2>Every villa, ranked</h2>" +
      "<p>Sort by any column, or search for a neighbour. Click a row to open that villa. " +
      "Society average is <b>" + dhm(SOC.all.mean) + "</b>; the median villa (" + esc(MEDIAN_VILLA) +
      ") averages <b>" + dhm(VILLA[MEDIAN_VILLA].all.mean) + "</b>.</p></div></div>" +
      '<div class="panel-b">' +
      '<div class="tbl-tools"><input id="tblQ" type="search" placeholder="Search villa…" value="' + esc(S.q) + '">' +
      '<span class="tbl-note">' + list.length + " of " + FLATS.length + " villas · fastest " +
      esc(FASTEST) + " (" + dhm(VILLA[FASTEST].all.mean) + ") · slowest " +
      esc(SLOWEST) + " (" + dhm(VILLA[SLOWEST].all.mean) + ")</span></div>" +
      '<div class="tbl-scroll"><table><thead><tr>' + head + "</tr></thead><tbody>" + body + "</tbody></table></div>" +
      "</div>" +
      fx("fx-table",
        "Build the whole league table in two columns, then sort it:",
        formula("List every villa once", XL.uniq + XL.first, F.uniqList) +
        formula("Average for that villa (fill down to row " + XL.uniqLast + ")", XL.avg + XL.first, F.perVilla) +
        formula("Rank it", "AC" + XL.first, "=RANK($" + XL.avg + XL.first + "," + AVGR + ",1)") +
        formula("Society average, all tickets", null, F.socAll)) +
      "</div>";
  }

  /* ---- ticket ledger ---- */
  function secLedger(f) {
    var v = VILLA[f];
    var body = v.tickets.map(function (t) {
      return "<tr><td class=\"n\">" + t.id + "</td>" +
        "<td>" + dt(t.created) + "</td>" +
        "<td>" + (t.open ? '<span class="pill open">Open</span>' : dt(t.closed)) + "</td>" +
        "<td>" + esc(catLabel(t.category)) + "</td>" +
        "<td>" + esc(t.subject) + "</td>" +
        '<td><span class="pill ' + (t.esc === "Level-1" ? "urgent" : "usual") + '">' + ESC[t.esc] + "</span></td>" +
        '<td class="n">' + dhm(t.mins) + "</td>" +
        '<td class="n">' + n0(t.mins) + "</td></tr>";
    }).join("");

    return '<div class="panel"><div class="panel-h"><div>' +
      "<h2>Every ticket from " + esc(f) + "</h2>" +
      "<p>The rows behind the averages above. <b>Minutes</b> is the value the helper column produces for that ticket &mdash; " +
      "check any one of them by hand.</p></div></div>" +
      '<div class="panel-b"><div class="tbl-scroll"><table><thead><tr>' +
      "<th>Id</th><th>Raised</th><th>Closed</th><th>Category</th><th>Subject</th>" +
      "<th>Escalation</th><th>Time taken</th><th>Minutes</th></tr></thead><tbody>" +
      body + "</tbody></table></div></div>" +
      fx("fx-ledger-" + f,
        "Every ticket for this villa, straight from the sheet — filter column <code>" + C.flat +
        "</code> to <code>" + esc(f) + "</code>. The Minutes column is:",
        formula("Minutes helper", XL.mins + XL.first, F.minutes) +
        formula("Readable form", null, F.display)) +
      "</div>";
  }

  /* ---- methodology ---- */
  function secMethod() {
    return '<div class="panel"><div class="panel-h"><div>' +
      "<h2>How every number on this page is built</h2>" +
      "<p>Open <code>" + esc(META.source) + "</code>, add the two cells below, and every figure here can be " +
      "reproduced with a standard Excel formula. Nothing is pre-computed anywhere else.</p></div></div>" +
      '<div class="panel-b">' +
      '<div class="meth">' +
      "<section><h3>1. The sheet as exported</h3><p>Row 1 is the report title and row 2 is blank. " +
      "Column headers sit on <b>row 3</b>, and the " + META.total + " tickets occupy <b>rows " +
      XL.first + " to " + XL.last + "</b>. Villa is column <code>" + C.flat + "</code>, escalation level is <code>" +
      C.esc + "</code>, resolution time is <code>" + C.rtat + "</code>.</p></section>" +

      "<section><h3>2. Turning the time text into a number</h3><p><code>" + C.rtat + "</code> holds text such as " +
      "<code>2 day(s): 3 hour(s): 5 minute(s)</code>, which Excel cannot average. The helper column converts it " +
      "to plain minutes.</p></section>" +

      "<section><h3>3. Tickets that are still open</h3><p>" + META.openCount + " tickets have never been closed " +
      "(<code>" + C.closed + "</code> is <code>-</code>). Their exported resolution time is stale, so they are " +
      "measured from when they were raised up to a fixed cutoff of <b>20 August 2026, 9:30 AM</b>. " +
      "They are included in every average on this page.</p></section>" +

      "<section><h3>4. Escalation levels</h3><p>The export uses two values. <code>Level-1</code> is shown here as " +
      "<b>Urgent</b> and <code>NA</code> as <b>Usual</b>. No other value appears in the data.</p></section>" +

      "<section><h3>5. Columns that were dropped</h3><p>These held one identical value on all " + META.total +
      " rows, so they carry no information and are not shown.</p>" +
      '<ul class="drop">' + META.droppedColumns.map(function (c) { return "<li>" + esc(c) + "</li>"; }).join("") +
      "</ul></section>" +

      "<section><h3>6. Averages</h3><p>Every average is a plain arithmetic mean of the helper column over the " +
      "matching rows &mdash; <code>AVERAGEIFS</code>, no weighting and no outlier removal. Medians are shown " +
      "alongside because a few very old tickets pull the mean up.</p></section>" +
      "</div></div>" +
      fx("fx-method",
        "Add these two cells to the sheet, then fill <code>" + XL.mins + XL.first + "</code> down to row " +
        XL.last + ". Every other formula on this page depends only on these.",
        formula("Cutoff for open tickets", "Z1", F.cutoff) +
        formula("Minutes — fill down " + XL.mins + XL.first + ":" + XL.mins + XL.last, XL.mins + XL.first, F.minutes) +
        formula("Excel 365 alternative (same result, shorter)", XL.mins + XL.first,
          '=IF($' + C.closed + N1 + '="-",\n' +
          '  ROUND(($' + XL.cutoff + '-(DATE(2000+VALUE(MID($' + C.created + N1 + ',7,2)),VALUE(MID($' +
          C.created + N1 + ',4,2)),VALUE(LEFT($' + C.created + N1 + ',2)))+TIMEVALUE(MID($' +
          C.created + N1 + ',11,8))))*1440,0),\n' +
          '  VALUE(TEXTBEFORE($' + C.rtat + N1 + '," "))*1440\n' +
          '   +VALUE(TEXTBEFORE(TEXTAFTER($' + C.rtat + N1 + ',": ")," "))*60\n' +
          '   +VALUE(TEXTBEFORE(TEXTAFTER($' + C.rtat + N1 + ',": ",2)," ")))')) +
      "</div>";
  }

  function blank() {
    return '<div class="panel"><div class="blank"><b>Search for your villa to begin</b>' +
      "Type a villa number above &mdash; for example <b>75</b> or <b>Villa-75</b>.<br>" +
      "This report covers " + META.total + " tickets from " + FLATS.length + " villas, raised between " +
      dt(META.rangeStart) + " and " + dt(META.rangeEnd) + ".</div></div>" + secTable(null) + secMethod();
  }

  /* ==================== render ==================== */
  var app = document.getElementById("app");
  var footEl = document.getElementById("foot");

  function render() {
    var f = S.flat;
    app.innerHTML = f
      ? secIdentity(f) + secEscalation(f) + secStanding(f) + secHeadToHead(f) +
        secCategories(f) + secTable(f) + secLedger(f) + secMethod()
      : blank();

    footEl.innerHTML = esc(META.society) + " help desk export &mdash; " + META.total + " tickets, " +
      FLATS.length + " villas, " + dt(META.rangeStart) + " to " + dt(META.rangeEnd) + ". " +
      "Open tickets are measured to a fixed cutoff of 20 Aug 2026, 9:30 AM. " +
      "Source file: " + esc(META.source) + ". Built " + dt(META.generated) + ".";
  }

  /* ==================== interactions ==================== */
  var input = document.getElementById("villaInput");
  var sugg = document.getElementById("sugg");
  var cur = -1, matches = [];

  function match(q) {
    q = q.trim().toLowerCase().replace(/^villa[-\s]*/, "");
    if (!q) return FLATS.slice(0, 30);
    return FLATS.filter(function (f) {
      return f.toLowerCase().indexOf(q) >= 0 || String(num(f)).indexOf(q) === 0;
    }).slice(0, 40);
  }

  function drawSugg() {
    if (!matches.length) {
      sugg.innerHTML = '<div class="none">No villa matches that.</div>';
    } else {
      sugg.innerHTML = matches.map(function (f, i) {
        var v = VILLA[f];
        return '<button type="button" role="option" aria-selected="' + (i === cur) + '" class="' +
          (i === cur ? "cur" : "") + '" data-pick="' + esc(f) + '">' +
          '<span class="sv">' + esc(f) + "</span>" +
          '<span class="sm">' + v.all.n + " tickets · " + dhm(v.all.mean) + " avg</span></button>";
      }).join("");
    }
    sugg.classList.add("on");
    input.setAttribute("aria-expanded", "true");
  }
  function closeSugg() {
    sugg.classList.remove("on");
    input.setAttribute("aria-expanded", "false");
    cur = -1;
  }
  function pick(f) {
    S.flat = f;
    S.peer = null;
    input.value = f;
    closeSugg();
    render();
    window.scrollTo({ top: 0, behavior: "smooth" });
    if (history.replaceState) history.replaceState(null, "", "#" + encodeURIComponent(f));
  }

  input.addEventListener("input", function () { matches = match(input.value); cur = -1; drawSugg(); });
  input.addEventListener("focus", function () { matches = match(input.value); drawSugg(); });
  input.addEventListener("keydown", function (e) {
    if (e.key === "ArrowDown" || e.key === "ArrowUp") {
      e.preventDefault();
      if (!sugg.classList.contains("on")) { matches = match(input.value); drawSugg(); return; }
      cur += e.key === "ArrowDown" ? 1 : -1;
      if (cur < 0) cur = matches.length - 1;
      if (cur >= matches.length) cur = 0;
      drawSugg();
      var el = sugg.children[cur];
      if (el && el.scrollIntoView) el.scrollIntoView({ block: "nearest" });
    } else if (e.key === "Enter") {
      e.preventDefault();
      if (matches.length) pick(matches[cur >= 0 ? cur : 0]);
    } else if (e.key === "Escape") { closeSugg(); }
  });
  sugg.addEventListener("mousedown", function (e) {
    var b = e.target.closest("[data-pick]");
    if (b) { e.preventDefault(); pick(b.getAttribute("data-pick")); }
  });
  document.addEventListener("click", function (e) {
    if (!e.target.closest(".lookup")) closeSugg();
  });

  /* delegated: formula drawers, copy, sorting, row clicks, peer select */
  document.addEventListener("click", function (e) {
    var fxb = e.target.closest("[data-fx]");
    if (fxb) {
      var id = fxb.getAttribute("data-fx");
      var body = document.getElementById(id);
      var on = !body.classList.contains("on");
      body.classList.toggle("on", on);
      fxb.setAttribute("aria-expanded", String(on));
      if (on) OPEN_FX[id] = true; else delete OPEN_FX[id];
      return;
    }
    var cp = e.target.closest("[data-copy]");
    if (cp) {
      var txt = cp.getAttribute("data-copy");
      var done = function () {
        cp.textContent = "Copied";
        cp.classList.add("done");
        setTimeout(function () { cp.textContent = "Copy"; cp.classList.remove("done"); }, 1600);
      };
      if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(txt).then(done, function () { fallbackCopy(txt, done); });
      } else { fallbackCopy(txt, done); }
      return;
    }
    var th = e.target.closest("[data-sort]");
    if (th) {
      var k = th.getAttribute("data-sort");
      if (S.sortKey === k) S.sortDir *= -1; else { S.sortKey = k; S.sortDir = 1; }
      render();
      return;
    }
    var tr = e.target.closest("[data-goto]");
    if (tr) { pick(tr.getAttribute("data-goto")); }
  });

  function fallbackCopy(txt, done) {
    var ta = document.createElement("textarea");
    ta.value = txt;
    ta.style.cssText = "position:fixed;opacity:0";
    document.body.appendChild(ta);
    ta.select();
    try { document.execCommand("copy"); done(); } catch (err) { /* clipboard unavailable */ }
    document.body.removeChild(ta);
  }

  document.addEventListener("change", function (e) {
    if (e.target.id === "peerSel") { S.peer = e.target.value; render(); }
  });
  document.addEventListener("input", function (e) {
    if (e.target.id === "tblQ") {
      S.q = e.target.value;
      render();
      var el = document.getElementById("tblQ");
      if (el) { el.focus(); el.setSelectionRange(el.value.length, el.value.length); }
    }
  });

  /* ==================== boot ==================== */
  document.getElementById("socName").textContent = META.society;
  var hash = decodeURIComponent((location.hash || "").replace(/^#/, ""));
  if (hash && VILLA[hash]) { S.flat = hash; input.value = hash; }
  render();
})();
