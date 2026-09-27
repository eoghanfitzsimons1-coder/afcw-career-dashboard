/* ============================================================
   Career Mode Dashboard — file parsers
   ============================================================ */
(function(){
  "use strict";

  const uid = () => Math.random().toString(36).slice(2, 10);

  function strip(v){
    if(v === null || v === undefined) return "";
    return String(v).trim();
  }
  function asInt(v, def){
    const n = parseFloat(v);
    return isNaN(n) ? (def === undefined ? 0 : def) : Math.trunc(n);
  }

  const LEAGUE_SHEETS = ["Premier League", "EFL Championship", "EFL League One", "EFL League Two"];
  const SHEETS = {
    dashboard: "Dashboard",
    squad:     "Squad Data",
    youth:     "Youth Academy",
    finances:  "Club Financials",
    cups:      "Cup Results",
    summary:   "Career Summary"
  };

  function isExcelReady(){
    return typeof window.XLSX !== "undefined" && window.XLSX && window.XLSX.read;
  }

  function sheetToRows(wb, name){
    if(!wb || !Array.isArray(wb.SheetNames)) return null;
    if(!wb.SheetNames.includes(name)) return null;
    const ws = wb.Sheets[name];
    if(!ws) return null;
    return window.XLSX.utils.sheet_to_json(ws, {
      header: 1, defval: "", blankrows: false, raw: false
    });
  }

  function findHeader(rows, required, maxScan){
    return findHeaderFrom(rows, required, 0, maxScan);
  }

  function findHeaderFrom(rows, required, startRow, maxScan){
    const start = Math.max(0, startRow || 0);
    const limit = Math.min(start + (maxScan || 60), rows.length);
    for(let i = start; i < limit; i++){
      const cells = (rows[i] || []).map(c => strip(c).toLowerCase());
      if(required.every(r => cells.includes(r.toLowerCase()))) return i;
    }
    return -1;
  }

  function colIndex(rows, headerRow, name){
    const row = rows[headerRow] || [];
    const target = name.toLowerCase();
    for(let j = 0; j < row.length; j++){
      if(strip(row[j]).toLowerCase() === target) return j;
    }
    return -1;
  }

  function colIndexAny(rows, headerRow, names){
    for(const n of names){
      const i = colIndex(rows, headerRow, n);
      if(i >= 0) return i;
    }
    return -1;
  }

  function cellAt(rows, rowIdx, colIdx){
    if(colIdx < 0) return "";
    const row = rows[rowIdx] || [];
    return strip(row[colIdx]);
  }

  /* ---------- League table (from a given row) ---------- */
  function parseLeagueFrom(rows, startRow){
    const hdr = findHeaderFrom(rows, ["Pos","Team","PTS"], startRow);
    if(hdr < 0) return { rows: [], nextStart: startRow };

    const idx = {};
    for(const n of ["Pos","Team","P","W","D","L","GF","GA","GD","PTS"]){
      idx[n] = colIndex(rows, hdr, n);
    }
    if(idx.Pos < 0 || idx.Team < 0 || idx.PTS < 0) return { rows: [], nextStart: startRow };

    const out = [];
    let i = hdr + 1;
    for(; i < rows.length; i++){
      const team = cellAt(rows, i, idx.Team);
      if(!team || team.toLowerCase() === "nan") break;
      const pos = asInt(cellAt(rows, i, idx.Pos), -1);
      if(pos < 0) break;
      out.push({
        pos,
        team,
        p:   asInt(cellAt(rows, i, idx.P)),
        w:   asInt(cellAt(rows, i, idx.W)),
        d:   asInt(cellAt(rows, i, idx.D)),
        l:   asInt(cellAt(rows, i, idx.L)),
        gf:  asInt(cellAt(rows, i, idx.GF)),
        ga:  asInt(cellAt(rows, i, idx.GA)),
        gd:  asInt(cellAt(rows, i, idx.GD)),
        pts: asInt(cellAt(rows, i, idx.PTS))
      });
    }
    return { rows: out, nextStart: i };
  }

  function parseLeague(rows){
    const r = parseLeagueFrom(rows, 0);
    return r.rows.map(x => [x.pos, x.team, x.p, x.w, x.d, x.l, x.gf, x.ga, x.gd, x.pts]);
  }

  /* ---------- Champions League (group + knockout) ----------
     Tolerates blank rows inside the group table and any spacing
     between the group table and the knockout table. Accepts a
     header with Pos|Team|PTS, or just Team|PTS (first column read
     as position). Group letter picked up from any nearby cell.
  */
  function parseChampionsLeague(rows){
    if(!rows || !rows.length) return null;

    const out = { group: "", table: [], knockout: [] };

    let hdr = findHeaderFrom(rows, ["Pos","Team","PTS"], 0);
    let posFromFirstCol = false;
    if(hdr < 0){
      hdr = findHeaderFrom(rows, ["Team","PTS"], 0);
      if(hdr >= 0) posFromFirstCol = true;
    }
    if(hdr < 0) return null;

    for(let r = Math.max(0, hdr - 2); r <= Math.min(rows.length - 1, hdr + 1); r++){
      const t = (rows[r] || []).map(c => strip(c)).join(" ");
      const m = t.match(/Group\s+([A-Z0-9]+)/i);
      if(m){ out.group = m[1].toUpperCase(); break; }
    }

    const iPos  = posFromFirstCol ? 0 : colIndex(rows, hdr, "Pos");
    const iTeam = colIndex(rows, hdr, "Team");
    const iP    = colIndex(rows, hdr, "P");
    const iW    = colIndex(rows, hdr, "W");
    const iD    = colIndex(rows, hdr, "D");
    const iL    = colIndex(rows, hdr, "L");
    const iGF   = colIndex(rows, hdr, "GF");
    const iGA   = colIndex(rows, hdr, "GA");
    const iGD   = colIndex(rows, hdr, "GD");
    const iPts  = colIndex(rows, hdr, "PTS");

    let i = hdr + 1;
    let blankRun = 0;
    for(; i < rows.length; i++){
      const row = rows[i] || [];
      const allBlank = row.every(c => !strip(c));
      if(allBlank){
        blankRun++;
        if(blankRun >= 2) break;
        continue;
      }
      blankRun = 0;

      const team = cellAt(rows, i, iTeam);
      if(!team || team.toLowerCase() === "nan") break;

      let pos = i - hdr;
      if(posFromFirstCol){
        const n = parseInt(cellAt(rows, i, 0), 10);
        if(isNaN(n)) continue;
        pos = n;
      } else {
        const n = parseInt(cellAt(rows, i, iPos), 10);
        if(isNaN(n)) continue;
        pos = n;
      }

      out.table.push({
        pos,
        team,
        p:   asInt(cellAt(rows, i, iP)),
        w:   asInt(cellAt(rows, i, iW)),
        d:   asInt(cellAt(rows, i, iD)),
        l:   asInt(cellAt(rows, i, iL)),
        gf:  asInt(cellAt(rows, i, iGF)),
        ga:  asInt(cellAt(rows, i, iGA)),
        gd:  asInt(cellAt(rows, i, iGD)),
        pts: asInt(cellAt(rows, i, iPts))
      });
    }

    const koHdr = findHeaderFrom(rows, ["Round","Opponent"], i);
    if(koHdr >= 0){
      const kR   = colIndex(rows, koHdr, "Round");
      const kO   = colIndex(rows, koHdr, "Opponent");
      const kGF  = colIndexAny(rows, koHdr, ["GF","Goals For","For"]);
      const kGA  = colIndexAny(rows, koHdr, ["GA","Goals Against","Against"]);
      const kS   = colIndexAny(rows, koHdr, ["Score","Agg","Aggregate"]);
      const kPW  = colIndexAny(rows, koHdr, ["Pens Won","Penalties Won","Pens For"]);
      const kPL  = colIndexAny(rows, koHdr, ["Pens Lost","Penalties Lost","Pens Against"]);
      const useSingleScore = kGF < 0 && kS >= 0;

      for(let j = koHdr + 1; j < rows.length; j++){
        const round = cellAt(rows, j, kR);
        const opp   = cellAt(rows, j, kO);
        if(!round || !opp) continue;
        if(round.toLowerCase() === "round" && opp.toLowerCase() === "opponent") continue;

        let gf = 0, ga = 0;
        if(useSingleScore){
          const raw = cellAt(rows, j, kS);
          const m = String(raw).match(/^(\d+)\s*[-–—]\s*(\d+)/);
          if(m){ gf = +m[1]; ga = +m[2]; }
        } else {
          gf = asInt(cellAt(rows, j, kGF));
          ga = asInt(cellAt(rows, j, kGA));
        }

        /* Penalties: parse if both cells have a number */
        let pensWon = null, pensLost = null;
        if(kPW >= 0 && kPL >= 0){
          const w = cellAt(rows, j, kPW);
          const l = cellAt(rows, j, kPL);
          if(w !== "" && l !== ""){
            const wn = parseInt(w, 10);
            const ln = parseInt(l, 10);
            if(!isNaN(wn) && !isNaN(ln)){ pensWon = wn; pensLost = ln; }
          }
        }

        out.knockout.push({ round, opponent: opp, gf, ga, pensWon, pensLost });
      }
    }

    return (out.table.length || out.knockout.length) ? out : null;
  }

  /* ---------- Leaderboards ---------- */
  function parseLeaderTable(rows, startRow, metricName){
    const hdr = findHeaderFrom(rows, ["Rank","Name","Team",metricName], startRow);
    if(hdr < 0) return { rows: [], nextStart: startRow };

    const iRank    = colIndex(rows, hdr, "Rank");
    const iName    = colIndex(rows, hdr, "Name");
    const iTeam    = colIndex(rows, hdr, "Team");
    const iMetric  = colIndex(rows, hdr, metricName);
    const iMatches = colIndex(rows, hdr, "Matches");

    const out = [];
    let i = hdr + 1;
    while(i < rows.length){
      const rankRaw = cellAt(rows, i, iRank);
      const name    = cellAt(rows, i, iName);
      if(!name || name.toLowerCase() === "nan") break;
      if(rankRaw === "" || isNaN(parseInt(rankRaw, 10))) break;
      out.push([
        asInt(rankRaw),
        name,
        cellAt(rows, i, iTeam),
        asInt(cellAt(rows, i, iMetric)),
        asInt(cellAt(rows, i, iMatches))
      ]);
      i++;
    }
    return { rows: out, nextStart: i };
  }

  function parseLeaderboards(rows){
    const goalsTbl      = parseLeaderTable(rows, 0, "Goals");
    const assistsTbl    = parseLeaderTable(rows, goalsTbl.nextStart, "Assists");
    const uclGoalsTbl   = parseLeaderTable(rows, assistsTbl.nextStart, "Goals");
    const uclAssistsTbl = parseLeaderTable(rows, uclGoalsTbl.nextStart, "Assists");
    return {
      goals:      goalsTbl.rows,
      assists:    assistsTbl.rows,
      uclGoals:   uclGoalsTbl.rows,
      uclAssists: uclAssistsTbl.rows
    };
  }

  /* ---------- Squad ---------- */
  function parseSquad(rows){
    const hdr = findHeader(rows, ["Name","Position","OVR"]);
    if(hdr < 0) return [];
    const idx = {};
    for(const n of ["Name","Position","Age","OVR","Market Value","Wage","Contract"]){
      idx[n] = colIndex(rows, hdr, n);
    }
    if(idx.Name < 0) return [];

    const iApp    = colIndexAny(rows, hdr, ["Appearances","Apps","App"]);
    const iGoal   = colIndexAny(rows, hdr, ["Goals","Gls"]);
    const iAssist = colIndexAny(rows, hdr, ["Assists","Ast"]);
    const iCS     = colIndexAny(rows, hdr, ["Clean Sheets","CS","CleanSheets"]);

    const out = [];
    for(let i = hdr + 1; i < rows.length; i++){
      const name = cellAt(rows, i, idx.Name);
      if(!name || name.toLowerCase() === "nan") continue;
      out.push({
        name,
        pos:      cellAt(rows, i, idx.Position),
        age:      asInt(cellAt(rows, i, idx.Age)),
        ovr:      idx.OVR >= 0 ? asInt(cellAt(rows, i, idx.OVR)) : null,
        value:    cellAt(rows, i, idx["Market Value"]) || "—",
        wage:     cellAt(rows, i, idx.Wage)           || "—",
        contract: cellAt(rows, i, idx.Contract)       || "—",
        apps:    iApp    >= 0 ? asInt(cellAt(rows, i, iApp))    : null,
        goals:   iGoal   >= 0 ? asInt(cellAt(rows, i, iGoal))   : null,
        assists: iAssist >= 0 ? asInt(cellAt(rows, i, iAssist)) : null,
        cs:      iCS     >= 0 ? asInt(cellAt(rows, i, iCS))     : null
      });
    }
    return out;
  }

  /* ---------- Youth ---------- */
  function parseYouth(rows){
    const hdr = findHeader(rows, ["Name","Position","OVR"]);
    if(hdr < 0) return [];
    const idx = {};
    for(const n of ["Name","Position","Age","OVR","Potential","Development Plan","Nationality"]){
      idx[n] = colIndex(rows, hdr, n);
    }
    if(idx.Name < 0) return [];
    const out = [];
    for(let i = hdr + 1; i < rows.length; i++){
      const name = cellAt(rows, i, idx.Name);
      if(!name || name.toLowerCase() === "nan") continue;
      out.push({
        name,
        pos:         cellAt(rows, i, idx.Position),
        age:         asInt(cellAt(rows, i, idx.Age)),
        ovr:         asInt(cellAt(rows, i, idx.OVR)),
        potential:   cellAt(rows, i, idx.Potential) || "—",
        plan:        cellAt(rows, i, idx["Development Plan"]) || "—",
        nationality: cellAt(rows, i, idx.Nationality) || ""
      });
    }
    return out;
  }

  /* ---------- Finances ---------- */
  function parseFinances(rows){
    const hdr = findHeader(rows, ["Metric","Value"]);
    if(hdr < 0) return {};
    const iM = colIndex(rows, hdr, "Metric");
    const iV = colIndex(rows, hdr, "Value");
    if(iM < 0 || iV < 0) return {};
    const map = {
      "Earnings":   "earnings",
      "Expenses":   "expenses",
      "Profit":     "profit",
      "Club Worth": "clubWorth",
      "Projection": "projection"
    };
    const out = {};
    for(let i = hdr + 1; i < rows.length; i++){
      const k = cellAt(rows, i, iM);
      const key = map[k];
      if(key) out[key] = cellAt(rows, i, iV);
    }
    return out;
  }

  /* ---------- Cups ---------- */
  function parseCups(rows){
    const hdr = findHeader(rows, ["Competition","Winner","Result"]);
    if(hdr < 0) return [];
    const idx = {};
    for(const n of ["Competition","Winner","Result"]) idx[n] = colIndex(rows, hdr, n);
    if(idx.Competition < 0) return [];
    const out = [];
    for(let i = hdr + 1; i < rows.length; i++){
      const comp = cellAt(rows, i, idx.Competition);
      if(!comp || comp.toLowerCase() === "nan") continue;
      out.push([
        comp,
        cellAt(rows, i, idx.Winner),
        cellAt(rows, i, idx.Result)
      ]);
    }
    return out;
  }

  /* ---------- Career Summary ---------- */
  function parseTransfer(str){
    if(!str || str === "N/A" || str === "—" || str === "-") return null;
    const m = str.match(/^(.+?)\s*[-–—]\s*(.+)$/);
    if(m) return { player: m[1].trim(), fee: m[2].trim() };
    return { player: null, fee: str };
  }

  function parseScoreline(str){
    if(!str || str === "N/A" || str === "—" || str === "-") return null;
    const m = String(str).match(
      /^\s*(\d+)\s*[-–—]\s*(\d+)\s*(?:vs\.?\s*)?(.+?)\s*(?:\(([HAN])\))?\s*$/i
    );
    if(!m) return { raw: str };
    return {
      gf: +m[1],
      ga: +m[2],
      opp: m[3].trim(),
      venue: (m[4] || "").toUpperCase() || null,
      raw: str
    };
  }

  function parseCareerSummary(rows){
    const hdr = findHeader(rows, ["Metric","Value"]);
    if(hdr < 0) return null;
    const iM = colIndex(rows, hdr, "Metric");
    const iV = colIndex(rows, hdr, "Value");
    if(iM < 0 || iV < 0) return null;

    const summary = {
      manager: "", seasonLabel: "", competitions: {},
      leaguePosition: null, managerOfMonth: 0, managerOfYear: false,
      recordTransferPaid: null, recordTransferReceived: null,
      biggestWin: null, biggestDefeat: null,
      allComps: null
    };
    const records = { p:0, w:0, d:0, l:0, gf:0, ga:0 };

    for(let i = hdr + 1; i < rows.length; i++){
      const k = cellAt(rows, i, iM);
      const v = cellAt(rows, i, iV);
      if(!k) continue;

      if(k === "Manager")                          summary.manager = v;
      else if(k === "Season")                      summary.seasonLabel = v;
      else if(k === "Manager of the Month Awards") summary.managerOfMonth = asInt(v);
      else if(k === "Manager of the Year Award")   summary.managerOfYear = /^yes$/i.test(v);
      else if(k === "Record Transfer Fee Paid")    summary.recordTransferPaid = parseTransfer(v);
      else if(k === "Record Transfer Fee Received") summary.recordTransferReceived = parseTransfer(v);
      else if(k === "Biggest Win")                 summary.biggestWin = parseScoreline(v);
      else if(k === "Biggest Defeat")              summary.biggestDefeat = parseScoreline(v);
      else if(k === "Club Record - Played")        records.p = asInt(v);
      else if(k === "Club Record - Won")           records.w = asInt(v);
      else if(k === "Club Record - Drawn")         records.d = asInt(v);
      else if(k === "Club Record - Lost")          records.l = asInt(v);
      else if(k === "Club Record - Goals For")     records.gf = asInt(v);
      else if(k === "Club Record - Goals Against") records.ga = asInt(v);
      else if(/^League Position/i.test(k))         summary.leaguePosition = v;
      else summary.competitions[k] = v;
    }

    if(records.p || records.w || records.gf || records.ga) summary.allComps = records;
    return summary;
  }

  /* ---------- Excel entry point ---------- */
  async function parseExcel(file, meta){
    if(!isExcelReady()) throw new Error("Excel library not loaded");
    const buf = await file.arrayBuffer();
    const wb = window.XLSX.read(buf, { type: "array" });

    const allLeagues = {};
    for(const lg of LEAGUE_SHEETS){
      const rows = sheetToRows(wb, lg);
      if(!rows) continue;
      const parsed = parseLeague(rows);
      if(parsed.length) allLeagues[lg] = parsed;
    }

    let clubLeague = "Unknown";
    let clubRecord = null;
    const clubName = String(meta.club || "").toLowerCase();
    for(const lg of Object.keys(allLeagues)){
      for(const r of allLeagues[lg]){
        if(r[1].toLowerCase() === clubName){
          clubLeague = lg;
          clubRecord = { pos:r[0], p:r[2], w:r[3], d:r[4], l:r[5], gf:r[6], ga:r[7], gd:r[8], pts:r[9] };
          break;
        }
      }
      if(clubRecord) break;
    }
    if(!clubRecord) clubRecord = { pos:0, p:0, w:0, d:0, l:0, gf:0, ga:0, gd:0, pts:0 };

    const dashboardRows = sheetToRows(wb, SHEETS.dashboard) || [];
    const squadRows     = sheetToRows(wb, SHEETS.squad)     || [];
    const youthRows     = sheetToRows(wb, SHEETS.youth)     || [];
    const finRows       = sheetToRows(wb, SHEETS.finances)  || [];
    const cupRows       = sheetToRows(wb, SHEETS.cups)      || [];
    const summaryRows   = sheetToRows(wb, SHEETS.summary)   || [];
      const EURO_SHEET_OPTIONS = [
      { names: ["Champions League", "UEFA Champions League", "UCL"],          type: "ucl",  label: "Champions League" },
      { names: ["Europa League", "UEFA Europa League", "UEL"],               type: "uel",  label: "Europa League" },
      { names: ["Conference League", "UEFA Europa Conference League", "UECL"], type: "uecl", label: "Conference League" }
    ];
    let euroComp = null;
    for(const opt of EURO_SHEET_OPTIONS){
      for(const name of opt.names){
        const rows = sheetToRows(wb, name);
        if(!rows) continue;
        const parsed = parseChampionsLeague(rows);
        if(parsed){
          parsed.type  = opt.type;
          parsed.label = opt.label;
          euroComp = parsed;
          break;
        }
      }
      if(euroComp) break;
    }

    const leaderboards = parseLeaderboards(dashboardRows);

    return normalizeSeason({
      season:       meta.season,
      club:         meta.club,
      manager:      meta.manager,
      promoted:     meta.promoted,
      league:       clubLeague,
      leagueRecord: clubRecord,
      allLeagues,
      squad:        parseSquad(squadRows),
      youth:        parseYouth(youthRows),
      finances:     parseFinances(finRows),
      cups:         parseCups(cupRows),
      topScorers:   leaderboards.goals,
      topAssists:   leaderboards.assists,
      uclGoals:     leaderboards.uclGoals,
      uclAssists:   leaderboards.uclAssists,
      championsLeague: euroComp,
      summary:      parseCareerSummary(summaryRows)
    });
  }

  /* ---------- JSON path ---------- */
  async function parseJson(file){
    const text = await file.text();
    const data = JSON.parse(text);
    const list = Array.isArray(data)         ? data
               : Array.isArray(data.seasons) ? data.seasons
               : [data];
    const out = [];
    for(const raw of list){
      if(!raw || typeof raw !== "object") continue;
      if(!raw.season && !raw.club) continue;
      const s = normalizeSeason(raw);
      if(!Object.keys(s.allLeagues).length && !s.squad.length) continue;
      out.push(s);
    }
    return out;
  }

  /* ---------- normalise ---------- */
  function normLeaderboard(arr){
    return (arr || []).map(t => {
      if(Array.isArray(t)){
        return {
          rank: +t[0] || 0,
          name: t[1] || "",
          team: t[2] || "",
          value: +t[3] || 0,
          matches: +t[4] || 0
        };
      }
      return {
        rank: +t.rank || 0,
        name: t.name || "",
        team: t.team || "",
        value: +(t.value ?? t.goals ?? t.assists ?? 0),
        matches: +t.matches || 0
      };
    });
  }

  function normalizeSeason(raw){
    const allLeagues = {};
    const src = raw.allLeagues || {};
    for(const name of Object.keys(src)){
      const arr = src[name] || [];
      allLeagues[name] = arr.map(r => Array.isArray(r) ? {
        pos:r[0], team:r[1], p:r[2], w:r[3], d:r[4],
        l:r[5], gf:r[6], ga:r[7], gd:r[8], pts:r[9]
      } : {
        pos:r.pos, team:r.team, p:r.p, w:r.w, d:r.d,
        l:r.l, gf:r.gf, ga:r.ga, gd:r.gd, pts:r.pts
      }).filter(r => r.team);
    }

    const squad = (raw.squad || []).map(p => {
      if(Array.isArray(p)){
        return {
          name: p[0] || "Unknown", pos: p[1] || "",
          age: +p[2] || 0, ovr: +p[3] || 0,
          value: p[4] || "—", wage: p[5] || "—", contract: p[6] || "—",
          apps:    p.length > 7  && p[7]  != null ? +p[7]  : null,
          goals:   p.length > 8  && p[8]  != null ? +p[8]  : null,
          assists: p.length > 9  && p[9]  != null ? +p[9]  : null,
          cs:      p.length > 10 && p[10] != null ? +p[10] : null
        };
      }
      return {
        name: p.name || "Unknown", pos: p.pos || "",
        age: +p.age || 0, ovr: p.ovr != null ? +p.ovr : null,
        value: p.value || "—", wage: p.wage || "—", contract: p.contract || "—",
        apps:    p.apps    != null ? +p.apps    : null,
        goals:   p.goals   != null ? +p.goals   : null,
        assists: p.assists != null ? +p.assists : null,
        cs:      p.cs      != null ? +p.cs      : null
      };
    });

    const youth = (raw.youth || []).map(y => ({
      name:        y.name || "Unknown",
      pos:         y.pos  || "",
      age:         +y.age || 0,
      ovr:         +y.ovr || 0,
      potential:   y.potential   || "—",
      plan:        y.plan        || "—",
      nationality: y.nationality || ""
    }));

    const cups = (raw.cups || []).map(c => Array.isArray(c) ? {
      competition:c[0] || "", winner:c[1] || "", result:c[2] || ""
    } : {
      competition:c.competition || "", winner:c.winner || "", result:c.result || ""
    });

    const topScorers = normLeaderboard(raw.topScorers);
    const topAssists = normLeaderboard(raw.topAssists);
    const uclGoals   = normLeaderboard(raw.uclGoals);
    const uclAssists = normLeaderboard(raw.uclAssists);

    let championsLeague = null;
    if(raw.championsLeague && (raw.championsLeague.table?.length || raw.championsLeague.knockout?.length)){
      championsLeague = {
        group: raw.championsLeague.group || "",
        table: (raw.championsLeague.table || []).map(r => ({
          pos: +r.pos || 0,
          team: r.team || "",
          p: +r.p || 0, w: +r.w || 0, d: +r.d || 0, l: +r.l || 0,
          gf: +r.gf || 0, ga: +r.ga || 0, gd: +r.gd || 0, pts: +r.pts || 0
        })),
         knockout: (raw.championsLeague.knockout || []).map(k => ({
          round: k.round || "",
          opponent: k.opponent || "",
          gf: +k.gf || 0,
          ga: +k.ga || 0,
          pensWon:  k.pensWon  != null ? +k.pensWon  : null,
          pensLost: k.pensLost != null ? +k.pensLost : null
        }))
      };
    }

    let summary = raw.summary || null;
    if(summary){
      summary = {
        manager: summary.manager || "",
        seasonLabel: summary.seasonLabel || "",
        competitions: summary.competitions || {},
        leaguePosition: summary.leaguePosition || null,
        managerOfMonth: summary.managerOfMonth || 0,
        managerOfYear: !!summary.managerOfYear,
        recordTransferPaid: summary.recordTransferPaid || null,
        recordTransferReceived: summary.recordTransferReceived || null,
        biggestWin: summary.biggestWin || null,
        biggestDefeat: summary.biggestDefeat || null,
        allComps: summary.allComps || null
      };
    }

    let record = raw.leagueRecord;
    if(!record && raw.league && allLeagues[raw.league]){
      const me = allLeagues[raw.league].find(r =>
        r.team.toLowerCase() === String(raw.club || "").toLowerCase()
      );
      if(me) record = { pos:me.pos, p:me.p, w:me.w, d:me.d, l:me.l, gf:me.gf, ga:me.ga, gd:me.gd, pts:me.pts };
    }
    if(!record) record = { pos:0, p:0, w:0, d:0, l:0, gf:0, ga:0, gd:0, pts:0 };

    return {
      _id: uid(),
      season:   raw.season || "Unknown",
      club:     raw.club   || "My Club",
      league:   raw.league || "League",
      manager:  raw.manager || "",
      promoted: !!raw.promoted,
      sortKey:  raw.sortKey != null ? +raw.sortKey : null,
      leagueRecord: record,
      allLeagues, squad, youth, finances: raw.finances || {}, cups,
      topScorers, topAssists, uclGoals, uclAssists,
      championsLeague,
      summary
    };
  }

  window.CareerParse = {
    isExcelReady,
    parseExcel,
    parseJson,
    normalizeSeason
  };
})();