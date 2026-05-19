// ============================================================
//  UP & DOWN THE RIVER — Google Apps Script Backend
//  Paste this into: script.google.com → New Project
// ============================================================

const SS_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

// Sheet names
const SHEET_ARCHIVE  = 'Game Archive';
const SHEET_TOURNEY  = 'Tournament Stats';
const SHEET_LEADER   = 'Leaderboard Stats';
const SHEET_META     = 'Meta';

// ── DOGET / DOPOST ──────────────────────────────────────────
function doGet(e) {
  const action = e.parameter.action || 'getData';
  if (action === 'getData') return respond(getData());
  if (action === 'manualRefresh') return respond(manualStatsRefresh());
  return respond({ error: 'Unknown action' });
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action  = payload.action || 'saveData';
    if (action === 'saveData') return respond(saveData(payload));
    return respond({ error: 'Unknown action' });
  } catch(err) {
    return respond({ error: err.message });
  }
}

function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── GET DATA ─────────────────────────────────────────────────
function getData() {
  const meta = getMeta();
  return {
    allTime:           meta.allTime           || { games: [] },
    currentTournament: meta.currentTournament || { games: [] }
  };
}

// ── SAVE DATA ────────────────────────────────────────────────
function saveData(payload) {
  const { allTime, currentTournament } = payload;
  setMeta('allTime', allTime);
  setMeta('currentTournament', currentTournament);

  // Write sheets
  writeArchiveSheet(allTime.games || []);
  writeTournamentSheet(currentTournament);
  writeLeaderboardSheet(allTime);

  return { ok: true };
}

// ── META (stores JSON blobs in Meta sheet) ───────────────────
function getMeta() {
  const ss    = SpreadsheetApp.openById(SS_ID);
  let sheet   = ss.getSheetByName(SHEET_META);
  if (!sheet) { sheet = ss.insertSheet(SHEET_META); sheet.hideSheet(); }
  const data  = sheet.getDataRange().getValues();
  const obj   = {};
  data.forEach(row => { if (row[0]) { try { obj[row[0]] = JSON.parse(row[1]); } catch(_){} } });
  return obj;
}

function setMeta(key, value) {
  const ss    = SpreadsheetApp.openById(SS_ID);
  let sheet   = ss.getSheetByName(SHEET_META);
  if (!sheet) { sheet = ss.insertSheet(SHEET_META); sheet.hideSheet(); }
  const data  = sheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === key) { sheet.getRange(i+1, 2).setValue(JSON.stringify(value)); return; }
  }
  sheet.appendRow([key, JSON.stringify(value)]);
}

// ── GAME ARCHIVE SHEET ───────────────────────────────────────
function writeArchiveSheet(games) {
  const ss    = SpreadsheetApp.openById(SS_ID);
  let sheet   = ss.getSheetByName(SHEET_ARCHIVE);
  if (!sheet) sheet = ss.insertSheet(SHEET_ARCHIVE);
  sheet.clearContents();

  // Collect all unique player names in order of first appearance
  const allPlayers = [];
  games.forEach(g => {
    (g.players||[]).forEach(p => { if (!allPlayers.includes(p)) allPlayers.push(p); });
  });

  // Header rows
  const staticCols = ['Tournament','Game #','Date','Round #','Cards','Trump'];
  const playerCols = [];
  allPlayers.forEach(p => playerCols.push(`${p} Bid`, `${p} Tricks`, `${p} Score`, `${p} Made`));
  sheet.appendRow([...staticCols, ...playerCols]);

  // Style header
  const hdr = sheet.getRange(1, 1, 1, staticCols.length + playerCols.length);
  hdr.setFontWeight('bold').setBackground('#0b1e13').setFontColor('#c9a84c');

  games.forEach(game => {
    const { tournamentId, gameNum, date, rounds, players } = game;
    const dateStr = date ? new Date(date).toLocaleDateString() : '';
    (rounds || []).forEach(rd => {
      const row = [tournamentId, gameNum, dateStr, rd.roundNum, rd.cards, rd.trump];
      allPlayers.forEach(p => {
        const s = (rd.scores||[]).find(x=>x.name===p);
        if (s) { row.push(s.bid, s.tricks, s.score, s.made?'Y':'N'); }
        else   { row.push('','','',''); }
      });
      sheet.appendRow(row);
    });
  });

  // Auto-resize
  try { sheet.autoResizeColumns(1, staticCols.length + playerCols.length); } catch(_) {}
}

// ── TOURNAMENT STATS SHEET ───────────────────────────────────
function writeTournamentSheet(currentTournament) {
  const ss    = SpreadsheetApp.openById(SS_ID);
  let sheet   = ss.getSheetByName(SHEET_TOURNEY);
  if (!sheet) sheet = ss.insertSheet(SHEET_TOURNEY);
  sheet.clearContents();

  const games  = currentTournament.games || [];
  const stats  = calcStatsGAS(games);
  const players = [...new Set(games.flatMap(g=>g.players||[]))];

  writeStatsToSheet(sheet, players, stats, `Tournament: ${currentTournament.id||'—'}`);
}

// ── LEADERBOARD STATS SHEET ──────────────────────────────────
function writeLeaderboardSheet(allTime) {
  const ss    = SpreadsheetApp.openById(SS_ID);
  let sheet   = ss.getSheetByName(SHEET_LEADER);
  if (!sheet) sheet = ss.insertSheet(SHEET_LEADER);
  sheet.clearContents();

  const games  = allTime.games || [];
  const stats  = calcStatsGAS(games);
  const players = [...new Set(games.flatMap(g=>g.players||[]))];

  writeStatsToSheet(sheet, players, stats, 'All-Time Leaderboard');
}

function writeStatsToSheet(sheet, players, stats, title) {
  if (!players.length) return;
  sheet.appendRow([title]);
  sheet.getRange(1,1).setFontWeight('bold').setFontSize(13).setFontColor('#c9a84c');
  sheet.appendRow([]);

  const STAT_ROWS = [
    { section:'POINTS DISTRIBUTION' },
    { key:'totalTournamentPoints', label:'Total Tournament Points' },
    { key:'tp5', label:'Games Earning 5 T-Points' },
    { key:'tp4', label:'Games Earning 4 T-Points' },
    { key:'tp3', label:'Games Earning 3 T-Points' },
    { key:'tp2', label:'Games Earning 2 T-Points' },
    { key:'tp1', label:'Games Earning 1 T-Point' },
    { section:'FINANCIALS' },
    { key:'moneyLosses',     label:'Money From Losses',         fmt:'$' },
    { key:'moneyPenalties',  label:'Money From Penalties',      fmt:'$' },
    { key:'totalPot',        label:'Total Money in Pot',        fmt:'$' },
    { key:'mostMoneyOneGame',label:'Most Money Paid One Game',  fmt:'$' },
    { section:'GENERAL SCORING' },
    { key:'avgGamePoints',   label:'Average Game Points' },
    { key:'totalGamePoints', label:'Total Game Points' },
    { key:'totalSets',       label:'Total Number of Sets' },
    { key:'totalTricks',     label:'Total Number of Tricks' },
    { section:'GAME RECORDS' },
    { key:'mostSetsOneGame',   label:'Most Sets in One Game' },
    { key:'leastSetsOneGame',  label:'Least Sets in One Game' },
    { key:'mostTricksOneGame', label:'Most Tricks in One Game' },
    { key:'leastTricksOneGame',label:'Least Tricks in One Game' },
    { key:'lowestScore',       label:'Lowest Score Ever' },
    { key:'highestScore',      label:'Highest Score Ever' },
    { section:'STREAKS' },
    { key:'longestWinStreakGames',  label:'Longest Winning Streak (Games)' },
    { key:'longestLoseStreakGames', label:'Longest Losing Streak (Games)' },
    { key:'longestWinStreakHands',  label:'Longest Winning Streak (Hands)' },
    { key:'longestLoseStreakHands', label:'Longest Losing Streak (Hands)' },
    { key:'longestStreakNoPay',     label:'Longest Streak Without Paying' },
    { key:'longestStreakPay',       label:'Longest Streak With Paying' },
  ];

  // Header
  const headerRow = ['Statistic', ...players];
  sheet.appendRow(headerRow);
  const hr = sheet.getRange(3, 1, 1, headerRow.length);
  hr.setFontWeight('bold').setBackground('#112418').setFontColor('#c9a84c');

  let rowNum = 4;
  STAT_ROWS.forEach(row => {
    if (row.section) {
      sheet.appendRow([row.section]);
      sheet.getRange(rowNum, 1).setFontWeight('bold').setFontColor('#7a9080').setBackground('#0b1e13');
      sheet.getRange(rowNum, 1, 1, players.length+1).setBackground('#0b1e13');
    } else {
      const cells = [row.label];
      players.forEach(p => {
        const s = stats[p];
        let val = s ? s[row.key] : '—';
        if (val === null || val === undefined) val = '—';
        if (row.fmt === '$' && val !== '—') val = `$${(+val).toFixed(2)}`;
        cells.push(val);
      });
      sheet.appendRow(cells);
    }
    rowNum++;
  });

  try { sheet.autoResizeColumns(1, players.length + 1); } catch(_){}
}

// ── STATS CALCULATOR (GAS version mirrors client) ────────────
function calcStatsGAS(games) {
  const TRUMP_SEQUENCE = ['Hearts','Clubs','Diamonds','Spades'];
  const stats = {};
  const playerGames = {};

  games.forEach(game => {
    const { players, rounds, totals, financials, tPoints } = game;
    if (!players || !rounds) return;

    players.forEach(name => {
      if (!stats[name]) stats[name] = blankStats(name);
      if (!playerGames[name]) playerGames[name] = { hw:0, hl:0, bw:0, bl:0, np:0, pp:0, bnp:0, bpp:0 };
      const s  = stats[name];
      const tp = (tPoints||{})[name] || 0;
      const gameScore = (totals||{})[name] || 0;
      const fin = (financials||{})[name] || { losses:0, penalties:0, total:0 };
      const pg  = playerGames[name];

      s.totalTournamentPoints += tp;
      if (tp===5) s.tp5++; else if (tp===4) s.tp4++; else if (tp===3) s.tp3++;
      else if (tp===2) s.tp2++; else if (tp===1) s.tp1++;

      s.moneyLosses    += fin.losses||0;
      s.moneyPenalties += fin.penalties||0;
      s.totalPot       += fin.total||0;
      s.mostMoneyOneGame = Math.max(s.mostMoneyOneGame, fin.total||0);
      s.totalGamePoints += gameScore;
      s.lowestScore     = Math.min(s.lowestScore, gameScore);
      s.highestScore    = Math.max(s.highestScore, gameScore);
      s._games++;

      let gameSets=0, gameTricks=0;
      (rounds||[]).forEach(rd => {
        const prd = (rd.scores||[]).find(x=>x.name===name);
        if (!prd) return;
        if (!prd.made) gameSets++;
        gameTricks += prd.tricks||0;
        if (prd.made) { pg.hw++; pg.hl=0; pg.bw=Math.max(pg.bw,pg.hw); }
        else          { pg.hl++; pg.hw=0; pg.bl=Math.max(pg.bl,pg.hl); }
      });

      s.totalSets       += gameSets;
      s.totalTricks     += gameTricks;
      s.mostSetsOneGame   = Math.max(s.mostSetsOneGame, gameSets);
      s.leastSetsOneGame  = Math.min(s.leastSetsOneGame, gameSets);
      s.mostTricksOneGame  = Math.max(s.mostTricksOneGame, gameTricks);
      s.leastTricksOneGame = Math.min(s.leastTricksOneGame, gameTricks);

      const rank = Object.values(tPoints||{}).filter(v=>v>tp).length;
      if (rank===0) { s._cw++; s._cl=0; s.longestWinStreakGames=Math.max(s.longestWinStreakGames,s._cw); }
      else if (rank===(players.length-1)) { s._cl++; s._cw=0; s.longestLoseStreakGames=Math.max(s.longestLoseStreakGames,s._cl); }

      if ((fin.total||0)>0) { pg.pp++; pg.np=0; pg.bpp=Math.max(pg.bpp,pg.pp); }
      else                  { pg.np++; pg.pp=0; pg.bnp=Math.max(pg.bnp,pg.np); }
    });
  });

  Object.keys(stats).forEach(name => {
    const s  = stats[name];
    const pg = playerGames[name];
    s.avgGamePoints       = s._games ? +(s.totalGamePoints/s._games).toFixed(1) : 0;
    if (s.leastSetsOneGame   === Infinity) s.leastSetsOneGame   = 0;
    if (s.leastTricksOneGame === Infinity) s.leastTricksOneGame = 0;
    if (s.lowestScore  === Infinity)  s.lowestScore  = 0;
    if (s.highestScore === -Infinity) s.highestScore = 0;
    if (pg) {
      s.longestWinStreakHands  = pg.bw;
      s.longestLoseStreakHands = pg.bl;
      s.longestStreakNoPay     = pg.bnp;
      s.longestStreakPay       = pg.bpp;
    }
  });
  return stats;
}

function blankStats(name) {
  return {
    name, _games:0, _cw:0, _cl:0,
    totalTournamentPoints:0, tp5:0,tp4:0,tp3:0,tp2:0,tp1:0,
    moneyLosses:0, moneyPenalties:0, totalPot:0, mostMoneyOneGame:0,
    avgGamePoints:0, totalGamePoints:0, totalSets:0, totalTricks:0,
    mostSetsOneGame:0, leastSetsOneGame:Infinity,
    mostTricksOneGame:0, leastTricksOneGame:Infinity,
    lowestScore:Infinity, highestScore:-Infinity,
    longestWinStreakGames:0,  longestLoseStreakGames:0,
    longestWinStreakHands:0,  longestLoseStreakHands:0,
    longestStreakNoPay:0,     longestStreakPay:0
  };
}

// ── MANUAL REFRESH (run from Apps Script editor) ─────────────
function manualStatsRefresh() {
  const ss    = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName(SHEET_ARCHIVE);
  if (!sheet) return { error: 'Game Archive sheet not found' };

  const rows  = sheet.getDataRange().getValues();
  if (rows.length < 2) return { error: 'No data in archive' };

  // Parse archive back into game objects
  const header    = rows[0];
  const staticCnt = 6; // Tournament, Game#, Date, Round#, Cards, Trump
  const playerHeaders = header.slice(staticCnt);
  // Group every 4 columns into a player
  const archivePlayers = [];
  for (let i = 0; i < playerHeaders.length; i += 4) {
    const pName = playerHeaders[i].replace(' Bid','');
    archivePlayers.push(pName);
  }

  const gameMap = {};
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    const [tid, gnum, date, rnum, cards, trump, ...pData] = row;
    if (!tid) continue;
    const key = `${tid}_${gnum}`;
    if (!gameMap[key]) {
      gameMap[key] = { tournamentId:String(tid), gameNum:gnum, date:String(date),
                       players:[], rounds:[], totals:{}, financials:{}, tPoints:{} };
    }
    const g = gameMap[key];
    const scores = [];
    archivePlayers.forEach((p, i) => {
      const bid    = pData[i*4];
      const tricks = pData[i*4+1];
      const score  = pData[i*4+2];
      const made   = pData[i*4+3] === 'Y';
      if (bid !== '') {
        if (!g.players.includes(p)) g.players.push(p);
        scores.push({ name:p, bid:+bid, tricks:+tricks, score:+score, made });
      }
    });
    g.rounds.push({ roundNum:+rnum, cards:+cards, trump:String(trump), scores });
  }

  // Re-derive totals, financials, tPoints for each game
  const games = Object.values(gameMap).sort((a,b)=>a.gameNum-b.gameNum);
  games.forEach(game => {
    const totals = {};
    game.players.forEach(p => totals[p]=0);
    game.rounds.forEach(rd => rd.scores.forEach(s=>{ if(totals[s.name]!==undefined) totals[s.name]+=s.score; }));
    game.totals = totals;
    const max = Math.min(10, Math.floor(52/game.players.length));
    const threshold = (10*(2*max-1)-20)-(10*game.players.length);
    game.threshold = threshold;
    // Financials
    const entries = Object.entries(totals);
    const minScore = Math.min(...entries.map(([,v])=>v));
    const lastP = entries.filter(([,v])=>v===minScore).map(([n])=>n);
    const fin = {};
    entries.forEach(([name,score]) => {
      let losses=0, penalties=0;
      if (lastP.includes(name) && entries.length>1) losses=1;
      if (score<threshold) penalties=Math.ceil((threshold-score)/10);
      fin[name]={losses,penalties,total:losses+penalties};
    });
    game.financials=fin;
    // T-Points
    const sorted=entries.slice().sort((a,b)=>b[1]-a[1]);
    const tp={};
    let i=0;
    while(i<sorted.length){
      let j=i;
      while(j<sorted.length-1 && sorted[j][1]===sorted[j+1][1]) j++;
      const pts=game.players.length-i;
      for(let k=i;k<=j;k++) tp[sorted[k][0]]=pts;
      i=j+1;
    }
    game.tPoints=tp;
  });

  // Rebuild allTime and currentTournament
  const allTime = { games };
  const tids = [...new Set(games.map(g=>g.tournamentId))];
  const lastTid = tids[tids.length-1] || '';
  const currentTournament = { id:lastTid, games:games.filter(g=>g.tournamentId===lastTid) };

  setMeta('allTime', allTime);
  setMeta('currentTournament', currentTournament);

  // Write sheets
  writeArchiveSheet(games);
  writeTournamentSheet(currentTournament);
  writeLeaderboardSheet(allTime);

  return { ok:true, gamesLoaded:games.length };
}
