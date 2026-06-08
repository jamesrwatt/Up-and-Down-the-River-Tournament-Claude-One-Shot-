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
    // Browser sends text/plain to avoid CORS preflight; body is still JSON
    const raw = (e.postData && e.postData.contents) ? e.postData.contents : '{}';
    const payload = JSON.parse(raw);
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

// ── META (chunked JSON storage — avoids 50k char cell limit) ─
// Large values are split into 40k-char chunks stored as key, key__1, key__2…
// Reads reassemble them transparently.
const META_CHUNK_SIZE = 40000;

function getMeta() {
  const ss    = SpreadsheetApp.openById(SS_ID);
  let sheet   = ss.getSheetByName(SHEET_META);
  if (!sheet) { sheet = ss.insertSheet(SHEET_META); sheet.hideSheet(); }

  const data = sheet.getDataRange().getValues();
  // Build a raw map of every row: key → value string
  const rawMap = {};
  data.forEach(row => {
    const k = String(row[0] || '').trim();
    const v = String(row[1] || '');
    if (k) rawMap[k] = v;
  });

  // Reassemble chunked keys: key, key__1, key__2 … → joined JSON string
  const obj = {};
  const baseKeys = Object.keys(rawMap).filter(k => !k.match(/__\d+$/));
  baseKeys.forEach(baseKey => {
    let json = rawMap[baseKey];
    let i = 1;
    while (rawMap[`${baseKey}__${i}`] !== undefined) {
      json += rawMap[`${baseKey}__${i}`];
      i++;
    }
    try { obj[baseKey] = JSON.parse(json); } catch(_) {}
  });
  return obj;
}

function setMeta(key, value) {
  const ss    = SpreadsheetApp.openById(SS_ID);
  let sheet   = ss.getSheetByName(SHEET_META);
  if (!sheet) { sheet = ss.insertSheet(SHEET_META); sheet.hideSheet(); }

  const json   = JSON.stringify(value);
  const chunks = [];
  for (let i = 0; i < json.length; i += META_CHUNK_SIZE) {
    chunks.push(json.slice(i, i + META_CHUNK_SIZE));
  }
  const chunkKeys = [key];
  for (let i = 1; i < chunks.length; i++) chunkKeys.push(key + '__' + i);

  // Read all current rows, remove this key's old chunks, append new chunks
  // Then rewrite the whole sheet in one setValues call
  const existing = sheet.getDataRange().getValues();
  const keyPattern = new RegExp('^' + key.replace(/[.*+?^${}()|[\]\\]/g,'\\$&') + '(__\\d+)?$');
  const kept = existing.filter(row => !keyPattern.test(String(row[0] || '').trim()));
  const newRows = chunkKeys.map(function(k, i) { return [k, chunks[i]]; });
  const allRows = kept.concat(newRows);

  sheet.clearContents();
  if (allRows.length > 0) {
    sheet.getRange(1, 1, allRows.length, 2).setValues(allRows);
  }
}

// ── GAME ARCHIVE SHEET ───────────────────────────────────────
function writeArchiveSheet(games) {
  const ss    = SpreadsheetApp.openById(SS_ID);
  let sheet   = ss.getSheetByName(SHEET_ARCHIVE);
  if (!sheet) sheet = ss.insertSheet(SHEET_ARCHIVE);
  sheet.clearContents();
  sheet.clearFormats();

  if (!games || games.length === 0) {
    sheet.getRange(1, 1).setValue('No games recorded yet.');
    return;
  }

  // Collect all unique player names
  const allPlayers = [];
  games.forEach(g => {
    (g.players || []).forEach(p => { if (!allPlayers.includes(p)) allPlayers.push(p); });
  });

  const staticCols = ['Tournament', 'Game #', 'Date', 'Round #', 'Cards', 'Trump'];
  const playerCols = [];
  allPlayers.forEach(p => playerCols.push(`${p} Bid`, `${p} Tricks`, `${p} Score`, `${p} Made`));
  const fullHeader = [...staticCols, ...playerCols];
  const numCols    = fullHeader.length;

  // ── Build ALL rows in memory first ──────────────────────────
  const allRows = [fullHeader]; // row 0 = header
  games.forEach(game => {
    const { tournamentId, gameNum, date, rounds } = game;
    const dateStr = date ? new Date(date).toLocaleDateString() : '';
    (rounds || []).forEach(rd => {
      const row = [
        tournamentId || '', gameNum || '', dateStr,
        rd.roundNum || '', rd.cards || '', rd.trump || ''
      ];
      allPlayers.forEach(p => {
        const s = (rd.scores || []).find(x => x.name === p);
        if (s) row.push(s.bid, s.tricks, s.score, s.made ? 'Y' : 'N');
        else   row.push('', '', '', '');
      });
      allRows.push(row);
    });
  });

  // ── Write all rows in one call ───────────────────────────────
  sheet.getRange(1, 1, allRows.length, numCols).setValues(allRows);

  // ── Formatting (bulk range ops, no per-row calls) ────────────
  sheet.getRange(1, 1, 1, numCols)
       .setFontWeight('bold')
       .setBackground('#2c3e50')
       .setFontColor('#ffffff');

  try { sheet.setFrozenRows(1); } catch(_) {}
  try { sheet.autoResizeColumns(1, numCols); } catch(_) {}
}

// ── TOURNAMENT STATS SHEET ───────────────────────────────────
function writeTournamentSheet(currentTournament) {
  const ss    = SpreadsheetApp.openById(SS_ID);
  let sheet   = ss.getSheetByName(SHEET_TOURNEY);
  if (!sheet) sheet = ss.insertSheet(SHEET_TOURNEY);
  sheet.clearContents();

  const games   = (currentTournament && currentTournament.games) ? currentTournament.games : [];
  const players = [...new Set(games.flatMap(g => g.players || []))];
  const maxPlayers = games.length ? Math.max.apply(null, games.map(function(g){ return (g.players||[]).length; })) : 5;

  if (players.length === 0) {
    sheet.appendRow(['No tournament data yet.']);
    return;
  }

  const stats = calcStatsGAS(games);
  writeStatsToSheet(sheet, players, stats, `Tournament: ${(currentTournament && currentTournament.id) || '—'}`, maxPlayers);
}

// ── LEADERBOARD STATS SHEET ──────────────────────────────────
function writeLeaderboardSheet(allTime) {
  const ss    = SpreadsheetApp.openById(SS_ID);
  let sheet   = ss.getSheetByName(SHEET_LEADER);
  if (!sheet) sheet = ss.insertSheet(SHEET_LEADER);
  sheet.clearContents();

  const games   = (allTime && allTime.games) ? allTime.games : [];
  const players = [...new Set(games.flatMap(g => g.players || []))];
  const maxPlayers = games.length ? Math.max.apply(null, games.map(function(g){ return (g.players||[]).length; })) : 5;

  if (players.length === 0) {
    sheet.appendRow(['No all-time data yet.']);
    return;
  }

  const stats = calcStatsGAS(games);
  writeStatsToSheet(sheet, players, stats, 'All-Time Leaderboard', maxPlayers);
}

// ── WRITE STATS TABLE TO SHEET ───────────────────────────────
function writeStatsToSheet(sheet, players, stats, title, maxPlayers) {
  if (!players || players.length === 0) return;
  const n = Math.max(1, maxPlayers || 5);

  // Sort players: T-Points DESC → totalPot ASC → avgGamePoints DESC
  const sorted = players.slice().sort(function(a, b) {
    const sa = stats[a] || {};
    const sb = stats[b] || {};
    var tpDiff = (sb.totalTournamentPoints || 0) - (sa.totalTournamentPoints || 0);
    if (tpDiff !== 0) return tpDiff;
    var potDiff = (sa.totalPot || 0) - (sb.totalPot || 0);
    if (potDiff !== 0) return potDiff;
    return (sb.avgGamePoints || 0) - (sa.avgGamePoints || 0);
  });

  // Build dynamic T-Points rows
  const tpRows = [];
  for (var pts = n; pts >= 1; pts--) {
    tpRows.push({ key:'tp'+pts, label:'Games Earning '+pts+' T-Point'+(pts!==1?'s':''), dir: pts===n?'high':pts===1?'low':null });
  }

  const STAT_ROWS = [
    { section:'POINTS DISTRIBUTION' },
    { key:'totalTournamentPoints', label:'Total Tournament Points', dir:'high' }
  ].concat(tpRows).concat([{ section:'FINANCIALS' },
    { key:'moneyLosses',      label:'Money From Losses',       fmt:'$',      dir:'low'  },
    { key:'moneyPenalties',   label:'Money From Penalties',    fmt:'$',      dir:'low'  },
    { key:'totalPot',         label:'Total Money in Pot',      fmt:'$',      dir:'low'  },
    { key:'mostMoneyOneGame', label:'Most Money Paid One Game',fmt:'$',      dir:'low'  },
    { section:'GENERAL SCORING' },
    { key:'avgGamePoints',    label:'Average Game Points',                   dir:'high' },
    { key:'totalGamePoints',  label:'Total Game Points',                     dir:'high' },
    { key:'totalSets',        label:'Total Number of Sets',                  dir:'low'  },
    { key:'totalTricks',      label:'Total Number of Tricks',                dir:'high' },
    { section:'GAME RECORDS' },
    { key:'mostSetsOneGame',    label:'Most Sets in One Game',               dir:'low'  },
    { key:'leastSetsOneGame',   label:'Least Sets in One Game',              dir:'low'  },
    { key:'mostTricksOneGame',  label:'Most Tricks in One Game',             dir:'high' },
    { key:'leastTricksOneGame', label:'Least Tricks in One Game',            dir:'high' },
    { key:'lowestScore',        label:'Lowest Score Ever',                   dir:'high' },
    { key:'highestScore',       label:'Highest Score Ever',                  dir:'high' },
    { section:'STREAKS' },
    { key:'longestWinStreakGames',   label:'Longest Winning Streak (Games)',        dir:'high' },
    { key:'longestLoseStreakGames',  label:'Longest Losing Streak (Games)',         dir:'low'  },
    { key:'longestWinStreakInGame',  label:'Longest Winning Streak (in 1 Game)',    dir:'high' },
    { key:'longestLoseStreakInGame', label:'Longest Losing Streak (in 1 Game)',     dir:'low'  },
    { key:'longestWinStreakHands',   label:'Longest Winning Streak (Across Games)', dir:'high' },
    { key:'longestLoseStreakHands',  label:'Longest Losing Streak (Across Games)',  dir:'low'  },
    { key:'longestStreakNoPay',      label:'Longest Streak Without Paying',         dir:'high' },
    { key:'longestStreakPay',        label:'Longest Streak With Paying',            dir:'low'  },
  ]);

  const numCols = sorted.length + 1; // stat label + one col per player

  // ── Row 1: Title ──────────────────────────────────────────
  // ── Row 2: Column headers ─────────────────────────────────
  // ── Rows 3+: Stats — build entire 2D array first ──────────
  const allData = [
    [title],          // row 1 — title (written separately for merge)
    ['Statistic', ...sorted], // row 2 — player headers
  ];

  // Metadata arrays for deferred formatting
  const sectionRows   = [];      // sheetRow numbers that are section dividers
  const bestCells     = [];      // {row, col} to colour green
  const worstCells    = [];      // {row, col} to colour red
  const evenDataRows  = [];      // sheetRow numbers for alternating shading
  let sheetRow = 3;

  STAT_ROWS.forEach(row => {
    if (row.section) {
      const sectionLabel = [row.section];
      for (let i = 1; i < numCols; i++) sectionLabel.push('');
      allData.push(sectionLabel);
      sectionRows.push(sheetRow);
    } else {
      const rawVals = sorted.map(function(p) {
        const s = stats[p];
        return (s && s[row.key] !== undefined && s[row.key] !== null) ? +s[row.key] : null;
      });

      let bestVal = null, worstVal = null;
      if (row.dir && sorted.length > 1) {
        const defined = rawVals.filter(function(v) { return v !== null; });
        if (defined.length > 1) {
          const maxV = Math.max.apply(null, defined);
          const minV = Math.min.apply(null, defined);
          if (maxV !== minV) {
            bestVal  = row.dir === 'high' ? maxV : minV;
            worstVal = row.dir === 'high' ? minV : maxV;
          }
        }
      }

      const cells = [row.label];
      rawVals.forEach(function(v, i) {
        if (v === null) { cells.push('—'); return; }
        cells.push(row.fmt === '$' ? '$' + Math.round(v) : Math.round(v));
        if (bestVal  !== null && v === bestVal)  bestCells.push({r: sheetRow, c: i + 2});
        if (worstVal !== null && v === worstVal) worstCells.push({r: sheetRow, c: i + 2});
      });
      allData.push(cells);

      if (sheetRow % 2 === 0) evenDataRows.push(sheetRow);
    }
    sheetRow++;
  });

  // ── Bulk write all data in one call ───────────────────────
  const totalRows = allData.length;

  // Title row (single cell, will be merged across all cols)
  sheet.getRange(1, 1, 1, numCols)
       .merge()
       .setValue(title)
       .setFontWeight('bold')
       .setFontSize(13)
       .setBackground('#2c3e50')
       .setFontColor('#ffffff')
       .setHorizontalAlignment('left');

  // Write rows 2+ (headers + data) in one setValues call
  if (allData.length > 1) {
    sheet.getRange(2, 1, allData.length - 1, numCols).setValues(allData.slice(1));
  }

  // ── Bulk formatting after write ────────────────────────────

  // Header row (row 2)
  sheet.getRange(2, 1, 1, numCols)
       .setFontWeight('bold')
       .setBackground('#3d5166')
       .setFontColor('#ffffff');

  // Section rows — apply all in one loop with minimal calls
  sectionRows.forEach(function(r) {
    sheet.getRange(r, 1, 1, numCols)
         .setFontWeight('bold')
         .setFontColor('#ffffff')
         .setBackground('#4a6070')
         .setFontSize(9);
  });

  // Alternate row shading — collect contiguous runs to minimise range calls
  if (evenDataRows.length > 0) {
    // Shade all even data rows using a single conditional batch
    evenDataRows.forEach(function(r) {
      // Only shade if not a section row
      if (sectionRows.indexOf(r) === -1) {
        sheet.getRange(r, 1, 1, numCols).setBackground('#f4f7f6');
      }
    });
  }

  // Centre-align all player-value columns (cols 2..numCols), rows 3..end
  if (sorted.length > 0 && totalRows > 2) {
    sheet.getRange(3, 2, totalRows - 2, sorted.length).setHorizontalAlignment('center');
  }

  // Green/red cell colouring — individual cells, but grouped by colour
  // Apply greens in one pass, reds in another
  bestCells.forEach(function(cell) {
    sheet.getRange(cell.r, cell.c).setFontColor('#27ae60').setFontWeight('bold');
  });
  worstCells.forEach(function(cell) {
    sheet.getRange(cell.r, cell.c).setFontColor('#e74c3c').setFontWeight('bold');
  });

  // ── Final touches ─────────────────────────────────────────
  try { sheet.setFrozenRows(2); } catch(_) {}
  try { sheet.autoResizeColumns(1, numCols); } catch(_) {}
}

// ── STATS CALCULATOR (GAS version mirrors client) ────────────
function calcStatsGAS(games) {
  const stats       = {};
  const playerGames = {};

  games.forEach(game => {
    const { players, rounds, financials, tPoints } = game;
    if (!players || !rounds || players.length === 0) return;

    // ── Recompute per-game score from rounds (don't trust stored totals) ──
    const gameScores = {};
    players.forEach(p => { gameScores[p] = 0; });
    (rounds || []).forEach(rd => {
      (rd.scores || []).forEach(s => {
        if (gameScores[s.name] !== undefined) {
          // Always derive score from bid/tricks — ignore any stored score value
          const made  = (Number(s.bid) === Number(s.tricks));
          const score = made ? 10 + Number(s.tricks) : Number(s.tricks);
          gameScores[s.name] += score;
        }
      });
    });

    players.forEach(name => {
      if (!stats[name]) stats[name] = blankStats(name);
      if (!playerGames[name]) playerGames[name] = { hw:0, hl:0, bw:0, bl:0, np:0, pp:0, bnp:0, bpp:0 };

      const s         = stats[name];
      const pg        = playerGames[name];
      const gameScore = gameScores[name] || 0;      // recomputed, reliable
      const tp        = Number((tPoints||{})[name]) || 0;
      const fin       = (financials||{})[name] || { losses:0, penalties:0, total:0 };

      // ── Tournament points ──
      s.totalTournamentPoints += tp;
      if (tp >= 1 && tp <= 10) s['tp' + tp] = (s['tp' + tp] || 0) + 1;

      // ── Financials ──
      s.moneyLosses     += Number(fin.losses)   || 0;
      s.moneyPenalties  += Number(fin.penalties) || 0;
      s.totalPot        += Number(fin.total)     || 0;
      s.mostMoneyOneGame = Math.max(s.mostMoneyOneGame, Number(fin.total) || 0);

      // ── Game totals (now recomputed, not from stored totals) ──
      s.totalGamePoints += gameScore;
      s.lowestScore      = Math.min(s.lowestScore,  gameScore);
      s.highestScore     = Math.max(s.highestScore, gameScore);
      s._games++;

      // ── Per-round stats + in-game streaks ──
      let gameSets = 0, gameTricks = 0;
      let inGameWin = 0, inGameLose = 0, bestInGameWin = 0, bestInGameLose = 0;
      (rounds || []).forEach(rd => {
        const prd = (rd.scores || []).find(x => x.name === name);
        if (!prd) return;
        const made   = (Number(prd.bid) === Number(prd.tricks));
        const tricks = Number(prd.tricks) || 0;
        if (!made) gameSets++;
        gameTricks += tricks;
        // Across-games hand streak
        if (made) { pg.hw++; pg.hl = 0; pg.bw = Math.max(pg.bw, pg.hw); }
        else      { pg.hl++; pg.hw = 0; pg.bl = Math.max(pg.bl, pg.hl); }
        // In-game streak (resets each game)
        if (made) { inGameWin++; inGameLose = 0; bestInGameWin  = Math.max(bestInGameWin,  inGameWin);  }
        else      { inGameLose++; inGameWin = 0; bestInGameLose = Math.max(bestInGameLose, inGameLose); }
      });

      s.totalSets        += gameSets;
      s.totalTricks      += gameTricks;
      s.mostSetsOneGame    = Math.max(s.mostSetsOneGame,   gameSets);
      s.leastSetsOneGame   = Math.min(s.leastSetsOneGame,  gameSets);
      s.mostTricksOneGame  = Math.max(s.mostTricksOneGame,  gameTricks);
      s.leastTricksOneGame = Math.min(s.leastTricksOneGame, gameTricks);
      s.longestWinStreakInGame  = Math.max(s.longestWinStreakInGame,  bestInGameWin);
      s.longestLoseStreakInGame = Math.max(s.longestLoseStreakInGame, bestInGameLose);

      // ── Game win/lose streaks ──
      // Win  = sole highest game score (no tie for first)
      // Loss = tied for or at the lowest game score (tie for last = loss for all at that score)
      const gsVals   = Object.values(gameScores);
      const minScore = Math.min.apply(null, gsVals);
      const maxScore = Math.max.apply(null, gsVals);
      const isWin    = gameScore === maxScore && gsVals.filter(v => v === maxScore).length === 1;
      const isLoss   = gameScore === minScore;

      if (isWin) {
        s._cw++; s._cl = 0;
        s.longestWinStreakGames  = Math.max(s.longestWinStreakGames, s._cw);
      } else if (isLoss) {
        s._cl++; s._cw = 0;
        s.longestLoseStreakGames = Math.max(s.longestLoseStreakGames, s._cl);
      } else {
        s._cw = 0; s._cl = 0;
      }

      // ── Pay streaks ──
      if ((Number(fin.total) || 0) > 0) {
        pg.pp++; pg.np = 0; pg.bpp = Math.max(pg.bpp, pg.pp);
      } else {
        pg.np++; pg.pp = 0; pg.bnp = Math.max(pg.bnp, pg.np);
      }
    });
  });

  // ── Finalise derived stats ──
  Object.keys(stats).forEach(name => {
    const s  = stats[name];
    const pg = playerGames[name];
    s.avgGamePoints       = s._games ? Math.round(s.totalGamePoints / s._games) : 0;
    if (s.leastSetsOneGame   === Infinity)  s.leastSetsOneGame   = 0;
    if (s.leastTricksOneGame === Infinity)  s.leastTricksOneGame = 0;
    if (s.lowestScore        === Infinity)  s.lowestScore        = 0;
    if (s.highestScore       === -Infinity) s.highestScore       = 0;
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
    totalTournamentPoints:0,
    tp10:0,tp9:0,tp8:0,tp7:0,tp6:0,tp5:0,tp4:0,tp3:0,tp2:0,tp1:0,
    moneyLosses:0, moneyPenalties:0, totalPot:0, mostMoneyOneGame:0,
    avgGamePoints:0, totalGamePoints:0, totalSets:0, totalTricks:0,
    mostSetsOneGame:0,   leastSetsOneGame:Infinity,
    mostTricksOneGame:0, leastTricksOneGame:Infinity,
    lowestScore:Infinity, highestScore:-Infinity,
    longestWinStreakGames:0,   longestLoseStreakGames:0,
    longestWinStreakHands:0,   longestLoseStreakHands:0,
    longestWinStreakInGame:0,  longestLoseStreakInGame:0,
    longestStreakNoPay:0,      longestStreakPay:0
  };
}

// ── MANUAL REFRESH (run from Apps Script editor) ─────────────
function manualStatsRefresh() {
  try {
    const ss    = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName(SHEET_ARCHIVE);
    if (!sheet) {
      Logger.log('ERROR: Game Archive sheet not found. Create it first by pushing data from the app.');
      return { error: 'Game Archive sheet not found' };
    }

    const allData = sheet.getDataRange().getValues();
    Logger.log('Archive rows found: ' + allData.length);

    // Need at least header row + one data row
    if (allData.length < 2) {
      Logger.log('ERROR: No data rows in Game Archive (only header or empty).');
      return { error: 'No data in archive — add games first' };
    }

    // ── Parse header to discover player columns ────────────
    const header    = allData[0];
    const staticCnt = 6; // Tournament, Game#, Date, Round#, Cards, Trump
    const playerHeaders = header.slice(staticCnt).map(String);

    // Build player list from header: "PlayerName Bid" → "PlayerName"
    const archivePlayers = [];
    for (let i = 0; i < playerHeaders.length; i += 4) {
      const raw = playerHeaders[i] || '';
      const pName = raw.replace(/\s*Bid\s*$/i, '').trim();
      if (pName) archivePlayers.push(pName);
    }

    Logger.log('Players detected in archive header: ' + archivePlayers.join(', '));
    if (archivePlayers.length === 0) {
      Logger.log('ERROR: Could not detect player columns in header row.');
      return { error: 'No player columns found in archive header' };
    }

    // ── Parse data rows into game map ─────────────────────
    const gameMap = {};
    let skippedRows = 0;

    for (let r = 1; r < allData.length; r++) {
      const row  = allData[r];
      const tid  = String(row[0] || '').trim();
      const gnum = row[1];
      const date = String(row[2] || '').trim();
      const rnum = row[3];
      const cards = row[4];
      const trump = String(row[5] || '').trim();

      // Skip rows missing required fields
      if (!tid || tid === '' || gnum === '' || gnum === null || gnum === undefined) {
        skippedRows++;
        continue;
      }

      const key = `${tid}__${gnum}`;
      if (!gameMap[key]) {
        gameMap[key] = {
          tournamentId: tid,
          gameNum:      Number(gnum),
          date:         date,
          players:      [],
          rounds:       [],
          totals:       {},
          financials:   {},
          tPoints:      {}
        };
      }

      const g      = gameMap[key];
      const pData  = row.slice(staticCnt);
      const scores = [];

      archivePlayers.forEach((p, i) => {
        const bidRaw    = pData[i * 4];
        const tricksRaw = pData[i * 4 + 1];
        // Score column (index i*4+2) is intentionally IGNORED —
        // we always recompute from bid & tricks to guard against
        // manually-entered or cumulative values in the sheet.
        const madeRaw   = pData[i * 4 + 3];

        // Only include player if they have bid data for this round
        if (bidRaw !== '' && bidRaw !== null && bidRaw !== undefined) {
          if (!g.players.includes(p)) g.players.push(p);
          const bid    = Number(bidRaw);
          const tricks = Number(tricksRaw);
          const made   = bid === tricks;                  // recompute
          const score  = made ? 10 + tricks : tricks;    // recompute
          scores.push({ name: p, bid, tricks, score, made });
        }
      });

      // Only push round if it has at least one score
      if (scores.length > 0) {
        g.rounds.push({
          roundNum: Number(rnum),
          cards:    Number(cards),
          trump:    trump,
          scores:   scores
        });
      }
    }

    Logger.log('Skipped rows (blank/invalid): ' + skippedRows);

    const games = Object.values(gameMap)
      .filter(g => g.players.length > 0 && g.rounds.length > 0)
      .sort((a, b) => {
        // Sort by tournament id first, then game number
        if (a.tournamentId < b.tournamentId) return -1;
        if (a.tournamentId > b.tournamentId) return  1;
        return a.gameNum - b.gameNum;
      });

    Logger.log('Valid games parsed: ' + games.length);
    if (games.length === 0) {
      Logger.log('No valid games found. Check archive data format.');
      return { error: 'No valid games parsed from archive' };
    }

    // ── Re-derive totals, financials, tPoints per game ────
    games.forEach(game => {
      const { players, rounds } = game;

      // Totals
      const totals = {};
      players.forEach(p => { totals[p] = 0; });
      rounds.forEach(rd => {
        rd.scores.forEach(s => {
          if (totals[s.name] !== undefined) totals[s.name] += s.score;
        });
      });
      game.totals = totals;

      // Threshold & financials
      const maxCards  = Math.min(10, Math.floor(52 / players.length));
      const threshold = (10 * (2 * maxCards - 1) - 20) - (10 * players.length);
      game.threshold  = threshold;
      game.maxCards   = maxCards;

      const entries  = Object.entries(totals);
      const minScore = Math.min(...entries.map(([, v]) => v));
      const lastP    = entries.filter(([, v]) => v === minScore).map(([n]) => n);
      const fin      = {};
      entries.forEach(([name, score]) => {
        const losses    = (lastP.includes(name) && entries.length > 1) ? 1 : 0;
        const penalties = score < threshold ? Math.ceil((threshold - score) / 10) : 0;
        fin[name] = { losses, penalties, total: losses + penalties };
      });
      game.financials = fin;

      // T-Points with tie-sharing
      const sorted = entries.slice().sort((a, b) => b[1] - a[1]);
      const tp = {};
      let i = 0;
      while (i < sorted.length) {
        let j = i;
        while (j < sorted.length - 1 && sorted[j][1] === sorted[j + 1][1]) j++;
        const pts = players.length - i;
        for (let k = i; k <= j; k++) tp[sorted[k][0]] = pts;
        i = j + 1;
      }
      game.tPoints = tp;
    });

    // ── Rebuild allTime and currentTournament ─────────────
    const allTime = { games };
    const tids    = [...new Set(games.map(g => g.tournamentId))];
    const lastTid = tids[tids.length - 1] || '';
    const currentTournament = {
      id:    lastTid,
      games: games.filter(g => g.tournamentId === lastTid)
    };

    Logger.log('Tournaments found: ' + tids.join(', '));
    Logger.log('Current tournament: ' + lastTid + ' (' + currentTournament.games.length + ' games)');

    // ── Persist to Meta ───────────────────────────────────
    setMeta('allTime',           allTime);
    setMeta('currentTournament', currentTournament);

    // ── Write all sheets ──────────────────────────────────
    // Re-write archive with clean formatting (keeps data unchanged)
    writeArchiveSheet(games);
    writeTournamentSheet(currentTournament);
    writeLeaderboardSheet(allTime);

    const result = {
      ok:             true,
      gamesLoaded:    games.length,
      tournaments:    tids,
      currentTournId: lastTid
    };
    Logger.log('manualStatsRefresh complete: ' + JSON.stringify(result));
    return result;

  } catch (err) {
    Logger.log('EXCEPTION in manualStatsRefresh: ' + err.message + '\n' + err.stack);
    return { error: err.message };
  }
}
