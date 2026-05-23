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

  // Build map of all chunk keys for this base key
  const chunkKeys = [key];
  for (let i = 1; i < chunks.length; i++) chunkKeys.push(`${key}__${i}`);

  // Read current sheet to find existing rows
  const data = sheet.getDataRange().getValues();

  // Delete ALL existing rows for this key (base + any old chunks)
  // Work backwards so row indices stay valid
  const rowsToDelete = [];
  data.forEach((row, idx) => {
    const k = String(row[0] || '').trim();
    if (k === key || k.match(new RegExp(`^${key}__\\d+$`))) {
      rowsToDelete.push(idx + 1); // 1-indexed
    }
  });
  for (let i = rowsToDelete.length - 1; i >= 0; i--) {
    sheet.deleteRow(rowsToDelete[i]);
  }

  // Append all chunks as new rows
  chunks.forEach((chunk, i) => {
    sheet.appendRow([chunkKeys[i], chunk]);
  });
}

// ── GAME ARCHIVE SHEET ───────────────────────────────────────
function writeArchiveSheet(games) {
  const ss    = SpreadsheetApp.openById(SS_ID);
  let sheet   = ss.getSheetByName(SHEET_ARCHIVE);
  if (!sheet) sheet = ss.insertSheet(SHEET_ARCHIVE);
  sheet.clearContents();

  if (!games || games.length === 0) {
    sheet.appendRow(['No games recorded yet.']);
    return;
  }

  // Collect all unique player names in order of first appearance
  const allPlayers = [];
  games.forEach(g => {
    (g.players||[]).forEach(p => { if (!allPlayers.includes(p)) allPlayers.push(p); });
  });

  // Build header
  const staticCols = ['Tournament','Game #','Date','Round #','Cards','Trump'];
  const playerCols = [];
  allPlayers.forEach(p => playerCols.push(`${p} Bid`, `${p} Tricks`, `${p} Score`, `${p} Made`));
  const fullHeader = [...staticCols, ...playerCols];
  sheet.appendRow(fullHeader);

  // Style header — navy background, white text
  const hdr = sheet.getRange(1, 1, 1, fullHeader.length);
  hdr.setFontWeight('bold').setBackground('#2c3e50').setFontColor('#ffffff');

  // Data rows
  games.forEach(game => {
    const { tournamentId, gameNum, date, rounds, players } = game;
    const dateStr = date ? new Date(date).toLocaleDateString() : '';
    (rounds || []).forEach(rd => {
      const row = [tournamentId || '', gameNum || '', dateStr, rd.roundNum || '', rd.cards || '', rd.trump || ''];
      allPlayers.forEach(p => {
        const s = (rd.scores||[]).find(x => x.name === p);
        if (s) { row.push(s.bid, s.tricks, s.score, s.made ? 'Y' : 'N'); }
        else   { row.push('', '', '', ''); }
      });
      sheet.appendRow(row);
    });
  });

  // Freeze header row and auto-resize
  try { sheet.setFrozenRows(1); } catch(_) {}
  try { sheet.autoResizeColumns(1, fullHeader.length); } catch(_) {}
}

// ── TOURNAMENT STATS SHEET ───────────────────────────────────
function writeTournamentSheet(currentTournament) {
  const ss    = SpreadsheetApp.openById(SS_ID);
  let sheet   = ss.getSheetByName(SHEET_TOURNEY);
  if (!sheet) sheet = ss.insertSheet(SHEET_TOURNEY);
  sheet.clearContents();

  const games   = (currentTournament && currentTournament.games) ? currentTournament.games : [];
  const players = [...new Set(games.flatMap(g => g.players || []))];

  if (players.length === 0) {
    sheet.appendRow(['No tournament data yet.']);
    return;
  }

  const stats = calcStatsGAS(games);
  writeStatsToSheet(sheet, players, stats, `Tournament: ${(currentTournament && currentTournament.id) || '—'}`);
}

// ── LEADERBOARD STATS SHEET ──────────────────────────────────
function writeLeaderboardSheet(allTime) {
  const ss    = SpreadsheetApp.openById(SS_ID);
  let sheet   = ss.getSheetByName(SHEET_LEADER);
  if (!sheet) sheet = ss.insertSheet(SHEET_LEADER);
  sheet.clearContents();

  const games   = (allTime && allTime.games) ? allTime.games : [];
  const players = [...new Set(games.flatMap(g => g.players || []))];

  if (players.length === 0) {
    sheet.appendRow(['No all-time data yet.']);
    return;
  }

  const stats = calcStatsGAS(games);
  writeStatsToSheet(sheet, players, stats, 'All-Time Leaderboard');
}

// ── WRITE STATS TABLE TO SHEET ───────────────────────────────
function writeStatsToSheet(sheet, players, stats, title) {
  if (!players || players.length === 0) return;

  const STAT_ROWS = [
    { section:'POINTS DISTRIBUTION' },
    { key:'totalTournamentPoints', label:'Total Tournament Points' },
    { key:'tp5',  label:'Games Earning 5 T-Points' },
    { key:'tp4',  label:'Games Earning 4 T-Points' },
    { key:'tp3',  label:'Games Earning 3 T-Points' },
    { key:'tp2',  label:'Games Earning 2 T-Points' },
    { key:'tp1',  label:'Games Earning 1 T-Point'  },
    { section:'FINANCIALS' },
    { key:'moneyLosses',     label:'Money From Losses',        fmt:'$' },
    { key:'moneyPenalties',  label:'Money From Penalties',     fmt:'$' },
    { key:'totalPot',        label:'Total Money in Pot',       fmt:'$' },
    { key:'mostMoneyOneGame',label:'Most Money Paid One Game', fmt:'$' },
    { section:'GENERAL SCORING' },
    { key:'avgGamePoints',  label:'Average Game Points' },
    { key:'totalGamePoints',label:'Total Game Points'   },
    { key:'totalSets',      label:'Total Number of Sets'   },
    { key:'totalTricks',    label:'Total Number of Tricks' },
    { section:'GAME RECORDS' },
    { key:'mostSetsOneGame',   label:'Most Sets in One Game'    },
    { key:'leastSetsOneGame',  label:'Least Sets in One Game'   },
    { key:'mostTricksOneGame', label:'Most Tricks in One Game'  },
    { key:'leastTricksOneGame',label:'Least Tricks in One Game' },
    { key:'lowestScore',       label:'Lowest Score Ever'        },
    { key:'highestScore',      label:'Highest Score Ever'       },
    { section:'STREAKS' },
    { key:'longestWinStreakGames',  label:'Longest Winning Streak (Games)'   },
    { key:'longestLoseStreakGames', label:'Longest Losing Streak (Games)'    },
    { key:'longestWinStreakHands',  label:'Longest Winning Streak (Hands)'   },
    { key:'longestLoseStreakHands', label:'Longest Losing Streak (Hands)'    },
    { key:'longestStreakNoPay',     label:'Longest Streak Without Paying'    },
    { key:'longestStreakPay',       label:'Longest Streak With Paying'       },
  ];

  const numCols = players.length + 1; // stat label + one col per player

  // ── Row 1: Title ──────────────────────────────────────────
  sheet.appendRow([title]);
  sheet.getRange(1, 1, 1, numCols)
       .merge()
       .setFontWeight('bold')
       .setFontSize(13)
       .setBackground('#2c3e50')
       .setFontColor('#ffffff')
       .setHorizontalAlignment('left');

  // ── Row 2: Column headers ─────────────────────────────────
  const headerRow = ['Statistic', ...players];
  sheet.appendRow(headerRow);
  sheet.getRange(2, 1, 1, numCols)
       .setFontWeight('bold')
       .setBackground('#3d5166')
       .setFontColor('#ffffff');

  // ── Rows 3+: Stats ────────────────────────────────────────
  let sheetRow = 3; // track actual sheet row (1-indexed)

  STAT_ROWS.forEach(row => {
    if (row.section) {
      // Section divider row
      const sectionLabel = [row.section];
      // Pad to full width so merge works cleanly
      for (let i = 1; i < numCols; i++) sectionLabel.push('');
      sheet.appendRow(sectionLabel);
      sheet.getRange(sheetRow, 1, 1, numCols)
           .setFontWeight('bold')
           .setFontColor('#ffffff')
           .setBackground('#4a6070')
           .setFontSize(9);
    } else {
      // Data row
      const cells = [row.label];
      players.forEach(p => {
        const s   = stats[p];
        let val   = (s && s[row.key] !== undefined && s[row.key] !== null) ? s[row.key] : '—';
        if (row.fmt === '$' && val !== '—') val = '$' + (+val).toFixed(2);
        cells.push(val);
      });
      sheet.appendRow(cells);

      // Right-align numeric columns
      if (players.length > 0) {
        sheet.getRange(sheetRow, 2, 1, players.length).setHorizontalAlignment('center');
      }
      // Money rows: green font
      if (row.fmt === '$') {
        sheet.getRange(sheetRow, 2, 1, players.length).setFontColor('#27ae60');
      }
    }
    sheetRow++;
  });

  // ── Final formatting ──────────────────────────────────────
  try { sheet.setFrozenRows(2); } catch(_) {}
  try { sheet.autoResizeColumns(1, numCols); } catch(_) {}

  // Alternate row shading on data rows (skip title + header)
  const dataStart = 3;
  const dataEnd   = sheetRow - 1;
  for (let r = dataStart; r <= dataEnd; r++) {
    // Only shade non-section rows (section rows are already colored)
    const bg = sheet.getRange(r, 1).getBackground();
    if (bg === '#4a6070' || bg === '#4a6070'.toLowerCase()) continue; // skip section rows
    if (r % 2 === 0) {
      sheet.getRange(r, 1, 1, numCols).setBackground('#f4f7f6');
    }
  }
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
      if      (tp === 5) s.tp5++;
      else if (tp === 4) s.tp4++;
      else if (tp === 3) s.tp3++;
      else if (tp === 2) s.tp2++;
      else if (tp === 1) s.tp1++;

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

      // ── Per-round stats ──
      let gameSets = 0, gameTricks = 0;
      (rounds || []).forEach(rd => {
        const prd = (rd.scores || []).find(x => x.name === name);
        if (!prd) return;
        const made   = (Number(prd.bid) === Number(prd.tricks));
        const tricks = Number(prd.tricks) || 0;
        if (!made) gameSets++;
        gameTricks += tricks;
        if (made) { pg.hw++; pg.hl = 0; pg.bw = Math.max(pg.bw, pg.hw); }
        else      { pg.hl++; pg.hw = 0; pg.bl = Math.max(pg.bl, pg.hl); }
      });

      s.totalSets        += gameSets;
      s.totalTricks      += gameTricks;
      s.mostSetsOneGame    = Math.max(s.mostSetsOneGame,   gameSets);
      s.leastSetsOneGame   = Math.min(s.leastSetsOneGame,  gameSets);
      s.mostTricksOneGame  = Math.max(s.mostTricksOneGame,  gameTricks);
      s.leastTricksOneGame = Math.min(s.leastTricksOneGame, gameTricks);

      // ── Game win/lose streaks ──
      const rank = Object.values(tPoints || {}).filter(v => Number(v) > tp).length;
      if (rank === 0) {
        s._cw++; s._cl = 0;
        s.longestWinStreakGames  = Math.max(s.longestWinStreakGames,  s._cw);
      } else if (rank === players.length - 1) {
        s._cl++; s._cw = 0;
        s.longestLoseStreakGames = Math.max(s.longestLoseStreakGames, s._cl);
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
    s.avgGamePoints       = s._games ? +(s.totalGamePoints / s._games).toFixed(1) : 0;
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
