/**
 * ICC Tennis DOUBLES Ladders — Google Apps Script
 *
 * Group-of-4 rotating-partners format. Each round, 4 players form a
 * group and play 3 sets (first to 6 games), switching partners each set:
 *     Set 1:  A & B   vs   C & D
 *     Set 2:  A & C   vs   B & D
 *     Set 3:  A & D   vs   B & C
 *
 * Every INDIVIDUAL accumulates their own Sets W/L and Games W/L.
 * Standings rank by Sets Win % first, then Games Win %.
 *
 * ── Setup ──────────────────────────────────────────────────
 *   1. Create a new Google Sheet named "Doubles Ladders Worksheet".
 *   2. Open Extensions > Apps Script, delete any existing code,
 *      and paste this whole file in. Save.
 *   3. Run the `setupLaddersTab` function once from the editor
 *      (click Run). Approve permissions when prompted.
 *   4. Deploy > New deployment > Web app
 *        - Execute as: Me
 *        - Who has access: Anyone
 *   5. Copy the Web App URL into index.html (APPS_SCRIPT_URL).
 *
 * ── Ladders config tab ─────────────────────────────────────
 *     Ladder Name | Active | Rounds | Players
 *     2025 Gold   | TRUE   | 12     | Ann, Beth, Carol, ...
 *   Players = comma-separated. Active = TRUE to show in the app.
 *
 * ── Match tab schema (auto-created) ────────────────────────
 *   "<Ladder Name> - Matches" columns:
 *     Date | Round | Player A | Player B | Player C | Player D |
 *     Set 1 | Set 2 | Set 3 | Submitted At
 *
 *   Each Set cell is stored as "t1-t2" where t1/t2 are the games
 *   for the two teams in that set's rotation pairing.
 */

// ═══════════════════════════════════════════════════════════
//  SETUP — run this once from the Apps Script editor
// ═══════════════════════════════════════════════════════════
function setupLaddersTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Ladders');

  if (!sheet) {
    sheet = ss.insertSheet('Ladders');
  }

  sheet.getRange(1, 1, 1, 4)
    .setValues([['Ladder Name', 'Active', 'Rounds', 'Players']])
    .setFontWeight('bold')
    .setBackground('#052d54')
    .setFontColor('#ffffff');

  if (sheet.getLastRow() < 2) {
    sheet.getRange(2, 1, 1, 4).setValues([[
      '2025 Ladies Gold Doubles',
      true,
      12,
      'Player One, Player Two, Player Three, Player Four'
    ]]);
  }

  sheet.setColumnWidth(1, 240);
  sheet.setColumnWidth(2, 80);
  sheet.setColumnWidth(3, 80);
  sheet.setColumnWidth(4, 500);
  sheet.setFrozenRows(1);

  SpreadsheetApp.getUi().alert(
    'Ladders tab ready!\n\n' +
    'Edit row 2 with your real ladder name, rounds, and the full ' +
    'comma-separated player list. Add more rows for more ladders.'
  );
}

// ═══════════════════════════════════════════════════════════
//  HELPERS
// ═══════════════════════════════════════════════════════════
function getLaddersConfig() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Ladders');
  if (!sheet || sheet.getLastRow() < 2) return [];

  var lastCol = Math.max(4, sheet.getLastColumn());
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues();
  var ladders = [];

  for (var i = 0; i < rows.length; i++) {
    var name    = String(rows[i][0] || '').trim();
    var active  = rows[i][1] === true || String(rows[i][1]).toUpperCase() === 'TRUE';
    var rounds  = Number(rows[i][2]) || 12;
    var playersRaw = String(rows[i][3] || '').trim();

    if (!name) continue;

    var players = playersRaw
      .split(',')
      .map(function(p) { return p.trim(); })
      .filter(function(p) { return p.length > 0; });

    ladders.push({
      name: name,
      active: active,
      rounds: rounds,
      format: 'doubles-rotation',
      players: players
    });
  }

  return ladders;
}

function getLadderByName(name) {
  var ladders = getLaddersConfig();
  for (var i = 0; i < ladders.length; i++) {
    if (ladders[i].name === name) return ladders[i];
  }
  return null;
}

function getMatchesSheet(ladderName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tabName = ladderName + ' - Matches';
  var sheet = ss.getSheetByName(tabName);

  if (!sheet) {
    sheet = ss.insertSheet(tabName);
    sheet.getRange(1, 1, 1, 10)
      .setValues([[
        'Date', 'Round',
        'Player A', 'Player B', 'Player C', 'Player D',
        'Set 1', 'Set 2', 'Set 3', 'Submitted At'
      ]])
      .setFontWeight('bold')
      .setBackground('#052d54')
      .setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 100);
    sheet.setColumnWidth(2, 60);
    sheet.setColumnWidth(3, 140);
    sheet.setColumnWidth(4, 140);
    sheet.setColumnWidth(5, 140);
    sheet.setColumnWidth(6, 140);
    sheet.setColumnWidth(7, 70);
    sheet.setColumnWidth(8, 70);
    sheet.setColumnWidth(9, 70);
    sheet.setColumnWidth(10, 180);
  }

  return sheet;
}

// The rotation: for each set index, which player slots team up.
// Players are [A, B, C, D] => indices [0, 1, 2, 3].
function rotationForSet(setIndex) {
  if (setIndex === 0) return { team1: [0, 1], team2: [2, 3] };
  if (setIndex === 1) return { team1: [0, 2], team2: [1, 3] };
  return { team1: [0, 3], team2: [1, 2] };
}

// ═══════════════════════════════════════════════════════════
//  POST — log a group's round (4 players, 3 sets)
// ═══════════════════════════════════════════════════════════
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    if (!data.ladder) throw new Error('Missing ladder name.');
    var ladder = getLadderByName(data.ladder);
    if (!ladder) throw new Error('Ladder not found: ' + data.ladder);

    var players = Array.isArray(data.players) ? data.players : [];
    if (players.length !== 4) throw new Error('Need exactly 4 players.');

    var sets = Array.isArray(data.sets) ? data.sets : [];
    if (sets.length !== 3) throw new Error('Need exactly 3 sets.');

    var setStrings = [];
    for (var s = 0; s < 3; s++) {
      var t1 = Number(sets[s].t1);
      var t2 = Number(sets[s].t2);
      if (isNaN(t1) || isNaN(t2)) throw new Error('Set ' + (s + 1) + ' is incomplete.');
      setStrings.push(t1 + '-' + t2);
    }

    var sheet = getMatchesSheet(data.ladder);
    sheet.appendRow([
      data.date || '',
      data.round || '',
      players[0], players[1], players[2], players[3],
      setStrings[0], setStrings[1], setStrings[2],
      new Date().toISOString()
    ]);

    return jsonOut({ status: 'ok' });

  } catch (err) {
    return jsonOut({ status: 'error', message: err.toString() });
  }
}

// ═══════════════════════════════════════════════════════════
//  GET — list ladders OR return standings for a ladder
// ═══════════════════════════════════════════════════════════
function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) || 'ladders';

    if (action === 'ladders') {
      var all = getLaddersConfig();
      var active = all.filter(function(l) { return l.active; });
      return jsonOut({ status: 'ok', ladders: active });
    }

    if (action === 'standings') {
      var ladderName = e.parameter.ladder;
      if (!ladderName) throw new Error('Missing ladder parameter.');
      return jsonOut(calculateStandings(ladderName));
    }

    throw new Error('Unknown action: ' + action);

  } catch (err) {
    return jsonOut({ status: 'error', message: err.toString() });
  }
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ═══════════════════════════════════════════════════════════
//  STANDINGS CALCULATION
// ═══════════════════════════════════════════════════════════
function parseSet(str) {
  var parts = String(str).split('-');
  return { t1: Number(parts[0]) || 0, t2: Number(parts[1]) || 0 };
}

function calculateStandings(ladderName) {
  var ladder = getLadderByName(ladderName);
  if (!ladder) throw new Error('Ladder not found: ' + ladderName);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(ladderName + ' - Matches');

  // Initialize stats for every roster player
  var stats = {};
  for (var i = 0; i < ladder.players.length; i++) {
    stats[ladder.players[i]] = {
      name: ladder.players[i],
      rounds: 0,
      setsWon: 0, setsLost: 0,
      gamesWon: 0, gamesLost: 0
    };
  }

  var recent = [];

  if (sheet && sheet.getLastRow() > 1) {
    var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues();

    for (var r = 0; r < rows.length; r++) {
      var date    = rows[r][0];
      var round   = rows[r][1];
      var players  = [rows[r][2], rows[r][3], rows[r][4], rows[r][5]];
      var setCells = [rows[r][6], rows[r][7], rows[r][8]];

      // Distribute each set's result across the 4 players per the rotation
      for (var si = 0; si < 3; si++) {
        var sc = parseSet(setCells[si]);
        var rot = rotationForSet(si);
        var t1won = sc.t1 > sc.t2;

        rot.team1.forEach(function(idx) {
          var p = stats[players[idx]];
          if (!p) return; // skip subs / non-roster
          p.gamesWon  += sc.t1;
          p.gamesLost += sc.t2;
          if (t1won) p.setsWon += 1; else p.setsLost += 1;
        });
        rot.team2.forEach(function(idx) {
          var p = stats[players[idx]];
          if (!p) return;
          p.gamesWon  += sc.t2;
          p.gamesLost += sc.t1;
          if (!t1won) p.setsWon += 1; else p.setsLost += 1;
        });
      }

      // Count a round played for each roster player in this group
      players.forEach(function(name) {
        if (stats[name]) stats[name].rounds += 1;
      });

      recent.push({
        date:    date instanceof Date ? Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(date),
        round:   round,
        players: players,
        sets:    [String(setCells[0]), String(setCells[1]), String(setCells[2])]
      });
    }
  }

  // Compute percentages
  var standings = Object.keys(stats).map(function(k) {
    var p = stats[k];
    var totalSets  = p.setsWon + p.setsLost;
    var totalGames = p.gamesWon + p.gamesLost;
    p.setsPct  = totalSets  > 0 ? p.setsWon  / totalSets  : 0;
    p.gamesPct = totalGames > 0 ? p.gamesWon / totalGames : 0;
    return p;
  });

  // Sort: Sets Win % first, then Games Win %, then name
  standings.sort(function(a, b) {
    if (b.setsPct  !== a.setsPct)  return b.setsPct  - a.setsPct;
    if (b.gamesPct !== a.gamesPct) return b.gamesPct - a.gamesPct;
    return a.name.localeCompare(b.name);
  });

  // Matchups: groups of 4 by rank (1-4, 5-8, ...)
  var matchups = [];
  for (var m = 0; m < standings.length; m += 4) {
    var group = standings.slice(m, m + 4).map(function(p, idx) {
      return { rank: m + idx + 1, name: p.name };
    });
    matchups.push(group);
  }

  // Most recent 15 groups
  recent.reverse();
  var recentGroups = recent.slice(0, 15);

  return {
    status: 'ok',
    ladder: ladderName,
    format: ladder.format,
    rounds: ladder.rounds,
    standings: standings,
    matchups: matchups,
    recent: recentGroups
  };
}
