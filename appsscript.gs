// ============================================================
// VTI SIMULATION — Google Apps Script Web App
// SETUP INSTRUCTIONS:
// 1. In your Google Sheet, go to Extensions > Apps Script
// 2. Delete any existing code, paste this entire file
// 3. Click Deploy > New Deployment
// 4. Type: Web App
// 5. Execute as: Me
// 6. Who has access: Anyone
// 7. Click Deploy, copy the Web App URL
// 8. Paste that URL into index.html where it says PASTE_YOUR_APPS_SCRIPT_URL_HERE
// ============================================================

const SHEET_ID = '1lrNaXhHcjLlPCEbyI8hcg5tj_TF5UgVvqEZmh-ZKVfE';

function doGet(e) {
  const p = e.parameter;
  let result;
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    switch (p.action) {
      case 'checkName':   result = checkName(ss, p.name); break;
      case 'register':    result = registerTeam(ss, p.name); break;
      case 'saveAnswer':  result = saveAnswer(ss, p.team, p.round, p.answer, p.points, p.time); break;
      case 'complete':    result = completeGame(ss, p.team, p.score, p.rounds); break;
      case 'rankings':    result = getRankings(ss); break;
      case 'teamAnswers': result = getTeamAnswers(ss, p.team); break;
      default:            result = { error: 'Unknown action' };
    }
  } catch (err) {
    result = { error: err.toString() };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#0B1F3A').setFontColor('#ffffff');
  }
  return sheet;
}

function checkName(ss, name) {
  const sheet = getOrCreateSheet(ss, 'Teams', ['TeamName','RegisteredAt','CurrentRound','TotalScore','Status','RoundsCompleted']);
  const data = sheet.getDataRange().getValues();
  const taken = data.slice(1).some(row => row[0].toString().toLowerCase().trim() === name.toLowerCase().trim());
  return { available: !taken };
}

function registerTeam(ss, name) {
  const lock = LockService.getScriptLock();
  lock.tryLock(6000);
  try {
    const check = checkName(ss, name);
    if (!check.available) return { success: false, reason: 'name_taken' };
    const sheet = getOrCreateSheet(ss, 'Teams', ['TeamName','RegisteredAt','CurrentRound','TotalScore','Status','RoundsCompleted']);
    sheet.appendRow([name, new Date().toISOString(), 0, 0, 'playing', 0]);
    return { success: true };
  } finally {
    lock.releaseLock();
  }
}

function saveAnswer(ss, team, round, answer, points, time) {
  const sheet = getOrCreateSheet(ss, 'Answers', ['TeamName','Round','AnswerGiven','PointsEarned','TimeTaken','Timestamp']);
  sheet.appendRow([team, parseInt(round), answer, parseInt(points), parseInt(time), new Date().toISOString()]);
  // Update team current round and score
  const teamsSheet = ss.getSheetByName('Teams');
  if (teamsSheet) {
    const data = teamsSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === team) {
        teamsSheet.getRange(i + 1, 3).setValue(parseInt(round));
        teamsSheet.getRange(i + 1, 4).setValue((data[i][3] || 0) + parseInt(points));
        teamsSheet.getRange(i + 1, 6).setValue(parseInt(round));
        break;
      }
    }
  }
  return { success: true };
}

function completeGame(ss, team, score, rounds) {
  const teamsSheet = ss.getSheetByName('Teams');
  if (teamsSheet) {
    const data = teamsSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === team) {
        teamsSheet.getRange(i + 1, 3).setValue(parseInt(rounds));
        teamsSheet.getRange(i + 1, 4).setValue(parseInt(score));
        teamsSheet.getRange(i + 1, 5).setValue('complete');
        teamsSheet.getRange(i + 1, 6).setValue(parseInt(rounds));
        break;
      }
    }
  }
  return { success: true };
}

function getRankings(ss) {
  const teamsSheet = ss.getSheetByName('Teams');
  if (!teamsSheet) return { teams: [] };
  const data = teamsSheet.getDataRange().getValues();
  const teams = data.slice(1).map(row => ({
    name: row[0], registeredAt: row[1], currentRound: row[2],
    totalScore: row[3], status: row[4], roundsCompleted: row[5]
  }));
  teams.sort((a, b) => b.totalScore - a.totalScore || b.roundsCompleted - a.roundsCompleted);
  return { teams };
}

function getTeamAnswers(ss, team) {
  const sheet = ss.getSheetByName('Answers');
  if (!sheet) return { answers: [] };
  const data = sheet.getDataRange().getValues();
  const answers = data.slice(1)
    .filter(row => row[0] === team)
    .map(row => ({ round: row[1], answer: row[2], points: row[3], time: row[4] }));
  return { answers };
}
