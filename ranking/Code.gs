// 구글 시트 > 확장 프로그램 > Apps Script 에 붙여넣고 "웹 앱"으로 배포하세요.
// 시트 탭(rps, sudoku)은 첫 기록이 들어올 때 자동으로 만들어집니다.
// 기록을 초기화하려면 해당 탭을 지우면 됩니다.

const GAMES = {
  rps:    { higherBetter: true,  maxAttempts: 3, min: 0, max: 30 },     // 10라운드 승점
  sudoku: { higherBetter: false, maxAttempts: 0, min: 1, max: 100000 }, // 시간(초), 0 = 도전 무제한
};
const MAX_NAME = 12;
const COL = { name: 0, score: 1, level: 2, attempts: 3 };

function doGet(e) {
  const p = e.parameter;
  const cfg = GAMES[p.game];
  let result;
  if (!cfg) result = { ok: false, error: "unknown game" };
  else if (p.action === "me") result = me(p);
  else if (p.action === "start") result = withLock(() => start(p, cfg));
  else if (p.action === "add") result = withLock(() => add(p, cfg));
  else result = top(p, cfg);
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function withLock(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function sheetFor(game) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(game);
  if (!sh) {
    sh = ss.insertSheet(game);
    sh.appendRow(["이름", "최고 기록", "난이도", "도전 횟수", "마지막 기록 시각"]);
  }
  return sh;
}

function cleanName(p) {
  return String(p.name || "").trim().replace(/^[=+\-@']+/, "").slice(0, MAX_NAME);
}

// 이름 + 난이도로 행 찾기 → { sh, row(1부터, 없으면 0), values }
function findRow(p) {
  const sh = sheetFor(p.game);
  const name = cleanName(p), level = String(p.level || "");
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][COL.name]) === name && String(rows[i][COL.level]) === level) {
      return { sh, row: i + 1, values: rows[i] };
    }
  }
  return { sh, row: 0, values: null };
}

function me(p) {
  const { values } = findRow(p);
  return { ok: true, attempts: values ? Number(values[COL.attempts]) : 0,
           best: values && values[COL.score] !== "" ? values[COL.score] : null };
}

// 게임 시작 시 도전 횟수 차감
function start(p, cfg) {
  const name = cleanName(p);
  if (!name) return { ok: false, error: "no name" };
  const { sh, row, values } = findRow(p);
  const attempts = values ? Number(values[COL.attempts]) : 0;
  if (cfg.maxAttempts && attempts >= cfg.maxAttempts) return { ok: false, error: "no attempts", attempts };
  if (row) sh.getRange(row, COL.attempts + 1).setValue(attempts + 1);
  else sh.appendRow([name, "", String(p.level || ""), 1, ""]);
  return { ok: true, attempts: attempts + 1 };
}

// 게임 끝나면 기록 저장 (최고 기록만 남김)
function add(p, cfg) {
  const score = Number(p.score);
  if (!isFinite(score) || score < cfg.min || score > cfg.max) return { ok: false, error: "bad score" };
  const { sh, row, values } = findRow(p);
  if (!row) return { ok: false, error: "not started" };
  const old = values[COL.score];
  const better = old === "" || (cfg.higherBetter ? score > old : score < old);
  if (better) sh.getRange(row, COL.score + 1).setValue(score);
  sh.getRange(row, 5).setValue(new Date());
  return { ok: true, best: better ? score : old };
}

function top(p, cfg) {
  const level = String(p.level || "");
  const rows = sheetFor(p.game).getDataRange().getValues().slice(1)
    .filter(r => r[COL.score] !== "" && String(r[COL.level]) === level)
    .sort((a, b) => cfg.higherBetter ? b[COL.score] - a[COL.score] : a[COL.score] - b[COL.score])
    .slice(0, 10)
    .map(r => ({ name: String(r[COL.name]), score: r[COL.score], attempts: r[COL.attempts] }));
  return { ok: true, top: rows };
}
