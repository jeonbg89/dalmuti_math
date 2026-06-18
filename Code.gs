/***********************************************************************
 *  달무티 수학 · 백엔드 (Google Apps Script + 스프레드시트)
 *  갈담초 5학년 · 학교 구글 계정 안에서 운영 (추가 비용 0)
 *
 *  - 학생 시트(명단) + 학습기록 시트(세트 1건당 1행)
 *  - 포인트/성실도/진행률/주간성장/달무티 계급을 '학습기록'에서 계산
 *  - 프론트엔드(index.html)와 계급·주간 로직이 동일합니다.
 *  설치/배포는 '배포가이드.md' 참고.
 *  ※ 프로젝트 시간대를 Asia/Seoul로 두세요(프로젝트 설정).
 ***********************************************************************/

const CONFIG = {
  TZ: "Asia/Seoul",
  TOKEN: "",                  // 보안용. 임의 문자열로 바꾸면 활성화(예: "galdam2026")
  SHEET_STUDENTS: "학생",
  SHEET_LOG: "학습기록",
  TOTAL_STAGES: 61,
  LEVELS: [
    [0, "첫걸음"], [300, "오름 Ⅰ"], [800, "오름 Ⅱ"], [1600, "징검다리"],
    [2800, "외나무다리"], [4500, "돌다리"], [7000, "구름다리"], [10000, "연산 마스터"],
  ],
  // 달무티 한국어판 13계급 (위 = 최고) — 프론트엔드와 동일
  DAL: ["달무티", "대주교", "시종장", "남작 부인", "백작", "기사", "재봉사", "석공", "요리사", "양치기", "광부", "농노", "어릿광대"],
  P_CORRECT: 1, P_SET: 5, P_PERFECT: 10, P_ATTEND: 20, P_STREAK7: 50, P_STAGE: 30,
  CLEAR_PCT: 80,   // 정확도 80% 이상이면 단계 완료(다음 단계 잠금 해제)
  // ▼ 전부 클리어(기사) 이후 누적 점수로 백작→남작 부인→시종장→대주교→달무티. 임시값 — 한 주 데이터 보고 숫자만 조정하세요.
  TOP_POINTS: [3000, 4500, 6500, 9000, 12000],
};

/* ─────────────────────────── 라우팅 ─────────────────────────── */
function doGet(e) {
  try {
    if (!checkToken(e)) return json({ ok: false, error: "unauthorized" });
    const action = (e.parameter.action || "class");
    if (action === "roster") return json({ ok: true, students: readRoster() });
    if (action === "dashboard") return json({ ok: true, data: buildDashboard(Number(e.parameter.no)) });
    if (action === "progress") return json({ ok: true, data: buildProgress(Number(e.parameter.no)) });
    if (action === "class") return json({ ok: true, data: buildClass() });
    return json({ ok: false, error: "unknown action" });
  } catch (err) { return json({ ok: false, error: String(err) }); }
}
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || "{}");
    if (!checkToken({ parameter: { token: body.token } })) return json({ ok: false, error: "unauthorized" });
    if (body.action === "log") return json({ ok: true, data: logSet(body) });
    return json({ ok: false, error: "unknown action" });
  } catch (err) { return json({ ok: false, error: String(err) }); }
}
function checkToken(e) { if (!CONFIG.TOKEN) return true; return (e.parameter && e.parameter.token) === CONFIG.TOKEN; }
function json(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }

/* ─────────────────────── 시트 I/O ─────────────────────── */
function ss() { return SpreadsheetApp.getActiveSpreadsheet(); }
function readRoster() {
  const sh = ss().getSheetByName(CONFIG.SHEET_STUDENTS);
  const rows = sh.getDataRange().getValues().slice(1);
  return rows.filter(r => r[0] !== "" && (r[3] === "" || r[3] === true || r[3] === "Y"))
    .map(r => ({ no: Number(r[0]), name: String(r[1]), email: String(r[2] || "") }));
}
function readRecords() {
  const sh = ss().getSheetByName(CONFIG.SHEET_LOG);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 10).getValues().filter(r => r[0] !== "").map(r => ({
    ts: r[0] instanceof Date ? r[0] : new Date(r[0]),
    no: Number(r[1]), name: String(r[2]), stageId: String(r[3]), stageName: String(r[4]),
    correct: Number(r[5]), total: Number(r[6]), perfect: (r[7] === "Y" || r[7] === true),
    seconds: Number(r[8]), setPoints: Number(r[9]) || 0,
  }));
}

/* ─────────────────────── 공통 헬퍼 ─────────────────────── */
const ymd = (d) => Utilities.formatDate(d, CONFIG.TZ, "yyyy-MM-dd");
const dayNum = (s) => { const p = s.split("-"); return Date.UTC(+p[0], +p[1] - 1, +p[2]) / 86400000; };
function isoWeek(d) {
  const t = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate())); const day = t.getUTCDay() || 7;
  t.setUTCDate(t.getUTCDate() + 4 - day); const ys = new Date(Date.UTC(t.getUTCFullYear(), 0, 1));
  const wk = Math.ceil(((t - ys) / 86400000 + 1) / 7); return t.getUTCFullYear() + "-W" + String(wk).padStart(2, "0");
}
function levelOf(p) {
  let i = 0; for (let k = 0; k < CONFIG.LEVELS.length; k++) if (p >= CONFIG.LEVELS[k][0]) i = k;
  const next = CONFIG.LEVELS[Math.min(i + 1, CONFIG.LEVELS.length - 1)][0];
  return { idx: i, name: CONFIG.LEVELS[i][1], cur: CONFIG.LEVELS[i][0], next, isMax: i === CONFIG.LEVELS.length - 1 };
}
function rankTitle(pos, total) {
  if (total <= 1) return CONFIG.DAL[0];
  const idx = Math.round(pos / (total - 1) * (CONFIG.DAL.length - 1));
  return CONFIG.DAL[Math.max(0, Math.min(CONFIG.DAL.length - 1, idx))];
}
function tierOf(title) {
  const i = CONFIG.DAL.indexOf(title);
  if (i === 0) return "최고 귀족";
  if (i <= 4) return "귀족";
  if (i <= 6) return "기사·장인";
  if (i <= 10) return "장인·노동";
  if (i === 11) return "농노";
  return "어릿광대";
}
/* 절대(협동) 계급: 남과 비교가 아니라 '내가 깬 단계 + 누적 점수'로만 결정
   - 아래 7계급(어릿광대→재봉사)은 클리어한 단계 수(월드 경계)로
   - 전부 클리어(61) = 기사, 그 위 5계급은 누적 점수(TOP_POINTS)로  */
function absRank(cleared, points) {
  const T = CONFIG.TOTAL_STAGES;
  if (cleared < T) {
    const byClear = [[54, "재봉사"], [47, "석공"], [34, "요리사"], [19, "양치기"], [9, "광부"], [1, "농노"], [0, "어릿광대"]];
    for (let i = 0; i < byClear.length; i++) if (cleared >= byClear[i][0]) return byClear[i][1];
    return "어릿광대";
  }
  const P = CONFIG.TOP_POINTS; // [백작, 남작 부인, 시종장, 대주교, 달무티] 오름차순
  if (points >= P[4]) return "달무티";
  if (points >= P[3]) return "대주교";
  if (points >= P[2]) return "시종장";
  if (points >= P[1]) return "남작 부인";
  if (points >= P[0]) return "백작";
  return "기사";
}

/* ─────────────────────── 집계 엔진 ─────────────────────── */
function computeStudentStats(recs, todayStr) {
  recs = recs.slice().sort((a, b) => a.ts - b.ts);
  const base = { sets: 0, problems: 0, correct: 0, perfectSets: 0, perfectRate: 0, accuracy: 0,
    days: 0, c30: 0, dilig: 0, streak: 0, stages: 0, clearedStages: 0, points: 0, level: levelOf(0), recent: [], sets30: 0, prob30: 0 };
  if (!recs.length) return base;
  let problems = 0, correct = 0, perfectSets = 0;
  const dateSet = {}, perfectStages = {}, bestByStage = {};
  recs.forEach(r => { problems += r.total; correct += r.correct;
    const acc = r.total ? Math.round(r.correct / r.total * 100) : 0;
    if (acc > (bestByStage[r.stageId] || 0)) bestByStage[r.stageId] = acc;
    if (r.perfect) { perfectSets++; perfectStages[r.stageId] = true; } dateSet[ymd(r.ts)] = true; });
  const dates = Object.keys(dateSet).sort(), dayNums = dates.map(dayNum);
  const numSet = {}; dayNums.forEach(n => numSet[n] = true);
  const todayN = dayNum(todayStr);
  const days = dates.length;
  const c30 = dayNums.filter(n => todayN - n < 30).length;
  const dilig = Math.round(c30 / 30 * 100);
  const recs30 = recs.filter(r => todayN - dayNum(ymd(r.ts)) < 30);
  const sets30 = recs30.length, prob30 = recs30.reduce((a, r) => a + r.total, 0);
  let streak = 0, cur = todayN; if (!numSet[cur]) cur -= 1; while (numSet[cur]) { streak++; cur -= 1; }
  let maxStreak = 1, run = 1;
  for (let i = 1; i < dayNums.length; i++) { run = (dayNums[i] - dayNums[i - 1] === 1) ? run + 1 : 1; if (run > maxStreak) maxStreak = run; }
  const sets = recs.length, stages = Object.keys(perfectStages).length;
  const clearedStages = Object.keys(bestByStage).filter(id => bestByStage[id] >= CONFIG.CLEAR_PCT).length;
  const perfectRate = Math.round(perfectSets / sets * 100), accuracy = Math.round(correct / problems * 100);
  const points = correct * CONFIG.P_CORRECT + sets * CONFIG.P_SET + perfectSets * CONFIG.P_PERFECT
    + days * CONFIG.P_ATTEND + Math.floor(maxStreak / 7) * CONFIG.P_STREAK7 + stages * CONFIG.P_STAGE;
  const recent = recs.slice(-5).reverse().map(r => ({ stageName: r.stageName, correct: r.correct, total: r.total,
    perfect: r.perfect, points: r.setPoints, date: ymd(r.ts) }));
  return { sets, problems, correct, perfectSets, perfectRate, accuracy, days, c30, dilig,
    streak, maxStreak, stages, clearedStages, points, level: levelOf(points), recent, sets30, prob30 };
}

/* 주간 성장(이번 주 setPoints 합) → 달무티 계급 순위 */
function weeklyGrowth(recs, weekStr) {
  return recs.filter(r => isoWeek(r.ts) === weekStr).reduce((a, r) => a + (r.setPoints || 0), 0);
}
function rankClass(roster, allRecs, weekStr) {
  const arr = roster.map(s => ({ no: s.no, name: s.name,
    weekly: weeklyGrowth(allRecs.filter(r => r.no === s.no), weekStr) }))
    .sort((a, b) => (b.weekly - a.weekly) || (a.no - b.no));
  arr.forEach((x, i) => { x.pos = i; x.title = rankTitle(i, arr.length); x.tier = tierOf(x.title); });
  const byNo = {}; arr.forEach(x => byNo[x.no] = x);
  return { ranking: arr, byNo };
}

/* ─────────────────────── 응답 빌더 ─────────────────────── */
function buildDashboard(no) {
  const roster = readRoster(), all = readRecords(), week = isoWeek(new Date());
  const me = roster.find(s => s.no === no) || { no, name: "?" };
  const mine = all.filter(r => r.no === no);
  const stats = computeStudentStats(mine, ymd(new Date()));
  const title = absRank(stats.clearedStages, stats.points), tier = tierOf(title);
  return Object.assign({ no: me.no, name: me.name, totalStages: CONFIG.TOTAL_STAGES, week,
    weekly: weeklyGrowth(mine, week), title: title, tier: tier, classSize: roster.length }, stats);
}
function buildClass() {
  const roster = readRoster(), all = readRecords(), week = isoWeek(new Date()), today = ymd(new Date());
  const ranking = rankClass(roster, all, week).ranking;
  const students = ranking.map(r => {
    const st = computeStudentStats(all.filter(x => x.no === r.no), today);
    return { no: r.no, name: r.name, weekly: r.weekly, title: r.title, tier: r.tier, pos: r.pos,
      points: st.points, level: st.level.name, perfectRate: st.perfectRate, accuracy: st.accuracy,
      dilig: st.dilig, streak: st.streak, studiedToday: st.recent.some(x => x.date === today) };
  });
  const n = students.length || 1;
  const summary = { week, dalmuti: students[0] ? students[0].name : "-",
    weeklyTotal: students.reduce((a, s) => a + s.weekly, 0),
    studiedToday: students.filter(s => s.studiedToday).length,
    avgPerfect: Math.round(students.reduce((a, s) => a + s.perfectRate, 0) / n), count: n };
  return { summary, students };
}

/* ─── 듀오링고식 단계 완료/왕관(메달) — 로그 기반, 단계별 '최고 정확도'만 ─── */
function medalOf(acc) { if (acc >= 100) return "gold"; if (acc >= 85) return "silver"; if (acc >= 70) return "bronze"; return "none"; }
function buildProgress(no) {
  const recs = readRecords().filter(r => r.no === no);
  const byStage = {};
  recs.forEach(r => {
    const acc = r.total ? Math.round(r.correct / r.total * 100) : 0;
    const cur = byStage[r.stageId] || { best: 0, plays: 0 };
    if (acc > cur.best) cur.best = acc;
    cur.plays += 1; // 그 단계를 푼 세트 수 = 반복 횟수
    byStage[r.stageId] = cur;
  });
  Object.keys(byStage).forEach(id => {
    const e = byStage[id];
    byStage[id] = { best: e.best, medal: medalOf(e.best), cleared: e.best >= CONFIG.CLEAR_PCT, plays: e.plays };
  });
  return byStage; // ← 맵을 그대로 반환 (프론트 mergeProgress 형식과 일치)
}

/* ─────────────────────── 쓰기: 세트 기록 ─────────────────────── */
function logSet(p) {
  const lock = LockService.getScriptLock(); lock.waitLock(8000);
  try {
    const sh = ss().getSheetByName(CONFIG.SHEET_LOG);
    const correct = Number(p.correct), total = Number(p.total), perfect = correct === total;
    const setPoints = correct * CONFIG.P_CORRECT + CONFIG.P_SET + (perfect ? CONFIG.P_PERFECT : 0);
    const name = p.name || (readRoster().find(s => s.no === Number(p.no)) || {}).name || "";
    sh.appendRow([new Date(), Number(p.no), name, String(p.stageId), String(p.stageName),
      correct, total, perfect ? "Y" : "N", Number(p.seconds || 0), setPoints]);
    SpreadsheetApp.flush();
    const all = readRecords(), week = isoWeek(new Date());
    const stats = computeStudentStats(all.filter(r => r.no === Number(p.no)), ymd(new Date()));
    const r = rankClass(readRoster(), all, week).byNo[Number(p.no)] || {};
    return { setPoints, perfect, stats, weekly: r.weekly, title: r.title, tier: r.tier, pos: r.pos };
  } finally { lock.releaseLock(); }
}

/* ─────────────────────── 최초 설치 ───────────────────────
   편집기에서 setup 함수를 한 번 실행 → 시트 2개 생성 + 1~18번 자리 */
function setup() {
  const s = ss();
  let stu = s.getSheetByName(CONFIG.SHEET_STUDENTS) || s.insertSheet(CONFIG.SHEET_STUDENTS);
  stu.clear();
  stu.getRange(1, 1, 1, 4).setValues([["번호", "이름", "이메일(구글계정)", "활성"]]).setFontWeight("bold");
  const seed = []; for (let i = 1; i <= 18; i++) seed.push([i, "학생" + i, "", "Y"]);
  stu.getRange(2, 1, seed.length, 4).setValues(seed); stu.setFrozenRows(1);
  let log = s.getSheetByName(CONFIG.SHEET_LOG) || s.insertSheet(CONFIG.SHEET_LOG);
  log.clear();
  log.getRange(1, 1, 1, 10).setValues([["시각", "학생번호", "이름", "단계ID", "단계명",
    "맞은개수", "총문제", "완전통과", "걸린초", "세트포인트"]]).setFontWeight("bold");
  log.setFrozenRows(1);
}
