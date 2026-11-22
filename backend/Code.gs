/* AllHikes sign-ups: Google Apps Script behind the Telegram mini app.

   Lives in the "AllHikes sign-ups" Google Sheet (Extensions → Apps Script).
   Script property BOT_TOKEN = the bot's token from @BotFather (never put it in this file).
   Run setup() once. Then Deploy → Web app, "Execute as: Me", "Who has access: Anyone".

   The mini app POSTs {action, initData, hike?, on?}. initData is the signed proof from
   Telegram of who the user is; nothing is saved unless that signature checks out.
   A daily trigger at 18:00 Yerevan time sends a reminder for tomorrow's hikes.
   A person's Telegram account is their AllHikes account: every signed-in call also updates
   their row in the "Users" tab. "booked" is set by the person themselves in My hikes. */

const APP_URL = "https://apricot-stone.github.io/allhikes/";
const TZ = "Asia/Yerevan";
const SHEET = "Signups";
const HEAD = ["time", "hike_id", "hike", "date", "club", "user_id", "name", "username", "lang", "status", "reminded", "booked"];
const USERS = "Users";
const UHEAD = ["user_id", "name", "username", "lang", "first_seen", "last_seen"];
const C = Object.fromEntries(HEAD.map((h, i) => [h, i]));

/* ---------- web app ---------- */
function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ok: false, error: "bad request"}); }
  const user = verify_(body.initData);
  if (!user) return json_({ok: false, error: "not signed by Telegram"});
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    saveUser_(user);
    if (body.action === "going") setGoing_(user, body.hike || {}, !!body.on);
    if (body.action === "booked") setBooked_(user, (body.hike || {}).id, !!body.on);
    const rows = rows_();
    return json_({ok: true, mine: mine_(rows, user.id), counts: counts_(rows)});
  } catch (err) {
    return json_({ok: false, error: String(err.message || err)});
  } finally { lock.releaseLock(); }
}

function doGet() {
  return json_({ok: true, counts: counts_(rows_())});
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- Telegram login check (core.telegram.org/bots/webapps#validating-data-received-via-the-mini-app) ---------- */
function verify_(initData) {
  if (!initData || typeof initData !== "string") return null;
  const token = PropertiesService.getScriptProperties().getProperty("BOT_TOKEN");
  if (!token) throw new Error("BOT_TOKEN script property is missing");
  const f = {};
  initData.split("&").forEach(p => {
    const i = p.indexOf("=");
    if (i > 0) f[decodeURIComponent(p.slice(0, i))] = decodeURIComponent(p.slice(i + 1).replace(/\+/g, "%20"));
  });
  const hash = f.hash;
  if (!hash || !f.user) return null;
  delete f.hash;
  const check = Object.keys(f).sort().map(k => k + "=" + f[k]).join("\n");
  const secret = Utilities.computeHmacSha256Signature(token, "WebAppData");
  const sig = Utilities.computeHmacSha256Signature(Utilities.newBlob(check).getBytes(), secret);
  const hex = sig.map(b => ((b + 256) % 256).toString(16).padStart(2, "0")).join("");
  if (hex !== hash) return null;
  if (Date.now() / 1000 - Number(f.auth_date) > 7 * 86400) return null; // older than a week
  try { return JSON.parse(f.user); } catch (err) { return null; }
}

/* ---------- the sheet ---------- */
function sheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) { sh = ss.insertSheet(SHEET); sh.appendRow(HEAD); sh.setFrozenRows(1); }
  if (sh.getLastColumn() < HEAD.length) sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]); // new columns added later
  return sh;
}
function rows_() {
  return sheet_().getDataRange().getValues().slice(1).map(r => r.concat(Array(Math.max(0, HEAD.length - r.length)).fill("")));
}
const fullName_ = u => [u.first_name, u.last_name].filter(Boolean).join(" ").slice(0, 120);

function saveUser_(user) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(USERS);
  if (!sh) { sh = ss.insertSheet(USERS); sh.appendRow(UHEAD); sh.setFrozenRows(1); }
  const rows = sh.getDataRange().getValues();
  const i = rows.findIndex((r, n) => n > 0 && String(r[0]) === String(user.id));
  const now = new Date();
  const row = [String(user.id), fullName_(user), user.username ? "@" + user.username : "", (user.language_code || "").slice(0, 2)];
  if (i > 0) sh.getRange(i + 1, 1, 1, 6).setValues([row.concat([rows[i][4], now])]);
  else sh.appendRow(row.concat([now, now]));
}

function setBooked_(user, id, on) {
  const rows = rows_();
  const i = rows.findIndex(r => String(r[C.hike_id]) === String(id) && String(r[C.user_id]) === String(user.id) && r[C.status] === "going");
  if (i < 0) throw new Error("not signed up for this hike");
  sheet_().getRange(i + 2, C.booked + 1).setValue(on ? "yes" : "");
}
const today_ = (days) => Utilities.formatDate(new Date(Date.now() + (days || 0) * 864e5), TZ, "yyyy-MM-dd");
const day_ = v => v instanceof Date ? Utilities.formatDate(v, TZ, "yyyy-MM-dd") : String(v);

function setGoing_(user, hike, on) {
  const id = String(hike.id || "").slice(0, 80);
  if (!/^[\w-]+$/.test(id) || !/^\d{4}-\d{2}-\d{2}$/.test(hike.date || "")) throw new Error("bad hike");
  const sh = sheet_();
  const rows = rows_();
  const i = rows.findIndex(r => String(r[C.hike_id]) === id && String(r[C.user_id]) === String(user.id));
  const row = [
    new Date(), id, String(hike.name || "").slice(0, 200), "'" + hike.date, String(hike.club || "").slice(0, 120),
    String(user.id), fullName_(user),
    user.username ? "@" + user.username : "", (user.language_code || "en").slice(0, 2),
    on ? "going" : "cancelled", i >= 0 ? rows[i][C.reminded] : "",
    on && i >= 0 && rows[i][C.status] === "going" ? rows[i][C.booked] : "",
  ];
  if (i >= 0) sh.getRange(i + 2, 1, 1, HEAD.length).setValues([row]);
  else if (on) sh.appendRow(row);
}

function mine_(rows, userId) {
  return rows.filter(r => String(r[C.user_id]) === String(userId) && r[C.status] === "going").map(r => ({
    id: String(r[C.hike_id]), name: String(r[C.hike]), date: day_(r[C.date]), club: String(r[C.club]), booked: r[C.booked] === "yes",
  }));
}

function counts_(rows) {
  const t = today_(), out = {};
  rows.forEach(r => {
    if (r[C.status] === "going" && day_(r[C.date]) >= t) out[r[C.hike_id]] = (out[r[C.hike_id]] || 0) + 1;
  });
  return out;
}

/* ---------- reminders (daily trigger) ---------- */
const MSG = {
  en: (h, club, others, booked) => `⛰ Tomorrow: ${h}\nwith ${club}.\n\n${others ? `You and ${others} more from AllHikes are going. ` : ""}${booked ? "Have a great hike!" : "You haven't marked this hike as booked yet: message the club to make sure you have a place."}`,
  ru: (h, club, others, booked) => `⛰ Завтра: ${h}\nс ${club}.\n\n${others ? `Вместе с вами идут ещё ${others} человек из AllHikes. ` : ""}${booked ? "Хорошего похода!" : "Вы ещё не отметили, что записались: напишите клубу, чтобы точно было место."}`,
  hy: (h, club, others, booked) => `⛰ Վաղը՝ ${h}\n${club}-ի հետ։\n\n${others ? `Ձեզ հետ AllHikes-ից գալիս է ևս ${others} հոգի։ ` : ""}${booked ? "Լավ արշավ։" : "Դուք դեռ չեք նշել, որ ամրագրել եք. գրեք ակումբին՝ տեղ ունենալու համար։"}`,
};
const BTN = {en: "Open AllHikes", ru: "Открыть AllHikes", hy: "Բացել AllHikes"};

function sendReminders() {
  const sh = sheet_(), rows = rows_(), tomorrow = today_(1), sentOn = today_();
  rows.forEach((r, i) => {
    if (r[C.status] !== "going" || day_(r[C.date]) !== tomorrow || r[C.reminded]) return;
    const others = rows.filter(x => x[C.hike_id] === r[C.hike_id] && x[C.status] === "going").length - 1;
    const lang = MSG[r[C.lang]] ? r[C.lang] : "en";
    const ok = tg_("sendMessage", {
      chat_id: String(r[C.user_id]),
      text: MSG[lang](r[C.hike], r[C.club], others, r[C.booked] === "yes"),
      reply_markup: {inline_keyboard: [[{text: BTN[lang], web_app: {url: APP_URL}}]]},
    });
    sh.getRange(i + 2, C.reminded + 1).setValue(ok ? sentOn : "failed " + sentOn);
  });
}

function tg_(method, payload) {
  const token = PropertiesService.getScriptProperties().getProperty("BOT_TOKEN");
  const res = UrlFetchApp.fetch(`https://api.telegram.org/bot${token}/${method}`, {
    method: "post", contentType: "application/json", payload: JSON.stringify(payload), muteHttpExceptions: true,
  });
  try { return JSON.parse(res.getContentText()).ok === true; } catch (err) { return false; }
}

/* ---------- run once ---------- */
function setup() {
  sheet_();
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === "sendReminders").forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("sendReminders").timeBased().atHour(18).everyDays(1).inTimezone(TZ).create();
  const token = PropertiesService.getScriptProperties().getProperty("BOT_TOKEN");
  if (!token) throw new Error("Add the BOT_TOKEN script property first (Project Settings → Script properties).");
  const me = JSON.parse(UrlFetchApp.fetch(`https://api.telegram.org/bot${token}/getMe`, {muteHttpExceptions: true}).getContentText());
  Logger.log(me.ok ? `All set. Bot: @${me.result.username}. Reminders go out daily at 18:00 Yerevan time.` : "The BOT_TOKEN looks wrong: " + me.description);
}
