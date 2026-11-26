// ============================================================
//  Secret Ovens – कल का सामान : Google Sheet server
//  यह कोड Google Sheet के Apps Script में पेस्ट करें।
// ============================================================

// मालिक का PIN – किसी की जमा की हुई लिस्ट हटाने के लिए। इसे बदल लें।
const OWNER_PIN = "1234";

const SHEET_NAME = "Entries";
const HEADERS = ["तारीख (किसके लिए)", "फ़ोन", "नाम", "कब भेजा", "कैटेगरी", "सामान", "मात्रा", "यूनिट"];

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange("A:A").setNumberFormat("@");
  }
  return sh;
}

// ---------- खरीदार का हिसाब (रेट) ----------
const BUY_SHEET = "Purchases";
const BUY_HEADERS = ["तारीख", "किसने भरा", "कब सेव किया", "कैटेगरी", "सामान", "लिस्ट में मात्रा", "लाया मात्रा", "यूनिट", "रेट (1 यूनिट)", "कुल ₹"];
// ---------- लॉगिन कोड: इन्हें सीधे Sheet के "Codes" टैब में बदलें ----------
const CODE_SHEET = "Codes";

function getBuySheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(BUY_SHEET);
  if (!sh) {
    sh = ss.insertSheet(BUY_SHEET);
    sh.appendRow(BUY_HEADERS);
    sh.setFrozenRows(1);
    sh.getRange("A:A").setNumberFormat("@");
  }
  return sh;
}

function getCodes_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(CODE_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CODE_SHEET);
    sh.getRange(1, 1, 3, 2).setValues([["लॉगिन", "कोड"], ["खरीदार", "'2580"], ["मालिक", "'9999"]]);
    sh.setFrozenRows(1);
  }
  const v = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 2).getValues();
  const c = {};
  v.forEach(function (r) {
    const k = String(r[0]).trim(), code = String(r[1]).replace(/^'/, "").trim();
    if (k === "खरीदार") c.buy = code;
    if (k === "मालिक") c.own = code;
  });
  return c;
}

// मालिक का कोड हर जगह चलता है; खरीदार का कोड सिर्फ़ खरीदार वाले काम में
function checkCode_(role, code) {
  const c = getCodes_();
  code = String(code || "").trim();
  if (!code) return false;
  if (role === "own") return code === c.own;
  if (role === "buy") return code === c.buy || code === c.own;
  return false;
}

function readBuys_(keepDate) {
  const sh = getBuySheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 10).getValues()
    .filter(function (r) { return keepDate(dateStr_(r[0])); })
    .map(function (r) {
      return { d: dateStr_(r[0]), who: String(r[1]), c: String(r[3]), n: String(r[4]),
        cq: r[5] === "" ? null : Number(r[5]), q: Number(r[6]), u: String(r[7]), r: Number(r[8]) };
    });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function dateStr_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return String(v).replace(/^'/, "");
}

function readRows_() {
  const sh = getSheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 8).getValues();
}

// rows -> list of submissions {d, dev, who, at, items:[{c,n,q,u}]}
function group_(rows, keepDate) {
  const subs = {};
  const order = [];
  rows.forEach(function (r) {
    const d = dateStr_(r[0]);
    if (!keepDate(d)) return;
    const key = d + "|" + r[1] + "|" + r[2];
    let s = subs[key];
    if (!s) {
      s = subs[key] = {
        d: d,
        dev: String(r[1]),
        who: String(r[2]),
        at: r[3] instanceof Date ? r[3].getTime() : Number(r[3]) || 0,
        items: []
      };
      order.push(key);
    }
    const q = (r[6] === "" || r[6] === null) ? null : Number(r[6]);
    s.items.push({ c: String(r[4]), n: String(r[5]), q: q, u: String(r[7]) });
  });
  return order.map(function (k) { return subs[k]; });
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  try {
    if (p.action === "day") {
      const keepD = function (d) { return d === p.date; };
      return json_({ ok: true, subs: group_(readRows_(), keepD), buys: readBuys_(keepD) });
    }
    if (p.action === "range") {
      const keepR = function (d) { return d >= p.from && d <= p.to; };
      return json_({ ok: true, subs: group_(readRows_(), keepR), buys: readBuys_(keepR) });
    }
    return json_({ ok: true, app: "secret-ovens", time: Date.now() });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const b = JSON.parse(e.postData.contents);

    // लॉगिन: कोड Sheet के "Codes" टैब से जाँचा जाता है
    if (b.action === "login") {
      if (!checkCode_(b.role, b.code) || (b.role === "buy" && String(b.code).trim() !== getCodes_().buy)) return json_({ ok: false, error: "code" });
      const out = { ok: true };
      if (b.role === "own") out.sheet = SpreadsheetApp.getActiveSpreadsheet().getUrl();
      return json_(out);
    }

    // खरीदार का हिसाब: उस तारीख की पुरानी एंट्री हटाकर नई लिखो
    if (b.action === "buy") {
      if (!/^\d{4}-\d{2}-\d{2}$/.test(String(b.date))) return json_({ ok: false, error: "date" });
      if (!checkCode_("buy", b.code)) return json_({ ok: false, error: "code" });
      const bs = getBuySheet_();
      const last = bs.getLastRow();
      if (last >= 2) {
        const dates = bs.getRange(2, 1, last - 1, 1).getValues();
        for (let k = dates.length - 1; k >= 0; k--) {
          if (dateStr_(dates[k][0]) === b.date) bs.deleteRow(k + 2);
        }
      }
      const rows = (b.rows || []).filter(function (x) { return Number(x.q) > 0 && Number(x.r) > 0; });
      if (rows.length) {
        const now = new Date();
        const data = rows.map(function (x) {
          const q = Number(x.q), r = Number(x.r);
          return ["'" + b.date, String(b.who || ""), now, String(x.c), String(x.n),
            (x.cq === null || x.cq === undefined || x.cq === "") ? "" : Number(x.cq), q, String(x.u), r, Math.round(q * r * 100) / 100];
        });
        bs.getRange(bs.getLastRow() + 1, 1, data.length, 10).setValues(data);
      }
      return json_({ ok: true, saved: rows.length });
    }

    if (b.action !== "submit" && b.action !== "delete") return json_({ ok: false, error: "action" });
    if (!/^\d{4}-\d{2}-\d{2}$/.test(String(b.date))) return json_({ ok: false, error: "date" });
    if (b.action === "delete" && String(b.pin) !== OWNER_PIN) return json_({ ok: false, error: "pin" });

    const sh = getSheet_();
    const who = String(b.who || "");
    const dev = String(b.dev || "");

    // इसी फ़ोन + नाम की इसी तारीख की पुरानी लिस्ट हटाएं (ताकि दोबारा भेजने पर डबल न हो)
    const rows = readRows_();
    const del = [];
    rows.forEach(function (r, i) {
      if (dateStr_(r[0]) === b.date && String(r[1]) === dev && String(r[2]) === who) del.push(i + 2);
    });
    let i = del.length - 1;
    while (i >= 0) {
      const end = del[i];
      let start = end;
      i--;
      while (i >= 0 && del[i] === start - 1) { start = del[i]; i--; }
      sh.deleteRows(start, end - start + 1);
    }

    // नई लिस्ट जोड़ें
    if (b.action === "submit" && b.items && b.items.length) {
      const now = new Date();
      const data = b.items.map(function (x) {
        const q = (x.q === null || x.q === undefined || x.q === "") ? "" : Number(x.q);
        return ["'" + b.date, dev, who, now, String(x.c), String(x.n), q, String(x.u)];
      });
      sh.getRange(sh.getLastRow() + 1, 1, data.length, 8).setValues(data);
    }
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
