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
      return json_({ ok: true, subs: group_(readRows_(), function (d) { return d === p.date; }) });
    }
    if (p.action === "range") {
      return json_({ ok: true, subs: group_(readRows_(), function (d) { return d >= p.from && d <= p.to; }) });
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
