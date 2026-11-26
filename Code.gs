// ============================================================
//  Secret Ovens – रेस्टोरेंट सामान ऐप : Google Sheet server (v10)
//  यह कोड Google Sheet के Apps Script में पेस्ट करें, फिर
//  Deploy → Manage deployments → ✏️ → New version → Deploy
// ============================================================
//
//  Sheet के टैब (अपने-आप बनते हैं):
//   Entries   – शेफ़/हेल्पर की भेजी लिस्ट
//   Purchases – खरीदार का हिसाब (मात्रा, दाम, कुल)
//   Cash      – खरीदार को दिए पैसे, खर्च, बचे पैसे, बिल की फ़ोटो
//   Codes     – खरीदार और मालिक के कोड (यहीं बदलें)
//   Items     – सामान की लिस्ट (यहीं जोड़ें/हटाएँ)
//   Settings  – रेस्टोरेंट का नाम और टैगलाइन
// ============================================================

// पुराने ऐप (v8) के लिए – लिस्ट हटाने का PIN। नया ऐप मालिक का कोड इस्तेमाल करता है।
const OWNER_PIN = "1234";

const SHEET_NAME = "Entries";
const HEADERS = ["तारीख (किसके लिए)", "फ़ोन", "नाम", "कब भेजा", "कैटेगरी", "सामान", "मात्रा", "यूनिट"];
const BUY_SHEET = "Purchases";
const BUY_HEADERS = ["तारीख", "किसने भरा", "कब सेव किया", "कैटेगरी", "सामान", "लिस्ट में मात्रा", "लाया मात्रा", "यूनिट", "रेट (1 यूनिट)", "कुल ₹", "स्थिति"];
const CASH_SHEET = "Cash";
const CASH_HEADERS = ["तारीख", "खरीदार", "दिए पैसे ₹", "कुल खर्च ₹", "बचे पैसे ₹", "कब सेव किया", "बिल की फ़ोटो"];
const CODE_SHEET = "Codes";
const ITEM_SHEET = "Items";
const ITEM_HEADERS = ["कैटेगरी", "सामान", "इकाई", "दुकान", "चालू (हाँ/नहीं)"];
const SET_SHEET = "Settings";
const BILL_FOLDER = "Secret Ovens – बिल की फ़ोटो";

// ---------------- छोटे काम ----------------
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function sheet_(name, headers) {
  let sh = ss_().getSheetByName(name);
  if (!sh) {
    sh = ss_().insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
    sh.setFrozenRows(1);
    sh.getRange("A:A").setNumberFormat("@");
  } else if (sh.getLastColumn() < headers.length) {
    // नए कॉलम जोड़ें (पुराना डेटा वैसा ही रहेगा)
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
  }
  return sh;
}
function getSheet_() { return sheet_(SHEET_NAME, HEADERS); }

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function dateStr_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
  return String(v).replace(/^'/, "");
}
function okDate_(d) { return /^\d{4}-\d{2}-\d{2}$/.test(String(d)); }
function rows_(sh, cols) {
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, cols).getValues();
}
function deleteDateRows_(sh, date) {
  const last = sh.getLastRow();
  if (last < 2) return;
  const d = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let k = d.length - 1; k >= 0; k--) if (dateStr_(d[k][0]) === date) sh.deleteRow(k + 2);
}

// ---------------- कोड (लॉगिन) ----------------
function getCodes_() {
  let sh = ss_().getSheetByName(CODE_SHEET);
  if (!sh) {
    sh = ss_().insertSheet(CODE_SHEET);
    sh.getRange(1, 1, 3, 2).setValues([["लॉगिन", "कोड"], ["खरीदार", "'2580"], ["मालिक", "'9999"]]);
    sh.getRange(1, 1, 1, 2).setFontWeight("bold");
    sh.setFrozenRows(1);
  }
  const c = {};
  rows_(sh, 2).forEach(function (r) {
    const k = String(r[0]).trim(), code = String(r[1]).replace(/^'/, "").trim();
    if (k === "खरीदार") c.buy = code;
    if (k === "मालिक") c.own = code;
  });
  return c;
}
// मालिक का कोड हर जगह चलता है; खरीदार का कोड सिर्फ़ खरीदार वाले काम में
function can_(role, code) {
  const c = getCodes_();
  code = String(code || "").trim();
  if (!code) return false;
  if (role === "own") return code === c.own;
  if (role === "buy") return code === c.buy || code === c.own;
  return false;
}
// गलत कोड: 5 बार के बाद 10 मिनट का ताला
function locked_(role) {
  return Number(CacheService.getScriptCache().get("fail_" + role) || 0) >= 5;
}
function fail_(role) {
  const cache = CacheService.getScriptCache(), k = "fail_" + role;
  cache.put(k, String(Number(cache.get(k) || 0) + 1), 600);
}

// ---------------- पढ़ना ----------------
// rows -> list of submissions {d, dev, who, at, items:[{c,n,q,u}]}
function group_(rows, keep) {
  const subs = {}, order = [];
  rows.forEach(function (r) {
    const d = dateStr_(r[0]);
    if (!keep(d)) return;
    const key = d + "|" + r[1] + "|" + r[2];
    let s = subs[key];
    if (!s) {
      s = subs[key] = { d: d, dev: String(r[1]), who: String(r[2]), at: r[3] instanceof Date ? r[3].getTime() : Number(r[3]) || 0, items: [] };
      order.push(key);
    }
    const q = (r[6] === "" || r[6] === null) ? null : Number(r[6]);
    s.items.push({ c: String(r[4]), n: String(r[5]), q: q, u: String(r[7]) });
  });
  return order.map(function (k) { return subs[k]; });
}
function readSubs_(keep) { return group_(rows_(getSheet_(), 8), keep); }

function readBuys_(keep) {
  return rows_(sheet_(BUY_SHEET, BUY_HEADERS), 11)
    .filter(function (r) { return keep(dateStr_(r[0])); })
    .map(function (r) {
      return { d: dateStr_(r[0]), who: String(r[1]), c: String(r[3]), n: String(r[4]),
        cq: r[5] === "" ? null : Number(r[5]), q: Number(r[6]) || 0, u: String(r[7]), r: Number(r[8]) || 0,
        na: String(r[10]) === "नहीं मिला" };
    });
}
function readCash_(keep) {
  return rows_(sheet_(CASH_SHEET, CASH_HEADERS), 7)
    .filter(function (r) { return keep(dateStr_(r[0])); })
    .map(function (r) {
      return { d: dateStr_(r[0]), who: String(r[1]), given: Number(r[2]) || 0, spent: Number(r[3]) || 0, left: Number(r[4]) || 0, photo: String(r[6] || "") };
    });
}
// हर सामान का पिछला दाम (इस तारीख से पहले का सबसे नया)
function lastRates_(before) {
  const best = {};
  readBuys_(function (d) { return d < before; }).forEach(function (b) {
    if (b.na || !(b.r > 0)) return;
    const k = b.n;
    if (!best[k] || best[k].d < b.d) best[k] = { r: b.r, d: b.d, u: b.u };
  });
  return best;
}

function readItems_() {
  const sh = sheet_(ITEM_SHEET, ITEM_HEADERS);
  return rows_(sh, 5)
    .filter(function (r) { return String(r[1]).trim() && String(r[4]).trim() !== "नहीं"; })
    .map(function (r) { return { c: String(r[0]).trim(), n: String(r[1]).trim(), u: String(r[2]).trim(), shop: String(r[3]).trim() }; });
}
function readSettings_() {
  let sh = ss_().getSheetByName(SET_SHEET);
  if (!sh) {
    sh = ss_().insertSheet(SET_SHEET);
    sh.getRange(1, 1, 3, 2).setValues([["सेटिंग", "वैल्यू"], ["रेस्टोरेंट का नाम", "SECRET OVENS"], ["टैगलाइन", "GOOD FOOD, GOOD MOOD"]]);
    sh.getRange(1, 1, 1, 2).setFontWeight("bold");
    sh.setFrozenRows(1);
  }
  const out = {};
  rows_(sh, 2).forEach(function (r) {
    const k = String(r[0]).trim(), v = String(r[1]).trim();
    if (k === "रेस्टोरेंट का नाम") out.name = v;
    if (k === "टैगलाइन") out.tagline = v;
  });
  return out;
}

// ---------------- GET (बिना कोड – सिर्फ़ लिस्ट, दाम नहीं) ----------------
function doGet(e) {
  const p = (e && e.parameter) || {};
  try {
    if (p.action === "day") return json_({ ok: true, subs: readSubs_(function (d) { return d === p.date; }) });
    if (p.action === "range") return json_({ ok: true, subs: readSubs_(function (d) { return d >= p.from && d <= p.to; }) });
    if (p.action === "items") return json_({ ok: true, items: readItems_(), settings: readSettings_() });
    return json_({ ok: true, app: "secret-ovens", v: 10, time: Date.now() });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

// ---------------- POST ----------------
function doPost(e) {
  let b;
  try { b = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: "bad" }); }

  // पढ़ने वाले काम – लॉक की ज़रूरत नहीं
  try {
    if (b.action === "login") {
      if (b.role !== "buy" && b.role !== "own") return json_({ ok: false, error: "code" });
      if (locked_(b.role)) return json_({ ok: false, error: "locked" });
      const c = getCodes_(), code = String(b.code || "").trim();
      if (!code || code !== c[b.role]) { fail_(b.role); return json_({ ok: false, error: "code" }); }
      const out = { ok: true };
      if (b.role === "own") out.sheet = ss_().getUrl();
      return json_(out);
    }
    if (b.action === "hisab") {
      if (!okDate_(b.date)) return json_({ ok: false, error: "date" });
      if (!can_("buy", b.code)) return json_({ ok: false, error: "code" });
      const keep = function (d) { return d === b.date; };
      const cash = readCash_(keep);
      return json_({ ok: true, subs: readSubs_(keep), buys: readBuys_(keep), cash: cash[0] || null, last: lastRates_(b.date) });
    }
    if (b.action === "report") {
      if (!can_("own", b.code)) return json_({ ok: false, error: "code" });
      const keep = function (d) { return d >= b.from && d <= b.to; };
      return json_({ ok: true, subs: readSubs_(keep), buys: readBuys_(keep), cash: readCash_(keep) });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }

  // लिखने वाले काम – एक समय में एक
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    // शेफ़/हेल्पर की लिस्ट भेजना या हटाना (पुराना और नया ऐप दोनों)
    if (b.action === "submit" || b.action === "delete") {
      if (!okDate_(b.date)) return json_({ ok: false, error: "date" });
      if (b.action === "delete" && String(b.pin) !== OWNER_PIN && !can_("own", b.pin)) return json_({ ok: false, error: "pin" });
      const sh = getSheet_();
      const who = String(b.who || ""), dev = String(b.dev || "");
      const rows = rows_(sh, 8), del = [];
      rows.forEach(function (r, i) { if (dateStr_(r[0]) === b.date && String(r[1]) === dev && String(r[2]) === who) del.push(i + 2); });
      for (let i = del.length - 1; i >= 0; i--) sh.deleteRow(del[i]);
      if (b.action === "submit" && b.items && b.items.length) {
        const now = new Date();
        const data = b.items.map(function (x) {
          const q = (x.q === null || x.q === undefined || x.q === "") ? "" : Number(x.q);
          return ["'" + b.date, dev, who, now, String(x.c), String(x.n), q, String(x.u)];
        });
        sh.getRange(sh.getLastRow() + 1, 1, data.length, 8).setValues(data);
      }
      return json_({ ok: true });
    }

    // खरीदार का हिसाब: उस तारीख का पुराना हटाकर नया लिखो
    if (b.action === "buy") {
      if (!okDate_(b.date)) return json_({ ok: false, error: "date" });
      if (!can_("buy", b.code)) return json_({ ok: false, error: "code" });
      const bs = sheet_(BUY_SHEET, BUY_HEADERS);
      deleteDateRows_(bs, b.date);
      const now = new Date();
      let spent = 0;
      const data = (b.rows || []).filter(function (x) { return x.na || (Number(x.q) > 0 && Number(x.r) > 0); }).map(function (x) {
        const q = x.na ? 0 : Number(x.q), r = x.na ? 0 : Number(x.r), t = Math.round(q * r * 100) / 100;
        spent += t;
        return ["'" + b.date, String(b.who || ""), now, String(x.c), String(x.n),
          (x.cq === null || x.cq === undefined || x.cq === "") ? "" : Number(x.cq), q, String(x.u), r, t, x.na ? "नहीं मिला" : "मिला"];
      });
      if (data.length) bs.getRange(bs.getLastRow() + 1, 1, data.length, 11).setValues(data);

      // पैसे का हिसाब (फ़ोटो का लिंक पहले जैसा रहे)
      const cs = sheet_(CASH_SHEET, CASH_HEADERS);
      const old = readCash_(function (d) { return d === b.date; })[0];
      deleteDateRows_(cs, b.date);
      const given = Number(b.given) || 0;
      spent = Math.round(spent * 100) / 100;
      cs.appendRow(["'" + b.date, String(b.who || ""), given, spent, Math.round((given - spent) * 100) / 100, now, old ? old.photo : ""]);
      return json_({ ok: true, saved: data.length, spent: spent });
    }

    // बिल की फ़ोटो → Google Drive
    if (b.action === "photo") {
      if (!okDate_(b.date)) return json_({ ok: false, error: "date" });
      if (!can_("buy", b.code)) return json_({ ok: false, error: "code" });
      const m = String(b.data || "").match(/^data:image\/(jpeg|png);base64,(.+)$/);
      if (!m) return json_({ ok: false, error: "photo" });
      const bytes = Utilities.base64Decode(m[2]);
      if (bytes.length > 3 * 1024 * 1024) return json_({ ok: false, error: "photo" });
      const it = DriveApp.getFoldersByName(BILL_FOLDER);
      const folder = it.hasNext() ? it.next() : DriveApp.createFolder(BILL_FOLDER);
      const file = folder.createFile(Utilities.newBlob(bytes, "image/" + m[1], "bill-" + b.date + "-" + Date.now() + "." + (m[1] === "png" ? "png" : "jpg")));
      const url = file.getUrl();
      const cs = sheet_(CASH_SHEET, CASH_HEADERS);
      const rows = rows_(cs, 7);
      let row = -1;
      rows.forEach(function (r, i) { if (dateStr_(r[0]) === b.date) row = i + 2; });
      if (row < 0) { cs.appendRow(["'" + b.date, String(b.who || ""), 0, 0, 0, new Date(), url]); }
      else {
        const prev = String(cs.getRange(row, 7).getValue() || "");
        cs.getRange(row, 7).setValue(prev ? prev + "\n" + url : url);
      }
      return json_({ ok: true, url: url });
    }

    // पहली बार: ऐप की लिस्ट से Items टैब भरना (सिर्फ़ मालिक, सिर्फ़ जब खाली हो)
    if (b.action === "seedItems") {
      if (!can_("own", b.code)) return json_({ ok: false, error: "code" });
      const sh = sheet_(ITEM_SHEET, ITEM_HEADERS);
      if (sh.getLastRow() >= 2) return json_({ ok: true, skipped: true });
      const data = (b.items || []).slice(0, 2000).map(function (x) { return [String(x.c), String(x.n), String(x.u), String(x.shop || ""), "हाँ"]; });
      if (data.length) sh.getRange(2, 1, data.length, 5).setValues(data);
      return json_({ ok: true, seeded: data.length });
    }

    return json_({ ok: false, error: "action" });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
