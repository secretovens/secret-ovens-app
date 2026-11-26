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
const ITEM_HEADERS = ["कैटेगरी", "सामान", "इकाई", "दुकान", "चालू (हाँ/नहीं)", "स्टेप (+ दबाने पर कितना बढ़े)", "फ़ोटो (ऐप से लगती है)"];
const SET_SHEET = "Settings";
const SALES_SHEET = "Sales";
const SALES_HEADERS = ["तारीख", "काउंटर (नकद) ₹", "UPI / कार्ड ₹", "Zomato ₹", "Swiggy ₹", "कुल बिक्री ₹", "किसने भरा", "कब सेव किया"];
const BILL_FOLDER = "Secret Ovens – बिल की फ़ोटो";
const ITEM_PHOTO_FOLDER = "Secret Ovens – सामान की फ़ोटो";

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
  return rows_(sh, 7)
    .filter(function (r) { return String(r[1]).trim() && String(r[4]).trim() !== "नहीं"; })
    .map(function (r) { return { c: String(r[0]).trim(), n: String(r[1]).trim(), u: String(r[2]).trim(), shop: String(r[3]).trim(), step: Number(r[5]) || "", img: String(r[6] || "").trim() }; });
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
    if (k === "सामान लिस्ट वर्ज़न") out.listVersion = Number(v) || 0;
  });
  return out;
}

function readSales_(keep) {
  return rows_(sheet_(SALES_SHEET, SALES_HEADERS), 8)
    .filter(function (r) { return keep(dateStr_(r[0])); })
    .map(function (r) { return { d: dateStr_(r[0]), cash: Number(r[1]) || 0, upi: Number(r[2]) || 0, zomato: Number(r[3]) || 0, swiggy: Number(r[4]) || 0, total: Number(r[5]) || 0 }; });
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
    // हफ़्ते/महीने का हिसाब Excel में (सिर्फ़ मालिक) – मालिक के काम की 5 शीट
    if (b.action === "export") {
      if (!can_("own", b.code)) return json_({ ok: false, error: "code" });
      if (!okDate_(b.from) || !okDate_(b.to)) return json_({ ok: false, error: "date" });
      return json_(exportXls_(b.from, b.to));
    }
    if (b.action === "report") {
      if (!can_("own", b.code)) return json_({ ok: false, error: "code" });
      const keep = function (d) { return d >= b.from && d <= b.to; };
      return json_({ ok: true, subs: readSubs_(keep), buys: readBuys_(keep), cash: readCash_(keep), sales: readSales_(keep) });
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

    // रोज़ की बिक्री (सिर्फ़ मालिक): उस तारीख की पुरानी हटाकर नई
    if (b.action === "sale") {
      if (!okDate_(b.date)) return json_({ ok: false, error: "date" });
      if (!can_("own", b.code)) return json_({ ok: false, error: "code" });
      const sh = sheet_(SALES_SHEET, SALES_HEADERS);
      deleteDateRows_(sh, b.date);
      const v = [Number(b.cash) || 0, Number(b.upi) || 0, Number(b.zomato) || 0, Number(b.swiggy) || 0];
      const tot = v[0] + v[1] + v[2] + v[3];
      if (tot > 0) sh.appendRow(["'" + b.date].concat(v).concat([tot, String(b.who || ""), new Date()]));
      return json_({ ok: true, total: tot });
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
        // ग्राम वाले सामान का दाम 1 kg का होता है
        const q = x.na ? 0 : Number(x.q), r = x.na ? 0 : Number(x.r), t = Math.round((String(x.u) === "ग्राम" ? q * r / 1000 : q * r) * 100) / 100;
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
      cs.appendRow(["'" + b.date, String(b.who || ""), given > 0 ? given : "", spent, given > 0 ? Math.round((given - spent) * 100) / 100 : "", now, old ? old.photo : ""]);
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

    // सामान की फ़ोटो (सिर्फ़ मालिक): Drive में रखकर Items टैब में उसका id
    if (b.action === "itemPhoto") {
      if (!can_("own", b.code)) return json_({ ok: false, error: "code" });
      const m = String(b.data || "").match(/^data:image\/(jpeg|png);base64,(.+)$/);
      if (!m) return json_({ ok: false, error: "photo" });
      const bytes = Utilities.base64Decode(m[2]);
      if (bytes.length > 1024 * 1024) return json_({ ok: false, error: "photo" });
      const sh = sheet_(ITEM_SHEET, ITEM_HEADERS);
      const rows = rows_(sh, 7);
      let row = -1;
      rows.forEach(function (r, i) { if (String(r[1]).trim() === String(b.name).trim()) row = i + 2; });
      if (row < 0) return json_({ ok: false, error: "item" });
      const it = DriveApp.getFoldersByName(ITEM_PHOTO_FOLDER);
      const folder = it.hasNext() ? it.next() : DriveApp.createFolder(ITEM_PHOTO_FOLDER);
      const file = folder.createFile(Utilities.newBlob(bytes, "image/" + m[1], String(b.name) + ".jpg"));
      // स्टाफ़ के फ़ोन पर दिखे इसलिए "लिंक वाला कोई भी देख सके" (सिर्फ़ ये सामान वाली फ़ोटो)
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      const prev = String(rows[row - 2][6] || "");
      if (prev) { try { DriveApp.getFileById(prev).setTrashed(true); } catch (e) {} }
      sh.getRange(row, 7).setValue(file.getId());
      return json_({ ok: true, img: file.getId() });
    }

    // Items टैब भरना: खाली हो, या ऐप में नई लिस्ट (ऊँचा वर्ज़न) हो। दुकान और चालू/बंद पहले जैसे रहते हैं।
    if (b.action === "seedItems") {
      if (!can_("own", b.code)) return json_({ ok: false, error: "code" });
      const sh = sheet_(ITEM_SHEET, ITEM_HEADERS);
      const ver = Number(b.version) || 0, cur = readSettings_().listVersion || 0;
      if (sh.getLastRow() >= 2 && ver <= cur) return json_({ ok: true, skipped: true });
      const old = {};
      rows_(sh, 7).forEach(function (r) { old[String(r[1]).trim()] = { shop: r[3], on: r[4], img: r[6] }; });
      if (sh.getLastRow() >= 2) sh.getRange(2, 1, sh.getLastRow() - 1, 7).clearContent();
      const data = (b.items || []).slice(0, 2000).map(function (x) {
        const o = old[String(x.n).trim()] || {};
        return [String(x.c), String(x.n), String(x.u), String(o.shop || x.shop || ""), String(o.on || "हाँ"), x.step ? Number(x.step) : "", String(o.img || "")];
      });
      if (data.length) sh.getRange(2, 1, data.length, 7).setValues(data);
      if (ver) {
        const st = ss_().getSheetByName(SET_SHEET);
        const rr = rows_(st, 2);
        let row = -1;
        rr.forEach(function (r, i) { if (String(r[0]).trim() === "सामान लिस्ट वर्ज़न") row = i + 2; });
        if (row < 0) st.appendRow(["सामान लिस्ट वर्ज़न", ver]); else st.getRange(row, 2).setValue(ver);
      }
      return json_({ ok: true, seeded: data.length });
    }

    return json_({ ok: false, error: "action" });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// ======================= Excel रिपोर्ट =======================
function exportXls_(from, to) {
  const tz = Session.getScriptTimeZone();
  const D = function (s) { const p = s.split("-"); return new Date(+p[0], +p[1] - 1, +p[2]); };
  const S = function (d) { return Utilities.formatDate(d, tz, "yyyy-MM-dd"); };
  const nDays = Math.round((D(to) - D(from)) / 864e5) + 1;
  const pTo = S(new Date(D(from).getTime() - 864e5)), pFrom = S(new Date(D(from).getTime() - nDays * 864e5));
  const amt = function (x) { return String(x.u) === "ग्राम" ? x.q * x.r / 1000 : x.q * x.r; };
  const ok = function (x) { return !x.na && x.q > 0 && x.r > 0; };
  const special = function (c) { return c === "एकमुश्त" || c === "अन्य खर्च" || c === "स्टाफ़ सब्ज़ी"; };
  const cur = readBuys_(function (d) { return d >= from && d <= to; }).filter(ok);
  const prev = readBuys_(function (d) { return d >= pFrom && d <= pTo; }).filter(ok);
  const shopOf = {}; readItems_().forEach(function (i) { shopOf[i.n] = i.shop; });
  const r2 = function (v) { return Math.round(v * 100) / 100; };
  const unitRate = function (u) { return u === "ग्राम" ? "1 kg" : "1 " + u; };

  // सामान-वार जोड़
  const agg = function (list) {
    const m = {};
    list.forEach(function (x) {
      const k = x.n + "|" + x.u;
      const e = m[k] || (m[k] = { c: x.c, n: x.n, u: x.u, q: 0, t: 0, days: {}, min: 1e12, max: 0, minD: "" });
      e.q += x.q; e.t += amt(x); e.days[x.d] = 1;
      if (x.r < e.min) { e.min = x.r; e.minD = x.d; }
      if (x.r > e.max) e.max = x.r;
    });
    Object.keys(m).forEach(function (k) { const e = m[k]; e.avg = e.q ? (e.u === "ग्राम" ? e.t * 1000 / e.q : e.t / e.q) : 0; });
    return m;
  };
  const A = agg(cur), P = agg(prev);
  const items = Object.keys(A).map(function (k) { return A[k]; }).filter(function (e) { return !special(e.c); }).sort(function (a, b) { return b.t - a.t; });
  const specials = Object.keys(A).map(function (k) { return A[k]; }).filter(function (e) { return special(e.c); });
  const total = cur.reduce(function (s, x) { return s + amt(x); }, 0);
  const pTotal = prev.reduce(function (s, x) { return s + amt(x); }, 0);
  const buyDays = {}; cur.forEach(function (x) { buyDays[x.d] = (buyDays[x.d] || 0) + amt(x); });
  const saleDays = {}; let saleTotal = 0;
  readSales_(function (d) { return d >= from && d <= to; }).forEach(function (x) { saleDays[x.d] = x.total; saleTotal += x.total; });

  const xs = SpreadsheetApp.create("Secret Ovens हिसाब " + from + " से " + to);
  const G = "#1f4d3a", Y = "#fff2cc";
  const head = function (sh, r, vals) { sh.getRange(r, 1, 1, vals.length).setValues([vals]).setFontWeight("bold").setBackground(G).setFontColor("#ffffff").setWrap(true); };

  // ---------- 1. सारांश ----------
  const s1 = xs.getSheets()[0]; s1.setName("सारांश");
  s1.getRange("A1").setValue("Secret Ovens – इन्वेंटरी का हिसाब").setFontSize(16).setFontWeight("bold").setFontColor(G);
  s1.getRange("A2").setValue(from + " से " + to + " (" + nDays + " दिन)   ·   पिछली अवधि: " + pFrom + " से " + pTo).setFontColor("#666666");
  const rows = [
    ["कुल खरीद (₹)", r2(total)],
    ["रोज़ की औसत खरीद (₹)", r2(total / nDays)],
    ["पिछली अवधि की खरीद (₹)", r2(pTotal)],
    ["पिछली अवधि से बदलाव", pTotal ? "=(B4-B6)/B6" : "—"],
    ["", ""],
    ["इस अवधि की कुल बिक्री (₹)", saleTotal ? "='रोज़ का खर्च'!D" + (nDays + 2) : ""],
    ["बचत % (सामान के बाद) = (बिक्री − खरीद) ÷ बिक्री", "=IF(B9>0,(B9-B4)/B9,\"बिक्री भरें\")"],
    ["हाल", "=IF(B9>0,IF(B10<0,\"🔴 घाटा\",IF(B10>0.8,\"सारी खरीद भरी है? जाँचें\",IF(B10>=0.68,\"✅ अच्छी (68% या ज़्यादा)\",IF(B10>=0.65,\"⚠ ध्यान दें (65–68%)\",\"❌ कम है – सामान महँगा या दाम कम\")))),\"\")"],
    ["बचत (₹) = बिक्री − खरीद", "=IF(B9>0,B9-B4,\"\")"],
    ["", ""],
    ["एकमुश्त / अन्य खर्च / स्टाफ़ सब्ज़ी (₹)", r2(specials.reduce(function (s, e) { return s + e.t; }, 0))]
  ];
  s1.getRange(4, 1, rows.length, 2).setValues(rows);
  s1.getRange("A4:A14").setFontWeight("bold");
  s1.getRange("B4:B6").setNumberFormat("₹#,##0"); s1.getRange("B12").setNumberFormat("₹#,##0"); s1.getRange("B14").setNumberFormat("₹#,##0");
  s1.getRange("B7").setNumberFormat("+0%;-0%;0%"); s1.getRange("B10").setNumberFormat("0.0%");
  s1.getRange("B9").setBackground(Y).setNumberFormat("₹#,##0").setNote("ऐप में भरी रोज़ की बिक्री का जोड़ ('रोज़ का खर्च' शीट से)। खाली हो तो यहाँ खुद भी लिख सकते हैं।");
  s1.getRange("A16").setValue("भारत के रेस्टोरेंट में सामान का खर्च बिक्री का 28–32% ठीक माना जाता है, यानी बचत 68% या ज़्यादा (स्रोत: Petpooja / NRAI)। किराया, तनख़्वाह, गैस इसमें नहीं घटे हैं।").setFontColor("#666666");
  head(s1, 18, ["सबसे ज़्यादा खर्च वाले 10 सामान", "कुल ₹", "हिस्सा %"]);
  const top = items.slice(0, 10).map(function (e) { return [e.n, r2(e.t), total ? e.t / total : 0]; });
  if (top.length) { s1.getRange(19, 1, top.length, 3).setValues(top); s1.getRange(19, 2, top.length, 1).setNumberFormat("₹#,##0"); s1.getRange(19, 3, top.length, 1).setNumberFormat("0%"); }
  s1.setColumnWidth(1, 330); s1.setColumnWidth(2, 140); s1.setColumnWidth(3, 90);

  // ---------- 2. सामान-वार (ABC + कितना लाना है) ----------
  const s2 = xs.insertSheet("सामान-वार (ABC)");
  head(s2, 1, ["सामान", "कैटेगरी", "दुकान", "कितना आया", "इकाई", "कुल ₹", "हिस्सा %", "जोड़ हिस्सा %", "ABC", "कितने दिन खरीदा", "रोज़ की औसत खपत", "हफ़्ते की ज़रूरत (अंदाज़ा)", "औसत दाम ₹", "दाम किसका", "पिछली अवधि औसत ₹", "दाम में बदलाव"]);
  const tCell = "$F$" + (items.length + 2);
  const r2rows = items.map(function (e, i) {
    const r = i + 2, pe = P[e.n + "|" + e.u];
    return [e.n, e.c, shopOf[e.n] || "", r2(e.q), e.u, r2(e.t), "=F" + r + "/" + tCell, "=SUM($G$2:G" + r + ")",
      "=IF(H" + r + "<=0.7,\"A\",IF(H" + r + "<=0.9,\"B\",\"C\"))", Object.keys(e.days).length, "=D" + r + "/" + nDays, "=K" + r + "*7",
      r2(e.avg), unitRate(e.u), pe ? r2(pe.avg) : "", pe ? "=IF(O" + r + ">0,M" + r + "/O" + r + "-1,\"\")" : ""];
  });
  if (r2rows.length) {
    s2.getRange(2, 1, r2rows.length, 16).setValues(r2rows);
    s2.getRange(2, 6, r2rows.length, 1).setNumberFormat("₹#,##0");
    s2.getRange(2, 7, r2rows.length, 2).setNumberFormat("0%");
    s2.getRange(2, 11, r2rows.length, 2).setNumberFormat("0.0");
    s2.getRange(2, 13, r2rows.length, 1).setNumberFormat("₹#,##0.0"); s2.getRange(2, 15, r2rows.length, 1).setNumberFormat("₹#,##0.0");
    s2.getRange(2, 16, r2rows.length, 1).setNumberFormat("+0%;-0%;0%");
  }
  const tr = r2rows.length + 2;
  s2.getRange(tr, 1, 1, 6).setValues([["कुल", "", "", "", "", r2rows.length ? "=SUM(F2:F" + (tr - 1) + ")" : 0]]).setFontWeight("bold");
  s2.getRange(tr, 6).setNumberFormat("₹#,##0");
  s2.getRange(tr + 2, 1).setValue("ABC: A = सबसे ज़्यादा पैसा (कुल का 70%) – इन्हीं पर दाम और मात्रा का ध्यान रखें · B = अगला 20% · C = बाकी छोटा सामान। 'हफ़्ते की ज़रूरत' = रोज़ की औसत खपत × 7, यानी अगले हफ़्ते के लिए लगभग कितना लाना है।").setFontColor("#666666");
  if (r2rows.length) {
    const rule = SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("A").setBackground("#f4cccc").setRanges([s2.getRange(2, 9, r2rows.length, 1)]).build();
    const up = SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(0.1).setFontColor("#cc0000").setRanges([s2.getRange(2, 16, r2rows.length, 1)]).build();
    s2.setConditionalFormatRules([rule, up]);
  }
  s2.setFrozenRows(1); s2.setFrozenColumns(1); s2.setColumnWidth(1, 200);

  // ---------- 3. दाम पर नज़र ----------
  const s3 = xs.insertSheet("दाम पर नज़र");
  head(s3, 1, ["सामान", "दाम किसका", "पिछली अवधि औसत ₹", "इस अवधि औसत ₹", "बदलाव", "सबसे कम ₹", "सबसे ज़्यादा ₹", "सबसे सस्ता कब मिला", "दुकान"]);
  const watch = items.map(function (e) { const pe = P[e.n + "|" + e.u]; const ch = pe && pe.avg ? e.avg / pe.avg - 1 : null; return { e: e, pe: pe, ch: ch, spread: e.min > 0 ? e.max / e.min - 1 : 0 }; })
    .filter(function (w) { return (w.ch !== null && Math.abs(w.ch) >= 0.1) || w.spread >= 0.15; })
    .sort(function (a, b) { return Math.abs(b.ch || b.spread) - Math.abs(a.ch || a.spread); });
  const wr = watch.map(function (w) { const e = w.e, u = e.u === "ग्राम" ? 1000 : 1; return [e.n, unitRate(e.u), w.pe ? r2(w.pe.avg) : "", r2(e.avg), w.ch === null ? "" : w.ch, r2(e.min), r2(e.max), e.minD, shopOf[e.n] || ""]; });
  if (wr.length) { s3.getRange(2, 1, wr.length, 9).setValues(wr); s3.getRange(2, 5, wr.length, 1).setNumberFormat("+0%;-0%;0%"); s3.getRange(2, 3, wr.length, 4).setNumberFormat("₹#,##0.0"); s3.getRange(2, 5, wr.length, 1).setNumberFormat("+0%;-0%;0%"); }
  else s3.getRange(2, 1).setValue("इस अवधि में किसी सामान का दाम ज़्यादा ऊपर-नीचे नहीं हुआ।");
  s3.getRange(wr.length + 3, 1).setValue("यहाँ वो सामान हैं जिनका दाम पिछली अवधि से 10% से ज़्यादा बदला, या इसी अवधि में 15% से ज़्यादा ऊपर-नीचे हुआ। ग्राम वाले सामान का दाम 1 kg का है।").setFontColor("#666666");
  s3.setFrozenRows(1); s3.setColumnWidth(1, 200);

  // ---------- 4. रोज़ का खर्च (बिक्री भरें तो रोज़ की बचत %) ----------
  const s4 = xs.insertSheet("रोज़ का खर्च");
  head(s4, 1, ["तारीख", "दिन", "खरीद ₹", "बिक्री ₹", "बचत %"]);
  const WD = ["रविवार", "सोमवार", "मंगलवार", "बुधवार", "गुरुवार", "शुक्रवार", "शनिवार"];
  const dr = [];
  for (let i = 0; i < nDays; i++) {
    const d = new Date(D(from).getTime() + i * 864e5), k = S(d), r = i + 2;
    dr.push(["'" + k, WD[d.getDay()], r2(buyDays[k] || 0), saleDays[k] ? r2(saleDays[k]) : "", "=IF(D" + r + ">0,(D" + r + "-C" + r + ")/D" + r + ",\"\")"]);
  }
  if (dr.length) { s4.getRange(2, 1, dr.length, 5).setValues(dr); s4.getRange(2, 4, dr.length, 1).setBackground(Y); s4.getRange(2, 3, dr.length, 2).setNumberFormat("₹#,##0"); s4.getRange(2, 5, dr.length, 1).setNumberFormat("0.0%"); }
  const t4 = dr.length + 2;
  s4.getRange(t4, 1, 1, 5).setValues([["कुल", "", "=SUM(C2:C" + (t4 - 1) + ")", "=SUM(D2:D" + (t4 - 1) + ")", "=IF(D" + t4 + ">0,(D" + t4 + "-C" + t4 + ")/D" + t4 + ",\"\")"]]).setFontWeight("bold");
  s4.getRange(t4, 3, 1, 2).setNumberFormat("₹#,##0"); s4.getRange(t4, 5).setNumberFormat("0.0%");
  s4.getRange(t4 + 2, 1).setValue("बिक्री ऐप से अपने-आप आती है; जो दिन छूटा हो, पीले खाने में खुद भर सकते हैं। ध्यान दें: एक दिन की खरीद अगले कई दिन काम आती है, इसलिए रोज़ का % ऊपर-नीचे होगा – हफ़्ते/महीने का कुल % ही असली तस्वीर है।").setFontColor("#666666");
  s4.setFrozenRows(1);

  // ---------- 5. कैटेगरी और दुकान ----------
  const s5 = xs.insertSheet("कैटेगरी और दुकान");
  head(s5, 1, ["कैटेगरी", "कुल ₹", "हिस्सा %", "पिछली अवधि ₹", "बदलाव"]);
  const cat = {}, pcat = {};
  cur.forEach(function (x) { cat[x.c] = (cat[x.c] || 0) + amt(x); }); prev.forEach(function (x) { pcat[x.c] = (pcat[x.c] || 0) + amt(x); });
  const cr = Object.keys(cat).sort(function (a, b) { return cat[b] - cat[a]; }).map(function (c, i) { const r = i + 2; return [c, r2(cat[c]), total ? cat[c] / total : 0, r2(pcat[c] || 0), pcat[c] ? "=B" + r + "/D" + r + "-1" : ""]; });
  if (cr.length) { s5.getRange(2, 1, cr.length, 5).setValues(cr); s5.getRange(2, 2, cr.length, 1).setNumberFormat("₹#,##0"); s5.getRange(2, 3, cr.length, 1).setNumberFormat("0%"); s5.getRange(2, 4, cr.length, 1).setNumberFormat("₹#,##0"); s5.getRange(2, 5, cr.length, 1).setNumberFormat("+0%;-0%;0%"); }
  const sh0 = cr.length + 3, shop = {};
  cur.forEach(function (x) { const k = shopOf[x.n] || (special(x.c) ? x.c : "दुकान नहीं लिखी"); shop[k] = (shop[k] || 0) + amt(x); });
  head(s5, sh0, ["दुकान", "कुल ₹", "हिस्सा %"]);
  const srows = Object.keys(shop).sort(function (a, b) { return shop[b] - shop[a]; }).map(function (k) { return [k, r2(shop[k]), total ? shop[k] / total : 0]; });
  if (srows.length) { s5.getRange(sh0 + 1, 1, srows.length, 3).setValues(srows); s5.getRange(sh0 + 1, 2, srows.length, 1).setNumberFormat("₹#,##0"); s5.getRange(sh0 + 1, 3, srows.length, 1).setNumberFormat("0%"); }
  s5.getRange(sh0 + srows.length + 2, 1).setValue("दुकान के हिसाब से खर्च देखने के लिए Google Sheet के Items टैब में हर सामान की 'दुकान' भरें।").setFontColor("#666666");
  s5.setColumnWidth(1, 200);

  SpreadsheetApp.flush();
  const file = DriveApp.getFileById(xs.getId());
  const it = DriveApp.getFoldersByName("Secret Ovens – हिसाब (Excel)");
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder("Secret Ovens – हिसाब (Excel)");
  file.moveTo(folder);
  return { ok: true, url: "https://docs.google.com/spreadsheets/d/" + xs.getId() + "/export?format=xlsx", sheet: xs.getUrl() };
}
