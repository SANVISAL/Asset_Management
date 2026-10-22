/**
 * Staff Asset Manager — Google Apps Script
 * Sheets used: AllAssetLists (master list of assets), Staffs (staff list), History (created by setup)
 *
 * Rule: edit assets ONLY in AllAssetLists (or through the dashboard).
 * The Staffs device-ID columns are filled by "Sync Staffs sheet".
 */

const CONFIG = {
  ASSET_SHEET: "AllAssetLists",
  STAFF_SHEET: "Staffs",
  HISTORY_SHEET: "History",
  // Staffs layout:
  //   A = Name, B = Position, C = Email, D = Contact
  //   Department = a banner row (only column A filled, e.g. "OPERATION / HR");
  //   every staff row below it belongs to that department until the next banner.
  //   Device columns come in pairs: a checkbox column ("Laptop") + an ID column ("Laptop ID").
  STAFF_HEADER_ROW: 1,
  STAFF_COLS: { name: "A", position: "B", email: "C", contact: "D" },
  // Pending = returned by its holder, waiting to be checked; set it to "In stock" to move it back to the stock area
  STATUSES: ["In use", "Pending", "In stock", "Repair", "Lost", "Retired"],
  CONDITIONS: ["New", "Good", "Fair", "Damaged"],
  OLD_DEVICE_YEARS: 4,
};

/* ---------------------------------------------------------------- Menu & UI */

function onOpen() {
  rememberSpreadsheet_();
  SpreadsheetApp.getUi()
    .createMenu("🗂 Asset Manager")
    .addItem("Open dashboard", "openDashboard")
    // .addItem('Open dashboard in new tab (full screen)', 'openDashboardTab')
    .addSeparator()
    // .addItem('Set up sheets (run once)', 'setupSheets')
    .addItem("Sync Staffs sheet from assets", "syncStaffsSheet")
    // .addItem('Import holders from Staffs ID columns', 'importHoldersFromStaffs')
    .addToUi();
}

function openDashboard() {
  const html = HtmlService.createHtmlOutputFromFile("Dashboard")
    .setWidth(1600)
    .setHeight(900);
  SpreadsheetApp.getUi().showModalDialog(html, "Asset Manager");
}

/**
 * Opens the dashboard as its own browser tab (the web app), which can fill the whole screen.
 * Needs one deployment first: Deploy → New deployment → Web app.
 */
function openDashboardTab() {
  rememberSpreadsheet_();
  const ui = SpreadsheetApp.getUi();
  const url = ScriptApp.getService().getUrl();
  if (!url) {
    ui.alert(
      "Deploy the dashboard once first",
      'In Extensions → Apps Script: Deploy → New deployment → type "Web app" → Deploy, and allow access.\n' +
        "Then reload this sheet and use this menu item again.",
      ui.ButtonSet.OK
    );
    return;
  }
  const html = HtmlService.createHtmlOutput(
    '<base target="_blank"><div style="font:14px sans-serif;padding:8px">' +
      "<p>If the dashboard did not open by itself:</p>" +
      '<p><a href="' +
      url +
      '" style="font-size:16px;font-weight:bold" onclick="setTimeout(google.script.host.close, 300)">Open dashboard ↗</a></p></div>' +
      '<script>if (window.open("' +
      url +
      '", "_blank")) google.script.host.close();</script>'
  )
    .setWidth(320)
    .setHeight(130);
  ui.showModalDialog(html, "Asset Manager");
}

/** The web app has no "active" spreadsheet, so remember which one this is. */
function rememberSpreadsheet_() {
  try {
    PropertiesService.getScriptProperties().setProperty(
      "SPREADSHEET_ID",
      SpreadsheetApp.getActive().getId()
    );
  } catch (e) {}
}

/** Lets you also deploy the dashboard as a web app (Deploy → New deployment → Web app). */
function doGet() {
  return HtmlService.createHtmlOutputFromFile("Dashboard")
    .setTitle("Asset Manager")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

/* ---------------------------------------------------------------- Setup */

function setupSheets() {
  const sh = sheet_(CONFIG.ASSET_SHEET);
  let map = headerMap_(sh);

  // 1. Add a Status column if missing, pre-filled from "Using By Name"
  if (col_(map, "status") < 0) {
    const newCol = sh.getLastColumn() + 1;
    sh.getRange(1, newCol).setValue("Status").setFontWeight("bold");
    const last = sh.getLastRow();
    const c = assetCols_(map);
    if (last > 1 && c.id >= 0 && c.user >= 0) {
      const rows = sh.getRange(2, 1, last - 1, newCol - 1).getValues();
      sh.getRange(2, newCol, last - 1, 1).setValues(
        rows.map((r) => [
          !String(r[c.id]).trim() || isAssetBanner_(r, c)
            ? ""
            : String(r[c.user]).trim()
            ? "In use"
            : "In stock",
        ])
      );
    }
    map = headerMap_(sh);
  }

  // 2. Dropdowns for Status and Condition
  const rows = Math.max(sh.getMaxRows() - 1, 1);
  sh.getRange(2, col_(map, "status") + 1, rows, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(CONFIG.STATUSES, true)
      .setAllowInvalid(false)
      .build()
  );
  const condCol = col_(map, "condition");
  if (condCol >= 0) {
    sh.getRange(2, condCol + 1, rows, 1).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(CONFIG.CONDITIONS, true)
        .setAllowInvalid(true)
        .build()
    );
  }

  // 3. History sheet
  let h = ss_().getSheetByName(CONFIG.HISTORY_SHEET);
  if (!h) {
    h = ss_().insertSheet(CONFIG.HISTORY_SHEET);
    h.appendRow([
      "Timestamp",
      "Asset ID",
      "Action",
      "From",
      "To",
      "Condition",
      "Done by",
      "Note",
    ]);
    h.setFrozenRows(1);
    h.getRange(1, 1, 1, 8).setFontWeight("bold");
  }

  SpreadsheetApp.getUi().alert(
    "Setup complete: Status column, dropdowns and History sheet are ready."
  );
}

/* ---------------------------------------------------------------- Read data */

// Per-run memo: Apps Script starts every call fresh, so this only saves repeated lookups within one action
const RUN_ = {};

function ss_() {
  if (RUN_.ss) return RUN_.ss;
  const active = SpreadsheetApp.getActive();
  if (active) return (RUN_.ss = active);
  const id =
    PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID"); // web app: see rememberSpreadsheet_
  if (!id)
    throw new Error(
      "Open the spreadsheet once (so the Asset Manager menu loads), then reload this page."
    );
  return (RUN_.ss = SpreadsheetApp.openById(id));
}

function tz_() {
  return RUN_.tz || (RUN_.tz = ss_().getSpreadsheetTimeZone());
}

function sheet_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh) throw new Error("Sheet not found: " + name);
  return sh;
}

function headerMap_(sh) {
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h, i) => {
    map[String(h).trim().toLowerCase()] = i;
  });
  return map;
}

function col_(map, ...names) {
  for (const n of names) {
    const k = n.toLowerCase();
    if (k in map) return map[k];
  }
  return -1;
}

function fmtDate_(d) {
  return d instanceof Date
    ? Utilities.formatDate(d, tz_(), "yyyy-MM-dd")
    : String(d || "");
}

/** A type banner row in AllAssetLists (e.g. "LAPTOP"): has text, but none of the device columns are filled. */
function isAssetBanner_(row, c) {
  const has = (i) =>
    i >= 0 && String(row[i] == null ? "" : row[i]).trim() !== "";
  return (
    !!bannerText_(row) &&
    ![
      c.type,
      c.user,
      c.model,
      c.desc,
      c.purchase,
      c.condition,
      c.serial,
      c.prev,
      c.remarks,
    ].some(has)
  );
}

function bannerText_(row) {
  const v = row.find((x) => String(x == null ? "" : x).trim() !== "");
  return v === undefined ? "" : String(v).trim();
}

function assetCols_(map) {
  return {
    id: col_(map, "asset id"),
    user: col_(map, "using by name"),
    type: col_(map, "asset type"),
    model: col_(map, "brand / model", "brand/model", "model"),
    desc: col_(map, "description"),
    purchase: col_(map, "purchase date"),
    condition: col_(map, "condition"),
    prev: col_(map, "previous holder"),
    remarks: col_(map, "remarks"),
    status: col_(map, "status"),
    serial: col_(map, "serial no.", "serial no", "serial"),
  };
}

function readAssets_() {
  const sh = sheet_(CONFIG.ASSET_SHEET);
  const values = sh.getDataRange().getValues();
  const c = assetCols_(headerMap_(sh));
  if (c.id < 0)
    throw new Error('Column "Asset ID" not found in ' + CONFIG.ASSET_SHEET);
  if (c.user < 0)
    throw new Error(
      'Column "Using By Name" not found in ' + CONFIG.ASSET_SHEET
    );

  const assets = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const get = (i) => (i >= 0 && row[i] !== null ? String(row[i]).trim() : "");
    const id = get(c.id);
    if (!id || isAssetBanner_(row, c)) continue; // empty row or type banner
    const user = get(c.user);
    const desc = get(c.desc);
    const pd = c.purchase >= 0 ? row[c.purchase] : "";
    const serialMatch = desc.match(/S\/N\s*:\s*([^\s]+)/i);
    assets.push({
      row: r + 1,
      id,
      user,
      type: get(c.type).toUpperCase(),
      model: get(c.model),
      description: desc,
      serial: get(c.serial) || (serialMatch ? serialMatch[1] : ""),
      purchaseDate: fmtDate_(pd),
      ageYears:
        pd instanceof Date
          ? (Date.now() - pd.getTime()) / (365.25 * 864e5)
          : null,
      condition: get(c.condition),
      previousHolder: get(c.prev),
      previousHolders: parseIds_(get(c.prev)),
      remarks: get(c.remarks),
      status: get(c.status) || (user ? "In use" : "In stock"),
    });
  }
  return { sh, c, assets, values };
}

function letterToIndex_(letter) {
  let n = 0;
  for (const ch of String(letter).toUpperCase())
    n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1; // 0-based
}

/** Works out which column holds what in Staffs (0-based indexes, -1 = not present). */
/** Staffs header row as trimmed strings (read once per run). */
function staffHeaders_(sh) {
  return (
    RUN_.staffHeaders ||
    (RUN_.staffHeaders = sh
      .getRange(CONFIG.STAFF_HEADER_ROW, 1, 1, sh.getLastColumn())
      .getValues()[0]
      .map((h) => String(h).trim()))
  );
}

function staffLayout_(sh) {
  const headers = staffHeaders_(sh).map((h) => h.toLowerCase());
  const find = (forced, ...words) => {
    if (forced) return letterToIndex_(forced);
    return headers.findIndex((h) =>
      words.some((w) => h === w || h.startsWith(w + " ") || h.endsWith(" " + w))
    );
  };
  const k = CONFIG.STAFF_COLS;
  const layout = {
    name: find(k.name, "name", "staff name", "full name"),
    position: find(k.position, "position", "title", "role"),
    email: find(k.email, "email", "e-mail"),
    contact: find(k.contact, "contact", "contract", "tel"),
  };
  if (layout.name < 0)
    throw new Error(
      'Could not find a "Name" header in ' +
        CONFIG.STAFF_SHEET +
        " — set CONFIG.STAFF_COLS.name to its column letter."
    );
  return layout;
}

/** Staff list (read once per run; addStaff clears RUN_.staff). The Staffs sync never changes it. */
function readStaff_() {
  return RUN_.staff || (RUN_.staff = readStaffFresh_());
}

function readStaffFresh_() {
  const sh = sheet_(CONFIG.STAFF_SHEET);
  const v = sh.getDataRange().getValues();
  const L = staffLayout_(sh);
  const staff = [];
  let dept = "";
  for (let r = CONFIG.STAFF_HEADER_ROW; r < v.length; r++) {
    const g = (i) =>
      i >= 0 ? String(v[r][i] == null ? "" : v[r][i]).trim() : "";
    const name = g(L.name);
    if (!name) continue; // empty row
    if (!g(L.position) && !g(L.email) && !g(L.contact)) {
      // banner row = department
      dept = name;
      continue;
    }
    staff.push({
      row: r + 1,
      name,
      dept,
      position: g(L.position),
      email: g(L.email),
      contact: g(L.contact),
    });
  }
  return staff;
}

function readHistory_(limit) {
  const h = ss_().getSheetByName(CONFIG.HISTORY_SHEET);
  if (!h || h.getLastRow() < 2) return [];
  const n = Math.min(limit, h.getLastRow() - 1);
  const rows = h
    .getRange(h.getLastRow() - n + 1, 1, n, 8)
    .getValues()
    .reverse();
  const tz = tz_();
  return rows.map((r) => ({
    time:
      r[0] instanceof Date
        ? Utilities.formatDate(r[0], tz, "yyyy-MM-dd HH:mm")
        : String(r[0]),
    id: String(r[1]),
    action: String(r[2]),
    from: String(r[3]),
    to: String(r[4]),
    condition: String(r[5]),
    by: String(r[6]),
    note: String(r[7]),
  }));
}

/* ---------------------------------------------------------------- Dashboard data */

/** Called by the dashboard with no argument; actions pass the assets they already read. */
function getDashboardData(data) {
  const { assets } = data && data.assets ? data : readAssets_();
  const staff = readStaff_();
  const staffByName = {};
  staff.forEach((s) => {
    staffByName[s.name.toLowerCase()] = s;
  });
  assets.forEach((a) => {
    const s = staffByName[a.user.toLowerCase()];
    a.dept = s ? s.dept : "";
  });

  const count = (f) => assets.filter(f).length;
  const kpis = {
    total: assets.length,
    inUse: count((a) => a.status === "In use"),
    inStock: count((a) => a.status === "In stock"),
    pending: count((a) => a.status === "Pending"),
    repair: count((a) => a.status === "Repair"),
    lostDamaged: count(
      (a) => a.status === "Lost" || a.condition.toLowerCase() === "damaged"
    ),
    staff: staff.length,
  };

  const byType = {};
  assets.forEach((a) => {
    const t = a.type || "OTHER";
    byType[t] = byType[t] || {
      type: t,
      total: 0,
      inUse: 0,
      inStock: 0,
      other: 0,
    };
    byType[t].total++;
    if (a.status === "In use") byType[t].inUse++;
    else if (a.status === "In stock") byType[t].inStock++;
    else byType[t].other++;
  });

  const tally = (list, key) => {
    const o = {};
    list.forEach((a) => {
      const k = key(a) || "Unknown";
      o[k] = (o[k] || 0) + 1;
    });
    return o;
  };

  const alerts = {
    available: assets.filter((a) => a.status === "In stock"),
    unknownHolder: assets.filter(
      (a) => a.user && !staffByName[a.user.toLowerCase()]
    ),
    statusConflict: assets.filter(
      (a) =>
        (a.status === "In use" && !a.user) ||
        (a.status === "In stock" && a.user)
    ),
    old: assets.filter(
      (a) => a.ageYears !== null && a.ageYears >= CONFIG.OLD_DEVICE_YEARS
    ),
    missingInfo: assets.filter(
      (a) => !a.purchaseDate || !a.condition || !a.serial
    ),
    repair: assets.filter((a) => a.status === "Repair"),
    pending: assets.filter((a) => a.status === "Pending"),
  };

  return {
    kpis,
    byType: Object.values(byType).sort((a, b) => b.total - a.total),
    byDept: tally(
      assets.filter((a) => a.status === "In use"),
      (a) => a.dept
    ),
    byCondition: tally(assets, (a) => a.condition),
    alerts,
    assets,
    staff,
    history: readHistory_(30),
    statuses: CONFIG.STATUSES,
    conditions: CONFIG.CONDITIONS,
    oldYears: CONFIG.OLD_DEVICE_YEARS,
    updated: Utilities.formatDate(
      new Date(),
      tz_(),
      "yyyy-MM-dd HH:mm"
    ),
  };
}

/* ---------------------------------------------------------------- Actions */

function withLock_(fn) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(15000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function findAsset_(assetId) {
  const data = readAssets_();
  const a = data.assets.find((x) => x.id === assetId);
  if (!a) throw new Error("Asset not found: " + assetId);
  return Object.assign(data, { a });
}

function setCell_(sh, row, colIndex, value) {
  if (colIndex >= 0) sh.getRange(row, colIndex + 1).setValue(value);
}

/* ---------------------------------------------------------------- AllAssetLists layout: USING area + stock area */

// Banner words that start the stock area (the USING area also ends where a type banner repeats, e.g. a 2nd "LAPTOP")
const STOCK_BANNERS = ["STOCK", "INSTOCK", "AVAILABLE", "SPARE", "SPARES"];
const normWord_ = (s) =>
  String(s == null ? "" : s)
    .toUpperCase()
    .replace(/[^A-Z]/g, "");

/**
 * Splits AllAssetLists into areas. With a "USING" banner (row 2), devices held by someone live under it,
 * grouped by type banners; everything else is the stock area. Without a USING banner it's all one stock area.
 * Returns { using, stock, areaOf } — each area { end, types: { BANNERWORD: { banner, last } } }, rows 1-based.
 */
function assetAreas_(values, c) {
  const areas = {
    using: null,
    stock: { end: 0, types: {} },
    areaOf: {},
    lastRow: values.length,
  };
  let area = "stock",
    key = null;
  for (let r = 1; r < values.length; r++) {
    const row = values[r],
      n = r + 1;
    if (!bannerText_(row)) continue; // empty row
    if (isAssetBanner_(row, c)) {
      const t = normWord_(bannerText_(row));
      if (t === "USING") {
        area = "using";
        areas.using = { end: n, types: {} };
        key = null;
        continue;
      }
      if (STOCK_BANNERS.indexOf(t) >= 0) {
        area = "stock";
        areas.stock.end = Math.max(areas.stock.end, n);
        key = null;
        continue;
      }
      if (area === "using" && areas.using.types[t]) area = "stock"; // same type banner again → stock area begins
      key = t;
      areas[area].types[t] = { banner: n, last: n };
      areas[area].end = n;
      continue;
    }
    areas[area].end = n;
    areas.areaOf[n] = area;
    const k = key || normWord_(row[c.type]);
    (areas[area].types[k] = areas[area].types[k] || {
      banner: 0,
      last: n,
    }).last = n;
  }
  return areas;
}

/** The type group inside an area for an asset type ("LAPTOP" ↔ banner "LAPTOP" or "LAPTOPS"), or null. */
function typeGroup_(area, type) {
  const k = normWord_(type);
  return area.types[k] || area.types[k + "S"] || null;
}

/**
 * Row index where a device of `type` belongs in `areaName` (insert before it).
 * Creates the type banner at the end of the area when it doesn't exist yet.
 * Returns { dest, inserted } — inserted = row of a new banner (0 if none), so callers can shift their own rows.
 */
function typeSlot_(sh, c, areas, areaName, type) {
  const A = areas[areaName];
  const g = typeGroup_(A, type);
  if (g) return { dest: g.last + 1, inserted: 0 };
  const after = A.end || areas.lastRow; // empty area → bottom of the sheet
  const at = after + 1;
  sh.insertRowAfter(after);
  const width = sh.getLastColumn();
  // style the new banner like an existing one
  const anyBanner = Object.values(areas.stock.types)
    .concat(areas.using ? Object.values(areas.using.types) : [])
    .map((g2) => g2.banner)
    .filter(Boolean)[0];
  if (anyBanner)
    sh.getRange(
      anyBanner >= at ? anyBanner + 1 : anyBanner,
      1,
      1,
      width
    ).copyTo(
      sh.getRange(at, 1, 1, width),
      SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
      false
    );
  sh.getRange(at, c.id + 1)
    .setValue(type)
    .setFontWeight("bold");
  return { dest: at + 1, inserted: at };
}

/**
 * Moves a device's row into the right area after a change:
 * has a holder → USING area; status "In stock" → stock area; anything else (Pending, Repair…) stays put.
 * Does nothing when the sheet has no USING banner.
 */
function placeAsset_(assetId, data) {
  // data = assets already read in this action (with user/status updated in memory) — saves a re-read.
  // Only valid until a row moves; returns true when it moved one, so callers re-read after that.
  const { sh, c, assets, values } = data || readAssets_();
  const areas = assetAreas_(values, c);
  if (!areas.using) return false;
  const a = assets.find((x) => x.id === assetId);
  if (!a) return false;
  const want = a.user ? "using" : a.status === "In stock" ? "stock" : null;
  if (!want || areas.areaOf[a.row] === want) return false;
  const slot = typeSlot_(sh, c, areas, want, a.type);
  const src = slot.inserted && a.row >= slot.inserted ? a.row + 1 : a.row;
  sh.moveRows(sh.getRange(src + ":" + src), slot.dest);
  return true;
}

/** Previous Holder keeps everyone who has had the device, oldest first ("SK, DARA"); a returning person moves to the end. */
function addPrevHolder_(sh, c, a, name) {
  if (!name) return;
  const list = parseIds_(a.previousHolder).filter(
    (n) => n.toLowerCase() !== name.toLowerCase()
  );
  list.push(name);
  a.previousHolder = list.join(", ");
  setCell_(sh, a.row, c.prev, a.previousHolder);
}

function log_(assetId, action, from, to, condition, note) {
  const h =
    RUN_.history || (RUN_.history = ss_().getSheetByName(CONFIG.HISTORY_SHEET));
  if (!h) throw new Error('History sheet missing — run "Set up sheets" first.');
  if (RUN_.email === undefined)
    RUN_.email = Session.getActiveUser().getEmail() || "";
  h.appendRow([
    new Date(),
    assetId,
    action,
    from || "",
    to || "",
    condition || "",
    RUN_.email,
    note || "",
  ]);
}

/** Give an asset to a staff member. */
function assignAsset(assetId, staffName, note) {
  return withLock_(() => {
    const data = findAsset_(assetId);
    const { sh, c, a } = data;
    if (!readStaff_().some((s) => s.name === staffName))
      throw new Error("Staff not found: " + staffName);
    if (a.status === "Lost" || a.status === "Retired")
      throw new Error(assetId + " is " + a.status + " and cannot be assigned.");
    if (a.user && a.user !== staffName) addPrevHolder_(sh, c, a, a.user);
    setCell_(sh, a.row, c.user, staffName);
    setCell_(sh, a.row, c.status, "In use");
    log_(
      assetId,
      a.user ? "Reassigned" : "Assigned",
      a.user,
      staffName,
      a.condition,
      note
    );
    a.user = staffName;
    a.status = "In use";
    placeAsset_(assetId, data);
    return afterChange_();
  });
}

/** Take an asset back from its holder: Pending (to be checked), or Repair if damaged. The row stays where it is. */
function returnAsset(assetId, condition, note) {
  return withLock_(() => {
    const { sh, c, a } = findAsset_(assetId);
    if (!a.user) throw new Error(assetId + " is not assigned to anyone.");
    const newStatus = condition === "Damaged" ? "Repair" : "Pending";
    addPrevHolder_(sh, c, a, a.user);
    setCell_(sh, a.row, c.user, "");
    setCell_(sh, a.row, c.status, newStatus);
    if (condition) setCell_(sh, a.row, c.condition, condition);
    log_(
      assetId,
      "Returned",
      a.user,
      newStatus,
      condition || a.condition,
      note
    );
    return afterChange_();
  });
}

/** Change status (Repair, Lost, Retired, In stock...). */
function setAssetStatus(assetId, status, note) {
  return withLock_(() => {
    if (CONFIG.STATUSES.indexOf(status) < 0)
      throw new Error("Unknown status: " + status);
    const data = findAsset_(assetId);
    const { sh, c, a } = data;
    setCell_(sh, a.row, c.status, status);
    if ((status === "In stock" || status === "Retired") && a.user) {
      addPrevHolder_(sh, c, a, a.user);
      setCell_(sh, a.row, c.user, "");
      a.user = "";
    }
    log_(assetId, "Status → " + status, a.status, status, a.condition, note);
    a.status = status;
    placeAsset_(assetId, data); // "In stock" moves it down to the stock area
    return afterChange_();
  });
}

/**
 * Add a new device to AllAssetLists, inside its type's section (USING area if it has a holder, else stock):
 * after the last device of that type, else right under a matching banner ("LAPTOP" / "Laptops"),
 * else under a new banner at the end of that area.
 */
function addAsset(f) {
  return withLock_(() => {
    const t = (k) => String(f[k] == null ? "" : f[k]).trim();
    const id = t("id"),
      type = t("type").toUpperCase(),
      user = t("user"),
      serial = t("serial");
    if (!id) throw new Error("Asset ID is required.");
    if (!type) throw new Error("Asset type is required.");
    const { sh, c, assets, values } = readAssets_();
    if (assets.some((a) => a.id.toLowerCase() === id.toLowerCase()))
      throw new Error("Asset ID already exists: " + id);
    if (user && !readStaff_().some((s) => s.name === user))
      throw new Error("Staff not found: " + user);

    let desc = t("description");
    if (serial && c.serial < 0 && !/S\/N\s*:/i.test(desc))
      desc = (desc ? desc + "\n" : "") + "S/N: " + serial;
    const pd = t("purchaseDate").match(/^(\d{4})-(\d{2})-(\d{2})$/);

    const width = sh.getLastColumn();
    const row = new Array(width).fill("");
    const put = (i, v) => {
      if (i >= 0) row[i] = v;
    };
    put(c.id, id);
    put(c.user, user);
    put(c.type, type);
    put(c.model, t("model"));
    put(c.desc, desc);
    put(c.serial, serial);
    put(c.purchase, pd ? new Date(+pd[1], pd[2] - 1, +pd[3]) : "");
    put(c.condition, t("condition"));
    put(c.remarks, t("remarks"));
    put(c.status, user ? "In use" : "In stock");

    // Where does it go? Under its type banner — in the USING area when it has a holder, else the stock area
    const areas = assetAreas_(values, c);
    const areaName = user && areas.using ? "using" : "stock";
    const g = typeGroup_(areas[areaName], type);
    const lastOfType = g && g.last !== g.banner ? g.last : -1; // a device row of this type, for formatting
    const slot = typeSlot_(sh, c, areas, areaName, type);
    const newRow = slot.dest;
    sh.insertRowAfter(newRow - 1);
    // Copy formatting and dropdowns from a device row (never from a banner)
    let tpl = lastOfType > 0 ? lastOfType : assets.length ? assets[0].row : -1;
    if (tpl > 0 && slot.inserted && tpl >= slot.inserted) tpl++; // shifted by a new banner
    if (tpl >= newRow) tpl++; // shifted by the new row
    if (tpl > 0) {
      const src = sh.getRange(tpl, 1, 1, width),
        dst = sh.getRange(newRow, 1, 1, width);
      src.copyTo(dst, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      src.copyTo(
        dst,
        SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION,
        false
      );
    }
    sh.getRange(newRow, 1, 1, width).setValues([row]);

    log_(id, "Added", "", user || "In stock", t("condition"), t("note"));
    return afterChange_();
  });
}

/**
 * Add a staff member to Staffs, inserted after the last person of their department.
 * A department that doesn't exist yet gets a new banner row at the bottom.
 */
function addStaff(f) {
  return withLock_(() => {
    const t = (k) => String(f[k] == null ? "" : f[k]).trim();
    const name = t("name"),
      dept = t("dept"),
      position = t("position"),
      email = t("email"),
      contact = t("contact");
    if (!name) throw new Error("Name is required.");
    if (!dept) throw new Error("Department is required.");
    if (!position && !email && !contact)
      throw new Error(
        "Fill in at least a position, email or contact — a row with only a name is read as a department heading."
      );
    const staff = readStaff_();
    if (staff.some((s) => s.name.toLowerCase() === name.toLowerCase()))
      throw new Error("Staff already exists: " + name);

    const sh = sheet_(CONFIG.STAFF_SHEET);
    const L = staffLayout_(sh);
    const v = sh.getDataRange().getValues();
    const g = (r, i) =>
      i >= 0 ? String(v[r][i] == null ? "" : v[r][i]).trim() : "";
    let cur = "",
      last = -1,
      lastIsBanner = false; // last = 1-based row of the department's last row (banner or staff)
    for (let r = CONFIG.STAFF_HEADER_ROW; r < v.length; r++) {
      const n = g(r, L.name);
      if (!n) continue;
      const banner = !g(r, L.position) && !g(r, L.email) && !g(r, L.contact);
      if (banner) cur = n;
      if (cur.toLowerCase() === dept.toLowerCase()) {
        last = r + 1;
        lastIsBanner = banner;
      }
    }

    const width = sh.getLastColumn();
    let row;
    if (last > 0) {
      sh.insertRowAfter(last);
      row = last + 1;
    } else {
      const banner = sh.getLastRow() + 1;
      sh.getRange(banner, L.name + 1)
        .setValue(dept)
        .setFontWeight("bold");
      row = banner + 1;
    }
    // Copy formatting and checkboxes from a staff row (never from a banner, which may be bold/merged)
    let tpl =
      last > 0 && !lastIsBanner ? last : staff.length ? staff[0].row : -1;
    if (tpl > 0 && last > 0 && tpl > last) tpl++; // that row moved down by the insert
    if (tpl > 0) {
      const src = sh.getRange(tpl, 1, 1, width),
        dst = sh.getRange(row, 1, 1, width);
      src.copyTo(dst, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      src.copyTo(
        dst,
        SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION,
        false
      );
    }
    const put = (i, val) => {
      if (i >= 0) sh.getRange(row, i + 1).setValue(val);
    };
    put(L.name, name);
    put(L.position, position);
    put(L.email, email);
    put(L.contact, contact);
    RUN_.staff = null; // the staff list just changed
    return afterChange_();
  });
}

/* ---------------------------------------------------------------- Sync Staffs sheet */

/**
 * Rebuilds the device columns of Staffs from AllAssetLists:
 *  - "<Type> ID" columns (Laptop ID, Phone ID…) get the asset IDs, each linked to its row in AllAssetLists
 *  - the matching checkbox column ("Laptop", "Phone"…) is ticked when the person holds one
 * Department banner rows and empty rows are left untouched.
 * The header word must match the Asset Type (LAPTOP ↔ "Laptop ID"); otherwise add it to TYPE_ALIASES.
 */
const TYPE_ALIASES = {
  // 'HEADER WORD': 'ASSET TYPE IN AllAssetLists', e.g. 'PHONE': 'MOBILE'
};

/** Finds the "<Type> ID" columns of Staffs and their checkbox partner ("<Type>"). Columns are 1-based. */
function staffIdColumns_(sh) {
  const headers = staffHeaders_(sh);
  const L = staffLayout_(sh);
  const infoCols = [L.name, L.position, L.email, L.contact];
  const targets = [];
  headers.forEach((h, i) => {
    const m = h.match(/^(.+?)\s*id$/i);
    if (!m || infoCols.indexOf(i) >= 0) return;
    const word = m[1].trim();
    const type = TYPE_ALIASES[word.toUpperCase()] || word.toUpperCase();
    const box = headers.findIndex(
      (x) => x.toLowerCase() === word.toLowerCase()
    );
    targets.push({
      col: i + 1,
      header: h,
      type,
      boxCol: box >= 0 ? box + 1 : 0,
    });
  });
  return targets;
}

/** Menu item: sync and tell the user what was synced. */
function syncStaffsSheet() {
  const types = syncStaffs_();
  SpreadsheetApp.getUi().alert(
    types.length
      ? "Staffs sheet synced: " + types.join(", ")
      : 'No columns ending in "ID" found in ' + CONFIG.STAFF_SHEET
  );
}

/**
 * After a dashboard action: return fresh data right away.
 * The dashboard then calls syncStaffsQuiet() in the background, so the user doesn't wait for the Staffs sync.
 */
function afterChange_() {
  return getDashboardData();
}

/** Called by the dashboard after each action. A sync problem never undoes the action. */
function syncStaffsQuiet() {
  withLock_(() => syncStaffs_());
}

/** Does the sync without any UI (safe from the web app). Returns the synced types. */
function syncStaffs_() {
  const sh = sheet_(CONFIG.STAFF_SHEET);
  const hr = CONFIG.STAFF_HEADER_ROW;
  const lastRow = sh.getLastRow();
  if (lastRow <= hr) return [];
  const targets = staffIdColumns_(sh);
  if (!targets.length) return [];

  const { sh: assetSh, assets } = readAssets_();
  const gid = assetSh.getSheetId();
  const assetRow = {};
  const idx = {}; // name -> type -> [ids]
  assets.forEach((a) => {
    assetRow[a.id] = a.row;
    if (!a.user || a.status !== "In use") return;
    const n = a.user.toLowerCase();
    idx[n] = idx[n] || {};
    (idx[n][a.type] = idx[n][a.type] || []).push(a.id);
  });

  const staffRows = {};
  readStaff_().forEach((s) => {
    staffRows[s.row] = s.name.toLowerCase();
  });

  const n = lastRow - hr;
  const link = (id) =>
    assetRow[id] ? "#gid=" + gid + "&range=A" + assetRow[id] : null;

  // One read for all ID + checkbox columns, then write back only the columns that changed
  const cols = [];
  targets.forEach((t) => cols.push(t.col, t.boxCol || t.col));
  const c0 = Math.min.apply(null, cols),
    w = Math.max.apply(null, cols) - c0 + 1;
  const block = sh.getRange(hr + 1, c0, n, w);
  const rich = block.getRichTextValues();
  const plain = block.getValues();
  // "text|start-end-url|…" — compares a cell's text and links
  const keyOf = (r) =>
    r
      ? r.getText() +
        r
          .getRuns()
          .filter((x) => x.getLinkUrl())
          .map((x) => "|" + x.getStartIndex() + "-" + x.getEndIndex() + "-" + x.getLinkUrl())
          .join("")
      : null;

  targets.forEach((t) => {
    const j = t.col - c0,
      bj = t.boxCol ? t.boxCol - c0 : -1;
    const cells = [],
      boxes = [];
    let idChanged = false,
      boxChanged = false;
    for (let i = 0; i < n; i++) {
      const cur = rich[i][j];
      const name = staffRows[i + hr + 1];
      if (name === undefined) {
        // banner / empty row: leave as is (non-text cells come back null — make them text so the write accepts them)
        cells.push([cur || SpreadsheetApp.newRichTextValue().setText(String(plain[i][j])).build()]);
        if (bj >= 0) boxes.push([plain[i][bj]]);
        continue;
      }
      const ids = (idx[name] || {})[t.type] || [];
      let b = SpreadsheetApp.newRichTextValue().setText(ids.join(", "));
      let pos = 0;
      ids.forEach((id) => {
        const url = link(id);
        if (url) b = b.setLinkUrl(pos, pos + id.length, url);
        pos += id.length + 2;
      });
      const want = b.build();
      if (keyOf(want) !== keyOf(cur)) idChanged = true;
      cells.push([want]);
      if (bj >= 0) {
        const tick = ids.length > 0;
        if ((plain[i][bj] === true) !== tick) boxChanged = true;
        boxes.push([tick]);
      }
    }
    if (idChanged) sh.getRange(hr + 1, t.col, n, 1).setRichTextValues(cells);
    if (boxChanged) sh.getRange(hr + 1, t.boxCol, n, 1).setValues(boxes);
  });

  return targets.map((t) => t.type);
}
/* ---------------------------------------------------------------- Import holders from Staffs */

/** "PHN-0005, PHN-0049" (commas, semicolons or new lines) → ['PHN-0005', 'PHN-0049'] */
function parseIds_(cell) {
  return String(cell == null ? "" : cell)
    .split(/[,;\n]+/)
    .map((x) => x.trim())
    .filter(Boolean);
}

/**
 * Simple trigger: typing in a Staffs "<Type> ID" cell updates AllAssetLists right away.
 *  - an ID added to the cell  → that person becomes the holder (In use)
 *  - an ID removed from the cell → the device goes back to stock
 * Then the Staffs sheet is re-synced, so the cell shows exactly what AllAssetLists says.
 */
function onEdit(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  if (
    sh.getName() !== CONFIG.STAFF_SHEET ||
    e.range.getLastRow() <= CONFIG.STAFF_HEADER_ROW
  )
    return;
  const cols = staffIdColumns_(sh).filter(
    (t) => t.col >= e.range.getColumn() && t.col <= e.range.getLastColumn()
  );
  if (!cols.length) return;
  try {
    const msgs = withLock_(() => applyStaffIdEdits_(sh, e.range, cols));
    if (msgs.length) ss_().toast(msgs.join("\n"), "Asset Manager", 8);
  } catch (err) {
    ss_().toast(err.message, "Asset Manager — not saved", 10);
  }
}

function applyStaffIdEdits_(sh, range, cols) {
  const byRow = {};
  readStaff_().forEach((s) => {
    byRow[s.row] = s;
  });
  const data = readAssets_();
  const { sh: assetSh, c, assets } = data;
  const byId = {};
  assets.forEach((a) => {
    byId[a.id.toLowerCase()] = a;
  });
  const same = (x, y) => x.toLowerCase() === y.toLowerCase();
  const msgs = [],
    moved = new Set();
  const r0 = Math.max(range.getRow(), CONFIG.STAFF_HEADER_ROW + 1),
    r1 = range.getLastRow();

  cols.forEach((col) => {
    sh.getRange(r0, col.col, r1 - r0 + 1, 1)
      .getValues()
      .forEach((v, i) => {
        const s = byRow[r0 + i];
        if (!s) return; // department banner or empty row
        const typed = parseIds_(v[0]);
        const keep = new Set(typed.map((x) => x.toLowerCase()));

        typed.forEach((id) => {
          const a = byId[id.toLowerCase()];
          if (!a)
            return msgs.push(
              id + " is not in " + CONFIG.ASSET_SHEET + " — removed"
            );
          if (a.status === "Lost" || a.status === "Retired")
            return msgs.push(a.id + " is " + a.status + " — not assigned");
          const mine = same(a.user, s.name);
          if (mine && a.status !== "In stock") return; // already theirs
          const from = mine ? "" : a.user;
          if (from) addPrevHolder_(assetSh, c, a, from);
          setCell_(assetSh, a.row, c.user, s.name);
          setCell_(assetSh, a.row, c.status, "In use");
          log_(
            a.id,
            from ? "Reassigned" : "Assigned",
            from,
            s.name,
            a.condition,
            "Typed in Staffs"
          );
          msgs.push(
            a.id +
              " → " +
              s.name +
              (from ? " (was " + from + ")" : "") +
              (a.type !== col.type
                ? " — it is a " + a.type + ", moved to its own column"
                : "")
          );
          a.user = s.name;
          a.status = "In use";
          moved.add(a.id);
        });

        // Devices of this column's type the person held, but no longer typed in the cell → back to stock
        assets
          .filter(
            (a) =>
              same(a.user, s.name) &&
              a.status === "In use" &&
              a.type === col.type &&
              !keep.has(a.id.toLowerCase())
          )
          .forEach((a) => {
            addPrevHolder_(assetSh, c, a, a.user);
            setCell_(assetSh, a.row, c.user, "");
            setCell_(assetSh, a.row, c.status, "Pending");
            log_(
              a.id,
              "Returned",
              a.user,
              "Pending",
              a.condition,
              "Removed in Staffs"
            );
            msgs.push(a.id + " returned — Pending (was " + a.user + ")");
            a.user = "";
            a.status = "Pending";
          });
      });
  });
  // Move rows only after all writes: moving shifts row numbers
  let fresh = data; // in-memory rows are valid until the first row actually moves
  moved.forEach((id) => {
    if (placeAsset_(id, fresh)) fresh = null;
  });
  syncStaffs_();
  return msgs;
}

/**
 * Menu item: the reverse of the sync. Reads the IDs typed into the Staffs "<Type> ID" cells
 * and writes that person as the holder in AllAssetLists. Shows a preview and asks first.
 * Devices held in AllAssetLists but not listed in Staffs are left unchanged.
 */
function importHoldersFromStaffs() {
  const ui = SpreadsheetApp.getUi();
  const p = planImport_();
  const lines = [];
  if (p.changes.length)
    lines.push(
      "Will assign " + p.changes.length + " device(s):",
      ...p.changes.map(
        (c) => "  " + c.id + ": " + (c.from || "(in stock)") + " → " + c.to
      )
    );
  if (p.problems.length)
    lines.push(
      "",
      "Skipped (these will disappear from Staffs after the import — note them down to fix):",
      ...p.problems.map((s) => "  " + s)
    );
  if (p.warnings.length)
    lines.push("", "Note:", ...p.warnings.map((s) => "  " + s));
  if (!p.changes.length) {
    ui.alert(
      "Import holders",
      (lines.length ? lines.join("\n") + "\n\n" : "") +
        "Nothing to change — AllAssetLists already matches Staffs.",
      ui.ButtonSet.OK
    );
    return;
  }
  if (
    ui.alert(
      "Import holders",
      lines.join("\n") + "\n\nApply?",
      ui.ButtonSet.YES_NO
    ) !== ui.Button.YES
  )
    return;

  const done = withLock_(() => {
    const plan = planImport_(); // re-read inside the lock in case something changed meanwhile
    const { sh, c, changes } = plan;
    changes.forEach((ch) => {
      if (ch.from) addPrevHolder_(sh, c, ch.asset, ch.from);
      setCell_(sh, ch.row, c.user, ch.to);
      setCell_(sh, ch.row, c.status, "In use");
      log_(
        ch.id,
        ch.from ? "Reassigned" : "Assigned",
        ch.from,
        ch.to,
        ch.condition,
        "Imported from Staffs"
      );
      ch.asset.user = ch.to;
      ch.asset.status = "In use";
    });
    // after all writes: moving shifts row numbers, so in-memory rows are only used until the first move
    let fresh = plan;
    changes.forEach((ch) => {
      if (placeAsset_(ch.id, fresh)) fresh = null;
    });
    syncStaffs_();
    return changes.length;
  });
  ui.alert("Import holders: " + done + " device(s) assigned.");
}

/** Works out what importHoldersFromStaffs would change. */
function planImport_() {
  const staffSh = sheet_(CONFIG.STAFF_SHEET);
  const { sh, c, assets, values } = readAssets_();
  const byId = {};
  assets.forEach((a) => {
    byId[a.id.toLowerCase()] = a;
  });
  const cols = staffIdColumns_(staffSh);
  const staff = readStaff_();
  const lastRow = staffSh.getLastRow();
  const hr = CONFIG.STAFF_HEADER_ROW;
  const problems = [],
    warnings = [];

  // id -> { name, header } as listed in Staffs; ids listed under two people are dropped
  const listed = {},
    clash = {};
  if (lastRow > hr) {
    cols.forEach((col) => {
      const vals = staffSh
        .getRange(hr + 1, col.col, lastRow - hr, 1)
        .getValues();
      staff.forEach((s) => {
        parseIds_(vals[s.row - hr - 1][0]).forEach((id) => {
          const k = id.toLowerCase();
          if (listed[k] && listed[k].name !== s.name)
            (clash[k] = clash[k] || new Set([listed[k].name])).add(s.name);
          else listed[k] = { id, name: s.name, col };
        });
      });
    });
  }

  const changes = [];
  Object.keys(listed).forEach((k) => {
    const l = listed[k],
      a = byId[k];
    if (clash[k])
      return problems.push(
        l.id + " is listed under " + [...clash[k]].join(" and ")
      );
    if (!a)
      return problems.push(
        l.id +
          " (" +
          l.name +
          ", " +
          l.col.header +
          ") is not in " +
          CONFIG.ASSET_SHEET
      );
    if (a.status === "Lost" || a.status === "Retired")
      return problems.push(
        a.id + " is " + a.status + " and cannot be assigned to " + l.name
      );
    if (a.type !== l.col.type)
      warnings.push(
        a.id +
          " is a " +
          (a.type || "device with no type") +
          ' but was typed in "' +
          l.col.header +
          '"'
      );
    const sameHolder = a.user.toLowerCase() === l.name.toLowerCase();
    if (!sameHolder || a.status === "In stock")
      // same holder but "In stock" = status conflict, fix it too
      changes.push({
        id: a.id,
        row: a.row,
        asset: a,
        from: sameHolder ? "" : a.user,
        to: l.name,
        condition: a.condition,
      });
  });
  return { sh, c, assets, values, changes, problems, warnings };
}
