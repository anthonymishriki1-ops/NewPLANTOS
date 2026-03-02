
/***************************************
 * PlantOS — Code.gs (Backend + Menus) — “Don’t break my UI” Edition
 *
 * Scope (locked):
 * - DOES NOT touch your HTML files.
 * - DOES NOT define doGet (so it cannot override your existing routing).
 * - Gives you the exact menu flow you asked for:
 *     1) Set Web App URL (manual prompt)
 *     2) Wipe previous deployment fields (IDs/URLs only)
 *     3) Rebuild links + folders + care docs + QR (chunked)
 *     4) Continue rebuild (resume cursor)
 *     5) STOP
 *     + Diagnostics
 *
 * Canonical Drive folder scheme: UID_<uid>
 * - This is best for your search-engine/indexing idea because UID is stable forever.
 *
 * IMPORTANT:
 * - Remove/disable any OTHER file that defines these functions:
 *     plantosUploadPlantPhoto, plantosGetLatestPhoto, onOpen
 *   Duplicates = override roulette.
 ***************************************/

const PLANTOS_BACKEND_CFG = {
  INVENTORY_SHEET: 'Plant Care Tracking + Inventory',
  SETTINGS_SHEET: 'PlantOS Settings',

  SETTINGS_KEYS: {
    ACTIVE_WEBAPP_URL: 'ACTIVE_WEBAPP_URL',
    REBUILD_CURSOR: 'REBUILD_CURSOR',

    DRIVE_ROOT_ID: 'DRIVE_ROOT_ID',
    DRIVE_PLANTS_ID: 'DRIVE_PLANTS_ID',
    DRIVE_QR_ID: 'DRIVE_QR_ID',
  },

  DRIVE_NAMES: {
    ROOT: 'PlantOS',
    PLANTS: 'Plants',
    QR: 'QR - Plant Pages',
  },

  CANONICAL_PLANT_FOLDER_PREFIX: 'UID_',
  PHOTOS_SUBFOLDER: 'Photos',

  REBUILD_CHUNK: 35,

  QR: {
    SIZE: '320x320',
    API: 'https://api.qrserver.com/v1/create-qr-code/',
  },

  // Column headers we will read/write if present (we do NOT auto-add columns)
  HEADERS: {
    UID: 'Plant UID',
    NICKNAME: 'Nick-name',
    GENUS: 'Genus',
    TAXON: 'Taxon Raw',
    LOCATION: 'Location',
    PLANT_ID: 'Plant ID',

    FOLDER_ID: 'Folder ID',
    FOLDER_URL: 'Folder URL',
    CARE_DOC_ID: 'Care Doc ID',
    CARE_DOC_URL: 'Care Doc URL',
    QR_FILE_ID: 'QR File ID',
    QR_URL: 'QR URL',
    PLANT_PAGE_URL: 'Plant Page URL',

    LAST_WATERED: 'Last Watered',
    WATER_EVERY_DAYS: 'Water Every Days',
    WATERED: 'Watered',

    LAST_FERTILIZED: 'Last Fertilized',
    FERT_EVERY_DAYS: 'Fertilize Every Days',
    FERTILIZED: 'Fertilized',

    POT_SIZE: 'Pot Size',
    MEDIUM: 'Medium',
    BIRTHDAY: 'Birthday',

    // Optional photo cache columns (nice-to-have)
    LATEST_PHOTO_ID: 'Latest Photo ID',
    LATEST_PHOTO_THUMB: 'Latest Photo Thumb',
    LATEST_PHOTO_VIEW: 'Latest Photo View',
    LATEST_PHOTO_UPDATED: 'Latest Photo Updated',
  }
};

/* ===================== MENU ===================== */

function onOpen() {
  plantosBuildMenu_();
}

function plantosBuildMenu_() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('PlantOS')
    .addItem('Set Web App URL (manual)', 'plantosMenuSetWebAppUrlManual')
    .addItem('Confirm Web App URL (auto)', 'plantosMenuConfirmWebAppUrlAuto')
    .addSeparator()
    .addItem('Wipe Previous Deployments (IDs/URLs)', 'plantosMenuWipeDeploymentFields')
    .addItem('Rebuild Deployments (links/folders/docs/QR)', 'plantosMenuRebuildStart')
    .addItem('Continue Rebuild (resume)', 'plantosMenuRebuildContinue')
    .addSeparator()
    .addItem('STOP (clear rebuild cursor)', 'plantosMenuStop')
    .addSeparator()
    .addItem('Diagnostics (sanity check)', 'plantosMenuDiagnostics')
    .addToUi();
}

function plantosMenuSetWebAppUrlManual() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt(
    'Set ACTIVE_WEBAPP_URL',
    'Paste your deployed Web App URL (the one ending in /exec).\nExample: https://script.google.com/macros/s/…/exec',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  const url = String(resp.getResponseText() || '').trim();
  const ok = plantosValidateWebAppUrl_(url);
  if (!ok.ok) {
    ui.alert('Nope.\n' + ok.reason);
    return;
  }

  plantosSetSetting_(PLANTOS_BACKEND_CFG.SETTINGS_KEYS.ACTIVE_WEBAPP_URL, url);
  ui.alert('Saved ACTIVE_WEBAPP_URL:\n' + url);
}

function plantosMenuConfirmWebAppUrlAuto() {
  const ui = SpreadsheetApp.getUi();
  const url = plantosGetCurrentWebAppUrl_();
  if (!url) {
    ui.alert('Could not detect current Web App URL.\nUse "Set Web App URL (manual)" instead.');
    return;
  }

  const ok = plantosValidateWebAppUrl_(url);
  if (!ok.ok) {
    ui.alert('Auto-detected URL does not look like a deployed /exec URL.\nUse manual entry.\n\nDetected:\n' + url);
    return;
  }

  plantosSetSetting_(PLANTOS_BACKEND_CFG.SETTINGS_KEYS.ACTIVE_WEBAPP_URL, url);
  ui.alert('Saved ACTIVE_WEBAPP_URL:\n' + url);
}

function plantosMenuWipeDeploymentFields() {
  plantosWipeDeploymentFields_();
  SpreadsheetApp.getUi().alert('Done.\nDeployment IDs/URLs cleared from the sheet.\n(Drive folders/files are NOT deleted.)');
}

function plantosMenuRebuildStart() {
  const res = plantosRebuildDeploymentAssets_({ resume: false });
  SpreadsheetApp.getUi().alert(res.message);
}

function plantosMenuRebuildContinue() {
  const res = plantosRebuildDeploymentAssets_({ resume: true });
  SpreadsheetApp.getUi().alert(res.message);
}

function plantosMenuStop() {
  plantosSetSetting_(PLANTOS_BACKEND_CFG.SETTINGS_KEYS.REBUILD_CURSOR, '');
  SpreadsheetApp.getUi().alert('Stopped.\nRebuild cursor cleared.');
}

function plantosMenuDiagnostics() {
  const report = plantosDiagnostics_();
  SpreadsheetApp.getUi().alert(report);
}

/* ===================== SETTINGS ===================== */

function plantosGetSetting_(key) {
  const sh = plantosGetSheet_(PLANTOS_BACKEND_CFG.SETTINGS_SHEET);
  const values = sh.getDataRange().getValues();
  for (let r = 1; r < values.length; r++) {
    const k = String(values[r][0] || '').trim();
    if (plantosNorm_(k) === plantosNorm_(key)) return String(values[r][1] || '').trim();
  }
  return '';
}

function plantosSetSetting_(key, value) {
  const sh = plantosGetSheet_(PLANTOS_BACKEND_CFG.SETTINGS_SHEET);
  const values = sh.getDataRange().getValues();
  for (let r = 1; r < values.length; r++) {
    const k = String(values[r][0] || '').trim();
    if (plantosNorm_(k) === plantosNorm_(key)) {
      sh.getRange(r + 1, 2).setValue(value);
      return;
    }
  }
  sh.appendRow([key, value]);
}

/* ===================== SHEET HELPERS ===================== */

function plantosGetSS_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function plantosGetSheet_(name) {
  const sh = plantosGetSS_().getSheetByName(name);
  if (!sh) throw new Error('Missing sheet: ' + name);
  return sh;
}

function plantosGetInventorySheet_() {
  return plantosGetSheet_(PLANTOS_BACKEND_CFG.INVENTORY_SHEET);
}

function plantosReadInventory_() {
  const sh = plantosGetInventorySheet_();
  const values = sh.getDataRange().getValues();
  const headers = values[0] || [];
  const hmap = plantosHeaderMap_(headers);
  return { sh, values, headers, hmap };
}

function plantosNorm_(s) {
  return String(s == null ? '' : s).trim().toLowerCase();
}

function plantosHeaderMap_(headers) {
  const map = {};
  headers.forEach((h, i) => {
    const k = plantosNorm_(h);
    if (!k) return;
    if (!(k in map)) map[k] = i;
  });
  return map;
}

function plantosCol_(hmap, headerName) {
  const idx = hmap[plantosNorm_(headerName)];
  return (typeof idx === 'number') ? idx : -1;
}

function plantosSafeStr_(v) {
  return (v == null) ? '' : String(v);
}

function plantosAsDate_(v) {
  if (!v) return null;
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) return v;
  const d = new Date(v);
  return isNaN(d) ? null : d;
}

function plantosFmtDate_(d) {
  if (!d) return '';
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function plantosAddDays_(d, days) {
  const x = new Date(d.getTime());
  x.setDate(x.getDate() + days);
  return x;
}

function plantosNow_() {
  return new Date();
}

/* ===================== URL ===================== */

function plantosGetCurrentWebAppUrl_() {
  try {
    return ScriptApp.getService().getUrl() || '';
  } catch (e) {
    return '';
  }
}

function plantosValidateWebAppUrl_(url) {
  const u = String(url || '').trim();
  if (!u) return { ok: false, reason: 'Empty URL.' };
  if (!u.startsWith('https://script.google.com/macros/s/')) {
    return { ok: false, reason: 'URL must start with https://script.google.com/macros/s/' };
  }
  if (!(u.includes('/exec') || u.includes('/dev'))) {
    return { ok: false, reason: 'URL should contain /exec (or /dev for test).' };
  }
  return { ok: true };
}

function plantosBuildPlantPageUrl_(baseUrl, uid) {
  const u = String(baseUrl || '').trim();
  const id = encodeURIComponent(String(uid || '').trim());
  if (!u || !id) return '';
  // Strip any existing querystring or fragment from the base URL.
  const clean = u.split('?')[0].split('#')[0];
  // Always append our own query string.
  return `${clean}?mode=uid${id}`;
}

/* ===================== WIPE DEPLOYMENT FIELDS ===================== */

function plantosWipeDeploymentFields_() {
  const { sh, hmap } = plantosReadInventory_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const H = PLANTOS_BACKEND_CFG.HEADERS;
  const targets = [
    H.FOLDER_ID, H.FOLDER_URL,
    H.CARE_DOC_ID, H.CARE_DOC_URL,
    H.QR_FILE_ID, H.QR_URL,
    H.PLANT_PAGE_URL
  ];

  const colIdxs = targets.map(h => plantosCol_(hmap, h)).filter(i => i >= 0);
  colIdxs.forEach(ci => sh.getRange(2, ci + 1, lastRow - 1, 1).clearContent());

  plantosSetSetting_(PLANTOS_BACKEND_CFG.SETTINGS_KEYS.REBUILD_CURSOR, '');
}

/* ===================== DRIVE ROOTS ===================== */

function plantosGetOrCreateRootFolder_() {
  const key = PLANTOS_BACKEND_CFG.SETTINGS_KEYS.DRIVE_ROOT_ID;
  const existingId = plantosGetSetting_(key);
  if (existingId) {
    try {
      return DriveApp.getFolderById(existingId);
    } catch (e) {
      // fall through
    }
  }

  const name = PLANTOS_BACKEND_CFG.DRIVE_NAMES.ROOT;
  const it = DriveApp.getFoldersByName(name);
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder(name);
  plantosSetSetting_(key, folder.getId());
  return folder;
}

function plantosGetOrCreateChildFolder_(parent, name, settingsKey) {
  if (settingsKey) {
    const existingId = plantosGetSetting_(settingsKey);
    if (existingId) {
      try {
        return DriveApp.getFolderById(existingId);
      } catch (e) {
        // fall through
      }
    }
  }

  const it = parent.getFoldersByName(name);
  const folder = it.hasNext() ? it.next() : parent.createFolder(name);

  if (settingsKey) plantosSetSetting_(settingsKey, folder.getId());
  return folder;
}

function plantosGetPlantsRoot_() {
  const root = plantosGetOrCreateRootFolder_();
  return plantosGetOrCreateChildFolder_(root, PLANTOS_BACKEND_CFG.DRIVE_NAMES.PLANTS, PLANTOS_BACKEND_CFG.SETTINGS_KEYS.DRIVE_PLANTS_ID);
}

function plantosGetQrRoot_() {
  const root = plantosGetOrCreateRootFolder_();
  return plantosGetOrCreateChildFolder_(root, PLANTOS_BACKEND_CFG.DRIVE_NAMES.QR, PLANTOS_BACKEND_CFG.SETTINGS_KEYS.DRIVE_QR_ID);
}

function plantosEnsureSubfolder_(folder, name) {
  const it = folder.getFoldersByName(name);
  return it.hasNext() ? it.next() : folder.createFolder(name);
}

/* ===================== FOLDER RESOLUTION (Canonical UID_<uid>) ===================== */

function plantosCanonicalFolderName_(uid) {
  return PLANTOS_BACKEND_CFG.CANONICAL_PLANT_FOLDER_PREFIX + String(uid || '').trim();
}

function plantosResolveOrCreatePlantFolder_(plantsRootFolder, uid) {
  const canonicalName = plantosCanonicalFolderName_(uid);

  // 1) Exact canonical name
  let it = plantsRootFolder.getFoldersByName(canonicalName);
  if (it.hasNext()) return it.next();

  // 2) Try to find legacy folder that starts with "<uid> —" or "<uid> -"
  const uidStr = String(uid || '').trim();
  const all = plantsRootFolder.getFolders();
  while (all.hasNext()) {
    const f = all.next();
    const n = String(f.getName() || '');
    if (n === uidStr) {
      try { f.setName(canonicalName); } catch (e) {}
      return f;
    }
    if (n.startsWith(uidStr + ' —') || n.startsWith(uidStr + ' -') || n.startsWith(uidStr + '—') || n.startsWith(uidStr + '-')) {
      try { f.setName(canonicalName); } catch (e) {}
      return f;
    }
  }

  // 3) Create new canonical folder
  const created = plantsRootFolder.createFolder(canonicalName);
  return created;
}

/* ===================== DEPLOYMENT REBUILD ===================== */

function plantosRebuildDeploymentAssets_(opts) {
  opts = opts || {};
  const resume = !!opts.resume;

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(15000)) {
    return { ok: false, message: 'Another PlantOS job is running. Try again in a moment.' };
  }

  try {
    const baseUrl = plantosGetSetting_(PLANTOS_BACKEND_CFG.SETTINGS_KEYS.ACTIVE_WEBAPP_URL);
    const ok = plantosValidateWebAppUrl_(baseUrl);
    if (!ok.ok) return { ok: false, message: 'ACTIVE_WEBAPP_URL not set or invalid.\nUse "Set Web App URL (manual)".\n\n' + ok.reason };

    const plantsRoot = plantosGetPlantsRoot_();
    const qrRoot = plantosGetQrRoot_();

    const { sh, headers, hmap } = plantosReadInventory_();
    const H = PLANTOS_BACKEND_CFG.HEADERS;

    const uidCol = plantosCol_(hmap, H.UID);
    if (uidCol < 0) return { ok: false, message: `Missing required header: "${H.UID}"` };

    const lastRow = sh.getLastRow();
    if (lastRow < 2) return { ok: true, message: 'No plants to rebuild (sheet has no rows).' };

    let cursor = 2;
    if (resume) {
      const c = Number(plantosGetSetting_(PLANTOS_BACKEND_CFG.SETTINGS_KEYS.REBUILD_CURSOR) || '');
      if (c && c >= 2) cursor = c;
    }

    if (cursor > lastRow) {
      plantosSetSetting_(PLANTOS_BACKEND_CFG.SETTINGS_KEYS.REBUILD_CURSOR, '');
      return { ok: true, message: 'Nothing to do. Cursor already past end.' };
    }

    const start = cursor;
    const end = Math.min(lastRow, start + PLANTOS_BACKEND_CFG.REBUILD_CHUNK - 1);

    const range = sh.getRange(start, 1, end - start + 1, headers.length);
    const block = range.getValues();

    const folderIdCol = plantosCol_(hmap, H.FOLDER_ID);
    const folderUrlCol = plantosCol_(hmap, H.FOLDER_URL);
    const careDocIdCol = plantosCol_(hmap, H.CARE_DOC_ID);
    const careDocUrlCol = plantosCol_(hmap, H.CARE_DOC_URL);
    const qrFileIdCol = plantosCol_(hmap, H.QR_FILE_ID);
    const qrUrlCol = plantosCol_(hmap, H.QR_URL);
    const plantPageUrlCol = plantosCol_(hmap, H.PLANT_PAGE_URL);

    for (let i = 0; i < block.length; i++) {
      const row = block[i];
      const uid = plantosSafeStr_(row[uidCol]).trim();
      if (!uid) continue;

      const primary = plantosComputePrimaryLabel_(hmap, row);
      const plantPageUrl = plantosBuildPlantPageUrl_(baseUrl, uid);

      // Plant Page URL
      if (plantPageUrlCol >= 0 && !plantosSafeStr_(row[plantPageUrlCol]).trim()) {
        row[plantPageUrlCol] = plantPageUrl;
      }

      // Resolve or create plant folder (prefer sheet Folder ID if valid)
      let plantFolder = null;
      if (folderIdCol >= 0) {
        const fid = plantosSafeStr_(row[folderIdCol]).trim();
        if (fid) {
          try {
            plantFolder = DriveApp.getFolderById(fid);
          } catch (e) {
            plantFolder = null;
          }
        }
      }
      if (!plantFolder) {
        plantFolder = plantosResolveOrCreatePlantFolder_(plantsRoot, uid);
        try {
          // Enforce canonical name (best-effort)
          const canonicalName = plantosCanonicalFolderName_(uid);
          if (plantFolder.getName() !== canonicalName) plantFolder.setName(canonicalName);
        } catch (e) {}

        if (folderIdCol >= 0) row[folderIdCol] = plantFolder.getId();
        if (folderUrlCol >= 0) row[folderUrlCol] = plantFolder.getUrl();
      } else {
        // Best-effort canonical rename if user wants stable indexing
        try {
          const canonicalName = plantosCanonicalFolderName_(uid);
          if (plantFolder.getName() !== canonicalName) plantFolder.setName(canonicalName);
        } catch (e) {}
        if (folderUrlCol >= 0 && !plantosSafeStr_(row[folderUrlCol]).trim()) row[folderUrlCol] = plantFolder.getUrl();
      }

      // Ensure Photos subfolder
      plantosEnsureSubfolder_(plantFolder, PLANTOS_BACKEND_CFG.PHOTOS_SUBFOLDER);

      // Care Doc
      if (careDocIdCol >= 0 && !plantosSafeStr_(row[careDocIdCol]).trim()) {
        const docFile = plantosEnsureCareDoc_(plantFolder, uid, primary);
        row[careDocIdCol] = docFile.getId();
        if (careDocUrlCol >= 0) row[careDocUrlCol] = docFile.getUrl();
      } else if (careDocUrlCol >= 0 && careDocIdCol >= 0 && plantosSafeStr_(row[careDocIdCol]).trim() && !plantosSafeStr_(row[careDocUrlCol]).trim()) {
        try {
          row[careDocUrlCol] = DriveApp.getFileById(String(row[careDocIdCol])).getUrl();
        } catch (e) {}
      }

      // QR
      if (qrFileIdCol >= 0 && !plantosSafeStr_(row[qrFileIdCol]).trim()) {
        const qrFile = plantosEnsurePlantQr_(qrRoot, uid, primary, plantPageUrl);
        row[qrFileIdCol] = qrFile.getId();
        if (qrUrlCol >= 0) row[qrUrlCol] = qrFile.getUrl();
      } else if (qrUrlCol >= 0 && qrFileIdCol >= 0 && plantosSafeStr_(row[qrFileIdCol]).trim() && !plantosSafeStr_(row[qrUrlCol]).trim()) {
        try {
          row[qrUrlCol] = DriveApp.getFileById(String(row[qrFileIdCol])).getUrl();
        } catch (e) {}
      }

      block[i] = row;
    }

    range.setValues(block);

    const nextCursor = end + 1;
    if (nextCursor <= lastRow) {
      plantosSetSetting_(PLANTOS_BACKEND_CFG.SETTINGS_KEYS.REBUILD_CURSOR, nextCursor);
      return { ok: true, message: `Rebuilt rows ${start}–${end}.\nRun "Continue Rebuild" to finish.` };
    }

    plantosSetSetting_(PLANTOS_BACKEND_CFG.SETTINGS_KEYS.REBUILD_CURSOR, '');
    return { ok: true, message: `Rebuilt rows ${start}–${end}.\nAll done ✅` };

  } catch (e) {
    return { ok: false, message: 'Rebuild failed: ' + (e && e.message ? e.message : e) };
  } finally {
    lock.releaseLock();
  }
}

function plantosComputePrimaryLabel_(hmap, row) {
  const H = PLANTOS_BACKEND_CFG.HEADERS;
  const nicknameCol = plantosCol_(hmap, H.NICKNAME);
  const genusCol = plantosCol_(hmap, H.GENUS);
  const taxonCol = plantosCol_(hmap, H.TAXON);
  const plantIdCol = plantosCol_(hmap, H.PLANT_ID);

  const nn = nicknameCol >= 0 ? plantosSafeStr_(row[nicknameCol]).trim() : '';
  if (nn) return nn;

  const genus = genusCol >= 0 ? plantosSafeStr_(row[genusCol]).trim() : '';
  const taxon = taxonCol >= 0 ? plantosSafeStr_(row[taxonCol]).trim() : '';
  const combo = [genus, taxon].filter(Boolean).join(' ').trim();
  if (combo) return combo;

  const pid = plantIdCol >= 0 ? plantosSafeStr_(row[plantIdCol]).trim() : '';
  return pid ? `Plant ${pid}` : 'Plant';
}

function plantosEnsureCareDoc_(plantFolder, uid, primary) {
  const canonical = plantosCanonicalFolderName_(uid);
  const desiredPrefix = `Care Log — ${canonical}`;
  const files = plantFolder.getFiles();

  while (files.hasNext()) {
    const f = files.next();
    if (f.getMimeType && f.getMimeType() === MimeType.GOOGLE_DOCS) {
      const name = String(f.getName() || '');
      if (name.startsWith(desiredPrefix) || name.includes(canonical)) return f;
      if (name.startsWith('Care Log')) return f; // fallback
    }
  }

  const title = `${desiredPrefix} — ${primary}`.substring(0, 180);
  const doc = DocumentApp.create(title);
  const body = doc.getBody();
  body.appendParagraph('PlantOS Care Log').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph(`UID: ${uid}`);
  body.appendParagraph(`Primary: ${primary}`);
  body.appendParagraph('');
  body.appendParagraph('Entries:').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  doc.saveAndClose();

  const file = DriveApp.getFileById(doc.getId());
  plantFolder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);
  return file;
}

function plantosEnsurePlantQr_(qrRootFolder, uid, primary, plantPageUrl) {
  const canonical = plantosCanonicalFolderName_(uid);
  const filename = `QR_${canonical}.png`;

  const it = qrRootFolder.getFilesByName(filename);
  if (it.hasNext()) return it.next();

  const url = `${PLANTOS_BACKEND_CFG.QR.API}?size=${encodeURIComponent(PLANTOS_BACKEND_CFG.QR.SIZE)}&data=${encodeURIComponent(plantPageUrl)}`;
  const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  const code = resp.getResponseCode();
  if (code < 200 || code >= 300) throw new Error('QR fetch failed: ' + code);

  const blob = resp.getBlob().setName(filename);
  const f = qrRootFolder.createFile(blob);
  f.setDescription(`PlantOS QR for ${primary} (${uid})`);

  // Optional sharing (not required but handy)
  try {
    f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {}

  return f;
}

/* ===================== DIAGNOSTICS ===================== */

function plantosDiagnostics_() {
  const lines = [];
  const url = plantosGetSetting_(PLANTOS_BACKEND_CFG.SETTINGS_KEYS.ACTIVE_WEBAPP_URL);
  const cursor = plantosGetSetting_(PLANTOS_BACKEND_CFG.SETTINGS_KEYS.REBUILD_CURSOR);

  lines.push('ACTIVE_WEBAPP_URL: ' + (url || '(not set)'));
  lines.push('REBUILD_CURSOR: ' + (cursor || '(none)'));

  // Drive roots
  try {
    const root = plantosGetOrCreateRootFolder_();
    const plants = plantosGetPlantsRoot_();
    const qr = plantosGetQrRoot_();
    lines.push('Drive ROOT: OK (' + root.getName() + ')');
    lines.push('Drive Plants: OK (' + plants.getName() + ')');
    lines.push('Drive QR: OK (' + qr.getName() + ')');
  } catch (e) {
    lines.push('Drive Roots: ERROR: ' + (e && e.message ? e.message : e));
  }

  // Inventory checks
  try {
    const { sh, hmap } = plantosReadInventory_();
    const H = PLANTOS_BACKEND_CFG.HEADERS;
    const uidCol = plantosCol_(hmap, H.UID);
    lines.push('Inventory Sheet: OK (' + sh.getName() + ')');
    lines.push('Plant UID col: ' + (uidCol >= 0 ? 'OK' : 'MISSING'));

    // Count missing key deployment fields (best-effort)
    if (uidCol >= 0) {
      const lastRow = sh.getLastRow();
      const data = (lastRow >= 2) ? sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues() : [];
      const fidCol = plantosCol_(hmap, H.FOLDER_ID);
      const ppCol = plantosCol_(hmap, H.PLANT_PAGE_URL);
      const qrCol = plantosCol_(hmap, H.QR_FILE_ID);

      let missingFolder = 0, missingPage = 0, missingQr = 0, total = 0;
      data.forEach(row => {
        const uid = plantosSafeStr_(row[uidCol]).trim();
        if (!uid) return;
        total++;
        if (fidCol >= 0 && !plantosSafeStr_(row[fidCol]).trim()) missingFolder++;
        if (ppCol >= 0 && !plantosSafeStr_(row[ppCol]).trim()) missingPage++;
        if (qrCol >= 0 && !plantosSafeStr_(row[qrCol]).trim()) missingQr++;
      });

      lines.push('Rows w/UID: ' + total);
      if (fidCol >= 0) lines.push('Missing Folder ID: ' + missingFolder);
      if (ppCol >= 0) lines.push('Missing Plant Page URL: ' + missingPage);
      if (qrCol >= 0) lines.push('Missing QR File ID: ' + missingQr);
    }
  } catch (e) {
    lines.push('Inventory: ERROR: ' + (e && e.message ? e.message : e));
  }

  return lines.join('\n');
}

/* ===================== API USED BY YOUR HTML ===================== */

function plantosListLocations() {
  const { values, hmap } = plantosReadInventory_();
  const locCol = plantosCol_(hmap, PLANTOS_BACKEND_CFG.HEADERS.LOCATION);
  if (locCol < 0) return [];
  const set = {};
  for (let r = 1; r < values.length; r++) {
    const loc = plantosSafeStr_(values[r][locCol]).trim();
    if (loc) set[loc] = true;
  }
  return Object.keys(set).sort((a, b) => a.localeCompare(b));
}

function plantosHome() {
  const { values, hmap } = plantosReadInventory_();
  const H = PLANTOS_BACKEND_CFG.HEADERS;

  const uidCol = plantosCol_(hmap, H.UID);
  const nicknameCol = plantosCol_(hmap, H.NICKNAME);
  const genusCol = plantosCol_(hmap, H.GENUS);
  const taxonCol = plantosCol_(hmap, H.TAXON);

  const lastWateredCol = plantosCol_(hmap, H.LAST_WATERED);
  const everyDaysCol = plantosCol_(hmap, H.WATER_EVERY_DAYS);
  const birthdayCol = plantosCol_(hmap, H.BIRTHDAY);

  const now = plantosNow_();
  const today = Utilities.formatDate(now, Session.getScriptTimeZone(), 'MM/dd');

  const dueNow = [];
  const upcoming = [];
  const birthdays = [];

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const uid = uidCol >= 0 ? plantosSafeStr_(row[uidCol]).trim() : '';
    if (!uid) continue;

    const nn = nicknameCol >= 0 ? plantosSafeStr_(row[nicknameCol]).trim() : '';
    const genus = genusCol >= 0 ? plantosSafeStr_(row[genusCol]).trim() : '';
    const taxon = taxonCol >= 0 ? plantosSafeStr_(row[taxonCol]).trim() : '';
    const primary = nn || [genus, taxon].filter(Boolean).join(' ') || uid;

    if (birthdayCol >= 0) {
      const bd = plantosAsDate_(row[birthdayCol]);
      if (bd) {
        const mmdd = Utilities.formatDate(bd, Session.getScriptTimeZone(), 'MM/dd');
        if (mmdd === today) birthdays.push(primary);
      }
    }

    const every = everyDaysCol >= 0 ? Number(row[everyDaysCol]) : NaN;
    const lw = lastWateredCol >= 0 ? plantosAsDate_(row[lastWateredCol]) : null;

    if (!isNaN(every) && every > 0) {
      if (!lw) {
        dueNow.push({ uid, primary, due: 'unknown' });
      } else {
        const dueDate = plantosAddDays_(lw, every);
        if (dueDate <= now) dueNow.push({ uid, primary, due: plantosFmtDate_(dueDate) });
        else {
          const diffDays = Math.ceil((dueDate.getTime() - now.getTime()) / (24 * 3600 * 1000));
          if (diffDays >= 1 && diffDays <= 7) upcoming.push({ uid, primary, due: plantosFmtDate_(dueDate) });
        }
      }
    }
  }

  dueNow.sort((a, b) => String(a.due || '').localeCompare(String(b.due || '')));
  upcoming.sort((a, b) => String(a.due || '').localeCompare(String(b.due || '')));

  return { dueNow, upcoming, birthdays };
}

function plantosGetPlantsByLocation(location) {
  const loc = plantosSafeStr_(location).trim();
  const { values, hmap } = plantosReadInventory_();

  const uidCol = plantosCol_(hmap, PLANTOS_BACKEND_CFG.HEADERS.UID);
  const locCol = plantosCol_(hmap, PLANTOS_BACKEND_CFG.HEADERS.LOCATION);
  if (uidCol < 0 || locCol < 0) return [];

  const out = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    if (plantosSafeStr_(row[locCol]).trim() !== loc) continue;
    const uid = plantosSafeStr_(row[uidCol]).trim();
    if (!uid) continue;
    out.push(plantosRowToPlant_(hmap, row));
  }
  return out;
}

function plantosGetPlant(uid) {
  const needle = plantosSafeStr_(uid).trim();
  if (!needle) return { ok: false, reason: 'Missing uid' };

  const { values, hmap } = plantosReadInventory_();
  const uidCol = plantosCol_(hmap, PLANTOS_BACKEND_CFG.HEADERS.UID);
  if (uidCol < 0) return { ok: false, reason: 'Missing Plant UID column' };

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    if (plantosSafeStr_(row[uidCol]).trim() === needle) {
      const plant = plantosRowToPlant_(hmap, row);
      plant._rowNumber = r + 1;
      return { ok: true, plant };
    }
  }
  return { ok: false, reason: 'Not found' };
}

function plantosSetNickname(uid, nickname) {
  const needle = plantosSafeStr_(uid).trim();
  const nn = plantosSafeStr_(nickname).trim();
  if (!needle) throw new Error('Missing uid');

  const { sh, values, hmap } = plantosReadInventory_();
  const uidCol = plantosCol_(hmap, PLANTOS_BACKEND_CFG.HEADERS.UID);
  const nicknameCol = plantosCol_(hmap, PLANTOS_BACKEND_CFG.HEADERS.NICKNAME);
  if (uidCol < 0 || nicknameCol < 0) throw new Error('Missing columns');

  for (let r = 1; r < values.length; r++) {
    if (plantosSafeStr_(values[r][uidCol]).trim() === needle) {
      sh.getRange(r + 1, nicknameCol + 1).setValue(nn);
      return { ok: true };
    }
  }
  throw new Error('Plant not found');
}

function plantosUpdatePlant(uid, patch) {
  const needle = plantosSafeStr_(uid).trim();
  if (!needle) throw new Error('Missing uid');
  patch = patch || {};

  const { sh, values, hmap } = plantosReadInventory_();
  const uidCol = plantosCol_(hmap, PLANTOS_BACKEND_CFG.HEADERS.UID);
  if (uidCol < 0) throw new Error('Missing Plant UID');

  const writable = {
    potSize: PLANTOS_BACKEND_CFG.HEADERS.POT_SIZE,
    substrate: PLANTOS_BACKEND_CFG.HEADERS.MEDIUM,
    location: PLANTOS_BACKEND_CFG.HEADERS.LOCATION,
    birthday: PLANTOS_BACKEND_CFG.HEADERS.BIRTHDAY,
    waterEveryDays: PLANTOS_BACKEND_CFG.HEADERS.WATER_EVERY_DAYS,
  };

  for (let r = 1; r < values.length; r++) {
    if (plantosSafeStr_(values[r][uidCol]).trim() !== needle) continue;

    Object.keys(writable).forEach(k => {
      if (!(k in patch)) return;
      const colName = writable[k];
      const c = plantosCol_(hmap, colName);
      if (c >= 0) sh.getRange(r + 1, c + 1).setValue(patch[k]);
    });

    return { ok: true };
  }

  throw new Error('Plant not found');
}

function plantosCreatePlant(payload) {
  payload = payload || {};
  const { sh, headers, hmap } = plantosReadInventory_();

  const uidCol = plantosCol_(hmap, PLANTOS_BACKEND_CFG.HEADERS.UID);
  if (uidCol < 0) throw new Error('Missing Plant UID');

  const uid = plantosGenerateNextUid_();
  const row = new Array(headers.length).fill('');
  row[uidCol] = uid;

  const setIf = (colName, val) => {
    const c = plantosCol_(hmap, colName);
    if (c >= 0) row[c] = val;
  };

  setIf(PLANTOS_BACKEND_CFG.HEADERS.GENUS, payload.genus || '');
  setIf(PLANTOS_BACKEND_CFG.HEADERS.TAXON, payload.taxonRaw || payload.taxon || '');
  setIf(PLANTOS_BACKEND_CFG.HEADERS.LOCATION, payload.location || '');
  setIf(PLANTOS_BACKEND_CFG.HEADERS.NICKNAME, payload.nickname || '');
  setIf(PLANTOS_BACKEND_CFG.HEADERS.MEDIUM, payload.medium || payload.substrate || '');
  setIf(PLANTOS_BACKEND_CFG.HEADERS.BIRTHDAY, payload.birthday || '');

  sh.appendRow(row);
  return { ok: true, uid };
}

function plantosGenerateNextUid_() {
  const { values, hmap } = plantosReadInventory_();
  const uidCol = plantosCol_(hmap, PLANTOS_BACKEND_CFG.HEADERS.UID);
  let max = 0;
  for (let r = 1; r < values.length; r++) {
    const v = plantosSafeStr_(values[r][uidCol]).trim();
    const n = Number(v);
    if (!isNaN(n)) max = Math.max(max, n);
  }
  if (max > 0) return String(max + 1);
  return String(Date.now());
}

function plantosQuickLog(uid, payload) {
  const needle = plantosSafeStr_(uid).trim();
  payload = payload || {};
  if (!needle) throw new Error('Missing uid');

  const { sh, values, hmap } = plantosReadInventory_();
  const uidCol = plantosCol_(hmap, PLANTOS_BACKEND_CFG.HEADERS.UID);
  if (uidCol < 0) throw new Error('Missing Plant UID');

  const wateredCol = plantosCol_(hmap, PLANTOS_BACKEND_CFG.HEADERS.WATERED);
  const lastWateredCol = plantosCol_(hmap, PLANTOS_BACKEND_CFG.HEADERS.LAST_WATERED);
  const fertilizedCol = plantosCol_(hmap, PLANTOS_BACKEND_CFG.HEADERS.FERTILIZED);
  const lastFertilizedCol = plantosCol_(hmap, PLANTOS_BACKEND_CFG.HEADERS.LAST_FERTILIZED);

  const now = plantosNow_();

  for (let r = 1; r < values.length; r++) {
    if (plantosSafeStr_(values[r][uidCol]).trim() !== needle) continue;

    if (payload.water === true) {
      if (wateredCol >= 0) sh.getRange(r + 1, wateredCol + 1).setValue(true);
      if (lastWateredCol >= 0) sh.getRange(r + 1, lastWateredCol + 1).setValue(now);
    }
    if (payload.fertilize === true) {
      if (fertilizedCol >= 0) sh.getRange(r + 1, fertilizedCol + 1).setValue(true);
      if (lastFertilizedCol >= 0) sh.getRange(r + 1, lastFertilizedCol + 1).setValue(now);
    }

    plantosTimelineAppend_(needle, payload, now);
    return { ok: true };
  }

  throw new Error('Plant not found');
}

function plantosBatchWater(uids, actionLabel) {
  uids = uids || [];
  if (!Array.isArray(uids) || !uids.length) return { ok: true, count: 0 };

  const label = plantosSafeStr_(actionLabel).trim();

  const { sh, values, hmap } = plantosReadInventory_();
  const uidCol = plantosCol_(hmap, PLANTOS_BACKEND_CFG.HEADERS.UID);
  const wateredCol = plantosCol_(hmap, PLANTOS_BACKEND_CFG.HEADERS.WATERED);
  const lastWateredCol = plantosCol_(hmap, PLANTOS_BACKEND_CFG.HEADERS.LAST_WATERED);
  if (uidCol < 0) throw new Error('Missing Plant UID');

  const set = {};
  uids.forEach(u => { const k = plantosSafeStr_(u).trim(); if (k) set[k] = true; });

  const now = plantosNow_();
  let count = 0;

  for (let r = 1; r < values.length; r++) {
    const uid = plantosSafeStr_(values[r][uidCol]).trim();
    if (!uid || !set[uid]) continue;

    if (wateredCol >= 0) sh.getRange(r + 1, wateredCol + 1).setValue(true);
    if (lastWateredCol >= 0) sh.getRange(r + 1, lastWateredCol + 1).setValue(now);

    plantosTimelineAppend_(uid, label ? { water: true, notes: label } : { water: true }, now);
    count++;
  }

  return { ok: true, count };
}function plantosSearch(q, limit) {
  const query = plantosNorm_(q);
  limit = Number(limit || 15);
  if (!query) return [];

  const { values, hmap } = plantosReadInventory_();
  const uidCol = plantosCol_(hmap, PLANTOS_BACKEND_CFG.HEADERS.UID);
  if (uidCol < 0) return [];

  const out = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const uid = plantosSafeStr_(row[uidCol]).trim();
    if (!uid) continue;

    const plant = plantosRowToPlant_(hmap, row);
    const hay = plantosNorm_([plant.uid, plant.nickname, plant.primary, plant.classification, plant.location].join(' '));
    if (hay.includes(query)) out.push(plant);
    if (out.length >= limit) break;
  }
  return out;
}

function plantosGetRecentLog(limit) {
  limit = Number(limit || 25);
  const props = PropertiesService.getScriptProperties().getProperties();
  const all = [];
  Object.keys(props).forEach(k => {
    if (!k.startsWith('PLANT_TIMELINE::')) return;
    try {
      const items = JSON.parse(props[k] || '[]');
      items.forEach(it => all.push(it));
    } catch (e) {}
  });
  all.sort((a, b) => String(b.ts || '').localeCompare(String(a.ts || '')));
  return all.slice(0, limit);
}

function plantosGetTimeline(uid, limit) {
  limit = Number(limit || 30);
  const key = 'PLANT_TIMELINE::' + plantosSafeStr_(uid).trim();
  const raw = PropertiesService.getScriptProperties().getProperty(key) || '[]';
  let items = [];
  try { items = JSON.parse(raw); } catch (e) { items = []; }
  return items.slice(0, limit);
}

/* ===================== PHOTO BACKEND (kept single, canonical, sheet-cache optional) ===================== */

function plantosUploadPlantPhoto(uid, dataUrl, originalName) {
  uid = String(uid || '').trim();
  if (!uid) return { ok: false, reason: 'Missing uid' };

  const parsed = plantosParseDataUrl_(dataUrl);
  if (!parsed || !parsed.bytes) return { ok: false, reason: 'Bad image data' };

  const plantsRoot = plantosGetPlantsRoot_();
  const plantFolder = plantosResolveOrCreatePlantFolder_(plantsRoot, uid);
  try {
    const canonical = plantosCanonicalFolderName_(uid);
    if (plantFolder.getName() !== canonical) plantFolder.setName(canonical);
  } catch (e) {}

  const photosFolder = plantosEnsureSubfolder_(plantFolder, PLANTOS_BACKEND_CFG.PHOTOS_SUBFOLDER);

  const safeName = (originalName && String(originalName).trim()) ? String(originalName).trim() : 'photo.jpg';
  const ext = safeName.toLowerCase().endsWith('.png') || parsed.mime === 'image/png' ? 'png' : 'jpg';

  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm-ss');
  const filename = `${ts}_UID${uid}.${ext}`;

  const blob = Utilities.newBlob(parsed.bytes, parsed.mime || (ext === 'png' ? 'image/png' : 'image/jpeg'), filename);
  const file = photosFolder.createFile(blob);

  // PUBLIC: anyone with link can view (for thumbnails in web app)
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {}

  const fileId = file.getId();
  const viewUrl = file.getUrl();
  const thumbUrl = plantosDriveThumbUrl_(fileId, 300);

  // Optional: store latest photo fields if columns exist
  plantosWriteLatestPhotoToSheet_(uid, {
    fileId,
    viewUrl,
    thumbUrl,
    name: filename,
    updated: new Date().toISOString()
  });

  return {
    ok: true,
    photo: { fileId, viewUrl, thumbUrl, name: filename, updated: new Date().toISOString() }
  };
}

function plantosGetLatestPhoto(uid) {
  uid = String(uid || '').trim();
  if (!uid) return { ok: false, reason: 'Missing uid' };

  const fromSheet = plantosReadLatestPhotoFromSheet_(uid);
  if (fromSheet) return { ok: true, photo: fromSheet };

  const plantsRoot = plantosGetPlantsRoot_();
  const plantFolder = plantosResolveOrCreatePlantFolder_(plantsRoot, uid);
  const photosFolder = plantosEnsureSubfolder_(plantFolder, PLANTOS_BACKEND_CFG.PHOTOS_SUBFOLDER);

  const files = photosFolder.getFiles();
  let newest = null;

  while (files.hasNext()) {
    const f = files.next();
    const mt = f.getMimeType ? f.getMimeType() : '';
    if (mt && !mt.startsWith('image/')) continue;
    const t = f.getLastUpdated ? f.getLastUpdated() : new Date(0);
    if (!newest || t > newest.t) newest = { f, t };
  }

  if (!newest) return { ok: true, photo: null };

  const fileId = newest.f.getId();
  return {
    ok: true,
    photo: {
      fileId,
      viewUrl: newest.f.getUrl(),
      thumbUrl: plantosDriveThumbUrl_(fileId, 300),
      name: newest.f.getName(),
      updated: newest.t.toISOString()
    }
  };
}

function plantosParseDataUrl_(dataUrl) {
  const s = String(dataUrl || '');
  const m = s.match(/^data:([^;]+);base64,(.+)$/);
  if (!m) return null;
  const mime = m[1];
  const b64 = m[2];
  const bytes = Utilities.base64Decode(b64);
  return { mime, bytes };
}

function plantosDriveThumbUrl_(fileId, sizePx) {
  const sz = sizePx || 300;
  return `https://drive.google.com/thumbnail?id=${encodeURIComponent(fileId)}&sz=w${encodeURIComponent(sz)}`;
}

function plantosWriteLatestPhotoToSheet_(uid, photo) {
  try {
    const { sh, hmap } = plantosReadInventory_();
    const H = PLANTOS_BACKEND_CFG.HEADERS;

    const uidCol = plantosCol_(hmap, H.UID);
    if (uidCol < 0) return;

    const idCol = plantosCol_(hmap, H.LATEST_PHOTO_ID);
    const thCol = plantosCol_(hmap, H.LATEST_PHOTO_THUMB);
    const vwCol = plantosCol_(hmap, H.LATEST_PHOTO_VIEW);
    const upCol = plantosCol_(hmap, H.LATEST_PHOTO_UPDATED);

    // Only write if at least ID column exists
    if (idCol < 0) return;

    const lastRow = sh.getLastRow();
    if (lastRow < 2) return;

    const uids = sh.getRange(2, uidCol + 1, lastRow - 1, 1).getValues().map(r => String(r[0] || '').trim());
    const idx = uids.findIndex(x => x === uid);
    if (idx < 0) return;

    const rowNum = 2 + idx;

    sh.getRange(rowNum, idCol + 1).setValue(photo.fileId || '');
    if (thCol >= 0) sh.getRange(rowNum, thCol + 1).setValue(photo.thumbUrl || '');
    if (vwCol >= 0) sh.getRange(rowNum, vwCol + 1).setValue(photo.viewUrl || '');
    if (upCol >= 0) sh.getRange(rowNum, upCol + 1).setValue(photo.updated || '');
  } catch (e) {
    // ignore
  }
}

function plantosReadLatestPhotoFromSheet_(uid) {
  try {
    const { sh, hmap } = plantosReadInventory_();
    const H = PLANTOS_BACKEND_CFG.HEADERS;

    const uidCol = plantosCol_(hmap, H.UID);
    const idCol = plantosCol_(hmap, H.LATEST_PHOTO_ID);
    const thCol = plantosCol_(hmap, H.LATEST_PHOTO_THUMB);
    const vwCol = plantosCol_(hmap, H.LATEST_PHOTO_VIEW);
    const upCol = plantosCol_(hmap, H.LATEST_PHOTO_UPDATED);

    if (uidCol < 0 || idCol < 0) return null;

    const lastRow = sh.getLastRow();
    if (lastRow < 2) return null;

    const data = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (String(row[uidCol] || '').trim() !== uid) continue;

      const fileId = String(row[idCol] || '').trim();
      if (!fileId) return null;

      const thumbUrl = (thCol >= 0) ? String(row[thCol] || '').trim() : plantosDriveThumbUrl_(fileId, 300);
      let viewUrl = (vwCol >= 0) ? String(row[vwCol] || '').trim() : '';
      if (!viewUrl) {
        try { viewUrl = DriveApp.getFileById(fileId).getUrl(); } catch (e) {}
      }

      return {
        fileId,
        thumbUrl,
        viewUrl,
        updated: (upCol >= 0) ? String(row[upCol] || '').trim() : ''
      };
    }
    return null;
  } catch (e) {
    return null;
  }
}

/* ===================== ROW -> PLANT OBJECT ===================== */

function plantosGetByHeader_(hmap, row, headerName) {
  const c = plantosCol_(hmap, headerName);
  if (c < 0) return '';
  return row[c];
}

function plantosGetByHeaderDate_(hmap, row, headerName) {
  const c = plantosCol_(hmap, headerName);
  if (c < 0) return '';
  const d = plantosAsDate_(row[c]);
  return d ? d : '';
}

function plantosRowToPlant_(hmap, row) {
  const H = PLANTOS_BACKEND_CFG.HEADERS;

  const uid = plantosGetByHeader_(hmap, row, H.UID);

  const nickname = plantosGetByHeader_(hmap, row, H.NICKNAME);
  const genus = plantosGetByHeader_(hmap, row, H.GENUS);
  const taxon = plantosGetByHeader_(hmap, row, H.TAXON);

  const location = plantosGetByHeader_(hmap, row, H.LOCATION);

  const folderId = plantosGetByHeader_(hmap, row, H.FOLDER_ID);
  const folderUrl = plantosGetByHeader_(hmap, row, H.FOLDER_URL);
  const careDocUrl = plantosGetByHeader_(hmap, row, H.CARE_DOC_URL);

  const plantPageUrl = plantosGetByHeader_(hmap, row, H.PLANT_PAGE_URL);

  const lastWatered = plantosGetByHeaderDate_(hmap, row, H.LAST_WATERED);
  const lastFertilized = plantosGetByHeaderDate_(hmap, row, H.LAST_FERTILIZED);
  const everyDays = plantosGetByHeader_(hmap, row, H.WATER_EVERY_DAYS);

  const potSize = plantosGetByHeader_(hmap, row, H.POT_SIZE);
  const medium = plantosGetByHeader_(hmap, row, H.MEDIUM);
  const birthday = plantosGetByHeaderDate_(hmap, row, H.BIRTHDAY);

  const primary = String(nickname || '').trim() || [genus, taxon].filter(Boolean).join(' ') || uid;

  let due = '';
  const lw = lastWatered ? plantosAsDate_(lastWatered) : null;
  const ev = Number(everyDays);
  if (lw && !isNaN(ev) && ev > 0) due = plantosFmtDate_(plantosAddDays_(lw, ev));

  return {
    uid,
    nickname: nickname || '',
    primary,

    classification: [genus, taxon].filter(Boolean).join(' ').trim(),
    gs: '',

    location: location || '',

    folderId: folderId || '',
    folderUrl: folderUrl || '',
    careDocUrl: careDocUrl || '',

    plantPageUrl: plantPageUrl || '',

    lastWatered: lastWatered ? plantosFmtDate_(plantosAsDate_(lastWatered)) : '',
    due,
    everyDays: everyDays || '',

    lastFertilized: lastFertilized ? plantosFmtDate_(plantosAsDate_(lastFertilized)) : '',

    potSize: potSize || '',
    substrate: medium || '',
    medium: medium || '',

    birthday: birthday ? plantosFmtDate_(plantosAsDate_(birthday)) : '',

    humanPlantId: plantosGetByHeader_(hmap, row, H.PLANT_ID) || '',
  };
}

/* ===================== TIMELINE STORAGE ===================== */

function plantosTimelineAppend_(uid, payload, when) {
  const key = 'PLANT_TIMELINE::' + uid;
  const raw = PropertiesService.getScriptProperties().getProperty(key) || '[]';
  let items = [];
  try { items = JSON.parse(raw); } catch (e) { items = []; }

  const ts = Utilities.formatDate(when, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  const action =
    payload.repot ? 'REPOT' :
    payload.water ? 'WATERED' :
    payload.fertilize ? 'FERTILIZED' :
    'UPDATE';

  let details = '';
  if (payload.repot) details = `Pot: ${payload.potSize || ''} • Substrate: ${payload.substrate || ''}`;
  if (payload.notes) details = (details ? (details + ' • ') : '') + payload.notes;

  items.unshift({ uid, ts, action, details });
  items = items.slice(0, 120);

  PropertiesService.getScriptProperties().setProperty(key, JSON.stringify(items));
}

/* ===================== OPTIONAL doGet (COMMENTED OUT) ===================== */
/*
  If you do NOT already have doGet elsewhere, you can copy/paste this into a separate WebApp.gs file.
  Leaving it commented here prevents accidental override of your existing doGet.

function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};
  const t = HtmlService.createTemplateFromFile('Index');
  t.baseUrl = plantosGetSetting_(PLANTOS_BACKEND_CFG.SETTINGS_KEYS.ACTIVE_WEBAPP_URL) || '';
  t.mode = params.mode || 'home';
  t.uid = params.uid || '';
  t.loc = params.loc || '';
  t.openAdd = params.openAdd || '';
  return t.evaluate().setTitle('PlantOS').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
*/

/* ===================== Web App routing (Active) ===================== */

/**
 * Handle HTTP GET requests for the PlantOS web app. This function is the
 * entry point for users accessing the web interface. It reads query
 * parameters and passes them as variables to the HTML template.
 *
 * If no ACTIVE_WEBAPP_URL has been configured yet, the script will
 * attempt to detect its own URL as a fallback. The mode is resolved from
 * the URL parameters or defaults to "home". When a uid parameter is
 * present without an explicit mode, the default mode becomes "plant".
 *
 * @param {Object} e Event parameter containing request information.
 * @return {HtmlOutput} Rendered HTML template for the app.
 */
function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};
  // Determine base URL from settings or current script URL
  let baseUrl = plantosGetSetting_(PLANTOS_BACKEND_CFG.SETTINGS_KEYS.ACTIVE_WEBAPP_URL);
  if (!baseUrl) baseUrl = plantosGetCurrentWebAppUrl_() || '';

  // Extract parameters with sane defaults
  const uid = String(params.uid || '').trim();
  let mode = String(params.mode || '').trim();
  const loc = String(params.loc || '').trim();
  const openAdd = String(params.openAdd || '').trim();

  // If a UID is present and no mode specified, default to plant page
  if (!mode && uid) mode = 'plant';
  if (!mode) mode = 'home';

  const t = HtmlService.createTemplateFromFile('App');
  t.baseUrl = baseUrl;
  t.mode = mode;
  t.uid = uid;
  t.loc = loc;
  t.openAdd = openAdd;
  return t.evaluate().setTitle('PlantOS').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ===================== Batch Fertilize ===================== */

/**
 * Perform a batch fertilize operation for multiple plants. This function
 * mirrors the behaviour of plantosBatchWater but writes to the
 * fertilization columns. Each UID is deduplicated and looked up once.
 *
 * @param {string[]} uids Array of Plant UIDs to mark as fertilized.
 * @return {Object} Result object with count of updated plants.
 */
function plantosBatchFertilize(uids) {
  uids = uids || [];
  if (!Array.isArray(uids) || !uids.length) return { ok: true, count: 0 };

  const { sh, values, hmap } = plantosReadInventory_();
  const H = PLANTOS_BACKEND_CFG.HEADERS;
  const uidCol = plantosCol_(hmap, H.UID);
  const fertilizedCol = plantosCol_(hmap, H.FERTILIZED);
  const lastFertilizedCol = plantosCol_(hmap, H.LAST_FERTILIZED);
  if (uidCol < 0) throw new Error('Missing Plant UID column');

  // Build a set of unique UIDs for quick lookup
  const set = {};
  uids.forEach(u => {
    const k = plantosSafeStr_(u).trim();
    if (k) set[k] = true;
  });

  const now = plantosNow_();
  let count = 0;

  for (let r = 1; r < values.length; r++) {
    const uid = plantosSafeStr_(values[r][uidCol]).trim();
    if (!uid || !set[uid]) continue;

    if (fertilizedCol >= 0) sh.getRange(r + 1, fertilizedCol + 1).setValue(true);
    if (lastFertilizedCol >= 0) sh.getRange(r + 1, lastFertilizedCol + 1).setValue(now);

    // Append to timeline
    plantosTimelineAppend_(uid, { fertilize: true }, now);
    count++;
  }

  return { ok: true, count };
}

/* ===================== WEB APP ROUTING (ACTIVE) ===================== */
/**
 * Web app entrypoint.
 * Supports:
 * - .../exec?mode=plant&uid=16740419
 * - .../exec?uid=16740419            (defaults to plant)
 * - .../exec?mode=uid16740419        (compact)
 * - .../exec?mode=locations          (legacy alias -> my-plants)
 */
function doGet(e) {
  const params = (e && e.parameter) ? e.parameter : {};

  let baseUrl = plantosGetSetting_(PLANTOS_BACKEND_CFG.SETTINGS_KEYS.ACTIVE_WEBAPP_URL);
  if (!baseUrl) {
    try { baseUrl = ScriptApp.getService().getUrl() || ''; } catch (err) { baseUrl = ''; }
  }

  let mode = String(params.mode || '').trim();
  let uid = String(params.uid || '').trim();
  const loc = String(params.loc || '').trim();
  const openAdd = String(params.openAdd || '').trim();

  // Compact: ?mode=uid16740419
  if (!uid) {
    const m = mode.match(/^uid(\d+)$/i);
    if (m && m[1]) {
      uid = m[1];
      mode = 'plant';
    }
  }

  // If UID present but no mode, default plant
  if (!mode && uid) mode = 'plant';
  if (!mode) mode = 'home';

  // Legacy aliases
  const ml = mode.toLowerCase();
  if (ml === 'locations' || ml === 'plants') mode = 'my-plants';

  const t = HtmlService.createTemplateFromFile('App');
  t.baseUrl = baseUrl;
  t.mode = mode;
  t.uid = uid;
  t.loc = loc;
  t.openAdd = openAdd;
  return t.evaluate().setTitle('PlantOS').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ===================== BATCH FERTILIZE ===================== */
/**
 * Mirrors plantosBatchWater:
 * - sets Fertilized=true (if present)
 * - sets Last Fertilized=now (column name: "Last Fertilized")
 * - appends to timeline
 */
function plantosBatchFertilize(uids, actionLabel) {
  uids = uids || [];
  if (!Array.isArray(uids) || !uids.length) return { ok: true, count: 0 };

  const label = plantosSafeStr_(actionLabel).trim();

  const { sh, values, hmap } = plantosReadInventory_();
  const uidCol = plantosCol_(hmap, PLANTOS_BACKEND_CFG.HEADERS.UID);
  const fertilizedCol = plantosCol_(hmap, PLANTOS_BACKEND_CFG.HEADERS.FERTILIZED);
  const lastFertilizedCol = plantosCol_(hmap, PLANTOS_BACKEND_CFG.HEADERS.LAST_FERTILIZED);
  if (uidCol < 0) throw new Error('Missing Plant UID');

  const set = {};
  uids.forEach(u => { const k = plantosSafeStr_(u).trim(); if (k) set[k] = true; });

  const now = plantosNow_();
  let count = 0;

  for (let r = 1; r < values.length; r++) {
    const uid = plantosSafeStr_(values[r][uidCol]).trim();
    if (!uid || !set[uid]) continue;

    if (fertilizedCol >= 0) sh.getRange(r + 1, fertilizedCol + 1).setValue(true);
    if (lastFertilizedCol >= 0) sh.getRange(r + 1, lastFertilizedCol + 1).setValue(now);

    plantosTimelineAppend_(uid, label ? { fertilize: true, notes: label } : { fertilize: true }, now);
    count++;
  }

  return { ok: true, count };
}
