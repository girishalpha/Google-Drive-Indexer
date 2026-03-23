/**
 * ╔══════════════════════════════════════════════════════════════╗
 * ║       ULTIMATE DRIVE INDEXER — Google Apps Script            ║
 * ║  High-performance sync with smart upserts & scheduled runs   ║
 * ╚══════════════════════════════════════════════════════════════╝
 */

// ── Sheet Names ─────────────────────────────────────────────────────
const SHEET_INDEX  = 'Zunofy Knowledge Base Drive Indexer';
const SHEET_LOG    = 'Sync Log';
const SHEET_CONFIG = 'Config';

// ── Column Headers (Order dictates output structure) ────────────────
const HEADERS = [
  '#', '✅ New', 'File Name', 'File Type', 'Folder Path',
  'Created At', 'Last Modified At', 'Owner / Creator',
  'File Size', 'Share Status', 'File Link', 'Direct Download',
  'Tags / Description', 'File ID', 'Last Synced'
];

// Column map for easy referencing
const C = {};
HEADERS.forEach((h, i) => C[h] = i);

// ═══════════════════════════════════════════════════════════════════
// UI MENU (Matches your screenshot)
// ═══════════════════════════════════════════════════════════════════
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📁 Drive Index')
    .addItem('▶ Full Sync (All Files)',              'runFullSync')
    .addItem('⚡ Quick Sync (New & Modified Only)',   'runQuickSync')
    .addSeparator()
    .addSubMenu(ui.createMenu('⏰ Schedule Auto-Refresh')
      .addItem('Every 15 Minutes', 'triggerEvery15Min')
      .addItem('Every Hour',       'triggerEveryHour')
      .addItem('Every 6 Hours',    'triggerEvery6Hours')
      .addItem('Daily at 8 AM',    'triggerDaily')
      .addSeparator()
      .addItem('🗑 Remove All Triggers', 'removeAllTriggers'))
    .addSeparator()
    .addItem('⚙️ Setup Config Sheet', 'initConfigSheet')
    .addItem('🧹 Clear Index',         'clearIndex')
    .addItem('📋 View Sync Log',       'goToSyncLog')
    .addToUi();
}

// ═══════════════════════════════════════════════════════════════════
// SYNC LOGIC
// ═══════════════════════════════════════════════════════════════════

function runFullSync() {
  executeSync_(true);
}

function runQuickSync() {
  executeSync_(false);
}

function executeSync_(isFullSync) {
  const cfg = readConfig_();
  if (!cfg.driveId) {
    SpreadsheetApp.getUi().alert('Configuration Missing', 'Please run "Setup Config Sheet" and enter a Drive/Folder ID.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet_(ss, SHEET_INDEX);
  setupHeaders_(sheet);

  const t0 = Date.now();
  const existingData = isFullSync ? {} : buildIdRowMap_(sheet);
  
  if (isFullSync && sheet.getLastRow() > 1) {
    sheet.deleteRows(2, sheet.getLastRow() - 1); // Wipe slate clean
  }

  // Fetch Drive Data
  const files = fetchDriveFiles_(cfg);
  
  // Clear previous 'New' flags
  if (!isFullSync && sheet.getLastRow() > 1) {
    sheet.getRange(2, C['✅ New'] + 1, sheet.getLastRow() - 1, 1).clearContent();
  }

  const newRows = [];
  const updateRows = []; // Format: {rowNum: num, data: []}
  const now = formatDate_(new Date().toISOString(), cfg.tz);
  
  let added = 0, updated = 0;

  files.forEach((f) => {
    const rowData = formatRowData_(f, cfg, now);
    
    if (!isFullSync && existingData[f.id]) {
      rowData[C['✅ New']] = ''; // Not new
      updateRows.push({ rowNum: existingData[f.id], data: rowData });
      updated++;
    } else {
      rowData[C['✅ New']] = '✅'; // Flag as new
      newRows.push(rowData);
      added++;
    }
  });

  // Apply Updates
  if (updateRows.length > 0) {
    // Write updates individually (can be optimized further with Sheets API for massive datasets)
    updateRows.forEach(u => {
      sheet.getRange(u.rowNum, 1, 1, HEADERS.length).setValues([u.data]);
    });
  }

  // Append New Rows
  if (newRows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, HEADERS.length).setValues(newRows);
  }

  // Post-Processing: Sort, Number, and Style
  if (sheet.getLastRow() > 1) {
    // Sort by Created At (Descending - Newest First)
    sheet.getRange(2, 2, sheet.getLastRow() - 1, HEADERS.length - 1).sort({column: C['Created At'] + 1, ascending: false});
    
    // Recalculate Row Numbers (Column A)
    const numRows = sheet.getLastRow() - 1;
    const numbering = Array.from({length: numRows}, (_, i) => [i + 1]);
    sheet.getRange(2, 1, numRows, 1).setValues(numbering);
    
    styleRows_(sheet);
  }

  const sec = ((Date.now() - t0) / 1000).toFixed(1);
  const syncType = isFullSync ? 'Full Sync' : `Quick Sync (+${added}, ~${updated})`;
  appendLog_(ss, syncType, added + updated, sec);
  toast_(`✅ ${syncType} complete in ${sec}s`);
}

// ═══════════════════════════════════════════════════════════════════
// DRIVE API FETCHING (Optimized Batch Crawl)
// ═══════════════════════════════════════════════════════════════════

function fetchDriveFiles_(cfg) {
  const results = [];
  const queryParts = ["trashed = false"];
  
  if (!cfg.isSharedDrive) {
    queryParts.push(`'${cfg.driveId}' in parents`);
  }

  // Queue structure prevents stack overflows on deeply nested shared drives
  const folderQueue = [{ id: cfg.driveId, path: cfg.isSharedDrive ? 'Shared Drive' : 'My Drive' }];
  
  while (folderQueue.length > 0) {
    const currentFolder = folderQueue.shift();
    let pageToken = null;
    
    do {
      const params = {
        q: `'${currentFolder.id}' in parents and trashed = false`,
        // OPTIMIZATION: Specifically requesting lastModifyingUser emails and owner emails
        fields: 'nextPageToken, files(id, name, mimeType, webViewLink, webContentLink, createdTime, modifiedTime, owners(displayName, emailAddress), lastModifyingUser(displayName, emailAddress), size, description, shared)',
        pageSize: 1000,
        supportsAllDrives: true,
        includeItemsFromAllDrives: true
      };

      if (cfg.isSharedDrive) {
        params.driveId = cfg.driveId;
        params.corpora = 'drive';
      }
      
      if (pageToken) params.pageToken = pageToken;
      
      let response;
      try {
        response = Drive.Files.list(params);
      } catch (e) {
        console.error("Drive fetch error: " + e.message);
        break; 
      }

      const items = response.files || [];
      
      items.forEach(f => {
        f._folderPath = currentFolder.path;
        results.push(f);
        
        // Add subfolders to queue if enabled
        if (cfg.includeSubfolders && f.mimeType === 'application/vnd.google-apps.folder') {
          folderQueue.push({ id: f.id, path: `${currentFolder.path} > ${f.name}` });
        }
      });
      
      pageToken = response.nextPageToken;
    } while (pageToken);
    
    // Safety guard: Prevents Apps Script 6-minute timeout limit on massive drives
    if (results.length > 8000) break; 
  }
  
  return results;
}

function formatRowData_(f, cfg, syncTime) {
  const isFolder = f.mimeType === 'application/vnd.google-apps.folder';
  const mimeLabel = f.mimeType.split('.').pop().replace('folder', 'Folder').replace('document', 'Doc').replace('spreadsheet', 'Sheet').replace('presentation', 'Slides');
  
  // ── FIX: SHARED DRIVE EMAIL EXTRACTION ──
  let userEmail = '';
  
  // 1. Try to get standard owner (Works for "My Drive")
  if (f.owners && f.owners.length > 0 && f.owners[0].emailAddress) {
    userEmail = f.owners[0].emailAddress;
  } 
  // 2. Fallback to Last Modifying User (Crucial for "Shared Drives")
  else if (f.lastModifyingUser && f.lastModifyingUser.emailAddress) {
    userEmail = f.lastModifyingUser.emailAddress;
  } 
  // 3. Absolute fallback to Name if email is hidden by Workspace Admin settings
  else if (f.lastModifyingUser && f.lastModifyingUser.displayName) {
    userEmail = f.lastModifyingUser.displayName;
  } else {
    userEmail = 'System / Unknown';
  }

  const size = parseInt(f.size) > 0 ? (parseInt(f.size) / 1048576).toFixed(2) + ' MB' : (isFolder ? '' : 'N/A');
  const dLink = f.webContentLink || (isFolder ? '' : `https://drive.google.com/uc?export=download&id=${f.id}`);

  const row = new Array(HEADERS.length).fill('');
  row[C['✅ New']]             = ''; 
  row[C['File Name']]          = f.name;
  row[C['File Type']]          = mimeLabel;
  row[C['Folder Path']]        = f._folderPath;
  row[C['Created At']]         = f.createdTime ? formatDate_(f.createdTime, cfg.tz) : '';
  row[C['Last Modified At']]   = f.modifiedTime ? formatDate_(f.modifiedTime, cfg.tz) : '';
  row[C['Owner / Creator']]    = userEmail; // <--- Now outputs the actual email!
  row[C['File Size']]          = size;
  row[C['Share Status']]       = f.shared ? 'Shared' : 'Private';
  row[C['File Link']]          = f.webViewLink || '';
  row[C['Direct Download']]    = dLink;
  row[C['Tags / Description']] = f.description || '';
  row[C['File ID']]            = f.id;
  row[C['Last Synced']]        = syncTime;
  
  return row;
}


// ═══════════════════════════════════════════════════════════════════
// SHEET & UTILITY FUNCTIONS
// ═══════════════════════════════════════════════════════════════════

function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function setupHeaders_(sheet) {
  if (sheet.getLastRow() > 0 && sheet.getRange(1,1).getValue() === HEADERS[0]) return;
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
    .setBackground('#34a853').setFontColor('#ffffff') // Google Drive Green
    .setFontWeight('bold').setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(3); // Freeze up to File Name
}

function buildIdRowMap_(sheet) {
  const map = {};
  const last = sheet.getLastRow();
  if (last < 2) return map;
  const ids = sheet.getRange(2, C['File ID'] + 1, last - 1, 1).getValues();
  ids.forEach((row, i) => { if (row[0]) map[row[0]] = i + 2; });
  return map;
}

function styleRows_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // Zebra Striping
  for (let r = 2; r <= lastRow; r++) {
    sheet.getRange(r, 1, 1, HEADERS.length)
         .setBackground(r % 2 === 0 ? '#f8f9fa' : '#ffffff');
  }

  // Hyperlinks for File Link
  const linkCol = C['File Link'] + 1;
  const nameCol = C['File Name'] + 1;
  const names = sheet.getRange(2, nameCol, lastRow - 1, 1).getValues();
  const links = sheet.getRange(2, linkCol, lastRow - 1, 1).getValues();
  
  const richVals = links.map((lr, i) => {
    const url = lr[0];
    if (!url) return SpreadsheetApp.newRichTextValue().setText('N/A').build();
    return SpreadsheetApp.newRichTextValue()
      .setText(names[i][0] || 'Open File')
      .setLinkUrl(url)
      .setTextStyle(SpreadsheetApp.newTextStyle().setForegroundColor('#1a73e8').setUnderline(true).build())
      .build();
  });
  sheet.getRange(2, linkCol, lastRow - 1, 1).setRichTextValues(richVals.map(v => [v]));
  
  // Auto-resize
  [1, 2, 3, 4, 5, 6, 7].forEach(c => sheet.autoResizeColumn(c));
}

// ═══════════════════════════════════════════════════════════════════
// CONFIG & LOGGING
// ═══════════════════════════════════════════════════════════════════

function initConfigSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet_(ss, SHEET_CONFIG);
  sheet.clearContents();
  const rows = [
    ['Setting', 'Value', 'Description'],
    ['TARGET_DRIVE_ID', '', 'Paste Folder ID or Shared Drive ID here'],
    ['IS_SHARED_DRIVE', 'FALSE', 'Set TRUE if the ID above is a Shared Drive'],
    ['INCLUDE_SUBFOLDERS', 'TRUE', 'TRUE = index sub-folders recursively'],
    ['TIMEZONE', Session.getScriptTimeZone(), 'e.g., Asia/Kolkata']
  ];
  sheet.getRange(1, 1, rows.length, 3).setValues(rows);
  sheet.getRange(1, 1, 1, 3).setBackground('#34a853').setFontColor('#ffffff').setFontWeight('bold');
  sheet.autoResizeColumns(1, 3);
  ss.setActiveSheet(sheet);
  toast_('⚙️ Config sheet generated. Please fill in the Drive ID.');
}

function readConfig_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s = ss.getSheetByName(SHEET_CONFIG);
  const cfg = { driveId: '', isSharedDrive: false, includeSubfolders: true, tz: Session.getScriptTimeZone() };
  if (!s) return cfg;
  
  s.getDataRange().getValues().slice(1).forEach(([k, v]) => {
    if (k === 'TARGET_DRIVE_ID') cfg.driveId = String(v).trim();
    if (k === 'IS_SHARED_DRIVE') cfg.isSharedDrive = (String(v).toUpperCase() === 'TRUE');
    if (k === 'INCLUDE_SUBFOLDERS') cfg.includeSubfolders = (String(v).toUpperCase() === 'TRUE');
    if (k === 'TIMEZONE' && v) cfg.tz = String(v).trim();
  });
  return cfg;
}

function appendLog_(ss, type, count, sec) {
  const s = getOrCreateSheet_(ss, SHEET_LOG);
  if (s.getLastRow() === 0) {
    s.appendRow(['Timestamp', 'Sync Type', 'Items Processed', 'Duration (s)', 'Status']);
    s.getRange(1,1,1,5).setBackground('#34a853').setFontColor('#ffffff').setFontWeight('bold');
  }
  s.appendRow([new Date().toLocaleString(), type, count, sec, '✅ Success']);
}

function goToSyncLog() {
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOG);
  if (s) SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(s);
}

function clearIndex() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('Clear Index?', 'Are you sure you want to wipe all indexed data? (The files in Drive remain safe).', ui.ButtonSet.YES_NO) === ui.Button.YES) {
    const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_INDEX);
    if (s && s.getLastRow() > 1) s.deleteRows(2, s.getLastRow() - 1);
    toast_('🧹 Index cleared');
  }
}

// ═══════════════════════════════════════════════════════════════════
// TRIGGERS
// ═══════════════════════════════════════════════════════════════════
function triggerEvery15Min()  { setTimeTrigger_(15, 'MINUTE'); }
function triggerEveryHour()   { setTimeTrigger_(1,  'HOUR');   }
function triggerEvery6Hours() { setTimeTrigger_(6,  'HOUR');   }
function triggerDaily() {
  removeAllTriggers();
  ScriptApp.newTrigger('runQuickSync').timeBased().everyDays(1).atHour(8).create();
  toast_('⏰ Daily sync set for 8:00 AM');
}

function setTimeTrigger_(n, unit) {
  removeAllTriggers();
  const b = ScriptApp.newTrigger('runQuickSync').timeBased();
  if (unit === 'MINUTE') b.everyMinutes(n); else b.everyHours(n);
  b.create();
  toast_(`⏰ Auto-sync set: Every ${n} ${unit.toLowerCase()}(s)`);
}

function removeAllTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  toast_('🗑 Triggers removed');
}

// Utility formatting
function formatDate_(iso, tz) {
  return Utilities.formatDate(new Date(iso), tz, 'yyyy-MM-dd HH:mm:ss');
}
function toast_(msg) {
  SpreadsheetApp.getActiveSpreadsheet().toast(msg, 'Drive Indexer', 5);
}
