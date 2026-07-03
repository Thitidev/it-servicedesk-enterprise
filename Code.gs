// ============================================================
// IT ServiceDesk — Google Apps Script Backend
// Deploy: Extensions → Apps Script → Deploy → Web App
//   Execute as: Me  |  Who has access: Anyone
// ============================================================

const SHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const SS       = SpreadsheetApp.getActiveSpreadsheet();

// ── Sheet names ──────────────────────────────────────────────
const SH = {
  TICKETS  : 'Tickets',
  ASSETS   : 'Assets',
  DOCUMENTS: 'Documents',
  USERS    : 'Users',
  AUDIT    : 'AuditLog',
  SETTINGS : 'Settings',
};

// ── CORS headers ─────────────────────────────────────────────
function corsHeaders() {
  return ContentService.createTextOutput()
    .setMimeType(ContentService.MimeType.JSON);
}
function json(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// ROUTER — GET
// ============================================================
function doGet(e) {
  try {
    const action = e.parameter.action || 'ping';
    switch (action) {
      case 'ping':        return json({ ok: true, ts: new Date().toISOString() });
      case 'getTickets':  return json(getTickets(e.parameter));
      case 'getAssets':   return json(getAssets(e.parameter));
      case 'getDocuments':return json(getDocuments(e.parameter));
      case 'getUsers':    return json(getUsers());
      case 'getAudit':    return json(getAudit());
      case 'getStats':    return json(getStats());
      case 'getSettings': return json(getSettings());
      default: return json({ error: 'Unknown action: ' + action });
    }
  } catch(err) {
    return json({ error: err.message, stack: err.stack });
  }
}

// ============================================================
// ROUTER — POST
// ============================================================
function doPost(e) {
  try {
    const body   = JSON.parse(e.postData.contents);
    const action = body.action;
    switch (action) {
      case 'createTicket':  return json(createTicket(body.data));
      case 'updateTicket':  return json(updateTicket(body.data));
      case 'deleteTicket':  return json(deleteRow(SH.TICKETS, body.id));
      case 'createAsset':   return json(createAsset(body.data));
      case 'updateAsset':   return json(updateAsset(body.data));
      case 'deleteAsset':   return json(deleteRow(SH.ASSETS, body.id));
      case 'createDocument':return json(createDocument(body.data));
      case 'deleteDocument':return json(deleteRow(SH.DOCUMENTS, body.id));
      case 'createUser':    return json(createUser(body.data));
      case 'updateUser':    return json(updateUser(body.data));
      case 'saveSettings':  return json(saveSettings(body.data));
      case 'addAudit':      return json(addAudit(body.data));
      default: return json({ error: 'Unknown action: ' + action });
    }
  } catch(err) {
    return json({ error: err.message });
  }
}

// ============================================================
// INIT — สร้าง Sheet structure ครั้งแรก
// ============================================================
function initSheets() {
  const headers = {
    [SH.TICKETS]: [
      'id','subject','requester','email','dept','branch','category','subcategory',
      'priority','impact','urgency','status','technicianId','technicianName','teamId',
      'assetCode','description','resolution','rootCause','preventive','tags',
      'slaPercent','slaDeadline','createdAt','updatedAt','closedAt','csatScore','csatNote'
    ],
    [SH.ASSETS]: [
      'code','name','type','brand','model','serialNo','cpu','ram','storage','os',
      'userId','userName','dept','location','floor','status','purchasePrice',
      'purchaseDate','vendor','warrantyExpiry','macAddress','ipAddress',
      'software','notes','createdAt','updatedAt'
    ],
    [SH.DOCUMENTS]: [
      'id','assetCode','assetName','docType','name','vendor','docDate','amount',
      'fileSize','fileExt','fileUrl','notes','uploadedAt','uploadedBy','isPublic'
    ],
    [SH.USERS]: [
      'id','firstName','lastName','email','employeeId','dept','branch','phone',
      'role','status','lastLogin','ticketCount','createdAt'
    ],
    [SH.AUDIT]: [
      'id','timestamp','userId','userName','action','targetType','targetId',
      'detail','ipAddress','userAgent'
    ],
    [SH.SETTINGS]: ['key','value','updatedAt'],
  };

  Object.entries(headers).forEach(([name, cols]) => {
    let sh = SS.getSheetByName(name);
    if (!sh) {
      sh = SS.insertSheet(name);
    }
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, cols.length).setValues([cols]);
      sh.getRange(1, 1, 1, cols.length)
        .setBackground('#1860C4')
        .setFontColor('#FFFFFF')
        .setFontWeight('bold');
      sh.setFrozenRows(1);
    }
  });

  // Seed sample data
  seedSampleData();
  return { ok: true, message: 'Sheets initialized!' };
}

// ============================================================
// SEED DATA
// ============================================================
function seedSampleData() {
  const ticketSh = SS.getSheetByName(SH.TICKETS);
  if (ticketSh.getLastRow() > 1) return; // already seeded

  const now = new Date().toISOString();
  const tickets = [
    ['TK-0089','ไม่สามารถเชื่อมต่อ WiFi หลัง Windows Update','อรุณี พัทนา','arune@co.th','การตลาด','สำนักงานใหญ่','Network','WiFi','high','ผู้ใช้คนเดียว','สูง','open','','','Network Team','','ไม่สามารถ connect WiFi ได้','','','','wifi,network,update',88,'2026-06-27T09:00:00Z',now,now,'','',''],
    ['TK-0088','Microsoft Teams เข้า Meeting ไม่ได้','ประยุทธ์ ชูศรี','prayuth@co.th','การเงิน','สำนักงานใหญ่','Microsoft Teams','Meeting','medium','ผู้ใช้คนเดียว','ปานกลาง','progress','USR-001','กฤต แสงทอง','Software Team','','เข้า Teams ไม่ได้','','','','teams,meeting',55,'2026-06-27T09:00:00Z',now,now,'','',''],
    ['TK-0087','Outlook ค้นหาอีเมลไม่ได้','สมศักดิ์ วงขาม','somsak@co.th','HR','สำนักงานใหม','อีเมล','Outlook','low','ผู้ใช้คนเดียว','ต่ำ','open','','','','','ค้นหาไม่ได้','','','','outlook,search',72,'2026-06-28T09:00:00Z',now,now,'','',''],
    ['TK-0086','เครื่องพิมพ์ออฟไลน์ ชั้น 8','มาลี ศรีสุข','malee@co.th','HR','สำนักงานใหญ่','เครื่องพิมพ์','Network Printer','medium','ทั้งแผนก','สูง','waiting','USR-003','ธนากร พัน','Hardware Team','PRT-HP-101','เครื่องพิมพ์ offline','','','','printer,offline',20,'2026-06-26T09:00:00Z',now,now,'','',''],
    ['TK-0082','หน้าจอมอนิเตอร์กระพริบ','รัตนา โพล','rattana@co.th','การขาย','สำนักงานใหญ่','Hardware','Monitor','critical','ผู้ใช้คนเดียว','สูง','progress','USR-001','กฤต แสงทอง','Hardware Team','DT-HP-0412','จอกระพริบ','','','','monitor,hardware,critical',8,'2026-06-25T09:00:00Z',now,now,'','',''],
  ];
  ticketSh.getRange(2,1,tickets.length,tickets[0].length).setValues(tickets);

  const assetSh = SS.getSheetByName(SH.ASSETS);
  if (assetSh.getLastRow() <= 1) {
    const assets = [
      ['NB-DELL-0412','LAPTOP-SJ-001','notebook','Dell','Latitude 5520','DL552-XK914','i7-1165G7','16GB DDR4','512GB NVMe','Windows 11 Pro','EMP-00412','สมชาย ใจดี','การเงิน','สำนักงานใหญ่','ชั้น 5','active',42500,'2024-01-15','Dell Thailand','2025-12-31','A4:C3:F0:12:45:7E','10.1.4.52','Microsoft 365,Chrome,VPN Client','',now,now],
      ['NB-HP-0381','LAPTOP-NW-002','notebook','HP','EliteBook 840','HP840-NW02','i5-1135G7','8GB DDR4','512GB SSD','Windows 11 Pro','USR-002','ณัฐยา วนิช','ไอที','สำนักงานใหญ่','ชั้น 1','active',38900,'2024-03-01','HP Thailand','2027-06-01','B2:D4:F0:34:56:8E','10.1.1.18','Microsoft 365,Teams','',now,now],
      ['SRV-HP-001','SERVER-MAIN-01','server','HP','ProLiant DL380','SRV380-001','Xeon E5-2690','128GB ECC','4TB RAID','Windows Server 2022','IT-TEAM','IT Team','ไอที','Server Room','ชั้น B1','active',285000,'2023-06-15','HP Thailand','2028-01-01','','10.1.0.1','IIS,SQL Server,AD','',now,now],
      ['PRT-HP-101','PRINTER-FIN-01','printer','HP','LaserJet Pro M404dn','PRT404-FIN','—','—','—','—','FIN-POOL','Finance Pool','การเงิน','สำนักงานใหญ่','ชั้น 5','active',12500,'2023-08-25','Office Depot','2026-08-20','','10.1.4.100','','',now,now],
    ];
    assetSh.getRange(2,1,assets.length,assets[0].length).setValues(assets);
  }

  const userSh = SS.getSheetByName(SH.USERS);
  if (userSh.getLastRow() <= 1) {
    const users = [
      ['USR-001','กฤต','แสงทอง','krit.s@company.co.th','IT-001','ไอที','สำนักงานใหญ่','081-111-1111','tech','active',now,48,now],
      ['USR-002','ณัฐยา','วนิช','nattaya.w@company.co.th','IT-002','ไอที','สำนักงานใหญ่','081-222-2222','tech','active',now,42,now],
      ['USR-003','ธนากร','พัน','thanakorn.p@company.co.th','IT-003','ไอที','สำนักงานใหญ่','081-333-3333','tech','active',now,31,now],
      ['ADM-001','Admin','User','admin@company.co.th','IT-000','ไอที','สำนักงานใหญ่','081-000-0000','admin','active',now,0,now],
      ['EMP-00412','สมชาย','ใจดี','somchai.j@company.co.th','EMP-00412','การเงิน','สำนักงานใหญ่','081-234-5678','req','active',now,12,now],
    ];
    userSh.getRange(2,1,users.length,users[0].length).setValues(users);
  }

  // Default settings
  const settingSh = SS.getSheetByName(SH.SETTINGS);
  if (settingSh.getLastRow() <= 1) {
    const settings = [
      ['company_name','Thai Enterprise Co., Ltd.',now],
      ['sla_critical_response','15',now],
      ['sla_critical_resolve','240',now],
      ['sla_high_response','60',now],
      ['sla_high_resolve','480',now],
      ['sla_medium_response','240',now],
      ['sla_medium_resolve','1440',now],
      ['sla_low_response','480',now],
      ['sla_low_resolve','4320',now],
      ['notify_teams','true',now],
      ['notify_email','true',now],
      ['session_timeout','30',now],
    ];
    settingSh.getRange(2,1,settings.length,3).setValues(settings);
  }
}

// ============================================================
// HELPERS
// ============================================================
function sheetToObjects(sheetName) {
  const sh   = SS.getSheetByName(sheetName);
  if (!sh) return [];
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}

function findRowById(sheetName, id, idCol = 0) {
  const sh   = SS.getSheetByName(sheetName);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) === String(id)) return i + 1; // 1-indexed
  }
  return -1;
}

function nextId(sheetName, prefix) {
  const sh      = SS.getSheetByName(sheetName);
  const lastRow = sh.getLastRow();
  const num     = lastRow; // row count = auto-increment
  const pad     = String(num).padStart(4, '0');
  return `${prefix}-${pad}`;
}

function deleteRow(sheetName, id) {
  const row = findRowById(sheetName, id);
  if (row < 0) return { ok: false, error: 'Not found' };
  SS.getSheetByName(sheetName).deleteRow(row);
  addAudit({ action: 'DELETE', targetType: sheetName, targetId: id, detail: 'Deleted' });
  return { ok: true };
}

// ============================================================
// TICKETS
// ============================================================
function getTickets(params) {
  const tickets = sheetToObjects(SH.TICKETS);
  let result = tickets;

  if (params.status && params.status !== 'all') {
    result = result.filter(t => t.status === params.status);
  }
  if (params.priority) {
    result = result.filter(t => t.priority === params.priority);
  }
  if (params.technicianId) {
    result = result.filter(t => t.technicianId === params.technicianId);
  }
  if (params.search) {
    const q = params.search.toLowerCase();
    result = result.filter(t =>
      String(t.id).toLowerCase().includes(q) ||
      String(t.subject).toLowerCase().includes(q) ||
      String(t.requester).toLowerCase().includes(q)
    );
  }

  return { ok: true, data: result, total: result.length };
}

function createTicket(data) {
  const sh  = SS.getSheetByName(SH.TICKETS);
  const now = new Date().toISOString();
  const id  = nextId(SH.TICKETS, 'TK');

  // Calculate SLA deadline
  const slaMap = { critical: 240, high: 480, medium: 1440, low: 4320 };
  const mins   = slaMap[data.priority] || 1440;
  const deadline = new Date(Date.now() + mins * 60000).toISOString();

  const row = [
    id, data.subject, data.requester, data.email, data.dept, data.branch,
    data.category, data.subcategory, data.priority, data.impact, data.urgency,
    'new', data.technicianId||'', data.technicianName||'', data.teamId||'',
    data.assetCode||'', data.description||'', '','','', data.tags||'',
    100, deadline, now, now, '', '', ''
  ];
  sh.appendRow(row);
  addAudit({ action: 'CREATE_TICKET', targetType: 'Ticket', targetId: id, detail: data.subject, userName: data.requester });
  return { ok: true, id, data: { ...data, id, status: 'new', createdAt: now } };
}

function updateTicket(data) {
  const sh      = SS.getSheetByName(SH.TICKETS);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const row     = findRowById(SH.TICKETS, data.id);
  if (row < 0) return { ok: false, error: 'Ticket not found' };

  const now = new Date().toISOString();
  headers.forEach((h, i) => {
    if (h in data && h !== 'id' && h !== 'createdAt') {
      sh.getRange(row, i + 1).setValue(data[h]);
    }
  });
  sh.getRange(row, headers.indexOf('updatedAt') + 1).setValue(now);
  addAudit({ action: 'UPDATE_TICKET', targetType: 'Ticket', targetId: data.id, detail: `Status: ${data.status||''}` });
  return { ok: true, id: data.id };
}

// ============================================================
// ASSETS
// ============================================================
function getAssets(params) {
  let assets = sheetToObjects(SH.ASSETS);
  if (params.type && params.type !== 'all') {
    assets = assets.filter(a => a.type === params.type);
  }
  if (params.status) {
    assets = assets.filter(a => a.status === params.status);
  }
  if (params.search) {
    const q = params.search.toLowerCase();
    assets = assets.filter(a =>
      String(a.code).toLowerCase().includes(q) ||
      String(a.brand).toLowerCase().includes(q) ||
      String(a.userName).toLowerCase().includes(q)
    );
  }
  return { ok: true, data: assets, total: assets.length };
}

function createAsset(data) {
  const sh  = SS.getSheetByName(SH.ASSETS);
  const now = new Date().toISOString();
  const row = [
    data.code, data.name, data.type, data.brand, data.model, data.serialNo,
    data.cpu||'', data.ram||'', data.storage||'', data.os||'',
    data.userId||'', data.userName||'', data.dept||'', data.location||'',
    data.floor||'', data.status||'active', data.purchasePrice||0,
    data.purchaseDate||'', data.vendor||'', data.warrantyExpiry||'',
    data.macAddress||'', data.ipAddress||'', data.software||'', data.notes||'',
    now, now
  ];
  sh.appendRow(row);
  addAudit({ action: 'CREATE_ASSET', targetType: 'Asset', targetId: data.code, detail: data.brand + ' ' + data.model });
  return { ok: true, code: data.code };
}

function updateAsset(data) {
  const sh      = SS.getSheetByName(SH.ASSETS);
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const row     = findRowById(SH.ASSETS, data.code, 0);
  if (row < 0) return { ok: false, error: 'Asset not found' };
  const now = new Date().toISOString();
  headers.forEach((h, i) => {
    if (h in data && h !== 'code' && h !== 'createdAt') {
      sh.getRange(row, i + 1).setValue(data[h]);
    }
  });
  sh.getRange(row, headers.indexOf('updatedAt') + 1).setValue(now);
  return { ok: true };
}

// ============================================================
// DOCUMENTS
// ============================================================
function getDocuments(params) {
  let docs = sheetToObjects(SH.DOCUMENTS);
  if (params.assetCode) {
    docs = docs.filter(d => d.assetCode === params.assetCode);
  }
  if (params.docType && params.docType !== 'all') {
    docs = docs.filter(d => d.docType === params.docType);
  }
  return { ok: true, data: docs, total: docs.length };
}

function createDocument(data) {
  const sh  = SS.getSheetByName(SH.DOCUMENTS);
  const now = new Date().toISOString();
  const id  = 'DOC-' + String(sh.getLastRow()).padStart(4, '0');
  const row = [
    id, data.assetCode, data.assetName, data.docType, data.name,
    data.vendor||'', data.docDate||'', data.amount||'',
    data.fileSize||'', data.fileExt||'pdf', data.fileUrl||'',
    data.notes||'', now, data.uploadedBy||'Admin', data.isPublic||false
  ];
  sh.appendRow(row);
  addAudit({ action: 'UPLOAD_DOC', targetType: 'Document', targetId: id, detail: data.name });
  return { ok: true, id };
}

// ============================================================
// USERS
// ============================================================
function getUsers() {
  return { ok: true, data: sheetToObjects(SH.USERS) };
}

function createUser(data) {
  const sh  = SS.getSheetByName(SH.USERS);
  const now = new Date().toISOString();
  const id  = 'USR-' + String(sh.getLastRow()).padStart(4, '0');
  const row = [
    id, data.firstName, data.lastName, data.email, data.employeeId||'',
    data.dept||'', data.branch||'', data.phone||'',
    data.role||'req', 'active', '', 0, now
  ];
  sh.appendRow(row);
  return { ok: true, id };
}

function updateUser(data) {
  const sh      = SS.getSheetByName(SH.USERS);
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const row     = findRowById(SH.USERS, data.id);
  if (row < 0) return { ok: false, error: 'User not found' };
  headers.forEach((h, i) => {
    if (h in data && h !== 'id' && h !== 'createdAt') {
      sh.getRange(row, i + 1).setValue(data[h]);
    }
  });
  return { ok: true };
}

// ============================================================
// AUDIT LOG
// ============================================================
function getAudit() {
  const rows = sheetToObjects(SH.AUDIT).reverse().slice(0, 100);
  return { ok: true, data: rows };
}

function addAudit(data) {
  const sh  = SS.getSheetByName(SH.AUDIT);
  const now = new Date().toISOString();
  const id  = 'AUD-' + String(sh.getLastRow()).padStart(6, '0');
  sh.appendRow([
    id, now,
    data.userId   || 'system',
    data.userName || 'System',
    data.action   || '',
    data.targetType || '',
    data.targetId   || '',
    data.detail     || '',
    data.ipAddress  || '',
    data.userAgent  || '',
  ]);
  return { ok: true, id };
}

// ============================================================
// STATS (Dashboard)
// ============================================================
function getStats() {
  const tickets = sheetToObjects(SH.TICKETS);
  const assets  = sheetToObjects(SH.ASSETS);
  const docs    = sheetToObjects(SH.DOCUMENTS);

  const now  = new Date();
  const mon0 = new Date(now.getFullYear(), now.getMonth(), 1);

  const byStatus = {};
  const byPri    = {};
  const byCat    = {};
  let   resolved = 0;
  let   slaBreached = 0;
  let   slaSum   = 0;

  tickets.forEach(t => {
    byStatus[t.status] = (byStatus[t.status] || 0) + 1;
    byPri[t.priority]  = (byPri[t.priority]  || 0) + 1;
    byCat[t.category]  = (byCat[t.category]  || 0) + 1;
    const pct = Number(t.slaPercent) || 100;
    slaSum += pct;
    if (pct < 1) slaBreached++;
    if (t.status === 'resolved' || t.status === 'closed') {
      if (new Date(t.updatedAt) >= mon0) resolved++;
    }
  });

  const open     = (byStatus['open']||0) + (byStatus['progress']||0) + (byStatus['assigned']||0);
  const waiting  = (byStatus['waiting']||0);
  const total    = tickets.length;
  const slaComp  = total > 0 ? Math.round(slaSum / total) : 100;

  // Asset warranty expiring within 90 days
  const soon90 = assets.filter(a => {
    if (!a.warrantyExpiry) return false;
    const parts = String(a.warrantyExpiry).split('-');
    const exp   = new Date(a.warrantyExpiry);
    const diff  = (exp - now) / 86400000;
    return diff > 0 && diff < 90;
  }).length;

  return {
    ok: true,
    data: {
      total, open, waiting, resolved,
      slaBreached, slaCompliance: slaComp,
      critical: byPri['critical'] || 0,
      byStatus, byPriority: byPri, byCategory: byCat,
      assets: { total: assets.length, warrantyExpiringSoon: soon90 },
      documents: { total: docs.length },
    }
  };
}

// ============================================================
// SETTINGS
// ============================================================
function getSettings() {
  const rows = sheetToObjects(SH.SETTINGS);
  const obj  = {};
  rows.forEach(r => { obj[r.key] = r.value; });
  return { ok: true, data: obj };
}

function saveSettings(data) {
  const sh      = SS.getSheetByName(SH.SETTINGS);
  const now     = new Date().toISOString();
  const rows    = sh.getDataRange().getValues();
  const keyCol  = 0;
  const valCol  = 1;

  Object.entries(data).forEach(([key, val]) => {
    let found = false;
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][keyCol] === key) {
        sh.getRange(i + 1, valCol + 1).setValue(val);
        sh.getRange(i + 1, 3).setValue(now);
        found = true;
        break;
      }
    }
    if (!found) {
      sh.appendRow([key, val, now]);
    }
  });
  return { ok: true };
}

// ============================================================
// TRIGGER: Auto SLA update every 15 min
// ============================================================
function setupTrigger() {
  // Delete existing
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  // Create new time-based trigger
  ScriptApp.newTrigger('updateSLAPercents')
    .timeBased()
    .everyMinutes(15)
    .create();
  return 'Trigger set: every 15 minutes';
}

function updateSLAPercents() {
  const sh      = SS.getSheetByName(SH.TICKETS);
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const data    = sh.getDataRange().getValues();
  const now     = new Date();
  const slaIdx  = headers.indexOf('slaPercent');
  const dlIdx   = headers.indexOf('slaDeadline');
  const crIdx   = headers.indexOf('createdAt');
  const stIdx   = headers.indexOf('status');
  const prIdx   = headers.indexOf('priority');
  const slaMap  = { critical:240, high:480, medium:1440, low:4320 };

  for (let i = 1; i < data.length; i++) {
    const st = data[i][stIdx];
    if (st === 'resolved' || st === 'closed') continue;

    const pri      = data[i][prIdx];
    const created  = new Date(data[i][crIdx]);
    const deadline = new Date(data[i][dlIdx]);
    const totalMs  = (slaMap[pri] || 1440) * 60000;
    const usedMs   = now - created;
    const pct      = Math.max(0, Math.round((1 - usedMs / totalMs) * 100));

    sh.getRange(i + 1, slaIdx + 1).setValue(pct);
  }
}
