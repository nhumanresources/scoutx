/**
 * ScoutX — Dual-Sheet Sync (Recruit + Marketing)
 * Deploy: Extensions → Apps Script → paste this → Deploy → New deployment
 *   Type: Web app · Execute as: Me · Who has access: Anyone
 * Paste the resulting /exec URL into ScoutX Settings.
 *
 * The app POSTs: { action:'sync', module:'recruit'|'marketing', sheetId:'...', payload:{...} }
 * Each sync fully rewrites the tabs so sheets always mirror the app.
 */

function doGet() {
  return ContentService.createTextOutput(JSON.stringify({ status: 'ScoutX API running' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var req = JSON.parse(e.postData.contents);
    if (req.action !== 'sync') return out({ error: 'unknown action' });
    var ss = SpreadsheetApp.openById(req.sheetId);
    if (req.module === 'recruit') syncRecruit(ss, req.payload);
    else if (req.module === 'marketing') syncMarketing(ss, req.payload);
    return out({ ok: true, module: req.module, at: new Date().toISOString() });
  } catch (err) {
    return out({ error: String(err) });
  }
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function writeTab(ss, name, headers, rows) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.clearContents();
  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  if (rows.length) sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
  sh.setFrozenRows(1);
}

function syncRecruit(ss, p) {
  writeTab(ss, 'Jobs',
    ['job_id', 'title', 'status', 'created'],
    (p.jobs || []).map(function (j) { return [j.id, j.title, j.status, ts(j.created)]; }));
  writeTab(ss, 'Candidates',
    ['id', 'name', 'phone', 'role', 'job_id', 'stage', 'ai_score', 'wa_sent', 'added'],
    (p.candidates || []).map(function (c) {
      return [c.id, c.name, c.phone, c.role, c.jobId || '', c.stage, c.score || '', c.waSent ? 'TRUE' : 'FALSE', ts(c.added)];
    }));
  writeTab(ss, 'Conversations',
    ['candidate_id', 'role', 'content', 'timestamp'],
    (p.conversations || []).map(function (m) { return [m.candidate_id, m.role, m.content, ts(m.time)]; }));
}

function syncMarketing(ss, p) {
  writeTab(ss, 'Leads',
    ['lead_id', 'name', 'company', 'role_title', 'phone', 'stage', 'heat', 'source', 'added'],
    (p.leads || []).map(function (l) {
      return [l.id, l.name, l.company || '', l.roleTitle || '', l.phone || '', l.stage || 'Prospect', l.heat || '', l.source || '', ts(l.added)];
    }));
  writeTab(ss, 'Campaigns',
    ['campaign_id', 'name', 'audience', 'sent', 'status', 'created'],
    (p.campaigns || []).map(function (c) { return [c.id, c.name, c.audience, c.sent, c.status, ts(c.created)]; }));
  writeTab(ss, 'Inbox',
    ['lead_id', 'role', 'content', 'timestamp'],
    (p.inbox || []).map(function (m) { return [m.lead_id, m.role, m.content, ts(m.time)]; }));
}

function ts(t) {
  return t ? new Date(t).toISOString().replace('T', ' ').slice(0, 16) : '';
}
