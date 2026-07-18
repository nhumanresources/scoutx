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

/* ═══ AUTH (central accounts + email verification) ═══
 * Verification emails are sent FROM the Google account that deploys this
 * script (krishna@nhrms.com). Users on Zoho/GoDaddy mail receive them fine.
 */
var ADMIN_EMAIL = 'krishna@nhrms.com';
var ALLOWED_DOMAINS = ['nhrms.in', 'nhrms.com', 'rytadvisory.com', 'rytadvisory.in', 'rytpro.com'];

function loadUsers() {
  var p = PropertiesService.getScriptProperties().getProperty('sx_users');
  return p ? JSON.parse(p) : [];
}
function persistUsers(u) {
  PropertiesService.getScriptProperties().setProperty('sx_users', JSON.stringify(u));
}
function domainOk(email) {
  var d = String(email).toLowerCase().split('@')[1] || '';
  return ALLOWED_DOMAINS.indexOf(d) !== -1;
}

function doPost(e) {
  try {
    var req = JSON.parse(e.postData.contents);
    switch (req.action) {
      case 'sync':    return handleSync(req);
      case 'signup':  return handleSignup(req);
      case 'verify':  return handleVerify(req);
      case 'login':   return handleLogin(req);
      case 'users':   return handleUsers(req);
      case 'setuser': return handleSetUser(req);
      default:        return out({ error: 'unknown action' });
    }
  } catch (err) {
    return out({ error: String(err) });
  }
}

function handleSignup(req) {
  var email = String(req.email || '').toLowerCase().trim();
  if (!req.name || !email || !req.hash || !req.role) return out({ error: 'Missing fields' });
  if (!domainOk(email)) return out({ error: 'Only company emails allowed (nhrms / rytadvisory / rytpro)' });
  var users = loadUsers();
  var existing = users.filter(function (u) { return u.email === email; })[0];
  if (existing && existing.verified) return out({ error: 'Account already exists — sign in instead' });
  var code = String(Math.floor(100000 + Math.random() * 900000));
  var user = {
    name: req.name, email: email, hash: req.hash,
    role: email === ADMIN_EMAIL ? 'admin' : req.role,
    verified: false, code: code, codeExp: Date.now() + 30 * 60 * 1000, created: Date.now()
  };
  if (existing) { users[users.indexOf(existing)] = user; } else { users.push(user); }
  persistUsers(users);
  MailApp.sendEmail({
    to: email,
    subject: 'ScoutX — your verification code: ' + code,
    htmlBody:
      '<div style="font-family:Arial,sans-serif;max-width:440px;margin:auto;padding:24px;border:1px solid #eee;border-radius:12px">' +
      '<h2 style="margin:0 0 4px">🎯 ScoutX</h2>' +
      '<p style="color:#666;margin:0 0 20px">NHRMS · AI Recruitment & Marketing Platform</p>' +
      '<p>Hi ' + user.name.split(' ')[0] + ', your verification code is:</p>' +
      '<p style="font-size:34px;font-weight:bold;letter-spacing:8px;text-align:center;background:#f4f6ff;border-radius:10px;padding:16px">' + code + '</p>' +
      '<p style="color:#666;font-size:13px">Enter this at <a href="https://scoutx.nhrms.com">scoutx.nhrms.com</a>. Valid for 30 minutes. If you did not request this, ignore this email.</p>' +
      '</div>'
  });
  return out({ ok: true });
}

function handleVerify(req) {
  var users = loadUsers();
  var u = users.filter(function (x) { return x.email === String(req.email || '').toLowerCase(); })[0];
  if (!u) return out({ error: 'No pending signup for this email' });
  if (u.verified) return out({ ok: true, role: u.role });
  if (Date.now() > u.codeExp) return out({ error: 'Code expired — sign up again to get a new one' });
  if (String(req.code) !== u.code) return out({ error: 'Wrong code — check the email' });
  u.verified = true; delete u.code; delete u.codeExp;
  persistUsers(users);
  return out({ ok: true, role: u.role });
}

function handleLogin(req) {
  var users = loadUsers();
  var u = users.filter(function (x) { return x.email === String(req.email || '').toLowerCase(); })[0];
  if (!u || u.hash !== req.hash) return out({ error: 'Wrong email or password' });
  if (!u.verified) return out({ error: 'Email not verified yet — check your inbox for the code' });
  return out({ ok: true, name: u.name, role: u.role });
}

function requireAdmin(req) {
  var users = loadUsers();
  var a = users.filter(function (x) { return x.email === String(req.email || '').toLowerCase(); })[0];
  return (a && a.verified && a.role === 'admin' && a.hash === req.hash) ? users : null;
}

function handleUsers(req) {
  var users = requireAdmin(req);
  if (!users) return out({ error: 'Admin authentication failed' });
  return out({ ok: true, users: users.map(function (u) {
    return { name: u.name, email: u.email, role: u.role, verified: !!u.verified, created: u.created };
  }) });
}

function handleSetUser(req) {
  var users = requireAdmin(req);
  if (!users) return out({ error: 'Admin authentication failed' });
  var t = users.filter(function (x) { return x.email === String(req.target || '').toLowerCase(); })[0];
  if (!t) return out({ error: 'User not found' });
  if (req.del) { users.splice(users.indexOf(t), 1); }
  else if (req.role) { t.role = req.role; }
  persistUsers(users);
  return out({ ok: true });
}

function handleSync(req) {
  var ss = SpreadsheetApp.openById(req.sheetId);
  if (req.module === 'recruit') syncRecruit(ss, req.payload);
  else if (req.module === 'marketing') syncMarketing(ss, req.payload);
  return out({ ok: true, module: req.module, at: new Date().toISOString() });
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
