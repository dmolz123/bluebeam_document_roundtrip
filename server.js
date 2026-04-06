/**
 * Bluebeam Studio API — Document Roundtrip PoC
 * Proof-of-concept reference implementation. Not for production use.
 *
 * XML-FALLBACK VERSION:
 *   - Keeps working check-in to Project copy
 *   - Keeps working exportmarkups XML generation
 *   - Replaces markuplist dependency with exported XML download + parse
 *   - Preserves existing flow/endpoints/UI compatibility
 *
 * Roundtrip flow:
 *   0a. /poc/setup-project          — Create folders + upload custom-columns.xml (once)
 *   0b. /poc/upload-to-project      — Upload PDF(s) from UI → project review folder
 *   0c. /poc/apply-custom-columns   — Apply custom-columns.xml to each uploaded file (optional / not used in UI)
 *   1.  /poc/trigger                — Simulate source-system workflow event
 *   2.  /poc/create-session         — Create Studio Session
 *   3.  /poc/register-webhook       — Subscribe to session events
 *   4.  /poc/checkout-to-session    — Check project file(s) out into session
 *   5.  /poc/invite-reviewers       — Invite reviewers
 *   6.  (Review in Bluebeam Revu — no API step)
 *   7.  /poc/checkin                — Check session file(s) back into project
 *   8.  /poc/export-markups         — Run exportmarkups job → XML in project
 *   9.  /poc/run-markuplist-job     — Compatibility route; now extracts structured data from exported XML
 *   9b. /poc/downstream-process     — Combined downstream step: checkin + export + XML parse
 *   10. /poc/finalize               — Finalize session
 *   11. /poc/snapshot               — Snapshot + download marked-up PDF
 *   12. /poc/cleanup                — Delete webhook + session
 */

require('dotenv').config();
const express = require('express');
const cors = require('cors');
const fs = require('fs');
const path = require('path');
const multer = require('multer');
const { parseStringPromise } = require('xml2js');
const TokenManager = require('./tokenManager');

const app = express();
const PORT = process.env.PORT || 3000;
const upload = multer({ storage: multer.memoryStorage() });

// -----------------------------------------------------------------------------
// API CONFIGURATION
// -----------------------------------------------------------------------------
const API_V1 = 'https://api.bluebeam.com/publicapi/v1';
const API_V2 = 'https://api.bluebeam.com/publicapi/v2';
const CLIENT_ID = process.env.BB_CLIENT_ID;

const WEBHOOK_CALLBACK_URL =
  process.env.WEBHOOK_CALLBACK_URL ||
  `http://localhost:${PORT}/webhook/studio-events`;

// Hardcoded project ID for this PoC
const POC_PROJECT_ID = '712-566-288';

// Project folder names
const FOLDER_RESOURCES = 'resources';
const FOLDER_REVIEW_DOCS = 'review-documents';
const FOLDER_MARKUP_EXPORTS = 'markup-exports';

// Path to the custom columns XML bundled with this repo
const CUSTOM_COLUMNS_XML_PATH = path.join(__dirname, 'resources', 'custom-columns.xml');

// Optional standalone markup config
const MARKUP_SESSION_ID = process.env.MARKUP_SESSION_ID || '';
const MARKUP_FILE_ID = process.env.MARKUP_FILE_ID || '';
const MARKUP_FILE_NAME = process.env.MARKUP_FILE_NAME || 'Sample Drawing.pdf';

// -----------------------------------------------------------------------------
// DEMO STUB
// -----------------------------------------------------------------------------
let demoStub = {
  documentId: process.env.DEMO_DOCUMENT_ID || 'DOC-001',
  description: process.env.DEMO_DESCRIPTION || 'Design review — coordination update',
  reviewers: [{ email: 'dmolz@bluebeam.com', hasStudioAccount: true }],
  sessionEndDate: new Date(Date.now() + 7 * 24 * 60 * 60 * 1000).toISOString()
};

// -----------------------------------------------------------------------------
// IN-MEMORY STATE
// -----------------------------------------------------------------------------
let pocState = {
  sessionId: null,
  subscriptionId: null,
  projectSetupDone: false,
  folderIds: {},
  customColumnsFileId: null,
  projectFiles: [],
  sessionFileIds: [],
  markupExports: [],
  markups: [],
  markupJobId: null,
  status: 'idle',
  log: [],
  createdAt: null,
  webhookEvents: []
};

function logStep(msg, type = 'info') {
  const entry = { time: new Date().toISOString(), msg, type };
  pocState.log.push(entry);
  console.log(`[${type.toUpperCase()}] ${msg}`);
  return entry;
}

app.use(cors());
app.use(express.json());
app.use(express.static('public'));

const fetch = (...args) =>
  import('node-fetch').then(({ default: fetch }) => fetch(...args));

const tokenManager = new TokenManager();

// -----------------------------------------------------------------------------
// HELPERS
// -----------------------------------------------------------------------------
function resetPocState() {
  pocState = {
    sessionId: null,
    subscriptionId: null,
    projectSetupDone: false,
    folderIds: {},
    customColumnsFileId: null,
    projectFiles: [],
    sessionFileIds: [],
    markupExports: [],
    markups: [],
    markupJobId: null,
    status: 'idle',
    log: [],
    createdAt: null,
    webhookEvents: []
  };
}

function isLocalhost(url) {
  return /^https?:\/\/(localhost|127\.0\.0\.1)/.test(url);
}

function authHeaders(accessToken, extra = {}) {
  return {
    Authorization: `Bearer ${accessToken}`,
    client_id: CLIENT_ID,
    'Content-Type': 'application/json',
    Accept: 'application/json',
    ...extra
  };
}

function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

async function pollJob(url, headers, maxAttempts = 20, intervalMs = 3000) {
  const inProgress = new Set([100, 130, 150]);

  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    await sleep(intervalMs);

    const res = await fetch(url, { headers });
    const data = await res.json();

    const status = data.Status ?? data.JobStatus;
    const msg = data.StatusMessage ?? data.JobStatusMessage ?? '';
    logStep(`Job poll ${attempt}/${maxAttempts}: status=${status} ${msg}`.trim(), 'info');

    if (status === 200) return data;
    if (!inProgress.has(status)) {
      throw new Error(`Job failed (status=${status}): ${msg}`);
    }
  }

  throw new Error(`Job did not complete after ${maxAttempts} attempts`);
}

async function listProjectFolders(accessToken) {
  const resp = await fetch(`${API_V1}/projects/${POC_PROJECT_ID}/folders`, {
    headers: authHeaders(accessToken)
  });
  if (!resp.ok) throw new Error(`Failed to list folders: ${resp.status} - ${await resp.text()}`);
  const data = await resp.json();
  return data.ProjectFolders || [];
}

async function createFolder(name, accessToken) {
  const resp = await fetch(`${API_V1}/projects/${POC_PROJECT_ID}/folders`, {
    method: 'POST',
    headers: authHeaders(accessToken),
    body: JSON.stringify({ Name: name })
  });
  if (!resp.ok) throw new Error(`Failed to create folder "${name}": ${resp.status} - ${await resp.text()}`);
  const data = await resp.json();
  return data.Id;
}

async function uploadFileToProject(fileBuffer, fileName, accessToken, folderId = null) {
  logStep(`Uploading "${fileName}" to project ${POC_PROJECT_ID}${folderId ? ` (folderId=${folderId})` : ''}...`, 'info');

  const metaBody = {
    Name: fileName,
    Size: fileBuffer.length,
    CRC: '0'
  };
  if (folderId) metaBody.ParentFolderId = folderId;

  const metaResp = await fetch(`${API_V1}/projects/${POC_PROJECT_ID}/files`, {
    method: 'POST',
    headers: authHeaders(accessToken),
    body: JSON.stringify(metaBody)
  });

  if (!metaResp.ok) {
    throw new Error(`Metadata block failed for "${fileName}": ${metaResp.status} - ${await metaResp.text()}`);
  }

  const meta = await metaResp.json();
  const projectFileId = meta.Id;
  const uploadUrl = meta.UploadUrl;
  const uploadContentType = meta.UploadContentType || 'application/pdf';

  logStep(`Metadata block created: projectFileId=${projectFileId}`, 'success');

  logStep(`Uploading ${fileBuffer.length} bytes to storage...`, 'info');
  const s3Resp = await fetch(uploadUrl, {
    method: 'PUT',
    headers: {
      'Content-Type': uploadContentType,
      'x-amz-server-side-encryption': 'AES256'
    },
    body: fileBuffer
  });

  if (!s3Resp.ok) {
    throw new Error(`S3 upload failed for "${fileName}": ${s3Resp.status}`);
  }

  logStep('S3 upload complete', 'success');

  const confirmResp = await fetch(
    `${API_V1}/projects/${POC_PROJECT_ID}/files/${projectFileId}/confirm-upload`,
    { method: 'POST', headers: authHeaders(accessToken), body: '{}' }
  );

  if (!confirmResp.ok) {
    throw new Error(`Confirm upload failed for "${fileName}": ${confirmResp.status} - ${await confirmResp.text()}`);
  }

  logStep(`"${fileName}" confirmed in project (projectFileId=${projectFileId})`, 'success');
  return { projectFileId, name: fileName, size: fileBuffer.length, folderId };
}

async function listProjectFiles(accessToken) {
  const resp = await fetch(`${API_V1}/projects/${POC_PROJECT_ID}/files`, {
    headers: authHeaders(accessToken)
  });
  if (!resp.ok) {
    throw new Error(`Failed to list project files: ${resp.status} - ${await resp.text()}`);
  }
  const data = await resp.json();
  return data.ProjectFiles || [];
}

async function getProjectFileByPath(accessToken, filePath) {
  const url = `${API_V1}/projects/${POC_PROJECT_ID}/files/by-path?path=${encodeURIComponent(filePath)}`;
  const resp = await fetch(url, {
    headers: authHeaders(accessToken)
  });

  if (!resp.ok) {
    throw new Error(`Failed to get file by path: ${resp.status} - ${await resp.text()}`);
  }

  return resp.json();
}

async function waitForProjectFileSettlement(accessToken, projectFileId, fileName, options = {}) {
  const {
    attempts = 6,
    intervalMs = 5000,
    finalExtraDelayMs = 5000
  } = options;

  logStep(`Waiting for project file settlement: "${fileName}" (projectFileId=${projectFileId})...`, 'info');

  let lastSeen = null;

  for (let attempt = 1; attempt <= attempts; attempt++) {
    const files = await listProjectFiles(accessToken);
    const match = files.find(f => String(f.Id) === String(projectFileId));

    if (match) {
      lastSeen = match;

      const observed = [
        `path=${match.Path || '(n/a)'}`,
        `name=${match.Name || '(n/a)'}`
      ];

      if (typeof match.Version !== 'undefined') observed.push(`version=${match.Version}`);
      if (typeof match.IsCheckedOut !== 'undefined') observed.push(`isCheckedOut=${match.IsCheckedOut}`);
      if (typeof match.CheckedOut !== 'undefined') observed.push(`checkedOut=${match.CheckedOut}`);
      if (typeof match.ParentFolderId !== 'undefined') observed.push(`parentFolderId=${match.ParentFolderId}`);

      logStep(
        `Project file re-query ${attempt}/${attempts}: found "${fileName}" (${observed.join(', ')})`,
        'info'
      );

      const checkedOutFlags = [match.IsCheckedOut, match.CheckedOut].filter(v => typeof v !== 'undefined');
      const explicitlyCheckedOut = checkedOutFlags.some(v => v === true || String(v).toLowerCase() === 'true');

      if (!explicitlyCheckedOut) {
        if (finalExtraDelayMs > 0) {
          logStep(`Project file appears settled. Waiting an extra ${Math.round(finalExtraDelayMs / 1000)}s before downstream jobs...`, 'info');
          await sleep(finalExtraDelayMs);
        }
        return lastSeen;
      }
    } else {
      logStep(`Project file re-query ${attempt}/${attempts}: projectFileId=${projectFileId} not found yet`, 'warn');
    }

    if (attempt < attempts) {
      await sleep(intervalMs);
    }
  }

  logStep(
    `Project file settlement window elapsed for "${fileName}". Proceeding with downstream jobs using best-known state.`,
    'warn'
  );

  if (finalExtraDelayMs > 0) {
    await sleep(finalExtraDelayMs);
  }

  return lastSeen;
}

async function waitForProjectFileReadyByPath(accessToken, filePath, expectedRevisionId = null) {
  for (let attempt = 1; attempt <= 8; attempt++) {
    const file = await getProjectFileByPath(accessToken, filePath);

    logStep(
      `by-path re-query ${attempt}/8: revision=${file.RevisionID}, inSession=${file.InSession}, isLocked=${file.IsLocked}`,
      'info'
    );

    const revisionOk = expectedRevisionId == null || Number(file.RevisionID) >= Number(expectedRevisionId);
    const ready = revisionOk && !file.InSession && !file.IsLocked;

    if (ready) {
      await sleep(3000);
      return file;
    }

    await sleep(4000);
  }

  throw new Error(`Project file did not become ready by path: ${filePath}`);
}

// -----------------------------------------------------------------------------
// XML PARSING HELPERS
// -----------------------------------------------------------------------------
function lowerKeyMap(obj) {
  const out = {};
  for (const [k, v] of Object.entries(obj || {})) out[k.toLowerCase()] = v;
  return out;
}

function firstDefined(obj, keys) {
  const map = lowerKeyMap(obj);
  for (const key of keys) {
    if (typeof map[key.toLowerCase()] !== 'undefined') return map[key.toLowerCase()];
  }
  return undefined;
}

function scalar(val) {
  if (Array.isArray(val)) return scalar(val[0]);
  if (val && typeof val === 'object') {
    if (typeof val._ !== 'undefined') return val._;
    return '';
  }
  return val;
}

function normalizeMarkupRecord(record, sourceFile) {
  const mapped = lowerKeyMap(record);

  const known = {
    Id: scalar(firstDefined(mapped, ['id', 'markupid', 'markup_id'])),
    Author: scalar(firstDefined(mapped, ['author', 'createdby', 'user', 'username'])),
    Type: scalar(firstDefined(mapped, ['type', 'markuptype'])),
    Subject: scalar(firstDefined(mapped, ['subject', 'label', 'title'])),
    Comment: scalar(firstDefined(mapped, ['comment', 'comments', 'note', 'message', 'reply'])),
    Status: scalar(firstDefined(mapped, ['status', 'state'])),
    Layer: scalar(firstDefined(mapped, ['layer'])),
    Page: scalar(firstDefined(mapped, ['page', 'pagenumber', 'pageindex'])),
    DateCreated: scalar(firstDefined(mapped, ['datecreated', 'created', 'createddate'])),
    DateModified: scalar(firstDefined(mapped, ['datemodified', 'modified', 'modifieddate'])),
    Color: scalar(firstDefined(mapped, ['color'])),
    Checked: scalar(firstDefined(mapped, ['checked'])),
    Locked: scalar(firstDefined(mapped, ['locked']))
  };

  const skip = new Set([
    'id', 'markupid', 'markup_id',
    'author', 'createdby', 'user', 'username',
    'type', 'markuptype',
    'subject', 'label', 'title',
    'comment', 'comments', 'note', 'message', 'reply',
    'status', 'state',
    'layer',
    'page', 'pagenumber', 'pageindex',
    'datecreated', 'created', 'createddate',
    'datemodified', 'modified', 'modifieddate',
    'color',
    'checked',
    'locked'
  ]);

  const extended = {};
  for (const [k, v] of Object.entries(mapped)) {
    if (!skip.has(k)) {
      const sv = scalar(v);
      if (typeof sv !== 'undefined' && sv !== null && String(sv) !== '') {
        extended[k] = sv;
      }
    }
  }

  return {
    ...known,
    ExtendedProperties: extended,
    _sourceFile: sourceFile
  };
}

function looksLikeMarkupRecord(obj) {
  if (!obj || typeof obj !== 'object' || Array.isArray(obj)) return false;

  const keys = Object.keys(lowerKeyMap(obj));
  const hits = [
    'author', 'subject', 'comment', 'status', 'page', 'layer', 'type',
    'markupid', 'id', 'markuptype'
  ].filter(k => keys.includes(k)).length;

  return hits >= 2;
}

function extractMarkupCandidates(node, sourceFile, results = []) {
  if (Array.isArray(node)) {
    for (const item of node) extractMarkupCandidates(item, sourceFile, results);
    return results;
  }

  if (!node || typeof node !== 'object') return results;

  if (looksLikeMarkupRecord(node)) {
    results.push(normalizeMarkupRecord(node, sourceFile));
  }

  for (const value of Object.values(node)) {
    extractMarkupCandidates(value, sourceFile, results);
  }

  return results;
}

async function downloadExportedMarkupXml(accessToken, exportFileName) {
  const xmlPath = `/${FOLDER_MARKUP_EXPORTS}/${exportFileName}`;
  const fileMeta = await getProjectFileByPath(accessToken, xmlPath);

  if (!fileMeta.DownloadUrl) {
    throw new Error(`DownloadUrl missing for exported XML: ${xmlPath}`);
  }

  logStep(`Downloading exported XML from path=${xmlPath} (fileId=${fileMeta.Id})...`, 'info');

  const xmlResp = await fetch(fileMeta.DownloadUrl);
  if (!xmlResp.ok) {
    throw new Error(`Failed to download exported XML "${exportFileName}": ${xmlResp.status}`);
  }

  return xmlResp.text();
}

async function parseBluebeamExportXml(xmlText, sourceFile) {
  const parsed = await parseStringPromise(xmlText, {
    explicitArray: false,
    mergeAttrs: true,
    trim: true
  });

  const candidates = extractMarkupCandidates(parsed, sourceFile);

  // De-duplicate similar rows
  const seen = new Set();
  const unique = [];

  for (const item of candidates) {
    const key = [
      item.Id || '',
      item.Author || '',
      item.Subject || '',
      item.Comment || '',
      item.Page || '',
      item.DateCreated || ''
    ].join('|');

    if (!seen.has(key)) {
      seen.add(key);
      unique.push(item);
    }
  }

  return unique;
}

// -----------------------------------------------------------------------------
// DOWNSTREAM HELPERS
// -----------------------------------------------------------------------------
async function performCheckin(accessToken) {
  if (!pocState.sessionId) throw new Error('No active session');
  if (pocState.sessionFileIds.length === 0) throw new Error('No session files — run checkout-to-session first');

  pocState.status = 'checking-in';
  const results = [];

  for (const sf of pocState.sessionFileIds) {
    logStep(`Checking in "${sf.name}" (sessionFileId=${sf.sessionFileId})...`, 'info');

    const resp = await fetch(
      `${API_V1}/sessions/${pocState.sessionId}/files/${sf.sessionFileId}/checkin`,
      {
        method: 'POST',
        headers: authHeaders(accessToken),
        body: JSON.stringify({ Comment: 'Session markup review complete' })
      }
    );

    if (!resp.ok) {
      const err = await resp.text();
      logStep(`Check-in failed for "${sf.name}": ${resp.status} - ${err}`, 'warn');
      results.push({ name: sf.name, success: false, error: err });
      continue;
    }

    logStep(`"${sf.name}" checked in to project`, 'success');

    await waitForProjectFileSettlement(accessToken, sf.projectFileId, sf.name, {
      attempts: 6,
      intervalMs: 5000,
      finalExtraDelayMs: 5000
    });

    results.push({ name: sf.name, success: true });
  }

  logStep('Check-in complete — project files updated with session markups', 'success');
  return results;
}

async function performExportMarkups(accessToken) {
  if (pocState.sessionFileIds.length === 0) {
    throw new Error('No session files — run checkout-to-session first');
  }

  logStep('Exporting markups to XML...', 'info');
  const results = [];

  if (!pocState.folderIds[FOLDER_MARKUP_EXPORTS]) {
    logStep('markup-exports folder ID not set — re-querying folders...', 'info');
    const folders = await listProjectFolders(accessToken);
    const found = folders.find(f => f.Name === FOLDER_MARKUP_EXPORTS);
    if (found) pocState.folderIds[FOLDER_MARKUP_EXPORTS] = found.Id;
    else throw new Error(`Folder "${FOLDER_MARKUP_EXPORTS}" not found — run setup-project first`);
  }

  for (const sf of pocState.sessionFileIds) {
    const exportFileName = `Markups-${sf.projectFileId}.xml`;

    logStep(`Submitting exportmarkups job for "${sf.name}" → ${exportFileName}...`, 'info');

    const jobResp = await fetch(
      `${API_V1}/projects/${POC_PROJECT_ID}/files/${sf.projectFileId}/jobs/exportmarkups`,
      {
        method: 'POST',
        headers: authHeaders(accessToken),
        body: JSON.stringify({
          OutputFileName: exportFileName,
          OutputPath: FOLDER_MARKUP_EXPORTS,
          Priority: 0
        })
      }
    );

    if (!jobResp.ok) {
      const err = await jobResp.text();
      logStep(`exportmarkups submission failed for "${sf.name}": ${jobResp.status} - ${err}`, 'warn');
      results.push({ name: sf.name, success: false, error: err });
      continue;
    }

    const { Id: jobId } = await jobResp.json();
    logStep(`exportmarkups job submitted: jobId=${jobId} — polling (15s interval)...`, 'success');

    const pollUrl = `${API_V1}/jobs/${jobId}`;
    await pollJob(pollUrl, authHeaders(accessToken), 15, 15000);

    logStep(`Markup XML exported: ${exportFileName}`, 'success');
    logStep('Allowing exported markup artifacts to settle for 5s...', 'info');
    await sleep(5000);

    // upsert-style update
    const existingIndex = pocState.markupExports.findIndex(m => m.exportFileName === exportFileName);
    const exportRecord = {
      name: sf.name,
      exportFileName,
      projectPath: FOLDER_MARKUP_EXPORTS
    };

    if (existingIndex >= 0) {
      pocState.markupExports[existingIndex] = exportRecord;
    } else {
      pocState.markupExports.push(exportRecord);
    }

    results.push({ name: sf.name, success: true, exportFileName });
  }

  return results;
}

async function performMarkupExtractionFromXml(accessToken) {
  if (!pocState.sessionFileIds.length) {
    throw new Error('No session files — run checkout-to-session first');
  }

  pocState.status = 'extracting-markups';
  pocState.markups = [];

  for (const sf of pocState.sessionFileIds) {
    const projectFile = pocState.projectFiles.find(
      pf => String(pf.projectFileId) === String(sf.projectFileId)
    );

    if (!projectFile) {
      throw new Error(`Could not find project file metadata for ${sf.name}`);
    }

    const reviewPath = `/${FOLDER_REVIEW_DOCS}/${projectFile.name}`;
    const currentFile = await waitForProjectFileReadyByPath(accessToken, reviewPath);

    logStep(
      `Resolved project readiness by path=${reviewPath}, byPathFileId=${currentFile?.Id}, usingCheckedInProjectFileId=${sf.projectFileId}, revision=${currentFile?.RevisionID}, inSession=${currentFile?.InSession}, isLocked=${currentFile?.IsLocked}`,
      'info'
    );

    const exportFileName = `Markups-${sf.projectFileId}.xml`;

    logStep(`Downloading and parsing exported XML for "${sf.name}" via ${exportFileName}...`, 'info');

    const xmlText = await downloadExportedMarkupXml(accessToken, exportFileName);
    const fileMarkups = await parseBluebeamExportXml(xmlText, sf.name);

    pocState.markups.push(...fileMarkups);
    logStep(`"${sf.name}" — ${fileMarkups.length} markup(s) extracted from exported XML`, 'success');
  }

  pocState.status = 'active';
  logStep(`XML extraction complete — ${pocState.markups.length} total markup(s)`, 'success');
  return pocState.markups;
}

// -----------------------------------------------------------------------------
// HEALTH CHECK
// -----------------------------------------------------------------------------
app.get('/health', (req, res) => {
  res.json({
    status: 'healthy',
    projectId: POC_PROJECT_ID,
    config: {
      hasClientId: Boolean(CLIENT_ID),
      webhookCallbackUrl: WEBHOOK_CALLBACK_URL,
      webhookIsLocalhost: isLocalhost(WEBHOOK_CALLBACK_URL),
      customColumnsXmlExists: fs.existsSync(CUSTOM_COLUMNS_XML_PATH)
    }
  });
});

// =============================================================================
// POC ROUTES
// =============================================================================
app.get('/poc/state', (req, res) => {
  res.json({ ...pocState, stub: demoStub, projectId: POC_PROJECT_ID });
});

app.get('/poc/stub', (req, res) => res.json(demoStub));

app.post('/poc/configure', (req, res) => {
  const { documentId, description, reviewerEmail } = req.body || {};

  if (documentId) demoStub.documentId = documentId;
  if (description) demoStub.description = description;

  if (reviewerEmail && reviewerEmail !== 'dmolz@bluebeam.com') {
    if (!demoStub.reviewers.some(r => r.email === reviewerEmail)) {
      demoStub.reviewers.push({ email: reviewerEmail, hasStudioAccount: false });
      logStep(`Added reviewer: ${reviewerEmail}`, 'info');
    }
  }

  demoStub.sessionEndDate = new Date(Date.now() + 7 * 24 * 60 * 60 * 1000).toISOString();
  res.json({ success: true, stub: demoStub });
});

app.post('/poc/remove-reviewer', (req, res) => {
  const { email } = req.body || {};
  if (email === 'dmolz@bluebeam.com') {
    return res.status(400).json({ error: 'Cannot remove primary reviewer' });
  }
  demoStub.reviewers = demoStub.reviewers.filter(r => r.email !== email);
  res.json({ success: true, stub: demoStub });
});

app.post('/poc/reset', (req, res) => {
  resetPocState();
  demoStub.reviewers = [{ email: 'dmolz@bluebeam.com', hasStudioAccount: true }];
  logStep('PoC state reset', 'info');
  res.json({ success: true });
});

// -----------------------------------------------------------------------------
// STEP 0a — Project Setup
// -----------------------------------------------------------------------------
app.post('/poc/setup-project', async (req, res) => {
  try {
    if (!fs.existsSync(CUSTOM_COLUMNS_XML_PATH)) {
      throw new Error(`custom-columns.xml not found at ${CUSTOM_COLUMNS_XML_PATH} — ensure resources/ folder is present`);
    }

    logStep('Running project setup...', 'info');
    const accessToken = await tokenManager.getValidAccessToken();

    const existing = await listProjectFolders(accessToken);
    const folderMap = {};
    existing.forEach(f => { folderMap[f.Name] = f.Id; });

    const needed = [FOLDER_RESOURCES, FOLDER_REVIEW_DOCS, FOLDER_MARKUP_EXPORTS];

    for (const name of needed) {
      if (folderMap[name]) {
        logStep(`Folder "${name}" already exists (id=${folderMap[name]})`, 'info');
        pocState.folderIds[name] = folderMap[name];
      } else {
        logStep(`Creating folder "${name}"...`, 'info');
        const id = await createFolder(name, accessToken);
        await sleep(1500);
        pocState.folderIds[name] = id;
        logStep(`Folder "${name}" created (id=${id})`, 'success');
      }
    }

    const allFiles = await listProjectFiles(accessToken);
    const existingXml = allFiles.find(f =>
      f.Name === 'custom-columns.xml' &&
      f.ParentFolderId === pocState.folderIds[FOLDER_RESOURCES]
    );

    if (existingXml) {
      pocState.customColumnsFileId = existingXml.Id;
      logStep(`custom-columns.xml already in resources folder (fileId=${existingXml.Id})`, 'info');
    } else {
      logStep('Uploading custom-columns.xml to resources folder...', 'info');
      const xmlBuffer = fs.readFileSync(CUSTOM_COLUMNS_XML_PATH);
      const result = await uploadFileToProject(
        xmlBuffer,
        'custom-columns.xml',
        accessToken,
        pocState.folderIds[FOLDER_RESOURCES]
      );
      pocState.customColumnsFileId = result.projectFileId;
      logStep(`custom-columns.xml uploaded (fileId=${pocState.customColumnsFileId})`, 'success');
    }

    pocState.projectSetupDone = true;
    logStep('Project setup complete', 'success');

    res.json({
      success: true,
      folderIds: pocState.folderIds,
      customColumnsFileId: pocState.customColumnsFileId,
      state: pocState
    });
  } catch (err) {
    pocState.status = 'error';
    logStep(err.message, 'error');
    res.status(500).json({ error: err.message });
  }
});

// -----------------------------------------------------------------------------
// STEP 0b — Upload PDF(s)
// -----------------------------------------------------------------------------
app.post('/poc/upload-to-project', upload.array('files'), async (req, res) => {
  try {
    if (!req.files || req.files.length === 0) {
      throw new Error('No files received');
    }

    pocState.status = 'uploading';
    logStep(`Received ${req.files.length} file(s) for upload`, 'info');

    const accessToken = await tokenManager.getValidAccessToken();
    const reviewFolderId = pocState.folderIds[FOLDER_REVIEW_DOCS] || null;
    const uploaded = [];

    for (const file of req.files) {
      const result = await uploadFileToProject(
        file.buffer,
        file.originalname,
        accessToken,
        reviewFolderId
      );
      uploaded.push(result);
      pocState.projectFiles.push(result);
    }

    if (uploaded.length > 0 && demoStub.documentId === 'DOC-001') {
      demoStub.documentId = uploaded[0].name.replace(/\.[^.]+$/, '');
    }

    logStep(`${uploaded.length} file(s) uploaded to project`, 'success');
    res.json({ success: true, uploaded, state: pocState });
  } catch (err) {
    pocState.status = 'error';
    logStep(err.message, 'error');
    res.status(500).json({ error: err.message });
  }
});

// -----------------------------------------------------------------------------
// STEP 0c — Apply Custom Columns (optional / not used in UI)
// -----------------------------------------------------------------------------
app.post('/poc/apply-custom-columns', async (req, res) => {
  try {
    if (!pocState.customColumnsFileId) {
      throw new Error('custom-columns.xml not uploaded — run setup-project first');
    }
    if (pocState.projectFiles.length === 0) {
      throw new Error('No project files — run upload-to-project first');
    }

    logStep('Applying custom columns to project files...', 'info');
    const accessToken = await tokenManager.getValidAccessToken();
    const results = [];

    for (const pf of pocState.projectFiles) {
      logStep(`Submitting importcustomcolumns job for "${pf.name}"...`, 'info');

      const jobResp = await fetch(
        `${API_V1}/projects/${POC_PROJECT_ID}/files/${pf.projectFileId}/jobs/importcustomcolumns`,
        {
          method: 'POST',
          headers: authHeaders(accessToken),
          body: JSON.stringify({
            CurrentPassword: '',
            CustomColumnsFileID: parseInt(pocState.customColumnsFileId, 10),
            OutputFileName: pf.name,
            OutputPath: FOLDER_REVIEW_DOCS,
            Priority: 0
          })
        }
      );

      if (!jobResp.ok) {
        const err = await jobResp.text();
        logStep(`importcustomcolumns submission failed for "${pf.name}": ${jobResp.status} - ${err}`, 'warn');
        results.push({ name: pf.name, success: false, error: err });
        continue;
      }

      const { Id: jobId } = await jobResp.json();
      logStep(`Job submitted: jobId=${jobId} — polling...`, 'success');

      const pollUrl = `${API_V1}/jobs/${jobId}`;
      await pollJob(pollUrl, authHeaders(accessToken));

      logStep(`Custom columns applied to "${pf.name}"`, 'success');
      results.push({ name: pf.name, success: true, jobId });
    }

    res.json({ success: true, results, state: pocState });
  } catch (err) {
    pocState.status = 'error';
    logStep(err.message, 'error');
    res.status(500).json({ error: err.message });
  }
});

// -----------------------------------------------------------------------------
// STEP 1 — Trigger
// -----------------------------------------------------------------------------
app.post('/poc/trigger', (req, res) => {
  pocState.status = 'triggered';
  pocState.log = [];

  logStep(`Workflow event received — document: ${demoStub.documentId}`, 'info');
  logStep(`Files: ${pocState.projectFiles.map(f => f.name).join(', ') || '(none)'}`, 'info');
  logStep(`Description: ${demoStub.description}`, 'info');
  logStep(`Reviewers: ${demoStub.reviewers.map(r => r.email).join(', ')}`, 'info');
  logStep(`Session end date: ${new Date(demoStub.sessionEndDate).toLocaleDateString()}`, 'info');

  res.json({ success: true, state: pocState });
});

// -----------------------------------------------------------------------------
// STEP 2 — Create Session
// -----------------------------------------------------------------------------
app.post('/poc/create-session', async (req, res) => {
  try {
    pocState.status = 'creating';
    logStep('Creating Bluebeam Studio Session...', 'info');

    const accessToken = await tokenManager.getValidAccessToken();
    const sessionName = `${demoStub.documentId}_Review_${new Date().toISOString().slice(0, 10)}`;

    const resp = await fetch(`${API_V1}/sessions`, {
      method: 'POST',
      headers: authHeaders(accessToken),
      body: JSON.stringify({
        Name: sessionName,
        Notification: true,
        Restricted: true,
        SessionEndDate: demoStub.sessionEndDate,
        DefaultPermissions: [
          { Type: 'Markup', Allow: 'Allow' },
          { Type: 'SaveCopy', Allow: 'Allow' },
          { Type: 'PrintCopy', Allow: 'Allow' },
          { Type: 'MarkupAlert', Allow: 'Allow' },
          { Type: 'AddDocuments', Allow: 'Deny' }
        ]
      })
    });

    if (!resp.ok) {
      throw new Error(`Session creation failed: ${resp.status} - ${await resp.text()}`);
    }

    const data = await resp.json();
    pocState.sessionId = data.Id;
    pocState.createdAt = new Date().toISOString();

    logStep(`Session created: ID=${pocState.sessionId}`, 'success');
    logStep(`Session name: ${sessionName}`, 'info');
    res.json({ success: true, sessionId: pocState.sessionId, state: pocState });
  } catch (err) {
    pocState.status = 'error';
    logStep(err.message, 'error');
    res.status(500).json({ error: err.message });
  }
});

// -----------------------------------------------------------------------------
// STEP 3 — Register Webhook
// -----------------------------------------------------------------------------
app.post('/poc/register-webhook', async (req, res) => {
  try {
    if (!pocState.sessionId) {
      throw new Error('No active session — run create-session first');
    }

    if (isLocalhost(WEBHOOK_CALLBACK_URL)) {
      logStep('Webhook skipped — WEBHOOK_CALLBACK_URL is localhost (Bluebeam requires public HTTPS)', 'warn');
      logStep('Set WEBHOOK_CALLBACK_URL to a public HTTPS URL (e.g. ngrok) to enable webhooks', 'warn');
      return res.json({ success: true, skipped: true, state: pocState });
    }

    logStep(`Registering webhook for session ${pocState.sessionId}...`, 'info');
    const accessToken = await tokenManager.getValidAccessToken();

    const resp = await fetch(`${API_V2}/subscriptions`, {
      method: 'POST',
      headers: authHeaders(accessToken),
      body: JSON.stringify({
        sourceType: 'session',
        resourceId: pocState.sessionId,
        callbackURI: WEBHOOK_CALLBACK_URL
      })
    });

    if (!resp.ok) {
      throw new Error(`Webhook registration failed: ${resp.status} - ${await resp.text()}`);
    }

    const data = await resp.json();
    pocState.subscriptionId = data.subscriptionId;

    logStep(`Webhook registered: subscriptionId=${pocState.subscriptionId}`, 'success');
    logStep(`Callback URL: ${WEBHOOK_CALLBACK_URL}`, 'info');
    res.json({ success: true, subscriptionId: pocState.subscriptionId, state: pocState });
  } catch (err) {
    pocState.status = 'error';
    logStep(err.message, 'error');
    res.status(500).json({ error: err.message });
  }
});

// -----------------------------------------------------------------------------
// STEP 4 — Checkout to Session
// -----------------------------------------------------------------------------
app.post('/poc/checkout-to-session', async (req, res) => {
  try {
    if (!pocState.sessionId) throw new Error('No active session — run create-session first');
    if (pocState.projectFiles.length === 0) throw new Error('No project files — run upload-to-project first');

    pocState.status = 'checking-out';
    logStep(`Checking ${pocState.projectFiles.length} file(s) out to session ${pocState.sessionId}...`, 'info');

    const accessToken = await tokenManager.getValidAccessToken();
    const checked = [];

    for (const pf of pocState.projectFiles) {
      logStep(`Checking out "${pf.name}" (projectFileId=${pf.projectFileId})...`, 'info');

      const resp = await fetch(
        `${API_V1}/projects/${POC_PROJECT_ID}/files/${pf.projectFileId}/checkout-to-session`,
        {
          method: 'POST',
          headers: authHeaders(accessToken),
          body: JSON.stringify({ SessionId: pocState.sessionId })
        }
      );

      if (!resp.ok) {
        const err = await resp.text();

        if (resp.status === 409) {
          logStep(`"${pf.name}" already checked out (409) — attempting to release and retry...`, 'warn');

          const releaseResp = await fetch(
            `${API_V1}/projects/${POC_PROJECT_ID}/files/${pf.projectFileId}/checkout`,
            { method: 'DELETE', headers: authHeaders(accessToken) }
          );

          if (releaseResp.ok) {
            logStep(`Checkout released for "${pf.name}" — retrying...`, 'info');

            const retry = await fetch(
              `${API_V1}/projects/${POC_PROJECT_ID}/files/${pf.projectFileId}/checkout-to-session`,
              {
                method: 'POST',
                headers: authHeaders(accessToken),
                body: JSON.stringify({ SessionId: pocState.sessionId })
              }
            );

            if (!retry.ok) {
              logStep(`Retry failed for "${pf.name}": ${retry.status}`, 'warn');
              continue;
            }
          } else {
            logStep(`Could not release checkout for "${pf.name}"`, 'warn');
            continue;
          }
        } else {
          logStep(`Checkout failed for "${pf.name}": ${resp.status} - ${err}`, 'warn');
          continue;
        }
      }

      await sleep(1000);

      const sessionFilesResp = await fetch(
        `${API_V1}/sessions/${pocState.sessionId}/files?includeDeleted=false`,
        { headers: authHeaders(accessToken) }
      );

      if (!sessionFilesResp.ok) {
        logStep(`Could not list session files after checkout: ${sessionFilesResp.status}`, 'warn');
        continue;
      }

      const sessionFilesData = await sessionFilesResp.json();
      const sessionFiles = sessionFilesData.SessionFiles || sessionFilesData.Files || [];

      const match = sessionFiles.find(f =>
        f.ProjectFileId === pf.projectFileId || f.Name === pf.name
      );

      if (!match) {
        logStep(`Could not find session file entry for "${pf.name}" after checkout`, 'warn');
        continue;
      }

      const entry = {
        sessionFileId: match.Id,
        projectFileId: pf.projectFileId,
        name: pf.name
      };

      pocState.sessionFileIds.push(entry);
      checked.push(entry);

      logStep(`"${pf.name}" checked out to session (sessionFileId=${match.Id})`, 'success');
    }

    logStep(`${checked.length} file(s) checked out to session`, 'success');
    res.json({ success: true, checked, state: pocState });
  } catch (err) {
    pocState.status = 'error';
    logStep(err.message, 'error');
    res.status(500).json({ error: err.message });
  }
});

// -----------------------------------------------------------------------------
// STEP 5 — Invite Reviewers
// -----------------------------------------------------------------------------
app.post('/poc/invite-reviewers', async (req, res) => {
  try {
    if (!pocState.sessionId) throw new Error('No active session');

    pocState.status = 'inviting';
    logStep(`Inviting ${demoStub.reviewers.length} reviewer(s)...`, 'info');

    const accessToken = await tokenManager.getValidAccessToken();
    const results = [];

    for (const reviewer of demoStub.reviewers) {
      const endpoint = reviewer.hasStudioAccount
        ? `${API_V1}/sessions/${pocState.sessionId}/users`
        : `${API_V1}/sessions/${pocState.sessionId}/invite`;

      logStep(
        `Inviting ${reviewer.email} via ${reviewer.hasStudioAccount ? 'direct-add' : 'email-invite'}`,
        'info'
      );

      const resp = await fetch(endpoint, {
        method: 'POST',
        headers: authHeaders(accessToken),
        body: JSON.stringify({
          Email: reviewer.email,
          Message: `You have been invited to review ${demoStub.documentId}: ${demoStub.description}`
        })
      });

      if (!resp.ok) {
        const err = await resp.text();
        logStep(`Failed to invite ${reviewer.email}: ${resp.status} - ${err}`, 'warn');
        results.push({ email: reviewer.email, success: false, error: err });
      } else {
        logStep(`Invited: ${reviewer.email}`, 'success');
        results.push({ email: reviewer.email, success: true });
      }
    }

    pocState.status = 'active';
    logStep('Session active — reviewers notified', 'success');
    logStep(`Join in Bluebeam Revu with session ID: ${pocState.sessionId}`, 'info');
    res.json({ success: true, results, state: pocState });
  } catch (err) {
    pocState.status = 'error';
    logStep(err.message, 'error');
    res.status(500).json({ error: err.message });
  }
});

// -----------------------------------------------------------------------------
// STEP 7 — Check In
// -----------------------------------------------------------------------------
app.post('/poc/checkin', async (req, res) => {
  try {
    const accessToken = await tokenManager.getValidAccessToken();
    const results = await performCheckin(accessToken);
    res.json({ success: true, results, state: pocState });
  } catch (err) {
    pocState.status = 'error';
    logStep(err.message, 'error');
    res.status(500).json({ error: err.message });
  }
});

// -----------------------------------------------------------------------------
// STEP 8 — Export Markups
// -----------------------------------------------------------------------------
app.post('/poc/export-markups', async (req, res) => {
  try {
    const accessToken = await tokenManager.getValidAccessToken();
    const results = await performExportMarkups(accessToken);

    res.json({
      success: true,
      results,
      markupExports: pocState.markupExports,
      state: pocState
    });
  } catch (err) {
    pocState.status = 'error';
    logStep(err.message, 'error');
    res.status(500).json({ error: err.message });
  }
});

// -----------------------------------------------------------------------------
// STEP 9 — Compatibility Route (now XML-backed extraction)
// -----------------------------------------------------------------------------
app.post('/poc/run-markuplist-job', async (req, res) => {
  try {
    const accessToken = await tokenManager.getValidAccessToken();

    if (!pocState.markupExports.length) {
      logStep('No exported XML tracked yet — running export-markups first...', 'info');
      await performExportMarkups(accessToken);
    }

    const markups = await performMarkupExtractionFromXml(accessToken);

    res.json({
      success: true,
      count: markups.length,
      markups,
      extractionMode: 'exportmarkups-xml',
      state: pocState
    });
  } catch (err) {
    pocState.status = 'error';
    logStep(err.message, 'error');
    res.status(500).json({ error: err.message });
  }
});

// -----------------------------------------------------------------------------
// STEP 9b — Combined Downstream Processing
// -----------------------------------------------------------------------------
app.post('/poc/downstream-process', async (req, res) => {
  try {
    if (!pocState.sessionId) throw new Error('No active session');
    if (pocState.sessionFileIds.length === 0) throw new Error('No session files — run checkout-to-session first');

    logStep('Starting downstream processing...', 'info');
    const accessToken = await tokenManager.getValidAccessToken();

    const checkinResults = await performCheckin(accessToken);

    logStep('Global post-checkin settle delay: 5s...', 'info');
    await sleep(5000);

    const exportResults = await performExportMarkups(accessToken);

    logStep('Global post-export settle delay: 5s...', 'info');
    await sleep(5000);

    const markups = await performMarkupExtractionFromXml(accessToken);

    logStep('Downstream processing complete', 'success');

    res.json({
      success: true,
      checkinResults,
      exportResults,
      count: markups.length,
      markups,
      extractionMode: 'exportmarkups-xml',
      markupExports: pocState.markupExports,
      state: pocState
    });
  } catch (err) {
    pocState.status = 'error';
    logStep(err.message, 'error');
    res.status(500).json({ error: err.message });
  }
});

// -----------------------------------------------------------------------------
// STEP 10 — Finalize Session
// -----------------------------------------------------------------------------
app.post('/poc/finalize', async (req, res) => {
  try {
    if (!pocState.sessionId) throw new Error('No active session');

    pocState.status = 'finalizing';
    logStep(`Setting session ${pocState.sessionId} to Finalizing...`, 'info');

    const accessToken = await tokenManager.getValidAccessToken();
    const resp = await fetch(`${API_V1}/sessions/${pocState.sessionId}`, {
      method: 'PUT',
      headers: authHeaders(accessToken),
      body: JSON.stringify({
        Name: `${demoStub.documentId}_Review_${new Date().toISOString().slice(0, 10)}`,
        Restricted: true,
        SessionEndDate: demoStub.sessionEndDate
      })
    });

    if (!resp.ok) {
      throw new Error(`Finalize failed: ${resp.status} - ${await resp.text()}`);
    }

    logStep('Session finalized', 'success');
    res.json({ success: true, state: pocState });
  } catch (err) {
    pocState.status = 'error';
    logStep(err.message, 'error');
    res.status(500).json({ error: err.message });
  }
});

// -----------------------------------------------------------------------------
// STEP 11 — Snapshot
// -----------------------------------------------------------------------------
app.post('/poc/snapshot', async (req, res) => {
  try {
    if (!pocState.sessionId || pocState.sessionFileIds.length === 0) {
      throw new Error('No active session or no files');
    }

    pocState.status = 'snapshotting';
    const accessToken = await tokenManager.getValidAccessToken();
    const downloads = [];

    for (const sf of pocState.sessionFileIds) {
      logStep(`Requesting snapshot for "${sf.name}"...`, 'info');

      const snapResp = await fetch(
        `${API_V1}/sessions/${pocState.sessionId}/files/${sf.sessionFileId}/snapshot`,
        { method: 'POST', headers: authHeaders(accessToken) }
      );

      if (!snapResp.ok) {
        logStep(`Snapshot request failed for "${sf.name}": ${snapResp.status}`, 'warn');
        continue;
      }

      let downloadUrl = null;

      for (let i = 0; i < 20; i++) {
        await sleep(5000);

        const pollToken = await tokenManager.getValidAccessToken();
        const pollResp = await fetch(
          `${API_V1}/sessions/${pocState.sessionId}/files/${sf.sessionFileId}/snapshot`,
          { headers: authHeaders(pollToken) }
        );

        if (!pollResp.ok) continue;

        const d = await pollResp.json();
        logStep(`Snapshot poll ${i + 1}: ${d.Status}`, 'info');

        if (d.Status === 'Complete') {
          downloadUrl = d.DownloadUrl;
          break;
        }
        if (d.Status === 'Error') {
          throw new Error(`Snapshot error: ${d.Message}`);
        }
      }

      if (!downloadUrl) {
        logStep(`Snapshot timed out for "${sf.name}"`, 'warn');
        continue;
      }

      const dlResp = await fetch(downloadUrl);
      if (!dlResp.ok) throw new Error(`Download failed: ${dlResp.status}`);

      const pdfBuffer = await dlResp.buffer();

      const publicDir = path.join(__dirname, 'public');
      if (!fs.existsSync(publicDir)) fs.mkdirSync(publicDir, { recursive: true });

      const outFile = `${demoStub.documentId}_${sf.name.replace(/\.[^.]+$/, '')}_Reviewed.pdf`;
      fs.writeFileSync(path.join(publicDir, outFile), pdfBuffer);

      logStep(`PDF saved: ${outFile} (${pdfBuffer.length} bytes)`, 'success');
      downloads.push({ name: outFile, path: `/${outFile}`, size: pdfBuffer.length });
    }

    pocState.status = 'complete';
    logStep('Snapshots complete', 'success');
    res.json({ success: true, downloads, state: pocState });
  } catch (err) {
    pocState.status = 'error';
    logStep(err.message, 'error');
    res.status(500).json({ error: err.message });
  }
});

// -----------------------------------------------------------------------------
// STEP 12 — Cleanup
// -----------------------------------------------------------------------------
app.post('/poc/cleanup', async (req, res) => {
  try {
    if (!pocState.sessionId) throw new Error('No active session to clean up');

    const accessToken = await tokenManager.getValidAccessToken();

    if (pocState.subscriptionId) {
      const subResp = await fetch(`${API_V2}/subscriptions/${pocState.subscriptionId}`, {
        method: 'DELETE',
        headers: authHeaders(accessToken)
      });

      logStep(
        subResp.ok ? 'Webhook subscription deleted' : `Sub delete: ${subResp.status}`,
        subResp.ok ? 'success' : 'warn'
      );
    }

    const sessResp = await fetch(`${API_V1}/sessions/${pocState.sessionId}`, {
      method: 'DELETE',
      headers: authHeaders(accessToken)
    });

    logStep(
      sessResp.ok ? 'Session deleted' : `Session delete: ${sessResp.status}`,
      sessResp.ok ? 'success' : 'warn'
    );

    logStep('Cleanup complete', 'success');
    pocState.sessionId = null;
    pocState.subscriptionId = null;

    res.json({ success: true, state: pocState });
  } catch (err) {
    logStep(err.message, 'error');
    res.status(500).json({ error: err.message });
  }
});

// -----------------------------------------------------------------------------
// WEBHOOK LISTENER
// -----------------------------------------------------------------------------
app.post('/webhook/studio-events', (req, res) => {
  const p = req.body || {};
  logStep(`Webhook: ${p.ResourceType || 'unknown'} / ${p.EventType || 'unknown'}`, 'webhook');
  pocState.webhookEvents.push({ ...p, receivedAt: new Date().toISOString() });
  res.sendStatus(200);
});

// -----------------------------------------------------------------------------
// STANDALONE ENDPOINTS
// -----------------------------------------------------------------------------
app.get('/api/project-markups', (req, res) => {
  if (!pocState.markups.length) {
    return res.status(404).json({ error: 'No markup data. Run downstream processing first.' });
  }

  res.json(
    pocState.markups.map(m => ({
      MarkupId: m.Id,
      Author: m.Author,
      Type: m.Type,
      Subject: m.Subject,
      Comment: m.Comment,
      Status: m.Status,
      Layer: m.Layer,
      Page: m.Page,
      DateCreated: m.DateCreated,
      DateModified: m.DateModified,
      Color: m.Color,
      Checked: m.Checked,
      Locked: m.Locked,
      ExtendedProperties: m.ExtendedProperties || {},
      SourceFile: m._sourceFile
    }))
  );
});

// -----------------------------------------------------------------------------
// START
// -----------------------------------------------------------------------------
app.listen(PORT, () => {
  console.log(`\nBluebeam Studio PoC  →  http://localhost:${PORT}`);
  console.log(`Project: ${POC_PROJECT_ID}`);

  if (isLocalhost(WEBHOOK_CALLBACK_URL)) {
    console.log('⚠  Webhook will be skipped (localhost URL)');
  }

  console.log(`\nFLOW:`);
  console.log(`  POST /poc/setup-project         — 0a: Create folders + upload custom-columns.xml`);
  console.log(`  POST /poc/upload-to-project     — 0b: Upload PDFs from UI → project`);
  console.log(`  POST /poc/apply-custom-columns  — 0c: Apply custom columns to project files (optional)`);
  console.log(`  POST /poc/trigger               — 1:  Workflow event`);
  console.log(`  POST /poc/create-session        — 2:  Create session`);
  console.log(`  POST /poc/register-webhook      — 3:  Register webhook`);
  console.log(`  POST /poc/checkout-to-session   — 4:  Check out files into session`);
  console.log(`  POST /poc/invite-reviewers      — 5:  Invite reviewers`);
  console.log(`       (6: Review in Revu)`);
  console.log(`  POST /poc/checkin               — 7:  Check in session files`);
  console.log(`  POST /poc/export-markups        — 8:  Export markups to XML`);
  console.log(`  POST /poc/run-markuplist-job    — 9:  Compatibility route, XML-backed extraction`);
  console.log(`  POST /poc/downstream-process    — 9b: Check in + export + XML extract`);
  console.log(`  POST /poc/finalize              — 10: Finalize session`);
  console.log(`  POST /poc/snapshot              — 11: Snapshot + download PDF`);
  console.log(`  POST /poc/cleanup               — 12: Delete webhook + session\n`);
});
