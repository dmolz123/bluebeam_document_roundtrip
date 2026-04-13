/**
 * Bluebeam Studio API — Document Roundtrip PoC
 * Proof-of-concept reference implementation. Not for production use.
 *
 * Changes vs. previous version:
 *   - Dynamic project creation (POST /v1/projects) — new project every run
 *   - Project-level permissions (PUT /v1/projects/{id}/permissions)
 *   - Per-folder permissions (POST /v1/projects/{id}/folders/{folderId}/permissions)
 *   - Per-user permissions (PUT /v1/projects/{id}/users/{userId}/permissions)
 *   - GET /v1/projects/{id}/users to retrieve user IDs after invite
 *   - CRC: null (B&McD spec — let AWS calculate)
 *   - Checkin body as x-www-form-urlencoded per B&McD spec
 *   - State model injection via pdf-lib (pre-upload, wipes known custom models first)
 *   - POST /poc/inject-state-model exposed as standalone endpoint
 *
 * Roundtrip flow:
 *   0a. /poc/setup-project          — Create project + folders + upload custom-columns.xml
 *   0b. /poc/upload-to-project      — Inject state model + upload PDF(s) to project
 *   0c. /poc/apply-custom-columns   — Apply custom columns (optional)
 *   1.  /poc/trigger                — Simulate source-system workflow event
 *   2.  /poc/create-session         — Create Studio Session
 *   3.  /poc/register-webhook       — Subscribe to session events
 *   4.  /poc/checkout-to-session    — Check project file(s) out into session
 *   5.  /poc/invite-reviewers       — Invite reviewers + fetch user IDs + set permissions
 *   6.  (Review in Bluebeam Revu)
 *   7.  /poc/checkin                — Check session files back into project
 *   8.  /poc/export-markups         — exportmarkups job → XML
 *   9.  /poc/run-markuplist-job     — XML-backed markup extraction
 *   9b. /poc/downstream-process     — Combined: checkin + export + extract
 *   10. /poc/finalize               — Finalize session
 *   11. /poc/snapshot               — Snapshot + download PDF
 *   12. /poc/cleanup                — Delete webhook + session
 */

require('dotenv').config();
const express            = require('express');
const cors               = require('cors');
const fs                 = require('fs');
const path               = require('path');
const multer             = require('multer');
const { parseStringPromise } = require('xml2js');
const { PDFDocument, PDFName, PDFDict, PDFArray, PDFString } = require('pdf-lib');
const TokenManager       = require('./tokenManager');

const app    = express();
const PORT   = process.env.PORT || 3000;
const upload = multer({ storage: multer.memoryStorage() });

// -----------------------------------------------------------------------------
// API CONFIGURATION
// -----------------------------------------------------------------------------
const API_V1    = 'https://api.bluebeam.com/publicapi/v1';
const API_V2    = 'https://api.bluebeam.com/publicapi/v2';
const CLIENT_ID = process.env.BB_CLIENT_ID;

const WEBHOOK_CALLBACK_URL =
  process.env.WEBHOOK_CALLBACK_URL ||
  `http://localhost:${PORT}/webhook/studio-events`;

// Folder names created inside each project
const FOLDER_RESOURCES      = 'resources';
const FOLDER_REVIEW_DOCS    = 'review-documents';
const FOLDER_MARKUP_EXPORTS = 'markup-exports';

// Custom columns XML path
const CUSTOM_COLUMNS_XML_PATH = path.join(__dirname, 'resources', 'custom-columns.xml');

// ---------------------------------------------------------------------------
// STATE MODEL — 5-step QC Review
// ---------------------------------------------------------------------------
// cName identifiers of custom models to REMOVE before injecting the correct one.
// Add any incorrect/demo model names here so they are wiped on each run.
const STATE_MODELS_TO_REMOVE = [
  '5_step_QC_Review',       // our own model — ensures idempotency on re-runs
  'Incorrect_Review_Model', // demo "wrong" model that will be pre-loaded to show removal
  'Bad_Model',
  'Old_Review'
];

const QC_STATE_MODEL = {
  cName:   '5_step_QC_Review',
  cUIName: '5-step QC Review',
  states: [
    { key: 'Step3_Agree',             label: 'Step 3 - Agree' },
    { key: 'Step3_Disagree',          label: 'Step 3 - Disagree' },
    { key: 'Step3_Address_Future',    label: 'Step 3 - Address in Future Submittal' },
    { key: 'Step4_Revisions_Made',    label: 'Step 4 - Revisions Made' },
    { key: 'Step5_Verified_Concur',   label: 'Step 5 - Revisions Verified/Concur' },
    { key: 'Step5_Incomplete_Disagr', label: 'Step 5 - Revisions Incomplete/Disagree' }
  ],
  defaultState: 'Step3_Agree'
};

// ---------------------------------------------------------------------------
// PDF STATE MODEL INJECTION (pdf-lib — no Bluebeam API required)
// ---------------------------------------------------------------------------

/**
 * Injects a Collab.addStateModel() call as a document-level JS action.
 * First removes known custom models so the PDF always has a clean state.
 * Bluebeam default models (Review, Migration) cannot be removed and are left alone.
 */
async function injectStateModel(pdfBuffer) {
  const pdfDoc = await PDFDocument.load(pdfBuffer, { ignoreEncryption: true });
  const ctx    = pdfDoc.context;
  const cat    = pdfDoc.catalog;

  // Build removal calls for all known custom models
  const removeCalls = STATE_MODELS_TO_REMOVE
    .map(n => `try{Collab.removeStateModel("${n}");}catch(e){}`)
    .join('\n');

  // Build oStates object literal
  const statesObj = QC_STATE_MODEL.states
    .map(s => `    "${s.key}": { cUIName: "${s.label}" }`)
    .join(',\n');

  const js = `${removeCalls}
try {
  Collab.addStateModel({
    cName:   "${QC_STATE_MODEL.cName}",
    cUIName: "${QC_STATE_MODEL.cUIName}",
    oStates: {
${statesObj}
    },
    cDefault: "${QC_STATE_MODEL.defaultState}"
  });
} catch(e) {}`.trim();

  const jsBytes  = Buffer.from(js, 'utf-8');
  const jsStream = ctx.stream(jsBytes, { Type: 'JavaScript', Length: jsBytes.length });
  const jsRef    = ctx.register(jsStream);

  // Names tree entry (document-level JS)
  const modelKey  = `BB_StateModel_${QC_STATE_MODEL.cName}`;
  const jsAction  = ctx.obj({ S: PDFName.of('JavaScript'), JS: jsRef });
  const jsActRef  = ctx.register(jsAction);

  let namesDict = cat.lookupMaybe(PDFName.of('Names'), PDFDict);
  if (!namesDict) {
    const ref = ctx.register(ctx.obj({}));
    cat.set(PDFName.of('Names'), ref);
    namesDict = ctx.lookup(ref, PDFDict);
  }

  let jsNamesDict = namesDict.lookupMaybe(PDFName.of('JavaScript'), PDFDict);
  if (!jsNamesDict) {
    const ref = ctx.register(ctx.obj({}));
    namesDict.set(PDFName.of('JavaScript'), ref);
    jsNamesDict = ctx.lookup(ref, PDFDict);
  }

  const existing = jsNamesDict.lookupMaybe(PDFName.of('Names'), PDFArray);
  if (existing) {
    existing.push(PDFString.of(modelKey));
    existing.push(jsActRef);
  } else {
    jsNamesDict.set(PDFName.of('Names'), ctx.obj([PDFString.of(modelKey), jsActRef]));
  }

  // Also set as OpenAction (belt + suspenders)
  cat.set(PDFName.of('OpenAction'), ctx.register(ctx.obj({ S: PDFName.of('JavaScript'), JS: jsRef })));

  return Buffer.from(await pdfDoc.save());
}

// -----------------------------------------------------------------------------
// DEMO STUB
// -----------------------------------------------------------------------------
let demoStub = {
  projectName:    process.env.DEMO_PROJECT_NAME || 'Demo Project',
  documentId:     process.env.DEMO_DOCUMENT_ID  || 'DOC-001',
  description:    process.env.DEMO_DESCRIPTION  || 'Design review — coordination update',
  reviewers:      [{ email: 'dmolz@bluebeam.com', hasStudioAccount: true }],
  sessionEndDate: new Date(Date.now() + 7 * 24 * 60 * 60 * 1000).toISOString()
};

// -----------------------------------------------------------------------------
// IN-MEMORY STATE
// -----------------------------------------------------------------------------
let pocState = {
  projectId:           null,   // created dynamically each run
  sessionId:           null,
  subscriptionId:      null,
  projectSetupDone:    false,
  folderIds:           {},
  customColumnsFileId: null,
  projectFiles:        [],
  sessionFileIds:      [],
  projectUserIds:      [],     // fetched after invite, used for per-user permissions
  markupExports:       [],
  markups:             [],
  markupJobId:         null,
  status:              'idle',
  log:                 [],
  createdAt:           null,
  webhookEvents:       []
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
  import('node-fetch').then(({ default: f }) => f(...args));

const tokenManager = new TokenManager();

// -----------------------------------------------------------------------------
// HELPERS
// -----------------------------------------------------------------------------
function resetPocState() {
  pocState = {
    projectId:           null,
    sessionId:           null,
    subscriptionId:      null,
    projectSetupDone:    false,
    folderIds:           {},
    customColumnsFileId: null,
    projectFiles:        [],
    sessionFileIds:      [],
    projectUserIds:      [],
    markupExports:       [],
    markups:             [],
    markupJobId:         null,
    status:              'idle',
    log:                 [],
    createdAt:           null,
    webhookEvents:       []
  };
}

function isLocalhost(url) {
  return /^https?:\/\/(localhost|127\.0\.0\.1)/.test(url);
}

function authHeaders(accessToken, extra = {}) {
  return {
    Authorization:  `Bearer ${accessToken}`,
    'client_id':    CLIENT_ID,
    'Content-Type': 'application/json',
    Accept:         'application/json',
    ...extra
  };
}

function sleep(ms) {
  return new Promise(r => setTimeout(r, ms));
}

// -----------------------------------------------------------------------------
// JOB POLLER
// Status codes: 100=Queued, 130=Running, 150=Finishing, 200=Success
// -----------------------------------------------------------------------------
async function pollJob(url, headers, maxAttempts = 20, intervalMs = 3000) {
  const inProgress = new Set([100, 130, 150]);
  for (let i = 1; i <= maxAttempts; i++) {
    await sleep(intervalMs);
    const res  = await fetch(url, { headers });
    const data = await res.json();
    const status = data.Status ?? data.JobStatus;
    const msg    = data.StatusMessage ?? data.JobStatusMessage ?? '';
    logStep(`Job poll ${i}/${maxAttempts}: status=${status} ${msg}`.trim(), 'info');
    if (status === 200) return data;
    if (!inProgress.has(status)) throw new Error(`Job failed (status=${status}): ${msg}`);
  }
  throw new Error(`Job did not complete after ${maxAttempts} attempts`);
}

// -----------------------------------------------------------------------------
// PROJECT HELPERS
// -----------------------------------------------------------------------------

/** Create a new Studio Project. Returns the new project ID string. */
async function createProject(name, accessToken) {
  const resp = await fetch(`${API_V1}/projects`, {
    method:  'POST',
    headers: authHeaders(accessToken),
    body:    JSON.stringify({ Name: name, Notification: false, Restricted: true })
  });
  if (!resp.ok) throw new Error(`Failed to create project: ${resp.status} - ${await resp.text()}`);
  const data = await resp.json();
  return data.Id; // format "123-456-789"
}

/** Set overall project permissions. */
async function setProjectPermission(projectId, type, allow, accessToken) {
  const resp = await fetch(`${API_V1}/projects/${projectId}/permissions`, {
    method:  'PUT',
    headers: authHeaders(accessToken),
    body:    JSON.stringify({ Type: type, Allow: allow })
  });
  if (!resp.ok) {
    logStep(`Project permission ${type}=${allow} returned ${resp.status}`, 'warn');
  }
}

/** List all folders in a project. */
async function listProjectFolders(projectId, accessToken) {
  const resp = await fetch(`${API_V1}/projects/${projectId}/folders`, {
    headers: authHeaders(accessToken)
  });
  if (!resp.ok) throw new Error(`Failed to list folders: ${resp.status} - ${await resp.text()}`);
  const data = await resp.json();
  return data.ProjectFolders || [];
}

/** Create a folder. parentFolderId=0 places it at root. */
async function createFolder(projectId, name, parentFolderId = 0, accessToken) {
  const resp = await fetch(`${API_V1}/projects/${projectId}/folders`, {
    method:  'POST',
    headers: authHeaders(accessToken),
    body:    JSON.stringify({ Name: name, ParentFolderId: parentFolderId, Comment: '' })
  });
  if (!resp.ok) throw new Error(`Failed to create folder "${name}": ${resp.status} - ${await resp.text()}`);
  const data = await resp.json();
  return data.Id;
}

/** Set folder permissions for a specific user. */
async function setFolderPermission(projectId, folderId, userId, permission, accessToken) {
  const resp = await fetch(`${API_V1}/projects/${projectId}/folders/${folderId}/permissions`, {
    method:  'POST',
    headers: authHeaders(accessToken),
    body:    JSON.stringify({ UserId: userId, Permission: permission })
  });
  if (!resp.ok) {
    logStep(`Folder permission for user ${userId} returned ${resp.status}`, 'warn');
  }
}

/** Set per-user project permissions. */
async function setUserPermission(projectId, userId, type, allow, accessToken) {
  const resp = await fetch(`${API_V1}/projects/${projectId}/users/${userId}/permissions`, {
    method:  'PUT',
    headers: authHeaders(accessToken),
    body:    JSON.stringify({ Type: type, Allow: allow })
  });
  if (!resp.ok) {
    logStep(`User permission ${type}=${allow} for user ${userId} returned ${resp.status}`, 'warn');
  }
}

/** Invite a user to the project. */
async function inviteProjectUser(projectId, email, accessToken) {
  const resp = await fetch(`${API_V1}/projects/${projectId}/users`, {
    method:  'POST',
    headers: authHeaders(accessToken),
    body:    JSON.stringify({ Email: email, SendEmail: true, Message: '' })
  });
  if (!resp.ok) {
    logStep(`Project invite for ${email} returned ${resp.status}`, 'warn');
  }
}

/** Get all project users. Returns array of { Id, Email, Name, IsProjectOwner }. */
async function getProjectUsers(projectId, accessToken) {
  const resp = await fetch(`${API_V1}/projects/${projectId}/users`, {
    headers: authHeaders(accessToken)
  });
  if (!resp.ok) throw new Error(`Failed to get project users: ${resp.status} - ${await resp.text()}`);
  const data = await resp.json();
  return data.ProjectUsers || [];
}

/**
 * Upload a file to the project (3-step).
 * Per B&McD spec: CRC must be null (let AWS calculate).
 * S3 PUT must NOT include auth headers.
 */
async function uploadFileToProject(fileBuffer, fileName, projectId, accessToken, folderId = null) {
  logStep(`Uploading "${fileName}" to project ${projectId}${folderId ? ` (folderId=${folderId})` : ''}...`, 'info');

  const metaBody = { Name: fileName, CRC: null };
  if (folderId) metaBody.ParentFolderId = folderId;

  const metaResp = await fetch(`${API_V1}/projects/${projectId}/files`, {
    method:  'POST',
    headers: authHeaders(accessToken),
    body:    JSON.stringify(metaBody)
  });
  if (!metaResp.ok)
    throw new Error(`Metadata block failed for "${fileName}": ${metaResp.status} - ${await metaResp.text()}`);

  const meta              = await metaResp.json();
  const projectFileId     = meta.Id;
  const uploadUrl         = meta.UploadUrl;
  const uploadContentType = meta.UploadContentType || 'application/pdf';

  logStep(`Metadata block created: projectFileId=${projectFileId}`, 'success');

  // S3 PUT — no auth headers
  const s3Resp = await fetch(uploadUrl, {
    method: 'PUT',
    headers: { 'Content-Type': uploadContentType, 'x-amz-server-side-encryption': 'AES256' },
    body:   fileBuffer
  });
  if (!s3Resp.ok)
    throw new Error(`S3 upload failed for "${fileName}": ${s3Resp.status}`);

  logStep('S3 upload complete', 'success');

  const confirmResp = await fetch(
    `${API_V1}/projects/${projectId}/files/${projectFileId}/confirm-upload`,
    { method: 'POST', headers: authHeaders(accessToken), body: '{}' }
  );
  if (!confirmResp.ok)
    throw new Error(`Confirm upload failed for "${fileName}": ${confirmResp.status} - ${await confirmResp.text()}`);

  logStep(`"${fileName}" confirmed in project (projectFileId=${projectFileId})`, 'success');
  return { projectFileId, name: fileName, size: fileBuffer.length, folderId };
}

/** List all files in the project. */
async function listProjectFiles(projectId, accessToken) {
  const resp = await fetch(`${API_V1}/projects/${projectId}/files`, {
    headers: authHeaders(accessToken)
  });
  if (!resp.ok)
    throw new Error(`Failed to list project files: ${resp.status} - ${await resp.text()}`);
  const data = await resp.json();
  return data.ProjectFiles || [];
}

async function getProjectFileByPath(projectId, accessToken, filePath) {
  const url  = `${API_V1}/projects/${projectId}/files/by-path?path=${encodeURIComponent(filePath)}`;
  const resp = await fetch(url, { headers: authHeaders(accessToken) });
  if (!resp.ok)
    throw new Error(`Failed to get file by path: ${resp.status} - ${await resp.text()}`);
  return resp.json();
}

async function waitForProjectFileSettlement(projectId, accessToken, projectFileId, fileName) {
  logStep(`Waiting for project file settlement: "${fileName}"...`, 'info');
  for (let i = 1; i <= 6; i++) {
    const files = await listProjectFiles(projectId, accessToken);
    const match = files.find(f => String(f.Id) === String(projectFileId));
    if (match) {
      const co = match.IsCheckedOut === true || match.CheckedOut === true;
      logStep(`Settlement check ${i}/6: found, checkedOut=${co}`, 'info');
      if (!co) { await sleep(5000); return match; }
    }
    await sleep(5000);
  }
  logStep('Settlement window elapsed — proceeding', 'warn');
  await sleep(5000);
}

// -----------------------------------------------------------------------------
// XML PARSE HELPERS
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
  if (val && typeof val === 'object') return typeof val._ !== 'undefined' ? val._ : '';
  return val;
}
function normalizeMarkupRecord(record, sourceFile) {
  const mapped = lowerKeyMap(record);
  const known = {
    Id:           scalar(firstDefined(mapped, ['id','markupid'])),
    Author:       scalar(firstDefined(mapped, ['author','createdby','user'])),
    Type:         scalar(firstDefined(mapped, ['type','markuptype'])),
    Subject:      scalar(firstDefined(mapped, ['subject','label','title'])),
    Comment:      scalar(firstDefined(mapped, ['comment','comments','note','contents'])),
    Status:       scalar(firstDefined(mapped, ['status','state'])),
    Layer:        scalar(firstDefined(mapped, ['layer'])),
    Page:         scalar(firstDefined(mapped, ['page','pagenumber','pageindex'])),
    DateCreated:  scalar(firstDefined(mapped, ['datecreated','creationdate','created'])),
    DateModified: scalar(firstDefined(mapped, ['datemodified','moddate','modified'])),
    Color:        scalar(firstDefined(mapped, ['color'])),
    Checked:      scalar(firstDefined(mapped, ['checked'])),
    Locked:       scalar(firstDefined(mapped, ['locked']))
  };

  // Status history — collect all History/StateHistory entries
  const historyRaw = mapped.history || mapped.statehistory || mapped.statushistory;
  const statusHistory = [];
  if (historyRaw) {
    const items = Array.isArray(historyRaw) ? historyRaw
      : (historyRaw.item ? (Array.isArray(historyRaw.item) ? historyRaw.item : [historyRaw.item]) : [historyRaw]);
    for (const item of items) {
      const h = lowerKeyMap(typeof item === 'object' ? item : {});
      const state  = scalar(firstDefined(h, ['state','status','value']));
      const who    = scalar(firstDefined(h, ['author','user','by']));
      const when   = scalar(firstDefined(h, ['date','time','timestamp','datetime']));
      if (state || who) statusHistory.push({ state, author: who, date: when });
    }
  }

  // Replies / comments with timestamps
  const repliesRaw = mapped.replies || mapped.reply || mapped.comments;
  const replies = [];
  if (repliesRaw) {
    const items = Array.isArray(repliesRaw) ? repliesRaw
      : (repliesRaw.item ? (Array.isArray(repliesRaw.item) ? repliesRaw.item : [repliesRaw.item]) : [repliesRaw]);
    for (const item of items) {
      const r = lowerKeyMap(typeof item === 'object' ? item : {});
      replies.push({
        author:  scalar(firstDefined(r, ['author','user','by'])),
        comment: scalar(firstDefined(r, ['comment','text','content','body'])),
        date:    scalar(firstDefined(r, ['date','time','timestamp','datetime']))
      });
    }
  }

  const skip = new Set([
    'id','markupid','author','createdby','user','type','markuptype','subject','label','title',
    'comment','comments','note','contents','status','state','layer','page','pagenumber','pageindex',
    'datecreated','creationdate','created','datemodified','moddate','modified','color','checked',
    'locked','custom','history','statehistory','statushistory','replies','reply'
  ]);

  const custom = {};
  if (mapped.custom && typeof mapped.custom === 'object' && !Array.isArray(mapped.custom)) {
    for (const [k, v] of Object.entries(mapped.custom)) {
      const sv = scalar(v);
      if (sv !== null && sv !== undefined && String(sv).trim()) custom[k] = String(sv).trim();
    }
  }

  const extended = {};
  for (const [k, v] of Object.entries(mapped)) {
    if (!skip.has(k)) {
      if (v && typeof v === 'object' && !Array.isArray(v)) {
        for (const [nk, nv] of Object.entries(v)) {
          const sv = scalar(nv);
          if (sv !== null && sv !== undefined && String(sv).trim()) extended[`${k}.${nk}`] = String(sv).trim();
        }
      } else {
        const sv = scalar(v);
        if (sv !== null && sv !== undefined && String(sv).trim()) extended[k] = String(sv).trim();
      }
    }
  }

  return {
    ...known,
    StatusHistory:      statusHistory,
    Replies:            replies,
    Custom:             custom,
    ExtendedProperties: { ...custom, ...extended },
    _sourceFile:        sourceFile
  };
}

function looksLikeMarkupRecord(obj) {
  if (!obj || typeof obj !== 'object' || Array.isArray(obj)) return false;
  const keys = Object.keys(lowerKeyMap(obj));
  return ['author','subject','comment','status','page','layer','type','markupid','id','markuptype']
    .filter(k => keys.includes(k)).length >= 2;
}

function extractMarkupCandidates(node, sourceFile, results = []) {
  if (Array.isArray(node)) { node.forEach(i => extractMarkupCandidates(i, sourceFile, results)); return results; }
  if (!node || typeof node !== 'object') return results;
  if (looksLikeMarkupRecord(node)) results.push(normalizeMarkupRecord(node, sourceFile));
  for (const v of Object.values(node)) extractMarkupCandidates(v, sourceFile, results);
  return results;
}

async function parseBluebeamExportXml(xmlText, sourceFile) {
  const parsed = await parseStringPromise(xmlText, { explicitArray: false, mergeAttrs: true, trim: true });
  const candidates = extractMarkupCandidates(parsed, sourceFile);
  const seen = new Set();
  return candidates.filter(m => {
    const key = [m.Id, m.Author, m.Subject, m.Comment, m.Page, m.DateCreated].join('|');
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

// -----------------------------------------------------------------------------
// DOWNSTREAM HELPERS
// -----------------------------------------------------------------------------
async function performCheckin(accessToken) {
  if (!pocState.sessionId)              throw new Error('No active session');
  if (!pocState.sessionFileIds.length)  throw new Error('No session files');

  pocState.status = 'checking-in';
  const results = [];

  for (const sf of pocState.sessionFileIds) {
    logStep(`Checking in "${sf.name}" (sessionFileId=${sf.sessionFileId})...`, 'info');

    // B&McD spec: checkin body is x-www-form-urlencoded
    const resp = await fetch(
      `${API_V1}/sessions/${pocState.sessionId}/files/${sf.sessionFileId}/checkin`,
      {
        method:  'POST',
        headers: {
          Authorization:  `Bearer ${accessToken}`,
          'client_id':    CLIENT_ID,
          'Content-Type': 'application/x-www-form-urlencoded'
        },
        body: `Comment=Session+markup+review+complete`
      }
    );

    if (!resp.ok) {
      const err = await resp.text();
      logStep(`Check-in failed for "${sf.name}": ${resp.status} - ${err}`, 'warn');
      results.push({ name: sf.name, success: false, error: err });
      continue;
    }

    logStep(`"${sf.name}" checked in`, 'success');
    await waitForProjectFileSettlement(pocState.projectId, accessToken, sf.projectFileId, sf.name);
    results.push({ name: sf.name, success: true });
  }

  logStep('Check-in complete', 'success');
  return results;
}

async function performExportMarkups(accessToken) {
  if (!pocState.sessionFileIds.length) throw new Error('No session files');

  logStep('Exporting markups to XML...', 'info');

  if (!pocState.folderIds[FOLDER_MARKUP_EXPORTS]) {
    const folders = await listProjectFolders(pocState.projectId, accessToken);
    const found   = folders.find(f => f.Name === FOLDER_MARKUP_EXPORTS);
    if (found) pocState.folderIds[FOLDER_MARKUP_EXPORTS] = found.Id;
    else throw new Error(`Folder "${FOLDER_MARKUP_EXPORTS}" not found — run setup-project first`);
  }

  const results = [];
  for (const sf of pocState.sessionFileIds) {
    const exportFileName = `Markups-${sf.projectFileId}.xml`;

    const jobResp = await fetch(
      `${API_V1}/projects/${pocState.projectId}/files/${sf.projectFileId}/jobs/exportmarkups`,
      {
        method:  'POST',
        headers: authHeaders(accessToken),
        body:    JSON.stringify({ OutputFileName: exportFileName, OutputPath: FOLDER_MARKUP_EXPORTS, Priority: 0 })
      }
    );

    if (!jobResp.ok) {
      const err = await jobResp.text();
      logStep(`exportmarkups submission failed: ${jobResp.status} - ${err}`, 'warn');
      results.push({ name: sf.name, success: false, error: err });
      continue;
    }

    const { Id: jobId } = await jobResp.json();
    logStep(`exportmarkups job ${jobId} — polling (15s interval)...`, 'success');
    await pollJob(`${API_V1}/jobs/${jobId}`, authHeaders(accessToken), 15, 15000);

    logStep(`Markup XML exported: ${exportFileName}`, 'success');
    await sleep(5000);

    const idx = pocState.markupExports.findIndex(m => m.exportFileName === exportFileName);
    const rec = { name: sf.name, exportFileName, projectPath: FOLDER_MARKUP_EXPORTS };
    if (idx >= 0) pocState.markupExports[idx] = rec; else pocState.markupExports.push(rec);
    results.push({ name: sf.name, success: true, exportFileName });
  }
  return results;
}

async function performMarkupExtractionFromXml(accessToken) {
  pocState.status  = 'extracting-markups';
  pocState.markups = [];

  for (const sf of pocState.sessionFileIds) {
    const exportFileName = `Markups-${sf.projectFileId}.xml`;
    logStep(`Downloading exported XML for "${sf.name}"...`, 'info');

    const xmlPath  = `/${FOLDER_MARKUP_EXPORTS}/${exportFileName}`;
    const fileMeta = await getProjectFileByPath(pocState.projectId, accessToken, xmlPath);
    if (!fileMeta.DownloadUrl) throw new Error(`DownloadUrl missing for ${xmlPath}`);

    const xmlResp = await fetch(fileMeta.DownloadUrl);
    if (!xmlResp.ok) throw new Error(`Failed to download XML: ${xmlResp.status}`);

    const xmlText    = await xmlResp.text();
    const fileMarkups = await parseBluebeamExportXml(xmlText, sf.name);
    pocState.markups.push(...fileMarkups);
    logStep(`"${sf.name}" — ${fileMarkups.length} markup(s) extracted`, 'success');
  }

  pocState.status = 'active';
  logStep(`Extraction complete — ${pocState.markups.length} total markup(s)`, 'success');
  return pocState.markups;
}

// -----------------------------------------------------------------------------
// HEALTH CHECK
// -----------------------------------------------------------------------------
app.get('/health', (req, res) => {
  res.json({
    status:    'healthy',
    projectId: pocState.projectId || '(created per run)',
    config: {
      hasClientId:            Boolean(CLIENT_ID),
      webhookCallbackUrl:     WEBHOOK_CALLBACK_URL,
      webhookIsLocalhost:     isLocalhost(WEBHOOK_CALLBACK_URL),
      customColumnsXmlExists: fs.existsSync(CUSTOM_COLUMNS_XML_PATH)
    }
  });
});

// =============================================================================
// POC ROUTES
// =============================================================================
app.get('/poc/state', (req, res) => res.json({ ...pocState, stub: demoStub }));
app.get('/poc/stub',  (req, res) => res.json(demoStub));

app.post('/poc/configure', (req, res) => {
  const { projectName, documentId, description, reviewerEmail } = req.body || {};
  if (projectName)   demoStub.projectName  = projectName;
  if (documentId)    demoStub.documentId   = documentId;
  if (description)   demoStub.description  = description;
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
  if (email === 'dmolz@bluebeam.com')
    return res.status(400).json({ error: 'Cannot remove primary reviewer' });
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
// Creates a brand new project, sets permissions, creates folders.
// -----------------------------------------------------------------------------
app.post('/poc/setup-project', async (req, res) => {
  try {
    if (!fs.existsSync(CUSTOM_COLUMNS_XML_PATH))
      throw new Error(`custom-columns.xml not found at ${CUSTOM_COLUMNS_XML_PATH}`);

    logStep('Running project setup...', 'info');
    const accessToken = await tokenManager.getValidAccessToken();

    // 1 — Create new project
    const projectName = demoStub.projectName || 'BB Roundtrip PoC';
    logStep(`Creating project "${projectName}"...`, 'info');
    const projectId = await createProject(projectName, accessToken);
    pocState.projectId = projectId;
    logStep(`Project created: ID=${projectId}`, 'success');

    // 2 — Set overall project permissions (B&McD step 5)
    const projectPermissions = [
      { type: 'CreateSessions',     allow: 'Allow' },
      { type: 'Invite',             allow: 'Allow' },
      { type: 'ManageParticipants', allow: 'Allow' },
      { type: 'ShareItems',         allow: 'Allow' }
    ];
    for (const p of projectPermissions) {
      await setProjectPermission(projectId, p.type, p.allow, accessToken);
      logStep(`Project permission: ${p.type}=${p.allow}`, 'info');
    }

    // 3 — Create folders (1500ms delay after each per developer guide)
    const needed = [FOLDER_RESOURCES, FOLDER_REVIEW_DOCS, FOLDER_MARKUP_EXPORTS];
    for (const name of needed) {
      logStep(`Creating folder "${name}"...`, 'info');
      const id = await createFolder(projectId, name, 0, accessToken);
      await sleep(1500);
      pocState.folderIds[name] = id;
      logStep(`Folder "${name}" created (id=${id})`, 'success');
    }

    // 4 — Upload custom-columns.xml to resources folder
    logStep('Uploading custom-columns.xml to resources folder...', 'info');
    const xmlBuffer = fs.readFileSync(CUSTOM_COLUMNS_XML_PATH);
    const result    = await uploadFileToProject(
      xmlBuffer, 'custom-columns.xml', projectId, accessToken,
      pocState.folderIds[FOLDER_RESOURCES]
    );
    pocState.customColumnsFileId = result.projectFileId;
    logStep(`custom-columns.xml uploaded (fileId=${pocState.customColumnsFileId})`, 'success');

    pocState.projectSetupDone = true;
    logStep('Project setup complete', 'success');
    res.json({ success: true, projectId, folderIds: pocState.folderIds, state: pocState });

  } catch (err) {
    pocState.status = 'error';
    logStep(err.message, 'error');
    res.status(500).json({ error: err.message });
  }
});

// -----------------------------------------------------------------------------
// STEP 0b — Upload PDF(s) from UI
// Optionally injects state model before upload.
// -----------------------------------------------------------------------------
app.post('/poc/upload-to-project', upload.array('files'), async (req, res) => {
  try {
    if (!req.files || req.files.length === 0) throw new Error('No files received');
    if (!pocState.projectId) throw new Error('No project — run setup-project first');

    pocState.status = 'uploading';
    const injectModel = req.body.injectStateModel !== 'false'; // default true
    logStep(`Received ${req.files.length} file(s) for upload (injectStateModel=${injectModel})`, 'info');

    const accessToken    = await tokenManager.getValidAccessToken();
    const reviewFolderId = pocState.folderIds[FOLDER_REVIEW_DOCS] || null;
    const uploaded       = [];

    for (const file of req.files) {
      let buffer = file.buffer;

      if (injectModel) {
        logStep(`Injecting 5-step QC Review state model into "${file.originalname}"...`, 'info');
        buffer = await injectStateModel(buffer);
        logStep(`State model injected (${buffer.length} bytes)`, 'success');
      }

      const result = await uploadFileToProject(buffer, file.originalname, pocState.projectId, accessToken, reviewFolderId);
      uploaded.push(result);
      pocState.projectFiles.push(result);
    }

    if (uploaded.length > 0 && demoStub.documentId === 'DOC-001')
      demoStub.documentId = uploaded[0].name.replace(/\.[^.]+$/, '');

    logStep(`${uploaded.length} file(s) uploaded to project`, 'success');
    res.json({ success: true, uploaded, state: pocState });

  } catch (err) {
    pocState.status = 'error';
    logStep(err.message, 'error');
    res.status(500).json({ error: err.message });
  }
});

// -----------------------------------------------------------------------------
// STEP 0c — Apply Custom Columns (optional)
// -----------------------------------------------------------------------------
app.post('/poc/apply-custom-columns', async (req, res) => {
  try {
    if (!pocState.customColumnsFileId) throw new Error('custom-columns.xml not uploaded');
    if (!pocState.projectFiles.length) throw new Error('No project files');

    const accessToken = await tokenManager.getValidAccessToken();
    const results     = [];

    for (const pf of pocState.projectFiles) {
      logStep(`Submitting importcustomcolumns job for "${pf.name}"...`, 'info');
      const jobResp = await fetch(
        `${API_V1}/projects/${pocState.projectId}/files/${pf.projectFileId}/jobs/importcustomcolumns`,
        {
          method:  'POST',
          headers: authHeaders(accessToken),
          body:    JSON.stringify({
            CurrentPassword:     '',
            CustomColumnsFileID: parseInt(pocState.customColumnsFileId, 10),
            OutputFileName:      pf.name,
            OutputPath:          FOLDER_REVIEW_DOCS,
            Priority:            0
          })
        }
      );
      if (!jobResp.ok) {
        const err = await jobResp.text();
        logStep(`importcustomcolumns failed: ${jobResp.status} - ${err}`, 'warn');
        results.push({ name: pf.name, success: false, error: err });
        continue;
      }
      const { Id: jobId } = await jobResp.json();
      await pollJob(`${API_V1}/jobs/${jobId}`, authHeaders(accessToken));
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
  pocState.log    = [];
  logStep(`Workflow event — project: ${demoStub.projectName}`, 'info');
  logStep(`Files: ${pocState.projectFiles.map(f => f.name).join(', ') || '(none)'}`, 'info');
  logStep(`Reviewers: ${demoStub.reviewers.map(r => r.email).join(', ')}`, 'info');
  res.json({ success: true, state: pocState });
});

// -----------------------------------------------------------------------------
// STEP 2 — Create Session
// -----------------------------------------------------------------------------
app.post('/poc/create-session', async (req, res) => {
  try {
    if (!pocState.projectId) throw new Error('No project — run setup-project first');
    pocState.status = 'creating';

    const accessToken = await tokenManager.getValidAccessToken();
    const sessionName = `${demoStub.projectName}_${new Date().toISOString().slice(0, 10)}`;

    const resp = await fetch(`${API_V1}/sessions`, {
      method:  'POST',
      headers: authHeaders(accessToken),
      body:    JSON.stringify({
        Name:           sessionName,
        Notification:   true,
        Restricted:     true,
        SessionEndDate: demoStub.sessionEndDate,
        DefaultPermissions: [
          { Type: 'Markup',       Allow: 'Allow' },
          { Type: 'SaveCopy',     Allow: 'Allow' },
          { Type: 'PrintCopy',    Allow: 'Allow' },
          { Type: 'MarkupAlert',  Allow: 'Allow' },
          { Type: 'AddDocuments', Allow: 'Deny'  }
        ]
      })
    });
    if (!resp.ok) throw new Error(`Session creation failed: ${resp.status} - ${await resp.text()}`);

    const data         = await resp.json();
    pocState.sessionId = data.Id;
    pocState.createdAt = new Date().toISOString();
    logStep(`Session created: ID=${pocState.sessionId}`, 'success');
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
    if (!pocState.sessionId) throw new Error('No active session');

    if (isLocalhost(WEBHOOK_CALLBACK_URL)) {
      logStep('Webhook skipped (localhost)', 'warn');
      return res.json({ success: true, skipped: true, state: pocState });
    }

    const accessToken = await tokenManager.getValidAccessToken();
    const resp = await fetch(`${API_V2}/subscriptions`, {
      method:  'POST',
      headers: authHeaders(accessToken),
      body:    JSON.stringify({ sourceType: 'session', resourceId: pocState.sessionId, callbackURI: WEBHOOK_CALLBACK_URL })
    });
    if (!resp.ok) throw new Error(`Webhook registration failed: ${resp.status} - ${await resp.text()}`);

    const data              = await resp.json();
    pocState.subscriptionId = data.subscriptionId;
    logStep(`Webhook registered: subscriptionId=${pocState.subscriptionId}`, 'success');
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
    if (!pocState.sessionId)          throw new Error('No active session');
    if (!pocState.projectFiles.length) throw new Error('No project files');

    pocState.status = 'checking-out';
    const accessToken = await tokenManager.getValidAccessToken();
    const checked     = [];

    for (const pf of pocState.projectFiles) {
      logStep(`Checking out "${pf.name}" (projectFileId=${pf.projectFileId})...`, 'info');

      const resp = await fetch(
        `${API_V1}/projects/${pocState.projectId}/files/${pf.projectFileId}/checkout-to-session`,
        { method: 'POST', headers: authHeaders(accessToken), body: JSON.stringify({ SessionId: pocState.sessionId }) }
      );

      if (!resp.ok) {
        const err = await resp.text();
        if (resp.status === 409) {
          logStep(`409 — releasing checkout and retrying...`, 'warn');
          const rel = await fetch(
            `${API_V1}/projects/${pocState.projectId}/files/${pf.projectFileId}/checkout`,
            { method: 'DELETE', headers: authHeaders(accessToken) }
          );
          if (rel.ok) {
            const retry = await fetch(
              `${API_V1}/projects/${pocState.projectId}/files/${pf.projectFileId}/checkout-to-session`,
              { method: 'POST', headers: authHeaders(accessToken), body: JSON.stringify({ SessionId: pocState.sessionId }) }
            );
            if (!retry.ok) { logStep(`Retry failed: ${retry.status}`, 'warn'); continue; }
          } else { logStep(`Release failed`, 'warn'); continue; }
        } else { logStep(`Checkout failed: ${resp.status} - ${err}`, 'warn'); continue; }
      }

      await sleep(1000);

      const sfResp = await fetch(
        `${API_V1}/sessions/${pocState.sessionId}/files?includeDeleted=false`,
        { headers: authHeaders(accessToken) }
      );
      if (!sfResp.ok) { logStep(`Could not list session files`, 'warn'); continue; }

      const sfData  = await sfResp.json();
      const sfList  = sfData.SessionFiles || sfData.Files || [];
      const match   = sfList.find(f => f.ProjectFileId === pf.projectFileId || f.Name === pf.name);
      if (!match) { logStep(`Session file entry not found for "${pf.name}"`, 'warn'); continue; }

      const entry = { sessionFileId: match.Id, projectFileId: pf.projectFileId, name: pf.name };
      pocState.sessionFileIds.push(entry);
      checked.push(entry);
      logStep(`"${pf.name}" checked out (sessionFileId=${match.Id})`, 'success');
    }

    res.json({ success: true, checked, state: pocState });
  } catch (err) {
    pocState.status = 'error';
    logStep(err.message, 'error');
    res.status(500).json({ error: err.message });
  }
});

// -----------------------------------------------------------------------------
// STEP 5 — Invite Reviewers + fetch user IDs + set per-user permissions
// (B&McD steps 4, 6, 7, 8 combined)
// -----------------------------------------------------------------------------
app.post('/poc/invite-reviewers', async (req, res) => {
  try {
    if (!pocState.sessionId) throw new Error('No active session');
    pocState.status = 'inviting';

    const accessToken = await tokenManager.getValidAccessToken();

    // 5a — Invite to session
    const sessionResults = [];
    for (const reviewer of demoStub.reviewers) {
      const endpoint = reviewer.hasStudioAccount
        ? `${API_V1}/sessions/${pocState.sessionId}/users`
        : `${API_V1}/sessions/${pocState.sessionId}/invite`;

      logStep(`Inviting ${reviewer.email} (session)...`, 'info');
      const resp = await fetch(endpoint, {
        method:  'POST',
        headers: authHeaders(accessToken),
        body:    JSON.stringify({ Email: reviewer.email, Message: `Review invitation: ${demoStub.projectName}` })
      });
      if (!resp.ok) {
        const err = await resp.text();
        logStep(`Session invite failed for ${reviewer.email}: ${resp.status}`, 'warn');
        sessionResults.push({ email: reviewer.email, success: false, error: err });
      } else {
        logStep(`Session invited: ${reviewer.email}`, 'success');
        sessionResults.push({ email: reviewer.email, success: true });
      }
    }

    // 5b — Invite to project (B&McD step 4)
    for (const reviewer of demoStub.reviewers) {
      await inviteProjectUser(pocState.projectId, reviewer.email, accessToken);
      logStep(`Project invited: ${reviewer.email}`, 'info');
    }

    await sleep(2000); // allow user records to propagate

    // 5c — GET project users to retrieve numeric user IDs (B&McD step 6)
    const projectUsers = await getProjectUsers(pocState.projectId, accessToken);
    pocState.projectUserIds = projectUsers;
    logStep(`Fetched ${projectUsers.length} project user(s)`, 'success');

    // 5d — Set folder permissions per user (B&McD step 7)
    const nonOwnerUsers = projectUsers.filter(u => !u.IsProjectOwner);
    for (const user of nonOwnerUsers) {
      // review-documents: ReadWrite, markup-exports: Read, resources: Read
      const folderPerms = [
        { folder: FOLDER_REVIEW_DOCS,    perm: 'ReadWrite' },
        { folder: FOLDER_MARKUP_EXPORTS, perm: 'Read'      },
        { folder: FOLDER_RESOURCES,      perm: 'Read'      }
      ];
      for (const fp of folderPerms) {
        if (pocState.folderIds[fp.folder]) {
          await setFolderPermission(pocState.projectId, pocState.folderIds[fp.folder], user.Id, fp.perm, accessToken);
          logStep(`Folder perm: ${user.Email} → ${fp.folder}=${fp.perm}`, 'info');
        }
      }

      // 5e — Per-user permissions (B&McD step 8)
      await setUserPermission(pocState.projectId, user.Id, 'CreateSessions', 'Deny',  accessToken);
      await setUserPermission(pocState.projectId, user.Id, 'ManagePermissions', 'Deny', accessToken);
      logStep(`User perms set for ${user.Email}`, 'info');
    }

    pocState.status = 'active';
    logStep(`Session active — ID: ${pocState.sessionId}`, 'success');
    res.json({ success: true, sessionResults, projectUsers, state: pocState });
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
    const results     = await performCheckin(accessToken);
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
    const results     = await performExportMarkups(accessToken);
    res.json({ success: true, results, markupExports: pocState.markupExports, state: pocState });
  } catch (err) {
    pocState.status = 'error';
    logStep(err.message, 'error');
    res.status(500).json({ error: err.message });
  }
});

// -----------------------------------------------------------------------------
// STEP 9 — Run Markup List Job (XML-backed)
// -----------------------------------------------------------------------------
app.post('/poc/run-markuplist-job', async (req, res) => {
  try {
    const accessToken = await tokenManager.getValidAccessToken();
    if (!pocState.markupExports.length) {
      logStep('No exports yet — running export-markups first...', 'info');
      await performExportMarkups(accessToken);
    }
    const markups = await performMarkupExtractionFromXml(accessToken);
    res.json({ success: true, count: markups.length, markups, extractionMode: 'exportmarkups-xml', state: pocState });
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
    if (!pocState.sessionId)               throw new Error('No active session');
    if (!pocState.sessionFileIds.length)   throw new Error('No session files');

    logStep('Starting downstream processing...', 'info');
    const accessToken = await tokenManager.getValidAccessToken();

    const checkinResults = await performCheckin(accessToken);
    logStep('Post-checkin settle: 5s...', 'info');
    await sleep(5000);

    const exportResults = await performExportMarkups(accessToken);
    logStep('Post-export settle: 5s...', 'info');
    await sleep(5000);

    const markups = await performMarkupExtractionFromXml(accessToken);
    logStep('Downstream processing complete', 'success');

    res.json({
      success: true,
      checkinResults, exportResults,
      count: markups.length, markups,
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
// STEP 10 — Finalize Session
// -----------------------------------------------------------------------------
app.post('/poc/finalize', async (req, res) => {
  try {
    if (!pocState.sessionId) throw new Error('No active session');
    pocState.status = 'finalizing';

    const accessToken = await tokenManager.getValidAccessToken();
    const resp = await fetch(`${API_V1}/sessions/${pocState.sessionId}`, {
      method:  'PUT',
      headers: authHeaders(accessToken),
      body:    JSON.stringify({ Name: `${demoStub.projectName}_Final`, Restricted: true, SessionEndDate: demoStub.sessionEndDate })
    });
    if (!resp.ok) throw new Error(`Finalize failed: ${resp.status} - ${await resp.text()}`);

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
    if (!pocState.sessionId || !pocState.sessionFileIds.length)
      throw new Error('No active session or no files');

    pocState.status = 'snapshotting';
    const accessToken = await tokenManager.getValidAccessToken();
    const downloads   = [];

    for (const sf of pocState.sessionFileIds) {
      const snapResp = await fetch(
        `${API_V1}/sessions/${pocState.sessionId}/files/${sf.sessionFileId}/snapshot`,
        { method: 'POST', headers: authHeaders(accessToken) }
      );
      if (!snapResp.ok) { logStep(`Snapshot request failed: ${snapResp.status}`, 'warn'); continue; }

      let dlUrl = null;
      for (let i = 0; i < 20; i++) {
        await sleep(5000);
        const pollToken = await tokenManager.getValidAccessToken();
        const p = await fetch(
          `${API_V1}/sessions/${pocState.sessionId}/files/${sf.sessionFileId}/snapshot`,
          { headers: authHeaders(pollToken) }
        );
        if (!p.ok) continue;
        const d = await p.json();
        logStep(`Snapshot poll ${i + 1}: ${d.Status}`, 'info');
        if (d.Status === 'Complete') { dlUrl = d.DownloadUrl; break; }
        if (d.Status === 'Error')    throw new Error(`Snapshot error: ${d.Message}`);
      }

      if (!dlUrl) { logStep('Snapshot timed out', 'warn'); continue; }

      const dlResp = await fetch(dlUrl);
      if (!dlResp.ok) throw new Error(`Download failed: ${dlResp.status}`);
      const pdfBuf = await dlResp.buffer();

      const publicDir = path.join(__dirname, 'public');
      if (!fs.existsSync(publicDir)) fs.mkdirSync(publicDir, { recursive: true });

      const outFile = `${demoStub.projectName.replace(/\s+/g,'_')}_${sf.name.replace(/\.[^.]+$/, '')}_Reviewed.pdf`;
      fs.writeFileSync(path.join(publicDir, outFile), pdfBuf);
      downloads.push({ name: outFile, path: `/${outFile}`, size: pdfBuf.length });
      logStep(`PDF saved: ${outFile}`, 'success');
    }

    pocState.status = 'complete';
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
    if (!pocState.sessionId) throw new Error('No active session');
    const accessToken = await tokenManager.getValidAccessToken();

    if (pocState.subscriptionId) {
      const r = await fetch(`${API_V2}/subscriptions/${pocState.subscriptionId}`,
        { method: 'DELETE', headers: authHeaders(accessToken) });
      logStep(r.ok ? 'Webhook deleted' : `Sub delete: ${r.status}`, r.ok ? 'success' : 'warn');
    }

    const r = await fetch(`${API_V1}/sessions/${pocState.sessionId}`,
      { method: 'DELETE', headers: authHeaders(accessToken) });
    logStep(r.ok ? 'Session deleted' : `Session delete: ${r.status}`, r.ok ? 'success' : 'warn');

    pocState.sessionId      = null;
    pocState.subscriptionId = null;
    res.json({ success: true, state: pocState });
  } catch (err) {
    logStep(err.message, 'error');
    res.status(500).json({ error: err.message });
  }
});

// -----------------------------------------------------------------------------
// STANDALONE: inject state model into a single PDF (for testing)
// POST /poc/inject-state-model  multipart, field "file"
// -----------------------------------------------------------------------------
app.post('/poc/inject-state-model', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'No file uploaded' });

    const modified   = await injectStateModel(req.file.buffer);
    const outName    = req.file.originalname.replace(/\.pdf$/i, '') + '_state_injected.pdf';

    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `attachment; filename="${outName}"`);
    res.setHeader('Content-Length', modified.length);
    res.send(modified);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// -----------------------------------------------------------------------------
// WEBHOOK
// -----------------------------------------------------------------------------
app.post('/webhook/studio-events', (req, res) => {
  const p = req.body || {};
  logStep(`Webhook: ${p.ResourceType || 'unknown'} / ${p.EventType || 'unknown'}`, 'webhook');
  pocState.webhookEvents.push({ ...p, receivedAt: new Date().toISOString() });
  res.sendStatus(200);
});

// -----------------------------------------------------------------------------
// DATA ENDPOINTS
// -----------------------------------------------------------------------------
app.get('/api/project-markups', (req, res) => {
  if (!pocState.markups.length)
    return res.status(404).json({ error: 'No markup data.' });
  res.json(pocState.markups.map(m => ({
    MarkupId:           m.Id,
    Author:             m.Author,
    Type:               m.Type,
    Subject:            m.Subject,
    Comment:            m.Comment,
    Status:             m.Status,
    StatusHistory:      m.StatusHistory || [],
    Replies:            m.Replies || [],
    Layer:              m.Layer,
    Page:               m.Page,
    DateCreated:        m.DateCreated,
    DateModified:       m.DateModified,
    Color:              m.Color,
    Checked:            m.Checked,
    Locked:             m.Locked,
    ExtendedProperties: m.ExtendedProperties || {},
    SourceFile:         m._sourceFile
  })));
});

// -----------------------------------------------------------------------------
// START
// -----------------------------------------------------------------------------
app.listen(PORT, () => {
  console.log(`\nBluebeam Studio PoC  →  http://localhost:${PORT}`);
  console.log(`Dynamic project creation — new project per run`);
  if (isLocalhost(WEBHOOK_CALLBACK_URL)) console.log('⚠  Webhook will be skipped (localhost)');
  console.log(`\nState Model: "${QC_STATE_MODEL.cUIName}" (${QC_STATE_MODEL.states.length} states)`);
  console.log(`Custom models removed before injection: ${STATE_MODELS_TO_REMOVE.join(', ')}\n`);
});
