/**
 * Bluebeam Studio API — Document Roundtrip PoC
 * Proof-of-concept reference implementation. Not for production use.
 *
 * CHANGES IN THIS VERSION:
 *   - /poc/upload-to-project: state model injection is now ON by default.
 *     Pass injectStateModel=false in the multipart body to skip.
 *   - /poc/configure: accepts sessionEndDate in body; no longer overwrites it on every call
 *   - /poc/hydrate:   NEW — rehydrates in-memory pocState from DB after a server restart
 *   - /poc/auth/refresh: NEW — accepts a Bluebeam OAuth refresh token and rotates the
 *                        tokenManager's access token without a server restart
 *
 * XML PARSING FIXES (this version):
 *   - normalizeMarkupRecord: <Contents> is the current state model value, not Comment.
 *     Comment now maps only from comment/comments/note/message/reply (not contents).
 *     Status maps from state/contents (state value confirmed from BAX format analysis).
 *   - looksLikeMarkupRecord: records with a 'parent' key are StatusHistory audit entries,
 *     not real markups — excluded from candidate set.
 *   - extractMarkupCandidates: skips 'statushistory' subtrees entirely to prevent
 *     audit child records (Subject: "Set to X") from inflating the markup count.
 *
 * Full roundtrip flow:
 *   0a. /poc/setup-project          — Create folders + upload custom-columns.xml (once)
 *   0b. /poc/upload-to-project      — Inject state model (default ON) + upload PDF(s) → project + DB
 *   0c. /poc/apply-custom-columns   — Apply custom-columns.xml to each file (optional, not in UI)
 *    H. /poc/hydrate                — Restore pocState from DB after restart (use before downstream steps)
 *   1.  /poc/trigger                — Simulate source-system workflow event
 *   2.  /poc/create-session         — Create Studio Session + DB write + schedule poller
 *   3.  /poc/register-webhook       — Subscribe to session events
 *   4.  /poc/checkout-to-session    — Check project files out into session + update DB files table
 *   5.  /poc/invite-reviewers       — Invite reviewers
 *   6.  (Review in Bluebeam Revu — no API step)
 *   7.  /poc/checkin                — Check session files back into project
 *   8.  /poc/export-markups         — Run exportmarkups job → XML in project
 *   9.  /poc/run-markuplist-job     — Compatibility route; extracts structured data from exported XML
 *   9b. /poc/downstream-process     — Combined: checkin + export + XML parse + DB snapshot
 *   10. /poc/finalize               — Finalize session + update DB status
 *   11. /poc/snapshot               — Snapshot + download marked-up PDF
 *   12. /poc/cleanup                — Delete webhook + session + clear poller
 *
 * DB endpoints:
 *   GET  /poc/projects                       — List all Atkins projects in DB
 *   GET  /poc/projects/:atkinsId/snapshot    — Full project snapshot for Tab 2 dashboard
 *
 * Auth:
 *   POST /poc/auth/refresh                   — Rotate access token via refresh token
 *
 * Utility:
 *   GET  /health                             — Server + DB status
 *   GET  /poc/state                          — Current in-memory pocState + demoStub
 *   POST /poc/hydrate                        — Restore pocState from DB after restart
 *   POST /poc/reset                          — Clear in-memory state
 *   POST /poc/configure                      — Update demoStub + upsert project in DB
 */

require('dotenv').config();
const express  = require('express');
const cors     = require('cors');
const fs       = require('fs');
const path     = require('path');
const multer   = require('multer');
const { parseStringPromise } = require('xml2js');
const { PDFDocument, PDFName, PDFDict, PDFArray, PDFString } = require('pdf-lib');
const TokenManager = require('./tokenManager');
const db           = require('./db');

const app    = express();
const PORT   = process.env.PORT || 3000;
const upload = multer({ storage: multer.memoryStorage() });

// =============================================================================
// API CONFIGURATION
// =============================================================================
const API_V1 = 'https://api.bluebeam.com/publicapi/v1';
const API_V2 = 'https://api.bluebeam.com/publicapi/v2';
const CLIENT_ID = process.env.BB_CLIENT_ID;

const WEBHOOK_CALLBACK_URL =
  process.env.WEBHOOK_CALLBACK_URL ||
  `http://localhost:${PORT}/webhook/studio-events`;

// Hardcoded Studio Project ID for this PoC
const POC_PROJECT_ID = '712-566-288';

// Project folder names
const FOLDER_RESOURCES     = 'resources';
const FOLDER_REVIEW_DOCS   = 'review-documents';
const FOLDER_MARKUP_EXPORTS = 'markup-exports';

// Path to custom columns XML bundled with this repo
const CUSTOM_COLUMNS_XML_PATH = path.join(__dirname, 'resources', 'custom-columns.xml');

// =============================================================================
// STATE MODEL — 5-step QC Review
// =============================================================================

const STATE_MODELS_TO_REMOVE = [
  '5_step_QC_Review',
  'Incorrect_Review_Model',
  'Bad_Model',
  'Old_Review'
];

const QC_STATE_MODEL = {
  cName:    '5_step_QC_Review',
  cUIName:  '5-step QC Review',
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

// =============================================================================
// PDF STATE MODEL INJECTION (pdf-lib — no Bluebeam API required)
// =============================================================================
async function injectStateModel(pdfBuffer) {
  const pdfDoc = await PDFDocument.load(pdfBuffer, { ignoreEncryption: true });
  const ctx    = pdfDoc.context;
  const cat    = pdfDoc.catalog;

  const removeCalls = STATE_MODELS_TO_REMOVE
    .map(name => `try{Collab.removeStateModel("${name}");}catch(e){}`)
    .join('\n');

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

  cat.set(
    PDFName.of('OpenAction'),
    ctx.register(ctx.obj({ S: PDFName.of('JavaScript'), JS: jsRef }))
  );

  return Buffer.from(await pdfDoc.save());
}

// =============================================================================
// DEMO STUB
// =============================================================================
let demoStub = {
  atkinsProjectId: process.env.DEMO_ATKINS_PROJECT_ID || '',
  documentId:      process.env.DEMO_DOCUMENT_ID        || 'DOC-001',
  description:     process.env.DEMO_DESCRIPTION         || 'Design review — coordination update',
  projectName:     '',
  region:          '',
  reviewType:      '',
  qaCategory:      '',
  discipline:      '',
  pollingInterval: 0,
  reviewers:       [{ email: 'dmolz@bluebeam.com', hasStudioAccount: true }],
  sessionEndDate:  new Date(Date.now() + 7 * 24 * 60 * 60 * 1000).toISOString()
};

// =============================================================================
// IN-MEMORY STATE
// =============================================================================
let pocState = {
  sessionId:           null,
  subscriptionId:      null,
  projectSetupDone:    false,
  folderIds:           {},
  customColumnsFileId: null,
  projectFiles:        [],
  sessionFileIds:      [],
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

const fetch        = (...args) => import('node-fetch').then(({ default: f }) => f(...args));
const tokenManager = new TokenManager();

// =============================================================================
// HELPERS
// =============================================================================
function resetPocState() {
  pocState = {
    sessionId:           null,
    subscriptionId:      null,
    projectSetupDone:    false,
    folderIds:           {},
    customColumnsFileId: null,
    projectFiles:        [],
    sessionFileIds:      [],
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
    client_id:      CLIENT_ID,
    'Content-Type': 'application/json',
    Accept:         'application/json',
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
    const res  = await fetch(url, { headers });
    const data = await res.json();
    const status = data.Status ?? data.JobStatus;
    const msg    = data.StatusMessage ?? data.JobStatusMessage ?? '';
    logStep(`Job poll ${attempt}/${maxAttempts}: status=${status} ${msg}`.trim(), 'info');
    if (status === 200) return data;
    if (!inProgress.has(status)) throw new Error(`Job failed (status=${status}): ${msg}`);
  }
  throw new Error(`Job did not complete after ${maxAttempts} attempts`);
}

async function listProjectFolders(accessToken) {
  const resp = await fetch(`${API_V1}/projects/${POC_PROJECT_ID}/folders`, {
    headers: authHeaders(accessToken)
  });
  if (!resp.ok) throw new Error(`Failed to list folders: ${resp.status} - ${await resp.text()}`);
  return (await resp.json()).ProjectFolders || [];
}

async function createFolder(name, accessToken) {
  const resp = await fetch(`${API_V1}/projects/${POC_PROJECT_ID}/folders`, {
    method:  'POST',
    headers: authHeaders(accessToken),
    body:    JSON.stringify({ Name: name })
  });
  if (!resp.ok) throw new Error(`Failed to create folder "${name}": ${resp.status} - ${await resp.text()}`);
  return (await resp.json()).Id;
}

async function uploadFileToProject(fileBuffer, fileName, accessToken, folderId = null) {
  logStep(`Uploading "${fileName}" to project ${POC_PROJECT_ID}${folderId ? ` (folderId=${folderId})` : ''}...`, 'info');

  const metaBody = { Name: fileName, Size: fileBuffer.length, CRC: '0' };
  if (folderId) metaBody.ParentFolderId = folderId;

  const metaResp = await fetch(`${API_V1}/projects/${POC_PROJECT_ID}/files`, {
    method:  'POST',
    headers: authHeaders(accessToken),
    body:    JSON.stringify(metaBody)
  });
  if (!metaResp.ok) {
    throw new Error(`Metadata block failed for "${fileName}": ${metaResp.status} - ${await metaResp.text()}`);
  }

  const meta              = await metaResp.json();
  const projectFileId     = meta.Id;
  const uploadUrl         = meta.UploadUrl;
  const uploadContentType = meta.UploadContentType || 'application/pdf';

  logStep(`Metadata block created: projectFileId=${projectFileId}`, 'success');

  const s3Resp = await fetch(uploadUrl, {
    method:  'PUT',
    headers: { 'Content-Type': uploadContentType, 'x-amz-server-side-encryption': 'AES256' },
    body:    fileBuffer
  });
  if (!s3Resp.ok) throw new Error(`S3 upload failed for "${fileName}": ${s3Resp.status}`);
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
  if (!resp.ok) throw new Error(`Failed to list project files: ${resp.status} - ${await resp.text()}`);
  return (await resp.json()).ProjectFiles || [];
}

async function getProjectFileByPath(accessToken, filePath) {
  const url  = `${API_V1}/projects/${POC_PROJECT_ID}/files/by-path?path=${encodeURIComponent(filePath)}`;
  const resp = await fetch(url, { headers: authHeaders(accessToken) });
  if (!resp.ok) throw new Error(`Failed to get file by path: ${resp.status} - ${await resp.text()}`);
  return resp.json();
}

async function waitForProjectFileSettlement(accessToken, projectFileId, fileName, options = {}) {
  const { attempts = 6, intervalMs = 5000, finalExtraDelayMs = 5000 } = options;
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
      if (typeof match.Version !== 'undefined')        observed.push(`version=${match.Version}`);
      if (typeof match.IsCheckedOut !== 'undefined')   observed.push(`isCheckedOut=${match.IsCheckedOut}`);
      if (typeof match.CheckedOut !== 'undefined')     observed.push(`checkedOut=${match.CheckedOut}`);
      if (typeof match.ParentFolderId !== 'undefined') observed.push(`parentFolderId=${match.ParentFolderId}`);

      logStep(`Project file re-query ${attempt}/${attempts}: found "${fileName}" (${observed.join(', ')})`, 'info');

      const checkedOutFlags      = [match.IsCheckedOut, match.CheckedOut].filter(v => typeof v !== 'undefined');
      const explicitlyCheckedOut = checkedOutFlags.some(v => v === true || String(v).toLowerCase() === 'true');

      if (!explicitlyCheckedOut) {
        if (finalExtraDelayMs > 0) {
          logStep(`Project file appears settled. Waiting an extra ${Math.round(finalExtraDelayMs / 1000)}s...`, 'info');
          await sleep(finalExtraDelayMs);
        }
        return lastSeen;
      }
    } else {
      logStep(`Project file re-query ${attempt}/${attempts}: projectFileId=${projectFileId} not found yet`, 'warn');
    }

    if (attempt < attempts) await sleep(intervalMs);
  }

  logStep(`Settlement window elapsed for "${fileName}". Proceeding with best-known state.`, 'warn');
  if (finalExtraDelayMs > 0) await sleep(finalExtraDelayMs);
  return lastSeen;
}

async function waitForProjectFileReadyByPath(accessToken, filePath, expectedRevisionId = null) {
  for (let attempt = 1; attempt <= 8; attempt++) {
    const file        = await getProjectFileByPath(accessToken, filePath);
    const revisionOk  = expectedRevisionId == null || Number(file.RevisionID) >= Number(expectedRevisionId);
    const ready       = revisionOk && !file.InSession && !file.IsLocked;
    logStep(`by-path re-query ${attempt}/8: revision=${file.RevisionID}, inSession=${file.InSession}, isLocked=${file.IsLocked}`, 'info');
    if (ready) { await sleep(3000); return file; }
    await sleep(4000);
  }
  throw new Error(`Project file did not become ready by path: ${filePath}`);
}

// =============================================================================
// XML PARSING HELPERS
// =============================================================================
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

// Known state model value strings used by Bluebeam's built-in Review model
// and the custom 5-step QC model. Contents field values matching these are
// treated as state values, not free-text comments.
const KNOWN_STATE_PREFIXES = [
  'Step3', 'Step4', 'Step5',           // 5-step QC model
  'Accepted', 'Rejected', 'Cancelled', // built-in Review model
  'Assigned', 'Completed', 'Approved', // common status values
  'None'
];

function looksLikeStateValue(str) {
  if (!str || typeof str !== 'string') return false;
  const s = str.trim();
  if (!s) return false;
  return KNOWN_STATE_PREFIXES.some(prefix => s.startsWith(prefix)) || /^Step\d/.test(s);
}

function normalizeMarkupRecord(record, sourceFile) {
  const mapped = lowerKeyMap(record);

  // ── CONTENTS field handling ──────────────────────────────────────────────
  // In Bluebeam's BAX/XML export format:
  //   <Contents> = the CURRENT state model value (e.g. "Assigned", "Step3_Address_Future")
  //                when a state has been set on the markup.
  //   <Contents> = free-text comment when the reviewer typed a note directly.
  //
  // Strategy: if Contents looks like a state value, use it for Status.
  // Otherwise use it for Comment. Never use it for both.
  const rawContents = scalar(firstDefined(mapped, ['contents']));
  const contentsIsState = looksLikeStateValue(rawContents);

  const known = {
    Id:           scalar(firstDefined(mapped, ['id', 'markupid', 'markup_id'])),
    Author:       scalar(firstDefined(mapped, ['author', 'createdby', 'user', 'username'])),
    Type:         scalar(firstDefined(mapped, ['type', 'markuptype'])),
    Subject:      scalar(firstDefined(mapped, ['subject', 'label', 'title'])),
    // Comment: explicitly exclude 'contents' — handled separately below
    Comment:      scalar(firstDefined(mapped, ['comment', 'comments', 'note', 'message', 'reply'])),
    // Status: try state/status first, then fall back to Contents if it looks like a state value
    Status:       scalar(firstDefined(mapped, ['state', 'status'])) || (contentsIsState ? rawContents : ''),
    Layer:        scalar(firstDefined(mapped, ['layer'])),
    Page:         scalar(firstDefined(mapped, ['page', 'pagenumber', 'pageindex'])),
    DateCreated:  scalar(firstDefined(mapped, ['datecreated', 'creationdate', 'created', 'createddate'])),
    DateModified: scalar(firstDefined(mapped, ['datemodified', 'moddate', 'modified', 'modifieddate'])),
    Color:        scalar(firstDefined(mapped, ['color'])),
    Checked:      scalar(firstDefined(mapped, ['checked'])),
    Locked:       scalar(firstDefined(mapped, ['locked']))
  };

  // If Comment is still empty and Contents was not a state value, use it as comment
  if (!known.Comment && rawContents && !contentsIsState) {
    known.Comment = rawContents;
  }

  // Additional Status fallback: check <Statuses> sub-element (some export formats)
  if (!known.Status) {
    const statusesRaw = mapped.statuses || mapped.markupstatus || mapped.markupstatuses;
    if (statusesRaw) {
      const statusesMap = lowerKeyMap(typeof statusesRaw === 'object' ? statusesRaw : {});
      const sv = scalar(firstDefined(statusesMap, [
        'statustext', 'statetext', 'status', 'state', 'value', 'text', 'name', 'label', 'cname', 'cuiname'
      ]));
      if (sv) known.Status = sv;
    }
    if (!known.Status) {
      const sv2 = scalar(firstDefined(mapped, ['statustext','statetext','statusvalue','statevalue','statusname','statename']));
      if (sv2) known.Status = sv2;
    }
  }

  const skip = new Set([
    'id','markupid','markup_id','author','createdby','user','username',
    'type','markuptype','subject','label','title','comment','comments',
    'note','message','reply','contents','status','state','layer','page',
    'pagenumber','pageindex','datecreated','creationdate','created','createddate',
    'datemodified','moddate','modified','modifieddate','color','checked','locked','custom',
    'statuses','markupstatus','markupstatuses','statustext','statetext','statusvalue','statevalue',
    // Skip StatusHistory — it's audit trail, handled at extractMarkupCandidates level
    'statushistory','parent','typeinternal','raw','index'
  ]);

  const custom   = {};
  const rawCustom = mapped.custom;
  if (rawCustom && typeof rawCustom === 'object' && !Array.isArray(rawCustom)) {
    for (const [k, v] of Object.entries(rawCustom)) {
      const sv = scalar(v);
      if (typeof sv !== 'undefined' && sv !== null && String(sv).trim() !== '') {
        custom[k] = String(sv).trim();
      }
    }
  }

  const extended = {};
  for (const [k, v] of Object.entries(mapped)) {
    if (!skip.has(k)) {
      if (v && typeof v === 'object' && !Array.isArray(v)) {
        for (const [nk, nv] of Object.entries(v)) {
          const sv = scalar(nv);
          if (typeof sv !== 'undefined' && sv !== null && String(sv).trim() !== '') {
            extended[`${k}.${nk}`] = String(sv).trim();
          }
        }
      } else {
        const sv = scalar(v);
        if (typeof sv !== 'undefined' && sv !== null && String(sv).trim() !== '') {
          extended[k] = String(sv).trim();
        }
      }
    }
  }

  return { ...known, Custom: custom, ExtendedProperties: { ...custom, ...extended }, _sourceFile: sourceFile };
}

function looksLikeMarkupRecord(obj) {
  if (!obj || typeof obj !== 'object' || Array.isArray(obj)) return false;
  const keys = Object.keys(lowerKeyMap(obj));

  // StatusHistory/Status child records have a 'parent' key pointing to the
  // parent annotation ID. These are audit trail entries ("Set to Assigned"),
  // not real markup annotations — exclude them to prevent count inflation.
  if (keys.includes('parent')) return false;

  const hits = ['author','subject','comment','status','page','layer','type','markupid','id','markuptype']
    .filter(k => keys.includes(k)).length;
  return hits >= 2;
}

function extractMarkupCandidates(node, sourceFile, results = [], _key = '') {
  if (Array.isArray(node)) {
    for (const item of node) extractMarkupCandidates(item, sourceFile, results, _key);
    return results;
  }
  if (!node || typeof node !== 'object') return results;

  // Skip StatusHistory subtrees entirely — they contain audit child records
  // (Type: Text, Subject: "Set to X") that score as markup candidates but are
  // not real annotations. The state value is already captured in Contents on
  // the parent annotation.
  if (_key.toLowerCase() === 'statushistory') return results;

  if (looksLikeMarkupRecord(node)) {
    results.push(normalizeMarkupRecord(node, sourceFile));
    // Don't recurse into a matched record's children — avoids double-counting
    // nested structures like Custom columns being extracted as separate records.
    return results;
  }

  for (const [key, value] of Object.entries(node)) {
    extractMarkupCandidates(value, sourceFile, results, key);
  }
  return results;
}

async function downloadExportedMarkupXml(accessToken, exportFileName) {
  const xmlPath  = `/${FOLDER_MARKUP_EXPORTS}/${exportFileName}`;
  const fileMeta = await getProjectFileByPath(accessToken, xmlPath);
  if (!fileMeta.DownloadUrl) throw new Error(`DownloadUrl missing for exported XML: ${xmlPath}`);
  logStep(`Downloading exported XML from path=${xmlPath} (fileId=${fileMeta.Id})...`, 'info');
  const xmlResp = await fetch(fileMeta.DownloadUrl);
  if (!xmlResp.ok) throw new Error(`Failed to download exported XML "${exportFileName}": ${xmlResp.status}`);
  return xmlResp.text();
}

async function parseBluebeamExportXml(xmlText, sourceFile) {
  const parsed     = await parseStringPromise(xmlText, { explicitArray: false, mergeAttrs: true, trim: true });
  const candidates = extractMarkupCandidates(parsed, sourceFile);
  const seen       = new Set();
  const unique     = [];
  for (const item of candidates) {
    const key = [item.Id||'', item.Author||'', item.Subject||'', item.Comment||'', item.Page||'', item.DateCreated||''].join('|');
    if (!seen.has(key)) { seen.add(key); unique.push(item); }
  }
  logStep(`XML parse: ${candidates.length} candidates → ${unique.length} unique markups after dedup`, 'info');
  return unique;
}

// =============================================================================
// DOWNSTREAM HELPERS
// =============================================================================
async function performCheckin(accessToken) {
  if (!pocState.sessionId)              throw new Error('No active session');
  if (pocState.sessionFileIds.length === 0) throw new Error('No session files — run checkout-to-session first (or /poc/hydrate after restart)');

  pocState.status = 'checking-in';
  const results   = [];

  for (const sf of pocState.sessionFileIds) {
    logStep(`Checking in "${sf.name}" (sessionFileId=${sf.sessionFileId})...`, 'info');

    const resp = await fetch(
      `${API_V1}/sessions/${pocState.sessionId}/files/${sf.sessionFileId}/checkin`,
      { method: 'POST', headers: authHeaders(accessToken), body: JSON.stringify({ Comment: 'Session markup review complete' }) }
    );

    if (!resp.ok) {
      const err = await resp.text();
      logStep(`Check-in failed for "${sf.name}": ${resp.status} - ${err}`, 'warn');
      results.push({ name: sf.name, success: false, error: err });
      continue;
    }

    logStep(`"${sf.name}" checked in to project`, 'success');
    await waitForProjectFileSettlement(accessToken, sf.projectFileId, sf.name, {
      attempts: 6, intervalMs: 5000, finalExtraDelayMs: 5000
    });
    results.push({ name: sf.name, success: true });
  }

  logStep('Check-in complete — project files updated with session markups', 'success');
  return results;
}

async function performExportMarkups(accessToken) {
  if (pocState.sessionFileIds.length === 0) throw new Error('No session files — run checkout-to-session first (or /poc/hydrate after restart)');

  logStep('Exporting markups to XML...', 'info');
  const results = [];

  if (!pocState.folderIds[FOLDER_MARKUP_EXPORTS]) {
    logStep('markup-exports folder ID not set — re-querying folders...', 'info');
    const folders = await listProjectFolders(accessToken);
    const found   = folders.find(f => f.Name === FOLDER_MARKUP_EXPORTS);
    if (found) pocState.folderIds[FOLDER_MARKUP_EXPORTS] = found.Id;
    else throw new Error(`Folder "${FOLDER_MARKUP_EXPORTS}" not found — run setup-project first`);
  }

  for (const sf of pocState.sessionFileIds) {
    const exportFileName = `Markups-${sf.projectFileId}.xml`;
    logStep(`Submitting exportmarkups job for "${sf.name}" → ${exportFileName}...`, 'info');

    const jobResp = await fetch(
      `${API_V1}/projects/${POC_PROJECT_ID}/files/${sf.projectFileId}/jobs/exportmarkups`,
      {
        method:  'POST',
        headers: authHeaders(accessToken),
        body:    JSON.stringify({ OutputFileName: exportFileName, OutputPath: FOLDER_MARKUP_EXPORTS, Priority: 0 })
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

    await pollJob(`${API_V1}/jobs/${jobId}`, authHeaders(accessToken), 15, 15000);

    logStep(`Markup XML exported: ${exportFileName}`, 'success');
    logStep('Allowing exported markup artifacts to settle for 5s...', 'info');
    await sleep(5000);

    const existingIndex = pocState.markupExports.findIndex(m => m.exportFileName === exportFileName);
    const exportRecord  = { name: sf.name, exportFileName, projectPath: FOLDER_MARKUP_EXPORTS };
    if (existingIndex >= 0) pocState.markupExports[existingIndex] = exportRecord;
    else pocState.markupExports.push(exportRecord);

    results.push({ name: sf.name, success: true, exportFileName });
  }

  return results;
}

async function performMarkupExtractionFromXml(accessToken) {
  if (!pocState.sessionFileIds.length) throw new Error('No session files — run checkout-to-session first (or /poc/hydrate after restart)');

  pocState.status  = 'extracting-markups';
  pocState.markups = [];

  for (const sf of pocState.sessionFileIds) {
    const projectFile = pocState.projectFiles.find(pf => String(pf.projectFileId) === String(sf.projectFileId));

    if (projectFile) {
      const reviewPath = `/${FOLDER_REVIEW_DOCS}/${projectFile.name}`;
      try {
        const currentFile = await waitForProjectFileReadyByPath(accessToken, reviewPath);
        logStep(
          `Resolved project readiness: path=${reviewPath}, revision=${currentFile?.RevisionID}, inSession=${currentFile?.InSession}, isLocked=${currentFile?.IsLocked}`,
          'info'
        );
      } catch (err) {
        logStep(`Project file not fully ready by path (${reviewPath}) — proceeding with XML parse: ${err.message}`, 'warn');
      }
    }

    const exportFileName = `Markups-${sf.projectFileId}.xml`;
    logStep(`Downloading and parsing exported XML for "${sf.name}"...`, 'info');

    const xmlText    = await downloadExportedMarkupXml(accessToken, exportFileName);
    const fileMarkups = await parseBluebeamExportXml(xmlText, sf.name);

    pocState.markups.push(...fileMarkups);
    logStep(`"${sf.name}" — ${fileMarkups.length} markup(s) extracted from exported XML`, 'success');
  }

  pocState.status = 'active';
  logStep(`XML extraction complete — ${pocState.markups.length} total markup(s)`, 'success');
  return pocState.markups;
}

// =============================================================================
// POLLING SCHEDULER
// =============================================================================
const activePollers = new Map();

function scheduleSessionPoller(bluebeamSessionId, atkinsProjectId, intervalHours) {
  if (activePollers.has(bluebeamSessionId)) {
    clearInterval(activePollers.get(bluebeamSessionId));
    activePollers.delete(bluebeamSessionId);
    logStep(`Cleared existing poller for session ${bluebeamSessionId}`, 'info');
  }

  if (!intervalHours || intervalHours <= 0) return;

  const intervalMs = intervalHours * 60 * 60 * 1000;
  logStep(`Scheduling poller for session ${bluebeamSessionId} every ${intervalHours}h`, 'info');

  const timer = setInterval(async () => {
    logStep(`[POLLER] Auto-polling session ${bluebeamSessionId} (project ${atkinsProjectId})...`, 'info');
    try {
      const accessToken = await tokenManager.getValidAccessToken();

      if (pocState.sessionId !== bluebeamSessionId) {
        logStep(`[POLLER] Session ${bluebeamSessionId} not current in-memory session — saving last known snapshot`, 'warn');
        db.updateSessionPolled(bluebeamSessionId);
        return;
      }

      await performExportMarkups(accessToken);
      await sleep(3000);
      const markups = await performMarkupExtractionFromXml(accessToken);

      db.insertSnapshot({
        atkinsProjectId,
        bluebeamSessionId,
        markupCount:  markups.length,
        snapshotJson: JSON.stringify(markups)
      });

      db.updateSessionPolled(bluebeamSessionId);
      logStep(`[POLLER] Snapshot saved — ${markups.length} markup(s) for session ${bluebeamSessionId}`, 'success');
    } catch (err) {
      logStep(`[POLLER] Error polling session ${bluebeamSessionId}: ${err.message}`, 'error');
    }
  }, intervalMs);

  activePollers.set(bluebeamSessionId, timer);
}

function restorePollers() {
  try {
    const sessions = db.db.prepare('SELECT * FROM sessions WHERE polling_interval > 0').all();
    if (sessions.length) {
      logStep(`Restoring ${sessions.length} polling scheduler(s) from DB...`, 'info');
      for (const s of sessions) {
        scheduleSessionPoller(s.bluebeam_session_id, s.atkins_project_id, s.polling_interval);
      }
    }
  } catch (err) {
    logStep(`Failed to restore pollers: ${err.message}`, 'warn');
  }
}

// =============================================================================
// HEALTH CHECK
// =============================================================================
app.get('/health', (req, res) => {
  let dbStatus    = 'ok';
  let projectCount = 0;
  let sessionCount = 0;

  try {
    projectCount = db.db.prepare('SELECT COUNT(*) as c FROM projects').get().c;
    sessionCount = db.db.prepare('SELECT COUNT(*) as c FROM sessions').get().c;
  } catch (err) {
    dbStatus = `error: ${err.message}`;
  }

  res.json({
    status:    'healthy',
    projectId: POC_PROJECT_ID,
    db: {
      status:       dbStatus,
      projectCount,
      sessionCount,
      path:         process.env.DB_PATH || './poc.db'
    },
    pollers: activePollers.size,
    pocState: {
      status:           pocState.status,
      sessionId:        pocState.sessionId,
      projectFiles:     pocState.projectFiles.length,
      sessionFiles:     pocState.sessionFileIds.length,
      markups:          pocState.markups.length,
      projectSetupDone: pocState.projectSetupDone
    },
    config: {
      hasClientId:            Boolean(CLIENT_ID),
      webhookCallbackUrl:     WEBHOOK_CALLBACK_URL,
      webhookIsLocalhost:     isLocalhost(WEBHOOK_CALLBACK_URL),
      customColumnsXmlExists: fs.existsSync(CUSTOM_COLUMNS_XML_PATH),
      stateModel: {
        enabled:       true,
        name:          QC_STATE_MODEL.cUIName,
        states:        QC_STATE_MODEL.states.length,
        modelsRemoved: STATE_MODELS_TO_REMOVE
      }
    }
  });
});

// =============================================================================
// POC UTILITY ROUTES
// =============================================================================
app.get('/poc/state', (req, res) => {
  res.json({ ...pocState, stub: demoStub, projectId: POC_PROJECT_ID });
});

app.get('/poc/stub', (req, res) => res.json(demoStub));

// -----------------------------------------------------------------------------
// AUTH REFRESH
// -----------------------------------------------------------------------------
app.post('/poc/auth/refresh', async (req, res) => {
  try {
    const { refreshToken } = req.body || {};
    if (!refreshToken) {
      return res.status(400).json({ error: 'refreshToken is required in request body' });
    }

    if (typeof tokenManager.refreshWithToken !== 'function') {
      return res.status(501).json({
        error: 'tokenManager does not implement refreshWithToken(token). ' +
               'Add this method to tokenManager.js — see server.js comments.'
      });
    }

    await tokenManager.refreshWithToken(refreshToken);
    logStep('Access token rotated via UI-supplied refresh token', 'success');

    res.json({ success: true, message: 'Token refreshed successfully' });
  } catch (err) {
    logStep(`Token refresh failed: ${err.message}`, 'error');
    res.status(500).json({ error: err.message });
  }
});

// -----------------------------------------------------------------------------
// CONFIGURE
// -----------------------------------------------------------------------------
app.post('/poc/configure', (req, res) => {
  const {
    documentId,
    description,
    reviewerEmail,
    atkinsProjectId,
    projectName,
    region,
    reviewType,
    qaCategory,
    discipline,
    pollingInterval,
    sessionEndDate
  } = req.body || {};

  if (documentId)      demoStub.documentId      = documentId;
  if (description)     demoStub.description     = description;
  if (atkinsProjectId) demoStub.atkinsProjectId = atkinsProjectId;
  if (projectName)     demoStub.projectName     = projectName;
  if (region)          demoStub.region          = region;
  if (reviewType)      demoStub.reviewType      = reviewType;
  if (qaCategory)      demoStub.qaCategory      = qaCategory;
  if (discipline)      demoStub.discipline      = discipline;

  if (typeof pollingInterval !== 'undefined') {
    demoStub.pollingInterval = parseInt(pollingInterval) || 0;
  }

  if (sessionEndDate) {
    try {
      demoStub.sessionEndDate = new Date(sessionEndDate).toISOString();
      logStep(`sessionEndDate set to: ${demoStub.sessionEndDate}`, 'info');
    } catch (_) {
      logStep(`Invalid sessionEndDate value ignored: ${sessionEndDate}`, 'warn');
    }
  }

  if (reviewerEmail && reviewerEmail !== 'dmolz@bluebeam.com') {
    if (!demoStub.reviewers.some(r => r.email === reviewerEmail)) {
      demoStub.reviewers.push({ email: reviewerEmail, hasStudioAccount: false });
      logStep(`Added reviewer: ${reviewerEmail}`, 'info');
    }
  }

  if (demoStub.atkinsProjectId) {
    db.upsertProject({
      atkinsProjectId:   demoStub.atkinsProjectId,
      bluebeamProjectId: POC_PROJECT_ID,
      projectName:       demoStub.projectName  || demoStub.description || null,
      region:            demoStub.region       || null,
      reviewType:        demoStub.reviewType   || null,
      qaCategory:        demoStub.qaCategory   || null
    });
    logStep(`Project upserted in DB: ${demoStub.atkinsProjectId}`, 'info');
  }

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

// =============================================================================
// HYDRATE — restore pocState from DB after a server restart
// =============================================================================
app.post('/poc/hydrate', async (req, res) => {
  try {
    const { atkinsProjectId, bluebeamSessionId } = req.body || {};

    let session;

    if (bluebeamSessionId) {
      session = db.db.prepare(
        'SELECT * FROM sessions WHERE bluebeam_session_id = ?'
      ).get(bluebeamSessionId);
      if (!session) throw new Error(`Session not found in DB: ${bluebeamSessionId}`);
    } else if (atkinsProjectId) {
      session = db.db.prepare(
        `SELECT * FROM sessions
         WHERE atkins_project_id = ?
         ORDER BY created_at DESC
         LIMIT 1`
      ).get(atkinsProjectId);
      if (!session) throw new Error(`No sessions found in DB for project: ${atkinsProjectId}`);
    } else {
      session = db.db.prepare(
        'SELECT * FROM sessions ORDER BY created_at DESC LIMIT 1'
      ).get();
      if (!session) throw new Error('No sessions found in DB. Start a review first.');
    }

    logStep(`Hydrating from DB: session=${session.bluebeam_session_id}, project=${session.atkins_project_id}`, 'info');

    pocState.sessionId  = session.bluebeam_session_id;
    pocState.status     = session.status || 'active';
    pocState.createdAt  = session.created_at;

    const dbFiles = db.db.prepare(
      'SELECT * FROM files WHERE bluebeam_session_id = ?'
    ).all(session.bluebeam_session_id);

    const dbProjectFiles = db.db.prepare(
      `SELECT * FROM files
       WHERE (bluebeam_session_id = ? OR bluebeam_session_id IS NULL)
         AND atkins_project_id = ?`
    ).all(session.bluebeam_session_id, session.atkins_project_id);

    pocState.projectFiles = dbProjectFiles.map(f => ({
      projectFileId: f.bluebeam_project_file_id,
      name:          f.file_name,
      size:          f.file_size || 0
    }));

    pocState.sessionFileIds = dbFiles
      .filter(f => f.bluebeam_session_file_id)
      .map(f => ({
        sessionFileId: f.bluebeam_session_file_id,
        projectFileId: f.bluebeam_project_file_id,
        name:          f.file_name
      }));

    logStep(`Restored ${pocState.projectFiles.length} project file(s), ${pocState.sessionFileIds.length} session file(s)`, 'success');

    const project = db.getProject(session.atkins_project_id);
    if (project) {
      demoStub.atkinsProjectId = project.atkins_project_id;
      demoStub.projectName     = project.project_name || '';
      demoStub.region          = project.region       || '';
      demoStub.reviewType      = project.review_type  || '';
      demoStub.qaCategory      = project.qa_category  || '';
    } else {
      demoStub.atkinsProjectId = session.atkins_project_id;
    }

    demoStub.discipline      = session.discipline       || '';
    demoStub.pollingInterval = session.polling_interval || 0;
    demoStub.documentId      = session.review_name      || demoStub.documentId;

    let folderResolutionStatus = 'skipped';
    try {
      const accessToken = await tokenManager.getValidAccessToken();
      const folders     = await listProjectFolders(accessToken);
      folders.forEach(f => { pocState.folderIds[f.Name] = f.Id; });
      pocState.projectSetupDone = true;
      folderResolutionStatus    = 'ok';
      logStep(`Folder IDs resolved: ${Object.entries(pocState.folderIds).map(([n, id]) => `${n}=${id}`).join(', ')}`, 'success');
    } catch (folderErr) {
      logStep(`Could not re-resolve folder IDs (token may not be ready): ${folderErr.message}`, 'warn');
    }

    if (demoStub.pollingInterval > 0 && !activePollers.has(pocState.sessionId)) {
      scheduleSessionPoller(pocState.sessionId, demoStub.atkinsProjectId, demoStub.pollingInterval);
      logStep(`Poller re-scheduled: every ${demoStub.pollingInterval}h`, 'info');
    }

    logStep(`Hydration complete — ready for downstream processing`, 'success');

    res.json({
      success:               true,
      sessionId:             pocState.sessionId,
      atkinsProjectId:       demoStub.atkinsProjectId,
      projectFiles:          pocState.projectFiles.length,
      sessionFiles:          pocState.sessionFileIds.length,
      folderResolutionStatus,
      pollerScheduled:       activePollers.has(pocState.sessionId),
      folderIds:             pocState.folderIds,
      projectFilesDetail:    pocState.projectFiles,
      sessionFilesDetail:    pocState.sessionFileIds,
      state:                 pocState
    });
  } catch (err) {
    logStep(`Hydration failed: ${err.message}`, 'error');
    res.status(500).json({ error: err.message });
  }
});

// =============================================================================
// STEP 0a — Project Setup
// =============================================================================
app.post('/poc/setup-project', async (req, res) => {
  try {
    if (!fs.existsSync(CUSTOM_COLUMNS_XML_PATH)) {
      throw new Error(`custom-columns.xml not found at ${CUSTOM_COLUMNS_XML_PATH} — ensure resources/ folder is present`);
    }

    logStep('Running project setup...', 'info');
    const accessToken = await tokenManager.getValidAccessToken();
    const existing    = await listProjectFolders(accessToken);
    const folderMap   = {};
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

    const allFiles    = await listProjectFiles(accessToken);
    const existingXml = allFiles.find(f =>
      f.Name === 'custom-columns.xml' && f.ParentFolderId === pocState.folderIds[FOLDER_RESOURCES]
    );

    if (existingXml) {
      pocState.customColumnsFileId = existingXml.Id;
      logStep(`custom-columns.xml already in resources folder (fileId=${existingXml.Id})`, 'info');
    } else {
      logStep('Uploading custom-columns.xml to resources folder...', 'info');
      const xmlBuffer = fs.readFileSync(CUSTOM_COLUMNS_XML_PATH);
      const result    = await uploadFileToProject(xmlBuffer, 'custom-columns.xml', accessToken, pocState.folderIds[FOLDER_RESOURCES]);
      pocState.customColumnsFileId = result.projectFileId;
      logStep(`custom-columns.xml uploaded (fileId=${pocState.customColumnsFileId})`, 'success');
    }

    pocState.projectSetupDone = true;
    logStep('Project setup complete', 'success');

    res.json({
      success:             true,
      folderIds:           pocState.folderIds,
      customColumnsFileId: pocState.customColumnsFileId,
      state:               pocState
    });
  } catch (err) {
    pocState.status = 'error';
    logStep(err.message, 'error');
    res.status(500).json({ error: err.message });
  }
});

// =============================================================================
// STEP 0b — Upload PDF(s)
// =============================================================================
app.post('/poc/upload-to-project', upload.array('files'), async (req, res) => {
  try {
    if (!req.files || req.files.length === 0) throw new Error('No files received');

    pocState.status      = 'uploading';
    const injectModel    = req.body.injectStateModel !== 'false';
    logStep(`Received ${req.files.length} file(s) for upload (injectStateModel=${injectModel})`, 'info');
    if (injectModel) {
      logStep(`State model injection: will remove [${STATE_MODELS_TO_REMOVE.join(', ')}] and inject "${QC_STATE_MODEL.cUIName}"`, 'info');
    }

    const accessToken    = await tokenManager.getValidAccessToken();
    const reviewFolderId = pocState.folderIds[FOLDER_REVIEW_DOCS] || null;
    const uploaded       = [];

    for (const file of req.files) {
      let buffer = file.buffer;

      if (injectModel) {
        logStep(`Injecting "${QC_STATE_MODEL.cUIName}" state model into "${file.originalname}"...`, 'info');
        buffer = await injectStateModel(buffer);
        logStep(`State model injected into "${file.originalname}" (${buffer.length} bytes)`, 'success');
      }

      const result = await uploadFileToProject(buffer, file.originalname, accessToken, reviewFolderId);
      uploaded.push(result);
      pocState.projectFiles.push(result);

      db.insertFile({
        atkinsProjectId:       demoStub.atkinsProjectId || 'UNKNOWN',
        bluebeamProjectId:     POC_PROJECT_ID,
        bluebeamProjectFileId: result.projectFileId,
        bluebeamSessionId:     null,
        bluebeamSessionFileId: null,
        fileName:              result.name,
        fileSize:              result.size
      });
      logStep(`File written to DB: ${result.name} (projectFileId=${result.projectFileId})`, 'info');
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

// =============================================================================
// STEP 0c — Apply Custom Columns (optional)
// =============================================================================
app.post('/poc/apply-custom-columns', async (req, res) => {
  try {
    if (!pocState.customColumnsFileId) throw new Error('custom-columns.xml not uploaded — run setup-project first');
    if (pocState.projectFiles.length === 0) throw new Error('No project files — run upload-to-project first');

    logStep('Applying custom columns to project files...', 'info');
    const accessToken = await tokenManager.getValidAccessToken();
    const results     = [];

    for (const pf of pocState.projectFiles) {
      logStep(`Submitting importcustomcolumns job for "${pf.name}"...`, 'info');

      const jobResp = await fetch(
        `${API_V1}/projects/${POC_PROJECT_ID}/files/${pf.projectFileId}/jobs/importcustomcolumns`,
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
        logStep(`importcustomcolumns submission failed for "${pf.name}": ${jobResp.status} - ${err}`, 'warn');
        results.push({ name: pf.name, success: false, error: err });
        continue;
      }

      const { Id: jobId } = await jobResp.json();
      logStep(`Job submitted: jobId=${jobId} — polling...`, 'success');
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

// =============================================================================
// STEP 1 — Trigger
// =============================================================================
app.post('/poc/trigger', (req, res) => {
  pocState.status  = 'triggered';
  pocState.log     = [];
  logStep(`Workflow event received — document: ${demoStub.documentId}`, 'info');
  logStep(`Atkins Project: ${demoStub.atkinsProjectId || '(not set)'}`, 'info');
  logStep(`Files: ${pocState.projectFiles.map(f => f.name).join(', ') || '(none)'}`, 'info');
  logStep(`Description: ${demoStub.description}`, 'info');
  logStep(`Reviewers: ${demoStub.reviewers.map(r => r.email).join(', ')}`, 'info');
  logStep(`Session end date: ${new Date(demoStub.sessionEndDate).toLocaleDateString()}`, 'info');
  res.json({ success: true, state: pocState });
});

// =============================================================================
// STEP 2 — Create Session
// =============================================================================
app.post('/poc/create-session', async (req, res) => {
  try {
    pocState.status  = 'creating';
    logStep('Creating Bluebeam Studio Session...', 'info');

    const accessToken = await tokenManager.getValidAccessToken();
    const sessionName = `${demoStub.documentId}_Review_${new Date().toISOString().slice(0, 10)}`;

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

    const data          = await resp.json();
    pocState.sessionId  = data.Id;
    pocState.createdAt  = new Date().toISOString();

    logStep(`Session created: ID=${pocState.sessionId}`, 'success');
    logStep(`Session name: ${sessionName}`, 'info');
    logStep(`Session end date: ${demoStub.sessionEndDate}`, 'info');

    db.insertSession({
      atkinsProjectId:   demoStub.atkinsProjectId || 'UNKNOWN',
      bluebeamSessionId: pocState.sessionId,
      reviewName:        demoStub.documentId,
      discipline:        demoStub.discipline    || null,
      reviewType:        demoStub.reviewType    || null,
      pollingInterval:   demoStub.pollingInterval || 0
    });
    logStep(`Session written to DB: ${pocState.sessionId}`, 'info');

    if (demoStub.pollingInterval > 0) {
      scheduleSessionPoller(pocState.sessionId, demoStub.atkinsProjectId || 'UNKNOWN', demoStub.pollingInterval);
    }

    res.json({ success: true, sessionId: pocState.sessionId, state: pocState });
  } catch (err) {
    pocState.status = 'error';
    logStep(err.message, 'error');
    res.status(500).json({ error: err.message });
  }
});

// =============================================================================
// STEP 3 — Register Webhook
// =============================================================================
app.post('/poc/register-webhook', async (req, res) => {
  try {
    if (!pocState.sessionId) throw new Error('No active session — run create-session first');

    if (isLocalhost(WEBHOOK_CALLBACK_URL)) {
      logStep('Webhook skipped — WEBHOOK_CALLBACK_URL is localhost (Bluebeam requires public HTTPS)', 'warn');
      logStep('Set WEBHOOK_CALLBACK_URL env var to a public HTTPS URL (e.g. ngrok) to enable webhooks', 'warn');
      return res.json({ success: true, skipped: true, state: pocState });
    }

    logStep(`Registering webhook for session ${pocState.sessionId}...`, 'info');
    const accessToken = await tokenManager.getValidAccessToken();

    const resp = await fetch(`${API_V2}/subscriptions`, {
      method:  'POST',
      headers: authHeaders(accessToken),
      body:    JSON.stringify({
        sourceType:  'session',
        resourceId:  pocState.sessionId,
        callbackURI: WEBHOOK_CALLBACK_URL
      })
    });

    if (!resp.ok) throw new Error(`Webhook registration failed: ${resp.status} - ${await resp.text()}`);

    const data               = await resp.json();
    pocState.subscriptionId  = data.subscriptionId;

    logStep(`Webhook registered: subscriptionId=${pocState.subscriptionId}`, 'success');
    logStep(`Callback URL: ${WEBHOOK_CALLBACK_URL}`, 'info');
    res.json({ success: true, subscriptionId: pocState.subscriptionId, state: pocState });
  } catch (err) {
    pocState.status = 'error';
    logStep(err.message, 'error');
    res.status(500).json({ error: err.message });
  }
});

// =============================================================================
// STEP 4 — Checkout to Session
// =============================================================================
app.post('/poc/checkout-to-session', async (req, res) => {
  try {
    if (!pocState.sessionId)              throw new Error('No active session — run create-session first');
    if (pocState.projectFiles.length === 0) throw new Error('No project files — run upload-to-project first');

    pocState.status = 'checking-out';
    logStep(`Checking ${pocState.projectFiles.length} file(s) out to session ${pocState.sessionId}...`, 'info');

    const accessToken = await tokenManager.getValidAccessToken();
    const checked     = [];

    for (const pf of pocState.projectFiles) {
      logStep(`Checking out "${pf.name}" (projectFileId=${pf.projectFileId})...`, 'info');

      const resp = await fetch(
        `${API_V1}/projects/${POC_PROJECT_ID}/files/${pf.projectFileId}/checkout-to-session`,
        { method: 'POST', headers: authHeaders(accessToken), body: JSON.stringify({ SessionId: pocState.sessionId }) }
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
              { method: 'POST', headers: authHeaders(accessToken), body: JSON.stringify({ SessionId: pocState.sessionId }) }
            );
            if (!retry.ok) { logStep(`Retry failed for "${pf.name}": ${retry.status}`, 'warn'); continue; }
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
      const sessionFiles     = sessionFilesData.SessionFiles || sessionFilesData.Files || [];
      const match            = sessionFiles.find(f => f.ProjectFileId === pf.projectFileId || f.Name === pf.name);

      if (!match) {
        logStep(`Could not find session file entry for "${pf.name}" after checkout`, 'warn');
        continue;
      }

      const entry = { sessionFileId: match.Id, projectFileId: pf.projectFileId, name: pf.name };
      pocState.sessionFileIds.push(entry);
      checked.push(entry);

      db.updateFileSession({
        sessionId:     pocState.sessionId,
        sessionFileId: match.Id,
        projectFileId: pf.projectFileId
      });
      logStep(`"${pf.name}" checked out to session (sessionFileId=${match.Id}) — DB updated`, 'success');
    }

    logStep(`${checked.length} file(s) checked out to session`, 'success');
    res.json({ success: true, checked, state: pocState });
  } catch (err) {
    pocState.status = 'error';
    logStep(err.message, 'error');
    res.status(500).json({ error: err.message });
  }
});

// =============================================================================
// STEP 5 — Invite Reviewers
// =============================================================================
app.post('/poc/invite-reviewers', async (req, res) => {
  try {
    if (!pocState.sessionId) throw new Error('No active session');

    pocState.status = 'inviting';
    logStep(`Inviting ${demoStub.reviewers.length} reviewer(s)...`, 'info');

    const accessToken = await tokenManager.getValidAccessToken();
    const results     = [];

    for (const reviewer of demoStub.reviewers) {
      const endpoint = reviewer.hasStudioAccount
        ? `${API_V1}/sessions/${pocState.sessionId}/users`
        : `${API_V1}/sessions/${pocState.sessionId}/invite`;

      logStep(`Inviting ${reviewer.email} via ${reviewer.hasStudioAccount ? 'direct-add' : 'email-invite'}`, 'info');

      const resp = await fetch(endpoint, {
        method:  'POST',
        headers: authHeaders(accessToken),
        body:    JSON.stringify({
          Email:   reviewer.email,
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

// =============================================================================
// STEP 7 — Check In (standalone)
// =============================================================================
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

// =============================================================================
// STEP 8 — Export Markups (standalone)
// =============================================================================
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

// =============================================================================
// STEP 9 — Compatibility Route
// =============================================================================
app.post('/poc/run-markuplist-job', async (req, res) => {
  try {
    const accessToken = await tokenManager.getValidAccessToken();
    if (!pocState.markupExports.length) {
      logStep('No exported XML tracked yet — running export-markups first...', 'info');
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

// =============================================================================
// STEP 9b — Combined Downstream Processing
// =============================================================================
app.post('/poc/downstream-process', async (req, res) => {
  try {
    if (!pocState.sessionId) {
      throw new Error(
        'No active session in memory. If the server was restarted, run POST /poc/hydrate first to restore state from the DB.'
      );
    }
    if (pocState.sessionFileIds.length === 0) {
      throw new Error(
        'No session files in memory. If the server was restarted, run POST /poc/hydrate first to restore file references from the DB.'
      );
    }

    logStep('Starting downstream processing...', 'info');
    const accessToken = await tokenManager.getValidAccessToken();

    const checkinResults = await performCheckin(accessToken);
    logStep('Global post-checkin settle delay: 5s...', 'info');
    await sleep(5000);

    const exportResults = await performExportMarkups(accessToken);
    logStep('Global post-export settle delay: 5s...', 'info');
    await sleep(5000);

    const markups = await performMarkupExtractionFromXml(accessToken);

    db.insertSnapshot({
      atkinsProjectId:   demoStub.atkinsProjectId || 'UNKNOWN',
      bluebeamSessionId: pocState.sessionId,
      markupCount:       markups.length,
      snapshotJson:      JSON.stringify(markups)
    });

    db.updateSessionPolled(pocState.sessionId);

    logStep(`Snapshot saved to DB — ${markups.length} markup(s)`, 'success');
    logStep('Downstream processing complete', 'success');

    res.json({
      success:        true,
      checkinResults,
      exportResults,
      count:          markups.length,
      markups,
      extractionMode: 'exportmarkups-xml',
      markupExports:  pocState.markupExports,
      state:          pocState
    });
  } catch (err) {
    pocState.status = 'error';
    logStep(err.message, 'error');
    res.status(500).json({ error: err.message });
  }
});

// =============================================================================
// STEP 10 — Finalize Session
// =============================================================================
app.post('/poc/finalize', async (req, res) => {
  try {
    if (!pocState.sessionId) throw new Error('No active session');

    pocState.status = 'finalizing';
    logStep(`Setting session ${pocState.sessionId} to Finalizing...`, 'info');

    const accessToken = await tokenManager.getValidAccessToken();
    const resp        = await fetch(`${API_V1}/sessions/${pocState.sessionId}`, {
      method:  'PUT',
      headers: authHeaders(accessToken),
      body:    JSON.stringify({
        Name:           `${demoStub.documentId}_Review_${new Date().toISOString().slice(0, 10)}`,
        Restricted:     true,
        SessionEndDate: demoStub.sessionEndDate
      })
    });

    if (!resp.ok) throw new Error(`Finalize failed: ${resp.status} - ${await resp.text()}`);

    db.updateSessionStatus(pocState.sessionId, 'finalized');
    logStep('Session finalized', 'success');
    res.json({ success: true, state: pocState });
  } catch (err) {
    pocState.status = 'error';
    logStep(err.message, 'error');
    res.status(500).json({ error: err.message });
  }
});

// =============================================================================
// STEP 11 — Snapshot PDF
// =============================================================================
app.post('/poc/snapshot', async (req, res) => {
  try {
    if (!pocState.sessionId || pocState.sessionFileIds.length === 0) {
      throw new Error('No active session or no files. Run /poc/hydrate if server was restarted.');
    }

    pocState.status   = 'snapshotting';
    const accessToken = await tokenManager.getValidAccessToken();
    const downloads   = [];

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
        const pollResp  = await fetch(
          `${API_V1}/sessions/${pocState.sessionId}/files/${sf.sessionFileId}/snapshot`,
          { headers: authHeaders(pollToken) }
        );
        if (!pollResp.ok) continue;
        const d = await pollResp.json();
        logStep(`Snapshot poll ${i + 1}: ${d.Status}`, 'info');
        if (d.Status === 'Complete') { downloadUrl = d.DownloadUrl; break; }
        if (d.Status === 'Error')    throw new Error(`Snapshot error: ${d.Message}`);
      }

      if (!downloadUrl) { logStep(`Snapshot timed out for "${sf.name}"`, 'warn'); continue; }

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
    db.updateSessionStatus(pocState.sessionId, 'complete');
    logStep('Snapshots complete', 'success');
    res.json({ success: true, downloads, state: pocState });
  } catch (err) {
    pocState.status = 'error';
    logStep(err.message, 'error');
    res.status(500).json({ error: err.message });
  }
});

// =============================================================================
// STEP 12 — Cleanup
// =============================================================================
app.post('/poc/cleanup', async (req, res) => {
  try {
    if (!pocState.sessionId) throw new Error('No active session to clean up');

    const accessToken = await tokenManager.getValidAccessToken();

    if (pocState.subscriptionId) {
      const subResp = await fetch(`${API_V2}/subscriptions/${pocState.subscriptionId}`, {
        method: 'DELETE', headers: authHeaders(accessToken)
      });
      logStep(subResp.ok ? 'Webhook subscription deleted' : `Sub delete: ${subResp.status}`, subResp.ok ? 'success' : 'warn');
    }

    const sessResp = await fetch(`${API_V1}/sessions/${pocState.sessionId}`, {
      method: 'DELETE', headers: authHeaders(accessToken)
    });
    logStep(sessResp.ok ? 'Session deleted' : `Session delete: ${sessResp.status}`, sessResp.ok ? 'success' : 'warn');

    db.updateSessionStatus(pocState.sessionId, 'cleaned');
    scheduleSessionPoller(pocState.sessionId, null, 0);

    logStep('Cleanup complete', 'success');
    pocState.sessionId      = null;
    pocState.subscriptionId = null;
    res.json({ success: true, state: pocState });
  } catch (err) {
    logStep(err.message, 'error');
    res.status(500).json({ error: err.message });
  }
});

// =============================================================================
// STANDALONE: inject state model into a single PDF for testing
// =============================================================================
app.post('/poc/inject-state-model', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
    const modified = await injectStateModel(req.file.buffer);
    const outName  = req.file.originalname.replace(/\.pdf$/i, '') + '_state_injected.pdf';
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `attachment; filename="${outName}"`);
    res.setHeader('Content-Length', modified.length);
    res.send(modified);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// =============================================================================
// DB QUERY ENDPOINTS
// =============================================================================
app.get('/poc/projects/:atkinsId/debug-snapshot', (req, res) => {
  try {
    const { atkinsId } = req.params;
    const sessions = db.getSessionsByProject(atkinsId);
    const results = sessions.map(s => {
      const snap = db.getLatestSnapshotBySession(atkinsId, s.bluebeam_session_id);
      if (!snap || !snap.snapshot_json) return { session: s.bluebeam_session_id, markups: [] };
      let markups = [];
      try { markups = JSON.parse(snap.snapshot_json); } catch (_) {}
      return {
        session: s.bluebeam_session_id,
        markup_count: markups.length,
        sample: markups.slice(0, 3).map(m => ({
          Id: m.Id, Author: m.Author, Subject: m.Subject,
          Status: m.Status, Type: m.Type, Page: m.Page,
          CustomKeys: Object.keys(m.Custom || {}),
          ExtendedKeys: Object.keys(m.ExtendedProperties || {}).slice(0, 20),
          ExtendedSample: Object.fromEntries(Object.entries(m.ExtendedProperties || {}).slice(0, 10))
        }))
      };
    });
    res.json({ atkinsId, results });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.get('/poc/projects', (req, res) => {
  try {
    const projects = db.listProjects();
    res.json({ projects });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.get('/poc/projects/:atkinsId/snapshot', (req, res) => {
  try {
    const { atkinsId } = req.params;
    const project      = db.getProject(atkinsId);
    const sessions     = db.getSessionsByProject(atkinsId);
    const files        = db.getFilesByProject(atkinsId);

    if (!project && !sessions.length) {
      return res.status(404).json({
        error: `No data found for project ${atkinsId}`,
        hint:  'Start a review from Tab 1 with this Atkins Project ID to populate data.'
      });
    }

    const sessionsWithSnapshots = sessions.map(s => {
      const snap    = db.getLatestSnapshotBySession(atkinsId, s.bluebeam_session_id);
      let markups   = [];
      if (snap && snap.snapshot_json) {
        try { markups = JSON.parse(snap.snapshot_json); } catch (_) {}
      }
      return {
        ...s,
        markup_count:  snap ? snap.markup_count : 0,
        last_snapshot: snap ? snap.created_at   : null,
        markups,
        files: files.filter(f => f.bluebeam_session_id === s.bluebeam_session_id)
      };
    });

    const totalMarkups = sessionsWithSnapshots.reduce((sum, s) => sum + (s.markup_count || 0), 0);

    res.json({
      project:      project || { atkins_project_id: atkinsId },
      sessions:     sessionsWithSnapshots,
      files,
      totalMarkups,
      sessionCount: sessions.length,
      fileCount:    files.length
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// =============================================================================
// WEBHOOK LISTENER
// =============================================================================
app.post('/webhook/studio-events', (req, res) => {
  const p = req.body || {};
  logStep(`Webhook: ${p.ResourceType || 'unknown'} / ${p.EventType || 'unknown'}`, 'webhook');
  pocState.webhookEvents.push({ ...p, receivedAt: new Date().toISOString() });
  res.sendStatus(200);
});

// =============================================================================
// STANDALONE MARKUP API
// =============================================================================
app.get('/api/project-markups', (req, res) => {
  if (!pocState.markups.length) {
    return res.status(404).json({ error: 'No markup data. Run downstream processing first.' });
  }
  res.json(pocState.markups.map(m => ({
    MarkupId:           m.Id,
    Author:             m.Author,
    Type:               m.Type,
    Subject:            m.Subject,
    Comment:            m.Comment,
    Status:             m.Status,
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

// =============================================================================
// START
// =============================================================================
app.listen(PORT, () => {
  console.log(`\nBluebeam Studio PoC  →  http://localhost:${PORT}`);
  console.log(`Project: ${POC_PROJECT_ID}`);
  console.log(`DB: ${process.env.DB_PATH || './poc.db'}`);

  if (isLocalhost(WEBHOOK_CALLBACK_URL)) console.log('⚠  Webhook will be skipped (localhost URL)');

  console.log(`\nState Model: "${QC_STATE_MODEL.cUIName}" (${QC_STATE_MODEL.states.length} states) — injection ON by default`);
  console.log(`Models removed before injection: ${STATE_MODELS_TO_REMOVE.join(', ')}`);

  console.log(`\nFLOW:`);
  console.log(`  POST /poc/setup-project         — 0a: Create folders + upload custom-columns.xml`);
  console.log(`  POST /poc/upload-to-project     — 0b: State model injection (default ON) + upload PDFs → project + DB`);
  console.log(`  POST /poc/apply-custom-columns  — 0c: Apply custom columns (optional)`);
  console.log(`  POST /poc/hydrate               — H:  Restore pocState from DB after restart`);
  console.log(`  POST /poc/trigger               — 1:  Workflow event`);
  console.log(`  POST /poc/create-session        — 2:  Create session + DB write + schedule poller`);
  console.log(`  POST /poc/register-webhook      — 3:  Register webhook`);
  console.log(`  POST /poc/checkout-to-session   — 4:  Check out files + update DB files table`);
  console.log(`  POST /poc/invite-reviewers      — 5:  Invite reviewers`);
  console.log(`       (6: Review in Revu)`);
  console.log(`  POST /poc/checkin               — 7:  Check in session files`);
  console.log(`  POST /poc/export-markups        — 8:  Export markups to XML`);
  console.log(`  POST /poc/run-markuplist-job    — 9:  XML-backed extraction`);
  console.log(`  POST /poc/downstream-process    — 9b: Checkin + export + XML extract + DB snapshot`);
  console.log(`  POST /poc/finalize              — 10: Finalize session + update DB status`);
  console.log(`  POST /poc/snapshot              — 11: Snapshot + download PDF`);
  console.log(`  POST /poc/cleanup               — 12: Delete webhook + session + clear poller`);
  console.log(`\nAUTH:`);
  console.log(`  POST /poc/auth/refresh          — Rotate access token via refresh token`);
  console.log(`\nDB ENDPOINTS:`);
  console.log(`  GET  /poc/projects                     — List all projects`);
  console.log(`  GET  /poc/projects/:atkinsId/snapshot  — Full project snapshot for Tab 2`);
  console.log(`\nUTILITY:`);
  console.log(`  GET  /health                           — Server + DB + pocState status`);
  console.log(`  GET  /poc/state                        — Current in-memory state`);
  console.log(`  POST /poc/reset                        — Clear in-memory state\n`);

  restorePollers();
});
