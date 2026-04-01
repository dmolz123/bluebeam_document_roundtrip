/**
 * Proof-of-concept reference implementation for a Bluebeam Studio roundtrip workflow.
 * Intended for evaluation and development reference only.
 *
 * Roundtrip flow:
 *   0.  /poc/upload-to-project      — Upload PDF(s) from UI → Studio Project 712-566-288
 *   1.  /poc/trigger               — Simulate source-system workflow event
 *   2.  /poc/create-session        — Create a Studio Session
 *   3.  /poc/register-webhook      — Subscribe to session events (skipped gracefully if localhost)
 *   4.  /poc/add-to-session        — Look up project file ID, add file to session
 *   5.  /poc/invite-reviewers      — Invite dmolz@bluebeam.com + any additional reviewers
 *   6.  (Review happens in Bluebeam Revu — no API step)
 *   7.  /poc/update-project-copy   — Push session markups back to project file
 *   8.  /poc/run-markuplist-job    — Run markuplist job on project file, poll, return markups
 *   9.  /poc/finalize              — Set session status to Finalizing
 *   10. /poc/snapshot              — Create + poll snapshot, download merged PDF
 *   11. /poc/cleanup               — Delete webhook subscription + session
 */

require('dotenv').config();
const express  = require('express');
const cors     = require('cors');
const fs       = require('fs');
const path     = require('path');
const multer   = require('multer');
const TokenManager = require('./tokenManager');

const app    = express();
const PORT   = process.env.PORT || 3000;
const upload = multer({ storage: multer.memoryStorage() }); // files held in memory, sent straight to BB

// -----------------------------------------------------------------------------
// API BASE URLs
// -----------------------------------------------------------------------------
const API_V1 = 'https://api.bluebeam.com/publicapi/v1';
const API_V2 = 'https://api.bluebeam.com/publicapi/v2';

const CLIENT_ID            = process.env.BB_CLIENT_ID;
const WEBHOOK_CALLBACK_URL =
  process.env.WEBHOOK_CALLBACK_URL ||
  `http://localhost:${PORT}/webhook/studio-events`;

// Hardcoded Studio Project for this PoC
const POC_PROJECT_ID = '712-566-288';

// -----------------------------------------------------------------------------
// OPTIONAL SAMPLE / REFERENCE CONSTANTS
// Populate in .env to enable the standalone /powerbi/markups endpoint.
// -----------------------------------------------------------------------------
const MARKUP_SESSION_ID = process.env.MARKUP_SESSION_ID || '';
const MARKUP_FILE_ID    = process.env.MARKUP_FILE_ID    || '';
const MARKUP_FILE_NAME  = process.env.MARKUP_FILE_NAME  || 'Sample Drawing.pdf';

// -----------------------------------------------------------------------------
// DEMO STUB
// -----------------------------------------------------------------------------
let demoStub = {
  documentId:   process.env.DEMO_DOCUMENT_ID  || 'DOC-001',
  description:  process.env.DEMO_DESCRIPTION  || 'Design review — coordination update',
  // Hardcoded primary reviewer; additional reviewers added via UI
  reviewers: [
    { email: 'dmolz@bluebeam.com', hasStudioAccount: true }
  ],
  sessionEndDate: new Date(Date.now() + 7 * 24 * 60 * 60 * 1000).toISOString()
};

// -----------------------------------------------------------------------------
// IN-MEMORY STATE
// -----------------------------------------------------------------------------
let pocState = {
  sessionId:        null,
  subscriptionId:   null,
  projectFiles:     [],   // [{ projectFileId, name, size }] — set after upload to project
  sessionFileIds:   [],   // [{ sessionFileId, projectFileId, name }] — set after add-to-session
  markups:          [],
  markupJobId:      null,
  status:           'idle',
  log:              [],
  createdAt:        null,
  webhookEvents:    []
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
    sessionId:      null,
    subscriptionId: null,
    projectFiles:   [],
    sessionFileIds: [],
    markups:        [],
    markupJobId:    null,
    status:         'idle',
    log:            [],
    createdAt:      null,
    webhookEvents:  []
  };
}

function ensureConfigured(value, name) {
  if (!value) throw new Error(`Missing required configuration: ${name}`);
}

function isLocalhost(url) {
  return /^https?:\/\/(localhost|127\.0\.0\.1)/.test(url);
}

/**
 * Generic async job poller for Bluebeam project file jobs.
 * Polls GET {url} every {intervalMs} ms until JobStatus reaches
 * 2 (Complete) or 3 (Error), or maxAttempts is exhausted.
 */
async function pollJob(url, headers, maxAttempts = 30, intervalMs = 3000) {
  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    await new Promise(resolve => setTimeout(resolve, intervalMs));
    const res  = await fetch(url, { headers });
    const data = await res.json();
    logStep(`Job poll ${attempt}/${maxAttempts}: ${data.JobStatusMessage} (${data.JobStatus})`, 'info');
    if (data.JobStatus === 2) return data;
    if (data.JobStatus === 3) throw new Error(`Job failed: ${data.JobStatusMessage}`);
  }
  throw new Error(`Job did not complete after ${maxAttempts} attempts`);
}

/**
 * Upload a single file buffer to the Studio Project (3-step).
 * Returns { projectFileId, name }.
 */
async function uploadFileToProject(fileBuffer, fileName, accessToken) {
  logStep(`Uploading "${fileName}" to project ${POC_PROJECT_ID}...`, 'info');

  // Step A — create metadata block in project
  const metaResp = await fetch(`${API_V1}/projects/${POC_PROJECT_ID}/files`, {
    method: 'POST',
    headers: {
      Authorization:  `Bearer ${accessToken}`,
      client_id:      CLIENT_ID,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({ Name: fileName })
  });

  if (!metaResp.ok) {
    const err = await metaResp.text();
    throw new Error(`Project metadata block failed for "${fileName}": ${metaResp.status} - ${err}`);
  }

  const meta              = await metaResp.json();
  const projectFileId     = meta.Id;
  const uploadUrl         = meta.UploadUrl;
  const uploadContentType = meta.UploadContentType || 'application/pdf';

  logStep(`Project metadata block created: projectFileId=${projectFileId}`, 'success');

  // Step B — PUT bytes to S3
  logStep(`Uploading binary (${fileBuffer.length} bytes)...`, 'info');
  const s3Resp = await fetch(uploadUrl, {
    method: 'PUT',
    headers: {
      'x-amz-server-side-encryption': 'AES256',
      'Content-Type':                  uploadContentType
    },
    body: fileBuffer
  });

  if (!s3Resp.ok) throw new Error(`S3 upload failed for "${fileName}": ${s3Resp.status}`);
  logStep('Binary upload to S3 complete', 'success');

  // Step C — confirm with Bluebeam
  const confirmResp = await fetch(
    `${API_V1}/projects/${POC_PROJECT_ID}/files/${projectFileId}/confirm-upload`,
    {
      method: 'POST',
      headers: { Authorization: `Bearer ${accessToken}`, client_id: CLIENT_ID }
    }
  );

  if (!confirmResp.ok) {
    const err = await confirmResp.text();
    throw new Error(`Confirm upload failed for "${fileName}": ${confirmResp.status} - ${err}`);
  }

  logStep(`"${fileName}" confirmed in project (projectFileId=${projectFileId})`, 'success');
  return { projectFileId, name: fileName, size: fileBuffer.length };
}

/**
 * Look up a file in the project by name.
 * Returns the project file object or null.
 */
async function findProjectFileByName(fileName, accessToken) {
  const resp = await fetch(
    `${API_V1}/projects/${POC_PROJECT_ID}/files`,
    {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        client_id:     CLIENT_ID,
        Accept:        'application/json'
      }
    }
  );

  if (!resp.ok) {
    const err = await resp.text();
    throw new Error(`Failed to list project files: ${resp.status} - ${err}`);
  }

  const data  = await resp.json();
  const files = data.ProjectFiles || [];
  return files.find(f => f.Name === fileName) || null;
}

// -----------------------------------------------------------------------------
// HEALTH CHECK
// -----------------------------------------------------------------------------
app.get('/health', (req, res) => {
  res.json({
    status: 'healthy',
    projectId: POC_PROJECT_ID,
    config: {
      hasClientId:        Boolean(CLIENT_ID),
      webhookCallbackUrl: WEBHOOK_CALLBACK_URL,
      webhookIsLocalhost: isLocalhost(WEBHOOK_CALLBACK_URL)
    }
  });
});

// =============================================================================
// POC ROUTES
// =============================================================================

app.get('/poc/state', (req, res) => {
  res.json({ ...pocState, stub: demoStub, projectId: POC_PROJECT_ID });
});

app.get('/poc/stub', (req, res) => {
  res.json(demoStub);
});

// POST configure — update description, documentId, or add extra reviewers
// Body: { documentId?, description?, reviewerEmail? }
// Note: dmolz@bluebeam.com is always present as the primary reviewer
app.post('/poc/configure', (req, res) => {
  const { documentId, description, reviewerEmail } = req.body || {};

  if (documentId)  demoStub.documentId  = documentId;
  if (description) demoStub.description = description;

  if (reviewerEmail && reviewerEmail !== 'dmolz@bluebeam.com') {
    // Add to list if not already present
    const exists = demoStub.reviewers.some(r => r.email === reviewerEmail);
    if (!exists) {
      demoStub.reviewers.push({ email: reviewerEmail, hasStudioAccount: false });
      logStep(`Added reviewer: ${reviewerEmail}`, 'info');
    }
  }

  demoStub.sessionEndDate = new Date(Date.now() + 7 * 24 * 60 * 60 * 1000).toISOString();
  res.json({ success: true, stub: demoStub });
});

// Remove a reviewer (cannot remove dmolz@bluebeam.com)
app.post('/poc/remove-reviewer', (req, res) => {
  const { email } = req.body || {};
  if (email === 'dmolz@bluebeam.com')
    return res.status(400).json({ error: 'Cannot remove primary reviewer' });

  demoStub.reviewers = demoStub.reviewers.filter(r => r.email !== email);
  res.json({ success: true, stub: demoStub });
});

app.post('/poc/reset', (req, res) => {
  resetPocState();
  // Reset reviewers to just the primary
  demoStub.reviewers = [{ email: 'dmolz@bluebeam.com', hasStudioAccount: true }];
  logStep('PoC state reset', 'info');
  res.json({ success: true });
});

// -----------------------------------------------------------------------------
// STEP 0 — Upload file(s) from UI to Studio Project
//
// Accepts multipart/form-data with field name "files".
// Each file is uploaded to project 712-566-288 via the 3-step flow,
// then the project file list is queried to confirm the file ID.
// Supports multiple files in a single request.
// -----------------------------------------------------------------------------
app.post('/poc/upload-to-project', upload.array('files'), async (req, res) => {
  try {
    ensureConfigured(CLIENT_ID, 'BB_CLIENT_ID');

    if (!req.files || req.files.length === 0)
      throw new Error('No files received — attach files with field name "files"');

    pocState.status = 'uploading';
    logStep(`Received ${req.files.length} file(s) for upload to project ${POC_PROJECT_ID}`, 'info');

    const accessToken = await tokenManager.getValidAccessToken();
    const uploaded    = [];

    for (const file of req.files) {
      const result = await uploadFileToProject(file.buffer, file.originalname, accessToken);
      uploaded.push(result);
      pocState.projectFiles.push(result);
    }

    // Update document name in stub to match first uploaded file
    if (uploaded.length > 0) {
      demoStub.documentId = demoStub.documentId || uploaded[0].name.replace(/\.[^.]+$/, '');
    }

    logStep(`${uploaded.length} file(s) uploaded to project successfully`, 'success');
    res.json({ success: true, uploaded, state: pocState });

  } catch (err) {
    pocState.status = 'error';
    logStep(err.message, 'error');
    res.status(500).json({ error: err.message });
  }
});

// -----------------------------------------------------------------------------
// STEP 1 — Simulate source-system workflow event
// -----------------------------------------------------------------------------
app.post('/poc/trigger', (req, res) => {
  pocState.status = 'triggered';
  pocState.log    = [];

  const fileNames = pocState.projectFiles.map(f => f.name).join(', ') || '(none uploaded yet)';
  logStep(`Workflow event received — document: ${demoStub.documentId}`, 'info');
  logStep(`Files staged for review: ${fileNames}`, 'info');
  logStep(`Description: ${demoStub.description}`, 'info');
  logStep(`Reviewers: ${demoStub.reviewers.map(r => r.email).join(', ')}`, 'info');
  logStep(`Session end date: ${new Date(demoStub.sessionEndDate).toLocaleDateString()}`, 'info');

  res.json({ success: true, state: pocState });
});

// -----------------------------------------------------------------------------
// STEP 2 — Create Studio Session
// -----------------------------------------------------------------------------
app.post('/poc/create-session', async (req, res) => {
  try {
    ensureConfigured(CLIENT_ID, 'BB_CLIENT_ID');

    pocState.status = 'creating';
    logStep('Creating Bluebeam Studio Session...', 'info');

    const accessToken = await tokenManager.getValidAccessToken();
    const sessionName = `${demoStub.documentId}_Review_${new Date().toISOString().slice(0,10)}`;

    const body = {
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
    };

    const response = await fetch(`${API_V1}/sessions`, {
      method: 'POST',
      headers: {
        Authorization:  `Bearer ${accessToken}`,
        client_id:      CLIENT_ID,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(body)
    });

    if (!response.ok) {
      const err = await response.text();
      throw new Error(`Session creation failed: ${response.status} - ${err}`);
    }

    const data         = await response.json();
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
// STEP 3 — Register Webhook Subscription
//
// Skipped gracefully if WEBHOOK_CALLBACK_URL is localhost — Bluebeam requires
// a publicly accessible HTTPS URL. Use ngrok or similar for local testing.
// -----------------------------------------------------------------------------
app.post('/poc/register-webhook', async (req, res) => {
  try {
    ensureConfigured(CLIENT_ID, 'BB_CLIENT_ID');

    if (!pocState.sessionId)
      throw new Error('No active session — run create-session first');

    // Graceful skip for localhost
    if (isLocalhost(WEBHOOK_CALLBACK_URL)) {
      logStep('Webhook skipped — WEBHOOK_CALLBACK_URL is localhost (not reachable by Bluebeam)', 'warn');
      logStep('Set WEBHOOK_CALLBACK_URL to a public HTTPS URL (e.g. ngrok) to enable webhooks', 'warn');
      return res.json({ success: true, skipped: true, reason: 'localhost', state: pocState });
    }

    logStep(`Registering webhook for session ${pocState.sessionId}...`, 'info');

    const accessToken = await tokenManager.getValidAccessToken();

    const response = await fetch(`${API_V2}/subscriptions`, {
      method: 'POST',
      headers: {
        Authorization:  `Bearer ${accessToken}`,
        client_id:      CLIENT_ID,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        sourceType:  'session',
        resourceId:  pocState.sessionId,
        callbackURI: WEBHOOK_CALLBACK_URL
      })
    });

    if (!response.ok) {
      const err = await response.text();
      throw new Error(`Webhook registration failed: ${response.status} - ${err}`);
    }

    const data              = await response.json();
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
// STEP 4 — Add Project File(s) to Session
//
// For each file uploaded to the project:
//   a. Query GET /projects/{projectId}/files to find the file by name and get its Id
//   b. Get the project file download URL via GET /projects/{projectId}/files/{fileId}
//   c. Create a session file metadata block using that URL as Source
//   d. Confirm the session file upload
//
// This is the "check out into session" operation — no bytes are re-uploaded.
// -----------------------------------------------------------------------------
app.post('/poc/add-to-session', async (req, res) => {
  try {
    ensureConfigured(CLIENT_ID, 'BB_CLIENT_ID');

    if (!pocState.sessionId)
      throw new Error('No active session — run create-session first');
    if (pocState.projectFiles.length === 0)
      throw new Error('No project files found — run upload-to-project first');

    pocState.status = 'adding-to-session';
    logStep(`Adding ${pocState.projectFiles.length} project file(s) to session ${pocState.sessionId}...`, 'info');

    const accessToken = await tokenManager.getValidAccessToken();
    const added       = [];

    for (const projectFile of pocState.projectFiles) {

      // 4a — look up file in project by name to get confirmed Id
      logStep(`Looking up "${projectFile.name}" in project ${POC_PROJECT_ID}...`, 'info');
      const found = await findProjectFileByName(projectFile.name, accessToken);

      if (!found) {
        logStep(`File "${projectFile.name}" not found in project — skipping`, 'warn');
        continue;
      }

      const projectFileId = found.Id;
      logStep(`Found in project: "${projectFile.name}" (projectFileId=${projectFileId})`, 'success');

      // 4b — get file download URL from project
      const fileDetailResp = await fetch(
        `${API_V1}/projects/${POC_PROJECT_ID}/files/${projectFileId}/download`,
        {
          headers: { Authorization: `Bearer ${accessToken}`, client_id: CLIENT_ID }
        }
      );

      if (!fileDetailResp.ok) {
        const err = await fileDetailResp.text();
        throw new Error(`Failed to get download URL for "${projectFile.name}": ${fileDetailResp.status} - ${err}`);
      }

      const fileDetail = await fileDetailResp.json();
      const downloadUrl = fileDetail.Url || fileDetail.DownloadUrl || fileDetail.url;

      if (!downloadUrl)
        throw new Error(`No download URL returned for "${projectFile.name}"`);

      logStep(`Download URL obtained for "${projectFile.name}"`, 'success');

      // 4c — create session file metadata block using project download URL as Source
      logStep(`Creating session file metadata block for "${projectFile.name}"...`, 'info');

      const sessionMetaResp = await fetch(`${API_V1}/sessions/${pocState.sessionId}/files`, {
        method: 'POST',
        headers: {
          Authorization:  `Bearer ${accessToken}`,
          client_id:      CLIENT_ID,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          Name:   projectFile.name,
          Source: downloadUrl
        })
      });

      if (!sessionMetaResp.ok) {
        const err = await sessionMetaResp.text();
        throw new Error(`Session file metadata failed for "${projectFile.name}": ${sessionMetaResp.status} - ${err}`);
      }

      const sessionMeta   = await sessionMetaResp.json();
      const sessionFileId = sessionMeta.Id;
      const uploadUrl     = sessionMeta.UploadUrl;

      logStep(`Session file metadata created: sessionFileId=${sessionFileId}`, 'success');

      // 4d — If UploadUrl is returned, the file bytes must still be transferred
      //      (Bluebeam may return an UploadUrl even for source-referenced files)
      if (uploadUrl) {
        logStep('UploadUrl returned — fetching source bytes and transferring to session storage...', 'info');

        const srcResp = await fetch(downloadUrl);
        if (!srcResp.ok) throw new Error(`Failed to fetch source file: ${srcResp.status}`);

        const srcBuffer = await srcResp.buffer();

        const s3Resp = await fetch(uploadUrl, {
          method: 'PUT',
          headers: {
            'x-amz-server-side-encryption': 'AES256',
            'Content-Type':                  'application/pdf'
          },
          body: srcBuffer
        });

        if (!s3Resp.ok) throw new Error(`Session S3 upload failed: ${s3Resp.status}`);
        logStep(`Bytes transferred to session storage (${srcBuffer.length} bytes)`, 'success');
      }

      // 4e — confirm session file upload
      logStep('Confirming session file upload...', 'info');

      const confirmResp = await fetch(
        `${API_V1}/sessions/${pocState.sessionId}/files/${sessionFileId}/confirm-upload`,
        {
          method: 'POST',
          headers: { Authorization: `Bearer ${accessToken}`, client_id: CLIENT_ID }
        }
      );

      if (!confirmResp.ok) {
        const err = await confirmResp.text();
        throw new Error(`Session confirm upload failed for "${projectFile.name}": ${confirmResp.status} - ${err}`);
      }

      const entry = { sessionFileId, projectFileId, name: projectFile.name };
      pocState.sessionFileIds.push(entry);
      added.push(entry);

      logStep(`"${projectFile.name}" active in session (sessionFileId=${sessionFileId})`, 'success');
    }

    logStep(`${added.length} file(s) added to session`, 'success');
    res.json({ success: true, added, state: pocState });

  } catch (err) {
    pocState.status = 'error';
    logStep(err.message, 'error');
    res.status(500).json({ error: err.message });
  }
});

// -----------------------------------------------------------------------------
// STEP 5 — Invite Reviewers
//
// dmolz@bluebeam.com is always the primary reviewer (hasStudioAccount: true → /users).
// Additional reviewers added via the UI use /invite (email flow).
// -----------------------------------------------------------------------------
app.post('/poc/invite-reviewers', async (req, res) => {
  try {
    ensureConfigured(CLIENT_ID, 'BB_CLIENT_ID');

    if (!pocState.sessionId) throw new Error('No active session');

    pocState.status = 'inviting';
    logStep(`Inviting ${demoStub.reviewers.length} reviewer(s)...`, 'info');

    const accessToken = await tokenManager.getValidAccessToken();
    const results     = [];

    for (const reviewer of demoStub.reviewers) {
      // hasStudioAccount → direct add via /users (no email sent, appears in their Revu immediately)
      // otherwise       → /invite sends an email with a join link
      const endpoint = reviewer.hasStudioAccount
        ? `${API_V1}/sessions/${pocState.sessionId}/users`
        : `${API_V1}/sessions/${pocState.sessionId}/invite`;

      logStep(
        `Inviting ${reviewer.email} via ${reviewer.hasStudioAccount ? 'direct-add (/users)' : 'email-invite (/invite)'}`,
        'info'
      );

      const resp = await fetch(endpoint, {
        method: 'POST',
        headers: {
          Authorization:  `Bearer ${accessToken}`,
          client_id:      CLIENT_ID,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          Email:     reviewer.email,
          SendEmail: true,
          Message:   `You have been invited to review ${demoStub.documentId}: ${demoStub.description}`
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
    logStep(`Join via Bluebeam Revu — session ID: ${pocState.sessionId}`, 'info');
    res.json({ success: true, results, state: pocState });

  } catch (err) {
    pocState.status = 'error';
    logStep(err.message, 'error');
    res.status(500).json({ error: err.message });
  }
});

// -----------------------------------------------------------------------------
// STEP 6 — Non-API: Review in Bluebeam Revu
// -----------------------------------------------------------------------------

// -----------------------------------------------------------------------------
// STEP 7 — Update the Project File Copy
//
// Pushes session markups back to each linked project file.
// Uses the SESSION file ID (from pocState.sessionFileIds), not the project file ID.
// Non-destructive — session remains active after this call.
// -----------------------------------------------------------------------------
app.post('/poc/update-project-copy', async (req, res) => {
  try {
    ensureConfigured(CLIENT_ID, 'BB_CLIENT_ID');

    if (!pocState.sessionId)              throw new Error('No active session');
    if (pocState.sessionFileIds.length === 0) throw new Error('No files in session — run add-to-session first');

    pocState.status = 'updating-project';

    const accessToken = await tokenManager.getValidAccessToken();
    const results     = [];

    for (const sf of pocState.sessionFileIds) {
      logStep(`Updating project copy for "${sf.name}" (sessionFileId=${sf.sessionFileId})...`, 'info');

      const resp = await fetch(
        `${API_V1}/sessions/${pocState.sessionId}/files/${sf.sessionFileId}/updateprojectcopy`,
        {
          method: 'POST',
          headers: { Authorization: `Bearer ${accessToken}`, client_id: CLIENT_ID }
        }
      );

      if (resp.status === 204) {
        logStep(`"${sf.name}" — project copy updated (204)`, 'success');
        results.push({ name: sf.name, success: true });
      } else {
        const err = await resp.text();
        logStep(`"${sf.name}" — update failed: ${resp.status} - ${err}`, 'warn');
        results.push({ name: sf.name, success: false, error: err });
      }
    }

    logStep('Project file copy update complete — session remains active', 'info');
    res.json({ success: true, results, state: pocState });

  } catch (err) {
    pocState.status = 'error';
    logStep(err.message, 'error');
    res.status(500).json({ error: err.message });
  }
});

// -----------------------------------------------------------------------------
// STEP 8 — Run Markup List Job on the Project File
//
// Runs against the first project file. If multiple files were uploaded,
// runs against each in sequence and merges the markup arrays.
// -----------------------------------------------------------------------------
app.post('/poc/run-markuplist-job', async (req, res) => {
  try {
    ensureConfigured(CLIENT_ID, 'BB_CLIENT_ID');

    if (pocState.sessionFileIds.length === 0)
      throw new Error('No session files — run add-to-session first');

    pocState.status  = 'extracting-markups';
    pocState.markups = [];

    const accessToken = await tokenManager.getValidAccessToken();

    const authHeaders = {
      Authorization:  `Bearer ${accessToken}`,
      client_id:      CLIENT_ID,
      'Content-Type': 'application/json',
      Accept:         'application/json'
    };

    for (const sf of pocState.sessionFileIds) {
      logStep(`Submitting markuplist job for "${sf.name}" (projectFileId=${sf.projectFileId})...`, 'info');

      const submitResp = await fetch(
        `${API_V1}/projects/${POC_PROJECT_ID}/files/${sf.projectFileId}/jobs/markuplist`,
        { method: 'POST', headers: authHeaders, body: JSON.stringify({}) }
      );

      if (!submitResp.ok) {
        const err = await submitResp.text();
        logStep(`Markuplist job submission failed for "${sf.name}": ${submitResp.status} - ${err}`, 'warn');
        continue;
      }

      const { JobId }      = await submitResp.json();
      pocState.markupJobId = JobId;

      logStep(`Job submitted: JobId=${JobId} — polling...`, 'success');

      const pollUrl = `${API_V1}/projects/${POC_PROJECT_ID}/files/${sf.projectFileId}/jobs/markuplist/${JobId}`;
      const result  = await pollJob(pollUrl, authHeaders);

      const fileMarkups = (result.Markups || []).map(m => ({ ...m, _sourceFile: sf.name }));
      pocState.markups.push(...fileMarkups);

      logStep(`"${sf.name}" — ${fileMarkups.length} markup(s) extracted`, 'success');
    }

    pocState.status = 'active';
    logStep(`Markuplist complete — ${pocState.markups.length} total markup(s) extracted`, 'success');

    res.json({
      success: true,
      count:   pocState.markups.length,
      markups: pocState.markups,
      state:   pocState
    });

  } catch (err) {
    pocState.status = 'error';
    logStep(err.message, 'error');
    res.status(500).json({ error: err.message });
  }
});

// -----------------------------------------------------------------------------
// STEP 9 — Finalize Session
// -----------------------------------------------------------------------------
app.post('/poc/finalize', async (req, res) => {
  try {
    ensureConfigured(CLIENT_ID, 'BB_CLIENT_ID');

    if (!pocState.sessionId) throw new Error('No active session');

    pocState.status = 'finalizing';
    logStep(`Setting session ${pocState.sessionId} to Finalizing...`, 'info');

    const accessToken = await tokenManager.getValidAccessToken();

    const resp = await fetch(`${API_V1}/sessions/${pocState.sessionId}`, {
      method: 'PUT',
      headers: {
        Authorization:  `Bearer ${accessToken}`,
        client_id:      CLIENT_ID,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ Status: 'Finalizing' })
    });

    if (!resp.ok) {
      const err = await resp.text();
      throw new Error(`Finalize failed: ${resp.status} - ${err}`);
    }

    logStep('Session set to Finalizing', 'success');
    res.json({ success: true, state: pocState });

  } catch (err) {
    pocState.status = 'error';
    logStep(err.message, 'error');
    res.status(500).json({ error: err.message });
  }
});

// -----------------------------------------------------------------------------
// STEP 10 — Create + Poll Snapshot, Download merged PDF
// -----------------------------------------------------------------------------
app.post('/poc/snapshot', async (req, res) => {
  try {
    ensureConfigured(CLIENT_ID, 'BB_CLIENT_ID');

    if (!pocState.sessionId || pocState.sessionFileIds.length === 0)
      throw new Error('No active session or no files in session');

    pocState.status = 'snapshotting';

    const accessToken = await tokenManager.getValidAccessToken();
    const downloads   = [];

    for (const sf of pocState.sessionFileIds) {
      logStep(`Requesting snapshot for "${sf.name}" (sessionFileId=${sf.sessionFileId})...`, 'info');

      const snapResp = await fetch(
        `${API_V1}/sessions/${pocState.sessionId}/files/${sf.sessionFileId}/snapshot`,
        {
          method: 'POST',
          headers: { Authorization: `Bearer ${accessToken}`, client_id: CLIENT_ID }
        }
      );

      if (!snapResp.ok) {
        const err = await snapResp.text();
        logStep(`Snapshot request failed for "${sf.name}": ${snapResp.status} - ${err}`, 'warn');
        continue;
      }

      logStep(`Polling snapshot for "${sf.name}"...`, 'info');

      const maxAttempts  = 20;
      const pollInterval = 5000;
      let attempts    = 0;
      let downloadUrl = null;

      while (attempts < maxAttempts) {
        await new Promise(resolve => setTimeout(resolve, pollInterval));
        attempts++;

        const pollToken = await tokenManager.getValidAccessToken();
        const pollResp  = await fetch(
          `${API_V1}/sessions/${pocState.sessionId}/files/${sf.sessionFileId}/snapshot`,
          { headers: { Authorization: `Bearer ${pollToken}`, client_id: CLIENT_ID } }
        );

        if (!pollResp.ok) {
          logStep(`Snapshot poll ${attempts} returned ${pollResp.status}`, 'warn');
          continue;
        }

        const pollData = await pollResp.json();
        logStep(`Snapshot poll ${attempts}/${maxAttempts}: ${pollData.Status}`, 'info');

        if (pollData.Status === 'Complete') {
          downloadUrl = pollData.DownloadUrl;
          logStep(`Snapshot ready for "${sf.name}"`, 'success');
          break;
        }

        if (pollData.Status === 'Error')
          throw new Error(`Snapshot error for "${sf.name}": ${pollData.Message || 'unknown'}`);
      }

      if (!downloadUrl) {
        logStep(`Snapshot timed out for "${sf.name}"`, 'warn');
        continue;
      }

      const dlResp = await fetch(downloadUrl);
      if (!dlResp.ok) throw new Error(`Download failed for "${sf.name}": ${dlResp.status}`);

      const pdfBuffer = await dlResp.buffer();

      const publicDir = path.join(__dirname, 'public');
      if (!fs.existsSync(publicDir)) fs.mkdirSync(publicDir, { recursive: true });

      const outputFileName = `${demoStub.documentId}_${sf.name.replace(/\.[^.]+$/, '')}_Reviewed.pdf`;
      fs.writeFileSync(path.join(publicDir, outputFileName), pdfBuffer);

      logStep(`PDF saved: ${outputFileName} (${pdfBuffer.length} bytes)`, 'success');
      downloads.push({ name: outputFileName, path: `/${outputFileName}`, size: pdfBuffer.length });
    }

    pocState.status = 'complete';
    logStep('All snapshots complete — reviewed documents ready', 'success');

    res.json({ success: true, downloads, state: pocState });

  } catch (err) {
    pocState.status = 'error';
    logStep(err.message, 'error');
    res.status(500).json({ error: err.message });
  }
});

// -----------------------------------------------------------------------------
// STEP 11 — Cleanup
// -----------------------------------------------------------------------------
app.post('/poc/cleanup', async (req, res) => {
  try {
    ensureConfigured(CLIENT_ID, 'BB_CLIENT_ID');

    if (!pocState.sessionId) throw new Error('No active session to clean up');

    const accessToken = await tokenManager.getValidAccessToken();

    if (pocState.subscriptionId) {
      logStep(`Deleting webhook subscription ${pocState.subscriptionId}...`, 'info');
      const subResp = await fetch(`${API_V2}/subscriptions/${pocState.subscriptionId}`, {
        method: 'DELETE',
        headers: { Authorization: `Bearer ${accessToken}`, client_id: CLIENT_ID }
      });
      logStep(
        subResp.ok ? 'Webhook subscription deleted' : `Subscription delete returned ${subResp.status}`,
        subResp.ok ? 'success' : 'warn'
      );
    }

    logStep(`Deleting session ${pocState.sessionId}...`, 'info');
    const sessResp = await fetch(`${API_V1}/sessions/${pocState.sessionId}`, {
      method: 'DELETE',
      headers: { Authorization: `Bearer ${accessToken}`, client_id: CLIENT_ID }
    });
    logStep(
      sessResp.ok ? 'Session deleted' : `Session delete returned ${sessResp.status}`,
      sessResp.ok ? 'success' : 'warn'
    );

    logStep('Cleanup complete', 'success');
    pocState.sessionId      = null;
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
  const payload = req.body || {};
  logStep(
    `Webhook: ResourceType=${payload.ResourceType || 'unknown'}, EventType=${payload.EventType || 'unknown'}`,
    'webhook'
  );
  pocState.webhookEvents.push({ ...payload, receivedAt: new Date().toISOString() });
  if (payload.ResourceType === 'Sessions' && payload.EventType === 'Update')
    logStep('Session update event — middleware could trigger next workflow step', 'webhook');
  res.sendStatus(200);
});

// =============================================================================
// STANDALONE ENDPOINTS
// =============================================================================

app.get('/powerbi/markups', async (req, res) => {
  try {
    ensureConfigured(CLIENT_ID,         'BB_CLIENT_ID');
    ensureConfigured(MARKUP_SESSION_ID, 'MARKUP_SESSION_ID');
    ensureConfigured(MARKUP_FILE_ID,    'MARKUP_FILE_ID');

    const accessToken = await tokenManager.getValidAccessToken();
    const response    = await fetch(
      `${API_V2}/sessions/${MARKUP_SESSION_ID}/files/${MARKUP_FILE_ID}/markups`,
      { headers: { Authorization: `Bearer ${accessToken}`, client_id: CLIENT_ID, Accept: 'application/json' } }
    );

    if (!response.ok) throw new Error(`Failed to get markups: ${response.status}`);

    const data    = await response.json();
    const markups = data.Markups || data || [];

    res.json(markups.map(m => ({
      MarkupId: m.Id, FileName: MARKUP_FILE_NAME, FileId: MARKUP_FILE_ID, SessionId: MARKUP_SESSION_ID,
      Type: m.Type, Subject: m.Subject, Comment: m.Comment, Author: m.Author,
      DateCreated: m.DateCreated, DateModified: m.DateModified, Page: m.Page,
      Status: m.Status, Color: m.Color, Layer: m.Layer
    })));
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.get('/api/project-markups', (req, res) => {
  if (pocState.markups.length === 0)
    return res.status(404).json({ error: 'No markup data. Run /poc/run-markuplist-job first.' });

  res.json(pocState.markups.map(m => ({
    MarkupId: m.Id, Author: m.Author, Type: m.Type, Subject: m.Subject,
    Comment: m.Comment, Status: m.Status, Layer: m.Layer, Page: m.Page,
    DateCreated: m.DateCreated, DateModified: m.DateModified, Color: m.Color,
    Checked: m.Checked, Locked: m.Locked, ExtendedProperties: m.ExtendedProperties || {},
    SourceFile: m._sourceFile
  })));
});

// -----------------------------------------------------------------------------
// START SERVER
// -----------------------------------------------------------------------------
app.listen(PORT, () => {
  console.log(`\nBluebeam Integration PoC  →  http://localhost:${PORT}`);
  console.log(`Studio Project: ${POC_PROJECT_ID}`);
  console.log(`Primary reviewer: dmolz@bluebeam.com`);
  if (isLocalhost(WEBHOOK_CALLBACK_URL))
    console.log(`\n⚠  WEBHOOK_CALLBACK_URL is localhost — webhook step will be skipped`);
  console.log(`\nROUNDTRIP FLOW:`);
  console.log(`   POST /poc/upload-to-project     — Step 0:  Upload PDF(s) from UI → project`);
  console.log(`   POST /poc/trigger               — Step 1:  Workflow event`);
  console.log(`   POST /poc/create-session        — Step 2:  Create Studio Session`);
  console.log(`   POST /poc/register-webhook      — Step 3:  Subscribe to session events`);
  console.log(`   POST /poc/add-to-session        — Step 4:  Check project file(s) into session`);
  console.log(`   POST /poc/invite-reviewers      — Step 5:  Invite reviewers`);
  console.log(`        (Step 6: Review in Bluebeam Revu — no API call)`);
  console.log(`   POST /poc/update-project-copy   — Step 7:  Push session markups → project`);
  console.log(`   POST /poc/run-markuplist-job    — Step 8:  Extract markup metadata`);
  console.log(`   POST /poc/finalize              — Step 9:  Finalize session`);
  console.log(`   POST /poc/snapshot              — Step 10: Snapshot + download PDF`);
  console.log(`   POST /poc/cleanup               — Step 11: Delete webhook + session\n`);
});
