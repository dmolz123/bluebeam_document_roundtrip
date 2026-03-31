/**
 * Proof-of-concept reference implementation for a Bluebeam Studio roundtrip workflow.
 * Intended for evaluation and development reference only.
 *
 * Roundtrip flow:
 *   1.  /poc/trigger               — Simulate source-system workflow event
 *   2.  /poc/create-session        — Create a Studio Session
 *   3.  /poc/register-webhook      — Subscribe to session events
 *   4.  /poc/upload-file           — Upload PDF to session (3-step)
 *   5.  /poc/invite-reviewers      — Invite reviewers to session
 *   6.  (Review happens in Bluebeam Revu — no API step)
 *   7.  /poc/update-project-copy   — Push session markups back to project file
 *   8.  /poc/run-markuplist-job    — Run markuplist job on project file, poll, return markups
 *   9.  /poc/finalize              — Set session status to Finalizing
 *   10. /poc/snapshot              — Create + poll snapshot, download merged PDF
 *   11. /poc/cleanup               — Delete webhook subscription + session
 */

require('dotenv').config();
const express = require('express');
const cors    = require('cors');
const fs      = require('fs');
const path    = require('path');
const TokenManager = require('./tokenManager');

const app  = express();
const PORT = process.env.PORT || 3000;

// -----------------------------------------------------------------------------
// API BASE URLs
// -----------------------------------------------------------------------------
const API_V1 = 'https://api.bluebeam.com/publicapi/v1';
const API_V2 = 'https://api.bluebeam.com/publicapi/v2';

const CLIENT_ID            = process.env.BB_CLIENT_ID;
const WEBHOOK_CALLBACK_URL =
  process.env.WEBHOOK_CALLBACK_URL ||
  `http://localhost:${PORT}/webhook/studio-events`;

// -----------------------------------------------------------------------------
// OPTIONAL SAMPLE / REFERENCE CONSTANTS
// Populate in .env to enable the standalone /powerbi/markups endpoint.
// -----------------------------------------------------------------------------
const MARKUP_SESSION_ID = process.env.MARKUP_SESSION_ID || '';
const MARKUP_FILE_ID    = process.env.MARKUP_FILE_ID    || '';
const MARKUP_FILE_NAME  = process.env.MARKUP_FILE_NAME  || 'Sample Drawing.pdf';

// Project-level identifiers used by Steps 7-8 (update project copy + markuplist job).
const BB_PROJECT_ID      = process.env.BB_PROJECT_ID      || '';
const BB_PROJECT_FILE_ID = process.env.BB_PROJECT_FILE_ID || '';

// -----------------------------------------------------------------------------
// DEMO STUB — generic source-system payload
// Simulates what an upstream integration (DMS, PLM, ERP, etc.) would provide.
// Fields can be overridden at runtime via POST /poc/configure.
// -----------------------------------------------------------------------------
const DEMO_ASSETS_PATH = process.env.DEMO_ASSETS_PATH || './demo-assets';

let demoStub = {
  documentId:   process.env.DEMO_DOCUMENT_ID   || 'DOC-001',
  documentName: process.env.DEMO_DOCUMENT_NAME || 'Sample-Drawing.pdf',
  description:  process.env.DEMO_DESCRIPTION   || 'Drawing review — coordination update',
  reviewers: [
    {
      email:            process.env.DEMO_REVIEWER_EMAIL || 'reviewer@example.com',
      hasStudioAccount: false
    }
  ],
  sessionEndDate: new Date(Date.now() + 7 * 24 * 60 * 60 * 1000).toISOString()
};

// -----------------------------------------------------------------------------
// IN-MEMORY STATE — tracks active PoC session
// -----------------------------------------------------------------------------
let pocState = {
  sessionId:      null,
  subscriptionId: null,
  fileIds:        [],      // [{ fileId, name, source }]
  markups:        [],      // populated by run-markuplist-job
  markupJobId:    null,
  status:         'idle',  // idle | triggered | creating | uploading | inviting | active
                           // | updating-project | extracting-markups | finalizing
                           // | snapshotting | complete | error
  log:            [],
  createdAt:      null,
  webhookEvents:  []
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
    fileIds:        [],
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

function fileExists(filePath) {
  try { return fs.existsSync(filePath); } catch { return false; }
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

// -----------------------------------------------------------------------------
// HEALTH CHECK
// -----------------------------------------------------------------------------
app.get('/health', (req, res) => {
  res.json({
    status: 'healthy',
    config: {
      hasClientId:        Boolean(CLIENT_ID),
      webhookCallbackUrl: WEBHOOK_CALLBACK_URL
    },
    projectConfig: {
      hasProjectId:     Boolean(BB_PROJECT_ID),
      hasProjectFileId: Boolean(BB_PROJECT_FILE_ID)
    },
    sampleConfig: {
      hasMarkupSessionId: Boolean(MARKUP_SESSION_ID),
      hasMarkupFileId:    Boolean(MARKUP_FILE_ID)
    }
  });
});

// =============================================================================
// POC ROUTES
// =============================================================================

// GET current PoC state (for UI polling)
app.get('/poc/state', (req, res) => {
  res.json({ ...pocState, stub: demoStub });
});

// GET current stub config
app.get('/poc/stub', (req, res) => {
  res.json(demoStub);
});

// POST configure stub fields from the UI
// Body: { documentId?, documentName?, description?, reviewerEmail?, hasStudioAccount? }
app.post('/poc/configure', (req, res) => {
  const { documentId, documentName, description, reviewerEmail, hasStudioAccount } = req.body || {};

  if (documentId)   demoStub.documentId   = documentId;
  if (documentName) demoStub.documentName = documentName;
  if (description)  demoStub.description  = description;
  if (reviewerEmail) {
    demoStub.reviewers = [{
      email:            reviewerEmail,
      hasStudioAccount: Boolean(hasStudioAccount)
    }];
  }

  demoStub.sessionEndDate = new Date(Date.now() + 7 * 24 * 60 * 60 * 1000).toISOString();

  logStep(`Stub reconfigured: doc=${demoStub.documentName}, reviewer=${demoStub.reviewers[0].email}`, 'info');
  res.json({ success: true, stub: demoStub });
});

// Reset PoC state
app.post('/poc/reset', (req, res) => {
  resetPocState();
  logStep('PoC state reset', 'info');
  res.json({ success: true });
});

// -----------------------------------------------------------------------------
// STEP 1 — Simulate source-system workflow event
// -----------------------------------------------------------------------------
app.post('/poc/trigger', (req, res) => {
  pocState.status = 'triggered';
  pocState.log    = [];

  logStep(`Workflow event received — document: ${demoStub.documentId} / ${demoStub.documentName}`, 'info');
  logStep(`Description: ${demoStub.description}`, 'info');
  logStep(`Reviewers resolved: ${demoStub.reviewers.map(r => r.email).join(', ')}`, 'info');
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
    const sessionName = `${demoStub.documentId}_${demoStub.documentName.replace('.pdf', '')}_Review`;

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
// -----------------------------------------------------------------------------
app.post('/poc/register-webhook', async (req, res) => {
  try {
    ensureConfigured(CLIENT_ID, 'BB_CLIENT_ID');

    if (!pocState.sessionId)
      throw new Error('No active session — run create-session first');

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
// STEP 4 — Upload PDF (3-step: metadata → S3 → confirm)
// -----------------------------------------------------------------------------
app.post('/poc/upload-file', async (req, res) => {
  try {
    ensureConfigured(CLIENT_ID, 'BB_CLIENT_ID');

    if (!pocState.sessionId) throw new Error('No active session');

    pocState.status = 'uploading';

    const docName = demoStub.documentName;
    const docPath = path.join(DEMO_ASSETS_PATH, docName);

    logStep(`Uploading ${docName} to session ${pocState.sessionId}...`, 'info');

    const accessToken = await tokenManager.getValidAccessToken();

    // 4a — Create metadata block
    logStep('Step 4a: Creating file metadata block...', 'info');

    const metaResp = await fetch(`${API_V1}/sessions/${pocState.sessionId}/files`, {
      method: 'POST',
      headers: {
        Authorization:  `Bearer ${accessToken}`,
        client_id:      CLIENT_ID,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        Name:   docName,
        Source: `source-system://${demoStub.documentId}/${docName}`
      })
    });

    if (!metaResp.ok) {
      const err = await metaResp.text();
      throw new Error(`Metadata block failed: ${metaResp.status} - ${err}`);
    }

    const metaData          = await metaResp.json();
    const fileId            = metaData.Id;
    const uploadUrl         = metaData.UploadUrl;
    const uploadContentType = metaData.UploadContentType || 'application/pdf';

    logStep(`Metadata block created: fileId=${fileId}`, 'success');

    // 4b — Upload binary to storage
    logStep('Step 4b: Uploading PDF binary to storage...', 'info');

    let pdfBuffer;
    if (fileExists(docPath)) {
      pdfBuffer = fs.readFileSync(docPath);
      logStep(`Loaded from disk: ${docPath}`, 'info');
    } else {
      pdfBuffer = Buffer.from(
        '%PDF-1.4\n1 0 obj<</Type/Catalog/Pages 2 0 R>>endobj\n' +
        '2 0 obj<</Type/Pages/Kids[3 0 R]/Count 1>>endobj\n' +
        '3 0 obj<</Type/Page/MediaBox[0 0 612 792]/Parent 2 0 R>>endobj\n' +
        'xref\n0 4\n0000000000 65535 f\n0000000009 00000 n\n' +
        '0000000058 00000 n\n0000000115 00000 n\n' +
        'trailer<</Size 4/Root 1 0 R>>\nstartxref\n190\n%%EOF'
      );
      logStep(`No file at ${docPath} — using minimal demo PDF`, 'info');
    }

    const s3Resp = await fetch(uploadUrl, {
      method: 'PUT',
      headers: {
        'x-amz-server-side-encryption': 'AES256',
        'Content-Type':                  uploadContentType
      },
      body: pdfBuffer
    });

    if (!s3Resp.ok) throw new Error(`Binary upload failed: ${s3Resp.status}`);
    logStep(`Binary upload complete (${pdfBuffer.length} bytes)`, 'success');

    // 4c — Confirm upload
    logStep('Step 4c: Confirming upload with Bluebeam...', 'info');

    const confirmResp = await fetch(
      `${API_V1}/sessions/${pocState.sessionId}/files/${fileId}/confirm-upload`,
      {
        method: 'POST',
        headers: { Authorization: `Bearer ${accessToken}`, client_id: CLIENT_ID }
      }
    );

    if (!confirmResp.ok) {
      const err = await confirmResp.text();
      throw new Error(`Confirm upload failed: ${confirmResp.status} - ${err}`);
    }

    pocState.fileIds.push({
      fileId,
      name:   docName,
      source: `source-system://${demoStub.documentId}/${docName}`
    });

    logStep(`File confirmed in session: ${docName} (fileId=${fileId})`, 'success');
    res.json({ success: true, fileId, state: pocState });

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
    ensureConfigured(CLIENT_ID, 'BB_CLIENT_ID');

    if (!pocState.sessionId) throw new Error('No active session');

    pocState.status = 'inviting';
    logStep(`Inviting ${demoStub.reviewers.length} reviewer(s)...`, 'info');

    const accessToken = await tokenManager.getValidAccessToken();
    const results     = [];

    for (const reviewer of demoStub.reviewers) {
      // /invite sends email — best for users without a Studio account
      // /users  direct-adds known existing Studio accounts
      const endpoint = reviewer.hasStudioAccount
        ? `${API_V1}/sessions/${pocState.sessionId}/users`
        : `${API_V1}/sessions/${pocState.sessionId}/invite`;

      logStep(
        `Inviting ${reviewer.email} via ${reviewer.hasStudioAccount ? 'direct-add' : 'email-invite'}`,
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
          Message:   `Please review document ${demoStub.documentId}: ${demoStub.description}`
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
    logStep(`Join via Bluebeam Revu using session ID: ${pocState.sessionId}`, 'info');
    res.json({ success: true, results, state: pocState });

  } catch (err) {
    pocState.status = 'error';
    logStep(err.message, 'error');
    res.status(500).json({ error: err.message });
  }
});

// -----------------------------------------------------------------------------
// STEP 6 — Non-API: Review in Bluebeam Revu. Poll /poc/state to monitor.
// -----------------------------------------------------------------------------

// -----------------------------------------------------------------------------
// STEP 7 — Update the Project File Copy
//
// Pushes session markups back to the linked project file without closing the
// session. Equivalent to "Update Server Copy" in Bluebeam Revu.
//
// IMPORTANT: {fileId} here is the SESSION file ID (pocState.fileIds[0].fileId),
// NOT the project file ID. These are different identifiers.
//
// Prerequisites:
//   - BB_PROJECT_ID and BB_PROJECT_FILE_ID set in .env
//   - Session must still be active (not finalized or deleted)
// -----------------------------------------------------------------------------
app.post('/poc/update-project-copy', async (req, res) => {
  try {
    ensureConfigured(CLIENT_ID,          'BB_CLIENT_ID');
    ensureConfigured(BB_PROJECT_ID,      'BB_PROJECT_ID');
    ensureConfigured(BB_PROJECT_FILE_ID, 'BB_PROJECT_FILE_ID');

    if (!pocState.sessionId)           throw new Error('No active session');
    if (pocState.fileIds.length === 0) throw new Error('No files uploaded to session');

    pocState.status = 'updating-project';

    const { fileId, name } = pocState.fileIds[0];

    logStep(`Updating project file copy from session file: ${name} (sessionFileId=${fileId})`, 'info');
    logStep(`Target: project=${BB_PROJECT_ID} / projectFile=${BB_PROJECT_FILE_ID}`, 'info');

    const accessToken = await tokenManager.getValidAccessToken();

    // POST /sessions/{sessionId}/files/{sessionFileId}/updateprojectcopy
    // Non-destructive — session remains active after this call.
    const resp = await fetch(
      `${API_V1}/sessions/${pocState.sessionId}/files/${fileId}/updateprojectcopy`,
      {
        method: 'POST',
        headers: { Authorization: `Bearer ${accessToken}`, client_id: CLIENT_ID }
      }
    );

    if (resp.status === 204) {
      logStep('Project file copy updated (204 No Content)', 'success');
      logStep('Session remains active — proceed to markup extraction', 'info');
      res.json({ success: true, state: pocState });
    } else {
      const err = await resp.text();
      throw new Error(`Update project copy failed: ${resp.status} - ${err}`);
    }

  } catch (err) {
    pocState.status = 'error';
    logStep(err.message, 'error');
    res.status(500).json({ error: err.message });
  }
});

// -----------------------------------------------------------------------------
// STEP 8 — Run Markup List Job on the Project File
//
// Submits a markuplist job against the updated project file, polls until
// complete, and stores the full markup array in pocState.markups.
//
// Flow:
//   POST .../jobs/markuplist         → { JobId }
//   GET  .../jobs/markuplist/{jobId} → poll until JobStatus === 2 (Complete)
//
// Prerequisites: BB_PROJECT_ID + BB_PROJECT_FILE_ID in .env,
// update-project-copy must have run successfully first.
// -----------------------------------------------------------------------------
app.post('/poc/run-markuplist-job', async (req, res) => {
  try {
    ensureConfigured(CLIENT_ID,          'BB_CLIENT_ID');
    ensureConfigured(BB_PROJECT_ID,      'BB_PROJECT_ID');
    ensureConfigured(BB_PROJECT_FILE_ID, 'BB_PROJECT_FILE_ID');

    pocState.status = 'extracting-markups';
    logStep(`Submitting markuplist job: project=${BB_PROJECT_ID} / file=${BB_PROJECT_FILE_ID}`, 'info');

    const accessToken = await tokenManager.getValidAccessToken();

    const authHeaders = {
      Authorization:  `Bearer ${accessToken}`,
      client_id:      CLIENT_ID,
      'Content-Type': 'application/json',
      Accept:         'application/json'
    };

    // 8a — Submit job
    const submitResp = await fetch(
      `${API_V1}/projects/${BB_PROJECT_ID}/files/${BB_PROJECT_FILE_ID}/jobs/markuplist`,
      { method: 'POST', headers: authHeaders, body: JSON.stringify({}) }
    );

    if (!submitResp.ok) {
      const err = await submitResp.text();
      throw new Error(`Markuplist job submission failed: ${submitResp.status} - ${err}`);
    }

    const { JobId }      = await submitResp.json();
    pocState.markupJobId = JobId;

    logStep(`Markuplist job submitted: JobId=${JobId}`, 'success');
    logStep('Polling for completion...', 'info');

    // 8b — Poll until complete
    const pollUrl = `${API_V1}/projects/${BB_PROJECT_ID}/files/${BB_PROJECT_FILE_ID}/jobs/markuplist/${JobId}`;
    const result  = await pollJob(pollUrl, authHeaders);

    pocState.markups = result.Markups || [];
    pocState.status  = 'active'; // session still open

    logStep(`Markuplist complete — ${pocState.markups.length} markup(s) extracted`, 'success');

    res.json({
      success: true,
      jobId:   JobId,
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

    logStep('Session set to Finalizing — Bluebeam will process the close', 'success');
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

    if (!pocState.sessionId || pocState.fileIds.length === 0)
      throw new Error('No active session or no files uploaded');

    pocState.status = 'snapshotting';

    const { fileId, name } = pocState.fileIds[0];
    logStep(`Requesting snapshot for ${name} (fileId=${fileId})...`, 'info');

    const accessToken = await tokenManager.getValidAccessToken();

    const snapResp = await fetch(
      `${API_V1}/sessions/${pocState.sessionId}/files/${fileId}/snapshot`,
      {
        method: 'POST',
        headers: { Authorization: `Bearer ${accessToken}`, client_id: CLIENT_ID }
      }
    );

    if (!snapResp.ok) {
      const err = await snapResp.text();
      throw new Error(`Snapshot request failed: ${snapResp.status} - ${err}`);
    }

    logStep('Snapshot requested — polling for completion...', 'info');

    // Snapshot uses string Status ("Complete"/"Error"), not numeric JobStatus codes
    const maxAttempts  = 20;
    const pollInterval = 5000;
    let attempts    = 0;
    let downloadUrl = null;

    while (attempts < maxAttempts) {
      await new Promise(resolve => setTimeout(resolve, pollInterval));
      attempts++;

      const pollToken = await tokenManager.getValidAccessToken();
      const pollResp  = await fetch(
        `${API_V1}/sessions/${pocState.sessionId}/files/${fileId}/snapshot`,
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
        logStep('Snapshot complete — download URL received', 'success');
        break;
      }

      if (pollData.Status === 'Error')
        throw new Error(`Snapshot error: ${pollData.Message || 'unknown'}`);
    }

    if (!downloadUrl)
      throw new Error(`Snapshot did not complete after ${maxAttempts} attempts`);

    logStep('Downloading marked-up PDF...', 'info');

    const dlResp = await fetch(downloadUrl);
    if (!dlResp.ok) throw new Error(`Download failed: ${dlResp.status}`);

    const pdfBuffer = await dlResp.buffer();

    const publicDir = path.join(__dirname, 'public');
    if (!fs.existsSync(publicDir)) fs.mkdirSync(publicDir, { recursive: true });

    const outputFileName = `${demoStub.documentId}_Reviewed.pdf`;
    fs.writeFileSync(path.join(publicDir, outputFileName), pdfBuffer);

    pocState.status = 'complete';
    logStep(`PDF saved: ${outputFileName} (${pdfBuffer.length} bytes)`, 'success');
    logStep('Reviewed document ready for return to source system', 'info');

    res.json({
      success:      true,
      downloadPath: `/${outputFileName}`,
      fileSize:     pdfBuffer.length,
      state:        pocState
    });

  } catch (err) {
    pocState.status = 'error';
    logStep(err.message, 'error');
    res.status(500).json({ error: err.message });
  }
});

// -----------------------------------------------------------------------------
// STEP 11 — Cleanup (delete webhook subscription + session)
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
// WEBHOOK LISTENER — receives Bluebeam Studio events
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

// GET markups from a live session file (requires MARKUP_* env vars)
app.get('/powerbi/markups', async (req, res) => {
  try {
    ensureConfigured(CLIENT_ID,         'BB_CLIENT_ID');
    ensureConfigured(MARKUP_SESSION_ID, 'MARKUP_SESSION_ID');
    ensureConfigured(MARKUP_FILE_ID,    'MARKUP_FILE_ID');

    const accessToken = await tokenManager.getValidAccessToken();

    const response = await fetch(
      `${API_V2}/sessions/${MARKUP_SESSION_ID}/files/${MARKUP_FILE_ID}/markups`,
      {
        headers: {
          Authorization: `Bearer ${accessToken}`,
          client_id:     CLIENT_ID,
          Accept:        'application/json'
        }
      }
    );

    if (!response.ok) {
      const errText = await response.text();
      throw new Error(`Failed to get markups: ${response.status} - ${errText}`);
    }

    const data    = await response.json();
    const markups = data.Markups || data || [];

    const flattened = markups.map(m => ({
      MarkupId:     m.Id           || null,
      FileName:     MARKUP_FILE_NAME,
      FileId:       MARKUP_FILE_ID,
      SessionId:    MARKUP_SESSION_ID,
      Type:         m.Type         || null,
      Subject:      m.Subject      || null,
      Comment:      m.Comment      || null,
      Author:       m.Author       || null,
      DateCreated:  m.DateCreated  || null,
      DateModified: m.DateModified || null,
      Page:         m.Page         || null,
      Status:       m.Status       || null,
      Color:        m.Color        || null,
      Layer:        m.Layer        || null
    }));

    res.json(flattened);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// GET markup data from the most recent markuplist job (served from memory)
app.get('/api/project-markups', (req, res) => {
  if (pocState.markups.length === 0) {
    return res.status(404).json({
      error: 'No markup data available. Run /poc/run-markuplist-job first.'
    });
  }

  const normalized = pocState.markups.map(m => ({
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
    ExtendedProperties: m.ExtendedProperties || {}
  }));

  res.json(normalized);
});

// -----------------------------------------------------------------------------
// START SERVER
// -----------------------------------------------------------------------------
app.listen(PORT, () => {
  console.log(`\nBluebeam Integration PoC  →  http://localhost:${PORT}`);

  console.log(`\nSTATUS / CONFIG:`);
  console.log(`   GET  /health`);
  console.log(`   GET  /poc/state`);
  console.log(`   GET  /poc/stub`);
  console.log(`   POST /poc/configure`);
  console.log(`   POST /poc/reset`);

  console.log(`\nROUNDTRIP FLOW:`);
  console.log(`   POST /poc/trigger                — Step 1:  Workflow event received`);
  console.log(`   POST /poc/create-session         — Step 2:  Create Studio Session`);
  console.log(`   POST /poc/register-webhook       — Step 3:  Subscribe to session events`);
  console.log(`   POST /poc/upload-file            — Step 4:  Upload PDF (3-step)`);
  console.log(`   POST /poc/invite-reviewers       — Step 5:  Invite reviewers`);
  console.log(`        (Step 6: Review in Bluebeam Revu — no API call)`);
  console.log(`   POST /poc/update-project-copy    — Step 7:  Push session markups → project file`);
  console.log(`   POST /poc/run-markuplist-job     — Step 8:  Extract markup metadata`);
  console.log(`   POST /poc/finalize               — Step 9:  Set session to Finalizing`);
  console.log(`   POST /poc/snapshot               — Step 10: Create + download merged PDF`);
  console.log(`   POST /poc/cleanup                — Step 11: Delete webhook + session`);
  console.log(`   POST /webhook/studio-events      —          Webhook receiver`);

  console.log(`\nSTANDALONE ENDPOINTS:`);
  console.log(`   GET  /powerbi/markups            — Live session markups (needs MARKUP_* env vars)`);
  console.log(`   GET  /api/project-markups        — Markups from last markuplist job\n`);
});
