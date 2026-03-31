# Bluebeam Studio — Document Roundtrip PoC

## Disclaimer

This repository contains a proof-of-concept integration demonstrating how an external system (DMS, PLM, ERP, or similar) can drive a complete document review lifecycle using the Bluebeam Studio API.

This project is **not an official Bluebeam product** and is not supported by Bluebeam. It is provided solely as a reference implementation for evaluation and development purposes.

---

## What This PoC Demonstrates

The full roundtrip lifecycle, driven entirely through the Bluebeam API:

| Step | What happens | API |
|------|-------------|-----|
| 1 | Workflow event received from source system (stubbed) | — |
| 2 | Create a Studio Session with permissions | `POST /v1/sessions` |
| 3 | Register a webhook subscription for session events | `POST /v2/subscriptions` |
| 4 | Upload a PDF to the session (3-step: metadata → S3 → confirm) | `POST /v1/sessions/{id}/files` |
| 5 | Invite reviewers via email invite or direct-add | `POST /v1/sessions/{id}/invite` |
| 6 | Reviewers mark up the document in Bluebeam Revu | *(non-API)* |
| 7 | Push session markups back to the project file | `POST /v1/sessions/{id}/files/{id}/updateprojectcopy` |
| 8 | Run a markup list job and extract all markup metadata | `POST /v1/projects/{id}/files/{id}/jobs/markuplist` |
| 9 | Finalize the session | `PUT /v1/sessions/{id}` |
| 10 | Create a snapshot and download the marked-up PDF | `POST /v1/sessions/{id}/files/{id}/snapshot` |
| 11 | Delete the webhook subscription and session | `DELETE /v2/subscriptions/{id}` |

The source system side of the workflow is **stubbed** — a configurable demo payload simulates what an upstream system (DMS, PLM, ERP, etc.) would provide. All Bluebeam API calls are real and require valid credentials.

---

## Architecture

```
Source System (stubbed)
        │
        │  workflow event
        ▼
  Express Backend (server.js)
  ├── OAuth 2.0 token management (tokenManager.js)
  ├── Session lifecycle (create → upload → invite → finalize → cleanup)
  ├── Project file roundtrip (updateprojectcopy → markuplist job)
  ├── Webhook receiver (/webhook/studio-events)
  └── Static file server → index.html (PoC UI)
        │
        │  HTTPS
        ▼
  Bluebeam Studio API
  ├── /publicapi/v1  — Sessions, Projects, File Jobs
  └── /publicapi/v2  — Webhook Subscriptions, Markups
```

The backend is a thin Express proxy. It handles token refresh automatically and exposes each roundtrip step as a discrete `POST /poc/*` endpoint. The frontend polls `/poc/state` every 3 seconds to keep the UI in sync.

---

## Requirements

- Node.js 18+
- Bluebeam API credentials (Client ID and Client Secret from the [Bluebeam Developer Portal](https://developers.bluebeam.com/))
- A Studio Project with a file already uploaded (required for Steps 7–8)
- A `.env` file configured as described below

---

## Quick Start

```bash
# 1. Clone the repository
git clone <repo-url>
cd <repo-directory>

# 2. Install dependencies
npm install

# 3. Configure environment
cp .env.example .env
# Edit .env with your credentials (see Environment Variables below)

# 4. Start the server
npm start
```

Open `http://localhost:3000` in your browser. Use the config strip at the top of the UI to set the document name, reviewer email, and description before running through the steps.

---

## Environment Variables

Copy `.env.example` to `.env` and fill in the values below.

### Required

| Variable | Description |
|----------|-------------|
| `BB_CLIENT_ID` | Your Bluebeam API Client ID |
| `BB_CLIENT_SECRET` | Your Bluebeam API Client Secret |

### Required for Steps 7–8 (project file roundtrip + markup extraction)

| Variable | Description |
|----------|-------------|
| `BB_PROJECT_ID` | Studio Project ID containing the source file |
| `BB_PROJECT_FILE_ID` | File ID within that project |

### Optional — Demo stub defaults

These set the initial values shown in the UI config strip. All can be overridden at runtime without restarting the server.

| Variable | Default | Description |
|----------|---------|-------------|
| `DEMO_DOCUMENT_ID` | `DOC-001` | Document identifier passed to Bluebeam session name and file source URI |
| `DEMO_DOCUMENT_NAME` | `Sample-Drawing.pdf` | Filename to use when uploading. Place the actual file in `./demo-assets/` |
| `DEMO_DESCRIPTION` | `Drawing review — coordination update` | Review description sent in reviewer invitation message |
| `DEMO_REVIEWER_EMAIL` | `reviewer@example.com` | Email address to invite to the session |
| `DEMO_ASSETS_PATH` | `./demo-assets` | Directory where the server looks for the source PDF |

### Optional — Webhook

| Variable | Default | Description |
|----------|---------|-------------|
| `WEBHOOK_CALLBACK_URL` | `http://localhost:{PORT}/webhook/studio-events` | Public URL Bluebeam will POST events to. For local testing use a tunnel such as [ngrok](https://ngrok.com/) |

### Optional — Standalone markups endpoint

Only needed if you use the standalone `/powerbi/markups` endpoint (not part of the main roundtrip flow).

| Variable | Description |
|----------|-------------|
| `MARKUP_SESSION_ID` | Session ID for the standalone markups endpoint |
| `MARKUP_FILE_ID` | File ID within that session |
| `MARKUP_FILE_NAME` | Display name for the file in markup output |

### Example `.env`

```env
# Required
BB_CLIENT_ID=your_client_id_here
BB_CLIENT_SECRET=your_client_secret_here

# Required for project roundtrip (Steps 7–8)
BB_PROJECT_ID=385-509-537
BB_PROJECT_FILE_ID=45231007

# Demo stub defaults (all overridable in the UI)
DEMO_DOCUMENT_ID=DOC-001
DEMO_DOCUMENT_NAME=Sample-Drawing.pdf
DEMO_DESCRIPTION=Drawing review — coordination update
DEMO_REVIEWER_EMAIL=reviewer@example.com
DEMO_ASSETS_PATH=./demo-assets

# Webhook (use ngrok or similar for local testing)
WEBHOOK_CALLBACK_URL=https://your-tunnel.ngrok.io/webhook/studio-events
```

---

## Source PDF

Place your test PDF in the `demo-assets/` directory with a filename matching `DEMO_DOCUMENT_NAME`:

```
demo-assets/
└── Sample-Drawing.pdf
```

If no file is found at that path, the server falls back to a minimal valid PDF for demo purposes. Real markup extraction in Step 8 requires a real PDF with markups added during the session review.

---

## API Endpoints

### Status

| Method | Path | Description |
|--------|------|-------------|
| `GET` | `/health` | Server health and configuration status |
| `GET` | `/poc/state` | Full current state (session IDs, markups, log, webhook events) |
| `GET` | `/poc/stub` | Current demo stub configuration |
| `POST` | `/poc/configure` | Update stub fields at runtime (no restart needed) |
| `POST` | `/poc/reset` | Reset all in-memory state |

### Roundtrip Flow

| Method | Path | Step |
|--------|------|------|
| `POST` | `/poc/trigger` | 1 — Simulate source-system event |
| `POST` | `/poc/create-session` | 2 — Create Studio Session |
| `POST` | `/poc/register-webhook` | 3 — Subscribe to session events |
| `POST` | `/poc/upload-file` | 4 — Upload PDF (3-step) |
| `POST` | `/poc/invite-reviewers` | 5 — Invite reviewers |
| `POST` | `/poc/update-project-copy` | 7 — Push session markups to project file |
| `POST` | `/poc/run-markuplist-job` | 8 — Run markup list job, poll, extract metadata |
| `POST` | `/poc/finalize` | 9 — Set session to Finalizing |
| `POST` | `/poc/snapshot` | 10 — Create snapshot, download marked-up PDF |
| `POST` | `/poc/cleanup` | 11 — Delete webhook subscription + session |
| `POST` | `/webhook/studio-events` | Webhook receiver (called by Bluebeam) |

### Standalone Endpoints

| Method | Path | Description |
|--------|------|-------------|
| `GET` | `/api/project-markups` | Returns markup metadata from the most recent markuplist job (in-memory) |
| `GET` | `/powerbi/markups` | Returns flattened markup list from a live session file (requires `MARKUP_*` env vars) |

---

## Project Structure

```
├── server.js          # Express backend — all roundtrip API logic
├── tokenManager.js    # OAuth 2.0 token management with auto-refresh
├── public/
│   └── index.html     # PoC UI — step runner, live log, markup metadata viewer
├── demo-assets/       # Place your test PDF here
├── .env               # Your credentials (git-ignored)
├── .env.example       # Template — copy to .env
└── package.json
```

---

## Notes on the Roundtrip Flow

**Step 4 (Upload)** uses a 3-step process required by the Bluebeam API: create a metadata placeholder, upload the binary directly to the returned S3 URL, then confirm the upload with Bluebeam. All three sub-steps are handled automatically by `/poc/upload-file`.

**Step 7 (Update Project Copy)** is a non-destructive operation — `updateprojectcopy` pushes session markups to the project file without closing the session. The session remains active and available for further work after this call. The `fileId` used here is the **session** file ID, not the project file ID.

**Step 8 (Markup List Job)** is asynchronous. The server submits the job, then polls the status endpoint every 3 seconds until `JobStatus` reaches `2` (Complete) or `3` (Error). Once complete, the full markup array — including author, type, comment, status, layer, position, and extended properties — is stored in memory and rendered in the UI.

**Step 6 (Review)** has no API equivalent. Reviewers open the session in Bluebeam Revu using the session ID, mark up the document, and when finished, you trigger Step 7 to pull their changes back to the project file.

**Webhook events** require a publicly accessible callback URL. For local testing, use a tunneling tool such as [ngrok](https://ngrok.com/) and set `WEBHOOK_CALLBACK_URL` in your `.env` accordingly.

---

## Adapting This PoC

The stub payload in `server.js` (`demoStub`) simulates what an upstream system would provide — document identifiers, reviewer lists, and session parameters. To adapt this PoC for a real integration, replace the stub with actual data from your source system (fetched via API, passed in a webhook payload, or read from a queue) and wire the session ID and job results back to that system after the roundtrip completes.

The backend proxy pattern (source system → Express → Bluebeam API) works with any upstream system that can make HTTP requests.

---

## See Also

- [Bluebeam Developer Portal](https://developers.bluebeam.com/)
- [Studio Session Guide](https://support.bluebeam.com/developer/studio-session-guide.html)
- [Authentication Guide](https://support.bluebeam.com/developer/authentication-guide.html)
- [Get Started in the Developer Portal](https://support.bluebeam.com/developer/getting-started-dev-portal.html)
