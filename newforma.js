// Direct push of Studio markups into Newforma Konekt via its BCF REST API
// (buildingSMART BCF 2.1: POST .../topics then .../topics/{guid}/comments).
// Reuses the same markupToIssue mapping as the .bcfzip export so the file and
// the API can never drift apart. Config comes from env vars; the token never
// lives in code.
//
//   NEWFORMA_BCF_BASE       e.g. https://bcfrestapi.bimtrackapp.co
//   NEWFORMA_BCF_VERSION    default "2.1"
//   NEWFORMA_BCF_PROJECT_ID the target project's id (from the Integrations page)
//   NEWFORMA_BCF_TOKEN      OAuth2 bearer token (from the Integrations page)
const { markupToIssue } = require('./bcf');

function newformaConfig() {
  return {
    base: (process.env.NEWFORMA_BCF_BASE || '').replace(/\/+$/, ''),
    version: process.env.NEWFORMA_BCF_VERSION || '2.1',
    projectId: process.env.NEWFORMA_BCF_PROJECT_ID || '',
    token: process.env.NEWFORMA_BCF_TOKEN || '',
  };
}
function newformaConfigured() {
  const c = newformaConfig();
  return !!(c.base && c.projectId && c.token);
}

// Remember what we already pushed this process, per session, so an auto-sync
// that re-runs only sends NEW markups instead of duplicating topics. In-memory
// only (resets on redeploy) — fine for the PoC; durable state would live in DB.
const pushedBySession = new Map(); // sid -> Set(issue.key)

async function postJson(url, token, body) {
  const res = await fetch(url, {
    method: 'POST',
    headers: {
      'Authorization': 'Bearer ' + token,
      'Content-Type': 'application/json',
      'Accept': 'application/json',
    },
    body: JSON.stringify(body),
  });
  const text = await res.text();
  let json = null; if (text) { try { json = JSON.parse(text); } catch { json = text; } }
  return { ok: res.ok, status: res.status, json };
}

// Push a session's markups into Newforma. Returns a summary.
async function pushSessionToNewforma(sid, session, file, markups) {
  const c = newformaConfig();
  if (!c.base || !c.projectId || !c.token) {
    throw new Error('Newforma is not configured. Set NEWFORMA_BCF_BASE, NEWFORMA_BCF_PROJECT_ID, and NEWFORMA_BCF_TOKEN.');
  }
  const topicsUrl = `${c.base}/bcf/${c.version}/projects/${encodeURIComponent(c.projectId)}/topics`;
  const seen = pushedBySession.get(sid) || new Set();
  const results = [];
  let pushed = 0, skipped = 0, failed = 0;

  for (const m of (markups || [])) {
    const issue = markupToIssue(m, session, file);
    if (seen.has(issue.key)) { skipped++; results.push({ key: issue.key, title: issue.title, status: 'skipped (already synced)' }); continue; }

    const topicBody = {
      topic_type: issue.topicType,
      topic_status: issue.topicStatus,
      title: issue.title,
      labels: issue.labels,
      description: issue.description,
    };
    const t = await postJson(topicsUrl, c.token, topicBody);
    if (!t.ok) {
      failed++;
      results.push({ key: issue.key, title: issue.title, status: 'topic failed (' + t.status + ')', detail: t.json });
      continue;
    }
    const guid = t.json && (t.json.guid || t.json.Guid);
    // Best-effort comment; a topic without its comment is still a success.
    if (guid && issue.comment) {
      const commentUrl = `${c.base}/bcf/${c.version}/projects/${encodeURIComponent(c.projectId)}/topics/${encodeURIComponent(guid)}/comments`;
      const cRes = await postJson(commentUrl, c.token, { comment: issue.comment });
      results.push({ key: issue.key, title: issue.title, status: 'pushed', guid, comment: cRes.ok ? 'added' : ('failed (' + cRes.status + ')') });
    } else {
      results.push({ key: issue.key, title: issue.title, status: 'pushed', guid: guid || null, comment: issue.comment ? 'skipped (no topic guid returned)' : 'none' });
    }
    seen.add(issue.key);
    pushed++;
  }
  pushedBySession.set(sid, seen);
  return { ok: failed === 0, pushed, skipped, failed, total: (markups || []).length, projectId: c.projectId, results };
}

module.exports = { pushSessionToNewforma, newformaConfigured, newformaConfig };
