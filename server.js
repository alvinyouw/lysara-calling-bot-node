import express from "express";
import axios from "axios";
import { ClientSecretCredential } from "@azure/identity";

const app = express();
app.use(express.json({ limit: "2mb" }));
function requireApiKey(req, res, next) {
  const key = req.header("x-api-key");
  if (!process.env.API_KEY) {
    return res.status(500).json({ error: "API_KEY not configured on server" });
  }
  if (key !== process.env.API_KEY) {
    return res.status(401).json({ error: "Unauthorized" });
  }
  next();
}
const BUILD_TAG = "2026-02-22-remove-top-v1";

// ============================
// ENV (Azure App Service -> Environment variables)
// ============================
const TENANT_ID = process.env.TENANT_ID; // GUID
const CLIENT_ID = process.env.MICROSOFT_APP_ID; // App registration (client) ID
const CLIENT_SECRET = process.env.MICROSOFT_APP_PASSWORD; // Client secret VALUE

// Must match what you set in Azure Bot -> Teams channel -> Calling webhook
const CALLING_CALLBACK_URI =
  process.env.CALLING_CALLBACK_URI ||
  "https://lysara-e3e2f0dydffnfefs.southeastasia-01.azurewebsites.net/api/calling";

// ============================
// Startup warning if env missing
// ============================
(function assertEnv() {
  const missing = [];
  if (!TENANT_ID) missing.push("TENANT_ID");
  if (!CLIENT_ID) missing.push("MICROSOFT_APP_ID");
  if (!CLIENT_SECRET) missing.push("MICROSOFT_APP_PASSWORD");

  if (missing.length) {
    console.warn(
      `[WARN] Missing env vars: ${missing.join(", ")}. /join will fail until set.`
    );
  }
})();

// ============================
// Helpers
// ============================
async function getGraphToken() {
  const cred = new ClientSecretCredential(TENANT_ID, CLIENT_ID, CLIENT_SECRET);
  const token = await cred.getToken("https://graph.microsoft.com/.default");
  return token.token;
}

/**
 * Extract Oid (organizer AAD object id) from Teams join URL context param.
 * Example context: {"Tid":"...","Oid":"..."}
 */
function tryExtractOrganizerOid(joinWebUrl) {
  try {
    const u = new URL(joinWebUrl);
    const contextParam = u.searchParams.get("context");
    if (!contextParam) return null;

    const ctx = JSON.parse(decodeURIComponent(contextParam));
    return ctx?.Oid || null;
  } catch {
    return null;
  }
}

/**
 * Graph lookup for onlineMeeting by JoinWebUrl (must be URL-encoded as per docs).
 * We query:
 * GET /users/{organizerUserId}/onlineMeetings?$filter=JoinWebUrl eq '{encodedJoinWebUrl}'
 */
async function findOnlineMeeting({ token, organizerUserId, joinWebUrl }) {
  // Graph expects the joinWebUrl value to be URL encoded inside the filter
  const encodedJoinWebUrl = encodeURIComponent(joinWebUrl);

  const url =
    `https://graph.microsoft.com/v1.0/users/${organizerUserId}/onlineMeetings` +
    `?$filter=JoinWebUrl%20eq%20'${encodedJoinWebUrl}'`;

  try {
    const resp = await axios.get(url, {
      headers: { Authorization: `Bearer ${token}` }
    });

    const meeting = resp.data?.value?.[0] || null;
    return { meeting, tried: url };
  } catch (e) {
    return { meeting: null, tried: url, error: e?.response?.data || e.message };
  }
}

// ============================
// Routes
// ============================

// Health
app.get("/health", (_req, res) => res.json({ ok: true, build: BUILD_TAG }));

// Quick env sanity (remove later if you want)
app.get("/debug/env", (_req, res) => {
  res.json({
    TENANT_ID: TENANT_ID || null,
    MICROSOFT_APP_ID: CLIENT_ID ? "SET" : null,
    MICROSOFT_APP_PASSWORD: CLIENT_SECRET ? "SET" : null,
    CALLING_CALLBACK_URI
  });
});

// (Optional) quick check Graph can list meetings for a user (for diagnosing policy issues)
app.get("/debug/onlineMeetings/:organizerUserId", async (req, res) => {
  try {
    const token = await getGraphToken();
    const url = `https://graph.microsoft.com/v1.0/users/${req.params.organizerUserId}/onlineMeetings`;
    const r = await axios.get(url, { headers: { Authorization: `Bearer ${token}` } });
    res.json(r.data);
  } catch (e) {
    res.status(500).json({ error: e?.response?.data || e.message });
  }
});

// Bot Framework messaging endpoint (Azure Bot -> Messaging endpoint)
app.post("/api/messages", (_req, res) => {
  res.sendStatus(200);
});

// Calling webhook endpoint (Azure Bot -> Teams channel -> Calling webhook)
app.post("/api/calling", (req, res) => {
  console.log("=== CALLING WEBHOOK EVENT RECEIVED ===");
  console.log(JSON.stringify(req.body, null, 2));
  res.sendStatus(202);
});

/**
 * POST /join
 * Body:
 *  - joinWebUrl (required)
 *  - organizerUserId (optional; if omitted, we try to parse Oid from joinWebUrl context)
 *
 * This will:
 *  1) Find the onlineMeeting via JoinWebUrl filter
 *  2) Extract chatInfo.threadId
 *  3) POST /communications/calls to join the scheduled meeting
 */
app.post("/join", requireApiKey, async (req, res) => {
  try {
    let { joinWebUrl, organizerUserId } = req.body || {};

    if (!joinWebUrl) {
      return res.status(400).json({ error: "Missing joinWebUrl (Teams meeting join link)." });
    }

    // If organizerUserId not provided, try to extract from join link (Oid)
    if (!organizerUserId) {
      organizerUserId = tryExtractOrganizerOid(joinWebUrl);
    }

    if (!organizerUserId) {
      return res.status(400).json({
        error:
          "Missing organizerUserId and could not extract Oid from joinWebUrl context. Provide organizerUserId explicitly."
      });
    }

    const token = await getGraphToken();

    // 1) Look up onlineMeeting
    const found = await findOnlineMeeting({ token, organizerUserId, joinWebUrl });

    if (!found.meeting) {
      return res.status(404).json({
        error:
          "Online meeting not found for this organizerUserId + joinWebUrl. Common causes: organizerUserId not correct, the meeting isn't under this organizer, or app access policy/admin consent not applied.",
        tried: [found.tried],
        lookupError: found.error || null,
        organizerUserIdUsed: organizerUserId
      });
    }

    const threadId = found.meeting?.chatInfo?.threadId;
    if (!threadId) {
      return res.status(400).json({
        error:
          "Online meeting found, but chatInfo.threadId is missing. Cannot join scheduled meeting without threadId.",
        meetingId: found.meeting?.id
      });
    }

    // after you have found.meeting successfully:
const joinMeetingId = found.meeting?.joinMeetingIdSettings?.joinMeetingId;
const passcode = found.meeting?.joinMeetingIdSettings?.passcode ?? null;

if (!joinMeetingId) {
  return res.status(400).json({
    error: "joinMeetingIdSettings.joinMeetingId is missing from onlineMeeting. Cannot join via joinMeetingIdMeetingInfo."
  });
}

const payload = {
  "@odata.type": "#microsoft.graph.call",
  callbackUri: CALLING_CALLBACK_URI,
  requestedModalities: ["audio"],
  mediaConfig: {
    "@odata.type": "#microsoft.graph.serviceHostedMediaConfig"
  },
  meetingInfo: {
    "@odata.type": "#microsoft.graph.joinMeetingIdMeetingInfo",
    joinMeetingId: joinMeetingId,
    passcode: passcode // can be null if not required
  },
  tenantId: TENANT_ID
};

// POST https://graph.microsoft.com/v1.0/communications/calls

const createCallUrl = "https://graph.microsoft.com/v1.0/communications/calls";    
const callResp = await axios.post(createCallUrl, payload, {
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json"
      }
    });

    return res.status(200).json({
      ok: true,
      organizerUserIdUsed: organizerUserId,
      meeting: {
        id: found.meeting?.id,
        subject: found.meeting?.subject || null,
        threadId
      },
      call: callResp.data
    });
  } catch (e) {
    return res.status(500).json({ error: e?.response?.data || e.message });
  }
});

app.post("/call/:callId/hangup", requireApiKey, async (req, res) => {
  try {
    const token = await getGraphToken();
    const url = `https://graph.microsoft.com/v1.0/communications/calls/${req.params.callId}`;

    await axios.delete(url, {
      headers: { Authorization: `Bearer ${token}` }
    });

    res.json({ ok: true, callId: req.params.callId });
  } catch (e) {
    res.status(500).json({ error: e?.response?.data || e.message });
  }
});

app.get("/transcripts", requireApiKey, async (req, res) => {
  try {
    const joinWebUrl = req.query.joinWebUrl;
    if (!joinWebUrl) return res.status(400).json({ error: "Missing joinWebUrl query param" });

    // Extract organizer Oid from link
    const organizerUserId = tryExtractOrganizerOid(joinWebUrl);
    if (!organizerUserId) {
      return res.status(400).json({ error: "Could not extract organizer Oid from joinWebUrl context" });
    }

    const token = await getGraphToken();

    // Find online meeting (same helper used by /join)
    const found = await findOnlineMeeting({ token, organizerUserId, joinWebUrl });
    if (!found.meeting) {
      return res.status(404).json({ error: "Online meeting not found", tried: found.tried, lookupError: found.error || null });
    }

    const meetingId = found.meeting.id;

    // 1) List transcripts
    const listUrl = `https://graph.microsoft.com/v1.0/users/${organizerUserId}/onlineMeetings/${meetingId}/transcripts`;
    const listResp = await axios.get(listUrl, { headers: { Authorization: `Bearer ${token}` } });

    const transcripts = listResp.data?.value || [];
    if (!transcripts.length) {
      return res.status(404).json({
        error: "No transcripts found yet. Usually means transcription wasn't started or it's not processed yet.",
        meetingId
      });
    }

    // 2) Get transcript content for the newest transcript
    const latest = transcripts.sort((a, b) => (a.createdDateTime || "").localeCompare(b.createdDateTime || "")).pop();
    const transcriptId = latest.id;

    // Content endpoint returns VTT content
    const contentUrl = `https://graph.microsoft.com/v1.0/users/${organizerUserId}/onlineMeetings/${meetingId}/transcripts/${transcriptId}/content`;
    const contentResp = await axios.get(contentUrl, {
  headers: {
    Authorization: `Bearer ${token}`,
    Accept: "text/vtt"
  },
  responseType: "text"
});

    res.setHeader("Content-Type", "text/vtt; charset=utf-8");
    return res.status(200).send(contentResp.data);
  } catch (e) {
    return res.status(500).json({ error: e?.response?.data || e.message });
  }
});

// ============================
// Start server
// ============================
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server listening on ${PORT}`));