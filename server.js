import express from "express";
import axios from "axios";
import { ClientSecretCredential } from "@azure/identity";

const app = express();
app.use(express.json({ limit: "2mb" }));

// ============================
// Environment variables (Azure App Service -> Environment variables)
// ============================
const TENANT_ID = process.env.TENANT_ID;
const CLIENT_ID = process.env.MICROSOFT_APP_ID; // Entra App Registration (client) ID
const CLIENT_SECRET = process.env.MICROSOFT_APP_PASSWORD; // Client secret value

// Your public calling webhook (must match what you set in Azure Bot -> Teams channel -> Calling)
const CALLING_CALLBACK_URI =
  process.env.CALLING_CALLBACK_URI ||
  "https://lysara-e3e2f0dydffnfefs.southeastasia-01.azurewebsites.net/api/calling";

// ============================
// Basic sanity checks at startup
// ============================
function assertEnv() {
  const missing = [];
  if (!TENANT_ID) missing.push("TENANT_ID");
  if (!CLIENT_ID) missing.push("MICROSOFT_APP_ID");
  if (!CLIENT_SECRET) missing.push("MICROSOFT_APP_PASSWORD");

  if (missing.length) {
    console.warn(
      `[WARN] Missing env vars: ${missing.join(", ")}. /join will fail until set.`
    );
  }
}
assertEnv();

// ============================
// Helpers
// ============================
async function getGraphToken() {
  const cred = new ClientSecretCredential(TENANT_ID, CLIENT_ID, CLIENT_SECRET);
  const token = await cred.getToken("https://graph.microsoft.com/.default");
  return token.token;
}

// Try to find the onlineMeeting by joinWebUrl using v1.0 first, then beta as fallback.
async function findOnlineMeeting({ token, organizerUserId, joinWebUrl }) {
  // OData: need to wrap string in single quotes; any single quote inside must be doubled.
  const safeUrl = String(joinWebUrl).replace(/'/g, "''");

  // Some environments accept joinWebUrl casing; some samples historically used JoinWebUrl.
  // We'll try both styles across v1.0 + beta.
  const candidates = [
    {
      label: "v1.0 joinWebUrl",
      url: `https://graph.microsoft.com/v1.0/users/${organizerUserId}/onlineMeetings?$filter=joinWebUrl eq '${encodeURIComponent(
        safeUrl
      )}'`,
      // NOTE: Above encode is for safety, but Graph can be picky. We'll also try without encode below.
      raw: `https://graph.microsoft.com/v1.0/users/${organizerUserId}/onlineMeetings?$filter=joinWebUrl eq '${safeUrl}'`
    },
    {
      label: "beta JoinWebUrl",
      url: `https://graph.microsoft.com/beta/users/${organizerUserId}/onlineMeetings?$filter=JoinWebUrl eq '${safeUrl}'`,
      raw: `https://graph.microsoft.com/beta/users/${organizerUserId}/onlineMeetings?$filter=JoinWebUrl eq '${safeUrl}'`
    }
  ];

  for (const c of candidates) {
    try {
      // Prefer raw (no encode). If it errors due to URL chars, try encoded variant.
      const resp = await axios.get(c.raw, {
        headers: { Authorization: `Bearer ${token}` }
      });

      const meeting = resp.data?.value?.[0];
      if (meeting) return { meeting, used: c.label, tried: c.raw };
    } catch (e1) {
      // Try encoded URL variant (for the one that has it)
      if (c.url && c.url !== c.raw) {
        try {
          const resp2 = await axios.get(c.url, {
            headers: { Authorization: `Bearer ${token}` }
          });
          const meeting2 = resp2.data?.value?.[0];
          if (meeting2) return { meeting: meeting2, used: c.label, tried: c.url };
        } catch (e2) {
          // continue to next candidate
        }
      }
      // continue to next candidate
    }
  }

  return { meeting: null, used: null, tried: candidates.map((x) => x.raw) };
}

// ============================
// Routes
// ============================

// Health check
app.get("/health", (_req, res) => {
  res.json({ ok: true });
});

// Bot Framework messaging endpoint (required by Azure Bot “Messaging endpoint”)
app.post("/api/messages", (_req, res) => {
  // If you're not using chat messages, a simple 200 OK is fine for now.
  res.sendStatus(200);
});

// Calling webhook endpoint (Graph posts call lifecycle notifications here)
app.post("/api/calling", (req, res) => {
  // IMPORTANT: For a production calling bot you should validate, store call state, and respond properly.
  // For now we log the body so you can see callbacks arriving.
  console.log("=== CALLING WEBHOOK EVENT RECEIVED ===");
  console.log(JSON.stringify(req.body, null, 2));
  res.sendStatus(202);
});

/**
 * Join meeting endpoint
 * Body requires:
 *  - organizerUserId (AAD user GUID of meeting organizer)
 *  - joinWebUrl (Teams meeting join URL)
 *
 * It will:
 *  1) Find onlineMeeting and extract chatInfo.threadId
 *  2) Create call via POST /communications/calls (join scheduled meeting)
 *
 * Docs: Create call requires threadId/messageId/organizerId/tenantId for scheduled meeting joins.  [oai_citation:2‡Microsoft Learn](https://learn.microsoft.com/en-us/graph/api/application-post-calls?view=graph-rest-1.0&utm_source=chatgpt.com)
 */
app.post("/join", async (req, res) => {
  try {
    const { organizerUserId, joinWebUrl } = req.body || {};

    if (!organizerUserId) {
      return res.status(400).json({ error: "Missing organizerUserId (AAD user GUID)." });
    }
    if (!joinWebUrl) {
      return res.status(400).json({ error: "Missing joinWebUrl (Teams meeting join link)." });
    }

    const token = await getGraphToken();

    // 1) Find meeting details
    const found = await findOnlineMeeting({ token, organizerUserId, joinWebUrl });
    if (!found.meeting) {
      return res.status(404).json({
        error:
          "Online meeting not found for this organizerUserId + joinWebUrl. Common causes: organizerUserId is wrong, the user isn't organizer/attendee, or application access policy/admin consent not applied.",
        tried: found.tried
      });
    }

    const threadId = found.meeting?.chatInfo?.threadId;
    if (!threadId) {
      return res.status(400).json({
        error:
          "Found onlineMeeting but missing chatInfo.threadId. Cannot join scheduled meeting without threadId.",
        meetingId: found.meeting?.id,
        used: found.used
      });
    }

    // 2) Create call (join scheduled meeting)
    const createCallUrl = "https://graph.microsoft.com/v1.0/communications/calls";

    const payload = {
      "@odata.type": "#microsoft.graph.call",
      callbackUri: CALLING_CALLBACK_URI,
      requestedModalities: ["audio"],
      mediaConfig: {
        "@odata.type": "#microsoft.graph.serviceHostedMediaConfig"
      },
      chatInfo: {
        "@odata.type": "#microsoft.graph.chatInfo",
        threadId: threadId,
        messageId: "0"
      },
      meetingInfo: {
        "@odata.type": "#microsoft.graph.organizerMeetingInfo",
        organizer: {
          "@odata.type": "#microsoft.graph.identitySet",
          user: {
            "@odata.type": "#microsoft.graph.identity",
            id: organizerUserId,
            tenantId: TENANT_ID
          }
        },
        allowConversationWithoutHost: true
      },
      tenantId: TENANT_ID
    };

    const callResp = await axios.post(createCallUrl, payload, {
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json"
      }
    });

    return res.status(200).json({
      ok: true,
      meetingLookup: { used: found.used, meetingId: found.meeting?.id, threadId },
      call: callResp.data
    });
  } catch (e) {
    const details = e?.response?.data || e?.message || String(e);
    return res.status(500).json({ error: details });
  }
});

// ============================
// Start server
// ============================
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server listening on port ${PORT}`);
});