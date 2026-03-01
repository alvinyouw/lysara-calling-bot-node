import express from "express";
import axios from "axios";
import { ClientSecretCredential } from "@azure/identity";

const app = express();
app.use(express.json({ limit: "2mb" }));
// Tracks who is still in each call
const callState = new Map();
// callState.get(callId) = { participantIds: Set<string>, lastUpdate: Date.now() }

// ===== Duplicate-join guard (in-memory) =====
const activeCallsByThreadId = new Map(); // threadId -> callId

function getCallIdFromNotification(n) {
  // resourceUrl: "/communications/calls/<callId>" or "/communications/calls/<callId>/participants"
  const url = n?.resourceUrl || n?.resource || "";
  const m = String(url).match(/calls\/([0-9a-fA-F-]{36})/);
  return m ? m[1] : null;
}

function countNonBotParticipants(participants) {
  // Count humans only (user/guest). Exclude application (the bot).
  let count = 0;
  for (const p of participants || []) {
    const ident = p?.info?.identity || {};
    const isApp = !!ident.application;
    const isHuman = !!(ident.user || ident.guest);
    if (!isApp && isHuman) count += 1;
  }
  return count;
}

async function hangupCall(callId) {
  // Use same internal API you already built
  // IMPORTANT: this is calling your own protected endpoint, so include x-api-key
  const url = `http://127.0.0.1:${process.env.PORT || 3000}/call/${callId}/hangup`;
  await axios.post(url, null, { headers: { "x-api-key": process.env.API_KEY } });
  console.log(`[auto-leave] hangup requested callId=${callId}`);
}

// ===== API key protection for your own endpoints =====
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

const BUILD_TAG = "2026-03-01-auto-leave-v1";

// ===== ENV (Azure App Service -> Environment variables) =====
const TENANT_ID = process.env.TENANT_ID; // GUID
const CLIENT_ID = process.env.MICROSOFT_APP_ID; // App registration (client) ID
const CLIENT_SECRET = process.env.MICROSOFT_APP_PASSWORD; // Client secret VALUE

// Must match Azure Bot -> Teams channel -> Calling webhook
const CALLING_CALLBACK_URI =
  process.env.CALLING_CALLBACK_URI ||
  "https://lysara-e3e2f0dydffnfefs.southeastasia-01.azurewebsites.net/api/calling";

// ===== Startup warning if env missing =====
(function assertEnv() {
  const missing = [];
  if (!TENANT_ID) missing.push("TENANT_ID");
  if (!CLIENT_ID) missing.push("MICROSOFT_APP_ID");
  if (!CLIENT_SECRET) missing.push("MICROSOFT_APP_PASSWORD");

  if (missing.length) {
    console.warn(
      `[WARN] Missing env vars: ${missing.join(", ")}. /join & /transcripts will fail until set.`
    );
  }
})();

// ===== Helpers =====
async function getGraphToken() {
  const cred = new ClientSecretCredential(TENANT_ID, CLIENT_ID, CLIENT_SECRET);
  const token = await cred.getToken("https://graph.microsoft.com/.default");
  return token.token;
}

// Extract organizer Oid from Teams join URL `context` param: {"Tid":"...","Oid":"..."}
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

// Find onlineMeeting by JoinWebUrl (Graph expects URL-encoded joinWebUrl in filter)
async function findOnlineMeeting({ token, organizerUserId, joinWebUrl }) {
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

function isGuid(s) {
  return /^[0-9a-fA-F-]{36}$/.test(String(s || ""));
}

// ===== Routes =====

// Health
app.get("/health", (_req, res) => res.json({ ok: true, build: BUILD_TAG }));

// Debug env (optional)
app.get("/debug/env", (_req, res) => {
  res.json({
    TENANT_ID: TENANT_ID || null,
    MICROSOFT_APP_ID: CLIENT_ID ? "SET" : null,
    MICROSOFT_APP_PASSWORD: CLIENT_SECRET ? "SET" : null,
    CALLING_CALLBACK_URI
  });
});

// Bot Framework messaging endpoint (leave unprotected)
app.post("/api/messages", (_req, res) => res.sendStatus(200));

app.post("/api/calling", async (req, res) => {
  try {
    // Always ACK fast
    res.sendStatus(202);

    const body = req.body;
    const notifications = body?.value || [];
    if (!Array.isArray(notifications) || notifications.length === 0) return;

    for (const n of notifications) {
      console.log(`[webhook] changeType=${n.changeType} resourceUrl=${n.resourceUrl}`);
    }

    for (const n of notifications) {
      const callId = getCallIdFromNotification(n);
      if (!callId) continue;

      // 1) Call deleted/terminated -> cleanup
      const changeType = n?.changeType;
      const resourceUrl = n?.resourceUrl || "";

      // If the call ends, clear trackers
      if (changeType === "deleted" || (n?.resourceData?.state === "terminated")) {
        console.log(`[call] ended callId=${callId} changeType=${changeType}`);
        callState.delete(callId);

        // also remove from activeCallsByThreadId if it points to this callId
        for (const [tId, cId] of activeCallsByThreadId.entries()) {
          if (cId === callId) {
            activeCallsByThreadId.delete(tId);
            break;
          }
        }
        continue;
      }

      // 2) Participants updates
      if (resourceUrl.endsWith("/participants") && Array.isArray(n?.resourceData)) {
        const participants = n.resourceData;

        const nonBotCount = countNonBotParticipants(participants);

        // Init state
        const state =
          callState.get(callId) ||
          { lastNonBotCount: null, emptySince: null, hangupTimer: null };

        state.lastNonBotCount = nonBotCount;

        // If no humans left -> start countdown (avoid flapping)
        if (nonBotCount === 0) {
          if (!state.emptySince) {
            state.emptySince = Date.now();
            console.log(`[auto-leave] no humans left callId=${callId}, starting timer`);
          }

          if (!state.hangupTimer) {
            state.hangupTimer = setTimeout(async () => {
              try {
                // Re-check state just before hangup
                const latest = callState.get(callId);
                if (!latest) return;
                if (latest.lastNonBotCount === 0) {
                  await hangupCall(callId);
                } else {
                  console.log(`[auto-leave] humans returned, skipping hangup callId=${callId}`);
                }
              } catch (e) {
                console.log(`[auto-leave] hangup failed callId=${callId} error=${e?.message || e}`);
              }
            }, 25000); // 25 seconds grace period
          }
        } else {
          // Humans present -> cancel timer
          if (state.hangupTimer) {
            clearTimeout(state.hangupTimer);
            state.hangupTimer = null;
          }
          state.emptySince = null;
        }

        callState.set(callId, state);
        console.log(`[participants] callId=${callId} nonBotCount=${nonBotCount}`);
      }
    }
  } catch (e) {
    // webhook must not fail the request; we already 202’d, just log
    console.log("[webhook] error", e?.message || e);
  }
});

/**
 * POST /join (protected)
 * Body: { joinWebUrl }
 * - Extract organizer Oid from link
 * - Find meeting via Graph onlineMeetings filter
 * - Guard: don't join twice for same threadId
 * - Join via joinMeetingIdMeetingInfo if available (guest/lobby style)
 * - else fallback to organizerMeetingInfo join
 */
app.post("/join", requireApiKey, async (req, res) => {
  try {
    const { joinWebUrl } = req.body || {};
    if (!joinWebUrl) {
      return res.status(400).json({ error: "Missing joinWebUrl (Teams meeting join link)." });
    }

    const organizerUserId = tryExtractOrganizerOid(joinWebUrl);
    if (!organizerUserId) {
      return res.status(400).json({
        error: "Could not extract organizer Oid from joinWebUrl context."
      });
    }

    const token = await getGraphToken();

    const found = await findOnlineMeeting({ token, organizerUserId, joinWebUrl });
    if (!found.meeting) {
      return res.status(404).json({
        error: "Online meeting not found for organizerUserId + joinWebUrl.",
        tried: [found.tried],
        lookupError: found.error || null,
        organizerUserIdUsed: organizerUserId
      });
    }

    const threadId = found.meeting?.chatInfo?.threadId;
    if (!threadId) {
      return res.status(400).json({
        error: "Online meeting found, but chatInfo.threadId is missing.",
        meetingId: found.meeting?.id
      });
    }

    // Duplicate join guard
    if (activeCallsByThreadId.has(threadId)) {
      return res.status(409).json({
        error: "Bot already joined this meeting",
        callId: activeCallsByThreadId.get(threadId),
        threadId
      });
    }

    const joinMeetingId = found.meeting?.joinMeetingIdSettings?.joinMeetingId;
    const passcode = found.meeting?.joinMeetingIdSettings?.passcode ?? null;

    const createCallUrl = "https://graph.microsoft.com/v1.0/communications/calls";

    let payload;
    if (joinMeetingId) {
      // Guest/lobby-style join (you already tested this works)
      payload = {
        "@odata.type": "#microsoft.graph.call",
        callbackUri: CALLING_CALLBACK_URI,
        requestedModalities: ["audio"],
        mediaConfig: { "@odata.type": "#microsoft.graph.serviceHostedMediaConfig" },
        meetingInfo: {
          "@odata.type": "#microsoft.graph.joinMeetingIdMeetingInfo",
          joinMeetingId,
          passcode
        },
        tenantId: TENANT_ID
      };
    } else {
      // Fallback scheduled meeting join
      payload = {
        "@odata.type": "#microsoft.graph.call",
        callbackUri: CALLING_CALLBACK_URI,
        requestedModalities: ["audio"],
        mediaConfig: { "@odata.type": "#microsoft.graph.serviceHostedMediaConfig" },
        chatInfo: { "@odata.type": "#microsoft.graph.chatInfo", threadId, messageId: "0" },
        meetingInfo: {
          "@odata.type": "#microsoft.graph.organizerMeetingInfo",
          organizer: {
            "@odata.type": "#microsoft.graph.identitySet",
            user: { "@odata.type": "#microsoft.graph.identity", id: organizerUserId, tenantId: TENANT_ID }
          },
          allowConversationWithoutHost: true
        },
        tenantId: TENANT_ID
      };
    }

    const callResp = await axios.post(createCallUrl, payload, {
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" }
    });

    // Store callId to prevent duplicate joins
    activeCallsByThreadId.set(threadId, callResp.data.id);

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

// Hang up call (protected)
app.post("/call/:callId/hangup", requireApiKey, async (req, res) => {
  try {
    const callId = req.params.callId;
    if (!isGuid(callId)) {
      return res.status(400).json({ error: "callId must be a GUID" });
    }

    const token = await getGraphToken();
    const url = `https://graph.microsoft.com/v1.0/communications/calls/${callId}`;

    await axios.delete(url, { headers: { Authorization: `Bearer ${token}` } });

    // Remove from map if present
    for (const [tId, cId] of activeCallsByThreadId.entries()) {
      if (cId === callId) {
        activeCallsByThreadId.delete(tId);
        break;
      }
    }

    return res.json({ ok: true, callId });
  } catch (e) {
    return res.status(500).json({ error: e?.response?.data || e.message });
  }
});

// Check call state (protected)
app.get("/call/:callId", requireApiKey, async (req, res) => {
  try {
    const callId = req.params.callId;
    if (!isGuid(callId)) {
      return res.status(400).json({ error: "callId must be a GUID" });
    }

    const token = await getGraphToken();
    const url = `https://graph.microsoft.com/v1.0/communications/calls/${callId}`;
    const r = await axios.get(url, { headers: { Authorization: `Bearer ${token}` } });

    return res.json({
      ok: true,
      id: r.data?.id,
      state: r.data?.state,
      terminationReason: r.data?.terminationReason || null
    });
  } catch (e) {
    const status = e?.response?.status;
    if (status === 404) return res.status(404).json({ ok: true, state: "not_found_or_ended" });
    return res.status(500).json({ error: e?.response?.data || e.message });
  }
});

// Fetch transcript as VTT (protected)
app.get("/transcripts", requireApiKey, async (req, res) => {
  try {
    const joinWebUrl = req.query.joinWebUrl;
    if (!joinWebUrl) return res.status(400).json({ error: "Missing joinWebUrl query param" });

    const organizerUserId = tryExtractOrganizerOid(joinWebUrl);
    if (!organizerUserId) {
      return res.status(400).json({ error: "Could not extract organizer Oid from joinWebUrl context" });
    }

    const token = await getGraphToken();

    const found = await findOnlineMeeting({ token, organizerUserId, joinWebUrl });
    if (!found.meeting) {
      return res.status(404).json({
        error: "Online meeting not found",
        tried: found.tried,
        lookupError: found.error || null
      });
    }

    const meetingId = found.meeting.id;

    // List transcripts
    const listUrl = `https://graph.microsoft.com/v1.0/users/${organizerUserId}/onlineMeetings/${meetingId}/transcripts`;
    const listResp = await axios.get(listUrl, { headers: { Authorization: `Bearer ${token}` } });

    const transcripts = listResp.data?.value || [];
    if (!transcripts.length) {
      return res.status(404).json({
        error: "No transcripts found yet. Transcription may not be started or still processing.",
        meetingId
      });
    }

    // Get latest transcript
    const latest = transcripts
      .sort((a, b) => (a.createdDateTime || "").localeCompare(b.createdDateTime || ""))
      .pop();

    const transcriptId = latest.id;

    // Get transcript content (VTT)
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

// ===== Start server =====
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server listening on ${PORT}`));