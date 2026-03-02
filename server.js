import express from "express";
import axios from "axios";
import { ClientSecretCredential } from "@azure/identity";

// =====================================================
// In-memory state
// =====================================================

// callId -> {
//   createdAt, firstHumanJoinAt, lastHumanSeenAt,
//   lastNonBotCount, emptySince, hangupTimer
// }
const callState = new Map();

// threadId -> callId (duplicate-join guard)
const activeCallsByThreadId = new Map();

// =====================================================
// App setup
// =====================================================
const app = express();
app.use(express.json({ limit: "2mb" }));

const BUILD_TAG = "2026-03-02-clean-status-leave-webhook-v1";

// =====================================================
// ENV
// =====================================================
const TENANT_ID = process.env.TENANT_ID;
const CLIENT_ID = process.env.MICROSOFT_APP_ID;
const CLIENT_SECRET = process.env.MICROSOFT_APP_PASSWORD;

const CALLING_CALLBACK_URI =
  process.env.CALLING_CALLBACK_URI ||
  "https://lysara-e3e2f0dydffnfefs.southeastasia-01.azurewebsites.net/api/calling";

(function assertEnv() {
  const missing = [];
  if (!TENANT_ID) missing.push("TENANT_ID");
  if (!CLIENT_ID) missing.push("MICROSOFT_APP_ID");
  if (!CLIENT_SECRET) missing.push("MICROSOFT_APP_PASSWORD");
  if (!process.env.API_KEY) missing.push("API_KEY");

  if (missing.length) {
    console.warn(`[WARN] Missing env vars: ${missing.join(", ")}.`);
  }
})();

// =====================================================
// Auth middleware
// =====================================================
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

function isGuid(s) {
  return /^[0-9a-fA-F-]{36}$/.test(String(s || ""));
}

// =====================================================
// Helpers
// =====================================================
async function getGraphToken() {
  const cred = new ClientSecretCredential(TENANT_ID, CLIENT_ID, CLIENT_SECRET);
  const token = await cred.getToken("https://graph.microsoft.com/.default");
  return token.token;
}

// Extract organizer Oid from Teams join URL context param: {"Tid":"...","Oid":"..."}
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
    const resp = await axios.get(url, { headers: { Authorization: `Bearer ${token}` } });
    const meeting = resp.data?.value?.[0] || null;
    return { meeting, tried: url };
  } catch (e) {
    return { meeting: null, tried: url, error: e?.response?.data || e.message };
  }
}

function getCallIdFromNotification(n) {
  const url = n?.resourceUrl || n?.resource || "";
  const m = String(url).match(/calls\/([0-9a-fA-F-]{36})/);
  return m ? m[1] : null;
}

// Count humans from webhook participant payload (resourceData array)
function countNonBotParticipants(participants) {
  let count = 0;
  for (const p of participants || []) {
    const ident = p?.info?.identity || {};
    const isApp = !!ident.application;

    // Some payloads have user/guest; some only have displayName somewhere
    const displayName =
      ident?.user?.displayName ||
      ident?.guest?.displayName ||
      ident?.encrypted?.displayName ||
      ident?.phone?.displayName ||
      p?.info?.displayName ||
      p?.info?.identity?.user?.displayName ||
      p?.info?.identity?.guest?.displayName ||
      null;

    const isHuman = !!(ident.user || ident.guest || displayName);
    if (!isApp && isHuman) count += 1;
  }
  return count;
}

async function hangupCall(callId) {
  // Call our own protected endpoint internally
  const port = process.env.PORT || 8080;
  const url = `http://127.0.0.1:${port}/call/${callId}/hangup`;
  await axios.post(url, null, { headers: { "x-api-key": process.env.API_KEY } });
  console.log(`[auto-leave] hangup requested callId=${callId}`);
}

function cleanupCall(callId) {
  const st = callState.get(callId);
  if (st?.hangupTimer) {
    clearTimeout(st.hangupTimer);
  }
  callState.delete(callId);

  // Remove from activeCallsByThreadId if present
  for (const [tId, cId] of activeCallsByThreadId.entries()) {
    if (cId === callId) {
      activeCallsByThreadId.delete(tId);
      break;
    }
  }
}

// =====================================================
// Routes
// =====================================================

// Health
app.get("/health", (_req, res) => res.json({ ok: true, build: BUILD_TAG }));

// Debug env (safe: don’t print secrets)
app.get("/debug/env", (_req, res) => {
  res.json({
    TENANT_ID: TENANT_ID || null,
    MICROSOFT_APP_ID: CLIENT_ID ? "SET" : null,
    MICROSOFT_APP_PASSWORD: CLIENT_SECRET ? "SET" : null,
    API_KEY: process.env.API_KEY ? "SET" : null,
    CALLING_CALLBACK_URI,
    build: BUILD_TAG,
  });
});

// Bot Framework messaging endpoint (leave unprotected)
app.post("/api/messages", (_req, res) => res.sendStatus(200));

/**
 * Calling webhook endpoint (leave unprotected)
 * Graph expects quick ACK. We:
 * - ACK 202 immediately
 * - then process notifications, track participants
 * - auto-leave when humans == 0 for 25s
 */
app.post("/api/calling", (req, res) => {
  res.sendStatus(202);

  setImmediate(async () => {
    try {
      const notifications = req.body?.value || [];
      if (!Array.isArray(notifications) || notifications.length === 0) return;

      for (const n of notifications) {
        console.log(`[webhook] changeType=${n.changeType} resourceUrl=${n.resourceUrl}`);
      }

      for (const n of notifications) {
        const callId = getCallIdFromNotification(n);
        if (!callId) continue;

        const changeType = n?.changeType;
        const resourceUrl = n?.resourceUrl || "";

        // Call ended
        if (changeType === "deleted" || n?.resourceData?.state === "terminated") {
          console.log(`[call] ended callId=${callId} changeType=${changeType}`);
          cleanupCall(callId);
          continue;
        }

        // Participants update
        if (resourceUrl.endsWith("/participants") && Array.isArray(n?.resourceData)) {
          const participants = n.resourceData;
          const nonBotCount = countNonBotParticipants(participants);

          const now = Date.now();
          const st = callState.get(callId) || {
            createdAt: now,
            firstHumanJoinAt: null,
            lastHumanSeenAt: null,
            lastNonBotCount: null,
            emptySince: null,
            hangupTimer: null,
          };

          st.lastNonBotCount = nonBotCount;

          if (nonBotCount > 0) {
            if (!st.firstHumanJoinAt) {
              st.firstHumanJoinAt = now;
              console.log(`[join] first human joined callId=${callId} at=${new Date(now).toISOString()}`);
            }
            st.lastHumanSeenAt = now;

            // cancel hangup timer
            if (st.hangupTimer) {
              clearTimeout(st.hangupTimer);
              st.hangupTimer = null;
            }
            st.emptySince = null;
          } else {
            // no humans
            if (!st.emptySince) {
              st.emptySince = now;
              console.log(`[auto-leave] no humans left callId=${callId}, starting timer`);
            }

            if (!st.hangupTimer) {
              st.hangupTimer = setTimeout(async () => {
                try {
                  const latest = callState.get(callId);
                  if (!latest) return;

                  if ((latest.lastNonBotCount ?? 0) === 0) {
                    await hangupCall(callId);
                  } else {
                    console.log(`[auto-leave] humans returned, skipping hangup callId=${callId}`);
                  }
                } catch (e) {
                  console.log(`[auto-leave] hangup failed callId=${callId} error=${e?.message || e}`);
                }
              }, 25_000);
            }
          }

          callState.set(callId, st);
          console.log(`[participants] callId=${callId} nonBotCount=${nonBotCount}`);
        }
      }
    } catch (e) {
      console.log("[webhook] error", e?.message || e);
    }
  });
});

/**
 * POST /join (protected)
 * Body: { joinWebUrl }
 * Returns: { call: { id, state, ... }, meeting: { threadId, ... } }
 */
app.post("/join", requireApiKey, async (req, res) => {
  try {
    const { joinWebUrl } = req.body || {};
    if (!joinWebUrl) {
      return res.status(400).json({ error: "Missing joinWebUrl (Teams meeting join link)." });
    }

    const organizerUserId = tryExtractOrganizerOid(joinWebUrl);
    if (!organizerUserId) {
      return res.status(400).json({ error: "Could not extract organizer Oid from joinWebUrl context." });
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

    const payload = {
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

    const callResp = await axios.post(createCallUrl, payload, {
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" }
    });

    // Save mapping + init local state
    activeCallsByThreadId.set(threadId, callResp.data.id);
    callState.set(callResp.data.id, {
      createdAt: Date.now(),
      firstHumanJoinAt: null,
      lastHumanSeenAt: null,
      lastNonBotCount: null,
      emptySince: null,
      hangupTimer: null,
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

/**
 * GET /status?call_id=<GUID>  (protected)
 * Poller expects stable fields:
 *   azure_status, call_state, termination_reason, human_count, is_ended
 */
app.get("/status", requireApiKey, async (req, res) => {
  try {
    const callId = req.query.call_id || req.query.bot_job_id || req.query.callId;
    if (!callId) return res.status(400).json({ error: "Missing call_id" });
    if (!isGuid(callId)) return res.status(400).json({ error: "call_id must be a GUID" });

    const token = await getGraphToken();

    // 1) Call state
    const callUrl = `https://graph.microsoft.com/v1.0/communications/calls/${callId}`;
    const callResp = await axios.get(callUrl, { headers: { Authorization: `Bearer ${token}` } });
    const callStateStr = callResp.data?.state || null; // establishing / established / terminated
    const ended = callStateStr === "terminated";

    // 2) Participants (best-effort). If it fails, use webhook-cached count if available.
    let humanCount = null;
    try {
      const partUrl = `https://graph.microsoft.com/v1.0/communications/calls/${callId}/participants`;
      const partResp = await axios.get(partUrl, { headers: { Authorization: `Bearer ${token}` } });
      const participants = partResp.data?.value || [];

      let count = 0;
      for (const p of participants) {
        const ident = p?.info?.identity || {};
        const isApp = !!ident.application;
        const isHuman = !!(ident.user || ident.guest);
        if (!isApp && isHuman) count += 1;
      }
      humanCount = count;
    } catch {
      const cached = callState.get(callId);
      if (typeof cached?.lastNonBotCount === "number") humanCount = cached.lastNonBotCount;
    }

    return res.json({
      ok: true,
      azure_status: ended ? "ended" : "running",
      call_state: callStateStr,
      termination_reason: callResp.data?.terminationReason || null,
      human_count: humanCount,
      is_ended: ended
    });
  } catch (e) {
    const status = e?.response?.status;

    // Graph returns 404 when call already ended / not found
    if (status === 404) {
      return res.status(404).json({
        ok: false,
        azure_status: "not_found",
        is_ended: true,
        call_state: "not_found",
        human_count: 0
      });
    }

    return res.status(502).json({
      ok: false,
      azure_status: "error",
      graph_status: status || null,
      error: e?.response?.data || e.message
    });
  }
});

/**
 * POST /leave  (protected)
 * body: { call_id: "<GUID>" }   OR query ?call_id=<GUID>
 */
app.post("/leave", requireApiKey, async (req, res) => {
  try {
    const callId = req.body?.call_id || req.body?.bot_job_id || req.query.call_id;
    if (!callId || !isGuid(callId)) {
      return res.status(400).json({ error: "Missing or invalid call_id (GUID)" });
    }

    const token = await getGraphToken();
    const url = `https://graph.microsoft.com/v1.0/communications/calls/${callId}`;
    await axios.delete(url, { headers: { Authorization: `Bearer ${token}` } });

    cleanupCall(callId);
    return res.json({ ok: true, call_id: callId });
  } catch (e) {
    const status = e?.response?.status;
    if (status === 404) {
      cleanupCall(req.body?.call_id || req.query.call_id);
      return res.json({ ok: true, call_id: req.body?.call_id || req.query.call_id, already_gone: true });
    }
    return res.status(500).json({ error: e?.response?.data || e.message });
  }
});

// Convenience: hangup by path param (protected)
app.post("/call/:callId/hangup", requireApiKey, async (req, res) => {
  try {
    const callId = req.params.callId;
    if (!isGuid(callId)) return res.status(400).json({ error: "callId must be a GUID" });

    const token = await getGraphToken();
    const url = `https://graph.microsoft.com/v1.0/communications/calls/${callId}`;
    await axios.delete(url, { headers: { Authorization: `Bearer ${token}` } });

    cleanupCall(callId);
    return res.json({ ok: true, callId });
  } catch (e) {
    return res.status(500).json({ error: e?.response?.data || e.message });
  }
});

// Convenience: check call by path param (protected)
app.get("/call/:callId", requireApiKey, async (req, res) => {
  try {
    const callId = req.params.callId;
    if (!isGuid(callId)) return res.status(400).json({ error: "callId must be a GUID" });

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
    const contentUrl =
      `https://graph.microsoft.com/v1.0/users/${organizerUserId}/onlineMeetings/${meetingId}` +
      `/transcripts/${transcriptId}/content`;

    const contentResp = await axios.get(contentUrl, {
      headers: { Authorization: `Bearer ${token}`, Accept: "text/vtt" },
      responseType: "text"
    });

    res.setHeader("Content-Type", "text/vtt; charset=utf-8");
    return res.status(200).send(contentResp.data);
  } catch (e) {
    return res.status(500).json({ error: e?.response?.data || e.message });
  }
});

// =====================================================
// Start server
// =====================================================
const PORT = process.env.PORT || 8080;
app.listen(PORT, () => console.log(`Server listening on ${PORT}`));