import express from "express";
import axios from "axios";
import { ClientSecretCredential } from "@azure/identity";

// =====================================================
// In-memory state
// =====================================================

// callId -> {
//   createdAt,
//   firstHumanJoinAt,
//   lastHumanSeenAt,
//   lastNonBotCount,
//   emptySince,
//   hangupTimer,
//   endedAt,
// }
const callState = new Map();

// threadId -> callId (duplicate-join guard)
const activeCallsByThreadId = new Map();

// =====================================================
// App setup
// =====================================================
const app = express();
app.use(express.json({ limit: "2mb" }));

const BUILD_TAG = "2026-03-02-bot-clean-v2";

// =====================================================
// ENV
// =====================================================
const TENANT_ID = process.env.TENANT_ID;
const CLIENT_ID = process.env.MICROSOFT_APP_ID;
const CLIENT_SECRET = process.env.MICROSOFT_APP_PASSWORD;
const API_KEY = process.env.API_KEY;

const CALLING_CALLBACK_URI =
  process.env.CALLING_CALLBACK_URI ||
  "https://lysara-e3e2f0dydffnfefs.southeastasia-01.azurewebsites.net/api/calling";

(function assertEnv() {
  const missing = [];
  if (!TENANT_ID) missing.push("TENANT_ID");
  if (!CLIENT_ID) missing.push("MICROSOFT_APP_ID");
  if (!CLIENT_SECRET) missing.push("MICROSOFT_APP_PASSWORD");
  if (!API_KEY) missing.push("API_KEY");
  if (missing.length) console.warn(`[WARN] Missing env vars: ${missing.join(", ")}`);
})();

// =====================================================
// Auth middleware
// =====================================================
function requireApiKey(req, res, next) {
  const key = req.header("x-api-key");
  if (!API_KEY) return res.status(500).json({ error: "API_KEY not configured on server" });
  if (key !== API_KEY) return res.status(401).json({ error: "Unauthorized" });
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

function getCallIdFromNotification(n) {
  const url = n?.resourceUrl || n?.resource || "";
  const m = String(url).match(/calls\/([0-9a-fA-F-]{36})/);
  return m ? m[1] : null;
}

function countNonBotParticipants(participants) {
  let count = 0;
  for (const p of participants || []) {
    const ident = p?.info?.identity || {};
    const isApp = !!ident.application;

    const displayName =
      ident?.user?.displayName ||
      ident?.guest?.displayName ||
      ident?.encrypted?.displayName ||
      ident?.phone?.displayName ||
      p?.info?.displayName ||
      null;

    const isHuman = !!(ident.user || ident.guest || displayName);
    if (!isApp && isHuman) count += 1;
  }
  return count;
}

function cleanupCall(callId) {
  const st = callState.get(callId);
  if (st?.hangupTimer) clearTimeout(st.hangupTimer);
  callState.delete(callId);

  for (const [tId, cId] of activeCallsByThreadId.entries()) {
    if (cId === callId) {
      activeCallsByThreadId.delete(tId);
      break;
    }
  }
}

async function graphHangup(callId) {
  const token = await getGraphToken();
  const url = `https://graph.microsoft.com/v1.0/communications/calls/${callId}`;
  await axios.delete(url, { headers: { Authorization: `Bearer ${token}` } });
}

async function hangupCall(callId) {
  // call our own internal endpoint
  const port = process.env.PORT || 8080;
  const url = `http://127.0.0.1:${port}/call/${callId}/hangup`;
  await axios.post(url, null, { headers: { "x-api-key": API_KEY } });
  console.log(`[auto-leave] hangup requested callId=${callId}`);
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

// =====================================================
// Routes
// =====================================================

// Health
app.get("/health", (_req, res) => res.json({ ok: true, build: BUILD_TAG }));

// Bot Framework messaging endpoint (leave unprotected)
app.post("/api/messages", (_req, res) => res.sendStatus(200));

/**
 * Calling webhook endpoint (unprotected)
 * - ACK immediately
 * - update in-memory participant state
 * - auto-leave only AFTER at least one human joined once
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
          const st = callState.get(callId);
          if (st) st.endedAt = Date.now();
          cleanupCall(callId);
          continue;
        }

        // Participants update (resourceData is array)
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
            endedAt: null,
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

            callState.set(callId, st);
            console.log(`[participants] callId=${callId} nonBotCount=${nonBotCount}`);
            continue;
          }

          // nonBotCount === 0
          // IMPORTANT: don't auto-leave until at least one human has joined once
          if (!st.firstHumanJoinAt) {
            callState.set(callId, st);
            console.log(`[participants] callId=${callId} nonBotCount=0 (waiting for first human; no auto-leave)`);
            continue;
          }

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
 */
app.post("/join", requireApiKey, async (req, res) => {
  try {
    const { joinWebUrl } = req.body || {};
    if (!joinWebUrl) return res.status(400).json({ error: "Missing joinWebUrl" });

    const organizerUserId = tryExtractOrganizerOid(joinWebUrl);
    if (!organizerUserId) return res.status(400).json({ error: "Could not extract organizer Oid from joinWebUrl" });

    const token = await getGraphToken();
    const found = await findOnlineMeeting({ token, organizerUserId, joinWebUrl });
    if (!found.meeting) {
      return res.status(404).json({
        error: "Online meeting not found",
        tried: found.tried,
        lookupError: found.error || null,
      });
    }

    const threadId = found.meeting?.chatInfo?.threadId;
    if (!threadId) return res.status(400).json({ error: "chatInfo.threadId missing" });

    if (activeCallsByThreadId.has(threadId)) {
      return res.status(409).json({ error: "Bot already joined this meeting", callId: activeCallsByThreadId.get(threadId) });
    }

    const joinMeetingId = found.meeting?.joinMeetingIdSettings?.joinMeetingId;
    const passcode = found.meeting?.joinMeetingIdSettings?.passcode ?? null;

    const createCallUrl = "https://graph.microsoft.com/v1.0/communications/calls";
    const payload = {
      "@odata.type": "#microsoft.graph.call",
      callbackUri: CALLING_CALLBACK_URI,
      requestedModalities: ["audio"],
      mediaConfig: { "@odata.type": "#microsoft.graph.serviceHostedMediaConfig" },
      meetingInfo: { "@odata.type": "#microsoft.graph.joinMeetingIdMeetingInfo", joinMeetingId, passcode },
      tenantId: TENANT_ID,
    };

    const callResp = await axios.post(createCallUrl, payload, {
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    });

    const callId = callResp.data?.id;
    if (callId) {
      activeCallsByThreadId.set(threadId, callId);
      callState.set(callId, {
        createdAt: Date.now(),
        firstHumanJoinAt: null,
        lastHumanSeenAt: null,
        lastNonBotCount: null,
        emptySince: null,
        hangupTimer: null,
        endedAt: null,
      });
    }

    return res.json({
      ok: true,
      organizerUserIdUsed: organizerUserId,
      meeting: { id: found.meeting?.id, subject: found.meeting?.subject || null, threadId },
      call: callResp.data,
    });
  } catch (e) {
    return res.status(500).json({ error: e?.response?.data || e.message });
  }
});

/**
 * GET /status?call_id=<GUID>  (protected)
 * IMPORTANT: DO NOT depend on Graph GET here.
 * Use webhook memory as truth.
 */
app.get("/status", requireApiKey, async (req, res) => {
  const callId = req.query.call_id || req.query.bot_job_id || req.query.callId;
  if (!callId) return res.status(400).json({ error: "Missing call_id" });
  if (!isGuid(callId)) return res.status(400).json({ error: "call_id must be a GUID" });

  const st = callState.get(callId);

  // If we never saw this callId in webhook memory, treat as ended/not_found
  if (!st) {
    return res.status(404).json({
      ok: false,
      azure_status: "not_found",
      is_ended: true,
      call_state: "not_found",
      human_count: 0,
    });
  }

  const humanCount = typeof st.lastNonBotCount === "number" ? st.lastNonBotCount : null;

  return res.json({
    ok: true,
    azure_status: "running",
    call_state: "established", // best-effort; webhook doesn't always tell state
    termination_reason: null,
    human_count: humanCount,
    is_ended: false,
    first_human_join_at: st.firstHumanJoinAt ? new Date(st.firstHumanJoinAt).toISOString() : null,
    last_human_seen_at: st.lastHumanSeenAt ? new Date(st.lastHumanSeenAt).toISOString() : null,
  });
});

/**
 * POST /leave  (protected)
 * body: { call_id: "<GUID>" }  OR query ?call_id=<GUID>
 * This DOES call Graph to hang up.
 */
app.post("/leave", requireApiKey, async (req, res) => {
  const callId = req.body?.call_id || req.query.call_id;
  if (!callId || !isGuid(callId)) return res.status(400).json({ error: "Missing or invalid call_id" });

  try {
    await graphHangup(callId);
  } catch (e) {
    const status = e?.response?.status;
    if (status !== 404) {
      return res.status(500).json({ error: e?.response?.data || e.message });
    }
  }

  cleanupCall(callId);
  return res.json({ ok: true, call_id: callId });
});

// Convenience: hangup by path param (protected)
app.post("/call/:callId/hangup", requireApiKey, async (req, res) => {
  const callId = req.params.callId;
  if (!isGuid(callId)) return res.status(400).json({ error: "callId must be a GUID" });

  try {
    await graphHangup(callId);
  } catch (e) {
    const status = e?.response?.status;
    if (status !== 404) return res.status(500).json({ error: e?.response?.data || e.message });
  }

  cleanupCall(callId);
  return res.json({ ok: true, callId });
});

// =====================================================
// Start server
// =====================================================
const PORT = process.env.PORT || 8080;
app.listen(PORT, () => console.log(`Server listening on ${PORT}`));