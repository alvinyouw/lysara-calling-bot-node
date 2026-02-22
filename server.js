import express from "express";
import axios from "axios";
import { ClientSecretCredential } from "@azure/identity";

const app = express();
app.use(express.json());

// ===== Required env vars you will configure in Azure App Service =====
const TENANT_ID = process.env.TENANT_ID;
const CLIENT_ID = process.env.MICROSOFT_APP_ID;        // Entra App Registration (client) ID
const CLIENT_SECRET = process.env.MICROSOFT_APP_PASSWORD;

// ----- Graph token helper (app-only) -----
async function getGraphToken() {
  const cred = new ClientSecretCredential(TENANT_ID, CLIENT_ID, CLIENT_SECRET);
  const token = await cred.getToken("https://graph.microsoft.com/.default");
  return token.token;
}

// Health check
app.get("/health", (_, res) => res.json({ ok: true }));

// Azure Bot "Messaging endpoint" must exist
app.post("/api/messages", (req, res) => {
  // If you don't need chat messages, keep this as a 200 OK.
  res.sendStatus(200);
});

// Teams Calling webhook must exist
app.post("/api/calling", (req, res) => {
  // Minimal acknowledge so Graph doesn't keep retrying.
  // You will expand this later to properly handle call lifecycle events.
  res.sendStatus(202);
});

// Endpoint Lovable (or your scheduler) calls to join a meeting
app.post("/join", async (req, res) => {
  try {
    const { joinUrl } = req.body;
    if (!joinUrl) return res.status(400).json({ error: "Missing joinUrl" });

    const token = await getGraphToken();

    // NOTE: The Create Call payload depends on what meeting identity you have.
    // We keep this intentionally as a placeholder until you provide the exact meeting info you store.
    const createCallUrl = "https://graph.microsoft.com/v1.0/communications/calls";
    const payload = {
      // TODO: Replace with correct meetingInfo payload for joinUrl / meeting identity
    };

    const resp = await axios.post(createCallUrl, payload, {
      headers: { Authorization: `Bearer ${token}` }
    });

    res.status(200).json({ ok: true, graph: resp.data });
  } catch (e) {
    res.status(500).json({ error: e?.response?.data || e.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Listening on ${PORT}`));