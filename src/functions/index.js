const admin = require("firebase-admin");
const { onRequest } = require("firebase-functions/v2/https");
const { defineSecret } = require("firebase-functions/params");

admin.initializeApp();

const GEMINI_API_KEYS = defineSecret("GEMINI_API_KEYS");

const FREE_AI_TABLE_LIMIT = 2;
const MODEL = "gemini-2.5-flash";

function sendCors(res) {
  res.set("Access-Control-Allow-Origin", "*");
  res.set("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.set("Access-Control-Allow-Headers", "Content-Type, Authorization");
}

function extractJson(text) {
  const cleaned = String(text || "")
    .replace(/```json/g, "")
    .replace(/```/g, "")
    .trim();

  const start = cleaned.indexOf("{");
  const end = cleaned.lastIndexOf("}");

  if (start === -1 || end === -1) {
    throw new Error("AI did not return JSON");
  }

  return JSON.parse(cleaned.slice(start, end + 1));
}

async function verifyUser(req) {
  const authHeader = req.headers.authorization || "";

  if (!authHeader.startsWith("Bearer ")) {
    throw new Error("Missing auth token");
  }

  const idToken = authHeader.replace("Bearer ", "");
  return admin.auth().verifyIdToken(idToken);
}

async function callGemini(apiKey, userPrompt) {
  const systemPrompt = `
You are an AI table generator for a spreadsheet app.
Return ONLY valid JSON. No markdown. No explanation.

JSON shape:
{
  "title": "Short table title",
  "columns": ["ID", "Column 1", "Column 2"],
  "rows": [
    ["1", "value", "value"],
    ["2", "value", "value"]
  ],
  "columnTypes": {
    "Email": "email",
    "Date": "date",
    "Status": "dropdown"
  }
}

Rules:
- Always include ID as the first column.
- Make useful business/data-entry tables.
- If user asks for row count, follow it. Otherwise create 8 rows.
- Keep columns practical and clean.
- Do not include code fences.
`;

  const url =
    `https://generativelanguage.googleapis.com/v1beta/models/${MODEL}:generateContent?key=${apiKey}`;

  const response = await fetch(url, {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      contents: [
        {
          role: "user",
          parts: [
            {
              text: `${systemPrompt}\n\nUser request: ${userPrompt}`
            }
          ]
        }
      ],
      generationConfig: {
        temperature: 0.4,
        responseMimeType: "application/json"
      }
    })
  });

  const data = await response.json();

  if (!response.ok) {
    const message = data?.error?.message || "Gemini request failed";
    const error = new Error(message);
    error.status = response.status;
    throw error;
  }

  const text = data?.candidates?.[0]?.content?.parts?.[0]?.text || "";
  return extractJson(text);
}

exports.aiTable = onRequest(
  {
    region: "us-central1",
    secrets: [GEMINI_API_KEYS],
    timeoutSeconds: 60,
    memory: "512MiB"
  },
  async (req, res) => {
    sendCors(res);

    if (req.method === "OPTIONS") {
      res.status(204).send("");
      return;
    }

    if (req.method !== "POST") {
      res.status(405).json({ error: "Method not allowed" });
      return;
    }

    try {
      const decoded = await verifyUser(req);
      const uid = decoded.uid;
      const email = decoded.email || "";

      const prompt = String(req.body?.prompt || "").trim();

      if (!prompt) {
        res.status(400).json({ error: "Prompt is required" });
        return;
      }

      const userRef = admin.firestore().doc(`users/${uid}`);
      const userSnap = await userRef.get();
      const userData = userSnap.exists ? userSnap.data() : {};

      const isAdmin = email === "iftia5061@gmail.com";
      const isPremium = Boolean(userData.isPremium) || isAdmin;
      const currentUsage = Number(userData.aiTableUsage || 0);

      if (!isPremium && currentUsage >= FREE_AI_TABLE_LIMIT) {
        res.status(402).json({
          error: `Free plan allows ${FREE_AI_TABLE_LIMIT} AI tables. Upgrade for unlimited AI tables.`,
          code: "AI_FREE_LIMIT_REACHED"
        });
        return;
      }

      const rawKeys = GEMINI_API_KEYS.value() || "";
      const apiKeys = rawKeys
        .split(",")
        .map(key => key.trim())
        .filter(Boolean);

      if (!apiKeys.length) {
        res.status(500).json({ error: "No Gemini API key configured" });
        return;
      }

      let result = null;
      let lastError = null;

      for (const key of apiKeys) {
        try {
          result = await callGemini(key, prompt);
          break;
        } catch (err) {
          lastError = err;

          if (err.status !== 429 && err.status !== 403) {
            break;
          }
        }
      }

      if (!result) {
        const status = lastError?.status === 429 ? 429 : 500;
        res.status(status).json({
          error: status === 429
            ? "AI quota limit reached today. Try again later."
            : "AI table generation failed."
        });
        return;
      }

      await userRef.set({
        aiTableUsage: admin.firestore.FieldValue.increment(1),
        lastAiTableAt: admin.firestore.FieldValue.serverTimestamp()
      }, { merge: true });

      res.json({
        title: result.title || "AI Generated Table",
        columns: Array.isArray(result.columns) ? result.columns : ["ID", "Name", "Email", "Status"],
        rows: Array.isArray(result.rows) ? result.rows : [],
        columnTypes: result.columnTypes || {},
        usage: {
          isPremium,
          used: currentUsage + 1,
          limit: isPremium ? null : FREE_AI_TABLE_LIMIT
        }
      });
    } catch (err) {
      console.error("aiTable error:", err);
      res.status(401).json({
        error: err.message || "Unauthorized"
      });
    }
  }
);