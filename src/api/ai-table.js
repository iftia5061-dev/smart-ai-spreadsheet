const admin = require("firebase-admin");

const FREE_AI_TABLE_LIMIT = 2;
const OPENROUTER_API_KEY = process.env.OPENROUTER_API_KEY || "";
const MODEL = "google/gemini-2.5-flash";

function initAdmin() {
  if (admin.apps.length) return;
  const raw = process.env.FIREBASE_SERVICE_ACCOUNT;
  if (!raw) throw new Error("Missing FIREBASE_SERVICE_ACCOUNT env variable");
  const serviceAccount = JSON.parse(raw);
  if (serviceAccount.private_key) {
    serviceAccount.private_key = serviceAccount.private_key.replace(/\\n/g, "\n");
  }
  admin.initializeApp({ credential: admin.credential.cert(serviceAccount) });
}

function setCors(res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");
}

function extractJson(text) {
  const cleaned = String(text || "").replace(/```json/g, "").replace(/```/g, "").trim();
  const start = cleaned.indexOf("{");
  const end = cleaned.lastIndexOf("}");
  if (start === -1 || end === -1) throw new Error("AI did not return JSON");
  return JSON.parse(cleaned.slice(start, end + 1));
}

async function verifyUser(req) {
  const authHeader = req.headers.authorization || "";
  if (!authHeader.startsWith("Bearer ")) throw new Error("Missing auth token");
  const idToken = authHeader.replace("Bearer ", "");
  return admin.auth().verifyIdToken(idToken);
}

function isUserPremium(userData, email) {
  if (email === "iftia5061@gmail.com") return true;
  if (!userData?.isPremium) return false;
  const expiry = userData.premiumExpiry;
  if (!expiry) return true;
  if (typeof expiry.toDate === "function") return expiry.toDate() > new Date();
  if (expiry.seconds) return new Date(expiry.seconds * 1000) > new Date();
  return true;
}

async function callOpenRouter(prompt) {
  const systemPrompt = `You are an AI table generator for a premium spreadsheet app.
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
- Create useful business/data-entry tables.
- If user asks for a row count, follow it.
- If no row count is given, create 8 rows.
- Keep columns practical, clean, and suitable for spreadsheet work.
- Return valid JSON only.`;

  const response = await fetch("https://openrouter.ai/api/v1/chat/completions", {
    method: "POST",
    headers: {
      "Authorization": `Bearer ${OPENROUTER_API_KEY}`,
      "Content-Type": "application/json",
      "HTTP-Referer": "https://sheetminds.xyz",
      "X-Title": "SheetMind",
    },
    body: JSON.stringify({
      model: MODEL,
      messages: [
        { role: "system", content: systemPrompt },
        { role: "user", content: prompt },
      ],
      temperature: 0.35,
    }),
  });

  const data = await response.json();

  if (!response.ok) {
    const message = data?.error?.message || "OpenRouter request failed";
    const error = new Error(message);
    error.status = response.status;
    throw error;
  }

  const text = data?.choices?.[0]?.message?.content || "";
  return extractJson(text);
}

module.exports = async function handler(req, res) {
  setCors(res);

  if (req.method === "OPTIONS") { res.status(204).send(""); return; }
  if (req.method !== "POST") { res.status(405).json({ error: "Method not allowed" }); return; }

  try {
    initAdmin();

    const decoded = await verifyUser(req);
    const uid = decoded.uid;
    const email = decoded.email || "";

    const prompt = String(req.body?.prompt || "").trim();
    if (!prompt) { res.status(400).json({ error: "Prompt is required" }); return; }

    const userRef = admin.firestore().doc(`users/${uid}`);
    const userSnap = await userRef.get();
    const userData = userSnap.exists ? userSnap.data() : {};

    const premium = isUserPremium(userData, email);
    const currentUsage = Number(userData.aiTableUsage || 0);

    if (!premium && currentUsage >= FREE_AI_TABLE_LIMIT) {
      res.status(402).json({
        code: "AI_FREE_LIMIT_REACHED",
        error: `Free plan allows ${FREE_AI_TABLE_LIMIT} AI tables. Upgrade for unlimited AI tables.`,
      });
      return;
    }

    if (!OPENROUTER_API_KEY) {
      res.status(500).json({ error: "No API key configured" });
      return;
    }

    const result = await callOpenRouter(prompt);

    await userRef.set({
      aiTableUsage: admin.firestore.FieldValue.increment(1),
      lastAiTableAt: admin.firestore.FieldValue.serverTimestamp(),
    }, { merge: true });

    res.status(200).json({
      title: result.title || "AI Generated Table",
      columns: Array.isArray(result.columns) ? result.columns : ["ID", "Name", "Email", "Status"],
      rows: Array.isArray(result.rows) ? result.rows : [],
      columnTypes: result.columnTypes || {},
      usage: {
        isPremium: premium,
        used: currentUsage + 1,
        limit: premium ? null : FREE_AI_TABLE_LIMIT,
      },
    });
  } catch (err) {
    console.error("AI table API error:", err);
    res.status(401).json({ code: "UNAUTHORIZED", error: err.message || "Unauthorized" });
  }
};