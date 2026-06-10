const admin = require("firebase-admin");

const FREE_AI_TABLE_LIMIT = 2;
const ADMIN_EMAIL = "iftia5061@gmail.com";

function setCors(res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");
}

function initAdmin() {
  if (admin.apps.length) return;

  const raw = process.env.FIREBASE_SERVICE_ACCOUNT;
  if (!raw) throw new Error("Missing FIREBASE_SERVICE_ACCOUNT env variable");

  const serviceAccount = JSON.parse(raw);
  if (serviceAccount.private_key) {
    serviceAccount.private_key = serviceAccount.private_key.replace(/\\n/g, "\n");
  }

  admin.initializeApp({
    credential: admin.credential.cert(serviceAccount),
  });
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

function isPremiumUser(userData, email) {
  if (email === ADMIN_EMAIL) return true;
  if (!userData || !userData.isPremium) return false;

  const expiry = userData.premiumExpiry;
  if (!expiry) return true;
  if (typeof expiry.toDate === "function") return expiry.toDate() > new Date();
  if (expiry.seconds) return new Date(expiry.seconds * 1000) > new Date();
  return true;
}

async function callOpenRouter(userPrompt) {
  const apiKey = process.env.OPENROUTER_API_KEY;
  if (!apiKey) {
    const err = new Error("No OpenRouter API key configured in Vercel Production Environment Variables");
    err.status = 500;
    throw err;
  }

  const systemPrompt = `
You are an AI table generator for a premium spreadsheet/data-entry app.
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
- Make practical company/freelancer data-entry tables.
- If user asks for row count, follow it. Otherwise create 8 rows.
- Keep column names clean and spreadsheet friendly.
- Return valid JSON only.
`;

  const response = await fetch("https://openrouter.ai/api/v1/chat/completions", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "Authorization": `Bearer ${apiKey}`,
      "HTTP-Referer": "https://smart-ai-two-phi.vercel.app",
      "X-Title": "Smart AI Table",
    },
    body: JSON.stringify({
      model: "google/gemini-flash-1.5",
      messages: [
        { role: "system", content: systemPrompt },
        { role: "user", content: userPrompt },
      ],
      temperature: 0.35,
    }),
  });

  const data = await response.json();
  if (!response.ok) {
    const err = new Error(data?.error?.message || "OpenRouter request failed");
    err.status = response.status;
    throw err;
  }

  const text = data?.choices?.[0]?.message?.content || "";
  return extractJson(text);
}

module.exports = async function handler(req, res) {
  setCors(res);

  if (req.method === "OPTIONS") {
    res.status(204).send("");
    return;
  }

  if (req.method !== "POST") {
    res.status(405).json({ error: "Method not allowed" });
    return;
  }

  try {
    initAdmin();

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
    const premium = isPremiumUser(userData, email);
    const currentUsage = Number(userData.aiTableUsage || 0);

    if (!premium && currentUsage >= FREE_AI_TABLE_LIMIT) {
      res.status(402).json({
        code: "AI_FREE_LIMIT_REACHED",
        error: `Free plan allows ${FREE_AI_TABLE_LIMIT} AI tables. Upgrade for unlimited AI tables.`,
      });
      return;
    }

    let result = null;
    let lastError = null;

    try {
      result = await callOpenRouter(prompt);
    } catch (err) {
      lastError = err;
    }

    if (!result) {
      console.error("OpenRouter generation failed", { status: lastError?.status, message: lastError?.message });
      const quotaHit = lastError?.status === 429 || lastError?.status === 403;
      res.status(quotaHit ? 429 : 500).json({
        code: quotaHit ? "AI_QUOTA_REACHED" : "AI_FAILED",
        error: quotaHit
          ? "AI quota limit reached today. Try again later."
          : (lastError?.message || "AI table generation failed."),
      });
      return;
    }

    await userRef.set({
      aiTableUsage: admin.firestore.FieldValue.increment(1),
      lastAiTableAt: admin.firestore.FieldValue.serverTimestamp(),
    }, { merge: true });

    res.status(200).json({
      title: result.title || "AI Generated Table",
      columns: Array.isArray(result.columns) && result.columns.length
        ? result.columns.map(String)
        : ["ID", "Name", "Email", "Status"],
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