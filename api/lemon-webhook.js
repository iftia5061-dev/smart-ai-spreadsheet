const crypto = require("crypto");
const admin = require("firebase-admin");

const ADMIN_EMAIL = "iftia5061@gmail.com";

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

function getRawBody(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    req.on("data", (chunk) => chunks.push(Buffer.from(chunk)));
    req.on("end", () => resolve(Buffer.concat(chunks)));
    req.on("error", reject);
  });
}

function verifySignature(rawBody, signature) {
  const secret = process.env.LEMONSQUEEZY_WEBHOOK_SECRET;
  if (!secret) throw new Error("Missing LEMONSQUEEZY_WEBHOOK_SECRET env variable");
  if (!signature) throw new Error("Missing Lemon Squeezy signature");

  const digest = crypto.createHmac("sha256", secret).update(rawBody).digest("hex");
  const expected = Buffer.from(digest, "utf8");
  const actual = Buffer.from(signature, "utf8");

  if (expected.length !== actual.length || !crypto.timingSafeEqual(expected, actual)) {
    throw new Error("Invalid Lemon Squeezy signature");
  }
}

function addMonths(date, months) {
  const next = new Date(date);
  next.setMonth(next.getMonth() + months);
  return next;
}

function getText(...values) {
  return values.filter(Boolean).map((value) => String(value).toLowerCase()).join(" ");
}

function getPlan(payload) {
  const eventName = payload?.meta?.event_name || "";
  const attrs = payload?.data?.attributes || {};
  const item = attrs.first_order_item || {};
  const text = getText(eventName, attrs.product_name, attrs.variant_name, item.product_name, item.variant_name, attrs.urls?.receipt);

  if (text.includes("year")) return { plan: "yearly", months: 12, source: "lemon" };
  if (text.includes("desktop") || text.includes("windows")) return { plan: "desktop", months: 1200, source: "lemon" };
  if (text.includes("phone") || text.includes("android") || text.includes("mobile")) return { plan: "phone", months: 1200, source: "lemon" };
  if (eventName.includes("subscription")) return { plan: "monthly", months: 1, source: "lemon" };
  return { plan: "monthly", months: 1, source: "lemon" };
}

function getBuyer(payload) {
  const custom = payload?.meta?.custom_data || {};
  const attrs = payload?.data?.attributes || {};
  const email = String(custom.email || custom.user_email || attrs.user_email || attrs.customer_email || "").trim().toLowerCase();
  const uid = String(custom.uid || custom.userId || custom.user_id || "").trim();
  return { email, uid };
}

async function activatePremium({ email, uid, payload }) {
  const db = admin.firestore();
  const now = new Date();
  const plan = getPlan(payload);
  const expiry = addMonths(now, plan.months);
  const eventName = payload?.meta?.event_name || "unknown";
  const orderId = payload?.data?.id || "";

  const premiumData = {
    isPremium: true,
    plan: "pro",
    premiumPlan: plan.plan,
    premiumExpiry: admin.firestore.Timestamp.fromDate(expiry),
    premiumGrantedAt: admin.firestore.FieldValue.serverTimestamp(),
    premiumGrantedBy: "lemon_webhook",
    paymentStatus: "approved",
    paymentProvider: "lemon_squeezy",
    lemonEvent: eventName,
    lemonOrderId: String(orderId),
  };

  let resolvedUid = uid;
  if (!resolvedUid && email) {
    try {
      const user = await admin.auth().getUserByEmail(email);
      resolvedUid = user.uid;
    } catch (err) {
      resolvedUid = "";
    }
  }

  if (resolvedUid) {
    await db.doc(`users/${resolvedUid}`).set({ ...premiumData, email }, { merge: true });
    await db.doc(`spreadsheets/${resolvedUid}`).set({ isPremium: true }, { merge: true });
  }

  if (email) {
    await db.doc(`premium_emails/${email}`).set({
      ...premiumData,
      email,
      uid: resolvedUid || "",
      updatedAt: admin.firestore.FieldValue.serverTimestamp(),
    }, { merge: true });
  }

  return { uid: resolvedUid, email, plan: plan.plan };
}

module.exports = async function handler(req, res) {
  if (req.method !== "POST") {
    res.status(405).json({ error: "Method not allowed" });
    return;
  }

  try {
    initAdmin();
    const rawBody = await getRawBody(req);
    verifySignature(rawBody, req.headers["x-signature"] || "");

    const payload = JSON.parse(rawBody.toString("utf8"));
    const eventName = payload?.meta?.event_name || req.headers["x-event-name"] || "";
    const allowedEvents = new Set([
      "order_created",
      "subscription_created",
      "subscription_payment_success",
      "subscription_resumed",
    ]);

    if (!allowedEvents.has(eventName)) {
      res.status(200).json({ ok: true, ignored: eventName });
      return;
    }

    const buyer = getBuyer(payload);
    if (!buyer.email && !buyer.uid) {
      res.status(400).json({ error: "No buyer email or uid found in Lemon payload" });
      return;
    }

    const result = await activatePremium({ ...buyer, payload });
    res.status(200).json({ ok: true, ...result });
  } catch (err) {
    console.error("Lemon webhook error:", err);
    res.status(400).json({ error: err.message || "Webhook failed" });
  }
};