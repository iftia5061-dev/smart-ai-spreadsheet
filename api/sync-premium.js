const admin = require("firebase-admin");

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

function isActivePremium(data) {
  if (!data || !data.isPremium) return false;
  const expiry = data.premiumExpiry;
  if (!expiry) return true;
  if (typeof expiry.toDate === "function") return expiry.toDate() > new Date();
  if (expiry.seconds) return new Date(expiry.seconds * 1000) > new Date();
  return true;
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

    const authHeader = req.headers.authorization || "";
    if (!authHeader.startsWith("Bearer ")) {
      res.status(401).json({ error: "Missing auth token" });
      return;
    }

    const decoded = await admin.auth().verifyIdToken(authHeader.replace("Bearer ", ""));
    const uid = decoded.uid;
    const email = String(decoded.email || "").toLowerCase();

    if (email === ADMIN_EMAIL) {
      res.status(200).json({ isPremium: true, admin: true });
      return;
    }

    const db = admin.firestore();
    const userRef = db.doc(`users/${uid}`);
    const userSnap = await userRef.get();
    const userData = userSnap.exists ? userSnap.data() : {};

    if (isActivePremium(userData)) {
      res.status(200).json({ isPremium: true, source: "users" });
      return;
    }

    if (!email) {
      res.status(200).json({ isPremium: false });
      return;
    }

    const paidSnap = await db.doc(`premium_emails/${email}`).get();
    const paidData = paidSnap.exists ? paidSnap.data() : null;

    if (!isActivePremium(paidData)) {
      res.status(200).json({ isPremium: false });
      return;
    }

    const premiumData = {
      isPremium: true,
      plan: "pro",
      premiumPlan: paidData.premiumPlan || "monthly",
      premiumExpiry: paidData.premiumExpiry || null,
      premiumGrantedAt: admin.firestore.FieldValue.serverTimestamp(),
      premiumGrantedBy: "lemon_sync",
      paymentStatus: "approved",
      paymentProvider: "lemon_squeezy",
    };

    await userRef.set({ ...premiumData, email, uid }, { merge: true });
    await db.doc(`spreadsheets/${uid}`).set({ isPremium: true }, { merge: true });
    await db.doc(`premium_emails/${email}`).set({ uid }, { merge: true });

    res.status(200).json({ isPremium: true, source: "premium_emails" });
  } catch (err) {
    console.error("Premium sync error:", err);
    res.status(500).json({ error: err.message || "Premium sync failed" });
  }
};