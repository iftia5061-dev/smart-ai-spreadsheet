import { initializeApp } from "firebase/app";
import {
  getAuth,
  GoogleAuthProvider,
  signInWithPopup,
  signOut,
  onAuthStateChanged,
} from "firebase/auth";
import {
  getFirestore,
  doc,
  setDoc,
  getDoc,
  updateDoc,
  serverTimestamp,
} from "firebase/firestore";

// ============================================================
// FIREBASE CONFIG
// ============================================================
const firebaseConfig = {
  apiKey: "AIzaSyBFqaj-vl_W0sBAde_-XdbJPtKYI_tA3Wk",
  authDomain: "smart-sheet-pro.firebaseapp.com",
  projectId: "smart-sheet-pro",
  storageBucket: "smart-sheet-pro.firebasestorage.app",
  messagingSenderId: "1005248265803",
  appId: "1:1005248265803:web:5b173115e8198633cdd7ce",
  measurementId: "G-NCY4VFHVV5",
};

// ============================================================
// FIREBASE INIT
// ============================================================
const app = initializeApp(firebaseConfig);
export const auth = getAuth(app);
export const db = getFirestore(app);
export const provider = new GoogleAuthProvider();
export { signInWithPopup, signOut, onAuthStateChanged };

// ============================================================
// SUBSCRIPTION HELPERS
// ============================================================

// Get user subscription status from Firestore
export const getUserStatus = async (uid) => {
  try {
    const ref = doc(db, "users", uid);
    const snap = await getDoc(ref);
    if (snap.exists()) {
      return snap.data();
    }
    return null;
  } catch (e) {
    console.error("getUserStatus error:", e);
    return null;
  }
};

// Create new user profile on first login
export const createUserProfile = async (user) => {
  try {
    const ref = doc(db, "users", user.uid);
    const snap = await getDoc(ref);
    if (!snap.exists()) {
      await setDoc(ref, {
        uid: user.uid,
        email: user.email,
        displayName: user.displayName,
        photoURL: user.photoURL,
        plan: "free",           // "free" or "pro"
        isPremium: false,
        joinedAt: serverTimestamp(),
        lastLogin: serverTimestamp(),
        tableCount: 0,
        paymentStatus: null,    // null, "pending", "approved"
        paymentRef: null,
      });
    } else {
      // Update last login
      await updateDoc(ref, {
        lastLogin: serverTimestamp(),
      });
    }
  } catch (e) {
    console.error("createUserProfile error:", e);
  }
};

// Check if user is premium
export const checkIsPremium = async (uid) => {
  try {
    const data = await getUserStatus(uid);
    if (!data) return false;
    return data.isPremium === true && data.plan === "pro";
  } catch (e) {
    return false;
  }
};

// Submit bKash payment request (pending admin approval)
export const submitPaymentRequest = async (uid, txnId, amount) => {
  try {
    const ref = doc(db, "users", uid);
    await updateDoc(ref, {
      paymentStatus: "pending",
      paymentRef: txnId,
      paymentAmount: amount,
      paymentSubmittedAt: serverTimestamp(),
    });

    // Also save to payments collection for admin
    const payRef = doc(db, "payments", `${uid}_${Date.now()}`);
    await setDoc(payRef, {
      uid,
      txnId,
      amount,
      status: "pending",
      submittedAt: serverTimestamp(),
    });

    return { success: true };
  } catch (e) {
    console.error("submitPaymentRequest error:", e);
    return { success: false, error: e.message };
  }
};

// Admin: approve user payment → upgrade to Pro
export const approveUserPayment = async (uid) => {
  try {
    const ref = doc(db, "users", uid);
    await updateDoc(ref, {
      plan: "pro",
      isPremium: true,
      paymentStatus: "approved",
      upgradedAt: serverTimestamp(),
    });
    return { success: true };
  } catch (e) {
    console.error("approveUserPayment error:", e);
    return { success: false };
  }
};