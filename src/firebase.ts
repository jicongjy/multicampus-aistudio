import { initializeApp } from 'firebase/app';
import {
  getAuth,
  signInWithPopup,
  signOut,
  GoogleAuthProvider,
  onAuthStateChanged,
  User
} from 'firebase/auth';
import {
  getFirestore,
  doc,
  collection,
  getDoc,
  getDocs,
  setDoc,
  updateDoc,
  deleteDoc,
  writeBatch,
  onSnapshot,
  query,
  orderBy,
  getDocFromServer
} from 'firebase/firestore';
import firebaseConfig from '../firebase-applet-config.json';

export const SCOPES = [
  'https://www.googleapis.com/auth/spreadsheets',
  'https://www.googleapis.com/auth/drive.file',
];

export const FIRESTORE_DATABASE_ID = (firebaseConfig as any).firestoreDatabaseId || 'ai-studio-39914201-2301-45fe-8210-48667d5b6dea';

// Initialize Firebase
export const app = initializeApp(firebaseConfig);
export const db = getFirestore(app, FIRESTORE_DATABASE_ID);
export const auth = getAuth(app);

export const ADMIN_EMAIL = 'dangmujy@gmail.com';

const googleProvider = new GoogleAuthProvider();
SCOPES.forEach(scope => googleProvider.addScope(scope));
googleProvider.setCustomParameters({ prompt: 'select_account' });

let cachedAccessToken: string | null = null;
let isSigningIn = false;

export enum OperationType {
  CREATE = 'create',
  UPDATE = 'update',
  DELETE = 'delete',
  LIST = 'list',
  GET = 'get',
  WRITE = 'write',
}

export interface FirestoreErrorInfo {
  error: string;
  operationType: OperationType;
  path: string | null;
  authInfo: {
    userId?: string | null;
    email?: string | null;
    emailVerified?: boolean | null;
    isAnonymous?: boolean | null;
    tenantId?: string | null;
    providerInfo?: {
      providerId?: string | null;
      email?: string | null;
    }[];
  };
}

export function handleFirestoreError(error: unknown, operationType: OperationType, path: string | null): never {
  const errInfo: FirestoreErrorInfo = {
    error: error instanceof Error ? error.message : String(error),
    authInfo: {
      userId: auth.currentUser?.uid,
      email: auth.currentUser?.email,
      emailVerified: auth.currentUser?.emailVerified,
      isAnonymous: auth.currentUser?.isAnonymous,
      tenantId: auth.currentUser?.tenantId,
      providerInfo: auth.currentUser?.providerData?.map(provider => ({
        providerId: provider.providerId,
        email: provider.email,
      })) || []
    },
    operationType,
    path
  };
  console.error('Firestore Error: ', JSON.stringify(errInfo));
  throw new Error(JSON.stringify(errInfo));
}

// Test connection on boot
export async function testConnection() {
  try {
    await getDocFromServer(doc(db, 'test', 'connection'));
  } catch (error) {
    if (error instanceof Error && error.message.includes('the client is offline')) {
      console.error('Please check your Firebase configuration.');
    }
  }
}
testConnection();

export interface RegistrationRecord {
  id: number;
  docId: string;
  name: string;
  phone: string;
  phoneDigits: string;
  email: string;
  count: number;
  memo: string;
  mkt: boolean;
  agree: boolean;
  agreedAt: string | null;
  createdAt: string;
  updatedAt: string | null;
}

export interface RegistrationInput {
  name: string;
  phone: string;
  email: string;
  count: number;
  memo?: string;
  mkt?: boolean;
  agree?: boolean;
}

function normalizePhone(phone: string): string {
  return phone.replace(/\D/g, '');
}

/**
 * Check if a phone number is already registered using the public phone_index
 */
export async function checkPhoneTaken(phone: string, exceptDocId?: string | null): Promise<boolean> {
  const digits = normalizePhone(phone);
  if (!digits) return false;
  if (exceptDocId && exceptDocId === `reg_${digits}`) return false;

  const path = `phone_index/${digits}`;
  try {
    const snap = await getDoc(doc(db, 'phone_index', digits));
    return snap.exists();
  } catch (err) {
    handleFirestoreError(err, OperationType.GET, path);
  }
}

/**
 * Create a new event registration in Firebase Firestore
 */
export async function createRegistration(v: RegistrationInput, nextSeq: number): Promise<RegistrationRecord> {
  const digits = normalizePhone(v.phone);
  const docId = `reg_${digits}`;
  const now = new Date().toISOString();

  const record: RegistrationRecord = {
    id: nextSeq,
    docId,
    name: v.name.trim(),
    phone: v.phone.trim(),
    phoneDigits: digits,
    email: v.email.trim().toLowerCase(),
    count: Number(v.count) || 1,
    memo: (v.memo || '').trim(),
    mkt: !!v.mkt,
    agree: true,
    agreedAt: now,
    createdAt: now,
    updatedAt: null,
  };

  const batch = writeBatch(db);
  batch.set(doc(db, 'registrations', docId), record);
  batch.set(doc(db, 'phone_index', digits), { registered: true, createdAt: now });

  try {
    await batch.commit();
    return record;
  } catch (err) {
    handleFirestoreError(err, OperationType.WRITE, `registrations/${docId}`);
  }
}

/**
 * Update an existing registration (Admin only)
 */
export async function updateRegistration(
  docId: string,
  v: RegistrationInput,
  oldRecord: RegistrationRecord
): Promise<void> {
  const now = new Date().toISOString();
  const newDigits = normalizePhone(v.phone);
  const oldDigits = oldRecord.phoneDigits || normalizePhone(oldRecord.phone);

  const updatedData: Partial<RegistrationRecord> = {
    name: v.name.trim(),
    phone: v.phone.trim(),
    phoneDigits: newDigits,
    email: v.email.trim().toLowerCase(),
    count: Number(v.count) || 1,
    memo: (v.memo || '').trim(),
    mkt: !!v.mkt,
    agree: !!v.agree,
    updatedAt: now,
  };

  if (v.agree && !oldRecord.agree) {
    updatedData.agreedAt = now;
  } else if (!v.agree) {
    updatedData.agreedAt = null;
  }

  const batch = writeBatch(db);

  // If phone changed, update phone_index
  if (newDigits !== oldDigits) {
    batch.delete(doc(db, 'phone_index', oldDigits));
    batch.set(doc(db, 'phone_index', newDigits), { registered: true, createdAt: now });
  }

  batch.update(doc(db, 'registrations', docId), updatedData);

  try {
    await batch.commit();
  } catch (err) {
    handleFirestoreError(err, OperationType.UPDATE, `registrations/${docId}`);
  }
}

/**
 * Delete a registration (Admin only)
 */
export async function deleteRegistration(docId: string, phone: string): Promise<void> {
  const digits = normalizePhone(phone);
  const batch = writeBatch(db);

  batch.delete(doc(db, 'registrations', docId));
  if (digits) {
    batch.delete(doc(db, 'phone_index', digits));
  }

  try {
    await batch.commit();
  } catch (err) {
    handleFirestoreError(err, OperationType.DELETE, `registrations/${docId}`);
  }
}

/**
 * Clear all registrations (Admin only)
 */
export async function clearAllRegistrations(records: RegistrationRecord[]): Promise<void> {
  const batch = writeBatch(db);
  for (const r of records) {
    batch.delete(doc(db, 'registrations', r.docId));
    const digits = r.phoneDigits || normalizePhone(r.phone);
    if (digits) {
      batch.delete(doc(db, 'phone_index', digits));
    }
  }

  try {
    await batch.commit();
  } catch (err) {
    handleFirestoreError(err, OperationType.DELETE, 'registrations');
  }
}

/**
 * Real-time listener for registrations (Admin only)
 */
export function subscribeToRegistrations(
  onData: (records: RegistrationRecord[]) => void,
  onError: (error: unknown) => void
): () => void {
  const isAuthAdmin = auth.currentUser && auth.currentUser.email?.toLowerCase() === ADMIN_EMAIL.toLowerCase();
  if (!isAuthAdmin) {
    console.warn('subscribeToRegistrations called without admin auth. Skipping subscription.');
    return () => {};
  }

  const q = query(collection(db, 'registrations'), orderBy('createdAt', 'desc'));
  return onSnapshot(
    q,
    snapshot => {
      const records: RegistrationRecord[] = [];
      snapshot.forEach(docSnap => {
        const d = docSnap.data() as RegistrationRecord;
        records.push({
          ...d,
          docId: docSnap.id,
        });
      });
      onData(records);
    },
    error => {
      console.warn('Real-time registrations subscription issue:', error);
      onError(error);
    }
  );
}

/**
 * Fetch all registrations once (Admin only)
 */
export async function fetchAllRegistrations(): Promise<RegistrationRecord[]> {
  const isAuthAdmin = auth.currentUser && auth.currentUser.email?.toLowerCase() === ADMIN_EMAIL.toLowerCase();
  if (!isAuthAdmin) {
    console.warn('fetchAllRegistrations called without admin auth. Returning empty list.');
    return [];
  }

  try {
    const q = query(collection(db, 'registrations'), orderBy('createdAt', 'desc'));
    const snap = await getDocs(q);
    const records: RegistrationRecord[] = [];
    snap.forEach(docSnap => {
      records.push({
        ...(docSnap.data() as RegistrationRecord),
        docId: docSnap.id,
      });
    });
    return records;
  } catch (err) {
    handleFirestoreError(err, OperationType.LIST, 'registrations');
  }
}

/**
 * Sign in as admin with Google and retrieve Workspace OAuth access token
 */
export async function signInAdminWithGoogle(): Promise<{ user: User; accessToken: string | null }> {
  try {
    isSigningIn = true;
    const cred = await signInWithPopup(auth, googleProvider);
    const googleCred = GoogleAuthProvider.credentialFromResult(cred);
    cachedAccessToken = googleCred?.accessToken || null;
    return { user: cred.user, accessToken: cachedAccessToken };
  } finally {
    isSigningIn = false;
  }
}

export function getCachedAccessToken(): string | null {
  return cachedAccessToken;
}

export function setCachedAccessToken(token: string | null): void {
  cachedAccessToken = token;
}

/**
 * Sign out admin and clear cached token
 */
export async function signOutAdmin(): Promise<void> {
  cachedAccessToken = null;
  await signOut(auth);
}

/**
 * Listen to auth state
 */
export function listenAuthState(callback: (user: User | null, isAdmin: boolean) => void): () => void {
  return onAuthStateChanged(auth, user => {
    if (!user) {
      cachedAccessToken = null;
    }
    const isAdminUser = !!user && user.email?.toLowerCase() === ADMIN_EMAIL.toLowerCase();
    callback(user, isAdminUser);
  });
}

export interface GoogleSheetsConfig {
  spreadsheetId: string;
  spreadsheetUrl: string;
  sheetTitle: string;
  autoSync: boolean;
  updatedAt: string;
}

export async function getSheetsConfig(): Promise<GoogleSheetsConfig | null> {
  try {
    const snap = await getDoc(doc(db, 'settings', 'google_sheets'));
    if (snap.exists()) {
      return snap.data() as GoogleSheetsConfig;
    }
    return null;
  } catch (e) {
    console.warn('Failed to load sheets config:', e);
    return null;
  }
}

export async function saveSheetsConfig(config: Partial<GoogleSheetsConfig> | null): Promise<void> {
  const ref = doc(db, 'settings', 'google_sheets');
  if (config === null) {
    await deleteDoc(ref);
  } else {
    await setDoc(
      ref,
      {
        ...config,
        updatedAt: new Date().toISOString(),
      },
      { merge: true }
    );
  }
}
