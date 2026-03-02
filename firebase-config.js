import { initializeApp } from "https://www.gstatic.com/firebasejs/10.12.5/firebase-app.js";
import {
  getDatabase,
  ref,
  onValue,
  set,
  update,
  remove,
  runTransaction,
} from "https://www.gstatic.com/firebasejs/10.12.5/firebase-database.js";

const firebaseConfig = {
  apiKey: "AIzaSyBK-4VEMXFiWk2rJ8gGUJRRNKDQDyi_5I",
  authDomain: "cartografia-9ca7b.firebaseapp.com",
  databaseURL: "https://cartografia-9ca7b-default-rtdb.firebaseio.com",
  projectId: "cartografia-9ca7b",
  storageBucket: "cartografia-9ca7b.firebasestorage.app",
  messagingSenderId: "950518445545",
  appId: "1:950518445545:web:1616cb2587b32c10c6c778",
};

const app = initializeApp(firebaseConfig);
const db = getDatabase(app);

window.firebaseApp = app;
window.db = db;

// seu app escuta essa coleção:
window.dbRef = ref(db, "solicitacoes");

// “firebaseFunctions” = funções RTDB que seu script.js usa
window.firebaseFunctions = { ref, onValue, set, update, remove, runTransaction };

console.log("✅ Firebase OK:", firebaseConfig.projectId);