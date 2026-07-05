// Firebase 초기화 (재헌님 프로젝트 설정)
import { initializeApp } from "firebase/app";
import { getFirestore } from "firebase/firestore";

const firebaseConfig = {
  apiKey: "AIzaSyDhrOQeaJDUiL9foC8JhAOTJ9KN1UnZQzw",
  authDomain: "laundry-2579e.firebaseapp.com",
  projectId: "laundry-2579e",
  storageBucket: "laundry-2579e.firebasestorage.app",
  messagingSenderId: "506865715505",
  appId: "1:506865715505:web:a9e38e68ccf124ff9e4810",
};

const app = initializeApp(firebaseConfig);
export const db = getFirestore(app);
