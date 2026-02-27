import React from "react";
import { Routes, Route, Navigate } from "react-router-dom";

import JobListPage from "./pages/JobListPage.jsx";
import UploadPage from "./pages/UploadPage.jsx";
import JobEditorPage from "./pages/JobEditorPage.jsx";

export default function App() {
  return (
    <Routes>
      {/* 一覧 */}
      <Route path="/" element={<JobListPage />} />

      {/* アップロード */}
      <Route path="/upload" element={<UploadPage />} />

      {/* 編集 */}
      <Route path="/job/:jobId" element={<JobEditorPage />} />

      {/* それ以外は一覧へ */}
      <Route path="*" element={<Navigate to="/" replace />} />
    </Routes>
  );
}
