import React from "react";
import { NavLink, Navigate, Route, Routes, useLocation } from "react-router-dom";

import JobListPage from "./pages/JobListPage.jsx";
import UploadPage from "./pages/UploadPage.jsx";
import JobEditorPage from "./pages/JobEditorPage.jsx";
import DiffPage from "./pages/DiffPage.jsx";
import LectureSearchGuidePage from "./pages/LectureSearchGuidePage.jsx";
import FolderZipPage from "./pages/FolderZipPage.jsx";
import PdfDocumentZipPage from "./pages/PdfDocumentZipPage.jsx";
import PdfSlideZipPage from "./pages/PdfSlideZipPage.jsx";
import PptVideoZipPage from "./pages/PptVideoZipPage.jsx";
import "./styles/shell.css";

function AppShell({ children }) {
  const location = useLocation();
  const isFragmentActive =
    location.pathname === "/" ||
    location.pathname === "/fragment" ||
    location.pathname.startsWith("/upload") ||
    location.pathname.startsWith("/job/") ||
    location.pathname.startsWith("/diff");

  return (
    <div className="app-shell">
      <aside className="app-sidebar" aria-label="ツールナビゲーション">
        <div className="app-sidebar__brand">
          <div className="app-sidebar__eyebrow">tools</div>
          <div className="app-sidebar__title">制作支援ツール</div>
        </div>

        <nav className="app-sidebar__nav">
          <section className="app-sidebar__group">
            <h2 className="app-sidebar__group-title">M社</h2>
            <div className="app-sidebar__group-links">
              <NavLink
                to="/"
                className={() =>
                  `app-sidebar__link${isFragmentActive ? " app-sidebar__link--active" : ""}`
                }
              >
                <span className="app-sidebar__link-main">フラグメントデザイン</span>
                <span className="app-sidebar__link-sub">デザイン生成・シート比較</span>
              </NavLink>

              <NavLink
                to="/lecture-search-guide"
                className={({ isActive }) =>
                  `app-sidebar__link${isActive ? " app-sidebar__link--active" : ""}`
                }
              >
                <span className="app-sidebar__link-main">講演会検索案内ツール</span>
                <span className="app-sidebar__link-sub">ZIP生成・Vault登録</span>
              </NavLink>
            </div>
          </section>

          <section className="app-sidebar__group">
            <h2 className="app-sidebar__group-title">I社</h2>
            <div className="app-sidebar__group-links">
              <NavLink
                to="/shared-slide-zip"
                className={({ isActive }) =>
                  `app-sidebar__link${isActive ? " app-sidebar__link--active" : ""}`
                }
              >
                <span className="app-sidebar__link-main">Shared付きスライド</span>
                <span className="app-sidebar__link-sub">ZIP生成・タイトル抽出</span>
              </NavLink>
            </div>
          </section>

          <section className="app-sidebar__group">
            <h2 className="app-sidebar__group-title">全社共通</h2>
            <div className="app-sidebar__group-links">
              <NavLink
                to="/pdf-document-zip"
                className={({ isActive }) =>
                  `app-sidebar__link${isActive ? " app-sidebar__link--active" : ""}`
                }
              >
                <span className="app-sidebar__link-main">PDFスライド</span>
                <span className="app-sidebar__link-sub">pdf.pdf形式ZIP生成</span>
              </NavLink>

              <NavLink
                to="/folder-zip"
                className={({ isActive }) =>
                  `app-sidebar__link${isActive ? " app-sidebar__link--active" : ""}`
                }
              >
                <span className="app-sidebar__link-main">フォルダZIP生成</span>
                <span className="app-sidebar__link-sub">レガシーorモダン自動判定</span>
              </NavLink>

              <NavLink
                to="/ppt-video-zip"
                className={({ isActive }) =>
                  `app-sidebar__link${isActive ? " app-sidebar__link--active" : ""}`
                }
              >
                <span className="app-sidebar__link-main">PPT動画ZIP生成</span>
                <span className="app-sidebar__link-sub">動画ページ抽出・ZIP生成</span>
              </NavLink>
            </div>
          </section>
        </nav>
      </aside>

      <main className="app-main">{children}</main>
    </div>
  );
}

export default function App() {
  return (
    <AppShell>
      <Routes>
        {/* フラグメントデザイン */}
        <Route path="/" element={<JobListPage />} />
        <Route path="/fragment" element={<JobListPage />} />

        {/* アップロード */}
        <Route path="/upload" element={<UploadPage />} />

        {/* 編集 */}
        <Route path="/job/:jobId" element={<JobEditorPage />} />

        <Route path="/diff" element={<DiffPage />} />

        {/* 講演会検索案内ツール */}
        <Route path="/lecture-search-guide" element={<LectureSearchGuidePage />} />

        {/* Shared付きスライドZIP生成 */}
        <Route path="/shared-slide-zip" element={<PdfSlideZipPage />} />

        {/* PDF資料ZIP生成 */}
        <Route path="/pdf-document-zip" element={<PdfDocumentZipPage />} />

        {/* フォルダZIP生成 */}
        <Route path="/folder-zip" element={<FolderZipPage />} />

        {/* PPT動画ZIP生成 */}
        <Route path="/ppt-video-zip" element={<PptVideoZipPage />} />

        {/* それ以外は一覧へ */}
        <Route path="*" element={<Navigate to="/" replace />} />
      </Routes>
    </AppShell>
  );
}
