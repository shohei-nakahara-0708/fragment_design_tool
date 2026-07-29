import React from "react";
import { NavLink, Navigate, Route, Routes, useLocation } from "react-router-dom";

import JobListPage from "./pages/JobListPage.jsx";
import UploadPage from "./pages/UploadPage.jsx";
import JobEditorPage from "./pages/JobEditorPage.jsx";
import DiffPage from "./pages/DiffPage.jsx";
import LectureSearchGuidePage from "./pages/LectureSearchGuidePage.jsx";
import StampMailToolPage from "./pages/StampMailToolPage.jsx";
import BannerMailToolPage from "./pages/BannerMailToolPage.jsx";
import CslLectureToolPage from "./pages/CslLectureToolPage.jsx";
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

          <NavLink
            to="/stamp-mail-tool"
            className={({ isActive }) =>
              `app-sidebar__link${isActive ? " app-sidebar__link--active" : ""} sidebar__link--none`
            }
          >
            <span className="app-sidebar__link-main">CSLスタンプメール生成</span>
            <span className="app-sidebar__link-sub">HTML・images.zip生成</span>
          </NavLink>

          <NavLink
            to="/banner-mail-tool"
            className={({ isActive }) =>
              `app-sidebar__link${isActive ? " app-sidebar__link--active" : ""} sidebar__link--none`
            }
          >
            <span className="app-sidebar__link-main">CSLバナーメール生成</span>
            <span className="app-sidebar__link-sub">bannerタブ・images.zip生成</span>
          </NavLink>

          <NavLink
            to="/csl-lecture-tool"
            className={({ isActive }) =>
              `app-sidebar__link${isActive ? " app-sidebar__link--active" : ""} sidebar__link--none`
            }
          >
            <span className="app-sidebar__link-main">CSL講演会ツール生成</span>
            <span className="app-sidebar__link-sub">講演会ツールタブ・ZIP生成</span>
          </NavLink>
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

        {/* CSLスタンプメール生成 */}
        <Route path="/stamp-mail-tool" element={<StampMailToolPage />} />

        {/* CSLバナーメール生成 */}
        <Route path="/banner-mail-tool" element={<BannerMailToolPage />} />

        {/* CSL講演会ツール生成 */}
        <Route path="/csl-lecture-tool" element={<CslLectureToolPage />} />

        {/* それ以外は一覧へ */}
        <Route path="*" element={<Navigate to="/" replace />} />
      </Routes>
    </AppShell>
  );
}
