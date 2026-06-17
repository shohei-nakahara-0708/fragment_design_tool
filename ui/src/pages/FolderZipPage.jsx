import React, { useMemo, useRef, useState } from "react";

import { filesFromDataTransfer, filesFromInput } from "../lib/folderDrop.js";

const API_BASE = import.meta.env.VITE_API_BASE || "";
const INVALID_FILENAME_CHARS = new Set(["<", ">", ":", "\"", "/", "\\", "|", "?", "*"]);
const IGNORED_FOLDER_PARTS = new Set([
  "__macosx",
  ".appledouble",
  ".fseventsd",
  ".spotlight-v100",
  ".temporaryitems",
  ".trashes",
]);
const IGNORED_FILE_NAMES = new Set([
  ".ds_store",
  ".localized",
  "desktop.ini",
  "icon\r",
  "thumbs.db",
]);

function formatSize(bytes) {
  if (!bytes) return "-";
  const kb = bytes / 1024;
  if (kb < 1024) return `${kb.toFixed(1)} KB`;
  return `${(kb / 1024).toFixed(1)} MB`;
}

function errorMessage(payload, fallback) {
  if (!payload) return fallback;
  if (typeof payload === "string") return payload;
  if (typeof payload.detail === "string") return payload.detail;
  return fallback;
}

function normalizeZipFilename(value) {
  return String(value || "")
    .trim()
    .replace(/\.zip$/i, "")
    .split("")
    .map((char) => (char.charCodeAt(0) < 32 || INVALID_FILENAME_CHARS.has(char) ? "_" : char))
    .join("")
    .replace(/_+/g, "_")
    .replace(/^[ ._]+|[ ._]+$/g, "");
}

function cleanPath(value) {
  return String(value || "")
    .replaceAll("\\", "/")
    .replace(/^\/+/, "")
    .replace(/\/+/g, "/");
}

function basename(path) {
  return cleanPath(path).split("/").pop() || "";
}

function isIgnoredPath(path) {
  const parts = cleanPath(path).split("/").filter(Boolean);
  if (!parts.length) return true;
  if (parts.some((part) => IGNORED_FOLDER_PARTS.has(part.toLowerCase()))) return true;

  const name = parts[parts.length - 1] || "";
  return IGNORED_FILE_NAMES.has(name.toLowerCase()) || name.startsWith("._");
}

function pathRoot(path) {
  const parts = cleanPath(path).split("/");
  return parts.length > 1 ? parts[0] : "";
}

const ZIP_MODE_OPTIONS = [
  ["auto", "自動判定"],
  ["include_root", "親フォルダごと"],
  ["contents_only", "中身だけ"],
];

function folderSummaries(items, folderModes) {
  const grouped = new Map();
  items.forEach((item) => {
    const root = pathRoot(item.relativePath);
    if (!root) return;
    if (!grouped.has(root)) grouped.set(root, []);
    grouped.get(root).push(item);
  });

  return Array.from(grouped, ([root, groupItems]) => {
    const groupPaths = groupItems.map((item) => item.relativePath);
    const detection = detectMode(groupPaths);
    const selectedMode = folderModes[root] || "auto";
    const resolvedMode = selectedMode === "auto" ? detection.mode : selectedMode;
    const safeFilename = normalizeZipFilename(root);
    return {
      root,
      safeFilename,
      fileCount: groupItems.length,
      size: groupItems.reduce((sum, item) => sum + (item.file?.size || 0), 0),
      detection,
      selectedMode,
      resolvedMode,
      previewPaths: groupPaths.slice(0, 3).map((path) => arcnamePreview(path, resolvedMode, root, safeFilename)),
    };
  });
}

function detectMode(paths) {
  const names = paths.map((path) => basename(path).toLowerCase());
  if (names.some((name) => /-thumb\.jpe?g$/i.test(name))) {
    return {
      mode: "include_root",
      label: "*-thumb.jpg を検出",
      description: "MAILTOOL形式として親フォルダごと含めます。",
    };
  }
  if (names.some((name) => name === "thumb.png")) {
    return {
      mode: "contents_only",
      label: "thumb.png を検出",
      description: "スライド/PDF形式としてフォルダの中身だけをZIP化します。",
    };
  }
  return {
    mode: "contents_only",
    label: "判定対象なし",
    description: "初期値としてフォルダの中身だけをZIP化します。",
  };
}

function arcnamePreview(path, mode, rootName, fallbackRoot) {
  const clean = cleanPath(path);
  if (!clean) return "";
  const parts = clean.split("/");
  if (mode === "contents_only" && rootName && parts[0] === rootName) {
    return parts.slice(1).join("/") || clean;
  }
  if (mode === "include_root" && !rootName) {
    return `${fallbackRoot}/${clean}`;
  }
  return clean;
}

function downloadUrl(value) {
  if (!value) return "";
  if (/^https?:\/\//i.test(value)) return value;
  return `${API_BASE}${value}`;
}

export default function FolderZipPage() {
  const folderInputRef = useRef(null);
  const [items, setItems] = useState([]);
  const [folderModes, setFolderModes] = useState({});
  const [dragOver, setDragOver] = useState(false);
  const [generating, setGenerating] = useState(false);
  const [error, setError] = useState("");
  const [ignoredCount, setIgnoredCount] = useState(0);
  const [result, setResult] = useState(null);

  const folderList = useMemo(() => folderSummaries(items, folderModes), [items, folderModes]);
  const hasRootlessItems = Boolean(items.length && items.some((item) => !pathRoot(item.relativePath)));
  const totalSize = items.reduce((sum, item) => sum + (item.file?.size || 0), 0);
  const canGenerate = Boolean(items.length && folderList.length && !hasRootlessItems && !generating);

  const applyItems = (nextItems) => {
    const candidates = nextItems
      .map((item) => ({
        file: item.file,
        relativePath: cleanPath(item.relativePath || item.file?.name),
      }))
      .filter((item) => item.file && item.relativePath && !item.relativePath.endsWith("/"));
    const normalized = candidates.filter((item) => !isIgnoredPath(item.relativePath));

    const normalizedByPath = new Map(normalized.map((item) => [item.relativePath.toLowerCase(), item]));
    setItems((current) => {
      const merged = current.filter((item) => !normalizedByPath.has(item.relativePath.toLowerCase()));
      merged.push(...normalized);
      return merged;
    });
    setIgnoredCount((current) => current + candidates.length - normalized.length);
    setResult(null);
    const hasRootless = normalized.some((item) => !pathRoot(item.relativePath));
    setError(
      candidates.length && !normalized.length
        ? "ZIP化できるファイルがありません。"
        : hasRootless
          ? "ZIPファイル名をフォルダ名に固定するため、フォルダごと追加してください。"
          : "",
    );
  };

  const handleDrop = async (event) => {
    event.preventDefault();
    setDragOver(false);
    if (generating) return;

    try {
      const droppedItems = await filesFromDataTransfer(event.dataTransfer);
      if (!droppedItems.length) {
        setError("ファイルを取得できませんでした。");
        return;
      }
      applyItems(droppedItems);
    } catch (err) {
      setError(err?.message || "フォルダを読み込めませんでした。");
    }
  };

  const clearItems = () => {
    setItems([]);
    setFolderModes({});
    setIgnoredCount(0);
    setResult(null);
    setError("");
  };

  const changeFolderMode = (root, value) => {
    setFolderModes((current) => ({ ...current, [root]: value }));
    setResult(null);
  };

  const generate = async () => {
    if (!canGenerate) return;
    setGenerating(true);
    setError("");
    setResult(null);

    try {
      const formData = new FormData();
      items.forEach((item) => formData.append("files", item.file, item.file.name));
      formData.append("relativePaths", JSON.stringify(items.map((item) => item.relativePath)));
      formData.append("mode", "auto");
      formData.append("folderModes", JSON.stringify(folderModes));

      const response = await fetch(`${API_BASE}/folder-zip-tool/generate-batch`, {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        const payload = await response.json().catch(() => null);
        throw new Error(errorMessage(payload, "ZIP生成に失敗しました。"));
      }

      const payload = await response.json();
      setResult({
        ...payload,
        results: Array.isArray(payload?.results) ? payload.results : [],
      });
    } catch (err) {
      setError(err?.message || "ZIP生成に失敗しました。");
    } finally {
      setGenerating(false);
    }
  };

  return (
    <div className="lecture-tool-page">
      <div className="lecture-tool-page__inner">
        <header className="lecture-tool-header">
          <div>
            <h1 className="lecture-tool-header__title">フォルダZIP生成</h1>
            <div className="lecture-tool-header__sub">
              thumb名の形式に合わせて、レガシーかモダンかを自動判定してZIP化します。
            </div>
          </div>
          <div className="lecture-tool-actions pdf-slide-actions">
            <button
              className="lecture-tool-button lecture-tool-button--primary"
              type="button"
              onClick={generate}
              disabled={!canGenerate}
            >
              {generating ? "生成中" : "ZIPを生成"}
            </button>
          </div>
        </header>

        {error ? <div className="lecture-tool-alert">{error}</div> : null}

        <div className="pdf-slide-grid">
          <section className="lecture-tool-panel lecture-tool-panel--wide">
            <div className="lecture-tool-panel__inner">
              <div className="lecture-tool-panel__head">
                <div>
                  <h2 className="lecture-tool-panel__title">フォルダ追加</h2>
                  <div className="lecture-tool-panel__sub">フォルダ選択またはドラッグ&ドロップに対応しています。</div>
                </div>
                {items.length ? (
                  <button
                    className="lecture-tool-button lecture-tool-button--small"
                    type="button"
                    onClick={clearItems}
                    disabled={generating}
                  >
                    削除
                  </button>
                ) : null}
              </div>

              <div
                className={`lecture-tool-drop${dragOver ? " lecture-tool-drop--active" : ""}${generating ? " lecture-tool-drop--disabled" : ""}`}
                onClick={() => {
                  if (!generating) folderInputRef.current?.click();
                }}
                onDragOver={(event) => {
                  event.preventDefault();
                  if (!generating) setDragOver(true);
                }}
                onDragLeave={() => setDragOver(false)}
                onDrop={handleDrop}
                onKeyDown={(event) => {
                  if ((event.key === "Enter" || event.key === " ") && !generating) {
                    event.preventDefault();
                    folderInputRef.current?.click();
                  }
                }}
                role="button"
                tabIndex={0}
              >
                <div className="lecture-tool-drop__title">
                  {items.length
                    ? `${folderList.length || 0}フォルダ / ${items.length}件のファイルを追加済み`
                    : "フォルダを選択 または ドラッグ&ドロップ"}
                </div>
                <div className="lecture-tool-drop__sub">
                  {items.length
                    ? `${formatSize(totalSize)}${ignoredCount ? ` / 除外 ${ignoredCount}件` : ""}`
                    : "空フォルダはブラウザ仕様上取り込まれません。"}
                </div>
                <input
                  ref={folderInputRef}
                  className="lecture-tool-hidden-input"
                  type="file"
                  webkitdirectory=""
                  multiple
                  onChange={(event) => {
                    applyItems(filesFromInput(event.target.files));
                    event.target.value = "";
                  }}
                  disabled={generating}
                />
              </div>
            </div>
          </section>

          <section className="lecture-tool-panel lecture-tool-panel--wide">
            <div className="lecture-tool-panel__inner">
              <div className="lecture-tool-panel__head">
                <div>
                  <h2 className="lecture-tool-panel__title">ZIP設定</h2>
                  <div className="lecture-tool-panel__sub">判定結果を見ながらZIP内の階層を指定します。</div>
                </div>
              </div>

              <div className="folder-zip-settings">
                <div className="folder-zip-fixed-name">
                  <div className="lecture-tool-select-label lecture-tool-select-label-sub">生成ZIP</div>
                  {folderList.length ? (
                    <div className="folder-zip-folder-list">
                      {folderList.map((folder) => (
                        <div className="folder-zip-folder-row" key={folder.root}>
                          <div className="pdf-slide-name-preview">{folder.safeFilename}.zip</div>
                          <div className="lecture-tool-panel__sub">
                            {folder.fileCount}ファイル / {formatSize(folder.size)}
                          </div>
                        </div>
                      ))}
                    </div>
                  ) : (
                    <div className="pdf-slide-name-preview">
                      {items.length ? "フォルダ名を取得できません" : "フォルダごとに自動設定"}
                    </div>
                  )}
                </div>

                <div className="folder-zip-mode">
                  <div className="lecture-tool-select-label">フォルダごとのZIP方法</div>
                  {folderList.length ? (
                    <div className="folder-zip-mode-list">
                      {folderList.map((folder) => (
                        <div className="folder-zip-mode-row" key={folder.root}>
                          <div className="folder-zip-mode-row__head">
                            <div>
                              <div className="folder-zip-mode-row__title">{folder.safeFilename}.zip</div>
                              <div className="folder-zip-mode-row__meta">{folder.detection.label}</div>
                            </div>
                            <span className="folder-zip-detection__mode">
                              実行時: {folder.resolvedMode === "include_root" ? "親フォルダごと" : "中身だけ"}
                            </span>
                          </div>
                          <div className="folder-zip-mode__buttons">
                            {ZIP_MODE_OPTIONS.map(([value, label]) => (
                              <button
                                className={`folder-zip-mode__button${folder.selectedMode === value ? " folder-zip-mode__button--active" : ""}`}
                                type="button"
                                key={value}
                                onClick={() => changeFolderMode(folder.root, value)}
                                disabled={generating}
                              >
                                {label}
                              </button>
                            ))}
                          </div>
                        </div>
                      ))}
                    </div>
                  ) : (
                    <div className="folder-zip-detection">
                      <div className="folder-zip-detection__title">フォルダ未追加</div>
                      <div className="folder-zip-detection__meta">追加後にフォルダごとのZIP方法を設定できます。</div>
                    </div>
                  )}
                </div>
              </div>
            </div>
          </section>

          <section className="lecture-tool-panel lecture-tool-panel--wide pdf-slide-result-panel">
            <div className="lecture-tool-panel__inner">
              <div className="lecture-tool-panel__head">
                <div>
                  <h2 className="lecture-tool-panel__title">ZIP内パス</h2>
                  <div className="lecture-tool-panel__sub">先頭数件の出力パスを確認できます。</div>
                </div>
                <button
                  className="lecture-tool-button lecture-tool-button--primary"
                  type="button"
                  onClick={generate}
                  disabled={!canGenerate}
                >
                  {generating ? "生成中" : "ZIPを生成"}
                </button>
              </div>

              {folderList.length ? (
                <div className="folder-zip-preview-groups">
                  {folderList.map((folder) => (
                    <div className="folder-zip-preview-group" key={folder.root}>
                      <div className="folder-zip-preview-group__head">
                        <div className="folder-zip-preview-group__title">{folder.safeFilename}.zip</div>
                        <div className="folder-zip-preview-group__meta">
                          {folder.resolvedMode === "include_root" ? "親フォルダごと" : "中身だけ"}
                        </div>
                      </div>
                      <div className="folder-zip-preview-list">
                        {folder.previewPaths.map((path) => (
                          <div className="folder-zip-preview-row" key={`${folder.root}:${path}`}>{path}</div>
                        ))}
                        {folder.fileCount > folder.previewPaths.length ? (
                          <div className="lecture-tool-panel__sub">
                            ほか {folder.fileCount - folder.previewPaths.length} 件
                          </div>
                        ) : null}
                      </div>
                    </div>
                  ))}
                </div>
              ) : (
                <div className="lecture-tool-empty">フォルダを追加してください。</div>
              )}

              {result?.results?.length ? (
                <div className="folder-zip-result-list">
                  {result.batch ? (
                    <div className="pdf-slide-result pdf-slide-result--batch folder-zip-result">
                      <div>
                        <div className="pdf-slide-result__title">{result.batch.filename}</div>
                        <div className="pdf-slide-result__meta">
                          一括ダウンロード / {result.folderCount} ZIP / {formatSize(result.batch.size)}
                        </div>
                      </div>
                      <a
                        className="lecture-tool-button lecture-tool-button--primary"
                        href={downloadUrl(result.batch.downloadUrl)}
                        download={result.batch.filename}
                      >
                        一括ダウンロード
                      </a>
                    </div>
                  ) : null}
                  {result.results.map((item) => (
                    <div className="pdf-slide-result pdf-slide-result--batch folder-zip-result" key={item.filename}>
                      <div>
                        <div className="pdf-slide-result__title">{item.filename}</div>
                        <div className="pdf-slide-result__meta">
                          {item.fileCount}ファイル / {item.mode === "include_root" ? "親フォルダごと" : "中身だけ"} / {formatSize(item.size)}
                        </div>
                      </div>
                      <a
                        className="lecture-tool-button lecture-tool-button--primary"
                        href={downloadUrl(item.downloadUrl)}
                        download={item.filename}
                      >
                        ZIPをダウンロード
                      </a>
                    </div>
                  ))}
                </div>
              ) : null}
            </div>
          </section>
        </div>
      </div>
    </div>
  );
}
