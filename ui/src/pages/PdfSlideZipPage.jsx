import React, { useCallback, useEffect, useRef, useState } from "react";

import { createPdfFirstPageThumbnail, createPdfPageThumbnail } from "../lib/pdfThumbnail.js";

const API_BASE = import.meta.env.VITE_API_BASE || "";

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

function normalizePresentationId(value) {
  return String(value || "")
    .trim()
    .replace(/\.zip$/i, "")
    .replace(/[^A-Za-z0-9._-]+/g, "_")
    .replace(/^[._-]+|[._-]+$/g, "");
}

function pageZipBaseName(presentationId, page, pageCount) {
  const safeId = normalizePresentationId(presentationId) || "PresentationId";
  const digits = Math.max(3, String(pageCount || 0).length);
  return `${safeId}_${String(page || 1).padStart(digits, "0")}`;
}

function downloadUrl(value) {
  if (!value) return "";
  if (/^https?:\/\//i.test(value)) return value;
  return `${API_BASE}${value}`;
}

function fileKey(file) {
  return `${file.name}:${file.size}:${file.lastModified}`;
}

function createDocument(file) {
  return {
    id: globalThis.crypto?.randomUUID?.() || `${Date.now()}-${Math.random()}`,
    file,
    presentationId: "",
    analyzing: true,
    titleError: "",
    titles: [],
    thumbnailUrl: "",
    thumbnailLoading: true,
    thumbnailError: "",
    copyThumbnailUrl: "",
    copyThumbnailPage: 0,
    copyThumbnailLoading: false,
    copyThumbnailError: "",
  };
}

async function copyText(text) {
  if (navigator.clipboard?.writeText) {
    await navigator.clipboard.writeText(text);
    return;
  }

  const textarea = document.createElement("textarea");
  textarea.value = text;
  textarea.style.position = "fixed";
  textarea.style.opacity = "0";
  document.body.appendChild(textarea);
  textarea.select();
  document.execCommand("copy");
  textarea.remove();
}

export default function PdfSlideZipPage() {
  const fileInputRef = useRef(null);
  const copiedTimerRef = useRef(null);
  const copyThumbnailRequestRef = useRef({});
  const [documents, setDocuments] = useState([]);
  const [dragOver, setDragOver] = useState(false);
  const [generating, setGenerating] = useState(false);
  const [error, setError] = useState("");
  const [copiedKey, setCopiedKey] = useState("");
  const [copyModeIndexes, setCopyModeIndexes] = useState({});
  const [openTitlePanels, setOpenTitlePanels] = useState({});
  const [result, setResult] = useState(null);

  useEffect(() => {
    return () => {
      if (copiedTimerRef.current) window.clearTimeout(copiedTimerRef.current);
    };
  }, []);

  const updateDocument = useCallback((documentId, changes) => {
    setDocuments((current) => current.map((item) => (item.id === documentId ? { ...item, ...changes } : item)));
  }, []);

  const ensureCopyModeThumbnail = useCallback((documentItem, pageNumber) => {
    if (!documentItem || !pageNumber) return;
    if (documentItem.copyThumbnailPage === pageNumber && documentItem.copyThumbnailUrl) return;
    if (
      documentItem.copyThumbnailPage === pageNumber &&
      (documentItem.copyThumbnailLoading || documentItem.copyThumbnailError)
    ) {
      return;
    }

    const requestKey = `${pageNumber}:${Date.now()}:${Math.random()}`;
    copyThumbnailRequestRef.current[documentItem.id] = requestKey;
    updateDocument(documentItem.id, {
      copyThumbnailUrl: "",
      copyThumbnailPage: pageNumber,
      copyThumbnailLoading: true,
      copyThumbnailError: "",
    });

    createPdfPageThumbnail(documentItem.file, pageNumber, 680)
      .then((thumbnailUrl) => {
        if (copyThumbnailRequestRef.current[documentItem.id] !== requestKey) return;
        updateDocument(documentItem.id, {
          copyThumbnailUrl: thumbnailUrl,
          copyThumbnailPage: pageNumber,
          copyThumbnailLoading: false,
          copyThumbnailError: "",
        });
      })
      .catch(() => {
        if (copyThumbnailRequestRef.current[documentItem.id] !== requestKey) return;
        updateDocument(documentItem.id, {
          copyThumbnailUrl: "",
          copyThumbnailPage: pageNumber,
          copyThumbnailLoading: false,
          copyThumbnailError: "生成できませんでした",
        });
      });
  }, [updateDocument]);

  useEffect(() => {
    documents.forEach((documentItem) => {
      if (!openTitlePanels[documentItem.id] || !documentItem.titles.length) return;
      const index = Math.min(
        Math.max(copyModeIndexes[documentItem.id] ?? 0, 0),
        documentItem.titles.length - 1,
      );
      const pageNumber = documentItem.titles[index]?.page;
      ensureCopyModeThumbnail(documentItem, pageNumber);
    });
  }, [copyModeIndexes, documents, ensureCopyModeThumbnail, openTitlePanels]);

  const analyzePdf = async (documentId, file) => {
    try {
      const formData = new FormData();
      formData.append("pdf", file, file.name);
      const response = await fetch(`${API_BASE}/pdf-slide-tool/analyze`, {
        method: "POST",
        body: formData,
      });
      const payload = await response.json().catch(() => null);

      if (!response.ok) {
        throw new Error(errorMessage(payload, "タイトル抽出に失敗しました。"));
      }
      updateDocument(documentId, {
        analyzing: false,
        titleError: "",
        titles: Array.isArray(payload?.titles) ? payload.titles : [],
      });
    } catch (err) {
      updateDocument(documentId, {
        analyzing: false,
        titleError: err?.message || "タイトル抽出に失敗しました。",
        titles: [],
      });
    }
  };

  const generateThumbnail = async (documentId, file) => {
    try {
      const thumbnailUrl = await createPdfFirstPageThumbnail(file);
      updateDocument(documentId, {
        thumbnailUrl,
        thumbnailLoading: false,
        thumbnailError: "",
      });
    } catch {
      updateDocument(documentId, {
        thumbnailUrl: "",
        thumbnailLoading: false,
        thumbnailError: "サムネイルを生成できませんでした。",
      });
    }
  };

  const addFiles = (fileList) => {
    const selectedFiles = Array.from(fileList || []);
    if (!selectedFiles.length) return;

    const invalidFiles = selectedFiles.filter(
      (file) => !file.name.toLowerCase().endsWith(".pdf") && file.type !== "application/pdf",
    );
    const existingKeys = new Set(documents.map((item) => fileKey(item.file)));
    const additions = selectedFiles
      .filter((file) => file.name.toLowerCase().endsWith(".pdf") || file.type === "application/pdf")
      .filter((file) => !existingKeys.has(fileKey(file)))
      .map(createDocument);

    if (invalidFiles.length) {
      setError("PDF以外のファイルは追加されませんでした。");
    } else if (!additions.length) {
      setError("選択したPDFはすでに追加されています。");
    } else {
      setError("");
    }

    if (!additions.length) return;
    setResult(null);
    setDocuments((current) => [...current, ...additions]);
    additions.forEach((item) => {
      analyzePdf(item.id, item.file);
      generateThumbnail(item.id, item.file);
    });
  };

  const removeDocument = (documentId) => {
    delete copyThumbnailRequestRef.current[documentId];
    setDocuments((current) => current.filter((item) => item.id !== documentId));
    setCopyModeIndexes((current) => {
      const next = { ...current };
      delete next[documentId];
      return next;
    });
    setOpenTitlePanels((current) => {
      const next = { ...current };
      delete next[documentId];
      return next;
    });
    setResult(null);
    setError("");
  };

  const changePresentationId = (documentId, value) => {
    updateDocument(documentId, { presentationId: value });
    setResult(null);
  };

  const copyTitle = async (documentId, page, title) => {
    if (!title) return false;
    try {
      await copyText(title);
      const key = `${documentId}:${page}`;
      setCopiedKey(key);
      setError("");
      if (copiedTimerRef.current) window.clearTimeout(copiedTimerRef.current);
      copiedTimerRef.current = window.setTimeout(() => setCopiedKey(""), 1600);
      return true;
    } catch {
      setError("コピーできませんでした。");
      return false;
    }
  };

  const copyModeIndex = (documentId, titles) => {
    if (!titles.length) return 0;
    const index = copyModeIndexes[documentId] ?? 0;
    return Math.min(Math.max(index, 0), titles.length - 1);
  };

  const setCopyModeIndex = (documentId, titles, nextIndex) => {
    if (!titles.length) return;
    const index = Math.min(Math.max(nextIndex, 0), titles.length - 1);
    setCopyModeIndexes((current) => ({ ...current, [documentId]: index }));
  };

  const copyTitleAndAdvance = async (documentId, titles) => {
    const index = copyModeIndex(documentId, titles);
    const item = titles[index];
    if (!item?.title) return;

    const copied = await copyTitle(documentId, item.page, item.title);
    if (copied && index < titles.length - 1) {
      setCopyModeIndex(documentId, titles, index + 1);
    }
  };

  const jumpToPage = (documentId, titles, value) => {
    const page = Number.parseInt(value, 10);
    if (!Number.isFinite(page)) return;
    setCopyModeIndex(documentId, titles, page - 1);
  };

  const normalizedIds = documents.map((item) => normalizePresentationId(item.presentationId));
  const normalizedIdKeys = normalizedIds.map((id) => id.toLocaleLowerCase());
  const duplicateIds = [
    ...new Set(
      normalizedIds.filter(
        (id, index) => id && normalizedIdKeys.indexOf(normalizedIdKeys[index]) !== index,
      ),
    ),
  ];
  const hasInvalidId = documents.some((item) => !normalizePresentationId(item.presentationId));
  const validationMessage = duplicateIds.length
    ? `Presentation IDが重複しています: ${duplicateIds.join(", ")}`
    : hasInvalidId && documents.length
      ? "すべてのPDFにPresentation IDを入力してください。"
      : "";

  const generate = async () => {
    if (!documents.length || hasInvalidId || duplicateIds.length || generating) return;
    setGenerating(true);
    setError("");
    setResult(null);

    try {
      const formData = new FormData();
      documents.forEach((item) => formData.append("pdfs", item.file, item.file.name));
      formData.append("presentationIds", JSON.stringify(documents.map((item) => item.presentationId.trim())));
      const response = await fetch(`${API_BASE}/pdf-slide-tool/generate-batch`, {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        const payload = await response.json().catch(() => null);
        throw new Error(errorMessage(payload, "一括ZIP生成に失敗しました。"));
      }

      const payload = await response.json();
      setResult({
        ...payload,
        results: Array.isArray(payload?.results)
          ? payload.results.map((item, index) => ({
            ...item,
            sourceName: documents[index]?.file.name || item.presentationId,
          }))
          : [],
      });
    } catch (err) {
      setError(err?.message || "一括ZIP生成に失敗しました。");
    } finally {
      setGenerating(false);
    }
  };

  const canGenerate = Boolean(documents.length && !hasInvalidId && !duplicateIds.length && !generating);
  const totalAnalyzing = documents.filter((item) => item.analyzing).length;

  return (
    <div className="lecture-tool-page">
      <div className="lecture-tool-page__inner">
        <header className="lecture-tool-header">
          <div>
            <h1 className="lecture-tool-header__title">Shared付きスライド</h1>
            <div className="lecture-tool-header__sub">複数PDFからページごとのスライドZIPを生成します。</div>
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
        {validationMessage ? <div className="lecture-tool-alert">{validationMessage}</div> : null}

        <div className="pdf-slide-grid">
          <section className="lecture-tool-panel lecture-tool-panel--wide">
            <div className="lecture-tool-panel__inner">
              <div className="lecture-tool-panel__head">
                <div>
                  <h2 className="lecture-tool-panel__title">PDF追加</h2>
                  <div className="lecture-tool-panel__sub">複数選択または追加のドラッグ&ドロップに対応しています。</div>
                </div>
                {documents.length ? (
                  <button
                    className="lecture-tool-button lecture-tool-button--small"
                    type="button"
                    onClick={() => {
                      copyThumbnailRequestRef.current = {};
                      setDocuments([]);
                      setCopyModeIndexes({});
                      setOpenTitlePanels({});
                      setResult(null);
                      setError("");
                    }}
                    disabled={generating}
                  >
                    すべて削除
                  </button>
                ) : null}
              </div>
              <div
                className={`lecture-tool-drop${dragOver ? " lecture-tool-drop--active" : ""}${generating ? " lecture-tool-drop--disabled" : ""}`}
                onClick={() => {
                  if (!generating) fileInputRef.current?.click();
                }}
                onDragOver={(event) => {
                  event.preventDefault();
                  if (!generating) setDragOver(true);
                }}
                onDragLeave={() => setDragOver(false)}
                onDrop={(event) => {
                  event.preventDefault();
                  setDragOver(false);
                  if (!generating) addFiles(event.dataTransfer.files);
                }}
                role="button"
                tabIndex={0}
              >
                <div className="lecture-tool-drop__title">
                  {documents.length ? `${documents.length}件のPDFを追加済み` : "PDFを複数選択 または ドラッグ&ドロップ"}
                </div>
                <div className="lecture-tool-drop__sub">
                  {totalAnalyzing ? `${totalAnalyzing}件のタイトルを解析中` : "選択後もPDFを追加できます。"}
                </div>
                <input
                  ref={fileInputRef}
                  className="lecture-tool-hidden-input"
                  type="file"
                  accept="application/pdf,.pdf"
                  multiple
                  onChange={(event) => {
                    addFiles(event.target.files);
                    event.target.value = "";
                  }}
                  disabled={generating}
                />
              </div>
            </div>
          </section>

          {documents.length ? (
            <section className="lecture-tool-panel lecture-tool-panel--wide pdf-slide-settings-panel">
              <div className="lecture-tool-panel__inner">
                <div className="lecture-tool-panel__head">
                  <div>
                    <h2 className="lecture-tool-panel__title">PDFごとの設定</h2>
                    <div className="lecture-tool-panel__sub">
                      追加したPDFごとにPresentation IDとページタイトルを管理します。
                    </div>
                  </div>
                  <span className="lecture-tool-progress__badge">{documents.length}件</span>
                </div>

                <div className="pdf-slide-document-list">
                  {documents.map((documentItem, documentIndex) => {
                    const modeIndex = copyModeIndex(documentItem.id, documentItem.titles);
                    const modeItem = documentItem.titles[modeIndex];
                    const modeZipName = pageZipBaseName(
                      documentItem.presentationId,
                      modeItem?.page || 1,
                      documentItem.titles.length,
                    );

                    return (
                      <article className="pdf-slide-document-field" key={documentItem.id}>
                        <div className="pdf-slide-document-field__head">
                          <div className="pdf-slide-document__heading">
                            <span className="pdf-slide-document__number">
                              PDF {String(documentIndex + 1).padStart(2, "0")}
                            </span>
                            <div className="pdf-slide-document__thumbnail">
                              {documentItem.thumbnailUrl ? (
                                <img src={documentItem.thumbnailUrl} alt={`${documentItem.file.name} 1ページ目`} />
                              ) : (
                                <div className="pdf-slide-document__thumbnail-placeholder">
                                  {documentItem.thumbnailLoading ? "生成中" : "PDF"}
                                </div>
                              )}
                            </div>
                            <div>
                              <h3 className="pdf-slide-document__filename">{documentItem.file.name}</h3>
                              <div className="lecture-tool-panel__sub">{formatSize(documentItem.file.size)}</div>
                            </div>
                          </div>
                          <button
                            className="lecture-tool-button lecture-tool-button--small"
                            type="button"
                            onClick={() => removeDocument(documentItem.id)}
                            disabled={generating}
                          >
                            削除
                          </button>
                        </div>

                        <div className="pdf-slide-document__config">
                          <label className="lecture-tool-select-label">
                            Presentation ID
                            <input
                              className="lecture-tool-search"
                              value={documentItem.presentationId}
                              onChange={(event) => changePresentationId(documentItem.id, event.target.value)}
                              placeholder="minjuvi-fl_slide_008"
                              disabled={generating}
                            />
                          </label>
                          <label className="lecture-tool-select-label">
                            先頭ページのZIP名
                            <div className="pdf-slide-name-preview">
                              {normalizePresentationId(documentItem.presentationId)
                                ? `${normalizePresentationId(documentItem.presentationId)}_001.zip`
                                : "PresentationId_001.zip"}
                            </div>
                          </label>
                        </div>

                        <details
                          className="pdf-slide-title-accordion"
                          onToggle={(event) => {
                            const isOpen = event.currentTarget.open;
                            setOpenTitlePanels((current) => ({
                              ...current,
                              [documentItem.id]: isOpen,
                            }));
                            if (isOpen) ensureCopyModeThumbnail(documentItem, modeItem?.page || 1);
                          }}
                        >
                          <summary className="pdf-slide-title-accordion__summary">
                            <div className="pdf-slide-title-accordion__label">
                              <span className="pdf-slide-title-accordion__mark">T</span>
                              <div>
                                <h3 className="pdf-slide-document__titles-title">ページタイトル</h3>
                                <div className="lecture-tool-panel__sub">
                                  抽出したタイトルの確認・コピー
                                </div>
                              </div>
                            </div>
                            <div className="pdf-slide-title-accordion__meta">
                              <span className="lecture-tool-progress__badge">
                                {documentItem.analyzing
                                  ? "解析中"
                                  : documentItem.titleError
                                    ? "取得失敗"
                                    : `${documentItem.titles.length}ページ`}
                              </span>
                              <span className="pdf-slide-title-accordion__action">
                                <span className="pdf-slide-title-accordion__open">+ 表示する</span>
                                <span className="pdf-slide-title-accordion__close">- 閉じる</span>
                              </span>
                            </div>
                          </summary>
                          <div className="pdf-slide-title-accordion__body">
                            {documentItem.analyzing ? (
                              <div className="lecture-tool-progress">
                                <div className="lecture-tool-progress__head">
                                  <div>
                                    <div className="lecture-tool-progress__title">タイトル抽出中</div>
                                    <div className="lecture-tool-progress__current">
                                      PDF内のテキストをページごとに解析しています。
                                    </div>
                                  </div>
                                  <span className="lecture-tool-progress__badge">解析中</span>
                                </div>
                              </div>
                            ) : documentItem.titleError ? (
                              <div className="lecture-tool-alert">{documentItem.titleError}</div>
                            ) : documentItem.titles.length ? (
                              <>
                                <div className="pdf-slide-copy-mode">
                                  <div className="pdf-slide-copy-mode__head">
                                    <div>
                                      <div className="pdf-slide-copy-mode__title">ページタイトルコピー</div>
                                      <div className="pdf-slide-copy-mode__sub">タイトルをコピーすると次のページへ進みます。</div>
                                    </div>
                                    <span className="pdf-slide-copy-mode__counter">
                                      {modeIndex + 1} / {documentItem.titles.length}
                                    </span>
                                  </div>

                                  <div className="pdf-slide-copy-mode__body">
                                    <div className="pdf-slide-copy-mode__thumbnail">
                                      {documentItem.copyThumbnailUrl ? (
                                        <img
                                          src={documentItem.copyThumbnailUrl}
                                          alt={`${documentItem.file.name} ${modeItem?.page || 1}ページ目`}
                                        />
                                      ) : (
                                        <div className="pdf-slide-copy-mode__thumbnail-placeholder">
                                          {documentItem.copyThumbnailLoading
                                            ? "生成中"
                                            : documentItem.copyThumbnailError || "PDF"}
                                        </div>
                                      )}
                                    </div>
                                    <div className="pdf-slide-copy-mode__page">
                                      {String(modeItem?.page || 1).padStart(3, "0")}
                                    </div>
                                    <textarea
                                      className="pdf-slide-copy-mode__text"
                                      value={modeItem?.title || ""}
                                      readOnly
                                      rows={3}
                                    />
                                  </div>

                                  <div className="pdf-slide-copy-mode__zip">
                                    <div className="pdf-slide-copy-mode__zip-label">ページZIP名（拡張子なし）</div>
                                    <div className="pdf-slide-copy-mode__zip-row">
                                      <input
                                        className="pdf-slide-copy-mode__zip-input"
                                        value={modeZipName}
                                        readOnly
                                      />
                                      <button
                                        className="lecture-tool-button lecture-tool-button--small"
                                        type="button"
                                        onClick={() => copyTitle(documentItem.id, `zip-${modeItem?.page}`, modeZipName)}
                                        disabled={!modeZipName}
                                      >
                                        {copiedKey === `${documentItem.id}:zip-${modeItem?.page}` ? "コピー済み" : "ZIP名コピー"}
                                      </button>
                                    </div>
                                  </div>

                                  <div className="pdf-slide-copy-mode__actions">
                                    <label className="pdf-slide-copy-mode__jump">
                                      ページ指定
                                      <input
                                        type="number"
                                        min="1"
                                        max={documentItem.titles.length}
                                        value={modeIndex + 1}
                                        onFocus={(event) => event.target.select()}
                                        onChange={(event) => jumpToPage(documentItem.id, documentItem.titles, event.target.value)}
                                      />
                                    </label>
                                    <button
                                      className="lecture-tool-button lecture-tool-button--small"
                                      type="button"
                                      onClick={() => setCopyModeIndex(documentItem.id, documentItem.titles, modeIndex - 1)}
                                      disabled={modeIndex <= 0}
                                    >
                                      前へ
                                    </button>
                                    <button
                                      className="lecture-tool-button lecture-tool-button--primary"
                                      type="button"
                                      onClick={() => copyTitleAndAdvance(documentItem.id, documentItem.titles)}
                                      disabled={!modeItem?.title}
                                    >
                                      {copiedKey === `${documentItem.id}:${modeItem?.page}` ? "コピー済み" : "コピーして次へ"}
                                    </button>
                                    <button
                                      className="lecture-tool-button lecture-tool-button--small"
                                      type="button"
                                      onClick={() => setCopyModeIndex(documentItem.id, documentItem.titles, modeIndex + 1)}
                                      disabled={modeIndex >= documentItem.titles.length - 1}
                                    >
                                      次へ
                                    </button>
                                  </div>
                                </div>

                                {/* <div className="pdf-slide-title-list">
                                {documentItem.titles.map((item, titleIndex) => (
                                  <div
                                    className={`pdf-slide-title-row${titleIndex === modeIndex ? " pdf-slide-title-row--active" : ""}`}
                                    key={item.page}
                                  >
                                    <div className="pdf-slide-title-row__page">{String(item.page).padStart(3, "0")}</div>
                                    <textarea
                                      className="pdf-slide-title-row__text"
                                      value={item.title || ""}
                                      placeholder="タイトルを取得できませんでした"
                                      readOnly
                                      rows={2}
                                    />
                                    <button
                                      className="lecture-tool-button lecture-tool-button--small"
                                      type="button"
                                      onClick={() => {
                                        setCopyModeIndex(documentItem.id, documentItem.titles, titleIndex);
                                        copyTitle(documentItem.id, item.page, item.title);
                                      }}
                                      disabled={!item.title}
                                    >
                                      {copiedKey === `${documentItem.id}:${item.page}` ? "コピー済み" : "コピー"}
                                    </button>
                                  </div>
                                ))}
                              </div> */}
                              </>
                            ) : (
                              <div className="lecture-tool-empty">タイトルを取得できませんでした。</div>
                            )}
                          </div>
                        </details>
                      </article>
                    );
                  })}
                </div>
              </div>
            </section>
          ) : null}

          <section className="lecture-tool-panel lecture-tool-panel--wide pdf-slide-result-panel">
            <div className="lecture-tool-panel__inner">
              <div className="lecture-tool-panel__head">
                <div>
                  <h2 className="lecture-tool-panel__title">生成結果</h2>
                  <div className="lecture-tool-panel__sub">PDF別または全PDF一括でダウンロードできます。</div>
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

              {generating ? (
                <div className="lecture-tool-progress">
                  <div className="lecture-tool-progress__head">
                    <div>
                      <div className="lecture-tool-progress__title">一括生成中</div>
                      <div className="lecture-tool-progress__current">PDFページを画像化し、個別ZIPを作成しています。</div>
                    </div>
                    <span className="lecture-tool-progress__badge">実行中</span>
                  </div>
                </div>
              ) : result ? (
                <div className="pdf-slide-result-groups">
                  <section className="pdf-slide-result-group">
                    <h3 className="pdf-slide-result-group__title">PDFごと</h3>
                    <div className="pdf-slide-result-list">
                      {result.results.map((item) => (
                        <div className="pdf-slide-result" key={item.presentationId}>
                          <div>
                            <div className="pdf-slide-result__title">{item.presentationId}</div>
                            <div className="pdf-slide-result__source">{item.sourceName}</div>
                            <div className="pdf-slide-result__meta">
                              {item.pageCount}ページ / {item.pageCount}個の個別ZIP / {formatSize(item.size)}
                            </div>
                          </div>
                          <a
                            className="lecture-tool-button lecture-tool-button--primary"
                            href={downloadUrl(item.downloadUrl)}
                            download={item.filename}
                          >
                            PDF別ZIP
                          </a>
                        </div>
                      ))}
                    </div>
                  </section>

                  <section className="pdf-slide-result-group">
                    <h3 className="pdf-slide-result-group__title">全PDF一括</h3>
                    <div className="pdf-slide-result pdf-slide-result--batch">
                      <div>
                        <div className="pdf-slide-result__title">{result.batch.filename}</div>
                        <div className="pdf-slide-result__meta">
                          {result.pdfCount}PDF / {result.pageCount}ページ / Presentation IDごとのフォルダ分け / {formatSize(result.batch.size)}
                        </div>
                      </div>
                      <a
                        className="lecture-tool-button lecture-tool-button--primary"
                        href={downloadUrl(result.batch.downloadUrl)}
                        download={result.batch.filename}
                      >
                        一括ZIP
                      </a>
                    </div>
                  </section>
                </div>
              ) : (
                <div className="lecture-tool-empty">PDFを追加し、PDFごとのPresentation IDを入力してください。</div>
              )}
            </div>
          </section>
        </div>
      </div>
    </div>
  );
}
