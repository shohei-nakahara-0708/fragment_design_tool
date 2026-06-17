import React, { useRef, useState } from "react";

import { createPdfFirstPageThumbnail } from "../lib/pdfThumbnail.js";

const API_BASE = import.meta.env.VITE_API_BASE || "";
const INVALID_FILENAME_CHARS = new Set(["<", ">", ":", "\"", "/", "\\", "|", "?", "*"]);

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
  const withoutExtension = String(value || "")
    .trim()
    .replace(/\.zip$/i, "");
  return Array.from(withoutExtension)
    .map((char) => (char.charCodeAt(0) < 32 || INVALID_FILENAME_CHARS.has(char) ? "_" : char))
    .join("")
    .replace(/_+/g, "_")
    .replace(/^[ ._]+|[ ._]+$/g, "");
}

function defaultZipFilename(file) {
  return normalizeZipFilename(String(file?.name || "").replace(/\.pdf$/i, "")) || "pdf_document";
}

function fileKey(file) {
  return `${file.name}:${file.size}:${file.lastModified}`;
}

function downloadUrl(value) {
  if (!value) return "";
  if (/^https?:\/\//i.test(value)) return value;
  return `${API_BASE}${value}`;
}

function uniqueZipFilename(baseName, usedNames) {
  const base = normalizeZipFilename(baseName) || "pdf_document";
  let candidate = base;
  let suffix = 2;
  while (usedNames.has(candidate.toLocaleLowerCase())) {
    candidate = `${base}_${String(suffix).padStart(2, "0")}`;
    suffix += 1;
  }
  return candidate;
}

function createDocument(file, zipFilename) {
  return {
    id: globalThis.crypto?.randomUUID?.() || `${Date.now()}-${Math.random()}`,
    file,
    zipFilename,
    thumbnailUrl: "",
    thumbnailLoading: true,
    thumbnailError: "",
  };
}

export default function PdfDocumentZipPage() {
  const fileInputRef = useRef(null);
  const [documents, setDocuments] = useState([]);
  const [dragOver, setDragOver] = useState(false);
  const [generating, setGenerating] = useState(false);
  const [error, setError] = useState("");
  const [result, setResult] = useState(null);

  const normalizedNames = documents.map((item) => normalizeZipFilename(item.zipFilename));
  const normalizedNameKeys = normalizedNames.map((name) => name.toLocaleLowerCase());
  const duplicateNames = [
    ...new Set(
      normalizedNames.filter(
        (name, index) => name && normalizedNameKeys.indexOf(normalizedNameKeys[index]) !== index,
      ),
    ),
  ];
  const hasInvalidName = documents.some((item) => !normalizeZipFilename(item.zipFilename));
  const validationMessage = duplicateNames.length
    ? `ZIPファイル名が重複しています: ${duplicateNames.join(", ")}`
    : hasInvalidName && documents.length
      ? "すべてのPDFにZIPファイル名を入力してください。"
      : "";
  const canGenerate = Boolean(documents.length && !hasInvalidName && !duplicateNames.length && !generating);

  const addFiles = (fileList) => {
    const selectedFiles = Array.from(fileList || []);
    if (!selectedFiles.length) return;

    const invalidFiles = selectedFiles.filter(
      (file) => !file.name.toLowerCase().endsWith(".pdf") && file.type !== "application/pdf",
    );
    const existingKeys = new Set(documents.map((item) => fileKey(item.file)));
    const usedNames = new Set(
      documents
        .map((item) => normalizeZipFilename(item.zipFilename))
        .filter(Boolean)
        .map((name) => name.toLocaleLowerCase()),
    );
    const additions = selectedFiles
      .filter((file) => file.name.toLowerCase().endsWith(".pdf") || file.type === "application/pdf")
      .filter((file) => !existingKeys.has(fileKey(file)))
      .map((file) => {
        const zipFilename = uniqueZipFilename(defaultZipFilename(file), usedNames);
        usedNames.add(zipFilename.toLocaleLowerCase());
        return createDocument(file, zipFilename);
      });

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
    additions.forEach((item) => generateThumbnail(item.id, item.file));
  };

  const updateDocument = (documentId, changes) => {
    setDocuments((current) => current.map((item) => (item.id === documentId ? { ...item, ...changes } : item)));
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

  const removeDocument = (documentId) => {
    setDocuments((current) => current.filter((item) => item.id !== documentId));
    setResult(null);
    setError("");
  };

  const clearDocuments = () => {
    setDocuments([]);
    setResult(null);
    setError("");
  };

  const changeZipFilename = (documentId, value) => {
    setDocuments((current) => current.map((item) => (item.id === documentId ? { ...item, zipFilename: value } : item)));
    setResult(null);
  };

  const generate = async () => {
    if (!canGenerate) return;
    setGenerating(true);
    setError("");
    setResult(null);

    try {
      const formData = new FormData();
      documents.forEach((item) => formData.append("pdfs", item.file, item.file.name));
      formData.append("zipFilenames", JSON.stringify(documents.map((item) => item.zipFilename.trim())));
      const response = await fetch(`${API_BASE}/pdf-document-zip-tool/generate-batch`, {
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
            sourceName: documents[index]?.file.name || item.zipName,
          }))
          : [],
      });
    } catch (err) {
      setError(err?.message || "一括ZIP生成に失敗しました。");
    } finally {
      setGenerating(false);
    }
  };

  return (
    <div className="lecture-tool-page">
      <div className="lecture-tool-page__inner">
        <header className="lecture-tool-header">
          <div>
            <h1 className="lecture-tool-header__title">PDFスライド</h1>
            <div className="lecture-tool-header__sub">複数PDFからpdf.pdf形式のZIPを生成します。</div>
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
                  <div className="lecture-tool-panel__sub">ZIP内ではPDF本体をpdf.pdfとして格納します。</div>
                </div>
                {documents.length ? (
                  <button
                    className="lecture-tool-button lecture-tool-button--small"
                    type="button"
                    onClick={clearDocuments}
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
                  {documents.length ? "追加後もPDFを足せます。" : "複数PDFをまとめて処理できます。"}
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
                    <h2 className="lecture-tool-panel__title">PDFごとのZIP設定</h2>
                    <div className="lecture-tool-panel__sub">追加したPDFごとに出力ZIP名を指定します。</div>
                  </div>
                  <span className="lecture-tool-progress__badge">{documents.length}件</span>
                </div>

                <div className="pdf-slide-document-list">
                  {documents.map((documentItem, documentIndex) => {
                    const safeName = normalizeZipFilename(documentItem.zipFilename);

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
                              <div className="lecture-tool-panel__sub">
                                {formatSize(documentItem.file.size)} / ZIP内のPDF名: pdf.pdf
                              </div>
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
                            ZIPファイル名
                            <input
                              className="lecture-tool-search"
                              value={documentItem.zipFilename}
                              onChange={(event) => changeZipFilename(documentItem.id, event.target.value)}
                              placeholder="minjuvi-fl_PDF_007_01"
                              disabled={generating}
                            />
                          </label>
                          <label className="lecture-tool-select-label">
                            出力ZIP名
                            <div className="pdf-slide-name-preview">
                              {safeName ? `${safeName}.zip` : "zip_filename.zip"}
                            </div>
                          </label>
                        </div>
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
                  <div className="lecture-tool-panel__sub">個別ZIPと一括ZIPを分けてダウンロードできます。</div>
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
                      <div className="lecture-tool-progress__current">PDF本体とプレビュー画像をZIP化しています。</div>
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
                        <div className="pdf-slide-result" key={item.filename}>
                          <div>
                            <div className="pdf-slide-result__title">{item.filename}</div>
                            <div className="pdf-slide-result__source">{item.sourceName}</div>
                            <div className="pdf-slide-result__meta">
                              {item.pageCount}ページ / {formatSize(item.size)}
                            </div>
                          </div>
                          <a
                            className="lecture-tool-button lecture-tool-button--primary"
                            href={downloadUrl(item.downloadUrl)}
                            download={item.filename}
                          >
                            個別ZIP
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
                          {result.pdfCount}PDF / {result.pageCount}ページ / 個別ZIPを格納 / {formatSize(result.batch.size)}
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
                <div className="lecture-tool-empty">PDFを追加し、PDFごとのZIPファイル名を指定してください。</div>
              )}
            </div>
          </section>
        </div>
      </div>
    </div>
  );
}
