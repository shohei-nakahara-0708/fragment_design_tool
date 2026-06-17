import React, { useMemo, useRef, useState } from "react";

const API_BASE = import.meta.env.VITE_API_BASE || "";
const INVALID_FILENAME_CHARS = new Set(["<", ">", ":", "\"", "/", "\\", "|", "?", "*"]);

function formatSize(bytes) {
  if (!bytes) return "-";
  const mb = bytes / 1024 / 1024;
  if (mb < 1024) return `${mb.toFixed(1)} MB`;
  return `${(mb / 1024).toFixed(2)} GB`;
}

function errorMessage(payload, fallback) {
  if (!payload) return fallback;
  if (typeof payload === "string") return payload;
  if (typeof payload.detail === "string") return payload.detail;
  return fallback;
}

function normalizeZipBase(value) {
  return String(value || "")
    .trim()
    .replace(/\.zip$/i, "")
    .split("")
    .map((char) => (char.charCodeAt(0) < 32 || INVALID_FILENAME_CHARS.has(char) ? "_" : char))
    .join("")
    .replace(/_+/g, "_")
    .replace(/^[ ._]+|[ ._]+$/g, "");
}

function defaultPackageBase(file) {
  return normalizeZipBase(String(file?.name || "").replace(/\.pptx$/i, "")) || "ppt_video";
}

function downloadUrl(value) {
  if (!value) return "";
  if (/^https?:\/\//i.test(value)) return value;
  return `${API_BASE}${value}`;
}

export default function PptVideoZipPage() {
  const fileInputRef = useRef(null);
  const [pptxFile, setPptxFile] = useState(null);
  const [dragOver, setDragOver] = useState(false);
  const [packageBase, setPackageBase] = useState("");
  const [analysis, setAnalysis] = useState(null);
  const [selectedPages, setSelectedPages] = useState([]);
  const [pageImages, setPageImages] = useState({});
  const [analyzing, setAnalyzing] = useState(false);
  const [generating, setGenerating] = useState(false);
  const [error, setError] = useState("");
  const [result, setResult] = useState(null);

  const safePackageBase = normalizeZipBase(packageBase);
  const videoSlides = Array.isArray(analysis?.slides) ? analysis.slides : [];
  const selectedSet = useMemo(() => new Set(selectedPages), [selectedPages]);
  const missingImagePages = selectedPages.filter((page) => !pageImages[page]?.file);
  const canAnalyze = Boolean(pptxFile && !analyzing && !generating);
  const canGenerate = Boolean(
    pptxFile &&
    videoSlides.length &&
    selectedPages.length &&
    !missingImagePages.length &&
    safePackageBase &&
    !analyzing &&
    !generating,
  );

  const clearPageImages = () => {
    setPageImages((current) => {
      Object.values(current).forEach((item) => {
        if (item?.previewUrl) URL.revokeObjectURL(item.previewUrl);
      });
      return {};
    });
  };

  const setFile = (file) => {
    if (!file) return;
    if (!file.name.toLowerCase().endsWith(".pptx")) {
      setError("PPTXを選択してください。");
      return;
    }
    setPptxFile(file);
    setPackageBase(defaultPackageBase(file));
    setAnalysis(null);
    setSelectedPages([]);
    clearPageImages();
    setResult(null);
    setError("");
  };

  const setPageImage = (page, file) => {
    if (!file) return;
    const validImage = file.type.startsWith("image/") || /\.(png|jpe?g|webp)$/i.test(file.name);
    if (!validImage) {
      setError(`${page}ページ目は画像ファイルを選択してください。`);
      return;
    }

    const previewUrl = URL.createObjectURL(file);
    setPageImages((current) => {
      if (current[page]?.previewUrl) URL.revokeObjectURL(current[page].previewUrl);
      return {
        ...current,
        [page]: { file, previewUrl },
      };
    });
    setResult(null);
    setError("");
  };

  const removePageImage = (page) => {
    setPageImages((current) => {
      if (current[page]?.previewUrl) URL.revokeObjectURL(current[page].previewUrl);
      const next = { ...current };
      delete next[page];
      return next;
    });
    setResult(null);
  };

  const analyze = async () => {
    if (!canAnalyze) return;
    setAnalyzing(true);
    setError("");
    setResult(null);

    try {
      const formData = new FormData();
      formData.append("pptx", pptxFile, pptxFile.name);
      const response = await fetch(`${API_BASE}/ppt-video-zip-tool/analyze`, {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        const payload = await response.json().catch(() => null);
        throw new Error(errorMessage(payload, "PPTX解析に失敗しました。"));
      }

      const payload = await response.json();
      const slides = Array.isArray(payload?.slides) ? payload.slides : [];
      setAnalysis({ ...payload, slides });
      setSelectedPages(slides.map((slide) => slide.page));
      clearPageImages();
      setPackageBase((current) => normalizeZipBase(current) || payload.packageBase || defaultPackageBase(pptxFile));
      if (!slides.length) {
        setError("動画が配置されているページが見つかりませんでした。");
      }
    } catch (err) {
      setAnalysis(null);
      setSelectedPages([]);
      setError(err?.message || "PPTX解析に失敗しました。");
    } finally {
      setAnalyzing(false);
    }
  };

  const togglePage = (page) => {
    setSelectedPages((current) => (
      current.includes(page)
        ? current.filter((item) => item !== page)
        : [...current, page].sort((a, b) => a - b)
    ));
    setResult(null);
  };

  const selectAll = () => {
    setSelectedPages(videoSlides.map((slide) => slide.page));
    setResult(null);
  };

  const clearSelection = () => {
    setSelectedPages([]);
    setResult(null);
  };

  const generate = async () => {
    if (!canGenerate) return;
    setGenerating(true);
    setError("");
    setResult(null);

    try {
      const formData = new FormData();
      formData.append("pptx", pptxFile, pptxFile.name);
      formData.append("selectedPages", JSON.stringify(selectedPages));
      selectedPages.forEach((page) => {
        const image = pageImages[page];
        if (image?.file) {
          formData.append("pageImages", image.file, `page-${page}-image.png`);
        }
      });
      formData.append("pageImagePages", JSON.stringify(selectedPages));
      formData.append("packageBase", safePackageBase);
      const response = await fetch(`${API_BASE}/ppt-video-zip-tool/generate`, {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        const payload = await response.json().catch(() => null);
        throw new Error(errorMessage(payload, "ZIP生成に失敗しました。"));
      }

      setResult(await response.json());
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
            <h1 className="lecture-tool-header__title">PPT動画ZIP生成</h1>
            <div className="lecture-tool-header__sub">PPTXから動画ページを解析し、アップロードしたimage.pngで動画入りスライドZIPを生成します。</div>
          </div>
          <div className="lecture-tool-actions pdf-slide-actions">
            <button className="lecture-tool-button" type="button" onClick={analyze} disabled={!canAnalyze}>
              {analyzing ? "解析中" : "動画ページ解析"}
            </button>
            <button className="lecture-tool-button lecture-tool-button--primary" type="button" onClick={generate} disabled={!canGenerate}>
              {generating ? "生成中" : "ZIPを生成"}
            </button>
          </div>
        </header>

        {error ? <div className="lecture-tool-alert">{error}</div> : null}
        {missingImagePages.length ? (
          <div className="lecture-tool-alert">
            生成対象ページのimage.pngをアップロードしてください: {missingImagePages.join(", ")}
          </div>
        ) : null}

        <div className="pdf-slide-grid">
          <section className="lecture-tool-panel lecture-tool-panel--wide">
            <div className="lecture-tool-panel__inner">
              <div className="lecture-tool-panel__head">
                <div>
                  <h2 className="lecture-tool-panel__title">PPTX追加</h2>
                  <div className="lecture-tool-panel__sub">動画付きページを解析して、ページごとのZIPを作成します。</div>
                </div>
                {pptxFile ? (
                  <button
                    className="lecture-tool-button lecture-tool-button--small"
                    type="button"
                    onClick={() => {
                      setPptxFile(null);
                      setPackageBase("");
                      setAnalysis(null);
                      setSelectedPages([]);
                      clearPageImages();
                      setResult(null);
                      setError("");
                    }}
                    disabled={analyzing || generating}
                  >
                    削除
                  </button>
                ) : null}
              </div>

              <div
                className={`lecture-tool-drop${dragOver ? " lecture-tool-drop--active" : ""}${analyzing || generating ? " lecture-tool-drop--disabled" : ""}`}
                onClick={() => {
                  if (!analyzing && !generating) fileInputRef.current?.click();
                }}
                onDragOver={(event) => {
                  event.preventDefault();
                  if (!analyzing && !generating) setDragOver(true);
                }}
                onDragLeave={() => setDragOver(false)}
                onDrop={(event) => {
                  event.preventDefault();
                  setDragOver(false);
                  if (!analyzing && !generating) setFile(event.dataTransfer.files?.[0]);
                }}
                role="button"
                tabIndex={0}
              >
                <div className="lecture-tool-drop__title">
                  {pptxFile ? pptxFile.name : "PPTXを選択 または ドラッグ&ドロップ"}
                </div>
                <div className="lecture-tool-drop__sub">
                  {pptxFile ? `${formatSize(pptxFile.size)} / 解析後にページごとのimage.pngを追加します。` : "動画を含むPPTXに対応しています。"}
                </div>
                <input
                  ref={fileInputRef}
                  className="lecture-tool-hidden-input"
                  type="file"
                  accept=".pptx,application/vnd.openxmlformats-officedocument.presentationml.presentation"
                  onChange={(event) => {
                    setFile(event.target.files?.[0]);
                    event.target.value = "";
                  }}
                  disabled={analyzing || generating}
                />
              </div>

              {pptxFile ? (
                <div className="ppt-video-file-summary">
                  <div className="pdf-slide-name-preview">{safePackageBase || "zip_base"} 13.zip</div>
                  <button className="lecture-tool-button lecture-tool-button--primary" type="button" onClick={analyze} disabled={!canAnalyze}>
                    {analyzing ? "解析中" : "動画ページ解析"}
                  </button>
                </div>
              ) : null}
            </div>
          </section>

          <section className="lecture-tool-panel lecture-tool-panel--wide">
            <div className="lecture-tool-panel__inner">
              <div className="lecture-tool-panel__head">
                <div>
                  <h2 className="lecture-tool-panel__title">ZIP設定</h2>
                  <div className="lecture-tool-panel__sub">ZIPファイル名、生成対象ページ、ページごとのimage.pngを指定します。</div>
                </div>
                {videoSlides.length ? <span className="lecture-tool-progress__badge">{videoSlides.length}ページ</span> : null}
              </div>

              <div className="ppt-video-settings">
                <label className="lecture-tool-select-label">
                  ZIPファイル名ベース
                  <input
                    className="lecture-tool-search"
                    value={packageBase}
                    onChange={(event) => {
                      setPackageBase(event.target.value);
                      setResult(null);
                    }}
                    placeholder="JPN-TSL-0420"
                    disabled={generating}
                  />
                </label>
                <label className="lecture-tool-select-label">
                  出力例
                  <div className="pdf-slide-name-preview">
                    {safePackageBase ? `${safePackageBase} 13.zip` : "JPN-TSL-0420 13.zip"}
                  </div>
                </label>
              </div>

              {analyzing ? (
                <div className="lecture-tool-progress ppt-video-progress">
                  <div className="lecture-tool-progress__head">
                    <div>
                      <div className="lecture-tool-progress__title">動画ページ解析中</div>
                      <div className="lecture-tool-progress__current">PPTX内の動画ファイルと配置情報を確認しています。</div>
                    </div>
                    <span className="lecture-tool-progress__badge">実行中</span>
                  </div>
                </div>
              ) : videoSlides.length ? (
                <div className="ppt-video-slide-list">
                  <div className="ppt-video-selection-actions">
                    <button className="lecture-tool-button lecture-tool-button--small" type="button" onClick={selectAll} disabled={generating}>
                      すべて選択
                    </button>
                    <button className="lecture-tool-button lecture-tool-button--small" type="button" onClick={clearSelection} disabled={generating}>
                      選択解除
                    </button>
                  </div>
                  {videoSlides.map((slide) => {
                    const selected = selectedSet.has(slide.page);
                    const sampleName = safePackageBase ? `${safePackageBase} ${slide.page}.zip` : `slide ${slide.page}.zip`;

                    const imageItem = pageImages[slide.page];
                    const imageInputId = `ppt-video-image-${slide.page}`;

                    return (
                      <div className={`ppt-video-slide-row${selected ? " ppt-video-slide-row--active" : ""}`} key={slide.page}>
                        <label className="ppt-video-slide-row__check" aria-label={`${slide.page}ページを生成対象にする`}>
                          <input
                            type="checkbox"
                            checked={selected}
                            onChange={() => togglePage(slide.page)}
                            disabled={generating}
                          />
                        </label>
                        <span className="ppt-video-slide-row__page">P{slide.page}</span>
                        <span className="ppt-video-slide-row__body">
                          <span className="ppt-video-slide-row__title">{sampleName}</span>
                          <span className="ppt-video-slide-row__meta">
                            {slide.videoCount || slide.videos?.length || 1}動画 / アップロード画像は生成時に動画部分を白抜き
                          </span>
                        </span>
                        <span className="ppt-video-slide-row__image">
                          <span className="ppt-video-slide-row__preview">
                            {imageItem?.previewUrl ? (
                              <img src={imageItem.previewUrl} alt={`${slide.page}ページ目のimage.png`} />
                            ) : (
                              <span>image.png</span>
                            )}
                          </span>
                          <span className="ppt-video-slide-row__image-body">
                            <span className={`ppt-video-slide-row__image-name${imageItem ? "" : " ppt-video-slide-row__image-name--missing"}`}>
                              {imageItem?.file?.name || "未アップロード"}
                            </span>
                            <span className="ppt-video-slide-row__image-actions">
                              <input
                                id={imageInputId}
                                className="lecture-tool-hidden-input"
                                type="file"
                                accept="image/png,image/jpeg,image/webp"
                                onChange={(event) => {
                                  setPageImage(slide.page, event.target.files?.[0]);
                                  event.target.value = "";
                                }}
                                disabled={generating}
                              />
                              <label className="lecture-tool-button lecture-tool-button--small" htmlFor={imageInputId}>
                                画像選択
                              </label>
                              {imageItem ? (
                                <button
                                  className="lecture-tool-button lecture-tool-button--small"
                                  type="button"
                                  onClick={() => removePageImage(slide.page)}
                                  disabled={generating}
                                >
                                  削除
                                </button>
                              ) : null}
                            </span>
                          </span>
                        </span>
                      </div>
                    );
                  })}
                </div>
              ) : (
                <div className="lecture-tool-empty">PPTXを追加して動画ページ解析を実行してください。</div>
              )}
            </div>
          </section>

          <section className="lecture-tool-panel lecture-tool-panel--wide pdf-slide-result-panel">
            <div className="lecture-tool-panel__inner">
              <div className="lecture-tool-panel__head">
                <div>
                  <h2 className="lecture-tool-panel__title">生成結果</h2>
                  <div className="lecture-tool-panel__sub">ページごとのZIPと、一括ZIPを分けてダウンロードできます。</div>
                </div>
                <button className="lecture-tool-button lecture-tool-button--primary" type="button" onClick={generate} disabled={!canGenerate}>
                  {generating ? "生成中" : "ZIPを生成"}
                </button>
              </div>

              {generating ? (
                <div className="lecture-tool-progress">
                  <div className="lecture-tool-progress__head">
                    <div>
                      <div className="lecture-tool-progress__title">ZIP生成中</div>
                      <div className="lecture-tool-progress__current">アップロード画像の白抜き、動画抽出、ZIP作成を実行しています。</div>
                    </div>
                    <span className="lecture-tool-progress__badge">実行中</span>
                  </div>
                </div>
              ) : result ? (
                <div className="pdf-slide-result-groups">
                  <section className="pdf-slide-result-group">
                    <h3 className="pdf-slide-result-group__title">ページごと</h3>
                    <div className="pdf-slide-result-list">
                      {result.results.map((item) => (
                        <div className="pdf-slide-result" key={item.filename}>
                          <div>
                            <div className="pdf-slide-result__title">{item.filename}</div>
                            <div className="pdf-slide-result__meta">
                              {item.videoCount}動画 / {formatSize(item.size)}
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
                    <h3 className="pdf-slide-result-group__title">一括</h3>
                    <div className="pdf-slide-result pdf-slide-result--batch">
                      <div>
                        <div className="pdf-slide-result__title">{result.batch.filename}</div>
                        <div className="pdf-slide-result__meta">
                          {result.pageCount}ZIPを格納 / {formatSize(result.batch.size)}
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
                <div className="lecture-tool-empty">動画ページを解析し、生成対象ページを選択してください。</div>
              )}
            </div>
          </section>
        </div>
      </div>
    </div>
  );
}
