import React, { useEffect, useMemo, useRef, useState } from "react";

const API_BASE = import.meta.env.VITE_API_BASE || "";
const INVALID_FILENAME_CHARS = new Set(["<", ">", ":", "\"", "/", "\\", "|", "?", "*"]);
const AUTO_HEIGHTS = [1080, 720, 540, 480, 360, 270, 240];

function formatSize(bytes) {
  if (!bytes) return "-";
  const mb = bytes / 1024 / 1024;
  if (mb < 1024) return `${mb.toFixed(1)} MB`;
  return `${(mb / 1024).toFixed(2)} GB`;
}

function formatDuration(seconds) {
  if (!Number.isFinite(seconds) || seconds <= 0) return "-";
  const hours = Math.floor(seconds / 3600);
  const minutes = Math.floor((seconds % 3600) / 60);
  const rest = Math.floor(seconds % 60);
  if (hours) return `${hours}:${String(minutes).padStart(2, "0")}:${String(rest).padStart(2, "0")}`;
  return `${minutes}:${String(rest).padStart(2, "0")}`;
}

function formatBitrate(kbps) {
  if (!Number.isFinite(kbps) || kbps <= 0) return "-";
  if (kbps >= 1000) return `${(kbps / 1000).toFixed(2)} Mbps`;
  return `${Math.round(kbps)} kbps`;
}

function normalizeFilename(value) {
  return String(value || "")
    .trim()
    .replace(/\.(mp4|mov|m4v|webm)$/i, "")
    .split("")
    .map((char) => (char.charCodeAt(0) < 32 || INVALID_FILENAME_CHARS.has(char) ? "_" : char))
    .join("")
    .replace(/_+/g, "_")
    .replace(/^[ ._]+|[ ._]+$/g, "");
}

function defaultOutputName(file) {
  const stem = normalizeFilename(String(file?.name || "").replace(/\.[^.]+$/i, ""));
  return stem ? `${stem}_encoded` : "encoded_video";
}

function errorMessage(payload, fallback) {
  if (!payload) return fallback;
  if (typeof payload === "string") return payload;
  if (typeof payload.detail === "string") return payload.detail;
  return fallback;
}

function autoAudioKbps(totalKbps, hasAudio = true) {
  if (!hasAudio) return 0;
  if (totalKbps < 80) return Math.max(0, totalKbps - 40);
  if (totalKbps < 160) return 32;
  if (totalKbps < 260) return 48;
  return 64;
}

function even(value) {
  const rounded = Math.round(value);
  return Math.max(2, rounded - (rounded % 2));
}

function estimateSettings({ file, duration, width, height, targetMegabytes, maxHeight, audioBitrateKbps }) {
  if (!file || !Number.isFinite(duration) || duration <= 0) return null;

  const targetTotalKbps = Math.max(1, Math.floor((Number(targetMegabytes || 0) * 8192 * 0.96) / duration));
  const originalKbps = Math.round((file.size * 8) / duration / 1000);
  const totalKbps = Math.min(targetTotalKbps, Math.max(1, Math.floor(originalKbps * 0.92)));
  const audioKbps = audioBitrateKbps === "auto" ? autoAudioKbps(totalKbps) : Number(audioBitrateKbps || 0);
  const videoKbps = totalKbps - audioKbps;

  let outputHeight = height || 0;
  let outputWidth = width || 0;
  if (width && height) {
    const aspect = width / height;
    const sizeForHeight = (candidateHeight) => {
      const nextHeight = Math.min(even(candidateHeight), even(height));
      return {
        width: even(nextHeight * aspect),
        height: nextHeight,
      };
    };

    if (maxHeight !== "auto") {
      const size = sizeForHeight(Number(maxHeight));
      outputHeight = size.height;
      outputWidth = size.width;
    } else {
      const candidates = Array.from(new Set([height, ...AUTO_HEIGHTS].filter((item) => item <= height))).sort((a, b) => b - a);
      const fps = 30;
      const selected = candidates.find((candidateHeight) => {
        const size = sizeForHeight(candidateHeight);
        return (videoKbps * 1000) / Math.max(1, size.width * size.height * fps) >= 0.02;
      }) || candidates[candidates.length - 1] || height;
      const size = sizeForHeight(selected);
      outputHeight = size.height;
      outputWidth = size.width;
    }
  }

  return {
    totalKbps,
    audioKbps,
    videoKbps,
    originalKbps,
    outputWidth,
    outputHeight,
  };
}

function filenameFromDisposition(value, fallback) {
  const text = value || "";
  const encodedMatch = text.match(/filename\*=UTF-8''([^;]+)/i);
  if (encodedMatch) {
    try {
      return decodeURIComponent(encodedMatch[1]);
    } catch {
      return fallback;
    }
  }
  const plainMatch = text.match(/filename="?([^";]+)"?/i);
  return plainMatch?.[1] || fallback;
}

function numberHeader(headers, name) {
  const value = Number(headers.get(name));
  return Number.isFinite(value) ? value : 0;
}

function revokeResult(result) {
  if (result?.url) URL.revokeObjectURL(result.url);
}

export default function VideoEncodePage() {
  const fileInputRef = useRef(null);
  const [videoFile, setVideoFile] = useState(null);
  const [videoUrl, setVideoUrl] = useState("");
  const [metadata, setMetadata] = useState({ duration: 0, width: 0, height: 0 });
  const [targetMegabytes, setTargetMegabytes] = useState(50);
  const [maxMegabytes, setMaxMegabytes] = useState(80);
  const [maxHeight, setMaxHeight] = useState("auto");
  const [audioBitrateKbps, setAudioBitrateKbps] = useState("auto");
  const [preset, setPreset] = useState("medium");
  const [outputName, setOutputName] = useState("");
  const [dragOver, setDragOver] = useState(false);
  const [encoding, setEncoding] = useState(false);
  const [error, setError] = useState("");
  const [result, setResult] = useState(null);

  const safeOutputName = normalizeFilename(outputName);
  const estimate = useMemo(
    () => estimateSettings({
      file: videoFile,
      duration: metadata.duration,
      width: metadata.width,
      height: metadata.height,
      targetMegabytes,
      maxHeight,
      audioBitrateKbps,
    }),
    [audioBitrateKbps, maxHeight, metadata.duration, metadata.height, metadata.width, targetMegabytes, videoFile],
  );
  const canEncode = Boolean(videoFile && !encoding && safeOutputName && Number(targetMegabytes) > 0 && Number(maxMegabytes) > 0);
  const severeCompression = Boolean(estimate && estimate.totalKbps < 180);

  useEffect(() => () => {
    if (videoUrl) URL.revokeObjectURL(videoUrl);
  }, [videoUrl]);

  useEffect(() => () => {
    revokeResult(result);
  }, [result]);

  const clearResult = () => {
    setResult((current) => {
      revokeResult(current);
      return null;
    });
  };

  const setFile = (file) => {
    if (!file) return;
    const validVideo = file.type.startsWith("video/") || /\.(mp4|mov|m4v|webm|ogv|avi|mkv)$/i.test(file.name);
    if (!validVideo) {
      setError("動画ファイルを選択してください。");
      return;
    }

    const nextUrl = URL.createObjectURL(file);
    setVideoUrl((current) => {
      if (current) URL.revokeObjectURL(current);
      return nextUrl;
    });
    setVideoFile(file);
    setMetadata({ duration: 0, width: 0, height: 0 });
    setOutputName(defaultOutputName(file));
    clearResult();
    setError("");
  };

  const clearVideo = () => {
    setVideoUrl((current) => {
      if (current) URL.revokeObjectURL(current);
      return "";
    });
    setVideoFile(null);
    setMetadata({ duration: 0, width: 0, height: 0 });
    setOutputName("");
    clearResult();
    setError("");
  };

  const encodeVideo = async () => {
    if (!canEncode) return;
    setEncoding(true);
    setError("");
    clearResult();

    try {
      const formData = new FormData();
      formData.append("video", videoFile, videoFile.name);
      formData.append("targetMegabytes", String(targetMegabytes));
      formData.append("maxMegabytes", String(maxMegabytes));
      formData.append("maxHeight", maxHeight);
      formData.append("audioBitrateKbps", audioBitrateKbps);
      formData.append("outputName", safeOutputName);
      formData.append("preset", preset);

      const response = await fetch(`${API_BASE}/video-encode-tool/encode`, {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        const payload = await response.json().catch(() => null);
        throw new Error(errorMessage(payload, "動画エンコードに失敗しました。"));
      }

      const blob = await response.blob();
      const filename = filenameFromDisposition(response.headers.get("Content-Disposition"), `${safeOutputName}.mp4`);
      const nextResult = {
        url: URL.createObjectURL(blob),
        filename,
        size: blob.size,
        duration: numberHeader(response.headers, "X-Duration"),
        originalSize: numberHeader(response.headers, "X-Original-Size"),
        originalWidth: numberHeader(response.headers, "X-Original-Width"),
        originalHeight: numberHeader(response.headers, "X-Original-Height"),
        outputWidth: numberHeader(response.headers, "X-Output-Width"),
        outputHeight: numberHeader(response.headers, "X-Output-Height"),
        originalBitrateKbps: numberHeader(response.headers, "X-Original-Bitrate-Kbps"),
        totalBitrateKbps: numberHeader(response.headers, "X-Total-Bitrate-Kbps"),
        videoBitrateKbps: numberHeader(response.headers, "X-Video-Bitrate-Kbps"),
        audioBitrateKbps: numberHeader(response.headers, "X-Audio-Bitrate-Kbps"),
        overMax: response.headers.get("X-Over-Max") === "1" || blob.size > Number(maxMegabytes) * 1024 * 1024,
      };
      setResult(nextResult);
    } catch (err) {
      setError(err?.message || "動画エンコードに失敗しました。");
    } finally {
      setEncoding(false);
    }
  };

  return (
    <div className="lecture-tool-page">
      <div className="lecture-tool-page__inner">
        <header className="lecture-tool-header">
          <div>
            <h1 className="lecture-tool-header__title">動画エンコード</h1>
            <div className="lecture-tool-header__sub">目標サイズに合わせてMP4へ圧縮します。</div>
          </div>
          <div className="lecture-tool-actions pdf-slide-actions">
            <button className="lecture-tool-button lecture-tool-button--primary" type="button" onClick={encodeVideo} disabled={!canEncode}>
              {encoding ? "エンコード中" : "エンコード開始"}
            </button>
          </div>
        </header>

        {error ? <div className="lecture-tool-alert">{error}</div> : null}

        <div className="video-encode-grid">
          <section className="lecture-tool-panel">
            <div className="lecture-tool-panel__inner">
              <div className="lecture-tool-panel__head">
                <div>
                  <h2 className="lecture-tool-panel__title">動画追加</h2>
                  <div className="lecture-tool-panel__sub">長尺動画はエンコードに時間がかかります。</div>
                </div>
                {videoFile ? (
                  <button className="lecture-tool-button lecture-tool-button--small" type="button" onClick={clearVideo} disabled={encoding}>
                    削除
                  </button>
                ) : null}
              </div>

              <div
                className={`lecture-tool-drop${dragOver ? " lecture-tool-drop--active" : ""}${encoding ? " lecture-tool-drop--disabled" : ""}`}
                onClick={() => {
                  if (!encoding) fileInputRef.current?.click();
                }}
                onDragOver={(event) => {
                  event.preventDefault();
                  if (!encoding) setDragOver(true);
                }}
                onDragLeave={() => setDragOver(false)}
                onDrop={(event) => {
                  event.preventDefault();
                  setDragOver(false);
                  if (!encoding) setFile(event.dataTransfer.files?.[0]);
                }}
                role="button"
                tabIndex={0}
              >
                <div className="lecture-tool-drop__title">
                  {videoFile ? videoFile.name : "動画を選択 または ドラッグ&ドロップ"}
                </div>
                <div className="lecture-tool-drop__sub">
                  {videoFile ? `${formatSize(videoFile.size)} / ${formatDuration(metadata.duration)}` : "MP4、MOV、WebMなどに対応しています。"}
                </div>
                <input
                  ref={fileInputRef}
                  className="lecture-tool-hidden-input"
                  type="file"
                  accept="video/*,.mp4,.mov,.m4v,.webm,.ogv,.avi,.mkv"
                  onChange={(event) => {
                    setFile(event.target.files?.[0]);
                    event.target.value = "";
                  }}
                  disabled={encoding}
                />
              </div>

              {videoUrl ? (
                <div className="video-encode-player">
                  <video
                    src={videoUrl}
                    controls
                    playsInline
                    preload="metadata"
                    onLoadedMetadata={(event) => {
                      const video = event.currentTarget;
                      setMetadata({
                        duration: Number.isFinite(video.duration) ? video.duration : 0,
                        width: video.videoWidth || 0,
                        height: video.videoHeight || 0,
                      });
                    }}
                  />
                </div>
              ) : null}
            </div>
          </section>

          <section className="lecture-tool-panel">
            <div className="lecture-tool-panel__inner">
              <div className="lecture-tool-panel__head">
                <div>
                  <h2 className="lecture-tool-panel__title">圧縮設定</h2>
                  <div className="lecture-tool-panel__sub">初期値は50MB目標、80MB上限です。</div>
                </div>
              </div>

              <div className="video-encode-settings">
                <div className="video-encode-field-row">
                  <label className="lecture-tool-select-label">
                    目標サイズ MB
                    <input
                      className="lecture-tool-search"
                      type="number"
                      min="1"
                      step="1"
                      value={targetMegabytes}
                      onChange={(event) => {
                        setTargetMegabytes(event.target.value);
                        clearResult();
                      }}
                      disabled={encoding}
                    />
                  </label>
                  <label className="lecture-tool-select-label">
                    上限サイズ MB
                    <input
                      className="lecture-tool-search"
                      type="number"
                      min="1"
                      step="1"
                      value={maxMegabytes}
                      onChange={(event) => {
                        setMaxMegabytes(event.target.value);
                        clearResult();
                      }}
                      disabled={encoding}
                    />
                  </label>
                </div>

                <div className="video-encode-field-row">
                  <label className="lecture-tool-select-label">
                    解像度
                    <select
                      className="lecture-tool-select"
                      value={maxHeight}
                      onChange={(event) => {
                        setMaxHeight(event.target.value);
                        clearResult();
                      }}
                      disabled={encoding}
                    >
                      <option value="auto">自動</option>
                      <option value="1080">1080p以下</option>
                      <option value="720">720p以下</option>
                      <option value="540">540p以下</option>
                      <option value="480">480p以下</option>
                      <option value="360">360p以下</option>
                      <option value="270">270p以下</option>
                      <option value="240">240p以下</option>
                    </select>
                  </label>

                  <label className="lecture-tool-select-label">
                    音声
                    <select
                      className="lecture-tool-select"
                      value={audioBitrateKbps}
                      onChange={(event) => {
                        setAudioBitrateKbps(event.target.value);
                        clearResult();
                      }}
                      disabled={encoding}
                    >
                      <option value="auto">自動</option>
                      <option value="24">24 kbps</option>
                      <option value="32">32 kbps</option>
                      <option value="48">48 kbps</option>
                      <option value="64">64 kbps</option>
                      <option value="96">96 kbps</option>
                      <option value="128">128 kbps</option>
                      <option value="0">音声なし</option>
                    </select>
                  </label>
                </div>

                <div className="video-encode-field-row">
                  <label className="lecture-tool-select-label">
                    エンコード速度
                    <select
                      className="lecture-tool-select"
                      value={preset}
                      onChange={(event) => setPreset(event.target.value)}
                      disabled={encoding}
                    >
                      <option value="fast">速い</option>
                      <option value="medium">標準</option>
                      <option value="slow">高圧縮</option>
                    </select>
                  </label>

                  <label className="lecture-tool-select-label">
                    出力ファイル名
                    <input
                      className="lecture-tool-search"
                      value={outputName}
                      onChange={(event) => {
                        setOutputName(event.target.value);
                        clearResult();
                      }}
                      placeholder="encoded_video"
                      disabled={encoding}
                    />
                  </label>
                </div>

                {estimate ? (
                  <div className="video-encode-estimate">
                    <div className="video-encode-metric">
                      <span>元</span>
                      <strong>{formatSize(videoFile.size)}</strong>
                      <em>{formatBitrate(estimate.originalKbps)}</em>
                    </div>
                    <div className="video-encode-metric">
                      <span>目標</span>
                      <strong>{formatSize(Number(targetMegabytes) * 1024 * 1024)}</strong>
                      <em>{formatBitrate(estimate.totalKbps)}</em>
                    </div>
                    <div className="video-encode-metric">
                      <span>映像</span>
                      <strong>{formatBitrate(estimate.videoKbps)}</strong>
                      <em>{estimate.outputWidth && estimate.outputHeight ? `${estimate.outputWidth} x ${estimate.outputHeight}` : "-"}</em>
                    </div>
                    <div className="video-encode-metric">
                      <span>音声</span>
                      <strong>{formatBitrate(estimate.audioKbps)}</strong>
                      <em>{formatDuration(metadata.duration)}</em>
                    </div>
                  </div>
                ) : null}

                {severeCompression ? (
                  <div className="lecture-tool-hint">50MB前後の長尺動画はかなり低いビットレートになります。自動解像度では画質の破綻を避けるため、360p以下まで下げることがあります。</div>
                ) : null}
              </div>
            </div>
          </section>

          <section className="lecture-tool-panel lecture-tool-panel--wide pdf-slide-result-panel">
            <div className="lecture-tool-panel__inner">
              <div className="lecture-tool-panel__head">
                <div>
                  <h2 className="lecture-tool-panel__title">生成結果</h2>
                  <div className="lecture-tool-panel__sub">エンコード後のMP4を確認してダウンロードできます。</div>
                </div>
                <button className="lecture-tool-button lecture-tool-button--primary" type="button" onClick={encodeVideo} disabled={!canEncode}>
                  {encoding ? "エンコード中" : result ? "再エンコード" : "エンコード開始"}
                </button>
              </div>

              {encoding ? (
                <div className="lecture-tool-progress">
                  <div className="lecture-tool-progress__head">
                    <div>
                      <div className="lecture-tool-progress__title">エンコード中</div>
                      <div className="lecture-tool-progress__current">長尺動画では数分以上かかることがあります。</div>
                    </div>
                    <span className="lecture-tool-progress__badge">実行中</span>
                  </div>
                </div>
              ) : result ? (
                <div className="video-encode-result">
                  <div className="video-encode-result__media">
                    <video src={result.url} controls playsInline />
                  </div>
                  <div className="video-encode-result__body">
                    {result.overMax ? (
                      <div className="lecture-tool-alert video-encode-result__alert">上限サイズを超えました。目標サイズまたは解像度を下げて再エンコードしてください。</div>
                    ) : null}
                    <div className="pdf-slide-result pdf-slide-result--batch">
                      <div>
                        <div className="pdf-slide-result__title">{result.filename}</div>
                        <div className="pdf-slide-result__meta">
                          {formatSize(result.size)} / {result.outputWidth} x {result.outputHeight} / {formatBitrate(result.totalBitrateKbps)}
                        </div>
                      </div>
                      <a className="lecture-tool-button lecture-tool-button--primary" href={result.url} download={result.filename}>
                        ダウンロード
                      </a>
                    </div>
                    <div className="video-encode-result__stats">
                      <div><strong>元サイズ</strong><span>{formatSize(result.originalSize)}</span></div>
                      <div><strong>圧縮後</strong><span>{formatSize(result.size)}</span></div>
                      <div><strong>映像</strong><span>{formatBitrate(result.videoBitrateKbps)}</span></div>
                      <div><strong>音声</strong><span>{formatBitrate(result.audioBitrateKbps)}</span></div>
                    </div>
                  </div>
                </div>
              ) : (
                <div className="lecture-tool-empty">動画を追加して圧縮設定を確認してください。</div>
              )}
            </div>
          </section>
        </div>
      </div>
    </div>
  );
}
