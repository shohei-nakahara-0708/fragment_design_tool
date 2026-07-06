import React, { useEffect, useMemo, useRef, useState } from "react";

const POSTER_SIZE = { width: 1024, height: 768 };
const THUMB_SIZE = { width: 200, height: 150 };
const FRAME_RATIO = 16 / 9;

function formatDuration(seconds) {
  if (!Number.isFinite(seconds)) return "0:00.0";
  const safeSeconds = Math.max(0, seconds);
  const minutes = Math.floor(safeSeconds / 60);
  const rest = safeSeconds - minutes * 60;
  return `${minutes}:${rest.toFixed(1).padStart(4, "0")}`;
}

function formatSize(bytes) {
  if (!bytes) return "-";
  const mb = bytes / 1024 / 1024;
  if (mb < 1024) return `${mb.toFixed(1)} MB`;
  return `${(mb / 1024).toFixed(2)} GB`;
}

function clamp(value, min, max) {
  if (!Number.isFinite(value)) return min;
  return Math.min(Math.max(value, min), max);
}

function revokeResult(result) {
  if (!result) return;
  if (result.posterUrl) URL.revokeObjectURL(result.posterUrl);
  if (result.thumbUrl) URL.revokeObjectURL(result.thumbUrl);
}

function canvasToBlob(canvas) {
  return new Promise((resolve, reject) => {
    canvas.toBlob((blob) => {
      if (blob) {
        resolve(blob);
      } else {
        reject(new Error("PNGを生成できませんでした。"));
      }
    }, "image/png");
  });
}

function waitForSeek(video, targetSeconds) {
  const duration = Number.isFinite(video.duration) ? video.duration : targetSeconds;
  const maxTarget = duration > 0.05 ? duration - 0.05 : duration;
  const safeTarget = clamp(targetSeconds, 0, maxTarget);

  return new Promise((resolve, reject) => {
    let settled = false;
    const timeoutId = window.setTimeout(() => {
      cleanup();
      reject(new Error("指定秒数への移動に失敗しました。"));
    }, 8000);

    const cleanup = () => {
      window.clearTimeout(timeoutId);
      video.removeEventListener("seeked", handleSeeked);
      video.removeEventListener("error", handleError);
    };

    const finish = () => {
      if (settled) return;
      settled = true;
      cleanup();
      resolve();
    };

    const handleSeeked = () => finish();
    const handleError = () => {
      if (settled) return;
      settled = true;
      cleanup();
      reject(new Error("動画を読み込めませんでした。"));
    };

    video.addEventListener("seeked", handleSeeked);
    video.addEventListener("error", handleError);

    if (Math.abs(video.currentTime - safeTarget) < 0.03 && video.readyState >= 2) {
      requestAnimationFrame(finish);
      return;
    }

    try {
      video.currentTime = safeTarget;
    } catch {
      cleanup();
      reject(new Error("指定秒数へ移動できませんでした。"));
    }
  });
}

function drawVideoCover(ctx, video, x, y, width, height) {
  const sourceWidth = video.videoWidth;
  const sourceHeight = video.videoHeight;
  if (!sourceWidth || !sourceHeight) {
    throw new Error("動画のサイズを取得できませんでした。");
  }

  const sourceRatio = sourceWidth / sourceHeight;
  let sx = 0;
  let sy = 0;
  let sw = sourceWidth;
  let sh = sourceHeight;

  if (sourceRatio > FRAME_RATIO) {
    sw = sourceHeight * FRAME_RATIO;
    sx = (sourceWidth - sw) / 2;
  } else if (sourceRatio < FRAME_RATIO) {
    sh = sourceWidth / FRAME_RATIO;
    sy = (sourceHeight - sh) / 2;
  }

  ctx.drawImage(video, sx, sy, sw, sh, x, y, width, height);
}

async function createThumbnailImages(video, seconds) {
  video.pause();
  await waitForSeek(video, seconds);

  const posterCanvas = document.createElement("canvas");
  posterCanvas.width = POSTER_SIZE.width;
  posterCanvas.height = POSTER_SIZE.height;
  const posterCtx = posterCanvas.getContext("2d");
  const frameHeight = POSTER_SIZE.width / FRAME_RATIO;
  const frameY = (POSTER_SIZE.height - frameHeight) / 2;

  posterCtx.fillStyle = "#000000";
  posterCtx.fillRect(0, 0, POSTER_SIZE.width, POSTER_SIZE.height);
  drawVideoCover(posterCtx, video, 0, frameY, POSTER_SIZE.width, frameHeight);

  const thumbCanvas = document.createElement("canvas");
  thumbCanvas.width = THUMB_SIZE.width;
  thumbCanvas.height = THUMB_SIZE.height;
  const thumbCtx = thumbCanvas.getContext("2d");
  thumbCtx.imageSmoothingQuality = "high";
  thumbCtx.drawImage(posterCanvas, 0, 0, THUMB_SIZE.width, THUMB_SIZE.height);

  const [posterBlob, thumbBlob] = await Promise.all([
    canvasToBlob(posterCanvas),
    canvasToBlob(thumbCanvas),
  ]);

  return {
    seconds,
    posterBlob,
    thumbBlob,
    posterUrl: URL.createObjectURL(posterBlob),
    thumbUrl: URL.createObjectURL(thumbBlob),
  };
}

export default function VideoThumbnailPage() {
  const fileInputRef = useRef(null);
  const videoRef = useRef(null);
  const [videoFile, setVideoFile] = useState(null);
  const [videoUrl, setVideoUrl] = useState("");
  const [duration, setDuration] = useState(0);
  const [metadataReady, setMetadataReady] = useState(false);
  const [captureSeconds, setCaptureSeconds] = useState(0);
  const [dragOver, setDragOver] = useState(false);
  const [generating, setGenerating] = useState(false);
  const [error, setError] = useState("");
  const [result, setResult] = useState(null);

  const maxSeconds = Number.isFinite(duration) && duration > 0 ? duration : 0;
  const safeCaptureSeconds = useMemo(
    () => clamp(captureSeconds, 0, maxSeconds || captureSeconds || 0),
    [captureSeconds, maxSeconds],
  );
  const canGenerate = Boolean(videoFile && videoUrl && metadataReady && !generating);

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
    const validVideo = file.type.startsWith("video/") || /\.(mp4|mov|m4v|webm|ogv)$/i.test(file.name);
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
    setDuration(0);
    setMetadataReady(false);
    setCaptureSeconds(0);
    clearResult();
    setError("");
  };

  const clearVideo = () => {
    setVideoUrl((current) => {
      if (current) URL.revokeObjectURL(current);
      return "";
    });
    setVideoFile(null);
    setDuration(0);
    setMetadataReady(false);
    setCaptureSeconds(0);
    clearResult();
    setError("");
  };

  const changeCaptureSeconds = (value) => {
    const numeric = Number.parseFloat(value);
    const nextSeconds = clamp(numeric, 0, maxSeconds || numeric || 0);
    setCaptureSeconds(Number.isFinite(nextSeconds) ? nextSeconds : 0);
    clearResult();

    const video = videoRef.current;
    if (video && Number.isFinite(nextSeconds)) {
      video.currentTime = nextSeconds;
    }
  };

  const useCurrentTime = () => {
    const video = videoRef.current;
    if (!video) return;
    changeCaptureSeconds(video.currentTime || 0);
  };

  const generate = async () => {
    const video = videoRef.current;
    if (!video || !canGenerate) return;
    setGenerating(true);
    setError("");
    clearResult();

    try {
      const nextResult = await createThumbnailImages(video, safeCaptureSeconds);
      setResult(nextResult);
    } catch (err) {
      setError(err?.message || "サムネイル生成に失敗しました。");
    } finally {
      setGenerating(false);
    }
  };

  return (
    <div className="lecture-tool-page">
      <div className="lecture-tool-page__inner">
        <header className="lecture-tool-header">
          <div>
            <h1 className="lecture-tool-header__title">動画サムネイル作成</h1>
            <div className="lecture-tool-header__sub">
              動画の指定秒数から、poster.pngとthumb.pngを生成します。
            </div>
          </div>
          <div className="lecture-tool-actions pdf-slide-actions">
            <button className="lecture-tool-button lecture-tool-button--primary" type="button" onClick={generate} disabled={!canGenerate}>
              {generating ? "生成中" : "サムネイル生成"}
            </button>
          </div>
        </header>

        {error ? <div className="lecture-tool-alert">{error}</div> : null}

        <div className="video-thumb-grid">
          <section className="lecture-tool-panel">
            <div className="lecture-tool-panel__inner">
              <div className="lecture-tool-panel__head">
                <div>
                  <h2 className="lecture-tool-panel__title">動画追加</h2>
                  <div className="lecture-tool-panel__sub">MP4、MOV、WebMなどの動画を選択できます。</div>
                </div>
                {videoFile ? (
                  <button className="lecture-tool-button lecture-tool-button--small" type="button" onClick={clearVideo} disabled={generating}>
                    削除
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
                  if (!generating) setFile(event.dataTransfer.files?.[0]);
                }}
                role="button"
                tabIndex={0}
              >
                <div className="lecture-tool-drop__title">
                  {videoFile ? videoFile.name : "動画を選択 または ドラッグ&ドロップ"}
                </div>
                <div className="lecture-tool-drop__sub">
                  {videoFile ? `${formatSize(videoFile.size)} / ${formatDuration(maxSeconds)}` : "指定秒数のフレームをPNGにします。"}
                </div>
                <input
                  ref={fileInputRef}
                  className="lecture-tool-hidden-input"
                  type="file"
                  accept="video/*,.mp4,.mov,.m4v,.webm,.ogv"
                  onChange={(event) => {
                    setFile(event.target.files?.[0]);
                    event.target.value = "";
                  }}
                  disabled={generating}
                />
              </div>

              {videoUrl ? (
                <div className="video-thumb-player">
                  <video
                    ref={videoRef}
                    src={videoUrl}
                    controls
                    playsInline
                    preload="metadata"
                    onLoadedMetadata={(event) => {
                      const nextDuration = event.currentTarget.duration;
                      setDuration(Number.isFinite(nextDuration) ? nextDuration : 0);
                      setMetadataReady(true);
                      setCaptureSeconds(0);
                    }}
                    onError={() => setError("この動画はブラウザで読み込めませんでした。")}
                  />
                </div>
              ) : null}
            </div>
          </section>

          <section className="lecture-tool-panel">
            <div className="lecture-tool-panel__inner">
              <div className="lecture-tool-panel__head">
                <div>
                  <h2 className="lecture-tool-panel__title">秒数指定</h2>
                  <div className="lecture-tool-panel__sub">生成したい場面の秒数を指定します。</div>
                </div>
                {videoFile ? <span className="lecture-tool-progress__badge">{formatDuration(safeCaptureSeconds)}</span> : null}
              </div>

              {videoFile ? (
                <div className="video-thumb-settings">
                  <label className="lecture-tool-select-label">
                    秒数
                    <input
                      className="lecture-tool-search"
                      type="number"
                      min="0"
                      max={maxSeconds || undefined}
                      step="0.1"
                      value={Number.isFinite(captureSeconds) ? captureSeconds : 0}
                      onChange={(event) => changeCaptureSeconds(event.target.value)}
                      disabled={generating}
                    />
                  </label>

                  <label className="lecture-tool-select-label">
                    シークバー
                    <input
                      className="video-thumb-range"
                      type="range"
                      min="0"
                      max={maxSeconds || 0}
                      step="0.1"
                      value={safeCaptureSeconds}
                      onChange={(event) => changeCaptureSeconds(event.target.value)}
                      disabled={generating || !maxSeconds}
                    />
                  </label>

                  <div className="video-thumb-actions">
                    <button className="lecture-tool-button" type="button" onClick={useCurrentTime} disabled={generating}>
                      現在位置を使う
                    </button>
                    <button className="lecture-tool-button lecture-tool-button--primary" type="button" onClick={generate} disabled={!canGenerate}>
                      {generating ? "生成中" : "サムネイル生成"}
                    </button>
                  </div>

                  <div className="video-thumb-specs">
                    <div>
                      <strong>poster.png</strong>
                      <span>1024 x 768</span>
                    </div>
                    <div>
                      <strong>thumb.png</strong>
                      <span>200 x 150</span>
                    </div>
                  </div>
                </div>
              ) : (
                <div className="lecture-tool-empty">動画を追加すると秒数を指定できます。</div>
              )}
            </div>
          </section>

          <section className="lecture-tool-panel lecture-tool-panel--wide pdf-slide-result-panel">
            <div className="lecture-tool-panel__inner">
              <div className="lecture-tool-panel__head">
                <div>
                  <h2 className="lecture-tool-panel__title">生成結果</h2>
                  <div className="lecture-tool-panel__sub">上下黒帯付きの4:3 PNGをダウンロードできます。</div>
                </div>
                <button className="lecture-tool-button lecture-tool-button--primary" type="button" onClick={generate} disabled={!canGenerate}>
                  {generating ? "生成中" : "再生成"}
                </button>
              </div>

              {generating ? (
                <div className="lecture-tool-progress">
                  <div className="lecture-tool-progress__head">
                    <div>
                      <div className="lecture-tool-progress__title">サムネイル生成中</div>
                      <div className="lecture-tool-progress__current">{formatDuration(safeCaptureSeconds)} のフレームを切り出しています。</div>
                    </div>
                    <span className="lecture-tool-progress__badge">実行中</span>
                  </div>
                </div>
              ) : result ? (
                <div className="video-thumb-result">
                  <div className="video-thumb-result__preview">
                    <img src={result.posterUrl} alt="生成したposter.png" />
                  </div>
                  <div className="video-thumb-result__downloads">
                    <div className="pdf-slide-result">
                      <div>
                        <div className="pdf-slide-result__title">poster.png</div>
                        <div className="pdf-slide-result__meta">1024 x 768 / {formatDuration(result.seconds)}</div>
                      </div>
                      <a className="lecture-tool-button lecture-tool-button--primary" href={result.posterUrl} download="poster.png">
                        ダウンロード
                      </a>
                    </div>
                    <div className="pdf-slide-result">
                      <div>
                        <div className="pdf-slide-result__title">thumb.png</div>
                        <div className="pdf-slide-result__meta">200 x 150 / {formatDuration(result.seconds)}</div>
                      </div>
                      <a className="lecture-tool-button lecture-tool-button--primary" href={result.thumbUrl} download="thumb.png">
                        ダウンロード
                      </a>
                    </div>
                  </div>
                </div>
              ) : (
                <div className="lecture-tool-empty">動画を追加して秒数を指定し、サムネイルを生成してください。</div>
              )}
            </div>
          </section>
        </div>
      </div>
    </div>
  );
}
