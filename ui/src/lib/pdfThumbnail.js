import { pdfjs } from "react-pdf";

pdfjs.GlobalWorkerOptions.workerSrc = new URL(
  "pdfjs-dist/build/pdf.worker.min.mjs",
  import.meta.url,
).toString();

export async function createPdfPageThumbnail(file, pageNumber = 1, maxWidth = 180) {
  const data = await file.arrayBuffer();
  const loadingTask = pdfjs.getDocument({ data });
  const pdf = await loadingTask.promise;

  try {
    const safePageNumber = Math.min(
      Math.max(Number.parseInt(pageNumber, 10) || 1, 1),
      pdf.numPages,
    );
    const page = await pdf.getPage(safePageNumber);
    const initialViewport = page.getViewport({ scale: 1 });
    const scale = maxWidth / initialViewport.width;
    const viewport = page.getViewport({ scale });
    const canvas = document.createElement("canvas");
    const context = canvas.getContext("2d", { alpha: false });

    canvas.width = Math.ceil(viewport.width);
    canvas.height = Math.ceil(viewport.height);
    await page.render({ canvasContext: context, viewport }).promise;
    return canvas.toDataURL("image/jpeg", 0.82);
  } finally {
    pdf.destroy();
  }
}

export async function createPdfFirstPageThumbnail(file, maxWidth = 180) {
  return createPdfPageThumbnail(file, 1, maxWidth);
}
