import init, { convertToPdf } from "../../pkg/office2pdf.js";

const fileInput = document.getElementById("fileInput");
const formatSelect = document.getElementById("formatSelect");
const convertButton = document.getElementById("convertButton");
const status = document.getElementById("status");
const viewer = document.getElementById("viewer");
const previewFrame = document.getElementById("previewFrame");
const downloadLink = document.getElementById("downloadLink");

const formatByExtension = new Map([
  ["docx", "docx"],
  ["pptx", "pptx"],
  ["xlsx", "xlsx"],
]);

let wasmInitPromise = null;
let pdfObjectUrl = null;
let wasmUrl = "";

function setStatus(message, isError = false) {
  status.textContent = message;
  status.style.color = isError ? "#b91c1c" : "#64748b";
}

function detectFormat(fileName) {
  const extension = fileName.split(".").pop()?.toLowerCase() ?? "";
  return formatByExtension.get(extension);
}

function getSelectedFormat(file) {
  if (formatSelect.value !== "auto") {
    return formatSelect.value;
  }
  return detectFormat(file.name) ?? null;
}

function getPdfFileName(inputFileName) {
  const baseName = inputFileName.replace(/\.[^.]+$/, "");
  const safeName = baseName.length > 0 ? baseName : "converted";
  return `${safeName}.pdf`;
}

async function ensureWasmReady() {
  if (wasmInitPromise === null) {
    if (window.location.protocol === "file:") {
      throw new Error(
        "Do not open this page with file://. Start a local server and use http://localhost (for example: python3 -m http.server).",
      );
    }

    wasmInitPromise = (async () => {
      const resolvedWasmUrl = new URL("../../pkg/office2pdf_bg.wasm", import.meta.url);
      wasmUrl = resolvedWasmUrl.href;
      setStatus(`Loading WASM module: ${resolvedWasmUrl.pathname}`);

      const response = await fetch(resolvedWasmUrl);
      if (!response.ok) {
        throw new Error(`Failed to fetch WASM (${response.status} ${response.statusText})`);
      }

      const wasmBytes = await response.arrayBuffer();
      await init({ module_or_path: wasmBytes });
      setStatus("WASM module loaded.");
    })();
  }
  await wasmInitPromise;
}

function updatePdfPreview(pdfBytes, sourceName) {
  if (pdfObjectUrl) {
    URL.revokeObjectURL(pdfObjectUrl);
  }

  const pdfBlob = new Blob([pdfBytes], { type: "application/pdf" });
  pdfObjectUrl = URL.createObjectURL(pdfBlob);

  previewFrame.src = pdfObjectUrl;
  downloadLink.href = pdfObjectUrl;
  downloadLink.download = getPdfFileName(sourceName);
  viewer.classList.add("visible");
}

async function handleConvertClick() {
  const file = fileInput.files?.[0];
  if (!file) {
    setStatus("Please select a DOCX, PPTX, or XLSX file first.", true);
    return;
  }

  const format = getSelectedFormat(file);
  if (format === null) {
    setStatus("Could not detect file format. Please choose it manually.", true);
    return;
  }

  convertButton.disabled = true;
  setStatus(`Converting ${file.name} as ${format.toUpperCase()}...`);

  try {
    await ensureWasmReady();
    const officeBytes = new Uint8Array(await file.arrayBuffer());
    const pdfBytes = convertToPdf(officeBytes, format);
    updatePdfPreview(pdfBytes, file.name);
    const wasmSourceHint = wasmUrl.length > 0 ? ` via ${wasmUrl}` : "";
    setStatus(`Done. Generated ${pdfBytes.length.toLocaleString()} bytes of PDF${wasmSourceHint}.`);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    setStatus(`Conversion failed: ${message}`, true);
  } finally {
    convertButton.disabled = false;
  }
}

convertButton.addEventListener("click", () => {
  void handleConvertClick();
});

window.addEventListener("beforeunload", () => {
  if (pdfObjectUrl) {
    URL.revokeObjectURL(pdfObjectUrl);
  }
});

void ensureWasmReady().catch((error) => {
  const message = error instanceof Error ? error.message : String(error);
  setStatus(`WASM preload failed: ${message}`, true);
});
