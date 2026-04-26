const QR_PAYLOAD_PREFIX = "GABMATH1:";
const STORAGE_KEY = "gabmath-current-proof";
const CARD_TARGET = {
  width: 820,
  height: 1160,
  leftMarkerX: 130,
  rightMarkerX: 690,
  topMarkerY: 180,
  bottomMarkerY: 1010,
};
const CARD_PROCESSING_MAX_WIDTH = 960;
const CARD_FRAME_INTERVAL_MS = 140;
const MIN_FOCUS_SCORE = 120;
const REQUIRED_STABLE_FRAMES = 4;
const MIN_CARD_RECTANGLE_SCORE = 0.12;
const MIN_MARKER_FILL_RATIO = 0.62;
const MIN_MARKER_SIDE = 16;
const MAX_MARKER_SIDE = 90;
const MIN_FILLED_BUBBLE_SCORE = 0.24;
const SECOND_BUBBLE_RELATIVE_LIMIT = 0.86;

// Leitura v2 (mais robusta para bolhas circulares, com limiar adaptativo).
const BUBBLE_CONTRAST_MIN = 0.085;
const BUBBLE_AMBIGUOUS_RELATIVE = 0.9;
const BUBBLE_AMBIGUOUS_GAP = 0.035;
const MARKER_CORNER_BONUS = 0.6;
const CARD_ASPECT = CARD_TARGET.width / CARD_TARGET.height;

// Marcadores (multiescala): aceitar quadrados pequenos (cartão longe) e grandes (cartão perto).
const MARKER_MIN_AREA_RATIO = 0.00002;
const MARKER_MAX_AREA_RATIO = 0.015;
const MARKER_MIN_SIDE_RATIO = 0.006;
const MARKER_MAX_SIDE_RATIO = 0.28;

const state = {
  stream: null,
  qrTimer: null,
  cardTimer: null,
  opencvReady: false,
  opencvKernel: null,
  proof: null,
  answers: {},
  stableSignature: "",
  stableCount: 0,
  lastDetection: null,
  cardMode: false,
  autoStartCardReading: false,
  debug: false,
  workflow: "idle",
  qrProcessing: false,
  lastQrValue: "",
  lastQrAt: 0,
  elements: {},
  page: "",
};

const captureCanvas = document.createElement("canvas");
const processingCanvas = document.createElement("canvas");

document.addEventListener("DOMContentLoaded", initializeApp);

function initializeApp() {
  state.page = document.body.dataset.page || "qr";
  if (state.page === "qr") {
    initializeQrPage();
    return;
  }
  if (state.page === "card") {
    initializeCardPage();
  }
}

function initializeQrPage() {
  state.elements = {
    modeBadge: document.getElementById("mode-badge"),
    video: document.getElementById("camera-video"),
    overlayCanvas: document.getElementById("camera-overlay"),
    cameraPanel: document.getElementById("camera-panel") || document.querySelector(".camera-stack"),
    startScanButton: document.getElementById("start-scan"),
    stopScanButton: document.getElementById("stop-scan"),
    loadProofButton: document.getElementById("load-proof"),
    proofIdInput: document.getElementById("proof-id-input"),
    scanStatus: document.getElementById("scan-status"),
    cardStatus: document.getElementById("card-status"),
    qrStep: document.getElementById("step-qr"),
    cardStep: document.getElementById("step-card"),
    enableCameraButton: document.getElementById("enable-camera"),
    capturePhotoButton: document.getElementById("capture-photo"),
    retakePhotoButton: document.getElementById("retake-photo"),
    confirmReadingButton: document.getElementById("confirm-reading"),
    cancelReadingButton: document.getElementById("cancel-reading"),
    disableCameraButton: document.getElementById("disable-camera"),
    proofSummary: document.getElementById("proof-summary"),
    answerGrid: document.getElementById("answer-grid"),
    resultPanel: document.getElementById("result-panel"),
    alignedCanvas: document.getElementById("aligned-preview"),
    backToQrButton: document.getElementById("back-to-qr"),
    debugToggle: document.getElementById("debug-toggle"),
  };

  state.elements.startScanButton.addEventListener("click", startQrCamera);
  state.elements.stopScanButton.addEventListener("click", stopWorkflow);
  state.elements.loadProofButton.addEventListener("click", () => commitQrAndGo(state.elements.proofIdInput.value));
  wireCardButtons();
  if (state.elements.backToQrButton) {
    state.elements.backToQrButton.addEventListener("click", backToQrMode);
  }

  waitForOpenCv();
  setStatus("Aponte a camera apenas para o QR Code.");
  state.debug = Boolean(getQueryFlag("debug")) || Boolean(state.elements.debugToggle?.checked);
  if (state.elements.debugToggle) {
    state.elements.debugToggle.addEventListener("change", () => {
      state.debug = Boolean(state.elements.debugToggle.checked);
    });
  }

  const stored = sessionStorage.getItem(STORAGE_KEY);
  if (stored) {
    try {
      const proof = JSON.parse(stored);
      if (proof && state.elements.cardStep) {
        enterCardMode(proof, { autoStart: false });
      }
    } catch {
      // ignore
    }
  }
}

function initializeCardPage() {
  state.elements = {
    modeBadge: document.getElementById("mode-badge"),
    video: document.getElementById("camera-video"),
    overlayCanvas: document.getElementById("camera-overlay"),
    cameraPanel: document.getElementById("camera-panel") || document.querySelector(".camera-stack"),
    alignedCanvas: document.getElementById("aligned-preview"),
    startScanButton: document.getElementById("start-scan"),
    stopScanButton: document.getElementById("stop-scan"),
    startCardScanButton: document.getElementById("start-card-scan"),
    enableCameraButton: document.getElementById("enable-camera"),
    capturePhotoButton: document.getElementById("capture-photo"),
    retakePhotoButton: document.getElementById("retake-photo"),
    confirmReadingButton: document.getElementById("confirm-reading"),
    cancelReadingButton: document.getElementById("cancel-reading"),
    disableCameraButton: document.getElementById("disable-camera"),
    proofSummary: document.getElementById("proof-summary"),
    scanStatus: document.getElementById("scan-status"),
    answerGrid: document.getElementById("answer-grid"),
    resultPanel: document.getElementById("result-panel"),
    debugToggle: document.getElementById("debug-toggle"),
  };

  wireCardButtons();

  waitForOpenCv();
  state.debug = Boolean(getQueryFlag("debug")) || Boolean(state.elements.debugToggle?.checked);
  if (state.elements.debugToggle) {
    state.elements.debugToggle.addEventListener("change", () => {
      state.debug = Boolean(state.elements.debugToggle.checked);
    });
  }
  const stored = sessionStorage.getItem(STORAGE_KEY);
  if (!stored) {
    setStatus("Nenhuma prova foi carregada. Volte e leia o QR Code primeiro.");
    state.elements.startCardScanButton.disabled = true;
    return;
  }

  state.proof = JSON.parse(stored);
  renderProofSummary();
  renderAnswerGrid();
  setWorkflowState("waitingUserToStartAnswerScan");
  updateCardControls();
  setStatus("Alinhe o celular com o cartão-resposta e toque em “Habilitar câmera”.");
}

async function startQrCamera() {
  setWorkflowState("qrScanning", "Abrindo câmera…");
  await startCamera();
  if (!state.stream) {
    setWorkflowState("error", "Falha ao abrir a câmera.");
    return;
  }
  setCameraPanelVisible(true);
  setStatus("Câmera aberta. Aponte apenas para o QR Code.");
  startQrLoop();
}

async function startCardCamera() {
  await startCamera();
  if (!state.stream) {
    setWorkflowState("error", "Falha ao abrir a câmera.");
    return;
  }
  setCameraPanelVisible(true);
  setWorkflowState("waitingUserToStartAnswerScan", "Câmera aberta. Alinhe o cartão e toque em “Iniciar leitura”.");
  updateCardControls();
}

async function startCamera() {
  if (!navigator.mediaDevices?.getUserMedia) {
    setStatus("Camera nao disponivel neste navegador.");
    return;
  }

  stopLoops();
  clearOverlay();
  disposeCameraResources();

  try {
    const stream = await navigator.mediaDevices.getUserMedia({
      video: {
        facingMode: { ideal: "environment" },
        width: { ideal: 1920 },
        height: { ideal: 1080 },
      },
      audio: false,
    });
    state.stream = stream;
    state.elements.video.srcObject = stream;
    await state.elements.video.play();

    // Tenta habilitar foco contínuo (quando suportado).
    try {
      const track = stream.getVideoTracks?.()[0];
      if (track?.applyConstraints) {
        await track.applyConstraints({ advanced: [{ focusMode: "continuous" }] });
      }
    } catch {
      // ignore
    }
  } catch (error) {
    setStatus(`Falha ao abrir a camera: ${error?.message || String(error)}`);
  }
}

function stopWorkflow() {
  if (state.workflow === "qrScanning") {
    stopQrCamera();
    setWorkflowState("idle", "Leitura do QR interrompida.");
    return;
  }
  if (state.workflow === "answerScanning") {
    stopAnswerCamera();
    setWorkflowState("waitingUserToStartAnswerScan", "Leitura do cartão interrompida.");
    return;
  }
  disposeCameraResources();
  stopLoops();
  clearOverlay();
  setWorkflowState("idle", "Leitura interrompida.");
}

function disposeCameraResources() {
  if (state.stream) {
    for (const track of state.stream.getTracks()) {
      try {
        track.stop();
      } catch {
        // ignore
      }
    }
    state.stream = null;
  }
  if (state.elements.video) {
    try {
      state.elements.video.pause?.();
    } catch {
      // ignore
    }
    state.elements.video.srcObject = null;
  }
}

function stopQrCamera() {
  stopLoops();
  state.qrProcessing = false;
  disposeCameraResources();
  clearOverlay();
  setCameraPanelVisible(false);
}

function stopAnswerCamera() {
  stopLoops();
  disposeCameraResources();
  clearOverlay();
  setCameraPanelVisible(false);
}

function stopLoops() {
  if (state.qrTimer) {
    clearInterval(state.qrTimer);
    state.qrTimer = null;
  }
  if (state.cardTimer) {
    clearInterval(state.cardTimer);
    cancelAnimationFrame(state.cardTimer);
    state.cardTimer = null;
  }
}

function startQrLoop() {
  stopLoops();
  state.qrProcessing = false;
  state.qrTimer = setInterval(async () => {
    if (!state.elements.video?.srcObject) {
      return;
    }
    drawQrGuide();
    if (state.qrProcessing) {
      return;
    }
    try {
      const rawValue = await detectQrCode(state.elements.video);
      if (!rawValue) {
        return;
      }
      const now = Date.now();
      if (rawValue === state.lastQrValue && (now - state.lastQrAt) < 2000) {
        return;
      }
      state.lastQrValue = rawValue;
      state.lastQrAt = now;
      state.qrProcessing = true;
      state.elements.proofIdInput.value = rawValue;
      commitQrAndGo(rawValue);
    } catch (error) {
      setStatus(`Falha na leitura do QR: ${error?.message || String(error)}`);
      state.qrProcessing = false;
    }
  }, 450);
}

function enterCardMode(proof, { autoStart } = { autoStart: false }) {
  state.cardMode = true;
  state.proof = proof;
  state.answers = {};
  state.stableSignature = "";
  state.stableCount = 0;
  state.lastDetection = null;
  state.autoStartCardReading = Boolean(autoStart);

  // Ao sair do QR, sempre libera a câmera e espera o usuário iniciar.
  stopQrCamera();
  setCameraPanelVisible(false);

  if (state.elements.qrStep) {
    state.elements.qrStep.classList.add("hidden");
  }
  if (state.elements.cardStep) {
    state.elements.cardStep.classList.remove("hidden");
  }
  if (state.elements.resultPanel) {
    state.elements.resultPanel.classList.add("hidden");
    state.elements.resultPanel.innerHTML = "";
  }

  renderProofSummary();
  renderAnswerGrid();
  setWorkflowState("waitingUserToStartAnswerScan");
  setBadge("Cartão-resposta");
  updateCardControls();

  if (!state.opencvReady) {
    setStatus("Carregando OpenCV... aguarde alguns segundos.");
    return;
  }

  setStatus("Alinhe o celular com o cartão-resposta e toque em “Iniciar leitura”.");
}

function backToQrMode() {
  state.cardMode = false;
  state.autoStartCardReading = false;
  stopLoops();
  stopAnswerCamera();
  setCameraPanelVisible(false);
  state.proof = null;
  state.answers = {};
  state.stableSignature = "";
  state.stableCount = 0;
  state.lastDetection = null;
  sessionStorage.removeItem(STORAGE_KEY);

  if (state.elements.cardStep) {
    state.elements.cardStep.classList.add("hidden");
  }
  if (state.elements.qrStep) {
    state.elements.qrStep.classList.remove("hidden");
  }
  setBadge("Lendo QR");
  setWorkflowState("idle", "Aponte a câmera apenas para o QR Code.");
  updateCardControls();
}

async function detectQrCode(videoElement) {
  if ("BarcodeDetector" in window) {
    const detector = new BarcodeDetector({ formats: ["qr_code"] });
    const codes = await detector.detect(videoElement);
    if (codes.length) {
      return String(codes[0].rawValue || "").trim();
    }
  }

  if (typeof window.jsQR === "function") {
    const videoWidth = videoElement.videoWidth || 0;
    const videoHeight = videoElement.videoHeight || 0;
    if (videoWidth < 40 || videoHeight < 40) {
      return "";
    }

    // Processa apenas a área central (melhora performance e reduz falso positivo).
    const cropScale = 0.68;
    const cropWidth = Math.round(videoWidth * cropScale);
    const cropHeight = Math.round(videoHeight * cropScale);
    const cropX = Math.round((videoWidth - cropWidth) / 2);
    const cropY = Math.round((videoHeight - cropHeight) / 2);

    // Downscale para acelerar (jsQR é CPU-bound).
    const maxSide = 720;
    const scale = Math.min(1, maxSide / Math.max(cropWidth, cropHeight, 1));
    captureCanvas.width = Math.max(1, Math.round(cropWidth * scale));
    captureCanvas.height = Math.max(1, Math.round(cropHeight * scale));
    const context = captureCanvas.getContext("2d", { willReadFrequently: true });
    context.drawImage(
      videoElement,
      cropX,
      cropY,
      cropWidth,
      cropHeight,
      0,
      0,
      captureCanvas.width,
      captureCanvas.height,
    );
    const imageData = context.getImageData(0, 0, captureCanvas.width, captureCanvas.height);
    const code = window.jsQR(imageData.data, imageData.width, imageData.height, {
      inversionAttempts: "attemptBoth",
    });
    if (code?.data) {
      return String(code.data).trim();
    }
  }

  return "";
}

function drawQrGuide() {
  if (!state.elements?.overlayCanvas || !state.elements?.video) {
    return;
  }
  const width = state.elements.video.clientWidth || state.elements.video.videoWidth || 1;
  const height = state.elements.video.clientHeight || state.elements.video.videoHeight || 1;
  state.elements.overlayCanvas.width = width;
  state.elements.overlayCanvas.height = height;
  const context = state.elements.overlayCanvas.getContext("2d");
  context.clearRect(0, 0, width, height);

  const boxW = Math.round(width * 0.62);
  const boxH = Math.round(height * 0.62);
  const x = Math.round((width - boxW) / 2);
  const y = Math.round((height - boxH) / 2);

  context.strokeStyle = "rgba(255,255,255,0.85)";
  context.lineWidth = 3;
  context.setLineDash([10, 8]);
  context.strokeRect(x, y, boxW, boxH);
  context.setLineDash([]);
  context.fillStyle = "rgba(0,0,0,0.15)";
  context.fillRect(0, 0, width, y);
  context.fillRect(0, y + boxH, width, height - (y + boxH));
  context.fillRect(0, y, x, boxH);
  context.fillRect(x + boxW, y, width - (x + boxW), boxH);
}

function setCameraPanelVisible(visible) {
  const panel = state.elements?.cameraPanel;
  if (!panel) {
    return;
  }
  panel.classList.toggle("hidden", !visible);
}

function drawAnswerGuide() {
  if (!state.elements?.overlayCanvas || !state.elements?.video) {
    return;
  }
  const width = state.elements.video.clientWidth || state.elements.video.videoWidth || 1;
  const height = state.elements.video.clientHeight || state.elements.video.videoHeight || 1;
  state.elements.overlayCanvas.width = width;
  state.elements.overlayCanvas.height = height;
  const context = state.elements.overlayCanvas.getContext("2d");
  context.clearRect(0, 0, width, height);

  // Proporção aproximada do cartão (CARD_TARGET.width / CARD_TARGET.height)
  const targetAspect = CARD_TARGET.width / CARD_TARGET.height;
  let boxW = Math.round(width * 0.92);
  let boxH = Math.round(boxW / Math.max(targetAspect, 0.01));
  if (boxH > Math.round(height * 0.78)) {
    boxH = Math.round(height * 0.78);
    boxW = Math.round(boxH * targetAspect);
  }
  const x = Math.round((width - boxW) / 2);
  const y = Math.round((height - boxH) / 2);

  context.strokeStyle = "rgba(52,199,89,0.92)";
  context.lineWidth = 4;
  context.setLineDash([12, 10]);
  context.strokeRect(x, y, boxW, boxH);
  context.setLineDash([]);

  // Marcação dos 4 cantos (onde os quadrados pretos devem ficar).
  const cornerSize = Math.max(18, Math.round(Math.min(boxW, boxH) * 0.08));
  const inset = Math.max(10, Math.round(cornerSize * 0.25));
  context.lineWidth = 5;
  context.strokeStyle = "rgba(52,199,89,0.95)";

  drawCorner(context, x + inset, y + inset, cornerSize, "tl");
  drawCorner(context, x + boxW - inset, y + inset, cornerSize, "tr");
  drawCorner(context, x + boxW - inset, y + boxH - inset, cornerSize, "br");
  drawCorner(context, x + inset, y + boxH - inset, cornerSize, "bl");
}

function drawCorner(context, x, y, size, where) {
  const half = Math.round(size / 2);
  const len = Math.round(size * 0.7);
  context.beginPath();
  if (where === "tl") {
    context.moveTo(x, y + len);
    context.lineTo(x, y);
    context.lineTo(x + len, y);
  } else if (where === "tr") {
    context.moveTo(x - len, y);
    context.lineTo(x, y);
    context.lineTo(x, y + len);
  } else if (where === "br") {
    context.moveTo(x, y - len);
    context.lineTo(x, y);
    context.lineTo(x - len, y);
  } else {
    context.moveTo(x + len, y);
    context.lineTo(x, y);
    context.lineTo(x, y - len);
  }
  context.stroke();
}

function commitQrAndGo(rawValue) {
  const normalized = String(rawValue || "").trim();
  if (!normalized) {
    setStatus("Informe ou leia um QR valido.");
    return;
  }

  try {
    const proof = parseQrPayload(normalized);
    sessionStorage.setItem(STORAGE_KEY, JSON.stringify(proof));
    if (state.elements.cardStep) {
      enterCardMode(proof, { autoStart: false });
    } else {
      window.location.href = "./card.html";
    }
    state.qrProcessing = false;
  } catch (error) {
    setStatus(error?.message || String(error));
    state.qrProcessing = false;
  }
}

function waitForOpenCv() {
  if (window.cv && typeof window.cv.Mat === "function") {
    state.opencvReady = true;
    state.opencvKernel = cv.getStructuringElement(cv.MORPH_RECT, new cv.Size(3, 3));
    onOpenCvReady();
    return;
  }

  const timer = setInterval(() => {
    if (window.cv && typeof window.cv.Mat === "function") {
      clearInterval(timer);
      state.opencvReady = true;
      state.opencvKernel = cv.getStructuringElement(cv.MORPH_RECT, new cv.Size(3, 3));
      if (state.page === "card") {
        setStatus("OpenCV carregado. Agora voce pode iniciar a leitura do cartao.");
      }
      onOpenCvReady();
    }
  }, 250);
}

function onOpenCvReady() {
  updateCardControls();
}

function renderProofSummary() {
  if (!state.proof) {
    return;
  }
  state.elements.proofSummary.innerHTML = `
    <div><strong>Prova:</strong> ${escapeHtml(state.proof.id_prova || "")}</div>
    <div><strong>Aluno:</strong> ${escapeHtml(state.proof.aluno || "")}</div>
    <div><strong>Questoes:</strong> ${Number(state.proof.quantidade_questoes || 0)}</div>
  `;
}

function renderAnswerGrid() {
  state.elements.answerGrid.innerHTML = "";
  for (const question of state.proof.questoes || []) {
    const row = document.createElement("div");
    row.className = "answer-row";
    row.dataset.question = String(question.numero);

    const label = document.createElement("div");
    label.className = "answer-label";
    label.textContent = `Q.${question.numero}`;
    row.appendChild(label);

    for (const letter of ["A", "B", "C", "D", "E"]) {
      const box = document.createElement("div");
      box.className = "choice-pill";
      box.dataset.letter = letter;
      box.textContent = letter;
      row.appendChild(box);
    }

    state.elements.answerGrid.appendChild(row);
  }
}

function updateRenderedAnswers() {
  for (const row of state.elements.answerGrid.querySelectorAll(".answer-row")) {
    const questionNumber = row.dataset.question;
    const selected = state.answers[String(questionNumber)] || "";
    for (const pill of row.querySelectorAll(".choice-pill")) {
      pill.classList.toggle("selected", pill.dataset.letter === selected);
    }
  }
}

function startCardReading() {
  // Fluxo antigo (tempo real) foi substituído por captura de foto.
  setStatus("Use o botão “Capturar foto” para fazer a leitura.");
}

function stopAnswerReading() {
  // No modo por foto, não há leitura contínua. Mantém por compatibilidade.
  stopLoops();
  clearOverlay();
  setWorkflowState("waitingUserToStartAnswerScan", "Leitura pausada.");
  updateCardControls();
}

function disableAnswerCamera() {
  stopAnswerReading();
  stopAnswerCamera();
  setWorkflowState("waitingUserToStartAnswerScan", "Câmera desabilitada. Toque em “Habilitar câmera”.");
  updateCardControls();
}

async function enableAnswerCamera() {
  if (state.stream) {
    setCameraPanelVisible(true);
    setWorkflowState("waitingUserToStartAnswerScan");
    updateCardControls();
    return;
  }
  await startCardCamera();
}

function wireCardButtons() {
  if (!state.elements) {
    return;
  }

  // Novo layout (index.html atualizado)
  if (state.elements.enableCameraButton && !state.elements.enableCameraButton.dataset.wired) {
    state.elements.enableCameraButton.dataset.wired = "1";
    state.elements.enableCameraButton.addEventListener("click", enableAnswerCamera);
  }
  if (state.elements.capturePhotoButton && !state.elements.capturePhotoButton.dataset.wired) {
    state.elements.capturePhotoButton.dataset.wired = "1";
    state.elements.capturePhotoButton.addEventListener("click", captureAndProcessPhoto);
  }
  if (state.elements.retakePhotoButton && !state.elements.retakePhotoButton.dataset.wired) {
    state.elements.retakePhotoButton.dataset.wired = "1";
    state.elements.retakePhotoButton.addEventListener("click", retakePhoto);
  }
  if (state.elements.confirmReadingButton && !state.elements.confirmReadingButton.dataset.wired) {
    state.elements.confirmReadingButton.dataset.wired = "1";
    state.elements.confirmReadingButton.addEventListener("click", confirmReading);
  }
  if (state.elements.cancelReadingButton && !state.elements.cancelReadingButton.dataset.wired) {
    state.elements.cancelReadingButton.dataset.wired = "1";
    state.elements.cancelReadingButton.addEventListener("click", cancelReading);
  }
  if (state.elements.disableCameraButton && !state.elements.disableCameraButton.dataset.wired) {
    state.elements.disableCameraButton.dataset.wired = "1";
    state.elements.disableCameraButton.addEventListener("click", disableAnswerCamera);
  }

  // Compatibilidade com card.html antigo (se existir).
  if (state.elements.startScanButton && state.elements.stopScanButton && state.elements.startCardScanButton) {
    if (!state.elements.startScanButton.dataset.wired) {
      state.elements.startScanButton.dataset.wired = "1";
      state.elements.startScanButton.addEventListener("click", startCardCamera);
    }
    if (!state.elements.stopScanButton.dataset.wired) {
      state.elements.stopScanButton.dataset.wired = "1";
      state.elements.stopScanButton.addEventListener("click", stopWorkflow);
    }
    if (!state.elements.startCardScanButton.dataset.wired) {
      state.elements.startCardScanButton.dataset.wired = "1";
      state.elements.startCardScanButton.addEventListener("click", captureAndProcessPhoto);
    }
  }
}

function detectCardAnswers(videoElement, questionCount) {
  if (videoElement.videoWidth < 100 || videoElement.videoHeight < 100) {
    return null;
  }

  const scale = Math.min(1, CARD_PROCESSING_MAX_WIDTH / Math.max(videoElement.videoWidth, 1));
  processingCanvas.width = Math.max(1, Math.round(videoElement.videoWidth * scale));
  processingCanvas.height = Math.max(1, Math.round(videoElement.videoHeight * scale));
  const captureContext = processingCanvas.getContext("2d", { willReadFrequently: true });
  captureContext.drawImage(videoElement, 0, 0, processingCanvas.width, processingCanvas.height);

  const src = cv.imread(processingCanvas);
  const gray = new cv.Mat();
  const blur = new cv.Mat();
  const thresh = new cv.Mat();

  try {
    cv.cvtColor(src, gray, cv.COLOR_RGBA2GRAY);
    cv.equalizeHist(gray, gray);
    cv.GaussianBlur(gray, blur, new cv.Size(5, 5), 0, 0, cv.BORDER_DEFAULT);
    cv.threshold(blur, thresh, 0, 255, cv.THRESH_BINARY_INV + cv.THRESH_OTSU);
    if (state.opencvKernel) {
      cv.morphologyEx(thresh, thresh, cv.MORPH_CLOSE, state.opencvKernel);
    }
    let markers = findMarkersV2(thresh, processingCanvas.width, processingCanvas.height);
    if (markers.length < 4) {
      // Fallback para o algoritmo anterior.
      markers = findMarkers(thresh, processingCanvas.width, processingCanvas.height);
    }

    if (markers.length < 4) {
      // Fallback: adaptativo costuma funcionar melhor com sombras/reflexo.
      cv.adaptiveThreshold(
        blur,
        thresh,
        255,
        cv.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv.THRESH_BINARY_INV,
        31,
        7,
      );
      if (state.opencvKernel) {
        cv.morphologyEx(thresh, thresh, cv.MORPH_CLOSE, state.opencvKernel);
      }
      markers = findMarkersV2(thresh, processingCanvas.width, processingCanvas.height);
      if (markers.length < 4) {
        markers = findMarkers(thresh, processingCanvas.width, processingCanvas.height);
      }
      if (markers.length < 4) {
        return null;
      }
    }

    const corners = orderCorners(markers.slice(0, 4).map((marker) => marker.center));
    if (rectangleScore(corners, processingCanvas.width, processingCanvas.height) < MIN_CARD_RECTANGLE_SCORE) {
      return null;
    }
    const warped = perspectiveWarp(gray, corners);
    const focusScore = measureFocus(warped);
    const basePreview = new cv.Mat();
    cv.cvtColor(warped, basePreview, cv.COLOR_GRAY2RGBA);

    let reading;
    try {
      reading = readAnswersFromWarpedV2(warped, questionCount);
    } catch {
      reading = readAnswersFromWarped(warped, questionCount);
    }
    const baseImageData = matToImageData(basePreview);

    warped.delete();
    basePreview.delete();

    if (state.debug) {
      drawGridOverlay(baseImageData, reading.rows);
    }

    return {
      corners,
      answers: reading.answers,
      rows: reading.rows,
      baseImageData,
      focusScore,
    };
  } finally {
    src.delete();
    gray.delete();
    blur.delete();
    thresh.delete();
  }
}

function detectCardAnswersFromCanvas(canvasElement, questionCount) {
  if (!canvasElement?.width || !canvasElement?.height) {
    return null;
  }

  const src = cv.imread(canvasElement);
  const gray = new cv.Mat();
  const blur = new cv.Mat();
  const thresh = new cv.Mat();

  try {
    cv.cvtColor(src, gray, cv.COLOR_RGBA2GRAY);
    cv.equalizeHist(gray, gray);
    // Multiescala: procura marcadores em vários tamanhos para tolerar distância.
    let markers = findMarkersMultiScale(gray);

    // Fallback: pipeline antigo por threshold único.
    if (markers.length < 4) {
      cv.GaussianBlur(gray, blur, new cv.Size(5, 5), 0, 0, cv.BORDER_DEFAULT);
      cv.threshold(blur, thresh, 0, 255, cv.THRESH_BINARY_INV + cv.THRESH_OTSU);
      if (state.opencvKernel) {
        cv.morphologyEx(thresh, thresh, cv.MORPH_CLOSE, state.opencvKernel);
      }
      markers = findMarkersV2(thresh, canvasElement.width, canvasElement.height);
      if (markers.length < 4) {
        cv.adaptiveThreshold(
          blur,
          thresh,
          255,
          cv.ADAPTIVE_THRESH_GAUSSIAN_C,
          cv.THRESH_BINARY_INV,
          31,
          7,
        );
        if (state.opencvKernel) {
          cv.morphologyEx(thresh, thresh, cv.MORPH_CLOSE, state.opencvKernel);
        }
        markers = findMarkersV2(thresh, canvasElement.width, canvasElement.height);
      }
      if (markers.length < 4) {
        return { ok: false, reason: "markers_not_found" };
      }
    }

    if (state.debug) {
      try {
        const preview = new cv.Mat();
        cv.cvtColor(gray, preview, cv.COLOR_GRAY2RGBA);
        const imageData = matToImageData(preview);
        preview.delete();
        drawMarkerSelectionDebug(imageData, markers.slice(0, 4));
      } catch {
        // ignore
      }
    }

    const corners = orderCorners(markers.slice(0, 4).map((marker) => marker.center));
    if (rectangleScore(corners, canvasElement.width, canvasElement.height) < MIN_CARD_RECTANGLE_SCORE) {
      return { ok: false, reason: "bad_geometry" };
    }

    const warped = perspectiveWarp(gray, corners);
    const focusScore = measureFocus(warped);
    if (focusScore < MIN_FOCUS_SCORE * 0.65) {
      warped.delete();
      return { ok: false, reason: "blur" };
    }

    const basePreview = new cv.Mat();
    cv.cvtColor(warped, basePreview, cv.COLOR_GRAY2RGBA);

    let reading;
    try {
      reading = readAnswersFromWarpedV2(warped, questionCount);
    } catch {
      reading = readAnswersFromWarped(warped, questionCount);
    }

    const baseImageData = matToImageData(basePreview);
    warped.delete();
    basePreview.delete();

    if (state.debug) {
      drawGridOverlay(baseImageData, reading.rows);
    }

    return {
      ok: true,
      corners,
      answers: reading.answers,
      rows: reading.rows,
      baseImageData,
      focusScore,
    };
  } finally {
    src.delete();
    gray.delete();
    blur.delete();
    thresh.delete();
  }
}

function drawMarkerSelectionDebug(imageData, markers) {
  if (!imageData || !state.elements?.alignedCanvas) {
    return;
  }
  // Reutiliza o canvas de preview alinhado para exibir a foto com marcações.
  state.elements.alignedCanvas.width = imageData.width;
  state.elements.alignedCanvas.height = imageData.height;
  const context = state.elements.alignedCanvas.getContext("2d");
  context.putImageData(imageData, 0, 0);
  if (!markers?.length) {
    return;
  }
  context.lineWidth = 4;
  context.strokeStyle = "#ffb300";
  for (const marker of markers) {
    const x = marker.center.x;
    const y = marker.center.y;
    context.beginPath();
    context.arc(x, y, 16, 0, Math.PI * 2);
    context.stroke();
  }
}

function findMarkers(binaryMat, width, height) {
  const contours = new cv.MatVector();
  const hierarchy = new cv.Mat();
  try {
    cv.findContours(binaryMat, contours, hierarchy, cv.RETR_EXTERNAL, cv.CHAIN_APPROX_SIMPLE);
    return collectMarkerCandidates(contours, width, height);
  } finally {
    contours.delete();
    hierarchy.delete();
  }
}

function findMarkersV2(binaryMat, width, height) {
  const contours = new cv.MatVector();
  const hierarchy = new cv.Mat();
  try {
    cv.findContours(binaryMat, contours, hierarchy, cv.RETR_EXTERNAL, cv.CHAIN_APPROX_SIMPLE);
    return collectMarkerCandidatesV3(contours, width, height);
  } finally {
    contours.delete();
    hierarchy.delete();
  }
}

function findMarkersMultiScale(grayMat) {
  const scales = [1.0, 0.8, 0.62, 0.48];
  const candidates = [];

  for (const scale of scales) {
    const scaled = new cv.Mat();
    const blur = new cv.Mat();
    const thresh = new cv.Mat();
    try {
      if (scale === 1.0) {
        grayMat.copyTo(scaled);
      } else {
        const size = new cv.Size(
          Math.max(1, Math.round(grayMat.cols * scale)),
          Math.max(1, Math.round(grayMat.rows * scale)),
        );
        cv.resize(grayMat, scaled, size, 0, 0, cv.INTER_AREA);
      }

      cv.GaussianBlur(scaled, blur, new cv.Size(5, 5), 0, 0, cv.BORDER_DEFAULT);
      cv.threshold(blur, thresh, 0, 255, cv.THRESH_BINARY_INV + cv.THRESH_OTSU);
      if (state.opencvKernel) {
        cv.morphologyEx(thresh, thresh, cv.MORPH_CLOSE, state.opencvKernel);
      }

      let markers = findMarkersV2(thresh, scaled.cols, scaled.rows);
      if (markers.length < 4) {
        markers = findMarkers(thresh, scaled.cols, scaled.rows);
      }

      for (const marker of markers) {
        candidates.push({
          ...marker,
          center: { x: marker.center.x / scale, y: marker.center.y / scale },
          area: (marker.area || 0) / (scale * scale),
          _scale: scale,
        });
      }
    } finally {
      scaled.delete();
      blur.delete();
      thresh.delete();
    }
  }

  if (candidates.length < 4) {
    return [];
  }

  // Dedupe no espaço original.
  candidates.sort((a, b) => ((b.squareScore || 0) - (a.squareScore || 0)) || ((b.cornerScore || 0) - (a.cornerScore || 0)));
  const unique = [];
  for (const candidate of candidates) {
    const duplicate = unique.some((item) => distance(item.center, candidate.center) < 22);
    if (!duplicate) {
      unique.push(candidate);
    }
    if (unique.length >= 22) {
      break;
    }
  }

  if (unique.length < 4) {
    return [];
  }

  let best = null;
  const limit = Math.min(unique.length, 14);
  for (let a = 0; a < limit - 3; a += 1) {
    for (let b = a + 1; b < limit - 2; b += 1) {
      for (let c = b + 1; c < limit - 1; c += 1) {
        for (let d = c + 1; d < limit; d += 1) {
          const combo = [unique[a], unique[b], unique[c], unique[d]];
          const ordered = orderCorners(combo.map((item) => item.center));
          const rectScore = rectangleScore(ordered, grayMat.cols, grayMat.rows);
          const cornerBoost =
            (combo.reduce((sum, item) => sum + (item.cornerScore || 0), 0) / 4) * MARKER_CORNER_BONUS;
          const squareBoost = combo.reduce((sum, item) => sum + (item.squareScore || 0), 0) / 4;
          const score = rectScore * (1 + cornerBoost) * (0.55 + (squareBoost * 0.45));
          if (!best || score > best.score) {
            best = { score, markers: combo };
          }
        }
      }
    }
  }

  return best ? best.markers : unique.slice(0, 4);
}

function collectMarkerCandidates(contours, width, height) {
  const candidates = [];
  const minArea = (width * height) * 0.00012;
  const minSide = Math.max(MIN_MARKER_SIDE, Math.round(Math.min(width, height) * 0.018));
  const maxSide = Math.max(MAX_MARKER_SIDE, Math.round(Math.min(width, height) * 0.14));

  for (let index = 0; index < contours.size(); index += 1) {
    const contour = contours.get(index);
    const area = cv.contourArea(contour);
    if (area < minArea) {
      contour.delete();
      continue;
    }

    const perimeter = cv.arcLength(contour, true);
    const approx = new cv.Mat();
    cv.approxPolyDP(contour, approx, perimeter * 0.02, true);
    const rect = cv.boundingRect(contour);
    const aspect = rect.width / Math.max(rect.height, 1);
    const fillRatio = area / Math.max(rect.width * rect.height, 1);
    const convex = cv.isContourConvex(approx);
    const hull = new cv.Mat();
    cv.convexHull(contour, hull, false, true);
    const hullArea = cv.contourArea(hull);
    const solidity = area / Math.max(hullArea, 1);

    if (
      approx.rows === 4 &&
      convex &&
      aspect > 0.8 &&
      aspect < 1.2 &&
      fillRatio > MIN_MARKER_FILL_RATIO &&
      solidity > 0.86 &&
      rect.width >= minSide &&
      rect.height >= minSide &&
      rect.width <= maxSide &&
      rect.height <= maxSide
    ) {
      const moments = cv.moments(contour);
      if (moments.m00 !== 0) {
        candidates.push({
          area,
          center: {
            x: moments.m10 / moments.m00,
            y: moments.m01 / moments.m00,
          },
        });
      }
    }

    hull.delete();
    approx.delete();
    contour.delete();
  }

  candidates.sort((left, right) => right.area - left.area);
  const unique = [];
  for (const candidate of candidates) {
    const duplicate = unique.some((item) => distance(item.center, candidate.center) < 20);
    if (!duplicate) {
      unique.push(candidate);
    }
    if (unique.length >= 10) {
      break;
    }
  }

  if (unique.length < 4) {
    return [];
  }

  let best = null;
  const limit = Math.min(unique.length, 8);
  for (let a = 0; a < limit - 3; a += 1) {
    for (let b = a + 1; b < limit - 2; b += 1) {
      for (let c = b + 1; c < limit - 1; c += 1) {
        for (let d = c + 1; d < limit; d += 1) {
          const combo = [unique[a], unique[b], unique[c], unique[d]];
          const ordered = orderCorners(combo.map((item) => item.center));
          const score = rectangleScore(ordered, width, height);
          if (!best || score > best.score) {
            best = { score, markers: combo };
          }
        }
      }
    }
  }

  return best ? best.markers : unique.slice(0, 4);
}

function cornerProximityScore(center, width, height) {
  const corners = [
    { x: 0, y: 0 },
    { x: width, y: 0 },
    { x: width, y: height },
    { x: 0, y: height },
  ];
  const diag = Math.hypot(width, height) || 1;
  let best = 0;
  for (const corner of corners) {
    const d = Math.hypot(center.x - corner.x, center.y - corner.y);
    const score = 1 - Math.min(1, d / (diag * 0.52));
    if (score > best) {
      best = score;
    }
  }
  return best;
}

function collectMarkerCandidatesV2(contours, width, height) {
  const candidates = [];
  const minArea = (width * height) * 0.000095;
  const minSide = Math.max(MIN_MARKER_SIDE, Math.round(Math.min(width, height) * 0.016));
  const maxSide = Math.max(MAX_MARKER_SIDE, Math.round(Math.min(width, height) * 0.16));

  for (let index = 0; index < contours.size(); index += 1) {
    const contour = contours.get(index);
    const area = cv.contourArea(contour);
    if (area < minArea) {
      contour.delete();
      continue;
    }

    const rect = cv.boundingRect(contour);
    const aspect = rect.width / Math.max(rect.height, 1);
    const fillRatio = area / Math.max(rect.width * rect.height, 1);

    if (
      aspect < 0.72 ||
      aspect > 1.38 ||
      rect.width < minSide ||
      rect.height < minSide ||
      rect.width > maxSide ||
      rect.height > maxSide ||
      fillRatio < (MIN_MARKER_FILL_RATIO - 0.1)
    ) {
      contour.delete();
      continue;
    }

    const hull = new cv.Mat();
    cv.convexHull(contour, hull, false, true);
    const hullArea = cv.contourArea(hull);
    const solidity = area / Math.max(hullArea, 1);
    hull.delete();

    if (solidity < 0.78) {
      contour.delete();
      continue;
    }

    const moments = cv.moments(contour);
    contour.delete();
    if (moments.m00 === 0) {
      continue;
    }
    const center = { x: moments.m10 / moments.m00, y: moments.m01 / moments.m00 };
    const cornerScore = cornerProximityScore(center, width, height);

    candidates.push({
      area,
      center,
      cornerScore,
    });
  }

  candidates.sort((left, right) => (right.area + right.cornerScore * 1000) - (left.area + left.cornerScore * 1000));
  const unique = [];
  for (const candidate of candidates) {
    const duplicate = unique.some((item) => distance(item.center, candidate.center) < 18);
    if (!duplicate) {
      unique.push(candidate);
    }
    if (unique.length >= 14) {
      break;
    }
  }

  if (unique.length < 4) {
    return [];
  }

  let best = null;
  const limit = Math.min(unique.length, 10);
  for (let a = 0; a < limit - 3; a += 1) {
    for (let b = a + 1; b < limit - 2; b += 1) {
      for (let c = b + 1; c < limit - 1; c += 1) {
        for (let d = c + 1; d < limit; d += 1) {
          const combo = [unique[a], unique[b], unique[c], unique[d]];
          const ordered = orderCorners(combo.map((item) => item.center));
          const rectScore = rectangleScore(ordered, width, height);
          const cornerBoost =
            (combo.reduce((sum, item) => sum + (item.cornerScore || 0), 0) / 4) * MARKER_CORNER_BONUS;
          const score = rectScore * (1 + cornerBoost);
          if (!best || score > best.score) {
            best = { score, markers: combo };
          }
        }
      }
    }
  }

  return best ? best.markers : unique.slice(0, 4);
}

function collectMarkerCandidatesV3(contours, width, height) {
  const candidates = [];
  const frameArea = width * height;
  const minArea = frameArea * MARKER_MIN_AREA_RATIO;
  const maxArea = frameArea * MARKER_MAX_AREA_RATIO;
  const minDim = Math.min(width, height);
  const minSide = Math.max(8, Math.round(minDim * MARKER_MIN_SIDE_RATIO));
  const maxSide = Math.max(24, Math.round(minDim * MARKER_MAX_SIDE_RATIO));

  for (let index = 0; index < contours.size(); index += 1) {
    const contour = contours.get(index);
    const area = cv.contourArea(contour);
    if (area < minArea || area > maxArea) {
      contour.delete();
      continue;
    }

    let rectWidth = 0;
    let rectHeight = 0;
    try {
      const rrect = cv.minAreaRect(contour);
      rectWidth = Math.max(rrect.size.width, rrect.size.height);
      rectHeight = Math.min(rrect.size.width, rrect.size.height);
    } catch {
      const rect = cv.boundingRect(contour);
      rectWidth = rect.width;
      rectHeight = rect.height;
    }

    const aspect = rectWidth / Math.max(rectHeight, 1);
    const fillRatio = area / Math.max(rectWidth * rectHeight, 1);

    if (
      aspect < 0.66 ||
      aspect > 1.52 ||
      rectWidth < minSide ||
      rectHeight < minSide ||
      rectWidth > maxSide ||
      rectHeight > maxSide ||
      fillRatio < 0.50
    ) {
      contour.delete();
      continue;
    }

    const hull = new cv.Mat();
    cv.convexHull(contour, hull, false, true);
    const hullArea = cv.contourArea(hull);
    const solidity = area / Math.max(hullArea, 1);
    hull.delete();

    if (solidity < 0.72) {
      contour.delete();
      continue;
    }

    const moments = cv.moments(contour);
    contour.delete();
    if (moments.m00 === 0) {
      continue;
    }

    const center = { x: moments.m10 / moments.m00, y: moments.m01 / moments.m00 };
    const cornerScore = cornerProximityScore(center, width, height);
    const aspectSquare = Math.min(rectWidth, rectHeight) / Math.max(rectWidth, rectHeight, 1);
    const squareScore = clamp(aspectSquare * fillRatio * solidity, 0, 1);

    candidates.push({
      area,
      center,
      cornerScore,
      squareScore,
    });
  }

  candidates.sort((left, right) => {
    const leftScore = (left.squareScore || 0) * 10000 + (left.cornerScore || 0) * 1200 + Math.sqrt(left.area || 0);
    const rightScore = (right.squareScore || 0) * 10000 + (right.cornerScore || 0) * 1200 + Math.sqrt(right.area || 0);
    return rightScore - leftScore;
  });

  const unique = [];
  for (const candidate of candidates) {
    const duplicate = unique.some((item) => distance(item.center, candidate.center) < 18);
    if (!duplicate) {
      unique.push(candidate);
    }
    if (unique.length >= 20) {
      break;
    }
  }

  if (unique.length < 4) {
    return [];
  }

  let best = null;
  const limit = Math.min(unique.length, 14);
  for (let a = 0; a < limit - 3; a += 1) {
    for (let b = a + 1; b < limit - 2; b += 1) {
      for (let c = b + 1; c < limit - 1; c += 1) {
        for (let d = c + 1; d < limit; d += 1) {
          const combo = [unique[a], unique[b], unique[c], unique[d]];

          const centers = combo.map((item) => item.center);
          let minDist = Infinity;
          for (let i = 0; i < centers.length; i += 1) {
            for (let j = i + 1; j < centers.length; j += 1) {
              minDist = Math.min(minDist, distance(centers[i], centers[j]));
            }
          }
          if (minDist < Math.min(width, height) * 0.06) {
            continue;
          }

          const ordered = orderCorners(combo.map((item) => item.center));
          const rectScore = rectangleScore(ordered, width, height);
          const cornerBoost =
            (combo.reduce((sum, item) => sum + (item.cornerScore || 0), 0) / 4) * MARKER_CORNER_BONUS;
          const squareBoost = combo.reduce((sum, item) => sum + (item.squareScore || 0), 0) / 4;
          const score = rectScore * (1 + cornerBoost) * (0.55 + (squareBoost * 0.45));
          if (!best || score > best.score) {
            best = { score, markers: combo };
          }
        }
      }
    }
  }

  return best ? best.markers : unique.slice(0, 4);
}

function rectangleScore(points, frameWidth, frameHeight) {
  const [topLeft, topRight, bottomRight, bottomLeft] = points;
  const widthTop = distance(topLeft, topRight);
  const widthBottom = distance(bottomLeft, bottomRight);
  const heightLeft = distance(topLeft, bottomLeft);
  const heightRight = distance(topRight, bottomRight);
  const widthBalance = 1 - Math.abs(widthTop - widthBottom) / Math.max(widthTop, widthBottom, 1);
  const heightBalance = 1 - Math.abs(heightLeft - heightRight) / Math.max(heightLeft, heightRight, 1);
  const areaScore = ((widthTop + widthBottom) / 2) * ((heightLeft + heightRight) / 2);
  const normalizedArea = areaScore / Math.max(frameWidth * frameHeight, 1);

  // Penaliza quadriláteros com proporção muito diferente do cartão.
  const quadAspect = ((widthTop + widthBottom) / 2) / Math.max((heightLeft + heightRight) / 2, 1);
  const aspectError = Math.abs(quadAspect - CARD_ASPECT) / Math.max(CARD_ASPECT, 1e-6);
  const aspectScore = clamp(1 - aspectError, 0, 1);

  return normalizedArea * widthBalance * heightBalance * (0.35 + (aspectScore * 0.65));
}

function perspectiveWarp(grayMat, corners) {
  const srcTri = cv.matFromArray(4, 1, cv.CV_32FC2, [
    corners[0].x, corners[0].y,
    corners[1].x, corners[1].y,
    corners[3].x, corners[3].y,
    corners[2].x, corners[2].y,
  ]);
  const dstTri = cv.matFromArray(4, 1, cv.CV_32FC2, [
    CARD_TARGET.leftMarkerX, CARD_TARGET.topMarkerY,
    CARD_TARGET.rightMarkerX, CARD_TARGET.topMarkerY,
    CARD_TARGET.leftMarkerX, CARD_TARGET.bottomMarkerY,
    CARD_TARGET.rightMarkerX, CARD_TARGET.bottomMarkerY,
  ]);
  const matrix = cv.getPerspectiveTransform(srcTri, dstTri);
  const warped = new cv.Mat();
  cv.warpPerspective(
    grayMat,
    warped,
    matrix,
    new cv.Size(CARD_TARGET.width, CARD_TARGET.height),
    cv.INTER_LINEAR,
    cv.BORDER_CONSTANT,
    new cv.Scalar(255, 255, 255, 255),
  );
  srcTri.delete();
  dstTri.delete();
  matrix.delete();
  return warped;
}

function measureFocus(grayMat) {
  const laplacian = new cv.Mat();
  const mean = new cv.Mat();
  const stddev = new cv.Mat();
  cv.Laplacian(grayMat, laplacian, cv.CV_64F);
  cv.meanStdDev(laplacian, mean, stddev);
  const score = stddev.doubleAt(0, 0) ** 2;
  laplacian.delete();
  mean.delete();
  stddev.delete();
  return score;
}

function readAnswersFromWarped(warpedGray, questionCount) {
  const answers = {};
  const rows = [];
  const thresholded = new cv.Mat();
  cv.threshold(warpedGray, thresholded, 0, 255, cv.THRESH_BINARY_INV + cv.THRESH_OTSU);

  const stepX = (CARD_TARGET.rightMarkerX - CARD_TARGET.leftMarkerX) / 6;
  const stepY = (CARD_TARGET.bottomMarkerY - CARD_TARGET.topMarkerY) / (questionCount + 1);
  const bubbleRadius = Math.min(stepX, stepY) * 0.34;
  const sampleRadius = bubbleRadius * 0.58;

  for (let questionIndex = 1; questionIndex <= questionCount; questionIndex += 1) {
    const centerY = CARD_TARGET.topMarkerY + (stepY * questionIndex);
    const scores = [];
    const centers = [];

    for (let optionIndex = 1; optionIndex <= 5; optionIndex += 1) {
      const centerX = CARD_TARGET.leftMarkerX + (stepX * optionIndex);
      centers.push({ x: centerX, y: centerY });
      scores.push(sampleBubbleScore(thresholded, centerX, centerY, sampleRadius));
    }

    const markedIndices = resolveMarkedIndices(scores);
    const selectedIndex = markedIndices.length === 1 ? markedIndices[0] : -1;
    const letter = selectedIndex >= 0 ? ["A", "B", "C", "D", "E"][selectedIndex] : "";

    if (letter) {
      answers[String(questionIndex)] = letter;
    }

    rows.push({
      questionNumber: questionIndex,
      centers,
      markedIndices,
      selectedIndex,
      bubbleRadius,
    });
  }

  thresholded.delete();
  return { answers, rows };
}

function readAnswersFromWarpedV2(warpedGray, questionCount) {
  const answers = {};
  const rows = [];

  const stepX = (CARD_TARGET.rightMarkerX - CARD_TARGET.leftMarkerX) / 6;
  const stepY = (CARD_TARGET.bottomMarkerY - CARD_TARGET.topMarkerY) / (questionCount + 1);
  const bubbleRadius = Math.min(stepX, stepY) * 0.34;
  const roiRadius = bubbleRadius * 0.62;
  const roiSize = Math.max(22, Math.round(roiRadius * 2.9));

  const innerMask = new cv.Mat.zeros(roiSize, roiSize, cv.CV_8UC1);
  const ringMask = new cv.Mat.zeros(roiSize, roiSize, cv.CV_8UC1);
  const center = new cv.Point(Math.floor(roiSize / 2), Math.floor(roiSize / 2));
  const innerR = Math.max(4, Math.round(roiRadius * 0.9));
  const outerR = Math.max(innerR + 2, Math.round(roiRadius * 1.25));
  cv.circle(innerMask, center, innerR, new cv.Scalar(255), -1);
  cv.circle(ringMask, center, outerR, new cv.Scalar(255), -1);
  cv.circle(ringMask, center, innerR + 1, new cv.Scalar(0), -1);

  try {
    for (let questionIndex = 1; questionIndex <= questionCount; questionIndex += 1) {
      const centerY = CARD_TARGET.topMarkerY + (stepY * questionIndex);
      const centers = [];
      const scores = [];

      for (let optionIndex = 1; optionIndex <= 5; optionIndex += 1) {
        const centerX = CARD_TARGET.leftMarkerX + (stepX * optionIndex);
        centers.push({ x: centerX, y: centerY });
        scores.push(sampleBubbleContrastScore(warpedGray, centerX, centerY, roiSize, innerMask, ringMask));
      }

      const { markedIndices, selectedIndex, confidence, adaptiveThreshold } = classifyBubbleScores(scores);
      const letter = selectedIndex >= 0 ? ["A", "B", "C", "D", "E"][selectedIndex] : "";
      if (letter) {
        answers[String(questionIndex)] = letter;
      }

      rows.push({
        questionNumber: questionIndex,
        centers,
        scores,
        markedIndices,
        selectedIndex,
        bubbleRadius,
        confidence,
        adaptiveThreshold,
      });
    }
  } finally {
    innerMask.delete();
    ringMask.delete();
  }

  return { answers, rows };
}

function clamp(value, min, max) {
  return Math.max(min, Math.min(max, value));
}

function sampleBubbleContrastScore(grayMat, centerX, centerY, roiSize, innerMask, ringMask) {
  const half = Math.floor(roiSize / 2);
  const x0 = clamp(Math.round(centerX) - half, 0, grayMat.cols - roiSize);
  const y0 = clamp(Math.round(centerY) - half, 0, grayMat.rows - roiSize);
  const rect = new cv.Rect(x0, y0, roiSize, roiSize);
  const roi = grayMat.roi(rect);
  try {
    const bubbleMean = cv.mean(roi, innerMask)[0];
    const ringMean = cv.mean(roi, ringMask)[0];
    const local = Math.max(8, ringMean);
    const contrast = clamp((ringMean - bubbleMean) / local, 0, 1);
    const darkness = clamp((255 - bubbleMean) / 255, 0, 1);
    return clamp((contrast * 0.78) + (darkness * 0.22), 0, 1);
  } finally {
    roi.delete();
  }
}

function classifyBubbleScores(scores) {
  const sorted = scores
    .map((score, index) => ({ score, index }))
    .sort((a, b) => b.score - a.score);

  const best = sorted[0];
  const second = sorted[1] || { score: 0, index: -1 };

  const mean = scores.reduce((sum, value) => sum + value, 0) / Math.max(scores.length, 1);
  const variance = scores.reduce((sum, value) => sum + ((value - mean) ** 2), 0) / Math.max(scores.length, 1);
  const std = Math.sqrt(variance);

  const adaptiveThreshold = Math.max(BUBBLE_CONTRAST_MIN, mean + (std * 0.75));
  if (best.score < adaptiveThreshold) {
    return { markedIndices: [], selectedIndex: -1, confidence: 0, adaptiveThreshold };
  }

  const relativeAmbiguous = second.score >= best.score * BUBBLE_AMBIGUOUS_RELATIVE;
  const gapAmbiguous = (best.score - second.score) <= BUBBLE_AMBIGUOUS_GAP;
  const ambiguous = relativeAmbiguous || gapAmbiguous;

  if (ambiguous) {
    const threshold = Math.max(adaptiveThreshold, best.score * 0.92);
    const markedIndices = scores
      .map((score, index) => ({ score, index }))
      .filter((item) => item.score >= threshold)
      .map((item) => item.index);
    return {
      markedIndices,
      selectedIndex: -1,
      confidence: clamp(best.score - second.score, 0, 1),
      adaptiveThreshold,
    };
  }

  return {
    markedIndices: [best.index],
    selectedIndex: best.index,
    confidence: clamp(best.score - second.score, 0, 1),
    adaptiveThreshold,
  };
}

function drawGridOverlay(baseImageData, rows) {
  if (!baseImageData || !rows?.length || !state.elements?.alignedCanvas) {
    return;
  }
  drawAlignedPreview(baseImageData);
  const context = state.elements.alignedCanvas.getContext("2d");
  context.lineWidth = 2;
  context.strokeStyle = "#ffb300";
  context.fillStyle = "rgba(255,179,0,0.25)";

  for (const row of rows) {
    for (const center of row.centers) {
      context.beginPath();
      context.arc(center.x, center.y, row.bubbleRadius * 0.95, 0, Math.PI * 2);
      context.stroke();
    }
  }
}

function sampleBubbleScore(binaryMat, centerX, centerY, radius) {
  let whitePixels = 0;
  let totalPixels = 0;

  const startX = Math.max(0, Math.floor(centerX - radius));
  const endX = Math.min(binaryMat.cols - 1, Math.ceil(centerX + radius));
  const startY = Math.max(0, Math.floor(centerY - radius));
  const endY = Math.min(binaryMat.rows - 1, Math.ceil(centerY + radius));
  const radiusSquared = radius * radius;

  for (let y = startY; y <= endY; y += 1) {
    for (let x = startX; x <= endX; x += 1) {
      const dx = x - centerX;
      const dy = y - centerY;
      if ((dx * dx) + (dy * dy) > radiusSquared) {
        continue;
      }
      totalPixels += 1;
      if (binaryMat.ucharPtr(y, x)[0] > 0) {
        whitePixels += 1;
      }
    }
  }

  return totalPixels ? whitePixels / totalPixels : 0;
}

function resolveMarkedIndices(scores) {
  const bestScore = Math.max(...scores);
  if (bestScore < MIN_FILLED_BUBBLE_SCORE) {
    return [];
  }
  const threshold = Math.max(MIN_FILLED_BUBBLE_SCORE, bestScore * SECOND_BUBBLE_RELATIVE_LIMIT);
  return scores
    .map((score, index) => ({ score, index }))
    .filter((item) => item.score >= threshold)
    .map((item) => item.index);
}

function answersSignature(answers, questionCount) {
  const values = [];
  for (let index = 1; index <= questionCount; index += 1) {
    values.push(answers[String(index)] || "-");
  }
  return values.join("");
}

function drawOverlay(corners) {
  const width = state.elements.video.clientWidth || state.elements.video.videoWidth || 1;
  const height = state.elements.video.clientHeight || state.elements.video.videoHeight || 1;
  state.elements.overlayCanvas.width = width;
  state.elements.overlayCanvas.height = height;
  const context = state.elements.overlayCanvas.getContext("2d");
  context.clearRect(0, 0, width, height);

  const scaleX = width / state.elements.video.videoWidth;
  const scaleY = height / state.elements.video.videoHeight;
  context.strokeStyle = "#34c759";
  context.lineWidth = 4;
  context.beginPath();
  context.moveTo(corners[0].x * scaleX, corners[0].y * scaleY);
  context.lineTo(corners[1].x * scaleX, corners[1].y * scaleY);
  context.lineTo(corners[2].x * scaleX, corners[2].y * scaleY);
  context.lineTo(corners[3].x * scaleX, corners[3].y * scaleY);
  context.closePath();
  context.stroke();
}

function clearOverlay() {
  if (state.elements.overlayCanvas) {
    const context = state.elements.overlayCanvas.getContext("2d");
    context.clearRect(0, 0, state.elements.overlayCanvas.width, state.elements.overlayCanvas.height);
  }
  if (state.elements.alignedCanvas) {
    const context = state.elements.alignedCanvas.getContext("2d");
    context.clearRect(0, 0, state.elements.alignedCanvas.width, state.elements.alignedCanvas.height);
  }
}

function drawAlignedPreview(imageData) {
  if (!imageData || !state.elements.alignedCanvas) {
    return;
  }
  state.elements.alignedCanvas.width = imageData.width;
  state.elements.alignedCanvas.height = imageData.height;
  const context = state.elements.alignedCanvas.getContext("2d");
  context.putImageData(imageData, 0, 0);
}

function drawCorrectionPreview(result) {
  if (!state.lastDetection?.baseImageData) {
    return;
  }
  drawAlignedPreview(state.lastDetection.baseImageData);
  const context = state.elements.alignedCanvas.getContext("2d");
  context.lineWidth = 6;

  for (const detail of result.detalhes) {
    const row = state.lastDetection.rows.find((item) => item.questionNumber === detail.numero);
    if (!row) {
      continue;
    }
    if (detail.acertou && row.selectedIndex >= 0) {
      drawCircle(context, row.centers[row.selectedIndex], row.bubbleRadius * 1.25, "#1ca44a");
      continue;
    }
    if (row.markedIndices.length) {
      for (const markedIndex of row.markedIndices) {
        drawCircle(context, row.centers[markedIndex], row.bubbleRadius * 1.25, "#d93025");
      }
    } else {
      const left = row.centers[0].x - (row.bubbleRadius * 3.3);
      const top = row.centers[0].y - (row.bubbleRadius * 1.2);
      const width = (row.centers[4].x - row.centers[0].x) + (row.bubbleRadius * 2.4);
      const height = row.bubbleRadius * 2.4;
      context.strokeStyle = "#d93025";
      context.strokeRect(left, top, width, height);
    }
  }
}

function drawCircle(context, center, radius, color) {
  context.strokeStyle = color;
  context.beginPath();
  context.arc(center.x, center.y, radius, 0, Math.PI * 2);
  context.stroke();
}

function correctProof() {
  const result = correctProofLocally(state.proof, state.answers);
  drawCorrectionPreview(result);
  state.elements.resultPanel.classList.remove("hidden");
  state.elements.resultPanel.classList.toggle("wrong", result.acertos !== result.total);
  state.elements.resultPanel.innerHTML = `
    <div><strong>Aluno:</strong> ${escapeHtml(result.aluno || "")}</div>
    <div><strong>Acertos:</strong> ${result.acertos} de ${result.total}</div>
    <div><strong>Nota:</strong> ${result.nota}</div>
    <div><strong>Legenda:</strong> verde = correta, vermelho = errada.</div>
  `;
  setStatus("Leitura concluída.");
}

function resetReadingUi() {
  state.answers = {};
  state.stableSignature = "";
  state.stableCount = 0;
  state.lastDetection = null;
  if (state.elements?.resultPanel) {
    state.elements.resultPanel.classList.add("hidden");
    state.elements.resultPanel.innerHTML = "";
  }
  clearOverlay();
  updateRenderedAnswers();
  if (state.elements?.alignedCanvas) {
    const context = state.elements.alignedCanvas.getContext("2d");
    context?.clearRect(0, 0, state.elements.alignedCanvas.width, state.elements.alignedCanvas.height);
  }
}

async function captureAndProcessPhoto() {
  if (!state.proof) {
    setStatus("Nenhuma prova carregada.");
    return;
  }
  if (!state.opencvReady) {
    setStatus("OpenCV ainda está carregando.");
    return;
  }
  if (!state.stream || !state.elements.video?.srcObject) {
    setStatus("Habilite a câmera antes de capturar a foto.");
    return;
  }

  resetReadingUi();
  setWorkflowState("answerScanning", "Capturando foto…");
  updateCardControls();

  // Captura frame estático do preview.
  const video = state.elements.video;
  const videoWidth = video.videoWidth || 0;
  const videoHeight = video.videoHeight || 0;
  if (videoWidth < 120 || videoHeight < 120) {
    setWorkflowState("error", "Vídeo ainda não está pronto. Tente novamente.");
    updateCardControls();
    return;
  }

  const scale = Math.min(1, CARD_PROCESSING_MAX_WIDTH / Math.max(videoWidth, 1));
  processingCanvas.width = Math.max(1, Math.round(videoWidth * scale));
  processingCanvas.height = Math.max(1, Math.round(videoHeight * scale));
  const ctx = processingCanvas.getContext("2d", { willReadFrequently: true });
  ctx.drawImage(video, 0, 0, processingCanvas.width, processingCanvas.height);

  // Fecha o preview para evitar travas durante o processamento.
  stopAnswerCamera();

  setWorkflowState("answerScanning", "Processando foto…");
  updateCardControls();

  let detection;
  try {
    detection = detectCardAnswersFromCanvas(processingCanvas, state.proof.quantidade_questoes);
  } catch {
    detection = { ok: false, reason: "unknown" };
  }

  if (!detection || detection.ok === false) {
    const reason = detection?.reason || "unknown";
    if (reason === "markers_not_found") {
      setWorkflowState(
        "error",
        "Não foi possível identificar os quatro quadradinhos de alinhamento. Tire uma nova foto com o cartão inteiro visível e boa iluminação.",
      );
    } else if (reason === "bad_geometry") {
      setWorkflowState(
        "error",
        "Os marcadores foram encontrados, mas o cartão está muito inclinado ou parcialmente fora da imagem. Tire uma nova foto.",
      );
    } else if (reason === "blur") {
      setWorkflowState(
        "error",
        "A foto ficou desfocada. Afaste um pouco o celular, estabilize e tire uma nova foto com boa iluminação.",
      );
    } else {
      setWorkflowState(
        "error",
        "Não foi possível ler o cartão com confiança. Alinhe melhor e tire uma nova foto.",
      );
    }
    updateCardControls();
    return;
  }

  state.lastDetection = detection;
  state.answers = { ...detection.answers };
  updateRenderedAnswers();
  drawAlignedPreview(detection.baseImageData);
  correctProof();
  setWorkflowState("answerDetected");
  updateCardControls();
}

function retakePhoto() {
  resetReadingUi();
  setWorkflowState("waitingUserToStartAnswerScan", "Toque em “Habilitar câmera” e alinhe o cartão.");
  updateCardControls();
  enableAnswerCamera();
}

function confirmReading() {
  setStatus("Leitura confirmada.");
  // Mantém resultados na tela; usuário pode tocar em "Ler outro QR" se desejar.
}

function cancelReading() {
  resetReadingUi();
  stopAnswerCamera();
  setWorkflowState("waitingUserToStartAnswerScan", "Leitura cancelada. Toque em “Habilitar câmera” para tentar novamente.");
  updateCardControls();
}

function correctProofLocally(proof, answers) {
  const details = [];
  let correct = 0;
  for (const question of proof.questoes || []) {
    const marked = String(answers[String(question.numero)] || "").toUpperCase();
    const expected = String(question.correta_letra || "").toUpperCase();
    const hit = marked !== "" && marked === expected;
    if (hit) {
      correct += 1;
    }
    details.push({
      numero: question.numero,
      marcada: marked,
      correta: expected,
      acertou: hit,
    });
  }
  const total = details.length;
  return {
    aluno: proof.aluno || "",
    acertos: correct,
    total,
    nota: total ? Math.round((correct / total) * 1000) / 100 : 0,
    detalhes: details,
  };
}

function parseQrPayload(rawValue) {
  const normalized = String(rawValue || "").trim();
  if (!normalized.startsWith(QR_PAYLOAD_PREFIX)) {
    throw new Error("QR Code incompativel com o GabMath.");
  }

  const encoded = normalized.slice(QR_PAYLOAD_PREFIX.length);
  const jsonText = decodeBase64Url(encoded);
  const payload = JSON.parse(jsonText);
  const answerKey = String(payload.g || "").toUpperCase();
  const questions = Array.from(answerKey).map((letter, index) => ({
    numero: index + 1,
    correta_letra: letter,
  }));

  return {
    id_prova: payload.i || payload.id || "",
    aluno: payload.s || payload.a || "",
    quantidade_questoes: questions.length,
    questoes: questions,
  };
}

function decodeBase64Url(value) {
  const normalized = value.replaceAll("-", "+").replaceAll("_", "/");
  const padded = normalized + "=".repeat((4 - (normalized.length % 4)) % 4);
  const binary = atob(padded);
  const bytes = Uint8Array.from(binary, (char) => char.charCodeAt(0));
  return new TextDecoder().decode(bytes);
}

function matToImageData(mat) {
  const rgba = new Uint8ClampedArray(mat.data);
  return new ImageData(rgba, mat.cols, mat.rows);
}

function orderCorners(points) {
  const bySum = [...points].sort((left, right) => (left.x + left.y) - (right.x + right.y));
  const topLeft = bySum[0];
  const bottomRight = bySum[3];
  const remaining = bySum.slice(1, 3).sort((left, right) => (left.x - left.y) - (right.x - right.y));
  const bottomLeft = remaining[0];
  const topRight = remaining[1];
  return [topLeft, topRight, bottomRight, bottomLeft];
}

function distance(left, right) {
  return Math.hypot(left.x - right.x, left.y - right.y);
}

function setStatus(text) {
  if (state.elements.scanStatus && !state.cardMode) {
    state.elements.scanStatus.textContent = text;
  }
  if (state.elements.cardStatus && state.cardMode) {
    state.elements.cardStatus.textContent = text;
  }
}

function setBadge(text) {
  if (state.elements.modeBadge) {
    state.elements.modeBadge.textContent = text;
  }
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function setWorkflowState(next, statusText) {
  state.workflow = next;
  if (statusText) {
    setStatus(statusText);
  }
  if (!state.elements?.modeBadge) {
    return;
  }
  const badgeText = {
    idle: "Aguardando",
    qrScanning: "Lendo QR",
    qrDetected: "QR detectado",
    waitingUserToStartAnswerScan: "Aguardando leitura",
    answerScanning: "Lendo cartão",
    answerDetected: "Leitura concluída",
    error: "Erro",
  }[next] || next;
  setBadge(badgeText);
}

function updateCardControls() {
  if (!state.cardMode) {
    setCameraPanelVisible(state.workflow === "qrScanning");
    return;
  }

  const cameraEnabled = Boolean(state.stream && state.elements.video?.srcObject);
  const scanning = state.workflow === "answerScanning";
  const photoMode = true;
  const photoReady = state.workflow === "answerDetected";

  setCameraPanelVisible(cameraEnabled);

  // Novo layout (index.html atualizado)
  if (state.elements?.enableCameraButton) {
    state.elements.enableCameraButton.classList.toggle("hidden", cameraEnabled);
    state.elements.enableCameraButton.disabled = scanning || !state.opencvReady;

    if (state.elements.capturePhotoButton) {
      state.elements.capturePhotoButton.classList.toggle("hidden", !cameraEnabled || scanning || photoReady);
      state.elements.capturePhotoButton.disabled = scanning || !cameraEnabled;
    }

    if (state.elements.retakePhotoButton) {
      state.elements.retakePhotoButton.classList.toggle("hidden", scanning || (!photoReady && cameraEnabled));
      state.elements.retakePhotoButton.disabled = scanning;
    }

    if (state.elements.confirmReadingButton) {
      state.elements.confirmReadingButton.classList.toggle("hidden", !photoReady);
      state.elements.confirmReadingButton.disabled = false;
    }

    if (state.elements.cancelReadingButton) {
      const showCancel = photoReady || cameraEnabled || scanning;
      state.elements.cancelReadingButton.classList.toggle("hidden", !showCancel);
      state.elements.cancelReadingButton.disabled = false;
    }

    if (state.elements.disableCameraButton) {
      state.elements.disableCameraButton.classList.toggle("hidden", !cameraEnabled || scanning || photoReady);
      state.elements.disableCameraButton.disabled = scanning;
    }

    if (cameraEnabled && !scanning && !photoReady) {
      drawAnswerGuide();
    }
    return;
  }

  // Compatibilidade: card.html antigo.
  if (state.elements?.startScanButton && state.elements?.stopScanButton && state.elements?.startCardScanButton) {
    state.elements.startScanButton.disabled = scanning;
    state.elements.stopScanButton.disabled = false;
    state.elements.startCardScanButton.disabled = !cameraEnabled || !state.opencvReady || scanning;
  }
}

async function startAnswerScan() {
  if (!state.proof) {
    setStatus("Nenhuma prova carregada.");
    return;
  }
  if (!state.opencvReady) {
    setStatus("OpenCV ainda está carregando.");
    return;
  }

  if (!state.stream) {
    await startCardCamera();
    return;
  }

  startCardReading();
}

function getQueryFlag(name) {
  try {
    const value = new URLSearchParams(window.location.search).get(name);
    if (value === null) {
      return false;
    }
    return value === "" || value === "1" || value === "true" || value === "on" || value === "yes";
  } catch {
    return false;
  }
}
