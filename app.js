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
const MIN_CARD_RECTANGLE_SCORE = 0.18;
const MIN_MARKER_FILL_RATIO = 0.72;
const MIN_MARKER_SIDE = 16;
const MAX_MARKER_SIDE = 90;
const MIN_FILLED_BUBBLE_SCORE = 0.24;
const SECOND_BUBBLE_RELATIVE_LIMIT = 0.86;

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
    video: document.getElementById("camera-video"),
    overlayCanvas: document.getElementById("camera-overlay"),
    startScanButton: document.getElementById("start-scan"),
    stopScanButton: document.getElementById("stop-scan"),
    loadProofButton: document.getElementById("load-proof"),
    proofIdInput: document.getElementById("proof-id-input"),
    scanStatus: document.getElementById("scan-status"),
  };

  state.elements.startScanButton.addEventListener("click", startQrCamera);
  state.elements.stopScanButton.addEventListener("click", stopWorkflow);
  state.elements.loadProofButton.addEventListener("click", () => commitQrAndGo(state.elements.proofIdInput.value));

  setStatus("Aponte a camera apenas para o QR Code.");
}

function initializeCardPage() {
  state.elements = {
    video: document.getElementById("camera-video"),
    overlayCanvas: document.getElementById("camera-overlay"),
    alignedCanvas: document.getElementById("aligned-preview"),
    startScanButton: document.getElementById("start-scan"),
    stopScanButton: document.getElementById("stop-scan"),
    startCardScanButton: document.getElementById("start-card-scan"),
    proofSummary: document.getElementById("proof-summary"),
    scanStatus: document.getElementById("scan-status"),
    answerGrid: document.getElementById("answer-grid"),
    resultPanel: document.getElementById("result-panel"),
  };

  state.elements.startScanButton.addEventListener("click", startCardCamera);
  state.elements.stopScanButton.addEventListener("click", stopWorkflow);
  state.elements.startCardScanButton.addEventListener("click", startCardReading);

  waitForOpenCv();
  const stored = sessionStorage.getItem(STORAGE_KEY);
  if (!stored) {
    setStatus("Nenhuma prova foi carregada. Volte e leia o QR Code primeiro.");
    state.elements.startCardScanButton.disabled = true;
    return;
  }

  state.proof = JSON.parse(stored);
  renderProofSummary();
  renderAnswerGrid();
  setStatus("Abra a camera e depois alinhe o cartao-resposta com calma.");
}

async function startQrCamera() {
  await startCamera();
  setStatus("Camera aberta. Aponte apenas para o QR Code.");
  startQrLoop();
}

async function startCardCamera() {
  await startCamera();
  setStatus("Camera aberta. Posicione o cartao com calma antes de iniciar a leitura.");
}

async function startCamera() {
  if (!navigator.mediaDevices?.getUserMedia) {
    setStatus("Camera nao disponivel neste navegador.");
    return;
  }

  stopLoops();
  clearOverlay();

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
  } catch (error) {
    setStatus(`Falha ao abrir a camera: ${error?.message || String(error)}`);
  }
}

function stopWorkflow() {
  stopLoops();
  if (state.stream) {
    for (const track of state.stream.getTracks()) {
      track.stop();
    }
    state.stream = null;
  }
  if (state.elements.video) {
    state.elements.video.srcObject = null;
  }
  clearOverlay();
  setStatus("Leitura interrompida.");
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
  state.qrTimer = setInterval(async () => {
    if (!state.elements.video?.srcObject) {
      return;
    }
    try {
      const rawValue = await detectQrCode(state.elements.video);
      if (!rawValue) {
        return;
      }
      state.elements.proofIdInput.value = rawValue;
      commitQrAndGo(rawValue);
    } catch (error) {
      setStatus(`Falha na leitura do QR: ${error?.message || String(error)}`);
    }
  }, 450);
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
    captureCanvas.width = videoElement.videoWidth;
    captureCanvas.height = videoElement.videoHeight;
    const context = captureCanvas.getContext("2d", { willReadFrequently: true });
    context.drawImage(videoElement, 0, 0, captureCanvas.width, captureCanvas.height);
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

function commitQrAndGo(rawValue) {
  const normalized = String(rawValue || "").trim();
  if (!normalized) {
    setStatus("Informe ou leia um QR valido.");
    return;
  }

  try {
    const proof = parseQrPayload(normalized);
    sessionStorage.setItem(STORAGE_KEY, JSON.stringify(proof));
    window.location.href = "./card.html";
  } catch (error) {
    setStatus(error?.message || String(error));
  }
}

function waitForOpenCv() {
  if (window.cv && typeof window.cv.Mat === "function") {
    state.opencvReady = true;
    state.opencvKernel = cv.getStructuringElement(cv.MORPH_RECT, new cv.Size(3, 3));
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
    }
  }, 250);
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
  if (!state.proof) {
    setStatus("Nenhuma prova carregada.");
    return;
  }
  if (!state.stream || !state.elements.video?.srcObject) {
    setStatus("Abra a camera antes de iniciar a leitura do cartao.");
    return;
  }
  if (!state.opencvReady) {
    setStatus("OpenCV ainda esta carregando.");
    return;
  }

  state.answers = {};
  state.stableSignature = "";
  state.stableCount = 0;
  state.lastDetection = null;
  state.elements.resultPanel.classList.add("hidden");
  state.elements.resultPanel.innerHTML = "";
  clearOverlay();
  stopLoops();

  setStatus("Alinhe os 4 quadrados pretos. A leitura so fecha quando a imagem estiver nitida.");

  let lastTimestamp = 0;
  const tick = (timestamp) => {
    if (!state.elements.video?.srcObject) {
      state.cardTimer = requestAnimationFrame(tick);
      return;
    }

    if (timestamp - lastTimestamp < CARD_FRAME_INTERVAL_MS) {
      state.cardTimer = requestAnimationFrame(tick);
      return;
    }
    lastTimestamp = timestamp;

    const detection = detectCardAnswers(state.elements.video, state.proof.quantidade_questoes);
    if (!detection) {
      state.stableSignature = "";
      state.stableCount = 0;
      setStatus("Nao foi possivel alinhar o cartao. Mantenha visiveis os 4 quadrados pretos.");
      state.cardTimer = requestAnimationFrame(tick);
      return;
    }

    state.lastDetection = detection;
    drawOverlay(detection.corners);
    drawAlignedPreview(detection.baseImageData);

    if (detection.focusScore < MIN_FOCUS_SCORE) {
      state.stableSignature = "";
      state.stableCount = 0;
      setStatus(`Imagem ainda sem nitidez suficiente (${Math.round(detection.focusScore)}). Aproxime e estabilize o celular.`);
      state.cardTimer = requestAnimationFrame(tick);
      return;
    }

    const signature = answersSignature(detection.answers, state.proof.quantidade_questoes);
    if (signature === state.stableSignature) {
      state.stableCount += 1;
    } else {
      state.stableSignature = signature;
      state.stableCount = 1;
    }

    state.answers = { ...detection.answers };
    updateRenderedAnswers();
    setStatus(`Cartao alinhado e nitido. Confirmando leitura... ${state.stableCount}/${REQUIRED_STABLE_FRAMES}`);

    if (state.stableCount >= REQUIRED_STABLE_FRAMES) {
      stopLoops();
      correctProof();
      return;
    }

    state.cardTimer = requestAnimationFrame(tick);
  };

  state.cardTimer = requestAnimationFrame(tick);
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
  const contours = new cv.MatVector();
  const hierarchy = new cv.Mat();

  try {
    cv.cvtColor(src, gray, cv.COLOR_RGBA2GRAY);
    cv.GaussianBlur(gray, blur, new cv.Size(5, 5), 0, 0, cv.BORDER_DEFAULT);
    cv.threshold(blur, thresh, 0, 255, cv.THRESH_BINARY_INV + cv.THRESH_OTSU);
    if (state.opencvKernel) {
      cv.morphologyEx(thresh, thresh, cv.MORPH_CLOSE, state.opencvKernel);
    }
    cv.findContours(thresh, contours, hierarchy, cv.RETR_EXTERNAL, cv.CHAIN_APPROX_SIMPLE);

    const markers = collectMarkerCandidates(contours, processingCanvas.width, processingCanvas.height);
    if (markers.length < 4) {
      return null;
    }

    const corners = orderCorners(markers.slice(0, 4).map((marker) => marker.center));
    if (rectangleScore(corners, processingCanvas.width, processingCanvas.height) < MIN_CARD_RECTANGLE_SCORE) {
      return null;
    }
    const warped = perspectiveWarp(gray, corners);
    const focusScore = measureFocus(warped);
    const basePreview = new cv.Mat();
    cv.cvtColor(warped, basePreview, cv.COLOR_GRAY2RGBA);

    const reading = readAnswersFromWarped(warped, questionCount);
    const baseImageData = matToImageData(basePreview);

    warped.delete();
    basePreview.delete();

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
    contours.delete();
    hierarchy.delete();
  }
}

function collectMarkerCandidates(contours, width, height) {
  const candidates = [];
  const minArea = (width * height) * 0.00012;

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

    if (
      approx.rows === 4 &&
      aspect > 0.8 &&
      aspect < 1.2 &&
      fillRatio > MIN_MARKER_FILL_RATIO &&
      rect.width >= MIN_MARKER_SIDE &&
      rect.height >= MIN_MARKER_SIDE &&
      rect.width <= MAX_MARKER_SIDE &&
      rect.height <= MAX_MARKER_SIDE
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
  return normalizedArea * widthBalance * heightBalance;
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
  setStatus("Leitura concluida.");
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
  if (state.elements.scanStatus) {
    state.elements.scanStatus.textContent = text;
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
