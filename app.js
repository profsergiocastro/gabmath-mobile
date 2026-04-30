const QR_PAYLOAD_PREFIX = "GABMATH1:";
const STORAGE_KEY = "gabmath-current-proof";
const RESULTS_STORAGE_KEY = "gabmath-exam-results-v1";
const CARD_TARGET = {
  width: 820,
  height: 1160,
  leftMarkerX: 130,
  rightMarkerX: 690,
  topMarkerY: 180,
  bottomMarkerY: 1010,
};
const CARD_PROCESSING_MAX_WIDTH = 960;
const PHOTO_PROCESSING_MAX_WIDTH = 2200;
const PHOTO_PROCESSING_FALLBACK_MAX_WIDTH = 3600;
// Modo tempo real: processamento amostrado (não processa todo frame).
// Aumentar um pouco o intervalo reduz travamentos em celulares mais fracos.
const CARD_FRAME_INTERVAL_MS = 200;
const MIN_FOCUS_SCORE = 120;
// Quantos frames consecutivos precisam estar estáveis antes de capturar.
const REQUIRED_STABLE_FRAMES = 3;
// Resolução máxima do frame usado para detectar marcadores em tempo real.
// Menor = mais rápido; a leitura final usa um frame maior.
const REALTIME_DETECTION_MAX_WIDTH = 560;
// Score mínimo para aceitar o quadrilátero dos 4 marcadores.
// Valor mais baixo permite foto com o cartão mais distante (menor na imagem),
// mas ainda rejeita casos onde a geometria fica inconsistente.
const MIN_CARD_RECTANGLE_SCORE = 0.015;
const GOOD_CARD_RECTANGLE_SCORE = 0.075;
const MIN_MARKER_FILL_RATIO = 0.62;
const MIN_MARKER_SIDE = 16;
const MAX_MARKER_SIDE = 90;
const MIN_FILLED_BUBBLE_SCORE = 0.24;
const SECOND_BUBBLE_RELATIVE_LIMIT = 0.86;

// Leitura v2 (mais robusta para bolhas circulares, com limiar adaptativo).
const BUBBLE_CONTRAST_MIN = 0.085;
const BUBBLE_AMBIGUOUS_RELATIVE = 0.9;
const BUBBLE_AMBIGUOUS_GAP = 0.035;
// Não assumimos que o cartão esteja próximo aos cantos da foto.
// (O aluno pode fotografar mais distante e o cartão ficar centralizado.)
// Mantido apenas por compatibilidade; não usamos mais como fator de decisão.
const MARKER_CORNER_BONUS = 0.0;
const CARD_ASPECT = CARD_TARGET.width / CARD_TARGET.height;
const MARKER_RECT_ASPECT = (CARD_TARGET.rightMarkerX - CARD_TARGET.leftMarkerX) / Math.max((CARD_TARGET.bottomMarkerY - CARD_TARGET.topMarkerY), 1);

// Marcadores (multiescala): aceitar quadrados pequenos (cartão longe) e grandes (cartão perto).
const MARKER_MIN_AREA_RATIO = 0.000006;
const MARKER_MAX_AREA_RATIO = 0.015;
// Permite marcadores menores quando a foto é tirada mais distante.
const MARKER_MIN_SIDE_RATIO = 0.0028;
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
  cardCaptureMode: "realtime", // "realtime" | "diagnosticPhoto"
  lastPhotoMeta: null,
  workflow: "idle",
  qrProcessing: false,
  lastQrValue: "",
  lastQrAt: 0,
  lastResult: null,
  editing: null,
  elements: {},
  page: "",
};

const captureCanvas = document.createElement("canvas");
const processingCanvas = document.createElement("canvas");
const captureCanvasContext = captureCanvas.getContext("2d", { willReadFrequently: true });
const processingCanvasContext = processingCanvas.getContext("2d", { willReadFrequently: true });

document.addEventListener("DOMContentLoaded", initializeApp);

function normalizeName(name) {
  const trimmed = String(name || "").trim().replace(/\s+/g, " ");
  if (!trimmed) {
    return "";
  }
  try {
    return trimmed
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "");
  } catch {
    return trimmed.toLowerCase();
  }
}

function parseExamIdDate(examId) {
  const raw = String(examId || "").trim();
  const match = raw.match(/^(\d{4})(\d{2})(\d{2})_(\d{2})(\d{2})(\d{2})_/);
  if (!match) {
    return { iso: "", label: "" };
  }
  const [, yyyy, mm, dd, hh, mi, ss] = match;
  const iso = `${yyyy}-${mm}-${dd}T${hh}:${mi}:${ss}`;
  const label = `${dd}/${mm}/${yyyy} às ${hh}h${mi}min${ss}s`;
  return { iso, label };
}

function parseExamAndStudentCardId(fullExamId) {
  const raw = String(fullExamId || "").trim();
  const match = raw.match(/^(.*?)-(\d+)$/);
  if (!match) {
    return { fullExamId: raw, examGroupId: raw, studentCardId: null };
  }
  const examGroupId = match[1];
  const studentCardId = match[2];
  if (!examGroupId) {
    return { fullExamId: raw, examGroupId: raw, studentCardId: null };
  }
  return { fullExamId: raw, examGroupId, studentCardId };
}

function createResultId() {
  try {
    if (window.crypto?.randomUUID) {
      return window.crypto.randomUUID();
    }
  } catch {
    // ignore
  }
  return `r_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 10)}`;
}

function openResultsPage() {
  window.location.href = "./resultados.html";
}

function initializeApp() {
  state.page = document.body.dataset.page || "qr";
  if (state.page === "qr") {
    initializeQrPage();
    return;
  }
  if (state.page === "card") {
    initializeCardPage();
    return;
  }
  if (state.page === "results") {
    initializeResultsPage();
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
    cardPhotoInput: document.getElementById("card-photo-input"),
    proofSummary: document.getElementById("proof-summary"),
    cardHint: document.getElementById("card-hint"),
    answerGrid: document.getElementById("answer-grid"),
    resultPanel: document.getElementById("result-panel"),
    alignedCanvas: document.getElementById("aligned-preview"),
    backToQrButton: document.getElementById("back-to-qr"),
    debugToggle: document.getElementById("debug-toggle"),
    debugFilePanel: document.getElementById("debug-file-panel"),
    openResultsButton: document.getElementById("open-results"),
  };

  state.elements.startScanButton.addEventListener("click", startQrCamera);
  state.elements.stopScanButton.addEventListener("click", stopWorkflow);
  state.elements.loadProofButton.addEventListener("click", () => commitQrAndGo(state.elements.proofIdInput.value));
  wireCardButtons();
  if (state.elements.openResultsButton) {
    state.elements.openResultsButton.addEventListener("click", openResultsPage);
  }
  if (state.elements.backToQrButton) {
    state.elements.backToQrButton.addEventListener("click", backToQrMode);
  }

  waitForOpenCv();
  setStatus("Aponte a camera apenas para o QR Code.");
  if (state.elements.debugToggle) {
    state.elements.debugToggle.addEventListener("change", () => {
      handleCardCaptureModeToggle();
    });
  }
  handleCardCaptureModeToggle();
  if (state.elements.cardPhotoInput && !state.elements.cardPhotoInput.dataset.wired) {
    state.elements.cardPhotoInput.dataset.wired = "1";
    state.elements.cardPhotoInput.addEventListener("change", async (event) => {
      const file = event?.target?.files?.[0];
      event.target.value = "";
      if (!file) return;
      await processPhotoFromFile(file);
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
    cardPhotoInput: document.getElementById("card-photo-input"),
    proofSummary: document.getElementById("proof-summary"),
    cardHint: document.getElementById("card-hint"),
    scanStatus: document.getElementById("scan-status"),
    answerGrid: document.getElementById("answer-grid"),
    resultPanel: document.getElementById("result-panel"),
    debugToggle: document.getElementById("debug-toggle"),
    debugFilePanel: document.getElementById("debug-file-panel"),
  };

  wireCardButtons();

  waitForOpenCv();
  if (state.elements.debugToggle) {
    state.elements.debugToggle.addEventListener("change", () => {
      handleCardCaptureModeToggle();
    });
  }
  handleCardCaptureModeToggle();
  if (state.elements.cardPhotoInput && !state.elements.cardPhotoInput.dataset.wired) {
    state.elements.cardPhotoInput.dataset.wired = "1";
    state.elements.cardPhotoInput.addEventListener("change", async (event) => {
      const file = event?.target?.files?.[0];
      event.target.value = "";
      if (!file) return;
      await processPhotoFromFile(file);
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
  setStatus('Alinhe o celular com o cartão-resposta e toque em "Tirar foto do cartão".');
}

async function initializeResultsPage() {
  state.elements = {
    modeBadge: document.getElementById("mode-badge"),
    examSelect: document.getElementById("exam-select"),
    examNameInput: document.getElementById("exam-name-input"),
    examSummary: document.getElementById("exam-summary"),
    resultsTbody: document.getElementById("results-tbody"),
    sortSelect: document.getElementById("sort-select"),
    studentSearch: document.getElementById("student-search"),
    questionStats: document.getElementById("question-stats"),
    exportStudentsCsv: document.getElementById("export-students-csv"),
    exportQuestionsCsv: document.getElementById("export-questions-csv"),
    exportJson: document.getElementById("export-json"),
    importJsonFile: document.getElementById("import-json-file"),
    importJsonButton: document.getElementById("import-json-button"),
    deleteExamFilter: document.getElementById("delete-exam-filter"),
    deleteExamList: document.getElementById("delete-exam-list"),
    deleteExamsButton: document.getElementById("delete-exams-button"),
    studentDialog: document.getElementById("student-dialog"),
    studentDialogBody: document.getElementById("student-dialog-body"),
    closeStudentDialog: document.getElementById("close-student-dialog"),
    editDialog: document.getElementById("edit-dialog"),
    editDialogBody: document.getElementById("edit-dialog-body"),
    saveEditDialog: document.getElementById("save-edit-dialog"),
    cancelEditDialog: document.getElementById("cancel-edit-dialog"),
  };

  setBadge("Resultados");

  const store = loadResultsStore();
  const examGroupIds = Object.keys(store.examGroups || {}).sort((a, b) => a.localeCompare(b));
  if (!examGroupIds.length) {
    state.elements.examSelect.innerHTML = "<option value=\"\">Nenhuma prova salva</option>";
    state.elements.examSummary.innerHTML = "<div>Nenhum resultado salvo ainda.</div>";
    return;
  }

  const selectedFromQuery = getQueryParam("examGroupId") || getQueryParam("examId");
  const selectedExamGroupId = selectedFromQuery && examGroupIds.includes(selectedFromQuery)
    ? selectedFromQuery
    : examGroupIds[0];

  state.elements.examSelect.innerHTML = examGroupIds
    .map((examGroupId) => {
      const group = store.examGroups?.[examGroupId];
      const name = String(group?.examName || examGroupId);
      return `<option value="${escapeHtml(examGroupId)}"${examGroupId === selectedExamGroupId ? " selected" : ""}>${escapeHtml(name)}</option>`;
    })
    .join("");

  state.elements.examSelect.addEventListener("change", () => {
    const next = state.elements.examSelect.value;
    window.history.replaceState({}, "", `./resultados.html?examGroupId=${encodeURIComponent(next)}`);
    renderResultsPage();
  });
  state.elements.sortSelect.addEventListener("change", renderResultsPage);
  state.elements.studentSearch.addEventListener("input", renderResultsPage);
  state.elements.exportStudentsCsv.addEventListener("click", () => exportStudentsCsv(state.elements.examSelect.value));
  state.elements.exportQuestionsCsv.addEventListener("click", () => exportQuestionsCsv(state.elements.examSelect.value));
  state.elements.exportJson.addEventListener("click", () => exportExamJson(state.elements.examSelect.value));
  if (state.elements.importJsonButton) {
    state.elements.importJsonButton.addEventListener("click", importResultsFromJsonFile);
  }
  state.elements.closeStudentDialog.addEventListener("click", () => state.elements.studentDialog.close());
  if (state.elements.cancelEditDialog) {
    state.elements.cancelEditDialog.addEventListener("click", () => state.elements.editDialog.close());
  }
  if (state.elements.saveEditDialog) {
    state.elements.saveEditDialog.addEventListener("click", saveEditedStudentResult);
  }
  if (state.elements.examNameInput) {
    state.elements.examNameInput.addEventListener("change", () => saveExamName(state.elements.examSelect.value));
    state.elements.examNameInput.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        event.preventDefault();
        saveExamName(state.elements.examSelect.value);
        state.elements.examNameInput.blur();
      }
    });
  }

  if (state.elements.deleteExamFilter) {
    state.elements.deleteExamFilter.addEventListener("input", renderDeleteExamList);
  }
  if (state.elements.deleteExamsButton) {
    state.elements.deleteExamsButton.addEventListener("click", deleteSelectedExamGroups);
  }

  renderResultsPage();
  renderDeleteExamList();
}

function getQueryParam(name) {
  try {
    return new URLSearchParams(window.location.search).get(name) || "";
  } catch {
    return "";
  }
}

function renderResultsPage() {
  const examGroupId = state.elements.examSelect.value;
  const store = loadResultsStore();
  const exam = store.examGroups?.[examGroupId];
  if (!exam) {
    state.elements.examSummary.innerHTML = "<div>Avaliação não encontrada.</div>";
    state.elements.resultsTbody.innerHTML = "";
    state.elements.questionStats.innerHTML = "";
    return;
  }

  const results = Array.isArray(exam.results) ? [...exam.results] : [];
  const filtered = filterResultsByStudentName(results, state.elements.studentSearch.value);
  const sorted = sortResults(filtered, state.elements.sortSelect.value);

  if (state.elements.examNameInput) {
    state.elements.examNameInput.value = String(exam.examName || "");
  }
  renderExamSummary(exam, results);
  renderResultsTable(examGroupId, sorted);
  renderQuestionStats(exam, results);
}

function filterResultsByStudentName(results, query) {
  const q = String(query || "").trim().toLowerCase();
  if (!q) {
    return results;
  }
  return results.filter((item) => String(item.studentName || "").toLowerCase().includes(q));
}

function sortResults(results, mode) {
  const parseDate = (value) => (value ? Date.parse(value) : 0);
  const compareText = (a, b) => String(a || "").localeCompare(String(b || ""), "pt-BR");
  const compareNumber = (a, b) => (Number(a) || 0) - (Number(b) || 0);

  const list = [...results];
  switch (mode) {
    case "name_desc":
      list.sort((a, b) => compareText(b.studentName, a.studentName));
      break;
    case "score_asc":
      list.sort((a, b) => compareNumber(a.score, b.score));
      break;
    case "score_desc":
      list.sort((a, b) => compareNumber(b.score, a.score));
      break;
    case "correct_asc":
      list.sort((a, b) => compareNumber(a.correctCount, b.correctCount));
      break;
    case "correct_desc":
      list.sort((a, b) => compareNumber(b.correctCount, a.correctCount));
      break;
    case "date_asc":
      list.sort((a, b) => parseDate(a.correctedAt) - parseDate(b.correctedAt));
      break;
    case "date_desc":
      list.sort((a, b) => parseDate(b.correctedAt) - parseDate(a.correctedAt));
      break;
    case "name_asc":
    default:
      list.sort((a, b) => compareText(a.studentName, b.studentName));
      break;
  }
  return list;
}

function calculateMedian(values) {
  const nums = values.map((v) => Number(v) || 0).sort((a, b) => a - b);
  if (!nums.length) {
    return 0;
  }
  const mid = Math.floor(nums.length / 2);
  if (nums.length % 2 === 0) {
    return (nums[mid - 1] + nums[mid]) / 2;
  }
  return nums[mid];
}

function calculateStandardDeviation(values) {
  const nums = values.map((v) => Number(v) || 0);
  if (!nums.length) {
    return 0;
  }
  const mean = nums.reduce((s, v) => s + v, 0) / nums.length;
  const variance = nums.reduce((s, v) => s + ((v - mean) ** 2), 0) / nums.length;
  return Math.sqrt(variance);
}

function renderExamSummary(exam, results) {
  const scores = results.map((r) => Number(r.score) || 0);
  const totalStudents = results.length;
  const mean = totalStudents ? (scores.reduce((s, v) => s + v, 0) / totalStudents) : 0;
  const median = calculateMedian(scores);
  const max = totalStudents ? Math.max(...scores) : 0;
  const min = totalStudents ? Math.min(...scores) : 0;
  const stdev = calculateStandardDeviation(scores);
  const parsed = parseExamIdDate(exam.examGroupId);

  const createdLine = parsed.label
    ? `<div><strong>Data de elaboração:</strong> ${escapeHtml(parsed.label)}</div>`
    : "";

  state.elements.examSummary.innerHTML = `
    <div><strong>Avaliação:</strong> ${escapeHtml(exam.examName || exam.examGroupId || "")}</div>
    <div><strong>ID base:</strong> ${escapeHtml(exam.examGroupId || "")}</div>
    ${createdLine}
    <div><strong>Cartões/alunos corrigidos:</strong> ${totalStudents}</div>
    <div><strong>Média:</strong> ${mean.toFixed(2)}</div>
    <div><strong>Mediana:</strong> ${median.toFixed(2)}</div>
    <div><strong>Maior:</strong> ${max.toFixed(2)}</div>
    <div><strong>Menor:</strong> ${min.toFixed(2)}</div>
    <div><strong>Desvio padrão:</strong> ${stdev.toFixed(2)}</div>
  `;
}

function formatDateTime(iso) {
  try {
    const date = new Date(iso);
    return date.toLocaleString("pt-BR");
  } catch {
    return String(iso || "");
  }
}

function renderResultsTable(examGroupId, results) {
  const rows = results.map((item) => {
    const percent = Number(item.percentage || 0);
    const pillClass = percent >= 70 ? "ok" : (percent >= 40 ? "neutral" : "bad");
    return `
      <tr>
        <td>${escapeHtml(item.studentCardId ?? "-")}</td>
        <td class="muted-cell">${escapeHtml(item.fullExamId || "")}</td>
        <td>${escapeHtml(item.studentName || "")}</td>
        <td>${Number(item.score || 0).toFixed(2)}</td>
        <td>${Number(item.correctCount || 0)}</td>
        <td>${Number(item.wrongCount || 0)}</td>
        <td><span class="pill ${pillClass}">${percent.toFixed(2)}%</span></td>
        <td>${escapeHtml(formatDateTime(item.correctedAt))}</td>
        <td class="actions-cell">
          <button data-action="edit" data-result="${escapeHtml(item.resultId || "")}">Editar</button>
          <button data-action="details" data-result="${escapeHtml(item.resultId || "")}">Detalhes</button>
          <button class="secondary" data-action="delete" data-result="${escapeHtml(item.resultId || "")}">Excluir</button>
        </td>
      </tr>
    `;
  }).join("");

  state.elements.resultsTbody.innerHTML = rows || "<tr><td colspan=\"9\">Nenhum aluno encontrado.</td></tr>";

  for (const button of state.elements.resultsTbody.querySelectorAll("button[data-action]")) {
    button.addEventListener("click", () => {
      const action = button.dataset.action;
      const resultId = button.dataset.result;
      if (action === "details") {
        openStudentDetails(examGroupId, resultId);
      } else if (action === "edit") {
        openEditStudentResult(examGroupId, resultId);
      } else if (action === "delete") {
        deleteStudentResult(examGroupId, resultId);
      }
    });
  }
}

function openStudentDetails(examGroupId, resultId) {
  const store = loadResultsStore();
  const exam = store.examGroups?.[examGroupId];
  const row = (exam?.results || []).find((item) => String(item.resultId || "") === String(resultId || ""));
  if (!row) {
    return;
  }

  const lines = [];
  lines.push(`<div class="dialog-body">`);
  lines.push(`<h2 style="margin:0 0 8px;">${escapeHtml(row.studentName || "")}</h2>`);
  lines.push(`<div><strong>Cartão:</strong> ${escapeHtml(row.studentCardId ?? "-")}</div>`);
  lines.push(`<div><strong>ID completo:</strong> ${escapeHtml(row.fullExamId || "")}</div>`);
  lines.push(`<div><strong>Nota:</strong> ${(Number(row.score || 0)).toFixed(2)} (${(Number(row.percentage || 0)).toFixed(2)}%)</div>`);
  lines.push(`<div><strong>Acertos:</strong> ${Number(row.correctCount || 0)} | <strong>Erros:</strong> ${Number(row.wrongCount || 0)}</div>`);
  lines.push(`<div><strong>Correção:</strong> ${escapeHtml(formatDateTime(row.correctedAt))}</div>`);

  for (const qid of exam.questionIds || []) {
    const meta = row.questionMeta?.[qid] || {};
    const status = String(meta.status || "em_branco");
    const css = status === "correta" ? "correct" : (status === "em_branco" ? "blank" : (status === "anulada" ? "annulled" : "wrong"));
    const marcada = status === "correta" || status === "errada" ? (meta.marcada || "") : "";
    const marcadas = Array.isArray(meta.marcadas) && meta.marcadas.length ? meta.marcadas.join(",") : "";
    lines.push(`
      <div class="question-line ${css}">
        <div><strong>${escapeHtml(qid)}</strong></div>
        <div><strong>Marcada:</strong> ${escapeHtml(marcada || marcadas || "-")}</div>
        <div><strong>Correta:</strong> ${escapeHtml(meta.correta || "")}</div>
        <div><strong>Status:</strong> ${escapeHtml(status.replaceAll("_", " "))}</div>
      </div>
    `);
  }

  lines.push(`</div>`);
  state.elements.studentDialogBody.innerHTML = lines.join("");
  state.elements.studentDialog.showModal();
}

function deleteStudentResult(examGroupId, resultId) {
  const store = loadResultsStore();
  const exam = store.examGroups?.[examGroupId];
  const row = (exam?.results || []).find((item) => String(item.resultId || "") === String(resultId || ""));
  if (!row) {
    return;
  }

  const confirmed = window.confirm(`Excluir o resultado de ${row.studentName}?`);
  if (!confirmed) {
    return;
  }

  exam.results = (exam.results || []).filter((item) => String(item.resultId || "") !== String(resultId || ""));
  saveResultsStore(store);
  renderResultsPage();
}

function saveExamName(examGroupId) {
  if (!examGroupId) {
    return;
  }
  const nextName = String(state.elements.examNameInput?.value || "").trim();
  if (!nextName) {
    return;
  }
  const store = loadResultsStore();
  const exam = store.examGroups?.[examGroupId];
  if (!exam) {
    return;
  }
  exam.examName = nextName;
  saveResultsStore(store);

  try {
    const selector = `option[value="${CSS.escape(examGroupId)}"]`;
    const option = state.elements.examSelect?.querySelector(selector);
    if (option) {
      option.textContent = nextName;
    }
  } catch {
    // ignore
  }
}

function renderDeleteExamList() {
  if (!state.elements?.deleteExamList) {
    return;
  }

  const store = loadResultsStore();
  const groups = store.examGroups && typeof store.examGroups === "object" ? store.examGroups : {};
  const query = String(state.elements.deleteExamFilter?.value || "").trim().toLowerCase();

  const items = Object.values(groups)
    .filter((group) => group && typeof group === "object")
    .map((group) => ({
      examGroupId: String(group.examGroupId || ""),
      examName: String(group.examName || group.examGroupId || ""),
      createdAtFromId: String(group.createdAtFromId || ""),
      resultsCount: Array.isArray(group.results) ? group.results.length : 0,
    }))
    .filter((item) => Boolean(item.examGroupId))
    .filter((item) => {
      if (!query) return true;
      const haystack = `${item.examGroupId} ${item.examName}`.toLowerCase();
      return haystack.includes(query);
    })
    .sort((a, b) => a.examGroupId.localeCompare(b.examGroupId, "pt-BR"));

  if (!items.length) {
    state.elements.deleteExamList.innerHTML = "<div class=\"hint\">Nenhuma avaliação encontrada.</div>";
    return;
  }

  state.elements.deleteExamList.innerHTML = items
    .map((item) => {
      const created = item.createdAtFromId ? ` (${escapeHtml(formatDateTime(item.createdAtFromId))})` : "";
      return `
        <label class="delete-item">
          <input type="checkbox" class="delete-exam-checkbox" value="${escapeHtml(item.examGroupId)}">
          <div>
            <div><strong>${escapeHtml(item.examName)}</strong></div>
            <div class="muted-cell">${escapeHtml(item.examGroupId)}${created}</div>
            <div class="muted-cell">${escapeHtml(String(item.resultsCount))} aluno(s)</div>
          </div>
        </label>
      `;
    })
    .join("");
}

function deleteSelectedExamGroups() {
  if (!state.elements?.deleteExamList) {
    return;
  }

  const checked = Array.from(state.elements.deleteExamList.querySelectorAll("input.delete-exam-checkbox:checked"))
    .map((input) => String(input.value || "").trim())
    .filter(Boolean);

  if (!checked.length) {
    alert("Selecione ao menos uma avaliação para excluir.");
    return;
  }

  const store = loadResultsStore();
  if (!store.examGroups || typeof store.examGroups !== "object") {
    return;
  }

  const confirmed = window.confirm(
    `Excluir ${checked.length} avaliação(ões) e todos os resultados salvos?\n\nEssa ação não pode ser desfeita.`
  );
  if (!confirmed) {
    return;
  }

  for (const examGroupId of checked) {
    delete store.examGroups[examGroupId];
  }
  saveResultsStore(store);

  // Recarrega para atualizar select, tabela e estatísticas.
  window.location.href = "./resultados.html";
}

function recalculateQuestionStatus(marked, correct) {
  const value = String(marked || "").toUpperCase();
  const expected = String(correct || "").toUpperCase();
  if (value === "__MULTIPLE__") {
    return "multipla_marcacao";
  }
  if (value === "__AMBIGUOUS__") {
    return "ambigua";
  }
  if (value === "__ANNULLED__") {
    return "anulada";
  }
  if (!value) {
    return "em_branco";
  }
  return value === expected ? "correta" : "errada";
}

function recalculateStudentResult(row, exam) {
  const answerKey = exam.answerKey || {};
  const questionIds = Array.isArray(exam.questionIds) && exam.questionIds.length
    ? exam.questionIds
    : Object.keys(answerKey).sort((a, b) => a.localeCompare(b));

  const previousMeta = row.questionMeta && typeof row.questionMeta === "object" ? row.questionMeta : {};
  const answers = { ...(row.answers || {}) };
  const questionStatus = {};
  const questionMeta = {};

  let correctCount = 0;
  let blankCount = 0;
  let multipleCount = 0;
  let ambiguousCount = 0;
  let annulledCount = 0;

  for (const qid of questionIds) {
    const expected = String(answerKey[qid] || "").toUpperCase();
    const rawMarked = String(answers[qid] || "").trim().toUpperCase();
    const status = recalculateQuestionStatus(rawMarked, expected);

    let marcada = "";
    let marcadas = [];
    if (status === "correta" || status === "errada") {
      marcada = rawMarked;
      marcadas = rawMarked ? [rawMarked] : [];
    } else if (status === "multipla_marcacao") {
      marcadas = Array.isArray(previousMeta?.[qid]?.marcadas) ? previousMeta[qid].marcadas : [];
      marcada = "";
    } else if (status === "ambigua") {
      marcadas = Array.isArray(previousMeta?.[qid]?.marcadas) ? previousMeta[qid].marcadas : [];
      marcada = "";
    } else {
      marcada = "";
      marcadas = [];
    }

    // Sanitiza answers: guarda apenas A-E quando houver resposta única.
    answers[qid] = (status === "correta" || status === "errada") ? marcada : "";

    if (status === "correta") {
      correctCount += 1;
    } else if (status === "em_branco") {
      blankCount += 1;
    } else if (status === "multipla_marcacao") {
      multipleCount += 1;
    } else if (status === "ambigua") {
      ambiguousCount += 1;
    } else if (status === "anulada") {
      annulledCount += 1;
    }
    const numero = Number.parseInt(String(qid).replace(/^\D+/, ""), 10) || 0;
    questionStatus[qid] = status;
    questionMeta[qid] = {
      numero,
      correta: expected,
      marcada,
      marcadas,
      status,
    };
  }

  const total = questionIds.length;
  const totalScored = Math.max(0, total - annulledCount);
  const wrongCount = Math.max(0, totalScored - correctCount);
  const calculatedScore = totalScored ? Math.round((correctCount / totalScored) * 1000) / 100 : 0;
  const percentage = totalScored ? Math.round((correctCount / totalScored) * 10000) / 100 : 0;

  const manual = Boolean(row.manualScoreOverride);
  const score = manual ? (Number(row.score) || 0) : calculatedScore;

  return {
    ...row,
    answers,
    questionStatus,
    questionMeta,
    correctCount,
    wrongCount,
    blankCount,
    multipleCount,
    ambiguousCount,
    annulledCount,
    calculatedScore,
    percentage,
    score,
  };
}

function openEditStudentResult(examGroupId, resultId) {
  if (!state.elements?.editDialog || !state.elements?.editDialogBody) {
    alert("Editor não disponível nesta versão.");
    return;
  }
  const store = loadResultsStore();
  const exam = store.examGroups?.[examGroupId];
  const row = (exam?.results || []).find((item) => String(item.resultId || "") === String(resultId || ""));
  if (!exam || !row) {
    return;
  }

  state.editing = { examGroupId, resultId };

  const qids = Array.isArray(exam.questionIds) && exam.questionIds.length
    ? exam.questionIds
    : Object.keys(exam.answerKey || {}).sort((a, b) => a.localeCompare(b));

  const manual = Boolean(row.manualScoreOverride);
  const scoreValue = manual ? Number(row.score || 0) : Number(row.calculatedScore ?? row.score ?? 0);

  const lines = [];
  lines.push(`<div class="dialog-body">`);
  lines.push(`<h2 style="margin:0 0 8px;">Editar resultado</h2>`);
  lines.push(`<div><strong>Avaliação:</strong> ${escapeHtml(exam.examName || exam.examGroupId || "")}</div>`);
  lines.push(`<div><strong>ID base:</strong> ${escapeHtml(exam.examGroupId || "")}</div>`);
  lines.push(`<div><strong>ID completo:</strong> ${escapeHtml(row.fullExamId || "")}</div>`);
  lines.push(`<div><strong>Cartão:</strong> ${escapeHtml(row.studentCardId ?? "-")}</div>`);

  lines.push(`<label class="field"><span>Nome do aluno</span><input id="edit-student-name" type="text" value="${escapeHtml(row.studentName || "")}"></label>`);

  lines.push(`<div class="edit-score-row">`);
  lines.push(`<label class="debug-toggle" style="margin: 10px 0 0;"><input id="edit-manual-score" type="checkbox"${manual ? " checked" : ""}>Sobrescrever nota manualmente</label>`);
  lines.push(`<label class="field" style="margin-top: 10px;"><span>Nota</span><input id="edit-score" type="number" step="0.01" value="${escapeHtml(String(scoreValue))}"></label>`);
  lines.push(`<label class="field" style="margin-top: 10px;"><span>Motivo (opcional)</span><input id="edit-score-reason" type="text" value="${escapeHtml(String(row.manualScoreReason || ""))}"></label>`);
  lines.push(`</div>`);

  lines.push(`<div id="edit-summary" class="summary"></div>`);

  lines.push(`<h3 style="margin:14px 0 6px;">Respostas</h3>`);
  for (const qid of qids) {
    const expected = String(exam.answerKey?.[qid] || "").toUpperCase();
    const current = String(row.answers?.[qid] || "").toUpperCase();
    const status = recalculateQuestionStatus(current, expected);
    const css = status === "correta" ? "correct" : (status === "em_branco" ? "blank" : "wrong");
    lines.push(`
      <div class="question-line ${css}">
        <div><strong>${escapeHtml(qid)}</strong></div>
        <div><strong>Correta:</strong> ${escapeHtml(expected)}</div>
        <div>
          <strong>Marcada:</strong>
          <select class="edit-answer" data-qid="${escapeHtml(qid)}">
            <option value=""></option>
            ${["A","B","C","D","E"].map((l) => `<option value="${l}"${l === current ? " selected" : ""}>${l}</option>`).join("")}
          </select>
        </div>
        <div><strong>Status:</strong> <span class="edit-status" data-qid="${escapeHtml(qid)}">${escapeHtml(status.replaceAll("_", " "))}</span></div>
      </div>
    `);
  }
  lines.push(`</div>`);

  state.elements.editDialogBody.innerHTML = lines.join("");

  const manualToggle = state.elements.editDialogBody.querySelector("#edit-manual-score");
  const scoreInput = state.elements.editDialogBody.querySelector("#edit-score");
  const reasonInput = state.elements.editDialogBody.querySelector("#edit-score-reason");
  const nameInput = state.elements.editDialogBody.querySelector("#edit-student-name");

  const setManualEnabled = () => {
    const enabled = Boolean(manualToggle?.checked);
    if (scoreInput) {
      scoreInput.disabled = !enabled;
    }
    if (reasonInput) {
      reasonInput.disabled = !enabled;
    }
  };

  const updatePreview = () => {
    const temp = {
      ...row,
      studentName: String(nameInput?.value || row.studentName || "").trim(),
      studentId: normalizeName(String(nameInput?.value || row.studentName || "")),
      manualScoreOverride: Boolean(manualToggle?.checked),
      score: Number(scoreInput?.value || 0),
      manualScoreReason: String(reasonInput?.value || ""),
      answers: {},
    };
    for (const select of state.elements.editDialogBody.querySelectorAll("select.edit-answer")) {
      const qid = select.dataset.qid;
      temp.answers[qid] = String(select.value || "").toUpperCase();
    }
    const updated = recalculateStudentResult(temp, exam);

    const summary = state.elements.editDialogBody.querySelector("#edit-summary");
    if (summary) {
      summary.innerHTML = `
        <div><strong>Calculada:</strong> ${Number(updated.calculatedScore || 0).toFixed(2)} (${Number(updated.percentage || 0).toFixed(2)}%)</div>
        <div><strong>Acertos:</strong> ${Number(updated.correctCount || 0)} | <strong>Erros:</strong> ${Number(updated.wrongCount || 0)} | <strong>Em branco:</strong> ${Number(updated.blankCount || 0)}</div>
      `;
    }

    for (const qid of qids) {
      const expected = String(exam.answerKey?.[qid] || "").toUpperCase();
      const marked = String(updated.answers?.[qid] || "");
      const status = recalculateQuestionStatus(marked, expected);
      const statusEl = Array.from(state.elements.editDialogBody.querySelectorAll(".edit-status"))
        .find((element) => String(element.dataset.qid || "") === String(qid));
      if (statusEl) {
        statusEl.textContent = status.replaceAll("_", " ");
      }
    }

    if (!updated.manualScoreOverride && scoreInput) {
      scoreInput.value = String(Number(updated.calculatedScore || 0));
    }
  };

  setManualEnabled();
  updatePreview();

  manualToggle?.addEventListener("change", () => {
    setManualEnabled();
    updatePreview();
  });
  scoreInput?.addEventListener("input", updatePreview);
  reasonInput?.addEventListener("input", updatePreview);
  nameInput?.addEventListener("input", updatePreview);
  for (const select of state.elements.editDialogBody.querySelectorAll("select.edit-answer")) {
    select.addEventListener("change", updatePreview);
  }

  state.elements.editDialog.showModal();
}

function saveEditedStudentResult() {
  const editing = state.editing;
  if (!editing?.examGroupId || !editing?.resultId) {
    return;
  }
  const store = loadResultsStore();
  const exam = store.examGroups?.[editing.examGroupId];
  if (!exam) {
    return;
  }
  const index = (exam.results || []).findIndex((item) => String(item.resultId || "") === String(editing.resultId));
  if (index < 0) {
    return;
  }

  const body = state.elements.editDialogBody;
  const name = String(body?.querySelector("#edit-student-name")?.value || "").trim();
  if (!name) {
    alert("Informe o nome do aluno.");
    return;
  }
  const manual = Boolean(body?.querySelector("#edit-manual-score")?.checked);
  const score = Number(body?.querySelector("#edit-score")?.value || 0);
  const reason = String(body?.querySelector("#edit-score-reason")?.value || "");

  const previous = exam.results[index];
  const nextRow = {
    ...previous,
    studentName: name,
    studentId: normalizeName(name),
    manualScoreOverride: manual,
    score,
    manualScoreReason: manual ? reason : "",
    answers: {},
  };
  for (const select of body.querySelectorAll("select.edit-answer")) {
    const qid = select.dataset.qid;
    nextRow.answers[qid] = String(select.value || "").toUpperCase();
  }

  const recalculated = recalculateStudentResult(nextRow, exam);
  if (manual) {
    recalculated.score = Number(score || 0);
  } else {
    recalculated.score = Number(recalculated.calculatedScore || 0);
    recalculated.manualScoreReason = "";
  }

  exam.results[index] = recalculated;
  saveResultsStore(store);
  state.elements.editDialog.close();
  renderResultsPage();
}

function calculateQuestionStats(exam, results) {
  const stats = [];
  const questionIds = exam.questionIds || [];
  for (const qid of questionIds) {
    const correctLetter = String(exam.answerKey?.[qid] || "").toUpperCase();
    const counts = { A: 0, B: 0, C: 0, D: 0, E: 0, blank: 0, multiple: 0, ambiguous: 0 };
    let correct = 0;
    for (const row of results) {
      const meta = row.questionMeta?.[qid];
      if (!meta) {
        counts.blank += 1;
        continue;
      }
      const status = String(meta.status || "em_branco");
      if (status === "correta") {
        correct += 1;
      }
      if (status === "em_branco") {
        counts.blank += 1;
        continue;
      }
      if (status === "multipla_marcacao") {
        counts.multiple += 1;
      }
      if (status === "ambigua") {
        counts.ambiguous += 1;
      }
      const marcada = String(meta.marcada || "");
      if (marcada && counts[marcada] !== undefined) {
        counts[marcada] += 1;
      }
    }
    const total = results.length || 1;
    const correctPct = (correct / total) * 100;
    const difficulty = correctPct >= 70 ? "Fácil" : (correctPct >= 40 ? "Média" : "Difícil");
    const entries = Object.entries({ A: counts.A, B: counts.B, C: counts.C, D: counts.D, E: counts.E, "Em branco": counts.blank })
      .sort((a, b) => b[1] - a[1]);
    const mostMarked = entries[0]?.[0] || "-";
    stats.push({
      qid,
      correctLetter,
      correct,
      wrong: total - correct,
      blank: counts.blank,
      correctPct,
      difficulty,
      mostMarked,
      counts,
    });
  }
  return stats;
}

function renderQuestionStats(exam, results) {
  const stats = calculateQuestionStats(exam, results);
  if (!stats.length) {
    state.elements.questionStats.innerHTML = "<div>Nenhuma questão encontrada.</div>";
    return;
  }
  const cards = stats.map((item) => {
    const pct = item.correctPct;
    const pillClass = pct >= 70 ? "ok" : (pct >= 40 ? "neutral" : "bad");
    return `
      <div class="question-line ${pct >= 70 ? "correct" : (pct >= 40 ? "blank" : "wrong")}">
        <div><strong>${escapeHtml(item.qid)}</strong></div>
        <div><strong>Correta:</strong> ${escapeHtml(item.correctLetter)}</div>
        <div><strong>Acerto:</strong> <span class="pill ${pillClass}">${pct.toFixed(1)}%</span></div>
        <div><strong>Dificuldade:</strong> ${escapeHtml(item.difficulty)}</div>
      </div>
    `;
  }).join("");
  state.elements.questionStats.innerHTML = cards;
}

function downloadBlob(filename, content, mime) {
  const blob = new Blob([content], { type: mime });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function exportStudentsCsv(examGroupId) {
  const store = loadResultsStore();
  const exam = store.examGroups?.[examGroupId];
  if (!exam || !(exam.results || []).length) {
    alert("Não há resultados para exportar.");
    return;
  }
  const qids = exam.questionIds || [];
  const header = ["Cartão", "ID completo", "Aluno", "Nota", "Acertos", "Erros", "Percentual", "Data/Hora", ...qids];
  const rows = (exam.results || []).map((r) => {
    const base = [
      r.studentCardId ?? "",
      r.fullExamId ?? "",
      r.studentName,
      Number(r.score || 0).toFixed(2),
      r.correctCount,
      r.wrongCount,
      Number(r.percentage || 0).toFixed(2),
      r.correctedAt,
    ];
    const answers = qids.map((qid) => String(r.answers?.[qid] || ""));
    return [...base, ...answers];
  });
  const csv = [header, ...rows].map((line) => line.map(csvEscape).join(",")).join("\n");
  downloadBlob(`${examGroupId}-alunos.csv`, csv, "text/csv;charset=utf-8");
}

function exportQuestionsCsv(examGroupId) {
  const store = loadResultsStore();
  const exam = store.examGroups?.[examGroupId];
  if (!exam || !(exam.results || []).length) {
    alert("Não há resultados para exportar.");
    return;
  }
  const stats = calculateQuestionStats(exam, exam.results || []);
  const header = ["Questão", "Correta", "% Acerto", "Acertos", "Erros", "Em branco", "Mais marcada", "Dificuldade"];
  const rows = stats.map((s) => [
    s.qid,
    s.correctLetter,
    s.correctPct.toFixed(2),
    s.correct,
    s.wrong,
    s.blank,
    s.mostMarked,
    s.difficulty,
  ]);
  const csv = [header, ...rows].map((line) => line.map(csvEscape).join(",")).join("\n");
  downloadBlob(`${examGroupId}-questoes.csv`, csv, "text/csv;charset=utf-8");
}

function exportExamJson(examGroupId) {
  const store = loadResultsStore();
  const exam = store.examGroups?.[examGroupId];
  if (!exam) {
    alert("Prova não encontrada.");
    return;
  }
  downloadBlob(`${examGroupId}.json`, JSON.stringify(exam, null, 2), "application/json;charset=utf-8");
}

function csvEscape(value) {
  const text = String(value ?? "");
  const needs = /[",\n]/.test(text);
  const escaped = text.replaceAll("\"", "\"\"");
  return needs ? `"${escaped}"` : escaped;
}

async function importResultsFromJsonFile() {
  const file = state.elements.importJsonFile?.files?.[0];
  if (!file) {
    alert("Selecione um arquivo JSON para importar.");
    return;
  }

  let text = "";
  try {
    text = await file.text();
  } catch {
    alert("Não foi possível ler o arquivo.");
    return;
  }

  let parsed;
  try {
    parsed = JSON.parse(text);
  } catch {
    alert("JSON inválido.");
    return;
  }

  const incoming = migrateResultsStoreV3(parsed);
  if (!incoming || !incoming.examGroups || typeof incoming.examGroups !== "object") {
    alert("Arquivo não contém resultados compatíveis.");
    return;
  }

  const store = loadResultsStore();
  store.examGroups ||= {};

  let addedGroups = 0;
  let addedResults = 0;
  let duplicateResults = 0;

  // Primeiro passe: conta duplicidades.
  for (const [examGroupId, group] of Object.entries(incoming.examGroups)) {
    const targetGroup = store.examGroups[examGroupId];
    if (!targetGroup) {
      continue;
    }
    const existing = Array.isArray(targetGroup.results) ? targetGroup.results : [];
    const incomingResults = Array.isArray(group?.results) ? group.results : [];
    for (const row of incomingResults) {
      const fullExamId = String(row.fullExamId || row.examId || "");
      const studentCardId = row.studentCardId ? String(row.studentCardId) : null;
      const studentId = String(row.studentId || normalizeName(row.studentName || ""));
      const match = existing.some((e) =>
        (row.resultId && String(e.resultId || "") === String(row.resultId)) ||
        (fullExamId && String(e.fullExamId || e.examId || "") === fullExamId) ||
        (studentCardId && String(e.studentCardId || "") === studentCardId) ||
        (studentId && String(e.studentId || "") === studentId)
      );
      if (match) {
        duplicateResults += 1;
      }
    }
  }

  const replaceDuplicates = duplicateResults
    ? window.confirm(`Foram encontrados ${duplicateResults} resultados duplicados. Deseja substituir pelos do arquivo?`)
    : false;

  for (const [examGroupId, group] of Object.entries(incoming.examGroups)) {
    if (!group || typeof group !== "object") {
      continue;
    }
    const baseDate = parseExamIdDate(examGroupId);
    store.examGroups[examGroupId] ||= {
      examGroupId,
      examName: String(group.examName || examGroupId),
      createdAtFromId: String(group.createdAtFromId || baseDate.iso || ""),
      answerKey: group.answerKey && typeof group.answerKey === "object" ? group.answerKey : {},
      questionIds: Array.isArray(group.questionIds) ? group.questionIds : [],
      results: [],
      annulled: group.annulled && typeof group.annulled === "object" ? group.annulled : {},
    };

    const targetGroup = store.examGroups[examGroupId];
    if (targetGroup.results && !Array.isArray(targetGroup.results)) {
      targetGroup.results = [];
    }
    if (!Array.isArray(targetGroup.results)) {
      targetGroup.results = [];
    }

    if (!targetGroup.examName || targetGroup.examName === examGroupId) {
      targetGroup.examName = String(group.examName || examGroupId);
    }
    if (!targetGroup.createdAtFromId) {
      targetGroup.createdAtFromId = String(group.createdAtFromId || baseDate.iso || "");
    }
    if (group.answerKey && Object.keys(group.answerKey).length) {
      targetGroup.answerKey = group.answerKey;
    }
    if (Array.isArray(group.questionIds) && group.questionIds.length) {
      targetGroup.questionIds = group.questionIds;
    }

    if (targetGroup === store.examGroups[examGroupId] && targetGroup.results.length === 0 && group.results?.length) {
      // group criado agora
      addedGroups += 1;
    }

    const incomingResults = Array.isArray(group.results) ? group.results : [];
    for (const row of incomingResults) {
      const normalizedRow = {
        ...row,
        resultId: String(row.resultId || createResultId()),
        fullExamId: String(row.fullExamId || row.examId || ""),
        studentCardId: row.studentCardId ?? parseExamAndStudentCardId(String(row.fullExamId || row.examId || "")).studentCardId,
        studentName: String(row.studentName || ""),
        studentId: String(row.studentId || normalizeName(row.studentName || "")),
      };

      const fullExamId = normalizedRow.fullExamId;
      const studentCardId = normalizedRow.studentCardId ? String(normalizedRow.studentCardId) : null;
      const studentId = normalizedRow.studentId;

      const existingIndex = targetGroup.results.findIndex((e) =>
        (normalizedRow.resultId && String(e.resultId || "") === normalizedRow.resultId) ||
        (fullExamId && String(e.fullExamId || e.examId || "") === fullExamId) ||
        (studentCardId && String(e.studentCardId || "") === studentCardId) ||
        (studentId && String(e.studentId || "") === studentId)
      );

      if (existingIndex >= 0) {
        if (replaceDuplicates) {
          targetGroup.results[existingIndex] = { ...targetGroup.results[existingIndex], ...normalizedRow };
        }
        continue;
      }
      targetGroup.results.push(normalizedRow);
      addedResults += 1;
    }
  }

  saveResultsStore(store);
  alert(`Importação concluída. Avaliações adicionadas: ${addedGroups}. Resultados adicionados: ${addedResults}.`);
  window.location.reload();
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
  setWorkflowState("waitingUserToStartAnswerScan", "Câmera aberta. Alinhe o cartão e tire uma foto.");
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

  setStatus('Alinhe o celular com o cartão-resposta e toque em "Tirar foto do cartão".');
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
    const context = captureCanvasContext || captureCanvas.getContext("2d", { willReadFrequently: true });
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

function computeAnswerGuideBox(width, height) {
  const targetAspect = CARD_TARGET.width / CARD_TARGET.height;
  let boxW = Math.round(width * 0.92);
  let boxH = Math.round(boxW / Math.max(targetAspect, 0.01));
  if (boxH > Math.round(height * 0.78)) {
    boxH = Math.round(height * 0.78);
    boxW = Math.round(boxH * targetAspect);
  }
  const x = Math.round((width - boxW) / 2);
  const y = Math.round((height - boxH) / 2);
  return { x, y, w: boxW, h: boxH };
}

function pickMarkersNearGuideCorners(markers, width, height) {
  if (!Array.isArray(markers) || markers.length < 4) {
    return { ok: false, corners: [], rectScore: 0 };
  }

  const guide = computeAnswerGuideBox(width, height);
  const expected = [
    { x: guide.x, y: guide.y }, // tl
    { x: guide.x + guide.w, y: guide.y }, // tr
    { x: guide.x + guide.w, y: guide.y + guide.h }, // br
    { x: guide.x, y: guide.y + guide.h }, // bl
  ];

  const maxDist = Math.min(guide.w, guide.h) * 0.55;
  const chosen = [null, null, null, null];

  for (const marker of markers) {
    const center = marker.center;
    if (!center || !Number.isFinite(center.x) || !Number.isFinite(center.y)) {
      continue;
    }
    let bestIndex = -1;
    let bestDistance = Infinity;
    for (let i = 0; i < expected.length; i += 1) {
      const d = distance(center, expected[i]);
      if (d < bestDistance) {
        bestDistance = d;
        bestIndex = i;
      }
    }
    if (bestIndex < 0 || bestDistance > maxDist) {
      continue;
    }
    const current = chosen[bestIndex];
    if (!current) {
      chosen[bestIndex] = marker;
      continue;
    }
    const currentScore = Number(current.squareScore || 0) + Math.sqrt(Number(current.area || 0)) * 0.00025;
    const markerScore = Number(marker.squareScore || 0) + Math.sqrt(Number(marker.area || 0)) * 0.00025;
    if (markerScore > currentScore) {
      chosen[bestIndex] = marker;
    }
  }

  if (chosen.some((item) => !item)) {
    return { ok: false, corners: [], rectScore: 0 };
  }

  const ordered = orderCorners(chosen.map((m) => m.center));
  const validation = validateCornerGeometry(ordered, width, height);
  if (!validation.ok) {
    return { ok: false, corners: [], rectScore: validation.rectScore || 0 };
  }

  return { ok: true, corners: ordered, rectScore: validation.rectScore || 0 };
}

function cornersMeanDistance(cornersA, cornersB) {
  if (!Array.isArray(cornersA) || !Array.isArray(cornersB) || cornersA.length !== 4 || cornersB.length !== 4) {
    return Infinity;
  }
  let sum = 0;
  for (let i = 0; i < 4; i += 1) {
    sum += distance(cornersA[i], cornersB[i]);
  }
  return sum / 4;
}

function detectCardCornersRealtime(videoElement) {
  if (videoElement.videoWidth < 160 || videoElement.videoHeight < 160) {
    return null;
  }

  const scale = Math.min(1, REALTIME_DETECTION_MAX_WIDTH / Math.max(videoElement.videoWidth, 1));
  processingCanvas.width = Math.max(1, Math.round(videoElement.videoWidth * scale));
  processingCanvas.height = Math.max(1, Math.round(videoElement.videoHeight * scale));
  const captureContext = processingCanvasContext || processingCanvas.getContext("2d", { willReadFrequently: true });
  captureContext.drawImage(videoElement, 0, 0, processingCanvas.width, processingCanvas.height);

  const src = cv.imread(processingCanvas);
  const gray = new cv.Mat();
  const blur = new cv.Mat();
  const thresh = new cv.Mat();

  try {
    cv.cvtColor(src, gray, cv.COLOR_RGBA2GRAY);
    // Kernel menor para manter desempenho no mobile (suficiente para reduzir ruído leve).
    cv.GaussianBlur(gray, blur, new cv.Size(3, 3), 0, 0, cv.BORDER_DEFAULT);
    cv.threshold(blur, thresh, 0, 255, cv.THRESH_BINARY_INV + cv.THRESH_OTSU);
    if (state.opencvKernel) {
      cv.morphologyEx(thresh, thresh, cv.MORPH_CLOSE, state.opencvKernel);
    }

    let markers = findMarkersRealtime(thresh, processingCanvas.width, processingCanvas.height);
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
      markers = findMarkersRealtime(thresh, processingCanvas.width, processingCanvas.height);
    }

    if (markers.length < 4) {
      return null;
    }

    const selection = pickMarkersNearGuideCorners(markers, processingCanvas.width, processingCanvas.height);
    if (!selection.ok || selection.rectScore < MIN_CARD_RECTANGLE_SCORE) {
      return null;
    }

    return { corners: selection.corners, rectScore: selection.rectScore };
  } finally {
    src.delete();
    gray.delete();
    blur.delete();
    thresh.delete();
  }
}

function scaleCorners(corners, factor) {
  if (!Array.isArray(corners) || corners.length !== 4 || !Number.isFinite(factor) || factor <= 0) {
    return [];
  }
  return corners.map((point) => ({
    x: Number(point?.x || 0) * factor,
    y: Number(point?.y || 0) * factor,
  }));
}

function detectCardAnswersFromCanvasWithCorners(canvasElement, questionCount, corners) {
  if (!canvasElement?.width || !canvasElement?.height) {
    return null;
  }
  if (!Array.isArray(corners) || corners.length !== 4) {
    return null;
  }

  const src = cv.imread(canvasElement);
  const gray = new cv.Mat();

  try {
    cv.cvtColor(src, gray, cv.COLOR_RGBA2GRAY);
    cv.equalizeHist(gray, gray);

    const ordered = orderCorners(corners);
    const validation = validateCornerGeometry(ordered, canvasElement.width, canvasElement.height);
    if (!validation.ok) {
      return null;
    }

    const warped = perspectiveWarp(gray, ordered);
    const warpValidation = validateWarpedCardMarkers(warped);
    if (!warpValidation.ok) {
      warped.delete();
      return null;
    }

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
      corners: ordered,
      answers: reading.answers,
      rows: reading.rows,
      baseImageData,
      focusScore,
      diagnostics: {
        rectScore: validation.rectScore ?? null,
        warpValidation: {
          ok: warpValidation.ok,
          avg: warpValidation.avg,
          min: warpValidation.min,
          scores: warpValidation.scores,
        },
      },
    };
  } finally {
    src.delete();
    gray.delete();
  }
}

function startCardReading() {
  if (!state.proof) {
    setStatus("Nenhuma prova carregada.");
    return;
  }
  if (!state.stream || !state.elements.video?.srcObject) {
    setStatus("Abra a câmera antes de iniciar a leitura do cartão.");
    return;
  }
  if (!state.opencvReady) {
    setStatus("OpenCV ainda está carregando.");
    return;
  }

  if (state.cardCaptureMode === "diagnosticPhoto") {
    setStatus('Use o botão "Tirar foto do cartão (diagnóstico)" para fazer a leitura.');
    return;
  }

  setWorkflowState("answerScanning");
  updateCardControls();
  state.answers = {};
  state.stableSignature = "";
  state.stableCount = 0;
  state.lastDetection = null;
  state.lastResult = null;
  if (state.elements.resultPanel) {
    state.elements.resultPanel.classList.add("hidden");
    state.elements.resultPanel.innerHTML = "";
  }
  clearOverlay();
  stopLoops();

  setStatus("Alinhe os 4 quadradinhos pretos nos cantos do retângulo verde. Aguarde a captura automática.");

  let lastTimestamp = 0;
  const startAt = performance.now();
  let warned = false;
  let referenceCorners = null;
  let referenceRect = 0;

  const tick = (timestamp) => {
    if (state.workflow !== "answerScanning") {
      return;
    }
    if (!state.elements.video?.srcObject) {
      state.cardTimer = requestAnimationFrame(tick);
      return;
    }
    if (timestamp - lastTimestamp < CARD_FRAME_INTERVAL_MS) {
      state.cardTimer = requestAnimationFrame(tick);
      return;
    }
    lastTimestamp = timestamp;

    drawAnswerGuide();

    const found = detectCardCornersRealtime(state.elements.video);
    if (!found) {
      referenceCorners = null;
      referenceRect = 0;
      state.stableCount = 0;
      if (!warned && (timestamp - startAt) > 12000) {
        warned = true;
        setStatus("Não foi possível estabilizar a leitura. Tente o modo diagnóstico (por foto).");
      } else {
        setStatus("Procurando os 4 marcadores... mantenha o cartão inteiro visível e bem iluminado.");
      }
      state.cardTimer = requestAnimationFrame(tick);
      return;
    }

    drawOverlay(found.corners);

    const minDim = Math.min(processingCanvas.width, processingCanvas.height);
    const tolerance = Math.max(8, minDim * 0.018);

    if (referenceCorners) {
      const dist = cornersMeanDistance(referenceCorners, found.corners);
      const rectDelta = Math.abs((found.rectScore || 0) - (referenceRect || 0));
      if (dist <= tolerance && rectDelta <= 0.03) {
        state.stableCount += 1;
      } else {
        state.stableCount = 1;
        referenceCorners = found.corners;
        referenceRect = found.rectScore || 0;
      }
    } else {
      state.stableCount = 1;
      referenceCorners = found.corners;
      referenceRect = found.rectScore || 0;
    }

    setStatus(`Cartão detectado. Estabilizando... ${state.stableCount}/${REQUIRED_STABLE_FRAMES}`);

    if (state.stableCount >= REQUIRED_STABLE_FRAMES) {
      // Captura o melhor frame (mais recente) e roda o pipeline completo UMA vez.
      // Otimiza: reaproveita os cantos detectados no modo tempo real, evitando re-detectar marcadores no frame final.
      const questionCount = state.proof.quantidade_questoes;
      const videoWidth = state.elements.video.videoWidth || 0;
      const videoHeight = state.elements.video.videoHeight || 0;
      const finalScale = Math.min(1, CARD_PROCESSING_MAX_WIDTH / Math.max(videoWidth, 1));
      captureCanvas.width = Math.max(1, Math.round(videoWidth * finalScale));
      captureCanvas.height = Math.max(1, Math.round(videoHeight * finalScale));
      const finalContext = captureCanvasContext || captureCanvas.getContext("2d", { willReadFrequently: true });
      finalContext.drawImage(state.elements.video, 0, 0, captureCanvas.width, captureCanvas.height);

      const cornerFactor = captureCanvas.width / Math.max(1, processingCanvas.width);
      const scaledCorners = scaleCorners(found.corners, cornerFactor);

      let detection = detectCardAnswersFromCanvasWithCorners(captureCanvas, questionCount, scaledCorners);
      if (!detection) {
        // Fallback: tenta a rotina antiga (redetecta marcadores) se o reaproveitamento falhar.
        detection = detectCardAnswers(state.elements.video, questionCount);
      }

      if (!detection) {
        state.stableCount = 0;
        referenceCorners = null;
        referenceRect = 0;
        setStatus("Falha ao ler o cartão. Ajuste o enquadramento e tente novamente.");
        state.cardTimer = requestAnimationFrame(tick);
        return;
      }

      state.lastDetection = detection;
      state.answers = { ...detection.answers };
      updateRenderedAnswers();
      drawAlignedPreview(detection.baseImageData);

      if (detection.focusScore < MIN_FOCUS_SCORE) {
        state.stableCount = 0;
        referenceCorners = null;
        referenceRect = 0;
        setStatus("Imagem ainda sem nitidez suficiente. Afaste um pouco o celular e estabilize.");
        state.cardTimer = requestAnimationFrame(tick);
        return;
      }

      stopAnswerCamera();
      correctProof();
      setWorkflowState("answerDetected");
      updateCardControls();
      return;
    }

    state.cardTimer = requestAnimationFrame(tick);
  };

  state.cardTimer = requestAnimationFrame(tick);
}

function stopAnswerReading() {
  stopLoops();
  clearOverlay();
  const diagnostic = state.cardCaptureMode === "diagnosticPhoto";
  setWorkflowState(
    "waitingUserToStartAnswerScan",
    diagnostic ? "Leitura pausada." : "Leitura pausada. Toque em iniciar quando estiver alinhado.",
  );
  updateCardControls();
}

function disableAnswerCamera() {
  stopAnswerReading();
  stopAnswerCamera();
  const diagnostic = state.cardCaptureMode === "diagnosticPhoto";
  setWorkflowState(
    "waitingUserToStartAnswerScan",
    diagnostic
      ? 'Pronto. Toque em "Tirar foto do cartão (diagnóstico)" para capturar novamente.'
      : 'Pronto. Toque em "Iniciar leitura (tempo real)" para tentar novamente.',
  );
  updateCardControls();
}

function openCardPhotoPicker() {
  const input = state.elements?.cardPhotoInput;
  if (!input) {
    setStatus("Captura de foto indisponível neste navegador.");
    return;
  }
  try {
    input.value = "";
  } catch {
    // ignore
  }
  try {
    input.click();
  } catch (error) {
    setStatus(`Não foi possível abrir a câmera: ${error?.message || String(error)}`);
  }
}

async function enableAnswerCamera() {
  if (!state.proof) {
    setStatus("Nenhuma prova carregada.");
    return;
  }
  if (!state.opencvReady) {
    setStatus("OpenCV ainda está carregando.");
    return;
  }

  if (state.cardCaptureMode === "diagnosticPhoto") {
    setWorkflowState("waitingUserToStartAnswerScan");
    updateCardControls();
    openCardPhotoPicker();
    return;
  }

  await startCardCamera();
  if (!state.stream) {
    setWorkflowState("error", "Falha ao abrir a câmera.");
    updateCardControls();
    return;
  }
  startCardReading();
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
  const captureContext = processingCanvasContext || processingCanvas.getContext("2d", { willReadFrequently: true });
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
    let markers = [];
    const illuminationNormalized = normalizeIllumination(gray);
    try {
      markers = findMarkersMultiScale(illuminationNormalized || gray);
    } finally {
      try {
        illuminationNormalized?.delete?.();
      } catch {
        // ignore
      }
    }

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

    // Reforça seleção: calcula contraste local (preto dentro, branco fora) por candidato.
    // Ajuda a evitar falsos positivos (textos/bolhas/sombras) quando o cartão está distante.
    const minDim = Math.min(canvasElement.width, canvasElement.height);
    const maxOuter = clamp(Math.round(minDim * 0.14), 48, 220);
    markers = markers.map((marker) => {
      const area = Math.max(Number(marker.area || 0), 1);
      const side = Math.sqrt(area);
      const outerSize = clamp(Math.round(side * 2.2), 16, maxOuter);
      const innerSize = clamp(Math.round(side * 0.9), 12, Math.max(12, outerSize - 4));
      let contrastScore = 0;
      try {
        contrastScore = markerContrastScoreAt(gray, marker.center.x, marker.center.y, outerSize, innerSize);
      } catch {
        contrastScore = 0;
      }
      return { ...marker, contrastScore };
    });

    const selection = pickCornerMarkers(markers, canvasElement.width, canvasElement.height);
    if (!selection.ok) {
      return { ok: false, reason: selection.reason || "markers_not_found", diagnostics: selection };
    }
    // Não rejeitar cedo demais: em fotos distantes os marcadores ficam pequenos e o score cai.
    // Preferimos tentar o warp e validar depois (warpValidation) antes de pedir nova foto.
    if (selection.confidence < 0.045) {
      return { ok: false, reason: "low_confidence", diagnostics: selection };
    }

  if (state.debug) {
    try {
      const preview = new cv.Mat();
      cv.cvtColor(gray, preview, cv.COLOR_GRAY2RGBA);
      const imageData = matToImageData(preview);
      preview.delete();
      drawMarkerSelectionDebug(imageData, selection.markers);
    } catch {
      // ignore
    }
  }

  const corners = selection.corners;
  let warped = null;
  try {
    warped = perspectiveWarp(gray, corners);
  } catch {
    warped = null;
  }
  if (!warped) {
    return { ok: false, reason: "bad_geometry", diagnostics: selection };
  }

  const warpValidation = validateWarpedCardMarkers(warped);
  if (!warpValidation.ok) {
    warped.delete();
    return { ok: false, reason: "warp_invalid", diagnostics: { ...selection, warpValidation } };
  }

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
    inputSize: { width: canvasElement.width, height: canvasElement.height },
    diagnostics: {
      confidence: selection.confidence,
      rectScore: selection.rectScore ?? null,
      warpValidation: {
        ok: warpValidation.ok,
        avg: warpValidation.avg,
        min: warpValidation.min,
        scores: warpValidation.scores,
      },
    },
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

  if (markers.length === 4) {
    const corners = orderCorners(markers.map((marker) => marker.center));
    context.strokeStyle = "#ff3d00";
    context.lineWidth = 4;
    context.beginPath();
    context.moveTo(corners[0].x, corners[0].y);
    context.lineTo(corners[1].x, corners[1].y);
    context.lineTo(corners[2].x, corners[2].y);
    context.lineTo(corners[3].x, corners[3].y);
    context.closePath();
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
    const found = collectMarkerCandidatesV3(contours, width, height);
    if (found.length >= 4) {
      return found;
    }
    // Fallback: componentes conectados costuma ser mais tolerante quando o contorno fica "quebrado" por blur/ruído.
    const byComponents = collectMarkerCandidatesConnectedComponents(binaryMat, width, height);
    return byComponents.length >= 4 ? byComponents : found;
  } finally {
    contours.delete();
    hierarchy.delete();
  }
}

// Versão otimizada para o modo tempo real:
// - não executa busca combinatória (N^4) dos candidatos;
// - retorna uma lista maior de candidatos para seleção por canto (guia verde).
function findMarkersRealtime(binaryMat, width, height) {
  const contours = new cv.MatVector();
  const hierarchy = new cv.Mat();
  try {
    cv.findContours(binaryMat, contours, hierarchy, cv.RETR_EXTERNAL, cv.CHAIN_APPROX_SIMPLE);
    const found = collectMarkerCandidatesFast(contours, width, height);
    if (found.length >= 4) {
      return found;
    }
    // Evita custo alto do connectedComponents quando ainda ha poucos sinais (ex.: cartao longe/fora do quadro).
    if (found.length < 3) {
      return found;
    }
    const byComponents = collectMarkerCandidatesConnectedComponents(binaryMat, width, height);
    return byComponents.length >= 4 ? byComponents : found;
  } finally {
    contours.delete();
    hierarchy.delete();
  }
}

function collectMarkerCandidatesConnectedComponents(binaryMat, width, height) {
  if (typeof cv.connectedComponentsWithStats !== "function") {
    return [];
  }

  const labels = new cv.Mat();
  const stats = new cv.Mat();
  const centroids = new cv.Mat();

  try {
    const num = cv.connectedComponentsWithStats(binaryMat, labels, stats, centroids, 8, cv.CV_32S);
    if (!num || num < 2) {
      return [];
    }

    const frameArea = width * height;
    const minArea = frameArea * MARKER_MIN_AREA_RATIO;
    const maxArea = frameArea * MARKER_MAX_AREA_RATIO;
    const minDim = Math.min(width, height);
    const minSide = Math.max(5, Math.round(minDim * MARKER_MIN_SIDE_RATIO));
    const maxSide = Math.max(24, Math.round(minDim * MARKER_MAX_SIDE_RATIO));

    const candidates = [];
    for (let i = 1; i < num; i += 1) {
      const w = stats.intAt(i, 2);
      const h = stats.intAt(i, 3);
      const area = stats.intAt(i, 4);
      if (area < minArea || area > maxArea) {
        continue;
      }
      const rectWidth = Math.max(w, h);
      const rectHeight = Math.min(w, h);
      if (rectWidth < minSide || rectHeight < minSide || rectWidth > maxSide || rectHeight > maxSide) {
        continue;
      }

      const aspect = rectWidth / Math.max(rectHeight, 1);
      const fillRatio = area / Math.max(w * h, 1);
      if (aspect < 0.62 || aspect > 1.62 || fillRatio < 0.42) {
        continue;
      }

      const cx = centroids.doubleAt(i, 0);
      const cy = centroids.doubleAt(i, 1);
      const aspectSquare = rectHeight / Math.max(rectWidth, 1);
      const squareScore = clamp(aspectSquare * fillRatio, 0, 1);

      candidates.push({
        area,
        center: { x: cx, y: cy },
        squareScore,
      });
    }

    candidates.sort((a, b) => ((b.squareScore || 0) - (a.squareScore || 0)) || ((b.area || 0) - (a.area || 0)));

    const unique = [];
    const dedupeDistance = Math.max(14, Math.round(minDim * 0.02));
    for (const candidate of candidates) {
      const duplicate = unique.some((item) => distance(item.center, candidate.center) < dedupeDistance);
      if (!duplicate) {
        unique.push(candidate);
      }
      if (unique.length >= 20) {
        break;
      }
    }

    return unique;
  } catch {
    return [];
  } finally {
    labels.delete();
    stats.delete();
    centroids.delete();
  }
}

function findMarkersMultiScale(grayMat) {
  // Multiescala: inclui upscaling leve para ajudar quando os marcadores ficam pequenos na foto.
  const baseMaxDim = Math.max(grayMat.cols, grayMat.rows, 1);
  const MAX_WORK_DIM = 2500; // evita explosão de memória/CPU no mobile ao fazer upscaling
  const scales = [1.0, 1.75, 1.5, 1.25, 0.9, 0.75, 0.6, 0.45, 0.34, 0.28];
  const candidates = [];

  for (const scale of scales) {
    if (baseMaxDim * scale > MAX_WORK_DIM) {
      continue;
    }
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
        const interpolation = scale < 1 ? cv.INTER_AREA : cv.INTER_LINEAR;
        cv.resize(grayMat, scaled, size, 0, 0, interpolation);
      }

      // Blur adaptativo: blur grande pode "apagar" marcadores muito pequenos.
      const blurK = scale <= 0.6 ? 3 : 5;
      cv.GaussianBlur(scaled, blur, new cv.Size(blurK, blurK), 0, 0, cv.BORDER_DEFAULT);

      const thresholdAttempts = [
        () => cv.threshold(blur, thresh, 0, 255, cv.THRESH_BINARY_INV + cv.THRESH_OTSU),
        () => cv.adaptiveThreshold(blur, thresh, 255, cv.ADAPTIVE_THRESH_GAUSSIAN_C, cv.THRESH_BINARY_INV, 21, 5),
        () => cv.adaptiveThreshold(blur, thresh, 255, cv.ADAPTIVE_THRESH_GAUSSIAN_C, cv.THRESH_BINARY_INV, 31, 7),
        () => cv.adaptiveThreshold(blur, thresh, 255, cv.ADAPTIVE_THRESH_GAUSSIAN_C, cv.THRESH_BINARY_INV, 51, 9),
        () => cv.adaptiveThreshold(blur, thresh, 255, cv.ADAPTIVE_THRESH_GAUSSIAN_C, cv.THRESH_BINARY_INV, 71, 11),
      ];

      let markers = [];
      for (const applyThreshold of thresholdAttempts) {
        applyThreshold();
        // Morfologia por escala: close para preencher falhas do contorno do marcador.
        const morphK = scale >= 1.25 ? 5 : 3;
        const kernel = cv.getStructuringElement(cv.MORPH_RECT, new cv.Size(morphK, morphK));
        try {
          cv.morphologyEx(thresh, thresh, cv.MORPH_CLOSE, kernel);
        } finally {
          kernel.delete();
        }
        markers = findMarkersV2(thresh, scaled.cols, scaled.rows);
        if (markers.length < 4) {
          markers = findMarkers(thresh, scaled.cols, scaled.rows);
        }
        if (markers.length >= 4) {
          break;
        }
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
  candidates.sort((a, b) => ((b.squareScore || 0) - (a.squareScore || 0)) || ((b.area || 0) - (a.area || 0)));
  const unique = [];
  const dedupeDistance = Math.max(14, Math.round(Math.min(grayMat.cols, grayMat.rows) * 0.02));
  for (const candidate of candidates) {
    const duplicate = unique.some((item) => distance(item.center, candidate.center) < dedupeDistance);
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

  // Retorna candidatos (deduplicados) para que a seleção final garanta 1 marcador por canto.
  return unique;
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
  const minSide = Math.max(5, Math.round(minDim * MARKER_MIN_SIDE_RATIO));
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
      fillRatio < 0.44
    ) {
      contour.delete();
      continue;
    }

    // Evita confundir bolhas circulares preenchidas com marcadores quadrados.
    // (círculo tende a ter circularidade ~1; quadrado ~pi/4 ≈ 0.785)
    const perimeter = cv.arcLength(contour, true);
    const circularity = (4 * Math.PI * area) / Math.max(perimeter * perimeter, 1e-6);
    if (circularity > 0.945) {
      contour.delete();
      continue;
    }

    const hull = new cv.Mat();
    cv.convexHull(contour, hull, false, true);
    const hullArea = cv.contourArea(hull);
    const solidity = area / Math.max(hullArea, 1);
    hull.delete();

    if (solidity < 0.68) {
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
      circularity,
    });
  }

  candidates.sort((left, right) => {
    const leftScore = (left.squareScore || 0) * 10000 + Math.sqrt(left.area || 0);
    const rightScore = (right.squareScore || 0) * 10000 + Math.sqrt(right.area || 0);
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
           if (minDist < Math.min(width, height) * 0.045) {
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

// Coleta candidatos a marcadores com foco em performance (tempo real).
// Mesma ideia do V3, mas sem etapa combinatória e com heurísticas mais leves.
function collectMarkerCandidatesFast(contours, width, height) {
  const candidates = [];
  const frameArea = width * height;
  const minArea = frameArea * MARKER_MIN_AREA_RATIO;
  const maxArea = frameArea * MARKER_MAX_AREA_RATIO;
  const minDim = Math.min(width, height);
  const minSide = Math.max(5, Math.round(minDim * MARKER_MIN_SIDE_RATIO));
  const maxSide = Math.max(24, Math.round(minDim * MARKER_MAX_SIDE_RATIO));

  for (let index = 0; index < contours.size(); index += 1) {
    const contour = contours.get(index);
    const area = cv.contourArea(contour);
    if (area < minArea || area > maxArea) {
      contour.delete();
      continue;
    }

    const rect = cv.boundingRect(contour);
    contour.delete();

    const rectWidth = Math.max(rect.width, rect.height);
    const rectHeight = Math.min(rect.width, rect.height);
    if (rectWidth < minSide || rectHeight < minSide || rectWidth > maxSide || rectHeight > maxSide) {
      continue;
    }

    const aspect = rectWidth / Math.max(rectHeight, 1);
    const fillRatio = area / Math.max(rect.width * rect.height, 1);
    if (aspect < 0.66 || aspect > 1.52 || fillRatio < 0.44) {
      continue;
    }

    const center = { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 };
    const cornerScore = cornerProximityScore(center, width, height);
    const aspectSquare = rectHeight / Math.max(rectWidth, 1);
    const squareScore = clamp(aspectSquare * fillRatio, 0, 1);

    candidates.push({
      area,
      center,
      cornerScore,
      squareScore,
    });
  }

  candidates.sort((left, right) => {
    const leftScore = (left.squareScore || 0) * 10000 + Math.sqrt(left.area || 0);
    const rightScore = (right.squareScore || 0) * 10000 + Math.sqrt(right.area || 0);
    return rightScore - leftScore;
  });

  const unique = [];
  const dedupeDistance = Math.max(12, Math.round(minDim * 0.028));
  for (const candidate of candidates) {
    const duplicate = unique.some((item) => distance(item.center, candidate.center) < dedupeDistance);
    if (!duplicate) {
      unique.push(candidate);
    }
    if (unique.length >= 28) {
      break;
    }
  }

  return unique;
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
  const aspectError = Math.abs(quadAspect - MARKER_RECT_ASPECT) / Math.max(MARKER_RECT_ASPECT, 1e-6);
  const aspectScore = clamp(1 - aspectError, 0, 1);

  return normalizedArea * widthBalance * heightBalance * (0.35 + (aspectScore * 0.65));
}

function isConvexQuad(points) {
  if (!points || points.length !== 4) {
    return false;
  }
  const cross = (a, b, c) => ((b.x - a.x) * (c.y - a.y)) - ((b.y - a.y) * (c.x - a.x));
  const z1 = cross(points[0], points[1], points[2]);
  const z2 = cross(points[1], points[2], points[3]);
  const z3 = cross(points[2], points[3], points[0]);
  const z4 = cross(points[3], points[0], points[1]);
  const hasPos = z1 > 0 || z2 > 0 || z3 > 0 || z4 > 0;
  const hasNeg = z1 < 0 || z2 < 0 || z3 < 0 || z4 < 0;
  return !(hasPos && hasNeg);
}

function validateCornerGeometry(orderedCorners, frameWidth, frameHeight) {
  const reasons = [];
  if (!orderedCorners || orderedCorners.length !== 4) {
    return { ok: false, rectScore: 0, reasons: ["corners_invalid"] };
  }
  if (!isConvexQuad(orderedCorners)) {
    reasons.push("quadrilatero_nao_convexo");
  }

  const [tl, tr, br, bl] = orderedCorners;
  const widthTop = distance(tl, tr);
  const widthBottom = distance(bl, br);
  const heightLeft = distance(tl, bl);
  const heightRight = distance(tr, br);

  const widthBalance = 1 - Math.abs(widthTop - widthBottom) / Math.max(widthTop, widthBottom, 1);
  const heightBalance = 1 - Math.abs(heightLeft - heightRight) / Math.max(heightLeft, heightRight, 1);
  if (widthBalance < 0.55) {
    reasons.push("larguras_incompativeis");
  }
  if (heightBalance < 0.55) {
    reasons.push("alturas_incompativeis");
  }

  const rectScore = rectangleScore(orderedCorners, frameWidth, frameHeight);
  if (rectScore < MIN_CARD_RECTANGLE_SCORE) {
    reasons.push("area_ou_proporcao_ruim");
  }

  const minDist = Math.min(
    distance(tl, tr),
    distance(tr, br),
    distance(br, bl),
    distance(bl, tl),
  );
  if (minDist < Math.min(frameWidth, frameHeight) * 0.045) {
    reasons.push("marcadores_muito_proximos");
  }

  return { ok: reasons.length === 0, rectScore, reasons };
}

function markerQualityScore(marker, width, height) {
  const square = clamp(Number(marker.squareScore || 0), 0, 1);
  const contrast = clamp(Number(marker.contrastScore || 0), 0, 1);
  const area = clamp(Math.sqrt(Number(marker.area || 0)) / Math.sqrt(Math.max(width * height, 1)), 0, 1);
  const circularityRaw = Number(marker.circularity);
  const circularity = Number.isFinite(circularityRaw) ? clamp(circularityRaw, 0, 1) : 0.9;
  // Penaliza formatos muito "circulares" (bolhas preenchidas). Quadrado tende a ~0.78; círculo ~1.0.
  const circularityScore = clamp(1 - Math.max(0, circularity - 0.86) / 0.14, 0, 1);
  // Peso de área menor: permite marcadores menores (foto mais distante) permanecerem no pool.
  return (square * 0.64) + (contrast * 0.22) + (circularityScore * 0.11) + (area * 0.03);
}

function pickCornerMarkers(candidates, width, height) {
  if (!Array.isArray(candidates) || candidates.length < 4) {
    return { ok: false, reason: "markers_not_found", markers: [], corners: [], confidence: 0 };
  }

  const scored = candidates.map((marker, index) => ({
    marker,
    index,
    quality: markerQualityScore(marker, width, height),
  }));

  scored.sort((a, b) =>
    (b.quality - a.quality) ||
    ((b.marker.squareScore || 0) - (a.marker.squareScore || 0)) ||
    ((b.marker.area || 0) - (a.marker.area || 0))
  );

  // Pool limitado para manter desempenho no mobile.
  const pool = scored.slice(0, Math.min(18, scored.length));

  let best = null;
  const minDim = Math.min(width, height);
  const minDistLimit = minDim * 0.042;

  for (let a = 0; a < pool.length - 3; a += 1) {
    for (let b = a + 1; b < pool.length - 2; b += 1) {
      for (let c = b + 1; c < pool.length - 1; c += 1) {
        for (let d = c + 1; d < pool.length; d += 1) {
          const combo = [pool[a], pool[b], pool[c], pool[d]];
          const centers = combo.map((item) => item.marker.center);

          // Evita escolher quadrados muito próximos (ruído/texto).
          let minDist = Infinity;
          for (let i = 0; i < centers.length; i += 1) {
            for (let j = i + 1; j < centers.length; j += 1) {
              minDist = Math.min(minDist, distance(centers[i], centers[j]));
            }
          }
          if (minDist < minDistLimit) {
            continue;
          }

          // Marcadores reais tendem a ter áreas parecidas (mesma impressão). Evita combinações com tamanhos muito discrepantes.
          const areas = combo.map((item) => Number(item.marker.area || 0)).filter((value) => value > 0);
          if (areas.length === 4) {
            const minA = Math.min(areas[0], areas[1], areas[2], areas[3]);
            const maxA = Math.max(areas[0], areas[1], areas[2], areas[3]);
            if (minA > 0 && (maxA / minA) > 9) {
              continue;
            }
          }

          const ordered = orderCorners(centers);
          const validation = validateCornerGeometry(ordered, width, height);
          if (!validation.ok) {
            continue;
          }

          const qualityAvg = combo.reduce((sum, item) => sum + item.quality, 0) / 4;
          const score = validation.rectScore * (1 + (qualityAvg * 1.25));
          if (!best || score > best.score) {
            best = {
              score,
              rectScore: validation.rectScore,
              qualityAvg,
              markers: combo.map((item) => item.marker),
              corners: ordered,
            };
          }
        }
      }
    }
  }

  if (!best) {
    return { ok: false, reason: "bad_geometry", markers: [], corners: [], confidence: 0 };
  }

  const rectConfidence = clamp(
    (best.rectScore - MIN_CARD_RECTANGLE_SCORE) / Math.max(GOOD_CARD_RECTANGLE_SCORE - MIN_CARD_RECTANGLE_SCORE, 1e-6),
    0,
    1,
  );
  const confidence = clamp((rectConfidence * 0.75) + (best.qualityAvg * 0.25), 0, 1);

  return {
    ok: true,
    markers: best.markers,
    corners: best.corners,
    rectScore: best.rectScore,
    confidence,
    qualityAvg: best.qualityAvg,
  };
}

function warpCardImage(grayMat, orderedCorners) {
  const [tl, tr, br, bl] = orderedCorners;
  const widthTop = distance(tl, tr);
  const widthBottom = distance(bl, br);
  const heightLeft = distance(tl, bl);
  const heightRight = distance(tr, br);

  let targetWidth = Math.round(Math.max(widthTop, widthBottom));
  let targetHeight = Math.round(Math.max(heightLeft, heightRight));
  if (!Number.isFinite(targetWidth) || !Number.isFinite(targetHeight) || targetWidth < 10 || targetHeight < 10) {
    return null;
  }

  // Mantém proporção ao aplicar limites (evita "esticadas").
  const MIN_W = 520;
  const MIN_H = 760;
  const MAX_W = 1600;
  const MAX_H = 2400;

  let scale = 1;
  if (targetWidth < MIN_W || targetHeight < MIN_H) {
    scale = Math.max(MIN_W / targetWidth, MIN_H / targetHeight);
  } else if (targetWidth > MAX_W || targetHeight > MAX_H) {
    scale = Math.min(MAX_W / targetWidth, MAX_H / targetHeight);
  }
  if (scale !== 1) {
    targetWidth = Math.round(targetWidth * scale);
    targetHeight = Math.round(targetHeight * scale);
  }
  targetWidth = Math.max(320, Math.min(MAX_W, targetWidth));
  targetHeight = Math.max(480, Math.min(MAX_H, targetHeight));

  const targetAspect = targetWidth / Math.max(targetHeight, 1);
  const aspectError = Math.abs(targetAspect - MARKER_RECT_ASPECT) / Math.max(MARKER_RECT_ASPECT, 1e-6);
  if (aspectError > 0.45) {
    return null;
  }

  const srcTri = cv.matFromArray(4, 1, cv.CV_32FC2, [
    tl.x, tl.y,
    tr.x, tr.y,
    bl.x, bl.y,
    br.x, br.y,
  ]);
  const dstTri = cv.matFromArray(4, 1, cv.CV_32FC2, [
    0, 0,
    targetWidth - 1, 0,
    0, targetHeight - 1,
    targetWidth - 1, targetHeight - 1,
  ]);
  const matrix = cv.getPerspectiveTransform(srcTri, dstTri);
  const warped = new cv.Mat();
  cv.warpPerspective(
    grayMat,
    warped,
    matrix,
    new cv.Size(targetWidth, targetHeight),
    cv.INTER_LINEAR,
    cv.BORDER_CONSTANT,
    new cv.Scalar(255, 255, 255, 255),
  );
  srcTri.delete();
  dstTri.delete();
  matrix.delete();
  return warped;
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

function markerContrastScoreAt(grayMat, centerX, centerY, outerSize, innerSize) {
  const outerHalf = Math.floor(outerSize / 2);
  const innerHalf = Math.floor(innerSize / 2);

  const outerX = clamp(Math.round(centerX) - outerHalf, 0, grayMat.cols - outerSize);
  const outerY = clamp(Math.round(centerY) - outerHalf, 0, grayMat.rows - outerSize);
  const innerX = clamp(Math.round(centerX) - innerHalf, 0, grayMat.cols - innerSize);
  const innerY = clamp(Math.round(centerY) - innerHalf, 0, grayMat.rows - innerSize);

  const outerRect = new cv.Rect(outerX, outerY, outerSize, outerSize);
  const innerRect = new cv.Rect(innerX, innerY, innerSize, innerSize);
  const outerRoi = grayMat.roi(outerRect);
  const innerRoi = grayMat.roi(innerRect);

  try {
    const outerMean = cv.mean(outerRoi)[0];
    const innerMean = cv.mean(innerRoi)[0];

    const outer = Math.max(12, outerMean);
    const contrast = clamp((outerMean - innerMean) / outer, 0, 1);
    const darkness = clamp((255 - innerMean) / 255, 0, 1);
    return clamp((contrast * 0.78) + (darkness * 0.22), 0, 1);
  } finally {
    outerRoi.delete();
    innerRoi.delete();
  }
}

function validateWarpedCardMarkers(warpedGray) {
  const minDim = Math.min(warpedGray.cols, warpedGray.rows);
  const outerSize = clamp(Math.round(minDim * 0.055), 30, 80);
  const innerSize = clamp(Math.round(outerSize * 0.56), 18, outerSize - 4);

  const points = [
    { key: "tl", x: CARD_TARGET.leftMarkerX, y: CARD_TARGET.topMarkerY },
    { key: "tr", x: CARD_TARGET.rightMarkerX, y: CARD_TARGET.topMarkerY },
    { key: "br", x: CARD_TARGET.rightMarkerX, y: CARD_TARGET.bottomMarkerY },
    { key: "bl", x: CARD_TARGET.leftMarkerX, y: CARD_TARGET.bottomMarkerY },
  ];

  const scores = {};
  let sum = 0;
  let min = 1;
  for (const point of points) {
    const score = markerContrastScoreAt(warpedGray, point.x, point.y, outerSize, innerSize);
    scores[point.key] = score;
    sum += score;
    min = Math.min(min, score);
  }

  const avg = sum / 4;
  // Conservador: se algum marcador ficou "claro" demais após o warp, rejeita a leitura.
  const ok = avg >= 0.22 && min >= 0.14;
  return { ok, avg, min, scores, outerSize, innerSize };
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

function normalizeIllumination(grayMat) {
  // Correção simples de iluminação/sombras: normaliza pela "imagem de fundo" (blur grande).
  // Ajuda quando o marcador fica pequeno e o threshold falha por causa de sombra/gradiente.
  const minDim = Math.min(grayMat.cols, grayMat.rows);
  const kernel = clamp(Math.round(minDim / 22) * 2 + 1, 21, 71);
  const blurred = new cv.Mat();
  const normalized = new cv.Mat();
  try {
    cv.GaussianBlur(grayMat, blurred, new cv.Size(kernel, kernel), 0, 0, cv.BORDER_DEFAULT);
    cv.divide(grayMat, blurred, normalized, 255);
    cv.normalize(normalized, normalized, 0, 255, cv.NORM_MINMAX);
    return normalized;
  } catch {
    normalized.delete();
    return null;
  } finally {
    blurred.delete();
  }
}

function readAnswersFromWarped(warpedGray, questionCount, bounds) {
  const answers = {};
  const rows = [];
  const thresholded = new cv.Mat();
  cv.threshold(warpedGray, thresholded, 0, 255, cv.THRESH_BINARY_INV + cv.THRESH_OTSU);

  const leftX = Number(bounds?.leftX ?? CARD_TARGET.leftMarkerX);
  const rightX = Number(bounds?.rightX ?? CARD_TARGET.rightMarkerX);
  const topY = Number(bounds?.topY ?? CARD_TARGET.topMarkerY);
  const bottomY = Number(bounds?.bottomY ?? CARD_TARGET.bottomMarkerY);

  const stepX = (rightX - leftX) / 6;
  const stepY = (bottomY - topY) / (questionCount + 1);
  const bubbleRadius = Math.min(stepX, stepY) * 0.34;
  const sampleRadius = bubbleRadius * 0.58;

  for (let questionIndex = 1; questionIndex <= questionCount; questionIndex += 1) {
    const centerY = topY + (stepY * questionIndex);
    const scores = [];
    const centers = [];

    for (let optionIndex = 1; optionIndex <= 5; optionIndex += 1) {
      const centerX = leftX + (stepX * optionIndex);
      centers.push({ x: centerX, y: centerY });
      scores.push(sampleBubbleScore(thresholded, centerX, centerY, sampleRadius));
    }

    const markedIndices = resolveMarkedIndices(scores);
    const selectedIndex = markedIndices.length === 1 ? markedIndices[0] : -1;
    const letter = selectedIndex >= 0 ? ["A", "B", "C", "D", "E"][selectedIndex] : "";
    const status = markedIndices.length === 0
      ? "em_branco"
      : (markedIndices.length === 1 ? "respondida" : "multipla_marcacao");
    const markedLetters = markedIndices.map((index) => ["A", "B", "C", "D", "E"][index] || "?");

    if (letter) {
      answers[String(questionIndex)] = letter;
    }

    rows.push({
      questionNumber: questionIndex,
      centers,
      markedIndices,
      markedLetters,
      selectedIndex,
      bubbleRadius,
      status,
    });
  }

  thresholded.delete();
  return { answers, rows };
}

function readAnswersFromWarpedV2(warpedGray, questionCount, bounds) {
  const answers = {};
  const rows = [];

  const leftX = Number(bounds?.leftX ?? CARD_TARGET.leftMarkerX);
  const rightX = Number(bounds?.rightX ?? CARD_TARGET.rightMarkerX);
  const topY = Number(bounds?.topY ?? CARD_TARGET.topMarkerY);
  const bottomY = Number(bounds?.bottomY ?? CARD_TARGET.bottomMarkerY);

  const stepX = (rightX - leftX) / 6;
  const stepY = (bottomY - topY) / (questionCount + 1);
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
      const centerY = topY + (stepY * questionIndex);
      const centers = [];
      const scores = [];

      for (let optionIndex = 1; optionIndex <= 5; optionIndex += 1) {
        const centerX = leftX + (stepX * optionIndex);
        centers.push({ x: centerX, y: centerY });
        scores.push(sampleBubbleContrastScore(warpedGray, centerX, centerY, roiSize, innerMask, ringMask));
      }

      const classification = classifyBubbleScores(scores);
      const { markedIndices, selectedIndex, confidence, adaptiveThreshold, status } = classification;
      const letter = status === "respondida" && selectedIndex >= 0
        ? ["A", "B", "C", "D", "E"][selectedIndex]
        : "";
      if (letter) {
        answers[String(questionIndex)] = letter;
      }
      const markedLetters = markedIndices.map((index) => ["A", "B", "C", "D", "E"][index] || "?");

      rows.push({
        questionNumber: questionIndex,
        centers,
        scores,
        markedIndices,
        markedLetters,
        selectedIndex,
        bubbleRadius,
        confidence,
        adaptiveThreshold,
        status,
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

  const best = sorted[0] || { score: 0, index: -1 };
  const second = sorted[1] || { score: 0, index: -1 };

  const mean = scores.reduce((sum, value) => sum + value, 0) / Math.max(scores.length, 1);
  const variance = scores.reduce((sum, value) => sum + ((value - mean) ** 2), 0) / Math.max(scores.length, 1);
  const std = Math.sqrt(variance);

  const adaptiveThreshold = Math.max(BUBBLE_CONTRAST_MIN, mean + (std * 0.75));
  const markedIndices = scores
    .map((score, index) => ({ score, index }))
    .filter((item) => item.score >= adaptiveThreshold)
    .map((item) => item.index);

  // Regras pedagógicas: nunca "forçar" resposta.
  if (markedIndices.length === 0) {
    return { status: "em_branco", markedIndices: [], selectedIndex: -1, confidence: 0, adaptiveThreshold };
  }

  // Se mais de uma alternativa passou do limiar, é múltipla marcação (errada).
  if (markedIndices.length >= 2) {
    const confidence = clamp(best.score - second.score, 0, 1);
    return { status: "multipla_marcacao", markedIndices, selectedIndex: -1, confidence, adaptiveThreshold };
  }

  // Apenas uma acima do limiar: ainda pode ser ambígua se muito próxima da segunda.
  const relativeAmbiguous = second.score >= best.score * BUBBLE_AMBIGUOUS_RELATIVE;
  const gapAmbiguous = (best.score - second.score) <= BUBBLE_AMBIGUOUS_GAP;
  if (relativeAmbiguous || gapAmbiguous) {
    return {
      status: "ambigua",
      markedIndices: [best.index, second.index].filter((value) => value >= 0),
      selectedIndex: -1,
      confidence: clamp(best.score - second.score, 0, 1),
      adaptiveThreshold,
    };
  }

  return {
    status: "respondida",
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
  // Importante: não limpar nem redimensionar aqui.
  // O guia verde (drawAnswerGuide) já limpa e desenha o overlay a cada tick.
  const width = state.elements.overlayCanvas.width || state.elements.video.clientWidth || state.elements.video.videoWidth || 1;
  const height = state.elements.overlayCanvas.height || state.elements.video.clientHeight || state.elements.video.videoHeight || 1;
  const context = state.elements.overlayCanvas.getContext("2d");

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
    return [];
  }
  drawAlignedPreview(state.lastDetection.baseImageData);
  const context = state.elements.alignedCanvas.getContext("2d");
  context.lineWidth = 6;
  const yellowPoints = [];

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

    // Questão errada (inclui em branco/múltipla/ambígua): marca a alternativa correta com ponto amarelo.
    const expectedLetter = String(detail.correta || "").toUpperCase();
    const expectedIndex = ["A", "B", "C", "D", "E"].indexOf(expectedLetter);
    if (expectedIndex >= 0 && row.centers[expectedIndex]) {
      const center = row.centers[expectedIndex];
      const radius = Math.max(3, row.bubbleRadius * 0.22);
      drawFilledDot(context, center, radius, "#ffcc00");
      yellowPoints.push({ numero: detail.numero, correta: expectedLetter, x: center.x, y: center.y });
    }
  }

  return yellowPoints;
}

function drawCircle(context, center, radius, color) {
  context.strokeStyle = color;
  context.beginPath();
  context.arc(center.x, center.y, radius, 0, Math.PI * 2);
  context.stroke();
}

function drawFilledDot(context, center, radius, color) {
  context.fillStyle = color;
  context.strokeStyle = "rgba(30,30,30,0.75)";
  context.lineWidth = Math.max(1, Math.round(radius * 0.35));
  context.beginPath();
  context.arc(center.x, center.y, radius, 0, Math.PI * 2);
  context.fill();
  context.stroke();
}

function correctProof() {
  const result = correctProofLocally(state.proof, state.answers, state.lastDetection?.rows || []);
  state.lastResult = result;
  const yellowPoints = drawCorrectionPreview(result);
  state.elements.resultPanel.classList.remove("hidden");
  state.elements.resultPanel.classList.toggle("wrong", result.acertos !== result.total);
  state.elements.resultPanel.innerHTML = `
    <div><strong>Aluno:</strong> ${escapeHtml(result.aluno || "")}</div>
    <div><strong>Acertos:</strong> ${result.acertos} de ${result.total}</div>
    <div><strong>Nota:</strong> ${result.nota}</div>
    <div><strong>Em branco:</strong> ${result.brancos}</div>
    <div><strong>Múltiplas:</strong> ${result.multiplas}</div>
    <div><strong>Ambíguas:</strong> ${result.ambiguas}</div>
    <div><strong>Legenda:</strong> verde = correta, vermelho = errada.</div>
  `;

  if (state.debug) {
    const debugPayload = result.detalhes.map((item) => ({
      numero: item.numero,
      status: item.status,
      marcada: item.marcada,
      marcadas: item.marcadas,
      correta: item.correta,
      acertou: item.acertou,
      scores: item.scores,
      threshold: item.threshold,
    }));
    const debugBundle = {
      foto: state.lastPhotoMeta,
      leitura: {
        inputSize: state.lastDetection?.inputSize || null,
        focusScore: state.lastDetection?.focusScore ?? null,
        corners: state.lastDetection?.corners || null,
        diagnostics: state.lastDetection?.diagnostics || null,
      },
      pontos_corretos: yellowPoints,
      questoes: debugPayload,
    };
    const details = document.createElement("details");
    details.style.marginTop = "10px";
    const summary = document.createElement("summary");
    summary.textContent = "Debug (scores por questão)";
    const pre = document.createElement("pre");
    pre.style.whiteSpace = "pre-wrap";
    pre.style.fontSize = "0.85rem";
    pre.textContent = JSON.stringify(debugBundle, null, 2);
    details.appendChild(summary);
    details.appendChild(pre);
    state.elements.resultPanel.appendChild(details);
  }
  setStatus("Leitura concluída.");
}

function resetReadingUi() {
  state.answers = {};
  state.stableSignature = "";
  state.stableCount = 0;
  state.lastDetection = null;
  state.lastResult = null;
  state.lastPhotoMeta = null;
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

function handleCardDetectionFailure(detection) {
  const reason = detection?.reason || "unknown";
  if (reason === "markers_not_found") {
    setWorkflowState(
      "error",
      "Não foi possível identificar os quatro quadradinhos de alinhamento. Tire uma nova foto com o cartão inteiro visível e boa iluminação.",
    );
  } else if (reason === "low_confidence") {
    setWorkflowState(
      "error",
      "Os marcadores foram encontrados, mas com baixa confiança. Certifique-se de que os 4 quadradinhos estão visíveis (sem sombras) e tire uma nova foto.",
    );
  } else if (reason === "warp_invalid") {
    setWorkflowState(
      "error",
      "Os marcadores foram detectados, mas a correção de perspectiva não ficou confiável. Certifique-se de que os 4 quadradinhos estejam bem visíveis e tire uma nova foto com boa iluminação.",
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
}

async function processPhotoFromFile(file) {
  if (!state.proof) {
    setStatus("Nenhuma prova carregada.");
    return;
  }
  if (!state.opencvReady) {
    setStatus("OpenCV ainda está carregando.");
    return;
  }
  if (!file) {
    setStatus("Nenhuma foto selecionada.");
    return;
  }

  resetReadingUi();
  stopAnswerCamera();
  setWorkflowState("answerScanning", "Processando foto…");
  updateCardControls();

  let bitmap;
  try {
    try {
      bitmap = await createImageBitmap(file, { imageOrientation: "from-image" });
    } catch {
      bitmap = await createImageBitmap(file);
    }
  } catch (error) {
    setWorkflowState("error", `Não foi possível abrir a foto selecionada: ${error?.message || String(error)}`);
    updateCardControls();
    return;
  }

  let detection;
  try {
    state.lastPhotoMeta = {
      original: { width: bitmap.width, height: bitmap.height },
      attempts: [],
    };

    const drawBitmapToCanvas = (maxDimension, rotationSteps = 0) => {
      const maxDim = Math.max(bitmap.width, bitmap.height, 1);
      const safeMax = clamp(maxDimension, 720, 3600);
      const scale = Math.min(1, safeMax / maxDim);
      const targetWidth = Math.max(1, Math.round(bitmap.width * scale));
      const targetHeight = Math.max(1, Math.round(bitmap.height * scale));
      const rotation = ((rotationSteps % 4) + 4) % 4;

      processingCanvas.width = rotation % 2 === 0 ? targetWidth : targetHeight;
      processingCanvas.height = rotation % 2 === 0 ? targetHeight : targetWidth;
      const ctx = processingCanvasContext || processingCanvas.getContext("2d", { willReadFrequently: true });
      ctx.imageSmoothingEnabled = true;
      try {
        ctx.imageSmoothingQuality = "high";
      } catch {
        // ignore
      }
      ctx.setTransform(1, 0, 0, 1, 0, 0);
      ctx.clearRect(0, 0, processingCanvas.width, processingCanvas.height);
      ctx.save();
      if (rotation === 0) {
        ctx.drawImage(bitmap, 0, 0, targetWidth, targetHeight);
      } else if (rotation === 1) {
        // 90° CW
        ctx.translate(processingCanvas.width, 0);
        ctx.rotate(Math.PI / 2);
        ctx.drawImage(bitmap, 0, 0, targetWidth, targetHeight);
      } else if (rotation === 2) {
        // 180°
        ctx.translate(processingCanvas.width, processingCanvas.height);
        ctx.rotate(Math.PI);
        ctx.drawImage(bitmap, 0, 0, targetWidth, targetHeight);
      } else {
        // 270° CW
        ctx.translate(0, processingCanvas.height);
        ctx.rotate(-Math.PI / 2);
        ctx.drawImage(bitmap, 0, 0, targetWidth, targetHeight);
      }
      ctx.restore();

      return { scale, rotation, targetWidth: processingCanvas.width, targetHeight: processingCanvas.height };
    };

    const shouldRetryHigherRes = (result) => {
      const reason = result?.reason || "unknown";
      return reason === "markers_not_found" || reason === "low_confidence" || reason === "bad_geometry" || reason === "warp_invalid";
    };

    try {
      const questionCount = state.proof.quantidade_questoes;
      const rotations = [0, 1, 2, 3];

      const runAttempt = (maxDimension, rotationSteps) => {
        const meta = drawBitmapToCanvas(maxDimension, rotationSteps);
        const result = detectCardAnswersFromCanvas(processingCanvas, questionCount);
        state.lastPhotoMeta.attempts.push({
          maxDimension,
          rotationSteps: meta.rotation,
          canvasWidth: meta.targetWidth,
          canvasHeight: meta.targetHeight,
          scale: meta.scale,
          ok: result?.ok === true,
          reason: result?.ok === false ? (result.reason || "unknown") : null,
        });
        return result;
      };

      // 1) Tentativa padrão (mais rápida): tenta rotações se necessário (alguns navegadores ignoram EXIF).
      for (const rotationSteps of rotations) {
        detection = runAttempt(PHOTO_PROCESSING_MAX_WIDTH, rotationSteps);
        if (detection?.ok === true || !shouldRetryHigherRes(detection)) {
          break;
        }
      }

      // 2) Fallback: mais resolução para fotos distantes (marcadores pequenos).
      if (detection?.ok !== true && shouldRetryHigherRes(detection)) {
        const maxDim = Math.max(bitmap.width, bitmap.height, 1);
        if (PHOTO_PROCESSING_FALLBACK_MAX_WIDTH > PHOTO_PROCESSING_MAX_WIDTH && maxDim > PHOTO_PROCESSING_MAX_WIDTH * 1.05) {
          for (const rotationSteps of rotations) {
            detection = runAttempt(PHOTO_PROCESSING_FALLBACK_MAX_WIDTH, rotationSteps);
            if (detection?.ok === true || !shouldRetryHigherRes(detection)) {
              break;
            }
          }
        }
      }
    } finally {
      try {
        bitmap.close?.();
      } catch {
        // ignore
      }
    }
  } catch {
    detection = { ok: false, reason: "unknown" };
  }

  if (!detection || detection.ok === false) {
    handleCardDetectionFailure(detection);
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

async function captureAndProcessPhoto() {
  // Mantido por compatibilidade (card.html antigo).
  // No fluxo atual, a captura é feita via câmera nativa pelo input file.
  await enableAnswerCamera();
}

function retakePhoto() {
  resetReadingUi();
  setWorkflowState("waitingUserToStartAnswerScan", 'Toque em "Tirar foto do cartão" para capturar novamente.');
  updateCardControls();
  enableAnswerCamera();
}

function confirmReading() {
  if (!state.proof || !state.lastResult) {
    setStatus("Nenhum resultado para confirmar.");
    return;
  }

  // Garante que o nome do aluno exista (para não sobrescrever como "Aluno").
  if (!String(state.proof.aluno || "").trim()) {
    const typed = window.prompt("Informe o nome do aluno:", "");
    if (!typed || !String(typed).trim()) {
      setStatus("Informe o nome do aluno para salvar o resultado.");
      return;
    }
    state.proof.aluno = String(typed).trim();
    renderProofSummary();
  }

  try {
    saveStudentResult({
      proof: state.proof,
      result: state.lastResult,
    });
  } catch (error) {
    setStatus(error?.message || String(error));
    return;
  }

  setStatus("Resultado salvo com sucesso.");
}

function cancelReading() {
  resetReadingUi();
  stopAnswerCamera();
  setWorkflowState("waitingUserToStartAnswerScan", 'Leitura cancelada. Toque em "Tirar foto do cartão" para tentar novamente.');
  updateCardControls();
}

function loadResultsStore() {
  try {
    const raw = localStorage.getItem(RESULTS_STORAGE_KEY);
    if (!raw) {
      return { version: 3, examGroups: {} };
    }
    const parsed = JSON.parse(raw);
    return migrateResultsStoreV3(parsed);
  } catch {
    return { version: 3, examGroups: {} };
  }
}

function saveResultsStore(store) {
  localStorage.setItem(RESULTS_STORAGE_KEY, JSON.stringify(store));
}

function migrateResultsStore(parsed) {
  const empty = { version: 2, exams: {} };
  if (!parsed || typeof parsed !== "object") {
    return empty;
  }

  // Formato atual: { exams: { [examId]: exam } }
  if (parsed.exams && typeof parsed.exams === "object") {
    const next = { version: Number(parsed.version) || 2, exams: {} };
    for (const [examId, exam] of Object.entries(parsed.exams)) {
      if (!exam || typeof exam !== "object") {
        continue;
      }
      const parsedDate = parseExamIdDate(examId);
      const normalized = {
        examId,
        examName: String(exam.examName || examId),
        createdAtFromId: String(exam.createdAtFromId || parsedDate.iso || ""),
        answerKey: exam.answerKey && typeof exam.answerKey === "object" ? exam.answerKey : {},
        questionIds: Array.isArray(exam.questionIds) ? exam.questionIds : [],
        results: Array.isArray(exam.results) ? exam.results : [],
        annulled: exam.annulled && typeof exam.annulled === "object" ? exam.annulled : {},
      };
      normalized.results = normalized.results.map((row) => ({
        ...row,
        studentId: String(row.studentId || normalizeName(row.studentName || "")),
      }));
      next.exams[examId] = normalized;
    }
    return next;
  }

  // Migração: exame salvo diretamente.
  if (typeof parsed.examId === "string") {
    const examId = parsed.examId;
    const parsedDate = parseExamIdDate(examId);
    const exam = {
      examId,
      examName: String(parsed.examName || examId),
      createdAtFromId: String(parsed.createdAtFromId || parsedDate.iso || ""),
      answerKey: parsed.answerKey && typeof parsed.answerKey === "object" ? parsed.answerKey : {},
      questionIds: Array.isArray(parsed.questionIds) ? parsed.questionIds : [],
      results: Array.isArray(parsed.results) ? parsed.results : [],
      annulled: parsed.annulled && typeof parsed.annulled === "object" ? parsed.annulled : {},
    };
    exam.results = exam.results.map((row) => ({
      ...row,
      studentId: String(row.studentId || normalizeName(row.studentName || "")),
    }));
    return { version: 2, exams: { [examId]: exam } };
  }

  // Migração: um único resultado.
  if (parsed.studentName && parsed.examId) {
    const examId = String(parsed.examId);
    const parsedDate = parseExamIdDate(examId);
    const row = {
      ...parsed,
      studentId: String(parsed.studentId || normalizeName(parsed.studentName || "")),
    };
    return {
      version: 2,
      exams: {
        [examId]: {
          examId,
          examName: String(parsed.examName || examId),
          createdAtFromId: String(parsed.createdAtFromId || parsedDate.iso || ""),
          answerKey: parsed.answerKey && typeof parsed.answerKey === "object" ? parsed.answerKey : {},
          questionIds: Array.isArray(parsed.questionIds) ? parsed.questionIds : [],
          results: [row],
          annulled: {},
        },
      },
    };
  }

  return empty;
}

function migrateResultsStoreV3(parsed) {
  const empty = { version: 3, examGroups: {} };
  if (!parsed || typeof parsed !== "object") {
    return empty;
  }

  // Se já estiver em v3, normaliza e garante campos.
  if (parsed.examGroups && typeof parsed.examGroups === "object") {
    const next = { version: Number(parsed.version) || 3, examGroups: {} };
    for (const [examGroupId, group] of Object.entries(parsed.examGroups)) {
      if (!group || typeof group !== "object") {
        continue;
      }
      const parsedDate = parseExamIdDate(examGroupId);
      const normalized = {
        examGroupId,
        examName: String(group.examName || examGroupId),
        createdAtFromId: String(group.createdAtFromId || parsedDate.iso || ""),
        answerKey: group.answerKey && typeof group.answerKey === "object" ? group.answerKey : {},
        questionIds: Array.isArray(group.questionIds) ? group.questionIds : [],
        results: Array.isArray(group.results) ? group.results : [],
        annulled: group.annulled && typeof group.annulled === "object" ? group.annulled : {},
      };
      normalized.results = normalized.results.map((row) => {
        const fullExamId = String(row.fullExamId || row.examId || examGroupId);
        const parsedIds = parseExamAndStudentCardId(fullExamId);
        return {
          ...row,
          resultId: String(row.resultId || createResultId()),
          fullExamId,
          studentCardId: row.studentCardId ?? parsedIds.studentCardId,
          studentId: String(row.studentId || normalizeName(row.studentName || "")),
        };
      });
      next.examGroups[examGroupId] = normalized;
    }
    return next;
  }

  // Converte v2 (exams) para v3 (examGroups).
  const v2 = migrateResultsStore(parsed);
  const next = { version: 3, examGroups: {} };
  const ensureGroup = (examGroupId, base) => {
    next.examGroups[examGroupId] ||= {
      examGroupId,
      examName: base.examName || examGroupId,
      createdAtFromId: base.createdAtFromId || "",
      answerKey: base.answerKey || {},
      questionIds: base.questionIds || [],
      results: [],
      annulled: base.annulled || {},
    };
    if (base.answerKey && Object.keys(base.answerKey).length) {
      next.examGroups[examGroupId].answerKey = base.answerKey;
    }
    if (Array.isArray(base.questionIds) && base.questionIds.length) {
      next.examGroups[examGroupId].questionIds = base.questionIds;
    }
    if (base.examName && (!next.examGroups[examGroupId].examName || next.examGroups[examGroupId].examName === examGroupId)) {
      next.examGroups[examGroupId].examName = base.examName;
    }
    if (base.createdAtFromId && !next.examGroups[examGroupId].createdAtFromId) {
      next.examGroups[examGroupId].createdAtFromId = base.createdAtFromId;
    }
    return next.examGroups[examGroupId];
  };

  for (const [fullExamId, exam] of Object.entries(v2.exams || {})) {
    if (!exam || typeof exam !== "object") {
      continue;
    }
    const ids = parseExamAndStudentCardId(fullExamId);
    const parsedDate = parseExamIdDate(ids.examGroupId);
    const group = ensureGroup(ids.examGroupId, {
      examName: String(exam.examName || ids.examGroupId),
      createdAtFromId: String(exam.createdAtFromId || parsedDate.iso || ""),
      answerKey: exam.answerKey && typeof exam.answerKey === "object" ? exam.answerKey : {},
      questionIds: Array.isArray(exam.questionIds) ? exam.questionIds : [],
      annulled: exam.annulled && typeof exam.annulled === "object" ? exam.annulled : {},
    });
    for (const row of Array.isArray(exam.results) ? exam.results : []) {
      const rowFull = String(row.fullExamId || row.examId || fullExamId);
      const rowIds = parseExamAndStudentCardId(rowFull);
      group.results.push({
        ...row,
        resultId: String(row.resultId || createResultId()),
        fullExamId: rowFull,
        studentCardId: row.studentCardId ?? rowIds.studentCardId,
        studentId: String(row.studentId || normalizeName(row.studentName || "")),
      });
    }
  }

  return next;
}

function questionIdForNumber(number) {
  return `Q${String(number).padStart(3, "0")}`;
}

function toIsoNow() {
  return new Date().toISOString();
}

function saveStudentResult({ proof, result }) {
  const fullExamId = String(proof.id_prova || "").trim() || "prova";
  const parsedIds = parseExamAndStudentCardId(fullExamId);
  const examGroupId = parsedIds.examGroupId || fullExamId;
  const studentCardId = parsedIds.studentCardId;
  const examName = String(proof.nome_prova || parsedIds.examGroupId || "").trim() || examGroupId;
  const studentName = String(proof.aluno || "").trim();
  if (!studentName) {
    throw new Error("Informe o nome do aluno para salvar o resultado.");
  }
  const studentId = String(proof.id_aluno || proof.studentId || "").trim() || normalizeName(studentName);

  const answerKey = {};
  const questionIds = [];
  for (const question of proof.questoes || []) {
    const qid = questionIdForNumber(question.numero);
    questionIds.push(qid);
    answerKey[qid] = String(question.correta_letra || "").toUpperCase();
  }

  const answers = {};
  const questionStatus = {};
  const questionMeta = {};

  for (const detail of result.detalhes || []) {
    const qid = questionIdForNumber(detail.numero);
    const status = String(detail.status || "em_branco");
    const expected = String(detail.correta || "").toUpperCase();
    const marked = String(detail.marcada || "").toUpperCase();

    const normalizedStatus = status === "respondida"
      ? (detail.acertou ? "correta" : "errada")
      : (status === "em_branco" ? "em_branco" : (status === "multipla_marcacao" ? "multipla_marcacao" : "ambigua"));

    answers[qid] = status === "respondida" ? marked : "";
    questionStatus[qid] = normalizedStatus;
    questionMeta[qid] = {
      numero: detail.numero,
      correta: expected,
      marcada: marked,
      marcadas: detail.marcadas || [],
      status: normalizedStatus,
    };
  }

  const store = loadResultsStore();
  store.examGroups ||= {};
  const parsedDate = parseExamIdDate(examGroupId);
  store.examGroups[examGroupId] ||= {
    examGroupId,
    examName,
    createdAtFromId: parsedDate.iso || "",
    answerKey,
    questionIds,
    results: [],
    annulled: {},
  };

  // Atualiza gabarito se já existir (mantém o mais recente).
  store.examGroups[examGroupId].answerKey = answerKey;
  store.examGroups[examGroupId].questionIds = questionIds;
  store.examGroups[examGroupId].examName = store.examGroups[examGroupId].examName || examName;
  store.examGroups[examGroupId].createdAtFromId = store.examGroups[examGroupId].createdAtFromId || parsedDate.iso || "";
  if (!Array.isArray(store.examGroups[examGroupId].results)) {
    store.examGroups[examGroupId].results = [];
  }

  const correctedAt = toIsoNow();
  const total = Number(result.total || 0);
  const correctCount = Number(result.acertos || 0);
  const wrongCount = Math.max(0, total - correctCount);
  const percentage = total ? Math.round((correctCount / total) * 10000) / 100 : 0;

  const calculatedScore = Number(result.nota || 0);
  const row = {
    resultId: createResultId(),
    examGroupId,
    fullExamId,
    studentCardId,
    studentId,
    studentName,
    examId: fullExamId,
    correctedAt,
    score: calculatedScore,
    calculatedScore,
    manualScoreOverride: false,
    correctCount,
    wrongCount,
    blankCount: Number(result.brancos || 0),
    multipleCount: Number(result.multiplas || 0),
    ambiguousCount: Number(result.ambiguas || 0),
    percentage,
    answers,
    answerKey,
    questionStatus,
    questionMeta,
  };

  const beforeCount = (store.examGroups[examGroupId].results || []).length;
  const existingIndex = (store.examGroups[examGroupId].results || []).findIndex((item) => {
    if (String(item.fullExamId || "") === String(fullExamId || "")) {
      return true;
    }
    if (studentCardId && String(item.studentCardId || "") === String(studentCardId)) {
      return true;
    }
    return String(item.studentId || "") === String(studentId || "");
  });
  if (existingIndex >= 0) {
    const replace = window.confirm(
      `Já existe um resultado para ${studentName} nesta prova. Deseja substituir?`
    );
    if (!replace) {
      throw new Error("Resultado não salvo.");
    }
    const previous = store.examGroups[examGroupId].results[existingIndex] || {};
    row.resultId = String(previous.resultId || row.resultId);
    store.examGroups[examGroupId].results[existingIndex] = row;
  } else {
    store.examGroups[examGroupId].results.push(row);
  }

  console.log(
    "[Resultados] examGroupId:",
    examGroupId,
    "| antes:",
    beforeCount,
    "| depois:",
    (store.examGroups[examGroupId].results || []).length,
    "| aluno:",
    studentName,
  );
  saveResultsStore(store);
}

function correctProofLocally(proof, answers, rows) {
  const details = [];
  let correct = 0;
  let brancos = 0;
  let multiplas = 0;
  let ambiguas = 0;

  const rowByNumber = new Map();
  for (const row of rows || []) {
    rowByNumber.set(Number(row.questionNumber), row);
  }
  for (const question of proof.questoes || []) {
    const row = rowByNumber.get(Number(question.numero));
    const status = row?.status || (answers[String(question.numero)] ? "respondida" : "em_branco");

    let marked = "";
    if (status === "respondida") {
      marked = String(answers[String(question.numero)] || "").toUpperCase();
    }
    const expected = String(question.correta_letra || "").toUpperCase();
    const hit = status === "respondida" && marked !== "" && marked === expected;
    if (hit) {
      correct += 1;
    }
    if (status === "em_branco") {
      brancos += 1;
    } else if (status === "multipla_marcacao") {
      multiplas += 1;
    } else if (status === "ambigua") {
      ambiguas += 1;
    }
    details.push({
      numero: question.numero,
      marcada: marked,
      marcadas: row?.markedLetters || [],
      correta: expected,
      acertou: hit,
      status,
      scores: row?.scores || [],
      threshold: row?.adaptiveThreshold ?? null,
    });
  }
  const total = details.length;
  return {
    aluno: proof.aluno || "",
    acertos: correct,
    total,
    nota: total ? Math.round((correct / total) * 1000) / 100 : 0,
    brancos,
    multiplas,
    ambiguas,
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
    studentId: payload.sid || payload.studentId || payload.matricula || payload.m || "",
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
  const target = state.cardMode ? state.elements.cardStatus : state.elements.scanStatus;
  if (!target) {
    return;
  }
  if (target.textContent !== text) {
    target.textContent = text;
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

function updateDebugUi() {
  if (state.elements?.debugFilePanel) {
    state.elements.debugFilePanel.classList.toggle("hidden", !state.debug);
  }
}

function handleCardCaptureModeToggle() {
  const diagnostic = Boolean(state.elements?.debugToggle?.checked);
  const nextMode = diagnostic ? "diagnosticPhoto" : "realtime";
  const previousMode = state.cardCaptureMode;

  state.cardCaptureMode = nextMode;
  // Diagnóstico sempre liga debug; também permite ?debug=1.
  state.debug = Boolean(getQueryFlag("debug")) || diagnostic;
  updateDebugUi();

  if (previousMode !== nextMode && state.workflow === "answerScanning") {
    // Evita travas: ao alternar modo no meio da leitura, para tudo.
    stopLoops();
    stopAnswerCamera();
    clearOverlay();
    setWorkflowState("waitingUserToStartAnswerScan", "Modo alterado. Inicie a leitura novamente.");
  }

  updateCardControls();
}

function updateCardControls() {
  if (!state.cardMode) {
    setCameraPanelVisible(state.workflow === "qrScanning");
    return;
  }

  const scanning = state.workflow === "answerScanning";
  const done = state.workflow === "answerDetected";
  const diagnostic = state.cardCaptureMode === "diagnosticPhoto";

  // Modo padrão: preview em tempo real. Diagnóstico: captura por foto (câmera nativa).
  setCameraPanelVisible(!diagnostic && scanning);

  if (state.elements?.cardHint) {
    if (diagnostic) {
      state.elements.cardHint.textContent =
        "Modo diagnóstico (por foto): tire uma foto do cartão com os 4 quadradinhos pretos visíveis.";
    } else if (scanning) {
      state.elements.cardHint.textContent =
        "Modo padrão (tempo real): alinhe os 4 quadradinhos nos cantos do retângulo verde e mantenha o celular estável.";
    } else {
      state.elements.cardHint.textContent =
        "Modo padrão (tempo real): alinhe o cartão no retângulo verde e toque em iniciar. A leitura fecha automaticamente quando estiver estável.";
    }
  }

  // Novo layout (index.html atualizado)
  if (state.elements?.enableCameraButton) {
    state.elements.enableCameraButton.classList.toggle("hidden", scanning || done);
    state.elements.enableCameraButton.disabled = !state.opencvReady;
    state.elements.enableCameraButton.textContent = diagnostic
      ? "Tirar foto do cartão (diagnóstico)"
      : "Iniciar leitura (tempo real)";

    if (state.elements.disableCameraButton) {
      state.elements.disableCameraButton.classList.toggle("hidden", !scanning);
      state.elements.disableCameraButton.disabled = false;
    }

    if (state.elements.retakePhotoButton) {
      // Em ambos os modos: permite tentar novamente depois que terminar.
      state.elements.retakePhotoButton.classList.toggle("hidden", scanning || !done);
      state.elements.retakePhotoButton.disabled = scanning;
      state.elements.retakePhotoButton.textContent = diagnostic ? "Tirar nova foto" : "Ler novamente";
    }

    if (state.elements.confirmReadingButton) {
      state.elements.confirmReadingButton.classList.toggle("hidden", !done);
      state.elements.confirmReadingButton.disabled = false;
    }

    if (state.elements.cancelReadingButton) {
      // Mantém "Cancelar" apenas quando o resultado já existe.
      state.elements.cancelReadingButton.classList.toggle("hidden", scanning || !done);
      state.elements.cancelReadingButton.disabled = false;
    }
    return;
  }

  // Compatibilidade: card.html antigo.
  if (state.elements?.startScanButton && state.elements?.stopScanButton && state.elements?.startCardScanButton) {
    state.elements.startScanButton.disabled = scanning;
    state.elements.stopScanButton.disabled = false;
    state.elements.startCardScanButton.disabled = !state.stream || !state.opencvReady || scanning;
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
