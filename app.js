const KS_TABLE = {
  A: { min: 3, std: 4, max: 6, label: "아주 어두운 분위기/단순 작업" },
  B: { min: 6, std: 10, max: 15, label: "저조도 공공장소" },
  C: { min: 15, std: 20, max: 30, label: "빈번하지 않은 작업" },
  D: { min: 30, std: 40, max: 60, label: "단순 작업장" },
  E: { min: 60, std: 100, max: 150, label: "보통 활동 공간" },
  F: { min: 150, std: 200, max: 300, label: "일반 시작업" },
  G: { min: 300, std: 400, max: 600, label: "일반 작업면 조명" },
  H: { min: 600, std: 1000, max: 1500, label: "저휘도/소물체 작업" },
  I: { min: 1500, std: 2000, max: 3000, label: "장시간/정밀 작업" },
  J: { min: 3000, std: 4000, max: 6000, label: "고난도 시작업" },
  K: { min: 6000, std: 10000, max: 15000, label: "매우 특별한 시작업" },
};

const CAD_DEFAULT_TEXT_HEIGHT_MM = 150;

const cadState = {
  module: null,
  file: null,
  extension: "",
  content: null,
  rooms: [],
  placements: [],
};

const els = {
  floorHeight: document.getElementById("floorHeight"),
  lumens: document.getElementById("lumens"),
  maintenanceFactor: document.getElementById("maintenanceFactor"),
  circuits: document.getElementById("circuits"),
  ksClass: document.getElementById("ksClass"),
  ksLevel: document.getElementById("ksLevel"),
  targetLux: document.getElementById("targetLux"),
  fontStyle: document.getElementById("fontStyle"),
  uMode: document.getElementById("uMode"),

  applyKsBtn: document.getElementById("applyKsBtn"),
  resetBtn: document.getElementById("resetBtn"),

  roomWidth: document.getElementById("roomWidth"),
  roomHeight: document.getElementById("roomHeight"),
  roomArea: document.getElementById("roomArea"),
  calcSingleBtn: document.getElementById("calcSingleBtn"),
  singleOutput: document.getElementById("singleOutput"),
  singlePreview: document.getElementById("singlePreview"),
  singleResultPanel: document.getElementById("singleResultPanel"),
  resArea: document.getElementById("resArea"),
  resHeff: document.getElementById("resHeff"),
  resRI: document.getElementById("resRI"),
  resU: document.getElementById("resU"),
  resCount: document.getElementById("resCount"),
  resNote: document.getElementById("resNote"),

  bulkInput: document.getElementById("bulkInput"),
  calcBulkBtn: document.getElementById("calcBulkBtn"),
  downloadCsvBtn: document.getElementById("downloadCsvBtn"),
  bulkTableBody: document.querySelector("#bulkTable tbody"),
  bulkTotal: document.getElementById("bulkTotal"),

  genLispBtn: document.getElementById("genLispBtn"),
  downloadLispBtn: document.getElementById("downloadLispBtn"),
  lispOutput: document.getElementById("lispOutput"),

  cadDropZone: document.getElementById("cadDropZone"),
  cadFile: document.getElementById("cadFile"),
  cadFileName: document.getElementById("cadFileName"),
  cadUnit: document.getElementById("cadUnit"),
  cadMinArea: document.getElementById("cadMinArea"),
  cadMaxArea: document.getElementById("cadMaxArea"),
  cadGrid: document.getElementById("cadGrid"),
  cadOutputLayer: document.getElementById("cadOutputLayer"),
  cadSymbolRadius: document.getElementById("cadSymbolRadius"),
  cadTextOffset: document.getElementById("cadTextOffset"),
  cadAnalyzeBtn: document.getElementById("cadAnalyzeBtn"),
  cadLayoutBtn: document.getElementById("cadLayoutBtn"),
  cadExportBtn: document.getElementById("cadExportBtn"),
  cadStatus: document.getElementById("cadStatus"),
  cadRoomTableBody: document.querySelector("#cadRoomTable tbody"),
  cadRecommendedTotal: document.getElementById("cadRecommendedTotal"),
  cadPlacedTotal: document.getElementById("cadPlacedTotal"),

  ksSummary: document.getElementById("ksSummary"),
  scrollTopBtn: document.getElementById("scrollTopBtn"),
};

const presetButtons = Array.from(document.querySelectorAll(".preset-btn"));
const navPills = Array.from(document.querySelectorAll(".pill"));
let bulkResults = [];

function toNum(value, fallback = 0) {
  const n = Number(value);
  return Number.isFinite(n) ? n : fallback;
}

function round(value, digits = 2) {
  const factor = Math.pow(10, digits);
  return Math.round(value * factor) / factor;
}

function fmt(value, digits = 2) {
  return round(value, digits).toLocaleString("ko-KR", {
    minimumFractionDigits: digits,
    maximumFractionDigits: digits,
  });
}

function parseCircuits(value) {
  return String(value || "")
    .split(",")
    .map((v) => v.trim())
    .filter(Boolean);
}

function unitToMmFactor(unit) {
  if (unit === "m") return 1000;
  if (unit === "cm") return 10;
  return 1;
}

function rawLengthToMm(rawLength, unit) {
  return rawLength * unitToMmFactor(unit);
}

function mmToRawLength(mmLength, unit) {
  return mmLength / unitToMmFactor(unit);
}

function rawAreaToM2(rawArea, unit) {
  const mmFactor = unitToMmFactor(unit);
  const areaMm2 = rawArea * mmFactor * mmFactor;
  return areaMm2 / 1_000_000;
}

function samePoint(a, b, eps = 1e-6) {
  return Math.abs(a.x - b.x) <= eps && Math.abs(a.y - b.y) <= eps;
}

function polygonArea(points) {
  if (!points || points.length < 3) return 0;
  let sum = 0;
  for (let i = 0; i < points.length; i += 1) {
    const p1 = points[i];
    const p2 = points[(i + 1) % points.length];
    sum += p1.x * p2.y - p2.x * p1.y;
  }
  return sum / 2;
}

function bboxOfPoints(points) {
  let minX = Infinity;
  let minY = Infinity;
  let maxX = -Infinity;
  let maxY = -Infinity;

  points.forEach((p) => {
    if (p.x < minX) minX = p.x;
    if (p.y < minY) minY = p.y;
    if (p.x > maxX) maxX = p.x;
    if (p.y > maxY) maxY = p.y;
  });

  return {
    minX,
    minY,
    maxX,
    maxY,
    width: maxX - minX,
    height: maxY - minY,
  };
}

function pointOnSegment(px, py, x1, y1, x2, y2, eps = 1e-6) {
  const cross = (px - x1) * (y2 - y1) - (py - y1) * (x2 - x1);
  if (Math.abs(cross) > eps) return false;

  const dot = (px - x1) * (px - x2) + (py - y1) * (py - y2);
  return dot <= eps;
}

function pointInPolygon(point, polygon) {
  const x = point.x;
  const y = point.y;

  let inside = false;
  for (let i = 0, j = polygon.length - 1; i < polygon.length; j = i, i += 1) {
    const xi = polygon[i].x;
    const yi = polygon[i].y;
    const xj = polygon[j].x;
    const yj = polygon[j].y;

    if (pointOnSegment(x, y, xi, yi, xj, yj)) {
      return true;
    }

    const intersect = yi > y !== yj > y && x < ((xj - xi) * (y - yi)) / ((yj - yi) || 1e-12) + xi;
    if (intersect) inside = !inside;
  }

  return inside;
}

function snapToGrid(value, grid) {
  if (grid <= 0) return value;
  const snapped = Math.round(value / grid) * grid;
  return snapped <= 0 ? grid : snapped;
}

function populateKsUi() {
  Object.keys(KS_TABLE).forEach((grade) => {
    const op = document.createElement("option");
    op.value = grade;
    op.textContent = `${grade} (${KS_TABLE[grade].std} lx 표준)`;
    els.ksClass.appendChild(op);
  });
  els.ksClass.value = "G";

  Object.entries(KS_TABLE).forEach(([grade, info]) => {
    const li = document.createElement("li");
    li.textContent = `${grade}: ${info.min} / ${info.std} / ${info.max} lx (${info.label})`;
    els.ksSummary.appendChild(li);
  });
}

function applyKsLux() {
  const grade = els.ksClass.value;
  const level = els.ksLevel.value;
  const lux = KS_TABLE[grade]?.[level];
  if (!lux) return;
  els.targetLux.value = String(lux);
}

const PRESET_MAP = {
  office: { ksClass: "G", ksLevel: "std", targetLux: 400, floorHeight: 2700, lumens: 4400 },
  corridor: { ksClass: "F", ksLevel: "std", targetLux: 200, floorHeight: 2700, lumens: 2200 },
  storage: { ksClass: "E", ksLevel: "std", targetLux: 100, floorHeight: 2700, lumens: 6600 },
};

function applyPreset(name) {
  const preset = PRESET_MAP[name];
  if (!preset) return;

  presetButtons.forEach((btn) => {
    btn.classList.toggle("active", btn.dataset.preset === name);
  });

  els.ksClass.value = preset.ksClass;
  els.ksLevel.value = preset.ksLevel;
  els.floorHeight.value = String(preset.floorHeight);
  applyKsLux();

  if (Number.isFinite(preset.targetLux)) {
    els.targetLux.value = String(preset.targetLux);
  }
  if (Number.isFinite(preset.lumens)) {
    els.lumens.value = String(preset.lumens);
  }

  validateAllInputs();
  calcSingle();
  calcBulk();
  generateLispCode();
}

const U_POINTS = [
  { ri: 0.0, u: 0.45 },
  { ri: 1.0, u: 0.55 },
  { ri: 1.5, u: 0.65 },
  { ri: 2.5, u: 0.75 },
  { ri: 4.0, u: 0.8 },
];

function getUByRiStep(ri) {
  if (ri >= 4.0) return 0.8;
  if (ri >= 2.5) return 0.75;
  if (ri >= 1.5) return 0.65;
  if (ri >= 1.0) return 0.55;
  return 0.45;
}

function getUByRiInterpolated(ri) {
  if (ri <= U_POINTS[0].ri) return U_POINTS[0].u;
  if (ri >= U_POINTS[U_POINTS.length - 1].ri) return U_POINTS[U_POINTS.length - 1].u;

  for (let i = 0; i < U_POINTS.length - 1; i += 1) {
    const left = U_POINTS[i];
    const right = U_POINTS[i + 1];
    if (ri >= left.ri && ri <= right.ri) {
      const t = (ri - left.ri) / (right.ri - left.ri || 1);
      return left.u + (right.u - left.u) * t;
    }
  }

  return getUByRiStep(ri);
}

function getUByRi(ri, mode = "interp") {
  return mode === "step" ? getUByRiStep(ri) : getUByRiInterpolated(ri);
}

function computeLighting({
  floorHeight,
  lux,
  lumens,
  maintenanceFactor,
  widthMm,
  heightMm,
  areaM2,
  uMode = "interp",
}) {
  if (floorHeight <= 850) throw new Error("층고는 850mm보다 커야 합니다.");
  if (lux <= 0) throw new Error("목표 조도는 0보다 커야 합니다.");
  if (lumens <= 0) throw new Error("등기구 광속은 0보다 커야 합니다.");
  if (maintenanceFactor <= 0 || maintenanceFactor > 1) {
    throw new Error("보수율은 0보다 크고 1 이하이어야 합니다.");
  }
  if (widthMm <= 0 || heightMm <= 0) throw new Error("가로/세로는 0보다 커야 합니다.");
  if (areaM2 <= 0) throw new Error("면적은 0보다 커야 합니다.");

  const hEff = (floorHeight - 850) / 1000;
  const widthM = widthMm / 1000;
  const heightM = heightMm / 1000;
  const ri = areaM2 / (hEff * (widthM + heightM));
  const u = getUByRi(ri, uMode);
  const rawCount = (lux * areaM2) / (lumens * u * maintenanceFactor);
  const recommendedCount = Math.max(1, Math.ceil(rawCount));

  return {
    hEff,
    ri,
    u,
    rawCount,
    recommendedCount,
  };
}

function getCommonInputs() {
  return {
    floorHeight: toNum(els.floorHeight.value),
    lumens: toNum(els.lumens.value),
    maintenanceFactor: toNum(els.maintenanceFactor.value),
    circuits: els.circuits.value.trim(),
    targetLux: toNum(els.targetLux.value),
    fontStyle: els.fontStyle.value.trim() || "Standard",
    uMode: els.uMode?.value || "interp",
  };
}

function normalizeRoomDimensions({ widthMm, heightMm, areaM2 }) {
  let w = widthMm;
  let h = heightMm;
  const area = areaM2;
  let assumedSquare = false;

  if (area > 0 && (w <= 0 || h <= 0)) {
    const side = Math.sqrt(area * 1_000_000);
    w = side;
    h = side;
    assumedSquare = true;
  }

  return {
    widthMm: w,
    heightMm: h,
    areaM2: area,
    assumedSquare,
  };
}

function setInputValidity(input, message = "") {
  if (!input) return;
  input.setCustomValidity(message || "");
  input.classList.toggle("is-invalid", Boolean(message));
}

function validateNumberInput(input) {
  if (!input) return true;
  const raw = input.value;

  if (raw === "") {
    setInputValidity(input, "");
    return true;
  }

  const n = Number(raw);
  if (!Number.isFinite(n)) {
    setInputValidity(input, "숫자 값을 입력해주세요.");
    return false;
  }

  if (input.id === "floorHeight" && n <= 850) {
    setInputValidity(input, "층고는 850mm보다 커야 합니다.");
    return false;
  }

  if (input.id === "maintenanceFactor" && (n <= 0 || n > 1)) {
    setInputValidity(input, "보수율은 0 초과 1 이하여야 합니다.");
    return false;
  }

  if (input.id === "cadMaxArea") {
    const min = Number(els.cadMinArea.value || 0);
    if (n < min) {
      setInputValidity(input, "최대 면적은 최소 면적보다 커야 합니다.");
      return false;
    }
  }

  if (input.id === "cadMinArea") {
    const max = Number(els.cadMaxArea.value || Number.POSITIVE_INFINITY);
    if (n > max) {
      setInputValidity(input, "최소 면적은 최대 면적보다 작아야 합니다.");
      return false;
    }
  }

  setInputValidity(input, "");
  return true;
}

function validateAllInputs() {
  const numberInputs = Array.from(document.querySelectorAll("input[type='number']"));
  return numberInputs.map(validateNumberInput).every(Boolean);
}

function drawSinglePreview(widthMm, heightMm, count) {
  const canvas = els.singlePreview;
  if (!canvas) return;

  const ctx = canvas.getContext("2d");
  if (!ctx) return;

  const w = canvas.width;
  const h = canvas.height;
  ctx.clearRect(0, 0, w, h);

  // background
  ctx.fillStyle = "#060b18";
  ctx.fillRect(0, 0, w, h);

  // subtle grid
  ctx.strokeStyle = "rgba(93,160,255,0.04)";
  ctx.lineWidth = 0.5;
  for (let gx = 0; gx < w; gx += 24) {
    ctx.beginPath();
    ctx.moveTo(gx, 0);
    ctx.lineTo(gx, h);
    ctx.stroke();
  }
  for (let gy = 0; gy < h; gy += 24) {
    ctx.beginPath();
    ctx.moveTo(0, gy);
    ctx.lineTo(w, gy);
    ctx.stroke();
  }

  const margin = 24;
  const drawW = w - margin * 2;
  const drawH = h - margin * 2;
  const ratio = widthMm / Math.max(heightMm, 1e-9);

  let roomW = drawW;
  let roomH = roomW / ratio;
  if (roomH > drawH) {
    roomH = drawH;
    roomW = roomH * ratio;
  }

  const ox = (w - roomW) / 2;
  const oy = (h - roomH) / 2;

  // room fill + border
  ctx.fillStyle = "rgba(93,160,255,0.03)";
  ctx.fillRect(ox, oy, roomW, roomH);
  ctx.strokeStyle = "#4080cc";
  ctx.lineWidth = 2;
  ctx.strokeRect(ox, oy, roomW, roomH);

  // dimensions
  ctx.fillStyle = "#6688bb";
  ctx.font = "12px sans-serif";
  ctx.textAlign = "center";
  ctx.fillText(`${(widthMm / 1000).toFixed(1)}m`, ox + roomW / 2, oy - 10);
  ctx.save();
  ctx.translate(ox - 14, oy + roomH / 2);
  ctx.rotate(-Math.PI / 2);
  ctx.fillText(`${(heightMm / 1000).toFixed(1)}m`, 0, 0);
  ctx.restore();
  ctx.textAlign = "left";

  if (count <= 0) return;

  const cols = Math.max(1, Math.round(Math.sqrt(count * ratio)));
  const rows = Math.max(1, Math.ceil(count / cols));

  // keep lights away from walls
  const insetX = roomW * 0.15;
  const insetY = roomH * 0.15;
  const innerW = Math.max(0, roomW - insetX * 2);
  const innerH = Math.max(0, roomH - insetY * 2);
  const dx = cols > 1 ? innerW / (cols - 1) : 0;
  const dy = rows > 1 ? innerH / (rows - 1) : 0;

  let placed = 0;
  for (let i = 0; i < cols && placed < count; i += 1) {
    for (let j = 0; j < rows && placed < count; j += 1) {
      const cx = cols > 1 ? ox + insetX + i * dx : ox + roomW / 2;
      const cy = rows > 1 ? oy + insetY + j * dy : oy + roomH / 2;

      const grad = ctx.createRadialGradient(cx, cy, 0, cx, cy, 20);
      grad.addColorStop(0, "rgba(255,215,106,0.25)");
      grad.addColorStop(1, "rgba(255,215,106,0)");
      ctx.fillStyle = grad;
      ctx.fillRect(cx - 20, cy - 20, 40, 40);

      ctx.fillStyle = "#ffd76a";
      ctx.beginPath();
      ctx.arc(cx, cy, 5, 0, Math.PI * 2);
      ctx.fill();

      ctx.strokeStyle = "rgba(255,215,106,0.6)";
      ctx.lineWidth = 1;
      ctx.beginPath();
      ctx.moveTo(cx - 9, cy);
      ctx.lineTo(cx + 9, cy);
      ctx.moveTo(cx, cy - 9);
      ctx.lineTo(cx, cy + 9);
      ctx.stroke();

      placed += 1;
    }
  }

  // legend
  ctx.fillStyle = "rgba(15,23,48,0.75)";
  const legendW = 190;
  const legendH = 28;
  const legendX = w - legendW - 12;
  const legendY = 8;
  ctx.fillRect(legendX, legendY, legendW, legendH);
  ctx.strokeStyle = "rgba(93,160,255,0.2)";
  ctx.lineWidth = 1;
  ctx.strokeRect(legendX, legendY, legendW, legendH);

  ctx.fillStyle = "#ffd76a";
  ctx.beginPath();
  ctx.arc(legendX + 14, legendY + 14, 4, 0, Math.PI * 2);
  ctx.fill();

  ctx.fillStyle = "#b0c4ee";
  ctx.font = "12px sans-serif";
  ctx.fillText(`${count}개 (${cols}×${rows} 배열)`, legendX + 26, legendY + 18);
}

function calcSingle() {
  const common = getCommonInputs();

  const singleInputs = [
    els.floorHeight,
    els.lumens,
    els.maintenanceFactor,
    els.targetLux,
    els.roomWidth,
    els.roomHeight,
    els.roomArea,
  ];
  const singleValid = singleInputs.map(validateNumberInput).every(Boolean);
  if (!singleValid) {
    els.singleOutput.textContent = "오류: 입력값을 확인해주세요.";
    if (els.singleResultPanel) {
      els.singleResultPanel.style.display = "none";
    }
    drawSinglePreview(1, 1, 0);
    return;
  }

  const widthInputMm = toNum(els.roomWidth.value);
  const heightInputMm = toNum(els.roomHeight.value);

  const areaManual = toNum(els.roomArea.value, NaN);
  const baseAreaM2 = Number.isFinite(areaManual) && areaManual > 0
    ? areaManual
    : (widthInputMm * heightInputMm) / 1_000_000;

  const normalized = normalizeRoomDimensions({
    widthMm: widthInputMm,
    heightMm: heightInputMm,
    areaM2: baseAreaM2,
  });

  try {
    const result = computeLighting({
      floorHeight: common.floorHeight,
      lux: common.targetLux,
      lumens: common.lumens,
      maintenanceFactor: common.maintenanceFactor,
      widthMm: normalized.widthMm,
      heightMm: normalized.heightMm,
      areaM2: normalized.areaM2,
      uMode: common.uMode,
    });

    const txt = [
      `[단일 실 계산 결과]`,
      `면적: ${fmt(normalized.areaM2, 2)} ㎡`,
      `유효높이(Hm): ${fmt(result.hEff, 2)} m`,
      `실지수(RI): ${fmt(result.ri, 2)}`,
      `이용률(U): ${fmt(result.u, 2)} (${common.uMode === "step" ? "계단식" : "보간"})`,
      `계산 수량(실수): ${fmt(result.rawCount, 2)} 개`,
      `추천 수량(올림): ${result.recommendedCount.toLocaleString("ko-KR")} 개`,
      normalized.assumedSquare ? `※ 가로/세로가 없어서 정사각형으로 가정해 계산했습니다.` : "",
    ]
      .filter(Boolean)
      .join("\n");

    els.singleOutput.textContent = txt;

    if (els.singleResultPanel) {
      els.resArea.textContent = `${fmt(normalized.areaM2, 2)} ㎡`;
      els.resHeff.textContent = `${fmt(result.hEff, 2)} m`;
      els.resRI.textContent = fmt(result.ri, 2);
      els.resU.textContent = fmt(result.u, 2);
      els.resCount.textContent = `${result.recommendedCount.toLocaleString("ko-KR")}개`;
      els.resNote.textContent = normalized.assumedSquare
        ? "※ 가로/세로 없이 정사각형 가정"
        : "";
      els.singleResultPanel.style.display = "block";
    }

    drawSinglePreview(normalized.widthMm, normalized.heightMm, result.recommendedCount);
  } catch (error) {
    els.singleOutput.textContent = `오류: ${error.message}`;
    if (els.singleResultPanel) {
      els.singleResultPanel.style.display = "none";
    }
    drawSinglePreview(1, 1, 0);
  }
}

function parseCsvLine(line) {
  const cols = line.split(",").map((s) => s.trim());
  if (cols.length < 1) return null;

  const name = cols[0] || `Room-${Math.random().toString(36).slice(2, 6)}`;
  const widthMm = toNum(cols[1], NaN);
  const heightMm = toNum(cols[2], NaN);
  const areaM2Raw = toNum(cols[3], NaN);
  const luxRaw = toNum(cols[4], NaN);

  const hasWH = Number.isFinite(widthMm) && Number.isFinite(heightMm) && widthMm > 0 && heightMm > 0;
  const hasArea = Number.isFinite(areaM2Raw) && areaM2Raw > 0;
  if (!hasWH && !hasArea) return null;

  return {
    name,
    widthMm,
    heightMm,
    areaM2Raw,
    luxRaw,
  };
}

function calcBulk() {
  const common = getCommonInputs();

  const bulkInputs = [els.floorHeight, els.lumens, els.maintenanceFactor, els.targetLux];
  const bulkValid = bulkInputs.map(validateNumberInput).every(Boolean);
  if (!bulkValid) {
    bulkResults = [];
    renderBulkTable([]);
    return;
  }

  const lines = els.bulkInput.value
    .split(/\r?\n/)
    .map((v) => v.trim())
    .filter(Boolean);

  const rows = [];
  for (const line of lines) {
    const item = parseCsvLine(line);
    if (!item) continue;

    const areaFromWH = Number.isFinite(item.widthMm) && Number.isFinite(item.heightMm)
      ? (item.widthMm * item.heightMm) / 1_000_000
      : NaN;

    const baseAreaM2 = Number.isFinite(item.areaM2Raw) && item.areaM2Raw > 0
      ? item.areaM2Raw
      : areaFromWH;

    const normalized = normalizeRoomDimensions({
      widthMm: item.widthMm,
      heightMm: item.heightMm,
      areaM2: baseAreaM2,
    });

    const lux = Number.isFinite(item.luxRaw) && item.luxRaw > 0 ? item.luxRaw : common.targetLux;

    try {
      const result = computeLighting({
        floorHeight: common.floorHeight,
        lux,
        lumens: common.lumens,
        maintenanceFactor: common.maintenanceFactor,
        widthMm: normalized.widthMm,
        heightMm: normalized.heightMm,
        areaM2: normalized.areaM2,
        uMode: common.uMode,
      });

      rows.push({
        name: item.name,
        areaM2: normalized.areaM2,
        ri: result.ri,
        u: result.u,
        lux,
        count: result.recommendedCount,
      });
    } catch (error) {
      rows.push({
        name: item.name,
        areaM2: normalized.areaM2,
        ri: NaN,
        u: NaN,
        lux,
        count: `오류: ${error.message}`,
      });
    }
  }

  bulkResults = rows;
  renderBulkTable(rows);
}

function appendCell(tr, label, value) {
  const td = document.createElement("td");
  td.textContent = String(value);
  if (label) {
    td.setAttribute("data-label", label);
  }
  tr.appendChild(td);
}

function renderBulkTable(rows) {
  els.bulkTableBody.innerHTML = "";
  let total = 0;

  for (const row of rows) {
    const tr = document.createElement("tr");
    const isError = typeof row.count !== "number";
    if (!isError) total += row.count;

    appendCell(tr, "실명", row.name);
    appendCell(tr, "면적(㎡)", Number.isFinite(row.areaM2) ? fmt(row.areaM2, 2) : "-");
    appendCell(tr, "RI", Number.isFinite(row.ri) ? fmt(row.ri, 2) : "-");
    appendCell(tr, "U", Number.isFinite(row.u) ? fmt(row.u, 2) : "-");
    appendCell(tr, "조도(Lx)", Number.isFinite(row.lux) ? fmt(row.lux, 0) : "-");
    appendCell(tr, "추천 수량(개)", isError ? row.count : row.count.toLocaleString("ko-KR"));

    els.bulkTableBody.appendChild(tr);
  }

  els.bulkTotal.textContent = total.toLocaleString("ko-KR");
}

function downloadBulkCsv() {
  if (!bulkResults.length) {
    alert("먼저 일괄 계산을 실행하세요.");
    return;
  }

  const lines = ["room,area_m2,ri,u,lux,recommended_count"];
  for (const row of bulkResults) {
    lines.push([
      row.name,
      Number.isFinite(row.areaM2) ? round(row.areaM2, 4) : "",
      Number.isFinite(row.ri) ? round(row.ri, 4) : "",
      Number.isFinite(row.u) ? round(row.u, 4) : "",
      Number.isFinite(row.lux) ? round(row.lux, 2) : "",
      typeof row.count === "number" ? row.count : `"${row.count}"`,
    ].join(","));
  }

  downloadFile("lighting_bulk_result.csv", lines.join("\n"), "text/csv;charset=utf-8");
}

function escapeLispString(value) {
  return String(value).replace(/\\/g, "\\\\").replace(/\"/g, '\\\"');
}

function lispNum(value, digits = 4) {
  const n = Number(value);
  if (!Number.isFinite(n)) return "0";
  return n.toFixed(digits).replace(/\.?0+$/, "");
}

function generateLispCode() {
  const common = getCommonInputs();

  const conf = {
    floorHeight: lispNum(common.floorHeight, 2),
    targetLux: lispNum(common.targetLux, 2),
    lumens: lispNum(common.lumens, 2),
    maintenanceFactor: lispNum(common.maintenanceFactor, 3),
    circuits: escapeLispString(common.circuits || "45, 46"),
    fontStyle: escapeLispString(common.fontStyle || "Standard"),
  };

  const lispUFunctionLines = common.uMode === "step"
    ? [
        ";; U mode: step",
        "(defun qqq_u_by_ri (ri)",
        "  (cond ((>= ri 4.0) 0.80)",
        "        ((>= ri 2.5) 0.75)",
        "        ((>= ri 1.5) 0.65)",
        "        ((>= ri 1.0) 0.55)",
        "        (t 0.45)))",
      ]
    : [
        ";; U mode: interpolated",
        "(defun qqq_u_by_ri (ri / t)",
        "  (cond",
        "    ((<= ri 0.0) 0.45)",
        "    ((<= ri 1.0) (+ 0.45 (* (/ (- ri 0.0) 1.0) (- 0.55 0.45))))",
        "    ((<= ri 1.5) (+ 0.55 (* (/ (- ri 1.0) 0.5) (- 0.65 0.55))))",
        "    ((<= ri 2.5) (+ 0.65 (* (/ (- ri 1.5) 1.0) (- 0.75 0.65))))",
        "    ((<= ri 4.0) (+ 0.75 (* (/ (- ri 2.5) 1.5) (- 0.80 0.75))))",
        "    (t 0.80)))",
      ];

  const code = [
    ";; qqq_auto.lsp - generated by lighting-layout-oneclick-app",
    "(vl-load-com)",
    "",
    ";; --- defaults ---",
    `(if (not *qqq_h*)    (setq *qqq_h* ${conf.floorHeight}))`,
    `(if (not *qqq_lux*)  (setq *qqq_lux* ${conf.targetLux}))`,
    `(if (not *qqq_lm*)   (setq *qqq_lm* ${conf.lumens}))`,
    `(if (not *qqq_cir*)  (setq *qqq_cir* "${conf.circuits}"))`,
    "(if (not *qqq_txth*) (setq *qqq_txth* 150.0))",
    "(if (not *qqq_txto*) (setq *qqq_txto* 300.0))",
    `(if (not *qqq_font*) (setq *qqq_font* "${conf.fontStyle}"))`,
    `(setq *qqq_m* ${conf.maintenanceFactor})`,
    "(setq *qqq_grid* 300.0)",
    "(if (not *qqq_blk*)  (setq *qqq_blk* \"\"))",
    "(if (not *qqq_rot*)  (setq *qqq_rot* 0.0))",
    "",
    ";; --- helpers ---",
    "(defun qqq_split (str delim / pos out)",
    "  (while (setq pos (vl-string-search delim str))",
    "    (setq out (cons (vl-string-trim \" \" (substr str 1 pos)) out)",
    "          str (substr str (+ pos 2))))",
    "  (reverse (cons (vl-string-trim \" \" str) out)))",
    "",
    ...lispUFunctionLines,
    "",
    "(defun qqq_inside (pt v_list / ang sum p1 p2 i)",
    "  (setq sum 0.0 i 0)",
    "  (while (< i (length v_list))",
    "    (setq p1 (nth i v_list)",
    "          p2 (if (= i (1- (length v_list))) (car v_list) (nth (1+ i) v_list)))",
    "    (setq ang (- (angle pt p2) (angle pt p1)))",
    "    (if (> ang pi) (setq ang (- ang (* 2 pi))))",
    "    (if (< ang (* -1 pi)) (setq ang (+ ang (* 2 pi))))",
    "    (setq sum (+ sum ang)",
    "          i (1+ i)))",
    "  (> (abs sum) 0.001))",
    "",
    "(defun qqq_closed_p (e / d f70)",
    "  (setq d (entget e)",
    "        f70 (cdr (assoc 70 d)))",
    "  (and f70 (/= 0 (logand f70 1))))",
    "",
    "(defun qqq_pick_block (/ blk)",
    "  (setq blk (car (entsel (strcat \"\\n전등 블록 선택 (현재: \" *qqq_blk* \") [변경 시 클릭/유지는 엔터]: \"))))",
    "  (if blk",
    "    (setq *qqq_blk* (cdr (assoc 2 (entget blk)))",
    "          *qqq_rot* (* (/ (cdr (assoc 50 (entget blk))) pi) 180.0))))",
    "",
    "(defun qqq_calc_core (area width height / h_eff ri u cnt)",
    "  (setq h_eff (/ (- *qqq_h* 850.0) 1000.0))",
    "  (if (<= h_eff 0.1) (setq h_eff 0.1))",
    "  (setq ri (/ area (* h_eff (+ (/ width 1000.0) (/ height 1000.0)))))",
    "  (setq u (qqq_u_by_ri ri))",
    "  (setq cnt (max 1 (fix (+ 0.999 (/ (* *qqq_lux* area) (* *qqq_lm* u *qqq_m*))))))",
    "  (list ri u cnt))",
    "",
    "(defun qqq_place_in_boundary (ent cir_list cir_idx / obj area bmin bmax min_pt max_pt width height calc ri u target ratio cols rows x_spacing y_spacing x_start y_start i j pt v_list pts_done txt_pt placed)",
    "  (setq obj (vlax-ename->vla-object ent)",
    "        area (/ (vla-get-area obj) 1000000.0))",
    "  (vla-getboundingbox obj 'bmin 'bmax)",
    "  (setq min_pt (vlax-safearray->list bmin)",
    "        max_pt (vlax-safearray->list bmax)",
    "        width (- (car max_pt) (car min_pt))",
    "        height (- (cadr max_pt) (cadr min_pt)))",
    "",
    "  (setq calc (qqq_calc_core area width height)",
    "        ri (nth 0 calc)",
    "        u (nth 1 calc)",
    "        target (nth 2 calc))",
    "",
    "  (setq ratio (/ width (max height 1e-6))",
    "        cols (fix (+ 0.5 (sqrt (* target ratio)))))",
    "  (if (< cols 1) (setq cols 1))",
    "  (setq rows (fix (+ 0.5 (/ (float target) (float cols)))))",
    "  (if (< rows 1) (setq rows 1))",
    "",
    "  (setq x_spacing (* (max 1 (fix (+ 0.5 (/ (/ width cols) *qqq_grid*)))) *qqq_grid*)",
    "        y_spacing (* (max 1 (fix (+ 0.5 (/ (/ height rows) *qqq_grid*)))) *qqq_grid*))",
    "",
    "  (setq x_start (* (fix (+ (/ (+ (car min_pt) (/ (- width (* x_spacing (1- cols))) 2.0)) *qqq_grid*) 0.5)) *qqq_grid*)",
    "        y_start (* (fix (+ (/ (+ (cadr min_pt) (/ (- height (* y_spacing (1- rows))) 2.0)) *qqq_grid*) 0.5)) *qqq_grid*))",
    "",
    "  (setq v_list (mapcar 'cdr (vl-remove-if-not '(lambda (x) (= (car x) 10)) (entget ent)))",
    "        pts_done nil",
    "        placed 0",
    "        i 0)",
    "",
    "  (while (< i cols)",
    "    (setq j 0)",
    "    (while (< j rows)",
    "      (setq pt (list (+ x_start (* i x_spacing)) (+ y_start (* j y_spacing)) 0.0))",
    "      (if (and (qqq_inside pt v_list) (not (member pt pts_done)))",
    "        (progn",
    "          (command \"-insert\" *qqq_blk* pt \"1\" \"1\" (rtos *qqq_rot* 2 2))",
    "          (setq txt_pt (list (+ (car pt) *qqq_txto*) (- (cadr pt) *qqq_txto*) 0.0))",
    "          (entmake (list '(0 . \"TEXT\")",
    "                         (cons 10 txt_pt)",
    "                         (cons 40 *qqq_txth*)",
    "                         (cons 1 (nth cir_idx cir_list))",
    "                         (cons 7 *qqq_font*)",
    "                         '(62 . 2)))",
    "          (setq cir_idx (if (>= (1+ cir_idx) (length cir_list)) 0 (1+ cir_idx))",
    "                pts_done (cons pt pts_done)",
    "                placed (1+ placed))))",
    "      (setq j (1+ j)))",
    "    (setq i (1+ i)))",
    "",
    "  (list cir_idx placed area ri u target))",
    "",
    ";; --- command: single room via boundary ---",
    "(defun c:QQQ (/ os pick_pt ent last_ent cir_list cir_idx ret)",
    "  (setq os (getvar \"OSMODE\"))",
    "  (setvar \"OSMODE\" 0)",
    "  (qqq_pick_block)",
    "  (if (and (/= *qqq_blk* \"\") (setq pick_pt (getpoint \"\\n실 내부 클릭: \")))",
    "    (progn",
    "      (setq last_ent (entlast))",
    "      (vl-cmdf \"-boundary\" pick_pt \"\")",
    "      (setq ent (entlast))",
    "      (if (not (equal last_ent ent))",
    "        (progn",
    "          (setq cir_list (qqq_split *qqq_cir* \",\") cir_idx 0)",
    "          (setq ret (qqq_place_in_boundary ent cir_list cir_idx))",
    "          (entdel ent)",
    "          (princ (strcat \"\\n[QQQ] 배치 완료 / 배치수량: \" (itoa (nth 1 ret)))))",
    "        (princ \"\\n경계 인식 실패.\")))",
    "    (princ \"\\n블록 또는 실 선택 오류.\"))",
    "  (setvar \"OSMODE\" os)",
    "  (princ))",
    "",
    ";; --- command: all rooms by selected closed boundaries ---",
    "(defun c:QQQALL (/ os ss i e cir_list cir_idx ret totalRooms totalPlaced)",
    "  (setq os (getvar \"OSMODE\"))",
    "  (setvar \"OSMODE\" 0)",
    "  (qqq_pick_block)",
    "  (if (/= *qqq_blk* \"\")",
    "    (progn",
    "      (setq ss (ssget \"_:L\" '((0 . \"LWPOLYLINE,POLYLINE\"))))",
    "      (if ss",
    "        (progn",
    "          (setq cir_list (qqq_split *qqq_cir* \",\")",
    "                cir_idx 0",
    "                i 0",
    "                totalRooms 0",
    "                totalPlaced 0)",
    "          (while (< i (sslength ss))",
    "            (setq e (ssname ss i))",
    "            (if (qqq_closed_p e)",
    "              (progn",
    "                (setq ret (qqq_place_in_boundary e cir_list cir_idx)",
    "                      cir_idx (nth 0 ret)",
    "                      totalPlaced (+ totalPlaced (nth 1 ret))",
    "                      totalRooms (1+ totalRooms))))",
    "            (setq i (1+ i)))",
    "          (princ (strcat \"\\n[QQQALL] 완료 / 실 수: \" (itoa totalRooms) \" / 총 배치 수량: \" (itoa totalPlaced))))",
    "        (princ \"\\n닫힌 경계 폴리라인을 선택하세요.\")))",
    "    (princ \"\\n전등 블록 선택이 필요합니다.\"))",
    "  (setvar \"OSMODE\" os)",
    "  (princ))",
    "",
    "(princ \"\\nQQQ / QQQALL loaded.\")",
    "(princ)",
  ].join("\n");

  els.lispOutput.value = code;
}

function downloadLisp() {
  if (!els.lispOutput.value.trim()) {
    generateLispCode();
  }
  downloadFile("qqq_auto.lsp", els.lispOutput.value, "text/plain;charset=utf-8");
}

function toDownloadPayload(data) {
  if (data instanceof Uint8Array) return data;
  if (data instanceof ArrayBuffer) return new Uint8Array(data);
  if (typeof data === "string") return new TextEncoder().encode(data);
  return new TextEncoder().encode(String(data ?? ""));
}

function downloadFile(filename, data, mimeType) {
  const payload = toDownloadPayload(data);
  const blob = new Blob([payload], { type: mimeType || "application/octet-stream" });
  const url = URL.createObjectURL(blob);

  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();

  // Delay revoke for better compatibility on Safari/iOS large-file downloads.
  setTimeout(() => URL.revokeObjectURL(url), 60_000);
}

function getCadConfig() {
  const unit = els.cadUnit.value;
  const minAreaM2 = toNum(els.cadMinArea.value, 0);
  const maxAreaM2 = toNum(els.cadMaxArea.value, Number.POSITIVE_INFINITY);

  return {
    unit,
    minAreaM2,
    maxAreaM2,
    gridMm: Math.max(10, toNum(els.cadGrid.value, 300)),
    outputLayer: els.cadOutputLayer.value.trim() || "LIGHT_AUTO",
    symbolRadiusMm: Math.max(1, toNum(els.cadSymbolRadius.value, 150)),
    textOffsetMm: Math.max(0, toNum(els.cadTextOffset.value, 300)),
  };
}

function setCadStatus(message, append = false) {
  if (append) {
    els.cadStatus.textContent = `${els.cadStatus.textContent}\n${message}`.trim();
  } else {
    els.cadStatus.textContent = message;
  }
}

function renderCadRooms(rows) {
  els.cadRoomTableBody.innerHTML = "";

  let recTotal = 0;
  let placedTotal = 0;

  rows.forEach((room) => {
    recTotal += room.recommendedCount || 0;
    placedTotal += room.placedCount || 0;

    const tr = document.createElement("tr");
    appendCell(tr, "방", room.name);
    appendCell(tr, "레이어", room.layer || "-");
    appendCell(tr, "면적(㎡)", fmt(room.areaM2, 2));
    appendCell(tr, "RI", fmt(room.ri, 2));
    appendCell(tr, "U", fmt(room.u, 2));
    appendCell(tr, "추천 수량", (room.recommendedCount || 0).toLocaleString("ko-KR"));
    appendCell(tr, "배치 수량", (room.placedCount || 0).toLocaleString("ko-KR"));
    els.cadRoomTableBody.appendChild(tr);
  });

  els.cadRecommendedTotal.textContent = recTotal.toLocaleString("ko-KR");
  els.cadPlacedTotal.textContent = placedTotal.toLocaleString("ko-KR");
}

const CAD_SCRIPT_URLS = [
  "https://cdn.jsdelivr.net/npm/@mlightcad/libdxfrw-web@0.1.0/dist/libdxfrw.js",
  "https://unpkg.com/@mlightcad/libdxfrw-web@0.1.0/dist/libdxfrw.js",
];

function loadScript(src, timeoutMs = 15000) {
  return new Promise((resolve, reject) => {
    const existing = document.querySelector(`script[data-cad-src=\"${src}\"]`);
    if (existing) {
      existing.addEventListener("load", () => resolve(), { once: true });
      existing.addEventListener("error", () => reject(new Error(`failed: ${src}`)), { once: true });
      return;
    }

    const script = document.createElement("script");
    script.src = src;
    script.async = true;
    script.dataset.cadSrc = src;

    const timer = setTimeout(() => {
      script.remove();
      reject(new Error(`timeout: ${src}`));
    }, timeoutMs);

    script.onload = () => {
      clearTimeout(timer);
      resolve();
    };

    script.onerror = () => {
      clearTimeout(timer);
      script.remove();
      reject(new Error(`failed: ${src}`));
    };

    document.head.appendChild(script);
  });
}

async function ensureCadScriptLoaded() {
  if (typeof window.createModule === "function") return;

  let lastError;
  for (const src of CAD_SCRIPT_URLS) {
    try {
      setCadStatus(`CAD 엔진 스크립트 로딩 중...\n- ${src}`);
      await loadScript(src, 18000);
      if (typeof window.createModule === "function") return;
    } catch (error) {
      lastError = error;
    }
  }

  throw new Error(
    `CAD 엔진 스크립트를 불러오지 못했습니다. 네트워크/광고차단/보안앱 설정을 확인해주세요. ` +
      `${lastError ? `(원인: ${lastError.message})` : ""}`,
  );
}

async function ensureCadModule() {
  if (cadState.module) return cadState.module;

  await ensureCadScriptLoaded();

  setCadStatus("CAD 엔진 로딩 중...");

  const timeoutMs = 25000;
  cadState.module = await Promise.race([
    window.createModule(),
    new Promise((_, reject) => {
      setTimeout(() => reject(new Error("CAD 엔진 초기화 시간이 초과되었습니다.")), timeoutMs);
    }),
  ]);

  return cadState.module;
}

function inferCadTypeFromName(fileName) {
  const name = String(fileName || "");
  const dot = name.lastIndexOf(".");
  if (dot >= 0) {
    const ext = name.slice(dot + 1).toLowerCase();
    if (ext === "dxf" || ext === "dwg") return ext;
  }
  return "";
}

function inferCadTypeFromBuffer(buffer) {
  const headBytes = new Uint8Array(buffer.slice(0, 128));
  const headText = new TextDecoder("ascii").decode(headBytes);

  // DWG signature examples: AC1015, AC1024, AC1032...
  if (/^AC10\d{2}/.test(headText)) {
    return "dwg";
  }

  // DXF ASCII often starts with: "0\nSECTION"
  if (/SECTION/i.test(headText) || /\nENTITIES\n/i.test(headText)) {
    return "dxf";
  }

  return "";
}

async function loadCadFileFromInput() {
  const file = els.cadFile.files?.[0];
  if (!file) {
    throw new Error("DXF 또는 DWG 파일을 먼저 업로드해주세요.");
  }

  let ext = inferCadTypeFromName(file.name);
  const buffer = await file.arrayBuffer();

  if (!ext) {
    ext = inferCadTypeFromBuffer(buffer);
  }

  if (!["dxf", "dwg"].includes(ext)) {
    throw new Error("파일 형식을 판별하지 못했습니다. 확장자를 .dxf 또는 .dwg로 지정해주세요.");
  }

  cadState.file = file;
  cadState.extension = ext;
  cadState.rooms = [];
  cadState.placements = [];

  // Keep original bytes for both DXF and DWG to avoid encoding loss on non-ASCII drawings.
  cadState.content = buffer;

  setCadStatus(`파일 로드 완료: ${file.name} (${ext.toUpperCase()})`);
}

async function parseCadDatabase(content, extension) {
  const lib = await ensureCadModule();
  const database = new lib.DRW_Database();
  const fileHandler = new lib.DRW_FileHandler();
  fileHandler.database = database;

  let ok = false;

  try {
    if (extension === "dxf") {
      const dxf = new lib.DRW_DxfRW(content);
      ok = dxf.read(fileHandler, false);
      dxf.delete();
    } else {
      const dwg = new lib.DRW_DwgR(content);
      ok = dwg.read(fileHandler, false);
      dwg.delete();
    }
  } catch (error) {
    fileHandler.delete();
    database.delete();
    throw error;
  }

  if (!ok) {
    const entitySize = database?.mBlock?.entities?.size ? database.mBlock.entities.size() : 0;
    if (entitySize <= 0) {
      fileHandler.delete();
      database.delete();
      throw new Error("도면 파싱에 실패했습니다. 손상 파일이거나 지원되지 않는 버전일 수 있습니다.");
    }
  }

  return { lib, database, fileHandler };
}

function extractPolylineVertices(entity, lib) {
  if (entity.eType === lib.DRW_ETYPE.LWPOLYLINE) {
    const list = entity.getVertexList();
    const points = [];
    for (let i = 0, size = list.size(); i < size; i += 1) {
      const v = list.get(i);
      if (!v) continue;
      points.push({ x: v.x, y: v.y });
    }
    return {
      points,
      flags: entity.flags || 0,
    };
  }

  if (entity.eType === lib.DRW_ETYPE.POLYLINE) {
    const list = entity.getVertexList();
    const points = [];
    for (let i = 0, size = list.size(); i < size; i += 1) {
      const v = list.get(i);
      if (!v) continue;
      const bp = v.basePoint;
      const x = bp?.x;
      const y = bp?.y;
      if (!Number.isFinite(x) || !Number.isFinite(y)) continue;
      points.push({ x, y });
    }
    return {
      points,
      flags: entity.flags || 0,
    };
  }

  return null;
}

function isClosedPolyline(points, flags) {
  if (points.length < 3) return false;
  if ((flags & 1) === 1) return true;
  return samePoint(points[0], points[points.length - 1]);
}

function buildRoomFromPolyline(polylineInfo, roomIndex, layer, common, cadCfg) {
  const points = [...polylineInfo.points];
  if (samePoint(points[0], points[points.length - 1])) {
    points.pop();
  }

  if (points.length < 3) return null;

  const areaRaw = Math.abs(polygonArea(points));
  const areaM2 = rawAreaToM2(areaRaw, cadCfg.unit);

  if (areaM2 < cadCfg.minAreaM2 || areaM2 > cadCfg.maxAreaM2) {
    return null;
  }

  const bboxRaw = bboxOfPoints(points);
  const widthMm = rawLengthToMm(bboxRaw.width, cadCfg.unit);
  const heightMm = rawLengthToMm(bboxRaw.height, cadCfg.unit);

  if (widthMm <= 0 || heightMm <= 0) return null;

  const result = computeLighting({
    floorHeight: common.floorHeight,
    lux: common.targetLux,
    lumens: common.lumens,
    maintenanceFactor: common.maintenanceFactor,
    widthMm,
    heightMm,
    areaM2,
    uMode: common.uMode,
  });

  return {
    id: roomIndex,
    name: `ROOM_${roomIndex}`,
    layer: layer || "0",
    points,
    bboxRaw,
    areaM2,
    ri: result.ri,
    u: result.u,
    recommendedCount: result.recommendedCount,
    placedCount: 0,
  };
}

function detectRooms(lib, database, common, cadCfg) {
  const block = database.mBlock;
  if (!block || !block.entities) {
    throw new Error("모델 공간 엔티티를 찾지 못했습니다.");
  }

  const rooms = [];
  const entities = block.entities;
  let roomIndex = 1;

  for (let i = 0, size = entities.size(); i < size; i += 1) {
    const entity = entities.get(i);
    if (!entity) continue;

    const polyline = extractPolylineVertices(entity, lib);
    if (!polyline) continue;
    if (!isClosedPolyline(polyline.points, polyline.flags)) continue;

    const room = buildRoomFromPolyline(polyline, roomIndex, entity.layer, common, cadCfg);
    if (!room) continue;

    rooms.push(room);
    roomIndex += 1;
  }

  rooms.sort((a, b) => b.areaM2 - a.areaM2);
  rooms.forEach((room, idx) => {
    room.id = idx + 1;
    room.name = `${room.layer || "ROOM"}_${String(idx + 1).padStart(3, "0")}`;
  });

  return rooms;
}

function pointKey(pt, digits = 4) {
  return `${pt.x.toFixed(digits)}:${pt.y.toFixed(digits)}`;
}

function generateRoomPlacementPoints(room, targetCount, cadCfg) {
  if (!targetCount || targetCount <= 0) return [];

  const gridRaw = mmToRawLength(cadCfg.gridMm, cadCfg.unit);
  const { minX, minY, maxX, maxY, width, height } = room.bboxRaw;

  if (width <= 0 || height <= 0) return [];

  const ratio = width / Math.max(height, 1e-9);
  let cols = Math.max(1, Math.round(Math.sqrt(targetCount * ratio)));
  let rows = Math.max(1, Math.round(targetCount / cols));

  const spacingX = Math.max(gridRaw, snapToGrid(width / Math.max(cols, 1), gridRaw));
  const spacingY = Math.max(gridRaw, snapToGrid(height / Math.max(rows, 1), gridRaw));

  const xStart = minX + Math.max(0, (width - spacingX * (cols - 1)) / 2);
  const yStart = minY + Math.max(0, (height - spacingY * (rows - 1)) / 2);

  const points = [];
  const keys = new Set();

  for (let i = 0; i < cols; i += 1) {
    for (let j = 0; j < rows; j += 1) {
      const candidate = {
        x: xStart + i * spacingX,
        y: yStart + j * spacingY,
      };

      if (!pointInPolygon(candidate, room.points)) continue;

      const key = pointKey(candidate);
      if (keys.has(key)) continue;

      keys.add(key);
      points.push(candidate);

      if (points.length >= targetCount) {
        return points;
      }
    }
  }

  // fallback: denser scan
  const fallbackStep = Math.max(gridRaw / 2, Math.min(width, height) / 20);
  for (let y = minY + fallbackStep / 2; y <= maxY - fallbackStep / 2; y += fallbackStep) {
    for (let x = minX + fallbackStep / 2; x <= maxX - fallbackStep / 2; x += fallbackStep) {
      const candidate = { x, y };
      if (!pointInPolygon(candidate, room.points)) continue;

      const key = pointKey(candidate);
      if (keys.has(key)) continue;

      keys.add(key);
      points.push(candidate);
      if (points.length >= targetCount) return points;
    }
  }

  // fallback: random fill
  let attempts = 0;
  const maxAttempts = targetCount * 200;
  while (points.length < targetCount && attempts < maxAttempts) {
    const candidate = {
      x: minX + Math.random() * width,
      y: minY + Math.random() * height,
    };

    if (!pointInPolygon(candidate, room.points)) {
      attempts += 1;
      continue;
    }

    const key = pointKey(candidate, 2);
    if (keys.has(key)) {
      attempts += 1;
      continue;
    }

    keys.add(key);
    points.push(candidate);
    attempts += 1;
  }

  return points;
}

function runCadLayout() {
  if (!cadState.rooms.length) {
    throw new Error("먼저 '방 자동 인식'을 실행해주세요.");
  }

  const circuits = parseCircuits(getCommonInputs().circuits);
  if (!circuits.length) circuits.push("1");

  const cadCfg = getCadConfig();
  let circuitIdx = 0;
  const placements = [];

  cadState.rooms.forEach((room) => {
    const pts = generateRoomPlacementPoints(room, room.recommendedCount, cadCfg);
    room.placedCount = pts.length;

    pts.forEach((pt) => {
      placements.push({
        x: pt.x,
        y: pt.y,
        roomId: room.id,
        roomName: room.name,
        circuit: circuits[circuitIdx % circuits.length],
      });
      circuitIdx += 1;
    });
  });

  cadState.placements = placements;
  renderCadRooms(cadState.rooms);

  const recTotal = cadState.rooms.reduce((sum, room) => sum + (room.recommendedCount || 0), 0);
  const placedTotal = cadState.rooms.reduce((sum, room) => sum + (room.placedCount || 0), 0);

  setCadStatus(
    `자동 배치 계산 완료\n` +
      `- 방 수: ${cadState.rooms.length}\n` +
      `- 추천 총수량: ${recTotal.toLocaleString("ko-KR")}\n` +
      `- 실제 배치수량: ${placedTotal.toLocaleString("ko-KR")}`,
  );
}

function ensureOutputLayer(lib, database, layerName) {
  const layers = database.layers;
  for (let i = 0, size = layers.size(); i < size; i += 1) {
    const layer = layers.get(i);
    if (layer?.name === layerName) return;
  }

  const layer = new lib.DRW_Layer();
  layer.name = layerName;
  layers.push_back(layer);
}

function appendPlacementEntities(lib, database, placements, cadCfg, common) {
  const block = database.mBlock;
  if (!block || !block.entities) {
    throw new Error("출력 대상 모델 공간을 찾지 못했습니다.");
  }

  ensureOutputLayer(lib, database, cadCfg.outputLayer);

  const radiusRaw = mmToRawLength(cadCfg.symbolRadiusMm, cadCfg.unit);
  const textOffsetRaw = mmToRawLength(cadCfg.textOffsetMm, cadCfg.unit);
  const textHeightRaw = mmToRawLength(CAD_DEFAULT_TEXT_HEIGHT_MM, cadCfg.unit);

  placements.forEach((item) => {
    const circle = new lib.DRW_Circle();
    circle.layer = cadCfg.outputLayer;
    circle.color = 2;
    circle.basePoint.x = item.x;
    circle.basePoint.y = item.y;
    circle.basePoint.z = 0;
    circle.radius = Math.max(radiusRaw, 1);
    block.entities.push_back(circle);

    const text = new lib.DRW_Text();
    text.layer = cadCfg.outputLayer;
    text.color = 1;
    text.text = item.circuit;
    text.style = common.fontStyle || "Standard";
    text.height = Math.max(textHeightRaw, 1);

    const tx = item.x + textOffsetRaw;
    const ty = item.y - textOffsetRaw;
    text.basePoint.x = tx;
    text.basePoint.y = ty;
    text.basePoint.z = 0;
    text.secPoint.x = tx;
    text.secPoint.y = ty;
    text.secPoint.z = 0;

    block.entities.push_back(text);
  });
}

function validateExportedDxf(lib, dxfContent) {
  const checkDb = new lib.DRW_Database();
  const checkHandler = new lib.DRW_FileHandler();
  checkHandler.database = checkDb;

  let ok = false;
  try {
    const reader = new lib.DRW_DxfRW(dxfContent);
    ok = reader.read(checkHandler, false);
    reader.delete();
  } catch (error) {
    ok = false;
  } finally {
    checkHandler.delete();
    checkDb.delete();
  }

  return ok;
}

async function analyzeCad() {
  if (!cadState.content || !cadState.extension) {
    await loadCadFileFromInput();
  }

  const common = getCommonInputs();
  const cadCfg = getCadConfig();

  setCadStatus("도면 분석 중...");

  let parsed;
  try {
    parsed = await parseCadDatabase(cadState.content, cadState.extension);
    const rooms = detectRooms(parsed.lib, parsed.database, common, cadCfg);
    cadState.rooms = rooms;
    cadState.placements = [];

    renderCadRooms(rooms);

    const recTotal = rooms.reduce((sum, room) => sum + room.recommendedCount, 0);
    setCadStatus(
      `도면 분석 완료\n` +
        `- 인식된 방 수: ${rooms.length}\n` +
        `- 추천 총수량: ${recTotal.toLocaleString("ko-KR")}\n` +
        `- 기준: 닫힌 폴리라인 + 면적 ${cadCfg.minAreaM2}~${cadCfg.maxAreaM2}㎡`,
    );
  } finally {
    if (parsed?.fileHandler) parsed.fileHandler.delete();
    if (parsed?.database) parsed.database.delete();
  }
}

async function exportCadResult() {
  if (!cadState.content || !cadState.extension || !cadState.file) {
    throw new Error("먼저 CAD 파일을 업로드해주세요.");
  }

  if (!cadState.rooms.length) {
    await analyzeCad();
  }

  if (!cadState.placements.length) {
    runCadLayout();
  }

  const cadCfg = getCadConfig();
  const common = getCommonInputs();

  setCadStatus("결과 DXF 생성 중...");

  let parsed;
  try {
    parsed = await parseCadDatabase(cadState.content, cadState.extension);
    appendPlacementEntities(parsed.lib, parsed.database, cadState.placements, cadCfg, common);

    const targetVersion = parsed.lib.DRW_Version.AC1015 || parsed.lib.DRW_Version.AC1021;
    const output = parsed.fileHandler.fileExport(
      targetVersion,
      false,
      parsed.database,
      false,
    );

    const isValid = validateExportedDxf(parsed.lib, output);
    if (!isValid) {
      throw new Error("생성된 DXF 파일 검증에 실패했습니다. 설정을 조정한 뒤 다시 시도해주세요.");
    }

    const baseName = cadState.file.name.replace(/\.[^.]+$/, "");
    const outputName = `${baseName}_lighting_auto.dxf`;
    const sizeMb = (toDownloadPayload(output).byteLength / (1024 * 1024)).toFixed(2);
    downloadFile(outputName, output, "application/dxf");

    setCadStatus(`결과 파일 생성 완료: ${outputName}\n파일 크기: ${sizeMb} MB\n다운로드를 확인해주세요.`);
  } finally {
    if (parsed?.fileHandler) parsed.fileHandler.delete();
    if (parsed?.database) parsed.database.delete();
  }
}

async function runWithButtonLock(button, action) {
  if (!button) {
    await action();
    return;
  }

  const prev = button.disabled;
  button.disabled = true;
  try {
    await action();
  } finally {
    button.disabled = prev;
  }
}

function resetDefaults() {
  els.floorHeight.value = "2700";
  els.lumens.value = "4400";
  els.maintenanceFactor.value = "0.9";
  els.circuits.value = "45, 46";
  els.ksClass.value = "G";
  els.ksLevel.value = "std";
  els.targetLux.value = "400";
  els.fontStyle.value = "Standard";
  els.uMode.value = "interp";

  els.roomWidth.value = "9000";
  els.roomHeight.value = "6000";
  els.roomArea.value = "";

  els.bulkInput.value = "회의실A,9000,6000,,400\n사무실1,12000,8000,,400\n복도,20000,2500,,200";
  bulkResults = [];
  renderBulkTable([]);

  els.cadUnit.value = "mm";
  els.cadMinArea.value = "3";
  els.cadMaxArea.value = "5000";
  els.cadGrid.value = "300";
  els.cadOutputLayer.value = "LIGHT_AUTO";
  els.cadSymbolRadius.value = "150";
  els.cadTextOffset.value = "300";
  els.cadFile.value = "";

  presetButtons.forEach((btn) => btn.classList.remove("active"));

  cadState.file = null;
  cadState.extension = "";
  cadState.content = null;
  cadState.rooms = [];
  cadState.placements = [];

  if (els.cadFileName) {
    els.cadFileName.textContent = "선택된 파일 없음";
  }

  renderCadRooms([]);
  setCadStatus("CAD 파일을 업로드한 뒤 ‘방 자동 인식’을 실행하세요.");
  validateAllInputs();
  calcSingle();
  calcBulk();
  generateLispCode();
}

function setupCollapsibleCards() {
  const cards = Array.from(document.querySelectorAll("main .card"));
  cards.forEach((card) => {
    const heading = card.querySelector(":scope > h2");
    if (!heading) return;

    let body = card.querySelector(":scope > .card-body");
    if (!body) {
      body = document.createElement("div");
      body.className = "card-body";
      const children = Array.from(card.children).filter((child) => child !== heading);
      children.forEach((child) => body.appendChild(child));
      card.appendChild(body);
    }

    heading.title = "클릭해서 접기/펼치기";
    heading.addEventListener("click", () => {
      card.classList.toggle("collapsed");
    });
  });
}

function setupTopNavigation() {
  navPills.forEach((pill) => {
    pill.addEventListener("click", () => {
      const targetId = pill.dataset.target;
      const target = document.getElementById(targetId);
      if (!target) return;

      target.classList.remove("collapsed");
      target.scrollIntoView({ behavior: "smooth", block: "start" });

      navPills.forEach((p) => p.classList.remove("active"));
      pill.classList.add("active");
    });
  });

  const sections = Array.from(document.querySelectorAll("main .card[id]"));
  const updateActiveByScroll = () => {
    let current = "";
    sections.forEach((section) => {
      if (section.getBoundingClientRect().top <= 120) {
        current = section.id;
      }
    });

    if (!current && sections.length) {
      current = sections[0].id;
    }

    navPills.forEach((pill) => {
      pill.classList.toggle("active", pill.dataset.target === current);
    });
  };

  window.addEventListener("scroll", updateActiveByScroll, { passive: true });
  updateActiveByScroll();
}

function setupScrollTopButton() {
  const btn = els.scrollTopBtn;
  if (!btn) return;

  const toggle = () => {
    btn.classList.toggle("visible", window.scrollY > 400);
  };

  window.addEventListener("scroll", toggle, { passive: true });
  toggle();

  btn.addEventListener("click", () => {
    window.scrollTo({ top: 0, behavior: "smooth" });
  });
}

function setupCadDropZone() {
  const zone = els.cadDropZone;
  if (!zone || !els.cadFile) return;

  zone.addEventListener("dragover", (event) => {
    event.preventDefault();
    zone.classList.add("drag-over");
  });

  zone.addEventListener("dragleave", () => {
    zone.classList.remove("drag-over");
  });

  zone.addEventListener("drop", (event) => {
    event.preventDefault();
    zone.classList.remove("drag-over");

    if (!event.dataTransfer?.files?.length) return;
    els.cadFile.files = event.dataTransfer.files;
    els.cadFile.dispatchEvent(new Event("change"));
  });
}

function bindEvents() {
  els.applyKsBtn.addEventListener("click", () => {
    applyKsLux();
    calcSingle();
    calcBulk();
    generateLispCode();
  });
  els.resetBtn.addEventListener("click", resetDefaults);

  presetButtons.forEach((button) => {
    button.addEventListener("click", (event) => {
      event.stopPropagation();
      const presetName = button.dataset.preset;
      applyPreset(presetName);
    });
  });

  els.calcSingleBtn.addEventListener("click", calcSingle);
  els.calcBulkBtn.addEventListener("click", calcBulk);
  els.downloadCsvBtn.addEventListener("click", downloadBulkCsv);

  els.genLispBtn.addEventListener("click", generateLispCode);
  els.downloadLispBtn.addEventListener("click", downloadLisp);

  els.ksClass.addEventListener("change", () => {
    applyKsLux();
    calcSingle();
    calcBulk();
    generateLispCode();
  });
  els.ksLevel.addEventListener("change", () => {
    applyKsLux();
    calcSingle();
    calcBulk();
    generateLispCode();
  });
  els.uMode.addEventListener("change", () => {
    calcSingle();
    calcBulk();
    generateLispCode();
  });

  const liveValidateInputs = Array.from(document.querySelectorAll("input[type='number'], textarea"));
  liveValidateInputs.forEach((input) => {
    input.addEventListener("input", () => {
      if (input.type === "number") {
        validateNumberInput(input);
      }
    });
  });

  const quickRecalcInputs = [
    els.floorHeight,
    els.lumens,
    els.maintenanceFactor,
    els.targetLux,
    els.roomWidth,
    els.roomHeight,
    els.roomArea,
  ];
  quickRecalcInputs.forEach((input) => {
    input.addEventListener("change", () => {
      calcSingle();
      calcBulk();
      generateLispCode();
    });
  });

  els.bulkInput.addEventListener("blur", calcBulk);

  els.cadFile.addEventListener("change", async () => {
    const name = els.cadFile.files?.[0]?.name || "선택된 파일 없음";
    if (els.cadFileName) {
      els.cadFileName.textContent = name;
    }

    try {
      await loadCadFileFromInput();
      renderCadRooms([]);
    } catch (error) {
      setCadStatus(`오류: ${error.message}`);
    }
  });

  els.cadAnalyzeBtn.addEventListener("click", async () => {
    await runWithButtonLock(els.cadAnalyzeBtn, async () => {
      try {
        await analyzeCad();
      } catch (error) {
        setCadStatus(`오류: ${error.message}`);
      }
    });
  });

  els.cadLayoutBtn.addEventListener("click", async () => {
    await runWithButtonLock(els.cadLayoutBtn, async () => {
      try {
        if (!cadState.rooms.length) {
          await analyzeCad();
        }
        runCadLayout();
      } catch (error) {
        setCadStatus(`오류: ${error.message}`);
      }
    });
  });

  els.cadExportBtn.addEventListener("click", async () => {
    await runWithButtonLock(els.cadExportBtn, async () => {
      try {
        await exportCadResult();
      } catch (error) {
        setCadStatus(`오류: ${error.message}`);
      }
    });
  });
}

(function init() {
  populateKsUi();
  setupCollapsibleCards();
  setupTopNavigation();
  setupScrollTopButton();
  setupCadDropZone();
  bindEvents();
  applyKsLux();
  validateAllInputs();

  if (els.singleResultPanel) {
    els.singleResultPanel.style.display = "none";
  }

  calcSingle();
  calcBulk();
  generateLispCode();
  renderCadRooms([]);
  setCadStatus("CAD 파일을 업로드한 뒤 ‘방 자동 인식’을 실행하세요.");
})();
