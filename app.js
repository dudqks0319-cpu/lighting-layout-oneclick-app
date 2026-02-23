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

const els = {
  floorHeight: document.getElementById("floorHeight"),
  lumens: document.getElementById("lumens"),
  maintenanceFactor: document.getElementById("maintenanceFactor"),
  circuits: document.getElementById("circuits"),
  ksClass: document.getElementById("ksClass"),
  ksLevel: document.getElementById("ksLevel"),
  targetLux: document.getElementById("targetLux"),
  fontStyle: document.getElementById("fontStyle"),

  applyKsBtn: document.getElementById("applyKsBtn"),
  resetBtn: document.getElementById("resetBtn"),

  roomWidth: document.getElementById("roomWidth"),
  roomHeight: document.getElementById("roomHeight"),
  roomArea: document.getElementById("roomArea"),
  calcSingleBtn: document.getElementById("calcSingleBtn"),
  singleOutput: document.getElementById("singleOutput"),

  bulkInput: document.getElementById("bulkInput"),
  calcBulkBtn: document.getElementById("calcBulkBtn"),
  downloadCsvBtn: document.getElementById("downloadCsvBtn"),
  bulkTableBody: document.querySelector("#bulkTable tbody"),
  bulkTotal: document.getElementById("bulkTotal"),

  genLispBtn: document.getElementById("genLispBtn"),
  downloadLispBtn: document.getElementById("downloadLispBtn"),
  lispOutput: document.getElementById("lispOutput"),

  ksSummary: document.getElementById("ksSummary"),
};

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

function getUByRi(ri) {
  if (ri >= 4.0) return 0.8;
  if (ri >= 2.5) return 0.75;
  if (ri >= 1.5) return 0.65;
  if (ri >= 1.0) return 0.55;
  return 0.45;
}

function computeLighting({ floorHeight, lux, lumens, maintenanceFactor, widthMm, heightMm, areaM2 }) {
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
  const u = getUByRi(ri);
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
  };
}

function calcSingle() {
  const common = getCommonInputs();
  const widthMm = toNum(els.roomWidth.value);
  const heightMm = toNum(els.roomHeight.value);

  const areaManual = toNum(els.roomArea.value, NaN);
  const areaM2 = Number.isFinite(areaManual) && areaManual > 0 ? areaManual : (widthMm * heightMm) / 1_000_000;

  try {
    const result = computeLighting({
      floorHeight: common.floorHeight,
      lux: common.targetLux,
      lumens: common.lumens,
      maintenanceFactor: common.maintenanceFactor,
      widthMm,
      heightMm,
      areaM2,
    });

    const txt = [
      `[단일 실 계산 결과]`,
      `면적: ${fmt(areaM2, 2)} ㎡`,
      `유효높이(Hm): ${fmt(result.hEff, 2)} m`,
      `실지수(RI): ${fmt(result.ri, 2)}`,
      `이용률(U): ${fmt(result.u, 2)}`,
      `계산 수량(실수): ${fmt(result.rawCount, 2)} 개`,
      `추천 수량(올림): ${result.recommendedCount.toLocaleString("ko-KR")} 개`,
    ].join("\n");

    els.singleOutput.textContent = txt;
  } catch (error) {
    els.singleOutput.textContent = `오류: ${error.message}`;
  }
}

function parseCsvLine(line) {
  const cols = line.split(",").map((s) => s.trim());
  if (cols.length < 3) return null;

  const name = cols[0] || `Room-${Math.random().toString(36).slice(2, 6)}`;
  const widthMm = toNum(cols[1], NaN);
  const heightMm = toNum(cols[2], NaN);
  const areaM2Raw = toNum(cols[3], NaN);
  const luxRaw = toNum(cols[4], NaN);

  if (!Number.isFinite(widthMm) || !Number.isFinite(heightMm)) return null;

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
  const lines = els.bulkInput.value
    .split(/\r?\n/)
    .map((v) => v.trim())
    .filter(Boolean);

  const rows = [];
  for (const line of lines) {
    const item = parseCsvLine(line);
    if (!item) continue;

    const areaM2 = Number.isFinite(item.areaM2Raw) && item.areaM2Raw > 0
      ? item.areaM2Raw
      : (item.widthMm * item.heightMm) / 1_000_000;

    const lux = Number.isFinite(item.luxRaw) && item.luxRaw > 0 ? item.luxRaw : common.targetLux;

    try {
      const result = computeLighting({
        floorHeight: common.floorHeight,
        lux,
        lumens: common.lumens,
        maintenanceFactor: common.maintenanceFactor,
        widthMm: item.widthMm,
        heightMm: item.heightMm,
        areaM2,
      });

      rows.push({
        name: item.name,
        areaM2,
        ri: result.ri,
        u: result.u,
        lux,
        count: result.recommendedCount,
      });
    } catch (error) {
      rows.push({
        name: item.name,
        areaM2,
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

function renderBulkTable(rows) {
  els.bulkTableBody.innerHTML = "";
  let total = 0;

  for (const row of rows) {
    const tr = document.createElement("tr");
    const isError = typeof row.count !== "number";
    if (!isError) total += row.count;

    tr.innerHTML = `
      <td>${row.name}</td>
      <td>${Number.isFinite(row.areaM2) ? fmt(row.areaM2, 2) : "-"}</td>
      <td>${Number.isFinite(row.ri) ? fmt(row.ri, 2) : "-"}</td>
      <td>${Number.isFinite(row.u) ? fmt(row.u, 2) : "-"}</td>
      <td>${Number.isFinite(row.lux) ? fmt(row.lux, 0) : "-"}</td>
      <td>${isError ? row.count : row.count.toLocaleString("ko-KR")}</td>
    `;

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
    "(defun qqq_u_by_ri (ri)",
    "  (cond ((>= ri 4.0) 0.80)",
    "        ((>= ri 2.5) 0.75)",
    "        ((>= ri 1.5) 0.65)",
    "        ((>= ri 1.0) 0.55)",
    "        (t 0.45)))",
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

function downloadFile(filename, text, mimeType) {
  const blob = new Blob([text], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
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

  els.roomWidth.value = "9000";
  els.roomHeight.value = "6000";
  els.roomArea.value = "";

  els.singleOutput.textContent = "";
  els.bulkInput.value = "회의실A,9000,6000,,400\n사무실1,12000,8000,,400\n복도,20000,2500,,200";
  bulkResults = [];
  renderBulkTable([]);
  els.lispOutput.value = "";
}

function bindEvents() {
  els.applyKsBtn.addEventListener("click", applyKsLux);
  els.resetBtn.addEventListener("click", resetDefaults);

  els.calcSingleBtn.addEventListener("click", calcSingle);
  els.calcBulkBtn.addEventListener("click", calcBulk);
  els.downloadCsvBtn.addEventListener("click", downloadBulkCsv);

  els.genLispBtn.addEventListener("click", generateLispCode);
  els.downloadLispBtn.addEventListener("click", downloadLisp);

  els.ksClass.addEventListener("change", applyKsLux);
  els.ksLevel.addEventListener("change", applyKsLux);
}

(function init() {
  populateKsUi();
  bindEvents();
  applyKsLux();
  calcSingle();
  calcBulk();
  generateLispCode();
})();
