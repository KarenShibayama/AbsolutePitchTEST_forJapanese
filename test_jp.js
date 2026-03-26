// ====== Configuration ======
const NOTE_NAMES = ["C","C#","D","D#","E","F","F#","G","G#","A","A#","B"];
const SOLFEGE_NAMES = ["do","do#","re","re#","mi","fa","fa#","sol","sol#","la","la#","si"];
const DISPLAY_LABELS_SOLFEGE = {
  "C":  "ド",
  "C#": "ド♯/レ♭",
  "D":  "レ",
  "D#": "レ♯/ミ♭",
  "E":  "ミ",
  "F":  "ファ",
  "F#": "ファ♯/ソ♭",
  "G":  "ソ",
  "G#": "ソ♯/ラ♭",
  "A":  "ラ",
  "A#": "ラ♯/シ♭",
  "B":  "シ"
};
const LABEL_TO_FILE = {
  "C": "C",
  "C#": "Cis",
  "D": "D",
  "D#": "Dis",
  "E": "E",
  "F": "F",
  "F#": "Fis",
  "G": "G",
  "G#": "Gis",
  "A": "A",
  "A#": "Ais",
  "B": "B"
};
const DISPLAY_LABELS = {
  "C":  "C",
  "C#": "C♯/D♭",
  "D":  "D",
  "D#": "D♯/E♭",
  "E":  "E",
  "F":  "F",
  "F#": "F♯/G♭",
  "G":  "G",
  "G#": "G♯/A♭",
  "A":  "A",
  "A#": "A♯/B♭",
  "B":  "B"
};
const AUDIO_DIR = "audio";

// Adjust to your MIDI range
const MIN_MIDI = 36;  // 036-C2.wav
const MAX_MIDI = 95;  // adjust as needed

const N_TRIALS = 60;     // number of trials
const TRIAL_MS = 4000;         // fixed 4s from tone onset to next tone
const START_DELAY_MS = 5000;   // 5s delay after volume check OK


// ====== State ======
let trials = [];
let trialIndex = -1;
let current = null;
let tSoundOn = null;
let canRespond = false;
let results = [];
let ID = "";
// Label mode: "sharp" = C/C#, "solfege" = do/re/mi
let LABEL_MODE = "sharp";  // change to "solfege" if needed
let runId = null;

const elStatus = document.getElementById("status");
const btnStart = document.getElementById("btnStart");
const btnDownload = document.getElementById("btnDownload");
const btnPdf = document.getElementById("btnPdf");
const elID = document.getElementById("ID");
const elSex = document.getElementById("sex");
const elAge = document.getElementById("age");
const elInstRows = document.getElementById("instRows");
const btnAddInst = document.getElementById("btnAddInst");
const btnVolPlay = document.getElementById("btnVolPlay");
const btnVolOK   = document.getElementById("btnVolOK");
const elLabelModeChooser = document.getElementById("labelModeChooser");
const btnLabelModeStart = document.getElementById("btnLabelModeStart");
const elSummary = document.getElementById("summary");
const canvasAcc = document.getElementById("accChart");
const canvasRT = document.getElementById("rtChart");
const elKeyboard = document.getElementById("keyboard");
const audioBufferCache = new Map(); // midi -> AudioBuffer
const VOLUME_CHECK_MIDI = 69; // assumes an A4(440Hz)-equivalent file exists; change if needed

const isMobile = /iPhone|iPad|Android/i.test(navigator.userAgent);// --- Device detection ---
if (isMobile && btnDownload) { btnDownload.style.display = "none"; }

let inVolumeCheck = false;
let audioCtx = null;

function prettyLabel(s) {
  // Display-only conversion (not used for analysis)
  return s.replace(/#/g, "♯");
}

// ====== UI generation ======
function buildChoiceButtons() {
  const keyboard = document.getElementById("keyboard");
  if (!keyboard) {
    console.error('id="keyboard" not found');
    return;
  }
  keyboard.innerHTML = "";

  // White (7) and black (5) key layout for one octave
  const whiteKeys = ["C","D","E","F","G","A","B"];
  const blackKeys = [
    { note: "C#", leftBase: 0, offset: 0.69 }, // between C and D
    { note: "D#", leftBase: 1, offset: 0.71 }, // between D and E
    { note: "F#", leftBase: 3, offset: 0.69 }, // between F and G
    { note: "G#", leftBase: 4, offset: 0.70 }, // between G and A
    { note: "A#", leftBase: 5, offset: 0.71 }, // between A and B
  ];

  const labelFor = (noteSharp) => {
    if (LABEL_MODE === "sharp") {
      return DISPLAY_LABELS[noteSharp] ?? noteSharp;
    }
    return DISPLAY_LABELS_SOLFEGE[noteSharp] ?? noteSharp;
  };

  const pressFlash = (btn) => {
    btn.classList.add("pressed");
    setTimeout(() => btn.classList.remove("pressed"), 120);
  };

  // White keys
  whiteKeys.forEach((note, i) => {
    const w = document.createElement("button");
    w.type = "button";
    w.className = "key white";
    w.style.left = `calc((100% / 7) * ${i})`;

    const span = document.createElement("span");
    span.className = "label";
    w.appendChild(span);
    span.textContent = prettyLabel(labelFor(note));
    w.addEventListener("click", () => {
      pressFlash(w);
      handleResponse(note); // internal label: C/D/E...
    });

    keyboard.appendChild(w);
  });

  // Black keys
  blackKeys.forEach(({ note, leftBase, offset }) => {
    const b = document.createElement("button");
    b.type = "button";
    b.className = "key black";

    // Place black keys slightly to the right of white-key boundaries
    // Fine-tuned per key for a more realistic keyboard look
    b.style.left = `calc((100% / 7) * (${leftBase} + ${offset}))`;

    const span = document.createElement("span");
    span.className = "label";
    const blackLabel = labelFor(note);
    span.textContent = blackLabel;
    b.appendChild(span);

    b.addEventListener("click", () => {
      pressFlash(b);
      handleResponse(note); // internal label: C#...
    });

    keyboard.appendChild(b);
  });
}buildChoiceButtons();

// ====== MIDI -> note/octave & filename ======
function midiToPcOct(m) {
  const pcSharp = NOTE_NAMES[m % 12];          // internal label for correctness checking
  const solfege = SOLFEGE_NAMES[m % 12];       // display label (solfege)
  const pcFile  = LABEL_TO_FILE[pcSharp];      // file label for audio filenames
  const oct = Math.floor(m / 12) - 1;
  return { pc: pcSharp, solfege, pcFile, oct };
}

function midiToFilename(m) {
  const { pcFile, oct } = midiToPcOct(m);
  const num = String(m).padStart(3, "0");
  return `${num}-${pcFile}${oct}.wav`;
}

function filePathForMidi(m) {
  return `${AUDIO_DIR}/${midiToFilename(m)}`;
}

// ====== Audio playback ======
async function ensureAudioCtx() {
  if (!audioCtx) audioCtx = new (window.AudioContext || window.webkitAudioContext)();
  if (audioCtx.state !== "running") await audioCtx.resume();
}

async function getAudioBuffer(m) {
  if (audioBufferCache.has(m)) return audioBufferCache.get(m);

  await ensureAudioCtx();
  const url = filePathForMidi(m);
  const res = await fetch(url);
  if (!res.ok) throw new Error(`音声ファイルが見つかりません: ${url}`);
  const buf = await res.arrayBuffer();
  const audioBuf = await audioCtx.decodeAudioData(buf);

  audioBufferCache.set(m, audioBuf);
  return audioBuf;
}

async function playMidi(m) {
  const audioBuf = await getAudioBuffer(m);

  const src = audioCtx.createBufferSource();
  src.buffer = audioBuf;
  src.connect(audioCtx.destination);

  // Schedule slightly ahead to reduce click noise / timing jitter
  const startAt = audioCtx.currentTime + 0.03;
  src.start(startAt);

  return startAt; // return the scheduled audio onset time
}

// ====== Randomization (build trials from MIDI range) ======
function shuffle(a) {
  for (let i = a.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [a[i], a[j]] = [a[j], a[i]];
  }
}

function makeMidiPool() {
  const arr = [];
  for (let m = MIN_MIDI; m <= MAX_MIDI; m++) arr.push(m);
  return arr;
}


// Use all 60 tones (36..95) exactly once and return an order that satisfies adjacent diff >= 13
// ====== Randomization (build trials from MIDI range) ======
const MIN_INTERVAL = 13; // fixed constraint: one octave + semitone

function okAdj(a, b) {
  return (
    a !== b &&
    Math.abs(b - a) >= MIN_INTERVAL &&
    (a % 12) !== (b % 12) // forbid same pitch class (octave-equivalent adjacency)
  );
}

// Build one sequence for a pool (evens only / odds only), using each note exactly once
async function solveOnePool(pool, label) {
  // Precompute adjacency graph
  const neighbors = new Map();
  for (const a of pool) {
    neighbors.set(a, pool.filter(b => okAdj(a, b)));
  }

  // Choose random start among top-K nodes with fewest neighbors (avoid fixed patterns)
  const K = Math.min(8, pool.length);
  const sorted = pool.slice().sort((x, y) => neighbors.get(x).length - neighbors.get(y).length);
  const start = sorted[Math.floor(Math.random() * K)];

  const used = new Set([start]);
  const path = [start];

  let steps = 0;
  const YIELD_EVERY = 3000;

  async function dfs(curr) {
    if (path.length === pool.length) return true;

    steps++;
    if (steps % YIELD_EVERY === 0) {
      elStatus.textContent = `試行を作成しています...（${label}） step=${steps} / length=${path.length}`;
      await new Promise(r => setTimeout(r, 0));
    }

    // Next candidates: unused only
    const cand = neighbors.get(curr).filter(v => !used.has(v));

    // Heuristic: fewer remaining unused neighbors first
    cand.sort((a, b) => {
      const da = neighbors.get(a).filter(v => !used.has(v)).length;
      const db = neighbors.get(b).filter(v => !used.has(v)).length;
      return da - db;
    });

    // Lightly shuffle top candidates to avoid deterministic tie behavior
    const top = Math.min(6, cand.length);
    for (let i = top - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      [cand[i], cand[j]] = [cand[j], cand[i]];
    }

    for (const nxt of cand) {
      used.add(nxt);
      path.push(nxt);

      if (await dfs(nxt)) return true;

      path.pop();
      used.delete(nxt);
    }
    return false;
  }

  const ok = await dfs(start);
  if (!ok) throw new Error(`No valid sequence found (${label})`);

  return path;
}

// Use all 60 tones (36..95) exactly once
// Build evens, then odds, and only validate the junction
async function makeTrials(n) {
  const poolAll = makeMidiPool(); // 36..95
  if (n !== poolAll.length) {
    throw new Error(`This mode assumes "use every tone in range exactly once". n=${n}, pool=${poolAll.length}`);
  }

  const evenPool = poolAll.filter(m => m % 2 === 0);
  const oddPool  = poolAll.filter(m => m % 2 === 1);

  const MAX_TRIES = 200; // retry limit on failure

  for (let t = 1; t <= MAX_TRIES; t++) {
    elStatus.textContent = `試行を作成しています... ${t}/${MAX_TRIES}`;

    // Build even and odd sequences separately
    const evenPath = await solveOnePool(evenPool, "even");
    const oddPath  = await solveOnePool(oddPool,  "odd");

    // Check junction compatibility
    const a = evenPath[evenPath.length - 1];
    const b = oddPath[0];

    if (okAdj(a, b)) {
      const path = evenPath.concat(oddPath);

      // Convert to trial objects
      return path.map((m) => {
        const { pc, oct } = midiToPcOct(m);
        return { midi: m, target: pc, oct, file: midiToFilename(m) };
      });
    }
    // Retry if junction fails
  }

  throw new Error("Failed to satisfy junction constraints (retry limit reached).");
}

// ====== Task flow ======
async function startTest() {
    alreadySent = false;

    // One runId per session
    runId = `${Date.now()}_${Math.random().toString(16).slice(2)}`;

    if (!validateStartRequirements(true)) {
      return;
    }
    ID = (elID.value || "").trim();
  
    // Lock UI
    if (btnDownload) btnDownload.disabled = true;
    btnStart.disabled = true;
    elID.disabled = true;
    elStatus.classList.remove("compact");
    if (elLabelModeChooser) elLabelModeChooser.style.display = "none";
  
    // Enter volume check
    inVolumeCheck = true;
    btnVolPlay.disabled = false;
    btnVolOK.disabled = false;
  
    elStatus.textContent = "音量確認です。聞きやすい音量に調整してください。";
  }


  let trialTimeoutId = null;
  let respondedThisTrial = false;
  
  async function nextTrial() {

    if (trialTimeoutId) {
      clearTimeout(trialTimeoutId);
      trialTimeoutId = null;
    }
  
    trialIndex++;
  
    if (trialIndex >= trials.length) {
      await finishTest();
      return;
    }
  
    current = trials[trialIndex];
    respondedThisTrial = false;
    canRespond = false;
  
    elStatus.textContent = `${trialIndex + 1} / ${trials.length} 読み込み中...`;
  
    try {
  
      const startAt = await playMidi(current.midi);
  
      const nowCtx = audioCtx.currentTime;
      const msUntilStart = Math.max(0, (startAt - nowCtx) * 1000);
  
      // At audio onset
      setTimeout(() => {
  
        tSoundOn = performance.now();
        canRespond = true;
  
        elStatus.textContent =
          `${trialIndex + 1} / ${trials.length}  回答してください`;
  
        // Important: always advance 5s after tone onset
        trialTimeoutId = setTimeout(() => {
  
          if (!respondedThisTrial) {
  
            results.push({
              ID,
              trial: trialIndex + 1,
              midi: current.midi,
              file: current.file,
              target: current.target,
              target_solfege: midiToPcOct(current.midi).solfege,
              response: "",
              response_solfege: "",
              correct: 0,
              rt_ms: "",
              no_response: 1
            });
  
          }
  
          canRespond = false;
          nextTrial();
  
        }, TRIAL_MS);
  
      }, msUntilStart);
  
    } catch (e) {
  
      elStatus.textContent = String(e.message || e);
      updateStartButtonState();
      elID.disabled = false;
  
    }
  }

  function handleResponse(resp) {
    if (!canRespond || !current) return;
    if (respondedThisTrial) return; // one response per trial
  
    const rt = performance.now() - tSoundOn;
    const correct = resp === current.target;
  
    const responseIdx = NOTE_NAMES.indexOf(resp);
    const responseSolfege = responseIdx >= 0 ? SOLFEGE_NAMES[responseIdx] : "";
  
    results.push({
      ID,
      trial: trialIndex + 1,
      midi: current.midi,
      file: current.file,
      target: current.target,
      target_solfege: midiToPcOct(current.midi).solfege,
      response: resp,
      response_solfege: responseSolfege,
      correct: correct ? 1 : 0,
      rt_ms: Math.round(rt),
      no_response: 0
    });
  
    respondedThisTrial = true;
  
    // Do not advance immediately: trialTimeout advances after 5s
    elStatus.textContent = `${trialIndex + 1} / ${trials.length} お待ちください...`;
  }


async function finishTest() {
  const { nCorrect, total } = calcAccuracy();
  const accPct = (nCorrect / total) * 100;
  const accPctText = accPct.toFixed(1);
  
  elStatus.innerHTML = `
    <b>課題が終了しました。</b><br>
    正答数：<b>${nCorrect} / ${total}</b><br>
    正答率：<b>${accPctText}%</b><br>
    データを自動保存しています...
    `;
    updateStartButtonState();
    elID.disabled = false;
    
    // --- Combined chart (accuracy bars + reaction-time dots) ---
    const { labels, rates, totals } = calcAccuracyByPitchClass();
    const { meansSec, counts } = calcMeanRTByPitchClass();
    const rtValuesMs = results
      .filter(r => typeof r.trial === "number" && r.correct === 1 && r.rt_ms !== "" && r.rt_ms != null && !Number.isNaN(Number(r.rt_ms)))
      .map(r => Number(r.rt_ms));
    const meanRtMs = rtValuesMs.length ? Math.round(rtValuesMs.reduce((a, b) => a + b, 0) / rtValuesMs.length) : null;
    const noResponseN = results.filter(r => r.no_response === 1).length;
    const noResponsePct = Math.round((noResponseN / total) * 100);
    const bestIdx = rates.indexOf(Math.max(...rates));
    const rtCandidates = meansSec
      .map((v, i) => ({ idx: i, v, n: counts[i] }))
      .filter(x => x.n > 0);
    const fastest = rtCandidates.length ? rtCandidates.reduce((a, b) => (a.v <= b.v ? a : b)) : null;

    elSummary.innerHTML = `
      <div class="result-panel">
        <div class="result-title">結果概要</div>
        <div class="kpi-grid">
          <div class="kpi-card">
            <div class="kpi-label">正答率</div>
            <div class="kpi-value">${accPctText}%</div>
            <div class="kpi-sub">${nCorrect} / ${total}</div>
          </div>
          <div class="kpi-card">
            <div class="kpi-label">平均反応時間（正答のみ）</div>
            <div class="kpi-value">${meanRtMs == null ? "-" : `${(meanRtMs / 1000).toFixed(2)}秒`}</div>
            <div class="kpi-sub">${rtValuesMs.length} 試行</div>
          </div>
          <div class="kpi-card">
            <div class="kpi-label">無回答</div>
            <div class="kpi-value">${noResponsePct}%</div>
            <div class="kpi-sub">${noResponseN} / ${total}</div>
          </div>
        </div>
        <div class="insight-row">最も正答率が高かった音：<b>${labelForDisplay(labels[bestIdx])}</b> (${Math.round(rates[bestIdx] * 100)}%)${fastest ? ` / 最速反応：<b>${labelForDisplay(labels[fastest.idx])}</b> (${fastest.v.toFixed(2)}秒)` : ""}</div>
      </div>
    `;

    canvasAcc.style.display = "block";
    drawBarChartAccuracy(canvasAcc, labels, rates, totals, meansSec, counts);
    canvasRT.style.display = "none";

    // Redraw on resize (single handler assignment)
    window.onresize = () => {
      drawBarChartAccuracy(canvasAcc, labels, rates, totals, meansSec, counts);
    };

    const ok = await sendOnce();
    elStatus.innerHTML = `
      <b>課題が終了しました。</b><br>
      正答数：<b>${nCorrect} / ${total}</b><br>
      正答率：<b>${accPctText}%</b><br>
      ${ok ? "データは自動保存されました。" : "データの自動保存に失敗しました。通信環境を確認してください。"}
    `;
}

function calcAccuracy() {
  const nCorrect = results.filter(r => r.correct === 1).length;
  const total = N_TRIALS;
  return { nCorrect, total };
}

function calcAccuracyByPitchClass() {
  // 12-note aggregation: per target (correct count / presented count)
  const stat = {};
  NOTE_NAMES.forEach(n => stat[n] = { total: 0, correct: 0 });

  for (const r of results) {
    // Only numeric trial rows (ignore any summary/misc rows)
    if (typeof r.trial !== "number") continue;
    if (!r.target || !(r.target in stat)) continue;

    stat[r.target].total += 1;
    if (r.correct === 1) stat[r.target].correct += 1;
  }

  const labels = NOTE_NAMES.slice();
  const rates = labels.map(n => {
    const { total, correct } = stat[n];
    return total ? (correct / total) : 0;
  });
  const totals = labels.map(n => stat[n].total);
  return { labels, rates, totals };
}

function labelForDisplay(pcSharp) {
  if (LABEL_MODE === "sharp") return pcSharp;
  return DISPLAY_LABELS_SOLFEGE[pcSharp] ?? pcSharp;
}

function drawBarChartAccuracy(canvas, labelsSharp, rates, totals, meansSec = [], rtCounts = []) {
  // Match keyboard width
  const cssW = elKeyboard.getBoundingClientRect().width;
  const cssH = 280;

  const dpr = window.devicePixelRatio || 1;
  canvas.style.width = cssW + "px";
  canvas.style.height = cssH + "px";
  canvas.width = Math.round(cssW * dpr);
  canvas.height = Math.round(cssH * dpr);

  const ctx = canvas.getContext("2d");
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);

  ctx.clearRect(0, 0, cssW, cssH);

  // Padding
  const padL = 48, padR = 52, padT = 48, padB = 46;
  const plotW = cssW - padL - padR;
  const plotH = cssH - padT - padB;

  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, cssW, cssH);

  ctx.fillStyle = "#333";
  ctx.font = "bold 13px system-ui, sans-serif";
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";
  ctx.fillText("音高別の正答率と平均反応時間", padL + plotW / 2, 13);
  ctx.font = "12px system-ui, sans-serif";
  ctx.fillStyle = "#5f6b7a";
  ctx.textAlign = "left";
  ctx.fillText("棒: 正答率 (%)", padL, 31);
  ctx.textAlign = "right";
  ctx.fillText("点: 平均反応時間（秒、正答のみ）", padL + plotW + 36, 31);

  // Axis (0–100%)
  ctx.strokeStyle = "#333";
  ctx.lineWidth = 1;
  ctx.beginPath();
  ctx.moveTo(padL, padT);
  ctx.lineTo(padL, padT + plotH);
  ctx.lineTo(padL + plotW, padT + plotH);
  ctx.stroke();

  // Ticks
  ctx.fillStyle = "#333";
  ctx.font = "12px system-ui, sans-serif";
  ctx.textAlign = "right";
  ctx.textBaseline = "middle";
  [0, 0.25, 0.5, 0.75, 1].forEach(v => {
    const y = padT + plotH - v * plotH;
    ctx.strokeStyle = "#ddd";
    ctx.beginPath();
    ctx.moveTo(padL, y);
    ctx.lineTo(padL + plotW, y);
    ctx.stroke();

    ctx.fillStyle = "#333";
    ctx.fillText(`${Math.round(v*100)}%`, padL - 6, y + 1);
  });

  // Bar
  const n = labelsSharp.length;
  const gap = 6;
  const barW = Math.max(6, (plotW - gap * (n - 1)) / n);

  const accColorWhite = "#ffffff";
  const accColorSharp = "#8a8a8a";
  ctx.strokeStyle = "#333";

  const accLabelYs = Array(n).fill(null);
  for (let i = 0; i < n; i++) {
    const rate = rates[i];          // 0..1
    const x = padL + i * (barW + gap);
    const h = rate * plotH;
    const y = padT + plotH - h;

    const pc = labelsSharp[i];
    const black = pc.includes("#");

    ctx.fillStyle = black ? accColorSharp : accColorWhite;
    ctx.strokeStyle = "#333";

    ctx.fillRect(x, y, barW, h);
    ctx.strokeRect(x, y, barW, h);

    // X labels (respect current label mode)
    const lab = prettyLabel(labelForDisplay(labelsSharp[i]));
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";
    ctx.fillStyle = "#333";
    ctx.fillText(lab, x + barW / 2, padT + plotH + 14);

    if (rate > 0) {
      const accY = Math.max(padT + 8, y - 7);
      accLabelYs[i] = accY;
      ctx.font = "10px system-ui, sans-serif";
      ctx.fillStyle = "#444";
      ctx.fillText(`${(rate * 100).toFixed(1)}%`, x + barW / 2, accY);
      ctx.font = "12px system-ui, sans-serif";
    }
  }

  if (meansSec.length === n) {
    const yMaxRT = 4.0; // right axis fixed at 4.0s for reaction time

    // Right axis (reaction time)
    ctx.strokeStyle = "#333";
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(padL + plotW, padT);
    ctx.lineTo(padL + plotW, padT + plotH);
    ctx.stroke();

    ctx.fillStyle = "#555";
    ctx.font = "11px system-ui, sans-serif";
    ctx.textAlign = "left";
    ctx.textBaseline = "middle";
    [0, 1, 2, 3, 4].forEach(rt => {
      const y = padT + plotH - (rt / yMaxRT) * plotH;
      ctx.fillText(`${rt.toFixed(1)}s`, padL + plotW + 6, y + 1);
    });

    // Reaction time: dots only (prioritize readability and reduce overlap)
    for (let i = 0; i < n; i++) {
      if ((rtCounts[i] || 0) <= 0) {
        continue;
      }
      const x = padL + i * (barW + gap) + barW / 2;
      const y = padT + plotH - (Math.min(meansSec[i], yMaxRT) / yMaxRT) * plotH;

      ctx.fillStyle = "#555";
      ctx.beginPath();
      ctx.arc(x, y, 3, 0, Math.PI * 2);
      ctx.fill();

      // If overlapping same-note accuracy label, shift RT label up/down
      let rtLabelY = y - 8;
      if (accLabelYs[i] != null && Math.abs(rtLabelY - accLabelYs[i]) < 12) {
        rtLabelY = accLabelYs[i] - 12;
        if (rtLabelY < padT + 8) rtLabelY = accLabelYs[i] + 12;
      }
      if (rtLabelY > padT + plotH - 4) rtLabelY = padT + plotH - 4;

      ctx.font = "9px system-ui, sans-serif";
      ctx.textAlign = "center";
      ctx.fillText(`${meansSec[i].toFixed(1)}s`, x, rtLabelY);
      ctx.font = "12px system-ui, sans-serif";
    }
  }
}

async function volumePlay() {
    try {
      await playMidi(VOLUME_CHECK_MIDI);
    } catch (e) {
      elStatus.textContent = String(e.message || e);
    }
  }
  
  async function volumeOK() {
    if (!inVolumeCheck) return;
  
    inVolumeCheck = false;
    btnVolPlay.disabled = true;
    btnVolOK.disabled = true;
  
    results = [];
    elStatus.textContent = "試行を準備しています...";
  
    try {
      trials = await makeTrials(N_TRIALS);
    } catch (e) {
      elStatus.textContent = `試行の準備に失敗しました：${e.message}`;
      // Restore buttons so user can retry
      btnVolPlay.disabled = false;
      btnVolOK.disabled = false;
      inVolumeCheck = true;
      return;
    }
  
    trialIndex = -1;
    canRespond = false;
    showLabelModeChooser();
  }

function showLabelModeChooser() {
  if (!elLabelModeChooser) {
    startMainAfterLabelChoice();
    return;
  }
  elLabelModeChooser.style.display = "block";
  elStatus.classList.add("compact");
  const current = document.querySelector(`input[name="labelMode"][value="${LABEL_MODE}"]`);
  if (current) current.checked = true;
  elStatus.textContent = "音量確認が完了しました。";
}

async function startMainAfterLabelChoice() {
  if (elLabelModeChooser) elLabelModeChooser.style.display = "none";
  elStatus.classList.remove("compact");
  elStatus.textContent = "5秒後に本試行を開始します...";

  setTimeout(async () => {
    elStatus.textContent = "本試行を開始します。";
    await nextTrial();
  }, START_DELAY_MS);
}

function wireLabelModeChooser() {
  const radios = Array.from(document.querySelectorAll('input[name="labelMode"]'));
  radios.forEach(r => {
    r.addEventListener("change", () => {
      if (!r.checked) return;
      LABEL_MODE = r.value === "solfege" ? "solfege" : "sharp";
      buildChoiceButtons();
    });
  });

  if (btnLabelModeStart) {
    btnLabelModeStart.addEventListener("click", startMainAfterLabelChoice);
  }
}

function setStatus(msg, isError=false) {
  elStatus.textContent = msg;
  elStatus.style.color = isError ? "crimson" : "#111";
}

// ===== Prevent duplicate submission (send exactly once) =====
let alreadySent = false;

async function sendOnce() {
  console.log("sendOnce called", new Date().toISOString());

  if (alreadySent) {
    console.log("Already sent. Skip sending.");
    return true;
  }

  setStatus("データを送信しています... しばらくお待ちください。");

  const ok = await sendDataToGAS({
    ID,
    test: "AP_Test_v1",
    payload: {
      runId,
      demographics: getDemographics(),
      summary: {
        totalTrials: N_TRIALS,
        ...calcAccuracy()
      },
      results
    },
  });

  if (ok) {
    alreadySent = true;
    setStatus("データの自動保存が完了しました。");
  } else {
    setStatus("データの自動保存に失敗しました。通信環境を確認してください。", true);
  }

  return ok;
}

function getDemographics() {
  const sex = (elSex?.value || "").trim();
  const age = (elAge?.value || "").trim();

  // instruments: collect multiple rows as [{name,start,end}]
  const rows = Array.from(document.querySelectorAll("#instRows .instRow"));
  const instruments = rows.map(r => {
    const nameSel = r.querySelector(".instName");
    const startEl = r.querySelector(".instStart");
    const endEl   = r.querySelector(".instEnd");
    const otherEl = r.querySelector(".instOther");

    let name = (nameSel?.value || "").trim();
    if (name === "other") {
      name = (otherEl?.value || "").trim();
    }

    const start = (startEl?.value || "").trim();
    const endRaw = (endEl?.value || "").trim();
    const end = endRaw.toLowerCase() === "present" ? "present" : endRaw;

    return { name, startAge: start, endAge: end };
  }).filter(x => x.name || x.startAge || x.endAge); // drop empty rows

  return { sex, age, instruments };
}

function getInstrumentRowValues(row) {
  const nameSel = row.querySelector(".instName");
  const startEl = row.querySelector(".instStart");
  const endEl   = row.querySelector(".instEnd");
  const otherEl = row.querySelector(".instOther");

  const rawName = (nameSel?.value || "").trim();
  const other = (otherEl?.value || "").trim();
  const start = (startEl?.value || "").trim();
  const end = (endEl?.value || "").trim();
  const name = rawName === "other" ? other : rawName;

  return { rawName, name, other, start, end };
}

function isInstrumentRowComplete(v) {
  // non-musician does not require age inputs
  if (v.rawName === "non-musician") return !!v.name;
  return !!v.name && !!v.start && !!v.end;
}

function clearInvalidMarks() {
  document.querySelectorAll(".invalid-field").forEach(el => el.classList.remove("invalid-field"));
}

function markInvalid(el) {
  if (!el) return;
  el.classList.add("invalid-field");
}

function validateStartRequirements(showMessage = false) {
  clearInvalidMarks();

  const id = (elID?.value || "").trim();
  const sex = (elSex?.value || "").trim();
  const age = (elAge?.value || "").trim();
  const rows = Array.from(document.querySelectorAll("#instRows .instRow"));
  let msg = "";
  let valid = true;

  if (!id) {
    markInvalid(elID);
    valid = false;
    msg ||= "学籍番号を入力してください。";
  }
  if (!sex) {
    markInvalid(elSex);
    valid = false;
    msg ||= "性別を選択してください。";
  }
  if (!age) {
    markInvalid(elAge);
    valid = false;
    msg ||= "年齢を入力してください。";
  }
  if (!rows.length) {
    valid = false;
    msg ||= "音楽歴を入力してください。";
  }

  for (let i = 0; i < rows.length; i++) {
    const v = getInstrumentRowValues(rows[i]);
    const nameSel = rows[i].querySelector(".instName");
    const startEl = rows[i].querySelector(".instStart");
    const endEl = rows[i].querySelector(".instEnd");
    const otherEl = rows[i].querySelector(".instOther");
    const isBlank = !v.rawName && !v.start && !v.end && !v.other;
    const isComplete = isInstrumentRowComplete(v);

    // Row 1 is required; additional rows must be either fully blank or fully completed
    if (i === 0) {
      if (!isComplete) {
        if (!v.rawName) markInvalid(nameSel);
        if (v.rawName === "other" && !v.other) markInvalid(otherEl);
        if (v.rawName !== "non-musician" && !v.start) markInvalid(startEl);
        if (v.rawName !== "non-musician" && !v.end) markInvalid(endEl);
        valid = false;
        msg ||= (v.rawName === "non-musician")
          ? "1行目の楽器を選択してください。"
          : "1行目の音楽歴をすべて入力してください。";
      }
    } else if (!isBlank && !isComplete) {
      if (!v.rawName) markInvalid(nameSel);
      if (v.rawName === "other" && !v.other) markInvalid(otherEl);
      if (v.rawName !== "non-musician" && !v.start) markInvalid(startEl);
      if (v.rawName !== "non-musician" && !v.end) markInvalid(endEl);
      valid = false;
      msg ||= `${i + 1} 行目の音楽歴をすべて入力するか、空欄のままにしてください。`;
    }
  }

  if (showMessage && !valid) elStatus.textContent = msg;
  return valid;
}

function updateStartButtonState() {
  if (!btnStart) return;
  // Keep Start disabled during volume check
  if (inVolumeCheck) {
    btnStart.disabled = true;
    return;
  }
  btnStart.disabled = !validateStartRequirements(false);
}

function wireStartValidation() {
  [elID, elSex, elAge].forEach(el => {
    if (!el) return;
    el.addEventListener("input", updateStartButtonState);
    el.addEventListener("change", updateStartButtonState);
  });
  updateStartButtonState();
}

function wireInstrumentUI() {
  if (!btnAddInst) return;

  // Show free-text field only when "other" is selected
  const toggleOther = (row) => {
    const sel = row.querySelector(".instName");
    const other = row.querySelector(".instOther");
    const start = row.querySelector(".instStart");
    const end = row.querySelector(".instEnd");
    if (!sel || !other) return;
    const isOther = sel.value === "other";
    const isNonMusician = sel.value === "non-musician";
    other.style.display = isOther ? "inline-block" : "none";
    if (start) {
      start.disabled = isNonMusician;
      if (isNonMusician) start.value = "";
    }
    if (end) {
      end.disabled = isNonMusician;
      if (isNonMusician) end.value = "";
    }
  };

  // Attach handlers to the initial row as well
  document.querySelectorAll("#instRows .instRow").forEach(r => {
    const sel = r.querySelector(".instName");
    const other = r.querySelector(".instOther");
    const start = r.querySelector(".instStart");
    const end = r.querySelector(".instEnd");
    if (sel) sel.addEventListener("change", () => {
      toggleOther(r);
      updateStartButtonState();
    });
    if (other) other.addEventListener("input", updateStartButtonState);
    if (start) start.addEventListener("input", updateStartButtonState);
    if (end) end.addEventListener("input", updateStartButtonState);
    toggleOther(r);
  });

  btnAddInst.addEventListener("click", () => {
    const base = document.querySelector("#instRows .instRow");
    if (!base) return;

    const clone = base.cloneNode(true);
    // Clear values
    clone.querySelectorAll("input").forEach(i => i.value = "");
    clone.querySelectorAll("select").forEach(s => s.value = "");
    // Hide "other" input
    const otherInit = clone.querySelector(".instOther");
    if (otherInit) otherInit.style.display = "none";

    // Rebind change handlers
    const sel = clone.querySelector(".instName");
    const start = clone.querySelector(".instStart");
    const end = clone.querySelector(".instEnd");
    const other = clone.querySelector(".instOther");
    if (sel) sel.addEventListener("change", () => {
      const o = clone.querySelector(".instOther");
      if (o) o.style.display = (sel.value === "other") ? "inline-block" : "none";
      updateStartButtonState();
    });
    if (start) start.addEventListener("input", updateStartButtonState);
    if (end) end.addEventListener("input", updateStartButtonState);
    if (other) other.addEventListener("input", updateStartButtonState);

    elInstRows.appendChild(clone);
    updateStartButtonState();
  });
}

function calcMeanRTByPitchClass() {
  const stat = {};
  NOTE_NAMES.forEach(n => stat[n] = { sum: 0, n: 0 });

  for (const r of results) {
    if (typeof r.trial !== "number") continue;
    if (!r.target || !(r.target in stat)) continue;
    if (r.correct !== 1) continue; // RT mean is computed from correct trials only
    if (r.rt_ms === "" || r.rt_ms == null || Number.isNaN(Number(r.rt_ms))) continue;

    stat[r.target].sum += Number(r.rt_ms);
    stat[r.target].n += 1;
  }

  const labels = NOTE_NAMES.slice();
  const meansSec = labels.map(n => {
    const s = stat[n];
    return s.n ? (s.sum / s.n) / 1000 : 0;
  });
  const counts = labels.map(n => stat[n].n);
  return { labels, meansSec, counts };
}

function drawBarChartRT(canvas, labelsSharp, meansSec, counts) {
  const cssW = elKeyboard.getBoundingClientRect().width;
  const cssH = 220;

  const dpr = window.devicePixelRatio || 1;
  canvas.style.width = cssW + "px";
  canvas.style.height = cssH + "px";
  canvas.width = Math.round(cssW * dpr);
  canvas.height = Math.round(cssH * dpr);

  const ctx = canvas.getContext("2d");
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);

  ctx.clearRect(0, 0, cssW, cssH);
  ctx.fillStyle = "#fff";
  ctx.fillRect(0, 0, cssW, cssH);

  const padL = 46, padR = 10, padT = 28, padB = 40;
  const plotW = cssW - padL - padR;
  const plotH = cssH - padT - padB;

  ctx.fillStyle = "#333";
  ctx.font = "bold 13px system-ui, sans-serif";
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";
  ctx.fillText("音高別の平均反応時間", padL + plotW / 2, 12);
  ctx.font = "12px system-ui, sans-serif";

  const maxVal = Math.max(1.0, ...meansSec);
  const yMax = Math.ceil(maxVal * 10) / 10;

  ctx.strokeStyle = "#333";
  ctx.lineWidth = 1;
  ctx.beginPath();
  ctx.moveTo(padL, padT);
  ctx.lineTo(padL, padT + plotH);
  ctx.lineTo(padL + plotW, padT + plotH);
  ctx.stroke();

  const ticks = 4;
  ctx.fillStyle = "#333";
  ctx.font = "12px system-ui, sans-serif";
  for (let i = 0; i <= ticks; i++) {
    const v = (yMax / ticks) * i;
    const y = padT + plotH - (v / yMax) * plotH;
    ctx.strokeStyle = "#ddd";
    ctx.beginPath();
    ctx.moveTo(padL, y);
    ctx.lineTo(padL + plotW, y);
    ctx.stroke();
    ctx.fillStyle = "#333";
    ctx.fillText(`${v.toFixed(1)}s`, 4, y + 4);
  }

  const n = labelsSharp.length;
  const gap = 6;
  const barW = Math.max(6, (plotW - gap * (n - 1)) / n);

  for (let i = 0; i < n; i++) {
    const val = meansSec[i];
    const x = padL + i * (barW + gap);
    const h = (val / yMax) * plotH;
    const y = padT + plotH - h;

    const pc = labelsSharp[i];
    const black = pc.includes("#");

    ctx.fillStyle = black ? "#888" : "#fff";
    ctx.strokeStyle = "#333";
    ctx.fillRect(x, y, barW, h);
    ctx.strokeRect(x, y, barW, h);

    const lab = prettyLabel(labelForDisplay(labelsSharp[i]));
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";
    ctx.fillStyle = "#333";
    ctx.fillText(lab, x + barW / 2, padT + plotH + 14);

    ctx.fillStyle = "#333";
    ctx.font = "10px system-ui, sans-serif";
    ctx.fillText(`n=${counts[i]}`, x + 1, padT + plotH + 32);
    ctx.font = "12px system-ui, sans-serif";
  }
}

// ===== Instruments UI (minimal) =====
window.addInstrumentRow = function addInstrumentRow() {
  const elInstRows = document.getElementById("instRows");
  if (!elInstRows) return;

  elInstRows.insertAdjacentHTML("beforeend", `
    <div class="row instRow" style="margin:0;">
      <select class="instName" onchange="toggleOtherInput(this)">
        <option value="">選択してください</option>
        <option value="non-musician">演奏歴なし</option>
        <option value="piano">ピアノ</option>
        <option value="violin">バイオリン</option>
        <option value="flute">フルート</option>
        <option value="clarinet">クラリネット</option>
        <option value="saxophone">サクソフォン</option>
        <option value="trumpet">トランペット</option>
        <option value="trombone">トロンボーン</option>
        <option value="cello">チェロ</option>
        <option value="guitar">ギター</option>
        <option value="voice">声楽</option>
        <option value="percussion">打楽器</option>
        <option value="other">その他</option>
      </select>

      <input class="instStart" type="number" min="0" max="120" placeholder="開始年齢" style="width:120px;" />
      <input class="instEnd" type="number" min="0" max="120" placeholder="終了年齢" style="width:120px;" />
      <input class="instOther" type="text" placeholder="その他の楽器名" style="width:180px; display:none;" />
    </div>
  `);
};

  // Show free-text input only when "other" is selected
  window.toggleOtherInput = function toggleOtherInput(sel) {
    const row = sel.closest(".instRow");
    const other = row.querySelector(".instOther");
    const start = row.querySelector(".instStart");
    const end = row.querySelector(".instEnd");
    if (!other) return;
    other.style.display = (sel.value === "other") ? "inline-block" : "none";
    const isNonMusician = sel.value === "non-musician";
    if (start) {
      start.disabled = isNonMusician;
      if (isNonMusician) start.value = "";
    }
    if (end) {
      end.disabled = isNonMusician;
      if (isNonMusician) end.value = "";
    }
    updateStartButtonState();
  };

btnStart.addEventListener("click", startTest);
btnVolPlay.addEventListener("click", volumePlay);
btnVolOK.addEventListener("click", volumeOK);
wireInstrumentUI();
wireLabelModeChooser();
wireStartValidation();
