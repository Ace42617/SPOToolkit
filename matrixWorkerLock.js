// Full-tab lock overlay for permissions matrix worker tab (progress + mini-game).
(function () {
  const OVERLAY_ID = "spot-matrix-worker-lock";
  const BRAND = "#37ae1c";
  const BRAND_LIGHT = "#86efac";

  let overlayEl = null;
  let barEl = null;
  let pctEl = null;
  let msgEl = null;
  let logEl = null;
  let statusEl = null;
  let gameLoopId = null;
  let snakeRafId = null;
  let snake = null;
  let snakeUserHighScore = 0;
  let snakeAutoHighScore = 0;
  let snakeWindowKeyHandler = null;

  function formatScoreLine(score) {
    return "Score: " + score + " | Auto High: " + snakeAutoHighScore + " | Your High: " + snakeUserHighScore;
  }

  function updateSnakeScoreDisplay(scoreEl, score) {
    if (scoreEl) scoreEl.textContent = formatScoreLine(score);
  }

  function recordSnakeHighOnFail(failedMode, finalScore) {
    if (failedMode === "auto" && finalScore > snakeAutoHighScore) snakeAutoHighScore = finalScore;
    if (failedMode === "manual" && finalScore > snakeUserHighScore) snakeUserHighScore = finalScore;
  }

  const CSS =
    "#" + OVERLAY_ID + "{position:fixed;inset:0;z-index:2147483646;background:linear-gradient(160deg,#0f172a 0%,#111827 45%,#1a2332 100%);color:#e6edf6;font-family:'Segoe UI',system-ui,sans-serif;display:flex;align-items:center;justify-content:center;padding:24px;box-sizing:border-box;overflow:auto;}" +
    "#" + OVERLAY_ID + " .spot-mwl-card{max-width:520px;width:100%;background:rgba(255,255,255,.04);border:1px solid rgba(255,255,255,.12);border-radius:16px;padding:22px 22px 18px;box-shadow:0 24px 64px rgba(0,0,0,.45);}" +
    "#" + OVERLAY_ID + " .spot-mwl-title{margin:0 0 8px;font-size:20px;font-weight:700;letter-spacing:-.02em;}" +
    "#" + OVERLAY_ID + " .spot-mwl-warn{margin:0 0 6px;font-size:13px;font-weight:600;color:" + BRAND_LIGHT + ";}" +
    "#" + OVERLAY_ID + " .spot-mwl-hint{margin:0 0 16px;font-size:12px;line-height:1.45;color:rgba(230,237,246,.72);}" +
    "#" + OVERLAY_ID + " .spot-mwl-bar-wrap{height:10px;background:rgba(255,255,255,.08);border-radius:999px;overflow:hidden;margin-bottom:10px;}" +
    "#" + OVERLAY_ID + " .spot-mwl-bar{height:100%;width:0;background:linear-gradient(90deg,#5fd247," + BRAND + ");transition:width .3s ease;border-radius:999px;}" +
    "#" + OVERLAY_ID + " .spot-mwl-head{display:flex;justify-content:space-between;align-items:center;gap:12px;margin-bottom:8px;font-size:13px;font-weight:600;}" +
    "#" + OVERLAY_ID + " .spot-mwl-pct{color:" + BRAND_LIGHT + ";font-variant-numeric:tabular-nums;}" +
    "#" + OVERLAY_ID + " .spot-mwl-msg{min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;color:" + BRAND_LIGHT + ";}" +
    "#" + OVERLAY_ID + " .spot-mwl-log{max-height:88px;overflow:auto;font:11px/1.4 Consolas,ui-monospace,monospace;color:rgba(230,237,246,.65);white-space:pre-wrap;margin:0 0 14px;padding:8px 10px;background:rgba(0,0,0,.22);border-radius:8px;}" +
    "#" + OVERLAY_ID + " .spot-mwl-game{border-top:1px solid rgba(255,255,255,.1);padding-top:14px;margin-top:4px;}" +
    "#" + OVERLAY_ID + " .spot-mwl-game-head{text-align:center;margin:0 0 10px;}" +
    "#" + OVERLAY_ID + " .spot-mwl-game-title{margin:0 0 4px;font-size:12px;font-weight:500;color:rgba(230,237,246,.85);line-height:1.4;}" +
    "#" + OVERLAY_ID + " .spot-mwl-game-name{font-weight:800;font-size:13px;letter-spacing:.06em;text-transform:uppercase;background:linear-gradient(90deg," + BRAND_LIGHT + "," + BRAND + ");-webkit-background-clip:text;background-clip:text;color:transparent;}" +
    "#" + OVERLAY_ID + " .spot-mwl-game-sub{margin:0;font-size:11px;color:rgba(230,237,246,.58);line-height:1.35;}" +
    "#" + OVERLAY_ID + " .spot-mwl-crt{display:flex;justify-content:center;}" +
    "#" + OVERLAY_ID + " .spot-mwl-crt-bezel{position:relative;padding:3px 3px 5px;background:linear-gradient(180deg,#32363c 0%,#1e2126 55%,#121418 100%);border-radius:8px;box-shadow:0 4px 14px rgba(0,0,0,.45),inset 0 1px 0 rgba(255,255,255,.08);border:1px solid #0b0d10;}" +
    "#" + OVERLAY_ID + " .spot-mwl-crt-bezel::before{content:'';position:absolute;top:2px;left:50%;transform:translateX(-50%);width:4px;height:4px;border-radius:50%;background:#101215;box-shadow:inset 0 1px 2px rgba(0,0,0,.75);opacity:.7;}" +
    "#" + OVERLAY_ID + " .spot-mwl-crt-bezel::after{content:'SPO-9000';position:absolute;bottom:1px;right:4px;font:600 6px/1 Consolas,ui-monospace,monospace;letter-spacing:.1em;color:rgba(134,239,172,.22);}" +
    "#" + OVERLAY_ID + " .spot-mwl-crt-screen{position:relative;overflow:hidden;border-radius:4px;box-shadow:inset 0 0 18px rgba(0,0,0,.88),0 0 14px rgba(55,174,28,.18);}" +
    "#" + OVERLAY_ID + " .spot-mwl-crt-screen canvas{display:block;border:none;border-radius:0;background:#020806;image-rendering:pixelated;image-rendering:crisp-edges;}" +
    "#" + OVERLAY_ID + " .spot-mwl-crt-scanlines{pointer-events:none;position:absolute;inset:0;z-index:2;background:repeating-linear-gradient(to bottom,rgba(0,0,0,0) 0,rgba(0,0,0,0) 2px,rgba(0,0,0,.26) 2px,rgba(0,0,0,.26) 4px);mix-blend-mode:multiply;animation:spot-mwl-crt-flicker 5s steps(1,end) infinite;}" +
    "#" + OVERLAY_ID + " .spot-mwl-crt-vignette{pointer-events:none;position:absolute;inset:0;z-index:3;background:radial-gradient(ellipse at center,transparent 52%,rgba(0,0,0,.62) 100%);}" +
    "#" + OVERLAY_ID + " .spot-mwl-crt-glare{pointer-events:none;position:absolute;inset:0;z-index:4;background:linear-gradient(135deg,rgba(255,255,255,.07) 0%,transparent 32%,transparent 100%);}" +
    "#" + OVERLAY_ID + " .spot-mwl-crt-noise{pointer-events:none;position:absolute;inset:0;z-index:5;opacity:.05;background-image:url(\"data:image/svg+xml,%3Csvg viewBox='0 0 256 256' xmlns='http://www.w3.org/2000/svg'%3E%3Cfilter id='n'%3E%3CfeTurbulence type='fractalNoise' baseFrequency='.95' numOctaves='3' stitchTiles='stitch'/%3E%3C/filter%3E%3Crect width='100%25' height='100%25' filter='url(%23n)'/%3E%3C/svg%3E\");mix-blend-mode:soft-light;animation:spot-mwl-crt-flicker 5s steps(1,end) infinite;}" +
    "@keyframes spot-mwl-crt-flicker{0%,100%{opacity:.88;}48%{opacity:.94;}49%{opacity:.78;}50%{opacity:.92;}92%{opacity:.9;}93%{opacity:.82;}}" +
    "#" + OVERLAY_ID + " .spot-mwl-score{margin:10px 0 0;font:11px/1.4 Consolas,ui-monospace,monospace;color:#5fd247;text-align:center;text-shadow:0 0 10px rgba(95,210,71,.45),0 0 2px rgba(134,239,172,.35);letter-spacing:.02em;}" +
    "#" + OVERLAY_ID + " .spot-mwl-actions{display:flex;gap:10px;margin-top:14px;justify-content:flex-end;}" +
    "#" + OVERLAY_ID + " .spot-mwl-cancel{padding:6px 12px;border-radius:8px;border:1px solid rgba(255,255,255,.2);background:rgba(255,255,255,.06);color:#e6edf6;font-size:12px;cursor:pointer;font-family:inherit;}" +
    "#" + OVERLAY_ID + " .spot-mwl-cancel:hover{background:rgba(255,255,255,.12);}" +
    "#" + OVERLAY_ID + ".spot-mwl-done .spot-mwl-warn{color:#86efac;}" +
    "#" + OVERLAY_ID + ".spot-mwl-failed .spot-mwl-warn{color:#f87171;}" +
    "#" + OVERLAY_ID + " .spot-mwl-close-note{margin:10px 0 0;font-size:11px;color:rgba(230,237,246,.55);text-align:center;}";

  function injectStyles() {
    if (document.getElementById("spot-mwl-styles")) return;
    const s = document.createElement("style");
    s.id = "spot-mwl-styles";
    s.textContent = CSS;
    (document.head || document.documentElement).appendChild(s);
  }

  function stopSnake() {
    if (gameLoopId) {
      clearInterval(gameLoopId);
      gameLoopId = null;
    }
    if (snakeRafId) {
      cancelAnimationFrame(snakeRafId);
      snakeRafId = null;
    }
    if (snakeWindowKeyHandler) {
      window.removeEventListener("keydown", snakeWindowKeyHandler, true);
      snakeWindowKeyHandler = null;
    }
    snake = null;
  }

  function startSnake(canvas) {
    stopSnake();
    const ctx = canvas.getContext("2d");
    const cols = 16;
    const rows = 16;
    const TICK_MS = 140;
    const AUTO_IDLE_MS = Math.round(0.8 * cols * TICK_MS);
    const AUTO_LABEL_PULSE_MS = 2200;
    const cell = Math.floor(Math.min(canvas.width / cols, canvas.height / rows));
    const ox = Math.floor((canvas.width - cell * cols) / 2);
    const oy = Math.floor((canvas.height - cell * rows) / 2);
    let dir = { x: 1, y: 0 };
    let nextDir = { x: 1, y: 0 };
    let body = [{ x: 4, y: 8 }, { x: 3, y: 8 }, { x: 2, y: 8 }];
    let food = { x: 10, y: 8 };
    let score = 0;
    let mode = "auto";
    let lastUserInputAt = 0;
    const scoreEl = overlayEl && overlayEl.querySelector(".spot-mwl-score");

    function randFood() {
      for (;;) {
        const p = { x: Math.floor(Math.random() * cols), y: Math.floor(Math.random() * rows) };
        if (!body.some(function (b) { return b.x === p.x && b.y === p.y; })) return p;
      }
    }

    function willCollide(nx, ny, eating) {
      if (nx < 0 || ny < 0 || nx >= cols || ny >= rows) return true;
      var limit = eating ? body.length : body.length - 1;
      for (var i = 0; i < limit; i++) {
        if (body[i].x === nx && body[i].y === ny) return true;
      }
      return false;
    }

    function pickAutoDir() {
      var head = body[0];
      var moves = [
        { x: 0, y: -1 },
        { x: 0, y: 1 },
        { x: -1, y: 0 },
        { x: 1, y: 0 }
      ];
      var valid = [];
      for (var i = 0; i < moves.length; i++) {
        var m = moves[i];
        if (m.x === -dir.x && m.y === -dir.y) continue;
        var nx = head.x + m.x;
        var ny = head.y + m.y;
        var eating = nx === food.x && ny === food.y;
        if (!willCollide(nx, ny, eating)) valid.push(m);
      }
      if (!valid.length) return dir;
      valid.sort(function (a, b) {
        var da = Math.abs(head.x + a.x - food.x) + Math.abs(head.y + a.y - food.y);
        var db = Math.abs(head.x + b.x - food.x) + Math.abs(head.y + b.y - food.y);
        return da - db;
      });
      return valid[0];
    }

    function resetSnake() {
      body = [{ x: 4, y: 8 }, { x: 3, y: 8 }, { x: 2, y: 8 }];
      dir = nextDir = { x: 1, y: 0 };
      score = 0;
      food = randFood();
      updateSnakeScoreDisplay(scoreEl, score);
    }

    function draw() {
      var phosphor = mode === "auto" ? "#7dff9a" : "#5dff67";
      var bg = "#020806";

      ctx.fillStyle = bg;
      ctx.fillRect(0, 0, canvas.width, canvas.height);

      ctx.strokeStyle = "rgba(55,174,28,.09)";
      ctx.lineWidth = 1;
      for (var gx = 0; gx <= cols; gx++) {
        ctx.beginPath();
        ctx.moveTo(ox + gx * cell + 0.5, oy + 0.5);
        ctx.lineTo(ox + gx * cell + 0.5, oy + rows * cell + 0.5);
        ctx.stroke();
      }
      for (var gy = 0; gy <= rows; gy++) {
        ctx.beginPath();
        ctx.moveTo(ox + 0.5, oy + gy * cell + 0.5);
        ctx.lineTo(ox + cols * cell + 0.5, oy + gy * cell + 0.5);
        ctx.stroke();
      }

      var bodyLen = body.length;
      for (var si = bodyLen - 1; si >= 0; si--) {
        var seg = body[si];
        var t = bodyLen > 1 ? si / (bodyLen - 1) : 0;
        var fade = 1 - t;
        var fadeCurve = fade * fade * (3 - 2 * fade);
        var alpha = 0.1 + fadeCurve * 0.9;
        var segSat = Math.round(48 + fadeCurve * 52);
        var segLight = Math.round(14 + fadeCurve * 50);
        var segBlur = 1 + fadeCurve * 11;
        var inset = 1 + Math.round(t * 1.2);
        var segSize = Math.max(2, cell - 2 - Math.round(t * 1.4));
        var sx = ox + seg.x * cell + inset;
        var sy = oy + seg.y * cell + inset;

        ctx.shadowColor = phosphor;
        ctx.globalAlpha = alpha * 0.4;
        ctx.shadowBlur = segBlur + fadeCurve * 8;
        ctx.fillStyle = "hsl(118, " + Math.round(segSat * 0.75) + "%, " + Math.round(segLight * 0.55) + "%)";
        ctx.fillRect(sx, sy, segSize, segSize);

        ctx.globalAlpha = alpha * (0.55 + fadeCurve * 0.45);
        ctx.shadowBlur = segBlur;
        ctx.fillStyle = "hsl(118, " + segSat + "%, " + segLight + "%)";
        ctx.fillRect(sx, sy, segSize, segSize);
      }
      ctx.globalAlpha = 1;
      ctx.shadowBlur = 0;
      ctx.shadowBlur = 14;
      ctx.fillStyle = BRAND_LIGHT;
      ctx.fillRect(ox + food.x * cell + 2, oy + food.y * cell + 2, cell - 4, cell - 4);
      ctx.shadowBlur = 0;

      if (mode === "auto") {
        var pulse = 0.5 + 0.5 * Math.sin((Date.now() * Math.PI * 2) / AUTO_LABEL_PULSE_MS);
        var autoAlpha = 0.38 + pulse * 0.62;
        var letterSat = Math.round(52 + pulse * 48);
        var letterLight = Math.round(26 + pulse * 44);
        var glowSat = Math.round(70 + pulse * 30);
        var glowLight = Math.round(42 + pulse * 38);
        ctx.font = "10px Consolas,ui-monospace,monospace";
        ctx.shadowColor = "hsl(118, " + glowSat + "%, " + glowLight + "%)";
        ctx.globalAlpha = autoAlpha * (0.4 + pulse * 0.35);
        ctx.shadowBlur = 16 + pulse * 24;
        ctx.fillStyle = "hsl(118, " + Math.round(letterSat * 0.85) + "%, " + Math.round(letterLight * 0.75) + "%)";
        ctx.fillText("AUTO", ox + 4, oy + 12);
        ctx.globalAlpha = autoAlpha;
        ctx.shadowBlur = 8 + pulse * 16;
        ctx.fillStyle = "hsl(118, " + letterSat + "%, " + letterLight + "%)";
        ctx.fillText("AUTO", ox + 4, oy + 12);
        ctx.shadowBlur = 0;
        ctx.globalAlpha = 1;
      }
    }

    function renderLoop() {
      if (!snake) return;
      draw();
      snakeRafId = requestAnimationFrame(renderLoop);
    }

    function tick() {
      if (mode === "manual" && lastUserInputAt && Date.now() - lastUserInputAt >= AUTO_IDLE_MS) {
        mode = "auto";
      }
      if (mode === "auto") {
        nextDir = pickAutoDir();
      }
      dir = nextDir;
      var head = { x: body[0].x + dir.x, y: body[0].y + dir.y };
      var eating = head.x === food.x && head.y === food.y;
      if (willCollide(head.x, head.y, eating)) {
        recordSnakeHighOnFail(mode, score);
        resetSnake();
        return;
      }
      body.unshift(head);
      if (eating) {
        score++;
        updateSnakeScoreDisplay(scoreEl, score);
        food = randFood();
      } else {
        body.pop();
      }
    }

    function onKey(e) {
      var k = e.key;
      if (k !== "ArrowUp" && k !== "ArrowDown" && k !== "ArrowLeft" && k !== "ArrowRight") return;
      e.preventDefault();
      mode = "manual";
      lastUserInputAt = Date.now();
      if (k === "ArrowUp" && dir.y === 0) nextDir = { x: 0, y: -1 };
      else if (k === "ArrowDown" && dir.y === 0) nextDir = { x: 0, y: 1 };
      else if (k === "ArrowLeft" && dir.x === 0) nextDir = { x: -1, y: 0 };
      else if (k === "ArrowRight" && dir.x === 0) nextDir = { x: 1, y: 0 };
    }

    snakeWindowKeyHandler = onKey;
    window.addEventListener("keydown", snakeWindowKeyHandler, true);
    snake = { canvas: canvas, onKey: onKey };
    updateSnakeScoreDisplay(scoreEl, score);
    snakeRafId = requestAnimationFrame(renderLoop);
    gameLoopId = setInterval(tick, TICK_MS);
  }

  function show() {
    injectStyles();
    if (overlayEl && overlayEl.parentNode) {
      overlayEl.classList.remove("spot-mwl-done", "spot-mwl-failed");
      return;
    }
    overlayEl = document.createElement("div");
    overlayEl.id = OVERLAY_ID;
    overlayEl.innerHTML =
      '<div class="spot-mwl-card">' +
      '<h1 class="spot-mwl-title">Permissions matrix export</h1>' +
      '<p class="spot-mwl-warn">Processing — do not close this tab</p>' +
      '<p class="spot-mwl-hint">Your page was opened in a new tab so you can keep working. This tab runs the export and will close automatically when finished.</p>' +
      '<div class="spot-mwl-bar-wrap"><div class="spot-mwl-bar"></div></div>' +
      '<div class="spot-mwl-head"><span class="spot-mwl-msg">Starting…</span><span class="spot-mwl-pct">0%</span></div>' +
      '<pre class="spot-mwl-log"></pre>' +
      '<div class="spot-mwl-game">' +
      '<div class="spot-mwl-game-head">' +
      '<p class="spot-mwl-game-title">While you wait — <span class="spot-mwl-game-name">Auto-Snake</span></p>' +
      '<p class="spot-mwl-game-sub">(arrow keys to take over)</p></div>' +
      '<div class="spot-mwl-crt"><div class="spot-mwl-crt-bezel"><div class="spot-mwl-crt-screen">' +
      '<canvas width="256" height="256" aria-label="Snake mini-game"></canvas>' +
      '<div class="spot-mwl-crt-scanlines" aria-hidden="true"></div>' +
      '<div class="spot-mwl-crt-vignette" aria-hidden="true"></div>' +
      '<div class="spot-mwl-crt-glare" aria-hidden="true"></div>' +
      '<div class="spot-mwl-crt-noise" aria-hidden="true"></div></div></div></div>' +
      '<p class="spot-mwl-score">Score: 0 | Auto High: 0 | Your High: 0</p></div>' +
      '<div class="spot-mwl-actions"><button type="button" class="spot-mwl-cancel">Cancel export</button></div>' +
      '<p class="spot-mwl-close-note"></p></div>';
    document.documentElement.appendChild(overlayEl);
    barEl = overlayEl.querySelector(".spot-mwl-bar");
    pctEl = overlayEl.querySelector(".spot-mwl-pct");
    msgEl = overlayEl.querySelector(".spot-mwl-msg");
    logEl = overlayEl.querySelector(".spot-mwl-log");
    statusEl = overlayEl.querySelector(".spot-mwl-close-note");
    overlayEl.querySelector(".spot-mwl-cancel").addEventListener("click", function () {
      try { chrome.runtime.sendMessage({ type: "SPCSVExportCancel" }); } catch (_) {}
    });
    const scoreElInit = overlayEl.querySelector(".spot-mwl-score");
    if (scoreElInit) scoreElInit.textContent = formatScoreLine(0);
    startSnake(overlayEl.querySelector(".spot-mwl-crt-screen canvas"));
  }

  function update(opts) {
    if (!overlayEl) return;
    opts = opts || {};
    if (opts.percent != null && barEl && pctEl) {
      const p = Math.max(0, Math.min(100, Math.round(opts.percent)));
      barEl.style.width = p + "%";
      pctEl.textContent = p + "%";
    }
    if (opts.message && msgEl) msgEl.textContent = String(opts.message);
    if (opts.logLine && logEl) {
      const line = String(opts.logLine).trim();
      if (line) {
        const ts = new Date().toLocaleTimeString([], { hour: "2-digit", minute: "2-digit", second: "2-digit" });
        const prev = logEl.textContent || "";
        logEl.textContent = (prev ? prev + "\n" : "") + "[" + ts + "] " + line;
        logEl.scrollTop = logEl.scrollHeight;
      }
    }
  }

  function finish(success, message, autoCloseMs) {
    if (!overlayEl) return;
    stopSnake();
    overlayEl.classList.toggle("spot-mwl-done", !!success);
    overlayEl.classList.toggle("spot-mwl-failed", !success);
    const warn = overlayEl.querySelector(".spot-mwl-warn");
    if (warn) warn.textContent = success ? "Export complete" : "Export failed";
    if (msgEl && message) msgEl.textContent = String(message);
    if (barEl && success) barEl.style.width = "100%";
    if (pctEl && success) pctEl.textContent = "100%";
    if (statusEl) {
      const sec = Math.max(1, Math.round((autoCloseMs || 5000) / 1000));
      statusEl.textContent = "Closing this tab in " + sec + " second" + (sec === 1 ? "" : "s") + "…";
    }
    const cancelBtn = overlayEl.querySelector(".spot-mwl-cancel");
    if (cancelBtn) cancelBtn.style.display = "none";
  }

  function hide() {
    stopSnake();
    if (overlayEl && overlayEl.parentNode) overlayEl.parentNode.removeChild(overlayEl);
    overlayEl = barEl = pctEl = msgEl = logEl = statusEl = null;
  }

  function isActive() {
    return !!(overlayEl && overlayEl.parentNode);
  }

  window.SPOToolkitMatrixWorkerLock = {
    show: show,
    update: update,
    finish: finish,
    hide: hide,
    isActive: isActive
  };
})();
