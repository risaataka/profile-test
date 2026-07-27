/**
 * ページ上に力士を出す。
 *
 * 方針（動作を軽く保つため）:
 *   - 通信は storage.onChanged のみ。常駐タイマーは横綱のときだけ。
 *   - 常時アニメは CSS（合成のみ）。JS の rAF は塩が積もる瞬間だけ動く。
 *   - タブが裏に回ったら animation-play-state: paused で完全停止。
 *   - shadow DOM + pointer-events:none なので、ページの見た目にも操作にも触らない。
 */
(function () {
  'use strict';

  if (window.top !== window) return;              // iframe には出さない
  if (document.documentElement.dataset.sumoTab) return;
  document.documentElement.dataset.sumoTab = '1';

  const STATE_KEY = 'sumoState';
  const SETTINGS_KEY = 'sumoSettings';
  const SALT_KEY = 'sumoSalt';

  const R = self.SUMO_RANKS;
  const RIKISHI = self.SUMO_RIKISHI;

  const COLS = 48;                                 // 塩の山の解像度
  const MAX_PARTICLES = 120;
  const THROW_INTERVAL = 3000;                     // 塩をまく間隔(ms)
  const GRAINS_PER_THROW = 18;

  let settings = { enabled: true, position: 'top-left', salt: true };
  let stage = 0;
  let tabCount = 0;
  let booted = false;

  let stageEl, rikishiEl, labelEl, pileEl, pilePathEl, pileEdgeEl;
  let cols = new Float32Array(COLS);
  let pileDirty = false;
  let pileRaf = 0;
  let pileVisible = false;
  let throwTimer = 0;
  let saveTimer = 0;
  let sinkTimer = 0;
  let liveParticles = 0;
  let labelTimer = 0;
  let flashTimer = 0;

  /* ------------------------------------------------------------------ 起動 */

  function safeStorageGet(keys) {
    return new Promise((resolve) => {
      try {
        chrome.storage.local.get(keys, (v) => {
          resolve(chrome.runtime.lastError ? {} : v || {});
        });
      } catch (e) {
        resolve({});
      }
    });
  }

  function safeStorageSet(obj) {
    try {
      chrome.storage.local.set(obj, () => void chrome.runtime.lastError);
    } catch (e) { /* 拡張がリロードされた等 */ }
  }

  async function boot() {
    const got = await safeStorageGet([STATE_KEY, SETTINGS_KEY, SALT_KEY]);
    if (got[SETTINGS_KEY]) Object.assign(settings, got[SETTINGS_KEY]);
    if (!settings.enabled) return;

    build();
    if (got[SALT_KEY] && Array.isArray(got[SALT_KEY].cols)) {
      const saved = got[SALT_KEY].cols;
      for (let i = 0; i < COLS; i++) cols[i] = saved[i] || 0;
      pileVisible = true;
      pileEl.classList.add('on');
      schedulePileRender();
    }

    const s = got[STATE_KEY];
    applyState(s ? s.tabCount : 1, s ? s.stage : 1, false);
    booted = true;
  }

  function build() {
    const host = document.createElement('div');
    host.id = 'chanko-tab-rikishi';
    host.setAttribute('aria-hidden', 'true');
    host.style.cssText =
      'all:initial;position:fixed;left:0;top:0;width:0;height:0;' +
      'z-index:2147483647;pointer-events:none;';

    const shadow = host.attachShadow({ mode: 'open' });
    const style = document.createElement('style');
    style.textContent = self.SUMO_STYLES.all;

    stageEl = document.createElement('div');
    stageEl.className = 'stage pos-' + settings.position;

    rikishiEl = document.createElement('div');
    rikishiEl.className = 'rikishi';

    labelEl = document.createElement('div');
    labelEl.className = 'label';

    pileEl = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
    pileEl.setAttribute('class', 'pile');
    pileEl.setAttribute('viewBox', '0 0 100 100');
    pileEl.setAttribute('preserveAspectRatio', 'none');
    pilePathEl = document.createElementNS('http://www.w3.org/2000/svg', 'path');
    pileEdgeEl = document.createElementNS('http://www.w3.org/2000/svg', 'path');
    pileEdgeEl.setAttribute('class', 'edge');
    pileEdgeEl.setAttribute('vector-effect', 'non-scaling-stroke');
    pileEl.appendChild(pilePathEl);
    pileEl.appendChild(pileEdgeEl);

    stageEl.appendChild(pileEl);
    stageEl.appendChild(rikishiEl);
    stageEl.appendChild(labelEl);
    shadow.appendChild(style);
    shadow.appendChild(stageEl);

    (document.body || document.documentElement).appendChild(host);

    document.addEventListener('visibilitychange', onVisibility, { passive: true });
    onVisibility();
  }

  /* ------------------------------------------------------------ 階級の反映 */

  /** 階級ごとの表示サイズ(px)。44px から 240px まで等比で大きくなる */
  function sizeFor(st) {
    const px = 44 * Math.pow(240 / 44, (st - 1) / 23);
    return Math.round(Math.min(px, window.innerHeight * 0.42));
  }

  function applyState(newTabCount, newStage, animated) {
    const prevStage = stage;
    tabCount = newTabCount;
    stage = newStage;

    if (prevStage === stage) {
      if (booted) showLabel(false);
      return;
    }

    const promoted = stage > prevStage;
    const render = () => {
      const rank = R.RANKS[stage - 1];
      rikishiEl.innerHTML = RIKISHI.svg({ stage: stage, rank: rank });
      stageEl.style.setProperty('--size', sizeFor(stage) + 'px');
      stageEl.style.setProperty('--rank-color', rank.color);
      updateSalt();
      showLabel(prevStage > 0);
    };

    if (!animated || prevStage === 0) {
      render();
      return;
    }

    if (promoted) {
      // 白く点滅 → それから大きくなる
      rikishiEl.classList.remove('flash');
      void rikishiEl.offsetWidth;
      rikishiEl.classList.add('flash');
      clearTimeout(flashTimer);
      flashTimer = setTimeout(() => {
        rikishiEl.classList.remove('flash');
        render();
      }, 510);
    } else {
      rikishiEl.classList.remove('demote');
      void rikishiEl.offsetWidth;
      rikishiEl.classList.add('demote');
      render();
    }
  }

  function showLabel(announce) {
    const rank = R.RANKS[stage - 1];
    const left = R.tabsToDemotion(tabCount);
    const tail = stage === R.MAX_STAGE
      ? 'タブ' + tabCount + '枚 — 塩をまいています'
      : (left ? 'タブ' + tabCount + '枚 — あと' + left + '枚閉じれば降格' : 'タブ' + tabCount + '枚');
    labelEl.innerHTML =
      '<span class="sw"></span>' + rank.name + ' ' +
      '<span style="opacity:.65">' + rank.kani + '</span> ・ ' + tail;

    labelEl.classList.add('show');
    clearTimeout(labelTimer);
    labelTimer = setTimeout(() => labelEl.classList.remove('show'), announce ? 3600 : 2400);
  }

  /* ----------------------------------------------------------------- 塩まき */

  function updateSalt() {
    const yokozuna = stage === R.MAX_STAGE && settings.salt !== false;
    if (yokozuna && !document.hidden) {
      if (!throwTimer) {
        throwTimer = setInterval(throwSalt, THROW_INTERVAL);
        setTimeout(throwSalt, 900);
      }
      clearInterval(sinkTimer); sinkTimer = 0;
    } else {
      clearInterval(throwTimer); throwTimer = 0;
      if (!yokozuna && pileVisible && !sinkTimer) startSinking();
    }
  }

  /** 横綱でなくなったら、積もった塩は片付けられる */
  function startSinking() {
    sinkTimer = setInterval(() => {
      let max = 0;
      for (let i = 0; i < COLS; i++) {
        cols[i] *= 0.9;
        if (cols[i] > max) max = cols[i];
      }
      schedulePileRender();
      if (max < 0.004) {
        clearInterval(sinkTimer); sinkTimer = 0;
        cols.fill(0);
        pileVisible = false;
        pileEl.classList.remove('on');
        schedulePileRender();
      }
    }, 90);
  }

  function throwSalt() {
    if (document.hidden || stage !== R.MAX_STAGE) return;
    rikishiEl.classList.remove('throwing');
    void rikishiEl.offsetWidth;
    rikishiEl.classList.add('throwing');
    setTimeout(() => rikishiEl.classList.remove('throwing'), 1200);
    setTimeout(spawnGrains, 280);
  }

  function spawnGrains() {
    if (document.hidden || stage !== R.MAX_STAGE) return;
    if (liveParticles > MAX_PARTICLES) return;

    const box = rikishiEl.getBoundingClientRect();
    const o = RIKISHI.saltOrigin(stage);
    const hx = box.left + box.width * o.x;
    const hy = box.top + box.height * o.y - box.height * 0.35; // 振りかぶった手の高さ
    const vw = window.innerWidth;
    const vh = window.innerHeight;
    const frag = document.createDocumentFragment();

    for (let i = 0; i < GRAINS_PER_THROW; i++) {
      const g = document.createElement('div');
      g.className = 'salt';
      // 着地点（画面下 = 手前）。まいた方向にすこし寄せる
      const bias = settings.position.indexOf('right') >= 0 ? 0.3 : 0.03;
      const xf = Math.min(0.99, Math.max(0.01, bias + Math.random() * 0.67));
      const x1 = xf * vw;
      const y1 = vh - pileTopPx(xf) - 2;
      const s = g.style;
      s.setProperty('--x0', hx.toFixed(1) + 'px');
      s.setProperty('--y0', hy.toFixed(1) + 'px');
      s.setProperty('--xm', (hx + (x1 - hx) * 0.45).toFixed(1) + 'px');
      s.setProperty('--ym', (hy + (y1 - hy) * 0.2 - 40 - Math.random() * 40).toFixed(1) + 'px');
      s.setProperty('--sm', (1.2 + Math.random() * 1.1).toFixed(2));
      s.setProperty('--x1', x1.toFixed(1) + 'px');
      s.setProperty('--y1', y1.toFixed(1) + 'px');
      s.setProperty('--s1', (1.6 + Math.random() * 1.6).toFixed(2));
      s.setProperty('--dur', (1.5 + Math.random() * 0.9).toFixed(2) + 's');
      g.dataset.xf = xf;
      g.addEventListener('animationend', onGrainLanded, { once: true });
      frag.appendChild(g);
      liveParticles++;
    }
    stageEl.appendChild(frag);
  }

  function onGrainLanded(e) {
    const g = e.currentTarget;
    liveParticles--;
    addGrain(parseFloat(g.dataset.xf));
    g.remove();
  }

  /* ------------------------------------------------------------ 積もった塩 */

  const PILE_MAX = 0.45;        // pile レイヤー(=40vh)に対する上限
  const GRAIN = 0.028;

  function addGrain(xf) {
    const i = Math.max(0, Math.min(COLS - 1, Math.round(xf * (COLS - 1))));
    cols[i] = Math.min(PILE_MAX, cols[i] + GRAIN);
    if (i > 0) cols[i - 1] = Math.min(PILE_MAX, cols[i - 1] + GRAIN * 0.4);
    if (i < COLS - 1) cols[i + 1] = Math.min(PILE_MAX, cols[i + 1] + GRAIN * 0.4);
    relax(i);
    if (!pileVisible) {
      pileVisible = true;
      pileEl.classList.add('on');
    }
    schedulePileRender();
    schedulePileSave();
  }

  /** 安息角っぽく、隣との段差をならす */
  function relax(center) {
    const SLOPE = GRAIN * 1.6;
    for (let pass = 0; pass < 2; pass++) {
      for (let i = Math.max(1, center - 4); i < Math.min(COLS - 1, center + 5); i++) {
        const d = cols[i] - cols[i - 1];
        if (d > SLOPE) { cols[i] -= d * 0.25; cols[i - 1] += d * 0.25; }
        const d2 = cols[i] - cols[i + 1];
        if (d2 > SLOPE) { cols[i] -= d2 * 0.25; cols[i + 1] += d2 * 0.25; }
      }
    }
  }

  function pileLayerPx() {
    return window.innerHeight * 0.4;   // CSS の .pile { height: 40vh } と対応
  }

  function pileTopPx(xf) {
    const i = Math.max(0, Math.min(COLS - 1, Math.round(xf * (COLS - 1))));
    return cols[i] * pileLayerPx();
  }

  function schedulePileRender() {
    pileDirty = true;
    if (pileRaf) return;
    pileRaf = requestAnimationFrame(() => {
      pileRaf = 0;
      if (!pileDirty) return;
      pileDirty = false;
      renderPile();
    });
  }

  function renderPile() {
    const step = 100 / (COLS - 1);
    let d = 'M 0 ' + (100 - cols[0] * 100).toFixed(2);
    for (let i = 1; i < COLS; i++) {
      const x0 = (i - 1) * step;
      const x1 = i * step;
      const y0 = 100 - cols[i - 1] * 100;
      const y1 = 100 - cols[i] * 100;
      const mx = (x0 + x1) / 2;
      d += ' C ' + mx.toFixed(2) + ' ' + y0.toFixed(2) +
           ', ' + mx.toFixed(2) + ' ' + y1.toFixed(2) +
           ', ' + x1.toFixed(2) + ' ' + y1.toFixed(2);
    }
    pileEdgeEl.setAttribute('d', d);            // 白い背景でも輪郭が見えるように
    pilePathEl.setAttribute('d', d + ' L 100 100 L 0 100 Z');
  }

  function schedulePileSave() {
    if (saveTimer) return;
    saveTimer = setTimeout(() => {
      saveTimer = 0;
      const out = new Array(COLS);
      for (let i = 0; i < COLS; i++) out[i] = Math.round(cols[i] * 1000) / 1000;
      safeStorageSet({ [SALT_KEY]: { cols: out } });
    }, 4000);
  }

  /* ------------------------------------------------------------ 省エネ制御 */

  function onVisibility() {
    if (!stageEl) return;
    if (document.hidden) {
      stageEl.classList.add('paused');
      clearInterval(throwTimer); throwTimer = 0;
    } else {
      stageEl.classList.remove('paused');
      updateSalt();
    }
  }

  /* -------------------------------------------------------------- 状態同期 */

  try {
    chrome.storage.onChanged.addListener((changes, area) => {
      if (area !== 'local') return;

      if (changes[SETTINGS_KEY]) {
        const next = changes[SETTINGS_KEY].newValue || {};
        const wasEnabled = settings.enabled;
        Object.assign(settings, next);
        if (!settings.enabled && wasEnabled) {
          teardown();
          return;
        }
        if (settings.enabled && !stageEl) {
          boot();
          return;
        }
        if (stageEl) {
          stageEl.className = 'stage pos-' + settings.position +
            (document.hidden ? ' paused' : '');
          updateSalt();
        }
      }

      if (changes[SALT_KEY] && !changes[SALT_KEY].newValue && pileVisible) {
        // 横綱から落ちた（background が塩を消した）
        if (!sinkTimer) startSinking();
      }

      if (changes[STATE_KEY] && stageEl) {
        const s = changes[STATE_KEY].newValue;
        if (s) applyState(s.tabCount, s.stage, true);
      }
    });
  } catch (e) { /* 拡張のコンテキストが無効 */ }

  function teardown() {
    clearInterval(throwTimer); throwTimer = 0;
    clearInterval(sinkTimer); sinkTimer = 0;
    clearTimeout(labelTimer);
    clearTimeout(flashTimer);
    const host = document.getElementById('chanko-tab-rikishi');
    if (host) host.remove();
    stageEl = null;
    stage = 0;
    booted = false;
  }

  boot();
})();
