/**
 * タブの枚数を数えて階級を storage に書くだけの軽量サービスワーカー。
 *
 * content script には messaging をせず storage.onChanged で伝える。
 * （タブ数ぶんの sendMessage が要らないので、タブが多い人ほど効く）
 */
importScripts('ranks.js');

const STATE_KEY = 'sumoState';
const SETTINGS_KEY = 'sumoSettings';
const SALT_KEY = 'sumoSalt';

const DEFAULT_SETTINGS = {
  enabled: true,
  position: 'top-left', // top-left | top-right | bottom-left | bottom-right
  salt: true
};

let recountTimer = null;

/** 連続したタブ操作をまとめて 1 回だけ数える */
function scheduleRecount() {
  if (recountTimer) return;
  recountTimer = setTimeout(() => {
    recountTimer = null;
    recount();
  }, 200);
}

async function recount() {
  let tabs;
  try {
    tabs = await chrome.tabs.query({ windowType: 'normal' });
  } catch (e) {
    return;
  }
  const tabCount = tabs.length;
  const stage = SUMO_RANKS.stageForTabs(tabCount);

  const stored = await chrome.storage.local.get(STATE_KEY);
  const prev = stored[STATE_KEY];
  if (prev && prev.tabCount === tabCount && prev.stage === stage) return;

  const state = {
    tabCount: tabCount,
    stage: stage,
    prevStage: prev ? prev.stage : stage,
    // 同じ階級のままでも content 側が変化に気づけるように連番を振る
    seq: (prev && prev.seq ? prev.seq : 0) + 1
  };

  await chrome.storage.local.set({ [STATE_KEY]: state });

  // 横綱から落ちたら、まいた塩は掃除される
  if (prev && prev.stage === SUMO_RANKS.YOKOZUNA && stage < SUMO_RANKS.YOKOZUNA) {
    await chrome.storage.local.remove(SALT_KEY);
  }

  updateBadge(state);
}

function updateBadge(state) {
  const rank = SUMO_RANKS.RANKS[state.stage - 1];
  chrome.action.setBadgeText({ text: String(state.tabCount) });
  chrome.action.setBadgeBackgroundColor({ color: rank.color });
  chrome.action.setTitle({
    title: `ちゃんこタブ力士 — ${rank.name}（${rank.kani}・${state.tabCount}タブ）`
  });
}

async function ensureSettings() {
  const got = await chrome.storage.local.get(SETTINGS_KEY);
  if (!got[SETTINGS_KEY]) {
    await chrome.storage.local.set({ [SETTINGS_KEY]: DEFAULT_SETTINGS });
  }
}

chrome.runtime.onInstalled.addListener(async () => {
  await ensureSettings();
  recount();
});
chrome.runtime.onStartup.addListener(async () => {
  await ensureSettings();
  recount();
});

chrome.tabs.onCreated.addListener(scheduleRecount);
chrome.tabs.onRemoved.addListener(scheduleRecount);
chrome.tabs.onAttached.addListener(scheduleRecount);
chrome.tabs.onDetached.addListener(scheduleRecount);
chrome.tabs.onReplaced.addListener(scheduleRecount);
chrome.windows.onCreated.addListener(scheduleRecount);
chrome.windows.onRemoved.addListener(scheduleRecount);

// サービスワーカーが起き直したときの取りこぼし対策
scheduleRecount();
