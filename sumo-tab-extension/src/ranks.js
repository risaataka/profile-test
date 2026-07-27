/**
 * 階級テーブル（24段階）
 *
 * まわしの色は冠位十二階（徳・仁・礼・信・義・智 / 各 大・小）に準拠。
 * 位の低い順 = 智 → 義 → 信 → 礼 → 仁 → 徳、各色に「小」「大」の2段階で 12 x 2 = 24。
 *
 * background.js からは importScripts()、content.js からは content_scripts の
 * 読み込み順で共有する。どちらも self に生やすだけのプレーンスクリプト。
 */
(function (root) {
  'use strict';

  // 冠位十二階の色（低い位 → 高い位）。
  // 1色 = 4階級（小X の薄い/濃い、大X の薄い/濃い）で、上に行くほど濃く鮮やかになる。
  const KANI = [
    { hue: '黒', kani: ['小智', '大智'], colors: ['#7a7a7a', '#4f4f4f', '#2a2a2a', '#000000'] },
    { hue: '白', kani: ['小義', '大義'], colors: ['#cfc7b2', '#e0dacb', '#f0ece1', '#ffffff'] },
    { hue: '黄', kani: ['小信', '大信'], colors: ['#e6d79a', '#ddc257', '#efb520', '#e09a00'] },
    { hue: '赤', kani: ['小礼', '大礼'], colors: ['#dda49a', '#d0736a', '#c33a42', '#9e0f26'] },
    { hue: '青', kani: ['小仁', '大仁'], colors: ['#93b2d6', '#5c8ec0', '#2c5f92', '#10396c'] },
    { hue: '紫', kani: ['小徳', '大徳'], colors: ['#a98fc6', '#8461b0', '#633a91', '#40166a'] }
  ];

  // 番付名（下から上へ 24 個）
  const BANZUKE = [
    '序ノ口', '序二段', '三段目',
    '幕下十五枚目', '幕下十枚目', '幕下五枚目', '幕下筆頭',
    '十両十三枚目', '十両九枚目', '十両五枚目', '十両二枚目', '十両筆頭',
    '前頭十六枚目', '前頭十三枚目', '前頭十枚目', '前頭七枚目',
    '前頭五枚目', '前頭三枚目', '前頭二枚目', '前頭筆頭',
    '小結', '関脇', '大関', '横綱'
  ];

  // 昇進に必要なタブ数（この枚数「以上」でその階級）
  const THRESHOLDS = [
    1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 12, 14,
    16, 18, 20, 23, 26, 30, 34, 38, 43, 48, 54, 60
  ];

  const RANKS = BANZUKE.map(function (name, i) {
    const k = KANI[i >> 2];              // 4 階級で 1 色
    const step = i & 3;                  // その色の中の 0..3
    return {
      stage: i + 1,
      name: name,
      kani: k.kani[step >> 1],           // 前半 2 つが 小X、後半 2 つが 大X
      hue: k.hue,
      color: k.colors[step],
      // まわしの陰影用（1 段濃い色。最上段は 1 段薄い色）
      shade: k.colors[step === 3 ? 2 : step + 1],
      minTabs: THRESHOLDS[i]
    };
  });

  const MAX_STAGE = RANKS.length; // 24
  const YOKOZUNA = MAX_STAGE;

  /** タブ枚数から階級(1..24)を求める */
  function stageForTabs(tabCount) {
    let stage = 1;
    for (let i = 0; i < THRESHOLDS.length; i++) {
      if (tabCount >= THRESHOLDS[i]) stage = i + 1;
      else break;
    }
    return stage;
  }

  /** 次の階級に上がるまでのタブ数（横綱なら null） */
  function tabsToNextStage(tabCount) {
    const stage = stageForTabs(tabCount);
    if (stage >= MAX_STAGE) return null;
    return THRESHOLDS[stage] - tabCount;
  }

  /** 降格まであと何枚閉じればよいか（序ノ口なら null） */
  function tabsToDemotion(tabCount) {
    const stage = stageForTabs(tabCount);
    if (stage <= 1) return null;
    return tabCount - THRESHOLDS[stage - 1] + 1;
  }

  root.SUMO_RANKS = {
    RANKS: RANKS,
    MAX_STAGE: MAX_STAGE,
    YOKOZUNA: YOKOZUNA,
    stageForTabs: stageForTabs,
    tabsToNextStage: tabsToNextStage,
    tabsToDemotion: tabsToDemotion
  };
})(typeof self !== 'undefined' ? self : this);
