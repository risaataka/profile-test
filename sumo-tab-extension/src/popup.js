(function () {
  'use strict';

  const R = self.SUMO_RANKS;
  const STATE_KEY = 'sumoState';
  const SETTINGS_KEY = 'sumoSettings';

  const $ = (id) => document.getElementById(id);

  function renderPreview(stage) {
    const rank = R.RANKS[stage - 1];
    $('preview').innerHTML =
      '<style>' + self.SUMO_STYLES.character + '</style>' +
      self.SUMO_RIKISHI.svg({ stage: stage, rank: rank });
  }

  function renderBanzuke(stage) {
    const html = R.RANKS.map(function (r) {
      const cls = r.stage === stage ? 'row on' : (r.stage < stage ? 'row done' : 'row');
      return '<div class="' + cls + '" data-stage="' + r.stage + '">' +
        '<span class="n">' + r.stage + '</span>' +
        '<i style="background:' + r.color + '"></i>' +
        '<span>' + r.name + '</span>' +
        '<span class="t">' + r.minTabs + '枚〜</span>' +
        '</div>';
    }).join('');
    const box = $('banzuke');
    box.innerHTML = html;
    const cur = box.querySelector('.row.on');
    if (cur) cur.scrollIntoView({ block: 'center' });
  }

  function render(tabCount) {
    const stage = R.stageForTabs(tabCount);
    const rank = R.RANKS[stage - 1];

    $('rankName').textContent = rank.name;
    $('kani').innerHTML = '<i style="background:' + rank.color + '"></i>' +
      rank.kani + '（' + rank.hue + 'のまわし）・第' + stage + '階級';
    $('count').textContent = 'いま ' + tabCount + ' タブ';

    const next = R.tabsToNextStage(tabCount);
    const back = R.tabsToDemotion(tabCount);

    const lo = rank.minTabs;
    const hi = stage < R.MAX_STAGE ? R.RANKS[stage].minTabs : lo + 1;
    const p = stage >= R.MAX_STAGE ? 1 : (tabCount - lo) / Math.max(1, hi - lo);
    $('bar').style.width = Math.round(Math.min(1, p) * 100) + '%';
    $('bar').style.background = rank.color;

    let hint;
    if (stage >= R.MAX_STAGE) {
      hint = '横綱。塩をまき散らしています。<b>' + (back || 0) + '枚</b>閉じれば土俵を降ります。';
    } else if (back) {
      hint = 'あと<b>' + next + '枚</b>開くと昇進 ／ <b>' + back + '枚</b>閉じれば降格。';
    } else {
      hint = 'あと<b>' + next + '枚</b>開くと昇進。ここが最軽量です。';
    }
    $('hint').innerHTML = hint;

    renderPreview(stage);
    renderBanzuke(stage);
  }

  function loadSettings() {
    chrome.storage.local.get(SETTINGS_KEY, function (got) {
      const s = Object.assign(
        { enabled: true, position: 'top-left', salt: true },
        got[SETTINGS_KEY] || {}
      );
      $('enabled').checked = !!s.enabled;
      $('salt').checked = s.salt !== false;
      $('position').value = s.position;
    });
  }

  function saveSettings() {
    chrome.storage.local.set({
      [SETTINGS_KEY]: {
        enabled: $('enabled').checked,
        salt: $('salt').checked,
        position: $('position').value
      }
    });
  }

  ['enabled', 'salt', 'position'].forEach(function (id) {
    $(id).addEventListener('change', saveSettings);
  });

  chrome.tabs.query({ windowType: 'normal' }, function (tabs) {
    render(tabs.length);
  });
  chrome.storage.local.get(STATE_KEY, function (got) {
    if (got[STATE_KEY]) render(got[STATE_KEY].tabCount);
  });
  chrome.storage.onChanged.addListener(function (c, area) {
    if (area === 'local' && c[STATE_KEY] && c[STATE_KEY].newValue) {
      render(c[STATE_KEY].newValue.tabCount);
    }
  });

  loadSettings();
})();
