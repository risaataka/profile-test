/**
 * shadow DOM に流し込む CSS。content script と popup で共有する。
 * アニメーションはすべて CSS（合成レイヤー）任せ。JS の rAF は塩が積もる時だけ。
 */
(function (root) {
  'use strict';

  // 力士そのもののアニメーション（popup でも使う）
  const CHARACTER = `
.rikishi-svg { display:block; width:100%; height:100%; overflow:visible; }
.rikishi-svg * { transform-box: view-box; }

[data-anim] .mouth { animation: sumo-mogu .62s ease-in-out infinite; }
@keyframes sumo-mogu {
  0%, 100% { transform: scaleY(1); }
  50%      { transform: scaleY(.3); }
}

[data-anim] .arm-r { animation: sumo-eat 2.6s ease-in-out infinite; }
@keyframes sumo-eat {
  0%, 28%  { transform: rotate(0deg); }
  45%      { transform: rotate(-12deg); }
  62%, 100%{ transform: rotate(0deg); }
}

[data-anim] .eye { animation: sumo-blink 5.4s ease-in-out infinite; }
@keyframes sumo-blink {
  0%, 94%, 100% { transform: scaleY(1); }
  97%           { transform: scaleY(.1); }
}

[data-anim] .steam path { animation: sumo-steam 3.2s ease-out infinite; }
[data-anim] .steam .s2 { animation-delay: 1.05s; }
[data-anim] .steam .s3 { animation-delay: 2.1s; }
@keyframes sumo-steam {
  0%   { opacity: 0;   transform: translateY(5px)  scaleY(.6); }
  25%  { opacity: .7; }
  100% { opacity: 0;   transform: translateY(-12px) scaleY(1.15); }
}

/* 横綱の塩まき */
.throwing .arm-l { animation: sumo-shio 1.15s cubic-bezier(.3,.1,.2,1); }
@keyframes sumo-shio {
  0%   { transform: rotate(0deg); }
  22%  { transform: rotate(112deg); }
  38%  { transform: rotate(96deg); }
  52%  { transform: rotate(140deg); }
  100% { transform: rotate(0deg); }
}
`;

  // ページに重ねる部分
  const OVERLAY = `
:host { all: initial; }
.stage {
  position: fixed; inset: 0;
  pointer-events: none;          /* ページの操作は絶対に邪魔しない */
  overflow: hidden;
  z-index: 2147483647;
  contain: layout style;
}

.rikishi {
  position: absolute;
  width: var(--size, 44px);
  height: var(--size, 44px);
  transition: width .55s cubic-bezier(.2,1.5,.35,1),
              height .55s cubic-bezier(.2,1.5,.35,1),
              opacity .3s linear;
  transform-origin: 50% 100%;
}
.pos-top-left     .rikishi { left: 14px;  top: 6px; }
.pos-top-right    .rikishi { right: 14px; top: 6px; }
.pos-bottom-left  .rikishi { left: 14px;  bottom: 4px; }
.pos-bottom-right .rikishi { right: 14px; bottom: 4px; }

/* 昇進：白く点滅してから大きくなる */
.rikishi.flash { animation: sumo-flash .17s steps(1) 3; }
@keyframes sumo-flash {
  0%, 49%   { filter: brightness(3.6) saturate(.12)
                      drop-shadow(0 0 10px rgba(255,214,106,.95))
                      drop-shadow(0 0 2px rgba(0,0,0,.45)); }
  50%, 100% { filter: none; }
}
.rikishi.demote { animation: sumo-demote .5s ease-out; }
@keyframes sumo-demote {
  0%   { transform: translateY(0)   rotate(0deg); }
  35%  { transform: translateY(4px) rotate(-3deg); }
  100% { transform: translateY(0)   rotate(0deg); }
}

.label {
  position: absolute;
  font: 600 12px/1.5 "Hiragino Kaku Gothic ProN", "Yu Gothic", "Noto Sans JP", sans-serif;
  color: #fff;
  background: rgba(28,24,22,.82);
  border-radius: 999px;
  padding: 3px 10px;
  white-space: nowrap;
  opacity: 0;
  transform: translateY(-4px);
  transition: opacity .25s, transform .25s;
  backdrop-filter: blur(2px);
}
.label.show { opacity: 1; transform: translateY(0); }
.label .sw {
  display: inline-block; width: 8px; height: 8px; border-radius: 2px;
  background: var(--rank-color, #fff);
  box-shadow: 0 0 0 1px rgba(255,255,255,.75);
  vertical-align: 0px; margin-right: 6px;
}
.pos-top-left     .label { left: 14px;  top: calc(6px + var(--size, 44px) - 4px); }
.pos-top-right    .label { right: 14px; top: calc(6px + var(--size, 44px) - 4px); text-align: right; }
.pos-bottom-left  .label { left: 14px;  bottom: calc(4px + var(--size, 44px) - 4px); }
.pos-bottom-right .label { right: 14px; bottom: calc(4px + var(--size, 44px) - 4px); }

/* まいた塩 */
.salt {
  position: absolute;
  left: 0; top: 0;
  width: 4px; height: 4px;
  margin: -2px 0 0 -2px;
  border-radius: 50%;
  background: #fff;
  box-shadow: 0 0 4px rgba(255,255,255,.85);
  will-change: transform, opacity;
  animation: sumo-salt var(--dur, 1.8s) cubic-bezier(.3,.5,.55,1) forwards;
}
@keyframes sumo-salt {
  0%   { transform: translate3d(var(--x0), var(--y0), 0) scale(.15); opacity: 0; }
  10%  { opacity: 1; }
  45%  { transform: translate3d(var(--xm), var(--ym), 0) scale(var(--sm)); opacity: 1; }
  100% { transform: translate3d(var(--x1), var(--y1), 0) scale(var(--s1)); opacity: .95; }
}

.pile {
  position: absolute;
  left: 0; bottom: 0;
  width: 100%;
  height: var(--pile-h, 40vh);
  opacity: 0;
  transition: opacity .6s linear;
}
.pile.on { opacity: 1; }
.pile path { fill: #fdfdfb; }
.pile path.edge { fill: none; stroke: rgba(120,132,150,.5); stroke-width: 1.4; }

/* タブが裏に回ったら全部止める（軽さのため） */
.paused * { animation-play-state: paused !important; }

@media (prefers-reduced-motion: reduce) {
  .rikishi-svg * { animation: none !important; }
  .salt { animation-duration: 2.6s; }
}
`;

  root.SUMO_STYLES = { character: CHARACTER, overlay: OVERLAY, all: CHARACTER + OVERLAY };
})(typeof self !== 'undefined' ? self : this);
