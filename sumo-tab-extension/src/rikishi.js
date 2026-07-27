/**
 * 力士のSVGを組み立てるだけのモジュール。
 * content script / popup の両方から使うのでプレーンスクリプト（self に生やす）。
 *
 * 階級(1..24)から「太さ」を決めて、同じテンプレートの座標を太らせる。
 * 序ノ口はやせた男性、横綱はまん丸。まわしの色は冠位十二階。
 */
(function (root) {
  'use strict';

  const SKIN = '#f1c49c';
  const SKIN_LINE = '#d99f70';
  const SKIN_SHADE = '#e0a87c';
  const INK = '#241f1b';
  const NABE = '#463a34';
  const NABE_RIM = '#6d5a53';
  const SOUP = '#e2a45c';

  let seq = 0;
  const n = (v) => Math.round(v * 100) / 100;

  /** 腰に巻くもの（まわし・綱）。胴でクリップして使う横帯 */
  function band(g, y0, y1) {
    const x0 = n(g.cx - g.bb - 8);
    const x1 = n(g.cx + g.bb + 8);
    const dip = 6;
    return `M ${x0} ${n(y0)} Q ${g.cx} ${n(y0 + dip)} ${x1} ${n(y0)}` +
           ` L ${x1} ${n(y1)} Q ${g.cx} ${n(y1 + dip)} ${x0} ${n(y1)} Z`;
  }

  /** 縁取り付きの手足。太い暗色の上に肌色を重ねるだけ */
  function limb(d, w) {
    return `<path d="${d}" fill="none" stroke="${SKIN_LINE}" stroke-width="${n(w + 1.8)}" stroke-linecap="round"/>` +
           `<path d="${d}" fill="none" stroke="${SKIN}" stroke-width="${w}" stroke-linecap="round"/>`;
  }

  /**
   * @param {object} o
   * @param {number} o.stage 1..24
   * @param {object} o.rank  ranks.js の RANKS[stage-1]
   * @param {boolean} [o.animate=true] アニメーション用の属性を付けるか
   */
  function svg(o) {
    const stage = Math.max(1, Math.min(24, o.stage | 0));
    const rank = o.rank;
    const g = geometry(stage);
    const uid = 'ctr' + (++seq);
    const isYokozuna = stage === 24;
    const fat = g.fat;

    const parts = [];

    // ---- 足（あぐら） ----
    parts.push(`<g stroke="${SKIN_LINE}" stroke-width="1.1">
      <ellipse cx="${n(g.cx - g.bb * 1.02)}" cy="${n(g.legCy - 2)}" rx="${g.kneeRx}" ry="${g.kneeRy}" fill="${SKIN}"/>
      <ellipse cx="${n(g.cx + g.bb * 1.02)}" cy="${n(g.legCy - 2)}" rx="${g.kneeRx}" ry="${g.kneeRy}" fill="${SKIN}"/>
      <path d="M ${n(g.cx - g.legRx)} ${g.legCy}
               Q ${n(g.cx - g.legRx)} ${n(g.legCy + g.legRy)} ${n(g.cx - g.legRx * 0.45)} ${n(g.legCy + g.legRy)}
               L ${n(g.cx + g.legRx * 0.45)} ${n(g.legCy + g.legRy)}
               Q ${n(g.cx + g.legRx)} ${n(g.legCy + g.legRy)} ${n(g.cx + g.legRx)} ${g.legCy}
               Q ${g.cx} ${n(g.legCy - g.legRy * 0.9)} ${n(g.cx - g.legRx)} ${g.legCy} Z" fill="${SKIN}"/>
    </g>`);
    // 組んだ足の重なり
    parts.push(`<path d="M ${n(g.cx - g.legRx * 0.55)} ${n(g.legCy + 2)} Q ${g.cx} ${n(g.legCy + 6)} ${n(g.cx + g.legRx * 0.55)} ${n(g.legCy + 2)}"
      fill="none" stroke="${SKIN_LINE}" stroke-width="1.2" stroke-linecap="round"/>`);

    // ---- 胴 ----
    parts.push(`<defs><clipPath id="${uid}"><path d="${g.torso}"/></clipPath></defs>`);
    parts.push(`<path d="${g.torso}" fill="${SKIN}" stroke="${SKIN_LINE}" stroke-width="1.1"/>`);
    // 腹（太るほどはっきり出る）
    if (fat > 0.04) {
      parts.push(`<g clip-path="url(#${uid})">
        <ellipse cx="${g.cx}" cy="${n(g.hy - 15)}" rx="${n(g.bb * 0.8)}" ry="${n(9 + 8 * fat)}"
          fill="${SKIN_SHADE}" opacity="${n(0.16 + 0.26 * fat)}"/>
      </g>`);
      parts.push(`<path d="M ${g.cx} ${n(g.hy - 22)} q 2.5 4 0 8" fill="none"
        stroke="${SKIN_LINE}" stroke-width="1.2" stroke-linecap="round" opacity="0.75"/>`);
    }

    // ---- 下がり（まわしの前垂れ） ----
    const sagariN = 5;
    const sagariW = g.bb * 0.95;
    for (let i = 0; i < sagariN; i++) {
      const x = n(g.cx - sagariW / 2 + (i * sagariW) / (sagariN - 1));
      const dy = n(Math.abs(i - (sagariN - 1) / 2) * 1.6);   // 中央ほど長く垂れる
      parts.push(`<rect x="${n(x - 0.9)}" y="${n(g.hy + 2 + dy)}" width="1.8" height="${n(11 + 4 * fat)}" rx="0.9"
        fill="${rank.color}" stroke="${rank.shade}" stroke-width="0.35"/>`);
    }

    // ---- まわし（冠位十二階の色）：胴のシルエットに沿わせる ----
    parts.push(`<g clip-path="url(#${uid})">
      <path d="${band(g, g.hy - 13, g.hy + 4)}" fill="${rank.color}"/>
      <path d="${band(g, g.hy - 13, g.hy - 9)}" fill="${rank.shade}" opacity="0.45"/>
      <path d="${band(g, g.hy - 3, g.hy - 1.5)}" fill="#000" opacity="0.12"/>
    </g>`);

    // ---- 横綱の綱 ----
    if (isYokozuna) {
      parts.push(`<g clip-path="url(#${uid})">
        <path d="${band(g, g.hy - 22, g.hy - 14)}" fill="#fdfbf4"/>
        <path d="${band(g, g.hy - 22, g.hy - 14)}" fill="none"
          stroke="#ddd6c3" stroke-width="2.4" stroke-dasharray="2.5 4.5"/>
      </g>`);
      // 綱の結び目と紙垂
      parts.push(`<g fill="#ffffff" stroke="#dcd6c4" stroke-width="0.5">
        <circle cx="${g.cx}" cy="${n(g.hy - 16)}" r="3.4"/>
        <path d="M ${n(g.cx - 5.5)} ${n(g.hy - 17)} q -3 -3 -6 -1 q 3 3 6 1 z"/>
        <path d="M ${n(g.cx + 5.5)} ${n(g.hy - 17)} q 3 -3 6 -1 q -3 3 -6 1 z"/>
        <path d="M ${n(g.cx - 1.6)} ${n(g.hy - 13)} l 3.2 0 l -0.6 3 l -2 0 z"/>
      </g>`);
    }

    // ---- 左腕（塩をまく方） ----
    parts.push(`<g class="arm-l" style="transform-origin:${g.shL.x}px ${g.shL.y}px">
      ${limb(`M ${g.shL.x} ${g.shL.y} Q ${n(g.cx - g.bb * 1.1)} ${n(g.sy + 24)} ${g.handL.x} ${g.handL.y}`, g.armW)}
      <circle cx="${g.handL.x}" cy="${g.handL.y}" r="${n(g.armW * 0.62)}" fill="${SKIN}" stroke="${SKIN_LINE}" stroke-width="0.9"/>
      ${isYokozuna ? `<g class="salt-in-hand" fill="#fff" stroke="#e6e2d6" stroke-width="0.4">
        <circle cx="${n(g.handL.x - 1.5)}" cy="${n(g.handL.y - 4)}" r="1.4"/>
        <circle cx="${n(g.handL.x + 2)}" cy="${n(g.handL.y - 5)}" r="1.1"/>
        <circle cx="${n(g.handL.x + 0.2)}" cy="${n(g.handL.y - 6.8)}" r="1"/>
      </g>` : ''}
    </g>`);

    // ---- 頭 ----
    parts.push(`<circle cx="${g.cx}" cy="${g.headCy}" r="${g.headR}" fill="${SKIN}" stroke="${SKIN_LINE}" stroke-width="1.1"/>`);
    parts.push(`<circle cx="${n(g.cx - g.headR)}" cy="${n(g.headCy + 3)}" r="${n(2 + 1.2 * fat)}" fill="${SKIN}" stroke="${SKIN_LINE}" stroke-width="0.9"/>`);
    parts.push(`<circle cx="${n(g.cx + g.headR)}" cy="${n(g.headCy + 3)}" r="${n(2 + 1.2 * fat)}" fill="${SKIN}" stroke="${SKIN_LINE}" stroke-width="0.9"/>`);
    // 髪とまげ
    const hairY = n(g.headCy - g.headR * 0.34);
    const hairHalf = n(Math.sqrt(Math.max(0, g.headR * g.headR - Math.pow(g.headR * 0.34, 2))));
    parts.push(`<path d="M ${n(g.cx - hairHalf)} ${hairY} A ${g.headR} ${g.headR} 0 0 1 ${n(g.cx + hairHalf)} ${hairY} Z" fill="${INK}"/>`);
    parts.push(`<g fill="${INK}">
      <rect x="${n(g.cx - 3)}" y="${n(g.headCy - g.headR - 4.5)}" width="6" height="7" rx="3"/>
      <path d="M ${n(g.cx + 1)} ${n(g.headCy - g.headR - 2.5)} q 8.5 -1 9.5 2.5 q -3.5 2.5 -9.5 1.5 z"/>
    </g>`);
    // 顔（シンプル）
    parts.push(`<g class="eyes" fill="${INK}">
      <ellipse class="eye" cx="${n(g.cx - g.eyeX)}" cy="${g.eyeY}" rx="1.8" ry="2.2"
        style="transform-origin:${n(g.cx - g.eyeX)}px ${g.eyeY}px"/>
      <ellipse class="eye" cx="${n(g.cx + g.eyeX)}" cy="${g.eyeY}" rx="1.8" ry="2.2"
        style="transform-origin:${n(g.cx + g.eyeX)}px ${g.eyeY}px"/>
    </g>`);
    parts.push(`<g fill="#e78e7c" opacity="0.4">
      <ellipse cx="${n(g.cx - g.eyeX - 2)}" cy="${n(g.eyeY + 5.5)}" rx="${n(2.6 + 1.2 * fat)}" ry="1.7"/>
      <ellipse cx="${n(g.cx + g.eyeX + 2)}" cy="${n(g.eyeY + 5.5)}" rx="${n(2.6 + 1.2 * fat)}" ry="1.7"/>
    </g>`);
    parts.push(`<ellipse class="mouth" cx="${g.cx}" cy="${g.mouthY}" rx="${g.mouthRx}" ry="2.9"
      fill="#8c4136" style="transform-origin:${g.cx}px ${g.mouthY}px"/>`);

    // ---- 右腕（箸で口へ運ぶ） ----
    parts.push(`<g class="arm-r" style="transform-origin:${g.shR.x}px ${g.shR.y}px">
      ${limb(`M ${g.shR.x} ${g.shR.y} Q ${n(g.cx + g.bb * 1.05)} ${n(g.sy + 18)} ${g.handR.x} ${g.handR.y}`, g.armW)}
      <g stroke="#caa066" stroke-width="1.2" stroke-linecap="round">
        <line x1="${n(g.handR.x - 1)}" y1="${n(g.handR.y - 1)}" x2="${n(g.cx + g.mouthRx + 1)}" y2="${n(g.mouthY - 1)}"/>
        <line x1="${n(g.handR.x - 1)}" y1="${n(g.handR.y + 1.5)}" x2="${n(g.cx + g.mouthRx + 1)}" y2="${n(g.mouthY + 1.5)}"/>
      </g>
      <circle cx="${g.handR.x}" cy="${g.handR.y}" r="${n(g.armW * 0.62)}" fill="${SKIN}" stroke="${SKIN_LINE}" stroke-width="0.9"/>
    </g>`);

    // ---- ちゃんこ鍋 ----
    parts.push(`<g class="steam" fill="none" stroke="#ffffff" stroke-width="1.7" stroke-linecap="round" opacity="0.55">
      <path class="s1" d="M ${n(g.cx - g.pr * 0.55)} ${n(g.py - 4)} q 3 -4 0 -7.5 q -3 -3.5 0 -6"/>
      <path class="s2" d="M ${g.cx} ${n(g.py - 6)} q 3.2 -4 0 -8 q -3.2 -3.5 0 -6"/>
      <path class="s3" d="M ${n(g.cx + g.pr * 0.55)} ${n(g.py - 4)} q 3 -4 0 -7.5 q -3 -3.5 0 -6"/>
    </g>`);
    parts.push(`<g class="nabe">
      <path d="M ${n(g.cx - g.pr)} ${g.py} A ${g.pr} ${g.ph} 0 0 0 ${n(g.cx + g.pr)} ${g.py} Z" fill="${NABE}"/>
      <ellipse cx="${n(g.cx - g.pr)}" cy="${n(g.py + 0.5)}" rx="2.8" ry="1.9" fill="${NABE_RIM}"/>
      <ellipse cx="${n(g.cx + g.pr)}" cy="${n(g.py + 0.5)}" rx="2.8" ry="1.9" fill="${NABE_RIM}"/>
      <ellipse cx="${g.cx}" cy="${g.py}" rx="${g.pr}" ry="${n(g.pr * 0.3)}" fill="${NABE_RIM}"/>
      <ellipse cx="${g.cx}" cy="${n(g.py + 0.4)}" rx="${n(g.pr * 0.84)}" ry="${n(g.pr * 0.23)}" fill="${SOUP}"/>
      <circle cx="${n(g.cx - g.pr * 0.42)}" cy="${n(g.py - 0.4)}" r="${n(1.6 + 0.6 * fat)}" fill="#7fa85c"/>
      <circle cx="${n(g.cx + g.pr * 0.1)}" cy="${n(g.py + 1.2)}" r="${n(1.8 + 0.7 * fat)}" fill="#f6eddc"/>
      <circle cx="${n(g.cx + g.pr * 0.55)}" cy="${n(g.py - 0.6)}" r="${n(1.4 + 0.5 * fat)}" fill="#d96a4a"/>
    </g>`);

    const anim = o.animate === false ? '' : ' data-anim="1"';
    return `<svg class="rikishi-svg" viewBox="0 0 120 122" width="100%" height="100%"
      xmlns="http://www.w3.org/2000/svg" aria-hidden="true"${anim}>${parts.join('')}</svg>`;
  }

  /** 階級ごとの座標。太さ(fat)だけで全部が決まる */
  function geometry(stage) {
    const fat = (Math.max(1, Math.min(24, stage)) - 1) / 23;
    const cx = 60;
    const headR = n(12.5 + 6 * fat);
    const headCy = 30;
    const sy = n(headCy + headR + 4);
    const bt = n(11 + 18 * fat);
    const bb = n(14 + 22 * fat);
    const hy = 94;

    const torso =
      `M ${n(cx - bt)} ${sy} Q ${cx} ${n(sy - 7)} ${n(cx + bt)} ${sy}` +
      ` C ${n(cx + bt + 3)} ${n(sy + 14)}, ${n(cx + bb)} ${n(hy - 16)}, ${n(cx + bb)} ${hy}` +
      ` Q ${cx} ${n(hy + 9)} ${n(cx - bb)} ${hy}` +
      ` C ${n(cx - bb)} ${n(hy - 16)}, ${n(cx - bt - 3)} ${n(sy + 14)}, ${n(cx - bt)} ${sy} Z`;

    return {
      fat, cx, headR, headCy, sy, bt, bb, hy, torso,
      armW: n(6 + 4.5 * fat),
      legRx: n(bb * 1.3),
      legRy: n(9 + 4 * fat),
      legCy: 100,
      kneeRx: n(8.5 + 5 * fat),
      kneeRy: n(6.5 + 3.5 * fat),
      pr: n(14 + 6.5 * fat),
      py: 105,
      ph: n(10 + 3.5 * fat),
      shL: { x: n(cx - bt * 0.82), y: n(sy + 6) },
      shR: { x: n(cx + bt * 0.82), y: n(sy + 6) },
      handL: { x: n(cx - bb * 1.02), y: n(hy + 3) },
      handR: { x: n(cx + headR + 5), y: n(headCy + headR * 0.75) },
      eyeX: n(headR * 0.4),
      eyeY: n(headCy + headR * 0.1),
      mouthY: n(headCy + headR * 0.62),
      mouthRx: n(3.6 + 1.6 * fat)
    };
  }

  /**
   * 塩をまく手の位置（viewBox に対する 0..1）。
   * 塩まきモーションで左腕を約120度振り上げた先を求める。
   */
  function saltOrigin(stage) {
    const g = geometry(stage);
    const vx = g.handL.x - g.shL.x;
    const vy = g.handL.y - g.shL.y;
    const a = (120 * Math.PI) / 180;
    const x = g.shL.x + vx * Math.cos(a) - vy * Math.sin(a);
    const y = g.shL.y + vx * Math.sin(a) + vy * Math.cos(a);
    return { x: x / 120, y: y / 122 };
  }

  root.SUMO_RIKISHI = { svg: svg, saltOrigin: saltOrigin, geometry: geometry };
})(typeof self !== 'undefined' ? self : this);
