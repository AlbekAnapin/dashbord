/* =====================================================================
   charts.js — маленькая библиотека графиков на чистом SVG (без зависимостей)
   ===================================================================== */
(function (root) {
  'use strict';
  const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const PALETTE = ['#4e79a7', '#f28e2b', '#e15759', '#76b7b2', '#59a14f', '#edc948', '#b07aa1', '#ff9da7', '#9c755f', '#bab0ac', '#1f77b4', '#17becf', '#8c564b', '#7f7f7f'];
  const trunc = (s, n) => (s.length > n ? s.slice(0, n - 1) + '…' : s);
  const nf = new Intl.NumberFormat('ru-RU', { maximumFractionDigits: 0 });
  const nf1 = new Intl.NumberFormat('ru-RU', { minimumFractionDigits: 1, maximumFractionDigits: 1 });
  const fmtN = (v) => nf.format(Math.round(v));
  const fmt1 = (v) => nf1.format(v);

  function niceMax(v) {
    if (v <= 0) return 1;
    const p = Math.pow(10, Math.floor(Math.log10(v)));
    const f = v / p;
    const n = f <= 1 ? 1 : f <= 2 ? 2 : f <= 5 ? 5 : 10;
    return n * p;
  }

  // Горизонтальные столбики (в т.ч. составные). items: [{id,label,title,vals:[..]}]
  function hbar(items, o) {
    o = o || {};
    const colors = o.colors || [PALETTE[0]];
    const rowH = o.rowH || 22, labelW = o.labelW || 250, barW = o.barW || 320, valW = 52;
    const W = labelW + barW + valW, H = items.length * rowH + 6;
    const totals = items.map((i) => i.vals.reduce((a, b) => a + b, 0));
    const max = o.max || Math.max(...totals, 1e-9);
    const fmt = o.fmt || fmt1;
    const chars = Math.floor(labelW / 7.3);
    let s = `<svg class="chart" viewBox="0 0 ${W} ${H}" role="img" preserveAspectRatio="xMinYMin meet">`;
    items.forEach((it, idx) => {
      const y = idx * rowH + 3;
      const attr = it.id ? ` data-vuz="${esc(it.id)}" class="clk"` : '';
      s += `<g${attr}><title>${esc(it.title || it.label)}</title>`;
      s += `<text x="0" y="${y + rowH / 2 + 4}" class="lbl">${esc(trunc(it.label, chars))}</text>`;
      let x = labelW;
      it.vals.forEach((v, j) => {
        const w = Math.max(0, (v / max) * barW);
        if (w > 0) s += `<rect x="${x}" y="${y + 3}" width="${w}" height="${rowH - 8}" rx="2" fill="${colors[j % colors.length]}"><title>${esc((o.names && o.names[j]) || '')} ${fmt(v)}</title></rect>`;
        x += w;
      });
      s += `<text x="${x + 5}" y="${y + rowH / 2 + 4}" class="val">${fmt(totals[idx])}</text></g>`;
    });
    return s + '</svg>';
  }

  // Кольцевая диаграмма. items: [{label,value,color}]
  function donut(items, o) {
    o = o || {};
    const size = o.size || 200, r = size / 2 - 8, r2 = r * 0.6, cx = size / 2, cy = size / 2;
    const total = items.reduce((a, b) => a + b.value, 0);
    if (!total) return '<div class="muted">Нет данных</div>';
    let a0 = -Math.PI / 2, s = `<svg class="donut" viewBox="0 0 ${size} ${size}" role="img">`;
    items.forEach((it) => {
      if (it.value <= 0) return;
      const frac = it.value / total, a1 = a0 + frac * Math.PI * 2;
      const large = frac > 0.5 ? 1 : 0;
      const p = (rad, a) => `${cx + rad * Math.cos(a)} ${cy + rad * Math.sin(a)}`;
      const d = frac > 0.9999
        ? `M ${cx} ${cy - r} A ${r} ${r} 0 1 1 ${cx - 0.01} ${cy - r} L ${cx - 0.01} ${cy - r2} A ${r2} ${r2} 0 1 0 ${cx} ${cy - r2} Z`
        : `M ${p(r, a0)} A ${r} ${r} 0 ${large} 1 ${p(r, a1)} L ${p(r2, a1)} A ${r2} ${r2} 0 ${large} 0 ${p(r2, a0)} Z`;
      s += `<path d="${d}" fill="${it.color}"><title>${esc(it.label)}: ${fmtN(it.value)} (${fmt1(frac * 100)}%)</title></path>`;
      a0 = a1;
    });
    s += `<text x="${cx}" y="${cy - 2}" class="dn-total" text-anchor="middle">${esc(o.center || fmtN(total))}</text>`;
    if (o.sub) s += `<text x="${cx}" y="${cy + 14}" class="dn-sub" text-anchor="middle">${esc(o.sub)}</text>`;
    return s + '</svg>';
  }

  function legend(items, total) {
    return '<div class="legend">' + items.filter((i) => i.value > 0).map((i) =>
      `<div class="lg"><span class="sw" style="background:${i.color}"></span><span class="lt">${esc(i.label)}</span><span class="lv">${total ? fmt1((i.value / total) * 100) : '0,0'}%</span></div>`).join('') + '</div>';
  }

  // Линейный график. series: [{name,color,values:[..|null]}], xLabels
  function line(series, xLabels, o) {
    o = o || {};
    const W = o.width || 620, H = o.height || 240, T = 14, B = 30;
    const all = series.flatMap((s) => s.values).filter((v) => v !== null && v !== undefined);
    if (!all.length) return '<div class="muted">Нет данных</div>';
    const ymax = o.max !== undefined ? o.max : niceMax(Math.max(...all) * 1.05);
    const ymin = o.min !== undefined ? o.min : 0;
    const fmt = o.fmt || fmtN;
    const L = Math.max(52, 8 + 7 * Math.max(fmt(ymax).length, fmt(ymin).length)), R = 38;
    const n = xLabels.length, xs = (i) => L + (n === 1 ? (W - L - R) / 2 : (i * (W - L - R)) / (n - 1));
    const ys = (v) => T + (1 - (v - ymin) / (ymax - ymin || 1)) * (H - T - B);
    let s = `<svg class="chart line" viewBox="0 0 ${W} ${H}" role="img">`;
    for (let i = 0; i <= 4; i++) {
      const v = ymin + ((ymax - ymin) * i) / 4, y = ys(v);
      s += `<line x1="${L}" x2="${W - R}" y1="${y}" y2="${y}" class="grid"/><text x="${L - 6}" y="${y + 4}" text-anchor="end" class="axis">${fmt(v)}</text>`;
    }
    xLabels.forEach((t, i) => { s += `<text x="${xs(i)}" y="${H - 10}" text-anchor="middle" class="axis">${esc(t)}</text>`; });
    series.forEach((se) => {
      let d = '', pen = false;
      se.values.forEach((v, i) => { if (v === null || v === undefined) { pen = false; return; } d += `${pen ? 'L' : 'M'} ${xs(i)} ${ys(v)} `; pen = true; });
      s += `<path d="${d}" fill="none" stroke="${se.color}" stroke-width="2.2" stroke-linejoin="round"/>`;
      se.values.forEach((v, i) => { if (v !== null && v !== undefined) s += `<circle cx="${xs(i)}" cy="${ys(v)}" r="3.6" fill="${se.color}"><title>${esc(se.name)} · ${esc(xLabels[i])}: ${fmt(v)}</title></circle>`; });
    });
    s += '</svg>';
    if (series.length > 1 || o.legend) {
      s += '<div class="legend row">' + series.map((se) => `<div class="lg"><span class="sw" style="background:${se.color}"></span><span class="lt">${esc(se.name)}</span></div>`).join('') + '</div>';
    }
    return s;
  }

  // Гистограмма. values: [0..100]
  function hist(values, o) {
    o = o || {};
    const bins = o.bins || 10, W = 440, H = 170, L = 30, B = 24, T = 12;
    const counts = new Array(bins).fill(0);
    values.forEach((v) => { counts[Math.min(bins - 1, Math.max(0, Math.floor((v / 100) * bins)))]++; });
    const max = Math.max(...counts, 1), bw = (W - L - 6) / bins;
    let s = `<svg class="chart" viewBox="0 0 ${W} ${H}" role="img">`;
    counts.forEach((c, i) => {
      const h = (c / max) * (H - T - B), x = L + i * bw, y = H - B - h;
      const hue = 120 - ((i + 0.5) / bins) * 120;
      s += `<rect x="${x + 2}" y="${y}" width="${bw - 4}" height="${h}" rx="2" fill="hsl(${hue},60%,50%)"><title>${(i * 100) / bins}–${((i + 1) * 100) / bins}: ${c} ОВПО</title></rect>`;
      if (c) s += `<text x="${x + bw / 2}" y="${y - 3}" text-anchor="middle" class="axis">${c}</text>`;
      s += `<text x="${x + bw / 2}" y="${H - 8}" text-anchor="middle" class="axis">${(i * 100) / bins}</text>`;
    });
    return s + '</svg>';
  }

  root.Charts = { PALETTE, hbar, donut, legend, line, hist, esc, fmtN, fmt1, trunc };
})(window);
