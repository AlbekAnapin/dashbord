/* =====================================================================
   app.js — интерфейс дашборда: вкладки, таблицы, загрузка файлов, настройки.
   Все данные обрабатываются в браузере, никуда не отправляются.
   ===================================================================== */
(function () {
  'use strict';
  const C = window.VuzCore, Ch = window.Charts, esc = Ch.esc, fmtN = Ch.fmtN, fmt1 = Ch.fmt1;
  const $ = (s, r) => (r || document).querySelector(s);
  const $$ = (s, r) => Array.from((r || document).querySelectorAll(s));
  const LS = 'vuzdash.v1', LS_THEME = 'vuzdash.theme';
  const BUNDLE = window.VUZ_BUNDLE || null;
  const RED_ZONE = 50;
  const TABS = [['overview', 'Обзор'], ['rating', 'Рейтинг по качеству'], ['unified', 'Сводный рейтинг ошибок'],
    ['dynamics', 'Динамика'], ['vuz', 'ОВПО'], ['data', 'Данные и настройки']];
  const CRIT_COLORS = ['#4e79a7', '#f28e2b', '#b07aa1'];

  const state = {
    snaps: [], settings: C.defaultSettings(), tab: 'overview', cur: null, cmp: 'auto', vuz: null,
    sort: { rating: { col: 'total', dir: -1 }, unified: { col: 'total', dir: -1 } },
    search: { rating: '', unified: '' }, noDig: false, dynMetric: '__total', cache: null, hideZero: true, detSort: true, sug: { list: [], idx: 0 },
  };

  /* ------------------------------------------------------------ утилиты */
  const LEGAL = /^(Некоммерческое\s+акционерное\s+общество|Акционерное\s+общество|Товарищество\s+с\s+ограниченной\s+ответственностью|Республиканское\s+государственное\s+учреждение|Автономная\s+организация\s+образования|Организация\s+образования|Негосударственное\s+учреждение|Частное\s+учреждение|Государственное\s+учреждение|Учреждение|высшего\s+образования|Образования|образования)\s*/i;
  const shortName = (n) => {
    let s = n.replace(/["«»“”]/g, '').trim();
    for (let i = 0; i < 3 && LEGAL.test(s); i++) s = s.replace(LEGAL, '');
    return s || n;
  };
  const sgn = (v, f) => (v > 0 ? '+' : v < 0 ? '−' : '') + (f || fmt1)(Math.abs(v));
  const heatStyle = (s) => (s === null || s === undefined ? '' : ` style="background:hsla(${Math.max(0, 120 - s * 1.2)},70%,48%,.30)"`);
  const toast = (msg) => { const t = $('#toast'); t.textContent = msg; t.classList.add('show'); clearTimeout(toast.t); toast.t = setTimeout(() => t.classList.remove('show'), 3800); };
  const tick = () => new Promise((r) => setTimeout(r, 30));
  const loading = (on, msg) => { const l = $('#loading'); l.textContent = msg || 'Обработка…'; l.classList.toggle('show', !!on); };
  function download(name, content, mime) {
    const blob = content instanceof Blob ? content : new Blob([content], { type: mime || 'text/plain' });
    const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name; document.body.appendChild(a); a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 500);
  }
  function exportXlsx(name, aoa, sheet) {
    const ws = XLSX.utils.aoa_to_sheet(aoa); const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, sheet || 'Данные'); XLSX.writeFile(wb, name);
  }
  const sortedSnaps = () => state.snaps.slice().sort((a, b) => (a.date || '9999').localeCompare(b.date || '9999') || a.added.localeCompare(b.added));
  const curSnap = () => state.snaps.find((s) => s.id === state.cur) || null;
  function cmpSnap() {
    if (state.cmp === 'none') return null;
    const cur = curSnap(); if (!cur) return null;
    if (state.cmp !== 'auto') return state.snaps.find((s) => s.id === state.cmp && s.id !== cur.id) || null;
    const all = sortedSnaps(); const i = all.findIndex((s) => s.id === cur.id);
    return i > 0 ? all[i - 1] : null;
  }

  /* ------------------------------------------------------------ хранилище */
  function mergeSettings(s) {
    const d = C.defaultSettings(); if (!s) return d;
    return Object.assign(d, s, { critWeights: Object.assign(d.critWeights, s.critWeights), indWeights: Object.assign(d.indWeights, s.indWeights), ucolOn: Object.assign(d.ucolOn, s.ucolOn) });
  }
  function save() {
    const pack = (strip) => JSON.stringify({ basedOn: BUNDLE ? BUNDLE.generated : null, settings: state.settings, cur: state.cur,
      snaps: strip ? state.snaps.map((s) => (s.id === state.cur ? s : Object.assign({}, s, { tables: undefined }))) : state.snaps });
    try { localStorage.setItem(LS, pack(false)); }
    catch (e) {
      try { localStorage.setItem(LS, pack(true)); toast('Память браузера ограничена: детализация по полям сохранена только для текущего среза'); }
      catch (e2) { toast('Не удалось сохранить данные в браузере (хранилище переполнено)'); }
    }
  }
  function load() {
    let loc = null; try { loc = JSON.parse(localStorage.getItem(LS) || 'null'); } catch (e) { loc = null; }
    if (BUNDLE && (!loc || loc.basedOn !== BUNDLE.generated)) {
      state.snaps = BUNDLE.snapshots || []; state.settings = mergeSettings(BUNDLE.settings); state.published = true;
    } else if (loc) { state.snaps = loc.snaps || []; state.settings = mergeSettings(loc.settings); state.cur = loc.cur; }
    state.published = !!BUNDLE;
    const all = sortedSnaps(); if (!state.snaps.find((s) => s.id === state.cur)) state.cur = all.length ? all[all.length - 1].id : null;
    try { if (localStorage.getItem(LS_THEME) === 'dark') document.documentElement.dataset.theme = 'dark'; } catch (e) { /* ignore */ }
  }
  const invalidate = () => { state.cache = null; };

  /* ------------------------------------------------------------ расчёты */
  function derive() {
    if (state.cache) return state.cache;
    const snap = curSnap(); if (!snap) return null;
    const S = state.settings, keys = C.includedKeys(snap, S);
    const rated = C.rate(snap, keys, S), uni = C.unified(snap, keys, S);
    const D = { snap, keys, rated, uni, byKey: new Map(rated.rows.map((r) => [r.key, r])), uByKey: new Map(uni.rows.map((r) => [r.key, r])), cmp: cmpSnap(), delta: null, udelta: null };
    if (D.cmp) {
      const [rc, rn] = C.rateAll([D.cmp, snap], S);
      const mc = new Map(rc.rows.map((r) => [r.key, r])); D.delta = new Map();
      for (const r of rn.rows) { const o = mc.get(r.key); if (o) D.delta.set(r.key, { d: r.total - o.total, prev: o.rank, cur: r.rank, dK: r.K - o.K, dZ: r.Z - o.Z, dI: r.I - o.I }); }
      const uc = C.unified(D.cmp, C.includedKeys(D.cmp, S), S);
      D.udelta = new Map(uc.rows.map((r) => [r.key, r.total])); D.ugrandPrev = uc.grand;
      D.avgPrev = rc.rows.length ? rc.rows.reduce((s, r) => s + r.total, 0) / rc.rows.length : 0;
      D.avgCur = rn.rows.length ? rn.rows.reduce((s, r) => s + r.total, 0) / rn.rows.length : 0;
    }
    state.cache = D; return D;
  }
  const avg = (a) => (a.length ? a.reduce((x, y) => x + y, 0) / a.length : 0);
  function trendTag(d) {
    if (d === null || d === undefined) return '<span class="tag flat">—</span>';
    const t = state.settings.trend;
    if (d <= -t) return '<span class="tag good">▼ улучшение</span>';
    if (d >= t) return '<span class="tag bad">▲ ухудшение</span>';
    return '<span class="tag flat">без изменений</span>';
  }
  const deltaHtml = (d, f) => (d === null || d === undefined ? '' : `<span class="${d > 0.05 ? 'up' : d < -0.05 ? 'down' : 'flat'}">${sgn(d, f)}</span>`);

  /* ------------------------------------------------------------ шапка */
  function renderTop() {
    const all = sortedSnaps(), cur = curSnap(), cmp = cmpSnap();
    const opts = (sel, extra) => (extra || '') + all.map((s) => `<option value="${s.id}"${s.id === sel ? ' selected' : ''}>${esc(s.label)}</option>`).join('');
    $('#snapSel').innerHTML = opts(state.cur) || '<option>— нет срезов —</option>';
    $('#cmpSel').innerHTML = `<option value="auto"${state.cmp === 'auto' ? ' selected' : ''}>предыдущий${cmp && state.cmp === 'auto' ? ' (' + esc(cmp.label) + ')' : ''}</option><option value="none"${state.cmp === 'none' ? ' selected' : ''}>не сравнивать</option>` +
      all.filter((s) => !cur || s.id !== cur.id).map((s) => `<option value="${s.id}"${state.cmp === s.id ? ' selected' : ''}>${esc(s.label)}</option>`).join('');
    $('#tabs').innerHTML = TABS.map(([id, t]) => `<button data-tab="${id}" class="${state.tab === id ? 'on' : ''}">${t}</button>`).join('');
  }

  /* ------------------------------------------------------------ пустое состояние */
  function emptyHtml() {
    return `<div class="empty"><div class="drop" id="dropEmpty"><div class="big">Перетащите сюда Excel-файл среза</div>
      <div class="muted">или <a href="#" id="pickLink">выберите файл</a>. Можно сразу несколько срезов — тогда появится динамика.</div></div>
      <p class="muted">Файл разбирается прямо в браузере и никуда не отправляется.<br>Ожидается тот же формат, что и в «Не корректные данные»: листы ГБД ФЛ, ОСМС, Генерация номеров дипломов, Дублирование записей по ИИН, Оцифровка дипломов, GPA, Полнота данных ППС, МТБ, Госуслуги, Административный отчёт.</p></div>`;
  }

  /* ------------------------------------------------------------ ОБЗОР */
  function sourceTotals(D) {
    const S = state.settings, by = {};
    D.uni.rows.forEach((r) => C.UCOLS.forEach((u) => { if (S.ucolOn[u.id] && r.parts[u.id] !== null && !(state.noDig && u.src === 'Оцифровка')) by[u.src] = (by[u.src] || 0) + r.parts[u.id]; }));
    return Object.entries(by).map(([label, value], i) => ({ label, value })).sort((a, b) => b.value - a.value).map((x, i) => Object.assign(x, { color: Ch.PALETTE[i % Ch.PALETTE.length] }));
  }
  function viewOverview(D) {
    const S = state.settings, W = S.critWeights, ws = W['К'] + W['З'] + W['И'] || 1;
    const rows = D.rated.rows, n = rows.length;
    const avgTotal = avg(rows.map((r) => r.total)), red = rows.filter((r) => r.total >= RED_ZONE).length;
    const src = sourceTotals(D), srcTotal = src.reduce((a, b) => a + b.value, 0);
    const topSrc = src[0];
    const dErr = D.cmp && D.ugrandPrev ? (D.uni.grand / D.ugrandPrev - 1) * 100 : null;
    const dAvg = D.cmp ? D.avgCur - D.avgPrev : null;
    const kpi = (l, v, d, sub) => `<div class="card kpi"><div class="l">${l}</div><div class="v">${v}</div><div class="d">${d || ''}${sub ? `<span class="muted"> ${sub}</span>` : ''}</div></div>`;
    const cmpTxt = D.cmp ? 'к ' + esc(D.cmp.label) : '';
    const top = rows.slice(0, 15).map((r) => ({ id: r.key, label: shortName(r.name), title: r.name + ` — итог ${fmt1(r.total)}`, vals: [(r.K * W['К']) / ws, (r.Z * W['З']) / ws, (r.I * W['И']) / ws] }));
    const idxAvg = C.INDICATORS.map((ind) => ({ ind, a: avg(rows.map((r) => r.scores[ind.id]).filter((v) => v !== null)) })).sort((a, b) => b.a - a.a)
      .map((x) => ({ label: `${x.ind.id} ${x.ind.label}`, vals: [x.a] }));
    const allSnaps = sortedSnaps();
    let dynHtml = '';
    if (allSnaps.length > 1) {
      const RA = C.rateAll(allSnaps, S);
      dynHtml = `<div class="card"><h2>Динамика среднего балла по срезам</h2><div class="sub">Единая шкала по всем срезам: чем ниже, тем лучше</div>${Ch.line([
        { name: 'Корректность', color: CRIT_COLORS[0], values: RA.map((r) => avg(r.rows.map((x) => x.K))) },
        { name: 'Заполненность', color: CRIT_COLORS[1], values: RA.map((r) => avg(r.rows.map((x) => x.Z))) },
        { name: 'Исполнение', color: CRIT_COLORS[2], values: RA.map((r) => avg(r.rows.map((x) => x.I))) },
        { name: 'Итого', color: '#222', values: RA.map((r) => avg(r.rows.map((x) => x.total))) }], allSnaps.map((s) => s.label), { fmt: fmt1, height: 230 })}</div>`;
    }
    return `
    <div class="grid g4">
      ${kpi('ОВПО в расчёте', fmtN(n), '', `из ${fmtN(Object.keys(D.snap.names).length)} в файле`)}
      ${kpi('Всего ошибок', fmtN(D.uni.grand), dErr === null ? '' : deltaHtml(dErr, (v) => fmt1(v) + '%'), cmpTxt)}
      ${kpi('Средний итоговый балл', fmt1(avgTotal), dAvg === null ? '' : deltaHtml(dAvg), cmpTxt)}
      ${kpi('«Красная зона» (балл ≥ ' + RED_ZONE + ')', fmtN(red), '', `ОВПО (${fmt1((red / (n || 1)) * 100)}%)`)}
      ${kpi('Главный источник ошибок', topSrc ? esc(topSrc.label) : '—', '', topSrc ? fmt1((topSrc.value / (srcTotal || 1)) * 100) + '% всех ошибок' : '')}
    </div>
    <div class="grid g2" style="margin-top:16px">
      <div class="card"><h2>Топ-15 ОВПО с наибольшим итоговым баллом</h2>
        <div class="sub">Вклад критериев: <span class="sw" style="background:${CRIT_COLORS[0]}"></span> корректность <span class="sw" style="background:${CRIT_COLORS[1]}"></span> заполненность <span class="sw" style="background:${CRIT_COLORS[2]}"></span> исполнение · клик по названию — карточка ОВПО</div>
        ${Ch.hbar(top, { colors: CRIT_COLORS, names: ['Корректность', 'Заполненность', 'Исполнение'], labelW: 300, barW: 280 })}</div>
      <div class="card"><h2>Структура ошибок по источникам</h2>
        <div class="sub"><label><input type="checkbox" id="noDig" ${state.noDig ? 'checked' : ''}> без оцифровки дипломов (она обычно занимает основную долю)</label></div>
        <div class="donut-wrap">${Ch.donut(src, { center: fmtN(srcTotal), sub: 'ошибок' })}${Ch.legend(src, srcTotal)}</div></div>
    </div>
    <div class="grid g3" style="margin-top:16px">
      <div class="card"><h2>Распределение итогового балла</h2><div class="sub">Число ОВПО по диапазонам балла</div>${Ch.hist(rows.map((r) => r.total))}</div>
      <div class="card"><h2>Средний балл по критериям</h2><div class="sub">0 — нет проблем, 100 — худший</div>
        ${Ch.hbar([{ label: 'Корректность', vals: [avg(rows.map((r) => r.K))] }, { label: 'Заполненность', vals: [avg(rows.map((r) => r.Z))] }, { label: 'Исполнение', vals: [avg(rows.map((r) => r.I))] }], { colors: ['#4e79a7'], max: 100, labelW: 120, barW: 200 })}
        <div class="sub" style="margin-top:8px">ОВПО без ошибок (итоговый балл 0): <b>${rows.filter((r) => r.total === 0).length}</b></div></div>
      ${dynHtml || `<div class="card"><h2>Динамика</h2><div class="muted">Загрузите ещё один срез на вкладке «Данные и настройки» — здесь появится динамика.</div></div>`}
    </div>
    <div class="card" style="margin-top:16px"><h2>Какие показатели в среднем «просаживают» рейтинг</h2><div class="sub">Средний балл по показателям (по всем ОВПО, где показатель применим)</div>
      ${Ch.hbar(idxAvg, { colors: ['#4e79a7'], max: 100, labelW: 400, barW: 520, rowH: 22 })}</div>`;
  }

  /* ------------------------------------------------------------ универсальная таблица */
  function sortRows(rows, cols, st) {
    const c = cols.find((x) => x.id === st.col) || cols[0]; const g = c.sortVal || c.get;
    return rows.slice().sort((a, b) => {
      const x = g(a), y = g(b);
      if (x === null || x === undefined) return 1; if (y === null || y === undefined) return -1;
      if (typeof x === 'string') return st.dir * x.localeCompare(y, 'ru');
      return st.dir * (x - y) || a.name.localeCompare(b.name, 'ru');
    });
  }
  function tableHtml(cols, rows, st, id) {
    const th = cols.map((c) => `<th class="${c.cls || ''}${st.col === c.id ? ' sorted' : ''}" data-sort="${id}:${c.id}" ${st.col === c.id ? `data-dir="${st.dir > 0 ? '▲' : '▼'}"` : ''} title="${esc(c.title || c.label)}">${c.label}</th>`).join('');
    const body = rows.map((r) => `<tr class="clk-row" data-vuz="${esc(r.key)}">${cols.map((c) => c.html(r)).join('')}</tr>`).join('');
    return `<table><thead><tr>${th}</tr></thead><tbody>${body}</tbody></table>`;
  }
  const nameCell = (r) => `<td class="name" title="${esc(r.name)}"><b>${esc(shortName(r.name))}</b></td>`;
  const filterRows = (rows, q) => { q = C.key(q || ''); return q ? rows.filter((r) => C.key(r.name).includes(q)) : rows; };

  /* ------------------------------------------------------------ РЕЙТИНГ ПО КАЧЕСТВУ */
  function ratingCols(D) {
    const cols = [
      { id: 'rank', label: '№', cls: 'rank', get: (r) => r.rank, html: (r) => `<td class="rank">${r.rank}</td>` },
      { id: 'name', label: 'ОВПО', cls: 'name', get: (r) => r.name, html: nameCell },
      { id: 'K', label: 'Корректность', get: (r) => r.K, html: (r) => `<td${heatStyle(r.K)}>${fmt1(r.K)}</td>` },
      { id: 'Z', label: 'Заполненность', get: (r) => r.Z, html: (r) => `<td${heatStyle(r.Z)}>${fmt1(r.Z)}</td>` },
      { id: 'I', label: 'Исполнение', get: (r) => r.I, html: (r) => `<td${heatStyle(r.I)}>${fmt1(r.I)}</td>` },
      { id: 'total', label: 'ИТОГО (балл)', get: (r) => r.total, html: (r) => `<td${heatStyle(r.total)}><b>${fmt1(r.total)}</b></td>` },
      { id: 'rankK', label: 'Ранг: корректн.', title: 'Место по корректности (1 — худшее)', get: (r) => r.rankK, html: (r) => `<td>${r.rankK}</td>` },
      { id: 'rankZ', label: 'Ранг: заполн.', get: (r) => r.rankZ, html: (r) => `<td>${r.rankZ}</td>` },
      { id: 'rankI', label: 'Ранг: исполн.', get: (r) => r.rankI, html: (r) => `<td>${r.rankI}</td>` },
    ];
    if (D.delta) {
      const dl = (r) => { const x = D.delta.get(r.key); return x ? x.d : null; };
      cols.push({ id: 'delta', label: `Δ балла к ${esc(D.cmp.label)}`, title: 'Изменение итогового балла на общей шкале двух срезов (минус — лучше)', get: dl, html: (r) => `<td>${deltaHtml(dl(r)) || '—'}</td>` });
      cols.push({ id: 'prevRank', label: 'Место в прошлом срезе', get: (r) => { const x = D.delta.get(r.key); return x ? x.prev : null; }, html: (r) => { const x = D.delta.get(r.key); return `<td>${x ? x.prev : '—'}</td>`; } });
      cols.push({ id: 'trend', label: 'Тренд', sortVal: dl, get: dl, html: (r) => `<td>${trendTag(dl(r))}</td>` });
    }
    return cols;
  }
  function viewRating(D) {
    const cols = ratingCols(D);
    return `<div class="toolbar"><input type="text" id="q-rating" placeholder="Поиск по названию…" value="${esc(state.search.rating)}">
      <button class="btn ghost" id="exp-rating">Скачать Excel</button>
      <span class="mini">Баллы 0–100: чем выше, тем хуже. Ранг 1 — худший. Клик по заголовку — сортировка, по строке — карточка ОВПО. Веса критериев — на вкладке «Данные и настройки».</span></div>
      <div class="tbl-wrap" id="tbl-rating"></div>`;
  }
  function fillRating(D) {
    const cols = ratingCols(D), rows = sortRows(filterRows(D.rated.rows, state.search.rating), cols, state.sort.rating);
    $('#tbl-rating').innerHTML = tableHtml(cols, rows, state.sort.rating, 'rating');
  }
  function exportRating(D) {
    const aoa = [['№', 'ОВПО', 'Корректность', 'Заполненность', 'Исполнение', 'ИТОГО (балл)', 'Ранг: корректность', 'Ранг: заполненность', 'Ранг: исполнение', D.delta ? 'Δ балла к ' + D.cmp.label : null].filter((x) => x !== null)];
    D.rated.rows.forEach((r) => { const row = [r.rank, r.name, +r.K.toFixed(2), +r.Z.toFixed(2), +r.I.toFixed(2), +r.total.toFixed(2), r.rankK, r.rankZ, r.rankI]; if (D.delta) { const x = D.delta.get(r.key); row.push(x ? +x.d.toFixed(2) : null); } aoa.push(row); });
    exportXlsx(`Рейтинг_по_качеству_${D.snap.label.replace(/[^\d.]/g, '')}.xlsx`, aoa, 'Рейтинг');
  }

  /* ------------------------------------------------------------ СВОДНЫЙ РЕЙТИНГ ОШИБОК */
  function unifiedCols(D) {
    const S = state.settings, max = Math.max(...D.uni.rows.map((r) => r.total), 1);
    const cols = [
      { id: 'rank', label: '№', cls: 'rank', get: (r) => r.rank, html: (r) => `<td class="rank">${r.rank}</td>` },
      { id: 'name', label: 'ОВПО', cls: 'name', get: (r) => r.name, html: nameCell },
      { id: 'total', label: 'Общее количество ошибок', get: (r) => r.total, html: (r) => `<td class="bar-cell" style="background-image:linear-gradient(90deg,rgba(78,121,167,.28),rgba(78,121,167,.28));background-size:${(r.total / max) * 100}% 70%"><b>${fmtN(r.total)}</b></td>` },
      { id: 'share', label: '% от общего числа ошибок', title: 'Доля ОВПО в общем объёме ошибок всех ОВПО', get: (r) => r.share, html: (r) => `<td>${fmt1(r.share)}%</td>` },
      { id: 'cum', label: 'Накопленный %', title: 'Накопленная доля при сортировке по убыванию (принцип Парето)', get: (r) => r.cum, html: (r) => `<td>${fmt1(r.cum)}%</td>` },
    ];
    if (D.udelta) {
      const dp = (r) => { const p = D.udelta.get(r.key); return p === undefined ? null : p === 0 ? (r.total === 0 ? 0 : null) : (r.total / p - 1) * 100; };
      cols.push({ id: 'dpct', label: `Δ ошибок к ${esc(D.cmp.label)}, %`, get: dp, html: (r) => { const v = dp(r); return `<td>${v === null ? '—' : deltaHtml(v, (x) => fmt1(x) + '%')}</td>`; } });
    }
    cols.push({ id: 'osmsRate', label: 'ОСМС: ошибок на 100 записей', title: 'Ошибки ОСМС / число записей (студентов) × 100', get: (r) => r.osmsRate, html: (r) => `<td>${r.osmsRate === null ? '—' : fmt1(r.osmsRate)}</td>` });
    C.UCOLS.forEach((u) => cols.push({ id: u.id, label: u.label + (S.ucolOn[u.id] ? '' : ' (не в сумме)'), title: u.label, get: (r) => r.parts[u.id], html: (r) => `<td${S.ucolOn[u.id] ? '' : ' class="muted"'}>${r.parts[u.id] === null ? '—' : fmtN(r.parts[u.id])}</td>` }));
    return cols;
  }
  function viewUnified(D) {
    const S = state.settings;
    return `<div class="toolbar"><input type="text" id="q-unified" placeholder="Поиск по названию…" value="${esc(state.search.unified)}">
      <button class="btn ghost" id="exp-unified">Скачать Excel</button></div>
      <div class="note">Какие источники входят в «Общее количество ошибок» (нажмите, чтобы включить/выключить):
        <div class="chips" style="margin-top:8px">${C.UCOLS.map((u) => `<span class="chip ${S.ucolOn[u.id] ? 'on' : ''}" data-ucol="${u.id}">${esc(u.label)}</span>`).join('')}</div></div>
      <div class="mini" style="margin:6px 0">Всего ошибок по ${fmtN(D.uni.rows.length)} ОВПО: <b>${fmtN(D.uni.grand)}</b>. Единицы в источниках разные (записи, поля, дипломы, отчёты), поэтому это объём проблем, а не число уникальных ошибок.</div>
      <div class="tbl-wrap" id="tbl-unified"></div>`;
  }
  function fillUnified(D) {
    const cols = unifiedCols(D), rows = sortRows(filterRows(D.uni.rows, state.search.unified), cols, state.sort.unified);
    $('#tbl-unified').innerHTML = tableHtml(cols, rows, state.sort.unified, 'unified');
  }
  function exportUnified(D) {
    const aoa = [['№', 'ОВПО', 'Общее количество ошибок', '% от общего', 'Накопленный %', 'ОСМС: ошибок на 100 записей'].concat(C.UCOLS.map((u) => u.label + (state.settings.ucolOn[u.id] ? '' : ' (не в сумме)')))];
    D.uni.rows.forEach((r) => aoa.push([r.rank, r.name, r.total, +r.share.toFixed(3), +r.cum.toFixed(3), r.osmsRate === null ? null : +r.osmsRate.toFixed(2)].concat(C.UCOLS.map((u) => r.parts[u.id]))));
    exportXlsx(`Сводный_рейтинг_ошибок_${D.snap.label.replace(/[^\d.]/g, '')}.xlsx`, aoa, 'Сводный рейтинг');
  }

  /* ------------------------------------------------------------ ДИНАМИКА */
  function viewDynamics(D) {
    const all = sortedSnaps(), S = state.settings;
    if (all.length < 2) return `<div class="card"><h2>Динамика</h2><p class="muted">Для динамики нужно минимум два среза. Загрузите ещё один файл на вкладке «Данные и настройки».</p></div>`;
    const RA = C.rateAll(all, S), labels = all.map((s) => s.label);
    const UA = all.map((s) => C.unified(s, C.includedKeys(s, S), S));
    const colSum = (u, i) => ((all[i].missing || []).includes(u.sheet) || UA[i].rows.every((r) => r.parts[u.id] === null) ? null : UA[i].rows.reduce((a, r) => a + (r.parts[u.id] || 0), 0));
    const metric = state.dynMetric;
    const series = metric === '__total' ? [{ name: 'Всего ошибок (по включённым источникам)', color: '#4e79a7', values: UA.map((u) => u.grand) }]
      : [{ name: (C.UCOLS.find((u) => u.id === metric) || {}).label, color: '#e15759', values: UA.map((_, i) => colSum(C.UCOLS.find((u) => u.id === metric), i)) }];
    const rowsT = C.UCOLS.map((u) => ({ u, v: all.map((_, i) => colSum(u, i)) }));
    const pc = (a, b) => (a && b !== null && a !== null ? (b / a - 1) * 100 : null);
    const pcHtml = (p) => (p === null ? '—' : deltaHtml(p, (x) => fmt1(x) + '%'));
    const tbl = `<table><thead><tr><th class="name">Источник ошибок (сумма по ОВПО)</th>${labels.map((l) => `<th>${esc(l)}</th>`).join('')}<th>Δ к пред. срезу</th><th>Δ к первому</th></tr></thead><tbody>` +
      rowsT.map(({ u, v }) => `<tr><td class="name">${esc(u.label)}${S.ucolOn[u.id] ? '' : ' <span class="muted">(не в сумме)</span>'}</td>${v.map((x) => `<td>${x === null ? '—' : fmtN(x)}</td>`).join('')}<td>${pcHtml(pc(v[v.length - 2], v[v.length - 1]))}</td><td>${pcHtml(pc(v[0], v[v.length - 1]))}</td></tr>`).join('') +
      `<tr><td class="name"><b>ВСЕГО (по включённым источникам)</b></td>${UA.map((u) => `<td><b>${fmtN(u.grand)}</b></td>`).join('')}<td>${pcHtml(pc(UA[UA.length - 2].grand, UA[UA.length - 1].grand))}</td><td>${pcHtml(pc(UA[0].grand, UA[UA.length - 1].grand))}</td></tr></tbody></table>`;
    let movers = '';
    if (D.delta) {
      const arr = Array.from(D.delta.entries()).map(([k, x]) => Object.assign({ key: k, name: D.snap.names[k] }, x));
      const best = arr.slice().sort((a, b) => a.d - b.d).slice(0, 10), worst = arr.slice().sort((a, b) => b.d - a.d).slice(0, 10);
      const list = (xs) => `<table><thead><tr><th class="name">ОВПО</th><th>Δ балла</th><th>Место: было → стало</th></tr></thead><tbody>${xs.map((x) => `<tr class="clk-row" data-vuz="${esc(x.key)}"><td class="name" title="${esc(x.name)}">${esc(shortName(x.name))}</td><td>${deltaHtml(x.d)}</td><td>${x.prev} → ${x.cur}</td></tr>`).join('')}</tbody></table>`;
      movers = `<div class="grid g2" style="margin-top:16px"><div class="card"><h2>Лучшая динамика (к ${esc(D.cmp.label)})</h2><div class="tbl-wrap" style="max-height:none">${list(best)}</div></div>
        <div class="card"><h2>Ухудшение (к ${esc(D.cmp.label)})</h2><div class="tbl-wrap" style="max-height:none">${list(worst)}</div></div></div>`;
    } else movers = '<div class="note" style="margin-top:12px">Выберите срез для сравнения в шапке страницы — здесь появятся ОВПО с лучшей и худшей динамикой.</div>';
    return `<div class="grid g2">
      <div class="card"><h2>Средний балл по критериям</h2><div class="sub">Единая шкала по всем срезам · чем ниже, тем лучше</div>${Ch.line([
        { name: 'Корректность', color: CRIT_COLORS[0], values: RA.map((r) => avg(r.rows.map((x) => x.K))) },
        { name: 'Заполненность', color: CRIT_COLORS[1], values: RA.map((r) => avg(r.rows.map((x) => x.Z))) },
        { name: 'Исполнение', color: CRIT_COLORS[2], values: RA.map((r) => avg(r.rows.map((x) => x.I))) },
        { name: 'Итого', color: '#222', values: RA.map((r) => avg(r.rows.map((x) => x.total))) }], labels, { fmt: fmt1 })}</div>
      <div class="card"><h2>Число ошибок по срезам</h2><div class="sub"><select id="dynMetric"><option value="__total">Всего (включённые источники)</option>${C.UCOLS.map((u) => `<option value="${u.id}"${metric === u.id ? ' selected' : ''}>${esc(u.label)}</option>`).join('')}</select></div>${Ch.line(series, labels, { legend: false })}</div>
    </div>${movers}
    <div class="card" style="margin-top:16px"><h2>Ошибки по источникам и срезам</h2><div class="sub">«—» — в этом срезе нет соответствующего листа. Срезы в старом формате (до 30.10) устроены иначе (например, в оцифровке считалось меньше полей), поэтому сравнение с ними ориентировочное.</div><div class="tbl-wrap" style="max-height:none">${tbl}</div></div>`;
  }

  /* ------------------------------------------------------------ КАРТОЧКА ОВПО */
  const SECTIONS = [['gbd', 'ГБД ФЛ'], ['osms', 'ОСМС'], ['gen', 'Генерация номеров дипломов'], ['dup', 'Дубли ИИН'], ['dig', 'Оцифровка дипломов'],
    ['gpa', 'GPA'], ['pps', 'Полнота данных ППС'], ['mtb', 'Материально-техническая база'], ['gos', 'Государственные услуги'], ['admin', 'Административные отчёты']];
  const SEC_UNIT = { gbd: 'ошибок', osms: 'ошибок', gen: 'дипломов', dup: 'записей-дублей', dig: 'ошибок', gpa: 'записей без GPA', pps: 'ошибок', mtb: 'проблем', gos: 'просроченных заявок', admin: 'неполных отчётов' };
  const SHORT = { gbd_stud: 'ГБД: обуч.', gbd_grad: 'ГБД: выпуск.', gbd_pps: 'ГБД: ППС', osms_stud: 'ОСМС: обуч.', osms_exp: 'ОСМС: отчисл.', osms_grad: 'ОСМС: выпуск.', dig_main: 'Оцифровка', pps: 'ППС', dup_stud: 'Дубли: обуч.', dup_pps: 'Дубли: ППС', gos: 'Госуслуги' };

  // несколько строк одного ОВПО в таблице-списке складываем в одну
  function collapse(t, rows) {
    const out = t.cols.map((c) => (c.t === 's' ? [] : 0));
    rows.forEach((r) => r.forEach((v, i) => { const c = t.cols[i]; if (c.t === 's') { if (v) out[i].push(v); } else out[i] += (v || 0); }));
    return out.map((v) => (Array.isArray(v) ? v.join('\n') : v));
  }
  // суммы по каждому столбцу по всем ОВПО в расчёте (для колонок «всего» и «доля ОВПО»)
  function colTotals(D, t) {
    D.ct = D.ct || {}; if (D.ct[t.id]) return D.ct[t.id];
    const sums = t.cols.map(() => 0);
    D.keys.forEach((k) => { const rows = t.rows[k]; if (!rows) return; const v = collapse(t, rows); t.cols.forEach((c, i) => { if (c.t === 'n') sums[i] += v[i]; }); });
    return (D.ct[t.id] = { sums });
  }
  const secTotal = (D, key, sheet) => { const u = D.uByKey.get(key); return u ? C.UCOLS.filter((c) => c.sheet === sheet).reduce((s, c) => s + (u.parts[c.id] || 0), 0) : 0; };
  const noRows = '<div class="muted">По этой таблице у ОВПО записей нет — ошибок не выявлено.</div>';

  function partnersHtml(label, text) {
    const lines = text.split('\n').map((x) => x.trim()).filter(Boolean).map((x) => { const m = /^(.*?)\s+—\s+(\d+)$/.exec(x); return m ? { n: m[1].replace(/["«»]/g, ''), c: +m[2] } : { n: x, c: null }; })
      .sort((p, q) => (q.c || 0) - (p.c || 0));
    return `<details open style="margin-top:8px"><summary>${esc(label)} — ${lines.length}</summary><div class="tbl-wrap" style="max-height:300px;margin-top:6px"><table class="fieldtbl"><tbody>${lines.map((l) => `<tr><td class="lbl">${esc(l.n)}</td><td>${l.c === null ? '' : fmtN(l.c)}</td></tr>`).join('')}</tbody></table></div></details>`;
  }
  // таблица «по полям»: название поля · число ошибок · % от итога · полоса · всего по всем ОВПО · доля этого ОВПО
  function listTableHtml(D, t, key) {
    const rows = t.rows[key]; if (!rows || !rows.length) return noRows;
    const v = collapse(t, rows), ti = t.cols.findIndex((c) => c.total);
    const items = t.cols.map((c, i) => ({ c, i, val: v[i] }));
    let fields = items.filter((e) => e.c.t === 'n' && !e.c.total);
    const total = ti >= 0 ? v[ti] : fields.reduce((s, e) => s + e.val, 0);
    const infos = items.filter((e) => e.c.t === 'i'), texts = items.filter((e) => e.c.t === 's' && e.val);
    const ct = colTotals(D, t);
    const zeros = fields.filter((e) => e.val === 0).length;
    if (state.detSort) fields = fields.slice().sort((p, q) => q.val - p.val);
    const shown = state.hideZero ? fields.filter((e) => e.val > 0) : fields;
    const max = Math.max(...fields.map((e) => e.val), 1);
    const body = shown.map((e) => `<tr class="${e.val ? '' : 'zero'}"><td class="lbl"${e.c.o ? ` title="Как в исходном файле: ${esc(e.c.o)}"` : ''}>${esc(e.c.h)}</td><td><b>${fmtN(e.val)}</b></td><td>${total ? fmt1((e.val / total) * 100) + '%' : '—'}</td>
      <td class="barc"><div class="mb" style="width:${(e.val / max) * 100}%"></div></td><td class="muted">${fmtN(ct.sums[e.i])}</td><td class="muted">${ct.sums[e.i] ? fmt1((e.val / ct.sums[e.i]) * 100) + '%' : '—'}</td></tr>`).join('');
    return `${infos.length ? `<div class="mini" style="margin-bottom:6px">${infos.map((e) => `${esc(e.c.h)}: <b>${e.val === null ? '—' : fmtN(e.val)}</b>`).join(' · ')}</div>` : ''}
      ${shown.length ? `<div class="tbl-wrap" style="max-height:none"><table class="fieldtbl"><thead><tr><th class="static lbl">Поле</th><th class="static">Ошибок</th><th class="static" title="Доля от итога таблицы">% от итога</th><th class="static"></th><th class="static" title="Сумма по всем ОВПО в расчёте">Всего по всем ОВПО</th><th class="static" title="Какая часть общего числа ошибок по этому полю приходится на данное ОВПО">Доля ОВПО</th></tr></thead>
      <tbody>${body}</tbody>${ti >= 0 ? `<tfoot><tr><td class="lbl"><b>Итого по таблице</b></td><td><b>${fmtN(total)}</b></td><td></td><td></td><td class="muted">${fmtN(ct.sums[ti])}</td><td class="muted">${ct.sums[ti] ? fmt1((total / ct.sums[ti]) * 100) + '%' : '—'}</td></tr></tfoot>` : ''}</table></div>` : '<div class="muted">Ошибок по полям нет.</div>'}
      ${state.hideZero && zeros ? `<div class="mini" style="margin-top:4px">Скрыто полей без ошибок: ${zeros}</div>` : ''}
      ${texts.map((e) => partnersHtml(e.c.h, e.val)).join('')}`;
  }
  const statusCls = (s) => (/^данные заполнены/i.test(s) ? 'good' : /^нет/i.test(s) ? 'bad' : 'warn');
  function adminHtml(t, rows) {
    const counts = {}; rows.forEach((r) => { const s = t.stat[r[3]]; counts[s] = (counts[s] || 0) + 1; });
    const list = state.hideZero ? rows.filter((r) => statusCls(t.stat[r[3]]) !== 'good') : rows;
    return `<div class="chips" style="margin-bottom:8px">${Object.entries(counts).map(([s, n]) => `<span class="tag ${statusCls(s)}">${esc(s)}: ${n}</span>`).join('')}</div>
      <div class="tbl-wrap" style="max-height:420px"><table class="fieldtbl"><thead><tr><th class="static lbl">Отчёт</th><th class="static lbl">Статус</th><th class="static">Таблиц</th><th class="static">С данными</th><th class="static">Без данных</th></tr></thead><tbody>
      ${list.map((r) => `<tr><td class="lbl"><b>${esc(r[0])}</b> ${esc(t.names[r[0]] || '')}</td><td class="lbl"><span class="tag ${statusCls(t.stat[r[3]])}">${esc(t.stat[r[3]])}</span></td><td>${fmtN(r[1])}</td><td>${fmtN(r[2])}</td><td>${fmtN(r[4])}</td></tr>`).join('')}</tbody></table></div>
      ${state.hideZero && list.length < rows.length ? `<div class="mini" style="margin-top:4px">Скрыто полностью заполненных отчётов: ${rows.length - list.length}</div>` : ''}`;
  }
  // таблица-сетка: строки (степени, общежития, категории) × столбцы (годы, поля)
  function gridTableHtml(D, t, key) {
    const rows = t.rows[key]; if (!rows || !rows.length) return noRows;
    if (t.id === 'admin') return adminHtml(t, rows);
    const cell = (c, v) => {
      if (c.t === 's') { const x = String(v || ''); return `<td class="lbl">${/^нет данных$/i.test(x) ? '<span class="tag bad">нет данных</span>' : /^есть$/i.test(x) ? '<span class="tag good">есть</span>' : esc(x) || '—'}</td>`; }
      return `<td>${v === null || v === undefined ? '—' : fmtN(v)}</td>`;
    };
    const hasTotals = rows.length > 1 && t.cols.some((c) => c.t === 'n');
    const sums = t.cols.map((c, i) => (c.t === 'n' ? rows.reduce((s, r) => s + (r[i] || 0), 0) : null));
    return `<div class="tbl-wrap" style="max-height:none"><table class="fieldtbl"><thead><tr>${t.cols.map((c) => `<th class="static${c.t === 's' ? ' lbl' : ''}">${esc(c.h)}</th>`).join('')}</tr></thead>
      <tbody>${rows.map((r) => `<tr>${t.cols.map((c, i) => cell(c, r[i])).join('')}</tr>`).join('')}</tbody>
      ${hasTotals ? `<tfoot><tr>${t.cols.map((c, i) => (i === 0 ? '<td class="lbl"><b>Итого</b></td>' : sums[i] === null ? '<td></td>' : `<td><b>${fmtN(sums[i])}</b></td>`)).join('')}</tr></tfoot>` : ''}</table></div>`;
  }
  function topFieldsHtml(D, key) {
    const items = [];
    Object.values(D.snap.tables || {}).forEach((t) => {
      if (t.layout !== 'list' || !SHORT[t.id]) return;
      const rows = t.rows[key]; if (!rows) return;
      const v = collapse(t, rows);
      t.cols.forEach((c, i) => { if (c.t === 'n' && !c.total && v[i] > 0) items.push({ label: `${SHORT[t.id]} · ${c.h}`, title: `${t.title}: ${c.h}`, vals: [v[i]] }); });
    });
    items.sort((a, b) => b.vals[0] - a.vals[0]);
    if (!items.length) return '<div class="muted">Ошибок по полям нет.</div>';
    return Ch.hbar(items.slice(0, 12), { colors: ['#e15759'], labelW: 470, barW: 300, rowH: 22, fmt: fmtN });
  }
  function sectionsHtml(D, key) {
    const tables = D.snap.tables;
    if (!tables || !Object.keys(tables).length) return '<div class="note warn">Для этого среза нет детализации по полям (он сохранён старой версией сайта). Загрузите исходный Excel-файл заново на вкладке «Данные и настройки».</div>';
    const secs = SECTIONS.map(([id, title]) => ({ id, title, ts: Object.values(tables).filter((t) => t.sec === id && !(t.note && !(t.rows[key] || []).length)) })).filter((s) => s.ts.length);
    const nav = secs.map((s) => { const tot = secTotal(D, key, s.id); return `<span class="chip" data-goto="sec-${s.id}">${esc(s.title)} <b class="${tot ? 'up' : 'down'}">${fmtN(tot)}</b></span>`; }).join('');
    const body = secs.map((s) => { const tot = secTotal(D, key, s.id);
      return `<div class="card sec" id="sec-${s.id}"><h2>${esc(s.title)} <span class="tag ${tot ? 'bad' : 'good'}">${fmtN(tot)} ${SEC_UNIT[s.id]}</span></h2>
        ${s.ts.map((t) => `<div class="subtab"><h3>${esc(t.title)}</h3>${t.note ? `<div class="mini">${esc(t.note)}</div>` : ''}${t.layout === 'list' ? listTableHtml(D, t, key) : gridTableHtml(D, t, key)}</div>`).join('')}</div>`; }).join('');
    return `<div class="chips secnav">${nav}</div><div class="grid" style="gap:16px;margin-top:12px">${body}</div>`;
  }

  function viewVuz(D) {
    const rows = D.rated.rows;
    if (!state.vuz || !D.byKey.has(state.vuz)) state.vuz = rows.length ? rows[0].key : null;
    const r = D.byKey.get(state.vuz); if (!r) return '<div class="card">Нет данных</div>';
    const u = D.uByKey.get(state.vuz), S = state.settings, all = sortedSnaps(), n = rows.length;
    const dl = D.delta ? D.delta.get(state.vuz) : null;
    const opts = rows.map((x) => `<option value="${esc(x.key)}"${x.key === state.vuz ? ' selected' : ''}>${x.rank}. ${esc(shortName(x.name))}</option>`).join('');
    const kpi = (l, v, sub) => `<div class="card kpi"><div class="l">${l}</div><div class="v">${v}</div><div class="d muted">${sub || ''}</div></div>`;
    const bars = C.INDICATORS.map((i) => { const sc = r.scores[i.id], v = r.vals[i.id]; return { label: `${i.id} ${i.label}`, title: `${i.label}: значение ${v === null ? '—' : fmtN(v)}, балл ${sc === null ? 'н/д' : fmt1(sc)}`, vals: [sc === null ? 0 : sc] }; });
    const bySrc = {}; C.UCOLS.forEach((c) => { if (S.ucolOn[c.id] && u.parts[c.id] !== null) bySrc[c.src] = (bySrc[c.src] || 0) + u.parts[c.id]; });
    const dn = Object.entries(bySrc).map(([label, value]) => ({ label, value })).sort((a, b) => b.value - a.value).map((x, i) => Object.assign(x, { color: Ch.PALETTE[i % Ch.PALETTE.length] }));
    const dnTotal = dn.reduce((a, b) => a + b.value, 0);
    let trend = '<div class="muted">Для динамики нужно минимум два среза.</div>';
    if (all.length > 1) {
      const RA = C.rateAll(all, S), labels = all.map((s) => s.label);
      const sc = RA.map((R) => { const x = R.rows.find((y) => y.key === state.vuz); return x ? x.total : null; });
      const er = all.map((s) => { const x = C.unified(s, C.includedKeys(s, S), S).rows.find((y) => y.key === state.vuz); return x ? x.total : null; });
      trend = `<div class="grid g2"><div><div class="sub">Итоговый балл (единая шкала)</div>${Ch.line([{ name: 'Балл', color: '#e15759', values: sc }], labels, { fmt: fmt1, height: 190 })}</div>
        <div><div class="sub">Общее количество ошибок</div>${Ch.line([{ name: 'Ошибок', color: '#4e79a7', values: er }], labels, { height: 190 })}</div></div>`;
    }
    const uq = all.map((s) => C.unified(s, [state.vuz], S).rows[0]);
    const metricsTbl = all.length ? `<table><thead><tr><th class="name static">Источник</th>${all.map((s) => `<th class="static">${esc(s.label)}</th>`).join('')}</tr></thead><tbody>${C.UCOLS.map((c) => `<tr><td class="name">${esc(c.label)}</td>${uq.map((q, qi) => { const v = q && all[qi].m[state.vuz] ? q.parts[c.id] : null; return `<td>${v === null || v === undefined ? '—' : fmtN(v)}</td>`; }).join('')}</tr>`).join('')}</tbody></table>` : '';
    return `<div class="toolbar"><button class="btn ghost sm" id="vPrev">← предыдущий</button>
        <div class="vsearch"><input type="text" id="vq" placeholder="🔍 Поиск ОВПО по названию…" autocomplete="off"><div id="vsug" class="suggest" hidden></div></div>
        <select id="vuzSel" style="max-width:520px">${opts}</select><button class="btn ghost sm" id="vNext">следующий →</button></div>
      <div class="card" style="margin-bottom:16px"><h2 style="font-size:17px">${esc(r.name)}</h2><div class="sub">Срез ${esc(D.snap.label)} · место ${r.rank} из ${n} по итоговому баллу (1 — худшее) ${dl ? '· Δ к ' + esc(D.cmp.label) + ': ' + deltaHtml(dl.d) : ''}</div></div>
      <div class="grid g4">
        ${kpi('Итоговый балл', `<span class="pill"${heatStyle(r.total)}>${fmt1(r.total)}</span>`, `место ${r.rank} из ${n}`)}
        ${kpi('Корректность', fmt1(r.K), `место ${r.rankK} из ${n}`)}${kpi('Заполненность', fmt1(r.Z), `место ${r.rankZ} из ${n}`)}${kpi('Исполнение', fmt1(r.I), `место ${r.rankI} из ${n}`)}
        ${kpi('Всего ошибок', fmtN(u.total), `место ${u.rank} · ${fmt1(u.share)}% от общего`)}
        ${kpi('ОСМС на 100 записей', u.osmsRate === null ? '—' : fmt1(u.osmsRate), u.osmsRec ? fmtN(u.osmsRec) + ' записей' : '')}
      </div>
      <div class="grid g2" style="margin-top:16px">
        <div class="card"><h2>Баллы по показателям</h2><div class="sub">Чем длиннее полоса, тем хуже (наведите курсор — фактическое значение)</div>${Ch.hbar(bars, { colors: ['#e15759'], max: 100, labelW: 330, barW: 260, rowH: 22 })}</div>
        <div class="card"><h2>Структура ошибок ОВПО</h2><div class="donut-wrap">${Ch.donut(dn, { center: fmtN(dnTotal), sub: 'ошибок' })}${Ch.legend(dn, dnTotal)}</div></div>
      </div>
      <div class="card" style="margin-top:16px"><h2>Где больше всего ошибок — топ-12 полей</h2><div class="sub">Конкретные поля из всех листов (по убыванию). Полная расшифровка по каждому листу — ниже.</div>${topFieldsHtml(D, state.vuz)}</div>
      <div class="card" style="margin-top:16px"><h2>Динамика по срезам</h2>${trend}</div>
      <h2 style="margin:22px 2px 6px;font-size:16px">Подробно по листам и полям</h2>
      <div class="toolbar"><label class="chip"><input type="checkbox" id="hideZero" ${state.hideZero ? 'checked' : ''}> скрывать поля без ошибок</label><label class="chip"><input type="checkbox" id="detSort" ${state.detSort ? 'checked' : ''}> сортировать по убыванию</label>
        <span class="mini">Нажмите на раздел, чтобы перейти к нему. «Доля ОВПО» — какую часть всех ошибок по этому полю даёт данное ОВПО.</span></div>
      ${sectionsHtml(D, state.vuz)}
      <div class="card" style="margin-top:16px"><h2>Значения по источникам и срезам</h2><div class="tbl-wrap" style="max-height:none">${metricsTbl}</div></div>`;
  }

  /* ------------------------------------------------------------ ДАННЫЕ И НАСТРОЙКИ */
  function viewData() {
    const all = sortedSnaps(), S = state.settings;
    const D = derive();
    const snapRows = all.map((s) => {
      const warn = s.diag && s.diag.warnings.length;
      return `<tr><td><input type="text" data-lbl="${s.id}" value="${esc(s.label)}"></td><td><input type="date" data-date="${s.id}" value="${s.date || ''}"></td>
        <td class="mini">${esc(s.fileName)}</td><td>${fmtN(Object.keys(s.names).length)}</td>
        <td>${warn ? `<span class="tag bad">${warn} предупр.</span>` : '<span class="tag good">ок</span>'}${s.missing && s.missing.length ? ' <span class="tag flat">нет листов: ' + s.missing.length + '</span>' : ''}${!s.tables ? ' <span class="tag warn">нет детализации по полям</span>' : ''}
          <details><summary>листы</summary><div class="mini">${(s.diag ? s.diag.sheets : []).map((x) => `${x.found ? '✓' : '✗'} ${esc(x.label)} — ${x.rows} строк`).join('<br>')}${s.diag && s.diag.warnings.length ? '<br><b>' + s.diag.warnings.map(esc).join('<br>') + '</b>' : ''}</div></details></td>
        <td><button class="btn danger sm" data-del="${s.id}">Удалить</button></td></tr>`;
    }).join('');
    let activeInfo = 'Сейчас в расчёт входят все ОВПО из файла.';
    if (S.activeNames && S.activeNames.length && D) {
      const miss = S.activeNames.filter((nm) => { const k = C.key(nm); return !(k in D.snap.names) && !(D.snap.alias && D.snap.alias[k] && D.snap.alias[k] in D.snap.names); });
      activeInfo = `Загружен список действующих ОВПО: <b>${S.activeNames.length}</b>. Найдено в текущем срезе: <b>${S.activeNames.length - miss.length}</b>.` + (miss.length ? `<details open><summary>Не найдены в срезе (${miss.length})</summary><div class="mini">${miss.map(esc).join('<br>')}</div></details>` : '');
    }
    const crit = ['К', 'З', 'И'].map((c) => `<label>${C.CRIT[c]}<input type="number" min="0" step="0.5" style="width:90px" data-cw="${c}" value="${S.critWeights[c]}"></label>`).join('');
    const indW = C.INDICATORS.map((i) => `<tr><td class="name" style="position:static;min-width:0">${i.id} · ${esc(i.label)}</td><td style="text-align:left">${C.CRIT[i.crit]}</td><td><input type="number" min="0" step="0.5" style="width:80px" data-iw="${i.id}" value="${S.indWeights[i.id]}"></td></tr>`).join('');
    const methRows = C.INDICATORS.map((i) => `<tr><td class="name" style="position:static;min-width:0">${i.id}</td><td style="text-align:left">${esc(i.label)}</td><td style="text-align:left">${C.CRIT[i.crit]}</td><td style="text-align:left">${i.kind === 'lin' ? 'сам процент (0–100)' : 'логарифмическая шкала'}</td></tr>`).join('');
    const methodology = `<div class="card" style="margin-top:16px"><h2>Методика расчёта</h2>
      <details><summary>Как считаются баллы и рейтинги</summary><div class="mini" style="line-height:1.65;margin-top:8px">
      <p><b>Балл показателя (0–100).</b> Для счётных показателей: LN(1 + значение) / LN(1 + максимум среди ОВПО) × 100. Логарифм нужен, чтобы несколько ОВПО-выбросов (например, 150 тыс. ошибок оцифровки) не «обнуляли» различия между остальными. Для процентных показателей балл равен самому проценту. 0 — проблем нет, 100 — худший.</p>
      <p><b>Критерий</b> — взвешенное среднее баллов его показателей (если показатель к ОВПО неприменим — он исключается, веса перенормируются). <b>Итоговый балл</b> — взвешенное среднее трёх критериев. Ранг 1 — худшее ОВПО; одинаковые баллы получают одинаковый ранг.</p>
      <p><b>Критерии:</b> корректность — данные есть, но неверные; заполненность — данных нет; исполнение — заявки по госуслугам с истёкшим сроком обработки.</p>
      <p><b>Сводный рейтинг ошибок</b> — простая сумма ошибок по выбранным источникам; «% от общего» — доля ОВПО в сумме по всем ОВПО в расчёте. Единицы в источниках разные (записи, поля, дипломы, отчёты), поэтому это объём проблем.</p>
      <p><b>Динамика</b> считается на единой шкале всех срезов, чтобы изменения баллов были сопоставимы. Срезы в старом формате (до 30.10) устроены иначе — для них сравнение ориентировочное.</p>
      <p><b>Ограничение:</b> в данных нет численности студентов и ППС, поэтому показатели абсолютные — крупные ВУЗы в невыгодном положении.</p></div>
      <div class="tbl-wrap" style="max-height:360px;margin-top:8px"><table><thead><tr><th class="name" style="position:static">Код</th><th style="text-align:left">Показатель</th><th style="text-align:left">Критерий</th><th style="text-align:left">Шкала</th></tr></thead><tbody>${methRows}</tbody></table></div></details></div>`;
    return `
    <div class="grid g2" style="margin-top:16px">
      <div class="card"><h2>Какие ОВПО учитывать</h2><div class="sub">Например, только действующие организации</div>
        <div class="note ${S.activeNames ? '' : 'warn'}">${activeInfo}</div>
        <div class="form-row"><button class="btn ghost" id="pickActive">Загрузить список ОВПО (xlsx)</button><button class="btn ghost" id="clearActive">Сбросить список</button></div>
        <div class="mini">В списке названия должны быть в первом столбце (первая строка — заголовок «ОВПО»). Файл можно также перетащить в зону загрузки выше — он распознается автоматически.</div>
        <div class="form-row"><label>Исключить ОВПО, в названии которых есть (через запятую)<input type="text" id="excl" style="width:340px" value="${esc((S.exclude || []).join(', '))}"></label></div></div>
      <div class="card"><h2>Веса критериев</h2><div class="sub">Итоговый балл — взвешенное среднее трёх критериев</div>
        <div class="form-row">${crit}</div>
        <div class="form-row"><label>Порог «существенного» изменения балла<input type="number" min="0" step="0.5" style="width:90px" id="trendT" value="${S.trend}"></label></div>
        <details><summary>Веса отдельных показателей</summary><div class="tbl-wrap" style="max-height:340px;margin-top:6px"><table><thead><tr><th class="name" style="position:static">Показатель</th><th style="text-align:left">Критерий</th><th>Вес</th></tr></thead><tbody>${indW}</tbody></table></div></details>
        <div class="form-row"><button class="btn ghost" id="resetSettings">Вернуть настройки по умолчанию</button></div></div>
    </div>
    ${methodology}`;
  }

   //  <div class="card"><h2>Загрузка срезов</h2><div class="sub">Файлы разбираются в браузере и никуда не отправляются. Можно загрузить несколько срезов сразу.</div>
   //    <div class="drop" id="dropMain"><div class="big">Перетащите Excel-файлы сюда</div><div class="muted">или <a href="#" id="pickLink">выберите файлы</a> (.xlsx). Здесь же можно загрузить сохранённый <code>bundle.js</code>.</div></div>
   //    ${all.length ? `<div class="tbl-wrap" style="max-height:none;margin-top:14px"><table class="snap-table"><thead><tr><th>Название среза</th><th>Дата</th><th>Файл</th><th>ОВПО</th><th>Проверка разбора</th><th></th></tr></thead><tbody>${snapRows}</tbody></table></div>` : ''}
   //  </div>
   // <div class="card" style="margin-top:16px"><h2>Публикация для сотрудников (GitHub Pages)</h2>
   //    <div class="sub">${state.published ? '<span class="tag pub">опубликованные данные загружены из data/bundle.js</span>' : 'Сейчас данные хранятся только в вашем браузере.'}</div>
   //    <ol class="mini" style="line-height:1.7"><li>Нажмите «Скачать bundle.js» — в файл попадут все срезы и настройки (только итоговые числа по ОВПО, без ИИН и ID).</li>
   //    <li>Положите файл в папку <code>data/</code> сайта (замените существующий) и отправьте изменения на GitHub.</li>
   //    <li>Сотрудники откроют ссылку и сразу увидят все дашборды — загружать файлы им не нужно.</li></ol>
   //    <div class="note warn">Сайт на GitHub Pages общедоступен по ссылке. Данные по ОВПО — служебные: используйте закрытый репозиторий с закрытым доступом к Pages или внутренний хостинг, если публиковать их открыто нельзя.</div>
   //    <div class="form-row"><button class="btn" id="exportBundle">Скачать bundle.js</button><button class="btn danger" id="wipe">Очистить данные в браузере</button></div></div>

  /* ------------------------------------------------------------ загрузка файлов */
  function parseBundleText(txt) {
    let t = txt.trim(); const i = t.indexOf('{'); const j = t.lastIndexOf('}');
    if (i < 0 || j < 0) throw new Error('не найден JSON');
    return JSON.parse(t.slice(i, j + 1));
  }
  function readNamesList(buf) {
    const wb = XLSX.read(buf, { type: 'array' });
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: null });
    return rows.map((r) => r[0]).filter((v) => typeof v === 'string' && v.trim()).map((v) => C.norm(v)).filter((v) => !/^(овпо|вуз|организация|наименование)$/i.test(v));
  }
  async function handleFiles(files) {
    const list = Array.from(files);   // копия: поле выбора файла очищается сразу после события
    loading(true, 'Обработка файлов…'); await tick();
    let added = 0;
    for (const f of list) {
      try {
        if (/\.(js|json)$/i.test(f.name)) {
          const b = parseBundleText(await f.text());
          const have = new Set(state.snaps.map((s) => s.id));
          (b.snapshots || []).forEach((s) => { if (!have.has(s.id)) { state.snaps.push(s); added++; } });
          if (b.settings) state.settings = mergeSettings(b.settings);
          toast(`Импортировано срезов: ${added}`);
        } else if (/\.xlsx?$/i.test(f.name)) {
          loading(true, `Разбираю «${f.name}»…`); await tick();
          const buf = new Uint8Array(await f.arrayBuffer());
          const snap = C.parseSnapshot(XLSX, buf, f.name);
          if (snap.diag.sheets.every((x) => !x.found)) {
            const names = readNamesList(buf);
            if (names.length) { state.settings.activeNames = names; toast(`Загружен список ОВПО: ${names.length}`); continue; }
          }
          state.snaps.push(snap); added++;
          if (snap.diag.warnings.length) toast(`«${f.name}»: ${snap.diag.warnings.length} предупр. — см. «Данные и настройки»`);
        } else toast(`Неподдерживаемый файл: ${f.name}`);
      } catch (e) { console.error(e); toast(`Ошибка при чтении «${f.name}»: ${e.message}`); }
    }
    if (added) { const all = sortedSnaps(); state.cur = all[all.length - 1].id; }
    invalidate(); save(); loading(false); render();
  }

  /* ------------------------------------------------------------ рендер и события */
  function render() {
    renderTop();
    const view = $('#view'), D = derive();
    if (!D && state.tab !== 'data') { view.innerHTML = emptyHtml(); bindDrop(); return; }
    const t = state.tab;
    view.innerHTML = t === 'overview' ? viewOverview(D) : t === 'rating' ? viewRating(D) : t === 'unified' ? viewUnified(D) : t === 'dynamics' ? viewDynamics(D) : t === 'vuz' ? viewVuz(D) : viewData();
    if (t === 'rating') fillRating(D);
    if (t === 'unified') fillUnified(D);
    bindDrop();
  }
  function bindDrop() {
    $$('.drop').forEach((el) => {
      ['dragenter', 'dragover'].forEach((ev) => el.addEventListener(ev, (e) => { e.preventDefault(); el.classList.add('over'); }));
      ['dragleave', 'drop'].forEach((ev) => el.addEventListener(ev, (e) => { e.preventDefault(); el.classList.remove('over'); }));
      el.addEventListener('drop', (e) => handleFiles(e.dataTransfer.files));
    });
  }
  function changeSettings() { invalidate(); save(); render(); }

  document.addEventListener('click', (e) => {
    const t = e.target.closest('[data-tab],[data-vuz],[data-pick],[data-goto],[data-sort],[data-ucol],[data-del],#pickLink,#pickActive,#clearActive,#exportBundle,#wipe,#resetSettings,#exp-rating,#exp-unified,#vPrev,#vNext,#themeBtn,#uploadBtn');
    if (!e.target.closest('.vsearch')) { const sg = $('#vsug'); if (sg) sg.hidden = true; }
    if (!t) return;
    const D = derive();
    if (t.dataset.tab) { state.tab = t.dataset.tab; render(); window.scrollTo(0, 0); }
    else if (t.dataset.pick) { state.vuz = t.dataset.pick; state.tab = 'vuz'; render(); window.scrollTo(0, 0); }
    else if (t.dataset.goto) { const el = document.getElementById(t.dataset.goto); if (el) el.scrollIntoView({ behavior: 'smooth', block: 'start' }); }
    else if (t.dataset.vuz) { state.vuz = t.dataset.vuz; state.tab = 'vuz'; render(); window.scrollTo(0, 0); }
    else if (t.dataset.sort) {
      const [tb, col] = t.dataset.sort.split(':'), st = state.sort[tb];
      if (st.col === col) st.dir *= -1; else { st.col = col; st.dir = col === 'name' ? 1 : -1; }
      if (D) (tb === 'rating' ? fillRating : fillUnified)(D);
    }
    else if (t.dataset.ucol) { state.settings.ucolOn[t.dataset.ucol] = state.settings.ucolOn[t.dataset.ucol] ? 0 : 1; changeSettings(); }
    else if (t.dataset.del) { state.snaps = state.snaps.filter((s) => s.id !== t.dataset.del); const all = sortedSnaps(); if (state.cur === t.dataset.del) state.cur = all.length ? all[all.length - 1].id : null; changeSettings(); }
    else if (t.id === 'pickLink' || t.id === 'uploadBtn') { e.preventDefault(); $('#fileInput').click(); }
    else if (t.id === 'pickActive') { $('#activeInput').click(); }
    else if (t.id === 'clearActive') { state.settings.activeNames = null; changeSettings(); }
    else if (t.id === 'exportBundle') {
      const b = { version: 1, generated: new Date().toISOString(), settings: state.settings, snapshots: state.snaps };
      download('bundle.js', '/* Данные дашборда ОВПО: только агрегаты по ОВПО, без ИИН и ID. Положите файл в папку data/ */\nwindow.VUZ_BUNDLE = ' + JSON.stringify(b) + ';\n', 'text/javascript');
      toast('bundle.js скачан — положите его в папку data/ и отправьте на GitHub');
    }
    else if (t.id === 'wipe') { if (confirm('Удалить все загруженные срезы и настройки из браузера?')) { try { localStorage.removeItem(LS); } catch (x) { /* ignore */ } state.snaps = []; state.cur = null; state.settings = C.defaultSettings(); changeSettings(); } }
    else if (t.id === 'resetSettings') { state.settings = Object.assign(C.defaultSettings(), { activeNames: state.settings.activeNames }); changeSettings(); }
    else if (t.id === 'exp-rating' && D) exportRating(D);
    else if (t.id === 'exp-unified' && D) exportUnified(D);
    else if ((t.id === 'vPrev' || t.id === 'vNext') && D) {
      const rows = D.rated.rows, i = rows.findIndex((x) => x.key === state.vuz), j = Math.max(0, Math.min(rows.length - 1, i + (t.id === 'vNext' ? 1 : -1)));
      state.vuz = rows[j].key; render();
    }
    else if (t.id === 'themeBtn') {
      const dark = document.documentElement.dataset.theme !== 'dark';
      document.documentElement.dataset.theme = dark ? 'dark' : ''; try { localStorage.setItem(LS_THEME, dark ? 'dark' : 'light'); } catch (x) { /* ignore */ }
    }
  });
  document.addEventListener('input', (e) => {
    const t = e.target;
    if (t.id === 'q-rating') { state.search.rating = t.value; const D = derive(); if (D) fillRating(D); }
    else if (t.id === 'q-unified') { state.search.unified = t.value; const D = derive(); if (D) fillUnified(D); }
    else if (t.id === 'vq') suggest(t.value);
  });
  document.addEventListener('change', (e) => {
    const t = e.target;
    if (t.id === 'snapSel') { state.cur = t.value; invalidate(); save(); render(); }
    else if (t.id === 'cmpSel') { state.cmp = t.value; invalidate(); render(); }
    else if (t.id === 'fileInput') { handleFiles(Array.from(t.files)); t.value = ''; }
    else if (t.id === 'activeInput') {
      const f = t.files[0]; if (!f) return;
      f.arrayBuffer().then((b) => { const names = readNamesList(new Uint8Array(b)); state.settings.activeNames = names; toast(`Загружен список ОВПО: ${names.length}`); changeSettings(); }); t.value = '';
    }
    else if (t.id === 'noDig') { state.noDig = t.checked; render(); }
    else if (t.id === 'dynMetric') { state.dynMetric = t.value; render(); }
    else if (t.id === 'vuzSel') { state.vuz = t.value; render(); }
    else if (t.id === 'hideZero') { state.hideZero = t.checked; render(); }
    else if (t.id === 'detSort') { state.detSort = t.checked; render(); }
    else if (t.id === 'excl') { state.settings.exclude = t.value.split(',').map((x) => x.trim()).filter(Boolean); changeSettings(); }
    else if (t.id === 'trendT') { state.settings.trend = Math.max(0, Number(t.value) || 0); changeSettings(); }
    else if (t.dataset.cw) { state.settings.critWeights[t.dataset.cw] = Math.max(0, Number(t.value) || 0); changeSettings(); }
    else if (t.dataset.iw) { state.settings.indWeights[t.dataset.iw] = Math.max(0, Number(t.value) || 0); changeSettings(); }
    else if (t.dataset.lbl) { const s = state.snaps.find((x) => x.id === t.dataset.lbl); if (s) { s.label = t.value || s.label; changeSettings(); } }
    else if (t.dataset.date) { const s = state.snaps.find((x) => x.id === t.dataset.date); if (s) { s.date = t.value || null; changeSettings(); } }
  });
  function suggest(q) {
    const D = derive(), box = $('#vsug'); if (!D || !box) return;
    const toks = C.key(q).split(' ').filter(Boolean);
    if (!toks.length) { box.hidden = true; state.sug = { list: [], idx: 0 }; return; }
    const list = D.rated.rows.filter((r) => { const k = C.key(r.name); return toks.every((t) => k.includes(t)); });
    state.sug = { list: list.slice(0, 12).map((r) => r.key), idx: 0 };
    box.innerHTML = list.length
      ? list.slice(0, 12).map((r, i) => `<div class="sug-item${i === 0 ? ' on' : ''}" data-pick="${esc(r.key)}"><span>${esc(shortName(r.name))}</span><span class="muted">№${r.rank}</span></div>`).join('') + (list.length > 12 ? `<div class="sug-item muted">…и ещё ${list.length - 12}. Уточните запрос</div>` : '')
      : '<div class="sug-item muted">Ничего не найдено</div>';
    box.hidden = false;
  }
  document.addEventListener('keydown', (e) => {
    if (e.target.id !== 'vq') return;
    const box = $('#vsug'), n = state.sug.list.length;
    if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
      e.preventDefault(); if (!n) return;
      state.sug.idx = (state.sug.idx + (e.key === 'ArrowDown' ? 1 : -1) + n) % n;
      $$('.sug-item', box).forEach((el, i) => el.classList.toggle('on', i === state.sug.idx));
    } else if (e.key === 'Enter') {
      e.preventDefault(); if (n) { state.vuz = state.sug.list[state.sug.idx]; render(); window.scrollTo(0, 0); }
    } else if (e.key === 'Escape') box.hidden = true;
  });
  // перетаскивание файла в любое место окна
  window.addEventListener('dragover', (e) => e.preventDefault());
  window.addEventListener('drop', (e) => { if (!e.target.closest('.drop')) { e.preventDefault(); if (e.dataTransfer.files.length) handleFiles(e.dataTransfer.files); } });

  /* ------------------------------------------------------------ старт */
  load(); render();
  window.__vuzdash = { state, derive, handleFiles, shortName };   // для отладки и тестов
})();
