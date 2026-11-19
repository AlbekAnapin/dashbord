/* =====================================================================
   core.js — разбор Excel-среза и расчёт рейтингов (без привязки к DOM).
   Работает и в браузере (window.VuzCore), и в Node (module.exports) — для тестов.
   ===================================================================== */
(function (root) {
  'use strict';

  /* ---------------------------------------------------------------- утилиты */
  const norm = (s) => String(s).trim().replace(/[«»“”]/g, '"').replace(/\s+/g, ' ');
  const key = (s) => norm(s).toLowerCase().replace(/ё/g, 'е').replace(/’/g, "'").replace(/["']/g, ' ').replace(/\s+/g, ' ').trim();
  const num = (v) => { if (v === null || v === undefined || v === '') return 0; const n = Number(v); return isNaN(n) ? 0 : n; };
  const isStr = (v) => typeof v === 'string' && v.trim() !== '';
  const kk = (v) => (isStr(v) ? key(v) : '');
  const isMarker = (v) => { const k = kk(v); return k === 'овпо' || k === 'вуз'; };
  const isTotal = (v) => kk(v).startsWith('итого');
  const isName = (v) => isStr(v) && !isMarker(v) && !isTotal(v);

  // Названия-варианты одного ОВПО (подстрока дубля, подстрока основного названия, исключение)
  const ALIAS_RULES = [
    ['академия кнб рк', 'государственное учреждение академия комитета национальной безопасности', 'пограничная'],
    ['gazi', 'филиал университета гази в республике', null],
  ];

  /* ---------------------------------------------------------------- накопитель */
  function makeAcc() {
    const names = {};              // key -> отображаемое название
    const m = {};                  // key -> {метрика: число}
    const tables = {};             // id -> детализация по полям (см. модуль детализации ниже)
    const touch = (name) => {
      const k = key(name);
      if (!(k in names)) names[k] = norm(name);
      if (!m[k]) m[k] = {};
      return k;
    };
    const add = (name, metric, val) => {
      const k = touch(name);
      m[k][metric] = (m[k][metric] || 0) + val;
    };
    return { names, m, tables, add, touch };
  }

  function findHeaderRow(rows, from) {
    for (let r = from || 0; r < rows.length; r++) if (rows[r] && isMarker(rows[r][0])) return r;
    return -1;
  }
  const colWhere = (row, pred, from) => {
    for (let c = from || 0; c < (row ? row.length : 0); c++) if (row[c] !== null && row[c] !== undefined && pred(kk(row[c]), row[c])) return c;
    return -1;
  };

  /* ---------------------------------------------------------------- парсеры листов */
  const P = {};

  // ГБД ФЛ: три блока (Обучающиеся / Выпускники / ППС), столбец «Кол-во ошибок»
  P.gbd = (rows, acc) => {
    let metric = null, valCol = -1, n = 0;
    const titles = { 'обучающиеся': 'gbd_stud', 'выпускники': 'gbd_grad', 'ппс': 'gbd_pps' };
    for (const row of rows) {
      if (!row) continue;
      const a = row[0], ka = kk(a);
      if (!ka) continue;
      if (titles[ka]) { metric = titles[ka]; continue; }
      if (isMarker(a)) { valCol = colWhere(row, (k) => k.includes('кол-во ошибок')); continue; }
      if (isTotal(a) || !metric || valCol < 0) continue;
      acc.add(a, metric, num(row[valCol])); n++;
    }
    return n;
  };

  // ОСМС: категории (обучающиеся / выпускники / отчисленные) — столбцы «Подсчёт ошибок» и «Записей (студентов)»
  function osmsGroups(rows, hr) {
    const hdr = rows[hr] || [], sub = rows[hr + 1] || [], groups = [];
    hdr.forEach((v, c) => {
      const k = kk(v); if (!k) return;
      const cat = k.includes('обучающ') ? 'stud' : k.includes('выпуск') ? 'grad' : k.includes('отчисл') ? 'exp' : null;
      if (cat) groups.push({ cat, c });
    });
    groups.forEach((g, i) => {
      const end = i + 1 < groups.length ? groups[i + 1].c : hdr.length;
      g.errCol = -1; g.recCol = -1;
      for (let c = g.c; c < end; c++) {
        const sk = kk(sub[c]);
        if (sk.includes('подсч') || sk.includes('ошиб')) g.errCol = c;
        if (sk.includes('записей')) g.recCol = c;
      }
      if (g.errCol < 0) g.errCol = g.c;      // старый формат: одна строка заголовка, категория = столбец
    });
    return groups;
  }
  P.osms = (rows, acc) => {
    const hr = findHeaderRow(rows); if (hr < 0) return 0;
    let totalCol = colWhere(rows[hr], (k) => k.startsWith('всего'));
    if (totalCol < 0 && rows[hr + 1]) totalCol = colWhere(rows[hr + 1], (k) => k.startsWith('всего'));
    if (totalCol < 0) return 0;
    const groups = osmsGroups(rows, hr);
    let n = 0;
    for (let r = hr + 1; r < rows.length; r++) {
      const row = rows[r]; if (!row || !isName(row[0])) continue;
      acc.add(row[0], 'osms', num(row[totalCol]));
      let rec = 0;
      groups.forEach((g) => { acc.add(row[0], 'osms_' + g.cat, num(row[g.errCol])); if (g.recCol >= 0) rec += num(row[g.recCol]); });
      acc.add(row[0], 'osms_rec', rec); n++;
    }
    return n;
  };

  // Генерация номеров дипломов: «Всего за 3 года»
  P.gen = (rows, acc) => {
    let hr = findHeaderRow(rows); if (hr < 0) hr = 0;           // в старом формате шапка без «ОВПО» в первом столбце
    let col = colWhere(rows[hr], (k) => k.startsWith('всего'));
    if (col < 0) col = colWhere(rows[hr], () => true, 1);
    if (col < 0) return 0;
    let n = 0;
    for (let r = hr + 1; r < rows.length; r++) {
      const row = rows[r]; if (!row || !isName(row[0])) continue;
      acc.add(row[0], 'gen', num(row[col])); n++;
    }
    return n;
  };

  // Дубли ИИН: слева обучающиеся, справа ППС (группы по столбцам «ВУЗ»)
  P.dup = (rows, acc) => {
    const hr = findHeaderRow(rows); if (hr < 0) return 0;
    const head = rows[hr], groups = [];
    head.forEach((v, c) => { if (isMarker(v)) groups.push(c); });
    const specs = groups.map((c, i) => ({
      name: c, val: colWhere(head, (k) => k.startsWith('всего') || k === 'кол-во', c + 1), metric: i === 0 ? 'dup_stud' : 'dup_pps',
    })).slice(0, 2);
    let n = 0;
    for (let r = hr + 1; r < rows.length; r++) {
      const row = rows[r]; if (!row) continue;
      for (const g of specs) if (isName(row[g.name]) && g.val >= 0) { acc.add(row[g.name], g.metric, num(row[g.val])); n++; }
    }
    return n;
  };

  // Оцифровка дипломов: первый блок — «Общее количество ошибок», второй — «не оцифровано дипломов 2024-26»
  P.dig = (rows, acc) => {
    let n = 0, mode = 'first', valCol = -1;
    for (const row of rows) {
      if (!row) continue;
      const a = row[0], ka = kk(a);
      if (!ka) continue;
      if (ka.startsWith('количество')) { mode = ka.includes('не оцифрован') ? 'second' : 'skip'; valCol = -1; continue; }
      if (isMarker(a)) {
        valCol = mode === 'first' ? colWhere(row, (k) => k.includes('общее количество ошибок'))
          : colWhere(row, (k) => k.startsWith('итого'));
        continue;
      }
      if (isTotal(a) || valCol < 0 || mode === 'skip') continue;
      acc.add(a, mode === 'first' ? 'dig_full' : 'dig_recent', num(row[valCol])); n++;
    }
    return n;
  };

  // GPA: слева обучающиеся (одна колонка), справа выпускники (годы выпуска 2024..2026)
  P.gpa = (rows, acc) => {
    const hr = findHeaderRow(rows); if (hr < 0) return 0;
    const head = rows[hr], mc = [];
    head.forEach((v, c) => { if (isMarker(v)) mc.push(c); });
    const valL = colWhere(head, (k) => k.includes('gpa'), mc[0] + 1);
    const yearCols = [];
    if (mc[1] !== undefined) head.forEach((v, c) => { if (c > mc[1] && /^20\d\d/.test(String(v === null ? '' : v).trim())) yearCols.push(c); });
    let n = 0;
    for (let r = hr + 1; r < rows.length; r++) {
      const row = rows[r]; if (!row) continue;
      if (isName(row[mc[0]]) && valL >= 0) { acc.add(row[mc[0]], 'gpa', num(row[valL])); acc.add(row[mc[0]], 'gpa_stud', num(row[valL])); n++; }
      if (mc[1] !== undefined && isName(row[mc[1]])) { const g = yearCols.reduce((s, c) => s + num(row[c]), 0); acc.add(row[mc[1]], 'gpa', g); acc.add(row[mc[1]], 'gpa_grad', g); n++; }
    }
    return n;
  };

  // Полнота данных ППС
  P.pps = (rows, acc) => {
    const hr = findHeaderRow(rows); if (hr < 0) return 0;
    const head = rows[hr];
    const iin = colWhere(head, (k) => k.startsWith('иин отсутствует'));
    const fio = colWhere(head, (k) => k.startsWith('фио не на кириллице'));
    const dr = colWhere(head, (k) => k.startsWith('дата рождения не соответствует'));
    const all = colWhere(head, (k) => k === 'всего ошибок');
    if (iin < 0 || all < 0) return 0;
    let n = 0;
    for (let r = hr + 1; r < rows.length; r++) {
      const row = rows[r]; if (!row || !isName(row[0])) continue;
      let corr = 0, compl = 0;
      for (let c = iin; c < all; c++) { const v = num(row[c]); if (c === fio || c === dr) corr += v; else compl += v; }
      acc.add(row[0], 'pps_corr', corr); acc.add(row[0], 'pps_compl', compl); n++;
    }
    return n;
  };

  // Материально-техническая база: слева общежития с причиной, справа корпуса «Нет данных»
  P.mtb = (rows, acc) => {
    const hr = findHeaderRow(rows); if (hr < 0) return 0;
    const head = rows[hr], mc = [];
    head.forEach((v, c) => { if (isMarker(v)) mc.push(c); });
    const reasonCol = colWhere(head, (k) => k === 'причина');
    let n = 0;
    for (let r = hr + 1; r < rows.length; r++) {
      const row = rows[r]; if (!row) continue;
      if (isName(row[mc[0]]) && reasonCol >= 0) {
        const reason = String(row[reasonCol] === null ? '' : row[reasonCol]).toLowerCase();
        acc.add(row[mc[0]], 'mtb_price', (reason.includes('больше') || reason.includes('меньше')) ? 1 : 0);
        acc.add(row[mc[0]], 'mtb_unfilled', reason.includes('не заполнен') ? 1 : 0); n++;
      }
      if (mc[1] !== undefined && isName(row[mc[1]])) {
        let cnt = 0;
        for (let c = mc[1] + 1; c <= mc[1] + 4; c++) if (String(row[c] === null ? '' : row[c]).trim().toLowerCase() === 'нет данных') cnt++;
        acc.add(row[mc[1]], 'mtb_nodata', cnt); acc.add(row[mc[1]], 'mtb_right', 1); n++;
      }
    }
    return n;
  };

  // Государственные услуги: «Итого (кол-во срок обработки заявки истёк)»
  P.gos = (rows, acc, warn) => {
    const hr = findHeaderRow(rows); if (hr < 0) return 0;
    const hdr2 = (rows[hr + 1] || []).concat(rows[hr] || []);
    if (hdr2.some((v) => /принято|отклонено/i.test(String(v === null ? '' : v)))) { warn.push('Лист «Госуслуги»: формат «Принято / Отклонено» не поддерживается — нужен столбец «Итого (кол-во срок обработки заявки истек)»'); return 0; }
    const col = colWhere(rows[hr], (k) => k.startsWith('итого'));
    if (col < 0) return 0;
    let n = 0;
    for (let r = hr + 1; r < rows.length; r++) {
      const row = rows[r]; if (!row) continue;
      if (isTotal(row[0])) break;
      if (isMarker(row[0])) { warn.push('Лист «Госуслуги»: в файле несколько таблиц — использована первая'); break; }
      if (!isName(row[0])) continue;
      acc.add(row[0], 'gos', num(row[col])); n++;
    }
    return n;
  };

  // Административный отчёт: статус каждого отчёта (НЕТ ДАННЫХ = 1, ЧАСТИЧНО = 0,5)
  P.admin = (rows, acc) => {
    const hr = findHeaderRow(rows); if (hr < 0) return 0;
    const col = colWhere(rows[hr], (k) => k === 'статус');
    if (col < 0) return 0;
    let n = 0;
    for (let r = hr + 1; r < rows.length; r++) {
      const row = rows[r]; if (!row || !isName(row[0])) continue;
      const s = String(row[col] === null ? '' : row[col]).trim().toUpperCase();
      let w = 0;
      if (s.startsWith('НЕТ')) w = 1; else if (s.includes('ЧАСТИЧНО')) w = 0.5;
      acc.add(row[0], 'admin_n', 1); acc.add(row[0], 'admin_sum', w); acc.add(row[0], 'admin_bad', w > 0 ? 1 : 0); n++;
    }
    return n;
  };


  /* ---------------------------------------------------------------- детализация по полям (вкладка «ОВПО») */
  // Таблица: {id, sec, title, layout:'list'|'grid', cols:[{h, t:'n'|'i'|'s'|'st', total?}], rows:{ключ ОВПО: [[значения по столбцам], ...]}}
  // t: n — число ошибок, i — справочное число (записей, стоимость…), s — текст, st — код статуса (админ. отчёты)
  const lab = (v) => String(v === null || v === undefined ? '' : v).replace(/\s+/g, ' ').trim();
  const isIdHdr = (h) => /^(id\b|айди)/i.test(h);
  const D = {};
  function defTable(acc, id, sec, title, layout, cols, extra) {
    acc.tables[id] = Object.assign({ id, sec, title, layout, cols, rows: {} }, extra || {});
    return acc.tables[id];
  }
  function pushRow(acc, id, name, arr) {
    const t = acc.tables[id]; if (!t) return;
    const k = acc.touch(name);
    (t.rows[k] = t.rows[k] || []).push(arr);
  }
  const nums = (row, cols) => cols.map((c) => (c.t === 's' ? (row[c.idx] === null || row[c.idx] === undefined ? '' : String(row[c.idx])) : c.t === 'i' && (row[c.idx] === null || row[c.idx] === '') ? null : num(row[c.idx])));
  // столбцы по строке заголовка: от from до to (включительно), без ID-столбцов
  function colsOf(head, from, to, opt) {
    const cols = [];
    for (let c = from; c <= to && c < head.length; c++) {
      const h = lab(head[c]); if (!h || isIdHdr(h)) continue;
      const k = h.toLowerCase();
      const col = { idx: c, h, t: 'n' };
      if (opt && opt.total && opt.total.test(k)) col.total = true;
      if (opt && opt.info && opt.info.test(k)) col.t = 'i';
      if (opt && opt.text && opt.text.test(k)) col.t = 's';
      cols.push(col);
    }
    return cols;
  }

  D.gbd = (rows, acc) => {
    const titles = { 'обучающиеся': ['gbd_stud', 'ГБД ФЛ: обучающиеся'], 'выпускники': ['gbd_grad', 'ГБД ФЛ: выпускники'], 'ппс': ['gbd_pps', 'ГБД ФЛ: ППС'] };
    let cur = null;
    for (const row of rows) {
      if (!row) continue;
      const a = row[0], ka = kk(a); if (!ka) continue;
      if (titles[ka]) { cur = { id: titles[ka][0], title: titles[ka][1], cols: null }; continue; }
      if (isMarker(a) && cur) { cur.cols = colsOf(row, 1, row.length, { total: /кол-во ошибок/ }); defTable(acc, cur.id, 'gbd', cur.title, 'list', cur.cols); continue; }
      if (isTotal(a) || !cur || !cur.cols) continue;
      pushRow(acc, cur.id, a, nums(row, cur.cols));
    }
  };

  D.osms = (rows, acc) => {
    const hr = findHeaderRow(rows); if (hr < 0) return;
    const groups = osmsGroups(rows, hr);
    let totalCol = colWhere(rows[hr], (k) => k.startsWith('всего'));
    if (totalCol < 0 && rows[hr + 1]) totalCol = colWhere(rows[hr + 1], (k) => k.startsWith('всего'));
    const order = ['stud', 'exp', 'grad'], names = { stud: 'Обучающиеся (назначение)', exp: 'Отчисленные (снятие)', grad: 'Выпускники (снятие)' };
    defTable(acc, 'osms_cat', 'osms', 'ОСМС: ошибки по категориям', 'grid', [{ h: 'Категория', t: 's' }, { h: 'Записей (студентов)', t: 'i' }, { h: 'Ошибок', t: 'n', total: true }]);
    for (let r = hr + 1; r < rows.length; r++) {
      const row = rows[r]; if (!row || !isName(row[0])) continue;
      order.forEach((cat) => {
        const g = groups.find((x) => x.cat === cat); if (!g) return;
        pushRow(acc, 'osms_cat', row[0], [names[cat], g.recCol >= 0 ? num(row[g.recCol]) : null, num(row[g.errCol])]);
      });
    }
  };
  const osmsSub = (id, title) => (rows, acc) => {
    const hr = findHeaderRow(rows); if (hr < 0) return;
    const head = rows[hr], totalCol = colWhere(head, (k) => k.startsWith('всего'));
    if (totalCol < 0) return;
    const cols = colsOf(head, 1, totalCol, { total: /^всего/, info: /^записей/ });
    defTable(acc, id, 'osms', title, 'list', cols);
    let after = false;
    for (let r = hr + 1; r < rows.length; r++) {
      const row = rows[r]; if (!row) continue;
      if (isTotal(row[0])) { after = true; continue; }          // строки под первой строкой «ИТОГО» — отдельный блок
      if (!isName(row[0])) continue;
      if (!after) { pushRow(acc, id, row[0], nums(row, cols)); continue; }
      const id2 = id + '_2';
      if (!acc.tables[id2]) defTable(acc, id2, 'osms', title + ' — дополнительный блок', 'list', cols,
        { note: 'Это отдельный блок строк под строкой «ИТОГО» на этом листе исходного файла. В итоговые показатели листа «ОСМС» он не входит.' });
      pushRow(acc, id2, row[0], nums(row, cols));
    }
  };
  D.osmsStud = osmsSub('osms_stud', 'ОСМС: обучающиеся — по полям');
  D.osmsExp = osmsSub('osms_exp', 'ОСМС: отчисленные — по полям');
  D.osmsGrad = osmsSub('osms_grad', 'ОСМС: выпускники — по полям');

  D.gen = (rows, acc) => {
    const hr = findHeaderRow(rows); if (hr < 0) return;
    const cols = colsOf(rows[hr], 1, rows[hr].length, { total: /^всего/ });
    defTable(acc, 'gen', 'gen', 'Диплом сгенерирован, но студент не выпущен — по годам', 'grid', cols);
    for (let r = hr + 1; r < rows.length; r++) { const row = rows[r]; if (row && isName(row[0])) pushRow(acc, 'gen', row[0], nums(row, cols)); }
  };

  D.dup = (rows, acc) => {
    const hr = findHeaderRow(rows); if (hr < 0) return;
    const head = rows[hr], marks = [];
    head.forEach((v, c) => { if (isMarker(v)) marks.push(c); });
    marks.slice(0, 2).forEach((c0, i) => {
      const end = i + 1 < marks.length ? marks[i + 1] - 1 : head.length;
      const cols = colsOf(head, c0 + 1, end, { total: /^всего/, text: /^с какими/, info: /^кол-во вузов/ });
      const id = i === 0 ? 'dup_stud' : 'dup_pps';
      defTable(acc, id, 'dup', i === 0 ? 'Дубли ИИН: обучающиеся' : 'Дубли ИИН: ППС', 'list', cols);
      for (let r = hr + 1; r < rows.length; r++) { const row = rows[r]; if (row && isName(row[c0])) pushRow(acc, id, row[c0], nums(row, cols)); }
    });
  };

  D.dig = (rows, acc) => {
    let mode = 'first', cols = null, id = null;
    for (const row of rows) {
      if (!row) continue;
      const a = row[0], ka = kk(a); if (!ka) continue;
      if (ka.startsWith('количество')) { mode = ka.includes('не оцифрован') ? 'second' : 'skip'; cols = null; continue; }
      if (isMarker(a)) {
        if (mode === 'first') {
          const tc = colWhere(row, (k) => k.includes('общее количество ошибок'));
          cols = colsOf(row, 1, tc < 0 ? row.length : tc, { total: /общее количество ошибок/, info: /^не оцифрован/ });
          id = 'dig_main'; defTable(acc, id, 'dig', 'Оцифровка дипломов: ошибки по полям', 'list', cols);
        } else if (mode === 'second') {
          const tc = colWhere(row, (k) => k.startsWith('итого'));
          cols = colsOf(row, 1, tc < 0 ? row.length : tc, { total: /^итого/ });
          id = 'dig_years'; defTable(acc, id, 'dig', 'Не оцифровано дипломов выпускников — по годам выпуска', 'grid', cols);
        } else cols = null;
        continue;
      }
      if (isTotal(a) || !cols) continue;
      pushRow(acc, id, a, nums(row, cols));
    }
  };

  D.gpa = (rows, acc) => {
    const hr = findHeaderRow(rows); if (hr < 0) return;
    const head = rows[hr], mc = [];
    head.forEach((v, c) => { if (isMarker(v)) mc.push(c); });
    const valL = colWhere(head, (k) => k.includes('gpa'), mc[0] + 1);
    const degL = colWhere(head, (k) => k.includes('степень') || k.includes('уровень'), mc[0] + 1);
    defTable(acc, 'gpa_stud', 'gpa', 'GPA не заполнен: обучающиеся', 'grid', [{ h: 'Академическая степень', t: 's' }, { h: 'Кол-во не заполнено GPA', t: 'n', total: true }]);
    let yc = [], lvl = -1;
    if (mc[1] !== undefined) {
      head.forEach((v, c) => { if (c > mc[1] && /^20\d\d/.test(lab(v))) yc.push({ c, h: lab(v) }); });
      lvl = colWhere(head, (k) => k.includes('уровень') || k.includes('степень'), mc[1] + 1);
      defTable(acc, 'gpa_grad', 'gpa', 'GPA не заполнен: выпускники — по годам выпуска', 'grid',
        [{ h: 'Уровень образования', t: 's' }].concat(yc.map((y) => ({ h: 'Выпуск ' + y.h, t: 'n' })), [{ h: 'Итого', t: 'n', total: true }]));
    }
    for (let r = hr + 1; r < rows.length; r++) {
      const row = rows[r]; if (!row) continue;
      if (isName(row[mc[0]]) && valL >= 0) pushRow(acc, 'gpa_stud', row[mc[0]], [degL >= 0 ? lab(row[degL]) : '', num(row[valL])]);
      if (mc[1] !== undefined && isName(row[mc[1]])) {
        const vs = yc.map((y) => num(row[y.c]));
        pushRow(acc, 'gpa_grad', row[mc[1]], [lvl >= 0 ? lab(row[lvl]) : ''].concat(vs, [vs.reduce((a, b) => a + b, 0)]));
      }
    }
  };

  D.pps = (rows, acc) => {
    const hr = findHeaderRow(rows); if (hr < 0) return;
    const head = rows[hr], all = colWhere(head, (k) => k === 'всего ошибок'); if (all < 0) return;
    const cols = colsOf(head, 1, all, { total: /^всего ошибок/ });
    defTable(acc, 'pps', 'pps', 'Полнота данных ППС: ошибки по полям', 'list', cols);
    for (let r = hr + 1; r < rows.length; r++) { const row = rows[r]; if (row && isName(row[0])) pushRow(acc, 'pps', row[0], nums(row, cols)); }
  };

  D.mtb = (rows, acc) => {
    const hr = findHeaderRow(rows); if (hr < 0) return;
    const head = rows[hr], mc = [];
    head.forEach((v, c) => { if (isMarker(v)) mc.push(c); });
    const left = colsOf(head, mc[0] + 1, (mc[1] !== undefined ? mc[1] : head.length) - 1, { info: /стоимость/, text: /наименование|причина/ });
    defTable(acc, 'mtb_dorm', 'mtb', 'Общежития: что не заполнено или указано некорректно', 'grid', left);
    let right = [];
    if (mc[1] !== undefined) { right = colsOf(head, mc[1] + 1, head.length, { text: /./ }); defTable(acc, 'mtb_bldg', 'mtb', 'Корпуса и сооружения: наличие данных', 'grid', right); }
    for (let r = hr + 1; r < rows.length; r++) {
      const row = rows[r]; if (!row) continue;
      if (isName(row[mc[0]])) pushRow(acc, 'mtb_dorm', row[mc[0]], nums(row, left));
      if (mc[1] !== undefined && isName(row[mc[1]])) pushRow(acc, 'mtb_bldg', row[mc[1]], nums(row, right));
    }
  };

  D.gos = (rows, acc) => {
    const hr = findHeaderRow(rows); if (hr < 0) return;
    const head = rows[hr], tc = colWhere(head, (k) => k.startsWith('итого')); if (tc < 0) return;
    const cols = colsOf(head, 1, tc, { total: /^итого/ });
    defTable(acc, 'gos', 'gos', 'Госуслуги: заявки с истёкшим сроком обработки — по услугам', 'list', cols);
    for (let r = hr + 1; r < rows.length; r++) {
      const row = rows[r]; if (!row) continue;
      if (isTotal(row[0]) || isMarker(row[0])) break;
      if (isName(row[0])) pushRow(acc, 'gos', row[0], nums(row, cols));
    }
  };

  D.admin = (rows, acc) => {
    const hr = findHeaderRow(rows); if (hr < 0) return;
    const head = rows[hr], sc = colWhere(head, (k) => k === 'статус'); if (sc < 0) return;
    const cName = colWhere(head, (k) => k.startsWith('название')), cTab = colWhere(head, (k) => k === 'таблицы'),
      cOk = colWhere(head, (k) => k.startsWith('таблиц с данными')), cNo = colWhere(head, (k) => k.startsWith('таблицы без'));
    const t = defTable(acc, 'admin', 'admin', 'Административные отчёты: заполненность', 'grid',
      [{ h: 'Отчёт', t: 's' }, { h: 'Таблиц', t: 'i' }, { h: 'Таблиц с данными', t: 'i' }, { h: 'Статус', t: 'st' }, { h: 'Таблиц без данных', t: 'i' }], { names: {}, stat: [] });
    for (let r = hr + 1; r < rows.length; r++) {
      const row = rows[r]; if (!row || !isName(row[0])) continue;
      const code = lab(row[1]), st = lab(row[sc]);
      if (cName >= 0 && !(code in t.names)) t.names[code] = lab(row[cName]);
      let si = t.stat.indexOf(st); if (si < 0) { t.stat.push(st); si = t.stat.length - 1; }
      pushRow(acc, 'admin', row[0], [code, num(row[cTab]), num(row[cOk]), si, cNo >= 0 && row[cNo] !== null ? num(row[cNo]) : 0]);
    }
  };

  const SHEETS = [
    { id: 'gbd', label: 'ГБД ФЛ', re: /^гбд/, fn: P.gbd, det: D.gbd },
    { id: 'osms', label: 'ОСМС', re: /^осмс$/, fn: P.osms, det: D.osms },
    { id: 'osms_stud', label: 'ОСМС (Обучающиеся)', re: /осмс.*обуч/, det: D.osmsStud, optional: true },
    { id: 'osms_exp', label: 'ОСМС (Отчисленные)', re: /осмс.*отчисл/, det: D.osmsExp, optional: true },
    { id: 'osms_grad', label: 'ОСМС (Выпускники)', re: /осмс.*выпуск/, det: D.osmsGrad, optional: true },
    { id: 'gen', label: 'Генерация номеров дипломов', re: /генерац/, fn: P.gen, det: D.gen },
    { id: 'dup', label: 'Дублирование записей по ИИН', re: /дублиров/, fn: P.dup, det: D.dup },
    { id: 'dig', label: 'Оцифровка дипломов', re: /оцифров/, fn: P.dig, det: D.dig },
    { id: 'gpa', label: 'GPA не заполнен', re: /gpa/, fn: P.gpa, det: D.gpa },
    { id: 'pps', label: 'Полнота данных ППС', re: /полнота.*ппс|ппс.*полнота/, fn: P.pps, det: D.pps },
    { id: 'mtb', label: 'Материально-техническая база', re: /материальн/, fn: P.mtb, det: D.mtb },
    { id: 'gos', label: 'Государственные услуги', re: /госуд?арств|госудаств/, fn: P.gos, det: D.gos },
    { id: 'admin', label: 'Административный отчёт', re: /административ/, fn: P.admin, det: D.admin },
  ];

  /* ---------------------------------------------------------------- разбор файла */
  function parseDateFromName(fileName) {
    const m = /(\d{1,2})[._-](\d{1,2})[._-](20\d{2})/.exec(fileName || '');
    if (!m) return { iso: null, label: (fileName || 'срез').replace(/\.xlsx?$/i, '') };
    const dd = m[1].padStart(2, '0'), mm = m[2].padStart(2, '0');
    let label = `${dd}.${mm}.${m[3]}`;
    if (/до[_ ]?обнов/i.test(fileName)) label += ' (до обновления)';
    return { iso: `${m[3]}-${mm}-${dd}`, label };
  }

  function buildAlias(names) {
    const alias = {}, keys = Object.keys(names);
    for (const [src, tgt, excl] of ALIAS_RULES) {
      const t = keys.filter((k) => k.includes(tgt) && !(excl && k.includes(excl)));
      if (t.length !== 1) continue;
      keys.forEach((k) => { if (k.includes(src) && k !== t[0]) alias[k] = t[0]; });
    }
    return alias;
  }

  // В исходном файле подписи ряда столбцов обрезаны (~34 символа). У обучающихся в ГБД ФЛ берём полные названия
  // из блока выпускников (те же столбцы в том же порядке); одинаковые названия в одной таблице различаем буквой столбца.
  function fixLabels(XLSX, tables) {
    const g = tables.gbd_grad, st = tables.gbd_stud;
    if (g && st) {
      for (let i = 0; i < Math.min(12, st.cols.length, g.cols.length); i++) {
        if (/^(у данного иин|данный иин (числится|имеет))/i.test(st.cols[i].h)) { st.cols[i].o = st.cols[i].h; st.cols[i].h = g.cols[i].h.replace(/^Данный иин/, 'Данный ИИН'); }
      }
    }
    for (const t of Object.values(tables)) t.cols.forEach((c) => { if (/^Данный иин/.test(c.h)) c.h = c.h.replace(/^Данный иин/, 'Данный ИИН'); if (/\(те(нг)?$/.test(c.h)) c.h = c.h.replace(/\(те(нг)?$/, '(тенге)'); });
    for (const t of Object.values(tables)) {
      const cnt = {};
      t.cols.forEach((c) => { cnt[c.h] = (cnt[c.h] || 0) + 1; });
      t.cols.forEach((c) => {
        if (cnt[c.h] > 1 && c.idx !== undefined) { c.o = c.o || c.h; c.h = `${c.h} (столбец ${XLSX.utils.encode_col(c.idx)})`; }
      });
    }
  }

  function parseSnapshot(XLSX, data, fileName) {
    const wb = XLSX.read(data, { type: 'array', cellDates: false });
    const acc = makeAcc();
    const diag = { sheets: [], warnings: [] }, missing = [];
    for (const spec of SHEETS) {
      const sn = wb.SheetNames.find((n) => spec.re.test(String(n).toLowerCase().trim()));
      if (!sn) {
        diag.sheets.push({ id: spec.id, label: spec.label, found: false, rows: 0, optional: !!spec.optional });
        if (!spec.optional) { diag.warnings.push(`Не найден лист «${spec.label}»`); missing.push(spec.id); }
        continue;
      }
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, defval: null, raw: true, blankrows: true });
      let n = 0;
      if (spec.fn) {
        try { n = spec.fn(rows, acc, diag.warnings); } catch (e) { diag.warnings.push(`Ошибка разбора листа «${sn}»: ${e.message}`); }
        if (n === 0) { diag.warnings.push(`Лист «${sn}»: не найдено ни одной строки данных (изменилась структура?)`); missing.push(spec.id); }
      }
      if (spec.det) { try { spec.det(rows, acc); } catch (e) { diag.warnings.push(`Детализация листа «${sn}»: ${e.message}`); } }
      diag.sheets.push({ id: spec.id, label: spec.label, sheetName: sn, found: true, rows: n || (spec.fn ? 0 : rows.filter((r) => r && isName(r[0])).length), optional: !!spec.optional });
    }
    // объединяем варианты названий
    const alias = buildAlias(acc.names);
    const names = {}, m = {};
    for (const k of Object.keys(acc.names)) {
      const t = alias[k] || k;
      if (!(t in names)) names[t] = acc.names[t] || acc.names[k];
      m[t] = m[t] || {};
      for (const [mk, v] of Object.entries(acc.m[k] || {})) m[t][mk] = (m[t][mk] || 0) + v;
    }
    fixLabels(XLSX, acc.tables);
    const outTables = {};
    for (const [id, t] of Object.entries(acc.tables)) {
      const rowsOut = {};
      for (const [k, arr] of Object.entries(t.rows)) { const tk = alias[k] || k; rowsOut[tk] = (rowsOut[tk] || []).concat(arr); }
      outTables[id] = Object.assign({}, t, { rows: rowsOut, cols: t.cols.map((c) => { const o = { h: c.h, t: c.t }; if (c.total) o.total = 1; if (c.o) o.o = c.o; return o; }) });
    }
    const d = parseDateFromName(fileName);
    return {
      id: 's' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6),
      fileName: fileName || '', label: d.label, date: d.iso, names, m, alias, diag, missing, tables: outTables, ver: 2,
      added: new Date().toISOString(),
    };
  }

  /* ---------------------------------------------------------------- показатели */
  const CRIT = { 'К': 'Корректность', 'З': 'Заполненность', 'И': 'Исполнение' };
  const INDICATORS = [
    { id: 'K1', sheet: 'gbd', crit: 'К', label: 'ГБД ФЛ: обучающиеся, ошибок', kind: 'log', get: (m) => m.gbd_stud || 0 },
    { id: 'K2', sheet: 'gbd', crit: 'К', label: 'ГБД ФЛ: выпускники, ошибок', kind: 'log', get: (m) => m.gbd_grad || 0 },
    { id: 'K3', sheet: 'gbd', crit: 'К', label: 'ГБД ФЛ: ППС, ошибок', kind: 'log', get: (m) => m.gbd_pps || 0 },
    { id: 'K4', sheet: 'dup', crit: 'К', label: 'Дубли ИИН: обучающиеся, записей', kind: 'log', get: (m) => m.dup_stud || 0 },
    { id: 'K5', sheet: 'dup', crit: 'К', label: 'Дубли ИИН: ППС, записей', kind: 'log', get: (m) => m.dup_pps || 0 },
    { id: 'K6', sheet: 'gen', crit: 'К', label: 'Дипломы сгенерированы, но не выпущены (3 года)', kind: 'log', get: (m) => m.gen || 0 },
    { id: 'K7', sheet: 'pps', crit: 'К', label: 'ППС: ФИО не кириллицей + дата рожд. ≠ ИИН', kind: 'log', get: (m) => m.pps_corr || 0 },
    { id: 'K8', sheet: 'mtb', crit: 'К', label: 'Общежития с недостоверной стоимостью, шт.', kind: 'log', get: (m) => m.mtb_price || 0 },
    { id: 'Z1', sheet: 'osms', crit: 'З', label: 'ОСМС: ошибок', kind: 'log', get: (m) => m.osms || 0 },
    { id: 'Z2', sheet: 'dig', crit: 'З', label: 'Оцифровка дипломов: незаполн. полей', kind: 'log', get: (m) => m.dig_full || 0 },
    { id: 'Z3', sheet: 'gpa', crit: 'З', label: 'GPA не заполнен (обуч. + выпуск. 2024–26)', kind: 'log', get: (m) => m.gpa || 0 },
    { id: 'Z4', sheet: 'pps', crit: 'З', label: 'ППС: незаполненные поля', kind: 'log', get: (m) => m.pps_compl || 0 },
    { id: 'Z5', sheet: 'admin', crit: 'З', label: 'Админ. отчёты: % недозаполнения', kind: 'lin', get: (m) => (m.admin_n ? (m.admin_sum / m.admin_n) * 100 : null) },
    { id: 'Z6', sheet: 'mtb', crit: 'З', label: 'МТБ: незаполн. общежития + корпуса «нет данных»', kind: 'log',
      get: (m) => ((m.mtb_right || 0) > 0 || (m.mtb_unfilled || 0) > 0 ? (m.mtb_unfilled || 0) + (m.mtb_nodata || 0) : null) },
    { id: 'Z7', sheet: 'dig', crit: 'З', label: 'Не оцифровано дипломов выпускников 2024–26', kind: 'log', get: (m) => m.dig_recent || 0 },
    { id: 'I1', sheet: 'gos', crit: 'И', label: 'Госуслуги: не обработано заявок (срок истёк)', kind: 'log', get: (m) => m.gos || 0 },
  ];

  // Столбцы «сводного рейтинга по ошибкам»
  const UCOLS = [
    { id: 'u_gbd_stud', sheet: 'gbd', label: 'ГБД ФЛ: обучающиеся', src: 'ГБД ФЛ', get: (m) => m.gbd_stud || 0, on: 1 },
    { id: 'u_gbd_grad', sheet: 'gbd', label: 'ГБД ФЛ: выпускники', src: 'ГБД ФЛ', get: (m) => m.gbd_grad || 0, on: 1 },
    { id: 'u_gbd_pps', sheet: 'gbd', label: 'ГБД ФЛ: ППС', src: 'ГБД ФЛ', get: (m) => m.gbd_pps || 0, on: 1 },
    { id: 'u_osms_stud', sheet: 'osms', group: 'osms', label: 'ОСМС: обучающиеся', src: 'ОСМС', get: (m) => m.osms_stud || 0, on: 1 },
    { id: 'u_osms_exp', sheet: 'osms', group: 'osms', label: 'ОСМС: отчисленные', src: 'ОСМС', get: (m) => m.osms_exp || 0, on: 1 },
    { id: 'u_osms_grad', sheet: 'osms', group: 'osms', label: 'ОСМС: выпускники', src: 'ОСМС', get: (m) => m.osms_grad || 0, on: 1 },
    { id: 'u_gen', sheet: 'gen', label: 'Дипломы сгенерированы, но не выпущены (3 года)', src: 'Генерация дипломов', get: (m) => m.gen || 0, on: 1 },
    { id: 'u_dup_stud', sheet: 'dup', label: 'Дубли ИИН: обучающиеся', src: 'Дубли ИИН', get: (m) => m.dup_stud || 0, on: 1 },
    { id: 'u_dup_pps', sheet: 'dup', label: 'Дубли ИИН: ППС', src: 'Дубли ИИН', get: (m) => m.dup_pps || 0, on: 1 },
    { id: 'u_dig_full', sheet: 'dig', label: 'Оцифровка дипломов (все поля)', src: 'Оцифровка', get: (m) => m.dig_full || 0, on: 1 },
    { id: 'u_gpa_stud', sheet: 'gpa', group: 'gpa', label: 'GPA не заполнен: обучающиеся', src: 'GPA', get: (m) => m.gpa_stud || 0, on: 1 },
    { id: 'u_gpa_grad', sheet: 'gpa', group: 'gpa', label: 'GPA не заполнен: выпускники 2024–26', src: 'GPA', get: (m) => m.gpa_grad || 0, on: 1 },
    { id: 'u_pps', sheet: 'pps', label: 'Полнота данных ППС (все ошибки)', src: 'ППС', get: (m) => (m.pps_compl || 0) + (m.pps_corr || 0), on: 1 },
    { id: 'u_admin', sheet: 'admin', label: 'Админ. отчёты: отчётов с неполными данными', src: 'Админ. отчёты', get: (m) => m.admin_bad || 0, on: 1 },
    { id: 'u_mtb', sheet: 'mtb', label: 'МТБ: общежития с неверной ценой + без данных + корпуса «нет данных»', src: 'МТБ',
      get: (m) => (m.mtb_price || 0) + (m.mtb_unfilled || 0) + (m.mtb_nodata || 0), on: 1 },
    { id: 'u_dig_recent', sheet: 'dig', label: 'Не оцифровано дипломов выпускников 2024–26', src: 'Оцифровка', get: (m) => m.dig_recent || 0, on: 1 },
    { id: 'u_gos', sheet: 'gos', label: 'Госуслуги: не обработано заявок (срок истёк)', src: 'Госуслуги', get: (m) => m.gos || 0, on: 1 },
  ];

  const defaultSettings = () => ({
    critWeights: { 'К': 1, 'З': 1, 'И': 1 },
    indWeights: Object.fromEntries(INDICATORS.map((i) => [i.id, 1])),
    ucolOn: Object.fromEntries(UCOLS.map((u) => [u.id, u.on])),
    exclude: ['военн'],
    activeNames: null,           // null = все ОВПО; иначе массив названий
    trend: 2,                    // порог «существенного» изменения балла
  });

  /* ---------------------------------------------------------------- фильтр ОВПО */
  function includedKeys(snap, settings) {
    const excl = (settings.exclude || []).map((x) => key(x)).filter(Boolean);
    let active = null;
    if (settings.activeNames && settings.activeNames.length) {
      active = new Set();
      settings.activeNames.forEach((n) => { const k = key(n); active.add(k); if (snap.alias && snap.alias[k]) active.add(snap.alias[k]); });
    }
    return Object.keys(snap.names).filter((k) => {
      if (excl.some((w) => k.includes(w))) return false;
      if (active && !active.has(k)) return false;
      return true;
    });
  }

  /* ---------------------------------------------------------------- расчёт баллов */
  // maxOverride — общие максимумы (для сравнения срезов по единой шкале)
  function rate(snap, keys, settings, maxOverride) {
    const S = settings || defaultSettings();
    const vals = {}, mx = {};
    for (const ind of INDICATORS) {
      vals[ind.id] = {};
      let mxv = -Infinity;
      for (const k of keys) {
        const v = (snap.missing || []).includes(ind.sheet) ? null : ind.get(snap.m[k] || {});
        vals[ind.id][k] = v;
        if (v !== null && v > mxv) mxv = v;
      }
      mx[ind.id] = maxOverride && maxOverride[ind.id] !== undefined ? maxOverride[ind.id] : (mxv === -Infinity ? 0 : mxv);
    }
    const rows = keys.map((k) => {
      const scores = {};
      for (const ind of INDICATORS) {
        const v = vals[ind.id][k];
        if (v === null) scores[ind.id] = null;
        else if (ind.kind === 'lin') scores[ind.id] = v;
        else scores[ind.id] = mx[ind.id] === 0 ? 0 : (Math.log1p(v) / Math.log1p(mx[ind.id])) * 100;
      }
      const crit = {};
      for (const c of ['К', 'З', 'И']) {
        let sw = 0, sv = 0;
        for (const ind of INDICATORS) {
          if (ind.crit !== c || scores[ind.id] === null) continue;
          const w = S.indWeights[ind.id] === undefined ? 1 : S.indWeights[ind.id];
          sw += w; sv += w * scores[ind.id];
        }
        crit[c] = sw > 0 ? sv / sw : 0;
      }
      const W = S.critWeights, ws = W['К'] + W['З'] + W['И'];
      const total = ws > 0 ? (W['К'] * crit['К'] + W['З'] * crit['З'] + W['И'] * crit['И']) / ws : 0;
      const row = { key: k, name: snap.names[k], vals: {}, scores, K: crit['К'], Z: crit['З'], I: crit['И'], total };
      for (const ind of INDICATORS) row.vals[ind.id] = vals[ind.id][k];
      return row;
    });
    rankBy(rows, 'total', 'rank');
    rankBy(rows, 'K', 'rankK'); rankBy(rows, 'Z', 'rankZ'); rankBy(rows, 'I', 'rankI');
    rows.sort((a, b) => b.total - a.total || a.name.localeCompare(b.name, 'ru'));
    return { rows, max: mx };
  }

  function rankBy(rows, field, out) {
    const sorted = rows.map((r) => r[field]).sort((a, b) => b - a);
    rows.forEach((r) => { r[out] = 1 + sorted.findIndex((v) => v <= r[field] + 1e-12 && v >= r[field] - 1e-12); });
  }

  /* ---------------------------------------------------------------- сводный рейтинг по ошибкам */
  function unified(snap, keys, settings) {
    const S = settings || defaultSettings();
    const has = (met) => Object.values(snap.m).some((x) => x[met] !== undefined);
    const splitOk = { osms: has('osms_stud') || has('osms_exp') || has('osms_grad'), gpa: has('gpa_stud') || has('gpa_grad') };
    const rows = keys.map((k) => {
      const m = snap.m[k] || {};
      const parts = {};
      let total = 0;
      for (const u of UCOLS) {
        const gone = (snap.missing || []).includes(u.sheet) || (u.group && !splitOk[u.group]);
        const v = gone ? null : u.get(m); parts[u.id] = v;
        if (S.ucolOn[u.id] && v !== null) total += v;
      }
      // срез, сохранённый старой версией сайта (без разбивки на категории): берём общий итог, чтобы сумма не потерялась
      for (const grp of ['osms', 'gpa']) {
        if (splitOk[grp] || (snap.missing || []).includes(grp)) continue;
        if (UCOLS.some((u) => u.group === grp && S.ucolOn[u.id])) total += m[grp] || 0;
      }
      const osmsRec = m.osms_rec || 0;
      return { key: k, name: snap.names[k], parts, total, osmsRate: osmsRec ? ((m.osms || 0) / osmsRec) * 100 : null, osmsRec };
    });
    const grand = rows.reduce((s, r) => s + r.total, 0);
    rows.sort((a, b) => b.total - a.total || a.name.localeCompare(b.name, 'ru'));
    let cum = 0;
    rows.forEach((r, i) => {
      r.rank = i && rows[i - 1].total === r.total ? rows[i - 1].rank : i + 1;
      r.share = grand ? (r.total / grand) * 100 : 0;
      cum += r.share; r.cum = cum;
    });
    return { rows, grand };
  }

  /* ---------------------------------------------------------------- сравнение срезов */
  // Общая шкала по всем срезам, затем баллы каждого среза на этой шкале
  function rateAll(snaps, settings) {
    const keysBySnap = snaps.map((s) => includedKeys(s, settings));
    const common = {};
    for (const ind of INDICATORS) {
      let mx = 0;
      snaps.forEach((s, i) => keysBySnap[i].forEach((k) => { const v = (s.missing || []).includes(ind.sheet) ? null : ind.get(s.m[k] || {}); if (v !== null && v > mx) mx = v; }));
      common[ind.id] = mx;
    }
    return snaps.map((s, i) => rate(s, keysBySnap[i], settings, common));
  }

  const api = {
    norm, key, ALIAS_RULES, SHEETS, INDICATORS, UCOLS, CRIT,
    parseSnapshot, parseDateFromName, includedKeys, rate, rateAll, unified, defaultSettings,
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.VuzCore = api;
})(typeof window !== 'undefined' ? window : globalThis);
