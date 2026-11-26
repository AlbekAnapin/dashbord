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
  const isMarker = (v) => { const k = kk(v); return k === 'овпо' || k === 'вуз' || k === 'университет'; };
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
    const ids = {};                // id таблицы -> ключ ОВПО -> [{c: колонка, r: строка, v: 'id1,id2,…'}]
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
    return { names, m, tables, ids, add, touch };
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
    const catOf = (k) => (k.includes('обучающ') ? 'stud' : k.includes('выпуск') ? 'grad' : k.includes('отчисл') ? 'exp' : null);
    // вариант 1 (с 23.11): заголовок в одну строку — «Назначение (обучающиеся) Подсчёт ошибок», «ВСЕГО Ошибок»
    let totalCol = -1;
    hdr.forEach((v, c) => {
      const k = kk(v); if (!k || c === 0) return;
      const isRec = k.includes('записей'), isErr = k.includes('подсч') || k.includes('ошибок');
      if (k.startsWith('всего')) { if (isErr) totalCol = c; return; }
      const cat = catOf(k);
      if (cat && (isRec || isErr)) {
        let g = groups.find((x) => x.cat === cat); if (!g) { g = { cat, c, errCol: -1, recCol: -1 }; groups.push(g); }
        if (isErr) g.errCol = c; else g.recCol = c;
      }
    });
    if (groups.length) { groups.totalCol = totalCol; return groups; }
    // вариант 2: заголовок в две строки (группа / «Записей», «Подсчёт ошибок») или в одну строку без подзаголовков (старый формат)
    hdr.forEach((v, c) => { const k = kk(v); if (!k) return; const cat = catOf(k); if (cat) groups.push({ cat, c }); });
    const totalC = hdr.findIndex((v) => kk(v).startsWith('всего'));      // «ВСЕГО» — граница последней категории
    groups.forEach((g, i) => {
      const end = i + 1 < groups.length ? groups[i + 1].c : (totalC > g.c ? totalC : hdr.length);
      g.errCol = -1; g.recCol = -1;
      for (let c = g.c; c < end; c++) {
        const sk = kk(sub[c]);
        if (sk.includes('подсч') || sk.includes('ошиб')) g.errCol = c;
        if (sk.includes('записей')) g.recCol = c;
      }
      if (g.errCol < 0) g.errCol = g.c;      // старый формат: одна строка заголовка, категория = столбец
    });
    groups.totalCol = -1;
    return groups;
  }
  P.osms = (rows, acc) => {
    const hr = findHeaderRow(rows); if (hr < 0) return 0;
    const groups = osmsGroups(rows, hr);
    let totalCol = groups.totalCol;
    if (totalCol === undefined || totalCol < 0) {
      totalCol = colWhere(rows[hr], (k) => k.startsWith('всего'));
      if (totalCol < 0 && rows[hr + 1]) totalCol = colWhere(rows[hr + 1], (k) => k.startsWith('всего'));
    }
    if (totalCol < 0) return 0;
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

  // Дубли ИИН. Новый формат (с 23.11): построчный список записей — слева обучающиеся (studentid), справа ППС (tutorid).
  // Старый формат: сводка по вузам с итогами «Всего…». Формат определяется по заголовку «studentid».
  function dupLayout(rows) {
    for (let r = 0; r < Math.min(rows.length, 6); r++) {
      const row = rows[r] || [];
      const sid = colWhere(row, (k) => k === 'studentid');
      if (sid >= 0) {
        const L = {
          hr: r, sid, name: colWhere(row, (k) => k.startsWith('наименование'), sid + 1), times: colWhere(row, (k) => k.startsWith('сколько раз')),
          list: colWhere(row, (k) => k.startsWith('вузы по иин')), type: colWhere(row, (k) => k === 'тип дубля'), reason: colWhere(row, (k) => k === 'причина'),
        };
        L.pname = colWhere(row, (k) => k === 'вуз' || k === 'овпо', (L.reason >= 0 ? L.reason : sid) + 1);
        L.tid = colWhere(row, (k) => k === 'tutorid', L.pname + 1);
        L.precs = colWhere(row, (k) => k.startsWith('записей в этом'), L.pname + 1);
        L.pall = colWhere(row, (k) => k.startsWith('записей всего'), L.pname + 1);
        L.pcnt = colWhere(row, (k) => k.startsWith('кол-во вузов'), L.pname + 1);
        L.plist = colWhere(row, (k) => k.startsWith('все вузы'), L.pname + 1);
        L.pinner = colWhere(row, (k) => k.startsWith('вузы с дублем'), L.pname + 1);
        L.ptype = colWhere(row, (k) => k === 'тип дубля', L.pname + 1);
        return L;
      }
    }
    return null;
  }
  P.dup = (rows, acc) => {
    const L = dupLayout(rows);
    if (L) {
      let n = 0;
      for (let r = L.hr + 1; r < rows.length; r++) {
        const row = rows[r]; if (!row) continue;
        if (row[L.sid] !== null && row[L.sid] !== undefined && isName(row[L.name])) { acc.add(row[L.name], 'dup_stud', 1); n++; }
        if (L.tid >= 0 && row[L.tid] !== null && row[L.tid] !== undefined && isName(row[L.pname])) { acc.add(row[L.pname], 'dup_pps', 1); n++; }
      }
      return n;
    }
    return P.dupLegacy(rows, acc);
  };
  P.dupLegacy = (rows, acc) => {
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

  // Оцифровка дипломов: первый блок — «Общее количество ошибок» + «Не оцифрованно кол-во дипломов 1991г» (считается как поле),
  // второй блок — «не оцифровано дипломов 2024-26». В новых файлах строка заголовков может повторяться дважды.
  P.dig = (rows, acc) => {
    let n = 0, mode = 'first', valCol = -1, x91 = -1;
    for (const row of rows) {
      if (!row) continue;
      const a = row[0], ka = kk(a);
      if (!ka) continue;
      if (ka.startsWith('количество')) { mode = ka.includes('не оцифрован') ? 'second' : 'skip'; valCol = -1; continue; }
      if (isMarker(a)) {
        valCol = mode === 'first' ? colWhere(row, (k) => k.includes('общее количество ошибок')) : colWhere(row, (k) => k.startsWith('итого'));
        x91 = mode === 'first' ? colWhere(row, (k) => k.startsWith('не оцифрован')) : -1;
        continue;
      }
      if (isTotal(a) || valCol < 0 || mode === 'skip') continue;
      acc.add(a, mode === 'first' ? 'dig_full' : 'dig_recent', num(row[valCol]) + (x91 >= 0 ? num(row[x91]) : 0)); n++;
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
  P.mtb = (rows, acc, warn) => {
    const hr = findHeaderRow(rows); if (hr < 0) return 0;
    const head = rows[hr], mc = [];
    head.forEach((v, c) => { if (isMarker(v)) mc.push(c); });
    const reasonCol = colWhere(head, (k) => k === 'причина');
    if (reasonCol < 0) { if (warn) warn.push('Лист «МТБ»: в файле реестр общежитий без колонки «Причина» (нет списка ошибок) — показатели МТБ не рассчитываются'); return 0; }
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
  const isIdHdr = (h) => /^(id\b|айди|studentid|tutorid)/i.test(h);
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

  // ID-столбцы листа (айди студентов / выпускников / ППС): значения хранятся отдельно от таблиц, в snapshot.ids
  const idLabel = (h) => lab(h).replace(/\s*\(.*$/, '').trim();
  function idColsOf(head) { const out = []; (head || []).forEach((v, c) => { const h = lab(v); if (h && isIdHdr(h)) out.push({ idx: c, h: idLabel(h) }); }); return out; }
  function parseIdList(v) { if (v === null || v === undefined || v === '') return []; return String(v).replace(/[{}\[\]]/g, ' ').split(/[\s,;]+/).filter(Boolean); }
  function grabIds(acc, tid, name, row, idc, rowLabel) {
    (idc || []).forEach((x) => {
      const list = parseIdList(row[x.idx]); if (!list.length) return;
      const k = acc.touch(name), t = (acc.ids[tid] = acc.ids[tid] || {});
      (t[k] = t[k] || []).push({ c: x.h, r: rowLabel || '', v: list.join(',') });
    });
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
    const cols = colsOf(rows[hr], 1, rows[hr].length, { total: /^всего/ }), idc = idColsOf(rows[hr]);
    defTable(acc, 'gen', 'gen', 'Диплом сгенерирован, но студент не выпущен — по годам', 'grid', cols);
    for (let r = hr + 1; r < rows.length; r++) { const row = rows[r]; if (row && isName(row[0])) { pushRow(acc, 'gen', row[0], nums(row, cols)); grabIds(acc, 'gen', row[0], row, idc, ''); } }
  };

  D.dupLegacy = (rows, acc) => {
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

  // Новый формат: собираем записи, таблицы строятся после разбора всех листов (нужны названия вузов для списка «с какими вузами»)
  D.dup = (rows, acc) => {
    const L = dupLayout(rows);
    if (!L) { D.dupLegacy(rows, acc); return; }
    const pend = { stud: [], pps: [], infoS: lab(rows[0] && rows[0][0]), infoP: lab(rows[0] && L.pname >= 0 ? rows[0][L.pname] : '') };
    for (let r = L.hr + 1; r < rows.length; r++) {
      const row = rows[r]; if (!row) continue;
      if (row[L.sid] !== null && row[L.sid] !== undefined && isName(row[L.name])) {
        pend.stud.push({ name: row[L.name], id: String(row[L.sid]).trim(), type: lab(row[L.type]) || 'Тип не указан', reason: L.reason >= 0 ? lab(row[L.reason]) || 'Причина не указана' : '', list: L.list >= 0 ? row[L.list] : '' });
      }
      if (L.tid >= 0 && row[L.tid] !== null && row[L.tid] !== undefined && isName(row[L.pname])) {
        pend.pps.push({ name: row[L.pname], id: String(row[L.tid]).trim(), type: L.ptype >= 0 ? lab(row[L.ptype]) || 'Тип не указан' : 'Тип не указан', list: L.plist >= 0 ? row[L.plist] : '' });
      }
    }
    acc._dupPending = pend;
  };
  function finalizeDup(acc) {
    const pend = acc._dupPending; if (!pend) return;
    const allNames = Object.values(acc.names).filter((n) => n.length > 4).sort((a, b) => b.length - a.length);
    const partnersOf = (text, own) => {
      let t = norm(text || ''); const found = [];
      for (const nm of allNames) { let i = t.indexOf(nm); while (i >= 0) { found.push(nm); t = t.slice(0, i) + '\u0001'.repeat(nm.length) + t.slice(i + nm.length); i = t.indexOf(nm); } }
      const ok = key(own); return found.filter((nm) => key(nm) !== ok);
    };
    const build = (tid, recs, groupBy, title, info, idCol, withPartners) => {
      const labels = []; recs.forEach((x) => { const g = x[groupBy]; if (g && !labels.includes(g)) labels.push(g); });
      if (!labels.length) return;
      const per = {};
      recs.forEach((x) => {
        const k = acc.touch(x.name), e = (per[k] = per[k] || { n: {}, p: {}, ids: {} });
        e.n[x[groupBy]] = (e.n[x[groupBy]] || 0) + 1;
        (e.ids[x.type] = e.ids[x.type] || []).push(...parseIdList(x.id));
        if (withPartners) partnersOf(x.list, x.name).forEach((nm) => { e.p[nm] = (e.p[nm] || 0) + 1; });
      });
      const cols = labels.map((h) => ({ h, t: 'n' })).concat([{ h: 'Всего записей-дублей', t: 'n', total: true }]);
      if (withPartners) cols.push({ h: 'С какими вузами (вуз — кол-во)', t: 's' });
      defTable(acc, tid, 'dup', title, 'list', cols, info ? { info } : {});
      Object.entries(per).forEach(([k, e]) => {
        const counts = labels.map((l) => e.n[l] || 0), row = counts.concat([counts.reduce((a, b) => a + b, 0)]);
        if (withPartners) row.push(Object.entries(e.p).sort((a, b) => b[1] - a[1]).map(([nm, c]) => `${nm} — ${c}`).join('\n'));
        (acc.tables[tid].rows[k] = acc.tables[tid].rows[k] || []).push(row);
        if (idCol) Object.entries(e.ids).forEach(([type, list]) => { const t = (acc.ids[tid] = acc.ids[tid] || {}); (t[k] = t[k] || []).push({ c: idCol, r: type, v: list.join(',') }); });
      });
    };
    build('dup_stud', pend.stud, 'type', 'Дубли ИИН: обучающиеся — по типу дубля', pend.infoS, 'studentid', true);
    build('dup_stud_reason', pend.stud, 'reason', 'Дубли ИИН: обучающиеся — по причине', '', '', false);
    build('dup_pps', pend.pps, 'type', 'Дубли ИИН: ППС (штатные в нескольких вузах) — по типу дубля', pend.infoP, 'tutorid', true);
  }

  D.dig = (rows, acc) => {
    let mode = 'first', fc = null, id = null, idc = [];
    for (const row of rows) {
      if (!row) continue;
      const a = row[0], ka = kk(a); if (!ka) continue;
      if (ka.startsWith('количество')) { mode = ka.includes('не оцифрован') ? 'second' : 'skip'; fc = null; id = null; continue; }
      if (isMarker(a)) {
        if (mode === 'first') {
          if (acc.tables.dig_main) continue;                         // повторная строка заголовков (новый формат)
          const tc = colWhere(row, (k) => k.includes('общее количество ошибок'));
          fc = colsOf(row, 1, (tc < 0 ? row.length : tc) - 1, {});   // поля, в т.ч. «Не оцифрованно кол-во дипломов 1991г»
          idc = idColsOf(row); id = 'dig_main';
          defTable(acc, id, 'dig', 'Оцифровка дипломов: ошибки по полям (включая не оцифрованные дипломы до 1991 г.)', 'list', fc.concat([{ idx: -1, h: 'Итого', t: 'n', total: true }]));
        } else if (mode === 'second') {
          if (acc.tables.dig_years) continue;
          const tc = colWhere(row, (k) => k.startsWith('итого'));
          fc = colsOf(row, 1, tc < 0 ? row.length : tc, { total: /^итого/ }); idc = []; id = 'dig_years';
          defTable(acc, id, 'dig', 'Не оцифровано дипломов выпускников — по годам выпуска', 'grid', fc);
        } else fc = null;
        continue;
      }
      if (isTotal(a) || !fc) continue;
      if (id === 'dig_main') { const arr = nums(row, fc); arr.push(arr.reduce((x, y) => x + y, 0)); pushRow(acc, id, a, arr); grabIds(acc, id, a, row, idc, ''); }
      else pushRow(acc, id, a, nums(row, fc));
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
    const idAll = idColsOf(head), idL = idAll.filter((x) => mc[1] === undefined || x.idx < mc[1]), idR = idAll.filter((x) => mc[1] !== undefined && x.idx > mc[1]);
    for (let r = hr + 1; r < rows.length; r++) {
      const row = rows[r]; if (!row) continue;
      if (isName(row[mc[0]]) && valL >= 0) {
        const deg = degL >= 0 ? lab(row[degL]) : '';
        pushRow(acc, 'gpa_stud', row[mc[0]], [deg, num(row[valL])]); grabIds(acc, 'gpa_stud', row[mc[0]], row, idL, deg);
      }
      if (mc[1] !== undefined && isName(row[mc[1]])) {
        const vs = yc.map((y) => num(row[y.c])), lv = lvl >= 0 ? lab(row[lvl]) : '';
        pushRow(acc, 'gpa_grad', row[mc[1]], [lv].concat(vs, [vs.reduce((a, b) => a + b, 0)])); grabIds(acc, 'gpa_grad', row[mc[1]], row, idR, lv);
      }
    }
  };

  D.pps = (rows, acc) => {
    const hr = findHeaderRow(rows); if (hr < 0) return;
    const head = rows[hr], all = colWhere(head, (k) => k === 'всего ошибок'); if (all < 0) return;
    const cols = colsOf(head, 1, all, { total: /^всего ошибок/ });
    defTable(acc, 'pps', 'pps', 'Полнота данных ППС: ошибки по полям', 'list', cols);
    const idc = idColsOf(head);
    for (let r = hr + 1; r < rows.length; r++) { const row = rows[r]; if (row && isName(row[0])) { pushRow(acc, 'pps', row[0], nums(row, cols)); grabIds(acc, 'pps', row[0], row, idc, ''); } }
  };

  D.mtb = (rows, acc) => {
    const hr = findHeaderRow(rows); if (hr < 0) return;
    const head = rows[hr], mc = [];
    head.forEach((v, c) => { if (isMarker(v)) mc.push(c); });
    if (colWhere(head, (k) => k === 'причина') < 0) return;      // реестр общежитий без списка ошибок
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
        if (cnt[c.h] > 1 && c.idx !== undefined && c.idx >= 0) { c.o = c.o || c.h; c.h = `${c.h} (столбец ${XLSX.utils.encode_col(c.idx)})`; }
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
    finalizeDup(acc);
    fixLabels(XLSX, acc.tables);
    const outTables = {};
    for (const [id, t] of Object.entries(acc.tables)) {
      const rowsOut = {};
      for (const [k, arr] of Object.entries(t.rows)) { const tk = alias[k] || k; rowsOut[tk] = (rowsOut[tk] || []).concat(arr); }
      outTables[id] = Object.assign({}, t, { rows: rowsOut, cols: t.cols.map((c) => { const o = { h: c.h, t: c.t }; if (c.total) o.total = 1; if (c.o) o.o = c.o; return o; }) });
    }
    const outIds = {};
    for (const [tid, byKey] of Object.entries(acc.ids)) {
      outIds[tid] = {};
      for (const [k, arr] of Object.entries(byKey)) { const tk = alias[k] || k; outIds[tid][tk] = (outIds[tid][tk] || []).concat(arr); }
    }
    const d = parseDateFromName(fileName);
    return {
      id: 's' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6),
      fileName: fileName || '', label: d.label, date: d.iso, names, m, alias, diag, missing, tables: outTables, ids: outIds, ver: 3,
      added: new Date().toISOString(),
    };
  }

  /* ---------------------------------------------------------------- показатели */
  const CRIT = { 'К': 'Корректность', 'З': 'Заполненность', 'И': 'Исполнение' };
  const INDICATORS = [
    { id: 'K1', tabs: ['gbd_stud'], sheet: 'gbd', crit: 'К', label: 'ГБД ФЛ: обучающиеся, ошибок', kind: 'log', get: (m) => m.gbd_stud || 0 },
    { id: 'K2', tabs: ['gbd_grad'], sheet: 'gbd', crit: 'К', label: 'ГБД ФЛ: выпускники, ошибок', kind: 'log', get: (m) => m.gbd_grad || 0 },
    { id: 'K3', tabs: ['gbd_pps'], sheet: 'gbd', crit: 'К', label: 'ГБД ФЛ: ППС, ошибок', kind: 'log', get: (m) => m.gbd_pps || 0 },
    { id: 'K4', tabs: ['dup_stud'], sheet: 'dup', crit: 'К', label: 'Дубли ИИН: обучающиеся, записей', kind: 'log', get: (m) => m.dup_stud || 0 },
    { id: 'K5', tabs: ['dup_pps'], sheet: 'dup', crit: 'К', label: 'Дубли ИИН: ППС, записей', kind: 'log', get: (m) => m.dup_pps || 0 },
    { id: 'K6', tabs: ['gen'], sheet: 'gen', crit: 'К', label: 'Дипломы сгенерированы, но не выпущены (3 года)', kind: 'log', get: (m) => m.gen || 0 },
    { id: 'K7', tabs: ['pps'], sheet: 'pps', crit: 'К', label: 'ППС: ФИО не кириллицей + дата рожд. ≠ ИИН', kind: 'log', get: (m) => m.pps_corr || 0 },
    { id: 'K8', tabs: ['mtb_dorm'], sheet: 'mtb', crit: 'К', label: 'Общежития с недостоверной стоимостью, шт.', kind: 'log', get: (m) => m.mtb_price || 0 },
    { id: 'Z1', tabs: ['osms_stud', 'osms_exp', 'osms_grad'], sheet: 'osms', crit: 'З', label: 'ОСМС: ошибок', kind: 'log', get: (m) => m.osms || 0 },
    { id: 'Z2', tabs: ['dig_main'], sheet: 'dig', crit: 'З', label: 'Оцифровка дипломов: ошибок (поля + не оцифровано до 1991 г.)', kind: 'log', get: (m) => m.dig_full || 0 },
    { id: 'Z3', tabs: ['gpa_stud', 'gpa_grad'], sheet: 'gpa', crit: 'З', label: 'GPA не заполнен (обуч. + выпуск. 2024–26)', kind: 'log', get: (m) => m.gpa || 0 },
    { id: 'Z4', tabs: ['pps'], sheet: 'pps', crit: 'З', label: 'ППС: незаполненные поля', kind: 'log', get: (m) => m.pps_compl || 0 },
    { id: 'Z5', tabs: ['admin'], sheet: 'admin', crit: 'З', label: 'Админ. отчёты: % недозаполнения', kind: 'lin', get: (m) => (m.admin_n ? (m.admin_sum / m.admin_n) * 100 : null) },
    { id: 'Z6', tabs: ['mtb_dorm', 'mtb_bldg'], sheet: 'mtb', crit: 'З', label: 'МТБ: незаполн. общежития + корпуса «нет данных»', kind: 'log',
      get: (m) => ((m.mtb_right || 0) > 0 || (m.mtb_unfilled || 0) > 0 ? (m.mtb_unfilled || 0) + (m.mtb_nodata || 0) : null) },
    { id: 'Z7', tabs: ['dig_years'], sheet: 'dig', crit: 'З', label: 'Не оцифровано дипломов выпускников 2024–26', kind: 'log', get: (m) => m.dig_recent || 0 },
    { id: 'I1', tabs: ['gos'], sheet: 'gos', crit: 'И', label: 'Госуслуги: не обработано заявок (срок истёк)', kind: 'log', get: (m) => m.gos || 0 },
  ];

  // Столбцы «сводного рейтинга по ошибкам»
  const UCOLS = [
    { id: 'u_gbd_stud', tabs: ['gbd_stud'], sheet: 'gbd', label: 'ГБД ФЛ: обучающиеся', src: 'ГБД ФЛ', get: (m) => m.gbd_stud || 0, on: 1 },
    { id: 'u_gbd_grad', tabs: ['gbd_grad'], sheet: 'gbd', label: 'ГБД ФЛ: выпускники', src: 'ГБД ФЛ', get: (m) => m.gbd_grad || 0, on: 1 },
    { id: 'u_gbd_pps', tabs: ['gbd_pps'], sheet: 'gbd', label: 'ГБД ФЛ: ППС', src: 'ГБД ФЛ', get: (m) => m.gbd_pps || 0, on: 1 },
    { id: 'u_osms_stud', tabs: ['osms_stud'], sheet: 'osms', group: 'osms', label: 'ОСМС: обучающиеся', src: 'ОСМС', get: (m) => m.osms_stud || 0, on: 1 },
    { id: 'u_osms_exp', tabs: ['osms_exp'], sheet: 'osms', group: 'osms', label: 'ОСМС: отчисленные', src: 'ОСМС', get: (m) => m.osms_exp || 0, on: 1 },
    { id: 'u_osms_grad', tabs: ['osms_grad'], sheet: 'osms', group: 'osms', label: 'ОСМС: выпускники', src: 'ОСМС', get: (m) => m.osms_grad || 0, on: 1 },
    { id: 'u_gen', tabs: ['gen'], sheet: 'gen', label: 'Дипломы сгенерированы, но не выпущены (3 года)', src: 'Генерация дипломов', get: (m) => m.gen || 0, on: 1 },
    { id: 'u_dup_stud', tabs: ['dup_stud'], sheet: 'dup', label: 'Дубли ИИН: обучающиеся', src: 'Дубли ИИН', get: (m) => m.dup_stud || 0, on: 1 },
    { id: 'u_dup_pps', tabs: ['dup_pps'], sheet: 'dup', label: 'Дубли ИИН: ППС', src: 'Дубли ИИН', get: (m) => m.dup_pps || 0, on: 1 },
    { id: 'u_dig_full', tabs: ['dig_main'], sheet: 'dig', label: 'Оцифровка дипломов (все поля)', src: 'Оцифровка', get: (m) => m.dig_full || 0, on: 1 },
    { id: 'u_gpa_stud', tabs: ['gpa_stud'], sheet: 'gpa', group: 'gpa', label: 'GPA не заполнен: обучающиеся', src: 'GPA', get: (m) => m.gpa_stud || 0, on: 1 },
    { id: 'u_gpa_grad', tabs: ['gpa_grad'], sheet: 'gpa', group: 'gpa', label: 'GPA не заполнен: выпускники 2024–26', src: 'GPA', get: (m) => m.gpa_grad || 0, on: 1 },
    { id: 'u_pps', tabs: ['pps'], sheet: 'pps', label: 'Полнота данных ППС (все ошибки)', src: 'ППС', get: (m) => (m.pps_compl || 0) + (m.pps_corr || 0), on: 1 },
    { id: 'u_admin', tabs: ['admin'], sheet: 'admin', label: 'Админ. отчёты: отчётов с неполными данными', src: 'Админ. отчёты', get: (m) => m.admin_bad || 0, on: 1 },
    { id: 'u_mtb', tabs: ['mtb_dorm', 'mtb_bldg'], sheet: 'mtb', label: 'МТБ: общежития с неверной ценой + без данных + корпуса «нет данных»', src: 'МТБ',
      get: (m) => (m.mtb_price || 0) + (m.mtb_unfilled || 0) + (m.mtb_nodata || 0), on: 1 },
    { id: 'u_dig_recent', tabs: ['dig_years'], sheet: 'dig', label: 'Не оцифровано дипломов выпускников 2024–26', src: 'Оцифровка', get: (m) => m.dig_recent || 0, on: 1 },
    { id: 'u_gos', tabs: ['gos'], sheet: 'gos', label: 'Госуслуги: не обработано заявок (срок истёк)', src: 'Госуслуги', get: (m) => m.gos || 0, on: 1 },
  ];

  const defaultSettings = () => ({
    critWeights: { 'К': 1, 'З': 1, 'И': 1 },
    indWeights: Object.fromEntries(INDICATORS.map((i) => [i.id, 1])),
    ucolOn: Object.fromEntries(UCOLS.map((u) => [u.id, u.on])),
    exclude: ['военн'],
    excl: { sheets: [], tables: [], fields: {} },   // что не учитывать: листы / таблицы / отдельные поля
    activeNames: null,           // null = все ОВПО; иначе массив названий
    trend: 2,                    // порог «существенного» изменения балла
  });


  /* ---------------------------------------------------------------- что учитывать: листы / таблицы / поля */
  const SECTION_DEFS = [
    { id: 'gbd', title: 'ГБД ФЛ', tables: [['gbd_stud', 'Обучающиеся'], ['gbd_grad', 'Выпускники'], ['gbd_pps', 'ППС']] },
    { id: 'osms', title: 'ОСМС', tables: [['osms_stud', 'Обучающиеся'], ['osms_exp', 'Отчисленные'], ['osms_grad', 'Выпускники']] },
    { id: 'gen', title: 'Генерация номеров дипломов', tables: [['gen', 'По годам']] },
    { id: 'dup', title: 'Дубли ИИН', tables: [['dup_stud', 'Обучающиеся (по типу дубля)'], ['dup_pps', 'ППС (по типу дубля)']] },
    { id: 'dig', title: 'Оцифровка дипломов', tables: [['dig_main', 'Ошибки по полям'], ['dig_years', 'Не оцифровано выпускников 2024–26 (по годам)']] },
    { id: 'gpa', title: 'GPA', tables: [['gpa_stud', 'Обучающиеся'], ['gpa_grad', 'Выпускники (по годам выпуска)']] },
    { id: 'pps', title: 'Полнота данных ППС', tables: [['pps', 'Ошибки по полям']] },
    { id: 'mtb', title: 'Материально-техническая база', tables: [['mtb_dorm', 'Общежития (цена, незаполненные)'], ['mtb_bldg', 'Корпуса и сооружения («нет данных»)']] },
    { id: 'gos', title: 'Государственные услуги', tables: [['gos', 'По услугам']] },
    { id: 'admin', title: 'Административные отчёты', tables: [['admin', 'По отчётам']] },
  ];
  const SHEET_OF = {};
  SECTION_DEFS.forEach((sec) => sec.tables.forEach(([id]) => { SHEET_OF[id] = sec.id; }));
  ['osms_cat', 'osms_stud_2', 'osms_exp_2', 'osms_grad_2'].forEach((id) => { SHEET_OF[id] = 'osms'; });
  SHEET_OF.dup_stud_reason = 'dup';
  const ALL_TABLES_OF = {};
  Object.entries(SHEET_OF).forEach(([t, sh]) => { (ALL_TABLES_OF[sh] = ALL_TABLES_OF[sh] || []).push(t); });
  // таблицы, из которых нельзя исключать отдельные поля (одно число на ОВПО или строки без полей)
  const NO_FIELDS = new Set(['gpa_stud', 'mtb_dorm', 'osms_cat', 'osms_stud_2', 'osms_exp_2', 'osms_grad_2', 'dup_stud_reason']);

  // «Поля» таблицы, которые можно исключать: столбцы с числом ошибок; для админ. отчётов — коды отчётов; для МТБ-корпусов — виды объектов
  function errorFields(t) {
    if (!t || NO_FIELDS.has(t.id)) return [];
    if (t.id === 'admin') return Object.keys(t.names || {});
    if (t.id === 'mtb_bldg') return t.cols.map((c) => c.h);
    return t.cols.filter((c) => c.t === 'n' && !c.total).map((c) => c.h);
  }
  const CORR_RE = /фио не на кириллице|дата рождения не соответствует/i;

  function scopeHelpers(S, tables) {
    const ex = (S && S.excl) || {};
    const sheets = new Set(ex.sheets || []), tabs = new Set(ex.tables || []), fields = {};
    Object.entries(ex.fields || {}).forEach(([k, v]) => { fields[k] = new Set(v); });
    const errF = (tid) => (tables && tables[tid] ? errorFields(tables[tid]) : []);
    const isTable = (tid) => {
      if (tabs.has(tid) || sheets.has(SHEET_OF[tid])) return true;
      const f = fields[tid];
      if (f && f.size) { const all = errF(tid); if (all.length && all.every((x) => f.has(x))) return true; }
      return false;
    };
    const scoped = (sh) => (SECTION_DEFS.find((x) => x.id === sh) || { tables: [] }).tables.map((x) => x[0]);
    const isSheet = (sh) => sheets.has(sh) || scoped(sh).every((id) => isTable(id));
    const isField = (tid, label) => isTable(tid) || !!(fields[tid] && fields[tid].has(label));
    const any = sheets.size > 0 || tabs.size > 0 || Object.values(fields).some((f) => f.size > 0);
    return { isSheet, isTable, isField, any, sheets, tabs, fields, errF };
  }

  const collapseRows = (t, rows) => {
    const out = t.cols.map((c) => (c.t === 's' ? '' : 0));
    rows.forEach((r) => r.forEach((v, i) => { if (t.cols[i].t !== 's') out[i] += (v || 0); }));
    return out;
  };
  const EFF = new WeakMap(), ACT = new WeakMap();
  const sigOf = (list) => list.length + ':' + list.reduce((h, x) => ((h * 31 + String(x).length) | 0), 7);
  // Список действующих ОВПО определяет состав расчёта: организации из списка, которых нет в файле, считаются ОВПО без ошибок
  function withActiveNames(snap, S) {
    const list = S && S.activeNames;
    if (!list || !list.length) return snap;
    const sig = sigOf(list), c = ACT.get(snap);
    if (c && c.sig === sig) return c.val;
    const names = Object.assign({}, snap.names), m = Object.assign({}, snap.m); let added = 0;
    for (const n of list) {
      const k = key(n), al = snap.alias && snap.alias[k];
      if (k in names || (al && al in names)) continue;
      names[k] = norm(n); m[k] = {}; added++;
    }
    const val = added ? Object.assign({}, snap, { names, m, baseSnap: snap }) : snap;
    ACT.set(snap, { sig, val });
    return val;
  }
  // «Эффективные» данные: метрики ОВПО пересчитаны с учётом исключённых листов, таблиц и полей
  function effective(snap, S) {
    snap = withActiveNames(snap, S);
    if (!S || !S.excl) return snap;
    const sig = JSON.stringify(S.excl);
    if (snap.effSig === sig) return snap;
    const H = scopeHelpers(S, snap.tables);
    if (!H.any) return snap;
    const cached = EFF.get(snap); if (cached && cached.sig === sig) return cached.val;
    const T = snap.tables || {};
    const excl = new Set(Object.keys(SHEET_OF).filter((id) => H.isTable(id)));
    const m2 = {};
    const fset = (id) => H.fields[id] || new Set();
    for (const k of Object.keys(snap.names)) {
      const m = snap.m[k] || {}, o = Object.assign({}, m);
      const incTotal = (id) => {
        const t = T[id]; const rows = t && t.rows[k]; if (!t || !rows) return 0;
        const v = collapseRows(t, rows), ti = t.cols.findIndex((c) => c.total), f = fset(id);
        if (!f.size) return ti >= 0 ? v[ti] : t.cols.reduce((s, c, i) => s + (c.t === 'n' ? v[i] : 0), 0);
        return t.cols.reduce((s, c, i) => s + (c.t === 'n' && !c.total && !f.has(c.h) ? v[i] : 0), 0);
      };
      const adj = (metric, id) => {
        if (excl.has(id)) { o[metric] = 0; return true; }
        if (fset(id).size && T[id]) { o[metric] = incTotal(id); return true; }
        return false;
      };
      adj('gbd_stud', 'gbd_stud'); adj('gbd_grad', 'gbd_grad'); adj('gbd_pps', 'gbd_pps');
      const a1 = adj('osms_stud', 'osms_stud'), a2 = adj('osms_exp', 'osms_exp'), a3 = adj('osms_grad', 'osms_grad');
      if (a1 || a2 || a3) o.osms = (o.osms_stud || 0) + (o.osms_exp || 0) + (o.osms_grad || 0);
      adj('gen', 'gen'); adj('dup_stud', 'dup_stud'); adj('dup_pps', 'dup_pps');
      adj('dig_full', 'dig_main'); adj('dig_recent', 'dig_years');
      const g1 = adj('gpa_stud', 'gpa_stud'), g2 = adj('gpa_grad', 'gpa_grad');
      if (g1 || g2) o.gpa = (o.gpa_stud || 0) + (o.gpa_grad || 0);
      adj('gos', 'gos');
      // ППС: «корректность» (ФИО, дата рождения) и «заполненность» (остальные поля) считаем по оставшимся полям
      if (excl.has('pps')) { o.pps_corr = 0; o.pps_compl = 0; }
      else if (fset('pps').size && T.pps && T.pps.rows[k]) {
        const t = T.pps, v = collapseRows(t, T.pps.rows[k]), f = fset('pps'); let corr = 0, compl = 0;
        t.cols.forEach((c, i) => { if (c.t !== 'n' || c.total || f.has(c.h)) return; if (CORR_RE.test(c.h)) corr += v[i]; else compl += v[i]; });
        o.pps_corr = corr; o.pps_compl = compl;
      } else if (fset('pps').size) { o.pps_corr = 0; o.pps_compl = 0; }
      // МТБ
      if (excl.has('mtb_dorm')) { o.mtb_price = 0; o.mtb_unfilled = 0; }
      if (excl.has('mtb_bldg')) { o.mtb_nodata = 0; o.mtb_right = 0; }
      else if (fset('mtb_bldg').size && T.mtb_bldg) {
        const t = T.mtb_bldg, f = fset('mtb_bldg'); let c = 0;
        (t.rows[k] || []).forEach((r) => t.cols.forEach((col, i) => { if (!f.has(col.h) && String(r[i] || '').trim().toLowerCase() === 'нет данных') c++; }));
        o.mtb_nodata = c;
      }
      // административные отчёты: пересчёт по оставшимся отчётам
      if (!excl.has('admin') && fset('admin').size && T.admin) {
        const t = T.admin; let n = 0, sum = 0, bad = 0;
        (t.rows[k] || []).forEach((r) => {
          if (fset('admin').has(r[0])) return;
          const st = String(t.stat[r[3]] || '').toUpperCase(); const w = st.startsWith('НЕТ') ? 1 : st.includes('ЧАСТИЧНО') ? 0.5 : 0;
          n++; sum += w; if (w > 0) bad++;
        });
        o.admin_n = n; o.admin_sum = sum; o.admin_bad = bad;
      }
      m2[k] = o;
    }
    const val = Object.assign({}, snap, { m: m2, excl, effSig: sig });
    EFF.set(snap, { sig, val });
    return val;
  }
  const goneBy = (snap, x) => (snap.missing || []).includes(x.sheet) || !!(snap.excl && x.tabs.every((id) => snap.excl.has(id)));

  /* ---------------------------------------------------------------- фильтр ОВПО */
  function includedKeys(snap, settings) {
    snap = withActiveNames(snap, settings);
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
    snap = effective(snap, S);
    const vals = {}, mx = {};
    for (const ind of INDICATORS) {
      vals[ind.id] = {};
      let mxv = -Infinity;
      for (const k of keys) {
        const v = goneBy(snap, ind) ? null : ind.get(snap.m[k] || {});
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
    snap = effective(snap, S);
    const has = (met) => Object.values(snap.m).some((x) => x[met] !== undefined);
    const splitOk = { osms: has('osms_stud') || has('osms_exp') || has('osms_grad'), gpa: has('gpa_stud') || has('gpa_grad') };
    const rows = keys.map((k) => {
      const m = snap.m[k] || {};
      const parts = {};
      let total = 0;
      for (const u of UCOLS) {
        const gone = goneBy(snap, u) || (u.group && !splitOk[u.group]);
        const v = gone ? null : u.get(m); parts[u.id] = v;
        if (S.ucolOn[u.id] && v !== null) total += v;
      }
      // срез, сохранённый старой версией сайта (без разбивки на категории): берём общий итог, чтобы сумма не потерялась
      for (const grp of ['osms', 'gpa']) {
        if (splitOk[grp] || (snap.missing || []).includes(grp) || (snap.excl && SHEET_OF[grp === 'osms' ? 'osms_stud' : 'gpa_stud'] && ['osms_stud', 'osms_exp', 'osms_grad', 'gpa_stud', 'gpa_grad'].filter((id) => SHEET_OF[id] === grp).every((id) => snap.excl.has(id)))) continue;
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
    snaps = snaps.map((s) => effective(s, settings));
    const keysBySnap = snaps.map((s) => includedKeys(s, settings));
    const common = {};
    for (const ind of INDICATORS) {
      let mx = 0;
      snaps.forEach((s, i) => keysBySnap[i].forEach((k) => { const v = goneBy(s, ind) ? null : ind.get(s.m[k] || {}); if (v !== null && v > mx) mx = v; }));
      common[ind.id] = mx;
    }
    return snaps.map((s, i) => rate(s, keysBySnap[i], settings, common));
  }

  const api = {
    norm, key, ALIAS_RULES, SHEETS, INDICATORS, UCOLS, CRIT,
    parseSnapshot, parseDateFromName, includedKeys, rate, rateAll, unified, defaultSettings,
    effective, withActiveNames, scopeHelpers, errorFields, SECTION_DEFS, SHEET_OF,
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.VuzCore = api;
})(typeof window !== 'undefined' ? window : globalThis);
