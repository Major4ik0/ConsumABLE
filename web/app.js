/**
 * ConsumABLE — Клиентская веб-версия строевой записки
 * Принцип: Strict In-Memory (Zero Persistence).
 * Все данные хранятся ТОЛЬКО в оперативной памяти браузера,
 * не используют localStorage / sessionStorage / IndexedDB и никуда не передаются.
 */

// 1. ИСХОДНОЕ СОСТОЯНИЕ (IN-MEMORY STATE)
const state = {
  settings: {
    reportDate: getTodayDateString(),
    unitTitle: "* курс * факультета (автоматизированных систем управления)",
    chiefTitle: (typeof localStorage !== 'undefined' && localStorage.getItem('consumable_chief_title')) || "Начальник * курса * факультета",
    commanderRank: (typeof localStorage !== 'undefined' && localStorage.getItem('consumable_commander_rank')) || "",
    commanderName: (typeof localStorage !== 'undefined' && localStorage.getItem('consumable_commander_name')) || "",
    chiefName: ""
  },
  cadets: [],
  absences: [],
  absenceTypes: [
    { id: 1, name: "Дежурный по курсу", code: "Н", category: "Наряд" },
    { id: 2, name: "Дневальный по курсу", code: "Н", category: "Наряд" },
    { id: 3, name: "Патруль", code: "Н", category: "Наряд" },
    { id: 4, name: "Увольнение", code: "УВ", category: "Увольнение" },
    { id: 5, name: "Госпиталь", code: "Г", category: "Госпиталь" },
    { id: 6, name: "Лазарет", code: "Л", category: "Лазарет" },
    { id: 7, name: "Командировка", code: "К", category: "Командировка" },
    { id: 8, name: "МОГ", code: "Н", category: "Наряд" }
  ],
  standardCategories: ["Наряд", "Госпиталь", "Лазарет", "Командировка", "Похороны", "Увольнение"],
  currentSlice: "total",
  activePreviewSlice: "total",
  stagingCadets: []
};

let nextCadetId = 1;
let nextAbsenceId = 1;
let nextTypeId = 8;
let cadetModalEditId = null;

// Вспомогательная функция для сегодняшней даты YYYY-MM-DD
function getTodayDateString() {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const day = String(now.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

// Форматирование даты в формат ДД.ММ.ГГГГ г.
function formatDateRussian(dateStr) {
  if (!dateStr) return "";
  const parts = dateStr.split('-');
  if (parts.length === 3) {
    return `${parts[2]}.${parts[1]}.${parts[0]} г.`;
  }
  return dateStr;
}

// Автоподстановка курса и факультета в шапку строевой записки
function updateHeaderText(text, faculty = null, course = null) {
  if (!text) return text;
  let res = String(text);
  if (course !== null && course !== undefined) {
    res = res.replace(/(\*|\d+)\s+курса(?![а-яА-ЯёЁ])/gi, `${course} курса`);
    res = res.replace(/(\*|\d+)\s+курс(?![а-яА-ЯёЁ])/gi, `${course} курс`);
  }
  if (faculty !== null && faculty !== undefined) {
    res = res.replace(/(\*|\d+)\s+факультет([а-яА-ЯёЁ]*)/gi, `${faculty} факультет$2`);
  }
  return res;
}

function updateHeaderSettings(faculty = null, course = null) {
  state.settings.unitTitle = updateHeaderText(state.settings.unitTitle, faculty, course);
  state.settings.chiefTitle = updateHeaderText(state.settings.chiefTitle, faculty, course);

  const unitTitleEl = document.getElementById("unitTitle");
  if (unitTitleEl) unitTitleEl.value = state.settings.unitTitle;

  const chiefTitleEl = document.getElementById("chiefTitle");
  if (chiefTitleEl) chiefTitleEl.value = state.settings.chiefTitle;
}

function syncChiefName() {
  const rk = (state.settings.commanderRank || "").trim();
  const nm = (state.settings.commanderName || "").trim();
  if (rk && nm) {
    state.settings.chiefName = `${rk}                          ${nm}`;
  } else if (rk) {
    state.settings.chiefName = rk;
  } else if (nm) {
    state.settings.chiefName = nm;
  } else {
    state.settings.chiefName = "";
  }
  const chiefNameEl = document.getElementById("chiefName");
  if (chiefNameEl) chiefNameEl.value = state.settings.chiefName;
}

const SIGNATURE_RANKS = [
  'генерал-майор', 'подполковник', 'старший лейтенант', 'младший лейтенант',
  'старший прапорщик', 'старший сержант', 'младший сержант',
  'полковник', 'капитан', 'майор', 'лейтенант', 'старшина',
  'прапорщик', 'сержант', 'ефрейтор', 'рядовой'
];

function parseSignatureFromWorkbook(wb) {
  let foundTitle = null;
  let foundRank = null;
  let foundName = null;

  if (!wb || !wb.SheetNames) return { foundTitle, foundRank, foundName };

  for (const sname of wb.SheetNames) {
    const ws = wb.Sheets[sname];
    if (!ws) continue;
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
    if (!rows || rows.length === 0) continue;

    for (const r of rows) {
      if (!r || !Array.isArray(r)) continue;
      const cells = r.filter(c => c !== null && c !== undefined && String(c).trim() !== '').map(c => String(c).trim());
      const line = cells.join(' ');
      if (!line) continue;

      if (!foundTitle && /начальник\s+.*курс/i.test(line)) {
        const mTitle = line.match(/(начальник\s+[^,\n]+)/i);
        if (mTitle) foundTitle = mTitle[1].trim();
      }

      if (!(foundRank && foundName)) {
        for (const rk of SIGNATURE_RANKS) {
          const re = new RegExp('(?:^|[^а-яёА-ЯЁ])(' + rk + ')(?:[^а-яёА-ЯЁ]|$)', 'i');
          if (re.test(line)) {
            foundRank = rk;
            const idx = line.toLowerCase().indexOf(rk);
            const rem = line.slice(idx + rk.length).trim();
            const mFio = rem.match(/([А-ЯЁа-яё\-]+(?:\s+[А-ЯЁ]\.){1,2}|[А-ЯЁ]\.\s*[А-ЯЁа-яё\-]+|[А-ЯЁ]\.\s*[А-ЯЁ]\.\s*[А-ЯЁа-яё\-]+|[А-ЯЁа-яё\-]{3,})/);
            if (mFio) {
              foundName = mFio[0].trim();
            }
            break;
          }
        }
      }

      if (foundTitle && foundRank && foundName) break;
    }
    if (foundRank && foundName) break;
  }

  return { foundTitle, foundRank, foundName };
}

// 2. АЛГОРИТМЫ ОБРАБОТКИ ФИО И ПОЛА (1-в-1 ИЗ EXCEL_MANAGER.PY)

/**
 * Преобразование полного ФИО в краткое: «Фамилия И.О.»
 */
function makeShortFio(fullName) {
  if (!fullName) return "";
  const name = String(fullName).trim();
  const m = name.match(/^([А-ЯЁа-яё\-]+)\s+([А-ЯЁа-яё])\.\s*([А-ЯЁа-яё])\.$/);
  if (m) {
    return `${m[1]} ${m[2].toUpperCase()}.${m[3].toUpperCase()}.`;
  }
  const parts = name.split(/\s+/).filter(Boolean);
  if (parts.length >= 3) {
    const i1 = parts[1][0].toUpperCase();
    const i2 = parts[2][0].toUpperCase();
    return `${parts[0]} ${i1}.${i2}.`;
  } else if (parts.length === 2) {
    if (parts[1].includes('.')) {
      const initials = parts[1].replace(/[^а-яёА-ЯЁa-zA-Z]/g, '').toUpperCase();
      if (initials.length >= 2) {
        return `${parts[0]} ${initials[0]}.${initials[1]}.`;
      } else if (initials.length === 1) {
        return `${parts[0]} ${initials[0]}.`;
      }
    }
    return `${parts[0]} ${parts[1][0].toUpperCase()}.`;
  }
  return name;
}

/**
 * Автоопределение пола курсанта по окончаниям отчества/фамилии или явному полю
 * По спецификации: -вич, -ич, -лы, -оглы -> М; -вна, -чна, -кызы -> Ж; по умолчанию -> М
 */
function detectGender(fio, explicitVal) {
  if (explicitVal !== undefined && explicitVal !== null && String(explicitVal).trim()) {
    const v = String(explicitVal).trim().toUpperCase();
    if (v.startsWith('Ж') || v.startsWith('F') || v.includes('ЖЕН')) {
      return 'Ж';
    }
    if (v.startsWith('М') || v.startsWith('M') || v.includes('МУЖ')) {
      return 'М';
    }
  }

  const fullLower = String(fio || '').trim().toLowerCase();
  if (fullLower.includes('кызы') || fullLower.endsWith('-кызы')) {
    return 'Ж';
  }
  if (fullLower.includes('оглы') || fullLower.endsWith('-оглы')) {
    return 'М';
  }

  const parts = fullLower.split(/\s+/).filter(Boolean);
  // Проверяем каждое слово начиная со второго (отчество может стоять 2-м или 3-м словом)
  for (let i = 1; i < parts.length; i++) {
    const p = parts[i];
    if (p.endsWith('вна') || p.endsWith('чна') || p.endsWith('кызы')) {
      return 'Ж';
    }
    if (p.endsWith('вич') || p.endsWith('ич') || p.endsWith('лы') || p.endsWith('оглы')) {
      return 'М';
    }
  }

  return 'М';
}

// 3. ПАРСИНГ EXCEL С ЛИЧНЫМ СОСТАВОМ (SheetJS)

function parseCadetsSheet(ws, defaultGroup = "", seenGroups = []) {
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
  if (!rows || rows.length === 0) return [];

  let headerIdx = -1;
  const colMap = {};

  const maxHeaderSearch = Math.min(15, rows.length);
  for (let r_i = 0; r_i < maxHeaderSearch; r_i++) {
    const row = (rows[r_i] || []).map(c => String(c || '').trim().toLowerCase());
    const hasFio = row.some(c => ['фио', 'ф.и.о.', 'фамили', 'курсант', 'слушател', 'обучающ'].some(k => c.includes(k)));
    const hasRank = row.some(c => ['звание', 'в/зв', 'в/звание', 'воинск'].some(k => c.includes(k)));
    const hasGrp = row.some(c => ['групп', 'взвод', 'подразд', 'отделен', 'к/о', 'класс'].some(k => c.includes(k)));
    const hasNum = row.some(c => c.includes('№') || c.includes('п/п'));

    if ((hasFio && (hasRank || hasGrp)) || (hasFio && hasNum)) {
      headerIdx = r_i;
      row.forEach((cellText, c_i) => {
        if (['в/зв', 'звание', 'воинск'].some(k => cellText.includes(k))) {
          colMap.rank = c_i;
        } else if (['фио', 'ф.и.о.', 'фамилия имя отчество', 'фамилия, имя, отчество', 'фамилия имя'].includes(cellText)) {
          colMap.fio = c_i;
        } else if (cellText.includes('сокр') && cellText.includes('фио')) {
          colMap.short_fio = c_i;
        } else if (['фио', 'ф.и.о.', 'фамилия'].some(k => cellText.includes(k)) && !('fio' in colMap)) {
          colMap.fio = c_i;
        } else if (cellText.includes('фамил') && !('fio' in colMap)) {
          colMap.surname = c_i;
        } else if (cellText === 'имя' || cellText.startsWith('имя ')) {
          colMap.name = c_i;
        } else if (cellText.includes('отчеств')) {
          colMap.patronymic = c_i;
        } else if (['групп', 'взвод', 'подразд', 'отделен', 'к/о', 'класс'].some(k => cellText.includes(k))) {
          colMap.group = c_i;
        } else if (['пол', 'м/ж'].some(k => cellText.includes(k))) {
          colMap.gender = c_i;
        }
      });
      break;
    }
  }

  const startRow = headerIdx !== -1 ? headerIdx + 1 : 0;

  if (!('fio' in colMap) && !('surname' in colMap)) {
    const colScores = {};
    const sampleRows = rows.slice(startRow, startRow + 15);
    sampleRows.forEach(r => {
      (r || []).forEach((cell, c_i) => {
        const val = String(cell || '').trim();
        if (!val) return;
        const words = val.split(/\s+/).filter(Boolean);
        if (words.length >= 2 && words.length <= 3 && /[а-яА-ЯёЁ]/.test(val)) {
          colScores[c_i] = (colScores[c_i] || 0) + 1;
        }
      });
    });

    const bestCols = Object.keys(colScores);
    if (bestCols.length > 0) {
      bestCols.sort((a, b) => colScores[b] - colScores[a]);
      const bestFioCol = Number(bestCols[0]);
      colMap.fio = bestFioCol;
      if (bestFioCol > 0) colMap.rank = bestFioCol - 1;
      colMap.group = bestFioCol + 1;
    } else {
      colMap.rank = 0;
      colMap.fio = 1;
      colMap.group = 2;
    }
  }

  const cadets = [];
  let currentGroup = defaultGroup || "1";

  for (let r_i = startRow; r_i < rows.length; r_i++) {
    const row = rows[r_i];
    if (!row || row.every(c => c === null || String(c).trim() === '')) {
      continue;
    }

    let fio = "";
    if ('fio' in colMap && colMap.fio < row.length && row[colMap.fio] !== null) {
      fio = String(row[colMap.fio]).trim();
    } else if ('surname' in colMap) {
      const sur = (colMap.surname < row.length && row[colMap.surname] !== null) ? String(row[colMap.surname]).trim() : "";
      const nm = (colMap.name < row.length && row[colMap.name] !== null) ? String(row[colMap.name]).trim() : "";
      const pat = (colMap.patronymic < row.length && row[colMap.patronymic] !== null) ? String(row[colMap.patronymic]).trim() : "";
      fio = [sur, nm, pat].filter(Boolean).join(" ").trim();
    }

    const fioLower = fio.toLowerCase();
    if (!fio || ['фио', 'ф.и.о.', 'фамилия', 'итого', 'всего'].includes(fioLower)) {
      continue;
    }

    const words = fio.split(/\s+/).filter(Boolean);
    if (words.length === 1 && (fioLower.includes('взвод') || fioLower.includes('группа') || fioLower.includes('отделение'))) {
      currentGroup = fio;
      continue;
    }

    let rank = "";
    if ('rank' in colMap && colMap.rank < row.length && row[colMap.rank] !== null) {
      rank = String(row[colMap.rank]).trim();
      if (/^\d+$/.test(rank)) {
        rank = "курсант";
      }
    }
    if (!rank) {
      rank = "курсант";
    }

    let groupRaw = null;
    if ('group' in colMap && colMap.group < row.length) {
      groupRaw = row[colMap.group];
    }

    let groupNum = currentGroup;
    if (groupRaw !== null && String(groupRaw).trim() !== '') {
      const grpStr = String(groupRaw).trim();
      const numVal = Number(grpStr);
      if (!isNaN(numVal) && Number.isInteger(numVal)) {
        groupNum = String(numVal);
      } else {
        groupNum = grpStr;
      }
    }

    if (!seenGroups.includes(groupNum)) {
      seenGroups.push(groupNum);
    }
    const sortOrder = seenGroups.indexOf(groupNum);

    let shortFio = "";
    if ('short_fio' in colMap && colMap.short_fio < row.length && row[colMap.short_fio]) {
      shortFio = String(row[colMap.short_fio]).trim();
    }
    if (!shortFio) {
      shortFio = makeShortFio(fio);
    }

    let rawGender = null;
    if ('gender' in colMap && colMap.gender < row.length) {
      rawGender = row[colMap.gender];
    }
    const gender = detectGender(fio, rawGender);

    const cg = getCourseFromGroup(groupNum);
    const crsVal = cg !== "?" ? parseInt(cg, 10) : 1;

    cadets.push({
      id: nextCadetId++,
      rank: rank,
      fio: fio,
      short_fio: shortFio,
      group_num: groupNum,
      course: crsVal,
      gender: gender,
      sort_order: sortOrder
    });
  }

  return cadets;
}

// 4. ГЕНЕРАЦИЯ EXCEL ЧЕРЕЗ EXCELJS (1-В-1 С EXCEL_MANAGER.PY)

function getColumnLetter(colNum) {
  let temp = colNum;
  let letter = '';
  while (temp > 0) {
    let rem = (temp - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    temp = Math.floor((temp - 1) / 26);
  }
  return letter;
}

function applyStyleToRange(ws, topRow, leftCol, bottomRow, rightCol, { font, alignment, border, fill } = {}) {
  for (let r = topRow; r <= bottomRow; r++) {
    for (let c = leftCol; c <= rightCol; c++) {
      const cell = ws.getCell(r, c);
      if (font) cell.font = font;
      if (alignment) cell.alignment = alignment;
      if (border) cell.border = border;
      if (fill) cell.fill = fill;
    }
  }
}

function buildExcelSheet(ws, sheetTitleSuffix, cadets, absences, categories) {
  const boldFont = { name: "Times New Roman", size: 12, bold: true };
  const normFont = { name: "Times New Roman", size: 12 };
  const titleFont = { name: "Times New Roman", size: 12, bold: true };

  const headerFill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFF2F2F2" }
  };

  const thinBorder = {
    top: { style: "thin", color: { argb: "FF000000" } },
    left: { style: "thin", color: { argb: "FF000000" } },
    bottom: { style: "thin", color: { argb: "FF000000" } },
    right: { style: "thin", color: { argb: "FF000000" } }
  };

  const centerAlign = { horizontal: "center", vertical: "middle", wrapText: true };

  // 1. Заголовки (строки 2-4)
  const fullTitle = ("РАЗВЕРНУТАЯ СТРОЕВАЯ ЗАПИСКА " + (sheetTitleSuffix || "")).trim();
  ws.mergeCells("B2:L2");
  ws.getCell("B2").value = fullTitle;
  applyStyleToRange(ws, 2, 2, 2, 12, { font: titleFont, alignment: centerAlign });

  ws.mergeCells("B3:L3");
  ws.getCell("B3").value = state.settings.unitTitle;
  applyStyleToRange(ws, 3, 2, 3, 12, { font: normFont, alignment: centerAlign });

  ws.mergeCells("B4:L4");
  const dtVal = formatDateRussian(state.settings.reportDate);
  ws.getCell("B4").value = dtVal;
  applyStyleToRange(ws, 4, 2, 4, 12, { font: normFont, alignment: centerAlign });
  ws.getRow(4).height = 17.0;

  // 2. Группы по порядку добавления / sort_order
  const groupOrderMap = {};
  cadets.forEach(c => {
    const grp = c.group_num || "1";
    if (!(grp in groupOrderMap)) {
      groupOrderMap[grp] = c.sort_order || 0;
    }
  });
  const groups = Object.keys(groupOrderMap).sort((a, b) => groupOrderMap[a] - groupOrderMap[b]);

  const startRow = 5;
  ws.getRow(startRow).height = 18.0;

  const headers = ["№ п/п", "Подразд.", "П/С", "Н/Л", ...categories];
  const v_col1 = 2 + headers.length;
  const v_col2 = v_col1 + 1;
  const v_col1_letter = getColumnLetter(v_col1);
  const v_col2_letter = getColumnLetter(v_col2);

  // Заголовки сводной таблицы
  headers.forEach((h, idx) => {
    const colIdx = 2 + idx;
    const cell = ws.getCell(startRow, colIdx);
    cell.value = h;
    cell.font = normFont;
    cell.fill = headerFill;
    cell.border = thinBorder;
    cell.alignment = centerAlign;
  });

  // Колонка «ВСЕГО» (объединение v_col1:v_col2)
  ws.mergeCells(startRow, v_col1, startRow, v_col2);
  ws.getCell(startRow, v_col1).value = "ВСЕГО";
  applyStyleToRange(ws, startRow, v_col1, startRow, v_col2, {
    font: normFont,
    fill: headerFill,
    border: thinBorder,
    alignment: centerAlign
  });

  const first_cat_col = 6;
  const last_cat_col = first_cat_col + categories.length - 1;
  const first_cat_letter = getColumnLetter(first_cat_col);
  const last_cat_letter = getColumnLetter(last_cat_col);

  let currRow = startRow + 1;
  const first_grp_row = currRow;

  // Строки подразделений
  groups.forEach((grp, idx) => {
    ws.getRow(currRow).height = 17.0;
    const grpCadets = cadets.filter(c => String(c.group_num) === String(grp));
    const ps = grpCadets.length;
    const grpAbs = absences.filter(a => String(a.group_num) === String(grp));

    const rowCounts = {};
    categories.forEach(cat => {
      rowCounts[cat] = grpAbs.filter(a => a.category === cat).length;
    });

    ws.getCell(currRow, 2).value = idx + 1;
    ws.getCell(currRow, 3).value = grp;
    ws.getCell(currRow, 4).value = ps;
    ws.getCell(currRow, 5).value = {
      formula: `D${currRow}-${v_col1_letter}${currRow}`,
      result: ps - grpAbs.length
    };

    categories.forEach((cat, c_offset) => {
      const colNum = first_cat_col + c_offset;
      const cVal = rowCounts[cat] > 0 ? rowCounts[cat] : null;
      ws.getCell(currRow, colNum).value = cVal;
    });

    ws.mergeCells(currRow, v_col1, currRow, v_col2);
    ws.getCell(currRow, v_col1).value = {
      formula: `SUM(${first_cat_letter}${currRow}:${last_cat_letter}${currRow})`,
      result: grpAbs.length
    };

    applyStyleToRange(ws, currRow, 2, currRow, v_col2, {
      font: normFont,
      border: thinBorder,
      alignment: centerAlign
    });

    currRow++;
  });

  const last_grp_row = groups.length > 0 ? currRow - 1 : currRow;

  // Строки раздельного расхода (Мужчины / Женщины), если на листе представлены оба пола
  const hasMales = cadets.some(c => c.gender === 'М' || c.gender === 'male');
  const hasFemales = cadets.some(c => c.gender === 'Ж' || c.gender === 'female');

  if (hasMales && hasFemales) {
    const genderConfigs = [
      { label: "Мужчины", val: "М", val2: "male", numIdx: groups.length + 1 },
      { label: "Женщины", val: "Ж", val2: "female", numIdx: groups.length + 2 }
    ];

    genderConfigs.forEach(g => {
      ws.getRow(currRow).height = 17.0;
      const gCadets = cadets.filter(c => c.gender === g.val || c.gender === g.val2);
      const gPs = gCadets.length;
      const gAbs = absences.filter(a => a.gender === g.val || a.gender === g.val2);
      const rowCounts = {};
      categories.forEach(cat => {
        rowCounts[cat] = gAbs.filter(a => a.category === cat).length;
      });

      ws.getCell(currRow, 2).value = g.numIdx;
      ws.getCell(currRow, 3).value = g.label;
      ws.getCell(currRow, 4).value = gPs;
      ws.getCell(currRow, 5).value = {
        formula: `D${currRow}-${v_col1_letter}${currRow}`,
        result: gPs - gAbs.length
      };

      categories.forEach((cat, c_offset) => {
        const colNum = first_cat_col + c_offset;
        const cVal = rowCounts[cat] > 0 ? rowCounts[cat] : null;
        ws.getCell(currRow, colNum).value = cVal;
      });

      ws.mergeCells(currRow, v_col1, currRow, v_col2);
      ws.getCell(currRow, v_col1).value = {
        formula: `SUM(${first_cat_letter}${currRow}:${last_cat_letter}${currRow})`,
        result: gAbs.length
      };

      applyStyleToRange(ws, currRow, 2, currRow, v_col2, {
        font: normFont,
        border: thinBorder,
        alignment: centerAlign
      });

      currRow++;
    });
  }

  // Строка «ИТОГО общий расход» с динамическими формулами
  ws.getRow(currRow).height = 18.0;
  const totalNumIdx = (hasMales && hasFemales) ? groups.length + 3 : groups.length + 1;
  ws.getCell(currRow, 2).value = totalNumIdx;
  ws.getCell(currRow, 3).value = "Всего";

  if (groups.length > 0) {
    const totalPsVal = cadets.length;
    const totalAbsVal = absences.length;
    const totalNlVal = totalPsVal - totalAbsVal;

    ws.getCell(currRow, 4).value = {
      formula: `SUM(D${first_grp_row}:D${last_grp_row})`,
      result: totalPsVal
    };
    ws.getCell(currRow, 5).value = {
      formula: `SUM(E${first_grp_row}:E${last_grp_row})`,
      result: totalNlVal
    };

    categories.forEach((cat, c_offset) => {
      const colLetter = getColumnLetter(first_cat_col + c_offset);
      const catCount = absences.filter(a => a.category === cat).length;
      ws.getCell(currRow, first_cat_col + c_offset).value = {
        formula: `SUM(${colLetter}${first_grp_row}:${colLetter}${last_grp_row})`,
        result: catCount
      };
    });

    ws.mergeCells(currRow, v_col1, currRow, v_col2);
    ws.getCell(currRow, v_col1).value = {
      formula: `SUM(${first_cat_letter}${currRow}:${last_cat_letter}${currRow})`,
      result: totalAbsVal
    };
  } else {
    ws.getCell(currRow, 4).value = 0;
    ws.getCell(currRow, 5).value = 0;
    ws.mergeCells(currRow, v_col1, currRow, v_col2);
    ws.getCell(currRow, v_col1).value = 0;
  }

  applyStyleToRange(ws, currRow, 2, currRow, v_col2, {
    font: boldFont,
    border: thinBorder,
    alignment: centerAlign
  });

  currRow += 2;

  // 3. Список отсутствующих
  const lastColLetter = v_col2_letter;
  ws.mergeCells(`B${currRow}:${lastColLetter}${currRow}`);
  ws.getCell(`B${currRow}`).value = "Отсутствующие";
  applyStyleToRange(ws, currRow, 2, currRow, v_col2, { font: normFont, alignment: centerAlign });
  ws.getRow(currRow).height = 18.0;
  currRow++;

  // Заголовки таблицы отсутствующих
  ws.getRow(currRow).height = 18.0;
  ws.getCell(currRow, 2).value = "№ п/п";
  ws.getCell(currRow, 3).value = "В/звание";

  ws.mergeCells(currRow, 4, currRow, 6);
  ws.getCell(currRow, 4).value = "Фамилия и инициалы";

  ws.mergeCells(currRow, 7, currRow, 8);
  ws.getCell(currRow, 7).value = "Причина отсутствия";

  ws.getCell(currRow, 9).value = "Подразд.";

  ws.mergeCells(currRow, 10, currRow, v_col2);
  ws.getCell(currRow, 10).value = "Примечание";

  applyStyleToRange(ws, currRow, 2, currRow, v_col2, {
    font: normFont,
    fill: headerFill,
    border: thinBorder,
    alignment: centerAlign
  });
  currRow++;

  // Строки отсутствующих
  absences.forEach((a, idx) => {
    ws.getRow(currRow).height = 18.0;

    const cNum = ws.getCell(currRow, 2);
    cNum.value = idx + 1;
    cNum.font = normFont;
    cNum.border = thinBorder;
    cNum.alignment = centerAlign;

    const cRank = ws.getCell(currRow, 3);
    cRank.value = a.rank || "курсант";
    cRank.font = normFont;
    cRank.border = thinBorder;
    cRank.alignment = centerAlign;

    ws.mergeCells(currRow, 4, currRow, 6);
    ws.getCell(currRow, 4).value = a.short_fio || a.fio || "";
    applyStyleToRange(ws, currRow, 4, currRow, 6, { font: normFont, border: thinBorder, alignment: centerAlign });

    ws.mergeCells(currRow, 7, currRow, 8);
    ws.getCell(currRow, 7).value = a.type_name || "";
    applyStyleToRange(ws, currRow, 7, currRow, 8, { font: normFont, border: thinBorder, alignment: centerAlign });

    const cGrp = ws.getCell(currRow, 9);
    cGrp.value = a.group_num || "";
    cGrp.font = normFont;
    cGrp.border = thinBorder;
    cGrp.alignment = centerAlign;

    ws.mergeCells(currRow, 10, currRow, v_col2);
    const noteText = String(a.note || "").trim();
    ws.getCell(currRow, 10).value = noteText;
    applyStyleToRange(ws, currRow, 10, currRow, v_col2, { font: normFont, border: thinBorder, alignment: centerAlign });

    currRow++;
  });

  // 4. Подписи начальника
  currRow++;
  ws.mergeCells(currRow, 2, currRow, v_col2);
  ws.getCell(currRow, 2).value = state.settings.chiefTitle.trim();
  applyStyleToRange(ws, currRow, 2, currRow, v_col2, { font: normFont, alignment: centerAlign });

  currRow++;
  ws.mergeCells(currRow, 2, currRow, v_col2);
  ws.getCell(currRow, 2).value = state.settings.chiefName;
  applyStyleToRange(ws, currRow, 2, currRow, v_col2, { font: normFont, alignment: centerAlign });

  // 5. Ширина колонок
  const colWidths = {
    1: 3.0,   // A
    2: 7.0,   // B
    3: 11.5,  // C
    4: 8.5,   // D
    5: 9.0,   // E
    6: 13.0,  // F
    7: 12.5,  // G
    8: 11.5,  // H
    9: 15.0,  // I
    10: 15.0, // J
    11: 9.0,  // K
    12: 13.0  // L
  };
  Object.entries(colWidths).forEach(([colIdx, width]) => {
    ws.getColumn(Number(colIdx)).width = width;
  });

  // 6. Параметры страницы для печати на лист А4 (1-в-1 с openpyxl)
  ws.pageSetup = {
    orientation: "portrait",
    paperSize: 9, // A4
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 0,
    horizontalCentered: true,
    margins: {
      left: 0.2,
      right: 0.2,
      top: 0.3,
      bottom: 0.3,
      header: 0.0,
      footer: 0.0
    },
    printArea: `B2:${lastColLetter}${currRow}`
  };
}

async function generateExcelReport() {
  if (state.cadets.length === 0) {
    showToast("Сначала загрузите базу личного состава!", "danger");
    return;
  }

  const wb = new ExcelJS.Workbook();
  wb.creator = "ConsumABLE Web";
  wb.created = new Date();

  // Определение категорий для сводки
  const standardCategories = ["Наряд", "Госпиталь", "Лазарет", "Командировка", "Похороны"];
  const activeAbsenceCategories = [...new Set(state.absences.map(a => a.category))];
  const orderedCats = standardCategories.filter(c => activeAbsenceCategories.includes(c) || true);
  const otherCats = activeAbsenceCategories.filter(c => !standardCategories.includes(c));
  const categories = [...orderedCats, ...otherCats];

  // Лист 1: «Сводный расход»
  const wsTotal = wb.addWorksheet("Сводный расход");
  buildExcelSheet(wsTotal, "", state.cadets, state.absences, categories);

  // Лист 2: «Мужской состав»
  const maleCadets = state.cadets.filter(c => c.gender === 'М' || c.gender === 'male');
  const maleAbsences = state.absences.filter(a => a.gender === 'М' || a.gender === 'male');
  const wsMale = wb.addWorksheet("Мужской состав");
  buildExcelSheet(wsMale, "(Мужской состав)", maleCadets, maleAbsences, categories);

  // Лист 3: «Женский состав»
  const femaleCadets = state.cadets.filter(c => c.gender === 'Ж' || c.gender === 'female');
  const femaleAbsences = state.absences.filter(a => a.gender === 'Ж' || a.gender === 'female');
  const wsFemale = wb.addWorksheet("Женский состав");
  buildExcelSheet(wsFemale, "(Женский состав)", femaleCadets, femaleAbsences, categories);

  // Выгрузка файла через нативный Blob
  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  const fileNameDate = state.settings.reportDate || "сегодня";
  a.download = `Строевая_записка_${fileNameDate}.xlsx`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1000);

  showToast("Строевая записка (3 листа) успешно сгенерирована и скачана!", "success");
}

// 5. УПРАВЛЕНИЕ ИНТЕРФЕЙСОМ И РЕНДЕРИНГ

function switchSliceTab(slice) {
  state.currentSlice = slice;
  document.querySelectorAll(".slice-tab-btn").forEach(btn => {
    btn.classList.toggle("active", btn.dataset.slice === slice);
  });
  updateKpi();
  renderAbsencesTable();
}

function updateKpi() {
  const allCadets = state.cadets;
  const allAbsences = state.absences;

  const maleCadets = allCadets.filter(c => c.gender === 'М' || c.gender === 'male');
  const femaleCadets = allCadets.filter(c => c.gender === 'Ж' || c.gender === 'female');

  const maleAbsences = allAbsences.filter(a => a.gender === 'М' || a.gender === 'male');
  const femaleAbsences = allAbsences.filter(a => a.gender === 'Ж' || a.gender === 'female');

  const totalTabCount = document.getElementById("sliceTotalCount");
  const maleTabCount = document.getElementById("sliceMaleCount");
  const femaleTabCount = document.getElementById("sliceFemaleCount");
  if (totalTabCount) totalTabCount.textContent = allCadets.length;
  if (maleTabCount) maleTabCount.textContent = maleCadets.length;
  if (femaleTabCount) femaleTabCount.textContent = femaleCadets.length;

  let currentCadets = allCadets;
  let currentAbsences = allAbsences;
  if (state.currentSlice === 'male') {
    currentCadets = maleCadets;
    currentAbsences = maleAbsences;
  } else if (state.currentSlice === 'female') {
    currentCadets = femaleCadets;
    currentAbsences = femaleAbsences;
  }

  const total = currentCadets.length;
  const absent = currentAbsences.length;
  const present = Math.max(0, total - absent);

  const kpiTotalEl = document.getElementById("kpiTotal");
  const kpiPresentEl = document.getElementById("kpiPresent");
  const kpiAbsentEl = document.getElementById("kpiAbsent");

  if (kpiTotalEl) kpiTotalEl.textContent = `${total} чел.`;
  if (kpiPresentEl) kpiPresentEl.textContent = `${present} чел.`;
  if (kpiAbsentEl) kpiAbsentEl.textContent = `${absent} чел.`;

  // 3 Информационные карточки со статистикой (Всего / Мужчины / Девушки)
  const statTotal = document.getElementById("statCardTotal");
  const statMale = document.getElementById("statCardMale");
  const statFemale = document.getElementById("statCardFemale");
  if (statTotal) {
    const ps = allCadets.length;
    const abs = allAbsences.length;
    const nl = Math.max(0, ps - abs);
    statTotal.textContent = `По списку: ${ps} | В строю: ${nl} | На расходе: ${abs}`;
  }
  if (statMale) {
    const ps = maleCadets.length;
    const abs = maleAbsences.length;
    const nl = Math.max(0, ps - abs);
    statMale.textContent = `По списку: ${ps} | В строю: ${nl} | На расходе: ${abs}`;
  }
  if (statFemale) {
    const ps = femaleCadets.length;
    const abs = femaleAbsences.length;
    const nl = Math.max(0, ps - abs);
    statFemale.textContent = `По списку: ${ps} | В строю: ${nl} | На расходе: ${abs}`;
  }

  // Контроль сквозного математического баланса
  const balanceBanner = document.getElementById("sliceBalanceBanner");
  if (balanceBanner) {
    const isSumBalanced = (allCadets.length === (maleCadets.length + femaleCadets.length)) &&
                          (allAbsences.length === (maleAbsences.length + femaleAbsences.length));
    if (isSumBalanced) {
      balanceBanner.className = "slice-balance-banner";
      balanceBanner.innerHTML = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" class="icon-sm" style="display:inline-block; vertical-align:middle; margin-right:4px;"><path d="M12 3v18"/><path d="M3 7l9-4 9 4"/><circle cx="6" cy="14" r="3"/><circle cx="18" cy="14" r="3"/></svg> <strong>Сквозной баланс строя подтвержден:</strong> Общий [П/С: ${allCadets.length}, В строю: ${Math.max(0, allCadets.length - allAbsences.length)}, Расход: ${allAbsences.length}] = Мужчины (${maleCadets.length}/${Math.max(0, maleCadets.length - maleAbsences.length)}/${maleAbsences.length}) + Женщины (${femaleCadets.length}/${Math.max(0, femaleCadets.length - femaleAbsences.length)}/${femaleAbsences.length})`;
    } else {
      balanceBanner.className = "slice-balance-banner error";
      balanceBanner.innerHTML = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" class="icon-sm" style="display:inline-block; vertical-align:middle; margin-right:4px;"><circle cx="12" cy="12" r="10"/><line x1="15" y1="9" x2="9" y2="15"/><line x1="9" y1="9" x2="15" y2="15"/></svg> <strong>Нарушение баланса:</strong> Сумма мужского и женского состава не сходится с общим!`;
    }
  }
}

function calculateAndRenderBalance() {
  updateKpi();
}

function updateSliceStats() {
  updateKpi();
}

function getAvailableCadets(query = "") {
  const absentCadetIds = new Set(state.absences.map(a => a.cadet_id));
  const filterText = String(query || "").trim().toLowerCase();
  return state.cadets.filter(c => {
    if (absentCadetIds.has(c.id)) return false;
    if (!filterText) return true;
    const fio = String(c.fio || "").toLowerCase();
    const shortFio = String(c.short_fio || "").toLowerCase();
    const rank = String(c.rank || "").toLowerCase();
    const grp = String(c.group_num || "").toLowerCase();
    return fio.includes(filterText) ||
           shortFio.includes(filterText) ||
           rank.includes(filterText) ||
           grp.includes(filterText);
  });
}

function detectReasonCategory(reasonName) {
  if (!reasonName) return "Наряд";
  const low = reasonName.trim().toLowerCase();

  const matched = state.absenceTypes.find(t => t.name.toLowerCase() === low);
  if (matched) return matched.category;

  if (low.includes("увольн") || low.includes("отпуск")) return "Увольнение";
  if (low.includes("госпитал")) return "Госпиталь";
  if (low.includes("лазарет") || low.includes("больнич") || low.includes("медпункт")) return "Лазарет";
  if (low.includes("командир")) return "Командировка";
  if (low.includes("похорон")) return "Похороны";
  return "Наряд";
}

function getCategoryBadgeClass(category) {
  if (category === "Увольнение") return "badge-green";
  if (category === "Госпиталь" || category === "Лазарет") return "badge-red";
  if (category === "Командировка") return "badge-purple";
  if (category === "Похороны") return "badge-yellow";
  return "badge-blue";
}

function updateCadetSelectOptions() {
  const cadetDropdown = document.getElementById("cadetAutocompleteDropdown");
  if (cadetDropdown) cadetDropdown.classList.remove("active");
}

function renderAbsencesTable() {
  const tbody = document.getElementById("absencesTableBody");
  const emptyState = document.getElementById("absencesEmptyState");
  const table = document.getElementById("absencesTable");
  const badgeCount = document.getElementById("absenceCountBadge");

  if (!tbody) return;

  let displayedAbsences = state.absences;
  if (state.currentSlice === 'male') {
    displayedAbsences = state.absences.filter(a => a.gender === 'М' || a.gender === 'male');
  } else if (state.currentSlice === 'female') {
    displayedAbsences = state.absences.filter(a => a.gender === 'Ж' || a.gender === 'female');
  }

  if (badgeCount) {
    badgeCount.textContent = `${displayedAbsences.length} чел.`;
  }

  if (displayedAbsences.length === 0) {
    if (table) table.style.display = "none";
    if (emptyState) emptyState.style.display = "block";
    tbody.innerHTML = "";
    return;
  }

  if (table) table.style.display = "table";
  if (emptyState) emptyState.style.display = "none";

  tbody.innerHTML = "";

  displayedAbsences.forEach((a, idx) => {
    const tr = document.createElement("tr");

    let catBadgeClass = "badge-blue";
    if (a.category === "Увольнение") catBadgeClass = "badge-green";
    else if (a.category === "Госпиталь" || a.category === "Лазарет") catBadgeClass = "badge-red";
    else if (a.category === "Командировка") catBadgeClass = "badge-purple";
    else if (a.category === "Похороны") catBadgeClass = "badge-yellow";

    tr.innerHTML = `
      <td class="col-center">${idx + 1}</td>
      <td>${escapeHtml(a.rank)}</td>
      <td><strong>${escapeHtml(a.short_fio)}</strong></td>
      <td class="col-center"><span class="badge badge-gray">${escapeHtml(a.group_num)}</span></td>
      <td>${escapeHtml(a.type_name)}</td>
      <td><span class="badge ${catBadgeClass}">${escapeHtml(a.category)}</span></td>
      <td>${escapeHtml(a.note || "—")}</td>
      <td class="col-center">
        <button class="btn btn-danger btn-sm" onclick="removeAbsence(${a.id})" title="Удалить из расхода">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>
        </button>
      </td>
    `;
    tbody.appendChild(tr);
  });
}

// -------------------------------------------------------------
// МОДАЛЬНОЕ ОКНО ВЫБОРА КУРСА ПРИ ИМПОРТЕ EXCEL
// -------------------------------------------------------------

let pendingImportCadets = [];
let selectedImportCourse = 3;

function openImportCourseModal(cadets) {
  pendingImportCadets = cadets || [];
  const modal = document.getElementById("importCourseModal");
  if (!modal) return;

  let detectedCourse = 2;
  const courseCounts = {};
  cadets.forEach(c => {
    const cVal = getCourseFromGroup(c.group_num);
    if (cVal !== '?' && !isNaN(parseInt(cVal, 10))) {
      courseCounts[cVal] = (courseCounts[cVal] || 0) + 1;
    }
  });

  let maxCount = 0;
  Object.entries(courseCounts).forEach(([crs, cnt]) => {
    if (cnt > maxCount) {
      maxCount = cnt;
      detectedCourse = parseInt(crs, 10);
    }
  });

  const promptEl = document.getElementById("importCourseModalPrompt");
  if (promptEl) {
    if (maxCount > 0) {
      promptEl.textContent = `Обнаружен ${detectedCourse} курс в номерах групп. Установить актуальный курс:`;
    } else {
      promptEl.textContent = `Какой курс установить для загружаемых групп?`;
    }
  }

  selectImportCourse(detectedCourse);
  modal.classList.add("active");
}

function closeImportCourseModal() {
  const modal = document.getElementById("importCourseModal");
  if (modal) modal.classList.remove("active");
  pendingImportCadets = [];
}

function selectImportCourse(course) {
  selectedImportCourse = course;
  const grid = document.querySelector(".course-choice-grid");
  if (grid) {
    grid.querySelectorAll(".course-choice-btn").forEach(btn => {
      const bCourse = btn.dataset.course;
      if (String(bCourse) === String(course)) {
        btn.classList.add("selected");
      } else {
        btn.classList.remove("selected");
      }
    });
  }

  const hint = document.getElementById("importCourseHint");
  if (hint) {
    if (course === 'keep') {
      hint.textContent = "Выбран режим: Оставить как в файле (номера групп не меняются)";
    } else {
      let sample = "1280";
      if (pendingImportCadets && pendingImportCadets.length > 0 && pendingImportCadets[0].group_num) {
        sample = pendingImportCadets[0].group_num;
      }
      const changed = changeGroupCourse(sample, course);
      hint.textContent = `Выбран: ${course} курс (группа ${sample} станет ${changed})`;
    }
  }
}

function confirmImportWithCourse() {
  if (!pendingImportCadets || pendingImportCadets.length === 0) {
    closeImportCourseModal();
    return;
  }

  const targetCourse = selectedImportCourse;
  state.cadets = pendingImportCadets.map((c, idx) => {
    let grp = c.group_num || "";
    let crs = c.course || 1;

    if (targetCourse !== 'keep') {
      const numCourse = parseInt(targetCourse, 10);
      grp = changeGroupCourse(grp, numCourse);
      crs = numCourse;
    } else {
      const detected = getCourseFromGroup(grp);
      if (detected !== "?") {
        crs = parseInt(detected, 10);
      }
    }

    return {
      id: nextCadetId++,
      rank: c.rank || "курсант",
      fio: c.fio,
      short_fio: c.short_fio || makeShortFio(c.fio),
      group_num: grp,
      course: crs,
      gender: (c.gender === 'female' || c.gender === 'Ж') ? 'Ж' : detectGender(c.fio, c.gender),
      status: 'active',
      sort_order: idx
    };
  });

  state.absences = []; // сброс расхода для нового списка личного состава

  // Автоподстановка курса и факультета в шапку строевой записки
  let detectedFaculty = null;
  let detectedCourse = targetCourse !== 'keep' ? parseInt(targetCourse, 10) : null;
  for (const c of state.cadets) {
    const parsed = parseGroupCode(c.group_num);
    if (parsed.faculty && !detectedFaculty) detectedFaculty = parsed.faculty;
    if (parsed.course && !detectedCourse) detectedCourse = parsed.course;
    if (detectedFaculty && detectedCourse) break;
  }
  if (detectedFaculty || detectedCourse) {
    updateHeaderSettings(detectedFaculty, detectedCourse);
  }

  closeImportCourseModal();

  updateKpi();
  updateCadetGroupFilterOptions();
  renderCadetsTable();
  updateCadetSelectOptions();
  renderAbsencesTable();
  calculateAndRenderBalance();
  renderLivePreview();

  showToast(`Успешно загружено курсантов: ${state.cadets.length} чел.${targetCourse !== 'keep' ? ` (${targetCourse} курс)` : ''}`, "success");
}

function incrementAllCourses() {
  if (state.cadets.length === 0) {
    showToast("База личного состава пуста!", "danger");
    return;
  }

  state.cadets = state.cadets.map(c => {
    const nextCourse = (c.course || 1) + 1;
    const nextGroup = changeGroupCourse(c.group_num, nextCourse);
    return {
      ...c,
      course: nextCourse,
      group_num: nextGroup,
      status: nextCourse > 5 ? 'graduated' : c.status
    };
  });

  state.absences = state.absences.map(a => {
    const nextCourse = (a.course || 1) + 1;
    const nextGroup = changeGroupCourse(a.group_num, nextCourse);
    return {
      ...a,
      course: nextCourse,
      group_num: nextGroup
    };
  });

  // Автоматически обновляем номер курса в шапке строевой записки
  let newCourse = null;
  if (state.cadets.length > 0) {
    newCourse = state.cadets[0].course;
  }
  if (newCourse !== null) {
    updateHeaderSettings(null, newCourse);
  }

  updateKpi();
  updateCadetGroupFilterOptions();
  renderCadetsTable();
  updateCadetSelectOptions();
  renderAbsencesTable();
  calculateAndRenderBalance();
  renderLivePreview();

  showToast("Курс всех учебных групп успешно повышен на +1!", "success");
}

window.openImportCourseModal = openImportCourseModal;
window.closeImportCourseModal = closeImportCourseModal;
window.selectImportCourse = selectImportCourse;
window.confirmImportWithCourse = confirmImportWithCourse;
window.incrementAllCourses = incrementAllCourses;

// -------------------------------------------------------------
// STAGING WIZARD ДЛЯ WEB-ВЕРСИИ
// -------------------------------------------------------------

function openWizardModal(stagingCadets) {
  state.stagingCadets = stagingCadets.map((c, i) => ({
    tempId: 't_' + (i + 1),
    course: c.course || (getCourseFromGroup(c.group_num) !== '?' ? Number(getCourseFromGroup(c.group_num)) : 1),
    status: (c.course > 5) ? 'graduated' : 'active',
    ...c
  }));
  renderWizardTable();
  const modal = document.getElementById("stagingWizardModal");
  if (modal) modal.classList.add("active");
}

function closeWizardModal() {
  const modal = document.getElementById("stagingWizardModal");
  if (modal) modal.classList.remove("active");
}

function parseGroupCode(groupCode) {
  if (!groupCode) return { faculty: null, course: null, subgroup: null };
  const grpStr = String(groupCode).trim();
  const m4 = grpStr.match(/^(\d)(\d)(\d{2})$/) || grpStr.match(/(\d)(\d)(\d{2})/);
  if (m4) {
    return {
      faculty: parseInt(m4[1], 10),
      course: parseInt(m4[2], 10),
      subgroup: m4[3]
    };
  }
  const m3 = grpStr.match(/^(\d)(\d{2})$/);
  if (m3) {
    return {
      faculty: null,
      course: parseInt(m3[1], 10),
      subgroup: m3[2]
    };
  }
  return { faculty: null, course: null, subgroup: null };
}

function getCourseFromGroup(groupStr) {
  const cleaned = String(groupStr || "").replace(/\D/g, "");
  if (cleaned.length === 4) {
    return cleaned[1]; // Вторая цифра — курс
  }
  return "?";
}

function changeGroupCourse(groupCode, targetCourse) {
  if (!groupCode) return "";
  const grpStr = String(groupCode).trim();
  const targetDigit = String(parseInt(targetCourse, 10));

  // 4-digit code e.g. 1280 -> 1380
  const m4 = grpStr.match(/^(\d)(\d)(\d{2})$/);
  if (m4) {
    return `${m4[1]}${targetDigit}${m4[3]}`;
  }
  const m4InText = grpStr.match(/(\b\d)\d(\d{2}\b)/);
  if (m4InText) {
    return grpStr.replace(/(\b\d)\d(\d{2}\b)/, `$1${targetDigit}$2`);
  }
  // 3-digit code e.g. 221 -> 321
  const m3 = grpStr.match(/^(\d)(\d{2})$/);
  if (m3) {
    return `${targetDigit}${m3[2]}`;
  }
  return grpStr;
}

function openBulkCourseModal() {
  const modal = document.getElementById("bulkCourseModal");
  if (modal) modal.classList.add("active");
}

function closeBulkCourseModal() {
  const modal = document.getElementById("bulkCourseModal");
  if (modal) modal.classList.remove("active");
}

function applyBulkCourseChange() {
  const select = document.getElementById("bulkTargetCourseSelect");
  if (!select) return;
  const targetCourse = parseInt(select.value, 10);
  if (!targetCourse || targetCourse < 1 || targetCourse > 5) return;

  state.cadets.forEach(c => {
    c.course = targetCourse;
    if (c.group_num) {
      c.group_num = changeGroupCourse(c.group_num, targetCourse);
    }
  });

  state.absences.forEach(a => {
    if (a.group_num) {
      a.group_num = changeGroupCourse(a.group_num, targetCourse);
    }
  });

  // Автоматически обновляем номер курса в шапке строевой записки
  updateHeaderSettings(null, targetCourse);

  renderCadetsTable();
  updateCadetGroupFilterOptions();
  renderAbsencesTable();
  updateKpi();
  updateSliceStats();
  calculateAndRenderBalance();
  renderLivePreview();
  closeBulkCourseModal();
  showToast(`Курс всех групп успешно изменен на ${targetCourse} курс!`, "success");
}

function renderWizardTable() {
  const tbody = document.getElementById("wizardTableBody");
  const badge = document.getElementById("wizardSummaryBadge");
  if (!tbody) return;

  const total = state.stagingCadets.length;
  const graduated = state.stagingCadets.filter(c => c.status === 'graduated' || c.course > 5).length;
  const males = state.stagingCadets.filter(c => c.gender === 'М' || c.gender === 'male').length;
  const females = state.stagingCadets.filter(c => c.gender === 'Ж' || c.gender === 'female').length;

  if (badge) {
    badge.textContent = `Всего: ${total} чел. | М: ${males} | Ж: ${females}` + (graduated ? ` | Выпускники: ${graduated}` : '');
  }

  tbody.innerHTML = "";
  state.stagingCadets.forEach((c, idx) => {
    const tr = document.createElement("tr");
    const isGraduated = c.status === 'graduated' || c.course > 5;
    const isFemale = c.gender === 'Ж' || c.gender === 'female';
    if (isGraduated) tr.style.backgroundColor = '#FFFBEB';

    tr.innerHTML = `
      <td style="text-align: center; color: var(--text-muted);">${idx + 1}</td>
      <td>${escapeHtml(c.rank || "курсант")}</td>
      <td><strong>${escapeHtml(c.fio)}</strong></td>
      <td style="text-align: center;"><span class="badge badge-gray">${escapeHtml(c.group_num)}</span></td>
      <td style="text-align: center;">
        <select class="form-select" style="padding: 2px 6px; font-size: 12px; width: auto;" onchange="wizardChangeCourse('${c.tempId}', this.value)">
          ${[1, 2, 3, 4, 5, 6].map(cr => `<option value="${cr}" ${c.course == cr ? 'selected' : ''}>${cr} курс</option>`).join('')}
        </select>
      </td>
      <td style="text-align: center;">
        <button class="btn btn-sm" style="padding: 2px 8px; font-weight: 700; ${isFemale ? 'background: #FDF2F8; color: #BE185D; border: 1px solid #FBCFE8;' : 'background: #EFF6FF; color: #1D4ED8; border: 1px solid #BFDBFE;'}" onclick="wizardToggleGender('${c.tempId}')">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" class="icon-sm" style="display:inline-block; vertical-align:middle; margin-right:2px;">
            ${isFemale ? '<circle cx="12" cy="9" r="5"/><line x1="12" y1="14" x2="12" y2="21"/><line x1="9" y1="18" x2="15" y2="18"/>' : '<circle cx="10" cy="14" r="5"/><line x1="19" y1="5" x2="13.6" y2="10.4"/><polyline points="15 5 19 5 19 9"/>'}
          </svg>
          ${isFemale ? 'Ж' : 'М'}
        </button>
      </td>
      <td style="text-align: center;">
        <span class="badge ${isGraduated ? 'badge-yellow' : 'badge-green'}">${isGraduated ? 'Выпускник' : 'В строю'}</span>
      </td>
      <td>
        ${isGraduated ? '<span style="color: #DC2626; font-size: 11px; font-weight: 600;">Курс > 5 (выпускник)</span>' : '<span style="color: #16A34A; font-size: 11px; font-weight: 600;">Норма</span>'}
      </td>
    `;
    tbody.appendChild(tr);
  });
}

function wizardIncrementCourseAll() {
  state.stagingCadets = state.stagingCadets.map(c => {
    if (c.status === 'active' || !c.status) {
      const nextCourse = Number(c.course || 1) + 1;
      return {
        ...c,
        course: nextCourse,
        group_num: nextCourse <= 5 ? changeGroupCourse(c.group_num, nextCourse) : c.group_num,
        status: nextCourse > 5 ? 'graduated' : 'active'
      };
    }
    return c;
  });
  renderWizardTable();
  showToast("Курс всех активных курсантов увеличен на +1. Номера групп обновлены.", "info");
}

function wizardBulkSetCourse() {
  const sel = document.getElementById("wizardBulkCourseSelect");
  const targetCourse = Number(sel ? sel.value : 1);
  state.stagingCadets = state.stagingCadets.map(c => ({
    ...c,
    course: targetCourse,
    group_num: changeGroupCourse(c.group_num, targetCourse),
    status: targetCourse > 5 ? 'graduated' : 'active'
  }));
  renderWizardTable();
  showToast(`Курс ${targetCourse} назначен. Номера групп обновлены (2-я цифра).`, "info");
}

function wizardRemoveGraduates() {
  const countBefore = state.stagingCadets.length;
  state.stagingCadets = state.stagingCadets.filter(c => c.status !== 'graduated' && c.course <= 5);
  renderWizardTable();
  showToast(`Исключено выпускников: ${countBefore - state.stagingCadets.length} чел.`, "info");
}

function wizardToggleGender(tempId) {
  state.stagingCadets = state.stagingCadets.map(c => {
    if (c.tempId === tempId) {
      const newGender = (c.gender === 'Ж' || c.gender === 'female') ? 'М' : 'Ж';
      return { ...c, gender: newGender };
    }
    return c;
  });
  renderWizardTable();
}

function wizardChangeCourse(tempId, val) {
  const num = Number(val);
  state.stagingCadets = state.stagingCadets.map(c => {
    if (c.tempId === tempId) {
      return {
        ...c,
        course: num,
        status: num > 5 ? 'graduated' : 'active'
      };
    }
    return c;
  });
  renderWizardTable();
}

function confirmWizardImport() {
  const graduates = state.stagingCadets.filter(c => c.status === 'graduated' || c.course > 5);
  if (graduates.length > 0) {
    if (confirm(`В списке обнаружено ${graduates.length} выпускников (курс > 5). Исключить их перед внесением в расход?`)) {
      state.stagingCadets = state.stagingCadets.filter(c => c.status !== 'graduated' && c.course <= 5);
    }
  }

  state.cadets = state.stagingCadets.map((c, idx) => ({
    id: nextCadetId++,
    rank: c.rank || "курсант",
    fio: c.fio,
    short_fio: c.short_fio || makeShortFio(c.fio),
    group_num: c.group_num,
    course: c.course || (getCourseFromGroup(c.group_num) !== '?' ? Number(getCourseFromGroup(c.group_num)) : 1),
    gender: (c.gender === 'female' || c.gender === 'Ж') ? 'Ж' : 'М',
    status: c.status || 'active',
    sort_order: idx
  }));

  state.absences = [];

  // Автоподстановка курса и факультета в шапку строевой записки
  let detectedFaculty = null;
  let detectedCourse = null;
  for (const c of state.cadets) {
    const parsed = parseGroupCode(c.group_num);
    if (parsed.faculty && !detectedFaculty) detectedFaculty = parsed.faculty;
    if (parsed.course && !detectedCourse) detectedCourse = parsed.course;
    if (detectedFaculty && detectedCourse) break;
  }
  if (detectedFaculty || detectedCourse) {
    updateHeaderSettings(detectedFaculty, detectedCourse);
  }

  closeWizardModal();

  updateKpi();
  updateCadetGroupFilterOptions();
  renderCadetsTable();
  updateCadetSelectOptions();
  renderAbsencesTable();
  renderLivePreview();

  showToast(`Успешно загружено курсантов: ${state.cadets.length} чел.`, "success");
}

function removeAbsence(absenceId) {
  state.absences = state.absences.filter(a => a.id !== absenceId);
  updateKpi();
  renderAbsencesTable();
  updateCadetSelectOptions();
  renderLivePreview();
  showToast("Курсант возвращен в строй", "info");
}

function updateCadetGroupFilterOptions() {
  const filterGroupSelect = document.getElementById("cadetGroupFilter");
  if (!filterGroupSelect) return;

  const currentFilter = filterGroupSelect.value;
  const groups = [...new Set(state.cadets.map(c => String(c.group_num || "").trim()))].filter(Boolean).sort((a, b) => {
    const na = Number(a), nb = Number(b);
    if (!isNaN(na) && !isNaN(nb)) return na - nb;
    return a.localeCompare(b);
  });

  filterGroupSelect.innerHTML = '<option value="">Все подразделения</option>';
  groups.forEach(g => {
    const opt = document.createElement("option");
    opt.value = g;
    opt.textContent = `${g} взвод / группа`;
    filterGroupSelect.appendChild(opt);
  });

  if (groups.includes(currentFilter)) {
    filterGroupSelect.value = currentFilter;
  }
}

function renderCadetsTable() {
  const tbody = document.getElementById("cadetsTableBody");
  const emptyState = document.getElementById("cadetsEmptyState");
  const table = document.getElementById("cadetsTable");
  const countEl = document.getElementById("cadetsCountBadge");
  const filterGroupSelect = document.getElementById("cadetGroupFilter");
  const searchInput = document.getElementById("cadetSearchInput");

  if (!tbody) return;

  const searchText = searchInput ? searchInput.value.trim().toLowerCase() : "";
  const selectedGroup = filterGroupSelect ? filterGroupSelect.value.trim() : "";

  const filtered = state.cadets.filter(c => {
    if (state.selectedCourseFilter && state.selectedCourseFilter !== 'all') {
      if (Number(c.course) !== Number(state.selectedCourseFilter)) {
        return false;
      }
    }
    if (selectedGroup && String(c.group_num || "").trim() !== selectedGroup) {
      return false;
    }
    if (searchText) {
      const fio = String(c.fio || "").toLowerCase();
      const shortFio = String(c.short_fio || "").toLowerCase();
      const rank = String(c.rank || "").toLowerCase();
      const grp = String(c.group_num || "").toLowerCase();
      return fio.includes(searchText) ||
             shortFio.includes(searchText) ||
             rank.includes(searchText) ||
             grp.includes(searchText);
    }
    return true;
  });

  if (countEl) {
    if (searchText || selectedGroup || (state.selectedCourseFilter && state.selectedCourseFilter !== 'all')) {
      countEl.textContent = `${filtered.length} из ${state.cadets.length} чел.`;
    } else {
      countEl.textContent = `${state.cadets.length} чел.`;
    }
  }

  if (filtered.length === 0) {
    if (table) table.style.display = "none";
    if (emptyState) emptyState.style.display = "block";
    tbody.innerHTML = "";
    return;
  }

  if (table) table.style.display = "table";
  if (emptyState) emptyState.style.display = "none";

  tbody.innerHTML = "";

  filtered.forEach((c, idx) => {
    const isFemale = (c.gender === 'Ж' || c.gender === 'female');
    const crs = getCourseFromGroup(c.group_num);
    const crsText = crs !== "?" ? ` (${crs} к.)` : (c.course ? ` (${c.course} к.)` : "");
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td class="col-center">${idx + 1}</td>
      <td>${escapeHtml(c.rank)}</td>
      <td><strong>${escapeHtml(c.fio)}</strong></td>
      <td>${escapeHtml(c.short_fio)}</td>
      <td class="col-center"><span class="badge badge-gray">${escapeHtml(c.group_num)}${crsText}</span></td>
      <td class="col-center">
        <button class="btn btn-sm" style="padding: 2px 8px; font-weight: 700; ${isFemale ? 'background: #FDF2F8; color: #BE185D; border: 1px solid #FBCFE8;' : 'background: #EFF6FF; color: #1D4ED8; border: 1px solid #BFDBFE;'}" onclick="toggleCadetGender(${c.id})" title="Нажмите для смены пола (М ↔ Ж)">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" class="icon-sm" style="display:inline-block; vertical-align:middle; margin-right:2px;">
            ${isFemale ? '<circle cx="12" cy="9" r="5"/><line x1="12" y1="14" x2="12" y2="21"/><line x1="9" y1="18" x2="15" y2="18"/>' : '<circle cx="10" cy="14" r="5"/><line x1="19" y1="5" x2="13.6" y2="10.4"/><polyline points="15 5 19 5 19 9"/>'}
          </svg>
          ${isFemale ? 'Ж' : 'М'}
        </button>
      </td>
      <td class="col-center" style="display: flex; gap: 6px; justify-content: center;">
        <button class="btn btn-secondary btn-sm" onclick="openEditCadetModal(${c.id})" title="Редактировать">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
        </button>
        <button class="btn btn-danger btn-sm" onclick="deleteCadet(${c.id})" title="Удалить">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>
        </button>
      </td>
    `;
    tbody.appendChild(tr);
  });
}

function toggleCadetGender(cadetId) {
  const cadet = state.cadets.find(c => c.id === cadetId);
  if (!cadet) return;
  cadet.gender = (cadet.gender === 'Ж' ? 'М' : 'Ж');
  state.absences.forEach(a => {
    if (a.cadet_id === cadetId) {
      a.gender = cadet.gender;
    }
  });
  renderCadetsTable();
  renderAbsencesTable();
  calculateAndRenderBalance();
  updateKpi();
  renderLivePreview();
  showToast(`Пол курсанта ${cadet.short_fio} изменен на ${cadet.gender}`, "info");
}

function filterCadetsByCourse(courseVal) {
  state.selectedCourseFilter = courseVal;
  document.querySelectorAll(".course-filter-bar .seg-btn").forEach(btn => {
    if (btn.getAttribute("data-course") === String(courseVal)) {
      btn.classList.add("active");
    } else {
      btn.classList.remove("active");
    }
  });
  renderCadetsTable();
}

function deleteCadet(cadetId) {
  if (!confirm("Удалить курсанта из базы? Связанный расход также будет удален.")) {
    return;
  }
  state.cadets = state.cadets.filter(c => c.id !== cadetId);
  state.absences = state.absences.filter(a => a.cadet_id !== cadetId);
  updateKpi();
  updateCadetGroupFilterOptions();
  renderCadetsTable();
  updateCadetSelectOptions();
  renderAbsencesTable();
  renderLivePreview();
  showToast("Курсант удален из базы", "info");
}

function renderReasonsTable() {
  const tbody = document.getElementById("reasonsTableBody");
  if (!tbody) return;

  tbody.innerHTML = "";
  state.absenceTypes.forEach((t, idx) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td class="col-center">${idx + 1}</td>
      <td><strong>${escapeHtml(t.name)}</strong></td>
      <td class="col-center"><span class="badge badge-gray">${escapeHtml(t.code || "—")}</span></td>
      <td><span class="badge badge-blue">${escapeHtml(t.category)}</span></td>
      <td class="col-center">
        ${t.id > 7 ? `
          <button class="btn btn-danger btn-sm" onclick="deleteReasonType(${t.id})" title="Удалить причину">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>
          </button>
        ` : `<span style="color: var(--text-light); font-size: 11px;">Базовая</span>`}
      </td>
    `;
    tbody.appendChild(tr);
  });

  // Обновление выпадающего списка причин в форме добавления
  const select = document.getElementById("absenceReasonSelect");
  if (select) {
    const currentVal = select.value;
    select.innerHTML = '<option value="">-- Выберите причину --</option>';
    state.absenceTypes.forEach(t => {
      const opt = document.createElement("option");
      opt.value = t.id;
      opt.textContent = `${t.name} (${t.category})`;
      select.appendChild(opt);
    });
    const customOpt = document.createElement("option");
    customOpt.value = "custom";
    customOpt.textContent = "+ Другая причина (вручную)";
    select.appendChild(customOpt);

    if (currentVal) select.value = currentVal;
  }
}

function deleteReasonType(typeId) {
  state.absenceTypes = state.absenceTypes.filter(t => t.id !== typeId);
  renderReasonsTable();
  showToast("Причина удалена", "info");
}

function switchPreviewSlice(slice) {
  state.activePreviewSlice = slice;
  document.querySelectorAll(".preview-tab-btn").forEach(btn => {
    btn.classList.toggle("active", btn.dataset.slice === slice);
  });
  renderLivePreview(slice);
}

function printActiveReport() {
  if (state.cadets.length === 0) {
    showToast("Сначала загрузите базу личного состава!", "danger");
    return;
  }
  window.print();
}

window.switchPreviewSlice = switchPreviewSlice;
window.printActiveReport = printActiveReport;
window.generateExcelReport = generateExcelReport;

function renderLivePreview(slice = state.activePreviewSlice || 'total') {
  const container = document.getElementById("livePreviewContainer");
  if (!container) return;

  let cadets = state.cadets;
  let absences = state.absences;
  let titleSuffix = "";

  if (slice === 'male') {
    cadets = state.cadets.filter(c => c.gender === 'М' || c.gender === 'male');
    absences = state.absences.filter(a => a.gender === 'М' || a.gender === 'male');
    titleSuffix = "(Мужской состав)";
  } else if (slice === 'female') {
    cadets = state.cadets.filter(c => c.gender === 'Ж' || c.gender === 'female');
    absences = state.absences.filter(a => a.gender === 'Ж' || a.gender === 'female');
    titleSuffix = "(Женский состав)";
  }

  if (cadets.length === 0) {
    const emptyMsg = slice === 'male'
      ? "В базе нет курсантов мужского пола"
      : slice === 'female'
      ? "В базе нет курсантов женского пола"
      : "База личного состава пуста";
    container.innerHTML = `
      <div class="empty-state">
        <svg class="empty-state-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>
        <div class="empty-state-title">${emptyMsg}</div>
        <p>Загрузите список курсантов из Excel или добавьте их вручную для формирования строевой записки.</p>
      </div>
    `;
    return;
  }

  const standardCategories = ["Наряд", "Госпиталь", "Лазарет", "Командировка", "Похороны"];
  const activeAbsenceCategories = [...new Set(absences.map(a => a.category))];
  const orderedCats = standardCategories.filter(c => activeAbsenceCategories.includes(c) || true);
  const otherCats = activeAbsenceCategories.filter(c => !standardCategories.includes(c));
  const categories = [...orderedCats, ...otherCats];

  const groupOrderMap = {};
  cadets.forEach(c => {
    if (!(c.group_num in groupOrderMap)) groupOrderMap[c.group_num] = c.sort_order || 0;
  });
  const groups = Object.keys(groupOrderMap).sort((a, b) => groupOrderMap[a] - groupOrderMap[b]);

  let totalPs = 0;
  let totalNl = 0;
  const catTotals = {};
  categories.forEach(c => { catTotals[c] = 0; });

  let groupRowsHtml = "";
  groups.forEach((grp, idx) => {
    const grpCadets = cadets.filter(c => c.group_num === grp);
    const ps = grpCadets.length;
    const grpAbs = absences.filter(a => a.group_num === grp);
    const grpAbsTotal = grpAbs.length;
    const nl = ps - grpAbsTotal;
    totalPs += ps;
    totalNl += nl;

    const rowCounts = {};
    categories.forEach(cat => {
      const cnt = grpAbs.filter(a => a.category === cat).length;
      rowCounts[cat] = cnt;
      catTotals[cat] += cnt;
    });

    groupRowsHtml += `
      <tr>
        <td>${idx + 1}</td>
        <td>${escapeHtml(grp)}</td>
        <td>${ps}</td>
        <td>${nl}</td>
        ${categories.map(c => `<td>${rowCounts[c] > 0 ? rowCounts[c] : ""}</td>`).join("")}
        <td colspan="2">${grpAbsTotal > 0 ? grpAbsTotal : ""}</td>
      </tr>
    `;
  });

  const hasMales = cadets.some(c => c.gender === 'М' || c.gender === 'male');
  const hasFemales = cadets.some(c => c.gender === 'Ж' || c.gender === 'female');

  let genderRowsHtml = "";
  if (slice === 'total' && hasMales && hasFemales) {
    genderRowsHtml = [
      { label: "Мужчины", val: "М", val2: "male", numIdx: groups.length + 1 },
      { label: "Женщины", val: "Ж", val2: "female", numIdx: groups.length + 2 }
    ].map(g => {
      const gCadets = cadets.filter(c => c.gender === g.val || c.gender === g.val2);
      const gPs = gCadets.length;
      const gAbs = absences.filter(a => a.gender === g.val || a.gender === g.val2);
      const gNl = gPs - gAbs.length;

      const rowCounts = {};
      categories.forEach(cat => {
        rowCounts[cat] = gAbs.filter(a => a.category === cat).length;
      });

      return `
        <tr>
          <td>${g.numIdx}</td>
          <td>${g.label}</td>
          <td>${gPs}</td>
          <td>${gNl}</td>
          ${categories.map(c => `<td>${rowCounts[c] > 0 ? rowCounts[c] : ""}</td>`).join("")}
          <td colspan="2">${gAbs.length > 0 ? gAbs.length : ""}</td>
        </tr>
      `;
    }).join("");
  }

  const totalAbsAll = Object.values(catTotals).reduce((sum, v) => sum + v, 0);
  const totalNumIdx = (slice === 'total' && hasMales && hasFemales) ? groups.length + 3 : groups.length + 1;
  const totalRowHtml = `
    <tr class="bold-cell">
      <td>${totalNumIdx}</td>
      <td>Всего</td>
      <td>${totalPs}</td>
      <td>${totalNl}</td>
      ${categories.map(c => `<td>${catTotals[c] > 0 ? catTotals[c] : ""}</td>`).join("")}
      <td colspan="2">${totalAbsAll > 0 ? totalAbsAll : ""}</td>
    </tr>
  `;

  let absencesRowsHtml = "";
  if (absences.length === 0) {
    absencesRowsHtml = `<tr><td colspan="${6 + categories.length}" style="text-align: center; padding: 12px;">Все курсанты находятся в строю</td></tr>`;
  } else {
    absences.forEach((a, idx) => {
      absencesRowsHtml += `
        <tr>
          <td>${idx + 1}</td>
          <td>${escapeHtml(a.rank || "курсант")}</td>
          <td colspan="3">${escapeHtml(a.short_fio || a.fio || "")}</td>
          <td colspan="2">${escapeHtml(a.type_name || "")}</td>
          <td>${escapeHtml(a.group_num || "")}</td>
          <td colspan="${categories.length - 2}">${escapeHtml(a.note || "")}</td>
        </tr>
      `;
    });
  }

  const fullTitle = titleSuffix ? `РАЗВЕРНУТАЯ СТРОЕВАЯ ЗАПИСКА ${titleSuffix}` : "РАЗВЕРНУТАЯ СТРОЕВАЯ ЗАПИСКА";

  container.innerHTML = `
    <div class="excel-preview-box">
      <div style="text-align: center; font-weight: bold; font-size: 14px; margin-bottom: 4px;">${fullTitle}</div>
      <div style="text-align: center; font-size: 13px; margin-bottom: 4px;">${escapeHtml(state.settings.unitTitle)}</div>
      <div style="text-align: center; font-size: 13px; margin-bottom: 16px;">${escapeHtml(formatDateRussian(state.settings.reportDate))}</div>

      <table class="excel-preview-table">
        <thead>
          <tr>
            <th>№ п/п</th>
            <th>Подразд.</th>
            <th>П/С</th>
            <th>Н/Л</th>
            ${categories.map(c => `<th>${escapeHtml(c)}</th>`).join("")}
            <th colspan="2">ВСЕГО</th>
          </tr>
        </thead>
        <tbody>
          ${groupRowsHtml}
          ${genderRowsHtml}
          ${totalRowHtml}
        </tbody>
      </table>

      <div style="text-align: center; font-size: 13px; font-weight: bold; margin: 16px 0 8px 0;">Отсутствующие</div>
      <table class="excel-preview-table">
        <thead>
          <tr>
            <th>№ п/п</th>
            <th>В/звание</th>
            <th colspan="3">Фамилия и инициалы</th>
            <th colspan="2">Причина отсутствия</th>
            <th>Подразд.</th>
            <th colspan="${categories.length - 2}">Примечание</th>
          </tr>
        </thead>
        <tbody>
          ${absencesRowsHtml}
        </tbody>
      </table>

      <div style="margin-top: 24px; font-size: 13px; text-align: center;">
        <div style="margin-bottom: 8px;">${escapeHtml(state.settings.chiefTitle)}</div>
        <div style="white-space: pre;">${escapeHtml(state.settings.chiefName)}</div>
      </div>
    </div>
  `;
}

// 6. ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ И МОДАЛЬНЫЕ ОКНА

function escapeHtml(str) {
  if (str === null || str === undefined) return "";
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function showToast(message, type = "info") {
  const container = document.getElementById("toastContainer");
  if (!container) return;

  const toast = document.createElement("div");
  toast.className = `toast ${type}`;
  toast.innerHTML = `
    <svg style="width: 18px; height: 18px; flex-shrink: 0;" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
      ${type === "success" 
        ? '<path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/>' 
        : type === "danger" 
        ? '<circle cx="12" cy="12" r="10"/><line x1="15" y1="9" x2="9" y2="15"/><line x1="9" y1="9" x2="15" y2="15"/>'
        : '<circle cx="12" cy="12" r="10"/><line x1="12" y1="16" x2="12" y2="12"/><line x1="12" y1="8" x2="12.01" y2="8"/>'}
    </svg>
    <span>${escapeHtml(message)}</span>
  `;
  container.appendChild(toast);

  setTimeout(() => {
    toast.style.opacity = "0";
    toast.style.transition = "opacity 0.3s ease";
    setTimeout(() => toast.remove(), 300);
  }, 3500);
}

function openAddCadetModal() {
  cadetModalEditId = null;
  const modal = document.getElementById("cadetModal");
  document.getElementById("cadetModalTitle").textContent = "Добавление курсанта";
  document.getElementById("cadetModalFio").value = "";
  document.getElementById("cadetModalRank").value = "курсант";
  document.getElementById("cadetModalGroup").value = "";
  document.getElementById("cadetModalGenderM").checked = true;
  if (modal) modal.classList.add("active");
}

function openEditCadetModal(cadetId) {
  const cadet = state.cadets.find(c => c.id === cadetId);
  if (!cadet) return;

  cadetModalEditId = cadetId;
  const modal = document.getElementById("cadetModal");
  document.getElementById("cadetModalTitle").textContent = "Редактирование курсанта";
  document.getElementById("cadetModalFio").value = cadet.fio;
  document.getElementById("cadetModalRank").value = cadet.rank;
  document.getElementById("cadetModalGroup").value = cadet.group_num;
  if (cadet.gender === "Ж") {
    document.getElementById("cadetModalGenderF").checked = true;
  } else {
    document.getElementById("cadetModalGenderM").checked = true;
  }
  if (modal) modal.classList.add("active");
}

function closeCadetModal() {
  const modal = document.getElementById("cadetModal");
  if (modal) modal.classList.remove("active");
  cadetModalEditId = null;
}

function saveCadetModal() {
  const fio = document.getElementById("cadetModalFio").value.trim();
  const rank = document.getElementById("cadetModalRank").value.trim() || "курсант";
  const group = document.getElementById("cadetModalGroup").value.trim() || "1";
  const gender = document.getElementById("cadetModalGenderF").checked ? "Ж" : "М";

  if (!fio) {
    showToast("Введите ФИО курсанта", "danger");
    return;
  }

  const shortFio = makeShortFio(fio);

  if (cadetModalEditId) {
    const c = state.cadets.find(item => item.id === cadetModalEditId);
    if (c) {
      c.fio = fio;
      c.rank = rank;
      c.group_num = group;
      c.gender = gender;
      c.short_fio = shortFio;

      // Обновить в отсутствующих если он там есть
      state.absences.forEach(a => {
        if (a.cadet_id === cadetModalEditId) {
          a.fio = fio;
          a.rank = rank;
          a.group_num = group;
          a.gender = gender;
          a.short_fio = shortFio;
        }
      });
      showToast("Данные курсанта обновлены", "success");
    }
  } else {
    const seenGroups = [...new Set(state.cadets.map(c => c.group_num))];
    if (!seenGroups.includes(group)) seenGroups.push(group);

    state.cadets.push({
      id: nextCadetId++,
      rank: rank,
      fio: fio,
      short_fio: shortFio,
      group_num: group,
      gender: gender,
      sort_order: seenGroups.indexOf(group)
    });
    showToast("Курсант добавлен в базу", "success");
  }

  closeCadetModal();
  updateKpi();
  updateCadetGroupFilterOptions();
  renderCadetsTable();
  updateCadetSelectOptions();
  renderAbsencesTable();
  renderLivePreview();
}

// 7. СБРОС ВСЕХ ДАННЫХ (ZERO PERSISTENCE CLEAR)

function resetAllData() {
  if (!confirm("Внимание! Вы действительно хотите очистить все списки и текущий расход? Все данные сессии будут удалены из памяти браузера.")) {
    return;
  }

  state.cadets = [];
  state.absences = [];
  state.settings.reportDate = getTodayDateString();
  state.settings.unitTitle = "* курс * факультета (автоматизированных систем управления)";
  state.settings.chiefTitle = "Начальник * курса * факультета";
  state.settings.commanderRank = "";
  state.settings.commanderName = "";
  state.settings.chiefName = "";

  if (typeof localStorage !== 'undefined') {
    localStorage.removeItem('consumable_commander_rank');
    localStorage.removeItem('consumable_commander_name');
    localStorage.removeItem('consumable_chief_title');
  }
  syncChiefName();

  const reportDateEl = document.getElementById("reportDate");
  if (reportDateEl) reportDateEl.value = state.settings.reportDate;
  const unitTitleEl = document.getElementById("unitTitle");
  if (unitTitleEl) unitTitleEl.value = state.settings.unitTitle;
  const chiefTitleEl = document.getElementById("chiefTitle");
  if (chiefTitleEl) chiefTitleEl.value = state.settings.chiefTitle;
  const rankEl = document.getElementById("commanderRank");
  if (rankEl) rankEl.value = "";
  const nameEl = document.getElementById("commanderName");
  if (nameEl) nameEl.value = "";

  updateKpi();
  updateCadetGroupFilterOptions();
  renderAbsencesTable();
  renderCadetsTable();
  updateCadetSelectOptions();
  renderLivePreview();

  showToast("Все данные сессии успешно удалены из памяти", "success");
}

// 8. ИНИЦИАЛИЗАЦИЯ И СЛУШАТЕЛИ СОБЫТИЙ

document.addEventListener("DOMContentLoaded", () => {
  // Инициализация полей параметров
  const reportDateEl = document.getElementById("reportDate");
  const unitTitleEl = document.getElementById("unitTitle");
  const chiefTitleEl = document.getElementById("chiefTitle");
  const commanderRankEl = document.getElementById("commanderRank");
  const commanderNameEl = document.getElementById("commanderName");

  const savedRank = (typeof localStorage !== 'undefined' && localStorage.getItem('consumable_commander_rank')) || "";
  const savedName = (typeof localStorage !== 'undefined' && localStorage.getItem('consumable_commander_name')) || "";
  const savedTitle = (typeof localStorage !== 'undefined' && localStorage.getItem('consumable_chief_title'));

  state.settings.commanderRank = savedRank;
  if (commanderRankEl) commanderRankEl.value = savedRank;

  state.settings.commanderName = savedName;
  if (commanderNameEl) commanderNameEl.value = savedName;

  if (savedTitle) {
    state.settings.chiefTitle = savedTitle;
    if (chiefTitleEl) chiefTitleEl.value = savedTitle;
  }
  syncChiefName();

  if (reportDateEl) {
    reportDateEl.value = state.settings.reportDate;
    reportDateEl.addEventListener("change", e => {
      state.settings.reportDate = e.target.value;
      renderLivePreview();
    });
  }

  if (unitTitleEl) {
    unitTitleEl.value = state.settings.unitTitle;
    unitTitleEl.addEventListener("input", e => {
      state.settings.unitTitle = e.target.value;
      renderLivePreview();
    });
  }

  if (chiefTitleEl) {
    chiefTitleEl.value = state.settings.chiefTitle;
    chiefTitleEl.addEventListener("input", e => {
      state.settings.chiefTitle = e.target.value;
      if (typeof localStorage !== 'undefined') {
        localStorage.setItem('consumable_chief_title', e.target.value);
      }
      renderLivePreview();
    });
  }

  if (commanderRankEl) {
    commanderRankEl.value = state.settings.commanderRank;
    commanderRankEl.addEventListener("change", e => {
      const val = e.target.value.trim();
      state.settings.commanderRank = val;
      if (typeof localStorage !== 'undefined') {
        if (val) {
          localStorage.setItem('consumable_commander_rank', val);
        } else {
          localStorage.removeItem('consumable_commander_rank');
        }
      }
      syncChiefName();
      renderLivePreview();
    });
  }

  if (commanderNameEl) {
    commanderNameEl.value = state.settings.commanderName;
    commanderNameEl.addEventListener("input", e => {
      const val = e.target.value.trim();
      state.settings.commanderName = val;
      if (typeof localStorage !== 'undefined') {
        if (val) {
          localStorage.setItem('consumable_commander_name', val);
        } else {
          localStorage.removeItem('consumable_commander_name');
        }
      }
      syncChiefName();
      renderLivePreview();
    });
  }

  // Переключение вкладок
  document.querySelectorAll(".tab-btn").forEach(btn => {
    btn.addEventListener("click", () => {
      const target = btn.dataset.tab;
      document.querySelectorAll(".tab-btn").forEach(b => b.classList.remove("active"));
      document.querySelectorAll(".tab-pane").forEach(p => p.classList.remove("active"));

      btn.classList.add("active");
      const targetPane = document.getElementById(target);
      if (targetPane) targetPane.classList.add("active");

      if (target === "previewTab") {
        renderLivePreview();
      }
    });
  });

  // Загрузка файла Excel с личным составом
  const fileInput = document.getElementById("excelFileInput");
  if (fileInput) {
    fileInput.addEventListener("change", async e => {
      const file = e.target.files[0];
      if (!file) return;

      try {
        const data = await file.arrayBuffer();
        const wb = XLSX.read(data, { type: "array" });

        // Поиск и парсинг подписи должностного лица (начальника курса)
        const sig = parseSignatureFromWorkbook(wb);
        if (sig.foundRank) {
          state.settings.commanderRank = sig.foundRank;
          if (commanderRankEl) commanderRankEl.value = sig.foundRank;
          if (typeof localStorage !== 'undefined') {
            localStorage.setItem('consumable_commander_rank', sig.foundRank);
          }
        } else {
          state.settings.commanderRank = "";
          if (commanderRankEl) commanderRankEl.value = "";
          if (typeof localStorage !== 'undefined') {
            localStorage.removeItem('consumable_commander_rank');
          }
        }

        if (sig.foundName) {
          state.settings.commanderName = sig.foundName;
          if (commanderNameEl) commanderNameEl.value = sig.foundName;
          if (typeof localStorage !== 'undefined') {
            localStorage.setItem('consumable_commander_name', sig.foundName);
          }
        } else {
          state.settings.commanderName = "";
          if (commanderNameEl) commanderNameEl.value = "";
          if (typeof localStorage !== 'undefined') {
            localStorage.removeItem('consumable_commander_name');
          }
        }

        if (sig.foundTitle) {
          state.settings.chiefTitle = sig.foundTitle;
          if (chiefTitleEl) chiefTitleEl.value = sig.foundTitle;
          if (typeof localStorage !== 'undefined') {
            localStorage.setItem('consumable_chief_title', sig.foundTitle);
          }
        }
        syncChiefName();

        const priorityNames = ['база', 'личный состав', 'курсанты', 'список', 'список личного состава', 'ведомость'];
        let targetSheets = [];

        for (const sheetName of wb.SheetNames) {
          const cleanName = sheetName.trim().toLowerCase();
          if (priorityNames.some(p => cleanName.includes(p))) {
            targetSheets.push(wb.Sheets[sheetName]);
            break;
          }
        }

        if (targetSheets.length === 0) {
          for (const sname of wb.SheetNames) {
            const ws = wb.Sheets[sname];
            targetSheets.push(ws);
          }
        }

        if (targetSheets.length === 0) {
          targetSheets = [wb.Sheets[wb.SheetNames[0]]];
        }

        let loadedCadets = [];
        const seenGroups = [];

        targetSheets.forEach(ws => {
          const cadets = parseCadetsSheet(ws, "", seenGroups);
          loadedCadets = loadedCadets.concat(cadets);
        });

        if (loadedCadets.length === 0) {
          showToast("Не удалось найти данные курсантов в файле. Проверьте заголовки.", "danger");
          return;
        }

        // Открываем модальное окно выбора курса при импорте
        openImportCourseModal(loadedCadets);
      } catch (err) {
        console.error(err);
        showToast("Ошибка при чтении Excel-файла: " + err.message, "danger");
      } finally {
        fileInput.value = "";
      }
    });
  }

  // -------------------------------------------------------------
  // АВТОПОДСТАНОВКА КУРСАНТА (ПРИ ВВОДЕ ФАМИЛИИ)
  // -------------------------------------------------------------
  const cadetInput = document.getElementById("absenceCadetInput");
  const cadetIdInput = document.getElementById("absenceCadetId");
  const cadetDropdown = document.getElementById("cadetAutocompleteDropdown");

  let cadetHighlightIndex = -1;

  function renderCadetAutocomplete(query = "") {
    if (!cadetDropdown) return;
    const available = getAvailableCadets(query);
    cadetHighlightIndex = -1;

    if (available.length === 0) {
      cadetDropdown.innerHTML = `<div class="autocomplete-empty">${state.cadets.length === 0 ? "Сначала загрузите базу личного состава" : "Курсанты не найдены"}</div>`;
      cadetDropdown.classList.add("active");
      return;
    }

    const maxItems = Math.min(available.length, 30);
    let html = "";
    for (let i = 0; i < maxItems; i++) {
      const c = available[i];
      const crs = getCourseFromGroup(c.group_num);
      const crsText = crs !== "?" ? `${crs} к.` : `${c.course || 1} к.`;
      html += `
        <div class="autocomplete-item" data-id="${c.id}" data-index="${i}">
          <div class="autocomplete-item-text">
            <span style="color: var(--text-muted); font-size: 12px;">${escapeHtml(c.rank)}</span>
            <strong>${escapeHtml(c.short_fio)}</strong>
            <span style="font-size: 11px; color: var(--text-light);">(${escapeHtml(c.fio)})</span>
          </div>
          <span class="badge badge-gray">гр. ${escapeHtml(c.group_num)}, ${crsText}</span>
        </div>
      `;
    }
    cadetDropdown.innerHTML = html;
    cadetDropdown.classList.add("active");

    cadetDropdown.querySelectorAll(".autocomplete-item").forEach(item => {
      item.addEventListener("mousedown", e => {
        e.preventDefault();
        const cid = Number(item.dataset.id);
        const c = state.cadets.find(x => x.id === cid);
        if (c) selectCadet(c);
      });
    });
  }

  function selectCadet(c) {
    if (!c) return;
    const crs = getCourseFromGroup(c.group_num);
    const crsText = crs !== "?" ? `${crs} к.` : `${c.course || 1} к.`;
    cadetIdInput.value = c.id;
    cadetInput.value = `${c.short_fio} | ${c.rank} (гр. ${c.group_num}, ${crsText})`;
    if (cadetDropdown) cadetDropdown.classList.remove("active");
    cadetHighlightIndex = -1;

    // Автоматический переход к вводу причины
    const reasonInput = document.getElementById("absenceReasonInput");
    if (reasonInput) {
      reasonInput.focus();
      renderReasonAutocomplete(reasonInput.value);
    }
  }

  function updateCadetHighlight(items) {
    items.forEach((it, idx) => {
      if (idx === cadetHighlightIndex) {
        it.classList.add("selected");
        it.scrollIntoView({ block: "nearest" });
      } else {
        it.classList.remove("selected");
      }
    });
  }

  if (cadetInput) {
    cadetInput.addEventListener("focus", () => {
      renderCadetAutocomplete(cadetInput.value);
    });

    cadetInput.addEventListener("input", () => {
      cadetIdInput.value = "";
      renderCadetAutocomplete(cadetInput.value);
    });

    cadetInput.addEventListener("keydown", e => {
      const items = cadetDropdown ? cadetDropdown.querySelectorAll(".autocomplete-item") : [];

      if (e.key === "ArrowDown") {
        e.preventDefault();
        if (!cadetDropdown || !cadetDropdown.classList.contains("active")) {
          renderCadetAutocomplete(cadetInput.value);
          return;
        }
        if (items.length > 0) {
          cadetHighlightIndex = (cadetHighlightIndex + 1) % items.length;
          updateCadetHighlight(items);
        }
      } else if (e.key === "ArrowUp") {
        e.preventDefault();
        if (items.length > 0) {
          cadetHighlightIndex = (cadetHighlightIndex - 1 + items.length) % items.length;
          updateCadetHighlight(items);
        }
      } else if (e.key === "Enter") {
        e.preventDefault();
        if (cadetHighlightIndex >= 0 && items[cadetHighlightIndex]) {
          const cid = Number(items[cadetHighlightIndex].dataset.id);
          const c = state.cadets.find(x => x.id === cid);
          if (c) selectCadet(c);
        } else {
          const available = getAvailableCadets(cadetInput.value);
          if (available.length > 0) {
            selectCadet(available[0]);
          } else {
            showToast("Курсант с такой фамилией не найден", "danger");
          }
        }
      } else if (e.key === "Escape") {
        if (cadetDropdown) cadetDropdown.classList.remove("active");
      }
    });

    cadetInput.addEventListener("blur", () => {
      setTimeout(() => {
        if (cadetDropdown) cadetDropdown.classList.remove("active");
      }, 200);
    });
  }

  // -------------------------------------------------------------
  // АВТОПОДСТАНОВКА ПРИЧИНЫ ОТСУТСТВИЯ
  // -------------------------------------------------------------
  const reasonInput = document.getElementById("absenceReasonInput");
  const reasonCategoryInput = document.getElementById("absenceReasonCategory");
  const reasonDropdown = document.getElementById("reasonAutocompleteDropdown");
  const categoryBox = document.getElementById("customReasonCategoryBox");
  const categorySelect = document.getElementById("customReasonCategorySelect");

  let reasonHighlightIndex = -1;

  function renderReasonAutocomplete(query = "") {
    if (!reasonDropdown) return;
    const q = String(query || "").trim().toLowerCase();
    reasonHighlightIndex = -1;

    const matched = state.absenceTypes.filter(t => {
      if (!q) return true;
      return t.name.toLowerCase().includes(q) || t.category.toLowerCase().includes(q);
    });

    let html = "";
    matched.forEach((t, i) => {
      html += `
        <div class="autocomplete-item" data-id="${t.id}" data-name="${escapeHtml(t.name)}" data-category="${escapeHtml(t.category)}" data-index="${i}">
          <div class="autocomplete-item-main"><strong>${escapeHtml(t.name)}</strong></div>
          <span class="badge ${getCategoryBadgeClass(t.category)}">${escapeHtml(t.category)}</span>
        </div>
      `;
    });

    if (q && !state.absenceTypes.some(t => t.name.toLowerCase() === q)) {
      const autoCat = detectReasonCategory(q);
      html += `
        <div class="autocomplete-item custom-reason-option" data-name="${escapeHtml(query.trim())}" data-category="${autoCat}" data-index="${matched.length}">
          <div class="autocomplete-item-main" style="color: var(--primary);">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" class="icon-sm" style="display:inline-block; vertical-align:middle; margin-right:4px;"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg> Использовать «<strong>${escapeHtml(query.trim())}</strong>»
          </div>
          <span class="badge ${getCategoryBadgeClass(autoCat)}">${autoCat}</span>
        </div>
      `;
    }

    if (!html) {
      html = `<div class="autocomplete-empty">Причины не найдены</div>`;
    }

    reasonDropdown.innerHTML = html;
    reasonDropdown.classList.add("active");

    reasonDropdown.querySelectorAll(".autocomplete-item").forEach(item => {
      item.addEventListener("mousedown", e => {
        e.preventDefault();
        const rName = item.dataset.name;
        const rCat = item.dataset.category || detectReasonCategory(rName);
        selectReason(rName, rCat);
      });
    });
  }

  function selectReason(name, cat) {
    if (!name) return;
    reasonInput.value = name;
    reasonCategoryInput.value = cat;
    if (categorySelect) categorySelect.value = cat;

    if (reasonDropdown) reasonDropdown.classList.remove("active");
    reasonHighlightIndex = -1;

    const isStandard = state.absenceTypes.some(t => t.name.toLowerCase() === name.toLowerCase());
    if (categoryBox) {
      categoryBox.style.display = isStandard ? "none" : "flex";
    }

    const noteInput = document.getElementById("absenceNoteInput");
    if (noteInput) noteInput.focus();
  }

  function updateReasonHighlight(items) {
    items.forEach((it, idx) => {
      if (idx === reasonHighlightIndex) {
        it.classList.add("selected");
        it.scrollIntoView({ block: "nearest" });
      } else {
        it.classList.remove("selected");
      }
    });
  }

  if (reasonInput) {
    reasonInput.addEventListener("focus", () => {
      renderReasonAutocomplete(reasonInput.value);
    });

    reasonInput.addEventListener("input", () => {
      const val = reasonInput.value.trim();
      const detected = detectReasonCategory(val);
      reasonCategoryInput.value = detected;
      if (categorySelect) categorySelect.value = detected;

      const isStandard = state.absenceTypes.some(t => t.name.toLowerCase() === val.toLowerCase());
      if (categoryBox) {
        categoryBox.style.display = (val && !isStandard) ? "flex" : "none";
      }

      renderReasonAutocomplete(reasonInput.value);
    });

    reasonInput.addEventListener("keydown", e => {
      const items = reasonDropdown ? reasonDropdown.querySelectorAll(".autocomplete-item") : [];

      if (e.key === "ArrowDown") {
        e.preventDefault();
        if (!reasonDropdown || !reasonDropdown.classList.contains("active")) {
          renderReasonAutocomplete(reasonInput.value);
          return;
        }
        if (items.length > 0) {
          reasonHighlightIndex = (reasonHighlightIndex + 1) % items.length;
          updateReasonHighlight(items);
        }
      } else if (e.key === "ArrowUp") {
        e.preventDefault();
        if (items.length > 0) {
          reasonHighlightIndex = (reasonHighlightIndex - 1 + items.length) % items.length;
          updateReasonHighlight(items);
        }
      } else if (e.key === "Enter") {
        e.preventDefault();
        if (reasonHighlightIndex >= 0 && items[reasonHighlightIndex]) {
          const it = items[reasonHighlightIndex];
          selectReason(it.dataset.name, it.dataset.category || detectReasonCategory(it.dataset.name));
        } else {
          const val = reasonInput.value.trim();
          if (val) {
            const matched = state.absenceTypes.find(t => t.name.toLowerCase() === val.toLowerCase());
            if (matched) {
              selectReason(matched.name, matched.category);
            } else {
              selectReason(val, detectReasonCategory(val));
            }
          }
        }
      } else if (e.key === "Escape") {
        if (reasonDropdown) reasonDropdown.classList.remove("active");
      }
    });

    reasonInput.addEventListener("blur", () => {
      setTimeout(() => {
        if (reasonDropdown) reasonDropdown.classList.remove("active");
      }, 200);
    });
  }

  if (categorySelect) {
    categorySelect.addEventListener("change", e => {
      reasonCategoryInput.value = e.target.value;
    });
  }

  // Примечание: нажатие Enter сразу инициирует добавление в расход
  const noteInput = document.getElementById("absenceNoteInput");
  if (noteInput) {
    noteInput.addEventListener("keydown", e => {
      if (e.key === "Enter") {
        e.preventDefault();
        const addBtn = document.getElementById("addAbsenceBtn");
        if (addBtn) addBtn.click();
      }
    });
  }

  // Закрытие выпадающих списков при клике вне полей
  document.addEventListener("click", e => {
    if (!e.target.closest("#absenceCadetInput") && !e.target.closest("#cadetAutocompleteDropdown")) {
      if (cadetDropdown) cadetDropdown.classList.remove("active");
    }
    if (!e.target.closest("#absenceReasonInput") && !e.target.closest("#reasonAutocompleteDropdown")) {
      if (reasonDropdown) reasonDropdown.classList.remove("active");
    }
  });

  // -------------------------------------------------------------
  // ДОБАВЛЕНИЕ В РАСХОД
  // -------------------------------------------------------------
  const addAbsenceBtn = document.getElementById("addAbsenceBtn");
  if (addAbsenceBtn) {
    addAbsenceBtn.addEventListener("click", () => {
      let cadetId = Number(cadetIdInput ? cadetIdInput.value : 0);

      // Если cadetId не выставлен, пробуем найти по введенному тексту
      if (!cadetId && cadetInput && cadetInput.value.trim()) {
        const available = getAvailableCadets(cadetInput.value);
        if (available.length > 0) {
          cadetId = available[0].id;
        }
      }

      if (!cadetId) {
        showToast("Введите фамилию курсанта!", "danger");
        if (cadetInput) cadetInput.focus();
        return;
      }

      const cadet = state.cadets.find(c => c.id === cadetId);
      if (!cadet) {
        showToast("Курсант не найден в базе!", "danger");
        return;
      }

      if (state.absences.some(a => a.cadet_id === cadetId)) {
        showToast("Этот курсант уже добавлен в расход!", "danger");
        return;
      }

      const typeName = reasonInput ? reasonInput.value.trim() : "";
      if (!typeName) {
        showToast("Введите или выберите причину отсутствия!", "danger");
        if (reasonInput) reasonInput.focus();
        return;
      }

      const category = (categorySelect && categoryBox && categoryBox.style.display !== "none")
        ? categorySelect.value
        : (reasonCategoryInput.value || detectReasonCategory(typeName));

      // Если введена новая причина, сохраняем в справочник типов
      if (!state.absenceTypes.some(t => t.name.toLowerCase() === typeName.toLowerCase())) {
        state.absenceTypes.push({
          id: nextTypeId++,
          name: typeName,
          code: typeName.slice(0, 2).toUpperCase(),
          category: category
        });
        renderReasonsTable();
      }

      const note = noteInput ? noteInput.value.trim() : "";

      state.absences.push({
        id: nextAbsenceId++,
        cadet_id: cadet.id,
        rank: cadet.rank,
        fio: cadet.fio,
        short_fio: cadet.short_fio,
        group_num: cadet.group_num,
        gender: cadet.gender,
        type_name: typeName,
        category: category,
        note: note
      });

      // Очистка полей формы
      cadetInput.value = "";
      cadetIdInput.value = "";
      reasonInput.value = "";
      reasonCategoryInput.value = "Наряд";
      if (categoryBox) categoryBox.style.display = "none";
      if (noteInput) noteInput.value = "";

      updateKpi();
      renderAbsencesTable();
      updateCadetSelectOptions();
      renderLivePreview();

      showToast(`Курсант ${cadet.short_fio} добавлен в расход (${typeName})`, "success");

      // Возврат фокуса на ввод фамилии для быстрого добавления следующего курсанта
      if (cadetInput) cadetInput.focus();
    });
  }

  // Фильтрация курсантов во вкладке Личный состав
  const cadetSearchInput = document.getElementById("cadetSearchInput");
  const cadetGroupFilter = document.getElementById("cadetGroupFilter");
  if (cadetSearchInput) cadetSearchInput.addEventListener("input", renderCadetsTable);
  if (cadetGroupFilter) cadetGroupFilter.addEventListener("change", renderCadetsTable);

  // Автоопределение пола при вводе ФИО в модалке
  const modalFioInput = document.getElementById("cadetModalFio");
  if (modalFioInput) {
    modalFioInput.addEventListener("input", e => {
      const gen = detectGender(e.target.value);
      if (gen === "Ж") {
        document.getElementById("cadetModalGenderF").checked = true;
      } else {
        document.getElementById("cadetModalGenderM").checked = true;
      }
    });
  }

  // Добавление новой причины в справочник
  const addReasonBtn = document.getElementById("addReasonBtn");
  if (addReasonBtn) {
    addReasonBtn.addEventListener("click", () => {
      const name = document.getElementById("newReasonName").value.trim();
      const cat = document.getElementById("newReasonCategory").value.trim() || "Наряд";
      const code = document.getElementById("newReasonCode").value.trim();

      if (!name) {
        showToast("Введите название причины!", "danger");
        return;
      }

      state.absenceTypes.push({
        id: nextTypeId++,
        name: name,
        code: code,
        category: cat
      });

      document.getElementById("newReasonName").value = "";
      document.getElementById("newReasonCode").value = "";

      renderReasonsTable();
      showToast("Новая причина добавлена в справочник", "success");
    });
  }

  // Инициализация стартовых таблиц
  updateKpi();
  updateCadetGroupFilterOptions();
  renderCadetsTable();
  renderReasonsTable();
  updateCadetSelectOptions();
  renderAbsencesTable();
  renderLivePreview();
});
