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
    unitTitle: "3 курс 1 факультета (автоматизированных систем управления)",
    chiefTitle: "Начальник 3 курса 1 факультета",
    chiefName: "майор                          В.Медков"
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
    { id: 7, name: "Командировка", code: "К", category: "Командировка" }
  ],
  standardCategories: ["Наряд", "Госпиталь", "Лазарет", "Командировка", "Похороны", "Увольнение"]
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
  const parts = String(fio || '').trim().split(/\s+/).filter(Boolean);
  if (parts.length >= 3) {
    const patronymic = parts[2].toLowerCase();
    if (patronymic.endsWith('на')) {
      return 'Ж';
    }
    if (patronymic.endsWith('ич') || patronymic.endsWith('ыч')) {
      return 'М';
    }
  }
  if (parts.length > 0) {
    const surname = parts[0].toLowerCase();
    if (surname.endsWith('а') || surname.endsWith('я')) {
      return 'Ж';
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

    cadets.push({
      id: nextCadetId++,
      rank: rank,
      fio: fio,
      short_fio: shortFio,
      group_num: groupNum,
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

async function generateExcelReport() {
  if (state.cadets.length === 0) {
    showToast("Сначала загрузите базу личного состава!", "danger");
    return;
  }

  const wb = new ExcelJS.Workbook();
  wb.creator = "ConsumABLE Web";
  wb.created = new Date();
  const ws = wb.addWorksheet("Расход");

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
  ws.mergeCells("B2:L2");
  ws.getCell("B2").value = "РАЗВЕРНУТАЯ СТРОЕВАЯ ЗАПИСКА";
  applyStyleToRange(ws, 2, 2, 2, 12, { font: titleFont, alignment: centerAlign });

  ws.mergeCells("B3:L3");
  ws.getCell("B3").value = state.settings.unitTitle;
  applyStyleToRange(ws, 3, 2, 3, 12, { font: normFont, alignment: centerAlign });

  ws.mergeCells("B4:L4");
  const dtVal = formatDateRussian(state.settings.reportDate);
  ws.getCell("B4").value = dtVal;
  applyStyleToRange(ws, 4, 2, 4, 12, { font: normFont, alignment: centerAlign });
  ws.getRow(4).height = 17.0;

  // 2. Определение категорий для сводки
  const standardCategories = ["Наряд", "Госпиталь", "Лазарет", "Командировка", "Похороны"];
  const activeAbsenceCategories = [...new Set(state.absences.map(a => a.category))];
  const orderedCats = standardCategories.filter(c => activeAbsenceCategories.includes(c) || true); // Все стандартные
  const otherCats = activeAbsenceCategories.filter(c => !standardCategories.includes(c));
  const categories = [...orderedCats, ...otherCats];

  // Группы по порядку добавления / sort_order
  const groupOrderMap = {};
  state.cadets.forEach(c => {
    if (!(c.group_num in groupOrderMap)) {
      groupOrderMap[c.group_num] = c.sort_order || 0;
    }
  });
  const groups = Object.keys(groupOrderMap).sort((a, b) => groupOrderMap[a] - groupOrderMap[b]);

  const startRow = 5;
  ws.getRow(startRow).height = 18.0;

  const headers = ["№ п/п", "Подразд.", "П/С", "Н/Л", ...categories];
  const v_col1 = 2 + headers.length;
  const v_col2 = v_col1 + 1;

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

  let currRow = startRow + 1;
  let totalPs = 0;
  let totalNl = 0;
  const catTotals = {};
  categories.forEach(cat => { catTotals[cat] = 0; });

  // Строки подразделений
  groups.forEach((grp, idx) => {
    ws.getRow(currRow).height = 17.0;
    const grpCadets = state.cadets.filter(c => c.group_num === grp);
    const ps = grpCadets.length;
    const grpAbs = state.absences.filter(a => a.group_num === grp);

    const rowCounts = {};
    categories.forEach(cat => {
      const cnt = grpAbs.filter(a => a.category === cat).length;
      rowCounts[cat] = cnt;
      catTotals[cat] += cnt;
    });

    const grpAbsTotal = grpAbs.length;
    const nl = ps - grpAbsTotal;
    totalPs += ps;
    totalNl += nl;

    const rowVals = [
      idx + 1,
      grp,
      ps,
      nl,
      ...categories.map(cat => rowCounts[cat] > 0 ? rowCounts[cat] : "")
    ];

    rowVals.forEach((val, c_i) => {
      const cell = ws.getCell(currRow, 2 + c_i);
      cell.value = val;
      cell.font = normFont;
      cell.border = thinBorder;
      cell.alignment = centerAlign;
    });

    ws.mergeCells(currRow, v_col1, currRow, v_col2);
    ws.getCell(currRow, v_col1).value = grpAbsTotal > 0 ? grpAbsTotal : "";
    applyStyleToRange(ws, currRow, v_col1, currRow, v_col2, {
      font: normFont,
      border: thinBorder,
      alignment: centerAlign
    });

    currRow++;
  });

  // Строки «Мужчины» и «Женщины»
  const genderConfigs = [
    { label: "Мужчины", val: "М", numIdx: groups.length + 1 },
    { label: "Женщины", val: "Ж", numIdx: groups.length + 2 }
  ];

  genderConfigs.forEach(g => {
    ws.getRow(currRow).height = 17.0;
    const gCadets = state.cadets.filter(c => c.gender === g.val);
    const gPs = gCadets.length;
    const gAbs = state.absences.filter(a => a.gender === g.val);
    const gNl = gPs - gAbs.length;

    const rowCounts = {};
    categories.forEach(cat => {
      rowCounts[cat] = gAbs.filter(a => a.category === cat).length;
    });

    const gVals = [
      g.numIdx,
      g.label,
      gPs,
      gNl,
      ...categories.map(cat => rowCounts[cat] > 0 ? rowCounts[cat] : "")
    ];

    gVals.forEach((val, c_i) => {
      const cell = ws.getCell(currRow, 2 + c_i);
      cell.value = val;
      cell.font = normFont;
      cell.border = thinBorder;
      cell.alignment = centerAlign;
    });

    ws.mergeCells(currRow, v_col1, currRow, v_col2);
    ws.getCell(currRow, v_col1).value = gAbs.length > 0 ? gAbs.length : "";
    applyStyleToRange(ws, currRow, v_col1, currRow, v_col2, {
      font: normFont,
      border: thinBorder,
      alignment: centerAlign
    });

    currRow++;
  });

  // Строка «Всего»
  ws.getRow(currRow).height = 18.0;
  const totalAbsAll = Object.values(catTotals).reduce((sum, v) => sum + v, 0);
  const summaryVals = [
    groups.length + 3,
    "Всего",
    totalPs,
    totalNl,
    ...categories.map(c => catTotals[c] > 0 ? catTotals[c] : "")
  ];

  summaryVals.forEach((val, c_i) => {
    const cell = ws.getCell(currRow, 2 + c_i);
    cell.value = val;
    cell.font = boldFont;
    cell.border = thinBorder;
    cell.alignment = centerAlign;
  });

  ws.mergeCells(currRow, v_col1, currRow, v_col2);
  ws.getCell(currRow, v_col1).value = totalAbsAll > 0 ? totalAbsAll : "";
  applyStyleToRange(ws, currRow, v_col1, currRow, v_col2, {
    font: boldFont,
    border: thinBorder,
    alignment: centerAlign
  });

  currRow++;
  ws.getRow(currRow).height = 14.0;
  currRow++;

  // 3. Список отсутствующих
  const lastColLetter = getColumnLetter(v_col2);
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
  state.absences.forEach((a, idx) => {
    ws.getRow(currRow).height = 18.0;

    const cNum = ws.getCell(currRow, 2);
    cNum.value = idx + 1;
    cNum.font = normFont;
    cNum.border = thinBorder;
    cNum.alignment = centerAlign;

    const cRank = ws.getCell(currRow, 3);
    cRank.value = a.rank;
    cRank.font = normFont;
    cRank.border = thinBorder;
    cRank.alignment = centerAlign;

    ws.mergeCells(currRow, 4, currRow, 6);
    ws.getCell(currRow, 4).value = a.short_fio;
    applyStyleToRange(ws, currRow, 4, currRow, 6, { font: normFont, border: thinBorder, alignment: centerAlign });

    ws.mergeCells(currRow, 7, currRow, 8);
    ws.getCell(currRow, 7).value = a.type_name;
    applyStyleToRange(ws, currRow, 7, currRow, 8, { font: normFont, border: thinBorder, alignment: centerAlign });

    const cGrp = ws.getCell(currRow, 9);
    cGrp.value = a.group_num;
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
  ws.getCell(currRow, 2).value = `  ${state.settings.chiefTitle}`;
  ws.getCell(currRow, 2).font = normFont;
  currRow++;
  ws.getCell(currRow, 2).value = state.settings.chiefName;
  ws.getCell(currRow, 2).font = normFont;

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

  // 6. Параметры страницы для идеальной печати на лист А4 (1-в-1 с openpyxl)
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

  showToast("Строевая записка успешно сгенерирована и скачана!", "success");
}

// 5. УПРАВЛЕНИЕ ИНТЕРФЕЙСОМ И РЕНДЕРИНГ

function updateKpi() {
  const total = state.cadets.length;
  const absent = state.absences.length;
  const present = Math.max(0, total - absent);

  const kpiTotalEl = document.getElementById("kpiTotal");
  const kpiPresentEl = document.getElementById("kpiPresent");
  const kpiAbsentEl = document.getElementById("kpiAbsent");

  if (kpiTotalEl) kpiTotalEl.textContent = `${total} чел.`;
  if (kpiPresentEl) kpiPresentEl.textContent = `${present} чел.`;
  if (kpiAbsentEl) kpiAbsentEl.textContent = `${absent} чел.`;
}

function updateCadetSelectOptions() {
  const select = document.getElementById("absenceCadetSelect");
  const filterInput = document.getElementById("cadetFilterInput");
  if (!select) return;

  const filterText = filterInput ? filterInput.value.trim().toLowerCase() : "";
  const currentVal = select.value;

  select.innerHTML = '<option value="">-- Выберите курсанта --</option>';

  const absentCadetIds = new Set(state.absences.map(a => a.cadet_id));

  const availableCadets = state.cadets.filter(c => {
    if (absentCadetIds.has(c.id)) return false;
    if (!filterText) return true;
    return c.fio.toLowerCase().includes(filterText) ||
           c.group_num.toLowerCase().includes(filterText) ||
           c.rank.toLowerCase().includes(filterText);
  });

  availableCadets.forEach(c => {
    const opt = document.createElement("option");
    opt.value = c.id;
    opt.textContent = `${c.rank} ${c.fio} (${c.group_num} взв.)`;
    select.appendChild(opt);
  });

  if (currentVal && availableCadets.some(c => String(c.id) === String(currentVal))) {
    select.value = currentVal;
  }
}

function renderAbsencesTable() {
  const tbody = document.getElementById("absencesTableBody");
  const emptyState = document.getElementById("absencesEmptyState");
  const table = document.getElementById("absencesTable");
  const badgeCount = document.getElementById("absenceCountBadge");

  if (!tbody) return;

  if (badgeCount) {
    badgeCount.textContent = `${state.absences.length} чел.`;
  }

  if (state.absences.length === 0) {
    if (table) table.style.display = "none";
    if (emptyState) emptyState.style.display = "block";
    tbody.innerHTML = "";
    return;
  }

  if (table) table.style.display = "table";
  if (emptyState) emptyState.style.display = "none";

  tbody.innerHTML = "";

  state.absences.forEach((a, idx) => {
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

function removeAbsence(absenceId) {
  state.absences = state.absences.filter(a => a.id !== absenceId);
  updateKpi();
  renderAbsencesTable();
  updateCadetSelectOptions();
  renderLivePreview();
  showToast("Курсант возвращен в строй", "info");
}

function renderCadetsTable() {
  const tbody = document.getElementById("cadetsTableBody");
  const emptyState = document.getElementById("cadetsEmptyState");
  const table = document.getElementById("cadetsTable");
  const countEl = document.getElementById("cadetsCountBadge");
  const filterGroupSelect = document.getElementById("cadetGroupFilter");
  const searchInput = document.getElementById("cadetSearchInput");

  if (!tbody) return;

  if (countEl) countEl.textContent = `${state.cadets.length} чел.`;

  // Обновление фильтра подразделений
  if (filterGroupSelect) {
    const currentFilter = filterGroupSelect.value;
    const groups = [...new Set(state.cadets.map(c => c.group_num))].sort();
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

  const searchText = searchInput ? searchInput.value.trim().toLowerCase() : "";
  const selectedGroup = filterGroupSelect ? filterGroupSelect.value : "";

  const filtered = state.cadets.filter(c => {
    if (selectedGroup && c.group_num !== selectedGroup) return false;
    if (searchText) {
      return c.fio.toLowerCase().includes(searchText) ||
             c.rank.toLowerCase().includes(searchText) ||
             c.group_num.toLowerCase().includes(searchText);
    }
    return true;
  });

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
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td class="col-center">${idx + 1}</td>
      <td>${escapeHtml(c.rank)}</td>
      <td><strong>${escapeHtml(c.fio)}</strong></td>
      <td>${escapeHtml(c.short_fio)}</td>
      <td class="col-center"><span class="badge badge-gray">${escapeHtml(c.group_num)}</span></td>
      <td class="col-center"><span class="badge ${c.gender === 'Ж' ? 'badge-purple' : 'badge-blue'}">${c.gender}</span></td>
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

function deleteCadet(cadetId) {
  if (!confirm("Удалить курсанта из базы? Связанный расход также будет удален.")) {
    return;
  }
  state.cadets = state.cadets.filter(c => c.id !== cadetId);
  state.absences = state.absences.filter(a => a.cadet_id !== cadetId);
  updateKpi();
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

function renderLivePreview() {
  const container = document.getElementById("livePreviewContainer");
  if (!container) return;

  if (state.cadets.length === 0) {
    container.innerHTML = `
      <div class="empty-state">
        <svg class="empty-state-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>
        <div class="empty-state-title">База личного состава пуста</div>
        <p>Загрузите список курсантов из Excel или добавьте их вручную для формирования строевой записки.</p>
      </div>
    `;
    return;
  }

  const standardCategories = ["Наряд", "Госпиталь", "Лазарет", "Командировка", "Похороны"];
  const activeAbsenceCategories = [...new Set(state.absences.map(a => a.category))];
  const orderedCats = standardCategories.filter(c => activeAbsenceCategories.includes(c) || true);
  const otherCats = activeAbsenceCategories.filter(c => !standardCategories.includes(c));
  const categories = [...orderedCats, ...otherCats];

  const groupOrderMap = {};
  state.cadets.forEach(c => {
    if (!(c.group_num in groupOrderMap)) groupOrderMap[c.group_num] = c.sort_order || 0;
  });
  const groups = Object.keys(groupOrderMap).sort((a, b) => groupOrderMap[a] - groupOrderMap[b]);

  let totalPs = 0;
  let totalNl = 0;
  const catTotals = {};
  categories.forEach(c => { catTotals[c] = 0; });

  let groupRowsHtml = "";
  groups.forEach((grp, idx) => {
    const grpCadets = state.cadets.filter(c => c.group_num === grp);
    const ps = grpCadets.length;
    const grpAbs = state.absences.filter(a => a.group_num === grp);
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

  const genderRowsHtml = [
    { label: "Мужчины", val: "М", numIdx: groups.length + 1 },
    { label: "Женщины", val: "Ж", numIdx: groups.length + 2 }
  ].map(g => {
    const gCadets = state.cadets.filter(c => c.gender === g.val);
    const gPs = gCadets.length;
    const gAbs = state.absences.filter(a => a.gender === g.val);
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

  const totalAbsAll = Object.values(catTotals).reduce((sum, v) => sum + v, 0);
  const totalRowHtml = `
    <tr class="bold-cell">
      <td>${groups.length + 3}</td>
      <td>Всего</td>
      <td>${totalPs}</td>
      <td>${totalNl}</td>
      ${categories.map(c => `<td>${catTotals[c] > 0 ? catTotals[c] : ""}</td>`).join("")}
      <td colspan="2">${totalAbsAll > 0 ? totalAbsAll : ""}</td>
    </tr>
  `;

  let absencesRowsHtml = "";
  if (state.absences.length === 0) {
    absencesRowsHtml = `<tr><td colspan="${6 + categories.length}" style="text-align: center; padding: 12px;">Все курсанты находятся в строю</td></tr>`;
  } else {
    state.absences.forEach((a, idx) => {
      absencesRowsHtml += `
        <tr>
          <td>${idx + 1}</td>
          <td>${escapeHtml(a.rank)}</td>
          <td colspan="3">${escapeHtml(a.short_fio)}</td>
          <td colspan="2">${escapeHtml(a.type_name)}</td>
          <td>${escapeHtml(a.group_num)}</td>
          <td colspan="${categories.length - 2}">${escapeHtml(a.note || "")}</td>
        </tr>
      `;
    });
  }

  container.innerHTML = `
    <div class="excel-preview-box">
      <div style="text-align: center; font-weight: bold; font-size: 14px; margin-bottom: 4px;">РАЗВЕРНУТАЯ СТРОЕВАЯ ЗАПИСКА</div>
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

      <div style="margin-top: 24px; font-size: 13px;">
        <div style="margin-bottom: 8px;">&nbsp;&nbsp;${escapeHtml(state.settings.chiefTitle)}</div>
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
  state.settings.unitTitle = "3 курс 1 факультета (автоматизированных систем управления)";
  state.settings.chiefTitle = "Начальник 3 курса 1 факультета";
  state.settings.chiefName = "майор                          В.Медков";

  document.getElementById("reportDate").value = state.settings.reportDate;
  document.getElementById("unitTitle").value = state.settings.unitTitle;
  document.getElementById("chiefTitle").value = state.settings.chiefTitle;
  document.getElementById("chiefName").value = state.settings.chiefName;

  updateKpi();
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
  const chiefNameEl = document.getElementById("chiefName");

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
      renderLivePreview();
    });
  }

  if (chiefNameEl) {
    chiefNameEl.value = state.settings.chiefName;
    chiefNameEl.addEventListener("input", e => {
      state.settings.chiefName = e.target.value;
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

        state.cadets = loadedCadets;
        state.absences = []; // Сброс расхода при новой загрузке базы

        updateKpi();
        renderCadetsTable();
        updateCadetSelectOptions();
        renderAbsencesTable();
        renderLivePreview();

        showToast(`Успешно загружено курсантов: ${loadedCadets.length} чел.`, "success");
      } catch (err) {
        console.error(err);
        showToast("Ошибка при чтении Excel-файла: " + err.message, "danger");
      } finally {
        fileInput.value = "";
      }
    });
  }

  // Фильтр при выборе курсанта для расхода
  const cadetFilterInput = document.getElementById("cadetFilterInput");
  if (cadetFilterInput) {
    cadetFilterInput.addEventListener("input", () => {
      updateCadetSelectOptions();
    });
  }

  // Обработка переключения причины отсутствия
  const reasonSelect = document.getElementById("absenceReasonSelect");
  const customReasonFields = document.getElementById("customReasonFields");
  if (reasonSelect && customReasonFields) {
    reasonSelect.addEventListener("change", () => {
      if (reasonSelect.value === "custom") {
        customReasonFields.style.display = "flex";
      } else {
        customReasonFields.style.display = "none";
      }
    });
  }

  // Добавление в расход
  const addAbsenceBtn = document.getElementById("addAbsenceBtn");
  if (addAbsenceBtn) {
    addAbsenceBtn.addEventListener("click", () => {
      const cadetSelect = document.getElementById("absenceCadetSelect");
      const cadetId = Number(cadetSelect.value);
      if (!cadetId) {
        showToast("Выберите курсанта из списка!", "danger");
        return;
      }

      const cadet = state.cadets.find(c => c.id === cadetId);
      if (!cadet) return;

      if (state.absences.some(a => a.cadet_id === cadetId)) {
        showToast("Этот курсант уже добавлен в расход!", "danger");
        return;
      }

      const reasonSelect = document.getElementById("absenceReasonSelect");
      let typeName = "";
      let category = "Наряд";

      if (reasonSelect.value === "custom") {
        typeName = document.getElementById("customReasonName").value.trim();
        category = document.getElementById("customReasonCategory").value.trim() || "Наряд";
        if (!typeName) {
          showToast("Введите название причины отсутствия!", "danger");
          return;
        }
      } else {
        const typeId = Number(reasonSelect.value);
        const reasonObj = state.absenceTypes.find(t => t.id === typeId);
        if (!reasonObj) {
          showToast("Выберите причину отсутствия!", "danger");
          return;
        }
        typeName = reasonObj.name;
        category = reasonObj.category;
      }

      const note = document.getElementById("absenceNoteInput").value.trim();

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

      // Очистка формы
      cadetSelect.value = "";
      if (cadetFilterInput) cadetFilterInput.value = "";
      document.getElementById("absenceNoteInput").value = "";
      if (reasonSelect.value === "custom") {
        document.getElementById("customReasonName").value = "";
      }

      updateKpi();
      renderAbsencesTable();
      updateCadetSelectOptions();
      renderLivePreview();

      showToast(`Курсант ${cadet.short_fio} добавлен в расход`, "success");
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
  renderReasonsTable();
  updateCadetSelectOptions();
  renderAbsencesTable();
  renderCadetsTable();
  renderLivePreview();
});
