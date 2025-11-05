// @ts-ignore - используем библиотеку со стилями
import XLSX from "xlsx-js-style";

export type FormValues = {
  pipeDiameter: string;
  deviceNumber: string;
  measurementsCount: {
    flow90: number;
    flow50: number;
    flow10: number;
    flow2: number;
  };
};

export const FLOW_ROWS = [90, 50, 10, 2];

// Допустимые погрешности для каждого процента расхода
export const TOLERANCE_BY_FLOW: Record<
  number,
  { value: number; display: string }
> = {
  90: { value: 0.5, display: "±0.5" },
  50: { value: 0.5, display: "±0.5" },
  10: { value: 1.0, display: "±1.0" },
  2: { value: 2.0, display: "±2.0" },
};

// Таблица основных параметров по диаметрам
export const DIAMETER_PARAMS = {
  15: { qMax: 6, qMin: 0.02 },
  20: { qMax: 12, qMin: 0.03 },
  25: { qMax: 18, qMin: 0.036 },
  32: { qMax: 30, qMin: 0.06 },
  40: { qMax: 45, qMin: 0.09 },
  50: { qMax: 70, qMin: 0.14 },
  80: { qMax: 181, qMin: 0.36 },
  100: { qMax: 283, qMin: 0.55 },
  125: { qMax: 400, qMin: 33 },
  150: { qMax: 636, qMin: 1.3 },
  200: { qMax: 1130, qMin: 2.3 },
};

// Опции для селекта диаметров
export const DIAMETER_OPTIONS = Object.keys(DIAMETER_PARAMS).map((d) => ({
  value: `${d} мм`,
  label: `${d} мм`,
}));

export function getDiameterParams(
  diameterStr: string
): { qMax: number; qMin: number } | null {
  // Извлекаем число из строки (например, "50 мм" -> 50)
  const match = diameterStr.match(/\d+/);
  if (!match) return null;

  const diameter = parseInt(match[0], 10);
  return DIAMETER_PARAMS[diameter as keyof typeof DIAMETER_PARAMS] || null;
}

// Вторая колонка (Gобр) — образцовое значение, очень точное (±0.01%)
export function generateReference(
  baseValue: number,
  decimalPlaces: number = 3
): number {
  const delta = baseValue * 0.0001; // ±0.01%
  const value = baseValue + (Math.random() * 2 - 1) * delta;
  return Number(value.toFixed(decimalPlaces));
}

// Третья колонка (Gизм) — измеренное значение с погрешностью относительно Gобр
// Генерируем так, чтобы погрешность была в разумных пределах допустимой
export function generateMeasured(
  gRef: number,
  allowedErrorPercent: number,
  decimalPlaces: number = 3
): number {
  // Генерируем погрешность от -80% до +80% от допустимой
  // Это даст реалистичные значения, близкие к допустимым, но не превышающие их
  const errorRange = (allowedErrorPercent / 100) * 0.8; // 0.5% -> 0.004, 1.0% -> 0.008, 2.0% -> 0.016
  const randomError = (Math.random() * 2 - 1) * errorRange; // от -errorRange до +errorRange

  // Gизм = Gобр * (1 + погрешность)
  const gMeas = gRef * (1 + randomError);
  return Number(gMeas.toFixed(decimalPlaces));
}

function buildHeader(ws: any, values: FormValues) {
  const params = getDiameterParams(values.pipeDiameter);
  const rangeStr = params
    ? `${values.pipeDiameter.replace(/\D/g, "")}мм / ${params.qMin}-${
        params.qMax
      } куб.м/час`
    : "50мм / 0.12-60 куб.м/час"; // значение по умолчанию

  // Шапка документа (минимально необходимая разметка + объединения)
  const headerRows: (string | number)[][] = [
    [],
    ["ПРОТОКОЛ"],
    ["поверки расходомера-счётчика"],
    ["по методике СЕНА 407112.002 МП"],
    [],
    ["Тип прибора", "Поток-Омега"],
    ["Заводской номер", values.deviceNumber || "—"],
    ["ДУ / Диапазон измерения", rangeStr],
    ["Предприятие-изготовитель", 'ТОО "СП Поток-К"'],
    ["Тип поверочной установки", '"Контур-Сервис"'],
    ["Температура окружающей среды", "22 °C"],
    [],
    ["Таблица погрешностей."],
  ];

  XLSX.utils.sheet_add_aoa(ws, headerRows, { origin: { r: 0, c: 0 } });

  // Объединение для заголовка в несколько колонок
  const merges: any[] = [
    { s: { r: 1, c: 0 }, e: { r: 1, c: 6 } },
    { s: { r: 2, c: 0 }, e: { r: 2, c: 6 } },
    { s: { r: 3, c: 0 }, e: { r: 3, c: 6 } },
    { s: { r: 12, c: 0 }, e: { r: 12, c: 6 } },
  ];

  ws["!merges"] = (ws["!merges"] || []).concat(merges);
  ws["!cols"] = [
    { wch: 28 },
    { wch: 28 },
    { wch: 14 },
    { wch: 14 },
    { wch: 14 },
    { wch: 14 },
    { wch: 14 },
  ];
}

function buildTable(
  ws: any,
  measurementsCount: FormValues["measurementsCount"],
  pipeDiameter: string
) {
  const params = getDiameterParams(pipeDiameter);
  if (!params) {
    // Если диаметр не найден, используем значения по умолчанию
    const defaultParams = { qMax: 60, qMin: 0.12 };
    return buildTableWithParams(ws, measurementsCount, defaultParams);
  }

  return buildTableWithParams(ws, measurementsCount, params);
}

function buildTableWithParams(
  ws: any,
  measurementsCount: FormValues["measurementsCount"],
  params: { qMax: number; qMin: number }
) {
  // Заголовки таблицы — ближе к оригиналу
  const tableHeaderRow = [
    "Расход, %",
    "Gобр, м3/ч",
    "Gизм, м3/ч",
    "Погр., %",
    "Погр. ср., %",
    "Доп. Погр., %",
  ];
  XLSX.utils.sheet_add_aoa(ws, [tableHeaderRow], { origin: { r: 14, c: 0 } });

  let currentRow = 15; // строка после заголовка таблицы

  for (const flow of FLOW_ROWS) {
    // Получаем количество измерений для данного процента расхода
    let countForFlow: number;
    switch (flow) {
      case 90:
        countForFlow = measurementsCount.flow90;
        break;
      case 50:
        countForFlow = measurementsCount.flow50;
        break;
      case 10:
        countForFlow = measurementsCount.flow10;
        break;
      case 2:
        countForFlow = measurementsCount.flow2;
        break;
      default:
        countForFlow = 0;
    }

    // Заголовочная строка блока расхода (будет объединена по вертикали)
    XLSX.utils.sheet_add_aoa(ws, [[flow]], { origin: { r: currentRow, c: 0 } });

    // Допустимая погрешность для данного процента расхода
    const tolerance = TOLERANCE_BY_FLOW[flow] || {
      value: 0.5,
      display: "±0.5",
    };

    // Для 10% и 2% используем 4 знака после запятой, для остальных - 3
    const decimalPlaces = flow === 10 || flow === 2 ? 4 : 3;

    const errors: number[] = [];
    const actualCount = countForFlow === 0 ? 3 : countForFlow;

    // Если количество измерений равно 0, создаем 3 строки с нулевыми значениями
    if (countForFlow === 0) {
      const zeroValue = 0;
      const zeroError = 0;

      for (let i = 0; i < 3; i++) {
        errors.push(zeroError);
        const rowValues = ["", zeroValue, zeroValue, zeroError, "", ""];
        XLSX.utils.sheet_add_aoa(ws, [rowValues], {
          origin: { r: currentRow + 1 + i, c: 0 },
        });

        // Применяем форматирование для ячеек с нулевыми значениями
        const row = currentRow + 1 + i;
        const numFmt = decimalPlaces === 4 ? "0.0000" : "0.000";

        // Gобр (колонка 1) - формат с нужным количеством знаков после запятой
        const gRefAddr = XLSX.utils.encode_cell({ r: row, c: 1 });
        let gRefCell = ws[gRefAddr];
        if (!gRefCell) {
          gRefCell = { t: "n", v: 0, s: {} };
          ws[gRefAddr] = gRefCell;
        }
        if (!gRefCell.s) gRefCell.s = {};
        gRefCell.s.numFmt = numFmt;
        gRefCell.z = numFmt; // Также устанавливаем z для совместимости

        // Gизм (колонка 2) - формат с нужным количеством знаков после запятой
        const gMeasAddr = XLSX.utils.encode_cell({ r: row, c: 2 });
        let gMeasCell = ws[gMeasAddr];
        if (!gMeasCell) {
          gMeasCell = { t: "n", v: 0, s: {} };
          ws[gMeasAddr] = gMeasCell;
        }
        if (!gMeasCell.s) gMeasCell.s = {};
        gMeasCell.s.numFmt = numFmt;
        gMeasCell.z = numFmt; // Также устанавливаем z для совместимости

        // Погр. (колонка 3) - формат с 2 знаками после запятой
        const errAddr = XLSX.utils.encode_cell({ r: row, c: 3 });
        let errCell = ws[errAddr];
        if (!errCell) {
          errCell = { t: "n", v: 0, s: {} };
          ws[errAddr] = errCell;
        }
        if (!errCell.s) errCell.s = {};
        errCell.s.numFmt = "0.00";
        errCell.z = "0.00"; // Также устанавливаем z для совместимости
      }
    } else {
      // Вычисляем базовое значение для данного процента расхода
      // qMin = 0%, qMax = 100%
      // Для flow%: baseValue = qMin + (qMax - qMin) * (flow / 100)
      const flowPercent = flow / 100;
      const baseValue = params.qMin + (params.qMax - params.qMin) * flowPercent;
      const allowedErrorPercent = tolerance.value;

      // N строк измерений
      const numFmt = decimalPlaces === 4 ? "0.0000" : "0.000";
      for (let i = 0; i < countForFlow; i++) {
        // Генерируем образцовое значение (очень точное)
        const gRef = generateReference(baseValue, decimalPlaces);
        // Генерируем измеренное значение с погрешностью относительно Gобр
        const gMeas = generateMeasured(
          gRef,
          allowedErrorPercent,
          decimalPlaces
        );
        // Погрешности всегда отображаются с 2 знаками после запятой
        const errPct = Number((((gMeas - gRef) / gRef) * 100).toFixed(2));
        errors.push(errPct);

        const rowValues = ["", gRef, gMeas, errPct, "", ""];
        XLSX.utils.sheet_add_aoa(ws, [rowValues], {
          origin: { r: currentRow + 1 + i, c: 0 },
        });

        // Применяем форматирование для числовых ячеек
        const row = currentRow + 1 + i;
        // Gобр (колонка 1)
        const gRefAddr = XLSX.utils.encode_cell({ r: row, c: 1 });
        const gRefCell = ws[gRefAddr];
        if (gRefCell && !gRefCell.s) gRefCell.s = {};
        if (gRefCell) gRefCell.s.numFmt = numFmt;

        // Gизм (колонка 2)
        const gMeasAddr = XLSX.utils.encode_cell({ r: row, c: 2 });
        const gMeasCell = ws[gMeasAddr];
        if (gMeasCell && !gMeasCell.s) gMeasCell.s = {};
        if (gMeasCell) gMeasCell.s.numFmt = numFmt;

        // Погр. (колонка 3)
        const errAddr = XLSX.utils.encode_cell({ r: row, c: 3 });
        const errCell = ws[errAddr];
        if (errCell && !errCell.s) errCell.s = {};
        if (errCell) errCell.s.numFmt = "0.00";
      }
    }

    // Средняя погрешность также отображается с 2 знаками после запятой
    const avgError = Number(
      (errors.reduce((a, b) => a + b, 0) / (errors.length || 1)).toFixed(2)
    );

    // Запишем среднюю и доп. погрешности в строку расхода (и объединим по вертикали)
    XLSX.utils.sheet_add_aoa(ws, [[, , , , avgError, tolerance.display]], {
      origin: { r: currentRow, c: 0 },
    });

    // Применяем форматирование для средней погрешности (колонка 4)
    const avgErrorAddr = XLSX.utils.encode_cell({ r: currentRow, c: 4 });
    const avgErrorCell = ws[avgErrorAddr] || { t: "n", v: avgError };
    if (!avgErrorCell.s) avgErrorCell.s = {};
    avgErrorCell.s.numFmt = "0.00";
    ws[avgErrorAddr] = avgErrorCell;

    const mergesForBlock: any[] = [
      {
        s: { r: currentRow, c: 0 },
        e: { r: currentRow + actualCount, c: 0 },
      },
      {
        s: { r: currentRow, c: 4 },
        e: { r: currentRow + actualCount, c: 4 },
      },
      {
        s: { r: currentRow, c: 5 },
        e: { r: currentRow + actualCount, c: 5 },
      },
    ];
    ws["!merges"] = (ws["!merges"] || []).concat(mergesForBlock);

    // Пропуск к следующему блоку (1 строка заголовка расхода + N строк измерений)
    currentRow += 1 + actualCount + 1; // +1 пустая строка между блоками для читаемости
  }

  // вернуть границы диапазона таблицы для стилизации
  return { startRow: 14, endRow: currentRow - 2, startCol: 0, endCol: 5 };
}

function applyTableStyles(
  ws: any,
  range: { startRow: number; endRow: number; startCol: number; endCol: number }
) {
  const align = { horizontal: "center", vertical: "center" } as const;
  const border = {
    top: { style: "thin", color: { rgb: "000000" } },
    right: { style: "thin", color: { rgb: "000000" } },
    bottom: { style: "thin", color: { rgb: "000000" } },
    left: { style: "thin", color: { rgb: "000000" } },
  } as const;

  // Применяем стили ко всем ячейкам в диапазоне таблицы
  // Это включает: заголовки таблицы, первый столбец (Расход, %), столбцы с погрешностями и все остальные
  for (let r = range.startRow; r <= range.endRow; r++) {
    for (let c = range.startCol; c <= range.endCol; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      let cell = ws[addr];

      // Создаем ячейку, если её нет (для пустых ячеек)
      if (!cell) {
        cell = { t: "s", v: "" };
        ws[addr] = cell;
      }

      // Сохраняем существующие стили (особенно numFmt) при применении новых
      const existingNumFmt = cell.s?.numFmt;
      const existingZ = cell.z;
      const existingFill = cell.s?.fill;
      const existingFont = cell.s?.font;

      // Применяем центрирование и границы ко всем ячейкам
      // Это гарантирует, что:
      // - Первый столбец (колонка 0) - центрирован
      // - Столбцы с погрешностями (колонка 3 и 4) - центрированы
      // - Все ячейки имеют границы
      cell.s = {
        alignment: align,
        border,
        ...(existingNumFmt && { numFmt: existingNumFmt }),
        ...(existingFill && { fill: existingFill }),
        ...(existingFont && { font: existingFont }),
      };

      // Сохраняем z для совместимости
      if (existingZ) {
        cell.z = existingZ;
      }

      // Убеждаемся, что ячейка обновлена в worksheet
      ws[addr] = cell;
    }
  }
}

export function exportReport(values: FormValues) {
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([[]]);

  buildHeader(ws, values);
  const tableRange = buildTable(
    ws,
    {
      flow90: Number(values.measurementsCount.flow90 || 0),
      flow50: Number(values.measurementsCount.flow50 || 0),
      flow10: Number(values.measurementsCount.flow10 || 0),
      flow2: Number(values.measurementsCount.flow2 || 0),
    },
    values.pipeDiameter
  );
  applyTableStyles(ws, tableRange);

  // Повторно применяем форматирование для нулевых значений после стилизации
  // чтобы гарантировать, что numFmt не был перезаписан
  let currentRow = 15;
  for (const flow of FLOW_ROWS) {
    let countForFlow: number;
    switch (flow) {
      case 90:
        countForFlow = Number(values.measurementsCount.flow90 || 0);
        break;
      case 50:
        countForFlow = Number(values.measurementsCount.flow50 || 0);
        break;
      case 10:
        countForFlow = Number(values.measurementsCount.flow10 || 0);
        break;
      case 2:
        countForFlow = Number(values.measurementsCount.flow2 || 0);
        break;
      default:
        countForFlow = 0;
    }

    if (countForFlow === 0) {
      const decimalPlaces = flow === 10 || flow === 2 ? 4 : 3;
      const numFmt = decimalPlaces === 4 ? "0.0000" : "0.000";

      for (let i = 0; i < 3; i++) {
        const row = currentRow + 1 + i;

        // Gобр (колонка 1)
        const gRefAddr = XLSX.utils.encode_cell({ r: row, c: 1 });
        const gRefCell = ws[gRefAddr];
        if (gRefCell) {
          if (!gRefCell.s) gRefCell.s = {};
          gRefCell.s.numFmt = numFmt;
          gRefCell.z = numFmt;
        }

        // Gизм (колонка 2)
        const gMeasAddr = XLSX.utils.encode_cell({ r: row, c: 2 });
        const gMeasCell = ws[gMeasAddr];
        if (gMeasCell) {
          if (!gMeasCell.s) gMeasCell.s = {};
          gMeasCell.s.numFmt = numFmt;
          gMeasCell.z = numFmt;
        }

        // Погр. (колонка 3)
        const errAddr = XLSX.utils.encode_cell({ r: row, c: 3 });
        const errCell = ws[errAddr];
        if (errCell) {
          if (!errCell.s) errCell.s = {};
          errCell.s.numFmt = "0.00";
          errCell.z = "0.00";
        }
      }
    }

    const actualCount = countForFlow === 0 ? 3 : countForFlow;
    currentRow += 1 + actualCount + 1;
  }

  XLSX.utils.book_append_sheet(wb, ws, "Протокол");
  const fileName = `ПРОТОКОЛ_№${values.deviceNumber || "без_номера"}.xlsx`;
  XLSX.writeFile(wb, fileName);
}
