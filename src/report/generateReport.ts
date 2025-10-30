import * as XLSX from "xlsx";

export type FormValues = {
  pipeDiameter: string;
  deviceNumber: string;
  measurementsCount: number;
};

const FLOW_ROWS = [90, 50, 10, 2];

function generateAround10WithDeltaPercent(deltaPercent: number): number {
  const base = 10;
  const delta = base * deltaPercent; // доля, не проценты
  const value = base + (Math.random() * 2 - 1) * delta;
  return Number(value.toFixed(5));
}

// Третья колонка (Gизм) — ±0.05%
function generateMeasured(): number {
  return generateAround10WithDeltaPercent(0.0005);
}

// Вторая колонка (Gобр) — ±0.01%
function generateReference(): number {
  return generateAround10WithDeltaPercent(0.0001);
}

function buildHeader(ws: XLSX.WorkSheet, values: FormValues) {
  // Шапка документа (минимально необходимая разметка + объединения)
  const headerRows: (string | number)[][] = [
    [],
    ["ПРОТОКОЛ"],
    ["поверки расходомера-счётчика"],
    ["по методике СЕНА 407112.002 МП"],
    [],
    ["Тип прибора", "Поток-Омега"],
    ["Заводской номер", values.deviceNumber || "—"],
    ["ДУ / Диапазон измерения", "50мм / 0.12-60 куб.м/час"],
    ["Предприятие-изготовитель", 'ТОО "СП Поток-К"'],
    ["Тип поверочной установки", '"Контур-Сервис"'],
    ["Температура окружающей среды", "22 °C"],
    [],
    ["Таблица погрешностей."],
  ];

  XLSX.utils.sheet_add_aoa(ws, headerRows, { origin: { r: 0, c: 0 } });

  // Объединение для заголовка в несколько колонок
  const merges: XLSX.Range[] = [
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

function buildTable(ws: XLSX.WorkSheet, measurementsCount: number) {
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
    // Заголовочная строка блока расхода (будет объединена по вертикали)
    XLSX.utils.sheet_add_aoa(ws, [[flow]], { origin: { r: currentRow, c: 0 } });

    const errors: number[] = [];

    // N строк измерений
    for (let i = 0; i < measurementsCount; i++) {
      const gRef = generateReference();
      const gMeas = generateMeasured();
      const errPct = Number((((gMeas - gRef) / gRef) * 100).toFixed(5));
      errors.push(errPct);

      const rowValues = ["", gRef, gMeas, errPct, "", ""];
      XLSX.utils.sheet_add_aoa(ws, [rowValues], {
        origin: { r: currentRow + 1 + i, c: 0 },
      });
    }

    const avgError = Number(
      (errors.reduce((a, b) => a + b, 0) / (errors.length || 1)).toFixed(5)
    );

    const toleranceByFlow: Record<number, string> = {
      90: "±0.5",
      50: "±0.5",
      10: "±1.0",
      2: "±2.0",
    };
    const tolerance = toleranceByFlow[flow] || "";

    // Запишем среднюю и доп. погрешности в строку расхода (и объединим по вертикали)
    XLSX.utils.sheet_add_aoa(ws, [[, , , , avgError, tolerance]], {
      origin: { r: currentRow, c: 0 },
    });

    const mergesForBlock: XLSX.Range[] = [
      {
        s: { r: currentRow, c: 0 },
        e: { r: currentRow + measurementsCount, c: 0 },
      },
      {
        s: { r: currentRow, c: 4 },
        e: { r: currentRow + measurementsCount, c: 4 },
      },
      {
        s: { r: currentRow, c: 5 },
        e: { r: currentRow + measurementsCount, c: 5 },
      },
    ];
    ws["!merges"] = (ws["!merges"] || []).concat(mergesForBlock);

    // Пропуск к следующему блоку (1 строка заголовка расхода + N строк измерений)
    currentRow += 1 + measurementsCount + 1; // +1 пустая строка между блоками для читаемости
  }
}

export function exportReport(values: FormValues) {
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([[]]);

  buildHeader(ws, values);
  buildTable(ws, Math.max(1, Number(values.measurementsCount || 1)));

  XLSX.utils.book_append_sheet(wb, ws, "Протокол");
  const fileName = `ПРОТОКОЛ_№${values.deviceNumber || "без_номера"}.xlsx`;
  XLSX.writeFile(wb, fileName);
}
