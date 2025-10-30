import * as XLSX from "xlsx";

export type FormValues = {
  pipeDiameter: string;
  deviceNumber: string;
  measurementsCount: number;
};

const FLOW_ROWS = [90, 50, 10, 2];

function generateRandomAround10(): number {
  const delta = 10 * 0.0005; // ±0.05% от 10
  const value = 10 + (Math.random() * 2 - 1) * delta;
  return Number(value.toFixed(5));
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
  // Заголовки таблицы как на фото (упрощённо)
  const tableHeaderRow1 = [
    "Расход, %",
    "Забр.",
    "Оизм.",
    "qизм.",
    "Погр., %",
    "Доп. Погр., %",
  ];
  XLSX.utils.sheet_add_aoa(ws, [tableHeaderRow1], { origin: { r: 14, c: 0 } });

  let currentRow = 15; // строка после заголовка таблицы

  for (const flow of FLOW_ROWS) {
    // Основная строка расхода
    XLSX.utils.sheet_add_aoa(ws, [[flow]], { origin: { r: currentRow, c: 0 } });

    // N дополнительных строк измерений
    for (let i = 0; i < measurementsCount; i++) {
      const rowValues = [
        "",
        generateRandomAround10(),
        generateRandomAround10(),
        generateRandomAround10(),
        generateRandomAround10(),
        "",
      ];
      XLSX.utils.sheet_add_aoa(ws, [rowValues], {
        origin: { r: currentRow + 1 + i, c: 0 },
      });
    }

    // Объединить ячейку расхода по вертикали с N строками ниже
    const merge: XLSX.Range = {
      s: { r: currentRow, c: 0 },
      e: { r: currentRow + measurementsCount, c: 0 },
    };
    ws["!merges"] = (ws["!merges"] || []).concat([merge]);

    // Пропустить к следующему блоку (1 строка заголовка расхода + N строк измерений)
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
