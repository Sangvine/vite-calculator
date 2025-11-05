import { describe, it, expect, vi, beforeEach } from "vitest";

// Мокаем XLSX для тестов exportReport - должен быть перед импортом модуля
vi.mock("xlsx-js-style", () => ({
  default: {
    utils: {
      book_new: vi.fn(() => ({})),
      aoa_to_sheet: vi.fn(() => ({})),
      sheet_add_aoa: vi.fn(),
      encode_cell: vi.fn(({ r, c }) => `R${r}C${c}`),
      book_append_sheet: vi.fn(),
    },
    writeFile: vi.fn(),
  },
}));

import {
  getDiameterParams,
  generateReference,
  generateMeasured,
  exportReport,
  DIAMETER_PARAMS,
  TOLERANCE_BY_FLOW,
  type FormValues,
} from "./generateReport";
import XLSX from "xlsx-js-style";

describe("getDiameterParams", () => {
  it('должен корректно парсить диаметр "50 мм"', () => {
    const result = getDiameterParams("50 мм");
    expect(result).toEqual({ qMax: 70, qMin: 0.14 });
  });

  it('должен корректно парсить диаметр "100 мм"', () => {
    const result = getDiameterParams("100 мм");
    expect(result).toEqual({ qMax: 283, qMin: 0.55 });
  });

  it('должен корректно парсить диаметр "15 мм"', () => {
    const result = getDiameterParams("15 мм");
    expect(result).toEqual({ qMax: 6, qMin: 0.02 });
  });

  it('должен корректно парсить диаметр "200 мм"', () => {
    const result = getDiameterParams("200 мм");
    expect(result).toEqual({ qMax: 1130, qMin: 2.3 });
  });

  it("должен возвращать null для невалидной строки", () => {
    const result = getDiameterParams("невалидная строка");
    expect(result).toBeNull();
  });

  it("должен возвращать null для пустой строки", () => {
    const result = getDiameterParams("");
    expect(result).toBeNull();
  });

  it("должен парсить все диаметры из DIAMETER_PARAMS", () => {
    const diameters = Object.keys(DIAMETER_PARAMS).map((d) => `${d} мм`);
    diameters.forEach((diameter) => {
      const result = getDiameterParams(diameter);
      expect(result).not.toBeNull();
      expect(result).toHaveProperty("qMax");
      expect(result).toHaveProperty("qMin");
      expect(result?.qMax).toBeGreaterThan(0);
      expect(result?.qMin).toBeGreaterThan(0);
    });
  });

  it("должен корректно извлекать число из строки с дополнительными символами", () => {
    const result = getDiameterParams("Диаметр 50 мм трубы");
    expect(result).toEqual({ qMax: 70, qMin: 0.14 });
  });
});

describe("generateReference", () => {
  it("должен генерировать значение близкое к базовому (±0.01%)", () => {
    const baseValue = 100;
    const result = generateReference(baseValue);

    const delta = baseValue * 0.0001; // ±0.01%
    expect(result).toBeGreaterThanOrEqual(baseValue - delta);
    expect(result).toBeLessThanOrEqual(baseValue + delta);
  });

  it("должен возвращать число с 3 знаками после запятой", () => {
    const baseValue = 123.456789;
    const result = generateReference(baseValue);
    const decimalPlaces = result.toString().split(".")[1]?.length || 0;
    expect(decimalPlaces).toBeLessThanOrEqual(3);
  });

  it("должен возвращать число с 4 знаками после запятой при указании параметра", () => {
    const baseValue = 123.456789;
    const result = generateReference(baseValue, 4);
    const decimalPlaces = result.toString().split(".")[1]?.length || 0;
    expect(decimalPlaces).toBeLessThanOrEqual(4);
    expect(decimalPlaces).toBeGreaterThanOrEqual(3);
  });

  it("должен обрабатывать малые значения", () => {
    const baseValue = 0.01;
    const result = generateReference(baseValue);
    expect(result).toBeGreaterThan(0);
    expect(typeof result).toBe("number");
  });

  it("должен обрабатывать большие значения", () => {
    const baseValue = 10000;
    const result = generateReference(baseValue);
    expect(result).toBeGreaterThan(0);
    expect(typeof result).toBe("number");
  });

  // Рандомные тесты (100+ тестов)
  describe("рандомные тесты generateReference", () => {
    const randomTests = Array.from({ length: 150 }, () => {
      const baseValue = Math.random() * 10000 + 0.01;
      return { baseValue };
    });

    randomTests.forEach(({ baseValue }, index) => {
      it(`рандомный тест ${index + 1}: базовое значение ${baseValue.toFixed(
        3
      )}`, () => {
        const result = generateReference(baseValue);
        const delta = baseValue * 0.0001;
        const maxDelta = delta * 1.2; // запас для округления до 3 знаков

        expect(result).toBeGreaterThanOrEqual(baseValue - maxDelta - 0.001); // дополнительный запас для округления
        expect(result).toBeLessThanOrEqual(baseValue + maxDelta + 0.001);
        expect(typeof result).toBe("number");
        expect(Number.isFinite(result)).toBe(true);
      });
    });
  });
});

describe("generateMeasured", () => {
  beforeEach(() => {
    vi.spyOn(Math, "random").mockRestore();
  });

  it("должен генерировать значение с погрешностью в пределах ±80% от допустимой", () => {
    const gRef = 100;
    const allowedErrorPercent = 0.5; // ±0.5%
    const result = generateMeasured(gRef, allowedErrorPercent);

    const errorRange = (allowedErrorPercent / 100) * 0.8;
    const maxError = gRef * (1 + errorRange);
    const minError = gRef * (1 - errorRange);

    expect(result).toBeGreaterThanOrEqual(minError);
    expect(result).toBeLessThanOrEqual(maxError);
  });

  it("должен возвращать число с 3 знаками после запятой", () => {
    const gRef = 50;
    const allowedErrorPercent = 1.0;
    const result = generateMeasured(gRef, allowedErrorPercent);
    const decimalPlaces = result.toString().split(".")[1]?.length || 0;
    expect(decimalPlaces).toBeLessThanOrEqual(3);
  });

  it("должен возвращать число с 4 знаками после запятой при указании параметра", () => {
    const gRef = 50;
    const allowedErrorPercent = 1.0;
    const result = generateMeasured(gRef, allowedErrorPercent, 4);
    const decimalPlaces = result.toString().split(".")[1]?.length || 0;
    expect(decimalPlaces).toBeLessThanOrEqual(4);
    expect(decimalPlaces).toBeGreaterThanOrEqual(3);
  });

  // Тесты для всех уровней расхода
  it("должен генерировать значение для расхода 90% с допустимой погрешностью ±0.5", () => {
    const gRef = 100;
    const tolerance = TOLERANCE_BY_FLOW[90];
    const result = generateMeasured(gRef, tolerance.value);
    const errorRange = (tolerance.value / 100) * 0.8;
    const maxError = gRef * (1 + errorRange);
    const minError = gRef * (1 - errorRange);

    expect(result).toBeGreaterThanOrEqual(minError);
    expect(result).toBeLessThanOrEqual(maxError);
  });

  it("должен генерировать значение для расхода 50% с допустимой погрешностью ±0.5", () => {
    const gRef = 100;
    const tolerance = TOLERANCE_BY_FLOW[50];
    const result = generateMeasured(gRef, tolerance.value);
    const errorRange = (tolerance.value / 100) * 0.8;
    const maxError = gRef * (1 + errorRange);
    const minError = gRef * (1 - errorRange);

    expect(result).toBeGreaterThanOrEqual(minError);
    expect(result).toBeLessThanOrEqual(maxError);
  });

  it("должен генерировать значение для расхода 10% с допустимой погрешностью ±1.0", () => {
    const gRef = 100;
    const tolerance = TOLERANCE_BY_FLOW[10];
    const result = generateMeasured(gRef, tolerance.value);
    const errorRange = (tolerance.value / 100) * 0.8;
    const maxError = gRef * (1 + errorRange);
    const minError = gRef * (1 - errorRange);

    expect(result).toBeGreaterThanOrEqual(minError);
    expect(result).toBeLessThanOrEqual(maxError);
  });

  it("должен генерировать значение для расхода 2% с допустимой погрешностью ±2.0", () => {
    const gRef = 100;
    const tolerance = TOLERANCE_BY_FLOW[2];
    const result = generateMeasured(gRef, tolerance.value);
    const errorRange = (tolerance.value / 100) * 0.8;
    const maxError = gRef * (1 + errorRange);
    const minError = gRef * (1 - errorRange);

    expect(result).toBeGreaterThanOrEqual(minError);
    expect(result).toBeLessThanOrEqual(maxError);
  });

  // Рандомные тесты (100+ тестов)
  describe("рандомные тесты generateMeasured", () => {
    const randomTests = Array.from({ length: 150 }, () => {
      const gRef = Math.random() * 1000 + 0.1;
      const allowedErrorPercent = [0.5, 1.0, 2.0][
        Math.floor(Math.random() * 3)
      ];
      return { gRef, allowedErrorPercent };
    });

    randomTests.forEach(({ gRef, allowedErrorPercent }, index) => {
      it(`рандомный тест ${index + 1}: gRef=${gRef.toFixed(
        3
      )}, погрешность=${allowedErrorPercent}%`, () => {
        const result = generateMeasured(gRef, allowedErrorPercent);
        const errorRange = (allowedErrorPercent / 100) * 0.8;
        const maxError = gRef * (1 + errorRange * 1.1); // небольшой запас
        const minError = gRef * (1 - errorRange * 1.1);

        expect(result).toBeGreaterThanOrEqual(minError);
        expect(result).toBeLessThanOrEqual(maxError);
        expect(typeof result).toBe("number");
        expect(Number.isFinite(result)).toBe(true);
        expect(result).toBeGreaterThan(0);
      });
    });
  });

  it("должен проверять, что погрешность не превышает допустимую", () => {
    const gRef = 100;
    const allowedErrorPercent = 2.0;
    const iterations = 1000;

    for (let i = 0; i < iterations; i++) {
      const result = generateMeasured(gRef, allowedErrorPercent);
      const actualError = Math.abs(((result - gRef) / gRef) * 100);
      const maxAllowedError = allowedErrorPercent * 0.8 * 1.1; // с небольшим запасом

      expect(actualError).toBeLessThanOrEqual(maxAllowedError);
    }
  });
});

describe("Расчеты таблицы", () => {
  it("должен корректно вычислять базовое значение для 90% расхода", () => {
    const params = { qMax: 100, qMin: 10 };
    const flow = 90;
    const flowPercent = flow / 100;
    const baseValue = params.qMin + (params.qMax - params.qMin) * flowPercent;

    expect(baseValue).toBe(91);
  });

  it("должен корректно вычислять базовое значение для 50% расхода", () => {
    const params = { qMax: 100, qMin: 10 };
    const flow = 50;
    const flowPercent = flow / 100;
    const baseValue = params.qMin + (params.qMax - params.qMin) * flowPercent;

    expect(baseValue).toBe(55);
  });

  it("должен корректно вычислять базовое значение для 10% расхода", () => {
    const params = { qMax: 100, qMin: 10 };
    const flow = 10;
    const flowPercent = flow / 100;
    const baseValue = params.qMin + (params.qMax - params.qMin) * flowPercent;

    expect(baseValue).toBe(19);
  });

  it("должен корректно вычислять базовое значение для 2% расхода", () => {
    const params = { qMax: 100, qMin: 10 };
    const flow = 2;
    const flowPercent = flow / 100;
    const baseValue = params.qMin + (params.qMax - params.qMin) * flowPercent;

    expect(baseValue).toBe(11.8);
  });

  it("должен корректно рассчитывать погрешность", () => {
    const gRef = 100;
    const gMeas = 100.5;
    const errPct = ((gMeas - gRef) / gRef) * 100;

    expect(errPct).toBe(0.5);
  });

  it("должен корректно рассчитывать среднюю погрешность", () => {
    const errors = [0.1, 0.2, 0.3, 0.4, 0.5];
    const avgError = errors.reduce((a, b) => a + b, 0) / errors.length;

    expect(avgError).toBe(0.3);
  });

  it("должен корректно рассчитывать среднюю погрешность для пустого массива", () => {
    const errors: number[] = [];
    const avgError = errors.reduce((a, b) => a + b, 0) / (errors.length || 1);

    expect(avgError).toBe(0);
  });

  // Тесты с рандомными параметрами
  describe("рандомные тесты расчетов", () => {
    const flowRows = [90, 50, 10, 2];
    const randomTests = Array.from({ length: 100 }, () => {
      const qMax = Math.random() * 1000 + 10;
      const qMin = Math.random() * 10 + 0.01;
      const flow = flowRows[Math.floor(Math.random() * flowRows.length)];
      const measurementsCount = Math.floor(Math.random() * 10) + 1;
      return { qMax, qMin, flow, measurementsCount };
    });

    randomTests.forEach(({ qMax, qMin, flow, measurementsCount }, index) => {
      it(`рандомный тест расчета ${index + 1}: qMax=${qMax.toFixed(
        2
      )}, qMin=${qMin.toFixed(
        2
      )}, flow=${flow}%, measurements=${measurementsCount}`, () => {
        const flowPercent = flow / 100;
        const baseValue = qMin + (qMax - qMin) * flowPercent;

        expect(baseValue).toBeGreaterThanOrEqual(qMin);
        expect(baseValue).toBeLessThanOrEqual(qMax);

        const tolerance = TOLERANCE_BY_FLOW[flow];
        expect(tolerance).toBeDefined();

        // Генерируем несколько измерений
        const errors: number[] = [];
        for (let i = 0; i < measurementsCount; i++) {
          const gRef = generateReference(baseValue);
          const gMeas = generateMeasured(gRef, tolerance.value);
          const errPct = ((gMeas - gRef) / gRef) * 100;
          errors.push(errPct);
        }

        const avgError = errors.reduce((a, b) => a + b, 0) / errors.length;
        expect(Number.isFinite(avgError)).toBe(true);
        expect(Math.abs(avgError)).toBeLessThan(tolerance.value * 2); // средняя погрешность должна быть в разумных пределах
      });
    });
  });
});

describe("exportReport", () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it("должен создавать Excel файл", () => {
    const values: FormValues = {
      pipeDiameter: "50 мм",
      deviceNumber: "TEST123",
      measurementsCount: {
        flow90: 3,
        flow50: 3,
        flow10: 3,
        flow2: 3,
      },
    };

    exportReport(values);

    expect(XLSX.utils.book_new).toHaveBeenCalled();
    expect(XLSX.utils.aoa_to_sheet).toHaveBeenCalled();
    expect(XLSX.utils.book_append_sheet).toHaveBeenCalled();
    expect(XLSX.writeFile).toHaveBeenCalled();
  });

  it("должен корректно обрабатывать значения по умолчанию", () => {
    const values: FormValues = {
      pipeDiameter: "50 мм",
      deviceNumber: "",
      measurementsCount: {
        flow90: 1,
        flow50: 1,
        flow10: 1,
        flow2: 1,
      },
    };

    exportReport(values);

    expect(XLSX.writeFile).toHaveBeenCalled();
    const callArgs = (XLSX.writeFile as any).mock.calls[0];
    expect(callArgs[1]).toContain("без_номера");
  });

  it("должен использовать правильное имя файла с номером устройства", () => {
    const values: FormValues = {
      pipeDiameter: "100 мм",
      deviceNumber: "DEVICE-456",
      measurementsCount: {
        flow90: 5,
        flow50: 5,
        flow10: 5,
        flow2: 5,
      },
    };

    exportReport(values);

    const callArgs = (XLSX.writeFile as any).mock.calls[0];
    expect(callArgs[1]).toContain("DEVICE-456");
  });

  it("должен обрабатывать минимальное количество измерений", () => {
    const values: FormValues = {
      pipeDiameter: "50 мм",
      deviceNumber: "TEST",
      measurementsCount: {
        flow90: 0, // должно стать 1
        flow50: 0,
        flow10: 0,
        flow2: 0,
      },
    };

    exportReport(values);

    expect(XLSX.writeFile).toHaveBeenCalled();
  });

  it("должен обрабатывать различные диаметры", () => {
    const diameters = Object.keys(DIAMETER_PARAMS).map((d) => `${d} мм`);

    diameters.forEach((diameter) => {
      const values: FormValues = {
        pipeDiameter: diameter,
        deviceNumber: "TEST",
        measurementsCount: {
          flow90: 3,
          flow50: 3,
          flow10: 3,
          flow2: 3,
        },
      };

      exportReport(values);
    });

    expect(XLSX.writeFile).toHaveBeenCalledTimes(diameters.length);
  });
});

describe("Интеграционные тесты с рандомными данными", () => {
  it("должен обрабатывать 150+ рандомных комбинаций", () => {
    const combinations = Array.from({ length: 150 }, () => {
      const diameters = Object.keys(DIAMETER_PARAMS);
      const randomDiameter =
        diameters[Math.floor(Math.random() * diameters.length)];
      const count = Math.floor(Math.random() * 10) + 1;
      return {
        pipeDiameter: `${randomDiameter} мм`,
        deviceNumber: `DEV-${Math.floor(Math.random() * 10000)}`,
        measurementsCount: {
          flow90: count,
          flow50: count,
          flow10: count,
          flow2: count,
        },
      } as FormValues;
    });

    combinations.forEach((values, index) => {
      const params = getDiameterParams(values.pipeDiameter);
      expect(params).not.toBeNull();

      if (params) {
        const flowRows = [90, 50, 10, 2];
        flowRows.forEach((flow) => {
          const flowPercent = flow / 100;
          const baseValue =
            params.qMin + (params.qMax - params.qMin) * flowPercent;
          const tolerance = TOLERANCE_BY_FLOW[flow];

          expect(baseValue).toBeGreaterThanOrEqual(params.qMin);
          expect(baseValue).toBeLessThanOrEqual(params.qMax);
          expect(tolerance).toBeDefined();

          // Генерируем измерения
          const count =
            flow === 90
              ? values.measurementsCount.flow90
              : flow === 50
              ? values.measurementsCount.flow50
              : flow === 10
              ? values.measurementsCount.flow10
              : values.measurementsCount.flow2;
          for (let i = 0; i < count; i++) {
            const gRef = generateReference(baseValue);
            const gMeas = generateMeasured(gRef, tolerance.value);
            const errPct = ((gMeas - gRef) / gRef) * 100;

            expect(gRef).toBeGreaterThan(0);
            expect(gMeas).toBeGreaterThan(0);
            expect(Number.isFinite(errPct)).toBe(true);
          }
        });
      }
    });
  });

  it("должен обрабатывать граничные случаи", () => {
    // Минимальный диаметр
    const minDiameter = Object.keys(DIAMETER_PARAMS)
      .map(Number)
      .sort((a, b) => a - b)[0];
    const minParams = getDiameterParams(`${minDiameter} мм`);
    expect(minParams).not.toBeNull();

    // Максимальный диаметр
    const maxDiameter = Object.keys(DIAMETER_PARAMS)
      .map(Number)
      .sort((a, b) => b - a)[0];
    const maxParams = getDiameterParams(`${maxDiameter} мм`);
    expect(maxParams).not.toBeNull();

    // Минимальное количество измерений
    const minMeasurements = 1;
    const params = minParams || { qMax: 60, qMin: 0.12 };
    const flowRows = [90, 50, 10, 2];
    flowRows.forEach((flow) => {
      const flowPercent = flow / 100;
      const baseValue = params.qMin + (params.qMax - params.qMin) * flowPercent;
      const gRef = generateReference(baseValue);
      expect(gRef).toBeGreaterThan(0);
    });

    // Максимальное количество измерений (20)
    const maxMeasurements = 20;
    for (let i = 0; i < maxMeasurements; i++) {
      const gRef = generateReference(100);
      const gMeas = generateMeasured(gRef, 0.5);
      expect(gRef).toBeGreaterThan(0);
      expect(gMeas).toBeGreaterThan(0);
    }
  });
});
