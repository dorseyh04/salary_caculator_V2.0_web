export const PERFORMANCE_BASE_BY_LEVEL = { "一级": 2000, "二级": 3000, "三级": 4000 };

const toNumber = (value, fallback = 0) => {
  const numberValue = Number(value);
  return Number.isFinite(numberValue) ? numberValue : fallback;
};

const normalizeHeader = (value) => String(value || "").replace(/\s+/g, "");

const findHeaderIndex = (headerRow, matcher) => {
  return headerRow.findIndex((cell) => matcher(normalizeHeader(cell)));
};

export function calculatePerformanceWage({ perfBase, perfScore }) {
  return toNumber(perfBase) * toNumber(perfScore);
}

export function parsePerformanceRows(rows) {
  if (!Array.isArray(rows) || rows.length < 2) {
    throw new Error("绩效分表数据不足");
  }

  const headerIndex = rows.findIndex((row) => {
    const firstCell = normalizeHeader(row?.[0]);
    return firstCell === "姓名" || firstCell.includes("姓名");
  });

  if (headerIndex < 0) {
    throw new Error("绩效分表缺少姓名表头");
  }

  const headerRow = rows[headerIndex] || [];
  const nameCol = findHeaderIndex(headerRow, (header) => header.includes("姓名"));
  const scoreCol = findHeaderIndex(
    headerRow,
    (header) => /绩效分$/.test(header) || header === "当月绩效分" || header === "个人绩效分",
  );
  const baseSalaryCol = findHeaderIndex(headerRow, (header) => header.includes("基本工资"));
  const levelCol = findHeaderIndex(headerRow, (header) => header.includes("绩效等级") || header === "等级");
  const perfBaseCol = findHeaderIndex(headerRow, (header) => header.includes("绩效工资基数") || header === "绩效基数");

  if (nameCol < 0 || scoreCol < 0 || baseSalaryCol < 0 || levelCol < 0 || perfBaseCol < 0) {
    throw new Error("绩效分表缺少必要列：姓名、当月绩效分、基本工资标准、绩效等级、绩效工资基数");
  }

  const persons = [];
  for (let i = headerIndex + 1; i < rows.length; i += 1) {
    const row = rows[i] || [];
    const name = String(row[nameCol] || "").trim();
    if (!name) continue;

    const level = String(row[levelCol] || "一级").trim();
    const fallbackPerfBase = PERFORMANCE_BASE_BY_LEVEL[level] || PERFORMANCE_BASE_BY_LEVEL["一级"];

    persons.push({
      name,
      defaultBaseSalary: toNumber(row[baseSalaryCol], 2000),
      defaultPerfBase: toNumber(row[perfBaseCol], fallbackPerfBase),
      perfScore: toNumber(row[scoreCol]),
      level,
    });
  }

  return persons;
}
