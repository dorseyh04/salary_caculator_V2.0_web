export const PERFORMANCE_BASE_BY_LEVEL = { "一级": 2000, "二级": 3000, "三级": 4000 };
export const SALES_COMMISSION_RATES = {
  普通个体防护: 0.005,
  工业吸附: 0.005,
  液滤: 0.005,
  "液滤-电驻极": 0.005,
  空滤: 0.005,
  生活擦拭: 0.006,
  高端个体防护: 0.007,
  耐高温材料: 0.008,
  透气弹性材料: 0.01,
};
export const NEGATIVE_MARGIN_RATE = 0.0025;
export const FALLBACK_SALES_COMMISSION_RATE = 0.005;

const toNumber = (value, fallback = 0) => {
  const numberValue = Number(value);
  return Number.isFinite(numberValue) ? numberValue : fallback;
};

const normalizeHeader = (value) => String(value || "").replace(/\s+/g, "");

const findHeaderIndex = (headerRow, matcher) => {
  return headerRow.findIndex((cell) => matcher(normalizeHeader(cell)));
};

const getByIndex = (row, index, fallbackIndex = -1) => {
  if (index >= 0) return row[index];
  if (fallbackIndex >= 0) return row[fallbackIndex];
  return "";
};

const toText = (value) => String(value || "").trim();

const toMoneyNumber = (value, fallback = 0) => {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  const cleaned = String(value ?? "").replace(/,/g, "").trim();
  if (!cleaned) return fallback;
  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : fallback;
};

export function calculatePerformanceWage({ perfBase, perfScore }) {
  return toNumber(perfBase) * toNumber(perfScore);
}

export function getSalesCommissionRate(category) {
  const c = toText(category);
  if (!c) return { rate: FALLBACK_SALES_COMMISSION_RATE, warning: "分类为空，按最低档0.5%处理" };
  if (c === "贸易") return { rate: 0, warning: "贸易订单不参与销售提成" };
  if (c === "副牌") return { rate: FALLBACK_SALES_COMMISSION_RATE, warning: "副牌按最低档0.5%" };
  if (SALES_COMMISSION_RATES[c] !== undefined) return { rate: SALES_COMMISSION_RATES[c], warning: null };
  return { rate: FALLBACK_SALES_COMMISSION_RATE, warning: `未知分类"${c}"，暂按最低档0.5%` };
}

export function parseSettlementRows(rows) {
  if (!Array.isArray(rows) || rows.length < 2) {
    throw new Error("回款明细表为空或格式错误");
  }

  const headerIndex = rows.findIndex((row) => {
    const cells = (row || []).map(normalizeHeader);
    return cells.includes("销售订单") && cells.includes("物料编码");
  });

  if (headerIndex < 0) {
    throw new Error("回款明细表缺少表头：销售订单、物料编码");
  }

  const headerRow = rows[headerIndex] || [];
  const col = {
    year: findHeaderIndex(headerRow, (header) => header === "年"),
    month: findHeaderIndex(headerRow, (header) => header === "月"),
    orderNo: findHeaderIndex(headerRow, (header) => header === "销售订单"),
    orderType: findHeaderIndex(headerRow, (header) => header === "订单类型描述"),
    postingDate: findHeaderIndex(headerRow, (header) => header === "过账日期"),
    customerId: findHeaderIndex(headerRow, (header) => header === "售达方"),
    customerName: findHeaderIndex(headerRow, (header) => header === "售达方描述"),
    contractNo: findHeaderIndex(headerRow, (header) => header === "销售合同号"),
    materialCode: findHeaderIndex(headerRow, (header) => header === "物料编码"),
    materialDesc: findHeaderIndex(headerRow, (header) => header === "物料描述"),
    unit: findHeaderIndex(headerRow, (header) => header === "单位"),
    quantity: findHeaderIndex(headerRow, (header) => header === "数量"),
    taxUnitPrice: findHeaderIndex(headerRow, (header) => header === "含税单价"),
    taxAmount: findHeaderIndex(headerRow, (header) => header === "含税金额"),
    receiptAmount: findHeaderIndex(headerRow, (header) => header === "收款金额"),
    transactionDate: findHeaderIndex(headerRow, (header) => header === "交易日期"),
    diff: findHeaderIndex(headerRow, (header) => header === "差额"),
    remark: findHeaderIndex(headerRow, (header) => header === "备注"),
    receiptMonth: findHeaderIndex(headerRow, (header) => header === "回款月份"),
    manager: findHeaderIndex(headerRow, (header) => header === "业务经理"),
    basePrice: findHeaderIndex(headerRow, (header) => header.includes("基价")),
    freightUnit: findHeaderIndex(headerRow, (header) => header.includes("运费单价")),
    activityCoef: findHeaderIndex(headerRow, (header) => header.includes("提成活跃")),
  };
  const classificationCol = findHeaderIndex(headerRow, (header) => header === "分类");
  const legacyCategoryCol = findHeaderIndex(headerRow, (header) => header === "类别");
  const categoryCol = classificationCol >= 0 ? classificationCol : legacyCategoryCol;

  const settlements = [];
  for (let i = headerIndex + 1; i < rows.length; i += 1) {
    const row = rows[i] || [];
    if (!row.some((cell) => cell !== "" && cell !== null && cell !== undefined)) continue;

    const classification = toText(getByIndex(row, categoryCol, 19));
    settlements.push({
      _rowIdx: i + 1,
      年: getByIndex(row, col.year, 0),
      月: getByIndex(row, col.month, 1),
      销售订单: toText(getByIndex(row, col.orderNo, 2)),
      订单类型描述: getByIndex(row, col.orderType, 3),
      过账日期: getByIndex(row, col.postingDate, 4),
      售达方: toText(getByIndex(row, col.customerId, 5)),
      售达方描述: getByIndex(row, col.customerName, 6),
      销售合同号: getByIndex(row, col.contractNo, 7),
      物料编码: toText(getByIndex(row, col.materialCode, 8)),
      物料描述: getByIndex(row, col.materialDesc, 9),
      单位: getByIndex(row, col.unit, 10),
      数量: getByIndex(row, col.quantity, 11),
      含税单价: getByIndex(row, col.taxUnitPrice, 12),
      含税金额: toMoneyNumber(getByIndex(row, col.taxAmount, 13)),
      收款金额: toMoneyNumber(getByIndex(row, col.receiptAmount, 14)),
      交易日期: getByIndex(row, col.transactionDate, 15),
      差额: getByIndex(row, col.diff, 16),
      备注: toText(getByIndex(row, col.remark, 17)),
      回款月份: getByIndex(row, col.receiptMonth, 18),
      分类: classification,
      类别: classification,
      分类列名: classificationCol >= 0 ? "分类" : "类别",
      业务经理: toText(getByIndex(row, col.manager, 20)),
      基价_V列: getByIndex(row, col.basePrice, 21),
      运费单价: toMoneyNumber(getByIndex(row, col.freightUnit, 22)),
      提成活跃系数: getByIndex(row, col.activityCoef, 23) === "" || getByIndex(row, col.activityCoef, 23) === null || getByIndex(row, col.activityCoef, 23) === undefined
        ? null
        : Number(getByIndex(row, col.activityCoef, 23)),
    });
  }

  return settlements;
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
