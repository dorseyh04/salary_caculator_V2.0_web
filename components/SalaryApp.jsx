import React, { useState, useEffect, useMemo, useCallback } from "react";
import * as XLSX from "xlsx";
import {
  Upload,
  FileSpreadsheet,
  Settings,
  Calculator,
  FileBarChart,
  Plus,
  Trash2,
  RotateCcw,
  AlertTriangle,
  CheckCircle2,
  ChevronRight,
  ChevronLeft,
  Download,
  FileWarning,
  Briefcase,
  Coins,
  Receipt,
  Users,
  TrendingUp,
  Activity,
} from "lucide-react";

// ============================================================
// 常量与系数表（依据《2026年康乃尔经营薪酬考核方案 V2.0》0515版）
// ============================================================
// 销售提成系数 — 按产品线（T列类别）
const PRODUCT_LINE_RATES = {
  普通个体防护: 0.004,
  高端个体防护: 0.006,
  工业吸附: 0.006,
  "液滤-电驻极": 0.006,
  液滤: 0.006,
  空滤: 0.006,
  生活擦拭: 0.006,
  耐高温材料: 0.008,
  透气弹性材料: 0.01,
};
const NEGATIVE_MARGIN_RATE = 0.0025;
const FUPAI_DEFAULT_RATE = 0.004;

const TRADE_RATES = { 熔喷料: 0.4, 母粒: 0.3, 无纺布: 0.2 };

const PERFORMANCE_BASE_BY_LEVEL = { 一级: 2000, 二级: 3000, 三级: 4000 };

const PRICE_DEDUCTIONS_KEY = "kne_price_deductions_v1";

// ============================================================
// 工具函数
// ============================================================
const fmtCNY = (n) => {
  if (n === null || n === undefined || isNaN(n)) return "—";
  return Number(n).toLocaleString("zh-CN", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
};
const fmtNum = (n, d = 2) => {
  if (n === null || n === undefined || isNaN(n)) return "—";
  return Number(n).toLocaleString("zh-CN", { minimumFractionDigits: d, maximumFractionDigits: d });
};
const fmtPct = (n, d = 2) => {
  if (n === null || n === undefined || isNaN(n)) return "—";
  return (Number(n) * 100).toFixed(d) + "%";
};
const fmtWan = (n) => fmtNum((n || 0) / 10000, 4);

function extractWeightFromDesc(desc) {
  if (!desc || typeof desc !== "string") return null;
  const matches = desc.match(/(\d+(?:\.\d+)?)\s*[Kk][Gg]/g);
  if (!matches || matches.length === 0) return null;
  return parseFloat(matches[matches.length - 1]) || null;
}

function convertToKg(qty, unit, desc) {
  if (qty === null || qty === undefined || isNaN(qty)) return { kg: null, source: "数量为空" };
  const q = Number(qty);
  if (!unit) return { kg: q, source: "单位缺失，按原值" };
  const u = String(unit).trim();
  if (u === "kg" || u === "KG" || u === "公斤" || u === "千克") return { kg: q, source: "K列=kg" };
  const w = extractWeightFromDesc(desc);
  if (w === null) return { kg: null, source: "无法从描述提取重量" };
  return { kg: q * w, source: `每${u}=${w}KG` };
}

function getProductLineRate(category) {
  if (!category) return { rate: FUPAI_DEFAULT_RATE, warning: "T列为空，按0.4%处理" };
  const c = String(category).trim();
  if (c === "贸易") return { rate: 0, warning: "贸易订单不参与销售提成" };
  if (c === "副牌") return { rate: FUPAI_DEFAULT_RATE, warning: "副牌按最低档0.4%" };
  if (PRODUCT_LINE_RATES[c] !== undefined) return { rate: PRODUCT_LINE_RATES[c], warning: null };
  return { rate: FUPAI_DEFAULT_RATE, warning: `未知类别"${c}"，暂按0.4%` };
}

function getPremiumRate(premiumRatio) {
  if (premiumRatio === null || premiumRatio === undefined || isNaN(premiumRatio)) return 0;
  if (premiumRatio <= 0) return 0;
  if (premiumRatio <= 0.05) return 0.15;
  if (premiumRatio <= 0.1) return 0.2;
  return 0.3;
}

function resolveBasePrice(vCellValue, materialCode, priceMap) {
  if (typeof vCellValue === "number" && !isNaN(vCellValue)) return { price: vCellValue, source: "V列直接值" };
  if (typeof vCellValue === "string") {
    const trimmed = vCellValue.trim();
    if (trimmed && !trimmed.startsWith("=")) {
      const num = parseFloat(trimmed);
      if (!isNaN(num)) return { price: num, source: "V列直接值" };
    }
  }
  if (materialCode && priceMap.has(String(materialCode))) return { price: priceMap.get(String(materialCode)), source: "基价表匹配" };
  return { price: null, source: "无法获取" };
}

// ============================================================
// Excel 解析
// ============================================================
async function parseSettlementWorkbook(file) {
  const arrayBuffer = await file.arrayBuffer();
  const wb = XLSX.read(arrayBuffer, { type: "array", cellDates: true, cellFormula: true });
  const sheet1Name = wb.SheetNames.find((n) => n.includes("回款明细")) || wb.SheetNames[0];
  const ws1 = wb.Sheets[sheet1Name];
  const rows1 = XLSX.utils.sheet_to_json(ws1, { header: 1, defval: "", raw: true });
  if (rows1.length < 2) throw new Error("回款明细表为空或格式错误");
  const dataRows1 = rows1.slice(1).filter((r) => r.some((c) => c !== "" && c !== null && c !== undefined));

  const settlements = dataRows1.map((row, idx) => ({
    _rowIdx: idx + 2,
    年: row[0], 月: row[1],
    销售订单: String(row[2] || "").trim(),
    订单类型描述: row[3], 过账日期: row[4],
    售达方: String(row[5] || "").trim(),
    售达方描述: row[6], 销售合同号: row[7],
    物料编码: String(row[8] || "").trim(),
    物料描述: row[9], 单位: row[10], 数量: row[11],
    含税单价: row[12],
    含税金额: typeof row[13] === "number" ? row[13] : parseFloat(row[13]) || 0,
    收款金额: typeof row[14] === "number" ? row[14] : parseFloat(row[14]) || 0,
    交易日期: row[15], 差额: row[16],
    备注: String(row[17] || ""),
    回款月份: row[18],
    类别: String(row[19] || "").trim(),
    业务经理: String(row[20] || "").trim(),
    基价_V列: row[21],
    运费单价: typeof row[22] === "number" ? row[22] : parseFloat(row[22]) || 0,
    提成活跃系数: row[23] === "" || row[23] === null || row[23] === undefined ? null : Number(row[23]),
  }));

  const sheet2Name = wb.SheetNames.find((n) => n.includes("基价")) || wb.SheetNames[1];
  const priceMap = new Map();
  if (sheet2Name) {
    const ws2 = wb.Sheets[sheet2Name];
    const rows2 = XLSX.utils.sheet_to_json(ws2, { header: 1, defval: "", raw: true });
    for (let i = 2; i < rows2.length; i++) {
      const code = rows2[i][0];
      const price = rows2[i][rows2[i].length - 1];
      if (code !== "" && code !== "/" && price !== "" && !isNaN(parseFloat(price))) {
        priceMap.set(String(code).trim(), parseFloat(price));
      }
    }
  }

  const monthGuess = settlements.find((s) => s.回款月份)?.回款月份 || settlements[0]?.月 || "";
  return { settlements, priceMap, month: String(monthGuess), sheet1Name, sheet2Name };
}

async function parsePerformanceWorkbook(file) {
  const arrayBuffer = await file.arrayBuffer();
  const wb = XLSX.read(arrayBuffer, { type: "array", cellDates: true });
  const sheetName = wb.SheetNames.find((n) => n.includes("汇总") || n.includes("绩效")) || wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: true });
  if (rows.length < 3) throw new Error("绩效考核表数据不足");

  const headerRow = rows[1] || [];
  // 自动跳过"抵扣"列（如果存在）— 找到完成率列的位置
  let colOffset = 1; // 默认 B 列起
  if (headerRow.some((h) => String(h || "").includes("抵扣"))) colOffset = 2; // 抵扣在 B 列，完成率从 C 列起

  const persons = [];
  for (let i = 2; i < rows.length; i++) {
    const row = rows[i];
    if (!row[0] || String(row[0]).trim() === "") continue;
    const name = String(row[0]).trim();
    const completionRate = Number(row[colOffset]) || 0;
    const perfScore = Number(row[colOffset + 1]) || 0;
    const baseSalary = Number(row[colOffset + 2]) || 2000;
    const level = String(row[colOffset + 3] || "一级").trim();
    const perfBase = Number(row[colOffset + 4]) || PERFORMANCE_BASE_BY_LEVEL[level] || 2000;
    persons.push({ name, defaultBaseSalary: baseSalary, defaultPerfBase: perfBase, completionRate, perfScore, level });
  }
  return { persons };
}

// ============================================================
// 计算引擎（无抵扣，使用含税金额）
// ============================================================
function runCalculation({ settlements, priceMap, persons, personSettings, priceDeductions, negativeOrders, tradeMarginInputs, activityOverrides }) {
  const warnings = [];
  const allLines = [], validRegular = [], tradeLines = [], dropped = [];
  const negativeSet = new Set(negativeOrders.map((o) => String(o).trim()).filter(Boolean));
  const priceDedMap = new Map(priceDeductions.map((d) => [String(d.customerId).trim(), Number(d.pricePerKg) || 0]));
  const activityMap = new Map(Object.entries(activityOverrides));
  const personNames = new Set(persons.map((p) => p.name));

  for (const r of settlements) {
    const lineBase = {
      ...r, _kg: null, _kgSource: null, _basePrice: null, _basePriceSource: null,
      _freightUnit: r.运费单价 || 0, _freightTotal: null, _exFactoryPrice: null,
      _commBase: null, _commRate: null, _commRateSource: null,
      _activityCoef: null, _activityOverride: false, _saleCommission: 0,
      _priceDeduction: 0, _adjPrice: null, _premiumRatio: null, _premiumRate: 0,
      _premiumBase: 0, _premiumCommission: 0, _isNegative: false,
      _flag: null, _dropReason: null,
    };

    if (!r.物料编码) { lineBase._flag = "dropped"; lineBase._dropReason = "物料编码为空"; dropped.push(lineBase); allLines.push(lineBase); continue; }
    const remarkStr = String(r.备注 || "");
    if (remarkStr.includes("预收款") || remarkStr.includes("无合同")) { lineBase._flag = "dropped"; lineBase._dropReason = `备注含"${remarkStr.includes("预收款") ? "预收款" : "无合同"}"`; dropped.push(lineBase); allLines.push(lineBase); continue; }

    const { kg, source: kgSource } = convertToKg(r.数量, r.单位, r.物料描述);
    lineBase._kg = kg; lineBase._kgSource = kgSource;
    if (kg === null || kg <= 0) { lineBase._flag = "dropped"; lineBase._dropReason = `单位折算失败：${kgSource}`; dropped.push(lineBase); allLines.push(lineBase); warnings.push({ type: "单位折算失败", order: r.销售订单, material: r.物料编码, person: r.业务经理, msg: `订单 ${r.销售订单} (物料 ${r.物料编码}) ${kgSource}` }); continue; }

    const { price, source: priceSource } = resolveBasePrice(r.基价_V列, r.物料编码, priceMap);
    lineBase._basePrice = price; lineBase._basePriceSource = priceSource;
    if (price === null || price <= 0) { lineBase._flag = "dropped"; lineBase._dropReason = "基价缺失"; dropped.push(lineBase); allLines.push(lineBase); warnings.push({ type: "基价缺失", order: r.销售订单, material: r.物料编码, person: r.业务经理, msg: `订单 ${r.销售订单} 物料 ${r.物料编码} 基价无法获取` }); continue; }

    lineBase._freightTotal = lineBase._freightUnit * kg;
    // ★ 核心变更：使用「含税金额」（N列）而非「收款金额」（O列）
    lineBase._exFactoryPrice = (Number(r.含税金额) - lineBase._freightTotal) / kg;

    if (r.类别 === "贸易") {
      const orderKey = `${r.销售订单}_${r.物料编码}`;
      const tradeInput = tradeMarginInputs[orderKey];
      lineBase._flag = "trade";
      lineBase._tradeCategory = tradeInput?.category || null;
      lineBase._tradeMargin = tradeInput && !isNaN(parseFloat(tradeInput.margin)) ? parseFloat(tradeInput.margin) : 0;
      lineBase._tradeRate = tradeInput?.category ? TRADE_RATES[tradeInput.category] || 0 : 0;
      lineBase._tradeCommission = lineBase._tradeMargin * lineBase._tradeRate;
      if (!tradeInput || !tradeInput.category) warnings.push({ type: "贸易订单未录入", order: r.销售订单, material: r.物料编码, person: r.业务经理, msg: `贸易订单 ${r.销售订单} 未录入毛利/品类，提成按0` });
      tradeLines.push(lineBase); allLines.push(lineBase); continue;
    }

    const isNegative = negativeSet.has(String(r.销售订单).trim());
    lineBase._isNegative = isNegative;
    let commRate, commRateSource;
    if (isNegative) { commRate = NEGATIVE_MARGIN_RATE; commRateSource = "负毛利(0.25%)"; }
    else { const pl = getProductLineRate(r.类别); commRate = pl.rate; commRateSource = `${r.类别}`; if (pl.warning) warnings.push({ type: "类别异常", order: r.销售订单, person: r.业务经理, msg: `订单 ${r.销售订单}: ${pl.warning}` }); }
    lineBase._commRate = commRate; lineBase._commRateSource = commRateSource;

    const effectivePrice = Math.min(lineBase._exFactoryPrice, price);
    lineBase._commBase = effectivePrice * kg;

    let activityCoef;
    const customerActOverride = activityMap.get(r.售达方);
    if (customerActOverride !== undefined && customerActOverride !== null && customerActOverride !== "") { activityCoef = Number(customerActOverride); lineBase._activityOverride = true; }
    else if (r.提成活跃系数 !== null && r.提成活跃系数 !== undefined && !isNaN(r.提成活跃系数)) { activityCoef = Number(r.提成活跃系数); }
    else { activityCoef = 1; }
    lineBase._activityCoef = activityCoef;
    lineBase._saleCommission = lineBase._commBase * commRate * activityCoef;

    const priceDed = priceDedMap.get(r.售达方) || 0;
    lineBase._priceDeduction = priceDed;
    lineBase._adjPrice = lineBase._exFactoryPrice - priceDed;
    if (lineBase._adjPrice > price) {
      lineBase._premiumBase = (lineBase._adjPrice - price) * kg;
      lineBase._premiumRatio = lineBase._adjPrice / price - 1;
      lineBase._premiumRate = getPremiumRate(lineBase._premiumRatio);
      lineBase._premiumCommission = lineBase._premiumBase * lineBase._premiumRate;
    } else {
      lineBase._premiumBase = 0;
      lineBase._premiumRatio = lineBase._adjPrice / price - 1;
      lineBase._premiumRate = 0; lineBase._premiumCommission = 0;
    }
    lineBase._flag = "valid";
    if (r.业务经理 && !personNames.has(r.业务经理)) warnings.push({ type: "人员未在绩效表", order: r.销售订单, person: r.业务经理, msg: `"${r.业务经理}" 不在绩效表中` });
    validRegular.push(lineBase); allLines.push(lineBase);
  }

  for (const d of priceDeductions) { const cid = String(d.customerId).trim(); if (cid && !settlements.some((s) => s.售达方 === cid)) warnings.push({ type: "居间单价无匹配", msg: `售达方 ${cid} 无匹配订单` }); }
  for (const o of negativeOrders) { const oid = String(o).trim(); if (oid && !settlements.some((s) => String(s.销售订单).trim() === oid)) warnings.push({ type: "负毛利无匹配", msg: `订单 ${oid} 无匹配` }); }

  const personResults = persons.map((p) => {
    const settings = personSettings[p.name] || {};
    const baseSalary = settings.baseSalary !== undefined ? Number(settings.baseSalary) : p.defaultBaseSalary;
    const perfBase = settings.perfBase !== undefined ? Number(settings.perfBase) : p.defaultPerfBase;
    const cappedRate = Math.min(p.completionRate, 1);
    const perfWage = (0.8 * cappedRate + 0.2 * p.perfScore) * perfBase;

    const myRegular = validRegular.filter((l) => l.业务经理 === p.name);
    const myTrade = tradeLines.filter((l) => l.业务经理 === p.name);
    const myDropped = dropped.filter((l) => l.业务经理 === p.name);
    const myAll = allLines.filter((l) => l.业务经理 === p.name);

    // ★ 无抵扣扣除
    const totalSaleCommission = myRegular.reduce((s, l) => s + l._saleCommission, 0);
    const totalPremium = myRegular.reduce((s, l) => s + l._premiumCommission, 0);
    const totalTrade = myTrade.reduce((s, l) => s + (l._tradeCommission || 0), 0);
    const totalSalary = baseSalary + perfWage + totalSaleCommission + totalPremium + totalTrade;

    const allReceipt = myAll.reduce((s, l) => s + (Number(l.含税金额) || 0), 0);
    const validReceipt = myRegular.reduce((s, l) => s + (Number(l.含税金额) || 0), 0);
    const tradeReceipt = myTrade.reduce((s, l) => s + (Number(l.含税金额) || 0), 0);
    const droppedReceipt = myDropped.reduce((s, l) => s + (Number(l.含税金额) || 0), 0);

    return { ...p, baseSalary, perfBase, perfWage, cappedRate, myRegular, myTrade, myDropped, totalSaleCommission, totalPremium, totalTrade, totalSalary, allReceipt, validReceipt, tradeReceipt, droppedReceipt };
  });

  return { warnings, personResults, allLines, validRegular, tradeLines, dropped };
}

// ============================================================
// 导出工具
// ============================================================
// ============================================================
// 个人报告 Excel 导出（HTML→XLS，单 Sheet，带完整样式）
// 用 HTML Table 渲染，Excel 可直接打开，样式完整保留
// ============================================================
function exportPersonXlsx(p, month) {
  // ── 通用样式常量 ───────────────────────────────────────────
  const S = {
    // 基础表格容器
    wrap:   `font-family:Microsoft YaHei,Arial,sans-serif;font-size:10pt;`,
    // 大标题
    title:  `background:#1e293b;color:#ffffff;font-size:14pt;font-weight:bold;padding:10px 14px;text-align:left;`,
    // 副标题（生成时间）
    sub:    `background:#334155;color:#cbd5e1;font-size:9pt;padding:4px 14px;text-align:left;`,
    // 区块标题（如"一、薪酬汇总"）
    sec:    `background:#475569;color:#ffffff;font-size:10pt;font-weight:bold;padding:6px 10px;`,
    // 普通表头
    th:     `background:#0f172a;color:#ffffff;font-weight:bold;padding:6px 8px;text-align:center;border:1px solid #475569;white-space:nowrap;`,
    // 蓝色表头（销售提成区）
    thBlue: `background:#1d4ed8;color:#ffffff;font-weight:bold;padding:6px 8px;text-align:center;border:1px solid #1e40af;white-space:nowrap;`,
    // 绿色表头（溢价奖金区）
    thGreen:`background:#059669;color:#ffffff;font-weight:bold;padding:6px 8px;text-align:center;border:1px solid #065f46;white-space:nowrap;`,
    // 琥珀色表头（贸易）
    thAmber:`background:#d97706;color:#ffffff;font-weight:bold;padding:6px 8px;text-align:center;border:1px solid #92400e;white-space:nowrap;`,
    // 普通单元格
    td:     `padding:5px 8px;border:1px solid #e2e8f0;vertical-align:middle;`,
    // 数字列（右对齐）
    tdNum:  `padding:5px 8px;border:1px solid #e2e8f0;text-align:right;vertical-align:middle;`,
    // 蓝色数据（销售提成）
    tdBlue: `padding:5px 8px;border:1px solid #bfdbfe;text-align:right;background:#eff6ff;vertical-align:middle;`,
    // 绿色数据（溢价奖金）
    tdGreen:`padding:5px 8px;border:1px solid #a7f3d0;text-align:right;background:#ecfdf5;vertical-align:middle;`,
    // 红色（负数溢价率）
    tdRed:  `padding:5px 8px;border:1px solid #fca5a5;text-align:right;background:#fef2f2;color:#dc2626;font-weight:bold;vertical-align:middle;`,
    // 合计行
    totTd:  `padding:6px 8px;border:1px solid #94a3b8;background:#f1f5f9;font-weight:bold;vertical-align:middle;`,
    totNum: `padding:6px 8px;border:1px solid #94a3b8;background:#f1f5f9;font-weight:bold;text-align:right;vertical-align:middle;`,
    totBlue:`padding:6px 8px;border:1px solid #93c5fd;background:#dbeafe;font-weight:bold;text-align:right;vertical-align:middle;`,
    totGreen:`padding:6px 8px;border:1px solid #6ee7b7;background:#d1fae5;font-weight:bold;text-align:right;vertical-align:middle;`,
    // 奇偶行底色
    even:   `background:#f8fafc;`,
    odd:    `background:#ffffff;`,
    // 标签格（汇总表左列）
    label:  `padding:6px 10px;border:1px solid #e2e8f0;font-weight:bold;background:#f8fafc;`,
    // 汇总值格
    val:    `padding:6px 10px;border:1px solid #e2e8f0;text-align:right;`,
    // 汇总合计行
    sumRow: `padding:7px 10px;border:1px solid #94a3b8;background:#0f172a;color:#ffffff;font-weight:bold;`,
    sumVal: `padding:7px 10px;border:1px solid #94a3b8;background:#0f172a;color:#ffffff;font-weight:bold;text-align:right;font-size:11pt;`,
    // 空行分隔
    gap:    `height:16px;`,
    // 公式文本
    formula:`padding:5px 12px;color:#374151;font-size:9.5pt;background:#fafafa;border:1px solid #e5e7eb;`,
  };

  const n2 = (v) => (v == null || isNaN(v) ? "" : Number(v).toFixed(2));
  const n4 = (v) => (v == null || isNaN(v) ? "" : Number(v).toFixed(4));
  const pct = (v, d=2) => (v == null || isNaN(v) ? "" : (Number(v)*100).toFixed(d)+"%");
  const cny = (v) => (v == null || isNaN(v) ? "" : Number(v).toLocaleString("zh-CN",{minimumFractionDigits:2,maximumFractionDigits:2}));

  // 计算合计
  const sumKg   = p.myRegular.reduce((s,l)=>s+l._kg, 0);
  const sumAmt  = p.myRegular.reduce((s,l)=>s+l.含税金额, 0);
  const sumSale = p.myRegular.reduce((s,l)=>s+l._saleCommission, 0);
  const sumPrem = p.myRegular.reduce((s,l)=>s+l._premiumCommission, 0);
  const sumTrd  = p.myTrade.reduce((s,l)=>s+(l._tradeCommission||0), 0);

  // ── 开始构建 HTML ──────────────────────────────────────────
  let html = `<html xmlns:o="urn:schemas-microsoft-com:office:office"
    xmlns:x="urn:schemas-microsoft-com:office:excel"
    xmlns="http://www.w3.org/TR/REC-html40">
<head><meta charset="UTF-8">
<xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet>
<x:Name>薪酬明细</x:Name>
<x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions>
</x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml>
</head><body>
<table style="${S.wrap}" cellspacing="0" cellpadding="0" border="0">`;

  // ① 大标题
  html += `<tr><td colspan="15" style="${S.title}">滨州康乃尔新材料科技有限公司 · 业务员薪酬核算明细</td></tr>`;
  html += `<tr><td colspan="15" style="${S.sub}">${p.name} · ${month} &nbsp;|&nbsp; 生成时间：${new Date().toLocaleString("zh-CN")} &nbsp;|&nbsp; 本报告依据《2026年康乃尔经营薪酬考核方案V2.0》核算</td></tr>`;

  // ② 空行
  html += `<tr><td colspan="15" style="${S.gap}"></td></tr>`;

  // ③ 一、薪酬汇总
  html += `<tr><td colspan="15" style="${S.sec}">一、月度薪酬汇总</td></tr>`;
  html += `<tr>
    <td style="${S.th}" width="130">项目</td>
    <td style="${S.th}" width="140">金额（元）</td>
    <td colspan="13" style="border:none;"></td>
  </tr>`;
  const summaryRows = [
    ["基本工资", p.baseSalary, ""],
    ["绩效工资", p.perfWage, `（完成率 ${pct(p.completionRate)} · 绩效分 ${p.perfScore.toFixed(2)} · 基数 ${p.perfBase}`],
    ["销售提成", p.totalSaleCommission, ""],
    ["溢价奖金", p.totalPremium, ""],
    ["贸易提成", p.totalTrade, ""],
  ];
  summaryRows.forEach(([label, val, note]) => {
    html += `<tr>
      <td style="${S.label}">${label}</td>
      <td style="${S.val}">${cny(val)}</td>
      <td colspan="13" style="padding:4px 10px;border:none;color:#64748b;font-size:9pt;">${note}</td>
    </tr>`;
  });
  html += `<tr>
    <td style="${S.sumRow}">合计薪酬</td>
    <td style="${S.sumVal}">${cny(p.totalSalary)}</td>
    <td colspan="13" style="border:none;"></td>
  </tr>`;

  // ④ 空行
  html += `<tr><td colspan="15" style="${S.gap}"></td></tr>`;

  // ⑤ 二、绩效工资推演
  html += `<tr><td colspan="15" style="${S.sec}">二、绩效工资推演</td></tr>`;
  html += `<tr><td colspan="15" style="${S.formula}">公式：绩效工资 = (80% × min(当月完成率, 100%) + 20% × 绩效分) × 绩效基数</td></tr>`;
  html += `<tr><td colspan="15" style="${S.formula}">
    = (80% × min(${pct(p.completionRate)}, 100%) + 20% × ${p.perfScore.toFixed(2)}) × ${p.perfBase.toFixed(0)}
    = <strong>${cny(p.perfWage)}</strong> 元
  </td></tr>`;

  // ⑥ 空行
  html += `<tr><td colspan="15" style="${S.gap}"></td></tr>`;

  // ⑦ 三、常规订单明细
  if (p.myRegular.length > 0) {
    html += `<tr><td colspan="15" style="${S.sec}">三、常规订单明细（共 ${p.myRegular.length} 笔）</td></tr>`;
    html += `<tr>
      <td style="${S.th}">订单号</td>
      <td style="${S.th}">类别</td>
      <td style="${S.th}">客户名称</td>
      <td style="${S.th}">核定公斤</td>
      <td style="${S.th}">含税金额</td>
      <td style="${S.th}">运费单价<br>(元/kg)</td>
      <td style="${S.th}">出厂售价<br>(元/kg)</td>
      <td style="${S.th}">基价<br>(元/kg)</td>
      <td style="${S.thBlue}">提成系数</td>
      <td style="${S.thBlue}">活跃度</td>
      <td style="${S.thBlue}">销售提成</td>
      <td style="${S.th}">居间单价<br>(元/kg)</td>
      <td style="${S.thGreen}">溢价率</td>
      <td style="${S.thGreen}">溢价奖金</td>
    </tr>`;
    p.myRegular.forEach((l, idx) => {
      const base = idx % 2 === 0 ? S.even : S.odd;
      const premNeg = l._premiumRatio !== null && l._premiumRatio < 0;
      html += `<tr style="${base}">
        <td style="${S.td}font-size:9pt;">${l.销售订单}</td>
        <td style="${S.td}">${l.类别}${l._isNegative ? '<br><span style="color:#d97706;font-size:8pt;">⚠负毛利</span>' : ""}</td>
        <td style="${S.td}max-width:160px;">${l.售达方描述}</td>
        <td style="${S.tdNum}">${n2(l._kg)}</td>
        <td style="${S.tdNum}">${cny(l.含税金额)}</td>
        <td style="${S.tdNum}">${n4(l._freightUnit)}</td>
        <td style="${S.tdNum}">${n4(l._exFactoryPrice)}</td>
        <td style="${S.tdNum}">${n4(l._basePrice)}</td>
        <td style="${S.tdBlue}">${pct(l._commRate)}</td>
        <td style="${S.tdBlue}">${l._activityCoef}</td>
        <td style="${S.tdBlue}font-weight:bold;">${cny(l._saleCommission)}</td>
        <td style="${S.tdNum}">${l._priceDeduction||0}</td>
        <td style="${premNeg ? S.tdRed : S.tdGreen}">${l._premiumRatio!==null ? pct(l._premiumRatio,1) : "—"}</td>
        <td style="${S.tdGreen}font-weight:bold;">${cny(l._premiumCommission)}</td>
      </tr>`;
    });
    // 合计行
    html += `<tr>
      <td style="${S.totTd}" colspan="3">合计</td>
      <td style="${S.totNum}">${n2(sumKg)}</td>
      <td style="${S.totNum}">${cny(sumAmt)}</td>
      <td colspan="3" style="border:none;"></td>
      <td colspan="2" style="border:none;"></td>
      <td style="${S.totBlue}">${cny(sumSale)}</td>
      <td style="border:none;"></td>
      <td style="border:none;"></td>
      <td style="${S.totGreen}">${cny(sumPrem)}</td>
    </tr>`;
  }

  // ⑧ 贸易订单（如有）
  if (p.myTrade.length > 0) {
    html += `<tr><td colspan="15" style="${S.gap}"></td></tr>`;
    html += `<tr><td colspan="15" style="${S.sec}">四、贸易业务明细（共 ${p.myTrade.length} 笔）</td></tr>`;
    html += `<tr>
      <td style="${S.th}">订单号</td>
      <td style="${S.th}" colspan="2">物料描述</td>
      <td style="${S.th}" colspan="2">客户名称</td>
      <td style="${S.th}">含税金额</td>
      <td style="${S.thAmber}">贸易品类</td>
      <td style="${S.thAmber}">毛利金额</td>
      <td style="${S.thAmber}">提成比例</td>
      <td style="${S.thAmber}">贸易提成</td>
      <td colspan="4" style="border:none;"></td>
    </tr>`;
    p.myTrade.forEach((l, idx) => {
      const base = idx % 2 === 0 ? S.even : S.odd;
      html += `<tr style="${base}">
        <td style="${S.td}font-size:9pt;">${l.销售订单}</td>
        <td style="${S.td}" colspan="2">${l.物料描述}</td>
        <td style="${S.td}" colspan="2">${l.售达方描述}</td>
        <td style="${S.tdNum}">${cny(l.含税金额)}</td>
        <td style="${S.td}">${l._tradeCategory||"—"}</td>
        <td style="${S.tdNum}">${cny(l._tradeMargin)}</td>
        <td style="${S.tdNum}">${l._tradeRate?pct(l._tradeRate,0):"—"}</td>
        <td style="${S.tdNum}font-weight:bold;">${cny(l._tradeCommission)}</td>
        <td colspan="4" style="border:none;"></td>
      </tr>`;
    });
    html += `<tr>
      <td style="${S.totTd}" colspan="9">贸易提成合计</td>
      <td style="${S.totNum}font-weight:bold;">${cny(sumTrd)}</td>
      <td colspan="4" style="border:none;"></td>
    </tr>`;
  }

  // ⑨ 底部说明
  html += `<tr><td colspan="15" style="${S.gap}"></td></tr>`;
  html += `<tr><td colspan="15" style="padding:6px 12px;color:#94a3b8;font-size:8.5pt;border-top:2px solid #e2e8f0;">
    本报告由康乃尔薪酬核算系统自动生成。蓝色列为销售提成相关数据，绿色列为溢价奖金相关数据。如有疑问请联系销售管理部门。
  </td></tr>`;

  html += `</table></body></html>`;

  // ── 下载为 .xls（Excel 直接支持 HTML Table 格式）──────────
  const blob = new Blob(["\uFEFF" + html], { type: "application/vnd.ms-excel;charset=utf-8" });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement("a");
  a.href     = url;
  a.download = `康乃尔薪酬_${month}_${p.name}.xls`;
  document.body.appendChild(a); a.click();
  document.body.removeChild(a); URL.revokeObjectURL(url);
}

function generateAllMarkdown(personResults, warnings, month) {
  const lines = [];
  lines.push(`# 康乃尔业务人员月度薪酬核算报告`);
  lines.push(`**核算月份**: ${month}  ·  **生成时间**: ${new Date().toLocaleString("zh-CN")}`);
  lines.push("");
  lines.push("## 薪酬汇总表");
  lines.push("| 业务员 | 基本工资 | 绩效工资 | 销售提成 | 溢价奖金 | 贸易提成 | **合计** |");
  lines.push("|---|---|---|---|---|---|---|");
  let tBase = 0, tPerf = 0, tSale = 0, tPrem = 0, tTrade = 0, tTotal = 0;
  personResults.forEach((p) => {
    tBase += p.baseSalary; tPerf += p.perfWage; tSale += p.totalSaleCommission; tPrem += p.totalPremium; tTrade += p.totalTrade; tTotal += p.totalSalary;
    lines.push(`| ${p.name} | ${fmtCNY(p.baseSalary)} | ${fmtCNY(p.perfWage)} | ${fmtCNY(p.totalSaleCommission)} | ${fmtCNY(p.totalPremium)} | ${fmtCNY(p.totalTrade)} | **${fmtCNY(p.totalSalary)}** |`);
  });
  lines.push(`| **合计** | **${fmtCNY(tBase)}** | **${fmtCNY(tPerf)}** | **${fmtCNY(tSale)}** | **${fmtCNY(tPrem)}** | **${fmtCNY(tTrade)}** | **${fmtCNY(tTotal)}** |`);
  lines.push("");
  personResults.forEach((p) => { lines.push("---"); lines.push(""); lines.push(generatePersonMarkdown(p, month)); });
  if (warnings.length > 0) { lines.push("---"); lines.push("## 异常预警"); lines.push(""); warnings.forEach((w) => lines.push(`- **[${w.type}]** ${w.msg}`)); }
  return lines.join("\n");
}

function downloadMd(content, filename) {
  const blob = new Blob([content], { type: "text/markdown;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a"); a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); document.body.removeChild(a); URL.revokeObjectURL(url);
}

// ============================================================
// UI 子组件
// ============================================================
function Stepper({ step }) {
  const steps = [
    { id: 1, label: "上传文件", icon: Upload },
    { id: 2, label: "确认参数", icon: Settings },
    { id: 3, label: "执行核算", icon: Calculator },
    { id: 4, label: "查看结果", icon: FileBarChart },
  ];
  return (
    <div className="flex items-center justify-center gap-2 md:gap-4 py-6">
      {steps.map((s, i) => {
        const Icon = s.icon;
        const active = step === s.id, done = step > s.id;
        return (
          <React.Fragment key={s.id}>
            <div className="flex items-center gap-2.5">
              <div className={`w-10 h-10 rounded-full flex items-center justify-center transition-all border-2 ${done ? "bg-emerald-600 border-emerald-600 text-white" : active ? "bg-slate-800 border-slate-800 text-white shadow-lg shadow-slate-800/30" : "bg-white border-slate-300 text-slate-400"}`}>
                {done ? <CheckCircle2 size={18} /> : <Icon size={18} />}
              </div>
              <div className={`hidden md:block text-sm font-medium ${active ? "text-slate-800" : done ? "text-emerald-600" : "text-slate-400"}`}>{s.label}</div>
            </div>
            {i < steps.length - 1 && <div className={`h-0.5 w-6 md:w-12 ${step > s.id ? "bg-emerald-600" : "bg-slate-200"}`} />}
          </React.Fragment>
        );
      })}
    </div>
  );
}

// ============================================================
// 步骤 1：上传
// ============================================================
function Step1Upload({ onComplete }) {
  const [f1, setF1] = useState(null);
  const [f2, setF2] = useState(null);
  const [parsing, setParsing] = useState(false);
  const [error, setError] = useState(null);
  const handleParse = async () => {
    if (!f1 || !f2) { setError("请同时上传两份 Excel"); return; }
    setParsing(true); setError(null);
    try { onComplete({ settlementData: await parseSettlementWorkbook(f1), perfData: await parsePerformanceWorkbook(f2) }); }
    catch (e) { setError("解析失败：" + (e.message || e)); }
    finally { setParsing(false); }
  };
  const FileBox = ({ label, file, onChange, hint }) => (
    <label className="block cursor-pointer">
      <div className={`border-2 border-dashed rounded-2xl p-7 transition-all ${file ? "border-emerald-500 bg-emerald-50/50" : "border-slate-300 hover:border-slate-500 hover:bg-slate-50"}`}>
        <div className="flex items-start gap-4">
          <div className={`w-14 h-14 rounded-xl flex items-center justify-center shrink-0 ${file ? "bg-emerald-600 text-white" : "bg-slate-100 text-slate-500"}`}><FileSpreadsheet size={26} /></div>
          <div className="flex-1 min-w-0">
            <div className="font-semibold text-slate-900 mb-1">{label}</div>
            <div className="text-xs text-slate-500 mb-2">{hint}</div>
            {file ? <div className="text-sm text-emerald-700 font-medium truncate">✓ {file.name}</div> : <div className="text-sm text-slate-400">点击或拖拽选择 .xlsx 文件</div>}
          </div>
        </div>
        <input type="file" accept=".xlsx,.xls" className="hidden" onChange={(e) => onChange(e.target.files?.[0] || null)} />
      </div>
    </label>
  );
  return (
    <div className="max-w-3xl mx-auto px-4">
      <div className="text-center mb-8"><h2 className="text-2xl font-bold text-slate-900 mb-2">步骤一 · 上传数据文件</h2><p className="text-slate-500">请上传当月的回款明细表与绩效考核表</p></div>
      <div className="space-y-4">
        <FileBox label="回款明细表" hint="包含「回款明细」和「基价表」Sheet" file={f1} onChange={setF1} />
        <FileBox label="绩效考核表" hint="包含人员姓名、业绩完成率、绩效分等" file={f2} onChange={setF2} />
      </div>
      {error && <div className="mt-5 p-4 bg-red-50 border border-red-200 text-red-700 rounded-xl text-sm flex gap-2"><AlertTriangle size={18} className="shrink-0 mt-0.5" />{error}</div>}
      <button onClick={handleParse} disabled={parsing || !f1 || !f2} className="w-full mt-6 py-4 bg-slate-900 text-white rounded-xl font-semibold hover:bg-slate-800 disabled:bg-slate-300 disabled:cursor-not-allowed transition-all flex items-center justify-center gap-2">
        {parsing ? "解析中..." : <>下一步 · 确认参数 <ChevronRight size={18} /></>}
      </button>
    </div>
  );
}

// ============================================================
// 步骤 2：参数确认（5 个 Tab，无抵扣）
// ============================================================
function Step2Parameters({ data, onBack, onComplete }) {
  const { settlementData, perfData } = data;
  const { settlements } = settlementData;
  const { persons } = perfData;
  const [activeTab, setActiveTab] = useState("salary");

  const [personSettings, setPersonSettings] = useState(() => {
    const init = {};
    persons.forEach((p) => { init[p.name] = { baseSalary: p.defaultBaseSalary, perfBase: p.defaultPerfBase }; });
    return init;
  });

  const [priceDeductions, setPriceDeductions] = useState([]);
  const [storageReady, setStorageReady] = useState(false);
  useEffect(() => { try { const raw = window.localStorage?.getItem(PRICE_DEDUCTIONS_KEY); if (raw) { const p = JSON.parse(raw); if (Array.isArray(p)) setPriceDeductions(p); } } catch {} finally { setStorageReady(true); } }, []);
  useEffect(() => { if (!storageReady) return; try { window.localStorage?.setItem(PRICE_DEDUCTIONS_KEY, JSON.stringify(priceDeductions)); } catch {} }, [priceDeductions, storageReady]);

  const [negativeOrders, setNegativeOrders] = useState([]);
  const tradeOrdersRaw = useMemo(() => settlements.filter((s) => s.类别 === "贸易" && s.物料编码), [settlements]);
  const [tradeMarginInputs, setTradeMarginInputs] = useState({});
  const customerActivityList = useMemo(() => {
    const map = new Map();
    for (const s of settlements) { if (!s.售达方 || !s.物料编码) continue; if (!map.has(s.售达方)) map.set(s.售达方, { customerId: s.售达方, name: s.售达方描述, defaultCoef: s.提成活跃系数 ?? 1, orderCount: 0 }); map.get(s.售达方).orderCount++; }
    return Array.from(map.values());
  }, [settlements]);
  const [activityOverrides, setActivityOverrides] = useState({});

  const updatePS = (name, key, val) => setPersonSettings((p) => ({ ...p, [name]: { ...p[name], [key]: val === "" ? "" : val } }));
  const restorePS = (name) => { const p = persons.find((x) => x.name === name); if (p) setPersonSettings((prev) => ({ ...prev, [name]: { baseSalary: p.defaultBaseSalary, perfBase: p.defaultPerfBase } })); };

  const tabs = [
    { id: "salary", label: "工资基数", icon: Coins, count: persons.length },
    { id: "interim", label: "居间单价", icon: Briefcase, count: priceDeductions.length },
    { id: "negative", label: "负毛利订单", icon: FileWarning, count: negativeOrders.length },
    { id: "trade", label: "贸易毛利", icon: TrendingUp, count: tradeOrdersRaw.length },
    { id: "activity", label: "活跃度系数", icon: Activity, count: customerActivityList.length },
  ];

  const handleSubmit = () => onComplete({ personSettings, priceDeductions: priceDeductions.filter((d) => d.customerId?.trim()), negativeOrders: negativeOrders.filter((o) => String(o).trim()), tradeMarginInputs, activityOverrides });

  return (
    <div className="max-w-7xl mx-auto px-4 pb-12">
      <div className="text-center mb-6"><h2 className="text-2xl font-bold text-slate-900 mb-1">步骤二 · 确认 / 修改参数</h2><p className="text-slate-500 text-sm">默认值已从文件读取，可逐项调整。居间单价跨会话自动保留。</p></div>
      <div className="flex flex-wrap gap-2 border-b border-slate-200 mb-6">
        {tabs.map((t) => { const Icon = t.icon; const active = activeTab === t.id; return (
          <button key={t.id} onClick={() => setActiveTab(t.id)} className={`px-4 py-2.5 -mb-px border-b-2 font-medium text-sm flex items-center gap-2 transition-colors ${active ? "border-slate-900 text-slate-900" : "border-transparent text-slate-500 hover:text-slate-700"}`}>
            <Icon size={16} />{t.label}<span className={`px-1.5 py-0.5 rounded-md text-xs ${active ? "bg-slate-900 text-white" : "bg-slate-100 text-slate-500"}`}>{t.count}</span>
          </button>); })}
      </div>
      <div className="bg-white rounded-2xl border border-slate-200 p-5 md:p-6 shadow-sm">
        {activeTab === "salary" && (
          <div>
            <h3 className="font-semibold text-slate-900 mb-1">基本工资 与 绩效工资基数</h3>
            <p className="text-sm text-slate-500 mb-4">默认值取自绩效表。绩效工资基数按等级：一级=2000、二级=3000、三级=4000。</p>
            <div className="overflow-x-auto"><table className="w-full text-sm"><thead><tr className="text-left bg-slate-50 border-y border-slate-200">
              <th className="px-3 py-2.5 font-medium">人员</th><th className="px-3 py-2.5 font-medium">等级</th><th className="px-3 py-2.5 font-medium">基本工资</th><th className="px-3 py-2.5 font-medium">绩效基数</th><th className="px-3 py-2.5 font-medium">完成率</th><th className="px-3 py-2.5 font-medium">绩效分</th><th className="px-3 py-2.5 font-medium">操作</th>
            </tr></thead><tbody>
              {persons.map((p) => (
                <tr key={p.name} className="border-b border-slate-100 hover:bg-slate-50/50">
                  <td className="px-3 py-2.5 font-semibold text-slate-900">{p.name}</td>
                  <td className="px-3 py-2.5 text-slate-600">{p.level}</td>
                  <td className="px-3 py-2.5"><input type="number" value={personSettings[p.name].baseSalary} onChange={(e) => updatePS(p.name, "baseSalary", e.target.value)} className="w-28 px-2 py-1 border border-slate-300 rounded-md text-sm focus:outline-none focus:ring-2 focus:ring-slate-400" /></td>
                  <td className="px-3 py-2.5"><input type="number" value={personSettings[p.name].perfBase} onChange={(e) => updatePS(p.name, "perfBase", e.target.value)} className="w-28 px-2 py-1 border border-slate-300 rounded-md text-sm focus:outline-none focus:ring-2 focus:ring-slate-400" /></td>
                  <td className="px-3 py-2.5 text-slate-600">{fmtPct(p.completionRate)}</td>
                  <td className="px-3 py-2.5 text-slate-600">{fmtNum(p.perfScore)}</td>
                  <td className="px-3 py-2.5"><button onClick={() => restorePS(p.name)} className="text-slate-500 hover:text-slate-900 inline-flex items-center gap-1 text-xs"><RotateCcw size={12} /> 恢复</button></td>
                </tr>))}
            </tbody></table></div>
          </div>)}

        {activeTab === "interim" && (
          <div>
            <div className="flex items-start justify-between mb-1"><h3 className="font-semibold text-slate-900">溢价奖金扣除项目（居间单价）</h3><span className="text-xs text-emerald-600 bg-emerald-50 px-2 py-1 rounded-md">✓ 自动持久化</span></div>
            <p className="text-sm text-slate-500 mb-4">针对特定客户从出厂售价中扣除的费用单价（元/kg），仅影响溢价奖金。</p>
            <div className="space-y-2 mb-3">
              {priceDeductions.length === 0 && <div className="text-center py-8 text-slate-400 text-sm bg-slate-50/50 rounded-lg border border-dashed border-slate-200">尚未录入</div>}
              {priceDeductions.map((d, i) => (
                <div key={i} className="flex items-center gap-2">
                  <input placeholder="售达方编号" value={d.customerId} onChange={(e) => { const arr = [...priceDeductions]; arr[i] = { ...arr[i], customerId: e.target.value }; setPriceDeductions(arr); }} className="flex-1 px-3 py-2 border border-slate-300 rounded-md text-sm" />
                  <input type="number" step="0.01" placeholder="元/kg" value={d.pricePerKg} onChange={(e) => { const arr = [...priceDeductions]; arr[i] = { ...arr[i], pricePerKg: e.target.value }; setPriceDeductions(arr); }} className="w-44 px-3 py-2 border border-slate-300 rounded-md text-sm" />
                  <button onClick={() => setPriceDeductions((p) => p.filter((_, idx) => idx !== i))} className="p-2 text-slate-400 hover:text-red-600 hover:bg-red-50 rounded-md"><Trash2 size={16} /></button>
                </div>))}
            </div>
            <button onClick={() => setPriceDeductions((p) => [...p, { customerId: "", pricePerKg: "" }])} className="text-sm px-3 py-2 border border-dashed border-slate-400 rounded-md hover:bg-slate-50 inline-flex items-center gap-1.5 text-slate-700"><Plus size={14} /> 添加</button>
          </div>)}

        {activeTab === "negative" && (
          <div>
            <h3 className="font-semibold text-slate-900 mb-1">负毛利订单</h3>
            <p className="text-sm text-slate-500 mb-4">被标记的订单统一按 <b>0.25%</b> 计算销售提成系数。</p>
            <div className="space-y-2 mb-3">
              {negativeOrders.length === 0 && <div className="text-center py-8 text-slate-400 text-sm bg-slate-50/50 rounded-lg border border-dashed border-slate-200">尚未标记</div>}
              {negativeOrders.map((o, i) => (
                <div key={i} className="flex items-center gap-2">
                  <input placeholder="订单编号" value={o} onChange={(e) => { const arr = [...negativeOrders]; arr[i] = e.target.value; setNegativeOrders(arr); }} className="flex-1 px-3 py-2 border border-slate-300 rounded-md text-sm" />
                  <button onClick={() => setNegativeOrders((n) => n.filter((_, idx) => idx !== i))} className="p-2 text-slate-400 hover:text-red-600 hover:bg-red-50 rounded-md"><Trash2 size={16} /></button>
                </div>))}
            </div>
            <button onClick={() => setNegativeOrders((n) => [...n, ""])} className="text-sm px-3 py-2 border border-dashed border-slate-400 rounded-md hover:bg-slate-50 inline-flex items-center gap-1.5 text-slate-700"><Plus size={14} /> 添加</button>
          </div>)}

        {activeTab === "trade" && (
          <div>
            <h3 className="font-semibold text-slate-900 mb-1">贸易业务毛利录入</h3>
            <p className="text-sm text-slate-500 mb-4">提成比例：熔喷料 40% / 母粒 30% / 无纺布 20%。</p>
            {tradeOrdersRaw.length === 0 ? <div className="text-center py-10 text-slate-400 text-sm bg-slate-50/50 rounded-lg border border-dashed border-slate-200">本月无贸易订单</div> : (
              <div className="overflow-x-auto"><table className="w-full text-sm"><thead><tr className="text-left bg-slate-50 border-y border-slate-200">
                <th className="px-3 py-2.5 font-medium">订单号</th><th className="px-3 py-2.5 font-medium">业务员</th><th className="px-3 py-2.5 font-medium">物料</th><th className="px-3 py-2.5 font-medium">含税金额</th><th className="px-3 py-2.5 font-medium">品类</th><th className="px-3 py-2.5 font-medium">毛利（元）</th>
              </tr></thead><tbody>
                {tradeOrdersRaw.map((o) => { const key = `${o.销售订单}_${o.物料编码}`; const cur = tradeMarginInputs[key] || {}; return (
                  <tr key={key} className="border-b border-slate-100 hover:bg-slate-50/50">
                    <td className="px-3 py-2.5 font-mono text-xs">{o.销售订单}</td><td className="px-3 py-2.5">{o.业务经理}</td><td className="px-3 py-2.5 text-xs max-w-xs truncate">{o.物料描述}</td><td className="px-3 py-2.5">{fmtCNY(o.含税金额)}</td>
                    <td className="px-3 py-2.5"><select value={cur.category || ""} onChange={(e) => setTradeMarginInputs((p) => ({ ...p, [key]: { ...p[key], category: e.target.value } }))} className="px-2 py-1 border border-slate-300 rounded-md text-sm bg-white"><option value="">-选择-</option><option value="熔喷料">熔喷料(40%)</option><option value="母粒">母粒(30%)</option><option value="无纺布">无纺布(20%)</option></select></td>
                    <td className="px-3 py-2.5"><input type="number" step="0.01" value={cur.margin || ""} onChange={(e) => setTradeMarginInputs((p) => ({ ...p, [key]: { ...p[key], margin: e.target.value } }))} className="w-32 px-2 py-1 border border-slate-300 rounded-md text-sm" /></td>
                  </tr>); })}
              </tbody></table></div>)}
          </div>)}

        {activeTab === "activity" && (
          <div>
            <h3 className="font-semibold text-slate-900 mb-1">提成活跃度系数（按售达方）</h3>
            <p className="text-sm text-slate-500 mb-4">仅适用于销售提成，溢价奖金不适用。</p>
            <div className="overflow-x-auto max-h-[60vh] overflow-y-auto"><table className="w-full text-sm"><thead className="sticky top-0 bg-white"><tr className="text-left bg-slate-50 border-y border-slate-200">
              <th className="px-3 py-2.5 font-medium">售达方</th><th className="px-3 py-2.5 font-medium">客户名称</th><th className="px-3 py-2.5 font-medium">笔数</th><th className="px-3 py-2.5 font-medium">默认</th><th className="px-3 py-2.5 font-medium">当前</th><th className="px-3 py-2.5 font-medium">操作</th>
            </tr></thead><tbody>
              {customerActivityList.map((c) => { const cur = activityOverrides[c.customerId]; const display = cur !== undefined && cur !== "" ? cur : c.defaultCoef; return (
                <tr key={c.customerId} className="border-b border-slate-100 hover:bg-slate-50/50">
                  <td className="px-3 py-2.5 font-mono text-xs">{c.customerId}</td><td className="px-3 py-2.5 max-w-xs truncate">{c.name}</td><td className="px-3 py-2.5 text-slate-500">{c.orderCount}</td><td className="px-3 py-2.5 text-slate-400">{c.defaultCoef}</td>
                  <td className="px-3 py-2.5"><input type="number" step="0.1" value={display} onChange={(e) => setActivityOverrides((p) => ({ ...p, [c.customerId]: e.target.value }))} className="w-24 px-2 py-1 border border-slate-300 rounded-md text-sm" /></td>
                  <td className="px-3 py-2.5"><button onClick={() => setActivityOverrides((p) => { const np = { ...p }; delete np[c.customerId]; return np; })} className="text-slate-500 hover:text-slate-900 text-xs inline-flex items-center gap-1"><RotateCcw size={12} /> 恢复</button></td>
                </tr>); })}
            </tbody></table></div>
          </div>)}
      </div>
      <div className="flex justify-between mt-6">
        <button onClick={onBack} className="px-5 py-3 bg-white border border-slate-300 text-slate-700 rounded-xl font-medium hover:bg-slate-50 flex items-center gap-2"><ChevronLeft size={18} /> 上一步</button>
        <button onClick={handleSubmit} className="px-6 py-3 bg-slate-900 text-white rounded-xl font-semibold hover:bg-slate-800 flex items-center gap-2"><Calculator size={18} /> 开始核算</button>
      </div>
    </div>
  );
}

// ============================================================
// 步骤 4：结果展示
// ============================================================
function ResultPanel({ result, params, data, onBack, onRestart }) {
  const { personResults, warnings, validRegular, tradeLines, dropped } = result;
  const [activeReport, setActiveReport] = useState("summary");
  const [expandedPerson, setExpandedPerson] = useState(null);

  const totals = personResults.reduce((a, p) => ({ base: a.base + p.baseSalary, perf: a.perf + p.perfWage, sale: a.sale + p.totalSaleCommission, premium: a.premium + p.totalPremium, trade: a.trade + p.totalTrade, total: a.total + p.totalSalary }), { base: 0, perf: 0, sale: 0, premium: 0, trade: 0, total: 0 });
  const month = data.settlementData.month || "";

  // 导出下拉：个人 → Excel(.xlsx)，全员 → Markdown(.md)
  const [exportTarget, setExportTarget] = useState("all");
  const handleExport = useCallback(() => {
    if (exportTarget === "all") {
      downloadMd(generateAllMarkdown(personResults, warnings, month), `康乃尔薪酬核算_${month}_全员.md`);
    } else {
      const p = personResults.find((x) => x.name === exportTarget);
      if (p) exportPersonXlsx(p, month);
    }
  }, [exportTarget, personResults, warnings, month]);

  // CSV 汇总
  const exportCSV = useCallback(() => {
    const rows = [["业务员", "基本工资", "绩效工资", "销售提成", "溢价奖金", "贸易提成", "合计"]];
    personResults.forEach((p) => rows.push([p.name, p.baseSalary.toFixed(2), p.perfWage.toFixed(2), p.totalSaleCommission.toFixed(2), p.totalPremium.toFixed(2), p.totalTrade.toFixed(2), p.totalSalary.toFixed(2)]));
    rows.push(["合计", totals.base.toFixed(2), totals.perf.toFixed(2), totals.sale.toFixed(2), totals.premium.toFixed(2), totals.trade.toFixed(2), totals.total.toFixed(2)]);
    const csv = "\uFEFF" + rows.map((r) => r.map((c) => `"${c}"`).join(",")).join("\r\n");
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a"); a.href = url; a.download = `康乃尔薪酬汇总_${month}.csv`;
    document.body.appendChild(a); a.click(); document.body.removeChild(a); URL.revokeObjectURL(url);
  }, [personResults, totals, month]);

  // 色彩：销售提成=浅蓝 溢价奖金=浅绿 溢价率负=红
  const saleColor = "bg-blue-50 text-blue-900";
  const premiumColor = "bg-emerald-50 text-emerald-900";
  const negPremColor = "text-red-600 font-semibold";

  return (
    <div className="max-w-7xl mx-auto px-4 pb-12">
      <div className="text-center mb-6">
        <div className="inline-flex items-center gap-2 mb-2"><CheckCircle2 className="text-emerald-600" size={22} /><h2 className="text-2xl font-bold text-slate-900">核算完成 · {month}</h2></div>
        <p className="text-slate-500 text-sm">共 {personResults.length} 名业务员，{validRegular.length + tradeLines.length} 笔有效订单</p>
      </div>

      {/* KPI */}
      <div className="grid grid-cols-2 md:grid-cols-6 gap-3 mb-6">
        {[{ label: "薪酬总额", value: totals.total, accent: "bg-slate-900 text-white" }, { label: "基本工资", value: totals.base }, { label: "绩效工资", value: totals.perf }, { label: "销售提成", value: totals.sale, cls: "bg-blue-50 border-blue-200" }, { label: "溢价奖金", value: totals.premium, cls: "bg-emerald-50 border-emerald-200" }, { label: "贸易提成", value: totals.trade }].map((k, i) => (
          <div key={i} className={`rounded-xl p-4 border ${k.accent || k.cls || "bg-white border-slate-200"}`}>
            <div className={`text-xs font-medium ${k.accent ? "text-slate-300" : "text-slate-500"}`}>{k.label}</div>
            <div className="text-lg md:text-xl font-bold mt-1">¥ {fmtNum(k.value, 2)}</div>
          </div>))}
      </div>

      {/* Report tabs */}
      <div className="flex flex-wrap gap-2 border-b border-slate-200 mb-6">
        {[{ id: "summary", label: "薪酬汇总", icon: Coins }, { id: "validation", label: "交叉验证", icon: CheckCircle2 }, { id: "details", label: "逐人详细", icon: Users }, { id: "warnings", label: `异常预警 (${warnings.length})`, icon: AlertTriangle }].map((t) => {
          const Icon = t.icon; const active = activeReport === t.id;
          return (<button key={t.id} onClick={() => setActiveReport(t.id)} className={`px-4 py-2.5 -mb-px border-b-2 font-medium text-sm flex items-center gap-2 ${active ? "border-slate-900 text-slate-900" : "border-transparent text-slate-500 hover:text-slate-700"}`}><Icon size={16} />{t.label}</button>);
        })}
      </div>

      {/* Summary */}
      {activeReport === "summary" && (
        <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden">
          <div className="overflow-x-auto"><table className="w-full text-sm"><thead><tr className="bg-slate-50 border-b border-slate-200">
            <th className="px-4 py-3 text-left font-semibold">业务员</th><th className="px-4 py-3 text-right font-semibold">基本工资</th><th className="px-4 py-3 text-right font-semibold">绩效工资</th>
            <th className={`px-4 py-3 text-right font-semibold ${saleColor}`}>销售提成</th><th className={`px-4 py-3 text-right font-semibold ${premiumColor}`}>溢价奖金</th>
            <th className="px-4 py-3 text-right font-semibold">贸易提成</th><th className="px-4 py-3 text-right font-semibold bg-slate-100">合计</th>
          </tr></thead><tbody>
            {personResults.map((p) => (
              <tr key={p.name} className="border-b border-slate-100 hover:bg-slate-50/50">
                <td className="px-4 py-3 font-semibold">{p.name}</td><td className="px-4 py-3 text-right">{fmtCNY(p.baseSalary)}</td><td className="px-4 py-3 text-right">{fmtCNY(p.perfWage)}</td>
                <td className={`px-4 py-3 text-right ${saleColor}`}>{fmtCNY(p.totalSaleCommission)}</td><td className={`px-4 py-3 text-right ${premiumColor}`}>{fmtCNY(p.totalPremium)}</td>
                <td className="px-4 py-3 text-right">{fmtCNY(p.totalTrade)}</td><td className="px-4 py-3 text-right font-bold bg-slate-50">{fmtCNY(p.totalSalary)}</td>
              </tr>))}
            <tr className="bg-slate-900 text-white font-semibold">
              <td className="px-4 py-3">合计</td><td className="px-4 py-3 text-right">{fmtCNY(totals.base)}</td><td className="px-4 py-3 text-right">{fmtCNY(totals.perf)}</td>
              <td className="px-4 py-3 text-right">{fmtCNY(totals.sale)}</td><td className="px-4 py-3 text-right">{fmtCNY(totals.premium)}</td>
              <td className="px-4 py-3 text-right">{fmtCNY(totals.trade)}</td><td className="px-4 py-3 text-right">{fmtCNY(totals.total)}</td>
            </tr>
          </tbody></table></div>
          <div className="p-4 flex flex-wrap items-center gap-2 bg-slate-50/50 border-t border-slate-200">
            <button onClick={exportCSV} className="px-4 py-2 bg-white border border-slate-300 rounded-lg text-sm font-medium hover:bg-slate-50 inline-flex items-center gap-1.5"><Download size={14} /> CSV 汇总</button>
            <div className="flex items-center gap-1.5 ml-auto">
              <select value={exportTarget} onChange={(e) => setExportTarget(e.target.value)} className="px-3 py-2 border border-slate-300 rounded-lg text-sm bg-white">
                <option value="all">全员汇总报告</option>
                {personResults.map((p) => <option key={p.name} value={p.name}>{p.name} 个人报告</option>)}
              </select>
              <button onClick={handleExport} className="px-4 py-2 bg-slate-900 text-white rounded-lg text-sm font-medium hover:bg-slate-800 inline-flex items-center gap-1.5"><Download size={14} /> {exportTarget === "all" ? "导出 Markdown（全员）" : "导出 Excel（个人）"}</button>
            </div>
          </div>
        </div>)}

      {/* Validation */}
      {activeReport === "validation" && (
        <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden">
          <div className="px-5 py-3 border-b border-slate-200 bg-slate-50/50"><p className="text-xs text-slate-600">单位：万元。<b>全部含税金额 = 常规 + 贸易 + 剔除</b></p></div>
          <div className="overflow-x-auto"><table className="w-full text-sm"><thead><tr className="bg-slate-50 border-b border-slate-200">
            <th className="px-4 py-3 text-left font-semibold">业务员</th><th className="px-4 py-3 text-right font-semibold">全部</th><th className="px-4 py-3 text-right font-semibold">常规</th><th className="px-4 py-3 text-right font-semibold">贸易</th><th className="px-4 py-3 text-right font-semibold">剔除</th><th className="px-4 py-3 text-right font-semibold">差异</th><th className="px-4 py-3 text-center font-semibold">校验</th>
          </tr></thead><tbody>
            {personResults.map((p) => { const sum = p.validReceipt + p.tradeReceipt + p.droppedReceipt; const diff = p.allReceipt - sum; const ok = Math.abs(diff) < 0.01; return (
              <tr key={p.name} className="border-b border-slate-100">
                <td className="px-4 py-3 font-semibold">{p.name}</td><td className="px-4 py-3 text-right tabular-nums">{fmtWan(p.allReceipt)}</td><td className="px-4 py-3 text-right tabular-nums">{fmtWan(p.validReceipt)}</td><td className="px-4 py-3 text-right tabular-nums text-amber-700">{fmtWan(p.tradeReceipt)}</td><td className="px-4 py-3 text-right tabular-nums text-slate-400">{fmtWan(p.droppedReceipt)}</td><td className="px-4 py-3 text-right tabular-nums">{fmtWan(diff)}</td>
                <td className="px-4 py-3 text-center">{ok ? <CheckCircle2 className="inline text-emerald-600" size={16} /> : <AlertTriangle className="inline text-amber-600" size={16} />}</td>
              </tr>); })}
          </tbody></table></div>
        </div>)}

      {/* Details */}
      {activeReport === "details" && (
        <div className="space-y-3">
          {personResults.map((p) => {
            const expanded = expandedPerson === p.name;
            return (
              <div key={p.name} className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden">
                <button onClick={() => setExpandedPerson(expanded ? null : p.name)} className="w-full px-5 py-4 flex items-center justify-between hover:bg-slate-50 text-left">
                  <div className="flex items-center gap-4">
                    <div className="w-10 h-10 rounded-full bg-slate-100 text-slate-700 flex items-center justify-center font-semibold">{p.name.slice(0, 1)}</div>
                    <div><div className="font-semibold text-slate-900">{p.name}</div><div className="text-xs text-slate-500">{p.level} · 完成率 {fmtPct(p.completionRate)} · 绩效分 {fmtNum(p.perfScore)} · 常规{p.myRegular.length}笔 · 贸易{p.myTrade.length}笔</div></div>
                  </div>
                  <div className="flex items-center gap-4">
                    <div className="text-right"><div className="text-xs text-slate-500">月度薪酬</div><div className="text-xl font-bold text-slate-900">¥ {fmtNum(p.totalSalary)}</div></div>
                    <ChevronRight size={18} className={`text-slate-400 transition-transform ${expanded ? "rotate-90" : ""}`} />
                  </div>
                </button>
                {expanded && (
                  <div className="border-t border-slate-200 p-5 space-y-5 bg-slate-50/30">
                    {/* 薪酬构成 */}
                    <div className="grid grid-cols-2 md:grid-cols-5 gap-2">
                      {[{ label: "基本工资", v: p.baseSalary }, { label: "绩效工资", v: p.perfWage }, { label: "销售提成", v: p.totalSaleCommission, cls: "border-blue-200 bg-blue-50" }, { label: "溢价奖金", v: p.totalPremium, cls: "border-emerald-200 bg-emerald-50" }, { label: "贸易提成", v: p.totalTrade }].map((x, i) => (
                        <div key={i} className={`border rounded-lg p-2.5 ${x.cls || "border-slate-200 bg-white"}`}><div className="text-xs text-slate-500">{x.label}</div><div className="font-semibold text-slate-900 text-sm">¥ {fmtNum(x.v)}</div></div>))}
                    </div>
                    {/* 绩效推演 */}
                    <div className="bg-white rounded-lg border border-slate-200 p-4">
                      <h4 className="text-sm font-semibold text-slate-900 mb-2">绩效工资推演</h4>
                      <div className="text-sm text-slate-700 font-mono">(80% × min({fmtPct(p.completionRate)}, 100%) + 20% × {fmtNum(p.perfScore)}) × {fmtCNY(p.perfBase)} = <b>{fmtCNY(p.perfWage)}</b></div>
                    </div>
                    {/* 常规订单逐笔 */}
                    {p.myRegular.length > 0 && (
                      <div className="bg-white rounded-lg border border-slate-200 overflow-hidden">
                        <div className="px-4 py-2.5 bg-slate-50 border-b border-slate-200 text-sm font-semibold">常规订单明细 ({p.myRegular.length} 笔)</div>
                        <div className="overflow-x-auto"><table className="w-full text-xs"><thead className="bg-slate-50/50"><tr className="text-left text-slate-600">
                          <th className="px-2 py-2">订单号</th><th className="px-2 py-2">类别</th><th className="px-2 py-2">售达方</th><th className="px-2 py-2 text-right">核定kg</th><th className="px-2 py-2 text-right">含税金额</th><th className="px-2 py-2 text-right">运费/kg</th><th className="px-2 py-2 text-right">出厂价</th><th className="px-2 py-2 text-right">基价</th>
                          <th className={`px-2 py-2 text-right ${saleColor}`}>系数</th><th className="px-2 py-2 text-right">活跃度</th><th className={`px-2 py-2 text-right ${saleColor}`}>销售提成</th>
                          <th className="px-2 py-2 text-right">居间</th><th className={`px-2 py-2 text-right`}>溢价率</th><th className={`px-2 py-2 text-right ${premiumColor}`}>溢价奖金</th>
                        </tr></thead><tbody>
                          {(() => { let sumSale = 0, sumPrem = 0, sumAmt = 0, sumKg = 0; return (<>
                            {p.myRegular.map((l) => { sumSale += l._saleCommission; sumPrem += l._premiumCommission; sumAmt += l.含税金额; sumKg += l._kg;
                              const premRatioNeg = l._premiumRatio !== null && l._premiumRatio < 0;
                              return (
                              <tr key={l._rowIdx} className={`border-t border-slate-100 ${l._isNegative ? "bg-amber-50/40" : ""}`}>
                                <td className="px-2 py-1.5 font-mono">{l.销售订单}</td>
                                <td className="px-2 py-1.5">{l.类别}{l._isNegative && <span className="ml-1 text-amber-700">⚠</span>}</td>
                                <td className="px-2 py-1.5 max-w-[120px] truncate" title={l.售达方描述}>{l.售达方描述}</td>
                                <td className="px-2 py-1.5 text-right tabular-nums">{fmtNum(l._kg, 1)}</td>
                                <td className="px-2 py-1.5 text-right tabular-nums">{fmtNum(l.含税金额)}</td>
                                <td className="px-2 py-1.5 text-right tabular-nums">{fmtNum(l._freightUnit)}</td>
                                <td className="px-2 py-1.5 text-right tabular-nums">{fmtNum(l._exFactoryPrice)}</td>
                                <td className="px-2 py-1.5 text-right tabular-nums">{fmtNum(l._basePrice)}</td>
                                <td className={`px-2 py-1.5 text-right tabular-nums ${saleColor}`}>{fmtPct(l._commRate)}</td>
                                <td className="px-2 py-1.5 text-right tabular-nums">{l._activityCoef}</td>
                                <td className={`px-2 py-1.5 text-right tabular-nums font-semibold ${saleColor}`}>{fmtNum(l._saleCommission)}</td>
                                <td className="px-2 py-1.5 text-right tabular-nums">{l._priceDeduction || 0}</td>
                                <td className={`px-2 py-1.5 text-right tabular-nums ${premRatioNeg ? negPremColor : ""}`}>{l._premiumRatio !== null ? fmtPct(l._premiumRatio, 1) : "—"}</td>
                                <td className={`px-2 py-1.5 text-right tabular-nums font-semibold ${premiumColor}`}>{fmtNum(l._premiumCommission)}</td>
                              </tr>); })}
                            <tr className="border-t-2 border-slate-300 bg-slate-100 font-semibold text-xs">
                              <td className="px-2 py-2" colSpan={3}>合计</td>
                              <td className="px-2 py-2 text-right tabular-nums">{fmtNum(sumKg, 1)}</td>
                              <td className="px-2 py-2 text-right tabular-nums">{fmtNum(sumAmt)}</td>
                              <td colSpan={5}></td>
                              <td className={`px-2 py-2 text-right tabular-nums ${saleColor}`}>{fmtNum(sumSale)}</td>
                              <td colSpan={2}></td>
                              <td className={`px-2 py-2 text-right tabular-nums ${premiumColor}`}>{fmtNum(sumPrem)}</td>
                            </tr>
                          </>); })()}
                        </tbody></table></div>
                      </div>)}
                    {/* 贸易订单 */}
                    {p.myTrade.length > 0 && (
                      <div className="bg-white rounded-lg border border-slate-200 overflow-hidden">
                        <div className="px-4 py-2.5 bg-amber-50 border-b border-amber-200 text-sm font-semibold text-amber-900">贸易订单 ({p.myTrade.length} 笔)</div>
                        <table className="w-full text-xs"><thead className="bg-slate-50"><tr className="text-left text-slate-600">
                          <th className="px-3 py-2">订单号</th><th className="px-3 py-2">物料</th><th className="px-3 py-2 text-right">含税金额</th><th className="px-3 py-2">品类</th><th className="px-3 py-2 text-right">毛利</th><th className="px-3 py-2 text-right">比例</th><th className="px-3 py-2 text-right">提成</th>
                        </tr></thead><tbody>
                          {(() => { let sumT = 0; return (<>
                            {p.myTrade.map((l) => { sumT += l._tradeCommission || 0; return (
                              <tr key={l._rowIdx} className="border-t border-slate-100">
                                <td className="px-3 py-1.5 font-mono">{l.销售订单}</td><td className="px-3 py-1.5 max-w-[200px] truncate">{l.物料描述}</td><td className="px-3 py-1.5 text-right tabular-nums">{fmtCNY(l.含税金额)}</td>
                                <td className="px-3 py-1.5">{l._tradeCategory || "—"}</td><td className="px-3 py-1.5 text-right tabular-nums">{fmtCNY(l._tradeMargin)}</td><td className="px-3 py-1.5 text-right tabular-nums">{l._tradeRate ? fmtPct(l._tradeRate, 0) : "—"}</td>
                                <td className="px-3 py-1.5 text-right tabular-nums font-semibold">{fmtCNY(l._tradeCommission)}</td>
                              </tr>); })}
                            <tr className="border-t-2 border-amber-200 bg-amber-50 font-semibold text-xs"><td className="px-3 py-2" colSpan={6}>合计</td><td className="px-3 py-2 text-right tabular-nums">{fmtCNY(sumT)}</td></tr>
                          </>); })()}
                        </tbody></table>
                      </div>)}
                  </div>)}
              </div>);
          })}
        </div>)}

      {/* Warnings */}
      {activeReport === "warnings" && (
        <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden">
          {warnings.length === 0 ? <div className="p-10 text-center text-slate-500"><CheckCircle2 className="mx-auto text-emerald-500 mb-2" size={32} />所有订单正常</div> : (
            <div className="divide-y divide-slate-100">{warnings.map((w, i) => (<div key={i} className="p-4 flex gap-3"><AlertTriangle size={18} className="text-amber-500 shrink-0 mt-0.5" /><div className="flex-1 min-w-0"><div className="text-xs font-semibold text-amber-700 mb-0.5">{w.type}</div><div className="text-sm text-slate-700">{w.msg}</div></div></div>))}</div>)}
        </div>)}

      <div className="flex justify-between mt-6">
        <button onClick={onBack} className="px-5 py-3 bg-white border border-slate-300 text-slate-700 rounded-xl font-medium hover:bg-slate-50 flex items-center gap-2"><ChevronLeft size={18} /> 返回参数</button>
        <button onClick={onRestart} className="px-5 py-3 bg-slate-900 text-white rounded-xl font-medium hover:bg-slate-800 flex items-center gap-2">重新开始</button>
      </div>
    </div>
  );
}

// ============================================================
// 主组件
// ============================================================
export default function App() {
  const [step, setStep] = useState(1);
  const [data, setData] = useState(null);
  const [params, setParams] = useState(null);
  const [result, setResult] = useState(null);
  const [calculating, setCalculating] = useState(false);

  const handleStep1Done = (d) => { setData(d); setStep(2); };
  const handleStep2Done = (p) => {
    setParams(p); setCalculating(true);
    setTimeout(() => {
      try { setResult(runCalculation({ settlements: data.settlementData.settlements, priceMap: data.settlementData.priceMap, persons: data.perfData.persons, ...p })); setStep(4); }
      catch (e) { alert("核算失败：" + (e.message || e)); }
      finally { setCalculating(false); }
    }, 50);
  };
  const handleRestart = () => { setStep(1); setData(null); setParams(null); setResult(null); };

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 to-slate-100">
      <header className="bg-white border-b border-slate-200">
        <div className="max-w-7xl mx-auto px-4 py-4 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-lg bg-gradient-to-br from-slate-800 to-slate-900 text-white flex items-center justify-center font-bold">KNE</div>
            <div><div className="font-bold text-slate-900">康乃尔薪酬核算系统</div><div className="text-xs text-slate-500">v7.0 · 依据 2026年康乃尔经营薪酬考核方案 V2.0 (0515)</div></div>
          </div>
          {data && <div className="hidden md:flex items-center gap-3 text-sm text-slate-600">
            <span className="px-2.5 py-1 bg-slate-100 rounded-md">核算月份：<b className="text-slate-900">{data.settlementData.month}</b></span>
            <span className="px-2.5 py-1 bg-slate-100 rounded-md">业务员：<b className="text-slate-900">{data.perfData.persons.length}</b></span>
          </div>}
        </div>
      </header>
      <Stepper step={step === 4 ? 4 : step} />
      {calculating && <div className="text-center py-20"><div className="inline-block w-10 h-10 border-4 border-slate-200 border-t-slate-900 rounded-full animate-spin" /><div className="mt-4 text-slate-600">核算中...</div></div>}
      {!calculating && step === 1 && <Step1Upload onComplete={handleStep1Done} />}
      {!calculating && step === 2 && <Step2Parameters data={data} onBack={() => setStep(1)} onComplete={handleStep2Done} />}
      {!calculating && step === 4 && result && <ResultPanel result={result} params={params} data={data} onBack={() => setStep(2)} onRestart={handleRestart} />}
      <footer className="max-w-7xl mx-auto px-4 py-8 text-center text-xs text-slate-400">居间单价数据保存在本地浏览器，下次打开仍可使用。</footer>
    </div>
  );
}
