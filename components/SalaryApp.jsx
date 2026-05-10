import React, { useState, useEffect, useMemo, useCallback, useRef } from "react";
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
// 常量与系数表（依据 PRD v6.0 第三章）
// ============================================================
const PRODUCT_LINE_RATES = {
  普通个体防护: 0.004,
  高端个体防护: 0.006,
  工业吸附: 0.006,
  "液滤-电驻极": 0.006,
  液滤: 0.006, // 数据表实际写法
  空滤: 0.006,
  生活擦拭: 0.006,
  耐高温材料: 0.008,
  透气弹性材料: 0.01,
};
const NEGATIVE_MARGIN_RATE = 0.0025;
const FUPAI_DEFAULT_RATE = 0.004; // 副牌默认按最低档

const TRADE_RATES = {
  熔喷料: 0.4,
  母粒: 0.3,
  无纺布: 0.2,
};

const PERFORMANCE_BASE_BY_LEVEL = {
  一级: 2000,
  二级: 3000,
  三级: 4000,
};

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

// 从物料描述提取重量 (KG/kg)
function extractWeightFromDesc(desc) {
  if (!desc || typeof desc !== "string") return null;
  // 匹配 "数字KG" 或 "数字kg"，允许小数点
  const matches = desc.match(/(\d+(?:\.\d+)?)\s*[Kk][Gg]/g);
  if (!matches || matches.length === 0) return null;
  // 取最后一个匹配项作为单位重量（一般在描述末端）
  const last = matches[matches.length - 1];
  const num = parseFloat(last);
  return isNaN(num) ? null : num;
}

// 单位换算为 kg
function convertToKg(qty, unit, desc) {
  if (qty === null || qty === undefined || isNaN(qty)) return { kg: null, source: "数量为空" };
  const q = Number(qty);
  if (!unit) return { kg: q, source: "单位缺失，按原值" };
  const u = String(unit).trim();
  if (u === "kg" || u === "KG" || u === "公斤" || u === "千克") {
    return { kg: q, source: "K列=kg" };
  }
  // 其他单位：箱/包/个/卷/张/片 等 → 从描述提取单位重量
  const w = extractWeightFromDesc(desc);
  if (w === null) return { kg: null, source: "无法从描述提取重量" };
  return { kg: q * w, source: `每${u}=${w}KG` };
}

// 销售提成系数（按 T 列类别）
function getProductLineRate(category) {
  if (!category) return { rate: FUPAI_DEFAULT_RATE, warning: "T列为空，按0.4%处理" };
  const c = String(category).trim();
  if (c === "贸易") return { rate: 0, warning: "贸易订单不参与销售提成" };
  if (c === "副牌") return { rate: FUPAI_DEFAULT_RATE, warning: "副牌按最低档0.4%" };
  if (PRODUCT_LINE_RATES[c] !== undefined) return { rate: PRODUCT_LINE_RATES[c], warning: null };
  return { rate: FUPAI_DEFAULT_RATE, warning: `未知产品线类别"${c}"，暂按0.4%` };
}

// 溢价提成系数（按溢价率）
function getPremiumRate(premiumRatio) {
  if (premiumRatio === null || premiumRatio === undefined || isNaN(premiumRatio)) return 0;
  if (premiumRatio <= 0) return 0;
  if (premiumRatio <= 0.05) return 0.15;
  if (premiumRatio <= 0.1) return 0.2;
  return 0.3;
}

// 解析 V 列基价（数值或 VLOOKUP 公式或空）
function resolveBasePrice(vCellValue, materialCode, priceMap) {
  // V 列若为数字（XLSX 通常会把缓存值返回数字）
  if (typeof vCellValue === "number" && !isNaN(vCellValue)) {
    return { price: vCellValue, source: "V列直接值" };
  }
  // 字符串：可能是数字字符串或公式
  if (typeof vCellValue === "string") {
    const trimmed = vCellValue.trim();
    if (trimmed && !trimmed.startsWith("=")) {
      const num = parseFloat(trimmed);
      if (!isNaN(num)) return { price: num, source: "V列直接值" };
    }
  }
  // 否则查基价表
  if (materialCode && priceMap.has(String(materialCode))) {
    return { price: priceMap.get(String(materialCode)), source: "基价表匹配" };
  }
  return { price: null, source: "无法获取" };
}

// ============================================================
// Excel 解析
// ============================================================
async function parseSettlementWorkbook(file) {
  const arrayBuffer = await file.arrayBuffer();
  const wb = XLSX.read(arrayBuffer, { type: "array", cellDates: true, cellFormula: true });

  // Sheet1: 回款明细
  const sheet1Name = wb.SheetNames.find((n) => n.includes("回款明细")) || wb.SheetNames[0];
  const ws1 = wb.Sheets[sheet1Name];
  const rows1 = XLSX.utils.sheet_to_json(ws1, { header: 1, defval: "", raw: true });

  if (rows1.length < 2) throw new Error("回款明细表为空或格式错误");
  const dataRows1 = rows1.slice(1).filter((r) => r.some((c) => c !== "" && c !== null && c !== undefined));

  const settlements = dataRows1.map((row, idx) => ({
    _rowIdx: idx + 2,
    年: row[0],
    月: row[1],
    销售订单: String(row[2] || "").trim(),
    订单类型描述: row[3],
    过账日期: row[4],
    售达方: String(row[5] || "").trim(),
    售达方描述: row[6],
    销售合同号: row[7],
    物料编码: String(row[8] || "").trim(),
    物料描述: row[9],
    单位: row[10],
    数量: row[11],
    含税单价: row[12],
    含税金额: row[13],
    收款金额: typeof row[14] === "number" ? row[14] : parseFloat(row[14]) || 0,
    交易日期: row[15],
    差额: row[16],
    备注: String(row[17] || ""),
    回款月份: row[18],
    类别: String(row[19] || "").trim(),
    业务经理: String(row[20] || "").trim(),
    基价_V列: row[21],
    运费单价: typeof row[22] === "number" ? row[22] : parseFloat(row[22]) || 0,
    提成活跃系数: row[23] === "" || row[23] === null || row[23] === undefined ? null : Number(row[23]),
  }));

  // Sheet2: 基价表
  const sheet2Name = wb.SheetNames.find((n) => n.includes("基价")) || wb.SheetNames[1];
  const priceMap = new Map();
  if (sheet2Name) {
    const ws2 = wb.Sheets[sheet2Name];
    const rows2 = XLSX.utils.sheet_to_json(ws2, { header: 1, defval: "", raw: true });
    // 数据从第3行开始（行1标题，行2列头）
    for (let i = 2; i < rows2.length; i++) {
      const code = rows2[i][0];
      const price = rows2[i][rows2[i].length - 1]; // 最后一列是基价（按实际4列结构 D 列）
      if (code !== "" && code !== "/" && price !== "" && !isNaN(parseFloat(price))) {
        priceMap.set(String(code).trim(), parseFloat(price));
      }
    }
  }

  // 提取月份
  const monthGuess = settlements.find((s) => s.回款月份)?.回款月份 || settlements[0]?.月 || "";

  return { settlements, priceMap, month: String(monthGuess), sheet1Name, sheet2Name };
}

async function parsePerformanceWorkbook(file) {
  const arrayBuffer = await file.arrayBuffer();
  const wb = XLSX.read(arrayBuffer, { type: "array", cellDates: true });
  const sheetName = wb.SheetNames.find((n) => n.includes("汇总") || n.includes("绩效")) || wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: true });

  // 第1行标题，第2行列头，第3行起数据
  if (rows.length < 3) throw new Error("绩效考核表数据不足");

  // 自动识别列结构：可能 6 列（PRD）或 7 列（实际包含抵扣）
  const headerRow = rows[1] || [];
  const hasDeduction = headerRow.some((h) => String(h || "").includes("抵扣"));

  const persons = [];
  for (let i = 2; i < rows.length; i++) {
    const row = rows[i];
    if (!row[0] || String(row[0]).trim() === "") continue;
    const name = String(row[0]).trim();
    let deduction = 0,
      completionRate,
      perfScore,
      baseSalary,
      level,
      perfBase;
    if (hasDeduction) {
      // 7 列：姓名 / 抵扣总额 / 完成率 / 绩效分 / 基本工资 / 等级 / 基数
      deduction = Number(row[1]) || 0;
      completionRate = Number(row[2]) || 0;
      perfScore = Number(row[3]) || 0;
      baseSalary = Number(row[4]) || 2000;
      level = String(row[5] || "一级").trim();
      perfBase = Number(row[6]) || PERFORMANCE_BASE_BY_LEVEL[level] || 2000;
    } else {
      // 6 列：姓名 / 完成率 / 绩效分 / 基本工资 / 等级 / 基数
      completionRate = Number(row[1]) || 0;
      perfScore = Number(row[2]) || 0;
      baseSalary = Number(row[3]) || 2000;
      level = String(row[4] || "一级").trim();
      perfBase = Number(row[5]) || PERFORMANCE_BASE_BY_LEVEL[level] || 2000;
    }
    persons.push({
      name,
      defaultBaseSalary: baseSalary,
      defaultPerfBase: perfBase,
      defaultDeduction: deduction,
      completionRate,
      perfScore,
      level,
    });
  }
  return { persons, hasDeduction };
}

// ============================================================
// 计算引擎
// ============================================================
function runCalculation({
  settlements,
  priceMap,
  persons,
  personSettings, // {name: {baseSalary, perfBase, deduction}}
  priceDeductions, // [{customerId, pricePerKg}]
  negativeOrders, // [orderId]
  tradeMarginInputs, // {orderKey: {category, margin}}
  activityOverrides, // {customerId: number}
}) {
  const warnings = [];
  const allLines = []; // 全部行处理结果（含剔除）
  const validRegular = []; // 参与常规提成
  const tradeLines = []; // 贸易订单
  const dropped = []; // 剔除

  const negativeSet = new Set(negativeOrders.map((o) => String(o).trim()).filter(Boolean));
  const priceDedMap = new Map(priceDeductions.map((d) => [String(d.customerId).trim(), Number(d.pricePerKg) || 0]));
  const activityMap = new Map(Object.entries(activityOverrides));
  const personNames = new Set(persons.map((p) => p.name));

  for (const r of settlements) {
    const lineBase = {
      ...r,
      _kg: null,
      _kgSource: null,
      _basePrice: null,
      _basePriceSource: null,
      _freightUnit: r.运费单价 || 0,
      _freightTotal: null,
      _exFactoryPrice: null,
      _commBase: null,
      _commRate: null,
      _commRateSource: null,
      _activityCoef: null,
      _activityOverride: false,
      _saleCommission: 0,
      _priceDeduction: 0,
      _adjPrice: null,
      _premiumRatio: null,
      _premiumRate: 0,
      _premiumBase: 0,
      _premiumCommission: 0,
      _isNegative: false,
      _hasDeductionApplied: false,
      _flag: null, // 'valid' | 'trade' | 'dropped'
      _dropReason: null,
    };

    // 1. 无效过滤
    if (!r.物料编码) {
      lineBase._flag = "dropped";
      lineBase._dropReason = "物料编码为空";
      dropped.push(lineBase);
      allLines.push(lineBase);
      continue;
    }
    const remarkStr = String(r.备注 || "");
    if (remarkStr.includes("预收款") || remarkStr.includes("无合同")) {
      lineBase._flag = "dropped";
      lineBase._dropReason = `备注含"${remarkStr.includes("预收款") ? "预收款" : "无合同"}"`;
      dropped.push(lineBase);
      allLines.push(lineBase);
      continue;
    }

    // 2. 单位换算
    const { kg, source: kgSource } = convertToKg(r.数量, r.单位, r.物料描述);
    lineBase._kg = kg;
    lineBase._kgSource = kgSource;
    if (kg === null || kg <= 0) {
      lineBase._flag = "dropped";
      lineBase._dropReason = `单位折算失败：${kgSource}`;
      dropped.push(lineBase);
      allLines.push(lineBase);
      warnings.push({
        type: "单位折算失败",
        order: r.销售订单,
        material: r.物料编码,
        person: r.业务经理,
        msg: `订单 ${r.销售订单} (物料 ${r.物料编码}) ${kgSource}`,
      });
      continue;
    }

    // 3. 基价
    const { price, source: priceSource } = resolveBasePrice(r.基价_V列, r.物料编码, priceMap);
    lineBase._basePrice = price;
    lineBase._basePriceSource = priceSource;
    if (price === null || price <= 0) {
      lineBase._flag = "dropped";
      lineBase._dropReason = "基价缺失";
      dropped.push(lineBase);
      allLines.push(lineBase);
      warnings.push({
        type: "基价缺失",
        order: r.销售订单,
        material: r.物料编码,
        person: r.业务经理,
        msg: `订单 ${r.销售订单} 物料 ${r.物料编码} 在 V 列和基价表中均无法获取基价`,
      });
      continue;
    }

    // 4. 运费总额
    lineBase._freightTotal = lineBase._freightUnit * kg;

    // 5. 出厂售价
    lineBase._exFactoryPrice = (Number(r.收款金额) - lineBase._freightTotal) / kg;

    // 6. 贸易订单单独流程
    if (r.类别 === "贸易") {
      const orderKey = `${r.销售订单}_${r.物料编码}`;
      const tradeInput = tradeMarginInputs[orderKey];
      lineBase._flag = "trade";
      lineBase._tradeCategory = tradeInput?.category || null;
      lineBase._tradeMargin = tradeInput && !isNaN(parseFloat(tradeInput.margin)) ? parseFloat(tradeInput.margin) : 0;
      lineBase._tradeRate = tradeInput?.category ? TRADE_RATES[tradeInput.category] || 0 : 0;
      lineBase._tradeCommission = lineBase._tradeMargin * lineBase._tradeRate;
      if (!tradeInput || !tradeInput.category) {
        warnings.push({
          type: "贸易订单未录入毛利/品类",
          order: r.销售订单,
          material: r.物料编码,
          person: r.业务经理,
          msg: `贸易订单 ${r.销售订单} 物料 ${r.物料编码} 未录入毛利或品类，提成按 0`,
        });
      }
      tradeLines.push(lineBase);
      allLines.push(lineBase);
      continue;
    }

    // 7. 常规订单：销售提成
    const isNegative = negativeSet.has(String(r.销售订单).trim());
    lineBase._isNegative = isNegative;
    let commRate, commRateSource;
    if (isNegative) {
      commRate = NEGATIVE_MARGIN_RATE;
      commRateSource = "负毛利订单(0.25%)";
    } else {
      const pl = getProductLineRate(r.类别);
      commRate = pl.rate;
      commRateSource = `产品线[${r.类别}]`;
      if (pl.warning) {
        warnings.push({
          type: "类别异常",
          order: r.销售订单,
          person: r.业务经理,
          msg: `订单 ${r.销售订单}: ${pl.warning}`,
        });
      }
    }
    lineBase._commRate = commRate;
    lineBase._commRateSource = commRateSource;

    // 提成基数 = min(出厂售价, 基价) × 核定公斤
    const effectivePrice = Math.min(lineBase._exFactoryPrice, price);
    lineBase._commBase = effectivePrice * kg;

    // 活跃度系数
    let activityCoef;
    const customerActOverride = activityMap.get(r.售达方);
    if (customerActOverride !== undefined && customerActOverride !== null && customerActOverride !== "") {
      activityCoef = Number(customerActOverride);
      lineBase._activityOverride = true;
    } else if (r.提成活跃系数 !== null && r.提成活跃系数 !== undefined && !isNaN(r.提成活跃系数)) {
      activityCoef = Number(r.提成活跃系数);
    } else {
      activityCoef = 1;
    }
    lineBase._activityCoef = activityCoef;

    lineBase._saleCommission = lineBase._commBase * commRate * activityCoef;

    // 溢价提成
    const priceDed = priceDedMap.get(r.售达方) || 0;
    lineBase._priceDeduction = priceDed;
    lineBase._adjPrice = lineBase._exFactoryPrice - priceDed;
    lineBase._hasDeductionApplied = priceDed > 0;
    if (lineBase._adjPrice > price) {
      lineBase._premiumBase = (lineBase._adjPrice - price) * kg;
      lineBase._premiumRatio = lineBase._adjPrice / price - 1;
      lineBase._premiumRate = getPremiumRate(lineBase._premiumRatio);
      lineBase._premiumCommission = lineBase._premiumBase * lineBase._premiumRate;
    } else {
      lineBase._premiumBase = 0;
      lineBase._premiumRatio = lineBase._adjPrice / price - 1;
      lineBase._premiumRate = 0;
      lineBase._premiumCommission = 0;
    }

    lineBase._flag = "valid";

    // 人员归属预警
    if (r.业务经理 && !personNames.has(r.业务经理)) {
      warnings.push({
        type: "人员未在绩效表",
        order: r.销售订单,
        person: r.业务经理,
        msg: `业务员 "${r.业务经理}" 不在绩效考核表中，订单参与交叉验证但不计提成`,
      });
    }

    validRegular.push(lineBase);
    allLines.push(lineBase);
  }

  // 居间单价/负毛利订单匹配检查
  for (const d of priceDeductions) {
    const cid = String(d.customerId).trim();
    if (!cid) continue;
    const match = settlements.some((s) => s.售达方 === cid);
    if (!match) {
      warnings.push({
        type: "居间单价无匹配",
        msg: `售达方编号 ${cid} 在回款明细中无匹配订单`,
      });
    }
  }
  for (const o of negativeOrders) {
    const oid = String(o).trim();
    if (!oid) continue;
    const match = settlements.some((s) => String(s.销售订单).trim() === oid);
    if (!match) {
      warnings.push({
        type: "负毛利订单无匹配",
        msg: `订单编号 ${oid} 在回款明细中无匹配`,
      });
    }
  }

  // 按业务员聚合
  const personResults = persons.map((p) => {
    const settings = personSettings[p.name] || {};
    const baseSalary = settings.baseSalary !== undefined ? Number(settings.baseSalary) : p.defaultBaseSalary;
    const perfBase = settings.perfBase !== undefined ? Number(settings.perfBase) : p.defaultPerfBase;
    const deduction = settings.deduction !== undefined ? Number(settings.deduction) : p.defaultDeduction;

    // 绩效工资
    const cappedRate = Math.min(p.completionRate, 1);
    const perfWage = (0.8 * cappedRate + 0.2 * p.perfScore) * perfBase;

    // 该人的常规订单
    const myRegular = validRegular.filter((l) => l.业务经理 === p.name);
    const myTrade = tradeLines.filter((l) => l.业务经理 === p.name);
    const myDropped = dropped.filter((l) => l.业务经理 === p.name);
    const myAll = allLines.filter((l) => l.业务经理 === p.name);

    const totalSaleCommission = myRegular.reduce((s, l) => s + l._saleCommission, 0);
    const totalCommBase = myRegular.reduce((s, l) => s + l._commBase, 0);
    const weightedAvgRate = totalCommBase > 0 ? totalSaleCommission / totalCommBase : 0;
    const deductionAmt = deduction * weightedAvgRate;
    const finalSaleCommission = Math.max(0, totalSaleCommission - deductionAmt);

    const totalPremium = myRegular.reduce((s, l) => s + l._premiumCommission, 0);
    const totalTrade = myTrade.reduce((s, l) => s + (l._tradeCommission || 0), 0);

    const totalSalary = baseSalary + perfWage + finalSaleCommission + totalPremium + totalTrade;

    // 交叉验证
    const allReceipt = myAll.reduce((s, l) => s + (Number(l.收款金额) || 0), 0);
    const validReceipt = myRegular.reduce((s, l) => s + (Number(l.收款金额) || 0), 0);
    const tradeReceipt = myTrade.reduce((s, l) => s + (Number(l.收款金额) || 0), 0);
    const droppedReceipt = myDropped.reduce((s, l) => s + (Number(l.收款金额) || 0), 0);

    return {
      ...p,
      baseSalary,
      perfBase,
      deduction,
      perfWage,
      cappedRate,
      myRegular,
      myTrade,
      myDropped,
      totalSaleCommission,
      totalCommBase,
      weightedAvgRate,
      deductionAmt,
      finalSaleCommission,
      totalPremium,
      totalTrade,
      totalSalary,
      allReceipt,
      validReceipt,
      tradeReceipt,
      droppedReceipt,
    };
  });

  return {
    warnings,
    personResults,
    allLines,
    validRegular,
    tradeLines,
    dropped,
  };
}

// ============================================================
// 子组件：步骤指示器
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
        const active = step === s.id;
        const done = step > s.id;
        return (
          <React.Fragment key={s.id}>
            <div className="flex items-center gap-2.5">
              <div
                className={`w-10 h-10 rounded-full flex items-center justify-center transition-all border-2 ${
                  done
                    ? "bg-emerald-600 border-emerald-600 text-white"
                    : active
                    ? "bg-slate-800 border-slate-800 text-white shadow-lg shadow-slate-800/30"
                    : "bg-white border-slate-300 text-slate-400"
                }`}
              >
                {done ? <CheckCircle2 size={18} /> : <Icon size={18} />}
              </div>
              <div className={`hidden md:block text-sm font-medium ${active ? "text-slate-800" : done ? "text-emerald-600" : "text-slate-400"}`}>
                {s.label}
              </div>
            </div>
            {i < steps.length - 1 && (
              <div className={`h-0.5 w-6 md:w-12 ${step > s.id ? "bg-emerald-600" : "bg-slate-200"}`}></div>
            )}
          </React.Fragment>
        );
      })}
    </div>
  );
}

// ============================================================
// 步骤 1：文件上传
// ============================================================
function Step1Upload({ onComplete }) {
  const [settlementFile, setSettlementFile] = useState(null);
  const [perfFile, setPerfFile] = useState(null);
  const [parsing, setParsing] = useState(false);
  const [error, setError] = useState(null);

  const handleParse = async () => {
    if (!settlementFile || !perfFile) {
      setError("请同时上传两份 Excel 文件");
      return;
    }
    setParsing(true);
    setError(null);
    try {
      const settlementData = await parseSettlementWorkbook(settlementFile);
      const perfData = await parsePerformanceWorkbook(perfFile);
      onComplete({ settlementData, perfData });
    } catch (e) {
      setError("解析失败：" + (e.message || String(e)));
    } finally {
      setParsing(false);
    }
  };

  const FileBox = ({ label, file, onChange, hint }) => (
    <label className="block cursor-pointer">
      <div
        className={`relative border-2 border-dashed rounded-2xl p-7 transition-all ${
          file ? "border-emerald-500 bg-emerald-50/50" : "border-slate-300 hover:border-slate-500 hover:bg-slate-50"
        }`}
      >
        <div className="flex items-start gap-4">
          <div
            className={`w-14 h-14 rounded-xl flex items-center justify-center shrink-0 ${
              file ? "bg-emerald-600 text-white" : "bg-slate-100 text-slate-500"
            }`}
          >
            <FileSpreadsheet size={26} />
          </div>
          <div className="flex-1 min-w-0">
            <div className="font-semibold text-slate-900 mb-1">{label}</div>
            <div className="text-xs text-slate-500 mb-2">{hint}</div>
            {file ? (
              <div className="text-sm text-emerald-700 font-medium truncate">✓ {file.name}</div>
            ) : (
              <div className="text-sm text-slate-400">点击或拖拽以选择 .xlsx 文件</div>
            )}
          </div>
        </div>
        <input
          type="file"
          accept=".xlsx,.xls"
          className="hidden"
          onChange={(e) => onChange(e.target.files?.[0] || null)}
        />
      </div>
    </label>
  );

  return (
    <div className="max-w-3xl mx-auto px-4">
      <div className="text-center mb-8">
        <h2 className="text-2xl font-bold text-slate-900 mb-2">步骤一 · 上传数据文件</h2>
        <p className="text-slate-500">请上传当月的回款明细表与绩效考核表</p>
      </div>

      <div className="space-y-4">
        <FileBox
          label="回款明细表"
          hint="包含「回款明细」和「基价表」两个 Sheet"
          file={settlementFile}
          onChange={setSettlementFile}
        />
        <FileBox
          label="绩效考核表"
          hint="包含人员姓名、业绩完成率、绩效分、抵扣金额等"
          file={perfFile}
          onChange={setPerfFile}
        />
      </div>

      {error && (
        <div className="mt-5 p-4 bg-red-50 border border-red-200 text-red-700 rounded-xl text-sm flex gap-2">
          <AlertTriangle size={18} className="shrink-0 mt-0.5" />
          {error}
        </div>
      )}

      <button
        onClick={handleParse}
        disabled={parsing || !settlementFile || !perfFile}
        className="w-full mt-6 py-4 bg-slate-900 text-white rounded-xl font-semibold hover:bg-slate-800 disabled:bg-slate-300 disabled:cursor-not-allowed transition-all flex items-center justify-center gap-2"
      >
        {parsing ? "解析中..." : <>下一步 · 确认参数 <ChevronRight size={18} /></>}
      </button>
    </div>
  );
}

// ============================================================
// 步骤 2：参数确认（多 Tab）
// ============================================================
function Step2Parameters({ data, onBack, onComplete }) {
  const { settlementData, perfData } = data;
  const { settlements } = settlementData;
  const { persons } = perfData;

  const [activeTab, setActiveTab] = useState("salary");

  // 2a + 2b: 人员设置（基本工资、绩效工资基数、抵扣）
  const [personSettings, setPersonSettings] = useState(() => {
    const init = {};
    persons.forEach((p) => {
      init[p.name] = {
        baseSalary: p.defaultBaseSalary,
        perfBase: p.defaultPerfBase,
        deduction: p.defaultDeduction,
      };
    });
    return init;
  });

  // 2c: 居间单价（持久化到 localStorage，跨会话保留）
  const [priceDeductions, setPriceDeductions] = useState([]);
  const [storageReady, setStorageReady] = useState(false);

  useEffect(() => {
    try {
      if (typeof window !== "undefined" && window.localStorage) {
        const raw = window.localStorage.getItem(PRICE_DEDUCTIONS_KEY);
        if (raw) {
          const parsed = JSON.parse(raw);
          if (Array.isArray(parsed)) setPriceDeductions(parsed);
        }
      }
    } catch (e) {
      // 读取失败 - 使用空数组即可
    } finally {
      setStorageReady(true);
    }
  }, []);

  // 持久化 priceDeductions
  useEffect(() => {
    if (!storageReady) return;
    try {
      if (typeof window !== "undefined" && window.localStorage) {
        window.localStorage.setItem(PRICE_DEDUCTIONS_KEY, JSON.stringify(priceDeductions));
      }
    } catch (e) {
      // ignore (隐私模式或存储已满)
    }
  }, [priceDeductions, storageReady]);

  // 2d: 负毛利订单
  const [negativeOrders, setNegativeOrders] = useState([]);

  // 2e: 贸易订单毛利录入
  const tradeOrdersRaw = useMemo(() => {
    return settlements.filter((s) => s.类别 === "贸易" && s.物料编码);
  }, [settlements]);
  const [tradeMarginInputs, setTradeMarginInputs] = useState({});

  // 2f: 活跃度系数（按售达方去重）
  const customerActivityList = useMemo(() => {
    const map = new Map();
    for (const s of settlements) {
      if (!s.售达方 || !s.物料编码) continue;
      if (!map.has(s.售达方)) {
        map.set(s.售达方, {
          customerId: s.售达方,
          name: s.售达方描述,
          defaultCoef: s.提成活跃系数 === null || s.提成活跃系数 === undefined ? 1 : Number(s.提成活跃系数),
          orderCount: 0,
        });
      }
      map.get(s.售达方).orderCount++;
    }
    return Array.from(map.values());
  }, [settlements]);

  const [activityOverrides, setActivityOverrides] = useState({});

  const updatePersonSetting = (name, key, value) => {
    setPersonSettings((prev) => ({
      ...prev,
      [name]: { ...prev[name], [key]: value === "" ? "" : value },
    }));
  };
  const restorePersonDefault = (name) => {
    const p = persons.find((x) => x.name === name);
    if (p) {
      setPersonSettings((prev) => ({
        ...prev,
        [name]: { baseSalary: p.defaultBaseSalary, perfBase: p.defaultPerfBase, deduction: p.defaultDeduction },
      }));
    }
  };

  const addPriceDeduction = () => setPriceDeductions((p) => [...p, { customerId: "", pricePerKg: "" }]);
  const updatePriceDeduction = (i, key, val) => {
    setPriceDeductions((p) => {
      const arr = [...p];
      arr[i] = { ...arr[i], [key]: val };
      return arr;
    });
  };
  const removePriceDeduction = (i) => setPriceDeductions((p) => p.filter((_, idx) => idx !== i));

  const addNegativeOrder = () => setNegativeOrders((n) => [...n, ""]);
  const updateNegativeOrder = (i, val) => {
    setNegativeOrders((n) => {
      const arr = [...n];
      arr[i] = val;
      return arr;
    });
  };
  const removeNegativeOrder = (i) => setNegativeOrders((n) => n.filter((_, idx) => idx !== i));

  const tabs = [
    { id: "salary", label: "工资基数", icon: Coins, count: persons.length },
    { id: "deduction", label: "抵扣项目", icon: Receipt, count: persons.length },
    { id: "interim", label: "居间单价", icon: Briefcase, count: priceDeductions.length },
    { id: "negative", label: "负毛利订单", icon: FileWarning, count: negativeOrders.length },
    { id: "trade", label: "贸易毛利", icon: TrendingUp, count: tradeOrdersRaw.length },
    { id: "activity", label: "活跃度系数", icon: Activity, count: customerActivityList.length },
  ];

  const handleSubmit = () => {
    onComplete({
      personSettings,
      priceDeductions: priceDeductions.filter((d) => d.customerId && d.customerId.trim()),
      negativeOrders: negativeOrders.filter((o) => o && String(o).trim()),
      tradeMarginInputs,
      activityOverrides,
    });
  };

  return (
    <div className="max-w-7xl mx-auto px-4 pb-12">
      <div className="text-center mb-6">
        <h2 className="text-2xl font-bold text-slate-900 mb-1">步骤二 · 确认 / 修改参数</h2>
        <p className="text-slate-500 text-sm">默认值已从文件读取，可逐项调整。居间单价跨会话自动保留。</p>
      </div>

      {/* Tab 导航 */}
      <div className="flex flex-wrap gap-2 border-b border-slate-200 mb-6">
        {tabs.map((t) => {
          const Icon = t.icon;
          const active = activeTab === t.id;
          return (
            <button
              key={t.id}
              onClick={() => setActiveTab(t.id)}
              className={`px-4 py-2.5 -mb-px border-b-2 font-medium text-sm flex items-center gap-2 transition-colors ${
                active
                  ? "border-slate-900 text-slate-900"
                  : "border-transparent text-slate-500 hover:text-slate-700"
              }`}
            >
              <Icon size={16} />
              {t.label}
              <span
                className={`px-1.5 py-0.5 rounded-md text-xs ${
                  active ? "bg-slate-900 text-white" : "bg-slate-100 text-slate-500"
                }`}
              >
                {t.count}
              </span>
            </button>
          );
        })}
      </div>

      {/* Tab 内容 */}
      <div className="bg-white rounded-2xl border border-slate-200 p-5 md:p-6 shadow-sm">
        {activeTab === "salary" && (
          <div>
            <h3 className="font-semibold text-slate-900 mb-1">基本工资 与 绩效工资基数</h3>
            <p className="text-sm text-slate-500 mb-4">
              默认值取自绩效表。两者独立可调；绩效工资基数按等级：一级=2000、二级=3000、三级=4000。
            </p>
            <div className="overflow-x-auto -mx-2">
              <table className="w-full text-sm">
                <thead>
                  <tr className="text-left bg-slate-50 border-y border-slate-200">
                    <th className="px-3 py-2.5 font-medium text-slate-700">人员</th>
                    <th className="px-3 py-2.5 font-medium text-slate-700">绩效等级</th>
                    <th className="px-3 py-2.5 font-medium text-slate-700">基本工资-默认</th>
                    <th className="px-3 py-2.5 font-medium text-slate-700">基本工资-当前</th>
                    <th className="px-3 py-2.5 font-medium text-slate-700">绩效基数-默认</th>
                    <th className="px-3 py-2.5 font-medium text-slate-700">绩效基数-当前</th>
                    <th className="px-3 py-2.5 font-medium text-slate-700">完成率</th>
                    <th className="px-3 py-2.5 font-medium text-slate-700">绩效分</th>
                    <th className="px-3 py-2.5 font-medium text-slate-700">操作</th>
                  </tr>
                </thead>
                <tbody>
                  {persons.map((p) => (
                    <tr key={p.name} className="border-b border-slate-100 hover:bg-slate-50/50">
                      <td className="px-3 py-2.5 font-semibold text-slate-900">{p.name}</td>
                      <td className="px-3 py-2.5 text-slate-600">{p.level}</td>
                      <td className="px-3 py-2.5 text-slate-400">{fmtCNY(p.defaultBaseSalary)}</td>
                      <td className="px-3 py-2.5">
                        <input
                          type="number"
                          value={personSettings[p.name].baseSalary}
                          onChange={(e) => updatePersonSetting(p.name, "baseSalary", e.target.value)}
                          className="w-28 px-2 py-1 border border-slate-300 rounded-md text-sm focus:outline-none focus:ring-2 focus:ring-slate-400"
                        />
                      </td>
                      <td className="px-3 py-2.5 text-slate-400">{fmtCNY(p.defaultPerfBase)}</td>
                      <td className="px-3 py-2.5">
                        <input
                          type="number"
                          value={personSettings[p.name].perfBase}
                          onChange={(e) => updatePersonSetting(p.name, "perfBase", e.target.value)}
                          className="w-28 px-2 py-1 border border-slate-300 rounded-md text-sm focus:outline-none focus:ring-2 focus:ring-slate-400"
                        />
                      </td>
                      <td className="px-3 py-2.5 text-slate-600">{fmtPct(p.completionRate)}</td>
                      <td className="px-3 py-2.5 text-slate-600">{fmtNum(p.perfScore)}</td>
                      <td className="px-3 py-2.5">
                        <button
                          onClick={() => restorePersonDefault(p.name)}
                          className="text-slate-500 hover:text-slate-900 inline-flex items-center gap-1 text-xs"
                        >
                          <RotateCcw size={12} /> 恢复
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {activeTab === "deduction" && (
          <div>
            <h3 className="font-semibold text-slate-900 mb-1">当月抵扣项目总额</h3>
            <p className="text-sm text-slate-500 mb-4">
              含运输费、差旅费、招待费、样品费的合计金额。{perfData.hasDeduction ? "默认值已从绩效表读取。" : "默认 0。"}
            </p>
            <div className="overflow-x-auto -mx-2">
              <table className="w-full text-sm">
                <thead>
                  <tr className="text-left bg-slate-50 border-y border-slate-200">
                    <th className="px-3 py-2.5 font-medium text-slate-700">人员</th>
                    <th className="px-3 py-2.5 font-medium text-slate-700">默认值（元）</th>
                    <th className="px-3 py-2.5 font-medium text-slate-700">当前值（元）</th>
                    <th className="px-3 py-2.5 font-medium text-slate-700">操作</th>
                  </tr>
                </thead>
                <tbody>
                  {persons.map((p) => (
                    <tr key={p.name} className="border-b border-slate-100 hover:bg-slate-50/50">
                      <td className="px-3 py-2.5 font-semibold text-slate-900">{p.name}</td>
                      <td className="px-3 py-2.5 text-slate-400">{fmtCNY(p.defaultDeduction)}</td>
                      <td className="px-3 py-2.5">
                        <input
                          type="number"
                          value={personSettings[p.name].deduction}
                          onChange={(e) => updatePersonSetting(p.name, "deduction", e.target.value)}
                          className="w-36 px-2 py-1 border border-slate-300 rounded-md text-sm focus:outline-none focus:ring-2 focus:ring-slate-400"
                        />
                      </td>
                      <td className="px-3 py-2.5">
                        <button
                          onClick={() => updatePersonSetting(p.name, "deduction", p.defaultDeduction)}
                          className="text-slate-500 hover:text-slate-900 inline-flex items-center gap-1 text-xs"
                        >
                          <RotateCcw size={12} /> 恢复
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {activeTab === "interim" && (
          <div>
            <div className="flex items-start justify-between mb-1">
              <h3 className="font-semibold text-slate-900">溢价提成扣除项目（居间单价）</h3>
              <span className="text-xs text-emerald-600 bg-emerald-50 px-2 py-1 rounded-md">
                ✓ 自动持久化保存
              </span>
            </div>
            <p className="text-sm text-slate-500 mb-4">
              针对特定客户从出厂售价中扣除的费用项目单价（元/kg）。仅影响溢价提成。下次打开 App 仍保留。
            </p>
            <div className="space-y-2 mb-3">
              {priceDeductions.length === 0 && (
                <div className="text-center py-8 text-slate-400 text-sm bg-slate-50/50 rounded-lg border border-dashed border-slate-200">
                  尚未录入任何居间单价记录
                </div>
              )}
              {priceDeductions.map((d, i) => (
                <div key={i} className="flex items-center gap-2">
                  <input
                    placeholder="售达方编号 (如 C523103)"
                    value={d.customerId}
                    onChange={(e) => updatePriceDeduction(i, "customerId", e.target.value)}
                    className="flex-1 px-3 py-2 border border-slate-300 rounded-md text-sm focus:outline-none focus:ring-2 focus:ring-slate-400"
                  />
                  <input
                    type="number"
                    step="0.01"
                    placeholder="居间单价 (元/kg)"
                    value={d.pricePerKg}
                    onChange={(e) => updatePriceDeduction(i, "pricePerKg", e.target.value)}
                    className="w-44 px-3 py-2 border border-slate-300 rounded-md text-sm focus:outline-none focus:ring-2 focus:ring-slate-400"
                  />
                  <button
                    onClick={() => removePriceDeduction(i)}
                    className="p-2 text-slate-400 hover:text-red-600 hover:bg-red-50 rounded-md"
                  >
                    <Trash2 size={16} />
                  </button>
                </div>
              ))}
            </div>
            <button
              onClick={addPriceDeduction}
              className="text-sm px-3 py-2 border border-dashed border-slate-400 rounded-md hover:bg-slate-50 inline-flex items-center gap-1.5 text-slate-700"
            >
              <Plus size={14} /> 添加一条
            </button>
          </div>
        )}

        {activeTab === "negative" && (
          <div>
            <h3 className="font-semibold text-slate-900 mb-1">负毛利订单</h3>
            <p className="text-sm text-slate-500 mb-4">
              录入订单编号。被标记的订单销售提成系数统一按 <b>0.25%</b> 计算（覆盖产品线系数），不影响溢价提成。
            </p>
            <div className="space-y-2 mb-3">
              {negativeOrders.length === 0 && (
                <div className="text-center py-8 text-slate-400 text-sm bg-slate-50/50 rounded-lg border border-dashed border-slate-200">
                  尚未标记任何负毛利订单
                </div>
              )}
              {negativeOrders.map((o, i) => (
                <div key={i} className="flex items-center gap-2">
                  <input
                    placeholder="订单编号 (如 1004587260)"
                    value={o}
                    onChange={(e) => updateNegativeOrder(i, e.target.value)}
                    className="flex-1 px-3 py-2 border border-slate-300 rounded-md text-sm focus:outline-none focus:ring-2 focus:ring-slate-400"
                  />
                  <button
                    onClick={() => removeNegativeOrder(i)}
                    className="p-2 text-slate-400 hover:text-red-600 hover:bg-red-50 rounded-md"
                  >
                    <Trash2 size={16} />
                  </button>
                </div>
              ))}
            </div>
            <button
              onClick={addNegativeOrder}
              className="text-sm px-3 py-2 border border-dashed border-slate-400 rounded-md hover:bg-slate-50 inline-flex items-center gap-1.5 text-slate-700"
            >
              <Plus size={14} /> 添加一条
            </button>
          </div>
        )}

        {activeTab === "trade" && (
          <div>
            <h3 className="font-semibold text-slate-900 mb-1">贸易业务毛利录入</h3>
            <p className="text-sm text-slate-500 mb-4">
              T 列为「贸易」的订单已自动识别，请逐笔录入毛利金额与品类。提成比例：熔喷料 40% / 母粒 30% / 无纺布 20%。
            </p>
            {tradeOrdersRaw.length === 0 ? (
              <div className="text-center py-10 text-slate-400 text-sm bg-slate-50/50 rounded-lg border border-dashed border-slate-200">
                本月无贸易业务订单
              </div>
            ) : (
              <div className="overflow-x-auto -mx-2">
                <table className="w-full text-sm">
                  <thead>
                    <tr className="text-left bg-slate-50 border-y border-slate-200">
                      <th className="px-3 py-2.5 font-medium text-slate-700">订单号</th>
                      <th className="px-3 py-2.5 font-medium text-slate-700">业务员</th>
                      <th className="px-3 py-2.5 font-medium text-slate-700">物料</th>
                      <th className="px-3 py-2.5 font-medium text-slate-700">收款金额</th>
                      <th className="px-3 py-2.5 font-medium text-slate-700">贸易品类</th>
                      <th className="px-3 py-2.5 font-medium text-slate-700">毛利金额（元）</th>
                    </tr>
                  </thead>
                  <tbody>
                    {tradeOrdersRaw.map((o) => {
                      const key = `${o.销售订单}_${o.物料编码}`;
                      const cur = tradeMarginInputs[key] || {};
                      return (
                        <tr key={key} className="border-b border-slate-100 hover:bg-slate-50/50">
                          <td className="px-3 py-2.5 font-mono text-xs text-slate-700">{o.销售订单}</td>
                          <td className="px-3 py-2.5">{o.业务经理}</td>
                          <td className="px-3 py-2.5 text-xs max-w-xs truncate">{o.物料描述}</td>
                          <td className="px-3 py-2.5">{fmtCNY(o.收款金额)}</td>
                          <td className="px-3 py-2.5">
                            <select
                              value={cur.category || ""}
                              onChange={(e) =>
                                setTradeMarginInputs((p) => ({ ...p, [key]: { ...p[key], category: e.target.value } }))
                              }
                              className="px-2 py-1 border border-slate-300 rounded-md text-sm bg-white"
                            >
                              <option value="">- 选择 -</option>
                              <option value="熔喷料">熔喷料 (40%)</option>
                              <option value="母粒">母粒 (30%)</option>
                              <option value="无纺布">无纺布 (20%)</option>
                            </select>
                          </td>
                          <td className="px-3 py-2.5">
                            <input
                              type="number"
                              step="0.01"
                              value={cur.margin || ""}
                              onChange={(e) =>
                                setTradeMarginInputs((p) => ({ ...p, [key]: { ...p[key], margin: e.target.value } }))
                              }
                              className="w-32 px-2 py-1 border border-slate-300 rounded-md text-sm"
                            />
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        )}

        {activeTab === "activity" && (
          <div>
            <h3 className="font-semibold text-slate-900 mb-1">提成活跃度系数（按售达方）</h3>
            <p className="text-sm text-slate-500 mb-4">
              自动从 X 列读取。同一售达方共用同一系数。仅适用于销售提成（溢价提成不适用）。
            </p>
            <div className="overflow-x-auto -mx-2 max-h-[60vh] overflow-y-auto">
              <table className="w-full text-sm">
                <thead className="sticky top-0 bg-white">
                  <tr className="text-left bg-slate-50 border-y border-slate-200">
                    <th className="px-3 py-2.5 font-medium text-slate-700">售达方</th>
                    <th className="px-3 py-2.5 font-medium text-slate-700">客户名称</th>
                    <th className="px-3 py-2.5 font-medium text-slate-700">订单笔数</th>
                    <th className="px-3 py-2.5 font-medium text-slate-700">默认系数</th>
                    <th className="px-3 py-2.5 font-medium text-slate-700">当前系数</th>
                    <th className="px-3 py-2.5 font-medium text-slate-700">操作</th>
                  </tr>
                </thead>
                <tbody>
                  {customerActivityList.map((c) => {
                    const cur = activityOverrides[c.customerId];
                    const display = cur !== undefined && cur !== "" ? cur : c.defaultCoef;
                    return (
                      <tr key={c.customerId} className="border-b border-slate-100 hover:bg-slate-50/50">
                        <td className="px-3 py-2.5 font-mono text-xs">{c.customerId}</td>
                        <td className="px-3 py-2.5 text-slate-700 max-w-xs truncate">{c.name}</td>
                        <td className="px-3 py-2.5 text-slate-500">{c.orderCount}</td>
                        <td className="px-3 py-2.5 text-slate-400">{c.defaultCoef}</td>
                        <td className="px-3 py-2.5">
                          <input
                            type="number"
                            step="0.1"
                            value={display}
                            onChange={(e) =>
                              setActivityOverrides((p) => ({ ...p, [c.customerId]: e.target.value }))
                            }
                            className="w-24 px-2 py-1 border border-slate-300 rounded-md text-sm"
                          />
                        </td>
                        <td className="px-3 py-2.5">
                          <button
                            onClick={() =>
                              setActivityOverrides((p) => {
                                const np = { ...p };
                                delete np[c.customerId];
                                return np;
                              })
                            }
                            className="text-slate-500 hover:text-slate-900 text-xs inline-flex items-center gap-1"
                          >
                            <RotateCcw size={12} /> 恢复
                          </button>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </div>

      {/* 底部按钮 */}
      <div className="flex justify-between mt-6">
        <button
          onClick={onBack}
          className="px-5 py-3 bg-white border border-slate-300 text-slate-700 rounded-xl font-medium hover:bg-slate-50 transition-all flex items-center gap-2"
        >
          <ChevronLeft size={18} /> 上一步
        </button>
        <button
          onClick={handleSubmit}
          className="px-6 py-3 bg-slate-900 text-white rounded-xl font-semibold hover:bg-slate-800 transition-all flex items-center gap-2"
        >
          <Calculator size={18} /> 开始核算
        </button>
      </div>
    </div>
  );
}

// ============================================================
// 步骤 4：结果展示
// ============================================================
function ResultPanel({ result, params, data, onBack, onRestart }) {
  const { personResults, warnings, validRegular, tradeLines, dropped, allLines } = result;
  const [activeReport, setActiveReport] = useState("summary");
  const [expandedPerson, setExpandedPerson] = useState(null);

  // 全员合计
  const totals = personResults.reduce(
    (a, p) => ({
      base: a.base + p.baseSalary,
      perf: a.perf + p.perfWage,
      sale: a.sale + p.finalSaleCommission,
      premium: a.premium + p.totalPremium,
      trade: a.trade + p.totalTrade,
      total: a.total + p.totalSalary,
    }),
    { base: 0, perf: 0, sale: 0, premium: 0, trade: 0, total: 0 }
  );

  const month = data.settlementData.month || "";

  // 生成 Markdown 报告
  const exportMarkdown = useCallback(() => {
    const lines = [];
    lines.push(`# 康乃尔业务人员月度薪酬核算报告`);
    lines.push(`**核算月份**: ${month}`);
    lines.push(`**生成时间**: ${new Date().toLocaleString("zh-CN")}`);
    lines.push("");
    lines.push("## 一、薪酬汇总表");
    lines.push("");
    lines.push("| 业务员 | 基本工资 | 绩效工资 | 销售提成 | 溢价提成 | 贸易提成 | **合计** |");
    lines.push("|---|---|---|---|---|---|---|");
    personResults.forEach((p) => {
      lines.push(
        `| ${p.name} | ${fmtCNY(p.baseSalary)} | ${fmtCNY(p.perfWage)} | ${fmtCNY(p.finalSaleCommission)} | ${fmtCNY(p.totalPremium)} | ${fmtCNY(p.totalTrade)} | **${fmtCNY(p.totalSalary)}** |`
      );
    });
    lines.push(
      `| **合计** | **${fmtCNY(totals.base)}** | **${fmtCNY(totals.perf)}** | **${fmtCNY(totals.sale)}** | **${fmtCNY(totals.premium)}** | **${fmtCNY(totals.trade)}** | **${fmtCNY(totals.total)}** |`
    );
    lines.push("");
    lines.push("## 二、交叉验证（万元）");
    lines.push("");
    lines.push("| 业务员 | 全部收款 | 常规有效 | 贸易收款 | 剔除收款 | 差异 |");
    lines.push("|---|---|---|---|---|---|");
    personResults.forEach((p) => {
      const sum = p.validReceipt + p.tradeReceipt + p.droppedReceipt;
      const diff = p.allReceipt - sum;
      lines.push(
        `| ${p.name} | ${fmtWan(p.allReceipt)} | ${fmtWan(p.validReceipt)} | ${fmtWan(p.tradeReceipt)} | ${fmtWan(p.droppedReceipt)} | ${fmtWan(diff)} |`
      );
    });
    lines.push("");

    // 逐人详细
    personResults.forEach((p) => {
      lines.push(`## ${p.name} · 详细报告`);
      lines.push(`**当月薪酬总计**：${fmtCNY(p.totalSalary)} 元`);
      lines.push("");
      lines.push("### 薪酬明细");
      lines.push("| 项目 | 金额（元） |");
      lines.push("|---|---|");
      lines.push(`| 基本工资 | ${fmtCNY(p.baseSalary)} |`);
      lines.push(`| 绩效工资 | ${fmtCNY(p.perfWage)} |`);
      lines.push(`| 销售提成 | ${fmtCNY(p.finalSaleCommission)} |`);
      lines.push(`| 溢价提成 | ${fmtCNY(p.totalPremium)} |`);
      lines.push(`| 贸易提成 | ${fmtCNY(p.totalTrade)} |`);
      lines.push(`| **合计** | **${fmtCNY(p.totalSalary)}** |`);
      lines.push("");
      lines.push("### 绩效工资推演");
      lines.push(
        `(80% × min(${fmtPct(p.completionRate)}, 100%) + 20% × ${fmtNum(p.perfScore)}) × ${fmtCNY(p.perfBase)} = **${fmtCNY(p.perfWage)}** 元`
      );
      lines.push("");
      if (p.myRegular.length > 0) {
        lines.push("### 销售提成推演");
        lines.push(`- 逐笔销售提成合计：${fmtCNY(p.totalSaleCommission)} 元`);
        lines.push(`- 总提成基数：${fmtCNY(p.totalCommBase)} 元`);
        lines.push(`- 加权平均提成系数：${fmtPct(p.weightedAvgRate, 4)}`);
        lines.push(`- 抵扣金额：${fmtCNY(p.deduction)} × ${fmtPct(p.weightedAvgRate, 4)} = ${fmtCNY(p.deductionAmt)} 元`);
        lines.push(`- **最终销售提成**：${fmtCNY(p.totalSaleCommission)} − ${fmtCNY(p.deductionAmt)} = **${fmtCNY(p.finalSaleCommission)}** 元`);
        lines.push("");
      }
      if (p.myTrade.length > 0) {
        lines.push("### 贸易订单明细");
        lines.push("| 订单号 | 物料 | 收款金额 | 品类 | 毛利 | 比例 | 提成 |");
        lines.push("|---|---|---|---|---|---|---|");
        p.myTrade.forEach((l) => {
          lines.push(
            `| ${l.销售订单} | ${l.物料描述} | ${fmtCNY(l.收款金额)} | ${l._tradeCategory || "—"} | ${fmtCNY(l._tradeMargin)} | ${l._tradeRate ? fmtPct(l._tradeRate) : "—"} | ${fmtCNY(l._tradeCommission)} |`
          );
        });
        lines.push("");
      }
    });

    if (warnings.length > 0) {
      lines.push("## 异常预警");
      lines.push("");
      warnings.forEach((w) => {
        lines.push(`- **[${w.type}]** ${w.msg}`);
      });
    }

    const blob = new Blob([lines.join("\n")], { type: "text/markdown;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `康乃尔薪酬核算_${month}_${new Date().toISOString().slice(0, 10)}.md`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }, [personResults, warnings, totals, month]);

  // CSV 导出（薪酬汇总）
  const exportCSV = useCallback(() => {
    const rows = [];
    rows.push(["业务员", "基本工资", "绩效工资", "销售提成", "溢价提成", "贸易提成", "合计"]);
    personResults.forEach((p) => {
      rows.push([
        p.name,
        p.baseSalary.toFixed(2),
        p.perfWage.toFixed(2),
        p.finalSaleCommission.toFixed(2),
        p.totalPremium.toFixed(2),
        p.totalTrade.toFixed(2),
        p.totalSalary.toFixed(2),
      ]);
    });
    rows.push([
      "合计",
      totals.base.toFixed(2),
      totals.perf.toFixed(2),
      totals.sale.toFixed(2),
      totals.premium.toFixed(2),
      totals.trade.toFixed(2),
      totals.total.toFixed(2),
    ]);
    const csv = "\uFEFF" + rows.map((r) => r.map((c) => `"${c}"`).join(",")).join("\r\n");
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `康乃尔薪酬汇总_${month}.csv`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }, [personResults, totals, month]);

  return (
    <div className="max-w-7xl mx-auto px-4 pb-12">
      <div className="text-center mb-6">
        <div className="inline-flex items-center gap-2 mb-2">
          <CheckCircle2 className="text-emerald-600" size={22} />
          <h2 className="text-2xl font-bold text-slate-900">核算完成 · {month}</h2>
        </div>
        <p className="text-slate-500 text-sm">共 {personResults.length} 名业务员，{validRegular.length + tradeLines.length} 笔有效订单</p>
      </div>

      {/* 顶部 KPI 卡片 */}
      <div className="grid grid-cols-2 md:grid-cols-6 gap-3 mb-6">
        {[
          { label: "薪酬总额", value: totals.total, accent: "bg-slate-900 text-white" },
          { label: "基本工资", value: totals.base },
          { label: "绩效工资", value: totals.perf },
          { label: "销售提成", value: totals.sale },
          { label: "溢价提成", value: totals.premium },
          { label: "贸易提成", value: totals.trade },
        ].map((k, i) => (
          <div key={i} className={`rounded-xl p-4 border ${k.accent || "bg-white border-slate-200"}`}>
            <div className={`text-xs font-medium ${k.accent ? "text-slate-300" : "text-slate-500"}`}>{k.label}</div>
            <div className="text-lg md:text-xl font-bold mt-1">¥ {fmtNum(k.value, 2)}</div>
          </div>
        ))}
      </div>

      {/* 报告 Tab */}
      <div className="flex flex-wrap gap-2 border-b border-slate-200 mb-6">
        {[
          { id: "summary", label: "薪酬汇总", icon: Coins },
          { id: "validation", label: "交叉验证", icon: CheckCircle2 },
          { id: "params", label: "参数总览", icon: Settings },
          { id: "details", label: "逐人详细", icon: Users },
          { id: "warnings", label: `异常预警 (${warnings.length})`, icon: AlertTriangle },
        ].map((t) => {
          const Icon = t.icon;
          const active = activeReport === t.id;
          return (
            <button
              key={t.id}
              onClick={() => setActiveReport(t.id)}
              className={`px-4 py-2.5 -mb-px border-b-2 font-medium text-sm flex items-center gap-2 ${
                active ? "border-slate-900 text-slate-900" : "border-transparent text-slate-500 hover:text-slate-700"
              }`}
            >
              <Icon size={16} />
              {t.label}
            </button>
          );
        })}
      </div>

      {activeReport === "summary" && (
        <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden">
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead>
                <tr className="bg-slate-50 border-b border-slate-200">
                  <th className="px-4 py-3 text-left font-semibold text-slate-700">业务员</th>
                  <th className="px-4 py-3 text-right font-semibold text-slate-700">基本工资</th>
                  <th className="px-4 py-3 text-right font-semibold text-slate-700">绩效工资</th>
                  <th className="px-4 py-3 text-right font-semibold text-slate-700">销售提成</th>
                  <th className="px-4 py-3 text-right font-semibold text-slate-700">溢价提成</th>
                  <th className="px-4 py-3 text-right font-semibold text-slate-700">贸易提成</th>
                  <th className="px-4 py-3 text-right font-semibold text-slate-900 bg-slate-100">合计</th>
                </tr>
              </thead>
              <tbody>
                {personResults.map((p) => (
                  <tr key={p.name} className="border-b border-slate-100 hover:bg-slate-50/50">
                    <td className="px-4 py-3 font-semibold text-slate-900">{p.name}</td>
                    <td className="px-4 py-3 text-right text-slate-700">{fmtCNY(p.baseSalary)}</td>
                    <td className="px-4 py-3 text-right text-slate-700">{fmtCNY(p.perfWage)}</td>
                    <td className="px-4 py-3 text-right text-slate-700">{fmtCNY(p.finalSaleCommission)}</td>
                    <td className="px-4 py-3 text-right text-slate-700">{fmtCNY(p.totalPremium)}</td>
                    <td className="px-4 py-3 text-right text-slate-700">{fmtCNY(p.totalTrade)}</td>
                    <td className="px-4 py-3 text-right font-bold text-slate-900 bg-slate-50">{fmtCNY(p.totalSalary)}</td>
                  </tr>
                ))}
                <tr className="bg-slate-900 text-white font-semibold">
                  <td className="px-4 py-3">合计</td>
                  <td className="px-4 py-3 text-right">{fmtCNY(totals.base)}</td>
                  <td className="px-4 py-3 text-right">{fmtCNY(totals.perf)}</td>
                  <td className="px-4 py-3 text-right">{fmtCNY(totals.sale)}</td>
                  <td className="px-4 py-3 text-right">{fmtCNY(totals.premium)}</td>
                  <td className="px-4 py-3 text-right">{fmtCNY(totals.trade)}</td>
                  <td className="px-4 py-3 text-right">{fmtCNY(totals.total)}</td>
                </tr>
              </tbody>
            </table>
          </div>
          <div className="p-4 flex flex-wrap gap-2 bg-slate-50/50 border-t border-slate-200">
            <button
              onClick={exportCSV}
              className="px-4 py-2 bg-white border border-slate-300 rounded-lg text-sm font-medium hover:bg-slate-50 inline-flex items-center gap-1.5"
            >
              <Download size={14} /> 导出 CSV (汇总表)
            </button>
            <button
              onClick={exportMarkdown}
              className="px-4 py-2 bg-white border border-slate-300 rounded-lg text-sm font-medium hover:bg-slate-50 inline-flex items-center gap-1.5"
            >
              <Download size={14} /> 导出 Markdown (完整报告)
            </button>
          </div>
        </div>
      )}

      {activeReport === "validation" && (
        <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden">
          <div className="px-5 py-3 border-b border-slate-200 bg-slate-50/50">
            <p className="text-xs text-slate-600">
              单位：万元（保留 4 位小数）。<b>全部收款 = 常规有效 + 贸易收款 + 剔除收款</b>
            </p>
          </div>
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead>
                <tr className="bg-slate-50 border-b border-slate-200">
                  <th className="px-4 py-3 text-left font-semibold text-slate-700">业务员</th>
                  <th className="px-4 py-3 text-right font-semibold text-slate-700">全部收款</th>
                  <th className="px-4 py-3 text-right font-semibold text-slate-700">常规有效</th>
                  <th className="px-4 py-3 text-right font-semibold text-slate-700">贸易收款</th>
                  <th className="px-4 py-3 text-right font-semibold text-slate-700">剔除收款</th>
                  <th className="px-4 py-3 text-right font-semibold text-slate-700">差异</th>
                  <th className="px-4 py-3 text-center font-semibold text-slate-700">校验</th>
                </tr>
              </thead>
              <tbody>
                {personResults.map((p) => {
                  const sum = p.validReceipt + p.tradeReceipt + p.droppedReceipt;
                  const diff = p.allReceipt - sum;
                  const ok = Math.abs(diff) < 0.01;
                  return (
                    <tr key={p.name} className="border-b border-slate-100">
                      <td className="px-4 py-3 font-semibold text-slate-900">{p.name}</td>
                      <td className="px-4 py-3 text-right tabular-nums">{fmtWan(p.allReceipt)}</td>
                      <td className="px-4 py-3 text-right tabular-nums">{fmtWan(p.validReceipt)}</td>
                      <td className="px-4 py-3 text-right tabular-nums text-amber-700">{fmtWan(p.tradeReceipt)}</td>
                      <td className="px-4 py-3 text-right tabular-nums text-slate-400">{fmtWan(p.droppedReceipt)}</td>
                      <td className="px-4 py-3 text-right tabular-nums">{fmtWan(diff)}</td>
                      <td className="px-4 py-3 text-center">
                        {ok ? (
                          <CheckCircle2 className="inline text-emerald-600" size={16} />
                        ) : (
                          <AlertTriangle className="inline text-amber-600" size={16} />
                        )}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {activeReport === "params" && (
        <div className="space-y-4">
          {/* 个人参数 */}
          <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden">
            <div className="px-5 py-3 border-b border-slate-200 bg-slate-50/50 font-semibold text-slate-900">
              人员参数
            </div>
            <div className="overflow-x-auto">
              <table className="w-full text-sm">
                <thead>
                  <tr className="bg-slate-50 border-b border-slate-200">
                    <th className="px-4 py-2.5 text-left text-slate-700">人员</th>
                    <th className="px-4 py-2.5 text-left text-slate-700">等级</th>
                    <th className="px-4 py-2.5 text-right text-slate-700">基本工资</th>
                    <th className="px-4 py-2.5 text-right text-slate-700">绩效基数</th>
                    <th className="px-4 py-2.5 text-right text-slate-700">完成率(封顶)</th>
                    <th className="px-4 py-2.5 text-right text-slate-700">绩效分</th>
                    <th className="px-4 py-2.5 text-right text-slate-700">抵扣总额</th>
                  </tr>
                </thead>
                <tbody>
                  {personResults.map((p) => (
                    <tr key={p.name} className="border-b border-slate-100">
                      <td className="px-4 py-2.5 font-semibold">{p.name}</td>
                      <td className="px-4 py-2.5">{p.level}</td>
                      <td className="px-4 py-2.5 text-right">{fmtCNY(p.baseSalary)}</td>
                      <td className="px-4 py-2.5 text-right">{fmtCNY(p.perfBase)}</td>
                      <td className="px-4 py-2.5 text-right">{fmtPct(p.cappedRate)}</td>
                      <td className="px-4 py-2.5 text-right">{fmtNum(p.perfScore)}</td>
                      <td className="px-4 py-2.5 text-right">{fmtCNY(p.deduction)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>

          {/* 系数表 */}
          <div className="grid md:grid-cols-2 gap-4">
            <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden">
              <div className="px-5 py-3 border-b border-slate-200 bg-slate-50/50 font-semibold text-slate-900">
                销售提成系数表
              </div>
              <table className="w-full text-sm">
                <tbody>
                  {Object.entries(PRODUCT_LINE_RATES).map(([k, v]) => (
                    <tr key={k} className="border-b border-slate-100">
                      <td className="px-4 py-2 text-slate-700">{k}</td>
                      <td className="px-4 py-2 text-right tabular-nums font-semibold">{fmtPct(v, 2)}</td>
                    </tr>
                  ))}
                  <tr className="bg-amber-50 border-b border-amber-100">
                    <td className="px-4 py-2 text-amber-700">负毛利订单（覆盖产品线）</td>
                    <td className="px-4 py-2 text-right tabular-nums font-semibold text-amber-700">
                      {fmtPct(NEGATIVE_MARGIN_RATE, 2)}
                    </td>
                  </tr>
                </tbody>
              </table>
            </div>
            <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden">
              <div className="px-5 py-3 border-b border-slate-200 bg-slate-50/50 font-semibold text-slate-900">
                溢价提成系数表
              </div>
              <table className="w-full text-sm">
                <tbody>
                  <tr className="border-b border-slate-100">
                    <td className="px-4 py-2 text-slate-700">溢价率 ≤ 0</td>
                    <td className="px-4 py-2 text-right tabular-nums">0</td>
                  </tr>
                  <tr className="border-b border-slate-100">
                    <td className="px-4 py-2 text-slate-700">0 &lt; 溢价率 ≤ 5%</td>
                    <td className="px-4 py-2 text-right tabular-nums font-semibold">0.15</td>
                  </tr>
                  <tr className="border-b border-slate-100">
                    <td className="px-4 py-2 text-slate-700">5% &lt; 溢价率 ≤ 10%</td>
                    <td className="px-4 py-2 text-right tabular-nums font-semibold">0.20</td>
                  </tr>
                  <tr className="border-b border-slate-100">
                    <td className="px-4 py-2 text-slate-700">溢价率 &gt; 10%</td>
                    <td className="px-4 py-2 text-right tabular-nums font-semibold">0.30</td>
                  </tr>
                </tbody>
              </table>
            </div>
            <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden">
              <div className="px-5 py-3 border-b border-slate-200 bg-slate-50/50 font-semibold text-slate-900">
                贸易业务提成比例
              </div>
              <table className="w-full text-sm">
                <tbody>
                  {Object.entries(TRADE_RATES).map(([k, v]) => (
                    <tr key={k} className="border-b border-slate-100">
                      <td className="px-4 py-2 text-slate-700">{k}</td>
                      <td className="px-4 py-2 text-right tabular-nums font-semibold">{fmtPct(v, 0)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden">
              <div className="px-5 py-3 border-b border-slate-200 bg-slate-50/50 font-semibold text-slate-900">
                居间单价 / 负毛利订单（本次生效）
              </div>
              <div className="p-4 text-sm">
                <div className="mb-3">
                  <div className="text-xs text-slate-500 mb-1">居间单价（{params.priceDeductions.length} 条）</div>
                  {params.priceDeductions.length === 0 ? (
                    <div className="text-slate-400 text-xs">无</div>
                  ) : (
                    <div className="flex flex-wrap gap-1.5">
                      {params.priceDeductions.map((d, i) => (
                        <span key={i} className="inline-flex items-center gap-1 px-2 py-0.5 bg-slate-100 rounded text-xs">
                          <span className="font-mono">{d.customerId}</span>
                          <span className="text-slate-500">{d.pricePerKg}元/kg</span>
                        </span>
                      ))}
                    </div>
                  )}
                </div>
                <div>
                  <div className="text-xs text-slate-500 mb-1">负毛利订单（{params.negativeOrders.length} 条）</div>
                  {params.negativeOrders.length === 0 ? (
                    <div className="text-slate-400 text-xs">无</div>
                  ) : (
                    <div className="flex flex-wrap gap-1.5">
                      {params.negativeOrders.map((o, i) => (
                        <span key={i} className="px-2 py-0.5 bg-amber-50 text-amber-700 rounded text-xs font-mono">
                          {o}
                        </span>
                      ))}
                    </div>
                  )}
                </div>
              </div>
            </div>
          </div>
        </div>
      )}

      {activeReport === "details" && (
        <div className="space-y-3">
          {personResults.map((p) => {
            const expanded = expandedPerson === p.name;
            return (
              <div key={p.name} className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden">
                <button
                  onClick={() => setExpandedPerson(expanded ? null : p.name)}
                  className="w-full px-5 py-4 flex items-center justify-between hover:bg-slate-50 text-left"
                >
                  <div className="flex items-center gap-4">
                    <div className="w-10 h-10 rounded-full bg-slate-100 text-slate-700 flex items-center justify-center font-semibold">
                      {p.name.slice(0, 1)}
                    </div>
                    <div>
                      <div className="font-semibold text-slate-900">{p.name}</div>
                      <div className="text-xs text-slate-500">
                        {p.level} · 完成率 {fmtPct(p.completionRate)} · 绩效分 {fmtNum(p.perfScore)} · 常规{p.myRegular.length}笔 · 贸易{p.myTrade.length}笔
                      </div>
                    </div>
                  </div>
                  <div className="flex items-center gap-4">
                    <div className="text-right">
                      <div className="text-xs text-slate-500">月度薪酬</div>
                      <div className="text-xl font-bold text-slate-900">¥ {fmtNum(p.totalSalary)}</div>
                    </div>
                    <ChevronRight
                      size={18}
                      className={`text-slate-400 transition-transform ${expanded ? "rotate-90" : ""}`}
                    />
                  </div>
                </button>
                {expanded && (
                  <div className="border-t border-slate-200 p-5 space-y-5 bg-slate-50/30">
                    {/* 薪酬明细 */}
                    <div>
                      <h4 className="text-sm font-semibold text-slate-900 mb-2">薪酬构成</h4>
                      <div className="grid grid-cols-2 md:grid-cols-5 gap-2">
                        {[
                          { label: "基本工资", v: p.baseSalary },
                          { label: "绩效工资", v: p.perfWage },
                          { label: "销售提成", v: p.finalSaleCommission },
                          { label: "溢价提成", v: p.totalPremium },
                          { label: "贸易提成", v: p.totalTrade },
                        ].map((x, i) => (
                          <div key={i} className="bg-white border border-slate-200 rounded-lg p-2.5">
                            <div className="text-xs text-slate-500">{x.label}</div>
                            <div className="font-semibold text-slate-900 text-sm">¥ {fmtNum(x.v)}</div>
                          </div>
                        ))}
                      </div>
                    </div>

                    {/* 绩效工资推演 */}
                    <div className="bg-white rounded-lg border border-slate-200 p-4">
                      <h4 className="text-sm font-semibold text-slate-900 mb-2">绩效工资推演</h4>
                      <div className="text-sm text-slate-700 font-mono">
                        (80% × min({fmtPct(p.completionRate)}, 100%) + 20% × {fmtNum(p.perfScore)}) × {fmtCNY(p.perfBase)} = <b>{fmtCNY(p.perfWage)}</b> 元
                      </div>
                    </div>

                    {/* 销售提成推演 */}
                    {p.myRegular.length > 0 && (
                      <div className="bg-white rounded-lg border border-slate-200 p-4">
                        <h4 className="text-sm font-semibold text-slate-900 mb-2">销售提成推演</h4>
                        <div className="text-sm space-y-1 text-slate-700">
                          <div>逐笔销售提成合计：<b>{fmtCNY(p.totalSaleCommission)}</b> 元</div>
                          <div>总提成基数：<b>{fmtCNY(p.totalCommBase)}</b> 元</div>
                          <div>加权平均提成系数：<b>{fmtPct(p.weightedAvgRate, 4)}</b></div>
                          <div>抵扣金额：{fmtCNY(p.deduction)} × {fmtPct(p.weightedAvgRate, 4)} = <b>{fmtCNY(p.deductionAmt)}</b> 元</div>
                          <div className="pt-2 border-t border-slate-100 font-medium">
                            最终销售提成 = {fmtCNY(p.totalSaleCommission)} − {fmtCNY(p.deductionAmt)} = <span className="text-emerald-700 font-bold">{fmtCNY(p.finalSaleCommission)}</span> 元
                          </div>
                        </div>
                      </div>
                    )}

                    {/* 常规订单逐笔 */}
                    {p.myRegular.length > 0 && (
                      <div className="bg-white rounded-lg border border-slate-200 overflow-hidden">
                        <div className="px-4 py-2.5 bg-slate-50 border-b border-slate-200 text-sm font-semibold text-slate-900">
                          常规订单逐笔明细 ({p.myRegular.length} 笔)
                        </div>
                        <div className="overflow-x-auto">
                          <table className="w-full text-xs">
                            <thead className="bg-slate-50/50">
                              <tr className="text-left text-slate-600">
                                <th className="px-2 py-2">订单号</th>
                                <th className="px-2 py-2">类别</th>
                                <th className="px-2 py-2">售达方</th>
                                <th className="px-2 py-2 text-right">原数量</th>
                                <th className="px-2 py-2">单位</th>
                                <th className="px-2 py-2 text-right">核定kg</th>
                                <th className="px-2 py-2 text-right">收款</th>
                                <th className="px-2 py-2 text-right">运费/kg</th>
                                <th className="px-2 py-2 text-right">出厂价</th>
                                <th className="px-2 py-2 text-right">基价</th>
                                <th className="px-2 py-2 text-right">提成基数</th>
                                <th className="px-2 py-2 text-right">系数</th>
                                <th className="px-2 py-2 text-right">活跃度</th>
                                <th className="px-2 py-2 text-right">销售提成</th>
                                <th className="px-2 py-2 text-right">居间</th>
                                <th className="px-2 py-2 text-right">溢价率</th>
                                <th className="px-2 py-2 text-right">溢价提成</th>
                              </tr>
                            </thead>
                            <tbody>
                              {p.myRegular.map((l) => (
                                <tr
                                  key={l._rowIdx}
                                  className={`border-t border-slate-100 ${
                                    l._isNegative ? "bg-amber-50/40" : ""
                                  }`}
                                >
                                  <td className="px-2 py-1.5 font-mono">{l.销售订单}</td>
                                  <td className="px-2 py-1.5">
                                    {l.类别}
                                    {l._isNegative && <span className="ml-1 text-amber-700">⚠负毛利</span>}
                                  </td>
                                  <td className="px-2 py-1.5 max-w-[120px] truncate" title={l.售达方描述}>
                                    {l.售达方描述}
                                  </td>
                                  <td className="px-2 py-1.5 text-right tabular-nums">{fmtNum(l.数量, 1)}</td>
                                  <td className="px-2 py-1.5">{l.单位}</td>
                                  <td className="px-2 py-1.5 text-right tabular-nums">{fmtNum(l._kg, 1)}</td>
                                  <td className="px-2 py-1.5 text-right tabular-nums">{fmtNum(l.收款金额)}</td>
                                  <td className="px-2 py-1.5 text-right tabular-nums">{fmtNum(l._freightUnit)}</td>
                                  <td className="px-2 py-1.5 text-right tabular-nums">{fmtNum(l._exFactoryPrice)}</td>
                                  <td className="px-2 py-1.5 text-right tabular-nums">{fmtNum(l._basePrice)}</td>
                                  <td className="px-2 py-1.5 text-right tabular-nums">{fmtNum(l._commBase)}</td>
                                  <td className="px-2 py-1.5 text-right tabular-nums">{fmtPct(l._commRate, 2)}</td>
                                  <td className="px-2 py-1.5 text-right tabular-nums">{l._activityCoef}</td>
                                  <td className="px-2 py-1.5 text-right tabular-nums font-semibold">{fmtNum(l._saleCommission)}</td>
                                  <td className="px-2 py-1.5 text-right tabular-nums">{l._priceDeduction || 0}</td>
                                  <td className="px-2 py-1.5 text-right tabular-nums">
                                    {l._premiumRatio !== null ? fmtPct(l._premiumRatio, 1) : "—"}
                                  </td>
                                  <td className="px-2 py-1.5 text-right tabular-nums font-semibold">{fmtNum(l._premiumCommission)}</td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>
                      </div>
                    )}

                    {/* 贸易订单 */}
                    {p.myTrade.length > 0 && (
                      <div className="bg-white rounded-lg border border-slate-200 overflow-hidden">
                        <div className="px-4 py-2.5 bg-amber-50 border-b border-amber-200 text-sm font-semibold text-amber-900">
                          贸易订单明细 ({p.myTrade.length} 笔)
                        </div>
                        <table className="w-full text-xs">
                          <thead className="bg-slate-50">
                            <tr className="text-left text-slate-600">
                              <th className="px-3 py-2">订单号</th>
                              <th className="px-3 py-2">物料</th>
                              <th className="px-3 py-2">售达方</th>
                              <th className="px-3 py-2 text-right">收款</th>
                              <th className="px-3 py-2">品类</th>
                              <th className="px-3 py-2 text-right">毛利</th>
                              <th className="px-3 py-2 text-right">比例</th>
                              <th className="px-3 py-2 text-right">提成</th>
                            </tr>
                          </thead>
                          <tbody>
                            {p.myTrade.map((l) => (
                              <tr key={l._rowIdx} className="border-t border-slate-100">
                                <td className="px-3 py-1.5 font-mono">{l.销售订单}</td>
                                <td className="px-3 py-1.5 max-w-[200px] truncate" title={l.物料描述}>
                                  {l.物料描述}
                                </td>
                                <td className="px-3 py-1.5 max-w-[120px] truncate">{l.售达方描述}</td>
                                <td className="px-3 py-1.5 text-right tabular-nums">{fmtCNY(l.收款金额)}</td>
                                <td className="px-3 py-1.5">{l._tradeCategory || "—"}</td>
                                <td className="px-3 py-1.5 text-right tabular-nums">{fmtCNY(l._tradeMargin)}</td>
                                <td className="px-3 py-1.5 text-right tabular-nums">
                                  {l._tradeRate ? fmtPct(l._tradeRate, 0) : "—"}
                                </td>
                                <td className="px-3 py-1.5 text-right tabular-nums font-semibold">{fmtCNY(l._tradeCommission)}</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    )}
                  </div>
                )}
              </div>
            );
          })}
        </div>
      )}

      {activeReport === "warnings" && (
        <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden">
          {warnings.length === 0 ? (
            <div className="p-10 text-center text-slate-500">
              <CheckCircle2 className="mx-auto text-emerald-500 mb-2" size={32} />
              <div>未发现异常 — 所有订单均已正确处理</div>
            </div>
          ) : (
            <div className="divide-y divide-slate-100">
              {warnings.map((w, i) => (
                <div key={i} className="p-4 flex gap-3">
                  <AlertTriangle size={18} className="text-amber-500 shrink-0 mt-0.5" />
                  <div className="flex-1 min-w-0">
                    <div className="text-xs font-semibold text-amber-700 mb-0.5">{w.type}</div>
                    <div className="text-sm text-slate-700">{w.msg}</div>
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>
      )}

      {/* 底部按钮 */}
      <div className="flex justify-between mt-6">
        <button
          onClick={onBack}
          className="px-5 py-3 bg-white border border-slate-300 text-slate-700 rounded-xl font-medium hover:bg-slate-50 flex items-center gap-2"
        >
          <ChevronLeft size={18} /> 返回参数
        </button>
        <button
          onClick={onRestart}
          className="px-5 py-3 bg-slate-900 text-white rounded-xl font-medium hover:bg-slate-800 flex items-center gap-2"
        >
          重新开始
        </button>
      </div>
    </div>
  );
}

// ============================================================
// 主组件
// ============================================================
export default function App() {
  const [step, setStep] = useState(1);
  const [data, setData] = useState(null); // {settlementData, perfData}
  const [params, setParams] = useState(null); // {personSettings, priceDeductions, ...}
  const [result, setResult] = useState(null);
  const [calculating, setCalculating] = useState(false);

  const handleStep1Done = (d) => {
    setData(d);
    setStep(2);
  };

  const handleStep2Done = (p) => {
    setParams(p);
    setCalculating(true);
    // 同步执行（量级小），用 setTimeout 让 UI 有反应
    setTimeout(() => {
      try {
        const res = runCalculation({
          settlements: data.settlementData.settlements,
          priceMap: data.settlementData.priceMap,
          persons: data.perfData.persons,
          ...p,
        });
        setResult(res);
        setStep(4);
      } catch (e) {
        alert("核算失败：" + (e.message || e));
      } finally {
        setCalculating(false);
      }
    }, 50);
  };

  const handleRestart = () => {
    setStep(1);
    setData(null);
    setParams(null);
    setResult(null);
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 to-slate-100">
      {/* 顶部品牌栏 */}
      <header className="bg-white border-b border-slate-200">
        <div className="max-w-7xl mx-auto px-4 py-4 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-lg bg-gradient-to-br from-slate-800 to-slate-900 text-white flex items-center justify-center font-bold">
              KNE
            </div>
            <div>
              <div className="font-bold text-slate-900">康乃尔薪酬核算系统</div>
              <div className="text-xs text-slate-500">v6.0 · 依据 2026年滨州康乃尔经营薪酬考核方案 V2.0</div>
            </div>
          </div>
          {data && (
            <div className="hidden md:flex items-center gap-3 text-sm text-slate-600">
              <span className="px-2.5 py-1 bg-slate-100 rounded-md">
                核算月份：<b className="text-slate-900">{data.settlementData.month}</b>
              </span>
              <span className="px-2.5 py-1 bg-slate-100 rounded-md">
                业务员：<b className="text-slate-900">{data.perfData.persons.length}</b>
              </span>
            </div>
          )}
        </div>
      </header>

      <Stepper step={step === 4 ? 4 : step} />

      {calculating && (
        <div className="text-center py-20">
          <div className="inline-block w-10 h-10 border-4 border-slate-200 border-t-slate-900 rounded-full animate-spin"></div>
          <div className="mt-4 text-slate-600">核算中...</div>
        </div>
      )}

      {!calculating && step === 1 && <Step1Upload onComplete={handleStep1Done} />}
      {!calculating && step === 2 && (
        <Step2Parameters data={data} onBack={() => setStep(1)} onComplete={handleStep2Done} />
      )}
      {!calculating && step === 4 && result && (
        <ResultPanel result={result} params={params} data={data} onBack={() => setStep(2)} onRestart={handleRestart} />
      )}

      {/* 底部说明 */}
      <footer className="max-w-7xl mx-auto px-4 py-8 text-center text-xs text-slate-400">
        居间单价数据已保存在本地浏览器（localStorage），下次打开仍可继续使用。清除浏览器数据会同时清除这些记录。
      </footer>
    </div>
  );
}
