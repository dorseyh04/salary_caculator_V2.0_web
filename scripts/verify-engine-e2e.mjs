// 端到端验证：从 SalaryApp.jsx 源码抽取 runCalculation 及依赖函数（同源，非复制），
// 喂入真实 8 月回款明细 + 8 月绩效分，验证负毛利(Z)/居间费(AA)/交叉验证
import * as XLSX from "xlsx";
import { readFileSync } from "node:fs";

// ── 1. 抽取组件源码中的纯逻辑段（从 TRADE_RATES 到 runCalculation 结束）──
const src = readFileSync(new URL("../components/SalaryApp.jsx", import.meta.url), "utf8");
const start = src.indexOf("const TRADE_RATES");
const endMarker = "return { warnings, personResults, allLines, validRegular, tradeLines, dropped };\n}";
const end = src.indexOf(endMarker);
if (start < 0 || end < 0) throw new Error("源码截取失败");
const logicSrc = src.slice(start, end + endMarker.length);

// 依赖：lib 层的三个函数/常量，由 import 提供，这里显式注入
import { calculatePerformanceWage, getSalesCommissionRate, NEGATIVE_MARGIN_RATE, parsePerformanceRows, parseSettlementRows } from "../lib/salaryRules.mjs";
const factory = new Function(
  "calculatePerformanceWage", "getSalesCommissionRate", "NEGATIVE_MARGIN_RATE",
  `${logicSrc}; return { runCalculation, convertToKg, resolveBasePrice, getPremiumRate };`
);
const { runCalculation } = factory(calculatePerformanceWage, getSalesCommissionRate, NEGATIVE_MARGIN_RATE);

// ── 2. 读取真实文件 ──
const DIR = "/Users/george/Vault/02_Projects/BKNE/工资薪酬/团队薪酬/8 月";
const wb1 = XLSX.read(readFileSync(`${DIR}/8月回款明细（含西斯）.xlsx`), { type: "buffer", cellDates: true });
const s1 = wb1.SheetNames.find((n) => n.includes("回款明细")) || wb1.SheetNames[0];
const rows1 = XLSX.utils.sheet_to_json(wb1.Sheets[s1], { header: 1, defval: "", raw: true });
const settlements = parseSettlementRows(rows1);
const priceMap = new Map(); // 新模板无基价表 Sheet，V 列全部为直接数值

const wb2 = XLSX.read(readFileSync(`${DIR}/8 月绩效分.xlsx`), { type: "buffer", cellDates: true });
const s2 = wb2.SheetNames.find((n) => n.includes("汇总") || n.includes("绩效")) || wb2.SheetNames[0];
const rows2 = XLSX.utils.sheet_to_json(wb2.Sheets[s2], { header: 1, defval: "", raw: true });
const persons = parsePerformanceRows(rows2);
console.log("绩效分表人员:", persons.map((p) => `${p.name}(分=${p.perfScore},基薪=${p.defaultBaseSalary},绩效基数=${p.defaultPerfBase})`).join("、"));

// ── 3. 跑引擎（personSettings/tradeMarginInputs/activityOverrides 全默认）──
const result = runCalculation({ settlements, priceMap, persons, personSettings: {}, tradeMarginInputs: {}, activityOverrides: {} });

// ── 4. 断言与报告 ──
const { personResults, warnings, validRegular, tradeLines, dropped } = result;
const negLines = validRegular.filter((l) => l._isNegative);
console.log(`\n有效常规 ${validRegular.length} 行 / 贸易 ${tradeLines.length} 行 / 剔除 ${dropped.length} 行`);
console.log(`负毛利行: ${negLines.length}（应=9）→ ${negLines.length === 9 ? "✅" : "❌"}`);
console.log(`负毛利行系数全部为 0.25%: ${negLines.every((l) => l._commRate === 0.0025) ? "✅" : "❌"}`);
const interimApplied = validRegular.filter((l) => l._priceDeduction > 0);
console.log(`居间费生效行: ${interimApplied.length}（8月应=0）→ ${interimApplied.length === 0 ? "✅" : "❌"}`);

console.log("\n—— 全员薪酬汇总 ——");
console.log("姓名 | 基本工资 | 绩效工资 | 销售提成 | 溢价奖金 | 贸易提成 | 合计");
let gTotal = 0;
for (const p of personResults) {
  gTotal += p.totalSalary;
  console.log(`${p.name} | ${p.baseSalary.toFixed(2)} | ${p.perfWage.toFixed(2)} | ${p.totalSaleCommission.toFixed(2)} | ${p.totalPremium.toFixed(2)} | ${p.totalTrade.toFixed(2)} | ${p.totalSalary.toFixed(2)}`);
}
console.log(`合计薪酬总额: ${gTotal.toFixed(2)}`);

console.log("\n—— 交叉验证（全部=常规+贸易+剔除，单位元）——");
let allOk = true;
for (const p of personResults) {
  const sum = p.validReceipt + p.tradeReceipt + p.droppedReceipt;
  const diff = p.allReceipt - sum;
  if (Math.abs(diff) >= 0.01) allOk = false;
  console.log(`${p.name}: 全部=${p.allReceipt.toFixed(2)} 常规=${p.validReceipt.toFixed(2)} 贸易=${p.tradeReceipt.toFixed(2)} 剔除=${p.droppedReceipt.toFixed(2)} diff=${diff.toFixed(4)}`);
}
console.log(`交叉验证: ${allOk ? "✅ 全员平账" : "❌ 存在差异"}`);

console.log(`\n异常预警 ${warnings.length} 条:`);
for (const w of warnings.slice(0, 15)) console.log(`  [${w.type}] ${w.msg}`);

// ── 5. 抽样手算核对一笔负毛利行 ──
const sample = negLines[0];
if (sample) {
  const manualComm = Math.min(sample._exFactoryPrice, sample._basePrice) * sample._kg * 0.0025 * sample._activityCoef;
  console.log(`\n抽样核对订单 ${sample.销售订单}: 出厂价=${sample._exFactoryPrice.toFixed(4)} 基价=${sample._basePrice} kg=${sample._kg} 活跃度=${sample._activityCoef}`);
  console.log(`  引擎销售提成=${sample._saleCommission.toFixed(2)} vs 手算=${manualComm.toFixed(2)} → ${Math.abs(sample._saleCommission - manualComm) < 0.01 ? "✅" : "❌"}`);
}
