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

// ── 3. 跑引擎（personSettings/tradeMarginInputs/activityOverrides 全默认；
//        注入商机兑现与误差调整测试值）──
const bonusInputs = { "崔春辉": 500 };
const adjustmentInputs = { "刘洋": -200, "张哲": 150 };
const result = runCalculation({ settlements, priceMap, persons, personSettings: {}, tradeMarginInputs: {}, activityOverrides: {}, bonusInputs, adjustmentInputs });

// ── 4. 断言与报告 ──
const { personResults, warnings, validRegular, tradeLines, dropped } = result;
const negLines = validRegular.filter((l) => l._isNegative);
console.log(`\n有效常规 ${validRegular.length} 行 / 贸易 ${tradeLines.length} 行 / 剔除 ${dropped.length} 行`);
console.log(`负毛利行: ${negLines.length}（应=9）→ ${negLines.length === 9 ? "✅" : "❌"}`);
console.log(`负毛利行系数全部为 0.25%: ${negLines.every((l) => l._commRate === 0.0025) ? "✅" : "❌"}`);
const interimApplied = validRegular.filter((l) => l._priceDeduction > 0);
console.log(`居间费生效行: ${interimApplied.length}（8月应=0）→ ${interimApplied.length === 0 ? "✅" : "❌"}`);

console.log("\n—— 全员薪酬汇总 ——");
console.log("姓名 | 基本工资 | 绩效工资 | 销售提成 | 溢价奖金 | 贸易提成 | 商机兑现 | 误差调整 | 合计");
let gTotal = 0;
for (const p of personResults) {
  gTotal += p.totalSalary;
  console.log(`${p.name} | ${p.baseSalary.toFixed(2)} | ${p.perfWage.toFixed(2)} | ${p.totalSaleCommission.toFixed(2)} | ${p.totalPremium.toFixed(2)} | ${p.totalTrade.toFixed(2)} | ${p.bonus.toFixed(2)} | ${p.adjustment.toFixed(2)} | ${p.totalSalary.toFixed(2)}`);
}
console.log(`合计薪酬总额: ${gTotal.toFixed(2)}`);

// ── 商机兑现/误差调整断言 ──
console.log("\n—— 商机兑现 / 误差调整断言 ——");
let adjOk = true;
for (const p of personResults) {
  const expectBonus = Number(bonusInputs[p.name]) || 0;
  const expectAdj = Number(adjustmentInputs[p.name]) || 0;
  const expectedTotal = p.baseSalary + p.perfWage + p.totalSaleCommission + p.totalPremium + p.totalTrade + expectBonus + expectAdj;
  const ok = Math.abs(p.totalSalary - expectedTotal) < 0.01 && p.bonus === expectBonus && p.adjustment === expectAdj;
  if (!ok) adjOk = false;
  console.log(`${p.name}: 商机=${p.bonus}(应${expectBonus}) 误差=${p.adjustment}(应${expectAdj}) 合计=${p.totalSalary.toFixed(2)}(应${expectedTotal.toFixed(2)}) → ${ok ? "✅" : "❌"}`);
}
console.log(`商机/误差计入薪酬总额: ${adjOk ? "✅ 全部正确" : "❌ 存在偏差"}`);

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
  const manualComm = Math.min(sample._adjPrice, sample._basePrice) * sample._kg * 0.0025 * sample._activityCoef;
  console.log(`\n抽样核对订单 ${sample.销售订单}: 出厂价=${sample._exFactoryPrice.toFixed(4)} 基价=${sample._basePrice} kg=${sample._kg} 活跃度=${sample._activityCoef}`);
  console.log(`  引擎销售提成=${sample._saleCommission.toFixed(2)} vs 手算=${manualComm.toFixed(2)} → ${Math.abs(sample._saleCommission - manualComm) < 0.01 ? "✅" : "❌"}`);
}

// ── 6. 新口径回归：居间费扣减销售提成基数（2026-10-12 制度修订）──
// 出厂修正单价 = 出厂售价 - 居间费；提成基数 = min(出厂修正单价, 基价) × 核定公斤
// 合成数据（不落真实业务数据到仓库）
const mkSyn = (o) => ({
  年: 2026, 月: 9, 销售订单: o.order, 订单类型描述: "", 过账日期: "", 售达方: o.cust, 售达方描述: o.cust,
  销售合同号: "", 物料编码: o.mat, 物料描述: "测试物料", 单位: "公斤", 数量: o.qty,
  含税单价: "", 含税金额: o.tax, 收款金额: o.tax, 交易日期: "", 差额: "", 备注: "", 回款月份: "2026-09",
  分类: o.cat, 类别: o.cat, 分类列名: "分类", 业务经理: "测试员",
  基价_V列: o.base, 运费单价: o.freight, 提成活跃系数: 0.9, 负毛利标记: "", 居间费: o.interim,
});
const synthetic = [
  // A: 出厂售价=(60000-0.5×1000)/1000=59.5；修正=58.5 < 基价70 → 基数=58500 → 提成=58500×0.8%×0.9=421.2
  mkSyn({ order: "TEST-A", cust: "C-TEST", mat: "M-1", qty: 1000, tax: 60000, cat: "耐高温材料", base: 70, freight: 0.5, interim: 1 }),
  // B: 居间费=0 不扣减 → 基数=59500 → 提成=428.4
  mkSyn({ order: "TEST-B", cust: "C-TEST", mat: "M-2", qty: 1000, tax: 60000, cat: "耐高温材料", base: 70, freight: 0.5, interim: 0 }),
  // C: 修正=100-1=99 > 基价95 → 基数封顶=95000 → 提成=684；溢价=(99-95)×1000×0.15=600（溢价口径不受影响）
  mkSyn({ order: "TEST-C", cust: "C-TEST", mat: "M-3", qty: 1000, tax: 100000, cat: "耐高温材料", base: 95, freight: 0, interim: 1 }),
];
const synResult = runCalculation({ settlements: synthetic, priceMap: new Map(), persons: [{ name: "测试员", defaultBaseSalary: 2000, defaultPerfBase: 2000, perfScore: 0.5, level: "一级" }], personSettings: {}, tradeMarginInputs: {}, activityOverrides: {}, bonusInputs: {}, adjustmentInputs: {} });
const synLine = (o) => synResult.validRegular.find((l) => l.销售订单 === o);
const synExpect = (label, got, want) => console.log(`${label}: 引擎=${got.toFixed(2)} 期望=${want} → ${Math.abs(got - want) < 0.01 ? "✅" : "❌"}`);
synExpect("TEST-A 修正单价", synLine("TEST-A")._adjPrice, 58.5);
synExpect("TEST-A 销售提成(居间1扣基数)", synLine("TEST-A")._saleCommission, 421.2);
synExpect("TEST-B 销售提成(居间0不扣)", synLine("TEST-B")._saleCommission, 428.4);
synExpect("TEST-C 销售提成(基数封顶)", synLine("TEST-C")._saleCommission, 684);
synExpect("TEST-C 溢价提成(口径不变)", synLine("TEST-C")._premiumCommission, 600);
if (synResult.warnings.length) { console.log("合成数据预警:", synResult.warnings.map((w) => w.msg).join("；")); }
