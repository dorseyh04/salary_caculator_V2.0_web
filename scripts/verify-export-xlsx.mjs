// 验证个人 Excel 导出（exportPersonXlsx）：溢价系数列位置、列对齐、系数值渲染
// 用真实 8 月数据跑引擎，取崔春辉（溢价多行）做导出，mock 浏览器全局对象捕获 HTML
import * as XLSX from "xlsx";
import { readFileSync } from "node:fs";
import { calculatePerformanceWage, getSalesCommissionRate, NEGATIVE_MARGIN_RATE, parsePerformanceRows, parseSettlementRows } from "../lib/salaryRules.mjs";

// ── 抽取 runCalculation ──
const src = readFileSync(new URL("../components/SalaryApp.jsx", import.meta.url), "utf8");
const s1 = src.indexOf("const TRADE_RATES");
const e1Marker = "return { warnings, personResults, allLines, validRegular, tradeLines, dropped };\n}";
const e1 = src.indexOf(e1Marker);
const runCalculation = new Function(
  "calculatePerformanceWage", "getSalesCommissionRate", "NEGATIVE_MARGIN_RATE",
  `${src.slice(s1, e1 + e1Marker.length)}; return runCalculation;`
)(calculatePerformanceWage, getSalesCommissionRate, NEGATIVE_MARGIN_RATE);

// ── 抽取 exportPersonXlsx ──
const s2 = src.indexOf("function exportPersonXlsx");
const e2 = src.indexOf("function generatePersonMarkdown");
let capturedHtml = null;
const anchor = { href: "", download: "", click() {} };
globalThis.document = { createElement: () => anchor, body: { appendChild() {}, removeChild() {} } };
globalThis.URL.createObjectURL = () => "blob:mock";
globalThis.URL.revokeObjectURL = () => {};
const OrigBlob = globalThis.Blob;
globalThis.Blob = class extends OrigBlob { constructor(parts, opts) { super(parts, opts); capturedHtml = parts.join(""); } };
const exportPersonXlsx = new Function(`${src.slice(s2, e2)}; return exportPersonXlsx;`)();

// ── 真实数据跑引擎 ──
const DIR = "/Users/george/Vault/02_Projects/BKNE/工资薪酬/团队薪酬/8 月";
const wb1 = XLSX.read(readFileSync(`${DIR}/8月回款明细（含西斯）.xlsx`), { type: "buffer", cellDates: true });
const rows1 = XLSX.utils.sheet_to_json(wb1.Sheets[wb1.SheetNames.find((n) => n.includes("回款明细"))], { header: 1, defval: "", raw: true });
const settlements = parseSettlementRows(rows1);
const wb2 = XLSX.read(readFileSync(`${DIR}/8 月绩效分.xlsx`), { type: "buffer", cellDates: true });
const rows2 = XLSX.utils.sheet_to_json(wb2.Sheets[wb2.SheetNames.find((n) => n.includes("汇总") || n.includes("绩效"))], { header: 1, defval: "", raw: true });
const persons = parsePerformanceRows(rows2);
const result = runCalculation({ settlements, priceMap: new Map(), persons, personSettings: {}, tradeMarginInputs: {}, activityOverrides: {}, bonusInputs: {}, adjustmentInputs: {} });

const cui = result.personResults.find((p) => p.name === "崔春辉");
exportPersonXlsx(cui, "2026年8月");
if (!capturedHtml) throw new Error("未捕获到导出 HTML");

// ── 断言（限定在「三、常规订单明细」区块内）──
const secStart = capturedHtml.indexOf("三、常规订单明细");
const secEnd = capturedHtml.indexOf("</table>", secStart);
const sec = capturedHtml.slice(secStart, secEnd);

const idxRatio = sec.indexOf("溢价率");
const idxCoef = sec.indexOf("溢价系数");
const idxComm = sec.indexOf("溢价奖金");
console.log(`列顺序(明细区块内): 溢价率(${idxRatio}) < 溢价系数(${idxCoef}) < 溢价奖金(${idxComm}) → ${idxRatio > 0 && idxRatio < idxCoef && idxCoef < idxComm ? "✅" : "❌"}`);

// 明细表头行 = 含「核定公斤」的那个 <tr>...</tr>
const headerTr = (sec.match(/<tr>[\s\S]*?核定公斤[\s\S]*?<\/tr>/) || [""])[0];
const headerCells = (headerTr.match(/<td /g) || []).length;
const firstDataRow = sec.split("</tr>").find((seg) => seg.includes(cui.myRegular[0].销售订单)) || "";
const dataCells = (firstDataRow.match(/<td /g) || []).length;
console.log(`明细表头单元格数: ${headerCells}（应=15）→ ${headerCells === 15 ? "✅" : "❌"}`);
console.log(`首行数据单元格数: ${dataCells}（应=15）→ ${dataCells === 15 ? "✅" : "❌"}`);

// 系数值渲染：崔春辉有溢价行，应出现 0.15/0.2/0.3 中的实际档
const usedRates = [...new Set(cui.myRegular.filter((l) => l._premiumRate > 0).map((l) => l._premiumRate))];
const rendered = usedRates.filter((r) => sec.includes(`>${r}<`));
console.log(`崔春辉实际溢价系数档: [${usedRates.join(", ")}]，HTML 中渲染出: [${rendered.join(", ")}] → ${usedRates.length > 0 && rendered.length === usedRates.length ? "✅" : "❌"}`);

// 合计行单元格数（明细区块内所有 <tr> 段中含 >合计< 的那一段）
const allTrs = sec.match(/<tr>[\s\S]*?<\/tr>/g) || [];
const totTr = allTrs.find((t) => t.includes(">合计<")) || "";
const totCells = (totTr.match(/<td /g) || []).length;
console.log(`合计行单元格数: ${totCells}（应=10: colspan3+kg+amt+colspan3+colspan2+销售提成+居间空+溢价率空+溢价系数空+溢价合计）→ ${totCells === 10 ? "✅" : "❌"}`);
