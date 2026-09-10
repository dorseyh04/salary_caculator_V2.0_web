// 用真实 8 月回款明细验证解析层（Z 列负毛利 / AA 列居间费 / 分类 / 基价）
import * as XLSX from "xlsx";
import { readFileSync } from "node:fs";
import { parseSettlementRows, getSalesCommissionRate } from "../lib/salaryRules.mjs";

const FILE = "/Users/george/Vault/02_Projects/BKNE/工资薪酬/团队薪酬/8 月/8月回款明细（含西斯）.xlsx";

const wb = XLSX.read(readFileSync(FILE), { type: "buffer", cellDates: true });
const sheetName = wb.SheetNames.find((n) => n.includes("回款明细")) || wb.SheetNames[0];
const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, defval: "", raw: true });
const settlements = parseSettlementRows(rows);

const neg = settlements.filter((s) => String(s.负毛利标记 || "").includes("负毛利"));
const interim = settlements.filter((s) => Number(s.居间费) > 0);
const emptyClass = settlements.filter((s) => !s.分类);
const emptyBase = settlements.filter((s) => s.基价_V列 === "" || s.基价_V列 === null);
const noManager = settlements.filter((s) => !s.业务经理);

console.log("解析总行数:", settlements.length);
console.log("负毛利行数:", neg.length, "涉及订单:", [...new Set(neg.map((s) => s.销售订单))].join(", "));
console.log("居间费>0 行数:", interim.length);
console.log("分类为空的行数:", emptyClass.length);
console.log("基价_V列为空的行数:", emptyBase.length);
console.log("业务经理为空的行数:", noManager.length);

// 负毛利行抽样：系数判定应为 0.25%（引擎逻辑复刻）
console.log("\n负毛利行抽样（前 3 行）:");
for (const s of neg.slice(0, 3)) {
  const normalRate = getSalesCommissionRate(s.分类).rate;
  console.log(`  订单 ${s.销售订单} | ${s.业务经理} | 分类=${s.分类} | 常规系数=${(normalRate * 100).toFixed(2)}% → 负毛利覆盖=0.25% | 含税金额=${s.含税金额} | 基价=${s.基价_V列}`);
}

// 业务员分布
const mgr = {};
for (const s of settlements) mgr[s.业务经理] = (mgr[s.业务经理] || 0) + 1;
console.log("\n业务员行数分布:", mgr);

// 分类分布与系数
const cls = {};
for (const s of settlements) { cls[s.分类] = (cls[s.分类] || 0) + 1; }
console.log("\n分类分布:", cls);
