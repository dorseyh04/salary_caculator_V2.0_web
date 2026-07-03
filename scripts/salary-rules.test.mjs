import test from "node:test";
import assert from "node:assert/strict";

import {
  calculatePerformanceWage,
  getSalesCommissionRate,
  parsePerformanceRows,
  parseSettlementRows,
} from "../lib/salaryRules.mjs";

test("parses the uploaded monthly performance score workbook shape", () => {
  const rows = [
    ["6月 绩效汇总", "", "", "", ""],
    ["姓名", "当月绩效分", "基本工资标准", "绩效等级", "绩效工资基数"],
    ["崔春辉", 0.6, 2000, "二级", 3000],
    ["刘洋", 0.51, 2000, "一级", 2000],
  ];

  assert.deepEqual(parsePerformanceRows(rows), [
    {
      name: "崔春辉",
      defaultBaseSalary: 2000,
      defaultPerfBase: 3000,
      perfScore: 0.6,
      level: "二级",
    },
    {
      name: "刘洋",
      defaultBaseSalary: 2000,
      defaultPerfBase: 2000,
      perfScore: 0.51,
      level: "一级",
    },
  ]);
});

test("calculates monthly performance wage from base times monthly score", () => {
  assert.equal(calculatePerformanceWage({ perfBase: 3000, perfScore: 0.6 }), 1800);
  assert.equal(calculatePerformanceWage({ perfBase: 2000, perfScore: 0.51 }), 1020);
});

test("uses policy sales commission rates by settlement classification", () => {
  assert.equal(getSalesCommissionRate("普通个体防护").rate, 0.005);
  assert.equal(getSalesCommissionRate("工业吸附").rate, 0.005);
  assert.equal(getSalesCommissionRate("液滤").rate, 0.005);
  assert.equal(getSalesCommissionRate("空滤").rate, 0.005);
  assert.equal(getSalesCommissionRate("生活擦拭").rate, 0.006);
  assert.equal(getSalesCommissionRate("高端个体防护").rate, 0.007);
  assert.equal(getSalesCommissionRate("耐高温材料").rate, 0.008);
  assert.equal(getSalesCommissionRate("透气弹性材料").rate, 0.01);
});

test("parses settlement classification from the 分类 column for commission rates", () => {
  const rows = [
    [
      "年",
      "月",
      "销售订单",
      "订单类型描述",
      "过账日期",
      "售达方",
      "售达方描述",
      "销售合同号",
      "物料编码",
      "物料描述",
      "单位",
      "数量",
      "含税单价",
      "含税金额",
      "收款金额",
      "交易日期",
      "差额",
      "备注",
      "回款月份",
      "类别",
      "业务经理",
      "基价(元/kg)",
      "运费单价（元/kg）",
      "提成活跃系数",
      "分类",
    ],
    [
      2026,
      6,
      "1004858896",
      "销售订单（国内）-集成创建",
      "2026-06-25",
      "C523837",
      "河北阿木森滤纸有限公司",
      "1285260600011",
      "252019141",
      "耐高温材料\\1365mm_TEST_37g/㎡",
      "kg",
      1733.5,
      31,
      53738.5,
      53738.5,
      "20260614",
      0,
      "",
      "2026年",
      "06月",
      "刘洋",
      30.8,
      0,
      1,
      "耐高温材料",
    ],
  ];

  const [settlement] = parseSettlementRows(rows);

  assert.equal(settlement.类别, "耐高温材料");
  assert.equal(settlement.分类, "耐高温材料");
  assert.equal(getSalesCommissionRate(settlement.分类).rate, 0.008);
});
