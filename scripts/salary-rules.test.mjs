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

test("parses 负毛利 (Z) and 居间费 (AA) columns from the new template", () => {
  const rows = [
    [
      "年", "月", "销售订单", "订单类型描述", "过账日期", "售达方", "售达方描述", "销售合同号",
      "物料编码", "物料描述", "单位", "数量", "含税单价", "含税金额", "收款金额", "交易日期",
      "差额", "备注", "回款月份", "类别", "业务经理", "基价(元/kg)", "运费单价（元/kg）",
      "提成活跃系数", "分类", "负毛利", "居间费",
    ],
    [
      2026, 8, "1005250839", "销售订单（国内）-集成创建", "2026-08-20", "C515443", "北京洁邦洁净技术有限公司",
      "1285260800024", "252080846", "吸油卷_500mm*450mm*2mm_白色_30米/卷_压点_2.43kg/卷", "卷", 20, 48.6, 972, 972,
      "20260820", 0, "合计15948", "08月", "工业吸附", "崔春辉", 20.5, 1.13, 0.9, "工业吸附", "负毛利", 0,
    ],
    [
      2026, 8, "1005250840", "销售订单（国内）-集成创建", "2026-08-20", "C515443", "北京洁邦洁净技术有限公司",
      "1285260800024", "252021293", "吸油棉_400mm*500mm*2mm_灰色_200片/箱_压点_7.2KG/箱", "箱", 50, 144, 7200, 7200,
      "20260820", 0, "", "08月", "工业吸附", "崔春辉", 20.5, 1.13, 0.9, "工业吸附", "", 2.5,
    ],
  ];

  const [s1, s2] = parseSettlementRows(rows);

  assert.equal(s1.负毛利标记, "负毛利");
  assert.equal(s1.居间费, 0);
  assert.equal(s2.负毛利标记, "");
  assert.equal(s2.居间费, 2.5);
});

test("falls back to 类别 column when 分类 cell is empty (row-level)", () => {
  const rows = [
    [
      "年", "月", "销售订单", "订单类型描述", "过账日期", "售达方", "售达方描述", "销售合同号",
      "物料编码", "物料描述", "单位", "数量", "含税单价", "含税金额", "收款金额", "交易日期",
      "差额", "备注", "回款月份", "类别", "业务经理", "基价(元/kg)", "运费单价（元/kg）",
      "提成活跃系数", "分类", "负毛利", "居间费",
    ],
    [
      2026, 8, "1005250839", "销售订单（国内）-集成创建", "2026-08-20", "C515443", "北京洁邦洁净技术有限公司",
      "1285260800024", "252080846", "吸油卷_2.43kg/卷", "卷", 20, 48.6, 972, 972,
      "20260820", 0, "", "08月", "工业吸附", "崔春辉", 20.5, 1.13, 0.9, "", "", 0,
    ],
  ];

  const [settlement] = parseSettlementRows(rows);

  assert.equal(settlement.分类, "工业吸附");
  assert.equal(getSalesCommissionRate(settlement.分类).rate, 0.005);
});

test("missing 负毛利/居间费 columns parse as empty/zero (old template compatible)", () => {
  const rows = [
    [
      "年", "月", "销售订单", "订单类型描述", "过账日期", "售达方", "售达方描述", "销售合同号",
      "物料编码", "物料描述", "单位", "数量", "含税单价", "含税金额", "收款金额", "交易日期",
      "差额", "备注", "回款月份", "类别", "业务经理", "基价(元/kg)", "运费单价（元/kg）",
      "提成活跃系数", "分类",
    ],
    [
      2026, 6, "1004858896", "销售订单（国内）-集成创建", "2026-06-25", "C523837", "河北阿木森滤纸有限公司",
      "1285260600011", "252019141", "耐高温材料\\1365mm_TEST_37g/㎡", "kg", 1733.5, 31, 53738.5, 53738.5,
      "20260614", 0, "", "2026年", "06月", "刘洋", 30.8, 0, 1, "耐高温材料",
    ],
  ];

  const [settlement] = parseSettlementRows(rows);

  assert.equal(settlement.负毛利标记, "");
  assert.equal(settlement.居间费, 0);
});
