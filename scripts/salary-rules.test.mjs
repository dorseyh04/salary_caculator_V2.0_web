import test from "node:test";
import assert from "node:assert/strict";

import {
  calculatePerformanceWage,
  parsePerformanceRows,
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
