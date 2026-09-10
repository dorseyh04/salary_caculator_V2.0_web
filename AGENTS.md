# AGENTS.md

> 本文档为项目级 Codex 指引。每次新会话 Codex 会自动读取本文件以保持上下文连贯。

## 项目定位

业务人员月度薪酬核算 Web 系统。
此系统替代原有手工 Excel 核算流程，目标是让销售总监每月只需上传 2 份 Excel + 录入少量人工字段即可一键产出全员薪酬报告。

## 技术栈

- **框架**: Next.js 15 (App Router)
- **语言**: TypeScript（主项目） + JSX（核心业务组件，便于快速迭代）
- **样式**: Tailwind CSS 3
- **包管理**: pnpm
- **核心库**:
  - `xlsx` (SheetJS) — Excel 解析
  - `lucide-react` — 图标
- **部署**: Vercel
- **目标平台**: macOS 浏览器（Safari / Chrome / Edge）

## 目录结构

```
salary_caculator_V2.0_web/
├── app/                    Next.js App Router
│   ├── layout.tsx          根布局（中文 lang、字体）
│   ├── page.tsx            入口页（仅渲染 SalaryApp）
│   └── globals.css         Tailwind directives + 全局样式
├── components/
│   └── SalaryApp.jsx       核心业务组件（2000+ 行单文件）
├── docs/
│   └── PRD_v6.0.md         业务规则 PRD（核算逻辑权威来源）
├── public/                 静态资源
├── package.json
├── tailwind.config.ts
├── tsconfig.json
└── README.md               用户文档（中文）
```

## 核心代码组织（components/SalaryApp.jsx）

文件按功能分块，从上至下顺序：

1. **常量与系数表** — 销售提成、溢价、贸易、绩效基数等所有硬编码业务参数。修改提成率改这里。
2. **工具函数** — 货币/百分比/万元格式化、重量提取正则、单位换算
3. **业务规则函数** — `getProductLineRate`、`getPremiumRate`、`resolveBasePrice`
4. **Excel 解析** — `parseSettlementWorkbook`、`parsePerformanceWorkbook`
5. **计算引擎** — `runCalculation`（接受所有参数，返回 `{ warnings, personResults, ... }`）
6. **UI 组件** — `Stepper` / `Step1Upload` / `Step2Parameters` / `ResultPanel`
7. **主组件** — `App` 默认导出

## 业务规则要点（PRD 速查）

```
月度薪酬 = 基本工资 + 绩效工资 + 销售提成 + 溢价提成 + 贸易业务提成
```

- **基价读取**: 优先 V 列直接数值 → 失败则 I 列匹配基价表 D 列
- **核定公斤**: kg 直用，否则从描述提取 `(\d+\.?\d*)\s*[Kk][Gg]`
- **销售提成基数**: `min(出厂售价, 基价) × 核定公斤`
- **加权抵扣**: `抵扣总额 × 该人加权平均提成系数`
- **溢价提成**: 用 `居间单价` 修正出厂售价后，超出基价部分按 0/15/20/30% 分档
- **贸易业务**: T 列="贸易" 单独流程，按品类 40/30/20% × 用户录入毛利
- **活跃度系数**: 仅作用于销售提成；从 X 列读取，空时默认 1
- **居间单价持久化**: localStorage 跨会话保留

## 开发约定

1. **macOS 优先** — 所有代码、配置、文档均针对 macOS 浏览器，不考虑 Windows 兼容性
2. **真实数据驱动** — 修改业务逻辑前先用实际 Excel 数据回归验证（参考 `docs/PRD_v6.0.md` 中的字段定义）
3. **PRD 是权威** — 任何业务规则疑问以 `docs/PRD_v6.0.md` 为准；修改业务规则需同步更新 PRD
4. **数据精度** — 中间计算保留完整精度，最终展示两位小数
5. **金额单位** — UI 默认元，交叉验证表用万元（4 位小数）
6. **不收集数据** — 系统是纯前端，不向任何后端发送用户上传的数据

## 常用命令

```bash
pnpm install          # 安装依赖
pnpm dev              # 启动开发服务器（默认 :3000）
pnpm build && pnpm start  # 生产构建并预览
pnpm lint             # 代码检查
```

## 部署到 Vercel

```bash
# 方式 1: 通过 GitHub 集成（推荐）
# 在 Vercel 中导入此仓库，无需任何配置即可一键部署

# 方式 2: CLI
vercel --prod
```

## 待办与扩展方向

- [ ] 历史核算记录归档（按月）
- [ ] 多月趋势图：销售提成率、加权平均率走势
- [ ] 月度绩效分明细联动（点击人名查看历史绩效分）
- [ ] 报告 PDF 导出（替代当前 Markdown）
- [ ] 服务端持久化（IndexedDB → Postgres，多人协作场景）
- [ ] 居间单价导入导出（CSV）

## 调试小技巧

- 如果某笔订单未出现在结果中，先看「异常预警」Tab 是否被识别为剔除
- 基价缺失通常是 V 列公式未刷新 + 基价表无该物料编码，可手工在 Excel 中刷新或在基价表中补登
- 单位换算失败的最常见原因：物料描述中没有 KG 标注（如 `_5KG/包` 缺失），需销售运营更新主数据
