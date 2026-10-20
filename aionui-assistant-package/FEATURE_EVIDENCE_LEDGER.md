# Feature Evidence Ledger — 邮局委办报表助手

审计基线：2026-07-20；本地 AionUi 源码与邮局 2026 年 3—5 月官方 Excel 样本。

| Feature ID | Parent / lifecycle | Feature / user outcome | Delivery state | Current state & evidence | Decision | Acceptance | Coverage | Historical gap | Remaining gap |
|---|---|---|---|---|---|---|---|---|---|
| PO-01 | 对话 / 意图识别 | 识别初版、拆分、盈亏入口并一次只收一个文件 | implemented | `assistant-rule.md` + 3 个 SKILL | ADAPT Claude Code 阶段式流程 | 文档审查；推荐提示词覆盖三个入口 | partial | inherited-unassessed | 尚未在真实 Aion 对话中跑完整文件收集 |
| PO-02 | 初版 / Step 1—6 | 生成 21 列核对初版并清楚标注问题 | operational | `build_draft.py`；4月正确版/expected/current 3245 条逐格一致 | REUSE 已验证脚本，保留结构发现/问题汇总/红橙黄 | `test_runner.py` draft 指标与 expected 全一致；`verify_draft.py` | covered | backfilled | 仅有4月逐格基准，3/5月未建立 expected |
| PO-03 | 闸门 / 人工核对 | 初版后停止，防止未经确认的数据进入下游 | implemented | 系统提示词、三个 SKILL、脚本输出均声明闸门 | ADOPT 强制阶段边界 | 静态规则审查 | partial | previously-overclaimed | Aion 模型是否实际停下尚未做运行时验证 |
| PO-04 | 拆分 / Step 7 常规路线 | 从核对后初版按模板输出17/18列常规Sheet | operational | `build_split.py` 动态复制两类表头；模板历史提示处理双Sheet | ADAPT：逐 Sheet 模板驱动，不统一列数 | 4月3245条行数守恒；3—5月表头模式审计 | covered | backfilled | 新月份出现第三种表头时会停止/需扩展 |
| PO-05 | 拆分 / 特殊业务Sheet | 处理二干基础、一干汇总、两趟合一趟、降吨位 | planned | 已识别为独立人工/业务生命周期；决策树要求显式报告 | DECLINE 自动猜测 | 文档中不得静默遗漏 | gap | inherited-unassessed | 尚未实现特殊Sheet自动生成；需要业务确认规则 |
| PO-06 | 盈亏 / Step 8 | 按最近模板生成一/二干盈亏、合计和环比 | operational | `build_pnl.py` 选最后有效月份；动态车型组/路线/额定吨位；生成环比 | ADAPT 最近模板而非第一张旧Sheet | 回归验证5/6车型组、48/23路线、分组和=本月汇总、抽样通过 | covered | backfilled | 4月有3条路线不在最近模板，当前隔离并报告 |
| PO-07 | 异常 / 标注与隔离 | 不确定项留空或隔离，方便人工二次核对 | implemented | `annotation-rules.yaml`、问题汇总、`※Sheet`、未归类路线报告 | COMBINE 数据化规则 + 模型判断 | 初版88条问题汇总；静默空值检查 | partial | inherited-unassessed | YAML 由模型读取，尚未由所有脚本统一执行 |
| PO-08 | 动态发现 / 版本变化 | 表头、价格档、车型组、路线随文件变化而适配 | operational | `inspect_sources.py`；`principles.md`；3—5月结构审计 | ADOPT runtime discovery | `analysis.py` 验证拆分和盈亏模板变化 | covered | backfilled | 未知全新结构仍需人工介入 |
| PO-09 | 知识 / 决策支持 | 用决策树和真实案例复现 Claude Code 的判断方式 | implemented | `decision-tree.md`、`casebook/` 6个案例、`model-requirements.md` | ADAPT 经验为可检索证据 | 文件存在性与规则一致性审查 | partial | backfilled | 模型是否按需检索案例需 Aion 运行时测试 |
| PO-10 | 配置 / Aion 导入 | 导入助手、技能和规则后可直接使用 | implemented | `setup.py`、`assistants-import.json`、导入指南 | ADAPT 使用公开导入/API路径 | JSON/脚本静态校验 | partial | previously-overclaimed | 未启动 Aion/aioncore 做真实导入与对话验收 |
| PO-11 | 验证 / 回归 | 改动后证明初版、拆分、盈亏不退化 | operational | `test_runner.py` + `tests/business-flow/analysis.py` | ADOPT 代码结果而非估算 | 全链路回归 exit 0；业务验证报告 PASS | covered | backfilled | 无 CI；测试依赖本机邮局数据路径 |

## 当前非覆盖项

- PO-05 特殊业务 Sheet 尚未自动化，不能宣称整个拆分正式文件所有生命周期均完全自动生成。
- PO-01/03/09/10 需要在真实 AionUi 运行时验证模型行为、技能绑定和人工闸门。
- 3 条 4 月路线未在最近盈亏模板中唯一匹配，脚本选择隔离并报告，而不是猜测灌入。

