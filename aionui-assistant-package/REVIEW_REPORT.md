# Review Report — AionUi 邮局委办报表助手

日期：2026-07-20

## Standards

- Python、JSON、YAML 静态校验通过。
- 三个阶段脚本职责分离；已移除会绕过人工闸门的旧 `full_check.py`。
- 模板结构、具体数值和外部映射分离：模板动态读取，公司映射外置，未知路线隔离并报告。
- 全链路回归覆盖初版、拆分和盈亏；数值指标由代码生成并带断言。

未发现阻断级代码质量或安全问题。

## Spec

- PASS：Step 1—6 保持 21 列初版、问题汇总和人工闸门。
- PASS：Step 7 按模板逐 Sheet 输出 17/18 列，4 月 3245 条数据行守恒。
- PASS：Step 8 使用最近有效月份 `26年3月` 模板，动态读取车型组/路线并生成环比区。
- PASS：补齐决策树、案例库、模型要求、标注规则和证据台账。
- PARTIAL：特殊 Sheet（`二干基础`、`一干汇总`、`两趟合一趟`、`降吨位`）尚未自动生成，必须显式交给人工处理。
- PARTIAL：4 月有 3 条路线不在最近盈亏模板中，脚本选择隔离并报告，未猜测写入。
- UNVERIFIED：尚未在运行中的 AionUi/aioncore 中完成真实导入和对话验收。

## Verification

- `test_runner.py`: PASS
- `tests/business-flow/analysis.py`: PASS
- `check_env.py`: PASS
- `assistants-import.json`: valid JSON
- `annotation-rules.yaml`: 16 条规则，valid YAML

