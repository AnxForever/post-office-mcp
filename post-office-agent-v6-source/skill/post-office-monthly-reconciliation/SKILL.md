---
name: post-office-monthly-reconciliation
description: 邮局月度委办运输报表核对与生成技能。用于用户要求从草滩/港务在线测量表、车管派车数据、一干/二干价格表、财务收入表和月度模板生成车型汇总表、基础报表初版、拆分版或一干/二干盈亏表，或对这些 Excel 文件进行月度核对时使用。
---

# 邮局月度委办运输报表核对

只运行 Skill 中经批准的脚本。不编写临时数据处理代码，不用其他表格工具重算或改写结果。

## 先判断任务

- 仅要车型汇总表：执行 `build_model_summary.py`，只需草滩和港务在线表。
- 要完整月度报表或基础报表初版：执行 `run_workflow.py --phase review`。
- 要拆分版：确认用户已人工核对初版，再执行 `run_workflow.py --phase final`。
- **只要一干/二干盈亏表**：执行 `build_profit_tables.py`，只需「基础报表原版」+「财务收入表」+ 盈亏模板目录，不依赖 run-state，可与全年流程分开单独跑。该模式不建立月度运行状态。

在首次执行完整流程前，读取 [input-contract.md](references/input-contract.md) 和 [reconciliation-rules.md](references/reconciliation-rules.md)。

## 执行约束

1. 使用用户明确给出的 `YYYY-MM` 作为目标月；未给出时先请用户确认。
2. 将当次动态 Excel 文件全部列入 `code.execute.inputFiles`。
3. 每个阶段只执行一次对应脚本，只依据末尾 JSON 汇报结果。
4. 保留同一任务的 `/workspace/output/run-state.json`；第二阶段不得换工作区。
5. 如脚本返回 `ok:false`，原样汇报 `errors` 和 `next_action`，不绕过阻断。
6. Windows 执行时设置 `PYTHONIOENCODING=utf-8` 和 `PYTHONUTF8=1`。

## 仅生成车型汇总表

运行：

```python
import runpy, sys
sys.path.insert(0, "/workspace/skills/post-office-monthly-reconciliation/scripts")
sys.argv = [
    "build_model_summary.py",
    "--caotan", "/workspace/草滩在线表.xlsx",
    "--gangwu", "/workspace/港务在线表.xlsx",
    "--output-dir", "/workspace/output",
    "--month", "2026-04",
]
runpy.run_path("/workspace/skills/post-office-monthly-reconciliation/scripts/build_model_summary.py", run_name="__main__")
```

该模式不建立月度运行状态，不能直接进入最终阶段。

## 阶段一：预检并生成初版

确认完整动态输入已上传后运行：

```python
import runpy, sys
sys.path.insert(0, "/workspace/skills/post-office-monthly-reconciliation/scripts")
sys.argv = [
    "run_workflow.py",
    "--phase", "review",
    "--workspace", "/workspace",
    "--month", "2026-04",
    "--bundle-root", "/workspace/skills/post-office-monthly-reconciliation",
]
runpy.run_path("/workspace/skills/post-office-monthly-reconciliation/scripts/run_workflow.py", run_name="__main__")
```

当 `phase` 为 `awaiting_review` 时：

- 向用户提供车型汇总表和基础报表初版。
- 汇报 `metrics`、`issues` 和 `warnings`。
- 立即停止，请用户人工核对初版中的颜色行与“问题汇总”。
- 不自动继续生成拆分版或盈亏表。

## 阶段二：生成最终报表

只在用户明确确认已核对并上传修订初版后运行：

```python
import runpy, sys
sys.path.insert(0, "/workspace/skills/post-office-monthly-reconciliation/scripts")
sys.argv = [
    "run_workflow.py",
    "--phase", "final",
    "--workspace", "/workspace",
    "--review-file", "/workspace/4月_02_基础报表初版.xlsx",
    "--bundle-root", "/workspace/skills/post-office-monthly-reconciliation",
]
runpy.run_path("/workspace/skills/post-office-monthly-reconciliation/scripts/run_workflow.py", run_name="__main__")
```

如用户上传了上月一干/二干盈亏表，分别增加 `--previous-first-profit` 和 `--previous-second-profit`。**未上传则环比区留空，严禁从模板或其他来源推测上月数据。**

当 `phase` 为 `final_completed` 时，交付拆分版、一干盈亏表和二干盈亏表，并汇报校验结果。

官方基础报表母版中的“二干基础”“一干汇总”“两趟合一趟”等特殊 Sheet 也会按规则
生成：二干基础按二干分类、 一干汇总按一干分类；两趟合一趟只接受人工在修订初版中
写明的合并指令。若终端 `warnings` 列出“两趟合一趟”待人工处理，必须明确告知用户，
不能把普通线路自动生成结果宣称为官方完整等价文件。

## 仅生成一干/二干盈亏表

适用场景：用户手上已有官方「基础报表原版」（含“一干汇总”“二干基础”）和当月「财务收入表」，
只需要两张盈亏表。该模式不读写 `run-state.json`，可与完整流程分开单独执行。

**当月这两份源文件是第一基准。** 吨位档、路线、各项金额全部由它们决定；
盈亏模板是可选的版式参考，不参与任何数值计算。

必需输入（2 份）：

| 输入 | 必需结构 |
|---|---|
| 基础报表原版 | 含「一干汇总」「二干基础」两个 sheet，有 邮路名称/车牌号/封车时间/吨位/运输量/运价 |
| 财务收入表 | 表头在第 1 行；邮路级别、自委办、往返、邮路名称、流水号、车牌号、额定吨位、结算收入、普服收入 |

可选输入：

| 输入 | 作用 |
|---|---|
| `--tpl-dir` 盈亏模板目录 | 只用来追加「当月无数据的历史路线」占位行；不改变吨位档和数值 |
| `--no-template-routes` | 连占位也不要，只输出当月实际出现的路线 |

运行（纯两份源文件）：

```python
import runpy, sys
sys.path.insert(0, "/workspace/skills/post-office-monthly-reconciliation/scripts")
sys.argv = [
    "build_profit_tables.py",
    "--base", "/workspace/8月基础报表原版.xlsx",
    "--finance", "/workspace/2026.08运输局.xlsx",
    "--month", "2026-08",
    "--out", "/workspace/output",
]
runpy.run_path("/workspace/skills/post-office-monthly-reconciliation/scripts/build_profit_tables.py", run_name="__main__")
```

该脚本的口径（详见 [reconciliation-rules.md](references/reconciliation-rules.md) 的「盈亏表」一节）：

- 趟次 / 装载量 / 运费来自“一干汇总”“二干基础”，**行数与这两个 sheet 完全守恒**。
- **吨位归类：先看备注，无备注取该车该路线的最低档。**
  备注写「调度通知按30吨结算」「两辆12吨车合并为一辆30吨车」的，该行整行（趟次、装载量、运费、收入）
  归类为 30 吨档；没有备注时按该车在该路线跑过的最低吨位档归类。
  基础报表的 `吨位` 列是车辆车型档，不等于结算档。
- 结算收入来自财务表 `自委办=委办` 且**排除 `往返=返程`**的行，逐趟按（车牌 + 流水号前 4 位 MMDD）一对一分摊。
- **跨月记录按「是否出现在本月财务表」判定，不看流水号日期。** 只要该行在当月运输局表里就属于当月；
  车牌本月无派车记录、无法逐趟配对时，按邮路名称归到该路线现有最高吨位档。
- 盈亏 = 结算收入 − 运费。
- 吨位档由归类后的当月数据决定（按数值升序），模板过期档名不会带进来。
- 路线顺序按当月数据出现顺序；模板里当月无数据的路线只在传了 `--tpl-dir` 时追加为占位。

汇报规则：

- 只依据脚本末尾 JSON 的 `income_match`（exact/near/miss）、`tonnage_groups`、`unmatched_samples` 汇报。
- `miss` 不为 0 时必须告知用户具体趟次，不得用估算值补足金额。
- 未提供上月盈亏表时，底部环比区留空标注“（暂无数据）”，**禁止猜测上月数据**。
- 该模式准确率经验值：趟次 100%，结算收入约 8—9 成（分档口径差异），汇报时不得宣称与官方完全一致。

## 禁止事项

- 禁止用在线表替代车管数据作为基础报表行来源。
- 禁止在没有人工修订初版和同次 `run-state.json` 时执行最终阶段。
- 禁止猜测缺失容间、吨位、价格、路线或收入；保持空值和问题标记。
- 禁止在 Excel 打开时覆盖文件。
- 禁止在 `build_profit_tables.py` 中把“基础报表原版”以外的来源当作行来源，也禁止用它反推或改写基础报表。
