---
name: post-office-monthly-reconciliation
description: 邮局月度委办运输报表核对与生成技能。用于用户要求从草滩/港务在线测量表、车管派车数据、一干/二干价格表、财务收入表和月度模板生成车型汇总表、基础报表初版、拆分版或一干/二干盈亏表，或对这些 Excel 文件进行月度核对时使用。
---

# 邮局月度委办运输报表核对

只运行 Skill 中经批准的脚本。不编写临时数据处理代码，不用其他表格工具重算或改写结果。

## 先判断任务

- 仅要车型汇总表：执行 `build_model_summary.py`，只需草滩和港务在线表。
- 要完整月度报表或基础报表初版：执行 `run_workflow.py --phase review`。
- 要拆分版或盈亏表：确认用户已人工核对初版，再执行 `run_workflow.py --phase final`。

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
sys.argv = [
    "build_model_summary.py",
    "--caotan", "/workspace/草滩在线表.xlsx",
    "--gangwu", "/workspace/港务在线表.xlsx",
    "--output-dir", "/workspace/output",
    "--month", "2026-04",
]
runpy.run_path("{skill_dir}/scripts/build_model_summary.py", run_name="__main__")
```

该模式不建立月度运行状态，不能直接进入最终阶段。

## 阶段一：预检并生成初版

确认完整动态输入已上传后运行：

```python
import runpy, sys
sys.argv = [
    "run_workflow.py",
    "--phase", "review",
    "--workspace", "/workspace",
    "--month", "2026-04",
    "--bundle-root", "{skill_dir}",
]
runpy.run_path("{skill_dir}/scripts/run_workflow.py", run_name="__main__")
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
sys.argv = [
    "run_workflow.py",
    "--phase", "final",
    "--workspace", "/workspace",
    "--review-file", "/workspace/4月_02_基础报表初版.xlsx",
    "--bundle-root", "{skill_dir}",
]
runpy.run_path("{skill_dir}/scripts/run_workflow.py", run_name="__main__")
```

如用户上传了上月一干/二干盈亏表，分别增加 `--previous-first-profit` 和 `--previous-second-profit`。

当 `phase` 为 `final_completed` 时，交付拆分版、一干盈亏表和二干盈亏表，并汇报校验结果。

官方基础报表母版中的“二干基础”“一干汇总”“两趟合一趟”等特殊 Sheet 属于人工
业务汇总，当前 Skill 只自动生成普通线路 Sheet、明确确认后的“降吨位”审计 Sheet
和盈亏表。若终端 `warnings` 列出这些特殊 Sheet 待人工处理，必须明确告知用户，
不能把普通线路自动生成结果宣称为官方完整等价文件。

## 禁止事项

- 禁止用在线表替代车管数据作为基础报表行来源。
- 禁止在没有人工修订初版和同次 `run-state.json` 时执行最终阶段。
- 禁止猜测缺失容间、吨位、价格、路线或收入；保持空值和问题标记。
- 禁止在 Excel 打开时覆盖文件。
