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
3. **运行任何业务脚本前，先执行 `--phase detect`**，确认每个文件被识别成什么角色。
   识别依据是表结构，不是文件名；返回 `status:"needs_input"` 时停下来问用户，
   不要自己猜文件名。上传目录随平台版本变化（旧版 `/workspace`，新版 `/inputs`），
   脚本会自动解析并在 `--phase list` 的 `input_dir` 中返回真实位置。
4. 每个阶段只执行一次对应脚本，只依据末尾 JSON 汇报结果。
5. 保留同一任务的 `/workspace/output/run-state.json`；第二阶段不得换工作区。
6. `status:"needs_input"`（`ok:true`）表示输入不全，不是执行失败：把
   `missing_inputs[].action` 逐条告诉用户，请其补充后重试；已识别文件会保留。
   只有 `ok:false` 才是真正的执行失败，此时原样汇报 `errors` 和 `next_action`。
7. Windows 执行时设置 `PYTHONIOENCODING=utf-8` 和 `PYTHONUTF8=1`。

## 第 0 步：识别文件角色

任何业务脚本之前，先跑一次预检：

```python
import runpy, sys
sys.path.insert(0, "/workspace/skills/post-office-monthly-reconciliation/scripts")
sys.argv = ["agentos_entrypoint.py", "--phase", "detect", "--task", "review",
                        "--workspace", "/workspace"]  # 不传 --month，脚本自动推断
runpy.run_path("/workspace/skills/post-office-monthly-reconciliation/scripts/agentos_entrypoint.py", run_name="__main__")
```

`--task` 取 `review`（完整月报）、`summary`（只要车型汇总表）或 `profit`（只要盈亏表）。

**不用传 `--month`**：脚本会从文件名和工作表名推断目标月，结果在
`target_month`（最终采用）和 `inferred_month`（推断来源）里。

- 文件名里的 `9月`、`2026.09`，工作表名里的 `9月1日`、`9.1` 都能识别。
- 财务收入表的月份通常滞后（业务上允许传上月的），所以它在推断里权重更低，
  不会被当成目标月。
- 推断不出、或两个月份打平时返回 `status:"needs_month"`，此时才需要问用户。
  **不要为了省事自己猜一个月份。**

脚本按**表结构**判断每个文件是什么，不看文件名：

| 识别结果 | 判断依据 |
|---|---|
| 基础报表原版 | 含 Sheet `一干汇总` 或 `二干基础` |
| 财务收入表 | 表头含 `邮路名称` + `结算收入` + `普服收入`（客户可能叫「运输局表」） |
| 草滩在线表 | 日 Sheet 命名为 `8月1日` 格式 |
| 港务在线表 | 日 Sheet 命名为 `8.1` 格式 |
| **车辆容积测量表**（合并） | 无日 Sheet，且有一列**没有表头**、值形如 `9.1` |
| 车管基础数据 | 表头含 `邮路名称` + `邮路执行日期` + `派车单号` + `邮运量` |
| **车管平台汇总导出**（合并） | 表头含 `邮路名称` + `邮路执行日期` + `派车单号` + `车牌号`，但**没有** `邮运量`/结算吨位 |
| 一干/二干价格表 | 表头含 `路线/路向` + `公里` + 吨位/系数，且表内标题或文件名标注一干/二干 |
| 一干/二干盈亏模板 | 含 7 列盈亏子表头，且表内标题或文件名标注一干/二干 |
| 基础报表初版模板 | 21 列表头完全匹配 |

### 两套输入口径，二选一，不能混用

| 口径 | 需要的文件 |
|---|---|
| 逐场地/逐公司（`split`） | 草滩在线表 + 港务在线表 + 6 家公司各自的派车记录 |
| 合并导出（`merged`） | 1 张车辆容积测量表（草滩+港务一次导出）+ 1 张车管平台汇总导出 |

平台现在默认给「合并导出」，只需 2 个文件。`--phase detect` 返回的 `source_bundle`
会写明这次走的哪一套。同时上传两套会按 `split` 处理并给出警告。

合并口径下**容间以车管平台那一列为准**，不是测量表的实测值——与官方报表逐单核对，
平台值一致 3210/3256，实测值只对得上一半。平台值缺失、或与实测相差超过 5 方时
才回退实测值，并在该行标橙提示核对。

### 模板和价格表不用每月上传

初版模板、一干/二干盈亏模板、一干/二干价格表都随 Skill 内置在 `assets/` 下
（`first-price-table.xlsx` 是一干 2026 版，`second-price-table.xlsx` 是二干 2026 合同价）。
用户只传当月的在线表、车管和财务即可；**上传了就用用户的**（客户换版式或调价时以用户的为准），
没上传就用内置版本。`--phase detect` 返回的 `bundled_templates` / `bundled_price_tables`
会列出这次用了哪些内置文件，汇报时可以说明。

`assets/route-supplement.json` 是**合同外专线的补充价目**（兰州、平凉、陇南、洛川港、
兴平港、泾阳港、庆阳港）。这几条不在两张合同价目表里，但每月都要结算；价格与
「容间 → 吨位档」区间取自官方基础报表逐单反查（`source_month` 记录来源月份）。
官方这些专线是按固定容间区间定档的（兰州 63~91 方走 12 吨、127 方走 20 吨、
130~143 方走 30 吨、153 方以上走 40A），**不是**按容间除以 5 估算。
只补价格表里完全不存在、且路线名不是主表别名指向的路线；合同调价时需同步改这个文件。

一干/二干成套使用：用户只传了其中一份（或传了一份分不出干线的候选），
**不会**用内置的另一份补齐——两版合同混用会让运价错得毫无痕迹。此时照常返回缺文件，
请用户把一干、二干一起传上来。

内置价格表是合同（很少变动）。一旦客户调价，让用户上传新版覆盖，并同步替换
`assets/` 里的文件重新打包，否则内置版本就过期了。

返回字段的用法：

- `resolved`：每个角色唯一命中的文件路径，可直接作为脚本参数。
- `files[].evidence`：为什么判成这个角色，向用户解释时引用它。
- `confidence`：`low` 表示月份对不上或站点无法区分，汇报时要说明。
- `missing_inputs` / `status:"needs_input"`：缺文件或多候选，**停下问用户**，不要自行挑选。
  这是正常的「等用户补文件」，退出码仍是 0，不要汇报成执行失败。

需要看目录里还有哪些文件时，再执行 `--phase list`：

```python
sys.argv = ["agentos_entrypoint.py", "--phase", "list", "--workspace", "/workspace"]
```

### 识别不出来的情况

- **一干/二干价格表、盈亏模板**：两者结构相同，可靠信号只有表内标题或文件名里的「一干/二干」。
  都没有时脚本返回 `price_candidate` / `profit_candidate` 并标 `low`，此时必须问用户，
  不能按档位名猜（一干 2025 版模板也有 8 吨档，二干 2026 版价格表也有 40B 档）。
- **草滩/港务在线表**：只按日 Sheet 命名区分。若两份表用同一种命名，脚本返回 `low`，
  汇报时要求用户确认哪个是草滩、哪个是港务。

## 仅生成车型汇总表

运行：

```python
import runpy, sys
sys.path.insert(0, "/workspace/skills/post-office-monthly-reconciliation/scripts")
sys.argv = [
    "agentos_entrypoint.py", "--phase", "summary",
    "--workspace", "/workspace",
    # 不传 --month，脚本自动推断
]
runpy.run_path("/workspace/skills/post-office-monthly-reconciliation/scripts/agentos_entrypoint.py", run_name="__main__")
```

未传 `--caotan` / `--gangwu` 时，脚本按日 Sheet 命名自动识别；识别不唯一会返回
`status:"needs_input"` 并列出候选，请用户确认或补齐文件。
只有 `--phase detect --task summary` 已经给出两个明确路径时，才用 `--caotan/--gangwu` 显式传入。

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
    # 不传 --month，脚本自动推断
    "--bundle-root", "/workspace/skills/post-office-monthly-reconciliation",
]
runpy.run_path("/workspace/skills/post-office-monthly-reconciliation/scripts/run_workflow.py", run_name="__main__")
```

当 `phase` 为 `awaiting_review` 时：

- 向用户提供车型汇总表和基础报表初版（`published` 里是平台「成果文件」中的文件名）。
- 汇报 `metrics`、`issues` 和 `warnings`。
- 立即停止，请用户人工核对初版中的颜色行与“问题汇总”。
- 不自动继续生成拆分版或盈亏表。

初版行底色含义（同一行有多种问题时取最深的）：

| 颜色 | 含义 |
|---|---|
| 红 | 硬缺失，相关数值算不出来：未匹配车型汇总、价格表未覆盖该路线、容间低于最低吨位档、往返价缺失 |
| 深橙 | 降吨位后对应吨位价格为空 |
| 浅蓝 `#BBD7EE` | 容间取自**同车牌其他日期**的在线测量记录，不是本趟实测，需复核车型是否一致 |
| 橙 | 在线记录确认未测量，容间为 0 |
| 黄 | 吨位偏离该车本月基线，疑似跨档 |

浅蓝行同样会在「问题汇总」里出现「容间取自同车牌其他日期的在线测量记录」条目，汇报时要说明这批容间是兜底推算的。

## 阶段二：生成最终报表

只在用户明确确认已核对并上传修订初版后运行：

```python
import runpy, sys
sys.path.insert(0, "/workspace/skills/post-office-monthly-reconciliation/scripts")
sys.argv = [
    "run_workflow.py",
    "--phase", "final",
    "--workspace", "/workspace",
    "--review-file", "<--phase detect 返回的初版真实路径>",
    "--bundle-root", "/workspace/skills/post-office-monthly-reconciliation",
]
runpy.run_path("/workspace/skills/post-office-monthly-reconciliation/scripts/run_workflow.py", run_name="__main__")
```

如用户上传了上月一干/二干盈亏表，分别增加 `--previous-first-profit` 和 `--previous-second-profit`。**未上传则环比区留空，严禁从模板或其他来源推测上月数据。**

当 `phase` 为 `final_completed` 时，交付拆分版、一干盈亏表和二干盈亏表，并汇报校验结果。
脚本会把这三个文件复制到平台成果目录 `/outputs`，在 `published` 中返回；平台「成果文件」
是用户实际下载交付件的位置，不要让用户去 `/workspace/output` 找。

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
    "agentos_entrypoint.py", "--phase", "profit",
    "--workspace", "/workspace",
    "--month", "2026-08",
]
runpy.run_path("/workspace/skills/post-office-monthly-reconciliation/scripts/agentos_entrypoint.py", run_name="__main__")
```

未传 `--base-file` / `--finance-file` 时，脚本按表结构自动识别：含 `一干汇总`/`二干基础`
的是基础报表原版，表头含 `结算收入` + `普服收入` 的是财务收入表。识别不唯一时会返回
`status:"needs_input"` 并列出 `missing_inputs`、`detected`、`missing`、`ambiguous`，
此时请用户确认或补充文件，**不要让模型自己拼文件名**。

> 财务收入表就是「运输局表」。客户有时把它命名为 `2026.08运输局.xlsx`，
> 有时叫 `2026.03财务收入表.xlsx`，两者是同一张表（31 列、Sheet 名 `sheet1`），按内容识别即可。

该脚本的口径（详见 [reconciliation-rules.md](references/reconciliation-rules.md) 的「盈亏表」一节）：

- 趟次 / 装载量 / 运费来自“一干汇总”“二干基础”，**行数与这两个 sheet 完全守恒**。
- **吨位归类：先看备注，无备注取该车该路线的最低档。**
  备注写「调度通知按30吨结算」「两辆12吨车合并为一辆30吨车」的，该行整行（趟次、装载量、运费、收入）
  归类为 30 吨档；没有备注时按该车在该路线跑过的最低吨位档归类。
  基础报表的 `吨位` 列是车辆车型档，不等于结算档。
- 结算收入来自财务表 `自委办=委办` 且**排除 `往返=返程`**的行，逐趟按（车牌 + 流水号前 4 位 MMDD）一对一分摊。
- **跨月记录按「是否出现在本月财务表」判定，不看流水号日期。** 只要该行在当月财务收入表里就属于当月；
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
