---
name: post-office-model-summary
description: 邮局车型汇总表专用技能。只上传草滩在线测量表和港务在线测量表，按日期+车牌合并成当月的车型汇总表（01）。用户提到车型汇总表、容积汇总、01 表时使用。
---

# 邮局车型汇总表

只运行本技能里的 `agentos_entrypoint.py`，不要自己写临时脚本、不要自己读写 Excel。

## 输入

只要两张表：

1. **草滩在线测量表** —— 日 Sheet 名类似 `9月1日`
2. **港务在线测量表** —— 日 Sheet 名类似 `9.1`

按表结构识别，不看文件名。缺一张或同一张识别到多份时，返回 `needs_input`，把
`missing_inputs[].action` 逐条告诉用户，不要自己挑文件。

## 一条命令做完

```python
import runpy, sys
sys.path.insert(0, "{skill_dir}/scripts")
sys.argv = ["agentos_entrypoint.py", "--workspace", "/workspace"]
runpy.run_path("{skill_dir}/scripts/agentos_entrypoint.py", run_name="__main__")
```

- 不传 `--month`：脚本从文件名和日 Sheet 名推断月份。
- 返回 `status:"summary_completed"` 就是做完了，**立刻汇报，不要重跑**。
- 返回 `status:"needs_month"` / `"needs_confirm"`：把问题原样转告用户，拿到月份后带
  `--month 2026-09` 重跑一次。
- 返回 `status:"needs_input"`：不是报错，把缺的文件告诉用户。

## 输出

- `N月_01_车型汇总表.xlsx`
  - Sheet「车型汇总」：日期、场地、邮路、车牌、挂车、长宽高、测量容积、车型问题、备注
  - Sheet「容积取证」：每条记录取的是源值还是长×宽×高重算，方便追查

同一个日期+车牌只保留一条：优先非零测量值，同场地按 Sheet/行顺序取最后一条。
车型汇总表不设颜色，问题写在「车型问题/备注」列。

## 口径

- 草滩容积是公式列，`data_only` 读不到时按长×宽×高重算，并在「容积取证」写明。
- 在线表里有、但车管没有对应派车的记录，不在本技能处理；留给「基础报表初版」技能标出来。
- 本技能不读车管数据、价格表、财务收入表。

## 停下来的情况

| status | 含义 | 怎么跟用户说 |
|---|---|---|
| `needs_month` | 推断不出月份 | 问“这是哪一年的哪个月？” |
| `needs_confirm` | 月份是推断的，证据不够强 | 问“我按 2026 年 9 月处理，对吗？” |
| `needs_input` | 缺草滩/港务表，或同一角色多份 | 逐条转述 `missing_inputs[].action` |
| `ok:false` | 文件读不了、月份参数不合法 | 原样转述 `errors` |
