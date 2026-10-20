# 邮局月度委办运输报表核对 — 脚本包

## 文件清单

```
_agentos_package/
  scripts/
    build_summary.py    # ① 车型汇总表
    build_draft.py      # ② 基础报表初版（核心）
    build_split.py      # ③ 拆分版
    build_pnl.py        # ④ 一干/二干盈亏表
  company-map.json      # 公司简称→全称映射
  README.md             # 本文件
  workflow.md           # 业务流程说明
  notes.md              # 注意事项/常见陷阱
```

## 环境要求

- Python >= 3.10（正式安装版，**不要用Windows Store版**）
- openpyxl >= 3.1.0
- 无需其他第三方库

**Windows运行前必须设置编码:**
```bash
set PYTHONIOENCODING=utf-8
set PYTHONUTF8=1
```

**可选环境变量:**
```bash
set POST_OFFICE_REPORT_YEAR=2026    # 报表年份，默认2026
```

## 四步工作流

按顺序执行，不可跳过。

### ① 车型汇总表

合并草滩+港务在线表当月Sheet，跨场地补容间。

```bash
python scripts/build_summary.py <月份> <输出目录> <草滩在线表.xlsx> <港务在线表.xlsx>

# 示例
python scripts/build_summary.py 4 ./output 草滩在线表.xlsx 港务在线表.xlsx
```

输出: `{月}月_01_车型汇总表.xlsx`

### ② 基础报表初版

**核心步骤**。车管数据为行来源，在线表取容积，价格表算运价。

```bash
python scripts/build_draft.py <月份> <输出目录> <草滩表> <港务表> <一干价格表> <二干价格表> <车管目录>

# 示例
python scripts/build_draft.py 4 ./output \
  草滩在线表.xlsx 港务在线表.xlsx \
  一干价格表.xlsx 二干运费价.xlsx \
  车管数据目录/
```

输出: `{月}月_02_基础报表初版.xlsx`
  - Sheet: 结构发现 + 问题汇总 + 基础报表（21列）
  - **生成后必须停下，等人工核对修正**

### ③ 拆分版

从人工核对后的初版生成路线×公司拆分。

```bash
python scripts/build_split.py <月份> <输出目录> <核对后初版.xlsx> <拆分模板.xlsx>

# 示例
python scripts/build_split.py 4 ./output 核对后初版.xlsx 4月基础报表原版.xlsx
```

输出: `{月}月_03_基础报表拆分版.xlsx`（~88个Sheet）

### ④ 盈亏表

从初版+财务收入表生成一干/二干盈亏表。

```bash
python scripts/build_pnl.py <月份> <输出目录> <核对后初版.xlsx> <财务收入表.xlsx> <盈亏模板目录>

# 示例
python scripts/build_pnl.py 4 ./output 核对后初版.xlsx 财务收入表.xlsx 模板目录/
```

输出: `{月}月_04_一干盈亏表.xlsx` + `{月}月_04_二干盈亏表.xlsx`

## 路径约定（重要）

AgentOS 里**没有** `/inputs` 和 `/outputs` 这两个目录。用错会直接报
`PermissionError` 或 `FileNotFoundError`。

| 用途 | 正确路径 |
|---|---|
| 用户上传的源文件 | `/workspace/<文件名>` |
| 脚本输出 | `/workspace/output` |

调用脚本前的三条硬规则：

1. **先用真实文件名**。不确定时先列目录，不要臆造文件名：
   ```python
   import os
   print(os.listdir('/workspace'))
   ```
2. **输出目录传 `/workspace/output`**，不要传 `/outputs`。
3. 脚本自带容错：输出目录不可写会自动回退到 `./output`；输入文件不存在会打印
   同级目录清单和相近文件名后退出。**看到 `[STOP]` 就按提示改正路径重跑，不要重复提交同样的路径。**

## 输入文件要求

| 文件 | 说明 | 示例 |
|------|------|------|
| 草滩在线表 | Sheet名格式 `M月D日` 或 `M月DD日` | `4月1日`...`4月30日` |
| 港务在线表 | Sheet名格式 `M.D` | `4.1`...`4.30` |
| 一干价格表 | 含路线/公里/吨位档/系数动态发现 | `2024年10月新一干价格表.xlsx` |
| 二干价格表 | 同上，用于判定二干路线 | `二干运费价.xlsx` |
| 车管数据 | 目录含6-7个公司xlsx | `则一.xlsx` `富海源.xlsx` 等 |
| 财务收入表 | 含结算收入/普服收入列 | `2026.04财务收入表.xlsx` |
| 拆分模板 | 官方基础报表原版，含~100 Sheet | `4月基础报表原版.xlsx` |
| 盈亏模板 | 目录含一干委办盈亏.xlsx和二干委办盈亏.xlsx | 官方模板目录 |
