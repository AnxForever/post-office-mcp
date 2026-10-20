# 邮局报表 Agent + Skill 设计说明

---

## 一、现状诊断

### 官方 V5.0 的设计（正确的架构）

```
Agent 包 (agent.json)
├── systemPrompt: 告诉模型怎么调用
├── skills: ["post-office-monthly-reports"]
└── skillFiles: 25个文件内嵌
     └── post-office-monthly-reports/
         ├── manifest.json (type=bound, 唯一入口 run_workflow)
         ├── SKILL.md (使用说明)
         ├── scripts/
         │   ├── run_workflow.py (唯一入口, --phase review/final)
         │   ├── build_review_workbook.py (阶段一: 生成初版)
         │   ├── build_final_reports.py (阶段二: 拆分+盈亏)
         │   ├── inspect_inputs.py (输入预检)
         │   ├── validate_outputs.py (输出验证)
         │   └── reconciliation/ (业务逻辑模块)
         │       ├── matching.py (车管↔在线表匹配)
         │       ├── templates.py (Excel输出模板)
         │       ├── excel_io.py (Excel读写)
         │       ├── normalize.py (路线标准化)
         │       └── ...
         ├── assets/ (常驻资产)
         │   ├── company-map.json
         │   ├── route-aliases.json
         │   └── split-report-template.xlsx
         └── references/ (业务规则)
             ├── input-contract.md
             └── reconciliation-rules.md
```

**执行流程**:
```
用户上传14个文件 → 模型调 run_workflow --phase review → 生成初版
用户核对后上传修订初版 → 模型调 run_workflow --phase final → 生成拆分+盈亏
```

### V5.3 的问题清单

| # | 问题 | 根因 | 严重度 |
|---|------|------|:---:|
| 1 | **templates.py 缩进断裂** | `issue_number = 1` 丢了缩进,整个问题写入循环被弹出函数体外 | **致命** |
| 2 | **_sorted 死代码** | 之前尝试改问题分组,代码放到模块级,永远不会执行 | 中 |
| 3 | **原始问题循环丢失** | 替换时把 `for issue in issues:` 循环结构破坏 | **致命** |
| 4 | **39列未截断(部分)** | layout_max_column 有截断但模板结构继承问题 | 中 |
| 5 | 备注纯数字无解读 | 原版设计如此,非我们引入 | 低 |

---

## 二、正确的修复方案

### 原则: 在官方 V5.0 基础上做最小修改

不重写、不重构。只对三个文件做精准的外科手术。

### 修改清单

#### 修改1: `reconciliation/matching.py` — 串行跨站兜底

**位置**: `match_vehicle_row()` 函数

**原代码**:
```python
return index.get((execution_date, plate))
```

**改为**:
```python
result = index.get((execution_date, plate))
if result is not None:
    return result
# 串行兜底: 同车牌跨站取非零容积
plate_rows = [row for (d, p), row in index.items() if p == plate]
if plate_rows:
    nz = [r for r in plate_rows if isinstance(r.get("measured_volume"), (int, float)) and r["measured_volume"] > 0]
    sel = nz[0] if nz else plate_rows[0]
    sel["_cross_station"] = True
    return sel
return None
```

#### 修改2: `reconciliation/templates.py` — 蓝标 + 21列截断

**位置A**: `COLOR_PRIORITY` 和 `SEVERITY_FILL`

```python
# 加 blue
COLOR_PRIORITY = {None: 0, "yellow": 1, "orange": 2, "red": 3, "blue": 2.5}
SEVERITY_FILL = {
    ...
    "blue": PatternFill("solid", fgColor="FFBDD7EE"),
}
```

**位置B**: `layout_max_column`

```python
layout_max_column = min(max(layout_columns), 21)  # 截断到21列
```

**不动**: 问题写入循环保持原版 `for issue in issues:` — 不要碰缩进

#### 修改3: `build_review_workbook.py` — 串行标注

**位置**: `_model_issues()` 调用后

```python
if model is not None and model.get("_cross_station"):
    issues.append(Issue("cross_station", "blue",
        "串行：同车牌跨站点/跨日期取容积",
        "核对容积是否适用"))
```

---

## 三、不能做的事

| 不能做 | 原因 |
|--------|------|
| 改问题写入循环的缩进 | openpyxl 写操作对缩进极敏感,改了就崩 |
| 加 _sorted 分组排序 | 需要重构整个写入逻辑,风险极高 |
| 改 systemPrompt 的路径 | 模型找不到脚本就全崩 |
| 重命名 skill | 引发一连串路径问题 |
| 加 emoji | Windows GBK 控制台崩 |

---

## 四、两套产物

### 给 AgentOS 导入的 Agent 包
`邮局月度委办运输报表核对V5.3.zip`
- agent.json (含 systemPrompt + skillFiles)
- prompt.md
- 内嵌全部25个skill文件

### 给 AgentOS 后台装的 Skill 包
`post-office-monthly-reports-skill.zip`
- manifest.json
- SKILL.md
- scripts/ (全部Python)
- assets/ (3个常驻资产)
- references/ (2个业务规则)

**两套包的脚本内容必须一致。** 改了一个必须同步另一个。

---

## 五、构建流程

```
1. 从官方V5.0原版提取 skillFiles
2. 只改 matching.py / templates.py / build_review_workbook.py
3. 同步到 agent.json 的 skillFiles 和 skill zip
4. 打包
```

**关键**: 每次只改一个文件,改完立即验证语法,不批量替换。
