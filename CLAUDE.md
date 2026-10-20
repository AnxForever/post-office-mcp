# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 项目概述

邮局月度委办运输报表处理系统。每月用6家物流公司派车记录 + 两个场地容积测量 + 价格合同，生成5个Excel交付件。

## 核心脚本

| 脚本 | 位置 | 功能 |
|------|------|------|
| build_summary.py | aionui-assistant-package/resources/scripts/ | 合并草滩+港务在线表 → 车型汇总表 |
| build_draft.py | 同上 | 车管数据为行来源 → 21列初版(带标记) |
| build_split.py | 同上 | 初版 → 拆分版(~100 Sheet) |
| build_pnl.py | 同上 | 初版+财务表 → 一干/二干盈亏表 |

## 关键业务规则

- **行来源必须用车管数据**(6公司派车记录)，不是在线表或车型表
- **初版固定21列**，列数只数非空表头，不信 `max_column`
- **容积取数**: (日期+车牌)精确 → 同车牌跨站兜底(串行) → 都没有标橙
- **草滩容积是Excel公式** `=长×宽×高`，data_only读不到时硬算
- **价格表吨位档/系数/路线全动态读取**，不写死
- **所有不确定必标注**，初版生成后立即停下等人核对
- **`matched matching.py` 修复**: match_vehicle_row 加 plate fallback，标记 `_cross_station`
- **`templates.py` 修复**: 加蓝标 BBD7EE，layout_max_column 截断到 21

## Agent 包

- `邮局月度委办运输报表核对V5.3.zip` — AgentOS 导入包
- `post-office-monthly-reports-skill.zip` — AgentOS 后台 skill 包
- 官方原版是 reconcile-monthly-transport-reports.zip，V5.3 在此基础上修了三处

## 数据文件

- 在线表: `数据/草滩在线表.xlsx`, `数据/港务在线表.xlsx`
- 车管数据: `数据/每月车管基础数据下载/N月/`
- 官方模板+财务: `数据/委办报表开发/N月/`

## 运行测试

```bash
python aionui-assistant-package/resources/scripts/build_summary.py 4 ./输出 草滩在线表.xlsx 港务在线表.xlsx
python aionui-assistant-package/resources/scripts/build_draft.py 4 ./输出 草滩在线表.xlsx 港务在线表.xlsx 一干价格表.xlsx 二干运费价.xlsx 车管目录/
python aionui-assistant-package/resources/scripts/build_split.py 4 ./输出 核对后初版.xlsx 基础报表原版.xlsx
python aionui-assistant-package/resources/scripts/build_pnl.py 4 ./输出 核对后初版.xlsx 财务收入表.xlsx 盈亏模板目录/
```

## 常见陷阱

- 草滩在线表 Sheet 名 `4月1日`，港务 `4.1`，正则不同
- 草滩容积是公式列，`data_only=True` 返回 None，必须 `长×宽×高` 计算
- 同车牌跨站补容间只覆盖 607 条(4月)，仍有 388 条真没测到(标橙)
- 价格表第二行是系数(如0.63)，第四行才是容积下限(≥55)，不要搞混
- emoji 在 Windows 控制台会崩，全部用 [OK]/[STOP] 替代
