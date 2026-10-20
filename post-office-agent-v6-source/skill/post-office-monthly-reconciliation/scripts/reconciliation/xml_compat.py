# -*- coding: utf-8 -*-
"""Windows 上 openpyxl 读取工作簿的稳定性补丁。

问题现象
--------
读取较大的工作簿（如 100 个 Sheet 的基础报表原版）时，进程偶发以
0xC0000005（access violation / 段错误）退出。崩溃点每次不同，都落在 XML 解析里：

    lxml / ElementTree  fromstring / iterparse
    openpyxl/reader/excel.py  read / read_strings
    openpyxl/worksheet/_read_only.py  _cells_by_row

与业务代码无关：直接 `load_workbook(...); for ws in wb.worksheets: list(ws.iter_rows())`
就能复现。

根因
----
openpyxl 在 import 时会探测 lxml，存在就把它当作默认 XML 后端
（`openpyxl.LXML = True`）。本机装的 lxml 6.0.0 与 Python 3.12 组合下不稳定。

实测（同一组文件、同一台机器、每轮 40 次）：

    lxml 后端       崩溃 7 / 40
    标准库后端      崩溃 0 / 40

处理方式
--------
在 import openpyxl 之前设置环境变量 `OPENPYXL_LXML=False`，
强制 openpyxl 使用标准库 `xml.etree.ElementTree`。

用法
----
每个会读取工作簿的入口脚本，在 `from __future__` 之后、任何 openpyxl 导入之前：

    from reconciliation import xml_compat  # noqa: F401  必须在 openpyxl 之前导入

幂等，重复导入无副作用。
排查问题时可用 `POST_OFFICE_LXML=1` 恢复 lxml 后端做对比。
"""
from __future__ import annotations

import os

if os.environ.get("POST_OFFICE_LXML") == "1":
    os.environ.pop("OPENPYXL_LXML", None)
else:
    # openpyxl 只在环境变量缺失或为假值时才会去探测 lxml；
    # 显式写成 "False" 可以让它跳过 lxml，直接使用标准库。
    os.environ["OPENPYXL_LXML"] = "False"
