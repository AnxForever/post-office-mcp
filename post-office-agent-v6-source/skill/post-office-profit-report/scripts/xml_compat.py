# -*- coding: utf-8 -*-
"""Windows 上 openpyxl 读取工作簿的稳定性补丁。

问题现象
--------
读取较大的工作簿（如 100 个 Sheet 的基础报表原版）时，进程偶发以
0xC0000005（access violation / 段错误）退出。崩溃点每次不同，都落在 XML 解析里：

    openpyxl/worksheet/_read_only.py  _cells_by_row
    openpyxl/worksheet/_reader.py     parse
    xml/etree/ElementTree.py          _fixname / _start / feed

与业务代码无关：直接 `load_workbook(...); for ws in wb.worksheets: list(ws.iter_rows())`
就能复现。

根因
----
openpyxl 在 import 时探测 XML 后端：lxml 可用就用 lxml，否则用标准库
`xml.etree.ElementTree`。两种后端都受这个 Windows + CPython 3.12 的原生问题影响，
但崩溃率差别很大。实测（同一台机器，`--phase detect` 读 3 月全套输入，各跑 20 次）：

    lxml 后端       崩溃 1 / 20
    标准库后端      崩溃 8 / 20

更早的一次测量（同样的机器，40 次）结论相反，说明这不是代码问题，而是
「哪个后端这次更稳」会随 lxml / CPython 补丁版本变化。所以这里不再写死后端，
而是默认用 lxml（可用时），并允许用环境变量切回来：

    POST_OFFICE_STDLIB_XML=1   强制标准库后端
    POST_OFFICE_LXML=0         同上（旧名字）

lxml 不可用时自动退回标准库。

用法
----
每个会读取工作簿的入口脚本，在 `from __future__` 之后、任何 openpyxl 导入之前：

    from reconciliation import xml_compat  # noqa: F401  必须在 openpyxl 之前导入

幂等，重复导入无副作用。
"""
from __future__ import annotations

import os


def _force_stdlib() -> bool:
    if os.environ.get("POST_OFFICE_STDLIB_XML") == "1":
        return True
    return os.environ.get("POST_OFFICE_LXML") == "0"


def _lxml_available() -> bool:
    try:
        import lxml.etree  # noqa: F401
    except Exception:
        return False
    return True


# openpyxl 只在环境变量缺失或为真值时才会去探测 lxml；
# 显式写 "False" 可以跳过 lxml，直接使用标准库。
if _force_stdlib() or not _lxml_available():
    os.environ["OPENPYXL_LXML"] = "False"
else:
    os.environ.pop("OPENPYXL_LXML", None)
