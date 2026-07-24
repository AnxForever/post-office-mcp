#!/usr/bin/env python3
"""
邮局报表核对 MCP Server — 4 tools, wraps proven V5.3 scripts.

stdio mode:  python3 server.py
SSE mode:    python3 server.py --sse  (binds 0.0.0.0:$PORT for Render)

AgentOS 添加方式:
  传输方式: 命令
  命令: python
  参数: /path/to/server.py
  环境变量: {"SCRIPTS_DIR": "/path/to/scripts"}
"""
import os, sys, json, subprocess, tempfile, shutil

# ============================================================
# Config
# ============================================================
SCRIPTS_DIR = os.environ.get("SCRIPTS_DIR", os.path.join(os.path.dirname(__file__), "..", "aionui-assistant-package", "resources", "scripts"))
PYTHON = sys.executable or "python"

def _run_script(script_name, args):
    """Run a build script and return result dict."""
    script_path = os.path.join(SCRIPTS_DIR, script_name)
    if not os.path.exists(script_path):
        return {"success": False, "error": "SCRIPT_NOT_FOUND", "hint": f"脚本不存在: {script_path}"}

    cmd = [PYTHON, script_path] + [str(a) for a in args]
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=600, encoding="utf-8", errors="replace")
        stdout = result.stdout or ""
        stderr = result.stderr or ""
        if result.returncode != 0:
            return {"success": False, "error": "SCRIPT_FAILED", "hint": stderr[-500:] or stdout[-500:]}

        # Parse key metrics from stdout
        metrics = {}
        for line in stdout.splitlines():
            line = line.strip()
            if line.startswith("[") and "]" in line:
                metrics["raw_output"] = metrics.get("raw_output", "") + line + "\n"
            for key, pattern in [
                ("total", "总="), ("red", "红"), ("orange", "橙"), ("yellow", "黄"), ("blue", "蓝"),
                ("km_pct", "公里="), ("price_pct", "标准价="), ("issues", "问题汇总"),
                ("caotan", "草滩="), ("gangwu", "港务="), ("merged", "合并="),
                ("filled", "补容积"),
            ]:
                if pattern in line:
                    metrics[key] = line

        return {"success": True, "output": stdout[-2000:], "metrics": metrics}
    except subprocess.TimeoutExpired:
        return {"success": False, "error": "TIMEOUT", "hint": "脚本执行超时(600秒)"}
    except Exception as e:
        return {"success": False, "error": "INTERNAL_ERROR", "hint": str(e)[:200]}


# ============================================================
# MCP Server
# ============================================================
try:
    from mcp.server.fastmcp import FastMCP
    from mcp.server.transport_security import TransportSecuritySettings

    mcp = FastMCP(
        "post-office-report-check",
        transport_security=TransportSecuritySettings(
            enable_dns_rebinding_protection=True,
            allowed_hosts=["post-office-mcp.onrender.com", "localhost", "127.0.0.1"],
        ),
    )

    @mcp.tool()
    def build_summary(
        month: int,
        caotan_path: str,
        gangwu_path: str,
        output_dir: str,
    ) -> dict:
        """
        生成车型汇总表(交付产物1)。合并草滩+港务在线表当月Sheet,跨场地补容间。

        参数:
            month: 月份(如4)
            caotan_path: 草滩在线表路径
            gangwu_path: 港务在线表路径
            output_dir: 输出目录

        返回: {success, output_path, metrics}
        输出文件: {output_dir}/{month}月_01_车型汇总表.xlsx
        """
        result = _run_script("build_summary.py", [month, output_dir, caotan_path, gangwu_path])
        if result.get("success"):
            result["output_path"] = os.path.join(output_dir, f"{month}月_01_车型汇总表.xlsx")
        return result

    @mcp.tool()
    def build_draft(
        month: int,
        caotan_path: str,
        gangwu_path: str,
        yigan_price_path: str,
        ergan_price_path: str,
        carrier_dir: str,
        output_dir: str,
    ) -> dict:
        """
        生成基础报表初版(交付产物2)。车管数据为行来源,在线表取容积,价格表算运价。
        输出21列固定结构,带红橙黄蓝四色标记和问题汇总Sheet。
        生成后必须停下等人工核对。

        参数:
            month: 月份(如4)
            caotan_path: 草滩在线表路径
            gangwu_path: 港务在线表路径
            yigan_price_path: 一干价格表路径
            ergan_price_path: 二干价格表路径
            carrier_dir: 车管数据目录(含6个公司xlsx)
            output_dir: 输出目录

        返回: {success, output_path, metrics}
        输出文件: {output_dir}/{month}月_02_基础报表初版.xlsx
        """
        result = _run_script("build_draft.py", [month, output_dir, caotan_path, gangwu_path, yigan_price_path, ergan_price_path, carrier_dir])
        if result.get("success"):
            result["output_path"] = os.path.join(output_dir, f"{month}月_02_基础报表初版.xlsx")
        return result

    @mcp.tool()
    def build_split(
        month: int,
        draft_path: str,
        template_path: str,
        output_dir: str,
    ) -> dict:
        """
        生成基础报表拆分版(交付产物3)。按模板逐Sheet拆分,容间取整。
        必须在初版人工核对后调用。

        参数:
            month: 月份(如4)
            draft_path: 核对后初版文件路径
            template_path: 拆分模板(基础报表原版.xlsx)路径
            output_dir: 输出目录

        返回: {success, output_path, metrics}
        输出文件: {output_dir}/{month}月_03_基础报表拆分版.xlsx
        """
        result = _run_script("build_split.py", [month, output_dir, draft_path, template_path])
        if result.get("success"):
            result["output_path"] = os.path.join(output_dir, f"{month}月_03_基础报表拆分版.xlsx")
        return result

    @mcp.tool()
    def build_pnl(
        month: int,
        draft_path: str,
        finance_path: str,
        pnl_template_dir: str,
        output_dir: str,
    ) -> dict:
        """
        生成一干/二干盈亏表(交付产物4+5)。照搬盈亏模板格式,含环比区。
        必须在初版人工核对后调用。

        参数:
            month: 月份(如4)
            draft_path: 核对后初版文件路径
            finance_path: 财务收入表路径
            pnl_template_dir: 盈亏模板目录(含一干委办盈亏.xlsx和二干委办盈亏.xlsx)
            output_dir: 输出目录

        返回: {success, output_paths, metrics}
        输出文件: {output_dir}/{month}月_04_一干盈亏表.xlsx + {month}月_04_二干盈亏表.xlsx
        """
        result = _run_script("build_pnl.py", [month, output_dir, draft_path, finance_path, pnl_template_dir])
        if result.get("success"):
            result["output_paths"] = [
                os.path.join(output_dir, f"{month}月_04_一干盈亏表.xlsx"),
                os.path.join(output_dir, f"{month}月_04_二干盈亏表.xlsx"),
            ]
        return result

    # ============================================================
    # Entry point
    # ============================================================
    if "--sse" in sys.argv:
        import uvicorn
        from starlette.middleware.trustedhost import TrustedHostMiddleware
        app = mcp.sse_app()
        app.add_middleware(TrustedHostMiddleware, allowed_hosts=["*"])
        port = int(os.environ.get("PORT", 8000))
        print(f"MCP Server starting on 0.0.0.0:{port} (SSE mode)")
        uvicorn.run(app, host="0.0.0.0", port=port)
    elif "--http" in sys.argv:
        import uvicorn
        port = int(os.environ.get("PORT", 8000))
        app = mcp.streamable_http_app()
        print(f"MCP Server starting on 0.0.0.0:{port} (streamable HTTP)")
        uvicorn.run(app, host="0.0.0.0", port=port)
    else:
        print(f"MCP Server starting (stdio mode), scripts dir: {SCRIPTS_DIR}", file=sys.stderr)
        mcp.run()

except ImportError:
    print("pip install fastmcp uvicorn", file=sys.stderr)
    sys.exit(1)
