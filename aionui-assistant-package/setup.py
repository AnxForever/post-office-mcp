#!/usr/bin/env python3
"""
AionUi 邮局报表助手 · 一键配置脚本

用法：
  1. 先启动 AionUi 应用（确保 aioncore 后端在运行）
  2. 在本目录运行: python3 setup.py

脚本会自动：
  - 扫描本地端口找到 aioncore
  - 导入三个技能 (post-office-build-draft/split/pnl)
  - 导入助手 (邮局委办报表助手)
  - 写入助手系统提示词
  - 验证配置成功

完成后刷新 AionUi 的助手列表即可看到「邮局委办报表助手」。
"""
import json
import os
import socket
import sys
import time
import urllib.request
import urllib.error
from pathlib import Path

# === 配置 ===
PACKAGE_DIR = Path(__file__).parent.resolve()
ASSISTANTS_JSON = PACKAGE_DIR / 'assistants-import.json'
ASSISTANT_RULE = PACKAGE_DIR / 'assistant-rule.md'
SKILLS_DIR = PACKAGE_DIR / 'skills'
RESOURCES_DIR = PACKAGE_DIR / 'resources'

ASSISTANT_ID = 'post-office-report-assistant'

# AionUi aioncore 常见端口范围（动态分配，扫描这些）
PORT_RANGE = list(range(31000, 32000)) + list(range(21300, 21500)) + list(range(8000, 8100)) + list(range(9000, 9100))


def log(msg, kind='info'):
    prefix = {'info': '•', 'ok': '✅', 'warn': '⚠', 'err': '❌', 'step': '▶'}.get(kind, '•')
    print(f"{prefix} {msg}")


def http_request(method, url, data=None, timeout=5):
    """简单 HTTP 请求，返回 (status, body_dict_or_text)"""
    try:
        headers = {'Content-Type': 'application/json'}
        body = json.dumps(data).encode() if data is not None else None
        req = urllib.request.Request(url, data=body, headers=headers, method=method)
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            raw = resp.read().decode('utf-8', errors='replace')
            try:
                return resp.status, json.loads(raw)
            except json.JSONDecodeError:
                return resp.status, raw
    except urllib.error.HTTPError as e:
        raw = e.read().decode('utf-8', errors='replace')
        try:
            return e.code, json.loads(raw)
        except json.JSONDecodeError:
            return e.code, raw
    except (urllib.error.URLError, socket.timeout, ConnectionError):
        return None, None


def find_aioncore_port():
    """扫描端口找 aioncore（响应 /api/auth/status 的）"""
    log('扫描本地端口寻找 aioncore...', 'step')
    # 先试最常见的几个
    quick_ports = [21337, 21338, 21339, 8080, 8000, 9000, 3000]
    for port in quick_ports + PORT_RANGE:
        status, body = http_request('GET', f'http://127.0.0.1:{port}/api/auth/status', timeout=0.3)
        if status == 200 and isinstance(body, dict):
            # 确认是 aioncore（响应里有 needs_setup 或 is_authenticated 字段）
            if 'needs_setup' in body or 'is_authenticated' in body or 'user_count' in body:
                log(f'找到 aioncore: http://127.0.0.1:{port}', 'ok')
                return port, body
    return None, None


def check_auth(base, auth_status):
    """检查是否需要登录"""
    if auth_status.get('needs_setup'):
        log('AionUi 首次使用，需要先在应用里设置管理员账号。', 'err')
        log('请先打开 AionUi 应用完成初始化设置，再跑本脚本。', 'err')
        return False
    if auth_status.get('is_authenticated'):
        log('已认证（本地会话）', 'ok')
        return True
    # 本地访问通常 session 自动信任，试着调一个需要 auth 的 API
    status, _ = http_request('GET', f'{base}/api/assistants', timeout=5)
    if status == 200:
        log('本地访问认证通过', 'ok')
        return True
    if status in (401, 403):
        log('需要登录认证。请在 AionUi 应用里登录后，从浏览器开发者工具复制 cookie 给本脚本。', 'err')
        log('或者：在 AionUi 应用里直接用「导入助手」功能（设置 → 助手 → 导入），选 assistants-import.json。', 'warn')
        return False
    return True


def import_skills(base):
    """导入三个技能"""
    log('导入技能...', 'step')
    skills = ['post-office-build-draft', 'post-office-build-split', 'post-office-build-pnl']
    for skill_name in skills:
        skill_md = SKILLS_DIR / skill_name / 'SKILL.md'
        if not skill_md.exists():
            log(f'技能文件不存在: {skill_md}', 'err')
            return False
        # 调 /api/skills/import，参数是 skill_path（SKILL.md 的绝对路径）
        status, body = http_request('POST', f'{base}/api/skills/import', data={
            'skill_path': str(skill_md)
        }, timeout=15)
        if status == 200:
            log(f'技能已导入: {skill_name}', 'ok')
        elif status == 409:
            log(f'技能已存在，跳过: {skill_name}', 'ok')
        else:
            log(f'导入技能 {skill_name} 失败: HTTP {status} {body}', 'err')
            return False
    return True


def import_assistant(base):
    """导入助手"""
    log('导入助手...', 'step')
    if not ASSISTANTS_JSON.exists():
        log(f'助手导入文件不存在: {ASSISTANTS_JSON}', 'err')
        return False
    data = json.loads(ASSISTANTS_JSON.read_text(encoding='utf-8'))
    status, body = http_request('POST', f'{base}/api/assistants/import', data=data, timeout=15)
    if status == 200:
        result = body if isinstance(body, dict) else {}
        log(f'助手已导入: imported={result.get("imported", "?")} skipped={result.get("skipped", 0)} failed={result.get("failed", 0)}', 'ok')
        if result.get('failed', 0) > 0:
            log(f'失败详情: {result.get("errors", [])}', 'warn')
            return False
        return True
    else:
        log(f'导入助手失败: HTTP {status} {body}', 'err')
        return False


def write_rule(base):
    """写入助手系统提示词"""
    log('写入助手系统提示词...', 'step')
    if not ASSISTANT_RULE.exists():
        log(f'规则文件不存在: {ASSISTANT_RULE}', 'err')
        return False
    content = ASSISTANT_RULE.read_text(encoding='utf-8')
    # 替换资源目录占位符为实际路径
    content = content.replace('<RESOURCES_DIR>', str(RESOURCES_DIR).replace('\\', '/'))
    status, body = http_request('POST', f'{base}/api/skills/assistant-rule/write', data={
        'assistant_id': ASSISTANT_ID,
        'content': content,
        'locale': 'zh-CN',
    }, timeout=15)
    if status == 200:
        log('系统提示词已写入', 'ok')
        return True
    else:
        log(f'写入规则失败: HTTP {status} {body}', 'err')
        return False


def verify(base):
    """验证助手和技能都已就位"""
    log('验证配置...', 'step')
    status, body = http_request('GET', f'{base}/api/assistants', timeout=5)
    if status != 200:
        log(f'获取助手列表失败: HTTP {status}', 'err')
        return False
    assistants = body if isinstance(body, list) else []
    target = next((a for a in assistants if a.get('id') == ASSISTANT_ID), None)
    if not target:
        log(f'助手列表中未找到 {ASSISTANT_ID}', 'err')
        return False
    log(f'助手已就位: {target.get("name", ASSISTANT_ID)}', 'ok')
    skills = target.get('custom_skill_names', []) or target.get('enabled_skills', [])
    for expected in ['post-office-build-draft', 'post-office-build-split', 'post-office-build-pnl']:
        if expected not in skills:
            log(f'技能未绑定: {expected}（请手动在助手编辑页启用）', 'warn')
    return True


def main():
    print('=' * 60)
    print('AionUi 邮局委办报表助手 · 一键配置')
    print('=' * 60)

    # 0. 预检
    if not ASSISTANTS_JSON.exists() or not ASSISTANT_RULE.exists():
        log('缺少必要文件，请确保在 aionui-assistant-package 目录下运行', 'err')
        sys.exit(1)

    # 1. 找 aioncore
    port, auth_status = find_aioncore_port()
    if port is None:
        log('未找到运行中的 AionUi (aioncore)。请先启动 AionUi 应用，再跑本脚本。', 'err')
        sys.exit(2)
    base = f'http://127.0.0.1:{port}'

    # 2. 认证
    if not check_auth(base, auth_status or {}):
        sys.exit(3)

    # 3. 导入技能
    if not import_skills(base):
        log('技能导入失败，终止', 'err')
        sys.exit(4)

    # 4. 导入助手
    if not import_assistant(base):
        log('助手导入失败，终止', 'err')
        sys.exit(5)

    # 5. 写规则
    if not write_rule(base):
        log('规则写入失败（可手动在 UI 里粘贴）', 'warn')

    # 6. 验证
    if verify(base):
        print()
        print('=' * 60)
        log('配置完成！请刷新 AionUi 助手列表，选择「邮局委办报表助手」开始使用', 'ok')
        print('=' * 60)
        sys.exit(0)
    else:
        sys.exit(6)


if __name__ == '__main__':
    main()
