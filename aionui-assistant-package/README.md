# 邮局委办报表助手 · AionUi 一键配置包

## 🚀 启动直接用（推荐）

**三步：**

```bash
# 1. 启动 AionUi 应用（双击桌面图标）
# 2. 首次启动完成账号设置（在应用里设管理员密码）
# 3. 双击 setup.bat
```

`setup.bat` 会自动：
- 扫描本地端口找到 aioncore 后端
- 导入三个技能（初版/拆分/盈亏）
- 导入「邮局委办报表助手」
- 写入系统提示词（已自动填好资源目录路径）
- 验证配置

完成后**刷新 AionUi 助手列表**，就能看到「邮局委办报表助手」，直接对话使用。

> macOS/Linux 用户：`python3 setup.py`

## 包内容

```
aionui-assistant-package/
├── setup.py                     一键配置脚本（核心）
├── setup.bat                    Windows 双击启动
├── assistants-import.json       助手导入数据
├── assistant-rule.md            助手系统提示词
├── skills/
│   ├── post-office-build-draft/SKILL.md
│   ├── post-office-build-split/SKILL.md
│   └── post-office-build-pnl/SKILL.md
├── resources/
│   ├── scripts/                 Python 处理脚本
│   │   ├── build_draft.py / build_split.py / build_pnl.py
│   │   ├── inspect_sources.py / verify_draft.py / test_runner.py
│   │   ├── check_env.py / verify.py
│   │   └── audit/               审计脚本
│   ├── annotation-rules.yaml    不确定项标注规则
│   ├── company-map.json         车管公司简称映射
│   └── references/              参考文档
│       ├── official-data-observations.md ⭐
│       ├── decision-tree.md / principles.md
│       └── casebook/            真实错误案例库
├── tests/                       4月回归基准与业务结构审计
└── 导入指南.md                  手动导入步骤（备用）
```

## 失败排查

### `未找到运行中的 AionUi (aioncore)`
→ AionUi 应用没启动。先打开应用，等任务栏出现图标，再跑 setup.bat。

### `需要登录认证`
→ 首次启动要在 AionUi 应用里设管理员账号。设完后跑 setup.bat。如果还失败，说明需要 session——在 AionUi 浏览器里登录后，改用手动导入（见 `导入指南.md`）。

### `导入技能失败`
→ 检查 `skills/<名字>/SKILL.md` 是否存在。aioncore 需要能读到本目录的绝对路径。

### `写入规则失败`
→ 助手已导入但系统提示词没写入。打开 `assistant-rule.md`，手动粘贴到 AionUi 助手编辑页的"规则"框里（先把 `<RESOURCES_DIR>` 替换成本目录 `resources` 的绝对路径）。

### 资源路径问题
助手系统提示词里的 `<RESOURCES_DIR>` 会被 setup.py 自动替换为 `resources/` 的绝对路径。如果手动配置，要自己替换。

## 卸载

在 AionUi 应用里：
1. 设置 → 助手 → 删除「邮局委办报表助手」
2. 设置 → 技能 → 删除三个 post-office-* 技能
3. 删除整个 `aionui-assistant-package` 目录

## 设计理念

助手按 Claude Code 处理邮局报表的工作方式运行：
- **脚本做粗活**（确定性逻辑用 Python 脚本跑）
- **自己补判断**（识别异常、对比历史、追查可疑项）
- **不确定就查**（跑 Python 读源数据验证，不凭记忆）
- **拿不准就标注**（留空+备注写明原因）
- **关键节点问用户**

不追求 100% 正确，重点是**遇到不确定/可能有问题的及时标注**，方便人工二次核对。

## 业务阶段

- Step 1—6：生成固定 **21 列初版**，供人工核对。
- Step 7：从核对后初版生成逐 Sheet **17/18 列拆分正式版**。
- Step 8：从最近有效月份模板动态读取车型组、路线和环比区，生成一/二干盈亏表。

## 模型建议

建议使用支持长上下文、工具调用和代码执行的中高端模型。提示词不能保证正确性；交付前仍需运行 `resources/scripts/test_runner.py` 并遵守人工核对闸门。
