# 巴菲特投资智慧 · 致股东信知识库阅读研究

一个**完全离线的单文件 Web 应用**：把巴菲特致股东信知识库（合伙人信 1956–1969、伯克希尔股东信 1965–2024、概念/公司/人物词条）内嵌为单个 HTML，支持多维度分类、全文搜索、阅读划线、笔记与 LLM 讨论。

## 功能

- **主页**（点击左上角「巴菲特投资智慧」随时返回）：统计总览、全文搜索、快速浏览入口（信件/概念/公司/人物）、分类索引五维入口、必读经典书单、最新信件、热门主题、继续阅读（自动记录上次阅读位置）

- **多维度分类排序**
  - 基础分类：索引 / 合伙人信 / 伯克希尔股东信 / 特别信件 / 概念 / 公司 / 人物
  - 年代分组（1950s–2020s）、主题标签（从 1.2 万条交叉链接自动提取）
  - **分类索引五维**（解析自 `巴菲特致股东信分类索引(1956-2025) .xlsx`）：坎宁安主题 ×10、行业 ×13（含纺织(已退出)）、市场事件时期 ×13（含 Go-Go 投机狂潮、1973-74 大熊市）、选股方法演进 ×7、逐年索引 **1956–2025** 全覆盖（含撰写人/信件系列/核心摘要/事件标签/原文链接；1965-1969 合伙与伯克希尔双系列自动合并），筛选时顶部展示该分类的完整解析（市场背景 / 巴菲特行动 / 经验教训等）
- **全文搜索**：标题+正文检索、命中高亮、与各类过滤自由组合
- **摘要与原文**：每封信件开头展示「核心主题摘要」引用块、文末附「原文链接」（伯克希尔官网 / chian.io 原站，可点击打开）
- **阅读**：内置 Markdown 渲染（表格/引用/脚注/交叉链接应用内跳转）、目录、上一篇/下一篇、进度条
- **笔记与划线**：选中文字 → 高亮（黄/蓝）/ 下划线 / 笔记；**笔记正文支持 Markdown 渲染**（标题/加粗/列表/引用/代码块/链接，AI 答复保存后直接呈现排版）；文章级笔记自动保存；一键导出 `.md` + `.json`
- **AI 讨论**：每篇文章旁聊天面板，自动携带文章全文与你的笔记；每条 AI 答复可一键「加入笔记」；OpenAI 兼容接口（默认 DeepSeek）
- **与巴菲特对话 / 与芒格对话**：基于 `celebrity-buffett` / `celebrity-charlie-munger` 人格（distilly 蒸馏，项目内 `skills/` 下）的双对话专栏，主页双入口（巴菲特头像 / 芒格头像），护城河、多元思维模型等框架先研究再回答；两套对话记忆独立持久化
- **主页 Hero 交互图**：巴菲特历年表现（1957–2025；1957–1964 为合伙基金收益，1965–2018 为伯克希尔每股账面价值变动率，2019–2025 为每股市场价值变动率——与官方年报业绩表的口径切换一致）——年度变动率柱状图（正琥珀/负深红）+ 累计净值对数轴曲线 + 年化收益虚线，十字线 tooltip、图例开关、滚轮/滑块缩放，右上角标注分段口径；数据来自 `buffet_return_data.csv`（OCR 提取 + 官方年报交叉核对 + 复算校验）。
- **侧栏隐藏/显示**：顶栏 `◧/◨` 按钮一键收起/展开左侧栏，正文与 Hero 图随之伸缩，偏好持久化
- **流式输出**：AI 讨论与两路人格对话均为 SSE 流式逐字渲染
- **背景解释**：选中文章中的词语 →「🔍 背景解释」，AI 结合文章内容与**写作年份的市场背景**（从年度索引自动注入：市场环境、重大事件、核心主题）解释金融概念；支持就解释继续追问；满意的解释可保存到笔记「背景解释」栏，随时回顾，导出时一并包含
- **编辑风索引纵览**：顶栏「📖 索引」进入独立的杂志风分类索引页（`#/index`），Playfair 衬线标题 + 暖米色排版，含 Hero 统计、5 个标签页（年度信件 / 主题分类 / 行业分类 / 事件时期 / 选股方法）、信件系列与事件类型筛选；信件卡片可直接跳转应用内阅读，与主应用共享全部数据，不替换原有界面

## 安装版（Mac / Windows 可安装 App）

### Mac 版（原生窗口 + Dock 图标）

```bash
python3 package_app.py          # 生成 dist/巴菲特投资智慧.app + dist/巴菲特投资智慧-vX.Y.dmg
```

- **DMG 安装**：把 `巴菲特投资智慧.app` 拖入 Applications（或任意位置）即可；首次打开如被 Gatekeeper 拦截，右键 → 打开，或 `xattr -dr com.apple.quarantine "/Applications/巴菲特投资智慧.app"`
- **原生体验**：双击 App → 打开原生窗口（Swift + WKWebView 承载页面），**Dock 图标停留**，后台自动拉起本地服务（首选 127.0.0.1:8666，端口占用时回退至 8685）；关闭窗口即退出并停止服务
- **外部链接**：原文链接等外部网页自动用系统默认浏览器打开，应用内始终停留在本地页面
- **自包含**：188 篇文章 / 分类索引 / 巴菲特人格 Skill / 构建脚本全部内嵌于 App 内，不依赖安装路径，完全离线可用
- **LLM 密钥**：应用内设置面板填写（仅存本机浏览器 localStorage），或先 `export DEEPSEEK_API_KEY=sk-xxx` 再启动（环境变量密钥不写入应用状态）

### Windows 版（NSIS 安装器 + 卸载器，内置 Python 运行时）

```bash
python3 package_windows.py      # 生成 dist/巴菲特投资智慧-vX.Y-Setup.exe（安装时自动生成 uninstall.exe 卸载器）
brew install makensis mingw-w64 # 一次性前置依赖（NSIS 3.x + MinGW 交叉编译器；首次打包还会下载 vendor 缓存）
```

- **安装**：双击 `巴菲特投资智慧-vX.Y-Setup.exe` → 按用户安装到 `%LOCALAPPDATA%\巴菲特投资智慧`（无需管理员权限），创建开始菜单/桌面快捷方式与「应用和功能」卸载条目；未签名 SmartScreen 提示点「更多信息 → 仍要运行」即可
- **卸载**：开始菜单「卸载巴菲特投资智慧」或「应用和功能」→ 卸载；自动停止本地服务、清理文件/快捷方式/注册表，**默认保留**笔记数据（`%APPDATA%\巴菲特投资智慧`），可勾选一并删除
- **启动体验**：双击「巴菲特投资智慧」→ 静默拉起内置 Python 本地服务（127.0.0.1:8666，端口占用自动回退 8667+）→ 打开 Edge 应用模式窗口（无地址栏，类原生窗口）；再次启动复用已有服务；「停止服务」单独结束后台服务
- **自包含**：无需目标机器安装 Python（捆绑 python.org embeddable 运行时，服务脚本零第三方依赖）；64 位 Windows 10/11
- **数据位置**：`%APPDATA%\巴菲特投资智慧\state.json`（与 Mac 版同结构；`serve_buffett_app.py` 已按平台自动分流）
- 完整说明见安装包内「使用说明.txt」

## 记忆材料持久化（笔记 / 收藏 / 已读 / AI 对话）

- 笔记、划线高亮、收藏、已读标记、文章 AI 讨论、与巴菲特对话、与芒格对话，都会**实时保存到本地文件**：`~/Library/Application Support/巴菲特投资智慧/state.json`
- 该目录属于**本机用户数据**，任何 Git 仓库都不会包含它——上传 GitHub 不涉及这部分信息；如需改存项目目录，可设 `BUFFETT_DATA_DIR`（此时项目 `.gitignore` 已忽略 `.buffett-data/` 兜底）
- 环境变量或项目根 `.env` 提供的 API Key 只进入本地服务内存，不写入应用状态；在设置面板手动填写的 Key 仅保存在本机浏览器 localStorage
- 直接双击 html（无服务）时自动降级为浏览器 localStorage，不影响使用

## 快速开始

**唯一入口：`python3 serve_buffett_app.py`**（本地服务 + 环境变量密钥注入 + 自动打开浏览器）

- 普通使用（推荐）：双击「巴菲特投资智慧.app」（原生窗口版，见上）
- 开发调试：双击「启动巴菲特知识库.command」或终端运行 `python3 serve_buffett_app.py`
- 直接双击 html 也可离线阅读（此时在应用「设置」面板手动填写 API Key，仅存本机浏览器）

服务首选地址为 `http://127.0.0.1:8666/巴菲特投资智慧.html`；端口可用 `BUFFETT_PORT` 调整，若被占用会自动尝试随后 19 个端口，请以启动日志打印的实际地址为准。

## LLM 密钥安全设计

- **任何受版本控制的文件和生成物都不包含 API Key**。`llm-config.js` 只有默认 base/model（`key: ""`）。本机 `.env` 已被 Git 忽略。
- `serve_buffett_app.py` 启动本地服务时从环境变量 `DEEPSEEK_API_KEY` 读取密钥（未设置则回退脚本同目录的 `.env`），前端通过受同源策略保护的 JSON 接口 `/api/llm-config` 在运行时读取；服务只暴露应用入口，不把仓库文件作为静态资源提供。
- 设置面板不会显示或回存服务端注入的 Key；只有用户主动手填的 Key 才会保存在当前浏览器 localStorage。
- 默认模型 `deepseek-v4-flash`；可在应用「设置」面板修改 Base / Key / Model。

## 重新构建

知识库或分类索引有更新后：

```bash
python3 build_buffett_app.py        # → 巴菲特投资智慧.html
python3 package_app.py              # → dist/巴菲特投资智慧.app + DMG（Mac）
python3 package_windows.py          # → dist/巴菲特投资智慧-vX.Y-Setup.exe（Windows）
python3 release.py --version X.Y    # → 把 Mac DMG + Windows Setup.exe 一起发布到 GitHub Release
```

仅依赖 Python 3 标准库（含纯标准库 xlsx 解析器，无需 openpyxl）；App 壳需 macOS Command Line Tools（自带 swiftc）；Windows 安装器需 `brew install makensis`，并首次联网下载 embeddable Python 到 `vendor/python-embed/` 缓存（之后可 `--no-python-download` 离线打包）。

## 发布 GitHub Release（Mac + Windows 安装包）

```bash
python3 release.py --version 2.1            # 安装包已存在时直接发布
python3 release.py --version 2.1 --build    # 先打包 Mac + Windows 再发布
python3 release.py --version 2.1 --dry-run  # 演练：只打印将执行的操作
python3 release.py --version 2.1 --repos owner/repo  # 明确发布到指定仓库
```

- 生成/更新标签 `vX.Y` 对应的 Release，附件名统一为 ASCII（`BuffettWisdom-vX.Y.dmg` / `BuffettWisdom-vX.Y-Setup.exe`，规避 GitHub 对非 ASCII 附件名的截断），重复执行自动覆盖
- 依赖 gh CLI 并已登录：`brew install gh && gh auth login`
- 发布地址：`https://github.com/pwu0125/buffett-letters-knowledge/releases/tag/vX.Y`

## 目录结构

```
├── 巴菲特投资智慧.html               # 单文件应用（数据全部内嵌，约 3.7MB）
├── 巴菲特致股东信知识库/             # 原始文章（md + 原站 html）
├── 巴菲特致股东信分类索引(1956-2025) .xlsx  # 五维分类索引数据源
├── buffet_return_data.csv            # 历年收益 1957–2025（合伙基金 / 每股账面价值 / 每股市场价值三段口径，Hero 图数据源）
├── build_buffett_app.py              # 构建脚本（内嵌 ECharts / 图标 / 收益率）
├── serve_buffett_app.py              # 本地启动器（同源运行时配置 + /api/state 记忆持久化）
├── 启动巴菲特知识库.command           # macOS 一键启动（终端方式）
├── macapp/main.swift                 # 原生窗口壳（Swift + WKWebView，Dock 图标）
├── package_app.py                    # App/DMG 打包器（Mac）
├── winapp/                           # Windows 版源文件：C 启动器/停止服务（launcher.c/stopsvc.c）、NSIS 模板、使用说明
├── package_windows.py                # NSIS 安装器打包器（Windows，macOS 上交叉打包）
├── release.py                        # 一键发布 GitHub Release（Mac DMG + Windows Setup.exe）
├── assets/                           # buffett.png / munger.png（128px 内嵌图标源）
├── vendor/echarts.min.js             # ECharts 5.5.1（构建时内嵌，离线可用）
├── vendor/python-embed/              # embeddable Python 缓存（Windows 打包用，可离线重建）
├── skills/celebrity-buffett/         # 项目级巴菲特人格 Skill（构建时嵌入应用）
├── skills/celebrity-charlie-munger/  # 项目级芒格人格 Skill（构建时嵌入应用）
└── llm-config.js                     # 默认 LLM 配置（无密钥）
```

## 数据来源与致谢

- 知识库内容整理自 [buffett-letters-eir.pages.dev](https://buffett-letters-eir.pages.dev/)（巴菲特致股东信中文整理），原文版权归伯克希尔·哈撒韦 / 沃伦·巴菲特及其译者所有，本仓库仅作个人学习研究用途。
- 分类索引由人工整理（坎宁安主题框架参考 Lawrence Cunningham 的《The Essays of Warren Buffett》分类体系）。
