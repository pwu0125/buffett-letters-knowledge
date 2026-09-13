# AGENTS.md

## 适用范围与目标

本文件适用于仓库根目录及全部子目录。`skills/*/SKILL.md` 是应用内的巴菲特/芒格人格资产，不是本仓库的协作说明。

本项目的核心产物是一个单文件中文阅读应用：知识库 Markdown、分类索引、收益数据和人格材料经 `build_buffett_app.py` 编译为 `巴菲特投资智慧.html`，再由本地服务或 Mac/Windows 壳运行。阅读、搜索和笔记可完全离线使用；可选的 LLM 讨论仍需用户配置联网接口。修改时优先保证：

1. 知识内容、分类和内部链接准确；
2. 单文件应用仍可离线使用；
3. API Key 与用户笔记、收藏、已读和对话数据不进入仓库；
4. 安装包确实包含当前生成物，而不是旧 HTML；
5. 围绕请求做最小改动，不顺带清洗语料、重构大型模板或扩大发布范围。

用户可见文案、说明和交付结论默认使用中文；代码标识符遵循现有文件风格。

## 开始工作前

- 先运行 `git status --short --branch`，保留并避开用户已有改动。
- 阅读根 `README.md`，再读与任务直接相关的源文件。知识内容任务还要查看 `巴菲特致股东信知识库/README.md`。
- 先区分源文件、受版本控制的生成物、原站快照和本机数据；不要在错误层修改。
- 仓库目前没有正式测试套件、CI、`pyproject.toml`、`requirements.txt`、`package.json` 或 `Makefile`。不要声称运行了不存在的 pytest/npm/CI，也不要为了形式完整新增工程设施。
- 普通内容或代码修改不授权打包、发布、真实 LLM 请求或覆盖用户状态；这些动作按下文边界单独处理。

## 项目地图与事实来源

| 路径 | 角色 | 编辑规则 |
| --- | --- | --- |
| `巴菲特致股东信知识库/01-索引` 至 `07-人物` 下第一层 `*.md` | 应用文章正文源 | 应用构建只读取这七个固定目录的第一层 Markdown，不递归 |
| 同目录同名 `*.html` | 原站页面快照 | 构建器不读取；不是 Markdown 自动生成物，不要因改 Markdown 就批量覆盖 |
| `巴菲特致股东信知识库/README.md` | 人工维护的知识库目录 | 构建器不读取；新增、删除或改名文章时同步维护 |
| `巴菲特致股东信分类索引(1956-2025) .xlsx` | 五维分类、年度摘要和原文链接源 | 文件名在 `.xlsx` 前有真实空格；用表格工具编辑，不要当普通文本改 |
| `buffet_return_data.csv` | 首页 Hero 收益图数据源 | 覆盖 1957–2025 三段口径：1957–1964 合伙基金收益、1965–2018 每股账面价值变动、2019–2025 每股市场价值变动（官方年报业绩表）；三段必须连续复利，图表注释需写明分段口径，不要删除 2019 年后数据 |
| `build_buffett_app.py` | 数据编译器，也是 HTML/CSS/JS 前端源码 | UI 与构建逻辑在这里改，不直接补丁生成 HTML |
| `assets/*-128.png`、`vendor/echarts.min.js` | 单文件应用构建输入 | 不格式化 vendor 文件；缺失时构建可能静默降级 |
| `skills/celebrity-buffett/{persona.md,work.md}`、`skills/celebrity-charlie-munger/{persona.md,work.md}` | 应用实际嵌入的人格正文 | 改后必须重建 HTML；前端每次对话最多取人格内容前 14,000 字符 |
| `巴菲特投资智慧.html`、`llm-config.js` | 已跟踪的生成物 | 不手改；由构建器生成并审阅 diff |
| `serve_buffett_app.py` | 仅监听本机回环地址的静态服务、同源运行时 LLM 配置和状态 API | 状态测试必须使用隔离目录 |
| `macapp/`、`winapp/`、`package_app.py`、`package_windows.py` | 平台壳与安装包构建 | 只在相关平台变更或明确交付安装包时验证 |
| `release.py` | GitHub Release 外部写入 | 只有得到明确发布授权后才实际运行 |

根目录的 `buffet_return_*.jpg/png` 是独立研究/展示图片，当前运行时不读取。`skills/*/knowledge/` 下的 transcript、SRT、PDF 和研究文档是溯源材料；未经明确任务，不批量重编码、重命名、去重或删除。

## 知识库合同

### 文章身份与目录

- 七个分类目录及其应用键在 `build_buffett_app.py` 的 `CATS` 中硬编码。新增文章必须直接放在相应目录第一层。
- 文章 ID 为“分类键 + 文件 stem”。文件名或分类移动会改变 ID，已有笔记、收藏、已读记录和内部链接可能因此失联；未经明确迁移设计，不重命名或移动既有文章。
- 文件 stem 在七类文章中应全局唯一。内部链接按目标 basename/stem 全局解析，跨分类重名会覆盖映射。
- 标题取正文中出现的第一个 `# ` 一级标题；年份取文件名中出现的第一个四位年份。新增内容要保持这两个约定。
- 现有 Markdown 多使用指向原站 `.html` 路径的相对链接。构建器会按 stem 改写为应用内 `#/a/...`；解析不到的相对链接会降级为纯文本。新增、改名或移动文章后检查相关入链与出链。
- 同名 HTML 是原始快照。是否补充或更新快照必须由任务目标决定，不能假设它能从 Markdown 重建。

### 分类索引与注入内容

- XLSX 的工作表名和列布局是构建合同。当前使用 `主题分类索引`、`行业分类索引`、`事件时期索引`、`选股方法演进`、`年度总索引`。
- 年度信件开头的“核心主题摘要”和末尾的“原文链接”由 `年度总索引` 在构建时注入。四篇专题/合集文章使用自身正文，不应套用同年份年度信件的摘要或原文链接。应优先修改 XLSX 对应字段，不要在源 Markdown 重复写入相同区块。
- 同年合伙人信和伯克希尔信会合并为一个年度索引项；Excel 未覆盖的新年份只会得到简化兜底，不能视为完整分类资料。
- `buffet_return_data.csv` 当前字段为 `年份,第N年,年收益率%,模拟累计净值,年化收益%`，覆盖 1957–2025 共 69 行，分三段口径：1957–1964 合伙基金收益率、1965–2018 伯克希尔每股账面价值变动率、2019–2025 每股市场价值变动率（与官方年报业绩表 2019 年起改用每股市场价值的口径切换一致）。三段按年连续复利；`年化收益%` = 累计净值^(1/第N年) − 1。调整时核对年份、累计净值和年化收益之间的一致性，并同步图表分段口径注释。

### 当前基线

以下数字用于发现意外漏读或数据丢失，不是永久不变的硬门禁；若任务有意新增或删除内容，应同步更新相关说明：

- 188 篇文章，分类数依次为 `4 / 17 / 60 / 4 / 35 / 61 / 7`；
- 年份范围 `1956–2025`；
- 分类索引为主题 10、行业 13、事件 13、方法 7、合并后年度 70；
- 构建时摘要注入 77 篇、原文链接注入 77 篇；四篇专题/合集文章不注入年度信件内容；
- 收益 CSV 为 1957–2025 共 69 行数据（三段口径连续复利）；
- 每篇 Markdown 当前都有一份同 stem 的原站 HTML 快照。

## 人格 Skill 合同

- 应用构建只消费每套 skill 的 `persona.md` 和 `work.md`。`SKILL.md` 是组合入口，`persona_skill.md` / `work_skill.md` 是独立入口，`manifest.json` / `meta.json` 是元数据。
- 修改前先明确目标消费者。只影响应用时改实际嵌入文件并重建；若意图影响所有 skill 入口，再同步复核 `SKILL.md`、两个 `*_skill.md` 及必要的 manifest/meta 字段。不要无条件整套重写。
- `knowledge/research/reviews/` 和 `merged/summary.md` 可能是某次研究阶段的历史快照，不自动代表当前全量语料。
- manifest/meta 中列出的 distilly `prompts/`、`references/`、`tools/research/*` 多数不在本仓库；它们是生成来源元数据，不是可直接运行的本地命令。
- 修改 JSON 后至少用 `python3 -m json.tool <file> >/dev/null` 校验涉及的 manifest/meta。

## 构建与生成物

### 常规重建

知识 Markdown、XLSX、收益 CSV、128px 头像、ECharts、人格正文或 `build_buffett_app.py` 有变化时，在仓库根运行：

```bash
python3 build_buffett_app.py --debug --no-llm-config
```

使用 `--no-llm-config` 是为了避免本机被忽略的 `.env` 中 `DEEPSEEK_API_BASE` 造成无关 diff。只有明确修改默认 LLM 配置时才省略该参数；此后必须确认 `llm-config.js` 的 `key` 仍为 `""`，且没有写入本机私有 endpoint。

`build_buffett_app.py` 没有 argparse 帮助处理；执行 `python3 build_buffett_app.py --help` 也会真实构建，不要把它当只读探测命令。

分类 XLSX 是必需输入，缺失会令当前构建失败。CSV、ECharts 或头像缺失时可能静默降级，人格缺失时会警告后降级；因此构建成功退出仍不保证内容完整。正式产物应检查控制台警告、上述基线、生成 diff 和实际页面。生成 HTML 会写入当天 `built` 日期，因此跨日重建可能只有日期变化；不要为无关修改顺手重建，也不要把“构建后 diff 必须为空”设为门禁。

### 生成物规则

- 不直接编辑 `巴菲特投资智慧.html`。它包含大型内嵌 JSON/CSS/JS，人工补丁会在下次构建丢失。
- `llm-config.js` 只保存默认 base/model，不保存 API Key。
- 源输入变化时提交相应源文件和必要的最新生成物；纯文档修改不需要重建应用。
- 重建后先看 `git diff --stat`，再针对源文件和生成物检查有意义的 diff；不要格式化整份单文件 HTML。

## 本地服务、密钥与用户数据

- 开发入口是 `python3 serve_buffett_app.py`；自动化或烟测加 `--no-browser`。服务必须继续只绑定 `127.0.0.1`，不要未经安全设计改为局域网或公网监听。
- 服务会在首选端口及随后 19 个端口中选择可用端口，以控制台打印的实际地址为准。
- `/api/state` 使用整份状态替换语义。当前前后端合同字段是 `notes`、`favs`、`read`、`chat`、`buffett_chat`、`munger_chat`；字段变化必须同步前端映射、服务端白名单和往返测试。远端返回的显式空对象/数组也是权威状态，不能被浏览器旧数据反向覆盖。
- 默认状态文件位于 macOS/Linux 的 `~/Library/Application Support/巴菲特投资智慧/state.json`，Windows 位于 `%APPDATA%\巴菲特投资智慧\state.json`。不得删除、重置或用测试请求覆盖真实用户数据。
- 服务测试必须把 `BUFFETT_DATA_DIR` 指向新建的临时目录，并选用非生产端口。例如在一个终端运行：

```bash
BUFFETT_DATA_DIR="$(mktemp -d)" python3 serve_buffett_app.py --port 8876 --no-browser
```

再在另一终端用控制台显示的实际端口检查 `/`、`/api/llm-config`、`/api/state`，需要时测试一次状态 PUT/GET 往返；仓库内其他文件应返回 404，非 loopback `Host` 应返回 403。

- `.env`、`*.env`、API Key 和真实用户状态不得提交或暴露。`*.log` 与 `.buffett-data/` 不得提交；诊断时只报告与问题直接相关且已脱敏的日志片段。应用设置中手动填写的 Key 只应留在浏览器 localStorage；服务端 Key 只从环境变量或项目根 `.env` 读取，经 `/api/llm-config` 注入页面内存，不回填 localStorage。静态服务只应暴露应用 HTML、无密钥 `llm-config.js` 和图标，不得开放任意仓库文件。
- 除非任务明确要求真实连接，不发送真实 LLM 请求；UI/服务烟测使用空 Key 或 dummy Key，并避免点击“测试连接”。

## 平台打包

两个打包脚本都只复制现有 `巴菲特投资智慧.html`，不会调用构建器；`release.py --build` 也只调用两个打包脚本。任何会影响 HTML 的源变更都必须先完成常规重建，再打包。

### macOS

```bash
python3 package_app.py --version X.Y --no-dmg
python3 package_app.py --version X.Y
```

- 依赖 macOS 的 `xcrun/swiftc`、`sips`、`iconutil`、`plutil`、`codesign`、`hdiutil`。
- 默认目标为 `arm64-apple-macos12.0`，可由 `BUFFETT_SWIFT_TARGET` 覆盖；目标 Mac 仍需有 `python3`。
- Swift 壳会扫描 8666–8685，并连接最低的已有应用服务；若没有可用服务，则让 Python 在同一范围内选择端口并使用其实际地址。修改端口策略时必须同步检查 `serve_buffett_app.py` 与 `macapp/main.swift`，不能只验证命令行服务。
- 退出 App 时，Swift 壳会先等待页面通过 `window.BUF.flushState()` 完成普通 PUT，再停止自己启动的服务，最多等待 3 秒。不要用 `pagehide` 的 `keepalive` 请求替代这段握手：较大的笔记/对话状态会超过浏览器 keepalive 请求体上限。
- `.app` 实际生成在 `dist/stage/巴菲特投资智慧.app`，DMG 在 `dist/巴菲特投资智慧-vX.Y.dmg`。
- `--no-dmg` 可作较轻的包结构验证；正式交付仍需在目标架构 Mac 上验证启动、外部链接、状态持久化和关闭窗口后服务退出。

### Windows

```bash
python3 package_windows.py --version X.Y
```

- 需要 Pillow、`x86_64-w64-mingw32-gcc` 和 `makensis`；首次还会从 python.org 下载固定的 Python 3.12.8 x64 embeddable 包到被忽略的 `vendor/python-embed/`。
- `--no-python-download` 只允许使用已有缓存，不代表无需其他工具。
- 产物为 `dist/巴菲特投资智慧-vX.Y-Setup.exe`。交叉编译成功不等于 Windows 验收；正式交付需在 64 位 Windows 10/11 上验证安装、启动、端口回退、停止服务、覆盖升级和卸载保留用户数据。

两个平台的 `BUNDLED` 清单需要保持意图一致。若新增必须随安装包分发的运行时文件，应同时检查两份清单。安装包中虽带有构建脚本，但没有复制全部仓库级原始构建输入；正式 HTML 必须在仓库根生成，不要把安装包内部当作等价构建环境。

`dist/`、DMG、EXE 和 `vendor/python-embed/` 均被忽略。打包脚本会删除并重建各自的 `dist/stage*`，所以它们是写操作，不是普通只读校验，也不是所有改动的必跑项。

## 发布边界

- 只有用户明确要求发布时，才运行非 dry-run 的 `release.py`。发布前人工确认目标仓库、当前 HEAD、版本号、两个精确命名的安装包和拟操作的 Release。
- `python3 release.py --version X.Y --dry-run` 不会写远端，但仍要求本地产物、已登录的 `gh` 并读取 GitHub；若同时加 `--build`，仍会真实写 `dist/`，也可能触发 Windows Python 下载。
- `--repos` 指定的仓库必须传给所有 `gh release` 查询、创建、下载和上传命令；`gh auth status` 不带仓库参数。
- 发布器必须确认当前 HEAD 已存在于目标仓库；新标签显式以该 SHA 为 `--target`，已有标签必须解析到同一 SHA。附件缺失或大小不符必须以非零状态退出。
- Release 成功只证明远端附件上传且字节大小匹配，不证明安装、启动或功能验收。

## 与改动相称的验证

| 改动 | 最小相关验证 |
| --- | --- |
| 仅文档 | `git diff --check`，人工核对路径和命令 |
| 知识 Markdown / XLSX / CSV / 人格 / 前端构建器 | `python3 build_buffett_app.py --debug --no-llm-config`，检查警告、基线和生成 diff；仅在渲染、链接、索引、图表或人格行为受影响时浏览器冒烟相关页面 |
| `serve_buffett_app.py` 或状态合同 | Python 语法检查；隔离 `BUFFETT_DATA_DIR` 的 GET 与 PUT/GET 烟测 |
| `启动巴菲特知识库.command` | `zsh -n 启动巴菲特知识库.command` |
| `macapp/main.swift` | `xcrun swiftc -typecheck -swift-version 5 -target arm64-apple-macos12.0 macapp/main.swift -framework Cocoa -framework WebKit`；交付安装包时再做目标机验证 |
| `winapp/*.c` / NSIS / Windows 打包器 | 有依赖时做对应编译/打包；交付时在 Windows 10/11 验收 |
| `manifest.json` / `meta.json` | 对修改的 JSON 运行 `python3 -m json.tool` |
| 发布流程 | 先 dry-run 和人工核对；获授权后发布，再检查远端附件和目标平台安装 |

UI 相关浏览器冒烟至少覆盖本次受影响路径，并按需检查：主页、全文搜索、文章阅读与内部链接、五维索引、收益图、笔记/收藏/已读持久化、设置，以及巴菲特/芒格对话入口。不要为无关改动重复跑完整平台打包。

## 完成与交付

- 运行 `git diff --check` 和 `git status --short`，确认没有 `.env`、状态文件、日志、缓存或打包产物进入变更。
- 说明实际修改了什么、用户可见影响、运行了哪些验证，以及哪些平台或真实联网环节未验证。
- 不把“构建成功”“交叉编译成功”“上传成功”分别冒充页面体验、目标机安装或正式发布验收。
- 未经要求不创建提交、标签或 Release，不清理已有语料和用户数据。
