# 宏观界域强化序列库 (Macro-Realm Sequence Library)

## 项目概述
TRPG（桌面角色扮演游戏）"宏观界域"的强化序列资料库。包含特质改造、职业、技能表、能量池、公共建筑等游戏资源。
现役维护制作人：无语（QQ: 1348697221）｜原初制作人：沧羽（QQ: 853304398）

## 目录结构
```
macro-realm-seqlib/
├── .agents/skills/       # 仓库级专属 Agent Skills（Antigravity / Codex 自动发现）
│   ├── seqlib-resource-workflow/  # 资源审稿、入库、泛匹配卡片规范化与只读校验器 lint_resources.py
│   └── seqlib-release/            # 版本发版、重排编号、build_config 与多端同步门禁
├── build_chm.py          # CHM/ZIP 构建脚本
├── build_config.json     # 各版本根目录打包文件清单配置
├── text_encoding.py      # 统一文本解码入口（兼容 UTF-8 / UTF-16 / GBK 回退）
├── .github/workflows/    # GitHub Actions 自动构建、前端检查与在线站部署
├── tools/
│   ├── hhw/              # 内置 HTML Help Workshop（hhc.exe + 依赖 DLL）
│   └── renumber_resources.py  # 资源编号重排与首行标题同步工具
├── web/                  # 在线站（trpg.chordvers.org）前后端源码
│   ├── backend/          # FastAPI 只读检索/历史对照/结团统计/规范化签审/后台维护接口
│   └── frontend/         # Vite + React + TS 高维终端风前端
├── wiki/                 # Fandom Wiki（macro-realm.fandom.com）同步脚本与页面
├── docs/                 # 项目规范与模块说明文档
├── 序列库/               # 当前版本正式在用的资源（约 695 份）
│   ├── 公共建筑/
│   ├── 技能表/
│   ├── 新过审序列/       # 体验服/指定新过审暂存（仅用户明确指定时写入）
│   ├── 能量池/
│   ├── 特质改造/
│   └── 职业/
├── 荣誉室/               # 已下架/归档的历史资源（约 377 份，默认只读保留）
├── 资源标准模板.txt      # 全库资源格式、泛匹配卡片与首部元信息唯一事实标准
├── 6.8序列库编者注.txt
├── V6.8序列库更新日志.txt
├── 第十三批下架名单.txt
└── 第十四批下架名单.txt  # （含历代批次下架名单与历史编者注/更新日志）
```

## 仓库专属 Skills 索引（优先按需加载）
1. **`seqlib-resource-workflow`**（`.agents/skills/seqlib-resource-workflow/SKILL.md`）：
   - 处理资源审稿、新过审入库、增量调整 vs 全量重置合并、以及**基于 AI 语义理解的无损格式规范化**（严禁写死板正则脚本硬改文本）时优先激活。
   - 内置 AI 规范化只读 Harness：
     - 结构探针与语法考官：`python3 .agents/skills/seqlib-resource-workflow/scripts/lint_resources.py --git-changed --strict`（或 `--all --summary` 全库摸底）；
     - 防幻觉与数值守恒对账器：`python3 .agents/skills/seqlib-resource-workflow/scripts/verify_lossless.py --git-changed`（核验数值/骰子/百分比 100% 守恒、正文零漏字、署名零篡改，支持 `--emit-review` 导出人工签审单）。
2. **`seqlib-release`**（`.agents/skills/seqlib-release/SKILL.md`）：
   - 准备版本发布、批量重排编号、更新编者注/更新日志/`build_config.json`、构建 CHM/ZIP 及同步 Wiki 与在线站时激活。

## 构建系统

### 构建命令
```bash
# 完整构建（CHM + ZIP，需先在 build_config.json 配好对应主次版本号）
python build_chm.py --version v6.8

# 只构建 ZIP（macOS / Linux 本地校验常用）
python build_chm.py --version v6.8 --skip-chm

# 只构建 CHM
python build_chm.py --version v6.8 --skip-zip
```

### 输出文件命名格式
`宏观界域强化序列库V{版本号}.chm` / `.zip`
例：`宏观界域强化序列库V6.8.chm`

### 构建依赖
- **Python 3**
- **pandoc** — docx/doc 转 HTML（`winget install JohnMacFarlane.Pandoc`）
- **hhc.exe** — CHM 编译器（已内置于 `tools/hhw/`，无需额外安装）
- **chmcmd** — CHM 编译器（Linux 备选，`sudo apt install fp-utils`，对 CJK 索引支持有缺陷）

### CHM 编码方案（重要）
CHM 格式不支持 UTF-8，构建过程统一转换为 GBK 编码：
1. **内部文件路径保留原始中文目录结构**：`序列库/职业/001】天师.html`，GBK 编码，支持中文全文检索
2. **项目文件（.hhp/.hhc/.hhk）用 GBK 编码**：Language=0x804，Windows CHM 查看器按 GBK 解码
3. **HTML 内容文件用 GBK 编码**：`charset=gbk`，GBK 不支持的字符用 `xmlcharrefreplace` 自动转成 HTML 实体
4. **CHM 编译输出文件名用 ASCII**（`output.chm`），编译后改名为中文
5. **CI 构建时需设置 ACP=936**：英文 Windows 默认 ACP 为 1252，通过注册表改为 936 + `chcp 936`

### CHM 与 ZIP 内容范围
- **CHM 只包含 `序列库/` 目录**（不含荣誉室）
- **ZIP 包含 `序列库/` + `荣誉室/` + `build_config.json` 指定版本的根目录文件**

## CI/CD（GitHub Actions）
- 发布工作流：`.github/workflows/release.yml`（`windows-latest`，推送 `v*` tag 自动创建 GitHub Release 并上传 CHM + ZIP）
- 前端与战术桌检查：`.github/workflows/frontend-check.yml`、`tabletop-check.yml`
- 在线站自动部署：`.github/workflows/deploy-web.yml`（监听 `main` 分支 `web/**` 变更及手动 `workflow_dispatch`）

## 资源更新与模板核心规范（完整版见 `资源标准模板.txt`）
当用户要求“更新资源”“正式更新进库”“按模板更新资源”时，默认目标是正式 `序列库/`，不是 `序列库/新过审序列/`（推荐直接调用 `seqlib-resource-workflow` 技能配合 `lint_resources.py` 执行）。

### 正式入库与荣誉室规则
1. **已有正式资源**：先在 `序列库/` 内按资源名主体匹配旧文件，再把新稿更新到旧文件所在的正式路径。
2. **重置 / 全面重置 / 全翻修**：按新稿全量替换正式库中的对应资源；如果旧资源只在 `荣誉室/`，视为“回归新增”，在正式 `序列库/` 中按分类和新编号新增，**不移动、不删除 `荣誉室/` 旧文件**。
3. **调整 / 加强 / 削弱 / 优化 / 明确描述 / 修 bug / 补充**：按投稿含义增量更新对应正式资源。若投稿只是一两个能力或条目，只替换正式资源中的对应卡片或段落，严禁把整张技能表覆盖成单个能力。
4. **新增资源**：根据正文类型、文件名标注和现有目录分类放入正式 `序列库/` 对应大类/侧别；编号使用目标目录现有最大编号 + 1。
5. **新过审目录与荣誉室边界**：只有用户明确要求放入新过审/体验服时才写入 `序列库/新过审序列/`；只有用户明确要求“下架 / 移入荣誉室”时才改动 `荣誉室/`。

### 泛匹配卡片、格式与首部元信息规则
1. **文件名与首行**：**仅文件名**严禁使用 `/ \ : * ? " < > |` 等系统路径非法字符（正文内部正常允许 `/`、`*`、引号与兼容半角 `:`）；TXT 统一保存为 UTF-8，首行标题必须严格等于文件名去掉 `.txt`，且标题后空一行。
2. **`[XX名称]` 泛匹配成卡原则**：
   - 普通能力/法术/战技统一优先写 `[能力名称]：`；
   - 称号、道具、能量池、模块、义体、符文、召唤物/单位、建筑奖励/商品、职阶、天赋、仪式等专用实体，保留各自 `[XX名称]：`（如 `[称号名称]：`、`[义体名称]：`、`[单位名称]：`、`[奖励名称]：`），解析器与校验器对一切 `[XX名称]` 泛匹配成卡，切勿把专用卡强行改写成 `[能力名称]`；
   - 卡片内部允许长效果用空行自然分段（遇到下一 `[字段名]：`、新 `[XX名称]：` 或小节 `【...】` 才切换）；分级效果使用 `[初级效果]：`～`[终极效果]：`；删除真正为空的可选字段；防止漏写左括号（如 `能力名称]：` 会导致吞卡）。
3. **单卡最近修改备注（方案 B）**：紧贴单张卡片 `[XX名称]` 正上方的最近修改说明，新稿与规范化输出统一加前缀写成 `【能力修改：...】`（存量旧稿的无前缀紧贴备注及卡内 `[能力备注]：` 由读取端兼容，未触及旧稿不强翻）；裸 `【...】` 整行保留给小节标题。
4. **技艺与主成长树**：`技艺规则` 固定写成卡片内独立结构字段 `[技艺规则]：`（后接冥想/学习/习得/传授难度），不另立孤立技艺专卡；特质主成长树统一为 `【E】级：...`～`【S/EX】级：...`。
5. **首部元信息与更新日志**：
   - 标准输出只使用 `（制作人：xxx）`、`（调整人：xxx）`、`（修改人：xxx）`、`（重置人：xxx）`、`（审核人：xxx）`（读取端兼容行首无括号裸写与 `投稿人/制作者/复查人` 等旧称谓）；`调整人`、`修改人`、`重置人` 按原稿性质保留不互换。
   - 统一采用“人员信息块在前、空一行、连续更新记录块在后”；更新记录按日期从早到晚升序排列，同日多条分号合并，行尾圆括号附注保留在同一行。实际内容有变但漏写日志时须对照旧版补写 `【日期：说明】`；规则语义或归属不清处一律保留原文并列入人工复核。

## Wiki（Fandom）
- 地址：**macro-realm.fandom.com**（中文路径 `/zh/`），同步脚本位于 `wiki/sync_to_wiki.py`。
- 常用命令：
  ```bash
  # 试运行预览（不推送）
  python wiki/sync_to_wiki.py --user BotName --password xxx --dry-run

  # 同步全部或仅正式序列库
  python wiki/sync_to_wiki.py --user BotName --password xxx
  python wiki/sync_to_wiki.py --user BotName --password xxx --skip-honor
  ```
- 手动维护页面（主页、名人堂等 `.wikitext`）切勿加入自动同步分类，以免被清理逻辑误删。

## 资源重排工具
发生新增、下架移入 `荣誉室/` 或重置追加后，使用 `tools/renumber_resources.py` 统一重排编号（详见 `docs/renumber_resources.md`）：
```bash
python tools/renumber_resources.py          # 试运行查看计划
python tools/renumber_resources.py --apply  # 确认无误后改名并同步 TXT 首行
```

## Web 在线站前端规范（trpg.chordvers.org）

在线站前后端位于 `web/`（服务启动与接口说明见 `web/README.md`，战术桌见 `docs/tactical-tabletop.md`）。后续迭代 `web/frontend/` 时必须遵守以下防回退底线，完整规范见 **`docs/frontend-style-and-interaction.md`**：

1. **视觉风格基调（无限流高维兑换终端）**：
   - 采用**虚空棱镜暗色**（默认：`#0b0d11` 底 + `#73ded0` 棱镜青 + `#e5c07b` 星辉金，与战术桌同源）与**白昼终端亮色**（`#f3f5f8` 底 + `#ffffff` 纸面 + `#0d7a73` 深青）双主题。
   - 全站坚持 `2px` 微切角与 `1px` 细线分割；**严禁**退回通用后台大圆角气泡卡片/浅蓝模板风，**严禁**偏向中式宣纸古籍等单一题材风，**严禁**在界面中塞无意义的纯大写英文装饰词。
2. **核心交互与无损红线**：
   - **原文 100% 无损**：`Reader/adaptive.ts` 按连续原文区间切分，全文复制与单段复制必须直接取原始 `content` / `raw`，严禁从展示字段反向拼接；文内精确查找高亮时须保留首部人员与更新日志的结构化排版。
   - **虚拟列表与标题契约**：中栏列表（`ResourceCard.tsx`）固定 `104px` 行高，阅读器 `h1` 严格保持原始资源标题，不得随意增删前后缀。
   - **全端自适应**：桌面支持左/中栏独立折叠、多档案标签、专注阅读、主/参照双栏独立滚动对照及历史版本比对；手机端（`<= 820px`）采用「检索列表 ↔ 全屏阅读」单栏切换、抽屉筛选、贴边通栏能力卡与 `16px` 防自动放大输入框。
3. **改后必跑校验**（在 `web/frontend/` 下执行）：
   - `npm run build` 与 `node --test tests/*.test.cjs`；涉及布局或响应式改动时需通过 `tests/browser.cjs`、`tests/workbench-browser.cjs`、`tests/reading-layout-browser.cjs` 多视口回归。
