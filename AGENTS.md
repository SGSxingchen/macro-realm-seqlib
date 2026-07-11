# 宏观界域强化序列库 (Macro-Realm Sequence Library)

## 项目概述
TRPG（桌面角色扮演游戏）"宏观界域"的强化序列资料库。包含特质、职业、技能表、能量池、公共建筑等游戏资源。
制作人：沧羽（QQ: 853304398）

## 目录结构
```
序列库方案/
├── build_chm.py          # CHM/ZIP 构建脚本
├── .github/workflows/    # GitHub Actions 自动发布
│   └── release.yml
├── tools/hhw/            # 内置的 HTML Help Workshop（hhc.exe + 依赖 DLL）
├── 序列库/               # 当前版本在用的资源（724个文件）
│   ├── 公共建筑/
│   ├── 技能表/
│   ├── 能量池/
│   ├── 特质改造/
│   └── 职业/
├── 荣誉室/               # 已下架/归档的历史资源（326个文件）
├── 6.6序列库编者注.txt
├── V6.6序列库更新日志.txt
├── 第十批下架名单.txt
└── 第十一批下架名单.txt
```

## 构建系统

### 构建命令
```bash
# 完整构建（CHM + ZIP）
python build_chm.py --version v6.6

# 只构建 ZIP
python build_chm.py --version v6.6 --skip-chm

# 只构建 CHM
python build_chm.py --version v6.6 --skip-zip
```

### 输出文件命名格式
`宏观界域强化序列库V{版本号}.chm` / `.zip`
例：`宏观界域强化序列库V6.6.chm`

### 构建依赖
- **Python 3**
- **pandoc** — docx/doc 转 HTML（`winget install JohnMacFarlane.Pandoc`）
- **hhc.exe** — CHM 编译器（已内置于 `tools/hhw/`，无需额外安装）
- **chmcmd** — CHM 编译器（Linux 备选，`sudo apt install fp-utils`，对 CJK 索引支持有缺陷）

### CHM 编码方案（重要）
CHM 格式不支持 UTF-8，全程统一使用 GBK 编码：
1. **内部文件路径保留原始中文目录结构**：`序列库/职业/001】天师.html`，GBK 编码，支持中文全文检索
2. **项目文件（.hhp/.hhc/.hhk）用 GBK 编码**：Language=0x804，Windows CHM 查看器按 GBK 解码
3. **HTML 内容文件用 GBK 编码**：`charset=gbk`，GBK 不支持的字符用 `xmlcharrefreplace` 自动转成 HTML 实体
4. **CHM 编译输出文件名用 ASCII**（`output.chm`），编译后改名为中文
5. **CI 构建时需设置 ACP=936**：英文 Windows 默认 ACP 为 1252，通过注册表改为 936 + `chcp 936`

### CHM 内容范围
- **CHM 只包含 `序列库/` 目录**（不含荣誉室）
- **ZIP 包含全部**（序列库 + 荣誉室 + 根目录文件）

## CI/CD（GitHub Actions）
- 工作流文件：`.github/workflows/release.yml`
- **运行器：`windows-latest`**（必须用 Windows，hhc.exe 是 Windows 工具）
- **触发方式：**
  - 推送 `v*` tag → 自动创建 GitHub Release 并附带 CHM + ZIP
  - 手动触发 (workflow_dispatch) → 上传为 Artifact 供下载测试
- CI 会自动安装 pandoc（choco），hhc.exe 已内置于仓库 `tools/hhw/` 中

### 发布流程
```bash
git add .
git commit -m "更新内容"
git tag -a v6.6 -m "发布 V6.6 序列库更新"
git push origin main
git push origin v6.6
# GitHub Actions 自动构建并创建 Release
```

## 内容规范（编者注）
1. 资源文件名和文件内部不要添加特殊字符
2. txt 首行为标题，标题后空一行
3. txt 保存为 UTF-8 编码
4. 支持 txt 以外的文件，但需另存为 html 格式
5. 文件编号格式：`001】名称`，用于排序

## 资源更新流程（重要）
当用户要求“更新资源”“正式更新进库”“按模板更新资源”时，默认目标是正式 `序列库/`，不是 `序列库/新过审序列/`。

### 正式入库规则
1. **已有正式资源**：先在 `序列库/` 内按资源名主体匹配旧文件，再把新稿更新到旧文件所在的正式路径。
2. **重置 / 全面重置 / 全翻修**：按新稿全量替换正式库中的对应资源；如果旧资源只在 `荣誉室/`，视为“回归新增”，在正式 `序列库/` 中按分类和新编号新增，不移动、不删除荣誉室旧文件。
3. **调整 / 加强 / 削弱 / 优化 / 明确描述 / 修 bug / 补充**：按投稿含义更新对应正式资源。若投稿只是一两个能力或条目（例如某个法术调整），只替换正式资源中的对应能力卡或段落，不要把整张技能表覆盖成单个能力。
4. **新增资源**：根据正文类型、文件名标注和现有目录分类放入正式 `序列库/` 对应大类/侧别；编号使用目标目录现有最大编号 + 1，除非该目录内已有明确编号规则。
5. **新过审目录**：只有用户明确要求“放入新过审”“更新新过审序列”“体验服目录”时，才写入 `序列库/新过审序列/`。不要因为来源文件夹名包含“新过审”就自动放入该目录。

### 荣誉室规则
1. `荣誉室/` 是已下架/历史资源留档，默认只读保留。
2. 正式库重置回归、新增同源新版时，不删除、不移动荣誉室旧条目。
3. 只有用户明确要求“下架”“移入荣誉室”“清理荣誉室”时，才改动 `荣誉室/`。

### 模板与格式规则
1. 按 `资源标准模板.txt` 做格式整理：首行标题与文件名一致、标题后空一行、UTF-8 编码、字段名优先使用 `[字段名]：`。
2. 格式整理不得影响数值、强度、冷却、消耗、持续时间、判定方式、适用对象、限制条件、规则优先级等实质规则。
3. 不要为了统一模板而改写投稿数值、删减效果、合并有歧义字段、改变技能/特质/职业/能量池的机制含义。
4. 只做无歧义的格式修正；涉及规则含义、数值平衡、审核归属不清或全量/增量边界不清时，保留原文并列入人工复核。
5. `.docx` 可保持 `.docx` 覆盖原文件；只有用户要求或目标目录规范需要时，才转换为 `.txt` / `.html`。
6. `技艺规则` 固定作为结构字段写成 `[技艺规则]：`；不得写成 `[技艺规则]`、`【技艺规则】` 或 `【技艺规则】：`。学习难度、习得难度、传授难度等内容写在该字段后或后续连续行中。

### 首部元信息规则
1. 首部人员字段统一只使用：（制作人：xxx）、（调整人：xxx）、（修改人：xxx）、（重置人：xxx）、（审核人：xxx）。
2. `制作者`、`投稿人` 归一为 `制作人`；`文本优化`、`复查人` 等按语义归入 `调整人` 或 `修改人`；`审核`、`审核者` 归一为 `审核人`。
3. `调整人`、`修改人`、`重置人` 不要机械互相替换，应按原稿更新性质保留。原文明确写“重置人”就保留 `重置人`；明确写“修改人”就保留 `修改人`；明确写“调整人”就保留 `调整人`。
4. 首部统一采用“人员信息块在前、更新记录块在后”：标题与分类/副标题之后，先按原稿顺序连续保留制作人、调整人、修改人、重置人、审核人；人员块结束后空一行，再连续写更新记录。人员与日志不要求逐条绑定，不要为了配对而重排、合并、去重或伪造人员信息。更新记录块内部必须连续，不得插入空行、人员字段或正文；多条记录按日期从早到晚排列，最早在上、最新在下；同一天的多条记录应在不丢失信息的前提下尽可能合并，用分号分隔；无法可靠判断日期的记录标记待复核并放在明确日期记录之后，不得伪造日期。
5. 投稿没有写更新说明、但实际内容发生变化时，应对照更新前正式资源与新稿，识别真实变动并补写 `【日期：说明】`。说明应覆盖实际修改的能力、数值方向、机制或文本范围；新增资源可写 `【日期：新增收录】`。无法可靠判断时保留原文并列入人工复核，不得编造更新内容。
6. 不要把更新说明、调整内容、重置内容等正文说明误改成人员字段；人员字段只放真实人名或 ID。

## Wiki（Fandom）

### 概述
项目在 Fandom 上维护了一个 Wiki：**macro-realm.fandom.com**（中文路径 `/zh/`）

### 本地文件
Wiki 相关文件存放在 `wiki/` 目录：
```
wiki/
├── sync_to_wiki.py          # Fandom Wiki 同步脚本（全量清理+重建）
├── xlsx_to_wikitext.py       # Excel 转 wikitext 工具
├── wiki_mainpage.wikitext    # Wiki 主页源码
├── 玩家名人堂.wikitext       # 界域玩家名人堂页面
└── *.xlsx / *.wikitext       # 其他待同步的资源页面
```

### 同步命令
```bash
# 同步全部（序列库 + 荣誉室）到 Fandom
python wiki/sync_to_wiki.py --user BotName --password xxx

# 只同步序列库（跳过荣誉室）
python wiki/sync_to_wiki.py --user BotName --password xxx --skip-honor

# 试运行（不推送，本地预览）
python wiki/sync_to_wiki.py --user BotName --password xxx --dry-run

# 只同步某个子目录
python wiki/sync_to_wiki.py --user BotName --password xxx --filter 职业/战技侧
```

### 同步依赖
- **mwclient** — MediaWiki API 客户端（`pip install mwclient`）
- **pandoc** — docx 转换（已有）

### Wiki 页面格式
- 使用 **MediaWiki wikitext** 语法
- 手动维护的页面（主页、名人堂、管理组等）：直接编辑 `.wikitext` 文件
- 自动同步的页面（序列库资源）：由 `sync_to_wiki.py` 从 txt/docx 自动转换
- 自动同步的分类标签：职业、战技侧、神秘侧、科技侧、特殊侧、技能表、能量池、公共建筑、特质改造、荣誉室等
- **注意**：手动创建的页面不要加入自动同步分类，否则会被清理脚本误删

## 已知问题
- `chmcmd`（Linux/Free Pascal）编译的 CHM 索引功能不正常（CJK 编码问题），CI 已改用 Windows + hhc.exe
- `.doc` 格式（非 `.docx`）pandoc 可能无法转换，会生成占位页面
- ~~GitHub Actions 安装 HTML Help Workshop 可能因下载源不稳定而失败~~（已解决：hhc.exe 内置于 `tools/hhw/`）

## 资源重排工具

下架资源移入 `荣誉室/`、重置资源追加回 `序列库/`、或任何会改变编号顺序的批量操作后，使用 `tools/renumber_resources.py` 统一重排编号。

推荐流程：
```bash
# 先试运行，只查看计划，不改文件
python tools/renumber_resources.py

# 确认计划无误后再落盘
python tools/renumber_resources.py --apply
```

工具规则：
- 默认处理 `序列库/` 与 `荣誉室/`。
- 每个目录独立排序，编号文件夹与编号文件一起参与排序。
- `.txt` 文件首行会同步为最终文件名（不含扩展名）。
- 默认是 dry run，只有加 `--apply` 才会实际改名和更新首行。
- 详细中文说明见 `docs/renumber_resources.md`。
