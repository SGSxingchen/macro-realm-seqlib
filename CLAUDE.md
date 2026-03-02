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
├── 序列库/               # 当前版本在用的资源（~820个文件）
│   ├── 公共建筑/
│   ├── 技能表/
│   ├── 能量池/
│   ├── 特质改造/
│   └── 职业/
├── 荣誉室/               # 已下架/归档的历史资源（~240个文件）
├── 6.2序列库编者注.txt
├── V6.2序列库更新日志.docx
└── 第六批下架名单.txt
```

## 构建系统

### 构建命令
```bash
# 完整构建（CHM + ZIP）
python build_chm.py --version v6.2

# 只构建 ZIP
python build_chm.py --version v6.2 --skip-chm

# 只构建 CHM
python build_chm.py --version v6.2 --skip-zip
```

### 输出文件命名格式
`宏观界域强化序列库V{版本号}.chm` / `.zip`
例：`宏观界域强化序列库V6.2.chm`

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
git tag v6.2
git push origin main --tags
# GitHub Actions 自动构建并创建 Release
```

## 内容规范（编者注）
1. 资源文件名和文件内部不要添加特殊字符
2. txt 首行为标题，标题后空一行
3. txt 保存为 UTF-8 编码
4. 支持 txt 以外的文件，但需另存为 html 格式
5. 文件编号格式：`001】名称`，用于排序

## 已知问题
- `chmcmd`（Linux/Free Pascal）编译的 CHM 索引功能不正常（CJK 编码问题），CI 已改用 Windows + hhc.exe
- `.doc` 格式（非 `.docx`）pandoc 可能无法转换，会生成占位页面
- ~~GitHub Actions 安装 HTML Help Workshop 可能因下载源不稳定而失败~~（已解决：hhc.exe 内置于 `tools/hhw/`）
