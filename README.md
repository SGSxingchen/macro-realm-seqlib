# 宏观界域强化序列库

[![Latest Release](https://img.shields.io/github/v/release/SGSxingchen/macro-realm-seqlib?label=%E6%9C%80%E6%96%B0%E7%89%88%E6%9C%AC&color=blue)](https://github.com/SGSxingchen/macro-realm-seqlib/releases/latest)

TRPG「宏观界域」的强化序列资料库，收录职业、特质改造、技能表、能量池、公共建筑等游戏资源，当前版本 **V6.5**，共 **783** 个资源文件。

## 下载使用

前往 [GitHub Releases](https://github.com/SGSxingchen/macro-realm-seqlib/releases/latest) 下载最新版本：

| 文件 | 说明 |
|------|------|
| `宏观界域强化序列库V6.5.chm` | Windows CHM 帮助文档，支持全文检索，推荐使用 |
| `宏观界域强化序列库V6.5.zip` | 完整资源包（含序列库 + 荣誉室 + 更新日志） |

> CHM 文件如果打开后显示空白，右键文件 → 属性 → 勾选「解除锁定」后重新打开。

## 在线查阅

**序列库查询网站**：https://trpg.chordvers.org/

**界域 Wiki**：https://macro-realm.fandom.com/zh/wiki/宏观界域TRPG_Wiki

可在线查询所有资源内容，欢迎投稿。

## 目录结构

```
序列库/                     当前版本在用的资源（783个文件）
├── 公共建筑/               公共设施（13个）
├── 技能表/                 技能表资源（229个）
│   ├── 其他及特殊/
│   ├── 战技侧/
│   ├── 特殊侧/
│   ├── 神秘侧/
│   └── 科技侧/
├── 新过审序列/             新过审资源说明（1个）
├── 能量池/                 能量池资源（24个）
├── 特质改造/               特质改造资源（342个）
│   ├── 异化改造类/
│   ├── 特化改造类/
│   ├── 生化改造类/
│   └── 特殊特质/
└── 职业/                   职业资源（174个）
    ├── 战技侧/
    ├── 特殊侧/
    ├── 神秘侧/
    └── 科技侧/

荣誉室/                     已下架/归档的历史资源（249个文件）
```

## 投稿规范

1. 文件名和文件内部**不要添加特殊字符**
2. txt 首行为标题，尽量与文件名一致，标题后**空一行**
3. txt 保存编码为 **UTF-8**
4. 支持 txt 以外的文件，但须另存为 **html 格式**（源文件也要保留）
5. 文件编号格式：`001】名称`，编号用于排序
6. 投稿修改请在**最新版本**的基础上进行

## 本地构建

### 依赖

- **Python 3**
- **pandoc** — docx/doc 转 HTML（`winget install JohnMacFarlane.Pandoc`）
- **hhc.exe** — CHM 编译器（已内置于 `tools/hhw/`，无需额外安装）

### 构建命令

```bash
# 完整构建（CHM + ZIP）
python build_chm.py --version v6.5

# 只构建 ZIP
python build_chm.py --version v6.5 --skip-chm

# 只构建 CHM
python build_chm.py --version v6.5 --skip-zip
```

输出文件：`宏观界域强化序列库V6.5.chm` / `.zip`

### CI/CD

推送 `v*` tag 到 GitHub 会自动触发 Actions 构建并创建 Release：

```bash
git tag v6.5
git push origin main --tags
```

## 联系方式

- 制作人：**沧羽**
- QQ：853304398
- 发现错漏或有投稿意向请直接联系
