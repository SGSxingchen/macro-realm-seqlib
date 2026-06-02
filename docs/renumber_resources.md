# 资源序号重排工具

工具路径：`tools/renumber_resources.py`

这个工具用于一键重排 `序列库/` 和 `荣誉室/` 下资源的编号，并把 `.txt` 资源首行同步替换为最终文件名。

## 使用方式

默认只预览，不修改文件：

```bash
python tools/renumber_resources.py
```

确认计划无误后执行：

```bash
python tools/renumber_resources.py --apply
```

只处理某个目录：

```bash
python tools/renumber_resources.py --root 序列库/特质改造/生化改造类
```

只重排序号，不更新 `.txt` 首行：

```bash
python tools/renumber_resources.py --skip-title-update
```

## 重排规则

1. 每个目录独立重排，不跨目录合并排序。
2. 同一目录下，编号文件夹和编号文件一起参与排序。
3. 只处理形如 `001】名称` 的编号项目。
4. 文件会保留原扩展名，文件夹不带扩展名。
5. 新编号统一为三位数：`001】`、`002】`、`003】`。
6. `.txt` 文件的首行会替换为最终文件名去掉扩展名后的内容。

## 安全设计

- 默认是 dry run，只打印计划。
- 必须显式加 `--apply` 才会改动文件。
- 实际重命名会先使用临时名中转，避免链式改名互相覆盖。
- 处理顺序从深层目录到浅层目录，保证编号文件夹内部资源先稳定，再重命名外层文件夹。

## 编码说明

源资源 `.txt` 按投稿规范保持 UTF-8。工具读取时会兼容尝试 UTF-8、UTF-8-SIG、GBK、GB2312、Big5；写回时统一保存为 UTF-8。
