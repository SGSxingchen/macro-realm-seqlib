# Wiki 同步 Action 设计

## 目标

在 GitHub Actions 中新增一个独立的“Wiki 同步”手动按钮，用于运行 `wiki/sync_to_wiki.py`。按钮支持 dry-run 预览和真实同步，默认 dry-run，并默认按最近 tag 到当前 HEAD 的差异做增量同步，让删除和重命名能走脚本的 diff 逻辑。

## 触发方式

工作流仅使用 `workflow_dispatch` 手动触发，不监听 push 或 tag。这样 Wiki 同步与发布构建互不影响，也不会因为普通提交自动写入线上 Wiki。

## 输入项

- `mode`：`dry-run` 或 `sync`，默认 `dry-run`。
- `sync_range`：同步范围，默认 `latest-tag`。可选 `latest-tag`、`custom-ref`、`last-commit`、`full`。
- `diff_from`：`custom-ref` 模式使用的起始 tag/commit。
- `skip_honor`：是否跳过荣誉室，默认 `false`。
- `filter`：可选，只同步某个子目录，例如 `职业/战技侧`。
- `delay`：Fandom API 请求间隔秒数，默认 `5`。

## 执行流程

1. 检出仓库，并使用完整历史和 tag，保证 diff 基准可用。
2. 安装 Python。
3. 安装 pandoc 和 `mwclient`。
4. 组装同步脚本参数。
5. dry-run 模式使用占位账号密码，并追加 `--dry-run`。
6. sync 模式读取 GitHub Secrets：`WIKI_USER`、`WIKI_PASSWORD`。
7. sync 模式缺少 Secrets 时直接失败并提示配置。
8. `latest-tag` 模式通过 `git describe --tags --abbrev=0` 找最近 tag，并传入 `--diff-from`。
9. `custom-ref` 模式使用 `diff_from` 并传入 `--diff-from`。
10. `last-commit` 模式传入 `--incremental`。
11. `full` 模式不传 diff 参数，走全量清理重建。
12. 按输入追加 `--skip-honor`、`--filter`、`--delay`。
13. 运行 `wiki/sync_to_wiki.py`，结果保留在 Actions 日志中。

## 风险控制

- 默认 dry-run。
- 默认 `latest-tag` 增量同步，删除/重命名依赖 `git diff` 判断。
- 真实同步必须显式选择 `sync`。
- 真实同步必须配置 Secrets，不使用占位凭据。
- `filter` 不能搭配 `full` 真实同步，避免清理全部自动分类后只重建部分目录。
- 工作流独立存在，不影响 Release 构建。

## 验证

- 用本地检查确认工作流包含手动触发、输入项、Secrets 校验和脚本参数。
- 可在 GitHub Actions 中先手动运行 dry-run，确认日志后再运行真实同步。
