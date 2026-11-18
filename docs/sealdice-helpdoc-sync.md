# 海豹帮助文档同步

## 同步范围

只从 `序列库/` 的五类正式资源构建海豹 JSON：公共建筑、技能表、能量池、特质改造、职业。不发布 `新过审序列/`、`荣誉室/` 或非 TXT 资源。`.tar.gz` 是 GitHub Actions 构建产物和上传用的传输包，不放进海豹的 `data/helpdoc`；同步后该目录保存解包后的正式分类和清单。同步不覆盖其他帮助文档分组；`--clean-legacy-zip` 只清理 HelpDoc 根目录的旧包 `宏观界域.zip`。

## GitHub Actions

[同步海豹帮助文档](../.github/workflows/sync-sealdice-helpdoc.yml) 在 `main` 推送涉及正式资源、相关脚本或工作流时触发（仅修改 `新过审序列/` 不触发），推送 `v*` 标签也会触发。手动运行可勾选 `force_reload`（默认关闭）。同一时间只运行一次同步，新任务不会取消正在运行的任务。

复用部署 Secrets：`DEPLOY_HOST`、`DEPLOY_USER`、`DEPLOY_PORT`（默认 `61149`）、`DEPLOY_SSH_KEY`。仓库变量 `SEALDICE_DIR` 未配置时使用 `/root/sealdice/15`。部署账号需能写入实例的 `data/helpdoc`、`backups` 和上传暂存目录 `/tmp/sealdice-helpdoc-deploy/`，并能访问服务器本机的海豹 API。工作流会在成功或失败后尝试清理本次上传内容；若连接中断导致目录残留，应先确认没有在跑的任务，再检查并清理该目录。SSH 使用 `ssh-keyscan` 临时取得主机公钥，与现有 Web 部署一致；这不是预先固定公钥指纹。

构建前先运行全库 `--all --strict`。当前已知全库错误来自不发布的 `新过审序列/【新过审序列】简介.txt`（零卡片）；如果全库校验未通过，工作流会发出警告，再用同一套 strict 规则检查五个发布目录，正式范围有错误就停止。构建成功后，Actions Summary 显示各类资源和词条数，Artifacts 提供 `sealdice-helpdoc.tar.gz`。

## 本地构建与服务器端同步

从仓库根目录构建和跑测试，无需连接服务器：

```bash
python3 tools/build_sealdice_helpdoc.py --check \
  --output-dir /tmp/sealdice-helpdoc-out \
  --archive /tmp/sealdice-helpdoc.tar.gz
python3 -m unittest discover -s tests -p 'test_*helpdoc*.py'
```

需要手动同步时，把 `tools/sync_sealdice_helpdoc.py` 和构建包放在服务器的非 HelpDoc 暂存目录，在服务器上运行：

```bash
python3 /tmp/sealdice-helpdoc-deploy/sync_sealdice_helpdoc.py \
  --bundle /tmp/sealdice-helpdoc-deploy/sealdice-helpdoc.tar.gz \
  --sealdice-dir /root/sealdice/15 --clean-legacy-zip
```

如果服务器上已有仓库源码，也可直接用 `--source-dir /path/to/序列库` 代替 `--bundle`，现场构建后同步。`--helpdoc-dir` 可覆盖默认路径 `<sealdice-dir>/data/helpdoc`。默认 API 地址为 `http://127.0.0.1:3211`，可用 `--api-url` 覆盖。API token 按 `--token` → `SEALDICE_API_TOKEN` → `<sealdice-dir>/data/dice.yaml` 的第一个 `accessTokens` 条目取用。避免在终端历史和日志中记录 token。

## 验收、跳过与回滚

- 已安装的清单通过校验且对应同一 Git 提交时，默认跳过同步和重载。需要强制更新时加 `--force`，或在 Actions 手动勾选 `force_reload`。`--skip-reload` 仅写文件，不会重载或验证索引，不代表服务已生效；之后用同一提交正式重载时须加 `--force`，否则会跳过。
- 正常同步通过 `POST /sd-api/helpdoc/reload` 重载，等待 `GET /sd-api/helpdoc/tree` 不再加载并检查五个分类的状态，然后用 `POST /sd-api/helpdoc/text_item_page` 确认各分类的搜索结果 `total > 0`。
- 上一份受管内容保留在 `<sealdice-dir>/backups/helpdoc-seqlib-prev`，只有一个备份槽；它也包含本次移除的旧 ZIP（如有）。如重载或验收失败，脚本会尝试恢复旧内容和 ZIP 并再次重载。如果自动恢复失败，先停止后续同步并检查日志及备份，只恢复受管分类、清单和本次移除的旧 ZIP，再重载和复核；不要整目录覆盖 `data/helpdoc`。
