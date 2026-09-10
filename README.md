# Sumsub PDF → Google Drive → Excel

一个本地、可断点续跑的 Python 自动化工具：从 Excel 读取 Sumsub report 链接，在已登录的浏览器中下载 PDF，上传至指定 Google Drive 文件夹，并将查看链接回填至新的 Excel 文件。

> **安全提示**：KYC / screening report 常含敏感个人资料。建议使用私有仓库；不要提交 Excel、PDF、`credentials.json`、`token.json`、浏览器 profile 或 manifest。

## 流程

```text
Excel (sumsub ID + Sumsub_Url)
  → Playwright 打开 Sumsub 并点 Get PDF
  → 本地 PDF
  → Google Drive API 上传
  → 输出 Excel 写入 Google Drive URL
```

## 功能

- 根据 `sumsub ID` 去重，同一 ID 只下载和上传一次。
- 使用 JSON manifest 保存进度，重跑时跳过已成功上传的项目。
- 识别 Sumsub 404 `Applicant not found`，在链接列写入 `not found`。
- 输出 Excel 保留原始内容，并添加 `PDF Processing Status (interim)` 状态列。
- Google OAuth 登录不保存密码；仅在本机保存可撤销的 token。

## Excel 格式

第一行应包含以下列。空格、下划线和大小写不敏感。

| 逻辑字段 | 可用列名 |
| --- | --- |
| Sumsub ID | `sumsub ID` |
| Sumsub report 链接 | `Sumsub_Url` |
| Google Drive 输出链接 | `Google Drive PDF URL`、`Google_Drive_Url` 或 `Google Drive URL` |

如果 `Sumsub_Url` 使用 Excel 公式，请先在 Microsoft Excel 中打开、等待计算完成、保存并关闭，再运行脚本。脚本读取 Excel 保存的计算结果，不会计算 Excel 公式。

## 安装

需要 Python 3.10+、可以登录 Sumsub 的账号、以及可向目标 Google Drive 文件夹上传的 Google 账号。

```bash
git clone https://github.com/YOUR-ORG/sumsub-drive-excel-automation.git
cd sumsub-drive-excel-automation
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python -m playwright install chromium
```

详细 Google 授权步骤见 [docs/GOOGLE_OAUTH.md](docs/GOOGLE_OAUTH.md)。

## 第一次运行：测试两条

不要把真实路径或 folder ID 写进仓库。以下仅为示例：

```bash
python -u src/sumsub_drive_excel.py \
  --input "/absolute/path/source.xlsx" \
  --output "/absolute/path/source_OUTPUT.xlsx" \
  --sheet "Sheet name" \
  --folder-id "YOUR_GOOGLE_DRIVE_FOLDER_ID" \
  --credentials "./credentials.json" \
  --manifest "./state/batch_manifest.json" \
  --downloads-dir "./state/pdfs" \
  --max-records 2
```

首次运行会打开两个授权/登录流程：Google OAuth 和 Sumsub Chromium。完成登录或 MFA 后保持浏览器打开。验证两份 PDF、Drive 链接和输出 Excel 后，删去 `--max-records 2` 执行全量任务。

## 重跑、断网与重复上传

每次成功下载和上传后，脚本都会原子写入 manifest，并保存输出 Excel。相同参数重跑会：

- 跳过已有 Drive 链接、已有 Drive file ID 和 `not found`；
- 继续处理失败或未开始的条目；
- 不重复上传已被 manifest 记录的文件。

极少数情况下，如果连接恰好在 Drive 上传完成但 manifest 尚未保存时中断，下一次可能产生一个重复 PDF。保留 manifest 并避免在上传提示出现时强制中断，可将该风险降到最低。

## 资料

- [架构与所用知识](docs/ARCHITECTURE.md)
- [Google OAuth 设置](docs/GOOGLE_OAUTH.md)
- [日常操作与批次管理](docs/OPERATIONS.md)
- [故障排查](docs/TROUBLESHOOTING.md)
- [安全与数据处理](SECURITY.md)

## 许可证

此仓库未附带默认开源许可证。上传前请根据公司政策选择私有仓库或添加经批准的许可证。
