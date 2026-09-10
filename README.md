# Sumsub_Drive_Excel
Sumsub PDF → Google Drive → Excel
这个脚本会按 Excel 的每一行执行以下流程：

1. 从 `Sumsub_Url` 打开 Sumsub report，点击 **Get PDF**；
2. 下载 PDF 到本地；
3. 上传至指定 Google Drive 文件夹；
4. 把对应的 Google Drive 查看链接写入 `Google Drive PDF URL`；
5. 在 `PDF Processing Status (interim)` 记录执行结果，便于检查或重跑。

它会按 **sumsub ID** 去重。相同 ID 出现多行时，只下载/上传一次，之后将同一个链接回填至全部重复行。

## 先决条件

- macOS，Python 3.10 或更新版本；
- 可以在浏览器登录 Sumsub，并有下载 report PDF 的权限；
- 有权向目标 Google Drive 文件夹上传；
- Excel 的第 1 行必须有以下列名（空格、下划线和大小写可不同）：
  - `sumsub ID`
  - `Sumsub_Url`
  - `Google Drive PDF URL`

脚本不要求也不会保存你的 Sumsub 密码。Sumsub 登录状态保存在本机的 `sumsub_browser_profile` 文件夹中，Google 的 OAuth 授权 token 保存在 `token.json`。两者都不要上传到网盘或 Git。

## 1. 安装依赖（首次一次）

在 Terminal 中进入此脚本文件夹，然后运行：

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python -m playwright install chromium
```

如果你的电脑没有 `python3`，请先安装 Python 3.10+，然后重新打开 Terminal。

## 2. 创建 Google Drive 授权文件（首次一次）

1. 打开 [Google Cloud Console](https://console.cloud.google.com/)；创建或选择一个项目。
2. 在 **APIs & Services → Library** 启用 **Google Drive API**。
3. 在 **APIs & Services → OAuth consent screen** 完成基本设置，并把自己的 Google 账号加入 Test users（若应用仍为 Testing）。
4. 在 **Credentials → Create Credentials → OAuth client ID** 中选择 **Desktop app**。
5. 下载 JSON，放入本脚本文件夹，命名为 `credentials.json`。

第一次脚本运行时会跳转至 Google 的官方授权页面；选择有该 Drive 文件夹权限的账号并授权即可。之后会自动复用 `token.json`。

## 3. 先跑两个做测试

下列示例使用你之前的 Google Drive 文件夹 ID。请将 Excel 路径替换成自己当前文件的完整路径：

```bash
python sumsub_drive_excel.py \
  --input "/Users/admin/Downloads/UP Business RFI.xlsx" \
  --output "/Users/admin/Downloads/UP Business RFI_已填充Drive链接.xlsx" \
  --sheet "UP Business RFI" \
  --folder-id "1JCZgcTyZ65klMZhxxcOLt3VraWe8Mb0A" \
  --credentials "./credentials.json" \
  --max-records 2
```

第一次访问 Sumsub 时，Chromium 会打开。请在该浏览器窗口完成 Sumsub 登录、SSO 或 MFA，然后保持窗口打开。脚本会等待页面出现 **Get PDF** 并继续处理。

确认 Excel 与 Drive 的两个结果都正确后，删除 `--max-records 2`，用同一命令跑完整批次。

## 4. 只上传你已经下载好的 PDF

把已经下载的 PDF 放进 `pdf_downloads`。文件名可使用以下任一形式：

```text
<sumsub ID>.pdf
applicant-summary-<sumsub ID>.pdf
```

脚本会寻找文件名中包含对应 sumsub ID 的 PDF；若有多个匹配文件，会优先使用最新下载的一个。

再使用：

```bash
python sumsub_drive_excel.py \
  --input "/Users/admin/Downloads/UP Business RFI.xlsx" \
  --output "/Users/admin/Downloads/UP Business RFI_已填充Drive链接.xlsx" \
  --sheet "UP Business RFI" \
  --folder-id "1JCZgcTyZ65klMZhxxcOLt3VraWe8Mb0A" \
  --credentials "./credentials.json" \
  --skip-download
```

## 中断、失败和重跑

- 每成功下载或上传一个文件，脚本都会保存 `sumsub_drive_manifest.json` 和输出 Excel。
- 中途按 `Ctrl+C`，或某一行失败，都可以直接运行**完全相同的命令**继续；已上传的 PDF 不会重复上传。
- 打不开 report、没有权限、或找不到 **Get PDF** 时，脚本会继续处理下一条，并在 Excel 状态列写 `ERROR: ...`。请人工确认后再重跑。
- 如果某行的 `Google Drive PDF URL` 已经有链接，脚本默认保留该链接，不再下载或上传。

## 链接权限：默认更安全

默认情况下，脚本只上传文件，**不把 PDF 公开给所有人**。文件会继承目标文件夹的共享策略。这通常更适合含个人资料的 KYC/report PDF。

只有在你明确已获授权且确实要生成公开链接时，才增加：

```text
--share-anyone-with-link
```

该选项会对每个新上传 PDF 设置 “Anyone with the link / Viewer”。如果 Workspace 管理员禁止公开分享，Google 会拒绝该权限操作并把错误写入 Excel 状态列。

## 常见问题

**脚本说找不到 Get PDF**：先确认打开的链接是有效的 Sumsub report；在脚本打开的 Chromium 中手动登录；再执行同一命令。不要用 `--headless` 做首次登录。

**Google 说 app is not verified / access blocked**：通常是 OAuth consent screen 尚未完成，或者当前 Google 账号未列入 Test users。回到第 2 步设置。

**Excel 打不开或格式变化**：请始终使用不同的 `--output` 文件，不要覆盖原始 Excel。此脚本适用于 `.xlsx` / `.xlsm`；复杂的原生 Excel 功能（例如某些嵌入对象）可能无法完整保留。

**要运行很大一批数据**：建议先用 `--max-records 2` 验证权限和列名，然后全量运行。处理过程中不要关闭脚本打开的 Sumsub 浏览器。
