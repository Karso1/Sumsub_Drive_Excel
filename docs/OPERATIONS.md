# 日常操作

## 每个新批次

为每批创建独立目录、独立 Drive 文件夹、独立 manifest 和 PDF 下载目录。不要复用旧批次 manifest。

```text
batches/2026-10/
  source.xlsx
  source_OUTPUT.xlsx
  manifest.json
  pdfs/
```

脚本本身、`.venv`、`credentials.json` 可以复用，只需替换 `--input`、`--output`、`--sheet`、`--folder-id`、`--manifest`、`--downloads-dir`。

## 公开链接

默认上传文件继承目标 Drive 文件夹的权限，适合含个人资料的 report。仅在已获明确授权时使用 `--share-anyone-with-link`，它会为每个新 PDF 添加“持有链接的任何人可查看”。

## 手工下载的 PDF

若已下载 PDF，可将文件放入 `--downloads-dir`，文件名含 Sumsub ID 即可，例如 `<ID>.pdf` 或 `applicant-summary-<ID>.pdf`，然后加 `--skip-download`。

如需用本地 PDF 覆盖输出 Excel 中的 `not found`，请以该输出 Excel 作为 `--input`，并加 `--retry-not-found-with-local-pdfs`。已有 Drive URL 会保留；没有匹配本地 PDF 的记录继续保持 `not found`。
