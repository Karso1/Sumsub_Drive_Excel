# 架构与知识说明

## 这是哪类技术？

这是一个 **RPA / Web Automation / ETL 风格的数据流水线**，不是计算机视觉（CV）也不是 AI。它依靠网页的 DOM、按钮角色和文字 `Get PDF` 定位操作，而非截图识别或鼠标坐标。

| 层 | 技术 | 作用 |
| --- | --- | --- |
| 输入/输出 | `openpyxl` | 读取 Excel，写入 Drive URL 和状态 |
| 浏览器 | Playwright | 打开 Sumsub report、等待 `Get PDF`、捕获浏览器下载事件 |
| 云存储 | Google Drive API | 创建 PDF 文件、取得 file ID 和查看 URL |
| 身份授权 | OAuth 2.0 | 用户在官方 Google 页面授权，本地缓存 refresh token |
| 可靠性 | JSON manifest + 原子写入 | 为每个 ID 存储下载路径、Drive file ID、URL、状态 |
| 操作界面 | shell / `.command` launcher | 把固定批次参数封装为可双击运行的入口 |

## 数据流与状态机

每个 Sumsub ID 的 manifest 记录可包含 `pdf_path`、`drive_file_id`、`url`、`status`、`error` 和 `updated_at`。

```text
new → downloaded → uploaded
                  ↘ error → retry
new → not_found (终态，后续跳过)
```

脚本在下载成功后保存 manifest；上传成功后再次保存。因此重跑时优先使用已有 Drive file ID 或 URL。这个模式常称为 **idempotency（幂等性）**：重复执行大多数情况下不会创建第二份结果。

## 重要实现选择

- **Persistent browser profile**：把 Sumsub session 存在本地 profile，便于登录一次后复用；它被 `.gitignore` 排除。
- **`drive.file` OAuth scope**：脚本只创建/管理它创建的 Drive 文件，遵循最小权限原则。
- **headful first run**：首次不使用 headless，避免 SSO/MFA 无法完成。
- **404 快速处理**：Sumsub 若 client-side 跳转至 `/404/`，写 `not found` 而非等待下载超时。

## 不包含什么

- 不读取或保存 Sumsub 密码。
- 不绕过 MFA、CAPTCHA、访问控制或 Drive 管理策略。
- 不使用 OCR、视觉识别、机器学习模型或远程服务器。
