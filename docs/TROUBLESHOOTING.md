# 故障排查

| 现象 | 原因与处理 |
| --- | --- |
| `Worksheet not found` | 使用实际工作表标签，而不是 Excel 文件名。 |
| `Missing required Excel column` | 检查三列列头是否存在。 |
| `Applicant not found` | Sumsub URL 无效/对象已删除；脚本写入 `not found` 并跳过。 |
| 找不到 `Get PDF` | 先确认在脚本打开的 Chromium 中已登录，并人工验证 report 可打开。 |
| Google 授权被阻止 | 检查 Drive API、OAuth consent screen 和 Test users。 |
| `credentials.json` 找不到 | 把 Desktop app 下载的 JSON 放到传入 `--credentials` 的路径。 |
| 中途中断 | 使用完全相同参数重跑；保留 manifest。 |

对于非 404 的下载失败，脚本将状态写入输出 Excel 后继续下一行。不要把临时网络、登录过期等问题直接判为 `not found`；先人工确认。
