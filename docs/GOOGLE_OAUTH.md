# Google Drive OAuth 设置

1. 在 [Google Cloud Console](https://console.cloud.google.com/) 创建或选择项目。
2. 在 **APIs & Services → Library** 启用 **Google Drive API**。
3. 在 **Google Auth Platform** 完成 consent screen 基础配置；若是测试应用，把运行脚本的 Google 账号加入 Test users。
4. 在 **Clients** 创建 OAuth client，类型选择 **Desktop app**。
5. 下载 JSON，保存到项目根目录并命名为 `credentials.json`。
6. 首次脚本运行时，浏览器会打开 Google 官方授权页。选择能写入目标文件夹的账号并允许。

随后脚本创建 `token.json`。若需撤销权限，可在 Google Account 的第三方访问页面撤销，并删除本地 `token.json`。

官方参考：

- https://developers.google.com/workspace/guides/create-credentials
- https://developers.google.com/workspace/drive/api/guides/manage-uploads
