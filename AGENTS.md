# 青序公开版

- 此目录是正在维护的 PowerShell + WPF 产品；只向本仓库提交程序代码。
- 不复制上层私有仓库的 Git 历史、原生版、个人工作簿、课表、缓存、备份、账号路径或凭据。
- `qingxu.local.json` 只供本机配置，不提交、不打包。
- 发布模板必须从零生成，不能以删除个人文件可见单元格代替完整脱敏。
- 先在临时副本测试；保留双版本／哈希校验、备份及原子替换。
- 中文 PowerShell 5.1 脚本使用 UTF-8 BOM。保留灰色斜线、黄色强调及按需同步。
- 运行 `test_reader.py`、`test_excel_links.py`、`test_sync_review.py`、`test_regressions.ps1`、`test_refresh_races.ps1`、`test_launcher.ps1`。
- 不自动更新发布包，只有明确要求发布／ZIP／Release 时才重新构建。
