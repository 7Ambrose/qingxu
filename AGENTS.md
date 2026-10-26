# 青序公开版 · native 分支

- 本分支只含 C# / .NET 10 原生版（`Qingxu/`、`Qingxu.Tests/`）；PowerShell 版在 `main` 分支，两边不混放。
- 只提交程序代码、合成测试和从零生成的模板。不复制私有仓库的 Git 历史、个人工作簿、课表、缓存、备份、账号路径或凭据。
- 两套实现共用同一份 Excel 格式和同一个单实例锁名 `Local\CodexTodayAgenda_27107`，改动时要保持一致，否则会互相覆盖。
- 状态列约定：`☐ 待完成` / `☑ 已完成` / `☒ 未完成`。`done` 与 `incomplete` 落在同一列，最终值由两者共同决定。
- 写入必须保留：版本号 + SHA-256 双校验 → 备份 → 原子替换 → 回执。不要改成直接覆盖。
- `scripts/build_native.ps1` 是唯一构建入口；发布包只走 `-Publish`，不手动拼装。
- 提交前跑 `scripts/build_native.ps1 -Test`（43 项断言）。测试只操作 QA 副本，绝不写真实工作簿。
- 不自动更新发布包，只有明确要求发布 / Release 时才重新构建。
- 中文脚本使用 UTF-8 BOM，否则 Windows PowerShell 5.1 会按 ANSI 读取导致乱码。
