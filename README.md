# 青序 · C# 原生版

Windows 桌面日程小窗，单文件免安装。用 .NET 10 / WPF 写成，**运行时不启动 Python、Node 或 PowerShell**，也不需要 `@oai/artifact-tool` 或其他第三方包。

本仓库用分支区分两套实现：

| 分支 | 实现 | 运行依赖 |
|---|---|---|
| `main` | 原有 PowerShell + WPF 版 | Windows PowerShell 5.1、Python 3、Node.js（写回 Excel 需要 `@oai/artifact-tool`） |
| **`native`（当前分支）** | C# / .NET 10 原生版 | 无，自包含单文件 |

两套实现读写**同一份 Excel**、共用同一套单元格约定和同一个单实例锁，所以不能同时开两个窗口编辑同一个工作簿。日常使用选一套即可。

## 下载与运行

1. 到 [Releases](https://github.com/7Ambrose/qingxu/releases) 下载 `Qingxu-Native-v0.1.0-win-x64.zip`，完整解压。
2. 把 `计划模板.xlsx` 复制到解压目录下的 `outputs\study-plan\计划.xlsx`，再填写自己的课程和任务。
3. 双击 `青序.exe`。首次启动会自动在上级目录寻找 `outputs\study-plan\计划.xlsx`。
4. 想放桌面：双击 `创建桌面快捷方式（原生版）.ps1`，或右键 `青序.exe` → 发送到 → 桌面快捷方式。

也可以用命令行指定工作簿：

```
青序.exe --book "D:\MyPlanner\计划.xlsx"
```

## 功能

**日程**

- 今天 / 明天 / 后天三页独立翻页，另有每周清单页
- 顶部进度条：`今日 3/7 已完成 · 还剩 4 项`（课程不计入待办）
- 底部快速添加框：输入任务名回车即建，不必先点新建
- 键盘快捷键：`←` `→` 翻日期，`T` 回今天，`N` 新建，`R` 重读 Excel，`1`–`4` 切页，`Esc` 关闭弹层

**任务状态**

- 单击切换完成（浅绿底 + 勾 + 删除线）
- 双击标记未完成（灰底 + 灰色斜线），与已完成互斥
- 右键菜单：标记未完成 / 标记已完成 / 清除状态 / 复制到明天
- `★` 强调：黄色底色 + 加粗，颜色沿用原版的 `#FFF7DF`

**排期与编辑**

- 拖动 `⠿` 在同一天待办或同周清单内排序
- 拖到时间点排进半小时网格，自动检查与课程冲突
- 周任务可拖到所属周的其他日期；日程页「撤」只取消安排，本周页「删」才删任务

**同步**

- 修改先写本地 `.qingxu.json`，点「同步并打开Excel」才写回工作簿
- 写回前校验版本号 + SHA-256，不一致就停下并要求逐项核对，不静默覆盖
- 原文件先备份成 `*.qingxu-backup-*.xlsx`，原子替换
- Excel 被别处修改后 2 秒内自动重读；窗口跨过午夜自动切到当天

## 数据兼容

- 适配 2026-10-18 至 2027-01-17 的固定模板；课程、任务的行列位置按约定，不是任意 Excel 都能导入。
- 只修改涉及的单元格，保留包内其他压缩条目、样式、公式和关系，不重新生成整份工作簿。
- 公式交给 Excel 下次打开时重算：**这不是公式计算引擎**。
- 首次运行会读取旧版 `.agenda.json` 里尚未同步的修改，存到独立的 `.qingxu.json`，不改动旧缓存。

## 构建

需要 [.NET 10 SDK](https://dotnet.microsoft.com/download)。

```powershell
# 只构建 + 跑回归（需要一份已填写的工作簿，空白模板不含断言所需数据）
powershell -NoProfile -File scripts/build_native.ps1 -Test

# 构建自包含单文件并打包
powershell -NoProfile -File scripts/build_native.ps1 -Publish -Version 0.1.0
```

回归测试共 43 项断言，覆盖工作簿保真（其他 zip 条目哈希不变）、日志恢复、周任务取消安排、日期边界、并发锁、冲突核对、旧版队列导入和界面。断言里引用了具体的课程行和周任务行，所以要拿一份已填好内容的工作簿跑；只用空白模板复现不了。

诊断：`青序.exe --book "副本路径" --smoke-test "报告.json"` 会保存界面截图，暖机后采样 10 秒 CPU / 内存并自动退出。只用于副本。

## 体积与性能

自包含单文件约 165 MiB，主要是 .NET 运行时，因此启动约 1 秒、工作集约 200 MiB。要更小的体积可以去掉 `-p:SelfContained=true`，改为依赖预装 .NET 10，但那就不再是"下载即用"。

## 隐私

- 只公开程序、合成测试和从零生成的模板；个人课表、任务、缓存、备份、日志一律不上传。
- 不在仓库里存放个人路径、账号或凭据。
- `.gitignore` 不能消除已经提交的历史，分享前请自行检查。

## 参考

实现时查阅过 [微软 WPF 示例](https://github.com/microsoft/WPF-Samples) 与 [.NET 单文件发布文档](https://learn.microsoft.com/dotnet/core/deploying/single-file/overview)，未复制其他待办项目的业务代码。
