# 青序 · PowerShell 版

Windows 桌面日程小窗：Excel 读取、三日独立翻页、每周任务、拖拽排期、绿色完成勾选、灰色未完成斜线、黄色强调、托盘恢复及按需同步。

**这是原有 PowerShell + WPF 版，不是 C# 原生实验版。** EXE 只是无控制台启动器；当前发布是需要运行依赖的预览版，不是免安装独立软件。

## 下载与运行

1. 到 [Releases](https://github.com/7Ambrose/qingxu/releases) 下载 `Qingxu-PowerShell-v0.1.0.zip`，完整解压，不要只下载单个 EXE。
2. Windows 10/11、Windows PowerShell 5.1、.NET Framework 4.x、Python 3 和 Node.js 必须可用。可用环境变量 `AGENDA_PYTHON`、`AGENDA_NODE` 指定解释器。
3. 将 `templates/计划模板.xlsx` 复制为 `outputs/study-plan/计划.xlsx`，再填写自己的课程和任务。不要覆盖已有工作簿。
4. 双击 `今日日程/启动青序.exe`。若系统阻止脚本，请先核实来源，再按本机可信脚本策略处理；启动器不自动更改系统执行策略。
5. Excel 保存后会刷新；软件修改先保存在本地 JSON，点击“同步并打开 Excel”写回。真正的冲突会先要求核对，不静默覆盖。

## 重要依赖限制

当前 Excel 写入使用 `@oai/artifact-tool`，**不随此仓库和发布包分发**，也不提供未经授权的下载方式。运行环境必须能够让 Node.js 加载该组件；缺少它时，Excel 写回不可用。Python 读取和本地任务缓存不依赖该组件。

因此本 Release 可用于具备相应运行环境的试用与共同调试，**不能承诺普通电脑下载即用**。无需这些特殊依赖的独立版仍待后续工作。公开源码不改变第三方组件的分发许可。

## 使用自己的现有 Excel

在仓库或解压目录根部新建 **不提交** 的 `qingxu.local.json`：

```json
{"bookPath":"D:\\MyPlanner\\计划.xlsx"}
```

也可设置 `AGENDA_BOOK`，或直接运行主脚本的 `-BookPath` 参数。优先级：脚本参数、环境变量、本地配置、默认路径。Python 工具支持环境变量／本地配置及显式 `--book`。

现版本兼容固定模板：2026-10-18 至 2027-01-17；课程和任务行列保持约定。模板没有真实课程、考试或个人任务；不是任意结构的 Excel 都能读取。

## 隐私与开发

- 只公开程序、合成测试和从零生成的模板；个人 Excel、JSON 队列、联动基线、日志和备份不上传。
- 不在本仓库存放个人路径／账户信息。分享文件前检查内容，`.gitignore` 不能消除已经提交的历史。
- 本仓库从干净快照开始，不含原私有仓库的历史或原生分支。后续修复只需提交这里。
- `scripts/build_release.ps1` 重新编译旧版启动器并以白名单打包，排除本地设置和数据；不会打包原生版。

程序在 PowerShell 中运行，启动器用 C# 编译并不表示产品换成了原生版。
