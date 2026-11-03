# 项目工作约定

## 构建与运行（固定流程）

`D:\nick\project\autoinvoice-ds` 是**源码目录**，只在这里改代码。
程序**不在**本目录运行 —— 本目录被 DSH 工作区沙箱限制（原因见下）。

1. 在 `autoinvoice-ds` 修改 `autoinvoice.py`
2. 在 `autoinvoice-ds` 里构建 exe
3. 把 `AutoInvoice.exe` 复制到 `D:\nick\project\autoinvoice\` 覆盖
4. 在 `D:\nick\project\autoinvoice\` 里运行

```powershell
# 构建（两个前置条件缺一不可）
$env:PATH = "D:\tools\miniconda3;" + $env:PATH                 # 默认 python 是 WindowsApps 占位符
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
cd D:\nick\project\autoinvoice-ds
.\build.ps1 -SkipDependencyInstall

# 部署（先关掉正在运行的实例，否则 exe 被锁、复制失败）
Get-Process AutoInvoice -ErrorAction SilentlyContinue | Stop-Process -Force
Copy-Item .\AutoInvoice.exe D:\nick\project\autoinvoice\AutoInvoice.exe -Force
```

如果构建在 PyInstaller 打出第一行 stderr 时就中断（`build.ps1` 里 `$ErrorActionPreference = "Stop"`
会把原生程序的 stderr 当成错误），改用 cmd 包装，让 stderr 直接落盘而不再转成 ErrorRecord：

```powershell
cmd /c "powershell -NoLogo -NoProfile -ExecutionPolicy Bypass -File D:\nick\project\autoinvoice-ds\build.ps1 -SkipDependencyInstall > build.log 2>&1"
```

验证方式：标题栏显示 `AutoInvoice 发票报销助手 v1.<MMDD>.<HHmm>`。
构建后可用 `AutoInvoice.exe --scan-json` 对 `invoices\` 做一次无界面冒烟测试。

`build.ps1` 会把 `APP_VERSION` 自动改写为构建时刻，这是预期改动，提交时一并带上。

## 为什么不能在本目录运行

本工作区被 DSH 沙箱打了专属 ACL（`S-1-4-*` 授权 + `Everyone` 拒绝「删除子目录及文件」），
镜像位于本目录内的进程只能写本目录。后果：

- 单文件 exe 无法解包到 `%TEMP%` → `Could not create temporary directory!`
- Explorer 启动未签名 exe 时前置「未知发布者」安全警告

同一份 exe 放到 `D:\nick\project\autoinvoice\` 运行完全正常。
**不要修改这些 ACL**，它们是 DSH 的预期设计。

## 代码与提交

- 提交信息遵循 `feat: release AutoInvoice vX.YYYY.MMDD.HHMM`，下空一行写变更说明。
- 产物（`AutoInvoice.exe`、`dist\`、`build\`、`*.spec`）在 `.gitignore` 里，不提交。
- 运行数据（`data\`、`invoices\`、`used_invoices\`、`print_jobs\` 等）不进版本库。
- 全量测试：`python -m unittest discover -s tests`。

## `D:\nick\project\autoinvoice\` 是运行目录，不是代码副本

- 那边**只运行 exe**。它的 `autoinvoice.py`、`tests\`、`README.md` 以及它的 git 状态都**不归本仓库管**。
- 不要读取或修改那边的源码，不要在那里提交，也不要把它当作构建或对账的输入。
- 那边代码如需更新，由用户手动处理；本仓库的职责仅限于把 `AutoInvoice.exe` 复制过去。
