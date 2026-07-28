# AutoInvoice

Windows 桌面发票报销助手。

## 使用

1. 把待报销 PDF 或 OFD 发票放到 `invoices` 目录。
2. 运行 `AutoInvoice.exe`，或开发时运行 `python autoinvoice.py`。
3. 输入报销金额，选择发票类型或保留“全部”，点击“计算组合”。
4. 点击“生成组合预览”，程序会先生成一个合并 PDF，并在右侧预览。
5. 在预览区点击“打印”，会弹出 Windows 打印设置窗口，可在里面选择打印机、份数、单双面、纸张等设置。
6. 打印后确认标记已使用，程序会把发票移动到 `used_invoices`，并写入 `data\invoice_records.sqlite3` 和 `data\used_records.json`。

组合规则：合计金额必须大于等于报销金额；在最少超额金额的 50 元范围内，优先选择开票日期更早的发票。

识别不到金额的 PDF/OFD 会被移动到 `failed_invoices`，并写入 `data\failed_records.json` 供人工处理。

扫描时会按发票号码去重：如果 `invoices` 中同时存在相同发票号码的 PDF/OFD 或不同文件名但发票号码重复，程序只保留一份可用发票，优先保留 PDF，其余重复文件移动到 `void_invoices`。

运行目录结构：

```text
AutoInvoice.exe
invoices
print_jobs
used_invoices
void_invoices
data
failed_invoices
```

配置文件、SQLite 数据库和运行记录都保存在 `data` 目录。旧版本留在程序目录下的 `app_settings.json`、`used_records.json`、`failed_records.json` 和 `invoice_records.sqlite3` 会在启动时自动迁移到 `data`。

## 打包

在 PowerShell 中执行：

```powershell
.\build.ps1
```

打包产物在 `dist\AutoInvoice.exe`，脚本也会复制一份到项目根目录的 `AutoInvoice.exe`。exe 会扫描自己所在目录下的 `invoices`，所以运行时需要让 exe 和 `invoices` 同级。
