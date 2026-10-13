# AutoInvoice

Windows 桌面发票报销助手。

## 使用

1. 把待报销 PDF、OFD 或图片发票放到 `invoices` 目录。图片支持 JPG/JPEG、PNG、BMP、TIF/TIFF、WebP，一个文件对应一张发票。
2. 运行 `AutoInvoice.exe`，或开发时运行 `python autoinvoice.py`。
3. 输入报销金额，勾选一个或多个发票类型（默认选中“全部”），点击“计算组合”。勾选具体类型后会取消“全部”；取消最后一个具体类型时恢复“全部”。
4. 点击“生成组合预览”，程序会生成合并 PDF 并在右侧预览，同时在 `print_jobs` 下生成同名 Excel 核对表。核对表按 PDF 顺序列出类型、号码、金额、日期、销售方、文件名及对应预览页码，并显示报销金额、合计、差额和张数。生成后可直接打开核对表。
5. 在预览区点击“打印”，会弹出 Windows 打印设置窗口，可在里面选择打印机、份数、单双面、纸张等设置。
6. 打印后确认标记已使用，程序会把发票移动到 `used_invoices`，并写入 `data\invoice_records.sqlite3` 和 `data\used_records.json`。

在“可用发票”列表选中一张发票后，可通过“设置类型”菜单或右键菜单修改类型，也可恢复自动识别。手动类型保存在 SQLite 数据库中，重新扫描、重启或改名后仍保留，并同步更新推荐组合中的对应发票。修改已生成预览中的发票类型后，需重新生成组合预览以应用新的类型排序。

组合规则：合计金额必须大于等于报销金额；在最少超额金额的 50 元范围内，优先选择开票日期更早的发票。

图片使用内置中文 OCR 离线识别金额、发票号码、日期和类型，无需安装其他识别程序或联网。识别结果缓存在 `data\ocr_cache`，图片转 PDF 的预览和打印缓存保存在 `print_jobs\image_cache`。首次识别图片需要加载 OCR 模型，扫描在后台运行。

识别不到金额的 PDF/OFD/图片会被移动到 `failed_invoices`，并写入 `data\failed_records.json` 供人工处理。图片应清晰完整，尤其是发票号码和价税合计区域。

扫描时会按发票号码去重：如果 `invoices` 中存在相同发票号码的不同格式或不同文件名，程序只保留一份可用发票，依次优先保留 PDF、OFD、图片，其余重复文件移动到 `void_invoices`。

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
