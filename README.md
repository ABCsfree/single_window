# Trade XML Studio

这是一个独立的 WPF 出口锂电包装报文生成工具，用于生成信息补录 `ELBP004` 和随附单据 `ELBP005` XML。项目不加载、修改或依赖原始 EXE。

## 功能

- 单笔生成：读取一个文件夹中的 4 张图片，按文件名排序映射到 `A1`–`A4`。
- 报文包：每个箱号生成 1 个 `ELBP004` 主报文，并为每个附件生成 1 个独立 `ELBP005` 文件。
- P0 文件：可选代理委托协议，业务类型为 `P0`，单独生成对应 `ELBP005` 并写入 `ELBP004` 附件清单。
- 企业配置：出口企业和申请单位分别手动填写并保存，不互相复用。
- 出口企业方案：支持新增、修改方案名、下拉切换和删除；每个方案独立保存企业名称、海关十位编码、统一社会信用代码及主管海关代码。新增方案为空白，旧配置自动转为“默认方案”。删除当前方案需确认，删除后切换到下一个方案，若无下一个则切换到前一个；至少保留一个方案。切换时保留当前编辑，点击“保存配置”保存全部方案、删除结果及当前选择，重启后恢复。
- Excel 批量：支持 `.xlsx` / `.xlsm`，读表模式可选 BC 栏（B 栏序号、C 栏箱号，默认）或 AB 栏（A 栏序号、B 栏箱号）。切换模式后自动刷新预览。
- 三种影像目录模式：
  - 每个箱号对应一个同名小文件夹；
  - 每个序号对应一个同名小文件夹；
  - 所有图片位于同一目录，使用图片名第一个 `-` 前的数字匹配箱号末尾数字。
- 批量预览、重复箱号拦截、每批 3/4/5 张图片数量校验、附件自动压缩、同箱号 XML 覆盖确认。
- 信息补录支持 `C（申报）` 与 `G（暂存）`；附件上传类型支持 `F（首次上传）` 与 `P（重传/补传）`，窗口内附有用途说明。
- 公共配置保存到程序目录下的 `trade-xml-config.json`。

## 构建与运行

需要 Windows 和 .NET 10 SDK：

```powershell
dotnet build .\TradeXmlStudio.slnx
dotnet run --project .\src\TradeXmlStudio.App\TradeXmlStudio.App.csproj
```

仓库内已配置可复用的本地 SDK 时，也可以直接使用：

```powershell
.\dotnet-local.cmd build .\TradeXmlStudio.slnx
.\dotnet-local.cmd run --project .\src\TradeXmlStudio.App\TradeXmlStudio.App.csproj
```

运行无第三方测试框架的核心自检：

```powershell
dotnet run --project .\tests\TradeXmlStudio.Core.SmokeTests\TradeXmlStudio.Core.SmokeTests.csproj
```

发布自包含的单文件 Windows 程序：

```powershell
dotnet publish .\src\TradeXmlStudio.App\TradeXmlStudio.App.csproj -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true
```

使用仓库内的本地 SDK 发布：

```powershell
.\dotnet-local.cmd publish .\src\TradeXmlStudio.App\TradeXmlStudio.App.csproj -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true
```

## Excel 约定

根据“读表模式”读取序号和箱号：默认 BC 栏，也可选择 AB 栏。程序跳过所选箱号栏为空的行，以及所选栏中的“序号/箱号”表头行。序号必须为大于 0 的整数，否则提示对应的行号和栏位。每一行按“本批上传张数”生成 A1–A3、A1–A4 或 A1–A4+B1（以及启用时的 P0）对应的 `ELBP005`；整批共用一个 `ELBP004`。点击“批量生成”时会按当前读表模式重新读取 Excel，因此外部修改后只需保存文件，不需要重启程序。

## 与参考程序的边界

### 图片格式和自动压缩

单笔与批量生成均根据图片实际内容识别 JPEG、PNG、BMP，生成时同步 ELBP004 与 ELBP005 的附件名称和类型；例如内容为 PNG、扩展名为 `.jpg` 的合格图片，保持原字节并将报文中的名称修正为 `.png`。不会修改磁盘上的原图。

超过处理后附件上限或报文安全预算的照片，会依次尝试 JPEG 质量 95、90、85、80，选择第一个达标结果。保持像素分辨率，按 EXIF 方向旋转，透明区域填充白色；输出 JPEG 不保留原元数据。小于目标的图片不重新编码。压缩有损，请抽查标签小字和条码。

“处理后附件上限”默认 2.25 MiB。程序另外将未签名 ELBP005 的安全目标设为 `3 MiB - 64 KiB`（3,080,192 字节），并在估算 Base64 负载时再预留 16 KiB XML 字段空间；生成后检查实际 UTF-8 XML 字节数。3 MiB 仅是根据失败样本推测的客户端上限，以上是本程序的保守保护值，不是已确认的官方限制。调大附件上限不会取消报文大小保护，补签后仍需确认平台回执。

质量降至 80 仍超限、图片损坏或格式不支持时，整包在写入 XML 前失败，并提示具体附件。PDF 等非图片附件不自动压缩，超限需手动处理。图片转换使用 Windows WPF 图像编解码器，因此核心库与测试项目也以 Windows 为目标。

可选地用本机 ELBP005 样本验证压缩流程（不修改样本，不将内容放入仓库）：

```powershell
.\dotnet-local.cmd run --project .\tests\TradeXmlStudio.Core.SmokeTests -- "C:\path\ELBP005-sample.xml"
```

本项目按 2026-08-25 字段说明生成 `ELBP004/ELBP005`；`EdocID` 使用出口企业统一社会信用代码、主管海关代码、附件代码、毫秒时间戳和 20 位流水号组成。签名字段留空并交由单一窗口导入客户端处理。官网字段表、XSD 注释和样例在个别长度与格式上仍存在差异，正式使用前应拿一组最新验收成功的报文做逐字段比对。
