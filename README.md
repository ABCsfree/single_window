# Trade XML Studio

这是一个独立的 WPF 出口锂电包装报文生成工具，用于生成信息补录 `ELBP004` 和随附单据 `ELBP005` XML。项目不加载、修改或依赖原始 EXE。

## 功能

- 单笔生成：读取一个文件夹中的 4 张图片，按文件名排序映射到 `A1`–`A4`。
- 报文包：每个箱号生成 1 个 `ELBP004` 主报文，并为每个附件生成 1 个独立 `ELBP005` 文件。
- P0 文件：可选代理委托协议，业务类型为 `P0`，单独生成对应 `ELBP005` 并写入 `ELBP004` 附件清单。
- 企业配置：出口企业和申请单位分别手动填写并保存，不互相复用。
- 出口企业方案：支持新增、修改方案名和下拉切换；每个方案独立保存企业名称、海关十位编码、统一社会信用代码及主管海关代码。新增方案为空白，旧配置自动转为“默认方案”。切换时保留当前编辑，点击“保存配置”保存全部方案及当前选择，重启后恢复。
- Excel 批量：支持 `.xlsx` / `.xlsm`，读表模式可选 BC 栏（B 栏序号、C 栏箱号，默认）或 AB 栏（A 栏序号、B 栏箱号）。切换模式后自动刷新预览。
- 三种影像目录模式：
  - 每个箱号对应一个同名小文件夹；
  - 每个序号对应一个同名小文件夹；
  - 所有图片位于同一目录，使用图片名第一个 `-` 前的数字匹配箱号末尾数字。
- 批量预览、重复箱号拦截、每批 3/4/5 张图片数量校验、原附件 2.25MB 大小校验、同箱号 XML 覆盖确认。
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

本项目按 2026-08-25 字段说明生成 `ELBP004/ELBP005`；`EdocID` 使用出口企业统一社会信用代码、主管海关代码、附件代码、毫秒时间戳和 20 位流水号组成。签名字段留空并交由单一窗口导入客户端处理。官网字段表、XSD 注释和样例在个别长度与格式上仍存在差异，正式使用前应拿一组最新验收成功的报文做逐字段比对。
