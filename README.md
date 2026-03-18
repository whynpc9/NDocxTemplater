# NDocxTemplater

一个基于 OpenXML 的模板渲染库，当前支持 `.docx + JSON` 和 `.xlsx + JSON`，目标是提供类似 `docxtemplater` 的 .NET 替代方案。

## 目标框架

- `netstandard2.0`
- `netstandard2.1`

测试工程使用 `net10.0 + xUnit`。

## 已实现能力

- 基础标签替换：`{a.b[0].c}`
- 条件分支：`{?expr} ... {/?expr}`
- 循环：`{#expr} ... {/expr}`
- 表格循环：把循环标记放在表格行中，可按行复制输出列表数据
- `.xlsx` 工作表行循环：把循环/条件标记放在工作表行中，可按行复制输出列表数据，并保留模板行样式
- `.xlsx` 工作表媒体占位符：在单元格中使用图片/条形码占位符，渲染为 worksheet drawing
- `.xlsx` 模板行复制时支持合并单元格重建，以及公式中的相对行引用修正
- 支持 Word 将标签拆分到多个 Run/Text 节点后的渲染（包含表格单元格内格式表达式）
- 图片标签（参考 docxtemplater image tag 风格）
  - inline：`{%imagePath}`
  - block/居中：`{%%imagePath}`
  - 数据支持：base64 data URI、base64 字符串、文件路径
  - 支持缩放参数（图片数据对象）：
    - `width` / `height`：目标尺寸（像素）
    - `maxWidth` / `maxHeight`：最大边界（像素）
    - `scale`：缩放倍率（如 `0.25`）
    - `preserveAspectRatio`（或 `keepAspectRatio`）：是否保持宽高比，避免拉伸变形
  - 当前约束：图片标签需独占一个段落（单独一行）
- 条形码标签（inline/block 与图片标签一致）
  - inline：`{%barcode:expr;type=code128;width=360;height=96}`
  - block/居中：`{%%barcode:expr;type=ean13;width=280;height=100}`
  - `expr` 为数据源路径（应解析为字符串/数字）
  - 支持参数：
    - `type` / `format`：`code128`、`code39`、`code93`、`codabar`、`ean13`、`ean8`、`upca`、`itf`
    - `width` / `height`：输出尺寸（像素）
    - `margin`：条码边距（像素）
    - `pure`：是否仅输出条码图形（`true/false`）
  - 实现基于 `NetBarcode`（`netstandard2.0` 资产），避免运行时加载 `PresentationCore` 的问题
- 表达式扩展（管道语法）
  - 排序：`|sort:key:asc` 或 `|sort:key:desc`
  - 截断：`|take:10`
  - 计数（可直接用于 inline 文本）：`|count`
  - inline 友好聚合/取值：
    - 首项/末项：`|first`、`|last`
    - 位次取项：`|nth:N`（1-based，如第3名）、`|at:index`（0-based，支持负数如 `-1`）
    - 按字段取最大/最小项：`|maxby:key`、`|minby:key`
    - 从当前值继续取字段：`|get:path`（别名：`|pick:path`）
  - inline 条件分支：
    - `|if:trueText:falseText`
    - `|if:trueText`（false 时输出空字符串）
  - 格式化：
    - 数值：`|format:number:0.00`
    - 百分比（原始值为小数）：`|format:percent:0.00` -> `1.23%`
    - 千分比（原始值为小数）：`|format:permille:0.00` -> `12.30‰`
    - 日期：`|format:date:yyyy-MM-dd`

## 模板语法示例

```text
患者: {patient.name}

{?patient.isVip}
VIP 客户
{/?patient.isVip}

{#orders|sort:amount:desc|take:2}
{id} -> {amount|format:number:0.00}
{/orders|sort:amount:desc|take:2}

{%logo}
{%%cover}
{%barcode:barcodes.code128;type=code128;width=360;height=96}
```

图片数据对象示例（路径 / data URI + 等比缩放）：

```json
{
  "fromPath": {
    "src": "chart.png",
    "maxWidth": 376,
    "preserveAspectRatio": true
  },
  "fromDataUri": {
    "src": "data:image/png;base64,...",
    "scale": 0.25,
    "preserveAspectRatio": true
  }
}
```

## Inline 文本段落友好写法

在报告正文段落里，通常希望用单行模板直接表达“区间 + 最大值/最小值”这类语句。可以组合 `sort/first/last/maxby/minby/get/format` 来完成，而不需要额外的块标签换行。

```text
统计数据包括了从{financeMonthly|sort:month:asc|first|get:month|format:date:yyyy年M月}
到{financeMonthly|sort:month:asc|last|get:month|format:date:yyyy年M月}的财务数据，
其中营收最高的是{financeMonthly|maxby:revenue|get:month|format:date:M月}，
营收为{financeMonthly|maxby:revenue|get:revenue|format:number:#,##0}元

在这些机构的对比数据中，其中营收最高的为{institutions|maxby:revenue|get:name}，
收入为{institutions|maxby:revenue|get:revenue|format:number:#,##0}元，
营收最低的为{institutions|minby:revenue|get:name}，
收入为{institutions|minby:revenue|get:revenue|format:number:#,##0}元

前10名机构中，第3名为{institutions|sort:revenue:desc|take:10|nth:3|get:name}，
收入为{institutions|sort:revenue:desc|take:10|nth:3|get:revenue|format:number:#,##0}元，
前10名末位为{institutions|sort:revenue:desc|take:10|at:-1|get:name}

本次样本共{institutions|count}家机构，状态：{flags.includeRates|if:包含比率指标:不包含比率指标}，
环比增长率{metrics.growthRate|format:percent:0.00}，
坏账率{metrics.badDebtRate|format:permille:0.00}
```

说明：
- `|count` 已支持 inline 文本（例如 `共有{items|count}项`）
- `|format:number:0.00%` / `|format:number:0.00‰` 也仍然可用；`percent/permille` 只是更直观的别名

## XLSX 模板行语法

`.xlsx` 目前遵从现有“表格部分”的控制标记语法，只是把载体换成了工作表行：

```text
第 1 行: Item | Qty | Amount
第 2 行: {#report.lines|sort:amount:desc|take:2} | | 
第 3 行: {name} | {qty} | {amount|format:number:0.00}
第 4 行: {/report.lines|sort:amount:desc|take:2} | |
第 5 行: {?showFooter} | |
第 6 行: Count | {report.lines|count} |
第 7 行: {/?showFooter} | |
```

规则：
- 控制标记行应只有一个非空单元格包含完整标记，其他单元格留空
- 循环块会复制中间模板行，并保留模板行/单元格样式
- 普通单元格内联表达式与 `.docx` 使用同一套路径和管道语法
- 图片/条形码占位符可放在单元格中，建议独占一个单元格：
  - 图片：`{%pathImage}`、`{%%pathImage}`
  - 条形码：`{%barcode:barcodes.ean13;type=ean13;width=220;height=80}`
- 模板行被循环复制时：
  - 行内公式会按最终行号重写相对引用
  - 模板中的合并单元格区域会随复制后的行块一起展开重建
  - 位于循环块之后的汇总公式，若引用了循环块行范围，也会按最终输出行范围扩展

## 快速使用

```csharp
using NDocxTemplater;

var engine = new DocxTemplateEngine();
var templateBytes = File.ReadAllBytes("template.docx");
var json = File.ReadAllText("data.json");

var outputBytes = engine.Render(templateBytes, json);
File.WriteAllBytes("output.docx", outputBytes);
```

`.xlsx` 用法：

```csharp
using NDocxTemplater;

var engine = new XlsxTemplateEngine();
var templateBytes = File.ReadAllBytes("template.xlsx");
var json = File.ReadAllText("data.json");

var outputBytes = engine.Render(templateBytes, json);
File.WriteAllBytes("output.xlsx", outputBytes);
```

## NuGet Package

- Package ID: `NDocxTemplater`
- Repository: `https://github.com/whynpc9/NDocxTemplater`

发布由 GitHub Actions 自动完成：

- CI：`build + test + lint(dotnet format)`
- 发布：打 tag（如 `v0.1.0`）或手动触发 `Publish NuGet` workflow
- NuGet API Key 使用仓库 Secret：`NUGET_API_KEY`

## Examples

`examples` 下每个子目录都是一个独立用例，至少包含：

- `template.docx` 或 `template.xlsx`：模板文件
- `data.json`：输入数据
- `output.docx` 或 `output.xlsx`：按模板渲染后的结果文档
- `example.cs`：最小调用代码

目录如下：

```text
examples/
  01-basic-tags/
  02-condition/
  03-loop/
  04-table-loop/
  05-extensions/
  06-images/
  07-table-date-format-split-runs/
  08-inline-friendly-expressions/
  09-inline-ranking-positions/
  10-inline-conditions-and-rates/
  11-images-file-and-datauri-scaling/
  12-barcodes/
  13-xlsx-row-loop/
  14-xlsx-media-placeholders/
  15-xlsx-merged-cells-and-formulas/
```

各示例说明：

- `01-basic-tags`：基础路径替换（包含数组下标）
- `02-condition`：条件分支
- `03-loop`：段落循环
- `04-table-loop`：表格行循环
- `05-extensions`：排序/截断/计数/格式化
- `06-images`：图片标签（inline/block）和循环中的图片渲染
- `07-table-date-format-split-runs`：表格单元格内被拆分 Run 的日期格式表达式渲染
- `08-inline-friendly-expressions`：面向正文段落的 inline 聚合/取值表达式（区间、最大/最小值）
- `09-inline-ranking-positions`：面向“前 N 名中的第 K 名”场景的位次表达式（`nth/at`）
- `10-inline-conditions-and-rates`：inline 条件分支、计数，以及百分比/千分比格式化
- `11-images-file-and-datauri-scaling`：使用真实 PNG 验证文件路径 / data URI 图片插入，以及缩放与等比适配
  - 该示例额外包含 `chart.png`（用于文件路径模式）
- `12-barcodes`：由模板占位符指定数据源路径与条码参数（类型、尺寸、边距）
- `13-xlsx-row-loop`：`.xlsx` 工作表按行循环、条件行和单元格表达式替换
- `14-xlsx-media-placeholders`：`.xlsx` 工作表中的图片/条形码占位符，包含文件路径、data URI 和条形码参数
  - 该示例额外包含 `chart.png`（用于文件路径模式）
- `15-xlsx-merged-cells-and-formulas`：`.xlsx` 循环块中的跨行合并单元格，以及模板复制后的公式引用修正

如需重新生成示例资产：

```bash
dotnet run --project tools/ExampleGenerator/ExampleGenerator.csproj --disable-build-servers
```

## 测试

```bash
dotnet test NDocxTemplater.sln --disable-build-servers -m:1
```

当前测试覆盖了：基础替换、条件、循环、表格映射、`.xlsx` 工作表行循环/条件/样式保留、`.xlsx` 图片/条形码占位符、`.xlsx` 合并单元格与公式引用修正、图片渲染（含文件路径/data URI真实 PNG、缩放与等比适配）、条形码渲染（类型/尺寸参数）、排序/截断/计数/格式化、inline 聚合/位次表达式、inline 条件分支、百分比/千分比格式化、表格内拆分 Run 标签格式化。

## Acknowledgements

- 本项目受 [docxtemplater](https://github.com/open-xml-templating/docxtemplater) 启发，感谢原项目作者与社区在文档模板化领域的工作。
- `NDocxTemplater` 为独立的 .NET 实现，不包含对原项目源码的直接移植。

## Code Provenance

- 本仓库的初始版本由 OpenAI Codex 协助生成，并由仓库维护者进行审阅、修改与测试。

## License

本项目使用 MIT License，详见 `LICENSE`。
