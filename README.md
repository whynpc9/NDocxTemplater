# NDocxTemplater

一个基于 OpenXML 的 `.docx + JSON` 模板渲染库，目标是提供类似 `docxtemplater` 的 .NET 替代方案。

## 目标框架

- `netstandard2.0`
- `netstandard2.1`

测试工程使用 `net10.0 + xUnit`。

## 已实现能力

- 基础标签替换：`{a.b[0].c}`
- 条件分支：`{?expr} ... {/?expr}`
- 循环：`{#expr} ... {/expr}`
- 表格循环：把循环标记放在表格行中，可按行复制输出列表数据
- 支持 Word 将标签拆分到多个 Run/Text 节点后的渲染（包含表格单元格内格式表达式）
- 图片标签（参考 docxtemplater image tag 风格）
  - inline：`{%imagePath}`
  - block/居中：`{%%imagePath}`
  - 数据支持：base64 data URI、base64 字符串、文件路径
  - 当前约束：图片标签需独占一个段落（单独一行）
- 表达式扩展（管道语法）
  - 排序：`|sort:key:asc` 或 `|sort:key:desc`
  - 截断：`|take:10`
  - 计数：`|count`
  - inline 友好聚合/取值：
    - 首项/末项：`|first`、`|last`
    - 按字段取最大/最小项：`|maxby:key`、`|minby:key`
    - 从当前值继续取字段：`|get:path`（别名：`|pick:path`）
  - 格式化：
    - 数值：`|format:number:0.00`
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
```

## 快速使用

```csharp
using NDocxTemplater;

var engine = new DocxTemplateEngine();
var templateBytes = File.ReadAllBytes("template.docx");
var json = File.ReadAllText("data.json");

var outputBytes = engine.Render(templateBytes, json);
File.WriteAllBytes("output.docx", outputBytes);
```

## NuGet Package

- Package ID: `NDocxTemplater`
- Repository: `https://github.com/whynpc9/NDocxTemplater`

发布由 GitHub Actions 自动完成：

- CI：`build + test + lint(dotnet format)`
- 发布：打 tag（如 `v0.1.0`）或手动触发 `Publish NuGet` workflow
- NuGet API Key 使用仓库 Secret：`NUGET_API_KEY`

## Examples

`examples` 下每个子目录都是一个独立用例，均包含：

- `template.docx`：模板文件
- `data.json`：输入数据
- `output.docx`：按模板渲染后的结果文档
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

如需重新生成示例资产：

```bash
dotnet run --project tools/ExampleGenerator/ExampleGenerator.csproj --disable-build-servers
```

## 测试

```bash
dotnet test NDocxTemplater.sln --disable-build-servers -m:1
```

当前测试覆盖了：基础替换、条件、循环、表格映射、图片渲染、排序/截断/计数/格式化、inline 聚合表达式、表格内拆分 Run 标签格式化。

## Acknowledgements

- 本项目受 [docxtemplater](https://github.com/open-xml-templating/docxtemplater) 启发，感谢原项目作者与社区在文档模板化领域的工作。
- `NDocxTemplater` 为独立的 .NET 实现，不包含对原项目源码的直接移植。

## Code Provenance

- 本仓库的初始版本由 OpenAI Codex 协助生成，并由仓库维护者进行审阅、修改与测试。

## License

本项目使用 MIT License，详见 `LICENSE`。
