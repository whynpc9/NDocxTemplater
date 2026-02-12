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
- 表达式扩展（管道语法）
  - 排序：`|sort:key:asc` 或 `|sort:key:desc`
  - 截断：`|take:10`
  - 计数：`|count`
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
```

各示例说明：

- `01-basic-tags`：基础路径替换（包含数组下标）
- `02-condition`：条件分支
- `03-loop`：段落循环
- `04-table-loop`：表格行循环
- `05-extensions`：排序/截断/计数/格式化

如需重新生成示例资产：

```bash
dotnet run --project tools/ExampleGenerator/ExampleGenerator.csproj --disable-build-servers
```

## 测试

```bash
dotnet test NDocxTemplater.sln --disable-build-servers -m:1
```

当前测试覆盖了：基础替换、条件、循环、表格映射、排序/截断/计数/格式化。

## Acknowledgements

- 本项目受 [docxtemplater](https://github.com/open-xml-templating/docxtemplater) 启发，感谢原项目作者与社区在文档模板化领域的工作。
- `NDocxTemplater` 为独立的 .NET 实现，不包含对原项目源码的直接移植。

## Code Provenance

- 本仓库的初始版本由 OpenAI Codex 协助生成，并由仓库维护者进行审阅、修改与测试。

## License

本项目使用 MIT License，详见 `LICENSE`。
