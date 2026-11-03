# 文档转换 / Document Converter

在 DeepSeek Harness 工作区里转换常用文档格式。插件提供两个入口，共用同一套转换服务：

- **侧边栏「文档转换」面板**：从本机拖入/选择文件上传，或浏览工作区已有文件，选目标格式，转换后直接下载。
- **`convert_document` 对话工具**：让 Agent 直接把某个文件转成另一种格式，结果写回工作区。

## 安装

```sh
dsh plugin --profile web add dsh-doc-convert
# 或直接从 GitHub 安装：
dsh plugin --profile web add github:zsy334/dsh-doc-convert
```

首次安装会下载约 160MB 的 LibreOffice Kit 引擎；Linux 没有原生引擎包，会使用 WASM 版本。

## 支持矩阵

| 源格式 | 可转目标 |
| --- | --- |
| docx | pdf、html、md、txt、odt |
| doc（Word 97-2003） | docx、pdf、html、md、txt、odt |
| odt | docx、pdf、html、md、txt |
| txt | docx、pdf、html、md |
| md | html、docx、pdf、txt |
| html | md、docx、pdf、txt |
| pdf | txt、md（仅文本抽取） |
| xlsx / xls / ods | csv、pdf、ods/xlsx、html、json |
| csv | xlsx、pdf、html、json |
| tsv | csv、xlsx、json |
| pptx / ppt / odp | pdf、pptx/odp |

说明：

- Office 与 PDF 版式转换由随插件安装的 **LibreOffice Kit** 完成；Markdown/HTML/TXT 与数据格式由 JS 管线完成。
- PDF 只做文本/Markdown 抽取，**不**做 PDF→Word/PPT 的版面重建；扫描件需要 OCR，本版不支持。
- 多工作表导出 CSV 时需填写准确的工作表名（面板中有输入框，工具使用 `sheet` 参数）。
- 输出默认写在源文件同目录，重名时自动加 `-1`、`-2`；所有写入都被限制在工作区内。

## 使用

面板：点击左侧栏的「文档转换」图标，选择工作区与文件，再选目标格式。

工具：

```json
{ "input_path": "docs/report.md", "target_format": "pdf" }
```

## 结构

- `index.js`：Host 入口，注册工具与 HTTP API。
- `lib/formats.js`：格式注册表与支持矩阵（唯一事实来源）。
- `lib/convert.js`：转换路由与 JS 管线。
- `lib/docx.js`：HTML/文本 → DOCX 生成、HTML → 文本。
- `lib/engine.js`：LibreOffice Kit 子进程桥接。
- `lib/tool.js`：`convert_document` 工具。
- `lib/http.js`：面板使用的 `/doc-convert/api/*` 接口。
- `client.js`：侧边栏图标与转换面板。

重新安装依赖：在本目录运行 `pnpm install --ignore-scripts`（引擎约 160MB）。
