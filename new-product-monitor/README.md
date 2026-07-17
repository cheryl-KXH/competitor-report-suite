# 竞品新品周报生成工具

这个项目用于从日常维护的钉钉 AI 表中拉取新品数据，生成接近示例版式的周报 Excel。

## 项目结构

```text
competitor-new-product-monitor/
├── README.md
├── .gitignore
├── requirements.txt
├── config/
│   ├── dingtalk.example.json
│   ├── dingtalk.json
│   ├── field_mapping.json
│   ├── report_rules.json
│   ├── excel_layout.json
│   ├── image_layout.json
│   └── font_files.json
├── assets/
│   ├── fonts/
│   │   └── README.md
│   └── logo/
│       └── README.md
├── scripts/
│   ├── generate_weekly_report.py
│   ├── generate_weekly_outputs.py
│   └── render_report_image.py
└── outputs/
    └── 竞品新品周报YYYY-MM-DD_YYYY-MM-DD/
        ├── 竞品新品周报YYYY-MM-DD_YYYY-MM-DD.xlsx
        ├── 竞品新品周报YYYY-MM-DD_YYYY-MM-DD.html
        └── 竞品新品周报YYYY-MM-DD_YYYY-MM-DD.png
```

- `config/dingtalk.example.json`：可提交的钉钉连接配置模板，不含真实 token。
- `config/dingtalk.json`：本地私密钉钉连接配置，放真实 Streamable HTTP URL 或使用本机 mcporter 服务名。这个文件已被 `.gitignore` 忽略。
- `config/field_mapping.json`：钉钉字段如何翻译成脚本标准字段。
- `config/report_rules.json`：这份周报的业务规则。
- `config/excel_layout.json`：Excel 怎么排版、数据写到哪里。
- `config/image_layout.json`：HTML 和 PNG 长图怎么排版、渲染。
- `config/font_files.json`：内部中英文字体文件路径配置。
- `assets/logo/`：周报顶部 logo 本地放置位置。logo 是内部专用资产，不提交、不上传外部。
- `assets/fonts/`：字体文件本地放置位置。字体不要提交、不要上传外部。
- `scripts/generate_weekly_report.py`：现有 Excel 生成脚本。
- `scripts/generate_weekly_outputs.py`：统一输出入口，控制 `--output-mode`。
- `scripts/render_report_image.py`：HTML/PNG 长图渲染模块。
- `outputs/`：默认 Excel 输出目录。

## 环境准备

脚本需要 Python 3、`openpyxl`、`Pillow`、`playwright`、本机 Google Chrome 和 `mcporter`。

```bash
pip install -r requirements.txt
```

首次使用 PNG 长图输出前，建议安装 Playwright 运行环境：

```bash
python -m playwright install chromium
```

当前脚本默认会优先调用本机 `/Applications/Google Chrome.app` 做长图截图。

安装依赖后，可以运行配置校验：

```bash
python3 scripts/generate_weekly_report.py --validate-config

```

## 配置钉钉连接

第一次使用时复制模板：

```bash
cp config/dingtalk.example.json config/dingtalk.json
```

`config/dingtalk.json` 里有两种配置方式：

1. 如果本机 `mcporter` 已经配置了 `dingtalk-ai-table`，保持 `streamableHttpUrl` 为空即可。
2. 如果交接给继任者或需要换账号，把新的 Streamable HTTP URL 填入 `streamableHttpUrl`。

注意：Streamable HTTP URL 带访问令牌，不要提交到 git，不要上传外部平台。

如果要用钉钉按钮生成并上传到知识库，还需要配置钉钉文档/知识库 MCP：

```bash
cp config/dingtalk_docs.example.json config/dingtalk_docs.json
```

`config/dingtalk_docs.json` 用于本机私密配置，不提交 git。推荐按业务填写知识库根目录：

```json
{
  "streamableHttpUrl": "钉钉文档 MCP Streamable HTTP URL",
  "workspaceId": "知识库 ID 或知识库 URL",
  "businessRootFolderIds": {
    "喜茶": "Gl6Pm2Db8D3mok1vFnG9zNaAJxLq0Ee4",
    "野萃山": "ydxXB52LJq7l1yqMi3QMAQLpWqjMp697",
    "茶坊": "jb9Y4gmKWr7l1yqji4mbrrBEVGXn6lpz"
  }
}
```

上传时会先按状态机行的 `业务` 选择对应根目录，再在根目录下查找四位年份文件夹，例如 `2025`；年份文件夹不存在时会自动创建。

也可以不写 `streamableHttpUrl`，改用环境变量：

```bash
export DINGTALK_DOCS_MCP_URL="钉钉文档 MCP Streamable HTTP URL"
```

## 字体配置

本项目使用以下字体：

- 中文：`方正FW筑紫黑简 R.ttf`
- 英文/数字：`Heytea Sans Serif Regular.otf`

默认字体配置写在 `config/font_files.json`：

字体来源统一由 `fontDirectory` 和各字体的 `fileName` 决定，不再支持本机绝对路径。把字体文件放到 `assets/fonts/` 后，在 `config/font_files.json` 中配置文件名，例如：

```json
{
  "fontDirectory": "assets/fonts",
  "defaultFonts": {
    "chineseExcelName": "Microsoft YaHei",
    "latinExcelName": "Arial",
    "cssFamily": "Microsoft YaHei, Arial, sans-serif"
  },
  "chineseFont": {
    "excelName": "方正FW筑紫黑简 R",
    "defaultExcelName": "Microsoft YaHei",
    "fileName": "方正FW筑紫黑简 R.ttf"
  },
  "latinFont": {
    "excelName": "Heytea Sans Serif Regular",
    "defaultExcelName": "Arial",
    "fileName": "Heytea Sans Serif Regular.otf"
  }
}
```
字体是内部独家字体，不提交、不上传外部；`.gitignore` 已忽略 `assets/fonts/` 下的字体文件。Excel 文件只记录字体名称，不会嵌入字体文件，打开 Excel 的电脑也需要安装对应字体。如果字体文件缺失，脚本会使用 `defaultFonts` 或 `defaultExcelName`，不会阻断生成。

## CLI 命令

推荐使用统一入口，默认输出 Excel、HTML 和 PNG 三种产物：

```bash
python scripts/generate_weekly_outputs.py
```

指定周期：

```bash
python scripts/generate_weekly_outputs.py --startDate 20260509 --endDate 20260515
```

指定输出模式：

```bash
python scripts/generate_weekly_outputs.py --output-mode excel
python scripts/generate_weekly_outputs.py --output-mode html
python scripts/generate_weekly_outputs.py --output-mode image
python scripts/generate_weekly_outputs.py --output-mode all
```

说明：

- 不写 `--output-mode` 时默认等同于 `--output-mode all`，生成 `.xlsx`、`.html`、`.png`。
- `--output-mode image` 只保留 PNG，截图用的 HTML 临时文件不会写入输出目录。
- 每次生成都会创建一个周报文件夹，例如 `outputs/竞品新品周报2026-05-09_2026-05-15/`。
- 如果同名文件夹已存在，会生成 `_2`、`_3` 后缀的新文件夹。

现有 Excel 单独生成脚本仍然保留：

默认按“周六到周五”生成：周五导出上周六到本周五，周一导出上一个完整周六到周五。

```bash
python scripts/generate_weekly_report.py
```

指定周期：

```bash
python scripts/generate_weekly_report.py --startDate 20260502 --endDate 20260508
```

指定品牌：

```bash
python scripts/generate_weekly_report.py --brands 霸王茶姬,古茗,乐乐茶
```

按业务输出：

```bash
python scripts/generate_weekly_report.py --business 喜茶
python scripts/generate_weekly_report.py --business 野萃山
python scripts/generate_weekly_report.py --business 茶坊
python scripts/generate_weekly_outputs.py --business 喜茶
```

同时指定周期和品牌：

```bash
python scripts/generate_weekly_report.py --startDate 20260502 --endDate 20260508 --brands 霸王茶姬,古茗
```

如果同时传 `--brands` 和 `--business`，脚本优先使用 `--brands` 的手动品牌范围。

指定输出文件：

```bash
python scripts/generate_weekly_report.py --output outputs/竞品新品周报2026-05-02_2026-05-08.xlsx
```

输出文件不会覆盖已有文件。如果同名文件已经存在，脚本会按当前最大序号继续生成 `_2.xlsx`、`_3.xlsx` 这样的新文件名；即使中间序号文件被删除，也不会回头复用旧序号。默认输出和显式 `--output` 都遵循这个规则。

校验配置：

```bash
python scripts/generate_weekly_report.py --validate-config
```

解释当前配置：

```bash
python scripts/generate_weekly_report.py --explain-config
```

## 钉钉按钮生成服务

按钮生成走本机 FastAPI 服务，适合配合 ngrok 暴露给钉钉调用。

无密钥模式：

```bash
export DINGTALK_DOCS_MCP_URL="钉钉文档 MCP Streamable HTTP URL"
uvicorn service.app:app --host 0.0.0.0 --port 8000
```

有密钥模式：

```bash
export REPORT_SERVICE_SECRET="自己设置一个密钥"
export DINGTALK_DOCS_MCP_URL="钉钉文档 MCP Streamable HTTP URL"
uvicorn service.app:app --host 0.0.0.0 --port 8000
```

如果不设置 `REPORT_SERVICE_SECRET`，服务不会校验请求里的 `secret`。ngrok 地址是公网地址，长期使用建议设置密钥。

再启动 ngrok：

```bash
ngrok http 8000
```

钉钉按钮/自动化请求：

```http
POST https://你的-ngrok域名/generate-weekly-report
Content-Type: application/json
```

无密钥请求体：

```json
{
  "recordId": "当前行记录ID"
}
```

有密钥请求体：

```json
{
  "recordId": "当前行记录ID",
  "secret": "自己设置的密钥"
}
```

服务会读取 `竞品新品周报` 状态机表当前行：

- `业务`：用于主表业务筛选、关注品牌小字、文件命名。
- `开始时间`、`截止时间`：用于筛产品主表 `上市日期`。
- `年份`、`周次`：只用于命名，不参与产品筛选。

生成文件夹命名为：

```text
喜茶25W02_20250110
```

知识库上传路径示例：

```text
原始文件：喜茶竞品新品周报/2025/喜茶25W02_20250110/
原始文件：野萃山竞品新品周报/2025/野萃山25W02_20250110/
原始文件：茶坊竞品新品周报/2025/茶坊25W02_20250110/
```

同一行重复点击会覆盖同名本地输出目录，并覆盖 AI 表 `周报` 字段的知识库目录链接。状态流：

```text
等待1-2分钟 -> 已生成
等待1-2分钟 -> 生成异常
等待1-2分钟 -> 生成失败
```

状态机表字段要求：

- `周报`：URL 字段。
- `生成状态`：单选字段，选项为 `未生成`、`等待1-2分钟`、`已生成`、`生成异常`、`生成失败`。
- `生成时间`：日期时间字段。
- `反馈信息`：富文本字段。

## 当前业务规则

默认跟踪品牌：

默认不传 `--brands` 或 `--business` 时，不按固定品牌清单筛选，按本次钉钉表取数结果里实际出现的品牌输出；品牌顺序按钉钉 AI 表 `品牌` 字段的标签列表顺序。传 `--business` 时，主表记录按钉钉 AI 表 `业务` 字段筛选，支持 `喜茶`、`野萃山`、`茶坊`。

汇总表下方“关注品牌包括”小字：

- 传 `--business` 时，脚本读取同一个 base 下的 `关注竞品品牌` 表。
- 只列出 `业务需求方` 包含当前业务且 `跟踪状态` 为 `进行中` 的品牌。
- 品牌按 `关注竞品品牌` 表里 `品牌` 字段的标签顺序排序。
- 不传 `--business` 时，小字仍使用本次实际输出的品牌列表。

产品排序：

- 品牌顺序按上面的默认/业务/手动品牌规则。
- 同一品牌内先按 `上市日期` 升序。
- 同一上市日期内按 `产品系列归属` 分组，同系列产品排在一起。
- 同一系列内按钉钉表格返回顺序，不再按新品名称排序。

日期规则：

- 默认按“周六到周五”取数；周五运行会导出上周六到本周五，周一运行会导出上一个完整周六到周五。
- 例如 2026-05-15 周五运行默认取 `2026-05-09` 到 `2026-05-15`；2026-05-18 周一运行也默认取 `2026-05-09` 到 `2026-05-15`。
- 可以用 `--startDate 20260509` 和 `--endDate 20260515` 覆盖。

品类规则：

- 品类只显示一个。
- 优先级：`茶特调 > 果蔬茶 > 柠檬茶 > 酒元素 > 主品类`。
- 先看钉钉字段 `附加品类` 是否命中前四个标签，命中多个时按优先级只取一个。
- 都未命中时显示 `主品类`。
- 后续优先级变化时，改 `config/report_rules.json` 的 `categoryRule.priority`。

备注规则：

- Excel 备注列按 `回归`、`联名`、`备注` 顺序拼接。
- `回归` 只显示“回归”，不显示“上新”。
- 钉钉 `联名` 字段只写品牌/IP，Excel 输出时自动补成 `{品牌IP}联名`。
- 非空项用中文逗号 `，` 分隔；都为空显示 `/`。

价格规则：

- 价格文本中括号里的价格加删除线，例如 `15元(17元)(中杯)` 只把 `17元` 加删除线。
- 只要价格片段出现在括号内就会加删除线，例如 `18元(20元)(中杯)/22元(24元)(大杯)` 中的 `20元` 和 `24元` 都会加删除线。
- 汇总表价格列如果需要换行，只在 `/` 后插入一个换行，不在普通文字中间断开。
- 脚本会先清理价格里的连续空行，避免换行后出现空白行。

日期规则：

- 汇总表“上市时间”默认显示为 `5月4日` 这类中文日期。
- 如需调整，改 `config/excel_layout.json` 的 `summary.dateFormat`，再让 Codex 同步脚本规则。

Logo 和页面底色：

- 顶部 logo 来自 `assets/logo/report_logo.png`。
- logo 默认按 `B2:H2` 区域视觉居中，高度约 `2cm`；该区域不使用横向合并单元格。
- 如需替换 logo，保持文件名不变即可，或修改 `config/excel_layout.json` 的 `logo.path`。
- logo 是内部专用资产，已在 `.gitignore` 中忽略，不要提交到 git 或上传外部平台。
- A:I 的无业务底纹区域会统一填充白色，并应用配置字体，避免 Excel 回退到宋体/默认字体。

列宽规则：

- `B/C/D/F` 为固定列宽，配置在 `config/excel_layout.json` 的 `columnWidthRules.fixedColumnsChars`。
- `E` 会根据本次输出的最长新品名称自动调整，目标是不换行。
- `G` 会根据价格中 `/` 前后的最长价格块自动调整；只要价格含 `/`，就只在 `/` 后换行，不切开价格块。
- `H` 自动补足 `B:H` 的剩余宽度；如果 `E/G` 太宽，会优先保证 `E/G` 可读，并允许整体宽度超过目标值。
- 标题、品牌明细标题、明细值区使用跨列居中；第 5 行说明使用左对齐、不换行，让文字自然溢出到右侧空单元格。
- 关注品牌说明行使用 `B:H` 合并单元格，左对齐、不预先换行。
- 除关注品牌说明行外，不使用横向合并单元格；汇总表主体只保留品牌列和新品数量列的纵向合并。

图片和行高：

- 产品外观图片高度统一 `4cm`。
- 图片会在“产品外观”行的 `C:H` 区域内居中；行高会在 4cm 图片高度外额外加内边距，避免图片挡住边框。
- 汇总表主要按 `G` 价格列和 `H` 备注列估算行高，保证多行价格和备注显示完整。
- 明细表中 `新品名称`、`产品系列归属`、`产品价格` 默认保持紧凑单行；`原料构成` 只有文本超出 `C:H` 宽度时才增高。
- `产品卖点介绍` 和 `原料构成` 会清理钉钉底表里的硬换行，整理为连续段落；多段拼接时，若上一段没有结尾标点，会自动补中文句号。
- `产品卖点介绍` 按 `C:H` 宽度保守估算行高，优先保证文字完整显示。
- 产品外观图片下载到系统临时目录用于写入 Excel，保存完成后自动清理，不会在 `outputs/` 下保留 `_image_cache`。

数据质量提醒：

- 生成时会检查本次实际输出记录的 `品类`、`新品名称`、`上市日期`、`产品价格`、`产品卖点介绍`、`原料构成`。
- 上述文字字段为空时，脚本会在命令行/Agent 输出里提醒品牌、新品、上市日期、recordId 和缺失字段名，但不会中断生成。
- `产品系列归属` 为空时显示 `/`，不作为缺失提醒。
- `备注` 为空时显示 `/`，这是有效占位，不作为缺失提醒。
- `产品外观` 只检查图片附件是否缺失或下载失败；如果提醒缺图，请回钉钉 AI 表检查对应记录的产品外观图片。
- 这些检查项配置在 `config/report_rules.json` 的 `dataQualityRule`，后续可以用自然语言要求 Codex 调整。

## 用自然语言调整配置

后续可以直接让 Codex 按自然语言修改配置。建议这样说：

- “把默认品牌加上沪上阿姨。”
- “品类优先级改成柠檬茶 > 果蔬茶 > 茶特调 > 主品类。”
- “产品图片高度改成 5cm。”
- “卖点介绍行高上限放宽到 120。”
- “C/D/F 固定列宽分别改成 8、7、9。”
- “H 列最小宽度改成 14。”
- “顶部 logo 改成 assets/logo/new_logo.png。”
- “上市时间改成 2026年5月4日 这种格式。”
- “产品价格行改回产品属性。”
- “钉钉字段‘主品类’改名叫‘核心品类’，同步字段映射。”

一般对应关系：

- 钉钉字段名、字段 ID、取值方式：改 `config/field_mapping.json`
- 日期、品牌、品类、备注、价格等业务规则：改 `config/report_rules.json`
- 字体、颜色、列宽、行高、图片高度、写入位置：改 `config/excel_layout.json`
- 字体文件路径：改 `config/font_files.json`

调整后建议运行：

```bash
python scripts/generate_weekly_report.py --validate-config
python scripts/generate_weekly_report.py --explain-config
```
