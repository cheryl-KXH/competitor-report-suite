# 竞品跟踪反馈报告

从钉钉 AI 表按钮触发，读取任务文件夹中的原始数据，分阶段生成：

1. 产品清单及上新日期标注表
2. 美团&饿了么、京东、社媒平台整理数表
3. HTML 竞品跟踪反馈报告

## 目录结构

```text
product-feedback-report/
├── config/
│   ├── dingtalk.example.json
│   ├── dingtalk_docs.example.json
│   ├── field_mapping.json
│   ├── platform_rules.json
│   └── report_rules.json
├── service/
│   ├── app.py
│   ├── dingtalk_docs.py
│   ├── dingtalk_table.py
│   ├── jobs.py
│   └── schemas.py
├── scripts/
│   ├── data_processing.py
│   ├── generate_data_tables.py
│   ├── generate_report.py
│   ├── prepare_product_menu.py
│   ├── report_html.py
│   └── validate_config.py
└── outputs/
```

## 配置

复制模板后填真实配置：

```bash
cp config/dingtalk.example.json config/dingtalk.json
cp config/dingtalk_docs.example.json config/dingtalk_docs.json
```

`config/dingtalk.json` 和 `config/dingtalk_docs.json` 包含本地私密 token 或 MCP URL，不要提交。

`/generate-delivery-tables` 需要钉钉文档 MCP 来下载附件、解析文件所在目录并上传统计结果。部署前必须设置 `DINGTALK_DOCS_MCP_URL`，或在 mcporter 中注册 `dingtalk-docs`；可运行 `python scripts/validate_config.py` 检查配置提醒。

## 统一入口和按钮接口

钉钉自动化 HTTP 请求分别配置：

```text
POST /generate-weekly-report
POST /prepare-product-menu
POST /generate-delivery-tables
POST /generate-report
```

“统计外卖数据/开始统计”按钮应向统一入口发送：

```http
POST /generate-delivery-tables
Content-Type: application/json

{"recordId":"当前行记录 ID"}
```

任务会读取当前行的外卖数据和产品清单，解析它们所在的钉钉父文件夹，将 `美团外卖数据.xlsx`、`饿了么外卖数据.xlsx`、`京东外卖数据.xlsx` 上传到同一文件夹，再以附件格式回填当前行的“美团”“饿了么”“京东”字段。若无法解析同一文件夹，任务会停止并将原因写入反馈信息，不会静默上传到其他目录。

如果当前项目和周报项目放在同一个大文件夹下，推荐启动统一入口：

```bash
uvicorn gateway.app:app --host 0.0.0.0 --port 8000
```

统一入口通过子进程分别调用周报项目和反馈报告项目，避免两个项目都叫 `service` 时出现 import 冲突。

## 本地调试

可以先不连钉钉，直接用样例文件夹测试文件处理：

```bash
python scripts/prepare_product_menu.py --record-id local-test --input-dir "/path/to/样例文件夹"
python scripts/generate_delivery_tables.py --record-id local-test --input-dir "/path/to/样例文件夹" --annotation outputs/local-test/产品清单及上新日期.xlsx
python scripts/generate_report.py --record-id local-test
```

`产品清单及上新日期.xlsx` 只需要核对 `近32日上新日期`。未填写日期的产品按完整 30 天计算。

外卖销量统计口径：

- 同一平台、品牌、门店、产品存在多次抓取时，只使用最新抓取记录；同日重复时使用最后一条。
- 美团、饿了么逐门店计算在售天数：`min(抓取日期 - 上架日期, 30)`，同日上架按 1 天，未填写上架日期按 30 天。产品日店均为总销量除以各有售门店在售天数之和；结果同时展示在售门店数和产品清单中标注的上新日期。
- 京东的门店销量是品牌总销量快照，产品总销量按各有售门店最新展示销量的平均值计算。
- 京东结果末列展示产品清单中按平台、品牌、产品精确匹配的上新日期。
- 月销量为 0 仍视为有售；缺少门店名称或上架日期晚于抓取日期时停止生成并提示具体异常数据。
- 三个平台都按各自口径下的总销量降序排名，并按总销量计算平台内占比；平台无数据时生成仅含表头的空表。
