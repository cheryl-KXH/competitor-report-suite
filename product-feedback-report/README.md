# 竞品跟踪反馈报告

从钉钉 AI 表按钮触发，读取任务文件夹中的原始数据，分阶段生成：

1. 产品清单及上新日期标注表
2. 美团&饿了么、京东、社媒平台整理数表
3. 参考既有 PDF 版式的自包含 HTML 竞品跟踪反馈报告

## 目录结构

```text
product-feedback-report/
├── config/
│   ├── dingtalk.example.json
│   ├── dingtalk_docs.example.json
│   ├── field_mapping.json
│   ├── font_files.json
│   ├── html_layout.json
│   ├── platform_rules.json
│   └── report_rules.json
├── service/
│   ├── app.py
│   ├── dingtalk_docs.py
│   ├── dingtalk_table.py
│   ├── jobs.py
│   └── schemas.py
├── scripts/
│   ├── delivery/
│   │   ├── processing.py
│   │   ├── generate_tables.py
│   │   └── prepare_product_menu.py
│   ├── social/
│   │   ├── processing.py
│   │   └── generate_tables.py
│   ├── reporting/
│   │   ├── generate_report.py
│   │   └── html.py
│   └── tools/
│       └── validate_config.py
└── outputs/
```

## 配置

复制模板后填真实配置：

```bash
cp config/dingtalk.example.json config/dingtalk.json
cp config/dingtalk_docs.example.json config/dingtalk_docs.json
```

`config/dingtalk.json` 和 `config/dingtalk_docs.json` 包含本地私密 token 或 MCP URL，不要提交。

`/generate-delivery-tables` 需要钉钉文档 MCP 来下载附件、解析文件所在目录并上传统计结果。部署前必须设置 `DINGTALK_DOCS_MCP_URL`，或在 mcporter 中注册 `dingtalk-docs`。

配置检查命令：

```bash
python -m scripts.tools.validate_config
```

钉钉节点下载顺序为：节点版本一致的任务缓存、钉钉文档 MCP。MCP 未返回下载 URL 时任务会明确失败。

AI 表“外卖数据”字段支持直接点击“+”从电脑上传单个 `.xlsx`。若当前行没有可复用的钉钉文档父文件夹，任务会读取“报告日期”“竞品品牌”“关注新品”，在 `localUploadRootFolderId` 指定的根目录下创建并复用 `{YYYY}年/{YYYY}年{M}月/{竞品品牌}：{关注新品}`；未配置根节点 ID 时，按 `localUploadRootFolderName`（默认 `原始文件：竞品新品跟踪反馈`）查找既有根目录。关注新品中的英文逗号和中文逗号会统一显示为 `、`。

## 统一入口和按钮接口

钉钉自动化 HTTP 请求分别配置：

```text
POST /generate-weekly-report
POST /prepare-product-menu
POST /generate-delivery-tables
POST /generate-consumer-feedback
POST /generate-report
```

“统计外卖数据/开始统计”按钮应向统一入口发送：

```http
POST /generate-delivery-tables
Content-Type: application/json

{"recordId":"当前行记录 ID"}
```

任务会读取当前行的外卖数据和产品清单，统一定位到 `{品牌}：{产品名1}、{产品名2} YYYYMMDD` 带报告日期的标准任务文件夹。产品清单若位于旧的无日期文件夹，会复制归档到标准文件夹并以链接格式回填；旧目录中的文件不自动删除。有数据平台的 `品牌-YYYYMMDD-美团.xlsx`、`品牌-YYYYMMDD-饿了么.xlsx`、`品牌-YYYYMMDD-京东.xlsx` 上传到该文件夹，再以链接格式回填对应字段。品牌和数表日期分别取外卖原始数据中的主品牌和最新抓取日期。无数据平台不上传数表，链接单元格保持空白。重新运行时复用同一归档目录和确定性文件名，上传前删除目录中的同名文件及 `(数字)` 副本，再回填新文件链接。

“消费者反馈”按钮使用：

```http
POST https://iodine-obtain-president.ngrok-free.dev/generate-consumer-feedback
Content-Type: application/json

{"recordId":"当前行记录 ID"}
```

任务读取当前行的微博、小红书、抖音、B站字段，大众点评暂不参与。四个平台允许缺失；生成结果仍保留固定平台版块。社媒原始文件和 `{品牌}-{新品}-社媒评论统计.xlsx` 统一按当前行“30日”所在月归档：优先复用该月目录下同时包含品牌名、当前产品名和同一“30日”的文件夹；找不到时，读取 AI 表中同品牌、同一“30日”的全部新品，创建 `{品牌}：{产品名1}、{产品名2} YYYYMMDD` 文件夹，其中日期前保留一个空格。上传成功后再把“跟踪报告”更新为文件 URL。AI 表“+”本地上传附件会统一重命名为 `{品牌}-{新品}-{平台}.xlsx`。已是钉钉文档节点的原始附件，若不在带日期的目标文件夹中或名称不规范，也会复制归档并回填新附件；旧文件不自动删除。

如果当前行某个平台的 Excel 在“产品名称”列中同时包含同品牌、同一“30日”的多款新品，任务会先按产品拆成 `{品牌}-{新品}-{平台}.xlsx`：当前产品文件覆盖当前行附件，其他产品文件覆盖对应产品行的同平台附件，然后为所有受影响产品分别生成社媒评论统计表。产品名按全半角、大小写、空格和常见分隔符规范化后精确匹配，不使用简称或包含匹配；缺少“产品名称”列、存在空产品名、未知产品、重复产品记录或文件不包含当前产品时会在附件回写前停止。

社媒统计工作簿与外卖数表一样，直接使用 Python 和 `openpyxl` 生成，不依赖 Node.js，也不会生成额外的渲染预览文件。

“报告生成”按钮调用 `POST /generate-report`。报告品牌、日期和产品清单直接取当前行的“竞品品牌”“报告日期”“关注新品”；社媒表和周报产品表只用于查找这些关注新品对应的资料，不会向报告追加其他产品。任务读取当前行美团、饿了么、京东数表，以及关注新品的社媒评论统计表和周报产品信息，生成 `{品牌}：{产品1}、{产品2} YYYYMMDD.html`。HTML 上传到同一个带日期的归档文件夹，并回填现有“跟踪反馈报告”字段；“跟踪反馈报告New”暂不写入。

报告生成期间，“反馈信息”依次显示：`1/4 正在读取当前行的品牌、报告日期和关注新品`、`2/4 正在下载外卖数表和关注新品的社媒评论统计表`、`3/4 正在查找关注新品的产品信息并生成跟踪报告`、`4/4 正在上传跟踪报告到钉钉文档`。完全成功时显示“报告已生成”；存在非阻断缺失项时在其后列出产品和缺失资料；失败时显示发生失败的阶段和具体原因。

HTML 的 Logo、字体和产品图片均以内嵌 data URI 保存，不依赖本机路径或临时图片 URL。排版在 `config/html_layout.json` 中维护，字体在 `config/font_files.json` 中维护；默认复用同级 `new-product-monitor` 的周报 Logo、方正FW筑紫黑简和 Heytea Sans Serif。缺失周报产品信息、产品图片或某个新品的社媒统计时，报告继续生成并在内容和反馈信息中提示；美团或饿了么核心数表缺失时停止生成。京东无数据时显示“品牌京东外卖销售数据暂无法获取”。

输入与旧结果处理规则：

- 提取产品清单时先清空当前行的“产品清单”链接字段；开始统计外卖数据时先清空“美团”“饿了么”“京东”链接字段。旧实体文件不会在任务开始时删除，新文件上传时才删除同名文件及 `(数字)` 副本并完成替换。
- 产品清单提取只使用“外卖数据”字段；外卖统计只使用“外卖数据”和“产品清单”字段，不使用本机文件或“所有数据”字段直接生成。
- 如果输入字段是 AI 表从电脑上传的附件，或保存的是可唯一识别的本机单个 `.xlsx`，任务会先将其转存到钉钉文档并回填原字段，再从该字段继续生成。目录、多文件、非 xlsx、附件下载地址缺失或同名匹配不唯一会直接报错。
- 钉钉节点链接始终优先于桌面同名文件。任务缓存只有在 `nodeId`、远端 `updateTime`、文件名和本地摘要全部一致时才可使用；不再根据 `createTime == updateTime` 猜测缓存有效。

如果当前项目和周报项目放在同一个大文件夹下，推荐启动统一入口：

```bash
uvicorn gateway.app:app --host 0.0.0.0 --port 8000
```

统一入口通过子进程分别调用周报项目和反馈报告项目，避免两个项目都叫 `service` 时出现 import 冲突。

## 本地调试

可以先不连钉钉，直接用样例文件夹测试文件处理：

```bash
python -m scripts.delivery.prepare_product_menu --record-id local-test --input-dir "/path/to/样例文件夹"
python -m scripts.delivery.generate_tables --record-id local-test --input-dir "/path/to/样例文件夹" --annotation outputs/local-test/产品清单及上新日期.xlsx
python -m scripts.social.generate_tables --record-id local-test --brand 品牌 --product 新品 --start-date 6.10 --end-date 7.9 --weibo "/path/to/微博.xlsx" --xiaohongshu "/path/to/小红书.xlsx" --douyin "/path/to/抖音.xlsx" --bilibili "/path/to/B站.xlsx"
python -m scripts.reporting.generate_report --record-id local-test --brand 品牌 --products 新品A,新品B --report-date 2026-07-18 --meituan "/path/to/品牌-20260718-美团.xlsx" --eleme "/path/to/品牌-20260718-饿了么.xlsx" --jd "/path/to/品牌-20260718-京东.xlsx"
```

`产品清单及上新日期.xlsx` 只需要核对 `近32日上新日期`。未填写日期的产品按完整 30 天计算。

外卖销量统计口径：

- 同一平台、品牌、门店、产品存在多次抓取时，只使用最新抓取记录；同日重复时使用最后一条。
- 美团、饿了么逐门店计算在售天数：有上新日期时为 `min(抓取日期 - 上新日期 + 1, 30)`，同日上新按 1 天，未填写上新日期按 30 天。产品日店均为总销量除以各有售门店在售天数之和；结果同时展示在售门店数、单个商品的在售天数和产品清单中标注的上新日期。单个商品的在售天数以该商品最新抓取日期计算，最多记 30 天。
- 京东的门店销量是品牌总销量快照，产品总销量按各有售门店最新展示销量的平均值计算。
- 京东结果在上新日期前同样展示单个商品的在售天数。
- 月销量为 0 仍视为有售；缺少门店名称或上架日期晚于抓取日期时停止生成并提示具体异常数据。
- 三个平台都按各自口径下的总销量降序排名，并按总销量计算平台内占比；平台无数据时不上传数表，对应链接字段保持空白。
