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

`/generate-delivery-tables` 和 `/generate-report` 使用 AI 表附件下载地址读取原始文件，钉钉文档 MCP 只用于创建归档目录并上传生成结果。部署前必须设置 `DINGTALK_DOCS_MCP_URL`，或在 mcporter 中注册 `dingtalk-docs`。

配置检查命令：

```bash
python -m scripts.tools.validate_config
```

报告生成前，外卖数据、可选产品清单和各社媒原始数据必须保持为 AI 表“+”上传的单个 `.xlsx` 附件。生成任务不使用钉钉文档中的中间数表，也不复用本机同记录缓存。

AI 表“外卖数据”字段支持直接点击“+”从电脑上传单个 `.xlsx`。若当前行没有可复用的钉钉文档父文件夹，任务会读取“报告日期”“竞品品牌”“关注新品”，在 `localUploadRootFolderId` 指定的根目录下创建并复用 `{YYYY}年/{YYYY}年{M}月/{竞品品牌}：{关注新品}`；未配置根节点 ID 时，按 `localUploadRootFolderName`（默认 `原始文件：竞品新品跟踪反馈`）查找既有根目录。关注新品中的英文逗号和中文逗号会统一显示为 `、`。

## 统一入口和按钮接口

钉钉自动化 HTTP 请求分别配置：

```text
POST /generate-weekly-report
POST /prepare-product-menu
POST /generate-delivery-tables
POST /generate-consumer-feedback
POST /generate-report
POST /finalize-report
```

“统计外卖数据/开始统计”按钮应向统一入口发送：

```http
POST /generate-delivery-tables
Content-Type: application/json

{"recordId":"当前行记录 ID"}
```

任务直接读取 AI 表上传的外卖数据。产品清单有人工上传附件时优先使用；为空时在本次临时目录自动标注，不回填、不保存为同记录缓存。有数据平台的 `品牌-YYYYMMDD-美团.xlsx`、`品牌-YYYYMMDD-饿了么.xlsx`、`品牌-YYYYMMDD-京东.xlsx` 会立即归档并回填结果链接；原始附件在最终确认前保持不变。

“消费者反馈”按钮使用：

```http
POST https://iodine-obtain-president.ngrok-free.dev/generate-consumer-feedback
Content-Type: application/json

{"recordId":"当前行记录 ID"}
```

任务读取当前行的微博、小红书、抖音、B站上传附件，大众点评暂不参与。四个平台允许缺失；社媒统计结果立即归档并回填“跟踪报告”，但原始社媒附件不在生成阶段归档或覆盖。

如果当前行某个平台的 Excel 在“产品名称”列中同时包含同品牌、同一“30日”的多款新品，任务会先按产品拆成 `{品牌}-{新品}-{平台}.xlsx`：当前产品文件覆盖当前行附件，其他产品文件覆盖对应产品行的同平台附件，然后为所有受影响产品分别生成社媒评论统计表。产品名按全半角、大小写、空格和常见分隔符规范化后精确匹配，不使用简称或包含匹配；缺少“产品名称”列、存在空产品名、未知产品、重复产品记录或文件不包含当前产品时会在附件回写前停止。

社媒统计工作簿与外卖数表一样，直接使用 Python 和 `openpyxl` 生成，不依赖 Node.js，也不会生成额外的渲染预览文件。

“报告生成”按钮调用 `POST /generate-report`。任务在同一次运行中从原始附件生成外卖数表、社媒统计表和 `{品牌}：{产品1}、{产品2} YYYYMMDD.html`。HTML 直接消费本次的内存统计模型，不再下载或读取刚生成的中间数表。

报告生成期间分为 6 步：`1/6 正在读取品牌、新品及报告日期`、`2/6 正在读取上传的外卖及社媒数据`、按产品清单情况显示 `3/6 正在自动标注在售不满30天的产品上新日期` 或 `3/6 正在读取人工配置的产品清单`、`4/6 正在进行外卖及社媒数据统计`、`5/6 正在生成跟踪反馈报告`、`6/6 正在上传至钉钉文档`。缺失京东或部分社媒平台时，最终反馈例如 `报告已生成（京东外卖无数据，某新品小红书、B站无数据）`；缺失美团或饿了么有效数据会停止生成。

“报告检查完毕”按钮调用 `POST /finalize-report`。它会先检查最终报告链接，再归档外卖原始数据、实际上传的产品清单和匹配的社媒原始附件。文件分别命名为 `{品牌}-{最新抓取日期}.xlsx`、`{品牌}-{最新抓取日期}-产品清单.xlsx`和 `{品牌}-{新品}-{平台}.xlsx`，全部上传成功后再回写原字段。已是钉钉文档链接的字段会跳过，脚本不发送 DING 或群消息。

HTML 的 Logo、字体和产品图片均以内嵌 data URI 保存，不依赖本机路径或临时图片 URL。排版在 `config/html_layout.json` 中维护，字体在 `config/font_files.json` 中维护；反馈报告使用自己的 `assets/logo` 和 `assets/fonts`，不再跨目录读取周报项目资源。两个项目当前各自保留一份相同的周报 Logo、方正FW筑紫黑简和 Heytea Sans Serif。缺失周报产品信息、产品图片或某个新品的社媒统计时，报告继续生成并在内容和反馈信息中提示；美团或饿了么核心数表缺失时停止生成。京东无数据时显示“品牌京东外卖销售数据暂无法获取”。

输入与旧结果处理规则：

- 提取产品清单时先清空当前行的“产品清单”附件和图片字段；开始统计外卖数据时先清空“美团”“饿了么”“京东”链接字段。旧实体文件不会在任务开始时删除，新文件上传时才删除同名文件及 `(数字)` 副本并完成替换。
- 产品清单提取只使用 AI 表上传的“外卖数据”；外卖统计和报告生成不搜索桌面文件或同记录历史缓存。
- 数表和报告生成不归档或覆盖原始附件字段；只有 `/finalize-report` 会将原始附件替换为钉钉文档链接。

如果当前项目和周报项目放在同一个大文件夹下，推荐启动统一入口：

```bash
uvicorn gateway.app:app --host 0.0.0.0 --port 8000
```

统一入口通过子进程分别调用周报项目和反馈报告项目，避免两个项目都叫 `service` 时出现 import 冲突。

## 本地调试

可以先不连钉钉，直接用样例文件夹测试文件处理：

钉钉 HTTP 调用 `/generate-report` 时，外卖数表、社媒统计表和最终 HTML 报告都在系统临时目录中生成，上传至钉钉文档后自动清理，不会在本机 `outputs/` 留副本。下方命令行生成方式保持现状，未传 `--output-dir` 时仍写入本机 `outputs/`。

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
