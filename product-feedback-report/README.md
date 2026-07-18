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

任务会读取当前行的外卖数据和产品清单，解析它们所在的钉钉父文件夹，将有数据平台的 `美团外卖数据.xlsx`、`饿了么外卖数据.xlsx`、`京东外卖数据.xlsx` 上传到同一文件夹，再以附件格式回填对应字段。无数据平台不上传数表，附件单元格保持空白。若无法解析同一文件夹，任务会停止并将原因写入反馈信息，不会静默上传到其他目录。

“消费者反馈”按钮使用：

```http
POST https://iodine-obtain-president.ngrok-free.dev/generate-consumer-feedback
Content-Type: application/json

{"recordId":"当前行记录 ID"}
```

任务只读取当前行的微博、小红书、抖音字段，大众点评和 B站暂不参与。三个平台允许缺失；生成结果仍保留固定平台版块。社媒原始文件和 `{品牌}-{新品}-社媒评论统计.xlsx` 统一按当前行“30日”所在月归档：优先复用该月目录下同时包含品牌名和当前产品名的文件夹；找不到时，读取 AI 表中同品牌、同一“30日”的全部新品，创建 `{品牌}：{产品名1}、{产品名2}` 文件夹。上传成功后再把“跟踪报告”更新为文件 URL。若附件由 AI 表“+”从电脑上传，任务会等待临时下载 URL、转存到钉钉文档并用文档节点附件替换原字段；不会按文件名搜索服务机器上的本机文件。

社媒统计工作簿与外卖数表一样，直接使用 Python 和 `openpyxl` 生成，不依赖 Node.js，也不会生成额外的渲染预览文件。

输入与旧结果处理规则：

- 提取产品清单时先清空当前行的“产品清单”附件字段；开始统计外卖数据时先清空“美团”“饿了么”“京东”附件字段。旧实体文件不会在任务开始时删除，新文件上传时才按同名替换。
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
python -m scripts.social.generate_tables --record-id local-test --brand 品牌 --product 新品 --start-date 6.10 --end-date 7.9 --weibo "/path/to/微博.xlsx" --xiaohongshu "/path/to/小红书.xlsx" --douyin "/path/to/抖音.xlsx"
python -m scripts.reporting.generate_report --record-id local-test
```

`产品清单及上新日期.xlsx` 只需要核对 `近32日上新日期`。未填写日期的产品按完整 30 天计算。

外卖销量统计口径：

- 同一平台、品牌、门店、产品存在多次抓取时，只使用最新抓取记录；同日重复时使用最后一条。
- 美团、饿了么逐门店计算在售天数：有上新日期时为 `min(抓取日期 - 上新日期 + 1, 30)`，同日上新按 1 天，未填写上新日期按 30 天。产品日店均为总销量除以各有售门店在售天数之和；结果同时展示在售门店数和产品清单中标注的上新日期。
- 京东的门店销量是品牌总销量快照，产品总销量按各有售门店最新展示销量的平均值计算。
- 京东结果末列展示产品清单中按平台、品牌、产品精确匹配的上新日期。
- 月销量为 0 仍视为有售；缺少门店名称或上架日期晚于抓取日期时停止生成并提示具体异常数据。
- 三个平台都按各自口径下的总销量降序排名，并按总销量计算平台内占比；平台无数据时不上传数表，对应附件字段保持空白。
