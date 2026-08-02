# Competitor Report Suite

竞品报告自动化工具集，用于从钉钉 AI 表读取数据，生成竞品新品周报和竞品跟踪反馈报告。

## 项目组成

| 目录 | 用途 |
| --- | --- |
| `new-product-monitor/` | 生成竞品新品周报，支持 Excel、HTML 和 PNG 输出，并提供 FastAPI 接口 |
| `product-feedback-report/` | 整理外卖、京东和社媒数据，生成竞品跟踪反馈报告，并提供 FastAPI 接口 |
| `gateway/` | 项目网关相关文件 |
| `output/` | 生成结果目录；通常不应提交到 Git |

每个核心模块还有自己的 README，包含更详细的业务规则和命令说明：

- [新品周报说明](new-product-monitor/README.md)
- [跟踪反馈报告说明](product-feedback-report/README.md)

## 环境要求

建议准备：

- macOS 或 Linux
- Git
- Python 3.10 或更高版本
- pip
- 钉钉 AI 表访问配置
- 已安装并配置的 `mcporter`
- Google Chrome（生成 PNG 长图时使用）
- ngrok（需要让钉钉从公网调用本机服务时使用）

## 1. 克隆项目

```bash
git clone https://github.com/cheryl-KXH/competitor-report-suite.git
cd competitor-report-suite
```

## 2. 安装统一运行环境

`gateway` 会使用当前 Python 解释器启动两个报告项目的任务，因此推荐在仓库根目录只创建一个共享虚拟环境，并把两个项目的依赖都安装进去：

```bash
cd ~/Desktop/competitor-report-suite
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
pip install -r new-product-monitor/requirements.txt
pip install -r product-feedback-report/requirements.txt
python -m playwright install chromium
```

Windows PowerShell 激活虚拟环境时使用：

```powershell
.venv\Scripts\Activate.ps1
```

以后启动 Gateway 或执行任一项目的命令前，都先在仓库根目录激活这个环境：

```bash
cd ~/Desktop/competitor-report-suite
source .venv/bin/activate
```

## 3. 配置新品周报模块

从仓库根目录复制配置模板：

```bash
cp new-product-monitor/config/dingtalk.example.json \
  new-product-monitor/config/dingtalk.json
```

打开 `new-product-monitor/config/dingtalk.json`，根据自己的钉钉 AI 表填写连接信息。真实 Streamable HTTP URL 可能包含访问令牌，不要提交到 GitHub。

### 生成新品周报：完整命令

以下命令都要在 `new-product-monitor/` 目录中运行，并确保虚拟环境已激活：

```bash
cd ~/Desktop/competitor-report-suite/new-product-monitor
source .venv/bin/activate
```

先校验并查看当前配置：

```bash
python scripts/generate_weekly_report.py --validate-config
python scripts/generate_weekly_report.py --explain-config
```

#### 默认生成全部格式

不带参数时，统一入口会生成 Excel、HTML 和 PNG：

```bash
python scripts/generate_weekly_outputs.py
```

默认结果位于：

```text
new-product-monitor/outputs/竞品新品周报YYYY-MM-DD_YYYY-MM-DD/
```

同名目录已经存在时会自动添加 `_2`、`_3`，不会覆盖旧结果。

#### 指定生成格式

```bash
# 只生成 Excel
python scripts/generate_weekly_outputs.py --output-mode excel

# 只生成 HTML
python scripts/generate_weekly_outputs.py --output-mode html

# 只生成 PNG 长图
python scripts/generate_weekly_outputs.py --output-mode image

# 明确生成全部格式
python scripts/generate_weekly_outputs.py --output-mode all
```

#### 指定报告日期范围

日期格式必须为 `YYYYMMDD`：

```bash
python scripts/generate_weekly_outputs.py \
  --startDate 20260509 \
  --endDate 20260515
```

不传日期时，脚本按“周六到周五”的规则计算最近完整周期。

#### 按业务生成

支持的业务包括 `喜茶`、`野萃山` 和 `茶坊`：

```bash
python scripts/generate_weekly_outputs.py --business 喜茶

python scripts/generate_weekly_outputs.py \
  --business 野萃山 \
  --startDate 20260509 \
  --endDate 20260515 \
  --output-mode all
```

#### 手动指定竞品品牌

品牌之间使用英文逗号分隔：

```bash
python scripts/generate_weekly_report.py \
  --brands 霸王茶姬,古茗,乐乐茶
```

同时指定日期和品牌：

```bash
python scripts/generate_weekly_report.py \
  --startDate 20260502 \
  --endDate 20260508 \
  --brands 霸王茶姬,古茗
```

如果同时传入 `--brands` 和 `--business`，脚本优先使用 `--brands`。

#### 只生成 Excel，并指定输出文件

```bash
python scripts/generate_weekly_report.py \
  --startDate 20260502 \
  --endDate 20260508 \
  --output outputs/竞品新品周报2026-05-02_2026-05-08.xlsx
```

通过钉钉按钮或 HTTP 调用生成周报时，不需要单独启动新品周报服务；请使用下方的统一 Gateway。

## 4. 配置跟踪反馈报告模块

从仓库根目录复制两个配置模板：

```bash
cp product-feedback-report/config/dingtalk.example.json \
  product-feedback-report/config/dingtalk.json

cp product-feedback-report/config/dingtalk_docs.example.json \
  product-feedback-report/config/dingtalk_docs.json
```

打开上述配置，填写自己的 `baseId`、`tableId`、钉钉文档目录和连接信息。

检查配置：

```bash
cd ~/Desktop/competitor-report-suite/product-feedback-report
python -m scripts.tools.validate_config
cd ..
```

## 5. 启动统一 Gateway（两个报告共用 8000 端口）

Gateway 统一提供新品周报和反馈报告接口。只启动这一个服务，不需要分别启动两个子项目：

```bash
cd ~/Desktop/competitor-report-suite
source .venv/bin/activate
uvicorn gateway.app:app --host 0.0.0.0 --port 8000
```

保持这个终端窗口运行。打开第二个终端窗口检查：

```bash
curl http://127.0.0.1:8000/health
```

正常响应会包含两个项目的实际目录，例如：

```json
{
  "ok": true,
  "weeklyRoot": "/Users/你的用户名/Desktop/competitor-report-suite/new-product-monitor",
  "feedbackRoot": "/Users/你的用户名/Desktop/competitor-report-suite/product-feedback-report"
}
```

如果 `weeklyRoot` 和 `feedbackRoot` 指向正确目录，说明 Gateway 已经找到两个报告项目。

Gateway 在同一个 8000 端口提供以下接口：

| 接口 | 所属报告 | 用途 |
| --- | --- | --- |
| `POST /generate-weekly-report` | 新品周报 | 根据钉钉记录生成新品周报 |
| `POST /prepare-product-menu` | 反馈报告 | 生成产品清单 |
| `POST /generate-delivery-tables` | 反馈报告 | 生成外卖统计数表 |
| `POST /generate-consumer-feedback` | 反馈报告 | 生成消费者反馈统计 |
| `POST /generate-report` | 反馈报告 | 生成完整反馈报告 |
| `POST /finalize-report` | 反馈报告 | 检查完成后归档原始文件 |

如果需要让钉钉访问本机，再打开第三个终端窗口：

```bash
ngrok http 8000
```

钉钉自动化中的所有接口使用同一个 ngrok 域名，只需要改变最后的路径。

### 通过 Gateway 生成新品周报

将 `你的记录ID` 替换成钉钉 AI 表当前行的真实记录 ID：

```bash
curl -X POST http://127.0.0.1:8000/generate-weekly-report \
  -H "Content-Type: application/json" \
  -d '{"recordId":"你的记录ID"}'
```

### 通过 Gateway 生成反馈报告

下面所有命令同样使用 8000 端口。将 `你的记录ID` 替换成钉钉 AI 表当前行的真实记录 ID。

#### 第一步：生成产品清单

```bash
curl -X POST http://127.0.0.1:8000/prepare-product-menu \
  -H "Content-Type: application/json" \
  -d '{"recordId":"你的记录ID"}'
```

该步骤会生成产品清单供人工检查上新日期。需要修改时，下载自动生成的清单，调整后上传到 AI 表的“产品清单Adjusted”。

#### 第二步：生成外卖统计数表

```bash
curl -X POST http://127.0.0.1:8000/generate-delivery-tables \
  -H "Content-Type: application/json" \
  -d '{"recordId":"你的记录ID"}'
```

运行前，AI 表当前行必须已经上传单个 `.xlsx` 外卖数据附件。

#### 第三步：生成完整反馈报告

```bash
curl -X POST http://127.0.0.1:8000/generate-report \
  -H "Content-Type: application/json" \
  -d '{"recordId":"你的记录ID"}'
```

该步骤会读取当前行的外卖数据、可选的人工调整产品清单和社媒清洗数据，生成统计表及最终 HTML/PDF 报告，并上传到钉钉文档。

#### 第四步：检查完成后归档原始文件

```bash
curl -X POST http://127.0.0.1:8000/finalize-report \
  -H "Content-Type: application/json" \
  -d '{"recordId":"你的记录ID"}'
```

`/finalize-report` 是同步接口。等它返回 `HTTP 200` 后，才代表上传和字段回写全部完成。

#### 使用服务密钥调用

如果启动服务前设置了 `REPORT_SERVICE_SECRET`，每次请求都要传相同的 `secret`：

```bash
curl -X POST http://127.0.0.1:8000/generate-report \
  -H "Content-Type: application/json" \
  -d '{"recordId":"你的记录ID","secret":"你的服务密钥"}'
```

### 不连接钉钉：使用本地文件调试反馈报告

以下命令也要在 `product-feedback-report/` 目录运行。把 `/path/to/...` 替换成真实文件或文件夹路径；路径含空格或中文时保留双引号。

#### 1. 从样例文件夹生成产品清单

```bash
python -m scripts.delivery.prepare_product_menu \
  --record-id local-test \
  --input-dir "/path/to/样例文件夹" \
  --report-date 2026-07-21
```

#### 2. 根据产品清单生成外卖数表

```bash
python -m scripts.delivery.generate_tables \
  --record-id local-test \
  --input-dir "/path/to/样例文件夹" \
  --annotation "outputs/local-test/产品清单及上新日期.xlsx"
```

#### 3. 生成社媒统计数表

```bash
python -m scripts.social.generate_tables \
  --record-id local-test \
  --brand "品牌名称" \
  --product "新品名称" \
  --start-date 6.10 \
  --end-date 7.9 \
  --cleaned-raw-data "/path/to/品牌-YYYYMMDD-新品-社媒清洗数据.xlsx"
```

#### 4. 使用已整理的数表生成最终报告

```bash
python -m scripts.reporting.generate_report \
  --record-id local-test \
  --brand "品牌名称" \
  --products "新品A,新品B" \
  --report-date 2026-07-18 \
  --meituan "/path/to/品牌-20260718-美团.xlsx" \
  --eleme "/path/to/品牌-20260718-饿了么.xlsx" \
  --jd "/path/to/品牌-20260718-京东.xlsx"
```

未传 `--output-dir` 时，本地命令行生成结果默认写入 `product-feedback-report/outputs/`。

## 6. 可选：设置服务密钥

服务支持通过环境变量设置请求密钥：

```bash
export REPORT_SERVICE_SECRET="请替换为安全的随机字符串"
```

设置后，请求体也需要携带相同的 `secret`：

```json
{
  "recordId": "你的记录ID",
  "secret": "请替换为安全的随机字符串"
}
```

如果使用 ngrok 将服务暴露到公网，强烈建议设置密钥。

## 7. 私密配置与内部资源

以下内容不应提交到公开仓库：

- `config/dingtalk.json`
- `config/dingtalk_docs.json`
- `.env`
- Streamable HTTP URL 和访问令牌
- 内部字体、Logo 和其他授权资产
- 自动生成的报告文件

提交前建议执行：

```bash
git status
git diff --cached
```

确认没有密钥、令牌或内部文件后再推送。

## 常见问题

### `ModuleNotFoundError`

确认已经进入正确的模块目录、激活对应的 `.venv`，并执行：

```bash
pip install -r requirements.txt
```

### 8000 端口已被占用

先检查是否已有 Gateway 在运行：

```bash
curl http://127.0.0.1:8000/health
```

如果能够返回 `weeklyRoot` 和 `feedbackRoot`，直接使用现有 Gateway，不要重复启动。需要停止当前终端中的服务时按 `Control + C`，然后重新运行：

```bash
uvicorn gateway.app:app --host 0.0.0.0 --port 8000
```

### PNG 无法生成

确认 Chrome 和 Playwright Chromium 已安装：

```bash
python -m playwright install chromium
```

### 钉钉数据读取失败

检查：

1. `config/dingtalk.json` 中的 `baseId` 和 `tableId`
2. Streamable HTTP URL 是否有效
3. `mcporter` 是否已配置并可用
4. 当前账号是否有目标 AI 表的访问权限

## License

本项目暂未声明开源许可证。在添加许可证前，请勿默认复制、分发或用于商业用途。
