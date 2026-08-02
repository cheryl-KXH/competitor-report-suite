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

## 2. 安装新品周报模块

进入模块并创建独立的 Python 虚拟环境：

```bash
cd new-product-monitor
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
pip install -r requirements.txt
python -m playwright install chromium
```

Windows PowerShell 激活虚拟环境时使用：

```powershell
.venv\Scripts\Activate.ps1
```

创建本地钉钉配置：

```bash
cp config/dingtalk.example.json config/dingtalk.json
```

打开 `config/dingtalk.json`，根据自己的钉钉 AI 表填写连接信息。真实 Streamable HTTP URL 可能包含访问令牌，不要提交到 GitHub。

校验配置：

```bash
python scripts/generate_weekly_report.py --validate-config
```

生成 Excel、HTML 和 PNG 周报：

```bash
python scripts/generate_weekly_outputs.py
```

指定日期范围：

```bash
python scripts/generate_weekly_outputs.py \
  --startDate 20260509 \
  --endDate 20260515
```

启动 HTTP 服务：

```bash
uvicorn service.app:app --host 0.0.0.0 --port 8000
```

验证服务：

```bash
curl http://127.0.0.1:8000/health
```

## 3. 安装跟踪反馈报告模块

先退出上一个虚拟环境并返回项目根目录：

```bash
deactivate
cd ..
cd product-feedback-report
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
pip install -r requirements.txt
```

创建本地钉钉配置：

```bash
cp config/dingtalk.example.json config/dingtalk.json
```

打开 `config/dingtalk.json`，填写自己的 `baseId`、`tableId` 和连接信息。

启动 HTTP 服务。因为 8000 端口可能已被新品周报服务占用，这里示例使用 8001：

```bash
uvicorn service.app:app --host 0.0.0.0 --port 8001
```

验证服务：

```bash
curl http://127.0.0.1:8001/health
```

该服务提供的主要接口包括：

- `POST /prepare-product-menu`
- `POST /generate-delivery-tables`
- `POST /generate-report`
- `POST /finalize-report`

请求需要传入钉钉记录 ID，例如：

```bash
curl -X POST http://127.0.0.1:8001/prepare-product-menu \
  -H "Content-Type: application/json" \
  -d '{"recordId":"你的记录ID"}'
```

## 4. 可选：设置服务密钥

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

## 5. 私密配置与内部资源

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

### 端口已被占用

为第二个服务更换端口，例如：

```bash
uvicorn service.app:app --host 0.0.0.0 --port 8001
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
