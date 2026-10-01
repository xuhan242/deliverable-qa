# dsh-deliverable-qa

DeepSeek Harness 插件:**交付物质检**。文档交付前自动过一遍排版、AI 腔、敏感内容三条线;
发现密钥/凭证类高危问题时,直接拦截 `present` 交付。

检测引擎是纯 Python 标准库脚本(见本仓库根目录 `qa_check.py`),本插件把它接进 DSH:
注册一个 `qa_check` 工具,并在工具执行管道的 `tools/pre-execute` 阶段挂上交付安检门。

## 安装

tarball(当前可用,npm 发布前的正式渠道):

```bash
dsh plugin --profile <你的 profile> add https://github.com/xuhan242/deliverable-qa/releases/download/v0.1.0/dsh-deliverable-qa-0.1.0.tgz
```

npm(发布后):

```bash
dsh plugin --profile <你的 profile> add dsh-deliverable-qa
```

本地路径(开发):

```bash
dsh plugin --profile <你的 profile> add ./plugin
```

## 它做什么

1. **`qa_check` 工具**:模型可以直接调用,对 md/txt/docx/pptx 出问题清单;
2. **交付前安检门**:任何 `present` 调用声明交付物之前,插件先跑一遍质检——
   - 发现高危问题(密钥、凭证、私钥)时**拦截交付**,把问题位置告诉模型;
   - 检测不可用时放行(不阻塞正常交付),只在控制台告警;
3. 同一文件内容未变化时不重复检测(按 mtime + 大小缓存)。

## 配置

在 profile 的 `cordis.patch.yml` 里覆盖(全部可省略):

```yaml
- id: deliverable-qa
  config:
    gate: block          # block 拦截高危 / warn 只告警 / off 关闭交付门
    python: ''           # 指定 Python 解释器;留空自动找 DSH 自带运行时,再退回 PATH
    configPath: ''       # 本地私有规则(人名映射/内部路径/内网域名),建议放仓库外
    timeoutMs: 20000     # 单次质检超时
```

Python 解析顺序:配置的 `python` → DSH 自带运行时(`$DSH_HOME/dsh-runtimes/*/dependencies/python`)→ PATH 上的 `python3`/`python`。

## 私有规则(不进仓库)

把仓库根目录 `examples/config.example.json` 复制到仓库外(如 `~/.qa_local.json`),
填自己的人名映射、内部路径前缀、内网域名、敏感词,然后用 `configPath` 指过去。

## 许可

MIT
