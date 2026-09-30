# deliverable-qa — 交付物质检

AI 生成的文档越来越多,没人审文档。这是输出侧的质检层:**排版 / AI 腔 / 敏感内容**,
三条检查线,一份问题清单,复查归零才交付。

## 特性

- 纯 Python 标准库实现,零第三方依赖(`zipfile + ElementTree` 直接解析 OOXML)
- 支持 md / txt / docx(pptx 尽力而为),中文场景优先
- 规则与代码解耦:`rules/*.json` 加一条规则 = 加一行数据
- 高危敏感问题(密钥/凭证)不回显原文,避免终端/日志二次泄露
- 用户私有规则(人名映射、内部路径、内网域名)放仓库外本地配置

## 两种用法

| 形态 | 适合 | 安装方式 |
|---|---|---|
| **CLI / Skill** | 任意 agent(ZCode、Claude Code、DSH…)或命令行 | 把 `SKILL.md`、`qa_check.py`、`rules/` 放到 `~/.agents/skills/deliverable-qa/` |
| **DSH 插件** | DeepSeek Harness 用户:原生 `qa_check` 工具 + **交付前自动安检** | `dsh plugin --profile <profile> add dsh-deliverable-qa`(见 [plugin/](plugin/)) |

插件形态多出来的能力:注册原生工具给模型直接调用;在 `present` 声明交付物之前自动质检,
发现密钥/凭证类高危问题时**拦截交付**(可配置为只告警)。详见 [plugin/README.md](plugin/README.md)。

## 安装与使用

```bash
# 直接用
python qa_check.py 报告.docx

# 只跑某条线
python qa_check.py --only sensitive 报告.md

# 带本地规则(人名/内网路径等)
python qa_check.py --config ~/.qa_local.json 报告.md
```

退出码:存在高危问题返回 1,否则 0。

### 作为 agent skill 安装(跨 agent 通用)

把 `SKILL.md`、`qa_check.py`、`rules/` 复制到
`~/.agents/skills/deliverable-qa/` 即可;对话里说「质检这份文档」触发。

## 规则总览(v1)

| 线 | 规则 | 严重度默认 |
|---|---|---|
| 排版 | T1 中文语境半角标点(代码/URL/表格豁免) | 中 |
| 排版 | T2 docx 字体不统一(run 级 eastAsia/ascii) | 中 |
| 排版 | T3 标题层级跳跃 | 中 |
| 排版 | T4 日期格式不统一 | 低 |
| 排版 | T5 数字风格(千分位/量词/单位空格,含 bp/亿/万 等金融单位) | 低 |
| 排版 | T6 中英文空格风格不统一 | 低 |
| AI腔 | A1 破折号密度超标(每千字 2 处,可调) | 中 |
| AI腔 | A2 套话与句式(值得注意的是/不仅…更/综上所述…) | 中 |
| AI腔 | A3 首先/其次/最后三段式(段首与句内两种形态) | 中 |
| AI腔 | A4 加粗标签列表化 | 中 |
| AI腔 | A5 三并列短句排比 | 低 |
| AI腔 | A6 空洞收尾(未来可期/值得期待/必将…) | 低 |
| 敏感 | S1 密钥/凭证(AKIA、sk-、私钥头、ghp_ 等;不回显) | 高 |
| 敏感 | S2 个人信息(手机号/身份证;显示打码) | 中 |
| 敏感 | S3 自定义(人名映射/内部路径/内网域名/敏感词,本地配置) | 中 |

## 开源红线声明

本仓库为 clean-room 项目:全部代码与示例从零编写;示例中的人名(张三/李四/王经理)、
路径(`D:\company\...`)、域名均为合成数据;用户真实的人名与内部路径配置放在仓库外,
通过 `--config` 传入,永远不进入本仓库。

## 路线图

- pptx 解析增强;更多排版规则(表格/图注编号)
- 规则包社区化(行业词表)
- DSH 插件封装(交付 present 钩子 = 交付即安检)

## 许可

MIT
