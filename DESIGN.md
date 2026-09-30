# 设计说明(DESIGN)

「交付物质检」对 AI 协作产出的文档做出厂前检查,三条检查线:**排版规范(typography)、AI 腔(aitone)、敏感内容(sensitive)**。本文说明目录结构、规则文件格式与判定口径。

## 1. 目标与非目标

- 目标:输出侧质检层。各类 agent 都在生成文档,本工具在交付前给出一份可逐条修复的问题清单。
- 非目标:不做自动修复(修复由 agent/人编辑文档完成,脚本只检测);不做语义级审校;v1 不支持 PDF。

## 2. 目录结构

```
deliverable-qa/
├── qa_check.py              # 检测脚本(仅 Python 3 标准库,zipfile + ElementTree 解 OOXML)
├── rules/                   # 规则数据(与逻辑解耦:加一条规则 = 加一行数据)
│   ├── typography.json      # 排版线内置规则 T1–T6
│   ├── aitone.json          # AI 腔线内置规则 A1–A6(阈值可调)
│   └── sensitive.json       # 敏感线内置规则 S1–S2;用户自定义项走 --config
├── examples/                # 合成示例(每条内置规则至少埋 2 处问题)
│   ├── report.md            # 合成问题报告(markdown)
│   ├── report.docx          # 同题材 docx(由 build_docx_example.py 生成)
│   ├── build_docx_example.py# docx 生成脚本(生成用 python-docx;qa_check.py 本身零依赖)
│   └── config.example.json  # 本地私有配置的合成示例(真实私有配置放仓库外)
├── selftest.py              # 自测:预期发现数 vs 实际;修复 3 条后复查计数减少;pptx 冒烟
├── SKILL.md                 # 技能说明(触发词、流程纪律)
├── README.md                # 定位、安装、用法、规则表、开源红线声明
└── DESIGN.md                # 本文件
```

## 3. 数据流

```
输入文件(md/txt/docx/pptx)
  → 解析器:统一产出 Block 序列(行 / 段落为最小单元)
  → 规则引擎:rules/*.json 数据 + qa_check.py 内注册的检测器(detector)
  → Issue 列表(文件|位置|检查线|严重度|问题|修改建议)
  → 渲染:markdown 清单(默认)或 JSON(--json,供脚本消费)
```

- Block 字段:`idx`(位置编号)、`text`(原文)、`stext`(风格检查用文本:去行内代码、去 URL)、`is_code`(markdown 代码块)、`heading`(标题层级)、`runs`(docx run 级字体/加粗信息)、`slide`(pptx 幻灯片号)。
- 位置口径:md/txt 报「第 N 行」;docx 报「第 N 段」(document.xml 文档序,含表格内段落);pptx 报「幻灯片 S·第 N 段」。

## 4. 规则文件格式

每条规则是 JSON 数组中的一个对象:

| 字段 | 类型 | 说明 |
| --- | --- | --- |
| id | str | 规则编号,全仓库唯一(T1…T6 / A1…A6 / S1…S2,S3-* 为配置生成) |
| line | str | 检查线:typography / aitone / sensitive |
| name | str | 规则名(输出用) |
| detector | str | 检测器名,对应 qa_check.py 中 `DETECTORS` 注册表 |
| severity | str | 默认严重度:高 / 中 / 低 |
| enabled | bool | 开关 |
| params | object | 检测器参数(词条、阈值、窗口等),均可被 --config 覆盖 |
| description | str | 规则描述 |
| advice | str | 默认修改建议 |

**解耦承诺**:新增一条「词条/模式类」规则(套话、密钥特征、敏感词等)只需在 rules/*.json 加一行;只有引入全新检测逻辑时才需要注册新 detector。

## 5. 检测器注册表(v1)

| detector | 规则 | 参数 | 说明 |
| --- | --- | --- | --- |
| halfwidth_punct | T1 | chars、cjk_window | 中文语境(窗口内含汉字)出现半角标点;豁免:代码块、行内代码、URL、数字间 `,`/`:`(千分位/时间)、盘符 `D:\` |
| font_consistency | T2 | min_runs | 仅 docx:run 级 eastAsia/ascii 字体与全文主流不一致,按段报告 |
| heading_jump | T3 | – | 标题层级向前跳级(差 >1) |
| date_mix | T4 | priority | 同文档混用 ISO/点分/斜线/中文四种日期;报非主流样式 |
| number_style | T5 | thousands、plain_min/max_digits、measures、units | 三个子检查:千分位风格、同量词前中文数字 vs 阿拉伯数字、数字-单位间空格风格 |
| cjk_latin_space | T6 | majority_min | 汉字与英文字母之间加/不加空格须全文一致(v1 不考察汉字-数字之间) |
| dash_density | A1 | per_1000 | 破折号(——)密度,默认阈值每千字 2 处(不含代码块) |
| regex_hits | A2/S1/S2 | patterns、include_code、raw_text、no_echo、mask | 通用词条/正则命中,每命中一处报一条 |
| tri_template | A3 | markers、min_markers | 段首「首先/其次/最后」等标记 ≥2 种时,逐段报告 |
| bold_labels | A4 | min_consecutive | 连续 ≥N 段以「**加粗标签**:」(docx 为加粗 run + 冒号)开头,按组报告 |
| parallel_clauses | A5 | min_clauses、max_clause_len | 单句内 ≥N 个长度 ≤max 的逗号分句(排比腔) |
| hollow_closing | A6 | scan_tail、patterns | 仅检查末尾 scan_tail 个非空段,命中空洞收尾模式 |

`--config` 的 `sensitive_extra` 生成 4 条动态规则:`S3-names`(人名映射,建议替换值)、`S3-paths`(内部路径前缀)、`S3-domains`(内网域名)、`S3-words`(敏感词表,默认高危且不回显)。

## 6. 统计口径与判定规则

一致性类检查(多数派 vs 少数派)统一口径:

| 检查 | 主流判定 | 平局规则 |
| --- | --- | --- |
| T2 字体 | 出现次数最多的 run 级字体 | 先出现者为主流 |
| T4 日期 | 出现次数最多的样式 | 按 priority:iso > dot > slash > cjk |
| T5a 千分位 | 多数风格 | 平局报「带千分位」一方 |
| T5b 中文/阿拉伯数字(按量词) | 多数风格 | 平局报中文数字一方 |
| T5c 单位前空格(按单位) | 多数风格 | 平局报「加空格」一方 |
| T6 中英文空格 | 多数风格(需 ≥majority_min 个样本) | 平局视「加空格」为主流 |

其余口径:字符数 = 去空白后的 stext 长度;密度 = 出现数 × 1000 ÷ 字符数。

## 7. 豁免矩阵

| 检查线 | 代码块 | 行内代码 | URL | 数字间半角标点 |
| --- | --- | --- | --- | --- |
| 排版(T1/T4/T5/T6) | 豁免 | 豁免 | 豁免 | 豁免(`1,234`、`22:00`) |
| AI 腔(A1–A6) | 豁免 | 豁免 | 豁免 | – |
| 敏感(S1–S3) | **不豁免** | 不豁免 | 不豁免 | – |

敏感线不豁免代码块:密钥写在代码里同样是泄露。

## 8. 高危脱敏口径

- S1(疑似密钥/凭证,高危)与 S3-words(本地敏感词,高危):`no_echo`,报告中只报「类型 + 位置」,不回显原文,避免终端/日志二次泄露。
- S2(手机号/身份证,中危):回显但掩码(保留前 3 后 2 位)。
- S3-names/paths/domains(中危):回显命中词,因为修改建议本身需要给出替换目标。

## 9. 退出码

| 退出码 | 含义 |
| --- | --- |
| 0 | 无问题,或只有中/低危问题 |
| 1 | 存在高危(高)问题 |
| 2 | 用法/输入错误(文件不存在、配置非法、无启用规则等) |

## 10. 配置文件格式(--config)

```json
{
  "rules_override": {
    "S2": { "enabled": false },
    "A1": { "params": { "per_1000": 3.0 } },
    "T4": { "severity": "低" }
  },
  "sensitive_extra": {
    "name_map":     { "王经理": "leader" },
    "path_prefixes": ["D:\\company\\reports"],
    "domains":      ["intranet.company.local"],
    "words":        ["某内部代号"],
    "severity":     { "names": "中", "paths": "中", "domains": "中", "words": "高" }
  }
}
```

- `rules_override`:按规则 id 覆盖 enabled / severity,params 做键级合并。
- `sensitive_extra`:用户私有数据,**放仓库外**,仓库只带 examples/config.example.json 合成示例。

## 11. v1 边界与已知局限

- docx 字体检查只看 run 级显式 rFonts,样式继承链(styles.xml 级联)未展开,继承统一记为「(继承默认)」。
- pptx 尽力而为:仅抽取幻灯片文本,不检查字体/母版。
- T6 只考察汉字-英文字母边界,不含数字(日期、编号会引入大量误报)。
- A5/A6 为启发式,严重度定为低;阈值可调,误报优先靠调参而非删规则。
- Windows 优先:路径用 pathlib;输出强制 UTF-8(`sys.stdout.reconfigure`)。
