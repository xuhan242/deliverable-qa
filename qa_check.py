#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""交付物质检(qa_check):对 AI 协作产出的文档做出厂前检查。

三条检查线:
  - typography 排版规范
  - aitone      AI 腔
  - sensitive   敏感内容

特性:
  - 仅依赖 Python 3 标准库(zipfile + ElementTree 解析 OOXML);
  - 规则与逻辑解耦:内置规则全部在 rules/*.json,新增同类规则只需加一行数据;
  - 输出 markdown 问题清单(文件|位置|检查线|严重度|问题|修改建议),末尾按检查线汇总;
  - 存在高危问题退出码为 1,否则为 0;
  - 高危敏感问题不回显原文,只报类型与位置,避免终端/日志二次泄露。

用法:
  python qa_check.py 文件1 [文件2 ...] [--config 本地配置.json] [--only typography|aitone|sensitive] [--json]
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

__version__ = '1.0.0'

# ---------- 常量 ----------

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'   # docx 主命名空间
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'         # pptx 文本命名空间
HAN = '\u4e00-\u9fff'                                                  # 汉字区间
SEVERITIES = ('高', '中', '低')
LINE_NAMES = {'typography': '排版', 'aitone': 'AI腔', 'sensitive': '敏感'}

# 反引号用 chr(96) 拼接,避免源码/模板里出现裸反引号
_BT = chr(96)
_EBT = re.escape(_BT)

# URL 字符集(ASCII;遇到中文/全角标点即停止,避免把整句吞进 URL)
_URL_BODY = r"[A-Za-z0-9._~:/?#\[\]@!$&'()*+,;%=-]"
RE_URL = re.compile(r'(?:https?|ftp)://' + _URL_BODY + r'+|www\.' + _URL_BODY + r'+')

# 行内代码:两个及以上反引号包裹,或单个反引号包裹
RE_CODESPAN = re.compile(
    _EBT + '{2,}[^' + _EBT + ']*' + _EBT + '{2,}|' + _EBT + '[^' + _EBT + ']+' + _EBT)

# markdown 代码围栏:三个及以上反引号,或三个及以上波浪号
RE_FENCE = re.compile(r'^\s*(?:' + _EBT + '{3,}|~{3,})')

CN_NUM = '一二两三四五六七八九十百千'   # 中文数字字符(用于量词前数字风格检查)


# ---------- 数据结构 ----------

@dataclass
class RunInfo:
    """docx 的一个 run(文本片段)及其字体/加粗信息。"""
    text: str = ''
    east: str = ''       # w:rFonts/@w:eastAsia(中文字体)
    ascii_: str = ''     # w:rFonts/@w:ascii(西文字体)
    bold: bool = False


@dataclass
class Block:
    """最小检查单元:markdown 的一行 / docx、pptx 的一个段落。"""
    idx: int                 # 位置编号:行号或段落号
    text: str                # 原文
    stext: str = ''          # 风格检查用文本(去行内代码、URL)
    is_code: bool = False    # markdown 代码块(敏感线仍会扫描)
    heading: int = 0         # 标题层级,0 表示非标题
    runs: list = field(default_factory=list)   # docx run 级信息
    slide: int = 0           # pptx 幻灯片号


@dataclass
class Rule:
    rid: str
    line: str
    line_name: str
    name: str
    detector: str
    severity: str = '中'
    enabled: bool = True
    params: dict = field(default_factory=dict)
    description: str = ''
    advice: str = ''


@dataclass
class Issue:
    file: str
    loc: str
    line_name: str
    rid: str
    severity: str
    message: str
    advice: str
    sort_key: tuple = (0, 0, 0, '')


# ---------- 文本工具 ----------

def read_text_guess(path: Path) -> str:
    """按常见编码依次尝试读取文本文件。"""
    for enc in ('utf-8-sig', 'utf-8', 'gb18030'):
        try:
            return path.read_text(encoding=enc)
        except UnicodeDecodeError:
            continue
    return path.read_text(encoding='utf-8', errors='replace')


def style_sanitize(text: str) -> str:
    """排版/AI 腔检查用文本:去掉行内代码与 URL(豁免区)。"""
    t = RE_CODESPAN.sub(' ', text)
    t = RE_URL.sub(' ', t)
    return t


def has_han(text: str) -> bool:
    return bool(re.search('[' + HAN + ']', text))


def _mask(s: str) -> str:
    """敏感内容掩码:保留前 3 后 2 位。"""
    return s[:3] + '****' + s[-2:] if len(s) > 6 else '****'


# ---------- 文件解析 ----------

def parse_md_txt(path: Path, is_md: bool):
    """markdown / 纯文本:每行一个 Block;markdown 识别代码围栏与标题层级。"""
    blocks = []
    in_code = False
    for i, line in enumerate(read_text_guess(path).splitlines(), 1):
        if is_md and RE_FENCE.match(line):
            blocks.append(Block(idx=i, text=line, is_code=True))
            in_code = not in_code
            continue
        if in_code:
            blocks.append(Block(idx=i, text=line, is_code=True))
            continue
        b = Block(idx=i, text=line, stext=style_sanitize(line))
        if is_md:
            m = re.match(r'^(#{1,6})\s+(.*)$', line)
            if m:
                b.heading = len(m.group(1))
        blocks.append(b)
    return blocks, ('md' if is_md else 'txt')


def _docx_heading_styles(z: zipfile.ZipFile) -> dict:
    """从 styles.xml 建立 styleId → 标题层级 映射(兼容中文 Word 的「标题 1」)。"""
    mapping = {}
    try:
        sm = ET.fromstring(z.read('word/styles.xml'))
    except KeyError:
        return mapping
    for st in sm.findall(W + 'style'):
        sid = st.get(W + 'styleId') or ''
        nm = st.find(W + 'name')
        name = (nm.get(W + 'val') if nm is not None else '') or ''
        m = re.match(r'(?i)^heading\s*(\d)', name) or re.match(r'^标题\s*(\d)', name)
        if m:
            mapping[sid] = int(m.group(1))
    return mapping


def parse_docx(path: Path):
    """docx:按 document.xml 文档序(含表格内段落)逐段建立 Block,保留 run 级信息。"""
    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read('word/document.xml'))
        style_levels = _docx_heading_styles(z)
    body = root.find(W + 'body')
    if body is None:
        raise ValueError('docx 缺少 body')
    blocks = []
    idx = 0
    for p in body.iter(W + 'p'):
        idx += 1
        runs = []
        for r in p.findall(W + 'r'):
            text = ''.join(t.text or '' for t in r.findall(W + 't'))
            if not text:
                continue
            east = ascii_ = ''
            bold = False
            rpr = r.find(W + 'rPr')
            if rpr is not None:
                rf = rpr.find(W + 'rFonts')
                if rf is not None:
                    east = rf.get(W + 'eastAsia') or ''
                    ascii_ = rf.get(W + 'ascii') or ''
                bel = rpr.find(W + 'b')
                if bel is not None and bel.get(W + 'val') not in ('false', '0', 'none'):
                    bold = True
            runs.append(RunInfo(text, east, ascii_, bold))
        text = ''.join(rr.text for rr in runs)
        heading = 0
        ppr = p.find(W + 'pPr')
        if ppr is not None:
            ps = ppr.find(W + 'pStyle')
            if ps is not None:
                v = ps.get(W + 'val') or ''
                if v in style_levels:
                    heading = style_levels[v]
                else:
                    m = re.match(r'(?i)^heading\s*(\d+)$', v) or re.match(r'^([1-9])$', v)
                    if m:
                        heading = int(m.group(1))
        blocks.append(Block(idx=idx, text=text, stext=style_sanitize(text),
                            heading=heading, runs=runs))
    return blocks, 'docx'


def parse_pptx(path: Path):
    """pptx(尽力而为):抽取每张幻灯片的段落文本,不做字体/母版分析。"""
    blocks = []
    with zipfile.ZipFile(path) as z:
        names = [n for n in z.namelist() if re.match(r'^ppt/slides/slide\d+\.xml$', n)]
        names.sort(key=lambda n: int(re.search(r'(\d+)', n).group(1)))
        for n in names:
            sno = int(re.search(r'(\d+)', n).group(1))
            root = ET.fromstring(z.read(n))
            i = 0
            for para in root.iter(A + 'p'):
                i += 1
                text = ''.join(t.text or '' for t in para.iter(A + 't'))
                blocks.append(Block(idx=i, text=text, stext=style_sanitize(text), slide=sno))
    return blocks, 'pptx'


def parse_file(path: Path):
    ext = path.suffix.lower()
    if ext in ('.md', '.markdown'):
        return parse_md_txt(path, True)
    if ext == '.txt':
        return parse_md_txt(path, False)
    if ext == '.docx':
        return parse_docx(path)
    if ext == '.pptx':
        return parse_pptx(path)
    raise ValueError('不支持的扩展名 %s(支持 md/txt/docx/pptx)' % ext)


# ---------- 检查上下文 ----------

class Ctx:
    """单个被检文件的上下文。"""

    def __init__(self, path: Path, display: str, kind: str, blocks: list, file_index: int = 0):
        self.path = path
        self.display = display
        self.kind = kind
        self.blocks = blocks
        self.file_index = file_index
        self._style = [b for b in blocks if not b.is_code and b.text.strip()]

    @property
    def style_blocks(self) -> list:
        """排版/AI 腔检查范围:排除代码块与空块。"""
        return self._style

    @property
    def all_blocks(self) -> list:
        """敏感线检查范围:包含代码块(密钥写在代码里同样是泄露)。"""
        return self.blocks

    def loc(self, b: Block) -> str:
        if self.kind == 'pptx':
            return '幻灯片%d·第%d段' % (b.slide, b.idx)
        if self.kind == 'docx':
            return '第%d段' % b.idx
        return '第%d行' % b.idx

    def char_count(self) -> int:
        return sum(len(re.sub(r'\s', '', b.stext)) for b in self._style)

    def issue(self, b: Block, rule: Rule, message: str, advice: str = None) -> Issue:
        return Issue(file=self.display, loc=self.loc(b), line_name=rule.line_name,
                     rid=rule.rid, severity=rule.severity, message=message,
                     advice=advice if advice is not None else rule.advice,
                     sort_key=(self.file_index, b.slide, b.idx, rule.rid))


# ---------- 检测器(规则数据引用这里的名字) ----------

def det_halfwidth_punct(ctx: Ctx, rule: Rule):
    """T1:中文语境半角标点。"""
    chars = rule.params.get('chars', ',;:!?\'"')
    window = int(rule.params.get('cjk_window', 6))

    def exempt(t: str, i: int, ch: str) -> bool:
        prev = t[i - 1] if i > 0 else ''
        nxt = t[i + 1] if i + 1 < len(t) else ''
        if ch in ',:' and prev.isdigit() and nxt.isdigit():
            return True                      # 千分位 / 时间 / 比例
        if ch == ':' and prev.isascii() and prev.isalpha() and nxt and nxt in '\\/':
            return True                      # Windows 盘符、URL 残余
        if ch in '\'"' and prev.isascii() and prev.isalpha() and nxt.isascii() and nxt.isalpha():
            return True                      # 英文缩写/所有格
        return False

    out = []
    for b in ctx.style_blocks:
        t = b.stext
        if not has_han(t):
            continue                          # 非中文语境,跳过
        found = []
        for i, ch in enumerate(t):
            if ch not in chars or exempt(t, i, ch):
                continue
            seg = t[max(0, i - window): i + window + 1]
            if has_han(seg):
                found.append(ch)
        if found:
            c = Counter(found)
            detail = '、'.join('「%s」×%d' % (k, v) for k, v in c.items())
            out.append(ctx.issue(b, rule, '中文语境发现 %d 处半角标点:%s' % (len(found), detail)))
    return out


def det_font_consistency(ctx: Ctx, rule: Rule):
    """T2:docx run 级字体与全文主流不一致。"""
    if ctx.kind != 'docx':
        return []
    min_runs = int(rule.params.get('min_runs', 3))
    all_runs = [r for b in ctx.style_blocks for r in b.runs if r.text.strip()]
    if len(all_runs) < min_runs:
        return []
    main_e = Counter((r.east or '(继承默认)') for r in all_runs).most_common(1)[0][0]
    main_a = Counter((r.ascii_ or '(继承默认)') for r in all_runs).most_common(1)[0][0]
    out = []
    for b in ctx.style_blocks:
        devs = []
        for r in b.runs:
            if not r.text.strip():
                continue
            de = r.east or '(继承默认)'
            da = r.ascii_ or '(继承默认)'
            if de != main_e or da != main_a:
                devs.append('中文「%s」/西文「%s」' % (de, da))
        if devs:
            out.append(ctx.issue(
                b, rule,
                '字体与全文主流(中文「%s」/西文「%s」)不一致:%s' %
                (main_e, main_a, '、'.join(sorted(set(devs))))))
    return out


def det_heading_jump(ctx: Ctx, rule: Rule):
    """T3:标题层级跳跃。"""
    out = []
    prev = 0
    for b in ctx.style_blocks:
        if b.heading > 0:
            if prev and b.heading - prev > 1:
                out.append(ctx.issue(b, rule, '标题层级从 H%d 直接跳到 H%d' % (prev, b.heading)))
            prev = b.heading
    return out


_DATE_PATS = {
    'iso': r'(?<![\d./-])\d{4}-\d{1,2}-\d{1,2}(?![\d./-])',
    'dot': r'(?<![\d.\-])\d{4}\.\d{1,2}\.\d{1,2}(?![\d.])',
    'slash': r'(?<![\d/])\d{4}/\d{1,2}/\d{1,2}(?![\d/])',
    'cjk': r'(?<!\d)\d{1,2}月\d{1,2}日(?!\d)',
}
_DATE_NAMES = {'iso': 'ISO 式(2026-09-30)', 'dot': '点分式(2026.9.30)',
               'slash': '斜线式(2026/9/30)', 'cjk': '中文式(9月30日)'}


def det_date_mix(ctx: Ctx, rule: Rule):
    """T4:同文档混用多种日期格式,报告非主流样式。"""
    prio = rule.params.get('priority', ['iso', 'dot', 'slash', 'cjk'])
    hits = {st: [] for st in _DATE_PATS}
    for b in ctx.style_blocks:
        for st, pat in _DATE_PATS.items():
            for m in re.finditer(pat, b.stext):
                hits[st].append((b, m.group()))
    present = [st for st in prio if hits[st]]
    if len(present) < 2:
        return []
    majority = max(present, key=lambda st: (len(hits[st]), -prio.index(st)))
    out = []
    for st in present:
        if st == majority:
            continue
        for b, s in hits[st]:
            out.append(ctx.issue(b, rule, '日期「%s」为%s,与全文主流%s不一致' %
                                 (s, _DATE_NAMES[st], _DATE_NAMES[majority])))
    return out


def det_number_style(ctx: Ctx, rule: Rule):
    """T5:数字风格(千分位 / 中文数字 vs 阿拉伯数字 / 单位前空格)。"""
    p = rule.params
    out = []
    blocks = ctx.style_blocks

    def find(pat):
        return [(b, m.group()) for b in blocks for m in re.finditer(pat, b.stext)]

    # a) 千分位
    if p.get('thousands', True):
        lo = int(p.get('plain_min_digits', 5))
        hi = int(p.get('plain_max_digits', 8))
        comma = find(r'(?<![\d.,A-Za-z])\d{1,3}(?:,\d{3})+(?![\d.,A-Za-z])')
        plain = find(r'(?<![\d.,A-Za-z])\d{%d,%d}(?![\d.,A-Za-z])' % (lo, hi))
        if comma and plain:
            minority = 'comma' if len(comma) <= len(plain) else 'plain'
            for b, s in (comma if minority == 'comma' else plain):
                out.append(ctx.issue(b, rule, '数字「%s」的千分位风格(%s)与全文主流(%s)不一致' %
                                     (s, '加千分位' if minority == 'comma' else '不加千分位',
                                      '不加千分位' if minority == 'comma' else '加千分位')))

    # b) 同一量词前:中文数字 vs 阿拉伯数字
    for mw in p.get('measures', []):
        esc = re.escape(mw)
        cn = find('[' + CN_NUM + ']{1,4}' + esc)
        ar = find(r'\d+(?:\.\d+)?\s*' + esc)
        if cn and ar:
            minority = 'cn' if len(cn) <= len(ar) else 'ar'
            for b, s in (cn if minority == 'cn' else ar):
                out.append(ctx.issue(b, rule, '量词「%s」前的数字「%s」用%s,与全文主流(%s)不一致' %
                                     (mw, s, '中文数字' if minority == 'cn' else '阿拉伯数字',
                                      '阿拉伯数字' if minority == 'cn' else '中文数字')))

    # c) 数字与单位之间的空格
    for u in p.get('units', []):
        esc = re.escape(u)
        tail = '(?![A-Za-z])' if re.fullmatch(r'[A-Za-z%]+', u) else ''
        sp = find(r'(?<![\d.])\d+(?:\.\d+)?\s+' + esc + tail)
        ti = find(r'(?<![\d.])\d+(?:\.\d+)?' + esc + tail)
        if sp and ti:
            minority = 'sp' if len(sp) <= len(ti) else 'ti'
            for b, s in (sp if minority == 'sp' else ti):
                out.append(ctx.issue(b, rule, '「%s」数字与单位「%s」之间%s,与全文主流(%s)不一致' %
                                     (s, u, '加了空格' if minority == 'sp' else '未加空格',
                                      '不加空格' if minority == 'sp' else '加空格')))
    return out


def det_cjk_latin_space(ctx: Ctx, rule: Rule):
    """T6:汉字与英文字母之间加/不加空格,全文须一致(v1 不考察汉字-数字)。"""
    majority_min = int(rule.params.get('majority_min', 2))
    sp_pat = re.compile('([' + HAN + r'])[ \t]+([A-Za-z])|([A-Za-z])[ \t]+([' + HAN + '])')
    ti_pat = re.compile('([' + HAN + r'])[A-Za-z]|[A-Za-z]([' + HAN + '])')
    sp = [(b, m) for b in ctx.style_blocks for m in sp_pat.finditer(b.stext)]
    ti = [(b, m) for b in ctx.style_blocks for m in ti_pat.finditer(b.stext)]
    if not sp or not ti:
        return []
    # 多数派 = 数量多者;平局视「加空格」为主流
    minority = 'ti' if len(sp) >= len(ti) else 'sp'
    majority_cnt = len(sp) if minority == 'ti' else len(ti)
    if majority_cnt < majority_min:
        return []
    out = []
    for b, m in (ti if minority == 'ti' else sp):
        t = b.stext
        snippet = t[max(0, m.start() - 5): m.end() + 5].strip()
        out.append(ctx.issue(
            b, rule, '「…%s…」中英文之间%s,与全文主流(%s)不一致' %
            (snippet, '未加空格' if minority == 'ti' else '加了空格',
             '加空格' if minority == 'ti' else '不加空格')))
    return out


def det_dash_density(ctx: Ctx, rule: Rule):
    """A1:破折号(——)密度超过每千字阈值,每个文件报一条。"""
    thr = float(rule.params.get('per_1000', 2.0))
    occ = [(b, b.stext.count('——')) for b in ctx.style_blocks]
    n = sum(c for _, c in occ)
    chars = ctx.char_count()
    if n <= 0 or chars <= 0:
        return []
    dens = n * 1000.0 / chars
    if dens <= thr:
        return []
    first = next(b for b, c in occ if c)
    return [ctx.issue(first, rule, '破折号(——)共 %d 处,密度 %.1f 处/千字,超过阈值 %.1f 处/千字' %
                      (n, dens, thr))]


def det_regex_hits(ctx: Ctx, rule: Rule):
    """通用词条/正则命中(A2 套话、S1 密钥、S2 个人信息等)。"""
    p = rule.params
    include_code = bool(p.get('include_code', False))
    raw = bool(p.get('raw_text', False))
    blocks = ctx.all_blocks if include_code else ctx.style_blocks
    out = []
    for pat in p.get('patterns', []):
        try:
            cre = re.compile(pat['re'])
        except re.error as e:
            print('[警告] 规则 %s 的正则非法:%s(%s)' % (rule.rid, pat['re'], e), file=sys.stderr)
            continue
        for b in blocks:
            t = b.text if raw else b.stext
            for m in cre.finditer(t):
                label = pat.get('label', '命中')
                if p.get('no_echo'):
                    msg = '%s(内容不回显)' % label
                elif p.get('mask'):
                    msg = '%s「%s」' % (label, _mask(m.group()))
                elif m.group() in label:
                    msg = label
                else:
                    msg = '%s:「%s」' % (label, m.group())
                out.append(ctx.issue(b, rule, msg, pat.get('advice', rule.advice)))
    return out


def det_tri_template(ctx: Ctx, rule: Rule):
    """A3:首先/其次/最后三段式(段首标记 ≥min_markers 种,或单段内句界分隔的同组标记)。"""
    markers = rule.params.get('markers', ['首先', '其次', '最后'])
    min_markers = int(rule.params.get('min_markers', 2))
    hits = []
    for b in ctx.style_blocks:
        t = re.sub(r'^\s*(?:[-*+>|]\s*)+', '', b.text.strip())
        for mk in markers:
            if t.startswith(mk):
                hits.append((b, mk))
                break
    out = []
    if len({mk for _, mk in hits}) >= min_markers:
        out.extend(ctx.issue(b, rule, '段落以「%s」开头,与同文档其他段首标记构成三段式模板' % mk)
                   for b, mk in hits)
    for b in ctx.style_blocks:
        t = re.sub(r'^\s*(?:[-*+>|]\s*)+', '', b.text.strip())
        found = [mk for mk in markers
                 if re.search(r'(?:^|[。;!?;,，;])\s*' + re.escape(mk), t)]
        if len(found) >= min_markers:
            out.append(ctx.issue(b, rule, '单段内「%s」构成句内三段式模板' % '/'.join(found)))
    return out


def _is_bold_label_block(ctx: Ctx, b: Block) -> bool:
    """是否为「加粗标签+冒号」开头的段落(markdown 或 docx 加粗 run)。"""
    if ctx.kind == 'md':
        return bool(re.match(r'^\*\*[^*\n]{1,20}\*\*[:\uFF1A]', b.text))
    if ctx.kind == 'docx':
        runs = [r for r in b.runs if r.text.strip()]
        if not runs:
            return False
        r0 = runs[0]
        return bool(r0.bold and re.match(r'^[^:\uFF1A\n]{1,20}[:\uFF1A]', b.text) and b.text.startswith(r0.text))
    return False


def det_bold_labels(ctx: Ctx, rule: Rule):
    """A4:连续 ≥N 段以加粗标签+冒号开头,按组报告。"""
    min_con = int(rule.params.get('min_consecutive', 3))
    flags = [(_is_bold_label_block(ctx, b), b) for b in ctx.blocks if b.text.strip()]
    out = []
    unit = '段' if ctx.kind in ('docx', 'pptx') else '行'
    start = None
    length = 0

    def flush(_end_desc=None):
        if length >= min_con and start is not None:
            end_b = flags[start + length - 1][1]
            out.append(ctx.issue(flags[start][1], rule,
                                 '%s起连续 %d %s以「加粗标签+冒号」开头(至%s),列表化写法' %
                                 (ctx.loc(flags[start][1]), length, unit, ctx.loc(end_b))))

    for flag, b in flags:
        if flag:
            if length == 0:
                start = flags.index((flag, b))
            length += 1
        else:
            flush()
            length = 0
            start = None
    flush()
    return out


def det_parallel_clauses(ctx: Ctx, rule: Rule):
    """A5:单句内 ≥N 个超短逗号分句(排比腔)。"""
    min_clauses = int(rule.params.get('min_clauses', 3))
    max_len = int(rule.params.get('max_clause_len', 8))
    out = []
    for b in ctx.style_blocks:
        if b.heading:
            continue
        for sent in re.split(r'[。;;!?!?…]+', b.stext):
            segs = [s.strip() for s in re.split(r'[,\uFF0C]', sent)]
            if len(segs) < min_clauses:
                continue
            i = 0
            while i < len(segs):
                if segs[i] and len(segs[i]) <= max_len:
                    j = i
                    while j < len(segs) and segs[j] and len(segs[j]) <= max_len:
                        j += 1
                    if j - i >= min_clauses:
                        out.append(ctx.issue(b, rule, '单句内 %d 个超短并列分句(排比):「%s」' %
                                             (j - i, '、'.join(segs[i:j]))))
                    i = j
                else:
                    i += 1
    return out


def det_hollow_closing(ctx: Ctx, rule: Rule):
    """A6:结尾段的空洞展望/总结。"""
    tail_n = int(rule.params.get('scan_tail', 1))
    tail = ctx.style_blocks[-tail_n:] if tail_n > 0 else []
    out = []
    for pat in rule.params.get('patterns', []):
        try:
            cre = re.compile(pat['re'])
        except re.error:
            continue
        for b in tail:
            for _m in cre.finditer(b.stext):
                out.append(ctx.issue(b, rule, '%s(结尾段无信息量表达)' % pat.get('label', '空洞收尾')))
    return out


DETECTORS = {
    'halfwidth_punct': det_halfwidth_punct,
    'font_consistency': det_font_consistency,
    'heading_jump': det_heading_jump,
    'date_mix': det_date_mix,
    'number_style': det_number_style,
    'cjk_latin_space': det_cjk_latin_space,
    'dash_density': det_dash_density,
    'regex_hits': det_regex_hits,
    'tri_template': det_tri_template,
    'bold_labels': det_bold_labels,
    'parallel_clauses': det_parallel_clauses,
    'hollow_closing': det_hollow_closing,
}


# ---------- 规则装载与配置 ----------

def _apply_override(rule: Rule, ov: dict) -> None:
    if 'enabled' in ov:
        rule.enabled = bool(ov['enabled'])
    if ov.get('severity') in SEVERITIES:
        rule.severity = ov['severity']
    if isinstance(ov.get('params'), dict):
        rule.params = {**rule.params, **ov['params']}


def _extra_sensitive_rules(se: dict):
    """把 --config 的 sensitive_extra 变成 S3-* 规则(用户私有,仓库只带合成示例)。"""
    rules = []
    if not se:
        return rules
    sev = se.get('severity', {}) or {}
    name_map = se.get('name_map', {}) or {}
    if name_map:
        pats = [{'re': re.escape(k), 'label': '疑似真实人名(命中本地映射)',
                 'advice': '替换为「%s」或删除' % v} for k, v in name_map.items()]
        rules.append(Rule('S3-names', 'sensitive', '敏感', '自定义人名映射', 'regex_hits',
                          sev.get('names', '中'), True,
                          {'patterns': pats, 'include_code': True, 'raw_text': True},
                          '命中本地人名映射(用户私有配置)', '按映射替换'))
    if se.get('path_prefixes'):
        pats = [{'re': re.escape(k), 'label': '内部路径(命中本地路径前缀)'}
                for k in se['path_prefixes']]
        rules.append(Rule('S3-paths', 'sensitive', '敏感', '自定义内部路径前缀', 'regex_hits',
                          sev.get('paths', '中'), True,
                          {'patterns': pats, 'include_code': True, 'raw_text': True},
                          '命中本地内部路径前缀(用户私有配置)', '替换为对外可读的合成路径或删除'))
    if se.get('domains'):
        pats = [{'re': re.escape(k), 'label': '内网域名(命中本地域名表)'}
                for k in se['domains']]
        rules.append(Rule('S3-domains', 'sensitive', '敏感', '自定义内网域名', 'regex_hits',
                          sev.get('domains', '中'), True,
                          {'patterns': pats, 'include_code': True, 'raw_text': True},
                          '命中本地内网域名表(用户私有配置)', '替换为对外域名或删除'))
    if se.get('words'):
        pats = [{'re': re.escape(k), 'label': '敏感词(命中本地敏感词表)'}
                for k in se['words']]
        rules.append(Rule('S3-words', 'sensitive', '敏感', '自定义敏感词表', 'regex_hits',
                          sev.get('words', '高'), True,
                          {'patterns': pats, 'include_code': True, 'raw_text': True, 'no_echo': True},
                          '命中本地敏感词表(用户私有配置;不回显)', '高危:替换为对外代号或删除;须人工确认'))
    return rules


def load_rules(rules_dir: Path, only: str = None, cfg: dict = None):
    rules = []
    for jf in sorted(rules_dir.glob('*.json')):
        data = json.loads(jf.read_text(encoding='utf-8'))
        for r in data.get('rules', []):
            rules.append(Rule(r['id'], data.get('line', ''), data.get('line_name', data.get('line', '')),
                              r.get('name', r['id']), r.get('detector', ''), r.get('severity', '中'),
                              r.get('enabled', True), dict(r.get('params', {})),
                              r.get('description', ''), r.get('advice', '')))
    overrides = (cfg or {}).get('rules_override', {}) or {}
    for r in rules:
        if r.rid in overrides:
            _apply_override(r, overrides[r.rid])
    for r in _extra_sensitive_rules((cfg or {}).get('sensitive_extra')):
        if r.rid in overrides:
            _apply_override(r, overrides[r.rid])
        rules.append(r)
    if only:
        rules = [r for r in rules if r.line == only]
    return [r for r in rules if r.enabled]


# ---------- 输出渲染 ----------

def _esc(s) -> str:
    return str(s).replace('|', '\\|').replace('\n', ' ')


def render_markdown(issues: list, files_meta: list, only: str, rules: list) -> str:
    out = ['# 交付物质检报告', '']
    out.append('- 检查时间:%s' % datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
    out.append('- 检查文件:%d 个(%s)' % (len(files_meta),
                                       '、'.join('%s [%s]' % (f, k) for f, k in files_meta)))
    out.append('- 检查线:%s' % (LINE_NAMES.get(only, '全部(排版 / AI腔 / 敏感)')))
    out.append('- 启用规则:%d 条;高危敏感问题不回显原文' % len(rules))
    out.append('')
    out.append('## 问题清单')
    out.append('')
    if not issues:
        out.append('未发现问题。')
    else:
        out.append('| 文件 | 位置 | 检查线 | 严重度 | 问题 | 修改建议 |')
        out.append('| --- | --- | --- | --- | --- | --- |')
        for i in issues:
            out.append('| %s | %s | %s | %s | [%s] %s | %s |' %
                       (_esc(i.file), i.loc, i.line_name, i.severity, i.rid,
                        _esc(i.message), _esc(i.advice)))
    out.append('')

    def _table(key_fn, title):
        rows = {}
        for i in issues:
            k = key_fn(i)
            d = rows.setdefault(k, {'高': 0, '中': 0, '低': 0})
            d[i.severity] += 1
        lines = ['## 汇总:%s' % title, '',
                 '| %s | 高 | 中 | 低 | 合计 |' % title,
                 '| --- | --- | --- | --- | --- |']
        for k in sorted(rows):
            d = rows[k]
            lines.append('| %s | %d | %d | %d | %d |' %
                         (_esc(k), d['高'], d['中'], d['低'], sum(d.values())))
        tot = {'高': 0, '中': 0, '低': 0}
        for d in rows.values():
            for s in tot:
                tot[s] += d[s]
        lines.append('| 合计 | %d | %d | %d | %d |' %
                     (tot['高'], tot['中'], tot['低'], sum(tot.values())))
        lines.append('')
        return lines, tot

    t1, tot = _table(lambda i: i.line_name, '按检查线')
    out += t1
    t2, _ = _table(lambda i: i.file, '按文件')
    out += t2
    out.append('结论:共 %d 个问题(高 %d / 中 %d / 低 %d)。%s' %
               (len(issues), tot['高'], tot['中'], tot['低'],
                '存在高危问题,退出码 1。' if tot['高'] else '无高危问题,退出码 0。'))
    return '\n'.join(out)


def render_json(issues: list, files_meta: list, only: str, rules: list) -> str:
    summary = {}
    for i in issues:
        d = summary.setdefault(i.line_name, {'高': 0, '中': 0, '低': 0, '合计': 0})
        d[i.severity] += 1
        d['合计'] += 1
    high = sum(1 for i in issues if i.severity == '高')
    return json.dumps({
        'check_time': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'tool': 'qa_check', 'version': __version__,
        'files': [{'path': f, 'kind': k} for f, k in files_meta],
        'line': only or 'all',
        'rules_enabled': len(rules),
        'issues': [{'file': i.file, 'loc': i.loc, 'line': i.line_name, 'rule': i.rid,
                    'severity': i.severity, 'message': i.message, 'advice': i.advice}
                   for i in issues],
        'summary': summary,
        'total': len(issues), 'high': high, 'exit_code': 1 if high else 0,
    }, ensure_ascii=False, indent=2)


# ---------- 主流程 ----------

def main(argv=None) -> int:
    if hasattr(sys.stdout, 'reconfigure'):
        try:
            sys.stdout.reconfigure(encoding='utf-8')
        except Exception:
            pass
    ap = argparse.ArgumentParser(
        prog='qa_check',
        description='交付物质检:排版规范 / AI 腔 / 敏感内容三条线(仅 Python 标准库,规则见 rules/*.json)',
        epilog='示例:python qa_check.py report.docx --config C:/qa_local.json --only sensitive')
    ap.add_argument('files', nargs='+', help='待检文件路径(md/txt/docx/pptx),可多个')
    ap.add_argument('--config', help='本地规则覆盖与私有敏感词配置(JSON),建议放仓库外')
    ap.add_argument('--only', choices=['typography', 'aitone', 'sensitive'], help='只跑指定检查线')
    ap.add_argument('--json', action='store_true', help='以 JSON 输出(供脚本/自测使用)')
    args = ap.parse_args(argv)

    cfg = {}
    if args.config:
        cpath = Path(args.config)
        if not cpath.is_file():
            print('错误:配置文件不存在:%s' % args.config, file=sys.stderr)
            return 2
        try:
            cfg = json.loads(cpath.read_text(encoding='utf-8-sig'))
        except (json.JSONDecodeError, OSError) as e:
            print('错误:配置文件解析失败:%s' % e, file=sys.stderr)
            return 2
        if not isinstance(cfg, dict):
            print('错误:配置文件顶层必须是 JSON 对象', file=sys.stderr)
            return 2

    rules_dir = Path(__file__).resolve().parent / 'rules'
    if not rules_dir.is_dir():
        print('错误:找不到规则目录:%s' % rules_dir, file=sys.stderr)
        return 2
    rules = load_rules(rules_dir, args.only, cfg)
    if not rules:
        print('错误:没有启用的规则(检查 rules/*.json 与 --only/--config)', file=sys.stderr)
        return 2

    issues = []
    files_meta = []
    for fi, fpath in enumerate(args.files):
        p = Path(fpath)
        if not p.is_file():
            print('[跳过] 文件不存在:%s' % fpath, file=sys.stderr)
            continue
        try:
            blocks, kind = parse_file(p)
        except (OSError, ValueError, zipfile.BadZipFile, ET.ParseError) as e:
            print('[跳过] 无法解析 %s:%s' % (fpath, e), file=sys.stderr)
            continue
        ctx = Ctx(p, str(p), kind, blocks, file_index=fi)
        files_meta.append((str(p), kind))
        for rule in rules:
            det = DETECTORS.get(rule.detector)
            if det is None:
                print('[警告] 未知检测器:%s(规则 %s)' % (rule.detector, rule.rid), file=sys.stderr)
                continue
            issues.extend(det(ctx, rule))

    if not files_meta:
        print('错误:没有可检查的文件', file=sys.stderr)
        return 2

    issues.sort(key=lambda i: i.sort_key)
    if args.json:
        print(render_json(issues, files_meta, args.only, rules))
    else:
        print(render_markdown(issues, files_meta, args.only, rules))
    return 1 if any(i.severity == '高' for i in issues) else 0


if __name__ == '__main__':
    sys.exit(main())
