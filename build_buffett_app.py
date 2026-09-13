#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
巴菲特致股东信知识库 → 单文件阅读研究应用 构建器
==================================================

功能：
  1. 扫描「巴菲特致股东信知识库/」七个固定分类目录第一层的 .md 文章
  2. 提取元数据：分类、年份、标题、关联主题标签（解析文中交叉链接）
  3. 把原站失效的相对链接（../concepts/xx.html 等）重写为应用内锚点 (#a/<article-id>)
  4. 生成完全自包含的单文件应用「巴菲特投资智慧.html」（无 CDN、无外部依赖，双击即用）
  5. 可选：从项目根 .env 生成无密钥 llm-config.js（仅默认 base/model）

用法：
  python3 build_buffett_app.py            # 构建 html（默认同时生成 llm-config.js）
  python3 build_buffett_app.py --no-llm-config   # 跳过 llm-config.js
  python3 build_buffett_app.py --debug    # 打印数据统计

仅依赖 Python 标准库。
"""

import base64
import json
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

HERE = os.path.dirname(os.path.abspath(__file__))
KB_ROOT = os.path.join(HERE, "巴菲特致股东信知识库")


def load_icon_b64(name="buffett-128.png"):
    """assets/<name> → base64（单文件内嵌，离线可用）。缺失时返回空串。"""
    p = os.path.join(HERE, "assets", name)
    try:
        with open(p, "rb") as f:
            return base64.b64encode(f.read()).decode("ascii")
    except OSError:
        return ""
OUT_HTML = os.path.join(HERE, "巴菲特投资智慧.html")
OUT_LLM = os.path.join(HERE, "llm-config.js")
XLSX_PATH = os.path.join(HERE, "巴菲特致股东信分类索引(1956-2025) .xlsx")

_XLSX_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_XLSX_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def read_xlsx_grids(path):
    """纯标准库读取 xlsx → {sheet名: [[单元格值, ...], ...]}（空单元格为 ''）。"""
    def col_to_idx(ref):
        m = re.match(r"([A-Z]+)", ref or "")
        if not m:
            return 0
        n = 0
        for ch in m.group(1):
            n = n * 26 + (ord(ch) - ord("A") + 1)
        return n - 1

    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        shared = []
        if "xl/sharedStrings.xml" in names:
            root = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in root.findall(_XLSX_NS + "si"):
                shared.append("".join(t.text or "" for t in si.iter(_XLSX_NS + "t")))
        rels = {}
        if "xl/_rels/workbook.xml.rels" in names:
            for rel in ET.fromstring(z.read("xl/_rels/workbook.xml.rels")):
                rels[rel.get("Id")] = rel.get("Target")
        sheets = [(sh.get("name"), sh.get(_XLSX_REL + "id"))
                  for sh in ET.fromstring(z.read("xl/workbook.xml")).iter(_XLSX_NS + "sheet")]
        out = {}
        for idx, (name, rid) in enumerate(sheets, 1):
            target = rels.get(rid, "worksheets/sheet%d.xml" % idx)
            if not target.startswith("xl/"):
                target = "xl/" + target
            target = target.replace("../", "")
            if target not in names:
                continue
            grid = []
            for row in ET.fromstring(z.read(target)).iter(_XLSX_NS + "row"):
                cells, maxcol = {}, -1
                for c in row.findall(_XLSX_NS + "c"):
                    col = col_to_idx(c.get("r"))
                    t = c.get("t")
                    v = c.find(_XLSX_NS + "v")
                    if t == "s":
                        val = shared[int(v.text)] if v is not None and v.text else ""
                    elif t == "inlineStr":
                        val = "".join(x.text or "" for x in c.iter(_XLSX_NS + "t"))
                    else:
                        val = v.text if v is not None and v.text else ""
                        if re.fullmatch(r"-?\d+(\.\d+)?", str(val)):
                            val = int(float(val))
                    cells[col] = val
                    maxcol = max(maxcol, col)
                if cells:
                    grid.append([cells.get(i, "") for i in range(maxcol + 1)])
            out[name] = grid
    return out


def parse_years(text):
    """解析 “1977, 1986, 1995-1998” → [1977, 1986, 1995..1998]（含区间展开）。"""
    years = set()
    for m in re.finditer(r"(\d{4})\s*(?:-\s*(\d{4}))?", text or ""):
        a, b = int(m.group(1)), int(m.group(2)) if m.group(2) else int(m.group(1))
        if b >= a and b - a <= 70:
            years.update(range(a, b + 1))
    return sorted(years)


def split_terms(text):
    return [t.strip() for t in re.split(r"[;；,，]+", text or "") if t.strip()]


def md_to_plain(md):
    """Markdown → 纯文本（供年度摘要合成等使用）。"""
    s = re.sub(r"!\[([^\]]*)\]\([^)]*\)", r"\1", md)
    s = re.sub(r"\[([^\]]*)\]\([^)]*\)", r"\1", s)
    s = re.sub(r"\[\^[^\]]*\]\([^)]*\)", "", s)
    s = re.sub(r"\[\^[^\]]*\]", "", s)
    s = re.sub(r"^#{1,6}\s+", "", s, flags=re.M)
    s = re.sub(r"^\s*>\s?", "", s, flags=re.M)
    s = re.sub(r"^\s*([-*+]|\d+\.)\s+", "", s, flags=re.M)
    s = re.sub(r"^---+$", "", s, flags=re.M)
    s = re.sub(r"[|*_`~]", " ", s)
    s = re.sub(r"\s+", " ", s)
    return s.strip()


def parse_classification_index(letter_years=None, by_id=None):
    """解析「巴菲特致股东信分类索引(1956-2025) .xlsx」→ (五维索引, 年度原始行)。

    letter_years: {年份: [文章id, ...]} 信件类文章按年份分组；用于给 Excel
    未覆盖的年份自动合成年度索引条目。
    返回 (idx, year_rows)；year_rows 为按 (年份, 信件系列) 的原始行，
    供 inject_summary_and_links() 做文章级摘要/原文链接注入。
    """
    letter_years = letter_years or {}
    by_id = by_id or {}
    if not os.path.isfile(XLSX_PATH):
        print("[skip] 未找到分类索引 xlsx，跳过", file=sys.stderr)
        return None
    grids = read_xlsx_grids(XLSX_PATH)

    def cell(row, i):
        return row[i] if i < len(row) else ""

    # ---- 主题分类索引（坎宁安主题）----
    topics, tmap = [], {}
    for i, row in enumerate(grids.get("主题分类索引", [])[1:], 1):
        name = str(cell(row, 0)).strip()
        canon = re.sub(r"^[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]+\s*", "", name)
        if not name:
            continue
        k = "T%d" % i
        topics.append({"k": k, "n": name, "c": canon,
                       "d": str(cell(row, 1)).strip(),
                       "y": parse_years(cell(row, 2)),
                       "rep": str(cell(row, 3)).strip(),
                       "con": split_terms(cell(row, 4))})
        tmap[canon] = k

    def canon_to_k(tok):
        tok = tok.strip()
        if tok in tmap:
            return tmap[tok]
        for canon, k in tmap.items():      # 前缀匹配（如 “美国经验” → “美国经验/美国顺风”）
            if canon.startswith(tok):
                return k
        return None

    # ---- 行业分类索引 ----
    industries = []
    for i, row in enumerate(grids.get("行业分类索引", [])[1:], 1):
        name = str(cell(row, 0)).strip()
        if not name:
            continue
        industries.append({"k": "I%d" % i, "n": name,
                           "co": str(cell(row, 1)).strip(),
                           "y": parse_years(cell(row, 2)),
                           "d": str(cell(row, 3)).strip()})

    # ---- 事件时期索引 ----
    events = []
    for i, row in enumerate(grids.get("事件时期索引", [])[1:], 1):
        name = str(cell(row, 0)).strip()
        if not name:
            continue
        events.append({"k": "E%d" % i, "n": name,
                       "rng": str(cell(row, 1)).strip(),
                       "y": parse_years(cell(row, 2)),
                       "bg": str(cell(row, 3)).strip(),
                       "act": str(cell(row, 4)).strip(),
                       "les": str(cell(row, 5)).strip()})

    # ---- 选股方法演进 ----
    methods = []
    for i, row in enumerate(grids.get("选股方法演进", [])[1:], 1):
        name = str(cell(row, 0)).strip()
        if not name:
            continue
        methods.append({"k": "M%d" % i, "n": name,
                        "m": str(cell(row, 1)).strip(),
                        "y": parse_years(cell(row, 2)),
                        "view": str(cell(row, 3)).strip(),
                        "cases": str(cell(row, 4)).strip(),
                        "shift": str(cell(row, 5)).strip()})

    # ---- 年度总索引（新表：A年份 B信件系列 C撰写人 D核心主题摘要 E坎宁安主题 F行业 G背景 H事件 I原文链接）----
    year_rows = []
    for i, row in enumerate(grids.get("年度总索引", [])[1:], 1):
        yv = cell(row, 0)
        if not str(yv).strip():
            continue
        try:
            y = int(yv)
        except (TypeError, ValueError):
            continue
        year_rows.append({
            "y": y,
            "series": str(cell(row, 1)).strip(),
            "a": str(cell(row, 2)).strip(),
            "s": str(cell(row, 3)).strip(),
            "t": [tok for tok in split_terms(cell(row, 4)) if canon_to_k(tok)],
            "i": split_terms(cell(row, 5)),
            "bg": str(cell(row, 6)).strip(),
            "e": str(cell(row, 7)).strip(),
            "link": str(cell(row, 8)).strip(),
        })
    # 同年多系列（1965-1969 合伙基金信+伯克希尔信并存）合并为一个年度条目
    years, seen_y = [], set()
    for y in sorted({r["y"] for r in year_rows}):
        rows = [r for r in year_rows if r["y"] == y]
        multi = len(rows) > 1

        def _uniq(xs):
            out = []
            for x in xs:
                if x and x not in out:
                    out.append(x)
            return out

        years.append({
            "k": "Y%d" % y,
            "y": y,
            "a": "、".join(_uniq([r["a"] for r in rows])),
            "s": "\n".join((("【%s】" % r["series"]) if multi else "") + r["s"] for r in rows),
            "t": list(dict.fromkeys(t for r in rows for t in r["t"])),
            "i": list(dict.fromkeys(x for r in rows for x in r["i"])),
            "bg": "；".join(_uniq([r["bg"] for r in rows])),
            "e": "、".join(_uniq([r["e"] for r in rows])),
            "link": "；".join(_uniq([r["link"] for r in rows if r["link"]])),
            "series": _uniq([r["series"] for r in rows]),
        })
        seen_y.add(y)
    # 兜底：Excel 未覆盖的年份（如未来新增信件），从知识库正文合成条目
    for y in sorted(k for k in letter_years if k not in seen_y):
        ids = letter_years[y]
        main = sorted(ids, key=lambda aid: (1 if re.search(r"年\d+月|年中", aid) else 0, aid))[0]
        art = by_id.get(main)
        summary = ""
        if art:
            plain = md_to_plain(art["md"])
            plain = re.sub(r"^\d{4}年\d{1,2}月\d{1,2}日\s*", "", plain)
            summary = plain[:200]
        years.append({"k": "Y%d" % y, "y": y, "a": "巴菲特",
                      "s": summary, "t": [], "i": [], "bg": "", "e": "", "link": "", "series": []})
    years.sort(key=lambda x: x["y"])

    print("[ok] 已解析分类索引：主题%d / 行业%d / 事件%d / 方法%d / 年度%d（%d-%d）"
          % (len(topics), len(industries), len(events), len(methods), len(years),
             years[0]["y"] if years else 0, years[-1]["y"] if years else 0))
    return ({"topic": topics, "industry": industries,
             "event": events, "method": methods, "year": years}, year_rows)


# 文章分类 → 信件系列（用于匹配年度总索引行）。
# 特别信件不是当年股东信，不按年份沿用伯克希尔年报摘要和链接。
_SERIES_OF_CAT = {"partnership": "合伙基金信", "berkshire": "伯克希尔信"}


def inject_summary_and_links(by_id, year_rows):
    """按 (年份, 信件系列) 匹配年度总索引行：
    核心主题摘要 → 文章开头（引用块）；原文链接 → 文章末尾（🔗 段落）。"""
    if not year_rows:
        return
    by_key = {(r["y"], r["series"]): r for r in year_rows}
    n_sum = n_link = 0
    for art in by_id.values():
        if art["catKey"] not in _SERIES_OF_CAT or not art["year"]:
            continue
        row = by_key.get((art["year"], _SERIES_OF_CAT[art["catKey"]]))
        if not row:
            cands = [r for r in year_rows if r["y"] == art["year"]]
            row = cands[0] if cands else None
        if not row:
            continue
        if row["s"]:
            art["md"] = "> 📌 **核心主题摘要**：%s\n\n%s" % (row["s"], art["md"])
            art["len"] += len(row["s"]) + 24
            n_sum += 1
        if row["link"]:
            art["md"] = "%s\n\n## 🔗 原文链接\n\n[%s](%s)\n" % (art["md"], row["link"], row["link"])
            art["len"] += len(row["link"]) * 2 + 28
            n_link += 1
    print("[ok] 文章摘要注入 %d 篇 / 原文链接 %d 篇" % (n_sum, n_link))

# 目录 → (catKey, 显示名, 排序权重)
CATS = [
    ("01-索引", "index", "索引", 0),
    ("02-合伙人信", "partnership", "合伙人信", 1),
    ("03-伯克希尔股东信", "berkshire", "伯克希尔股东信", 2),
    ("04-特别信件", "special", "特别信件", 3),
    ("05-概念", "concept", "概念", 4),
    ("06-公司", "company", "公司", 5),
    ("07-人物", "person", "人物", 6),
]
CAT_KEY2NAME = {k: name for _, k, name, _ in CATS}

# 原站链接前缀 → 本地分类（用于链接重写时的目标定位）
PREFIX2CAT = {
    "../index-pages/": "index",
    "../partnership/": "partnership",
    "../berkshire/": "berkshire",
    "../special/": "special",
    "../concepts/": "concept",
    "../companies/": "company",
    "../people/": "person",
}

YEAR_RE = re.compile(r"(19|20)\d{2}")
LINK_RE = re.compile(r"\[([^\]]*)\]\(([^)]*)\)")
FN_INLINE_RE = re.compile(r"\[\^([^\]]+)\]\(([^)]*)\)")   # 行内脚注 [^1](定义)
FN_REF_RE = re.compile(r"^\[\^([^\]]+)\]:\s*(.+)$")        # 文末引用式脚注 [^1]: 定义


# ---------------------------------------------------------------- 扫描与元数据

def scan_articles():
    """第一遍：收集所有 md 文件 → 建立 stem→id 映射 + 原始内容。"""
    stem2id, raw = {}, []
    for folder, cat_key, name, order in CATS:
        d = os.path.join(KB_ROOT, folder)
        if not os.path.isdir(d):
            print(f"[warn] 缺少目录: {d}", file=sys.stderr)
            continue
        for fn in sorted(os.listdir(d)):
            if not fn.endswith(".md"):
                continue
            stem = fn[:-3]
            aid = f"{cat_key}-{stem}"
            stem2id[stem] = aid
            with open(os.path.join(d, fn), encoding="utf-8") as f:
                raw.append({"id": aid, "catKey": cat_key, "stem": stem, "md": f.read()})
    return stem2id, raw


def extract_meta(article, stem2id):
    """第二遍：元数据 + 链接重写。"""
    md = article["md"]

    # 标题：第一个 # 一级标题
    m = re.search(r"^# (.+)$", md, re.M)
    title = m.group(1).strip() if m else article["stem"]
    # 丢弃首个标题之前的分类标签行（如 “Economic Moat 概念” / “合伙人信”）
    if m:
        md = md[m.end():].lstrip("\n")

    # 年份：文件名中第一个四位年份
    ym = YEAR_RE.search(article["stem"])
    year = int(ym.group(0)) if ym else None

    # ---- 脚注抽取（必须在链接重写之前，否则行内脚注 [^X](定义) 会被误判为链接）----
    fns, seen = [], set()
    def _reg_fn(key, text):
        t = text.strip()
        if t and (key, t) not in seen:
            seen.add((key, t))
            fns.append([key, t])
    md = FN_INLINE_RE.sub(lambda m: (_reg_fn(m.group(1), m.group(2)), "[%s]" % m.group(1))[1], md)
    md = re.sub(r"^\[\^([^\]]+)\]:\s*.+$", lambda m: (_reg_fn(m.group(1), m.group(0).split(":", 1)[1]), "")[1], md, flags=re.M)

    # 链接重写（md 与脚注定义共用）：收集内部链接
    links, tags = [], set()
    def _rewrite_links(text):
        def _rep(mm):
            label, href = mm.group(1), mm.group(2).strip()
            if mm.group(0).startswith("!"):
                return mm.group(0)  # 图片，保持原样
            if href.startswith("#"):
                return mm.group(0)
            path = href.split("#")[0]
            if path.startswith(("http://", "https://", "mailto:")):
                return mm.group(0)
            stem = os.path.splitext(os.path.basename(path))[0]
            target = stem2id.get(stem)
            if target:
                links.append([target, label])
                tcat = target.split("-", 1)[0]
                if tcat in ("concept", "company", "person"):
                    tags.add(target)
                return f"[{label}](#/a/{target})"
            return f"[{label}]"  # 无法解析的目标：降级为纯文本
        return LINK_RE.sub(_rep, text)

    md = _rewrite_links(md)
    fns = [[k, _rewrite_links(t)] for k, t in fns]

    # 标签去重、去自身
    tags.discard(article["id"])
    return {
        "id": article["id"],
        "catKey": article["catKey"],
        "year": year,
        "title": title,
        "md": md,
        "len": len(md),
        "links": links,
        "tags": sorted(tags),
        "fns": fns,
    }


def build():
    stem2id, raw = scan_articles()
    by_id = {}
    for art in raw:
        meta = extract_meta(art, stem2id)
        by_id[meta["id"]] = meta

    # 标签用目标文章标题（干净名称），并统计标签频次
    tag_count = {}
    for art in by_id.values():
        named = []
        for tid in art["tags"]:
            t = by_id.get(tid)
            name = t["title"] if t else tid
            # 去掉标题中的英文括号后缀，如 “可口可乐（Coca-Cola）” → “可口可乐”
            m2 = re.match(r"^(.*?)（[^（）]*）$", name)
            if m2:
                name = m2.group(1).strip()
            named.append(name)
            tag_count[name] = tag_count.get(name, 0) + 1
        # 去重保序
        art["tags"] = list(dict.fromkeys(named))

    articles = [by_id[k] for k in sorted(by_id)]
    cats = []
    for folder, cat_key, name, order in CATS:
        n = sum(1 for a in articles if a["catKey"] == cat_key)
        if n:
            cats.append({"key": cat_key, "name": name, "count": n, "order": order})

    years = sorted({a["year"] for a in articles if a["year"]})
    # 年度信件按年份分组（年度索引补全早年用）
    letter_years = {}
    for a in articles:
        if a["year"] and a["catKey"] in ("berkshire", "partnership"):
            letter_years.setdefault(a["year"], []).append(a["id"])
    idx, year_rows = parse_classification_index(letter_years, by_id)
    if idx:
        inject_summary_and_links(by_id, year_rows)
    data = {
        "title": "巴菲特投资智慧",
        "subtitle": "巴菲特致股东信知识库 · 阅读研究",
        "built": __import__("datetime").date.today().isoformat(),
        "cats": cats,
        "yearRange": [years[0], years[-1]] if years else [],
        "articles": articles,
        "idx": idx,
        "buffettPersona": load_buffett_persona(),
        "mungerPersona": load_munger_persona(),
    }
    return data, tag_count


def load_buffett_persona():
    """加载 distilly 生成的 celebrity-buffett 人格内容（构建时嵌入，供「与巴菲特对话」专栏使用）。

    候选路径：
      1. 用户级 DSH 技能源：~/.agents/skills/celebrity/buffett/{persona,work}.md
      2. DSH 安装副本：~/.dsh/skills/celebrity-buffett/SKILL.md（剥离 frontmatter）
    全部缺失时返回空串（应用内会显示提示）。
    """
    candidates = [
        # 项目级副本优先（随项目打包移动，日后再构建也用它）
        (os.path.join(HERE, "skills", "celebrity-buffett"), ["persona.md", "work.md"]),
        (os.path.expanduser("~/.agents/skills/celebrity/buffett"), ["persona.md", "work.md"]),
        (os.path.expanduser("~/.dsh/skills/celebrity-buffett"), ["SKILL.md"]),
    ]
    for base, files in candidates:
        parts, ok = [], True
        for fn in files:
            p = os.path.join(base, fn)
            if not os.path.isfile(p):
                ok = False
                break
            text = open(p, encoding="utf-8").read()
            if fn == "SKILL.md":
                m = re.match(r"^---\n.*?\n---\n", text, re.S)  # 剥离 YAML frontmatter
                if m:
                    text = text[m.end():]
            parts.append(text)
        if ok:
            return "\n\n".join(parts)
    print("[warn] 未找到 celebrity-buffett 人格文件，跳过「与巴菲特对话」内容嵌入", file=sys.stderr)
    return ""


def load_munger_persona():
    """加载 distilly 生成的 celebrity-charlie-munger 人格内容（构建时嵌入，供「与芒格对话」专栏使用）。

    候选路径：
      1. 项目级副本（随项目打包移动）：skills/celebrity-charlie-munger/{persona,work}.md
      2. 用户级 DSH 技能源：~/.agents/skills/celebrity/charlie-munger/{persona,work}.md
    全部缺失时返回空串（应用内会显示提示）。
    """
    candidates = [
        (os.path.join(HERE, "skills", "celebrity-charlie-munger"), ["persona.md", "work.md"]),
        (os.path.expanduser("~/.agents/skills/celebrity/charlie-munger"), ["persona.md", "work.md"]),
    ]
    for base, files in candidates:
        parts, ok = [], True
        for fn in files:
            p = os.path.join(base, fn)
            if not os.path.isfile(p):
                ok = False
                break
            parts.append(open(p, encoding="utf-8").read())
        if ok:
            return "\n\n".join(parts)
    print("[warn] 未找到 celebrity-charlie-munger 人格文件，跳过「与芒格对话」内容嵌入", file=sys.stderr)
    return ""


# ---------------------------------------------------------------- 应用模板

APP_CSS = r"""
:root{
  --bg:#f6f3ec; --panel:#fffdf8; --ink:#2a2620; --ink2:#6f675a; --ink3:#9a917f;
  --line:#e6dfd1; --line2:#efe9dd;
  --accent:#a16207; --accent2:#b45309; --green:#3f6212; --blue:#1e4f9e;
  --hl:#ffebb3; --hl2:#bfe3ff; --hl3:#ffe1d8;
  --serif:Georgia,"Songti SC","STSong","SimSun",serif;
  --sans:-apple-system,BlinkMacSystemFont,"PingFang SC","Hiragino Sans GB","Microsoft YaHei","Noto Sans CJK SC",sans-serif;
  --mono:ui-monospace,SFMono-Regular,Menlo,Consolas,monospace;
}
*{box-sizing:border-box;margin:0;padding:0}
[hidden]{display:none !important}
html{scroll-behavior:smooth}
body{background:var(--bg);color:var(--ink);font:15px/1.65 var(--sans);}
button{font:inherit;color:inherit;background:none;border:none;cursor:pointer}
select,input,textarea{font:inherit;color:inherit}
a{color:var(--accent2);text-decoration:none}
a:hover{text-decoration:underline}
::-webkit-scrollbar{width:10px;height:10px}
::-webkit-scrollbar-thumb{background:#d8cfbc;border-radius:6px;border:2px solid var(--bg)}
::-webkit-scrollbar-track{background:transparent}

/* ---------- 顶栏 ---------- */
#topbar{position:sticky;top:0;z-index:50;display:flex;align-items:center;gap:14px;
  padding:10px 18px;background:rgba(246,243,236,.94);backdrop-filter:blur(8px);
  border-bottom:1px solid var(--line);flex-wrap:wrap}
.brand{display:flex;align-items:baseline;gap:8px;white-space:nowrap;cursor:pointer;border-radius:8px;padding:2px 6px;margin-left:-6px}
.brand:hover{background:var(--line2)}
.brand .logo{font-size:19px;font-weight:700;letter-spacing:.5px;color:var(--accent)}
.brand .logo em{font-style:normal;color:var(--accent2)}
.brand .brand-img{width:22px;height:22px;border-radius:6px;vertical-align:-4px;margin-right:6px}
.brand .sub{font-size:12px;color:var(--ink3)}
.search-wrap{flex:1;min-width:180px;max-width:560px;position:relative;display:flex;align-items:center}
#q{width:100%;padding:8px 34px 8px 12px;border:1px solid var(--line);border-radius:9px;
  background:var(--panel);font-size:14px;outline:none;transition:border-color .15s,box-shadow .15s}
#q:focus{border-color:var(--accent2);box-shadow:0 0 0 3px rgba(180,83,9,.12)}
#qCount{position:absolute;right:10px;font-size:12px;color:var(--ink3);pointer-events:none}
.top-actions{display:flex;align-items:center;gap:8px;flex-wrap:wrap}
.tbtn{padding:7px 11px;border:1px solid var(--line);border-radius:9px;background:var(--panel);
  font-size:13px;color:var(--ink2);display:inline-flex;align-items:center;gap:5px;white-space:nowrap}
.tbtn:hover{border-color:var(--accent2);color:var(--accent2)}
.tbtn.active{background:var(--accent);border-color:var(--accent);color:#fff}
.tbtn svg{width:14px;height:14px;fill:currentColor}
select.tbtn{-webkit-appearance:none;appearance:none;padding-right:24px;
  background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='10' height='6'%3E%3Cpath d='M0 0l5 6 5-6z' fill='%236f675a'/%3E%3C/svg%3E");
  background-repeat:no-repeat;background-position:right 9px center}

/* ---------- 布局 ---------- */
#layout{display:flex;align-items:flex-start;gap:0}
#sidebar{width:236px;flex:0 0 236px;position:sticky;top:57px;height:calc(100vh - 57px);
  overflow-y:auto;padding:16px 14px 40px;border-right:1px solid var(--line);background:var(--panel)}
#sbToggle{display:none}
@media (min-width:861px){
  #sbToggle{display:inline-flex}
  #sidebar.collapsed{display:none}
}
.sb-sec{margin-bottom:18px}
.sb-title{font-size:12px;font-weight:700;color:var(--ink3);letter-spacing:1px;margin-bottom:8px;
  display:flex;justify-content:space-between;align-items:center}
.sb-title-range{font-weight:400;letter-spacing:0}
.sb-item{display:flex;align-items:center;justify-content:space-between;gap:6px;width:100%;
  padding:5px 8px;border-radius:7px;font-size:13.5px;color:var(--ink2);text-align:left}
.sb-item:hover{background:var(--line2);color:var(--ink)}
.sb-item.active{background:#f3e6c8;color:var(--accent2);font-weight:600}
.sb-item .n{font-size:11px;color:var(--ink3)}
.sb-item.active .n{color:var(--accent2)}
.tag-cloud{display:flex;flex-wrap:wrap;gap:5px}
.tag-chip{font-size:12px;padding:3px 9px;border-radius:20px;border:1px solid var(--line);
  background:var(--bg);color:var(--ink2);cursor:pointer;white-space:nowrap}
.tag-chip:hover{border-color:var(--accent2);color:var(--accent2)}
.tag-chip.active{background:var(--green);border-color:var(--green);color:#fff}
.sb-check{display:flex;align-items:center;gap:8px;padding:5px 8px;font-size:13.5px;color:var(--ink2);cursor:pointer}
.sb-check input{accent-color:var(--accent2)}
.sb-foot{margin-top:10px;padding-top:10px;border-top:1px dashed var(--line);font-size:11.5px;color:var(--ink3);line-height:1.7}

/* 分类索引（xlsx 五维） */
.idx-group{margin-bottom:6px}
.idx-gtitle{display:flex;justify-content:space-between;align-items:center;width:100%;
  padding:6px 9px;border-radius:7px;font-size:12.5px;font-weight:700;color:var(--ink2);
  background:var(--line2);cursor:pointer;text-align:left}
.idx-gtitle:hover{color:var(--accent2)}
.idx-gtitle .n{font-size:11px;font-weight:400;color:var(--ink3)}
.idx-items{display:none;padding:3px 0}
.idx-group.open .idx-items{display:block}
.idx-items .sb-item{font-size:12.5px;padding:4px 8px;border-radius:6px}
.idx-items .sb-item.active{background:#f3e6c8;color:var(--accent2);font-weight:600}
.idx-items .sb-item .n{font-size:10.5px}
.idx-banner{background:#faf6ea;border:1px solid #e3d5b8;border-left:4px solid var(--accent2);
  border-radius:10px;padding:14px 44px 14px 18px;margin-bottom:16px;position:relative}
.idx-banner h3{font:700 16px/1.5 var(--serif);margin-bottom:8px}
.idx-banner .idx-rng{font-size:12px;color:var(--ink3);font-weight:400;margin-left:6px}
.idx-banner p{font-size:13px;color:#4d4638;line-height:1.8;margin:.35em 0;text-align:justify}
.idx-banner p b{color:var(--accent2)}
.idx-banner .idx-rep{color:var(--ink2);font-size:12.5px}
.idx-banner .idx-sum{font-size:13.5px}
.idx-banner .idx-chips{display:flex;flex-wrap:wrap;align-items:center;gap:6px;margin-top:9px}
.idx-banner .idx-lbl{font-size:12px;color:var(--ink2);font-weight:600}
.idx-banner .mini-tag{cursor:pointer}
.idx-close{position:absolute;top:10px;right:12px;color:var(--ink3);font-size:15px;padding:2px 7px;border-radius:6px}
.idx-close:hover{color:#b91c1c;background:#fde8e8}
.r-idx{display:flex;align-items:center;gap:7px;flex-wrap:wrap;margin-top:7px;font-size:12px;
  color:var(--ink2);background:#faf6ea;border:1px solid var(--line2);border-radius:8px;padding:5px 10px}
.r-idx .idx-ev{color:var(--ink2)}
.r-idx b{color:var(--accent2)}
.r-idx .mini-tag{cursor:pointer}

/* ---------- 主页 ---------- */
#home{padding-top:8px}
.home-hero{text-align:center;padding:36px 20px 30px;background:linear-gradient(180deg,#faf4e4,transparent);
  border-radius:16px;border:1px solid var(--line2)}
.home-hero .hh-icon{font-size:46px;line-height:1}
.home-hero .hh-img{width:96px;height:96px;border-radius:18px;box-shadow:0 4px 16px rgba(139,111,71,.28);margin-bottom:2px}
.home-hero .hh-chart-wrap{position:relative;width:min(1280px,100%);height:310px;margin:10px auto 4px}
.home-hero .hh-chart{width:100%;height:100%}
.home-hero .chart-note{position:absolute;right:10px;top:4px;max-width:calc(100% - 20px);font-size:10px;line-height:1.35;
  text-align:right;color:var(--ink3);z-index:5;pointer-events:none}
.home-hero h1{font:700 34px/1.4 var(--serif);color:var(--accent);margin:10px 0 6px}
.home-hero .hh-sub{color:var(--ink2);font-size:14px}
.hh-stats{display:flex;justify-content:center;gap:38px;margin:24px 0 20px;flex-wrap:wrap}
.hh-stats b{display:block;font:700 25px/1.2 var(--serif);color:var(--accent2)}
.hh-stats span{font-size:12px;color:var(--ink3)}
.hh-search{display:flex;gap:8px;justify-content:center;max-width:540px;margin:0 auto}
.hh-search input{flex:1;padding:11px 15px;border:1px solid var(--line);border-radius:10px;
  background:var(--panel);font-size:14.5px;outline:none;min-width:0}
.hh-search input:focus{border-color:var(--accent2);box-shadow:0 0 0 3px rgba(180,83,9,.12)}
.hh-search button{padding:11px 24px;border-radius:10px;background:var(--accent2);color:#fff;font-weight:600;flex:0 0 auto}
.hh-search button:hover{background:#92400e}
.home-sec{margin-top:28px}
.home-sec>h2{font:700 17.5px var(--serif);margin-bottom:13px;display:flex;align-items:center;gap:9px}
.home-sec>h2::after{content:"";flex:1;height:1px;background:var(--line)}
.home-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(210px,1fr));gap:12px}
.home-tile{background:var(--panel);border:1px solid var(--line);border-radius:12px;padding:15px 17px;
  cursor:pointer;transition:transform .12s,box-shadow .12s,border-color .12s;text-align:left}
.home-tile:hover{transform:translateY(-2px);box-shadow:0 6px 18px rgba(90,70,30,.09);border-color:#d8c9a8}
.home-tile .ht-ic{font-size:24px;line-height:1}
.home-tile .ht-name{font-weight:700;margin-top:8px;font-size:15px}
.home-tile .ht-desc{font-size:12px;color:var(--ink3);margin-top:4px;line-height:1.65}
.home-list{display:flex;flex-direction:column;gap:8px}
.home-row{display:flex;align-items:baseline;gap:12px;padding:11px 15px;background:var(--panel);
  border:1px solid var(--line);border-radius:10px;cursor:pointer}
.home-row:hover{border-color:#d8c9a8;background:#fffdf8;transform:translateY(-1px)}
.home-row .hr-year{font-weight:700;color:var(--accent2);flex:0 0 74px;font-size:13.5px}
.home-row .hr-title{font-weight:600;font-size:14.5px}
.home-row .hr-why{font-size:12px;color:var(--ink3);margin-left:auto;text-align:right;max-width:48%;line-height:1.6}
.home-chips{display:flex;flex-wrap:wrap;gap:8px}
.home-chips .tag-chip{font-size:13px;padding:5px 14px}
.home-last{display:flex;align-items:center;gap:14px;background:#faf6ea;border:1px solid #e3d5b8;
  border-radius:12px;padding:15px 19px;cursor:pointer;margin-top:28px;transition:border-color .12s}
.home-last:hover{border-color:var(--accent2)}
.home-last .hl-ic{font-size:26px}
.home-last .hl-title{font-weight:700;font-size:15px}
.home-last .hl-time{font-size:12px;color:var(--ink3);margin-left:auto;white-space:nowrap}

/* ---------- 与巴菲特对话（专栏）---------- */
.chat-entry{display:flex;align-items:center;gap:16px;margin-top:22px;padding:18px 22px;
  background:linear-gradient(90deg,#faf4e4,#fdf9f0);border:1px solid #e3d5b8;border-left:5px solid var(--accent2);
  border-radius:14px;cursor:pointer;transition:transform .12s,box-shadow .12s}
.chat-entry:hover{transform:translateY(-2px);box-shadow:0 8px 22px rgba(90,70,30,.12)}
.chat-entry .ce-ic{font-size:32px;line-height:1}
.chat-entry .ce-img{width:42px;height:42px;border-radius:12px;box-shadow:0 2px 8px rgba(139,111,71,.25);flex:0 0 auto}
.chat-entry .ce-body{flex:1;min-width:0}
.chat-entry .ce-title{font:700 18px var(--serif);color:var(--accent);margin-bottom:3px}
.chat-entry .ce-desc{font-size:13px;color:var(--ink2);line-height:1.7}
.chat-entry .ce-arrow{font-size:14px;color:var(--accent2);font-weight:600;white-space:nowrap}
.chat-head{display:flex;align-items:center;gap:12px;padding-bottom:12px;border-bottom:1px solid var(--line);flex-wrap:wrap}
.chat-title{font:700 19px var(--serif);flex:1;min-width:220px}
.chat-title .chat-img{width:26px;height:26px;border-radius:7px;vertical-align:-5px;margin-right:8px}
.chat-title .chat-sub{display:block;font:400 12px var(--sans);color:var(--ink3);margin-top:3px}
.chat-intro{background:#faf6ea;border:1px solid #e3d5b8;border-radius:10px;padding:10px 15px;
  font-size:12.5px;color:#4d4638;line-height:1.8;margin:12px 0}
.chat-intro.warn{background:#fdecec;border-color:#f5c2c2;color:#9f1239}
.chat-body{margin-top:12px;display:flex;flex-direction:column;height:calc(100vh - 215px);min-height:430px;
  background:var(--panel);border:1px solid var(--line);border-radius:14px;overflow:hidden}
.chat-msgs{flex:1;overflow-y:auto;padding:16px;display:flex;flex-direction:column;gap:10px}

/* ---------- 主区 ---------- */
#main{flex:1;min-width:0;padding:22px 26px 80px}
#libHead{display:flex;align-items:baseline;gap:10px;margin-bottom:16px;flex-wrap:wrap}
#libTitle{font-size:20px;font-weight:700}
#libCount{font-size:13px;color:var(--ink3)}
.group-title{margin:22px 0 10px;font-size:14px;font-weight:700;color:var(--ink2);
  display:flex;align-items:center;gap:8px}
.group-title::before{content:"";width:4px;height:14px;background:var(--accent2);border-radius:2px}
.group-title .n{font-weight:400;color:var(--ink3);font-size:12px}
.cards{display:grid;grid-template-columns:repeat(auto-fill,minmax(268px,1fr));gap:14px}
.card{background:var(--panel);border:1px solid var(--line);border-radius:12px;padding:14px 16px;
  cursor:pointer;transition:transform .12s,box-shadow .12s,border-color .12s;display:flex;flex-direction:column;gap:8px}
.card:hover{transform:translateY(-2px);box-shadow:0 6px 18px rgba(90,70,30,.09);border-color:#d8c9a8}
.card-head{display:flex;align-items:center;gap:6px}
.badge{font-size:11px;padding:2px 8px;border-radius:20px;font-weight:600;white-space:nowrap}
.badge.cat{background:#f3e6c8;color:var(--accent2)}
.badge.year{background:#e8e4d8;color:var(--ink2)}
.badge.kind{background:#e3ecf7;color:var(--blue)}
.card-fav{margin-left:auto;font-size:15px;color:#d9b64e;cursor:pointer;padding:0 2px}
.card-read{font-size:12px;color:var(--green)}
.rd-tick{color:var(--green);font-weight:700}
.card.is-read .card-title a{color:#a39b8a}
.card-note{font-size:12px;color:var(--green)}
.card-title{font-size:15.5px;font-weight:700;line-height:1.45}
.card-title a{color:var(--ink)}
.card-title a:hover{color:var(--accent2);text-decoration:none}
.card-excerpt{font-size:13px;color:var(--ink2);line-height:1.6;display:-webkit-box;
  -webkit-line-clamp:3;-webkit-box-orient:vertical;overflow:hidden}
.card-excerpt mark{background:var(--hl);color:inherit;border-radius:2px;padding:0 1px}
.card-tags{display:flex;flex-wrap:wrap;gap:5px;margin-top:auto}
.mini-tag{font-size:11px;color:var(--ink3);border:1px solid var(--line2);padding:1px 7px;border-radius:12px}
.rows{display:flex;flex-direction:column}
.row{display:flex;align-items:center;gap:10px;padding:9px 12px;border-bottom:1px solid var(--line2);
  cursor:pointer;border-radius:8px}
.row:hover{background:var(--panel)}
.row .r-title{font-weight:600;font-size:14.5px}
.row .r-meta{font-size:12px;color:var(--ink3);margin-left:auto;white-space:nowrap}
.empty{padding:60px 20px;text-align:center;color:var(--ink3)}
.empty .big{font-size:40px;margin-bottom:10px}

/* ---------- 阅读视图 ---------- */
#reader{display:flex;flex-direction:column;min-height:calc(100vh - 140px)}
#progress{position:fixed;top:57px;left:0;height:3px;background:linear-gradient(90deg,var(--accent2),#d97706);
  width:0;z-index:49;transition:width .1s linear}
.reader-top{display:flex;gap:14px;align-items:flex-start;padding-bottom:14px;border-bottom:1px solid var(--line);flex-wrap:wrap}
.reader-title{flex:1;min-width:240px}
.reader-title h1{font:700 24px/1.4 var(--serif);margin-bottom:6px}
.r-meta{display:flex;gap:8px;align-items:center;flex-wrap:wrap;font-size:12.5px;color:var(--ink3)}
.r-tags{margin-top:8px;display:flex;flex-wrap:wrap;gap:6px}
.r-actions{display:flex;gap:6px;align-items:center}
.reader-body{display:flex;gap:22px;align-items:flex-start;margin-top:18px}
#article{flex:1 1 0%;min-width:0;max-width:none;background:var(--panel);border:1px solid var(--line);
  border-radius:14px;padding:34px 42px 44px;box-shadow:0 2px 10px rgba(90,70,30,.04)}
#article>*+*{margin-top:1em}
#article h1{font:700 22px/1.5 var(--serif);margin-top:.2em;padding-bottom:.35em;border-bottom:2px solid var(--line2)}
#article h2{font:700 19px/1.5 var(--serif);margin-top:1.4em;padding-bottom:.3em;border-bottom:1px solid var(--line2)}
#article h3{font:700 16.5px/1.5 var(--serif);margin-top:1.2em}
#article h4{font-weight:700;margin-top:1.1em;font-size:15px}
#article p{line-height:1.95;text-align:justify}
#article ul,#article ol{margin-left:1.6em;line-height:1.9}
#article li{margin:.25em 0}
#article blockquote{margin:.6em 0;padding:.5em 1.1em;border-left:3px solid var(--accent2);
  background:#faf6ea;border-radius:0 8px 8px 0;color:#4d4638}
#article blockquote p{margin:.35em 0}
#article table{border-collapse:collapse;width:100%;font-size:13.5px;margin:.5em 0;display:block;overflow-x:auto}
#article th,#article td{border:1px solid var(--line);padding:6px 10px;text-align:left;white-space:nowrap}
#article th{background:#f4efe3;font-weight:600}
#article tr:nth-child(even) td{background:#fbf9f3}
#article hr{border:none;border-top:1px dashed var(--line);margin:1.6em 0}
#article code{font:13px/1.5 var(--mono);background:#f1ece0;padding:1px 6px;border-radius:5px;color:#7c2d12}
#article img{max-width:100%;border-radius:8px}
#article a{color:var(--accent2);text-decoration:underline;text-underline-offset:3px;text-decoration-color:#d9c9a8}
#article a:hover{color:#7c2d12}
mark.hl{background:var(--hl);color:inherit;border-radius:2px;padding:0 1px;cursor:pointer}
mark.hl.blue{background:var(--hl2)}
mark.hl.underline{background:transparent;text-decoration:underline;text-decoration-color:#15803d;text-decoration-thickness:2px;text-underline-offset:3px}
#article .fnref{font-size:11px;color:var(--blue);cursor:help;padding:0 1px}
.reader-foot{margin-top:26px;padding-top:16px;border-top:1px dashed var(--line);display:flex;justify-content:space-between;gap:10px;flex-wrap:wrap}
.reader-foot a{font-size:13px}

/* ---------- 右侧面板 ---------- */
#rpanel{flex:0 1 32%;min-width:270px;max-width:440px;position:sticky;top:74px;max-height:calc(100vh - 92px);
  display:flex;flex-direction:column;background:var(--panel);border:1px solid var(--line);border-radius:14px;overflow:hidden}
.rp-tabs{display:flex;border-bottom:1px solid var(--line);background:#faf7f0}
.rp-tab{flex:1;padding:10px 4px;font-size:13.5px;color:var(--ink2);text-align:center;border-bottom:2px solid transparent}
.rp-tab.active{color:var(--accent2);font-weight:700;border-bottom-color:var(--accent2)}
.rp-tab .cnt{font-size:11px;color:var(--ink3);margin-left:2px}
.tabpane{display:none;flex:1;overflow-y:auto;padding:14px}
.tabpane.active{display:block}
#tab-toc .toc-item{display:block;padding:5px 8px;border-radius:6px;font-size:13px;color:var(--ink2);cursor:pointer}
#tab-toc .toc-item:hover{background:var(--line2);color:var(--ink)}
#tab-toc .toc-item.l2{padding-left:20px;font-size:12.5px}
#tab-toc .toc-item.l3{padding-left:32px;font-size:12px;color:var(--ink3)}
.toc-empty{color:var(--ink3);font-size:13px;padding:8px}

/* 笔记 */
.note-sec{margin-bottom:16px}
.note-sec>h4{font-size:12.5px;color:var(--ink3);letter-spacing:.5px;margin-bottom:8px;font-weight:700}
.hl-item,.nt-item{display:flex;gap:8px;align-items:flex-start;padding:8px 9px;border:1px solid var(--line2);
  border-radius:9px;margin-bottom:7px;background:var(--bg);font-size:13px}
.hl-item .sw{width:10px;height:10px;border-radius:3px;margin-top:5px;flex:0 0 10px}
.hl-item .sw.yellow{background:var(--hl)}
.hl-item .sw.blue{background:var(--hl2)}
.hl-item .sw.underline{background:#15803d}
.hl-item .qt,.nt-item .qt{flex:1;color:var(--ink);line-height:1.6}
.hl-item .qt::before{content:"“";color:var(--accent2)}
.hl-item .qt::after{content:"”";color:var(--accent2)}
.nt-item .qt{color:var(--ink2);font-size:12px;border-left:2px solid var(--line);padding-left:7px;margin-bottom:4px}
.nt-item .body{white-space:normal}
.nt-item .body h4,.nt-item .body h5,.ai-msg h4,.ai-msg h5{font-weight:700;margin:.55em 0 .25em;font-size:14px}
.nt-item .body p,.ai-msg p{margin:.35em 0}
.nt-item .body ul,.nt-item .body ol,.ai-msg ul,.ai-msg ol{margin-left:1.3em;padding-left:0}
.nt-item .body li,.ai-msg li{margin:.15em 0}
.nt-item .body pre,.ai-msg pre{background:#f1ece0;border-radius:8px;padding:8px 12px;overflow-x:auto;
  font:12.5px var(--mono);margin:.45em 0;white-space:pre-wrap;word-break:break-word}
.nt-item .body blockquote,.ai-msg blockquote{border-left:3px solid var(--accent2);padding:3px 11px;
  margin:.45em 0;background:#faf6ea;color:#4d4638;border-radius:0 8px 8px 0}
.nt-item .body hr,.ai-msg hr{border:none;border-top:1px dashed var(--line);margin:.7em 0}
.nt-item .body a,.ai-msg a{color:var(--accent2);text-decoration:underline;text-underline-offset:2px}
.x-btn{flex:0 0 auto;color:var(--ink3);padding:1px 5px;border-radius:5px;font-size:13px}
.x-btn:hover{color:#b91c1c;background:#fde8e8}
.edit-btn{flex:0 0 auto;color:var(--ink3);padding:1px 5px;border-radius:5px;font-size:13px}
.edit-btn:hover{color:var(--blue);background:#e8f0fb}
#articleNote{width:100%;min-height:110px;border:1px solid var(--line);border-radius:9px;padding:9px 11px;
  font-size:13px;line-height:1.7;resize:vertical;background:var(--panel);outline:none}
#articleNote:focus{border-color:var(--accent2)}
.note-saved{font-size:11px;color:var(--ink3);margin-top:4px}
.note-empty{color:var(--ink3);font-size:12.5px;line-height:1.8}
.nt-tag{display:inline-block;font-size:10.5px;font-weight:600;padding:1px 7px;border-radius:10px;margin-bottom:4px;letter-spacing:.3px}
.nt-tag-ai{background:#e8f0fb;color:#2563eb}
.nt-tag-bg{background:#f0f6e6;color:#4d7c0f}
.note-sec h4{display:flex;align-items:center;justify-content:space-between}
.note-clear-btn{font-size:11px;font-weight:400;color:var(--ink3);border:1px solid var(--line);
  border-radius:6px;padding:1px 8px;background:var(--panel);cursor:pointer}
.note-clear-btn:hover{color:#b91c1c;border-color:#f5c2c2;background:#fde8e8}
.bg-saved-item{display:flex;align-items:center;gap:7px;padding:4px 8px;border-radius:6px;font-size:12.5px;
  color:var(--ink2);width:100%;text-align:left}
.bg-saved-item:hover{background:var(--line2)}
.bg-saved-label{flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;cursor:pointer}
.bg-saved-del{flex:0 0 auto;color:var(--ink3);padding:1px 5px;border-radius:5px;font-size:12px}
.bg-saved-del:hover{color:#b91c1c;background:#fde8e8}

/* AI 讨论 */
#tab-ai{display:none;flex-direction:column;padding:0}
#tab-ai.active{display:flex}
.ai-msgs{flex:1;overflow-y:auto;padding:14px;display:flex;flex-direction:column;gap:10px}
.ai-msg{max-width:92%;padding:9px 12px;border-radius:12px;font-size:13.5px;line-height:1.75;white-space:pre-wrap;word-break:break-word}
.ai-msg.user{align-self:flex-end;background:var(--accent);color:#fff;border-bottom-right-radius:4px}
.ai-msg.assistant{align-self:flex-start;background:#f1ecdf;border-bottom-left-radius:4px}
.ai-msg.err{align-self:flex-start;background:#fdecec;color:#9f1239;font-size:12.5px;border:1px solid #f5c2c2}
.ai-note-btn{margin-top:7px;font-size:11.5px;color:var(--accent2);border:1px dashed #cbb78d;
  border-radius:12px;padding:2px 10px;background:#fbf7ee;white-space:nowrap;align-self:flex-start}
.ai-note-btn:hover{background:#f3e6c8}
.ai-note-btn.saved{color:var(--green);border-color:#a7c08a;background:#f0f6e6}
.ai-note-btn.saved:hover{background:#e3edd3}
.ai-msg .qref{display:block;font-size:11.5px;color:var(--ink3);margin-bottom:5px;border-left:2px solid var(--line);padding-left:7px}
.ai-msg p{margin:.4em 0}
.ai-msg ul,.ai-msg ol{margin-left:1.3em}
.ai-msg code{font:12px var(--mono);background:#e6dfd0;padding:1px 5px;border-radius:4px}
.ai-chips{display:flex;flex-wrap:wrap;gap:6px;padding:0 14px 8px}
.ai-chip{font-size:12px;padding:5px 10px;border-radius:16px;border:1px dashed #cbb78d;color:var(--accent2);background:#fbf7ee}
.ai-chip:hover{background:#f3e6c8}
.ai-input{display:flex;gap:8px;padding:10px 12px;border-top:1px solid var(--line);background:#faf7f0}
.ai-input textarea{flex:1;border:1px solid var(--line);border-radius:9px;padding:8px 11px;font-size:13.5px;
  line-height:1.6;resize:none;max-height:120px;outline:none;background:var(--panel)}
.ai-input textarea:focus{border-color:var(--accent2)}
#aiSend{padding:8px 16px;border-radius:9px;background:var(--accent2);color:#fff;font-size:13.5px;font-weight:600}
#aiSend:disabled{opacity:.5;cursor:not-allowed}
.ai-status{font-size:12px;color:var(--ink3);padding:0 14px 8px}
.ai-welcome{padding:16px 14px;color:var(--ink3);font-size:13px;line-height:1.8}

/* 背景解释 */
#tab-bg{display:none;flex-direction:column;padding:0}
#tab-bg.active{display:flex}
.bg-head{padding:11px 12px 10px;border-bottom:1px solid var(--line);background:#faf7f0}
.bg-term-row{display:flex;gap:6px;align-items:center}
.bg-term-row label{font-size:12px;color:var(--ink3);white-space:nowrap}
#bgTermInput{flex:1;min-width:0;border:1px solid var(--line);border-radius:8px;padding:6px 10px;
  font-size:13px;background:var(--panel);outline:none}
#bgTermInput:focus{border-color:var(--accent2)}
#bgExplainBtn{padding:6px 14px;border-radius:8px;background:var(--accent2);color:#fff;font-size:13px;font-weight:600;white-space:nowrap}
#bgExplainBtn:disabled{opacity:.5;cursor:not-allowed}
.bg-msgs{flex:1;overflow-y:auto;padding:14px;display:flex;flex-direction:column;gap:10px}
.bg-msg{max-width:96%;padding:10px 13px;border-radius:12px;font-size:13.5px;line-height:1.75;word-break:break-word}
.bg-msg.term{align-self:stretch;max-width:100%;background:#f3e6c8;border:1px solid #e3d5b8;
  border-radius:10px;font-weight:600;color:var(--accent2);text-align:center}
.bg-msg.assistant{align-self:flex-start;background:#f1ecdf;border-bottom-left-radius:4px}
.bg-msg.user{align-self:flex-end;background:var(--accent);color:#fff;border-bottom-right-radius:4px}
.bg-msg.err{align-self:flex-start;background:#fdecec;color:#9f1239;font-size:12.5px;border:1px solid #f5c2c2}
.bg-msg p{margin:.4em 0}
.bg-msg ul,.bg-msg ol{margin-left:1.3em}
.bg-msg code{font:12px var(--mono);background:#e6dfd0;padding:1px 5px;border-radius:4px}
.bg-note-btn{margin-top:8px;font-size:11.5px;color:var(--accent2);border:1px dashed #cbb78d;
  border-radius:12px;padding:3px 12px;background:#fbf7ee;white-space:nowrap;align-self:flex-start}
.bg-note-btn:hover{background:#f3e6c8}
.bg-note-btn.saved{color:var(--green);border-color:#a7c08a;background:#f0f6e6}
.bg-input{display:flex;gap:8px;padding:10px 12px;border-top:1px solid var(--line);background:#faf7f0}
.bg-input textarea{flex:1;border:1px solid var(--line);border-radius:9px;padding:8px 11px;font-size:13.5px;
  line-height:1.6;resize:none;max-height:100px;outline:none;background:var(--panel)}
.bg-input textarea:focus{border-color:var(--accent2)}
#bgSend{padding:8px 16px;border-radius:9px;background:var(--accent2);color:#fff;font-size:13.5px;font-weight:600}
#bgSend:disabled{opacity:.5;cursor:not-allowed}
.bg-status{font-size:12px;color:var(--ink3);padding:0 14px 8px;min-height:20px}
.bg-welcome{padding:8px 4px;color:var(--ink3);font-size:13px;line-height:1.9}
.bg-welcome .bg-tip{background:#faf6ea;border:1px solid var(--line2);border-radius:10px;
  padding:11px 13px;margin-top:12px;font-size:12.5px;line-height:1.8}
.bg-saved{border-top:1px solid var(--line);padding:8px 14px 10px;background:#fdfbf5;max-height:132px;overflow-y:auto}
.bg-saved h4{font-size:11px;color:var(--ink3);font-weight:700;margin-bottom:5px;letter-spacing:.5px}
.bg-saved-item{display:flex;align-items:center;gap:7px;padding:4px 8px;border-radius:6px;font-size:12.5px;
  color:var(--ink2);cursor:pointer;width:100%;text-align:left}
.bg-saved-item:hover{background:var(--line2);color:var(--accent2)}
.bg-saved-item .bg-dot{width:6px;height:6px;border-radius:50%;background:var(--accent2);flex:0 0 6px}
.bg-saved-item .bg-time{margin-left:auto;font-size:10.5px;color:var(--ink3);white-space:nowrap}
/* 笔记中的背景解释条目 */
.bg-item{border:1px solid var(--line2);border-radius:9px;padding:10px 12px;margin-bottom:8px;
  background:var(--bg);font-size:13px}
.bg-item-term{font-weight:700;color:var(--accent2);font-size:13.5px;margin-bottom:5px;display:flex;align-items:center;gap:6px}
.bg-item-body{color:var(--ink);line-height:1.7;font-size:12.5px}
.bg-item-body p{margin:.3em 0}
.bg-item-body ul,.bg-item-body ol{margin-left:1.3em}
.bg-item-qa{margin-top:7px;padding-top:7px;border-top:1px dashed var(--line);font-size:12px;color:var(--ink2)}
.bg-qa-user,.bg-qa-assistant{margin:.35em 0;line-height:1.65}
.bg-qa-user{color:var(--blue)}
.bg-item-actions{display:flex;gap:6px;justify-content:flex-end;margin-top:6px}

/* ---------- 选择工具栏 / 弹窗 / toast ---------- */
#selToolbar{position:fixed;z-index:80;display:flex;gap:2px;background:var(--ink);color:#fff;
  border-radius:9px;padding:4px;box-shadow:0 6px 20px rgba(0,0,0,.25)}
#selToolbar button{padding:5px 10px;border-radius:6px;font-size:12.5px;color:#fff}
#selToolbar button:hover{background:rgba(255,255,255,.18)}
#selToolbar .sep{width:1px;background:rgba(255,255,255,.25);margin:3px 2px}
#modalBackdrop{position:fixed;inset:0;background:rgba(42,38,32,.45);z-index:90;display:flex;align-items:center;justify-content:center;padding:20px}
#modalBackdrop[hidden]{display:none}
#settingsModal{position:fixed;inset:0;z-index:90;display:flex;align-items:center;justify-content:center;
  background:rgba(42,38,32,.45);padding:20px}
#settingsModal[hidden]{display:none}
.modal{background:var(--panel);border-radius:14px;width:100%;max-width:560px;max-height:86vh;overflow-y:auto;
  padding:22px 24px;box-shadow:0 18px 50px rgba(0,0,0,.25)}
.modal h3{font-size:16.5px;margin-bottom:12px}
.modal .field{margin-bottom:14px}
.modal label{display:block;font-size:12.5px;color:var(--ink2);margin-bottom:5px;font-weight:600}
.modal input[type=text],.modal input[type=password],.modal textarea{width:100%;border:1px solid var(--line);
  border-radius:9px;padding:9px 11px;font-size:14px;outline:none;background:var(--bg)}
.modal input:focus,.modal textarea:focus{border-color:var(--accent2)}
.modal .quote-box{background:#faf6ea;border:1px solid var(--line2);border-radius:9px;padding:9px 11px;
  font-size:12.5px;color:var(--ink2);max-height:120px;overflow-y:auto;margin-bottom:10px;line-height:1.7}
.modal-actions{display:flex;gap:8px;justify-content:flex-end;margin-top:16px}
.btn{padding:8px 18px;border-radius:9px;font-size:14px;font-weight:600}
.btn.primary{background:var(--accent2);color:#fff}
.btn.primary:hover{background:#92400e}
.btn.ghost{border:1px solid var(--line);color:var(--ink2)}
.btn.ghost:hover{border-color:var(--accent2);color:var(--accent2)}
.btn.danger{background:#fee2e2;color:#b91c1c}
.hint{font-size:12px;color:var(--ink3);line-height:1.7;margin-top:4px}
#toast{position:fixed;bottom:26px;left:50%;transform:translateX(-50%) translateY(80px);z-index:100;
  background:var(--ink);color:#fff;padding:10px 20px;border-radius:10px;font-size:13.5px;opacity:0;
  transition:transform .25s,opacity .25s;box-shadow:0 8px 24px rgba(0,0,0,.3);max-width:80vw}
#toast.show{transform:translateX(-50%) translateY(0);opacity:1}

/* ---------- 编辑风索引页（#idxView） ---------- */
#idxView{background:#faf8f4;color:#1a1a1a;font:15px/1.7 var(--sans);min-height:calc(100vh - 57px);
  -webkit-font-smoothing:antialiased}
#idxView .iv-hero{padding:60px 24px 44px;text-align:center;border-bottom:1px solid #e8e3db;
  background:linear-gradient(180deg,#fdfbf7 0%,#faf8f4 100%)}
#idxView .iv-hero-inner{max-width:860px;margin:0 auto}
#idxView .iv-eyebrow{font-size:12px;font-weight:500;letter-spacing:3px;text-transform:uppercase;
  color:#8b6f47;margin-bottom:18px}
#idxView .iv-hero h1{font:700 clamp(28px,4.5vw,42px)/1.25 var(--serif);margin-bottom:14px}
#idxView .iv-hero h1 em{font-style:italic;color:#8b6f47}
#idxView .iv-hero-sub{font-size:14.5px;color:#4a4a4a;max-width:560px;margin:0 auto 28px;line-height:1.8}
#idxView .iv-stats{display:flex;justify-content:center;gap:38px;flex-wrap:wrap}
#idxView .iv-stat{text-align:center}
#idxView .iv-stat-num{font:700 30px/1 var(--serif)}
#idxView .iv-stat-label{font-size:11px;color:#8a8580;letter-spacing:1px;text-transform:uppercase;margin-top:6px}
#idxView .iv-nav{position:sticky;top:57px;z-index:40;background:rgba(250,248,244,.94);
  backdrop-filter:blur(10px);-webkit-backdrop-filter:blur(10px);border-bottom:1px solid #e8e3db;padding:0 24px}
#idxView .iv-nav-inner{max-width:860px;margin:0 auto;display:flex;align-items:center;gap:2px;overflow-x:auto}
#idxView .iv-nav-brand{font:700 14px/1 var(--serif);white-space:nowrap;margin-right:18px;padding:13px 0;color:#8b6f47}
#idxView .iv-tab{font-size:13px;font-weight:500;padding:13px 15px;border:none;background:none;
  color:#8a8580;cursor:pointer;white-space:nowrap;position:relative;transition:color .2s}
#idxView .iv-tab:hover{color:#1a1a1a}
#idxView .iv-tab.active{color:#1a1a1a;font-weight:600}
#idxView .iv-tab.active::after{content:'';position:absolute;bottom:0;left:15px;right:15px;height:2px;background:#8b6f47}
#idxView .iv-pane{display:none}
#idxView .iv-pane.active{display:block}
#idxView .iv-section{max-width:860px;margin:0 auto;padding:36px 24px 16px}
#idxView .iv-sec-head{display:flex;align-items:baseline;gap:12px;margin-bottom:4px;flex-wrap:wrap}
#idxView .iv-sec-title{font:700 21px/1.3 var(--serif)}
#idxView .iv-sec-meta{font-size:13px;color:#8a8580}
#idxView .iv-sec-desc{font-size:13.5px;color:#8a8580;margin-bottom:20px;padding-bottom:16px;border-bottom:1px solid #e8e3db}
#idxView .iv-filter-bar{max-width:860px;margin:0 auto;padding:14px 24px;display:flex;gap:8px;
  flex-wrap:wrap;align-items:center;border-bottom:1px solid #e8e3db}
#idxView .iv-filter-label{font-size:11.5px;color:#8a8580;letter-spacing:.5px;text-transform:uppercase;margin-right:2px}
#idxView .iv-filter-group{display:flex;gap:5px;flex-wrap:wrap}
#idxView .iv-filter-btn{font-size:12px;font-weight:500;padding:4px 12px;border:1px solid #e8e3db;
  border-radius:20px;background:transparent;color:#4a4a4a;cursor:pointer;transition:all .2s;white-space:nowrap}
#idxView .iv-filter-btn:hover{border-color:#8b6f47;color:#8b6f47}
#idxView .iv-filter-btn.active{background:#1a1a1a;color:#fff;border-color:#1a1a1a}
#idxView .iv-filter-div{width:1px;height:18px;background:#e8e3db;margin:0 5px}
#idxView .iv-letter{display:grid;grid-template-columns:72px 1fr;gap:20px;padding:24px 0;
  border-bottom:1px solid #e8e3db;transition:background .15s}
#idxView .iv-letter:hover{background:#f5f2ec;margin:0 -24px;padding-left:24px;padding-right:24px}
#idxView .iv-letter-year{font:700 28px/1 var(--serif);text-align:right;padding-top:3px}
#idxView .iv-letter-meta{display:flex;align-items:center;gap:8px;margin-bottom:6px;flex-wrap:wrap}
#idxView .iv-series-badge{font-size:10px;font-weight:600;letter-spacing:.5px;padding:2px 8px;border-radius:3px}
#idxView .iv-series-badge.partnership{background:rgba(107,91,149,.1);color:#6b5b95}
#idxView .iv-series-badge.berkshire{background:rgba(44,95,124,.1);color:#2c5f7c}
#idxView .iv-letter-author{font-size:12.5px;color:#8a8580}
#idxView .iv-letter-title{font:600 17px/1.4 var(--serif);margin-bottom:6px}
#idxView .iv-letter-title a{color:#1a1a1a}
#idxView .iv-letter-title a:hover{color:#8b6f47;text-decoration:none}
#idxView .iv-letter-summary{font-size:13.5px;color:#4a4a4a;line-height:1.75;margin-bottom:9px}
#idxView .iv-letter-tags{display:flex;gap:5px;flex-wrap:wrap;align-items:center}
#idxView .iv-tag{font-size:11px;padding:2px 8px;border-radius:3px;background:#f0ece4;color:#4a4a4a;letter-spacing:.3px}
#idxView .iv-tag.theme{background:rgba(139,111,71,.1);color:#8b6f47}
#idxView .iv-tag.ev-crisis{background:rgba(192,57,43,.08);color:#c0392b;font-weight:500}
#idxView .iv-tag.ev-bubble{background:rgba(211,84,0,.08);color:#d35400;font-weight:500}
#idxView .iv-tag.ev-inflation{background:rgba(184,134,11,.08);color:#b8860b;font-weight:500}
#idxView .iv-tag.ev-war{background:rgba(93,78,55,.08);color:#5d4e37;font-weight:500}
#idxView .iv-tag.ev-pandemic{background:rgba(41,128,185,.08);color:#2980b9;font-weight:500}
#idxView .iv-tag.ev-normal{background:rgba(39,99,42,.06);color:#27632a;font-weight:500}
#idxView .iv-letter-link{display:inline-flex;align-items:center;gap:4px;font-size:12.5px;font-weight:500;
  color:#8b6f47;text-decoration:none;margin-left:auto;transition:gap .2s}
#idxView .iv-letter-link:hover{gap:7px}
#idxView .iv-letter-ctx{font-size:12px;color:#8a8580;margin-top:6px;font-style:italic}
#idxView .iv-card{padding:24px 0;border-bottom:1px solid #e8e3db}
#idxView .iv-card:hover{background:#f5f2ec;margin:0 -24px;padding-left:24px;padding-right:24px;transition:background .15s}
#idxView .iv-card-head{display:flex;align-items:baseline;gap:10px;margin-bottom:9px;flex-wrap:wrap}
#idxView .iv-card-num{font:600 13px/1 var(--serif);color:#8b6f47;min-width:26px}
#idxView .iv-card-title{font:600 17px/1.4 var(--serif)}
#idxView .iv-card-sub{font-size:12.5px;color:#8a8580}
#idxView .iv-card-body{font-size:13.5px;color:#4a4a4a;line-height:1.75;margin-bottom:9px}
#idxView .iv-card-label{font-size:11px;font-weight:600;color:#8a8580;letter-spacing:.5px;
  text-transform:uppercase;margin-top:9px;margin-bottom:4px}
#idxView .iv-card-text{font-size:13px;color:#4a4a4a;line-height:1.7}
#idxView .iv-card-tags{display:flex;gap:5px;flex-wrap:wrap;margin-top:7px}
#idxView .iv-year-tag{font-size:11px;padding:2px 8px;border-radius:3px;background:#f0ece4;color:#4a4a4a}
#idxView .iv-card-row{display:grid;grid-template-columns:100px 1fr;gap:14px;margin-top:7px}
#idxView .iv-card-row-label{font-size:12px;font-weight:600;color:#8a8580;letter-spacing:.3px;padding-top:2px}
#idxView .iv-card-row-value{font-size:13px;color:#4a4a4a;line-height:1.7}
#idxView .iv-event-period{display:inline-block;font-size:12px;font-weight:600;padding:2px 10px;
  border-radius:3px;margin-bottom:7px}
#idxView .iv-period-crisis{background:rgba(192,57,43,.08);color:#c0392b}
#idxView .iv-period-bubble{background:rgba(211,84,0,.08);color:#d35400}
#idxView .iv-period-inflation{background:rgba(184,134,11,.08);color:#b8860b}
#idxView .iv-period-war{background:rgba(93,78,55,.08);color:#5d4e37}
#idxView .iv-period-pandemic{background:rgba(41,128,185,.08);color:#2980b9}
#idxView .iv-period-normal{background:rgba(39,99,42,.06);color:#27632a}
#idxView .iv-method-era{font:600 14px/1 var(--serif);color:#8b6f47;margin-bottom:4px}
#idxView .iv-method-method{font-size:13.5px;font-weight:600;color:#1a1a1a;margin-bottom:7px}
#idxView .iv-footer{max-width:860px;margin:36px auto 0;padding:24px;text-align:center;
  font-size:12.5px;color:#8a8580;line-height:1.8;border-top:1px solid #e8e3db}
#idxView .iv-empty{text-align:center;padding:44px 24px;color:#8a8580}

/* ---------- 响应式 ---------- */
@media (max-width:1100px){
  #rpanel{flex:0 1 34%;min-width:250px;max-width:400px}
  #article{padding:26px 28px}
}
@media (max-width:860px){
  #sidebar{position:fixed;left:0;top:57px;bottom:0;z-index:60;transform:translateX(-100%);
    transition:transform .2s;box-shadow:8px 0 24px rgba(0,0,0,.12);height:auto}
  #sidebar.open{transform:none}
  #sidebarBackdrop{position:fixed;inset:57px 0 0 0;background:rgba(42,38,32,.35);z-index:55;display:none}
  #sidebarBackdrop.show{display:block}
  #menuBtn{display:inline-flex}
  .reader-body{flex-direction:column}
  #rpanel{position:static;width:100%;flex:none;max-height:none;order:2}
  #tab-ai{min-height:420px}
  #tab-bg{min-height:420px}
  #article{max-width:none}
  #idxView .iv-letter{grid-template-columns:56px 1fr;gap:12px}
  #idxView .iv-letter-year{font-size:22px}
  #idxView .iv-card-row{grid-template-columns:1fr;gap:3px}
  #idxView .iv-stats{gap:22px}
  #idxView .iv-stat-num{font-size:24px}
  #idxView .iv-nav{top:57px}
}
@media (min-width:861px){#menuBtn{display:none}}
"""

APP_JS = r"""
'use strict';
/* ================= 工具 ================= */
const $  = s => document.querySelector(s);
const $$ = s => Array.from(document.querySelectorAll(s));
const esc = s => String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2,8);
const store = {
  get(k,d){ try{ const v=localStorage.getItem(k); return v?JSON.parse(v):d; }catch(e){ return d; } },
  set(k,v){ try{ localStorage.setItem(k,JSON.stringify(v)); }catch(e){} },
  del(k){ try{ localStorage.removeItem(k); }catch(e){} }
};
const NOTES_KEY='bf_notes', FAVS_KEY='bf_favs', SETTINGS_KEY='bf_settings', CHAT_KEY='bf_chat';
const CAT_NAME = {index:'索引',partnership:'合伙人信',berkshire:'伯克希尔股东信',special:'特别信件',concept:'概念',company:'公司',person:'人物'};

/* 服务启动时的密钥只从同源 JSON API 读入内存，不回写 localStorage。
   file:// 直接打开时请求失败，继续使用无密钥的 llm-config.js 默认值。 */
let runtimeLlmConfig=null;
async function loadRuntimeLlmConfig(){
  try{
    const r=await fetch('/api/llm-config',{cache:'no-store'});
    if(!r.ok) return false;
    const cfg=await r.json();
    if(!cfg || cfg.app!=='buffett-wisdom') return false;
    runtimeLlmConfig={base:cfg.base||'',key:cfg.key||'',model:cfg.model||''};
    return true;
  }catch(e){ return false; }
}

/* ============ 记忆材料本地持久化（经 /api/state 写本地文件） ============
   笔记 / 收藏 / 已读 / AI 对话 存到 serve 端本地文件（默认用户目录，
   不进 Git）；settings（可能含手动填写的 API Key）只留在浏览器 localStorage，
   不写入 state.json 或仓库文件。 */
const API_STATE='/api/state';
const STATE_KEYS=['notes','favs','read','chat','buffett_chat','munger_chat'];
const STATE_LS_MAP={notes:NOTES_KEY,favs:FAVS_KEY,read:'bf_read',chat:CHAT_KEY,buffett_chat:'bf_buffett_chat',munger_chat:'bf_munger_chat'};
const STATE_DEFAULTS={notes:{},favs:[],read:[],chat:{},buffett_chat:[],munger_chat:[]};
const hasOwn=(obj,key)=>Object.prototype.hasOwnProperty.call(obj,key);
const _nonEmpty = v => { try{ const s=JSON.stringify(v); return s!==undefined && s!=='null' && s!=='{}' && s!=='[]'; }catch(e){ return false; } };
const statePayload=()=>{
  const payload={};
  STATE_KEYS.forEach(k=>payload[k]=store.get(STATE_LS_MAP[k],STATE_DEFAULTS[k]));
  return payload;
};
const pushRemoteStateNow=async(keepalive=false)=>{
  try{
    const r=await fetch(API_STATE,{
      method:'PUT',headers:{'Content-Type':'application/json'},body:JSON.stringify(statePayload()),keepalive
    });
    return r.ok;
  }catch(e){ return false; }
};
const pushRemoteState = (()=>{ let t=null; return ()=>{
  clearTimeout(t);
  t=setTimeout(()=>pushRemoteStateNow(),300);
}; })();
async function loadRemoteState(){
  try{
    const r=await fetch(API_STATE,{cache:'no-store'});
    if(!r.ok) return false;
    const st=await r.json();
    const fileHas=STATE_KEYS.some(k=>hasOwn(st,k));
    if(fileHas){                                  // 本地文件为权威，显式空值也要覆盖 localStorage
      STATE_KEYS.forEach(k=>{ if(hasOwn(st,k)) store.set(STATE_LS_MAP[k],st[k]); });
      // 旧版文件可能缺新字段；保留其 localStorage 值并补齐六键文件。
      if(STATE_KEYS.some(k=>!hasOwn(st,k))) pushRemoteState();
      return true;
    }
    if(STATE_KEYS.some(k=>_nonEmpty(store.get(STATE_LS_MAP[k],null)))) pushRemoteState(); // 首次迁移
    return true;
  }catch(e){ return false; }                      // file:// 直开（无服务）→ 继续用 localStorage
}

/* ================= 数据 ================= */
const DATA = window.BUFFETT_DATA;
const ART = DATA.articles;
const BYID = {};
ART.forEach(a => BYID[a.id]=a);
const IDX = DATA.idx || {topic:[],industry:[],event:[],method:[],year:[]};
const IDX_DIM_NAME = {topic:'主题分类', industry:'行业分类', event:'事件时期', method:'选股方法', year:'年度索引'};
const IDX_DIM_DESC = {topic:'坎宁安主题分类', industry:'涉及行业', event:'市场事件/时期', method:'投资方法演进', year:'逐年索引'};
const plainCache = {};
function toPlain(md){
  let s = md
    .replace(/!\[([^\]]*)\]\([^)]*\)/g,'$1')
    .replace(/\[([^\]]*)\]\([^)]*\)/g,'$1')
    .replace(/\[\^[^\]]*\]\([^)]*\)/g,'')
    .replace(/\[\^[^\]]*\]/g,'')
    .replace(/^#{1,6}\s+/gm,'')
    .replace(/^\s*>\s?/gm,'')
    .replace(/^\s*([-*+]|\d+\.)\s+/gm,'')
    .replace(/^---+$/gm,'')
    .replace(/\|/g,' ')
    .replace(/[*_`~]/g,'')
    .replace(/\s+/g,' ')
    .trim();
  return s;
}
function plainOf(a){ return plainCache[a.id] || (plainCache[a.id]=toPlain(a.md)); }
function decadeOf(a){ return a.year ? Math.floor(a.year/10)*10 + '年代' : '其他'; }
const backlinks = {};
ART.forEach(a => backlinks[a.id]=[]);
ART.forEach(a => (a.links||[]).forEach(([tid])=>{ if(backlinks[tid]) backlinks[tid].push(a.id); }));

/* ================= 本地存储 ================= */
const notesDb = () => store.get(NOTES_KEY,{});
const notesOf  = id => notesDb()[id] || {hls:[],notes:[],articleNote:null,bg:[]};
const saveNotes = (id,obj) => { const db=notesDb(); db[id]=obj; store.set(NOTES_KEY,db); pushRemoteState(); };
const favs = () => store.get(FAVS_KEY,[]);
const isFav = id => favs().includes(id);
const toggleFav = id => {
  let f=favs();
  f = f.includes(id) ? f.filter(x=>x!==id) : f.concat(id);
  store.set(FAVS_KEY,f); pushRemoteState(); return f.includes(id);
};
const READ_KEY='bf_read';
const reads = () => store.get(READ_KEY,[]);
const isRead = id => reads().includes(id);
const toggleRead = id => {
  let r=reads();
  r = r.includes(id) ? r.filter(x=>x!==id) : r.concat(id);
  store.set(READ_KEY,r); pushRemoteState(); return r.includes(id);
};
const hasNotes = id => { const n=notesOf(id); return !!((n.hls&&n.hls.length)||(n.notes&&n.notes.length)||(n.articleNote&&n.articleNote.text)||(n.bg&&n.bg.length)); };
const chatOf = id => store.get(CHAT_KEY,{})[id] || [];
const saveChat = (id,msgs) => { const db=store.get(CHAT_KEY,{}); db[id]=msgs; store.set(CHAT_KEY,db); pushRemoteState(); };
const settings = () => {
  const s = store.get(SETTINGS_KEY,{});
  const c = runtimeLlmConfig || window.BUFFETT_LLM_CONFIG || {};
  return {
    base: s.base || c.base || 'https://api.deepseek.com/v1',
    key:  s.key  || c.key  || '',
    model:s.model|| c.model|| 'deepseek-v4-flash',
  };
};

/* ================= 状态 ================= */
const state = {
  cat:'all', decade:'all', tag:'', q:'', favOnly:false, notedOnly:false,
  sort:'year-asc', group:'none', view:'grid', cur:null, tab:'toc',
  idxDim:null, idxKey:null, idxOpenSet:{},
};

/* ================= Markdown 渲染 ================= */
let _fns = new Map();   // 当前文章的脚注 key → 首个定义（用于角标 tooltip）
let _fnList = [];       // 当前文章的脚注全量列表（用于文末脚注区）
function inlineMd(s){
  let out = esc(s);
  out = out.replace(/`([^`]+)`/g,(m,c)=>{ return '<code>'+c+'</code>'; });
  out = out.replace(/\[\^([^\]]+)\]\(([^)]*)\)/g,(m,k,t)=>{
    if(!_fns.has(k)) _fns.set(k,t);
    _fnList.push([k,t]);
    return '<sup class="fnref" title="'+esc(t)+'">['+esc(k)+']</sup>';
  });
  out = out.replace(/\[\^([^\]]+)\]/g,(m,k)=>{
    const t=_fns.get(k);
    return '<sup class="fnref"'+(t?' title="'+esc(t)+'"':'')+'>['+esc(k)+']</sup>';
  });
  out = out.replace(/!\[([^\]]*)\]\(([^)]+)\)/g,(m,alt,src)=>{
    return '<img alt="'+esc(alt)+'" src="'+esc(src)+'" loading="lazy">';
  });
  out = out.replace(/\[([^\]]+)\]\(([^)]+)\)/g,(m,txt,href)=>{
    const h = href.trim();
    if (h.startsWith('#/a/')) return '<a href="'+esc(h)+'" data-nav="1">'+txt+'</a>';
    if (/^(https?:|mailto:)/.test(h)) return '<a href="'+esc(h)+'" target="_blank" rel="noopener">'+txt+'</a>';
    return txt;
  });
  out = out.replace(/\*\*\*([^*]+)\*\*\*/g,'<strong><em>$1</em></strong>')
           .replace(/\*\*([^*]+)\*\*/g,'<strong>$1</strong>')
           .replace(/(^|[^*])\*([^*\n]+)\*(?!\*)/g,'$1<em>$2</em>')
           .replace(/__([^_\n]+)__/g,'<strong>$1</strong>')
           .replace(/(^|[^_])_([^_\n]+)_(?!_)/g,'$1<em>$2</em>');
  return out;
}

function parseListBlock(lines){
  const items = [];
  for (const ln of lines){
    const m = ln.match(/^(\s*)([-*+]|\d+\.)\s+(.*)$/);
    if (!m){
      if (items.length) items[items.length-1].rest.push(ln.trim());
      continue;
    }
    items.push({ind: m[1].replace(/\t/g,'  ').length, ordered:/^\d/.test(m[2]), text:m[3], rest:[]});
  }
  let html='', curInd=-1;
  for (const it of items){
    if (it.ind > curInd){ html += it.ordered?'<ol>':'<ul>'; }
    else if (it.ind < curInd){
      const diff = curInd - it.ind;
      for (let i=0;i<diff;i++) html += it.ordered?'</ol>':'</ul>';
    }
    const body = it.rest.length ? it.text + '<br>' + it.rest.join('<br>') : it.text;
    html += '<li>'+inlineMd(body)+'</li>';
    curInd = it.ind;
  }
  for (let i=0;i<=curInd;i++) html += '</ul>';
  html = html.replace(/<\/ul><\/ul>$/,'</ul>').replace(/<\/ol><\/ol>$/,'</ol>');
  return html;
}

function mdToHtml(md, fns){
  _fns = new Map();
  _fnList = [];
  if (fns) fns.forEach(([k,t])=>{ if(!_fns.has(k)) _fns.set(k,t); _fnList.push([k,t]); });
  const toc = [];
  let secN = 0;
  const blocks = md.replace(/\r\n?/g,'\n').split(/\n{2,}/);
  let html = '';
  for (let blk of blocks){
    blk = blk.replace(/\n+$/,'');
    if (!blk.trim()) continue;
    const lines = blk.split('\n');

    // 标题
    let m = blk.match(/^(#{1,6})\s+(.*)$/);
    if (m){
      const level = m[1].length, text = m[2];
      secN++;
      const id = 'sec-'+secN;
      if (level>=2 && level<=4) toc.push({l:level, t:text.replace(/[*_`]/g,''), id});
      html += '<h'+level+' id="'+id+'">'+inlineMd(text)+'</h'+level+'>';
      continue;
    }
    // 分隔线
    if (/^\s{0,3}(---|\*\*\*|___)\s*$/.test(blk)){ html += '<hr>'; continue; }
    // 表格
    if (/^\s*\|/.test(lines[0])){
      const rows = lines.map(l=>l.trim().replace(/^\|/,'').replace(/\|$/,'').split('|').map(c=>c.trim()));
      let header=null, body=rows;
      if (rows.length>1 && rows[1].every(c=>/^:?-{2,}:?$/.test(c))){ header=rows[0]; body=rows.slice(2); }
      let t='<table>';
      if (header) t += '<thead><tr>'+header.map(c=>'<th>'+inlineMd(c)+'</th>').join('')+'</tr></thead>';
      t += '<tbody>'+body.map(r=>'<tr>'+r.map(c=>'<td>'+inlineMd(c)+'</td>').join('')+'</tr>').join('')+'</tbody></table>';
      html += t; continue;
    }
    // 引用
    if (lines.every(l=>/^\s*>/.test(l))){
      const inner = lines.map(l=>l.replace(/^\s*>\s?/,'')).filter(x=>x.trim());
      if (inner.length) html += '<blockquote>'+inner.map(x=>'<p>'+inlineMd(x)+'</p>').join('')+'</blockquote>';
      continue;
    }
    // 列表
    if (/^\s{0,4}([-*+]|\d+\.)\s+/.test(lines[0])){
      html += parseListBlock(lines); continue;
    }
    // 段落
    html += '<p>'+inlineMd(lines.join(' '))+'</p>';
  }
  // 文末脚注区
  if (_fnList.length){
    const seen = new Set();
    const items = _fnList.filter(([k,t])=>{ const s=k+'|'+t; if(seen.has(s)) return false; seen.add(s); return true; });
    html += '<section class="footnotes"><hr><h3>脚注</h3><ol>'+
      items.map(([k,t])=>'<li id="fn-'+esc(k)+'"><sup>['+esc(k)+']</sup> '+inlineMd(t)+'</li>').join('')+
      '</ol></section>';
  }
  return {html, toc};
}

/* ================= 搜索 ================= */
let searchIdx = null; // {id: {score, pos}}
function doSearch(q){
  q = q.trim().toLowerCase();
  searchIdx = null;
  if (!q) return;
  searchIdx = {};
  const qlen = q.length;
  for (const a of ART){
    const title = a.title.toLowerCase();
    const tags  = (a.tags||[]).join(' ').toLowerCase();
    const plain = plainOf(a).toLowerCase();
    let score = 0, pos = -1;
    if (title.includes(q)){ score += 100; pos = title.indexOf(q); }
    else if (tags.includes(q)){ score += 40; pos = 0; }
    const p = plain.indexOf(q);
    if (p >= 0){
      score += Math.max(0, 30 - Math.floor(p/2000));
      if (pos < 0) pos = p;
    }
    if (score > 0) searchIdx[a.id] = {score, pos, len:qlen};
  }
}
function snippet(a){
  const q = state.q.trim();
  const plain = plainOf(a);
  const p = plain.toLowerCase().indexOf(q.toLowerCase());
  if (p < 0) return plain.slice(0,130);
  const s = Math.max(0, p-48);
  const e = Math.min(plain.length, p+q.length+90);
  return (s>0?'…':'') + plain.slice(s,e) + (e<plain.length?'…':'');
}

/* ================= 过滤 / 排序 / 分组 ================= */
function visibleList(){
  let list = ART.filter(a=>{
    if (state.cat!=='all' && a.catKey!==state.cat) return false;
    if (state.decade!=='all' && decadeOf(a)!==state.decade) return false;
    if (state.tag && !(a.tags||[]).includes(state.tag)) return false;
    if (state.favOnly && !isFav(a.id)) return false;
    if (state.notedOnly && !hasNotes(a.id)) return false;
    if (state.q && !(searchIdx && searchIdx[a.id])) return false;
    if (state.idxDim && state.idxKey){
      const item=(IDX[state.idxDim]||[]).find(x=>x.k===state.idxKey);
      if (!item) return false;
      const ys = state.idxDim==='year' ? [item.y] : (item.y||[]);
      if (!ys.includes(a.year)) return false;
    }
    return true;
  });
  const s = state.sort;
  const cmp = (x,y) => {
    if (s==='year-desc' || s==='year-asc'){
      const xv = x.year, yv = y.year;
      if (xv !== yv){
        if (xv === null) return 1;      // 无年份文章始终排最后
        if (yv === null) return -1;
        return s==='year-desc' ? yv-xv : xv-yv;
      }
      return x.title.localeCompare(y.title,'zh');
    }
    if (s==='title') return x.title.localeCompare(y.title,'zh');
    if (s==='len') return y.len - x.len;
    if (s==='fav'){ const fx=isFav(x.id)?1:0, fy=isFav(y.id)?1:0; if(fx!==fy) return fy-fx; return (y.year||-1)-(x.year||-1); }
    if (s==='cat'){
      const cx=CAT_ORDER.indexOf(x.catKey), cy=CAT_ORDER.indexOf(y.catKey);
      if (cx!==cy) return cx-cy;
      return (y.year||-1)-(x.year||-1);
    }
    return 0;
  };
  list.sort(cmp);
  if (state.q && searchIdx){
    list.sort((a,b)=> searchIdx[b.id].score - searchIdx[a.id].score);
  }
  return list;
}
const CAT_ORDER = DATA.cats.map(c=>c.key);
function groupKey(a){
  if (state.group==='cat') return CAT_NAME[a.catKey];
  if (state.group==='decade') return decadeOf(a);
  if (state.group==='tag') return (a.tags&&a.tags[0]) || '未标注主题';
  return '';
}
const GROUP_ORDER = {cat:['索引','合伙人信','伯克希尔股东信','特别信件','概念','公司','人物'],
  decade:['1950年代','1960年代','1970年代','1980年代','1990年代','2000年代','2010年代','2020年代','其他'],
  tag:[]};

/* ================= 渲染：库视图 ================= */
const libEl = $('#library'), readerEl = $('#reader');
function renderSidebar(){
  // 分类
  const catEl = $('#sbCats');
  catEl.innerHTML = '';
  const mk = (key,name,n) => {
    const b=document.createElement('button');
    b.className='sb-item'+(state.cat===key?' active':'');
    b.innerHTML='<span>'+name+'</span><span class="n">'+n+'</span>';
    b.onclick=()=>{ state.cat=key; state.tag=''; state.idxDim=null; state.idxKey=null; renderAll(); };
    catEl.appendChild(b);
  };
  mk('all','全部', ART.length);
  DATA.cats.forEach(c=>mk(c.key,c.name,c.count));
  // 年代
  const decEl = $('#sbDecades');
  decEl.innerHTML='';
  const decades=[];
  ART.forEach(a=>{ const d=decadeOf(a); const f=decades.find(x=>x.d===d); if(f) f.n++; else decades.push({d,n:1}); });
  decades.sort((a,b)=>{ const na=parseInt(a.d), nb=parseInt(b.d); if(isNaN(na))return 1; if(isNaN(nb))return -1; return na-nb; });
  const mkd=(d,n)=>{ const b=document.createElement('button'); b.className='sb-item'+(state.decade===d?' active':'');
    b.innerHTML='<span>'+d+'</span><span class="n">'+n+'</span>';
    b.onclick=()=>{ state.decade=(state.decade===d?'all':d); renderAll(); }; decEl.appendChild(b); };
  mkd('all', ART.length);
  decades.forEach(x=>mkd(x.d,x.n));
  // 标签云
  const tagEl = $('#sbTags');
  tagEl.innerHTML='';
  const counts={};
  ART.forEach(a=>(a.tags||[]).forEach(t=>counts[t]=(counts[t]||0)+1));
  const top = Object.entries(counts).sort((a,b)=>b[1]-a[1]).slice(0,40);
  top.forEach(([t,n])=>{
    const c=document.createElement('button');
    c.className='tag-chip'+(state.tag===t?' active':'');
    c.textContent=t+' '+n;
    c.onclick=()=>{ state.tag=(state.tag===t?'':t); renderAll(); };
    tagEl.appendChild(c);
  });
  if(!top.length){ const d=document.createElement('div'); d.className='hint'; d.textContent='（暂无可选主题）'; tagEl.appendChild(d); }
  // 复选框
  $('#sbFav').checked=state.favOnly; $('#sbNoted').checked=state.notedOnly;
  $('#sbFootInfo').textContent = ART.length+' 篇文章 · '+DATA.yearRange[0]+'–'+DATA.yearRange[1]+' · 构建于 '+DATA.built;
  renderIdxSidebar();
}

/* ---------- 分类索引（xlsx 五维）---------- */
function idxYearsOf(it, dim){ return dim==='year' ? [it.y] : (it.y||[]); }
function idxItemCount(it, dim){
  const ys = idxYearsOf(it, dim);
  if(!ys.length) return 0;
  return ART.filter(a=>a.year && ys.includes(a.year)).length;
}
function renderIdxSidebar(){
  const el=$('#sbIdx');
  if(!el) return;
  el.innerHTML='';
  const yItems=IDX.year||[];
  const yRangeEl=$('#idxRange');
  if(yRangeEl){
    const ys=yItems.map(x=>x.y).filter(Boolean);
    yRangeEl.textContent = ys.length ? '('+Math.min.apply(null,ys)+'-'+Math.max.apply(null,ys)+')' : '';
  }
  ['topic','industry','event','method','year'].forEach(dim=>{
    const items=IDX[dim]||[];
    const g=document.createElement('div');
    g.className='idx-group'+(state.idxOpenSet[dim]?' open':'');
    const body=document.createElement('div');
    body.className='idx-items';
    items.forEach(it=>{
      const b=document.createElement('button');
      b.className='sb-item idx-item'+(state.idxDim===dim&&state.idxKey===it.k?' active':'');
      const label = dim==='year' ? (it.y+' 年'+(it.e?' · '+it.e:'')) : it.n;
      b.innerHTML='<span>'+esc(label)+'</span><span class="n">'+idxItemCount(it, dim)+'</span>';
      b.title = IDX_DIM_DESC[dim]+'：'+it.n;
      b.onclick=()=>{
        if(state.idxDim===dim&&state.idxKey===it.k){ state.idxDim=null; state.idxKey=null; }
        else { state.idxDim=dim; state.idxKey=it.k; }
        renderAll();
      };
      body.appendChild(b);
    });
    g.innerHTML='<button class="idx-gtitle"><span>'+IDX_DIM_NAME[dim]+'</span><span class="n">'+items.length+'</span></button>';
    g.appendChild(body);
    g.querySelector('.idx-gtitle').onclick=()=>{
      // 展开状态记入 state，重渲染后保持（点击条目不再自动收回）
      state.idxOpenSet[dim] = !state.idxOpenSet[dim];
      g.classList.toggle('open', !!state.idxOpenSet[dim]);
    };
    el.appendChild(g);
  });
}
function idxActiveItem(){
  if(!state.idxDim||!state.idxKey) return null;
  return (IDX[state.idxDim]||[]).find(x=>x.k===state.idxKey)||null;
}
function idxChips(terms, cls){
  return terms.map(t=>'<button class="mini-tag '+(cls||'')+'" data-it="'+esc(t)+'">'+esc(t)+'</button>').join(' ');
}
function idxBannerHtml(it){
  if(!it) return '';
  const close='<button class="idx-close" title="清除分类">✕</button>';
  let inner='';
  if(state.idxDim==='topic'){
    inner='<h3>'+esc(it.n)+'</h3><p>'+esc(it.d||'')+'</p>'+
      (it.con&&it.con.length?'<div class="idx-chips">'+idxChips(it.con)+'</div>':'')+
      (it.rep?'<p class="idx-rep">代表信件/段落：'+esc(it.rep)+'</p>':'');
  } else if(state.idxDim==='industry'){
    inner='<h3>'+esc(it.n)+'</h3><p>'+esc(it.d||'')+'</p>'+
      (it.co?'<p class="idx-rep">核心公司/标的：'+esc(it.co)+'</p>':'');
  } else if(state.idxDim==='event'){
    inner='<h3>'+esc(it.n)+(it.rng?' <span class="idx-rng">'+esc(it.rng)+'</span>':'')+'</h3>'+
      (it.bg?'<p><b>市场背景：</b>'+esc(it.bg)+'</p>':'')+
      (it.act?'<p><b>巴菲特的观点与行动：</b>'+esc(it.act)+'</p>':'')+
      (it.les?'<p><b>经验教训：</b>'+esc(it.les)+'</p>':'');
  } else if(state.idxDim==='method'){
    inner='<h3>'+esc(it.n)+'</h3>'+(it.m?'<p><b>方法论：</b>'+esc(it.m)+'</p>':'')+
      (it.view?'<p><b>核心观点：</b>'+esc(it.view)+'</p>':'')+
      (it.cases?'<p><b>代表案例：</b>'+esc(it.cases)+'</p>':'')+
      (it.shift?'<p><b>关键转变：</b>'+esc(it.shift)+'</p>':'');
  } else if(state.idxDim==='year'){
    inner='<h3>'+it.y+' 年'+(it.a?' · 撰写人：'+esc(it.a):'')+'</h3>'+
      (it.e?'<p><b>事件标签：</b>'+esc(it.e)+'</p>':'')+
      (it.bg?'<p><b>市场/经济背景：</b>'+esc(it.bg)+'</p>':'')+
      (it.s?'<p class="idx-sum">'+esc(it.s)+'</p>':'')+
      (it.t&&it.t.length?'<div class="idx-chips"><span class="idx-lbl">坎宁安主题：</span>'+idxChips(it.t)+'</div>':'');
  }
  return '<div class="idx-banner">'+inner+close+'</div>';
}
function setIdxByCanonTopic(tok){
  for(const t of (IDX.topic||[])){
    if(t.c===tok || t.c.startsWith(tok)){ state.idxDim='topic'; state.idxKey=t.k; return true; }
  }
  return false;
}

/* ================= 主页 ================= */
const homeEl = $('#home');
function goLibrary(patch){
  // 主页入口 = 全新起点：按入口语义重置其余筛选
  if (patch.cat!==undefined){
    state.cat=patch.cat; state.tag=''; state.idxDim=null; state.idxKey=null; state.decade='all';
    state.q=''; searchIdx=null; $('#q').value=''; $('#qCount').textContent='';
  } else if (patch.idxOpen!==undefined){
    state.cat='all'; state.tag=''; state.idxDim=null; state.idxKey=null; state.decade='all';
    state.idxOpenSet = {}; state.idxOpenSet[patch.idxOpen] = true;   // 主页入口：仅展开所选维度并保持
    state.sort='year-asc'; $('#sortSel').value='year-asc';           // 分类入口：文章按时间从旧到新
    state.q=''; searchIdx=null; $('#q').value=''; $('#qCount').textContent='';
  } else if (patch.tag!==undefined){
    state.tag=patch.tag; state.idxDim=null; state.idxKey=null;
    state.q=''; searchIdx=null; $('#q').value=''; $('#qCount').textContent='';
  }
  if (patch.q!==undefined){
    state.q=patch.q||'';
    doSearch(state.q);
    $('#q').value=state.q;
    $('#qCount').textContent = state.q && searchIdx ? Object.keys(searchIdx).length+' 条' : '';
  }
  location.hash = '#/library';
}
function fmtAgo(ts){
  const d=Date.now()-ts, m=Math.floor(d/60000);
  if(m<1) return '刚刚';
  if(m<60) return m+' 分钟前';
  const h=Math.floor(m/60);
  if(h<24) return h+' 小时前';
  return Math.floor(h/24)+' 天前';
}
function renderHome(){
  const el=homeEl;
  const yearItems=IDX.year||[];
  const years=yearItems.map(x=>x.y).filter(Boolean);
  const yMin=years.length?Math.min.apply(null,years):'';
  const yMax=years.length?Math.max.apply(null,years):'';
  const nLetters=ART.filter(a=>['berkshire','partnership','special'].includes(a.catKey)).length;
  const nTags=new Set();
  ART.forEach(a=>(a.tags||[]).forEach(t=>nTags.add(t)));
  const catCount=k=>{ const c=(DATA.cats||[]).find(x=>x.key===k); return c?c.count:0; };
  const recent=ART.filter(a=>a.catKey==='berkshire'&&a.year&&a.year>=2020).sort((a,b)=>b.year-a.year);
  const tagCounts={};
  ART.forEach(a=>(a.tags||[]).forEach(t=>tagCounts[t]=(tagCounts[t]||0)+1));
  const hotTags=Object.entries(tagCounts).sort((a,b)=>b[1]-a[1]).slice(0,10);
  const last=store.get('bf_last',null);
  const lastArt=last&&BYID[last.id]?BYID[last.id]:null;
  const classics=[
    ['berkshire-1983-巴菲特致股东信','1983','商誉与所有者收益','收购标准、经济商誉、回购的系统论述'],
    ['berkshire-1984-巴菲特致股东信','1984','超级投资者','九位价值投资者业绩反驳有效市场假说'],
    ['berkshire-1987-巴菲特致股东信','1987','市场先生','黑色星期一：波动是机会而非风险'],
    ['berkshire-1994-巴菲特致股东信','1994','大师级框架','理解伯克希尔哲学的最佳单封信'],
    ['berkshire-2008-巴菲特致股东信','2008','买入美国','金融危机中的行动逻辑'],
    ['berkshire-2016-巴菲特致股东信','2016','苹果','能力圈扩展到科技平台'],
    ['berkshire-2021-巴菲特致股东信','2021','美国顺风','长期主义与反对投机'],
    ['berkshire-2023-巴菲特致股东信','2023','纪念芒格','思想伙伴关系的回顾'],
  ].filter(x=>BYID[x[0]]);
  const tiles=[
    ['全部文章','📚',ART.length+' 篇','cat:all'],
    ['合伙人信','📜',catCount('partnership')+' 篇','cat:partnership'],
    ['伯克希尔股东信','🏛',catCount('berkshire')+' 篇','cat:berkshire'],
    ['概念词条','💡',catCount('concept')+' 个','cat:concept'],
    ['公司案例','🏢',catCount('company')+' 家','cat:company'],
    ['人物','👤',catCount('person')+' 位','cat:person'],
  ];
  const dims=[
    ['主题分类','🏷',(IDX.topic||[]).length+' 类','主题：投资/估值/并购…','topic'],
    ['行业分类','🏭',(IDX.industry||[]).length+' 类','保险/铁路/消费/金融…','industry'],
    ['事件时期','🌪',(IDX.event||[]).length+' 段','泡沫、危机、疫情…','event'],
    ['选股方法','🧭',(IDX.method||[]).length+' 期','烟蒂→特许权→护城河…','method'],
    ['年度索引','📅',yearItems.length+' 年',''+yMin+'–'+yMax+' 逐年','year'],
  ];
  el.innerHTML =
    '<div class="home-hero">'+
      (typeof APP_ICON_B64!=='undefined'
        ? '<img class="hh-img" src="data:image/png;base64,'+APP_ICON_B64+'" alt="巴菲特">'
        : '<div class="hh-icon">🏛</div>')+
      '<h1>巴菲特投资智慧</h1>'+
      '<div class="hh-sub">巴菲特致股东信知识库 · 1956–2025 · 完全离线 · 阅读 / 划线 / 笔记 / AI 讨论</div>'+
      '<div class="hh-stats">'+
        '<div><b>'+ART.length+'</b><span>篇文章</span></div>'+
        '<div><b>'+(yMax-yMin+1)+'</b><span>年跨度</span></div>'+
        '<div><b>'+nLetters+'</b><span>封信件</span></div>'+
        '<div><b>'+nTags.size+'</b><span>主题标签</span></div>'+
      '</div>'+
      '<div class="hh-chart-wrap"><div class="hh-chart" id="heroChart"></div><div class="chart-note">口径：1957–1964 合伙基金收益率｜1965–2025 伯克希尔每股账面价值变动率（Book Value；2019–2025 按 10-K 股东权益 ÷ A 股等值股本推算）</div></div>'+
      '<div class="hh-search"><input id="homeQ" placeholder="搜索文章、概念、公司、人物…" autocomplete="off"><button id="homeQGo">搜索</button></div>'+
    '</div>'+
    '<div class="chat-entry" data-chat="1">'+
      (typeof APP_ICON_B64!=='undefined'
        ? '<img class="ce-img" src="data:image/png;base64,'+APP_ICON_B64+'" alt="巴菲特">'
        : '<span class="ce-ic">🗣</span>')+
      '<div class="ce-body"><div class="ce-title">与巴菲特对话</div><div class="ce-desc">以 celebrity-buffett 人格回答你的投资问题——护城河 / 安全边际 / 能力圈 / 决策启发式，先研究再回答</div></div><span class="ce-arrow">开始对话 →</span></div>'+
    '<div class="chat-entry" data-chat="2">'+
      (typeof MUNGER_ICON_B64!=='undefined'
        ? '<img class="ce-img" src="data:image/png;base64,'+MUNGER_ICON_B64+'" alt="芒格">'
        : (typeof APP_ICON_B64!=='undefined'
            ? '<img class="ce-img" src="data:image/png;base64,'+APP_ICON_B64+'" alt="芒格" style="filter:hue-rotate(210deg) saturate(1.2)">'
            : '<span class="ce-ic">🧠</span>'))+
      '<div class="ce-body"><div class="ce-title">与芒格对话</div><div class="ce-desc">以 celebrity-charlie-munger 人格回答投资与人生问题——多元思维模型 / 逆向思维 / 误判心理学，先研究再回答</div></div><span class="ce-arrow">开始对话 →</span></div>'+
    '<section class="home-sec"><h2>📚 快速浏览</h2><div class="home-grid">'+
      tiles.map(t=>'<button class="home-tile" data-go=\''+JSON.stringify({cat:t[3].split(':')[1]})+'\'><div class="ht-ic">'+t[1]+'</div><div class="ht-name">'+t[0]+'</div><div class="ht-desc">'+t[2]+'</div></button>').join('')+
    '</div></section>'+
    '<section class="home-sec"><h2>🗂 分类索引 <span class="n" style="font-size:12px;color:var(--ink3);font-weight:400">'+yMin+'–'+yMax+'</span></h2><div class="home-grid">'+
      dims.map(d=>'<button class="home-tile" data-go=\''+JSON.stringify({idxOpen:d[4]})+'\'><div class="ht-ic">'+d[1]+'</div><div class="ht-name">'+d[0]+'</div><div class="ht-desc">'+d[2]+' · '+d[3]+'</div></button>').join('')+
    '</div></section>'+
    '<section class="home-sec"><h2>⭐ 必读经典</h2><div class="home-list">'+
      classics.map(c=>'<div class="home-row" data-art="'+c[0]+'"><span class="hr-year">'+c[1]+'</span><span class="hr-title">'+esc(c[2])+'</span><span class="hr-why">'+esc(c[3])+'</span></div>').join('')+
    '</div></section>'+
    '<section class="home-sec"><h2>🕘 最新信件</h2><div class="home-list">'+
      recent.map(a=>'<div class="home-row" data-art="'+a.id+'"><span class="hr-year">'+a.year+'</span><span class="hr-title">'+esc(a.title)+'</span><span class="hr-why">'+(a.tags||[]).slice(0,3).map(esc).join(' · ')+'</span></div>').join('')+
    '</div></section>'+
    '<section class="home-sec"><h2>🔥 热门主题</h2><div class="home-chips">'+
      hotTags.map(t=>'<button class="tag-chip" data-tag="'+esc(t[0])+'">'+esc(t[0])+' · '+t[1]+'</button>').join('')+
    '</div></section>'+
    (lastArt?'<div class="home-last" data-art="'+lastArt.id+'"><span class="hl-ic">📖</span><div><div class="hl-title">继续阅读：'+esc(lastArt.title)+'</div><div class="hl-time">上次读到 '+fmtAgo(last.ts||Date.now())+'</div></div><span class="hl-time">→</span></div>':'');
  // 主页搜索
  const hq=$('#homeQ');
  if(hq){
    hq.addEventListener('keydown',e=>{ if(e.key==='Enter'){ goLibrary({q:hq.value.trim()}); } });
    const go=$('#homeQGo');
    if(go) go.onclick=()=>goLibrary({q:hq.value.trim()});
  }
  initHeroChart();
}

/* ============ 主页 Hero 交互图（ECharts，巴菲特历年年度表现） ============ */
let _heroChart=null, _heroChartResize=false;
function initHeroChart(){
  const el=document.getElementById('heroChart');
  if(!el) return;
  const R=window.BUFFETT_RETURNS;
  if(!R || !R.length || typeof echarts==='undefined'){ el.style.display='none'; return; }
  el.style.display='';
  if(_heroChart) _heroChart.dispose();
  _heroChart=echarts.init(el);
  const years=R.map(r=>r[0]);
  const last=R[R.length-1];
  _heroChart.setOption({
    backgroundColor:'transparent',
    tooltip:{
      trigger:'axis',
      backgroundColor:'#fffdf8', borderColor:'#e6dfd1', borderWidth:1,
      padding:[8,12], textStyle:{color:'#2a2620',fontSize:12},
      axisPointer:{type:'cross',lineStyle:{color:'#c9bfa8'},label:{backgroundColor:'#8a7a5c'}}
    },
    legend:{
      top:0, left:72, itemWidth:16, itemHeight:8, itemGap:18,
      textStyle:{color:'#6f675a',fontSize:11},
      data:['年度变动率','累计净值','年化收益'],
      selected:{'年度变动率':true,'累计净值':true,'年化收益':false}
    },
    grid:{left:50, right:62, top:46, bottom:46},
    xAxis:{
      type:'category', data:years,
      axisLine:{lineStyle:{color:'#c9bfa8'}},
      axisTick:{show:false},
      axisLabel:{color:'#9a917f',fontSize:10,interval:4}
    },
    yAxis:[
      {type:'value', name:'年度变动率 %', nameTextStyle:{color:'#9a917f',fontSize:10,padding:[0,0,0,30]},
       splitLine:{lineStyle:{color:'#efe9dd'}},
       axisLabel:{color:'#9a917f',fontSize:10}},
      {type:'log', name:'累计净值', nameTextStyle:{color:'#9a917f',fontSize:10},
       splitLine:{show:false},
       axisLabel:{color:'#9a917f',fontSize:10,
         formatter:function(v){ return v>=10000?Math.round(v/1000)+'k':v; }}}
    ],
    dataZoom:[
      {type:'inside', xAxisIndex:0, start:0, end:100},
      {type:'slider', xAxisIndex:0, height:16, bottom:4,
       borderColor:'#e6dfd1', backgroundColor:'#fffdf8',
       fillerColor:'rgba(161,98,7,.14)', handleStyle:{color:'#a16207'},
       textStyle:{color:'#9a917f',fontSize:9}}
    ],
    series:[
      {name:'年度变动率', type:'bar', data:R.map(r=>r[1]),
       barWidth:'52%',
       itemStyle:{color:function(p){ return p.value<0?'#7c1d1d':'#a16207'; },
                  borderRadius:[2,2,0,0]},
       markLine:{silent:true, symbol:'none',
         data:[{yAxis:0}], lineStyle:{color:'#c9bfa8',type:'dashed',width:1}},
       markPoint:{
         symbol:'pin', symbolSize:44,
         label:{fontSize:10,color:'#fff'},
         data:[
           {coord:[1976,59.3], label:{formatter:'+59.3%'}, itemStyle:{color:'#b45309'}},
           {coord:[2008,-9.6], label:{formatter:'-9.6%'}, itemStyle:{color:'#7c1d1d'}}
         ]}},
      {name:'累计净值', type:'line', yAxisIndex:1, data:R.map(r=>r[2]),
       smooth:true, symbol:'none',
       lineStyle:{color:'#4a6a8a',width:2},
       itemStyle:{color:'#4a6a8a'},
       markPoint:{
         symbol:'pin', symbolSize:44, label:{fontSize:10,color:'#fff'},
         data:[{coord:[last[0],last[2]],
                label:{formatter:function(p){ return Math.round(p.value).toLocaleString()+'×'; }},
                itemStyle:{color:'#4a6a8a'}}]}},
      {name:'年化收益', type:'line', data:R.map(r=>r[3]),
       smooth:true, symbol:'none',
       lineStyle:{color:'#9a917f',width:1.2,type:'dashed'},
       itemStyle:{color:'#9a917f'}}
    ]
  });
  if(!_heroChartResize){
    _heroChartResize=true;
    window.addEventListener('resize',()=>{ if(_heroChart) _heroChart.resize(); });
  }
}
homeEl.addEventListener('click',e=>{
  const ce=e.target.closest('[data-chat]');
  if(ce){ chatMode=(ce.dataset.chat==='2'?'munger':'buffett'); location.hash='#/chat'; return; }
  const art=e.target.closest('[data-art]');
  if(art){ location.hash='#/a/'+art.dataset.art; return; }
  const tg=e.target.closest('[data-tag]');
  if(tg){ goLibrary({tag:tg.dataset.tag}); return; }
  const go=e.target.closest('[data-go]');
  if(go){ try{ goLibrary(JSON.parse(go.dataset.go)); }catch(err){} }
});

function cardHtml(a){
  const n = notesOf(a.id);
  const hasN = hasNotes(a.id);
  const excerpt = state.q ? snippet(a) : plainOf(a).slice(0,130);
  const exHtml = state.q ? esc(excerpt).replace(new RegExp('('+state.q.replace(/[.*+?^${}()|[\]\\]/g,'\\$&')+')','gi'),'<mark>$1</mark>') : esc(excerpt);
  const tags=(a.tags||[]).slice(0,6).map(t=>'<span class="mini-tag">'+esc(t)+'</span>').join('');
  return '<article class="card'+(isRead(a.id)?' is-read':'')+'" data-id="'+esc(a.id)+'">'+
    '<div class="card-head">'+
      '<span class="badge cat">'+CAT_NAME[a.catKey]+'</span>'+
      (a.year?'<span class="badge year">'+a.year+'</span>':'')+
      (a.catKey==='concept'?'<span class="badge kind">概念</span>':'')+
      (a.catKey==='company'?'<span class="badge kind">公司</span>':'')+
      (a.catKey==='person'?'<span class="badge kind">人物</span>':'')+
      (hasN?'<span class="card-note" title="有笔记/高亮">📝</span>':'')+
      (isRead(a.id)?'<span class="card-read" title="已读">📖</span>':'')+
      '<span class="card-fav" title="收藏">'+(isFav(a.id)?'★':'☆')+'</span>'+
    '</div>'+
    '<h3 class="card-title"><a href="#/a/'+esc(a.id)+'">'+(isRead(a.id)?'<span class="rd-tick">✓ </span>':'')+esc(a.title)+'</a></h3>'+
    '<p class="card-excerpt">'+exHtml+'</p>'+
    (tags?'<div class="card-tags">'+tags+'</div>':'')+
  '</article>';
}
function rowHtml(a){
  const tags=(a.tags||[]).slice(0,4).map(t=>'<span class="mini-tag">'+esc(t)+'</span>').join(' ');
  return '<div class="row" data-id="'+esc(a.id)+'">'+
    '<span class="badge cat">'+CAT_NAME[a.catKey]+'</span>'+
    (a.year?'<span class="badge year">'+a.year+'</span>':'')+
    '<span class="r-title">'+esc(a.title)+'</span>'+
    '<span class="r-meta">'+(tags?' '+tags+' ':'')+(isRead(a.id)?'✓ ':'')+(isFav(a.id)?'★':'')+(hasNotes(a.id)?' 📝':'')+'</span>'+
  '</div>';
}

function renderLibrary(){
  readerEl.hidden = true; libEl.hidden = false; homeEl.hidden = true; idxEl.hidden = true; chatViewEl.hidden = true;
  layoutEl.hidden = false;
  const list = visibleList();
  const it = idxActiveItem();
  $('#libTitle').textContent = it
    ? (IDX_DIM_NAME[state.idxDim]||'')+'：'+(it.n||it.y)
    : (state.cat==='all'?'全部文章':CAT_NAME[state.cat]);
  const parts=[];
  if (state.decade!=='all') parts.push(state.decade);
  if (state.tag) parts.push('主题：'+state.tag);
  if (it) parts.push(IDX_DIM_NAME[state.idxDim]);
  if (state.favOnly) parts.push('仅收藏');
  if (state.notedOnly) parts.push('有笔记');
  $('#libCount').textContent = [list.length+' 篇', parts.join(' · ')].filter(Boolean).join(' — ');
  // 分类索引详解横幅
  const banner = $('#idxBanner');
  banner.innerHTML = it ? idxBannerHtml(it) : '';
  banner.onclick = e=>{
    const c=e.target.closest('.idx-close');
    if(c){ state.idxDim=null; state.idxKey=null; renderAll(); return; }
    const chip=e.target.closest('[data-it]');
    if(chip && state.idxDim==='year'){
      if(setIdxByCanonTopic(chip.dataset.it)) renderAll();
    }
  };
  const wrap = $('#libList');
  if (!list.length){
    wrap.innerHTML = '<div class="empty"><div class="big">🔍</div>没有符合条件的文章<br><button class="btn ghost" id="resetFilters" style="margin-top:14px">重置筛选</button></div>';
    const r = $('#resetFilters'); if (r) r.onclick = ()=>{ state.cat='all'; state.decade='all'; state.tag=''; state.q=''; state.favOnly=false; state.notedOnly=false; state.idxDim=null; state.idxKey=null; $('#q').value=''; renderAll(); };
    return;
  }
  if (state.group==='none'){
    wrap.innerHTML = (state.view==='grid'?'<div class="cards">':'<div class="rows">') + list.map(state.view==='grid'?cardHtml:rowHtml).join('') + '</div>';
  } else {
    const groups = {};
    list.forEach(a=>{ const k=groupKey(a); (groups[k]=groups[k]||[]).push(a); });
    const keys = Object.keys(groups);
    const order = GROUP_ORDER[state.group]||[];
    keys.sort((a,b)=>{
      const ia=order.indexOf(a), ib=order.indexOf(b);
      if (ia>=0&&ib>=0) return ia-ib;
      if (ia>=0) return -1; if (ib>=0) return 1;
      return a.localeCompare(b,'zh');
    });
    wrap.innerHTML = keys.map(k=>
      '<div class="group-title">'+esc(k)+'<span class="n">'+groups[k].length+' 篇</span></div>'+
      (state.view==='grid'?'<div class="cards">':'<div class="rows">')+
      groups[k].map(state.view==='grid'?cardHtml:rowHtml).join('')+'</div>'
    ).join('');
  }
  // 事件委托：卡片/行点击
  wrap.onclick = e => {
    const fav = e.target.closest('.card-fav');
    if (fav){
      e.preventDefault(); e.stopPropagation();
      const card = fav.closest('.card'); const id = card.dataset.id;
      const on = toggleFav(id);
      fav.textContent = on?'★':'☆';
      toast(on?'已收藏':'已取消收藏');
      if (state.favOnly) renderLibrary();
      return;
    }
    const el = e.target.closest('.card, .row');
    if (el) location.hash = '#/a/'+el.dataset.id;
  };
}

/* ================= 渲染：阅读视图 ================= */
function openArticle(id){
  const a = BYID[id];
  if (!a) return;
  state.cur = id;
  libEl.hidden = true; readerEl.hidden = false; homeEl.hidden = true; idxEl.hidden = true; chatViewEl.hidden = true;
  layoutEl.hidden = false;
  store.set('bf_last', {id, ts: Date.now()});
  $('#rTitle').textContent = a.title;
  $('#rMeta').innerHTML =
    '<span class="badge cat">'+CAT_NAME[a.catKey]+'</span>'+
    (a.year?'<span class="badge year">'+a.year+'</span>':'')+
    '<span>'+Math.max(1,Math.round(plainOf(a).length/500))+' 千字</span>'+
    '<span>'+(a.links||[]).length+' 处关联</span>';
  $('#rTags').innerHTML = (a.tags||[]).map(t=>
    '<button class="tag-chip" data-tag="'+esc(t)+'">'+esc(t)+'</button>').join('');
  $('#rTags').onclick = e=>{
    const b=e.target.closest('[data-tag]');
    if(!b) return;
    state.tag=b.dataset.tag; state.q=''; $('#q').value=''; $('#qCount').textContent='';
    location.hash='#/library';
  };
  // 年度索引行（信件类文章且有年度总索引条目时显示）
  const yi = (a.catKey==='partnership'||a.catKey==='berkshire')
    ? (IDX.year||[]).find(x=>x.y===a.year) : null;
  const idxLine = $('#rIdxLine');
  if (yi){
    idxLine.style.display='';
    idxLine.innerHTML = '<span>📋 年度索引</span>'+
      (yi.e?'<span class="idx-ev">事件：<b>'+esc(yi.e)+'</b></span>':'')+
      (yi.t&&yi.t.length?'<span class="idx-ev">坎宁安主题：'+(yi.t||[]).map(t=>'<button class="mini-tag" data-ct="'+esc(t)+'">'+esc(t)+'</button>').join(' ')+'</span>':'')+
      '<button class="mini-tag" id="rIdxYear" title="查看年度摘要">年度摘要 →</button>';
    idxLine.onclick = e=>{
      const ct=e.target.closest('[data-ct]');
      if(ct){ if(setIdxByCanonTopic(ct.dataset.ct)){ state.q=''; $('#q').value=''; $('#qCount').textContent=''; location.hash='#/library'; } return; }
      if(e.target.closest('#rIdxYear')){ state.idxDim='year'; state.idxKey=yi.k; location.hash='#/library'; }
    };
  } else {
    idxLine.innerHTML='';
    idxLine.style.display='none';
  }
  $('#rRead').textContent = isRead(id)?'✅ 已读':'📖 未读';
  $('#rRead').classList.toggle('active', isRead(id));
  $('#rRead').onclick = ()=>{
    const on=toggleRead(id);
    $('#rRead').textContent=on?'✅ 已读':'📖 未读';
    $('#rRead').classList.toggle('active', on);
    toast(on?'已标记为已读':'已取消已读标记');
  };
  $('#rFav').textContent = isFav(id)?'★':'☆';
  $('#rFav').onclick = ()=>{ const on=toggleFav(id); $('#rFav').textContent=on?'★':'☆'; toast(on?'已收藏':'已取消收藏'); };

  // 正文
  const {html, toc} = mdToHtml(a.md, a.fns||[]);
  const artEl = $('#article');
  artEl.innerHTML = html;
  // 脚注数量显示 & 尾部说明
  const fnCount = artEl.querySelectorAll('.fnref').length;
  if (fnCount) {
    const note = document.createElement('p');
    note.className='footnote-hint';
    note.style.cssText='font-size:12px;color:#9a917f;margin-top:2em';
    note.textContent = '本文含 '+fnCount+' 处脚注（悬停角标查看）。';
    artEl.appendChild(note);
  }
  applyHighlights(artEl, notesOf(id).hls||[]);
  buildToc(toc);
  renderNotesPanel();
  renderChatPanel();
  bgRequestSeq++; bgState.busy=false;         // 切文立即使旧背景请求失效，由新文章重新接管状态
  bgState.term=''; bgState.ctx=''; bgState.msgs=[]; bgState.saved=false;
  $('#bgTermInput').value='';
  renderBgPanel();
  updateReaderNav();
  // 上一篇/下一篇
  $('#rPrev').onclick = ()=>navArticle(-1);
  $('#rNext').onclick = ()=>navArticle(1);
  $('#backBtn').onclick = ()=>{ location.hash='#/library'; };
  document.title = a.title+' · 巴菲特投资智慧';
  window.scrollTo(0,0);
}
function navArticle(dir){
  const list = visibleList();
  const idx = list.findIndex(x=>x.id===state.cur);
  if (idx<0) return;
  const nxt = list[idx+dir];
  if (nxt) location.hash = '#/a/'+nxt.id;
  else toast(dir>0?'已经是最后一篇':'已经是第一篇');
}
function updateReaderNav(){
  const list = visibleList();
  const idx = list.findIndex(x=>x.id===state.cur);
  const cur = BYID[state.cur];
  const prev = idx>0?list[idx-1]:null, next = idx>=0&&idx<list.length-1?list[idx+1]:null;
  $('#rPrev').textContent = prev?('← '+(prev.title.length>14?prev.title.slice(0,14)+'…':prev.title)):'← 已是第一篇';
  $('#rPrev').disabled = !prev;
  $('#rNext').textContent = next?((next.title.length>14?next.title.slice(0,14)+'…':next.title)+' →'):'已是最后一篇 →';
  $('#rNext').disabled = !next;
  $('#rNavInfo').textContent = cur?('第 '+(idx+1)+' / '+list.length+' 篇'):'';
}
function buildToc(toc){
  const el = $('#tab-toc');
  if (!toc.length){ el.innerHTML='<div class="toc-empty">本文无小节目录</div>'; return; }
  el.innerHTML = toc.map(t=>
    '<a class="toc-item l'+(t.l-1)+'" data-sec="'+t.id+'">'+esc(t.t)+'</a>').join('');
  el.onclick = e=>{
    const it=e.target.closest('.toc-item');
    if(!it) return;
    const h=document.getElementById(it.dataset.sec);
    if(h) h.scrollIntoView({behavior:'smooth',block:'start'});
  };
}

/* ================= 高亮 ================= */
function findTextRange(root, target){
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  const nodes=[], starts=[];
  let acc='';
  while(walker.nextNode()){
    const n=walker.currentNode;
    const d = n.data.replace(/\s+/g,' ');
    if(!d) continue;
    nodes.push({n,d}); starts.push(acc.length); acc+=d;
  }
  const idx = acc.indexOf(target);
  if(idx<0) return null;
  const end = idx+target.length;
  let si=0; while(si<nodes.length-1 && starts[si+1]<=idx) si++;
  let ei=nodes.length-1; while(ei>0 && starts[ei]>=end) ei--;
  return {startNode:nodes[si].n, startOff:idx-starts[si], endNode:nodes[ei].n, endOff:end-starts[ei]};
}
function applyHighlights(container, hls){
  if(!hls||!hls.length) return;
  for(const hl of hls){
    if(!hl.text) continue;
    const f = findTextRange(container, hl.text.replace(/\s+/g,' '));
    if(!f) continue;
    try{
      const range=document.createRange();
      range.setStart(f.startNode,f.startOff);
      range.setEnd(f.endNode,f.endOff);
      const mark=document.createElement('mark');
      mark.className='hl '+(hl.color||'yellow');
      mark.dataset.hlid=hl.id;
      const frag=range.extractContents();
      mark.appendChild(frag);
      range.insertNode(mark);
    }catch(e){}
  }
}
function currentSelection(){
  const sel=window.getSelection();
  if(!sel||sel.isCollapsed) return null;
  if(!$('#article').contains(sel.anchorNode)||!$('#article').contains(sel.focusNode)) return null;
  const text=sel.toString().replace(/\s+/g,' ').trim();
  return text?text:null;
}
function selectionContext(){
  const sel=window.getSelection();
  if(!sel||sel.isCollapsed) return null;
  if(!$('#article').contains(sel.anchorNode)||!$('#article').contains(sel.focusNode)) return null;
  const text=sel.toString().replace(/\s+/g,' ').trim();
  if(!text) return null;
  let node=sel.anchorNode;
  while(node&&node.nodeType!==1) node=node.parentNode;
  while(node&&node!==$('#article')&&!/^(P|LI|BLOCKQUOTE|TD|TH|H[1-6])$/.test(node.tagName)) node=node.parentNode;
  const para=node?node.textContent.replace(/\s+/g,' ').trim().slice(0,500):'';
  return {term:text,context:para};
}
function showSelToolbar(x,y){
  const tb=$('#selToolbar');
  tb.hidden=false;
  const w=tb.offsetWidth;
  let left=Math.min(x, window.innerWidth-w-10);
  let top=y+14;
  if(top+tb.offsetHeight>window.innerHeight) top=y-tb.offsetHeight-10;
  tb.style.left=left+'px'; tb.style.top=top+'px';
}
function hideSelToolbar(){ $('#selToolbar').hidden=true; }
$('#article').addEventListener('mouseup', e=>{
  setTimeout(()=>{
    const text=currentSelection();
    if(text) showSelToolbar(e.clientX,e.clientY); else hideSelToolbar();
  },10);
});
document.addEventListener('mousedown', e=>{
  if(!e.target.closest('#selToolbar')) hideSelToolbar();
});
document.addEventListener('keydown', e=>{ if(e.key==='Escape'){ hideSelToolbar(); closeModal(); } });

/* ---------- 笔记 ---------- */
let editingNoteId=null;
function addHighlight(color){
  const text=currentSelection();
  if(!text){ toast('请先选中文字'); return; }
  const firstLine = text.split('\n')[0].slice(0,300);
  if(!firstLine){ hideSelToolbar(); return; }
  const id=state.cur;
  const n=notesOf(id);
  n.hls=n.hls||[]; n.hls.push({id:uid(), text:firstLine, color, ts:Date.now()});
  saveNotes(id,n);
  hideSelToolbar();
  toast('已添加'+(color==='underline'?'下划线':'高亮'));
  reapplyHighlights();
}
function reapplyHighlights(){
  const artEl=$('#article');
  // 清掉旧 mark（按 data-hlid 精确移除，避免误伤脚注等）
  $$('#article mark.hl').forEach(m=>m.replaceWith(document.createTextNode(m.textContent)));
  applyHighlights(artEl, notesOf(state.cur).hls||[]);
  renderNotesPanel();
}
function openNoteModal(quote, noteId, prefillText){
  editingNoteId=noteId||null;
  $('#nmQuote').textContent=quote||'';
  $('#nmText').value=prefillText||'';
  if(noteId){
    const n=notesOf(state.cur);
    const it=(n.notes||[]).find(x=>x.id===noteId);
    if(it){ $('#nmQuote').textContent=it.quote||''; $('#nmText').value=it.text||''; }
  }
  $('#nmTitle').textContent = noteId ? '编辑笔记' : (prefillText ? '将内容加入笔记' : '新建笔记');
  $('#modalBackdrop').hidden=false;
  $('#nmText').focus();
}
function closeModal(){ $('#modalBackdrop').hidden=true; editingNoteId=null; }
function saveNoteFromModal(){
  const text=$('#nmText').value.trim();
  const quote=$('#nmQuote').textContent.trim();
  if(!text&&!quote){ toast('笔记为空'); return; }
  const id=state.cur, n=notesOf(id);
  n.notes=n.notes||[];
  if(editingNoteId){
    const it=n.notes.find(x=>x.id===editingNoteId);
    if(it){ it.text=text; it.quote=quote; }
  } else {
    n.notes.push({id:uid(), quote, text, ts:Date.now()});
  }
  saveNotes(id,n);
  closeModal();
  toast('笔记已保存');
  renderNotesPanel();
}
let articleNoteSaveTimer=null;
let pendingArticleNote=null;
function flushArticleNoteSave(){
  if(!pendingArticleNote) return;
  clearTimeout(articleNoteSaveTimer);
  articleNoteSaveTimer=null;
  const pending=pendingArticleNote;
  pendingArticleNote=null;
  const latest=notesOf(pending.id);
  latest.articleNote={text:pending.text,ts:Date.now()};
  saveNotes(pending.id,latest);
  if(state.cur===pending.id){
    const saved=$('#noteSaved');
    if(saved) saved.textContent='已保存 '+new Date().toLocaleTimeString('zh-CN',{hour:'2-digit',minute:'2-digit'});
  }
}
function renderNotesPanel(){
  const id=state.cur;
  if(!id) return;
  // 若其他笔记操作触发重绘，先保存用户已输入但尚未到 600ms 的文章笔记。
  flushArticleNoteSave();
  const n=notesOf(id);
  const hlEl=$('#ntHighlights'), ntEl=$('#ntNotes');
  hlEl.innerHTML='';
  (n.hls||[]).forEach(h=>{
    const d=document.createElement('div');
    d.className='hl-item';
    d.innerHTML='<span class="sw '+(h.color||'yellow')+'"></span><span class="qt">'+esc(h.text)+'</span>'+
      '<button class="x-btn" title="删除高亮">✕</button>';
    d.querySelector('.x-btn').onclick=()=>{
      n.hls=n.hls.filter(x=>x.id!==h.id); saveNotes(id,n);
      reapplyHighlights(); toast('已删除高亮');
    };
    hlEl.appendChild(d);
  });
  if(!(n.hls||[]).length) hlEl.innerHTML='<div class="note-empty">还没有划线高亮。阅读时选中文字，即可高亮/下划线。</div>';
  ntEl.innerHTML='';
  (n.notes||[]).forEach(nt=>{
    const d=document.createElement('div');
    d.className='nt-item'+(nt.source?' nt-src-'+nt.source:'');
    const srcLabel=nt.source==='ai'?'<span class="nt-tag nt-tag-ai">🤖 AI 答复</span>'
      :nt.source==='bg'?'<span class="nt-tag nt-tag-bg">🔍 背景解释</span>':'';
    d.innerHTML='<div style="flex:1">'+srcLabel+
      (nt.quote?'<div class="qt">'+esc(nt.quote.slice(0,180))+'</div>':'')+
      '<div class="body">'+mdRich(nt.text||'')+'</div></div>'+
      '<button class="edit-btn" title="编辑">✎</button>'+
      '<button class="x-btn" title="删除">✕</button>';
    d.querySelector('.edit-btn').onclick=()=>openNoteModal(null,nt.id);
    d.querySelector('.x-btn').onclick=()=>{
      n.notes=n.notes.filter(x=>x.id!==nt.id); saveNotes(id,n);
      // 同步清除 AI 聊天 / 背景解释中对应消息的 noteId
      const cm=chatOf(id);
      cm.forEach(x=>{ if(x.noteId===nt.id) x.noteId=null; });
      saveChat(id,cm);
      if(bgState.msgs) bgState.msgs.forEach(x=>{ if(x.noteId===nt.id) x.noteId=null; });
      renderNotesPanel(); renderChatPanel(); renderBgPanel();
      toast('已删除笔记');
    };
    ntEl.appendChild(d);
  });
  if(!(n.notes||[]).length) ntEl.innerHTML='<div class="note-empty">还没有笔记。选中文字 → 「笔记」，或点击下方按钮。</div>'+
    '<button class="btn ghost" style="margin-top:8px;font-size:12.5px" id="newNoteBtn">+ 写一条笔记</button>';
  else ntEl.insertAdjacentHTML('beforeend','<button class="btn ghost" style="margin-top:8px;font-size:12.5px" id="newNoteBtn">+ 写一条笔记</button>');
  const nb=$('#newNoteBtn'); if(nb) nb.onclick=()=>openNoteModal('');
  // 背景解释
  const bgEl=$('#ntBg');
  bgEl.innerHTML='';
  (n.bg||[]).slice().sort((a,b)=>b.ts-a.ts).forEach(b=>{
    const d=document.createElement('div');
    d.className='bg-item';
    let qaHtml='';
    if(b.qa&&b.qa.length){
      qaHtml='<div class="bg-item-qa">'+b.qa.map(m=>
        '<div class="bg-qa-'+m.role+'">'+(m.role==='user'?'🙋 我：':'🤖 AI：')+esc(m.content.slice(0,400))+'</div>'
      ).join('')+'</div>';
    }
    d.innerHTML='<div class="bg-item-term">🔍 '+esc(b.term)+'</div>'+
      '<div class="bg-item-body">'+mdLight(b.explanation)+'</div>'+qaHtml+
      '<div class="bg-item-actions"><button class="edit-btn" data-reopen="'+b.id+'">查看</button>'+
      '<button class="x-btn" data-delbg="'+b.id+'">✕</button></div>';
    bgEl.appendChild(d);
  });
  if(!(n.bg||[]).length) bgEl.innerHTML='<div class="note-empty">还没有保存背景解释。阅读时选中词语 →「🔍 背景解释」，可把解释保存到这里。</div>';
  bgEl.onclick=e=>{
    const del=e.target.closest('[data-delbg]');
    if(del){ n.bg=n.bg.filter(x=>x.id!==del.dataset.delbg); saveNotes(id,n); renderNotesPanel(); renderBgPanel(); toast('已删除'); return; }
    const reopen=e.target.closest('[data-reopen]');
    if(reopen){
      const b=(n.bg||[]).find(x=>x.id===reopen.dataset.reopen);
      if(b){
        bgState.term=b.term; bgState.ctx=''; bgState.saved=true;
        bgState.msgs=[{role:'assistant',content:b.explanation,ts:b.ts}].concat(b.qa||[]);
        $('#bgTermInput').value=b.term;
        switchToBgTab();
        renderBgPanel();
      }
    }
  };
  // 文章笔记
  const an=$('#articleNote');
  an.value = (n.articleNote&&n.articleNote.text)||'';
  const saved=$('#noteSaved');
  an.oninput=()=>{
    saved.textContent='保存中…';
    clearTimeout(articleNoteSaveTimer);
    pendingArticleNote={id,text:an.value};
    articleNoteSaveTimer=setTimeout(flushArticleNoteSave,600);
  };
  saved.textContent = n.articleNote&&n.articleNote.text ? '上次保存 '+new Date(n.articleNote.ts).toLocaleString('zh-CN',{hour:'2-digit',minute:'2-digit'}) : '自动保存';
  const clrBtn=$('#clearArticleNote');
  if(clrBtn) clrBtn.onclick=()=>{
    if(!an.value.trim()){ toast('文章笔记为空'); return; }
    if(pendingArticleNote&&pendingArticleNote.id===id){
      clearTimeout(articleNoteSaveTimer);
      articleNoteSaveTimer=null;
      pendingArticleNote=null;
    }
    an.value=''; n.articleNote=null; saveNotes(id,n);
    saved.textContent='已清空';
    renderNotesPanel(); toast('已清空文章笔记');
  };
  const cnt=(n.hls||[]).length+(n.notes||[]).length+(n.bg||[]).length+(n.articleNote&&n.articleNote.text?1:0);
  $('#notesCnt').textContent=cnt?'('+cnt+')':'';
}
$('#nmCancel').onclick=closeModal;
$('#nmSave').onclick=saveNoteFromModal;
$('#modalBackdrop').addEventListener('mousedown',e=>{ if(e.target.id==='modalBackdrop') closeModal(); });
$('#hlYellow').onclick=()=>addHighlight('yellow');
$('#hlBlue').onclick=()=>addHighlight('blue');
$('#hlUnderline').onclick=()=>addHighlight('underline');
$('#hlNote').onclick=()=>{
  const text=currentSelection();
  hideSelToolbar();
  if(text) openNoteModal(text.slice(0,400));
  else toast('请先选中文字');
};
$('#hlCopy').onclick=()=>{
  const text=currentSelection();
  hideSelToolbar();
  if(!text) return;
  if(navigator.clipboard&&navigator.clipboard.writeText){ navigator.clipboard.writeText(text).then(()=>toast('已复制'),()=>toast('复制失败')); }
  else toast('复制失败');
};
$('#hlBg').onclick=()=>{
  const sel=selectionContext();
  hideSelToolbar();
  if(!sel){ toast('请先选中要解释的词语'); return; }
  switchToBgTab();
  explainBgTerm(sel.term, sel.context);
};

/* ================= 导出笔记 ================= */
function download(name, content, mime){
  const blob=new Blob([content],{type:mime||'text/plain;charset=utf-8'});
  const url=URL.createObjectURL(blob);
  const a=document.createElement('a');
  a.href=url; a.download=name;
  document.body.appendChild(a); a.click();
  setTimeout(()=>{ URL.revokeObjectURL(url); a.remove(); },300);
}
function exportNotes(){
  const db=notesDb();
  const ids=Object.keys(db).filter(id=>hasNotes(id));
  if(!ids.length){ toast('还没有任何笔记'); return; }
  const lines=['# 巴菲特投资智慧 · 我的笔记','', '导出时间：'+new Date().toLocaleString('zh-CN'),'共 '+ids.length+' 篇文章有笔记',''];
  ids.sort((a,b)=>((BYID[b]||{}).year||-1)-((BYID[a]||{}).year||-1));
  ids.forEach(id=>{
    const a=BYID[id]; if(!a) return;
    const n=db[id];
    lines.push('## '+a.title+'（'+(CAT_NAME[a.catKey]||'')+(a.year?' · '+a.year:'')+'）');
    (n.hls||[]).forEach(h=>lines.push('- 高亮：「'+h.text+'」'));
    (n.notes||[]).forEach(nt=>{
      lines.push('- 笔记：'+(nt.quote?'\n  > 「'+nt.quote+'」\n  ':'')+nt.text);
    });
    (n.bg||[]).forEach(b=>{
      lines.push('- 背景解释：「'+b.term+'」');
      lines.push('  '+b.explanation.replace(/\n/g,'\n  '));
      (b.qa||[]).forEach(m=>{
        lines.push('  - '+(m.role==='user'?'追问：':'回答：')+m.content.replace(/\n/g,'\n    '));
      });
    });
    if(n.articleNote&&n.articleNote.text) lines.push('- 文章笔记：\n  '+n.articleNote.text.replace(/\n/g,'\n  '));
    lines.push('');
  });
  download('巴菲特投资智慧-笔记.md', lines.join('\n'));
  const json={exportedAt:new Date().toISOString(), notes:db};
  download('巴菲特投资智慧-笔记.json', JSON.stringify(json,null,2), 'application/json');
  toast('已导出笔记（.md 与 .json）');
}

/* ================= AI 讨论 ================= */
function settingsOk(){ return !!(settings().key); }
function renderChatPanel(){
  const id=state.cur;
  const a=BYID[id]||{title:''};
  const box=$('#aiMsgs');
  box.innerHTML='';
  const msgs=chatOf(id);
  if(!settingsOk()){
    box.innerHTML='<div class="ai-welcome">🤖 与 AI 讨论需要配置大模型接口（DeepSeek 等 OpenAI 兼容接口）。<br><br>点击下方按钮打开设置填写 API Key（仅保存在本机浏览器）。<br><br><button class="btn primary" id="aiOpenSettings" style="font-size:13px">⚙ 打开设置</button></div>';
    const b=$('#aiOpenSettings'); if(b) b.onclick=openSettings;
    $('#aiChips').innerHTML=''; $('#aiInput').style.display='none'; $('#aiStatus').textContent='';
    return;
  }
  $('#aiInput').style.display='flex';
  $('#aiStatus').textContent='';
  msgs.forEach(m=>{
    const d=document.createElement('div');
    d.className='ai-msg '+m.role;
    if(m.quote) d.innerHTML='<span class="qref">'+esc(m.quote.slice(0,200))+'</span>'+mdRich(m.content);
    else d.innerHTML=mdRich(m.content);
    if(m.role==='assistant'&&!m.pending){
      // 检查该条答复是否已保存到笔记（笔记可能在笔记面板被删除）
      const curNotes=notesOf(id);
      const stillSaved=m.noteId && (curNotes.notes||[]).some(x=>x.id===m.noteId);
      const btn=document.createElement('button');
      btn.className='ai-note-btn'+(stillSaved?' saved':'');
      btn.textContent=stillSaved?'✓ 已保存笔记':'📝 保存到笔记';
      btn.title=stillSaved?'点击从笔记中删除这条答复':'仅把这条 AI 答复保存到笔记';
      btn.onclick=()=>toggleAiNote(m,btn,id);
      d.appendChild(btn);
    }
    box.appendChild(d);
  });
  if(!msgs.length){
    $('#aiChips').innerHTML=[
      '如何把本文思想应用到 A 股选股？',
      '结合本文给我 3 条可落地实践建议',
      '本文观点如何对应量化因子/回测假设？',
      '结合我的笔记，提炼一份行动清单',
    ].map(t=>'<button class="ai-chip">'+esc(t)+'</button>').join('');
    $('#aiChips').onclick=e=>{
      const c=e.target.closest('.ai-chip');
      if(c){ $('#aiInputBox').value=c.textContent; sendAi(); }
    };
  } else $('#aiChips').innerHTML='';
  box.scrollTop=box.scrollHeight;
}
function mdLight(t){
  return esc(t)
    .replace(/\*\*([^*]+)\*\*/g,'<strong>$1</strong>')
    .replace(/(^|[^*])\*([^*\n]+)\*(?!\*)/g,'$1<em>$2</em>')
    .replace(/`([^`]+)`/g,'<code>$1</code>')
    .replace(/\n/g,'<br>');
}
/* 富 Markdown 渲染（笔记/对话消息）：标题/加粗/斜体/列表/引用/代码块/链接/分隔线 */
function inlineMdNote(s){
  let out=esc(s);
  out=out.replace(/`([^`]+)`/g,'<code>$1</code>');
  out=out.replace(/\[([^\]]+)\]\((https?:[^)]+)\)/g,'<a href="$2" target="_blank" rel="noopener">$1</a>');
  out=out.replace(/\*\*([^*]+)\*\*/g,'<strong>$1</strong>');
  out=out.replace(/(^|[^*])\*([^*\n]+)\*(?!\*)/g,'$1<em>$2</em>');
  return out;
}
function mdRich(s){
  if(!s) return '';
  const parts=[];
  // 块级结构在“未转义”文本上识别，转义下沉到行内渲染（避免 > # - 被 &gt; 干扰）
  let out=s;
  out=out.replace(/```([\s\S]*?)```/g,(m,c)=>{ parts.push('<pre><code>'+esc(c)+'</code></pre>'); return '\u0000'+(parts.length-1)+'\u0000'; });
  const lines=out.split('\n');
  let html='', inUl=false, inOl=false, inBq=false;
  const closeAll=()=>{ if(inUl){html+='</ul>';inUl=false;} if(inOl){html+='</ol>';inOl=false;} if(inBq){html+='</blockquote>';inBq=false;} };
  for(const line of lines){
    const m=line.match(/^(#{1,4})\s+(.*)$/);
    if(m){ closeAll(); const lv=Math.min(6,m[1].length+2); html+='<h'+lv+'>'+inlineMdNote(m[2])+'</h'+lv+'>'; continue; }
    if(/^\s*---+\s*$/.test(line)){ closeAll(); html+='<hr>'; continue; }
    if(/^>\s?/.test(line)){ if(!inBq){html+='<blockquote>';inBq=true;} html+=inlineMdNote(line.replace(/^>\s?/,''))+'<br>'; continue; }
    if(/^\s*[-*+]\s+/.test(line)){ if(inOl){ closeAll(); } if(!inUl){html+='<ul>';inUl=true;} html+='<li>'+inlineMdNote(line.replace(/^\s*[-*+]\s+/,''))+'</li>'; continue; }
    if(/^\s*\d+\.\s+/.test(line)){ if(inUl){ closeAll(); } if(!inOl){html+='<ol>';inOl=true;} html+='<li>'+inlineMdNote(line.replace(/^\s*\d+\.\s+/,''))+'</li>'; continue; }
    closeAll();
    if(line.trim()) html+='<p>'+inlineMdNote(line)+'</p>';
  }
  closeAll();
  return html.replace(/\u0000(\d+)\u0000/g,(m,i)=>parts[+i]);
}
function toggleAiNote(m,btn,id){
  const n=notesOf(id);
  n.notes=n.notes||[];
  if(m.noteId && n.notes.some(x=>x.id===m.noteId)){
    // 已保存 → 删除该条笔记
    n.notes=n.notes.filter(x=>x.id!==m.noteId);
    m.noteId=null;
    saveNotes(id,n);
    const msgs=chatOf(id);
    const idx=msgs.findIndex(x=>x===m || (x.ts===m.ts && x.content===m.content));
    if(idx>=0){ msgs[idx].noteId=null; saveChat(id,msgs); }
    btn.classList.remove('saved');
    btn.textContent='📝 保存到笔记';
    btn.title='仅把这条 AI 答复保存到笔记';
    toast('已从笔记中删除');
  } else {
    // 未保存 → 仅保存该条 AI 答复（不含问题和其他答复）
    const noteId=uid();
    n.notes.push({id:noteId, quote:'', text:m.content, source:'ai', label:'🤖 AI 答复', ts:Date.now()});
    m.noteId=noteId;
    saveNotes(id,n);
    const msgs=chatOf(id);
    const idx=msgs.findIndex(x=>x===m || (x.ts===m.ts && x.content===m.content));
    if(idx>=0){ msgs[idx].noteId=noteId; saveChat(id,msgs); }
    btn.classList.add('saved');
    btn.textContent='✓ 已保存笔记';
    btn.title='点击从笔记中删除这条答复';
    toast('已保存到笔记');
  }
  renderNotesPanel();
}
function buildAiMessages(userText){
  const a=BYID[state.cur];
  const plain=plainOf(a).slice(0,14000);
  const n=notesOf(a.id);
  const notes=[];
  if(n.articleNote&&n.articleNote.text) notes.push('文章笔记：'+n.articleNote.text);
  (n.notes||[]).slice(-8).forEach(x=>notes.push('笔记：'+(x.quote?'「'+x.quote.slice(0,120)+'」 ':'')+x.text));
  const sys=[
    '你是「巴菲特投资智慧助教」，服务于一位 A 股量化投资者（使用选股因子、仓位管理、回测与择时信号体系）。',
    '任务：帮助用户把巴菲特致股东信中的思想，结合到其股票投资实践中。',
    '规则：',
    '1. 严格基于【当前文章】的原文回答；引用原文用「」括起，不得编造巴菲特没有表达过的观点；',
    '2. 每条建议给出「可操作步骤」或「如何验证」（如转化为选股条件、过滤规则、回测假设、仓位规则、风控阈值）；',
    '3. 若文章内容不足以支撑问题，明确说明，并给出合理推演框架；',
    '4. 注意巴菲特思想与 A 股市场（散户结构、涨跌停、T+1、政策市、行业轮动快）的差异并给出适配建议；',
    '5. 中文回答，分点清晰，篇幅适中。',
  ].join('\n');
  const ctx=[
    '【当前文章】标题：《'+a.title+'》｜分类：'+(CAT_NAME[a.catKey]||'')+(a.year?'｜年份：'+a.year:''),
    '【文章内容】',
    plain,
    notes.length?('【用户在这篇文章的笔记/高亮】\n'+notes.join('\n')):'',
  ].filter(Boolean).join('\n\n');
  const msgs=[
    {role:'system',content:sys},
    {role:'user',content:ctx},
  ];
  (chatOf(a.id)||[]).filter(m=>(m.role==='user'||(m.role==='assistant'&&!m.pending))).slice(-10)
    .forEach(m=>msgs.push({role:m.role,content:m.content.slice(0,3000)}));
  msgs.push({role:'user',content:userText});
  return msgs;
}
let aiBusy=false;
async function sendAi(){
  if(aiBusy) return;
  const box=$('#aiInputBox');
  const text=box.value.trim();
  if(!text){ toast('请输入问题'); return; }
  if(!settingsOk()){ openSettings(); return; }
  const id=state.cur;
  const requestMessages=buildAiMessages(text); // 此时历史中尚未加入当前问题，避免重复发送
  const msgs=chatOf(id);
  msgs.push({role:'user',content:text,quote:currentSelection()||'',ts:Date.now()});
  const aidMsg={role:'assistant',content:'',pending:true,ts:Date.now()};
  msgs.push(aidMsg);
  saveChat(id,msgs);
  box.value='';
  renderChatPanel();
  aiBusy=true;
  const sendBtn=$('#aiSend'); sendBtn.disabled=true;
  $('#aiStatus').textContent='思考中…（'+(settings().model.split('/').pop()||'')+'）';
  let saveTimer=null;
  try{
    const reply=await callLLMStream(requestMessages, acc=>{
      aidMsg.content=acc;
      if(state.cur===id){
        const el=$('#aiMsgs .ai-msg.assistant:last-of-type');
        if(el) el.innerHTML=mdRich(acc);
      }
      if(!saveTimer){ saveTimer=setTimeout(()=>{ saveChat(id,msgs); saveTimer=null; },500); }
    });
    aidMsg.content=reply; aidMsg.pending=false;
    clearTimeout(saveTimer);
    saveChat(id,msgs);
    if(state.cur===id) renderChatPanel();
  }catch(err){
    clearTimeout(saveTimer);
    const msgs2=chatOf(id).filter(m=>!(m.role==='assistant'&&m.pending));
    msgs2.push({role:'err',content:'请求失败：'+err.message,ts:Date.now()});
    saveChat(id,msgs2);
    if(state.cur===id) renderChatPanel();
  }
  aiBusy=false; sendBtn.disabled=false;
  $('#aiStatus').textContent='';
}
async function callLLM(messages){
  const s=settings();
  if(!s.key) throw new Error('未配置 API Key');
  let model=s.model||'';
  if((s.base||'').includes('deepseek')&&model.includes('/')) model=model.split('/').pop();
  const url=(s.base||'').replace(/\/+$/,'')+'/chat/completions';
  let resp;
  try{
    resp=await fetch(url,{
      method:'POST',
      headers:{'Content-Type':'application/json','Authorization':'Bearer '+s.key},
      body:JSON.stringify({model, messages, temperature:0.4}),
    });
  }catch(e){
    throw new Error('网络/跨域错误：'+e.message+'（如持续失败，可尝试把 Key 配置到支持 CORS 的代理，或用本地服务器打开本页）');
  }
  if(!resp.ok){
    let msg='HTTP '+resp.status;
    try{ const j=await resp.json(); msg+='：'+((j.error&&(j.error.message||j.error.code))||JSON.stringify(j).slice(0,200)); }catch(e){}
    throw new Error(msg);
  }
  const j=await resp.json();
  const c=j.choices&&j.choices[0]&&j.choices[0].message&&j.choices[0].message.content;
  if(!c) throw new Error('响应为空');
  return c;
}
/* 流式调用：SSE 解析，onDelta 持续回调累计文本 */
async function callLLMStream(messages, onDelta){
  const s=settings();
  if(!s.key) throw new Error('未配置 API Key');
  let model=s.model||'';
  if((s.base||'').includes('deepseek')&&model.includes('/')) model=model.split('/').pop();
  const url=(s.base||'').replace(/\/+$/,'')+'/chat/completions';
  let resp;
  try{
    resp=await fetch(url,{
      method:'POST',
      headers:{'Content-Type':'application/json','Authorization':'Bearer '+s.key},
      body:JSON.stringify({model, messages, temperature:0.4, stream:true}),
    });
  }catch(e){
    throw new Error('网络/跨域错误：'+e.message+'（如持续失败，可尝试把 Key 配置到支持 CORS 的代理，或用本地服务器打开本页）');
  }
  if(!resp.ok){
    let msg='HTTP '+resp.status;
    try{ const j=await resp.json(); msg+='：'+((j.error&&(j.error.message||j.error.code))||JSON.stringify(j).slice(0,200)); }catch(e){}
    throw new Error(msg);
  }
  const ct=(resp.headers.get('content-type')||'');
  if(!ct.includes('text/event-stream')){
    // 兼容非流式响应
    const j=await resp.json();
    const c=j.choices&&j.choices[0]&&j.choices[0].message&&j.choices[0].message.content;
    if(!c) throw new Error('响应为空');
    if(onDelta) onDelta(c);
    return c;
  }
  const reader=resp.body.getReader();
  const dec=new TextDecoder();
  let buf='', full='';
  const flush=line=>{
    if(!line.startsWith('data:')) return;
    const data=line.slice(5).trim();
    if(data==='[DONE]') return;
    try{
      const j=JSON.parse(data);
      const delta=j.choices&&j.choices[0]&&j.choices[0].delta&&j.choices[0].delta.content;
      if(delta){ full+=delta; if(onDelta) onDelta(full); }
    }catch(e){}
  };
  for(;;){
    const {done,value}=await reader.read();
    if(done) break;
    buf+=dec.decode(value,{stream:true});
    let idx;
    while((idx=buf.indexOf('\n'))>=0){
      flush(buf.slice(0,idx).trim());
      buf=buf.slice(idx+1);
    }
  }
  buf=buf.trim();
  if(buf) flush(buf);
  if(!full) throw new Error('响应为空');
  return full;
}

$('#aiSend').onclick=sendAi;
$('#aiInputBox').addEventListener('keydown',e=>{
  if(e.key==='Enter'&&!e.shiftKey){ e.preventDefault(); sendAi(); }
});

/* ================= 与巴菲特对话 / 与芒格对话（celebrity 双人格）================= */
let chatMode='buffett';
const chatViewEl = $('#chatView');
const CHAT_KEYS={buffett:'bf_buffett_chat', munger:'bf_munger_chat'};
const chatMsgs = (mode=chatMode) => store.get(CHAT_KEYS[mode], []);
const saveChatMsgs = (msgs,mode=chatMode) => { store.set(CHAT_KEYS[mode], msgs); pushRemoteState(); };
const CHAT_META={
  buffett:{name:'巴菲特',title:'与巴菲特对话',
    loaded:'🧬 巴菲特人格已加载（celebrity-buffett：6 个心智模型 / 8 条决策启发式 / Agentic Protocol 先研究再回答）。回答严格基于致股东信语料与人格框架，超出语料的推断会明确标注；发送「退出」结束对话。',
    missing:'⚠️ 未嵌入 celebrity-buffett 人格数据——请先运行 distilly 生成技能并重新构建本应用。',
    status:'巴菲特思考中…',placeholder:'用巴菲特的方式问我任何投资问题…（发送「退出」结束对话）',
    chips:['用巴菲特框架评估：我该集中持有还是分散？','如何看待当前 A 股的市场波动？','什么是护城河？怎么判断一家公司有没有？','如果我只能记住三条投资原则，应该是什么？','你犯过最大的错误是什么？学到了什么？']},
  munger:{name:'芒格',title:'与芒格对话',
    loaded:'🧬 芒格人格已加载（celebrity-charlie-munger：6 个心智模型 / 8 条决策启发式 / Agentic Protocol 先研究再回答）。回答严格基于芒格语料（演讲/Wesco 信/年会问答）与人格框架，超出语料的推断会明确标注；发送「退出」结束对话。',
    missing:'⚠️ 未嵌入 celebrity-charlie-munger 人格数据——请先运行 distilly 生成技能并重新构建本应用。',
    status:'芒格思考中…',placeholder:'用芒格的方式问我任何投资或人生问题…（发送「退出」结束对话）',
    chips:['怎么避免犯错？','多元思维模型怎么入门？','市场波动时应该怎么做？','你怎么看比特币和 AI？','如果我 30 岁，你会给我什么人生建议？']}
};
function showChatView(){
  readerEl.hidden = true; libEl.hidden = true; homeEl.hidden = true; idxEl.hidden = true;
  layoutEl.hidden = false; chatViewEl.hidden = false;
  document.title = CHAT_META[chatMode].title+' · '+DATA.title;
  renderChatView();
  window.scrollTo(0,0);
}
function renderChatView(){
  const meta=CHAT_META[chatMode];
  $('#chatStatus').textContent='';
  const tName=$('#chatTitleName'); if(tName) tName.textContent=meta.title;
  const tSub=$('#chatSub'); if(tSub) tSub.textContent='基于 '+ (chatMode==='buffett'?'celebrity-buffett':'celebrity-charlie-munger') +' 人格 · 仅供参考，不构成投资建议';
  const tIcon=$('#chatTitleIcon');
  if(tIcon){
    const b64 = chatMode==='munger'
      ? (typeof MUNGER_ICON_B64!=='undefined'?MUNGER_ICON_B64:'')
      : (typeof APP_ICON_B64!=='undefined'?APP_ICON_B64:'');
    if(b64){ tIcon.style.display=''; tIcon.src='data:image/png;base64,'+b64; }
    else tIcon.style.display='none';
  }
  const tBox=$('#chatInputBox'); if(tBox) tBox.placeholder=meta.placeholder;
  const intro=$('#chatIntro');
  const hasPersona=!!((chatMode==='buffett'?DATA.buffettPersona:DATA.mungerPersona)||'').trim();
  intro.innerHTML = hasPersona
    ? '<div class="chat-intro">'+meta.loaded+'</div>'
    : '<div class="chat-intro warn">'+meta.missing+'</div>';
  const box=$('#chatMsgs');
  box.innerHTML='';
  const msgs=chatMsgs();
  if(!settingsOk()){
    box.innerHTML='<div class="ai-welcome">🤖 「'+meta.title+'」需要配置大模型接口（DeepSeek 等 OpenAI 兼容接口）。<br><br><button class="btn primary" id="chatOpenSettings" style="font-size:13px">⚙ 打开设置</button></div>';
    const b=$('#chatOpenSettings'); if(b) b.onclick=openSettings;
    $('#chatChips').innerHTML=''; $('#chatInput').style.display='none'; $('#chatStatus').textContent='';
    return;
  }
  $('#chatInput').style.display='flex';
  msgs.forEach(m=>{
    const d=document.createElement('div');
    d.className='ai-msg '+m.role;
    d.innerHTML=mdRich(m.content);
    box.appendChild(d);
  });
  if(!msgs.length){
    $('#chatChips').innerHTML=meta.chips.map(t=>'<button class="ai-chip">'+esc(t)+'</button>').join('');
    $('#chatChips').onclick=e=>{ const c=e.target.closest('.ai-chip'); if(c){ $('#chatInputBox').value=c.textContent; sendChat(); } };
  } else $('#chatChips').innerHTML='';
  box.scrollTop=box.scrollHeight;
}
function buildChatMessages(userText,mode=chatMode){
  const isM=mode==='munger';
  const persona=((isM?DATA.mungerPersona:DATA.buffettPersona)||'').slice(0,14000);
  const sys=[
    isM
      ? '你正在扮演「查理·芒格（celebrity-charlie-munger）」——一个基于芒格演讲文集（1986–2017）、Wesco 致股东信（1997–2009）与 Wesco/DJCO 年会问答（1999–2023）蒸馏的决策与判断框架助手。'
      : '你正在扮演「巴菲特（celebrity-buffett）」——一个基于巴菲特致股东信知识库（1956–2025，81 封信件 + 35 概念 + 61 公司案例）蒸馏的投资决策与判断框架助手。',
    '【角色与人格（persona）】',
    persona,
    '【附加规则】',
    ...(isM ? [
      '1. 用芒格的口吻与框架回答：毒舌直白、三连词（"自大、愚蠢、疯狂"）、格言化收尾、自嘲式认错（"我和沃伦犯过很多错"）、政治不正确；',
      '2. 先按 Agentic Protocol 研究再回答：逆向思维（怎样会失败）→ 激励机制 → 能力圈 → 25 种误判清单 → 机会成本 → 复利时间视角；',
      '3. 涉及具体事实/数字/年份必须来自上述语料；语料没有的，明确说"语料中没有，这是我的框架推演"；绝不编造芒格说过的具体话；',
      '4. 被问 2024 年后的新事物（AI 大爆发、2026 年世界），先承认"我 2023 年就去世了，后面的世界没看到"，再给框架外推；',
      '5. 回答不构成投资建议；中文回答，结构清晰；用户说「退出」时简短告别。'
    ] : [
      '1. 用巴菲特的口吻与框架回答：生意比喻（护城河/裸泳/称重机）、"我们"体、先讲道理后给结论、自嘲式承认局限；',
      '2. 先按 Agentic Protocol 研究再回答：可预测性（能力圈）→ 护城河 → 管理层 → 价格安全边际 → 叙事触发器 → 错误清单反查；',
      '3. 涉及具体事实/数字/年份必须来自上述语料；语料没有的，明确说"语料中没有，这是我的框架推演"；绝不编造巴菲特说过的具体话；',
      '4. 回答不构成投资建议；涉及用户具体持仓时给出分析框架而非买卖指令；',
      '5. 中文回答，结构清晰；用户说「退出」时简短告别。'
    ]),
  ].join('\n');
  const msgs=[{role:'system',content:sys}];
  chatMsgs(mode).filter(m=>(m.role==='user'||(m.role==='assistant'&&!m.pending))).slice(-12)
    .forEach(m=>msgs.push({role:m.role,content:m.content.slice(0,3000)}));
  msgs.push({role:'user',content:userText});
  return msgs;
}
let chatBusy=false;
const chatRequestSeq={buffett:0,munger:0};
async function sendChat(){
  if(chatBusy) return;
  const reqMode=chatMode;
  const meta=CHAT_META[reqMode];
  const box=$('#chatInputBox');
  const text=box.value.trim();
  if(!text){ toast('请输入问题'); return; }
  if(!settingsOk()){ openSettings(); return; }
  const requestMessages=buildChatMessages(text,reqMode); // 只包含此前历史
  const msgs=chatMsgs(reqMode);
  msgs.push({role:'user',content:text,ts:Date.now()});
  const aidMsg={role:'assistant',content:'',pending:true,ts:Date.now()};
  msgs.push(aidMsg);
  saveChatMsgs(msgs,reqMode);
  box.value='';
  renderChatView();
  chatBusy=true;
  const reqId=++chatRequestSeq[reqMode];
  $('#chatSend').disabled=true;
  $('#chatStatus').textContent=meta.status+'（'+(settings().model.split('/').pop()||'')+'）';
  let saveTimer=null;
  try{
    const reply=await callLLMStream(requestMessages, acc=>{
      if(reqId!==chatRequestSeq[reqMode]) return;
      aidMsg.content=acc;
      if(chatMode===reqMode){
        const el=$('#chatMsgs .ai-msg.assistant:last-of-type');
        if(el) el.innerHTML=mdRich(acc);
      }
      if(!saveTimer){ saveTimer=setTimeout(()=>{
        if(reqId===chatRequestSeq[reqMode]) saveChatMsgs(msgs,reqMode);
        saveTimer=null;
      },500); }
    });
    clearTimeout(saveTimer);
    if(reqId===chatRequestSeq[reqMode]){
      aidMsg.content=reply; aidMsg.pending=false;
      saveChatMsgs(msgs,reqMode);
      if(chatMode===reqMode) renderChatView();
    }
  }catch(err){
    clearTimeout(saveTimer);
    if(reqId===chatRequestSeq[reqMode]){
      const msgs2=chatMsgs(reqMode).filter(m=>!(m.role==='assistant'&&m.pending));
      msgs2.push({role:'err',content:'请求失败：'+err.message,ts:Date.now()});
      saveChatMsgs(msgs2,reqMode);
      if(chatMode===reqMode) renderChatView();
    }
  }
  chatBusy=false;
  $('#chatSend').disabled=false;
  if(chatMode===reqMode) $('#chatStatus').textContent='';
}
$('#chatSend').onclick=sendChat;
$('#chatInputBox').addEventListener('keydown',e=>{ if(e.key==='Enter'&&!e.shiftKey){ e.preventDefault(); sendChat(); } });
$('#chatBack').onclick=()=>{ location.hash='#/'; };
$('#chatClear').onclick=()=>{
  chatRequestSeq[chatMode]++;               // 使当前人格正在返回的旧回调失效，避免清空后复活
  saveChatMsgs([],chatMode);
  renderChatView();
  toast('对话已清空');
};

/* ================= 背景解释 ================= */
const bgState = { term:'', ctx:'', msgs:[], busy:false, saved:false };
let bgRequestSeq=0;

function switchToBgTab(){
  $$('.rp-tab').forEach(x=>x.classList.remove('active'));
  const t=document.querySelector('.rp-tab[data-tab="bg"]');
  if(t) t.classList.add('active');
  state.tab='bg';
  $$('.tabpane').forEach(p=>p.classList.remove('active'));
  $('#tab-bg').classList.add('active');
}
function switchToNotesTab(){
  $$('.rp-tab').forEach(x=>x.classList.remove('active'));
  const t=document.querySelector('.rp-tab[data-tab="notes"]');
  if(t) t.classList.add('active');
  state.tab='notes';
  $$('.tabpane').forEach(p=>p.classList.remove('active'));
  $('#tab-notes').classList.add('active');
  renderNotesPanel();
}

function buildBgMessages(term, ctx, followUp){
  const a=BYID[state.cur];
  const plain=plainOf(a).slice(0,12000);
  const yi=(a.catKey==='partnership'||a.catKey==='berkshire')
    ? (IDX.year||[]).find(x=>x.y===a.year) : null;
  const yearBg = yi ? [
    '写作年份：'+yi.y,
    '市场/经济背景：'+(yi.bg||'（无详细记录）'),
    '当年重大事件：'+(yi.e||'（无详细记录）'),
    '当年核心主题摘要：'+(yi.s||'（无详细记录）'),
  ].join('\n') : '（该文章无明确年份或无年度背景记录）';
  const sys=[
    '你是一位金融史与投资概念专家，服务于阅读巴菲特致股东信的投资者。',
    '任务：解释用户选中的词语/概念。要求：',
    '1. 给出该词语的金融/投资学定义（简明准确，2-3句）；',
    '2. 结合【文章所在年份】的市场背景与历史事件，说明该概念在当时的含义、用法和市场语境；',
    '3. 结合文章内容，说明巴菲特在文中如何使用该概念、想表达什么；',
    '4. 如该概念在当代有新发展或理解差异，简要指出；',
    '5. 若词语是公司名、人名、政策名等专有名词，介绍其背景及其与文章的关系；',
    '6. 中文回答，用 Markdown 分点（**加粗**小标题），控制在400字以内，重点突出，不堆砌。',
  ].join('\n');
  const parts=[
    '【当前文章】《'+a.title+'》'+(a.year?'（'+a.year+'年）':''),
    '【年份背景】\n'+yearBg,
    '【文章内容】\n'+plain,
  ];
  if(ctx) parts.push('【选中词语所在段落】\n'+ctx);
  const base=parts.join('\n\n');
  const msgs=[{role:'system',content:sys}];
  if(followUp){
    msgs.push({role:'user',content:base+'\n\n请解释：「'+term+'」'});
    bgState.msgs.forEach(m=>{
      if(m.role==='user'||m.role==='assistant') msgs.push({role:m.role,content:m.content.slice(0,3000)});
    });
    msgs.push({role:'user',content:followUp});
  } else {
    msgs.push({role:'user',content:base+'\n\n请解释：「'+term+'」'});
  }
  return msgs;
}

function renderBgPanel(){
  const id=state.cur;
  if(!id) return;
  const a=BYID[id];
  const box=$('#bgMsgs');
  const savedBox=$('#bgSaved');
  $('#bgSend').disabled=bgState.busy;
  $('#bgExplainBtn').disabled=bgState.busy;
  // 已保存的背景解释（来自笔记中 source='bg' 的条目）
  const n=notesOf(id);
  const saved=(n.notes||[]).filter(x=>x.source==='bg').sort((x,y)=>y.ts-x.ts);
  if(saved.length){
    savedBox.hidden=false;
    const list=savedBox.querySelector('#bgSavedList');
    list.innerHTML=saved.map(b=>
      '<div class="bg-saved-item" data-bgid="'+b.id+'"><span class="bg-dot"></span>'+
      '<span class="bg-saved-label">'+esc(b.quote||b.label||'背景解释')+'</span>'+
      '<button class="bg-saved-del" data-delbg="'+b.id+'" title="删除">✕</button></div>'
    ).join('');
    list.onclick=e=>{
      const del=e.target.closest('[data-delbg]');
      if(del){
        e.stopPropagation();
        n.notes=n.notes.filter(x=>x.id!==del.dataset.delbg);
        saveNotes(id,n);
        // 同步清除 bgState 中对应消息的 noteId
        bgState.msgs.forEach(m=>{ if(m.noteId===del.dataset.delbg) m.noteId=null; });
        toast('已删除');
        renderBgPanel(); renderNotesPanel();
        return;
      }
      // 点击条目切换到笔记 tab
      const it=e.target.closest('[data-bgid]');
      if(it){ switchToNotesTab(); renderNotesPanel(); }
    };
  } else {
    savedBox.hidden=true;
  }
  // 消息区
  box.innerHTML='';
  if(!bgState.term&&!bgState.msgs.length){
    box.innerHTML='<div class="bg-welcome">🔍 <b>背景解释</b><br><br>'+
      '选中文章中的词语，点击浮动工具栏的「🔍 背景解释」，或在上方输入词语，AI 将结合文章内容与<b>写作年份的市场背景</b>解释其金融概念。<br><br>'+
      '你可以就解释继续追问，并把满意的解释保存到笔记，随时回顾。'+
      '<div class="bg-tip">💡 解释会结合该文写作年份'+(a.year?'（'+a.year+'年）':'')+'的市场环境、重大事件，以及巴菲特在文中的用法。追问同样基于这些上下文。</div></div>';
    $('#bgInput').style.display='none';
    $('#bgStatus').textContent='';
    return;
  }
  $('#bgInput').style.display='flex';
  // 词语条
  const termEl=document.createElement('div');
  termEl.className='bg-msg term';
  termEl.textContent='📌 '+bgState.term;
  box.appendChild(termEl);
  // 消息（每条 AI 答复独立保存/删除）
  bgState.msgs.forEach(m=>{
    const d=document.createElement('div');
    d.className='bg-msg '+m.role;
    if(m.role==='assistant'){
      d.innerHTML=mdLight(m.content);
      const curNotes=notesOf(id);
      const stillSaved=m.noteId && (curNotes.notes||[]).some(x=>x.id===m.noteId);
      const btn=document.createElement('button');
      btn.className='bg-note-btn'+(stillSaved?' saved':'');
      btn.textContent=stillSaved?'✓ 已保存笔记':'📝 保存到笔记';
      btn.title=stillSaved?'点击从笔记中删除这条解释':'仅把这条解释保存到笔记';
      btn.onclick=()=>toggleBgNote(m,btn,id);
      d.appendChild(btn);
    } else {
      d.textContent=m.content;
    }
    box.appendChild(d);
  });
  box.scrollTop=box.scrollHeight;
}

async function explainBgTerm(term, ctx){
  term=(term||'').trim();
  if(!term){ toast('请输入要解释的词语'); return; }
  if(!settingsOk()){ openSettings(); return; }
  const reqArticle=state.cur;
  const reqId=++bgRequestSeq;
  bgState.term=term;
  bgState.ctx=ctx||'';
  bgState.msgs=[];
  bgState.busy=true;
  bgState.saved=false;
  $('#bgTermInput').value=term;
  renderBgPanel();
  $('#bgStatus').textContent='正在检索背景资料并解释…（'+(settings().model.split('/').pop()||'')+'）';
  try{
    const reply=await callLLM(buildBgMessages(term,ctx));
    if(state.cur!==reqArticle||reqId!==bgRequestSeq) return;
    bgState.msgs.push({role:'assistant',content:reply,ts:Date.now()});
  }catch(err){
    if(state.cur!==reqArticle||reqId!==bgRequestSeq) return;
    bgState.msgs.push({role:'err',content:'解释失败：'+err.message,ts:Date.now()});
  }
  bgState.busy=false;
  $('#bgStatus').textContent='';
  renderBgPanel();
}

async function sendBgFollowUp(){
  if(bgState.busy) return;
  const input=$('#bgInputBox');
  const text=input.value.trim();
  if(!text) return;
  if(!settingsOk()){ openSettings(); return; }
  const reqArticle=state.cur;
  const reqId=++bgRequestSeq;
  const requestMessages=buildBgMessages(bgState.term,bgState.ctx,text); // 此时历史中尚未加入当前追问
  bgState.msgs.push({role:'user',content:text,ts:Date.now()});
  input.value='';
  bgState.busy=true;
  bgState.saved=false;
  renderBgPanel();
  $('#bgStatus').textContent='思考中…';
  try{
    const reply=await callLLM(requestMessages);
    if(state.cur!==reqArticle||reqId!==bgRequestSeq) return;
    bgState.msgs.push({role:'assistant',content:reply,ts:Date.now()});
  }catch(err){
    if(state.cur!==reqArticle||reqId!==bgRequestSeq) return;
    bgState.msgs.push({role:'err',content:'追问失败：'+err.message,ts:Date.now()});
  }
  bgState.busy=false;
  $('#bgStatus').textContent='';
  renderBgPanel();
}

function toggleBgNote(m,btn,id){
  const n=notesOf(id);
  n.notes=n.notes||[];
  if(m.noteId && n.notes.some(x=>x.id===m.noteId)){
    // 已保存 → 删除该条笔记
    n.notes=n.notes.filter(x=>x.id!==m.noteId);
    m.noteId=null;
    saveNotes(id,n);
    btn.classList.remove('saved');
    btn.textContent='📝 保存到笔记';
    btn.title='仅把这条解释保存到笔记';
    toast('已从笔记中删除');
  } else {
    // 未保存 → 仅保存该条解释（不含追问和其他回答）
    const noteId=uid();
    n.notes.push({id:noteId, quote:bgState.term||'背景解释', text:m.content, source:'bg', label:'🔍 背景解释', ts:Date.now()});
    m.noteId=noteId;
    saveNotes(id,n);
    btn.classList.add('saved');
    btn.textContent='✓ 已保存笔记';
    btn.title='点击从笔记中删除这条解释';
    toast('已保存到笔记');
  }
  renderBgPanel();
  renderNotesPanel();
}

$('#bgExplainBtn').onclick=()=>explainBgTerm($('#bgTermInput').value);
$('#bgTermInput').addEventListener('keydown',e=>{
  if(e.key==='Enter'){ e.preventDefault(); explainBgTerm($('#bgTermInput').value); }
});
$('#bgSend').onclick=sendBgFollowUp;
$('#bgInputBox').addEventListener('keydown',e=>{
  if(e.key==='Enter'&&!e.shiftKey){ e.preventDefault(); sendBgFollowUp(); }
});

/* ================= 设置 ================= */
function openSettings(){
  const s=settings();
  const local=store.get(SETTINGS_KEY,{});
  const cfg=runtimeLlmConfig||window.BUFFETT_LLM_CONFIG||null;
  $('#setBase').value=s.base;
  // 服务端注入的 Key 只留在内存；输入框只回显用户主动保存的浏览器 Key。
  $('#setKey').value=local.key||'';
  $('#setModel').value=s.model;
  $('#setHint').innerHTML = runtimeLlmConfig&&runtimeLlmConfig.key
    ? '已从本地服务器的同源 /api/llm-config 加载密钥（只在内存，不会回填或保存到 localStorage）。如需覆盖，可在此手动填写。'
    : (cfg
        ? '已加载无密钥默认配置。用本地服务器启动可从环境变量 DEEPSEEK_API_KEY 注入，或在此手动填写。'
        : '未检测到默认配置。请填写 API Key（仅保存在本浏览器 localStorage）。');
  $('#setTestResult').textContent='';
  $('#settingsModal').hidden=false;
}
function closeSettings(){ $('#settingsModal').hidden=true; }
$('#settingsBtn').onclick=openSettings;
$('#setCancel').onclick=closeSettings;
$('#setSave').onclick=()=>{
  const next={base:$('#setBase').value.trim(),model:$('#setModel').value.trim()};
  const key=$('#setKey').value.trim();
  if(key) next.key=key;
  store.set(SETTINGS_KEY,next);
  closeSettings(); toast('设置已保存');
  if(!chatViewEl.hidden) renderChatView();
  else if(state.cur) renderChatPanel();
};
$('#setTest').onclick=async ()=>{
  const btn=$('#setTest');
  btn.disabled=true; btn.textContent='测试中…';
  $('#setTestResult').textContent='';
  try{
    const base=$('#setBase').value.trim();
    const key=$('#setKey').value.trim()||settings().key;
    const model=$('#setModel').value.trim();
    if(!key) throw new Error('未配置 API Key');
    let m=model; if(base.includes('deepseek')&&m.includes('/')) m=m.split('/').pop();
    const resp=await fetch(base.replace(/\/+$/,'')+'/chat/completions',{
      method:'POST',
      headers:{'Content-Type':'application/json','Authorization':'Bearer '+key},
      body:JSON.stringify({model:m,messages:[{role:'user',content:'你好'}],max_tokens:8}),
    });
    if(!resp.ok) throw new Error('HTTP '+resp.status);
    $('#setTestResult').textContent='✅ 连接成功，配置可用';
    $('#setTestResult').style.color='#15803d';
  }catch(e){
    $('#setTestResult').textContent='❌ '+e.message;
    $('#setTestResult').style.color='#b91c1c';
  }
  btn.disabled=false; btn.textContent='测试连接';
};

/* ================= 全局事件 ================= */
$('#q').addEventListener('input',e=>{
  state.q=e.target.value;
  clearTimeout(window.__qTimer);
  window.__qTimer=setTimeout(()=>{
    doSearch(state.q);
    $('#qCount').textContent=state.q&&searchIdx?Object.keys(searchIdx).length+' 条':'';
    renderAll();
  },120);
});
$('#sortSel').onchange=e=>{ state.sort=e.target.value; renderAll(); };
$('#groupSel').onchange=e=>{ state.group=e.target.value; renderAll(); };
$('#viewSel').onchange=e=>{ state.view=e.target.value; renderAll(); };
$('#sbFav').onchange=e=>{ state.favOnly=e.target.checked; renderAll(); };
$('#sbNoted').onchange=e=>{ state.notedOnly=e.target.checked; renderAll(); };
$('#menuBtn').onclick=()=>{ $('#sidebar').classList.toggle('open'); $('#sidebarBackdrop').classList.toggle('show'); };
$('#sidebarBackdrop').onclick=()=>{ $('#sidebar').classList.remove('open'); $('#sidebarBackdrop').classList.remove('show'); };

/* 左侧栏隐藏/显示（桌面端）：点击一次隐藏，再点一次显示，正文随之伸缩 */
const SB_KEY='bf_sidebar_show';
const sbToggleEl=$('#sbToggle');
const applySidebar = show => {
  $('#sidebar').classList.toggle('collapsed', !show);
  if(sbToggleEl) sbToggleEl.textContent = show ? '◧' : '◨';
};
if(sbToggleEl){
  applySidebar(store.get(SB_KEY,1)!==0);
  sbToggleEl.onclick=()=>{
    const nowCollapsed = $('#sidebar').classList.toggle('collapsed');
    const show = !nowCollapsed;
    store.set(SB_KEY, show?1:0);
    applySidebar(show);
    window.dispatchEvent(new Event('resize'));   // hero 图等随正文宽度自适应
  };
}
$('#notesExport').onclick=exportNotes;
$('#brandBtn').onclick=()=>{ if(location.hash!=='#/') location.hash='#/'; else showHome(); };
document.addEventListener('keydown',e=>{
  if(e.key==='/'&&document.activeElement!==$('#q')&&!e.target.closest('input,textarea')){ e.preventDefault(); $('#q').focus(); }
});
window.addEventListener('scroll',()=>{
  const h=document.documentElement;
  const p=h.scrollTop/(h.scrollHeight-h.clientHeight||1);
  $('#progress').style.width=(p*100).toFixed(1)+'%';
},{passive:true});
// 阅读视图内点击内部链接
document.addEventListener('click',e=>{
  const lk=e.target.closest('a[data-nav]');
  if(lk){ e.preventDefault(); location.hash=lk.getAttribute('href'); }
});
// 标签页切换
$$('.rp-tab').forEach(t=>{
  t.onclick=()=>{
    $$('.rp-tab').forEach(x=>x.classList.remove('active'));
    t.classList.add('active');
    state.tab=t.dataset.tab;
    $$('.tabpane').forEach(p=>p.classList.remove('active'));
    const pane=$('#tab-'+state.tab);
    pane.classList.add('active');
    if(state.tab==='notes') renderNotesPanel();
    if(state.tab==='ai') renderChatPanel();
    if(state.tab==='bg') renderBgPanel();
  };
});

/* ================= 编辑风索引视图 ================= */
const idxEl = $('#idxView');
const layoutEl = $('#layout');
const IV = {tab:'letters', series:'all', event:'all', rendered:{}};

const IV_EV_CLS = {
  '危机/股灾':'ev-crisis','泡沫/狂热':'ev-bubble','通胀/加息':'ev-inflation',
  '战争/恐袭':'ev-war','疫情':'ev-pandemic','正常':'ev-normal'
};
const IV_PERIOD_CLS = {
  '危机/股灾':'iv-period-crisis','泡沫/狂热':'iv-period-bubble','通胀/加息':'iv-period-inflation',
  '战争/恐袭':'iv-period-war','疫情':'iv-period-pandemic','正常':'iv-period-normal'
};
const IV_SERIES_CLS = {'合伙基金信':'partnership','伯克希尔信':'berkshire'};

function ivEsc(s){return String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');}

function ivArticleForYear(y, seriesName){
  const ck = seriesName==='合伙基金信' ? 'partnership' : 'berkshire';
  return ART.find(a=>a.year===y && a.catKey===ck);
}

function ivPrimaryArticle(yr){
  // Prefer berkshire, then partnership
  return ivArticleForYear(yr.y,'伯克希尔信') || ivArticleForYear(yr.y,'合伙基金信');
}

function ivCleanSummary(s){
  // Strip 【series】 prefixes from merged-year summaries
  return String(s||'').replace(/【[^】]+】/g,'').trim();
}

function ivEventTags(yr){
  return String(yr.e||'').split(/[、,，;；]/).map(x=>x.trim()).filter(Boolean);
}

function ivAllEventTags(){
  const set = new Set();
  (IDX.year||[]).forEach(yr=>ivEventTags(yr).forEach(t=>set.add(t)));
  return Array.from(set);
}

function showIndexView(){
  if(state.cur) state.cur=null;
  readerEl.hidden=true; libEl.hidden=true; homeEl.hidden=true; chatViewEl.hidden=true;
  layoutEl.hidden=true; idxEl.hidden=false;
  document.title='巴菲特致股东信分类索引 · 编辑风';
  if(!IV.rendered.stats) renderIvStats();
  if(!IV.rendered[IV.tab]) renderIvTab(IV.tab);
  renderIvFilterBar();
  window.scrollTo(0,0);
}

function renderIvStats(){
  const nLetters = ART.filter(a=>a.catKey==='partnership'||a.catKey==='berkshire').length;
  const stats = [
    [nLetters, 'Letters'],
    [(IDX.topic||[]).length, 'Themes'],
    [(IDX.industry||[]).length, 'Industries'],
    [(IDX.event||[]).length, 'Eras'],
    [(IDX.method||[]).length, 'Methods'],
  ];
  $('#ivStats').innerHTML = stats.map(s=>
    '<div class="iv-stat"><div class="iv-stat-num">'+s[0]+'</div><div class="iv-stat-label">'+s[1]+'</div></div>'
  ).join('');
  $('#ivFooter').innerHTML =
    '数据来源：巴菲特致合伙人及股东信（1956–2025）· 坎宁安《巴菲特致股东的信》主题分类<br>'+
    '本页为应用内编辑风索引视图，与主应用共享全部数据与功能 · <a href="#/" style="color:#8b6f47">返回主视图</a>';
  IV.rendered.stats=true;
}

function renderIvFilterBar(){
  const bar = $('#ivFilterBar');
  if(IV.tab!=='letters'){bar.innerHTML=''; bar.style.display='none'; return;}
  bar.style.display='flex';
  const series = [['all','全部'],['合伙基金信','合伙基金信'],['伯克希尔信','伯克希尔信']];
  const evTags = ivAllEventTags();
  bar.innerHTML =
    '<span class="iv-filter-label">信件系列</span>'+
    '<div class="iv-filter-group">'+
      series.map(s=>'<button class="iv-filter-btn'+(IV.series===s[0]?' active':'')+'" data-sfilter="'+s[0]+'">'+s[1]+'</button>').join('')+
    '</div>'+
    '<div class="iv-filter-div"></div>'+
    '<span class="iv-filter-label">事件类型</span>'+
    '<div class="iv-filter-group">'+
      '<button class="iv-filter-btn'+(IV.event==='all'?' active':'')+'" data-efilter="all">全部</button>'+
      evTags.map(t=>'<button class="iv-filter-btn'+(IV.event===t?' active':'')+'" data-efilter="'+ivEsc(t)+'">'+ivEsc(t)+'</button>').join('')+
    '</div>';
  bar.querySelectorAll('[data-sfilter]').forEach(b=>b.onclick=()=>{IV.series=b.dataset.sfilter; renderIvFilterBar(); renderIvLetters();});
  bar.querySelectorAll('[data-efilter]').forEach(b=>b.onclick=()=>{IV.event=b.dataset.efilter; renderIvFilterBar(); renderIvLetters();});
}

function renderIvTab(tab){
  IV.tab=tab;
  document.querySelectorAll('#idxView .iv-tab').forEach(t=>t.classList.toggle('active',t.dataset.ivtab===tab));
  ['letters','themes','industries','events','methods'].forEach(t=>{
    const el=$('#iv'+t.charAt(0).toUpperCase()+t.slice(1));
    if(el) el.classList.toggle('active',t===tab);
  });
  renderIvFilterBar();
  if(tab==='letters') renderIvLetters();
  else if(tab==='themes') renderIvThemes();
  else if(tab==='industries') renderIvIndustries();
  else if(tab==='events') renderIvEvents();
  else if(tab==='methods') renderIvMethods();
  IV.rendered[tab]=true;
}

function renderIvLetters(){
  const years = (IDX.year||[]).slice().sort((a,b)=>a.y-b.y);
  const filtered = years.filter(yr=>{
    if(IV.series!=='all'){
      const has = (yr.series||[]).some(s=>s===IV.series);
      if(!has) return false;
    }
    if(IV.event!=='all'){
      const tags = ivEventTags(yr);
      if(!tags.includes(IV.event)) return false;
    }
    return true;
  });
  const container = $('#ivLetters');
  if(!filtered.length){
    container.innerHTML='<div class="iv-empty">没有符合筛选条件的信件</div>';
    return;
  }
  let html = '<div class="iv-section"><div class="iv-sec-head"><span class="iv-sec-title">年度信件</span>'+
    '<span class="iv-sec-meta">'+filtered.length+' 封 · 1956–2025</span></div>'+
    '<div class="iv-sec-desc">按年份排列的巴菲特致合伙人及股东信，点击标题可在应用内阅读全文。</div></div>';
  filtered.forEach(yr=>{
    const art = ivPrimaryArticle(yr);
    const tags = ivEventTags(yr);
    const seriesArr = yr.series||[];
    const badges = seriesArr.map(s=>{
      const cls = IV_SERIES_CLS[s]||'';
      return '<span class="iv-series-badge '+cls+'">'+ivEsc(s)+'</span>';
    }).join('');
    const topicTags = (yr.t||[]).slice(0,3).map(t=>'<span class="iv-tag theme">'+ivEsc(t)+'</span>').join('');
    const evTags = tags.map(t=>{
      const cls = IV_EV_CLS[t]||'';
      return '<span class="iv-tag '+cls+'">'+ivEsc(t)+'</span>';
    }).join('');
    const title = art ? '<a href="#/a/'+encodeURIComponent(art.id)+'">'+yr.y+' 年'+(seriesArr.includes('合伙基金信')?'致合伙人信':'致股东信')+'</a>' : yr.y+' 年信';
    const extLink = (yr.link||'').split('；').filter(Boolean)[0];
    const readLink = art ? '<a href="#/a/'+encodeURIComponent(art.id)+'" class="iv-letter-link">阅读全文 →</a>' :
      (extLink ? '<a href="'+ivEsc(extLink)+'" target="_blank" rel="noopener" class="iv-letter-link">阅读原文 →</a>' : '');
    const ctx = yr.bg ? '<div class="iv-letter-ctx">'+ivEsc(yr.bg)+'</div>' : '';
    html += '<div class="iv-section" style="padding-top:0;padding-bottom:0">'+
      '<div class="iv-letter">'+
        '<div class="iv-letter-year">'+yr.y+'</div>'+
        '<div class="iv-letter-body">'+
          '<div class="iv-letter-meta">'+badges+
            '<span class="iv-letter-author">'+ivEsc(yr.a||'')+'</span>'+
          '</div>'+
          '<div class="iv-letter-title">'+title+'</div>'+
          '<div class="iv-letter-summary">'+ivEsc(ivCleanSummary(yr.s))+'</div>'+
          '<div class="iv-letter-tags">'+topicTags+evTags+readLink+'</div>'+
          ctx+
        '</div>'+
      '</div></div>';
  });
  container.innerHTML=html;
}

function renderIvThemes(){
  const themes = IDX.topic||[];
  let html = '<div class="iv-section"><div class="iv-sec-head"><span class="iv-sec-title">主题分类</span>'+
    '<span class="iv-sec-meta">'+themes.length+' 个主题</span></div>'+
    '<div class="iv-sec-desc">基于坎宁安《巴菲特致股东的信》的主题分类法，将数十年信件按核心投资主题归类。</div></div>';
  themes.forEach((th,i)=>{
    const years = (th.y||[]).slice(0,8);
    const concepts = (th.con||[]).slice(0,6);
    html += '<div class="iv-section" style="padding-top:0;padding-bottom:0"><div class="iv-card">'+
      '<div class="iv-card-head">'+
        '<span class="iv-card-num">'+String(i+1).padStart(2,'0')+'</span>'+
        '<span class="iv-card-title">'+ivEsc(th.n||th.c||'')+'</span>'+
      '</div>'+
      '<div class="iv-card-body">'+ivEsc(th.d||'')+'</div>'+
      (th.rep?'<div class="iv-card-label">代表信件</div><div class="iv-card-text">'+ivEsc(th.rep)+'</div>':'')+
      (concepts.length?'<div class="iv-card-label">关键概念</div><div class="iv-card-tags">'+
        concepts.map(c=>'<span class="iv-tag theme">'+ivEsc(c)+'</span>').join('')+'</div>':'')+
      (years.length?'<div class="iv-card-label">重点年份</div><div class="iv-card-tags">'+
        years.map(y=>'<span class="iv-year-tag">'+y+'</span>').join('')+'</div>':'')+
    '</div></div>';
  });
  $('#ivThemes').innerHTML=html;
}

function renderIvIndustries(){
  const inds = IDX.industry||[];
  let html = '<div class="iv-section"><div class="iv-sec-head"><span class="iv-sec-title">行业分类</span>'+
    '<span class="iv-sec-meta">'+inds.length+' 个行业</span></div>'+
    '<div class="iv-sec-desc">巴菲特投资过的核心行业与代表性公司，按行业归类相关信件。</div></div>';
  inds.forEach((ind,i)=>{
    const years = (ind.y||[]).slice(0,10);
    html += '<div class="iv-section" style="padding-top:0;padding-bottom:0"><div class="iv-card">'+
      '<div class="iv-card-head">'+
        '<span class="iv-card-num">'+String(i+1).padStart(2,'0')+'</span>'+
        '<span class="iv-card-title">'+ivEsc(ind.n||'')+'</span>'+
        (ind.co?'<span class="iv-card-sub">· '+ivEsc(ind.co)+'</span>':'')+
      '</div>'+
      '<div class="iv-card-body">'+ivEsc(ind.d||'')+'</div>'+
      (years.length?'<div class="iv-card-label">重点年份</div><div class="iv-card-tags">'+
        years.map(y=>'<span class="iv-year-tag">'+y+'</span>').join('')+'</div>':'')+
    '</div></div>';
  });
  $('#ivIndustries').innerHTML=html;
}

function renderIvEvents(){
  const events = IDX.event||[];
  let html = '<div class="iv-section"><div class="iv-sec-head"><span class="iv-sec-title">事件时期</span>'+
    '<span class="iv-sec-meta">'+events.length+' 个时期</span></div>'+
    '<div class="iv-sec-desc">巴菲特投资生涯中经历的重大市场事件与时期，及其观点与教训。</div></div>';
  events.forEach((ev,i)=>{
    const years = (ev.y||[]);
    // Determine period class from first event tag
    const firstTag = ivEventTags({e:ev.n}).concat(ivEventTags({e:ev.bg})).find(t=>IV_PERIOD_CLS[t]) ||
                     (ev.n&&IV_PERIOD_CLS[Object.keys(IV_PERIOD_CLS).find(k=>ev.n.includes(k))]) || '';
    const periodCls = IV_PERIOD_CLS[firstTag] || '';
    html += '<div class="iv-section" style="padding-top:0;padding-bottom:0"><div class="iv-card">'+
      '<div class="iv-card-head">'+
        '<span class="iv-card-num">'+String(i+1).padStart(2,'0')+'</span>'+
        '<span class="iv-card-title">'+ivEsc(ev.n||'')+'</span>'+
      '</div>'+
      (ev.rng?'<span class="iv-event-period '+(periodCls||'')+'">'+ivEsc(ev.rng)+'</span>':'')+
      (ev.bg?'<div class="iv-card-label">市场背景</div><div class="iv-card-text">'+ivEsc(ev.bg)+'</div>':'')+
      (ev.act?'<div class="iv-card-label">观点与行动</div><div class="iv-card-text">'+ivEsc(ev.act)+'</div>':'')+
      (ev.les?'<div class="iv-card-label">经验教训</div><div class="iv-card-text">'+ivEsc(ev.les)+'</div>':'')+
      (years.length?'<div class="iv-card-label">相关信件年份</div><div class="iv-card-tags">'+
        years.map(y=>'<span class="iv-year-tag">'+y+'</span>').join('')+'</div>':'')+
    '</div></div>';
  });
  $('#ivEvents').innerHTML=html;
}

function renderIvMethods(){
  const methods = IDX.method||[];
  let html = '<div class="iv-section"><div class="iv-sec-head"><span class="iv-sec-title">选股方法演进</span>'+
    '<span class="iv-sec-meta">'+methods.length+' 个阶段</span></div>'+
    '<div class="iv-sec-desc">巴菲特从"烟蒂投资"到"护城河"的选股方法演变历程。</div></div>';
  methods.forEach((m,i)=>{
    const years = (m.y||[]);
    html += '<div class="iv-section" style="padding-top:0;padding-bottom:0"><div class="iv-card">'+
      '<div class="iv-card-head">'+
        '<span class="iv-card-num">'+String(i+1).padStart(2,'0')+'</span>'+
        '<span class="iv-card-title">'+ivEsc((m.n||'').replace(/\n/g,' '))+'</span>'+
      '</div>'+
      (m.m?'<div class="iv-method-method">'+ivEsc((m.m||'').replace(/\n/g,' '))+'</div>':'')+
      (m.view?'<div class="iv-card-label">核心观点</div><div class="iv-card-text">'+ivEsc(m.view)+'</div>':'')+
      (m.cases?'<div class="iv-card-label">代表案例</div><div class="iv-card-text">'+ivEsc(m.cases)+'</div>':'')+
      (m.shift?'<div class="iv-card-label">关键转变</div><div class="iv-card-text">'+ivEsc(m.shift)+'</div>':'')+
      (years.length?'<div class="iv-card-label">代表年份</div><div class="iv-card-tags">'+
        years.map(y=>'<span class="iv-year-tag">'+y+'</span>').join('')+'</div>':'')+
    '</div></div>';
  });
  $('#ivMethods').innerHTML=html;
}

// Tab click delegation
document.addEventListener('click', function(e){
  const tab = e.target.closest('#idxView .iv-tab');
  if(tab){ renderIvTab(tab.dataset.ivtab); }
});
// Index button
(function(){
  const btn = document.getElementById('indexBtn');
  if(btn) btn.onclick = function(){ location.hash = '#/index'; };
})();

/* ================= 路由 / 启动 ================= */
function showHome(){
  if (state.cur) state.cur=null;
  readerEl.hidden = true; libEl.hidden = true; homeEl.hidden = false; idxEl.hidden = true; chatViewEl.hidden = true;
  layoutEl.hidden = false;
  document.title = DATA.title+' · 致股东信知识库';
  renderHome();
  window.scrollTo(0,0);
}
function renderAllNoHash(){
  renderSidebar(); renderLibrary();
}
function renderAll(){
  if (location.hash !== '#/library') location.hash = '#/library';
  renderAllNoHash();
}
function onHash(){
  flushArticleNoteSave();
  const m=location.hash.match(/^#\/a\/(.+)$/);
  const id=m?decodeURIComponent(m[1]):null;
  if(id&&BYID[id]) openArticle(id);
  else if(location.hash==='#/library') renderAllNoHash();
  else if(location.hash==='#/index') showIndexView();
  else if(location.hash==='#/chat') showChatView();
  else showHome();
}
async function flushState(){
  flushArticleNoteSave();
  return pushRemoteStateNow(false);
}
window.addEventListener('hashchange',onHash);
window.addEventListener('pagehide',()=>{
  flushArticleNoteSave();
  pushRemoteStateNow(true);                 // 绕过 300ms debounce，给关闭窗口一次即时落盘机会
});
let toastTimer=null;
function toast(msg){
  const t=$('#toast');
  t.textContent=msg;
  t.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer=setTimeout(()=>t.classList.remove('show'),2200);
}
renderSidebar();
$('#sortSel').value = state.sort;   // 默认排序与状态保持一致
Promise.all([loadRemoteState(),loadRuntimeLlmConfig()]).then(onHash); // 先恢复记忆与同源运行时配置，再渲染

/* 测试钩子 */
window.BUF={state,openArticle,doSearch,visibleList,mdToHtml,toPlain,plainOf,applyHighlights,
  addHighlight,notesOf,saveNotes,settings,callLLM,exportNotes,store,BYID,ART,
  bgState,renderBgPanel,explainBgTerm,sendBgFollowUp,toggleBgNote,toggleAiNote,buildBgMessages,
  switchToBgTab,switchToNotesTab,selectionContext,renderChatPanel,renderNotesPanel,flushState};
"""

HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<meta name="description" content="巴菲特致股东信知识库 · 分类/搜索/阅读/笔记/划线/LLM 讨论 单文件应用">
<title>巴菲特投资智慧 · 致股东信知识库</title>
<style>
__CSS__
</style>
</head>
<body>
<header id="topbar">
  <button class="tbtn" id="menuBtn" title="筛选面板">☰</button>
  <button class="tbtn" id="sbToggle" title="隐藏/显示左侧栏">◧</button>
  <div class="brand" id="brandBtn" title="回到主页"><span class="logo">__BRAND_ICON__巴菲特<em>投资智慧</em></span><span class="sub">致股东信知识库</span></div>
  <div class="search-wrap">
    <input id="q" type="search" placeholder="搜索文章、概念、公司、人物…（/ 聚焦）" autocomplete="off">
    <span id="qCount"></span>
  </div>
  <div class="top-actions">
    <select class="tbtn" id="sortSel" title="排序">
      <option value="year-asc">时间 ↑</option>
      <option value="year-desc">时间 ↓</option>
      <option value="title">标题</option>
      <option value="cat">分类</option>
      <option value="len">篇幅</option>
      <option value="fav">收藏优先</option>
    </select>
    <select class="tbtn" id="groupSel" title="分组">
      <option value="none">不分组</option>
      <option value="cat">按分类分组</option>
      <option value="decade">按年代分组</option>
      <option value="tag">按主题分组</option>
    </select>
    <select class="tbtn" id="viewSel" title="视图">
      <option value="grid">卡片</option>
      <option value="rows">列表</option>
    </select>
    <button class="tbtn" id="indexBtn" title="编辑风索引纵览">📖 索引</button>
    <button class="tbtn" id="notesExport" title="导出全部笔记">📥 导出笔记</button>
    <button class="tbtn" id="settingsBtn" title="LLM 设置">⚙ 设置</button>
  </div>
</header>
<div id="sidebarBackdrop"></div>
<div id="layout">
  <aside id="sidebar">
    <div class="sb-sec">
      <div class="sb-title">分类索引 <span class="sb-title-range" id="idxRange"></span></div>
      <div id="sbIdx"></div>
    </div>
    <div class="sb-sec">
      <div class="sb-title">分类</div>
      <div id="sbCats"></div>
    </div>
    <div class="sb-sec">
      <div class="sb-title">年代</div>
      <div id="sbDecades"></div>
    </div>
    <div class="sb-sec">
      <div class="sb-title">主题标签</div>
      <div class="tag-cloud" id="sbTags"></div>
    </div>
    <div class="sb-sec">
      <div class="sb-title">状态</div>
      <label class="sb-check"><input type="checkbox" id="sbFav"> ★ 仅看收藏</label>
      <label class="sb-check"><input type="checkbox" id="sbNoted"> 📝 有笔记/高亮</label>
    </div>
    <div class="sb-foot" id="sbFootInfo"></div>
  </aside>
  <main id="main">
    <div id="home" hidden></div>
    <div id="chatView" hidden>
      <div class="chat-head">
        <button class="tbtn" id="chatBack">← 返回</button>
        <div class="chat-title">__CHAT_ICON__<span id="chatTitleName">与巴菲特对话</span>
          <span class="chat-sub" id="chatSub">基于 celebrity-buffett 人格（致股东信知识库蒸馏）· 仅供参考，不构成投资建议</span>
        </div>
        <button class="tbtn" id="chatClear" title="清空本对话记录">清空对话</button>
      </div>
      <div id="chatIntro"></div>
      <div class="chat-body">
        <div class="chat-msgs" id="chatMsgs"></div>
        <div class="ai-chips" id="chatChips"></div>
        <div class="ai-status" id="chatStatus"></div>
        <div class="ai-input" id="chatInput">
          <textarea id="chatInputBox" rows="2" placeholder="用巴菲特的方式问我任何投资问题…（发送「退出」结束对话）"></textarea>
          <button id="chatSend">发送</button>
        </div>
      </div>
    </div>
    <div id="library" hidden>
      <div id="libHead">
        <span id="libTitle">全部文章</span>
        <span id="libCount"></span>
      </div>
      <div id="idxBanner"></div>
      <div id="libList"></div>
    </div>
    <div id="reader" hidden>
      <div id="progress"></div>
      <div class="reader-top">
        <button class="tbtn" id="backBtn">← 返回列表</button>
        <div class="reader-title">
          <h1 id="rTitle"></h1>
          <div class="r-meta" id="rMeta"></div>
          <div class="r-tags" id="rTags"></div>
          <div class="r-idx" id="rIdxLine"></div>
        </div>
        <div class="r-actions">
          <button class="tbtn" id="rRead" title="标记已读 / 再点取消">📖 未读</button>
          <button class="tbtn" id="rFav" title="收藏">☆</button>
          <button class="tbtn" id="rPrev">← 上一篇</button>
          <button class="tbtn" id="rNext">下一篇 →</button>
        </div>
      </div>
      <div class="reader-body">
        <article id="article"></article>
        <aside id="rpanel">
          <div class="rp-tabs">
            <button class="rp-tab active" data-tab="toc">目录</button>
            <button class="rp-tab" data-tab="notes">笔记<span class="cnt" id="notesCnt"></span></button>
            <button class="rp-tab" data-tab="ai">AI 讨论</button>
            <button class="rp-tab" data-tab="bg">背景解释</button>
          </div>
          <div class="tabpane active" id="tab-toc"></div>
          <div class="tabpane" id="tab-notes">
            <div class="note-sec"><h4>划线高亮</h4><div id="ntHighlights"></div></div>
            <div class="note-sec"><h4>笔记</h4><div id="ntNotes"></div></div>
            <div class="note-sec"><h4>背景解释</h4><div id="ntBg"></div></div>
            <div class="note-sec"><h4>文章笔记（自动保存）<button class="note-clear-btn" id="clearArticleNote" title="清空文章笔记">清空</button></h4>
              <textarea id="articleNote" placeholder="记录你对这篇文章的理解、与 A 股实践的关联…"></textarea>
              <div class="note-saved" id="noteSaved"></div>
            </div>
          </div>
          <div class="tabpane" id="tab-ai">
            <div class="ai-msgs" id="aiMsgs"></div>
            <div class="ai-chips" id="aiChips"></div>
            <div class="ai-status" id="aiStatus"></div>
            <div class="ai-input" id="aiInput" style="display:none">
              <textarea id="aiInputBox" rows="1" placeholder="问 AI：这段思想如何应用到我的股票投资？"></textarea>
              <button id="aiSend">发送</button>
            </div>
          </div>
          <div class="tabpane" id="tab-bg">
            <div class="bg-head">
              <div class="bg-term-row">
                <label>词语</label>
                <input id="bgTermInput" type="text" placeholder="输入或选中词语后点「解释」">
                <button id="bgExplainBtn">解释</button>
              </div>
            </div>
            <div class="bg-msgs" id="bgMsgs"></div>
            <div class="bg-saved" id="bgSaved" hidden>
              <h4>📚 本文已保存的背景解释</h4>
              <div id="bgSavedList"></div>
            </div>
            <div class="bg-status" id="bgStatus"></div>
            <div class="bg-input" id="bgInput" style="display:none">
              <textarea id="bgInputBox" rows="1" placeholder="就这个解释继续追问…（Enter 发送）"></textarea>
              <button id="bgSend">发送</button>
            </div>
          </div>
        </aside>
      </div>
      <div class="reader-foot">
        <span class="tbtn" id="rNavInfo" style="border:none;background:none"></span>
        <span style="font-size:12px;color:#9a917f" id="rFootInfo">选中文字可高亮 / 下划线 / 背景解释 / 记笔记</span>
      </div>
    </div>
  </main>
</div>

<div id="selToolbar" hidden>
  <button id="hlYellow" title="黄色高亮">🟡 高亮</button>
  <button id="hlBlue" title="蓝色高亮">🔵 高亮</button>
  <button id="hlUnderline" title="绿色下划线">🟢 划线</button>
  <span class="sep"></span>
  <button id="hlBg" title="用 AI 解释选中词语的金融概念与时代背景">🔍 背景解释</button>
  <span class="sep"></span>
  <button id="hlNote" title="基于选中文字写笔记">📝 笔记</button>
  <button id="hlCopy" title="复制选中文字">⧉ 复制</button>
</div>

<div id="modalBackdrop" hidden>
  <div class="modal">
    <h3 id="nmTitle">新建笔记</h3>
    <div class="quote-box" id="nmQuote"></div>
    <div class="field">
      <textarea id="nmText" rows="5" style="width:100%;border:1px solid var(--line);border-radius:9px;padding:9px 11px;font-size:14px;outline:none;background:var(--bg);resize:vertical"
        placeholder="写下你的想法…（会随文章保存，可导出）"></textarea>
    </div>
    <div class="modal-actions">
      <button class="btn ghost" id="nmCancel">取消</button>
      <button class="btn primary" id="nmSave">保存笔记</button>
    </div>
  </div>
</div>

<div id="settingsModal" hidden>
  <div class="modal">
    <h3>⚙ LLM 讨论设置</h3>
    <div class="field">
      <label>API Base（OpenAI 兼容）</label>
      <input type="text" id="setBase" placeholder="https://api.deepseek.com/v1">
    </div>
    <div class="field">
      <label>API Key</label>
      <input type="password" id="setKey" placeholder="sk-...">
    </div>
    <div class="field">
      <label>模型</label>
      <input type="text" id="setModel" placeholder="deepseek-v4-flash">
    </div>
    <div class="hint" id="setHint"></div>
    <div class="hint" style="color:#15803d" id="setTestResult"></div>
    <div class="modal-actions">
      <button class="btn ghost" id="setCancel">取消</button>
      <button class="btn ghost" id="setTest">测试连接</button>
      <button class="btn primary" id="setSave">保存</button>
    </div>
  </div>
</div>

<div id="idxView" hidden>
  <header class="iv-hero">
    <div class="iv-hero-inner">
      <div class="iv-eyebrow">Warren Buffett · 1956–2025</div>
      <h1>巴菲特致股东信<em>分类索引</em></h1>
      <p class="iv-hero-sub">按年度、主题、行业、事件时期、选股方法五个维度，系统梳理巴菲特投资思想的演进脉络。</p>
      <div class="iv-stats" id="ivStats"></div>
    </div>
  </header>
  <nav class="iv-nav">
    <div class="iv-nav-inner">
      <span class="iv-nav-brand">索引</span>
      <button class="iv-tab active" data-ivtab="letters">📅 年度信件</button>
      <button class="iv-tab" data-ivtab="themes">🏷 主题分类</button>
      <button class="iv-tab" data-ivtab="industries">🏭 行业分类</button>
      <button class="iv-tab" data-ivtab="events">🌪 事件时期</button>
      <button class="iv-tab" data-ivtab="methods">🧭 选股方法</button>
    </div>
  </nav>
  <div class="iv-filter-bar" id="ivFilterBar"></div>
  <div id="ivLetters" class="iv-pane active"></div>
  <div id="ivThemes" class="iv-pane"></div>
  <div id="ivIndustries" class="iv-pane"></div>
  <div id="ivEvents" class="iv-pane"></div>
  <div id="ivMethods" class="iv-pane"></div>
  <footer class="iv-footer" id="ivFooter"></footer>
</div>

<div id="toast"></div>

<script>
window.BUFFETT_DATA = __DATA__;
</script>
<script src="llm-config.js" onerror="window.__noLlmCfg=1"></script>
<script>
__ECHARTS__
</script>
<script>
window.BUFFETT_RETURNS = __RETURNS__;
</script>
<script>
__JS__
</script>
</body>
</html>
"""


# ---------------------------------------------------------------- 组装

def load_echarts():
    """读取 vendor/echarts.min.js（缺失时返回空串 → 主页图表隐藏）。"""
    p = os.path.join(HERE, "vendor", "echarts.min.js")
    try:
        with open(p, encoding="utf-8") as f:
            return f.read()
    except OSError:
        return ""


def load_returns():
    """读取 buffet_return_data.csv（巴菲特历年收益率，OCR 提取并经复算校验）。

    返回 JSON 数组 [[年份, 年收益率%, 累计净值, 年化%], ...]；缺失时返回 "null"。
    """
    p = os.path.join(HERE, "buffet_return_data.csv")
    try:
        import csv as _csv
        rows = []
        with open(p, encoding="utf-8-sig") as f:
            for row in _csv.DictReader(f):
                rows.append([int(row["年份"]), float(row["年收益率%"]),
                             float(row["模拟累计净值"]), float(row["年化收益%"])])
        return json.dumps(rows, ensure_ascii=False)
    except (OSError, KeyError, ValueError):
        return "null"


def build_html(data_json: str, css: str, js: str, icon_b64: str = "",
               echarts_js: str = "", returns_js: str = "null") -> str:
    brand_icon = ('<img class="brand-img" src="data:image/png;base64,' + icon_b64
                  + '" alt="">' if icon_b64 else "🏛 ")
    chat_icon = ('<img class="chat-img" id="chatTitleIcon" src="data:image/png;base64,' + icon_b64
                 + '" alt="">' if icon_b64 else "🗣 ")
    # 内嵌第三方 JS 必须转义 </script（压缩代码里可能出现）
    echarts_js = echarts_js.replace("</", "<\\/")
    return (
        HTML_TEMPLATE
        .replace("__DATA__", data_json)
        .replace("__CSS__", css)
        .replace("__JS__", js)
        .replace("__BRAND_ICON__", brand_icon)
        .replace("__CHAT_ICON__", chat_icon)
        .replace("__ECHARTS__", echarts_js)
        .replace("__RETURNS__", returns_js)
    )


def find_root_env():
    """读取构建脚本同目录的项目 .env（解析为 dict，不含则返回 {}）。"""
    p = os.path.join(HERE, ".env")
    if not os.path.isfile(p):
        return {}
    env = {}
    try:
        with open(p, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith("#") and "=" in line:
                    k, v = line.split("=", 1)
                    env[k.strip()] = v.strip().strip('"').strip("'")
    except OSError:
        return {}
    return env


def build_llm_config() -> bool:
    """生成不含密钥的 llm-config.js（仅默认 base/model）。

    服务端密钥不会写入应用状态或生成物：serve_buffett_app.py 从环境变量
    DEEPSEEK_API_KEY（或项目根 .env）读取，并通过同源 /api/llm-config
    仅注入页面内存；在应用「设置」面板手动填写的值则保存在浏览器
    localStorage。
    """
    env = find_root_env()
    base = env.get("DEEPSEEK_API_BASE", "https://api.deepseek.com/v1")
    model = "deepseek-v4-flash"   # 默认模型（可在设置面板覆盖）
    js = (
        "/* 可选 LLM 默认配置（由 build_buffett_app.py 生成，不含任何密钥）。\n"
        " * 服务端密钥不写入应用状态或生成物，仅通过同源 /api/llm-config 注入页面内存；\n"
        " * 也可在应用「设置」面板手动填写（仅保存在本浏览器 localStorage）。\n"
        " * 修改本文件后刷新页面即可生效。 */\n"
        "window.BUFFETT_LLM_CONFIG = {\n"
        '  base: %s,\n'
        '  key: "",\n'
        '  model: %s\n'
        "};\n"
    ) % (json.dumps(base), json.dumps(model))
    with open(OUT_LLM, "w", encoding="utf-8") as f:
        f.write(js)
    print(f"[ok] 已生成 {os.path.basename(OUT_LLM)}（无密钥，仅默认 base/model）")
    return True


def main():
    debug = "--debug" in sys.argv
    no_llm = "--no-llm-config" in sys.argv

    data, tag_count = build()

    # 数据 JSON：转义 </script
    data_json = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    data_json = data_json.replace("</", "<\\/")

    # 主页标志：注入 buffett.png / munger.png 的 base64（常量前置到 APP_JS）
    icon_b64 = load_icon_b64()
    munger_b64 = load_icon_b64("munger-128.png")
    js = APP_JS
    if icon_b64:
        js = "const APP_ICON_B64='" + icon_b64 + "';\n" + js
    if munger_b64:
        js = "const MUNGER_ICON_B64='" + munger_b64 + "';\n" + js

    html = build_html(data_json, APP_CSS, js, icon_b64,
                      echarts_js=load_echarts(), returns_js=load_returns())
    with open(OUT_HTML, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"[ok] 已生成 {os.path.basename(OUT_HTML)}（{len(html)/1024/1024:.2f} MB，{len(data['articles'])} 篇文章）")
    if debug:
        print("  分类: " + ", ".join(f"{c['name']}({c['count']})" for c in data['cats']))
        print(f"  年份范围: {data['yearRange']}")
        print("  主题标签数: %d，Top10: %s" % (len(tag_count), sorted(tag_count.items(), key=lambda x: -x[1])[:10]))

    if not no_llm:
        build_llm_config()


if __name__ == "__main__":
    main()
