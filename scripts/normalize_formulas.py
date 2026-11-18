# -*- coding: utf-8 -*-
"""把 LaTeX 公式转成「Markdown 里能渲染 + 人也看得懂」的样式

问题（用户实测反馈）：
    31593.8 那份切出来的公式长这样——
        \\frac{\\sqrt{2gH_{v}}}{A_{vent}}
    或者被塞进代码块里。丢进 Obsidian / Cherry Studio / 语料库平台，
    **要么原样显示一堆反斜杠，要么根本不渲染**——人看不懂，模型也读不利索。

本脚本做两件事：
    1. 数学区（$…$ / $$…$$ / \\(…\\) / \\[…\\] / \\begin{equation}…）→
       统一成 **可被 MathJax/KaTeX 渲染** 的 `$…$` / `$$…$$`，
       并把 `\\_` 这类「Markdown 转义残留」清掉（数学区里的 `_` 是下标运算符）。
    2. 每条**独立公式**后面补一行 `> 纯文本：…`，
       把 LaTeX 翻成 ρ、√、×、≤、上标 ²、下标 的下标写法 ——
       平台不渲染数学时，人和模型还能读。

三种风格（--style）：
    both   （默认）$…$ 渲染式 + 纯文本对照，两边都有
    math   只保留 MathJax 式（平台支持数学公式时）
    plain  全部转纯文本，不留 LaTeX（平台不渲染数学、且要给人读时）

用法：
    python normalize_formulas.py <目录或文件> --check        # 体检：有多少裸 LaTeX
    python normalize_formulas.py <目录> --in-place            # 就地修（默认 both）
    python normalize_formulas.py <目录> --out 新目录 --style plain
"""
import argparse
import re
import sys
from pathlib import Path

for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8")
    except Exception:
        pass

GREEK = {
    "alpha": "α", "beta": "β", "gamma": "γ", "delta": "δ", "epsilon": "ε",
    "varepsilon": "ε", "zeta": "ζ", "eta": "η", "theta": "θ", "vartheta": "ϑ",
    "iota": "ι", "kappa": "κ", "lambda": "λ", "mu": "μ", "nu": "ν", "xi": "ξ",
    "pi": "π", "rho": "ρ", "sigma": "σ", "tau": "τ", "upsilon": "υ", "phi": "φ",
    "varphi": "φ", "chi": "χ", "psi": "ψ", "omega": "ω",
    "Gamma": "Γ", "Delta": "Δ", "Theta": "Θ", "Lambda": "Λ", "Xi": "Ξ",
    "Pi": "Π", "Sigma": "Σ", "Upsilon": "Υ", "Phi": "Φ", "Psi": "Ψ", "Omega": "Ω",
}

SYMS = {
    r"\times": "×", r"\cdot": "·", r"\pm": "±", r"\mp": "∓", r"\div": "÷",
    r"\le": "≤", r"\leq": "≤", r"\ge": "≥", r"\geq": "≥", r"\ne": "≠",
    r"\neq": "≠", r"\approx": "≈", r"\equiv": "≡", r"\sim": "~",
    r"\infty": "∞", r"\partial": "∂", r"\nabla": "∇", r"\sum": "Σ",
    r"\prod": "Π", r"\int": "∫", r"\propto": "∝", r"\degree": "°",
    r"\rightarrow": "→", r"\to": "→", r"\leftarrow": "←",
    r"\Rightarrow": "⇒", r"\Leftrightarrow": "⇔", r"\ldots": "…",
    r"\dots": "…", r"\cdots": "…", r"\prime": "′", r"\mid": "|",
    r"\quad": " ", r"\qquad": "  ", r"\left": "", r"\right": "",
    r"\displaystyle": "", r"\limits": "", r"\!": "", r"\,:": " ",
    r"\,": " ", r"\;": " ", r"\:": " ", r"\%": "%", r"\&": "&",
    r"\#": "#", r"\$": "$", r"\{": "{", r"\}": "}",
}

SUP = {"0": "⁰", "1": "¹", "2": "²", "3": "³", "4": "⁴", "5": "⁵",
       "6": "⁶", "7": "⁷", "8": "⁸", "9": "⁹", "+": "⁺", "-": "⁻",
       "n": "ⁿ", "i": "ⁱ"}


def _match_brace(s, i):
    """s[i] 应为 '{'，返回与之匹配的 '}' 下标（找不到返回 -1）"""
    if i >= len(s) or s[i] != "{":
        return -1
    depth = 0
    for j in range(i, len(s)):
        if s[j] == "{":
            depth += 1
        elif s[j] == "}":
            depth -= 1
            if depth == 0:
                return j
    return -1


def _take_arg(s, i):
    """取一个参数：{...} 或单个字符。返回 (内容, 结束位置+1)"""
    while i < len(s) and s[i] == " ":
        i += 1
    if i < len(s) and s[i] == "{":
        j = _match_brace(s, i)
        if j > 0:
            return s[i + 1:j], j + 1
    if i < len(s):
        return s[i], i + 1
    return "", i


def _sup(txt):
    """上标：能映射就映射，否则 ^(txt)"""
    if all(c in SUP for c in txt):
        return "".join(SUP[c] for c in txt)
    return "^" + ("(" + txt + ")" if len(txt) > 1 else txt)


def _sub(txt):
    return "_" + ("(" + txt + ")" if len(txt) > 1 else txt)


def latex_to_plain(s):
    """LaTeX 片段 → 纯文本（尽量保留数学含义，不追求零误差）"""
    if not s:
        return ""
    # 0) 重音：\dot{m}→ṁ、\bar{x}→x̄、\hat{x}→x̂、\vec{x}→x⃗、\overline{..}
    COMB = {"dot": "\u0307", "bar": "\u0304", "hat": "\u0302",
            "vec": "\u20d7", "tilde": "\u0303"}
    for name, comb in COMB.items():
        pat = re.compile(r"\\" + name + r"\s*")
        while True:
            m = pat.search(s)
            if not m:
                break
            body, p = _take_arg(s, m.end())
            txt = latex_to_plain(body)
            if name == "dot" and txt == "m":
                txt = "ṁ"
            elif name == "overline":
                txt = txt
            else:
                txt = txt + comb if txt else txt
            s = s[:m.start()] + txt + s[p:]
    # 1) \frac{a}{b} / \dfrac
    while True:
        m = re.search(r"\\(?:d|t)?frac\s*", s)
        if not m:
            break
        num, p = _take_arg(s, m.end())
        den, p2 = _take_arg(s, p)
        s = s[:m.start()] + "(" + latex_to_plain(num) + ")/(" + latex_to_plain(den) + ")" + s[p2:]
    # 2) \sqrt[n]{x}
    while True:
        m = re.search(r"\\sqrt\s*", s)
        if not m:
            break
        i = m.end()
        idx = ""
        if i < len(s) and s[i] == "[":
            j = s.find("]", i)
            if j > 0:
                idx = latex_to_plain(s[i + 1:j])
                i = j + 1
        body, p = _take_arg(s, i)
        piece = ("[" + idx + "]" if idx else "") + "√(" + latex_to_plain(body) + ")"
        s = s[:m.start()] + piece + s[p:]
    # 3) \text{..} / \mathrm{..} / \mbox
    s = re.sub(r"\\(?:text|mathrm|mbox|operatorname)\s*", "", s)
    # 4) 希腊字母
    def gr(m):
        return GREEK.get(m.group(1), m.group(0))
    s = re.sub(r"\\([A-Za-z]+)", gr, s)
    # 5) 符号（长键优先）
    for k in sorted(SYMS, key=len, reverse=True):
        s = s.replace(k, SYMS[k])
    # 6) 上标 ^{..} / ^x
    s = re.sub(r"\^\s*\{([^{}]*)\}", lambda m: _sup(m.group(1)), s)
    s = re.sub(r"\^(\w)", lambda m: _sup(m.group(1)), s)
    # 7) 下标 _{..} / _x（\rho_{s,i} → ρ_(s,i)）
    s = re.sub(r"_\s*\{([^{}]*)\}", lambda m: _sub(m.group(1)), s)
    s = re.sub(r"_(\w)", lambda m: "_" + m.group(1), s)
    # 8) 收尾：括号、换行、多余空白
    s = s.replace("\\\\", " ； ").replace("\\", "")
    s = s.replace("{", "").replace("}", "")
    s = re.sub(r"[ \t]+", " ", s)
    return s.strip()


# ---------------------------------------------------------------- 识别
# (正则, 正文所在的 group 号)
DISPLAY_PATTERNS = [
    (re.compile(r"\$\$(.+?)\$\$", re.S), 1),
    (re.compile(r"\\\[(.+?)\\\]", re.S), 1),
    (re.compile(r"\\begin\{(equation|align|gather|eqnarray)\*?\}(.+?)"
                r"\\end\{\1\*?\}", re.S), 2),
]
INLINE_PATTERNS = [
    (re.compile(r"(?<!\\)\$(?!\$)([^$\n]+?)(?<!\\)\$"), 1),
    (re.compile(r"\\\((.+?)\\\)", re.S), 1),
]

# 裸 LaTeX 特征：出现这些命令但没有数学定界符
BARE_CMD = re.compile(
    r"\\(?:frac|dfrac|sqrt|rho|Delta|sigma|alpha|beta|gamma|lambda|mu|"
    r"epsilon|theta|omega|phi|tau|eta|nu|xi|pi|zeta|chi|psi|iota|kappa|"
    r"times|cdot|pm|le|leq|ge|geq|ne|neq|approx|infty|partial|nabla|sum|"
    r"int|propto|degree|rightarrow|leftarrow|Rightarrow|mathrm|text|"
    r"begin|end|left|right|displaystyle|overline|hat|bar|dot|vec)")
CODE_FENCE = re.compile(r"```[\s\S]*?```")
MATH_SPAN = re.compile(r"\$\$.+?\$\$|\$[^$\n]+?\$|\\\(.+?\\\)|\\\[.+?\\\]", re.S)
# 上一次运行留下的「纯文本」对照行（重跑时要先删掉，否则会重复叠加）
STALE_PLAIN = re.compile(r"(\$\$\n.+?\n\$\$\n)\n> 纯文本：[^\n]*\n", re.S)


def _has_bare(line):
    """去掉已定界的数学区后，还残留 LaTeX 命令 = 裸 LaTeX"""
    rest = MATH_SPAN.sub("", line)
    return bool(BARE_CMD.search(rest))


def _wrap_bare(text, style):
    """把没有定界符的 LaTeX 命令包成数学区（math/both），或直接转纯文本（plain）。"""
    out_lines = []
    for ln in text.split("\n"):
        if not _has_bare(ln):
            out_lines.append(ln)
            continue
        if style == "plain":
            out_lines.append(latex_to_plain(ln))
            continue
        # 行内：按中文/标点切段，含裸 LaTeX 的非中文段包进 $…$
        segs = re.split(r"([\u4e00-\u9fff，。；：、（）「」？！]{1,})", ln)
        rebuilt = []
        for seg in segs:
            if seg and not re.fullmatch(r"[\u4e00-\u9fff，。；：、（）「」？！]+", seg) \
                    and _has_bare(seg):
                keep_lead = seg[:len(seg) - len(seg.lstrip())]
                keep_tail = seg[len(seg.rstrip()):]
                core = seg.strip()
                if core:
                    rebuilt.append(keep_lead + "$" + core + "$" + keep_tail)
                else:
                    rebuilt.append(seg)
            else:
                rebuilt.append(seg)
        out_lines.append("".join(rebuilt))
    return "\n".join(out_lines)


def normalize_text(text, style="both"):
    """返回 (新文本, 统计dict)。幂等：重跑不会重复叠加「纯文本」行。

    处理顺序很关键（实测踩过）：
        先把**独立公式**整块换成占位符，再把行内公式换成占位符，
        最后才处理「裸 LaTeX」。
        否则 `$$\\n<公式>\\n$$` 里的公式体是单独一行，
        会被「裸 LaTeX」那一步再包一层 `$…$`，变成 `$\\n$…$\\n$`。
    """
    st = dict(display=0, inline=0, bare=0, unescaped=0)
    slots = {}

    def unescape_math(body):
        if "\\_" in body:
            st["unescaped"] += body.count("\\_")
            body = body.replace("\\_", "_")
        return body

    def stash(payload):
        k = f"\x00M{len(slots)}\x00"
        slots[k] = payload
        return k

    # 保护代码块（代码块里的内容不动）
    fences = []

    def keep(m):
        fences.append(m.group(0))
        return f"\x00FENCE{len(fences) - 1}\x00"
    text = CODE_FENCE.sub(keep, text)

    # 0) 清掉上一次生成的「纯文本」对照行（幂等：后面会按当前内容重新生成）
    text = STALE_PLAIN.sub(r"\1", text)

    # 1) 独立公式 → 占位符
    for pat, gi in DISPLAY_PATTERNS:
        def repl(m, gi=gi):
            st["display"] += 1
            body = unescape_math(m.group(gi)).strip()
            if style == "plain":
                return stash(latex_to_plain(body))
            block = "$$\n" + body + "\n$$"
            if style == "both":
                plain = latex_to_plain(body)
                if plain:
                    block += "\n\n> 纯文本：" + plain
            return stash(block)
        text = pat.sub(repl, text)

    # 2) 行内公式 → 占位符
    for pat, gi in INLINE_PATTERNS:
        def repl2(m, gi=gi):
            st["inline"] += 1
            body = unescape_math(m.group(gi)).strip()
            if style == "plain":
                return stash(latex_to_plain(body))
            return stash("$" + body + "$")
        text = pat.sub(repl2, text)

    # 3) 裸 LaTeX（此时所有公式都已是占位符，不会误伤）
    before = sum(1 for ln in text.split("\n") if _has_bare(ln))
    if before:
        st["bare"] = before
        text = _wrap_bare(text, style)

    # 4) 还原
    for k, v in slots.items():
        text = text.replace(k, v)
    for i, f in enumerate(fences):
        text = text.replace(f"\x00FENCE{i}\x00", f)
    return text, st


def _blank_math(text):
    """把数学区替换成等行数的空行（保留行号，供逐行体检用）"""
    def rep(m):
        return "\x00" + "\n" * m.group(0).count("\n") + "\x00"
    return MATH_SPAN.sub(rep, text)


BARE_STRUCT = re.compile(r"\^\{|_\{|\\\(")


# ---------------------------------------------------------------- 体检
def check_text(text):
    """返回 [(行号, 问题)]"""
    probs = []
    clean = CODE_FENCE.sub(lambda m: "\n" * m.group(0).count("\n"), text)

    # 数学区先抹掉，再逐行找残留 LaTeX —— 否则 `$$\n公式体\n$$` 的公式体会被误报
    blanked = _blank_math(clean)
    for i, ln in enumerate(blanked.split("\n"), 1):
        if _has_bare(ln):
            probs.append((i, "裸 LaTeX（无 $…$ 定界，平台不渲染）"))
        elif BARE_STRUCT.search(ln):
            probs.append((i, "疑似公式未加定界（出现 ^{ 或 _{）"))

    # 数学区内部不应出现 Markdown 转义 \_
    for m in MATH_SPAN.finditer(clean):
        if "\\_" in m.group(0):
            probs.append((clean[:m.start()].count("\n") + 1,
                          "数学区内有 Markdown 转义 \\_（应为 _）"))

    # 公式被塞进代码块
    for m in CODE_FENCE.finditer(text):
        seg = m.group(0)
        if BARE_CMD.search(seg) or "$$" in seg:
            probs.append((text[:m.start()].count("\n") + 1,
                          "公式被包进代码块（不渲染）"))
    # 去重（同一行多个问题只报一次主问题）+ 按行号排序
    seen = set()
    uniq = []
    for ln, msg in probs:
        if ln in seen:
            continue
        seen.add(ln)
        uniq.append((ln, msg))
    return sorted(uniq, key=lambda x: x[0])


def main():
    ap = argparse.ArgumentParser(description="LaTeX 公式 → 可渲染 Markdown + 纯文本对照")
    ap.add_argument("path", help=".md 文件或目录")
    ap.add_argument("--style", default="both", choices=["both", "math", "plain"],
                    help="both=渲染式+纯文本（默认）；math=只留 MathJax；plain=全转纯文本")
    ap.add_argument("--in-place", action="store_true")
    ap.add_argument("--out", default="", help="另存目录")
    ap.add_argument("--check", action="store_true", help="只体检；有问题退出码 1")
    ap.add_argument("--quiet", action="store_true")
    a = ap.parse_args()

    p = Path(a.path)
    if not p.exists():
        sys.exit(f"找不到 {p}")
    files = [p] if p.is_file() else sorted(p.rglob("*.md"))
    if not files:
        sys.exit(f"{p} 下没有 .md")

    outdir = Path(a.out) if a.out else None
    if outdir:
        outdir.mkdir(parents=True, exist_ok=True)

    tot = dict(display=0, inline=0, bare=0, unescaped=0, changed=0, files=0)
    bad = []

    for f in files:
        raw = f.read_text(encoding="utf-8")
        if a.check:
            probs = check_text(raw)
            if probs:
                bad.append((f, probs))
                if not a.quiet:
                    print(f"\n❌ {f.name}（{len(probs)} 处）")
                    for ln, msg in probs[:6]:
                        print(f"   L{ln}: {msg}")
            continue
        new, st = normalize_text(raw, a.style)
        tot["files"] += 1
        for k in ("display", "inline", "bare", "unescaped"):
            tot[k] += st[k]
        if new != raw:
            tot["changed"] += 1
            if a.in_place:
                f.write_text(new, encoding="utf-8")
            elif outdir:
                (outdir / f.name).write_text(new, encoding="utf-8")
            if not a.quiet and not outdir:
                print(f"  {f.name}: 独立公式{st['display']} 行内{st['inline']} "
                      f"裸LaTeX{st['bare']} 去转义{st['unescaped']}")

    if a.check:
        if bad:
            print(f"\n共 {len(bad)} 个文件有问题（"
                  f"{sum(len(x[1]) for x in bad)} 处）")
            print("修法：python normalize_formulas.py <该目录> --in-place")
            return 1
        print(f"✅ {len(files)} 个 md 公式格式全部合规")
        return 0

    print(f"\n处理 {tot['files']} 个文件，改动 {tot['changed']} 个（style={a.style}）")
    print(f"  独立公式 {tot['display']}｜行内 {tot['inline']}｜裸 LaTeX {tot['bare']}｜"
          f"去转义 {tot['unescaped']}")
    if not a.in_place and not outdir:
        print("  要写盘：加 --in-place 或 --out <目录>")
    return 0


if __name__ == "__main__":
    sys.exit(main())
