#!/usr/bin/env python3
"""Pure-Python LaTeX Math to Office Math Markup Language (OMML) Converter (`latex_to_omml.py`).

Provides zero-Pandoc native Word equation rendering (`<m:oMath>` & `<m:oMathPara>`) so that
inline `$...$` and display `$$...$$` formulas render as genuine editable Word equations
(`Cambria Math` + `宋体` + `#000000`) on machines where Pandoc is not installed.
"""

from __future__ import annotations

from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
W = f"{{{W_NS}}}"
M = f"{{{M_NS}}}"

LATEX_SYMBOLS: dict[str, str] = {
    # Lowercase Greek
    "alpha": "α",
    "beta": "β",
    "gamma": "γ",
    "delta": "δ",
    "epsilon": "ϵ",
    "varepsilon": "ε",
    "zeta": "ζ",
    "eta": "η",
    "theta": "θ",
    "vartheta": "ϑ",
    "iota": "ι",
    "kappa": "κ",
    "lambda": "λ",
    "mu": "μ",
    "nu": "ν",
    "xi": "ξ",
    "pi": "π",
    "varpi": "ϖ",
    "rho": "ρ",
    "varrho": "ϱ",
    "sigma": "σ",
    "varsigma": "ς",
    "tau": "τ",
    "upsilon": "υ",
    "phi": "ϕ",
    "varphi": "φ",
    "chi": "χ",
    "psi": "ψ",
    "omega": "ω",
    # Uppercase Greek
    "Gamma": "Γ",
    "Delta": "Δ",
    "Theta": "Θ",
    "Lambda": "Λ",
    "Xi": "Ξ",
    "Pi": "Π",
    "Sigma": "Σ",
    "Upsilon": "Υ",
    "Phi": "Φ",
    "Psi": "Ψ",
    "Omega": "Ω",
    # Operators & Relations
    "leq": "≤",
    "le": "≤",
    "geq": "≥",
    "ge": "≥",
    "neq": "≠",
    "ne": "≠",
    "approx": "≈",
    "sim": "∼",
    "simeq": "≃",
    "equiv": "≡",
    "propto": "∝",
    "pm": "±",
    "mp": "∓",
    "times": "×",
    "cdot": "⋅",
    "div": "÷",
    "ast": "∗",
    "star": "⋆",
    "in": "∈",
    "notin": "∉",
    "subset": "⊂",
    "subseteq": "⊆",
    "supset": "⊃",
    "supseteq": "⊇",
    "cup": "∪",
    "cap": "∩",
    "emptyset": "∅",
    "varnothing": "∅",
    "infty": "∞",
    "partial": "∂",
    "nabla": "∇",
    "ell": "ℓ",
    "forall": "∀",
    "exists": "∃",
    "neg": "¬",
    "to": "→",
    "rightarrow": "→",
    "leftarrow": "←",
    "leftrightarrow": "↔",
    "Rightarrow": "⇒",
    "Leftarrow": "⇐",
    "Leftrightarrow": "⇔",
    "mapsto": "↦",
    "lessgtr": "≶",
    "gtrless": "≷",
    "ll": "≪",
    "gg": "≫",
    "ldots": "…",
    "cdots": "⋯",
    "dots": "…",
    "vdots": "⋮",
    "ddots": "⋱",
    "mid": "∣",
    "parallel": "∥",
    "perp": "⊥",
    "prime": "′",
    "dag": "†",
    "ddag": "‡",
    # Spacing
    ",": "\u2009",
    ";": "\u2004",
    ":": "\u2005",
    "quad": "\u2003",
    "qquad": "\u2003\u2003",
    "!": "",
    "%": "%",
    "{": "{",
    "}": "}",
    "_": "_",
    "&": "&",
}

NARY_OPERATORS: dict[str, str] = {
    "sum": "∑",
    "prod": "∏",
    "coprod": "∐",
    "int": "∫",
    "iint": "∬",
    "iiint": "∭",
    "oint": "∮",
    "bigcup": "⋃",
    "bigcap": "⋂",
}

ACCENT_COMBINING: dict[str, str] = {
    "hat": "\u0302",
    "widehat": "\u0302",
    "bar": "\u0304",
    "overline": "\u0305",
    "tilde": "\u0303",
    "widetilde": "\u0303",
    "vec": "\u20d7",
    "dot": "\u0307",
    "ddot": "\u0308",
    "check": "\u030c",
    "breve": "\u0306",
}

MATHBB_MAP: dict[str, str] = {
    "1": "𝟙",
    "C": "ℂ",
    "E": "𝔼",
    "N": "ℕ",
    "P": "ℙ",
    "Q": "ℚ",
    "R": "ℝ",
    "V": "𝕍",
    "Z": "ℤ",
}

NAMED_FUNCTIONS: set[str] = {
    "ln",
    "log",
    "exp",
    "sin",
    "cos",
    "tan",
    "min",
    "max",
    "sup",
    "inf",
    "lim",
    "arg",
    "argmin",
    "argmax",
    "Pr",
    "Var",
    "Cov",
    "Corr",
    "logit",
    "probit",
}


def _make_math_run(text: str, normal: bool = False, bold: bool = False) -> etree._Element:
    """Create `<m:r>` with Cambria Math + SimSun (`宋体`) and pure black `#000000`."""
    mr = etree.Element(f"{M}r")
    if normal or bold:
        mrpr = etree.SubElement(mr, f"{M}rPr")
        if normal:
            etree.SubElement(mrpr, f"{M}nor")
        if bold and not normal:
            sty = etree.SubElement(mrpr, f"{M}sty")
            sty.set(f"{M}val", "b")

    wrpr = etree.SubElement(mr, f"{W}rPr")
    rf = etree.SubElement(wrpr, f"{W}rFonts")
    rf.set(f"{W}ascii", "Cambria Math")
    rf.set(f"{W}hAnsi", "Cambria Math")
    rf.set(f"{W}eastAsia", "宋体")
    rf.set(f"{W}cs", "Cambria Math")
    if bold and normal:
        etree.SubElement(wrpr, f"{W}b")
    col = etree.SubElement(wrpr, f"{W}color")
    col.set(f"{W}val", "000000")

    mt = etree.SubElement(mr, f"{M}t")
    if text.startswith(" ") or text.endswith(" "):
        mt.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    mt.text = text
    return mr


class _LatexMathParser:
    """Compact recursive-descent LaTeX math parser emitting Word OMML `<m:*>` elements."""

    def __init__(self, src: str) -> None:
        self.src = src.strip()
        self.pos = 0
        self.n = len(self.src)

    def _skip_ws(self) -> None:
        while self.pos < self.n and self.src[self.pos] in (" ", "\t", "\r", "\n"):
            self.pos += 1

    def _peek(self) -> str:
        self._skip_ws()
        return self.src[self.pos] if self.pos < self.n else ""

    def _read_command(self) -> str:
        """Assuming `self.src[self.pos] == '\\'`, consume and return command name."""
        self.pos += 1
        if self.pos >= self.n:
            return ""
        ch = self.src[self.pos]
        if not ch.isalpha():
            self.pos += 1
            return ch
        start = self.pos
        while self.pos < self.n and self.src[self.pos].isalpha():
            self.pos += 1
        return self.src[start : self.pos]

    def _read_raw_braced_text(self) -> str:
        """Read literal `{...}` text preserving internal spaces (used by `\\text{...}`)."""
        self._skip_ws()
        if self.pos >= self.n or self.src[self.pos] != "{":
            return self.src[self.pos : self.pos + 1]
        self.pos += 1
        depth = 1
        out: list[str] = []
        while self.pos < self.n and depth > 0:
            ch = self.src[self.pos]
            self.pos += 1
            if ch == "{":
                depth += 1
                out.append(ch)
            elif ch == "}":
                depth -= 1
                if depth > 0:
                    out.append(ch)
            else:
                out.append(ch)
        return "".join(out)

    def parse_nodes(self, stop_Right: bool = False, stop_char: str | None = None) -> list[etree._Element]:
        nodes: list[etree._Element] = []
        while True:
            self._skip_ws()
            if self.pos >= self.n:
                break
            if stop_char and self.src[self.pos] == stop_char:
                self.pos += 1
                break
            if stop_Right and self.src.startswith(r"\right", self.pos):
                break
            node = self.parse_scripted_atom()
            if node is not None:
                nodes.append(node)
        return nodes

    def parse_group_or_atom(self) -> list[etree._Element]:
        self._skip_ws()
        if self.pos < self.n and self.src[self.pos] == "{":
            self.pos += 1
            return self.parse_nodes(stop_char="}")
        atom = self.parse_atom()
        return [atom] if atom is not None else []

    def _parse_sub_sup(self) -> tuple[list[etree._Element] | None, list[etree._Element] | None]:
        sub_nodes: list[etree._Element] | None = None
        sup_nodes: list[etree._Element] | None = None
        while True:
            ch = self._peek()
            if ch == "_" and sub_nodes is None:
                self.pos += 1
                sub_nodes = self.parse_group_or_atom()
            elif ch == "^" and sup_nodes is None:
                self.pos += 1
                sup_nodes = self.parse_group_or_atom()
            else:
                break
        return sub_nodes, sup_nodes

    def parse_scripted_atom(self) -> etree._Element | None:
        self._skip_ws()
        if self.pos >= self.n:
            return None

        # Check N-ary operators (`\sum`, `\prod`, `\int`) before normal atom so limits attach cleanly
        saved_pos = self.pos
        if self.src[self.pos] == "\\":
            cmd = self._read_command()
            if cmd in NARY_OPERATORS:
                sub_nodes, sup_nodes = self._parse_sub_sup()
                body_atom = self.parse_scripted_atom()
                nary = etree.Element(f"{M}nary")
                npr = etree.SubElement(nary, f"{M}naryPr")
                chr_el = etree.SubElement(npr, f"{M}chr")
                chr_el.set(f"{M}val", NARY_OPERATORS[cmd])
                sub_hide = etree.SubElement(npr, f"{M}subHide")
                sub_hide.set(f"{M}val", "0" if sub_nodes else "1")
                sup_hide = etree.SubElement(npr, f"{M}supHide")
                sup_hide.set(f"{M}val", "0" if sup_nodes else "1")

                sub_el = etree.SubElement(nary, f"{M}sub")
                for el in sub_nodes or []:
                    sub_el.append(el)
                sup_el = etree.SubElement(nary, f"{M}sup")
                for el in sup_nodes or []:
                    sup_el.append(el)
                e_el = etree.SubElement(nary, f"{M}e")
                if body_atom is not None:
                    e_el.append(body_atom)
                else:
                    e_el.append(_make_math_run("\u200b"))
                return nary
            self.pos = saved_pos

        base_nodes = self.parse_group_or_atom()
        if not base_nodes:
            return None

        sub_nodes, sup_nodes = self._parse_sub_sup()
        if sub_nodes is None and sup_nodes is None:
            if len(base_nodes) == 1:
                return base_nodes[0]
            # Wrap multi-node group in an unbordered `<m:box>` so scripts/flow stay grouped
            box = etree.Element(f"{M}box")
            e_el = etree.SubElement(box, f"{M}e")
            for el in base_nodes:
                e_el.append(el)
            return box

        if sub_nodes is not None and sup_nodes is not None:
            wrap = etree.Element(f"{M}sSubSup")
            e_el = etree.SubElement(wrap, f"{M}e")
            for el in base_nodes:
                e_el.append(el)
            sub_el = etree.SubElement(wrap, f"{M}sub")
            for el in sub_nodes:
                sub_el.append(el)
            sup_el = etree.SubElement(wrap, f"{M}sup")
            for el in sup_nodes:
                sup_el.append(el)
            return wrap

        if sub_nodes is not None:
            wrap = etree.Element(f"{M}sSub")
            e_el = etree.SubElement(wrap, f"{M}e")
            for el in base_nodes:
                e_el.append(el)
            sub_el = etree.SubElement(wrap, f"{M}sub")
            for el in sub_nodes:
                sub_el.append(el)
            return wrap

        wrap = etree.Element(f"{M}sSup")
        e_el = etree.SubElement(wrap, f"{M}e")
        for el in base_nodes:
            e_el.append(el)
        sup_el = etree.SubElement(wrap, f"{M}sup")
        for el in sup_nodes or []:
            sup_el.append(el)
        return wrap

    def parse_atom(self) -> etree._Element | None:
        self._skip_ws()
        if self.pos >= self.n:
            return None

        ch = self.src[self.pos]
        if ch == "\\":
            cmd = self._read_command()
            if not cmd:
                return None

            if cmd == "frac":
                num_nodes = self.parse_group_or_atom()
                den_nodes = self.parse_group_or_atom()
                f_el = etree.Element(f"{M}f")
                num_el = etree.SubElement(f_el, f"{M}num")
                for el in num_nodes or [_make_math_run("")]:
                    num_el.append(el)
                den_el = etree.SubElement(f_el, f"{M}den")
                for el in den_nodes or [_make_math_run("")]:
                    den_el.append(el)
                return f_el

            if cmd == "sqrt":
                deg_nodes: list[etree._Element] | None = None
                if self._peek() == "[":
                    self.pos += 1
                    deg_nodes = self.parse_nodes(stop_char="]")
                rad_nodes = self.parse_group_or_atom()
                rad_el = etree.Element(f"{M}rad")
                rpr = etree.SubElement(rad_el, f"{M}radPr")
                dh = etree.SubElement(rpr, f"{M}degHide")
                dh.set(f"{M}val", "0" if deg_nodes else "1")
                deg_el = etree.SubElement(rad_el, f"{M}deg")
                for el in deg_nodes or []:
                    deg_el.append(el)
                e_el = etree.SubElement(rad_el, f"{M}e")
                for el in rad_nodes or [_make_math_run("")]:
                    e_el.append(el)
                return rad_el

            if cmd in ACCENT_COMBINING:
                inner_nodes = self.parse_group_or_atom()
                acc_el = etree.Element(f"{M}acc")
                apr = etree.SubElement(acc_el, f"{M}accPr")
                chr_el = etree.SubElement(apr, f"{M}chr")
                chr_el.set(f"{M}val", ACCENT_COMBINING[cmd])
                e_el = etree.SubElement(acc_el, f"{M}e")
                for el in inner_nodes or [_make_math_run("")]:
                    e_el.append(el)
                return acc_el

            if cmd == "left":
                beg = self._read_delimiter()
                inner = self.parse_nodes(stop_Right=True)
                end = ")"
                if self.src.startswith(r"\right", self.pos):
                    self.pos += len(r"\right")
                    end = self._read_delimiter()
                d_el = etree.Element(f"{M}d")
                dpr = etree.SubElement(d_el, f"{M}dPr")
                b_el = etree.SubElement(dpr, f"{M}begChr")
                b_el.set(f"{M}val", "" if beg == "." else beg)
                e_chr = etree.SubElement(dpr, f"{M}endChr")
                e_chr.set(f"{M}val", "" if end == "." else end)
                e_el = etree.SubElement(d_el, f"{M}e")
                for el in inner or [_make_math_run("")]:
                    e_el.append(el)
                return d_el

            if cmd in ("text", "textrm", "mathrm", "operatorname", "textit"):
                txt = self._read_raw_braced_text()
                return _make_math_run(txt, normal=True)

            if cmd in ("textbf", "mathbf", "boldsymbol", "bm"):
                txt = self._read_raw_braced_text()
                # Map inner Greek commands if e.g. \boldsymbol{\beta}
                if txt.startswith("\\") and txt[1:] in LATEX_SYMBOLS:
                    txt = LATEX_SYMBOLS[txt[1:]]
                return _make_math_run(txt, normal=False, bold=True)

            if cmd == "mathbb":
                raw_bb = self._read_raw_braced_text().strip()
                mapped = "".join(MATHBB_MAP.get(c, c) for c in raw_bb)
                return _make_math_run(mapped, normal=False)

            if cmd in NAMED_FUNCTIONS:
                return _make_math_run(cmd, normal=True)

            if cmd in LATEX_SYMBOLS:
                val = LATEX_SYMBOLS[cmd]
                return _make_math_run(val) if val else None

            # Fallback unknown command: render command name upright so nothing silently disappears
            return _make_math_run(cmd, normal=True)

        # Plain character or run of digits/letters
        self.pos += 1
        if ch == "-":
            return _make_math_run("−")
        return _make_math_run(ch)

    def _read_delimiter(self) -> str:
        self._skip_ws()
        if self.pos >= self.n:
            return ""
        if self.src[self.pos] == "\\":
            cmd = self._read_command()
            delim_map = {"{": "{", "}": "}", "langle": "⟨", "rangle": "⟩", "|": "‖", "Vert": "‖"}
            return delim_map.get(cmd, cmd)
        ch = self.src[self.pos]
        self.pos += 1
        return ch


def latex_to_omml_element(latex: str, display: bool = False) -> etree._Element:
    """Convert a LaTeX math string into `<m:oMath>` (inline) or `<m:oMathPara>` (display)."""
    parser = _LatexMathParser(latex)
    nodes = parser.parse_nodes()
    omath = etree.Element(f"{M}oMath")
    for el in nodes or [_make_math_run("")]:
        omath.append(el)
    if not display:
        return omath
    para = etree.Element(f"{M}oMathPara")
    para.append(omath)
    return para
