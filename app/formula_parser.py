"""Excel 公式引用解析。

从公式文本中抽取：
  - 单元格引用 / 矩形区域引用（含跨工作表、带引号工作表名、整行/整列区域）；
  - 名称管理器中的**定义名称**（含 `工作表!名称` 形式的显式限定）。

输出尚未展开的 RawRef / RawName 列表，由 graph 模块结合工作表实际尺寸、
名称作用域展开为具体单元格。
"""
from __future__ import annotations

import re
from dataclasses import dataclass

from openpyxl.utils import column_index_from_string

# 工作表名前缀：'月报 2026'! 或 Sheet1! 或 参数!
_SHEET = r"(?:'(?P<qsheet>[^']+(?:''[^']*)*)'|(?P<usheet>[A-Za-z0-9_.一-鿿]+))!"
# 显式限定名称用的工作表前缀（分组名不能与上面的复用）
_QNAME_SHEET = (
    r"(?:'(?P<qname_qsheet>[^']+(?:''[^']*)*)'"
    r"|(?P<qname_usheet>[A-Za-z0-9_.一-鿿]+))!"
)
_CELL = r"\$?[A-Za-z]{1,3}\$?[0-9]{1,7}"
_COL = r"\$?[A-Za-z]{1,3}"
_ROW = r"\$?[0-9]{1,7}"
# 定义名称：首字符为字母 / 下划线 / 反斜杠 / CJK，不能以数字开头
_NAME = r"(?![0-9])[A-Za-z_\\一-鿿][A-Za-z0-9_.\\一-鿿]*"

_TOKEN_RE = re.compile(
    r"""
      (?P<string>"(?:[^"]|"")*")                        # 字符串字面量，忽略
    | (?P<error>\#[A-Za-z0-9!/?]+)                      # #REF! 等错误字面量，忽略
    | (?P<sheetref>{sheet}(?:{cell}(?::{cell})?|{col}:{col}|{row}:{row}))
    | (?P<qname>{qname_sheet}(?P<qname_name>{name}))   # 显式限定的定义名称 Sheet!Name（须排在单元格引用之后）
    | (?P<range>(?<![0-9.]){cell}:{cell}|{col}:{col}|{row}:{row})  # 本表区域
    | (?P<struct>{name}\[[^\[\]]*\])                    # 表结构化引用 表名[列名]，忽略
    | (?P<func>[A-Za-z_][A-Za-z0-9_.]*(?=\s*\())        # 函数名，忽略
    | (?P<cell>(?<![0-9.]){cell})(?!\s*\()              # 本表单单元格（排除 1E2 科学计数法）
    | (?P<name>(?<![0-9A-Za-z_.\\一-鿿]){name})(?!\s*\()  # 裸定义名称（函数调用除外）
    | (?P<other>.)
    """.format(sheet=_SHEET, qname_sheet=_QNAME_SHEET,
               cell=_CELL, col=_COL, row=_ROW, name=_NAME),
    re.VERBOSE,
)
_CELL_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?([0-9]{1,7})$")

# Excel 内置布尔常量，不属于用户定义名称
_BUILTIN_NAMES = {"TRUE", "FALSE"}


@dataclass(frozen=True)
class RawRef:
    """一条尚未展开的引用。行/列为 None 表示整列/整行区域，需按工作表尺寸截断。"""

    sheet: str | None  # None 表示公式所在工作表
    kind: str  # "cell" | "range" | "cols" | "rows"
    col1: int | None
    row1: int | None
    col2: int | None
    row2: int | None
    text: str  # 原始文本，用于展示


@dataclass(frozen=True)
class RawName:
    """公式中的一个定义名称记号。sheet 非 None 表示 `工作表!名称` 显式限定。"""

    sheet: str | None
    name: str
    text: str


def _split_cell(token: str) -> tuple[int, int]:
    m = _CELL_RE.match(token)
    if not m:
        raise ValueError(f"非法单元格引用: {token!r}")
    return column_index_from_string(m.group(1).upper()), int(m.group(2))


def _split_col(token: str) -> int:
    return column_index_from_string(token.replace("$", "").upper())


def _split_row(token: str) -> int:
    return int(token.replace("$", ""))


def _build_ref(sheet: str | None, body: str, full_text: str) -> RawRef | None:
    """把引用主体（不含工作表前缀）解析为 RawRef。"""
    if ":" in body:
        left, right = body.split(":", 1)
        if _CELL_RE.match(left) and _CELL_RE.match(right):
            c1, r1 = _split_cell(left)
            c2, r2 = _split_cell(right)
            return RawRef(sheet, "range", min(c1, c2), min(r1, r2), max(c1, c2), max(r1, r2), full_text)
        if re.fullmatch(_COL, left) and re.fullmatch(_COL, right):
            c1, c2 = _split_col(left), _split_col(right)
            return RawRef(sheet, "cols", min(c1, c2), None, max(c1, c2), None, full_text)
        if re.fullmatch(_ROW, left) and re.fullmatch(_ROW, right):
            r1, r2 = _split_row(left), _split_row(right)
            return RawRef(sheet, "rows", None, min(r1, r2), None, max(r1, r2), full_text)
        return None
    if _CELL_RE.match(body):
        c, r = _split_cell(body)
        return RawRef(sheet, "cell", c, r, c, r, full_text)
    return None


def parse_formula(formula: str) -> list[RawRef]:
    """解析公式文本，返回引用列表（未展开）。解析失败时返回已识别部分。"""
    return [r for r, _ in parse_formula_tokens(formula)]


def parse_formula_tokens(
    formula: str,
) -> list[tuple[RawRef | RawName, str]]:
    """解析公式，返回 (引用/名称, 类型) 二元组列表，类型为 "ref" 或 "name"。"""
    if not formula:
        return []
    text = formula[1:] if formula.startswith("=") else formula
    out: list[tuple[RawRef | RawName, str]] = []
    for m in _TOKEN_RE.finditer(text):
        kind = m.lastgroup
        if kind == "sheetref":
            sheet = m.group("qsheet")
            if sheet is not None:
                sheet = sheet.replace("''", "'")
            else:
                sheet = m.group("usheet")
            body = m.group("sheetref")
            body = body[body.index("!") + 1:]
            ref = _build_ref(sheet, body, m.group("sheetref"))
        elif kind == "range":
            ref = _build_ref(None, m.group("range"), m.group("range"))
        elif kind == "cell":
            ref = _build_ref(None, m.group("cell"), m.group("cell"))
        elif kind == "qname":
            qs = m.group("qname_qsheet")
            sheet = qs.replace("''", "'") if qs is not None else m.group("qname_usheet")
            out.append((RawName(sheet, m.group("qname_name"), m.group("qname")), "name"))
            continue
        elif kind == "name":
            token = m.group("name")
            if token.upper() in _BUILTIN_NAMES:
                continue
            out.append((RawName(None, token, token), "name"))
            continue
        else:
            continue
        if ref is not None:
            out.append((ref, "ref"))
    return out
