"""单元格依赖图：构图（含名称管理器解析）、循环引用检测、影响面/依赖链查询。"""
from __future__ import annotations

import re
from collections import deque
from dataclasses import dataclass, field

from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName

from .formula_parser import RawName, RawRef, parse_formula, parse_formula_tokens

# 单个区域最多展开的单元格数，防止 SUM(A:XFD) 之类打爆内存
MAX_RANGE_CELLS = 5000
# 影响面查询的默认上限
MAX_IMPACT_NODES = 5000
# 依赖健康审计：扇出（直接引用单元格数）告警的默认阈值
DEFAULT_FANOUT_THRESHOLD = 20
# 单条审计问题附带的证据边上限，防止超大区域把响应撑爆
MAX_AUDIT_EVIDENCE = 100

# 审计问题级别 -> 排序权重（越严重越靠前）
SEVERITY_WEIGHT = {"error": 0, "warning": 1, "info": 2}
AUDIT_LEVELS = ("error", "warning", "info")

# 定义名称解析状态
NAME_OK = "ok"                    # 成功解析到本工作簿的单元格 / 区域
NAME_CONSTANT = "constant"        # 定义为常量 / 非引用公式，无依赖边
NAME_UNDEFINED = "undefined"      # 按作用域找不到该名称
NAME_BAD_TARGET = "bad_target"    # 定义目标损坏（#REF! / 外部工作簿 / 空定义）
NAME_MISSING_SHEET = "missing_sheet"  # 目标指向不存在的工作表
# 不会产生依赖边（即不进入影响面）的名称解析状态
NAME_UNRESOLVED_STATES = frozenset(
    {NAME_UNDEFINED, NAME_BAD_TARGET, NAME_MISSING_SHEET}
)


def node_id(sheet: str, coord: str) -> str:
    return f"{sheet}!{coord}"


def split_node(node: str) -> tuple[str, str]:
    """工作表名本身可含 '!'，坐标部分不会，故从右侧切分。"""
    sheet, coord = node.rsplit("!", 1)
    return sheet, coord


def display_id(sheet: str, coord: str) -> str:
    """展示用 id：含特殊字符的工作表名加引号。"""
    if sheet.replace("_", "").isalnum():
        return f"{sheet}!{coord}"
    return f"'{sheet}'!{coord}"


@dataclass
class SheetInfo:
    name: str
    max_row: int
    max_col: int
    formula_count: int = 0


@dataclass(frozen=True)
class DefinedNameInfo:
    """名称管理器中的一条定义名称。"""

    name: str
    scope: str | None               # None=工作簿级，否则为工作表名
    target: str                     # 定义原文，如 参数!$B$2 / 3.14 / #REF!
    status: str                     # ok / constant / bad_target / missing_sheet
    refs: tuple[RawRef, ...]        # 目标解析出的引用（constant 为空元组）


@dataclass(frozen=True)
class NameUse:
    """公式中的一次名称使用及其解析结果。"""

    source: str                     # 公式单元格内部 id
    formula: str                    # 公式原文
    name: str                       # 名称本身
    qualifier: str | None           # `工作表!名称` 中的限定工作表，None 为裸用
    text: str                       # 公式中的原始记号，如 参数!税率
    scope: str | None               # 解析命中的定义作用域；未解析为 None
    target: str | None              # 命中的定义原文；未解析为 None
    status: str                     # 解析状态，见 NAME_* 常量


@dataclass(frozen=True)
class MissingSheetRef:
    """跨表引用指向不存在的工作表（构图时直接跳过，只在此留证）。"""

    source: str        # 公式单元格内部 id
    formula: str       # 公式原文
    sheet: str         # 被引用但不存在的工作表名
    ref_text: str      # 引用原文，如 NoSuch!B2


@dataclass(frozen=True)
class OutOfRangeRef:
    """显式区域 / 单元格引用超出目标工作表的已用范围。"""

    source: str        # 公式单元格内部 id
    formula: str       # 公式原文
    sheet: str         # 目标工作表名
    ref_text: str      # 引用原文，如 Data!A99:A101
    coord: str         # 第一个越界坐标（展示用），如 A99


def _parse_target_refs(target: str) -> list[RawRef]:
    """把名称管理器中名称的定义文本解析为引用列表。

    仅当整个定义是**纯引用**（单元格 / 区域，或以逗号分隔的多个引用并集，
    如 `参数!$B$2:B5`、`Sheet1!$A$1,Sheet1!$C$1`、整行整列）时返回 RawRef；
    常量（`3.14`）或含运算的表达式（`A1+1`）返回空列表——这类名称不指向
    确定的单元格 / 区域，不能展开为依赖边。
    """
    text = target[1:] if target.startswith("=") else target
    tokens = parse_formula_tokens(text)
    if not tokens or any(kind != "ref" for _, kind in tokens):
        return []
    refs = [r for r, _ in tokens]
    compact = re.sub(r"\s+", "", text)
    if ",".join(r.text for r in refs) != compact:
        return []
    return refs


@dataclass
class WorkbookGraph:
    sheets: list[SheetInfo] = field(default_factory=list)
    formulas: dict[str, str] = field(default_factory=dict)      # node -> 公式文本
    values: dict[str, str] = field(default_factory=dict)        # node -> 缓存计算值
    precedents: dict[str, set[str]] = field(default_factory=dict)  # node -> 它引用的节点
    dependents: dict[str, set[str]] = field(default_factory=dict)  # node -> 引用它的节点
    cycles: list[list[str]] = field(default_factory=list)
    cycle_nodes: set[str] = field(default_factory=set)
    edge_count: int = 0
    missing_sheet_refs: list[MissingSheetRef] = field(default_factory=list)
    out_of_range_refs: list[OutOfRangeRef] = field(default_factory=list)
    # 名称管理器
    defined_names: list[DefinedNameInfo] = field(default_factory=list)
    name_uses: list[NameUse] = field(default_factory=list)

    # ------------------------------------------------------------------ 构图
    @classmethod
    def from_workbooks(cls, wb_formula, wb_values) -> "WorkbookGraph":
        g = cls()
        bounds: dict[str, tuple[int, int]] = {}
        for ws in wb_formula.worksheets:
            bounds[ws.title] = (ws.max_row or 1, ws.max_column or 1)
            g.sheets.append(SheetInfo(ws.title, ws.max_row or 1, ws.max_column or 1))

        ws_values = {ws.title: ws for ws in wb_values.worksheets}
        sheet_names = set(bounds)

        g._collect_defined_names(wb_formula, sheet_names)
        global_names, local_names = g._name_indexes()
        seen_name_uses: set[tuple] = set()

        for ws in wb_formula.worksheets:
            wsv = ws_values.get(ws.title)
            info = next(s for s in g.sheets if s.name == ws.title)
            for row in ws.iter_rows():
                for cell in row:
                    formula = cell.value
                    if not isinstance(formula, str) or not formula.startswith("="):
                        continue
                    coord = cell.coordinate
                    node = node_id(ws.title, coord)
                    g.formulas[node] = formula
                    info.formula_count += 1
                    if wsv is not None:
                        v = wsv[coord].value
                        if v is not None:
                            g.values[node] = str(v)
                    for token, kind in parse_formula_tokens(formula):
                        if kind == "ref":
                            g._link_ref(token, node, formula, ws.title,
                                        bounds, sheet_names, ws_values)
                        else:
                            g._link_name(token, node, formula, ws.title,
                                         bounds, sheet_names, ws_values,
                                         global_names, local_names, seen_name_uses)

        g._detect_cycles()
        return g

    # ---------------------------------------------------------- 名称管理器
    def _collect_defined_names(self, wb_formula, sheet_names) -> None:
        """读取工作簿级与工作表级定义名称，并预先解析其目标。

        openpyxl 加载后：工作簿级名称在 wb.defined_names，工作表级名称在
        ws.defined_names（localSheetId 指向该表）。
        """
        sheet_by_index = {i: ws.title for i, ws in enumerate(wb_formula.worksheets)}

        def add(dn: DefinedName, scope: str | None) -> None:
            raw = dn.attr_text if dn.attr_text is not None else dn.value
            # openpyxl 在空定义时可能保留 None 或序列化为字面量 "none"
            target = (raw or "").strip()
            if not target or target.lower() == "none":
                status, refs = NAME_BAD_TARGET, ()
            elif "#REF!" in target.upper():
                status, refs = NAME_BAD_TARGET, ()
            elif target.lstrip().startswith("["):
                # 外部工作簿引用 [Book.xlsx]Sheet!A1，本工作簿内无法定位
                status, refs = NAME_BAD_TARGET, ()
            else:
                # 未限定工作表的相对引用：局部名称归属其作用域表，
                # 工作簿级名称在实际使用时归属公式所在表（此处保留 None）。
                default_sheet = scope
                parsed = tuple(
                    self._qualify_ref(r, default_sheet)
                    for r in _parse_target_refs(target)
                )
                if not parsed:
                    status = NAME_CONSTANT  # 常量或非引用公式（如 3.14 / A1+1）
                elif any(r.sheet not in sheet_names for r in parsed):
                    status = NAME_MISSING_SHEET
                else:
                    status = NAME_OK
                refs = parsed
            self.defined_names.append(
                DefinedNameInfo(dn.name, scope, target, status, refs)
            )

        for dn in wb_formula.defined_names.values():
            add(dn, None)
        for ws in wb_formula.worksheets:
            scope = ws.title
            for dn in ws.defined_names.values():
                # 防御：localSheetId 与所在表不一致时以 localSheetId 为准
                if dn.localSheetId is not None and dn.localSheetId in sheet_by_index:
                    scope = sheet_by_index[dn.localSheetId]
                add(dn, scope)

    def _name_indexes(
        self,
    ) -> tuple[dict[str, DefinedNameInfo], dict[str, dict[str, DefinedNameInfo]]]:
        global_names: dict[str, DefinedNameInfo] = {}
        local_names: dict[str, dict[str, DefinedNameInfo]] = {}
        for dn in self.defined_names:
            if dn.scope is None:
                global_names.setdefault(dn.name, dn)
            else:
                local_names.setdefault(dn.scope, {})[dn.name] = dn
        return global_names, local_names

    @staticmethod
    def _qualify_ref(ref: RawRef, default_sheet: str | None) -> RawRef:
        if ref.sheet is not None:
            return ref
        return RawRef(default_sheet, ref.kind, ref.col1, ref.row1,
                      ref.col2, ref.row2, ref.text)

    def _resolve_name(
        self,
        token: RawName,
        owner_sheet: str,
        global_names: dict[str, DefinedNameInfo],
        local_names: dict[str, dict[str, DefinedNameInfo]],
    ) -> DefinedNameInfo | None:
        """按 Excel 作用域规则解析名称。

        - 裸名称：先取公式所在工作表的工作表级同名名称，再回退工作簿级；
        - `工作表!名称`：限定名优先取该工作表的局部名，再回退工作簿级。
        """
        if token.sheet is not None:
            hit = local_names.get(token.sheet, {}).get(token.name)
            return hit if hit is not None else global_names.get(token.name)
        hit = local_names.get(owner_sheet, {}).get(token.name)
        return hit if hit is not None else global_names.get(token.name)

    def _link_name(self, token: RawName, node: str, formula: str,
                   owner_sheet: str, bounds, sheet_names, ws_values,
                   global_names, local_names, seen_uses: set) -> None:
        """处理公式中的一个名称记号：解析、登记使用记录、展开依赖边。"""
        dn = self._resolve_name(token, owner_sheet, global_names, local_names)
        if dn is None:
            scope, target, status = None, None, NAME_UNDEFINED
        else:
            scope, target, status = dn.scope, dn.target, dn.status

        key = (node, token.text, status)
        if key not in seen_uses:
            seen_uses.add(key)
            self.name_uses.append(
                NameUse(node, formula, token.name, token.sheet,
                        token.text, scope, target, status)
            )

        if status != NAME_OK:
            return  # 无法解析 / 目标无效：不产生任何边，不进入影响面

        # 名称定义里未限定工作表的引用，工作簿级名称归属公式所在表
        for ref in dn.refs:
            sheet = ref.sheet if ref.sheet is not None else owner_sheet
            fixed = ref if ref.sheet == sheet else self._qualify_ref(ref, sheet)
            self._expand_and_link(fixed, node, formula, bounds,
                                  sheet_names, ws_values)

    # ------------------------------------------------------------- 边展开
    def _link_ref(self, ref: RawRef, node: str, formula: str, owner_sheet: str,
                  bounds, sheet_names, ws_values) -> None:
        """处理公式中的一条直接引用。"""
        target_sheet = ref.sheet if ref.sheet is not None else owner_sheet
        fixed = ref if ref.sheet is not None else self._qualify_ref(ref, owner_sheet)
        if target_sheet not in sheet_names:
            self.missing_sheet_refs.append(
                MissingSheetRef(node, formula, target_sheet, ref.text)
            )
            return  # 引用不存在的工作表，跳过
        self._expand_and_link(fixed, node, formula, bounds, sheet_names, ws_values)

    def _expand_and_link(self, ref: RawRef, node: str, formula: str,
                         bounds, sheet_names, ws_values) -> None:
        """展开 RawRef（其 sheet 必须为现存工作表）并登记依赖边。"""
        target_sheet = ref.sheet
        wst = ws_values.get(target_sheet)
        coords, out_coord = self._expand(ref, bounds[target_sheet])
        if out_coord is not None:
            self.out_of_range_refs.append(
                OutOfRangeRef(node, formula, target_sheet, ref.text, out_coord)
            )
        for tcoord in coords:
            tnode = node_id(target_sheet, tcoord)
            g_prec = self.precedents.setdefault(node, set())
            if tnode not in g_prec:
                g_prec.add(tnode)
                self.dependents.setdefault(tnode, set()).add(node)
                self.edge_count += 1
            # 叶子单元格也记录值，便于前端展示
            if tnode not in self.formulas and tnode not in self.values and wst is not None:
                v = wst[tcoord].value
                if v is not None:
                    self.values[tnode] = str(v)

    @staticmethod
    def _expand(
        ref: RawRef, bounds: tuple[int, int]
    ) -> tuple[list[str], str | None]:
        """把 RawRef 展开为 (坐标列表, 第一个越界坐标)。

        显式区域（A1:B10）按原样展开；整行/整列区域按工作表已用范围截断
        （语义上即覆盖全表，越界属正常，不报问题）；显式单元格 / 区域
        超出已用范围时截断到范围内，并返回越界证据坐标。
        超过 MAX_RANGE_CELLS 时等比收缩，防止内存被打爆。
        """
        max_row, max_col = bounds
        out_coord: str | None = None
        if ref.kind == "cell":
            if ref.row1 <= max_row and ref.col1 <= max_col:
                return [f"{get_column_letter(ref.col1)}{ref.row1}"], None
            return [], f"{get_column_letter(ref.col1)}{ref.row1}"
        if ref.kind == "range":
            r1, r2, c1, c2 = ref.row1, ref.row2, ref.col1, ref.col2
            if r2 > max_row or c2 > max_col:
                out_coord = f"{get_column_letter(c2)}{r2}"
                r2, c2 = min(r2, max_row), min(c2, max_col)
            # 整个区域都在已用范围之外
            if r1 > r2 or c1 > c2:
                return [], out_coord
        elif ref.kind == "cols":
            r1, r2 = 1, max_row
            c1, c2 = ref.col1, min(ref.col2, max_col)
        else:  # rows
            r1, r2 = ref.row1, ref.row2
            c1, c2 = 1, max_col
        while (r2 - r1 + 1) * (c2 - c1 + 1) > MAX_RANGE_CELLS and r2 > r1:
            r2 -= 1
        while (r2 - r1 + 1) * (c2 - c1 + 1) > MAX_RANGE_CELLS and c2 > c1:
            c2 -= 1
        coords = [
            f"{get_column_letter(c)}{r}"
            for r in range(r1, r2 + 1)
            for c in range(c1, c2 + 1)
        ]
        return coords, out_coord

    # -------------------------------------------------------------- 环检测
    def _detect_cycles(self) -> None:
        """Tarjan 强连通分量（迭代实现）。大小>1 的分量或自环即循环引用。"""
        adj = self.precedents
        index_of: dict[str, int] = {}
        lowlink: dict[str, int] = {}
        on_stack: set[str] = set()
        stack: list[str] = []
        counter = 0

        for root in list(adj):
            if root in index_of:
                continue
            work = [(root, iter(adj.get(root, ())))]
            index_of[root] = lowlink[root] = counter
            counter += 1
            stack.append(root)
            on_stack.add(root)
            while work:
                node, it = work[-1]
                advanced = False
                for nxt in it:
                    if nxt not in adj and nxt not in index_of:
                        continue  # 非公式叶子节点不参与环
                    if nxt not in index_of:
                        index_of[nxt] = lowlink[nxt] = counter
                        counter += 1
                        stack.append(nxt)
                        on_stack.add(nxt)
                        work.append((nxt, iter(adj.get(nxt, ()))))
                        advanced = True
                        break
                    elif nxt in on_stack:
                        lowlink[node] = min(lowlink[node], index_of[nxt])
                if advanced:
                    continue
                work.pop()
                if work:
                    parent = work[-1][0]
                    lowlink[parent] = min(lowlink[parent], lowlink[node])
                if lowlink[node] == index_of[node]:
                    scc = []
                    while True:
                        w = stack.pop()
                        on_stack.discard(w)
                        scc.append(w)
                        if w == node:
                            break
                    is_cycle = len(scc) > 1 or any(
                        n in adj.get(n, ()) for n in scc
                    )
                    if is_cycle:
                        self.cycles.append(sorted(scc))
                        self.cycle_nodes.update(scc)

    # -------------------------------------------------------------- 查询
    def neighbors(self, node: str, direction: str) -> set[str]:
        adj = self.dependents if direction == "downstream" else self.precedents
        return adj.get(node, set())

    def impact(
        self,
        node: str,
        direction: str = "downstream",
        max_depth: int = 50,
        max_nodes: int = MAX_IMPACT_NODES,
    ) -> dict:
        """BFS 求某单元格的影响面（下游）或依赖链（上游）。

        返回每个可达节点的最短路径；被环截断的节点会标记 in_cycle。
        """
        adj = self.dependents if direction == "downstream" else self.precedents
        depth_of = {node: 0}
        parent = {node: None}
        truncated = False
        queue = deque([node])
        while queue:
            cur = queue.popleft()
            d = depth_of[cur]
            if d >= max_depth:
                continue
            for nxt in sorted(adj.get(cur, ())):
                if nxt in depth_of:
                    continue
                if len(depth_of) >= max_nodes:
                    truncated = True
                    break
                depth_of[nxt] = d + 1
                parent[nxt] = cur
                queue.append(nxt)
            if truncated:
                break

        def path_to(n: str) -> list[str]:
            path = []
            while n is not None:
                path.append(n)
                n = parent[n]
            return list(reversed(path))

        items = [
            {
                "cell": n,
                "sheet": split_node(n)[0],
                "coord": split_node(n)[1],
                "display": display_id(*split_node(n)),
                "depth": depth_of[n],
                "formula": self.formulas.get(n),
                "value": self.values.get(n),
                "in_cycle": n in self.cycle_nodes,
                "path": [display_id(*split_node(p)) for p in path_to(n)],
                "path_raw": path_to(n),
            }
            for n in depth_of
            if n != node
        ]
        items.sort(key=lambda x: (x["depth"], x["cell"]))
        direct = sorted(adj.get(node, set()))
        return {
            "root": node,
            "direction": direction,
            "direct": [
                {
                    "cell": n,
                    "display": display_id(*split_node(n)),
                    "formula": self.formulas.get(n),
                    "value": self.values.get(n),
                    "in_cycle": n in self.cycle_nodes,
                }
                for n in direct
            ],
            "impacted": items,
            "truncated": truncated,
            "total": len(items),
        }

    # ------------------------------------------------------ 依赖健康审计
    def _edge(self, src: str, dst: str | None = None,
              dst_display: str | None = None, ref: str | None = None) -> dict:
        e: dict = {
            "from": src,
            "from_display": display_id(*split_node(src)) if src else None,
        }
        if dst is not None:
            e["to"] = dst
        if dst_display is not None:
            e["to_display"] = dst_display
        if ref is not None:
            e["ref"] = ref
        return e

    def audit(
        self,
        levels: tuple[str, ...] = AUDIT_LEVELS,
        sheet: str | None = None,
        fanout_threshold: int = DEFAULT_FANOUT_THRESHOLD,
        evidence_limit: int | None = MAX_AUDIT_EVIDENCE,
    ) -> dict:
        """扫描依赖图，生成依赖健康问题清单。

        问题类型：
          missing_sheet       error   跨表公式指向不存在的工作表
          undefined_name      error   公式中的定义名称按作用域无法解析
          invalid_name_target error   定义名称目标损坏（#REF! / 外部工作簿）或
                                       指向不存在的工作表
          empty_ref           warning 被公式引用却为空的单元格
          out_of_range        warning 引用超出目标工作表的已用范围
          high_fanout         warning 公式直接引用的单元格数超过阈值
          orphan_formula      info    未被任何公式引用的公式单元格
        sheet 筛选按问题所属工作表（名称 / 缺失工作表类按公式所在工作表）。
        """
        issues: list[dict] = []

        def add(issue: dict) -> None:
            if sheet is not None and issue["sheet"] != sheet:
                return
            if issue["level"] not in levels:
                return
            ev = issue["evidence"]
            if evidence_limit is not None and len(ev) > evidence_limit:
                issue["evidence"] = ev[:evidence_limit]
                issue["evidence_truncated"] = True
                issue["evidence_total"] = len(ev)
            issues.append(issue)

        # 1) 跨表引用指向不存在的工作表（同一公式 / 工作表 / 引用去重）
        seen: set[tuple] = set()
        for m in self.missing_sheet_refs:
            key = (m.source, m.sheet, m.ref_text)
            if key in seen:
                continue
            seen.add(key)
            add({
                "code": "missing_sheet",
                "level": "error",
                "sheet": split_node(m.source)[0],
                "cell": m.source,
                "cell_display": display_id(*split_node(m.source)),
                "message": f"引用了不存在的工作表 {m.sheet!r}（{m.ref_text}）",
                "formula": m.formula,
                "target_sheet": m.sheet,
                "target_display": m.ref_text,
                "evidence": [self._edge(m.source, None, m.ref_text, m.ref_text)],
            })

        # 1b) 名称管理器：无法解析 / 目标无效的名称使用（每处使用一条，可定位）
        def _scope_label(scope: str | None) -> str:
            return "工作簿级" if scope is None else f"工作表级（{scope}）"

        for u in self.name_uses:
            if u.status not in NAME_UNRESOLVED_STATES:
                continue
            owner = split_node(u.source)[0]
            if u.qualifier:
                q = u.qualifier
                qualifier = f"{q}!" if q.replace("_", "").isalnum() else f"'{q}'!"
            else:
                qualifier = ""
            name_display = f"{qualifier}{u.name}"
            if u.status == NAME_UNDEFINED:
                code = "undefined_name"
                scope_hint = (
                    f"{_scope_label(owner)}同名名称" if u.qualifier is None
                    else f"工作表 {u.qualifier!r} 的工作表级或工作簿级同名名称"
                )
                message = (f"名称 {u.name!r} 无法解析：工作簿中不存在"
                           f"{scope_hint}（记号 {name_display}）")
                target_display = None
            else:
                code = "invalid_name_target"
                # missing_sheet 时 target 中的缺失工作表名
                missing = next(
                    (r.sheet for dn in self.defined_names
                     if dn.name == u.name and dn.scope == u.scope
                     for r in dn.refs if r.sheet not in {s.name for s in self.sheets}),
                    None,
                )
                if u.status == NAME_MISSING_SHEET:
                    reason = f"名称定义指向不存在的工作表 {missing!r}"
                elif (u.target or "").upper().find("#REF!") >= 0:
                    reason = "名称定义为损坏的 #REF! 引用"
                elif (u.target or "").lstrip().startswith("["):
                    reason = "名称定义指向外部工作簿，本文件内无法定位"
                else:
                    reason = "名称定义为空或无法解析"
                message = (f"名称 {u.name!r} 的定义无效：{reason}"
                           f"（{_scope_label(u.scope)}，定义为 {u.target!r}）")
                target_display = u.target
            add({
                "code": code,
                "level": "error",
                "sheet": owner,
                "cell": u.source,
                "cell_display": display_id(owner, split_node(u.source)[1]),
                "message": message,
                "formula": u.formula,
                "name": u.name,
                "name_display": name_display,
                "name_scope": u.scope,
                "name_qualifier": u.qualifier,
                "name_target": u.target,
                "name_status": u.status,
                "target_display": target_display,
                "evidence": [self._edge(u.source, None, target_display, u.text)],
            })

        # 2) 显式引用超出目标工作表的已用范围
        seen.clear()
        for o in self.out_of_range_refs:
            key = (o.source, o.sheet, o.ref_text)
            if key in seen:
                continue
            seen.add(key)
            tinfo = next((s for s in self.sheets if s.name == o.sheet), None)
            target_display = display_id(o.sheet, o.coord)
            add({
                "code": "out_of_range",
                "level": "warning",
                "sheet": split_node(o.source)[0],
                "cell": o.source,
                "cell_display": display_id(*split_node(o.source)),
                "message": "引用 {ref} 超出工作表「{s}」已用范围"
                           "（{r} 行 × {c} 列）".format(
                    ref=o.ref_text, s=o.sheet,
                    r=tinfo.max_row if tinfo else "?",
                    c=tinfo.max_col if tinfo else "?"),
                "formula": o.formula,
                "target_sheet": o.sheet,
                "target_display": target_display,
                "evidence": [
                    self._edge(o.source, node_id(o.sheet, o.coord),
                               target_display, o.ref_text)
                ],
            })

        # 3) 被公式引用却为空的单元格（无公式、无值；按目标单元格聚合证据边）
        for target in sorted(self.dependents):
            if target in self.formulas or target in self.values:
                continue
            sources = sorted(self.dependents[target])
            s_sheet, s_coord = split_node(target)
            add({
                "code": "empty_ref",
                "level": "warning",
                "sheet": s_sheet,
                "cell": target,
                "cell_display": display_id(s_sheet, s_coord),
                "message": f"单元格 {display_id(s_sheet, s_coord)} 为空，"
                           f"但被 {len(sources)} 个公式引用",
                "formula": None,
                "referenced_by": sources,
                "evidence": [self._edge(s, target) for s in sources],
            })

        # 4) 未被任何公式引用的公式单元格（孤立公式）
        for node in sorted(self.formulas):
            if self.dependents.get(node):
                continue
            n_sheet, n_coord = split_node(node)
            add({
                "code": "orphan_formula",
                "level": "info",
                "sheet": n_sheet,
                "cell": node,
                "cell_display": display_id(n_sheet, n_coord),
                "message": "该公式未被任何其他公式引用，其结果不参与其他计算",
                "formula": self.formulas[node],
                "evidence": [],
            })

        # 5) 扇出过大的公式节点（直接依赖单元格数 > 阈值）
        for node in sorted(self.precedents):
            fanout = len(self.precedents[node])
            if fanout <= fanout_threshold:
                continue
            f_sheet, f_coord = split_node(node)
            add({
                "code": "high_fanout",
                "level": "warning",
                "sheet": f_sheet,
                "cell": node,
                "cell_display": display_id(f_sheet, f_coord),
                "message": f"直接引用 {fanout} 个单元格，超过扇出阈值 "
                           f"{fanout_threshold}",
                "formula": self.formulas[node],
                "fanout": fanout,
                "threshold": fanout_threshold,
                "evidence": [self._edge(node, t)
                             for t in sorted(self.precedents[node])],
            })

        issues.sort(key=lambda x: (
            SEVERITY_WEIGHT[x["level"]], x["code"], x["sheet"], x["cell"]
        ))
        counts_by_level = {lv: 0 for lv in AUDIT_LEVELS}
        counts_by_code: dict[str, int] = {}
        for it in issues:
            counts_by_level[it["level"]] += 1
            counts_by_code[it["code"]] = counts_by_code.get(it["code"], 0) + 1
        return {
            "issues": issues,
            "total": len(issues),
            "counts_by_level": counts_by_level,
            "counts_by_code": counts_by_code,
            "fanout_threshold": fanout_threshold,
            "filters": {
                "levels": list(levels),
                "sheet": sheet,
            },
        }

    # -------------------------------------------------------------- 汇总
    def summary(self) -> dict:
        return {
            "sheets": [
                {
                    "name": s.name,
                    "max_row": s.max_row,
                    "max_col": s.max_col,
                    "formula_count": s.formula_count,
                }
                for s in self.sheets
            ],
            "formula_count": len(self.formulas),
            "edge_count": self.edge_count,
            "cycle_count": len(self.cycles),
            "defined_name_count": len(self.defined_names),
            "cycles": [
                [display_id(*split_node(n)) for n in cyc] for cyc in self.cycles
            ],
        }

    def defined_names_export(self) -> dict:
        """名称管理器的完整解析结果：定义清单 + 公式中的使用与解析状态。"""
        sheet_names = {s.name for s in self.sheets}

        def ref_texts(dn: DefinedNameInfo) -> list[dict]:
            out = []
            for r in dn.refs:
                out.append({
                    "ref": r.text,
                    "sheet": r.sheet,
                    "sheet_exists": r.sheet in sheet_names,
                    "kind": r.kind,
                })
            return out

        return {
            "definitions": [
                {
                    "name": dn.name,
                    "scope": dn.scope,
                    "scope_label": "工作簿级" if dn.scope is None else f"工作表级（{dn.scope}）",
                    "target": dn.target,
                    "status": dn.status,
                    "refs": ref_texts(dn),
                }
                for dn in sorted(self.defined_names,
                                 key=lambda d: (d.scope is None, d.scope or "", d.name))
            ],
            "uses": [
                {
                    "source": u.source,
                    "source_display": display_id(*split_node(u.source)),
                    "formula": u.formula,
                    "name": u.name,
                    "name_display": u.text,
                    "qualifier": u.qualifier,
                    "scope": u.scope,
                    "target": u.target,
                    "status": u.status,
                    "resolved": u.status == NAME_OK,
                }
                for u in self.name_uses
            ],
            "counts_by_status": self._name_status_counts(),
        }

    def _name_status_counts(self) -> dict[str, int]:
        counts = {NAME_OK: 0, NAME_CONSTANT: 0, NAME_UNDEFINED: 0,
                  NAME_BAD_TARGET: 0, NAME_MISSING_SHEET: 0}
        for u in self.name_uses:
            counts[u.status] = counts.get(u.status, 0) + 1
        return counts

    def cell_list(
        self,
        sheet: str | None = None,
        query: str = "",
        formulas_only: bool = True,
        limit: int = 1000,
        offset: int = 0,
    ) -> dict:
        nodes = set(self.formulas)
        if not formulas_only:
            nodes |= set(self.dependents)  # 被引用的非公式单元格
        rows = []
        q = query.lower()
        for n in sorted(nodes):
            s, coord = split_node(n)
            if sheet and s != sheet:
                continue
            disp = display_id(s, coord)
            if q and q not in disp.lower() and q not in (self.formulas.get(n) or "").lower():
                continue
            rows.append(
                {
                    "cell": n,
                    "sheet": s,
                    "coord": coord,
                    "display": disp,
                    "formula": self.formulas.get(n),
                    "value": self.values.get(n),
                    "in_cycle": n in self.cycle_nodes,
                    "dependent_count": len(self.dependents.get(n, ())),
                    "precedent_count": len(self.precedents.get(n, ())),
                }
            )
        return {"total": len(rows), "cells": rows[offset : offset + limit]}

    def export(self, filename: str) -> dict:
        """完整分析结果，供导出 JSON。"""
        edges = [
            {"from": src, "to": dst}
            for src, dsts in sorted(self.dependents.items())
            for dst in sorted(dsts)
        ]
        return {
            "filename": filename,
            "summary": self.summary(),
            "formulas": {
                display_id(*split_node(n)): {
                    "formula": f,
                    "value": self.values.get(n),
                    "in_cycle": n in self.cycle_nodes,
                }
                for n, f in sorted(self.formulas.items())
            },
            "edges": [
                {"from": display_id(*split_node(e["from"])), "to": display_id(*split_node(e["to"]))}
                for e in edges
            ],
            "cycles": [
                [display_id(*split_node(n)) for n in cyc] for cyc in self.cycles
            ],
            "defined_names": self.defined_names_export(),
            "audit": self.audit(evidence_limit=None),
        }
