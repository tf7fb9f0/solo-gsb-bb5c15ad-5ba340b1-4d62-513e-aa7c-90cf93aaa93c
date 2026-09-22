"""名称管理器（定义名称）解析、作用域与审计测试。

运行：python -m pytest tests/test_names.py -v
"""
import io

from openpyxl import Workbook, load_workbook
from openpyxl.workbook.defined_name import DefinedName

from app.formula_parser import RawName, parse_formula, parse_formula_tokens
from app.graph import (
    NAME_BAD_TARGET,
    NAME_CONSTANT,
    NAME_MISSING_SHEET,
    NAME_OK,
    NAME_UNDEFINED,
    WorkbookGraph,
)


# ---------------------------------------------------------------- 解析器
def names(formula):
    return [t for t, kind in parse_formula_tokens(formula) if kind == "name"]


def refs(formula):
    # parse_formula 仅返回引用，名称加入后行为应保持不变
    return parse_formula(formula)


def test_bare_name_token():
    r = names("=税率*2")
    assert len(r) == 1
    assert r[0] == RawName(None, "税率", "税率")


def test_qualified_name_token():
    r = names("=参数!折扣+1")
    assert r[0] == RawName("参数", "折扣", "参数!折扣")


def test_quoted_qualified_name():
    r = names("='月报 2026'!名称+1")
    assert r[0].sheet == "月报 2026" and r[0].name == "名称"


def test_cell_refs_still_take_precedence_over_names():
    # 参数!B1:B2 必须解析为区域引用而不是限定名称
    assert names("=SUM(参数!B1:B2)") == []
    assert len(refs("=SUM(参数!B1:B2)")) == 1
    assert names("=SUM(数据!A:A)") == []
    assert names("=参数!B1") == []


def test_function_and_builtin_not_name():
    assert names("=LOG10(A1)") == []
    assert names("=TRUE+FALSE") == []
    assert names("=SUM(A1)") == []


def test_name_not_matched_inside_string_or_scientific():
    assert names('=IF(A1>0,"税率",B2)') == []
    assert names("=1E2+A1") == []


def test_structured_reference_not_name():
    assert names("=SUM(表1[销售额])+A1") == []
    assert [(t.col1, t.row1) for t in refs("=SUM(表1[销售额])+A1")] == [(1, 1)]


# ---------------------------------------------------------------- 依赖图
def _save_and_load(wb):
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    wbf = load_workbook(buf, data_only=False)
    buf.seek(0)
    wbv = load_workbook(buf, data_only=True)
    return WorkbookGraph.from_workbooks(wbf, wbv)


def build_name_workbook() -> WorkbookGraph:
    """参数!B2=单价输入；工作簿级与两张表各自的工作表级同名「折扣」。"""
    wb = Workbook()
    p = wb.active
    p.title = "参数"
    p["B2"] = 10
    p["B4"] = 0.9
    p["A1"] = "x"
    c = wb.create_sheet("计算")
    c["A1"] = 0.21
    c["B1"] = "=税率*2"            # 工作簿级 税率 -> 参数!$B$2
    c["B2"] = "=折扣"              # 裸名：解析为计算表局部 折扣 -> 计算!$A$1
    c["B3"] = "=参数!折扣"         # 限定名：解析为参数表局部 折扣 -> 参数!$B$4
    c["B4"] = "=计算!税率"         # 限定一个工作簿级名称（计算表无局部税率）
    c["B5"] = "=SUM(数据区)"       # 工作簿级区域名称 参数!$A$1:$B$4
    c["B6"] = "=常数+1"            # 常量名称，无边
    c["B7"] = "=坏REF+1"           # #REF! 名称，无边
    c["B8"] = "=幽灵+1"            # 指向不存在工作表，无边
    c["B9"] = "=没这名字+1"        # 未定义名称，无边

    wb.defined_names.add(DefinedName("税率", attr_text="参数!$B$2"))
    wb.defined_names.add(DefinedName("数据区", attr_text="参数!$A$1:$B$4"))
    wb.defined_names.add(DefinedName("常数", attr_text="3.14"))
    wb.defined_names.add(DefinedName("坏REF", attr_text="#REF!"))
    wb.defined_names.add(DefinedName("幽灵", attr_text="不存在!$A$1"))
    p.defined_names.add(DefinedName("折扣", attr_text="参数!$B$4"))
    c.defined_names.add(DefinedName("折扣", attr_text="计算!$A$1"))
    return _save_and_load(wb)


def _uses(g):
    return {(u.source, u.text): u for u in g.name_uses}


def test_defined_names_collected_with_scope():
    g = build_name_workbook()
    by_key = {(dn.scope, dn.name): dn for dn in g.defined_names}
    assert by_key[(None, "税率")].status == NAME_OK
    assert by_key[("参数", "折扣")].status == NAME_OK
    assert by_key[("计算", "折扣")].status == NAME_OK
    assert by_key[(None, "常数")].status == NAME_CONSTANT
    assert by_key[(None, "坏REF")].status == NAME_BAD_TARGET
    assert by_key[(None, "幽灵")].status == NAME_MISSING_SHEET
    assert g.summary()["defined_name_count"] == 7


def test_same_named_local_name_resolves_by_scope():
    g = build_name_workbook()
    u = _uses(g)
    # 裸名优先公式所在表的工作表级名称
    assert u[("计算!B2", "折扣")].scope == "计算"
    assert u[("计算!B2", "折扣")].status == NAME_OK
    # 限定名 优先被限定表的工作表级名称
    assert u[("计算!B3", "参数!折扣")].scope == "参数"
    # 限定表没有同名局部名时回退工作簿级
    assert u[("计算!B4", "计算!税率")].scope is None
    assert u[("计算!B1", "税率")].scope is None


def test_resolved_names_expand_to_edges():
    g = build_name_workbook()
    assert g.precedents["计算!B1"] == {"参数!B2"}
    assert g.precedents["计算!B2"] == {"计算!A1"}
    assert g.precedents["计算!B3"] == {"参数!B4"}
    assert g.precedents["计算!B4"] == {"参数!B2"}
    # 区域名称展开为 8 个坐标（参数 A1:B4）
    assert len(g.precedents["计算!B5"]) == 8
    # 常量名称 / 无效名称不产生边
    for node in ("计算!B6", "计算!B7", "计算!B8", "计算!B9"):
        assert node not in g.precedents


def test_name_edges_participate_in_impact():
    g = build_name_workbook()
    down = {i["cell"] for i in g.impact("参数!B2")["impacted"]}
    assert {"计算!B1", "计算!B4"} <= down
    # 区域名称把 参数!B2 的下游也连到 计算!B5
    assert "计算!B5" in down


def test_unresolved_names_have_no_impact_footprint():
    g = build_name_workbook()
    # 无效 / 未定义名称的公式不应凭空创造目标节点
    bogus_targets = {
        n for n in g.dependents
        if n.startswith("不存在!")
    }
    assert bogus_targets == set()
    # 这些公式本身没有 precedent 边，不出现在 precedents 中
    for node in ("计算!B7", "计算!B8", "计算!B9"):
        assert node not in g.precedents


def test_name_uses_record_status_and_evidence():
    g = build_name_workbook()
    u = _uses(g)
    assert u[("计算!B6", "常数")].status == NAME_CONSTANT
    bad = u[("计算!B7", "坏REF")]
    assert bad.status == NAME_BAD_TARGET and bad.target == "#REF!"
    ghost = u[("计算!B8", "幽灵")]
    assert ghost.status == NAME_MISSING_SHEET and ghost.target == "不存在!$A$1"
    undef = u[("计算!B9", "没这名字")]
    assert undef.status == NAME_UNDEFINED and undef.scope is None


# ---------------------------------------------------------------- 审计
def test_audit_name_issues():
    g = build_name_workbook()
    res = g.audit()
    by = {}
    for it in res["issues"]:
        if it["code"] in {"undefined_name", "invalid_name_target"}:
            by[(it["code"], it["cell"])] = it

    issue = by[("undefined_name", "计算!B9")]
    assert issue["level"] == "error"
    assert issue["name"] == "没这名字"
    assert issue["sheet"] == "计算"
    assert issue["evidence"][0]["from"] == "计算!B9"
    assert issue["evidence"][0]["ref"] == "没这名字"
    assert "to" not in issue["evidence"][0]

    bad = by[("invalid_name_target", "计算!B7")]
    assert bad["name_status"] == NAME_BAD_TARGET
    assert "#REF!" in bad["message"]
    assert bad["name_target"] == "#REF!"

    ghost = by[("invalid_name_target", "计算!B8")]
    assert ghost["name_status"] == NAME_MISSING_SHEET
    assert "不存在" in ghost["message"]


def test_audit_sheet_filter_for_name_issues():
    g = build_name_workbook()
    res = g.audit(sheet="计算")
    codes = {i["code"] for i in res["issues"]}
    assert "undefined_name" in codes
    assert "invalid_name_target" in codes
    # 参数表自身没有名称问题
    assert g.audit(sheet="参数")["counts_by_code"].get("undefined_name", 0) == 0


def test_constant_and_ok_names_are_not_issues():
    g = build_name_workbook()
    # 名称类问题仅应来自三处无效 / 未定义使用
    name_issue_cells = {
        i["cell"] for i in g.audit()["issues"]
        if i["code"] in {"undefined_name", "invalid_name_target"}
    }
    assert name_issue_cells == {"计算!B7", "计算!B8", "计算!B9"}


def test_duplicate_name_use_in_same_formula_reported_once():
    wb = Workbook()
    ws = wb.active
    ws.title = "S1"
    ws["A1"] = "=没这名字+没这名字"
    g = _save_and_load(wb)
    bad = [i for i in g.audit()["issues"] if i["code"] == "undefined_name"]
    assert len(bad) == 1


def test_expression_name_is_constant_not_ref():
    # 名称定义为含运算的表达式（A1+1）时不指向确定区域，不展开为边
    wb = Workbook()
    ws = wb.active
    ws.title = "S1"
    ws["A1"] = 5
    ws["B1"] = "=表达式名+1"
    wb.defined_names.add(DefinedName("表达式名", attr_text="$A$1+1"))
    g = _save_and_load(wb)
    dn = next(d for d in g.defined_names if d.name == "表达式名")
    assert dn.status == NAME_CONSTANT
    assert "S1!B1" not in g.precedents
    assert "S1!A1" not in g.dependents


def test_union_name_target_expands_all_refs():
    wb = Workbook()
    ws = wb.active
    ws.title = "S1"
    ws["A1"] = 1
    ws["C1"] = 2
    ws["B2"] = "=SUM(并集名)"
    wb.defined_names.add(DefinedName("并集名", attr_text="S1!$A$1,S1!$C$1"))
    g = _save_and_load(wb)
    assert g.precedents["S1!B2"] == {"S1!A1", "S1!C1"}


def test_external_and_empty_name_targets_are_invalid():
    wb = Workbook()
    ws = wb.active
    ws.title = "S1"
    ws["A1"] = "=外部名"
    ws["A2"] = "=空名"
    wb.defined_names.add(
        DefinedName("外部名", attr_text="[其他.xlsx]数据!$B$1"))
    empty = DefinedName("空名")
    empty.attr_text = None
    wb.defined_names.add(empty)
    g = _save_and_load(wb)
    by = {dn.name: dn for dn in g.defined_names}
    assert by["外部名"].status == NAME_BAD_TARGET
    assert by["空名"].status == NAME_BAD_TARGET
    codes = {(i["code"], i["cell"]) for i in g.audit()["issues"]}
    assert ("invalid_name_target", "S1!A1") in codes
    assert ("invalid_name_target", "S1!A2") in codes


def test_whole_column_name_respects_range_cap_and_used_bounds():
    wb = Workbook()
    ws = wb.active
    ws.title = "数据"
    for i in range(1, 6):
        ws.cell(row=i, column=2, value=i)
    c = wb.create_sheet("计算")
    c["B1"] = "=整列名"
    wb.defined_names.add(DefinedName("整列名", attr_text="数据!$B:$B"))
    g = _save_and_load(wb)
    # 整列名称按目标表已用范围截断为 5 个单元格
    assert g.precedents["计算!B1"] == {f"数据!B{i}" for i in range(1, 6)}


def test_qualified_undefined_name_message():
    wb = Workbook()
    ws = wb.active
    ws.title = "S1"
    ws["A1"] = "=其他表!某名"
    wb.create_sheet("其他表")
    g = _save_and_load(wb)
    issue = next(i for i in g.audit()["issues"] if i["code"] == "undefined_name")
    assert issue["name_display"] == "其他表!某名"
    assert issue["name_qualifier"] == "其他表"


def test_name_cycle_detected():
    wb = Workbook()
    ws = wb.active
    ws.title = "S1"
    ws["A1"] = "=B1+1"
    ws["B1"] = "=A名+1"           # B1 经名称回到 A1，形成 A1 <-> B1
    wb.defined_names.add(DefinedName("A名", attr_text="S1!$A$1"))
    g = _save_and_load(wb)
    flat = {n for cyc in g.cycles for n in cyc}
    assert {"S1!A1", "S1!B1"} <= flat


# ---------------------------------------------------------------- 导出
def test_export_contains_names_and_status():
    g = build_name_workbook()
    data = g.export("t.xlsx")
    assert "defined_names" in data
    defs = {(d["scope"], d["name"]): d for d in data["defined_names"]["definitions"]}
    assert defs[(None, "税率")]["status"] == NAME_OK
    assert defs[(None, "税率")]["refs"][0]["sheet"] == "参数"
    assert defs[(None, "数据区")]["refs"][0]["kind"] == "range"
    uses = {(u["source"], u["name"]): u for u in data["defined_names"]["uses"]}
    assert uses[("计算!B9", "没这名字")]["resolved"] is False
    assert uses[("计算!B1", "税率")]["resolved"] is True
    counts = data["defined_names"]["counts_by_status"]
    assert counts[NAME_OK] == 5
    assert counts[NAME_UNDEFINED] == 1
    assert counts[NAME_BAD_TARGET] == 1
    assert counts[NAME_MISSING_SHEET] == 1
    # 名称问题也出现在完整审计段
    audit_codes = {i["code"] for i in data["audit"]["issues"]}
    assert {"undefined_name", "invalid_name_target"} <= audit_codes
