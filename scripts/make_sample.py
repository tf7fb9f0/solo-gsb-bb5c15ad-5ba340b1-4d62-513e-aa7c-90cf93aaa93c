"""生成示例工作簿 samples/sample_impact.xlsx。

覆盖场景：
  - 跨工作表引用（=参数!B1*参数!B2）
  - 带引号/空格的工作表名（='月报 2026'!...）
  - 矩形区域引用（=SUM(计算!B1:B4)）
  - 整列引用（=SUM(参数!B:B)）
  - 名称管理器：工作簿级 / 工作表级定义名称、同名名称按作用域解析、
    区域名称展开、常量名称、损坏名称（#REF!）、指向不存在工作表的名称、
    公式中使用未定义名称
  - 多级依赖链（参数 -> 计算 -> 汇总 -> 月报）
  - 循环引用（循环!A1 <-> 循环!B1，以及 循环!C1 自环）
  - 依赖健康审计（审计示例：缺失工作表 / 越界引用 / 空引用 / 高扇出 / 孤立公式）
"""
from pathlib import Path

from openpyxl import Workbook
from openpyxl.workbook.defined_name import DefinedName

OUT = Path(__file__).resolve().parent.parent / "samples" / "sample_impact.xlsx"


def main() -> None:
    wb = Workbook()

    # ---- 参数表：纯输入 ----
    ws = wb.active
    ws.title = "参数"
    rows = [
        ("项目", "值"),
        ("单价", 10),
        ("数量", 5),
        ("折扣率", 0.9),
        ("税率", 0.13),
        ("月数", 12),
    ]
    for r in rows:
        ws.append(r)

    # ---- 计算表：跨表引用 + 链式依赖 ----
    ws = wb.create_sheet("计算")
    ws.append(["项目", "值"])
    ws["A2"] = "小计"
    ws["B2"] = "=参数!B2*参数!B3"          # 单价 * 数量
    ws["A3"] = "折后"
    ws["B3"] = "=B2*参数!B4"               # 小计 * 折扣率
    ws["A4"] = "税额"
    ws["B4"] = "=B3*参数!B5"               # 折后 * 税率
    ws["A5"] = "月合计"
    ws["B5"] = "=B3+B4"                    # 折后 + 税额
    ws["A6"] = "年合计"
    ws["B6"] = "=B5*参数!B6"               # 月合计 * 月数

    # ---- 汇总表：跨表矩形区域引用 ----
    ws = wb.create_sheet("汇总")
    ws.append(["指标", "值"])
    ws["A2"] = "年合计(跨表)"
    ws["B2"] = "=计算!B6"
    ws["A3"] = "季度合计"
    ws["B3"] = "=计算!B6/4"
    ws["A4"] = "计算区求和(区域)"
    ws["B4"] = "=SUM(计算!B2:B5)"
    ws["A5"] = "参数区求和(区域)"
    ws["B5"] = "=SUM(参数!B2:B6)"
    ws["A6"] = "双区合计"
    ws["B6"] = "=SUM(计算!B2:B5)+SUM(参数!B2:B5)"
    ws["A7"] = "整列引用"
    ws["B7"] = "=SUM(参数!B:B)"

    # ---- 月报 2026：带空格的工作表名，公式内需引号 ----
    ws = wb.create_sheet("月报 2026")
    ws.append(["月份", "月均"])
    ws["A2"] = "月均(来自年合计)"
    ws["B2"] = "=计算!B6/12"
    ws["A3"] = "月均(来自汇总)"
    ws["B3"] = "=汇总!B2/12"
    ws["A4"] = "校验"
    ws["B4"] = "=B2-B3"

    # ---- 循环表：故意制造循环引用，用于演示检测 ----
    ws = wb.create_sheet("循环")
    ws["A1"] = "=B1+1"
    ws["B1"] = "=A1+1"
    ws["C1"] = "=C1*2"                     # 自环
    ws["D1"] = "=A1+10"                    # 受循环影响的下游单元格

    # ---- 审计数据表：200 行输入，供高扇出公式引用 ----
    ws = wb.create_sheet("数据")
    ws["A1"] = "序号"
    ws["B1"] = "值"
    for i in range(1, 201):
        ws.cell(row=i + 1, column=1, value=i)
        ws.cell(row=i + 1, column=2, value=i * 1.5)

    # ---- 名称示例表：名称管理器（工作簿级 / 工作表级定义名称）----
    wsn = wb.create_sheet("名称")
    wsn["A1"] = 0.21                        # 名称!A1：工作表级「折扣」指向本表
    ws_name = wsn
    wsn["A3"] = "场景"
    wsn["B3"] = "公式"
    wsn["A4"] = "工作簿级税率"
    wsn["B4"] = "=税率*100"                 # 税率（工作簿级）-> 参数!B5
    wsn["A5"] = "本表工作表级折扣"
    wsn["B5"] = "=折扣*100"                 # 折扣（工作表级，名称表）-> 名称!A1
    wsn["A6"] = "限定参数表同名折扣"
    wsn["B6"] = "=参数!折扣*100"            # 参数!折扣（工作表级）-> 参数!B4
    wsn["A7"] = "区域名称求和"
    wsn["B7"] = "=SUM(参数表_单价区)"       # 区域名称 -> 参数!B2:B3
    wsn["A8"] = "常量名称"
    wsn["B8"] = "=常量_PI*2"                # 常量名称 3.14，无依赖边
    wsn["A9"] = "损坏名称(#REF!)"
    wsn["B9"] = "=坏名_引用错误+1"          # error: invalid_name_target（#REF!）
    wsn["A10"] = "名称指向缺失工作表"
    wsn["B10"] = "=坏名_幽灵表+1"           # error: invalid_name_target（工作表不存在）
    wsn["A11"] = "未定义名称"
    wsn["B11"] = "=从未定义的名称+1"        # error: undefined_name

    # 工作簿级定义名称
    wb.defined_names.add(DefinedName("税率", attr_text="参数!$B$5"))
    wb.defined_names.add(DefinedName("参数表_单价区", attr_text="参数!$B$2:$B$3"))
    wb.defined_names.add(DefinedName("常量_PI", attr_text="3.14"))
    wb.defined_names.add(DefinedName("坏名_引用错误", attr_text="#REF!"))
    wb.defined_names.add(DefinedName("坏名_幽灵表", attr_text="不存在的表!$A$1"))
    # 工作表级定义名称（两个同名「折扣」，作用域不同）
    ws_name.defined_names.add(DefinedName("折扣", attr_text="名称!$A$1"))
    wb["参数"].defined_names.add(DefinedName("折扣", attr_text="参数!$B$4"))

    # ---- 审计示例表：刻意覆盖全部审计问题类型 ----
    ws = wb.create_sheet("审计示例")
    ws["A1"] = "缺失工作表"
    ws["B1"] = "=不存在的表!A1+1"           # error: missing_sheet
    ws["A2"] = "越界引用"
    ws["B2"] = "=数据!B999+1"               # warning: out_of_range
    ws["A3"] = "空单元格引用"
    ws["B3"] = "=参数!B8*2"                 # warning: empty_ref（参数表已用至第 7 行，B8 留空）
    ws["A4"] = "高扇出"
    ws["B4"] = "=SUM(数据!B2:B201)"         # warning: high_fanout（200 个直接依赖）
    ws["A5"] = "孤立公式"
    ws["B5"] = "=1+1"                       # info: orphan_formula
    ws["A6"] = "正常公式"
    ws["B6"] = "=参数!B2+1"                 # 引用合法非空
    ws["A7"] = "正常公式(下游)"
    ws["B7"] = "=B6*2"                      # 无问题：引用 B6，使 B6 不是孤立公式

    # 参数表追加标签行，使 B8 落在已用范围内；B8 本身留空，供空引用审计演示
    wb["参数"]["A8"] = "（空值行）"

    OUT.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUT)
    print(f"已生成 {OUT}")


if __name__ == "__main__":
    main()
