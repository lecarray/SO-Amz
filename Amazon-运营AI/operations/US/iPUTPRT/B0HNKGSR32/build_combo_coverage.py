"""Reproduce the course's WPS 'contains every keyword' full-table preview.

Inputs are the unchanged uploaded workbook profile and its parsed, value-only copy.
This is a string search, not an Amazon Ads match prediction.
"""

from __future__ import annotations

import csv
import hashlib
import json
import re
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


BASE = Path(__file__).resolve().parent
ATTRIBUTES = ["bluetooth", "thermal", "80mm", "usb", "ethernet", "wireless"]
CORES = ["receipt printer", "pos printer", "kitchen printer"]
SUMMARY_PATH = BASE / "2026-11-11-二级词广泛覆盖统计.csv"
WORKBOOK_PATH = BASE / "2026-11-11-二级词广泛覆盖统计.xlsx"
README_PATH = BASE / "2026-11-11-二级词广泛覆盖统计.md"


def tokens(value: str) -> list[str]:
    return re.split(r"\s+", value.replace("\u00a0", " ").strip().lower())


def matches(keyword: str, target: str) -> bool:
    name = keyword.lower()
    return all(token in name for token in tokens(target))


def main() -> None:
    profile = json.loads((BASE / "source-profile.json").read_text())
    original = BASE / "raw" / "IPT811准词表.xlsx"
    assert hashlib.sha256(original.read_bytes()).hexdigest() == profile["sha256"]
    data = json.loads((BASE / "analysis-input.json").read_text())
    full_rows = data["全词表"][1:]
    assert len(full_rows) == 1830
    assert len({str(row["v"][0]).lower() for row in full_rows}) == 1830

    summary: list[list[object]] = []
    details: list[list[object]] = []
    for attribute in ATTRIBUTES:
        group = "用户示范组" if attribute in ATTRIBUTES[:3] else "新增接口属性组"
        for core in CORES:
            target = f"{attribute} {core}"
            hit_rows = [row for row in full_rows if matches(str(row["v"][0]), target)]
            volumes = [int(row["v"][4] or 0) for row in hit_rows]
            false_pos = 0
            for row, volume in zip(hit_rows, volumes):
                keyword = str(row["v"][0])
                warning = ""
                if "pos" in tokens(target) and "pos" in keyword.lower() and not re.search(
                    r"\bpos\b", keyword.lower()
                ):
                    warning = "pos仅命中词内子串（例如postage），需人工排除"
                    false_pos += 1
                details.append(
                    [
                        group,
                        target,
                        row["row"],
                        keyword,
                        volume,
                        row["v"][1],
                        warning,
                    ]
                )
            summary.append(
                [
                    group,
                    attribute,
                    core,
                    target,
                    len(hit_rows),
                    sum(volume > 0 for volume in volumes),
                    sum(volume == 0 for volume in volumes),
                    sum(volumes),
                    false_pos,
                ]
            )

    headers = [
        "组合组别",
        "属性词",
        "核心词",
        "拟投放词",
        "全词表覆盖词数",
        "非零量词数",
        "零量词数",
        "表内搜索量合计（周期未核）",
        "pos子串误命中词数",
    ]
    with SUMMARY_PATH.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.writer(file)
        writer.writerow(headers)
        writer.writerows(summary)

    wb = Workbook()
    ws = wb.active
    ws.title = "覆盖统计"
    ws.append(headers)
    for line in summary:
        ws.append(line)
    ws.freeze_panes = "E2"
    ws.auto_filter.ref = ws.dimensions

    detail_ws = wb.create_sheet("命中明细")
    detail_ws.append(
        ["组合组别", "拟投放词", "全词表源行", "表内关键词", "表内搜索量", "竞品表现得分", "字面匹配警示"]
    )
    for line in details:
        detail_ws.append(line)
    detail_ws.freeze_panes = "D2"
    detail_ws.auto_filter.ref = detail_ws.dimensions

    guide = wb.create_sheet("统计口径")
    for line in [
        ("统计对象", "用户上传IPT811准词表.xlsx的全词表全部1830个不同关键词，含零量词；未按竞品得分筛选"),
        ("原件SHA256", profile["sha256"]),
        ("搜索方式", "拟投放词按空格分词，忽略大小写；关键词文本须同时包含每个词的字面子串"),
        ("与课程的关系", "近似WPS搜索：同时包含所有关键词；不是亚马逊广泛匹配真实触发规则"),
        ("否定词", "此表尚未应用任何否定词；先保留原始覆盖，再逐词审核错意图"),
        ("搜索量", "表内第三方估值；统计周期未在源文件核实，不能当实际广告曝光或订单"),
        ("重叠", "每个组合分别计数，同一关键词可落入多个组合；不可直接将各行加总当去重总量"),
        ("子串误命中", "例如pos可命中postage；警示列只标已识别的pos词边界问题"),
        ("写法差异", "80mm不会自动命中80 mm；此表保持用户的80mm写法"),
        ("无线含义", "wireless是词库属性词，仅用于研究；不能由此推定本SKU支持Wi-Fi"),
        ("站点/产品", "美国站 B0HNKGSR32 / IPT811-UEB"),
        ("统计日期", "2026-11-11（Asia/Shanghai）"),
    ]:
        guide.append(line)

    widths = {
        ws.title: [18, 16, 22, 34, 20, 18, 16, 30, 20],
        detail_ws.title: [18, 34, 16, 52, 18, 18, 48],
        guide.title: [20, 98],
    }
    for sheet in wb:
        sheet.sheet_view.showGridLines = False
        for cell in sheet[1]:
            cell.fill = PatternFill("solid", fgColor="17365D")
            cell.font = Font(color="FFFFFF", bold=True)
            cell.alignment = Alignment(vertical="center")
        sheet.row_dimensions[1].height = 25
        for index, width in enumerate(widths[sheet.title], start=1):
            sheet.column_dimensions[get_column_letter(index)].width = width
        for row in sheet.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(vertical="top", wrap_text=True)
    for row in ws.iter_rows(min_row=2):
        if row[0].value == "新增接口属性组":
            for cell in row:
                cell.fill = PatternFill("solid", fgColor="E8F3F6")
    wb.save(WORKBOOK_PATH)

    expected = {
        "bluetooth receipt printer": 65,
        "thermal receipt printer": 77,
        "80mm receipt printer": 17,
        "usb receipt printer": 11,
        "ethernet receipt printer": 2,
        "wireless receipt printer": 14,
    }
    found = {str(line[3]): int(line[4]) for line in summary}
    for target, count in expected.items():
        assert found[target] == count
    check = load_workbook(WORKBOOK_PATH, read_only=True, data_only=True)
    assert check["覆盖统计"].max_row == len(summary) + 1
    assert check["命中明细"].max_row == len(details) + 1
    check.close()

    grid = {str(line[3]): line for line in summary}
    text_lines = [
        "# 二级词广泛覆盖统计（2026-11-11）",
        "",
        "从用户给出的三类核心词（receipt printer、pos printer、kitchen printer）出发，原示范组用Bluetooth、Thermal、80mm组合；新增一组用词根表的USB、Ethernet、Wireless组合。统计全词表1830词，未做得分筛选或否定。",
        "",
        "**口径**：按课程所用WPS‘同时包含所有关键词’近似预判：忽略大小写，投放词按空格拆开，只要全词表关键词文本同时包含所有子串就计1词。关键词可被多组合重复覆盖；这不是亚马逊实际广泛匹配、广告曝光或可获得销量。表内搜索量统计周期未核。",
        "",
        "|属性词|receipt printer|pos printer|kitchen printer|",
        "|---|---:|---:|---:|",
    ]
    for attribute in ATTRIBUTES:
        counts = [grid[f"{attribute} {core}"][4] for core in CORES]
        text_lines.append(f"|{attribute}|{counts[0]}|{counts[1]}|{counts[2]}|")
    text_lines += [
        "",
        "每格是**覆盖词数**，不是搜索量。各组合的表内搜索量、零量词数、源行和命中原词见同名XLSX/CSV。",
        "",
        "注意：`thermal pos printer`的14词/325量中，`thermal postage label printer`占307量，是`pos`命中`postage`的子串误匹配；`80mm`不会自动匹配`80 mm`；`wireless`不能证明本SKU支持Wi-Fi。零量不等于市场无需求。",
        "",
        "课程来源：[R08 P0026](../../../../knowledge/sources/R08.md#p0026)（全词表同时包含所有关键词）；[R08 P0046](../../../../knowledge/sources/R08.md#p0046)（子串误命中须核）。",
        "",
    ]
    README_PATH.write_text("\n".join(text_lines), encoding="utf-8")
    print(f"combinations={len(summary)} detail_rows={len(details)}")
    print(f"summary={SUMMARY_PATH}")
    print(f"workbook={WORKBOOK_PATH}")
    print(f"markdown={README_PATH}")


if __name__ == "__main__":
    main()
