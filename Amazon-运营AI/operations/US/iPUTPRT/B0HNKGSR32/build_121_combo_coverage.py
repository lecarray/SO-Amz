"""Count all 11×11 core/attribute combinations with the course's WPS preview.

The search is a case-insensitive, all-token substring search over the uploaded
full keyword table. It is not Amazon Ads' actual broad-match behavior.
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
STEM = "2026-11-11-全词表121组合广泛覆盖统计"
EXPECTED_CORES = [
    "receipt printer",
    "pos printer",
    "impresora de recibos",
    "ticket printer",
    "kitchen printer",
    "receipt maker",
    "receipt machine",
    "bill printer",
    "impresora de facturas",
    "impresora termica",
    "pos receipt printer",
]
EXPECTED_ATTRIBUTES = [
    "thermal",
    "bluetooth",
    "usb",
    "80mm",
    "ipad",
    "pos",
    "ethernet",
    "para negocios",
    "wireless",
    "bluethoot",
    "small business",
]


def normalize(value: object) -> str:
    return re.sub(r"\s+", " ", str(value).replace("\u00a0", " ").strip()).lower()


def matches(keyword: str, target: str) -> bool:
    text = normalize(keyword)
    return all(part in text for part in normalize(target).split())


def style_sheet(sheet, widths: list[int]) -> None:
    sheet.sheet_view.showGridLines = False
    sheet.freeze_panes = "B2"
    for cell in sheet[1]:
        cell.fill = PatternFill("solid", fgColor="17365D")
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    sheet.row_dimensions[1].height = 30
    for index, width in enumerate(widths, 1):
        sheet.column_dimensions[get_column_letter(index)].width = width
    for row in sheet.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)


def main() -> None:
    profile = json.loads((BASE / "source-profile.json").read_text())
    source_path = BASE / "raw" / "IPT811准词表.xlsx"
    assert hashlib.sha256(source_path.read_bytes()).hexdigest() == profile["sha256"]
    parsed = json.loads((BASE / "analysis-input.json").read_text())
    full_rows = parsed["全词表"][1:]
    assert len(full_rows) == 1830
    assert len({normalize(row["v"][0]) for row in full_rows}) == 1830

    roots = parsed["词根表"][1:]
    cores = [normalize(row["v"][1]) for row in roots]
    attributes = [normalize(row["v"][2]) for row in roots]
    assert cores == EXPECTED_CORES
    assert attributes == EXPECTED_ATTRIBUTES
    assert len(cores) == len(attributes) == 11

    counts: dict[tuple[str, str], int] = {}
    summary: list[list[object]] = []
    details: list[list[object]] = []
    union_rows: set[int] = set()
    for core in cores:
        for attribute in attributes:
            target = f"{attribute} {core}"
            hits = [row for row in full_rows if matches(str(row["v"][0]), target)]
            union_rows.update(int(row["row"]) for row in hits)
            volumes = [int(row["v"][4] or 0) for row in hits]
            redundant = all(part in core.split() for part in attribute.split())
            substring_risks = 0
            for row, volume in zip(hits, volumes):
                keyword = str(row["v"][0])
                warning = ""
                if "pos" in target.split() and "pos" in keyword.lower() and not re.search(
                    r"\bpos\b", keyword.lower()
                ):
                    warning = "pos只在词内子串命中，例如postage"
                    substring_risks += 1
                details.append(
                    [
                        core,
                        attribute,
                        target,
                        row["row"],
                        keyword,
                        volume,
                        row["v"][1],
                        warning,
                    ]
                )
            counts[(core, attribute)] = len(hits)
            summary.append(
                [
                    core,
                    attribute,
                    target,
                    len(hits),
                    sum(volume > 0 for volume in volumes),
                    sum(volume == 0 for volume in volumes),
                    sum(volumes),
                    "属性词已包含在核心词内，未缩小范围" if redundant else "",
                    substring_risks,
                ]
            )
    assert len(summary) == 121

    summary_headers = [
        "核心词（词根表）",
        "属性词（词根表）",
        "拟投放词：属性＋核心",
        "全词表覆盖词数",
        "非零量词数",
        "零量词数",
        "表内搜索量合计（周期未核）",
        "组合重复提示",
        "pos子串误命中词数",
    ]
    csv_path = BASE / f"{STEM}.csv"
    with csv_path.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.writer(file)
        writer.writerow(summary_headers)
        writer.writerows(summary)

    workbook = Workbook()
    matrix = workbook.active
    matrix.title = "11乘11覆盖矩阵"
    matrix.append(["核心词 / 属性词", *attributes])
    for core in cores:
        matrix.append([core, *(counts[(core, attribute)] for attribute in attributes)])
    style_sheet(matrix, [28, *([17] * 11)])
    for row in matrix.iter_rows(min_row=2, min_col=2):
        for cell in row:
            if cell.value == 0:
                cell.fill = PatternFill("solid", fgColor="F0F0F0")
            elif cell.value >= 10:
                cell.fill = PatternFill("solid", fgColor="E5F2E8")

    stats = workbook.create_sheet("121组合统计")
    stats.append(summary_headers)
    for line in summary:
        stats.append(line)
    style_sheet(stats, [28, 18, 42, 20, 18, 17, 27, 43, 23])
    stats.auto_filter.ref = stats.dimensions

    detail = workbook.create_sheet("命中明细")
    detail.append(
        [
            "核心词",
            "属性词",
            "拟投放词",
            "全词表源行",
            "实际命中关键词",
            "表内搜索量",
            "竞品表现得分",
            "子串误命中提示",
        ]
    )
    for line in details:
        detail.append(line)
    style_sheet(detail, [28, 18, 42, 17, 55, 18, 18, 46])
    detail.auto_filter.ref = detail.dimensions

    guide = workbook.create_sheet("统计口径")
    for line in [
        ("来源", "用户上传IPT811准词表.xlsx中的全词表与词根表，原件保存在raw/"),
        ("原件SHA256", profile["sha256"]),
        ("组合数量", "词根表11个核心词×11个属性词＝121个组合，均按属性＋核心组成投放词"),
        ("覆盖词数定义", "全词表1830个不同关键词中，每个关键词只要同时包含组合的所有空格分词，就对该组合计数1"),
        ("搜索规则", "忽略大小写、清理多余空格；每个词按字面子串匹配，近似课程WPS‘同时包含所有关键词’"),
        ("未扣除", "尚未按得分、业务相关性或用户拟定否词过滤；含零量词"),
        ("不可直接加总", "同一全词表关键词可以被多个组合命中；121行覆盖词数与搜索量不去重"),
        ("Amazon规则", "表格筛选不是亚马逊广告真实广泛匹配；实际近义扩展、资格、出价和否定需看后台搜索词报告"),
        ("子串风险", "pos可命中postage；80mm不会自动匹配80 mm；单复数/拼写差异需逐词核对"),
        ("同义/重复", "如pos printer＋pos，属性pos已在核心内；可能与不加属性的投放词得到相同范围"),
        ("第三方量", "表内搜索量的统计周期未核，不等于广告曝光、独立买家或订单"),
        ("市场与时间", "美国站B0HNKGSR32；2026-11-11 Asia/Shanghai"),
    ]:
        guide.append(line)
    style_sheet(guide, [22, 103])

    xlsx_path = BASE / f"{STEM}.xlsx"
    workbook.save(xlsx_path)

    md_lines = [
        "# 全词表121组合广泛覆盖统计（2026-11-11）",
        "",
        "按用户词根表的**11个核心词×11个属性词**逐一组成`属性词＋核心词`，在全词表1830个关键词里用课程WPS“同时包含所有关键词”方法预判。以下为**覆盖词数**；尚未按得分、业务相关性或否定词过滤，也不是Amazon广告真实触发量。",
        "",
        "|核心词 / 属性词|" + "|".join(attributes) + "|",
        "|---|" + "|".join(["---:"] * len(attributes)) + "|",
    ]
    for core in cores:
        md_lines.append("|" + core + "|" + "|".join(str(counts[(core, a)]) for a in attributes) + "|")
    md_lines += [
        "",
        f"121个组合各自的覆盖词数合计为{sum(int(row[3]) for row in summary)}，但互有重叠，去重后只涉及{len(union_rows)}个全词表关键词，不能把命中数之和当独立搜索词。每个组合的表内搜索量、非零量词数、零量词数，以及每条命中的源行和关键词，见同名XLSX/CSV。",
        "",
        "特别注意：`pos printer＋pos`与`pos receipt printer＋pos`的属性词重复，未缩小核心范围；`pos`子串可能命中`postage`；`80mm`不自动命中`80 mm`；用户原词`bluethoot`按原拼写统计，没有悄悄改成`bluetooth`。覆盖0仅代表本次全词表导出没有字面命中。",
        "",
        "课程来源：[R08 P0026](../../../../knowledge/sources/R08.md#p0026)（同时包含所有关键词）；[R08 P0046](../../../../knowledge/sources/R08.md#p0046)（词边界误命中须复核）。",
        "",
    ]
    md_path = BASE / f"{STEM}.md"
    md_path.write_text("\n".join(md_lines), encoding="utf-8")

    check = load_workbook(xlsx_path, read_only=True, data_only=True)
    assert check["11乘11覆盖矩阵"].max_row == 12
    assert check["121组合统计"].max_row == 122
    assert check["命中明细"].max_row == len(details) + 1
    check.close()
    print(f"combinations={len(summary)} total_matches={len(details)} unique_keywords={len(union_rows)}")
    print(xlsx_path)
    print(csv_path)
    print(md_path)


if __name__ == "__main__":
    main()
