import math
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from pathlib import Path

XLSX_PATH = Path("je_samples.xlsx")
OUTPUT_DIR = Path("output")
REPORT_PATH = OUTPUT_DIR / "benford_report.md"
OBS_VS_EXP_SVG = OUTPUT_DIR / "first_digit_observed_vs_expected.svg"
DEVIATION_SVG = OUTPUT_DIR / "first_digit_deviation.svg"

NS = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


@dataclass
class BenfordResult:
    counts: dict
    total: int
    observed_pct: dict
    expected_pct: dict
    mad: float
    chi_square: float


def _load_shared_strings(zip_file: zipfile.ZipFile) -> list[str]:
    shared = ET.fromstring(zip_file.read("xl/sharedStrings.xml"))
    strings = []
    for si in shared.findall("a:si", NS):
        text_node = si.find("a:t", NS)
        if text_node is not None:
            strings.append(text_node.text or "")
        else:
            runs = si.findall("a:r/a:t", NS)
            strings.append("".join(run.text or "" for run in runs))
    return strings


def _column_headers(sheet_root: ET.Element, strings: list[str]) -> dict[str, str]:
    headers = {}
    row1 = sheet_root.find(".//a:sheetData/a:row", NS)
    if row1 is None:
        return headers
    for cell in row1.findall("a:c", NS):
        cell_ref = cell.attrib.get("r", "")
        col = "".join(ch for ch in cell_ref if ch.isalpha())
        value_node = cell.find("a:v", NS)
        if value_node is None:
            continue
        if cell.attrib.get("t") == "s":
            headers[col] = strings[int(value_node.text)]
        else:
            headers[col] = value_node.text or ""
    return headers


def _cell_value(cell: ET.Element, strings: list[str]) -> str | None:
    value_node = cell.find("a:v", NS)
    if value_node is None:
        return None
    if cell.attrib.get("t") == "s":
        return strings[int(value_node.text)]
    return value_node.text


def _first_digit(value: float) -> int | None:
    if value == 0:
        return None
    value = abs(value)
    text = f"{value:.12g}"
    text = text.lstrip("0").lstrip(".")
    for char in text:
        if char.isdigit() and char != "0":
            return int(char)
    return None


def _benford_expected() -> dict[int, float]:
    return {digit: math.log10(1 + 1 / digit) for digit in range(1, 10)}


def analyze_first_digits(values: list[float]) -> BenfordResult:
    counts = {digit: 0 for digit in range(1, 10)}
    total = 0
    for value in values:
        digit = _first_digit(value)
        if digit is None:
            continue
        counts[digit] += 1
        total += 1
    expected_pct = _benford_expected()
    observed_pct = {
        digit: (counts[digit] / total if total else 0.0) for digit in counts
    }
    mad = sum(abs(observed_pct[d] - expected_pct[d]) for d in counts) / 9
    chi_square = sum(
        ((counts[d] - total * expected_pct[d]) ** 2) / (total * expected_pct[d])
        for d in counts
        if total * expected_pct[d] > 0
    )
    return BenfordResult(
        counts=counts,
        total=total,
        observed_pct=observed_pct,
        expected_pct=expected_pct,
        mad=mad,
        chi_square=chi_square,
    )


def _read_amounts() -> list[float]:
    with zipfile.ZipFile(XLSX_PATH) as zip_file:
        strings = _load_shared_strings(zip_file)
        sheet_root = ET.fromstring(zip_file.read("xl/worksheets/sheet1.xml"))
        headers = _column_headers(sheet_root, strings)
        column_for_abs = None
        column_for_amount = None
        for col, header in headers.items():
            if header == "AbsoluteAmount":
                column_for_abs = col
            elif header == "Amount":
                column_for_amount = col
        if column_for_abs is None and column_for_amount is None:
            raise ValueError("No amount column found in spreadsheet")

        values = []
        for row in sheet_root.findall(".//a:sheetData/a:row", NS)[1:]:
            cells = {cell.attrib.get("r", ""): cell for cell in row.findall("a:c", NS)}
            amount = None
            if column_for_abs:
                cell = cells.get(f"{column_for_abs}{row.attrib.get('r')}")
                if cell is not None:
                    raw = _cell_value(cell, strings)
                    if raw not in (None, ""):
                        amount = float(raw)
            if amount is None and column_for_amount:
                cell = cells.get(f"{column_for_amount}{row.attrib.get('r')}")
                if cell is not None:
                    raw = _cell_value(cell, strings)
                    if raw not in (None, ""):
                        amount = float(raw)
            if amount is None:
                continue
            if amount == 0:
                continue
            values.append(amount)
        return values


def _svg_bar_chart(
    path: Path,
    title: str,
    series: dict[str, dict[int, float]],
    y_label: str,
    y_max: float,
):
    width, height = 900, 500
    margin = 60
    chart_width = width - 2 * margin
    chart_height = height - 2 * margin

    digits = list(range(1, 10))
    group_width = chart_width / len(digits)
    bar_width = group_width / (len(series) + 1)

    def y_pos(value: float) -> float:
        return height - margin - (value / y_max) * chart_height

    lines = [
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{width}" height="{height}">',
        f'<rect width="100%" height="100%" fill="#ffffff"/>',
        f'<text x="{width/2}" y="30" font-size="18" text-anchor="middle">{title}</text>',
        f'<line x1="{margin}" y1="{height-margin}" x2="{width-margin}" y2="{height-margin}" stroke="#333"/>',
        f'<line x1="{margin}" y1="{margin}" x2="{margin}" y2="{height-margin}" stroke="#333"/>',
    ]

    colors = ["#4c78a8", "#f58518", "#54a24b"]
    series_items = list(series.items())

    for idx, digit in enumerate(digits):
        x_base = margin + idx * group_width
        for s_idx, (name, values) in enumerate(series_items):
            value = values[digit]
            x = x_base + (s_idx + 0.5) * bar_width
            y = y_pos(value)
            bar_height = height - margin - y
            lines.append(
                f'<rect x="{x:.2f}" y="{y:.2f}" width="{bar_width*0.8:.2f}" height="{bar_height:.2f}" fill="{colors[s_idx % len(colors)]}"/>'
            )
        lines.append(
            f'<text x="{x_base + group_width/2}" y="{height-margin+20}" font-size="12" text-anchor="middle">{digit}</text>'
        )

    for i in range(6):
        val = y_max * i / 5
        y = y_pos(val)
        lines.append(
            f'<line x1="{margin}" y1="{y:.2f}" x2="{width-margin}" y2="{y:.2f}" stroke="#ddd"/>'
        )
        lines.append(
            f'<text x="{margin-10}" y="{y+4:.2f}" font-size="10" text-anchor="end">{val:.1f}%</text>'
        )

    lines.append(
        f'<text x="{margin}" y="{margin-20}" font-size="12">{y_label}</text>'
    )

    legend_x = width - margin - 150
    legend_y = margin
    for idx, (name, _) in enumerate(series_items):
        y = legend_y + idx * 20
        lines.append(
            f'<rect x="{legend_x}" y="{y-10}" width="12" height="12" fill="{colors[idx % len(colors)]}"/>'
        )
        lines.append(
            f'<text x="{legend_x+18}" y="{y}" font-size="12" alignment-baseline="middle">{name}</text>'
        )

    lines.append("</svg>")
    path.write_text("\n".join(lines))


def _svg_deviation_chart(path: Path, title: str, deviations: dict[int, float]):
    width, height = 900, 500
    margin = 60
    chart_width = width - 2 * margin
    chart_height = height - 2 * margin

    digits = list(range(1, 10))
    max_dev = max(abs(v) for v in deviations.values())
    y_max = max_dev * 1.2 if max_dev else 0.05

    def y_pos(value: float) -> float:
        return height - margin - ((value + y_max) / (2 * y_max)) * chart_height

    lines = [
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{width}" height="{height}">',
        f'<rect width="100%" height="100%" fill="#ffffff"/>',
        f'<text x="{width/2}" y="30" font-size="18" text-anchor="middle">{title}</text>',
        f'<line x1="{margin}" y1="{height-margin}" x2="{width-margin}" y2="{height-margin}" stroke="#333"/>',
        f'<line x1="{margin}" y1="{margin}" x2="{margin}" y2="{height-margin}" stroke="#333"/>',
    ]

    group_width = chart_width / len(digits)
    bar_width = group_width * 0.6

    for idx, digit in enumerate(digits):
        dev = deviations[digit]
        x = margin + idx * group_width + (group_width - bar_width) / 2
        y = y_pos(max(dev, 0))
        bar_height = abs(dev) / (2 * y_max) * chart_height
        fill = "#54a24b" if dev >= 0 else "#e45756"
        lines.append(
            f'<rect x="{x:.2f}" y="{y:.2f}" width="{bar_width:.2f}" height="{bar_height:.2f}" fill="{fill}"/>'
        )
        lines.append(
            f'<text x="{x + bar_width/2:.2f}" y="{height-margin+20}" font-size="12" text-anchor="middle">{digit}</text>'
        )

    for i in range(5):
        val = y_max * (i - 2) / 2
        y = y_pos(val)
        lines.append(
            f'<line x1="{margin}" y1="{y:.2f}" x2="{width-margin}" y2="{y:.2f}" stroke="#ddd"/>'
        )
        lines.append(
            f'<text x="{margin-10}" y="{y+4:.2f}" font-size="10" text-anchor="end">{val:.1f}%</text>'
        )

    lines.append(
        f'<text x="{margin}" y="{margin-20}" font-size="12">Observed - Expected</text>'
    )
    lines.append("</svg>")
    path.write_text("\n".join(lines))


def _write_report(result: BenfordResult):
    lines = [
        "# Benford Analysis Report",
        "",
        f"Total non-zero amounts analyzed: **{result.total:,}**",
        "",
        "## First-digit distribution",
        "",
        "![Observed vs Expected](first_digit_observed_vs_expected.svg)",
        "",
        "![Observed - Expected Deviation](first_digit_deviation.svg)",
        "",
        "| Digit | Observed Count | Observed % | Expected % | Deviation (Obs - Exp) |",
        "| --- | --- | --- | --- | --- |",
    ]
    for digit in range(1, 10):
        obs_pct = result.observed_pct[digit] * 100
        exp_pct = result.expected_pct[digit] * 100
        dev = obs_pct - exp_pct
        lines.append(
            f"| {digit} | {result.counts[digit]:,} | {obs_pct:.2f}% | {exp_pct:.2f}% | {dev:+.2f}% |"
        )
    lines.extend(
        [
            "",
            f"Mean absolute deviation (MAD): **{result.mad:.4f}**",
            f"Chi-square statistic: **{result.chi_square:.2f}**",
            "",
            "MAD guidance (Nigrini):\n",
            "- 0.000–0.006: Close conformity\n",
            "- 0.006–0.012: Acceptable conformity\n",
            "- 0.012–0.015: Marginally acceptable\n",
            "- >0.015: Nonconformity",
        ]
    )
    REPORT_PATH.write_text("\n".join(lines))


def main():
    OUTPUT_DIR.mkdir(exist_ok=True)
    values = _read_amounts()
    result = analyze_first_digits(values)

    _write_report(result)

    observed_pct = {k: v * 100 for k, v in result.observed_pct.items()}
    expected_pct = {k: v * 100 for k, v in result.expected_pct.items()}
    _svg_bar_chart(
        OBS_VS_EXP_SVG,
        "First-Digit Distribution (Observed vs Expected)",
        {"Observed": observed_pct, "Expected": expected_pct},
        "Percent of totals",
        max(max(observed_pct.values()), max(expected_pct.values())) * 1.2,
    )

    deviations = {k: observed_pct[k] - expected_pct[k] for k in observed_pct}
    _svg_deviation_chart(
        DEVIATION_SVG,
        "Deviation from Benford (Observed - Expected)",
        deviations,
    )


if __name__ == "__main__":
    main()
