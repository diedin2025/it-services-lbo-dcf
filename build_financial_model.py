from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import xml.sax.saxutils as saxutils


OUT_DIR = Path("/Users/ian/Documents/Playground")


def col_name(index: int) -> str:
    result = ""
    while index:
        index, rem = divmod(index - 1, 26)
        result = chr(65 + rem) + result
    return result


def cell_ref(row: int, col: int) -> str:
    return f"{col_name(col)}{row}"


def escape(text: str) -> str:
    return saxutils.escape(str(text))


@dataclass
class Cell:
    value: object | None = None
    formula: str | None = None
    style: int = 0
    cell_type: str | None = None


def number_cell(value: float | int, style: int = 0) -> Cell:
    return Cell(value=value, style=style)


def text_cell(value: str, style: int = 0) -> Cell:
    return Cell(value=value, style=style, cell_type="inlineStr")


def formula_cell(formula: str, value: float | int | str | None, style: int = 0, cell_type: str | None = None) -> Cell:
    return Cell(value=value, formula=formula, style=style, cell_type=cell_type)


def blank_cell(style: int = 0) -> Cell:
    return Cell(value=None, style=style)


def cell_xml(ref: str, cell: Cell) -> str:
    attrs = [f'r="{ref}"']
    if cell.style:
        attrs.append(f's="{cell.style}"')

    if cell.formula is not None:
        if cell.cell_type == "str":
            attrs.append('t="str"')
        elif cell.cell_type == "inlineStr":
            attrs.append('t="inlineStr"')
        xml = [f"<c {' '.join(attrs)}>"]
        xml.append(f"<f>{escape(cell.formula)}</f>")
        if cell.cell_type == "inlineStr":
            xml.append(f"<is><t>{escape(cell.value or '')}</t></is>")
        elif cell.cell_type == "str":
            xml.append(f"<v>{escape(cell.value or '')}</v>")
        elif cell.value is not None:
            xml.append(f"<v>{cell.value}</v>")
        xml.append("</c>")
        return "".join(xml)

    if cell.value is None:
        return f"<c {' '.join(attrs)}/>"

    if cell.cell_type == "inlineStr":
        attrs.append('t="inlineStr"')
        return f"<c {' '.join(attrs)}><is><t>{escape(cell.value)}</t></is></c>"

    return f"<c {' '.join(attrs)}><v>{cell.value}</v></c>"


def build_sheet_xml(cells: dict[tuple[int, int], Cell], widths: dict[int, float] | None = None) -> str:
    row_map: dict[int, list[tuple[int, Cell]]] = {}
    for (row, col), cell in cells.items():
        row_map.setdefault(row, []).append((col, cell))

    xml = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        '<sheetViews><sheetView workbookViewId="0"/></sheetViews>',
        '<sheetFormatPr defaultRowHeight="15"/>',
    ]

    if widths:
        xml.append("<cols>")
        for col_idx, width in sorted(widths.items()):
            xml.append(f'<col min="{col_idx}" max="{col_idx}" width="{width}" customWidth="1"/>')
        xml.append("</cols>")

    xml.append("<sheetData>")
    for row in sorted(row_map):
        xml.append(f'<row r="{row}">')
        for col, cell in sorted(row_map[row], key=lambda item: item[0]):
            xml.append(cell_xml(cell_ref(row, col), cell))
        xml.append("</row>")
    xml.append("</sheetData>")
    xml.append("</worksheet>")
    return "".join(xml)


def workbook_xml(sheet_names: list[str]) -> str:
    sheets = []
    for idx, name in enumerate(sheet_names, start=1):
        sheets.append(
            f'<sheet name="{escape(name)}" sheetId="{idx}" r:id="rId{idx}"/>'
        )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<fileVersion appName="Codex"/>'
        '<workbookPr/>'
        '<bookViews><workbookView activeTab="0"/></bookViews>'
        f"<sheets>{''.join(sheets)}</sheets>"
        '<calcPr calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/>'
        "</workbook>"
    )


def workbook_rels(sheet_count: int) -> str:
    rels = []
    for idx in range(1, sheet_count + 1):
        rels.append(
            f'<Relationship Id="rId{idx}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            f'Target="worksheets/sheet{idx}.xml"/>'
        )
    rels.append(
        f'<Relationship Id="rId{sheet_count + 1}" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
        'Target="styles.xml"/>'
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f"{''.join(rels)}"
        "</Relationships>"
    )


def root_rels() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="xl/workbook.xml"/>'
        '<Relationship Id="rId2" '
        'Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" '
        'Target="docProps/core.xml"/>'
        '<Relationship Id="rId3" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" '
        'Target="docProps/app.xml"/>'
        "</Relationships>"
    )


def content_types(sheet_count: int) -> str:
    overrides = [
        '<Override PartName="/xl/workbook.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
        '<Override PartName="/xl/styles.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
        '<Override PartName="/docProps/core.xml" '
        'ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
        '<Override PartName="/docProps/app.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>',
    ]
    for idx in range(1, sheet_count + 1):
        overrides.append(
            f'<Override PartName="/xl/worksheets/sheet{idx}.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        f"{''.join(overrides)}"
        "</Types>"
    )


def styles_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<fonts count="4">'
        '<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/></font>'
        '<font><b/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/></font>'
        '<font><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font>'
        '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font>'
        '</fonts>'
        '<fills count="5">'
        '<fill><patternFill patternType="none"/></fill>'
        '<fill><patternFill patternType="gray125"/></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FF000000"/><bgColor indexed="64"/></patternFill></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FFBDD7EE"/><bgColor indexed="64"/></patternFill></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FFD9EAD3"/><bgColor indexed="64"/></patternFill></fill>'
        '</fills>'
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="6">'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
        '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
        '<xf numFmtId="0" fontId="2" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>'
        '<xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1"/>'
        '<xf numFmtId="10" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
        '<xf numFmtId="4" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
        '</cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        '</styleSheet>'
    )


def core_xml() -> str:
    created = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" '
        'xmlns:dcterms="http://purl.org/dc/terms/" '
        'xmlns:dcmitype="http://purl.org/dc/dcmitype/" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        '<dc:creator>Codex</dc:creator>'
        '<cp:lastModifiedBy>Codex</cp:lastModifiedBy>'
        f'<dcterms:created xsi:type="dcterms:W3CDTF">{created}</dcterms:created>'
        f'<dcterms:modified xsi:type="dcterms:W3CDTF">{created}</dcterms:modified>'
        "</cp:coreProperties>"
    )


def app_xml(sheet_count: int) -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
        'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
        '<Application>Codex</Application>'
        f"<Sheets>{sheet_count}</Sheets>"
        "</Properties>"
    )


def make_inputs_sheet() -> tuple[dict[tuple[int, int], Cell], dict[int, float]]:
    cells: dict[tuple[int, int], Cell] = {}
    widths = {1: 28, 2: 14, 3: 12, 4: 12, 5: 12, 6: 18, 7: 14}

    cells[(1, 1)] = text_cell("Calculation", 1)
    cells[(1, 2)] = text_cell("Black", 2)
    cells[(2, 1)] = text_cell("Input", 1)
    cells[(2, 2)] = text_cell("Blue", 3)

    headers = ["Company Financials (£m)", "Dec 24A", "Dec 25B", "Dec 26F", "Dec 27F", "Dec 28F", "Dec 29F"]
    for idx, header in enumerate(headers, start=1):
        cells[(4, idx)] = text_cell(header, 1)

    revenue = [9.99, 15.7, 19.0, 22.7, 27.3, 32.7]
    ebitda = [1.034, 2.312, 2.377, 2.961, 3.762, 4.743]
    cells[(5, 1)] = text_cell("Revenue", 1)
    cells[(6, 1)] = text_cell("EBITDA", 1)
    for idx, value in enumerate(revenue, start=2):
        cells[(5, idx)] = number_cell(value, 5)
    for idx, value in enumerate(ebitda, start=2):
        cells[(6, idx)] = number_cell(value, 5)

    cells[(9, 1)] = text_cell("Current Shareholdings (Diluted)", 1)
    cells[(9, 6)] = text_cell("Other assumptions", 1)

    shareholders = [
        ("Chairman", 521805.281, 0.331),
        ("CEO", 553334.301, 0.351),
        ("Sales Director", 102469.315, 0.065),
        ("Operations Director", 95000.02474, 0.0602619585),
        ("Head of US", 0, 0),
        ("External shareholders", 165000.0427, 0.1046655067),
        ("Employees", 97999.02532, 0.06216433325),
        ("Option holders", 40843.01054, 0.02590820174),
    ]
    for row, (name, shares, pct) in enumerate(shareholders, start=10):
        cells[(row, 1)] = text_cell(name, 1)
        cells[(row, 2)] = number_cell(shares, 5)
        cells[(row, 3)] = number_cell(pct, 4)
    cells[(18, 2)] = formula_cell("SUM(B10:B17)", 1576451, 5)
    cells[(18, 3)] = formula_cell("SUM(C10:C17)", 1, 4)

    assumptions = [
        ("Tax", 0.2),
        ("CAPEX", 0.01),
        ("Working capital", 0.1),
        ("Yield on senior debt", 0.05),
        ("Deal fees", 0.015),
        ("Cash at Bank", 2.3),
    ]
    for row, (label, value) in enumerate(assumptions, start=10):
        cells[(row, 6)] = text_cell(label, 1)
        cells[(row, 7)] = number_cell(value, 4 if label != "Cash at Bank" else 5)

    return cells, widths


def make_lbo_sheet() -> tuple[dict[tuple[int, int], Cell], dict[int, float]]:
    cells: dict[tuple[int, int], Cell] = {}
    widths = {1: 34, 2: 12, 3: 12, 4: 12, 5: 12, 6: 12}

    cells[(1, 1)] = text_cell("LBO", 1)
    cells[(2, 1)] = text_cell("Acquisition Date", 1)
    cells[(2, 2)] = text_cell("2025-09-30")

    years = ["2025B", "2026F", "2027F", "2028F", "2029F"]
    for idx, year in enumerate(years, start=2):
        cells[(4, idx)] = text_cell(year, 1)

    row_labels = [
        "Revenue",
        "EBITDA",
        "EBITDA Margin",
        "Taxes",
        "CAPEX",
        "Working Capital",
        "Change in Working Capital",
        "Unlevered Free Cash Flow",
    ]
    for row, label in enumerate(row_labels, start=5):
        cells[(row, 1)] = text_cell(label, 1)

    revenue_vals = [15.7, 19.0, 22.7, 27.3, 32.7]
    ebitda_vals = [2.312, 2.377, 2.961, 3.762, 4.743]
    tax_vals = [0.4624, 0.4754, 0.5922, 0.7524, 0.9486]
    capex_vals = [0.157, 0.19, 0.227, 0.273, 0.327]
    wc_vals = [1.57, 1.9, 2.27, 2.73, 3.27]
    delta_wc_vals = [0.571, 0.33, 0.37, 0.46, 0.54]
    fcf_vals = [1.1216, 1.3816, 1.7718, 2.2766, 2.9274]

    for idx in range(5):
        col = idx + 2
        source_col = chr(ord("C") + idx)
        cells[(5, col)] = formula_cell(f"Inputs!{source_col}5", revenue_vals[idx], 5)
        cells[(6, col)] = formula_cell(f"Inputs!{source_col}6", ebitda_vals[idx], 5)
        cells[(7, col)] = formula_cell(f"{col_name(col)}6/{col_name(col)}5", ebitda_vals[idx] / revenue_vals[idx], 4)
        cells[(8, col)] = formula_cell(f"{col_name(col)}6*Inputs!G10", tax_vals[idx], 5)
        cells[(9, col)] = formula_cell(f"{col_name(col)}5*Inputs!G11", capex_vals[idx], 5)
        cells[(10, col)] = formula_cell(f"{col_name(col)}5*Inputs!G12", wc_vals[idx], 5)
        if idx == 0:
            cells[(11, col)] = formula_cell(f"{col_name(col)}10-Inputs!B5*Inputs!G12", delta_wc_vals[idx], 5)
        else:
            prev_col = col_name(col - 1)
            cells[(11, col)] = formula_cell(f"{col_name(col)}10-{prev_col}10", delta_wc_vals[idx], 5)
        cells[(12, col)] = formula_cell(
            f"{col_name(col)}6-{col_name(col)}8-{col_name(col)}9-{col_name(col)}11",
            fcf_vals[idx],
            5,
        )

    cells[(15, 1)] = text_cell("Entry Enterprise Value", 1)
    cells[(15, 2)] = text_cell("Undisclosed")
    cells[(16, 1)] = text_cell("Entry EBITDA Multiple", 1)
    cells[(16, 2)] = text_cell("Undisclosed")
    cells[(17, 1)] = text_cell("Cash at Bank", 1)
    cells[(17, 2)] = formula_cell("Inputs!G15", 2.3, 5)
    cells[(18, 1)] = text_cell("Deal Fees", 1)
    cells[(18, 2)] = text_cell("Undisclosed")
    cells[(19, 1)] = text_cell("Senior Debt at Close", 1)
    cells[(19, 2)] = text_cell("Undisclosed")
    cells[(20, 1)] = text_cell("Equity at Close", 1)
    cells[(20, 2)] = text_cell("Undisclosed")
    cells[(22, 1)] = text_cell("Notes", 1)
    cells[(23, 1)] = text_cell("Taxes are applied to EBITDA because no depreciation or amortization was provided.")
    cells[(24, 1)] = text_cell("Leverage, purchase price, and exit assumptions were not provided, so sponsor return metrics remain Undisclosed.")

    return cells, widths


def make_dcf_sheet() -> tuple[dict[tuple[int, int], Cell], dict[int, float]]:
    cells: dict[tuple[int, int], Cell] = {}
    widths = {1: 34, 2: 12, 3: 12, 4: 12, 5: 12, 6: 12}

    cells[(1, 1)] = text_cell("DCF", 1)
    years = ["2025B", "2026F", "2027F", "2028F", "2029F"]
    for idx, year in enumerate(years, start=2):
        cells[(4, idx)] = text_cell(year, 1)

    row_labels = [
        "Revenue",
        "EBITDA",
        "EBITDA Margin",
        "Taxes",
        "CAPEX",
        "Working Capital",
        "Change in Working Capital",
        "Unlevered Free Cash Flow",
        "Discount Period",
        "Discount Factor",
        "Present Value of UFCF",
    ]
    for row, label in enumerate(row_labels, start=5):
        cells[(row, 1)] = text_cell(label, 1)

    revenue_vals = [15.7, 19.0, 22.7, 27.3, 32.7]
    ebitda_vals = [2.312, 2.377, 2.961, 3.762, 4.743]
    tax_vals = [0.4624, 0.4754, 0.5922, 0.7524, 0.9486]
    capex_vals = [0.157, 0.19, 0.227, 0.273, 0.327]
    wc_vals = [1.57, 1.9, 2.27, 2.73, 3.27]
    delta_wc_vals = [0.571, 0.33, 0.37, 0.46, 0.54]
    fcf_vals = [1.1216, 1.3816, 1.7718, 2.2766, 2.9274]
    discount_periods = [0.25, 1.25, 2.25, 3.25, 4.25]

    for idx in range(5):
        col = idx + 2
        source_col = chr(ord("C") + idx)
        cells[(5, col)] = formula_cell(f"Inputs!{source_col}5", revenue_vals[idx], 5)
        cells[(6, col)] = formula_cell(f"Inputs!{source_col}6", ebitda_vals[idx], 5)
        cells[(7, col)] = formula_cell(f"{col_name(col)}6/{col_name(col)}5", ebitda_vals[idx] / revenue_vals[idx], 4)
        cells[(8, col)] = formula_cell(f"{col_name(col)}6*Inputs!G10", tax_vals[idx], 5)
        cells[(9, col)] = formula_cell(f"{col_name(col)}5*Inputs!G11", capex_vals[idx], 5)
        cells[(10, col)] = formula_cell(f"{col_name(col)}5*Inputs!G12", wc_vals[idx], 5)
        if idx == 0:
            cells[(11, col)] = formula_cell(f"{col_name(col)}10-Inputs!B5*Inputs!G12", delta_wc_vals[idx], 5)
        else:
            prev_col = col_name(col - 1)
            cells[(11, col)] = formula_cell(f"{col_name(col)}10-{prev_col}10", delta_wc_vals[idx], 5)
        cells[(12, col)] = formula_cell(
            f"{col_name(col)}6-{col_name(col)}8-{col_name(col)}9-{col_name(col)}11",
            fcf_vals[idx],
            5,
        )
        cells[(13, col)] = number_cell(discount_periods[idx], 5)
        cells[(14, col)] = text_cell("Undisclosed")
        cells[(15, col)] = text_cell("Undisclosed")

    cells[(18, 1)] = text_cell("WACC", 1)
    cells[(18, 2)] = text_cell("Undisclosed")
    cells[(19, 1)] = text_cell("Exit EBITDA Multiple", 1)
    cells[(19, 2)] = text_cell("Undisclosed")
    cells[(20, 1)] = text_cell("Terminal Value", 1)
    cells[(20, 2)] = text_cell("Undisclosed")
    cells[(21, 1)] = text_cell("DCF Enterprise Value (Exit Multiple Case)", 1)
    cells[(21, 2)] = text_cell("Undisclosed")
    cells[(23, 1)] = text_cell("Notes", 1)
    cells[(24, 1)] = text_cell("The forecast free cash flow is fully formula-driven from revenue, EBITDA, tax, CAPEX, and working-capital assumptions.")
    cells[(25, 1)] = text_cell("WACC and terminal multiple were not provided, so the discounted valuation case remains Undisclosed.")

    return cells, widths


def make_public_comps_sheet() -> tuple[dict[tuple[int, int], Cell], dict[int, float]]:
    cells: dict[tuple[int, int], Cell] = {}
    widths = {1: 26, 2: 14, 3: 14, 4: 14, 5: 14, 6: 16}

    cells[(1, 1)] = text_cell("Public Comps", 1)
    headers = ["Company", "Enterprise Value", "Revenue", "EBITDA", "EV / Revenue", "EV / EBITDA"]
    for idx, header in enumerate(headers, start=1):
        cells[(3, idx)] = text_cell(header, 1)
    for row in range(4, 8):
        cells[(row, 1)] = text_cell(f"Comparable {row - 3}")
        for col in range(2, 7):
            cells[(row, col)] = text_cell("Undisclosed")
    cells[(10, 1)] = text_cell("Median EV / EBITDA", 1)
    cells[(10, 2)] = text_cell("Undisclosed")
    cells[(12, 1)] = text_cell("No public comparable set was provided in the prompt, so trading multiples are Undisclosed.")

    return cells, widths


def make_precedents_sheet() -> tuple[dict[tuple[int, int], Cell], dict[int, float]]:
    cells: dict[tuple[int, int], Cell] = {}
    widths = {1: 26, 2: 14, 3: 14, 4: 14, 5: 14, 6: 16}

    cells[(1, 1)] = text_cell("Precedent Transactions", 1)
    headers = ["Transaction", "Enterprise Value", "Revenue", "EBITDA", "EV / Revenue", "EV / EBITDA"]
    for idx, header in enumerate(headers, start=1):
        cells[(3, idx)] = text_cell(header, 1)
    for row in range(4, 8):
        cells[(row, 1)] = text_cell(f"Transaction {row - 3}")
        for col in range(2, 7):
            cells[(row, col)] = text_cell("Undisclosed")
    cells[(10, 1)] = text_cell("Lower-Quartile EV / EBITDA", 1)
    cells[(10, 2)] = text_cell("Undisclosed")
    cells[(12, 1)] = text_cell("No precedent transaction set was provided in the prompt, so transaction multiples are Undisclosed.")

    return cells, widths


def make_summary_sheet() -> tuple[dict[tuple[int, int], Cell], dict[int, float]]:
    cells: dict[tuple[int, int], Cell] = {}
    widths = {1: 44, 2: 18}

    cells[(1, 1)] = text_cell("Valuation Summary", 1)
    rows = [
        "LBO Entry Multiple",
        "DCF Exit-Multiple Case",
        "Median Public Company EBITDA Multiple",
        "Lower-Quartile Precedent Transaction EBITDA Multiple",
        "Median of the Four Indicators Above",
    ]
    for row, label in enumerate(rows, start=3):
        cells[(row, 1)] = text_cell(label, 1)
        cells[(row, 2)] = text_cell("Undisclosed")
    cells[(10, 1)] = text_cell("The valuation summary is constrained by missing market and deal inputs in the prompt.")

    return cells, widths


def write_xlsx(path: Path, sheets: list[tuple[str, dict[tuple[int, int], Cell], dict[int, float]]]) -> None:
    sheet_names = [sheet[0] for sheet in sheets]
    with ZipFile(path, "w", ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types(len(sheets)))
        zf.writestr("_rels/.rels", root_rels())
        zf.writestr("xl/workbook.xml", workbook_xml(sheet_names))
        zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels(len(sheets)))
        zf.writestr("xl/styles.xml", styles_xml())
        zf.writestr("docProps/core.xml", core_xml())
        zf.writestr("docProps/app.xml", app_xml(len(sheets)))
        for idx, (_, cells, widths) in enumerate(sheets, start=1):
            zf.writestr(f"xl/worksheets/sheet{idx}.xml", build_sheet_xml(cells, widths))


def main() -> None:
    inputs_cells, inputs_widths = make_inputs_sheet()
    lbo_cells, lbo_widths = make_lbo_sheet()
    dcf_cells, dcf_widths = make_dcf_sheet()
    public_cells, public_widths = make_public_comps_sheet()
    precedent_cells, precedent_widths = make_precedents_sheet()
    summary_cells, summary_widths = make_summary_sheet()

    fcp_sheets = [
        ("Inputs", inputs_cells, inputs_widths),
        ("LBO", lbo_cells, lbo_widths),
        ("DCF", dcf_cells, dcf_widths),
        ("Public Comps", public_cells, public_widths),
        ("Precedent Transactions", precedent_cells, precedent_widths),
        ("Valuation Summary", summary_cells, summary_widths),
    ]
    inputs_only_sheets = [("Inputs", inputs_cells, inputs_widths)]

    write_xlsx(OUT_DIR / "FCP_eval.xlsx", fcp_sheets)
    write_xlsx(OUT_DIR / "inputs.xlsx", inputs_only_sheets)


if __name__ == "__main__":
    main()
