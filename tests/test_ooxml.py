from pathlib import Path
from zipfile import ZipFile
import xml.etree.ElementTree as ET

from openpyxl import Workbook

from trino_excel_client.ooxml import MAIN_NS, RESULT_TEMPORAL_FORMAT_RANGE, patch_standard_temporal_formats
from trino_excel_client.template import client_sheet_names


def _q(name: str) -> str:
    return f"{{{MAIN_NS}}}{name}"


def test_patch_standard_temporal_formats_adds_openxml_rules(tmp_path: Path) -> None:
    path = tmp_path / "client.xlsx"
    names = client_sheet_names()
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = names.result
    workbook.create_sheet(names.schema)
    workbook.save(path)

    patch_standard_temporal_formats(path)

    with ZipFile(path, "r") as archive:
        styles = ET.fromstring(archive.read("xl/styles.xml"))
        sheet = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))

    dxfs = styles.find(_q("dxfs"))
    assert dxfs is not None
    assert int(dxfs.attrib["count"]) >= 3
    format_codes = [
        num_format.attrib["formatCode"]
        for num_format in dxfs.iter(_q("numFmt"))
        if num_format.attrib["formatCode"] in {"dd.mm.yyyy", "dd.mm.yyyy hh:mm", "hh:mm:ss"}
    ]
    assert format_codes == ["dd.mm.yyyy", "dd.mm.yyyy hh:mm", "hh:mm:ss"]

    conditional_blocks = [
        block
        for block in sheet.findall(_q("conditionalFormatting"))
        if block.attrib.get("sqref") == RESULT_TEMPORAL_FORMAT_RANGE
    ]
    assert len(conditional_blocks) == 1
    formulas = [formula.text for formula in conditional_blocks[0].iter(_q("formula"))]
    assert formulas == [
        'IFERROR(INDEX(TrinoSchemaExcelFormat,MATCH(A$5,TrinoSchemaColumnName,0))="date",FALSE)',
        (
            'IFERROR(OR('
            'INDEX(TrinoSchemaExcelFormat,MATCH(A$5,TrinoSchemaColumnName,0))="datetime",'
            'INDEX(TrinoSchemaExcelFormat,MATCH(A$5,TrinoSchemaColumnName,0))="datetimezone"'
            '),FALSE)'
        ),
        'IFERROR(INDEX(TrinoSchemaExcelFormat,MATCH(A$5,TrinoSchemaColumnName,0))="time",FALSE)',
    ]
