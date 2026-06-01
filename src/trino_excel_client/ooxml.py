from __future__ import annotations

from pathlib import Path
from tempfile import NamedTemporaryFile
from zipfile import ZIP_DEFLATED, ZipFile
import xml.etree.ElementTree as ET

from .template import client_sheet_names

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
RESULT_TEMPORAL_FORMAT_RANGE = "A6:XFD1048576"

ET.register_namespace("", MAIN_NS)
ET.register_namespace("r", REL_NS)


def _q(name: str) -> str:
    return f"{{{MAIN_NS}}}{name}"


def _rel_q(name: str) -> str:
    return f"{{{PKG_REL_NS}}}{name}"


def _find_child_index(root: ET.Element, tag_names: tuple[str, ...]) -> int:
    wanted = {_q(tag_name) for tag_name in tag_names}
    for idx, child in enumerate(list(root)):
        if child.tag in wanted:
            return idx
    return len(root)


def _find_sheet_path(files: dict[str, bytes], sheet_name: str) -> str:
    workbook = ET.fromstring(files["xl/workbook.xml"])
    relationships = ET.fromstring(files["xl/_rels/workbook.xml.rels"])

    relationship_targets = {
        relationship.attrib["Id"]: relationship.attrib["Target"]
        for relationship in relationships.findall(_rel_q("Relationship"))
    }

    for sheet in workbook.find(_q("sheets")) or []:
        if sheet.attrib.get("name") != sheet_name:
            continue

        relationship_id = sheet.attrib[f"{{{REL_NS}}}id"]
        target = relationship_targets[relationship_id]
        if target.startswith("/"):
            return target.lstrip("/")
        if target.startswith("xl/"):
            return target
        return f"xl/{target}"

    raise ValueError(f"Worksheet does not exist in workbook: {sheet_name}")


def _append_temporal_dxfs(styles_xml: bytes) -> tuple[bytes, tuple[int, int, int]]:
    root = ET.fromstring(styles_xml)
    dxfs = root.find(_q("dxfs"))

    if dxfs is None:
        dxfs = ET.Element(_q("dxfs"), {"count": "0"})
        insert_at = _find_child_index(root, ("tableStyles", "colors", "extLst"))
        root.insert(insert_at, dxfs)

    first_dxf_id = len(list(dxfs))
    formats = ("dd.mm.yyyy", "dd.mm.yyyy hh:mm", "hh:mm:ss")

    for offset, format_code in enumerate(formats):
        dxf = ET.SubElement(dxfs, _q("dxf"))
        ET.SubElement(
            dxf,
            _q("numFmt"),
            {
                "numFmtId": str(50000 + first_dxf_id + offset),
                "formatCode": format_code,
            },
        )

    dxfs.attrib["count"] = str(len(list(dxfs)))
    return ET.tostring(root, encoding="utf-8", xml_declaration=True), (
        first_dxf_id,
        first_dxf_id + 1,
        first_dxf_id + 2,
    )


def _remove_existing_temporal_rules(root: ET.Element) -> None:
    for child in list(root):
        if child.tag != _q("conditionalFormatting"):
            continue
        if child.attrib.get("sqref") == RESULT_TEMPORAL_FORMAT_RANGE:
            root.remove(child)


def _next_priority(root: ET.Element) -> int:
    priorities: list[int] = []
    for rule in root.iter(_q("cfRule")):
        try:
            priorities.append(int(rule.attrib.get("priority", "0")))
        except ValueError:
            pass
    return max(priorities, default=0) + 1


def _append_temporal_rules(sheet_xml: bytes, dxf_ids: tuple[int, int, int]) -> bytes:
    root = ET.fromstring(sheet_xml)
    _remove_existing_temporal_rules(root)

    conditional_formatting = ET.Element(
        _q("conditionalFormatting"),
        {"sqref": RESULT_TEMPORAL_FORMAT_RANGE},
    )
    first_priority = _next_priority(root)
    formulas = (
        'IFERROR(INDEX(TrinoSchemaExcelFormat,MATCH(A$5,TrinoSchemaColumnName,0))="date",FALSE)',
        (
            'IFERROR(OR('
            'INDEX(TrinoSchemaExcelFormat,MATCH(A$5,TrinoSchemaColumnName,0))="datetime",'
            'INDEX(TrinoSchemaExcelFormat,MATCH(A$5,TrinoSchemaColumnName,0))="datetimezone"'
            '),FALSE)'
        ),
        'IFERROR(INDEX(TrinoSchemaExcelFormat,MATCH(A$5,TrinoSchemaColumnName,0))="time",FALSE)',
    )

    for offset, (dxf_id, formula) in enumerate(zip(dxf_ids, formulas)):
        rule = ET.SubElement(
            conditional_formatting,
            _q("cfRule"),
            {
                "type": "expression",
                "priority": str(first_priority + offset),
                "dxfId": str(dxf_id),
            },
        )
        ET.SubElement(rule, _q("formula")).text = formula

    insert_at = _find_child_index(
        root,
        (
            "dataValidations",
            "hyperlinks",
            "printOptions",
            "pageMargins",
            "pageSetup",
            "headerFooter",
            "rowBreaks",
            "colBreaks",
            "customProperties",
            "cellWatches",
            "ignoredErrors",
            "smartTags",
            "drawing",
            "legacyDrawing",
            "legacyDrawingHF",
            "picture",
            "oleObjects",
            "controls",
            "webPublishItems",
            "tableParts",
            "extLst",
        ),
    )
    root.insert(insert_at, conditional_formatting)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def patch_standard_temporal_formats(workbook_path: Path, sheet_prefix: str = "Trino") -> None:
    workbook_path = workbook_path.resolve()
    result_sheet_name = client_sheet_names(sheet_prefix).result

    with ZipFile(workbook_path, "r") as source:
        files = {name: source.read(name) for name in source.namelist()}
        file_order = source.namelist()

    sheet_path = _find_sheet_path(files, result_sheet_name)
    styles_xml, dxf_ids = _append_temporal_dxfs(files["xl/styles.xml"])
    sheet_xml = _append_temporal_rules(files[sheet_path], dxf_ids)
    files["xl/styles.xml"] = styles_xml
    files[sheet_path] = sheet_xml

    with NamedTemporaryFile(
        prefix=f".{workbook_path.stem}.",
        suffix=workbook_path.suffix,
        dir=workbook_path.parent,
        delete=False,
    ) as temp_file:
        temp_path = Path(temp_file.name)

    try:
        with ZipFile(temp_path, "w", ZIP_DEFLATED) as target:
            for name in file_order:
                target.writestr(name, files[name])
        temp_path.replace(workbook_path)
    finally:
        try:
            temp_path.unlink()
        except FileNotFoundError:
            pass
