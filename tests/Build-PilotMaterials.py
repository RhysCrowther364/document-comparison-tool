"""Create local, shareable pilot instructions and synthetic Word fixtures."""
from pathlib import Path
import re
import zipfile
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parent
APP_VERSION = "1.0.20.0"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
ET.register_namespace("w", W)

def node(name, text=None, **attrs):
    element = ET.Element("{" + W + "}" + name)
    for key, value in attrs.items():
        element.set("{" + W + "}" + key, str(value))
    if text is not None:
        element.text = text
    return element

def paragraph(text, style=None, bold=False, colour=None):
    p = node("p")
    if style:
        props = node("pPr")
        props.append(node("pStyle", val=style))
        p.append(props)
    run = node("r")
    if bold or colour:
        props = node("rPr")
        if bold:
            props.append(node("b"))
        if colour:
            props.append(node("color", val=colour))
        run.append(props)
    value = node("t", text)
    value.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    run.append(value)
    p.append(run)
    return p

def save_docx(path, blocks):
    document = node("document")
    body = node("body")
    document.append(body)
    body.extend(blocks)
    section = node("sectPr")
    section.append(node("pgSz", w=11906, h=16838))
    section.append(node("pgMar", top=1134, right=1134, bottom=1134, left=1134))
    body.append(section)
    styles = node("styles")
    definitions = [("Normal", 21, "181818", False), ("Title", 44, "3E016F", True),
                   ("Heading1", 27, "3E016F", True), ("TableText", 18, "181818", False)]
    for name, size, colour, bold in definitions:
        style = node("style", type="paragraph", styleId=name)
        style.append(node("name", val=name))
        if name == "Normal":
            style.set("{" + W + "}default", "1")
        props = node("pPr")
        if name in ("Title", "Heading1"):
            props.append(node("keepNext"))
            props.append(node("spacing", before=180, after=120))
        else:
            props.append(node("spacing", after=100))
        style.append(props)
        props = node("rPr")
        props.append(node("rFonts", ascii="Arial", hAnsi="Arial", cs="Arial"))
        props.append(node("sz", val=size))
        props.append(node("color", val=colour))
        if bold:
            props.append(node("b"))
        style.append(props)
        styles.append(style)
    types = '''<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>'''
    rels = '''<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="%s/officeDocument" Target="word/document.xml"/></Relationships>''' % R
    document_rels = '''<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdStyles" Type="%s/styles" Target="styles.xml"/></Relationships>''' % R
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", types)
        archive.writestr("_rels/.rels", rels)
        archive.writestr("word/_rels/document.xml.rels", document_rels)
        archive.writestr("word/styles.xml", ET.tostring(styles, encoding="utf-8", xml_declaration=True))
        archive.writestr("word/document.xml", ET.tostring(document, encoding="utf-8", xml_declaration=True))

def clean(text):
    return text.replace("**", "").replace("`", "")

def table(rows):
    result = node("tbl")
    props = node("tblPr")
    props.append(node("tblW", w=9638, type="dxa"))
    borders = node("tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        borders.append(node(side, val="single", sz=4, color="DADADA"))
    props.append(borders)
    result.append(props)
    count = len(rows[0])
    widths = [1800, 7838] if count == 2 else [1200, 4000, 4438]
    grid = node("tblGrid")
    for width in widths:
        grid.append(node("gridCol", w=width))
    result.append(grid)
    for index, row in enumerate(rows):
        tr = node("tr")
        props = node("trPr")
        props.append(node("cantSplit"))
        if index == 0:
            props.append(node("tblHeader"))
        tr.append(props)
        for column, text in enumerate(row):
            cell = node("tc")
            props = node("tcPr")
            props.append(node("tcW", w=widths[column], type="dxa"))
            if index == 0:
                props.append(node("shd", fill="3E016F", val="clear"))
            cell.append(props)
            cell.append(paragraph(clean(text), "TableText", bold=index == 0,
                                  colour="FFFFFF" if index == 0 else None))
            tr.append(cell)
        result.append(tr)
    return result

def build_plan():
    lines = ROOT.joinpath("PilotTestPlan.md").read_text(encoding="utf-8").splitlines()
    blocks = []
    i = 0
    while i < len(lines):
        line = lines[i]
        if line.startswith("|"):
            rows = []
            while i < len(lines) and lines[i].startswith("|"):
                row = [cell.strip() for cell in lines[i].strip().strip("|").split("|")]
                if not all(re.fullmatch(r"[-: ]+", cell) for cell in row):
                    rows.append(row)
                i += 1
            blocks.append(table(rows))
            blocks.append(paragraph(""))
            continue
        if line.startswith("# "):
            blocks.append(paragraph(clean(line[2:]), "Title"))
        elif line.startswith("## "):
            blocks.append(paragraph(clean(line[3:]), "Heading1"))
        elif line:
            blocks.append(paragraph(clean(line)))
        i += 1
    save_docx(ROOT / "PilotTestPlan.docx", blocks)

def build_fixtures():
    directory = ROOT / "PilotFixtures"
    directory.mkdir(exist_ok=True)
    def create(name, text):
        save_docx(directory / name, [paragraph(value) for value in text])
    create("WF-four-archive.docx", ["Alpha beta gamma delta."])
    create("WF-four-identical.docx", ["Alpha beta gamma delta."])
    create("WF-five-added.docx", ["Alpha beta gamma delta epsilon."])
    create("WF-three-deleted.docx", ["Alpha beta gamma."])
    create("WF-three-archive.docx", ["Alpha beta gamma."])
    create("WF-three-unrelated.docx", ["River stone cloud."])
    p = node("p")
    p.extend(list(paragraph("Al", bold=True)))
    p.extend(list(paragraph("pha beta gamma delta.", colour="E60054")))
    save_docx(directory / "WF-four-formatting.docx", [p])
    opening = "The opening section is unchanged and provides a clear starting point for this comparison."
    moved = "This entire paragraph describes the earlier archive process and will be moved to a different position without changing any of its original words."
    anchor = "The next section contains detailed information about project delivery schedules, review responsibilities, expected turnaround times, and the local document storage arrangements used by the team."
    closing = "The final fixed section explains the approval process in detail. Each reviewer checks the document carefully before the final version is released for use by the wider team."
    create("WF-move-archive.docx", [opening, moved, anchor, closing])
    create("WF-move-new.docx", [opening, anchor, closing, moved])
    old_description = ["Alpha beta gamma.", "Keep this stable paragraph.", "Old ending."]
    new_description = ["Alpha bright beta gamma.", "Keep this stable paragraph.", "Fresh ending."]
    old_claims = ["1. A device with a blue housing.", "2. A process requiring heat."]
    new_claims = ["1. A device with a green housing.", "2. A process requiring heat.", "3. A system controlling pressure."]
    prepared = ["Description"] + old_description + ["Claims"] + old_claims
    create("EP-prepared-archive.docx", prepared)
    create("EP-prepared-identical.docx", prepared)
    create("EP-uncropped-archive.docx", ["Synthetic cover material to exclude."] + prepared +
           ["Ansprüche", "1. Synthetic German translation to exclude.", "Revendications", "1. Synthetic French translation to exclude.",
            "Drawings", "Synthetic drawing caption to exclude.", "References cited", "Synthetic bibliography to exclude."])
    create("EP-description-changed.docx", ["Description"] + new_description + ["Claims"] + old_claims)
    create("EP-claims-changed.docx", ["Description"] + old_description + ["Claims"] + new_claims)
    create("EP-both-changed.docx", ["Description"] + new_description + ["Claims"] + new_claims)
    directory.joinpath("README.txt").write_text(
        "Synthetic pilot fixtures for Document Compare " + APP_VERSION + ".\n"
        "No client content is present. Do not edit these reference files.\n"
        "See ../PilotTestPlan.docx or ../PilotTestPlan.md for archive/new pairs, WF/EP selection and expected results.\n"
        "Use approved copies of real EP/WF PDFs for PDF/OCR and complex-layout checks.\n"
        "For EP new files received as PDF or converted from PDF into Word, prepare both inputs to the same intended scope and choose WF.\n"
        "Browse supports local, OneDrive and accessible shared-drive paths; results can use these destinations too. Private processing remains local.\n", encoding="utf-8")

if __name__ == "__main__":
    build_plan()
    build_fixtures()
    print("Created pilot plan DOCX and 15 synthetic Word fixtures locally.")
