import os
import zipfile
import xml
from xml.etree import ElementTree

class Relationship:
    """
    * §9.2
    * .xml.rels file's component
    * Assign a rid to a target like a sheet or an image
    """
    @staticmethod
    def from_xml(el: xml.etree.ElementTree.Element):
        return Relationship(
            el.get('Id'),
            el.get('Target'),
            el.get('Type'),
        )

    def __init__(self, id: str, target: str, type: str):
        self.id = id
        self.target = target
        self.type = type

    def __str__(self) -> str:
        return str(self.__dict__)

class Relationships:
    """
    * §9.2
    * *.xml.rels file (e.g. xl/_rels/workbook.xml.rels)
    """
    @staticmethod
    def from_archive(xlsx: zipfile.ZipFile, path: str):
        try:
            rels_xml_str = xlsx.read(path).decode()
            rels_el = ElementTree.fromstring(rels_xml_str)
            ns = {
                '': 'http://schemas.openxmlformats.org/package/2006/relationships',
            }
            relationships = rels_el.findall('.//Relationship', ns)
            items = [];
            for r in relationships:
                items.append(Relationship.from_xml(r))
            return Relationships(items)
        except:
            return Relationships([])

    def __init__(self, items: list[Relationship]):
        self.items = items

    def __getitem__(self, index):
        return self.items[index]

    def __setitem__(self, index):
        return self.items[index]

    def __len__(self):
        return len(self.items)

    def get(self, rid: str) -> Relationship:
        for i in self.items:
            if i.id == rid:
                return i
        return None

class AnchorPoint:
    """
    * §20.5.2.15 (from) or v20.5.2.32 (to)
    * xdr:from or xdr:to
    """
    def __init__(self, col: int, row: int):
        self.col = col # 0-based
        self.row = row # 0-based

class TwoCellAnchor:
    """
    * §20.5.2.33
    * xl/drawings/drawing1.xml file's component
    * An anchor (position) information for a group, a shape, or a drawing element
    """
    @staticmethod
    def from_xml(el: xml.etree.ElementTree, rels: Relationships, file_path: str):
        dir_path = os.path.dirname(file_path)
        ns = {
            'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        }
        fromPt = AnchorPoint(
            col=int(el.find('.//xdr:from/xdr:col', ns).text),
            row=int(el.find('.//xdr:from/xdr:row', ns).text),
        )
        toPt = AnchorPoint(
            col=int(el.find('.//xdr:to/xdr:col', ns).text),
            row=int(el.find('.//xdr:to/xdr:row', ns).text),
        )
        title = el.find('.//xdr:pic/xdr:nvPicPr/xdr:cNvPr', ns).get('name')
        cx = el.find('.//xdr:pic/xdr:spPr/a:xfrm/a:ext', ns).get('cx')
        cy = el.find('.//xdr:pic/xdr:spPr/a:xfrm/a:ext', ns).get('cy')
        embed_rid = el.find('.//*/xdr:blipFill/a:blip', ns).get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
        embed = rels.get(embed_rid)
        image_path = os.path.normpath(os.path.join(dir_path, embed.target))
        return TwoCellAnchor(fromPt, toPt, title, cx, cy, embed, image_path)

    def __init__(
            self,
            from_pt: AnchorPoint,
            to_pt: AnchorPoint,
            title: str,
            cx: int, cy: int,
            embed: Relationship,
            image_path: str):
        self.from_pt = from_pt
        self.to_pt = to_pt
        self.title = title
        self.cx = cx
        self.cy = cy
        self.embed = embed
        self.image_path = image_path

    def __str__(self) -> str:
        return f'from:{self.from_pt.col},{self.from_pt.row}, to:{self.to_pt.col},{self.to_pt.row}, embed:{self.embed}'

class SpreadsheetDrawing:
    """
    * §20.5
    * xl/drawings/drawing.xml
    """
    @staticmethod
    def from_archive(xlsx: zipfile.ZipFile, path: str):
        dir_path = os.path.dirname(path)
        file_name = os.path.basename(path)
        drawing_xml_str = xlsx.read(path).decode()
        rels_path = os.path.join(dir_path, '_rels', f'{file_name}.rels')

        drawing_xml_rels = Relationships.from_archive(xlsx, rels_path)

        drawing_el = ElementTree.fromstring(drawing_xml_str)
        ns = {
            'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
        }
        two_cell_anchors_el = drawing_el.findall('.//xdr:twoCellAnchor', ns)
        items = []
        min_from_row = 0
        max_from_row = 0
        min_from_col = 0
        max_from_col = 0
        is_first = True
        for anchor_el in two_cell_anchors_el:
            anchor = TwoCellAnchor.from_xml(anchor_el, drawing_xml_rels, path);
            items.append(anchor)

            if is_first == True:
                min_from_row = anchor.from_pt.row
                max_from_row = anchor.from_pt.row
                min_from_col = anchor.from_pt.col
                max_from_col = anchor.from_pt.col
            else:
                min_from_row = min(min_from_row, anchor.from_pt.row)
                max_from_row = max(max_from_row, anchor.from_pt.row)
                min_from_col = min(min_from_col, anchor.from_pt.col)
                max_from_col = max(max_from_col, anchor.from_pt.col)
            is_first = False
        return SpreadsheetDrawing(items, min_from_row, max_from_row, min_from_col, max_from_col)

    def __init__(self, two_cell_anchors: list[TwoCellAnchor], min_from_row: int, max_from_row: int, min_from_col, max_from_col: int):
        self.two_cell_anchors = two_cell_anchors
        self.min_from_row = min_from_row
        self.max_from_row = max_from_row
        self.min_from_col = min_from_col
        self.max_from_col = max_from_col

class Sheet:
    """
    * $18.3
    * xl/worksheets/sheet1.xml
    """
    @staticmethod
    def from_archive(xlsx: zipfile.ZipFile, path: str, name: str, id: str):
        dir_path = os.path.dirname(path)
        file_name = os.path.basename(path)
        sheet_xml_str = xlsx.read(path).decode()
        rels_path = os.path.join(dir_path, '_rels', f'{file_name}.rels')

        sheet_xml_rels = Relationships.from_archive(xlsx, rels_path)

        ns = {
            '': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
        }
        sheet_xml_el = ElementTree.fromstring(sheet_xml_str)
        drawing_el = sheet_xml_el.find('.//drawing', ns)
        drawing = None
        if drawing_el != None:
            rid = drawing_el.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
            rel = sheet_xml_rels.get(rid)
            if rel != None:
                drawing_path = os.path.normpath(os.path.join(dir_path, rel.target))
                drawing = SpreadsheetDrawing.from_archive(xlsx, drawing_path)

        return Sheet(name, id, drawing)

    def __init__(self, name: str, id: str, drawing: SpreadsheetDrawing):
        self.name = name
        self.id = id
        self.drawing = drawing

class Workbook:
    """
    * $18.2
    * xl/workbook.xml
    """
    @staticmethod
    def parse(path: str):
        xlsx = zipfile.ZipFile(path)
        workbook_xml_str = xlsx.read('xl/workbook.xml').decode()
        workbook_rels = Relationships.from_archive(xlsx, 'xl/_rels/workbook.xml.rels')

        ns = {
            '': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
        }
        workbook_el = ElementTree.fromstring(workbook_xml_str)
        sheet_els = workbook_el.findall('.//sheets/sheet', ns)

        sheets: list[Sheet] = []
        for sheet_el in sheet_els:
            name = sheet_el.get('name')
            id = sheet_el.get('sheetId')
            rid = sheet_el.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
            rel = workbook_rels.get(rid);
            sheets.append(Sheet.from_archive(xlsx, f'xl/{rel.target}', name, id))

        # load images
        media = {}
        for sheet in sheets:
            if sheet.drawing != None:
                for anchor in sheet.drawing.two_cell_anchors:
                    fs = xlsx.open(anchor.image_path)
                    if fs.readable():
                        media[anchor.image_path] = fs.read()

        return Workbook(sheets, media)

    def __init__(self, sheets: list[Sheet], media: dict):
        self.sheets = sheets
        self.media = media

    def get_sheet(self, name: str) -> Sheet:
        for sheet in self.sheets:
            if sheet.name == name:
                return sheet
        return None
