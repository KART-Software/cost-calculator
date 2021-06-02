from typing import List
from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LAParams, LTContainer, LTTextBox, LTComponent
from pdfminer.pdfinterp import PDFPageInterpreter, PDFResourceManager
from pdfminer.pdfpage import PDFPage


class SupplPdf:
    STRINGS_MUST_INCLUDED = [
        "SOLIDWORKS", "裏付け資料", "複合材部品使用プリプレグ", "System Code", "Date:",
        "Cost Report"
    ]
    PAGES_TO_CHECK = [0]

    filePath: str
    isSupplPDF: bool
    boxesInPages: List[List[LTTextBox]]

    def __init__(self, path: str):
        self.filePath = path
        self._judge()
        if self.isSupplPDF == True:
            self._getTextBoxes(path)

    def _judge(self):
        self.isSupplPDF = False
        boxes: List[LTTextBox]
        laparams = LAParams()
        resource_manager = PDFResourceManager()
        device = PDFPageAggregator(resource_manager, laparams=laparams)
        interpreter = PDFPageInterpreter(resource_manager, device)
        with open(self.filePath, "rb") as f:
            pages = PDFPage.get_pages(f, pagenos=SupplPdf.PAGES_TO_CHECK)
            boxes = []
            for page in pages:
                interpreter.process_page(page)
                layout = device.get_result()
                boxes.extend(findTextBoxesRecursively(layout))
        for box in boxes:
            for string in SupplPdf.STRINGS_MUST_INCLUDED:
                if string in box.get_text().strip():
                    self.isSupplPDF = True

    def _getTextBoxes(self, path: str):
        laparams = LAParams()  # Layout Analysisの設定で縦書きの検出を有効にする。
        resource_manager = PDFResourceManager()  # 共有のリソースを管理するリソースマネージャーを作成。
        # ページを集めるPageAggregatorオブジェクトを作成。
        device = PDFPageAggregator(resource_manager, laparams=laparams)
        interpreter = PDFPageInterpreter(resource_manager, device)
        with open(path, "rb") as f:
            pages = PDFPage.get_pages(f)
            self.boxesInPages = []
            for pageNumber, page in enumerate(pages):
                interpreter.process_page(page)  # ページを処理する。
                layout = device.get_result()  # LTPageオブジェクトを取得。

                boxes = findTextBoxesRecursively(
                    layout)  # ページ内のテキストボックスのリストを取得する。
                self.boxesInPages.append(boxes)

    def pageOfId(self, id: str) -> int:
        for page, boxes in enumerate(self.boxesInPages):
            for box in boxes:
                if id in box.get_text().strip():
                    return page + 1


def findTextBoxesRecursively(component: LTComponent) -> List[LTTextBox]:
    """
   再帰的にテキストボックス（LTTextBox）を探して、テキストボックスのリストを取得する。
   """
    # LTTextBoxを継承するオブジェクトの場合は1要素のリストを返す。
    if isinstance(component, LTTextBox):
        return [component]

    # LTContainerを継承するオブジェクトは子要素を含むので、再帰的に探す。
    if isinstance(component, LTContainer):
        boxes = []
        for child in component:
            boxes.extend(findTextBoxesRecursively(child))

        return boxes

    return []  # その他の場合は空リストを返す。