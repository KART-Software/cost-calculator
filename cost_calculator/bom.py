from cost_calculator.fca import FcaSheet
from typing import List
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from cost_calculator import FcaSheet
from cost_calculator.categories import CostCategory, SystemAssemblyCategory
import openpyxl


class BomSheet:
    bomBook: Workbook
    bomSheet: Worksheet
    isNotBomSheet: bool
    costColumns: List[int]
    componentColumn: int
    quantityColumn: int
    systemAssemblyRowRanges: List[tuple]

    def __init__(self, path: str):
        self.filePath = path
        self.bomBook = openpyxl.load_workbook(path)
        self.bomSheet = self.bomBook.worksheets[1]
        self._detectBaseRowAndColumns()
        self._detectSystemAssemblyRowRanges()

    def _detectBaseRowAndColumns(self):
        for row in range(1, 10):
            if self.bomSheet.cell(row, 1).value == 1:
                if self.bomSheet.cell(row + 1, 1).value == 2:
                    self.baseRow = row - 1
                    break
            if row >= 9:
                self.isNotBomSheet = True
                #error

        self.costColumns = [None, None, None, None, None]
        for column in range(1, self.bomSheet.max_column + 1):
            cellValue = self.bomSheet.cell(self.baseRow, column).value
            if cellValue == "Component":
                self.componentColumn = column
            if cellValue == "Quantity":
                self.quantityColumn = column
            if cellValue == CostCategory.Material.categoryName + " Cost":
                self.costColumns[CostCategory.Material] = column
                self.costColumns[CostCategory.Process] = column + 1
                self.costColumns[CostCategory.Fastener] = column + 2
                self.costColumns[CostCategory.Tooling] = column + 3
            if cellValue == "Link to FCA Sheet":
                self.linkToFcaSheetColumn = column

    def _detectSystemAssemblyRowRanges(self):
        self.systemAssemblyRowRanges = [
            None, None, None, None, None, None, None, None
        ]
        startRow = self.baseRow + 1
        for row in range(startRow, self.bomSheet.max_row + 1):
            if self.bomSheet.cell(row, 1).value == None:
                for category in SystemAssemblyCategory:
                    if self.bomSheet.cell(row,
                                          2).value in category.categoryName:
                        endRow = row - 1
                        self.systemAssemblyRowRanges[category] = (startRow,
                                                                  endRow)
                        startRow = row + 1
                        break

    def enterCost(self, fcaSheet: FcaSheet):
        rowRange = self.systemAssemblyRowRanges[
            fcaSheet.systemAssemblyCategory]
        component = fcaSheet.fcaSheet.title
        for row in range(rowRange[0], rowRange[1] + 1):
            if self.bomSheet.cell(row, self.componentColumn) == component:
                for category in CostCategory:
                    if category != CostCategory.ProcessMultiplier:
                        self.bomSheet.cell(
                            row,
                            self.costColumns[category],
                            value=fcaSheet.fcaSheet.cell(
                                fcaSheet.categoryRowRanges[1] + 1,
                                fcaSheet.subTotalColumns[category]))

    # def save(se)