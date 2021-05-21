from typing import List, Tuple
import openpyxl
from openpyxl.worksheet.worksheet import Worksheet

from cost_calculator import CostTable
from cost_calculator.cost import Cost, CostCategory


class FcaSheet:
    CATEGORY_COLUMN = 2
    UNIT_COST_COLUMN = 4
    MULTIPLIER_COLUMN = 7
    MULTVAL_COLUMN = 8
    CATEGORY_ROW_TO_CHECK_FROM = 9
    categoryRowRanges: List[tuple]
    subTotalColumns: List[int]

    def __init__(self, fcaSheet: Worksheet):
        self.fcaSheet = fcaSheet
        self._detectCategoryRowRanges()
        self._detectSubTotalColumns()

    def _detectCategoryRowRanges(self):
        category: CostCategory
        startRow: int
        self.categoryRowRanges = [None, None, None, None, None]
        startRow = None
        for row in range(FcaSheet.CATEGORY_ROW_TO_CHECK_FROM,
                         self.fcaSheet.max_row + 1):
            cellValue = self.fcaSheet.cell(row, FcaSheet.CATEGORY_COLUMN).value
            if cellValue in Cost.CATEGORY_NAMES:
                category = Cost.CATEGORY_NAMES.index(cellValue)
                startRow = row
            if startRow:
                if row > startRow and cellValue == None:
                    endRow = row - 1
                    self.categoryRowRanges[category] = (startRow, endRow)
                    startRow = None

    def _detectSubTotalColumns(self):
        self.subTotalColumns = list(range(5))
        self.subTotalColumns[CostCategory.ProcessMultiplier] = None
        for category in CostCategory:
            if category != CostCategory.ProcessMultiplier:
                for column in range(1, self.fcaSheet.max_column + 1):
                    if self.fcaSheet.cell(self.categoryRowRanges[category][0],
                                          column).value == "Sub Total":
                        self.subTotalColumns[category] = column
                        column += 1
                        break
                    column += 1

    def enterCost(self, category: CostCategory, costTable: CostTable):
        if category == CostCategory.Process:
            # error
            pass
        row = self.categoryRowRanges[category][0] + 1
        while True:
            if (self.fcaSheet.cell(row,
                                   FcaSheet.CATEGORY_COLUMN).value == None):
                break
            if (self.fcaSheet.cell(row,
                                   FcaSheet.CATEGORY_COLUMN).value == "None"):
                self.fcaSheet.cell(row, FcaSheet.UNIT_COST_COLUMN, value=0)
                break
            self.fcaSheet.cell(row,
                               FcaSheet.UNIT_COST_COLUMN,
                               value=costTable.getCost(
                                   self.fcaSheet.cell(
                                       row, FcaSheet.CATEGORY_COLUMN).value))
            row += 1

    def enterProcessCost(self, tableProcesses: CostTable,
                         tableProcessMultipliers: CostTable):
        MULTIPLIER_PREFIXES = ["", "Machine - ", "Material - "]
        row = self.categoryRowRanges[CostCategory.Process][0] + 1
        while True:
            if (self.fcaSheet.cell(row,
                                   FcaSheet.CATEGORY_COLUMN).value == None):
                break
            cost = tableProcesses.getCost(
                self.fcaSheet.cell(row, FcaSheet.CATEGORY_COLUMN).value)
            self.fcaSheet.cell(row, FcaSheet.UNIT_COST_COLUMN, value=cost)
            if self.fcaSheet.cell(row,
                                  FcaSheet.MULTIPLIER_COLUMN).value == None:
                multiplier = Cost(1.0)
            else:
                multiplier = tableProcessMultipliers.getCost(
                    self.fcaSheet.cell(row, FcaSheet.MULTIPLIER_COLUMN).value)
                for prefix in MULTIPLIER_PREFIXES:
                    multiplier_ = tableProcessMultipliers.getCost(
                        prefix + self.fcaSheet.cell(
                            row, FcaSheet.MULTIPLIER_COLUMN).value)
                    if multiplier_ != None:
                        multiplier = multiplier_
                        break
                self.fcaSheet.cell(row,
                                   FcaSheet.MULTVAL_COLUMN,
                                   value=multiplier)
            row += 1

    def deleteCost(self, category: CostCategory):
        if category == CostCategory.Process:
            # error
            pass
        row = self.categoryRowRanges[category][0] + 1
        while True:
            if (self.fcaSheet.cell(row,
                                   FcaSheet.CATEGORY_COLUMN).value == None):
                break
            self.fcaSheet.cell(row, FcaSheet.UNIT_COST_COLUMN, value="")
            row += 1

    def deleteProcessCost(self):
        row = self.categoryRowRanges[CostCategory.Process][0] + 1
        while True:
            if (self.fcaSheet.cell(row,
                                   FcaSheet.CATEGORY_COLUMN).value == None):
                break
            self.fcaSheet.cell(row, FcaSheet.UNIT_COST_COLUMN, value="")
            self.fcaSheet.cell(row, FcaSheet.MULTVAL_COLUMN, value="")
            row += 1


class Fca:
    fcaSheets: List[FcaSheet]

    def __init__(self, path: str):
        self.filePath = path
        self.fcaBook = openpyxl.load_workbook(path)
        # self.fcaSheets = [FcaSheet(sheet) for sheet in self.fcaBook.worksheets]
        self.fcaSheets = []
        for sheet in self.fcaBook.worksheets:
            if sheet["A1"].value == "University":
                self.fcaSheets.append(FcaSheet(sheet))