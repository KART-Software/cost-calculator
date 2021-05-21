from typing import List
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
    categoryRows: List[int]
    subTotalColumns: List[int]

    def __init__(self, fcaSheet: Worksheet):
        self.fcaSheet = fcaSheet
        self._detectCategoryRows()
        self._detectSubTotalColumn()

    def _detectCategoryRows(self):
        self.categoryRows = list(range(5))
        self.categoryRows[CostCategory.ProcessMultiplier] = None
        row = FcaSheet.CATEGORY_ROW_TO_CHECK_FROM
        for category in CostCategory:
            if category != CostCategory.ProcessMultiplier:
                while True:
                    if self.fcaSheet.cell(row, FcaSheet.CATEGORY_COLUMN
                                          ).value == category.categoryName:
                        self.categoryRows[category] = row
                        row += 1
                        break
                    row += 1

    def _detectSubTotalColumn(self):
        self.subTotalColumns = list(range(5))
        self.subTotalColumns[CostCategory.ProcessMultiplier] = None
        for category in CostCategory:
            if category != CostCategory.ProcessMultiplier:
                for column in range(1, self.fcaSheet.max_column + 1):
                    if self.fcaSheet.cell(self.categoryRows[category],
                                          column).value == "Sub Total":
                        self.subTotalColumns[category] = column
                        column += 1
                        break
                    column += 1

    def enterCost(self, category: CostCategory, costTable: CostTable):
        if category == CostCategory.Process:
            # error
            pass
        row = self.categoryRows[category] + 1
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
        row = self.categoryRows[CostCategory.Process] + 1
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
        row = self.categoryRows[category] + 1
        while True:
            if (self.fcaSheet.cell(row,
                                   FcaSheet.CATEGORY_COLUMN).value == None):
                break
            self.fcaSheet.cell(row, FcaSheet.UNIT_COST_COLUMN, value="")
            row += 1

    def deleteProcessCost(self):
        row = self.categoryRows[CostCategory.Process] + 1
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