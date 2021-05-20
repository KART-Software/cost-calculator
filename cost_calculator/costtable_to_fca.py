from enum import IntEnum
from typing import List
import openpyxl
from glob import glob
from openpyxl.worksheet.worksheet import Worksheet


def costTableToFca(costTablesDirectryPath: str,
                   fcaDirectryPath: str,
                   deleteMode=False):
    costTableToFca = CostTableToFca()
    if deleteMode == False:
        costTableToFca.setCostTables(costTablesDirectryPath)
    fcaFilePaths = glob(fcaDirectryPath + "/*")
    for path in fcaFilePaths:
        costTableToFca.setFca(path)
        if deleteMode == True:
            costTableToFca.deleteCost()
        else:
            costTableToFca.start()
        costTableToFca.save()


class CostCategory(IntEnum):
    Material = 0
    Process = 1
    ProcessMultiplier = 2
    Fastener = 3
    Tooling = 4

    @property
    def categoryName(self) -> str:
        CATEGORY_NAMES = [
            "Material", "Process", "ProcessMultiplier", "Fastener", "Tooling"
        ]
        return CATEGORY_NAMES[self]


class Cost(float):
    def __add__(self, other):
        return Cost(float(self) + float(other))

    def __sub__(self, other):
        return Cost(float(self) - float(other))

    def __mul__(self, other):
        return Cost(float(self) * float(other))


class CostTableToFca:
    def setCostTables(self, costTablesDirectryPath: str):
        costTableFilePaths = glob(costTablesDirectryPath + "/*")
        if len(costTableFilePaths) != 5:
            #error
            pass
        costTables = [CostTable(path) for path in costTableFilePaths]
        categoryOfTables = [table.category for table in costTables]
        for i in range(5):
            for j in range(5):
                if i != j and categoryOfTables[i] == categoryOfTables[j]:
                    #error
                    pass
        costTablesSorted: List[CostTable]
        costTablesSorted = list(range(5))
        for i in range(5):
            costTablesSorted[categoryOfTables[i]] = costTables[i]
        self.tableMaterials = costTablesSorted[0]
        self.tableProesses = costTablesSorted[1]
        self.tableProcessMultipliers = costTablesSorted[2]
        self.tableFasteners = costTablesSorted[3]
        self.tableTooling = costTablesSorted[4]
        #self.tableMaterials = next(filter(lambda x : x.category == CostCategory.Material, costTables), None)
        #elm = next(filter(lambda x: x.endswith("n"), fruits), None)

    def setFca(self, path: str):
        self.fca = Fca(path)

    def start(self):
        for sheet in self.fca.fcaSheets:
            sheet.enterCost(CostCategory.Material, self.tableMaterials)
            sheet.enterProcessCost(self.tableProesses,
                                   self.tableProcessMultipliers)
            sheet.enterCost(CostCategory.Fastener, self.tableFasteners)
            sheet.enterCost(CostCategory.Tooling, self.tableTooling)

    def deleteCost(self):
        for sheet in self.fca.fcaSheets:
            sheet.deleteCost(CostCategory.Material)
            sheet.deleteProcessCost()
            sheet.deleteCost(CostCategory.Fastener)
            sheet.deleteCost(CostCategory.Tooling)

    def save(self):
        self.fca.fcaBook.save(self.fca.filePath)


class CostTable:
    # GENERICTERM_VALUENAME_SHEETTITLE = {
    #     CostCategory.Material:
    #     ("Material", ("Table Price", "Calc Value"), "tblMaterials"),
    #     CostCategory.Process: ("Process", ("Unit Cost", ), "tblProcesses"),
    #     CostCategory.ProcessMultiplier:
    #     ("Process Multiplier", ("Multiplier", ), "tblProcessMultipliers"),
    #     CostCategory.Fastener:
    #     ("Fastener", ("Table Price", "Calc Price"), "tblFasteners"),
    #     CostCategory.Tooling: ("Process", ("Cost", ), "tblToolings"),
    # }
    GENERIC_TERM = [
        "Material", "Process", "Process Multiplier", "Fastener", "Process"
    ]
    VALUE_NAME = [("Table Price", "Calc Value"), ("Unit Cost", ),
                  ("Multiplier", ), ("Table Price", "Calc Price"), ("Cost", )]
    SHEET_TITLE = [
        "tblMaterials", "tblProcesses", "tblProcessMultipliers",
        "tblFasteners", "tblTooling"
    ]

    GENERIC_TERM_COLUMN = 1

    def __init__(self, path: str):
        self.costSheet = openpyxl.load_workbook(path,
                                                data_only=True).worksheets[0]
        self._detectCategory()
        self._detectBaseRowAndCollum()

    def _detectCategory(self):
        isNotCostTable = True
        for category in CostCategory:
            if self.costSheet.title == CostTable.SHEET_TITLE[category]:
                self.category = category
                break
            isNotCostTable = isNotCostTable and self.costSheet.title != CostTable.SHEET_TITLE[
                category]
        if isNotCostTable == True:
            #error
            pass

    def _detectBaseRowAndCollum(self):
        for i in range(1, 5):
            if (self.costSheet[i][CostTable.GENERIC_TERM_COLUMN].value ==
                    CostTable.GENERIC_TERM[self.category]):
                self.baseRow = i
                break
            if i >= 4:
                # error
                break
        numbers = []
        for j, cell in enumerate(self.costSheet[self.baseRow]):
            if cell.value in CostTable.VALUE_NAME[self.category]:
                numbers.append(j)
        self.valueCollums = tuple(numbers)

    def getCost(self, costName: str) -> Cost:
        for i in range(self.baseRow + 1, self.costSheet.max_row + 1):
            if self.costSheet[i][CostTable.GENERIC_TERM_COLUMN].value == None:
                # error
                break
            if self.costSheet[i][
                    CostTable.GENERIC_TERM_COLUMN].value == costName:
                for j in self.valueCollums:
                    if (type(self.costSheet[i][j].value) == float
                            or type(self.costSheet[i][j].value) == int):
                        return Cost(float(self.costSheet[i][j].value))


class FcaSheet:
    CATEGORY_COLUMN = 1
    UNIT_COST_COLUMN = 3
    MULTIPLIER_COLUMN = 6
    MULTVAL_COLUMN = 7

    def __init__(self, fcaSheet: Worksheet):
        self.fcaSheet = fcaSheet
        self._detectCategoryRows()

    def _detectCategoryRows(self):
        self.categoryRows = {}
        row = 9
        for category in CostCategory:
            if category == CostCategory.ProcessMultiplier:
                continue
            while True:
                if self.fcaSheet[row][
                        FcaSheet.
                        CATEGORY_COLUMN].value == category.categoryName:
                    self.categoryRows[category] = row
                    row += 1
                    break
                row += 1

    def enterCost(self, category: CostCategory, costTable: CostTable):
        if category == CostCategory.Process:
            # error
            pass
        row = self.categoryRows[category] + 1
        while True:
            if (self.fcaSheet[row][FcaSheet.CATEGORY_COLUMN].value == None):
                break
            if (self.fcaSheet[row][FcaSheet.CATEGORY_COLUMN].value == "None"):
                self.fcaSheet.cell(row=row,
                                   column=FcaSheet.UNIT_COST_COLUMN + 1,
                                   value=0)
                break
            self.fcaSheet.cell(
                row=row,
                column=FcaSheet.UNIT_COST_COLUMN + 1,
                value=costTable.getCost(
                    self.fcaSheet[row][FcaSheet.CATEGORY_COLUMN].value))
            row += 1

    def enterProcessCost(self, tableProcesses: CostTable,
                         tableProcessMultipliers: CostTable):

        row = self.categoryRows[CostCategory.Process] + 1
        while True:
            if (self.fcaSheet[row][FcaSheet.CATEGORY_COLUMN].value == None):
                break
            cost = tableProcesses.getCost(
                self.fcaSheet[row][FcaSheet.CATEGORY_COLUMN].value)
            self.fcaSheet.cell(row=row,
                               column=FcaSheet.UNIT_COST_COLUMN + 1,
                               value=cost)
            if self.fcaSheet[row][FcaSheet.MULTIPLIER_COLUMN].value == None:
                multiplier = Cost(1.0)
            else:
                multiplier = tableProcessMultipliers.getCost(
                    self.fcaSheet[row][FcaSheet.MULTIPLIER_COLUMN].value)
                self.fcaSheet.cell(row=row,
                                   column=FcaSheet.MULTVAL_COLUMN + 1,
                                   value=multiplier)
            row += 1

    def deleteCost(self, category: CostCategory):
        if category == CostCategory.Process:
            # error
            pass
        row = self.categoryRows[category] + 1
        while True:
            if (self.fcaSheet[row][FcaSheet.CATEGORY_COLUMN].value == None):
                break
            self.fcaSheet.cell(row=row,
                               column=FcaSheet.UNIT_COST_COLUMN + 1,
                               value="")
            row += 1

    def deleteProcessCost(self):
        row = self.categoryRows[CostCategory.Process] + 1
        while True:
            if (self.fcaSheet[row][FcaSheet.CATEGORY_COLUMN].value == None):
                break
            self.fcaSheet.cell(row=row,
                               column=FcaSheet.UNIT_COST_COLUMN + 1,
                               value="")
            self.fcaSheet.cell(row=row,
                               column=FcaSheet.MULTVAL_COLUMN + 1,
                               value="")
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
