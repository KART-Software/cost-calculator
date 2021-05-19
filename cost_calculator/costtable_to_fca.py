from enum import Enum
import openpyxl
from openpyxl.worksheet.worksheet import Worksheet


class CostCategory(Enum):
    Material = "Material"
    Process = "Process"
    ProcessMultiplier = "ProcessMultiPlier"
    Fastener = "Fastener"
    Tooling = "Tooling"


class Cost(float):
    def __add__(self, other):
        return Cost(float(self) + float(other))

    def __sub__(self, other):
        return Cost(float(self) - float(other))

    def __mul__(self, other):
        return Cost(float(self) * float(other))


class CostTableToFca:
    def setCostTables(
        self,
        tableMaterialsPath: str,
        tableProcessesPath: str,
        tableProcessMultipliersPath: str,
        TableFastenersPath: str,
        TableToolingPath: str,
    ):
        self.tableMaterials = CostTable(CostCategory.Material,
                                        tableMaterialsPath)
        self.tableProesses = CostTable(CostCategory.Process,
                                       tableProcessesPath)
        self.tableProcessMultipliers = CostTable(
            CostCategory.ProcessMultiplier, tableProcessMultipliersPath)
        self.tableFasteners = CostTable(CostCategory.Fastener,
                                        TableFastenersPath)
        self.tableTooling = CostTable(CostCategory.Tooling, TableToolingPath)

    def setFca(self, path: str):
        self.fca = Fca(path)

    def start(self):
        for sheet in self.fca.fcaSheets:
            sheet.enterCost(CostCategory.Material, self.tableMaterials)
            sheet.enterProcessCost(self.tableProesses,
                                   self.tableProcessMultipliers)
            sheet.enterCost(CostCategory.Fastener, self.tableFasteners)
            sheet.enterCost(CostCategory.Tooling, self.tableTooling)

    def save(self):
        self.fca.fcaBook.save(self.fca.filePath)
        del self.fca


class CostTable:
    GENERICTERM_VALUENAME_SHEETTITLE = {
        CostCategory.Material:
        ("Material", ("Table Price", "Calc Value"), "tblMaterials"),
        CostCategory.Process: ("Process", ("Unit Cost", ), "tblProcesses"),
        CostCategory.ProcessMultiplier:
        ("Process Multiplier", ("Multiplier", ), "tblProcessMultipliers"),
        CostCategory.Fastener:
        ("Fastener", ("Table Price", "Calc Price"), "tblFasteners"),
        CostCategory.Tooling: ("Process", ("Cost", ), "tblToolings"),
    }
    NAME_COLUMN = 1

    def __init__(self, path: str):
        self.costSheet = openpyxl.load_workbook(path,
                                                data_only=True).worksheets[0]
        self._detectCategory()
        self._detectBaseRowAndCollum()

    def _detectCategory(self):
        isNotCostTable = True
        for category in CostCategory:
            if self.costSheet.title == CostTable.GENERICTERM_VALUENAME_SHEETTITLE[
                    category][2]:
                self.category = category
                break
            isNotCostTable = isNotCostTable and self.costSheet.title == CostTable.GENERICTERM_VALUENAME_SHEETTITLE[
                category][2]
        if isNotCostTable:
            #error
            pass

    def _detectBaseRowAndCollum(self):
        for i in range(1, 5):
            if (self.costSheet[i][CostTable.NAME_COLUMN].value == CostTable.
                    GENERICTERM_VALUENAME_SHEETTITLE[self.category][0]):
                self.baseRow = i
                break
            if i >= 4:
                # error
                break
        numbers = []
        for j, cell in enumerate(self.costSheet[self.baseRow]):
            if cell.value in CostTable.GENERICTERM_VALUENAME_SHEETTITLE[
                    self.category][1]:
                numbers.append(j)
        self.valueCollums = tuple(numbers)

    def getCost(self, costName: str) -> Cost:
        for i in range(self.baseRow + 1, self.costSheet.max_row + 1):
            if self.costSheet[i][CostTable.NAME_COLUMN].value == None:
                # error
                break
            if self.costSheet[i][CostTable.NAME_COLUMN].value == costName:
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
                        FcaSheet.CATEGORY_COLUMN].value == category.value:
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
            if (self.fcaSheet[row][FcaSheet.CATEGORY_COLUMN].value == "None"
                    or self.fcaSheet[row][FcaSheet.CATEGORY_COLUMN].value
                    == None):
                break
            self.fcaSheet.cell(
                row=row,
                column=FcaSheet.UNIT_COST_COLUMN,
                value=costTable.getCost(
                    self.fcaSheet[row][FcaSheet.CATEGORY_COLUMN].value),
            )

    def enterProcessCost(self, tableProcesses: CostTable,
                         tableProcessMultipliers: CostTable):

        row = self.categoryRows[CostCategory.Process] + 1
        while True:
            if (self.fcaSheet[row][FcaSheet.CATEGORY_COLUMN].value == "None"
                    or self.fcaSheet[row][FcaSheet.CATEGORY_COLUMN].value
                    == None):
                break
            cost = tableProcesses.getCost(
                self.fcaSheet[row][FcaSheet.CATEGORY_COLUMN].value)
            self.fcaSheet.cell(row=row,
                               column=FcaSheet.UNIT_COST_COLUMN,
                               value=cost)
            # multiplier = tableProcessMultipliers.getCost(
            #     self.fcaSheet[row][FcaSheet.MULTIPLIER_COLUMN].value
            # ) if self.fcaSheet[row][FcaSheet.MULTIPLIER_COLUMN].value != None else
            if self.fcaSheet[row][FcaSheet.MULTIPLIER_COLUMN].value == None:
                multiplier = Cost(1.0)
            else:
                multiplier = tableProcessMultipliers.getCost(
                    self.fcaSheet[row][FcaSheet.MULTIPLIER_COLUMN].value)
                self.fcaSheet.cell(row=row,
                                   column=FcaSheet.MULTVAL_COLUMN,
                                   value=multiplier)


class Fca:
    def __init__(self, path: str):
        self.filePath = path
        self.fcaBook = openpyxl.load_workbook(path)
        self.fcaSheets = [FcaSheet(sheet) for sheet in self.fcaBook.worksheets]
