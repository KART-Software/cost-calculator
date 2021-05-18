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
        self.tableMaterials = CostTable(CostCategory.Material, tableMaterialsPath)
        self.tableProesses = CostTable(CostCategory.Process, tableProcessesPath)
        self.tableProcessMultipliers = CostTable(
            CostCategory.ProcessMultiplier, tableProcessMultipliersPath
        )
        self.tableFasteners = CostTable(CostCategory.Fastener, TableFastenersPath)
        self.tableTooling = CostTable(CostCategory.Tooling, TableToolingPath)

    def setFca(self, path: str):
        self.fca = Fca(path)

    def start():
        pass


class CostTable:
    NAME_AND_VALUE_NAME = {
        CostCategory.Material: ("Material", ("Table Price", "Calc Value")),
        CostCategory.Process: ("Process", ("Unit Cost",)),
        CostCategory.ProcessMultiplier: ("Process Multiplier", ("Multiplier",)),
        CostCategory.Fastener: ("Fastener", ("Table Price", "Calc Price")),
        CostCategory.Tooling: ("Process", ("Cost",)),
    }
    NAME_COLUMN = 1

    def __init__(self, costCategory: CostCategory, path: str):
        self.costSheet = openpyxl.load_workbook(path, data_only=True).worksheets[0]
        self.costCategory = costCategory
        self._setBaseRowAndCollum()

    def _setBaseRowAndCollum(self):
        for i in range(1, 5):
            if (
                self.costSheet[i][CostTable.NAME_COLUMN].value
                == CostTable.NAME_AND_VALUE_NAME[self.costCategory][0]
            ):
                self.baseRow = i
                break
            if i >= 4:
                # error
                break
        numbers = []
        for j, cell in enumerate(self.costSheet[self.baseRow]):
            if cell.value in CostTable.NAME_AND_VALUE_NAME[self.costCategory][1]:
                numbers.append(j)
        self.valueCollums = tuple(numbers)

    def getCost(self, costName: str) -> Cost:
        for i in range(self.baseRow + 1, self.costSheet.max_row + 1):
            if self.costSheet[i][CostTable.NAME_COLUMN].value == None:
                # error
                break
            if self.costSheet[i][CostTable.NAME_COLUMN].value == costName:
                for j in self.valueCollums:
                    if (
                        type(self.costSheet[i][j].value) == float
                        or type(self.costSheet[i][j].value) == int
                    ):
                        return Cost(float(self.costSheet[i][j].value))


class FcaSheet:
    CATEGORY_COLUMN = 1
    UNIT_COST_COLUMN = 3
    MULTIPLIER_COLUMN = 6

    def __init__(self, fcaSheet: Worksheet):
        self.fcaSheet = fcaSheet
        self._setCategoryRows()

    def _setCategoryRows(self):
        self.categoryRows = {}
        row = 9
        for category in CostCategory:
            if category == CostCategory.ProcessMultiplier:
                continue
            while True:
                if self.fcaSheet[row][FcaSheet.CATEGORY_COLUMN].value == category.value:
                    self.categoryRows[category] = row
                    row += 1
                    break
                row += 1

    def enterCost(self, row: int, cost: Cost):
        self.fcaSheet.cell(row=row, column=FcaSheet.UNIT_COST_COLUMN, value=Cost)


class Fca:
    def __init__(self, path: str):
        self.filePath = path
        self.fcaBook = openpyxl.load_workbook(path)
        self.fcaSheets = [FcaSheet(sheet) for sheet in self.fcaBook.worksheets]
