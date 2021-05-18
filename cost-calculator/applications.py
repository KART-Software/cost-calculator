from enum import Enum
import openpyxl
import pathlib

from openpyxl.worksheet.worksheet import Worksheet


class CostCategory(Enum):
    Material = "Material"
    Process = "Process"
    ProcessMultiplier = "ProcessMultiPlier"
    Fastener = "Fastener"
    Tooling = "Tooling"


class Cost(float):
    pass


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

    def setFca(path: str):
        pass

    def start():
        pass


class FcaToBom:
    def setFca(path: str):
        pass

    def setBom(path: str):
        pass

    def start():
        pass


class CostTable(Worksheet):
    NAME_AND_VALUE_NAME = {
        CostCategory.Material: ("Material", ("Table Price", "Calc Value")),
        CostCategory.Process: ("Process", ("Unit Cost")),
        CostCategory.ProcessMultiplier: ("Process Multiplier", "Multiplier"),
        CostCategory.Fastener: ("Fastener", ("Table Price", "Calc Price")),
        CostCategory.Tooling: ("Process", ("Cost")),
    }
    NAME_COLUMN_NUMBER = 1

    def __init__(self, costCategory: CostCategory, path: str):
        self = openpyxl.load_workbook(path).worksheets[0]
        self.costCategory = costCategory
        self._setBaseRowAndCollumNumber()

    def _setBaseRowAndCollumNumber(self):
        for i in range(1, 5):
            if (
                self[i][CostTable.NAME_COLUMN_NUMBER].value
                == CostTable.NAME_AND_VALUE_NAME[self.costCategory][0]
            ):
                self.baseRowNumber = i
                break
            if i >= 4:
                # error
                break
        numbers = []
        for j, cell in enumerate(self[self.baseRowNumber]):
            if cell.value in CostTable.NAME_AND_VALUE_NAME[self.costCategory][1]:
                numbers.append(j)
        self.valueCollumNumbers = tuple(numbers)

    def getCost(self, costName: str):
        for i in range(self.baseRowNumber + 1, self.max_row):
            if self[i][CostTable.NAME_COLUMN_NUMBER] == None:
                # error
                break
            if self[i][CostTable.NAME_COLUMN_NUMBER] == costName:
                for j in self.valueCollumNumbers:
                    if type(self[i][j]) == float:
                        return Cost(self[i][j])
