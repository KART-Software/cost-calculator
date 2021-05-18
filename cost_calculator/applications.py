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


class CostTable:
    NAME_AND_VALUE_NAME = {
        CostCategory.Material: ("Material", ("Table Price", "Calc Value")),
        CostCategory.Process: ("Process", ("Unit Cost",)),
        CostCategory.ProcessMultiplier: ("Process Multiplier", ("Multiplier",)),
        CostCategory.Fastener: ("Fastener", ("Table Price", "Calc Price")),
        CostCategory.Tooling: ("Process", ("Cost",)),
    }
    NAME_COLUMN_NUMBER = 1

    def __init__(self, costCategory: CostCategory, path: str):
        self.costSheet = openpyxl.load_workbook(path, data_only=True).worksheets[0]
        self.costCategory = costCategory
        self._setBaseRowAndCollumNumber()

    def _setBaseRowAndCollumNumber(self):
        for i in range(1, 5):
            if (
                self.costSheet[i][CostTable.NAME_COLUMN_NUMBER].value
                == CostTable.NAME_AND_VALUE_NAME[self.costCategory][0]
            ):
                self.baseRowNumber = i
                break
            if i >= 4:
                # error
                break
        numbers = []
        for j, cell in enumerate(self.costSheet[self.baseRowNumber]):
            if cell.value in CostTable.NAME_AND_VALUE_NAME[self.costCategory][1]:
                numbers.append(j)
        self.valueCollumNumbers = tuple(numbers)

    def getCost(self, costName: str) -> Cost:
        for i in range(self.baseRowNumber + 1, self.costSheet.max_row):
            if self.costSheet[i][CostTable.NAME_COLUMN_NUMBER].value == None:
                # error
                break
            if self.costSheet[i][CostTable.NAME_COLUMN_NUMBER].value == costName:
                # print(i)
                for j in self.valueCollumNumbers:
                    if (
                        type(self.costSheet[i][j].value) == float
                        or type(self.costSheet[i][j].value) == int
                    ):
                        return Cost(float(self.costSheet[i][j].value))
