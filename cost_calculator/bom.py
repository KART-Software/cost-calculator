from cost_calculator.cost import CostCategory
import openpyxl
from enum import IntEnum


class BomSheet:
    COST_COLUMNS = [9, 10, None, 11, 12]
    QUANTITY_COLUMN = 8

    def __init__(self, path: str):
        self.filePath = path
        self.book = openpyxl.load_workbook(path)
        self.sheet = self.book.worksheets[1]
        self._detectBaseRowAndColumns()
        self._detectSystemAssemblyRowRanges()

    def _detectBaseRowAndColumns(self):
        for row in range(1, 10):
            if self.sheet.cell(row, 1).value == 1:
                if self.sheet.cell(row + 1, 1).value == 2:
                    self.baseRow = row - 1
                    break
            if row >= 9:
                self.isNotBom = True
                #error

        self.costColumns = [None, None, None, None, None]
        for column in range(1, self.sheet.max_column + 1):
            cellValue = self.sheet.cell(self.baseRow, column).value
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
        for row in range(startRow, self.sheet.max_row + 1):
            if self.sheet.cell(row, 1).value == None:
                for category in SystemAssemblyCategory:
                    if self.sheet.cell(row, 2).value == category.categoryName:
                        endRow = row - 1
                        self.systemAssemblyRowRanges[category] = (startRow,
                                                                  endRow)
                        startRow = row + 1
                        break

    # def enterCost


class SystemAssemblyCategory(IntEnum):
    BreakSystem = 0
    EngineAndDrivetrain = 1
    FrameAndBody = 2
    Electrical = 3
    Miscellaneous_FinishAndAssembly = 4
    SteeringSystem = 5
    SuspensionSystem = 6
    Wheels_WheelBearingsAndTires = 7

    @property
    def categoryName(self) -> str:
        CATEGORY_NAMES = [
            "Brake System", "Engine & Drivetrain", "Frame & Body",
            "Electrical", "Miscellaneous, Finish & Assembly",
            "Steering System", "Suspension System",
            "Wheels, Wheel Bearings and Tires"
        ]
        return CATEGORY_NAMES[self]