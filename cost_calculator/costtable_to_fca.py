from enum import IntEnum
from typing import List
import openpyxl
from glob import glob
from openpyxl.worksheet.worksheet import Worksheet

from cost_calculator import CostTable, Fca
from cost_calculator.categories import CostCategory


def costTableToFca(costTablesDirectoryPath: str,
                   fcaDirectoryPath: str,
                   deleteMode=False):
    costTableToFca = CostTableToFca()
    if deleteMode == False:
        if costTableToFca.setCostTables(costTablesDirectoryPath):
            pass
    fcaFilePaths = glob(fcaDirectoryPath + "/*.xlsx")
    fcaFilePaths.extend(glob(fcaDirectoryPath + "/*/*.xlsx"))
    fcaFilePaths.extend(glob(fcaDirectoryPath + "/*/*/*.xlsx"))

    for path in fcaFilePaths:
        if costTableToFca.setFca(path):
            if deleteMode == True:
                costTableToFca.deleteCost()
            else:
                costTableToFca.start()
            costTableToFca.save()


class CostTableToFca:

    tableMaterials: Worksheet
    tableProesses: Worksheet
    tableProcessMultipliers: Worksheet
    tableFasteners: Worksheet
    tableTooling: Worksheet
    fca: Fca

    def setCostTables(self, costTablesDirectoryPath: str) -> bool:
        costTableFilePaths = glob(costTablesDirectoryPath + "/*")
        costTables: List[CostTable]
        costTables = []
        for path in costTableFilePaths:
            costTable = CostTable(path)
            if costTable.isNotCostTable == False:
                costTables.append(costTable)
        # costTables = [
        #     CostTable(path) for path in costTableFilePaths
        #     if CostTable(path).isNotCostTable == False
        # ]
        categoryOfTables = [table.category for table in costTables]
        if len(categoryOfTables) != 5:
            return False
        for i in range(5):
            for j in range(5):
                if i != j and categoryOfTables[i] == categoryOfTables[j]:
                    return False
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
        return True

    def setFca(self, path: str) -> bool:
        self.fca = Fca(path)
        return self.fca.isFca

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
        self.fca.save()