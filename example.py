from cost_calculator.categories import CostCategory
from cost_calculator.costtable_to_fca import CostTableToFca
from cost_calculator import FcaSheet, BomSheet
from cost_calculator.cost_table import CostTable
import openpyxl
from cost_calculator import costTableToFca

sheet = openpyxl.load_workbook(
    "example/fca_files/015_Kyoto University_FSAEJ_CR_FCA_BR-A1010-AA.xlsx",
    data_only=True).worksheets[1]
fcaSheet = FcaSheet(sheet)
fcaSheet.putfcaFilePath(
    "example/fca_files/015_Kyoto University_FSAEJ_CR_FCA_BR-A1010-AA.xlsx")
print(fcaSheet.subTotalColumns)
print(fcaSheet.categoryRowRanges)
print(fcaSheet.getSubTotal(CostCategory.Material))
# print(fcaSheet.parent)
# print(fcaSheet.isNotFcaSheet)
# print(fcaSheet.systemAssemblyCategory)
# print(fcaSheet.categoryRowRanges)
# print(fcaSheet.categoryRows)
# print(fcaSheet.subTotalColumns)
# costTable = CostTable("example/cost_table_files/tblFasteners_2020J_v11.0.xlsx")
# print(costTable.valueCollums)

# costTableToFca("example/cost_table_files", "example/fca_files")
# costTableToFca("", "example/fca_files", deleteMode=True)
# costTableToFca("", "fca", deleteMode=True)
# costTable = CostTable("tables/tblmpl.xlsx")
# print(costTable.getCost("Aluminium"))

# ctf = CostTableToFca()
# ctf.setFca("fca/fca.xlsx")
# ctf.deleteCost()
# ctf.save()

bom = BomSheet("example/Brake System.xlsx")
bom.enterData(fcaSheet)
bom.save()