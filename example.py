import openpyxl
from cost_calculator import costTableToFca

# fcaSheet = FcaSheet(openpyxl.load_workbook("fca/fca.xlsx").worksheets[0])
# print(fcaSheet.categoryRows)
# print(fcaSheet.subTotalColumns)
# costTable = CostTable("example/cost_table_files/tblFasteners_2020J_v11.0.xlsx")
# print(costTable.valueCollums)

costTableToFca("example/cost_table_files", "example/fca_files")
# costTableToFca("", "example/fca_files", deleteMode=True)
