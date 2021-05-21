from cost_calculator.costtable_to_fca import CostTable, FcaSheet, costTableToFca
import openpyxl

# fcaSheet = FcaSheet(openpyxl.load_workbook("fca/fca.xlsx").worksheets[0])
# print(fcaSheet.categoryRows)
# print(fcaSheet.subTotalColumns)
# costTable = CostTable("example/cost_table_files/tblFasteners_2020J_v11.0.xlsx")
# print(costTable.valueCollums)

costTableToFca("example/cost_table_files", "example/fca_files")
costTableToFca("", "example/fca_files")
