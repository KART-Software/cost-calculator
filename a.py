from cost_calculator.applications import CostTableToFca

costTableToFca = CostTableToFca()
costTableToFca.setCostTables(
    tableMaterialsPath="tables/tblmtr.xlsx",
    tableProcessesPath="tables/tblprc.xlsx",
    tableProcessMultipliersPath="tables/tblmpl.xlsx",
    TableFastenersPath="tables/tblfsn.xlsx",
    TableToolingPath="tables/tbltl.xlsx",
)
print(costTableToFca.tableFasteners.getCost("Galvanized Steel Loop Straps"))
print(
    costTableToFca.tableMaterials.getCost(
        "Boost Solenoid Valve, MITSUBISHI Motors OEM, MR561312"
    )
)
print(costTableToFca.tableMaterials.getCost("Motor, RC Servo"))
