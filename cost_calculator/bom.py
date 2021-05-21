import openpyxl

class BomSheet:
    def __init__(self, path: str):
        self.filePath = path
        self.book = openpyxl.load_workbook(path)
        self.bomSheet = self.bomBook[1]
    
    # def enterCost