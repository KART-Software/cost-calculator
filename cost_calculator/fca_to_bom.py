from cost_calculator import Fca, BomSheet


class FcaToBom:
    def setFca(self, path: str):
        self.fca = Fca(path)

    def setBom(self, path: str):
        self.bomSheet = BomSheet(path)

    def start():
        pass

    def save(self):
        self.bomSheet.book.save(self.bomSheet.filePath)
