from glob import glob
from cost_calculator import Fca, BomSheet


def fcaToBom(fcaDirectoryPath: str, bomFilePath: str):
    fcaToBom = FcaToBom()
    fcaToBom.setBom(bomFilePath)
    fcaFilePaths = glob(fcaDirectoryPath + "/*.xlsx")
    fcaFilePaths.extend(glob(fcaDirectoryPath + "/*/*.xlsx"))
    fcaFilePaths.extend(glob(fcaDirectoryPath + "/*/*/*.xlsx"))
    for path in fcaFilePaths:
        fcaToBom.setFca(path)
        fcaToBom.start()
    fcaToBom.save()


class FcaToBom:
    fca: Fca
    bomSheet: BomSheet

    def setFca(self, path: str):
        self.fca = Fca(path, data_only=True)

    def setBom(self, path: str):
        self.bomSheet = BomSheet(path)

    def start(self):
        for sheet in self.fca.fcaSheets:
            self.bomSheet.enterData(sheet)

    def save(self):
        self.bomSheet.save()
