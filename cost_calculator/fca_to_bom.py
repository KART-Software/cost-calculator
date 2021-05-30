from glob import glob
from cost_calculator import Fca, BomSheet


def fcaToBom(fcaDirectoryPath: str, bomFilePath: str):
    fcaToBom = FcaToBom()
    if fcaToBom.setBom(bomFilePath):
        fcaFilePaths = glob(fcaDirectoryPath + "/*.xlsx")
        fcaFilePaths.extend(glob(fcaDirectoryPath + "/*/*.xlsx"))
        fcaFilePaths.extend(glob(fcaDirectoryPath + "/*/*/*.xlsx"))
        for path in fcaFilePaths:
            if fcaToBom.setFca(path):
                fcaToBom.start()
        fcaToBom.save()
    else:
        print("正しいBOMファイルを指定してください。")


class FcaToBom:
    fca: Fca
    bomSheet: BomSheet

    def setFca(self, path: str) -> bool:
        self.fca = Fca(path, data_only=True)
        return self.fca.isFca

    def setBom(self, path: str) -> bool:
        self.bomSheet = BomSheet(path)
        return not self.bomSheet.isNotBomSheet

    def start(self):
        for sheet in self.fca.fcaSheets:
            self.bomSheet.enterData(sheet)

    def save(self):
        self.bomSheet.save()
