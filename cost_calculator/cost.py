from enum import IntEnum


class CostCategory(IntEnum):
    Material = 0
    Process = 1
    ProcessMultiplier = 2
    Fastener = 3
    Tooling = 4

    @property
    def categoryName(self) -> str:
        return Cost.CATEGORY_NAMES[self]


class Cost(float):

    CATEGORY_NAMES = [
        "Material", "Process", "ProcessMultiplier", "Fastener", "Tooling"
    ]

    def __add__(self, other):
        return Cost(float(self) + float(other))

    def __sub__(self, other):
        return Cost(float(self) - float(other))

    def __mul__(self, other):
        return Cost(float(self) * float(other))