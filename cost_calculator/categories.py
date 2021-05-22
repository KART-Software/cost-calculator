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


class SystemAssemblyCategory(IntEnum):
    BreakSystem = 0
    EngineAndDrivetrain = 1
    FrameAndBody = 2
    Electrical = 3
    Miscellaneous_FinishAndAssembly = 4
    SteeringSystem = 5
    SuspensionSystem = 6
    Wheels_WheelBearingsAndTires = 7

    @property
    def categoryName(self) -> str:
        CATEGORY_NAMES = [
            "Brake System", "Engine & Drivetrain", "Frame & Body",
            "Electrical", "Miscellaneous, Finish & Assembly",
            "Steering System", "Suspension System",
            "Wheels, Wheel Bearings and Tires"
        ]
        return CATEGORY_NAMES[self]