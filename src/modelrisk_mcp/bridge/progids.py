"""ProgIDs and CLSIDs for the four ModelRisk COM coclasses.

Source of truth: `ModelRisk_Project/VBAProject/ModelRiskAtl/*.rgs` in the
ModelRisk source tree. Verified 2026-05-19. See spec §8.0.
"""

from __future__ import annotations

from typing import Final

PROGID_DISTRIBUTIONS: Final = "ModelRisk"
PROGID_SIMULATION: Final = "ModelRisk.ModelRiskSimulation"
PROGID_SIMULATION_RESULTS: Final = "ModelRisk.ModelRiskSimulationResults"
PROGID_SIMULATION_SETTINGS: Final = "ModelRisk.ModelRiskSimulationSettings"

CLSID_DISTRIBUTIONS: Final = "{570013C9-8251-44CF-AF83-EDD333725537}"
CLSID_SIMULATION: Final = "{59530ADE-E690-4802-A6E4-890B72596310}"
CLSID_SIMULATION_RESULTS: Final = "{B1EEBA78-BE81-4d37-8FEA-FC3AE14BE755}"
CLSID_SIMULATION_SETTINGS: Final = "{389CD5FB-F265-467e-A255-90C206CE7220}"

TYPELIB: Final = "{ECC429DA-26E6-4D86-9B2D-1E14E0461749}"
