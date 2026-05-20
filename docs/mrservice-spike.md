# MRService.dll spike

Ran at 2026-05-20T13:03:39.886576+00:00.
Python bitness: 64-bit.

**DLL located** at `C:\Users\timou\source\repos\ModelRisk\ModelRisk_Project\ModelRiskSDK\MRLibrary\_x64\MRService.dll`.

DLL loaded and signatures configured.

Single-key activation: OK (MRLIB_SetOfflineActivationKey).

## Phase B test: open .xlsx headlessly
**NOT VIABLE** — `C:\Users\timou\source\repos\ModelRisk\Documentation\ModelRiskHelp\contents\Models\First Simple Model Stage 1 Complete.xlsx` → MRLIB_OpenSimulationModel returned FALSE

## Read-path test: open .vmrs results
**VIABLE** — `C:\Users\timou\source\repos\ModelRisk\ModelRisk_Project\ModelRisk\Book1_2.vmrs` → model_ptr=1678976156096, GetModelDataLength(sim=0)=1000

## Disposition

- If the .xlsx test is VIABLE, the v0.3 architecture is fully headless: MRService.dll opens the workbook, runs the simulation, and reads results — no Excel needed.
- If only the .vmrs test is VIABLE, the v0.3 architecture is Excel-triggered + headless-read: the XLL in Excel runs the simulation, MRService.dll reads the resulting .vmrs.
- If neither is VIABLE, the issue is likely activation (MRLIB_SetOfflineActivationKeyEx2 required) or a DLL dependency that didn't resolve.