# ONGEN Excel VBA Engineering Suite

This is the Excel-first delivery. It builds the combined engineering suite from controlled copies of the supplied `.xlsm` workbooks and installs the reviewed VBA changes. The original source files remain unchanged.

## Build the Excel suite

1. Download and extract the complete `ONGEN_Excel_VBA_0.2.zip` package.
2. Use a Windows computer with desktop Microsoft Excel installed.
3. In Excel, enable **File > Options > Trust Center > Trust Center Settings > Macro Settings > Trust access to the VBA project object model**. This is required only so the builder can install the reviewed modules.
4. Close any open ONGEN workbooks, then double-click `Start-Build.cmd`.
5. Open `output\build-YYYYMMDD-HHMMSS\ONGEN Excel Engineering Suite V0.2.xlsm` from the newest build folder and select **Enable Content** only after confirming the file came from this repository/package. `output\LATEST-BUILD.txt` also records the latest successful folder.
6. Open the `ONGEN CONTROL CENTER` worksheet and run **System self-check** before an engineering workflow.

Every run uses a new timestamped folder, so an earlier workbook left open in Excel cannot block or overwrite the next build. The build stops and removes partially generated workbooks if a VBA module cannot be installed or the control center cannot run. Its detailed log remains in that run's folder as `ONGEN-build.log`, including the original failure rather than a generic replacement error.

## What the installer changes

- Replaces `mod09_GeneratePD2` with a Phase 3 bounded stale-row cleanup to avoid whole-column work.
- Replaces `mod14_ImportPD6Power` so Phase 4D failures propagate to the caller and cannot be reported as successful.
- Replaces `mod20_DeviceWorkflow` with the Phase 5A row-counter correction that prevents existing output rows from being overwritten.
- Replaces the Field Devices `EXPORT_TO_IOGEN` module with validated headers, composite-key matching, bulk transfer, and count checks.
- Adds `modONGEN_SuiteControl`, an Excel worksheet control center with preflight checks, guarded phase buttons, error reporting, and an `ONGEN RUN LOG` sheet.

The supporting IO Generator, EAS, function library, and function-code lookup files are copied beside the generated master workbook because the existing VBA workflow depends on companion workbooks.

## Controlled limitation

`ONGEN Power Calculation V4.6.xlsm` could not be included because the uploaded file exceeded the available ingestion limit. Phase 4 code was audited and patched where possible, but the Power workbook and the full Phase 4 round trip were not runtime tested. The control center shows this limitation and does not claim Phase 4 validation.

Microsoft Excel and the VBA runtime are unavailable in the Linux build environment used to prepare this package. Static checks and workbook-container validation are included, while final VBA compilation and workflow acceptance must be performed in Windows Excel. In the VBA editor, use **Debug > Compile VBAProject**, then execute the phases against a controlled project copy and compare row counts and engineering totals.

See `docs\ONGEN_Audit_Report.pdf` and `docs\ONGEN_Engineering_Review.xlsx` for the evidence, findings, and change record.
