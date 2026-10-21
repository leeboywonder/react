# ONGEN Engineering Suite

## Excel VBA delivery (recommended)

`ONGEN_Excel_VBA_0.2.zip` is the Excel-first engineering delivery. It contains the supplied macro-enabled workbooks, reviewed VBA patches, an Excel worksheet control center, audit documents, and a Windows builder. Extract the ZIP and run `Start-Build.cmd` on a Windows computer with desktop Microsoft Excel.

The browser/desktop package remains available below as an earlier review artifact. It is not the primary Excel engineering deliverable.

## Browser review artifact 0.1.0

This branch contains the ONGEN engineering-preview delivery requested for the
PD Definition, Field Devices, IO Generator, EAS, and function-library workflow.

## Downloads

- `ONGEN_Suite_0.1.0.zip` — complete browser application, Windows desktop
  launcher/build source, audit evidence, proposed VBA module replacements,
  seeded project, and generated review outputs.
- `ONGEN_Audit_Report.pdf` — full workbook/code audit, changes, verified
  results, limitations, and Windows/Excel acceptance plan.
- `ONGEN_Engineering_Review.xlsx` — generated Phase 3, function-device,
  Phase 4 power-exchange, and issue-review workbook.

## Verification

| File | SHA-256 |
|---|---|
| `ONGEN_Excel_VBA_0.2.zip` | `a991096ee0734bd856d128fa19fc18db0b7ea23f5bd3caa96e9aa52aa568021c` |
| `ONGEN_Suite_0.1.0.zip` | `6b4f67c93911827d08d8f96cff19892192a5bab725f30b14e9b95b9611419dac` |
| `ONGEN_Audit_Report.pdf` | `71d70e8531543d4edc1d4f49ba22a9a92ae848a5ceae5041218c62f3381191da` |
| `ONGEN_Engineering_Review.xlsx` | `538f2b3c5e0abc20981f841b9f2e6f3985e370594f2729a0c7851e69daa3f482` |

The ZIP and XLSX passed archive-integrity checks before upload. The clean
packaged suite passed 39 Python engine/API/library tests and restored the
seeded six-source project successfully. Nine browser workflow checks passed
in Linux Chromium. Proposed VBA patches passed model/static checks only.

## Validation limits

This is an engineering preview. The original accessible workbooks remain
unchanged. Microsoft Excel/VBA and the native Windows desktop build were not
executed in the audit environment. Power Calculation V4.6 exceeded the
attachment transfer limit and remains uninspected. The generated IO packet is
blocked pending resolution of duplicate legacy device identities; this is
recorded in the workbook and audit report. Do not treat generated outputs as
an approved electrical design without the Windows/Excel and engineering
acceptance steps in the report.
