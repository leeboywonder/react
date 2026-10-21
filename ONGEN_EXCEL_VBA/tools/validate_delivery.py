#!/usr/bin/env python3
"""Static validation for the ONGEN Excel VBA delivery package."""

from pathlib import Path
import re
import sys
import zipfile

ROOT = Path(__file__).resolve().parents[1]

WORKBOOKS = [
    "ONGEN PD_DEFINITION_MASTER_WORKFLOW_V5.1.6.5.xlsm",
    "ONGEN Field Devices Specification -v2.1.7.1.xlsm",
    "ONGEN IO GENERATOR - B05.xlsm",
    "ONGEN EAS - V2.5.3.xlsm",
    "ONGEN_EAS_MASTER_FUNCTION_LIBRARY.xlsm",
]

MODULES = {
    "vba/master/mod09_GeneratePD2.bas": "mod09_GeneratePD2",
    "vba/master/mod14_ImportPD6Power.bas": "mod14_ImportPD6Power",
    "vba/master/mod20_DeviceWorkflow.bas": "mod20_DeviceWorkflow",
    "vba/master/modONGEN_SuiteControl.bas": "modONGEN_SuiteControl",
    "vba/field/EXPORT_TO_IOGEN.bas": "EXPORT_TO_IOGEN",
}


def fail(message: str) -> None:
    print(f"FAIL: {message}")
    raise SystemExit(1)


for name in WORKBOOKS:
    path = ROOT / "source" / name
    if not path.is_file():
        fail(f"missing workbook: {name}")
    with zipfile.ZipFile(path) as archive:
        names = set(archive.namelist())
        if "xl/vbaProject.bin" not in names:
            fail(f"workbook has no VBA project: {name}")
        if "[Content_Types].xml" not in names or "xl/workbook.xml" not in names:
            fail(f"invalid Excel package: {name}")

for relative, expected_name in MODULES.items():
    path = ROOT / relative
    if not path.is_file():
        fail(f"missing VBA module: {relative}")
    text = path.read_text(encoding="latin-1")
    match = re.search(r'^Attribute VB_Name = "([^"]+)"', text, re.MULTILINE)
    if not match or match.group(1) != expected_name:
        fail(f"wrong VBA module name in {relative}")
    if "Option Explicit" not in text:
        fail(f"Option Explicit missing from {relative}")

control = (ROOT / "vba/master/modONGEN_SuiteControl.bas").read_text(encoding="latin-1")
for procedure in (
    "ONGEN_InstallControlCenter",
    "ONGEN_SelfCheck",
    "ONGEN_RunGeneratePD2",
    "ONGEN_RunPower4D",
    "ONGEN_RunDevice5A",
):
    if not re.search(rf"Public Sub {procedure}\b", control):
        fail(f"control-center procedure missing: {procedure}")

builder = (ROOT / "Build-ONGEN-Excel.ps1").read_text(encoding="utf-8")
for module in ("mod09_GeneratePD2.bas", "mod14_ImportPD6Power.bas", "mod20_DeviceWorkflow.bas", "modONGEN_SuiteControl.bas", "EXPORT_TO_IOGEN.bas"):
    if module not in builder:
        fail(f"builder does not install: {module}")

print(f"PASS: {len(WORKBOOKS)} macro-enabled workbooks and {len(MODULES)} VBA modules validated")
