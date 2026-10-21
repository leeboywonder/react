Attribute VB_Name = "mod20_DeviceWorkflow"
Option Explicit

Private Const SH_FUNC As String = "FUNCTION DEVICE pd2"
Private Const SH_SPEC As String = "Specification Devices"
Private Const SH_FIELD As String = "Field Device IO ONGEN"
Private Const SH_PD6 As String = "PD6 System Designed"
Private Const SH_PD2 As String = "SIS_TAG_IMPORT_PD2"
Private Const SH_FUNCDEF As String = "FUNCTION DEFINITION"
Private Const SH_CAD As String = "SIS_TAG_EXPORT"

'====================================================================
' STEP 5A - one entry / one bulk output
'   1) Rebuild Specification Devices from FUNCTION DEVICE pd2.
'   2) Resolve the parent conveyor SIS_SEQUENCE.
'      Example: Device Name 1060-01B-B + Item Type B -> 1060-01B.
'   3) Apply final AREA_NUMBER / SWITCHGROUP / PANEL from PD6.
'   4) INPUT / OUTPUT / S_IN / S_OUT come from FUNCTION DEVICE pd2 when supplied.
'      Existing Specification values are only used when the generated source is blank.
'   5) Preserve existing Standard Block / Safety Block engineering fields.
'====================================================================
Public Sub DeviceWF_5A_Update_Specification_FAST()
    If PDPerf_IsYielding Then Exit Sub
    Dim wsFunc As Worksheet, wsSpec As Worksheet, wsPD6 As Worksheet
    Dim wsPD2 As Worksheet, wsFuncDef As Worksheet
    Dim fData As Variant, oldData As Variant, pd2Data As Variant
    Dim fMap As Object, oldMap As Object, pd2Map As Object
    Dim pdAssign As Object, oldEng As Object, seen As Object, vfdCFIO As Object
    Dim fLR As Long, fLC As Long, oldLR As Long, oldLC As Long, pd2LR As Long, pd2LC As Long
    Dim cDev As Long, cBase As Long, cType As Long, cArea As Long
    Dim cPD2Seq As Long, cPD2Type As Long, cPD2CFN As Long
    Dim oldDev As Long, oldType As Long, r As Long, outRow As Long, targetRow As Long
    Dim normalCount As Long, qualifyingVFD As Long, maxRows As Long
    Dim deviceName As String, itemType As String, baseSeq As String, k As String, cfn As String
    Dim a As Variant, eng As Variant, ioVal As Variant
    Dim outData() As Variant, finalData() As Variant, hdr(1 To 1, 1 To 12) As Variant
    Dim oldCalc As XlCalculation, oldScreen As Boolean, oldEvents As Boolean
    Dim matchedPD6 As Long, missingPD6 As Long, matchedPanel As Long, missingPanel As Long, preserved As Long
    Dim ioFromPD2 As Long, ioFallbackOld As Long, vfdAdded As Long, vfdUpdated As Long, vfdSkipped As Long

    Set wsFunc = DW_Sheet(ThisWorkbook, SH_FUNC)
    If wsFunc Is Nothing Then MsgBox "Missing '" & SH_FUNC & "'.", vbCritical: Exit Sub
    Set wsPD6 = DW_Sheet(ThisWorkbook, SH_PD6)
    If wsPD6 Is Nothing Then MsgBox "Missing '" & SH_PD6 & "'. Run Step 4D first.", vbCritical: Exit Sub
    Set wsPD2 = DW_Sheet(ThisWorkbook, SH_PD2)
    If wsPD2 Is Nothing Then MsgBox "Missing '" & SH_PD2 & "'. Run the PD2 generation step first.", vbCritical: Exit Sub
    Set wsFuncDef = DW_Sheet(ThisWorkbook, SH_FUNCDEF)
    If wsFuncDef Is Nothing Then MsgBox "Missing '" & SH_FUNCDEF & "'.", vbCritical: Exit Sub
    Set wsSpec = DW_GetOrCreateSheet(ThisWorkbook, SH_SPEC)

    oldScreen = Application.ScreenUpdating: oldEvents = Application.EnableEvents: oldCalc = Application.Calculation
    Application.ScreenUpdating = False: Application.EnableEvents = False: Application.Calculation = xlCalculationManual
    Application.StatusBar = "Step 5A: rebuilding Specification Devices, including VFD control functions..."
    On Error GoTo Fail

    'Preserve existing engineering values before rebuilding the worksheet.
    Set oldEng = CreateObject("Scripting.Dictionary"): oldEng.CompareMode = vbTextCompare
    oldLR = DW_LastRow(wsSpec): oldLC = DW_LastCol(wsSpec)
    If oldLR >= 2 And oldLC >= 1 Then
        oldData = wsSpec.Range(wsSpec.Cells(1, 1), wsSpec.Cells(oldLR, oldLC)).Value2
        Set oldMap = DW_HeaderMap(oldData, oldLC)
        oldDev = DW_ColAny(oldMap, Array("DEVICE NAME", "SIS_SEQUENCE"))
        oldType = DW_ColAny(oldMap, Array("ITEM TYPE", "TYPE"))
        If oldDev > 0 Then
            For r = 2 To oldLR
                PDPerf_Pulse "mod20_DeviceWorkflow.DeviceWF_5A_Update_Specification_FAST [L64]" & " row " & CStr(r)
                deviceName = DW_Text(oldData(r, oldDev))
                If oldType > 0 Then itemType = DW_Text(oldData(r, oldType)) Else itemType = vbNullString
                If Len(deviceName) > 0 Then
                    k = UCase$(deviceName & "|" & itemType)
                    oldEng(k) = Array( _
                        DW_Arr(oldData, r, oldMap, "INPUT"), DW_Arr(oldData, r, oldMap, "OUTPUT"), _
                        DW_Arr(oldData, r, oldMap, "S_IN"), DW_Arr(oldData, r, oldMap, "S_OUT"), _
                        DW_Arr(oldData, r, oldMap, "Standard Block"), DW_Arr(oldData, r, oldMap, "Safety Block"), _
                        DW_Arr(oldData, r, oldMap, "AREA_NUMBER"), DW_Arr(oldData, r, oldMap, "SWITCHGROUP"), _
                        DW_Arr(oldData, r, oldMap, "PANEL"))
                End If
            Next r
        End If
    End If

    Set pdAssign = DW_BuildPD6AssignmentMap(wsPD6)
    If DW_CountUsableAssignment(pdAssign, 2) = 0 Then
        Err.Raise vbObjectError + 6104, , _
                  "PD6 System Designed contains no usable Panel assignments. Step 4C/4D must complete before Step 5A."
    End If
    Set vfdCFIO = DW_BuildControlFunctionIOMap(wsFuncDef)

    'Load normal function-device rows.
    fLR = DW_LastRow(wsFunc): fLC = DW_LastCol(wsFunc)
    If fLR < 2 Or fLC < 1 Then Err.Raise vbObjectError + 6101, , SH_FUNC & " contains no device rows."
    fData = wsFunc.Range(wsFunc.Cells(1, 1), wsFunc.Cells(fLR, fLC)).Value2
    Set fMap = DW_HeaderMap(fData, fLC)
    cDev = DW_ColAny(fMap, Array("SIS_SEQUENCE", "DEVICE NAME"))
    cBase = DW_ColAny(fMap, Array("SEQUENCE", "PARENT SIS_SEQUENCE", "BASE SIS_SEQUENCE"))
    cType = DW_ColAny(fMap, Array("CONTROL_FUNCTION_TAG", "ITEM TYPE", "TYPE"))
    cArea = DW_ColAny(fMap, Array("AREA_NUMBER", "AREA"))
    If cDev = 0 Then Err.Raise vbObjectError + 6102, , SH_FUNC & " is missing SIS_SEQUENCE / Device Name."
    For r = 2 To fLR
        PDPerf_Pulse "mod20_DeviceWorkflow.DeviceWF_5A_Update_Specification_FAST [L97]" & " row " & CStr(r)
        If Len(DW_Text(fData(r, cDev))) > 0 Then normalCount = normalCount + 1
    Next r

    'Load VFD parents from SIS_TAG_IMPORT_PD2.
    pd2LR = DW_LastRow(wsPD2): pd2LC = DW_LastCol(wsPD2)
    If pd2LR >= 2 And pd2LC >= 1 Then
        pd2Data = wsPD2.Range(wsPD2.Cells(1, 1), wsPD2.Cells(pd2LR, pd2LC)).Value2
        Set pd2Map = DW_HeaderMap(pd2Data, pd2LC)
        cPD2Seq = DW_ColAny(pd2Map, Array("SIS_SEQUENCE", "SEQUENCE", "DEVICE NAME"))
        cPD2Type = DW_ColAny(pd2Map, Array("TYPE", "ITEM TYPE"))
        cPD2CFN = DW_ColAny(pd2Map, Array("CONTROL_FUNCTION_NUMBER", "CONTROL FUNCTION NUMBER", "CONTROLFUNCTIONNUMBER"))
        If cPD2Seq = 0 Then Err.Raise vbObjectError + 6110, , SH_PD2 & " is missing SIS_SEQUENCE."
        If cPD2Type = 0 Then Err.Raise vbObjectError + 6111, , SH_PD2 & " is missing TYPE."
        If cPD2CFN = 0 Then Err.Raise vbObjectError + 6112, , SH_PD2 & " is missing CONTROL_FUNCTION_NUMBER."
        For r = 2 To pd2LR
            PDPerf_Pulse "mod20_DeviceWorkflow.DeviceWF_5A_Update_Specification_FAST [L112]" & " row " & CStr(r)
            If StrComp(DW_Text(pd2Data(r, cPD2Type)), "VFD", vbTextCompare) = 0 _
               And DW_Usable(pd2Data(r, cPD2CFN)) _
               And DW_Usable(pd2Data(r, cPD2Seq)) Then qualifyingVFD = qualifyingVFD + 1
        Next r
    End If

    maxRows = normalCount + qualifyingVFD
    If maxRows = 0 Then Err.Raise vbObjectError + 6103, , "No Specification Device rows were found."

    hdr(1, 1) = "DEVICE NAME": hdr(1, 2) = "ITEM TYPE": hdr(1, 3) = "SIS_SEQUENCE"
    hdr(1, 4) = "AREA_NUMBER": hdr(1, 5) = "SWITCHGROUP": hdr(1, 6) = "PANEL"
    hdr(1, 7) = "INPUT": hdr(1, 8) = "OUTPUT": hdr(1, 9) = "S_IN": hdr(1, 10) = "S_OUT"
    hdr(1, 11) = "Standard Block": hdr(1, 12) = "Safety Block"
    ReDim outData(1 To maxRows, 1 To 12)
    Set seen = CreateObject("Scripting.Dictionary"): seen.CompareMode = vbTextCompare

    'Pass 1: retain the original FUNCTION DEVICE pd2 behavior.
    For r = 2 To fLR
        PDPerf_Pulse "mod20_DeviceWorkflow.DeviceWF_5A_Update_Specification_FAST [L130]" & " row " & CStr(r)
        deviceName = DW_Text(fData(r, cDev))
        If Len(deviceName) > 0 Then
            outRow = outRow + 1
            If cType > 0 Then itemType = DW_Text(fData(r, cType)) Else itemType = vbNullString
            If cBase > 0 Then baseSeq = DW_Text(fData(r, cBase)) Else baseSeq = vbNullString
            If Len(baseSeq) = 0 Then baseSeq = DW_BaseSequence(deviceName, itemType)

            outData(outRow, 1) = deviceName
            outData(outRow, 2) = itemType
            outData(outRow, 3) = baseSeq

            k = UCase$(deviceName & "|" & itemType)
            If oldEng.Exists(k) Then
                eng = oldEng(k)
            Else
                eng = Array(vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, _
                            vbNullString, vbNullString, vbNullString)
            End If

            Call DW_FillSpecAssignment(outData, outRow, baseSeq, pdAssign, fData, r, fMap, eng, cArea, _
                                       matchedPD6, missingPD6, matchedPanel, missingPanel)

            ioVal = DW_Arr(fData, r, fMap, "INPUT")
            If DW_Usable(ioVal) Then outData(outRow, 7) = ioVal: ioFromPD2 = ioFromPD2 + 1 Else outData(outRow, 7) = eng(0): If DW_Usable(eng(0)) Then ioFallbackOld = ioFallbackOld + 1
            ioVal = DW_Arr(fData, r, fMap, "OUTPUT")
            If DW_Usable(ioVal) Then outData(outRow, 8) = ioVal: ioFromPD2 = ioFromPD2 + 1 Else outData(outRow, 8) = eng(1): If DW_Usable(eng(1)) Then ioFallbackOld = ioFallbackOld + 1
            ioVal = DW_Arr(fData, r, fMap, "S_IN")
            If DW_Usable(ioVal) Then outData(outRow, 9) = ioVal: ioFromPD2 = ioFromPD2 + 1 Else outData(outRow, 9) = eng(2): If DW_Usable(eng(2)) Then ioFallbackOld = ioFallbackOld + 1
            ioVal = DW_Arr(fData, r, fMap, "S_OUT")
            If DW_Usable(ioVal) Then outData(outRow, 10) = ioVal: ioFromPD2 = ioFromPD2 + 1 Else outData(outRow, 10) = eng(3): If DW_Usable(eng(3)) Then ioFallbackOld = ioFallbackOld + 1

            outData(outRow, 11) = eng(4): outData(outRow, 12) = eng(5)
            If oldEng.Exists(k) Then preserved = preserved + 1
            seen(DW_DeviceKey(deviceName, itemType)) = outRow
        End If
    Next r

    'Pass 2: add/update parent VFD rows from SIS_TAG_IMPORT_PD2.
    'ONGEN audit fix: outRow remains the total; targetRow identifies an existing VFD.
    'Eligibility: TYPE=VFD AND CONTROL_FUNCTION_NUMBER is populated.
    For r = 2 To pd2LR
        PDPerf_Pulse "mod20_DeviceWorkflow.DeviceWF_5A_Update_Specification_FAST [L170]" & " row " & CStr(r)
        If StrComp(DW_Text(pd2Data(r, cPD2Type)), "VFD", vbTextCompare) = 0 _
           And DW_Usable(pd2Data(r, cPD2CFN)) Then

            baseSeq = DW_Text(pd2Data(r, cPD2Seq))
            cfn = DW_Text(pd2Data(r, cPD2CFN))
            If Len(baseSeq) = 0 Then
                vfdSkipped = vfdSkipped + 1
            Else
                deviceName = baseSeq
                k = DW_DeviceKey(deviceName, "VFD")

                If seen.Exists(k) Then
                    targetRow = CLng(seen(k))
                    vfdUpdated = vfdUpdated + 1
                Else
                    outRow = outRow + 1
                    targetRow = outRow
                    outData(targetRow, 1) = deviceName
                    outData(targetRow, 2) = "VFD"
                    outData(targetRow, 3) = baseSeq
                    seen(k) = targetRow
                    vfdAdded = vfdAdded + 1
                End If

                'Final assignment priority: PD6, then values already carried on PD2.
                If pdAssign.Exists(UCase$(baseSeq)) Then
                    a = pdAssign(UCase$(baseSeq))
                    outData(targetRow, 4) = a(0): outData(targetRow, 5) = a(1): outData(targetRow, 6) = a(2)
                Else
                    outData(targetRow, 4) = DW_Arr(pd2Data, r, pd2Map, "AREA_NUMBER")
                    outData(targetRow, 5) = DW_Arr(pd2Data, r, pd2Map, "SWITCHGROUP")
                    outData(targetRow, 6) = DW_Arr(pd2Data, r, pd2Map, "PANEL")
                    If Not DW_Usable(outData(targetRow, 6)) Then outData(targetRow, 6) = DW_Arr(pd2Data, r, pd2Map, "PNL")
                End If

                'VFD IO priority: explicit PD2 value, then FUNCTION DEFINITION by CONTROL_FUNCTION_NUMBER.
                outData(targetRow, 7) = DW_ResolveVFDIO(pd2Data, r, pd2Map, cfn, "INPUT", vfdCFIO)
                outData(targetRow, 8) = DW_ResolveVFDIO(pd2Data, r, pd2Map, cfn, "OUTPUT", vfdCFIO)
                outData(targetRow, 9) = DW_ResolveVFDIO(pd2Data, r, pd2Map, cfn, "S_IN", vfdCFIO)
                outData(targetRow, 10) = DW_ResolveVFDIO(pd2Data, r, pd2Map, cfn, "S_OUT", vfdCFIO)

                'Preserve existing block selections if this VFD already existed in Specification Devices.
                k = UCase$(deviceName & "|VFD")
                If oldEng.Exists(k) Then
                    eng = oldEng(k)
                    outData(targetRow, 11) = eng(4)
                    outData(targetRow, 12) = eng(5)
                End If
            End If
        End If
    Next r

    'Trim the oversized work array after de-duplication.
    ReDim finalData(1 To outRow, 1 To 12)
    Dim rr As Long, cc As Long
    For rr = 1 To outRow
        PDPerf_Pulse "mod20_DeviceWorkflow.DeviceWF_5A_Update_Specification_FAST [L225]" & " row " & CStr(rr)
        For cc = 1 To 12
            finalData(rr, cc) = outData(rr, cc)
        Next cc
    Next rr

    DW_ClearUsed wsSpec
    wsSpec.Cells(1, 1).Resize(1, 12).Value2 = hdr
    wsSpec.Cells(2, 1).Resize(outRow, 12).Value2 = finalData
    DW_FormatHeader wsSpec, 12

    Application.StatusBar = False
    Application.ScreenUpdating = oldScreen: Application.EnableEvents = oldEvents: Application.Calculation = oldCalc
    MsgBox "Step 5A complete." & vbCrLf & vbCrLf & _
           "Specification rows rebuilt: " & Format$(outRow, "#,##0") & vbCrLf & _
           "Qualifying VFD rows found in SIS_TAG_IMPORT_PD2: " & Format$(qualifyingVFD, "#,##0") & vbCrLf & _
           "VFD rows added to Specification Devices: " & Format$(vfdAdded, "#,##0") & vbCrLf & _
           "Existing VFD rows refreshed: " & Format$(vfdUpdated, "#,##0") & vbCrLf & _
           "VFD rows skipped for missing SIS_SEQUENCE: " & Format$(vfdSkipped, "#,##0") & vbCrLf & vbCrLf & _
           "VFD rule: TYPE=VFD + populated CONTROL_FUNCTION_NUMBER." & vbCrLf & _
           "VFD IO: SIS_TAG_IMPORT_PD2 first, then FUNCTION DEFINITION by CONTROL_FUNCTION_NUMBER.", _
           vbInformation, "Phase 5A - Specification + VFD"
    Exit Sub
Fail:
    Application.StatusBar = False
    Application.ScreenUpdating = oldScreen: Application.EnableEvents = oldEvents: Application.Calculation = oldCalc
    MsgBox "Step 5A failed. Error " & Err.number & ": " & Err.description, vbCritical, "Phase 5A"
End Sub

Private Sub DW_FillSpecAssignment(ByRef outData As Variant, ByVal outRow As Long, ByVal baseSeq As String, _
                                  ByVal pdAssign As Object, ByRef fData As Variant, ByVal r As Long, _
                                  ByVal fMap As Object, ByRef eng As Variant, ByVal cArea As Long, _
                                  ByRef matchedPD6 As Long, ByRef missingPD6 As Long, _
                                  ByRef matchedPanel As Long, ByRef missingPanel As Long)
    Dim a As Variant
    If pdAssign.Exists(UCase$(baseSeq)) Then
        a = pdAssign(UCase$(baseSeq))
        outData(outRow, 4) = a(0)
        outData(outRow, 5) = a(1)
        outData(outRow, 6) = a(2)
        matchedPD6 = matchedPD6 + 1
        If DW_Usable(a(2)) Then matchedPanel = matchedPanel + 1 Else missingPanel = missingPanel + 1
    Else
        If cArea > 0 And DW_Usable(fData(r, cArea)) Then outData(outRow, 4) = fData(r, cArea) Else outData(outRow, 4) = eng(6)
        outData(outRow, 5) = eng(7)
        outData(outRow, 6) = eng(8)
        missingPD6 = missingPD6 + 1
        If DW_Usable(outData(outRow, 6)) Then matchedPanel = matchedPanel + 1 Else missingPanel = missingPanel + 1
    End If
End Sub

'====================================================================
' STEP 5B - one entry / one bulk output
' Build the B03 input worksheet from the already-finalized Specification
' Devices sheet. No file picker and no external macro are called here.
'====================================================================
Public Sub DeviceWF_5B_Build_FieldIO_FAST()
    If PDPerf_IsYielding Then Exit Sub
    Dim wsSpec As Worksheet, wsField As Worksheet, wsPD6 As Worksheet, wsPD2 As Worksheet, wsFuncDef As Worksheet
    Dim data As Variant, m As Object, pdAssign As Object
    Dim pd2Data As Variant, pd2Map As Object, seen As Object, vfdCFIO As Object
    Dim outData() As Variant, finalData() As Variant, hdr(1 To 1, 1 To 12) As Variant
    Dim lr As Long, lc As Long, r As Long, outRow As Long, rowCount As Long, rr As Long, cc As Long
    Dim pd2LR As Long, pd2LC As Long, qualifyingVFD As Long
    Dim cDev As Long, cType As Long, cBase As Long
    Dim cPD2Seq As Long, cPD2Type As Long, cPD2CFN As Long
    Dim baseSeq As String, deviceName As String, itemType As String, cfn As String, k As String
    Dim a As Variant
    Dim pd6Matched As Long, pd6Missing As Long, panelFilled As Long, panelMissing As Long
    Dim vfdAdded As Long, vfdAlreadyPresent As Long, vfdSkippedNoSeq As Long
    Dim oldCalc As XlCalculation, oldScreen As Boolean, oldEvents As Boolean
    Dim cadHandleMap As Object

    Set wsSpec = DW_Sheet(ThisWorkbook, SH_SPEC)
    If wsSpec Is Nothing Then MsgBox "Missing '" & SH_SPEC & "'. Run Step 5A first.", vbCritical: Exit Sub
    Set wsField = DW_GetOrCreateSheet(ThisWorkbook, SH_FIELD)
    Set wsPD6 = DW_Sheet(ThisWorkbook, SH_PD6)
    If wsPD6 Is Nothing Then MsgBox "Missing '" & SH_PD6 & "'. Run Step 4D before Step 5B.", vbCritical: Exit Sub

    'SIS_TAG_IMPORT_PD2 is optional for legacy projects, but when present it is the
    'authoritative source for parent VFD rows that carry a CONTROL_FUNCTION_NUMBER.
    Set wsPD2 = DW_Sheet(ThisWorkbook, SH_PD2)
    Set wsFuncDef = DW_Sheet(ThisWorkbook, SH_FUNCDEF)
    If wsFuncDef Is Nothing Then
        MsgBox "Missing '" & SH_FUNCDEF & "'. VFD CONTROL_FUNCTION_NUMBER IO cannot be resolved.", vbCritical
        Exit Sub
    End If
    Set vfdCFIO = DW_BuildControlFunctionIOMap(wsFuncDef)
    Set cadHandleMap = DW_BuildCADHandleMap()


    oldScreen = Application.ScreenUpdating: oldEvents = Application.EnableEvents: oldCalc = Application.Calculation
    Application.ScreenUpdating = False: Application.EnableEvents = False: Application.Calculation = xlCalculationManual
    On Error GoTo Fail

    lr = DW_LastRow(wsSpec): lc = DW_LastCol(wsSpec)
    If lr < 2 Or lc < 1 Then Err.Raise vbObjectError + 6201, , SH_SPEC & " contains no rows."
    data = wsSpec.Range(wsSpec.Cells(1, 1), wsSpec.Cells(lr, lc)).Value2
    Set m = DW_HeaderMap(data, lc)
    cDev = DW_ColAny(m, Array("DEVICE NAME", "SIS_SEQUENCE"))
    cType = DW_ColAny(m, Array("ITEM TYPE", "TYPE"))
    cBase = DW_ColAny(m, Array("SIS_SEQUENCE", "SEQUENCE", "PARENT SIS_SEQUENCE"))
    If cDev = 0 Then Err.Raise vbObjectError + 6202, , "Specification Devices is missing DEVICE NAME."
    If cBase = 0 Then Err.Raise vbObjectError + 6204, , "Specification Devices is missing parent SIS_SEQUENCE."

    Set pdAssign = DW_BuildPD6AssignmentMap(wsPD6)
    If DW_CountUsableAssignment(pdAssign, 2) = 0 Then
        Err.Raise vbObjectError + 6205, , _
                  "PD6 System Designed contains no usable Panel assignments. Run Step 4C, Step 4D, and Step 5A before Step 5B."
    End If

    'Count the normal Specification rows first.
    For r = 2 To lr
        PDPerf_Pulse "mod20_DeviceWorkflow.DeviceWF_5B_Build_FieldIO_FAST [L336]" & " row " & CStr(r)
        If Len(DW_Text(data(r, cDev))) > 0 Then rowCount = rowCount + 1
    Next r

    'Count qualifying VFD rows from SIS_TAG_IMPORT_PD2.
    'Rule: TYPE = VFD AND CONTROL_FUNCTION_NUMBER is populated.
    If Not wsPD2 Is Nothing Then
        pd2LR = DW_LastRow(wsPD2): pd2LC = DW_LastCol(wsPD2)
        If pd2LR >= 2 And pd2LC >= 1 Then
            pd2Data = wsPD2.Range(wsPD2.Cells(1, 1), wsPD2.Cells(pd2LR, pd2LC)).Value2
            Set pd2Map = DW_HeaderMap(pd2Data, pd2LC)
            cPD2Seq = DW_ColAny(pd2Map, Array("SIS_SEQUENCE", "SEQUENCE", "DEVICE NAME"))
            cPD2Type = DW_ColAny(pd2Map, Array("TYPE", "ITEM TYPE"))
            cPD2CFN = DW_ColAny(pd2Map, Array("CONTROL_FUNCTION_NUMBER", "CONTROL FUNCTION NUMBER", "CONTROLFUNCTIONNUMBER"))

            If cPD2Seq = 0 Then Err.Raise vbObjectError + 6206, , SH_PD2 & " is missing SIS_SEQUENCE."
            If cPD2Type = 0 Then Err.Raise vbObjectError + 6207, , SH_PD2 & " is missing TYPE."
            If cPD2CFN = 0 Then Err.Raise vbObjectError + 6208, , SH_PD2 & " is missing CONTROL_FUNCTION_NUMBER."

            For r = 2 To pd2LR
                PDPerf_Pulse "mod20_DeviceWorkflow.DeviceWF_5B_Build_FieldIO_FAST [L355]" & " row " & CStr(r)
                If StrComp(DW_Text(pd2Data(r, cPD2Type)), "VFD", vbTextCompare) = 0 Then
                    If DW_Usable(pd2Data(r, cPD2CFN)) Then qualifyingVFD = qualifyingVFD + 1
                End If
            Next r
        End If
    End If

    If rowCount + qualifyingVFD = 0 Then Err.Raise vbObjectError + 6203, , "No device rows were found for Field Device IO ONGEN."
    ReDim outData(1 To rowCount + qualifyingVFD, 1 To 12)
    Set seen = CreateObject("Scripting.Dictionary"): seen.CompareMode = vbTextCompare

    hdr(1, 1) = "DEVICE NAME": hdr(1, 2) = "ITEM TYPE": hdr(1, 3) = "INPUT": hdr(1, 4) = "OUTPUT"
    hdr(1, 5) = "S_IN": hdr(1, 6) = "S_OUT": hdr(1, 7) = "PNL": hdr(1, 8) = "PWG"
    hdr(1, 9) = "LOCATION": hdr(1, 10) = "Standard Block": hdr(1, 11) = "Safety Block": hdr(1, 12) = "HANDLE"

    Application.StatusBar = "Step 5B: building B03 input, final PD6 assignments, and PD2 VFD control-function IO..."

    '----------------------------------------------------------------
    'Pass 1 - existing Specification Devices behavior.
    '----------------------------------------------------------------
    For r = 2 To lr
        PDPerf_Pulse "mod20_DeviceWorkflow.DeviceWF_5B_Build_FieldIO_FAST [L376]" & " row " & CStr(r)
        deviceName = DW_Text(data(r, cDev))
        If Len(deviceName) > 0 Then
            outRow = outRow + 1
            If cType > 0 Then itemType = DW_Text(data(r, cType)) Else itemType = vbNullString

            outData(outRow, 1) = data(r, cDev)
            outData(outRow, 2) = itemType
            outData(outRow, 3) = DW_Arr(data, r, m, "INPUT")
            outData(outRow, 4) = DW_Arr(data, r, m, "OUTPUT")
            outData(outRow, 5) = DW_Arr(data, r, m, "S_IN")
            outData(outRow, 6) = DW_Arr(data, r, m, "S_OUT")

            baseSeq = DW_Text(data(r, cBase))
            If Len(baseSeq) = 0 Then baseSeq = DW_BaseSequence(deviceName, itemType)

            If pdAssign.Exists(UCase$(baseSeq)) Then
                a = pdAssign(UCase$(baseSeq))
                outData(outRow, 7) = a(2)
                outData(outRow, 8) = a(1)
                outData(outRow, 9) = a(0)
                pd6Matched = pd6Matched + 1
                If DW_Usable(a(2)) Then panelFilled = panelFilled + 1 Else panelMissing = panelMissing + 1
            Else
                outData(outRow, 7) = DW_Arr(data, r, m, "PANEL")
                outData(outRow, 8) = DW_Arr(data, r, m, "SWITCHGROUP")
                outData(outRow, 9) = DW_Arr(data, r, m, "AREA_NUMBER")
                If DW_Usable(outData(outRow, 7)) Then panelFilled = panelFilled + 1 Else panelMissing = panelMissing + 1
                pd6Missing = pd6Missing + 1
            End If

            outData(outRow, 10) = DW_Arr(data, r, m, "Standard Block")
            outData(outRow, 11) = DW_Arr(data, r, m, "Safety Block")
            outData(outRow, 12) = DW_FindCADHandle(DW_Text(outData(outRow, 1)), cadHandleMap)

            k = DW_DeviceKey(deviceName, itemType)
            If Len(k) > 0 Then seen(k) = True
        End If
    Next r

    '----------------------------------------------------------------
    'Pass 2 - append VFD parent rows from SIS_TAG_IMPORT_PD2.
    'A VFD is eligible only when CONTROL_FUNCTION_NUMBER is usable.
    'IO priority: SIS_TAG_IMPORT_PD2 first; blank values fall back to FUNCTION DEFINITION by CONTROL_FUNCTION_NUMBER.
    '----------------------------------------------------------------
    If qualifyingVFD > 0 Then
        For r = 2 To pd2LR
            PDPerf_Pulse "mod20_DeviceWorkflow.DeviceWF_5B_Build_FieldIO_FAST [L422]" & " row " & CStr(r)
            itemType = DW_Text(pd2Data(r, cPD2Type))
            cfn = DW_Text(pd2Data(r, cPD2CFN))

            If StrComp(itemType, "VFD", vbTextCompare) = 0 And DW_Usable(cfn) Then
                baseSeq = DW_Text(pd2Data(r, cPD2Seq))
                If Len(baseSeq) = 0 Then
                    vfdSkippedNoSeq = vfdSkippedNoSeq + 1
                Else
                    'Use the PD2 SIS_SEQUENCE as the VFD device identifier. Do not append
                    'another -VFD suffix because TYPE already identifies the B03 device class.
                    deviceName = baseSeq
                    k = DW_DeviceKey(deviceName, "VFD")

                    If seen.Exists(k) Then
                        'If a VFD already came through Specification Devices, refresh its IO
                        'from authoritative SIS_TAG_IMPORT_PD2 instead of creating a duplicate.
                        Call DW_RefreshExistingVFDRow(outData, outRow, deviceName, pd2Data, r, pd2Map, pdAssign, baseSeq, cfn, vfdCFIO)
                        vfdAlreadyPresent = vfdAlreadyPresent + 1
                    Else
                        outRow = outRow + 1
                        outData(outRow, 1) = deviceName
                        outData(outRow, 2) = "VFD"
                        outData(outRow, 3) = DW_ResolveVFDIO(pd2Data, r, pd2Map, cfn, "INPUT", vfdCFIO)
                        outData(outRow, 4) = DW_ResolveVFDIO(pd2Data, r, pd2Map, cfn, "OUTPUT", vfdCFIO)
                        outData(outRow, 5) = DW_ResolveVFDIO(pd2Data, r, pd2Map, cfn, "S_IN", vfdCFIO)
                        outData(outRow, 6) = DW_ResolveVFDIO(pd2Data, r, pd2Map, cfn, "S_OUT", vfdCFIO)

                        If pdAssign.Exists(UCase$(baseSeq)) Then
                            a = pdAssign(UCase$(baseSeq))
                            outData(outRow, 7) = a(2)
                            outData(outRow, 8) = a(1)
                            outData(outRow, 9) = a(0)
                            pd6Matched = pd6Matched + 1
                            If DW_Usable(a(2)) Then panelFilled = panelFilled + 1 Else panelMissing = panelMissing + 1
                        Else
                            outData(outRow, 7) = DW_Arr(pd2Data, r, pd2Map, "PNL")
                            If Not DW_Usable(outData(outRow, 7)) Then outData(outRow, 7) = DW_Arr(pd2Data, r, pd2Map, "PANEL")
                            outData(outRow, 8) = DW_Arr(pd2Data, r, pd2Map, "SWITCHGROUP")
                            outData(outRow, 9) = DW_Arr(pd2Data, r, pd2Map, "AREA_NUMBER")
                            If DW_Usable(outData(outRow, 7)) Then panelFilled = panelFilled + 1 Else panelMissing = panelMissing + 1
                            pd6Missing = pd6Missing + 1
                        End If

                        seen(k) = True
                        vfdAdded = vfdAdded + 1
                    End If
                End If
            End If
        Next r
    End If

    DW_ClearUsed wsField
    wsField.Cells(1, 1).Resize(1, 12).Value2 = hdr
    wsField.Columns(12).NumberFormat = "@"

    If outRow > 0 Then
        'The working array is intentionally sized for the maximum possible row count.
        'Trim it before the bulk write because qualifying VFD rows can be de-duplicated.
        ReDim finalData(1 To outRow, 1 To 12)
        For rr = 1 To outRow
            PDPerf_Pulse "mod20_DeviceWorkflow.DeviceWF_5B_Build_FieldIO_FAST [L482]" & " row " & CStr(rr)
            For cc = 1 To 12
                finalData(rr, cc) = outData(rr, cc)
            Next cc
        Next rr
        wsField.Cells(2, 1).Resize(outRow, 12).Value2 = finalData
    End If
    DW_FormatHeader wsField, 11

    Application.StatusBar = False
    Application.ScreenUpdating = oldScreen: Application.EnableEvents = oldEvents: Application.Calculation = oldCalc
    MsgBox "Step 5B complete: " & Format$(outRow, "#,##0") & " B03 device rows were built." & vbCrLf & vbCrLf & _
           "Qualifying PD2 VFD rows (TYPE=VFD + CONTROL_FUNCTION_NUMBER): " & Format$(qualifyingVFD, "#,##0") & vbCrLf & _
           "New VFD rows added from SIS_TAG_IMPORT_PD2: " & Format$(vfdAdded, "#,##0") & vbCrLf & _
           "Existing VFD rows refreshed instead of duplicated: " & Format$(vfdAlreadyPresent, "#,##0") & vbCrLf & _
           "VFD rows skipped because SIS_SEQUENCE was blank: " & Format$(vfdSkippedNoSeq, "#,##0") & vbCrLf & vbCrLf & _
           "Rows matched directly to final PD6: " & Format$(pd6Matched, "#,##0") & vbCrLf & _
           "Rows using source-sheet fallback: " & Format$(pd6Missing, "#,##0") & vbCrLf & _
           "Rows receiving a usable PNL: " & Format$(panelFilled, "#,##0") & vbCrLf & _
           "Rows still missing PNL: " & Format$(panelMissing, "#,##0") & vbCrLf & vbCrLf & _
           "PNL = final PD6 Panel" & vbCrLf & _
           "PWG = final PD6 SwitchGroup" & vbCrLf & _
           "LOCATION = final PD6 Area" & vbCrLf & vbCrLf & _
           "VFD IO priority = SIS_TAG_IMPORT_PD2, then FUNCTION DEFINITION by CONTROL_FUNCTION_NUMBER." & vbCrLf & _
           "Step 5C exports these A:K values directly to B03 IO_DEVICE_LIST.", _
           vbInformation, "Phase 5B - B03 Input + VFD"
    Exit Sub
Fail:
    Application.StatusBar = False
    Application.ScreenUpdating = oldScreen: Application.EnableEvents = oldEvents: Application.Calculation = oldCalc
    MsgBox "Step 5B failed. Error " & Err.number & ": " & Err.description, vbCritical, "Phase 5B"
End Sub

Private Sub DW_RefreshExistingVFDRow(ByRef outData As Variant, ByVal usedRows As Long, _
                                     ByVal deviceName As String, ByRef pd2Data As Variant, _
                                     ByVal pd2Row As Long, ByVal pd2Map As Object, _
                                     ByVal pdAssign As Object, ByVal baseSeq As String, _
                                     ByVal cfn As String, ByVal vfdCFIO As Object)
    Dim i As Long, a As Variant
    For i = 1 To usedRows
        If DW_DeviceKey(DW_Text(outData(i, 1)), DW_Text(outData(i, 2))) = DW_DeviceKey(deviceName, "VFD") Then

            outData(i, 3) = DW_ResolveVFDIO(pd2Data, pd2Row, pd2Map, cfn, "INPUT", vfdCFIO)
            outData(i, 4) = DW_ResolveVFDIO(pd2Data, pd2Row, pd2Map, cfn, "OUTPUT", vfdCFIO)
            outData(i, 5) = DW_ResolveVFDIO(pd2Data, pd2Row, pd2Map, cfn, "S_IN", vfdCFIO)
            outData(i, 6) = DW_ResolveVFDIO(pd2Data, pd2Row, pd2Map, cfn, "S_OUT", vfdCFIO)

            If pdAssign.Exists(UCase$(baseSeq)) Then
                a = pdAssign(UCase$(baseSeq))
                outData(i, 7) = a(2)
                outData(i, 8) = a(1)
                outData(i, 9) = a(0)
            Else
                outData(i, 7) = DW_Arr(pd2Data, pd2Row, pd2Map, "PNL")
                If Not DW_Usable(outData(i, 7)) Then outData(i, 7) = DW_Arr(pd2Data, pd2Row, pd2Map, "PANEL")
                outData(i, 8) = DW_Arr(pd2Data, pd2Row, pd2Map, "SWITCHGROUP")
                outData(i, 9) = DW_Arr(pd2Data, pd2Row, pd2Map, "AREA_NUMBER")
            End If
            Exit Sub
        End If
    Next i
End Sub

Private Function DW_BuildControlFunctionIOMap(ByVal ws As Worksheet) As Object
    Dim d As Object, data As Variant, m As Object
    Dim lr As Long, lc As Long, r As Long
    Dim cCFN As Long, cIn As Long, cOut As Long, cSIn As Long, cSOut As Long
    Dim key As String, a As Variant

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare

    lr = DW_LastRow(ws): lc = DW_LastCol(ws)
    If lr < 2 Or lc < 1 Then
        Set DW_BuildControlFunctionIOMap = d
        Exit Function
    End If

    data = ws.Range(ws.Cells(1, 1), ws.Cells(lr, lc)).Value2
    Set m = DW_HeaderMap(data, lc)

    cCFN = DW_ColAny(m, Array("CONTROL_FUNCTION_NUMBER", "CONTROL FUNCTION NUMBER", "CONTROLFUNCTIONNUMBER"))
    cIn = DW_ColAny(m, Array("INPUT"))
    cOut = DW_ColAny(m, Array("OUTPUT"))
    cSIn = DW_ColAny(m, Array("S_IN", "S IN", "SIN"))
    cSOut = DW_ColAny(m, Array("S_OUT", "S OUT", "SOUT"))

    If cCFN = 0 Then
        Err.Raise vbObjectError + 6210, , SH_FUNCDEF & " is missing CONTROL_FUNCTION_NUMBER."
    End If

    For r = 2 To lr
        PDPerf_Pulse "mod20_DeviceWorkflow.DW_BuildControlFunctionIOMap [L573]" & " row " & CStr(r)
        key = DW_Text(data(r, cCFN))
        If Len(key) > 0 Then
            If d.Exists(key) Then
                a = d(key)
            Else
                'Array positions: 0=INPUT, 1=OUTPUT, 2=S_IN, 3=S_OUT.
                'Empty means no value has been defined for that IO category.
                a = Array(Empty, Empty, Empty, Empty)
            End If

            If cIn > 0 Then DW_AddIOValue a, 0, data(r, cIn)
            If cOut > 0 Then DW_AddIOValue a, 1, data(r, cOut)
            If cSIn > 0 Then DW_AddIOValue a, 2, data(r, cSIn)
            If cSOut > 0 Then DW_AddIOValue a, 3, data(r, cSOut)

            d(key) = a
        End If
    Next r

    Set DW_BuildControlFunctionIOMap = d
End Function

Private Sub DW_AddIOValue(ByRef a As Variant, ByVal pos As Long, ByVal v As Variant)
    Dim n As Double

    If Not DW_Usable(v) Then Exit Sub

    If IsNumeric(v) Then
        n = CDbl(v)
        If IsEmpty(a(pos)) Or Len(DW_Text(a(pos))) = 0 Then
            a(pos) = n
        ElseIf IsNumeric(a(pos)) Then
            a(pos) = CDbl(a(pos)) + n
        End If
    Else
        'Non-numeric engineering values are preserved rather than discarded.
        If IsEmpty(a(pos)) Or Len(DW_Text(a(pos))) = 0 Then a(pos) = v
    End If
End Sub

Private Function DW_ResolveVFDIO(ByRef pd2Data As Variant, ByVal pd2Row As Long, _
                                 ByVal pd2Map As Object, ByVal cfn As String, _
                                 ByVal ioName As String, ByVal vfdCFIO As Object) As Variant
    Dim v As Variant, a As Variant, pos As Long

    'Priority 1: an explicit IO value already carried on SIS_TAG_IMPORT_PD2.
    v = DW_Arr(pd2Data, pd2Row, pd2Map, ioName)
    If DW_Usable(v) Then
        DW_ResolveVFDIO = v
        Exit Function
    End If

    'Priority 2: resolve the VFD CONTROL_FUNCTION_NUMBER against FUNCTION DEFINITION.
    If Len(DW_Text(cfn)) = 0 Then Exit Function
    If vfdCFIO Is Nothing Then Exit Function
    If Not vfdCFIO.Exists(DW_Text(cfn)) Then Exit Function

    Select Case UCase$(Trim$(ioName))
        Case "INPUT": pos = 0
        Case "OUTPUT": pos = 1
        Case "S_IN": pos = 2
        Case "S_OUT": pos = 3
        Case Else: Exit Function
    End Select

    a = vfdCFIO(DW_Text(cfn))
    If Not IsEmpty(a(pos)) And Len(DW_Text(a(pos))) > 0 Then
        DW_ResolveVFDIO = a(pos)
    End If
End Function

Private Function DW_DeviceKey(ByVal deviceName As String, ByVal itemType As String) As String
    Dim suffix As String
    deviceName = UCase$(Trim$(deviceName))
    itemType = UCase$(Trim$(itemType))

    'Treat 1060-01B and 1060-01B-VFD as the same VFD identity.
    'This prevents a Function Device VFD and its parent PD2 VFD from duplicating.
    If itemType = "VFD" Then
        suffix = "-VFD"
        If Len(deviceName) > Len(suffix) Then
            If Right$(deviceName, Len(suffix)) = suffix Then
                deviceName = Left$(deviceName, Len(deviceName) - Len(suffix))
            End If
        End If
    End If

    If Len(deviceName) > 0 Then DW_DeviceKey = deviceName & "|" & itemType
End Function

'====================================================================
' STEP 5C - external handoff only.
' Does NOT rebuild Specification or Field Device IO.
'====================================================================
Public Sub DeviceWF_5C_Send_Run_B03()
    If PDPerf_IsYielding Then Exit Sub
    Dim wsSrc As Worksheet, wbB03 As Workbook, wsDest As Worksheet
    Dim fd As Object, p As String, openedHere As Boolean
    Dim lr As Long, oldLR As Long
    Dim eText As String
    Dim oldCalc As XlCalculation, oldScreen As Boolean, oldEvents As Boolean, oldAlerts As Boolean

    Set wsSrc = DW_Sheet(ThisWorkbook, SH_FIELD)
    If wsSrc Is Nothing Then MsgBox "Missing '" & SH_FIELD & "'. Run Step 5B first.", vbCritical: Exit Sub
    lr = DW_LastRow(wsSrc)
    If lr < 2 Then MsgBox SH_FIELD & " contains no device rows. Run Step 5B first.", vbExclamation: Exit Sub

    Set fd = Application.FileDialog(3)
    With fd
        .title = "Step 5C - Select ONGEN B03 IO Generator workbook"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel Workbooks", "*.xls;*.xlsx;*.xlsm;*.xlsb"
        If .Show <> -1 Then Exit Sub
        p = CStr(.SelectedItems(1))
    End With
    If Len(p) = 0 Then Exit Sub

    oldScreen = Application.ScreenUpdating: oldEvents = Application.EnableEvents
    oldAlerts = Application.DisplayAlerts: oldCalc = Application.Calculation
    Application.ScreenUpdating = False: Application.EnableEvents = False
    Application.DisplayAlerts = False: Application.Calculation = xlCalculationManual
    On Error GoTo Fail

    Set wbB03 = DW_OpenWorkbook(p, False, openedHere)
    Set wsDest = DW_Sheet(wbB03, "IO_DEVICE_LIST")
    If wsDest Is Nothing Then Set wsDest = wbB03.Worksheets.Add(After:=wbB03.Worksheets(wbB03.Worksheets.count)): wsDest.name = "IO_DEVICE_LIST"

    oldLR = DW_LastRow(wsDest)
    If oldLR < lr Then oldLR = lr
    If oldLR < 1 Then oldLR = 1
    wsDest.Range("A1:K" & oldLR).ClearContents
    wsDest.Range("A1:K" & lr).Value2 = wsSrc.Range("A1:K" & lr).Value2
    wbB03.Save

    Application.StatusBar = "Step 5C: launching B03 generator once..."
    On Error Resume Next
    Application.Run "'" & Replace(wbB03.name, "'", "''") & "'!Run_IO_Generator_V6"
    If Err.number <> 0 Then
        eText = Err.description
        Err.Clear
        On Error GoTo Fail
        wbB03.Save
        If openedHere Then wbB03.Close SaveChanges:=True
        Application.StatusBar = False
        Application.ScreenUpdating = oldScreen: Application.EnableEvents = oldEvents
        Application.DisplayAlerts = oldAlerts: Application.Calculation = oldCalc
        MsgBox "The B03 input was transferred in one bulk write, but Run_IO_Generator_V6 could not be launched automatically." & vbCrLf & _
               "Open the selected workbook, run its generator, then use Step 5D." & vbCrLf & vbCrLf & eText, vbExclamation, "Phase 5C"
        Exit Sub
    End If
    On Error GoTo Fail

    wbB03.Save
    If openedHere Then wbB03.Close SaveChanges:=True
    Application.StatusBar = False
    Application.ScreenUpdating = oldScreen: Application.EnableEvents = oldEvents
    Application.DisplayAlerts = oldAlerts: Application.Calculation = oldCalc
    MsgBox "Step 5C complete. B03 input was transferred once and the generator was launched once." & vbCrLf & _
           "Use Step 5D to import IO_Block_Summary.", vbInformation, "Phase 5C"
    Exit Sub
Fail:
    On Error Resume Next
    If Not wbB03 Is Nothing Then If openedHere Then wbB03.Close SaveChanges:=False
    Application.StatusBar = False
    Application.ScreenUpdating = oldScreen: Application.EnableEvents = oldEvents
    Application.DisplayAlerts = oldAlerts: Application.Calculation = oldCalc
    MsgBox "Step 5C failed. Error " & Err.number & ": " & Err.description, vbCritical, "Phase 5C"
End Sub

'====================================================================
' STEP 5D - bulk import of B03 return. Existing merge module is called
' once after the bulk import; that merge module is optimized in V4.6.

Private Function DW_BuildCADHandleMap() As Object
    Dim d As Object, ws As Worksheet, data As Variant, hm As Object
    Dim lr As Long, lc As Long, r As Long
    Dim cSeq As Long, cHandle As Long, cType As Long
    Dim seq As String, typ As String, h As String, k As String
    Dim rec As Variant, oldRec As Variant

    Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare
    Set ws = DW_Sheet(ThisWorkbook, SH_CAD)
    If ws Is Nothing Then Set DW_BuildCADHandleMap = d: Exit Function

    lr = DW_LastRow(ws): lc = DW_LastCol(ws)
    If lr < 2 Then Set DW_BuildCADHandleMap = d: Exit Function
    data = ws.Range(ws.Cells(1, 1), ws.Cells(lr, lc)).Value2
    Set hm = DW_HeaderMap(data, lc)

    cSeq = DW_ColAny(hm, Array("SIS_SEQUENCE", "SEQUENCE"))
    cHandle = DW_ColAny(hm, Array("HANDLE"))
    cType = DW_ColAny(hm, Array("TYPE", "ITEM TYPE"))
    If cSeq = 0 Or cHandle = 0 Then Set DW_BuildCADHandleMap = d: Exit Function

    For r = 2 To lr
        PDPerf_Pulse "mod20_DeviceWorkflow.DW_BuildCADHandleMap [L768]" & " row " & CStr(r)
        seq = DW_Text(data(r, cSeq))
        If cType > 0 Then typ = DW_Text(data(r, cType)) Else typ = vbNullString
        If Len(seq) > 0 Then
            h = CStr(ws.Cells(r, cHandle).Text)
            If Len(h) = 0 Or InStr(1, h, "E+", vbTextCompare) > 0 Then h = DW_Text(data(r, cHandle))
            If Len(h) > 0 Then
                k = UCase$(seq)
                rec = Array(h, typ, seq)
                If Not d.Exists(k) Then
                    d.Add k, rec
                Else
                    oldRec = d(k)
                    'Prefer top-level CAD rows (blank TYPE) over nested device rows.
                    If Len(Trim$(CStr(oldRec(1)))) > 0 And Len(Trim$(typ)) = 0 Then d(k) = rec
                End If
            End If
        End If
    Next r

    Set DW_BuildCADHandleMap = d
End Function

Private Function DW_FindCADHandle(ByVal seq As String, ByVal d As Object) As String
    Dim candidate As String, k As String, p As Long, first4 As String
    Dim rec As Variant, key As Variant, best As Variant, haveBest As Boolean

    candidate = Trim$(seq)
    If Len(candidate) = 0 Or d Is Nothing Then Exit Function

    k = UCase$(candidate)
    If d.Exists(k) Then
        rec = d(k): DW_FindCADHandle = CStr(rec(0)): Exit Function
    End If

    'Progressively remove device suffixes.
    Do
        p = InStrRev(candidate, "-")
        If p <= 0 Then Exit Do
        candidate = Left$(candidate, p - 1)
        k = UCase$(candidate)
        If d.Exists(k) Then
            rec = d(k): DW_FindCADHandle = CStr(rec(0)): Exit Function
        End If
    Loop

    'Final four-digit family fallback.
    first4 = Left$(Trim$(seq), 4)
    If Len(first4) <> 4 Or Not IsNumeric(first4) Then Exit Function

    For Each key In d.Keys
        PDPerf_Pulse "mod20_DeviceWorkflow.DW_FindCADHandle [L818]"
        If Left$(CStr(key), 4) = UCase$(first4) Then
            rec = d(key)
            If Not haveBest Then
                best = rec: haveBest = True
            ElseIf Len(Trim$(CStr(best(1)))) > 0 And Len(Trim$(CStr(rec(1)))) = 0 Then
                best = rec
            ElseIf (Len(Trim$(CStr(best(1)))) = 0) = (Len(Trim$(CStr(rec(1)))) = 0) Then
                If Len(CStr(rec(2))) < Len(CStr(best(2))) Then best = rec
            End If
        End If
    Next key

    If haveBest Then DW_FindCADHandle = CStr(best(0))
End Function

'====================================================================
Public Sub DeviceWF_5D_Import_B03_Return_FAST()
    If PDPerf_IsYielding Then Exit Sub
    Dim fd As Object, p As String, wbB03 As Workbook, wsSrc As Worksheet, wsOut As Worksheet
    Dim data As Variant, m As Object, pdAssign As Object, outData() As Variant, hdr(1 To 1, 1 To 24) As Variant
    Dim lr As Long, lc As Long, r As Long, outRow As Long, rowCount As Long
    Dim cBlock As Long, cGroup As Long, cPanel As Long, cPart As Long
    Dim blockID As String, baseSeq As String, itemType As String, a As Variant
    Dim oldCalc As XlCalculation, oldScreen As Boolean, oldEvents As Boolean, oldAlerts As Boolean
    Dim openedHere As Boolean

    Set fd = Application.FileDialog(3)
    With fd
        .title = "Step 5D - Select completed ONGEN B03 IO Generator workbook"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel Workbooks", "*.xls;*.xlsx;*.xlsm;*.xlsb"
        If .Show <> -1 Then Exit Sub
        p = CStr(.SelectedItems(1))
    End With
    If Len(p) = 0 Then Exit Sub

    oldScreen = Application.ScreenUpdating: oldEvents = Application.EnableEvents
    oldAlerts = Application.DisplayAlerts: oldCalc = Application.Calculation
    Application.ScreenUpdating = False: Application.EnableEvents = False
    Application.DisplayAlerts = False: Application.Calculation = xlCalculationManual
    On Error GoTo Fail

    Set wbB03 = DW_OpenWorkbook(p, True, openedHere)
    Set wsSrc = DW_Sheet(wbB03, "IO_Block_Summary")
    If wsSrc Is Nothing Then Err.Raise vbObjectError + 6401, , "B03 sheet 'IO_Block_Summary' was not found."
    Set wsOut = DW_GetOrCreateSheet(ThisWorkbook, "IO DEVICE EM")

    lr = DW_LastRow(wsSrc): lc = DW_LastCol(wsSrc)
    If lr < 2 Then Err.Raise vbObjectError + 6402, , "IO_Block_Summary contains no rows."
    data = wsSrc.Range(wsSrc.Cells(1, 1), wsSrc.Cells(lr, lc)).Value2
    Set m = DW_HeaderMap(data, lc)
    cBlock = DW_ColAny(m, Array("BLOCK ID", "BLOCK_ID", "DEVICE NAME"))
    cGroup = DW_ColAny(m, Array("Power Group", "Power Group (SwitchGroup)", "SWITCHGROUP", "PWG"))
    cPanel = DW_ColAny(m, Array("Panel", "PNL"))
    cPart = DW_ColAny(m, Array("Module Part Number", "PART_NUM_1"))
    If cBlock = 0 Then Err.Raise vbObjectError + 6403, , "IO_Block_Summary is missing BLOCK ID."

    Set pdAssign = DW_BuildPD6AssignmentMap(DW_Sheet(ThisWorkbook, SH_PD6))
    For r = 2 To lr
        PDPerf_Pulse "mod20_DeviceWorkflow.DeviceWF_5D_Import_B03_Return_FAST [L877]" & " row " & CStr(r)
        If Len(DW_Text(data(r, cBlock))) > 0 Then rowCount = rowCount + 1
    Next r
    If rowCount = 0 Then Err.Raise vbObjectError + 6404, , "IO_Block_Summary contains no BLOCK ID rows."
    ReDim outData(1 To rowCount, 1 To 24)

    hdr(1, 1) = "SEQUENCE": hdr(1, 2) = "TYPE": hdr(1, 3) = "AREA_NUMBER": hdr(1, 4) = "SWITCHGROUP"
    hdr(1, 5) = "Panel": hdr(1, 6) = "NETWORK_IP": hdr(1, 7) = "LAYER": hdr(1, 8) = "BLOCK"
    hdr(1, 9) = "X": hdr(1, 10) = "Y": hdr(1, 11) = "Z": hdr(1, 12) = "ROTATION_DEG"
    hdr(1, 13) = "CONTROL_FUNCTION_NUMBER": hdr(1, 14) = "CONTROL_FUNCTION_TAG"
    hdr(1, 15) = "PART_NUM_1": hdr(1, 16) = "MANUF_NAME_1": hdr(1, 17) = "PART_NUM_2": hdr(1, 18) = "MANUF_NAME_2"
    hdr(1, 19) = "PART_NUM_3": hdr(1, 20) = "MANUF_NAME_3": hdr(1, 21) = "PART_NUM_4": hdr(1, 22) = "MANUF_NAME_4"
    hdr(1, 23) = "PART_NUM_5": hdr(1, 24) = "MANUF_NAME_5"

    For r = 2 To lr
        PDPerf_Pulse "mod20_DeviceWorkflow.DeviceWF_5D_Import_B03_Return_FAST [L891]" & " row " & CStr(r)
        blockID = DW_Text(data(r, cBlock))
        If Len(blockID) > 0 Then
            outRow = outRow + 1
            DW_SplitDevice blockID, baseSeq, itemType
            outData(outRow, 1) = baseSeq
            outData(outRow, 2) = itemType
            If pdAssign.Exists(UCase$(baseSeq)) Then
                a = pdAssign(UCase$(baseSeq))
                outData(outRow, 3) = a(0)
                outData(outRow, 4) = a(1)
                outData(outRow, 5) = a(2)
            Else
                If cGroup > 0 Then outData(outRow, 4) = data(r, cGroup)
                If cPanel > 0 Then outData(outRow, 5) = data(r, cPanel)
            End If
            If cPart > 0 Then outData(outRow, 15) = data(r, cPart)
        End If
    Next r

    DW_ClearUsed wsOut
    wsOut.Cells(1, 1).Resize(1, 24).Value2 = hdr
    If rowCount > 0 Then wsOut.Cells(2, 1).Resize(rowCount, 24).Value2 = outData
    DW_FormatHeader wsOut, 24

    If openedHere Then wbB03.Close SaveChanges:=False
    Set wbB03 = Nothing

    Application.StatusBar = "Step 5D: merging coordinates and device-part data in memory..."
    mod15_IOReturnMerge.IOReturn_Update_IO_DEVICE_EM_From_PD2_And_FunctionDefinition

    Application.StatusBar = False
    Application.ScreenUpdating = oldScreen: Application.EnableEvents = oldEvents
    Application.DisplayAlerts = oldAlerts: Application.Calculation = oldCalc
    MsgBox "Step 5D complete: " & Format$(rowCount, "#,##0") & " B03 return rows imported with one bulk write.", vbInformation, "Phase 5D"
    Exit Sub
Fail:
    On Error Resume Next
    If Not wbB03 Is Nothing Then If openedHere Then wbB03.Close SaveChanges:=False
    Application.StatusBar = False
    Application.ScreenUpdating = oldScreen: Application.EnableEvents = oldEvents
    Application.DisplayAlerts = oldAlerts: Application.Calculation = oldCalc
    MsgBox "Step 5D failed. Error " & Err.number & ": " & Err.description, vbCritical, "Phase 5D"
End Sub

Private Function DW_BuildPD6AssignmentMap(ByVal wsPD6 As Worksheet) As Object
    Dim d As Object, data As Variant, m As Object, lr As Long, lc As Long, r As Long
    Dim cSeq As Long, cArea As Long, cSw As Long, cPanel As Long
    Dim seq As String, cur As Variant, areaVal As Variant, swVal As Variant, panelVal As Variant
    Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare
    If wsPD6 Is Nothing Then Set DW_BuildPD6AssignmentMap = d: Exit Function
    lr = DW_LastRow(wsPD6): lc = DW_LastCol(wsPD6)
    If lr < 2 Or lc < 1 Then Set DW_BuildPD6AssignmentMap = d: Exit Function
    data = wsPD6.Range(wsPD6.Cells(1, 1), wsPD6.Cells(lr, lc)).Value2
    Set m = DW_HeaderMap(data, lc)
    cSeq = DW_ColAny(m, Array("SIS_SEQUENCE", "SEQUENCE"))
    cArea = DW_ColAny(m, Array("AREA_NUMBER", "AREA", "LOCATION"))
    cSw = DW_ColAny(m, Array("SWITCHGROUP", "SWG", "PWG"))
    cPanel = DW_ColAny(m, Array("PANEL", "PNL", "FEEDING PANEL"))
    If cSeq = 0 Then Set DW_BuildPD6AssignmentMap = d: Exit Function

    For r = 2 To lr
        PDPerf_Pulse "mod20_DeviceWorkflow.DW_BuildPD6AssignmentMap [L952]" & " row " & CStr(r)
        seq = DW_Text(data(r, cSeq))
        If Len(seq) > 0 Then
            If cArea > 0 Then areaVal = data(r, cArea) Else areaVal = vbNullString
            If cSw > 0 Then swVal = data(r, cSw) Else swVal = vbNullString
            If cPanel > 0 Then panelVal = data(r, cPanel) Else panelVal = vbNullString
            If d.Exists(UCase$(seq)) Then
                cur = d(UCase$(seq))
                If Not DW_Usable(cur(0)) And DW_Usable(areaVal) Then cur(0) = areaVal
                If Not DW_Usable(cur(1)) And DW_Usable(swVal) Then cur(1) = swVal
                If Not DW_Usable(cur(2)) And DW_Usable(panelVal) Then cur(2) = panelVal
                d(UCase$(seq)) = cur
            Else
                d(UCase$(seq)) = Array(areaVal, swVal, panelVal)
            End If
        End If
    Next r
    Set DW_BuildPD6AssignmentMap = d
End Function

Private Function DW_BaseSequence(ByVal deviceName As String, ByVal itemType As String) As String
    Dim d As String, t As String, suffix As String
    d = Trim$(deviceName): t = Trim$(itemType)
    If Len(d) = 0 Then Exit Function
    If Len(t) > 0 Then
        suffix = "-" & t
        If Len(d) > Len(suffix) Then
            If StrComp(Right$(d, Len(suffix)), suffix, vbTextCompare) = 0 Then
                DW_BaseSequence = Left$(d, Len(d) - Len(suffix))
                Exit Function
            End If
        End If
    End If
    DW_BaseSequence = d
End Function

Private Sub DW_SplitDevice(ByVal deviceName As String, ByRef baseSeq As String, ByRef itemType As String)
    Dim p As Long
    p = InStrRev(deviceName, "-")
    If p > 0 Then
        baseSeq = Left$(deviceName, p - 1)
        itemType = Mid$(deviceName, p + 1)
    Else
        baseSeq = deviceName
        itemType = vbNullString
    End If
End Sub

Private Function DW_HeaderMap(ByRef data As Variant, ByVal lastCol As Long) As Object
    Dim d As Object, c As Long, k As String
    Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare
    For c = 1 To lastCol
        k = DW_Key(data(1, c))
        If Len(k) > 0 Then If Not d.Exists(k) Then d(k) = c
    Next c
    Set DW_HeaderMap = d
End Function

Private Function DW_ColAny(ByVal m As Object, ByVal names As Variant) As Long
    Dim i As Long, k As String
    If m Is Nothing Then Exit Function
    For i = LBound(names) To UBound(names)
        k = DW_Key(names(i))
        If m.Exists(k) Then DW_ColAny = CLng(m(k)): Exit Function
    Next i
End Function

Private Function DW_Arr(ByRef data As Variant, ByVal r As Long, ByVal m As Object, ByVal headerName As String) As Variant
    Dim c As Long
    c = DW_ColAny(m, Array(headerName))
    If c > 0 Then DW_Arr = data(r, c) Else DW_Arr = vbNullString
End Function

Private Function DW_Key(ByVal v As Variant) As String
    Dim s As String
    s = UCase$(Trim$(CStr(v)))
    s = Replace(s, " ", ""): s = Replace(s, "_", ""): s = Replace(s, "-", "")
    s = Replace(s, "(", ""): s = Replace(s, ")", ""): s = Replace(s, "/", "")
    DW_Key = s
End Function

Private Function DW_Text(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then Exit Function
    DW_Text = Trim$(CStr(v))
End Function

Private Function DW_CountUsableAssignment(ByVal d As Object, ByVal itemIndex As Long) As Long
    Dim k As Variant, a As Variant
    If d Is Nothing Then Exit Function
    For Each k In d.Keys
        PDPerf_Pulse "mod20_DeviceWorkflow.DW_CountUsableAssignment [L1041]"
        a = d(k)
        If IsArray(a) Then
            If itemIndex >= LBound(a) And itemIndex <= UBound(a) Then
                If DW_Usable(a(itemIndex)) Then DW_CountUsableAssignment = DW_CountUsableAssignment + 1
            End If
        End If
    Next k
End Function

Private Function DW_Usable(ByVal v As Variant) As Boolean
    Dim s As String
    s = UCase$(DW_Text(v))
    DW_Usable = (Len(s) > 0 And s <> "NOT FOUND" And s <> "NOTFOUND" And s <> "N/A")
End Function

Private Function DW_Sheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        PDPerf_Pulse "mod20_DeviceWorkflow.DW_Sheet [L1059]"
        If StrComp(Trim$(ws.name), Trim$(sheetName), vbTextCompare) = 0 Then Set DW_Sheet = ws: Exit Function
    Next ws
End Function

Private Function DW_GetOrCreateSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Set DW_GetOrCreateSheet = DW_Sheet(wb, sheetName)
    If DW_GetOrCreateSheet Is Nothing Then
        Set DW_GetOrCreateSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
        DW_GetOrCreateSheet.name = sheetName
        If wb Is ThisWorkbook Then mod23_WorksheetNavigation.NAV_Ensure_Worksheet_Navigation DW_GetOrCreateSheet
    End If
End Function

Private Function DW_OpenWorkbook(ByVal FilePath As String, ByVal readOnlyMode As Boolean, ByRef openedHere As Boolean) As Workbook
    Dim wb As Workbook
    openedHere = False
    For Each wb In Application.Workbooks
        PDPerf_Pulse "mod20_DeviceWorkflow.DW_OpenWorkbook [L1076]"
        On Error Resume Next
        If StrComp(wb.FullName, FilePath, vbTextCompare) = 0 Then Set DW_OpenWorkbook = wb: On Error GoTo 0: Exit Function
        On Error GoTo 0
    Next wb
    Set DW_OpenWorkbook = Workbooks.Open(FilePath, ReadOnly:=readOnlyMode, UpdateLinks:=False, AddToMru:=False)
    openedHere = True
End Function

Private Function DW_LastRow(ByVal ws As Worksheet) As Long
    Dim f As Range
    On Error Resume Next
    Set f = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    On Error GoTo 0
    If f Is Nothing Then DW_LastRow = 0 Else DW_LastRow = f.Row
End Function

Private Function DW_LastCol(ByVal ws As Worksheet) As Long
    Dim f As Range
    On Error Resume Next
    Set f = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    On Error GoTo 0
    If f Is Nothing Then DW_LastCol = 0 Else DW_LastCol = f.Column
End Function

Private Sub DW_ClearUsed(ByVal ws As Worksheet)
    Dim lr As Long, lc As Long
    lr = DW_LastRow(ws): lc = DW_LastCol(ws)
    If lr > 0 And lc > 0 Then ws.Range(ws.Cells(1, 1), ws.Cells(lr, lc)).ClearContents
End Sub

Private Sub DW_FormatHeader(ByVal ws As Worksheet, ByVal lastCol As Long)
    With ws.Cells(1, 1).Resize(1, lastCol)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Sub
