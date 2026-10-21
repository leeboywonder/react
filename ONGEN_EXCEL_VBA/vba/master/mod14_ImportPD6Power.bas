Attribute VB_Name = "mod14_ImportPD6Power"
Option Explicit

'==========================================================
' Macro Name : Fill_PD6_From_PowerCalculation_RELEASED
' Author     : Lee Pierre
' Version    : Released
'
' Purpose:
'   Opens the Power Calculation workbook selected by the user.
'   Reads worksheet "Export PD6 Designed Definition".
'   Matches source rows to the current workbook worksheet
'   "PD6 System Designed" by SIS_SEQUENCE.
'
' Fields Filled In Destination "PD6 System Designed":
'   - SIS_ELEMENT_NUMBER  : filled from Power Calculation export
'   - SWITCHGROUP         : filled from Power Calculation export
'   - Panel               : filled from Power Calculation export
'   - AREA_NUMBER         : only filled when destination AREA_NUMBER is blank
'
' Device-sheet assignment is handled in Phase 5A/5B so Step 4D remains a single PD6 import pass.
'
' Match Rule:
'   - Source SIS_SEQUENCE must match destination SIS_SEQUENCE.
'   - Matching is case-insensitive and trimmed.
'
' Safety Rules:
'   - The source Power Calculation workbook is opened Read-Only.
'   - The source workbook is closed without saving after the update.
'   - Destination headers are created if missing, except SIS_SEQUENCE.
'   - The macro stops if destination SIS_SEQUENCE is missing.
'==========================================================

Private Const DEST_SHEET_PD6 As String = "PD6 System Designed"
Private Const DEST_SHEET_FIELD_IO As String = "Field Device IO ONGEN"
Private Const SRC_SHEET_POWER As String = "Export PD6 Designed Definition"

Private Const OVERWRITE_ELEMENT_SWITCH_PANEL As Boolean = True
Private Const ONLY_FILL_MISSING_AREA As Boolean = True
Private Const OVERWRITE_FIELD_DEVICE_IO As Boolean = True

Public Sub Fill_PD6_From_PowerCalculation_RELEASED()
    If PDPerf_IsYielding Then Exit Sub
    Dim FilePath As String

    FilePath = mod11_PowerRoundTrip.PowerRT_SelectPowerCalcPath("Phase 4D - Import Final Power Calculation Return")
    If Len(FilePath) = 0 Then Exit Sub

    'This button entry owns the user-facing error message. The path entry below
    'propagates errors to its caller so Phase 4D cannot continue with backfill.
    On Error GoTo DirectFail
    Fill_PD6_From_PowerCalculation_Path_RELEASED FilePath
    Exit Sub
DirectFail:
    MsgBox "PD6 update from Power Calculation failed." & vbCrLf & _
           "Error " & Err.number & ": " & Err.description, _
           vbCritical, "Power Calculation PD6 Update"
End Sub

Public Sub Fill_PD6_From_PowerCalculation_Path_RELEASED(ByVal FilePath As String)

    Dim wbDest As Workbook
    Dim wbPower As Workbook
    Dim wsDest As Worksheet
    Dim wsPower As Worksheet
    Dim openedPowerHere As Boolean

    Dim oldCalc As XlCalculation
    Dim oldScreen As Boolean
    Dim oldEvents As Boolean
    Dim oldAlerts As Boolean

    Dim updatedElement As Long
    Dim updatedSwitch As Long
    Dim updatedPanel As Long
    Dim updatedArea As Long
    Dim matchedRows As Long
    Dim missingMatches As Long
    Dim appendedPSURows As Long
    Dim errNum As Long, errDesc As String, errSource As String

    Set wbDest = ThisWorkbook

    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    oldCalc = Application.Calculation
    oldAlerts = Application.DisplayAlerts

    On Error GoTo CleanFail

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    Set wsDest = SheetByNameLoose(wbDest, DEST_SHEET_PD6)

    If wsDest Is Nothing Then
        Err.Raise vbObjectError + 1000, , "Destination worksheet not found: " & DEST_SHEET_PD6
    End If

    If Len(Trim$(FilePath)) = 0 Then GoTo CleanExit

    Set wbPower = mod11_PowerRoundTrip.PowerRT_OpenWorkbook(FilePath, True, openedPowerHere)
    If Not mod11_PowerRoundTrip.PowerRT_ValidatePowerCalcWorkbook(wbPower, "4D") Then
        Err.Raise vbObjectError + 1090, , "The selected workbook is not a compatible Power Calculation return workbook for Phase 4D."
    End If
    Set wsPower = SheetByNameLoose(wbPower, SRC_SHEET_POWER)

    If wsPower Is Nothing Then
        Err.Raise vbObjectError + 1001, , "Source worksheet not found in Power Calculation workbook: " & SRC_SHEET_POWER
    End If

    UpdatePD6FromPowerSheet wsDest, wsPower, matchedRows, missingMatches, _
                            updatedElement, updatedSwitch, updatedPanel, updatedArea, appendedPSURows

CleanExit:
    If Not wbPower Is Nothing Then
        If openedPowerHere Then wbPower.Close SaveChanges:=False
    End If

    Application.ScreenUpdating = oldScreen
    Application.EnableEvents = oldEvents
    Application.Calculation = oldCalc
    Application.DisplayAlerts = oldAlerts

    If Len(CStr(FilePath)) > 0 Then
        MsgBox BuildCompletionMessage(matchedRows, missingMatches, updatedElement, _
                                      updatedSwitch, updatedPanel, updatedArea, appendedPSURows), _
               vbInformation, "Power Calculation PD6 Update"
    End If

    Exit Sub

CleanFail:
    errNum = Err.number
    errDesc = Err.description
    errSource = Err.Source
    On Error Resume Next
    If Not wbPower Is Nothing Then
        If openedPowerHere Then wbPower.Close SaveChanges:=False
    End If

    Application.ScreenUpdating = oldScreen
    Application.EnableEvents = oldEvents
    Application.Calculation = oldCalc
    Application.DisplayAlerts = oldAlerts
    On Error GoTo 0

    'Do not swallow an import failure: the Phase 4D caller must not backfill.
    If Len(errSource) = 0 Then errSource = "mod14_ImportPD6Power.Fill_PD6_From_PowerCalculation_Path_RELEASED"
    Err.Raise errNum, errSource, errDesc

End Sub


Private Function BuildCompletionMessage( _
    ByVal matchedRows As Long, _
    ByVal missingMatches As Long, _
    ByVal updatedElement As Long, _
    ByVal updatedSwitch As Long, _
    ByVal updatedPanel As Long, _
    ByVal updatedArea As Long, _
    ByVal appendedPSURows As Long) As String

    Dim msg As String

    msg = "PD6 update from Power Calculation complete." & vbCrLf & vbCrLf
    msg = msg & "PD6 System Designed:" & vbCrLf
    msg = msg & "Matched SIS_SEQUENCE rows: " & matchedRows & vbCrLf
    msg = msg & "Rows with no source match: " & missingMatches & vbCrLf
    msg = msg & "SIS_ELEMENT_NUMBER updated: " & updatedElement & vbCrLf
    msg = msg & "SWITCHGROUP updated: " & updatedSwitch & vbCrLf
    msg = msg & "Panel updated: " & updatedPanel & vbCrLf
    msg = msg & "Missing AREA_NUMBER filled: " & updatedArea & vbCrLf
    msg = msg & "New PSU rows appended from Power Calc: " & appendedPSURows & vbCrLf & vbCrLf
    msg = msg & "Device-sheet updates are intentionally deferred to Phase 5A/5B." & vbCrLf
    msg = msg & "This keeps Step 4D as one Power Calc -> PD6 import pass."

    BuildCompletionMessage = msg
End Function

Private Sub UpdatePD6FromPowerSheet( _
    ByVal wsDest As Worksheet, _
    ByVal wsPower As Worksheet, _
    ByRef matchedRows As Long, _
    ByRef missingMatches As Long, _
    ByRef updatedElement As Long, _
    ByRef updatedSwitch As Long, _
    ByRef updatedPanel As Long, _
    ByRef updatedArea As Long, _
    ByRef appendedPSURows As Long)

    Dim destLastRow As Long
    Dim destLastCol As Long
    Dim destHeaderRow As Long
    Dim srcLastRow As Long
    Dim srcLastCol As Long
    Dim srcHeaderRow As Long

    Dim destExact As Object
    Dim destNorm As Object
    Dim srcExact As Object
    Dim srcNorm As Object
    Dim powerDict As Object

    Dim destData As Variant
    Dim srcData As Variant
    Dim outData As Variant
    Dim rowValues As Variant

    Dim destSeqCol As Long
    Dim destElementCol As Long
    Dim destSwitchCol As Long
    Dim destPanelCol As Long
    Dim destAreaCol As Long

    Dim srcSeqCol As Long
    Dim srcElementCol As Long
    Dim srcSwitchCol As Long
    Dim srcPanelCol As Long
    Dim srcAreaCol As Long

    Dim r As Long
    Dim DataRowCount As Long
    Dim seq As String
    Dim seqKey As String

    destLastRow = LastUsedRow(wsDest)
    destLastCol = LastUsedCol(wsDest)
    srcLastRow = LastUsedRow(wsPower)
    srcLastCol = LastUsedCol(wsPower)

    If destLastRow < 2 Or destLastCol < 1 Then
        Err.Raise vbObjectError + 1100, , "Destination PD6 sheet has no data rows to update."
    End If

    If srcLastRow < 2 Or srcLastCol < 1 Then
        Err.Raise vbObjectError + 1101, , "Power Calculation export sheet has no data rows to read. Step 4C did not finish successfully; run Step 4C before Step 4D."
    End If

    destHeaderRow = LocateHeaderRow(wsDest, destLastRow, destLastCol)
    srcHeaderRow = LocateHeaderRow(wsPower, srcLastRow, srcLastCol)

    If destHeaderRow = 0 Then
        Err.Raise vbObjectError + 1102, , "Could not locate SIS_SEQUENCE header in destination PD6 sheet."
    End If

    If srcHeaderRow = 0 Then
        Err.Raise vbObjectError + 1103, , "Could not locate SIS_SEQUENCE header in Power Calculation export sheet."
    End If

    Set destExact = CreateObject("Scripting.Dictionary")
    destExact.CompareMode = vbTextCompare
    Set destNorm = CreateObject("Scripting.Dictionary")
    destNorm.CompareMode = vbTextCompare
    Set srcExact = CreateObject("Scripting.Dictionary")
    srcExact.CompareMode = vbTextCompare
    Set srcNorm = CreateObject("Scripting.Dictionary")
    srcNorm.CompareMode = vbTextCompare

    BuildHeaderMapsFromSheet wsDest, destHeaderRow, destLastCol, destExact, destNorm
    BuildHeaderMapsFromSheet wsPower, srcHeaderRow, srcLastCol, srcExact, srcNorm

    destSeqCol = FindHeaderColumn(destExact, destNorm, "SIS_SEQUENCE")
    If destSeqCol = 0 Then
        Err.Raise vbObjectError + 1104, , "Destination header missing: SIS_SEQUENCE"
    End If

    destElementCol = GetOrCreateHeaderColumn(wsDest, destExact, destNorm, destHeaderRow, destLastCol, "SIS_ELEMENT_NUMBER")
    destSwitchCol = GetOrCreateHeaderColumn(wsDest, destExact, destNorm, destHeaderRow, destLastCol, "SWITCHGROUP")
    destPanelCol = GetOrCreateHeaderColumn(wsDest, destExact, destNorm, destHeaderRow, destLastCol, "Panel")
    destAreaCol = GetOrCreateHeaderColumn(wsDest, destExact, destNorm, destHeaderRow, destLastCol, "AREA_NUMBER")

    destLastCol = LastUsedCol(wsDest)

    srcSeqCol = FindHeaderColumn(srcExact, srcNorm, "SIS_SEQUENCE")
    srcElementCol = FindHeaderColumn(srcExact, srcNorm, "SIS_ELEMENT_NUMBER")
    srcSwitchCol = FindHeaderColumn(srcExact, srcNorm, "SWITCHGROUP")
    srcPanelCol = FindHeaderColumn(srcExact, srcNorm, "Panel")
    srcAreaCol = FindHeaderColumn(srcExact, srcNorm, "AREA_NUMBER")

    If srcSeqCol = 0 Then Err.Raise vbObjectError + 1105, , "Source header missing: SIS_SEQUENCE"
    If srcElementCol = 0 Then Err.Raise vbObjectError + 1106, , "Source header missing: SIS_ELEMENT_NUMBER"
    If srcSwitchCol = 0 Then Err.Raise vbObjectError + 1107, , "Source header missing: SWITCHGROUP"
    If srcPanelCol = 0 Then Err.Raise vbObjectError + 1108, , "Source header missing: Panel or PNL"
    If srcAreaCol = 0 Then Err.Raise vbObjectError + 1109, , "Source header missing: AREA_NUMBER"

    srcData = wsPower.Range(wsPower.Cells(1, 1), wsPower.Cells(srcLastRow, srcLastCol)).Value2

    Set powerDict = CreateObject("Scripting.Dictionary")
    powerDict.CompareMode = vbTextCompare

    For r = srcHeaderRow + 1 To srcLastRow
        PDPerf_Pulse "mod14_ImportPD6Power.UpdatePD6FromPowerSheet [L284]" & " row " & CStr(r)
        seq = TrimSafe(srcData(r, srcSeqCol))

        If Len(seq) > 0 Then
            seqKey = UCase$(seq)

            ReDim rowValues(1 To 4)
            rowValues(1) = srcData(r, srcElementCol)
            rowValues(2) = srcData(r, srcSwitchCol)
            rowValues(3) = srcData(r, srcPanelCol)
            rowValues(4) = srcData(r, srcAreaCol)

            If Not powerDict.Exists(seqKey) Then
                powerDict.Add seqKey, rowValues
            End If
        End If
    Next r

    destData = wsDest.Range(wsDest.Cells(1, 1), wsDest.Cells(destLastRow, destLastCol)).Value2
    DataRowCount = destLastRow - destHeaderRow

    If DataRowCount < 1 Then
        Err.Raise vbObjectError + 1110, , "Destination PD6 sheet has no data rows below the header."
    End If

    ReDim outData(1 To DataRowCount, 1 To destLastCol)

    For r = 1 To DataRowCount
        PDPerf_Pulse "mod14_ImportPD6Power.UpdatePD6FromPowerSheet [L311]" & " row " & CStr(r)
        CopyDestRow destData, destHeaderRow + r, outData, r, destLastCol

        seq = TrimSafe(destData(destHeaderRow + r, destSeqCol))
        seqKey = UCase$(seq)

        If Len(seqKey) > 0 And powerDict.Exists(seqKey) Then
            rowValues = powerDict(seqKey)
            matchedRows = matchedRows + 1

            If Not IsBlankValue(rowValues(1)) Then
                If OVERWRITE_ELEMENT_SWITCH_PANEL Or IsBlankValue(outData(r, destElementCol)) Then
                    outData(r, destElementCol) = rowValues(1)
                    updatedElement = updatedElement + 1
                End If
            End If

            If Not IsBlankValue(rowValues(2)) Then
                If OVERWRITE_ELEMENT_SWITCH_PANEL Or IsBlankValue(outData(r, destSwitchCol)) Then
                    outData(r, destSwitchCol) = rowValues(2)
                    updatedSwitch = updatedSwitch + 1
                End If
            End If

            If Not IsBlankValue(rowValues(3)) Then
                If OVERWRITE_ELEMENT_SWITCH_PANEL Or IsBlankValue(outData(r, destPanelCol)) Then
                    outData(r, destPanelCol) = rowValues(3)
                    updatedPanel = updatedPanel + 1
                End If
            End If

            If Not IsBlankValue(rowValues(4)) Then
                If ONLY_FILL_MISSING_AREA Then
                    If IsBlankValue(outData(r, destAreaCol)) Then
                        outData(r, destAreaCol) = rowValues(4)
                        updatedArea = updatedArea + 1
                    End If
                Else
                    outData(r, destAreaCol) = rowValues(4)
                    updatedArea = updatedArea + 1
                End If
            End If

        ElseIf Len(seqKey) > 0 Then
            missingMatches = missingMatches + 1
        End If

    Next r

    wsDest.Range(wsDest.Cells(destHeaderRow + 1, 1), wsDest.Cells(destLastRow, destLastCol)).Value = outData

    AppendMissingPSURowsFromPower wsDest, wsPower, destHeaderRow, srcHeaderRow, appendedPSURows
    destLastRow = LastUsedRow(wsDest)
    destLastCol = LastUsedCol(wsDest)

    FormatUpdatedPD6 wsDest, destHeaderRow, destLastRow, destLastCol

End Sub

Private Sub AppendMissingPSURowsFromPower(ByVal wsDest As Worksheet, _
                                               ByVal wsPower As Worksheet, _
                                               ByVal destHeaderRow As Long, _
                                               ByVal srcHeaderRow As Long, _
                                               ByRef appendedPSURows As Long)
    Dim destLastRow As Long, destLastCol As Long, srcLastRow As Long, srcLastCol As Long
    Dim destExact As Object, destNorm As Object, srcExact As Object, srcNorm As Object
    Dim destSeqCol As Long, srcSeqCol As Long, srcTypeCol As Long
    Dim destSet As Object, srcData As Variant, appendData() As Variant
    Dim r As Long, c As Long, appendCount As Long, outRow As Long
    Dim seq As String, typ As String, rawHeader As String, srcCol As Long

    destLastRow = LastUsedRow(wsDest): destLastCol = LastUsedCol(wsDest)
    srcLastRow = LastUsedRow(wsPower): srcLastCol = LastUsedCol(wsPower)
    If srcLastRow <= srcHeaderRow Then Exit Sub

    Set destExact = CreateObject("Scripting.Dictionary"): destExact.CompareMode = vbTextCompare
    Set destNorm = CreateObject("Scripting.Dictionary"): destNorm.CompareMode = vbTextCompare
    Set srcExact = CreateObject("Scripting.Dictionary"): srcExact.CompareMode = vbTextCompare
    Set srcNorm = CreateObject("Scripting.Dictionary"): srcNorm.CompareMode = vbTextCompare
    BuildHeaderMapsFromSheet wsDest, destHeaderRow, destLastCol, destExact, destNorm
    BuildHeaderMapsFromSheet wsPower, srcHeaderRow, srcLastCol, srcExact, srcNorm

    destSeqCol = FindHeaderColumn(destExact, destNorm, "SIS_SEQUENCE")
    srcSeqCol = FindHeaderColumn(srcExact, srcNorm, "SIS_SEQUENCE")
    srcTypeCol = FindHeaderColumn(srcExact, srcNorm, "TYPE")
    If destSeqCol = 0 Or srcSeqCol = 0 Or srcTypeCol = 0 Then Exit Sub

    Set destSet = CreateObject("Scripting.Dictionary"): destSet.CompareMode = vbTextCompare
    For r = destHeaderRow + 1 To destLastRow
        PDPerf_Pulse "mod14_ImportPD6Power.AppendMissingPSURowsFromPower [L399]" & " row " & CStr(r)
        seq = TrimSafe(wsDest.Cells(r, destSeqCol).Value2)
        If Len(seq) > 0 Then destSet(UCase$(seq)) = True
    Next r

    srcData = wsPower.Range(wsPower.Cells(1, 1), wsPower.Cells(srcLastRow, srcLastCol)).Value2
    For r = srcHeaderRow + 1 To srcLastRow
        PDPerf_Pulse "mod14_ImportPD6Power.AppendMissingPSURowsFromPower [L405]" & " row " & CStr(r)
        seq = TrimSafe(srcData(r, srcSeqCol))
        typ = UCase$(TrimSafe(srcData(r, srcTypeCol)))
        If typ = "PSU" And Len(seq) > 0 Then
            If Not destSet.Exists(UCase$(seq)) Then appendCount = appendCount + 1
        End If
    Next r
    If appendCount = 0 Then Exit Sub

    ReDim appendData(1 To appendCount, 1 To destLastCol)
    outRow = 0
    For r = srcHeaderRow + 1 To srcLastRow
        PDPerf_Pulse "mod14_ImportPD6Power.AppendMissingPSURowsFromPower [L416]" & " row " & CStr(r)
        seq = TrimSafe(srcData(r, srcSeqCol))
        typ = UCase$(TrimSafe(srcData(r, srcTypeCol)))
        If typ = "PSU" And Len(seq) > 0 Then
            If Not destSet.Exists(UCase$(seq)) Then
                outRow = outRow + 1
                For c = 1 To destLastCol
                    rawHeader = TrimSafe(wsDest.Cells(destHeaderRow, c).Value2)
                    If Len(rawHeader) > 0 Then
                        srcCol = FindHeaderColumn(srcExact, srcNorm, rawHeader)
                        If srcCol > 0 Then appendData(outRow, c) = srcData(r, srcCol)
                    End If
                Next c
                destSet(UCase$(seq)) = True
            End If
        End If
    Next r

    wsDest.Cells(destLastRow + 1, 1).Resize(appendCount, destLastCol).Value2 = appendData
    appendedPSURows = appendedPSURows + appendCount
End Sub

Private Sub UpdateFieldDeviceIOFromPD6Sheet( _
    ByVal wsField As Worksheet, _
    ByVal wsPD6 As Worksheet, _
    ByRef fieldMatchedRows As Long, _
    ByRef fieldMissingMatches As Long, _
    ByRef updatedFieldPNL As Long, _
    ByRef updatedFieldPWG As Long, _
    ByRef updatedFieldLocation As Long)

    Dim fieldLastRow As Long
    Dim fieldLastCol As Long
    Dim fieldHeaderRow As Long
    Dim pd6LastRow As Long
    Dim pd6LastCol As Long
    Dim pd6HeaderRow As Long

    Dim fieldExact As Object
    Dim fieldNorm As Object
    Dim pd6Exact As Object
    Dim pd6Norm As Object
    Dim pd6Dict As Object

    Dim fieldData As Variant
    Dim pd6Data As Variant
    Dim outData As Variant
    Dim rowValues As Variant

    Dim deviceNameCol As Long
    Dim pnlCol As Long
    Dim pwgCol As Long
    Dim locationCol As Long

    Dim pd6SeqCol As Long
    Dim pd6PanelCol As Long
    Dim pd6SwitchCol As Long
    Dim pd6AreaCol As Long

    Dim r As Long
    Dim DataRowCount As Long
    Dim seq As String
    Dim seqKey As String

    fieldLastRow = LastUsedRow(wsField)
    fieldLastCol = LastUsedCol(wsField)
    pd6LastRow = LastUsedRow(wsPD6)
    pd6LastCol = LastUsedCol(wsPD6)

    If fieldLastRow < 2 Or fieldLastCol < 1 Then Exit Sub
    If pd6LastRow < 2 Or pd6LastCol < 1 Then Exit Sub

    fieldHeaderRow = LocateHeaderRowByHeader(wsField, fieldLastRow, fieldLastCol, "Device Name")
    pd6HeaderRow = LocateHeaderRow(wsPD6, pd6LastRow, pd6LastCol)

    If fieldHeaderRow = 0 Then
        Err.Raise vbObjectError + 1200, , _
                  "Could not locate Device Name header in Field Device IO ONGEN."
    End If

    If pd6HeaderRow = 0 Then
        Err.Raise vbObjectError + 1201, , _
                  "Could not locate SIS_SEQUENCE header in PD6 System Designed sheet."
    End If

    Set fieldExact = CreateObject("Scripting.Dictionary")
    fieldExact.CompareMode = vbTextCompare
    Set fieldNorm = CreateObject("Scripting.Dictionary")
    fieldNorm.CompareMode = vbTextCompare
    Set pd6Exact = CreateObject("Scripting.Dictionary")
    pd6Exact.CompareMode = vbTextCompare
    Set pd6Norm = CreateObject("Scripting.Dictionary")
    pd6Norm.CompareMode = vbTextCompare

    BuildHeaderMapsFromSheet wsField, fieldHeaderRow, fieldLastCol, fieldExact, fieldNorm
    BuildHeaderMapsFromSheet wsPD6, pd6HeaderRow, pd6LastCol, pd6Exact, pd6Norm

    deviceNameCol = FindHeaderColumn(fieldExact, fieldNorm, "Device Name")
    If deviceNameCol = 0 Then
        Err.Raise vbObjectError + 1202, , "Field Device IO ONGEN header missing: Device Name"
    End If

    pnlCol = GetOrCreateHeaderColumn(wsField, fieldExact, fieldNorm, fieldHeaderRow, fieldLastCol, "PNL")
    pwgCol = GetOrCreateHeaderColumn(wsField, fieldExact, fieldNorm, fieldHeaderRow, fieldLastCol, "PWG")
    locationCol = GetOrCreateHeaderColumn(wsField, fieldExact, fieldNorm, fieldHeaderRow, fieldLastCol, "LOCATION")

    fieldLastCol = LastUsedCol(wsField)

    pd6SeqCol = FindHeaderColumn(pd6Exact, pd6Norm, "SIS_SEQUENCE")
    pd6PanelCol = FindHeaderColumn(pd6Exact, pd6Norm, "Panel")
    pd6SwitchCol = FindHeaderColumn(pd6Exact, pd6Norm, "SWITCHGROUP")
    pd6AreaCol = FindHeaderColumn(pd6Exact, pd6Norm, "AREA_NUMBER")

    If pd6SeqCol = 0 Then Err.Raise vbObjectError + 1203, , "PD6 header missing: SIS_SEQUENCE"
    If pd6PanelCol = 0 Then Err.Raise vbObjectError + 1204, , "PD6 header missing: Panel or PNL"
    If pd6SwitchCol = 0 Then Err.Raise vbObjectError + 1205, , "PD6 header missing: SWITCHGROUP or PWG"
    If pd6AreaCol = 0 Then Err.Raise vbObjectError + 1206, , "PD6 header missing: AREA_NUMBER or LOCATION"

    pd6Data = wsPD6.Range(wsPD6.Cells(1, 1), wsPD6.Cells(pd6LastRow, pd6LastCol)).Value2

    Set pd6Dict = CreateObject("Scripting.Dictionary")
    pd6Dict.CompareMode = vbTextCompare

    For r = pd6HeaderRow + 1 To pd6LastRow
        PDPerf_Pulse "mod14_ImportPD6Power.UpdateFieldDeviceIOFromPD6Sheet [L539]" & " row " & CStr(r)
        seq = TrimSafe(pd6Data(r, pd6SeqCol))

        If Len(seq) > 0 Then
            seqKey = UCase$(seq)

            ReDim rowValues(1 To 3)
            rowValues(1) = pd6Data(r, pd6PanelCol)
            rowValues(2) = pd6Data(r, pd6SwitchCol)
            rowValues(3) = pd6Data(r, pd6AreaCol)

            If Not pd6Dict.Exists(seqKey) Then
                pd6Dict.Add seqKey, rowValues
            End If
        End If
    Next r

    fieldData = wsField.Range(wsField.Cells(1, 1), _
                              wsField.Cells(fieldLastRow, fieldLastCol)).Value2
    DataRowCount = fieldLastRow - fieldHeaderRow

    If DataRowCount < 1 Then Exit Sub

    ReDim outData(1 To DataRowCount, 1 To fieldLastCol)

    For r = 1 To DataRowCount
        PDPerf_Pulse "mod14_ImportPD6Power.UpdateFieldDeviceIOFromPD6Sheet [L564]" & " row " & CStr(r)
        CopyDestRow fieldData, fieldHeaderRow + r, outData, r, fieldLastCol

        seq = TrimSafe(fieldData(fieldHeaderRow + r, deviceNameCol))
        seqKey = UCase$(seq)

        If Len(seqKey) > 0 And pd6Dict.Exists(seqKey) Then
            rowValues = pd6Dict(seqKey)
            fieldMatchedRows = fieldMatchedRows + 1

            If IsUsableImportValue(rowValues(1)) Then
                If OVERWRITE_FIELD_DEVICE_IO Or IsBlankValue(outData(r, pnlCol)) Then
                    outData(r, pnlCol) = rowValues(1)
                    updatedFieldPNL = updatedFieldPNL + 1
                End If
            End If

            If IsUsableImportValue(rowValues(2)) Then
                If OVERWRITE_FIELD_DEVICE_IO Or IsBlankValue(outData(r, pwgCol)) Then
                    outData(r, pwgCol) = rowValues(2)
                    updatedFieldPWG = updatedFieldPWG + 1
                End If
            End If

            If IsUsableImportValue(rowValues(3)) Then
                If OVERWRITE_FIELD_DEVICE_IO Or IsBlankValue(outData(r, locationCol)) Then
                    outData(r, locationCol) = rowValues(3)
                    updatedFieldLocation = updatedFieldLocation + 1
                End If
            End If

        ElseIf Len(seqKey) > 0 Then
            fieldMissingMatches = fieldMissingMatches + 1
        End If

    Next r

    wsField.Range(wsField.Cells(fieldHeaderRow + 1, 1), _
                  wsField.Cells(fieldLastRow, fieldLastCol)).Value = outData

    FormatUpdatedPD6 wsField, fieldHeaderRow, fieldLastRow, fieldLastCol

End Sub

Private Sub CopyDestRow( _
    ByRef destData As Variant, _
    ByVal sourceRow As Long, _
    ByRef outData As Variant, _
    ByVal outRow As Long, _
    ByVal lastCol As Long)

    Dim c As Long

    For c = 1 To lastCol
        outData(outRow, c) = destData(sourceRow, c)
    Next c

End Sub

Private Function GetOrCreateHeaderColumn( _
    ByVal ws As Worksheet, _
    ByVal exactMap As Object, _
    ByVal normMap As Object, _
    ByVal headerRow As Long, _
    ByRef lastCol As Long, _
    ByVal headerName As String) As Long

    Dim colNum As Long
    Dim exactKey As String
    Dim NormKey As String

    colNum = FindHeaderColumn(exactMap, normMap, headerName)

    If colNum > 0 Then
        GetOrCreateHeaderColumn = colNum
        Exit Function
    End If

    lastCol = lastCol + 1
    ws.Cells(headerRow, lastCol).Value = headerName

    exactKey = UCase$(Trim$(headerName))
    NormKey = NormalizeHeader(headerName)

    If Not exactMap.Exists(exactKey) Then exactMap.Add exactKey, lastCol
    If Not normMap.Exists(NormKey) Then normMap.Add NormKey, lastCol

    GetOrCreateHeaderColumn = lastCol

End Function

Private Sub FormatUpdatedPD6( _
    ByVal ws As Worksheet, _
    ByVal headerRow As Long, _
    ByVal lastRow As Long, _
    ByVal lastCol As Long)

    With ws.rows(headerRow)
        .Font.Bold = True
        .Interior.Color = RGB(255, 242, 204)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    'Keep the existing column widths and body formatting. Full-range borders/AutoFit
    'were intentionally removed because they dominate runtime on large PD6 datasets.

End Sub

Private Sub BuildHeaderMapsFromSheet( _
    ByVal ws As Worksheet, _
    ByVal headerRow As Long, _
    ByVal lastCol As Long, _
    ByVal exactMap As Object, _
    ByVal normMap As Object)

    Dim c As Long
    Dim rawHeader As String
    Dim exactKey As String
    Dim NormKey As String

    For c = 1 To lastCol

        rawHeader = TrimSafe(ws.Cells(headerRow, c).Value)

        If Len(rawHeader) > 0 Then

            exactKey = UCase$(rawHeader)
            NormKey = NormalizeHeader(rawHeader)

            If Len(exactKey) > 0 Then
                If Not exactMap.Exists(exactKey) Then exactMap.Add exactKey, c
            End If

            If Len(NormKey) > 0 Then
                If Not normMap.Exists(NormKey) Then normMap.Add NormKey, c
            End If

        End If

    Next c

End Sub

Private Function FindHeaderColumn( _
    ByVal exactMap As Object, _
    ByVal normMap As Object, _
    ByVal wantedHeader As String) As Long

    Dim aliases As Variant
    Dim i As Long
    Dim exactKey As String
    Dim NormKey As String

    aliases = HeaderAliases(wantedHeader)

    For i = LBound(aliases) To UBound(aliases)

        exactKey = UCase$(Trim$(CStr(aliases(i))))

        If exactMap.Exists(exactKey) Then
            FindHeaderColumn = CLng(exactMap(exactKey))
            Exit Function
        End If

    Next i

    For i = LBound(aliases) To UBound(aliases)

        NormKey = NormalizeHeader(CStr(aliases(i)))

        If normMap.Exists(NormKey) Then
            FindHeaderColumn = CLng(normMap(NormKey))
            Exit Function
        End If

    Next i

    FindHeaderColumn = 0

End Function

Private Function HeaderAliases(ByVal targetHeader As String) As Variant

    Dim key As String
    Dim s As String

    key = NormalizeHeader(targetHeader)
    s = targetHeader

    Select Case key

        Case "SISSEQUENCE"
            s = "SIS_SEQUENCE|SIS SEQUENCE|SIS_SEQUENCE_NUMBER|SEQUENCE|SIS #|SIS|SIS_Sequence"

        Case "DEVICENAME"
            s = "Device Name|DEVICE NAME|Device name|SIS_SEQUENCE|SIS SEQUENCE|SEQUENCE"

        Case "SISELEMENTNUMBER"
            s = "SIS_ELEMENT_NUMBER|SIS ELEMENT NUMBER|ELEMENT_NUMBER|ELEMENT NUMBER|SIS Element Number"

        Case "AREANUMBER"
            s = "AREA_NUMBER|AREA NUMBER|AREA|Area_Number"

        Case "SWITCHGROUP"
            s = "SWITCHGROUP|SWITCH GROUP|SWITCH_GROUP|PWG|SwitchGroup"

        Case "PANEL"
            s = "Panel|PNL|PANEL"

        Case "PNL"
            s = "PNL|Panel|PANEL"

        Case "PWG"
            s = "PWG|SWITCHGROUP|SWITCH GROUP|SWITCH_GROUP|SwitchGroup"

        Case "LOCATION"
            s = "LOCATION|Location|AREA_NUMBER|AREA NUMBER|Area_Number|AREA"

    End Select

    HeaderAliases = Split(s, "|")

End Function

Private Function LocateHeaderRow( _
    ByVal ws As Worksheet, _
    ByVal lastRow As Long, _
    ByVal lastCol As Long) As Long

    Dim r As Long
    Dim c As Long
    Dim maxScanRow As Long
    Dim cellText As String

    maxScanRow = Application.Min(lastRow, 25)

    For r = 1 To maxScanRow
        PDPerf_Pulse "mod14_ImportPD6Power.LocateHeaderRow [L801]" & " row " & CStr(r)

        For c = 1 To lastCol

            cellText = NormalizeHeader(ws.Cells(r, c).Value)

            If cellText = "SISSEQUENCE" _
               Or cellText = "SEQUENCE" _
               Or cellText = "SIS" Then

                LocateHeaderRow = r
                Exit Function

            End If

        Next c

    Next r

    LocateHeaderRow = 0

End Function


Private Function LocateHeaderRowByHeader( _
    ByVal ws As Worksheet, _
    ByVal lastRow As Long, _
    ByVal lastCol As Long, _
    ByVal wantedHeader As String) As Long

    Dim r As Long
    Dim c As Long
    Dim maxScanRow As Long
    Dim exactMap As Object
    Dim normMap As Object

    maxScanRow = Application.Min(lastRow, 25)

    For r = 1 To maxScanRow
        PDPerf_Pulse "mod14_ImportPD6Power.LocateHeaderRowByHeader [L839]" & " row " & CStr(r)
        Set exactMap = CreateObject("Scripting.Dictionary")
        exactMap.CompareMode = vbTextCompare
        Set normMap = CreateObject("Scripting.Dictionary")
        normMap.CompareMode = vbTextCompare

        For c = 1 To lastCol
            AddHeaderToMaps ws.Cells(r, c).Value, c, exactMap, normMap
        Next c

        If FindHeaderColumn(exactMap, normMap, wantedHeader) > 0 Then
            LocateHeaderRowByHeader = r
            Exit Function
        End If
    Next r

    LocateHeaderRowByHeader = 0

End Function

Private Sub AddHeaderToMaps( _
    ByVal headerValue As Variant, _
    ByVal colNum As Long, _
    ByVal exactMap As Object, _
    ByVal normMap As Object)

    Dim rawHeader As String
    Dim exactKey As String
    Dim NormKey As String

    rawHeader = TrimSafe(headerValue)

    If Len(rawHeader) = 0 Then Exit Sub

    exactKey = UCase$(rawHeader)
    NormKey = NormalizeHeader(rawHeader)

    If Len(exactKey) > 0 Then
        If Not exactMap.Exists(exactKey) Then exactMap.Add exactKey, colNum
    End If

    If Len(NormKey) > 0 Then
        If Not normMap.Exists(NormKey) Then normMap.Add NormKey, colNum
    End If

End Sub

Private Function SheetByNameLoose( _
    ByVal wb As Workbook, _
    ByVal sheetName As String) As Worksheet

    Dim ws As Worksheet

    For Each ws In wb.Worksheets
        PDPerf_Pulse "mod14_ImportPD6Power.SheetByNameLoose [L892]"

        If UCase$(Trim$(ws.name)) = UCase$(Trim$(sheetName)) Then
            Set SheetByNameLoose = ws
            Exit Function
        End If

    Next ws

    Set SheetByNameLoose = Nothing

End Function

Private Function LastUsedRow(ByVal ws As Worksheet) As Long

    Dim f As Range

    On Error Resume Next
    Set f = ws.Cells.Find(What:="*", _
                          LookIn:=xlFormulas, _
                          SearchOrder:=xlByRows, _
                          SearchDirection:=xlPrevious)
    On Error GoTo 0

    If f Is Nothing Then
        LastUsedRow = 0
    Else
        LastUsedRow = f.Row
    End If

End Function

Private Function LastUsedCol(ByVal ws As Worksheet) As Long

    Dim f As Range

    On Error Resume Next
    Set f = ws.Cells.Find(What:="*", _
                          LookIn:=xlFormulas, _
                          SearchOrder:=xlByColumns, _
                          SearchDirection:=xlPrevious)
    On Error GoTo 0

    If f Is Nothing Then
        LastUsedCol = 0
    Else
        LastUsedCol = f.Column
    End If

End Function

Private Function NormalizeHeader(ByVal valueIn As Variant) As String

    Dim s As String

    If IsError(valueIn) Then
        NormalizeHeader = vbNullString
        Exit Function
    End If

    s = UCase$(Trim$(CStr(valueIn)))
    s = Replace(s, Chr$(160), " ")
    s = Replace(s, "_", "")
    s = Replace(s, " ", "")
    s = Replace(s, "-", "")
    s = Replace(s, "(", "")
    s = Replace(s, ")", "")
    s = Replace(s, "/", "")
    s = Replace(s, "\", "")
    s = Replace(s, ".", "")
    s = Replace(s, "#", "")
    s = Replace(s, ":", "")

    NormalizeHeader = s

End Function

Private Function TrimSafe(ByVal valueIn As Variant) As String

    If IsError(valueIn) Then
        TrimSafe = vbNullString
    ElseIf IsNull(valueIn) Then
        TrimSafe = vbNullString
    ElseIf IsEmpty(valueIn) Then
        TrimSafe = vbNullString
    Else
        TrimSafe = Trim$(CStr(valueIn))
    End If

End Function


Private Function IsUsableImportValue(ByVal valueIn As Variant) As Boolean

    Dim s As String

    If IsBlankValue(valueIn) Then
        IsUsableImportValue = False
        Exit Function
    End If

    s = UCase$(TrimSafe(valueIn))

    If s = "NOT FOUND" Or s = "N/A" Or s = "NA" Then
        IsUsableImportValue = False
    Else
        IsUsableImportValue = True
    End If

End Function

Private Function IsBlankValue(ByVal valueIn As Variant) As Boolean

    If IsError(valueIn) Then
        IsBlankValue = True
    ElseIf IsNull(valueIn) Then
        IsBlankValue = True
    ElseIf IsEmpty(valueIn) Then
        IsBlankValue = True
    ElseIf VarType(valueIn) = vbString Then
        IsBlankValue = Len(Trim$(CStr(valueIn))) = 0
    Else
        IsBlankValue = False
    End If

End Function


