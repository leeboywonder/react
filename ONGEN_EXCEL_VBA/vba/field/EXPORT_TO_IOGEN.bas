Attribute VB_Name = "EXPORT_TO_IOGEN"
Option Explicit

' Proposed replacement. Desktop Excel compile/run acceptance is still required.
' Preserves the public entry point used by the existing workbook.
Public Sub Export_IO_Device_List()
    Dim sourceSheet As Worksheet, targetSheet As Worksheet, targetBook As Workbook
    Dim chosen As Variant, openedHere As Boolean, changedHere As Boolean
    Dim oldScreen As Boolean, oldEvents As Boolean, oldAlerts As Boolean
    Dim oldCalc As XlCalculation, oldStatus As Variant, oldSecurity As Long
    Dim stateCaptured As Boolean, resultMessage As String, failureMessage As String
    Dim sourceCols(0 To 5) As Long, targetCols(0 To 5) As Long
    Dim sourceLast As Long, targetLast As Long, outputLast As Long
    Dim sourceData As Variant, savedColumns(0 To 5) As Variant
    Dim outColumn As Variant, sourceKeys As Object, targetKeys As Object
    Dim rowPlan() As Long, writePlan() As Boolean
    Dim r As Long, j As Long, targetRow As Long, sourceCount As Long
    Dim addCount As Long, updateCount As Long, skipCount As Long
    Dim key As String, device As String, itemType As String, countValue As Variant
    Dim policy As VbMsgBoxResult, matchCount As Long, keyCount As Long
    Dim srcHeader As Variant, dstHeader As Variant, wb As Workbook
    Dim requiredTarget As Variant, requiredHeader As Variant

    On Error GoTo Failed
    Set sourceSheet = ThisWorkbook.Worksheets("EM Designed function")
    srcHeader = Array(Array("Conv.ID", "DEVICE NAME", "CONV ID"), _
                      Array("Items Type", "ITEM TYPE", "ITEMS TYPE"), _
                      Array("INPUT"), Array("OUTPUT"), Array("S_IN"), Array("S_OUT"))
    dstHeader = Array(Array("DEVICE NAME", "Conv.ID", "CONV ID"), _
                      Array("ITEM TYPE", "Items Type", "ITEMS TYPE"), _
                      Array("INPUT"), Array("OUTPUT"), Array("S_IN"), Array("S_OUT"))
    For j = 0 To 5
        sourceCols(j) = OI_HeaderColumn(sourceSheet, srcHeader(j))
    Next j
    sourceLast = OI_LastRow(sourceSheet)
    If sourceLast < 2 Then
        MsgBox "No Specification device rows exist. Nothing was exported.", vbInformation, "ONGEN IO Export"
        Exit Sub
    End If
    sourceData = sourceSheet.Range(sourceSheet.Cells(1, 1), _
                 sourceSheet.Cells(sourceLast, sourceSheet.Cells(1, sourceSheet.Columns.Count).End(xlToLeft).Column)).Value2
    Set sourceKeys = CreateObject("Scripting.Dictionary")
    sourceKeys.CompareMode = vbTextCompare
    ReDim rowPlan(2 To sourceLast)
    ReDim writePlan(2 To sourceLast)

    ' Complete source validation before opening/changing the destination.
    For r = 2 To sourceLast
        If OI_HasMappedValue(sourceData, r, sourceCols) Then
            device = OI_Text(sourceData(r, sourceCols(0)))
            itemType = OI_Text(sourceData(r, sourceCols(1)))
            If Len(device) = 0 Or Len(itemType) = 0 Then
                Err.Raise vbObjectError + 5201, , "Source row " & r & ": DEVICE NAME/Conv.ID and ITEM TYPE are both required. Review the row; no identity was invented."
            End If
            key = OI_Key(device, itemType)
            If sourceKeys.Exists(key) Then
                Err.Raise vbObjectError + 5202, , "Source rows " & sourceKeys(key) & " and " & r & " share DEVICE NAME + ITEM TYPE. Resolve distinct device/variant identities before export."
            End If
            sourceKeys.Add key, r
            For j = 2 To 5
                countValue = OI_ValidatedCount(sourceData(r, sourceCols(j)), r, CStr(srcHeader(j)(0)))
                sourceData(r, sourceCols(j)) = countValue
            Next j
            sourceCount = sourceCount + 1
        End If
    Next r
    If sourceCount = 0 Then
        MsgBox "No Specification device rows exist. Nothing was exported.", vbInformation, "ONGEN IO Export"
        Exit Sub
    End If

    chosen = Application.GetOpenFilename("Excel Workbooks (*.xlsx;*.xlsm),*.xlsx;*.xlsm", , "Select ONGEN IO Generator workbook")
    If VarType(chosen) = vbBoolean Then Exit Sub
    If StrComp(CStr(chosen), ThisWorkbook.FullName, vbTextCompare) = 0 Then
        Err.Raise vbObjectError + 5203, , "Choose the IO Generator workbook, not the Specification workbook."
    End If

    oldScreen = Application.ScreenUpdating: oldEvents = Application.EnableEvents
    oldAlerts = Application.DisplayAlerts: oldCalc = Application.Calculation
    oldStatus = Application.StatusBar: oldSecurity = Application.AutomationSecurity
    stateCaptured = True
    Application.ScreenUpdating = False: Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.AutomationSecurity = 3 ' msoAutomationSecurityForceDisable for data-only opening
    Application.StatusBar = "ONGEN: validating IO export..."

    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, CStr(chosen), vbTextCompare) = 0 Then
            Set targetBook = wb
            Exit For
        End If
    Next wb
    If targetBook Is Nothing Then
        Set targetBook = Workbooks.Open(Filename:=CStr(chosen), UpdateLinks:=0, ReadOnly:=False)
        openedHere = True
    ElseIf Not targetBook.Saved Then
        Err.Raise vbObjectError + 5204, , "The selected IO workbook has unsaved edits. Save those edits before exporting, so rollback cannot discard them."
    End If
    If targetBook.ReadOnly Then Err.Raise vbObjectError + 5205, , "The selected IO workbook is read-only."
    Set targetSheet = targetBook.Worksheets("IO_DEVICE_LIST")
    For j = 0 To 5
        targetCols(j) = OI_HeaderColumn(targetSheet, dstHeader(j))
    Next j
    ' B05 contract: validate all 11 columns, preserve its five assignment columns.
    requiredTarget = Array("PNL", "PWG", "LOCATION", "Standard Block", "Safety Block")
    For Each requiredHeader In requiredTarget
        keyCount = OI_HeaderColumn(targetSheet, Array(CStr(requiredHeader)))
    Next requiredHeader
    OI_AssertLiteralInputColumns targetSheet, targetCols
    targetLast = OI_LastRow(targetSheet)
    If targetLast < 1 Then targetLast = 1
    outputLast = targetLast
    Set targetKeys = CreateObject("Scripting.Dictionary")
    targetKeys.CompareMode = vbTextCompare
    For r = 2 To targetLast
        device = OI_Text(targetSheet.Cells(r, targetCols(0)).Value2)
        itemType = OI_Text(targetSheet.Cells(r, targetCols(1)).Value2)
        If Len(device) > 0 Or Len(itemType) > 0 Then
            If Len(device) = 0 Or Len(itemType) = 0 Then
                Err.Raise vbObjectError + 5206, , "Destination row " & r & ": incomplete device identity. Review before export."
            End If
            key = OI_Key(device, itemType)
            If targetKeys.Exists(key) Then
                Err.Raise vbObjectError + 5207, , "Destination rows " & targetKeys(key) & " and " & r & " share DEVICE NAME + ITEM TYPE. Resolve duplicate identities before export."
            End If
            targetKeys.Add key, r
        End If
    Next r
    For r = 2 To sourceLast
        If OI_HasMappedValue(sourceData, r, sourceCols) Then
            key = OI_Key(OI_Text(sourceData(r, sourceCols(0))), OI_Text(sourceData(r, sourceCols(1))))
            If targetKeys.Exists(key) Then matchCount = matchCount + 1
        End If
    Next r
    policy = vbNo
    If matchCount > 0 Then
        policy = MsgBox(matchCount & " existing device/type matches." & vbCrLf & _
                 "Yes: update the six mapped fields and append new devices." & vbCrLf & _
                 "No: append new devices and skip existing matches." & vbCrLf & _
                 "PNL, PWG, LOCATION and block assignments are preserved." & vbCrLf & _
                 "Cancel: make no changes.", vbYesNoCancel + vbQuestion, "ONGEN IO Export")
        If policy = vbCancel Then GoTo CleanExit
    End If
    For r = 2 To sourceLast
        If OI_HasMappedValue(sourceData, r, sourceCols) Then
            key = OI_Key(OI_Text(sourceData(r, sourceCols(0))), OI_Text(sourceData(r, sourceCols(1))))
            If targetKeys.Exists(key) Then
                targetRow = CLng(targetKeys(key))
                If policy = vbYes Then
                    writePlan(r) = True: updateCount = updateCount + 1
                Else
                    skipCount = skipCount + 1
                End If
            Else
                outputLast = outputLast + 1: targetRow = outputLast
                writePlan(r) = True: addCount = addCount + 1
            End If
            rowPlan(r) = targetRow
        End If
    Next r
    If addCount + updateCount = 0 Then
        resultMessage = "No changes. Existing device/type matches skipped: " & skipCount
        GoTo CleanExit
    End If
    If outputLast > targetSheet.Rows.Count Then Err.Raise vbObjectError + 5208, , "Export would exceed Excel row capacity."

    ' Snapshot literal input values in mapped columns; no other columns are rewritten.
    For j = 0 To 5
        savedColumns(j) = targetSheet.Range(targetSheet.Cells(2, targetCols(j)), targetSheet.Cells(outputLast, targetCols(j))).Value2
    Next j
    changedHere = True
    For j = 0 To 5
        ReDim outColumn(1 To outputLast - 1, 1 To 1)
        If outputLast = 2 Then
            outColumn(1, 1) = savedColumns(j)
        Else
            For r = 1 To outputLast - 1
                outColumn(r, 1) = savedColumns(j)(r, 1)
            Next r
        End If
        For r = 2 To sourceLast
            If writePlan(r) Then outColumn(rowPlan(r) - 1, 1) = sourceData(r, sourceCols(j))
        Next r
        targetSheet.Cells(2, targetCols(j)).Resize(outputLast - 1, 1).Value2 = outColumn
    Next j
    targetBook.Save
    changedHere = False
    resultMessage = "IO export saved." & vbCrLf & "Added: " & addCount & "; updated: " & updateCount & "; skipped: " & skipCount & vbCrLf & _
                    "PNL, PWG, LOCATION, Standard Block and Safety Block were preserved. New rows require assignment review."
    GoTo CleanExit
Failed:
    failureMessage = "IO export failed: " & Err.Description
    If changedHere Then
        On Error Resume Next
        Err.Clear
        For j = 0 To 5
            targetSheet.Range(targetSheet.Cells(2, targetCols(j)), targetSheet.Cells(outputLast, targetCols(j))).Value2 = savedColumns(j)
        Next j
        If Err.Number <> 0 Then failureMessage = failureMessage & vbCrLf & "Rollback could not be confirmed. Close without saving and reopen the previously saved workbook."
        On Error GoTo 0
    End If
CleanExit:
    On Error Resume Next
    If openedHere And Not targetBook Is Nothing Then targetBook.Close SaveChanges:=False
    If stateCaptured Then
        Application.ScreenUpdating = oldScreen: Application.EnableEvents = oldEvents
        Application.DisplayAlerts = oldAlerts: Application.Calculation = oldCalc
        Application.StatusBar = oldStatus: Application.AutomationSecurity = oldSecurity
    End If
    On Error GoTo 0
    If Len(failureMessage) > 0 Then
        MsgBox failureMessage, vbCritical, "ONGEN IO Export"
    ElseIf Len(resultMessage) > 0 Then
        MsgBox resultMessage, vbInformation, "ONGEN IO Export"
    End If
End Sub

Private Function OI_HeaderColumn(ByVal ws As Worksheet, ByVal aliases As Variant) As Long
    Dim c As Long, candidate As Variant, found As Long, lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        For Each candidate In aliases
            If OI_HeaderKey(OI_Text(ws.Cells(1, c).Value2)) = OI_HeaderKey(CStr(candidate)) Then
                If found <> 0 And found <> c Then Err.Raise vbObjectError + 5210, , "Ambiguous headers on " & ws.Name & ": " & CStr(aliases(0))
                found = c
                Exit For
            End If
        Next candidate
    Next c
    If found = 0 Then Err.Raise vbObjectError + 5211, , "Missing header on " & ws.Name & ": " & CStr(aliases(0))
    OI_HeaderColumn = found
End Function

Private Function OI_HeaderKey(ByVal text As String) As String
    text = UCase$(Trim$(Replace(text, ChrW(160), " ")))
    text = Replace(Replace(Replace(text, " ", ""), "_", ""), ".", "")
    OI_HeaderKey = text
End Function

Private Function OI_Text(ByVal value As Variant) As String
    If IsError(value) Then Err.Raise vbObjectError + 5212, , "An identity/header cell contains an Excel error. Review source and destination data."
    OI_Text = Trim$(Replace(CStr(value), ChrW(160), " "))
End Function

Private Function OI_Key(ByVal device As String, ByVal itemType As String) As String
    OI_Key = CStr(Len(device)) & ":" & UCase$(device) & CStr(Len(itemType)) & ":" & UCase$(itemType)
End Function

Private Function OI_ValidatedCount(ByVal value As Variant, ByVal sourceRow As Long, ByVal header As String) As Variant
    Dim n As Double
    If IsError(value) Then GoTo InvalidCount
    If Len(Trim$(CStr(value))) = 0 Then
        OI_ValidatedCount = Empty
        Exit Function
    End If
    If VarType(value) = vbBoolean Or Not IsNumeric(value) Then GoTo InvalidCount
    n = CDbl(value)
    If n < 0 Or n > 2147483647# Then GoTo InvalidCount
    If n <> Fix(n) Then GoTo InvalidCount
    OI_ValidatedCount = n
    Exit Function
InvalidCount:
    Err.Raise vbObjectError + 5213, , "Source row " & sourceRow & " has invalid " & header & ". Use a nonnegative integer or blank."
End Function

Private Function OI_HasMappedValue(ByRef data As Variant, ByVal r As Long, ByRef cols() As Long) As Boolean
    Dim j As Long
    For j = 0 To 5
        If IsError(data(r, cols(j))) Then
            OI_HasMappedValue = True
            Exit Function
        End If
        If Len(Trim$(CStr(data(r, cols(j))))) > 0 Then
            OI_HasMappedValue = True
            Exit Function
        End If
    Next j
End Function

Private Function OI_LastRow(ByVal ws As Worksheet) As Long
    Dim found As Range
    Set found = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, _
                             SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False)
    If found Is Nothing Then OI_LastRow = 1 Else OI_LastRow = found.Row
End Function

Private Sub OI_AssertLiteralInputColumns(ByVal ws As Worksheet, ByRef cols() As Long)
    Dim j As Long, lastRow As Long, formulas As Range
    lastRow = OI_LastRow(ws)
    If lastRow < 2 Then Exit Sub
    For j = 0 To 5
        Set formulas = Nothing
        On Error Resume Next
        Set formulas = ws.Range(ws.Cells(2, cols(j)), ws.Cells(lastRow, cols(j))).SpecialCells(xlCellTypeFormulas)
        On Error GoTo 0
        If Not formulas Is Nothing Then
            Err.Raise vbObjectError + 5214, , "The IO input column contains formulas at " & formulas.Cells(1, 1).Address(False, False) & ". Review/convert only the intended managed input values before export; no formulas were overwritten."
        End If
    Next j
End Sub
