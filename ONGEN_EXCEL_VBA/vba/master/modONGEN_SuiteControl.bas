Attribute VB_Name = "modONGEN_SuiteControl"
Option Explicit

Private Const CONTROL_SHEET As String = "ONGEN CONTROL CENTER"
Private Const LOG_SHEET As String = "ONGEN RUN LOG"
Private mBusy As Boolean

Public Sub ONGEN_InstallControlCenter()
    Dim ws As Worksheet
    Dim i As Long
    Dim labels As Variant, macros As Variant, notes As Variant

    Application.ScreenUpdating = False
    Set ws = GetOrCreateSheet(CONTROL_SHEET)
    ws.Cells.Clear
    Do While ws.Shapes.Count > 0
        ws.Shapes(1).Delete
    Loop

    With ws
        .Tab.Color = RGB(0, 120, 104)
        .Range("A1:H2").Merge
        .Range("A1").Value = "ONGEN  |  ENGINEERING CONTROL CENTER"
        .Range("A1").Font.Size = 22
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Color = vbWhite
        .Range("A1").Interior.Color = RGB(18, 45, 64)
        .Range("A4").Value = "WORKFLOW"
        .Range("B4").Value = "ACTION"
        .Range("E4").Value = "ENGINEERING NOTE"
        .Range("A4:H4").Font.Bold = True
        .Range("A4:H4").Interior.Color = RGB(221, 235, 247)
        .Columns("A").ColumnWidth = 18
        .Columns("B:D").ColumnWidth = 16
        .Columns("E:H").ColumnWidth = 18
        .Rows("1:2").RowHeight = 28
    End With

    labels = Array( _
        "System self-check", "Phase 1 - Validate", "Phase 2 - Build PD5", _
        "Phase 3 - Generate PD2", "Phase 4A - Send PD5", "Phase 4B - Build PD6", _
        "Phase 4C - Send PD6", "Phase 4D - Import Power", _
        "Phase 5A - Update Devices", "Phase 5B - Build Field I/O", _
        "Phase 5C - Run B03", "Phase 5D - Import B03", _
        "Export PD6 CSV", "Export I/O Block CSV", "Open legacy dashboard")
    macros = Array( _
        "ONGEN_SelfCheck", "ONGEN_RunValidate", "ONGEN_RunBuildPD5", _
        "ONGEN_RunGeneratePD2", "ONGEN_RunPower4A", "ONGEN_RunPower4B", _
        "ONGEN_RunPower4C", "ONGEN_RunPower4D", _
        "ONGEN_RunDevice5A", "ONGEN_RunDevice5B", "ONGEN_RunDevice5C", _
        "ONGEN_RunDevice5D", "ONGEN_RunExportPD6", "ONGEN_RunExportIO", _
        "ONGEN_OpenLegacyDashboard")
    notes = Array( _
        "Checks required sheets and companion files.", _
        "Runs the existing definition workflow validation.", _
        "Creates PD5 through the existing reviewed workflow.", _
        "Uses the bounded-clear optimization in mod09.", _
        "Power Calculation V4.6 was not supplied; select the workbook when prompted.", _
        "Builds PD6 locally.", _
        "Power Calculation V4.6 was not supplied; select the workbook when prompted.", _
        "Uses the Phase 4D error-propagation patch.", _
        "Uses the Phase 5A row-counter/data-preservation patch.", _
        "Builds the Field I/O transfer set.", _
        "Runs the existing B03 handoff.", _
        "Imports the B03 return into the master.", _
        "Exports the released PD6 system CSV.", _
        "Exports I/O block engineering data.", _
        "Opens the workbook's existing dashboard.")

    For i = LBound(labels) To UBound(labels)
        AddActionRow ws, 6 + i * 2, CStr(labels(i)), CStr(macros(i)), CStr(notes(i))
    Next i

    ws.Range("A38:H40").Merge
    ws.Range("A38").Value = "Controlled limitation: ONGEN Power Calculation V4.6 was larger than the available upload limit and is not included. Phase 4 must not be treated as validated until that workbook is added and tested in Microsoft Excel."
    ws.Range("A38").WrapText = True
    ws.Range("A38").Interior.Color = RGB(255, 235, 156)
    ws.Range("A38").Font.Bold = True
    ws.Range("A42").Value = "Installed"
    ws.Range("B42").Value = Now
    ws.Range("A43").Value = "Workbook"
    ws.Range("B43").Value = ThisWorkbook.FullName
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    Application.ScreenUpdating = True
End Sub

Public Sub ONGEN_OpenControlCenter()
    On Error GoTo InstallIt
    ThisWorkbook.Worksheets(CONTROL_SHEET).Activate
    Exit Sub
InstallIt:
    Err.Clear
    ONGEN_InstallControlCenter
End Sub

Public Sub ONGEN_SelfCheck()
    Dim required As Variant, i As Long, report As String, missing As Long
    Dim companion As Variant, p As String
    required = Array("MAIN", "FUNCTION DEFINITION", "SIS_TAG_EXPORT", "BOQ", _
                     "UI_Config", "SIS_TAG_PD5")
    report = "ONGEN workbook check" & vbCrLf & String(34, "-") & vbCrLf
    For i = LBound(required) To UBound(required)
        If SheetExists(CStr(required(i))) Then
            report = report & "OK  Sheet: " & CStr(required(i)) & vbCrLf
        Else
            report = report & "CHECK  Sheet not found: " & CStr(required(i)) & vbCrLf
            missing = missing + 1
        End If
    Next i

    companion = Array("ONGEN Field Devices Specification -v2.1.7.1.xlsm", _
                      "ONGEN IO GENERATOR - B05.xlsm", _
                      "ONGEN EAS - V2.5.3.xlsm", _
                      "ONGEN_EAS_MASTER_FUNCTION_LIBRARY.xlsm", _
                      "ONGEN_FUNCTION_CODE_LOOKUP.csv")
    p = ThisWorkbook.Path & Application.PathSeparator
    For i = LBound(companion) To UBound(companion)
        If FileExistsSafe(p & CStr(companion(i))) Then
            report = report & "OK  File: " & CStr(companion(i)) & vbCrLf
        Else
            report = report & "CHECK  File not found: " & CStr(companion(i)) & vbCrLf
            missing = missing + 1
        End If
    Next i
    report = report & "CHECK  Power Calculation V4.6 is not included or validated." & vbCrLf
    AppendLog "SELF CHECK", IIf(missing = 0, "RETURNED WITH POWER LIMITATION", "CHECK REQUIRED"), report
    MsgBox report, IIf(missing = 0, vbInformation, vbExclamation), "ONGEN self-check"
End Sub

Public Sub ONGEN_RunValidate(): RunStep "Validate_Definition_Workflow", "Phase 1 validation", "": End Sub
Public Sub ONGEN_RunBuildPD5(): RunStep "Run_Create_PD5", "Phase 2 build PD5", "": End Sub
Public Sub ONGEN_RunGeneratePD2(): RunStep "Run_Generate_PD2", "Phase 3 generate PD2", "": End Sub
Public Sub ONGEN_RunPower4A(): RunStep "PowerRT_4A_Send_PD5_FAST", "Phase 4A send PD5", "SIS_TAG_PD5": End Sub
Public Sub ONGEN_RunPower4B(): RunStep "PowerRT_4B_Build_PD6_FAST", "Phase 4B build PD6", "SIS_TAG_PD5": End Sub
Public Sub ONGEN_RunPower4C(): RunStep "PowerRT_4C_Send_PD6_FAST", "Phase 4C send PD6", "": End Sub
Public Sub ONGEN_RunPower4D(): RunStep "PowerRT_4D_Import_Power_FAST", "Phase 4D import Power", "": End Sub
Public Sub ONGEN_RunDevice5A(): RunStep "DeviceWF_5A_Update_Specification_FAST", "Phase 5A update devices", "FUNCTION DEVICE pd2": End Sub
Public Sub ONGEN_RunDevice5B(): RunStep "DeviceWF_5B_Build_FieldIO_FAST", "Phase 5B build Field I/O", "Specification Devices": End Sub
Public Sub ONGEN_RunDevice5C(): RunStep "DeviceWF_5C_Send_Run_B03", "Phase 5C run B03", "Field Device IO ONGEN": End Sub
Public Sub ONGEN_RunDevice5D(): RunStep "DeviceWF_5D_Import_B03_Return_FAST", "Phase 5D import B03", "Field Device IO ONGEN": End Sub
Public Sub ONGEN_RunExportPD6(): RunStep "Export_PD6_Export_System_To_CSV", "Export PD6 CSV", "PD6 System Designed": End Sub
Public Sub ONGEN_RunExportIO(): RunStep "Export_IOBLOCK_To_CSV", "Export I/O block CSV", "": End Sub
Public Sub ONGEN_OpenLegacyDashboard(): RunStep "Build_Main_Dashboard", "Open legacy dashboard", "": End Sub

Private Sub RunStep(ByVal macroName As String, ByVal stepName As String, ByVal requiredSheet As String)
    Dim oldCalc As XlCalculation, oldEvents As Boolean, oldScreen As Boolean
    Dim oldAlerts As Boolean, started As Date, message As String
    If mBusy Then
        MsgBox "Another ONGEN action is already running.", vbExclamation, "ONGEN"
        Exit Sub
    End If
    If Len(requiredSheet) > 0 And Not SheetExists(requiredSheet) Then
        message = "Required sheet not found: " & requiredSheet
        AppendLog stepName, "BLOCKED", message
        MsgBox message, vbExclamation, "ONGEN preflight"
        Exit Sub
    End If
    If MsgBox("Run " & stepName & "?" & vbCrLf & vbCrLf & _
              "Save a controlled backup before accepting generated engineering outputs.", _
              vbQuestion + vbYesNo, "ONGEN") <> vbYes Then Exit Sub

    On Error GoTo Failed
    mBusy = True
    started = Now
    oldCalc = Application.Calculation
    oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating
    oldAlerts = Application.DisplayAlerts
    Application.StatusBar = "ONGEN: " & stepName
    Application.Run "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & macroName
    AppendLog stepName, "RETURNED", "Macro returned control. Review its output and validation messages."
    MsgBox stepName & " returned control to ONGEN." & vbCrLf & _
           "Review the generated sheets, logs, and totals before release.", vbInformation, "ONGEN"
CleanExit:
    Application.Calculation = oldCalc
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    Application.DisplayAlerts = oldAlerts
    Application.StatusBar = False
    mBusy = False
    Exit Sub
Failed:
    message = "Error " & CStr(Err.Number) & ": " & Err.Description
    AppendLog stepName, "ERROR", message
    MsgBox stepName & " stopped." & vbCrLf & message, vbCritical, "ONGEN"
    Resume CleanExit
End Sub

Private Sub AddActionRow(ByVal ws As Worksheet, ByVal rowNumber As Long, ByVal label As String, ByVal macroName As String, ByVal note As String)
    Dim shp As Shape
    ws.Cells(rowNumber, 1).Value = label
    ws.Cells(rowNumber, 1).Font.Bold = True
    ws.Range(ws.Cells(rowNumber, 5), ws.Cells(rowNumber + 1, 8)).Merge
    ws.Cells(rowNumber, 5).Value = note
    ws.Cells(rowNumber, 5).WrapText = True
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, ws.Cells(rowNumber, 2).Left, _
                                 ws.Cells(rowNumber, 2).Top, 210, 28)
    With shp
        .Name = "ongen_" & Format$(rowNumber, "00")
        .TextFrame2.TextRange.Text = "RUN"
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = vbWhite
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .Fill.ForeColor.RGB = RGB(0, 120, 104)
        .Line.Visible = msoFalse
        .OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & macroName
    End With
End Sub

Private Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        GetOrCreateSheet.Name = sheetName
    End If
End Function

Private Function SheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

Private Function FileExistsSafe(ByVal filePath As String) As Boolean
    Dim fileSystem As Object
    On Error GoTo NotFound
    If Len(Trim$(filePath)) = 0 Then Exit Function
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    FileExistsSafe = fileSystem.FileExists(filePath)
CleanExit:
    Set fileSystem = Nothing
    Exit Function
NotFound:
    FileExistsSafe = False
    Resume CleanExit
End Function

Private Sub AppendLog(ByVal actionName As String, ByVal result As String, ByVal details As String)
    Dim ws As Worksheet, nextRow As Long
    Set ws = GetOrCreateSheet(LOG_SHEET)
    If Len(CStr(ws.Cells(1, 1).Value)) = 0 Then
        ws.Range("A1:D1").Value = Array("Timestamp", "Action", "Result", "Details")
        ws.Range("A1:D1").Font.Bold = True
    End If
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(nextRow, 1).Value = Now
    ws.Cells(nextRow, 2).Value = actionName
    ws.Cells(nextRow, 3).Value = result
    ws.Cells(nextRow, 4).Value = details
    ws.Columns("A:D").EntireColumn.AutoFit
End Sub
