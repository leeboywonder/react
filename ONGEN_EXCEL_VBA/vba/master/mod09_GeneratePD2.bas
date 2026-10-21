Attribute VB_Name = "mod09_GeneratePD2"
Option Explicit

'====================================================================================
' RELEASED V1.6.21.26
'
' Macro:
'   Copy_PD5_To_Import_And_Fill_From_BOQ
'
' Purpose:
'   - Copy matching columns from "SIS_TAG_PD5" to "SIS_TAG_IMPORT_PD2"
'   - Create "SIS_TAG_IMPORT_PD2" if missing
'   - Add required headers if missing
'   - Fill BOQ data by matching:
'       SIS_TAG_IMPORT_PD2[LPN] = BOQ[Layout Position Number]
'   - TYPE from SIS_TAG_PD5 is preserved
'   - BOQ only fills TYPE if destination TYPE is blank
'   - Fills CONTROL_FUNCTION_NUMBER from FUNCTION DEFINITION by matching CONVEYOR TYPE
'   - Carries INPUT / OUTPUT / S_IN / S_OUT as part of the PD2 data contract
'   - If PD5 already supplies an IO value, that value is preserved
'   - If a PD2 IO value is blank and FUNCTION DEFINITION has IO headers,
'     the parent-level total for the resolved CONTROL_FUNCTION_NUMBER is filled
'   - Makes all headers bold
'   - Makes CONTROL_FUNCTION_NUMBER column yellow
'
' Required Sheets:
'   - SIS_TAG_PD5
'   - BOQ
'   - FUNCTION DEFINITION
'
'====================================================================================

Public Sub Copy_PD5_To_Import_And_Fill_From_BOQ()
    If PDPerf_IsYielding Then Exit Sub

    Dim wb As Workbook
    Dim wsSrc As Worksheet
    Dim wsDest As Worksheet
    Dim wsBOQ As Worksheet
    Dim wsFuncDef As Worksheet
    
    Dim srcHeaderRow As Long
    Dim destHeaderRow As Long
    Dim boqHeaderRow As Long
    Dim funcDefHeaderRow As Long
    
    Dim lastRowSrc As Long
    Dim lastRowDest As Long
    Dim lastRowExisting As Long
    Dim lastRowBOQ As Long
    Dim lastColSrc As Long
    Dim lastColDest As Long
    Dim lastColBOQ As Long
    Dim lastRowFuncDef As Long
    Dim lastColFuncDef As Long
    
    Dim srcData As Variant
    Dim boqData As Variant
    Dim destData As Variant
    Dim funcDefData As Variant
    
    Dim srcHeaders As Object
    Dim destHeaders As Object
    Dim boqHeaders As Object
    Dim dictBOQ As Object
    Dim funcDefHeaders As Object
    Dim dictCFNByConveyorType As Object
    Dim dictIOByCFN As Object
    
    Dim srcCol As Long
    Dim destCol As Long
    Dim r As Long
    Dim addedHeaders As Long
    Dim copiedCols As Long
    Dim filledBOQRows As Long
    Dim filledCFNRows As Long
    Dim filledIORows As Long
    
    Dim key As String
    Dim cfnKey As String
    Dim arrData As Variant
    Dim ioAgg As Variant
    Dim ioKey As String
    Dim lastFuncKey As String
    
    Dim colDestLPN As Long
    Dim colDestPower As Long
    Dim colDestVoltage As Long
    Dim colDestType As Long
    Dim colDestConveyorType As Long
    Dim colDestCFN As Long
    Dim colDestModuleLength As Long
    Dim colDestInput As Long, colDestOutput As Long
    Dim colDestSIn As Long, colDestSOut As Long
    
    Dim colBOQLPN As Long
    Dim colBOQPower As Long
    Dim colBOQVoltage As Long
    Dim colBOQType As Long
    Dim colBOQConveyorType As Long
    Dim colBOQModuleLength As Long

    Dim colFuncDefConveyorType As Long
    Dim colFuncDefCFN As Long
    Dim colFuncDefInput As Long, colFuncDefOutput As Long
    Dim colFuncDefSIn As Long, colFuncDefSOut As Long
    
    On Error GoTo ErrHandler
    
    Set wb = ThisWorkbook
    
    Set wsSrc = wb.Worksheets("SIS_TAG_PD5")
    Set wsBOQ = wb.Worksheets("BOQ")
    Set wsFuncDef = GetWorksheetByName(wb, "FUNCTION DEFINITION")
    Set wsDest = GetOrCreatePD2Sheet(wb, "SIS_TAG_IMPORT_PD2")
    
    srcHeaderRow = 1
    destHeaderRow = 1
    boqHeaderRow = 1
    funcDefHeaderRow = 1
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    '---------------------------------------
    ' Build destination required headers
    '---------------------------------------
    EnsurePD2Headers wsDest
    FormatPD2Sheet wsDest
    
    '---------------------------------------
    ' Get sheet sizes
    '---------------------------------------
    lastRowSrc = LastUsedRow(wsSrc)
    lastColSrc = LastUsedCol(wsSrc, srcHeaderRow)
    
    lastRowBOQ = LastUsedRow(wsBOQ)
    lastColBOQ = LastUsedCol(wsBOQ, boqHeaderRow)

    If Not wsFuncDef Is Nothing Then
        lastRowFuncDef = LastUsedRow(wsFuncDef)
        lastColFuncDef = LastUsedCol(wsFuncDef, funcDefHeaderRow)
    End If
    
    '---------------------------------------
    ' Build source/destination/BOQ header maps
    '---------------------------------------
    Set srcHeaders = headerMap(wsSrc, srcHeaderRow)
    Set destHeaders = headerMap(wsDest, destHeaderRow)
    Set boqHeaders = headerMap(wsBOQ, boqHeaderRow)
    If Not wsFuncDef Is Nothing Then Set funcDefHeaders = headerMap(wsFuncDef, funcDefHeaderRow)
    
    '---------------------------------------
    ' Add any missing headers from SIS_TAG_PD5
    '---------------------------------------
    For srcCol = 1 To lastColSrc
        
        Dim h As String
        h = Trim$(CStr(wsSrc.Cells(srcHeaderRow, srcCol).Value))
        
        If Len(h) > 0 Then
            If Not destHeaders.Exists(UCase$(h)) Then
                wsDest.Cells(destHeaderRow, LastUsedCol(wsDest, destHeaderRow) + 1).Value = h
                addedHeaders = addedHeaders + 1
            End If
        End If
        
    Next srcCol
    
    'Rebuild destination header map after adding headers
    Set destHeaders = headerMap(wsDest, destHeaderRow)
    lastColDest = LastUsedCol(wsDest, destHeaderRow)
    
    '---------------------------------------
    ' Copy matching columns from SIS_TAG_PD5 to destination
    ' Faster: copies each whole column range at once
    '---------------------------------------
    If lastRowSrc >= 2 Then
        lastRowExisting = LastUsedRow(wsDest)
        
        For srcCol = 1 To lastColSrc
            
            h = Trim$(CStr(wsSrc.Cells(srcHeaderRow, srcCol).Value))
            
            If Len(h) > 0 Then
                
                If destHeaders.Exists(UCase$(h)) Then
                    
                    destCol = CLng(destHeaders(UCase$(h)))
                    
                    If lastRowExisting >= 2 Then
                        wsDest.Range(wsDest.Cells(2, destCol), wsDest.Cells(lastRowExisting, destCol)).ClearContents
                    End If
                    
                    wsDest.Range(wsDest.Cells(2, destCol), wsDest.Cells(lastRowSrc, destCol)).Value = _
                        wsSrc.Range(wsSrc.Cells(2, srcCol), wsSrc.Cells(lastRowSrc, srcCol)).Value
                    
                    copiedCols = copiedCols + 1
                    
                End If
                
            End If
            
        Next srcCol
        
    End If
    
    '---------------------------------------
    ' Ensure and find destination columns
    '---------------------------------------
    colDestPower = EnsureHeader(wsDest, destHeaderRow, "POWER")
    colDestVoltage = EnsureHeader(wsDest, destHeaderRow, "VOLTAGE")
    colDestType = EnsureHeader(wsDest, destHeaderRow, "TYPE")
    colDestConveyorType = EnsureHeader(wsDest, destHeaderRow, "CONVEYOR TYPE")
    colDestCFN = EnsureHeader(wsDest, destHeaderRow, "CONTROL_FUNCTION_NUMBER")
    colDestModuleLength = EnsureHeader(wsDest, destHeaderRow, "Module Length (ML)")
    colDestInput = EnsureHeader(wsDest, destHeaderRow, "INPUT")
    colDestOutput = EnsureHeader(wsDest, destHeaderRow, "OUTPUT")
    colDestSIn = EnsureHeader(wsDest, destHeaderRow, "S_IN")
    colDestSOut = EnsureHeader(wsDest, destHeaderRow, "S_OUT")
    
    Set destHeaders = headerMap(wsDest, destHeaderRow)
    
    If Not destHeaders.Exists("LPN") Then
        MsgBox "Destination header 'LPN' not found in SIS_TAG_IMPORT_PD2.", vbExclamation
        GoTo SafeExit
    End If
    
    colDestLPN = CLng(destHeaders("LPN"))
    
    '---------------------------------------
    ' Find BOQ columns
    '---------------------------------------
    If Not boqHeaders.Exists("LAYOUT POSITION NUMBER") Then
        MsgBox "BOQ header 'Layout Position Number' not found.", vbExclamation
        GoTo SafeExit
    End If
    
    If Not boqHeaders.Exists("MOTOR POWER (MAIN DRIVE)") Then
        MsgBox "BOQ header 'Motor Power (Main Drive)' not found.", vbExclamation
        GoTo SafeExit
    End If
    
    If Not boqHeaders.Exists("MOTOR VOLTAGE (MAIN DRIVE)") Then
        MsgBox "BOQ header 'Motor Voltage (Main Drive)' not found.", vbExclamation
        GoTo SafeExit
    End If
    
    If Not boqHeaders.Exists("CONTROL CARD ABBREVIATION") Then
        MsgBox "BOQ header 'Control Card Abbreviation' not found.", vbExclamation
        GoTo SafeExit
    End If
    
    If Not boqHeaders.Exists("CONVEYOR TYPE") Then
        MsgBox "BOQ header 'Conveyor Type' not found.", vbExclamation
        GoTo SafeExit
    End If
    
    If Not boqHeaders.Exists("MODULE LENGTH (ML)") Then
        MsgBox "BOQ header 'Module Length (ML)' not found.", vbExclamation
        GoTo SafeExit
    End If
    
    colBOQLPN = CLng(boqHeaders("LAYOUT POSITION NUMBER"))
    colBOQPower = CLng(boqHeaders("MOTOR POWER (MAIN DRIVE)"))
    colBOQVoltage = CLng(boqHeaders("MOTOR VOLTAGE (MAIN DRIVE)"))
    colBOQType = CLng(boqHeaders("CONTROL CARD ABBREVIATION"))
    colBOQConveyorType = CLng(boqHeaders("CONVEYOR TYPE"))
    colBOQModuleLength = CLng(boqHeaders("MODULE LENGTH (ML)"))
    
    '---------------------------------------
    ' Build BOQ dictionary
    ' Faster: read BOQ to array once
    '---------------------------------------
    Set dictBOQ = CreateObject("Scripting.Dictionary")
    dictBOQ.CompareMode = vbTextCompare
    
    If lastRowBOQ >= 2 Then
        
        boqData = wsBOQ.Range(wsBOQ.Cells(1, 1), wsBOQ.Cells(lastRowBOQ, lastColBOQ)).Value
        
        For r = 2 To UBound(boqData, 1)
            PDPerf_Pulse "mod09_GeneratePD2.Copy_PD5_To_Import_And_Fill_From_BOQ [L278]" & " row " & CStr(r)
            
            key = NormalizeKey(boqData(r, colBOQLPN))
            
            If Len(key) > 0 Then
                
                arrData = Array( _
                    boqData(r, colBOQPower), _
                    boqData(r, colBOQVoltage), _
                    boqData(r, colBOQType), _
                    boqData(r, colBOQConveyorType), _
                    boqData(r, colBOQModuleLength))
                
                dictBOQ(key) = arrData
                
            End If
            
        Next r
        
    End If
    
    '---------------------------------------
    ' Build FUNCTION DEFINITION dictionaries
    ' 1) Conveyor Type -> CONTROL_FUNCTION_NUMBER
    ' 2) CONTROL_FUNCTION_NUMBER -> parent IO totals
    '
    ' IO precedence at Step 3A:
    '   - Existing/copied SIS_TAG_PD5 value wins.
    '   - FUNCTION DEFINITION only fills a blank PD2 IO value.
    '---------------------------------------
    Set dictCFNByConveyorType = CreateObject("Scripting.Dictionary")
    dictCFNByConveyorType.CompareMode = vbTextCompare
    Set dictIOByCFN = CreateObject("Scripting.Dictionary")
    dictIOByCFN.CompareMode = vbTextCompare

    If Not wsFuncDef Is Nothing Then
        If funcDefHeaders.Exists("CONTROL_FUNCTION_NUMBER") Then
            colFuncDefCFN = CLng(funcDefHeaders("CONTROL_FUNCTION_NUMBER"))
            If funcDefHeaders.Exists("CONVEYOR TYPE") Then colFuncDefConveyorType = CLng(funcDefHeaders("CONVEYOR TYPE"))
            If funcDefHeaders.Exists("INPUT") Then colFuncDefInput = CLng(funcDefHeaders("INPUT"))
            If funcDefHeaders.Exists("OUTPUT") Then colFuncDefOutput = CLng(funcDefHeaders("OUTPUT"))
            If funcDefHeaders.Exists("S_IN") Then colFuncDefSIn = CLng(funcDefHeaders("S_IN"))
            If funcDefHeaders.Exists("S_OUT") Then colFuncDefSOut = CLng(funcDefHeaders("S_OUT"))

            If lastRowFuncDef >= 2 Then
                funcDefData = wsFuncDef.Range(wsFuncDef.Cells(1, 1), wsFuncDef.Cells(lastRowFuncDef, lastColFuncDef)).Value
                lastFuncKey = vbNullString

                For r = 2 To UBound(funcDefData, 1)
                    PDPerf_Pulse "mod09_GeneratePD2.Copy_PD5_To_Import_And_Fill_From_BOQ [L326]" & " row " & CStr(r)
                    'The function number may be repeated on every row or only on the first row of a group.
                    If Len(Trim$(CStr(funcDefData(r, colFuncDefCFN)))) > 0 Then
                        lastFuncKey = NormalizeKey(funcDefData(r, colFuncDefCFN))
                    End If

                    'Conveyor type lookup used by Step 3A.
                    If colFuncDefConveyorType > 0 Then
                        key = NormalizeKey(funcDefData(r, colFuncDefConveyorType))
                        If Len(key) > 0 And Len(lastFuncKey) > 0 Then
                            If Not dictCFNByConveyorType.Exists(key) Then
                                dictCFNByConveyorType.Add key, funcDefData(r, colFuncDefCFN)
                                If Len(Trim$(CStr(dictCFNByConveyorType(key)))) = 0 Then dictCFNByConveyorType(key) = lastFuncKey
                            End If
                        End If
                    End If

                    'Parent-level IO totals. These are useful at SIS_TAG_IMPORT_PD2 level.
                    'Per-device IO remains row-level and is applied by Step 3B.
                    If Len(lastFuncKey) > 0 Then
                        If dictIOByCFN.Exists(lastFuncKey) Then
                            ioAgg = dictIOByCFN(lastFuncKey)
                        Else
                            ioAgg = Array(0#, 0#, 0#, 0#, False)
                        End If

                        If colFuncDefInput > 0 Then PD2_AddIOValue ioAgg, 0, funcDefData(r, colFuncDefInput)
                        If colFuncDefOutput > 0 Then PD2_AddIOValue ioAgg, 1, funcDefData(r, colFuncDefOutput)
                        If colFuncDefSIn > 0 Then PD2_AddIOValue ioAgg, 2, funcDefData(r, colFuncDefSIn)
                        If colFuncDefSOut > 0 Then PD2_AddIOValue ioAgg, 3, funcDefData(r, colFuncDefSOut)
                        dictIOByCFN(lastFuncKey) = ioAgg
                    End If
                Next r
            End If
        End If
    End If

    '---------------------------------------
    ' Fill destination from BOQ
    ' Faster: read destination into array, update array, write back once
    '---------------------------------------
    lastRowDest = LastUsedRow(wsDest)
    lastColDest = LastUsedCol(wsDest, destHeaderRow)
    
    If lastRowDest >= 2 Then
        
        destData = wsDest.Range(wsDest.Cells(1, 1), wsDest.Cells(lastRowDest, lastColDest)).Value
        
        For r = 2 To UBound(destData, 1)
            PDPerf_Pulse "mod09_GeneratePD2.Copy_PD5_To_Import_And_Fill_From_BOQ [L374]" & " row " & CStr(r)
            
            key = NormalizeKey(destData(r, colDestLPN))
            
            If Len(key) > 0 Then
                
                If dictBOQ.Exists(key) Then
                    
                    arrData = dictBOQ(key)
                    
                    destData(r, colDestPower) = arrData(0)
                    destData(r, colDestVoltage) = arrData(1)
                    
                    'TYPE is preserved from SIS_TAG_PD5.
                    'Only fill TYPE from BOQ if destination TYPE is blank.
                    If Len(Trim$(CStr(destData(r, colDestType)))) = 0 Then
                        destData(r, colDestType) = arrData(2)
                    End If
                    
                    destData(r, colDestConveyorType) = arrData(3)
                    destData(r, colDestModuleLength) = arrData(4)
                    
                    filledBOQRows = filledBOQRows + 1
                    
                End If
                
            End If

            'Fill CONTROL_FUNCTION_NUMBER using the final CONVEYOR TYPE value.
            'Example: SIS_TAG_IMPORT_PD2[CONVEYOR TYPE] = Divert -> lookup Divert in FUNCTION DEFINITION.
            cfnKey = NormalizeKey(destData(r, colDestConveyorType))

            If Len(cfnKey) > 0 Then
                If dictCFNByConveyorType.Exists(cfnKey) Then
                    destData(r, colDestCFN) = dictCFNByConveyorType(cfnKey)
                    filledCFNRows = filledCFNRows + 1
                End If
            End If

            'Fill missing parent IO totals from FUNCTION DEFINITION.
            'Never overwrite a nonblank value that arrived from SIS_TAG_PD5.
            ioKey = NormalizeKey(destData(r, colDestCFN))
            If Len(ioKey) > 0 Then
                If dictIOByCFN.Exists(ioKey) Then
                    ioAgg = dictIOByCFN(ioKey)
                    If CBool(ioAgg(4)) Then
                        If Not PD2_HasValue(destData(r, colDestInput)) Then destData(r, colDestInput) = ioAgg(0)
                        If Not PD2_HasValue(destData(r, colDestOutput)) Then destData(r, colDestOutput) = ioAgg(1)
                        If Not PD2_HasValue(destData(r, colDestSIn)) Then destData(r, colDestSIn) = ioAgg(2)
                        If Not PD2_HasValue(destData(r, colDestSOut)) Then destData(r, colDestSOut) = ioAgg(3)
                        filledIORows = filledIORows + 1
                    End If
                End If
            End If
            
        Next r
        
        wsDest.Range(wsDest.Cells(1, 1), wsDest.Cells(lastRowDest, lastColDest)).Value = destData
        
    End If
    
    FormatPD2Sheet wsDest
    
    MsgBox "Done." & vbCrLf & vbCrLf & _
           "Destination Sheet: SIS_TAG_IMPORT_PD2" & vbCrLf & _
           "Headers added from SIS_TAG_PD5: " & addedHeaders & vbCrLf & _
           "Columns copied from SIS_TAG_PD5: " & copiedCols & vbCrLf & _
           "Rows filled from BOQ using LPN: " & filledBOQRows & vbCrLf & _
           "CONTROL_FUNCTION_NUMBER rows filled from FUNCTION DEFINITION: " & filledCFNRows & vbCrLf & _
           "PD2 rows with INPUT/OUTPUT/S_IN/S_OUT evaluated: " & filledIORows, vbInformation

SafeExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    If Not PDPerf_IsRunning Then Application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.description, vbCritical
    Resume SafeExit

End Sub

'====================================================================================
' Gets worksheet by name using a case-insensitive match.
' Returns Nothing if the sheet is missing.
'====================================================================================
Private Function GetWorksheetByName(wb As Workbook, sheetName As String) As Worksheet

    Dim ws As Worksheet

    For Each ws In wb.Worksheets
        PDPerf_Pulse "mod09_GeneratePD2.GetWorksheetByName [L465]"
        If StrComp(Trim$(ws.name), Trim$(sheetName), vbTextCompare) = 0 Then
            Set GetWorksheetByName = ws
            Exit Function
        End If
    Next ws

    Set GetWorksheetByName = Nothing

End Function

'====================================================================================
' Creates destination sheet if missing.
'====================================================================================
Private Function GetOrCreatePD2Sheet(wb As Workbook, sheetName As String) As Worksheet

    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
        ws.name = sheetName
        If wb Is ThisWorkbook Then mod23_WorksheetNavigation.NAV_Ensure_Worksheet_Navigation ws
        WritePD2Headers ws
    End If
    
    Set GetOrCreatePD2Sheet = ws

End Function

'====================================================================================
' Writes released headers to new SIS_TAG_IMPORT_PD2.
' Uses Split() to avoid too many line continuations.
'====================================================================================
Private Sub WritePD2Headers(ws As Worksheet)

    Dim headers As Variant
    Dim i As Long
    
    headers = PD2HeaderList()
    
    For i = LBound(headers) To UBound(headers)
        ws.Cells(1, i + 1).Value = headers(i)
    Next i

End Sub

'====================================================================================
' Adds missing required headers.
'====================================================================================
Private Sub EnsurePD2Headers(ws As Worksheet)

    Dim headers As Variant
    Dim dict As Object
    Dim i As Long
    Dim nextCol As Long
    
    headers = PD2HeaderList()
    Set dict = headerMap(ws, 1)
    
    For i = LBound(headers) To UBound(headers)
        
        If Not dict.Exists(UCase$(CStr(headers(i)))) Then
            nextCol = LastUsedCol(ws, 1) + 1
            ws.Cells(1, nextCol).Value = headers(i)
            dict.Add UCase$(CStr(headers(i))), nextCol
        End If
        
    Next i

End Sub

'====================================================================================
' Required headers list.
' Split() prevents VBA "too many line continuations" error.
'====================================================================================
Private Function PD2HeaderList() As Variant

    Dim s As String
    
    s = "HANDLE|BLOCK_NAME|SIS_SEQUENCE|SIS_ELEMENT_NUMBER|AREA_NUMBER|SWITCHGROUP|TYPE"
    s = s & "|CONVEYOR TYPE|CONTROL_FUNCTION_NUMBER|CONTROL_FUNCTION_TAG|PNL|NETWORK_IP|LPN|LAYER"
    s = s & "|X|Y|Z|ROTATION_DEG|Module Length (ML)|Motor Power (Main Drive)"
    s = s & "|motor Voltage (Main Drive)|motor brake voltage (Main drive)|Motor Protection (Main Drive)"
    s = s & "|SEQUENCE_CLEANED|CONTROL_ID|CONTROL_TYPE|Load|POWER|VOLTAGE|SYSTEM TYPE"
    s = s & "|Encoder Type (Main Drive)|Total Qty of Sensors|Control Card|Qty of Control Card|Total Qty of Drives"
    s = s & "|INPUT|OUTPUT|S_IN|S_OUT"
    
    PD2HeaderList = Split(s, "|")

End Function

'====================================================================================
' Adds one FUNCTION DEFINITION IO value to a parent-level aggregate.
' Only numeric values are summed. Blank/non-numeric entries are ignored.
' ioAgg(4) records whether at least one real IO value was supplied.
'====================================================================================
Private Sub PD2_AddIOValue(ByRef ioAgg As Variant, ByVal index As Long, ByVal v As Variant)
    Dim txt As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then Exit Sub
    txt = Trim$(CStr(v))
    If Len(txt) = 0 Then Exit Sub
    If IsNumeric(v) Then
        ioAgg(index) = CDbl(ioAgg(index)) + CDbl(v)
        ioAgg(4) = True
    ElseIf IsNumeric(txt) Then
        ioAgg(index) = CDbl(ioAgg(index)) + CDbl(txt)
        ioAgg(4) = True
    End If
End Sub

Private Function PD2_HasValue(ByVal v As Variant) As Boolean
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then Exit Function
    PD2_HasValue = (Len(Trim$(CStr(v))) > 0)
End Function

'====================================================================================
' Formats destination:
'   - All headers bold
'   - CONTROL_FUNCTION_NUMBER column yellow
'====================================================================================
Private Sub FormatPD2Sheet(ws As Worksheet)

    Dim colCFN As Long
    Dim lastCol As Long
    
    lastCol = LastUsedCol(ws, 1)
    
    ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol)).Font.Bold = True
    
    colCFN = FindHeaderColumn(ws, 1, "CONTROL_FUNCTION_NUMBER")
    
    If colCFN > 0 Then
        ws.Columns(colCFN).Interior.Color = RGB(255, 255, 0)
        ws.Cells(1, colCFN).Interior.Color = RGB(255, 192, 0)
    End If
    
    ws.UsedRange.Columns.AutoFit

End Sub

'====================================================================================
' Creates a dictionary of headers:
' Key = UCASE(header text)
' Item = column number
'====================================================================================
Private Function headerMap(ws As Worksheet, headerRow As Long) As Object

    Dim dict As Object
    Dim lastCol As Long
    Dim c As Long
    Dim h As String
    
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    
    lastCol = LastUsedCol(ws, headerRow)
    
    For c = 1 To lastCol
        
        h = Trim$(CStr(ws.Cells(headerRow, c).Value))
        
        If Len(h) > 0 Then
            If Not dict.Exists(UCase$(h)) Then
                dict.Add UCase$(h), c
            End If
        End If
        
    Next c
    
    Set headerMap = dict

End Function

'====================================================================================
' Ensures one header exists and returns column.
'====================================================================================
Private Function EnsureHeader(ws As Worksheet, headerRow As Long, headerText As String) As Long

    Dim dict As Object
    Dim nextCol As Long
    
    Set dict = headerMap(ws, headerRow)
    
    If dict.Exists(UCase$(headerText)) Then
        EnsureHeader = CLng(dict(UCase$(headerText)))
    Else
        nextCol = LastUsedCol(ws, headerRow) + 1
        ws.Cells(headerRow, nextCol).Value = headerText
        EnsureHeader = nextCol
    End If

End Function

'====================================================================================
' Normalizes matching keys.
'====================================================================================
Private Function NormalizeKey(ByVal v As Variant) As String

    Dim s As String
    
    s = CStr(v)
    s = Replace(s, Chr(160), " ")
    s = Replace(s, vbTab, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Trim$(s)
    
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    
    NormalizeKey = UCase$(s)

End Function

'====================================================================================
' Finds header column.
'====================================================================================
Private Function FindHeaderColumn(ws As Worksheet, headerRow As Long, headerText As String) As Long

    Dim dict As Object
    
    Set dict = headerMap(ws, headerRow)
    
    If dict.Exists(UCase$(headerText)) Then
        FindHeaderColumn = CLng(dict(UCase$(headerText)))
    Else
        FindHeaderColumn = 0
    End If

End Function

'====================================================================================
' Last used row.
'====================================================================================
Private Function LastUsedRow(ws As Worksheet) As Long

    Dim f As Range
    
    On Error Resume Next
    
    Set f = ws.Cells.Find(What:="*", _
                          After:=ws.Cells(1, 1), _
                          LookIn:=xlFormulas, _
                          LookAt:=xlPart, _
                          SearchOrder:=xlByRows, _
                          SearchDirection:=xlPrevious)
    
    On Error GoTo 0
    
    If f Is Nothing Then
        LastUsedRow = 1
    Else
        LastUsedRow = f.Row
    End If

End Function

'====================================================================================
' Last used column on a header row.
'====================================================================================
Private Function LastUsedCol(ws As Worksheet, Optional headerRow As Long = 1) As Long

    Dim f As Range
    
    On Error Resume Next
    
    Set f = ws.rows(headerRow).Find(What:="*", _
                                    After:=ws.Cells(headerRow, 1), _
                                    LookIn:=xlFormulas, _
                                    LookAt:=xlPart, _
                                    SearchOrder:=xlByColumns, _
                                    SearchDirection:=xlPrevious)
    
    On Error GoTo 0
    
    If f Is Nothing Then
        LastUsedCol = 1
    Else
        LastUsedCol = f.Column
    End If

End Function


