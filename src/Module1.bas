Attribute VB_Name = "Module1"
Option Explicit

'--- WIN64/WIN32 API Safe Sleep Declaration (Fixed PtrSafe Typo) ---
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

' Global message filter tracking variable
Private oMsgFilter As Object

Public Sub KillMessageFilter()
    On Error Resume Next
    Set oMsgFilter = Nothing
    On Error GoTo 0
End Sub

'==============================================================================
' MAIN MASTER ENTRY POINT
'==============================================================================
Public Sub RunAll()

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Call Stage1_RunZCO05
    Call Stage2_CleanExcel
    Call Stage3_COR2_VARUpdate

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

'==============================================================================
' STAGE 1 — Export ZCO05 (Using Menu Recording & Dynamic File Tracking)
'==============================================================================
Public Sub Stage1_RunZCO05()

    Dim SapGuiAuto  As Object
    Dim Application As Object
    Dim Connection  As Object
    Dim Session     As Object
    Dim waitSecs    As Long
    Dim sLatestFile As String

    ' Direct error handling connection gate
    On Error GoTo ErrorHandler

    If SapGuiAuto Is Nothing Then
        Set SapGuiAuto = GetObject("SAPGUI")
    End If
    
    If Application Is Nothing Then
        Set Application = SapGuiAuto.GetScriptingEngine
    End If
    
    If Connection Is Nothing Then
        Set Connection = Application.Children(0)
    End If
    
    If Session Is Nothing Then
        Set Session = Connection.Children(0)
    End If

    ' --- Start of your recorded SAP Actions ---
    Session.findById("wnd[0]").maximize
    Session.findById("wnd[0]/tbar[0]/okcd").Text = "/NZCO05"
    Session.findById("wnd[0]/tbar[0]/btn[0]").press
    Session.findById("wnd[0]/usr/ctxtP_PLANT-LOW").Text = "1900"
    Session.findById("wnd[0]/usr/txtP_USER-LOW").Text = "WIP"
    Session.findById("wnd[0]/usr/txtP_USER-LOW").SetFocus
    Session.findById("wnd[0]/usr/txtP_USER-LOW").CaretPosition = 3
    Session.findById("wnd[0]/tbar[1]/btn[8]").press
    
    Sleep 2000
    
    ' Trigger Menu Export Operation
    Session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").Select
    Sleep 1500
    
    ' Set target directory path properties
    Session.findById("wnd[1]/usr/ctxtDY_PATH").Text = "C:\TEMP"
    Session.findById("wnd[1]/usr/ctxtDY_PATH").SetFocus
    Session.findById("wnd[1]/usr/ctxtDY_PATH").CaretPosition = 7
    
    ' Execute download selection click
    Session.findById("wnd[1]/tbar[0]/btn[0]").press
    ' --- End of your recorded SAP Actions ---
    
    ' Dynamic Wait Loop: Wait up to 30 seconds for the newest EXPORT file to land
    waitSecs = 0
    Do While waitSecs < 30
        DoEvents
        sLatestFile = GetLatestSAPExport("C:\TEMP\")
        If sLatestFile <> "" Then Exit Do
        Sleep 1000
        waitSecs = waitSecs + 1
    Loop

    If sLatestFile = "" Then
        MsgBox "Export failed: No recent EXPORT_*.xlsx asset found in C:\TEMP folder.", vbCritical, "Data Export Error"
        End
    End If
    
    Exit Sub

ErrorHandler:
    MsgBox "Error in Stage 1. Please make sure SAP is open and logged in. Error: " & Err.Description, vbCritical, "SAP Error"
    End
End Sub

'==============================================================================
' STAGE 2 — Open Newest Excel File, Filter Matrix, and Isolate Targets
'==============================================================================
Public Sub Stage2_CleanExcel()

    Dim wbSource    As Workbook
    Dim wsSource    As Worksheet
    Dim wsFormat    As Worksheet
    Dim wsTemp      As Worksheet
    Dim lastRow     As Long
    Dim filterRange As String
    Dim sTargetFile As String

    ' Automatically identify the exact file SAP just produced
    sTargetFile = GetLatestSAPExport("C:\TEMP\")

    If sTargetFile = "" Then
        MsgBox "Source data file asset missing inside C:\TEMP folder location.", vbCritical, "Execution Exception"
        Exit Sub
    End If

    ' Safe Sheet Handling: Avoid "Name Taken" errors by cleaning instead of deleting
    On Error Resume Next
    Set wsFormat = ThisWorkbook.Sheets("Format")
    On Error GoTo 0
    If wsFormat Is Nothing Then
        Set wsFormat = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
        wsFormat.Name = "Format"
    Else
        wsFormat.Cells.Clear
    End If

    On Error Resume Next
    Set wsTemp = ThisWorkbook.Sheets("Temp")
    On Error GoTo 0
    If wsTemp Is Nothing Then
        Set wsTemp = ThisWorkbook.Sheets.Add(After:=wsFormat)
        wsTemp.Name = "Temp"
    Else
        wsTemp.Cells.Clear
    End If

    ' Open natively as a true Excel Workbook to preserve column alignment
    Set wbSource = Workbooks.Open(FileName:=sTargetFile, UpdateLinks:=False, ReadOnly:=True)
    Set wsSource = wbSource.Sheets(1)

    ' Mirror data layout cleanly
    wsSource.Cells.Copy Destination:=wsFormat.Range("A1")
    wbSource.Close SaveChanges:=False

    lastRow = wsFormat.Cells(wsFormat.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "The extracted layout contains no data rows to process.", vbExclamation
        Exit Sub
    End If
    
    filterRange = "$A$1:$T$" & lastRow

    ' Setup header labels
    wsTemp.Range("A1").Value = wsFormat.Range("C1").Value
    wsTemp.Range("B1").Value = "Results"
    wsTemp.Range("A1").Font.Bold = True
    wsTemp.Range("B1").Font.Bold = True

    If wsFormat.AutoFilterMode Then wsFormat.AutoFilterMode = False

    ' --- RUN FILTER PROFILE SEQUENCE 1 ---
    wsFormat.Range(filterRange).AutoFilter Field:=2, Criteria1:=Array( _
        "032", "033", "040", "061", "065"), Operator:=xlFilterValues
    wsFormat.Range(filterRange).AutoFilter Field:=5, Criteria1:=Array( _
        "ZP01", "ZP03", "ZP04", "ZP06", "ZP07", "ZP14"), Operator:=xlFilterValues
    wsFormat.Range(filterRange).AutoFilter Field:=7, Criteria1:="<>"
    wsFormat.Range(filterRange).AutoFilter Field:=8, Criteria1:="<>"
    wsFormat.Range(filterRange).AutoFilter Field:=18, Criteria1:="<>"

    Call CopyVisibleColC(wsFormat, wsTemp, lastRow)

    If wsFormat.AutoFilterMode Then wsFormat.AutoFilterMode = False

    ' --- RUN FILTER PROFILE SEQUENCE 2 ---
    wsFormat.Range(filterRange).AutoFilter Field:=2, Criteria1:=Array( _
        "060", "066", "070", "080", "085", "100"), Operator:=xlFilterValues
    wsFormat.Range(filterRange).AutoFilter Field:=7, Criteria1:="<>"
    wsFormat.Range(filterRange).AutoFilter Field:=18, Criteria1:="<>"

    Call CopyVisibleColC(wsFormat, wsTemp, lastRow)

    If wsFormat.AutoFilterMode Then wsFormat.AutoFilterMode = False

    ' Remove processing duplicates
    Dim tempLastRow As Long
    tempLastRow = wsTemp.Cells(wsTemp.Rows.Count, "A").End(xlUp).Row
    If tempLastRow > 1 Then
        wsTemp.Range("A1:B" & tempLastRow).RemoveDuplicates Columns:=1, Header:=xlYes
    End If

    wsTemp.Columns("A:B").EntireColumn.AutoFit

End Sub

'==============================================================================
' STAGE 3 — Loop Filtered Order Records Through COR2 System Updates
'==============================================================================
Public Sub Stage3_COR2_VARUpdate()

    Dim SapGuiAuto  As Object
    Dim Application As Object
    Dim Connection  As Object
    Dim Session     As Object
    Dim wsTemp      As Worksheet
    Dim tempLastRow As Long

    If Not IsObject(Application) Then
       Set SapGuiAuto = GetObject("SAPGUI")
       Set Application = SapGuiAuto.GetScriptingEngine
    End If
    If Not IsObject(Connection) Then
       Set Connection = Application.Children(0)
    End If
    If Not IsObject(Session) Then
       Set Session = Connection.Children(0)
    End If

    On Error Resume Next
    Set wsTemp = ThisWorkbook.Sheets("Temp")
    On Error GoTo 0

    If wsTemp Is Nothing Then Exit Sub

    tempLastRow = wsTemp.Cells(wsTemp.Rows.Count, "A").End(xlUp).Row
    If tempLastRow < 2 Then
        MsgBox "No valid actionable rows staged in workspace queue.", vbInformation
        Exit Sub
    End If

    Call KillMessageFilter

    Dim i As Long, a As Integer
    Dim dd As String, orderNo As String, varFound As Boolean

    For i = 2 To tempLastRow
        orderNo = Trim(CStr(wsTemp.Cells(i, 1).Value))
        If orderNo = "" Then GoTo NextOrder

        On Error Resume Next
        Session.findById("wnd[0]").maximize
        If Err.Number <> 0 Then
            Set SapGuiAuto = GetObject("SAPGUI")
            Set Application = SapGuiAuto.GetScriptingEngine
            Set Connection = Application.Children(0)
            Set Session = Connection.Children(0)
            Session.findById("wnd[0]").maximize
            Err.Clear
        End If
        On Error GoTo 0

        Session.findById("wnd[0]/tbar[0]/okcd").Text = "/nCOR2"
        Session.findById("wnd[0]").sendVKey 0

        Session.findById("wnd[0]/usr/radR62CLORD-FLG_COMPL").Select
        Session.findById("wnd[0]/usr/ctxtCAUFVD-AUFNR").Text = orderNo
        Session.findById("wnd[0]/usr/radR62CLORD-FLG_COMPL").SetFocus
        Session.findById("wnd[0]").sendVKey 0

        If Session.ActiveWindow.Name = "wnd[1]" Then
            If InStr(Session.findById("wnd[1]").Text, "Information") > 0 Then
                Session.findById("wnd[1]/tbar[0]/btn[0]").press
                wsTemp.Cells(i, 2).Value = "Error - Info Popup"
                GoTo NextOrder
            End If
        End If

        Session.findById("wnd[0]/usr/btnPUSH_STATUS").press
        varFound = False

        For a = 0 To 9
            dd = ""
            On Error Resume Next
            dd = Session.findById("wnd[0]/usr/tabsTABSTRIP_0300/tabpANWS/ssubSUBSCREEN:SAPLBSVA:0302/tblSAPLBSVATC_E/txtJEST_BUF_E-ETX04[1," & a & "]").Text
            On Error GoTo 0

            If dd = "VAR" Then
                varFound = True
                Session.findById("wnd[0]/usr/tabsTABSTRIP_0300/tabpANWS/ssubSUBSCREEN:SAPLBSVA:0302/tblSAPLBSVATC_E/radJ_STMAINT-ANWS[0," & a & "]").Selected = True
                wsTemp.Cells(i, 2).Value = "Success"
                Exit For
            End If
        Next a

        If Not varFound Then wsTemp.Cells(i, 2).Value = "Error - VAR not found"

        If varFound Then
            Session.findById("wnd[0]/tbar[0]/btn[3]").press
            Session.findById("wnd[0]/tbar[0]/btn[11]").press
        Else
            Session.findById("wnd[0]/tbar[0]/btn[3]").press
        End If

NextOrder:
        Session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
        Session.findById("wnd[0]").sendVKey 0
    Next i

    ' Summary Metrics Display
    Dim successCount As Long, errorCount As Long, r As Long
    For r = 2 To tempLastRow
        If wsTemp.Cells(r, 2).Value = "Success" Then successCount = successCount + 1 Else errorCount = errorCount + 1
    Next r

    ' Delete the processed export file to keep C:\TEMP clean
    Dim sTargetFile As String
    sTargetFile = GetLatestSAPExport("C:\TEMP\")
    On Error Resume Next
    If sTargetFile <> "" Then Kill sTargetFile
    On Error GoTo 0

    MsgBox "Execution Pipeline Complete!" & vbLf & vbLf & _
           "Staged Count: " & (tempLastRow - 1) & vbLf & _
           "Successful  : " & successCount & vbLf & _
           "Exceptions  : " & errorCount, vbInformation
End Sub

'==============================================================================
' CORE HELPER UTILITIES
'==============================================================================
Private Sub CopyVisibleColC(wsFormat As Worksheet, wsTemp As Worksheet, lastRow As Long)
    Dim visibleCells As Range, dataRange As Range, cell As Range, pasteRow As Long
    If lastRow < 2 Then Exit Sub
    Set dataRange = wsFormat.Range("C2:C" & lastRow)
    On Error Resume Next
    Set visibleCells = dataRange.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    If visibleCells Is Nothing Then Exit Sub
    pasteRow = wsTemp.Cells(wsTemp.Rows.Count, "A").End(xlUp).Row + 1
    For Each cell In visibleCells
        If Trim(CStr(cell.Value)) <> "" Then
            wsTemp.Cells(pasteRow, 1).Value = cell.Value
            pasteRow = pasteRow + 1
        End If
    Next cell
End Sub

' Helper function to fetch the newest timestamped file generated by SAP
Private Function GetLatestSAPExport(FolderPath As String) As String
    Dim FileName As String, MostRecentFile As String
    Dim LatestDate As Date, FileDate As Date
    If Right(FolderPath, 1) <> "\" Then FolderPath = FolderPath & "\"
    
    FileName = Dir(FolderPath & "EXPORT_*.xlsx")
    Do While FileName <> ""
        FileDate = FileDateTime(FolderPath & FileName)
        If FileDate > LatestDate Then
            LatestDate = FileDate
            MostRecentFile = FolderPath & FileName
        End If
        FileName = Dir
    Loop
    GetLatestSAPExport = MostRecentFile
End Function

