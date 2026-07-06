Attribute VB_Name = "Module3"
Sub PublishToDiscord()
    Dim wsMaster As Worksheet
    Dim wsAnalysisMaster As Worksheet
    Dim wbPublic As Workbook
    Dim wsPublic As Worksheet
    Dim wsPublicAnalysis As Worksheet
    Dim PublicFilePath As String
    Dim lastRow As Long
    Dim i As Long
    Dim tbl As ListObject
    
    Set wsMaster = ThisWorkbook.Sheets("CwlJuly")
    Set wsAnalysisMaster = ThisWorkbook.Sheets("Analysis1")
    
    PublicFilePath = ThisWorkbook.Path & "\Public_War_Log.xlsx"
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    On Error Resume Next
    Set wbPublic = Workbooks.Open(PublicFilePath)
    On Error GoTo 0
    
    If wbPublic Is Nothing Then
        MsgBox "Failed! Could not find 'Public_War_Log.xlsx'", vbCritical
        Application.ScreenUpdating = True
        Exit Sub
    End If

    wsMaster.Copy Before:=wbPublic.Sheets(1)
    Set wsPublic = wbPublic.Sheets(1)
    
    lastRow = wsPublic.Cells.SpecialCells(xlCellTypeLastCell).Row
    For i = lastRow To 1 Step -1
        If wsPublic.Rows(i).Hidden Then
            wsPublic.Rows(i).Delete
        End If
    Next i
    
    For Each tbl In wsPublic.ListObjects
        tbl.Unlist
    Next tbl
    
    wsPublic.AutoFilterMode = False
    
    wsPublic.UsedRange.Value = wsPublic.UsedRange.Value
    
    wsPublic.Columns("D:G").Hidden = True
    wsPublic.Columns("L:XFD").Hidden = True
    
    wsAnalysisMaster.Copy After:=wbPublic.Sheets(1)
    Set wsPublicAnalysis = wbPublic.Sheets(2)
    
    wsPublicAnalysis.UsedRange.Value = wsPublicAnalysis.UsedRange.Value
    
    wsPublicAnalysis.Columns("I:XFD").Hidden = True
    
    Do While wbPublic.Sheets.Count > 2
        wbPublic.Sheets(3).Delete
    Loop
    
    wbPublic.Close SaveChanges:=True
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox "Complete!", vbInformation
End Sub
