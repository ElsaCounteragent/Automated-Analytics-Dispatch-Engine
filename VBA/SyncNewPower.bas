Attribute VB_Name = "Module2"
Sub SyncNewPower()
    Dim wsAnalysis As Worksheet
    Dim wsDatabase As Worksheet
    Dim i As Long, lastRow As Long
    Dim playerName As String
    Dim newPower As Variant
    Dim foundCell As Range
    
    Set wsAnalysis = Sheets("Analysis1")
    Set wsDatabase = Sheets("Database")
    
    lastRow = wsAnalysis.Cells(wsAnalysis.Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To lastRow
        playerName = wsAnalysis.Cells(i, 1).Value
        newPower = wsAnalysis.Cells(i, 10).Value
        
        If newPower <> "" Then
            Set foundCell = wsDatabase.Range("A:A").Find(What:=playerName, LookAt:=xlWhole)
            
            If Not foundCell Is Nothing Then
                If wsDatabase.Cells(foundCell.Row, 4).Value <> newPower Then
                    wsDatabase.Cells(foundCell.Row, 4).Value = newPower
                End If
            End If
        End If
    Next i
    
    MsgBox "Database Power Ratings Synchronized.", vbInformation, "Command Console"
End Sub
