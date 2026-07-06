Attribute VB_Name = "Module5"
Sub MassAssignStrategy()
    Dim ws As Worksheet
    Dim currentDay As Integer
    Dim lastRow As Long
    Dim i As Long
    Dim userChoice As Integer
    Dim starValue As String
    Dim starEmoji As String
    
    starEmoji = ChrW(&H2B50)
    Set ws = ThisWorkbook.Sheets("CwlJuly")
    
    If IsNumeric(ws.Range("N2").Value) And ws.Range("N2").Value <> "" Then
        currentDay = ws.Range("N2").Value
    Else
        MsgBox "Error: Cell N2 is broken", vbCritical
        Exit Sub
    End If
    
    userChoice = MsgBox("Select the default strategy for Day " & currentDay & ":" & vbCrLf & vbCrLf & _
                        "[YES] = Assign 3 Stars" & vbCrLf & _
                        "[NO] = Assign 2 Stars", vbYesNoCancel + vbQuestion, "Tactical Command")
                        
    If userChoice = vbYes Then
        starValue = starEmoji & starEmoji & starEmoji
    ElseIf userChoice = vbNo Then
        starValue = starEmoji & starEmoji
    Else
        MsgBox "Assignment aborted.", vbInformation
        Exit Sub
    End If
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    
    For i = 2 To lastRow
        If ws.Range("A" & i).Value = currentDay And ws.Range("C" & i).Value <> "" Then
            ws.Range("I" & i).Value = starValue
        End If
    Next i
    
    MsgBox "Strategy assigned for all active players on Day " & currentDay & "!", vbInformation
End Sub

