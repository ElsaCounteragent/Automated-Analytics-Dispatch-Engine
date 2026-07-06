Attribute VB_Name = "Module1"
Sub AutoScanLiveWar()
    Dim HTTP As Object
    Dim JSON As String, ClanData As String, OpponentData As String
    Dim APIToken As String, ClanTag As String, URL As String
    Dim wsTracker As Worksheet, wsDatabase As Worksheet
    Dim PrepTime As String
    
    APIToken = "sometoken"
    ClanTag = "%23sometag"
    
    Set wsTracker = ThisWorkbook.Sheets("somesheet")
    Set wsDatabase = ThisWorkbook.Sheets("othersheet")
    URL = "https://api.clashofclans.com/v1/clans/" & ClanTag & "/currentwar"
    
    On Error GoTo SafeExit
    Set HTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
    HTTP.Open "GET", URL, False
    HTTP.setRequestHeader "Authorization", "Bearer " & APIToken
    HTTP.setRequestHeader "Accept", "application/json"
    HTTP.send
    JSON = HTTP.responseText
    
    If InStr(JSON, "reason") > 0 Or JSON = "" Then Exit Sub
    If InStr(JSON, """state"":""inWar""") = 0 And InStr(JSON, """state"":""warEnded""") = 0 Then Exit Sub
    
    PrepTime = Split(Split(JSON, """preparationStartTime"":""")(1), """,")(0)
    
    ClanData = Split(JSON, """opponent"":{")(0)
    OpponentData = Split(JSON, """opponent"":{")(1)
    
    Dim oppDict As Object
    Set oppDict = CreateObject("Scripting.Dictionary")
    
    Dim oppBlocks() As String
    Dim oppTag As String, oppMap As Integer
    Dim i As Integer
    
    oppBlocks = Split(OpponentData, "{""tag"":""")
    For i = 1 To UBound(oppBlocks)
        oppTag = Split(oppBlocks(i), """")(0)
        If InStr(oppBlocks(i), """mapPosition"":") > 0 Then
            oppMap = Val(Split(Split(oppBlocks(i), """mapPosition"":")(1), ",")(0))
            oppDict.Add oppTag, oppMap
        End If
    Next i
    
    Dim clanBlocks() As String
    Dim playerName As String, attackerMap As Integer
    Dim attacksBlock As String, attackChunks() As String
    Dim stars1 As Integer, stars2 As Integer
    Dim defTag1 As String, defTag2 As String
    Dim passedTrial As Boolean
    
    clanBlocks = Split(ClanData, "{""tag"":""")
    
    For i = 1 To UBound(clanBlocks)
        passedTrial = False
        
        If InStr(clanBlocks(i), """mapPosition"":") > 0 Then
            playerName = Split(Split(clanBlocks(i), """name"":""")(1), """")(0)
            attackerMap = Val(Split(Split(clanBlocks(i), """mapPosition"":")(1), ",")(0))
            
            If InStr(clanBlocks(i), """attacks"":[") > 0 Then
                attacksBlock = Split(clanBlocks(i), """attacks"":[")(1)
                attacksBlock = Split(attacksBlock, "]")(0)
                attackChunks = Split(attacksBlock, "{""attackerTag"":")
                
                If UBound(attackChunks) = 2 Then
                    stars1 = Val(Split(Split(attackChunks(1), """stars"":")(1), ",")(0))
                    defTag1 = Split(Split(attackChunks(1), """defenderTag"":""")(1), """")(0)
                    stars2 = Val(Split(Split(attackChunks(2), """stars"":")(1), ",")(0))
                    defTag2 = Split(Split(attackChunks(2), """defenderTag"":""")(1), """")(0)
                    
                    If stars1 = 3 And stars2 = 3 Then
                        If oppDict.exists(defTag1) And oppDict.exists(defTag2) Then
                            
                            If attackerMap = 1 Then
                                If oppDict(defTag1) = 1 Or oppDict(defTag2) = 1 Then
                                    passedTrial = True
                                End If
                            Else
                                If oppDict(defTag1) <= attackerMap And oppDict(defTag2) <= attackerMap Then
                                    passedTrial = True
                                End If
                            End If
                            
                        End If
                    End If
                End If
            End If
            
            If passedTrial = True Then
                Dim foundCell As Range, dbCell As Range
                Dim nextRow As Long
                
                Set foundCell = wsTracker.Range("A:A").Find(What:=playerName, LookAt:=xlWhole)
                
                If foundCell Is Nothing Then
                    nextRow = wsTracker.Cells(wsTracker.Rows.Count, "A").End(xlUp).Row + 1
                    wsTracker.Cells(nextRow, 1).Value = playerName
                    wsTracker.Cells(nextRow, 2).Value = 1
                    wsTracker.Cells(nextRow, 3).Value = PrepTime
                Else
                    If foundCell.Offset(0, 2).Value <> PrepTime Then
                        foundCell.Offset(0, 1).Value = foundCell.Offset(0, 1).Value + 1
                        foundCell.Offset(0, 2).Value = PrepTime
                        
                        If foundCell.Offset(0, 1).Value >= 3 Then
                            Set dbCell = wsDatabase.Range("A:A").Find(What:=playerName, LookAt:=xlWhole)
                            If dbCell Is Nothing Then
                                Dim dbNextRow As Long
                                dbNextRow = wsDatabase.Cells(wsDatabase.Rows.Count, "A").End(xlUp).Row + 1
                                wsDatabase.Cells(dbNextRow, 1).Value = playerName
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next i

SafeExit:
    Set HTTP = Nothing
    Set oppDict = Nothing
End Sub
