Attribute VB_Name = "Module4"
Sub SendPingsToDiscord()
    Dim WebhookURL As String
    Dim HTTP As Object
    Dim JSON As String
    Dim cell As Range
    Dim CommsRange As Range
    
    WebhookURL = "somehook"
    Set CommsRange = Sheets("CwlJuly").Range("M3:M22")
    
    Set HTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    For Each cell In CommsRange
        If cell.Value <> "" Then
            JSON = "{""content"": """ & cell.Value & """}"
            
            HTTP.Open "POST", WebhookURL, False
            HTTP.setRequestHeader "Content-Type", "application/json"
            HTTP.send JSON
            
            Application.Wait (Now + TimeValue("0:00:01"))
            DoEvents
        End If
    Next cell
    
    Application.OnTime Now + TimeValue("00:03:00"), "SendRolePing"
    
    MsgBox "Phase 1 Complete: Tactical orders successfully transmitted." & vbCrLf & vbCrLf & "Leave this Excel file open. The Verified Role will be pinged in exactly 3 minutes.", vbInformation, "Command Console"
End Sub

Sub SendRolePing()
    Dim SecondWebhookURL As String
    Dim RoleID As String
    Dim HTTP As Object
    Dim RolePingJSON As String
    
    SecondWebhookURL = "hookinghook"
    RoleID = "12312312323123"
    
    RolePingJSON = "{""content"": ""<@&" & RoleID & "> Assignments for the current CWL have just been posted in Ingame Chat!""}"
    
    Set HTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
    HTTP.Open "POST", SecondWebhookURL, False
    HTTP.setRequestHeader "Content-Type", "application/json"
    HTTP.send RolePingJSON
    
    Set HTTP = Nothing
End Sub
