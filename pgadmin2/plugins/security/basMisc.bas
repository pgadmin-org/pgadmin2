Attribute VB_Name = "basMisc"
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence

Option Explicit

Public Sub LogError(lError As Long, szError As String, szRoutine As String)
'No logging here, if anythings going wrong then we want the real error

  svr.LogEvent "Error in " & App.Title & ":" & szRoutine & ": " & lError & " - " & szError, etErrors
  MsgBox "An error has occured in " & App.Title & ":" & szRoutine & ":" & vbCrLf & vbCrLf & "Number: " & lError & vbCrLf & "Description: " & szError, vbExclamation, App.Title & " Error"
  
End Sub

Public Sub StartMsg(ByVal szMsg As String)
'Logging code, so no internal logging...

  svr.LogEvent szMsg, etMiniDebug
  Screen.MousePointer = vbHourglass
  sb.Panels("info").Text = szMsg
  sb.Refresh
  sTimer = Timer
  
End Sub

Public Sub EndMsg()
'Logging code, so no internal logging...

Dim szMsg As String

  szMsg = "Done - " & Fix((Timer - sTimer) * 100) / 100 & " Secs."
  If Right(sb.Panels("info").Text, 5) <> "Done." Then
    svr.LogEvent szMsg, etMiniDebug
    sb.Panels("timer").Text = Fix((Timer - sTimer) * 100) / 100 & " Secs."
    sb.Panels("info").Text = sb.Panels("info").Text & " Done."
    sb.Refresh
  End If
  Screen.MousePointer = vbDefault
  
End Sub

'Parse an ACL and return | delimited User/Access lists
Public Sub ParseACL(ByVal szACL As String, ByRef szUserList As String, ByRef szAccessList As String)
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":basMisc.ParseACL(" & QUOTE & szACL & QUOTE & ", " & QUOTE & szUserList & QUOTE & ", " & QUOTE & szAccessList & QUOTE & ")", etFullDebug

Dim szEntries() As String
Dim szEntry As Variant
Dim szName As String
Dim szAccess As String
Dim szSQL As String
Dim szTemp As String
  
  szUserList = ""
  szAccessList = ""
  If szACL = "" Then Exit Sub
  szACL = Mid(szACL, 2, Len(szACL) - 2)
  szACL = Replace(szACL, QUOTE, "")
  szEntries = Split(szACL, ",")
  For Each szEntry In szEntries
  
    'Get the username
    szName = Left(szEntry, InStr(1, szEntry, "=") - 1)
    If szName = "" Then
      szName = "PUBLIC"
    ElseIf Len(szName) > 6 Then
      If Left(UCase(szName), 6) = "GROUP " Then
        szName = "GROUP " & Mid(szName, 7)
      Else
        szName = szName
      End If
    Else
      szName = szName
    End If
    
    'Get the Access
    szAccess = Mid(szEntry, InStr(1, szEntry, "=") + 1)
    szTemp = ""
    
    'ACLs are different in 7.2+
    If svr.dbVersion.VersionNum < 7.2 Then
      
      Select Case szAccess
        Case "arwR"
          szAccess = "All"
        Case ""
          szAccess = "None"
        Case Else
          If InStr(1, szAccess, "a") <> 0 Then szTemp = szTemp & "Insert, "
          If InStr(1, szAccess, "r") <> 0 Then szTemp = szTemp & "Select, "
          If InStr(1, szAccess, "w") <> 0 Then szTemp = szTemp & "Update, Delete, "
          If InStr(1, szAccess, "R") <> 0 Then szTemp = szTemp & "Rule, "
          szAccess = Left(szTemp, Len(szTemp) - 2)
      End Select
    
    Else
      
      Select Case szAccess
        Case "arwdRxt"
          szAccess = "All"
        Case ""
          szAccess = "None"
        Case Else
          If InStr(1, szAccess, "a") <> 0 Then szTemp = szTemp & "Insert, "
          If InStr(1, szAccess, "r") <> 0 Then szTemp = szTemp & "Select, "
          If InStr(1, szAccess, "w") <> 0 Then szTemp = szTemp & "Update, "
          If InStr(1, szAccess, "d") <> 0 Then szTemp = szTemp & "Delete, "
          If InStr(1, szAccess, "R") <> 0 Then szTemp = szTemp & "Rule, "
          If InStr(1, szAccess, "x") <> 0 Then szTemp = szTemp & "References, "
          If InStr(1, szAccess, "t") <> 0 Then szTemp = szTemp & "Trigger, "
          szAccess = Left(szTemp, Len(szTemp) - 2)
      End Select
    
    End If

    If szName <> "All" And szAccess <> "None" Then 'Don't include REVOKE ALL
      szUserList = szUserList & szName & "|"
      szAccessList = szAccessList & szAccess & "|"
    End If
  Next szEntry
  
  szUserList = Left(szUserList, Len(szUserList) - 1)
  szAccessList = Left(szAccessList, Len(szAccessList) - 1)
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basMisc.ParseACL"
End Sub
