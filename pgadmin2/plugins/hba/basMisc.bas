Attribute VB_Name = "basMisc"
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
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
Public Function GetID() As String
On Error GoTo Err_Handler
'Don't log, accessed *all* the time.

Static lID As Long
  lID = lID + 1
  GetID = lID
  
Exit Function
Err_Handler:
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basMisc.GetID"
End Function

Public Function SearchListview(lvListview As ListView, szSearchstring As String) As Boolean
On Error GoTo Err_Handler
Dim itmSearchFor As ListItem

Set itmSearchFor = lvListview.FindItem(szSearchstring)
If itmSearchFor Is Nothing Then
  SearchListview = False
Else
  SearchListview = True
End If

Exit Function
Err_Handler:
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basMisc.SearchListview"
End Function
