Attribute VB_Name = "basMisc"
' pgAdmin II Migration Wizard
' This code is based on code from the original pgAdmin Project
' Copyright (C) 1998 - 2002, Dave Page & others

' This program is free software; you can redistribute it and/or
' modify it under the terms of the GNU General Public License
' as published by the Free Software Foundation; either version 2
' of the License, or (at your option) any later version.

' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.

' You should have received a copy of the GNU General Public License
' along with this program; if not, write to the Free Software
' Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.

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

'Format an identifier as required
'This code is based on fmtID from the pg_dump code
Public Function fmtID(ByVal szData As String) As String
On Error Resume Next

Dim X As Integer
Dim iVal As Integer

  'Replace double quotes
  szData = Replace(szData, QUOTE, QUOTE & QUOTE)
    
  For X = 1 To Len(szData)
    iVal = Asc(Mid(szData, X, 1))
    If Not ((iVal >= 48) And (iVal <= 57)) And _
       Not ((iVal >= 97) And (iVal <= 122)) And _
       Not (iVal = 95) Then
      szData = QUOTE & szData & QUOTE
      Exit For
    End If
  Next X
  
  fmtID = szData

End Function
