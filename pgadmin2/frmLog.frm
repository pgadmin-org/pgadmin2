VERSION 5.00
Begin VB.Form frmLog 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Log View"
   ClientHeight    =   1800
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   9105
   ControlBox      =   0   'False
   Icon            =   "frmLog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   9105
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtLog 
      Height          =   1770
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      ToolTipText     =   "Displays the Progeny's rolling log."
      Top             =   0
      Width           =   9105
   End
End
Attribute VB_Name = "frmLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence
'
' frmLog.frm - Displays the rolling log

Option Explicit

Public Sub LogMsg(szMessage As String)
'Note - No function entry logging is done here 'cos we'd enter a loop then...

Dim X As Long
  szMessage = Replace(szMessage, vbCrLf, " ")
  szMessage = Replace(szMessage, vbLf, " ")
  If Len(txtLog.Text) + Len(Now & " - " & szMessage) > 32767 Then
    txtLog.Text = Mid(txtLog.Text, InStr(Len(szMessage), txtLog.Text, vbCrLf) + 2, Len(txtLog.Text))
  End If
  X = Len(txtLog.Text)
  If txtLog.Text = "" Then
    txtLog.Text = Now & " - " & szMessage
  Else
    txtLog.Text = txtLog.Text & vbCrLf & Now & " - " & szMessage
  End If
  txtLog.SelStart = X + 2

End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering frmLog.Form_Load()", etFullDebug

  'Size & position the form
  Me.Left = Val(RegRead(HKEY_CURRENT_USER, "Software\" & App.Title & "\Log Window", "Left", "0"))
  Me.Top = Val(RegRead(HKEY_CURRENT_USER, "Software\" & App.Title & "\Log Window", "Top", "0"))
  Me.Width = Val(RegRead(HKEY_CURRENT_USER, "Software\" & App.Title & "\Log Window", "Width", "9000"))
  Me.Height = Val(RegRead(HKEY_CURRENT_USER, "Software\" & App.Title & "\Log Window", "Height", "2200"))
  
  'Set the form topmost if required
  If UCase(RegRead(HKEY_CURRENT_USER, "Software\" & App.Title & "\Log Window", "Always On Top", "Y")) = "Y" Then
    SetTopMostWindow Me.hWnd, True
  Else
    SetTopMostWindow Me.hWnd, False
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, "frmLog.Form_Load"
End Sub

Private Sub Form_Resize()
On Error GoTo Err_Handler
'Don't log this as if the user is resizing this window, they probably don't want to see resize messages!

  txtLog.Width = Me.ScaleWidth
  txtLog.Height = Me.ScaleHeight
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, "frmLog.Form_Resize"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering frmLog.Form_Unload(" & Cancel & ")", etFullDebug

  'Stop writing Log Messages
  ctx.LogView = False
  
  'Save the size/position
  RegWrite HKEY_CURRENT_USER, "Software\" & App.Title & "\Log Window", "Top", regString, Me.Top
  RegWrite HKEY_CURRENT_USER, "Software\" & App.Title & "\Log Window", "Left", regString, Me.Left
  RegWrite HKEY_CURRENT_USER, "Software\" & App.Title & "\Log Window", "Width", regString, Me.Width
  RegWrite HKEY_CURRENT_USER, "Software\" & App.Title & "\Log Window", "Height", regString, Me.Height
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, "frmLog.Form_Unload"
End Sub
