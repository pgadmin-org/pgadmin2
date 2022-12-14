VERSION 5.00
Begin VB.Form frmPassword 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Change Password"
   ClientHeight    =   1728
   ClientLeft      =   48
   ClientTop       =   252
   ClientWidth     =   3468
   Icon            =   "frmPassword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1728
   ScaleWidth      =   3468
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2385
      TabIndex        =   6
      Top             =   1305
      Width           =   1050
   End
   Begin VB.TextBox txtConfirm 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1575
      PasswordChar    =   "*"
      TabIndex        =   5
      ToolTipText     =   "Re-enter your new password."
      Top             =   945
      Width           =   1860
   End
   Begin VB.TextBox txtNew 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1575
      PasswordChar    =   "*"
      TabIndex        =   4
      ToolTipText     =   "Enter your new password."
      Top             =   540
      Width           =   1860
   End
   Begin VB.TextBox txtCurrent 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1575
      PasswordChar    =   "*"
      TabIndex        =   3
      ToolTipText     =   "Enter your current password."
      Top             =   90
      Width           =   1860
   End
   Begin VB.Label Label1 
      Caption         =   "Confirm Password"
      Height          =   240
      Index           =   2
      Left            =   45
      TabIndex        =   2
      Top             =   945
      Width           =   1545
   End
   Begin VB.Label Label1 
      Caption         =   "New Password"
      Height          =   240
      Index           =   1
      Left            =   45
      TabIndex        =   1
      Top             =   540
      Width           =   1545
   End
   Begin VB.Label Label1 
      Caption         =   "Current Password"
      Height          =   240
      Index           =   0
      Left            =   45
      TabIndex        =   0
      Top             =   135
      Width           =   1545
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the Artistic Licence
'
' frmPassword.frm - Change Password

Option Explicit

Private Sub cmdOK_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmPassword.cmdOK_Click()", etFullDebug

  If txtCurrent.Text <> ctx.Password Then
    MsgBox ??TrasLang??("Incorrect Password!"), vbExclamation, ??TrasLang??("Error")
    Exit Sub
  End If
  If txtNew.Text <> txtConfirm.Text Then
    MsgBox ??TrasLang??("Passwords do not match!"), vbExclamation, ??TrasLang??("Error")
    Exit Sub
  End If
  If InStr(1, txtNew.Text, " ") Or InStr(1, txtNew.Text, "'") Or InStr(1, txtNew.Text, QUOTE) Then
    MsgBox ??TrasLang??("Illegal characters in password!"), vbExclamation, ??TrasLang??("Error")
    Exit Sub
  End If
  
  frmMain.svr.Users(ctx.Username).Password = txtNew.Text
  MsgBox ??TrasLang??("Password successfully changed!"), vbInformation, ??TrasLang??("Success!!")
  
  Unload Me
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmPassword.cmdOK_Click"
End Sub

Private Sub Form_Load()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmPassword.Form_Load()", etFullDebug

  PatchForm Me
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmPassword.Form_Load"
End Sub
