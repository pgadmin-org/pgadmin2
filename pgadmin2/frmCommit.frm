VERSION 5.00
Object = "{44F33AC4-8757-4330-B063-18608617F23E}#12.4#0"; "HighLightBox.ocx"
Begin VB.Form frmCommit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Commit changes"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmCommit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3555
      TabIndex        =   2
      Top             =   2790
      Width           =   1095
   End
   Begin HighlightBox.HBX hbxComments 
      Height          =   2715
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4650
      _ExtentX        =   8202
      _ExtentY        =   4789
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Revision Log Comments"
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2385
      TabIndex        =   0
      Top             =   2790
      Width           =   1095
   End
End
Attribute VB_Name = "frmCommit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001, 2002, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence
'
' frmCommit.frm - Commit object changes to RC

Option Explicit

Dim objCurrent As Object

Private Sub cmdCancel_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmCommit.cmdCancel_Click()", etFullDebug

  Unload Me
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmCommit.cmdCancel_Click"
End Sub

Private Sub cmdOK_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmCommit.cmdOK_Click()", etFullDebug

  StartMsg "Committing changes..."
  objCurrent.Commit rcUpdate, hbxComments.Text
  frmMain.tv_NodeClick frmMain.tv.SelectedItem
  EndMsg
  Unload Me
  Exit Sub
  
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmCommit.cmdOK_Click"
End Sub

Public Sub Initialise(objCurr As Object)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmCommit.Initialise(" & objCurr.Identifier & ")", etFullDebug

  Set objCurrent = objCurr
  Me.Caption = "Commit changes: " & objCurrent.Identifier & " (" & objCurrent.ObjectType & ")"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmCommit.Initialise"
End Sub
