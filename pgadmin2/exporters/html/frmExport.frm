VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmExport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Basic HTML "
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmExport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdExport 
      Caption         =   "&Export"
      Height          =   465
      Left            =   1260
      TabIndex        =   3
      Top             =   990
      Width           =   2175
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   330
      Left            =   4095
      TabIndex        =   2
      Top             =   360
      Width           =   330
   End
   Begin VB.TextBox txtFileName 
      Height          =   285
      Left            =   1035
      TabIndex        =   0
      Top             =   360
      Width           =   3030
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   180
      Top             =   1035
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblfileName 
      AutoSize        =   -1  'True
      Caption         =   "Export File"
      Height          =   195
      Left            =   180
      TabIndex        =   1
      Top             =   405
      Width           =   735
   End
End
Attribute VB_Name = "frmExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001, 2002, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence

Option Explicit

Private Sub cmdBrowse_Click()
  With CommonDialog1
    .FileName = txtFileName.Text
    .DialogTitle = "Save HTML File"
    .Filter = "HTML Files (*.htm;*.html)|*.htm;*.html"
    .ShowSave
  End With
  txtFileName.Text = CommonDialog1.FileName
End Sub

Private Sub cmdExport_Click()
  If txtFileName.Text = "" Then
    MsgBox "You must specify a filename!", vbExclamation, "Error"
    Exit Sub
  End If
  If Dir(txtFileName.Text) <> "" Then
    If MsgBox("File exists - overwrite?", vbYesNo + vbQuestion, "Overwrite File") = vbNo Then Exit Sub
  End If
  Me.Hide
End Sub
