VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmExport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MsAccess"
   ClientHeight    =   1860
   ClientLeft      =   2265
   ClientTop       =   2130
   ClientWidth     =   4710
   Icon            =   "frmExport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   4710
   Begin VB.ComboBox cboCond 
      Height          =   315
      ItemData        =   "frmExport.frx":12FA
      Left            =   1215
      List            =   "frmExport.frx":1307
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox txtTableName 
      Height          =   285
      Left            =   1200
      TabIndex        =   4
      Text            =   "Result"
      Top             =   480
      Width           =   3030
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "&Export"
      Height          =   465
      Left            =   1560
      TabIndex        =   3
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   330
      Left            =   4320
      TabIndex        =   2
      Top             =   120
      Width           =   330
   End
   Begin VB.TextBox txtFileName 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   3030
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblTable 
      AutoSize        =   -1  'True
      Caption         =   "If table exists"
      Height          =   195
      Left            =   90
      TabIndex        =   6
      Top             =   900
      Width           =   915
   End
   Begin VB.Label lblTableName 
      Caption         =   "Table Name"
      Height          =   195
      Left            =   60
      TabIndex        =   5
      Top             =   540
      Width           =   915
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblfileName 
      AutoSize        =   -1  'True
      Caption         =   "Export File"
      Height          =   195
      Left            =   60
      TabIndex        =   1
      Top             =   165
      Width           =   735
   End
End
Attribute VB_Name = "frmExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the Artistic Licence

Option Explicit

Private Sub cmdBrowse_Click()
On Error Resume Next
  With CommonDialog1
    .FileName = txtFileName.Text
    .DialogTitle = "Save MDB File"
    .Filter = "Access Files (*.mdb)|*.mdb"
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
    If MsgBox("File exists - continue with export?", vbYesNo + vbQuestion, "Export to existing file") = vbNo Then Exit Sub
  End If
  Me.Hide
End Sub

Private Sub Form_Load()
On Error Resume Next
  cboCond.ListIndex = 0
End Sub
