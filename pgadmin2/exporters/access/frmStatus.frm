VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmStatus 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Status Exporting"
   ClientHeight    =   825
   ClientLeft      =   7530
   ClientTop       =   2175
   ClientWidth     =   3945
   ControlBox      =   0   'False
   Icon            =   "frmStatus.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   825
   ScaleWidth      =   3945
   Begin MSComctlLib.ProgressBar pb 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Min             =   1e-4
   End
   Begin VB.Label lblCount 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the Artistic Licence
