VERSION 5.00
Begin VB.Form frmStatus 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Exporting to ASCII Text File..."
   ClientHeight    =   795
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3045
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   795
   ScaleWidth      =   3045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label lblCount 
      Height          =   195
      Left            =   1080
      TabIndex        =   1
      Top             =   270
      Width           =   1545
   End
   Begin VB.Label lblRecord 
      AutoSize        =   -1  'True
      Caption         =   "Record: "
      Height          =   195
      Left            =   405
      TabIndex        =   0
      Top             =   270
      Width           =   615
   End
End
Attribute VB_Name = "frmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence

Option Explicit
