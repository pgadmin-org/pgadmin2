VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{44F33AC4-8757-4330-B063-18608617F23E}#12.4#0"; "HighlightBox.ocx"
Begin VB.Form frmServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Server"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   5520
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3285
      TabIndex        =   10
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4410
      TabIndex        =   11
      Top             =   6480
      Width           =   1095
   End
   Begin TabDlg.SSTab tabProperties 
      Height          =   6360
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   5460
      _ExtentX        =   9631
      _ExtentY        =   11218
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Properties"
      TabPicture(0)   =   "frmServer.frx":014A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblProperties(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblProperties(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblProperties(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblProperties(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblProperties(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblProperties(5)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblProperties(6)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblProperties(7)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "hbxProperties(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtProperties(6)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtProperties(4)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtProperties(3)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtProperties(2)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtProperties(1)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtProperties(0)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtProperties(7)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtProperties(5)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).ControlCount=   17
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   5
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   6
         ToolTipText     =   "The ODBC driver version."
         Top             =   2700
         Width           =   3390
      End
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   510
         Index           =   7
         Left            =   1935
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   8
         ToolTipText     =   "The description of the PostgreSQL database."
         Top             =   3510
         Width           =   3390
      End
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   0
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   1
         ToolTipText     =   "The servers hostname or TCP/IP address."
         Top             =   675
         Width           =   3390
      End
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   1
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   2
         ToolTipText     =   "The TCP/IP port on the server."
         Top             =   1080
         Width           =   3390
      End
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   2
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   3
         ToolTipText     =   "The current username."
         Top             =   1485
         Width           =   3390
      End
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   3
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   4
         ToolTipText     =   "The last OID generated by initdb."
         Top             =   1890
         Width           =   3390
      End
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   4
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   5
         ToolTipText     =   "The ODBC driver name."
         Top             =   2295
         Width           =   3390
      End
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   6
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   7
         ToolTipText     =   "The version of PostgreSQL we're connected to."
         Top             =   3105
         Width           =   3390
      End
      Begin HighlightBox.HBX hbxProperties 
         Height          =   1860
         Index           =   0
         Left            =   135
         TabIndex        =   9
         ToolTipText     =   "The ODBC connection string used by the primary connection."
         Top             =   4140
         Width           =   5190
         _ExtentX        =   9155
         _ExtentY        =   3281
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
         Caption         =   "ODBC Connection String"
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "ODBC driver version"
         Height          =   195
         Index           =   7
         Left            =   135
         TabIndex        =   19
         Top             =   2745
         Width           =   1440
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "PostgreSQL description"
         Height          =   195
         Index           =   6
         Left            =   135
         TabIndex        =   18
         Top             =   3555
         Width           =   1665
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "PostgreSQL version"
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   17
         Top             =   3150
         Width           =   1410
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "ODBC driver name"
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   16
         Top             =   2340
         Width           =   1320
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Last system OID"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   15
         Top             =   1935
         Width           =   1155
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Username"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   14
         Top             =   1530
         Width           =   720
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Port"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   13
         Top             =   1125
         Width           =   285
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Hostname/IP Address"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   12
         Top             =   720
         Width           =   1560
      End
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001, 2002, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence
'
' frmServer.frm - Edit/Create a Server

Option Explicit

Dim objServer As pgServer

Private Sub cmdCancel_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmServer.cmdCancel_Click()", etFullDebug

  Unload Me
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmServer.cmdCancel_Click"
End Sub

Private Sub cmdOK_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmServer.cmdOK_Click()", etFullDebug

  Unload Me
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmServer.cmdOK_Click"
End Sub

Public Sub Initialise(obj As pgServer)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmServer.Initialise(" & QUOTE & obj.Identifier & QUOTE & ")", etFullDebug

Dim X As Integer

  Set objServer = obj
  
  'Set the font
  For X = 0 To 7
    Set txtProperties(X).Font = ctx.Font
  Next X
  Set hbxProperties(0).Font = ctx.Font
  
  Me.Caption = "Server: " & objServer.Identifier

  txtProperties(0).Text = objServer.Server
  txtProperties(1).Text = objServer.Port
  txtProperties(2).Text = objServer.Username
  txtProperties(3).Text = objServer.LastSystemOID
  txtProperties(4).Text = objServer.DriverName
  txtProperties(5).Text = objServer.DriverVersion.Major & "." & objServer.DriverVersion.Minor & "." & objServer.DriverVersion.Revision
  txtProperties(6).Text = objServer.dbVersion.Major & "." & objServer.dbVersion.Minor & "." & objServer.dbVersion.Revision
  txtProperties(7).Text = objServer.dbVersion.Description
  hbxProperties(0).Text = objServer.ConnectionString
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmServer.Initialise"
End Sub
