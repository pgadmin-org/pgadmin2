VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmWizard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database Migration Wizard"
   ClientHeight    =   4320
   ClientLeft      =   2325
   ClientTop       =   1455
   ClientWidth     =   6885
   Icon            =   "frmWizard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   6885
   Begin VB.PictureBox picStrip 
      BorderStyle     =   0  'None
      Height          =   4335
      Left            =   0
      Picture         =   "frmWizard.frx":0BC2
      ScaleHeight     =   289
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   33
      Top             =   0
      Width           =   465
   End
   Begin VB.CommandButton cmdTypeMap 
      Caption         =   "&Edit Type Map"
      Height          =   330
      Left            =   540
      TabIndex        =   18
      ToolTipText     =   "Edit the data Type Map."
      Top             =   3960
      Visible         =   0   'False
      Width           =   1230
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2205
      Top             =   3915
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "&Previous"
      Enabled         =   0   'False
      Height          =   330
      Left            =   3300
      TabIndex        =   17
      ToolTipText     =   "Move back a step."
      Top             =   3960
      Width           =   1140
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next"
      Height          =   330
      Left            =   4500
      TabIndex        =   16
      ToolTipText     =   "Proceed to the next step."
      Top             =   3960
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   330
      Left            =   5700
      TabIndex        =   32
      ToolTipText     =   "Accept the completed migration"
      Top             =   3960
      Visible         =   0   'False
      Width           =   1140
   End
   Begin TabDlg.SSTab tabWizard 
      Height          =   3840
      Left            =   540
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   90
      Width           =   6300
      _ExtentX        =   11113
      _ExtentY        =   6773
      _Version        =   393216
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   176
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   " "
      TabPicture(0)   =   "frmWizard.frx":1861
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(2)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraODBC"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraAccess"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "optType(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "chkNotNull"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "optType(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "chkIndexes"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "chkLCaseColumns"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "chkLCaseTables"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "chkLCaseIndexes"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "chkPrimaryKey"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   " "
      TabPicture(1)   =   "frmWizard.frx":187D
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lstDatabase"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   " "
      TabPicture(2)   =   "frmWizard.frx":1899
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label1(1)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "lstTables"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "cmdSelect(0)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "cmdDeselect(0)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      TabCaption(3)   =   " "
      TabPicture(3)   =   "frmWizard.frx":18B5
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label1(9)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "lstData"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "cmdSelect(1)"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "cmdDeselect(1)"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).ControlCount=   4
      TabCaption(4)   =   " "
      TabPicture(4)   =   "frmWizard.frx":18D1
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label1(10)"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Label1(8)"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "lstForeignKeys"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "cmdSelect(2)"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "cmdDeselect(2)"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).ControlCount=   5
      TabCaption(5)   =   " "
      TabPicture(5)   =   "frmWizard.frx":18ED
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "pbStatus"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "txtStatus"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).ControlCount=   2
      Begin VB.CommandButton cmdDeselect 
         Caption         =   "&Deselect All"
         Height          =   330
         Index           =   0
         Left            =   -74790
         TabIndex        =   21
         Top             =   975
         Width           =   1230
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "&Select All"
         Height          =   330
         Index           =   0
         Left            =   -74790
         TabIndex        =   20
         ToolTipText     =   "Select all tables"
         Top             =   555
         Width           =   1230
      End
      Begin VB.ListBox lstTables 
         Height          =   3435
         Left            =   -73410
         Style           =   1  'Checkbox
         TabIndex        =   22
         Top             =   315
         Width           =   4650
      End
      Begin VB.CommandButton cmdDeselect 
         Caption         =   "&Deselect All"
         Height          =   330
         Index           =   1
         Left            =   -74790
         TabIndex        =   24
         Top             =   975
         Width           =   1230
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "&Select All"
         Height          =   330
         Index           =   1
         Left            =   -74790
         TabIndex        =   23
         ToolTipText     =   "Select all tables"
         Top             =   555
         Width           =   1230
      End
      Begin VB.ListBox lstData 
         Height          =   3435
         Left            =   -73410
         Style           =   1  'Checkbox
         TabIndex        =   25
         Top             =   315
         Width           =   4650
      End
      Begin VB.CommandButton cmdDeselect 
         Caption         =   "&Deselect All"
         Height          =   330
         Index           =   2
         Left            =   -74790
         TabIndex        =   27
         Top             =   975
         Width           =   1230
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "&Select All"
         Height          =   330
         Index           =   2
         Left            =   -74820
         TabIndex        =   26
         ToolTipText     =   "Select all foreign keys"
         Top             =   585
         Width           =   1230
      End
      Begin VB.ListBox lstForeignKeys 
         Height          =   3435
         Left            =   -73410
         Style           =   1  'Checkbox
         TabIndex        =   28
         Top             =   315
         Width           =   4650
      End
      Begin VB.TextBox txtStatus 
         Height          =   3480
         Left            =   -74955
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   29
         ToolTipText     =   "Displays the status of the migration process"
         Top             =   135
         Width           =   6180
      End
      Begin VB.CheckBox chkPrimaryKey 
         Caption         =   "Create Primary Keys on Migrated Tables"
         Height          =   240
         Left            =   660
         TabIndex        =   12
         ToolTipText     =   "Select this to attempt to migrate Primary Keys from the source database."
         Top             =   2700
         Value           =   1  'Checked
         Width           =   4380
      End
      Begin VB.CheckBox chkLCaseIndexes 
         Caption         =   "Convert Index/Key Names to Lower Case"
         Height          =   240
         Left            =   660
         TabIndex        =   15
         ToolTipText     =   "Select this to convert index names to lower case."
         Top             =   3420
         Width           =   4380
      End
      Begin VB.CheckBox chkLCaseTables 
         Caption         =   "Convert Table Names to Lower Case"
         Height          =   240
         Left            =   660
         TabIndex        =   13
         ToolTipText     =   "Select this to convert table names to lower case."
         Top             =   2940
         Width           =   4380
      End
      Begin VB.CheckBox chkLCaseColumns 
         Caption         =   "Convert Column Names to Lower Case"
         Height          =   240
         Left            =   660
         TabIndex        =   14
         ToolTipText     =   "Select this to convert column names to lower case."
         Top             =   3180
         Width           =   4380
      End
      Begin VB.CheckBox chkIndexes 
         Caption         =   "Create Indexes on Migrated Tables"
         Height          =   240
         Left            =   660
         TabIndex        =   11
         ToolTipText     =   "Select this to attempt to migrate Indexes from the source database."
         Top             =   2460
         Value           =   1  'Checked
         Width           =   4380
      End
      Begin VB.OptionButton optType 
         Caption         =   "&ODBC"
         Height          =   240
         Index           =   1
         Left            =   3150
         TabIndex        =   2
         ToolTipText     =   "Migrate an ODBC Datasource"
         Top             =   285
         Width           =   1500
      End
      Begin VB.CheckBox chkNotNull 
         Caption         =   "Create columns as 'NOT NULL' where applicable"
         Height          =   240
         Left            =   675
         TabIndex        =   10
         ToolTipText     =   "Select this to attempt to migrate 'NOT NULL' rules from the source database."
         Top             =   2220
         Value           =   1  'Checked
         Width           =   4380
      End
      Begin VB.OptionButton optType 
         Caption         =   "&Access"
         Height          =   240
         Index           =   0
         Left            =   2070
         TabIndex        =   1
         ToolTipText     =   "Migrate an MS Access Database"
         Top             =   285
         Value           =   -1  'True
         Width           =   1500
      End
      Begin MSComctlLib.ProgressBar pbStatus 
         Height          =   195
         Left            =   -74955
         TabIndex        =   30
         Top             =   3600
         Width           =   6180
         _ExtentX        =   10901
         _ExtentY        =   344
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComctlLib.ListView lstDatabase 
         Height          =   3300
         Left            =   -74955
         TabIndex        =   19
         Top             =   450
         Width           =   6180
         _ExtentX        =   10901
         _ExtentY        =   5821
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Database"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Comment"
            Object.Width           =   7056
         EndProperty
      End
      Begin VB.Frame fraAccess 
         Caption         =   "Access Database"
         Height          =   1455
         Left            =   585
         TabIndex        =   34
         Top             =   600
         Width           =   4965
         Begin VB.CommandButton cmdBrowse 
            Caption         =   "..."
            Height          =   285
            Left            =   4500
            TabIndex        =   4
            ToolTipText     =   "Browse for the database to migrate"
            Top             =   315
            Width           =   330
         End
         Begin VB.TextBox txtFile 
            Height          =   285
            Left            =   1080
            TabIndex        =   3
            ToolTipText     =   "Enter the filename of the database to migrate."
            Top             =   315
            Width           =   3435
         End
         Begin VB.TextBox txtUID 
            Height          =   285
            Index           =   0
            Left            =   1080
            TabIndex        =   5
            ToolTipText     =   "Enter a username for this database if required."
            Top             =   675
            Width           =   3435
         End
         Begin VB.TextBox txtPWD 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Index           =   0
            Left            =   1080
            PasswordChar    =   "*"
            TabIndex        =   6
            ToolTipText     =   "Enter a password for this database if required."
            Top             =   1035
            Width           =   3435
         End
         Begin VB.Label Label1 
            Caption         =   ".mdb File"
            Height          =   195
            Index           =   0
            Left            =   135
            TabIndex        =   37
            Top             =   360
            Width           =   1365
         End
         Begin VB.Label Label1 
            Caption         =   "Username"
            Height          =   195
            Index           =   3
            Left            =   135
            TabIndex        =   36
            Top             =   720
            Width           =   1365
         End
         Begin VB.Label Label1 
            Caption         =   "Password"
            Height          =   195
            Index           =   4
            Left            =   135
            TabIndex        =   35
            Top             =   1080
            Width           =   1365
         End
      End
      Begin VB.Frame fraODBC 
         Caption         =   "ODBC Database"
         Height          =   1455
         Left            =   600
         TabIndex        =   39
         Top             =   600
         Visible         =   0   'False
         Width           =   4965
         Begin VB.TextBox txtPWD 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Index           =   1
            Left            =   1080
            PasswordChar    =   "*"
            TabIndex        =   9
            ToolTipText     =   "Enter a valid password for this datasource"
            Top             =   1035
            Width           =   3435
         End
         Begin VB.TextBox txtUID 
            Height          =   285
            Index           =   1
            Left            =   1080
            TabIndex        =   8
            ToolTipText     =   "Enter a valid username for this datasource"
            Top             =   675
            Width           =   3435
         End
         Begin VB.ComboBox cboDatasource 
            Height          =   315
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   7
            ToolTipText     =   "Select a datasource to migrate"
            Top             =   315
            Width           =   3705
         End
         Begin VB.Label Label1 
            Caption         =   "Password"
            Height          =   195
            Index           =   5
            Left            =   135
            TabIndex        =   42
            Top             =   1080
            Width           =   1365
         End
         Begin VB.Label Label1 
            Caption         =   "Username"
            Height          =   195
            Index           =   6
            Left            =   135
            TabIndex        =   41
            Top             =   720
            Width           =   1365
         End
         Begin VB.Label Label1 
            Caption         =   "Datasource"
            Height          =   195
            Index           =   7
            Left            =   135
            TabIndex        =   40
            Top             =   360
            Width           =   1365
         End
      End
      Begin VB.Label Label2 
         Caption         =   "Select the database to migrate into."
         Height          =   240
         Left            =   -74910
         TabIndex        =   47
         Top             =   225
         Width           =   3255
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tables to migrate:"
         Height          =   195
         Index           =   1
         Left            =   -74910
         TabIndex        =   46
         Top             =   315
         Width           =   1365
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Migrate data from:"
         Height          =   195
         Index           =   9
         Left            =   -74910
         TabIndex        =   45
         Top             =   315
         Width           =   1365
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Foreign Keys:"
         Height          =   195
         Index           =   8
         Left            =   -74910
         TabIndex        =   44
         Top             =   315
         Width           =   1365
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Note: There may be more Foreign Keys than are listed, these are just those eligible for Migration."
         Height          =   2100
         Index           =   10
         Left            =   -74790
         TabIndex        =   43
         Top             =   1395
         Width           =   1185
      End
      Begin VB.Label Label1 
         Caption         =   "Database Type"
         Height          =   195
         Index           =   2
         Left            =   720
         TabIndex        =   38
         Top             =   285
         Width           =   1365
      End
   End
   Begin VB.CommandButton cmdMigrate 
      Caption         =   "&Migrate db"
      Enabled         =   0   'False
      Height          =   330
      Left            =   5700
      TabIndex        =   31
      ToolTipText     =   "Start the database migration."
      Top             =   3960
      Width           =   1140
   End
End
Attribute VB_Name = "frmWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgAdmin II Migration Wizard
' This code is based on code from the original pgAdmin Project
' Copyright (C) 1998 - 2001, Dave Page & others

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
Dim cnLocal As New Connection
Dim catLocal As New Catalog
Dim bButtonPress As Boolean
Dim bProgramPress As Boolean
Dim szQuoteChar As String
Dim szDatabase As String

Public Sub Initialise()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.Initialise()", etFullDebug

  tabWizard.Tab = 0
  cmdPrevious.Enabled = False
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.Form_Load"
End Sub

Private Sub cmdBrowse_Click()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.cmdBrowse_Click()", etFullDebug

Dim X As Integer
  lstTables.Clear
  With CommonDialog1
    .Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
    .Filter = "Access Databases (*.mdb)|*.mdb"
    .ShowOpen
  End With
  If CommonDialog1.FileName = "" Then Exit Sub
  txtFile.Text = CommonDialog1.FileName
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.cmdBrowse_Click"
End Sub

Private Function dbConnect() As Integer
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.dbConnect()", etFullDebug

Dim tblTemp As Table
  If cnLocal.State <> adStateClosed Then cnLocal.Close
  If optType(0).Value = True Then
    If txtFile.Text = "" Then
      MsgBox "You must select a database to migrate!", vbExclamation, "Error"
      dbConnect = 1
      Exit Function
    End If
        
    StartMsg "Opening and Examining Source Database..."
    svr.LogEvent "Opening File: " & txtFile.Text, etMiniDebug
    cnLocal.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & txtFile.Text & ";Persist Security Info=False", txtUID(0).Text, txtPWD(0).Text
    szQuoteChar = "`"
  Else
    If cboDatasource.Text = "" Then
      MsgBox "You must select a database to migrate!", vbExclamation, "Error"
      dbConnect = 1
      Exit Function
    End If
    StartMsg "Opening and Examining Source Database..."
    svr.LogEvent "Opening DSN: " & cboDatasource.Text, etMiniDebug
    cnLocal.Open "DSN=" & cboDatasource.Text & ";UID=" & txtUID(1).Text & ";PWD=" & txtPWD(1).Text, txtUID(1).Text, txtPWD(1).Text
    szQuoteChar = GetQuoteChar("DSN=" & cboDatasource.Text & ";UID=" & txtUID(1).Text & ";PWD=" & txtPWD(1).Text)
  End If
  svr.LogEvent "Opened connection: " & cnLocal.ConnectionString, etMiniDebug
  svr.LogEvent "Provider: " & cnLocal.Provider & " v" & cnLocal.Version, etMiniDebug
  svr.LogEvent "Quote Character: '" & szQuoteChar & "'", etMiniDebug
  On Error Resume Next
  Set catLocal.ActiveConnection = cnLocal
  On Error GoTo Err_Handler
  lstTables.Clear
  For Each tblTemp In catLocal.Tables
    If tblTemp.Type = "TABLE" Or tblTemp.Type = "VIEW" Then lstTables.AddItem tblTemp.Name
  Next
  EndMsg
  dbConnect = 0
  
  Exit Function
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.dbConnect"
  dbConnect = 1
End Function

Private Sub cmdDeSelect_Click(Index As Integer)
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.cmdDeSelect_Click()", etFullDebug

Dim X As Integer
  
'1/15/2001 Rod Childers
'Rewrote to use case not Elseif

  Select Case Index
    Case 0 'Tables to migrate
      For X = 0 To lstTables.ListCount - 1
        lstTables.Selected(X) = False
      Next
    Case 1 'Data to migrate
      For X = 0 To lstData.ListCount - 1
        lstData.Selected(X) = False
      Next
    Case 2 'Foreign Keys
      For X = 0 To lstForeignKeys.ListCount - 1
        lstForeignKeys.Selected(X) = False
      Next
  End Select
    
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.cmdDeSelect_Click"
End Sub

Private Sub cmdMigrate_Click()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.cmdMigrate_Click()", etFullDebug

  bButtonPress = True
  cmdNext.Visible = False
  cmdPrevious.Visible = False
  cmdMigrate.Visible = False
  cmdTypeMap.Visible = False
  cmdOK.Visible = True
  tabWizard.Tab = 5
  Migrate_Data
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.cmdMigrate_Click"
End Sub

Private Sub cmdOK_Click()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.cmdOK_Click()", etFullDebug
  
  txtStatus.Text = ""
  bButtonPress = True
  cmdNext.Enabled = True
  cmdNext.Visible = True
  cmdPrevious.Visible = True
  cmdOK.Visible = False
  cmdMigrate.Visible = True
  cmdMigrate.Enabled = False
  tabWizard.Tab = 0
  
  bRunning = False
  Unload Me
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.cmdOK_Click"
End Sub

Private Sub cmdSelect_Click(Index As Integer)
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.cmdSelect_Click()", etFullDebug

Dim X As Integer
  
'1/15/2001 Rod Childers
'Rewrote to use case not Elseif

  Select Case Index
    Case 0 'Tables to migrate
      For X = 0 To lstTables.ListCount - 1
        lstTables.Selected(X) = True
      Next
    Case 1 'Data to migrate
      For X = 0 To lstData.ListCount - 1
        lstData.Selected(X) = True
      Next
    Case 2 'Foreign Keys
      For X = 0 To lstForeignKeys.ListCount - 1
        lstForeignKeys.Selected(X) = True
      Next
  End Select
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.cmdSelect_Click"
End Sub

Private Sub cmdTypeMap_Click()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.cmdTypeMap_Click()", etFullDebug

  Load frmTypeMap
  frmTypeMap.Show vbModal, Me
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.cmdTypeMap_Click"
End Sub

Private Sub Form_Unload(Cancel As Integer)

  bRunning = False

End Sub

Private Sub optType_Click(Index As Integer)
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.optType_Click()", etFullDebug

  If Index = 0 Then
    fraAccess.Visible = True
    fraODBC.Visible = False
    chkIndexes.Value = 1
    chkIndexes.Enabled = True
    
    chkPrimaryKey.Value = 1
    chkPrimaryKey.Enabled = True
      
  Else
    fraAccess.Visible = False
    fraODBC.Visible = True
    chkIndexes.Value = 0
    chkIndexes.Enabled = False
    
    chkPrimaryKey.Value = 0
    chkPrimaryKey.Enabled = False
        
    On Error Resume Next
    
    Dim i As Integer
    Dim sDSNItem As String * 1024
    Dim sDRVItem As String * 1024
    Dim sDSN As String
    Dim sDRV As String
    Dim iDSNLen As Integer
    Dim iDRVLen As Integer
    Dim lHenv As Long         'handle to the environment

    'get the DSNs
    If SQLAllocEnv(lHenv) <> -1 Then
        Do Until i <> SQL_SUCCESS
            sDSNItem = Space(1024)
            sDRVItem = Space(1024)
            i = SQLDataSources(lHenv, SQL_FD_FETCH_NEXT, sDSNItem, 1024, iDSNLen, sDRVItem, 1024, iDRVLen)
            sDSN = VBA.Left(sDSNItem, iDSNLen)
            sDRV = VBA.Left(sDRVItem, iDRVLen)
                
            If sDSN <> Space(iDSNLen) Then cboDatasource.AddItem sDSN
        Loop
    End If

    cboDatasource.ListIndex = 0
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.optType_Click"
End Sub

Private Sub cmdNext_Click()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.cmdNext_Click()", etFullDebug

  bButtonPress = True
  
  '1/16/2001 Rod Childers
  'Use case now, more tabs now
  Select Case tabWizard.Tab
    Case 0  'Database select tab
      If dbConnect <> 0 Then Exit Sub
      Call GetTargetDatabases
      tabWizard.Tab = 1
      cmdMigrate.Enabled = False
      cmdNext.Enabled = True
      cmdPrevious.Enabled = True
      cmdTypeMap.Visible = False

    Case 1  'Target Database
      If lstDatabase.SelectedItem Is Nothing Then
        MsgBox "You must select a target database!", vbExclamation, "Error"
        Exit Sub
      End If
      szDatabase = lstDatabase.SelectedItem.Text
      tabWizard.Tab = 2
      cmdMigrate.Enabled = True
      cmdNext.Enabled = True
      cmdPrevious.Enabled = True
      cmdTypeMap.Visible = True
      
    Case 2  'lstTables tab
      Call Load_Data  'Display selected tables
      tabWizard.Tab = 3
      cmdMigrate.Enabled = False
      cmdNext.Enabled = True
      cmdPrevious.Enabled = True
      cmdTypeMap.Visible = True
    
    Case 3  'lstData tab
      Call GetEligibleForeignKeys
      tabWizard.Tab = 4
      cmdMigrate.Enabled = False
      cmdNext.Enabled = True
      cmdPrevious.Enabled = True
      cmdTypeMap.Visible = True
      
    Case 4  'Foreign Keys tab
      tabWizard.Tab = 5
      cmdMigrate.Enabled = True
      cmdNext.Enabled = False
      cmdPrevious.Enabled = True
      cmdTypeMap.Visible = True
    
    Case 5  'txtStatus tab
  
  End Select
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.cmdNext_Click"
End Sub

Private Sub cmdPrevious_Click()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.cmdPrevious_Click()", etFullDebug

Dim X As Integer
  bButtonPress = True
  
  '1/16/2001 Rod Childers
  'Use case now, more tabs now
  Select Case tabWizard.Tab
    Case 0  'Database select tab

    Case 1  'Target Database
      tabWizard.Tab = 0
      cmdMigrate.Enabled = False
      cmdNext.Enabled = True
      cmdPrevious.Enabled = False
      cmdTypeMap.Visible = False
      
    Case 2  'lstTables tab
      lstTables.Clear
      tabWizard.Tab = 1
      cmdMigrate.Enabled = False
      cmdNext.Enabled = True
      cmdPrevious.Enabled = True
      cmdTypeMap.Visible = True
    
    Case 3  'lstData tab
      lstData.Clear
      tabWizard.Tab = 2
      cmdMigrate.Enabled = False
      cmdNext.Enabled = True
      cmdPrevious.Enabled = True
      cmdTypeMap.Visible = True
    
    Case 4  'Foreign Keys tab
      lstForeignKeys.Clear
      tabWizard.Tab = 3
      cmdMigrate.Enabled = False
      cmdNext.Enabled = True
      cmdPrevious.Enabled = True
      cmdTypeMap.Visible = True
      
    Case 5  'txtStatus tab
      tabWizard.Tab = 4
      cmdMigrate.Enabled = False
      cmdNext.Enabled = True
      cmdPrevious.Enabled = True
      cmdTypeMap.Visible = True
      
  End Select
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.cmdPrevious_Click"
End Sub

Private Sub Load_Data()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.Load_Data()", etFullDebug

lstData.Clear
Dim X As Integer
  For X = 0 To lstTables.ListCount - 1
    If lstTables.Selected(X) = True Then
      lstData.AddItem lstTables.List(X)
    End If
  Next
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.Load_Data"
End Sub

Private Sub Migrate_Data()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.Migrate_Data()", etFullDebug

Dim W As Integer
Dim X As Integer
Dim Y As Integer
Dim i As Integer
Dim j As Integer
Dim Z As Integer

Dim bPrimaryKeyAdded As Boolean
Dim bIsForeignKey As Boolean

Dim szRelatedCols As String
Dim szTmpFKName As String

Dim szQryStr As String
Dim szTemp1 As String
Dim szTemp2 As String
Dim Start As Single
Dim rsTemp As New Recordset
Dim loFlag As Boolean
Dim Tuples As Long
Dim Fields As String
Dim Values As String
Dim fNum As Integer

'   06/29/01 Matthew MacSuga (AutoIncrement Fix)
'   Check for existance of an auto increment field
Dim auto_increment_on As Integer
Dim auto_increment_field_name As String
Dim auto_increment_count As Long
Dim auto_increment_table As String
Dim auto_increment_query As String
Dim auto_increment_rs As New Recordset


  StartMsg "Migrating database..."
  pbStatus.Max = lstData.ListCount
  pbStatus.Value = 0
  Start = Timer
  szDatabase = lstDatabase.SelectedItem.Text
  svr.LogEvent "Migration from " & cnLocal.ConnectionString & " to " & szDatabase & " starting.", etMiniDebug
  
  If chkNotNull.Value = 1 Then svr.LogEvent "NOT NULL rules being honoured.", etMiniDebug
  If chkLCaseTables.Value = 1 Then svr.LogEvent "Table names being converted to lowercase.", etMiniDebug
  If chkLCaseColumns.Value = 1 Then svr.LogEvent "Column names being converted to lowercase.", etMiniDebug
  If chkLCaseIndexes.Value = 1 Then svr.LogEvent "Index names being converted to lowercase.", etMiniDebug

  'Begin a transaction.
  svr.Databases(szDatabase).Execute "BEGIN"
          
  For X = 0 To lstData.ListCount - 1
    svr.LogEvent "Creating table: " & lstData.List(X), etMiniDebug
    txtStatus.Text = txtStatus.Text & "Creating table: " & lstData.List(X) & vbCrLf
    txtStatus.SelStart = Len(txtStatus.Text)
    Me.Refresh
    
    'Create the table
    
    szTemp1 = ""  'Added 1/30/2001 Rod Childers Variables not being set to ""
    szTemp2 = ""
    
    loFlag = False
    If chkLCaseTables.Value = 0 Then
      szQryStr = "CREATE TABLE " & QUOTE & lstData.List(X) & QUOTE & " ( "
    Else
      szQryStr = "CREATE TABLE " & QUOTE & LCase(lstData.List(X)) & QUOTE & " ( "
    End If
    
    '   06/29/01 Matthew MacSuga (AutoIncrement Fix)
    '   Check for existance of an auto increment field
    auto_increment_on = 0
    auto_increment_count = 0
    auto_increment_field_name = ""
    If chkLCaseTables.Value = 0 Then
      auto_increment_table = lstData.List(X)
    Else
      auto_increment_table = LCase(lstData.List(X))
    End If
    auto_increment_query = ""
    
    '   Only do this if it's an access database
    If InStr(1, cnLocal.ConnectionString, "MSDASQL") = 0 Then
      For Y = 0 To catLocal.Tables(lstData.List(X)).Columns.Count - 1
        If catLocal.Tables(lstData.List(X)).Columns(Y).Type = adInteger Then
          If catLocal.Tables(lstData.List(X)).Columns(Y).Properties("AutoIncrement") = True Then
            auto_increment_on = 1
            
            If chkLCaseColumns.Value = 0 Then
              auto_increment_field_name = catLocal.Tables(lstData.List(X)).Columns(Y).Name
            Else
              auto_increment_field_name = LCase(catLocal.Tables(lstData.List(X)).Columns(Y).Name)
            End If
            
            Exit For
          End If
        End If
      Next Y
    
      If auto_increment_on = 1 Then
        auto_increment_query = "SELECT MAX(" & szQuoteChar & auto_increment_field_name & szQuoteChar & ") AS RECCOUNT FROM " & szQuoteChar & auto_increment_table & szQuoteChar
            
        '   Perform the query
        auto_increment_rs.Open auto_increment_query, cnLocal, 3, 1
        If auto_increment_rs.RecordCount = 1 Then
          '   Set auto_increment_count = MAX(fieldname) + 1 (to start at next record)
          auto_increment_count = auto_increment_rs("RECCOUNT") + 1
        End If
            
        '   Destroy what I created
        If auto_increment_rs.State <> adStateClosed Then auto_increment_rs.Close
        Set auto_increment_rs = Nothing
        
        '   Set the PostgreSQL query
        auto_increment_query = "CREATE SEQUENCE " & QUOTE & auto_increment_table & "_" & auto_increment_field_name & "_key" & QUOTE & " START " & auto_increment_count
      Else
        auto_increment_query = ""
      End If
    End If
    '   End AutoIncrement Fix
            
    '   07/02/01 - Matthew MacSuga - Put columns in original order fix
    Dim rsTemp_Column As New Recordset
    Dim K As Integer
    Dim sqlQ As String
    Dim newColumnArray()
    Dim ColCount As Integer
    
    '   Generate query to get column names
    sqlQ = "SELECT * FROM " & szQuoteChar & lstData.List(X) & szQuoteChar & " WHERE 1=2"
    rsTemp_Column.Open sqlQ, cnLocal, 3, 1
    
    ColCount = rsTemp_Column.Fields.Count - 1
    '   Set up the temporary copy array - sloppy sort
    ReDim newColumnArray(ColCount)
    For K = 0 To ColCount
        For Y = 0 To catLocal.Tables(lstData.List(X)).Columns.Count - 1
            If catLocal.Tables(lstData.List(X)).Columns(Y).Name = rsTemp_Column.Fields(K).Name Then
                newColumnArray(K) = catLocal.Tables(lstData.List(X)).Columns(Y).Name
                
                Exit For
            End If
        Next
    Next
    
    If rsTemp_Column.State <> adStateClosed Then rsTemp_Column.Close
    Set rsTemp_Column = Nothing
    
    For Y = 0 To catLocal.Tables(lstData.List(X)).Columns.Count - 1
      'DJP 2001-07-02 Don't migrate the oid column on PostgreSQL Databases!
      If Not ((cnLocal.Properties("DBMS Name") = "PostgreSQL") And (newColumnArray(Y) = "oid")) Then
        If chkLCaseColumns.Value = 0 Then
          szTemp1 = szTemp1 & QUOTE & catLocal.Tables(lstData.List(X)).Columns(newColumnArray(Y)).Name & QUOTE
        Else
          szTemp1 = szTemp1 & QUOTE & LCase(catLocal.Tables(lstData.List(X)).Columns(newColumnArray(Y)).Name) & QUOTE
        End If
        Select Case catLocal.Tables(lstData.List(X)).Columns(newColumnArray(Y)).Type
          Case adBigInt
            szTemp2 = RegRead(HKEY_CURRENT_USER, "Software\pgAdmin\Type Map", "BigInt", "int8")
          Case adBinary
            szTemp2 = RegRead(HKEY_CURRENT_USER, "Software\pgAdmin\Type Map", "Binary", "text")
          Case adBoolean
            szTemp2 = RegRead(HKEY_CURRENT_USER, "Software\pgAdmin\Type Map", "Boolean", "bool")
          Case adBSTR
            szTemp2 = RegRead(HKEY_CURRENT_USER, "Software\pgAdmin\Type Map", "BSTR", "bytea")
          Case adChapter
            szTemp2 = RegRead(HKEY_CURRENT_USER, "Software\pgAdmin\Type Map", "Chapter", "int4")
          Case adChar
            szTemp2 = RegRead(HKEY_CURRENT_USER, "Software\pgAdmin\Type Map", "Char", "char")
          Case adCurrency
            szTemp2 = RegRead(HKEY_CURRENT_USER, "Software\pgAdmin\Type Map", "Currency", "money")
          Case adDate
            szTemp2 = RegRead(HKEY_CURRENT_USER, "Software\pgAdmin\Type Map", "Date", "date")
          Case adDBDate
            szTemp2 = RegRead(HKEY_CURRENT_USER, "Software\pgAdmin\Type Map", "DBDate", "date")
          Case adDBTime
            szTemp2 = RegRead(HKEY_CURRENT_USER, "Software\pgAdmin\Type Map", "DBTime", "time")
          Case adDBTimeStamp
            szTemp2 = RegRead(HKEY_CURRENT_USER, "Software\pgAdmin\Type Map", "DBTimestamp", "timestamp")
          Case adDecimal
            szTemp2 = RegRead(HKEY_CURRENT_USER, "Software\pgAdmin\Type Map", "Decimal", "numeric")
          Case adDouble
            szTemp2 = RegRead(HKEY_CURRENT_USER, "Software\pgAdmin\Type Map", "Double", "float8")
          Case adEmpty
            szTemp2 = RegRead(HKEY_CURRENT_USER, "Software\pgAdmin\Type Map", "Empty", "text")
          Case adError
            szTemp2 = RegRead(HKEY_CURRENT_USER, "Software\pgAdmin\Type Map", "Error", "int4")
          Case adFileTime
            szTemp2 = RegRead(HKEY_CURRENT_USER, "Software\pgAdmin\Type Map", "FileTime", "datetime")
          Case adGUID
            szTemp2 = RegRead(HKEY_CURRENT_USER, "Software\pgAdmin\Type Map", "GUID", "text")
          Case adInteger
            szTemp2 = RegRead(HKEY_CURRENT_USER, "Software\pgAdmin\Type Map", "Integer", "int4")
          Case adLongVarBinary
            szTemp2 = RegRead(HKEY_CURRENT_USER, "Software\pgAdmin\Type Map", "LongVarBinary", "lo")
            loFlag = True
          Case adLongVarChar
            szTemp2 = RegRead(HKEY_CURRENT_USER, "Software\pgAdmin\Type Map", "LongVarChar", "text")
          Case adLongVarWChar
            szTemp2 = RegRead(HKEY_CURRENT_USER, "Software\pgAdmin\Type Map", "LongVarWChar", "text")
          Case adPropVariant
            szTemp2 = RegRead(HKEY_CURRENT_USER, "Software\pgAdmin\Type Map", "PropVariant", "text")
          Case adSingle
            szTemp2 = RegRead(HKEY_CURRENT_USER, "Software\pgAdmin\Type Map", "Single", "float4")
          Case adSmallInt
            szTemp2 = RegRead(HKEY_CURRENT_USER, "Software\pgAdmin\Type Map", "SmallInt", "int2")
          Case adTinyInt
            szTemp2 = RegRead(HKEY_CURRENT_USER, "Software\pgAdmin\Type Map", "TinyInt", "int2")
          Case adUnsignedBigInt
            szTemp2 = RegRead(HKEY_CURRENT_USER, "Software\pgAdmin\Type Map", "UnsignedBigInt", "int8")
          Case adUnsignedInt
            szTemp2 = RegRead(HKEY_CURRENT_USER, "Software\pgAdmin\Type Map", "UnsignedInt", "int4")
          Case adUnsignedSmallInt
            szTemp2 = RegRead(HKEY_CURRENT_USER, "Software\pgAdmin\Type Map", "UnsignedSmallInt", "int2")
          Case adUnsignedTinyInt
            szTemp2 = RegRead(HKEY_CURRENT_USER, "Software\pgAdmin\Type Map", "UnsignedTinyInt", "int2")
          Case adUserDefined
            szTemp2 = RegRead(HKEY_CURRENT_USER, "Software\pgAdmin\Type Map", "UserDefined", "text")
          Case adVarBinary
            szTemp2 = RegRead(HKEY_CURRENT_USER, "Software\pgAdmin\Type Map", "VarBinary", "lo")
            loFlag = True
          Case adVarChar
            '1/16/2001 Rod Childers
            'Changed VarChar to default to VarChar
            'Text in Access is = VarChar in PostgreSQL
            'Memo in Access is = text in PostgreSQL
            szTemp2 = RegRead(HKEY_CURRENT_USER, "Software\pgAdmin\Type Map", "VarChar", "varchar")
          Case adVarWChar
            '1/16/2001 Rod Childers
            'Changed VarWChar to default to VarChar
            'Text in Access is = VarChar in PostgreSQL
            'Memo in Access is = text in PostgreSQL
            szTemp2 = RegRead(HKEY_CURRENT_USER, "Software\pgAdmin\Type Map", "VarWChar", "varchar")
          Case adWChar
            szTemp2 = RegRead(HKEY_CURRENT_USER, "Software\pgAdmin\Type Map", "WChar", "text")
          Case Else
          szTemp2 = "text"
        End Select
        If szTemp2 = "bpchar" Or szTemp2 = "char" Or szTemp2 = "varchar" Then
          If catLocal.Tables(lstData.List(X)).Columns(newColumnArray(Y)).DefinedSize = 0 Then
            szTemp2 = szTemp2 & "(1)"
          Else
            'Varchar cannot exceed 8088 chars!
            If catLocal.Tables(lstData.List(X)).Columns(Y).DefinedSize > 8088 Then
              txtStatus.Text = txtStatus.Text & "  The 'varchar' field " & catLocal.Tables(lstData.List(X)).Columns(newColumnArray(Y)).Name & " is too long and has been converted to type 'text'" & vbCrLf
              txtStatus.SelStart = Len(txtStatus.Text)
              svr.LogEvent "The 'varchar' field " & catLocal.Tables(lstData.List(X)).Columns(newColumnArray(Y)).Name & " is too long and has been converted to type 'text'", etMiniDebug
              szTemp2 = "text"
            Else
              szTemp2 = szTemp2 & "(" & catLocal.Tables(lstData.List(X)).Columns(newColumnArray(Y)).DefinedSize & ")"
            End If
          End If
        End If
        If szTemp2 = "numeric" Then
          szTemp2 = szTemp2 & "(" & catLocal.Tables(lstData.List(X)).Columns(newColumnArray(Y)).NumericScale & "," & catLocal.Tables(lstData.List(X)).Columns(newColumnArray(Y)).Precision & ")"
        End If
      
        ' Matthew MacSuga Auto Increment Fix
        If auto_increment_on = 1 Then
          If LCase(newColumnArray(Y)) = LCase(auto_increment_field_name) Then
            szTemp2 = "int4 DEFAULT nextval('" & auto_increment_table & "_" & auto_increment_field_name & "_key')"
          End If
        End If
        
        szTemp1 = szTemp1 & " " & szTemp2
        If chkNotNull.Value = 1 Then
          If catLocal.Tables(lstData.List(X)).Columns(newColumnArray(Y)).Attributes And adColNullable = False Then szTemp1 = szTemp1 & " NOT NULL"
        End If
        szTemp1 = szTemp1 & ", "
      End If
    Next Y
    
    If Len(szTemp1) > 2 Then
      
      '1/14/2001 Rod Childers
      'See if the user wants PrimaryKeys created
      bPrimaryKeyAdded = False
      If chkPrimaryKey.Value = 1 Then
        
        'loop through indexes for table, look for Primary Key
        For j = 0 To catLocal.Tables(lstData.List(X)).Indexes.Count - 1
          If catLocal.Tables(lstData.List(X)).Indexes(j).PrimaryKey = True Then
            'Primary Key found, set flag
            bPrimaryKeyAdded = True
            
            'Primary key will be added, keep the extra , at the end of field list
            'and add it to the query string
            szQryStr = szQryStr & szTemp1
            
            szQryStr = szQryStr & " PRIMARY KEY("
            
            'Get the field names of the fields in the primary key
            For i = 0 To catLocal.Tables(lstData.List(X)).Indexes(j).Columns.Count - 1
              If chkLCaseColumns.Value = 0 Then
                szQryStr = szQryStr & QUOTE & catLocal.Tables(lstData.List(X)).Indexes(j).Columns(i).Name & QUOTE & ", "
              Else
                szQryStr = szQryStr & QUOTE & LCase(catLocal.Tables(lstData.List(X)).Indexes(j).Columns(i).Name) & QUOTE & ", "
              End If
            Next i
          End If
        Next j
      End If
       
      If bPrimaryKeyAdded = True Then
        'Trim off the extra , at the end
        szQryStr = Left(szQryStr, (Len(szQryStr) - 2))
        'add a ) to close the field statment fo the PRIMARY KEY
        szQryStr = szQryStr & ")"
      Else
        'No Primary key will be added, trim off the extra , at the end of the fields
        szTemp1 = Mid(szTemp1, 1, Len(szTemp1) - 2)
        szQryStr = szQryStr & szTemp1
      End If
       
      szQryStr = szQryStr & " )"
      
      ' Matthew MacSuga Auto-Increment Fix
      If auto_increment_on = 1 Then svr.Databases(szDatabase).Execute auto_increment_query
      ' End Auto-Increment Fix
      
      svr.Databases(szDatabase).Execute szQryStr
      
      '
      'Copy the data if required
      '
      If lstData.Selected(X) = True Then
      
        'Warn that BLOBS are being ignored.
        If loFlag = True Then
          txtStatus.Text = txtStatus.Text & "  BLOB data was found and NOT copied." & vbCrLf
          txtStatus.SelStart = Len(txtStatus.Text)
          svr.LogEvent "BLOB data was found and NOT copied.", etMiniDebug
          Me.Refresh
        End If
        Tuples = 0
        txtStatus.Text = txtStatus.Text & "  Copying data..." & vbCrLf
        txtStatus.SelStart = Len(txtStatus.Text)
        Me.Refresh
        svr.LogEvent "Migrating Data from: " & lstData.List(X), etMiniDebug
        svr.LogEvent "Executing: SELECT * FROM " & szQuoteChar & lstData.List(X) & szQuoteChar, etMiniDebug
        rsTemp.Open "SELECT * FROM " & szQuoteChar & lstData.List(X) & szQuoteChar, cnLocal, adOpenForwardOnly
        While Not rsTemp.EOF
          If chkLCaseTables.Value = 0 Then
            szQryStr = "INSERT INTO " & QUOTE & lstData.List(X) & QUOTE
          Else
            szQryStr = "INSERT INTO " & QUOTE & LCase(lstData.List(X)) & QUOTE
          End If
        
          For Z = 0 To rsTemp.Fields.Count - 1
            If rsTemp.Fields(Z).Value & "" <> "" Then
                            
              If chkLCaseColumns.Value = 0 Then
                Fields = Fields & QUOTE & rsTemp.Fields(Z).Name & QUOTE & ", "
              Else
                Fields = Fields & QUOTE & LCase(rsTemp.Fields(Z).Name) & QUOTE & ", "
              End If
            
              Select Case rsTemp.Fields(Z).Type
                 ' 04/24/2001 Jean-Michel POURE
                 ' Useful tricks to avoid bugs in non-English systems :
                 ' replace comma with dots in numerical values
                 ' and get rid of money acronyms (like FF for example)
                  Case adCurrency, adDouble, adSingle, adDecimal
                      Values = Values & "'" & Str(Val(Replace(rsTemp.Fields(Z).Value, ",", "."))) & "', "
                 
                 ' Another useful trick to avoid bugs in non-English systems :
                 ' Convert 'True' or 'Vrai' or 'T' into -1
                 ' and 'False' or 'Faux' or 'F' into 0
                 ' In PostgreSQL driver uncheck Bool as Char
                  Case adBoolean
                      Dim tempValue As String
                      tempValue = rsTemp.Fields(Z).Value
                      If (tempValue = "F") Then tempValue = "False"
                      If (tempValue = "T") Then tempValue = "True"
                      Values = Values & "'" & CBool(tempValue) * "-1" & "', "
                  '1/20/2001 Rod Childers
                  'See if this a date field that only contains a Time, if so add Old date to it
                  'so postgress will accept it into a timestamp field
                   Case adDate, adDBDate, adDBTimeStamp
                       If Len(rsTemp.Fields(Z).Value) < 12 And Right(rsTemp.Fields(Z).Value, 1) = "M" Then
                          'Only contains the time
                          Values = Values & "'1899-12-30 " & Replace(Replace((rsTemp.Fields(Z).Value & ""), "\", "\\"), "'", "''") & "', "
                       Else
                          'Valid date,treat like any other field
                          Values = Values & "'" & Replace(Replace((rsTemp.Fields(Z).Value & ""), "\", "\\"), "'", "''") & "', "
                       End If
                     
                    ' Text values and others
                    Case Else
                    Values = Values & "'" & Replace(Replace((rsTemp.Fields(Z).Value & ""), "\", "\\"), "'", "''") & "', "
               End Select
             End If
          Next
        
          Fields = Mid(Fields, 1, Len(Fields) - 2)
          Values = Mid(Values, 1, Len(Values) - 2)
          
          szQryStr = szQryStr & " (" & Fields & ") VALUES (" & Values & ")"
        
          svr.Databases(szDatabase).Execute szQryStr
          Tuples = Tuples + 1
          Fields = ""
          Values = ""
          DoEvents
          rsTemp.MoveNext
        Wend
        If rsTemp.State <> adStateClosed Then rsTemp.Close
        txtStatus.Text = txtStatus.Text & "  Records Copied: " & Tuples & vbCrLf
        svr.LogEvent "Records Copied: " & Tuples, etMiniDebug
        txtStatus.SelStart = Len(txtStatus.Text)
        Me.Refresh
      End If
      
      '
      'Copy indexes if required
      '
      If chkIndexes.Value = 1 Then
             
        For Y = 0 To catLocal.Tables(lstData.List(X)).Indexes.Count - 1
        
          '1/14/2001 Rod Childers
          'If primary keys were created above, check each index
          'if it is a primary key do not recreate the index
          If chkPrimaryKey.Value = 1 And catLocal.Tables(lstData.List(X)).Indexes(Y).PrimaryKey = True Then
            '------Do nothing, skip this index, it was created above
          Else
                          
          '1/14/2001 Rod Childers
          'Keep ForeignKeys from being migrated as an index
          'loop throught all the Keys, if this index is a forigen key, don't create
          bIsForeignKey = False
          For i = 0 To catLocal.Tables(lstData.List(X)).Keys.Count - 1
            If catLocal.Tables(lstData.List(X)).Keys(i).Name = catLocal.Tables(lstData.List(X)).Indexes(Y) And catLocal.Tables(lstData.List(X)).Keys(i).Type = adKeyForeign Then
              'This is not an index, it is a ForeignKey, set flag
              bIsForeignKey = True
            End If
          Next i
            
              
          If bIsForeignKey = False Then
            txtStatus.Text = txtStatus.Text & "Creating index: " & catLocal.Tables(lstData.List(X)).Indexes(Y).Name & vbCrLf
            txtStatus.SelStart = Len(txtStatus.Text)
            Me.Refresh
            svr.LogEvent "Creating index: " & catLocal.Tables(lstData.List(X)).Indexes(Y).Name, etMiniDebug
            szQryStr = "CREATE "
              
            If catLocal.Tables(lstData.List(X)).Indexes(Y).Unique = True Then
              szQryStr = szQryStr & "UNIQUE "
            End If
                
            If Len(lstData.List(X) & "_" & catLocal.Tables(lstData.List(X)).Indexes(Y).Name) > 27 Then
              If chkLCaseIndexes.Value = 0 Then
                szQryStr = szQryStr & "INDEX " & QUOTE & Mid(lstData.List(X) & "_" & catLocal.Tables(lstData.List(X)).Indexes(Y).Name & "_idx", 1, 26) & "-" & Y & QUOTE
              Else
                szQryStr = szQryStr & "INDEX " & QUOTE & LCase(Mid(lstData.List(X) & "_" & catLocal.Tables(lstData.List(X)).Indexes(Y).Name & "_idx", 1, 26) & "-" & Y) & QUOTE
              End If
            Else
              If chkLCaseIndexes.Value = 0 Then
                szQryStr = szQryStr & "INDEX " & QUOTE & lstData.List(X) & "_" & catLocal.Tables(lstData.List(X)).Indexes(Y).Name & "_idx" & QUOTE
              Else
                szQryStr = szQryStr & "INDEX " & QUOTE & LCase(lstData.List(X) & "_" & catLocal.Tables(lstData.List(X)).Indexes(Y).Name & "_idx") & QUOTE
              End If
            End If
            If chkLCaseTables.Value = 0 Then
              szQryStr = szQryStr & " ON " & QUOTE & lstData.List(X) & QUOTE & " USING btree ("
            Else
              szQryStr = szQryStr & " ON " & QUOTE & LCase(lstData.List(X)) & QUOTE & " USING btree ("
            End If
            For W = 0 To catLocal.Tables(lstData.List(X)).Indexes(Y).Columns.Count - 1
              If chkLCaseColumns.Value = 0 Then
                szQryStr = szQryStr & QUOTE & catLocal.Tables(lstData.List(X)).Indexes(Y).Columns(W).Name & QUOTE & ", "
              Else
                szQryStr = szQryStr & QUOTE & LCase(catLocal.Tables(lstData.List(X)).Indexes(Y).Columns(W).Name) & QUOTE & ", "
              End If
            Next
            szQryStr = Mid(szQryStr, 1, Len(szQryStr) - 2) & ")"
            svr.Databases(szDatabase).Execute szQryStr
          End If
        End If
        
        Next
        szTemp1 = ""
        szQryStr = ""
        pbStatus.Value = pbStatus.Value + 1
        Me.Refresh
      End If
    
    Else
      txtStatus.Text = txtStatus.Text & "  " & "Table skipped - no columns found!" & vbCrLf
      svr.LogEvent "Table skipped - no columns found!", etMiniDebug
    End If
  Next
    
  
  '1/16/2001 Rod Childers
  'Migrate Eligible selected Foreign Keys
  'Making All Foreign keys Lower case
  For j = 0 To lstForeignKeys.ListCount - 1
    If lstForeignKeys.Selected(j) = True Then
      
      txtStatus.Text = txtStatus.Text & "Creating Foreign Key: " & lstForeignKeys.List(j) & vbCrLf
      txtStatus.SelStart = Len(txtStatus.Text)
      Me.Refresh
      svr.LogEvent "Creating Foreign Key: " & lstForeignKeys.List(j), etMiniDebug
   
        'loop through the tables and find which table it belongs to
        For X = 0 To catLocal.Tables.Count - 1
          If catLocal.Tables(X).Type = "TABLE" Then
            'Go through all the Keys in table
            For i = 0 To (catLocal.Tables(X).Keys.Count - 1)
                            
              If catLocal.Tables(X).Keys(i).Name = lstForeignKeys.List(j) Then
                If chkLCaseTables.Value = 0 Then
                  szQryStr = "ALTER TABLE " & QUOTE & catLocal.Tables(X).Name & QUOTE
                Else
                  szQryStr = "ALTER TABLE " & QUOTE & LCase(catLocal.Tables(X).Name) & QUOTE
                End If
                              
                'Reduce in size if necessary and ad _fk to end
                szTmpFKName = Left(lstForeignKeys.List(j), 28) & "_fk"
                If chkLCaseIndexes.Value = 0 Then
                  szQryStr = szQryStr & " ADD CONSTRAINT " & QUOTE & szTmpFKName & QUOTE & " FOREIGN KEY("
                Else
                  szQryStr = szQryStr & " ADD CONSTRAINT " & QUOTE & LCase(szTmpFKName) & QUOTE & " FOREIGN KEY("
                End If
                
                'Get Columns involved with FK
                szRelatedCols = ""
                For Y = 0 To catLocal.Tables(X).Keys(i).Columns.Count - 1
                  If chkLCaseColumns.Value = 0 Then
                    szQryStr = szQryStr & QUOTE & catLocal.Tables(X).Keys(i).Columns(Y).Name & QUOTE & ","
                  Else
                    szQryStr = szQryStr & QUOTE & LCase(catLocal.Tables(X).Keys(i).Columns(Y).Name) & QUOTE & ","
                  End If
                  
                  'Get the related column name while we are on this comumn
                  'The Related column belongs to the Comumns collection in the table collection
                  If chkLCaseColumns.Value = 0 Then
                    szRelatedCols = szRelatedCols & QUOTE & catLocal.Tables(X).Keys(i).Columns(catLocal.Tables(X).Keys(i).Columns(Y).Name).RelatedColumn & QUOTE & ","
                  Else
                    szRelatedCols = szRelatedCols & QUOTE & LCase(catLocal.Tables(X).Keys(i).Columns(catLocal.Tables(X).Keys(i).Columns(Y).Name).RelatedColumn) & QUOTE & ","
                  End If
                Next Y
                
                'Trim extra , off end of column names, add ) to end
                szQryStr = Left(szQryStr, (Len(szQryStr) - 1)) & ")"
                szRelatedCols = Left(szRelatedCols, (Len(szRelatedCols) - 1)) & ")"
                If chkLCaseTables.Value = 0 Then
                  szQryStr = szQryStr & " REFERENCES " & QUOTE & catLocal.Tables(X).Keys(i).RelatedTable & QUOTE & " (" & szRelatedCols
                Else
                  szQryStr = szQryStr & " REFERENCES " & QUOTE & LCase(catLocal.Tables(X).Keys(i).RelatedTable) & QUOTE & " (" & szRelatedCols
                End If
                
                'Set action to do when referenced row is being deleted
                szQryStr = szQryStr & " ON DELETE "
                Select Case catLocal.Tables(X).Keys(i).DeleteRule
                  Case adRINone
                    szQryStr = szQryStr & "NO ACTION"
                  Case adRICascade
                    szQryStr = szQryStr & "CASCADE"
                  Case adRISetNull
                    szQryStr = szQryStr & "SET NULL"
                  Case adRISetDefault
                    szQryStr = szQryStr & "SET DEFAULT"
                End Select
                
                'Set action to do when referenced row is being Updated
                szQryStr = szQryStr & " ON UPDATE "
                Select Case catLocal.Tables(X).Keys(i).UpdateRule
                  Case adRINone
                    szQryStr = szQryStr & "NO ACTION"
                  Case adRICascade
                    szQryStr = szQryStr & "CASCADE"
                  Case adRISetNull
                    szQryStr = szQryStr & "SET NULL"
                  Case adRISetDefault
                    szQryStr = szQryStr & "SET DEFAULT"
                End Select
                
                svr.Databases(szDatabase).Execute szQryStr
                
              End If
            Next i
          End If
        Next X
    End If
  Next j
    
  svr.Databases(szDatabase).Execute "COMMIT"
  txtStatus.Text = txtStatus.Text & vbCrLf & "Migration finished at: " & Now & ", taking " & Fix((Timer - Start) * 100) / 100 & " seconds."
  txtStatus.SelStart = Len(txtStatus.Text)
  svr.LogEvent "Migration Completed!", etMiniDebug
  cmdOK.Enabled = True
  cmdOK.SetFocus
  EndMsg
  
  Exit Sub
Err_Handler:
  txtStatus.Text = txtStatus.Text & vbCrLf & "An error occured at: " & Now & ": " & vbCrLf & Err.Number & ": " & Replace(Err.Description, vbLf, vbCrLf) & vbCrLf & vbCrLf & "Rolling back..."
  txtStatus.SelStart = Len(txtStatus.Text)
  svr.Databases(szDatabase).Execute "ROLLBACK"
  txtStatus.Text = txtStatus.Text & " Done."
  txtStatus.SelStart = Len(txtStatus.Text)
  cmdNext.Visible = True
  cmdPrevious.Visible = True
  cmdMigrate.Visible = True
  cmdTypeMap.Visible = True
  cmdOK.Visible = False
  EndMsg
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.Migrate_Data"
End Sub

Private Sub tabWizard_Click(PreviousTab As Integer)
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.tabWizard_Click()", etFullDebug
    
  If bButtonPress = False And bProgramPress = False Then
    bProgramPress = True
    tabWizard.Tab = PreviousTab
  Else
    bProgramPress = False
  End If
  bButtonPress = False
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.tabWizard_Click"
End Sub

Private Sub GetEligibleForeignKeys()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.GetEligibleForeignKeys()", etFullDebug

'1/16/2001 Rod Childers
'This sub will:
'1 Look at the MS Access database and see if there are any foreign keys
'2 See if the target PostgreSQL database has the necessary tables to migrate these keys
'3 Load a list box of "eligible" foregin keys to be selected for migration

Dim tblTemp As Table
Dim i As Integer
Dim X As Integer

lstForeignKeys.Clear

StartMsg "Searching for Foreign Keys..."
'Loop Through all Tables in database
For X = 0 To catLocal.Tables.Count - 1
  If catLocal.Tables(X).Type = "TABLE" Then
    'Go through all the Keys in table, find foreign keys
    For i = 0 To (catLocal.Tables(X).Keys.Count - 1)
      If catLocal.Tables(X).Keys(i).Type = adKeyForeign Then
        'See if both tables needed exist in PostgreSQL, or are to be migrated
        'if so add it to the list
        'If the table with the Forgein key is to be migrated or it is already in the PostgreSQL database
        If isTableToBeMigrated((catLocal.Tables(X).Name)) = True Or svr.Databases(szDatabase).Tables.Exists(catLocal.Tables(X).Name) Then
          'If the Related table is to be migrated or it is already in the PostgreSQL database
          If isTableToBeMigrated((catLocal.Tables(X).Keys(i).RelatedTable)) Or svr.Databases(szDatabase).Tables.Exists(catLocal.Tables(X).Keys(i).RelatedTable) Then
            lstForeignKeys.AddItem catLocal.Tables(X).Keys(i).Name
          End If
        End If
      End If
    Next i
  End If
Next X
EndMsg

Exit Sub

Err_Handler:

If Err.Number <> 0 Then
  If Err.Number = 3251 Then
    'Foreign keys can not be found using this provider
    EndMsg
    svr.LogEvent "Foreign Keys are not supported with this provider.", etMiniDebug
    MsgBox "Foreign Keys are not supported with this provider.", vbInformation, "Warning"
    Exit Sub
  Else
    EndMsg
    If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.GetEligibleForeignKeys"
  End If
End If

End Sub


Private Function isTableToBeMigrated(szTableName As String)
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.isTableToBeMigrated()", etFullDebug

'1/16/2001 Rod Childers
'This function checks if a table is to be migrated
'the lstData should contain a list of all tables that
'were selected to be migrated

Dim X As Integer

isTableToBeMigrated = False

  For X = 0 To lstData.ListCount - 1
    If lstData.List(X) = szTableName Then
      isTableToBeMigrated = True
      Exit For
    End If
  Next X

  Exit Function
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.Load_Data"
End Function

Private Sub txtStatus_Change()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.txtStatus_Change()", etFullDebug

'1/25/2001  Rod Childers
'Clear before textbox gets to 32K limit
If Len(txtStatus.Text) >= 30000 Then
  txtStatus.Text = "Log Truncated" & vbCrLf & vbCrLf & Right(txtStatus.Text, 30000)
End If

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.txtStatus_Change"
End Sub

Private Sub GetTargetDatabases()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.GetTargetDatabases()", etFullDebug

Dim objDatabase As pgDatabase
Dim objItem As ListItem

  StartMsg "Looking for possible target databases..."
  lstDatabase.ListItems.Clear
  For Each objDatabase In svr.Databases
    If Not objDatabase.SystemObject Then
      Set objItem = lstDatabase.ListItems.Add(, , objDatabase.Identifier)
      objItem.SubItems(1) = Replace(objDatabase.Comment, vbCrLf, " ")
    End If
  Next objDatabase
  EndMsg

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.GetTargetDatabases"
End Sub
Private Function GetQuoteChar(szConnect As String) As String
'This may well go wrong :-(
On Error GoTo Cleanup
Dim iStatus As Integer
Dim iSize As Integer
Dim lEnv As Long
Dim lDBC As Long
Dim szResult As String * 8

  'Initialise the ODBC subsystem
  If SQLAllocEnv(lEnv) <> 0 Then
    Exit Function
  End If

  'Allocate space for the connection object
  If SQLAllocConnect(lEnv, lDBC) <> 0 Then
    GoTo Cleanup
  End If

  'Connect
  SQLDriverConnect lDBC, Me.hWnd, szConnect, Len(szConnect), szResult, Len(szResult), iSize, SQL_DRIVER_NOPROMPT

  'Get the quote char
  szResult = ""
  SQLGetInfoString lDBC, SQL_IDENTIFIER_QUOTE_CHAR, szResult, Len(szResult), iSize
  
  GetQuoteChar = Left(szResult, iSize)
  
Cleanup:
  On Error Resume Next
  If lDBC <> 0 Then SQLDisconnect lDBC
  SQLFreeConnect lDBC
  If lEnv <> 0 Then SQLFreeEnv lEnv
End Function
