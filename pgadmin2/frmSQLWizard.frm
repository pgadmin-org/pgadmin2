VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmSQLWizard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SQL Wizard"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7530
   Icon            =   "frmSQLWizard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   7530
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   330
      Left            =   6480
      TabIndex        =   49
      ToolTipText     =   "Return SQL and exit."
      Top             =   3960
      Visible         =   0   'False
      Width           =   960
   End
   Begin TabDlg.SSTab tabWizard 
      Height          =   3840
      Left            =   495
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   45
      Width           =   6990
      _ExtentX        =   12330
      _ExtentY        =   6773
      _Version        =   393216
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   176
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   " "
      TabPicture(0)   =   "frmSQLWizard.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lstAllTables"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lstIncTables"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdAddTable"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdRemoveTable"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   " "
      TabPicture(1)   =   "frmSQLWizard.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1(1)"
      Tab(1).Control(1)=   "Label2(0)"
      Tab(1).Control(2)=   "Label2(7)"
      Tab(1).Control(3)=   "Label2(1)"
      Tab(1).Control(4)=   "cmdRemoveJoin"
      Tab(1).Control(5)=   "cmdAddJoin"
      Tab(1).Control(6)=   "lstJoins"
      Tab(1).Control(7)=   "cboJColumn1"
      Tab(1).Control(8)=   "cboJColumn2"
      Tab(1).Control(9)=   "txtPrimaryTable"
      Tab(1).Control(10)=   "Frame1"
      Tab(1).Control(11)=   "Frame2"
      Tab(1).ControlCount=   12
      TabCaption(2)   =   " "
      TabPicture(2)   =   "frmSQLWizard.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label1(2)"
      Tab(2).Control(1)=   "Label2(4)"
      Tab(2).Control(2)=   "cmdRemoveColumn"
      Tab(2).Control(3)=   "cmdAddColumn"
      Tab(2).Control(4)=   "lstIncColumns"
      Tab(2).Control(5)=   "lstAllColumns"
      Tab(2).Control(6)=   "cmdColumnUp"
      Tab(2).Control(7)=   "cmdColumnDown"
      Tab(2).Control(8)=   "cmdAddCustomColumn"
      Tab(2).Control(9)=   "cboCustomColumn"
      Tab(2).ControlCount=   10
      TabCaption(3)   =   " "
      TabPicture(3)   =   "frmSQLWizard.frx":091E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label1(3)"
      Tab(3).Control(1)=   "Label2(2)"
      Tab(3).Control(2)=   "Label2(3)"
      Tab(3).Control(3)=   "lblBoolean"
      Tab(3).Control(4)=   "lblValue"
      Tab(3).Control(5)=   "cboWhereCols"
      Tab(3).Control(6)=   "lstCriteria"
      Tab(3).Control(7)=   "cmdAddCriteria"
      Tab(3).Control(8)=   "cmdRemoveCriteria"
      Tab(3).Control(9)=   "cboOperator"
      Tab(3).Control(10)=   "cboBoolean"
      Tab(3).Control(11)=   "txtValue"
      Tab(3).ControlCount=   12
      TabCaption(4)   =   " "
      TabPicture(4)   =   "frmSQLWizard.frx":093A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label1(4)"
      Tab(4).Control(1)=   "cmdRemoveSortCol"
      Tab(4).Control(2)=   "cmdAddAsc"
      Tab(4).Control(3)=   "lstIncSortCols"
      Tab(4).Control(4)=   "lstAllSortCols"
      Tab(4).Control(5)=   "cmdAddDesc"
      Tab(4).Control(6)=   "cmdSortColDown"
      Tab(4).Control(7)=   "cmdSortColUp"
      Tab(4).ControlCount=   8
      TabCaption(5)   =   " "
      TabPicture(5)   =   "frmSQLWizard.frx":0956
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label1(7)"
      Tab(5).Control(1)=   "Label2(5)"
      Tab(5).Control(2)=   "Label2(6)"
      Tab(5).Control(3)=   "chkDistinct"
      Tab(5).Control(4)=   "chkLimit"
      Tab(5).Control(5)=   "txtLimit"
      Tab(5).Control(6)=   "chkOffset"
      Tab(5).Control(7)=   "txtOffset"
      Tab(5).ControlCount=   8
      Begin VB.Frame Frame2 
         Caption         =   "Join Type"
         Height          =   1905
         Left            =   -74865
         TabIndex        =   69
         Top             =   1170
         Width           =   1815
         Begin VB.OptionButton optJType 
            Caption         =   "Full Join"
            Height          =   285
            Index           =   3
            Left            =   135
            TabIndex        =   9
            ToolTipText     =   "Select the type of join to add."
            Top             =   1350
            Width           =   1320
         End
         Begin VB.OptionButton optJType 
            Caption         =   "Right Outer Join"
            Height          =   285
            Index           =   2
            Left            =   135
            TabIndex        =   8
            ToolTipText     =   "Select the type of join to add."
            Top             =   1035
            Width           =   1545
         End
         Begin VB.OptionButton optJType 
            Caption         =   "Left Outer Join"
            Height          =   285
            Index           =   1
            Left            =   135
            TabIndex        =   7
            ToolTipText     =   "Select the type of join to add."
            Top             =   720
            Width           =   1500
         End
         Begin VB.OptionButton optJType 
            Caption         =   "Inner Join"
            Height          =   285
            Index           =   0
            Left            =   135
            TabIndex        =   6
            ToolTipText     =   "Select the type of join to add."
            Top             =   405
            Value           =   -1  'True
            Width           =   1320
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Operator"
         Height          =   1905
         Left            =   -73020
         TabIndex        =   68
         Top             =   1170
         Width           =   870
         Begin VB.OptionButton OptJOperator 
            Caption         =   "="
            Height          =   285
            Index           =   0
            Left            =   180
            TabIndex        =   10
            ToolTipText     =   "Select the join operator to use."
            Top             =   225
            Value           =   -1  'True
            Width           =   555
         End
         Begin VB.OptionButton OptJOperator 
            Caption         =   ">"
            Height          =   285
            Index           =   1
            Left            =   180
            TabIndex        =   11
            ToolTipText     =   "Select the join operator to use."
            Top             =   495
            Width           =   555
         End
         Begin VB.OptionButton OptJOperator 
            Caption         =   "<"
            Height          =   285
            Index           =   2
            Left            =   180
            TabIndex        =   12
            ToolTipText     =   "Select the join operator to use."
            Top             =   765
            Width           =   555
         End
         Begin VB.OptionButton OptJOperator 
            Caption         =   ">="
            Height          =   285
            Index           =   3
            Left            =   180
            TabIndex        =   13
            ToolTipText     =   "Select the join operator to use."
            Top             =   1035
            Width           =   555
         End
         Begin VB.OptionButton OptJOperator 
            Caption         =   "<="
            Height          =   285
            Index           =   4
            Left            =   180
            TabIndex        =   14
            ToolTipText     =   "Select the join operator to use."
            Top             =   1305
            Width           =   555
         End
         Begin VB.OptionButton OptJOperator 
            Caption         =   "<>"
            Height          =   285
            Index           =   5
            Left            =   180
            TabIndex        =   15
            ToolTipText     =   "Select the join operator to use."
            Top             =   1575
            Width           =   555
         End
      End
      Begin VB.TextBox txtPrimaryTable 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   -71535
         Locked          =   -1  'True
         TabIndex        =   19
         ToolTipText     =   "Displays the Primary Table for the join clauses."
         Top             =   450
         Width           =   3390
      End
      Begin VB.ComboBox cboJColumn2 
         Height          =   315
         Left            =   -74865
         Style           =   2  'Dropdown List
         TabIndex        =   16
         ToolTipText     =   "Select the second column in the join."
         Top             =   3420
         Width           =   2715
      End
      Begin VB.ComboBox cboCustomColumn 
         Height          =   315
         ItemData        =   "frmSQLWizard.frx":0972
         Left            =   -74865
         List            =   "frmSQLWizard.frx":0985
         TabIndex        =   22
         ToolTipText     =   "Select or Enter a custom column name."
         Top             =   3375
         Width           =   2760
      End
      Begin VB.TextBox txtOffset 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         Height          =   285
         Left            =   -71670
         TabIndex        =   48
         Text            =   "100"
         ToolTipText     =   "Enter the amount of records to limit the resultset to."
         Top             =   2070
         Width           =   915
      End
      Begin VB.CheckBox chkOffset 
         Caption         =   "Offset resultset by"
         Height          =   195
         Left            =   -73290
         TabIndex        =   46
         ToolTipText     =   "Offset the resultset by n rows."
         Top             =   2115
         Width           =   1635
      End
      Begin VB.TextBox txtLimit 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         Height          =   285
         Left            =   -71670
         TabIndex        =   45
         Text            =   "100"
         ToolTipText     =   "Enter the amount of records to limit the resultset to."
         Top             =   1575
         Width           =   915
      End
      Begin VB.CheckBox chkLimit 
         Caption         =   "Limit resultset to"
         Height          =   195
         Left            =   -73290
         TabIndex        =   44
         Top             =   1620
         Width           =   1545
      End
      Begin VB.CheckBox chkDistinct 
         Caption         =   "&Select DISTINCT values only."
         Height          =   285
         Left            =   -73290
         TabIndex        =   43
         ToolTipText     =   "Select if you only want to retrieve distinct values."
         Top             =   1035
         Width           =   3165
      End
      Begin VB.CommandButton cmdSortColUp 
         Height          =   540
         Left            =   -68610
         Picture         =   "frmSQLWizard.frx":09C7
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Move the selected Column up the list"
         Top             =   495
         Width           =   435
      End
      Begin VB.CommandButton cmdSortColDown 
         Height          =   540
         Left            =   -68610
         Picture         =   "frmSQLWizard.frx":0E09
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Move the selected Column down the list"
         Top             =   3150
         Width           =   435
      End
      Begin VB.CommandButton cmdAddDesc 
         Caption         =   ">> (Desc)"
         Height          =   375
         Left            =   -72255
         TabIndex        =   38
         ToolTipText     =   "Add the selected column for descending sort."
         Top             =   1575
         Width           =   915
      End
      Begin VB.ListBox lstAllSortCols 
         Height          =   3180
         Left            =   -74865
         TabIndex        =   36
         ToolTipText     =   "Lists the available columns."
         Top             =   495
         Width           =   2535
      End
      Begin VB.ListBox lstIncSortCols 
         Height          =   3180
         Left            =   -71265
         TabIndex        =   40
         ToolTipText     =   "Lists the selected selected sort columns."
         Top             =   495
         Width           =   2535
      End
      Begin VB.CommandButton cmdAddAsc 
         Caption         =   ">> (Asc)"
         Height          =   375
         Left            =   -72255
         TabIndex        =   37
         ToolTipText     =   "Add the selected column for ascending sort."
         Top             =   1125
         Width           =   915
      End
      Begin VB.CommandButton cmdRemoveSortCol 
         Caption         =   "<<"
         Height          =   375
         Left            =   -72255
         TabIndex        =   39
         ToolTipText     =   "Remove the selected column."
         Top             =   2385
         Width           =   915
      End
      Begin VB.CommandButton cmdAddCustomColumn 
         Caption         =   ">>"
         Height          =   375
         Left            =   -72030
         TabIndex        =   25
         ToolTipText     =   "Add the custom column."
         Top             =   3105
         Width           =   420
      End
      Begin VB.TextBox txtValue 
         Height          =   285
         Left            =   -74865
         TabIndex        =   32
         ToolTipText     =   "Enter the value to use in the selection criteria."
         Top             =   2970
         Width           =   2760
      End
      Begin VB.ComboBox cboBoolean 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmSQLWizard.frx":124B
         Left            =   -74865
         List            =   "frmSQLWizard.frx":1255
         Style           =   2  'Dropdown List
         TabIndex        =   29
         ToolTipText     =   "Select a boolean operator."
         Top             =   990
         Width           =   1005
      End
      Begin VB.ComboBox cboOperator 
         Height          =   315
         ItemData        =   "frmSQLWizard.frx":1262
         Left            =   -74865
         List            =   "frmSQLWizard.frx":1284
         Style           =   2  'Dropdown List
         TabIndex        =   31
         ToolTipText     =   "Select an Operator to use."
         Top             =   2340
         Width           =   1995
      End
      Begin VB.CommandButton cmdRemoveCriteria 
         Caption         =   "<<"
         Height          =   375
         Left            =   -71985
         TabIndex        =   34
         ToolTipText     =   "Remove the selected criteria."
         Top             =   2250
         Width           =   420
      End
      Begin VB.CommandButton cmdAddCriteria 
         Caption         =   ">>"
         Height          =   375
         Left            =   -71985
         TabIndex        =   33
         ToolTipText     =   "Add the defined criteria."
         Top             =   1305
         Width           =   420
      End
      Begin VB.ListBox lstCriteria 
         Height          =   3375
         Left            =   -71490
         TabIndex        =   35
         ToolTipText     =   "Lists the query's selection criteria.."
         Top             =   270
         Width           =   3345
      End
      Begin VB.ComboBox cboWhereCols 
         Height          =   315
         ItemData        =   "frmSQLWizard.frx":12C3
         Left            =   -74865
         List            =   "frmSQLWizard.frx":12C5
         Style           =   2  'Dropdown List
         TabIndex        =   30
         ToolTipText     =   "Select a column to include in the 'WHERE' clause."
         Top             =   1665
         Width           =   2760
      End
      Begin VB.CommandButton cmdColumnDown 
         Height          =   540
         Left            =   -68610
         Picture         =   "frmSQLWizard.frx":12C7
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Move the selected Column down the list"
         Top             =   3150
         Width           =   435
      End
      Begin VB.CommandButton cmdColumnUp 
         Height          =   540
         Left            =   -68610
         Picture         =   "frmSQLWizard.frx":1709
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Move the selected Column up the list"
         Top             =   495
         Width           =   435
      End
      Begin VB.ListBox lstAllColumns 
         Height          =   2595
         Left            =   -74865
         TabIndex        =   21
         ToolTipText     =   "Lists the columns available for inclusion in the query."
         Top             =   495
         Width           =   2760
      End
      Begin VB.ListBox lstIncColumns 
         Height          =   3180
         Left            =   -71535
         TabIndex        =   26
         ToolTipText     =   "Lists the columns to be included in the query."
         Top             =   495
         Width           =   2805
      End
      Begin VB.CommandButton cmdAddColumn 
         Caption         =   ">>"
         Height          =   375
         Left            =   -72030
         TabIndex        =   24
         ToolTipText     =   "Add the selected column."
         Top             =   990
         Width           =   420
      End
      Begin VB.CommandButton cmdRemoveColumn 
         Caption         =   "<<"
         Height          =   375
         Left            =   -72030
         TabIndex        =   23
         ToolTipText     =   "Remove the selected column."
         Top             =   2070
         Width           =   420
      End
      Begin VB.ComboBox cboJColumn1 
         Height          =   315
         Left            =   -74865
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Select the first column in the join."
         Top             =   675
         Width           =   2715
      End
      Begin VB.ListBox lstJoins 
         Height          =   2985
         Left            =   -71535
         TabIndex        =   20
         ToolTipText     =   "Lists the selected joins."
         Top             =   765
         Width           =   3390
      End
      Begin VB.CommandButton cmdAddJoin 
         Caption         =   ">>"
         Height          =   375
         Left            =   -72030
         TabIndex        =   17
         ToolTipText     =   "Add the defined join."
         Top             =   1305
         Width           =   420
      End
      Begin VB.CommandButton cmdRemoveJoin 
         Caption         =   "<<"
         Height          =   375
         Left            =   -72030
         TabIndex        =   18
         ToolTipText     =   "Remove the selected join."
         Top             =   2250
         Width           =   420
      End
      Begin VB.CommandButton cmdRemoveTable 
         Caption         =   "<<"
         Height          =   375
         Left            =   3285
         TabIndex        =   2
         ToolTipText     =   "Remove the selected table."
         Top             =   2160
         Width           =   420
      End
      Begin VB.CommandButton cmdAddTable 
         Caption         =   ">>"
         Height          =   375
         Left            =   3285
         TabIndex        =   1
         ToolTipText     =   "Add the selected table."
         Top             =   1215
         Width           =   420
      End
      Begin VB.ListBox lstIncTables 
         Height          =   3180
         Left            =   3780
         TabIndex        =   3
         ToolTipText     =   "Lists the selected tables."
         Top             =   495
         Width           =   3075
      End
      Begin VB.ListBox lstAllTables 
         Height          =   3180
         Left            =   135
         Sorted          =   -1  'True
         TabIndex        =   0
         ToolTipText     =   "Lists the available tables."
         Top             =   495
         Width           =   3075
      End
      Begin VB.Label Label2 
         Caption         =   "Column 2"
         Height          =   195
         Index           =   1
         Left            =   -74865
         TabIndex        =   67
         Top             =   3195
         Width           =   780
      End
      Begin VB.Label Label2 
         Caption         =   "Primary Join Table"
         Height          =   195
         Index           =   7
         Left            =   -71535
         TabIndex        =   66
         Top             =   225
         Width           =   1500
      End
      Begin VB.Label Label2 
         Caption         =   "rows"
         Height          =   195
         Index           =   6
         Left            =   -70635
         TabIndex        =   65
         Top             =   2115
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "rows"
         Height          =   195
         Index           =   5
         Left            =   -70635
         TabIndex        =   64
         Top             =   1620
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Do you require any additional options? (DISTINCT, LIMIT, OFFSET)"
         Height          =   240
         Index           =   7
         Left            =   -74865
         TabIndex        =   63
         Top             =   225
         Width           =   5730
      End
      Begin VB.Label Label2 
         Caption         =   "Custom Column or Function"
         Height          =   195
         Index           =   4
         Left            =   -74865
         TabIndex        =   62
         Top             =   3150
         Width           =   2040
      End
      Begin VB.Label lblValue 
         Caption         =   "Value"
         Height          =   195
         Left            =   -74865
         TabIndex        =   61
         Top             =   2745
         Width           =   780
      End
      Begin VB.Label lblBoolean 
         Caption         =   "Boolean"
         Enabled         =   0   'False
         Height          =   195
         Left            =   -74865
         TabIndex        =   60
         Top             =   765
         Width           =   780
      End
      Begin VB.Label Label2 
         Caption         =   "Operator"
         Height          =   195
         Index           =   3
         Left            =   -74865
         TabIndex        =   59
         Top             =   2115
         Width           =   780
      End
      Begin VB.Label Label2 
         Caption         =   "Column"
         Height          =   195
         Index           =   2
         Left            =   -74865
         TabIndex        =   58
         Top             =   1440
         Width           =   780
      End
      Begin VB.Label Label1 
         Caption         =   "How do you want the data sorted? (ORDER BY)"
         Height          =   240
         Index           =   4
         Left            =   -74865
         TabIndex        =   57
         Top             =   225
         Width           =   4515
      End
      Begin VB.Label Label1 
         Caption         =   "What selection criteria do you require? (WHERE)"
         Height          =   465
         Index           =   3
         Left            =   -74865
         TabIndex        =   56
         Top             =   225
         Width           =   3300
      End
      Begin VB.Label Label1 
         Caption         =   "What columns do you wish to include in the query?"
         Height          =   240
         Index           =   2
         Left            =   -74865
         TabIndex        =   55
         Top             =   225
         Width           =   3705
      End
      Begin VB.Label Label2 
         Caption         =   "Column 1"
         Height          =   195
         Index           =   0
         Left            =   -74865
         TabIndex        =   54
         Top             =   450
         Width           =   780
      End
      Begin VB.Label Label1 
         Caption         =   "How are the selected tables joined?"
         Height          =   240
         Index           =   1
         Left            =   -74865
         TabIndex        =   53
         Top             =   225
         Width           =   3705
      End
      Begin VB.Label Label1 
         Caption         =   "What tables do you want to include in your query?"
         Height          =   240
         Index           =   0
         Left            =   135
         TabIndex        =   52
         Top             =   225
         Width           =   3705
      End
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next"
      Height          =   330
      Left            =   6480
      TabIndex        =   51
      ToolTipText     =   "Move forward a stage"
      Top             =   3960
      Width           =   960
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "&Previous"
      Enabled         =   0   'False
      Height          =   330
      Left            =   5445
      TabIndex        =   47
      ToolTipText     =   "Move back a stage"
      Top             =   3960
      Width           =   960
   End
   Begin VB.PictureBox picStrip 
      BorderStyle     =   0  'None
      Height          =   4335
      Left            =   0
      Picture         =   "frmSQLWizard.frx":1B4B
      ScaleHeight     =   289
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   50
      Top             =   0
      Width           =   465
   End
End
Attribute VB_Name = "frmSQLWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence
'
' frmSQLWizard.frm - Does exactly what it says on the tin! (UK joke...)

Option Explicit
Dim bButtonPress As Boolean
Dim bProgramPress As Boolean
Dim szDatabase As String

Private Sub Get_Tables()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.Get_Tables()", etFullDebug

Dim objTable As pgTable

  StartMsg "Getting Tables..."
  lstAllTables.Clear
  lstIncTables.Clear
  
  For Each objTable In frmMain.svr.Databases(szDatabase).Tables
    lstAllTables.AddItem QUOTE & objTable.Name & QUOTE
  Next objTable

  EndMsg
  
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.Get_Tables"
End Sub

Private Sub Get_JoinCols()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.Get_JoinCols()", etFullDebug

Dim X As Integer
Dim objColumn As pgColumn

  StartMsg "Getting Columns..."
  cboJColumn1.Clear
  cboJColumn2.Clear
  lstJoins.Clear
  txtPrimaryTable.Text = ""
  
  For X = 0 To lstIncTables.ListCount - 1
    For Each objColumn In frmMain.svr.Databases(szDatabase).Tables(Mid(lstIncTables.List(X), 2, Len(lstIncTables.List(X)) - 2)).Columns
      cboJColumn1.AddItem lstIncTables.List(X) & "." & QUOTE & objColumn.Name & QUOTE
      cboJColumn2.AddItem lstIncTables.List(X) & "." & QUOTE & objColumn.Name & QUOTE
    Next objColumn
  Next X
  
  EndMsg
  
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.Get_JoinCols"
End Sub

Private Sub Get_ValidJoinCols()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.Get_ValidJoinCols()", etFullDebug

Dim objColumn As pgColumn
Dim X As Integer
Dim Y As Integer
Dim szTable As String
Dim iStart As Integer
Dim bFlag As Boolean

  StartMsg "Getting Columns..."
  
  'Clear down
  cboJColumn1.Clear
  cboJColumn2.Clear
  
  'Add columns from the primary table to list1
  szTable = Mid(txtPrimaryTable.Text, 2, Len(txtPrimaryTable.Text) - 2)
  For Each objColumn In frmMain.svr.Databases(szDatabase).Tables(szTable).Columns
    cboJColumn1.AddItem QUOTE & szTable & QUOTE & "." & QUOTE & objColumn.Name & QUOTE
  Next objColumn
  
  'Add columns from other tables to list1
  For X = 0 To lstJoins.ListCount - 1
    iStart = InStr(1, lstJoins.List(X), "JOIN " & QUOTE) + 6
    szTable = Mid(lstJoins.List(X), iStart, InStr(iStart + 1, lstJoins.List(X), QUOTE) - iStart)
    For Each objColumn In frmMain.svr.Databases(szDatabase).Tables(szTable).Columns
      cboJColumn1.AddItem QUOTE & szTable & QUOTE & "." & QUOTE & objColumn.Name & QUOTE
    Next objColumn
  Next
  
  'Now we need to add columns to list2 that aren't in list1
  For X = 0 To lstIncTables.ListCount - 1
    bFlag = False
    For Y = 0 To cboJColumn1.ListCount - 1
      If Mid(cboJColumn1.List(Y), 1, InStr(2, cboJColumn1.List(Y), QUOTE & "." & QUOTE)) = lstIncTables.List(X) Then
        bFlag = True
        Exit For
      End If
    Next Y
    If bFlag = False Then
      For Each objColumn In frmMain.svr.Databases(szDatabase).Tables(Mid(lstIncTables.List(X), 2, Len(lstIncTables.List(X)) - 2)).Columns
        cboJColumn2.AddItem lstIncTables.List(X) & "." & QUOTE & objColumn.Name & QUOTE
      Next objColumn
    End If
  Next X
  
  EndMsg
  
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.Get_ValidJoinCols"
End Sub

Private Sub Get_Columns()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.Get_Columns()", etFullDebug

Dim X As Integer
Dim objColumn As pgColumn

  StartMsg "Getting Columns..."
  lstAllColumns.Clear
  lstIncColumns.Clear
  For X = 0 To lstIncTables.ListCount - 1
    For Each objColumn In frmMain.svr.Databases(szDatabase).Tables(Mid(lstIncTables.List(X), 2, Len(lstIncTables.List(X)) - 2)).Columns
      lstAllColumns.AddItem lstIncTables.List(X) & "." & QUOTE & objColumn.Name & QUOTE
    Next objColumn
  Next X
  
  EndMsg
  
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.Get_Columns"
End Sub

Private Sub Get_WhereCols()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.Get_WhereCols()", etFullDebug

Dim X As Integer
Dim objColumn As pgColumn

  StartMsg "Getting Columns..."
  cboWhereCols.Clear
  lstCriteria.Clear
  
  For X = 0 To lstIncTables.ListCount - 1
    For Each objColumn In frmMain.svr.Databases(szDatabase).Tables(Mid(lstIncTables.List(X), 2, Len(lstIncTables.List(X)) - 2)).Columns
      cboWhereCols.AddItem lstIncTables.List(X) & "." & QUOTE & objColumn.Name & QUOTE
    Next objColumn
  Next X

  EndMsg
  
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.Get_WhereCols"
End Sub

Private Sub Get_SortCols()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.Get_SortCols()", etFullDebug

Dim X As Integer
Dim objColumn As pgColumn

  StartMsg "Getting Columns..."
  lstAllSortCols.Clear
  lstIncSortCols.Clear
  
  For X = 0 To lstIncTables.ListCount - 1
    For Each objColumn In frmMain.svr.Databases(szDatabase).Tables(Mid(lstIncTables.List(X), 2, Len(lstIncTables.List(X)) - 2)).Columns
      lstAllSortCols.AddItem lstIncTables.List(X) & "." & QUOTE & objColumn.Name & QUOTE
    Next objColumn
  Next X
  
  EndMsg
  
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.Get_SortCols"
End Sub

Private Sub cboJColumn1_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.cboJColumn1_Click()", etFullDebug

  cboJColumn1.ToolTipText = cboJColumn1.Text
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.cboJColumn1_Click"
End Sub

Private Sub cboJColumn2_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.cboJColumn2_Click()", etFullDebug

  cboJColumn2.ToolTipText = cboJColumn2.Text
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.cboJColumn2_Click"
End Sub

Private Sub cboOperator_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.cboOperator_Click()", etFullDebug

  If cboOperator.Text = "IS NULL" Or cboOperator.Text = "IS NOT NULL" Then
    txtValue.Enabled = False
    lblValue.Enabled = False
  Else
    txtValue.Enabled = True
    lblValue.Enabled = True
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.cboOperator_Change"
End Sub

Private Sub chkLimit_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.chkLimit_Click()", etFullDebug

  If chkLimit.Value = 1 Then
    txtLimit.Enabled = True
    chkOffset.Enabled = True
  Else
    txtLimit.Enabled = False
    chkOffset.Value = 0
    chkOffset.Enabled = False
    txtOffset.Enabled = False
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.chkLimit_Click"
End Sub

Private Sub chkOffset_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.chkOffset_Click()", etFullDebug

  If chkOffset.Value = 1 Then
    txtOffset.Enabled = True
  Else
    txtOffset.Enabled = False
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.chkOffset_Click"
End Sub

Private Sub cmdAddAsc_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.cmdAddAsc_Click()", etFullDebug

Dim iItem As Integer

  If lstAllSortCols.Text = "" Then
    MsgBox "You must select a column to add!", vbExclamation, "Error"
    Exit Sub
  End If
  lstIncSortCols.AddItem lstAllSortCols.Text & " ASC"
  iItem = lstAllSortCols.ListIndex - 1
  If iItem < 0 Then iItem = 0
  lstAllSortCols.RemoveItem lstAllSortCols.ListIndex
  If lstAllSortCols.ListCount > 0 Then lstAllSortCols.Selected(iItem) = True
  lstAllSortCols.ToolTipText = lstAllSortCols.Text
  lstIncSortCols.ToolTipText = lstIncSortCols.Text
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.cmdAddAsc_Click"
End Sub

Private Sub cmdAddDesc_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.cmdAddDesc_Click()", etFullDebug

Dim iItem As Integer

  If lstAllSortCols.Text = "" Then
    MsgBox "You must select a column to add!", vbExclamation, "Error"
    Exit Sub
  End If
  lstIncSortCols.AddItem lstAllSortCols.Text & " DESC"
  iItem = lstAllSortCols.ListIndex - 1
  If iItem < 0 Then iItem = 0
  lstAllSortCols.RemoveItem lstAllSortCols.ListIndex
  If lstAllSortCols.ListCount > 0 Then lstAllSortCols.Selected(iItem) = True
  lstAllSortCols.ToolTipText = lstAllSortCols.Text
  lstIncSortCols.ToolTipText = lstIncSortCols.Text
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.cmdAddDesc_Click"
End Sub

Private Sub cmdAddColumn_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.cmdAddColumn_Click()", etFullDebug

Dim iItem As Integer

  If lstAllColumns.Text = "" Then
    MsgBox "You must select a column to add!", vbExclamation, "Error"
    Exit Sub
  End If
  lstIncColumns.AddItem lstAllColumns.Text
  iItem = lstAllColumns.ListIndex - 1
  If iItem < 0 Then iItem = 0
  lstAllColumns.RemoveItem lstAllColumns.ListIndex
  If lstAllColumns.ListCount > 0 Then lstAllColumns.Selected(iItem) = True
  lstAllColumns.ToolTipText = lstAllColumns.Text
  lstIncColumns.ToolTipText = lstIncColumns.Text
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.cmdAddColumn_Click"
End Sub

Private Sub cmdAddCriteria_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.cmdAddCriteria_Click()", etFullDebug

  If cboBoolean.Enabled = True And cboBoolean.Text = "" Then
    MsgBox "You must select a boolean operator!", vbExclamation, "Error"
    Exit Sub
  End If
  If cboWhereCols.Text = "" Then
    MsgBox "You must select a column!", vbExclamation, "Error"
    Exit Sub
  End If
  If cboOperator.Text = "" Then
    MsgBox "You must select an operator!", vbExclamation, "Error"
    Exit Sub
  End If
  If txtValue.Text = "" And cboOperator.Text <> "IS NULL" And cboOperator.Text <> "IS NOT NULL" Then
    MsgBox "You must enter a value for the criteria!", vbExclamation, "Error"
    Exit Sub
  End If
  If cboBoolean.Enabled = False Then
    If cboOperator.Text = "LIKE" Then
      lstCriteria.AddItem cboWhereCols.Text & " ~~ " & txtValue.Text
    ElseIf cboOperator.Text = "NOT LIKE" Then
      lstCriteria.AddItem cboWhereCols.Text & " !~~ " & txtValue.Text
    ElseIf cboOperator.Text = "IS NULL" Or cboOperator = "IS NOT NULL" Then
      lstCriteria.AddItem cboWhereCols.Text & " " & cboOperator.Text
    Else
      lstCriteria.AddItem cboWhereCols.Text & " " & cboOperator.Text & " " & txtValue.Text
    End If
  Else
    If cboOperator.Text = "LIKE" Then
      lstCriteria.AddItem cboBoolean.Text & " " & cboWhereCols.Text & " ~~ " & txtValue.Text
    ElseIf cboOperator.Text = "NOT LIKE" Then
      lstCriteria.AddItem cboBoolean.Text & " " & cboWhereCols.Text & " !~~ " & txtValue.Text
    ElseIf cboOperator.Text = "IS NULL" Or cboOperator = "IS NOT NULL" Then
      lstCriteria.AddItem cboBoolean.Text & " " & cboWhereCols.Text & " " & cboOperator.Text
    Else
      lstCriteria.AddItem cboBoolean.Text & " " & cboWhereCols.Text & " " & cboOperator.Text & " " & txtValue.Text
    End If
  End If
  lblBoolean.Enabled = True
  cboBoolean.Enabled = True
  lstCriteria.ToolTipText = lstCriteria.Text
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.cmdAddCriteria_Click"
End Sub

Private Sub cmdAddCustomColumn_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.cmdAddCustomColumn_Click()", etFullDebug

  If cboCustomColumn.Text = "" Then Exit Sub
  lstIncColumns.AddItem cboCustomColumn.Text
  lstAllColumns.ToolTipText = lstAllColumns.Text
  lstIncColumns.ToolTipText = lstIncColumns.Text
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.cmdAddColumn_Click"
End Sub

Private Sub cmdAddJoin_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.cmdAddJoin_Click()", etFullDebug

Dim szTable1 As String
Dim szTable2 As String
Dim szOperator As String
Dim szType As String
  
  'Error Checks
  If cboJColumn1.Text = "" Then
    MsgBox "You must select the first join column!", vbExclamation, "Error"
    Exit Sub
  End If
  If cboJColumn2.Text = "" Then
    MsgBox "You must select the second join column!", vbExclamation, "Error"
    Exit Sub
  End If
  If cboJColumn2.Text = cboJColumn1.Text Then
    MsgBox "You cannot join a column to itself!", vbExclamation, "Error"
    Exit Sub
  End If
  
  'Get the table names
  szTable1 = Mid(cboJColumn1.Text, 1, InStr(1, cboJColumn1.Text, QUOTE & "." & QUOTE))
  szTable2 = Mid(cboJColumn2.Text, 1, InStr(1, cboJColumn2.Text, QUOTE & "." & QUOTE))
  
  'If this is the first join then set the primary table.
  If txtPrimaryTable.Text = "" Then txtPrimaryTable.Text = szTable1
  
  'Get the Join Type
  If optJType(0).Value = True Then
    szType = "INNER"
  ElseIf optJType(1).Value = True Then
    szType = "LEFT OUTER"
  ElseIf optJType(2).Value = True Then
    szType = "RIGHT OUTER"
  ElseIf optJType(3).Value = True Then
    szType = "FULL"
  End If
  
  'Get the Join Operator
  If OptJOperator(0).Value = True Then
    szOperator = "="
  ElseIf OptJOperator(1).Value = True Then
    szOperator = ">"
  ElseIf OptJOperator(2).Value = True Then
    szOperator = "<"
  ElseIf OptJOperator(3).Value = True Then
    szOperator = ">="
  ElseIf OptJOperator(4).Value = True Then
    szOperator = "<="
  ElseIf OptJOperator(5).Value = True Then
    szOperator = "<>"
  End If
  
  'Add the Join and reset for next.
  lstJoins.AddItem szType & " JOIN " & szTable2 & " ON " & cboJColumn1.Text & " " & szOperator & " " & cboJColumn2.Text
  lstJoins.ToolTipText = lstJoins.Text
  Get_ValidJoinCols
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.cmdAddJoin_Click"
End Sub

Private Sub cmdAddTable_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.cmdAddTable_Click()", etFullDebug

Dim iItem As Integer

  If lstAllTables.Text = "" Then
    MsgBox "You must select a table to add!", vbExclamation, "Error"
    Exit Sub
  End If
  lstIncTables.AddItem lstAllTables.Text
  iItem = lstAllTables.ListIndex - 1
  If iItem < 0 Then iItem = 0
  lstAllTables.RemoveItem lstAllTables.ListIndex
  If lstAllTables.ListCount > 0 Then lstAllTables.Selected(iItem) = True
  lstAllTables.ToolTipText = lstAllTables.Text
  lstIncTables.ToolTipText = lstIncTables.Text
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.cmdAddTable_Click"
End Sub

Private Sub cmdColumnDown_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.cmdColumnDown_Click()", etFullDebug

Dim szTemp As String

  If lstIncColumns.ListIndex = -1 Then
    MsgBox "You must select a column to move!", vbExclamation, "Error"
    Exit Sub
  End If
  If lstIncColumns.ListIndex = lstIncColumns.ListCount - 1 Then
    MsgBox "This column is already at the bottom!", vbExclamation, "Error"
    Exit Sub
  End If
  szTemp = lstIncColumns.List(lstIncColumns.ListIndex + 1)
  lstIncColumns.List(lstIncColumns.ListIndex + 1) = lstIncColumns.List(lstIncColumns.ListIndex)
  lstIncColumns.List(lstIncColumns.ListIndex) = szTemp
  lstIncColumns.ListIndex = lstIncColumns.ListIndex + 1
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.cmdColumnDown_Click"
End Sub

Private Sub cmdColumnUp_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.cmdColumnUp_Click()", etFullDebug

Dim szTemp As String

  If lstIncColumns.ListIndex = -1 Then
    MsgBox "You must select a column to move!", vbExclamation, "Error"
    Exit Sub
  End If
  If lstIncColumns.ListIndex = 0 Then
    MsgBox "This column is already at the top!", vbExclamation, "Error"
    Exit Sub
  End If
  szTemp = lstIncColumns.List(lstIncColumns.ListIndex - 1)
  lstIncColumns.List(lstIncColumns.ListIndex - 1) = lstIncColumns.List(lstIncColumns.ListIndex)
  lstIncColumns.List(lstIncColumns.ListIndex) = szTemp
  lstIncColumns.ListIndex = lstIncColumns.ListIndex - 1
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.cmdColumnUp_Click"
End Sub

Private Sub cmdOK_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.cmdOK_Click()", etFullDebug

Dim szSQL As String
Dim szColumns As String
Dim szJoins As String
Dim szTables As String
Dim szFrom As String
Dim szCriteria As String
Dim szOrderBy As String
Dim X As Integer

  'Column
  For X = 0 To lstIncColumns.ListCount - 1
    szColumns = szColumns & "  " & lstIncColumns.List(X) & ", " & vbCrLf
  Next
  If Len(szColumns) > 4 Then szColumns = Mid(szColumns, 1, Len(szColumns) - 4)
  
  'Joins
  If lstJoins.ListCount >= 1 Then szJoins = "  " & txtPrimaryTable.Text & vbCrLf
  For X = 0 To lstJoins.ListCount - 1
    szJoins = szJoins & "  " & lstJoins.List(X) & vbCrLf
  Next
  
  'Only add tables that aren't in any joins
  For X = 0 To lstIncTables.ListCount - 1
    If InStr(1, szJoins, lstIncTables.List(X)) = 0 Then
      szTables = szTables & "  " & lstIncTables.List(X) & ", " & vbCrLf
    End If
  Next
  If Len(szJoins) > 7 Then szJoins = Mid(szJoins, 1, Len(szJoins) - 1)
  If szJoins <> "" Then
    szFrom = szJoins & "," & vbCrLf & szTables
  Else
    szFrom = szTables
  End If
  If Len(szFrom) > 4 Then szFrom = Mid(szFrom, 1, Len(szFrom) - 4)
  
  'Criteria
  For X = 0 To lstCriteria.ListCount - 1
    szCriteria = szCriteria & "  " & lstCriteria.List(X) & " " & vbCrLf
  Next
  
  'Sorting
  For X = 0 To lstIncSortCols.ListCount - 1
    szOrderBy = szOrderBy & "  " & lstIncSortCols.List(X) & ", " & vbCrLf
  Next
  If Len(szOrderBy) > 4 Then szOrderBy = Mid(szOrderBy, 1, Len(szOrderBy) - 4)
  
  'Select Type
  If chkDistinct.Value = 1 Then
    szSQL = "SELECT DISTINCT" & vbCrLf & szColumns & vbCrLf
  Else
    szSQL = "SELECT " & vbCrLf & szColumns & vbCrLf
  End If
  
  'Build the main query
  szSQL = szSQL & "FROM " & vbCrLf & szFrom & vbCrLf
  If szCriteria <> "" Then szSQL = szSQL & "WHERE " & vbCrLf & szCriteria
  If szOrderBy <> "" Then szSQL = szSQL & "ORDER BY " & vbCrLf & szOrderBy & vbCrLf
  
  'Add any options
  If chkLimit.Value = 1 Then szSQL = szSQL & "LIMIT " & txtLimit.Text & " " & vbCrLf
  If chkOffset.Value = 1 Then szSQL = szSQL & "OFFSET " & txtOffset.Text & " " & vbCrLf
    
  For X = 0 To Forms.Count - 1
    If Forms(X).hWnd = Me.Tag Then Exit For
  Next
  If X = Forms.Count Then
    MsgBox "The SQL dialog that this wizard was initiated from appears to have been closed!", vbCritical, "Fatal Error"
    Unload Me
    Exit Sub
  End If
  Forms(X).txtSQL.Text = szSQL
  Unload Me
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.cmdOK_Click"
End Sub

Private Sub cmdRemoveColumn_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.cmdRemoveColumn_Click()", etFullDebug

Dim iItem As Integer

  If lstIncColumns.Text = "" Then
    MsgBox "You must select a column to remove!", vbExclamation, "Error"
    Exit Sub
  End If
  lstAllColumns.AddItem lstIncColumns.Text
  iItem = lstIncColumns.ListIndex - 1
  If iItem < 0 Then iItem = 0
  lstIncColumns.RemoveItem lstIncColumns.ListIndex
  If lstIncColumns.ListCount > 0 Then lstIncColumns.Selected(iItem) = True
  lstAllColumns.ToolTipText = lstAllColumns.Text
  lstIncColumns.ToolTipText = lstIncColumns.Text
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.cmdRemoveColumn_Click"
End Sub

Private Sub cmdRemoveCriteria_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.cmdRemoveCriteria_Click()", etFullDebug

Dim iItem As Integer

  If lstCriteria.Text = "" Then
    MsgBox "You must select a join to remove!", vbExclamation, "Error"
    Exit Sub
  End If
  If lstCriteria.ListCount > 1 And lstCriteria.ListIndex = 0 Then
    MsgBox "You must remove all other criteria before you can remove the first!", vbExclamation, "Error"
    Exit Sub
  End If
  If lstCriteria.ListCount = 1 And lstCriteria.ListIndex = 0 Then
    cboBoolean.Enabled = False
    lblBoolean.Enabled = False
  End If
  iItem = lstCriteria.ListIndex - 1
  If iItem < 0 Then iItem = 0
  lstCriteria.RemoveItem lstCriteria.ListIndex
  If lstCriteria.ListCount > 0 Then lstCriteria.Selected(iItem) = True
  lstCriteria.ToolTipText = lstCriteria.Text
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.cmdRemoveCriteria_Click"
End Sub

Private Sub cmdRemoveJoin_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.cmdRemoveJoin_Click()", etFullDebug

Dim iItem As Integer

  If lstJoins.Text = "" Then
    MsgBox "You must select a join to remove!", vbExclamation, "Error"
    Exit Sub
  End If
  iItem = lstJoins.ListIndex - 1
  If iItem < 0 Then iItem = 0
  lstJoins.RemoveItem lstJoins.ListIndex
  'Set the selected item if there is one, else clear the primary table
  If lstJoins.ListCount > 0 Then
    lstJoins.Selected(iItem) = True
    Get_ValidJoinCols
  Else
    txtPrimaryTable.Text = ""
    Get_JoinCols
  End If
  lstJoins.ToolTipText = lstJoins.Text
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.cmdRemoveJoin_Click"
End Sub

Private Sub cmdRemoveTable_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.cmdRemoveTable_Click()", etFullDebug

Dim iItem As Integer

  If lstIncTables.Text = "" Then
    MsgBox "You must select a table to remove!", vbExclamation, "Error"
    Exit Sub
  End If
  lstAllTables.AddItem lstIncTables.Text
  iItem = lstIncTables.ListIndex - 1
  If iItem < 0 Then iItem = 0
  lstIncTables.RemoveItem lstIncTables.ListIndex
  If lstIncTables.ListCount > 0 Then lstIncTables.Selected(iItem) = True
  lstAllTables.ToolTipText = lstAllTables.Text
  lstIncTables.ToolTipText = lstIncTables.Text
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.cmdRemoveTable_Click"
End Sub

Private Sub cmdNext_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.cmdNext_Click()", etFullDebug

  bButtonPress = True
  Select Case tabWizard.Tab
    Case 0
      If lstIncTables.ListCount = 0 Then Exit Sub
      If lstIncTables.ListCount = 1 Then
        tabWizard.Tab = 2
        Get_Columns
      Else
        tabWizard.Tab = 1
        Get_JoinCols
      End If
      cmdNext.Enabled = True
      cmdPrevious.Enabled = True
    Case 1
      tabWizard.Tab = 2
      Get_Columns
      cmdNext.Enabled = True
      cmdPrevious.Enabled = True
    Case 2
      If lstIncColumns.ListCount = 0 Then Exit Sub
      tabWizard.Tab = 3
      Get_WhereCols
      cmdNext.Enabled = True
      cmdPrevious.Enabled = True
    Case 3
      tabWizard.Tab = 4
      Get_SortCols
      cmdNext.Enabled = True
      cmdPrevious.Enabled = True
    Case 4
      tabWizard.Tab = 5
      cmdNext.Enabled = False
      cmdNext.Visible = False
      cmdOK.Enabled = True
      cmdOK.Visible = True
      cmdPrevious.Enabled = True
  End Select
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.cmdNext_Click"
End Sub

Private Sub cmdPrevious_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.cmdPrevious_Click()", etFullDebug

  bButtonPress = True
  Select Case tabWizard.Tab
    Case 5
      tabWizard.Tab = 4
      cmdNext.Enabled = True
      cmdNext.Visible = True
      cmdOK.Enabled = False
      cmdOK.Visible = False
      cmdPrevious.Enabled = True
    Case 4
      tabWizard.Tab = 3
      cmdNext.Enabled = True
      cmdPrevious.Enabled = True
    Case 3
      tabWizard.Tab = 2
      cmdNext.Enabled = True
      cmdPrevious.Enabled = True
    Case 2
      If lstIncTables.ListCount = 1 Then
        tabWizard.Tab = 0
      Else
        tabWizard.Tab = 1
      End If
      cmdNext.Enabled = True
      cmdPrevious.Enabled = True
    Case 1
      tabWizard.Tab = 0
      cmdNext.Enabled = True
      cmdPrevious.Enabled = False
  End Select
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.cmdPrevious_Click"
End Sub

Private Sub cmdRemoveSortCol_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.cmdRemoveSortCol_Click()", etFullDebug

Dim iItem As Integer

  If lstIncSortCols.Text = "" Then
    MsgBox "You must select column to remove!", vbExclamation, "Error"
    Exit Sub
  End If
  iItem = lstIncSortCols.ListIndex - 1
  If iItem < 0 Then iItem = 0
  If Mid(lstIncSortCols.Text, Len(lstIncSortCols.Text) - 3, 4) = "DESC" Then
    lstAllSortCols.AddItem Mid(lstIncSortCols.Text, 1, Len(lstIncSortCols.Text) - 5)
  Else
    lstAllSortCols.AddItem Mid(lstIncSortCols.Text, 1, Len(lstIncSortCols.Text) - 4)
  End If
  lstIncSortCols.RemoveItem lstIncSortCols.ListIndex
  If lstIncSortCols.ListCount > 0 Then lstIncSortCols.Selected(iItem) = True
  lstAllSortCols.ToolTipText = lstAllSortCols.Text
  lstIncSortCols.ToolTipText = lstIncSortCols.Text
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.cmdRemoveSortCol_Click"
End Sub

Private Sub cmdSortColDown_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.cmdSortColDown_Click()", etFullDebug

Dim Temp As String

  If lstIncSortCols.ListIndex = -1 Then
    MsgBox "You must select a column to move!", vbExclamation, "Error"
    Exit Sub
  End If
  If lstIncSortCols.ListIndex = lstIncSortCols.ListCount - 1 Then
    MsgBox "This column is already at the bottom!", vbExclamation, "Error"
    Exit Sub
  End If
  Temp = lstIncSortCols.List(lstIncSortCols.ListIndex + 1)
  lstIncSortCols.List(lstIncSortCols.ListIndex + 1) = lstIncSortCols.List(lstIncSortCols.ListIndex)
  lstIncSortCols.List(lstIncSortCols.ListIndex) = Temp
  lstIncSortCols.ListIndex = lstIncSortCols.ListIndex + 1
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.cmdSortColDown_Click"
End Sub

Private Sub cmdSortColUp_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.cmdSortColUp_Click()", etFullDebug

Dim Temp As String

  If lstIncSortCols.ListIndex = -1 Then
    MsgBox "You must select a column to move!", vbExclamation, "Error"
    Exit Sub
  End If
  If lstIncSortCols.ListIndex = 0 Then
    MsgBox "This column is already at the top!", vbExclamation, "Error"
    Exit Sub
  End If
  Temp = lstIncSortCols.List(lstIncSortCols.ListIndex - 1)
  lstIncSortCols.List(lstIncSortCols.ListIndex - 1) = lstIncSortCols.List(lstIncSortCols.ListIndex)
  lstIncSortCols.List(lstIncSortCols.ListIndex) = Temp
  lstIncSortCols.ListIndex = lstIncSortCols.ListIndex - 1
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.cmdSortColUp_Click"
End Sub

Private Sub lstAllColumns_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.lstAllColumns_Click()", etFullDebug

  lstAllColumns.ToolTipText = lstAllColumns.Text
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.lstAllColumns_Click"
End Sub

Private Sub lstAllSortCols_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.lstAllSortCols_Click()", etFullDebug

  lstAllSortCols.ToolTipText = lstAllSortCols.Text
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.lstAllSortCols_Click"
End Sub

Private Sub lstAllTables_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.lstAllTables_Click()", etFullDebug

  lstAllTables.ToolTipText = lstAllTables.Text
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.lstAllTables_Click"
End Sub

Private Sub lstCriteria_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.lstCriteria_Click()", etFullDebug

  lstCriteria.ToolTipText = lstCriteria.Text
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.lstCriteria_Click"
End Sub

Private Sub lstCriteria_DblClick()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.lstCriteria_DblClick()", etFullDebug

  cmdRemoveCriteria_Click
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.lstCriteria_DblClick"
End Sub

Private Sub lstIncColumns_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.lstIncColumns_Click()", etFullDebug

  lstIncColumns.ToolTipText = lstIncColumns.Text
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.lstIncColumns_Click"
End Sub

Private Sub lstIncSortCols_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.lstIncSortCols_Click()", etFullDebug

  lstIncSortCols.ToolTipText = lstIncSortCols.Text
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.lstIncSortCols_Click"
End Sub

Private Sub lstIncsortCols_DblClick()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.lstIncsortCols_DblClick()", etFullDebug

  cmdRemoveSortCol_Click
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.lstIncsortCols_DblClick"
End Sub

Private Sub lstAllSortCols_DblClick()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.lstAllSortCols_DblClick()", etFullDebug

  cmdAddAsc_Click
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.lstAllsortCols_DblClick"
End Sub

Private Sub lstIncColumns_DblClick()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.lstIncColumns_DblClick()", etFullDebug

  cmdRemoveColumn_Click
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.lstIncColumns_DblClick"
End Sub

Private Sub lstAllColumns_DblClick()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.lstAllColumns_DblClick()", etFullDebug

  cmdAddColumn_Click
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.lstAllColumns_DblClick"
End Sub

Private Sub lstIncTables_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.lstIncTables_Click()", etFullDebug

  lstIncTables.ToolTipText = lstIncTables.Text
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.lstIncTables_Click"
End Sub

Private Sub lstIncTables_DblClick()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.lstIncTables_DblClick()", etFullDebug

  cmdRemoveTable_Click
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.lstIncTables_DblClick"
End Sub

Private Sub lstAllTables_DblClick()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.lstAllTables_DblClick()", etFullDebug

  cmdAddTable_Click
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.lstAllTables_DblClick"
End Sub

Private Sub lstJoins_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.lstJoins_Click()", etFullDebug

  lstJoins.ToolTipText = lstJoins.Text
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.lstJoins_Click"
End Sub

Private Sub tabWizard_Click(PreviousTab As Integer)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.tabWizard_Click(" & PreviousTab & ")", etFullDebug

  If bButtonPress = False And bProgramPress = False Then
    bProgramPress = True
    tabWizard.Tab = PreviousTab
  Else
    bProgramPress = False
  End If
  bButtonPress = False
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.tabWizard_Click"
End Sub

Public Sub Initialise(szDB As String)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.Initialise(" & QUOTE & szDB & QUOTE & ")", etFullDebug

Dim sVersion As Single

  szDatabase = szDB
  tabWizard.Tab = 0
  
  'Can only do OJ's on PostgreSQL 7.1+
  sVersion = Val(frmMain.svr.dbVersion.Major & "." & frmMain.svr.dbVersion.Minor)
  If sVersion >= 7.1 Then
    optJType(1).Enabled = True
    optJType(2).Enabled = True
    optJType(3).Enabled = True
  Else
    optJType(1).Enabled = False
    optJType(2).Enabled = False
    optJType(3).Enabled = False
  End If
  Get_Tables
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.Form_Load"
End Sub


