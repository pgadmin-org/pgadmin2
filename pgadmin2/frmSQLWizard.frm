VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
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
      TabPicture(0)   =   "frmSQLWizard.frx":0BC2
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
      TabPicture(1)   =   "frmSQLWizard.frx":0BDE
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
      TabPicture(2)   =   "frmSQLWizard.frx":0BFA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cboCustomColumn"
      Tab(2).Control(1)=   "cmdAddCustomColumn"
      Tab(2).Control(2)=   "cmdColumnDown"
      Tab(2).Control(3)=   "cmdColumnUp"
      Tab(2).Control(4)=   "lstAllColumns"
      Tab(2).Control(5)=   "lstIncColumns"
      Tab(2).Control(6)=   "cmdAddColumn"
      Tab(2).Control(7)=   "cmdRemoveColumn"
      Tab(2).Control(8)=   "Label2(4)"
      Tab(2).Control(9)=   "Label1(2)"
      Tab(2).ControlCount=   10
      TabCaption(3)   =   " "
      TabPicture(3)   =   "frmSQLWizard.frx":0C16
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
      TabPicture(4)   =   "frmSQLWizard.frx":0C32
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
      TabPicture(5)   =   "frmSQLWizard.frx":0C4E
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
      Begin MSComctlLib.ImageCombo cboJColumn2 
         Height          =   330
         Left            =   -74865
         TabIndex        =   16
         ToolTipText     =   "Select the second column in the join."
         Top             =   3420
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483643
         Locked          =   -1  'True
      End
      Begin MSComctlLib.ImageCombo cboCustomColumn 
         Height          =   330
         Left            =   -74865
         TabIndex        =   22
         ToolTipText     =   "Select or Enter a custom column name."
         Top             =   3375
         Width           =   2760
         _ExtentX        =   4868
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483643
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
         Picture         =   "frmSQLWizard.frx":0C6A
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Move the selected Column up the list"
         Top             =   495
         Width           =   435
      End
      Begin VB.CommandButton cmdSortColDown 
         Height          =   540
         Left            =   -68610
         Picture         =   "frmSQLWizard.frx":10AC
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
      Begin MSComctlLib.ListView lstAllSortCols 
         Height          =   3180
         Left            =   -74865
         TabIndex        =   36
         ToolTipText     =   "Lists the available columns."
         Top             =   495
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   5609
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483630
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   35278
         EndProperty
      End
      Begin MSComctlLib.ListView lstIncSortCols 
         Height          =   3180
         Left            =   -71265
         TabIndex        =   40
         ToolTipText     =   "Lists the selected selected sort columns."
         Top             =   495
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   5609
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483630
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   35278
         EndProperty
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
      Begin MSComctlLib.ImageCombo cboBoolean 
         Height          =   330
         Left            =   -74865
         TabIndex        =   29
         ToolTipText     =   "Select a boolean operator."
         Top             =   990
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Enabled         =   0   'False
         Locked          =   -1  'True
      End
      Begin MSComctlLib.ImageCombo cboOperator 
         Height          =   330
         Left            =   -74865
         TabIndex        =   31
         ToolTipText     =   "Select an Operator to use."
         Top             =   2340
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483643
         Locked          =   -1  'True
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
      Begin MSComctlLib.ListView lstCriteria 
         Height          =   3375
         Left            =   -71490
         TabIndex        =   35
         ToolTipText     =   "Lists the query's selection criteria.."
         Top             =   270
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   5953
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483630
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   35278
         EndProperty
      End
      Begin MSComctlLib.ImageCombo cboWhereCols 
         Height          =   330
         Left            =   -74865
         TabIndex        =   30
         ToolTipText     =   "Select a column to include in the 'WHERE' clause."
         Top             =   1665
         Width           =   2760
         _ExtentX        =   4868
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483643
         Locked          =   -1  'True
      End
      Begin VB.CommandButton cmdColumnDown 
         Height          =   540
         Left            =   -68610
         Picture         =   "frmSQLWizard.frx":14EE
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Move the selected Column down the list"
         Top             =   3150
         Width           =   435
      End
      Begin VB.CommandButton cmdColumnUp 
         Height          =   540
         Left            =   -68610
         Picture         =   "frmSQLWizard.frx":1930
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Move the selected Column up the list"
         Top             =   495
         Width           =   435
      End
      Begin MSComctlLib.ListView lstAllColumns 
         Height          =   2595
         Left            =   -74865
         TabIndex        =   21
         ToolTipText     =   "Lists the columns available for inclusion in the query."
         Top             =   495
         Width           =   2760
         _ExtentX        =   4868
         _ExtentY        =   4577
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483630
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   35278
         EndProperty
      End
      Begin MSComctlLib.ListView lstIncColumns 
         Height          =   3180
         Left            =   -71535
         TabIndex        =   26
         ToolTipText     =   "Lists the columns to be included in the query."
         Top             =   495
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   5609
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483630
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   35278
         EndProperty
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
      Begin MSComctlLib.ImageCombo cboJColumn1 
         Height          =   330
         Left            =   -74865
         TabIndex        =   5
         ToolTipText     =   "Select the first column in the join."
         Top             =   675
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483643
         Locked          =   -1  'True
      End
      Begin MSComctlLib.ListView lstJoins 
         Height          =   2985
         Left            =   -71535
         TabIndex        =   20
         ToolTipText     =   "Lists the selected joins."
         Top             =   765
         Width           =   3390
         _ExtentX        =   5980
         _ExtentY        =   5265
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483630
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   35279
         EndProperty
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
      Begin MSComctlLib.ListView lstIncTables 
         Height          =   3180
         Left            =   3780
         TabIndex        =   3
         ToolTipText     =   "Lists the selected tables."
         Top             =   495
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   5609
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483630
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   35279
         EndProperty
      End
      Begin MSComctlLib.ListView lstAllTables 
         Height          =   3180
         Left            =   135
         TabIndex        =   0
         ToolTipText     =   "Lists the available tables."
         Top             =   495
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   5609
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483630
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   35279
         EndProperty
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
      Picture         =   "frmSQLWizard.frx":1D72
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
' Copyright (C) 2001, 2002, The pgAdmin Development Team
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

Dim objNamespace As pgNamespace
Dim objTable As pgTable
Dim objItem As ListItem

  StartMsg "Getting Tables..."
  lstAllTables.ListItems.Clear
  lstIncTables.ListItems.Clear
  
  For Each objNamespace In frmMain.svr.Databases(szDatabase).Namespaces
    If (Not objNamespace.SystemObject) Or (objNamespace.Name = "public") Then
      For Each objTable In objNamespace.Tables
        If Not objTable.SystemObject Then
          Set objItem = lstAllTables.ListItems.Add(, , objTable.FormattedID)
          Set objItem.Tag = objTable
        End If
      Next objTable
    End If
  Next objNamespace

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
Dim objItem As ComboItem

  StartMsg "Getting Columns..."
  cboJColumn1.ComboItems.Clear
  cboJColumn2.ComboItems.Clear
  lstJoins.ListItems.Clear
  txtPrimaryTable.Text = ""
  
  For X = 1 To lstIncTables.ListItems.Count
    For Each objColumn In frmMain.svr.Databases(szDatabase).Namespaces(lstIncTables.ListItems(X).Tag.Namespace).Tables(lstIncTables.ListItems(X).Tag.Name).Columns
      If Not objColumn.SystemObject Then
        Set objItem = cboJColumn1.ComboItems.Add(, , lstIncTables.ListItems(X).Tag.FormattedID & "." & objColumn.FormattedID)
        Set objItem.Tag = objColumn
      End If
      If Not objColumn.SystemObject Then
        Set objItem = cboJColumn2.ComboItems.Add(, , lstIncTables.ListItems(X).Tag.FormattedID & "." & objColumn.FormattedID)
        Set objItem.Tag = objColumn
      End If
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
Dim bInQuotes As Boolean
Dim szTemp As String
Dim szSchema As String
Dim objItem As ComboItem

  StartMsg "Getting Columns..."
  
  'Clear down
  cboJColumn1.ComboItems.Clear
  cboJColumn2.ComboItems.Clear
  
  'Split the table & schema name
  szTable = txtPrimaryTable.Text
  bInQuotes = False
  For Y = 1 To Len(szTable)
    szTemp = Mid(szTable, Y, 1)
    If szTemp = QUOTE Then
      bInQuotes = Not bInQuotes
    ElseIf szTemp = "." And Not bInQuotes Then
      szSchema = Mid(szTable, 1, Y - 1)
      szTable = Mid(szTable, Y + 1)
    End If
  Next Y
  If szSchema = "" Then szSchema = "public"
    
  'Add columns from the primary table to list1
  For Each objColumn In frmMain.svr.Databases(szDatabase).Namespaces(szSchema).Tables(szTable).Columns
    If Not objColumn.SystemObject Then
      Set objItem = cboJColumn1.ComboItems.Add(, , txtPrimaryTable.Text & "." & objColumn.FormattedID)
      Set objItem.Tag = objColumn
    End If
  Next objColumn
  
  'Add columns from other tables to list1
  For X = 1 To lstJoins.ListItems.Count
    For Each objColumn In frmMain.svr.Databases(szDatabase).Namespaces(lstJoins.ListItems(X).Tag.Namespace).Tables(lstJoins.ListItems(X).Tag.Table).Columns
      If Not objColumn.SystemObject Then
        If ctx.dbVer >= 7.3 Then
          Set objItem = cboJColumn1.ComboItems.Add(, , fmtID(lstJoins.ListItems(X).Tag.Namespace) & "." & fmtID(lstJoins.ListItems(X).Tag.Table) & "." & objColumn.FormattedID)
        Else
          Set objItem = cboJColumn1.ComboItems.Add(, , fmtID(lstJoins.ListItems(X).Tag.Table) & "." & objColumn.FormattedID)
        End If
        Set objItem.Tag = objColumn
      End If
    Next objColumn
  Next
  
  'Add all columns to list2. Previously we only added those that weren't in list1
  'but that prevented multiple links.
  For X = 1 To lstIncTables.ListItems.Count
    For Each objColumn In frmMain.svr.Databases(szDatabase).Namespaces(lstIncTables.ListItems(X).Tag.Namespace).Tables(lstIncTables.ListItems(X).Tag.Name).Columns
      If Not objColumn.SystemObject Then
        Set objItem = cboJColumn2.ComboItems.Add(, , lstIncTables.ListItems(X).Tag.FormattedID & "." & objColumn.FormattedID)
        Set objItem.Tag = objColumn
      End If
    Next objColumn
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
Dim Y As Integer
Dim objColumn As pgColumn
Dim bInQuotes As Boolean
Dim szTemp As String
Dim szTable As String
Dim szSchema As String
Dim objItem As ListItem

  StartMsg "Getting Columns..."
  lstAllColumns.ListItems.Clear
  lstIncColumns.ListItems.Clear
  For X = 1 To lstIncTables.ListItems.Count
  
    'Split the table & schema name
    szTable = lstIncTables.ListItems(X).Text
    bInQuotes = False
    For Y = 1 To Len(szTable)
      szTemp = Mid(szTable, Y, 1)
      If szTemp = QUOTE Then
        bInQuotes = Not bInQuotes
      ElseIf szTemp = "." And Not bInQuotes Then
        szSchema = Mid(szTable, 1, Y - 1)
        szTable = Mid(szTable, Y + 1)
      End If
    Next Y
    If szSchema = "" Then szSchema = "public"
  
    'Add the columns
    For Each objColumn In frmMain.svr.Databases(szDatabase).Namespaces(szSchema).Tables(szTable).Columns
      If Not objColumn.SystemObject Then
        Set objItem = lstAllColumns.ListItems.Add(, , lstIncTables.ListItems(X).Tag.FormattedID & "." & objColumn.FormattedID)
        Set objItem.Tag = lstIncTables.ListItems(X).Tag
      End If
    Next objColumn
  Next X
  
  'Load some default functions 'n' tings
  cboCustomColumn.ComboItems.Add , , "count(*)"
  cboCustomColumn.ComboItems.Add , , "version()"
  cboCustomColumn.ComboItems.Add , , "current_timestamp"
  cboCustomColumn.ComboItems.Add , , "current_user"
  
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
  cboWhereCols.ComboItems.Clear
  lstCriteria.ListItems.Clear
  cboOperator.ComboItems.Clear
  cboBoolean.ComboItems.Clear
  
  For X = 1 To lstIncTables.ListItems.Count
    For Each objColumn In frmMain.svr.Databases(szDatabase).Namespaces(lstIncTables.ListItems(X).Tag.Namespace).Tables(lstIncTables.ListItems(X).Tag.Name).Columns
      If Not objColumn.SystemObject Then cboWhereCols.ComboItems.Add , , lstIncTables.ListItems(X).Tag.FormattedID & "." & objColumn.FormattedID
    Next objColumn
  Next X
  
  'Add some operators etc.
  cboOperator.ComboItems.Add , , "="
  cboOperator.ComboItems.Add , , "!="
  cboOperator.ComboItems.Add , , ">"
  cboOperator.ComboItems.Add , , ">="
  cboOperator.ComboItems.Add , , "<"
  cboOperator.ComboItems.Add , , "<="
  cboOperator.ComboItems.Add , , "LIKE"
  cboOperator.ComboItems.Add , , "NOT LIKE"
  cboOperator.ComboItems.Add , , "IS NULL"
  cboOperator.ComboItems.Add , , "IS NOT NULL"
  cboBoolean.ComboItems.Add , , "AND"
  cboBoolean.ComboItems.Add , , "OR"

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
  lstAllSortCols.ListItems.Clear
  lstIncSortCols.ListItems.Clear
  
  For X = 1 To lstIncTables.ListItems.Count
    For Each objColumn In frmMain.svr.Databases(szDatabase).Namespaces(lstIncTables.ListItems(X).Tag.Namespace).Tables(lstIncTables.ListItems(X).Tag.Name).Columns
      If Not objColumn.SystemObject Then lstAllSortCols.ListItems.Add , , lstIncTables.ListItems(X).Tag.FormattedID & "." & objColumn.FormattedID
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

  If lstAllSortCols.SelectedItem Is Nothing Then
    MsgBox "You must select a column to add!", vbExclamation, "Error"
    Exit Sub
  End If
  lstIncSortCols.ListItems.Add , , lstAllSortCols.SelectedItem.Text & " ASC"
  lstAllSortCols.ListItems.Remove lstAllSortCols.SelectedItem.Index
  If lstAllSortCols.SelectedItem Is Nothing Then
    lstAllSortCols.ToolTipText = ""
  Else
    lstAllSortCols.ToolTipText = lstAllSortCols.SelectedItem.Text
  End If
  If lstIncSortCols.SelectedItem Is Nothing Then
    lstIncSortCols.ToolTipText = ""
  Else
    lstIncSortCols.ToolTipText = lstIncSortCols.SelectedItem.Text
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.cmdAddAsc_Click"
End Sub

Private Sub cmdAddDesc_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.cmdAddDesc_Click()", etFullDebug

  If lstAllSortCols.SelectedItem Is Nothing Then
    MsgBox "You must select a column to add!", vbExclamation, "Error"
    Exit Sub
  End If
  lstIncSortCols.ListItems.Add , , lstAllSortCols.SelectedItem.Text & " DESC"
  lstAllSortCols.ListItems.Remove lstAllSortCols.SelectedItem.Index
  If lstAllSortCols.SelectedItem Is Nothing Then
    lstAllSortCols.ToolTipText = ""
  Else
    lstAllSortCols.ToolTipText = lstAllSortCols.SelectedItem.Text
  End If
  If lstIncSortCols.SelectedItem Is Nothing Then
    lstIncSortCols.ToolTipText = ""
  Else
    lstIncSortCols.ToolTipText = lstIncSortCols.SelectedItem.Text
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.cmdAddDesc_Click"
End Sub

Private Sub cmdAddColumn_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.cmdAddColumn_Click()", etFullDebug

  If lstAllColumns.SelectedItem Is Nothing Then
    MsgBox "You must select a column to add!", vbExclamation, "Error"
    Exit Sub
  End If
  lstIncColumns.ListItems.Add , , lstAllColumns.SelectedItem.Text
  lstAllColumns.ListItems.Remove lstAllColumns.SelectedItem.Index
  If Not (lstAllColumns.SelectedItem Is Nothing) Then
    lstAllColumns.ToolTipText = lstAllColumns.SelectedItem.Text
  Else
    lstAllColumns.ToolTipText = ""
  End If
  If Not (lstIncColumns.SelectedItem Is Nothing) Then
    lstIncColumns.ToolTipText = lstIncColumns.SelectedItem.Text
  Else
    lstIncColumns.ToolTipText = ""
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.cmdAddColumn_Click"
End Sub

Private Sub cmdAddCriteria_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.cmdAddCriteria_Click()", etFullDebug

  If cboBoolean.Enabled = True And cboBoolean.SelectedItem Is Nothing Then
    MsgBox "You must select a boolean operator!", vbExclamation, "Error"
    Exit Sub
  End If
  If cboWhereCols.SelectedItem Is Nothing Then
    MsgBox "You must select a column!", vbExclamation, "Error"
    Exit Sub
  End If
  If cboOperator.SelectedItem Is Nothing Then
    MsgBox "You must select an operator!", vbExclamation, "Error"
    Exit Sub
  End If
  If txtValue.Text = "" And cboOperator.Text <> "IS NULL" And cboOperator.Text <> "IS NOT NULL" Then
    MsgBox "You must enter a value for the criteria!", vbExclamation, "Error"
    Exit Sub
  End If
  If cboBoolean.Enabled = False Then
    If cboOperator.Text = "IS NULL" Or cboOperator.Text = "IS NOT NULL" Then
      lstCriteria.ListItems.Add , , cboWhereCols.Text & " " & cboOperator.Text
    Else
      lstCriteria.ListItems.Add , , cboWhereCols.Text & " " & cboOperator.Text & " " & txtValue.Text
    End If
  Else
    If cboOperator.Text = "IS NULL" Or cboOperator.Text = "IS NOT NULL" Then
      lstCriteria.ListItems.Add , , cboBoolean.Text & " " & cboWhereCols.Text & " " & cboOperator.Text
    Else
      lstCriteria.ListItems.Add , , cboBoolean.Text & " " & cboWhereCols.Text & " " & cboOperator.Text & " " & txtValue.Text
    End If
  End If
  lblBoolean.Enabled = True
  cboBoolean.Enabled = True
  cboBoolean.BackColor = &H80000005
  If lstCriteria.SelectedItem Is Nothing Then
    lstCriteria.ToolTipText = ""
  Else
    lstCriteria.ToolTipText = lstCriteria.SelectedItem.Text
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.cmdAddCriteria_Click"
End Sub

Private Sub cmdAddCustomColumn_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.cmdAddCustomColumn_Click()", etFullDebug

  If cboCustomColumn.Text = "" Then Exit Sub
  lstIncColumns.ListItems.Add , , cboCustomColumn.Text
  If Not (lstAllColumns.SelectedItem Is Nothing) Then
    lstAllColumns.ToolTipText = lstAllColumns.SelectedItem.Text
  Else
    lstAllColumns.ToolTipText = ""
  End If
  If Not (lstIncColumns.SelectedItem Is Nothing) Then
    lstIncColumns.ToolTipText = lstIncColumns.SelectedItem.Text
  Else
    lstIncColumns.ToolTipText = ""
  End If
  
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
Dim objItem As ListItem
  
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
  If ctx.dbVer >= 7.3 Then
    szTable1 = fmtID(cboJColumn1.SelectedItem.Tag.Namespace) & "." & fmtID(cboJColumn1.SelectedItem.Tag.Table)
    szTable2 = fmtID(cboJColumn2.SelectedItem.Tag.Namespace) & "." & fmtID(cboJColumn2.SelectedItem.Tag.Table)
  Else
    szTable1 = fmtID(cboJColumn1.SelectedItem.Tag.Table)
    szTable2 = fmtID(cboJColumn2.SelectedItem.Tag.Table)
  End If

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
  Set objItem = lstJoins.ListItems.Add(, , szType & " JOIN " & szTable2 & " ON " & cboJColumn1.Text & " " & szOperator & " " & cboJColumn2.Text)
  Set objItem.Tag = cboJColumn2.SelectedItem.Tag
  lstJoins.ToolTipText = lstJoins.SelectedItem.Text
  Get_ValidJoinCols
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.cmdAddJoin_Click"
End Sub

Private Sub cmdAddTable_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.cmdAddTable_Click()", etFullDebug

Dim iItem As Integer
Dim objItem As ListItem

  If lstAllTables.SelectedItem Is Nothing Then
    MsgBox "You must select a table to add!", vbExclamation, "Error"
    Exit Sub
  End If
  Set objItem = lstIncTables.ListItems.Add(, , lstAllTables.SelectedItem.Text)
  Set objItem.Tag = lstAllTables.SelectedItem.Tag
  lstAllTables.ListItems.Remove lstAllTables.SelectedItem.Index
  If Not (lstAllTables.SelectedItem Is Nothing) Then
    lstAllTables.ToolTipText = lstAllTables.SelectedItem.Text
  Else
    lstAllTables.ToolTipText = ""
  End If
  If Not (lstIncTables.SelectedItem Is Nothing) Then
    lstIncTables.ToolTipText = lstIncTables.SelectedItem.Text
  Else
    lstIncTables.ToolTipText = ""
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.cmdAddTable_Click"
End Sub

Private Sub cmdColumnDown_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.cmdColumnDown_Click()", etFullDebug

Dim szTemp As String

  If lstIncColumns.SelectedItem Is Nothing Then
    MsgBox "You must select a column to move!", vbExclamation, "Error"
    Exit Sub
  End If
  If lstIncColumns.SelectedItem.Index = lstIncColumns.ListItems.Count Then
    MsgBox "This column is already at the bottom!", vbExclamation, "Error"
    Exit Sub
  End If
  szTemp = lstIncColumns.ListItems(lstIncColumns.SelectedItem.Index + 1).Text
  lstIncColumns.ListItems(lstIncColumns.SelectedItem.Index + 1).Text = lstIncColumns.ListItems(lstIncColumns.SelectedItem.Index).Text
  lstIncColumns.ListItems(lstIncColumns.SelectedItem.Index).Text = szTemp
  lstIncColumns.ListItems(lstIncColumns.SelectedItem.Index + 1).Selected = True
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.cmdColumnDown_Click"
End Sub

Private Sub cmdColumnUp_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.cmdColumnUp_Click()", etFullDebug

Dim szTemp As String

  If lstIncColumns.SelectedItem Is Nothing Then
    MsgBox "You must select a column to move!", vbExclamation, "Error"
    Exit Sub
  End If
  If lstIncColumns.SelectedItem.Index = 1 Then
    MsgBox "This column is already at the top!", vbExclamation, "Error"
    Exit Sub
  End If
  szTemp = lstIncColumns.ListItems(lstIncColumns.SelectedItem.Index - 1).Text
  lstIncColumns.ListItems(lstIncColumns.SelectedItem.Index - 1).Text = lstIncColumns.ListItems(lstIncColumns.SelectedItem.Index).Text
  lstIncColumns.ListItems(lstIncColumns.SelectedItem.Index).Text = szTemp
  lstIncColumns.ListItems(lstIncColumns.SelectedItem.Index - 1).Selected = True
  
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
  For X = 1 To lstIncColumns.ListItems.Count
    szColumns = szColumns & "  " & lstIncColumns.ListItems(X).Text & ", " & vbCrLf
  Next
  If Len(szColumns) > 4 Then szColumns = Mid(szColumns, 1, Len(szColumns) - 4)
  
  'Joins
  If lstJoins.ListItems.Count >= 1 Then szJoins = "  " & txtPrimaryTable.Text & vbCrLf
  For X = 1 To lstJoins.ListItems.Count
    szJoins = szJoins & "  " & lstJoins.ListItems(X).Text & vbCrLf
  Next
  
  'Only add tables that aren't in any joins
  For X = 1 To lstIncTables.ListItems.Count
    If InStr(1, szJoins, lstIncTables.ListItems(X).Text) = 0 Then
      szTables = szTables & "  " & lstIncTables.ListItems(X).Text & ", " & vbCrLf
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
  For X = 1 To lstCriteria.ListItems.Count
    szCriteria = szCriteria & "  " & lstCriteria.ListItems(X).Text & " " & vbCrLf
  Next
  
  'Sorting
  For X = 1 To lstIncSortCols.ListItems.Count
    szOrderBy = szOrderBy & "  " & lstIncSortCols.ListItems(X).Text & ", " & vbCrLf
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

  If lstIncColumns.SelectedItem Is Nothing Then
    MsgBox "You must select a column to remove!", vbExclamation, "Error"
    Exit Sub
  End If
  lstAllColumns.ListItems.Add , , lstIncColumns.SelectedItem.Text
  lstIncColumns.ListItems.Remove lstIncColumns.SelectedItem.Index
  If Not (lstAllColumns.SelectedItem Is Nothing) Then
    lstAllColumns.ToolTipText = lstAllColumns.SelectedItem.Text
  Else
    lstAllColumns.ToolTipText = ""
  End If
  If Not (lstIncColumns.SelectedItem Is Nothing) Then
    lstIncColumns.ToolTipText = lstIncColumns.SelectedItem.Text
  Else
    lstIncColumns.ToolTipText = ""
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.cmdRemoveColumn_Click"
End Sub

Private Sub cmdRemoveCriteria_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.cmdRemoveCriteria_Click()", etFullDebug

  If lstCriteria.SelectedItem Is Nothing Then
    MsgBox "You must select a join to remove!", vbExclamation, "Error"
    Exit Sub
  End If
  If lstCriteria.ListItems.Count > 1 And lstCriteria.SelectedItem.Index = 1 Then
    MsgBox "You must remove all other criteria before you can remove the first!", vbExclamation, "Error"
    Exit Sub
  End If
  If lstCriteria.ListItems.Count = 1 And lstCriteria.SelectedItem.Index = 1 Then
    cboBoolean.Enabled = False
    lblBoolean.Enabled = False
    cboBoolean.BackColor = &H8000000F
  End If
  lstCriteria.ListItems.Remove lstCriteria.SelectedItem.Index
  If lstCriteria.SelectedItem Is Nothing Then
    lstCriteria.ToolTipText = ""
  Else
    lstCriteria.ToolTipText = lstCriteria.SelectedItem.Text
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.cmdRemoveCriteria_Click"
End Sub

Private Sub cmdRemoveJoin_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.cmdRemoveJoin_Click()", etFullDebug

  If lstJoins.SelectedItem Is Nothing Then
    MsgBox "You must select a join to remove!", vbExclamation, "Error"
    Exit Sub
  End If
  lstJoins.ListItems.Remove lstJoins.SelectedItem.Index
  'Set the selected item if there is one, else clear the primary table
  If lstJoins.ListItems.Count > 0 Then
    Get_ValidJoinCols
  Else
    txtPrimaryTable.Text = ""
    Get_JoinCols
  End If
  If lstJoins.SelectedItem Is Nothing Then
    lstJoins.ToolTipText = ""
  Else
    lstJoins.ToolTipText = lstJoins.SelectedItem.Text
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.cmdRemoveJoin_Click"
End Sub

Private Sub cmdRemoveTable_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.cmdRemoveTable_Click()", etFullDebug

Dim iItem As Integer
Dim objItem As ListItem

  If lstIncTables.SelectedItem Is Nothing Then
    MsgBox "You must select a table to remove!", vbExclamation, "Error"
    Exit Sub
  End If
  Set objItem = lstAllTables.ListItems.Add(, , lstIncTables.SelectedItem.Text)
  Set objItem.Tag = lstIncTables.SelectedItem.Tag
  lstIncTables.ListItems.Remove lstIncTables.SelectedItem.Index
  If lstAllTables.SelectedItem Is Nothing Then
    lstAllTables.ToolTipText = ""
  Else
    lstAllTables.ToolTipText = lstAllTables.SelectedItem.Text
  End If
  If lstIncTables.SelectedItem Is Nothing Then
    lstIncTables.ToolTipText = ""
  Else
    lstIncTables.ToolTipText = lstIncTables.SelectedItem.Text
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.cmdRemoveTable_Click"
End Sub

Private Sub cmdNext_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.cmdNext_Click()", etFullDebug

  bButtonPress = True
  Select Case tabWizard.Tab
    Case 0
      If lstIncTables.ListItems.Count = 0 Then Exit Sub
      If lstIncTables.ListItems.Count = 1 Then
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
      If lstIncColumns.ListItems.Count = 0 Then Exit Sub
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
      If lstIncTables.ListItems.Count = 1 Then
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

  If lstIncSortCols.SelectedItem Is Nothing Then
    MsgBox "You must select column to remove!", vbExclamation, "Error"
    Exit Sub
  End If
  If Mid(lstIncSortCols.SelectedItem.Text, Len(lstIncSortCols.SelectedItem.Text) - 3, 4) = "DESC" Then
    lstAllSortCols.ListItems.Add , , Mid(lstIncSortCols.SelectedItem.Text, 1, Len(lstIncSortCols.SelectedItem.Text) - 5)
  Else
    lstAllSortCols.ListItems.Add , , Mid(lstIncSortCols.SelectedItem.Text, 1, Len(lstIncSortCols.SelectedItem.Text) - 4)
  End If
  lstIncSortCols.ListItems.Remove lstIncSortCols.SelectedItem.Index
  If lstAllSortCols.SelectedItem Is Nothing Then
    lstAllSortCols.ToolTipText = ""
  Else
    lstAllSortCols.ToolTipText = lstAllSortCols.SelectedItem.Text
  End If
  If lstIncSortCols.SelectedItem Is Nothing Then
    lstIncSortCols.ToolTipText = ""
  Else
    lstIncSortCols.ToolTipText = lstIncSortCols.SelectedItem.Text
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.cmdRemoveSortCol_Click"
End Sub

Private Sub cmdSortColDown_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.cmdSortColDown_Click()", etFullDebug

Dim szTemp As String

  If lstIncSortCols.SelectedItem Is Nothing Then
    MsgBox "You must select a column to move!", vbExclamation, "Error"
    Exit Sub
  End If
  If lstIncSortCols.SelectedItem.Index = lstIncSortCols.ListItems.Count Then
    MsgBox "This column is already at the bottom!", vbExclamation, "Error"
    Exit Sub
  End If
  szTemp = lstIncSortCols.ListItems(lstIncSortCols.SelectedItem.Index + 1).Text
  lstIncSortCols.ListItems(lstIncSortCols.SelectedItem.Index + 1).Text = lstIncSortCols.ListItems(lstIncSortCols.SelectedItem.Index).Text
  lstIncSortCols.ListItems(lstIncSortCols.SelectedItem.Index).Text = szTemp
  lstIncSortCols.ListItems(lstIncSortCols.SelectedItem.Index + 1).Selected = True
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.cmdSortColDown_Click"
End Sub

Private Sub cmdSortColUp_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.cmdSortColUp_Click()", etFullDebug

Dim szTemp As String

  If lstIncSortCols.SelectedItem Is Nothing Then
    MsgBox "You must select a column to move!", vbExclamation, "Error"
    Exit Sub
  End If
  If lstIncSortCols.SelectedItem.Index = 1 Then
    MsgBox "This column is already at the top!", vbExclamation, "Error"
    Exit Sub
  End If
  szTemp = lstIncSortCols.ListItems(lstIncSortCols.SelectedItem.Index - 1).Text
  lstIncSortCols.ListItems(lstIncSortCols.SelectedItem.Index - 1).Text = lstIncSortCols.ListItems(lstIncSortCols.SelectedItem.Index).Text
  lstIncSortCols.ListItems(lstIncSortCols.SelectedItem.Index).Text = szTemp
  lstIncSortCols.ListItems(lstIncSortCols.SelectedItem.Index - 1).Selected = True
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.cmdSortColUp_Click"
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.Form_Load()", etFullDebug

  'Set the font
  Set cboBoolean.Font = ctx.Font
  Set cboCustomColumn.Font = ctx.Font
  Set cboJColumn1.Font = ctx.Font
  Set cboJColumn2.Font = ctx.Font
  Set cboOperator.Font = ctx.Font
  Set cboWhereCols.Font = ctx.Font
  Set lstAllColumns.Font = ctx.Font
  Set lstAllSortCols.Font = ctx.Font
  Set lstAllTables.Font = ctx.Font
  Set lstCriteria.Font = ctx.Font
  Set lstIncColumns.Font = ctx.Font
  Set lstIncSortCols.Font = ctx.Font
  Set lstIncTables.Font = ctx.Font
  Set lstJoins.Font = ctx.Font
  Set txtLimit.Font = ctx.Font
  Set txtOffset.Font = ctx.Font
  Set txtPrimaryTable.Font = ctx.Font
  Set txtValue.Font = ctx.Font
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.Form_Load"
End Sub

Private Sub lstAllColumns_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.lstAllColumns_Click()", etFullDebug

  If lstAllColumns.SelectedItem Is Nothing Then
    lstAllColumns.ToolTipText = ""
  Else
    lstAllColumns.ToolTipText = lstAllColumns.SelectedItem.Text
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.lstAllColumns_Click"
End Sub

Private Sub lstAllSortCols_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.lstAllSortCols_Click()", etFullDebug

  If lstAllSortCols.SelectedItem Is Nothing Then
    lstAllSortCols.ToolTipText = ""
  Else
    lstAllSortCols.ToolTipText = lstAllSortCols.SelectedItem.Text
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.lstAllSortCols_Click"
End Sub

Private Sub lstAllTables_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.lstAllTables_Click()", etFullDebug

  If lstAllTables.SelectedItem Is Nothing Then
    lstAllTables.ToolTipText = ""
  Else
    lstAllTables.ToolTipText = lstAllTables.SelectedItem.Text
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.lstAllTables_Click"
End Sub

Private Sub lstCriteria_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.lstCriteria_Click()", etFullDebug

  If lstCriteria.SelectedItem Is Nothing Then
    lstCriteria.ToolTipText = ""
  Else
    lstCriteria.ToolTipText = lstCriteria.SelectedItem.Text
  End If
  
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

  If lstIncColumns.SelectedItem Is Nothing Then
    lstIncColumns.ToolTipText = ""
  Else
    lstIncColumns.ToolTipText = lstIncColumns.SelectedItem.Text
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.lstIncColumns_Click"
End Sub

Private Sub lstIncSortCols_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.lstIncSortCols_Click()", etFullDebug

  If lstIncSortCols.SelectedItem Is Nothing Then
    lstIncSortCols.ToolTipText = ""
  Else
    lstIncSortCols.ToolTipText = lstIncSortCols.SelectedItem.Text
  End If
  
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

  If lstIncTables.SelectedItem Is Nothing Then
    lstIncTables.ToolTipText = ""
  Else
    lstIncTables.ToolTipText = lstIncTables.SelectedItem.Text
  End If
  
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

  If lstJoins.SelectedItem Is Nothing Then
    lstJoins.ToolTipText = ""
  Else
    lstJoins.ToolTipText = lstJoins.SelectedItem.Text
  End If
  
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

  szDatabase = szDB
  tabWizard.Tab = 0
  
  'Can only do OJ's on PostgreSQL 7.1+
  If ctx.dbVer >= 7.1 Then
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


