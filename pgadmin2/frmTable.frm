VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{44F33AC4-8757-4330-B063-18608617F23E}#12.4#0"; "HighlightBox.ocx"
Begin VB.Form frmTable 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Table"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   Icon            =   "frmTable.frx":0000
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
      TabIndex        =   6
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4410
      TabIndex        =   7
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
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "&Properties"
      TabPicture(0)   =   "frmTable.frx":014A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblProperties(2)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblProperties(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblProperties(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblProperties(4)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "hbxProperties(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtProperties(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtProperties(1)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtProperties(2)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cboProperties(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "chkProperties(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "&Columns"
      TabPicture(1)   =   "frmTable.frx":0166
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lvProperties(0)"
      Tab(1).Control(1)=   "cmdColAdd"
      Tab(1).Control(2)=   "cmdColRemove"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "C&hecks"
      TabPicture(2)   =   "frmTable.frx":0182
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtCheck(0)"
      Tab(2).Control(1)=   "cmdChkAdd"
      Tab(2).Control(2)=   "cmdChkRemove"
      Tab(2).Control(3)=   "hbxCheck(0)"
      Tab(2).Control(4)=   "lvProperties(1)"
      Tab(2).Control(5)=   "lblProperties(5)"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "&Foreign Keys"
      TabPicture(3)   =   "frmTable.frx":019E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdFkyRemove"
      Tab(3).Control(1)=   "cmdFkyAdd"
      Tab(3).Control(2)=   "lvProperties(2)"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "&Inherits"
      TabPicture(4)   =   "frmTable.frx":01BA
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "cboInheritedTables(0)"
      Tab(4).Control(1)=   "cmdInhRemove"
      Tab(4).Control(2)=   "cmdInhAdd"
      Tab(4).Control(3)=   "lvProperties(3)"
      Tab(4).Control(4)=   "lblProperties(6)"
      Tab(4).ControlCount=   5
      TabCaption(5)   =   "&Security"
      TabPicture(5)   =   "frmTable.frx":01D6
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "lvProperties(4)"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "cmdAdd"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "fraAdd"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "cmdRemove"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).ControlCount=   4
      Begin VB.CheckBox chkProperties 
         Alignment       =   1  'Right Justify
         Caption         =   "OIDs?"
         Height          =   240
         Index           =   0
         Left            =   135
         TabIndex        =   4
         ToolTipText     =   "Does this table have an OID column? (Prior to PostgreSQL 7.2, there is always an OID column)."
         Top             =   2340
         Value           =   1  'Checked
         Width           =   1995
      End
      Begin MSComctlLib.ImageCombo cboProperties 
         Height          =   330
         Index           =   0
         Left            =   1935
         TabIndex        =   43
         ToolTipText     =   "The tables owner."
         Top             =   1440
         Width           =   3390
         _ExtentX        =   5980
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Locked          =   -1  'True
         ImageList       =   "il"
      End
      Begin MSComctlLib.ImageCombo cboInheritedTables 
         Height          =   330
         Index           =   0
         Left            =   -73065
         TabIndex        =   22
         Top             =   5850
         Width           =   3390
         _ExtentX        =   5980
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483633
         Locked          =   -1  'True
         ImageList       =   "il"
      End
      Begin VB.CommandButton cmdFkyRemove 
         Caption         =   "&Remove"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -73560
         TabIndex        =   18
         ToolTipText     =   "Remove the selected foreign key."
         Top             =   5805
         Width           =   1230
      End
      Begin VB.CommandButton cmdFkyAdd 
         Caption         =   "&Add"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -74865
         TabIndex        =   17
         ToolTipText     =   "Add the defined foreign key."
         Top             =   5805
         Width           =   1230
      End
      Begin VB.CommandButton cmdInhRemove 
         Caption         =   "&Remove"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -73560
         TabIndex        =   21
         Top             =   5355
         Width           =   1230
      End
      Begin VB.CommandButton cmdInhAdd 
         Caption         =   "&Add"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -74865
         TabIndex        =   20
         Top             =   5355
         Width           =   1230
      End
      Begin VB.TextBox txtCheck 
         Height          =   285
         Index           =   0
         Left            =   -73065
         TabIndex        =   14
         ToolTipText     =   "Enter a name for the check."
         Top             =   4950
         Width           =   3390
      End
      Begin VB.CommandButton cmdChkAdd 
         Caption         =   "&Add"
         Height          =   375
         Left            =   -74865
         TabIndex        =   12
         ToolTipText     =   "Add the defined check."
         Top             =   4410
         Width           =   1230
      End
      Begin VB.CommandButton cmdChkRemove 
         Caption         =   "&Remove"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -73560
         TabIndex        =   13
         ToolTipText     =   "Remove the selected check."
         Top             =   4410
         Width           =   1230
      End
      Begin VB.CommandButton cmdColRemove 
         Caption         =   "&Remove"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -73560
         TabIndex        =   10
         ToolTipText     =   "Remove the selected column."
         Top             =   5805
         Width           =   1230
      End
      Begin VB.CommandButton cmdColAdd 
         Caption         =   "&Add"
         Height          =   375
         Left            =   -74865
         TabIndex        =   9
         ToolTipText     =   "Add a new column."
         Top             =   5805
         Width           =   1230
      End
      Begin MSComctlLib.ListView lvProperties 
         Height          =   5280
         Index           =   0
         Left            =   -74865
         TabIndex        =   8
         Top             =   450
         Width           =   5190
         _ExtentX        =   9155
         _ExtentY        =   9313
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         Icons           =   "il"
         SmallIcons      =   "il"
         ColHdrIcons     =   "il"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   3087
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Pos"
            Object.Width           =   794
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Type"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Length"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Default"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Not Null"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Primary Key"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Comment"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   2
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   3
         ToolTipText     =   "The number of tuples (rows) in the table."
         Top             =   1890
         Width           =   3390
      End
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   1
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   2
         ToolTipText     =   "The tables OID (Object ID) in the PostgreSQL Database."
         Top             =   1080
         Width           =   3390
      End
      Begin VB.TextBox txtProperties 
         Height          =   285
         Index           =   0
         Left            =   1935
         TabIndex        =   1
         ToolTipText     =   "The name of the table."
         Top             =   675
         Width           =   3390
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "&Remove"
         Height          =   375
         Left            =   -73560
         TabIndex        =   25
         ToolTipText     =   "Remove the selected entry."
         Top             =   3915
         Width           =   1230
      End
      Begin VB.Frame fraAdd 
         Caption         =   "Define Privilege"
         Height          =   1815
         Left            =   -74865
         TabIndex        =   35
         Top             =   4410
         Width           =   5190
         Begin VB.CheckBox chkPrivilege 
            Caption         =   "&Trigger"
            Height          =   195
            Index           =   7
            Left            =   3420
            TabIndex        =   34
            ToolTipText     =   "Give trigger privilege to the selected entity."
            Top             =   1530
            Width           =   1590
         End
         Begin VB.CheckBox chkPrivilege 
            Caption         =   "R&eferences"
            Height          =   195
            Index           =   6
            Left            =   3420
            TabIndex        =   33
            ToolTipText     =   "Give references privilege to the selected entity."
            Top             =   1260
            Width           =   1590
         End
         Begin VB.CheckBox chkPrivilege 
            Caption         =   "&Delete"
            Height          =   195
            Index           =   3
            Left            =   225
            TabIndex        =   30
            ToolTipText     =   "Give delete privilege to the selected entity."
            Top             =   1530
            Width           =   1590
         End
         Begin VB.CheckBox chkPrivilege 
            Caption         =   "&All"
            Height          =   195
            Index           =   0
            Left            =   225
            TabIndex        =   27
            ToolTipText     =   "Give all privileges to the selected entity."
            Top             =   720
            Width           =   1590
         End
         Begin VB.CheckBox chkPrivilege 
            Caption         =   "&Select"
            Height          =   195
            Index           =   1
            Left            =   225
            TabIndex        =   28
            ToolTipText     =   "Give select privilege to the selected entity."
            Top             =   990
            Width           =   1590
         End
         Begin VB.CheckBox chkPrivilege 
            Caption         =   "&Update"
            Height          =   195
            Index           =   2
            Left            =   225
            TabIndex        =   29
            ToolTipText     =   "Give update privilege to the selected entity."
            Top             =   1260
            Width           =   1590
         End
         Begin VB.CheckBox chkPrivilege 
            Caption         =   "&Insert"
            Height          =   195
            Index           =   4
            Left            =   3420
            TabIndex        =   31
            ToolTipText     =   "Give insert privilege to the selected entity."
            Top             =   720
            Width           =   1590
         End
         Begin VB.CheckBox chkPrivilege 
            Caption         =   "&Rule"
            Height          =   195
            Index           =   5
            Left            =   3420
            TabIndex        =   32
            ToolTipText     =   "Give rule privilege to the selected entity."
            Top             =   990
            Width           =   1590
         End
         Begin MSComctlLib.ImageCombo cboEntities 
            Height          =   330
            Left            =   1260
            TabIndex        =   26
            ToolTipText     =   "Select a user, group or 'PUBLIC'."
            Top             =   315
            Width           =   3705
            _ExtentX        =   6535
            _ExtentY        =   582
            _Version        =   393216
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Locked          =   -1  'True
            ImageList       =   "il"
         End
         Begin VB.Label lblProperties 
            AutoSize        =   -1  'True
            Caption         =   "User/Group"
            Height          =   195
            Index           =   3
            Left            =   180
            TabIndex        =   36
            Top             =   360
            Width           =   840
         End
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   375
         Left            =   -74865
         TabIndex        =   24
         ToolTipText     =   "Add the defined entry."
         Top             =   3915
         Width           =   1230
      End
      Begin MSComctlLib.ListView lvProperties 
         Height          =   3390
         Index           =   4
         Left            =   -74865
         TabIndex        =   23
         ToolTipText     =   "The access control list for the view."
         Top             =   450
         Width           =   5190
         _ExtentX        =   9155
         _ExtentY        =   5980
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         Icons           =   "il"
         SmallIcons      =   "il"
         ColHdrIcons     =   "il"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "User/Group name"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Privileges"
            Object.Width           =   4939
         EndProperty
      End
      Begin HighlightBox.HBX hbxProperties 
         Height          =   3480
         Index           =   0
         Left            =   135
         TabIndex        =   5
         ToolTipText     =   "Comments about the table."
         Top             =   2700
         Width           =   5190
         _ExtentX        =   9155
         _ExtentY        =   6138
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Comments"
      End
      Begin MSComctlLib.ListView lvProperties 
         Height          =   5235
         Index           =   2
         Left            =   -74865
         TabIndex        =   16
         Top             =   450
         Width           =   5190
         _ExtentX        =   9155
         _ExtentY        =   9234
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         Icons           =   "il"
         SmallIcons      =   "il"
         ColHdrIcons     =   "il"
         ForeColor       =   -2147483640
         BackColor       =   -2147483633
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "References"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Columns"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Referenced columns"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "On Delete"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "On Update"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Deferrable?"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Initially"
            Object.Width           =   1764
         EndProperty
      End
      Begin MSComctlLib.ListView lvProperties 
         Height          =   4785
         Index           =   3
         Left            =   -74865
         TabIndex        =   19
         Top             =   450
         Width           =   5190
         _ExtentX        =   9155
         _ExtentY        =   8440
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         Icons           =   "il"
         SmallIcons      =   "il"
         ColHdrIcons     =   "il"
         ForeColor       =   -2147483640
         BackColor       =   -2147483633
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Inherited Table Name"
            Object.Width           =   7937
         EndProperty
      End
      Begin HighlightBox.HBX hbxCheck 
         Height          =   870
         Index           =   0
         Left            =   -74865
         TabIndex        =   15
         ToolTipText     =   "The check definition."
         Top             =   5355
         Width           =   5190
         _ExtentX        =   9155
         _ExtentY        =   1535
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Definition"
      End
      Begin MSComctlLib.ListView lvProperties 
         Height          =   3885
         Index           =   1
         Left            =   -74865
         TabIndex        =   11
         Top             =   450
         Width           =   5190
         _ExtentX        =   9155
         _ExtentY        =   6853
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         Icons           =   "il"
         SmallIcons      =   "il"
         ColHdrIcons     =   "il"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   5292
         EndProperty
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Inherited table name"
         Height          =   195
         Index           =   6
         Left            =   -74820
         TabIndex        =   42
         Top             =   5940
         Width           =   1440
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Check name"
         Height          =   195
         Index           =   5
         Left            =   -74820
         TabIndex        =   41
         Top             =   4995
         Width           =   900
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Tuples"
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   40
         Top             =   1935
         Width           =   480
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   39
         Top             =   720
         Width           =   420
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "OID"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   38
         Top             =   1125
         Width           =   285
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Owner"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   37
         Top             =   1530
         Width           =   465
      End
   End
   Begin MSComctlLib.ImageList il 
      Left            =   0
      Top             =   6300
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTable.frx":01F2
            Key             =   "column"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTable.frx":078C
            Key             =   "table"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTable.frx":08E6
            Key             =   "foreignkey"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTable.frx":0E80
            Key             =   "public"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTable.frx":0FDA
            Key             =   "group"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTable.frx":1574
            Key             =   "user"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTable.frx":1B0E
            Key             =   "check"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTable.frx":20A8
            Key             =   "sequence"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001, 2002, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence
'
' frmTable.frm - Edit/Create a Table

Option Explicit

Dim bNew As Boolean
Dim szDatabase As String
Dim szUsers() As String
Dim objTable As pgTable

Private Sub cmdCancel_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmTable.cmdCancel_Click()", etFullDebug

  Unload Me
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmTable.cmdCancel_Click"
End Sub

Private Sub cmdChkAdd_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmTable.cmdChkAdd_Click()", etFullDebug

Dim objItem As ListItem

  If txtCheck(0).Text = "" Then
    MsgBox "You must enter a name for the check!", vbExclamation, "Error"
    tabProperties.Tab = 2
    txtCheck(0).SetFocus
    Exit Sub
  End If
  If hbxCheck(0).Text = "" Then
    MsgBox "You must enter a definition for the check!", vbExclamation, "Error"
    tabProperties.Tab = 2
    hbxCheck(0).SetFocus
    Exit Sub
  End If
  For Each objItem In lvProperties(1).ListItems
    If objItem.Text = txtCheck(0).Text Then
      MsgBox "This check name is already in the list!", vbExclamation, "Error"
      tabProperties.Tab = 2
      txtCheck(0).SetFocus
      Exit Sub
    End If
  Next objItem
  
  Set objItem = lvProperties(1).ListItems.Add(, , txtCheck(0).Text, "check", "check")
  objItem.SubItems(1) = hbxCheck(0).Text
  lvProperties(1).Tag = "Y"
  
  txtCheck(0).Text = ""
  hbxCheck(0).Text = ""
  
  Exit Sub
  
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmTable.cmdChkAdd_Click"
End Sub

Private Sub cmdChkRemove_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmTable.cmdChkRemove_Click()", etFullDebug

  If lvProperties(1).SelectedItem Is Nothing Then
    MsgBox "You must select a check to remove!", vbExclamation, "Error"
    tabProperties.Tab = 2
    lvProperties(1).SetFocus
    Exit Sub
  End If
  
  lvProperties(1).ListItems.Remove lvProperties(1).SelectedItem.Index
  lvProperties(1).Tag = "Y"
  If lvProperties(1).SelectedItem Is Nothing Then
    cmdChkRemove.Enabled = False
  Else
    lvProperties_ItemClick 1, lvProperties(1).SelectedItem
  End If
  
  Exit Sub
  
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmTable.cmdChkRemove_Click"
End Sub

Private Sub cmdColAdd_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmTable.cmdColAdd_Click()", etFullDebug

Dim objColumnForm As New frmColumn
  
  Load objColumnForm
  If objTable Is Nothing Then
    objColumnForm.Initialise ctx.CurrentDB, "TA", , Me, False
  Else
    objColumnForm.Initialise ctx.CurrentDB, "TA", , Me, True
  End If
  objColumnForm.Show

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmTable.cmdColAdd_Click"
End Sub

Private Sub cmdColRemove_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmTable.cmdColRemove_Click()", etFullDebug

  If lvProperties(0).SelectedItem Is Nothing Then
    MsgBox "You must select a column to remove!", vbExclamation, "Error"
    tabProperties.Tab = 1
    lvProperties(0).SetFocus
    Exit Sub
  End If
  
  lvProperties(0).ListItems.Remove lvProperties(0).SelectedItem.Index
  lvProperties(0).Tag = "Y"
  If lvProperties(0).SelectedItem Is Nothing Then
    cmdColRemove.Enabled = False
  Else
    lvProperties_ItemClick 0, lvProperties(0).SelectedItem
  End If
  
  Exit Sub

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmTable.cmdColRemove_Click"
End Sub

Private Sub cmdFkyAdd_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmTable.cmdFkyAdd_Click()", etFullDebug

Dim objForeignKeyForm As New frmForeignKey
  
  Load objForeignKeyForm
  objForeignKeyForm.Initialise ctx.CurrentDB, "TA", , Me
  objForeignKeyForm.Show
  
  Exit Sub
  
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmTable.cmdFkyAdd_Click"
End Sub

Private Sub cmdFkyRemove_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmTable.cmdFkyRemove_Click()", etFullDebug

Dim colTemp As New Collection
Dim vTemp As Variant
Dim objItem As ListItem
Dim szForeignKey As String
Dim szTable As String
    
  If lvProperties(2).SelectedItem Is Nothing Then
    MsgBox "You must select a foreign key to remove!", vbExclamation, "Error"
    tabProperties.Tab = 3
    lvProperties(2).SetFocus
    Exit Sub
  End If
  
  lvProperties(2).ListItems.Remove lvProperties(2).SelectedItem.Index

  Exit Sub
  
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmTable.cmdFkyRemove_Click"
End Sub

Private Sub cmdInhAdd_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmTable.cmdInhAdd_Click()", etFullDebug

Dim objItem As ListItem

  If cboInheritedTables(0).Text = "" Then
    MsgBox "You must select a table to add!", vbExclamation, "Error"
    tabProperties.Tab = 4
    cboInheritedTables(0).SetFocus
    Exit Sub
  End If
  For Each objItem In lvProperties(3).ListItems
    If objItem.Text = cboInheritedTables(0).Text Then
      MsgBox "This table is already in the list!", vbExclamation, "Error"
      tabProperties.Tab = 4
      cboInheritedTables(0).SetFocus
      Exit Sub
    End If
  Next objItem
  
  Set objItem = lvProperties(3).ListItems.Add(, , cboInheritedTables(0).Text, "table", "table")
  lvProperties(3).Tag = "Y"
  
  Exit Sub
  
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmTable.cmdInhAdd_Click"
End Sub

Private Sub cmdInhRemove_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmTable.cmdInhRemove_Click()", etFullDebug

  If lvProperties(3).SelectedItem Is Nothing Then
    MsgBox "You must select a table to remove!", vbExclamation, "Error"
    tabProperties.Tab = 4
    lvProperties(3).SetFocus
    Exit Sub
  End If
  
  lvProperties(3).ListItems.Remove lvProperties(3).SelectedItem.Index
  lvProperties(3).Tag = "Y"

  Exit Sub
  
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmTable.cmdInhRemove_Click"
End Sub

Private Sub cmdOK_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmTable.cmdOK_Click()", etFullDebug

Dim objNode As Node
Dim objItem As ListItem
Dim lACL As Long
Dim szEntity As String
Dim vEntity As Variant
Dim szOldName As String
Dim szDataType As String
Dim szColumns As String
Dim szPrimaryKeys As String
Dim szChecks As String
Dim szForeignKeys As String
Dim szInherits As String

  'Check the data
  If txtProperties(0).Text = "" Then
    MsgBox "You must specify a Table name!", vbExclamation, "Error"
    tabProperties.Tab = 0
    txtProperties(0).SetFocus
    Exit Sub
  End If
  If lvProperties(0).ListItems.Count = 0 Then
    MsgBox "You must define at least one column!", vbExclamation, "Error"
    tabProperties.Tab = 1
    lvProperties(0).SetFocus
    Exit Sub
  End If
  
  If bNew Then
    StartMsg "Creating Table..."
    
    'Build the column list
    For Each objItem In lvProperties(0).ListItems
      szColumns = szColumns & QUOTE & objItem.Text & QUOTE & " " & objItem.SubItems(2)
      If objItem.SubItems(3) <> "" Then szColumns = szColumns & "(" & objItem.SubItems(3) & ")"
      If objItem.SubItems(4) <> "" Then szColumns = szColumns & " DEFAULT " & objItem.SubItems(4)
      If objItem.SubItems(5) <> "No" Then szColumns = szColumns & " NOT NULL"
      szColumns = szColumns & ", "
      
      'Add to the Primary Key list if required.
      If objItem.SubItems(6) <> "No" Then szPrimaryKeys = szPrimaryKeys & QUOTE & objItem.Text & QUOTE & ", "
    Next objItem
    szColumns = Left(szColumns, Len(szColumns) - 2)
    
    'Add the Primary Keys
    If Len(szPrimaryKeys) > 2 Then szPrimaryKeys = Left(szPrimaryKeys, Len(szPrimaryKeys) - 2)

    'Add Checks
    For Each objItem In lvProperties(1).ListItems
      szChecks = szChecks & "CONSTRAINT " & QUOTE & objItem.Text & QUOTE & " CHECK (" & objItem.SubItems(1) & "), "
    Next objItem
    If Len(szChecks) > 2 Then szChecks = Left(szChecks, Len(szChecks) - 2)
    
    'Add Foreign Keys
    For Each objItem In lvProperties(2).ListItems
      szForeignKeys = szForeignKeys & "CONSTRAINT " & QUOTE & objItem.Text & QUOTE & " FOREIGN KEY (" & objItem.SubItems(2) & ") "
      szForeignKeys = szForeignKeys & "REFERENCES " & QUOTE & objItem.SubItems(1) & QUOTE & " (" & objItem.SubItems(3) & ")"
      szForeignKeys = szForeignKeys & " ON DELETE " & UCase(objItem.SubItems(4))
      szForeignKeys = szForeignKeys & " ON UPDATE " & UCase(objItem.SubItems(5))
      If objItem.SubItems(6) = "Yes" Then szForeignKeys = szForeignKeys & " DEFERRABLE"
      szForeignKeys = szForeignKeys & " INITIALLY " & UCase(objItem.SubItems(7)) & ", "
    Next objItem
    If Len(szForeignKeys) > 2 Then szForeignKeys = Left(szForeignKeys, Len(szForeignKeys) - 2)
    
    'Add Inherits
    For Each objItem In lvProperties(3).ListItems
      szInherits = szInherits & QUOTE & objItem.Text & QUOTE & ", "
    Next objItem
    If Len(szInherits) > 2 Then szInherits = Left(szInherits, Len(szInherits) - 2)
     
    frmMain.svr.Databases(szDatabase).Tables.Add txtProperties(0).Text, szColumns, szPrimaryKeys, szChecks, szForeignKeys, szInherits, hbxProperties(0).Text, Bin2Bool(chkProperties(0).Value)
    
    'Add any comments for the columns.
    For Each objItem In lvProperties(0).ListItems
      If objItem.SubItems(7) <> "" Then frmMain.svr.Databases(szDatabase).Tables(txtProperties(0).Text).Columns(objItem.Text).Comment = objItem.SubItems(7)
    Next objItem
    
    'Add a new node and update the text on the parent
    For Each objNode In frmMain.tv.Nodes
      If Left(objNode.Key, 4) <> "SVR-" Then
        If (Left(objNode.Key, 4) = "TBL+") And (objNode.Parent.Text = szDatabase) Then
          frmMain.tv.Nodes.Add objNode.Key, tvwChild, "TBL-" & GetID, txtProperties(0).Text, "table"
          objNode.Text = "Tables (" & objNode.Children & ")"
        End If
      End If
    Next objNode
    
  Else
    StartMsg "Updating Table..."
    
    'Update the tablename if required
    If txtProperties(0).Tag = "Y" Then
      szOldName = objTable.Name
      frmMain.svr.Databases(szDatabase).Tables.Rename szOldName, txtProperties(0).Text
        
      'Update the node text
      For Each objNode In frmMain.tv.Nodes
        If (InStr(1, objNode.FullPath, "\" & szDatabase & "\") <> 0) Then
          If (Left(objNode.Key, 4) = "TBL-") And (objNode.Parent.Parent.Text = szDatabase) And (objNode.Text = szOldName) Then
            objNode.Text = txtProperties(0).Text
          End If
        End If
      Next objNode
    End If
    
    'Add any new columns
    If lvProperties(0).Tag = "Y" Then
      For Each objItem In lvProperties(0).ListItems
        If objItem.Tag <> "ORIG" Then
          If objItem.SubItems(3) = "" Then
            szDataType = objItem.SubItems(2)
          Else
            szDataType = objItem.SubItems(2) & "(" & objItem.SubItems(3) & ")"
          End If
          If objItem.SubItems(5) = "Yes" Then szDataType = szDataType & " NOT NULL"
          frmMain.svr.Databases(szDatabase).Tables(txtProperties(0).Text).Columns.Add objItem.Text, szDataType, objItem.SubItems(4), objItem.SubItems(7)
          For Each objNode In frmMain.tv.Nodes
            If InStr(1, objNode.FullPath, "\" & szDatabase & "\") <> 0 Then
              If (Left(objNode.Key, 4) = "COL+") And (objNode.Parent.Text = txtProperties(0).Text) And (objNode.Parent.Parent.Parent.Text = szDatabase) Then
                frmMain.tv.Nodes.Add objNode.Key, tvwChild, "COL-" & GetID, objItem.Text, "column"
                objNode.Text = "Columns (" & objNode.Children & ")"
              End If
            End If
          Next objNode
        End If
      Next objItem
    End If
    
    'Add any new checks
    If lvProperties(1).Tag = "Y" Then
      For Each objItem In lvProperties(1).ListItems
        If objItem.Tag <> "ORIG" Then
          frmMain.svr.Databases(szDatabase).Tables(txtProperties(0).Text).Checks.Add objItem.Text, objItem.SubItems(1)
          For Each objNode In frmMain.tv.Nodes
            If InStr(1, objNode.FullPath, "\" & szDatabase & "\") <> 0 Then
              If (Left(objNode.Key, 4) = "CHK+") And (objNode.Parent.Text = txtProperties(0).Text) And (objNode.Parent.Parent.Parent.Text = szDatabase) Then
                frmMain.tv.Nodes.Add objNode.Key, tvwChild, "CHK-" & GetID, objItem.Text, "check"
                objNode.Text = "Checks (" & objNode.Children & ")"
              End If
            End If
          Next objNode
        End If
      Next objItem
    End If
    
    'Update the comment
    If hbxProperties(0).Tag = "Y" Then objTable.Comment = hbxProperties(0).Text
  End If
  
  'Set the ACL on the Table as required
  If lvProperties(4).Tag = "Y" Then
    'Revoke all from existing entries
    For Each vEntity In szUsers
      If vEntity <> "" Then frmMain.svr.Databases(szDatabase).Tables(txtProperties(0).Text).Revoke vEntity, aclAll
    Next vEntity
    
    'Now Grant the new permissions
    For Each objItem In lvProperties(4).ListItems
      If objItem.Icon = "group" Then
        szEntity = "GROUP " & QUOTE & objItem.Text & QUOTE
      ElseIf objItem.Icon = "public" Then
        szEntity = "PUBLIC"
      Else
        szEntity = QUOTE & objItem.Text & QUOTE
      End If
      lACL = 0
      If InStr(1, objItem.SubItems(1), "All") <> 0 Then lACL = lACL + aclAll
      If InStr(1, objItem.SubItems(1), "Select") <> 0 Then lACL = lACL + aclSelect
      If InStr(1, objItem.SubItems(1), "Update") <> 0 Then lACL = lACL + aclUpdate
      If InStr(1, objItem.SubItems(1), "Delete") <> 0 Then lACL = lACL + aclDelete
      If InStr(1, objItem.SubItems(1), "Insert") <> 0 Then lACL = lACL + aclInsert
      If InStr(1, objItem.SubItems(1), "Rule") <> 0 Then lACL = lACL + aclRule
      If InStr(1, objItem.SubItems(1), "References") <> 0 Then lACL = lACL + aclReferences
      If InStr(1, objItem.SubItems(1), "Trigger") <> 0 Then lACL = lACL + aclTrigger
      frmMain.svr.Databases(szDatabase).Tables(txtProperties(0).Text).Grant szEntity, lACL
    Next objItem
  End If
  
  'Finally, alter the username if required.
  If cboProperties(0).Tag = "Y" Then frmMain.svr.Databases(szDatabase).Tables(txtProperties(0).Text).Owner = cboProperties(0).Text
  
  'Simulate a node click to refresh the ListTable
  frmMain.tv_NodeClick frmMain.tv.SelectedItem
    
  EndMsg
  Unload Me
  Exit Sub
  
Err_Handler:
  If Err.Number = 35606 Then Resume Next
  EndMsg
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmTable.cmdOK_Click"
End Sub

Public Sub Initialise(szDB As String, Optional Table As pgTable)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmTable.Initialise(" & QUOTE & szDB & QUOTE & ")", etFullDebug

Dim X As Integer
Dim objItem As ListItem
Dim objUser As pgUser
Dim objGroup As pgGroup
Dim objColumn As pgColumn
Dim objCheck As pgCheck
Dim objForeignKey As pgForeignKey
Dim objRelationship As pgRelationship
Dim bFirstRow As Boolean
Dim vInheritedTable As Variant
Dim szUserlist As String
Dim szAccesslist As String
Dim szAccess() As String
  
  szDatabase = szDB
  hbxCheck(0).Wordlist = ctx.AutoHighlight
  
  'ACLs are different in 7.2+ and have 2 extra privileges
  If frmMain.svr.dbVersion.VersionNum < 7.2 Then
    chkPrivilege(6).Enabled = False
    chkPrivilege(7).Enabled = False
  End If
  
  For Each objUser In frmMain.svr.Users
    cboProperties(0).ComboItems.Add , objUser.Name, objUser.Name, "user"
  Next objUser
  
  If Table Is Nothing Then
  
    'Create a new Table
    bNew = True
    Me.Caption = "Create Table"
    
    'Unlock the edittable fields
    cmdInhAdd.Enabled = True
    cmdInhRemove.Enabled = True
    cmdFkyAdd.Enabled = True
    cmdFkyRemove.Enabled = True
    lvProperties(2).BackColor = &H80000005
    lvProperties(3).BackColor = &H80000005
    cboInheritedTables(0).BackColor = &H80000005
    
    'Populate the Combos
    For Each objTable In frmMain.svr.Databases(szDatabase).Tables
      If Not objTable.SystemObject Then
        cboInheritedTables(0).ComboItems.Add , , objTable.Identifier, "table"
      End If
    Next objTable
    
    'Default the owner
    cboProperties(0).ComboItems(ctx.Username).Selected = True
    
    'Redim the userlist so it doesn't cause an error later.
    ReDim szUsers(0)
    
  Else
  
    'Display/Edit the specified Table.
    Set objTable = Table
    bNew = False
    
    If objTable.SystemObject Then  'Lock the permissions Add/Remove buttons if it's a system object
      cmdAdd.Enabled = False
      cmdRemove.Enabled = False
    End If
    
    Me.Caption = "Table: " & objTable.Identifier
    txtProperties(0).Text = objTable.Name
    txtProperties(1).Text = objTable.OID
    cboProperties(0).ComboItems(objTable.Owner).Selected = True
    txtProperties(2).Text = objTable.Rows
    chkProperties(0).Value = Bool2Bin(objTable.HasOIDs)
    hbxProperties(0).Text = objTable.Comment
    
    For Each objColumn In frmMain.svr.Databases(szDatabase).Tables(objTable.Name).Columns
      If Not objColumn.SystemObject Then
        Set objItem = lvProperties(0).ListItems.Add(, , objColumn.Name, "column", "column")
        objItem.SubItems(1) = objColumn.Position
        objItem.SubItems(2) = objColumn.DataType
        If objColumn.DataType = "numeric" Then
          objItem.SubItems(3) = objColumn.Length & ", " & objColumn.NumericScale
        ElseIf objColumn.Length > 0 Then
          objItem.SubItems(3) = objColumn.Length
        End If
        objItem.SubItems(4) = objColumn.Default
        If objColumn.NotNull Then
          objItem.SubItems(5) = "Yes"
        Else
          objItem.SubItems(5) = "No"
        End If
        If objColumn.PrimaryKey Then
          objItem.SubItems(6) = "Yes"
        Else
          objItem.SubItems(6) = "No"
        End If
        objItem.SubItems(7) = objColumn.Comment
        objItem.Tag = "ORIG"
      End If
    Next objColumn
    
    For Each objCheck In frmMain.svr.Databases(szDatabase).Tables(objTable.Name).Checks
      Set objItem = lvProperties(1).ListItems.Add(, , objCheck.Name, "check", "check")
      objItem.SubItems(1) = objCheck.Definition
      objItem.Tag = "ORIG"
    Next objCheck
    
    For Each objForeignKey In frmMain.svr.Databases(szDatabase).Tables(objTable.Name).ForeignKeys
      Set objItem = lvProperties(2).ListItems.Add(, , objForeignKey.Name, "foreignkey", "foreignkey")
      objItem.SubItems(1) = objForeignKey.ReferencedTable
      For Each objRelationship In objForeignKey.Relationships
        objItem.SubItems(2) = objItem.SubItems(2) & objRelationship.LocalColumn & ", "
        objItem.SubItems(3) = objItem.SubItems(3) & objRelationship.ReferencedColumn & ", "
      Next objRelationship
      If Len(objItem.SubItems(2)) > 2 Then objItem.SubItems(2) = Left(objItem.SubItems(2), Len(objItem.SubItems(2)) - 2)
      If Len(objItem.SubItems(3)) > 2 Then objItem.SubItems(3) = Left(objItem.SubItems(3), Len(objItem.SubItems(3)) - 2)
      objItem.SubItems(4) = objForeignKey.OnDelete
      objItem.SubItems(5) = objForeignKey.OnUpdate
      If objForeignKey.Deferrable Then
        objItem.SubItems(6) = "Yes"
      Else
        objItem.SubItems(6) = "No"
      End If
      objItem.SubItems(7) = objForeignKey.Initially
    Next objForeignKey
    
    For Each vInheritedTable In frmMain.svr.Databases(szDatabase).Tables(objTable.Name).InheritedTables
      Set objItem = lvProperties(3).ListItems.Add(, , vInheritedTable, "table", "table")
    Next vInheritedTable
    
    ParseACL objTable.ACL, szUserlist, szAccesslist
    szUsers = Split(szUserlist, "|")
    szAccess = Split(szAccesslist, "|")
    For X = 0 To UBound(szUsers)
      If UCase(Left(szUsers(X), 6)) = "GROUP " Then
        Set objItem = lvProperties(4).ListItems.Add(, , Mid(szUsers(X), 7), "group", "group")
      Else
        If UCase(szUsers(X)) = "PUBLIC" Then
          Set objItem = lvProperties(4).ListItems.Add(, , szUsers(X), "public", "public")
        Else
          Set objItem = lvProperties(4).ListItems.Add(, , szUsers(X), "user", "user")
        End If
      End If
      objItem.SubItems(1) = szAccess(X)
    Next X
  End If
  
  'Load the Entities combo
  cboEntities.ComboItems.Add , , "PUBLIC", "public"
  For Each objUser In frmMain.svr.Users
    cboEntities.ComboItems.Add , , objUser.Name, "user"
  Next objUser
  For Each objGroup In frmMain.svr.Groups
    cboEntities.ComboItems.Add , , objGroup.Name, "group"
  Next objGroup
  cboEntities.ComboItems(1).Selected = True
  
  'Reset the Tags
  txtProperties(0).Tag = "N"
  cboProperties(0).Tag = "N"
  hbxProperties(0).Tag = "N"
  lvProperties(4).Tag = "N"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmTable.Initialise"
End Sub

Private Sub cmdRemove_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmTable.cmdRemove_Click()", etFullDebug

  If lvProperties(4).SelectedItem Is Nothing Then Exit Sub
  lvProperties(4).ListItems.Remove lvProperties(4).SelectedItem.Index
  lvProperties(4).Tag = "Y"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmTable.cmdRemove_Click"
End Sub

Private Sub cmdAdd_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmTable.cmdAdd_Click()", etFullDebug

Dim szAccess As String
Dim objItem As ListItem

  If cboEntities.Text = "" Then Exit Sub
  
  'Check the entry doesn't already exist
  For Each objItem In lvProperties(4).ListItems
    If (objItem.Text = cboEntities.SelectedItem.Text) And (objItem.SmallIcon = cboEntities.SelectedItem.Image) Then
      MsgBox "'" & objItem.Text & "' already appears in the Access Control List. If you wish to modify this entry, it must be removed, and then replaced.", vbExclamation, "Error"
      Exit Sub
    End If
  Next objItem
  
  'Build the access string
  If chkPrivilege(0).Value = 1 Then
    szAccess = "All, "
  Else
    'ACLs are different in 7.2+
    If frmMain.svr.dbVersion.VersionNum < 7.2 Then
      If chkPrivilege(1).Value = 1 Then szAccess = szAccess & "Select, "
      If chkPrivilege(2).Value = 1 Then szAccess = szAccess & "Update/Delete, "
      If chkPrivilege(4).Value = 1 Then szAccess = szAccess & "Insert, "
      If chkPrivilege(5).Value = 1 Then szAccess = szAccess & "Rule, "
    Else
      If chkPrivilege(1).Value = 1 Then szAccess = szAccess & "Select, "
      If chkPrivilege(2).Value = 1 Then szAccess = szAccess & "Update, "
      If chkPrivilege(3).Value = 1 Then szAccess = szAccess & "Delete, "
      If chkPrivilege(4).Value = 1 Then szAccess = szAccess & "Insert, "
      If chkPrivilege(5).Value = 1 Then szAccess = szAccess & "Rule, "
      If chkPrivilege(6).Value = 1 Then szAccess = szAccess & "References, "
      If chkPrivilege(7).Value = 1 Then szAccess = szAccess & "Trigger, "
    End If
  End If
  If Len(szAccess) > 2 Then szAccess = Left(szAccess, Len(szAccess) - 2)
  If szAccess = "" Then szAccess = "None"
  
  Set objItem = lvProperties(4).ListItems.Add(, , cboEntities.SelectedItem.Text, cboEntities.SelectedItem.Image, cboEntities.SelectedItem.Image)
  objItem.SubItems(1) = szAccess
  lvProperties(4).Tag = "Y"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmTable.cmdAdd_Click"
End Sub

Private Sub hbxProperties_Change(Index As Integer)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmTable.hbxProperties_Change(" & Index & ")", etFullDebug

  hbxProperties(Index).Tag = "Y"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmTable.hbxProperties_Change"
End Sub

Private Sub lvProperties_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmTable.lvProperties_ItemClick(" & Index & ", " & Item.Text & ")", etFullDebug

  'Don't allow removal of existing columns
  If Index = 0 Then
    If Item.Tag = "ORIG" Then
      cmdColRemove.Enabled = False
    Else
      cmdColRemove.Enabled = True
    End If
  End If
  
  'Don't allow removal of existing checks
  If Index = 1 Then
    If Item.Tag = "ORIG" Then
      cmdChkRemove.Enabled = False
    Else
      cmdChkRemove.Enabled = True
    End If
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmTable.lvProperties_ItemClick"
End Sub

Private Sub txtProperties_Change(Index As Integer)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmTable.txtProperties_Change(" & Index & ")", etFullDebug

  txtProperties(Index).Tag = "Y"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmTable.txtProperties_Change"
End Sub

Private Sub chkPrivilege_Click(Index As Integer)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmTable.chkPrivilege_Click(" & Index & ")", etFullDebug

Dim X As Integer

  If Index = 0 Then
    'ACLs are different in 7.2+
    If frmMain.svr.dbVersion.VersionNum < 7.2 Then
      If chkPrivilege(0).Value = 1 Then
        For X = 1 To 5
          chkPrivilege(X).Enabled = False
        Next X
      Else
        For X = 1 To 5
          chkPrivilege(X).Enabled = True
        Next X
      End If
    Else
      If chkPrivilege(0).Value = 1 Then
        For X = 1 To 7
          chkPrivilege(X).Enabled = False
        Next X
      Else
        For X = 1 To 7
          chkPrivilege(X).Enabled = True
        Next X
      End If
    End If
  End If
  
  'Link Update/Delete for older versions
  If frmMain.svr.dbVersion.VersionNum < 7.2 Then
    If Index = 2 Then chkPrivilege(3).Value = chkPrivilege(2).Value
    If Index = 3 Then chkPrivilege(2).Value = chkPrivilege(3).Value
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmTable.chkPrivilege_Click"
End Sub

Private Sub cboProperties_Click(Index As Integer)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmTable.cboProperties_Click(" & Index & ")", etFullDebug

  cboProperties(Index).Tag = "Y"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmTable.cboProperties_Click"
End Sub

Private Sub chkProperties_Click(Index As Integer)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmTable.chkProperties_Click(" & Index & ")", etFullDebug

  If frmMain.svr.dbVersion.VersionNum < 7.2 Then
    chkProperties(0).Value = 1
  ElseIf Not (objTable Is Nothing) Then
    chkProperties(0).Value = Bool2Bin(objTable.HasOIDs)
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmUser.chkProperties_Click"
End Sub
