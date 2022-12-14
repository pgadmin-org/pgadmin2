VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{44F33AC4-8757-4330-B063-18608617F23E}#12.4#0"; "HighlightBox.ocx"
Begin VB.Form frmTable 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Table"
   ClientHeight    =   6876
   ClientLeft      =   4992
   ClientTop       =   2496
   ClientWidth     =   5520
   Icon            =   "frmTable.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6876
   ScaleWidth      =   5520
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
      TabPicture(0)   =   "frmTable.frx":06C2
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
      TabPicture(1)   =   "frmTable.frx":06DE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lvProperties(0)"
      Tab(1).Control(1)=   "cmdColAdd"
      Tab(1).Control(2)=   "cmdColRemove"
      Tab(1).Control(3)=   "cmdImport"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "C&hecks"
      TabPicture(2)   =   "frmTable.frx":06FA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblProperties(5)"
      Tab(2).Control(1)=   "lvProperties(1)"
      Tab(2).Control(2)=   "hbxCheck(0)"
      Tab(2).Control(3)=   "cmdChkRemove"
      Tab(2).Control(4)=   "cmdChkAdd"
      Tab(2).Control(5)=   "txtCheck(0)"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "&Foreign Keys"
      TabPicture(3)   =   "frmTable.frx":0716
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lvProperties(2)"
      Tab(3).Control(1)=   "cmdFkyAdd"
      Tab(3).Control(2)=   "cmdFkyRemove"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "&Inherits"
      TabPicture(4)   =   "frmTable.frx":0732
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "lblProperties(6)"
      Tab(4).Control(1)=   "lvProperties(3)"
      Tab(4).Control(2)=   "cmdInhAdd"
      Tab(4).Control(3)=   "cmdInhRemove"
      Tab(4).Control(4)=   "cboInheritedTables(0)"
      Tab(4).ControlCount=   5
      TabCaption(5)   =   "&Security"
      TabPicture(5)   =   "frmTable.frx":074E
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "cmdRemove"
      Tab(5).Control(1)=   "fraAdd"
      Tab(5).Control(2)=   "cmdAdd"
      Tab(5).Control(3)=   "lvProperties(4)"
      Tab(5).ControlCount=   4
      Begin VB.CommandButton cmdImport 
         Caption         =   "&Import"
         Height          =   375
         Left            =   -72240
         TabIndex        =   44
         ToolTipText     =   "Import column from table."
         Top             =   5805
         Width           =   1230
      End
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
         Height          =   300
         Index           =   0
         Left            =   1932
         TabIndex        =   43
         ToolTipText     =   "The tables owner."
         Top             =   1440
         Width           =   3396
         _ExtentX        =   5990
         _ExtentY        =   529
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483633
         Locked          =   -1  'True
         ImageList       =   "il"
      End
      Begin MSComctlLib.ImageCombo cboInheritedTables 
         Height          =   300
         Index           =   0
         Left            =   -73068
         TabIndex        =   22
         Top             =   5856
         Width           =   3396
         _ExtentX        =   5990
         _ExtentY        =   529
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
         Left            =   -74880
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
         _ExtentX        =   9165
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
            Height          =   300
            Left            =   1260
            TabIndex        =   26
            ToolTipText     =   "Select a user, group or 'PUBLIC'."
            Top             =   312
            Width           =   3708
            _ExtentX        =   6541
            _ExtentY        =   529
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
         _ExtentX        =   9165
         _ExtentY        =   5990
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
         _ExtentX        =   9165
         _ExtentY        =   6138
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
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
         _ExtentX        =   9165
         _ExtentY        =   9229
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
         _ExtentX        =   9165
         _ExtentY        =   8446
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
         _ExtentX        =   9165
         _ExtentY        =   1545
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
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
         _ExtentX        =   9165
         _ExtentY        =   6858
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
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTable.frx":076A
            Key             =   "column"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTable.frx":0D04
            Key             =   "table"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTable.frx":0E5E
            Key             =   "foreignkey"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTable.frx":13F8
            Key             =   "public"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTable.frx":1552
            Key             =   "group"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTable.frx":1AEC
            Key             =   "user"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTable.frx":2086
            Key             =   "check"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTable.frx":2620
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
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the Artistic Licence
'
' frmTable.frm - Edit/Create a Table

Option Explicit

Dim bNew As Boolean
Dim szDatabase As String
Dim szNamespace As String
Dim szDropCheckList As String
Dim szDropColumnList As String
Dim szDropForeignKeyList As String
Dim szUsers() As String
Public objTable As pgTable

Private Sub cmdCancel_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmTable.cmdCancel_Click()", etFullDebug

  Unload Me
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmTable.cmdCancel_Click"
End Sub

Private Sub cmdChkAdd_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmTable.cmdChkAdd_Click()", etFullDebug

Dim objItem As ListItem

  If txtCheck(0).Text = "" Then
    MsgBox ??TrasLang??("You must enter a name for the check!"), vbExclamation, ??TrasLang??("Error")
    tabProperties.Tab = 2
    txtCheck(0).SetFocus
    Exit Sub
  End If
  If hbxCheck(0).Text = "" Then
    MsgBox ??TrasLang??("You must enter a definition for the check!"), vbExclamation, ??TrasLang??("Error")
    tabProperties.Tab = 2
    hbxCheck(0).SetFocus
    Exit Sub
  End If
  For Each objItem In lvProperties(1).ListItems
    If objItem.Text = txtCheck(0).Text Then
      MsgBox ??TrasLang??("This check name is already in the list!"), vbExclamation, ??TrasLang??("Error")
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
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmTable.cmdChkRemove_Click()", etFullDebug

  If lvProperties(1).SelectedItem Is Nothing Then
    MsgBox ??TrasLang??("You must select a check to remove!"), vbExclamation, ??TrasLang??("Error")
    tabProperties.Tab = 2
    lvProperties(1).SetFocus
    Exit Sub
  End If
  
  If objTable Is Nothing Then
    lvProperties(1).ListItems.Remove lvProperties(1).SelectedItem.Index
    lvProperties(1).Tag = "Y"
    If lvProperties(1).SelectedItem Is Nothing Then
      cmdChkRemove.Enabled = False
    Else
      lvProperties_ItemClick 1, lvProperties(1).SelectedItem
    End If
  Else
    szDropCheckList = szDropCheckList & lvProperties(1).SelectedItem.Text & "!|!"
    lvProperties(1).ListItems.Remove lvProperties(1).SelectedItem.Index
    lvProperties(1).Tag = "Y"
    If lvProperties(1).SelectedItem Is Nothing Then
      cmdChkRemove.Enabled = False
    Else
      lvProperties_ItemClick 1, lvProperties(1).SelectedItem
    End If
  End If
  
  Exit Sub
  
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmTable.cmdChkRemove_Click"
End Sub

Private Sub cmdColAdd_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmTable.cmdColAdd_Click()", etFullDebug

Dim objColumnForm As New frmColumn
  
  Load objColumnForm
  objColumnForm.Initialise szDatabase, szNamespace, "TA", , Me
  objColumnForm.Show

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmTable.cmdColAdd_Click"
End Sub

Private Sub cmdColRemove_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmTable.cmdColRemove_Click()", etFullDebug

  If lvProperties(0).SelectedItem Is Nothing Then
    MsgBox "You must select a column to remove!", vbExclamation, "Error"
    tabProperties.Tab = 1
    lvProperties(0).SetFocus
    Exit Sub
  End If
  
  If objTable Is Nothing Then
    lvProperties(0).ListItems.Remove lvProperties(0).SelectedItem.Index
    lvProperties(0).Tag = "Y"
    If lvProperties(0).SelectedItem Is Nothing Then
      cmdColRemove.Enabled = False
    Else
      lvProperties_ItemClick 0, lvProperties(0).SelectedItem
    End If
  Else
    szDropColumnList = szDropColumnList & lvProperties(0).SelectedItem.Text & "!|!"
    lvProperties(0).ListItems.Remove lvProperties(0).SelectedItem.Index
    lvProperties(0).Tag = "Y"
    If lvProperties(1).SelectedItem Is Nothing Then
      cmdColRemove.Enabled = False
    Else
      lvProperties_ItemClick 0, lvProperties(0).SelectedItem
    End If
  End If
  
  Exit Sub

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmTable.cmdColRemove_Click"
End Sub

Private Sub cmdFkyAdd_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmTable.cmdFkyAdd_Click()", etFullDebug

Dim objForeignKeyForm As New frmForeignKey
  
  Load objForeignKeyForm
  objForeignKeyForm.Initialise szDatabase, szNamespace, "TA", , Me
  objForeignKeyForm.Show
  
  Exit Sub
  
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmTable.cmdFkyAdd_Click"
End Sub

Private Sub cmdFkyRemove_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
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
  
  If Not objTable Is Nothing Then
    szDropForeignKeyList = szDropForeignKeyList & lvProperties(2).SelectedItem.Text & "!|!"
  End If
  lvProperties(2).ListItems.Remove lvProperties(2).SelectedItem.Index

  Exit Sub
  
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmTable.cmdFkyRemove_Click"
End Sub

Private Sub cmdImport_Click()
If inIDE Then:  On Error GoTo 0: Else: On Error GoTo Err_Handler:
frmMain.svr.LogEvent "Entering " & App.Title & ":frmTable.cmdImport_Click()", etFullDebug

Dim objImportColumnForm As New frmImportColumn
  
  Load objImportColumnForm
  objImportColumnForm.Initialise Me
  objImportColumnForm.Show
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmTable.cmdImport_Click"
End Sub

Private Sub cmdInhAdd_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmTable.cmdInhAdd_Click()", etFullDebug

Dim objItem As ListItem

  If cboInheritedTables(0).Text = "" Then
    MsgBox ??TrasLang??("You must select a table to add!"), vbExclamation, ??TrasLang??("Error")
    tabProperties.Tab = 4
    cboInheritedTables(0).SetFocus
    Exit Sub
  End If
  For Each objItem In lvProperties(3).ListItems
    If objItem.Text = cboInheritedTables(0).Text Then
      MsgBox ??TrasLang??("This table is already in the list!"), vbExclamation, ??TrasLang??("Error")
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
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmTable.cmdInhRemove_Click()", etFullDebug

  If lvProperties(3).SelectedItem Is Nothing Then
    MsgBox ??TrasLang??("You must select a table to remove!"), vbExclamation, ??TrasLang??("Error")
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
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmTable.cmdOK_Click()", etFullDebug

Dim objNode As Node
Dim objItem As ListItem
Dim objNewTable As pgTable
Dim objNewColumn As pgColumn
Dim objNewCheck As pgCheck
Dim objNewForeignKey As pgForeignKey
Dim lACL As Long
Dim szEntity As String
Dim vEntity As Variant
Dim szOldName As String
Dim szDataType As String
Dim szColumns As String
Dim szPrimaryKeys As String
Dim szChecks As String
Dim szDropChecks() As String
Dim szDropColumns() As String
Dim szDropForeignKeys() As String
Dim szForeignKeys As String
Dim szInherits As String
Dim X As Integer
Dim bFlag As Boolean

  'Check the data
  If txtProperties(0).Text = "" Then
    MsgBox ??TrasLang??("You must specify a Table name!"), vbExclamation, ??TrasLang??("Error")
    tabProperties.Tab = 0
    txtProperties(0).SetFocus
    Exit Sub
  End If
  If lvProperties(0).ListItems.Count = 0 And lvProperties(3).ListItems.Count = 0 Then
    MsgBox ??TrasLang??("You must define at least one column or inherited table!"), vbExclamation, ??TrasLang??("Error")
    tabProperties.Tab = 1
    lvProperties(0).SetFocus
    Exit Sub
  End If
  
  If bNew Then
    StartMsg ??TrasLang??("Creating Table...")
    
    'Build the column list
    For Each objItem In lvProperties(0).ListItems
      szColumns = szColumns & fmtID(objItem.Text) & " " & objItem.SubItems(2)
      If objItem.SubItems(3) <> "" Then szColumns = szColumns & "(" & objItem.SubItems(3) & ")"
      If objItem.SubItems(4) <> "" Then szColumns = szColumns & " DEFAULT " & objItem.SubItems(4)
      If objItem.SubItems(5) <> "No" Then szColumns = szColumns & " NOT NULL"
      szColumns = szColumns & ", "
      
      'Add to the Primary Key list if required.
      If objItem.SubItems(6) <> "No" Then szPrimaryKeys = szPrimaryKeys & fmtID(objItem.Text) & ", "
    Next objItem
    If Len(szColumns) > 2 Then szColumns = Left(szColumns, Len(szColumns) - 2)
    
    'Add the Primary Keys
    If Len(szPrimaryKeys) > 2 Then szPrimaryKeys = Left(szPrimaryKeys, Len(szPrimaryKeys) - 2)

    'Add Checks
    For Each objItem In lvProperties(1).ListItems
      szChecks = szChecks & "CONSTRAINT " & fmtID(objItem.Text) & " CHECK (" & objItem.SubItems(1) & "), "
    Next objItem
    If Len(szChecks) > 2 Then szChecks = Left(szChecks, Len(szChecks) - 2)
    
    'Add Foreign Keys
    For Each objItem In lvProperties(2).ListItems
      szForeignKeys = szForeignKeys & "CONSTRAINT " & fmtID(objItem.Text) & " FOREIGN KEY (" & objItem.SubItems(2) & ") "
      szForeignKeys = szForeignKeys & "REFERENCES " & objItem.SubItems(1) & " (" & objItem.SubItems(3) & ")"
      szForeignKeys = szForeignKeys & " ON DELETE " & UCase(objItem.SubItems(4))
      szForeignKeys = szForeignKeys & " ON UPDATE " & UCase(objItem.SubItems(5))
      If objItem.SubItems(6) = "Yes" Then szForeignKeys = szForeignKeys & " DEFERRABLE"
      szForeignKeys = szForeignKeys & " INITIALLY " & UCase(objItem.SubItems(7)) & ", "
    Next objItem
    If Len(szForeignKeys) > 2 Then szForeignKeys = Left(szForeignKeys, Len(szForeignKeys) - 2)
    
    'Add Inherits
    For Each objItem In lvProperties(3).ListItems
      szInherits = szInherits & objItem.Text & ", "
    Next objItem
    If Len(szInherits) > 2 Then szInherits = Left(szInherits, Len(szInherits) - 2)
     
    Set objNewTable = frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Tables.Add(txtProperties(0).Text, szColumns, szPrimaryKeys, szChecks, szForeignKeys, szInherits, hbxProperties(0).Text, Bin2Bool(chkProperties(0).Value))
    
    'Add any comments for the columns.
    For Each objItem In lvProperties(0).ListItems
      If objItem.SubItems(7) <> "" Then frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Tables(txtProperties(0).Text).Columns(objItem.Text).Comment = objItem.SubItems(7)
    Next objItem
    
    'Add a new node and update the text on the parent
    On Error Resume Next
    Set objNode = frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Tables.Tag
    Set objNewTable.Tag = frmMain.tv.Nodes.Add(objNode.Key, tvwChild, "TBL-" & GetID, txtProperties(0).Text, "table")
    objNode.Text = ??TrasLang??("Tables (") & objNode.Children & ")"
    If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
 
  Else
    StartMsg ??TrasLang??("Updating Table...")
    
    'Update the tablename if required
    If txtProperties(0).Tag = "Y" Then
      szOldName = objTable.Name
      frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Tables.Rename szOldName, txtProperties(0).Text
        
      'Update the node text
      frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Tables(txtProperties(0).Text).Tag.Text = txtProperties(0).Text
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
          Set objNewColumn = frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Tables(txtProperties(0).Text).Columns.Add(objItem.Text, szDataType, objItem.SubItems(4), objItem.SubItems(7))
          If objItem.SubItems(5) = "Yes" Then objNewColumn.NotNull = True
          If objItem.SubItems(6) = "Yes" Then objNewColumn.PrimaryKey = True
          If Len(frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Tables(txtProperties(0).Text).Columns.Tag) > 0 Then
            Set objNode = frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Tables(txtProperties(0).Text).Columns.Tag
            Set objNewColumn.Tag = frmMain.tv.Nodes.Add(objNode.Key, tvwChild, "COL-" & GetID, objItem.Text, "column")
            objNode.Text = ??TrasLang??("Columns (") & objNode.Children & ")"
          End If
        End If
      Next objItem
    End If
    
    'Drop any old columns
    If Len(szDropColumnList) > 3 Then
      szDropColumns = Split(szDropColumnList, "!|!")
      For X = 0 To UBound(szDropColumns)
        If szDropColumns(X) <> "" Then
          If frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Tables(txtProperties(0).Text).Columns.Exists(szDropColumns(X)) Then
            If IsObject(frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Tables(txtProperties(0).Text).Columns(szDropColumns(X)).Tag) Then
              Set objNode = frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Tables(txtProperties(0).Text).Columns(szDropColumns(X)).Tag
              bFlag = True
            Else
              bFlag = False
            End If
            frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Tables(txtProperties(0).Text).Columns.Remove szDropColumns(X)
            If bFlag Then
              objNode.Parent.Text = ??TrasLang??("Columns (") & objNode.Children - 1 & ")"
              frmMain.tv.Nodes.Remove objNode.Index
            End If
          End If
        End If
      Next X
    End If
    
    'Add any new checks
    If lvProperties(1).Tag = "Y" Then
      For Each objItem In lvProperties(1).ListItems
        If objItem.Tag <> "ORIG" Then
          Set objNewCheck = frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Tables(txtProperties(0).Text).Checks.Add(objItem.Text, objItem.SubItems(1))
          If Len(frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Tables(txtProperties(0).Text).Checks.Tag) > 0 Then
            Set objNode = frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Tables(txtProperties(0).Text).Checks.Tag
            Set objNewCheck.Tag = frmMain.tv.Nodes.Add(objNode.Key, tvwChild, "CHK-" & GetID, objItem.Text, "check")
            objNode.Text = ??TrasLang??("Checks (") & objNode.Children & ")"
          End If
        End If
      Next objItem
    End If
    
    'Drop any old checks
    If Len(szDropCheckList) > 3 Then
      szDropChecks = Split(szDropCheckList, "!|!")
      For X = 0 To UBound(szDropChecks)
        If szDropChecks(X) <> "" Then
          If frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Tables(txtProperties(0).Text).Checks.Exists(szDropChecks(X)) Then
            If IsObject(frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Tables(txtProperties(0).Text).Checks(szDropChecks(X)).Tag) Then
              Set objNode = frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Tables(txtProperties(0).Text).Checks(szDropChecks(X)).Tag
              bFlag = True
            Else
              bFlag = False
            End If
            frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Tables(txtProperties(0).Text).Checks.Remove szDropChecks(X)
            If bFlag Then
              objNode.Parent.Text = ??TrasLang??("Checks (") & objNode.Children - 1 & ")"
              frmMain.tv.Nodes.Remove objNode.Index
            End If
          End If
        End If
      Next X
    End If
    
    'Add new Foreign Keys
    If lvProperties(2).Tag = "Y" Then
      For Each objItem In lvProperties(2).ListItems
        If objItem.Tag <> "ORIG" Then
          szForeignKeys = " FOREIGN KEY (" & objItem.SubItems(2) & ") "
          szForeignKeys = szForeignKeys & "REFERENCES " & objItem.SubItems(1) & " (" & objItem.SubItems(3) & ")"
          szForeignKeys = szForeignKeys & " ON DELETE " & UCase(objItem.SubItems(4))
          szForeignKeys = szForeignKeys & " ON UPDATE " & UCase(objItem.SubItems(5))
          If objItem.SubItems(6) = "Yes" Then szForeignKeys = szForeignKeys & " DEFERRABLE"
          szForeignKeys = szForeignKeys & " INITIALLY " & UCase(objItem.SubItems(7))
          
          Set objNewForeignKey = frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Tables(txtProperties(0).Text).ForeignKeys.Add(objItem.Text, szForeignKeys)
          If Len(frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Tables(txtProperties(0).Text).ForeignKeys.Tag) > 0 Then
            Set objNode = frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Tables(txtProperties(0).Text).ForeignKeys.Tag
            Set objNewForeignKey.Tag = frmMain.tv.Nodes.Add(objNode.Key, tvwChild, "FKY-" & GetID, objItem.Text, "foreignkey")
            objNode.Text = ??TrasLang??("Foreign Keys (") & objNode.Children & ")"
          End If
        End If
      Next objItem
    End If
    
    'Drop any old ForeignKey
    If Len(szDropForeignKeyList) > 3 Then
      szDropForeignKeys = Split(szDropForeignKeyList, "!|!")
      For X = 0 To UBound(szDropForeignKeys)
        If szDropForeignKeys(X) <> "" Then
          If frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Tables(txtProperties(0).Text).ForeignKeys.Exists(szDropForeignKeys(X)) Then
            If IsObject(frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Tables(txtProperties(0).Text).ForeignKeys(szDropForeignKeys(X)).Tag) Then
              Set objNode = frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Tables(txtProperties(0).Text).ForeignKeys(szDropForeignKeys(X)).Tag
              bFlag = True
            Else
              bFlag = False
            End If
            frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Tables(txtProperties(0).Text).ForeignKeys.Remove szDropForeignKeys(X)
            If bFlag Then
              objNode.Parent.Text = ??TrasLang??("Foreign Keys (") & objNode.Children - 1 & ")"
              frmMain.tv.Nodes.Remove objNode.Index
            End If
          End If
        End If
      Next X
    End If
    
    'Update the comment
    If hbxProperties(0).Tag = "Y" Then objTable.Comment = hbxProperties(0).Text
  End If
  
  'Set the ACL on the Table as required
  If lvProperties(4).Tag = "Y" Then
    'Revoke all from existing entries
    For Each vEntity In szUsers
      If vEntity <> "" Then
        If vEntity = "PUBLIC" Then
          frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Tables(txtProperties(0).Text).Revoke vEntity, aclAll
        ElseIf Left(vEntity, 6) = "GROUP " Then
          frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Tables(txtProperties(0).Text).Revoke "GROUP " & fmtID(Mid(vEntity, 7)), aclAll
        Else
          frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Tables(txtProperties(0).Text).Revoke fmtID(vEntity), aclAll
        End If
      End If
    Next vEntity
    
    'Now Grant the new permissions
    For Each objItem In lvProperties(4).ListItems
      If objItem.Icon = "group" Then
        szEntity = "GROUP " & fmtID(objItem.Text)
      ElseIf objItem.Icon = "public" Then
        szEntity = "PUBLIC"
      Else
        szEntity = fmtID(objItem.Text)
      End If
      lACL = 0
      If InStr(1, objItem.SubItems(1), ??TrasLang??("All")) <> 0 Then lACL = lACL + aclAll
      If InStr(1, objItem.SubItems(1), ??TrasLang??("Select")) <> 0 Then lACL = lACL + aclSelect
      If InStr(1, objItem.SubItems(1), ??TrasLang??("Update")) <> 0 Then lACL = lACL + aclUpdate
      If InStr(1, objItem.SubItems(1), ??TrasLang??("Delete")) <> 0 Then lACL = lACL + aclDelete
      If InStr(1, objItem.SubItems(1), ??TrasLang??("Insert")) <> 0 Then lACL = lACL + aclInsert
      If InStr(1, objItem.SubItems(1), ??TrasLang??("Rule")) <> 0 Then lACL = lACL + aclRule
      If InStr(1, objItem.SubItems(1), ??TrasLang??("References")) <> 0 Then lACL = lACL + aclReferences
      If InStr(1, objItem.SubItems(1), ??TrasLang??("Trigger")) <> 0 Then lACL = lACL + aclTrigger
      frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Tables(txtProperties(0).Text).Grant szEntity, lACL
    Next objItem
  End If
  
  'Finally, alter the username if required.
  If (cboProperties(0).Tag = "Y") And Not (frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Tables(txtProperties(0).Text).SystemObject) Then
    frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Tables(txtProperties(0).Text).Owner = cboProperties(0).Text
  End If
  
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

Public Sub Initialise(szDB As String, szNS As String, Optional Table As pgTable)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmTable.Initialise(" & Quote & szDB & Quote & ")", etFullDebug

Dim X As Integer
Dim objItem As ListItem
Dim objUser As pgUser
Dim objColumn As pgColumn
Dim objCheck As pgCheck
Dim objForeignKey As pgForeignKey
Dim objRelationship As pgRelationship
Dim objNamespace As pgNamespace
Dim bFirstRow As Boolean
Dim vInheritedTable As Variant
Dim szUserlist As String
Dim szAccesslist As String
Dim szAccess() As String
  
  szDatabase = szDB
  szNamespace = szNS
  
  PatchForm Me
  hbxCheck(0).Wordlist = ctx.AutoHighlight
  
  'ACLs are different in 7.2+ and have 2 extra privileges
  If ctx.dbVer < 7.2 Then
    chkPrivilege(6).Enabled = False
    chkPrivilege(7).Enabled = False
  End If
  
  For Each objUser In frmMain.svr.Users
    cboProperties(0).ComboItems.Add , "U~" & objUser.Name, objUser.Name, "user"
  Next objUser
  
  If Table Is Nothing Then
  
    'Create a new Table
    bNew = True
    Me.Caption = ??TrasLang??("Create Table")
    
    'Unlock the edittable fields
    cmdInhAdd.Enabled = True
    cmdInhRemove.Enabled = True
    cmdFkyAdd.Enabled = True
    cmdFkyRemove.Enabled = True
    lvProperties(2).BackColor = &H80000005
    lvProperties(3).BackColor = &H80000005
    cboInheritedTables(0).BackColor = &H80000005
    cboProperties(0).BackColor = &H80000005
    
    'Populate the Combos
    If ctx.dbVer >= 7.3 Then
      For Each objNamespace In frmMain.svr.Databases(szDatabase).Namespaces
        If (Not objNamespace.SystemObject) Or (objNamespace.Name = "public") Then
          For Each objTable In objNamespace.Tables
            If Not objTable.SystemObject Then
              cboInheritedTables(0).ComboItems.Add , , objTable.FormattedID, "table"
            End If
          Next objTable
        End If
      Next objNamespace
    Else
      For Each objTable In frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Tables
        If Not objTable.SystemObject Then
          cboInheritedTables(0).ComboItems.Add , , objTable.FormattedID, "table"
        End If
      Next objTable
    End If
    
    'Default the owner
    cboProperties(0).ComboItems("U~" & ctx.Username).Selected = True
    
    'Redim the userlist so it doesn't cause an error later.
    ReDim szUsers(0)
    
  Else
  
    'Display/Edit the specified Table.
    Set objTable = Table
    bNew = False
    
    If objTable.SystemObject Then  'Lock the permissions Add/Remove buttons if it's a system object
      cmdAdd.Enabled = False
      cmdRemove.Enabled = False
    Else
      cboProperties(0).BackColor = &H80000005
    End If
    
    'Allow DROP CHECK for 7.2+
    If ctx.dbVer >= 7.2 Then cmdChkRemove.Enabled = True
    
    If ctx.dbVer >= 7.3 Then
      cmdFkyAdd.Enabled = True
      cmdFkyRemove.Enabled = True
      lvProperties(2).BackColor = &H80000005
    End If
    
    Me.Caption = ??TrasLang??("Table: ") & objTable.Identifier
    txtProperties(0).Text = objTable.Name
    txtProperties(1).Text = objTable.Oid
    If objTable.SystemObject Then
      cboProperties(0).ComboItems.Clear
      cboProperties(0).ComboItems.Add , "U~" & objTable.Owner, objTable.Owner, "user", "user"
    End If
    cboProperties(0).ComboItems("U~" & objTable.Owner).Selected = True
    txtProperties(2).Text = objTable.Rows
    chkProperties(0).Value = Bool2Bin(objTable.HasOIDs)
    hbxProperties(0).Text = objTable.Comment
    
    For Each objColumn In objTable.Columns
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
        objItem.SubItems(5) = BoolToYesNo(objColumn.NotNull)
        objItem.SubItems(6) = BoolToYesNo(objColumn.PrimaryKey)
        objItem.SubItems(7) = objColumn.Comment
        objItem.Tag = "ORIG"
      End If
    Next objColumn
    
    For Each objCheck In objTable.Checks
      Set objItem = lvProperties(1).ListItems.Add(, , objCheck.Name, "check", "check")
      objItem.SubItems(1) = objCheck.Definition
      objItem.Tag = "ORIG"
    Next objCheck
    
    For Each objForeignKey In objTable.ForeignKeys
      Set objItem = lvProperties(2).ListItems.Add(, , objForeignKey.Identifier, "foreignkey", "foreignkey")
      objItem.SubItems(1) = objForeignKey.ReferencedTable
      For Each objRelationship In objForeignKey.Relationships
        objItem.SubItems(2) = objItem.SubItems(2) & objRelationship.LocalColumn & ", "
        objItem.SubItems(3) = objItem.SubItems(3) & objRelationship.ReferencedColumn & ", "
      Next objRelationship
      If Len(objItem.SubItems(2)) > 2 Then objItem.SubItems(2) = Left(objItem.SubItems(2), Len(objItem.SubItems(2)) - 2)
      If Len(objItem.SubItems(3)) > 2 Then objItem.SubItems(3) = Left(objItem.SubItems(3), Len(objItem.SubItems(3)) - 2)
      objItem.SubItems(4) = objForeignKey.OnDelete
      objItem.SubItems(5) = objForeignKey.OnUpdate
      objItem.SubItems(6) = BoolToYesNo(objForeignKey.Deferrable)
      objItem.SubItems(7) = objForeignKey.Initially
      objItem.Tag = "ORIG"
    Next objForeignKey
    
    For Each vInheritedTable In objTable.InheritedTables
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
  LoadUGACL cboEntities
  
  'Reset the Tags
  txtProperties(0).Tag = "N"
  cboProperties(0).Tag = "N"
  hbxProperties(0).Tag = "N"
  lvProperties(4).Tag = "N"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmTable.Initialise"
End Sub

Private Sub cmdRemove_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmTable.cmdRemove_Click()", etFullDebug

  If lvProperties(4).SelectedItem Is Nothing Then Exit Sub
  lvProperties(4).ListItems.Remove lvProperties(4).SelectedItem.Index
  lvProperties(4).Tag = "Y"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmTable.cmdRemove_Click"
End Sub

Private Sub cmdAdd_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmTable.cmdAdd_Click()", etFullDebug

Dim szAccess As String
Dim objItem As ListItem

  If cboEntities.Text = "" Then Exit Sub
  
  'Check the entry doesn't already exist
  For Each objItem In lvProperties(4).ListItems
    If (objItem.Text = cboEntities.SelectedItem.Text) And (objItem.SmallIcon = cboEntities.SelectedItem.Image) Then
      MsgBox "'" & objItem.Text & ??TrasLang??("' already appears in the Access Control List. If you wish to modify this entry, it must be removed, and then replaced."), vbExclamation, ??TrasLang??("Error")
      Exit Sub
    End If
  Next objItem
  
  'Build the access string
  If chkPrivilege(0).Value = 1 Then
    szAccess = "All, "
  Else
    'ACLs are different in 7.2+
    If ctx.dbVer < 7.2 Then
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
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmTable.hbxProperties_Change(" & Index & ")", etFullDebug

  hbxProperties(Index).Tag = "Y"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmTable.hbxProperties_Change"
End Sub

Private Sub lvProperties_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmTable.lvProperties_ItemClick(" & Index & ", " & Item.Text & ")", etFullDebug

  'Don't allow removal of existing columns on pre 7.3 dbs
  If Index = 0 Then
    If ((Item.Tag = "ORIG") And (ctx.dbVer < 7.3)) Then
      cmdColRemove.Enabled = False
    Else
      cmdColRemove.Enabled = True
    End If
  End If
  
  'Don't allow removal of existing checks
  If Index = 1 And ctx.dbVer < 7.2 Then
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
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmTable.txtProperties_Change(" & Index & ")", etFullDebug

  txtProperties(Index).Tag = "Y"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmTable.txtProperties_Change"
End Sub

Private Sub chkPrivilege_Click(Index As Integer)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmTable.chkPrivilege_Click(" & Index & ")", etFullDebug

Dim X As Integer

  If Index = 0 Then
    'ACLs are different in 7.2+
    If ctx.dbVer < 7.2 Then
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
  If ctx.dbVer < 7.2 Then
    If Index = 2 Then chkPrivilege(3).Value = chkPrivilege(2).Value
    If Index = 3 Then chkPrivilege(2).Value = chkPrivilege(3).Value
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmTable.chkPrivilege_Click"
End Sub

Private Sub cboProperties_Click(Index As Integer)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmTable.cboProperties_Click(" & Index & ")", etFullDebug

  cboProperties(Index).Tag = "Y"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmTable.cboProperties_Click"
End Sub

Private Sub chkProperties_Click(Index As Integer)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmTable.chkProperties_Click(" & Index & ")", etFullDebug

  If ctx.dbVer < 7.2 Then
    chkProperties(0).Value = 1
  ElseIf Not (objTable Is Nothing) Then
    chkProperties(0).Value = Bool2Bin(objTable.HasOIDs)
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmUser.chkProperties_Click"
End Sub
