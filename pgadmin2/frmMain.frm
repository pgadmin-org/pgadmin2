VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{44F33AC4-8757-4330-B063-18608617F23E}#12.4#0"; "HighlightBox.ocx"
Begin VB.Form frmMain 
   Caption         =   "pgAdmin II"
   ClientHeight    =   6660
   ClientLeft      =   3120
   ClientTop       =   1668
   ClientWidth     =   9684
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6660
   ScaleWidth      =   9684
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   8550
      Top             =   90
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin HighlightBox.HBX txtDefinition 
      Height          =   1635
      Left            =   3825
      TabIndex        =   3
      ToolTipText     =   "Displays the SQL Definition of the currently selected object."
      Top             =   4275
      Width           =   5820
      _ExtentX        =   10266
      _ExtentY        =   2879
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
      Caption         =   "Definition"
   End
   Begin MSComctlLib.ImageList ilTB 
      Left            =   9090
      Top             =   45
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0A02
            Key             =   "connect"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15D4
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1EAE
            Key             =   "create"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2A80
            Key             =   "drop"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3652
            Key             =   "properties"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4224
            Key             =   "sql"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4DF6
            Key             =   "viewdata"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":59C8
            Key             =   "vacuum"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":659A
            Key             =   "record"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6E74
            Key             =   "stop"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":774E
            Key             =   "statistics"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8320
            Key             =   "reindex"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tb 
      Align           =   1  'Align Top
      Height          =   480
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9684
      _ExtentX        =   17082
      _ExtentY        =   847
      ButtonWidth     =   826
      ButtonHeight    =   804
      Appearance      =   1
      Style           =   1
      ImageList       =   "ilTB"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   16
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "connect"
            Description     =   "Connect"
            Object.ToolTipText     =   "Connect to a Server."
            ImageKey        =   "connect"
            Style           =   5
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "refresh"
            Description     =   "Refresh"
            Object.ToolTipText     =   "Refresh the data below the selected object."
            ImageKey        =   "refresh"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "sep1"
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "create"
            Description     =   "Create"
            Object.ToolTipText     =   "Create a new object."
            ImageKey        =   "create"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   18
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "aggregate"
                  Text            =   "&Aggregate"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "cast"
                  Text            =   "&Cast"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "conversion"
                  Text            =   "&Conversion"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "database"
                  Text            =   "&Database"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "domain"
                  Text            =   "Do&main"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "function"
                  Text            =   "&Function"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "group"
                  Text            =   "&Group"
               EndProperty
               BeginProperty ButtonMenu8 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "index"
                  Text            =   "&Index"
               EndProperty
               BeginProperty ButtonMenu9 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "language"
                  Text            =   "&Language"
               EndProperty
               BeginProperty ButtonMenu10 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "operator"
                  Text            =   "&Operator"
               EndProperty
               BeginProperty ButtonMenu11 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "rule"
                  Text            =   "&Rule"
               EndProperty
               BeginProperty ButtonMenu12 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "namespace"
                  Text            =   "Sc&hema"
               EndProperty
               BeginProperty ButtonMenu13 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "sequence"
                  Text            =   "&Sequence"
               EndProperty
               BeginProperty ButtonMenu14 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "table"
                  Text            =   "&Table"
               EndProperty
               BeginProperty ButtonMenu15 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "trigger"
                  Text            =   "T&rigger"
               EndProperty
               BeginProperty ButtonMenu16 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "type"
                  Text            =   "T&ype"
               EndProperty
               BeginProperty ButtonMenu17 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "user"
                  Text            =   "&User"
               EndProperty
               BeginProperty ButtonMenu18 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "view"
                  Text            =   "&View"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "drop"
            Description     =   "Drop"
            Object.ToolTipText     =   "Drop the selected object."
            ImageKey        =   "drop"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "properties"
            Description     =   "Properties"
            Object.ToolTipText     =   "View/Edit the properties for the selected object."
            ImageKey        =   "properties"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "sep2"
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "sql"
            Description     =   "SQL"
            Object.ToolTipText     =   "Execute arbitrary SQL queries."
            ImageKey        =   "sql"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "viewdata"
            Description     =   "View Data"
            Object.ToolTipText     =   "View the data in the selected table/view"
            ImageKey        =   "viewdata"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "sep3"
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "resetstatistics"
            Object.ToolTipText     =   "Reset Database Statistics"
            ImageKey        =   "statistics"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "reindex"
            Object.ToolTipText     =   "Reindex the current object"
            ImageKey        =   "reindex"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "vacuum"
            Description     =   "Vacuum"
            Object.ToolTipText     =   "Vacuum the selected object."
            ImageKey        =   "vacuum"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "vacuum"
                  Text            =   "&Vacuum"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "analyse"
                  Text            =   "Vacuum (&Analyse)"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "sep4"
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "record"
            Description     =   "Record Log"
            Object.ToolTipText     =   "Record a query log."
            ImageKey        =   "record"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "stop"
            Description     =   "Stop Recording"
            Object.ToolTipText     =   "Stop recording."
            ImageKey        =   "stop"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sb 
      Align           =   2  'Align Bottom
      Height          =   288
      Left            =   0
      TabIndex        =   2
      Top             =   6372
      Width           =   9684
      _ExtentX        =   17082
      _ExtentY        =   508
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5675
            MinWidth        =   2
            Text            =   "Ready"
            TextSave        =   "Ready"
            Key             =   "info"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1101
            MinWidth        =   2
            Text            =   "0 Secs."
            TextSave        =   "0 Secs."
            Key             =   "timer"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   3048
            MinWidth        =   2
            Text            =   "Object: Not Connected"
            TextSave        =   "Object: Not Connected"
            Key             =   "currentobject"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   3450
            MinWidth        =   2
            Text            =   "Database: Not Connected"
            TextSave        =   "Database: Not Connected"
            Key             =   "currentdb"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   3260
            MinWidth        =   2
            Text            =   "Schema: Not Connected"
            TextSave        =   "Schema: Not Connected"
            Key             =   "currentns"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList il 
      Left            =   4320
      Top             =   1080
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   29
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9AB2
            Key             =   "aggregate"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A184
            Key             =   "check"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A856
            Key             =   "column"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":AF28
            Key             =   "function"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B5FA
            Key             =   "group"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":BCCC
            Key             =   "index"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C266
            Key             =   "indexcolumn"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C938
            Key             =   "foreignkey"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D00A
            Key             =   "language"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D6DC
            Key             =   "operator"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":DDAE
            Key             =   "property"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E348
            Key             =   "relationship"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E4A2
            Key             =   "rule"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":EB74
            Key             =   "server"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":ECCE
            Key             =   "sequence"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":F3A0
            Key             =   "table"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":FA72
            Key             =   "trigger"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10144
            Key             =   "type"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10816
            Key             =   "user"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10970
            Key             =   "view"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11042
            Key             =   "hiproperty"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":115DC
            Key             =   "database"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11736
            Key             =   "closeddatabase"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11890
            Key             =   "baddatabase"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":119EA
            Key             =   "statistics"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":125BC
            Key             =   "domain"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12C8E
            Key             =   "namespace"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13860
            Key             =   "cast"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14432
            Key             =   "conversion"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tv 
      Height          =   5235
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   3390
      _ExtentX        =   5990
      _ExtentY        =   9229
      _Version        =   393217
      Indentation     =   441
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "il"
      Appearance      =   1
   End
   Begin TabDlg.SSTab prop 
      Height          =   3255
      Left            =   3720
      TabIndex        =   4
      Top             =   720
      Width           =   5895
      _ExtentX        =   10393
      _ExtentY        =   5736
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "&Properties"
      TabPicture(0)   =   "frmMain.frx":14D0C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lv"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Statistics"
      TabPicture(1)   =   "frmMain.frx":14D28
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "sv"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&Dependencies"
      TabPicture(2)   =   "frmMain.frx":14D44
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "tvDep"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "&Locks"
      TabPicture(3)   =   "frmMain.frx":14D60
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lvLock"
      Tab(3).ControlCount=   1
      Begin MSComctlLib.ListView lv 
         Height          =   2655
         Left            =   45
         TabIndex        =   5
         Top             =   45
         Width           =   5940
         _ExtentX        =   10478
         _ExtentY        =   4678
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "il"
         SmallIcons      =   "il"
         ColHdrIcons     =   "il"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView sv 
         Height          =   2655
         Left            =   -74955
         TabIndex        =   6
         Top             =   45
         Width           =   5940
         _ExtentX        =   10478
         _ExtentY        =   4678
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "il"
         SmallIcons      =   "il"
         ColHdrIcons     =   "il"
         ForeColor       =   -2147483640
         BackColor       =   -2147483633
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.TreeView tvDep 
         Height          =   2655
         Left            =   -74955
         TabIndex        =   7
         Top             =   45
         Width           =   5940
         _ExtentX        =   10478
         _ExtentY        =   4678
         _Version        =   393217
         Indentation     =   441
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "il"
         Appearance      =   1
      End
      Begin MSComctlLib.ListView lvLock 
         Height          =   2655
         Left            =   -74955
         TabIndex        =   8
         Top             =   45
         Width           =   5940
         _ExtentX        =   10478
         _ExtentY        =   4678
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "il"
         SmallIcons      =   "il"
         ColHdrIcons     =   "il"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Image splVertical 
      DragMode        =   1  'Automatic
      Height          =   5325
      Left            =   3600
      MousePointer    =   9  'Size W E
      Top             =   585
      Width           =   45
   End
   Begin VB.Image splHorizontal 
      DragMode        =   1  'Automatic
      Height          =   45
      Left            =   3825
      MousePointer    =   7  'Size N S
      Top             =   4095
      Width           =   5760
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileConnect 
         Caption         =   "&Connect..."
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileChangePassword 
         Caption         =   "Change &Password..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSaveDefinition 
         Caption         =   "&Save Definition..."
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveDBSchema 
         Caption         =   "S&ave DB Schema..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Enabled         =   0   'False
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Enabled         =   0   'False
         Shortcut        =   +{INSERT}
      End
   End
   Begin VB.Menu mnuPlugins 
      Caption         =   "&Plugins"
      Begin VB.Menu mnuPluginsPlg 
         Caption         =   "You should never see this!"
         Index           =   0
      End
      Begin VB.Menu mnuPluginsPlg 
         Caption         =   "Plugin1"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPluginsPlg 
         Caption         =   "Plugin2"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPluginsPlg 
         Caption         =   "Plugin3"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPluginsPlg 
         Caption         =   "Plugin4"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPluginsPlg 
         Caption         =   "Plugin5"
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPluginsPlg 
         Caption         =   "Plugin6"
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPluginsPlg 
         Caption         =   "Plugin7"
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPluginsPlg 
         Caption         =   "Plugin8"
         Index           =   8
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPluginsPlg 
         Caption         =   "Plugin9"
         Index           =   9
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPluginsPlg 
         Caption         =   "Plugin10"
         Index           =   10
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPluginsPlg 
         Caption         =   "Plugin11"
         Index           =   11
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPluginsPlg 
         Caption         =   "Plugin12"
         Index           =   12
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPluginsPlg 
         Caption         =   "Plugin13"
         Index           =   13
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPluginsPlg 
         Caption         =   "Plugin14"
         Index           =   14
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPluginsPlg 
         Caption         =   "Plugin15"
         Index           =   15
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPluginsPlg 
         Caption         =   "Plugin16"
         Index           =   16
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPluginsPlg 
         Caption         =   "Plugin17"
         Index           =   17
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPluginsPlg 
         Caption         =   "Plugin18"
         Index           =   18
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPluginsPlg 
         Caption         =   "Plugin19"
         Index           =   19
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPluginsPlg 
         Caption         =   "Plugin20"
         Index           =   20
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuToolsFindObject 
         Caption         =   "&Find Object"
         Enabled         =   0   'False
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuToolsUpgradeWizard 
         Caption         =   "&Upgrade Wizard..."
      End
      Begin VB.Menu mnuToolsSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsOptions 
         Caption         =   "&Options..."
         Shortcut        =   ^O
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewSystemObjects 
         Caption         =   "System Objects"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewShowDefinitionPane 
         Caption         =   "Show &Definition Pane"
      End
      Begin VB.Menu mnuViewShowLogWindow 
         Caption         =   "Show &Log Window"
      End
      Begin VB.Menu mnuViewShowStatusBar 
         Caption         =   "Show &Status Bar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewShowToolBar 
         Caption         =   "Show &Tool Bar"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents"
      End
      Begin VB.Menu mnuHelpTipOfTheDay 
         Caption         =   "&Tip of the Day"
      End
      Begin VB.Menu mnuHelpSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "&Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuPopupConnect 
         Caption         =   "&Connect to server..."
      End
      Begin VB.Menu mnuPopupRefresh 
         Caption         =   "&Refresh below selection"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuPopupSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupCopy 
         Caption         =   "&Copy"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuPopupPaste 
         Caption         =   "&Paste"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuPopupSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupCreate 
         Caption         =   "&Create object"
         Enabled         =   0   'False
         Begin VB.Menu mnuPopupCreateAggregate 
            Caption         =   "&Aggregate..."
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuPopupCreateCast 
            Caption         =   "&Cast..."
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuPopupCreateConversion 
            Caption         =   "&Conversion..."
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuPopupCreateDatabase 
            Caption         =   "&Database..."
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuPopupCreateDomain 
            Caption         =   "Do&main..."
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuPopupCreateFunction 
            Caption         =   "&Function..."
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuPopupCreateGroup 
            Caption         =   "&Group..."
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuPopupCreateIndex 
            Caption         =   "&Index..."
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuPopupCreateLanguage 
            Caption         =   "&Language..."
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuPopupCreateOperator 
            Caption         =   "&Operator..."
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuPopupCreateRule 
            Caption         =   "&Rule..."
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuPopupCreateNamespace 
            Caption         =   "&Schema"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuPopupCreateSequence 
            Caption         =   "&Sequence..."
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuPopupCreateTable 
            Caption         =   "&Table..."
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuPopupCreateTrigger 
            Caption         =   "Tri&gger..."
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuPopupCreateType 
            Caption         =   "T&ype..."
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuPopupCreateUser 
            Caption         =   "&User..."
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuPopupCreateView 
            Caption         =   "&View..."
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnuPopupDrop 
         Caption         =   "&Drop object"
         Enabled         =   0   'False
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuPopupProperties 
         Caption         =   "&Properties..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuPopupSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupSQL 
         Caption         =   "&SQL..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuPopupViewData 
         Caption         =   "&View Data"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuPopupSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupResetStatistics 
         Caption         =   "R&eset Database Statistics"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuPopupReindex 
         Caption         =   "Re&Index"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuPopupVacuum 
         Caption         =   "Vac&uum"
         Enabled         =   0   'False
         Begin VB.Menu mnuPopupVacuumVacuum 
            Caption         =   "&Vacuum"
         End
         Begin VB.Menu mnuPopupVacuumAnalyse 
            Caption         =   "Vacuum &Analyse"
         End
      End
      Begin VB.Menu mnuPopupSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupRecordLog 
         Caption         =   "&Record Log..."
      End
      Begin VB.Menu mnuPopupStopRecording 
         Caption         =   "&Stop Recording"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the Artistic Licence
'
' frmMain.frm - The primary form.

Option Explicit

'The Global Server Object. This must be in a form to be declared WithEvents
Public WithEvents svr As pgServer
Attribute svr.VB_VarHelpID = -1

Private Sub Form_Resize()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.Form_Resize()", etFullDebug

  If Me.WindowState <> 1 Then
    txtDefinition.Minimise
    Resize splVertical.Left, splHorizontal.Top
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.Form_Resize"
End Sub

Public Sub Resize(VPos As Single, HPos As Single)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.Resize(" & HPos & ", " & VPos & ")", etFullDebug

Dim siTop As Single
Dim siLeft As Single
Dim siHeight As Single
Dim siWidth As Single
  
  'Check the form size
  If Me.Height < 4500 Then Me.Height = 4500
  If Me.Width < 5000 Then Me.Width = 5000
  
  'Size to the form
  If tb.Visible Then
    siTop = tb.Height
  Else
    siTop = 0
  End If
  siLeft = 0
  If sb.Visible Then
    siHeight = Me.ScaleHeight - sb.Height
  Else
    siHeight = Me.ScaleHeight
  End If
  siWidth = Me.ScaleWidth
  
  'Set the Min/Max positions
  If VPos < siLeft + 1000 Then VPos = siLeft + 1000
  If VPos > siWidth - 1000 Then VPos = siWidth - 1000
  If HPos < siTop + 1000 Then HPos = siTop + 1000
  If HPos > siHeight - 1000 Then HPos = siHeight - 1000
  
  'Set Verticals
  tv.Top = siTop
  tv.Height = siHeight - tv.Top
  
  prop.Top = siTop
  If txtDefinition.Visible And ((HPos - prop.Top) > 0) Then
    prop.Height = HPos - prop.Top
  Else
    prop.Height = tv.Height
  End If
  
  txtDefinition.Top = HPos + 50
  txtDefinition.Height = siHeight - txtDefinition.Top
  
  splVertical.Top = -(siHeight * 2)
  splVertical.Height = siHeight * 5
  splVertical.Left = VPos
  
  'Set Horizontals
  tv.Left = siLeft
  tv.Width = VPos - tv.Left
  
  prop.Left = VPos + 50
  prop.Width = siWidth - prop.Left
  
  txtDefinition.Left = prop.Left
  txtDefinition.Width = prop.Width
  
  splHorizontal.Left = -(siWidth * 2)
  splHorizontal.Width = siWidth * 5
  splHorizontal.Top = HPos
  
  'Set the properties listview size
  lv.Width = prop.Width - 45
  lv.Height = prop.Height - 450
  sv.Width = prop.Width - 45
  sv.Height = prop.Height - 450
  tvDep.Width = prop.Width - 45
  tvDep.Height = prop.Height - 450
  lvLock.Width = prop.Width - 45
  lvLock.Height = prop.Height - 450
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.Resize"
End Sub

Private Sub lv_KeyDown(KeyCode As Integer, Shift As Integer)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.lv_KeyDown(" & KeyCode & "," & Shift & ")", etFullDebug
  
  If KeyCode = vbKeyReturn And Shift = vbAltMask Then
    If mnuPopupProperties.Enabled Then mnuPopupProperties_Click
  ElseIf KeyCode = vbKeyDelete And Shift = 0 Then
    If mnuPopupDrop.Enabled Then mnuPopupDrop_Click
  End If

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.lv_KeyDown"
End Sub

Private Sub mnuEditCopy_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuEditCopy_Click()", etFullDebug

  CopyObjDb
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuEditCopy_Click"
End Sub

Private Sub mnuEditPaste_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuEditPaste_Click()", etFullDebug

  PasteObjDb
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuEditPaste_Click"
End Sub

Private Sub mnuPopupCopy_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuPopupCopy_Click()", etFullDebug

  CopyObjDb
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuPopupCopy_Click"
End Sub

Private Sub mnuPopupPaste_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuPopupPaste_Click()", etFullDebug

  PasteObjDb
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuPopupPaste_Click"
End Sub

Private Sub mnuPopupCreateDomain_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuPopupCreateDomain_Click()", etFullDebug

Dim objDomainForm As New frmDomain

  Load objDomainForm
  objDomainForm.Initialise ctx.CurrentDB, ctx.CurrentNS
  objDomainForm.Show
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuPopupCreateDomain_Click"
End Sub

Private Sub mnuPopupRecordLog_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuPopupRecordLog_Click()", etFullDebug

  Load frmRecordLog
  If InStr(1, Command, "-wine") <> 0 Then
    frmRecordLog.Show
  Else
    frmRecordLog.Show vbModal, Me
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuPopupRecordLog_Click"
End Sub

Private Sub mnuPopupStopRecording_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuPopupStopRecording_Click()", etFullDebug

  svr.LogEvent "Stopping recording query log.", etMiniDebug
  svr.UserLog = False
  tb.Buttons("record").Enabled = True
  tb.Buttons("stop").Enabled = False
  mnuPopupRecordLog.Enabled = True
  mnuPopupStopRecording = False
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuPopupStopRecording_Click"
End Sub

Private Sub mnuToolsFindObject_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuToolsFindObjDb_Click()", etFullDebug
Dim objFindForm As New frmFind

  Load objFindForm
  objFindForm.Initialise
  objFindForm.Show
  Exit Sub

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuToolsFindObject_Click"
End Sub

Private Sub tv_DragDrop(Source As Control, X As Single, Y As Single)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tv_DragDrop(" & Source.Name & ", " & X & ", " & Y & ")", etFullDebug

  If Source.Name = "splVertical" Then
    Resize tv.Left + X, splHorizontal.Top
  ElseIf Source.Name = "splHorizontal" Then
    Resize splVertical.Left, tv.Top + Y
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tv_DragDrop"
End Sub

Private Sub prop_DragDrop(Source As Control, X As Single, Y As Single)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.prop_DragDrop(" & Source.Name & ", " & X & ", " & Y & ")", etFullDebug

  If Source.Name = "splVertical" Then
    Resize prop.Left + X, splHorizontal.Top
  ElseIf Source.Name = "splHorizontal" Then
    Resize splVertical.Left, prop.Top + Y
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.prop_DragDrop"
End Sub

Private Sub lv_DragDrop(Source As Control, X As Single, Y As Single)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.lv_DragDrop(" & Source.Name & ", " & X & ", " & Y & ")", etFullDebug

  If Source.Name = "splVertical" Then
    Resize lv.Left + prop.Left + X, splHorizontal.Top
  ElseIf Source.Name = "splHorizontal" Then
    Resize splVertical.Left, lv.Top + prop.Top + Y
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.lv_DragDrop"
End Sub

Private Sub sv_DragDrop(Source As Control, X As Single, Y As Single)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.sv_DragDrop(" & Source.Name & ", " & X & ", " & Y & ")", etFullDebug

  If Source.Name = "splVertical" Then
    Resize sv.Left + prop.Left + X, splHorizontal.Top
  ElseIf Source.Name = "splHorizontal" Then
    Resize splVertical.Left, sv.Top + prop.Top + Y
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.sv_DragDrop"
End Sub

Private Sub tv_KeyDown(KeyCode As Integer, Shift As Integer)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tv_KeyDown(" & KeyCode & "," & Shift & ")", etFullDebug
  
  If KeyCode = vbKeyReturn And Shift = vbAltMask Then
    If mnuPopupProperties.Enabled Then mnuPopupProperties_Click
  ElseIf KeyCode = vbKeyDelete And Shift = 0 Then
    If mnuPopupDrop.Enabled Then mnuPopupDrop_Click
  End If

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tv_KeyDown"
End Sub

Private Sub txtDefinition_DragDrop(Source As Control, X As Single, Y As Single)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.txtDefinition_DragDrop(" & Source.Name & ", " & X & ", " & Y & ")", etFullDebug

  If Source.Name = "splVertical" Then
    Resize txtDefinition.Left + X, splHorizontal.Top
  ElseIf Source.Name = "splHorizontal" Then
    Resize splVertical.Left, txtDefinition.Top + Y
  End If
    
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.txtDefinition_DragDrop"
End Sub

Private Sub lv_DblClick()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.lv_DblClick()", etFullDebug

  mnuPopupProperties_Click
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.lv_DblClick"
End Sub

Private Sub lv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.lv_MouseUp(" & Button & ", " & Shift & ", " & X & ", " & Y & ")", etFullDebug

  If Button = 2 Then PopupMenu frmMain.mnuPopup
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.lv_MouseUp"
End Sub

Private Sub mnuHelpContents_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuHelpContents_Click()", etFullDebug

  HtmlHelp hwnd, App.Path & "\" & "help\pgadmin2.chm", HH_DISPLAY_TOPIC, 0

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuHelpContents_Click"
End Sub

Private Sub mnuHelpTipOfTheDay_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuHelpTipOfTheDay_Click()", etFullDebug

  Load frmTip
  frmTip.Show

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuHelpTipOfTheDay_Click"
End Sub

Private Sub mnuPluginsPlg_Click(Index As Integer)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuPluginsPlg_Click(" & Index & ")", etFullDebug

Dim szPlugin As String

  If Index = 0 Then Exit Sub
  szPlugin = Left(mnuPluginsPlg(Index).Caption, Len(mnuPluginsPlg(Index).Caption) - 3)
  svr.LogEvent "Executing Plugin: " & plg(szPlugin).Description & " v" & plg(szPlugin).Version, etMiniDebug
  plg(szPlugin).Execute sb, svr

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuPluginsPlg_Click"
End Sub

Private Sub mnuToolsUpgradeWizard_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuToolsUpgradeWizard_Click()", etFullDebug

  Load frmUpgradeWizard
  frmUpgradeWizard.Show

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuToolsUpgradeWizard_Click"
End Sub

Private Sub mnuViewSystemObjects_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuViewSystemObjects_Click()", etFullDebug

Dim objNode As Node

  If tv.Nodes.Count > 0 Then
    If MsgBox("This will cause the treeview to be collapsed and rebuilt. Are you sure you wish to continue?", vbQuestion + vbYesNo, "Collapse Treeview") = vbNo Then Exit Sub
  End If
  
  If mnuViewSystemObjects.Checked = False Then
    ctx.IncludeSys = True
    RegWrite HKEY_CURRENT_USER, "Software\" & App.Title, "Hide System Objects", regString, "N"
    mnuViewSystemObjects.Checked = True
  Else
    ctx.IncludeSys = False
    RegWrite HKEY_CURRENT_USER, "Software\" & App.Title, "Hide System Objects", regString, "Y"
    mnuViewSystemObjects.Checked = False
  End If
  
  'Clear all nodes, and re-create the server node
  If tv.Nodes.Count > 0 Then
    svr.Connect
    tv.Nodes.Clear
    Set objNode = frmMain.tv.Nodes.Add(, , "SVR-" & GetID, svr.Server, "server")
    tv_NodeClick objNode
    objNode.Expanded = True
  End If

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuViewSystemObjects_Click"
End Sub

Private Sub tv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tv_MouseUp(" & Button & ", " & Shift & ", " & X & ", " & Y & ")", etFullDebug

  If Button = 2 Then PopupMenu frmMain.mnuPopup

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tv_MouseUp"
End Sub

Private Sub mnuFileChangePassword_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuFileChangePassword_Click()", etFullDebug

  Load frmPassword
  If InStr(1, Command, "-wine") <> 0 Then
    frmPassword.Show
  Else
    frmPassword.Show vbModal, Me
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuFileChangePassword_Click"
End Sub

Private Sub mnuFileSaveDbSchema_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuFileSaveDBSchema_Click()", etFullDebug

Dim fNum As Integer
Dim bResetSequences As Boolean

  'Reset Sequences
  If MsgBox("Do you wish to reset Sequence values to zero in the output file?", vbQuestion + vbYesNo, "Reset Sequences") = vbYes Then bResetSequences = True
  
  With cdlg
    .DialogTitle = "Save Database Schema"
    .Filter = "SQL Scripts (*.sql)|*.sql"
    .CancelError = True
    .ShowSave
  End With
  If cdlg.FileName = "" Then
    MsgBox "No filename specified - Database Schema not saved.", vbExclamation, "Warning"
    Exit Sub
  End If
  If Dir(cdlg.FileName) <> "" Then
    If MsgBox("File exists - overwrite?", vbYesNo + vbQuestion, "Overwrite File") = vbNo Then mnuFileSaveDbSchema_Click
  End If
  fNum = FreeFile
  svr.LogEvent "Writing " & cdlg.FileName, etMiniDebug
  Open cdlg.FileName For Output As #fNum
  StartMsg "Saving Database Schema..."
  Print #fNum, "-- " & App.Title & " v" & App.Major & "." & App.Minor & "." & App.Revision & " Database Schema Dump" & vbCrLf
  Print #fNum, svr.Databases(ctx.CurrentDB).Schema(bResetSequences)
  EndMsg
  Close #fNum

  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number = 32755 Then
    svr.LogEvent "Save Database Schema operation cancelled.", etMiniDebug
    Exit Sub
  End If
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuFileSaveDBSchema_Click"
End Sub

Private Sub mnuFileSaveDefinition_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuFileSaveDefinition_Click()", etFullDebug

Dim fNum As Integer

  With cdlg
    .DialogTitle = "Save Object Definition"
    .Filter = "SQL Scripts (*.sql)|*.sql"
    .CancelError = True
    .ShowSave
  End With
  If cdlg.FileName = "" Then
    MsgBox "No filename specified - Object Definition not saved.", vbExclamation, "Warning"
    Exit Sub
  End If
  If Dir(cdlg.FileName) <> "" Then
    If MsgBox("File exists - overwrite?", vbYesNo + vbQuestion, "Overwrite File") = vbNo Then mnuFileSaveDefinition_Click
  End If
  fNum = FreeFile
  svr.LogEvent "Writing " & cdlg.FileName, etMiniDebug
  Open cdlg.FileName For Output As #fNum
  StartMsg "Saving Object Definition..."
  Print #fNum, txtDefinition.Text
  EndMsg
  Close #fNum

  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number = 32755 Then
    svr.LogEvent "Save Object Definition operation cancelled.", etMiniDebug
    Exit Sub
  End If
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuFileSaveDefinition_Click"
End Sub

Private Sub mnuHelpAbout_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuHelpAbout_Click()", etFullDebug

  Load frmAbout
  If InStr(1, Command, "-wine") <> 0 Then
    frmAbout.Show
  Else
    frmAbout.Show vbModal, Me
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuHelpAbout_Click"
End Sub

Private Sub mnuToolsOptions_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuToolsOptions_Click()", etFullDebug

  Load frmOptions
  If InStr(1, Command, "-wine") <> 0 Then
    frmOptions.Show
  Else
    frmOptions.Show vbModal, Me
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuToolsOptions_Click"
End Sub

Private Sub mnuViewShowDefinitionPane_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuViewShowDefinitionPane_Click()", etFullDebug

  txtDefinition.Text = ""
  If mnuViewShowDefinitionPane.Checked = True Then
    RegWrite HKEY_CURRENT_USER, "Software\" & App.Title, "Show Definition Pane", regString, "N"
    mnuViewShowDefinitionPane.Checked = False
    txtDefinition.Visible = False
  Else
    RegWrite HKEY_CURRENT_USER, "Software\" & App.Title, "Show Definition Pane", regString, "Y"
    mnuViewShowDefinitionPane.Checked = True
    txtDefinition.Visible = True
  End If
  Resize splVertical.Left, splHorizontal.Top
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuViewShowDefinitionPane_Click"
End Sub

Private Sub mnuViewShowLogWindow_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuViewShowLogWindow_Click()", etFullDebug

  If mnuViewShowLogWindow.Checked = True Then
    ctx.LogView = False
    frmLog.Hide
    mnuViewShowLogWindow.Checked = False
  Else
    ctx.LogView = True
    frmLog.Show
    mnuViewShowLogWindow.Checked = True
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuViewShowLogWindow_Click"
End Sub

Private Sub mnuViewShowStatusBar_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuViewShowStatusBar_Click()", etFullDebug

  If mnuViewShowStatusBar.Checked = True Then
    RegWrite HKEY_CURRENT_USER, "Software\" & App.Title, "Show Status Bar", regString, "N"
    mnuViewShowStatusBar.Checked = False
    sb.Visible = False
  Else
    RegWrite HKEY_CURRENT_USER, "Software\" & App.Title, "Show Status Bar", regString, "Y"
    mnuViewShowStatusBar.Checked = True
    sb.Visible = True
  End If
  Resize splVertical.Left, splHorizontal.Top
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuViewShowStatusBar_Click"
End Sub

Private Sub mnuViewShowToolBar_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuViewShowToolBar_Click()", etFullDebug

  If mnuViewShowToolBar.Checked = True Then
    RegWrite HKEY_CURRENT_USER, "Software\" & App.Title, "Show Tool Bar", regString, "N"
    mnuViewShowToolBar.Checked = False
    tb.Visible = False
  Else
    RegWrite HKEY_CURRENT_USER, "Software\" & App.Title, "Show Tool Bar", regString, "Y"
    mnuViewShowToolBar.Checked = True
    tb.Visible = True
  End If
  Resize splVertical.Left, splHorizontal.Top
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuViewShowToolBar_Click"
End Sub

Private Sub mnuPopupConnect_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuPopupConnect_Click()", etFullDebug

  Load frmConnect
  frmConnect.Load_Defaults
  If InStr(1, Command, "-wine") <> 0 Then
    frmConnect.Show
  Else
    frmConnect.Show vbModal, Me
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuPopupConnect_Click"
End Sub

Private Sub mnuPopupRefresh_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuPopupRefresh_Click()", etFullDebug

Dim objNode As Node

  'We refresh from collection nodes, or the Server. If anything else is selected, refresh from the parent
  If (Left(tv.SelectedItem.Key, 4) = "SVR-") Or (Mid(tv.SelectedItem.Key, 4, 1) = "+") Then
    Set objNode = tv.SelectedItem
  Else
    Set objNode = tv.SelectedItem.Parent
  End If
  
  'Now refresh the required part of the svr object
  Select Case Left(objNode.Key, 4)
    Case "SVR-"
      svr.Refresh
    Case "DAT+"
      svr.Databases.Refresh
    Case "GRP+"
      svr.Groups.Refresh
    Case "USR+"
      svr.Users.Refresh
    Case "CST+"
      svr.Databases(objNode.Parent.Text).Casts.Refresh
    Case "LNG+"
      svr.Databases(objNode.Parent.Text).Languages.Refresh
    Case "NSP+"
      svr.Databases(objNode.Parent.Text).Namespaces.Refresh
    Case "AGG+"
      svr.Databases(objNode.Parent.Parent.Parent.Text).Namespaces(objNode.Parent.Text).Aggregates.Refresh
    Case "CNV+"
      svr.Databases(objNode.Parent.Parent.Parent.Text).Namespaces(objNode.Parent.Text).Conversions.Refresh
    Case "DOM+"
      svr.Databases(objNode.Parent.Parent.Parent.Text).Namespaces(objNode.Parent.Text).Domains.Refresh
    Case "FNC+"
      svr.Databases(objNode.Parent.Parent.Parent.Text).Namespaces(objNode.Parent.Text).Functions.Refresh
    Case "OPR+"
      svr.Databases(objNode.Parent.Parent.Parent.Text).Namespaces(objNode.Parent.Text).Operators.Refresh
    Case "SEQ+"
      svr.Databases(objNode.Parent.Parent.Parent.Text).Namespaces(objNode.Parent.Text).Sequences.Refresh
    Case "TBL+"
      svr.Databases(objNode.Parent.Parent.Parent.Text).Namespaces(objNode.Parent.Text).Tables.Refresh
    Case "CHK+"
      svr.Databases(objNode.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(objNode.Parent.Parent.Parent.Text).Tables(objNode.Parent.Text).Checks.Refresh
    Case "COL+"
      svr.Databases(objNode.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(objNode.Parent.Parent.Parent.Text).Tables(objNode.Parent.Text).Columns.Refresh
    Case "FKY+"
      svr.Databases(objNode.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(objNode.Parent.Parent.Parent.Text).Tables(objNode.Parent.Text).ForeignKeys.Refresh
    Case "REL+"
      svr.Databases(objNode.Parent.Parent.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(objNode.Parent.Parent.Parent.Parent.Parent.Text).Tables(objNode.Parent.Parent.Parent.Text).ForeignKeys(objNode.Parent.Text).Relationships.Refresh
    Case "IND+"
      svr.Databases(objNode.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(objNode.Parent.Parent.Parent.Text).Tables(objNode.Parent.Text).Indexes.Refresh
    Case "RUL+"
      'verify if rule is for table or view
      If svr.Databases(objNode.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(objNode.Parent.Parent.Parent.Text).Tables.Exists(objNode.Parent.Text) Then
        svr.Databases(objNode.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(objNode.Parent.Parent.Parent.Text).Tables(objNode.Parent.Text).Rules.Refresh
      ElseIf svr.Databases(objNode.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(objNode.Parent.Parent.Parent.Text).Views.Exists(objNode.Parent.Text) Then
        svr.Databases(objNode.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(objNode.Parent.Parent.Parent.Text).Views(objNode.Parent.Text).Rules.Refresh
      End If
    Case "TRG+"
      svr.Databases(objNode.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(objNode.Parent.Parent.Parent.Text).Tables(objNode.Parent.Text).Triggers.Refresh
    Case "TYP+"
      svr.Databases(objNode.Parent.Parent.Parent.Text).Namespaces(objNode.Parent.Text).Types.Refresh
    Case "VIE+"
      svr.Databases(objNode.Parent.Parent.Parent.Text).Namespaces(objNode.Parent.Text).Views.Refresh
  End Select
  
  'Clear the child nodes
  While objNode.Children > 0
    tv.Nodes.Remove objNode.Child.Index
  Wend
  'Simulate a node click to refresh the immediate children
  tv_NodeClick objNode
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuPopupRefresh_Click"
End Sub

Private Sub mnuPopupResetStatistics_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuPopupResetStat_Click()", etFullDebug

  'reset statistic
  If MsgBox("Are you sure you wish to reset the database statistics?", vbApplicationModal + vbYesNo + vbQuestion) = vbYes Then
    svr.Databases(ctx.CurrentDB).Execute "SELECT pg_stat_reset()"
  End If
  
  'Reset the stats etc.
  tv_NodeClick tv.SelectedItem
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuPopupResetStatistics_Click"
End Sub

Private Sub mnuPopupReindex_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuPopupReindex_Click()", etFullDebug

  Reindex
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuPopupReindex_Click"
End Sub

Private Sub mnuPopupDrop_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuPopupDrop_Click()", etFullDebug

  Drop
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuPopupDrop_Click"
End Sub

Private Sub mnuPopupProperties_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuPopupProperties_Click()", etFullDebug
      
      If ctx.CurrentObject Is Nothing Then Exit Sub

      Select Case ctx.CurrentObject.ObjectType
        Case "Aggregate"
          Dim objAggregateForm As New frmAggregate
          Load objAggregateForm
          objAggregateForm.Initialise ctx.CurrentDB, ctx.CurrentNS, ctx.CurrentObject
          objAggregateForm.Show
          
        Case "Cast"
          Dim objCastForm As New frmCast
          Load objCastForm
          objCastForm.Initialise ctx.CurrentDB, ctx.CurrentObject
          objCastForm.Show
          
        Case "Column"
          Dim objColumnForm As New frmColumn
          Load objColumnForm
          objColumnForm.Initialise ctx.CurrentDB, ctx.CurrentNS, "MP", ctx.CurrentObject
          objColumnForm.Show
          
        Case "Database"
          Dim objDatabaseForm As New frmDatabase
          Load objDatabaseForm
          objDatabaseForm.Initialise ctx.CurrentObject
          objDatabaseForm.Show
          
        Case "Domain"
          Dim objDomainForm As New frmDomain
          Load objDomainForm
          objDomainForm.Initialise ctx.CurrentDB, ctx.CurrentNS, ctx.CurrentObject
          objDomainForm.Show
          
        Case "Conversion"
          Dim objConversionForm As New frmConversion
          Load objConversionForm
          objConversionForm.Initialise ctx.CurrentDB, ctx.CurrentNS, ctx.CurrentObject
          objConversionForm.Show
          
        Case "Foreign Key"
          Dim objForeignKeyForm As New frmForeignKey
          Load objForeignKeyForm
          objForeignKeyForm.Initialise ctx.CurrentDB, ctx.CurrentNS, "MP", ctx.CurrentObject
          objForeignKeyForm.Show
          
        Case "Function"
          Dim objFunctionForm As New frmFunction
          Load objFunctionForm
          objFunctionForm.Initialise ctx.CurrentDB, ctx.CurrentNS, ctx.CurrentObject
          objFunctionForm.Show

        Case "Group"
          Dim objGroupForm As New frmGroup
          Load objGroupForm
          objGroupForm.Initialise ctx.CurrentObject
          objGroupForm.Show
    
        Case "Index"
          Dim objIndexForm As New frmIndex
          Load objIndexForm
          objIndexForm.Initialise ctx.CurrentDB, ctx.CurrentNS, ctx.CurrentObject
          objIndexForm.Show
          
        Case "Language"
          Dim objLanguageForm As New frmLanguage
          Load objLanguageForm
          objLanguageForm.Initialise ctx.CurrentDB, ctx.CurrentObject
          objLanguageForm.Show
          
        Case "Schema"
          Dim objNamespaceForm As New frmNamespace
          Load objNamespaceForm
          objNamespaceForm.Initialise ctx.CurrentDB, ctx.CurrentObject
          objNamespaceForm.Show
          
        Case "Operator"
          Dim objOperatorForm As New frmOperator
          Load objOperatorForm
          objOperatorForm.Initialise ctx.CurrentDB, ctx.CurrentNS, ctx.CurrentObject
          objOperatorForm.Show
          
        Case "Rule"
          Dim objRuleForm As New frmRule
          Load objRuleForm
          objRuleForm.Initialise ctx.CurrentDB, ctx.CurrentNS, ctx.CurrentObject
          objRuleForm.Show
          
        Case "Server"
          Dim objServerForm As New frmServer
          Load objServerForm
          objServerForm.Initialise ctx.CurrentObject
          objServerForm.Show
          
        Case "Sequence"
          Dim objSequenceForm As New frmSequence
          Load objSequenceForm
          objSequenceForm.Initialise ctx.CurrentDB, ctx.CurrentNS, ctx.CurrentObject
          objSequenceForm.Show

        Case "Table"
          Dim objTableForm As New frmTable
          Load objTableForm
          objTableForm.Initialise ctx.CurrentDB, ctx.CurrentNS, ctx.CurrentObject
          objTableForm.Show
          
        Case "Trigger"
          Dim objTriggerForm As New frmTrigger
          Load objTriggerForm
          objTriggerForm.Initialise ctx.CurrentDB, ctx.CurrentNS, ctx.CurrentObject
          objTriggerForm.Show
          
        Case "Type"
          Dim objTypeForm As New frmType
          Load objTypeForm
          objTypeForm.Initialise ctx.CurrentDB, ctx.CurrentNS, ctx.CurrentObject
          objTypeForm.Show
          
        Case "User"
          Dim objUserForm As New frmUser
          Load objUserForm
          objUserForm.Initialise ctx.CurrentObject
          objUserForm.Show
          
        Case "View"
          Dim objViewForm As New frmView
          Load objViewForm
          objViewForm.Initialise ctx.CurrentDB, ctx.CurrentNS, ctx.CurrentObject
          objViewForm.Show
          
        Case Else
          MsgBox "Cannot display properties for the current object.", vbExclamation, "Error"
      End Select
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuPopupProperties_Click"
End Sub

Private Sub mnuPopupSQL_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuPopupSQL_Click()", etFullDebug
  
Dim Y As Integer
Dim X As Integer

  Y = 1
  For X = 0 To Forms.Count - 1
    If Forms(X).Name = "frmSQLInput" Then
      Y = Val(Forms(X).Tag) + 1
    End If
  Next
  Dim objSQLInputForm As New frmSQLInput
  Load objSQLInputForm
  objSQLInputForm.Tag = Y
  objSQLInputForm.Caption = "SQL " & Y & ": " & ctx.CurrentDB & " ()"
  objSQLInputForm.Show
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuPopupSQL_Click"
End Sub

Private Sub mnuPopupViewData_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuPopupViewData_Click()", etFullDebug
  
Dim objOutputForm As New frmSQLOutput
Dim rsQuery As New Recordset
Dim iMsgBoxResult As VbMsgBoxResult
Dim iLimit As Long
Dim szLimit As String
Dim szTemp As String

  'count row
  StartMsg "Counting Records..."
  Set rsQuery = frmMain.svr.Databases(ctx.CurrentDB).Execute("SELECT count(*) AS count FROM " & ctx.CurrentObject.FormattedID)
  EndMsg
  
  'verify limit output
  szLimit = ""
  iLimit = Val(RegRead(HKEY_CURRENT_USER, "Software\" & App.Title, "Row Limit", "1000"))
  If Not rsQuery.EOF Then
    If rsQuery!Count > iLimit Then
      iMsgBoxResult = MsgBox("The query will return " & rsQuery!Count & " rows. Do you wish to LIMIT the output?", vbApplicationModal + vbYesNoCancel + vbQuestion, "Row limit")
      If iMsgBoxResult = vbCancel Then
        Exit Sub
      ElseIf iMsgBoxResult = vbYes Then
        iLimit = Val(InputBox("Enter a row limit" & vbCrLf & "The table or view contains " & rsQuery!Count & " row(s)", "Row limit", iLimit))
        szLimit = " LIMIT " & iLimit
        RegWrite HKEY_CURRENT_USER, "Software\" & App.Title, "Row Limit", regString, iLimit
      End If
    End If
  End If

  StartMsg "Executing SQL Query..."
  Set rsQuery = frmMain.svr.Databases(ctx.CurrentDB).Execute("SELECT * FROM " & ctx.CurrentObject.FormattedID & szLimit)
  Load objOutputForm
  objOutputForm.Display rsQuery, ctx.CurrentDB, "(" & ctx.CurrentObject.ObjectType & ": " & ctx.CurrentObject.FormattedID & ")"
  objOutputForm.Show

  EndMsg
  
  Exit Sub
  
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuPopupViewData_Click"
End Sub

Private Sub tb_ButtonClick(ByVal Button As MSComctlLib.Button)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tb_ButtonClick(" & Button & ")", etFullDebug

  Select Case Button.Key
    Case "connect"
      mnuPopupConnect_Click
    Case "refresh"
      mnuPopupRefresh_Click
    Case "create"
      If ctx.CurrentObject.ObjectType <> "Server" And _
         ctx.CurrentObject.ObjectType <> "Check" And _
         ctx.CurrentObject.ObjectType <> "Column" And _
         ctx.CurrentObject.ObjectType <> "Schema" And _
         ctx.CurrentObject.ObjectType <> "Foreign Key" Then
        tb_ButtonMenuClick Button.ButtonMenus(LCase(ctx.CurrentObject.ObjectType))
      End If
    Case "drop"
      mnuPopupDrop_Click
    Case "properties"
      mnuPopupProperties_Click
    Case "sql"
      mnuPopupSQL_Click
    Case "viewdata"
      mnuPopupViewData_Click
    Case "resetstatistics"
      mnuPopupResetStatistics_Click
    Case "reindex"
      mnuPopupReindex_Click
    Case "vacuum"
      Vacuum False
    Case "record"
      mnuPopupRecordLog_Click
    Case "stop"
      mnuPopupStopRecording_Click
    Case Else
      MsgBox "Unknown menu button pressed.", vbExclamation, "Error"
  End Select
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tb_ButtonClick"
End Sub

Private Sub mnuPopupVacuumVacuum_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuPopupVacuumVacuum_Click()", etFullDebug

  Vacuum False
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuPopupVacuumVacuum_Click"
End Sub

Private Sub mnuPopupVacuumAnalyse_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuPopupVacuumAnalyse_Click()", etFullDebug

  Vacuum True
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuPopupVacuumAnalyse_Click"
End Sub

Private Sub mnuPopupCreateAggregate_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuPopupCreateAggregate_Click()", etFullDebug

Dim objAggregateForm As New frmAggregate

  Load objAggregateForm
  objAggregateForm.Initialise ctx.CurrentDB, ctx.CurrentNS
  objAggregateForm.Show
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuPopupCreateAggregate_Click"
End Sub

Private Sub mnuPopupCreateCast_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuPopupCreateCast_Click()", etFullDebug

Dim objCastForm As New frmCast

  Load objCastForm
  objCastForm.Initialise ctx.CurrentDB
  objCastForm.Show
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuPopupCreateCast_Click"
End Sub

Private Sub mnuPopupCreateConversion_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuPopupCreateConversion_Click()", etFullDebug

Dim objConversionForm As New frmConversion

  Load objConversionForm
  objConversionForm.Initialise ctx.CurrentDB, ctx.CurrentNS
  objConversionForm.Show
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuPopupCreateConversion_Click"
End Sub

Private Sub mnuPopupCreateDatabase_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuPopupCreateDatabase_Click()", etFullDebug

Dim objDatabaseForm As New frmDatabase

  Load objDatabaseForm
  objDatabaseForm.Initialise
  objDatabaseForm.Show
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuPopupCreateDatabase_Click"
End Sub

Private Sub mnuPopupCreateFunction_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuPopupCreateFunction_Click()", etFullDebug

Dim objFunctionForm As New frmFunction

  Load objFunctionForm
  objFunctionForm.Initialise ctx.CurrentDB, ctx.CurrentNS
  objFunctionForm.Show
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuPopupCreateFunction_Click"
End Sub

Private Sub mnuPopupCreateGroup_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuPopupCreateGroup_Click()", etFullDebug

Dim objGroupForm As New frmGroup

  Load objGroupForm
  objGroupForm.Initialise
  objGroupForm.Show
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuPopupCreateGroup_Click"
End Sub

Private Sub mnuPopupCreateIndex_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuPopupCreateIndex_Click()", etFullDebug

Dim objIndexForm As New frmIndex

  Load objIndexForm
  objIndexForm.Initialise ctx.CurrentDB, ctx.CurrentNS
  objIndexForm.Show
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuPopupCreateIndex_Click"
End Sub

Private Sub mnuPopupCreateLanguage_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuPopupCreateLanguage_Click()", etFullDebug

Dim objLanguageForm As New frmLanguage

  Load objLanguageForm
  objLanguageForm.Initialise ctx.CurrentDB
  objLanguageForm.Show
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuPopupCreateLanguage_Click"
End Sub

Private Sub mnuPopupCreateNamespace_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuPopupCreateNamespace_Click()", etFullDebug

Dim objNamespaceForm As New frmNamespace

  Load objNamespaceForm
  objNamespaceForm.Initialise ctx.CurrentDB
  objNamespaceForm.Show
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuPopupCreateNamespace_Click"
End Sub

Private Sub mnuPopupCreateOperator_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuPopupCreateOperator_Click()", etFullDebug

Dim objOperatorForm As New frmOperator

  Load objOperatorForm
  objOperatorForm.Initialise ctx.CurrentDB, ctx.CurrentNS
  objOperatorForm.Show
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuPopupCreateOperator_Click"
End Sub

Private Sub mnuPopupCreateRule_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuPopupCreateRule_Click()", etFullDebug

Dim objRuleForm As New frmRule

  Load objRuleForm
  objRuleForm.Initialise ctx.CurrentDB, ctx.CurrentNS
  objRuleForm.Show
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuPopupCreateRule_Click"
End Sub

Private Sub mnuPopupCreateSequence_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuPopupCreateSequence_Click()", etFullDebug

Dim objSequenceForm As New frmSequence

  Load objSequenceForm
  objSequenceForm.Initialise ctx.CurrentDB, ctx.CurrentNS
  objSequenceForm.Show
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuPopupCreateSequence_Click"
End Sub

Private Sub mnuPopupCreateTable_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuPopupCreateTable_Click()", etFullDebug

Dim objTableForm As New frmTable

  Load objTableForm
  objTableForm.Initialise ctx.CurrentDB, ctx.CurrentNS
  objTableForm.Show
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuPopupCreateTable_Click"
End Sub

Private Sub mnuPopupCreateTrigger_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuPopupCreateTrigger_Click()", etFullDebug

Dim objTriggerForm As New frmTrigger

  Load objTriggerForm
  objTriggerForm.Initialise ctx.CurrentDB, ctx.CurrentNS
  objTriggerForm.Show
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuPopupCreateTrigger_Click"
End Sub

Private Sub mnuPopupCreateType_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuPopupCreateType_Click()", etFullDebug

Dim objTypeForm As New frmType

  Load objTypeForm
  objTypeForm.Initialise ctx.CurrentDB, ctx.CurrentNS
  objTypeForm.Show
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuPopupCreateType_Click"
End Sub

Private Sub mnuPopupCreateUser_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuPopupCreateUser_Click()", etFullDebug

Dim objUserForm As New frmUser

  Load objUserForm
  objUserForm.Initialise
  objUserForm.Show
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuPopupCreateUser_Click"
End Sub

Private Sub mnuPopupCreateView_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuPopupCreateView_Click()", etFullDebug

Dim objViewForm As New frmView

  Load objViewForm
  objViewForm.Initialise ctx.CurrentDB, ctx.CurrentNS
  objViewForm.Show
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuPopupCreateView_Click"
End Sub

Private Sub tb_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tb_ButtonMenuClick(" & ButtonMenu & ")", etFullDebug

  Select Case ButtonMenu.Parent.Key
    Case "connect"
      Load frmConnect
      frmConnect.Load_Defaults Val(Mid(ButtonMenu.Key, 1, InStr(1, ButtonMenu.Key, "|") - 1))
      If InStr(1, Command, "-wine") <> 0 Then
        frmConnect.Show
      Else
        frmConnect.Show vbModal, Me
      End If
    
    Case "create"
    
      'For each of these just call the popup menu function
      Select Case ButtonMenu.Key
        Case "aggregate"
          mnuPopupCreateAggregate_Click
        Case "conversion"
          mnuPopupCreateConversion_Click
        Case "cast"
          mnuPopupCreateCast_Click
        Case "database"
          mnuPopupCreateDatabase_Click
        Case "domain"
          mnuPopupCreateDomain_Click
        Case "function"
          mnuPopupCreateFunction_Click
        Case "group"
          mnuPopupCreateGroup_Click
        Case "index"
          mnuPopupCreateIndex_Click
        Case "language"
          mnuPopupCreateLanguage_Click
        Case "operator"
          mnuPopupCreateOperator_Click
        Case "rule"
          mnuPopupCreateRule_Click
        Case "namespace"
          mnuPopupCreateNamespace_Click
        Case "sequence"
          mnuPopupCreateSequence_Click
        Case "table"
          mnuPopupCreateTable_Click
        Case "trigger"
          mnuPopupCreateTrigger_Click
        Case "type"
          mnuPopupCreateType_Click
        Case "user"
          mnuPopupCreateUser_Click
        Case "view"
          mnuPopupCreateView_Click
      End Select
      
    Case "vacuum"
      Select Case ButtonMenu.Key
        Case "vacuum"
          Vacuum False
        Case "analyse"
          Vacuum True
      End Select
      
    Case Else
      MsgBox "Unknown button menu option pressed."
      
  End Select
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tb_ButtonMenuClick"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
svr.LogEvent "Entering " & App.Title & ":frmMain.Form_Unload(" & Cancel & ")", etFullDebug

Dim objForm As Form
Dim lTop As Long
Dim lLeft As Long
  
  'Close child forms.
  For Each objForm In Forms
    Unload objForm
  Next objForm
  
  'If child forms haven't been closed, then the user probably pressed cancel on a save dialogue
  If Forms.Count > 1 Then
    Cancel = 1
    Exit Sub
  End If
  
  'Convert to Pixels
  lTop = Me.ScaleY(Me.Top, 1, 3)
  lLeft = Me.ScaleX(Me.Left, 1, 3)
  
  'Check the position
  If ((lTop < 0) Or (lTop > Screen.Height - 10)) Then lTop = 0
  If ((lLeft < 0) Or (lLeft > Screen.Width - 10)) Then lLeft = 0
  
  'Convert back to Twips
  lTop = Me.ScaleY(lTop, 3, 1)
  lLeft = Me.ScaleX(lLeft, 3, 1)
  
  'Save the Window size/position
  RegWrite HKEY_CURRENT_USER, "Software\" & App.Title, "Top", regString, lTop
  RegWrite HKEY_CURRENT_USER, "Software\" & App.Title, "Left", regString, lLeft
  RegWrite HKEY_CURRENT_USER, "Software\" & App.Title, "Width", regString, Me.Width
  RegWrite HKEY_CURRENT_USER, "Software\" & App.Title, "Height", regString, Me.Height
  RegWrite HKEY_CURRENT_USER, "Software\" & App.Title, "Horizontal Splitter", regString, splHorizontal.Top
  RegWrite HKEY_CURRENT_USER, "Software\" & App.Title, "Vertical Splitter", regString, splVertical.Left
  
  'Clear the Server, then Context objects last as the forms may be using them for logging
  Set svr = Nothing
  Set ctx = Nothing
End Sub

Private Sub mnuFileExit_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuFileExit_Click()", etFullDebug

  Unload Me
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuFileExit_Click"
End Sub

Private Sub mnuFileConnect_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuFileConnect_Click()", etFullDebug

  Load frmConnect
  frmConnect.Load_Defaults
  If InStr(1, Command, "-wine") <> 0 Then
    frmConnect.Show
  Else
    frmConnect.Show vbModal, Me
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuFileConnect_Click"
End Sub

Private Sub svr_EventLog(EventLevel As pgSchema.LogLevel, EventMessage As String)
'Note - No function entry logging is done here 'cos we'd enter a loop then...

  If ctx.LogView Then If EventLevel <= ctx.LogLevel Then frmLog.LogMsg EventMessage

End Sub

Private Sub tvServer(ByVal Node As MSComctlLib.Node)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvServer(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
    
  If Node.Children = 0 Then
    Set ctx.CurrentObject.Databases.Tag = tv.Nodes.Add(Node.Key, tvwChild, "DAT+" & GetID, "Databases (" & ctx.CurrentObject.Databases.Count(Not ctx.IncludeSys) & ")", "database")
    Set ctx.CurrentObject.Groups.Tag = tv.Nodes.Add(Node.Key, tvwChild, "GRP+" & GetID, "Groups (" & ctx.CurrentObject.Groups.Count & ")", "group")
    Set ctx.CurrentObject.Users.Tag = tv.Nodes.Add(Node.Key, tvwChild, "USR+" & GetID, "Users (" & ctx.CurrentObject.Users.Count & ")", "user")
  End If
  lv.ColumnHeaders.Add , , "Property"
  lv.ColumnHeaders.Add , , "Value"
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Hostname", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Server & ""
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Port", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Port
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Username", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Username
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Last system OID", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.LastSystemOID
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "ODBC driver", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.DriverName
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "ODBC driver version", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.DriverVersion.Major & "." & ctx.CurrentObject.DriverVersion.Minor & "." & ctx.CurrentObject.DriverVersion.Revision
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "PostgreSQL version", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.dbVersion.Major & "." & ctx.CurrentObject.dbVersion.Minor & "." & ctx.CurrentObject.dbVersion.Revision
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "DBMS", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.dbVersion.Description
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Connection string", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.ConnectionString
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvServer"
End Sub

Private Sub svServer(ByVal Node As MSComctlLib.Node)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.svServer(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
Dim rsStat As New Recordset

  ' Statistics.
  ' These don't come from pgSchema because they aren't really schema related.
  If ctx.dbVer >= 7.2 Then
    Set rsStat = svr.Databases(svr.MasterDB).Execute("SELECT datname, procpid, usename, current_query FROM pg_stat_activity")
    sv.ColumnHeaders.Add , , "Database"
    sv.ColumnHeaders.Add , , "PID"
    sv.ColumnHeaders.Add , , "Username"
    sv.ColumnHeaders.Add , , "Current Query"
  
    While Not rsStat.EOF
      If Not (svr.Databases(rsStat!datname).SystemObject And Not ctx.IncludeSys) Then
        Set lvItem = sv.ListItems.Add(, "STA-" & GetID, rsStat!datname & "", "statistics", "statistics")
        lvItem.SubItems(1) = rsStat!procpid & ""
        lvItem.SubItems(2) = rsStat!usename & ""
        lvItem.SubItems(3) = rsStat!current_query & ""
      End If
      rsStat.MoveNext
    Wend
    If rsStat.State <> adStateClosed Then rsStat.Close
    Set rsStat = Nothing
  Else
    sv.ColumnHeaders.Add , , "Statistics"
    Set lvItem = sv.ListItems.Add(, "STA-" & GetID, "Statistics are only available with PostgreSQL 7.2 or higher.", "server", "server")
  End If
  
  Exit Sub
Err_Handler:
  If rsStat.State <> adStateClosed Then rsStat.Close
  Set rsStat = Nothing
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.svServer"
End Sub

Private Sub tvDatabases(ByVal Node As MSComctlLib.Node)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvDatabases(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
Dim dat As pgDatabase

  If Node.Children = 0 Or Node.Children <> svr.Databases.Count(Not ctx.IncludeSys) Then
    While Not (Node.Child Is Nothing)
      tv.Nodes.Remove Node.Child.Index
    Wend
    For Each dat In svr.Databases
      If Not (dat.SystemObject And Not ctx.IncludeSys) Then
        
        'Connect now when not deferring to get a valid status
        If Not svr.DeferConnection Then dat.dbConnect
        
        If dat.Status <> statInaccessible Then
          If svr.DeferConnection Then
            Set dat.Tag = tv.Nodes.Add(Node.Key, tvwChild, "DAT-" & GetID, dat.Identifier, "closeddatabase")
          Else
            Set dat.Tag = tv.Nodes.Add(Node.Key, tvwChild, "DAT-" & GetID, dat.Identifier, "database")
          End If
        Else
          Set dat.Tag = tv.Nodes.Add(Node.Key, tvwChild, "DAT-" & GetID, dat.Identifier, "baddatabase")
        End If
      End If
    Next dat
    Node.Text = "Databases (" & Node.Children & ")"
  End If
  lv.ColumnHeaders.Add , , "Database"
  lv.ColumnHeaders.Add , , "Comment"
  For Each dat In svr.Databases
    If Not (dat.SystemObject And Not ctx.IncludeSys) Then
      Set lvItem = lv.ListItems.Add(, "DAT-" & GetID, dat.Identifier, "database", "database")
      lvItem.SubItems(1) = Replace(dat.Comment, vbCrLf, " ")
    End If
  Next dat
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvDatabases"
End Sub

Private Sub svDatabases(ByVal Node As MSComctlLib.Node)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.svDatabases(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
Dim rsStat As New Recordset

  ' Statistics.
  ' These don't come from pgSchema because they aren't really schema related.
  If ctx.dbVer >= 7.2 Then
    Set rsStat = svr.Databases(svr.MasterDB).Execute("SELECT datname, numbackends, xact_commit, xact_rollback, blks_read, blks_hit FROM pg_stat_database ORDER BY datname")
    sv.ColumnHeaders.Add , , "Database", 2000
    sv.ColumnHeaders.Add , , "Backends", 1500
    sv.ColumnHeaders.Add , , "Xact Committed", 1500
    sv.ColumnHeaders.Add , , "Xact Rolled Back", 1500
    sv.ColumnHeaders.Add , , "Blocks Read", 1500
    sv.ColumnHeaders.Add , , "Blocks Hit", 1500
  
    While Not rsStat.EOF
      If svr.Databases.Exists(rsStat!datname) Then
        If Not (svr.Databases(rsStat!datname).SystemObject And Not ctx.IncludeSys) Then
          Set lvItem = sv.ListItems.Add(, "STA+" & GetID, rsStat!datname & "", "statistics", "statistics")
          lvItem.SubItems(1) = rsStat!numbackends & ""
          lvItem.SubItems(2) = rsStat!xact_commit & ""
          lvItem.SubItems(3) = rsStat!xact_rollback & ""
          lvItem.SubItems(4) = rsStat!blks_read & ""
          lvItem.SubItems(5) = rsStat!blks_hit & ""
        End If
      End If
      rsStat.MoveNext
    Wend
    If rsStat.State <> adStateClosed Then rsStat.Close
    Set rsStat = Nothing
  Else
    sv.ColumnHeaders.Add , , "Statistics"
    Set lvItem = sv.ListItems.Add(, "STA+" & GetID, "Statistics are only available with PostgreSQL 7.2 or higher.", "server", "server")
  End If
  
  Exit Sub
Err_Handler:
  If rsStat.State <> adStateClosed Then rsStat.Close
  Set rsStat = Nothing
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.svDatabases"
End Sub

Private Sub tvDatabase(ByVal Node As MSComctlLib.Node)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvDatabase(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
Dim objVar As pgVar
Dim szTemp As String

  'Connect if required
  If svr.DeferConnection And svr.Databases(Node.Text).Status <> statOpen Then
    If Not svr.Databases(Node.Text).dbConnect Then
      If svr.Databases(Node.Text).Status = statClosed Then
        Node.Image = "closeddatabase"
      Else
        Node.Image = "baddatabase"
      End If
    Else
      Node.Image = "database"
    End If
  Else
    Node.Image = "database"
  End If
  
  If svr.Databases(Node.Text).Status = statOpen Then
    If Node.Children = 0 Then
      If ctx.dbVer >= 7.3 Then Set ctx.CurrentObject.Casts.Tag = tv.Nodes.Add(Node.Key, tvwChild, "CST+" & GetID, "Casts (" & ctx.CurrentObject.Casts.Count(Not ctx.IncludeSys) & ")", "cast")
      Set ctx.CurrentObject.Languages.Tag = tv.Nodes.Add(Node.Key, tvwChild, "LNG+" & GetID, "Languages (" & ctx.CurrentObject.Languages.Count(Not ctx.IncludeSys) & ")", "language")
      Set ctx.CurrentObject.Namespaces.Tag = tv.Nodes.Add(Node.Key, tvwChild, "NSP+" & GetID, "Schemas (" & ctx.CurrentObject.Namespaces.Count(Not ctx.IncludeSys) & ")", "namespace")
    End If
  Else
    Node.Image = "baddatabase"
  End If
  lv.ColumnHeaders.Add , , "Property"
  lv.ColumnHeaders.Add , , "Value"
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Name", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Name
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "OID", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Oid
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Owner", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Owner
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "ACL", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.ACL
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Path", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Path
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Server Encoding", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.ServerEncoding
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Variables", "property", "property")
  If ctx.CurrentObject.Status = statOpen Then
    For Each objVar In ctx.CurrentObject.DatabaseVars
      szTemp = szTemp & "{" & objVar.Name & " = " & objVar.Value & "}, "
    Next objVar
    If Len(szTemp) > 2 Then szTemp = Mid(szTemp, 1, Len(szTemp) - 2)
    lvItem.SubItems(1) = szTemp
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Allow Connections?", "property", "property")
  If ctx.CurrentObject.AllowConnections Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Connection Status?", "property", "property")
  If ctx.CurrentObject.Status = statInaccessible Then
    lvItem.SubItems(1) = "Inaccessible"
  ElseIf ctx.CurrentObject.Status = statOpen Then
    lvItem.SubItems(1) = "Connected"
  Else
    lvItem.SubItems(1) = "Not connected"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "System Database?", "property", "property")
  If ctx.CurrentObject.SystemObject Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Comment", "property", "property")
  lvItem.SubItems(1) = Replace(ctx.CurrentObject.Comment, vbCrLf, " ")
  
  'Set the Definition Pane
  
  If txtDefinition.Visible And (ctx.CurrentObject.Status = statOpen) Then txtDefinition.Text = ctx.CurrentObject.SQL
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvDatabase"
End Sub

Private Sub svDatabase(ByVal Node As MSComctlLib.Node)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.svDatabase(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
Dim rsStat As New Recordset

  ' Statistics.
  ' These don't come from pgSchema because they aren't really schema related.
  If ctx.dbVer >= 7.2 Then
    Set rsStat = svr.Databases(svr.MasterDB).Execute("SELECT numbackends, xact_commit, xact_rollback, blks_read, blks_hit FROM pg_stat_database WHERE datname = '" & Node.Text & "'")
    sv.ColumnHeaders.Add , , "Statistic"
    sv.ColumnHeaders.Add , , "Value"
  
    If Not rsStat.EOF Then
      Set lvItem = sv.ListItems.Add(, "STA-" & GetID, "Backends", "statistics", "statistics")
      lvItem.SubItems(1) = rsStat!numbackends & ""
      Set lvItem = sv.ListItems.Add(, "STA-" & GetID, "Xact Committed", "statistics", "statistics")
      lvItem.SubItems(1) = rsStat!xact_commit & ""
      Set lvItem = sv.ListItems.Add(, "STA-" & GetID, "Xact Rolled Back", "statistics", "statistics")
      lvItem.SubItems(1) = rsStat!xact_rollback & ""
      Set lvItem = sv.ListItems.Add(, "STA-" & GetID, "Blocks Read", "statistics", "statistics")
      lvItem.SubItems(1) = rsStat!blks_read & ""
      Set lvItem = sv.ListItems.Add(, "STA-" & GetID, "Blocks Hit", "statistics", "statistics")
      lvItem.SubItems(1) = rsStat!blks_hit & ""
    Else
      ClearStats
    End If
    If rsStat.State <> adStateClosed Then rsStat.Close
    Set rsStat = Nothing
  Else
    sv.ColumnHeaders.Add , , "Statistics"
    Set lvItem = sv.ListItems.Add(, "STA-" & GetID, "Statistics are only available with PostgreSQL 7.2 or higher.", "server", "server")
  End If
  
  Exit Sub
Err_Handler:
  If rsStat.State <> adStateClosed Then rsStat.Close
  Set rsStat = Nothing
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.svDatabase"
End Sub

Private Sub tvGroups(ByVal Node As MSComctlLib.Node)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvGroups(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
Dim szTemp As String
Dim vData As Variant
Dim grp As pgGroup

  If Node.Children = 0 Or Node.Children <> svr.Groups.Count Then
    While Not (Node.Child Is Nothing)
      tv.Nodes.Remove Node.Child.Index
    Wend
    For Each grp In svr.Groups
      Set grp.Tag = tv.Nodes.Add(Node.Key, tvwChild, "GRP-" & GetID, grp.Identifier, "group")
    Next grp
    Node.Text = "Groups (" & Node.Children & ")"
  End If
  lv.ColumnHeaders.Add , , "Group"
  lv.ColumnHeaders.Add , , "Group ID"
  lv.ColumnHeaders.Add , , "Members"
  For Each grp In svr.Groups
    Set lvItem = lv.ListItems.Add(, "GRP-" & GetID, grp.Identifier, "group", "group")
    lvItem.SubItems(1) = grp.ID
    szTemp = ""
    For Each vData In grp.Members
      szTemp = szTemp & vData & ", "
    Next vData
    If Len(szTemp) > 2 Then lvItem.SubItems(2) = Left(szTemp, Len(szTemp) - 2)
  Next grp
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvGroups"
End Sub

Private Sub tvGroup(ByVal Node As MSComctlLib.Node)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvGroup(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
Dim szTemp As String
Dim vData As Variant

  lv.ColumnHeaders.Add , , "Property"
  lv.ColumnHeaders.Add , , "Value"
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Name", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Name
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Group ID", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.ID
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Member Count", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Members.Count
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Members", "property", "property")
  For Each vData In ctx.CurrentObject.Members
    szTemp = szTemp & vData & ", "
  Next vData
  If Len(szTemp) > 2 Then lvItem.SubItems(1) = Left(szTemp, Len(szTemp) - 2)
  
  'Set the Definition Pane
  If txtDefinition.Visible Then txtDefinition.Text = ctx.CurrentObject.SQL
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvGroup"
End Sub

Private Sub tvUsers(ByVal Node As MSComctlLib.Node)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvUsers(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
Dim usr As pgUser

  If Node.Children = 0 Or Node.Children <> svr.Users.Count Then
    While Not (Node.Child Is Nothing)
      tv.Nodes.Remove Node.Child.Index
    Wend
    For Each usr In svr.Users
      Set usr.Tag = tv.Nodes.Add(Node.Key, tvwChild, "USR-" & GetID, usr.Identifier, "user")
    Next usr
    Node.Text = "Users (" & Node.Children & ")"
  End If
  lv.ColumnHeaders.Add , , "Username"
  lv.ColumnHeaders.Add , , "User ID"
  lv.ColumnHeaders.Add , , "Account Expires"
  For Each usr In svr.Users
    Set lvItem = lv.ListItems.Add(, "USR-" & GetID, usr.Identifier, "user", "user")
    lvItem.SubItems(1) = usr.ID
    lvItem.SubItems(2) = usr.AccountExpires
  Next usr
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvUsers"
End Sub

Private Sub tvUser(ByVal Node As MSComctlLib.Node)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvUser(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
Dim szTemp As String
Dim objVar As pgVar

  lv.ColumnHeaders.Add , , "Property"
  lv.ColumnHeaders.Add , , "Value"
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Name", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Name
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "User ID", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.ID
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Account Expires", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.AccountExpires
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Create Databases?", "property", "property")
  If ctx.CurrentObject.CreateDatabases Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Superuser?", "property", "property")
  If ctx.CurrentObject.Superuser Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Update Catalogues", "property", "property")
  If ctx.CurrentObject.UpdateCatalogues Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Variables", "property", "property")
  For Each objVar In ctx.CurrentObject.UserVars
    szTemp = szTemp & "{" & objVar.Name & " = " & objVar.Value & "}, "
  Next objVar
  If Len(szTemp) > 2 Then szTemp = Mid(szTemp, 1, Len(szTemp) - 2)
  lvItem.SubItems(1) = szTemp
  
  'Set the Definition Pane
  If txtDefinition.Visible Then txtDefinition.Text = ctx.CurrentObject.SQL
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvUser"
End Sub

Private Sub tvCasts(ByVal Node As MSComctlLib.Node)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvCasts(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
Dim cst As pgCast

  If Node.Children = 0 Or Node.Children <> svr.Databases(Node.Parent.Text).Casts.Count(Not ctx.IncludeSys) Then
    While Not (Node.Child Is Nothing)
      tv.Nodes.Remove Node.Child.Index
    Wend
    For Each cst In svr.Databases(Node.Parent.Text).Casts
      If Not (cst.SystemObject And Not ctx.IncludeSys) Then Set cst.Tag = tv.Nodes.Add(Node.Key, tvwChild, "CST-" & GetID, cst.Identifier, "cast")
    Next cst
    Node.Text = "Casts (" & Node.Children & ")"
  End If
  lv.ColumnHeaders.Add , , "Cast", lv.Width
  For Each cst In svr.Databases(Node.Parent.Text).Casts
    If Not (cst.SystemObject And Not ctx.IncludeSys) Then
      Set lvItem = lv.ListItems.Add(, "CST-" & GetID, cst.Identifier, "cast", "cast")
    End If
  Next cst
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvCasts"
End Sub

Private Sub tvCast(ByVal Node As MSComctlLib.Node)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvCast(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem

  lv.ColumnHeaders.Add , , "Property"
  lv.ColumnHeaders.Add , , "Value"
  
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Name", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Name
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "OID", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Oid
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Type source", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Source
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Type target", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Target
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Function", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Funct
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Context", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Context
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "System Cast?", "property", "property")
  If ctx.CurrentObject.SystemObject Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  
  'Set the Definition Pane
  If txtDefinition.Visible Then txtDefinition.Text = ctx.CurrentObject.SQL
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvCast"
End Sub

Private Sub tvLanguages(ByVal Node As MSComctlLib.Node)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvLanguages(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
Dim lng As pgLanguage

  If Node.Children = 0 Or Node.Children <> svr.Databases(Node.Parent.Text).Languages.Count(Not ctx.IncludeSys) Then
    While Not (Node.Child Is Nothing)
      tv.Nodes.Remove Node.Child.Index
    Wend
    For Each lng In svr.Databases(Node.Parent.Text).Languages
      If Not (lng.SystemObject And Not ctx.IncludeSys) Then Set lng.Tag = tv.Nodes.Add(Node.Key, tvwChild, "LNG-" & GetID, lng.Identifier, "language")
    Next lng
    Node.Text = "Languages (" & Node.Children & ")"
  End If
  lv.ColumnHeaders.Add , , "Language", lv.Width
  For Each lng In svr.Databases(Node.Parent.Text).Languages
    If Not (lng.SystemObject And Not ctx.IncludeSys) Then
      Set lvItem = lv.ListItems.Add(, "LNG-" & GetID, lng.Identifier, "language", "language")
    End If
  Next lng
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvLanguages"
End Sub

Private Sub tvLanguage(ByVal Node As MSComctlLib.Node)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvLanguage(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem

  lv.ColumnHeaders.Add , , "Property"
  lv.ColumnHeaders.Add , , "Value"
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Name", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Name
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "OID", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Oid
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "ACL", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.ACL
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Handler", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Handler
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Trusted?", "property", "property")
  If ctx.CurrentObject.Trusted Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "System Language?", "property", "property")
  If ctx.CurrentObject.SystemObject Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  
  'Set the Definition Pane
  If txtDefinition.Visible Then txtDefinition.Text = ctx.CurrentObject.SQL
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvLanguage"
End Sub

Private Sub tvNamespaces(ByVal Node As MSComctlLib.Node)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvNamespaces(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
Dim nsp As pgNamespace

  If Node.Children = 0 Or Node.Children <> svr.Databases(Node.Parent.Text).Namespaces.Count(Not ctx.IncludeSys) Then
    While Not (Node.Child Is Nothing)
      tv.Nodes.Remove Node.Child.Index
    Wend
    For Each nsp In svr.Databases(Node.Parent.Text).Namespaces
      If Not (nsp.SystemObject And Not ctx.IncludeSys) Then Set nsp.Tag = tv.Nodes.Add(Node.Key, tvwChild, "NSP-" & GetID, nsp.Identifier, "namespace")
    Next nsp
    Node.Text = "Schemas (" & Node.Children & ")"
  End If
  lv.ColumnHeaders.Add , , "Schema"
  lv.ColumnHeaders.Add , , "Comment"
  For Each nsp In svr.Databases(Node.Parent.Text).Namespaces
    If Not (nsp.SystemObject And Not ctx.IncludeSys) Then
      Set lvItem = lv.ListItems.Add(, "NSP-" & GetID, nsp.Identifier, "namespace", "namespace")
      lvItem.SubItems(1) = Replace(nsp.Comment, vbCrLf, " ")
    End If
  Next nsp
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvNamespaces"
End Sub

Private Sub tvNamespace(ByVal Node As MSComctlLib.Node)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvNamespace(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem

  If Node.Children = 0 Then
    Set ctx.CurrentObject.Aggregates.Tag = tv.Nodes.Add(Node.Key, tvwChild, "AGG+" & GetID, "Aggregates (" & ctx.CurrentObject.Aggregates.Count(Not ctx.IncludeSys) & ")", "aggregate")
    If ctx.dbVer >= 7.3 Then Set ctx.CurrentObject.Conversions.Tag = tv.Nodes.Add(Node.Key, tvwChild, "CNV+" & GetID, "Conversions (" & ctx.CurrentObject.Conversions.Count(Not ctx.IncludeSys) & ")", "conversion")
    If ctx.dbVer >= 7.3 Then Set ctx.CurrentObject.Domains.Tag = tv.Nodes.Add(Node.Key, tvwChild, "DOM+" & GetID, "Domains (" & ctx.CurrentObject.Domains.Count(Not ctx.IncludeSys) & ")", "domain")
    Set ctx.CurrentObject.Functions.Tag = tv.Nodes.Add(Node.Key, tvwChild, "FNC+" & GetID, "Functions (" & ctx.CurrentObject.Functions.Count(Not ctx.IncludeSys) & ")", "function")
    Set ctx.CurrentObject.Operators.Tag = tv.Nodes.Add(Node.Key, tvwChild, "OPR+" & GetID, "Operators (" & ctx.CurrentObject.Operators.Count(Not ctx.IncludeSys) & ")", "operator")
    Set ctx.CurrentObject.Sequences.Tag = tv.Nodes.Add(Node.Key, tvwChild, "SEQ+" & GetID, "Sequences (" & ctx.CurrentObject.Sequences.Count(Not ctx.IncludeSys) & ")", "sequence")
    Set ctx.CurrentObject.Tables.Tag = tv.Nodes.Add(Node.Key, tvwChild, "TBL+" & GetID, "Tables (" & ctx.CurrentObject.Tables.Count(Not ctx.IncludeSys) & ")", "table")
    Set ctx.CurrentObject.Types.Tag = tv.Nodes.Add(Node.Key, tvwChild, "TYP+" & GetID, "Types (" & ctx.CurrentObject.Types.Count(Not ctx.IncludeSys) & ")", "type")
    Set ctx.CurrentObject.Views.Tag = tv.Nodes.Add(Node.Key, tvwChild, "VIE+" & GetID, "Views (" & ctx.CurrentObject.Views.Count(Not ctx.IncludeSys) & ")", "view")
  End If
    
  lv.ColumnHeaders.Add , , "Property"
  lv.ColumnHeaders.Add , , "Value"
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Name", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Name
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "OID", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Oid
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Owner", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Owner
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "ACL", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.ACL
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "System Schema?", "property", "property")
  If ctx.CurrentObject.SystemObject Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Comment", "property", "property")
  lvItem.SubItems(1) = Replace(ctx.CurrentObject.Comment, vbCrLf, " ")
  
  'Set the Definition Pane
  If txtDefinition.Visible Then txtDefinition.Text = ctx.CurrentObject.SQL
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvNamespace"
End Sub

Private Sub tvAggregates(ByVal Node As MSComctlLib.Node)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvAggregates(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
Dim agg As pgAggregate

  If Node.Children = 0 Or Node.Children <> svr.Databases(Node.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Text).Aggregates.Count(Not ctx.IncludeSys) Then
    While Not (Node.Child Is Nothing)
      tv.Nodes.Remove Node.Child.Index
    Wend
    For Each agg In svr.Databases(Node.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Text).Aggregates
      If Not (agg.SystemObject And Not ctx.IncludeSys) Then Set agg.Tag = tv.Nodes.Add(Node.Key, tvwChild, "AGG-" & GetID, agg.Identifier, "aggregate")
    Next agg
    Node.Text = "Aggregates (" & Node.Children & ")"
  End If
  lv.ColumnHeaders.Add , , "Aggregate"
  lv.ColumnHeaders.Add , , "Comment"
  For Each agg In svr.Databases(Node.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Text).Aggregates
    If Not (agg.SystemObject And Not ctx.IncludeSys) Then
      Set lvItem = lv.ListItems.Add(, "AGG-" & GetID, agg.Identifier, "aggregate", "aggregate")
      lvItem.SubItems(1) = Replace(agg.Comment, vbCrLf, " ")
    End If
  Next agg
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvAggregates"
End Sub

Private Sub tvAggregate(ByVal Node As MSComctlLib.Node)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvAggregate(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem

  lv.ColumnHeaders.Add , , "Property"
  lv.ColumnHeaders.Add , , "Value"
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Name", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Name
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Input Type", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.InputType
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "OID", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Oid
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Owner", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Owner
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "State Type", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.StateType
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "State Function", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.StateFunction
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Final Type", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.FinalType
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Final Function", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.FinalFunction
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Initial Condition", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.InitialCondition
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "System Aggregate?", "property", "property")
  If ctx.CurrentObject.SystemObject Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Comment", "property", "property")
  lvItem.SubItems(1) = Replace(ctx.CurrentObject.Comment, vbCrLf, " ")
  
  'Set the Definition Pane
  If txtDefinition.Visible Then txtDefinition.Text = ctx.CurrentObject.SQL
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvAggregate"
End Sub

Private Sub tvDomains(ByVal Node As MSComctlLib.Node)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvDomains(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
Dim dom As pgDomain


  If Node.Children = 0 Or Node.Children <> svr.Databases(Node.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Text).Domains.Count(Not ctx.IncludeSys) Then
    While Not (Node.Child Is Nothing)
      tv.Nodes.Remove Node.Child.Index
    Wend
    For Each dom In svr.Databases(Node.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Text).Domains
      If Not (dom.SystemObject And Not ctx.IncludeSys) Then Set dom.Tag = tv.Nodes.Add(Node.Key, tvwChild, "DOM-" & GetID, dom.Identifier, "domain")
    Next dom
    Node.Text = "Domains (" & Node.Children & ")"
  End If
  lv.ColumnHeaders.Add , , "Domain"
  lv.ColumnHeaders.Add , , "Comment"
  For Each dom In svr.Databases(Node.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Text).Domains
    If Not (dom.SystemObject And Not ctx.IncludeSys) Then
      Set lvItem = lv.ListItems.Add(, "DOM-" & GetID, dom.Identifier, "domain", "domain")
      lvItem.SubItems(1) = Replace(dom.Comment, vbCrLf, " ")
    End If
  Next dom
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvDomains"
End Sub

Private Sub tvDomain(ByVal Node As MSComctlLib.Node)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvDomain(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem

  lv.ColumnHeaders.Add , , "Property"
  lv.ColumnHeaders.Add , , "Value"
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Name", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Name
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "OID", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Oid
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Owner", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Owner
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Base Type", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.BaseType
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Length", "property", "property")
  If ctx.CurrentObject.Length = 0 Then
    lvItem.SubItems(1) = "Variable"
  Else
    lvItem.SubItems(1) = ctx.CurrentObject.Length
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Numeric Scale", "property", "property")
  If ctx.CurrentObject.BaseType = "numeric" Then
    lvItem.SubItems(1) = ctx.CurrentObject.NumericScale
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Default", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Default
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Not Null?", "property", "property")
  If ctx.CurrentObject.NotNull Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "System Domain?", "property", "property")
  If ctx.CurrentObject.SystemObject Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Comment", "property", "property")
  lvItem.SubItems(1) = Replace(ctx.CurrentObject.Comment, vbCrLf, " ")

  'Set the Definition Pane
  If txtDefinition.Visible Then txtDefinition.Text = ctx.CurrentObject.SQL
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvDomain"
End Sub

Private Sub tvFunctions(ByVal Node As MSComctlLib.Node)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvFunctions(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
Dim szTemp As String
Dim vData As Variant
Dim fnc As pgFunction

  If Node.Children = 0 Or Node.Children <> svr.Databases(Node.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Text).Functions.Count(Not ctx.IncludeSys) Then
    While Not (Node.Child Is Nothing)
      tv.Nodes.Remove Node.Child.Index
    Wend
    For Each fnc In svr.Databases(Node.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Text).Functions
      If Not (fnc.SystemObject And Not ctx.IncludeSys) Then Set fnc.Tag = tv.Nodes.Add(Node.Key, tvwChild, "FNC-" & GetID, fnc.Identifier, "function")
    Next fnc
    Node.Text = "Functions (" & Node.Children & ")"
  End If
  lv.ColumnHeaders.Add , , "Function"
  lv.ColumnHeaders.Add , , "Comment"
  For Each fnc In svr.Databases(Node.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Text).Functions
    If Not (fnc.SystemObject And Not ctx.IncludeSys) Then
      Set lvItem = lv.ListItems.Add(, "FNC-" & GetID, fnc.Identifier, "function", "function")
      szTemp = ""
      For Each vData In fnc.Arguments
        szTemp = szTemp & vData & ", "
      Next vData
      If Len(szTemp) > 2 Then szTemp = Left(szTemp, Len(szTemp) - 2)
      lvItem.SubItems(1) = Replace(fnc.Comment, vbCrLf, " ")
    End If
  Next fnc
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvFunctions"
End Sub

Private Sub tvFunction(ByVal Node As MSComctlLib.Node)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvFunction(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
Dim szTemp As String
Dim vData As Variant

  lv.ColumnHeaders.Add , , "Property"
  lv.ColumnHeaders.Add , , "Value"
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Name", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Name
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Argument Count", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Arguments.Count
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Arguments", "property", "property")
  szTemp = ""
  For Each vData In ctx.CurrentObject.Arguments
    szTemp = szTemp & vData & ", "
  Next vData
  If Len(szTemp) > 2 Then szTemp = Left(szTemp, Len(szTemp) - 2)
  lvItem.SubItems(1) = szTemp
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Returns", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Returns
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "OID", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Oid
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Owner", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Owner
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "ACL", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.ACL
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Language", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Language
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Source", "property", "property")
  lvItem.SubItems(1) = Replace(ctx.CurrentObject.Source, vbCrLf, " ")
  If ctx.dbVer < 7.3 Then
    Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Cachable?", "property", "property")
    If ctx.CurrentObject.Cachable Then
      lvItem.SubItems(1) = "Yes"
    Else
      lvItem.SubItems(1) = "No"
    End If
  Else
    Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Returns a Set?", "property", "property")
    If ctx.CurrentObject.RetSet Then
      lvItem.SubItems(1) = "Yes"
    Else
      lvItem.SubItems(1) = "No"
    End If
    Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Volatility", "property", "property")
    lvItem.SubItems(1) = ctx.CurrentObject.Volatility
      Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Security Definer?", "property", "property")
    If ctx.CurrentObject.SecDef Then
      lvItem.SubItems(1) = "Yes"
    Else
      lvItem.SubItems(1) = "No"
    End If
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Strict?", "property", "property")
  If ctx.CurrentObject.Strict Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "System Function?", "property", "property")
  If ctx.CurrentObject.SystemObject Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Comment", "property", "property")
  lvItem.SubItems(1) = Replace(ctx.CurrentObject.Comment, vbCrLf, " ")
  
  'Set the Definition Pane
  If txtDefinition.Visible Then txtDefinition.Text = ctx.CurrentObject.SQL
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvFunction"
End Sub

Private Sub tvOperators(ByVal Node As MSComctlLib.Node)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvOperators(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
Dim opr As pgOperator

  If Node.Children = 0 Or Node.Children <> svr.Databases(Node.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Text).Operators.Count(Not ctx.IncludeSys) Then
    While Not (Node.Child Is Nothing)
      tv.Nodes.Remove Node.Child.Index
    Wend
    For Each opr In svr.Databases(Node.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Text).Operators
      If Not (opr.SystemObject And Not ctx.IncludeSys) Then Set opr.Tag = tv.Nodes.Add(Node.Key, tvwChild, "OPR-" & GetID, opr.Identifier, "operator")
    Next opr
    Node.Text = "Operators (" & Node.Children & ")"
  End If
  lv.ColumnHeaders.Add , , "Operator"
  lv.ColumnHeaders.Add , , "Comment"
  For Each opr In svr.Databases(Node.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Text).Operators
    If Not (opr.SystemObject And Not ctx.IncludeSys) Then
      Set lvItem = lv.ListItems.Add(, "OPR-" & GetID, opr.Identifier, "operator", "operator")
      lvItem.SubItems(1) = Replace(opr.Comment, vbCrLf, " ")
    End If
  Next opr
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvOperators"
End Sub

Private Sub tvOperator(ByVal Node As MSComctlLib.Node)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvOperator(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem

  lv.ColumnHeaders.Add , , "Property"
  lv.ColumnHeaders.Add , , "Value"
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Name", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Name
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Left Type", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.LeftOperandType
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Right Type", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.RightOperandType
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "OID", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Oid
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Owner", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Owner
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Operator Function", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.OperatorFunction
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Join Function", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.JoinFunction
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Restrict Function", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.RestrictFunction
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Result Type", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.ResultType
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Commutator", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Commutator
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Negator", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Negator
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Kind", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Kind
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Left Sort Operator", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.LeftTypeSortOperator
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Right Sort Operator", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.RightTypeSortOperator
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Hash Joins?", "property", "property")
  If ctx.CurrentObject.HashJoins Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "System Operator?", "property", "property")
  If ctx.CurrentObject.SystemObject Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Comment", "property", "property")
  lvItem.SubItems(1) = Replace(ctx.CurrentObject.Comment, vbCrLf, " ")
  
  'Set the Definition Pane
  If txtDefinition.Visible Then txtDefinition.Text = ctx.CurrentObject.SQL
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvOperator"
End Sub

Private Sub tvSequences(ByVal Node As MSComctlLib.Node)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvSequences(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
Dim seq As pgSequence

  If Node.Children = 0 Or Node.Children <> svr.Databases(Node.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Text).Sequences.Count(Not ctx.IncludeSys) Then
    While Not (Node.Child Is Nothing)
      tv.Nodes.Remove Node.Child.Index
    Wend
    For Each seq In svr.Databases(Node.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Text).Sequences
      If Not (seq.SystemObject And Not ctx.IncludeSys) Then Set seq.Tag = tv.Nodes.Add(Node.Key, tvwChild, "SEQ-" & GetID, seq.Identifier, "sequence")
    Next seq
    Node.Text = "Sequences (" & Node.Children & ")"
  End If
  lv.ColumnHeaders.Add , , "Sequence"
  lv.ColumnHeaders.Add , , "Comment"
  For Each seq In svr.Databases(Node.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Text).Sequences
    If Not (seq.SystemObject And Not ctx.IncludeSys) Then
      Set lvItem = lv.ListItems.Add(, "SEQ-" & GetID, seq.Identifier, "sequence", "sequence")
      lvItem.SubItems(1) = Replace(seq.Comment, vbCrLf, " ")
    End If
  Next seq
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvSequences"
End Sub

Private Sub svSequences(ByVal Node As MSComctlLib.Node)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.svSequences(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
Dim rsStat As New Recordset
Dim szSQL As String

  ' Statistics.
  ' These don't come from pgSchema because they aren't really schema related.
  If ctx.dbVer >= 7.2 Then
    If ctx.dbVer >= 7.3 Then
      szSQL = "SELECT relname, blks_read, blks_hit FROM pg_statio_all_sequences where schemaname='" & ctx.CurrentNS & "' ORDER BY relname"
    Else
      szSQL = "SELECT relname, blks_read, blks_hit FROM pg_statio_all_sequences ORDER BY relname"
    End If
    Set rsStat = svr.Databases(ctx.CurrentDB).Execute(szSQL)
    sv.ColumnHeaders.Add , , "Sequence", 2000
    sv.ColumnHeaders.Add , , "Blocks Read", 2000
    sv.ColumnHeaders.Add , , "Blocks Hit", 2000
  
    While Not rsStat.EOF
      If svr.Databases(ctx.CurrentDB).Namespaces(ctx.CurrentNS).Sequences.Exists(rsStat!relname) Then
        If Not (svr.Databases(ctx.CurrentDB).Namespaces(ctx.CurrentNS).Sequences(rsStat!relname).SystemObject And Not ctx.IncludeSys) Then
          Set lvItem = sv.ListItems.Add(, "STA+" & GetID, rsStat!relname & "", "statistics", "statistics")
          lvItem.SubItems(1) = rsStat!blks_read & ""
          lvItem.SubItems(2) = rsStat!blks_hit & ""
        End If
      End If
      rsStat.MoveNext
    Wend
    If rsStat.State <> adStateClosed Then rsStat.Close
    Set rsStat = Nothing
  Else
    sv.ColumnHeaders.Add , , "Statistics"
    Set lvItem = sv.ListItems.Add(, "STA+" & GetID, "Statistics are only available with PostgreSQL 7.2 or higher.", "server", "server")
  End If
  
  Exit Sub
Err_Handler:
  If rsStat.State <> adStateClosed Then rsStat.Close
  Set rsStat = Nothing
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.svSequences"
End Sub

Private Sub tvSequence(ByVal Node As MSComctlLib.Node)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvSequence(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem

  lv.ColumnHeaders.Add , , "Property"
  lv.ColumnHeaders.Add , , "Value"
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Name", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Name
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "OID", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Oid
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Owner", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Owner
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "ACL", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.ACL
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Last Value", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.LastValue
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Minimum", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Minimum
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Maximum", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Maximum
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Increment", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Increment
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Cache", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Cache
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Cycled?", "property", "property")
  If ctx.CurrentObject.Cycled Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "System Sequence?", "property", "property")
  If ctx.CurrentObject.SystemObject Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Comment", "property", "property")
  lvItem.SubItems(1) = Replace(ctx.CurrentObject.Comment, vbCrLf, " ")
  
  'Set the Definition Pane
  If txtDefinition.Visible Then txtDefinition.Text = ctx.CurrentObject.SQL
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvSequence"
End Sub

Private Sub svSequence(ByVal Node As MSComctlLib.Node)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.svSequence(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
Dim rsStat As New Recordset

  ' Statistics.
  ' These don't come from pgSchema because they aren't really schema related.
  If ctx.dbVer >= 7.2 Then
    Set rsStat = svr.Databases(ctx.CurrentDB).Execute("SELECT blks_read, blks_hit FROM pg_statio_all_sequences WHERE relid = " & ctx.CurrentObject.Oid & "::oid")
    sv.ColumnHeaders.Add , , "Statistic"
    sv.ColumnHeaders.Add , , "Value"
  
    If Not rsStat.EOF Then
      Set lvItem = sv.ListItems.Add(, "STA-" & GetID, "Blocks Read", "statistics", "statistics")
      lvItem.SubItems(1) = rsStat!blks_read & ""
      Set lvItem = sv.ListItems.Add(, "STA-" & GetID, "Blocks Hit", "statistics", "statistics")
      lvItem.SubItems(1) = rsStat!blks_hit & ""
    Else
      ClearStats
    End If
    If rsStat.State <> adStateClosed Then rsStat.Close
    Set rsStat = Nothing
  Else
    sv.ColumnHeaders.Add , , "Statistics"
    Set lvItem = sv.ListItems.Add(, "STA-" & GetID, "Statistics are only available with PostgreSQL 7.2 or higher.", "server", "server")
  End If
  
  Exit Sub
Err_Handler:
  If rsStat.State <> adStateClosed Then rsStat.Close
  Set rsStat = Nothing
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.svSequence"
End Sub

Private Sub tvTables(ByVal Node As MSComctlLib.Node)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvTables(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
Dim tbl As pgTable

  If Node.Children = 0 Or Node.Children <> svr.Databases(Node.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Text).Tables.Count(Not ctx.IncludeSys) Then
    While Not (Node.Child Is Nothing)
      tv.Nodes.Remove Node.Child.Index
    Wend
    For Each tbl In svr.Databases(Node.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Text).Tables
      If Not (tbl.SystemObject And Not ctx.IncludeSys) Then Set tbl.Tag = tv.Nodes.Add(Node.Key, tvwChild, "TBL-" & GetID, tbl.Identifier, "table")
    Next tbl
    Node.Text = "Tables (" & Node.Children & ")"
  End If
  lv.ColumnHeaders.Add , , "Table"
  lv.ColumnHeaders.Add , , "Comment"
  For Each tbl In svr.Databases(Node.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Text).Tables
    If Not (tbl.SystemObject And Not ctx.IncludeSys) Then
      Set lvItem = lv.ListItems.Add(, "TBL-" & GetID, tbl.Identifier, "table", "table")
      lvItem.SubItems(1) = Replace(tbl.Comment, vbCrLf, " ")
    End If
  Next tbl
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvTables"
End Sub

Private Sub svTables(ByVal Node As MSComctlLib.Node)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.svTables(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
Dim rsStat As New Recordset
Dim szSQL As String

  ' Statistics.
  ' These don't come from pgSchema because they aren't really schema related.
  If ctx.dbVer >= 7.2 Then
    If ctx.dbVer >= 7.3 Then
      szSQL = "SELECT relname, n_tup_ins, n_tup_upd, n_tup_del FROM pg_stat_all_tables where schemaname='" & ctx.CurrentNS & "' ORDER BY relname"
    Else
      szSQL = "SELECT relname, n_tup_ins, n_tup_upd, n_tup_del FROM pg_stat_all_tables ORDER BY relname"
    End If
    Set rsStat = svr.Databases(ctx.CurrentDB).Execute(szSQL)
    sv.ColumnHeaders.Add , , "Table", 2000
    sv.ColumnHeaders.Add , , "Tuples Inserted", 2000
    sv.ColumnHeaders.Add , , "Tuples Updated", 2000
    sv.ColumnHeaders.Add , , "Tuples Deleted", 2000
  
    While Not rsStat.EOF
      If svr.Databases(ctx.CurrentDB).Namespaces(ctx.CurrentNS).Tables.Exists(rsStat!relname) Then
        If Not (svr.Databases(ctx.CurrentDB).Namespaces(ctx.CurrentNS).Tables(rsStat!relname).SystemObject And Not ctx.IncludeSys) Then
          Set lvItem = sv.ListItems.Add(, "STA+" & GetID, rsStat!relname & "", "statistics", "statistics")
          lvItem.SubItems(1) = rsStat!n_tup_ins & ""
          lvItem.SubItems(2) = rsStat!n_tup_upd & ""
          lvItem.SubItems(3) = rsStat!n_tup_del & ""
        End If
      End If
      rsStat.MoveNext
    Wend
    If rsStat.State <> adStateClosed Then rsStat.Close
    Set rsStat = Nothing
  Else
    sv.ColumnHeaders.Add , , "Statistics"
    Set lvItem = sv.ListItems.Add(, "STA+" & GetID, "Statistics are only available with PostgreSQL 7.2 or higher.", "server", "server")
  End If
  
  Exit Sub
Err_Handler:
  If rsStat.State <> adStateClosed Then rsStat.Close
  Set rsStat = Nothing
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.svTables"
End Sub

Private Sub tvTable(ByVal Node As MSComctlLib.Node)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvTable(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
Dim szTemp As String
Dim vData As Variant

  lv.ColumnHeaders.Add , , "Property"
  lv.ColumnHeaders.Add , , "Value"
  If Node.Children = 0 Then
    Set ctx.CurrentObject.Checks.Tag = tv.Nodes.Add(Node.Key, tvwChild, "CHK+" & GetID, "Checks (" & ctx.CurrentObject.Checks.Count & ")", "check")
    Set ctx.CurrentObject.Columns.Tag = tv.Nodes.Add(Node.Key, tvwChild, "COL+" & GetID, "Columns (" & ctx.CurrentObject.Columns.Count(Not ctx.IncludeSys) & ")", "column")
    Set ctx.CurrentObject.ForeignKeys.Tag = tv.Nodes.Add(Node.Key, tvwChild, "FKY+" & GetID, "Foreign Keys (" & ctx.CurrentObject.ForeignKeys.Count(Not ctx.IncludeSys) & ")", "foreignkey")
    Set ctx.CurrentObject.Indexes.Tag = tv.Nodes.Add(Node.Key, tvwChild, "IND+" & GetID, "Indexes (" & ctx.CurrentObject.Indexes.Count(Not ctx.IncludeSys) & ")", "index")
    Set ctx.CurrentObject.Rules.Tag = tv.Nodes.Add(Node.Key, tvwChild, "RUL+" & GetID, "Rules (" & ctx.CurrentObject.Rules.Count(Not ctx.IncludeSys) & ")", "rule")
    Set ctx.CurrentObject.Triggers.Tag = tv.Nodes.Add(Node.Key, tvwChild, "TRG+" & GetID, "Triggers (" & ctx.CurrentObject.Triggers.Count(Not ctx.IncludeSys) & ")", "trigger")
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Name", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Name
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "OID", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Oid
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Owner", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Owner
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "ACL", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.ACL
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Rows", "property", "property")
  If ctx.AutoRowCount Then
    lvItem.SubItems(1) = ctx.CurrentObject.Rows
  Else
    lvItem.SubItems(1) = "Unknown"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Inherited Tables Count", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.InheritedTables.Count
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Inherited Tables", "property", "property")
  For Each vData In ctx.CurrentObject.InheritedTables
    szTemp = szTemp & vData & ", "
  Next vData
  If Len(szTemp) > 2 Then szTemp = Left(szTemp, Len(szTemp) - 2)
  lvItem.SubItems(1) = szTemp
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "OIDs?", "property", "property")
  If ctx.CurrentObject.HasOIDs Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "System Table?", "property", "property")
  If ctx.CurrentObject.SystemObject Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Comment", "property", "property")
  lvItem.SubItems(1) = Replace(ctx.CurrentObject.Comment, vbCrLf, " ")
  
  'Set the Definition Pane
  If txtDefinition.Visible Then txtDefinition.Text = ctx.CurrentObject.SQL
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvTable"
End Sub

Private Sub svTable(ByVal Node As MSComctlLib.Node)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.svTable(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
Dim rsStat As New Recordset

  ' Statistics.
  ' These don't come from pgSchema because they aren't really schema related.
  If ctx.dbVer >= 7.2 Then
    Set rsStat = svr.Databases(ctx.CurrentDB).Execute("SELECT seq_scan, seq_tup_read, idx_scan, idx_tup_fetch, n_tup_ins, n_tup_upd, n_tup_del, heap_blks_read, heap_blks_hit, idx_blks_read, idx_blks_hit, toast_blks_read, toast_blks_hit, tidx_blks_read, tidx_blks_hit FROM pg_stat_all_tables stat, pg_statio_all_tables statio WHERE stat.relid = statio.relid AND stat.relid = " & ctx.CurrentObject.Oid & "::oid")
    sv.ColumnHeaders.Add , , "Statistic"
    sv.ColumnHeaders.Add , , "Value"
  
    If Not rsStat.EOF Then
      Set lvItem = sv.ListItems.Add(, "STA-" & GetID, "Sequential Scans", "statistics", "statistics")
      lvItem.SubItems(1) = rsStat!seq_scan & ""
      Set lvItem = sv.ListItems.Add(, "STA-" & GetID, "Sequential Tuples Read", "statistics", "statistics")
      lvItem.SubItems(1) = rsStat!seq_tup_read & ""
      Set lvItem = sv.ListItems.Add(, "STA-" & GetID, "Index Scans", "statistics", "statistics")
      lvItem.SubItems(1) = rsStat!idx_scan & ""
      Set lvItem = sv.ListItems.Add(, "STA-" & GetID, "Index Tuples Fetched", "statistics", "statistics")
      lvItem.SubItems(1) = rsStat!idx_tup_fetch & ""
      Set lvItem = sv.ListItems.Add(, "STA-" & GetID, "Tuples Inserted", "statistics", "statistics")
      lvItem.SubItems(1) = rsStat!n_tup_ins & ""
      Set lvItem = sv.ListItems.Add(, "STA-" & GetID, "Tuples Updated", "statistics", "statistics")
      lvItem.SubItems(1) = rsStat!n_tup_upd & ""
      Set lvItem = sv.ListItems.Add(, "STA-" & GetID, "Tuples Deleted", "statistics", "statistics")
      lvItem.SubItems(1) = rsStat!n_tup_del & ""
      Set lvItem = sv.ListItems.Add(, "STA-" & GetID, "Heap Blocks Read", "statistics", "statistics")
      lvItem.SubItems(1) = rsStat!heap_blks_read & ""
      Set lvItem = sv.ListItems.Add(, "STA-" & GetID, "Heap Blocks Hit", "statistics", "statistics")
      lvItem.SubItems(1) = rsStat!heap_blks_hit & ""
      Set lvItem = sv.ListItems.Add(, "STA-" & GetID, "Index Blocks Read", "statistics", "statistics")
      lvItem.SubItems(1) = rsStat!idx_blks_read & ""
      Set lvItem = sv.ListItems.Add(, "STA-" & GetID, "Index Blocks Hit", "statistics", "statistics")
      lvItem.SubItems(1) = rsStat!idx_blks_hit & ""
      Set lvItem = sv.ListItems.Add(, "STA-" & GetID, "Toast Index Blocks Read", "statistics", "statistics")
      lvItem.SubItems(1) = rsStat!tidx_blks_read & ""
      Set lvItem = sv.ListItems.Add(, "STA-" & GetID, "Toast Index Blocks Hit", "statistics", "statistics")
      lvItem.SubItems(1) = rsStat!tidx_blks_hit & ""
    Else
      ClearStats
    End If
    If rsStat.State <> adStateClosed Then rsStat.Close
    Set rsStat = Nothing
  Else
    sv.ColumnHeaders.Add , , "Statistics"
    Set lvItem = sv.ListItems.Add(, "STA-" & GetID, "Statistics are only available with PostgreSQL 7.2 or higher.", "server", "server")
  End If
  
  Exit Sub
Err_Handler:
  If rsStat.State <> adStateClosed Then rsStat.Close
  Set rsStat = Nothing
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.svTable"
End Sub

Private Sub tvChecks(ByVal Node As MSComctlLib.Node)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvChecks(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
Dim chk As pgCheck

  If Node.Children = 0 Or Node.Children <> svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Text).Tables(Node.Parent.Text).Checks.Count Then
    While Not (Node.Child Is Nothing)
      tv.Nodes.Remove Node.Child.Index
    Wend
    For Each chk In svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Text).Tables(Node.Parent.Text).Checks
      Set chk.Tag = tv.Nodes.Add(Node.Key, tvwChild, "CHK-" & GetID, chk.Identifier, "check")
    Next chk
    Node.Text = "Checks (" & Node.Children & ")"
  End If
  lv.ColumnHeaders.Add , , "Check", lv.Width
  For Each chk In svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Text).Tables(Node.Parent.Text).Checks
    Set lvItem = lv.ListItems.Add(, "CHK-" & GetID, chk.Identifier, "check", "check")
  Next chk
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvChecks"
End Sub

Private Sub tvCheck(ByVal Node As MSComctlLib.Node)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvCheck(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem

  lv.ColumnHeaders.Add , , "Property"
  lv.ColumnHeaders.Add , , "Value"
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Name", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Name
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Definition", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Definition
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvCheck"
End Sub

Private Sub tvColumns(ByVal Node As MSComctlLib.Node)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvColumns(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
Dim col As pgColumn

  If Node.Children = 0 Or Node.Children <> svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Text).Tables(Node.Parent.Text).Columns.Count(Not ctx.IncludeSys) Then
    While Not (Node.Child Is Nothing)
      tv.Nodes.Remove Node.Child.Index
    Wend
    For Each col In svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Text).Tables(Node.Parent.Text).Columns
     If Not (col.SystemObject And Not ctx.IncludeSys) Then Set col.Tag = tv.Nodes.Add(Node.Key, tvwChild, "COL-" & GetID, col.Identifier, "column")
    Next col
    Node.Text = "Columns (" & Node.Children & ")"
  End If
  lv.ColumnHeaders.Add , , "Column"
  lv.ColumnHeaders.Add , , "Type"
  lv.ColumnHeaders.Add , , "Comment"
  For Each col In svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Text).Tables(Node.Parent.Text).Columns
    If Not (col.SystemObject And Not ctx.IncludeSys) Then
      Set lvItem = lv.ListItems.Add(, "COL-" & GetID, col.Identifier, "column", "column")
      lvItem.SubItems(1) = col.DataType
      lvItem.SubItems(2) = Replace(col.Comment, vbCrLf, " ")
    End If
  Next col
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvColumns"
End Sub

Private Sub tvColumn(ByVal Node As MSComctlLib.Node)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvColumn(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem

  lv.ColumnHeaders.Add , , "Property"
  lv.ColumnHeaders.Add , , "Value"
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Name", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Name
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Position", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Position
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Data Type", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.DataType
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Length", "property", "property")
  If ctx.CurrentObject.Length = 0 Then
    lvItem.SubItems(1) = "Variable"
  Else
    lvItem.SubItems(1) = ctx.CurrentObject.Length
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Numeric Precision", "property", "property")
  If ctx.CurrentObject.DataType = "numeric" Then
    lvItem.SubItems(1) = ctx.CurrentObject.NumericScale
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Default", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Default
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Restrict Nulls?", "property", "property")
  If ctx.CurrentObject.NotNull Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Primary Key?", "property", "property")
  If ctx.CurrentObject.PrimaryKey Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Statistics", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Statistics
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "System Column?", "property", "property")
  If ctx.CurrentObject.SystemObject Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Comment", "property", "property")
  lvItem.SubItems(1) = Replace(ctx.CurrentObject.Comment, vbCrLf, " ")
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvColumn"
End Sub

Private Sub svColumn(ByVal Node As MSComctlLib.Node)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.svDatabase(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
Dim rsStat As New Recordset
Dim szSQL As String

  ' Statistics.
  ' These don't come from pgSchema because they aren't really schema related.
  If ctx.dbVer >= 7.2 Then
    If ctx.dbVer >= 7.3 Then
      szSQL = "SELECT null_frac, avg_width, n_distinct, most_common_vals, most_common_freqs, histogram_bounds, correlation FROM pg_stats "
      szSQL = szSQL & "WHERE tablename = '" & Node.Parent.Parent.Text & "' AND attname = '" & Node.Text & "' and schemaname='" & ctx.CurrentNS & "'"
    Else
      szSQL = "SELECT null_frac, avg_width, n_distinct, most_common_vals, most_common_freqs, histogram_bounds, correlation FROM pg_stats WHERE tablename = '" & Node.Parent.Parent.Text & "' AND attname = '" & Node.Text & "'"
    End If
    Set rsStat = svr.Databases(ctx.CurrentDB).Execute(szSQL)
    sv.ColumnHeaders.Add , , "Statistic"
    sv.ColumnHeaders.Add , , "Value"
  
    If Not rsStat.EOF Then
      Set lvItem = sv.ListItems.Add(, "STA-" & GetID, "Null Fraction", "statistics", "statistics")
      lvItem.SubItems(1) = rsStat!null_frac & ""
      Set lvItem = sv.ListItems.Add(, "STA-" & GetID, "Average Width", "statistics", "statistics")
      lvItem.SubItems(1) = rsStat!avg_width & ""
      Set lvItem = sv.ListItems.Add(, "STA-" & GetID, "Distinct Values", "statistics", "statistics")
      lvItem.SubItems(1) = rsStat!n_distinct & ""
      Set lvItem = sv.ListItems.Add(, "STA-" & GetID, "Most Column Values", "statistics", "statistics")
      lvItem.SubItems(1) = rsStat!most_common_vals & ""
      Set lvItem = sv.ListItems.Add(, "STA-" & GetID, "Most Common Frequencies", "statistics", "statistics")
      lvItem.SubItems(1) = rsStat!most_common_freqs & ""
      Set lvItem = sv.ListItems.Add(, "STA-" & GetID, "Histogram Bounds", "statistics", "statistics")
      lvItem.SubItems(1) = rsStat!histogram_bounds & ""
      Set lvItem = sv.ListItems.Add(, "STA-" & GetID, "Correlation", "statistics", "statistics")
      lvItem.SubItems(1) = rsStat!correlation & ""
    Else
      ClearStats
    End If
    If rsStat.State <> adStateClosed Then rsStat.Close
    Set rsStat = Nothing
  Else
    sv.ColumnHeaders.Add , , "Statistics"
    Set lvItem = sv.ListItems.Add(, "STA-" & GetID, "Statistics are only available with PostgreSQL 7.2 or higher.", "server", "server")
  End If
  
  Exit Sub
Err_Handler:
  If rsStat.State <> adStateClosed Then rsStat.Close
  Set rsStat = Nothing
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.svColumn"
End Sub

Private Sub tvConversions(ByVal Node As MSComctlLib.Node)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvConversions(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
Dim conv As pgConversion

  If Node.Children = 0 Or Node.Children <> svr.Databases(Node.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Text).Conversions.Count(Not ctx.IncludeSys) Then
    While Not (Node.Child Is Nothing)
      tv.Nodes.Remove Node.Child.Index
    Wend
    For Each conv In svr.Databases(Node.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Text).Conversions
      If Not (conv.SystemObject And Not ctx.IncludeSys) Then Set conv.Tag = tv.Nodes.Add(Node.Key, tvwChild, "CNV-" & GetID, conv.Identifier, "conversion")
    Next
    Node.Text = "Conversions (" & Node.Children & ")"
  End If
  lv.ColumnHeaders.Add , , "Conversion"
  For Each conv In svr.Databases(Node.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Text).Conversions
    If Not (conv.SystemObject And Not ctx.IncludeSys) Then
      Set lvItem = lv.ListItems.Add(, "CNV-" & GetID, conv.Identifier, "conversion", "conversion")
    End If
  Next
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvConversions"
End Sub

Private Sub tvConversion(ByVal Node As MSComctlLib.Node)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvConversion(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
  
  lv.ColumnHeaders.Add , , "Property"
  lv.ColumnHeaders.Add , , "Value"
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Name", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Name
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "OID", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Oid
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Owner", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Owner
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Default", "property", "property")
  If ctx.CurrentObject.Default Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Source", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.ForEncoding
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Destination", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.ToEncoding
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Function", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Proc
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "System Conversion?", "property", "property")
  If ctx.CurrentObject.SystemObject Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If

  'Set the Definition Pane
  If txtDefinition.Visible Then txtDefinition.Text = ctx.CurrentObject.SQL
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvConversion"
End Sub

Private Sub tvForeignKeys(ByVal Node As MSComctlLib.Node)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvForeignKeys(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
Dim fky As pgForeignKey

  If Node.Children = 0 Or Node.Children <> svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Text).Tables(Node.Parent.Text).ForeignKeys.Count(Not ctx.IncludeSys) Then
    While Not (Node.Child Is Nothing)
      tv.Nodes.Remove Node.Child.Index
    Wend
    For Each fky In svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Text).Tables(Node.Parent.Text).ForeignKeys
      If Not (fky.SystemObject And Not ctx.IncludeSys) Then Set fky.Tag = tv.Nodes.Add(Node.Key, tvwChild, "FKY-" & GetID, fky.Identifier, "foreignkey")
    Next fky
    Node.Text = "Foreign Keys (" & Node.Children & ")"
  End If
  lv.ColumnHeaders.Add , , "Foreign Key"
  lv.ColumnHeaders.Add , , "References"
  For Each fky In svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Text).Tables(Node.Parent.Text).ForeignKeys
    If Not (fky.SystemObject And Not ctx.IncludeSys) Then
      Set lvItem = lv.ListItems.Add(, "FKY-" & GetID, fky.Identifier, "foreignkey", "foreignkey")
      lvItem.SubItems(1) = fky.ReferencedTable
    End If
  Next fky
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvForeignKeys"
End Sub

Private Sub tvForeignKey(ByVal Node As MSComctlLib.Node)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvForeignKey(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem

  If Node.Children = 0 Then tv.Nodes.Add Node.Key, tvwChild, "REL+" & GetID, "Relationships (" & ctx.CurrentObject.Relationships.Count & ")", "relationship"
  lv.ColumnHeaders.Add , , "Property"
  lv.ColumnHeaders.Add , , "Value"
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Name", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Name
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "OID", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Oid
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "References", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.ReferencedTable
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "On Delete", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.OnDelete
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "On Update", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.OnUpdate
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Deferrable", "property", "property")
  If ctx.CurrentObject.Deferrable Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Initially", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Initially
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "System Foreign Key?", "property", "property")
  If ctx.CurrentObject.SystemObject Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvForeignKey"
End Sub

Private Sub tvRelationships(ByVal Node As MSComctlLib.Node)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvRelationships(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
Dim rel As pgRelationship

  lv.ColumnHeaders.Add , , "Local Column"
  lv.ColumnHeaders.Add , , "Referenced Column"
  Node.Text = "Relationships (" & svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Parent.Text).ForeignKeys(Node.Parent.Text).Relationships.Count & ")"
  For Each rel In svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Parent.Text).ForeignKeys(Node.Parent.Text).Relationships
    Set lvItem = lv.ListItems.Add(, "REL-" & GetID, rel.LocalColumn, "relationship", "relationship")
    lvItem.SubItems(1) = rel.ReferencedColumn
  Next rel
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvRelationships"
End Sub

Private Sub tvIndexes(ByVal Node As MSComctlLib.Node)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvIndexes(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
Dim ind As pgIndex

  If Node.Children = 0 Or Node.Children <> svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Text).Tables(Node.Parent.Text).Indexes.Count(Not ctx.IncludeSys) Then
    While Not (Node.Child Is Nothing)
      tv.Nodes.Remove Node.Child.Index
    Wend
    For Each ind In svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Text).Tables(Node.Parent.Text).Indexes
      If Not (ind.SystemObject And Not ctx.IncludeSys) Then Set ind.Tag = tv.Nodes.Add(Node.Key, tvwChild, "IND-" & GetID, ind.Identifier, "index")
    Next ind
    Node.Text = "Indexes (" & Node.Children & ")"
  End If
  lv.ColumnHeaders.Add , , "Index"
  lv.ColumnHeaders.Add , , "Comment"
  For Each ind In svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Text).Tables(Node.Parent.Text).Indexes
    If Not (ind.SystemObject And Not ctx.IncludeSys) Then
      Set lvItem = lv.ListItems.Add(, "IND-" & GetID, ind.Identifier, "index", "index")
      lvItem.SubItems(1) = Replace(ind.Comment, vbCrLf, " ")
    End If
  Next ind
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvIndexes"
End Sub

Private Sub svIndexes(ByVal Node As MSComctlLib.Node)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.svIndexes(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
Dim rsStat As New Recordset
Dim szSQL As String

  ' Statistics.
  ' These don't come from pgSchema because they aren't really schema related.
  If ctx.dbVer >= 7.2 Then
    If ctx.dbVer >= 7.3 Then
      szSQL = "SELECT relname, indexrelname, idx_blks_read, idx_blks_hit FROM pg_statio_all_indexes "
      szSQL = szSQL & "WHERE relname = '" & Node.Parent.Text & "' and schemaname='" & ctx.CurrentNS & "' ORDER BY indexrelname"
    Else
      szSQL = "SELECT relname, indexrelname, idx_blks_read, idx_blks_hit FROM pg_statio_all_indexes WHERE relname = '" & Node.Parent.Text & "' ORDER BY indexrelname"
    End If
    Set rsStat = svr.Databases(ctx.CurrentDB).Execute(szSQL)
    sv.ColumnHeaders.Add , , "Index", 2000
    sv.ColumnHeaders.Add , , "Index Blocks Read", 2000
    sv.ColumnHeaders.Add , , "Index Blocks Hit", 2000
  
    While Not rsStat.EOF
      If svr.Databases(ctx.CurrentDB).Namespaces(ctx.CurrentNS).Tables(rsStat!relname).Indexes.Exists(rsStat!indexrelname) Then
        If Not (svr.Databases(ctx.CurrentDB).Namespaces(ctx.CurrentNS).Tables(rsStat!relname).Indexes(rsStat!indexrelname).SystemObject And Not ctx.IncludeSys) Then
          Set lvItem = sv.ListItems.Add(, "STA+" & GetID, rsStat!indexrelname & "", "statistics", "statistics")
          lvItem.SubItems(1) = rsStat!idx_blks_read & ""
          lvItem.SubItems(2) = rsStat!idx_blks_hit & ""
        End If
      End If
      rsStat.MoveNext
    Wend
    If rsStat.State <> adStateClosed Then rsStat.Close
    Set rsStat = Nothing
  Else
    sv.ColumnHeaders.Add , , "Statistics"
    Set lvItem = sv.ListItems.Add(, "STA+" & GetID, "Statistics are only available with PostgreSQL 7.2 or higher.", "server", "server")
  End If
  
  Exit Sub
Err_Handler:
  If rsStat.State <> adStateClosed Then rsStat.Close
  Set rsStat = Nothing
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.svIndexes"
End Sub

Private Sub tvIndex(ByVal Node As MSComctlLib.Node)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvIndex(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
Dim szTemp As String
Dim vData As Variant

  lv.ColumnHeaders.Add , , "Property"
  lv.ColumnHeaders.Add , , "Value"
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Name", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Name
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "OID", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Oid
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Index Type", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.IndexType
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Unique?", "property", "property")
  If ctx.CurrentObject.Unique Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Primary?", "property", "property")
  If ctx.CurrentObject.Primary Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Column Count", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.IndexedColumns.Count
  For Each vData In ctx.CurrentObject.IndexedColumns
    szTemp = szTemp & vData & ", "
  Next vData
  If Len(szTemp) > 2 Then szTemp = Left(szTemp, Len(szTemp) - 2)
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Columns", "property", "property")
  lvItem.SubItems(1) = szTemp
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Constraint", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Constraint
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "System Index?", "property", "property")
  If ctx.CurrentObject.SystemObject Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Comment", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Comment
  
  'Set the Definition Pane
  If txtDefinition.Visible Then txtDefinition.Text = ctx.CurrentObject.SQL

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvIndex"
End Sub

Private Sub svIndex(ByVal Node As MSComctlLib.Node)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.svIndex(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
Dim rsStat As New Recordset

  ' Statistics.
  ' These don't come from pgSchema because they aren't really schema related.
  If ctx.dbVer >= 7.2 Then
    Set rsStat = svr.Databases(ctx.CurrentDB).Execute("SELECT idx_scan, idx_tup_read, idx_tup_fetch, idx_blks_read, idx_blks_hit FROM pg_stat_all_indexes stat, pg_statio_all_indexes statio WHERE stat.relid = statio.relid AND stat.indexrelid = statio.indexrelid AND statio.indexrelid = " & ctx.CurrentObject.Oid & "::oid")
    sv.ColumnHeaders.Add , , "Statistic"
    sv.ColumnHeaders.Add , , "Value"
  
    If Not rsStat.EOF Then
      Set lvItem = sv.ListItems.Add(, "STA-" & GetID, "Index Scans", "statistics", "statistics")
      lvItem.SubItems(1) = rsStat!idx_scan & ""
      Set lvItem = sv.ListItems.Add(, "STA-" & GetID, "Index Tuples Read", "statistics", "statistics")
      lvItem.SubItems(1) = rsStat!idx_tup_read & ""
      Set lvItem = sv.ListItems.Add(, "STA-" & GetID, "Index Tuples Fetched", "statistics", "statistics")
      lvItem.SubItems(1) = rsStat!idx_tup_fetch & ""
      Set lvItem = sv.ListItems.Add(, "STA-" & GetID, "Index Blocks Read", "statistics", "statistics")
      lvItem.SubItems(1) = rsStat!idx_blks_read & ""
      Set lvItem = sv.ListItems.Add(, "STA-" & GetID, "Index Blocks Hit", "statistics", "statistics")
      lvItem.SubItems(1) = rsStat!idx_blks_hit & ""
    Else
      ClearStats
    End If
    If rsStat.State <> adStateClosed Then rsStat.Close
    Set rsStat = Nothing
  Else
    sv.ColumnHeaders.Add , , "Statistics"
    Set lvItem = sv.ListItems.Add(, "STA-" & GetID, "Statistics are only available with PostgreSQL 7.2 or higher.", "server", "server")
  End If
  
  Exit Sub
Err_Handler:
  If rsStat.State <> adStateClosed Then rsStat.Close
  Set rsStat = Nothing
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.svIndex"
End Sub

Private Sub tvRules(ByVal Node As MSComctlLib.Node)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvRules(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
Dim rul As pgRule
Dim objTmp

  'verify if rule is for table or view
  If svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Text).Tables.Exists(Node.Parent.Text) Then
    Set objTmp = svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Text).Tables
  ElseIf svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Text).Views.Exists(Node.Parent.Text) Then
    Set objTmp = svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Text).Views
  End If

  If Node.Children = 0 Or Node.Children <> objTmp(Node.Parent.Text).Rules.Count(Not ctx.IncludeSys) Then
    While Not (Node.Child Is Nothing)
      tv.Nodes.Remove Node.Child.Index
    Wend
    For Each rul In objTmp(Node.Parent.Text).Rules
      If Not (rul.SystemObject And Not ctx.IncludeSys) Then Set rul.Tag = tv.Nodes.Add(Node.Key, tvwChild, "RUL-" & GetID, rul.Identifier, "rule")
    Next rul
    Node.Text = "Rules (" & Node.Children & ")"
  End If
  lv.ColumnHeaders.Add , , "Rule"
  lv.ColumnHeaders.Add , , "Comment"
  For Each rul In objTmp(Node.Parent.Text).Rules
    If Not (rul.SystemObject And Not ctx.IncludeSys) Then
      Set lvItem = lv.ListItems.Add(, "RUL-" & GetID, rul.Identifier, "rule", "rule")
      lvItem.SubItems(1) = Replace(rul.Comment, vbCrLf, " ")
    End If
  Next rul
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvRules"
End Sub

Private Sub tvRule(ByVal Node As MSComctlLib.Node)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvRule(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem

  lv.ColumnHeaders.Add , , "Property"
  lv.ColumnHeaders.Add , , "Value"
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Name", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Name
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "OID", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Oid
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Event", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.RuleEvent
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Condition", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Condition
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Do Instead?", "property", "property")
  If ctx.CurrentObject.DoInstead Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Action", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Action
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Definition", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Definition
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "System Rule?", "property", "property")
  If ctx.CurrentObject.SystemObject Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Comment", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Comment
  
  'Set the Definition Pane
  If txtDefinition.Visible Then txtDefinition.Text = ctx.CurrentObject.SQL
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvRule"
End Sub

Private Sub tvTriggers(ByVal Node As MSComctlLib.Node)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvTriggers(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
Dim trg As pgTrigger

  If Node.Children = 0 Or Node.Children <> svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Text).Tables(Node.Parent.Text).Triggers.Count(Not ctx.IncludeSys) Then
    While Not (Node.Child Is Nothing)
      tv.Nodes.Remove Node.Child.Index
    Wend
    For Each trg In svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Text).Tables(Node.Parent.Text).Triggers
      If Not (trg.SystemObject And Not ctx.IncludeSys) Then Set trg.Tag = tv.Nodes.Add(Node.Key, tvwChild, "TRG-" & GetID, trg.Identifier, "trigger")
    Next trg
    Node.Text = "Triggers (" & Node.Children & ")"
  End If
  lv.ColumnHeaders.Add , , "Trigger"
  lv.ColumnHeaders.Add , , "Comment"
  For Each trg In svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Text).Tables(Node.Parent.Text).Triggers
    If Not (trg.SystemObject And Not ctx.IncludeSys) Then
      Set lvItem = lv.ListItems.Add(, "TRG-" & GetID, trg.Identifier, "trigger", "trigger")
      lvItem.SubItems(1) = Replace(trg.Comment, vbCrLf, " ")
    End If
  Next trg
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvTriggers"
End Sub

Private Sub tvTrigger(ByVal Node As MSComctlLib.Node)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvTrigger(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem

  lv.ColumnHeaders.Add , , "Property"
  lv.ColumnHeaders.Add , , "Value"
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Name", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Name
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "OID", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Oid
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Executes", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Executes
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Event", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.TriggerEvent
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "For Each", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.ForEach
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Function", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.TriggerFunction
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "System Trigger?", "property", "property")
  If ctx.CurrentObject.SystemObject Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Comment", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Comment
  
  'Set the Definition Pane
  If txtDefinition.Visible Then txtDefinition.Text = ctx.CurrentObject.SQL
    
  'add function to call trigger
  If Node.Children = 0 Then
    tv.Nodes.Add Node.Key, tvwChild, "FNT-" & GetID, ctx.CurrentObject.TriggerFunction, "function"
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvTrigger"
End Sub

Private Sub tvTypes(ByVal Node As MSComctlLib.Node)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvTypes(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
Dim typ As pgType

  If Node.Children = 0 Or Node.Children <> svr.Databases(Node.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Text).Types.Count(Not ctx.IncludeSys) Then
    While Not (Node.Child Is Nothing)
      tv.Nodes.Remove Node.Child.Index
    Wend
    For Each typ In svr.Databases(Node.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Text).Types
      If Not (typ.SystemObject And Not ctx.IncludeSys) Then Set typ.Tag = tv.Nodes.Add(Node.Key, tvwChild, "TYP-" & GetID, typ.Identifier, "type")
    Next typ
    Node.Text = "Types (" & Node.Children & ")"
  End If
  lv.ColumnHeaders.Add , , "Type"
  lv.ColumnHeaders.Add , , "Comment"
  For Each typ In svr.Databases(Node.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Text).Types
    If Not (typ.SystemObject And Not ctx.IncludeSys) Then
      Set lvItem = lv.ListItems.Add(, "TYP-" & GetID, typ.Identifier, "type", "type")
      lvItem.SubItems(1) = Replace(typ.Comment, vbCrLf, " ")
    End If
  Next typ
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvTypes"
End Sub

Private Sub tvType(ByVal Node As MSComctlLib.Node)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvType(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem

  lv.ColumnHeaders.Add , , "Property"
  lv.ColumnHeaders.Add , , "Value"
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Name", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Name
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "OID", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Oid
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Owner", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Owner
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Input Function", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.InputFunction
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Output Function", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.OutputFunction
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Internal Length", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.InternalLength
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Default", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Default
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Element", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Element
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Delimiter", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Delimiter
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Passed by Value?", "property", "property")
  If ctx.CurrentObject.PassedByValue Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Alignment", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Alignment
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Storage", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Storage
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "System Type?", "property", "property")
  If ctx.CurrentObject.SystemObject Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Comment", "property", "property")
  lvItem.SubItems(1) = Replace(ctx.CurrentObject.Comment, vbCrLf, " ")

  'Set the Definition Pane
  If txtDefinition.Visible Then txtDefinition.Text = ctx.CurrentObject.SQL
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvType"
End Sub

Private Sub tvViews(ByVal Node As MSComctlLib.Node)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvViews(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
Dim vie As pgView
  
  If Node.Children = 0 Or Node.Children <> svr.Databases(Node.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Text).Views.Count(Not ctx.IncludeSys) Then
    While Not (Node.Child Is Nothing)
      tv.Nodes.Remove Node.Child.Index
    Wend
    For Each vie In svr.Databases(Node.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Text).Views
      If Not (vie.SystemObject And Not ctx.IncludeSys) Then Set vie.Tag = tv.Nodes.Add(Node.Key, tvwChild, "VIE-" & GetID, vie.Identifier, "view")
    Next vie
    Node.Text = "Views (" & Node.Children & ")"
  End If
  lv.ColumnHeaders.Add , , "View"
  lv.ColumnHeaders.Add , , "Comment"
  For Each vie In svr.Databases(Node.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Text).Views
    If Not (vie.SystemObject And Not ctx.IncludeSys) Then
      Set lvItem = lv.ListItems.Add(, "VIE-" & GetID, vie.Identifier, "view", "view")
      lvItem.SubItems(1) = Replace(vie.Comment, vbCrLf, " ")
    End If
  Next vie
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvViews"
End Sub

Private Sub tvView(ByVal Node As MSComctlLib.Node)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvView(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem

  lv.ColumnHeaders.Add , , "Property"
  lv.ColumnHeaders.Add , , "Value"
  If Node.Children = 0 Then
    Set ctx.CurrentObject.Rules.Tag = tv.Nodes.Add(Node.Key, tvwChild, "RUL+" & GetID, "Rules (" & ctx.CurrentObject.Rules.Count(Not ctx.IncludeSys) & ")", "rule")
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Name", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Name
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "OID", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Oid
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Owner", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Owner
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "ACL", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.ACL
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Rows", "property", "property")
  If ctx.AutoRowCount Then
    lvItem.SubItems(1) = ctx.CurrentObject.Rows
  Else
    lvItem.SubItems(1) = "Unknown"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Definition", "property", "property")
  lvItem.SubItems(1) = ctx.CurrentObject.Definition
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "System View?", "property", "property")
  If ctx.CurrentObject.SystemObject Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Comment", "property", "property")
  lvItem.SubItems(1) = Replace(ctx.CurrentObject.Comment, vbCrLf, " ")
  
  'Set the Definition Pane
  If txtDefinition.Visible Then txtDefinition.Text = ctx.CurrentObject.SQL
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvView"
End Sub

Public Sub ClearStats()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.ClearStats()", etFullDebug

  sv.ColumnHeaders.Clear
  sv.ListItems.Clear

  sv.ColumnHeaders.Add , , "Statistics", sv.Width
  sv.ListItems.Add , , "No Statistics are available for the current selection", "statistics", "statistics"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.ClearStats"
End Sub

Public Sub tv_NodeClick(ByVal Node As MSComctlLib.Node)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tv_NodeClick(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
Dim szTemp As String
Dim vData As Variant

  StartMsg "Examining database..."
  
  lv.ColumnHeaders.Clear
  lv.ListItems.Clear
  lv.Tag = Node.FullPath
  sv.ColumnHeaders.Clear
  sv.ListItems.Clear
  If txtDefinition.Visible Then txtDefinition.Text = ""
  
  'Stats are only on 7.2+
  If ctx.dbVer < 7.2 Then
    sv.ColumnHeaders.Add , , "Statistics", sv.Width
    sv.ListItems.Add , , "Statistics are only available with PostgreSQL 7.2 or higher", "database", "database"
  End If
  
  Select Case Left(Node.Key, 4)

    Case "SVR-" 'Server
      ctx.CurrentDB = ""
      ctx.CurrentNS = ""
      Set ctx.CurrentObject = svr
      tvServer Node
      If ctx.dbVer >= 7.2 Then svServer Node
      tvDepend Node
      lvLocks Node

    Case "DAT+" 'Databases
      ctx.CurrentDB = ""
      ctx.CurrentNS = ""
      tvDatabases Node
      If ctx.dbVer >= 7.2 Then svDatabases Node
      tvDepend Node
      lvLocks Node
        
    Case "DAT-" 'Database
      ctx.CurrentDB = Node.Text
      ctx.CurrentNS = ""
      Set ctx.CurrentObject = svr.Databases(Node.Text)
      tvDatabase Node
      If ctx.dbVer >= 7.2 Then svDatabase Node
      tvDepend Node
      lvLocks Node
      
    Case "GRP+" 'Groups
      ctx.CurrentDB = ""
      ctx.CurrentNS = ""
      tvGroups Node
      If ctx.dbVer >= 7.2 Then ClearStats
      tvDepend Node
      lvLocks Node
      
    Case "GRP-" 'Group
      ctx.CurrentDB = ""
      ctx.CurrentNS = ""
      Set ctx.CurrentObject = svr.Groups(Node.Text)
      tvGroup Node
      If ctx.dbVer >= 7.2 Then ClearStats
      tvDepend Node
      lvLocks Node
      
    Case "USR+" 'Users
      ctx.CurrentDB = ""
      ctx.CurrentNS = ""
      tvUsers Node
      If ctx.dbVer >= 7.2 Then ClearStats
      tvDepend Node
      lvLocks Node

    Case "USR-" 'User
      ctx.CurrentDB = ""
      ctx.CurrentNS = ""
      Set ctx.CurrentObject = svr.Users(Node.Text)
      tvUser Node
      If ctx.dbVer >= 7.2 Then ClearStats
      tvDepend Node
      lvLocks Node
      
    Case "CST+" 'Casts
      ctx.CurrentDB = Node.Parent.Text
      ctx.CurrentNS = ""
      tvCasts Node
      If ctx.dbVer >= 7.2 Then ClearStats
      tvDepend Node
      lvLocks Node
    
    Case "CST-" 'Cast
      ctx.CurrentDB = Node.Parent.Parent.Text
      ctx.CurrentNS = ""
      Set ctx.CurrentObject = svr.Databases(Node.Parent.Parent.Text).Casts(Node.Text)
      tvCast Node
      If ctx.dbVer >= 7.2 Then ClearStats
      tvDepend Node
      lvLocks Node
      
    Case "LNG+" 'Languages
      ctx.CurrentDB = Node.Parent.Text
      ctx.CurrentNS = ""
      tvLanguages Node
      If ctx.dbVer >= 7.2 Then ClearStats
      tvDepend Node
      lvLocks Node

    Case "LNG-" 'Language
      ctx.CurrentDB = Node.Parent.Parent.Text
      ctx.CurrentNS = ""
      Set ctx.CurrentObject = svr.Databases(Node.Parent.Parent.Text).Languages(Node.Text)
      tvLanguage Node
      If ctx.dbVer >= 7.2 Then ClearStats
      tvDepend Node
      lvLocks Node
      
    Case "NSP+" 'Namespaces
      ctx.CurrentDB = Node.Parent.Text
      ctx.CurrentNS = ""
      tvNamespaces Node
      If ctx.dbVer >= 7.2 Then ClearStats
      tvDepend Node
      lvLocks Node

    Case "NSP-" 'Namespaces
      ctx.CurrentDB = Node.Parent.Parent.Text
      ctx.CurrentNS = Node.Text
      Set ctx.CurrentObject = svr.Databases(Node.Parent.Parent.Text).Namespaces(Node.Text)
      tvNamespace Node
      If ctx.dbVer >= 7.2 Then ClearStats
      tvDepend Node
      lvLocks Node
      
    Case "AGG+" 'Aggregates
      ctx.CurrentDB = Node.Parent.Parent.Parent.Text
      ctx.CurrentNS = Node.Parent.Text
      tvAggregates Node
      If ctx.dbVer >= 7.2 Then ClearStats
      tvDepend Node
      lvLocks Node
      
    Case "AGG-" 'Aggregate
      ctx.CurrentDB = Node.Parent.Parent.Parent.Parent.Text
      Set ctx.CurrentObject = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Aggregates(Node.Text)
      ctx.CurrentNS = Node.Parent.Parent.Text
      tvAggregate Node
      If ctx.dbVer >= 7.2 Then ClearStats
      tvDepend Node
      lvLocks Node
      
    Case "DOM+" 'Domains
      ctx.CurrentDB = Node.Parent.Parent.Parent.Text
      ctx.CurrentNS = Node.Parent.Text
      tvDomains Node
      If ctx.dbVer >= 7.2 Then ClearStats
      tvDepend Node
      lvLocks Node
      
    Case "DOM-" 'Domain
      ctx.CurrentDB = Node.Parent.Parent.Parent.Parent.Text
      Set ctx.CurrentObject = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Domains(Node.Text)
      ctx.CurrentNS = Node.Parent.Parent.Text
      tvDomain Node
      If ctx.dbVer >= 7.2 Then ClearStats
      tvDepend Node
      lvLocks Node
      
    Case "CNV+" 'Conversion
      ctx.CurrentDB = Node.Parent.Parent.Parent.Text
      ctx.CurrentNS = Node.Parent.Text
      tvConversions Node
      If ctx.dbVer >= 7.2 Then ClearStats
      tvDepend Node
      lvLocks Node
      
    Case "CNV-" 'Conversion
      ctx.CurrentDB = Node.Parent.Parent.Parent.Parent.Text
      Set ctx.CurrentObject = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Conversions(Node.Text)
      ctx.CurrentNS = Node.Parent.Parent.Text
      tvConversion Node
      If ctx.dbVer >= 7.2 Then ClearStats
      tvDepend Node
      lvLocks Node
      
    Case "FNC+" 'Functions
      ctx.CurrentDB = Node.Parent.Parent.Parent.Text
      ctx.CurrentNS = Node.Parent.Text
      tvFunctions Node
      If ctx.dbVer >= 7.2 Then ClearStats
      tvDepend Node
      lvLocks Node
      
    Case "FNC-" 'Function
      ctx.CurrentDB = Node.Parent.Parent.Parent.Parent.Text
      Set ctx.CurrentObject = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Functions(Node.Text)
      ctx.CurrentNS = Node.Parent.Parent.Text
      tvFunction Node
      If ctx.dbVer >= 7.2 Then ClearStats
      tvDepend Node
      lvLocks Node
      
    Case "FNT-" 'Function trigger
      ctx.CurrentDB = Node.Parent.Parent.Parent.Parent.Parent.Parent.Parent.Text
      Set ctx.CurrentObject = svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Parent.Parent.Text).Functions(Node.Text)
      ctx.CurrentNS = Node.Parent.Parent.Parent.Parent.Parent.Text
      tvFunction Node
      If ctx.dbVer >= 7.2 Then ClearStats
      tvDepend Node
      lvLocks Node
    
    Case "OPR+" 'Operators
      ctx.CurrentDB = Node.Parent.Parent.Parent.Text
      ctx.CurrentNS = Node.Parent.Text
      tvOperators Node
      If ctx.dbVer >= 7.2 Then ClearStats
      tvDepend Node
      lvLocks Node
      
    Case "OPR-" 'Operator
      ctx.CurrentDB = Node.Parent.Parent.Parent.Parent.Text
      Set ctx.CurrentObject = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Operators(Node.Text)
      ctx.CurrentNS = Node.Parent.Parent.Text
      tvOperator Node
      If ctx.dbVer >= 7.2 Then ClearStats
      tvDepend Node
      lvLocks Node
      
    Case "SEQ+" 'Sequences
      ctx.CurrentDB = Node.Parent.Parent.Parent.Text
      ctx.CurrentNS = Node.Parent.Text
      tvSequences Node
      If ctx.dbVer >= 7.2 Then svSequences Node
      tvDepend Node
      lvLocks Node

    Case "SEQ-" 'Sequence
      ctx.CurrentDB = Node.Parent.Parent.Parent.Parent.Text
      Set ctx.CurrentObject = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Sequences(Node.Text)
      ctx.CurrentNS = Node.Parent.Parent.Text
      tvSequence Node
      If ctx.dbVer >= 7.2 Then svSequence Node
      tvDepend Node
      lvLocks Node
      
    Case "TBL+" 'Tables
      ctx.CurrentDB = Node.Parent.Parent.Parent.Text
      ctx.CurrentNS = Node.Parent.Text
      tvTables Node
      If ctx.dbVer >= 7.2 Then svTables Node
      tvDepend Node
      lvLocks Node
      
    Case "TBL-" 'Table
      ctx.CurrentDB = Node.Parent.Parent.Parent.Parent.Text
      Set ctx.CurrentObject = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Tables(Node.Text)
      ctx.CurrentNS = Node.Parent.Parent.Text
      tvTable Node
      If ctx.dbVer >= 7.2 Then svTable Node
      tvDepend Node
      lvLocks Node
      
    Case "CHK+" 'Checks
      ctx.CurrentDB = Node.Parent.Parent.Parent.Parent.Parent.Text
      ctx.CurrentNS = Node.Parent.Parent.Parent.Text
      tvChecks Node
      If ctx.dbVer >= 7.2 Then ClearStats
      tvDepend Node
      lvLocks Node
      
    Case "CHK-" 'Check
      ctx.CurrentDB = Node.Parent.Parent.Parent.Parent.Parent.Parent.Text
      Set ctx.CurrentObject = svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Checks(Node.Text)
      ctx.CurrentNS = Node.Parent.Parent.Parent.Parent.Text
      tvCheck Node
      If ctx.dbVer >= 7.2 Then ClearStats
      tvDepend Node
      lvLocks Node
    
    Case "COL+" 'Columns
      ctx.CurrentDB = Node.Parent.Parent.Parent.Parent.Parent.Text
      ctx.CurrentNS = Node.Parent.Parent.Parent.Text
      tvColumns Node
      If ctx.dbVer >= 7.2 Then ClearStats
      tvDepend Node
      lvLocks Node
      
    Case "COL-" 'Column
      ctx.CurrentDB = Node.Parent.Parent.Parent.Parent.Parent.Parent.Text
      Set ctx.CurrentObject = svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Columns(Node.Text)
      ctx.CurrentNS = Node.Parent.Parent.Parent.Parent.Text
      tvColumn Node
      If ctx.dbVer >= 7.2 Then svColumn Node
      tvDepend Node
      lvLocks Node
      
    Case "FKY+" 'Foreign Keys
      ctx.CurrentDB = Node.Parent.Parent.Parent.Parent.Parent.Text
      ctx.CurrentNS = Node.Parent.Parent.Parent.Text
      tvForeignKeys Node
      If ctx.dbVer >= 7.2 Then ClearStats
      tvDepend Node
      lvLocks Node
      
    Case "FKY-" 'Foreign Key
      ctx.CurrentDB = Node.Parent.Parent.Parent.Parent.Parent.Parent.Text
      Set ctx.CurrentObject = svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).ForeignKeys(Node.Text)
      ctx.CurrentNS = Node.Parent.Parent.Parent.Parent.Text
      tvForeignKey Node
      If ctx.dbVer >= 7.2 Then ClearStats
      tvDepend Node
      lvLocks Node
      
    Case "REL+" 'Relationships
      ctx.CurrentDB = Node.Parent.Parent.Parent.Parent.Parent.Parent.Parent.Text
      ctx.CurrentNS = Node.Parent.Parent.Parent.Parent.Parent.Text
      tvRelationships Node
      If ctx.dbVer >= 7.2 Then ClearStats
      tvDepend Node
      lvLocks Node
      
    Case "IND+" 'Indexes
      ctx.CurrentDB = Node.Parent.Parent.Parent.Parent.Parent.Text
      ctx.CurrentNS = Node.Parent.Parent.Parent.Text
      tvIndexes Node
      If ctx.dbVer >= 7.2 Then svIndexes Node
      tvDepend Node
      lvLocks Node
      
    Case "IND-" 'Index
      ctx.CurrentDB = Node.Parent.Parent.Parent.Parent.Parent.Parent.Text
      Set ctx.CurrentObject = svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Indexes(Node.Text)
      ctx.CurrentNS = Node.Parent.Parent.Parent.Parent.Text
      tvIndex Node
      If ctx.dbVer >= 7.2 Then svIndex Node
      tvDepend Node
      lvLocks Node

    Case "RUL+" 'Rules
      ctx.CurrentDB = Node.Parent.Parent.Parent.Parent.Parent.Text
      ctx.CurrentNS = Node.Parent.Parent.Parent.Text
      tvRules Node
      If ctx.dbVer >= 7.2 Then ClearStats
      tvDepend Node
      lvLocks Node
  
    Case "RUL-" 'Rule
      ctx.CurrentDB = Node.Parent.Parent.Parent.Parent.Parent.Parent.Text
      'verify if rule is for table or view
      If svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Parent.Text).Tables.Exists(Node.Parent.Parent.Text) Then
        Set ctx.CurrentObject = svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Rules(Node.Text)
      ElseIf svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Parent.Text).Views.Exists(Node.Parent.Parent.Text) Then
        Set ctx.CurrentObject = svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Parent.Text).Views(Node.Parent.Parent.Text).Rules(Node.Text)
      End If
      ctx.CurrentNS = Node.Parent.Parent.Parent.Parent.Text
      tvRule Node
      If ctx.dbVer >= 7.2 Then ClearStats
      tvDepend Node
      lvLocks Node
      
    Case "TRG+" 'Triggers
      ctx.CurrentDB = Node.Parent.Parent.Parent.Parent.Parent.Text
      ctx.CurrentNS = Node.Parent.Parent.Parent.Text
      tvTriggers Node
      If ctx.dbVer >= 7.2 Then ClearStats
      tvDepend Node
      lvLocks Node
      
    Case "TRG-" 'Trigger
      ctx.CurrentDB = Node.Parent.Parent.Parent.Parent.Parent.Parent.Text
      Set ctx.CurrentObject = svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Triggers(Node.Text)
      ctx.CurrentNS = Node.Parent.Parent.Parent.Parent.Text
      tvTrigger Node
      If ctx.dbVer >= 7.2 Then ClearStats
      tvDepend Node
      lvLocks Node
      
    Case "TYP+" 'Types
      ctx.CurrentDB = Node.Parent.Parent.Parent.Text
      ctx.CurrentNS = Node.Parent.Text
      tvTypes Node
      If ctx.dbVer >= 7.2 Then ClearStats
      tvDepend Node
      lvLocks Node

    Case "TYP-" 'Type
      ctx.CurrentDB = Node.Parent.Parent.Parent.Parent.Text
      Set ctx.CurrentObject = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Types(Node.Text)
      ctx.CurrentNS = Node.Parent.Parent.Text
      tvType Node
      If ctx.dbVer >= 7.2 Then ClearStats
      tvDepend Node
      lvLocks Node
      
    Case "VIE+" 'Views
      ctx.CurrentDB = Node.Parent.Parent.Parent.Text
      ctx.CurrentNS = Node.Parent.Text
      tvViews Node
      If ctx.dbVer >= 7.2 Then ClearStats
      tvDepend Node
      lvLocks Node
      
    Case "VIE-" 'View
      ctx.CurrentDB = Node.Parent.Parent.Parent.Parent.Text
      Set ctx.CurrentObject = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Views(Node.Text)
      ctx.CurrentNS = Node.Parent.Parent.Text
      tvView Node
      If ctx.dbVer >= 7.2 Then ClearStats
      tvDepend Node
      lvLocks Node
    
  End Select
    
  AutoSizeColumnLv lv
  AutoSizeColumnLv sv
  EndMsg
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvNodeClick"
End Sub

Public Sub lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.lv_ItemClick(" & QUOTE & Item.Text & QUOTE & ")", etFullDebug

Dim szPath() As String

  'Get the elements of the node path. This will indicate the path through the pgSchema hierarchy
  szPath = Split(lv.Tag, "\")
  
  Select Case Left(Item.Key, 4)

    Case "SVR-" 'Server
      Set ctx.CurrentObject = svr
      ctx.CurrentDB = ""
      ctx.CurrentNS = ""
      If txtDefinition.Visible Then txtDefinition.Text = ""
        
    Case "DAT-" 'Database
      Set ctx.CurrentObject = svr.Databases(Item.Text)
      ctx.CurrentDB = Item.Text
      ctx.CurrentNS = ""
      If txtDefinition.Visible Then txtDefinition.Text = ctx.CurrentObject.SQL
      
    Case "GRP-" 'Group
      Set ctx.CurrentObject = svr.Groups(Item.Text)
      ctx.CurrentDB = ""
      ctx.CurrentNS = ""
      If txtDefinition.Visible Then txtDefinition.Text = ctx.CurrentObject.SQL
      
    Case "USR-" 'User
      Set ctx.CurrentObject = svr.Users(Item.Text)
      ctx.CurrentDB = ""
      ctx.CurrentNS = ""
      If txtDefinition.Visible Then txtDefinition.Text = ctx.CurrentObject.SQL
      
    Case "CST-" 'Cast
      Set ctx.CurrentObject = svr.Databases(szPath(2)).Casts(Item.Text)
      ctx.CurrentDB = ctx.CurrentObject.Database
      ctx.CurrentNS = ""
      If txtDefinition.Visible Then txtDefinition.Text = ctx.CurrentObject.SQL
      
    Case "LNG-" 'Language
      Set ctx.CurrentObject = svr.Databases(szPath(2)).Languages(Item.Text)
      ctx.CurrentDB = ctx.CurrentObject.Database
      ctx.CurrentNS = ""
      If txtDefinition.Visible Then txtDefinition.Text = ctx.CurrentObject.SQL

    Case "NSP-" 'Namespace
      Set ctx.CurrentObject = svr.Databases(szPath(2)).Namespaces(Item.Text)
      ctx.CurrentDB = ctx.CurrentObject.Database
      ctx.CurrentNS = Item.Text
      If txtDefinition.Visible Then txtDefinition.Text = ctx.CurrentObject.SQL
      
    Case "AGG-" 'Aggregate
      Set ctx.CurrentObject = svr.Databases(szPath(2)).Namespaces(szPath(4)).Aggregates(Item.Text)
      ctx.CurrentDB = ctx.CurrentObject.Database
      ctx.CurrentNS = ctx.CurrentObject.Namespace
      If txtDefinition.Visible Then txtDefinition.Text = ctx.CurrentObject.SQL
      
    Case "DOM-" 'Domain
      Set ctx.CurrentObject = svr.Databases(szPath(2)).Namespaces(szPath(4)).Domains(Item.Text)
      ctx.CurrentDB = ctx.CurrentObject.Database
      ctx.CurrentNS = ctx.CurrentObject.Namespace
      If txtDefinition.Visible Then txtDefinition.Text = ctx.CurrentObject.SQL

    Case "CNV-" 'Conversion
      Set ctx.CurrentObject = svr.Databases(szPath(2)).Namespaces(szPath(4)).Conversions(Item.Text)
      ctx.CurrentDB = ctx.CurrentObject.Database
      ctx.CurrentNS = ctx.CurrentObject.Namespace
      If txtDefinition.Visible Then txtDefinition.Text = ctx.CurrentObject.SQL

    Case "FNC-" 'Function
      Set ctx.CurrentObject = svr.Databases(szPath(2)).Namespaces(szPath(4)).Functions(Item.Text)
      ctx.CurrentDB = ctx.CurrentObject.Database
      ctx.CurrentNS = ctx.CurrentObject.Namespace
      If txtDefinition.Visible Then txtDefinition.Text = ctx.CurrentObject.SQL
      
    Case "OPR-" 'Operator
      Set ctx.CurrentObject = svr.Databases(szPath(2)).Namespaces(szPath(4)).Operators(Item.Text)
      ctx.CurrentDB = ctx.CurrentObject.Database
      ctx.CurrentNS = ctx.CurrentObject.Namespace
      If txtDefinition.Visible Then txtDefinition.Text = ctx.CurrentObject.SQL
 
    Case "SEQ-" 'Sequence
      Set ctx.CurrentObject = svr.Databases(szPath(2)).Namespaces(szPath(4)).Sequences(Item.Text)
      ctx.CurrentDB = ctx.CurrentObject.Database
      ctx.CurrentNS = ctx.CurrentObject.Namespace
      If txtDefinition.Visible Then txtDefinition.Text = ctx.CurrentObject.SQL

    Case "TBL-" 'Table
      Set ctx.CurrentObject = svr.Databases(szPath(2)).Namespaces(szPath(4)).Tables(Item.Text)
      ctx.CurrentDB = ctx.CurrentObject.Database
      ctx.CurrentNS = ctx.CurrentObject.Namespace
      If txtDefinition.Visible Then txtDefinition.Text = ctx.CurrentObject.SQL

    Case "CHK-" 'Check
      Set ctx.CurrentObject = svr.Databases(szPath(2)).Namespaces(szPath(4)).Tables(szPath(6)).Checks(Item.Text)
      ctx.CurrentDB = ctx.CurrentObject.Database
      ctx.CurrentNS = ctx.CurrentObject.Namespace
      If txtDefinition.Visible Then txtDefinition.Text = svr.Databases(ctx.CurrentDB).Namespaces(ctx.CurrentNS).Tables(ctx.CurrentObject.Table).SQL
      
    Case "COL-" 'Column
      Set ctx.CurrentObject = svr.Databases(szPath(2)).Namespaces(szPath(4)).Tables(szPath(6)).Columns(Item.Text)
      ctx.CurrentDB = ctx.CurrentObject.Database
      ctx.CurrentNS = ctx.CurrentObject.Namespace
      If txtDefinition.Visible Then txtDefinition.Text = svr.Databases(ctx.CurrentDB).Namespaces(ctx.CurrentNS).Tables(ctx.CurrentObject.Table).SQL

    Case "FKY-" 'Foreign Key
      Set ctx.CurrentObject = svr.Databases(szPath(2)).Namespaces(szPath(4)).Tables(szPath(6)).ForeignKeys(Item.Text)
      ctx.CurrentDB = ctx.CurrentObject.Database
      ctx.CurrentNS = ctx.CurrentObject.Namespace
      If txtDefinition.Visible Then txtDefinition.Text = svr.Databases(ctx.CurrentDB).Namespaces(ctx.CurrentNS).Tables(ctx.CurrentObject.Table).SQL
      
    Case "IND-" 'Index
      Set ctx.CurrentObject = svr.Databases(szPath(2)).Namespaces(szPath(4)).Tables(szPath(6)).Indexes(Item.Text)
      ctx.CurrentDB = ctx.CurrentObject.Database
      ctx.CurrentNS = ctx.CurrentObject.Namespace
      If txtDefinition.Visible Then txtDefinition.Text = ctx.CurrentObject.SQL
  
    Case "RUL-" 'Rule
      'verify if rule is for table or view
      If svr.Databases(szPath(2)).Namespaces(szPath(4)).Tables.Exists(szPath(6)) Then
        Set ctx.CurrentObject = svr.Databases(szPath(2)).Namespaces(szPath(4)).Tables(szPath(6)).Rules(Item.Text)
      Else
        Set ctx.CurrentObject = svr.Databases(szPath(2)).Namespaces(szPath(4)).Views(szPath(6)).Rules(Item.Text)
      End If
      ctx.CurrentDB = ctx.CurrentObject.Database
      ctx.CurrentNS = ctx.CurrentObject.Namespace
      If txtDefinition.Visible Then txtDefinition.Text = ctx.CurrentObject.SQL
      
    Case "TRG-" 'Trigger
      Set ctx.CurrentObject = svr.Databases(szPath(2)).Namespaces(szPath(4)).Tables(szPath(6)).Triggers(Item.Text)
      ctx.CurrentDB = ctx.CurrentObject.Database
      ctx.CurrentNS = ctx.CurrentObject.Namespace
      If txtDefinition.Visible Then txtDefinition.Text = ctx.CurrentObject.SQL

    Case "TYP-" 'Type
      Set ctx.CurrentObject = svr.Databases(szPath(2)).Namespaces(szPath(4)).Types(Item.Text)
      ctx.CurrentDB = ctx.CurrentObject.Database
      ctx.CurrentNS = ctx.CurrentObject.Namespace
      If txtDefinition.Visible Then txtDefinition.Text = ctx.CurrentObject.SQL
      
    Case "VIE-" 'View
      Set ctx.CurrentObject = svr.Databases(szPath(2)).Namespaces(szPath(4)).Views(Item.Text)
      ctx.CurrentDB = ctx.CurrentObject.Database
      ctx.CurrentNS = ctx.CurrentObject.Namespace
      If txtDefinition.Visible Then txtDefinition.Text = ctx.CurrentObject.SQL
  
  End Select
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.lv_ItemClick"
End Sub

Private Sub txtDefinition_Change()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.txtDefinition_Change()", etFullDebug
  
  If txtDefinition.Text = "" Then
    mnuFileSaveDefinition.Enabled = False
  Else
    mnuFileSaveDefinition.Enabled = True
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.txtDefinition_Change"
End Sub

Private Sub lv_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmMain.lv_ColumnClick(" & QUOTE & ColumnHeader.Text & QUOTE & ")", etFullDebug

  lv.Sorted = True
  'Sort by the select column. If we already are, then switch the direction.
  If lv.SortKey = (ColumnHeader.Index - 1) Then
    If lv.SortOrder = lvwAscending Then
      lv.SortOrder = lvwDescending
    Else
      lv.SortOrder = lvwAscending
    End If
  Else
    lv.SortOrder = lvwAscending
    lv.SortKey = (ColumnHeader.Index - 1)
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.lv_ColumnClick"
End Sub

Private Sub lvLock_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmMain.lvLock_ColumnClick(" & QUOTE & ColumnHeader.Text & QUOTE & ")", etFullDebug

  lvLock.Sorted = True
  'Sort by the select column. If we already are, then switch the direction.
  If lvLock.SortKey = (ColumnHeader.Index - 1) Then
    If lvLock.SortOrder = lvwAscending Then
      lvLock.SortOrder = lvwDescending
    Else
      lvLock.SortOrder = lvwAscending
    End If
  Else
    lvLock.SortOrder = lvwAscending
    lvLock.SortKey = (ColumnHeader.Index - 1)
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.lvLock_ColumnClick"
End Sub

Private Sub sv_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmMain.sv_ColumnClick(" & QUOTE & ColumnHeader.Text & QUOTE & ")", etFullDebug

  sv.Sorted = True
  'Sort by the select column. If we already are, then switch the direction.
  If sv.SortKey = (ColumnHeader.Index - 1) Then
    If sv.SortOrder = lvwAscending Then
      sv.SortOrder = lvwDescending
    Else
      sv.SortOrder = lvwAscending
    End If
  Else
    sv.SortOrder = lvwAscending
    sv.SortKey = (ColumnHeader.Index - 1)
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.sv_ColumnClick"
End Sub

Private Sub prop_Click(PreviousTab As Integer)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.prop_Click(" & PreviousTab & ")", etFullDebug
    
  If prop.Tab = 2 Then
    'refresh depending
    'Simulate a node click to refresh the ListDomain
    If Not tv.SelectedItem Is Nothing Then tv_NodeClick tv.SelectedItem
  ElseIf prop.Tab = 3 Then
    'refresh lock
    'Simulate a node click to refresh the ListDomain
    If Not tv.SelectedItem Is Nothing Then tv_NodeClick tv.SelectedItem
  End If

  Exit Sub
Err_Handler: LogError Err.Number, Err.Description, App.Title & ":frmMain.prop_Click"
End Sub

'show dependig object database
Private Sub tvDepend(ByVal Node As MSComctlLib.Node)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvDepend(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim objTmp
Dim objDep
Dim szKey As String

  ' Depending.
  tvDep.Nodes.Clear
  If prop.Tab <> 2 Then
    tvDep.Nodes.Add , , "DEP-" & GetID, "Dependencies are not applicable to the selected object.", "property", "property"
    Exit Sub
  ElseIf Len(ctx.CurrentDB) = 0 Then
    tvDep.Nodes.Add , , "DEP-" & GetID, "Dependencies are not applicable to the selected object.", "property", "property"
    Exit Sub
  ElseIf ctx.dbVer < 7.3 Then
    tvDep.Nodes.Add , , "DEP-" & GetID, "Dependencies are only available with PostgreSQL 7.3 or higher.", "property", "property"
    Exit Sub
  End If
  
  Select Case Left(Node.Key, 4)
    Case "CST-", "LNG-", "NSP-", "AGG-", "DOM-", "CNV-", "FNC-", "OPR-", "SEQ-", "TBL-", "TYP-", "VIE-"
      AddDepRef ctx.CurrentObject
    
    Case "CST+" 'Casts
      For Each objTmp In svr.Databases(ctx.CurrentDB).Casts
        AddDepRef objTmp
      Next
    
    Case "LNG+" 'Languages
      For Each objTmp In svr.Databases(ctx.CurrentDB).Languages
        AddDepRef objTmp
      Next

    Case "NSP+" 'Namespaces
      For Each objTmp In svr.Databases(ctx.CurrentDB).Namespaces
        AddDepRef objTmp
      Next
    
    Case "AGG+" 'Aggregates
      For Each objTmp In svr.Databases(ctx.CurrentDB).Namespaces(ctx.CurrentNS).Aggregates
        AddDepRef objTmp
      Next
      
    Case "DOM+" 'Domains
      For Each objTmp In svr.Databases(ctx.CurrentDB).Namespaces(ctx.CurrentNS).Domains
        AddDepRef objTmp
      Next
      
    Case "CNV+" 'Conversion
      For Each objTmp In svr.Databases(ctx.CurrentDB).Namespaces(ctx.CurrentNS).Conversions
        AddDepRef objTmp
      Next
      
    Case "FNC+" 'Functions
      For Each objTmp In svr.Databases(ctx.CurrentDB).Namespaces(ctx.CurrentNS).Functions
        AddDepRef objTmp
      Next
      
    Case "OPR+" 'Operators
      For Each objTmp In svr.Databases(ctx.CurrentDB).Namespaces(ctx.CurrentNS).Operators
        AddDepRef objTmp
      Next
      
    Case "SEQ+" 'Sequences
      For Each objTmp In svr.Databases(ctx.CurrentDB).Namespaces(ctx.CurrentNS).Sequences
        AddDepRef objTmp
      Next
      
    Case "TBL+" 'Tables
      For Each objTmp In svr.Databases(ctx.CurrentDB).Namespaces(ctx.CurrentNS).Tables
        AddDepRef objTmp
      Next
    
    Case "TYP+" 'Types
      For Each objTmp In svr.Databases(ctx.CurrentDB).Namespaces(ctx.CurrentNS).Types
        AddDepRef objTmp
      Next
    
    Case "VIE+" 'Views
      For Each objTmp In svr.Databases(ctx.CurrentDB).Namespaces(ctx.CurrentNS).Views
        AddDepRef objTmp
      Next
    
    Case Else
      tvDep.Nodes.Add , , "DEP-" & GetID, "Dependencies are not applicable to the selected object.", "property", "property"
      
  End Select
  
  Exit Sub
Err_Handler: LogError Err.Number, Err.Description, App.Title & ":frmMain.tvDepend"
End Sub

'add depend and reference
Private Sub AddDepRef(CurrentObj)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.AddDepRef(" & QUOTE & CurrentObj.ObjectType & QUOTE & ")", etFullDebug

Dim objDep
Dim szKey As String
Dim szKey1 As String
Dim szIdentifier As String
Dim szImg As String

  szKey = "DEP-" & GetID
  szImg = NameImageByObjectType(CurrentObj.ObjectType)
  tvDep.Nodes.Add , , szKey, CurrentObj.Identifier, szImg, szImg
  
  'add depend
  If CurrentObj.Dependent.Count > 0 Then
    szKey1 = "DEP-" & GetID
    tvDep.Nodes.Add szKey, tvwChild, szKey1, "Dependent Upon", "property", "property"
    For Each objDep In CurrentObj.Dependent
      
      szIdentifier = objDep.Identifier
      Select Case objDep.ObjectType
        Case "Aggregate", "Domain", "Conversion", "Function", "Operator", "Sequence", "Table", "Type", "View"
          szIdentifier = objDep.Namespace & "." & szIdentifier
      End Select
      
      szImg = NameImageByObjectType(objDep.ObjectType)
      tvDep.Nodes.Add szKey1, tvwChild, "DEP-" & GetID, szIdentifier, szImg, szImg
    Next
  End If
  
  'add reference
  If CurrentObj.Referenced.Count > 0 Then
    szKey1 = "REF-" & GetID
    tvDep.Nodes.Add szKey, tvwChild, szKey1, "Dependencies", "property", "property"
    For Each objDep In CurrentObj.Referenced
      
      szIdentifier = objDep.Identifier
      Select Case objDep.ObjectType
        Case "Aggregate", "Domain", "Conversion", "Function", "Operator", "Sequence", "Table", "Type", "View"
          szIdentifier = objDep.Namespace & "." & szIdentifier
      End Select
      
      szImg = NameImageByObjectType(objDep.ObjectType)
      tvDep.Nodes.Add szKey1, tvwChild, "REF-" & GetID, szIdentifier, szImg, szImg
    Next
  End If

  Exit Sub
Err_Handler: LogError Err.Number, Err.Description, App.Title & ":frmMain.AddDepRef"
End Sub

'show lock object database
Private Sub lvLocks(ByVal Node As MSComctlLib.Node)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.lvLocks(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug
 
  ' Lock.
  lvLock.ColumnHeaders.Clear
  lvLock.ListItems.Clear
  If prop.Tab <> 3 Then
    lvLock.ColumnHeaders.Add , , "Locks", lvLock.Width
    lvLock.ListItems.Add , , "Locks are not applicable to the selected object.", "property", "property"
    Exit Sub
  ElseIf ctx.dbVer < 7.3 Then
    lvLock.ColumnHeaders.Add , , "Locks", lvLock.Width
    lvLock.ListItems.Add , , "Locks are only available with PostgreSQL 7.3 or higher.", "property", "property"
    Exit Sub
  End If
  
  Select Case Left(Node.Key, 4)
    Case "SVR-", "USR+", "USR-", "DAT+", "DAT-", "NSP+", "NSP-"
      ShowLocks Left(Node.Key, 4)
    
    Case Else
      lvLock.ColumnHeaders.Add , , "Locks", lvLock.Width
      lvLock.ListItems.Add , , "Locks are not applicable to the selected object.", "property", "property"
    
  End Select
  Exit Sub

Err_Handler:
  EndMsg
  LogError Err.Number, Err.Description, App.Title & ":frmMain.lvLocks"
End Sub

Private Sub ShowLocks(ObjectType As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.ShowLocks(" & QUOTE & ObjectType & QUOTE & ")", etFullDebug

Dim szSQL As String
Dim szSqlLocks As String
Dim rsLocks As New Recordset
Dim rs As Recordset
Dim lvItem As ListItem
Dim szImg As String
Dim szUser As String
Dim szDatabase As String
Dim szNamespace As String
Dim szRelation As String
Dim iColumn As Integer

  If ObjectType <> "DAT-" And Left(ObjectType, 3) <> "NSP" Then lvLock.ColumnHeaders.Add , , "Database", 1500
  If ObjectType <> "NSP-" Then lvLock.ColumnHeaders.Add , , "Schema Name", 1500
  lvLock.ColumnHeaders.Add , , "Object Name", 1500
  If ObjectType <> "USR-" Then lvLock.ColumnHeaders.Add , , "User", 1500
  lvLock.ColumnHeaders.Add , , "Pid", 1500
  lvLock.ColumnHeaders.Add , , "Lock Mode", 2000

  StartMsg "Examining Locks..."
  
  szSqlLocks = "SELECT relation , database, transaction, pid, Mode, granted FROM pg_locks WHERE database IS NOT NULL"
  If ObjectType = "DAT-" Then
    'specify database
    szSqlLocks = szSqlLocks & " AND database=" & ctx.CurrentObject.Oid
  ElseIf Left(ObjectType, 3) = "NSP" Then
    'specify database for name space
    szSqlLocks = szSqlLocks & " AND database=" & svr.Databases(ctx.CurrentDB).Oid
  End If
  Set rsLocks = svr.Databases(svr.MasterDB).Execute(szSqlLocks & " ORDER BY pid")
  
  While Not rsLocks.EOF
    szUser = ""
    szSQL = "select usename from pg_stat_activity where procpid=" & rsLocks!pid
    Set rs = svr.Databases(svr.MasterDB).Execute(szSQL)
    If Not rs.EOF Then szUser = rs!usename & ""
    
    'filter user
    If ObjectType = "USR-" Then
      If ctx.CurrentObject.Name <> szUser Then GoTo NextLock
    End If
    
    szDatabase = ""
    If VarType(rsLocks!Database) <> vbNull Then
      szSQL = "SELECT datname FROM pg_database where oid=" & rsLocks!Database
      Set rs = svr.Databases(svr.MasterDB).Execute(szSQL)
      szDatabase = rs!datname & ""
    End If
    
    szNamespace = ""
    szRelation = ""
    szImg = "property"
    If VarType(rsLocks!relation) <> vbNull Then
      szSQL = "SELECT (SELECT n.nspname FROM pg_namespace n WHERE n.oid=c.relnamespace) as namespace, c.relname, c.relkind"
      szSQL = szSQL & " from pg_class c where oid=" & rsLocks!relation
      Set rs = svr.Databases(szDatabase).Execute(szSQL)
      If Not rs.EOF Then
        szNamespace = rs!Namespace & ""
        szRelation = rs!relname & ""

        Select Case rs!relkind
          Case "r"
            szImg = "table"
          Case "i"
            szImg = "index"
          Case "S"
            szImg = "sequence"
          Case "v"
            szImg = "view"
        End Select
      End If
    End If
    
    'filter name space
    If ObjectType = "NSP-" Then
      If ctx.CurrentObject.Name <> szNamespace Then GoTo NextLock
    End If
    
    If ObjectType <> "DAT-" And Left(ObjectType, 3) <> "NSP" Then
      Set lvItem = lvLock.ListItems.Add(, , szDatabase)
      lvItem.SubItems(1) = szNamespace
      lvItem.SubItems(2) = szRelation
      iColumn = 3
    ElseIf ObjectType = "NSP-" Then
      Set lvItem = lvLock.ListItems.Add(, , szRelation)
      iColumn = 1
    Else
      Set lvItem = lvLock.ListItems.Add(, , szNamespace)
      lvItem.SubItems(1) = szRelation
      iColumn = 2
    End If
    
    If ObjectType <> "USR-" Then lvItem.SubItems(iColumn) = szUser: iColumn = iColumn + 1
    lvItem.SubItems(iColumn) = rsLocks!pid: iColumn = iColumn + 1
    lvItem.SubItems(iColumn) = rsLocks!Mode: iColumn = iColumn + 1
    lvItem.SmallIcon = szImg
    lvItem.Icon = szImg

NextLock:
    rsLocks.MoveNext
  Wend
  
  If rsLocks.State <> adStateClosed Then rsLocks.Close
  Set rsLocks = Nothing
  
  AutoSizeColumnLv lvLock
  EndMsg
  Exit Sub

Err_Handler:
  EndMsg
  Set rsLocks = Nothing
  LogError Err.Number, Err.Description, App.Title & ":frmMain.ShowLocks"
End Sub

