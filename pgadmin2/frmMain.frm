VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{44F33AC4-8757-4330-B063-18608617F23E}#12.4#0"; "HighlightBox.ocx"
Begin VB.Form frmMain 
   Caption         =   "pgAdmin II"
   ClientHeight    =   6675
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9675
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6675
   ScaleWidth      =   9675
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   8550
      Top             =   90
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin HighlightBox.HBX txtDefinition 
      Height          =   2130
      Left            =   3825
      TabIndex        =   5
      ToolTipText     =   "Displays the SQL Definition of the currently selected object."
      Top             =   4185
      Width           =   5820
      _ExtentX        =   10266
      _ExtentY        =   3757
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
      Caption         =   "Definition"
   End
   Begin MSComctlLib.ImageList ilTB 
      Left            =   9090
      Top             =   45
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":08CA
            Key             =   "connect"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11A4
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A7E
            Key             =   "create"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2358
            Key             =   "drop"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3AEA
            Key             =   "properties"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":43C4
            Key             =   "sql"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4C9E
            Key             =   "vacuum"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4FB8
            Key             =   "viewdata"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      Height          =   5475
      Left            =   3645
      ScaleHeight     =   2384.051
      ScaleMode       =   0  'User
      ScaleWidth      =   780
      TabIndex        =   4
      Top             =   630
      Visible         =   0   'False
      Width           =   72
   End
   Begin MSComctlLib.Toolbar tb 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   1058
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      ImageList       =   "ilTB"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
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
               NumButtonMenus  =   14
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "aggregate"
                  Text            =   "&Aggregate"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "database"
                  Text            =   "&Database"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "function"
                  Text            =   "&Function"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "group"
                  Text            =   "&Group"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "index"
                  Text            =   "&Index"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "language"
                  Text            =   "&Language"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "operator"
                  Text            =   "&Operator"
               EndProperty
               BeginProperty ButtonMenu8 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "rule"
                  Text            =   "&Rule"
               EndProperty
               BeginProperty ButtonMenu9 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "sequence"
                  Text            =   "&Sequence"
               EndProperty
               BeginProperty ButtonMenu10 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "table"
                  Text            =   "&Table"
               EndProperty
               BeginProperty ButtonMenu11 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "trigger"
                  Text            =   "T&rigger"
               EndProperty
               BeginProperty ButtonMenu12 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "type"
                  Text            =   "T&ype"
               EndProperty
               BeginProperty ButtonMenu13 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "user"
                  Text            =   "&User"
               EndProperty
               BeginProperty ButtonMenu14 {66833FEE-8583-11D1-B16A-00C0F0283628} 
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
      EndProperty
   End
   Begin MSComctlLib.StatusBar sb 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   3
      Top             =   6390
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8918
            MinWidth        =   2
            Text            =   "Ready"
            TextSave        =   "Ready"
            Key             =   "info"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1111
            MinWidth        =   2
            Text            =   "0 Secs."
            TextSave        =   "0 Secs."
            Key             =   "timer"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   3043
            MinWidth        =   2
            Text            =   "Object: Not Connected"
            TextSave        =   "Object: Not Connected"
            Key             =   "currentobject"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   3440
            MinWidth        =   2
            Text            =   "Database: Not Connected"
            TextSave        =   "Database: Not Connected"
            Key             =   "currentdb"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList il 
      Left            =   4320
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   22
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5892
            Key             =   "aggregate"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":59EC
            Key             =   "check"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5F86
            Key             =   "column"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6520
            Key             =   "database"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":667A
            Key             =   "function"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6C14
            Key             =   "group"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":71AE
            Key             =   "index"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7748
            Key             =   "indexcolumn"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7CE2
            Key             =   "foreignkey"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":827C
            Key             =   "language"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8816
            Key             =   "operator"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8DB0
            Key             =   "property"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":934A
            Key             =   "relationship"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":98E4
            Key             =   "rule"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9A3E
            Key             =   "server"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9B98
            Key             =   "sequence"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A132
            Key             =   "table"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A28C
            Key             =   "trigger"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A826
            Key             =   "type"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":ADC0
            Key             =   "user"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B35A
            Key             =   "view"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B4B4
            Key             =   "baddatabase"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lv 
      Height          =   3480
      Left            =   3825
      TabIndex        =   2
      Top             =   675
      Width           =   5820
      _ExtentX        =   10266
      _ExtentY        =   6138
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
   Begin MSComctlLib.TreeView tv 
      Height          =   5460
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   3390
      _ExtentX        =   5980
      _ExtentY        =   9631
      _Version        =   393217
      Indentation     =   441
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "il"
      Appearance      =   1
   End
   Begin VB.Image imgSplitter 
      Height          =   5550
      Left            =   3510
      Top             =   630
      Width           =   60
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileConnect 
         Caption         =   "&Connect..."
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
      Begin VB.Menu mnuToolsUpgradeWizard 
         Caption         =   "&Upgrade Wizard..."
      End
      Begin VB.Menu mnuToolsSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsOptions 
         Caption         =   "&Options..."
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
      Begin VB.Menu mnuPopupHideSystemObjects 
         Caption         =   "Hide system objects"
      End
      Begin VB.Menu mnuPopupRefresh 
         Caption         =   "&Refresh below selection"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuPopupSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupCreate 
         Caption         =   "&Create object"
         Enabled         =   0   'False
         Begin VB.Menu mnuPopupCreateAggregate 
            Caption         =   "&Aggregate..."
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuPopupCreateDatabase 
            Caption         =   "&Database..."
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
      End
      Begin VB.Menu mnuPopupProperties 
         Caption         =   "&Properties..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuPopupSep2 
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
      Begin VB.Menu mnuPopupVacuum 
         Caption         =   "Vac&uum"
         Enabled         =   0   'False
         Begin VB.Menu mnuPopupVacuumVacuum 
            Caption         =   "&Vacuum"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuPopupVacuumAnalyse 
            Caption         =   "Vacuum &Analyse"
            Enabled         =   0   'False
         End
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence
'
' frmMain.frm - The primary form.

Option Explicit

'The Global Server Object. This must be in a form to be declared WithEvents
Public WithEvents svr As pgServer
Attribute svr.VB_VarHelpID = -1

'Indicates whether we are moving
Dim bMoving As Boolean


Private Sub Form_Resize()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.Form_Resize()", etFullDebug

  On Error Resume Next
  txtDefinition.Minimise
  If Me.Width < 8000 Then Me.Width = 8000
  If Me.Height < 6000 Then Me.Height = 6000
  SizeControls RegRead(HKEY_CURRENT_USER, "Software\" & App.Title, "Splitter Position")
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.Form_Resize"
End Sub

Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.imgSplitter_MouseDown(" & Button & ", " & Shift & ", " & X & ", " & Y & ")", etFullDebug

  With imgSplitter
    picSplitter.Move .Left, .Top, .Width \ 2, .Height - 20
  End With
  picSplitter.Visible = True
  bMoving = True
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.imgSplitter_MouseDown"
End Sub

Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.imgSplitter_MouseMove(" & Button & ", " & Shift & ", " & X & ", " & Y & ")", etFullDebug

Dim sglPos As Single
  
  If bMoving Then
    sglPos = X + imgSplitter.Left
    If sglPos < 500 Then
      picSplitter.Left = 500
    ElseIf sglPos > Me.Width - 500 Then
      picSplitter.Left = Me.Width - 500
    Else
      picSplitter.Left = sglPos
    End If
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.imgSplitter_MouseMove"
End Sub

Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.imgSplitter_MouseUp(" & Button & ", " & Shift & ", " & X & ", " & Y & ")", etFullDebug

  SizeControls picSplitter.Left
  RegWrite HKEY_CURRENT_USER, "Software\" & App.Title, "Splitter Position", regString, picSplitter.Left
  picSplitter.Visible = False
  bMoving = False
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.imgSplitter_MouseUp"
End Sub

Private Sub lv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.lv_MouseUp(" & Button & ", " & Shift & ", " & X & ", " & Y & ")", etFullDebug

  If Button = 2 Then PopupMenu frmMain.mnuPopup
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.lv_MouseUp"
End Sub

Private Sub mnuHelpContents_Click()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuHelpContents_Click()", etFullDebug

  HtmlHelp hWnd, App.Path & "\" & "help\pgadmin2.chm", HH_DISPLAY_TOPIC, 0

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuHelpContents_Click"
End Sub

Private Sub mnuHelpTipOfTheDay_Click()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuHelpTipOfTheDay_Click()", etFullDebug

  Load frmTip
  frmTip.Show

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuHelpTipOfTheDay_Click"
End Sub

Private Sub mnuPluginsPlg_Click(Index As Integer)
On Error GoTo Err_Handler
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
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuToolsUpgradeWizard_Click()", etFullDebug

  Load frmUpgradeWizard
  frmUpgradeWizard.Show

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuToolsUpgradeWizard_Click"
End Sub

Private Sub mnuViewSystemObjects_Click()
On Error GoTo Err_Handler
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
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tv_MouseUp(" & Button & ", " & Shift & ", " & X & ", " & Y & ")", etFullDebug

  If Button = 2 Then PopupMenu frmMain.mnuPopup

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tv_MouseUp"
End Sub

Private Sub mnuFileChangePassword_Click()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuFileChangePassword_Click()", etFullDebug

  Load frmPassword
  frmPassword.Show vbModal, Me
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuFileChangePassword_Click"
End Sub

Private Sub mnuFileSaveDbSchema_Click()
On Error GoTo Err_Handler
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
On Error GoTo Err_Handler
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
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuHelpAbout_Click()", etFullDebug

  Load frmAbout
  frmAbout.Show vbModal, Me
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuHelpAbout_Click"
End Sub

Private Sub mnuToolsOptions_Click()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuToolsOptions_Click()", etFullDebug

  Load frmOptions
  frmOptions.Show vbModal, Me
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuToolsOptions_Click"
End Sub

Private Sub mnuViewShowDefinitionPane_Click()
On Error GoTo Err_Handler
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
  SizeControls imgSplitter.Left
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuViewShowDefinitionPane_Click"
End Sub

Private Sub mnuViewShowLogWindow_Click()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuViewShowLogWindow_Click()", etFullDebug

  If mnuViewShowLogWindow.Checked = True Then
    ctx.LogView = False
    RegWrite HKEY_CURRENT_USER, "Software\" & App.Title & "\Log Window", "Visible", regString, "N"
    mnuViewShowLogWindow.Checked = False
    frmLog.Hide
  Else
    ctx.LogView = True
    RegWrite HKEY_CURRENT_USER, "Software\" & App.Title & "\Log Window", "Visible", regString, "Y"
    mnuViewShowLogWindow.Checked = True
    frmLog.Show
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuViewShowLogWindow_Click"
End Sub

Private Sub mnuViewShowStatusBar_Click()
On Error GoTo Err_Handler
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
  SizeControls imgSplitter.Left
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuViewShowStatusBar_Click"
End Sub

Private Sub mnuViewShowToolBar_Click()
On Error GoTo Err_Handler
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
  SizeControls imgSplitter.Left
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuViewShowToolBar_Click"
End Sub

Private Sub mnuPopupConnect_Click()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuPopupConnect_Click()", etFullDebug

  Load frmConnect
  frmConnect.Load_Defaults
  frmConnect.Show vbModal, Me
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuPopupConnect_Click"
End Sub

Private Sub mnuPopupRefresh_Click()
On Error GoTo Err_Handler
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
      svr.Connect
    Case "DAT+"
      svr.Databases.Refresh
    Case "GRP+"
      svr.Groups.Refresh
    Case "USR+"
      svr.Users.Refresh
    Case "AGG+"
      svr.Databases(objNode.Parent.Text).Aggregates.Refresh
    Case "FNC+"
      svr.Databases(objNode.Parent.Text).Functions.Refresh
    Case "LNG+"
      svr.Databases(objNode.Parent.Text).Languages.Refresh
    Case "OPR+"
      svr.Databases(objNode.Parent.Text).Operators.Refresh
    Case "SEQ+"
      svr.Databases(objNode.Parent.Text).Sequences.Refresh
    Case "TBL+"
      svr.Databases(objNode.Parent.Text).Tables.Refresh
    Case "CHK+"
      svr.Databases(objNode.Parent.Parent.Parent.Text).Tables(objNode.Parent.Text).Checks.Refresh
    Case "COL+"
      svr.Databases(objNode.Parent.Parent.Parent.Text).Tables(objNode.Parent.Text).Columns.Refresh
    Case "FKY+"
      svr.Databases(objNode.Parent.Parent.Parent.Text).Tables(objNode.Parent.Text).ForeignKeys.Refresh
    Case "REL+"
      svr.Databases(objNode.Parent.Parent.Parent.Parent.Parent.Text).Tables(objNode.Parent.Parent.Parent.Text).ForeignKeys(objNode.Parent.Text).Relationships.Refresh
    Case "IND+"
      svr.Databases(objNode.Parent.Parent.Parent.Text).Tables(objNode.Parent.Text).Indexes.Refresh
    Case "RUL+"
      svr.Databases(objNode.Parent.Parent.Parent.Text).Tables(objNode.Parent.Text).Rules.Refresh
    Case "TRG+"
      svr.Databases(objNode.Parent.Parent.Parent.Text).Tables(objNode.Parent.Text).Triggers.Refresh
    Case "TYP+"
      svr.Databases(objNode.Parent.Text).Types.Refresh
    Case "VIE+"
      svr.Databases(objNode.Parent.Text).Views.Refresh
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

Private Sub mnuPopupDrop_Click()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuPopupDrop_Click()", etFullDebug

  Drop
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuPopupDrop_Click"
End Sub

Private Sub mnuPopupProperties_Click()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuPopupProperties_Click()", etFullDebug

      Select Case ctx.CurrentObject.ObjectType
        Case "Aggregate"
          Dim objAggregateForm As New frmAggregate
          Load objAggregateForm
          objAggregateForm.Initialise ctx.CurrentDB, ctx.CurrentObject
          objAggregateForm.Show
          
        Case "Column"
          Dim objColumnForm As New frmColumn
          Load objColumnForm
          objColumnForm.Initialise ctx.CurrentDB, "MP", ctx.CurrentObject
          objColumnForm.Show
          
        Case "Database"
          Dim objDatabaseForm As New frmDatabase
          Load objDatabaseForm
          objDatabaseForm.Initialise ctx.CurrentObject
          objDatabaseForm.Show
          
        Case "Foreign Key"
          Dim objForeignKeyForm As New frmForeignKey
          Load objForeignKeyForm
          objForeignKeyForm.Initialise ctx.CurrentDB, "MP", ctx.CurrentObject
          objForeignKeyForm.Show
          
        Case "Function"
          Dim objFunctionForm As New frmFunction
          Load objFunctionForm
          objFunctionForm.Initialise ctx.CurrentDB, ctx.CurrentObject
          objFunctionForm.Show

        Case "Group"
          Dim objGroupForm As New frmGroup
          Load objGroupForm
          objGroupForm.Initialise ctx.CurrentObject
          objGroupForm.Show
    
        Case "Index"
          Dim objIndexForm As New frmIndex
          Load objIndexForm
          objIndexForm.Initialise ctx.CurrentDB, ctx.CurrentObject
          objIndexForm.Show
          
        Case "Language"
          Dim objLanguageForm As New frmLanguage
          Load objLanguageForm
          objLanguageForm.Initialise ctx.CurrentDB, ctx.CurrentObject
          objLanguageForm.Show
          
        Case "Operator"
          Dim objOperatorForm As New frmOperator
          Load objOperatorForm
          objOperatorForm.Initialise ctx.CurrentDB, ctx.CurrentObject
          objOperatorForm.Show
          
        Case "Rule"
          Dim objRuleForm As New frmRule
          Load objRuleForm
          objRuleForm.Initialise ctx.CurrentDB, ctx.CurrentObject
          objRuleForm.Show
          
        Case "Server"
          Dim objServerForm As New frmServer
          Load objServerForm
          objServerForm.Initialise ctx.CurrentObject
          objServerForm.Show
          
        Case "Sequence"
          Dim objSequenceForm As New frmSequence
          Load objSequenceForm
          objSequenceForm.Initialise ctx.CurrentDB, ctx.CurrentObject
          objSequenceForm.Show

        Case "Table"
          Dim objTableForm As New frmTable
          Load objTableForm
          objTableForm.Initialise ctx.CurrentDB, ctx.CurrentObject
          objTableForm.Show
          
        Case "Trigger"
          Dim objTriggerForm As New frmTrigger
          Load objTriggerForm
          objTriggerForm.Initialise ctx.CurrentDB, ctx.CurrentObject
          objTriggerForm.Show
          
        Case "Type"
          Dim objTypeForm As New frmType
          Load objTypeForm
          objTypeForm.Initialise ctx.CurrentDB, ctx.CurrentObject
          objTypeForm.Show
          
        Case "User"
          Dim objUserForm As New frmUser
          Load objUserForm
          objUserForm.Initialise ctx.CurrentObject
          objUserForm.Show
          
        Case "View"
          Dim objViewForm As New frmView
          Load objViewForm
          objViewForm.Initialise ctx.CurrentDB, ctx.CurrentObject
          objViewForm.Show
          
        Case Else
          MsgBox "Unknown object type for the current object.", vbExclamation, "Error"
      End Select
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuPopupProperties_Click"
End Sub

Private Sub mnuPopupSQL_Click()
On Error GoTo Err_Handler
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
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuPopupViewData_Click()", etFullDebug
  
Dim objOutputForm As New frmSQLOutput
Dim szQuery As String
Dim rsQuery As New Recordset

  StartMsg "Counting Records..."
  Set rsQuery = frmMain.svr.Databases(ctx.CurrentDB).Execute("SELECT count(*) AS count FROM " & QUOTE & ctx.CurrentObject.Identifier & QUOTE)
  EndMsg
  If Not rsQuery.EOF Then
    If rsQuery!Count > 5000 Then If MsgBox("This " & ctx.CurrentObject.ObjectType & " contains " & rsQuery!Count & " records which may take some time to load." & vbCrLf & "Do you wish to continue?", vbQuestion + vbYesNo, "Continue?") = vbNo Then Exit Sub
  End If
  
  StartMsg "Executing SQL Query..."

  Set rsQuery = frmMain.svr.Databases(ctx.CurrentDB).Execute("SELECT * FROM " & QUOTE & ctx.CurrentObject.Identifier & QUOTE)
  Load objOutputForm
  objOutputForm.Display rsQuery, ctx.CurrentDB, "(" & ctx.CurrentObject.ObjectType & ": " & ctx.CurrentObject.Identifier & ")"
  objOutputForm.Show
  
  If rsQuery.State <> adStateClosed Then rsQuery.Close
  Set rsQuery = Nothing
  EndMsg
  Exit Sub
  
Err_Handler:
  If rsQuery.State <> adStateClosed Then rsQuery.Close
  Set rsQuery = Nothing
  EndMsg
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuPopupViewData_Click"
End Sub

Private Sub tb_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Err_Handler
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
    Case "vacuum"
      Vacuum False
    Case Else
      MsgBox "Unknown menu button pressed.", vbExclamation, "Error"
  End Select
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tb_ButtonClick"
End Sub

Private Sub mnuPopupVacuumVacuum_Click()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuPopupVacuumVacuum_Click()", etFullDebug

  Vacuum False
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuPopupVacuumVacuum_Click"
End Sub

Private Sub mnuPopupVacuumAnalyse_Click()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuPopupVacuumAnalyse_Click()", etFullDebug

  Vacuum True
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuPopupVacuumAnalyse_Click"
End Sub

Private Sub mnuPopupCreateAggregate_Click()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuPopupCreateAggregate_Click()", etFullDebug

Dim objAggregateForm As New frmAggregate

  Load objAggregateForm
  objAggregateForm.Initialise ctx.CurrentDB
  objAggregateForm.Show
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuPopupCreateAggregate_Click"
End Sub

Private Sub mnuPopupCreateDatabase_Click()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuPopupCreateDatabase_Click()", etFullDebug

Dim objDatabaseForm As New frmDatabase

  Load objDatabaseForm
  objDatabaseForm.Initialise
  objDatabaseForm.Show
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuPopupCreateDatabase_Click"
End Sub

Private Sub mnuPopupCreateFunction_Click()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuPopupCreateFunction_Click()", etFullDebug

Dim objFunctionForm As New frmFunction

  Load objFunctionForm
  objFunctionForm.Initialise ctx.CurrentDB
  objFunctionForm.Show
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuPopupCreateFunction_Click"
End Sub

Private Sub mnuPopupCreateGroup_Click()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuPopupCreateGroup_Click()", etFullDebug

Dim objGroupForm As New frmGroup

  Load objGroupForm
  objGroupForm.Initialise
  objGroupForm.Show
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuPopupCreateGroup_Click"
End Sub

Private Sub mnuPopupCreateIndex_Click()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuPopupCreateIndex_Click()", etFullDebug

Dim objIndexForm As New frmIndex

  Load objIndexForm
  objIndexForm.Initialise ctx.CurrentDB
  objIndexForm.Show
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuPopupCreateIndex_Click"
End Sub

Private Sub mnuPopupCreateLanguage_Click()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuPopupCreateLanguage_Click()", etFullDebug

Dim objLanguageForm As New frmLanguage

  Load objLanguageForm
  objLanguageForm.Initialise ctx.CurrentDB
  objLanguageForm.Show
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuPopupCreateLanguage_Click"
End Sub

Private Sub mnuPopupCreateOperator_Click()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuPopupCreateOperator_Click()", etFullDebug

Dim objOperatorForm As New frmOperator

  Load objOperatorForm
  objOperatorForm.Initialise ctx.CurrentDB
  objOperatorForm.Show
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuPopupCreateOperator_Click"
End Sub

Private Sub mnuPopupCreateRule_Click()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuPopupCreateRule_Click()", etFullDebug

Dim objRuleForm As New frmRule

  Load objRuleForm
  objRuleForm.Initialise ctx.CurrentDB
  objRuleForm.Show
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuPopupCreateRule_Click"
End Sub

Private Sub mnuPopupCreateSequence_Click()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuPopupCreateSequence_Click()", etFullDebug

Dim objSequenceForm As New frmSequence

  Load objSequenceForm
  objSequenceForm.Initialise ctx.CurrentDB
  objSequenceForm.Show
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuPopupCreateSequence_Click"
End Sub

Private Sub mnuPopupCreateTable_Click()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuPopupCreateTable_Click()", etFullDebug

Dim objTableForm As New frmTable

  Load objTableForm
  objTableForm.Initialise ctx.CurrentDB
  objTableForm.Show
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuPopupCreateTable_Click"
End Sub

Private Sub mnuPopupCreateTrigger_Click()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuPopupCreateTrigger_Click()", etFullDebug

Dim objTriggerForm As New frmTrigger

  Load objTriggerForm
  objTriggerForm.Initialise ctx.CurrentDB
  objTriggerForm.Show
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuPopupCreateTrigger_Click"
End Sub

Private Sub mnuPopupCreateType_Click()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuPopupCreateType_Click()", etFullDebug

Dim objTypeForm As New frmType

  Load objTypeForm
  objTypeForm.Initialise ctx.CurrentDB
  objTypeForm.Show
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuPopupCreateType_Click"
End Sub

Private Sub mnuPopupCreateUser_Click()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuPopupCreateUser_Click()", etFullDebug

Dim objUserForm As New frmUser

  Load objUserForm
  objUserForm.Initialise
  objUserForm.Show
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuPopupCreateUser_Click"
End Sub

Private Sub mnuPopupCreateView_Click()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuPopupCreateView_Click()", etFullDebug

Dim objViewForm As New frmView

  Load objViewForm
  objViewForm.Initialise ctx.CurrentDB
  objViewForm.Show
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuPopupCreateView_Click"
End Sub

Private Sub tb_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tb_ButtonMenuClick(" & ButtonMenu & ")", etFullDebug

  Select Case ButtonMenu.Parent.Key
    Case "connect"
      Load frmConnect
      frmConnect.Load_Defaults Val(Mid(ButtonMenu.Key, 1, InStr(1, ButtonMenu.Key, "|") - 1))
      frmConnect.Show vbModal, Me
    
    Case "create"
    
      'For each of these just call the popup menu function
      Select Case ButtonMenu.Key
        Case "aggregate"
          mnuPopupCreateAggregate_Click
        Case "database"
          mnuPopupCreateDatabase_Click
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

Private Sub tv_DragDrop(Source As Control, X As Single, Y As Single)
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tv_DragDrop(" & QUOTE & Source.Name & QUOTE & ", " & X & ", " & Y & ")", etFullDebug

  If Source = imgSplitter Then
    SizeControls X
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tv_DragDrop"
End Sub

Public Sub SizeControls(X As Single)
On Error Resume Next
svr.LogEvent "Entering " & App.Title & ":frmMain.SizeControls(" & X & ")", etFullDebug

  'Set the width
  If X < 1500 Then X = 1500
  If X > (Me.Width - 1500) Then X = Me.Width - 1500
  tv.Width = X
  imgSplitter.Left = X
  lv.Left = X + 40
  lv.Width = Me.Width - (tv.Width + 140)
  txtDefinition.Left = lv.Left
  txtDefinition.Width = lv.Width

  'Set the top
  If tb.Visible Then
    tv.Top = tb.Height
  Else
    tv.Top = 0
  End If
  lv.Top = tv.Top
  
  'Set the height
  If sb.Visible Then
    tv.Height = Me.ScaleHeight - (tv.Top - sb.Height) - 575
  Else
    tv.Height = Me.ScaleHeight - tv.Top
  End If
  
  If txtDefinition.Visible Then
    lv.Height = tv.Height - txtDefinition.Height
  Else
    lv.Height = tv.Height
  End If
  txtDefinition.Top = lv.Height + lv.Top
  imgSplitter.Top = tv.Top
  imgSplitter.Height = tv.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
svr.LogEvent "Entering " & App.Title & ":frmMain.Form_Unload(" & Cancel & ")", etFullDebug

Dim objform As Form
  
  'Close child forms.
  For Each objform In Forms
    Unload objform
  Next objform
  
  'Save the Window size/position
  RegWrite HKEY_CURRENT_USER, "Software\" & App.Title, "Top", regString, Me.Top
  RegWrite HKEY_CURRENT_USER, "Software\" & App.Title, "Left", regString, Me.Left
  RegWrite HKEY_CURRENT_USER, "Software\" & App.Title, "Width", regString, Me.Width
  RegWrite HKEY_CURRENT_USER, "Software\" & App.Title, "Height", regString, Me.Height
  
  'Clear the Server, then Context objects last as the forms may be using them for logging
  Set svr = Nothing
  Set ctx = Nothing
End Sub

Private Sub mnuFileExit_Click()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuFileExit_Click()", etFullDebug

  Unload Me
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuFileExit_Click"
End Sub

Private Sub mnuFileConnect_Click()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.mnuFileConnect_Click()", etFullDebug

  Load frmConnect
  frmConnect.Load_Defaults
  frmConnect.Show vbModal, Me
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.mnuFileConnect_Click"
End Sub

Private Sub svr_EventLog(EventLevel As pgSchema.LogLevel, EventMessage As String)
'Note - No function entry logging is done here 'cos we'd enter a loop then...

  If ctx.LogView Then
    If EventLevel <= ctx.LogLevel Then frmLog.LogMsg EventMessage
  End If

End Sub

Private Sub tvServer(ByVal Node As MSComctlLib.Node)
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvServer(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
  If Node.Children = 0 Then
    tv.Nodes.Add Node.Key, tvwChild, "DAT+" & GetID, "Databases", "database"
    tv.Nodes.Add Node.Key, tvwChild, "GRP+" & GetID, "Groups", "group"
    tv.Nodes.Add Node.Key, tvwChild, "USR+" & GetID, "Users", "user"
  End If
  lv.ColumnHeaders.Add , , "Property", 2000
  lv.ColumnHeaders.Add , , "Value", lv.Width - 2100
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Hostname", "property", "property")
  lvItem.SubItems(1) = svr.Server & ""
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Port", "property", "property")
  lvItem.SubItems(1) = svr.Port
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Username", "property", "property")
  lvItem.SubItems(1) = svr.Username
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "DBMS", "property", "property")
  lvItem.SubItems(1) = svr.dbVersion.Description
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvServer"
End Sub

Private Sub tvDatabases(ByVal Node As MSComctlLib.Node)
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvDatabases(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
Dim dat As pgDatabase

  If Node.Children = 0 Or Node.Children <> svr.Databases.Count(Not ctx.IncludeSys) Then
    While Not (Node.Child Is Nothing)
      tv.Nodes.Remove Node.Child.Index
    Wend
    For Each dat In svr.Databases
      If Not (dat.SystemObject And Not ctx.IncludeSys) Then
        If dat.Status <> statInaccessible Then
          tv.Nodes.Add Node.Key, tvwChild, "DAT-" & GetID, dat.Identifier, "database"
        Else
          tv.Nodes.Add Node.Key, tvwChild, "DAT-" & GetID, dat.Identifier, "baddatabase"
        End If
      End If
    Next dat
    Node.Text = "Databases (" & Node.Children & ")"
  End If
  lv.ColumnHeaders.Add , , "Database", 2000
  lv.ColumnHeaders.Add , , "Comment", lv.Width - 2100
  For Each dat In svr.Databases
    If Not (dat.SystemObject And Not ctx.IncludeSys) Then
      Set lvItem = lv.ListItems.Add(, "DAT-" & GetID, dat.Identifier, "database", "database")
      lvItem.SubItems(1) = Replace(dat.Comment, vbCrLf, " ")
    End If
  Next dat
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvDatabases"
End Sub

Private Sub tvDatabase(ByVal Node As MSComctlLib.Node)
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvDatabase(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem

  If svr.Databases(Node.Text).Status <> statInaccessible Then
    If Node.Children = 0 Then
      tv.Nodes.Add Node.Key, tvwChild, "AGG+" & GetID, "Aggregates", "aggregate"
      tv.Nodes.Add Node.Key, tvwChild, "FNC+" & GetID, "Functions", "function"
      tv.Nodes.Add Node.Key, tvwChild, "LNG+" & GetID, "Languages", "language"
      tv.Nodes.Add Node.Key, tvwChild, "OPR+" & GetID, "Operators", "operator"
      tv.Nodes.Add Node.Key, tvwChild, "SEQ+" & GetID, "Sequences", "sequence"
      tv.Nodes.Add Node.Key, tvwChild, "TBL+" & GetID, "Tables", "table"
      tv.Nodes.Add Node.Key, tvwChild, "TYP+" & GetID, "Types", "type"
      tv.Nodes.Add Node.Key, tvwChild, "VIE+" & GetID, "Views", "view"
    End If
  Else
    Node.Image = "baddatabase"
  End If
  lv.ColumnHeaders.Add , , "Property", 2000
  lv.ColumnHeaders.Add , , "Value", lv.Width - 2100
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Name", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Text).Name
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "OID", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Text).OID
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Owner", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Text).Owner
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Path", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Text).Path
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Encoding", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Text).EncodingName
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Accessible?", "property", "property")
  If svr.Databases(Node.Text).Status <> statInaccessible Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "System Database?", "property", "property")
  If svr.Databases(Node.Text).SystemObject Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Comment", "property", "property")
  lvItem.SubItems(1) = Replace(svr.Databases(Node.Text).Comment, vbCrLf, " ")
  
  'Set the Definition Pane
  If txtDefinition.Visible Then txtDefinition.Text = svr.Databases(Node.Text).SQL
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvDatabase"
End Sub

Private Sub tvGroups(ByVal Node As MSComctlLib.Node)
On Error GoTo Err_Handler
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
      tv.Nodes.Add Node.Key, tvwChild, "GRP-" & GetID, grp.Identifier, "group"
    Next grp
    Node.Text = "Groups (" & Node.Children & ")"
  End If
  lv.ColumnHeaders.Add , , "Group", 2000
  lv.ColumnHeaders.Add , , "Group ID", 1000
  lv.ColumnHeaders.Add , , "Members", lv.Width - 3100
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
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvGroup(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
Dim szTemp As String
Dim vData As Variant

  lv.ColumnHeaders.Add , , "Property", 2000
  lv.ColumnHeaders.Add , , "Value", lv.Width - 2100
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Name", "property", "property")
  lvItem.SubItems(1) = svr.Groups(Node.Text).Name
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Group ID", "property", "property")
  lvItem.SubItems(1) = svr.Groups(Node.Text).ID
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Member Count", "property", "property")
  lvItem.SubItems(1) = svr.Groups(Node.Text).Members.Count
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Members", "property", "property")
  For Each vData In svr.Groups(Node.Text).Members
    szTemp = szTemp & vData & ", "
  Next vData
  If Len(szTemp) > 2 Then lvItem.SubItems(1) = Left(szTemp, Len(szTemp) - 2)
  
  'Set the Definition Pane
  If txtDefinition.Visible Then txtDefinition.Text = svr.Groups(Node.Text).SQL
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvGroup"
End Sub

Private Sub tvUsers(ByVal Node As MSComctlLib.Node)
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvUsers(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
Dim usr As pgUser

  If Node.Children = 0 Or Node.Children <> svr.Users.Count Then
    While Not (Node.Child Is Nothing)
      tv.Nodes.Remove Node.Child.Index
    Wend
    For Each usr In svr.Users
      tv.Nodes.Add Node.Key, tvwChild, "USR-" & GetID, usr.Identifier, "user"
    Next usr
    Node.Text = "Users (" & Node.Children & ")"
  End If
  lv.ColumnHeaders.Add , , "Username", 2000
  lv.ColumnHeaders.Add , , "User ID", 1500
  lv.ColumnHeaders.Add , , "Account Expires", lv.Width - 3600
  For Each usr In svr.Users
    Set lvItem = lv.ListItems.Add(, "USR-" & GetID, usr.Identifier, "user", "user")
    lvItem.SubItems(1) = usr.ID
    lvItem.SubItems(2) = usr.AccountExpires
  Next usr
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvUsers"
End Sub

Private Sub tvUser(ByVal Node As MSComctlLib.Node)
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvUser(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
Dim szTemp As String
Dim vData As Variant

  lv.ColumnHeaders.Add , , "Property", 2000
  lv.ColumnHeaders.Add , , "Value", lv.Width - 2100
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Name", "property", "property")
  lvItem.SubItems(1) = svr.Users(Node.Text).Name
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "User ID", "property", "property")
  lvItem.SubItems(1) = svr.Users(Node.Text).ID
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Account Expires", "property", "property")
  lvItem.SubItems(1) = svr.Users(Node.Text).AccountExpires
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Create Databases?", "property", "property")
  If svr.Users(Node.Text).CreateDatabases Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Superuser?", "property", "property")
  If svr.Users(Node.Text).Superuser Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Update Catalogues", "property", "property")
  If svr.Users(Node.Text).UpdateCatalogues Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  
  'Set the Definition Pane
  If txtDefinition.Visible Then txtDefinition.Text = svr.Users(Node.Text).SQL
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvUser"
End Sub

Private Sub tvAggregates(ByVal Node As MSComctlLib.Node)
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvAggregates(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
Dim agg As pgAggregate

  If Node.Children = 0 Or Node.Children <> svr.Databases(Node.Parent.Text).Aggregates.Count(Not ctx.IncludeSys) Then
    While Not (Node.Child Is Nothing)
      tv.Nodes.Remove Node.Child.Index
    Wend
    For Each agg In svr.Databases(Node.Parent.Text).Aggregates
      If Not (agg.SystemObject And Not ctx.IncludeSys) Then tv.Nodes.Add Node.Key, tvwChild, "AGG-" & GetID, agg.Identifier, "aggregate"
    Next agg
    Node.Text = "Aggregates (" & Node.Children & ")"
  End If
  lv.ColumnHeaders.Add , , "Aggregate", 2000
  lv.ColumnHeaders.Add , , "Comment", lv.Width - 2100
  For Each agg In svr.Databases(Node.Parent.Text).Aggregates
    If Not (agg.SystemObject And Not ctx.IncludeSys) Then
      Set lvItem = lv.ListItems.Add(, "AGG-" & GetID, agg.Identifier, "aggregate", "aggregate")
      lvItem.SubItems(1) = Replace(agg.Comment, vbCrLf, " ")
    End If
  Next agg
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvAggregates"
End Sub

Private Sub tvAggregate(ByVal Node As MSComctlLib.Node)
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvAggregate(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem

  lv.ColumnHeaders.Add , , "Property", 2000
  lv.ColumnHeaders.Add , , "Value", lv.Width - 2100
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Name", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Aggregates(Node.Text).Name
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Input Type", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Aggregates(Node.Text).InputType
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "OID", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Aggregates(Node.Text).OID
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Owner", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Aggregates(Node.Text).Owner
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "State Type", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Aggregates(Node.Text).StateType
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "State Function", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Aggregates(Node.Text).StateFunction
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Final Type", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Aggregates(Node.Text).FinalType
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Final Function", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Aggregates(Node.Text).FinalFunction
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Initial Condition", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Aggregates(Node.Text).InitialCondition
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "System Aggregate?", "property", "property")
  If svr.Databases(Node.Parent.Parent.Text).Aggregates(Node.Text).SystemObject Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Comment", "property", "property")
  lvItem.SubItems(1) = Replace(svr.Databases(Node.Parent.Parent.Text).Aggregates(Node.Text).Comment, vbCrLf, " ")
  
  'Set the Definition Pane
  If txtDefinition.Visible Then txtDefinition.Text = svr.Databases(Node.Parent.Parent.Text).Aggregates(Node.Text).SQL
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvAggregate"
End Sub

Private Sub tvFunctions(ByVal Node As MSComctlLib.Node)
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvFunctions(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
Dim szTemp As String
Dim vData As Variant
Dim fnc As pgFunction

  If Node.Children = 0 Or Node.Children <> svr.Databases(Node.Parent.Text).Functions.Count(Not ctx.IncludeSys) Then
    While Not (Node.Child Is Nothing)
      tv.Nodes.Remove Node.Child.Index
    Wend
    For Each fnc In svr.Databases(Node.Parent.Text).Functions
      If Not (fnc.SystemObject And Not ctx.IncludeSys) Then tv.Nodes.Add Node.Key, tvwChild, "FNC-" & GetID, fnc.Identifier, "function"
    Next fnc
    Node.Text = "Functions (" & Node.Children & ")"
  End If
  lv.ColumnHeaders.Add , , "Function", 2000
  lv.ColumnHeaders.Add , , "Comment", lv.Width - 2100
  For Each fnc In svr.Databases(Node.Parent.Text).Functions
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
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvFunction(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
Dim szTemp As String
Dim vData As Variant

  lv.ColumnHeaders.Add , , "Property", 2000
  lv.ColumnHeaders.Add , , "Value", lv.Width - 2100
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Name", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Functions(Node.Text).Name
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Argument Count", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Functions(Node.Text).Arguments.Count
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Arguments", "property", "property")
  szTemp = ""
  For Each vData In svr.Databases(Node.Parent.Parent.Text).Functions(Node.Text).Arguments
    szTemp = szTemp & vData & ", "
  Next vData
  If Len(szTemp) > 2 Then szTemp = Left(szTemp, Len(szTemp) - 2)
  lvItem.SubItems(1) = szTemp
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Returns", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Functions(Node.Text).Returns
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "OID", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Functions(Node.Text).OID
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Owner", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Functions(Node.Text).Owner
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Language", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Functions(Node.Text).Language
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Source", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Functions(Node.Text).Source
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Cachable?", "property", "property")
  If svr.Databases(Node.Parent.Parent.Text).Functions(Node.Text).Cachable Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Strict?", "property", "property")
  If svr.Databases(Node.Parent.Parent.Text).Functions(Node.Text).Strict Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "System Function?", "property", "property")
  If svr.Databases(Node.Parent.Parent.Text).Functions(Node.Text).SystemObject Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Comment", "property", "property")
  lvItem.SubItems(1) = Replace(svr.Databases(Node.Parent.Parent.Text).Functions(Node.Text).Comment, vbCrLf, " ")
  
  'Set the Definition Pane
  If txtDefinition.Visible Then txtDefinition.Text = svr.Databases(Node.Parent.Parent.Text).Functions(Node.Text).SQL
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvFunction"
End Sub

Private Sub tvLanguages(ByVal Node As MSComctlLib.Node)
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvLanguages(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
Dim lng As pgLanguage

  If Node.Children = 0 Or Node.Children <> svr.Databases(Node.Parent.Text).Languages.Count(Not ctx.IncludeSys) Then
    While Not (Node.Child Is Nothing)
      tv.Nodes.Remove Node.Child.Index
    Wend
    For Each lng In svr.Databases(Node.Parent.Text).Languages
      If Not (lng.SystemObject And Not ctx.IncludeSys) Then tv.Nodes.Add Node.Key, tvwChild, "LNG-" & GetID, lng.Identifier, "language"
    Next lng
    Node.Text = "Languages (" & Node.Children & ")"
  End If
  lv.ColumnHeaders.Add , , "Language", 2000
  lv.ColumnHeaders.Add , , "Comment", lv.Width - 2100
  For Each lng In svr.Databases(Node.Parent.Text).Languages
    If Not (lng.SystemObject And Not ctx.IncludeSys) Then
      Set lvItem = lv.ListItems.Add(, "LNG-" & GetID, lng.Identifier, "language", "language")
      lvItem.SubItems(1) = Replace(lng.Comment, vbCrLf, " ")
    End If
  Next lng
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvLanguages"
End Sub

Private Sub tvLanguage(ByVal Node As MSComctlLib.Node)
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvLanguage(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem

  lv.ColumnHeaders.Add , , "Property", 2000
  lv.ColumnHeaders.Add , , "Value", lv.Width - 2100
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Name", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Languages(Node.Text).Name
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "OID", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Languages(Node.Text).OID
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Handler", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Languages(Node.Text).Handler
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Trusted?", "property", "property")
  If svr.Databases(Node.Parent.Parent.Text).Languages(Node.Text).Trusted Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "System Language?", "property", "property")
  If svr.Databases(Node.Parent.Parent.Text).Languages(Node.Text).SystemObject Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Comment", "property", "property")
  lvItem.SubItems(1) = Replace(svr.Databases(Node.Parent.Parent.Text).Languages(Node.Text).Comment, vbCrLf, " ")
  
  'Set the Definition Pane
  If txtDefinition.Visible Then txtDefinition.Text = svr.Databases(Node.Parent.Parent.Text).Languages(Node.Text).SQL
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvLanguage"
End Sub

Private Sub tvOperators(ByVal Node As MSComctlLib.Node)
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvOperators(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
Dim opr As pgOperator

  If Node.Children = 0 Or Node.Children <> svr.Databases(Node.Parent.Text).Operators.Count(Not ctx.IncludeSys) Then
    While Not (Node.Child Is Nothing)
      tv.Nodes.Remove Node.Child.Index
    Wend
    For Each opr In svr.Databases(Node.Parent.Text).Operators
      If Not (opr.SystemObject And Not ctx.IncludeSys) Then tv.Nodes.Add Node.Key, tvwChild, "OPR-" & GetID, opr.Identifier, "operator"
    Next opr
    Node.Text = "Operators (" & Node.Children & ")"
  End If
  lv.ColumnHeaders.Add , , "Operator", 2000
  lv.ColumnHeaders.Add , , "Comment", lv.Width - 2100
  For Each opr In svr.Databases(Node.Parent.Text).Operators
    If Not (opr.SystemObject And Not ctx.IncludeSys) Then
      Set lvItem = lv.ListItems.Add(, "OPR-" & GetID, opr.Identifier, "operator", "operator")
      lvItem.SubItems(1) = Replace(opr.Comment, vbCrLf, " ")
    End If
  Next opr
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvOperators"
End Sub

Private Sub tvOperator(ByVal Node As MSComctlLib.Node)
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvOperator(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem

  lv.ColumnHeaders.Add , , "Property", 2000
  lv.ColumnHeaders.Add , , "Value", lv.Width - 2100
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Name", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Operators(Node.Text).Name
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Left Type", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Operators(Node.Text).LeftOperandType
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Right Type", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Operators(Node.Text).RightOperandType
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "OID", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Operators(Node.Text).OID
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Owner", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Operators(Node.Text).Owner
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Operator Function", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Operators(Node.Text).OperatorFunction
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Join Function", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Operators(Node.Text).JoinFunction
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Restrict Function", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Operators(Node.Text).RestrictFunction
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Result Type", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Operators(Node.Text).ResultType
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Commutator", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Operators(Node.Text).Commutator
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Negator", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Operators(Node.Text).Negator
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Kind", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Operators(Node.Text).Kind
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Left Sort Operator", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Operators(Node.Text).LeftTypeSortOperator
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Right Sort Operator", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Operators(Node.Text).RightTypeSortOperator
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Hash Joins?", "property", "property")
  If svr.Databases(Node.Parent.Parent.Text).Operators(Node.Text).HashJoins Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "System Operator?", "property", "property")
  If svr.Databases(Node.Parent.Parent.Text).Operators(Node.Text).SystemObject Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Comment", "property", "property")
  lvItem.SubItems(1) = Replace(svr.Databases(Node.Parent.Parent.Text).Operators(Node.Text).Comment, vbCrLf, " ")
  
  'Set the Definition Pane
  If txtDefinition.Visible Then txtDefinition.Text = svr.Databases(Node.Parent.Parent.Text).Operators(Node.Text).SQL
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvOperator"
End Sub

Private Sub tvSequences(ByVal Node As MSComctlLib.Node)
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvSequences(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
Dim seq As pgSequence

  If Node.Children = 0 Or Node.Children <> svr.Databases(Node.Parent.Text).Sequences.Count(Not ctx.IncludeSys) Then
    While Not (Node.Child Is Nothing)
      tv.Nodes.Remove Node.Child.Index
    Wend
    For Each seq In svr.Databases(Node.Parent.Text).Sequences
      If Not (seq.SystemObject And Not ctx.IncludeSys) Then tv.Nodes.Add Node.Key, tvwChild, "SEQ-" & GetID, seq.Identifier, "sequence"
    Next seq
    Node.Text = "Sequences (" & Node.Children & ")"
  End If
  lv.ColumnHeaders.Add , , "Sequence", 2000
  lv.ColumnHeaders.Add , , "Comment", lv.Width - 2100
  For Each seq In svr.Databases(Node.Parent.Text).Sequences
    If Not (seq.SystemObject And Not ctx.IncludeSys) Then
      Set lvItem = lv.ListItems.Add(, "SEQ-" & GetID, seq.Identifier, "sequence", "sequence")
      lvItem.SubItems(1) = Replace(seq.Comment, vbCrLf, " ")
    End If
  Next seq
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvSequences"
End Sub

Private Sub tvSequence(ByVal Node As MSComctlLib.Node)
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvSequence(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem

  lv.ColumnHeaders.Add , , "Property", 2000
  lv.ColumnHeaders.Add , , "Value", lv.Width - 2100
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Name", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Sequences(Node.Text).Name
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "OID", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Sequences(Node.Text).OID
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Owner", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Sequences(Node.Text).Owner
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "ACL", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Sequences(Node.Text).ACL
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Last Value", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Sequences(Node.Text).LastValue
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Minimum", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Sequences(Node.Text).Minimum
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Maximum", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Sequences(Node.Text).Maximum
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Increment", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Sequences(Node.Text).Increment
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Cache", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Sequences(Node.Text).Cache
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Cycled?", "property", "property")
  If svr.Databases(Node.Parent.Parent.Text).Sequences(Node.Text).Cycled Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "System Sequence?", "property", "property")
  If svr.Databases(Node.Parent.Parent.Text).Sequences(Node.Text).SystemObject Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Comment", "property", "property")
  lvItem.SubItems(1) = Replace(svr.Databases(Node.Parent.Parent.Text).Sequences(Node.Text).Comment, vbCrLf, " ")
  
  'Set the Definition Pane
  If txtDefinition.Visible Then txtDefinition.Text = svr.Databases(Node.Parent.Parent.Text).Sequences(Node.Text).SQL
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvSequence"
End Sub

Private Sub tvTables(ByVal Node As MSComctlLib.Node)
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvTables(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
Dim tbl As pgTable

  If Node.Children = 0 Or Node.Children <> svr.Databases(Node.Parent.Text).Tables.Count(Not ctx.IncludeSys) Then
    While Not (Node.Child Is Nothing)
      tv.Nodes.Remove Node.Child.Index
    Wend
    For Each tbl In svr.Databases(Node.Parent.Text).Tables
      If Not (tbl.SystemObject And Not ctx.IncludeSys) Then tv.Nodes.Add Node.Key, tvwChild, "TBL-" & GetID, tbl.Identifier, "table"
    Next tbl
    Node.Text = "Tables (" & Node.Children & ")"
  End If
  lv.ColumnHeaders.Add , , "Table", 2000
  lv.ColumnHeaders.Add , , "Comment", lv.Width - 2100
  For Each tbl In svr.Databases(Node.Parent.Text).Tables
    If Not (tbl.SystemObject And Not ctx.IncludeSys) Then
      Set lvItem = lv.ListItems.Add(, "TBL-" & GetID, tbl.Identifier, "table", "table")
      lvItem.SubItems(1) = Replace(tbl.Comment, vbCrLf, " ")
    End If
  Next tbl
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvTables"
End Sub

Private Sub tvTable(ByVal Node As MSComctlLib.Node)
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvTable(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
Dim szTemp As String
Dim vData As Variant

  lv.ColumnHeaders.Add , , "Property", 2000
  lv.ColumnHeaders.Add , , "Value", lv.Width - 2100
  If Node.Children = 0 Then
    tv.Nodes.Add Node.Key, tvwChild, "CHK+" & GetID, "Checks", "check"
    tv.Nodes.Add Node.Key, tvwChild, "COL+" & GetID, "Columns", "column"
    tv.Nodes.Add Node.Key, tvwChild, "FKY+" & GetID, "Foreign Keys", "foreignkey"
    tv.Nodes.Add Node.Key, tvwChild, "IND+" & GetID, "Indexes", "index"
    tv.Nodes.Add Node.Key, tvwChild, "RUL+" & GetID, "Rules", "rule"
    tv.Nodes.Add Node.Key, tvwChild, "TRG+" & GetID, "Triggers", "trigger"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Name", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Tables(Node.Text).Name
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "OID", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Tables(Node.Text).OID
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Owner", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Tables(Node.Text).Owner
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "ACL", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Tables(Node.Text).ACL
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Rows", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Tables(Node.Text).Rows
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Inherited Tables Count", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Tables(Node.Text).InheritedTables.Count
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Inherited Tables", "property", "property")
  For Each vData In svr.Databases(Node.Parent.Parent.Text).Tables(Node.Text).InheritedTables
    szTemp = szTemp & vData & ", "
  Next vData
  If Len(szTemp) > 2 Then szTemp = Left(szTemp, Len(szTemp) - 2)
  lvItem.SubItems(1) = szTemp
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "OIDs?", "property", "property")
  If svr.Databases(Node.Parent.Parent.Text).Tables(Node.Text).HasOIDs Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "System Table?", "property", "property")
  If svr.Databases(Node.Parent.Parent.Text).Tables(Node.Text).SystemObject Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Comment", "property", "property")
  lvItem.SubItems(1) = Replace(svr.Databases(Node.Parent.Parent.Text).Tables(Node.Text).Comment, vbCrLf, " ")
  
  'Set the Definition Pane
  If txtDefinition.Visible Then txtDefinition.Text = svr.Databases(Node.Parent.Parent.Text).Tables(Node.Text).SQL
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvTable"
End Sub

Private Sub tvChecks(ByVal Node As MSComctlLib.Node)
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvChecks(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
Dim chk As pgCheck

  If Node.Children = 0 Or Node.Children <> svr.Databases(Node.Parent.Parent.Parent.Text).Tables(Node.Parent.Text).Checks.Count Then
    While Not (Node.Child Is Nothing)
      tv.Nodes.Remove Node.Child.Index
    Wend
    For Each chk In svr.Databases(Node.Parent.Parent.Parent.Text).Tables(Node.Parent.Text).Checks
      tv.Nodes.Add Node.Key, tvwChild, "CHK-" & GetID, chk.Identifier, "check"
    Next chk
    Node.Text = "Checks (" & Node.Children & ")"
  End If
  lv.ColumnHeaders.Add , , "Check", lv.Width
  For Each chk In svr.Databases(Node.Parent.Parent.Parent.Text).Tables(Node.Parent.Text).Checks
    Set lvItem = lv.ListItems.Add(, "CHK-" & GetID, chk.Identifier, "check", "check")
  Next chk
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvChecks"
End Sub

Private Sub tvCheck(ByVal Node As MSComctlLib.Node)
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvCheck(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem

  lv.ColumnHeaders.Add , , "Property", 2000
  lv.ColumnHeaders.Add , , "Value", lv.Width - 2100
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Name", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Checks(Node.Text).Name
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Definition", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Checks(Node.Text).Definition
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvCheck"
End Sub

Private Sub tvColumns(ByVal Node As MSComctlLib.Node)
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvColumns(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
Dim col As pgColumn

  If Node.Children = 0 Or Node.Children <> svr.Databases(Node.Parent.Parent.Parent.Text).Tables(Node.Parent.Text).Columns.Count(Not ctx.IncludeSys) Then
    While Not (Node.Child Is Nothing)
      tv.Nodes.Remove Node.Child.Index
    Wend
    For Each col In svr.Databases(Node.Parent.Parent.Parent.Text).Tables(Node.Parent.Text).Columns
     If Not (col.SystemObject And Not ctx.IncludeSys) Then tv.Nodes.Add Node.Key, tvwChild, "COL-" & GetID, col.Identifier, "column"
    Next col
    Node.Text = "Columns (" & Node.Children & ")"
  End If
  lv.ColumnHeaders.Add , , "Column", 2000
  lv.ColumnHeaders.Add , , "Type", 1000
  lv.ColumnHeaders.Add , , "Comment", lv.Width - 3100
  For Each col In svr.Databases(Node.Parent.Parent.Parent.Text).Tables(Node.Parent.Text).Columns
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
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvColumn(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem

  lv.ColumnHeaders.Add , , "Property", 2000
  lv.ColumnHeaders.Add , , "Value", lv.Width - 2100
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Name", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Columns(Node.Text).Name
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Position", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Columns(Node.Text).Position
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Data Type", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Columns(Node.Text).DataType
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Size", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Columns(Node.Text).Length
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Numeric Precision", "property", "property")
  If svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Columns(Node.Text).DataType = "numeric" Then
    lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Columns(Node.Text).NumericScale
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Default", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Columns(Node.Text).Default
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Restrict Nulls?", "property", "property")
  If svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Columns(Node.Text).NotNull Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Primary Key?", "property", "property")
  If svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Columns(Node.Text).PrimaryKey Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "System Column?", "property", "property")
  If svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Columns(Node.Text).SystemObject Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Comment", "property", "property")
  lvItem.SubItems(1) = Replace(svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Columns(Node.Text).Comment, vbCrLf, " ")
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvColumn"
End Sub

Private Sub tvForeignKeys(ByVal Node As MSComctlLib.Node)
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvForeignKeys(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
Dim fky As pgForeignKey

  If Node.Children = 0 Or Node.Children <> svr.Databases(Node.Parent.Parent.Parent.Text).Tables(Node.Parent.Text).ForeignKeys.Count(Not ctx.IncludeSys) Then
    While Not (Node.Child Is Nothing)
      tv.Nodes.Remove Node.Child.Index
    Wend
    For Each fky In svr.Databases(Node.Parent.Parent.Parent.Text).Tables(Node.Parent.Text).ForeignKeys
      If Not (fky.SystemObject And Not ctx.IncludeSys) Then tv.Nodes.Add Node.Key, tvwChild, "FKY-" & GetID, fky.Identifier, "foreignkey"
    Next fky
    Node.Text = "Foreign Keys (" & Node.Children & ")"
  End If
  lv.ColumnHeaders.Add , , "Foreign Key", 2000
  lv.ColumnHeaders.Add , , "References", lv.Width - 2100
  For Each fky In svr.Databases(Node.Parent.Parent.Parent.Text).Tables(Node.Parent.Text).ForeignKeys
    If Not (fky.SystemObject And Not ctx.IncludeSys) Then
      Set lvItem = lv.ListItems.Add(, "FKY-" & GetID, fky.Identifier, "foreignkey", "foreignkey")
      lvItem.SubItems(1) = fky.ReferencedTable
    End If
  Next fky
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvForeignKeys"
End Sub

Private Sub tvForeignKey(ByVal Node As MSComctlLib.Node)
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvForeignKey(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem

  If Node.Children = 0 Then tv.Nodes.Add Node.Key, tvwChild, "REL+" & GetID, "Relationships (" & svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).ForeignKeys(Node.Text).Relationships.Count & ")", "relationship"
  lv.ColumnHeaders.Add , , "Property", 2000
  lv.ColumnHeaders.Add , , "Value", lv.Width - 2100
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Name", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).ForeignKeys(Node.Text).Name
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "OID", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).ForeignKeys(Node.Text).OID
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "References", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).ForeignKeys(Node.Text).ReferencedTable
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "On Delete", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).ForeignKeys(Node.Text).OnDelete
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "On Update", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).ForeignKeys(Node.Text).OnUpdate
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Deferrable", "property", "property")
  If svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).ForeignKeys(Node.Text).Deferrable Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Initially", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).ForeignKeys(Node.Text).Initially
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "System Foreign Key?", "property", "property")
  If svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).ForeignKeys(Node.Text).SystemObject Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvForeignKey"
End Sub

Private Sub tvRelationships(ByVal Node As MSComctlLib.Node)
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvRelationships(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
Dim rel As pgRelationship

  lv.ColumnHeaders.Add , , "Local Column", 2000
  lv.ColumnHeaders.Add , , "Referenced Column", lv.Width - 2600
  Node.Text = "Relationships (" & svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Parent.Text).ForeignKeys(Node.Parent.Text).Relationships.Count & ")"
  For Each rel In svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Parent.Text).ForeignKeys(Node.Parent.Text).Relationships
    Set lvItem = lv.ListItems.Add(, "REL-" & GetID, rel.LocalColumn, "relationship", "relationship")
    lvItem.SubItems(1) = rel.ReferencedColumn
  Next rel
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvRelationships"
End Sub

Private Sub tvIndexes(ByVal Node As MSComctlLib.Node)
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvIndexes(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
Dim ind As pgIndex

  If Node.Children = 0 Or Node.Children <> svr.Databases(Node.Parent.Parent.Parent.Text).Tables(Node.Parent.Text).Indexes.Count(Not ctx.IncludeSys) Then
    While Not (Node.Child Is Nothing)
      tv.Nodes.Remove Node.Child.Index
    Wend
    For Each ind In svr.Databases(Node.Parent.Parent.Parent.Text).Tables(Node.Parent.Text).Indexes
      If Not (ind.SystemObject And Not ctx.IncludeSys) Then tv.Nodes.Add Node.Key, tvwChild, "IND-" & GetID, ind.Identifier, "index"
    Next ind
    Node.Text = "Indexes (" & Node.Children & ")"
  End If
  lv.ColumnHeaders.Add , , "Index", 2000
  lv.ColumnHeaders.Add , , "Comment", lv.Width - 2100
  For Each ind In svr.Databases(Node.Parent.Parent.Parent.Text).Tables(Node.Parent.Text).Indexes
    If Not (ind.SystemObject And Not ctx.IncludeSys) Then
      Set lvItem = lv.ListItems.Add(, "IND-" & GetID, ind.Identifier, "index", "index")
      lvItem.SubItems(1) = Replace(ind.Comment, vbCrLf, " ")
    End If
  Next ind
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvIndexes"
End Sub

Private Sub tvIndex(ByVal Node As MSComctlLib.Node)
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvIndex(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
Dim szTemp As String
Dim vData As Variant

  lv.ColumnHeaders.Add , , "Property", 2000
  lv.ColumnHeaders.Add , , "Value", lv.Width - 2100
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Name", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Indexes(Node.Text).Name
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "OID", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Indexes(Node.Text).OID
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Index Type", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Indexes(Node.Text).IndexType
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Unique?", "property", "property")
  If svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Indexes(Node.Text).Unique Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Primary?", "property", "property")
  If svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Indexes(Node.Text).Primary Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Column Count", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Indexes(Node.Text).IndexedColumns.Count
  For Each vData In svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Indexes(Node.Text).IndexedColumns
    szTemp = szTemp & vData & ", "
  Next vData
  If Len(szTemp) > 2 Then szTemp = Left(szTemp, Len(szTemp) - 2)
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Columns", "property", "property")
  lvItem.SubItems(1) = szTemp
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Constraint", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Indexes(Node.Text).Constraint
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "System Index?", "property", "property")
  If svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Indexes(Node.Text).SystemObject Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Comment", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Indexes(Node.Text).Comment
  
  'Set the Definition Pane
  If txtDefinition.Visible Then txtDefinition.Text = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Indexes(Node.Text).SQL

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvIndex"
End Sub

Private Sub tvRules(ByVal Node As MSComctlLib.Node)
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvRules(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
Dim rul As pgRule

  If Node.Children = 0 Or Node.Children <> svr.Databases(Node.Parent.Parent.Parent.Text).Tables(Node.Parent.Text).Rules.Count(Not ctx.IncludeSys) Then
    While Not (Node.Child Is Nothing)
      tv.Nodes.Remove Node.Child.Index
    Wend
    For Each rul In svr.Databases(Node.Parent.Parent.Parent.Text).Tables(Node.Parent.Text).Rules
      If Not (rul.SystemObject And Not ctx.IncludeSys) Then tv.Nodes.Add Node.Key, tvwChild, "RUL-" & GetID, rul.Identifier, "rule"
    Next rul
    Node.Text = "Rules (" & Node.Children & ")"
  End If
  lv.ColumnHeaders.Add , , "Rule", 2000
  lv.ColumnHeaders.Add , , "Comment", lv.Width - 2100
  For Each rul In svr.Databases(Node.Parent.Parent.Parent.Text).Tables(Node.Parent.Text).Rules
    If Not (rul.SystemObject And Not ctx.IncludeSys) Then
      Set lvItem = lv.ListItems.Add(, "RUL-" & GetID, rul.Identifier, "rule", "rule")
      lvItem.SubItems(1) = Replace(rul.Comment, vbCrLf, " ")
    End If
  Next rul
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvRules"
End Sub

Private Sub tvRule(ByVal Node As MSComctlLib.Node)
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvRule(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem

  lv.ColumnHeaders.Add , , "Property", 2000
  lv.ColumnHeaders.Add , , "Value", lv.Width - 2100
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Name", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Rules(Node.Text).Name
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "OID", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Rules(Node.Text).OID
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Event", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Rules(Node.Text).RuleEvent
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Condition", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Rules(Node.Text).Condition
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Do Instead?", "property", "property")
  If svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Rules(Node.Text).DoInstead Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Action", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Rules(Node.Text).Action
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Definition", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Rules(Node.Text).Definition
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "System Rule?", "property", "property")
  If svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Rules(Node.Text).SystemObject Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Comment", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Rules(Node.Text).Comment
  
  'Set the Definition Pane
  If txtDefinition.Visible Then txtDefinition.Text = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Rules(Node.Text).SQL
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvRule"
End Sub

Private Sub tvTriggers(ByVal Node As MSComctlLib.Node)
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvTriggers(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
Dim trg As pgTrigger

  If Node.Children = 0 Or Node.Children <> svr.Databases(Node.Parent.Parent.Parent.Text).Tables(Node.Parent.Text).Triggers.Count(Not ctx.IncludeSys) Then
    While Not (Node.Child Is Nothing)
      tv.Nodes.Remove Node.Child.Index
    Wend
    For Each trg In svr.Databases(Node.Parent.Parent.Parent.Text).Tables(Node.Parent.Text).Triggers
      If Not (trg.SystemObject And Not ctx.IncludeSys) Then tv.Nodes.Add Node.Key, tvwChild, "TRG-" & GetID, trg.Identifier, "trigger"
    Next trg
    Node.Text = "Triggers (" & Node.Children & ")"
  End If
  lv.ColumnHeaders.Add , , "Trigger", 2000
  lv.ColumnHeaders.Add , , "Comment", lv.Width - 2100
  For Each trg In svr.Databases(Node.Parent.Parent.Parent.Text).Tables(Node.Parent.Text).Triggers
    If Not (trg.SystemObject And Not ctx.IncludeSys) Then
      Set lvItem = lv.ListItems.Add(, "TRG-" & GetID, trg.Identifier, "trigger", "trigger")
      lvItem.SubItems(1) = Replace(trg.Comment, vbCrLf, " ")
    End If
  Next trg
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvTriggers"
End Sub

Private Sub tvTrigger(ByVal Node As MSComctlLib.Node)
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvTrigger(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem

  lv.ColumnHeaders.Add , , "Property", 2000
  lv.ColumnHeaders.Add , , "Value", lv.Width - 2100
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Name", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Triggers(Node.Text).Name
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "OID", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Triggers(Node.Text).OID
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Executes", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Triggers(Node.Text).Executes
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Event", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Triggers(Node.Text).TriggerEvent
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "For Each", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Triggers(Node.Text).ForEach
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Function", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Triggers(Node.Text).TriggerFunction
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "System Trigger?", "property", "property")
  If svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Triggers(Node.Text).SystemObject Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Comment", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Triggers(Node.Text).Comment
  
  'Set the Definition Pane
  If txtDefinition.Visible Then txtDefinition.Text = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Triggers(Node.Text).SQL
    
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvTrigger"
End Sub

Private Sub tvTypes(ByVal Node As MSComctlLib.Node)
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvTypes(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
Dim typ As pgType

  If Node.Children = 0 Or Node.Children <> svr.Databases(Node.Parent.Text).Types.Count(Not ctx.IncludeSys) Then
    While Not (Node.Child Is Nothing)
      tv.Nodes.Remove Node.Child.Index
    Wend
    For Each typ In svr.Databases(Node.Parent.Text).Types
      If Not (typ.SystemObject And Not ctx.IncludeSys) Then tv.Nodes.Add Node.Key, tvwChild, "TYP-" & GetID, typ.Identifier, "type"
    Next typ
    Node.Text = "Types (" & Node.Children & ")"
  End If
  lv.ColumnHeaders.Add , , "Type", 2000
  lv.ColumnHeaders.Add , , "Comment", lv.Width - 2100
  For Each typ In svr.Databases(Node.Parent.Text).Types
    If Not (typ.SystemObject And Not ctx.IncludeSys) Then
      Set lvItem = lv.ListItems.Add(, "TYP-" & GetID, typ.Identifier, "type", "type")
      lvItem.SubItems(1) = Replace(typ.Comment, vbCrLf, " ")
    End If
  Next typ
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvTypes"
End Sub

Private Sub tvType(ByVal Node As MSComctlLib.Node)
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvType(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem

  lv.ColumnHeaders.Add , , "Property", 2000
  lv.ColumnHeaders.Add , , "Value", lv.Width - 2100
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Name", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Types(Node.Text).Name
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "OID", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Types(Node.Text).OID
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Owner", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Types(Node.Text).Owner
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Input Function", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Types(Node.Text).InputFunction
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Output Function", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Types(Node.Text).OutputFunction
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Internal Length", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Types(Node.Text).InternalLength
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "External Length", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Types(Node.Text).ExternalLength
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Default", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Types(Node.Text).Default
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Element", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Types(Node.Text).Element
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Delimiter", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Types(Node.Text).Delimiter
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Send Function", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Types(Node.Text).SendFunction
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Receive Function", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Types(Node.Text).ReceiveFunction
    Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Passed by Value?", "property", "property")
  If svr.Databases(Node.Parent.Parent.Text).Types(Node.Text).PassedByValue Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Alignment", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Types(Node.Text).Alignment
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Storage", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Types(Node.Text).Storage
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "System Type?", "property", "property")
  If svr.Databases(Node.Parent.Parent.Text).Types(Node.Text).SystemObject Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Comment", "property", "property")
  lvItem.SubItems(1) = Replace(svr.Databases(Node.Parent.Parent.Text).Types(Node.Text).Comment, vbCrLf, " ")

  'Set the Definition Pane
  If txtDefinition.Visible Then txtDefinition.Text = svr.Databases(Node.Parent.Parent.Text).Types(Node.Text).SQL
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvType"
End Sub

Private Sub tvViews(ByVal Node As MSComctlLib.Node)
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvViews(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
  Dim vie As pgView
  If Node.Children = 0 Or Node.Children <> svr.Databases(Node.Parent.Text).Views.Count(Not ctx.IncludeSys) Then
    While Not (Node.Child Is Nothing)
      tv.Nodes.Remove Node.Child.Index
    Wend
    For Each vie In svr.Databases(Node.Parent.Text).Views
      If Not (vie.SystemObject And Not ctx.IncludeSys) Then tv.Nodes.Add Node.Key, tvwChild, "VIE-" & GetID, vie.Identifier, "view"
    Next vie
    Node.Text = "Views (" & Node.Children & ")"
  End If
  lv.ColumnHeaders.Add , , "View", 2000
  lv.ColumnHeaders.Add , , "Comment", lv.Width - 2100
  For Each vie In svr.Databases(Node.Parent.Text).Views
    If Not (vie.SystemObject And Not ctx.IncludeSys) Then
      Set lvItem = lv.ListItems.Add(, "VIE-" & GetID, vie.Identifier, "view", "view")
      lvItem.SubItems(1) = Replace(vie.Comment, vbCrLf, " ")
    End If
  Next vie
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvViews"
End Sub

Private Sub tvView(ByVal Node As MSComctlLib.Node)
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tvView(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem

  lv.ColumnHeaders.Add , , "Property", 2000
  lv.ColumnHeaders.Add , , "Value", lv.Width - 2100
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Name", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Views(Node.Text).Name
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "OID", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Views(Node.Text).OID
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Owner", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Views(Node.Text).Owner
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "ACL", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Views(Node.Text).ACL
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Definition", "property", "property")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Views(Node.Text).Definition
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "System View?", "property", "property")
  If svr.Databases(Node.Parent.Parent.Text).Views(Node.Text).SystemObject Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lv.ListItems.Add(, "PRO-" & GetID, "Comment", "property", "property")
  lvItem.SubItems(1) = Replace(svr.Databases(Node.Parent.Parent.Text).Views(Node.Text).Comment, vbCrLf, " ")
  
  'Set the Definition Pane
  If txtDefinition.Visible Then txtDefinition.Text = svr.Databases(Node.Parent.Parent.Text).Views(Node.Text).SQL
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvView"
End Sub

Public Sub tv_NodeClick(ByVal Node As MSComctlLib.Node)
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.tv_NodeClick(" & QUOTE & Node.FullPath & QUOTE & ")", etFullDebug

Dim lvItem As ListItem
Dim szTemp As String
Dim vData As Variant

  StartMsg "Examining database..."
  
  lv.ColumnHeaders.Clear
  lv.ListItems.Clear
  lv.Tag = Node.FullPath
  If txtDefinition.Visible Then txtDefinition.Text = ""
  
  Select Case Left(Node.Key, 4)

    Case "SVR-" 'Server
      tvServer Node
      Set ctx.CurrentObject = svr
      ctx.CurrentDB = ""
    
    Case "DAT+" 'Databases
      tvDatabases Node
      ctx.CurrentDB = ""
        
    Case "DAT-" 'Database
      tvDatabase Node
      Set ctx.CurrentObject = svr.Databases(Node.Text)
      ctx.CurrentDB = Node.Text
      
    Case "GRP+" 'Groups
      tvGroups Node
      ctx.CurrentDB = ""
      
    Case "GRP-" 'Group
      tvGroup Node
      Set ctx.CurrentObject = svr.Groups(Node.Text)
      ctx.CurrentDB = ""
      
    Case "USR+" 'Users
      tvUsers Node
      ctx.CurrentDB = ""
      
    Case "USR-" 'User
      tvUser Node
      Set ctx.CurrentObject = svr.Users(Node.Text)
      ctx.CurrentDB = ""
      
    Case "AGG+" 'Aggregates
      tvAggregates Node
      ctx.CurrentDB = Node.Parent.Text
      
    Case "AGG-" 'Aggregate
      tvAggregate Node
      Set ctx.CurrentObject = svr.Databases(Node.Parent.Parent.Text).Aggregates(Node.Text)
      ctx.CurrentDB = Node.Parent.Parent.Text
      
    Case "FNC+" 'Functions
      tvFunctions Node
      ctx.CurrentDB = Node.Parent.Text
      
    Case "FNC-" 'Function
      tvFunction Node
      Set ctx.CurrentObject = svr.Databases(Node.Parent.Parent.Text).Functions(Node.Text)
      ctx.CurrentDB = Node.Parent.Parent.Text
      
    Case "LNG+" 'Languages
      tvLanguages Node
      ctx.CurrentDB = Node.Parent.Text

    Case "LNG-" 'Language
      tvLanguage Node
      Set ctx.CurrentObject = svr.Databases(Node.Parent.Parent.Text).Languages(Node.Text)
      ctx.CurrentDB = Node.Parent.Parent.Text
      
    Case "OPR+" 'Operators
      tvOperators Node
      ctx.CurrentDB = Node.Parent.Text
      
    Case "OPR-" 'Operator
      tvOperator Node
      Set ctx.CurrentObject = svr.Databases(Node.Parent.Parent.Text).Operators(Node.Text)
      ctx.CurrentDB = Node.Parent.Parent.Text
      
    Case "SEQ+" 'Sequences
      tvSequences Node
      ctx.CurrentDB = Node.Parent.Text

    Case "SEQ-" 'Sequence
      tvSequence Node
      Set ctx.CurrentObject = svr.Databases(Node.Parent.Parent.Text).Sequences(Node.Text)
      ctx.CurrentDB = Node.Parent.Parent.Text
      
    Case "TBL+" 'Tables
      tvTables Node
      ctx.CurrentDB = Node.Parent.Text
      
    Case "TBL-" 'Table
      tvTable Node
      Set ctx.CurrentObject = svr.Databases(Node.Parent.Parent.Text).Tables(Node.Text)
      ctx.CurrentDB = Node.Parent.Parent.Text
      
    Case "CHK+" 'Checks
      tvChecks Node
      ctx.CurrentDB = Node.Parent.Parent.Parent.Text
      
    Case "CHK-" 'Check
      tvCheck Node
      Set ctx.CurrentObject = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Checks(Node.Text)
      ctx.CurrentDB = Node.Parent.Parent.Parent.Parent.Text
    
    Case "COL+" 'Columns
      tvColumns Node
      ctx.CurrentDB = Node.Parent.Parent.Parent.Text
      
    Case "COL-" 'Column
      tvColumn Node
      Set ctx.CurrentObject = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Columns(Node.Text)
      ctx.CurrentDB = Node.Parent.Parent.Parent.Parent.Text
      
    Case "FKY+" 'Foreign Keys
      tvForeignKeys Node
      ctx.CurrentDB = Node.Parent.Parent.Parent.Text
      
    Case "FKY-" 'Foreign Key
      tvForeignKey Node
      Set ctx.CurrentObject = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).ForeignKeys(Node.Text)
      ctx.CurrentDB = Node.Parent.Parent.Parent.Parent.Text
      
    Case "REL+" 'Relationships
      tvRelationships Node
      ctx.CurrentDB = Node.Parent.Parent.Parent.Text
      
    Case "IND+" 'Indexes
      tvIndexes Node
      ctx.CurrentDB = Node.Parent.Parent.Parent.Text
      
    Case "IND-" 'Index
      tvIndex Node
      Set ctx.CurrentObject = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Indexes(Node.Text)
      ctx.CurrentDB = Node.Parent.Parent.Parent.Parent.Text

    Case "RUL+" 'Rules
      tvRules Node
      ctx.CurrentDB = Node.Parent.Parent.Parent.Text
  
    Case "RUL-" 'Rule
      tvRule Node
      Set ctx.CurrentObject = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Rules(Node.Text)
      ctx.CurrentDB = Node.Parent.Parent.Parent.Parent.Text
      
    Case "TRG+" 'Triggers
      tvTriggers Node
      ctx.CurrentDB = Node.Parent.Parent.Parent.Text
      
    Case "TRG-" 'Trigger
      tvTrigger Node
      Set ctx.CurrentObject = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Triggers(Node.Text)
      ctx.CurrentDB = Node.Parent.Parent.Parent.Parent.Text
      
    Case "TYP+" 'Types
      tvTypes Node
      ctx.CurrentDB = Node.Parent.Text

    Case "TYP-" 'Type
      tvType Node
      Set ctx.CurrentObject = svr.Databases(Node.Parent.Parent.Text).Types(Node.Text)
      ctx.CurrentDB = Node.Parent.Parent.Text
      
    Case "VIE+" 'Views
      tvViews Node
      ctx.CurrentDB = Node.Parent.Text
      
    Case "VIE-" 'View
      tvView Node
      Set ctx.CurrentObject = svr.Databases(Node.Parent.Parent.Text).Views(Node.Text)
      ctx.CurrentDB = Node.Parent.Parent.Text
      
  End Select
  EndMsg
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.tvNodeClick"
End Sub

Private Sub lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.lv_ItemClick(" & QUOTE & Item.Text & QUOTE & ")", etFullDebug

Dim szPath() As String

  'Get the elements of the node path. This will indicate the path through the pgSchema hierarchy
  szPath = Split(lv.Tag, "\")
  
  Select Case Left(Item.Key, 4)

    Case "SVR-" 'Server
      Set ctx.CurrentObject = svr
      ctx.CurrentDB = ""
      If txtDefinition.Visible Then txtDefinition.Text = ""
        
    Case "DAT-" 'Database
      Set ctx.CurrentObject = svr.Databases(Item.Text)
      ctx.CurrentDB = Item.Text
      If txtDefinition.Visible Then txtDefinition.Text = ctx.CurrentObject.SQL
      
    Case "GRP-" 'Group
      Set ctx.CurrentObject = svr.Groups(Item.Text)
      ctx.CurrentDB = ""
      If txtDefinition.Visible Then txtDefinition.Text = ctx.CurrentObject.SQL
      
    Case "USR-" 'User
      Set ctx.CurrentObject = svr.Users(Item.Text)
      ctx.CurrentDB = ""
      If txtDefinition.Visible Then txtDefinition.Text = ctx.CurrentObject.SQL
      
    Case "AGG-" 'Aggregate
      Set ctx.CurrentObject = svr.Databases(szPath(2)).Aggregates(Item.Text)
      ctx.CurrentDB = szPath(2)
      If txtDefinition.Visible Then txtDefinition.Text = ctx.CurrentObject.SQL

    Case "FNC-" 'Function
      Set ctx.CurrentObject = svr.Databases(szPath(2)).Functions(Item.Text)
      ctx.CurrentDB = szPath(2)
      If txtDefinition.Visible Then txtDefinition.Text = ctx.CurrentObject.SQL

    Case "LNG-" 'Language
      Set ctx.CurrentObject = svr.Databases(szPath(2)).Languages(Item.Text)
      ctx.CurrentDB = szPath(2)
      If txtDefinition.Visible Then txtDefinition.Text = ctx.CurrentObject.SQL
      
    Case "OPR-" 'Operator
      Set ctx.CurrentObject = svr.Databases(szPath(2)).Operators(Item.Text)
      ctx.CurrentDB = szPath(2)
      If txtDefinition.Visible Then txtDefinition.Text = ctx.CurrentObject.SQL
 
    Case "SEQ-" 'Sequence
      Set ctx.CurrentObject = svr.Databases(szPath(2)).Sequences(Item.Text)
      ctx.CurrentDB = szPath(2)
      If txtDefinition.Visible Then txtDefinition.Text = ctx.CurrentObject.SQL

    Case "TBL-" 'Table
      Set ctx.CurrentObject = svr.Databases(szPath(2)).Tables(Item.Text)
      ctx.CurrentDB = szPath(2)
      If txtDefinition.Visible Then txtDefinition.Text = ctx.CurrentObject.SQL

    Case "CHK-" 'Check
      Set ctx.CurrentObject = svr.Databases(szPath(2)).Tables(szPath(4)).Checks(Item.Text)
      ctx.CurrentDB = szPath(2)
      If txtDefinition.Visible Then txtDefinition.Text = svr.Databases(ctx.CurrentDB).Tables(ctx.CurrentObject.Table).SQL
      
    Case "COL-" 'Column
      Set ctx.CurrentObject = svr.Databases(szPath(2)).Tables(szPath(4)).Columns(Item.Text)
      ctx.CurrentDB = szPath(2)
      If txtDefinition.Visible Then txtDefinition.Text = svr.Databases(ctx.CurrentDB).Tables(ctx.CurrentObject.Table).SQL

    Case "FKY-" 'Foreign Key
      Set ctx.CurrentObject = svr.Databases(szPath(2)).Tables(szPath(4)).ForeignKeys(Item.Text)
      ctx.CurrentDB = szPath(2)
      If txtDefinition.Visible Then txtDefinition.Text = svr.Databases(ctx.CurrentDB).Tables(ctx.CurrentObject.Table).SQL
      
    Case "IND-" 'Index
      Set ctx.CurrentObject = svr.Databases(szPath(2)).Tables(szPath(4)).Indexes(Item.Text)
      ctx.CurrentDB = szPath(2)
      If txtDefinition.Visible Then txtDefinition.Text = ctx.CurrentObject.SQL
  
    Case "RUL-" 'Rule
      Set ctx.CurrentObject = svr.Databases(szPath(2)).Tables(szPath(4)).Rules(Item.Text)
      ctx.CurrentDB = szPath(2)
      If txtDefinition.Visible Then txtDefinition.Text = ctx.CurrentObject.SQL
      
    Case "TRG-" 'Trigger
      Set ctx.CurrentObject = svr.Databases(szPath(2)).Tables(szPath(4)).Triggers(Item.Text)
      ctx.CurrentDB = szPath(2)
      If txtDefinition.Visible Then txtDefinition.Text = ctx.CurrentObject.SQL

    Case "TYP-" 'Type
      Set ctx.CurrentObject = svr.Databases(szPath(2)).Types(Item.Text)
      ctx.CurrentDB = szPath(2)
      If txtDefinition.Visible Then txtDefinition.Text = ctx.CurrentObject.SQL
      
    Case "VIE-" 'View
      Set ctx.CurrentObject = svr.Databases(szPath(2)).Views(Item.Text)
      ctx.CurrentDB = szPath(2)
      If txtDefinition.Visible Then txtDefinition.Text = ctx.CurrentObject.SQL
      
  End Select
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.lv_ItemClick"
End Sub

Private Sub txtDefinition_Change()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmMain.txtDefinition_Change()", etFullDebug
  
  If txtDefinition.Text = "" Then
    mnuFileSaveDefinition.Enabled = False
  Else
    mnuFileSaveDefinition.Enabled = True
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmMain.txtDefinition_Change"
End Sub

