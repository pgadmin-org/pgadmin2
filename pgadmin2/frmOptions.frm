VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{44F33AC4-8757-4330-B063-18608617F23E}#12.4#0"; "HighlightBox.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   6888
   ClientLeft      =   3240
   ClientTop       =   1860
   ClientWidth     =   5532
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6888
   ScaleWidth      =   5532
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   0
      Top             =   6390
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab tabOptions 
      Height          =   6360
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   5460
      _ExtentX        =   9631
      _ExtentY        =   11218
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "&Logging"
      TabPicture(0)   =   "frmOptions.frx":0A02
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtLogFile"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdBrowse"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraLogLevel"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "chkLogWindow"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "chkMaskPassword"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "&Text"
      TabPicture(1)   =   "frmOptions.frx":0A1E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame8"
      Tab(1).Control(1)=   "Frame7"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "&Exporters"
      TabPicture(2)   =   "frmOptions.frx":0A3A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lstExporters"
      Tab(2).Control(1)=   "Frame1"
      Tab(2).Control(2)=   "cmdExpInstall"
      Tab(2).Control(3)=   "cmdExpUninstall"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "&Plugins"
      TabPicture(3)   =   "frmOptions.frx":0A56
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdPlgUninstall"
      Tab(3).Control(1)=   "cmdPlgInstall"
      Tab(3).Control(2)=   "Frame2"
      Tab(3).Control(3)=   "lstPlugins"
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "&PostgreSQL"
      TabPicture(4)   =   "frmOptions.frx":0A72
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame6"
      Tab(4).Control(1)=   "Frame5"
      Tab(4).Control(2)=   "Frame4"
      Tab(4).Control(3)=   "Frame3"
      Tab(4).ControlCount=   4
      TabCaption(5)   =   "Misc"
      TabPicture(5)   =   "frmOptions.frx":0A8E
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label5"
      Tab(5).Control(1)=   "Label3"
      Tab(5).Control(2)=   "UpDownMaxSqlQuery"
      Tab(5).Control(3)=   "chkAskDeleteObjectDatabase"
      Tab(5).Control(4)=   "chkShowUsersForPrivileges"
      Tab(5).Control(5)=   "UpDownMaxRecViewData"
      Tab(5).Control(6)=   "txtMaxRecordViewData"
      Tab(5).Control(7)=   "txtMaxSqlQuery"
      Tab(5).Control(8)=   "Frame9"
      Tab(5).Control(9)=   "Frame10"
      Tab(5).ControlCount=   10
      Begin VB.Frame Frame10 
         Caption         =   "Language Traslation"
         Height          =   612
         Left            =   -74880
         TabIndex        =   79
         Top             =   1800
         Width           =   5172
         Begin VB.ComboBox cboLang 
            Height          =   288
            Left            =   60
            Style           =   2  'Dropdown List
            TabIndex        =   80
            Top             =   240
            Width           =   2952
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Generate Documentation Database"
         Height          =   1692
         Left            =   -74880
         TabIndex        =   72
         Top             =   2460
         Visible         =   0   'False
         Width           =   5172
         Begin VB.TextBox txtDocWorkPath 
            Height          =   285
            Left            =   120
            TabIndex        =   77
            ToolTipText     =   "Enter a path on work directory"
            Top             =   1236
            Width           =   4515
         End
         Begin VB.CommandButton cmdDocWorkBrw 
            Caption         =   "..."
            Height          =   375
            Left            =   4668
            TabIndex        =   76
            ToolTipText     =   "Browse for a file."
            Top             =   1200
            Width           =   375
         End
         Begin VB.TextBox txtDocHHWPath 
            Height          =   285
            Left            =   120
            TabIndex        =   74
            ToolTipText     =   "Enter a path on program hhc.exe"
            Top             =   576
            Width           =   4515
         End
         Begin VB.CommandButton cmdDocHHWBrw 
            Caption         =   "..."
            Height          =   375
            Left            =   4668
            TabIndex        =   73
            ToolTipText     =   "Browse for a file."
            Top             =   528
            Width           =   375
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Work directory"
            Height          =   192
            Index           =   6
            Left            =   120
            TabIndex        =   78
            Top             =   960
            Width           =   1044
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "HTML Help Workshop directory"
            Height          =   192
            Index           =   5
            Left            =   120
            TabIndex        =   75
            Top             =   300
            Width           =   2268
         End
      End
      Begin VB.TextBox txtMaxSqlQuery 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -74880
         TabIndex        =   69
         Text            =   "0"
         Top             =   1380
         Width           =   492
      End
      Begin VB.TextBox txtMaxRecordViewData 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -74880
         TabIndex        =   67
         Text            =   "0"
         Top             =   1020
         Width           =   492
      End
      Begin MSComCtl2.UpDown UpDownMaxRecViewData 
         Height          =   312
         Left            =   -74400
         TabIndex        =   66
         Top             =   1020
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   550
         _Version        =   393216
         BuddyControl    =   "txtMaxRecordViewData"
         BuddyDispid     =   196618
         OrigLeft        =   720
         OrigTop         =   1020
         OrigRight       =   960
         OrigBottom      =   1332
         Max             =   100000
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.CheckBox chkShowUsersForPrivileges 
         Caption         =   "Show users for privileges"
         Height          =   240
         Left            =   -74880
         TabIndex        =   65
         Top             =   420
         Width           =   3345
      End
      Begin VB.CheckBox chkAskDeleteObjectDatabase 
         Caption         =   "Ask delete object database"
         Height          =   240
         Left            =   -74880
         TabIndex        =   64
         Top             =   708
         Width           =   3345
      End
      Begin VB.Frame Frame8 
         Caption         =   "Font"
         Height          =   1320
         Left            =   -74775
         TabIndex        =   59
         Top             =   495
         Width           =   5010
         Begin VB.CommandButton cmdBrowseFont 
            Caption         =   "..."
            Height          =   330
            Left            =   4095
            TabIndex        =   62
            Top             =   270
            Width           =   420
         End
         Begin VB.TextBox txtFont 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   540
            Locked          =   -1  'True
            TabIndex        =   60
            ToolTipText     =   "Enter the name of a database to use as the Master Connection."
            Top             =   270
            Width           =   3525
         End
         Begin VB.Label Label4 
            Caption         =   "This is the font used for display of data in the Treeview, Listview, text boxes and the Data Grid."
            Height          =   600
            Index           =   4
            Left            =   135
            TabIndex        =   61
            Top             =   630
            Width           =   4695
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Auto Highlight"
         Height          =   4020
         Left            =   -74775
         TabIndex        =   50
         Top             =   2025
         Width           =   5010
         Begin VB.CommandButton cmdDefault 
            Caption         =   "Default"
            Height          =   345
            Left            =   2880
            TabIndex        =   63
            ToolTipText     =   "Restore default words."
            Top             =   720
            Width           =   945
         End
         Begin VB.CheckBox chkItalic 
            Caption         =   "Italic"
            Height          =   285
            Left            =   1080
            TabIndex        =   56
            ToolTipText     =   "Should the word be made italic?"
            Top             =   750
            Width           =   675
         End
         Begin VB.CheckBox chkBold 
            Caption         =   "Bold"
            Height          =   285
            Left            =   225
            TabIndex        =   55
            ToolTipText     =   "Should the word be made bold?"
            Top             =   750
            Width           =   690
         End
         Begin VB.CommandButton cmdColour 
            Caption         =   "&Colour"
            Height          =   330
            Left            =   2880
            Style           =   1  'Graphical
            TabIndex        =   54
            ToolTipText     =   "Select a colour for the word."
            Top             =   270
            Width           =   945
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Add"
            Height          =   345
            Left            =   3915
            TabIndex        =   53
            ToolTipText     =   "Add the selected word."
            Top             =   270
            Width           =   945
         End
         Begin VB.TextBox txtWord 
            Height          =   285
            Left            =   720
            TabIndex        =   52
            ToolTipText     =   "Enter a word to highlight."
            Top             =   300
            Width           =   2055
         End
         Begin VB.CommandButton cmdRemove 
            Caption         =   "Remove"
            Height          =   345
            Left            =   3915
            TabIndex        =   51
            ToolTipText     =   "Remove the selected word."
            Top             =   705
            Width           =   945
         End
         Begin MSComctlLib.ListView lvWords 
            Height          =   2715
            Left            =   90
            TabIndex        =   57
            ToolTipText     =   "Displays the Text Formatting rules."
            Top             =   1170
            Width           =   4785
            _ExtentX        =   8446
            _ExtentY        =   4784
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
            NumItems        =   0
         End
         Begin VB.Label Label2 
            Caption         =   "Word"
            Height          =   255
            Left            =   180
            TabIndex        =   58
            Top             =   315
            Width           =   465
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Defer Connection"
         Height          =   1320
         Left            =   -74775
         TabIndex        =   47
         Top             =   4815
         Width           =   5010
         Begin VB.CheckBox chkDeferConnection 
            Caption         =   "Don't connect until necessary."
            Height          =   240
            Left            =   810
            TabIndex        =   48
            Top             =   315
            Width           =   3345
         End
         Begin VB.Label Label4 
            Caption         =   $"frmOptions.frx":0AAA
            Height          =   600
            Index           =   3
            Left            =   135
            TabIndex        =   49
            Top             =   630
            Width           =   4695
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Auto Row Count"
         Height          =   1320
         Left            =   -74775
         TabIndex        =   44
         Top             =   3375
         Width           =   5010
         Begin VB.CheckBox chkAutoRowCount 
            Caption         =   "Use auto row count on tables and views."
            Height          =   240
            Left            =   810
            TabIndex        =   45
            Top             =   315
            Width           =   3345
         End
         Begin VB.Label Label4 
            Caption         =   $"frmOptions.frx":0B52
            Height          =   600
            Index           =   2
            Left            =   135
            TabIndex        =   46
            Top             =   630
            Width           =   4695
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Security"
         Height          =   1320
         Left            =   -74775
         TabIndex        =   41
         Top             =   1935
         Width           =   5010
         Begin VB.CheckBox chkEncryptPasswords 
            Caption         =   "Use Encrypted passwords where possible."
            Height          =   240
            Left            =   810
            TabIndex        =   42
            Top             =   315
            Width           =   3345
         End
         Begin VB.Label Label4 
            Caption         =   $"frmOptions.frx":0C17
            Height          =   645
            Index           =   1
            Left            =   225
            TabIndex        =   43
            Top             =   585
            Width           =   4695
         End
      End
      Begin VB.CommandButton cmdPlgUninstall 
         Caption         =   "&Uninstall Plugin"
         Height          =   330
         Left            =   -73200
         TabIndex        =   37
         ToolTipText     =   "Uninstall the selected Plugin."
         Top             =   5895
         Width           =   1590
      End
      Begin VB.CommandButton cmdPlgInstall 
         Caption         =   "&Install Plugin"
         Height          =   330
         Left            =   -74910
         TabIndex        =   36
         ToolTipText     =   "Install a new Plugin."
         Top             =   5895
         Width           =   1590
      End
      Begin VB.Frame Frame2 
         Caption         =   "Details"
         Height          =   1950
         Left            =   -74910
         TabIndex        =   30
         Top             =   3870
         Width           =   5235
         Begin VB.TextBox txtPlgVersion 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   33
            Top             =   540
            Width           =   4110
         End
         Begin VB.TextBox txtPlgDescription 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   32
            Top             =   225
            Width           =   4110
         End
         Begin HighlightBox.TBX txtPlgAuthor 
            Height          =   945
            Left            =   90
            TabIndex        =   31
            Top             =   900
            Width           =   5055
            _ExtentX        =   8911
            _ExtentY        =   1672
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
            Caption         =   "Author"
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Version"
            Height          =   195
            Index           =   4
            Left            =   135
            TabIndex        =   35
            Top             =   540
            Width           =   525
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Description"
            Height          =   195
            Index           =   3
            Left            =   135
            TabIndex        =   34
            Top             =   270
            Width           =   795
         End
      End
      Begin VB.ListBox lstPlugins 
         Height          =   3120
         ItemData        =   "frmOptions.frx":0CAB
         Left            =   -74910
         List            =   "frmOptions.frx":0CAD
         TabIndex        =   29
         Top             =   450
         Width           =   5235
      End
      Begin VB.ListBox lstExporters 
         Height          =   3120
         ItemData        =   "frmOptions.frx":0CAF
         Left            =   -74910
         List            =   "frmOptions.frx":0CB1
         TabIndex        =   20
         Top             =   450
         Width           =   5235
      End
      Begin VB.Frame Frame1 
         Caption         =   "Details"
         Height          =   1950
         Left            =   -74910
         TabIndex        =   26
         Top             =   3870
         Width           =   5235
         Begin HighlightBox.TBX txtExpAuthor 
            Height          =   945
            Left            =   90
            TabIndex        =   23
            Top             =   900
            Width           =   5055
            _ExtentX        =   8911
            _ExtentY        =   1672
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
            Caption         =   "Author"
         End
         Begin VB.TextBox txtExpDescription 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   21
            Top             =   225
            Width           =   4110
         End
         Begin VB.TextBox txtExpVersion 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   540
            Width           =   4110
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Description"
            Height          =   195
            Index           =   2
            Left            =   135
            TabIndex        =   28
            Top             =   270
            Width           =   795
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Version"
            Height          =   195
            Index           =   1
            Left            =   135
            TabIndex        =   27
            Top             =   540
            Width           =   525
         End
      End
      Begin VB.CommandButton cmdExpInstall 
         Caption         =   "&Install Exporter"
         Height          =   330
         Left            =   -74910
         TabIndex        =   24
         ToolTipText     =   "Install a new Exporter."
         Top             =   5895
         Width           =   1590
      End
      Begin VB.CommandButton cmdExpUninstall 
         Caption         =   "&Uninstall Exporter"
         Height          =   330
         Left            =   -73200
         TabIndex        =   25
         ToolTipText     =   "Uninstall the selected Exporter."
         Top             =   5895
         Width           =   1590
      End
      Begin VB.CheckBox chkMaskPassword 
         Caption         =   "&Mask the Password in Logs?"
         Height          =   285
         Left            =   225
         TabIndex        =   9
         ToolTipText     =   "Check to replace the occurance of the user's password in any logs with *********."
         Top             =   5355
         Width           =   4155
      End
      Begin VB.CheckBox chkLogWindow 
         Caption         =   "Log Window 'Always on top'?"
         Height          =   285
         Left            =   225
         TabIndex        =   8
         ToolTipText     =   "Make the Log Window always appear on top of other windows regardless of whether it has focus."
         Top             =   4725
         Width           =   4155
      End
      Begin VB.Frame fraLogLevel 
         Caption         =   "Log Level"
         Height          =   2175
         Left            =   450
         TabIndex        =   19
         Top             =   1980
         Width           =   4560
         Begin VB.OptionButton optLogLevel 
            Caption         =   "&Full debug"
            Height          =   240
            Index           =   4
            Left            =   1260
            TabIndex        =   7
            ToolTipText     =   "Log everything. Warning - this can be *very* slow and can create huge logfiles."
            Top             =   1665
            Width           =   3120
         End
         Begin VB.OptionButton optLogLevel 
            Caption         =   "&Debug"
            Height          =   240
            Index           =   3
            Left            =   1260
            TabIndex        =   6
            ToolTipText     =   "Log errors, SQL queries and important debug messages."
            Top             =   1350
            Width           =   3120
         End
         Begin VB.OptionButton optLogLevel 
            Caption         =   "Errors and &SQL queries"
            Height          =   240
            Index           =   2
            Left            =   1260
            TabIndex        =   5
            ToolTipText     =   "Log errors and SQL queries."
            Top             =   1035
            Width           =   3120
         End
         Begin VB.OptionButton optLogLevel 
            Caption         =   "&Errors only"
            Height          =   240
            Index           =   1
            Left            =   1260
            TabIndex        =   4
            ToolTipText     =   "Log errors only."
            Top             =   720
            Width           =   3120
         End
         Begin VB.OptionButton optLogLevel 
            Caption         =   "&No logging"
            Height          =   240
            Index           =   0
            Left            =   1260
            TabIndex        =   3
            ToolTipText     =   "Don't perform any logging."
            Top             =   405
            Width           =   3120
         End
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "..."
         Height          =   375
         Left            =   4770
         TabIndex        =   2
         ToolTipText     =   "Browse for a file."
         Top             =   1125
         Width           =   375
      End
      Begin VB.TextBox txtLogFile 
         Height          =   285
         Left            =   225
         TabIndex        =   1
         ToolTipText     =   "Enter a path & filename to write the logfile to."
         Top             =   1170
         Width           =   4515
      End
      Begin VB.Frame Frame3 
         Caption         =   "Master Connection Database"
         Height          =   1320
         Left            =   -74775
         TabIndex        =   38
         Top             =   495
         Width           =   5010
         Begin VB.TextBox txtMasterDB 
            Height          =   285
            Left            =   540
            TabIndex        =   39
            ToolTipText     =   "Enter the name of a database to use as the Master Connection."
            Top             =   270
            Width           =   3930
         End
         Begin VB.Label Label4 
            Caption         =   $"frmOptions.frx":0CB3
            Height          =   600
            Index           =   0
            Left            =   135
            TabIndex        =   40
            Top             =   630
            Width           =   4695
         End
      End
      Begin MSComCtl2.UpDown UpDownMaxSqlQuery 
         Height          =   288
         Left            =   -74400
         TabIndex        =   70
         Top             =   1380
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   508
         _Version        =   393216
         BuddyControl    =   "txtMaxSqlQuery"
         BuddyDispid     =   196617
         OrigLeft        =   720
         OrigTop         =   1020
         OrigRight       =   960
         OrigBottom      =   1332
         Max             =   10000
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label Label3 
         Caption         =   "Max number of sql command to memorize (Query). "
         Height          =   240
         Left            =   -74100
         TabIndex        =   71
         Top             =   1440
         Width           =   3972
      End
      Begin VB.Label Label5 
         Caption         =   "Max number of record for view data. "
         Height          =   240
         Left            =   -74100
         TabIndex        =   68
         Top             =   1080
         Width           =   4032
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Logfile (%ID will be replaced with the Process ID)"
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   18
         Top             =   900
         Width           =   3450
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5688
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   17
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5688
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   16
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5688
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   1785
         Left            =   645
         TabIndex        =   15
         Top             =   300
         Width           =   2055
      End
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
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3312
      TabIndex        =   10
      Top             =   6480
      Width           =   1095
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the Artistic Licence
'
' frmLog.frm - Displays the rolling log

Option Explicit


Private Sub cmdAdd_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmOptions.cmdAdd_Click()", etFullDebug

Dim itmX As ListItem

  If txtWord.Text = "" Then
    MsgBox ??TrasLang??("You must enter a word to add!"), vbExclamation, ??TrasLang??("Error")
    txtWord.SetFocus
    Exit Sub
  End If
  For Each itmX In lvWords.ListItems
    If itmX.Text = txtWord.Text Then
      MsgBox ??TrasLang??("That word is already in the list!"), vbExclamation, ??TrasLang??("Error")
      txtWord.SetFocus
      Exit Sub
    End If
  Next itmX

  'Add the new listitem
  Set itmX = lvWords.ListItems.Add(, , txtWord.Text)
  itmX.SubItems(1) = txtWord.ForeColor
  If chkBold = "1" Then
    itmX.SubItems(2) = "Y"
  Else
    itmX.SubItems(2) = "N"
  End If
  If chkItalic.Value = "1" Then
    itmX.SubItems(3) = "Y"
  Else
    itmX.SubItems(3) = "N"
  End If

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmOptions.cmdAdd_Click"
End Sub

Private Sub cmdBrowse_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmOptions.cmdBrowse_Click()", etFullDebug

  With cdlg
    .FileName = txtLogFile.Text
    .DialogTitle = ??TrasLang??("Log File")
    .Filter = "All Files (*.*)|*.*"
    .CancelError = False
    .FLAGS = &H4
    .ShowOpen
  End With
  txtLogFile.Text = cdlg.FileName

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmOptions.cmdBrowse_Click"
End Sub

Private Sub cmdBrowseFont_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmOptions.cmdBrowseFont_Click()", etFullDebug

Dim szFont() As String

  'Extract the
  cdlg.CancelError = True
  cdlg.DialogTitle = ??TrasLang??("Data Font")
  cdlg.FLAGS = cdlCFBoth
  cdlg.ShowFont
  txtFont.Tag = cdlg.FontName & "|" & cdlg.FontSize & "|" & cdlg.FontBold & "|" & cdlg.FontItalic
  txtFont.Text = cdlg.FontName & ", " & cdlg.FontSize & "pt"
  If cdlg.FontBold Then txtFont.Text = txtFont.Text & ", bold"
  If cdlg.FontItalic Then txtFont.Text = txtFont.Text & ", italic"
  
  Exit Sub
Err_Handler:
  If Err.Number = 32755 Then Exit Sub
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmOptions.cmdCancel_Click"
End Sub

Private Sub cmdCancel_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmOptions.cmdCancel_Click()", etFullDebug

  Unload Me
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmOptions.cmdCancel_Click"
End Sub

Private Sub cmdDefault_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmOptions.cmdDefault_Click()", etFullDebug

  'load default
  If MsgBox(??TrasLang??("Are you sure you wish to restore default word?"), vbQuestion + vbYesNo, ??TrasLang??("Restore default Word")) = vbNo Then Exit Sub
  LoadWord szDefaultAutoHighlight
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmOptions.cmdDefault_Click"
End Sub

Private Sub cmdDocHHWBrw_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmOptions.cmdDocHHWBrw_Click()", etFullDebug

Dim szTemp As String

  szTemp = BrowseFolder(0, ??TrasLang??("Select folder on HTML Help Workshop"))
  If Len(szTemp) > 0 Then txtDocHHWPath.Text = szTemp

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmDoc.cmdSelDir_Click"
End Sub

Private Sub cmdDocWorkBrw_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmOptions.cmdDocWorkBrw_Click()", etFullDebug

Dim szTemp As String

  szTemp = BrowseFolder(0, ??TrasLang??("Select folder to work"))
  If Len(szTemp) > 0 Then txtDocWorkPath.Text = szTemp

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmDoc.cmdSelDir_Click"
End Sub

Private Sub cmdOK_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmOptions.cmdOK_Click()", etFullDebug

Dim iLogLevel As Integer
Dim objForm As Form
Dim szTextColours As String
Dim itmX As ListItem
Dim szFont() As String
Dim objFont As New StdFont

  'Save settings, and make them live
  'Logfile
  frmMain.svr.Logfile = txtLogFile.Text
  RegWrite HKEY_CURRENT_USER, "Software\" & App.Title, "Log File", regString, txtLogFile.Text
  
  'Log Level
  For iLogLevel = 0 To 4
    If optLogLevel(iLogLevel).Value = True Then Exit For
  Next iLogLevel
  ctx.LogLevel = iLogLevel
  frmMain.svr.LogLevel = ctx.LogLevel
  RegWrite HKEY_CURRENT_USER, "Software\" & App.Title, "Log Level", regString, iLogLevel
  
  'Log Window Always On Top
  'Find the log window if it's open
  For Each objForm In Forms
    If objForm.Name = "frmLog" Then Exit For
  Next objForm
  
  If chkLogWindow.Value = 1 Then
    If Not (objForm Is Nothing) Then SetTopMostWindow objForm.hwnd, True
    RegWrite HKEY_CURRENT_USER, "Software\" & App.Title & "\Log Window", "Always On Top", regString, "Y"
  Else
    If Not (objForm Is Nothing) Then SetTopMostWindow objForm.hwnd, False
    RegWrite HKEY_CURRENT_USER, "Software\" & App.Title & "\Log Window", "Always On Top", regString, "N"
  End If
  
  'Mask Password
  If chkLogWindow.Value = 1 Then
    frmMain.svr.ShowPassword = False
    RegWrite HKEY_CURRENT_USER, "Software\" & App.Title, "Mask Password", regString, "Y"
  Else
    frmMain.svr.ShowPassword = True
    RegWrite HKEY_CURRENT_USER, "Software\" & App.Title, "Mask Password", regString, "N"
  End If
  
  'Font
  szFont = Split(txtFont.Tag, "|")
  objFont.Name = szFont(0)
  objFont.Size = Val(szFont(1))
  objFont.Bold = CBool(szFont(2))
  objFont.Italic = CBool(szFont(3))
  Set ctx.Font = objFont
  PatchForm frmMain
  RegWrite HKEY_CURRENT_USER, "Software\" & App.Title, "Font", regString, CStr(txtFont.Tag)
  
  'Autohighlight Colours
  For Each itmX In lvWords.ListItems
    szTextColours = szTextColours & itmX.Text & "|"
    If itmX.SubItems(2) = "Y" Then
      szTextColours = szTextColours & "1|"
    Else
      szTextColours = szTextColours & "0|"
    End If
    If itmX.SubItems(3) = "Y" Then
      szTextColours = szTextColours & "1|"
    Else
      szTextColours = szTextColours & "0|"
    End If
    szTextColours = szTextColours & itmX.SubItems(1) & ";"
  Next itmX
  ctx.AutoHighlight = szTextColours
  RegWrite HKEY_CURRENT_USER, "Software\" & App.Title, "AutoHighlight", regString, CStr(ctx.AutoHighlight)
    
  'Master DB
  If txtMasterDB.Text <> RegRead(HKEY_CURRENT_USER, "Software\" & App.Title, "Master DB", "template1") And _
     frmMain.svr.ConnectionString <> "" Then
    MsgBox "The change to the Master Connection Database will not take effect until you reconnect to the server.", vbInformation, "Master Connection Database"
  End If
  RegWrite HKEY_CURRENT_USER, "Software\" & App.Title, "Master DB", regString, txtMasterDB.Text
  
  'Encrypted passwords
  If chkEncryptPasswords.Value = 1 Then
    frmMain.svr.EncryptPasswords = True
    RegWrite HKEY_CURRENT_USER, "Software\" & App.Title, "Encrypt Passwords", regString, "Y"
  Else
    frmMain.svr.EncryptPasswords = False
    RegWrite HKEY_CURRENT_USER, "Software\" & App.Title, "Encrypt Passwords", regString, "N"
  End If
  
  'Auto Rowcount
  If chkAutoRowCount.Value = 1 Then
    ctx.AutoRowCount = True
    RegWrite HKEY_CURRENT_USER, "Software\" & App.Title, "Auto Row Count", regString, "Y"
  Else
    ctx.AutoRowCount = False
    RegWrite HKEY_CURRENT_USER, "Software\" & App.Title, "Auto Row Count", regString, "N"
  End If
  
  'Defer Connections
  If chkDeferConnection.Value = 1 Then
    frmMain.svr.DeferConnection = True
    RegWrite HKEY_CURRENT_USER, "Software\" & App.Title, "Defer Connection", regString, "Y"
  Else
    frmMain.svr.DeferConnection = False
    RegWrite HKEY_CURRENT_USER, "Software\" & App.Title, "Defer Connection", regString, "N"
  End If
  
  'Show users for Privileges
  If chkShowUsersForPrivileges.Value = 1 Then
    ctx.ShowUsersForPrivileges = True
    RegWrite HKEY_CURRENT_USER, "Software\" & App.Title, "Show Users For Privileges", regString, "Y"
  Else
    ctx.ShowUsersForPrivileges = False
    RegWrite HKEY_CURRENT_USER, "Software\" & App.Title, "Show Users For Privileges", regString, "N"
  End If
  
  'Ask delete object database
  If chkAskDeleteObjectDatabase.Value = 1 Then
    ctx.AskDeleteObjectDatabase = True
    RegWrite HKEY_CURRENT_USER, "Software\" & App.Title, "Ask Delete Object Database", regString, "Y"
  Else
    ctx.AskDeleteObjectDatabase = False
    RegWrite HKEY_CURRENT_USER, "Software\" & App.Title, "Ask Delete Object Database", regString, "N"
  End If
  
  'max number of sql command in query
  ctx.MaxNumberSqlQuery = Val(txtMaxSqlQuery.Text)
  RegWrite HKEY_CURRENT_USER, "Software\" & App.Title, "Max Number Sql Query", regString, Val(txtMaxSqlQuery.Text)
  
  'max number of Record in View Data
  ctx.MaxRecordViewData = Val(txtMaxRecordViewData.Text)
  RegWrite HKEY_CURRENT_USER, "Software\" & App.Title, "Row Limit", regString, Val(txtMaxRecordViewData.Text)
  
  '///////////////////////////////////
  'Generate Documentation
  'HTML Help Workshop directory
  RegWrite HKEY_CURRENT_USER, "Software\" & App.Title & "\GenDbDoc", "HTML Help Workshop directory", regString, CStr(txtDocHHWPath.Text)
  'Work directory
  RegWrite HKEY_CURRENT_USER, "Software\" & App.Title & "\GenDbDoc", "Work directory", regString, CStr(txtDocWorkPath.Text)
  
  'save lang
  If cboLang.Text <> RegRead(HKEY_CURRENT_USER, "Software\" & App.Title, "Current Lang", "") Then
    RegWrite HKEY_CURRENT_USER, "Software\" & App.Title, "Current Lang", regString, cboLang.Text
    InitLang cboLang.Text
    MsgBox ??TrasLang??("For applay change of lang restart pgAdmin2!"), vbInformation
  End If

  Unload Me
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmOptions.cmdOK_Click"
End Sub

Private Sub cmdRemove_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmOptions.cmdRemove_Click()", etFullDebug

  If MsgBox(??TrasLang??("Are you sure you wish to remove the selected word?"), vbQuestion + vbYesNo, ??TrasLang??("Remove Word")) = vbNo Then Exit Sub
  lvWords.ListItems.Remove lvWords.SelectedItem.Index
      
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmOptions.cmdRemove_Click"
End Sub

Private Sub Form_Load()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmOptions.Form_Load()", etFullDebug

Dim szFont() As String
Dim szTemp As String
Dim szCurrentLang As String

  PatchForm Me
  
  'Get the current settings.
  'We use the registry settings because (for example) frmMain.svr.Logfile will return the actual filename, not the code.
  txtLogFile.Text = RegRead(HKEY_CURRENT_USER, "Software\" & App.Title, "Log File", "C:\" & App.Title & "_%ID.Log")
  Select Case Val(RegRead(HKEY_CURRENT_USER, "Software\" & App.Title, "Log Level", "2"))
    Case 0
      optLogLevel(0).Value = True
    Case 1
      optLogLevel(1).Value = True
    Case 2
      optLogLevel(2).Value = True
    Case 3
      optLogLevel(3).Value = True
    Case 4
      optLogLevel(4).Value = True
  End Select
  If UCase(RegRead(HKEY_CURRENT_USER, "Software\" & App.Title & "\Log Window", "Always On Top", "Y")) = "Y" Then
    chkLogWindow.Value = 1
  Else
    chkLogWindow.Value = 0
  End If
  If UCase(RegRead(HKEY_CURRENT_USER, "Software\" & App.Title, "Mask Password", "Y")) = "Y" Then
    chkMaskPassword.Value = 1
  Else
    chkMaskPassword.Value = 0
  End If
  If UCase(RegRead(HKEY_CURRENT_USER, "Software\" & App.Title, "Show Users For Privileges", "Y")) = "Y" Then
    chkShowUsersForPrivileges.Value = 1
  Else
    chkShowUsersForPrivileges.Value = 0
  End If
  If UCase(RegRead(HKEY_CURRENT_USER, "Software\" & App.Title, "Ask Delete Object Database", "Y")) = "Y" Then
    chkAskDeleteObjectDatabase.Value = 1
  Else
    chkAskDeleteObjectDatabase.Value = 0
  End If
  
  'Setup the Font Details
  txtFont.Tag = RegRead(HKEY_CURRENT_USER, "Software\" & App.Title, "Font", "MS Sans Serif|8|False|False")
  szFont = Split(txtFont.Tag, "|")
  cdlg.FontName = szFont(0)
  cdlg.FontSize = Val(szFont(1))
  cdlg.FontBold = CBool(szFont(2))
  cdlg.FontItalic = CBool(szFont(3))
  txtFont.Text = cdlg.FontName & ", " & cdlg.FontSize & "pt"
  If cdlg.FontBold Then txtFont.Text = txtFont.Text & ", bold"
  If cdlg.FontItalic Then txtFont.Text = txtFont.Text & ", italic"
    
  'load the Word List
  LoadWord RegRead(HKEY_CURRENT_USER, "Software\" & App.Title, "AutoHighlight", szDefaultAutoHighlight)

  'Master DB
  txtMasterDB.Text = RegRead(HKEY_CURRENT_USER, "Software\" & App.Title, "Master DB", "template1")
  
  'Encryted Passwords
  If UCase(RegRead(HKEY_CURRENT_USER, "Software\" & App.Title, "Encrypt Passwords", "Y")) = "Y" Then
    chkEncryptPasswords.Value = 1
  Else
    chkEncryptPasswords.Value = 0
  End If
  
  'Auto Row Count
  If UCase(RegRead(HKEY_CURRENT_USER, "Software\" & App.Title, "Auto Row Count", "Y")) = "Y" Then
    chkAutoRowCount.Value = 1
  Else
    chkAutoRowCount.Value = 0
  End If
  
  'Defer Connection
  If UCase(RegRead(HKEY_CURRENT_USER, "Software\" & App.Title, "Defer Connection", "Y")) = "Y" Then
    chkDeferConnection.Value = 1
  Else
    chkDeferConnection.Value = 0
  End If
  
  'max number of sql command in query
  UpDownMaxSqlQuery.Value = Val(RegRead(HKEY_CURRENT_USER, "Software\" & App.Title, "Max Number Sql Query", "50"))
  
  'max number of Record in View Data
  UpDownMaxRecViewData.Value = Val(RegRead(HKEY_CURRENT_USER, "Software\" & App.Title, "Row Limit", "1000"))
  
  '///////////////////////////////////
  'Generate Documentation
  'HTML Help Workshop directory
  txtDocHHWPath.Text = RegRead(HKEY_CURRENT_USER, "Software\" & App.Title & "\GenDbDoc", "HTML Help Workshop directory")
  'Work directory
  txtDocWorkPath.Text = RegRead(HKEY_CURRENT_USER, "Software\" & App.Title & "\GenDbDoc", "Work directory")
  
  'load combo lang
  szCurrentLang = RegRead(HKEY_CURRENT_USER, "Software\" & App.Title, "Current Lang", "English")
  cboLang.Clear
  cboLang.AddItem "English"
  szTemp = Dir(App.Path & "\*.lng")
  While szTemp <> ""
    szTemp = Left(szTemp, Len(szTemp) - 4)
    cboLang.AddItem szTemp
    If szTemp = szCurrentLang Then cboLang.ListIndex = cboLang.NewIndex
    szTemp = Dir
  Wend
  If cboLang.ListCount > 0 And cboLang.ListIndex = -1 Then cboLang.ListIndex = 0
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmOptions.Form_Load"
End Sub

Private Sub cmdColour_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmOptions.cmdColour_Click()", etFullDebug

  cdlg.ShowColor
  txtWord.ForeColor = cdlg.Color

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmOptions.cmdColour_Click"
End Sub

Private Sub GetExporters()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmOptions.GetExporters()", etFullDebug

Dim objExporter As pgExporter

  lstExporters.Clear
  txtExpAuthor.Text = ""
  txtExpVersion.Text = ""
  txtExpDescription.Text = ""
  
  For Each objExporter In exp
    lstExporters.AddItem objExporter.Description
  Next objExporter

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmOptions.GetExporters"
End Sub

Private Sub cmdExpInstall_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmOptions.cmdExpInstall_Click()", etFullDebug

  cdlg.FLAGS = cdlOFNHideReadOnly
  cdlg.Filter = "pgAdmin Exporters (*.dll)|*.dll|All Files (*.*)|*.*"
  cdlg.ShowOpen
  If cdlg.FileName = "" Then
    MsgBox ??TrasLang??("No Exporter selected - operation aborted!"), vbExclamation, ??TrasLang??("Error")
    Exit Sub
  Else
    exp.Install cdlg.FileName
  End If
  GetExporters

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmOptions.cmdExpInstall_Click"
End Sub

Private Sub cmdExpUninstall_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmOptions.cmdExpUninstall_Click()", etFullDebug

  If lstExporters.Text = "" Then
    MsgBox ??TrasLang??("You must select a Exporter to uninstall!"), vbExclamation, ??TrasLang??("Error")
    Exit Sub
  End If
  
  If MsgBox(??TrasLang??("Are you sure you wish to uninstall: ") & lstExporters.Text & "?", vbYesNo + vbQuestion, ??TrasLang??("Confirm")) = vbYes Then
    exp.Uninstall lstExporters.Text
    GetExporters
  End If

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmOptions.cmdExpUninstall_Click"
End Sub

Private Sub lstExporters_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmOptions.lstExporters_Click()", etFullDebug

  txtExpDescription.Text = exp(lstExporters.Text).Description
  txtExpVersion.Text = exp(lstExporters.Text).Version
  txtExpAuthor.Text = exp(lstExporters.Text).Author

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmOptions.lstExporters_Click"
End Sub

Private Sub GetPlugins()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmOptions.GetPlugins()", etFullDebug

Dim objPlugin As pgPlugin

  lstPlugins.Clear
  txtPlgAuthor.Text = ""
  txtPlgVersion.Text = ""
  txtPlgDescription.Text = ""
  
  For Each objPlugin In plg
    lstPlugins.AddItem objPlugin.Description
  Next objPlugin
  
  'Rebuild the Plugins Menu
  BuildPluginsMenu

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmOptions.GetPlugins"
End Sub

Private Sub cmdPlgInstall_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmOptions.cmdPlgInstall_Click()", etFullDebug

  cdlg.FLAGS = cdlOFNHideReadOnly
  cdlg.Filter = "pgAdmin Plugins (*.dll)|*.dll|All Files (*.*)|*.*"
  cdlg.ShowOpen
  If cdlg.FileName = "" Then
    MsgBox ??TrasLang??("No Plugin selected - operation aborted!"), vbExclamation, ??TrasLang??("Error")
    Exit Sub
  Else
    plg.Install cdlg.FileName
  End If
  GetPlugins

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmOptions.cmdPlgInstall_Click"
End Sub

Private Sub cmdPlgUninstall_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmOptions.cmdPlgUninstall_Click()", etFullDebug

  If lstPlugins.Text = "" Then
    MsgBox ??TrasLang??("You must select a Plugin to uninstall!"), vbExclamation, ??TrasLang??("Error")
    Exit Sub
  End If
  
  If MsgBox(??TrasLang??("Are you sure you wish to uninstall: ") & lstPlugins.Text & "?", vbYesNo + vbQuestion, ??TrasLang??("Confirm")) = vbYes Then
    plg.Uninstall lstPlugins.Text
    GetPlugins
  End If

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmOptions.cmdPlgUninstall_Click"
End Sub

Private Sub lstPlugins_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmOptions.lstPlugins_Click()", etFullDebug

  txtPlgDescription.Text = plg(lstPlugins.Text).Description
  txtPlgVersion.Text = plg(lstPlugins.Text).Version
  txtPlgAuthor.Text = plg(lstPlugins.Text).Author

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmOptions.lstPlugins_Click"
End Sub

Private Sub tabOptions_Click(PreviousTab As Integer)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmOptions.tabOptions_Click(" & PreviousTab & ")", etFullDebug

  Select Case tabOptions.Tab
    Case 0
    
    Case 1
    
    Case 2
      If lstExporters.ListCount = 0 Then GetExporters
    Case 3
      If lstPlugins.ListCount = 0 Then GetPlugins
  End Select

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmOptions.tabOptions_Click"
End Sub

Private Sub LoadWord(Words As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmOptions.LoadWord(" & Words & ")", etFullDebug

Dim szStrings() As String
Dim szValues() As String
Dim itmX As ListItem
Dim iLoop As Integer

  'Sort out the Word List
  txtWord.ForeColor = RGB(0, 0, 0)
  lvWords.ColumnHeaders.Add , , "Word", (lvWords.Width / 11) * 5
  lvWords.ColumnHeaders.Add , , "Colour", (lvWords.Width / 11) * 3
  lvWords.ColumnHeaders.Add , , "B", (lvWords.Width / 11)
  lvWords.ColumnHeaders.Add , , "I", (lvWords.Width / 11)
  
  'Load the text colours into the grid.
  lvWords.ListItems.Clear
  szStrings = Split(Words, ";")
  For iLoop = 0 To UBound(szStrings) - 1
    szValues = Split(szStrings(iLoop), "|")
    Set itmX = lvWords.ListItems.Add(, , szValues(0))
    itmX.ForeColor = szValues(3)
    itmX.SubItems(1) = szValues(3)
    If szValues(2) = "1" Then
      itmX.SubItems(3) = "Y"
    Else
      itmX.SubItems(3) = "N"
    End If
    If szValues(1) = "1" Then
      itmX.SubItems(2) = "Y"
    Else
      itmX.SubItems(2) = "N"
    End If
  Next iLoop
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmOptions.LoadWord"
End Sub
