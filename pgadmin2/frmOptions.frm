VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{44F33AC4-8757-4330-B063-18608617F23E}#12.4#0"; "HighlightBox.ocx"
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   6885
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   5520
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   5520
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   0
      Top             =   6390
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab tabOptions 
      Height          =   6360
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   5460
      _ExtentX        =   9631
      _ExtentY        =   11218
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
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
      Tab(1).Control(0)=   "Frame7"
      Tab(1).Control(1)=   "Frame8"
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
            _ExtentX        =   8440
            _ExtentY        =   4789
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
            Caption         =   $"frmOptions.frx":0A8E
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
            Caption         =   $"frmOptions.frx":0B36
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
            Caption         =   $"frmOptions.frx":0BFB
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
            _ExtentX        =   8916
            _ExtentY        =   1667
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
         Height          =   3375
         ItemData        =   "frmOptions.frx":0C8F
         Left            =   -74910
         List            =   "frmOptions.frx":0C91
         TabIndex        =   29
         Top             =   450
         Width           =   5235
      End
      Begin VB.ListBox lstExporters 
         Height          =   3375
         ItemData        =   "frmOptions.frx":0C93
         Left            =   -74910
         List            =   "frmOptions.frx":0C95
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
            _ExtentX        =   8916
            _ExtentY        =   1667
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
            Caption         =   $"frmOptions.frx":0C97
            Height          =   600
            Index           =   0
            Left            =   135
            TabIndex        =   40
            Top             =   630
            Width           =   4695
         End
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
      ScaleWidth      =   5685
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
      ScaleWidth      =   5685
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
      ScaleWidth      =   5685
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
      Left            =   3285
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
    MsgBox "You must enter a word to add!", vbExclamation, "Error"
    txtWord.SetFocus
    Exit Sub
  End If
  For Each itmX In lvWords.ListItems
    If itmX.Text = txtWord.Text Then
      MsgBox "That word is already in the list!", vbExclamation, "Error"
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
    .DialogTitle = "Log File"
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
  cdlg.DialogTitle = "Data Font"
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
  
  Unload Me
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmOptions.cmdOK_Click"
End Sub

Private Sub cmdRemove_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmOptions.cmdRemove_Click()", etFullDebug

  If MsgBox("Are you sure you wish to remove the selected word?", vbQuestion + vbYesNo, "Remove Word") = vbNo Then Exit Sub
  lvWords.ListItems.Remove lvWords.SelectedItem.Index
      
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmOptions.cmdRemove_Click"
End Sub

Private Sub Form_Load()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmOptions.Form_Load()", etFullDebug

Dim iLoop As Integer
Dim itmX As ListItem
Dim szStrings() As String
Dim szValues() As String
Dim szFont() As String

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
    
  'Sort out the Word List
  txtWord.ForeColor = RGB(0, 0, 0)
  lvWords.ColumnHeaders.Add , , "Word", (lvWords.Width / 11) * 5
  lvWords.ColumnHeaders.Add , , "Colour", (lvWords.Width / 11) * 3
  lvWords.ColumnHeaders.Add , , "B", (lvWords.Width / 11)
  lvWords.ColumnHeaders.Add , , "I", (lvWords.Width / 11)
  
  'Load the text colours into the grid.
  lvWords.ListItems.Clear
  szStrings = Split(RegRead(HKEY_CURRENT_USER, "Software\" & App.Title, "AutoHighlight", DEFAULT_AUTOHIGHLIGHT), ";")
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
    MsgBox "No Exporter selected - operation aborted!", vbExclamation, "Error"
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
    MsgBox "You must select a Exporter to uninstall!", vbExclamation, "Error"
    Exit Sub
  End If
  
  If MsgBox("Are you sure you wish to uninstall: " & lstExporters.Text & "?", vbYesNo + vbQuestion, "Confirm") = vbYes Then
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
    MsgBox "No Plugin selected - operation aborted!", vbExclamation, "Error"
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
    MsgBox "You must select a Plugin to uninstall!", vbExclamation, "Error"
    Exit Sub
  End If
  
  If MsgBox("Are you sure you wish to uninstall: " & lstPlugins.Text & "?", vbYesNo + vbQuestion, "Confirm") = vbYes Then
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

