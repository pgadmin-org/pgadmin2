VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{44F33AC4-8757-4330-B063-18608617F23E}#12.4#0"; "HighlightBox.ocx"
Begin VB.Form frmSQLInput 
   Caption         =   "SQL"
   ClientHeight    =   3195
   ClientLeft      =   5070
   ClientTop       =   2955
   ClientWidth     =   7245
   Icon            =   "frmSQLInput.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   7245
   Begin VB.ComboBox cboExporters 
      Height          =   315
      Left            =   4500
      Style           =   2  'Dropdown List
      TabIndex        =   6
      ToolTipText     =   "Select the query target."
      Top             =   2843
      Width           =   2715
   End
   Begin VB.CommandButton cmdExplain 
      Caption         =   "E&xplain"
      Height          =   330
      Left            =   2565
      TabIndex        =   4
      ToolTipText     =   "Execute the SQL query to the selected output option."
      Top             =   2835
      Width           =   810
   End
   Begin HighlightBox.HBX txtSQL 
      Height          =   2805
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Enter an SQL query or statement to execute."
      Top             =   0
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   4948
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ControlBarVisible=   0   'False
   End
   Begin VB.CommandButton cmdSQLWizard 
      Caption         =   "&Wizard"
      Height          =   330
      Left            =   1710
      TabIndex        =   3
      ToolTipText     =   "Run the SQL Wizard."
      Top             =   2835
      Width           =   810
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&Load"
      Height          =   330
      Left            =   0
      TabIndex        =   1
      ToolTipText     =   "Load a query."
      Top             =   2835
      Width           =   810
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   330
      Left            =   855
      TabIndex        =   2
      ToolTipText     =   "Save the current query."
      Top             =   2835
      Width           =   795
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "&Execute to:"
      Height          =   330
      Left            =   3420
      TabIndex        =   5
      ToolTipText     =   "Execute the SQL query to the selected output option."
      Top             =   2835
      Width           =   1035
   End
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Select SQL File"
      Filter          =   "All Files (*.*)|*.*"
   End
   Begin VB.Menu mnuLoadCmd 
      Caption         =   "&Previous"
      Index           =   0
   End
   Begin VB.Menu mnuLoadCmd 
      Caption         =   "&Next"
      Index           =   1
   End
End
Attribute VB_Name = "frmSQLInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence
'
' frmSQLInput.frm - Input Arbitrary SQL

Option Explicit

Dim bDirty As Boolean
Dim szDatabase As String
Const MAXCMDSQL As Integer = 50

Private Sub cmdExecute_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLInput.cmdExecute_Click()", etFullDebug

Dim rsQuery As New Recordset
Dim szBits() As String
Dim vBit As Variant
Dim szSQL As String

  If Len(txtSQL.Text) < 5 Then Exit Sub
  
  If txtSQL.SelLength > 5 Then
    szSQL = Mid(txtSQL.Text, txtSQL.SelStart + 1, txtSQL.SelLength)
  Else
    szSQL = txtSQL.Text
  End If
  
  RegWrite HKEY_CURRENT_USER, "Software\" & App.Title, "Recordset Viewer", regString, cboExporters.Text
  
  StartMsg "Executing SQL Query..."
  Set rsQuery = frmMain.svr.Databases(szDatabase).Execute(szSQL, , , qryUser)
  If rsQuery.Fields.Count > 0 Then
    Select Case cboExporters.Text
      Case "Screen"
        Dim objOutputForm As New frmSQLOutput
        Load objOutputForm
        objOutputForm.Display rsQuery, szDatabase, Me.Tag
        objOutputForm.Show
        EndMsg
      Case Else
        EndMsg
        frmMain.svr.LogEvent "Running Exporter: " & exp(cboExporters.Text).Description & " v" & exp(cboExporters.Text).Version, etMiniDebug
        exp(cboExporters.Text).Export rsQuery
    End Select
  Else
    EndMsg
    MsgBox "Query Executed OK!", vbInformation
  End If
  StoreCmdSql szSQL

  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLInput.cmdExecute_Click"
End Sub

Private Sub cmdExplain_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLInput.cmdExplain_Click()", etFullDebug

Dim objQueryPlanForm As New frmSQLExplain

  'Check for blank query
  If txtSQL.Text = "" Then Exit Sub

  Load objQueryPlanForm
  objQueryPlanForm.Explain txtSQL.Text, szDatabase
  objQueryPlanForm.Show

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLInput.cmdExplain_Click"
End Sub

Private Sub cmdLoad_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLInput.cmdLoad_Click()", etFullDebug

Dim szLine As String
Dim szFile As String
Dim fNum As Integer

  If bDirty = True Then
    If MsgBox("This query has been edited - do you wish to save it?", vbQuestion + vbYesNo, "Save Query") = vbYes Then cmdSave_Click
  End If
  
  With cdlg
    .DialogTitle = "Load SQL Query"
    .FLAGS = cdlOFNFileMustExist + cdlOFNHideReadOnly
    .Filter = "SQL Scripts (*.sql)|*.sql|All Files (*.*)|*.*"
    .FileName = ""
    .CancelError = True
    .ShowOpen
  End With
  
  If cdlg.FileName = "" Then Exit Sub
  txtSQL.Text = ""
  fNum = FreeFile
  frmMain.svr.LogEvent "Loading " & cdlg.FileName, etMiniDebug
  Open cdlg.FileName For Input As #fNum
  While Not EOF(fNum)
    Line Input #fNum, szLine
     szFile = szFile & szLine & vbCrLf
  Wend
  If Len(szFile) > 2 Then szFile = Left(szFile, Len(szFile) - 2)
  
  Close #fNum
  txtSQL.Text = szFile
  Me.Caption = "SQL " & Me.Tag & ": " & szDatabase & " (" & GetFilename & ")"
  bDirty = False

  Exit Sub
Err_Handler:
  If Err.Number = 32755 Then
    frmMain.svr.LogEvent "Load Query operation cancelled.", etMiniDebug
    Exit Sub
  End If
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLInput.cmdLoad_Click"
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLInput.Form_KeyUp(" & KeyCode & ", " & Shift & ")", etFullDebug

  Select Case KeyCode
    Case vbKeyF5
      cmdExecute_Click
  End Select
    
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLInput.Form_KeyUp"
End Sub

Private Sub cmdSave_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLInput.cmdSave_Click()", etFullDebug

Dim fNum As Integer

  With cdlg
    .DialogTitle = "Save SQL Query"
    .Filter = "SQL Scripts (*.sql)|*.sql"
    .CancelError = True
    .ShowSave
  End With
  If cdlg.FileName = "" Then
    MsgBox "No filename specified - SQL query not saved.", vbExclamation, "Warning"
    Exit Sub
  End If
  If Dir(cdlg.FileName) <> "" Then
    If MsgBox("File exists - overwrite?", vbYesNo + vbQuestion, "Overwrite File") = vbNo Then
      cmdSave_Click
      Exit Sub
    End If
  End If
  fNum = FreeFile
  frmMain.svr.LogEvent "Writing " & cdlg.FileName, etMiniDebug
  Open cdlg.FileName For Output As #fNum
  Print #fNum, txtSQL.Text
  Close #fNum
  Me.Caption = "SQL " & Me.Tag & ": " & szDatabase & " (" & GetFilename & ")"
  bDirty = False

  Exit Sub
Err_Handler:
  If Err.Number = 32755 Then
    frmMain.svr.LogEvent "Save Query operation cancelled.", etMiniDebug
    Exit Sub
  End If
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLInput.cmdSave_Click"
End Sub

Private Sub cmdSQLWizard_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLInput.cmdSave_Click()", etFullDebug

Dim objSQLWizardForm As New frmSQLWizard
  Load objSQLWizardForm
  objSQLWizardForm.Tag = Me.hwnd
  objSQLWizardForm.Caption = "SQL Wizard " & Me.Tag & ": " & szDatabase
  objSQLWizardForm.Initialise szDatabase
  objSQLWizardForm.Show

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLInput.cmdSQLWizard_Click"
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLInput.cmdSave_Click()", etFullDebug

Dim X As Integer
Dim objExporter As pgExporter
Dim szExporter As String

  cboExporters.AddItem "Screen"
  
  'Load Exporters
  For Each objExporter In exp
    cboExporters.AddItem objExporter.Description
  Next objExporter

  szExporter = RegRead(HKEY_CURRENT_USER, "Software\" & App.Title, "Recordset Viewer", "Screen")
  For X = 0 To cboExporters.ListCount - 1
    If cboExporters.List(X) = szExporter Then
      cboExporters.ListIndex = X
      Exit For
    End If
  Next X
  
  Set txtSQL.Font = ctx.Font
  txtSQL.Wordlist = ctx.AutoHighlight
  szDatabase = ctx.CurrentDB
  bDirty = False
  Me.Height = 3600
  Me.Width = 6705

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLInput.Form_Load"
End Sub

Private Sub Form_Resize()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLInput.Form_Resize()", etFullDebug

  If Me.WindowState <> 1 And Me.ScaleHeight > 0 Then
    If Me.WindowState = 0 Then
      If Me.Width < 7365 Then Me.Width = 7365
      If Me.Height < 3600 Then Me.Height = 3600
    End If
    
    txtSQL.Width = Me.ScaleWidth
    txtSQL.Height = Me.ScaleHeight - cmdExecute.Height - 50
    cmdExecute.Top = Me.ScaleHeight - cmdExecute.Height
    cmdExplain.Top = cmdExecute.Top
    cmdLoad.Top = cmdExecute.Top
    cmdSave.Top = cmdExecute.Top
    cmdSQLWizard.Top = cmdExecute.Top
    cboExporters.Top = cmdExecute.Top - ((cmdExecute.Height - cboExporters.Height) / 2)
    cboExporters.Left = Me.ScaleWidth - cboExporters.Width
    cmdExecute.Left = cboExporters.Left - cmdExecute.Width - 50
    cboExporters.Left = Me.ScaleWidth - cboExporters.Width

  End If

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLInput.Form_Resize"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLInput.Form_Unload()", etFullDebug

  If bDirty = True Then
    Select Case MsgBox("This query has been edited - do you wish to save it?", vbQuestion + vbYesNoCancel, "Save Query")
      Case vbYes
        cmdSave_Click
      Case vbCancel
        Cancel = 1
        Exit Sub
    End Select
  End If

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLInput.Form_Unload"
End Sub

Private Sub txtSQL_Change()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLInput.txtSQL_Change()", etFullDebug

  Me.Caption = "SQL " & Me.Tag & ": " & szDatabase & " (" & GetFilename & ")*"
  bDirty = True

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLInput.txtSQL_Change"
End Sub

Private Function GetFilename() As String
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLInput.GetFilename()", etFullDebug

Dim szParts() As String
  
  szParts = Split(cdlg.FileName, "\")
  If UBound(szParts) >= 0 Then GetFilename = szParts(UBound(szParts))

  Exit Function
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLInput.GetFilename"
End Function

'Load next/previous command sql
Private Sub mnuLoadCmd_Click(Index As Integer)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLInput.mnuLoadCmd_Click(" & Index & ")", etFullDebug

Dim szCmdSql() As String
Dim ii As Integer
Dim iCmdSql As Integer
Dim szTemp As String

  ReDim szCmdSql(MAXCMDSQL) As String
  iCmdSql = -1
  
  'load data
  For ii = 0 To MAXCMDSQL
    szTemp = RegRead(HKEY_CURRENT_USER, "Software\" & App.Title & "\Command SQL", "CmdSQL" & ii, "")
    If Len(Trim(szTemp)) > 0 Then
      iCmdSql = iCmdSql + 1
      szCmdSql(iCmdSql) = szTemp
    End If
  Next
  If iCmdSql < 0 Then Exit Sub
  ReDim Preserve szCmdSql(iCmdSql) As String
  
  If Len(txtSQL.Text) = 0 Then
    If Index = 1 Then
      'next
      szTemp = szCmdSql(0)
    Else
      szTemp = szCmdSql(UBound(szCmdSql))
    End If
  Else
    szTemp = ""
    For ii = 0 To UBound(szCmdSql)
      If szCmdSql(ii) = txtSQL.Text Then
        If Index = 1 Then
          'next
          If ii < UBound(szCmdSql) Then
            szTemp = szCmdSql(ii + 1)
          Else
            szTemp = szCmdSql(0)
          End If
        Else
          If ii > 0 Then
            szTemp = szCmdSql(ii - 1)
          Else
            szTemp = szCmdSql(UBound(szCmdSql))
          End If
        End If
        Exit For
      End If
    Next
    If Len(szTemp) = 0 Then
      If Index = 1 Then
        'next
        szTemp = szCmdSql(0)
      Else
        szTemp = szCmdSql(UBound(szCmdSql))
      End If
    End If
  End If
  txtSQL.Text = szTemp

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLInput.mnuLoadCmd_Click"
End Sub

'Store command sql
Private Sub StoreCmdSql(szCommandSql As String)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLInput.StoreCmdSql(" & QUOTE & szCommandSql & QUOTE & ")", etFullDebug

Dim ii As Integer
Dim szTemp As String
Dim iPosFree As Integer
  
  iPosFree = -1
  For ii = 0 To MAXCMDSQL
    szTemp = RegRead(HKEY_CURRENT_USER, "Software\" & App.Title & "\Command SQL", "CmdSQL" & ii, "")
    If Len(Trim(szTemp)) = 0 And iPosFree = -1 Then
      iPosFree = ii
    ElseIf szTemp = szCommandSql Then
      Exit Sub
    End If
  Next
  'cycle story
  If iPosFree = -1 Then iPosFree = 0
  
  'store command sql
  RegWrite HKEY_CURRENT_USER, "Software\" & App.Title & "\Command Sql", "CmdSql" & iPosFree, regString, szCommandSql

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLInput.StoreCmdSql"
End Sub
