VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{44F33AC4-8757-4330-B063-18608617F23E}#12.4#0"; "HighlightBox.ocx"
Begin VB.Form frmVisualQueryBuilder 
   Caption         =   "Visual Query Builder"
   ClientHeight    =   6492
   ClientLeft      =   2208
   ClientTop       =   2136
   ClientWidth     =   9120
   Icon            =   "frmVisualQueryBuilder.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6492
   ScaleWidth      =   9120
   Visible         =   0   'False
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   8580
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin HighlightBox.HBX txtSQL 
      Height          =   1272
      Left            =   4200
      TabIndex        =   7
      Top             =   60
      Width           =   4872
      _ExtentX        =   8594
      _ExtentY        =   2244
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "SQL Query"
   End
   Begin VB.Frame fraTypeQuery 
      Caption         =   "Type Query"
      Height          =   1332
      Left            =   2880
      TabIndex        =   2
      Top             =   0
      Width           =   1272
      Begin VB.OptionButton OptTypeQuery 
         Caption         =   "&Insert"
         Enabled         =   0   'False
         Height          =   192
         Index           =   3
         Left            =   120
         TabIndex        =   6
         Top             =   1020
         Width           =   1032
      End
      Begin VB.OptionButton OptTypeQuery 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   192
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   780
         Width           =   1032
      End
      Begin VB.OptionButton OptTypeQuery 
         Caption         =   "&Update"
         Enabled         =   0   'False
         Height          =   192
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   540
         Width           =   1032
      End
      Begin VB.OptionButton OptTypeQuery 
         Caption         =   "&Select"
         Height          =   192
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   300
         Value           =   -1  'True
         Width           =   1032
      End
   End
   Begin MSComctlLib.ImageList il 
      Left            =   1920
      Top             =   0
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisualQueryBuilder.frx":0BC2
            Key             =   "table"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisualQueryBuilder.frx":1294
            Key             =   "namespace"
         EndProperty
      EndProperty
   End
   Begin pgAdmin2.RelationObj RelQuery 
      Height          =   5052
      Left            =   60
      TabIndex        =   0
      Top             =   1380
      Width           =   9012
      _ExtentX        =   17590
      _ExtentY        =   8911
   End
   Begin MSComctlLib.TreeView tv 
      Height          =   1260
      Left            =   60
      TabIndex        =   1
      ToolTipText     =   "Add table to selection"
      Top             =   60
      Width           =   2772
      _ExtentX        =   4890
      _ExtentY        =   2223
      _Version        =   393217
      Indentation     =   441
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "il"
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin VB.Menu mnuRelQuery 
      Caption         =   "Relation Query"
      Visible         =   0   'False
      Begin VB.Menu mnuRelQueryViewSql 
         Caption         =   "View SQL"
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileClearQuery 
         Caption         =   "Clear query"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuFileExecQuery 
         Caption         =   "Execute Query"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuFileRetunQuery 
         Caption         =   "Return Query"
         Shortcut        =   ^K
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
         Shortcut        =   ^Q
      End
   End
End
Attribute VB_Name = "frmVisualQueryBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the Artistic Licence
'
' frmVisualQueryBuilder.frm - Visual Query Builder

Option Explicit

Dim szDB As String
Dim WithEvents FGridVQB As MSFlexGridLib.MSFlexGrid
Attribute FGridVQB.VB_VarHelpID = -1
Dim WithEvents objFGrid As ClsSuperFGrid
Attribute objFGrid.VB_VarHelpID = -1
Dim frmCallingForm As Form

Const HEADER_FILE_VBQ As String = "Visual Query Builder pgAdmin2"
Const TAG_RELATION As String = "[RELATION]"
Const TAG_JOIN As String = "[JOIN]"
Const TAG_COLUMN_FG As String = "[COLUMN FG]"
Const TAG_CURRENT_VERSION As String = "Version 1.0.0"
Const TAG_VERSION_1_0_0 As String = "Version 1.0.0"

Public Sub Initialise(Database As String, frmCF As Form)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmVisualQueryBuilder.Initialise(" & QUOTE & Database & QUOTE & ")", etFullDebug

Dim objNS As pgNamespace
Dim objTable As pgTable
Dim NodeNs As Node
  
  PatchForm Me
  
  szDB = Database
  Set frmCallingForm = frmCF
  txtSql.Wordlist = ctx.AutoHighlight
  
  'load structure
  tv.Nodes.Clear
  For Each objNS In frmMain.svr.Databases(szDB).Namespaces
    If Not (objNS.SystemObject And Not ctx.IncludeSys) Then
      If objNS.Tables.Count(Not ctx.IncludeSys) > 0 Then
        Set NodeNs = tv.Nodes.Add(, , "NSP-" & GetID, objNS.Identifier, "namespace")
        For Each objTable In objNS.Tables
          If Not (objTable.SystemObject And Not ctx.IncludeSys) Then
            tv.Nodes.Add NodeNs.Key, tvwChild, "TBL-" & GetID, objTable.Identifier, "table"
          End If
        Next
      End If
    End If
  Next
  
  With RelQuery
    .MenuActionGrid(1).Caption = §§TrasLang§§("Delete")
    .MenuActionGridEnable = True
  End With
  
  PrepareFGrid
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmVisualQueryBuilder.Initialise"
End Sub

Private Sub PrepareFGrid()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmVisualQueryBuilder.PrepareFGrid()", etFullDebug

Dim ii As Integer
  
  Set FGridVQB = RelQuery.GetGridCompose
  With FGridVQB
    .Visible = True
    .FixedRows = 1
    .FixedCols = 1
    .Rows = 6
    .Cols = 64
    .Height = Me.TextHeight("0") * (.Rows + 4)
    .RowHeight(0) = 100
    .TextMatrix(1, 0) = §§TrasLang§§("Table")
    .TextMatrix(2, 0) = §§TrasLang§§("Column")
    .TextMatrix(3, 0) = §§TrasLang§§("Order")
    .TextMatrix(4, 0) = §§TrasLang§§("Visible")
    .TextMatrix(5, 0) = §§TrasLang§§("Where")
    .HighLight = flexHighlightNever
  End With
  AutoSizeColumnFGrid FGridVQB

  'fix grid using super grid
  If objFGrid Is Nothing Then
    Set objFGrid = New ClsSuperFGrid
    Set objFGrid.FlexGrid = FGridVQB
  End If
  For ii = 1 To FGridVQB.Cols - 1
    With objFGrid
      .AddAction ii, 3, TAFG_COMBO, "|0|1|Ascending|1|0|Discending|2|0"
      .AddAction ii, 4, TAFG_CHECK, "UNCHECKED"
      .AddAction ii, 5, TAFG_TEXT
      With .FlexGrid
        .Col = ii
        .Row = 5
        .CellAlignment = flexAlignLeftCenter
      End With
    End With
  Next
  Exit Sub

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmVisualQueryBuilder.PrepareFGrid"
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmVisualQueryBuilder.Form_KeyUp(" & KeyCode & ", " & Shift & ")", etFullDebug

  Select Case KeyCode
    Case vbKeyF5
      mnuFileExecQuery_Click
    
    Case vbKeyF4
      mnuFileClearQuery_Click
    
  End Select
    
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmVisualQueryBuilder.Form_KeyUp"
End Sub

Private Sub Form_Resize()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmVisualQueryBuilder.Form_Resize()", etFullDebug
  
  If Me.WindowState <> vbMinimized Then
    tv.Left = RelQuery.Left
    RelQuery.Width = Me.ScaleWidth - 100
    
    If txtSql.Top = 0 Then
      txtSql.Width = Me.ScaleWidth
      txtSql.Height = Me.ScaleHeight
    Else
      If Me.ScaleWidth - txtSql.Left > 0 Then txtSql.Width = Me.ScaleWidth - txtSql.Left - 50
    End If
    RelQuery.Height = Me.ScaleHeight - tv.Top - tv.Height - 100
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmVisualQueryBuilder.Form_Resize"
End Sub

Private Sub mnuFileClearQuery_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmVisualQueryBuilder.mnuFileClearQuery_Click()", etFullDebug

  If MsgBox(§§TrasLang§§("Do you wish clear Query?"), vbQuestion + vbYesNo, §§TrasLang§§("Clear Query")) = vbNo Then Exit Sub
  RelQuery.Clear
  txtSql.Text = ""
  PrepareFGrid

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmVisualQueryBuilder.mnuFileClearQuery_Click"
End Sub

Private Sub mnuFileExecQuery_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmVisualQueryBuilder.mnuFileExecQuery_Click()", etFullDebug

Dim rsQuery As New Recordset
Dim szSQL As String

  mnuRelQueryViewSql_Click
  If Len(txtSql.Text) < 5 Then Exit Sub
  
  If txtSql.SelLength > 5 Then
    szSQL = Mid(txtSql.Text, txtSql.SelStart + 1, txtSql.SelLength)
  Else
    szSQL = txtSql.Text
  End If
  
  StartMsg §§TrasLang§§("Executing SQL Query...")
  
  'change CRLF -> LF
  szSQL = Replace(szSQL, vbCrLf, vbLf)
  Set rsQuery = frmMain.svr.Databases(szDB).Execute(szSQL, , , qryUser)
  If rsQuery.Fields.Count > 0 Then
    Dim objOutputForm As New frmSQLOutput
    Load objOutputForm
    objOutputForm.Display rsQuery, szDB, Me.Tag
    objOutputForm.Show
  End If
  EndMsg

  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmVisualQueryBuilder.mnuFileExecQuery_Click", False
End Sub

Private Sub mnuFileExit_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmVisualQueryBuilder.mnuFileExit_Click()", etFullDebug

  Unload Me
  Exit Sub

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmVisualQueryBuilder.mnuFileExit_Click"
End Sub

Private Sub mnuFileRetunQuery_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmVisualQueryBuilder.mnuFileRetunQuery_Click()", etFullDebug

  If Not frmCallingForm Is Nothing Then
    If Not frmCallingForm.Visible Then
      MsgBox §§TrasLang§§("The form that called this form has been destroyed!"), vbExclamation, §§TrasLang§§("Error")
      Exit Sub
    End If
  End If
  
  mnuRelQueryViewSql_Click
  frmCallingForm.txtSql.Text = txtSql.Text
  frmCallingForm.ZOrder
  Exit Sub

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmVisualQueryBuilder.mnuFileRetunQuery_Click"
End Sub

'load query definiton
Private Sub mnuFileOpen_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmVisualQueryBuilder.mnuFileOpen_Click()", etFullDebug

Dim iFile As Integer
Dim vData, vData1
Dim szTemp As String
Dim ii As Integer

  'clear
  If RelQuery.GetRelation.Count > 0 Then mnuFileClearQuery_Click
  If RelQuery.GetRelation.Count > 0 Then Exit Sub

  'load file
  With cdlg
    .FileName = "file.vbq"
    .DialogTitle = §§TrasLang§§("Visual Query Builder Open File")
    .Filter = "Visual Query Builder File|*.vbq"
    .FLAGS = &H4
    .CancelError = True
    .ShowOpen
  End With
  
  If Dir(cdlg.FileName) = "" Then Exit Sub
  
  iFile = FreeFile
  Open cdlg.FileName For Input As #iFile
  vData = Split(Input(LOF(iFile), #iFile), vbCrLf)
  Close #iFile
  
  If UBound(vData) > 0 Then
    If vData(0) <> HEADER_FILE_VBQ Then
      MsgBox §§TrasLang§§("This file is not a ") & HEADER_FILE_VBQ, vbExclamation, §§TrasLang§§("Error")
      Exit Sub
    End If
    
    szTemp = §§TrasLang§§("Do you wish load this file?") & vbCrLf
    szTemp = szTemp & String(30, "=") & vbCrLf
    For ii = 0 To 4
      szTemp = szTemp & vData(ii) & vbCrLf
    Next
    If MsgBox(szTemp, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    Select Case vData(1)
      Case TAG_VERSION_1_0_0
        
        'read structure file
        For ii = 0 To UBound(vData)
          Select Case vData(ii)
            Case TAG_RELATION
              vData1 = Split(vData(ii + 2), ".")
            
              'verify if schema Exists
              If Not frmMain.svr.Databases(szDB).Namespaces.Exists(CStr(vData1(0))) Then
                MsgBox §§TrasLang§§("Schema '") & vData1(0) & §§TrasLang§§("' not exists in this database!"), vbCritical, §§TrasLang§§("Error")
                Exit Sub
              End If
            
              'verify if table Exists
              If Not frmMain.svr.Databases(szDB).Namespaces(CStr(vData1(0))).Tables.Exists(CStr(vData1(1))) Then
                MsgBox §§TrasLang§§("Table '") & vData1(0) & §§TrasLang§§("' not exists in this schema!"), vbCritical, §§TrasLang§§("Error")
                Exit Sub
              End If
    
              'add relation
              AddRelation CStr(vData1(0)), CStr(vData1(1)), CStr(vData(ii + 1))
              ii = ii + 2
              
            Case TAG_JOIN
              RelQuery.AddJoin CStr(vData(ii + 1)), CStr(vData(ii + 2)), CStr(vData(ii + 3)), CStr(vData(ii + 4))
              ii = ii + 4
          
            Case TAG_COLUMN_FG
              RelQuery_AddElementInGridCompose CInt(vData(ii + 1)), CStr(vData(ii + 2)), "", CStr(vData(ii + 3))
              objFGrid.SetCurrentSetting CInt(vData(ii + 1)), 3, TAFG_COMBO, CStr(vData(ii + 4))   'order
              objFGrid.SetCheckBoxes CInt(vData(ii + 1)), 4, CBool(vData(ii + 5))            'visible
              FGridVQB.TextMatrix(5, CInt(vData(ii + 1))) = vData(ii + 6)                   'where
          
          End Select
        Next
    
      Case Else
        MsgBox §§TrasLang§§("Version file not valid!"), vbCritical, §§TrasLang§§("Error")
        Exit Sub
        
    End Select
  End If
  
  'view sql command
  mnuRelQueryViewSql_Click
  
  Exit Sub

Err_Handler:
  If Err.Number = 32755 Then
    frmMain.svr.LogEvent "Open Visual Query Builder operation cancelled.", etMiniDebug
    Exit Sub
  End If
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmVisualQueryBuilder.mnuFileOpen_Click"
End Sub

'save query definiton
Private Sub mnuFileSave_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmVisualQueryBuilder.mnuFileSave_Click()", etFullDebug

Dim szTemp As String
Dim ii As Integer
Dim iFile As Integer
Dim szTitle As String
Dim colRel As Collection
Dim vData

  Set colRel = RelQuery.GetRelation
  If colRel.Count = 0 Then
    MsgBox §§TrasLang§§("Is not present relation!"), vbInformation, §§TrasLang§§("Error")
    Exit Sub
  End If

  With cdlg
    .FileName = "file.vbq"
    .DialogTitle = §§TrasLang§§("Visual Query Builder Save File")
    .Filter = "Visual Query Builder File|*.vbq"
    .FLAGS = &H4
    .CancelError = True
    .ShowSave
  End With
  
  szTitle = InputBox(§§TrasLang§§("Insert title query"))
  
  szTemp = HEADER_FILE_VBQ & vbCrLf
  szTemp = szTemp & TAG_CURRENT_VERSION & vbCrLf
  szTemp = szTemp & "Title: " & szTitle & vbCrLf
  szTemp = szTemp & "Date: " & Now & vbCrLf
  szTemp = szTemp & "Database: " & szDB & vbCrLf
  
  'get relation
  For Each vData In colRel
    vData = Split(vData, ",")
    szTemp = szTemp & TAG_RELATION & vbCrLf
    szTemp = szTemp & vData(0) & vbCrLf                           'name realtion
    szTemp = szTemp & vData(1) & vbCrLf                           'tag relation
  Next

  'get join relation
  vData = Split(RelQuery.GetJoinRelation, "|")
  For ii = 0 To ((UBound(vData) + 1) / 6) - 1
    szTemp = szTemp & TAG_JOIN & vbCrLf
    szTemp = szTemp & vData(ii * 3 + ii * 3) & vbCrLf                  'name realtion
    szTemp = szTemp & vData(ii * 3 + 2 + ii * 3) & vbCrLf              'column name
    szTemp = szTemp & vData(ii * 3 + 3 + ii * 3) & vbCrLf              'name realtion
    szTemp = szTemp & vData(ii * 3 + 5 + ii * 3) & vbCrLf              'column name
  Next
  
  'read flex grid
  For ii = 1 To FGridVQB.Cols - 1
    If Len(FGridVQB.TextMatrix(1, ii)) > 0 Then
      szTemp = szTemp & TAG_COLUMN_FG & vbCrLf
      szTemp = szTemp & ii & vbCrLf                                             'column number
      szTemp = szTemp & FGridVQB.TextMatrix(1, ii) & vbCrLf                     'relation
      szTemp = szTemp & FGridVQB.TextMatrix(2, ii) & vbCrLf                     'column
      szTemp = szTemp & objFGrid.GetCurrentSetting(ii, 3, TAFG_COMBO) & vbCrLf  'order
      szTemp = szTemp & CInt(objFGrid.IsChecked(ii, 4)) & vbCrLf                'visible
      szTemp = szTemp & FGridVQB.TextMatrix(5, ii) & vbCrLf                     'where
    End If
  Next
  
  'save file
  iFile = FreeFile
  Open cdlg.FileName For Output As #iFile
  Print #iFile, szTemp
  Close #iFile
  
  Exit Sub

Err_Handler:
  If Err.Number = 32755 Then
    frmMain.svr.LogEvent "Save Visual Query Builder operation cancelled.", etMiniDebug
    Exit Sub
  End If
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmVisualQueryBuilder.mnuFileSave_Click"
End Sub

Private Sub RelQuery_MenuActionGridCompose(Index As Integer, Col As Integer)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmVisualQueryBuilder.RelQuery_MenuActionGridCompose(" & Index & "," & Col & ")", etFullDebug

Dim iCol As Integer

  Select Case Index
    
    Case 1    'remove action
      objFGrid.RemoveColumn FGridVQB.ColSel
      For iCol = 1 To FGridVQB.Cols - 1
        If Len(FGridVQB.TextMatrix(1, iCol)) = 0 Then
          objFGrid.RemoveAction iCol, 1, TAFG_COMBO
          objFGrid.RemoveAction iCol, 2, TAFG_COMBO
        End If
      Next
  
  End Select
  Exit Sub

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmVisualQueryBuilder.RelQuery_RemoveElementinGridCompose"
End Sub

Private Sub RelQuery_RenameRelation(OldName As String, NewName As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmVisualQueryBuilder.RelQuery_RenameRelation(" & QUOTE & OldName & QUOTE & "," & QUOTE & NewName & QUOTE & ")", etFullDebug

Dim iCol As Integer
  
  'change name in flexgrid
  For iCol = 1 To FGridVQB.Cols - 1
    If FGridVQB.TextMatrix(1, iCol) = OldName Then
      FGridVQB.TextMatrix(1, iCol) = NewName
    End If
  Next
  
  RebildComboTable
  AutoSizeColumnFGrid FGridVQB
  Exit Sub

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmVisualQueryBuilder.RelQuery_RenameRelation"
End Sub

'load table relation
Private Sub tv_DblClick()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmVisualQueryBuilder.tv_DblClick()", etFullDebug

  If tv.SelectedItem Is Nothing Then Exit Sub
  If Left(tv.SelectedItem.Key, 3) <> "TBL" Then Exit Sub
  
  AddRelation tv.SelectedItem.Parent.Text, tv.SelectedItem.Text, tv.SelectedItem.Text
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmVisualQueryBuilder.tv_DblClick"
End Sub

'add relation in query builder
Private Sub AddRelation(Schema As String, Table As String, RelationName As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmVisualQueryBuilder.AddRelation(" & QUOTE & Schema & QUOTE & "," & QUOTE & Table & QUOTE & ")", etFullDebug

Dim szColumn As String
Dim objColumn As pgColumn
  
  'load column table
  szColumn = ""
  For Each objColumn In frmMain.svr.Databases(szDB).Namespaces(Schema).Tables(Table).Columns
    If Not (objColumn.SystemObject And Not ctx.IncludeSys) Then
      If Len(szColumn) > 0 Then szColumn = szColumn & "|"
      szColumn = szColumn & objColumn.Name
    End If
  Next
  
  'add new table in relation
  RelQuery.AddElement RelationName, Schema & "." & Table, §§TrasLang§§("Schema: ") & Schema & §§TrasLang§§("  Table: ") & Table, Split(szColumn, "|")

  RebildComboTable
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmVisualQueryBuilder.tv_DblClick"
End Sub

'rebuild any combo table in flexgrid
Private Sub RebildComboTable()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmVisualQueryBuilder.RebildComboTable()", etFullDebug

Dim iCol As Integer
Dim colRel As Collection
Dim szTemp As String
Dim vData
  
  'create string table
  Set colRel = RelQuery.GetRelation
  szTemp = ""
  For Each vData In colRel
    vData = Split(vData, ",")
    If Len(szTemp) > 0 Then szTemp = szTemp & "|"
    szTemp = szTemp & vData(0) & "|0|0"
  Next
  
  'remove and add combo table
  For iCol = 1 To FGridVQB.Cols - 1
    If Len(FGridVQB.TextMatrix(1, iCol)) > 0 Then
      objFGrid.RemoveAction iCol, 1, TAFG_COMBO
      objFGrid.AddAction iCol, 1, TAFG_COMBO, szTemp
    End If
  Next
  Exit Sub

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmVisualQueryBuilder.RebildComboTable"
End Sub

Private Sub RelQuery_AddElementInGridCompose(Col As Integer, Name As String, Tag As String, Element As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmVisualQueryBuilder.tv_DblClick(" & Col & "," & QUOTE & Name & QUOTE & "," & QUOTE & Tag & QUOTE & "," & QUOTE & Element & QUOTE & ")", etFullDebug
  
Dim objTable As pgTable
Dim szTemp As String
Dim vData
Dim ii As Integer
Dim colRel As Collection
  
  If Col <= 0 Then Exit Sub
  With FGridVQB
    If Len(.TextMatrix(1, Col)) > 0 Then
      'add new column
      objFGrid.AddColumn Col
    End If
    .TextMatrix(1, Col) = Name
    .TextMatrix(2, Col) = Element
  End With
  objFGrid.SetCheckBoxes Col, 4, True

  LoadComboColumn Name, Col

  'load combo table
  Set colRel = RelQuery.GetRelation
  szTemp = ""
  For Each vData In colRel
    vData = Split(vData, ",")
    If Len(szTemp) > 0 Then szTemp = szTemp & "|"
    szTemp = szTemp & vData(0) & "|0|0"
  Next
  objFGrid.RemoveAction Col, 1, TAFG_COMBO
  objFGrid.AddAction Col, 1, TAFG_COMBO, szTemp
  AutoSizeColumnFGrid FGridVQB
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmVisualQueryBuilder.RelQuery_AddElementInGridCompose"
End Sub

Private Sub LoadComboColumn(Name As String, Col As Integer)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmVisualQueryBuilder.LoadComboColumn(" & QUOTE & Name & QUOTE & "," & Col & ")", etFullDebug

Dim objColumn As pgColumn
Dim szTemp As String
Dim bFound As Boolean
Dim szNamespace As String
Dim szTable As String
Dim colRel As Collection
Dim iCol As Integer
Dim vData

  'create string table
  Set colRel = RelQuery.GetRelation
  szTemp = ""
  For Each vData In colRel
    vData = Split(vData, ",")
    If vData(0) = Name Then
      vData = Split(vData(1), ".")
      szNamespace = vData(0)
      szTable = vData(1)
      Exit For
    End If
    If Len(szTemp) > 0 Then szTemp = szTemp & "|"
    szTemp = szTemp & vData(0) & "|0|0"
  Next

  'load combo column
  bFound = False
  szTemp = "*|0|0"
  For Each objColumn In frmMain.svr.Databases(szDB).Namespaces(szNamespace).Tables(szTable).Columns
    If Not (objColumn.SystemObject And Not ctx.IncludeSys) Then
      szTemp = szTemp & "|" & objColumn.Name & "|0|0"
      If objColumn.Name = FGridVQB.TextMatrix(2, Col) Then bFound = True
    End If
  Next

  'if not found column insert default * (all)
  If Not bFound Then FGridVQB.TextMatrix(2, Col) = "*"

  objFGrid.RemoveAction Col, 2, TAFG_COMBO
  objFGrid.AddAction Col, 2, TAFG_COMBO, szTemp
  Exit Sub

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmVisualQueryBuilder.LoadComboColumn"
End Sub

Private Sub objFGrid_ComboChange(Col As Integer, Row As Integer)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmVisualQueryBuilder.objFGrid_ComboChange(" & Col & "," & Row & ")", etFullDebug

  If Row = 1 Then
    'change table reload column
    LoadComboColumn FGridVQB.TextMatrix(1, Col), Col
  End If
  AutoSizeColumnFGrid FGridVQB

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmVisualQueryBuilder.objFGrid_ComboChange"
End Sub

'remove relation from flexgrid
Private Sub RelQuery_RemoveRelation(Name As String, Tag As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmVisualQueryBuilder.RelQuery_RemoveRelation(" & QUOTE & Name & QUOTE & "," & QUOTE & Tag & QUOTE & ")", etFullDebug

Dim iCol As Integer
Dim vData

  vData = Split(Tag, ".")
  For iCol = 1 To FGridVQB.Cols - 1
    If FGridVQB.TextMatrix(1, iCol) = Name Then
      RelQuery_MenuActionGridCompose 1, iCol
    End If
  Next
  
  RebildComboTable
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmVisualQueryBuilder.objFGrid_ComboChange"
End Sub

Private Sub RelQuery_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmVisualQueryBuilder.RelQuery_MouseUp(" & Button & "," & Shift & "," & X & "," & Y & ")", etFullDebug

  If Button = vbRightButton Then
    PopupMenu mnuRelQuery
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmVisualQueryBuilder.RelQuery_MouseUp"
End Sub

Private Sub mnuRelQueryViewSql_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmVisualQueryBuilder.mnuRelQueryViewSql_Click()", etFullDebug

Dim vData, vData1
Dim ii As Integer
Dim szWhere As String
Dim szWhereFlexGrid As String
Dim szFrom As String
Dim szSelect As String
Dim szOrder As String
Dim colForm As New Collection
Dim iCol As Integer
Dim szSQL As String
    
  'create select/order/where
  szSelect = ""
  szOrder = ""
  szWhereFlexGrid = ""
  For iCol = 1 To FGridVQB.Cols - 1
    If Len(FGridVQB.TextMatrix(1, iCol)) > 0 Then
      'select
      If objFGrid.IsChecked(iCol, 4) Then
        If Len(szSelect) > 0 Then szSelect = szSelect & ", "
        szSelect = szSelect & FGridVQB.TextMatrix(1, iCol) & "." & FGridVQB.TextMatrix(2, iCol)
      End If
      
      'order
      If Len(FGridVQB.TextMatrix(3, iCol)) > 0 Then
        If Len(szOrder) > 0 Then szOrder = szOrder & ", "
        szOrder = szOrder & FGridVQB.TextMatrix(1, iCol) & "." & FGridVQB.TextMatrix(2, iCol)
        Select Case objFGrid.GetCurrentSetting(iCol, 3, TAFG_COMBO)
          Case 1
            szOrder = szOrder & " ASC"
          
          Case 2
            szOrder = szOrder & " DESC"
          
        End Select
      End If
      
      'where
      If Len(FGridVQB.TextMatrix(5, iCol)) > 0 Then
        If Len(szWhereFlexGrid) > 0 Then szWhereFlexGrid = szWhereFlexGrid & " AND "
        szWhereFlexGrid = szWhereFlexGrid & "(" & FGridVQB.TextMatrix(5, iCol) & ")"
      End If
    End If
  Next
    
  'create from
  szFrom = ""
  Set colForm = RelQuery.GetRelation
  For Each vData In colForm
    If Len(szFrom) > 0 Then szFrom = szFrom & ", "
    vData1 = Split(vData, ",")
    szFrom = szFrom & " " & vData1(1) & " AS " & vData1(0)
  Next
  
  'create where
  szWhere = ""
  vData = Split(RelQuery.GetJoinRelation, "|")
  For ii = 0 To ((UBound(vData) + 1) / 6) - 1
    If Len(szWhere) > 0 Then szWhere = szWhere & vbCrLf & " AND "
    szWhere = szWhere & vData(ii * 3 + ii * 3) & "." & vData(ii * 3 + 2 + ii * 3)
    szWhere = szWhere & " = "
    szWhere = szWhere & vData(ii * 3 + 3 + ii * 3) & "." & vData(ii * 3 + 5 + ii * 3)
  Next
    
  'create sql command
  If Len(szSelect) > 0 Then szSQL = "SELECT " & szSelect & vbCrLf
  If Len(szFrom) > 0 Then szSQL = szSQL & "FROM " & szFrom & vbCrLf
  If Len(szWhere) > 0 Then szSQL = szSQL & "WHERE (" & szWhere & ")" & vbCrLf
  If Len(szWhereFlexGrid) > 0 Then
    szSQL = szSQL & IIf(Len(szWhere) > 0, " AND ", " WHERE ")
    szSQL = szSQL & "(" & szWhereFlexGrid & ")" & vbCrLf
  End If
  If Len(szOrder) > 0 Then szSQL = szSQL & "ORDER BY " & szOrder & vbCrLf
  
  txtSql.Text = szSQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmVisualQueryBuilder.mnuRelQueryViewSql_Click"
End Sub

Private Sub Form_Paint()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmVisualQueryBuilder.Form_Paint()", etFullDebug

  RelQuery.Refresh
  Exit Sub

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmVisualQueryBuilder.Form_Paint"
End Sub

Private Sub Form_Unload(Cancel As Integer)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmVisualQueryBuilder.Form_Unload()", etFullDebug

  If RelQuery.GetRelation.Count > 0 Then
    Select Case MsgBox(§§TrasLang§§("This query has been edited - do you wish to save it?"), vbQuestion + vbYesNoCancel, §§TrasLang§§("Save Query"))
      Case vbYes
        mnuFileSave_Click
      
      Case vbCancel
        Cancel = 1
        Exit Sub
    
    End Select
  End If

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmVisualQueryBuilder.Form_Unload"
End Sub

