VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsSuperFGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the Artistic Licence
'
' ClsSuperFGrid.cls - Add new functionality a standard flgrid

Option Explicit

Public Enum TypeActionFlexGrid
  TAFG_TEXT = 0
  TAFG_CHECK = 1
  TAFG_COMBO = 2
  TAFG_NULL = 3
End Enum

Private Type DetailAction
  Col As Integer
  Row As Integer
  Valid As Boolean
  TypeA As TypeActionFlexGrid
  Data As String
End Type

Dim objAction() As DetailAction
Dim WithEvents objFGrid As MSFlexGrid
Attribute objFGrid.VB_VarHelpID = -1
Dim WithEvents objTextDummy As VB.TextBox
Attribute objTextDummy.VB_VarHelpID = -1
Dim WithEvents objComboDummy As VB.ComboBox
Attribute objComboDummy.VB_VarHelpID = -1
Event ComboChange(Col As Integer, Row As Integer)

Private Sub Class_Initialize()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":ClsSuperFGrid.Class_Initialize()", etFullDebug
  
  ReDim objAction(0) As DetailAction
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":ClsSuperFGrid.Class_Initialize"
End Sub

'Object Flex Grid
Public Property Set FlexGrid(FGrid As MSFlexGrid)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":ClsSuperFGrid.Property Set FlexGrid(" & QUOTE & FGrid.Name & QUOTE & ")", etFullDebug
  
  Set objFGrid = FGrid
  
  'add combo dummy
  Set objComboDummy = objFGrid.Parent.Controls.Add("VB.ComboBox", "CboDummyFGrid")
  
  'add text dummy
  Set objTextDummy = objFGrid.Parent.Controls.Add("VB.TextBox", "TxtDummyFGrid")
  objTextDummy.Height = 1
  
  Exit Property
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":clsPlugins.Property Set FlexGrid"
End Property
Public Property Get FlexGrid() As MSFlexGrid
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":ClsSuperFGrid.Property Get FlexGrid()", etFullDebug
  
  Set FlexGrid = objFGrid

  Exit Property
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":clsPlugins.Property Get FlexGrid"
End Property

'add action
Public Sub AddAction(Col As Integer, Row As Integer, TypeA As TypeActionFlexGrid, Optional Data As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":ClsSuperFGrid.AddAction(" & Col & "," & Row & "," & TypeA & ")", etFullDebug

Dim ii As Integer
Dim bData As Boolean

  If ActionPosition(Col, Row, TypeA) = 0 Then
    ii = UBound(objAction) + 1
    ReDim Preserve objAction(ii) As DetailAction
    With objAction(ii)
      .Col = Col
      .Row = Row
      .Valid = True
      .TypeA = TypeA
      
      Select Case .TypeA
        Case TAFG_CHECK
          bData = False
          If Not IsEmpty(Data) Then bData = (UCase(Data) = "CHECKED")
          SetCheckBoxes .Col, .Row, bData
        
        Case TAFG_COMBO
          .Data = Data
          
          'fix row Height
          If objFGrid.RowHeight(.Row) < objComboDummy.Height Then objFGrid.RowHeight(.Row) = objComboDummy.Height
      
        Case TAFG_TEXT
          'fix row Height
          If objFGrid.RowHeight(.Row) < objTextDummy.Height Then objFGrid.RowHeight(.Row) = objTextDummy.Height
      
      End Select
    End With
  End If

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":ClsSuperFGrid.AddAction"
End Sub
'remove action
Public Sub RemoveAction(Col As Integer, Row As Integer, TypeA As TypeActionFlexGrid)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":ClsSuperFGrid.RemoveAction(" & Col & "," & Row & "," & TypeA & ")", etFullDebug
  
  objAction(ActionPosition(Col, Row, TypeA)).Valid = False

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":ClsSuperFGrid.RemoveAction"
End Sub

'retun position on action
Private Function ActionPosition(Col As Integer, Row As Integer, TypeA As TypeActionFlexGrid) As Integer
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":ClsSuperFGrid.ActionPosition(" & Col & "," & Row & "," & TypeA & ")", etFullDebug

Dim ii As Integer

  ActionPosition = 0
  For ii = 1 To UBound(objAction)
    With objAction(ii)
      If .Col = Col And .Row = Row And .Valid = True And .TypeA = TypeA Then
        ActionPosition = ii
        Exit For
      End If
    End With
  Next
  
  Exit Function
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":ClsSuperFGrid.ActionPosition"
End Function

'create CheckBoxes
Public Property Get IsChecked(ByVal Col As Integer, ByVal Row As Integer) As Boolean
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":ClsSuperFGrid.IsChecked(" & Col & "," & Row & ")", etFullDebug

Dim iRow As Integer
Dim iCol As Integer
  
  If ActionPosition(Col, Row, TAFG_CHECK) <= 0 Then Exit Property
  With objFGrid
    iRow = .Row
    iCol = .Col
    .Row = Row
    .Col = Col
    IsChecked = (.CellPicture = frmDummy.imgChecked.Picture)
    .Row = iRow
    .Col = iCol
  End With
  Exit Property
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":ClsSuperFGrid.IsChecked"
End Property

'create CheckBoxes
Public Sub SetCheckBoxes(ByVal Col As Integer, ByVal Row As Integer, Checked As Boolean)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":ClsSuperFGrid.SetCheckBoxes(" & Col & "," & Row & "," & Checked & ")", etFullDebug

Dim iRow As Integer
Dim iCol As Integer
  
  If ActionPosition(Col, Row, TAFG_CHECK) <= 0 Then Exit Sub
  With objFGrid
    iRow = .Row
    iCol = .Col
    .Row = Row
    .Col = Col
    .CellPictureAlignment = 4 'center
    If Checked Then
      Set .CellPicture = frmDummy.imgChecked.Picture
    Else
      Set .CellPicture = frmDummy.imgUnchecked.Picture
    End If
    .CellForeColor = vbWhite
    .Row = iRow
    .Col = iCol
  End With
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":ClsSuperFGrid.SetCheckBoxes"
End Sub

'retun type action
Private Function GetTypeAction(Col As Integer, Row As Integer) As TypeActionFlexGrid
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":ClsSuperFGrid.GetTypeAction(" & Col & "," & Row & ")", etFullDebug

Dim ii As Integer

  GetTypeAction = TAFG_NULL
  For ii = 1 To UBound(objAction)
    With objAction(ii)
      If .Col = Col And .Row = Row And .Valid = True Then
        GetTypeAction = .TypeA
        Exit For
      End If
    End With
  Next
  
  Exit Function
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":ClsSuperFGrid.GetTypeAction"
End Function

Private Sub objComboDummy_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":ClsSuperFGrid.objComboDummy_Click()", etFullDebug
  
  If objComboDummy.Visible = True Then
    objFGrid.Text = objComboDummy.Text
    objComboDummy.Text = ""
    objComboDummy.Visible = False
    RaiseEvent ComboChange(objFGrid.Col, objFGrid.Row)
  End If

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":ClsSuperFGrid.objComboDummy_Click"
End Sub

Private Sub objComboDummy_KeyUp(KeyCode As Integer, Shift As Integer)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":ClsSuperFGrid.objComboDummy_KeyUp(" & KeyCode & "," & Shift & ")", etFullDebug
  
  Select Case KeyCode
    Case vbKeyEscape
      objComboDummy.Visible = False
  End Select

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":ClsSuperFGrid.objComboDummy_KeyUp"
End Sub

Private Sub objComboDummy_LostFocus()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":ClsSuperFGrid.objComboDummy_LostFocus()", etFullDebug
  
  objComboDummy.Visible = False

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":ClsSuperFGrid.objComboDummy_LostFocus"
End Sub

Private Sub objFGrid_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":ClsSuperFGrid.objFGrid_MouseUp(" & Button & "," & Shift & "," & x & "," & y & ")", etFullDebug

Dim iRow As Integer
Dim iCol As Integer
Dim vData
Dim ii As Integer
    
  If Button = vbRightButton Then Exit Sub
    
  With objFGrid
    Select Case GetTypeAction(.Col, .Row)
      Case TAFG_NULL
        Exit Sub

      Case TAFG_CHECK
        Select Case .CellPicture
          Case frmDummy.imgUnchecked.Picture
            Set .CellPicture = frmDummy.imgChecked.Picture
          Case frmDummy.imgChecked.Picture
            Set .CellPicture = frmDummy.imgUnchecked.Picture
        End Select
    
      Case TAFG_COMBO
        If .MouseCol > 0 And .MouseRow > 0 Then
          objComboDummy.Clear
          vData = Split(objAction(ActionPosition(.Col, .Row, TAFG_COMBO)).Data, "|")
          For ii = 0 To Int(UBound(vData) / 3)
            objComboDummy.AddItem vData(ii * 3), vData(ii * 3 + 1)
            If vData(ii * 3 + 2) = "1" Then objComboDummy.ListIndex = objComboDummy.NewIndex
          Next
          objComboDummy.BackColor = .BackColor
          objComboDummy.ForeColor = .ForeColor
          Set objComboDummy.Font = .Font
          objComboDummy.Width = .CellWidth
          objComboDummy.Left = .CellLeft + .Left
          objComboDummy.Top = .CellTop + .Top
          objComboDummy.Text = .Text
          objComboDummy.Visible = True
          objComboDummy.ZOrder
          objComboDummy.SetFocus
          objComboDummy.SelStart = 0
          objComboDummy.SelLength = Len(objComboDummy.Text)
        End If
    
      Case TAFG_TEXT
        If .MouseCol > 0 And .MouseRow > 0 Then
          objTextDummy.BackColor = .BackColor
          objTextDummy.ForeColor = .ForeColor
          Set objTextDummy.Font = .Font
          objTextDummy.Width = .CellWidth
          objTextDummy.Height = .CellHeight
          objTextDummy.Left = .CellLeft + .Left
          objTextDummy.Top = .CellTop + .Top
          objTextDummy.Text = .Text
          objTextDummy.Visible = True
          objTextDummy.SetFocus
          objTextDummy.SelStart = 0
          objTextDummy.SelLength = Len(objTextDummy.Text)
          objTextDummy.ZOrder
        End If
    End Select
  End With
  Exit Sub

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":ClsSuperFGrid.objFGrid_MouseUp"
End Sub

Private Sub objTextDummy_KeyUp(KeyCode As Integer, Shift As Integer)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":ClsSuperFGrid.objTextDummy_KeyUp(" & KeyCode & "," & Shift & ")", etFullDebug
  
  Select Case KeyCode
    Case vbKeyReturn
      objFGrid.Text = objTextDummy.Text
      objTextDummy.Visible = False
    
    Case vbKeyEscape
      objTextDummy.Visible = False

  End Select
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":ClsSuperFGrid.objTextDummy_KeyUp"
End Sub

Private Sub objTextDummy_LostFocus()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":ClsSuperFGrid.objTextDummy_LostFocus()", etFullDebug
  
  objTextDummy.Visible = False
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":ClsSuperFGrid.objTextDummy_LostFocus"
End Sub

'add new column
Public Sub AddColumn(Col As Integer)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":ClsSuperFGrid.AddColumn(" & Col & ")", etFullDebug
  
Dim ii As Integer
Dim iCol As Integer

  With objFGrid
    iCol = Col
    If iCol < 0 Or iCol < .FixedCols Then Exit Sub
    If Len(.TextMatrix(0, .Cols - 1)) > 0 Then Exit Sub
    
    'move column
    For ii = .Cols - 1 To iCol + 1 Step -1
      MoveColumn ii - 1, ii
    Next
  End With
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":ClsSuperFGrid.AddColumn"
End Sub

'remove column
Public Sub RemoveColumn(Col As Integer)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":ClsSuperFGrid.RemoveColumn(" & Col & ")", etFullDebug
  
Dim ii As Integer
Dim iCol As Integer
  
  With objFGrid
    iCol = Col
    If iCol < 0 Or iCol < .FixedCols Then Exit Sub
    
    'move column
    For ii = iCol To .Cols - 2
      MoveColumn ii + 1, ii
    Next
  End With
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":ClsSuperFGrid.RemoveColumn"
End Sub

'move column from to
Private Sub MoveColumn(ColFrom As Integer, ColTo As Integer)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":ClsSuperFGrid.MoveColumn(" & ColFrom & "," & ColTo & ")", etFullDebug
  
Dim ii As Integer
Dim TypeA As TypeActionFlexGrid
Dim iCol As Integer
Dim iRow As Integer
Dim bExists As Boolean
  
  With objFGrid
    If ColFrom < 0 Or ColFrom < .FixedCols Then Exit Sub
    If ColTo < 0 Or ColTo < .FixedCols Then Exit Sub

    'save current postion
    iCol = .Col
    iRow = .Row

    'move column
    For ii = 0 To .Rows - 1
      'change column action
      TypeA = GetTypeAction(ColFrom, ii)
      bExists = False
      
      bExists = (GetTypeAction(ColTo, ii) = TypeA)
      If Not bExists Then objAction(ActionPosition(ColFrom, ii, TypeA)).Col = ColTo
      
      Select Case TypeA
        Case TAFG_CHECK
          'copy check image
          .Col = ColFrom
          .Row = ii
          
          Select Case .CellPicture
            Case frmDummy.imgUnchecked.Picture
              .Col = ColTo
              Set .CellPicture = frmDummy.imgUnchecked.Picture
            Case frmDummy.imgChecked.Picture
              .Col = ColTo
              Set .CellPicture = frmDummy.imgChecked.Picture
          End Select
          If Not bExists Then
            .Col = ColFrom
            .CellPicture = LoadPicture("")
          End If
          
        Case Else
          'Copy text data
          .TextMatrix(ii, ColTo) = .TextMatrix(ii, ColFrom)
          .TextMatrix(ii, ColFrom) = ""
        
      End Select
    Next
    
    'resotre old postion
    .Col = iCol
    .Row = iRow
  End With
  Exit Sub

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":ClsSuperFGrid.MoveColumn"
End Sub
