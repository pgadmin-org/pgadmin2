VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSQLOutput 
   Caption         =   "SQL Output"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8160
   Icon            =   "frmSQLOutput.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   8160
   Begin VB.PictureBox picTools 
      Height          =   465
      Left            =   0
      ScaleHeight     =   405
      ScaleWidth      =   4905
      TabIndex        =   5
      Top             =   1215
      Width           =   4965
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   330
         Left            =   1755
         TabIndex        =   8
         ToolTipText     =   "Delete the selected record."
         Top             =   45
         Width           =   825
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Enabled         =   0   'False
         Height          =   330
         Left            =   900
         TabIndex        =   7
         ToolTipText     =   "Edit the selected record."
         Top             =   45
         Width           =   825
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   330
         Left            =   45
         TabIndex        =   6
         ToolTipText     =   "Add a new record."
         Top             =   45
         Width           =   825
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   330
         Left            =   2610
         TabIndex        =   13
         ToolTipText     =   "Delete the selected record."
         Top             =   45
         Width           =   825
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   330
         Left            =   1755
         TabIndex        =   10
         ToolTipText     =   "Add a new record."
         Top             =   45
         Visible         =   0   'False
         Width           =   1680
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   330
         Left            =   45
         TabIndex        =   9
         ToolTipText     =   "Add a new record."
         Top             =   45
         Visible         =   0   'False
         Width           =   1680
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "0 Records"
         Height          =   195
         Left            =   3510
         TabIndex        =   11
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.PictureBox picEdit 
      Height          =   1005
      Left            =   0
      ScaleHeight     =   63
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   297
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   4515
      Begin VB.PictureBox picScroll 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   885
         Left            =   0
         ScaleHeight     =   59
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   242
         TabIndex        =   2
         Top             =   0
         Width           =   3630
         Begin VB.TextBox txtField 
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   0
            Left            =   45
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   3
            Top             =   225
            Width           =   3300
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            Caption         =   "Field Label"
            Height          =   195
            Index           =   0
            Left            =   45
            TabIndex        =   4
            Top             =   45
            Width           =   765
         End
      End
      Begin VB.VScrollBar scScroll 
         Height          =   780
         LargeChange     =   100
         Left            =   3960
         SmallChange     =   10
         TabIndex        =   1
         Top             =   0
         Width           =   240
      End
   End
   Begin MSComctlLib.ListView lvData 
      Height          =   1185
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   5010
      _ExtentX        =   8837
      _ExtentY        =   2090
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
End
Attribute VB_Name = "frmSQLOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001, 2002, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence
'
' frmSQLOutput.frm - View/Edit SQL Query Results

Option Explicit
Dim rsSQL As New Recordset
Dim szDatabase As String
Dim szTable As String
Dim szSchema As String
Dim szWhere As String
Dim iUnique As Integer
Dim bUpdateable As Boolean

Private Sub cmdAdd_Click()
'On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLOutput.cmdAdd_Click()", etFullDebug

  BuildEditBox
  lblInfo.Caption = "Add Record"

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLOutput.cmdAdd_Click"
End Sub

Private Sub cmdCancel_Click()
'On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLOutput.cmdCancel_Click()", etFullDebug

  HideEditBox

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLOutput.cmdCancel_Click"
End Sub

Private Sub cmdDelete_Click()
'On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLOutput.cmdDelete_Click()", etFullDebug

Dim X As Integer
Dim Y As Integer
Dim Z As Integer
Dim szCriteria As String
Dim szFullCriteria As String
Dim rsCount As New Recordset
Dim szQuery As String
Dim szValues() As String
Dim szKeys() As String
Dim bFlag As Boolean
  If MsgBox("Are you sure you wish to delete the selected record?", vbQuestion + vbYesNo, "Delete Record?") = vbNo Then Exit Sub
  
  'Build the most concise WHERE clause we can. adDate and adDBDate fields should be
  'formatted as ISO dates. Not required if we found a suitable unique column.
  If iUnique = 0 Then
    For X = 0 To lvData.ColumnHeaders.Count - 1
      If X = 0 Then
        If lvData.SelectedItem.Text <> "" Then
          Select Case Val(Mid(lvData.ColumnHeaders(X + 1).Key, InStr(1, lvData.ColumnHeaders(X + 1).Key, ":") + 1))
            Case adDate
              szCriteria = szCriteria & QUOTE & lvData.ColumnHeaders(X + 1).Text & QUOTE & " = '" & Format(lvData.SelectedItem.Text, "yyyy-MM-dd") & "' AND "
            Case adDBDate
              szCriteria = szCriteria & QUOTE & lvData.ColumnHeaders(X + 1).Text & QUOTE & " = '" & Format(lvData.SelectedItem.Text, "yyyy-MM-dd") & "' AND "
            Case adDBTimeStamp
              szCriteria = szCriteria & QUOTE & lvData.ColumnHeaders(X + 1).Text & QUOTE & " = '" & Format(lvData.SelectedItem.Text, "yyyy-MM-dd hh:mm:ss") & "' AND "
            Case Else
              If ((InStr(1, lvData.SelectedItem.Text, vbCrLf) <> 0) + (InStr(1, lvData.SelectedItem.Text, "\n")) <> 0) = 0 Then
                szCriteria = szCriteria & QUOTE & lvData.ColumnHeaders(X + 1).Text & QUOTE & " = '" & dbSZ(lvData.SelectedItem.Text) & "' AND "
              End If
          End Select
        End If
      Else
        If lvData.SelectedItem.SubItems(X) <> "" Then
          Select Case Val(Mid(lvData.ColumnHeaders(X + 1).Key, InStr(1, lvData.ColumnHeaders(X + 1).Key, ":") + 1))
            Case adDate
              szCriteria = szCriteria & QUOTE & lvData.ColumnHeaders(X + 1).Text & QUOTE & " = '" & Format(lvData.SelectedItem.SubItems(X), "yyyy-MM-dd") & "' AND "
            Case adDBDate
              szCriteria = szCriteria & QUOTE & lvData.ColumnHeaders(X + 1).Text & QUOTE & " = '" & Format(lvData.SelectedItem.SubItems(X), "yyyy-MM-dd") & "' AND "
            Case adDBTimeStamp
              szCriteria = szCriteria & QUOTE & lvData.ColumnHeaders(X + 1).Text & QUOTE & " = '" & Format(lvData.SelectedItem.SubItems(X), "yyyy-MM-dd hh:mm:ss") & "' AND "
            Case Else
              If ((InStr(1, lvData.SelectedItem.Text, vbCrLf) <> 0) + (InStr(1, lvData.SelectedItem.Text, "\n")) <> 0) = 0 Then
                szCriteria = szCriteria & QUOTE & lvData.ColumnHeaders(X + 1).Text & QUOTE & " = '" & dbSZ(lvData.SelectedItem.SubItems(X)) & "' AND "
              End If
          End Select
        End If
      End If
    Next
  End If
  
  'Find out how many records would be affected. Abort if zero, update if 1 or
  'give the option to update all if > 1
  StartMsg "Counting matching records..."
  If Len(szCriteria) > 5 Then szCriteria = Mid(szCriteria, 1, Len(szCriteria) - 5)
  If frmMain.svr.dbVersion.VersionNum >= 7.3 Then
    szQuery = "SELECT count(*) AS count FROM " & szSchema & "." & szTable
  Else
    szQuery = "SELECT count(*) AS count FROM " & szTable
  End If
  
  'WHERE Clase
  If iUnique > 0 Then
    If txtField(iUnique - 1).Tag <> "Y" Then
      If iUnique = 1 Then
        szFullCriteria = " WHERE " & QUOTE & lvData.ColumnHeaders(iUnique).Text & QUOTE & " = '" & dbSZ(lvData.SelectedItem.Text) & "'"
      Else
        szFullCriteria = " WHERE " & QUOTE & lvData.ColumnHeaders(iUnique).Text & QUOTE & " = '" & dbSZ(lvData.SelectedItem.SubItems(iUnique - 1)) & "'"
      End If
    Else
      If (szCriteria <> "") Or (szWhere <> "") Then
        szFullCriteria = " WHERE " & szCriteria
        If (szCriteria <> "") And (szWhere <> "") Then szFullCriteria = szFullCriteria & " AND "
        If szWhere <> "" Then szFullCriteria = szFullCriteria & szWhere
      End If
    End If
  Else
    If (szCriteria <> "") Or (szWhere <> "") Then
      szFullCriteria = " WHERE " & szCriteria
      If (szCriteria <> "") And (szWhere <> "") Then szFullCriteria = szFullCriteria & " AND "
      If szWhere <> "" Then szFullCriteria = szFullCriteria & szWhere
    End If
  End If
  
  szQuery = szQuery & szFullCriteria
  Set rsCount = frmMain.svr.Databases(szDatabase).Execute(szQuery)
  
  'Prepare the delete query for later
  If frmMain.svr.dbVersion.VersionNum >= 7.3 Then
    szQuery = "DELETE FROM " & szSchema & "." & szTable & szFullCriteria
  Else
    szQuery = "DELETE FROM " & szTable & szFullCriteria
  End If
  
  EndMsg
  If Not rsCount.EOF Then
    Select Case rsCount!Count
      Case 0
        MsgBox "Could not locate the record for deletion in the database!", vbExclamation, "Error"
        GoTo Done
      Case 1
        StartMsg "Deleting record..."
        frmMain.svr.Databases(szDatabase).Execute szQuery, , , qryData
        lvData.ListItems.Remove (lvData.SelectedItem.Index)
        GoTo Done
      Case Else
        If MsgBox("The selected record could not be uniquely identified. " & rsCount!Count & " records match, and will all be deleted if you proceed. Do you wish to continue?", vbQuestion + vbYesNo, "Delete Multiple Records") = vbNo Then Exit Sub
        StartMsg "Deleting records..."
        frmMain.svr.Databases(szDatabase).Execute szQuery, , , qryData
        
        'Get all the values in the selected row, then iterate through all rows and delete matching
        ReDim szValues(lvData.ColumnHeaders.Count - 1)
        For X = 0 To lvData.ColumnHeaders.Count - 1
          If X = 0 Then
            szValues(X) = lvData.SelectedItem.Text
          Else
            szValues(X) = lvData.SelectedItem.SubItems(X)
          End If
        Next X
        
        'Delete matching rows.
        For X = lvData.ListItems.Count To 1 Step -1
          bFlag = False
          For Y = 1 To lvData.ColumnHeaders.Count - 1
            If szValues(Y) <> lvData.ListItems(X).SubItems(Y) Then
              bFlag = True
              Exit For
            End If
          Next Y
          If Not (bFlag Or szValues(0) <> lvData.ListItems(X).Text) Then
            lvData.ListItems.Remove lvData.ListItems(X).Index
          End If
        Next X
        GoTo Done
    End Select
  End If
Done:
  EndMsg
  If lvData.ListItems.Count > 0 Then
    lblInfo.Caption = "Record " & lvData.SelectedItem.Index & " of " & lvData.ListItems.Count
  Else
    lblInfo.Caption = "Record 0 of 0"
  End If
  If lvData.ListItems.Count > 0 And bUpdateable = True Then cmdEdit.Enabled = True
  If lvData.ListItems.Count > 0 And bUpdateable = True Then cmdDelete.Enabled = True
  If rsCount.State <> adStateClosed Then rsCount.Close
  Set rsCount = Nothing
  
  Exit Sub
Err_Handler:
  EndMsg
  If rsCount.State <> adStateClosed Then rsCount.Close
  Set rsCount = Nothing
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLOutput.cmdDelete_Click"
End Sub

Private Sub cmdRefresh_Click()
'On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLOutput.cmdSave_Click()", etFullDebug

  rsSQL.Requery
  RefreshData

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLOutput.cmdSave_Click"
End Sub

Private Sub cmdSave_Click()
'On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLOutput.cmdSave_Click()", etFullDebug

Dim szQuery As String
Dim szColumns As String
Dim szValues As String
Dim szCriteria As String
Dim szFullCriteria As String
Dim szCells() As String
Dim X As Integer
Dim Y As Integer
Dim Z As Integer
Dim bFlag As Boolean
Dim itmX As ListItem
Dim rsCount As New Recordset

  If lblInfo.Caption = "Add Record" Then
    'Add new record
    'First build lists of columns and values
    For X = 0 To lblField.Count - 1
      If txtField(X).Text <> "" Then
        szColumns = szColumns & QUOTE & lblField(X).Caption & QUOTE & ", "
        Select Case Val(Mid(lvData.ColumnHeaders(X + 1).Key, InStr(1, lvData.ColumnHeaders(X + 1).Key, ":") + 1))
          Case adDate
            szValues = szValues & "'" & Format(txtField(X).Text, "yyyy-MM-dd") & "', "
          Case adDBDate
            szValues = szValues & "'" & Format(txtField(X).Text, "yyyy-MM-dd") & "', "
          Case adDBTimeStamp
            szValues = szValues & "'" & Format(txtField(X).Text, "yyyy-MM-dd hh:mm:ss") & "', "
          Case Else
            szValues = szValues & "'" & dbSZ(txtField(X).Text) & "', "
        End Select
      End If
    Next X
    
    'Check the data, then trim the ', ' from the end of each string and create the SQL query
    If szColumns = "" Then
      EndMsg
      MsgBox "No data has been entered!", vbExclamation, "Error"
      Exit Sub
    End If
    If Len(szColumns) > 2 Then szColumns = "(" & Mid(szColumns, 1, Len(szColumns) - 2) & ")"
    If Len(szValues) > 2 Then szValues = "(" & Mid(szValues, 1, Len(szValues) - 2) & ")"
    If frmMain.svr.dbVersion.VersionNum >= 7.3 Then
      szQuery = "INSERT INTO " & szSchema & "." & szTable & " " & szColumns & " VALUES " & szValues
    Else
      szQuery = "INSERT INTO " & szTable & " " & szColumns & " VALUES " & szValues
    End If
    
    'Execute the query
    frmMain.svr.Databases(szDatabase).Execute szQuery, , , qryData
    
    'Now add the record to the grid. If the query failed, we won't get to here 'cos
    'we'll be in the error handler.
    Set itmX = lvData.ListItems.Add(, , txtField(0).Text)
    For X = 1 To lblField.Count - 1
      itmX.SubItems(X) = txtField(X).Text
    Next X
    GoTo Done
  Else
    'Update record
    'First build lists of columns and values.
    For X = 0 To lblField.Count - 1
      If txtField(X).Tag = "Y" Then
        Select Case Val(Mid(lvData.ColumnHeaders(X + 1).Key, InStr(1, lvData.ColumnHeaders(X + 1).Key, ":") + 1))
          Case adDate
            szValues = szValues & QUOTE & lblField(X).Caption & QUOTE & " = '" & Format(txtField(X).Text, "yyyy-MM-dd") & "', "
          Case adDBDate
            szValues = szValues & QUOTE & lblField(X).Caption & QUOTE & " = '" & Format(txtField(X).Text, "yyyy-MM-dd") & "', "
          Case adDBTimeStamp
            szValues = szValues & QUOTE & lblField(X).Caption & QUOTE & " = '" & Format(txtField(X).Text, "yyyy-MM-dd hh:mm:ss") & "', "
          Case Else
            szValues = szValues & QUOTE & lblField(X).Caption & QUOTE & " = '" & dbSZ(txtField(X).Text) & "', "
        End Select
      End If
    Next X
    
    'Build the most concise WHERE clause we can. adDate and adDBDate fields should be
    'formatted as ISO dates. This is not required if we found a unique column AND it's
    'not been modified
    If iUnique > 0 Then
      If txtField(iUnique - 1).Tag = "Y" Then
        For X = 0 To lvData.ColumnHeaders.Count - 1
          If X = 0 Then
            If lvData.SelectedItem.Text <> "" Then
              Select Case Val(Mid(lvData.ColumnHeaders(X + 1).Key, InStr(1, lvData.ColumnHeaders(X + 1).Key, ":") + 1))
                Case adDate
                  szCriteria = szCriteria & QUOTE & lvData.ColumnHeaders(X + 1).Text & QUOTE & " = '" & Format(lvData.SelectedItem.Text, "yyyy-MM-dd") & "' AND "
                Case adDBDate
                  szCriteria = szCriteria & QUOTE & lvData.ColumnHeaders(X + 1).Text & QUOTE & " = '" & Format(lvData.SelectedItem.Text, "yyyy-MM-dd") & "' AND "
                Case adDBTimeStamp
                  szCriteria = szCriteria & QUOTE & lvData.ColumnHeaders(X + 1).Text & QUOTE & " = '" & Format(lvData.SelectedItem.Text, "yyyy-MM-dd hh:mm:ss") & "' AND "
                Case Else
                  If ((InStr(1, lvData.SelectedItem.Text, vbCrLf) <> 0) + (InStr(1, lvData.SelectedItem.Text, "\n")) <> 0) = 0 Then
                    szCriteria = szCriteria & QUOTE & lvData.ColumnHeaders(X + 1).Text & QUOTE & " = '" & dbSZ(lvData.SelectedItem.Text) & "' AND "
                  End If
              End Select
            End If
          Else
            If lvData.SelectedItem.SubItems(X) <> "" Then
              Select Case Val(Mid(lvData.ColumnHeaders(X + 1).Key, InStr(1, lvData.ColumnHeaders(X + 1).Key, ":") + 1))
                Case adDate
                  szCriteria = szCriteria & QUOTE & lvData.ColumnHeaders(X + 1).Text & QUOTE & " = '" & Format(lvData.SelectedItem.SubItems(X), "yyyy-MM-dd") & "' AND "
                Case adDBDate
                  szCriteria = szCriteria & QUOTE & lvData.ColumnHeaders(X + 1).Text & QUOTE & " = '" & Format(lvData.SelectedItem.SubItems(X), "yyyy-MM-dd") & "' AND "
                Case adDBTimeStamp
                  szCriteria = szCriteria & QUOTE & lvData.ColumnHeaders(X + 1).Text & QUOTE & " = '" & Format(lvData.SelectedItem.SubItems(X), "yyyy-MM-dd hh:mm:ss") & "' AND "
                Case Else
                  If ((InStr(1, lvData.SelectedItem.SubItems(X), vbCrLf) <> 0) + (InStr(1, lvData.SelectedItem.SubItems(X), "\n")) <> 0) = 0 Then
                    szCriteria = szCriteria & QUOTE & lvData.ColumnHeaders(X + 1).Text & QUOTE & " = '" & dbSZ(lvData.SelectedItem.SubItems(X)) & "' AND "
                  End If
              End Select
            End If
          End If
        Next
      End If
    Else
      For X = 0 To lvData.ColumnHeaders.Count - 1
        If X = 0 Then
          If lvData.SelectedItem.Text <> "" Then
            Select Case Val(Mid(lvData.ColumnHeaders(X + 1).Key, InStr(1, lvData.ColumnHeaders(X + 1).Key, ":") + 1))
              Case adDate
                szCriteria = szCriteria & QUOTE & lvData.ColumnHeaders(X + 1).Text & QUOTE & " = '" & Format(lvData.SelectedItem.Text, "yyyy-MM-dd") & "' AND "
              Case adDBDate
                szCriteria = szCriteria & QUOTE & lvData.ColumnHeaders(X + 1).Text & QUOTE & " = '" & Format(lvData.SelectedItem.Text, "yyyy-MM-dd") & "' AND "
              Case adDBTimeStamp
                szCriteria = szCriteria & QUOTE & lvData.ColumnHeaders(X + 1).Text & QUOTE & " = '" & Format(lvData.SelectedItem.Text, "yyyy-MM-dd hh:mm:ss") & "' AND "
              Case Else
                If ((InStr(1, lvData.SelectedItem.Text, vbCrLf) <> 0) + (InStr(1, lvData.SelectedItem.Text, "\n")) <> 0) = 0 Then
                  szCriteria = szCriteria & QUOTE & lvData.ColumnHeaders(X + 1).Text & QUOTE & " = '" & dbSZ(lvData.SelectedItem.Text) & "' AND "
                End If
            End Select
          End If
        Else
          If lvData.SelectedItem.SubItems(X) <> "" Then
            Select Case Val(Mid(lvData.ColumnHeaders(X + 1).Key, InStr(1, lvData.ColumnHeaders(X + 1).Key, ":") + 1))
              Case adDate
                szCriteria = szCriteria & QUOTE & lvData.ColumnHeaders(X + 1).Text & QUOTE & " = '" & Format(lvData.SelectedItem.SubItems(X), "yyyy-MM-dd") & "' AND "
              Case adDBDate
                szCriteria = szCriteria & QUOTE & lvData.ColumnHeaders(X + 1).Text & QUOTE & " = '" & Format(lvData.SelectedItem.SubItems(X), "yyyy-MM-dd") & "' AND "
              Case adDBTimeStamp
                szCriteria = szCriteria & QUOTE & lvData.ColumnHeaders(X + 1).Text & QUOTE & " = '" & Format(lvData.SelectedItem.SubItems(X), "yyyy-MM-dd hh:mm:ss") & "' AND "
              Case Else
                If ((InStr(1, lvData.SelectedItem.SubItems(X), vbCrLf) <> 0) + (InStr(1, lvData.SelectedItem.SubItems(X), "\n")) <> 0) = 0 Then
                  szCriteria = szCriteria & QUOTE & lvData.ColumnHeaders(X + 1).Text & QUOTE & " = '" & dbSZ(lvData.SelectedItem.SubItems(X)) & "' AND "
                End If
            End Select
          End If
        End If
      Next
    End If

    'Check the data
    If szValues = "" Then
      EndMsg
      MsgBox "No data has been modified!", vbExclamation, "Error"
      Exit Sub
    End If
    
    'Find out how many records would be affected. Abort if zero, update if 1 or
    'give the option to update all if > 1
    StartMsg "Counting matching records..."
    If Len(szValues) > 2 Then szValues = Mid(szValues, 1, Len(szValues) - 2)
    If Len(szCriteria) > 5 Then szCriteria = Mid(szCriteria, 1, Len(szCriteria) - 5)
    If frmMain.svr.dbVersion.VersionNum >= 7.3 Then
      szQuery = "SELECT count(*) AS count FROM " & szSchema & "." & szTable
    Else
      szQuery = "SELECT count(*) AS count FROM " & szTable
    End If
    
    'WHERE Clase
    If iUnique > 0 Then
      If txtField(iUnique - 1).Tag <> "Y" Then
        If iUnique = 1 Then
          szFullCriteria = " WHERE " & QUOTE & lvData.ColumnHeaders(iUnique).Text & QUOTE & " = '" & dbSZ(lvData.SelectedItem.Text) & "'"
        Else
          szFullCriteria = " WHERE " & QUOTE & lvData.ColumnHeaders(iUnique).Text & QUOTE & " = '" & dbSZ(lvData.SelectedItem.SubItems(iUnique - 1)) & "'"
        End If
      Else
        If (szCriteria <> "") Or (szWhere <> "") Then
          szFullCriteria = " WHERE " & szCriteria
          If (szCriteria <> "") And (szWhere <> "") Then szFullCriteria = szFullCriteria & " AND "
          If szWhere <> "" Then szFullCriteria = szFullCriteria & szWhere
        End If
      End If
    Else
      If (szCriteria <> "") Or (szWhere <> "") Then
        szFullCriteria = " WHERE " & szCriteria
        If (szCriteria <> "") And (szWhere <> "") Then szFullCriteria = szFullCriteria & " AND "
        If szWhere <> "" Then szFullCriteria = szFullCriteria & szWhere
      End If
    End If
    
    szQuery = szQuery & szFullCriteria
    Set rsCount = frmMain.svr.Databases(szDatabase).Execute(szQuery)

    'Prepare the update query for later
    If frmMain.svr.dbVersion.VersionNum >= 7.3 Then
      szQuery = "UPDATE " & szSchema & "." & szTable & " SET " & szValues & szFullCriteria
    Else
      szQuery = "UPDATE " & szTable & " SET " & szValues & szFullCriteria
    End If
    
    EndMsg
    If Not rsCount.EOF Then
      Select Case rsCount!Count
        Case 0
          MsgBox "Could not locate the record for updating in the database!", vbExclamation, "Error"
          GoTo Done
        Case 1
          StartMsg "Updating record..."
          frmMain.svr.Databases(szDatabase).Execute szQuery, , , qryData
          'Update the grid
          For X = 0 To lblField.Count - 1
            If X = 0 Then
              lvData.SelectedItem.Text = txtField(X).Text
            Else
              lvData.SelectedItem.SubItems(X) = txtField(X).Text
            End If
          Next X
          GoTo Done
        Case Else
          If MsgBox("The selected record could not be uniquely identified. " & rsCount!Count & " records match, and will all be updated if you proceed. Do you wish to continue?", vbQuestion + vbYesNo, "Update Multiple Records") = vbNo Then Exit Sub
          StartMsg "Updating records..."
          frmMain.svr.Databases(szDatabase).Execute szQuery, , , qryData

          'Get all the values in the selected row, then iterate through all rows and update matching
          ReDim szCells(lvData.ColumnHeaders.Count - 1)
          For X = 0 To lvData.ColumnHeaders.Count - 1
            If X = 0 Then
              szCells(X) = lvData.SelectedItem.Text
            Else
              szCells(X) = lvData.SelectedItem.SubItems(X)
            End If
          Next X
          
          'Update matching rows.
          For X = lvData.ListItems.Count To 1 Step -1
            bFlag = False
            For Y = 1 To lvData.ColumnHeaders.Count - 1
              If szCells(Y) <> lvData.ListItems(X).SubItems(Y) Then
                bFlag = True
                Exit For
              End If
            Next Y
            If Not (bFlag Or szCells(0) <> lvData.ListItems(X).Text) Then
              For Z = 0 To lblField.Count - 1
                If Z = 0 Then
                  lvData.ListItems(X).Text = txtField(Z).Text
                Else
                  lvData.ListItems(X).SubItems(Z) = txtField(Z).Text
                End If
              Next Z
            End If
          Next X
          GoTo Done
      End Select
    End If
  End If
Done:
  EndMsg
  HideEditBox
  If lvData.ListItems.Count > 0 Then
    lblInfo.Caption = "Record " & lvData.SelectedItem.Index & " of " & lvData.ListItems.Count
  Else
    lblInfo.Caption = "Record 0 of 0"
  End If
  If lvData.ListItems.Count > 0 And bUpdateable = True Then cmdEdit.Enabled = True
  If lvData.ListItems.Count > 0 And bUpdateable = True Then cmdDelete.Enabled = True
  If rsCount.State <> adStateClosed Then rsCount.Close
  Set rsCount = Nothing
  
  Exit Sub
Err_Handler:
  EndMsg
  If rsCount.State <> adStateClosed Then rsCount.Close
  Set rsCount = Nothing
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLOutput.cmdSave_Click"
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Error Resume Next
  If rsSQL.State <> adStateClosed Then rsSQL.Close
  Set rsSQL = Nothing
End Sub

Public Sub Display(rsQuery As Recordset, szDB As String, szID As String)
'On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLOutput.Display(" & QUOTE & rsQuery.Source & QUOTE & ")", etFullDebug

Dim iStart As Integer
Dim iEnd As Integer
Dim iTemp As Integer
Dim X As Integer
Dim szQuery As String
Dim szChar As String
Dim szBits() As String
Dim bInQuotes As Boolean
Dim bFlag As Boolean
Dim objView As pgView
Dim szTemp As String
Dim szTempTable As String
Dim rsSchema As New Recordset
Dim szSchemas() As String

  Set rsSQL = rsQuery
  szDatabase = szDB
  
  'Set the Grid Font
  Set lvData.Font = ctx.Font

  'Figure out if the query is updateable. This is the case if:
  '1) There is one and only one table
  '2) The table is not actually a view
  'We must also get the tablename, and any WHERE clause to help
  'with update queries.
  
  'Start by converting any spaces inside double quotes to tabs which
  'should never appear in the SQL
  szQuery = ""
  bInQuotes = False
  For X = 1 To Len(rsSQL.Source)
    szChar = Mid(rsSQL.Source, X, 1)
    If szChar = QUOTE Then
      szQuery = szQuery & QUOTE
      bInQuotes = Not bInQuotes
    ElseIf szChar = " " And bInQuotes Then
      szQuery = szQuery & vbTab
    Else
      szQuery = szQuery & szChar
    End If
  Next X
  
  'Find the FROM clause. If it is inside single quotes then we
  'should try again - it won't in doubles as there are no spaces
  'in doubles anymore.
  iStart = 0
  bFlag = False
  bInQuotes = False
  While bFlag = False
    iStart = InStr(iStart + 1, UCase(szQuery), " FROM ")
    'If not found, look at the start of a line.
    If iStart = 0 Then iStart = InStr(iStart + 1, UCase(szQuery), vbCrLf & "FROM ")
    If iStart = 0 Then 'No FROMs found
      bFlag = True
    Else 'Found a FROM, check it's not in quotes
      For X = 1 To iStart
        If Mid(szQuery, X, 1) = "'" Then bInQuotes = Not bInQuotes
      Next X
      If Not bInQuotes Then bFlag = True
    End If
  Wend
  
  'If FROM is not found then we must have a tableless query
  '(eg. SELECT version()), otherwise, check for '(' before the
  'FROM. If one exists, then it must indicate a calculation [(col1 + col2)],
  'a function call [count(*)] or a sub select [(SELECT ...)]. We can also
  'check for occurances of the AS keyword which will prevent updates.
  'After all that, increment iStart past the FROM
  If iStart = 0 Then
    szTable = ""
    szWhere = ""
    bUpdateable = False
    GoTo GotInfo
  Else
    'Look for Brackets...
    bInQuotes = False
    iTemp = InStr(1, Mid(szQuery, 1, iStart), "(")
    If iTemp > 0 Then
      For X = 1 To iTemp
        If ((Mid(szQuery, X, 1) = "'") Or (Mid(szQuery, X, 1) = QUOTE)) Then bInQuotes = Not bInQuotes
      Next X
      If Not bInQuotes Then
        szTable = ""
        szWhere = ""
        bUpdateable = False
        GoTo GotInfo
      End If
    End If
    
    'Look for AS...
    bInQuotes = False
    iTemp = InStr(1, UCase(Mid(szQuery, 1, iStart)), " AS ")
    If iTemp = 0 Then iTemp = InStr(1, UCase(Mid(szQuery, 1, iStart)), vbCrLf & "AS ")
    If iTemp > 0 Then
      For X = 1 To iTemp
        If ((Mid(szQuery, X, 1) = "'") Or (Mid(szQuery, X, 1) = QUOTE)) Then bInQuotes = Not bInQuotes
      Next X
      If Not bInQuotes Then
        szTable = ""
        szWhere = ""
        bUpdateable = False
        GoTo GotInfo
      End If
    End If
    
    'All done, wind past the from.
    iStart = iStart + 6
  End If
  
  'Find the end of the FROM clause. This will be delimited by one of the
  'following, or the end of the string:
  'WHERE GROUP HAVING UNION INTERSECT EXCEPT ORDER FOR LIMIT
  iEnd = InStr(iStart, UCase(szQuery), " WHERE ")
  iTemp = InStr(iStart, UCase(szQuery), " GROUP ")
  If iTemp <> 0 And iTemp < iEnd Then iEnd = iTemp
  iTemp = InStr(iStart, UCase(szQuery), " HAVING ")
  If iTemp <> 0 And iTemp < iEnd Then iEnd = iTemp
  iTemp = InStr(iStart, UCase(szQuery), " UNION ")
  If iTemp <> 0 And iTemp < iEnd Then iEnd = iTemp
  iTemp = InStr(iStart, UCase(szQuery), " INTERSECT ")
  If iTemp <> 0 And iTemp < iEnd Then iEnd = iTemp
  iTemp = InStr(iStart, UCase(szQuery), " EXCEPT ")
  If iTemp <> 0 And iTemp < iEnd Then iEnd = iTemp
  iTemp = InStr(iStart, UCase(szQuery), " ORDER ")
  If iTemp <> 0 And iTemp < iEnd Then iEnd = iTemp
  iTemp = InStr(iStart, UCase(szQuery), " FOR ")
  If iTemp <> 0 And iTemp < iEnd Then iEnd = iTemp
  iTemp = InStr(iStart, UCase(szQuery), " LIMIT ")
  If iTemp <> 0 And iTemp < iEnd Then iEnd = iTemp
  If iEnd = 0 Then iEnd = Len(szQuery) + 1

  'Split the FROM clause by space. We can then iterate through each element of
  'the array to figure out whether we have more than one table. The following
  'conditions could determine that we have more than one table:
  '1) A trailing , on any element
  '2) An element containing JOIN INNER OUTER LEFT RIGHT FULL CROSS or [(]SELECT
  szBits = Split(Mid(szQuery, iStart, iEnd - iStart), " ")
  For X = 0 To UBound(szBits)
    If Right(szBits(X), 1) = "," Then
      szTable = ""
      szWhere = ""
      bUpdateable = False
      GoTo GotInfo
    End If
    If UCase(szBits(X)) = "JOIN" Or _
       UCase(szBits(X)) = "INNER" Or _
       UCase(szBits(X)) = "OUTER" Or _
       UCase(szBits(X)) = "LEFT" Or _
       UCase(szBits(X)) = "RIGHT" Or _
       UCase(szBits(X)) = "FULL" Or _
       UCase(szBits(X)) = "CROSS" Or _
       UCase(szBits(X)) = "SELECT" Or _
       UCase(szBits(X)) = "(SELECT" Then
      szTable = ""
      szWhere = ""
      bUpdateable = False
      GoTo GotInfo
    End If
  Next X

  'If we got this far then we should only have one table so we should
  'get it's name. It should be the first item in the array unless
  'ONLY was specified

  If UCase(szBits(0)) = "ONLY" Then
    szTable = Replace(szBits(1), vbTab, " ")
  Else
    szTable = Replace(szBits(0), vbTab, " ")
  End If
  
  'If the last character is a semi-colon, then the user probably entered a
  'psql style query: SELECT * FROM pg_class;
  If Right(szTable, 1) = ";" Then szTable = Mid(szTable, 1, Len(szTable) - 1)
  
  'At this point, szTable could look like pretty much any of the following:
  'schema.table, "schema"."table", "schema".table, schema."table", table, "table"
  'We need to search for a . outside any quote pairs. If there is one, then
  'we can get the schema and table
  bInQuotes = False
  For X = 1 To Len(szTable)
    szTemp = Mid(szTable, X, 1)
    If szTemp = QUOTE Then
      bInQuotes = Not bInQuotes
    ElseIf szTemp = "." And Not bInQuotes Then
      szSchema = Mid(szTable, 1, X - 1)
      szTable = Mid(szTable, X + 1)
    End If
  Next X
  
  'Check to see if our table is actually a view or sequence. If it is then we can't
  'update :-(
  
  If Left(szTable, 1) = QUOTE And Right(szTable, 1) = QUOTE Then
    szTempTable = Mid(szTable, 2, Len(szTable) - 2)
  Else
    szTempTable = szTable
  End If
  If frmMain.svr.dbVersion.VersionNum < 7.3 Then
    If frmMain.svr.Databases(szDatabase).Namespaces("public").Views.Exists(szTempTable) Or _
       frmMain.svr.Databases(szDatabase).Namespaces("public").Sequences.Exists(szTempTable) Then
      szTable = ""
      szWhere = ""
      bUpdateable = False
      GoTo GotInfo
    End If
  Else
    'This is a PITA with 7.3 as we can have multiple schemas in the search path.
    'However, if szSchema isn't empty then it's easy
    If szSchema <> "" Then
      ReDim szSchemas(0)
      szSchemas(0) = szSchema
    Else
      Set rsSchema = frmMain.svr.Databases(szDatabase).Execute("SELECT current_schemas(true) AS path", , , qrySystem)
      szTemp = rsSchema!Path & ""
      If Len(szTemp) > 2 Then szTemp = Mid(szTemp, 2, Len(szTemp) - 2)
      If rsSchema.State <> adStateClosed Then rsSchema.Close
      Set rsSchema = Nothing
      bInQuotes = False
      ReDim szSchemas(0)
      For X = 1 To Len(szTemp)
        If Mid(szTemp, X, 1) = "," Then
          If bInQuotes Then
            szSchemas(UBound(szSchemas)) = szSchemas(UBound(szSchemas)) & ","
          Else
            ReDim Preserve szSchemas(UBound(szSchemas) + 1)
          End If
        ElseIf Mid(szTemp, X, 1) = QUOTE Then
          bInQuotes = Not bInQuotes
        Else
          szSchemas(UBound(szSchemas)) = szSchemas(UBound(szSchemas)) & Mid(szTemp, X, 1)
        End If
      Next X
    End If
      
    'Now check for the existance of a view, sequence or table with the current name.
    'If a table exists, then we're all go. If it's as view or sequence then it's not
    'updateable
    For X = 0 To UBound(szSchemas)
      If frmMain.svr.Databases(szDatabase).Namespaces(szSchemas(X)).Views.Exists(szTempTable) Or _
         frmMain.svr.Databases(szDatabase).Namespaces(szSchemas(X)).Sequences.Exists(szTempTable) Then
        szTable = ""
        szWhere = ""
        bUpdateable = False
        GoTo GotInfo
      Else
        If frmMain.svr.Databases(szDatabase).Namespaces(szSchemas(X)).Tables.Exists(szTempTable) Then
          'Reset the schema & table names to nicely formatted ones
          szTable = fmtID(frmMain.svr.Databases(szDatabase).Namespaces(szSchemas(X)).Tables(szTempTable).Name)
          szSchema = fmtID(frmMain.svr.Databases(szDatabase).Namespaces(szSchemas(X)).Tables(szTempTable).Namespace)
          Exit For
        End If
      End If
    Next X
  End If
  
  'Yippee!
  bUpdateable = True
  
  'As we're updateable we should also extract any WHERE clause
  'to add to any query based updates we may do. This will help
  'us to locate the exact record that the user wants to update.
  iStart = 0
  bFlag = False
  bInQuotes = False
  While bFlag = False
    iStart = InStr(iStart + 1, UCase(szQuery), " WHERE ")
    If iStart = 0 Then 'No WHEREs found
      bFlag = True
    Else 'Found a WHERE, check it's not in quotes
      For X = 1 To iStart
        If Mid(szQuery, X, 1) = "'" Then bInQuotes = Not bInQuotes
      Next X
      If Not bInQuotes Then bFlag = True
    End If
  Wend

  'If WHERE is not found then we must have an 'all records' query
  'otherwise increment iStart past the WHERE
  If iStart = 0 Then
    szWhere = ""
    GoTo GotInfo
  Else
    iStart = iStart + 7
  End If
  
  'Find the end of the WHERE clause. This will be delimited by one of the
  'following, or the end of the string:
  'GROUP HAVING UNION INTERSECT EXCEPT ORDER FOR LIMIT
  iEnd = InStr(iStart, UCase(szQuery), " GROUP ")
  iTemp = InStr(iStart, UCase(szQuery), " HAVING ")
  If iTemp <> 0 And iTemp < iEnd Then iEnd = iTemp
  iTemp = InStr(iStart, UCase(szQuery), " UNION ")
  If iTemp <> 0 And iTemp < iEnd Then iEnd = iTemp
  iTemp = InStr(iStart, UCase(szQuery), " INTERSECT ")
  If iTemp <> 0 And iTemp < iEnd Then iEnd = iTemp
  iTemp = InStr(iStart, UCase(szQuery), " EXCEPT ")
  If iTemp <> 0 And iTemp < iEnd Then iEnd = iTemp
  iTemp = InStr(iStart, UCase(szQuery), " ORDER ")
  If iTemp <> 0 And iTemp < iEnd Then iEnd = iTemp
  iTemp = InStr(iStart, UCase(szQuery), " FOR ")
  If iTemp <> 0 And iTemp < iEnd Then iEnd = iTemp
  iTemp = InStr(iStart, UCase(szQuery), " LIMIT ")
  If iTemp <> 0 And iTemp < iEnd Then iEnd = iTemp
  If iEnd = 0 Then iEnd = Len(szQuery) + 1
  
  szWhere = Trim(Mid(szQuery, iStart, iEnd - iStart))

GotInfo:

  'Setup the form
  Me.Caption = "SQL Output " & szID & ": " & rsQuery.Source
  If bUpdateable Then
    frmMain.svr.LogEvent "Recordset appears to be updateable.", etMiniDebug
    cmdEdit.Enabled = True
    cmdAdd.Enabled = True
    cmdDelete.Enabled = True
  Else
    frmMain.svr.LogEvent "Recordset appears to NOT be updateable.", etMiniDebug
    cmdEdit.Enabled = False
    cmdAdd.Enabled = False
    cmdDelete.Enabled = False
  End If
  LoadGrid

  'Attempt to figure out a Unique Column for safer updating
  If bUpdateable Then
    If frmMain.svr.dbVersion.VersionNum >= 7.3 Then
      szTemp = szSchema
      If Left(szTemp, 1) = QUOTE Then szTemp = Right(szTemp, Len(szTemp) - 1)
      If Right(szTemp, 1) = QUOTE Then szTemp = Left(szTemp, Len(szTemp) - 1)
    Else
      szTemp = "public"
    End If
    
    szTempTable = szTable
    If Left(szTempTable, 1) = QUOTE Then szTempTable = Right(szTempTable, Len(szTempTable) - 1)
    If Right(szTempTable, 1) = QUOTE Then szTempTable = Left(szTempTable, Len(szTempTable) - 1)
    For iTemp = 1 To lvData.ColumnHeaders.Count
      If frmMain.svr.Databases(szDatabase).Namespaces(szTemp).Tables.Exists(szTempTable) Then
        If frmMain.svr.Databases(szDatabase).Namespaces(szTemp).Tables(szTempTable).Columns.Exists(lvData.ColumnHeaders(iTemp).Text) Then
          If ((frmMain.svr.Databases(szDatabase).Namespaces(szTemp).Tables(szTempTable).Columns(lvData.ColumnHeaders(iTemp).Text).PrimaryKey) And _
             (rsSQL.Fields(iTemp - 1).Type <> adDate) And _
             (rsSQL.Fields(iTemp - 1).Type <> adDBDate) And _
             (rsSQL.Fields(iTemp - 1).Type <> adDBTimeStamp)) Then
            iUnique = iTemp
            Exit For
          End If
        End If
      End If
    Next iTemp
    If iUnique = 0 Then
      frmMain.svr.LogEvent "Couldn't find a suitable unique column for use as a key.", etMiniDebug
    Else
      frmMain.svr.LogEvent "Found column: " & lvData.ColumnHeaders(iUnique).Text & " for use as a key.", etMiniDebug
    End If
  End If
  
  Exit Sub
Err_Handler:
  If rsSchema.State <> adStateClosed Then rsSchema.Close
  Set rsSchema = Nothing
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLOutput.Display"
End Sub

Private Sub Form_Resize()
'On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLOutput.Form_Resize()", etFullDebug

  If Me.WindowState <> 1 And Me.ScaleHeight > 0 Then
    If Me.WindowState = 0 Then
      If Me.Width < 5820 Then Me.Width = 5820
      If Me.Height < 3600 Then Me.Height = 3600
    End If
    
    picTools.Visible = True
    picTools.Width = Me.ScaleWidth
    picTools.Top = Me.ScaleHeight - picTools.Height
    lvData.Width = Me.ScaleWidth
    lvData.Height = Me.ScaleHeight - picTools.Height
    picEdit.Height = lvData.Height
    picEdit.Width = lvData.Width
    picScroll.Width = picEdit.ScaleWidth - scScroll.Width
    scScroll.Left = picScroll.Width
    scScroll.Height = picEdit.ScaleHeight
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLOutput.Form_Resize"
End Sub

Private Sub LoadGrid()
'On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLOutput.LoadGrid()", etFullDebug

Dim X As Long

  cmdSave.Visible = False
  cmdCancel.Visible = False
  
  'Load Headers
  lvData.ColumnHeaders.Clear
  For X = 0 To rsSQL.Fields.Count - 1
    lvData.ColumnHeaders.Add , "C" & X & ":" & rsSQL.Fields(X).Type, rsSQL.Fields(X).Name & ""
  Next X
      
  RefreshData
  
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLOutput.LoadGrid"
End Sub

Private Sub RefreshData()
'On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLOutput.RefreshData()", etFullDebug

Dim itmX As ListItem
Dim X As Long

  'Load Data
  StartMsg "Loading data..."
  lvData.ListItems.Clear
  lblInfo.Caption = "Record 0 of 0"
  If Not (rsSQL.EOF And rsSQL.BOF) Then
    While Not rsSQL.EOF
    
      'Add the listitem
      Select Case rsSQL.Fields(0).Type
        Case adDBTime
          Set itmX = lvData.ListItems.Add(, , Format(rsSQL.Fields(0).Value & "", "ttttt"))
        Case Else
          Set itmX = lvData.ListItems.Add(, , rsSQL.Fields(0).Value & "")
      End Select
        
      'Add the extra fields
      For X = 1 To rsSQL.Fields.Count - 1
        Select Case rsSQL.Fields(X).Type
          Case adDBTime
            itmX.SubItems(X) = Format(rsSQL.Fields(X).Value & "", "ttttt")
          Case Else
            itmX.SubItems(X) = rsSQL.Fields(X).Value & ""
        End Select
      Next
      rsSQL.MoveNext
    Wend
    lblInfo.Caption = "Record " & lvData.SelectedItem.Index & " of " & lvData.ListItems.Count
  End If
  
  'Set Buttons
  If lvData.ListItems.Count > 0 And bUpdateable = True Then cmdEdit.Enabled = True
  If lvData.ListItems.Count > 0 And bUpdateable = True Then cmdDelete.Enabled = True
  
  EndMsg
  
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLOutput.RefreshData"
End Sub

Private Sub lvData_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLOutput.lvData_ColumnClick(" & QUOTE & ColumnHeader.Text & QUOTE & ")", etFullDebug

  lvData.Sorted = True
  'Sort by the select column. If we already are, then switch the direction.
  If lvData.SortKey = (ColumnHeader.Index - 1) Then
    If lvData.SortOrder = lvwAscending Then
      lvData.SortOrder = lvwDescending
    Else
      lvData.SortOrder = lvwAscending
    End If
  Else
    lvData.SortOrder = lvwAscending
    lvData.SortKey = (ColumnHeader.Index - 1)
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLOutput.lvData_ColumnClick"
End Sub

Private Sub lvData_ItemClick(ByVal Item As MSComctlLib.ListItem)
'On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLOutput.lvData_ItemClick(" & QUOTE & Item.Text & QUOTE & ")", etFullDebug

  lblInfo.Caption = "Record " & lvData.SelectedItem.Index & " of " & lvData.ListItems.Count
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLOutput.lvData_ItemClick"
End Sub

Private Sub BuildEditBox()
'On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLOutput.BuildEditBox()", etFullDebug

Dim X As Integer

  lblField(0).Top = 3
  txtField(0).Top = lblField(0).Top + lblField(0).Height
  txtField(0).Width = picScroll.Width - 6
  Set txtField(0).Font = ctx.Font
  lblField(0).Caption = lvData.ColumnHeaders(1).Text
  If lblField(0).Caption = "oid" Or _
     lblField(0).Caption = "cmax" Or _
     lblField(0).Caption = "xmax" Or _
     lblField(0).Caption = "cmin" Or _
     lblField(0).Caption = "xmin" Or _
     lblField(0).Caption = "ctid" Then
    txtField(0).Locked = True
  Else
    txtField(0).Locked = False
  End If
  For X = 2 To lvData.ColumnHeaders.Count
    Load lblField(X - 1)
    Load txtField(X - 1)
    lblField(X - 1).Visible = True
    txtField(X - 1).Visible = True
    lblField(X - 1).Top = txtField(X - 2).Top + txtField(X - 2).Height + 1
    txtField(X - 1).Top = lblField(X - 1).Top + lblField(X - 1).Height
    txtField(X - 1).Width = picScroll.Width - 6
    Set txtField(X - 1).Font = ctx.Font
    txtField(X - 1).TabIndex = txtField(X - 2).TabIndex + 1
    lblField(X - 1).Caption = lvData.ColumnHeaders(X).Text
    If lblField(X - 1).Caption = "oid" Or _
       lblField(0).Caption = "cmax" Or _
       lblField(0).Caption = "xmax" Or _
       lblField(0).Caption = "cmin" Or _
       lblField(0).Caption = "xmin" Or _
       lblField(0).Caption = "ctid" Then
      txtField(X - 1).Locked = True
    Else
      txtField(X - 1).Locked = False
    End If
  Next
  picScroll.Height = txtField(X - 2).Top + txtField(X - 2).Height + 1
  picEdit.Visible = True
  scScroll.Max = picScroll.ScaleHeight - picEdit.ScaleHeight
  cmdAdd.Visible = False
  cmdEdit.Visible = False
  cmdDelete.Visible = False
  cmdRefresh.Visible = False
  cmdSave.Visible = True
  cmdCancel.Visible = True
  txtField(0).SetFocus
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLOutput.BuildEditBox"
End Sub

Private Sub cmdEdit_Click()
'On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLOutput.cmdEdit_Click()", etFullDebug

Dim X As Long

  BuildEditBox
  For X = 0 To lvData.ColumnHeaders.Count - 1
    If X = 0 Then
      txtField(X).Text = lvData.SelectedItem.Text
    Else
      txtField(X).Text = lvData.SelectedItem.SubItems(X)
    End If
    txtField(X).Tag = ""
  Next
  lblInfo.Caption = "Edit Record"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLOutput.cmdEdit_Click"
End Sub

Private Sub scScroll_Change()
'On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLOutput.scScroll_Change()", etFullDebug

  picScroll.Top = -scScroll.Value
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLOutput.scScroll_Change"
End Sub

Private Sub txtField_Change(Index As Integer)
'On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLOutput.txtField_Change(" & Index & ")", etFullDebug

  txtField(Index).Tag = "Y"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLOutput.txtField_Change"
End Sub

Private Sub txtField_GotFocus(Index As Integer)
'On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLOutput.txtField_GotFocus(" & Index & ")", etFullDebug

Dim X As Long

  For X = 0 To txtField.Count - 1
    If X = Index Then
      txtField(X).BackColor = &H8000000E
    Else
      txtField(X).BackColor = &H8000000F
    End If
  Next
  If txtField(Index).Top + txtField(Index).Height > picEdit.ScaleHeight - picScroll.Top Then
    If lblField(Index).Top > scScroll.Max Then
      picScroll.Top = scScroll.Max
      scScroll.Value = scScroll.Max
    Else
      picScroll.Top = -lblField(Index).Top
      scScroll.Value = -picScroll.Top
    End If
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLOutput.txtField_GotFocus"
End Sub

Private Sub HideEditBox()
'On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSQLOutput.HideEditBox()", etFullDebug

Dim X As Integer

  txtField(0).Text = ""
  txtField(0).Tag = ""
  For X = 2 To lvData.ColumnHeaders.Count
    Unload lblField(X - 1)
    Unload txtField(X - 1)
  Next
  cmdAdd.Visible = True
  cmdEdit.Visible = True
  cmdDelete.Visible = True
  cmdRefresh.Visible = True
  cmdSave.Visible = False
  cmdCancel.Visible = False
  picEdit.Visible = False
  If lvData.ListItems.Count > 0 Then
    lblInfo.Caption = "Record " & lvData.SelectedItem.Index & " of " & lvData.ListItems.Count
  Else
    lblInfo.Caption = "Record 0 of 0"
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLOutput.HideEditBox"
End Sub

