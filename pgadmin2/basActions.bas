Attribute VB_Name = "basActions"
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence
'
' basActions.bas - Things that we do to objects (for want of a better description!)

Option Explicit

Public Sub Vacuum(bAnalyse As Boolean)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":basActions.Vacuum(" & bAnalyse & ")", etFullDebug
  
  'If a table is selected then Vacuum it alone, otherwise vacuum the entire database. We don't do columns
  'because there is no easy way to get the table name.
  Select Case ctx.CurrentObject.ObjectType
    Case "Table"
      If MsgBox("WARNING: Table vacuuming should only be performed when there is no one using the table." & vbCrLf & vbCrLf & "Do you wish to continue?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
      StartMsg "Vacuuming " & ctx.CurrentObject.Identifier & " in database: " & ctx.CurrentDB & "..."
      frmMain.svr.Databases(ctx.CurrentDB).Vacuum bAnalyse, ctx.CurrentObject.Identifier
      EndMsg
      MsgBox "The table '" & ctx.CurrentObject.Identifier & "' in database '" & ctx.CurrentDB & "' has been vacuumed.", vbInformation
    Case Else
      If MsgBox("WARNING: Database vacuuming should only be performed when there is no one using the database." & vbCrLf & vbCrLf & "Do you wish to continue?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
      StartMsg "Vacuuming " & ctx.CurrentDB & "..."
      frmMain.svr.Databases(ctx.CurrentDB).Vacuum bAnalyse
      EndMsg
      MsgBox "The database '" & ctx.CurrentDB & "' has been vacuumed.", vbInformation
  End Select
  
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basActions.Vacuum"
End Sub

Public Sub Drop()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":basActions.Drop()", etFullDebug
 
Dim objItem As ListItem
Dim objNode As Node
Dim szType As String
Dim szIdentifier As String
Dim szPath() As String

  If ctx.CurrentObject.SystemObject Then
    MsgBox "You cannot drop system objects!", vbExclamation, "Error"
    Exit Sub
  End If
  
  If MsgBox("Are you sure you wish to drop the " & ctx.CurrentObject.ObjectType & " '" & ctx.CurrentObject.Identifier & "'?" & vbCrLf & vbCrLf & "This action cannot be undone.", vbYesNo + vbQuestion, "Drop " & ctx.CurrentObject.ObjectType) = vbNo Then Exit Sub
  StartMsg "Dropping " & ctx.CurrentObject.ObjectType & ": " & ctx.CurrentObject.Identifier
        
  'Store the Identifier for later.
  szIdentifier = ctx.CurrentObject.Identifier
  
  'We need to figure out what type of object we are trying to drop to know where it is in the
  'object hierarchy
  
  Select Case ctx.CurrentObject.ObjectType
    Case "User"

      szType = "USR-"
      frmMain.svr.Users.Remove ctx.CurrentObject.Identifier

    Case "Group"
      szType = "GRP-"
      frmMain.svr.Groups.Remove ctx.CurrentObject.Identifier
      
      'Delete any matching tree nodes
      For Each objNode In frmMain.tv.Nodes
        If Left(objNode.Key, 4) = szType And objNode.Text = szIdentifier Then frmMain.tv.Nodes.Remove objNode.Index
      Next objNode
      
    Case "Database"
    
      'TODO - Dropping datbases seems to be nigh-on impossible. pgSchema *appears* to open more
      'connections to each database than we know about to close. Needs investigation.
    
      szType = "DAT-"
      frmMain.svr.Databases.Remove ctx.CurrentObject.Identifier
      
      'Delete any matching tree nodes
      For Each objNode In frmMain.tv.Nodes
        If Left(objNode.Key, 4) = szType And objNode.Text = szIdentifier Then frmMain.tv.Nodes.Remove objNode.Index
      Next objNode
      
    Case "Aggregate"
      szType = "AGG-"
      frmMain.svr.Databases(ctx.CurrentObject.Database).Aggregates.Remove ctx.CurrentObject.Identifier
      
      'Delete any matching tree nodes
      For Each objNode In frmMain.tv.Nodes
        szPath = Split(objNode.FullPath, "\")
        If UBound(szPath) >= 2 Then
          If (Left(objNode.Key, 4) = szType) And (szPath(2) = ctx.CurrentObject.Database) And (objNode.Text = szIdentifier) Then
            objNode.Parent.Text = "Aggregates (" & objNode.Parent.Children - 1 & ")"
            frmMain.tv.Nodes.Remove objNode.Index
          End If
        End If
      Next objNode
    
    Case "Function"
      szType = "FNC-"
      frmMain.svr.Databases(ctx.CurrentObject.Database).Functions.Remove ctx.CurrentObject.Identifier
      
      'Delete any matching tree nodes
      For Each objNode In frmMain.tv.Nodes
        szPath = Split(objNode.FullPath, "\")
        If UBound(szPath) >= 2 Then
          If (Left(objNode.Key, 4) = szType) And (szPath(2) = ctx.CurrentObject.Database) And (objNode.Text = szIdentifier) Then
            objNode.Parent.Text = "Functions (" & objNode.Parent.Children - 1 & ")"
            frmMain.tv.Nodes.Remove objNode.Index
          End If
        End If
      Next objNode
    
    Case "Index"
      szType = "IND-"
      frmMain.svr.Databases(ctx.CurrentObject.Database).Tables(ctx.CurrentObject.Table).Indexes.Remove ctx.CurrentObject.Identifier
      
      'Delete any matching tree nodes
      For Each objNode In frmMain.tv.Nodes
        szPath = Split(objNode.FullPath, "\")
        If UBound(szPath) >= 2 Then
          If (Left(objNode.Key, 4) = szType) And (szPath(2) = ctx.CurrentObject.Database) And (objNode.Text = szIdentifier) Then
            objNode.Parent.Text = "Indexes (" & objNode.Parent.Children - 1 & ")"
            frmMain.tv.Nodes.Remove objNode.Index
          End If
        End If
      Next objNode
      
    Case "Language"
      szType = "LNG-"
      frmMain.svr.Databases(ctx.CurrentObject.Database).Languages.Remove ctx.CurrentObject.Identifier
      
      'Delete any matching tree nodes
      For Each objNode In frmMain.tv.Nodes
        szPath = Split(objNode.FullPath, "\")
        If UBound(szPath) >= 2 Then
          If (Left(objNode.Key, 4) = szType) And (szPath(2) = ctx.CurrentObject.Database) And (objNode.Text = szIdentifier) Then
            objNode.Parent.Text = "Languages (" & objNode.Parent.Children - 1 & ")"
            frmMain.tv.Nodes.Remove objNode.Index
          End If
        End If
      Next objNode
    
    Case "Operator"
      szType = "OPR-"
      frmMain.svr.Databases(ctx.CurrentObject.Database).Operators.Remove ctx.CurrentObject.Identifier
      
      'Delete any matching tree nodes
      For Each objNode In frmMain.tv.Nodes
        szPath = Split(objNode.FullPath, "\")
        If UBound(szPath) >= 2 Then
          If (Left(objNode.Key, 4) = szType) And (szPath(2) = ctx.CurrentObject.Database) And (objNode.Text = szIdentifier) Then
            objNode.Parent.Text = "Operators (" & objNode.Parent.Children - 1 & ")"
            frmMain.tv.Nodes.Remove objNode.Index
          End If
        End If
      Next objNode
    
    Case "Rule"
      szType = "RUL-"
      frmMain.svr.Databases(ctx.CurrentObject.Database).Tables(ctx.CurrentObject.Table).Rules.Remove ctx.CurrentObject.Identifier
      
      'Delete any matching tree nodes
      For Each objNode In frmMain.tv.Nodes
        szPath = Split(objNode.FullPath, "\")
        If UBound(szPath) >= 2 Then
          If (Left(objNode.Key, 4) = szType) And (szPath(2) = ctx.CurrentObject.Database) And (objNode.Text = szIdentifier) Then
            objNode.Parent.Text = "Rules (" & objNode.Parent.Children - 1 & ")"
            frmMain.tv.Nodes.Remove objNode.Index
          End If
        End If
      Next objNode
    
    Case "Sequence"
      szType = "SEQ-"
      frmMain.svr.Databases(ctx.CurrentObject.Database).Sequences.Remove ctx.CurrentObject.Identifier
      
      'Delete any matching tree nodes
      For Each objNode In frmMain.tv.Nodes
        szPath = Split(objNode.FullPath, "\")
        If UBound(szPath) >= 2 Then
          If (Left(objNode.Key, 4) = szType) And (szPath(2) = ctx.CurrentObject.Database) And (objNode.Text = szIdentifier) Then
            objNode.Parent.Text = "Sequences (" & objNode.Parent.Children - 1 & ")"
            frmMain.tv.Nodes.Remove objNode.Index
          End If
        End If
      Next objNode
    
    Case "Table"
      szType = "TBL-"
      frmMain.svr.Databases(ctx.CurrentObject.Database).Tables.Remove ctx.CurrentObject.Identifier
      
      'Delete any matching tree nodes
      For Each objNode In frmMain.tv.Nodes
        szPath = Split(objNode.FullPath, "\")
        If UBound(szPath) >= 2 Then
          If (Left(objNode.Key, 4) = szType) And (szPath(2) = ctx.CurrentObject.Database) And (objNode.Text = szIdentifier) Then
            objNode.Parent.Text = "Tables (" & objNode.Parent.Children - 1 & ")"
            frmMain.tv.Nodes.Remove objNode.Index
          End If
        End If
      Next objNode
    
    Case "Trigger"
      szType = "TRG-"
      frmMain.svr.Databases(ctx.CurrentObject.Database).Tables(ctx.CurrentObject.Table).Indexes.Remove ctx.CurrentObject.Identifier
      
      'Delete any matching tree nodes
      For Each objNode In frmMain.tv.Nodes
        szPath = Split(objNode.FullPath, "\")
        If UBound(szPath) >= 2 Then
          If (Left(objNode.Key, 4) = szType) And (szPath(2) = ctx.CurrentObject.Database) And (objNode.Text = szIdentifier) Then
            objNode.Parent.Text = "Triggers (" & objNode.Parent.Children - 1 & ")"
            frmMain.tv.Nodes.Remove objNode.Index
          End If
        End If
      Next objNode
      
    Case "Type"
      szType = "TYP-"
      frmMain.svr.Databases(ctx.CurrentObject.Database).Types.Remove ctx.CurrentObject.Identifier
      
      'Delete any matching tree nodes
      For Each objNode In frmMain.tv.Nodes
        szPath = Split(objNode.FullPath, "\")
        If UBound(szPath) >= 2 Then
          If (Left(objNode.Key, 4) = szType) And (szPath(2) = ctx.CurrentObject.Database) And (objNode.Text = szIdentifier) Then
            objNode.Parent.Text = "Types (" & objNode.Parent.Children - 1 & ")"
            frmMain.tv.Nodes.Remove objNode.Index
          End If
        End If
      Next objNode
    
    Case "View"
      szType = "VIE-"
      frmMain.svr.Databases(ctx.CurrentObject.Database).Views.Remove ctx.CurrentObject.Identifier
      
      'Delete any matching tree nodes
      For Each objNode In frmMain.tv.Nodes
        szPath = Split(objNode.FullPath, "\")
        If UBound(szPath) >= 2 Then
          If (Left(objNode.Key, 4) = szType) And (szPath(2) = ctx.CurrentObject.Database) And (objNode.Text = szIdentifier) Then
            objNode.Parent.Text = "Views (" & objNode.Parent.Children - 1 & ")"
            frmMain.tv.Nodes.Remove objNode.Index
          End If
        End If
      Next objNode
  
    Case Else
      MsgBox ctx.CurrentObject.ObjectType & " objects cannot be dropped.", vbExclamation, "Error"
      Exit Sub
  End Select
  
  'Clear the current object
  Set ctx.CurrentObject = frmMain.svr
  
  'Simulate a click on the treeview to sort out the listview
  frmMain.tv_NodeClick frmMain.tv.SelectedItem
  
  EndMsg
 
  Exit Sub
Err_Handler:
  If Err.Number = 35606 Then Resume Next
  EndMsg
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basActions.Drop"
End Sub
