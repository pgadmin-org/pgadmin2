Attribute VB_Name = "basActions"
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the Artistic Licence
'
' basActions.bas - Things that we do to objects (for want of a better description!)

Option Explicit

Public Sub Vacuum(bAnalyse As Boolean)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":basActions.Vacuum(" & bAnalyse & ")", etFullDebug
  
  'If a table is selected then Vacuum it alone, otherwise vacuum the entire database. We don't do columns
  'because there is no easy way to get the table name.
  Select Case ctx.CurrentObject.ObjectType
    Case "Table"
      If frmMain.svr.dbVersion.VersionNum < 7.2 Then If MsgBox(§§TrasLang§§("WARNING: Table vacuuming should only be performed when there is no one using the table.") & vbCrLf & vbCrLf & §§TrasLang§§("Do you wish to continue?"), vbQuestion + vbYesNo) = vbNo Then Exit Sub
      StartMsg §§TrasLang§§("Vacuuming ") & ctx.CurrentObject.Identifier & §§TrasLang§§(" in database: ") & ctx.CurrentDB & "..."
      frmMain.svr.Databases(ctx.CurrentDB).Vacuum bAnalyse, ctx.CurrentObject.Identifier
      EndMsg
      MsgBox §§TrasLang§§("The table '") & ctx.CurrentObject.Identifier & §§TrasLang§§("' in database '") & ctx.CurrentDB & §§TrasLang§§("' has been vacuumed."), vbInformation
    Case Else
      If frmMain.svr.dbVersion.VersionNum < 7.2 Then If MsgBox(§§TrasLang§§("WARNING: Database vacuuming should only be performed when there is no one using the database.") & vbCrLf & vbCrLf & §§TrasLang§§("Do you wish to continue?"), vbQuestion + vbYesNo) = vbNo Then Exit Sub
      StartMsg §§TrasLang§§("Vacuuming ") & ctx.CurrentDB & "..."
      frmMain.svr.Databases(ctx.CurrentDB).Vacuum bAnalyse
      EndMsg
      MsgBox §§TrasLang§§("The database '") & ctx.CurrentDB & §§TrasLang§§("' has been vacuumed."), vbInformation
  End Select
  
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basActions.Vacuum"
End Sub

Public Sub Reindex()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":basActions.Reindex()", etFullDebug

  Select Case ctx.CurrentObject.ObjectType
    Case "Table"
      If MsgBox(§§TrasLang§§("Are you sure you wish to reindex the table: ") & ctx.CurrentObject.Identifier & "?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
      StartMsg §§TrasLang§§("Reindexing ") & ctx.CurrentObject.Identifier & §§TrasLang§§(" in database: ") & ctx.CurrentDB & "..."
      ctx.CurrentObject.Reindex
      EndMsg
      MsgBox §§TrasLang§§("The table '") & ctx.CurrentObject.Identifier & §§TrasLang§§("' in database '") & ctx.CurrentDB & §§TrasLang§§("' has been reindexed."), vbInformation
    Case "Index"
      If MsgBox(§§TrasLang§§("Are you sure you wish to rebuild the index: ") & ctx.CurrentObject.Identifier & "?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
      StartMsg §§TrasLang§§("Rebuilding ") & ctx.CurrentObject.Identifier & §§TrasLang§§(" in database: ") & ctx.CurrentDB & "..."
      ctx.CurrentObject.Reindex
      EndMsg
      MsgBox §§TrasLang§§("The index '") & ctx.CurrentObject.Identifier & §§TrasLang§§("' in database '") & ctx.CurrentDB & §§TrasLang§§("' has been rebuilt."), vbInformation
    Case Else
      If MsgBox(§§TrasLang§§("Are you sure you wish to reindex the database: ") & ctx.CurrentDB & "?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
      If MsgBox(§§TrasLang§§("Do you want to force the reindex operation?"), vbQuestion + vbYesNo) = vbNo Then
        StartMsg §§TrasLang§§("Reindexing ") & ctx.CurrentDB & "..."
        frmMain.svr.Databases(ctx.CurrentDB).Reindex False
      Else
        StartMsg §§TrasLang§§("Reindexing ") & ctx.CurrentDB & §§TrasLang§§(" (forced)...")
        frmMain.svr.Databases(ctx.CurrentDB).Reindex True
      End If
      EndMsg
      MsgBox §§TrasLang§§("The database '") & ctx.CurrentDB & §§TrasLang§§("' has been reindexed"), vbInformation
  End Select
  
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basActions.Reindex"
End Sub

Public Sub Drop()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":basActions.Drop()", etFullDebug
 
Dim objItem As ListItem
Dim objNode As Node
Dim szType As String
Dim szIdentifier As String
Dim szPath() As String

  If ctx.CurrentObject Is Nothing Then Exit Sub
  If ctx.CurrentObject.ObjectType <> "User" And ctx.CurrentObject.ObjectType <> "Group" Then
    If ctx.CurrentObject.SystemObject Then
      MsgBox §§TrasLang§§("You cannot drop system objects!"), vbExclamation, §§TrasLang§§("Error")
      Exit Sub
    End If
    If ctx.CurrentObject.ObjectType = "Table" Then
      If ctx.AskDeleteObjectDatabase Then
        If MsgBox(§§TrasLang§§("Are you sure you wish to drop the table '") & ctx.CurrentObject.Identifier & "'? All Indexes, Rules and Triggers on this table will also be dropped." & vbCrLf & vbCrLf & "This action cannot be undone.", vbYesNo + vbQuestion, §§TrasLang§§("Drop ") & ctx.CurrentObject.ObjectType) = vbNo Then Exit Sub
      End If
    Else
      If ctx.AskDeleteObjectDatabase Then
        If MsgBox(§§TrasLang§§("Are you sure you wish to drop the ") & ctx.CurrentObject.ObjectType & " '" & ctx.CurrentObject.Identifier & "'?" & vbCrLf & vbCrLf & "This action cannot be undone.", vbYesNo + vbQuestion, §§TrasLang§§("Drop ") & ctx.CurrentObject.ObjectType) = vbNo Then Exit Sub
      End If
    End If
  Else
    If ctx.AskDeleteObjectDatabase Then
      If MsgBox(§§TrasLang§§("Are you sure you wish to drop the ") & ctx.CurrentObject.ObjectType & " '" & ctx.CurrentObject.Identifier & "'?", vbYesNo + vbQuestion, §§TrasLang§§("Drop ") & ctx.CurrentObject.ObjectType) = vbNo Then Exit Sub
    End If
  End If
  
  StartMsg §§TrasLang§§("Dropping ") & ctx.CurrentObject.ObjectType & ": " & ctx.CurrentObject.Identifier
        
  'Store the Identifier & node for later.
  szIdentifier = ctx.CurrentObject.Identifier
  Set objNode = ctx.CurrentObject.Tag
  
  'We need to figure out what type of object we are trying to drop to know where it is in the
  'object hierarchy
  
  Select Case ctx.CurrentObject.ObjectType
    Case "User"
      szType = "USR-"
      frmMain.svr.Users.Remove ctx.CurrentObject.Identifier
      objNode.Parent.Text = §§TrasLang§§("Users (") & objNode.Parent.Children - 1 & ")"
      frmMain.tv.Nodes.Remove objNode.Index
      
    Case "Group"
      szType = "GRP-"
      frmMain.svr.Groups.Remove ctx.CurrentObject.Identifier
      objNode.Parent.Text = §§TrasLang§§("Groups (") & objNode.Parent.Children - 1 & ")"
      frmMain.tv.Nodes.Remove objNode.Index
      
    Case "Database"
    
      'TODO - Dropping datbases seems to be nigh-on impossible. pgSchema *appears* to open more
      'connections to each database than we know about to close. Needs investigation.
    
      szType = "DAT-"
      frmMain.svr.Databases.Remove ctx.CurrentObject.Identifier
      objNode.Parent.Text = §§TrasLang§§("Databases (") & objNode.Parent.Children - 1 & ")"
      frmMain.tv.Nodes.Remove objNode.Index
      
    Case "Aggregate"
      szType = "AGG-"
      frmMain.svr.Databases(ctx.CurrentObject.Database).Namespaces(ctx.CurrentObject.Namespace).Aggregates.Remove ctx.CurrentObject.Identifier
      objNode.Parent.Text = §§TrasLang§§("Aggregates (") & objNode.Parent.Children - 1 & ")"
      frmMain.tv.Nodes.Remove objNode.Index
      
    Case "Cast"
      szType = "CST-"
      frmMain.svr.Databases(ctx.CurrentObject.Database).Casts.Remove ctx.CurrentObject.Identifier
      objNode.Parent.Text = §§TrasLang§§("Casts (") & objNode.Parent.Children - 1 & ")"
      frmMain.tv.Nodes.Remove objNode.Index

    Case "Conversion"
      szType = "CNV-"
      frmMain.svr.Databases(ctx.CurrentObject.Database).Namespaces(ctx.CurrentObject.Namespace).Conversions.Remove ctx.CurrentObject.Identifier
      objNode.Parent.Text = §§TrasLang§§("Conversions (") & objNode.Parent.Children - 1 & ")"
      frmMain.tv.Nodes.Remove objNode.Index
      
    Case "Domain"
      szType = "DOM-"
      frmMain.svr.Databases(ctx.CurrentObject.Database).Namespaces(ctx.CurrentObject.Namespace).Domains.Remove ctx.CurrentObject.Identifier
      objNode.Parent.Text = §§TrasLang§§("Domains (") & objNode.Parent.Children - 1 & ")"
      frmMain.tv.Nodes.Remove objNode.Index
    
    Case "Function"
      szType = "FNC-"
      frmMain.svr.Databases(ctx.CurrentObject.Database).Namespaces(ctx.CurrentObject.Namespace).Functions.Remove ctx.CurrentObject.Identifier
      objNode.Parent.Text = §§TrasLang§§("Functions (") & objNode.Parent.Children - 1 & ")"
      frmMain.tv.Nodes.Remove objNode.Index
    
    Case "Index"
      szType = "IND-"
      frmMain.svr.Databases(ctx.CurrentObject.Database).Namespaces(ctx.CurrentObject.Namespace).Tables(ctx.CurrentObject.Table).Indexes.Remove ctx.CurrentObject.Identifier
      objNode.Parent.Text = §§TrasLang§§("Indexes (") & objNode.Parent.Children - 1 & ")"
      frmMain.tv.Nodes.Remove objNode.Index
      
    Case "Language"
      szType = "LNG-"
      frmMain.svr.Databases(ctx.CurrentObject.Database).Languages.Remove ctx.CurrentObject.Identifier
      objNode.Parent.Text = §§TrasLang§§("Languages (") & objNode.Parent.Children - 1 & ")"
      frmMain.tv.Nodes.Remove objNode.Index
      
    Case "Schema"
      szType = "NSP-"
      frmMain.svr.Databases(ctx.CurrentObject.Database).Namespaces.Remove ctx.CurrentObject.Identifier
      objNode.Parent.Text = §§TrasLang§§("Schemas (") & objNode.Parent.Children - 1 & ")"
      frmMain.tv.Nodes.Remove objNode.Index
    
    Case "Operator"
      szType = "OPR-"
      frmMain.svr.Databases(ctx.CurrentObject.Database).Namespaces(ctx.CurrentObject.Namespace).Operators.Remove ctx.CurrentObject.Identifier
      objNode.Parent.Text = §§TrasLang§§("Operators (") & objNode.Parent.Children - 1 & ")"
      frmMain.tv.Nodes.Remove objNode.Index
    
    Case "OperatorClass"
      szType = "OPC-"
      frmMain.svr.Databases(ctx.CurrentObject.Database).Namespaces(ctx.CurrentObject.Namespace).OperatorsClass.Remove ctx.CurrentObject.Identifier
      objNode.Parent.Text = §§TrasLang§§("Operators Class (") & objNode.Parent.Children - 1 & ")"
      frmMain.tv.Nodes.Remove objNode.Index
    
    Case "Rule"
      szType = "RUL-"
      'verify if rule is for table or view
      If frmMain.svr.Databases(ctx.CurrentObject.Database).Namespaces(ctx.CurrentObject.Namespace).Tables.Exists(ctx.CurrentObject.Table) Then
        frmMain.svr.Databases(ctx.CurrentObject.Database).Namespaces(ctx.CurrentObject.Namespace).Tables(ctx.CurrentObject.Table).Rules.Remove ctx.CurrentObject.Identifier
      ElseIf frmMain.svr.Databases(ctx.CurrentObject.Database).Namespaces(ctx.CurrentObject.Namespace).Views.Exists(ctx.CurrentObject.Table) Then
        frmMain.svr.Databases(ctx.CurrentObject.Database).Namespaces(ctx.CurrentObject.Namespace).Views(ctx.CurrentObject.Table).Rules.Remove ctx.CurrentObject.Identifier
      End If
      objNode.Parent.Text = §§TrasLang§§("Rules (") & objNode.Parent.Children - 1 & ")"
      frmMain.tv.Nodes.Remove objNode.Index
    
    Case "Sequence"
      szType = "SEQ-"
      frmMain.svr.Databases(ctx.CurrentObject.Database).Namespaces(ctx.CurrentObject.Namespace).Sequences.Remove ctx.CurrentObject.Identifier
      objNode.Parent.Text = §§TrasLang§§("Sequences (") & objNode.Parent.Children - 1 & ")"
      frmMain.tv.Nodes.Remove objNode.Index
    
    Case "Table"
      szType = "TBL-"
      frmMain.svr.Databases(ctx.CurrentObject.Database).Namespaces(ctx.CurrentObject.Namespace).Tables.Remove ctx.CurrentObject.Identifier, True
      objNode.Parent.Text = §§TrasLang§§("Tables (") & objNode.Parent.Children - 1 & ")"
      frmMain.tv.Nodes.Remove objNode.Index
      
    Case "Check"
      szType = "CHK-"
      frmMain.svr.Databases(ctx.CurrentObject.Database).Namespaces(ctx.CurrentObject.Namespace).Tables(ctx.CurrentObject.Table).Checks.Remove ctx.CurrentObject.Identifier
      objNode.Parent.Text = §§TrasLang§§("Checks (") & objNode.Parent.Children - 1 & ")"
      frmMain.tv.Nodes.Remove objNode.Index
      
    Case "Column"
      szType = "COL-"
      frmMain.svr.Databases(ctx.CurrentObject.Database).Namespaces(ctx.CurrentObject.Namespace).Tables(ctx.CurrentObject.Table).Columns.Remove ctx.CurrentObject.Identifier
      objNode.Parent.Text = §§TrasLang§§("Columns (") & objNode.Parent.Children - 1 & ")"
      frmMain.tv.Nodes.Remove objNode.Index
      
    Case "Trigger"
      szType = "TRG-"
      frmMain.svr.Databases(ctx.CurrentObject.Database).Namespaces(ctx.CurrentObject.Namespace).Tables(ctx.CurrentObject.Table).Triggers.Remove ctx.CurrentObject.Identifier
      objNode.Parent.Text = §§TrasLang§§("Triggers (") & objNode.Parent.Children - 1 & ")"
      frmMain.tv.Nodes.Remove objNode.Index
      
    Case "Type"
      szType = "TYP-"
      frmMain.svr.Databases(ctx.CurrentObject.Database).Namespaces(ctx.CurrentObject.Namespace).Types.Remove ctx.CurrentObject.Identifier
      objNode.Parent.Text = §§TrasLang§§("Types (") & objNode.Parent.Children - 1 & ")"
      frmMain.tv.Nodes.Remove objNode.Index
    
    Case "View"
      szType = "VIE-"
      frmMain.svr.Databases(ctx.CurrentObject.Database).Namespaces(ctx.CurrentObject.Namespace).Views.Remove ctx.CurrentObject.Identifier
      objNode.Parent.Text = §§TrasLang§§("Views (") & objNode.Parent.Children - 1 & ")"
      frmMain.tv.Nodes.Remove objNode.Index
  
    Case Else
      MsgBox ctx.CurrentObject.ObjectType & §§TrasLang§§(" objects cannot be dropped."), vbExclamation, §§TrasLang§§("Error")
      Exit Sub
  End Select
  
  'Clear the current object
  Set ctx.CurrentObject = frmMain.svr
  
  'Simulate a click on the treeview to sort out the listview
  frmMain.tv_NodeClick frmMain.tv.SelectedItem
  If Not frmMain.lv.SelectedItem Is Nothing Then
    frmMain.lv_ItemClick frmMain.lv.SelectedItem
  End If
  
  EndMsg
 
  Exit Sub
Err_Handler:
  If Err.Number = 35606 Then Resume Next
  EndMsg
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basActions.Drop"
End Sub
