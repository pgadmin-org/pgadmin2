Attribute VB_Name = "basClone"
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the Artistic Licence
'
' basClone.bas - Contains clone object function

Option Explicit

Private ObjDbClone

'Inizialize clone object
Public Sub InitClone()
  ClearObjDb
End Sub

'clear object database
Public Sub ClearObjDb()
  Set ObjDbClone = Nothing
End Sub

'copy object database
Public Sub CopyObjDb()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":basClone.CopyObjDb", etFullDebug

  'vetify type object
  Select Case ctx.CurrentObject.ObjectType
    Case "Domain", "Table", "View", "Group", "User", "Function", "Aggregate", "Operator", "Cast", "Type", "Conversion"
      Set ObjDbClone = ctx.CurrentObject
      frmMain.mnuEditPaste.Enabled = True
      frmMain.mnuPopupPaste.Enabled = True

    Case Else
      MsgBox "The current object type cannot be copied!", vbExclamation, "Error"
  End Select

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basClone.CopyObjDb"
End Sub

'paste object database
Public Sub PasteObjDb()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":basClone.PasteObjDb", etFullDebug

  If ObjDbClone Is Nothing Then
    MsgBox "You must select an object to copy!", vbExclamation, "Error"
    Exit Sub
  End If
  
  Select Case ObjDbClone.ObjectType
    Case "Domain", "Table", "View", "Function", "Aggregate", "Operator", "Type"
      If ctx.CurrentNS = "" Then
        MsgBox "You must select a schema to paste the object into!", vbExclamation, "Error"
        Exit Sub
      End If
    
    Case "Cast"
      If ctx.CurrentDB = "" Then
        MsgBox "You must select a database to paste the new object into!", vbExclamation, "Error"
        Exit Sub
      End If
    
  End Select

  frmClone.Initialise ObjDbClone
  frmClone.Show vbModal

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basClone.PasteObjDb"
End Sub

'clone type
Public Function CloneType(szNewName As String, szDatabase As String, szNamespace As String) As pgType
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":basClone.CloneType(" & QUOTE & szNewName & QUOTE & "," & QUOTE & szDatabase & QUOTE & "," & QUOTE & szNamespace & QUOTE & ")", etFullDebug

Dim objNewType As pgType

  Set objNewType = frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Types.Add(szNewName, ObjDbClone.InputFunction, ObjDbClone.OutputFunction, ObjDbClone.InternalLength, ObjDbClone.Default, ObjDbClone.Element, ObjDbClone.Delimiter, ObjDbClone.PassedByValue, ObjDbClone.Alignment, ObjDbClone.Storage, ObjDbClone.Comment)
  Set CloneType = objNewType
 
  Exit Function
  
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basClone.CloneType"
End Function

'clone cast
Public Function CloneCast(szDatabase As String) As pgCast
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":basClone.CloneCast(" & QUOTE & szDatabase & QUOTE & ")", etFullDebug

Dim objNewCast As pgCast

  Set objNewCast = frmMain.svr.Databases(szDatabase).Casts.Add(ObjDbClone.Source, ObjDbClone.Target, ObjDbClone.Funct, ObjDbClone.Context)
  Set CloneCast = objNewCast
 
  Exit Function
  
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basClone.CloneCast"
End Function

'clone operator
Public Function CloneOperator(szNewName As String, szDatabase As String, szNamespace As String) As pgOperator
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":basClone.CloneOperator(" & QUOTE & szNewName & QUOTE & "," & QUOTE & szDatabase & QUOTE & "," & QUOTE & szNamespace & QUOTE & ")", etFullDebug

Dim objNewOperator As pgOperator

  Set objNewOperator = frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Operators.Add(szNewName, ObjDbClone.OperatorFunction, ObjDbClone.LeftOperandType, ObjDbClone.RightOperandType, ObjDbClone.Commutator, ObjDbClone.Negator, ObjDbClone.RestrictFunction, ObjDbClone.JoinFunction, ObjDbClone.HashJoins, ObjDbClone.LeftTypeSortOperator, ObjDbClone.RightTypeSortOperator, ObjDbClone.Comment)
  Set CloneOperator = objNewOperator
 
  Exit Function
  
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basClone.CloneOperator"
End Function

'clone aggregate
Public Function CloneAggregate(szNewName As String, szDatabase As String, szNamespace As String) As pgAggregate
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":basClone.CloneAggregate(" & QUOTE & szNewName & QUOTE & "," & QUOTE & szDatabase & QUOTE & "," & QUOTE & szNamespace & QUOTE & ")", etFullDebug

Dim objNewAggregate As pgAggregate

  Set objNewAggregate = frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Aggregates.Add(szNewName, ObjDbClone.InputType, ObjDbClone.StateFunction, ObjDbClone.StateType, ObjDbClone.FinalFunction, ObjDbClone.InitialCondition, ObjDbClone.Comment)
  Set CloneAggregate = objNewAggregate

  Exit Function
  
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basClone.CloneAggregate"
End Function

'clone function
Public Function CloneFunction(szNewName As String, szDatabase As String, szNamespace As String) As pgFunction
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":basClone.CloneFunction(" & QUOTE & szNewName & QUOTE & "," & QUOTE & szDatabase & QUOTE & "," & QUOTE & szNamespace & QUOTE & ")", etFullDebug

Dim objNewFunction As pgFunction
Dim szArguments As String
Dim vData

  'Get the identifier/arguments in case we need it
  For Each vData In ObjDbClone.Arguments
    szArguments = szArguments & vData & ", "
  Next
  If Len(szArguments) > 2 Then szArguments = Left(szArguments, Len(szArguments) - 2)
  
  Set objNewFunction = frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Functions.Add(szNewName, szArguments, ObjDbClone.Returns, ObjDbClone.Source, ObjDbClone.Language, ObjDbClone.Cachable, ObjDbClone.Strict, ObjDbClone.Comment, ObjDbClone.Volatility, ObjDbClone.SecDef.ObjDbClone.RetSet)
  
  'clone acl
  CloneAcl objNewFunction
  
  Set CloneFunction = objNewFunction

  Exit Function
  
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basClone.CloneFunction"
End Function

'clone domain
Public Function CloneDomain(szNewName As String, szDatabase As String, szNamespace As String) As pgDomain
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":basClone.CloneDomain(" & QUOTE & szNewName & QUOTE & "," & QUOTE & szDatabase & QUOTE & "," & QUOTE & szNamespace & QUOTE & ")", etFullDebug

Dim objNewDomain As pgDomain

  Set objNewDomain = frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Domains.Add(szNewName, ObjDbClone.BaseType, ObjDbClone.Length, ObjDbClone.NumericScale, ObjDbClone.Default, ObjDbClone.NotNull, ObjDbClone.Comment)
  Set CloneDomain = objNewDomain

  Exit Function
  
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basClone.CloneDomain"
End Function

'clone user
Public Function CloneUser(szNewName As String) As pgUser
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":basClone.CloneUser(" & QUOTE & szNewName & QUOTE & ")", etFullDebug

Dim objMember As Variant
Dim objGroup As pgGroup
Dim lNextID As Long
Dim objNewUser As pgUser
Dim objVar As pgVar
      
  'Set defaults
  For Each objNewUser In frmMain.svr.Users
    If objNewUser.ID > lNextID Then lNextID = objNewUser.ID
  Next
    
  Set objNewUser = frmMain.svr.Users.Add(szNewName, lNextID + 1, , ObjDbClone.CreateDatabases, ObjDbClone.Superuser, ObjDbClone.AccountExpires)
    
  'add user to group
  For Each objGroup In frmMain.svr.Groups
    For Each objMember In objGroup.Members
      If objMember = ObjDbClone.Name Then
        objGroup.Members.Add szNewName
      End If
    Next
  Next
      
  'clone variable user
  For Each objVar In ObjDbClone.UserVars
    objNewUser.UserVars.AddOrUpdate objVar.Name, objVar.Value
  Next
      
  Set CloneUser = objNewUser
  Exit Function
  
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basClone.CloneUser"
End Function

'clone group
Public Function CloneGroup(szNewName As String) As pgGroup
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":basClone.CloneGroup(" & QUOTE & szNewName & QUOTE & ")", etFullDebug

Dim objMember As Variant
Dim lNextID As Long
Dim objNewGroup As pgGroup
      
  'Set defaults
  For Each objNewGroup In frmMain.svr.Groups
    If objNewGroup.ID > lNextID Then lNextID = objNewGroup.ID
  Next
    
  Set objNewGroup = frmMain.svr.Groups.Add(szNewName, lNextID + 1)

  'add user in group
  For Each objMember In ObjDbClone.Members
    frmMain.svr.Groups(szNewName).Members.Add objMember
  Next
      
  Set CloneGroup = objNewGroup
  Exit Function
  
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basClone.CloneGroup"
End Function

'clone conversion
Public Function CloneConversion(szNewName As String, szDatabase As String, szNamespace As String) As pgConversion
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":basClone.CloneConversion(" & QUOTE & szNewName & QUOTE & "," & QUOTE & szDatabase & QUOTE & "," & QUOTE & szNamespace & QUOTE & ")", etFullDebug

Dim objNewConversion As pgConversion

  Set objNewConversion = frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Conversions.Add(szNewName, False, ObjDbClone.ForEncoding, ObjDbClone.ToEncoding, ObjDbClone.Proc)
  
  Set CloneConversion = objNewConversion
  Exit Function
  
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basClone.CloneConversion"
End Function

'clone view
Public Function CloneView(szNewName As String, szDatabase As String, szNamespace As String) As pgView
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":basClone.CloneView(" & QUOTE & szNewName & QUOTE & "," & QUOTE & szDatabase & QUOTE & "," & QUOTE & szNamespace & QUOTE & ")", etFullDebug

Dim objNewView As pgView
Dim objRule As pgRule

  Set objNewView = frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Views.Add(szNewName, ObjDbClone.Definition, ObjDbClone.Comment)
  
  'clone acl
  CloneAcl objNewView

  'create rule
  StartMsg "Creating Rules..."
  For Each objRule In ObjDbClone.Rules
    If Not objNewView.Rules.Exists(objRule.Name) Then
      objNewView.Rules.Add objRule.Name, objRule.RuleEvent, objRule.Condition, objRule.DoInstead, objRule.Action, objRule.Comment
    End If
  Next
  
  Set CloneView = objNewView
  Exit Function
  
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basClone.CloneView"
End Function

'clone table
Public Function CloneTable(szNewName As String, szDatabase As String, szNamespace As String, Optional bCopyData As Boolean = False) As pgTable
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":basClone.CloneTable(" & QUOTE & szNewName & QUOTE & "," & QUOTE & szDatabase & QUOTE & "," & QUOTE & szNamespace & QUOTE & ")", etFullDebug

Dim objColumn As pgColumn
Dim objCheck As pgCheck
Dim objForeignKey As pgForeignKey
Dim objRelationship As pgRelationship
Dim szColumns As String
Dim szPrimaryKeys As String
Dim szChecks As String
Dim szForeignKeys As String
Dim szLocalColumn As String
Dim szReferencedColumn As String
Dim vInheritedTable As Variant
Dim szInherits As String
Dim objNewTable As pgTable
Dim szSQL As String
Dim objIndex As pgIndex
Dim vData
Dim szNameIndex As String
Dim objRule As pgRule
Dim objTrigger As pgTrigger

  'Build the column list
  For Each objColumn In ObjDbClone.Columns
    'no system column
    If Not objColumn.SystemObject Then
      szColumns = szColumns & objColumn.Identifier & " " & objColumn.DataType
            
      'verify if column require length
      If LCase(Left(objColumn.DataType, 4)) = "char" Or _
          LCase(Left(objColumn.DataType, 7)) = "varchar" Then
        szColumns = szColumns & "(" & objColumn.Length & ")"
      ElseIf LCase(Left(objColumn.DataType, 7)) = "numeric" Then
        szColumns = szColumns & "(" & objColumn.Length & ", " & objColumn.NumericScale & ")"
      End If
      
      If objColumn.Default <> "" Then szColumns = szColumns & " DEFAULT " & objColumn.Default
      If objColumn.NotNull Then szColumns = szColumns & " NOT NULL"
      szColumns = szColumns & ", "
      
      'Add to the Primary Key list if required.
      If objColumn.PrimaryKey Then szPrimaryKeys = szPrimaryKeys & fmtID(objColumn.Name) & ", "
    End If
  Next
  If Len(szColumns) > 2 Then szColumns = Left(szColumns, Len(szColumns) - 2)
    
  'Add the Primary Keys
  If Len(szPrimaryKeys) > 2 Then szPrimaryKeys = Left(szPrimaryKeys, Len(szPrimaryKeys) - 2)

  'Add Checks
  For Each objCheck In ObjDbClone.Checks
    szChecks = szChecks & "CONSTRAINT " & objCheck.Identifier & " CHECK (" & objCheck.Definition & "), "
  Next
  If Len(szChecks) > 2 Then szChecks = Left(szChecks, Len(szChecks) - 2)
    
  'Add Foreign Keys
  For Each objForeignKey In ObjDbClone.ForeignKeys
    For Each objRelationship In objForeignKey.Relationships
      szLocalColumn = szLocalColumn & objRelationship.LocalColumn & ", "
      szReferencedColumn = szReferencedColumn & objRelationship.ReferencedColumn & ", "
    Next objRelationship
    If Len(szLocalColumn) > 2 Then szLocalColumn = Left(szLocalColumn, Len(szLocalColumn) - 2)
    If Len(szReferencedColumn) > 2 Then szReferencedColumn = Left(szReferencedColumn, Len(szReferencedColumn) - 2)
    
    szForeignKeys = szForeignKeys & "CONSTRAINT " & objForeignKey.Identifier & " FOREIGN KEY (" & szLocalColumn & ") "
    szForeignKeys = szForeignKeys & "REFERENCES " & objForeignKey.ReferencedTable & " (" & szReferencedColumn & ")"
    szForeignKeys = szForeignKeys & " ON DELETE " & objForeignKey.OnDelete
    szForeignKeys = szForeignKeys & " ON UPDATE " & objForeignKey.OnUpdate
    If objForeignKey.Deferrable Then szForeignKeys = szForeignKeys & " DEFERRABLE"
    szForeignKeys = szForeignKeys & " INITIALLY " & objForeignKey.Initially & ", "
  Next
  If Len(szForeignKeys) > 2 Then szForeignKeys = Left(szForeignKeys, Len(szForeignKeys) - 2)
    
  'Add Inherits
  For Each vInheritedTable In ObjDbClone.InheritedTables
    szInherits = szInherits & vInheritedTable & ", "
  Next
  If Len(szInherits) > 2 Then szInherits = Left(szInherits, Len(szInherits) - 2)
  
  Set objNewTable = frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Tables.Add(szNewName, szColumns, szPrimaryKeys, szChecks, szForeignKeys, szInherits, ObjDbClone.Comment, ObjDbClone.HasOIDs)
  
  'Add any comments for the columns.
  For Each objColumn In ObjDbClone.Columns
    If objColumn.Comment <> "" Then objNewTable.Columns(objColumn.Name).Comment = objColumn.Comment
  Next
  
  'clone acl
  CloneAcl objNewTable
  
  'copy data table
  If bCopyData Then
    If ObjDbClone.Database = objNewTable.Database Then
      StartMsg "Copying data...."
      szSQL = "INSERT INTO " & fmtID(objNewTable.Namespace) & "." & fmtID(objNewTable.Name) & " SELECT * FROM " & fmtID(ObjDbClone.Namespace) & "." & fmtID(ObjDbClone.Name)
      frmMain.svr.Databases(objNewTable.Database).Execute szSQL
    End If
  End If
  
  'create index
  StartMsg "Creating Indexes..."
  szColumns = ""
  For Each objIndex In ObjDbClone.Indexes
    'no sistem index
    If Not objIndex.SystemObject Then
      'create name index and column list
      szNameIndex = objNewTable.Name
      szColumns = ""
      For Each vData In objIndex.IndexedColumns
        szColumns = szColumns & fmtID(vData) & ", "
        szNameIndex = szNameIndex & "_" & vData
      Next
      If Len(szColumns) > 2 Then szColumns = Left(szColumns, Len(szColumns) - 2)
      
      objNewTable.Indexes.Add szNameIndex, objIndex.Unique, szColumns, objIndex.IndexType, objIndex.Comment, objIndex.Constraint
    End If
  Next

  'create rule
  StartMsg "Creating Rules..."
  For Each objRule In ObjDbClone.Rules
    objNewTable.Rules.Add objRule.Name, objRule.RuleEvent, objRule.Condition, objRule.DoInstead, objRule.Action, objRule.Comment
  Next

  'create trigger
  StartMsg "Creating Triggers..."
  For Each objTrigger In ObjDbClone.Triggers
    objNewTable.Triggers.Add objTrigger.Name, objTrigger.TriggerFunction, objTrigger.Executes, objTrigger.TriggerEvent, objTrigger.ForEach, objTrigger.Comment
  Next
  
  EndMsg
  Set CloneTable = objNewTable
  Exit Function
  
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basClone.CloneTable"
End Function

Public Sub CloneAcl(objDb As Variant)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":basClone.CloneAcl(" & QUOTE & objDb.ObjectType & QUOTE & ")", etFullDebug

Dim szUserlist As String
Dim szAccesslist As String
Dim ii As Integer
Dim szUsers
Dim szAccess
Dim szEntity As String
Dim lACL As Long

  'Set the ACL on the View as required
  ParseACL ObjDbClone.ACL, szUserlist, szAccesslist
  szUsers = Split(szUserlist, "|")
  szAccess = Split(szAccesslist, "|")
  For ii = 0 To UBound(szUsers)
    If szUsers(ii) <> "" Then
      If szAccess(ii) = "none" Then
        'revoke
        If szUsers(ii) = "PUBLIC" Then
          objDb.Revoke szUsers(ii), aclAll
        ElseIf Left(szUsers(ii), 6) = "GROUP " Then
          objDb.Revoke "GROUP " & fmtID(Mid(szUsers(ii), 7)), aclAll
        Else
          objDb.Revoke fmtID(szUsers(ii)), aclAll
        End If
      Else
        'grant
        If szUsers(ii) = "PUBLIC" Then
          szEntity = "PUBLIC"
        ElseIf Left(szUsers(ii), 6) = "GROUP " Then
          szEntity = szUsers(ii)
        Else
          szEntity = fmtID(szUsers(ii))
        End If
        lACL = 0
        If InStr(1, szAccess(ii), "All") <> 0 Then lACL = lACL + aclAll
        If InStr(1, szAccess(ii), "Select") <> 0 Then lACL = lACL + aclSelect
        If InStr(1, szAccess(ii), "Update") <> 0 Then lACL = lACL + aclUpdate
        If InStr(1, szAccess(ii), "Delete") <> 0 Then lACL = lACL + aclDelete
        If InStr(1, szAccess(ii), "Insert") <> 0 Then lACL = lACL + aclInsert
        If InStr(1, szAccess(ii), "Rule") <> 0 Then lACL = lACL + aclRule
        If InStr(1, szAccess(ii), "References") <> 0 Then lACL = lACL + aclReferences
        If InStr(1, szAccess(ii), "Trigger") <> 0 Then lACL = lACL + aclTrigger
        If InStr(1, szAccess(ii), "Execute") <> 0 Then lACL = lACL + aclExecute
        objDb.Grant szEntity, lACL
      End If
    End If
  Next
  Exit Sub

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basClone.CloneAcl"
End Sub
