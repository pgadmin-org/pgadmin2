VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pgView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' pgSchema - PostgreSQL Schema Objects
' Copyright (C) 2001, 2002, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence

Option Explicit

'Properties
Private lOID As Long
Private szName As String
Private szIdentifier As String
Private bSystemObject As Boolean
Private szOwner As String
Private szACL As String
Private szDefinition As String
Private cnDatabase As Connection
Private lNamespaceOID As Long
Private szNamespace As String
Private vTag As Variant

Public Property Get FormattedID() As String
Attribute FormattedID.VB_Description = "Returns a formatted identifier for use in SQL statements."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgView.Property Get FormattedID()", etFullDebug
 
  If objServer.dbVersion.VersionNum >= 7.3 Then
    FormattedID = fmtID(szNamespace) & "." & fmtID(szName)
  Else
    FormattedID = fmtID(szName)
  End If
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Tag
Public Property Let Tag(ByVal vData As Variant)
Attribute Tag.VB_Description = "Set/Returns a variant for any use."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgView.Property Let Tag(" & vData & ")", etFullDebug

  vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Set Tag(ByVal vData As Variant)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgView.Property Set Tag(" & vData & ")", etFullDebug

  Set vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Tag() As Variant
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgView.Property Get Tag()", etFullDebug

  If IsObject(vTag) Then
    Set Tag = vTag
  Else
    Tag = vTag
  End If
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Namespace OID
Friend Property Let NamespaceOID(ByVal vData As Long)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgView.Property Let NamespaceOID(" & vData & ")", etFullDebug

  lNamespaceOID = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get NamespaceOID() As Long
Attribute NamespaceOID.VB_Description = "Returns the OID of the parent Namespace."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgView.Property Get NamespaceOID()", etFullDebug

  NamespaceOID = lNamespaceOID
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Namespace Name
Friend Property Let Namespace(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgView.Property Let Namespace(" & vData & ")", etFullDebug

  szNamespace = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Namespace() As String
Attribute Namespace.VB_Description = "Returns the name of the parent Namespace."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgView.Property Get Namespace()", etFullDebug

  Namespace = szNamespace
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property


'Set the Connection to use
Friend Property Set Connection(vData As Connection)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgView.Property Set Connection(Connection: " & vData & ")", etFullDebug

  Set cnDatabase = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Get the Database Name
Public Property Get Database() As String
Attribute Database.VB_Description = "Returns the parent database name."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgView.Property Get Database()", etFullDebug

Static szDB As String

  If szDB = "" Then szDB = GetDatabase(cnDatabase.ConnectionString)
  Database = szDB
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the OID
Friend Property Let OID(ByVal vData As Long)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgView.Property Let OID(" & vData & ")", etFullDebug

  lOID = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get OID() As Long
Attribute OID.VB_Description = "Returns the View OID."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgView.Property Get OID()", etFullDebug

  OID = lOID
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

Friend Property Let iName(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgView.Property Let iName(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szName = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

Friend Property Let Name(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgView.Property Let Name(" & vData & ")", etFullDebug
Dim szSQL As String
   
  szSQL = "ALTER TABLE " & FormattedID & " RENAME TO " & QUOTE & vData & QUOTE
  objServer.ExecSQL szSQL, cnDatabase
  szName = vData
  Identifier = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

Public Property Get Name() As String
Attribute Name.VB_Description = "Returns the View name."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgView.Property Get Name()", etFullDebug

  Name = szName
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Identifier
Friend Property Let Identifier(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgView.Property Let Identifier(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szIdentifier = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Identifier() As String
Attribute Identifier.VB_Description = "Returns the object identifier."
Attribute Identifier.VB_MemberFlags = "200"
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgView.Property Get Identifier()", etFullDebug

  Identifier = szIdentifier
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the Object Type
Public Property Get ObjectType() As String
Attribute ObjectType.VB_Description = "Returns the object type."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgView.Property Get ObjectType()", etFullDebug

  ObjectType = "View"
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return a boolean indicating whether this is a System Object
Friend Property Let SystemObject(ByVal vData As Boolean)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgView.Property Let SystemObject(" & vData & ")", etFullDebug

  bSystemObject = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get SystemObject() As Boolean
Attribute SystemObject.VB_Description = "Indicates whether this is a system object."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgView.Property Get SystemObject()", etFullDebug

  SystemObject = bSystemObject
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Owner
Friend Property Let Owner(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Let Owner(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szOwner = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Owner() As String
Attribute Owner.VB_Description = "Returns the View Owner."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get Owner()", etFullDebug

  Owner = szOwner
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the ACL
Friend Property Let ACL(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgView.Property Let ACL(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szACL = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get ACL() As String
Attribute ACL.VB_Description = "Returns the View ACL."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgView.Property Get ACL()", etFullDebug

  ACL = szACL
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the View definition
Public Property Let Definition(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgView.Property Let Definition(" & QUOTE & vData & QUOTE & ")", etFullDebug

Dim rs As New Recordset
Dim szFakeViewName As String
Dim szSQL As String

  'Create a fake view for testing purpose
  szFakeViewName = "tmp_" & Format(Now, "yymmdd_hhnnss")
  objServer.Databases(Database).Execute iSQL(szFakeViewName, vData)
    
  If objServer.dbVersion.VersionNum >= 7.3 Then
    'If we are here, then drop view and rename fake view
    szSQL = "DROP VIEW " & FormattedID & ";" & vbCrLf
    szSQL = szSQL & "ALTER TABLE " & fmtID(szNamespace) & "." & fmtID(szFakeViewName) & " RENAME TO " & fmtID(szName) & ";"
    objServer.Databases(Database).Execute szSQL
    
    'Select oid and definition back from the DB...
    szSQL = SQL_GET_VIEWS7_3 & " AND relname = '" & szName & "' AND relnamespace = " & lNamespaceOID & ";"
  Else
    'If we are here, then drop view and rename fake view
    szSQL = "DROP VIEW " & FormattedID & ";" & vbCrLf
    szSQL = szSQL & "ALTER TABLE " & fmtID(szFakeViewName) & " RENAME TO " & fmtID(szName) & ";"
    objServer.Databases(Database).Execute szSQL
    
    'Select oid and definition back from the DB...
    szSQL = SQL_GET_VIEWS7_1 & " AND relname = '" & szName & "';"
  End If
  
  Set rs = objServer.ExecSQL(szSQL, cnDatabase)
  If Not rs.EOF Then
    lOID = rs!OID & ""
    szDefinition = rs!Definition & ""
  End If
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  
  Exit Property
  
Err_Handler:
If rs.State <> adStateClosed Then rs.Close
Set rs = Nothing
objServer.iLogError Err
End Property

Friend Property Let iDefinition(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgViews.Property Let iDefinition(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szDefinition = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

Public Property Get Definition() As String
Attribute Definition.VB_Description = "Returns the View definition."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgView.Property Get Definition()", etFullDebug

  Definition = szDefinition
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Comment
Public Property Let Comment(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgView.Property Let Comment(" & vData & ")", etFullDebug

Dim szSQL As String

  szSQL = "COMMENT ON VIEW " & FormattedID & " IS '" & dbSZ(vData) & "'"
  objServer.ExecSQL szSQL, cnDatabase
  objServer.Databases(Database).KillCommentCache
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Comment() As String
Attribute Comment.VB_Description = "Sets/Returns the View comment."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgView.Property Get Comment()", etFullDebug

  If (szName <> "") Then
    Comment = objServer.Databases(Database).LookupComment(lOID)
  Else
    Comment = ""
  End If
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Get the row count
Public Property Get Rows() As Long
Attribute Rows.VB_Description = "Returns the number of rows represented in the view."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgView.Property Get Rows()", etFullDebug

Dim rs As New Recordset

  Set rs = objServer.ExecSQL("SELECT count(*) AS rows FROM " & FormattedID, cnDatabase)

  If Not rs.EOF Then Rows = rs!Rows
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
 
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Grant Permissions
Public Sub Grant(ByVal UserOrGroup As String, ByVal Privilege As aclPrivileges)
Attribute Grant.VB_Description = "Grant privileges on the View."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgView.Grant(" & QUOTE & UserOrGroup & QUOTE & ", " & Privilege & ")", etFullDebug

Dim szSQL As String
Dim rs As New Recordset
  
  'Generate and execute the Grant Query
  szSQL = "GRANT "
  If (Privilege And aclAll) = aclAll Then
    szSQL = szSQL & "ALL "
  Else
    If (Privilege And aclSelect) = aclSelect Then szSQL = szSQL & "SELECT, "
    If (Privilege And aclInsert) = aclInsert Then szSQL = szSQL & "INSERT, "
    If (Privilege And aclUpdate) = aclUpdate Then szSQL = szSQL & "UPDATE, "
    If (Privilege And aclDelete) = aclDelete Then szSQL = szSQL & "DELETE, "
    If (Privilege And aclRule) = aclRule Then szSQL = szSQL & "RULE, "
    If objServer.dbVersion.VersionNum >= 7.2 Then
      If (Privilege And aclReferences) = aclReferences Then szSQL = szSQL & "REFERENCES, "
      If (Privilege And aclTrigger) = aclTrigger Then szSQL = szSQL & "TRIGGER, "
    End If
    If Right(szSQL, 2) = ", " Then
      szSQL = Left(szSQL, Len(szSQL) - 2) & " "
    Else
      Exit Sub
    End If
  End If
  
  szSQL = szSQL & "ON " & FormattedID & " TO " & UserOrGroup
  objServer.ExecSQL szSQL, cnDatabase
  
  'Get the current ACL
  szSQL = "SELECT relacl FROM pg_class WHERE oid = " & lOID
  Set rs = objServer.ExecSQL(szSQL, cnDatabase)
  If Not rs.EOF Then
    szACL = rs!relacl & ""
  End If
  
  'Cleanup
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  
  Exit Sub
Err_Handler:
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  objServer.iLogError Err
End Sub

'Revoke Permissions
Public Sub Revoke(ByVal UserOrGroup As String, ByVal Privilege As aclPrivileges)
Attribute Revoke.VB_Description = "Revoke privileges from the View."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgView.Revoke(" & QUOTE & UserOrGroup & QUOTE & ", " & Privilege & ")", etFullDebug

Dim szSQL As String
Dim rs As New Recordset
  
  'Generate and execute the Grant Query
  szSQL = "REVOKE "
  If (Privilege And aclAll) = aclAll Then
    szSQL = szSQL & "ALL "
  Else
    If (Privilege And aclSelect) = aclSelect Then szSQL = szSQL & "SELECT, "
    If (Privilege And aclInsert) = aclInsert Then szSQL = szSQL & "INSERT, "
    If (Privilege And aclUpdate) = aclUpdate Then szSQL = szSQL & "UPDATE, "
    If (Privilege And aclDelete) = aclDelete Then szSQL = szSQL & "DELETE, "
    If (Privilege And aclRule) = aclRule Then szSQL = szSQL & "RULE, "
    If objServer.dbVersion.VersionNum >= 7.2 Then
      If (Privilege And aclReferences) = aclReferences Then szSQL = szSQL & "REFERENCES, "
      If (Privilege And aclTrigger) = aclTrigger Then szSQL = szSQL & "TRIGGER, "
    End If
    If Right(szSQL, 2) = ", " Then
      szSQL = Left(szSQL, Len(szSQL) - 2) & " "
    Else
      Exit Sub
    End If
  End If
  
  szSQL = szSQL & "ON " & FormattedID & " FROM " & UserOrGroup
  objServer.ExecSQL szSQL, cnDatabase
  
  'Get the current ACL
  szSQL = "SELECT relacl FROM pg_class WHERE oid = " & lOID
  Set rs = objServer.ExecSQL(szSQL, cnDatabase)
  If Not rs.EOF Then
    szACL = rs!relacl & ""
  End If
  
  'Cleanup
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  
  Exit Sub
Err_Handler:
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  objServer.iLogError Err
End Sub

Public Property Get iSQL(Optional szViewName As String, Optional szViewDefinition As String) As String
Attribute iSQL.VB_Description = "Returns the object's reverse engineered SQL."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgView.Property Get iSQL(" & QUOTE & szViewName & QUOTE & ", " & QUOTE & szViewDefinition & QUOTE & ")", etFullDebug
 
Dim szSQL As String
    
    'Set default values
    If szViewName = "" Then szViewName = szName
    If szViewDefinition = "" Then szViewDefinition = szDefinition
    
    szSQL = "-- View: " & FormattedID & vbCrLf & _
            "CREATE VIEW " & FormattedID & " AS " & szViewDefinition & ";" & vbCrLf
    If szACL <> "" Then szSQL = szSQL & ParseACL(FormattedID, szACL)
    If Comment <> "" Then szSQL = szSQL & "COMMENT ON VIEW " & FormattedID & " IS '" & Comment & "';" & vbCrLf

    iSQL = szSQL

Exit Property
Err_Handler:  objServer.iLogError Err
End Property

Public Property Get SQL() As String
Attribute SQL.VB_Description = "Returns the object's reverse engineered SQL."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgView.Property Get SQL()", etFullDebug
 
    SQL = iSQL

Exit Property
Err_Handler:  objServer.iLogError Err
End Property
