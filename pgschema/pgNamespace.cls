VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pgNamespace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' pgSchema - PostgreSQL Schema Objects
' Copyright (C) 2001, 2002, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence

Option Explicit

'Collections
Private colAggregates As Aggregates
Private colSequences As Sequences
Private colViews As Views
Private colTypes As Types
Private colFunctions As Functions
Private colOperators As Operators
Private colTables As Tables
Private colDomains As Domains

'Properties
Private dOID As Double
Private szName As String
Private szIdentifier As String
Private bSystemObject As Boolean
Private szOwner As String
Private szACL As String
Private cnDatabase As Connection
Private vTag As Variant

Public Property Get FormattedID() As String
Attribute FormattedID.VB_Description = "Returns a formatted identifier for use in SQL statements."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgNamespace.Property Get FormattedID()", etFullDebug

  FormattedID = fmtID(szName)
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Tag
Public Property Let Tag(ByVal vData As Variant)
Attribute Tag.VB_Description = "Set/Returns a variant for any use."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgNamespace.Property Let Tag(" & vData & ")", etFullDebug

  vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Set Tag(ByVal vData As Variant)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgNamespace.Property Set Tag(" & vData & ")", etFullDebug

  Set vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Tag() As Variant
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgNamespace.Property Get Tag()", etFullDebug

  If IsObject(vTag) Then
    Set Tag = vTag
  Else
    Tag = vTag
  End If
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set the Connection to use
Friend Property Set Connection(vData As Connection)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgNamespace.Property Set Connection(Connection: " & vData & ")", etFullDebug

  Set cnDatabase = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Get the Database Name
Public Property Get Database() As String
Attribute Database.VB_Description = "Returns the parent database name."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgNamespace.Property Get Database()", etFullDebug

Static szDB As String

  If szDB = "" Then szDB = GetDatabase(cnDatabase.ConnectionString)
  Database = szDB
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Return the Aggregates Collection
Public Property Get Aggregates() As Aggregates
Attribute Aggregates.VB_Description = "Returns a collection of pgAggregate objects."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get Aggregates()", etFullDebug

  If colAggregates Is Nothing Then
    Set colAggregates = New Aggregates
    
    'Pass the connection to the class
    Set colAggregates.Connection = cnDatabase
    colAggregates.Namespace = szName
    colAggregates.NamespaceOID = dOID
    colAggregates.Initialize
  End If
  Set Aggregates = colAggregates
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Return the Tables Collection
Public Property Get Tables() As Tables
Attribute Tables.VB_Description = "Returns a a collection of pgTable objects."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get Tables()", etFullDebug

  If colTables Is Nothing Then
    Set colTables = New Tables
    
    'Pass the connection to the class
    Set colTables.Connection = cnDatabase
    colTables.Namespace = szName
    colTables.NamespaceOID = dOID
    colTables.Initialize
  End If
  Set Tables = colTables
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Return the Operators Collection
Public Property Get Operators() As Operators
Attribute Operators.VB_Description = "Returns a collection of pgOperator objects."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get Operators()", etFullDebug

  If colOperators Is Nothing Then
    Set colOperators = New Operators
        
    'Pass the connection to the class
    Set colOperators.Connection = cnDatabase
    colOperators.Namespace = szName
    colOperators.NamespaceOID = dOID
    colOperators.Initialize
  End If
  Set Operators = colOperators
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Return the Functions Collection
Public Property Get Functions() As Functions
Attribute Functions.VB_Description = "Returns a collection of pgFunction objects."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get Functions()", etFullDebug

  If colFunctions Is Nothing Then
    Set colFunctions = New Functions
        
    'Pass the connection to the class
    Set colFunctions.Connection = cnDatabase
    colFunctions.Namespace = szName
    colFunctions.NamespaceOID = dOID
    colFunctions.Initialize
  End If
  Set Functions = colFunctions
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Return the Types Collection
Public Property Get Types() As Types
Attribute Types.VB_Description = "Returns a collection of pgType objects."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get Types()", etFullDebug

  If colTypes Is Nothing Then
    Set colTypes = New Types
        
    'Pass the connection to the class
    Set colTypes.Connection = cnDatabase
    colTypes.Namespace = szName
    colTypes.NamespaceOID = dOID
    colTypes.Initialize
  End If
  Set Types = colTypes
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Return the Views Collection
Public Property Get Views() As Views
Attribute Views.VB_Description = "Returns a collection of pgView objects."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get Views()", etFullDebug

  If colViews Is Nothing Then
    Set colViews = New Views
        
    'Pass the connection to the class
    Set colViews.Connection = cnDatabase
    colViews.Namespace = szName
    colViews.NamespaceOID = dOID
    colViews.Initialize
  End If
  Set Views = colViews
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Return the Sequences Collection
Public Property Get Sequences() As Sequences
Attribute Sequences.VB_Description = "Returns a collection of pgSequence objects."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get Sequences()", etFullDebug

  If colSequences Is Nothing Then
    Set colSequences = New Sequences
    
    'Pass the connection to the class
    Set colSequences.Connection = cnDatabase
    colSequences.Namespace = szName
    colSequences.NamespaceOID = dOID
    colSequences.Initialize
  End If
  Set Sequences = colSequences
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Return the Domains Collection
Public Property Get Domains() As Domains
Attribute Domains.VB_Description = "Returns a collection of pgDomain objects."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get Domains()", etFullDebug

  If colDomains Is Nothing Then
    Set colDomains = New Domains
        
    'Pass the connection to the class
    Set colDomains.Connection = cnDatabase
    colDomains.Namespace = szName
    colDomains.NamespaceOID = dOID
    colDomains.Initialize
  End If
  Set Domains = colDomains
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the OID
Friend Property Let OID(ByVal vData As Double)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgNamespace.Property Let OID(" & vData & ")", etFullDebug

  dOID = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get OID() As Double
Attribute OID.VB_Description = "Returns the Namespace OID."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgNamespace.Property Get OID()", etFullDebug

  OID = dOID
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

Friend Property Let Name(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgNamespace.Property Let Name(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szName = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Name() As String
Attribute Name.VB_Description = "Returns the Namespace name."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgNamespace.Property Get Name()", etFullDebug

  Name = szName
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Identifier
Friend Property Let Identifier(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgNamespace.Property Let Identifier(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szIdentifier = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Identifier() As String
Attribute Identifier.VB_Description = "Returns the object identifier."
Attribute Identifier.VB_MemberFlags = "200"
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgNamespace.Property Get Identifier()", etFullDebug

  Identifier = szIdentifier
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Return the Object Type
Public Property Get ObjectType() As String
Attribute ObjectType.VB_Description = "Returns the object type."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgNamespace.Property Get ObjectType()", etFullDebug

  ObjectType = "Schema"
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return a boolean indicating whether this is a System Object
Friend Property Let SystemObject(ByVal vData As Boolean)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgNamespace.Property Let SystemObject(" & vData & ")", etFullDebug

  bSystemObject = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get SystemObject() As Boolean
Attribute SystemObject.VB_Description = "Indicates whether this is a system object."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgNamespace.Property Get SystemObject()", etFullDebug

  SystemObject = bSystemObject
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Owner
Friend Property Let Owner(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Let Owner(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szOwner = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Owner() As String
Attribute Owner.VB_Description = "Returns the Namespace Owner."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get Owner()", etFullDebug

  Owner = szOwner
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the ACL
Friend Property Let ACL(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgNamespace.Property Let ACL(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szACL = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get ACL() As String
Attribute ACL.VB_Description = "Returns the Namespace ACL."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgNamespace.Property Get ACL()", etFullDebug

  ACL = szACL
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Comment
Public Property Let Comment(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgNamespace.Property Let Comment(" & vData & ")", etFullDebug

Dim szSQL As String

  If objVersion.VersionNum < 7.3 Then
    Err.Raise vbObjectError + 7, App.Title, "You cannot modify dummy schemas in PostgreSQL versions prior to 7.3."
    Exit Property
  End If
  
  szSQL = "COMMENT ON SCHEMA " & FormattedID & " IS '" & dbSZ(vData) & "'"
  objServer.ExecSQL szSQL, cnDatabase
  objServer.Databases(Database).KillCommentCache
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Comment() As String
Attribute Comment.VB_Description = "Sets/Returns the Namespace comment."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgNamespace.Property Get Comment()", etFullDebug

  If objVersion.VersionNum >= 7.3 Then
    Comment = objServer.Databases(Database).LookupComment(dOID)
  Else
    Comment = "This is a dummy schema for backwards compatibility with PostgreSQL 7.2 & below."
  End If
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Grant Permissions
Public Sub Grant(ByVal UserOrGroup As String, ByVal Privilege As aclPrivileges)
Attribute Grant.VB_Description = "Grant privileges on the Namespace."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgNamespace.Grant(" & QUOTE & UserOrGroup & QUOTE & ", " & Privilege & ")", etFullDebug

Dim szSQL As String
Dim rs As New Recordset
  
  If objVersion.VersionNum < 7.3 Then
    Err.Raise vbObjectError + 7, App.Title, "You cannot modify dummy schemas in PostgreSQL versions prior to 7.3."
    Exit Sub
  End If
  
  'Generate and execute the Grant Query
  szSQL = "GRANT "
  If (Privilege And aclAll) = aclAll Then
    szSQL = szSQL & "ALL "
  Else
    If (Privilege And aclCreate) = aclCreate Then szSQL = szSQL & "CREATE, "
    If (Privilege And aclUsage) = aclUsage Then szSQL = szSQL & "USAGE, "
    If Right(szSQL, 2) = ", " Then
      szSQL = Left(szSQL, Len(szSQL) - 2) & " "
    Else
      Exit Sub
    End If
  End If
  
  szSQL = szSQL & "ON SCHEMA " & FormattedID & " TO " & UserOrGroup
  objServer.ExecSQL szSQL, cnDatabase
  
  'Get the current ACL
  szSQL = "SELECT nspacl FROM pg_namespace WHERE oid = " & dOID & "::oid"
  Set rs = objServer.ExecSQL(szSQL, cnDatabase)
  If Not rs.EOF Then
    szACL = rs!nspacl & ""
  End If
  
  'Cleanup
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  
  Exit Sub
Err_Handler:
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  objServer.iLogError Err.Number, Err.Description
End Sub

'Revoke Permissions
Public Sub Revoke(ByVal UserOrGroup As String, ByVal Privilege As aclPrivileges)
Attribute Revoke.VB_Description = "Revoke privileges from the Namespace."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgNamespace.Revoke(" & QUOTE & UserOrGroup & QUOTE & ", " & Privilege & ")", etFullDebug

Dim szSQL As String
Dim rs As New Recordset

  If objVersion.VersionNum < 7.3 Then
    Err.Raise vbObjectError + 7, App.Title, "You cannot modify dummy schemas in PostgreSQL versions prior to 7.3."
    Exit Sub
  End If
  
  'Generate and execute the Grant Query
  szSQL = "REVOKE "
  If (Privilege And aclAll) = aclAll Then
    szSQL = szSQL & "ALL "
  Else
    If (Privilege And aclCreate) = aclCreate Then szSQL = szSQL & "CREATE, "
    If (Privilege And aclUsage) = aclUsage Then szSQL = szSQL & "USAGE, "
    If Right(szSQL, 2) = ", " Then
      szSQL = Left(szSQL, Len(szSQL) - 2) & " "
    Else
      Exit Sub
    End If
  End If
  
  szSQL = szSQL & "ON SCHEMA " & FormattedID & " FROM " & UserOrGroup
  objServer.ExecSQL szSQL, cnDatabase
  
  'Get the current ACL
  szSQL = "SELECT nspacl FROM pg_namespace WHERE oid = " & dOID & "::oid"
  Set rs = objServer.ExecSQL(szSQL, cnDatabase)
  If Not rs.EOF Then
    szACL = rs!nspacl & ""
  End If
  
  'Cleanup
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  
  Exit Sub
Err_Handler:
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  objServer.iLogError Err.Number, Err.Description
End Sub

Public Property Get SQL() As String
Attribute SQL.VB_Description = "Returns the object's reverse engineered SQL."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgNamespace.Property Get SQL()", etFullDebug
 
Dim szSQL As String
    
    If objVersion.VersionNum >= 7.3 Then
      szSQL = "-- Schema: " & szName & vbCrLf
      szSQL = szSQL & "CREATE SCHEMA " & FormattedID & " AUTHORIZATION " & fmtID(szOwner) & ";" & vbCrLf
      If szACL <> "" Then szSQL = szSQL & ParseACL(FormattedID, szACL, aclSchema)
      If Comment <> "" Then szSQL = szSQL & "COMMENT ON SCHEMA " & FormattedID & " IS '" & Comment & "';" & vbCrLf
    End If
      
    SQL = szSQL

Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Cleanup
Private Sub Class_Terminate()
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Class_Terminate()", etFullDebug

  Set colAggregates = Nothing
  Set colSequences = Nothing
  Set colViews = Nothing
  Set colTypes = Nothing
  Set colFunctions = Nothing
  Set colOperators = Nothing
  Set colTables = Nothing
  Set colDomains = Nothing
  
  Exit Sub
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Sub
