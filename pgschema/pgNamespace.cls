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
Private lOID As Long
Private szName As String
Private szIdentifier As String
Private bSystemObject As Boolean
Private szOwner As String
Private szACL As String
Private cnDatabase As Connection
Private vTag As Variant

'Set/Return the Tag
Friend Property Let Tag(ByVal vData As Variant)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgNamespace.Property Let Tag(" & vData & ")", etFullDebug

  vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Tag() As Variant
Attribute Tag.VB_Description = "Set/Returns a variant for any use."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgNamespace.Property Get Tag()", etFullDebug

  Tag = vTag
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set the Connection to use
Friend Property Set Connection(vData As Connection)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgNamespace.Property Set Connection(Connection: " & vData & ")", etFullDebug

  Set cnDatabase = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err
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
Err_Handler:  objServer.iLogError Err
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
    colAggregates.NamespaceOID = lOID
    colAggregates.Initialize
  End If
  Set Aggregates = colAggregates
    
  Exit Property
Err_Handler:  objServer.iLogError Err
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
    colTables.NamespaceOID = lOID
    colTables.Initialize
  End If
  Set Tables = colTables
    
  Exit Property
Err_Handler:  objServer.iLogError Err
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
    colOperators.NamespaceOID = lOID
    colOperators.Initialize
  End If
  Set Operators = colOperators
    
  Exit Property
Err_Handler:  objServer.iLogError Err
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
    colFunctions.NamespaceOID = lOID
    colFunctions.Initialize
  End If
  Set Functions = colFunctions
    
  Exit Property
Err_Handler:  objServer.iLogError Err
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
    colTypes.NamespaceOID = lOID
    colTypes.Initialize
  End If
  Set Types = colTypes
    
  Exit Property
Err_Handler:  objServer.iLogError Err
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
    colViews.NamespaceOID = lOID
    colViews.Initialize
  End If
  Set Views = colViews
    
  Exit Property
Err_Handler:  objServer.iLogError Err
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
    colSequences.NamespaceOID = lOID
    colSequences.Initialize
  End If
  Set Sequences = colSequences
    
  Exit Property
Err_Handler:  objServer.iLogError Err
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
    colDomains.NamespaceOID = lOID
    colDomains.Initialize
  End If
  Set Domains = colDomains
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the OID
Friend Property Let OID(ByVal vData As Long)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgNamespace.Property Let OID(" & vData & ")", etFullDebug

  lOID = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get OID() As Long
Attribute OID.VB_Description = "Returns the Namespace OID."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgNamespace.Property Get OID()", etFullDebug

  OID = lOID
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

Friend Property Let Name(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgNamespace.Property Let Name(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szName = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Name() As String
Attribute Name.VB_Description = "Returns the Namespace name."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgNamespace.Property Get Name()", etFullDebug

  Name = szName
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Identifier
Friend Property Let Identifier(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgNamespace.Property Let Identifier(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szIdentifier = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Identifier() As String
Attribute Identifier.VB_Description = "Returns the object identifier."
Attribute Identifier.VB_MemberFlags = "200"
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgNamespace.Property Get Identifier()", etFullDebug

  Identifier = szIdentifier
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the Object Type
Public Property Get ObjectType() As String
Attribute ObjectType.VB_Description = "Returns the object type."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgNamespace.Property Get ObjectType()", etFullDebug

  ObjectType = "Schema"
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return a boolean indicating whether this is a System Object
Friend Property Let SystemObject(ByVal vData As Boolean)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgNamespace.Property Let SystemObject(" & vData & ")", etFullDebug

  bSystemObject = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get SystemObject() As Boolean
Attribute SystemObject.VB_Description = "Indicates whether this is a system object."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgNamespace.Property Get SystemObject()", etFullDebug

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
Attribute Owner.VB_Description = "Returns the Namespace Owner."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get Owner()", etFullDebug

  Owner = szOwner
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the ACL
Friend Property Let ACL(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgNamespace.Property Let ACL(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szACL = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get ACL() As String
Attribute ACL.VB_Description = "Returns the Namespace ACL."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgNamespace.Property Get ACL()", etFullDebug

  ACL = szACL
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Comment
Public Property Let Comment(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgNamespace.Property Let Comment(" & vData & ")", etFullDebug

Dim szSQL As String

  If objServer.dbVersion.VersionNum < 7.3 Then
    Err.Raise 5432007, App.Title, "You cannot modify dummy schemas in PostgreSQL versions prior to 7.3."
    Exit Property
  End If
  
  szSQL = "COMMENT ON SCHEMA " & QUOTE & szName & QUOTE & " IS '" & dbSZ(vData) & "'"
  objServer.ExecSQL szSQL, cnDatabase
  objServer.Databases(Database).KillCommentCache
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Comment() As String
Attribute Comment.VB_Description = "Sets/Returns the Namespace comment."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgNamespace.Property Get Comment()", etFullDebug

  If objServer.dbVersion.VersionNum >= 7.3 Then
    Comment = objServer.Databases(Database).LookupComment(lOID)
  Else
    Comment = "This is a dummy schema for backwards compatibility with PostgreSQL 7.2 & below."
  End If
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Grant Permissions
Public Sub Grant(ByVal UserOrGroup As String, ByVal Privilege As aclPrivileges)
Attribute Grant.VB_Description = "Grant privileges on the Namespace."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgNamespace.Grant(" & QUOTE & UserOrGroup & QUOTE & ", " & Privilege & ")", etFullDebug

Dim szSQL As String
Dim rs As New Recordset
  
  If objServer.dbVersion.VersionNum < 7.3 Then
    Err.Raise 5432007, App.Title, "You cannot modify dummy schemas in PostgreSQL versions prior to 7.3."
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
  
  szSQL = szSQL & "ON SCHEMA " & QUOTE & szName & QUOTE & " TO " & UserOrGroup
  objServer.ExecSQL szSQL, cnDatabase
  
  'Get the current ACL
  szSQL = "SELECT nspacl FROM pg_namespace WHERE oid = " & lOID
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
  objServer.iLogError Err
End Sub

'Revoke Permissions
Public Sub Revoke(ByVal UserOrGroup As String, ByVal Privilege As aclPrivileges)
Attribute Revoke.VB_Description = "Revoke privileges from the Namespace."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgNamespace.Revoke(" & QUOTE & UserOrGroup & QUOTE & ", " & Privilege & ")", etFullDebug

Dim szSQL As String
Dim rs As New Recordset

  If objServer.dbVersion.VersionNum < 7.3 Then
    Err.Raise 5432007, App.Title, "You cannot modify dummy schemas in PostgreSQL versions prior to 7.3."
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
  
  szSQL = szSQL & "ON SCHEMA " & QUOTE & szName & QUOTE & " FROM " & UserOrGroup
  objServer.ExecSQL szSQL, cnDatabase
  
  'Get the current ACL
  szSQL = "SELECT nspacl FROM pg_namespace WHERE oid = " & lOID
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
  objServer.iLogError Err
End Sub

Public Property Get SQL() As String
Attribute SQL.VB_Description = "Returns the object's reverse engineered SQL."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgNamespace.Property Get SQL()", etFullDebug
 
Dim szSQL As String
    
    szSQL = "-- Schema: " & szName & vbCrLf
    szSQL = szSQL & "CREATE SCHEMA " & QUOTE & szName & QUOTE & " AUTHORIZATION " & QUOTE & szOwner & QUOTE & ";" & vbCrLf
    If szACL <> "" Then szSQL = szSQL & ParseACL(szName, szACL, aclSchema)
    If Comment <> "" Then szSQL = szSQL & "COMMENT ON SCHEMA " & QUOTE & szName & QUOTE & " IS '" & Comment & "';" & vbCrLf
      
    SQL = szSQL

Exit Property
Err_Handler:  objServer.iLogError Err
End Property

Public Property Get Schema(Optional ResetSequences As Boolean) As String
Attribute Schema.VB_Description = "Returns the reverse engineered SQL for all the objects within this Namespace."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Property Get pgDatabase.Schema(" & ResetSequences & ")", etFullDebug
Dim szSQL As String
Dim vObject As Variant
Dim vChildObject As Variant
Dim arrObjects() As Variant
Dim X As Long
Dim Y As Long
Dim vTemp As Variant

  szSQL = "-- pgSchema:  " & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & _
          "-- Timestamp: " & Now & vbCrLf & _
          "-- Server:    " & objServer.Server & vbCrLf & _
          "-- Port:      " & objServer.Port & vbCrLf & _
          "-- Database:  " & szName & vbCrLf & vbCrLf
  
  'We will output the Schema object by object in OID order. This should work
  'most of the time unless (for example) a table is altered to use a later
  'created function as a default. Hopefully future releases of PostgreSQL
  'will include a pg_dependency table that we can use instead.
  
  'First, copy all objects into a single array: Aggregates, Functions,
  'Indexes, Languages, Operators, Rules, Sequences, Tables, Triggers,
  'Types & Views
  
  ReDim arrObjects(0)
  
  'Aggregates
  For Each vObject In Aggregates
    If Not vObject.SystemObject Then
      Set arrObjects(UBound(arrObjects)) = vObject
      ReDim Preserve arrObjects(UBound(arrObjects) + 1)
    End If
  Next vObject
  
  'Domains
  For Each vObject In Domains
    If Not vObject.SystemObject Then
      Set arrObjects(UBound(arrObjects)) = vObject
      ReDim Preserve arrObjects(UBound(arrObjects) + 1)
    End If
  Next vObject

  'Functions
  For Each vObject In Functions
    If Not vObject.SystemObject Then
      Set arrObjects(UBound(arrObjects)) = vObject
      ReDim Preserve arrObjects(UBound(arrObjects) + 1)
    End If
  Next vObject

  'Operators
  For Each vObject In Operators
    If Not vObject.SystemObject Then
      Set arrObjects(UBound(arrObjects)) = vObject
      ReDim Preserve arrObjects(UBound(arrObjects) + 1)
    End If
  Next vObject

  'Sequences
  For Each vObject In Sequences
    If Not vObject.SystemObject Then
      Set arrObjects(UBound(arrObjects)) = vObject
      ReDim Preserve arrObjects(UBound(arrObjects) + 1)
    End If
  Next vObject

  'Tables
  For Each vObject In Tables
    If Not vObject.SystemObject Then
      Set arrObjects(UBound(arrObjects)) = vObject
      ReDim Preserve arrObjects(UBound(arrObjects) + 1)
    
      'Indexes
      For Each vChildObject In vObject.Indexes
        If Not vChildObject.SystemObject Then
          Set arrObjects(UBound(arrObjects)) = vChildObject
          ReDim Preserve arrObjects(UBound(arrObjects) + 1)
        End If
      Next vChildObject
      
      'Rules
      For Each vChildObject In vObject.Rules
        If Not vChildObject.SystemObject Then
          Set arrObjects(UBound(arrObjects)) = vChildObject
          ReDim Preserve arrObjects(UBound(arrObjects) + 1)
        End If
      Next vChildObject
      
      'Triggers
      For Each vChildObject In vObject.Triggers
        If Not vChildObject.SystemObject Then
          Set arrObjects(UBound(arrObjects)) = vChildObject
          ReDim Preserve arrObjects(UBound(arrObjects) + 1)
        End If
      Next vChildObject
    End If
  Next vObject

  'Types
  For Each vObject In Types
    If Not vObject.SystemObject Then
      Set arrObjects(UBound(arrObjects)) = vObject
      ReDim Preserve arrObjects(UBound(arrObjects) + 1)
    End If
  Next vObject

  'Views
  For Each vObject In Views
    If Not vObject.SystemObject Then
      Set arrObjects(UBound(arrObjects)) = vObject
      ReDim Preserve arrObjects(UBound(arrObjects) + 1)
    End If
  Next vObject
  
  'Lose the last empty element
  If UBound(arrObjects) > 0 Then ReDim Preserve arrObjects(UBound(arrObjects) - 1)
  
  'Now bubble sort the array by OID.
  For X = UBound(arrObjects) To LBound(arrObjects) Step -1
    For Y = LBound(arrObjects) + 1 To X
      If arrObjects(Y - 1).OID > arrObjects(Y).OID Then
        Set vTemp = arrObjects(Y - 1)
        Set arrObjects(Y - 1) = arrObjects(Y)
        Set arrObjects(Y) = vTemp
      End If
    Next Y
  Next X

  'Now loop through the array and output the SQL
  For X = 0 To UBound(arrObjects)
    If arrObjects(X).ObjectType = "Sequence" Then
      szSQL = szSQL & arrObjects(X).SQL(ResetSequences) & vbCrLf
    Else
      szSQL = szSQL & arrObjects(X).SQL & vbCrLf
    End If
  Next X
    
  Schema = szSQL
  
  Exit Property
Err_Handler:  objServer.iLogError Err
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
Err_Handler:  objServer.iLogError Err
End Sub
