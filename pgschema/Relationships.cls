VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Relationships"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' pgSchema - PostgreSQL Schema Objects
' Copyright (C) 2001, 2002, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence

Option Explicit

'Private Variables
Private mCol As Collection
Private cnDatabase As Connection
Private szLocCols As String
Private szRefCols As String
Private lNamespaceOID As Long
Private szNamespace As String

'Set/Return the Namespace OID
Friend Property Let NamespaceOID(ByVal vData As Long)
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Relationships.Property Let NamespaceOID(" & vData & ")", etFullDebug

  lNamespaceOID = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get NamespaceOID() As Long
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Relationships.Property Get NamespaceOID()", etFullDebug

  NamespaceOID = lNamespaceOID
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Namespace Name
Friend Property Let Namespace(ByVal vData As String)
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Relationships.Property Let Namespace(" & vData & ")", etFullDebug

  szNamespace = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Namespace() As String
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Relationships.Property Get Namespace()", etFullDebug

  Namespace = szNamespace
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property


'Does the specified item exist?
Public Function Exists(Identifier As String) As Boolean
Attribute Exists.VB_Description = "Does the specified object exist?"
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Relationships.Exists(" & QUOTE & Identifier & QUOTE & ")", etFullDebug

Dim objRelationship As pgRelationship

  For Each objRelationship In mCol
    If objRelationship.Identifier = Identifier Then
      Exists = True
      Exit For
    End If
  Next objRelationship
  
  Exit Function
Err_Handler:  objServer.iLogError Err
End Function

'Get the Database Name
Public Property Get Database() As String
Attribute Database.VB_Description = "Returns the parent database name."
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Relationships.Property Get Database()", etFullDebug

Static szDB As String

  If szDB = "" Then szDB = GetDatabase(cnDatabase.ConnectionString)
  Database = szDB
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set the Connection to use
Friend Property Set Connection(vData As Connection)
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Relationships.Property Set Connection(Connection: " & vData & ")", etFullDebug

  Set cnDatabase = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Internal Add
Private Function iAdd(szLocalColumn As String, szReferencedColumn As String) As pgRelationship
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Relationships.iAdd(" & QUOTE & szLocalColumn & QUOTE & ", " & QUOTE & szReferencedColumn & QUOTE & ")", etFullDebug

Dim objNewMember As pgRelationship

  'Create a new object
  Set objNewMember = New pgRelationship
  objNewMember.NamespaceOID = lNamespaceOID
  objNewMember.Namespace = szNamespace
  objNewMember.LocalColumn = szLocalColumn
  objNewMember.ReferencedColumn = szReferencedColumn
  objNewMember.Identifier = szLocalColumn & " = " & szReferencedColumn
  Set objNewMember.Connection = cnDatabase
  mCol.Add objNewMember, ULEncode(objNewMember.Identifier)
  Set iAdd = objNewMember
  
  Exit Function
Err_Handler:  objServer.iLogError Err
End Function

'Return the specified Item
Public Property Get Item(Key As String) As pgOperator
Attribute Item.VB_Description = "Returns a Relationship object."
Attribute Item.VB_UserMemId = 0
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Relationships.Property Get Item(" & QUOTE & Key & QUOTE & ")", etFullDebug

  Set Item = mCol(ULEncode(Key))
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the Item count
Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns a count of the Relationships on the current Foreign Key."
Attribute Count.VB_MemberFlags = "200"
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Relationships.Property Get Count()", etFullDebug

  Count = mCol.Count
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Allow enumeration using For...Each
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Relationships.Property Get NewEnum()", etFullDebug

  Set NewEnum = mCol.[_NewEnum]
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Refresh the collection
Public Sub Refresh()
Attribute Refresh.VB_Description = "Refrsh the collection."
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Relationships.Refresh()", etFullDebug

  'Kill the caches first
  objServer.Databases(Database).KillCommentCache
  objServer.Databases(Database).KillLanguageCache
  objServer.Databases(Database).KillOperatorCache
  objServer.Databases(Database).KillTypeCache
  
  'Initialize will basically do a refresh so let's use that!
  Initialize szLocCols, szRefCols
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub

'Initialise the collection if possible
Friend Sub Initialize(szLocalColumns As String, szReferencedColumns As String)
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Relationships.Initialize(" & QUOTE & szLocalColumns & QUOTE & ", " & QUOTE & szReferencedColumns & QUOTE & ")", etFullDebug

Dim szLocal() As String
Dim szReferenced() As String
Dim X As Long

  szLocCols = szLocalColumns
  szRefCols = szReferencedColumns
    
  'Create a new collection
  Set mCol = New Collection
  
  'Get the data and loop through it
  szLocal = Split(szLocalColumns, "|")
  szReferenced = Split(szReferencedColumns, "|")
  For X = 0 To UBound(szLocal) - 1
    iAdd szLocal(X), szReferenced(X)
  Next X
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub

'Cleanup
Private Sub Class_Terminate()
'On Error Goto Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Relationships.Class_Terminate()", etFullDebug

  Set mCol = Nothing
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub
