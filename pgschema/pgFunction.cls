VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pgFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' pgSchema - PostgreSQL Schema Objects
' Copyright (C) 2001, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence

Option Explicit

'Collections
Private colArguments As Arguments

'Properties
Private lOID As Long
Private szName As String
Private szIdentifier As String
Private bSystemObject As Boolean
Private szOwner As String
Private szReturns As String
Private szSource As String
Private szLanguage As String
Private bCachable As Boolean
Private bStrict As Boolean
Private szArgumentList As String
Private cnDatabase As Connection

'Set the Connection to use
Friend Property Set Connection(vData As Connection)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Set Connection(Connection: " & vData & ")", etFullDebug

  Set cnDatabase = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the Arguments Collection
Public Property Get Arguments() As Arguments
Attribute Arguments.VB_Description = "Returns a collection of Function Arguments."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get Arguments()", etFullDebug

  If colArguments Is Nothing Then
    Set colArguments = New Arguments
    
    'Pass the connection to the class
    Set colArguments.Connection = cnDatabase
    colArguments.Initialize szArgumentList
  End If
  Set Arguments = colArguments
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Get the Database Name
Public Property Get Database() As String
Attribute Database.VB_Description = "Returns the parent database name."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Get Database()", etFullDebug

  Database = GetDatabase(cnDatabase.ConnectionString)
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the OID
Friend Property Let OID(ByVal vData As Long)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Let OID(" & vData & ")", etFullDebug

  lOID = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get OID() As Long
Attribute OID.VB_Description = "Returns the Type OID."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Get OID()", etFullDebug

  OID = lOID
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Name
Friend Property Let Name(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Let Name(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szName = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Name() As String
Attribute Name.VB_Description = "Returns the Type name."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Get Name()", etFullDebug

  Name = szName
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Identifier
Friend Property Let Identifier(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Let Identifier(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szIdentifier = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Identifier() As String
Attribute Identifier.VB_Description = "Returns the object identifier."
Attribute Identifier.VB_MemberFlags = "200"
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Get Identifier()", etFullDebug

  Identifier = szIdentifier
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the Object Type
Public Property Get ObjectType() As String
Attribute ObjectType.VB_Description = "Returns the object type."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Get ObjectType()", etFullDebug

  ObjectType = "Function"
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return a boolean indicating whether this is a System Object
Friend Property Let SystemObject(ByVal vData As Boolean)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Let SystemObject(" & vData & ")", etFullDebug

  bSystemObject = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get SystemObject() As Boolean
Attribute SystemObject.VB_Description = "Indicates whether this is a system object."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Get SystemObject()", etFullDebug

  SystemObject = bSystemObject
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Owner
Friend Property Let Owner(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Let Owner(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szOwner = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Owner() As String
Attribute Owner.VB_Description = "Returns an object representing the Type's owner."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Get Owner()", etFullDebug

  Owner = szOwner
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Return Type
Friend Property Let Returns(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Let Returns(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szReturns = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Returns() As String
Attribute Returns.VB_Description = "Returns the Function Return data type."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Get Returns()", etFullDebug

  Returns = szReturns
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Function Source
Friend Property Let Source(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Let Source(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szSource = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Source() As String
Attribute Source.VB_Description = "Returns the Function Source. This may be the path to an object library, or Procedural Language code."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Get Source()", etFullDebug

  Source = szSource
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Language
Friend Property Let Language(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Let Language(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szLanguage = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Language() As String
Attribute Language.VB_Description = "RFeturns the Function Language."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Get Language()", etFullDebug

  Language = szLanguage
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return Cachable?
Friend Property Let Cachable(ByVal vData As Boolean)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgLanguage.Property Let Cachable(" & vData & ")", etFullDebug

  bCachable = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Cachable() As Boolean
Attribute Cachable.VB_Description = "Returns a value indicating whether the Function result may be cached."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgLanguage.Property Get Cachable()", etFullDebug

  Cachable = bCachable
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return Strict?
Friend Property Let Strict(ByVal vData As Boolean)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgLanguage.Property Let Strict(" & vData & ")", etFullDebug

  bStrict = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Strict() As Boolean
Attribute Strict.VB_Description = "Returns a value indicating whether the Function is Strict."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgLanguage.Property Get Strict()", etFullDebug

  Strict = bStrict
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Comment
Public Property Let Comment(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Let Comment(" & vData & ")", etFullDebug

Dim szSQL As String
Dim szArgs As String
Dim vArg As Variant
Dim X As Long

  'Get the arguments if we don't have them
  If colArguments Is Nothing Then
    Set colArguments = New Arguments
    
    'Pass the connection to the class
    Set colArguments.Connection = cnDatabase
    colArguments.Initialize szArgumentList
  End If
  
  szArgs = "("
  For Each vArg In colArguments
    szArgs = szArgs & QUOTE & vArg & QUOTE & ", "
  Next vArg
  If Len(szArgs) > 3 Then szArgs = Left(szArgs, Len(szArgs) - 2)
  szArgs = szArgs & ")"

  szSQL = "COMMENT ON FUNCTION " & QUOTE & szName & QUOTE & szArgs & " IS '" & dbSZ(vData) & "'"
  objServer.ExecSQL szSQL, cnDatabase
  objServer.Databases(Database).KillCommentCache
  If Not SystemObject Then Commit rlUpdate, "Updated object comment."
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Comment() As String
Attribute Comment.VB_Description = "Sets/Returns the Type comment."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Get Comment()", etFullDebug

  Comment = objServer.Databases(Database).LookupComment(lOID)
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the ArgumentList
Friend Property Let ArgumentList(ByVal vData As String)
Attribute ArgumentList.VB_Description = "Returns the Function Arguments."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Let ArgumentList(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szArgumentList = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

Public Property Get SQL() As String
Attribute SQL.VB_Description = "Returns the object's reverse engineered SQL."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Property Get SQL()", etFullDebug
Dim szSQL As String
Dim szArgs As String
Dim vArg As Variant
Dim X As Long
  'Get the arguments if we don't have them
  If colArguments Is Nothing Then
    Set colArguments = New Arguments
    
    'Pass the connection to the class
    Set colArguments.Connection = cnDatabase
    colArguments.Initialize szArgumentList
  End If
  
  szArgs = "("
  For Each vArg In colArguments
    szArgs = szArgs & QUOTE & vArg & QUOTE & ", "
  Next vArg
  If Len(szArgs) > 3 Then szArgs = Left(szArgs, Len(szArgs) - 2)
  szArgs = szArgs & ")"
  
  szSQL = "-- Function: " & szIdentifier & vbCrLf & _
          "CREATE FUNCTION " & QUOTE & szName & QUOTE & szArgs & _
          " RETURNS " & QUOTE & szReturns & QUOTE & _
          " AS '" & szSource & "'" & _
          " LANGUAGE '" & dbSZ(szLanguage) & "'"
          
  If bCachable Or bStrict Then szSQL = szSQL & " WITH ("
  If bCachable Then szSQL = szSQL & "ISCACHABLE, "
  If bStrict Then szSQL = szSQL & "ISSTRICT, "
  If bCachable Or bStrict Then szSQL = Mid(szSQL, 1, Len(szSQL) - 2) & ")"
  szSQL = szSQL & ";" & vbCrLf
  
  If Comment <> "" Then szSQL = szSQL & "COMMENT ON FUNCTION " & QUOTE & szName & QUOTE & szArgs & " IS '" & dbSZ(Comment) & "';" & vbCrLf
    
  SQL = szSQL
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Property

Public Sub Commit(UpdateType As rlLogType, UpdateComment As String)
Attribute Commit.VB_Description = "Commit object changes."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgFunction.Commit(" & UpdateType & ", " & QUOTE & UpdateComment & QUOTE & ")", etFullDebug

  objServer.Databases(Database).LogRevision ObjectType, Identifier, SQL, UpdateComment, UpdateType
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub

