VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pgType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' pgSchema - PostgreSQL Schema Objects
' Copyright (C) 2001, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence

Option Explicit

'Properties
Private lOID As Long
Private szName As String
Private szIdentifier As String
Private bSystemObject As Boolean
Private szOwner As String
Private szInputFunction As String
Private szOutputFunction As String
Private lInternalLength As Long
Private lExternalLength As Long
Private szDefault As String
Private szElement As String
Private szDelimiter As String
Private szSendFunction As String
Private szReceiveFunction As String
Private bPassedByValue As Boolean
Private szAlignment As String
Private szStorage As String
Private cnDatabase As Connection
Private dTimeStamp As Date

'History
Private colHistory As History

'Return the History collection.
Public Property Get History() As History
'On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Get History()", etFullDebug

  If objServer.Databases(Database).RevisionControl Then
    If colHistory Is Nothing Then
      Set colHistory = New History
     
      'Pass the options to the class
      Set colHistory.Connection = cnDatabase
      colHistory.ObjIdentifier = szIdentifier
      colHistory.objType = ObjectType
      colHistory.Initialize
    End If
    Set History = colHistory
  End If
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the TimeStamp
Friend Property Let TimeStamp(ByVal vData As Date)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Let TimeStamp(" & vData & ")", etFullDebug

  dTimeStamp = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get TimeStamp() As Date
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Get TimeStamp()", etFullDebug

  TimeStamp = dTimeStamp
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set the Connection to use
Friend Property Set Connection(vData As Connection)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Set Connection(Connection: " & vData & ")", etFullDebug

  Set cnDatabase = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Get the Database Name
Public Property Get Database() As String
Attribute Database.VB_Description = "Returns the parent database name."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Get Database()", etFullDebug

  Database = GetDatabase(cnDatabase.ConnectionString)
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the OID
Friend Property Let OID(ByVal vData As Long)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Let OID(" & vData & ")", etFullDebug

  lOID = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get OID() As Long
Attribute OID.VB_Description = "Returns the Type OID."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Get OID()", etFullDebug

  OID = lOID
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Name
Friend Property Let Name(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Let Name(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szName = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Name() As String
Attribute Name.VB_Description = "Returns the Type name."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Get Name()", etFullDebug

  Name = szName
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Identifier
Friend Property Let Identifier(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Let Identifier(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szIdentifier = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Identifier() As String
Attribute Identifier.VB_Description = "Returns the onject identifier."
Attribute Identifier.VB_MemberFlags = "200"
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Get Identifier()", etFullDebug

  Identifier = szIdentifier
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the Object Type
Public Property Get ObjectType() As String
Attribute ObjectType.VB_Description = "Returns the object type."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Get ObjectType()", etFullDebug

  ObjectType = "Type"
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return a boolean indicating whether this is a System Object
Friend Property Let SystemObject(ByVal vData As Boolean)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Let SystemObject(" & vData & ")", etFullDebug

  bSystemObject = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get SystemObject() As Boolean
Attribute SystemObject.VB_Description = "Indicates whether this is a system object."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Get SystemObject()", etFullDebug

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
Attribute Owner.VB_Description = "Returns the Type's owner."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get Owner()", etFullDebug

  Owner = szOwner
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the InputFunction
Friend Property Let InputFunction(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Let InputFunction(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szInputFunction = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get InputFunction() As String
Attribute InputFunction.VB_Description = "Returns the Input Function name used by the Type."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Get InputFunction()", etFullDebug

  InputFunction = szInputFunction
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the OutputFunction
Friend Property Let OutputFunction(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Let OutputFunction(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szOutputFunction = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get OutputFunction() As String
Attribute OutputFunction.VB_Description = "Returns the Output Function name used by the Type."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Get OutputFunction()", etFullDebug

  OutputFunction = szOutputFunction
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the InternalLength
Friend Property Let InternalLength(ByVal vData As Long)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Let InternalLength(" & vData & ")", etFullDebug

  lInternalLength = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get InternalLength() As Long
Attribute InternalLength.VB_Description = "Returns the Internal Length of the data Type."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Get InternalLength()", etFullDebug

  InternalLength = lInternalLength
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the ExternalLength
Friend Property Let ExternalLength(ByVal vData As Long)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Let ExternalLength(" & vData & ")", etFullDebug

  lExternalLength = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get ExternalLength() As Long
Attribute ExternalLength.VB_Description = "Returns the External Length of the data Type."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Get ExternalLength()", etFullDebug

  ExternalLength = lExternalLength
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Default
Friend Property Let Default(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Let Default(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szDefault = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Default() As String
Attribute Default.VB_Description = "Returns the Default value for the Type."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Get Default()", etFullDebug

  Default = szDefault
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Element
Friend Property Let Element(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Let Element(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szElement = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Element() As String
Attribute Element.VB_Description = "Returns the Element Type if this is an array Type."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Get Element()", etFullDebug

  Element = szElement
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Delimiter
Friend Property Let Delimiter(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Let Delimiter(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szDelimiter = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Delimiter() As String
Attribute Delimiter.VB_Description = "Returns the delimiter used for arrays of this Type."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Get Delimiter()", etFullDebug

  Delimiter = szDelimiter
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the SendFunction
Friend Property Let SendFunction(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Let SendFunction(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szSendFunction = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get SendFunction() As String
Attribute SendFunction.VB_Description = "Returns the Send Function name used by the Type."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Get SendFunction()", etFullDebug

  SendFunction = szSendFunction
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the ReceiveFunction
Friend Property Let ReceiveFunction(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Let ReceiveFunction(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szReceiveFunction = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get ReceiveFunction() As String
Attribute ReceiveFunction.VB_Description = "Returns the Receive Function name used by the Type."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Get ReceiveFunction()", etFullDebug

  ReceiveFunction = szReceiveFunction
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return PassedByValue?
Friend Property Let PassedByValue(ByVal vData As Boolean)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Let PassedByValue(" & vData & ")", etFullDebug

  bPassedByValue = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get PassedByValue() As Boolean
Attribute PassedByValue.VB_Description = "Returns a value indicating whether the Type is Passed By Value (or By Reference if False)."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Get PassedByValue()", etFullDebug

  PassedByValue = bPassedByValue
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Alignment
Friend Property Let Alignment(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Let Alignment(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szAlignment = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Alignment() As String
Attribute Alignment.VB_Description = "Returns trhe Alignment of the Type."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Get Alignment()", etFullDebug

  Alignment = szAlignment
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Storage
Friend Property Let Storage(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Let Storage(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szStorage = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Storage() As String
Attribute Storage.VB_Description = "Returns the Type storage method (PostgreSQL 7.1+)."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Get Storage()", etFullDebug

  Storage = szStorage
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Comment
Public Property Let Comment(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Let Comment(" & vData & ")", etFullDebug

Dim szSQL As String

  szSQL = "COMMENT ON TYPE " & QUOTE & szName & QUOTE & " IS '" & dbSZ(vData) & "'"
  objServer.ExecSQL szSQL, cnDatabase
  objServer.Databases(GetDatabase(cnDatabase.ConnectionString)).KillCommentCache
  If Not SystemObject Then Commit rcUpdate, "Updated object comment."
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Comment() As String
Attribute Comment.VB_Description = "Sets/Returns the Type comment."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Get Comment()", etFullDebug

  Comment = objServer.Databases(GetDatabase(cnDatabase.ConnectionString)).LookupComment(lOID)
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

Public Property Get SQL() As String
Attribute SQL.VB_Description = "Returns the object's reverse engineered SQL."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Get SQL()", etFullDebug
Dim szSQL As String
  
  szSQL = "-- Type: " & szName & vbCrLf & _
          "CREATE TYPE " & QUOTE & szName & QUOTE & "(" & _
          "INPUT = " & QUOTE & szInputFunction & QUOTE & ", " & _
          "OUTPUT = " & QUOTE & szOutputFunction & QUOTE
        
  If lInternalLength = 0 Then
    szSQL = szSQL & ", INTERNALLENGTH = VARIABLE"
  Else
    szSQL = szSQL & ", INTERNALLENGTH = " & lInternalLength
  End If
          
  If lExternalLength = 0 Then
    szSQL = szSQL & ", EXTERNALLENGTH = VARIABLE"
  Else
    szSQL = szSQL & ", EXTERNALLENGTH = " & lExternalLength
  End If
  
  If szDefault <> "" Then szSQL = szSQL & ", DEFAULT = " & QUOTE & dbSZ(szDefault) & QUOTE
  If szElement <> "" Then szSQL = szSQL & ", ELEMENT = " & QUOTE & szElement & QUOTE
  If szDelimiter <> "" Then szSQL = szSQL & ", DELIMITER = " & QUOTE & szDelimiter & QUOTE
  If szSendFunction <> "" Then szSQL = szSQL & ", SEND = " & QUOTE & szSendFunction & QUOTE
  If szReceiveFunction <> "" Then szSQL = szSQL & ", RECEIVE = " & QUOTE & szReceiveFunction & QUOTE
  If bPassedByValue Then szSQL = szSQL & ", PASSEDBYVALUE"
  If szAlignment <> "" Then szSQL = szSQL & ", ALIGNMENT = " & szAlignment
  If szStorage <> "" Then szSQL = szSQL & ", STORAGE = " & szStorage
  
  szSQL = szSQL & ");" & vbCrLf
          
  If Comment <> "" Then szSQL = szSQL & "COMMENT ON TYPE " & QUOTE & szName & QUOTE & " IS '" & dbSZ(Comment) & "';" & vbCrLf
    
  SQL = szSQL
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Property

Public Sub Commit(UpdateType As rcLogType, UpdateComment As String)
Attribute Commit.VB_Description = "Commit object changes."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Commit(" & UpdateType & ", " & QUOTE & UpdateComment & QUOTE & ")", etFullDebug

  objServer.Databases(Database).LogRevision ObjectType, Identifier, OID, SQL, UpdateComment, UpdateType
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub

Public Property Get RCStatus() As rcStat
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.RCStatus()", etFullDebug

Dim szSQL As String
Dim rs As New Recordset

  If Not objServer.Databases(Database).RevisionControl Then
    RCStatus = rcNotApplicable
    Exit Property
  End If
  
  szSQL = "SELECT rc_timestamp, rc_definition FROM pgadmin_rclog WHERE rc_type = '" & dbSZ(ObjectType) & "' AND rc_identifier = '" & dbSZ(szIdentifier) & "' ORDER BY rc_version DESC LIMIT 1"
  Set rs = objServer.ExecSQL(szSQL, cnDatabase)
  If rs.EOF Then
    RCStatus = rcNotInRC
  Else
    If SQL = rs!rc_definition & "" Then
      RCStatus = rcUpToDate
    Else
      If rs!rc_timestamp > dTimeStamp Then
        RCStatus = rcNeedsRefresh
      Else
        RCStatus = rcOutOfDate
      End If
    End If
  End If
  
  Set rs = Nothing
  
  Exit Property
Err_Handler:
  Set rs = Nothing
  objServer.iLogError Err
End Property


