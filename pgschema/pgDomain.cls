VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pgDomain"
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
Private szBaseType As String
Private lLength As Long
Private lNumericScale As Long
Private szDefault As String
Private bNotNull As Boolean
Private cnDatabase As Connection
Private dTimeStamp As Date

'History
Private colHistory As History

'Return the History collection.
Public Property Get History() As History
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDomain.Property Get History()", etFullDebug

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
objServer.iLogEvent "Entering " & App.Title & ":pgDomain.Property Let TimeStamp(" & vData & ")", etFullDebug

  dTimeStamp = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get TimeStamp() As Date
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDomain.Property Get TimeStamp()", etFullDebug

  TimeStamp = dTimeStamp
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set the Connection to use
Friend Property Set Connection(vData As Connection)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDomain.Property Set Connection(Connection: " & vData & ")", etFullDebug

  Set cnDatabase = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Get the Database Name
Public Property Get Database() As String
Attribute Database.VB_Description = "Returns the parent database name."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDomain.Property Get Database()", etFullDebug

Static szDB As String

  If szDB = "" Then szDB = GetDatabase(cnDatabase.ConnectionString)
  Database = szDB
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the OID
Friend Property Let OID(ByVal vData As Long)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDomain.Property Let OID(" & vData & ")", etFullDebug

  lOID = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get OID() As Long
Attribute OID.VB_Description = "Returns the Type OID."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDomain.Property Get OID()", etFullDebug

  OID = lOID
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Name
Friend Property Let Name(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDomain.Property Let Name(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szName = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Name() As String
Attribute Name.VB_Description = "Returns the Type name."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDomain.Property Get Name()", etFullDebug

  Name = szName
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Identifier
Friend Property Let Identifier(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDomain.Property Let Identifier(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szIdentifier = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Identifier() As String
Attribute Identifier.VB_Description = "Returns the onject identifier."
Attribute Identifier.VB_MemberFlags = "200"
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDomain.Property Get Identifier()", etFullDebug

  Identifier = szIdentifier
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the Object Type
Public Property Get ObjectType() As String
Attribute ObjectType.VB_Description = "Returns the object type."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDomain.Property Get ObjectType()", etFullDebug

  ObjectType = "Domain"
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return a boolean indicating whether this is a System Object
Friend Property Let SystemObject(ByVal vData As Boolean)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDomain.Property Let SystemObject(" & vData & ")", etFullDebug

  bSystemObject = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get SystemObject() As Boolean
Attribute SystemObject.VB_Description = "Indicates whether this is a system object."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDomain.Property Get SystemObject()", etFullDebug

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

'Set/Return the Base Type
Friend Property Let BaseType(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDomain.Property Let BaseType(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szBaseType = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get BaseType() As String
Attribute BaseType.VB_Description = "Returns the Base Type used by the Domain."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDomain.Property Get BaseType()", etFullDebug

  BaseType = szBaseType
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Length
Friend Property Let Length(ByVal vData As Long)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDomain.Property Let Length(" & vData & ")", etFullDebug

  lLength = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Length() As Long
Attribute Length.VB_Description = "Returns the Length of the Domain or 0 if not applicable."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDomain.Property Get Length()", etFullDebug

  Length = lLength
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the NumericScale
Friend Property Let NumericScale(ByVal vData As Long)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDomain.Property Let NumericScale(" & vData & ")", etFullDebug

  lNumericScale = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get NumericScale() As Long
Attribute NumericScale.VB_Description = "Returns the Numeric Scale of the Domain or 0 if not applicable."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDomain.Property Get NumericScale()", etFullDebug

  NumericScale = lNumericScale
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Default
Friend Property Let Default(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDomain.Property Let Default(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szDefault = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Default() As String
Attribute Default.VB_Description = "Returns the Default value for the Domain."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDomain.Property Get Default()", etFullDebug

  Default = szDefault
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return NotNull?
Friend Property Let NotNull(ByVal vData As Boolean)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDomain.Property Let NotNull(" & vData & ")", etFullDebug

  bNotNull = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get NotNull() As Boolean
Attribute NotNull.VB_Description = "Returns a value indicating whether the Domain will disallow NULL values."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDomain.Property Get NotNull()", etFullDebug

  NotNull = bNotNull
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Comment
Public Property Let Comment(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDomain.Property Let Comment(" & vData & ")", etFullDebug

Dim szSQL As String

  szSQL = "COMMENT ON DOMAIN " & QUOTE & szName & QUOTE & " IS '" & dbSZ(vData) & "'"
  objServer.ExecSQL szSQL, cnDatabase
  objServer.Databases(Database).KillCommentCache
  If Not SystemObject Then Commit rcUpdate, "Updated object comment."
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Comment() As String
Attribute Comment.VB_Description = "Sets/Returns the Type comment."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDomain.Property Get Comment()", etFullDebug

  Comment = objServer.Databases(Database).LookupComment(lOID)
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

Public Property Get SQL() As String
Attribute SQL.VB_Description = "Returns the object's reverse engineered SQL."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDomain.Property Get SQL()", etFullDebug
Dim szSQL As String
  
  szSQL = "-- Domain: " & szName & vbCrLf & _
          "CREATE DOMAIN " & QUOTE & szName & QUOTE & " AS " & szBaseType
          
  If szBaseType = "numeric" Then szSQL = szSQL & "(" & lLength & ", " & lNumericScale & ")"
  If szBaseType = "char" Or szBaseType = "varchar" Then szSQL = szSQL & "(" & lLength & ")"
  If szDefault <> "" Then szSQL = szSQL & " DEFAULT " & szDefault
  If bNotNull Then szSQL = szSQL & " NOT NULL"
  
  szSQL = szSQL & vbCrLf
          
  If Comment <> "" Then szSQL = szSQL & "COMMENT ON DOMAIN " & QUOTE & szName & QUOTE & " IS '" & dbSZ(Comment) & "';" & vbCrLf
    
  SQL = szSQL
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Property

Public Sub Commit(UpdateType As rcLogType, UpdateComment As String)
Attribute Commit.VB_Description = "Commit object changes."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDomain.Commit(" & UpdateType & ", " & QUOTE & UpdateComment & QUOTE & ")", etFullDebug

  objServer.Databases(Database).LogRevision ObjectType, Identifier, OID, SQL, UpdateComment, UpdateType
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub

Public Property Get RCStatus() As rcStat
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDomain.RCStatus()", etFullDebug

Dim szSQL As String
Dim rs As New Recordset

  If (Not objServer.Databases(Database).RevisionControl) Or SystemObject Then
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


