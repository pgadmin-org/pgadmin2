VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pgSequence"
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
Private bSystemObject As String
Private szOwner As String
Private szACL As String
Private szIncrement As String
Private szMaximum As String
Private szMinimum As String
Private szCache As String
Private bCycled As Boolean
Private cnDatabase As Connection
Private dTimeStamp As Date

'History
Private colHistory As History

'Return the History collection.
Public Property Get History() As History
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgSequence.Property Get History()", etFullDebug

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
objServer.iLogEvent "Entering " & App.Title & ":pgSequence.Property Let TimeStamp(" & vData & ")", etFullDebug

  dTimeStamp = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get TimeStamp() As Date
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgSequence.Property Get TimeStamp()", etFullDebug

  TimeStamp = dTimeStamp
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set the Connection to use
Friend Property Set Connection(vData As Connection)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgSequence.Property Set Connection(Connection: " & vData & ")", etFullDebug

  Set cnDatabase = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Get the Database Name
Public Property Get Database() As String
Attribute Database.VB_Description = "Returns the parent database name."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgSequence.Property Get Database()", etFullDebug

  Database = GetDatabase(cnDatabase.ConnectionString)
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the OID
Friend Property Let OID(ByVal vData As Long)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgSequence.Property Let OID(" & vData & ")", etFullDebug

  lOID = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get OID() As Long
Attribute OID.VB_Description = "Returns the Sequence OID."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgSequence.Property Get OID()", etFullDebug

  OID = lOID
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Name
Friend Property Let Name(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgSequence.Property Let Name(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szName = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Name() As String
Attribute Name.VB_Description = "Returns the Sequence name."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgSequence.Property Get Name()", etFullDebug

  Name = szName
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Identifier
Friend Property Let Identifier(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgSequence.Property Let Identifier(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szIdentifier = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Identifier() As String
Attribute Identifier.VB_Description = "Returns the object identifier."
Attribute Identifier.VB_MemberFlags = "200"
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgSequence.Property Get Identifier()", etFullDebug

  Identifier = szIdentifier
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the Object Type
Public Property Get ObjectType() As String
Attribute ObjectType.VB_Description = "Returns the object type."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgSequence.Property Get ObjectType()", etFullDebug

  ObjectType = "Sequence"
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return a boolean indicating whether this is a System Object
Friend Property Let SystemObject(ByVal vData As Boolean)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgSequence.Property Let SystemObject(" & vData & ")", etFullDebug

  bSystemObject = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get SystemObject() As Boolean
Attribute SystemObject.VB_Description = "Indicates whether this is a system object."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgSequence.Property Get SystemObject()", etFullDebug

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
Attribute Owner.VB_Description = "Returns the Sequence Owner."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get Owner()", etFullDebug

  Owner = szOwner
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the ACL
Friend Property Let ACL(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgSequence.Property Let ACL(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szACL = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get ACL() As String
Attribute ACL.VB_Description = "Returns the Sequence's ACL."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgSequence.Property Get ACL()", etFullDebug

  ACL = szACL
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the LastValue (Always get the latest value)
Public Property Let LastValue(ByVal vData As Variant)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgSequence.Property Let LastValue(" & vData & ")", etFullDebug

  'Use setval to set the last value
  objServer.ExecSQL "SELECT setval('" & dbSZ(szName) & "', " & vData & ")", cnDatabase
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get LastValue() As Variant
Attribute LastValue.VB_Description = "Returns the Sequence's Last value."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgSequence.Property Get LastValue()", etFullDebug

Dim rs As New Recordset
Dim szSQL As String
  
  szSQL = "SELECT last_value FROM " & QUOTE & szName & QUOTE
  Set rs = objServer.ExecSQL(szSQL, cnDatabase)
  If Not rs.EOF Then
    LastValue = rs!last_value & ""
  Else
    LastValue = 0
  End If
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
    
  Exit Property
Err_Handler:
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  objServer.iLogError Err
  
End Property

'Set/Return the Increment
Friend Property Let Increment(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgSequence.Property Let Increment(" & vData & ")", etFullDebug

  szIncrement = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Increment() As String
Attribute Increment.VB_Description = "Returns the Sequence's Increment value."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgSequence.Property Get Increment()", etFullDebug

  Increment = szIncrement
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Maximum
Friend Property Let Maximum(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgSequence.Property Let Maximum(" & vData & ")", etFullDebug

  szMaximum = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Maximum() As String
Attribute Maximum.VB_Description = "Returns the Sequence's Maximum value."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgSequence.Property Get Maximum()", etFullDebug

  Maximum = szMaximum
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Minimum
Friend Property Let Minimum(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgSequence.Property Let Minimum(" & vData & ")", etFullDebug

  szMinimum = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Minimum() As String
Attribute Minimum.VB_Description = "Returns the Sequence's Minimum value."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgSequence.Property Get Minimum()", etFullDebug

  Minimum = szMinimum
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Cache
Friend Property Let Cache(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgSequence.Property Let Cache(" & vData & ")", etFullDebug

  szCache = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Cache() As String
Attribute Cache.VB_Description = "Returns the Sequence's Cache value."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgSequence.Property Get Cache()", etFullDebug

  Cache = szCache
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return Cycled?
Friend Property Let Cycled(ByVal vData As Boolean)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgSequence.Property Let Cycled(" & vData & ")", etFullDebug

  bCycled = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Cycled() As Boolean
Attribute Cycled.VB_Description = "Returns a value indicating whether the Sequence value is cycled."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgSequence.Property Get Cycled()", etFullDebug

  Cycled = bCycled
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set/Return the Comment
Public Property Let Comment(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgSequence.Property Let Comment(" & vData & ")", etFullDebug

Dim szSQL As String

  szSQL = "COMMENT ON SEQUENCE " & QUOTE & szName & QUOTE & " IS '" & dbSZ(vData) & "'"
  objServer.ExecSQL szSQL, cnDatabase
  objServer.Databases(GetDatabase(cnDatabase.ConnectionString)).KillCommentCache
  If Not SystemObject Then Commit rcUpdate, "Updated object comment."
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property
Public Property Get Comment() As String
Attribute Comment.VB_Description = "Sets/Returns the Sequence's comment."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgSequence.Property Get Comment()", etFullDebug

  Comment = objServer.Databases(GetDatabase(cnDatabase.ConnectionString)).LookupComment(lOID)
    
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Grant Permissions
Public Sub Grant(ByVal UserOrGroup As String, ByVal Privilege As aclPrivileges)
Attribute Grant.VB_Description = "Grant privileges on the Sequence"
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgSequence.Grant(" & QUOTE & UserOrGroup & QUOTE & ", " & Privilege & ")", etFullDebug

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
  
  szSQL = szSQL & " ON " & QUOTE & szName & QUOTE & " TO " & UserOrGroup
  objServer.ExecSQL szSQL, cnDatabase
  
  'Get the current ACL
  szSQL = "SELECT relacl FROM pg_class WHERE oid = " & lOID
  Set rs = objServer.ExecSQL(szSQL, cnDatabase)
  If Not rs.EOF Then
    szACL = rs!relacl & ""
  End If
  
  If Not SystemObject Then Commit rcUpdate, "Updated object ACL."
  
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
Attribute Revoke.VB_Description = "Revoke privileges on the Sequence."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgSequence.Revoke(" & QUOTE & UserOrGroup & QUOTE & ", " & Privilege & ")", etFullDebug

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
  
  szSQL = szSQL & " ON " & QUOTE & szName & QUOTE & " FROM " & UserOrGroup
  objServer.ExecSQL szSQL, cnDatabase
  
  'Get the current ACL
  szSQL = "SELECT relacl FROM pg_class WHERE oid = " & lOID
  Set rs = objServer.ExecSQL(szSQL, cnDatabase)
  If Not rs.EOF Then
    szACL = rs!relacl & ""
  End If
  
  If Not SystemObject Then Commit rcUpdate, "Updated object ACL."
  
  'Cleanup
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  
  Exit Sub
Err_Handler:
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  objServer.iLogError Err
End Sub

Public Property Get SQL(Optional Reset As Boolean) As String
Attribute SQL.VB_Description = "Returns the object's reverse engineered SQL."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgSequence.Property Get SQL()", etFullDebug
Dim szSQL As String
  
  szSQL = "-- Sequence: " & szName & vbCrLf & _
          "CREATE SEQUENCE " & QUOTE & szName & QUOTE & " " & _
          "INCREMENT " & szIncrement & " " & _
          "MINVALUE " & szMinimum & " " & _
          "MAXVALUE " & szMaximum & " " & _
          "CACHE " & szCache
          
  If bCycled Then szSQL = szSQL & " CYCLE"
  szSQL = szSQL & ";" & vbCrLf
  If Not Reset Then szSQL = szSQL & "SELECT setval('" & dbSZ(szName) & "', " & LastValue & ");" & vbCrLf
  
  If szACL <> "" Then szSQL = szSQL & ParseACL(szName, szACL)
  If Comment <> "" Then szSQL = szSQL & "COMMENT ON SEQUENCE " & QUOTE & szName & QUOTE & " IS '" & dbSZ(Comment) & "';" & vbCrLf
    
  SQL = szSQL
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

Public Sub Commit(UpdateType As rcLogType, UpdateComment As String)
Attribute Commit.VB_Description = "Commit object changes."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgSequence.Commit(" & UpdateType & ", " & QUOTE & UpdateComment & QUOTE & ")", etFullDebug

  objServer.Databases(Database).LogRevision ObjectType, Identifier, OID, SQL(True), UpdateComment, UpdateType
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub

Public Property Get RCStatus() As rcStat
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgSequence.RCStatus()", etFullDebug

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
    If SQL(True) = rs!rc_definition & "" Then
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


