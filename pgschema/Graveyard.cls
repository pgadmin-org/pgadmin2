VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Graveyard"
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

'Get the Database Name
Public Property Get Database() As String
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Graveyard.Property Get Database()", etFullDebug

Static szDB As String

  If szDB = "" Then szDB = GetDatabase(cnDatabase.ConnectionString)
  Database = szDB
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set the Connection to use
Friend Property Set Connection(vData As Connection)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Graveyard.Property Set Connection(Connection: " & vData & ")", etFullDebug

  Set cnDatabase = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Internal Add
Private Sub iAdd(szType As String, szIdentifier As String, szTable As String, dTimeStamp As Date, szUser As String, szAction As String, lOID As Long, lVersion As Long, szDefinition As String, szComment As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Graveyard.iAdd(" & QUOTE & szType & QUOTE & ", " & QUOTE & szIdentifier & QUOTE & ", " & QUOTE & szTable & QUOTE & ", " & QUOTE & dTimeStamp & QUOTE & ", " & QUOTE & szUser & QUOTE & ", " & QUOTE & szAction & QUOTE & ", " & lOID & ", " & lVersion & ", " & QUOTE & szDefinition & QUOTE & ", " & QUOTE & szComment & QUOTE & ")", etFullDebug

Dim objNewMember As rcItem

  'Create a new object
  Set objNewMember = New rcItem
  objNewMember.ObjectType = szType
  objNewMember.Identifier = szIdentifier
  objNewMember.ParentTable = szTable
  objNewMember.TimeStamp = dTimeStamp
  objNewMember.User = szUser
  objNewMember.Action = szAction
  objNewMember.OID = lOID
  objNewMember.Version = lVersion
  objNewMember.Definition = szDefinition
  objNewMember.Comment = szComment

  mCol.Add objNewMember
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub

'Return the specified Item
Public Property Get Item(Index As Long) As rcItem
Attribute Item.VB_Description = "Returns a Graveyard item."
Attribute Item.VB_UserMemId = 0
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Graveyard.Property Get Item(" & Index & ")", etFullDebug

  Set Item = mCol(Index)
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the Item count
Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of Graveyard items in the collection."
Attribute Count.VB_MemberFlags = "200"
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Graveyard.Property Get Count()", etFullDebug

    Count = mCol.Count

  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Allow enumeration using For...Each
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Graveyard.Property Get NewEnum()", etFullDebug

  Set NewEnum = mCol.[_NewEnum]
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Refresh the collection
Public Sub Refresh()
Attribute Refresh.VB_Description = "Refresh the collection."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Graveyard.Refresh()", etFullDebug

  'Initialize will basically do a refresh so let's use that!
  Initialize
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub

'Restore an object
Public Sub Restore(Index As Long)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Graveyard.Restore()", etFullDebug

  'Restore the object
  objServer.Databases(Database).Execute mCol(Index).Definition
  
  'Now refresh the relevant bit of the object model, and log the restore.
  Select Case mCol(Index).ObjectType
    Case "Aggregate"
      objServer.Databases(Database).Aggregates.Refresh
      objServer.Databases(Database).Aggregates(mCol(Index).Identifier).Commit rcAdd, "Object restored from Revision Log (Version " & mCol(Index).Version & ")."
    Case "Function"
      objServer.Databases(Database).Functions.Refresh
      objServer.Databases(Database).Functions(mCol(Index).Identifier).Commit rcAdd, "Object restored from Revision Log (Version " & mCol(Index).Version & ")."
    Case "Index"
      objServer.Databases(Database).Tables(mCol(Index).ParentTable).Indexes.Refresh
      objServer.Databases(Database).Tables(mCol(Index).ParentTable).Indexes(mCol(Index).Identifier).Commit rcAdd, "Object restored from Revision Log (Version " & mCol(Index).Version & ")."
    Case "Language"
      objServer.Databases(Database).Languages.Refresh
      objServer.Databases(Database).Languages(mCol(Index).Identifier).Commit rcAdd, "Object restored from Revision Log (Version " & mCol(Index).Version & ")."
    Case "Operator"
      objServer.Databases(Database).Operators.Refresh
      objServer.Databases(Database).Operators(mCol(Index).Identifier).Commit rcAdd, "Object restored from Revision Log (Version " & mCol(Index).Version & ")."
    Case "Rule"
      objServer.Databases(Database).Tables(mCol(Index).ParentTable).Rules.Refresh
      objServer.Databases(Database).Tables(mCol(Index).ParentTable).Rules(mCol(Index).Identifier).Commit rcAdd, "Object restored from Revision Log (Version " & mCol(Index).Version & ")."
    Case "Sequence"
      objServer.Databases(Database).Sequences.Refresh
      objServer.Databases(Database).Sequences(mCol(Index).Identifier).Commit rcAdd, "Object restored from Revision Log (Version " & mCol(Index).Version & ")."
    Case "Table"
      objServer.Databases(Database).Tables.Refresh
      objServer.Databases(Database).Tables(mCol(Index).Identifier).Commit rcAdd, "Object restored from Revision Log (Version " & mCol(Index).Version & ")."
    Case "Trigger"
      objServer.Databases(Database).Tables(mCol(Index).ParentTable).Triggers.Refresh
      objServer.Databases(Database).Tables(mCol(Index).ParentTable).Triggers(mCol(Index).Identifier).Commit rcAdd, "Object restored from Revision Log (Version " & mCol(Index).Version & ")."
    Case "Type"
      objServer.Databases(Database).Types.Refresh
      objServer.Databases(Database).Types(mCol(Index).Identifier).Commit rcAdd, "Object restored from Revision Log (Version " & mCol(Index).Version & ")."
    Case "View"
      objServer.Databases(Database).Views.Refresh
      objServer.Databases(Database).Views(mCol(Index).Identifier).Commit rcAdd, "Object restored from Revision Log (Version " & mCol(Index).Version & ")."
  End Select
  
  'Now remove the Graveyard object
  mCol.Remove Index
  
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub

'Initialise the collection if possible
Friend Sub Initialize()
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Checks.Initialize()", etFullDebug

Dim rs As New Recordset
Dim bExists As Boolean
Dim objTable As pgTable

  'Create a new collection
  Set mCol = New Collection
  
  'Get the data and loop through it
  If objServer.Connection.State = adStateOpen Then
    Set rs = objServer.ExecSQL(SQL_GRAVEYARD, cnDatabase)
    
    'Refresh the database so everythig's up-to-date
    objServer.Databases(Database).Aggregates.Refresh
    objServer.Databases(Database).Functions.Refresh
    objServer.Databases(Database).Languages.Refresh
    objServer.Databases(Database).Operators.Refresh
    objServer.Databases(Database).Sequences.Refresh
    objServer.Databases(Database).Tables.Refresh
    objServer.Databases(Database).Types.Refresh
    objServer.Databases(Database).Views.Refresh
    
    While Not rs.EOF
      If rs!rc_action = "D" Then 'Just add an delete items
        iAdd rs!rc_type & "", rs!rc_identifier & "", rs!rc_table & "", rs!rc_timestamp, rs!rc_user & "", rs!rc_action & "", rs!rc_oid, rs!rc_version, rs!rc_definition & "", rs!rc_comment & ""
      Else 'Check that the item doesn't exist and more.
        bExists = False
        Select Case rs!rc_type & ""
          Case "Aggregate"
            bExists = objServer.Databases(Database).Aggregates.Exists(rs!rc_identifier & "")
          Case "Function"
            bExists = objServer.Databases(Database).Functions.Exists(rs!rc_identifier & "")
          Case "Language"
            bExists = objServer.Databases(Database).Languages.Exists(rs!rc_identifier & "")
          Case "Operator"
            bExists = objServer.Databases(Database).Operators.Exists(rs!rc_identifier & "")
          Case "Sequence"
            bExists = objServer.Databases(Database).Sequences.Exists(rs!rc_identifier & "")
          Case "Table"
            bExists = objServer.Databases(Database).Tables.Exists(rs!rc_identifier & "")
          Case "Type"
            bExists = objServer.Databases(Database).Types.Exists(rs!rc_identifier & "")
          Case "View"
            bExists = objServer.Databases(Database).Views.Exists(rs!rc_identifier & "")
          Case "Index"
            bExists = objServer.Databases(Database).Tables(rs!rc_table & "").Indexes.Exists(rs!rc_identifier & "")
          Case "Rule"
            bExists = objServer.Databases(Database).Tables(rs!rc_table & "").Rules.Exists(rs!rc_identifier & "")
          Case "Trigger"
            bExists = objServer.Databases(Database).Tables(rs!rc_table & "").Triggers.Exists(rs!rc_identifier & "")
        End Select
        
        'Add the item if we couldn't find it
        If Not bExists Then iAdd rs!rc_type & "", rs!rc_identifier & "", rs!rc_table & "", rs!rc_timestamp, rs!rc_user & "", rs!rc_action & "", rs!rc_oid, rs!rc_version, rs!rc_definition & "", rs!rc_comment & ""
      End If
      rs.MoveNext
    Wend
  End If
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  
  Exit Sub
Err_Handler:
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  objServer.iLogError Err
End Sub

'Cleanup
Private Sub Class_Terminate()
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Checks.Class_Terminate()", etFullDebug

  Set mCol = Nothing
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub

