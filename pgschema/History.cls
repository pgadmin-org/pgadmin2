VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "History"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' pgSchema - PostgreSQL Schema Objects
' Copyright (C) 2001, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence

Option Explicit

'Private Variables
Private mCol As Collection
Private cnDatabase As Connection

Dim szType As String
Dim szIdentifier As String

'Set the Connection to use
Friend Property Set Connection(vData As Connection)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":History.Property Set Connection(Connection: " & vData & ")", etFullDebug

  Set cnDatabase = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set the Type for internal use.
Friend Property Let objType(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":History.Property Let ObjType(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szType = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set the Identifier for internal use.
Friend Property Let ObjIdentifier(ByVal vData As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":History.Property Let ObjIdentifier(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szIdentifier = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Internal Add
Private Sub iAdd(szType As String, szIdentifier As String, szTable As String, dTimeStamp As Date, szUser As String, szAction As String, lOID As Long, lVersion As Long, szDefinition As String, szComment As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":History.iAdd(" & QUOTE & szType & QUOTE & ", " & QUOTE & szIdentifier & QUOTE & ", " & QUOTE & szTable & QUOTE & ", " & QUOTE & dTimeStamp & QUOTE & ", " & QUOTE & szUser & QUOTE & ", " & QUOTE & szAction & QUOTE & ", " & lOID & ", " & lVersion & ", " & QUOTE & szDefinition & QUOTE & ", " & QUOTE & szComment & QUOTE & ")", etFullDebug

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
Attribute Item.VB_Description = "Returns a history item."
Attribute Item.VB_UserMemId = 0
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":History.Property Get Item(" & Index & ")", etFullDebug

  Set Item = mCol(Index)
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the Item count
Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of history items in the collection."
Attribute Count.VB_MemberFlags = "200"
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":History.Property Get Count()", etFullDebug

    Count = mCol.Count

  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Allow enumeration using For...Each
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":History.Property Get NewEnum()", etFullDebug

  Set NewEnum = mCol.[_NewEnum]
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Refresh the collection
Public Sub Refresh()
Attribute Refresh.VB_Description = "Refresh the collection."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":History.Refresh()", etFullDebug

  'Initialize will basically do a refresh so let's use that!
  Initialize
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub

'Initialise the collection if possible
Friend Sub Initialize()
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Checks.Initialize()", etFullDebug

Dim rs As New Recordset
Dim szSQL As String

  'Create a new collection
  Set mCol = New Collection
  
  'Build the SQL query.
  szSQL = "SELECT * FROM pgadmin_rclog WHERE rc_type = '" & dbSZ(szType) & "' AND rc_identifier = '" & dbSZ(szIdentifier) & "' ORDER BY rc_version DESC"
  
  'Get the data and loop through it
  If objServer.Connection.State = adStateOpen Then
    Set rs = objServer.ExecSQL(szSQL, cnDatabase)
    While Not rs.EOF
      iAdd rs!rc_type & "", rs!rc_identifier & "", rs!rc_table & "", rs!rc_timestamp, rs!rc_user & "", rs!rc_action & "", rs!rc_oid, rs!rc_version, rs!rc_definition & "", rs!rc_comment & ""
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

