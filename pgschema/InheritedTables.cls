VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InheritedTables"
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

'Get the Database Name
Public Property Get Database() As String
Attribute Database.VB_Description = "Returns the parent database name."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":InheritedTables.Property Get Database()", etFullDebug

  Database = GetDatabase(cnDatabase.ConnectionString)
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Set the Connection to use
Friend Property Set Connection(vData As Connection)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":InheritedTables.Property Set Connection(Connection: " & vData & ")", etFullDebug

  Set cnDatabase = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Internal Add
Private Sub iAdd(szTable As String)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":InheritedTables.iAdd(" & QUOTE & szTable & QUOTE & ")", etFullDebug

  mCol.Add szTable
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub

'Return the specified Item
Public Property Get Item(Index As Long) As String
Attribute Item.VB_Description = "Returns an InheritedTable Item."
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":InheritedTables.Property Get Item(" & Index & ")", etFullDebug

  Item = mCol(Index)
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Return the Item count
Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns a count of the Inherited Tables in the current table."
Attribute Count.VB_MemberFlags = "200"
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":InheritedTables.Property Get Count()", etFullDebug

  Count = mCol.Count
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Allow enumeration using For...Each
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":InheritedTables.Property Get NewEnum()", etFullDebug

  Set NewEnum = mCol.[_NewEnum]
  
  Exit Property
Err_Handler:  objServer.iLogError Err
End Property

'Initialise the collection if possible
Friend Sub Initialize(lTableOID As Long)
On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":InheritedTables.Initialize(" & QUOTE & lTableOID & QUOTE & ")", etFullDebug

Dim rs As New Recordset

  'Create a new collection
  Set mCol = New Collection
  
  Set rs = objServer.ExecSQL(SQL_GET_INHERITED_TABLES & " AND i.inhrelid = " & lTableOID & " ORDER BY inhseqno", cnDatabase)
  While Not rs.EOF
    iAdd rs!relname & ""
    rs.MoveNext
  Wend
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
objServer.iLogEvent "Entering " & App.Title & ":InheritedTables.Class_Terminate()", etFullDebug

  Set mCol = Nothing
  
  Exit Sub
Err_Handler:  objServer.iLogError Err
End Sub
