Attribute VB_Name = "basDepend"
' pgSchema - PostgreSQL Schema Objects
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence

' basDepend.bas - pg_depend function

Option Explicit

Public Enum EDepRef
  EDR_Reference
  EDR_Depend
End Enum
 
'Return the dependent/referenced object is in
Public Function DepRef(Oid As Double, cnDatabase As Connection, Database As String, TypeDR As EDepRef) As Collection
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":basDepend.DepRef(" & Oid & "," & QUOTE & cnDatabase.ConnectionString & QUOTE & "," & QUOTE & Database & QUOTE & "," & TypeDR & ")", etFullDebug

Dim szSQL As String
Dim rs As Recordset
Dim rsDep As Recordset
Dim colDep As New Collection
Dim objTmp
Dim dOID As Double
Dim szTableName As String

  If objVersion.VersionNum < 7.3 Then
    Set DepRef = colDep
    Exit Function
  End If

  If TypeDR = EDR_Depend Then
    szSQL = "SELECT refclassid, refobjid , refobjsubid, deptype FROM pg_depend WHERE refclassid>0 AND objid=" & Oid & " ORDER BY refclassid"
  ElseIf TypeDR = EDR_Reference Then
    szSQL = "SELECT classid, objid, objsubid, deptype FROM pg_depend WHERE classid > 0 AND refobjid=" & Oid & " ORDER BY classid"
  End If
  Set rsDep = objServer.ExecSQL(szSQL, cnDatabase)
  While Not rsDep.EOF
    dOID = rsDep.Fields(1).Value
  
    'get name table by oid
    szSQL = "SELECT relname FROM pg_class WHERE oid=" & rsDep.Fields(0).Value
    Set rs = objServer.ExecSQL(szSQL, cnDatabase)
    szTableName = rs!relname
       
    Select Case szTableName
'pg_opclass
      
      Case "pg_attrdef"   'default value
        szSQL = "SELECT c.relname,c.oid,(select attname FROM pg_attribute WHERE attnum=a.adnum AND attrelid=c.oid) as attname FROM pg_class c, pg_attrdef a WHERE a.adrelid=c.oid AND a.oid=" & dOID
        Set rs = objServer.ExecSQL(szSQL, cnDatabase)
        Set objTmp = GetObjectTypePgClass(rs!Oid, cnDatabase, Database)
        colDep.Add objTmp(rs!relname).Columns(rs!attname)
      
      Case "pg_trigger"   'trigger
        szSQL = "SELECT c.relname,c.oid FROM pg_class c, pg_trigger t WHERE t.tgrelid=c.oid AND t.oid=" & dOID
        Set rs = objServer.ExecSQL(szSQL, cnDatabase)
        Set objTmp = GetObjectTypePgClass(rs!Oid, cnDatabase, Database)
        AddObjDepend dOID, objTmp(rs!relname).Triggers, colDep
      
      Case "pg_rewrite"   'rule
        szSQL = "SELECT c.relname,c.oid FROM pg_class c, pg_rewrite r WHERE r.ev_class=c.oid AND r.oid=" & dOID
        Set rs = objServer.ExecSQL(szSQL, cnDatabase)
        Set objTmp = GetObjectTypePgClass(rs!Oid, cnDatabase, Database)
        AddObjDepend dOID, objTmp(rs!relname).Rules, colDep
        
      Case "pg_language"  'language
        AddObjDepend dOID, objServer.Databases(Database).Languages, colDep
      
      Case "pg_cast"      'cast
        AddObjDepend dOID, objServer.Databases(Database).Casts, colDep
      
      Case "pg_namespace" 'namespace
        AddObjDepend dOID, objServer.Databases(Database).Namespaces, colDep
      
      Case "pg_proc"      'function
        szSQL = "SELECT n.nspname FROM pg_namespace n ,pg_proc p WHERE p.pronamespace=n.oid AND p.oid=" & dOID
        Set rs = objServer.ExecSQL(szSQL, cnDatabase)
        AddObjDepend dOID, objServer.Databases(Database).Namespaces(rs!nspname).Functions, colDep
      
      Case "pg_type"
        szSQL = "SELECT t.typtype, n.nspname FROM pg_namespace n,pg_type t WHERE t.typnamespace=n.oid AND t.oid=" & dOID
        Set rs = objServer.ExecSQL(szSQL, cnDatabase)
        Select Case rs!typtype
          
          Case "b", "p"   'base type, pseudo-type
            AddObjDepend dOID, objServer.Databases(Database).Namespaces(rs!nspname).Types, colDep
          
          Case "c"        'complex type
          
          Case "d"        'domain
            AddObjDepend dOID, objServer.Databases(Database).Namespaces(rs!nspname).Domains, colDep
      
          Case Else
            Err.Raise -1, "", "Not found type pg_type " & rs!typtype
        
        End Select
      
      Case "pg_class"
        AddObjDepend dOID, GetObjectTypePgClass(dOID, cnDatabase, Database), colDep
      
    End Select
    
    rsDep.MoveNext
  Wend
  Set DepRef = colDep

  Exit Function
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Function

'Find and Add in collection the object by Oid
Private Function AddObjDepend(dOID As Double, ObjFind, colDep As Collection) As Boolean
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":basDepend.AddObjDepend(" & dOID & "," & colDep.Count & ")", etFullDebug

Dim objTmp
  
  If ObjFind Is Nothing Then Exit Function
  
  For Each objTmp In ObjFind
    If objTmp.Oid = dOID Then
      AddObjDepend = True
      colDep.Add objTmp
      Exit Function
    End If
  Next
  AddObjDepend = False

  Exit Function
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Function

'get object type on relation into pg_class
Private Function GetObjectTypePgClass(dOID As Double, cnDatabase As Connection, Database As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":basDepend.GetObjectTypePgClass(" & dOID & "," & QUOTE & cnDatabase.ConnectionString & QUOTE & "," & QUOTE & Database & ")", etFullDebug

Dim szSQL As String
Dim rs As Recordset
Dim objTmp

  szSQL = "SELECT c.relkind, n.nspname FROM pg_namespace n,pg_class c WHERE c.relnamespace=n.oid AND c.oid=" & dOID
  Set rs = objServer.ExecSQL(szSQL, cnDatabase)
  Select Case rs!relkind
    Case "r", "t"   'table, TOAST
      Set GetObjectTypePgClass = objServer.Databases(Database).Namespaces(rs!nspname).Tables
         
    Case "i"        'index
      szSQL = "SELECT c.relname,c.oid FROM pg_index i, pg_class c WHERE i.indrelid=c.oid AND indexrelid=" & dOID
      Set rs = objServer.ExecSQL(szSQL, cnDatabase)
      Set objTmp = GetObjectTypePgClass(rs!Oid, cnDatabase, Database)
      Set GetObjectTypePgClass = objTmp(rs!relname).Indexes
          
    Case "S"        'sequence
      Set GetObjectTypePgClass = objServer.Databases(Database).Namespaces(rs!nspname).Sequences
            
    Case "v"        'view
      Set GetObjectTypePgClass = objServer.Databases(Database).Namespaces(rs!nspname).Views
      
    Case Else
      Set GetObjectTypePgClass = Nothing
      
    End Select
  
  Exit Function
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Function
