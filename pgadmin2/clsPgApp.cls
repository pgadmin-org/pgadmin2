VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPgApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence
'
' clsPgApp.cls - class export application function/form

Public frmAbout As pgAdmin2.frmAbout                        'frmAbout.frm
Public frmAggregate As pgAdmin2.frmAggregate                'frmAggregate.frm
Public frmCast As pgAdmin2.frmCast                          'frmCast.frm
Public frmColumn As pgAdmin2.frmColumn                      'frmColumn.frm
Public frmConnect As pgAdmin2.frmConnect                    'frmConnect.frm
Public frmDatabase As pgAdmin2.frmDatabase                  'frmDatabase.frm
Public frmDomain As pgAdmin2.frmDomain                      'frmDomain.frm
Public frmError As pgAdmin2.frmError                        'frmError.frm
Public frmFind As pgAdmin2.frmFind                          'frmFind.frm
Public frmForeignKey As pgAdmin2.frmForeignKey              'frmForeignKey.frm
Public frmFunction As pgAdmin2.frmFunction                  'frmFunction.frm
Public frmGroup As pgAdmin2.frmGroup                        'frmGroup.frm
Public frmIndex As pgAdmin2.frmIndex                        'frmIndex.frm
Public frmLanguage As pgAdmin2.frmLanguage                  'frmLanguage.frm
Public frmLog As pgAdmin2.frmLog                            'frmLog.frm
Public frmMain As pgAdmin2.frmMain                          'frmMain.frm
Public frmNamespace As pgAdmin2.frmNamespace                'frmNamespace.frm
Public frmOperator As pgAdmin2.frmOperator                  'frmOperator.frm
Public frmOptions As pgAdmin2.frmOptions                    'frmOptions.frm
Public frmPassword As pgAdmin2.frmPassword                  'frmPassword.frm
Public frmRecordLog As pgAdmin2.frmRecordLog                'frmRecordLog.frm
Public frmRule As pgAdmin2.frmRule                          'frmRule.frm
Public frmSequence As pgAdmin2.frmSequence                  'frmSequence.frm
Public frmServer As pgAdmin2.frmServer                      'frmServer.frm
Public frmSplash As pgAdmin2.frmSplash                      'frmSplash.frm
Public frmSQLExplain As pgAdmin2.frmSQLExplain              'frmSQLExplain.frm
Public frmSQLInput As pgAdmin2.frmSQLInput                  'frmSQLInput.frm
Public frmSQLOutput As pgAdmin2.frmSQLOutput                'frmSQLOutput.frm
Public frmSQLWizard As pgAdmin2.frmSQLWizard                'frmSQLWizard.frm
Public frmTable As pgAdmin2.frmTable                        'frmTable.frm
Public frmTip As pgAdmin2.frmTip                            'frmTip.frm
Public frmTrigger As pgAdmin2.frmTrigger                    'frmTrigger.frm
Public frmType As pgAdmin2.frmType                          'frmType.frm
Public frmUpgradeWizard As pgAdmin2.frmUpgradeWizard        'frmUpgradeWizard.frm
Public frmUser As pgAdmin2.frmUser                          'frmUser.frm
Public frmView As pgAdmin2.frmView                          'frmView.frm

''''''''''''''''''''''''''''''''''''''''''''''''''
'Module: basActions
'File: basActions.bas
Public Sub Vacuum(bAnalyse As Boolean)
  Call basActions.Vacuum(bAnalyse)
End Sub

Public Sub Drop()
  Call basActions.Drop()

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''
'Module: basAPI
'File: basAPI.bas
''''''''''''''''''''''''''''''''''''''''''''''''''
'Module: basGlobal
'File: basGlobal.bas
''''''''''''''''''''''''''''''''''''''''''''''''''
'Module: basMisc
'File: basMisc.bas
Public Sub Main()
  Call basMisc.Main()

End Sub

Public Function GetID() As String
  GetID = basMisc.GetID()
End Function

Public Function SetTopMostWindow(hwnd As Long, Topmost As Boolean) As Long
  SetTopMostWindow = basMisc.SetTopMostWindow(hwnd,Topmost)
End Function

Public Sub BuildConnectionMenu()
  Call basMisc.BuildConnectionMenu()

End Sub

Public Sub BuildPluginsMenu()
  Call basMisc.BuildPluginsMenu()

End Sub

Public Sub LogError(lError As Long, szError As String, szRoutine As String)
  Call basMisc.LogError(lError,szError,szRoutine)
End Sub

Public Sub StartMsg(ByVal szMsg As String)
  Call basMisc.StartMsg(szMsg)
End Sub

Public Sub EndMsg(Optional ByVal szErr As String)
  Call basMisc.EndMsg(szErr)
End Sub

Public Function dbSZ(szData As String) As String
  dbSZ = basMisc.dbSZ(szData)
End Function

'This code is based on fmtID from the pg_dump code
Public Function fmtID(ByVal szData As String) As String
  fmtID = basMisc.fmtID(szData)
End Function

Public Function Bool2Bin(bData As Boolean) As Integer
  Bool2Bin = basMisc.Bool2Bin(bData)
End Function

Public Function Bin2Bool(iData As Integer) As Boolean
  Bin2Bool = basMisc.Bin2Bool(iData)
End Function

'Parse an ACL and return | delimited User/Access lists
Public Sub ParseACL(ByVal szACL As String, ByRef szUserlist As String, ByRef szAccesslist As String)
  Call basMisc.ParseACL(szACL,szUserlist,szAccesslist)
End Sub

'Format a typename
Public Function fmtTypeName(objType As pgType) As String
  fmtTypeName = basMisc.fmtTypeName(objType)
End Function

Public Sub AutoSizeColumnLv(lv As ListView)
  Call basMisc.AutoSizeColumnLv(lv)
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''
'Module: basPatch
'File: basPatch.bas
'Patch form
Public Sub PatchForm(objForm As Form)
  Call basPatch.PatchForm(objForm)
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''
'Module: basRegistry
'File: basRegistry.bas
Public Function RegRead(ByVal Hive As RegHives, ByVal Section As String, ByVal Key As String, Optional Default As Variant) As String
  RegRead = basRegistry.RegRead(Hive,Section,Key,Default)
End Function

Public Sub RegWrite(ByVal Hive As RegHives, ByVal Section As String, ByVal Key As String, ByVal ValType As RegTypes, ByVal Value As Variant)
  Call basRegistry.RegWrite(Hive,Section,Key,ValType,Value)
End Sub

Public Function RegGetSubkey(ByVal Hive As RegHives, ByVal Section As String, Idx As Long) As String
  RegGetSubkey = basRegistry.RegGetSubkey(Hive,Section,Idx)
End Function

Public Function RegReadAll(ByVal Hive As RegHives, ByVal Section As String, Idx As Long) As Variant
  RegReadAll = basRegistry.RegReadAll(Hive,Section,Idx)
End Function

Public Sub RegDelSubkey(ByVal Hive As RegHives, ByVal Section As String)
  Call basRegistry.RegDelSubkey(Hive,Section)
End Sub

Public Sub RegDelValue(ByVal Hive As RegHives, ByVal Section As String, ByVal Key As String)
  Call basRegistry.RegDelValue(Hive,Section,Key)
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''
'Module: basVarDb
'File: basVarDb.bas
'Initalization
Public Sub InitVarDb()
  Call basVarDb.InitVarDb()

End Sub

'return the definition var db from Collection
Public Function GetVarDb(szName As String) 'As VarDb
'  GetVarDb = basVarDb.GetVarDb(szName)
End Function

'return the image name of value
Public Function GetImageFromVal(szValue As String, TypeVar As TypeVarDb) As String
  GetImageFromVal = basVarDb.GetImageFromVal(szValue,TypeVar)
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''
'Internal routine
'create collection forms activate
Public Function FormsActivate() As Collection
Dim objCol As New Collection
Dim objTmp

  For Each objTmp In VB.Forms
    objCol.Add objTmp
  Next
  Set FormsActivate = objCol
End Function

'get forms by name
Public Function FormByName(vData As String) As Form
Dim objTmp As Form
Dim objFrm As Form

  For Each objTmp In VB.Forms
    If LCase(objTmp.Name) = LCase(vData) Then
      Set objFrm = objTmp
      Exit For
    End If
  Next
  Set GetFormByName = objFrm
End Function

