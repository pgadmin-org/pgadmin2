' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence
'
' crtPgAppClass.vbs - Create class clsPgApp.cls

Option Explicit

Sub Main
dim fso,f,fc,szWorkDir,f1,f2,vData,ii,szTemp,szDataForm,szDataModule,szModuleName,bInRoutine,szTypeRoutine,szRoutine,szHeaderClass
dim szTypeReturn,vOutFunctionValid,szDataAdd,szConst
Const ForReading=1
Const KeyForm="Begin VB.Form"
Const KeyModule="Attribute VB_Name = "
Const KeyPublicSub="Public Sub "
Const KeyPublicFunction="Public Function "
Const Tab="  "
Const KeyGlobalConst="Global Const "
Const KeyPublicConst="Public Const "
Const KeyPublicEnum="Public Enum "
Const KeyEndEnum="End Enum"

Const Quote=""""

szHeaderClass="VERSION 1.0 CLASS" & vbcrlf & _
				"BEGIN" & vbcrlf & _
  				"  MultiUse = -1  'True"& vbcrlf & _
  				"  Persistable = 0  'NotPersistable"& vbcrlf & _
  				"  DataBindingBehavior = 0  'vbNone"& vbcrlf & _
  				"  DataSourceBehavior  = 0  'vbNone"& vbcrlf & _
  				"  MTSTransactionMode  = 0  'NotAnMTSObject"& vbcrlf & _
				"END"& vbcrlf & _
				"Attribute VB_Name = ""clsPgApp"""& vbcrlf & _
				"Attribute VB_GlobalNameSpace = False"& vbcrlf & _
				"Attribute VB_Creatable = True"& vbcrlf & _
				"Attribute VB_PredeclaredId = False"& vbcrlf & _
				"Attribute VB_Exposed = False"& vbcrlf & _
				"' pgAdmin II - PostgreSQL Tools"& vbcrlf & _
				"' Copyright (C) 2001 - 2003, The pgAdmin Development Team"& vbcrlf & _
				"' This software is released under the pgAdmin Public Licence"& vbcrlf & _
				"'"& vbcrlf & _
				"' clsPgApp.cls - class export application function/form"& vbcrlf 


szDataAdd=string(50,"'") & vbcrlf & "'Internal routine" & vbcrlf

szDataAdd=szDataAdd & "'create collection forms activate" & vbcrlf & _
			"Public Function FormsActivate() As Collection" & vbcrlf & _ 
			"Dim objCol As New Collection" & vbcrlf & _
			"Dim objTmp" & vbcrlf & _
			"" & vbcrlf & _
			"  For Each objTmp In VB.Forms" & vbcrlf & _
			"    objCol.Add objTmp" & vbcrlf & _
			"  Next" & vbcrlf & _
			"  Set FormsActivate = objCol" & vbcrlf & _
			"End Function" & vbcrlf & _
			"" & vbcrlf 

szDataAdd=szDataAdd & "'get forms by name" & vbcrlf & _
			"Public Function FormByName(vData As String) As Form" & vbcrlf & _
			"Dim objTmp As Form" & vbcrlf & _
			"Dim objFrm As Form" & vbcrlf & _
			"" & vbcrlf & _
			"  For Each objTmp In VB.Forms" & vbcrlf & _
			"    If LCase(objTmp.Name) = LCase(vData) Then" & vbcrlf & _
			"      Set objFrm = objTmp" & vbcrlf & _
			"      Exit For" & vbcrlf & _
			"    End If" & vbcrlf & _
			"  Next" & vbcrlf & _
			"  Set GetFormByName = objFrm" & vbcrlf & _
			"End Function" & vbcrlf & vbcrlf


vOutFunctionValid=array("Byte","Boolean","Integer","Long","Single","Double","Currency", "Decimal","Date","Object","String","Variant")

	Set fso = CreateObject("Scripting.FileSystemObject")
	szWorkDir=left(WScript.ScriptFullName,len(WScript.ScriptFullName)-len(WScript.ScriptName))


	'read file in directory
	Set f = fso.GetFolder(szWorkDir)
   	Set fc = f.Files
	For Each f1 in fc
		'verify extension
		select case lcase(fso.GetExtensionName(f1.name))
			case "cls"

			case "frm"

				'read form name
			   	Set f2 = fso.OpenTextFile(f1.path, ForReading)
			   	vData = split(f2.ReadAll,vbcrlf)
				for ii=0 to ubound(vData)
					if left(vData(ii),len(KeyForm))=KeyForm then
						szTemp=trim(mid(vData(ii),len(KeyForm)+1))
						szTemp="Public " & szTemp & " As pgAdmin2." & szTemp 
						'add file name
						szTemp=sztemp & space(60-len(sztemp)) & "'" & f1.name
						szDataForm=szDataForm & szTemp & vbcrlf
						exit for
					end if
				next

			case "bas"
				'read module name
			   	Set f2 = fso.OpenTextFile(f1.path, ForReading)
			   	vData = split(f2.ReadAll,vbcrlf)
				bInRoutine=False
				szModuleName=""
				szTypeRoutine=""
				szRoutine=""
				for ii=0 to ubound(vData)
					if left(vData(ii),len(KeyModule))=KeyModule then
						szModuleName=trim(mid(vData(ii),len(KeyModule)+2,len(trim(vData(ii)))-len(KeyModule)-2))
						szDataModule=szDataModule & string(50,"'") & vbcrlf
						szDataModule=szDataModule & "'Module: " & szModuleName & vbcrlf 
						szDataModule=szDataModule & "'File: " & f1.name & vbcrlf 
					elseif left(vData(ii),len(KeyPublicSub))=KeyPublicSub or left(vData(ii),len(KeyPublicFunction))=KeyPublicFunction then
						bInRoutine=True
						if left(vData(ii),len(KeyPublicSub))=KeyPublicSub then
							szTypeRoutine="Sub"
						else
							szTypeRoutine="Function"
						end if
		
						'comment routine
						if left(vdata(ii-1),1)="'" then szDataModule=szDataModule & vdata(ii-1) & vbcrlf
						szRoutine= vdata(ii) & vbcrlf
					elseif bInRoutine then
						if Right(trim(vData(ii)),1) = "_" then
							'param routine in more line
							szRoutine=szRoutine & vdata(ii) & vbcrlf
						else
							'param routine end
							if szTypeRoutine="Sub" then
								szDataModule=szDataModule & szRoutine 
							else
								szTypeReturn=trim(mid(szRoutine,instrrev(szRoutine,"As")+2))  'get type return
								szTypeReturn=replace(szTypeReturn,vbcrlf,"")
								if ubound(filter(vOutFunctionValid,szTypeReturn)) <=-1 then
									'comment type return
									szDataModule=szDataModule & mid(szRoutine,1,instrrev(szRoutine,")")) & " '" & trim(mid(szRoutine,instrrev(szRoutine,"As")-1)) 
								else
									szDataModule=szDataModule & szRoutine 
								end if
							end if

							'create call
							if szTypeRoutine="Sub" then
								'sub
								szTemp=trim(mid(szRoutine,len(KeyPublicSub)+1))
								szTemp=ParseArg(szTemp)
								szDataModule=szDataModule & Tab & "Call " & szModuleName & "." & szTemp & vbcrlf
							else
								'function 
								szTemp=trim(mid(szRoutine,len(KeyPublicFunction)+1))
								szTemp=mid(szTemp,1,instrrev(szTemp,")"))				'truncate type return
								szTemp=ParseArg(szTemp)
								szTemp= Tab & mid(szTemp,1,instr(szTemp,"(")-1) & " = " &  szModuleName & "." & szTemp & vbcrlf

								'comment output non valid in class
								if ubound(filter(vOutFunctionValid,szTypeReturn)) <=-1 then szTemp= "'" & sztemp
								szDataModule=szDataModule & sztemp 
							end if

							'end routine
							szDataModule=szDataModule & "End " & szTypeRoutine & vbcrlf & vbcrlf
							bInRoutine=False
							szRoutine=""
						end if
					elseif left(vData(ii),len(KeyGlobalConst))=KeyGlobalConst or left(vData(ii),len(KeyPublicConst))=KeyPublicConst then
						'global/public const
						szTemp=mid(vData(ii),len(KeyGlobalConst)+1)
						szTemp=mid(sztemp,1,instr(szTemp," ")-1)
						szConst=szConst & szTemp & vbcrlf
					elseif left(vData(ii),len(KeyPublicEnum))=KeyPublicEnum then
						'enumeration convert in constatnt			
						ii=ii+1
						while vData(ii) <> KeyEndEnum
							szTemp=trim(vdata(ii))
							szTemp=trim(mid(sztemp,1,instr(szTemp," ")))
							if len(szTemp) > 0 then szConst=szConst & szTemp & vbcrlf
							ii=ii+1
						wend
					end if
				next

		end select
   	Next

	'create function for const/enumerate
szDataAdd=szDataAdd & "'return value of the const/enumerate by name" & vbcrlf & _
		"Public Function ConstByName(szName As String)" & vbcrlf & _
		"" & vbcrlf & _
		"  Select Case LCase(szName)" & vbcrlf 

	vData=split(szConst,vbcrlf)
	for ii=0 to ubound(vData)-1
		szDataAdd=szDataAdd & "    case LCase(" & Quote & vData(ii) & Quote & ")" & vbcrlf
		szDataAdd=szDataAdd & "      ConstByName=" & vData(ii) & vbcrlf
		szDataAdd=szDataAdd & "" & vbcrlf
	next

	szDataAdd=szDataAdd & "  End Select" & vbcrlf & _
		"End Function" & vbcrlf & vbcrlf


	'echo data
	wscript.echo szHeaderClass
	wscript.echo szDataForm
	wscript.echo szDataModule
	wscript.echo szDataAdd
End Sub

'drop defination variable
Function ParseArg(szArg)
dim szTemp,vData,ii,szRoutine,szVar

	szTemp=trim(szArg)
    szRoutine=mid(szTemp,1,instr(szTemp,"("))
    szTemp=mid(szTemp,instr(szTemp,"(")+1)
    szTemp=mid(szTemp,1,instrrev(szTemp,")")-1)

	if len(sztemp) >0 then
		vData=split(szTemp)
		for ii=0 to ubound(vData)
			if vData(ii)="Optional" then
			elseif vData(ii)="ByVal" then
			elseif vData(ii)="ByRef" then
			elseif vData(ii)="As" then
				ii=ii+1
			elseif vData(ii)="=" then
				ii=ii+1
			else
				szVar= szVar & vData(ii) & ","
			end if
		next 
		szTemp=szRoutine & left(szVar,len(szvar)-1) & ")"
	else
		szTemp=szArg
	end if
	ParseArg=szTemp
End Function


'''''''''''''''''''''''''''''''''''
Main()
