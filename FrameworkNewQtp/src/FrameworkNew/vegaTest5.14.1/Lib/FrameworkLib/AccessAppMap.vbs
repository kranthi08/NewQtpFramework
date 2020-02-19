''
' @# author DSTWS
' @# Version 7.0.7

	''
    ' This function connects to App Map  and gets the testObject constructed for the UIName.
	' @author DSTWS
	' @param appdataSource - String the path of AppMap MSAccess dB along with filename
	' @param UIName - String specifying the UIName of the object in the AppMap sheet
	' @return String for constructed TestObject for given UIName.
	' @Modified By DT77734
	' @Modified on: 16 Jul 2009 

Function AccessAppMapDB(appdataSource,UIName)
	Dim objArray() 
	Const adCmdText=1
	Const adStateOpen=1
	Const adOpenKeySet=1
	Const adUseClient=3
	Const adLockOptimistic=3
	Const adOpenStatic = 3
	On error resume next
	Dim conAppMap
	Dim cmdQuery
	Dim rsSup
	Dim blnFound
	blnFound=False
	Err.Clear
	Set conAppMap=CreateObject("ADODB.Connection")
	If Err.Number<>0 Then
		Reporter.ReportEvent micWarning,"Error occured while connecting to the AppMap DB",Err.Description
		WriteToEvent("Fail" & vbtab & "Error occured while connecting to the AppMap DB")
		Err.Clear
		Set conAppMap=Nothing
	End If
    conAppMap.ConnectionString= "DRIVER={Microsoft Excel Driver (*.xls)};DBQ=" & appdataSource & ";Readonly=True"
	conAppMap.Open
	Set objRecordset=CreateObject("ADODB.Recordset")
	ObjRecordset.CursorLocation=3      
	err.clear
    'ObjRecordset.Open "Select * From [KeywordRepository$]  Where [UIName]="&"'"&UIName&"'",conAppMap, 1, 3
	If Cstr(Environment("WindowName"))<>Null Or Cstr(Environment("WindowName"))<>"" Or Cstr(Environment("WindowName"))<>Empty Then		'Added to get the values by making window name mandatory	Fanweb 7.0.11
		ObjRecordset.Open "Select * From [KeywordRepository$]  Where [UIName]="&"'"&UIName&"'"&" and [WindowName]="&"'"&Cstr(Environment("WindowName"))&"'",conAppMap, 1, 3
	Else
		ObjRecordset.Open "Select * From [KeywordRepository$]  Where [UIName]="&"'"&UIName&"'",conAppMap, 1, 3
	End If

	'ObjRecordset.Open "Select * From [KeywordRepository$]  Where [UIName]="&"'"&UIName&"'"&"and [WindowName]="&"'"&Cstr(Environment("WindowName"))&"'",conAppMap, 1, 3
'	If err.Number<>0 Then
'		Reporter.ReportEvent micFail,"There is no AppMap sheet in the excel or no UIName column in the AppMap","Please verify and run again"
'		WriteToEvent("Fail" & vbtab & "There is no AppMap sheet in the excel or no UIName column in the AppMap")
'		ExitTest
'	End If
	err.clear
	Environment("ParentObject")=CStr(ObjRecordset.Fields("ParentObject"))
	AccessAppMapDB=Cstr(ObjRecordset.Fields("ChildObject"))

	If err.number<>0 or IsEmpty(AccessAppMapDB) Then 
        		
		ObjRecordset.close
		Err.clear

		ObjRecordset.Open "Select * From [KeywordRepository$]  Where [UIName]="&"'"&UIName&"'",conAppMap, 1, 3
		Environment("ParentObject")=CStr(ObjRecordset.Fields("ParentObject"))
		AccessAppMapDB=Cstr(ObjRecordset.Fields("ChildObject"))
	End If

	If err.number<>0 and IsEmpty(AccessAppMapDB) Then 
		Environment("ParentObject") =""
		AccessAppMapDB = "" 
        Call ReportResult (UIName , "UIName should exists in the AppMap", "","There are no UIName in the AppMap" , "Failed" ,objParent)
		Reporter.ReportEvent micFail,"There are no UIName specified in the AppMap","Please specify the UIName and run again"
		WriteToEvent("Fail" & vbtab & "There are no UIName specified in the AppMap")
	elseif err.number<>0 then
		  Environment("ParentObject") =""
       End If
	set ObjRecordset=Nothing
	conAppMap.Close
	Set conAppMap=Nothing
End Function
