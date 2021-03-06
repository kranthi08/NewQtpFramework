''
' @# author DSTWS
' @# Version 7.0.7
	
	''
	'  This function GetDataFromTestDB connects to Test dB and gets the test data for the corresponding Data_ID
	' @author DSTWS
	' @param testdataSource string specifying the path of test data MSAccess dB along with filenames
	' @param ValueToEnter string contains the Data_Id which represents a specific test data in test data Db.
	' @Return Value: 1. Data_Value - Test data corresponding to the input Data_Id
	' @ Remarks: '  The format of the input is "Query_Data_Id","Query" is the keyword used by driver script to
	'  distinguish the test data from the test case specific input values. Data_Id is the unique value
	'  in Test Data dB which has a corresponding test data under Data_Value column.
	' @Modified By DT77734, DT77742
	' @Modified on: 16 Jul 2009 , 4 Aug 2009 

Function GetDataFromTestDB(testdataSource,ValueToEnter)
	Dim conAppMap
	Dim cmdQuery
	Dim rsSup
	If Left(LCase(Trim(ValueToEnter)),6)="query_" Then
		Data_Id=Mid(ValueToEnter,InStr(LCase(Trim(ValueToEnter)),"query_")+6)
	end If
	Err.Clear
	Set objTDConn=CreateObject("ADODB.Connection")
	If Err.Number<>0 Then
		Reporter.ReportEvent micWarning,"Error occured while connecting to the TestData DB",Err.Description
		WriteToEvent("Warning" & vbtab & "Error occured while connecting to the TestData DB")
		Err.Clear
		Set objTDConn=Nothing
	End If
	objTDConn.ConnectionString= "DRIVER={Microsoft Excel Driver (*.xls)};DBQ="& testdataSource & ";Readonly=True"
	objTDConn.Open
	Set objRecordset=CreateObject("ADODB.Recordset")
	ObjRecordset.CursorLocation=3                     
	ObjRecordset.Open "Select  *  From [TestData$]  Where [Data_Id]="&"'"& Data_Id &"'",objTDConn, 1, 3
	If ObjRecordset.recordcount>0 then   
		GetDataFromTestDB=ObjRecordset.Fields("Data_Value")
	Else
	   Reporter.ReportEvent micFail,"There is  no Data_Id  specified in the TestData","Please specify the Data_Id and run again"
		WriteToEvent("Fail" & vbtab & "There is  no Data_Id  specified in the TestData")
		GetDataFromTestDB=-1
		ExitTest
	End If
	Set ObjRecordset=Nothing
	objTDConn.close
	Set objTDConn=Nothing
End Function




''
	' This function is to Set the value in the Edit Filed.
	' @author DSTWS
	' @param testdataSource String specifying the path of the test data source file
	' @param sheetnames String specifying test data index names
	' @param keyword String specifying the keword name.
	' @param parameter String specifying  the name of the test data parameter
	' @Modified By DT77734, DT77742
	' @Modified on: 16 Jul 2009 , 12 Aug 2009


Function GetDataFromExcel(testdataSource,sheetnames,keyword,parameter)

   strOriginalParam=parameter

If VerifyEnvVariable("DynamicDataFromTestData") Then

   If Left(LCase(Trim(parameter)),5)="svar_"  Then											'-------------Added by Trao, TAF10 feature
		   If  LCase(Trim(Environment("DynamicDataFromTestData")))<>"yes" Then         'If DynamicDataFromTestData set to 'Yes' in settings, It will look for test data column to get the data , rather than return value from test step
			   GetDataFromExcel=parameter																		'If the test data coumn does not found, it will take the return value from the test step. 
			   Exit Function
			Else
			   parameter=Mid(parameter,InStr(LCase(Trim(parameter)),"svar_")+5)
		   End If
   End If
Else
     GetDataFromExcel=parameter
	 Exit Function
End If

   	If Left(LCase(Trim(parameter)),6)="param_" Then
		parameter=Mid(parameter,InStr(LCase(Trim(parameter)),"param_")+6)
	End If

    Err.Clear
	Set objTDConn=CreateObject("ADODB.Connection")
	If Err.Number<>0 Then
		Reporter.ReportEvent micWarning,"Error occured while connecting to the TestData DB",Err.Description
		WriteToEvent("Warning" & vbtab & "Error occured while connecting to the TestData DB")
		Err.Clear
		Set objTDConn=Nothing
	End If
	''TAF 10.1 code modification Start
	'objTDConn.ConnectionString= "DRIVER={Microsoft Excel Driver (*.xls)};DBQ="& testdataSource & ";Readonly=True"
    objTDConn.ConnectionString= "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & testdataSource &  ";" & "Extended Properties=""Excel 8.0;ImportMixedTypes=Text;IMEX=1"""
	''TAF 10.1 code modification End
		on error resume next 	
	err.clear										
	objTDConn.Open
	If err.number<>0 Then		
		GetDataFromExcel=-1
       	Exit function						
	End If
	Set objRecordset=CreateObject("ADODB.Recordset")
	ObjRecordset.CursorLocation=3                        'set the cursor to use adUseClient - disconnected recordset
	If sheetnames="" Then
		GetDataFromExcel=-1
		Environment("DataSheetExist") = False
		Reporter.ReportEvent micWarning,"TestData","Test Data Sheet is Blank"
		'Call ReportResult  (strUIName, strExpectedResult,strValueToEnter ,"Selected item:"&strValueToEnter &"in the radio button"&strUIName , "Passed" ,objParent)
        Exit Function
	End If
	arrSheetNames=Split(sheetnames,";")
	arrkeyword=Split(keyword,";")
	
	queryStatement=""
	Set ObjRecordset=Nothing
	objTDConn.close
	Set objTDConn=Nothing
	If  parameter="" Then
		For sheet=0 to Ubound(arrSheetNames)
			'Environment("TestDataSheetName") = arrSheetNames(sheet)	'Hima
			blnDataSheetExist = VerifySheetExists(testdataSource,arrSheetNames(sheet))
			If  not(blnDataSheetExist) Then
				GetDataFromExcel=-1
				Environment("DataSheetExist") = False		
		 		Exit Function
			End If
			If LCase(Trim(arrkeyword(0)))="all" Then
				arrkeyword(0)="*"
				queryStatement =queryStatement & "Select *  From "& "["& arrSheetNames(sheet) &"$] Where Keyword Is NOT NULL"
			Else
				queryStatement =queryStatement & "Select *  From "& "["& arrSheetNames(sheet) &"$] Where Keyword='" & Join(arrkeyword,"'Or Keyword='") &"'"
			End If
			
			If  sheet <> Ubound(arrSheetNames)Then
				queryStatement=queryStatement &"UNION ALL "
			End If
		Next
	Else
		If  keyword="" Then
			For sheet=0 to Ubound(arrSheetNames)
            	queryStatement =queryStatement &"Select "&parameter &"  From "& "["& arrSheetNames(sheet) &"$] "
				intFlag=1
				If  sheet <> Ubound(arrSheetNames)Then
					queryStatement=queryStatement &"UNION ALL "
				End If
			Next
		Else
			For sheet=0 to Ubound(arrSheetNames)
				If LCase(Trim(arrkeyword(0)))="all" Then
					arrkeyword(0)="*"
					queryStatement =queryStatement & "Select "&parameter &"  From "& "["& arrSheetNames(sheet) &"$] "
					'intFlag=1
				Else
					queryStatement =queryStatement &"Select "&parameter &"  From "& "["& arrSheetNames(sheet) &"$] Where Keyword='" & Join(arrkeyword,"'Or Keyword='") &"'"
				End If
				intFlag=1
				If  sheet <> Ubound(arrSheetNames)Then
					queryStatement=queryStatement &"UNION ALL "
				End If
			Next
		End If
	End If
	Set objTDConn=CreateObject("ADODB.Connection")
	''TAF 10.1 code modification Start
	'objTDConn.ConnectionString= "DRIVER={Microsoft Excel Driver (*.xls)};DBQ="& testdataSource & ";Readonly=True"
	objTDConn.ConnectionString= "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & testdataSource &  ";" & "Extended Properties=""Excel 8.0;ImportMixedTypes=Text;IMEX=1"""
	''TAF 10.1 code modification End
	objTDConn.Open
	Set objRecordset=CreateObject("ADODB.Recordset")
	objRecordset.CursorLocation=3   
	objRecordset.Open queryStatement ,objTDConn, 1, 3
	If intFlag=1 and  objRecordset.recordcount=0 Then
		Reporter.ReportEvent micWarning, "Get data from Excel", "data is not present for the column " & parameter
		'Call ReportResult("", "Should get the data from excel sheet","",ErrorHandler(300,parameter,""), "Failed","")
		Call WriteToEvent("Failed" & vbtab &"data is not available for the column. " & parameter &"check for existance of special characters in UIname/Coulmn name")
		intFlag = 0
		If Left(LCase(Trim(strOriginalParam)),5)="svar_"  Then
			GetDataFromExcel=strOriginalParam
		Else
		  GetDataFromExcel=-1
		End If
	
       Exit Function	
	ElseIf objRecordset.recordcount=0 then
	   GetDataFromExcel=-1
       Exit Function
	Else
		arrParams=split(objRecordset.GetString,vbCr )
		'MsgBox arrParams
		GetDataFromExcel=arrParams
		If  Environment("IsDataDriven")="no" Then
			Environment("NoOfIterations") =   ubound(arrkeyword)+1
		else
		    Environment("NoOfIterations")=objRecordset.recordcount
		End If
		
	End If
	Set objRecordset=Nothing
	objTDConn.close
	Set objTDConn=Nothing
End Function



''
	' This function is to Set the value in the Edit Filed.
	' @author DT77734
	' @param strFilePath String specifying the path of the file
	' @param strSheet String specifying sheet name.

Public Function  VerifySheetExists(strFilePath,strSheet)
''   Services.StartTransaction "FetchStart"
  ''TAF 10.1 new code Start
   VerifySheetExists = False
   If not isobject(appExcel) Then
	   Set appExcel = CreateObject("Excel.Application")
   End If
    Set objWorkbook = appExcel.Workbooks.Open(strFilePath) 
	intSheetCount = objWorkbook.Sheets.count  
    For intSheetCounter=1 to intSheetCount  
		If Lcase(trim(objWorkbook.Sheets(intSheetCounter).name)) = Lcase(trim(strSheet)) then 
			VerifySheetExists = True 
            Exit for 
		End if  
	Next  
	objWorkbook.Close
	'TAF 10.1 new code End
''Services.EndTransaction "FetchStart"

''''''''''''''' Viswa ''''''''''''''''''''''''''
'
''Services.StartTransaction "FetchStart"
'	Set objTCConn=CreateObject("ADODB.Connection")
'	''''''''''' Changes 9A - Start ''''''''''''''''''''''''
'	If Round(Trim(Environment("ExcelVersion")),0) = 12 Or Round(Trim(Environment("ExcelVersion")),0) = 14 Then
'		objTCConn.ConnectionString= "Provider=Microsoft.ACE.OLEDB.12.0;" & "Data Source=" & strFilePath & ";" & " Extended Properties=Excel 12.0;Persist Security Info=False"
'	Else
'		objTCConn.ConnectionString= "DRIVER={Microsoft Excel Driver (*.xls)};DBQ="& strFilePath & ";Readonly=True"
'	End If 
'	''''''''''' Changes 9A - End ''''''''''''''''''''''''''
'	On Error resume next
'	objTCConn.Open
'	sSQL = "Select * from ["&Lcase(trim(strSheet))&"$]"
'	Set objRecordset=CreateObject("ADODB.Recordset")
'	objRecordset.CursorLocation=3
'	objRecordset.Open sSQL,objTCConn,1,3
'	If Err.Number = 0 Then
'		VerifySheetExists = True 
'	Else
'		VerifySheetExists = False
'	End If
'	Set objRecordset=Nothing
'	Set objTDConn=Nothing
''Services.EndTransaction "FetchStart"
'
''''''''''''''' Viswa '''''''''''''''''''''
End Function

''
' @# author DSTWS
' @# Version 7.0.7
	
	''
	' This function is to Set the value in the Edit Filed.
	' @author DSTWS
	' @param resKey String specifying the result key name
	' @param resItem String specifying the result item name
	' @Modified By DT77734
	' @Modified on: 16 Jul 2009 
    
Dim mobjDict
Set oActivityResult=CreateObject("Scripting.Dictionary")

Public Function AddRes(resKey,resItem)
	ActivityResult.Add resKey, resItem'mobjDicto
End Function

'*******************************************************************************************************************

Function GetDataRows(testdataSource,sheetnames)
   	
	Err.Clear
	Set objTDConn=CreateObject("ADODB.Connection")
	
	If Err.Number<>0 Then
		Reporter.ReportEvent micWarning,"Error occured while connecting to the TestData DB",Err.Description
		WriteToEvent("Warning" & vbtab & "Error occured while connecting to the TestData DB")
		Err.Clear
		Set objTDConn=Nothing
	End If
	'TAF10.1 modification start
	'objTDConn.ConnectionString= "DRIVER={Microsoft Excel Driver (*.xls)};DBQ="& testdataSource & ";Readonly=True"
    objTDConn.ConnectionString= "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & testdataSource &  ";" & "Extended Properties=""Excel 8.0;ImportMixedTypes=Text;IMEX=1"""
	'TAF10.1 modification end
	on error resume next 	
	err.clear										
	objTDConn.Open
	If err.number<>0 Then		
		GetDataRows=-1
       	Exit function						
	End If
	Set objRecordset=CreateObject("ADODB.Recordset")
	ObjRecordset.CursorLocation=3                        'set the cursor to use adUseClient - disconnected recordset
	If sheetnames="" Then
		GetDataRows=-1
        Reporter.ReportEvent micWarning,"TestData","Test Data Sheet is Blank"
        Exit Function
	End If
	arrSheetNames=Split(sheetnames,";")
    Set ObjRecordset=Nothing
	objTDConn.close
	Set objTDConn=Nothing

			blnDataSheetExist = VerifySheetExists(testdataSource,arrSheetNames(sheet))
			If  not(blnDataSheetExist) Then
				GetDataRows=-1
				Environment("DataSheetExist") = False		
		 		Exit Function
			End If
				queryStatement =queryStatement & "Select Keyword  From "& "["& arrSheetNames(sheet) &"$] Where Keyword Is NOT NULL"
	Set objTDConn=CreateObject("ADODB.Connection")
	'TAF10.1 modification start
	'objTDConn.ConnectionString= "DRIVER={Microsoft Excel Driver (*.xls)};DBQ="& testdataSource & ";Readonly=True"
	objTDConn.ConnectionString= "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & testdataSource &  ";" & "Extended Properties=""Excel 8.0;ImportMixedTypes=Text;IMEX=1"""
	'TAF10.1 modification end
	objTDConn.Open
	Set objRecordset=CreateObject("ADODB.Recordset")
	objRecordset.CursorLocation=3   
	objRecordset.Open queryStatement ,objTDConn, 1, 3
	If intFlag=1 and  objRecordset.recordcount=0 Then
		Reporter.ReportEvent micWarning, "Get data rows", "data is not present for the column keywords"
		Call WriteToEvent("Failed" & vbtab &"data is not available for the column keywords. check for existance of keywords Coulmn name")
		intFlag = 0
		GetDataRows=-1
       Exit Function	
	ElseIf objRecordset.recordcount=0 then
	   GetDataRows=-1
       Exit Function
	Else
		arrParams=split(objRecordset.GetString,vbCr )
		GetDataRows=arrParams
	End If
	Set objRecordset=Nothing
	objTDConn.close
	Set objTDConn=Nothing
End Function
