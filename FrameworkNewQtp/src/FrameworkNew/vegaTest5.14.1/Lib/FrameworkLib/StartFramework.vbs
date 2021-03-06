''
' @# author DSTWS
' @# Version 7.0.7a
' @# Version 7.0.8 - Revision History : Changes made to DecParametrization() inlcude "param%_" feature
' @# Version 7.0.8 - Revision History : Changes made to GetDynamicParameter() inlcude "GetDynamicParam" environment variable
' @# Version 7.0.9- Modified by DT 77471: Added environment variables for recovery to handle popup recovery @Bug 63@ line 302,305
' @# Version 7.0.9- Modified by DT 77471: Disable the QTP results for optional step
' @# Version 7.0.9- Modified by DT 77471: Enable QTP results at the beginning of step execution
' @# Version 7.0.9 -Revision History : Enhancement to Version feature (See below)
' @# Version 7.0.9 -Revision History : 1.Handling vesrions like "3V5", "3V10", "3V8","3V8c" etc
' @# Version 7.0.9 -Revision History : 2.If version is a derived version dynamic parameter should be original version , version compare should be based derived version.
' @# Version 7.0.9 -Revision History : Handled if the runstatus column not there in the test script sheet
' @# Version 7.0.9A Keys Added
Dim fso
Set fso=CreateObject("scripting.filesystemobject")

Set DictParentObj = CreateObject("Scripting.Dictionary")
Set DictChildObj = CreateObject("Scripting.Dictionary")

	
	''
	' This function DecParametrization is to declare the parametrization to the test parameters
	' @author DSTWS
	' @param param String specifying whether a particular field to be parametrized Yes or not No.
	' @param TestDataPath String specifying the path of the test data sheet.
	' @param testdataSourceToEnter String specifying the value to Set.
	' @Modified By DT77734, DT77742
	' @Modified on: 03 Aug 2009

Function DecParametrization(param,TestDataPath,testdataSource,UIName)
	Dim paramSheet, paramDB
	If Left(LCase(Trim(param)),6)="param_" Then
		Environment("DataFromDataSheet") = "True"
		param_data=GetDataFromExcel(testdataSource,Environment("TestDataSheetName"),Environment("Keywords"),param)
		DecParametrization=param_data
	ElseIf Left(LCase(Trim(param)),5)="svar_" Then    '-------------Added by Trao, TAF10 feature
	   Environment("DataFromDataSheet") = "True"             'If DynamicDataFromTestData set to 'Yes' in settings, It will look for test data column to get the data , rather than return value from test step
																										 	 'If the test data coumn does not found, it will take the return value from the test step. 
		param_data=GetDataFromExcel(testdataSource,Environment("TestDataSheetName"),Environment("Keywords"),param)
		DecParametrization=param_data
  	ElseIf Left(Lcase(Trim(param)),6) = "param#" Then
		param="param_"&UIName
		Environment("DataFromDataSheet") = "True"
		param_data=GetDataFromExcel(testdataSource,Environment("TestDataSheetName"),Environment("Keywords"),param)
		DecParametrization=param_data
	ElseIf  Left(LCase(Trim(param)),6)="query_" Then
		paramDB=GetDataFromTestDB(testdataSource,param)
       		DecParametrization=paramDB
	ElseIf  Left(LCase(Trim(param)),6)="param%" Then
		varID=Mid(param,8,len(param))
		param="param_"&Environment(varID)
		Environment("DataFromDataSheet") = "True"
		param_data=GetDataFromExcel(testdataSource,Environment("TestDataSheetName"),Environment("Keywords"),param)
		DecParametrization=param_data
	ElseIf Left(LCase(Trim(param)),6)="space_" Then
		NoofSpaces = Replace(LCase(Trim(param)),"space_","")
		NoofSpaces=Cint(Trim(NoofSpaces))

		For i = 1 to NoofSpaces
			Spc = Spc&chr(32)
		Next
		DecParametrization=Spc
	Else
		If Environment("GetDynamicParam") Then
          		 DecParametrization=GetDynamicParameter(param)
		Else
           		DecParametrization=param
		End If
	End If 
End Function



Function RegExpTest(patrn, strng)
   Dim regEx, Match, Matches   ' Create variable.
   Set regEx = New RegExp   ' Create a regular expression.
   regEx.Pattern = patrn   ' Set pattern.
   regEx.IgnoreCase = True   ' Set case insensitivity.
   regEx.Global = True   ' Set global applicability.
   Set Matches = regEx.Execute(strng)   ' Execute search.
   For Each Match in Matches   ' Iterate Matches collection.
	         RetStr = RetStr & Match.Value & "!" 
   Next
   RegExpTest = RetStr
End Function

Function GetDynamicParameter(strParameter)
  ' msgbox "GETDynamicParameter " & strParameter
	If not(Environment("GetDynamicParam")) Then
		GetDynamicParameter = strParameter
		Exit Function
	End If
				strGerRegExpr = RegExpTest("%.*?%", strParameter)
		If  strGerRegExpr <>"" and  strGerRegExpr <> false Then
							arr = split(strGerRegExpr,"!")
				For i= lbound(arr) to ubound(arr)-1
						ReDim Preserve brr(i)
						brr(i) = Replace( arr(i),"%","")
				Next
        		For i= lbound(brr) to ubound(brr)
					If  VerifyEnvVariable(brr(i)) Then
						 strParameter = Replace(strParameter,"%" & Trim(brr(i)) & "%" ,Environment(Trim(brr(i))))
					End If
				Next
		End If
				GetDynamicParameter  =  strParameter
End Function


Function VerifyEnvVariable(EnvName)
   Dim tempEnvl
   On error resume next
   tempEnvl= Environment(EnvName)
   If Err.Number<>0 Then
	   VerifyEnvVariable=FALSE
	Else
		VerifyEnvVariable=TRUE
   End If
   On error GoTo 0
End Function







	''
	' This function is to introduce a tesmperory excel path and copy the test case script sheet when Imposheet fails to import the test script 
	' sheet to destination sheet
	' @author DSTWS
	' @param SheetSource String specifying the test case script
	' @param SheetDestination String specifying the sheet in the data table to copy the source sheet.
	' @Modified By DT77734
	' @Modified on: 16 Jul 2009 
	
Function ImportFromTemp(SheetSource,SheetDestination)
	tcPath=DataTable("TestCaseFilePath","Global")
	tempPath=StrReverse(mid(StrReverse(tcPath),Instr(StrReverse(tcPath),"\")))&"temp.xls"
	Set fso = CreateObject("Scripting.FileSystemObject")
	If fso.FileExists(tempPath) then
		fso.DeleteFile(tempPath)
	End If
	Set fso=Nothing
	Set objExcel = CreateObject("Excel.Application")
	Set objWorkbook1= objExcel.Workbooks.Open(tcPath)
	Set objWorkbook2= objExcel.Workbooks.add
	sname=objWorkbook1.Worksheets(SheetSource).name
	objWorkbook1.Worksheets(sname).UsedRange.Copy
	objWorkbook2.Worksheets("Sheet1").name=sname
	objWorkbook2.Worksheets(sname).Range("A1").PasteSpecial Paste =xlValues
	objWorkbook1.save
	objWorkbook2.saveas tempPath
	objWorkbook1.close
	objWorkbook2.close
	set objExcel=nothing
	DataTable.ImportSheet tempPath,SheetSource,SheetDestination
End Function




	''
	' This function is to import the external resources like Test Case Script, Test Data sheet and AppMap sheet.
	' @author DSTWS
	' @Modified By DT77742, DT77734
	' @Modified on: 13 Aug 2009

Function startTest()

    DictParentObj.RemoveAll         'Added by Srikanth for multiple appmap
   DictChildObj.RemoveAll

	Dim Param1,Param2,Param3,Param4,Param5,Param6,Param7,Param8,Param9,Param10
	TestCasePath=environment("TestCasePath")
	Environment("DataFromDataSheet") = "False"
    SheetSource=DataTable.GlobalSheet.GetParameter("TestCaseSheetName")
	SheetDestination="TestScriptSheet"

	DataTable.AddSheet SheetDestination
	
	'----------------------------------------Dict APPMAP Change---------------------------------------------------------------
	On error resume next
	Dim conAppMap
	Dim cmdQuery
	Dim rsSup
	Dim blnFound
	blnFound=False
	Err.Clear

'	*************************** Modified by Srikanth 0n 06/04/2012 to read Multiple AppMap
For appread = 1 to Environment("AppMapCount")
       Err.Clear
		Set conAppMap=CreateObject("ADODB.Connection")
	If Err.Number<>0 Then
		Reporter.ReportEvent micWarning,"Error occured while connecting to the AppMap DB",Err.Description
		WriteToEvent("Fail" & vbtab & "Error occured while connecting to the AppMap DB")
		Err.Clear
		Set conAppMap=Nothing
	End If
    conAppMap.ConnectionString= "DRIVER={Microsoft Excel Driver (*.xls)};DBQ="& Environment("AppMapPath"&appread) & ";Readonly=True"
	conAppMap.Open
	
	Set RecordSet = conAppMap.Execute( "Select * From [KeywordRepository$]" )
	If err.Number<>0 Then
		Reporter.ReportEvent micFail,"There is no AppMap sheet in the excel or no UIName column in the AppMap","Please verify and run again"
		WriteToEvent("Fail" & vbtab & "There is no AppMap sheet in the excel or no UIName column in the AppMap")
		ExitTest
	End If
	'Modified by Srikanth on 07172012 to get objectnames from AppMap based on Window and withoutwindow
	On Error Resume Next
    Set RecordSet = conAppMap.Execute( "Select * From [KeywordRepository$]  Where [WindowName]")
	If err.Number<>0 Then
		Environment("AppMapType") ="WithOutWindowName"
	Else
		Environment("AppMapType") = "WithWindowName"
		Err.clear
		'On Error Goto 0
	End If

	TotalRowsCnt = 0
    RecordSet.MoveFirst
    Do Until RecordSet.EOF
        TotalRowsCnt = TotalRowsCnt+1
        RecordSet.MoveNext
    Loop
    
    If TotalRowsCnt > 0 Then    	
    	RecordSet.MoveFirst
    	For AppMapRow = 1 To TotalRowsCnt
			If Environment("AppMapType") ="WithOutWindowName" Then
				DictParentObj.Add RecordSet.fields("UIName").Value, RecordSet.fields("ParentObject").Value
				DictChildObj.Add RecordSet.fields("UIName").Value, RecordSet.fields("ChildObject").Value
			Else
				DictParentObj.Add RecordSet.fields("WindowName").Value&"!"&RecordSet.fields("UIName").Value, RecordSet.fields("ParentObject").Value
				DictChildObj.Add RecordSet.fields("WindowName").Value&"!"&RecordSet.fields("UIName").Value, RecordSet.fields("ChildObject").Value	
			End If
    		RecordSet.MoveNext
    	Next
    End If
    Set conAppMap=Nothing
Next
'*************************** Modification Completed by Srikanth 0n 06/04/2012 to read Multiple AppMap
	
	
	
	'***********Load test case from local temp file  7.0.11   Fanweb**********************
    If not Environment("RunFromQC") Then	
		If LoadFromTempLS Then
			fso.CopyFile TestCasePath,Environment("TempLS")&"\"
			arrTempCase = Split(TestCasePath,"\")
			tempCase=Environment("TempLS")&"\"&arrTempCase(Ubound(arrTempCase))
			DataTable.ImportSheet tempCase,SheetSource,SheetDestination      
		Else
			DataTable.ImportSheet TestCasePath,SheetSource,SheetDestination
		End If
		If Err.Number<>0 Then
			DataTable.ImportSheet TestCasePath,SheetSource,SheetDestination
		End If
	Else
		DataTable.ImportSheet TestCasePath,SheetSource,SheetDestination	
	End If
'********************************************************************************************************
	err.clear
	'DataTable.ImportSheet TestCasePath,SheetSource,SheetDestination
	If err.Number=20012 Then
		Call ImportFromTemp(SheetSource,SheetDestination)
		Reporter.ReportEvent micDone,"Imposheet failed and called the ImportFromTemp function","Imposheet failed and called the ImportFromTemp function"
		WriteToEvent("Failed" & vbtab & "Imposheet failed and called the ImportFromTemp function")
		err.clear
	End If
	Environment("TestCaseName")=CStr(DataTable.GetSheet(SheetDestination).GetParameter("TestCaseName"))
	
	If Environment("TestCaseName") = "" Then
		DataTable.GetSheet(SheetDestination).SetCurrentRow(1)
		Environment("TestCaseName")=CStr(DataTable.GetSheet(SheetDestination).GetParameter("TestCaseName"))
	End If
	
	' PRint Environment("TestCaseName")
	rowCount=DataTable.GetSheet(SheetDestination).GetRowCount
	Environment("NoOfIterations")=1

	If Environment("WithoutTestData") Then
		intNoOfIterations=1
	Else
    			arrDataRowsAndData = GetDataFromExcel(Environment("TestDataPath"),Environment("TestDataSheetName"),Environment("Keywords"),"")  ' Argus
	
				if LCase(Trim(Environment("DataDriven")))="yes" Then
					intNoOfIterations=Cint(Environment("NoOfIterations"))	
				Else	
					intNoOfIterations=1	
				End If	
	
	End If

	'TAF 10.1 new code Start
	Environment("NoOfDatarows")=Environment("NoOfDatarows")+intNoOfIterations
	'TAF 10.1 new code End
	
    For ite=  0 to intNoOfIterations-1

			'TAF 10.1 new code Start
			If Environment("BlnCodeGeneration")=1 Then
				Call UpdateScript("'********************"&"Iteration: "&(ite+1)&"  start"&"***********************************************************")   '10.1
				Call UpdateScript("")
			End If
			'TAF 10.1 new code End
			
				

    	Environment("intite")  = ite				' Argus
    	
    	' ARgus Code begining
    		If isarray(arrDataRowsAndData) Then 
			 
				arrDataRow=split(arrDataRowsAndData(ite),VBTAB )
				RowKey = arrDataRow(0)
				Environment("RowKey") = arrDataRow(0)	
							
				'Print   "Ite - " &  Environment.value("intite")+1  & " &"   & " DataRow -" & RowKey 
				Environment("ExecelResultsKeywords") =  "Iteration - " &  Environment.value("intite")+1  & vbtab  & " Key -" & RowKey   & vbtab  & "Test Case Name - " & Environment("TestCaseSheetNameToResults")   ' Argus 

		else
					If Environment("WithoutTestData")=True Then
						'Print   "Ite - " &  Environment.value("intite")+1  & " & "  & " DataRow - Default DataRow"  ' today
				        Environment("ExecelResultsKeywords") =  "Iteration - " &  Environment.value("intite")+1  & vbtab  &  " Key - No Data Row used" 	  & vbtab  & "Test Case Name - " & Environment("TestCaseSheetNameToResults")   ' Argus 
					Else
					'Print   "Ite - " &  Environment.value("intite")+1  & " & "  & " DataRow - Default DataRow"  ' today
				    Environment("ExecelResultsKeywords") =  "Iteration - " &  Environment.value("intite")+1  & vbtab  &  " Key - Default Data Row" 	  & vbtab  & "Test Case Name - " & Environment("TestCaseSheetNameToResults")   ' Argus 
					End If
				

		End If
		If Lcase(Environment("SaveScreenshotsinLocalDrive")) = "yes" Then
			If IsEmpty(Environment("RowKey")) = True Then
					Environment("RowKey") = "0"
			End If	
			Call CreateLocalDriveScreenshotFolder()
		End If
	' ARgus Code end
		
    	Environment.value("RepeatStep")=FALSE			'Initiate the environment variables for recovery
		Environment.value("Counter")=0				
		Environment.value("elsecondition") = False ' by Srikanth 06/18/2014
		Environment.value("reusescript") = False ' by Srikanth 06/18/2014
		
		For i=1 to rowCount
			Environment.value("tcrow") = i ' by Srikanth 06/18/2014
			Reporter.Filter = rfEnableAll 				'To enable the results at the beginning of step execution
			If Environment("RepeatStep") and Environment("Counter")=1 Then		'To roll back the environment variable values
				i=i-1															' To repeat the step only once
				Environment.Value("RepeatStep")=FALSE
				Environment.Value("Counter")=0
			End If
			Environment("StepNumber")=i	+1				'Added to get the step number in Excel results	FanWeb		7.0.11
			Environment("FlagTCStep")=True
            Environment("Optional") = "False"
			DataTable.GetSheet(SheetDestination).SetCurrentRow(i)
			blnExecuteStep = False

			Err.clear
			On Error resume Next
            strParamWindowName  = DataTable.GetSheet(SheetDestination).GetParameter("WindowName")
			If err.number = 0  Then
			If (DataTable.GetSheet(SheetDestination).GetParameter("WindowName"))<>"" Then		'Fanweb 7.0.11  Making window name mandatory
					Environment("WindowName")=Cstr(DataTable.GetSheet(SheetDestination).GetParameter("WindowName")) 'Modification to use Window name
				ElseIf i=1 Then
					Environment("WindowName")=""
				End If
			End If
			On Error Goto 0
            
			On Error Resume Next
			If  Lcase(Trim(DataTable.GetSheet(SheetDestination).GetParameter("RunStatus"))) <> "skip" Then
				If IsEmpty(DataTable.GetSheet(SheetDestination).GetParameter("RunStatus")) Then
					DataTable.GetSheet(SheetDestination).GetParameter("RunStatus")=" "
	                ElseIf Lcase(Trim(DataTable.GetSheet(SheetDestination).GetParameter("RunStatus"))) ="optional" Then
								Environment("Optional")=True
								Reporter.Filter = rfDisableAll					'To disable the QTP results for optional step
				
				End If

				
			
				'Below block of code updated by Rajesh in 7.0.8a version
				TcVerFrm = LCase(Trim(CStr(DataTable.GetSheet(SheetDestination).GetParameter("VersionFrom")))) ' New
				TcVerTo  = LCase(Trim(CStr(DataTable.GetSheet(SheetDestination).GetParameter("VersionTo"))))  ' New
				Version  = LCase(Trim(CStr(Environment("Version")))) ' New
				
				' If its is Derived version , then the VF and VT should be verified against Derived version
				If  VerifyEnvVariable("DerivedVersion") Then
					Version = LCase(Trim(CStr(Environment("DerivedVersion")))) 
				End If
								
				ConvertVersion TcVerFrm, TcVerTo,Version, VF, VT,Ver
				
				If  (TcVerFrm = "") and (TcVerTo = "") Then   ' all 
					blnExecuteStep = True
					elseif  (TcVerFrm = "") and (TcVerTo <>  "") Then        ' upto
					
             			If   (cdbl(Ver) <= cdbl( VT))   Then
										blnExecuteStep = True
						End If
						
                    elseif  (TcVerFrm <>  "") and (TcVerTo =   "") Then   ' later
                    
						If   (cdbl(Ver) >= cdbl(VF))   Then		
										blnExecuteStep = True
						End If 
					elseif  (TcVerFrm <>  "") and (TcVerTo <>  "") Then   ' between
					
						If   (cdbl(Ver) >=cdbl(VF))  and  (cdbl(Ver)  <= cdbl( VT))  Then
										blnExecuteStep = True
						End If
				End If
				'above block of code updated by Rajesh
				err.clear
				If  blnExecuteStep Then
					functionName=DataTable.GetSheet(SheetDestination).GetParameter("Activity")
			        functionName=Cstr(functionName)

			If LCase(Trim(functionName))="endofrow" or LCase(Trim(functionName))="eof"  or LCase(Trim(functionName))="x" Then
				startTest=-1
				Exit For
			ElseIf Left(LCase(Trim(functionName)),6)="reuse_" Then
				ReUseSeverity=DataTable.GetSheet(SheetDestination).GetParameter("Severity")
				ReuableTestcasefilepath =  DataTable.GetSheet(SheetDestination).GetParameter("Data")

			'	TAF10  strat Reuse_XX_YY fix
				'msgbox ReuableTestcasefilepath
'				reusableSheetName =""
'				reusableDetails=Split(functionName,"_")
'				For m = 1 to ubound(reusableDetails)
'					reusableSheetName =  reusableSheetName & reusableDetails(m)
'				Next

				reusableSheetName = Replace(functionName,"reuse_","",1,1,1)

		'	TAF10  End  Reuse_XX_YY fix

				DestSheetName="Reusable"
				DataTable.AddSheet(DestSheetName)
               	Set objFileSys = CreateObject("Scripting.FileSystemObject")
				Set objExcelobj = createobject("Excel.Application") 
				Flag = False
				' Verfiy in the provided test Case source  file
    	if  Trim(ReuableTestcasefilepath) <> ""Then

					' From QC
			If Environment("RunFromQC")  Then

			arrTemp = Split( Environment("vegaTestSuite") & ReuableTestcasefilepath,"\")
			strReuseTestCaseFileName = arrTemp(ubound(arrTemp))
			arrTemp = Split( Environment("vegaTestSuite") & ReuableTestcasefilepath,"\" & strReuseTestCaseFileName)
			strReuseTestCaseFileFolder = arrTemp(lbound(arrTemp))
			'msgbox strReuseTestCaseFileName
			'msgbox strReuseTestCaseFileFolder
            DownloadAttachment  strReuseTestCaseFileName,strReuseTestCaseFileFolder,Environment("WorkingDirectory") & "\TempAutomation","TRUE"
			ReuableTestcasefilepath = Environment("WorkingDirectory") & "\TempAutomation\" & strReuseTestCaseFileName

			Set Fso1 = Createobject("Scripting.filesystemobject")
					If Fso1.FileExists( ReuableTestcasefilepath)  Then
						Set objWorkbook = objExcelobj.Workbooks.Open( ReuableTestcasefilepath)  
					sheetcount = objWorkbook.Sheets.count 
                    For intSheetCounter=1 to sheetcount 
						If objWorkbook.Sheets(intSheetCounter).name = reusableSheetName then 
							Flag=True 
							Exit for
						end if 
					Next 
					else
					    Reporter.ReportEvent micWarning, "Importing Reusable Test Case " & reusableSheetName &"from" & ReuableTestcasefilepath   ,  ReuableTestcasefilepath   & "or  " & " reusable test case " & reusableSheetName  & "does not exist" 
                  End If
			End If
				
					' Form LS 
					' If  file exits then else warming message.
						If not( Environment("RunFromQC"))  Then
					Set Fso1 = Createobject("Scripting.filesystemobject")
					If Fso1.FileExists( Environment("vegaTestSuite") & ReuableTestcasefilepath)  Then
						Set objWorkbook = objExcelobj.Workbooks.Open( Environment("vegaTestSuite") & ReuableTestcasefilepath)  
					sheetcount = objWorkbook.Sheets.count 
                    For intSheetCounter=1 to sheetcount 
						If objWorkbook.Sheets(intSheetCounter).name = reusableSheetName then 
							Flag=True 
							Exit for
						end if 
					Next 
					else
					    Reporter.ReportEvent micWarning, "Importing Reusable Test Case " & reusableSheetName &"from" &Environment("vegaTestSuite") & ReuableTestcasefilepath   , Environment("vegaTestSuite") & ReuableTestcasefilepath   & "or  " & " reusable test case " & reusableSheetName  & "does not exist" 
                  End If
                    
				End If

			
				
					If Flag Then 
						objWorkbook.Close 
						If Environment("RunFromQC")  Then
						'msgbox "taken from Direct path"
                         datatable.ImportSheet  Environment("WorkingDirectory") & "\TempAutomation\" & strReuseTestCaseFileName, reusableSheetName,DestSheetName 
						else
                         datatable.ImportSheet Environment("vegaTestSuite") & ReuableTestcasefilepath, reusableSheetName,DestSheetName 
						End If 

					End If  

				else
				' Verify in the resuable folder
                Set ResueFolder = objFileSys.GetFolder(Environment("ReusableTestCaseFolderPath"))
				Set GetFiles = ResueFolder.Files 
				intFileCount = GetFiles.count-1
                ReDim Preserve arrReuseFile(intFileCount)
				intFileCounter = 0 
				For Each File1 in GetFiles
					arrReuseFile(intFileCounter) = File1.name 
                    intResueTestPath = Environment("ReusableTestCaseFolderPath") &"\" &arrReuseFile(intFileCounter)
                    Set objWorkbook = objExcelobj.Workbooks.Open(intResueTestPath) 
					intSheetCount = objWorkbook.Sheets.count  
                    For intSheetCounter=1 to intSheetCount  
						If objWorkbook.Sheets(intSheetCounter).name = reusableSheetName then 
							Flag=True 
							Exit for 
						end if  
					Next  
					If Flag Then 
						'msgbox "taken from Reusable folder"
						objWorkbook.Close 
                        datatable.ImportSheet intResueTestPath, reusableSheetName,DestSheetName 
						Exit for  
					End If  
					objWorkbook.Close 
					intFileCounter = intFileCounter +1 
				Next 
				End If
        

				' Verify in Current test case file
    			If not Flag Then
					Set objWorkbook = objExcelobj.Workbooks.Open(TestCasePath)
					sheetcount = objWorkbook.Sheets.count 
                    For intSheetCounter=1 to sheetcount 
						If objWorkbook.Sheets(intSheetCounter).name = reusableSheetName then 
							Flag=True 
							Exit for
						end if 
					Next 
                    If Flag Then 
						'msgbox "taken from Current file"
						objWorkbook.Close 
                        datatable.ImportSheet TestCasePath, reusableSheetName,DestSheetName 
                    else 
						objWorkbook.Close 
						Call ReportResult  (reusableSheetName, "Should perform "& reusableSheetName ,"","Failed to perform the opeartion as the Reusable Testcase is not available" , "Failed" ,"")'Hima
                    	Reporter.ReportEvent micFail, "Importing Reusable Test Case "&reusableSheetName , reusableSheetName & " reusable test case does not exist" 
						WriteToEvent("Failed" & vbtab & reusableSheetName & " reusable test case does not exist" )
						Environment("ExecutionStarted") = False
						Exit for
					End If 
                End If 

				Set objExcelobj = Nothing 
				Set objFileSys = Nothing 
				If err.Number=20012 Then
					Reporter.ReportEvent micDone,"Imposheet failed and called the ImportFromTemp function","Imposheet failed and called the ImportFromTemp function"
					WriteToEvent("Failed" & vbtab & reusableSheetName & "Imposheet failed and called the ImportFromTemp function" )
					Call ImportFromTemp(reusableSheetName,DestSheetName)
					err.clear
				End If
				rowCountReuse=DataTable.GetSheet(DestSheetName).GetRowCount
				Environment.value("reusescript") = True ' Srikanth 06/18/2014
                For j=1 To rowCountReuse
				Environment.value("reusetcrow") = j ' Srikanth 06/18/2014
				DataTable.GetSheet(DestSheetName).SetCurrentRow(j)
				blnExecuteStep = False
				On error resume next
			If  Lcase(Trim(DataTable.GetSheet(DestSheetName).GetParameter("RunStatus"))) <> "skip" Then
                If Lcase(Trim(DataTable.GetSheet(DestSheetName).GetParameter("RunStatus"))) ="optional" Then
							Environment("Optional")=true
				End If
				If  (Trim(DataTable.GetSheet(DestSheetName).GetParameter("VersionFrom")) = "") and (Trim(DataTable.GetSheet(DestSheetName).GetParameter("VersionTo")) = "") Then   ' all 
					blnExecuteStep = True
					elseif  (Trim(DataTable.GetSheet(DestSheetName).GetParameter("VersionFrom")) = "") and (Trim(DataTable.GetSheet(DestSheetName).GetParameter("VersionTo")) <>  "") Then        ' upto
             			If  (lcase(Trim(cstr(Environment("Version")))) <= Lcase(Trim(DataTable.GetSheet(DestSheetName).GetParameter("VersionTo"))) )  Then
										blnExecuteStep = True
						End If
                    elseif  (Trim(DataTable.GetSheet(DestSheetName).GetParameter("VersionFrom")) <>  "") and (Trim(DataTable.GetSheet(DestSheetName).GetParameter("VersionTo")) =   "") Then   ' later
						If  (lcase(Trim(cstr(Environment("Version")))) >= Lcase(Trim(DataTable.GetSheet(DestSheetName).GetParameter("VersionFrom"))))  Then		
										blnExecuteStep = True
						End If 
					elseif  (Trim(DataTable.GetSheet(DestSheetName).GetParameter("VersionFrom")) <>  "") and (Trim(DataTable.GetSheet(DestSheetName).GetParameter("VersionTo")) <>  "") Then   ' between
						If  (lcase(Trim(cstr(Environment("Version")))) >= Lcase(Trim(DataTable.GetSheet(DestSheetName).GetParameter("VersionFrom")))) and  (lcase(Trim(cstr(Environment("Version")))) <= Lcase(Trim(DataTable.GetSheet(SheetDestination).GetParameter("VersionTo"))) )  Then
										blnExecuteStep = True
						End If
				End If
				err.clear

				If  blnExecuteStep Then

					On Error resume Next  ' TAF10 start - For TA Desktop AppMap Window name column mandate changes
					Reuse_WindowName=DataTable.GetSheet(DestSheetName).GetParameter("WindowName")
					If err.number = 0 Then
						If (DataTable.GetSheet(DestSheetName).GetParameter("WindowName"))<>"" Then 'Fanweb 7.0.11 Making window name mandatory
							Environment("WindowName")=Cstr(DataTable.GetSheet(DestSheetName).GetParameter("WindowName")) 'Modification to use Window name
						ElseIf j=1 Then
							Environment("WindowName")=""
						End If
					End If
					On Error Goto 0	 ' TAF10 End 

					Reuse_UIName=DataTable.GetSheet(DestSheetName).GetParameter("UIName")
					Reuse_functionName=DataTable.GetSheet(DestSheetName).GetParameter("Activity")
					Reuse_functionName=Cstr(Reuse_functionName)
					testdataSource=DataTable.GlobalSheet.GetParameter("TestDataDBPath")
					datasheetVals=GetDataSheetValues(reusableSheetName,DestSheetName,j)
					Reuse_ExpectedResult=datasheetVals(0,0)
					Reuse_stepResultVar=datasheetVals(0,1)

					If Trim(Len(Reuse_stepResultVar))<>0 Then
						Reuse_stepResultVar=Reuse_stepResultVar
					End If
					Reuse_StopEvent=datasheetVals(0,2)
					R_Param1=datasheetVals(0,4)
					R_Param2=datasheetVals(0,5)
					R_Param3=datasheetVals(0,6)
					R_Param4=datasheetVals(0,7)	
					R_Param5=datasheetVals(0,8)
					R_Param6=datasheetVals(0,9) 'chakri
					R_Param7=datasheetVals(0,10)'chakri
					R_Param8=datasheetVals(0,11)'chakri
					R_Param9=datasheetVals(0,12)'chakri
					R_Param10=datasheetVals(0,13)'chakri
					Reuse_captureScreen=datasheetVals(0,3)
					If Trim(LCase(Reuse_functionName))="endofrow" or Trim(LCase(Reuse_functionName))="eof"  or Trim(LCase(Reuse_functionName))="x" Then
						Exit For
					End If
                    Reuse_Param1=DecParametrization(R_Param1,TestCasePath,testdataSource,Reuse_UIName)
						Reuse_Param2=DecParametrization(R_Param2,TestCasePath,testdataSource,Reuse_UIName)
						Reuse_Param3=DecParametrization(R_Param3,TestCasePath,testdataSource,Reuse_UIName)
						Reuse_Param4=DecParametrization(R_Param4,TestCasePath,testdataSource,Reuse_UIName)
						Reuse_Param5=DecParametrization(R_Param5,TestCasePath,testdataSource,Reuse_UIName)
						Reuse_Param6=DecParametrization(R_Param6,TestCasePath,testdataSource,Reuse_UIName)
						Reuse_Param7=DecParametrization(R_Param7,TestCasePath,testdataSource,Reuse_UIName)
						Reuse_Param8=DecParametrization(R_Param8,TestCasePath,testdataSource,Reuse_UIName)
						Reuse_Param9=DecParametrization(R_Param9,TestCasePath,testdataSource,Reuse_UIName)
						Reuse_Param10=DecParametrization(R_Param10,TestCasePath,testdataSource,Reuse_UIName)

                	If  Not IsArray(Reuse_Param1) Then
						Parameter1=Reuse_Param1
					Else
						Parameter1=Reuse_Param1(ite)
					End If
					'TAF 10.1 code modification Start. If a test data in perticular test data column contains a text with '<> ' seperator, means user wants to give all the parameters required for this test step with <> seperator,
				'instead of  each parameter in seperate test data columns. 
					If  Instr(1,Parameter1,"<>") = 0  Then
							If not  IsArray(Reuse_Param2) Then
									Parameter2=Reuse_Param2
								Else
									Parameter2=Reuse_Param2(ite)
								End If
								If not IsArray(Reuse_Param3) Then
									Parameter3=Reuse_Param3
								else
									Parameter3=Reuse_Param3(ite)
								End If
								If not IsArray(Reuse_Param4) Then
									Parameter4=Reuse_Param4
								else
									Parameter4=Reuse_Param4(ite)
								End If
								If not  IsArray(Reuse_Param5) Then
									Parameter5=Reuse_Param5
								else
									Parameter5=Reuse_Param5(ite)
								End If
								If not  IsArray(Reuse_Param6) Then 'chakri
									Parameter6=Reuse_Param6
								Else
									Parameter6=Reuse_Param6(ite)
								End If
								If not  IsArray(Reuse_Param7) Then 'chakri
									Parameter7=Reuse_Param7
								Else
									Parameter7=Reuse_Param7(ite)
								End If
								If not  IsArray(Reuse_Param8) Then 'chakri
									Parameter8=Reuse_Param8
								Else
									Parameter8=Reuse_Param8(ite)
								End If
								If not  IsArray(Reuse_Param9) Then 'chakri
									Parameter9=Reuse_Param9
								Else
									Parameter9=Reuse_Param9(ite)
								End If
								If not  IsArray(Reuse_Param10) Then 'chakri
									Parameter10=Reuse_Param10
								Else
									Parameter10=Reuse_Param10(ite)
								End If
					Else

									On Error resume Next
									vvr = Split(Parameter1,"<>")
									Parameter1= vvr(0)
									Parameter2 = vvr(1)
									Parameter3 = vvr(2)
									Parameter4 = vvr(3)
									Parameter5 = vvr(4)
									Parameter6 = vvr(5)
									Parameter7 = vvr(6)
									Parameter8 = vvr(7)
									Parameter9 = vvr(8)
									Parameter10 = vvr(9)
					End If
					'TAF 10.1 code modification End. 

                    If Parameter1 = -1 or  Parameter2 =-1 or Parameter3 = -1 or Parameter4 = -1 or Parameter5 = -1 or Parameter6 = -1 or  Parameter7 =-1 or Parameter8 = -1 or Parameter9 = -1 or Parameter10 = -1 Then    'chakri
						Call ReportResult  (Reuse_UIName, Reuse_functionName,R_Param1,"Failed To Perform the opeartion as the Test Data sheet is either empty or not available in Test Data File" , "Failed" ,"")
						WriteToEvent("Failed" & vbtab &"Failed To Perform the opeartion as the Test Data sheet is either empty or not available in Test Data File" )
						If lcase(trim(Reuse_StopEvent))="showstopper" Then  
							Environment("FlagStepFailureOccured")=True
                        	Exit Function   
						else
							Environment("FlagStepFailureOccured")=True
							If  Environment("strTCSeverity")="showstopper" Then
									Exit Function
							End If
						End If 
					Else 
						Call ExecuteFramework(Environment("AppMapPath"),Reuse_UIName,Reuse_functionName,Reuse_ExpectedResult,Reuse_stepResultVar,Reuse_StopEvent,Parameter1,Parameter2,Parameter3,Parameter4,Parameter5,Parameter6,Parameter7,Parameter8,Parameter9,Parameter10,Reuse_captureScreen)
						If not Environment("FlagTCStep") Then   
							If ReUseSeverity="showstopper" Or LCase(trim(Reuse_StopEvent))="showstopper" Or Environment("strTCSeverity")="showstopper"  Then
								Environment("FlagStepFailureOccured")=True
								Exit Function
							End If
						End If
					End If 
					'Writing to exit the start test function if any of the test is failed and its re-use test case is show stopper
					If not(environment("FlagTCStep")) and SeverityName="showstopper" Then   
						Environment("FlagStepFailureOccured")=True
						Exit function   
'						else
						
					End If  
					End If  'rajesh
					End If  'rajesh 
					j = Environment.value("reusetcrow") ' Srikanth 06/18/2014
				Next   
				Environment.value("reusescript") = False ' Srikanth 06/18/2014
			Else
				UIName=DataTable.GetSheet(SheetDestination).GetParameter("UIName")
				testdataSource=DataTable.GlobalSheet.GetParameter("TestDataDBPath")
				datasheetVals=GetDataSheetValues(SheetSource,SheetDestination,i)
				ExpectedResult=datasheetVals(0,0)
				stepResultVar=datasheetVals(0,1)
				If Trim(Len(stepResultVar))<>0 Then
					stepResultVar=stepResultVar
				End If
				StopEvent=datasheetVals(0,2)
				captureScreen=datasheetVals(0,3)
				AParam1=datasheetVals(0,4)
				AParam2=datasheetVals(0,5)
				AParam3=datasheetVals(0,6)
				AParam4=datasheetVals(0,7)
				AParam5=datasheetVals(0,8)
				AParam6=datasheetVals(0,9)
				AParam7=datasheetVals(0,10)
				AParam8=datasheetVals(0,11)
				AParam9=datasheetVals(0,12)
				AParam10=datasheetVals(0,13)
                Param1=DecParametrization(AParam1,TestCasePath,testdataSource,UIName)
            	Param2=DecParametrization(AParam2,TestCasePath,testdataSource,UIName)
				Param3=DecParametrization(AParam3,TestCasePath,testdataSource,UIName)
				Param4=DecParametrization(AParam4,TestCasePath,testdataSource,UIName)
				Param5=DecParametrization(AParam5,TestCasePath,testdataSource,UIName)
				Param6=DecParametrization(AParam6,TestCasePath,testdataSource,UIName)
				Param7=DecParametrization(AParam7,TestCasePath,testdataSource,UIName)
				Param8=DecParametrization(AParam8,TestCasePath,testdataSource,UIName)
				Param9=DecParametrization(AParam9,TestCasePath,testdataSource,UIName)
				Param10=DecParametrization(AParam10,TestCasePath,testdataSource,UIName)
				If  Not IsArray(Param1) Then
					Parameter1=Param1
				else
					Parameter1=Param1(ite)
					If LCase(Trim(Parameter1))="endofrow" or LCase(Trim(Parameter1))="eof"  Then
						startTest=-1
						Exit For
					End If
				End If
				'TAF 10.1 code modification Start. If a test data in perticular test data column contains a text with '<> ' seperator, means user wants to give all the parameters required for this test step with <> seperator,
				'instead of  each parameter in seperate test data columns. 
	If  Instr(1,Parameter1,"<>") = 0 Then        
				If not  IsArray(Param2) Then
					Parameter2=Param2
				else
					Parameter2=Param2(ite)
				End If
				If not IsArray(Param3) Then
					Parameter3=Param3
				else
					Parameter3=Param3(ite)
				End If
				If not IsArray(Param4) Then
					Parameter4=Param4
				else
					Parameter4=Param4(ite)
				End If
				If not  IsArray(Param5) Then
					Parameter5=Param5
				else
					Parameter5=Param5(ite)
				End If					
					If not  IsArray(Param6) Then
					Parameter6=Param6
				else
					Parameter6=Param6(ite)
				End If					
				If not  IsArray(Param7) Then
					Parameter7=Param7
				else
					Parameter7=Param7(ite)
				End If					
				If not  IsArray(Param8) Then
					Parameter8=Param8
				else
					Parameter8=Param8(ite)
				End If					
				If not  IsArray(Param9) Then
					Parameter9=Param9
				else
					Parameter9=Param9(ite)
				End If					
				If not  IsArray(Param10) Then
					Parameter10=Param10
				else
					Parameter10=Param10(ite)
				End If	
			Else
					
						On Error resume Next
						vvr = Split(Parameter1,"<>")
						Parameter1= vvr(0)
						Parameter2 = vvr(1)
						Parameter3 = vvr(2)
						Parameter4 = vvr(3)
						Parameter5 = vvr(4)
						Parameter6 = vvr(5)
						Parameter7 = vvr(6)
						Parameter8 = vvr(7)
						Parameter9 = vvr(8)
						Parameter10 = vvr(9)
			End If
			'TAF 10.1 code modification End
				If Parameter1 = -1 or  Parameter2 =-1 or Parameter3 = -1 or Parameter4 = -1 or Parameter5 = -1 or Parameter6 = -1 or  Parameter7 =-1 or Parameter8 = -1 or Parameter9 = -1 or Parameter10 = -1 Then
					Call ReportResult  (UIName, functionName,AParam1,"Failed To Perform the opeartion as the Test Data sheet is either empty or not available in Test Data File" , "Failed" ,"")
					WriteToEvent("Failed" & vbtab &"Failed To Perform the opeartion as the Test Data sheet is either empty or not available in Test Data File" )
					If LCase(trim(StopEvent))="showstopper" Then   
						Environment("FlagStepFailureOccured")=True  'santosh 9 sept
						Exit Function   
						else
						Environment("FlagStepFailureOccured")=True
						If  Environment("strTCSeverity")="showstopper" Then
							Exit Function
						End If
					End If   
				Else
												
					Call ExecuteFramework(Environment("AppMapPath"),UIName,functionName,ExpectedResult,stepResultVar,StopEvent,Parameter1,Parameter2,Parameter3,Parameter4,Parameter5,Parameter6,Parameter7,Parameter8,Parameter9,Parameter10,captureScreen)
					If not environment("FlagTCStep") Then   
						If lcase(trim(StopEvent))="showstopper" Then   
							Environment("FlagStepFailureOccured")=True  'santosh 9 sept
							Exit Function   
						else
							Environment("FlagStepFailureOccured")=True
							If  Environment("strTCSeverity")="showstopper" Then
								Exit Function
							End If
						End If   
					End If

												

				End If' For parameter
		End If   'Reuse or normal test case

						If i=rowCount Then
							startTest=-1
						End If
		End If  ' booleanstep
    			
			'' end
		End If  'Skip/Run
		i = Environment.value("tcrow") ' by Srikanth 06/18/2014						
	Next
Next   
End Function


	
	''
	' This function is to Set the value in the Edit Filed.
	' @author DSTWS
	' @param SheetSource String specifying the test script sheet
	' @param sSheetDestination String specifying the data table sheet to copy the test script sheet
	' @param row number specifying from which two to extract from the data table.
	' @Modified By DT77734
	' @Modified on: 16 Jul 2009 

Public Function GetDataSheetValues(SheetSource,SheetDestination,row)
	Dim datasheetVals()
	ReDim datasheetVals(0,13)
	DataTable.GetSheet(SheetDestination).SetCurrentRow(row)
	result=DataTable.GetSheet(SheetDestination).GetParameter("ExpectedResult")
	If Len(Trim(result))=0 Then
		result=""
	End If
	datasheetVals(0,0)=result
	stepResultVar=DataTable.GetSheet(SheetDestination).GetParameter("ReturnValue")
	If Len(Trim(stepResultVar))=0 Then
		stepResultVar=""
	End If
	datasheetVals(0,1)=stepResultVar
	StopEvent=DataTable.GetSheet(SheetDestination).GetParameter("Severity")
	If Len(Trim(StopEvent))=0 Then
		StopEvent=""
	End If
	datasheetVals(0,2)=StopEvent
	captureScreen=DataTable.GetSheet(SheetDestination).GetParameter("CaptureScreen")
	If Len(Trim(captureScreen))=0 Then
		captureScreen="no"
	End If
	datasheetVals(0,3)=captureScreen
	Params1=DataTable.GetSheet(SheetDestination).GetParameter("Data")
	If Len(Trim(Params1))=0 Then
		Params=""
	else
		Params=Split(Params1,"!")
		For i=0 to Ubound(Params)
			If Left(LCase(Trim(Params(i))),5)="svar_" Then
				datasheetVals(0,i+4)=Trim(Params(i))
			else
				datasheetVals(0,i+4)=Trim(Params(i))
			End If
		Next
	End If
	GetDataSheetValues=datasheetVals
End Function


Public Function ConvertVersion(TcVerFrm,TcVerTo,Version,ByRef VF,ByRef VT,ByRef Ver)	'Addde 
		Ver = Version
		VF = TcVerFrm
		VT =  TcVerTo
counter = 0
If (Version <> "")  Then
	For i= 1 to len(Version)
		If  not(isnumeric(Mid(Version,i,1))) Then
				counter = counter+1
				If Mid(Version,i,1) <> "."  and  counter>1 Then
						str =  Mid(Version,i,1)
						Ver = Replace(Ver,str,"." & Asc(str))
				else
						str =  Mid(Version,i,1)
						Ver = Replace(Ver,str,Asc(str))
			End If
		End If
	Next
End If
counter = 0
If (TcVerFrm <> "")  Then
	For i= 1 to len(TcVerFrm)
			If  not(isnumeric(Mid(TcVerFrm,i,1))) Then
				counter = counter+1
				If Mid(TcVerFrm,i,1) <> "."  and  counter>1 Then
						str =  Mid(TcVerFrm,i,1)
						VF = Replace(VF,str,"."& Asc(str))
				else
						str =  Mid(TcVerFrm,i,1)
						VF = Replace(VF,str,Asc(str))
			End If
			eND iF
	Next
End If
counter = 0
If (TcVerTo <> "")  Then
	For i= 1 to len(TcVerTo)
		If  not(isnumeric(Mid(TcVerTo,i,1))) Then
				counter = counter+1
				If Mid(TcVerTo,i,1) <> "."  and  counter>1 Then
						str =  Mid(TcVerTo,i,1)
						VT = Replace(VT,str,"." & Asc(str))
				else
						str =  Mid(TcVerTo,i,1)
						VT = Replace(VT,str,Asc(str))
			End If
		eND iF
	Next
End If

End Function


Function PerformExecuteTest
   'Set oFileExist=CreateObject("TAFCore.CoreEngine")
        If (VerifyFileExists(Environment("strTestDataFilePath"))) Then
            blnDataSheetExists = VerifySheetExists(Environment("TestDataPath"), Environment("TestDataSheetName"))
            If blnDataSheetExists Then
                If Environment("ExecutionStarted") Then
                    Call WritetoResult_DataPath()
                    Call Footer(Environment("EventLogPath"), LCase(Trim(Environment("EventLog"))))
                    Call Footer(Environment("ScriptLogPath"), LCase(Trim(Environment("ScriptLog"))))
                End If
            End If
'            If Not (Environment("DataSheetExist")) Then
'
'                If Environment("RunFromQC") Then
'                    'Reporter.ReportEvent(micFail, "Fetching the Test Data from Data Sheet : " & Environment("TestDataSheetName") & " from TD file : " & Environment("QCTestDataFilePath"), "Test Data Sheet :" & Environment("TestDataSheetName") & " doesn't exists in the TD file : " & Environment("QCTestDataFilePath") & "for the Test Case : " & Environment("TestCaseName1"))
'                    WriteToEvent("Failed" & vbTab & "Test Data Sheet :" & Environment("TestDataSheetName") & " doesn't exists in the TD file : " & Environment("QCTestDataFilePath") & "for the Test Case : " & Environment("TestCaseName1"))
'
'                Else
'                    'Reporter.ReportEvent(micFail, "Fetching the Test Data from Data Sheet : " & Environment("TestDataSheetName") & " from TD file : " & Environment("strTestDataFilePath"), "Test Data Sheet :" & Environment("TestDataSheetName") & " doesn't exists in the TD file : " & Environment("strTestDataFilePath") & "for the Test Case : " & Environment("TestCaseName1"))
'                    WriteToEvent("Failed" & vbTab & "Test Data Sheet :" & Environment("TestDataSheetName") & " doesn't exists in the TD file : " & Environment("strTestDataFilePath") & "for the Test Case : " & Environment("TestCaseName1"))
'                End If
'            End If
        End If
End Function


'*********************************************************************************************************************************************************************************
' Function Name:FormatData
' Description: This function formats the data to taf format
' param Param Parameter value for the test pobect functions
' returns : Formatted parameter value
' Author : DSTWS TA2000 Automation Team
' Creation Date : April, 2011
' Reviewed By : DSTWS TA2000 Automation Team
'
' Modified By :
' Modified Date : 
'*********************************************************************************************************************************************************************************
Function FormatData(Param)
Param=LCase(Param)
If Param="" Then
FormatData=""
Exit Function
Else
If Mid (Param,1,8)="property" Or Mid (Param,1,3)="row" Or Mid (Param,1,3)="col" Or Mid (Param,1,5)="value" Or Mid (Param,1,8)="validate" Or Mid (Param,1,4)="save" Or Mid (Param,1,6)="select" Or Mid (Param,1,9)="selection" Or Mid (Param,1,4)="type" Or Mid(Param,1,7)="timeout" Then
If Left(LCase(Trim(Param)),5)="save=" Or Left(LCase(Trim(Param)),6)="save ="Then
arrVar=Split(LCase(Trim(Param)),"=")
stepResultVar=arrVar(1)

Else
arrParam=Split(Trim(Param),"=")
Param=arrParam(1)
End If
End If
End If
If Mid(Param,1,1)="#" Then
Param=Mid(Param,2)
If oActivityResult.Exists(Param) Then 
Param=oActivityResult.Item(Param) 
End If
End If
FormatData=Param
End Function



Public Function Evaluate(ByVal val)


						If CInt(val) >= 1 Then
							Evaluate=True
						Else
						   Evaluate=False
						End If
'	Select Case vartype(val) 
'	
'			Case 2,3,4,5:
'            
'						If val >= 1 Then
'							Evaluate=True
'						Else
'						   Evaluate=False
'						End If
'        
'			Case 8:
'
'						If UCase(val)="TRUE" Then
'							Evaluate=True
'						Else
'						    Evaluate=False
'						End If
'    
'			Case 0:
'
'						Evaluate=False
'			Case 1:
'
'						Evaluate=False
'
'			Case 11:
'
'						If val=True Then
'							Evaluate=True
'						Else
'						    Evaluate=False
'						End If
'
'	End Select

End Function

Public Function SaveScreenshotsinLocalDrive()
   	ScreenShotPath = Environment("SSPath")&"\Screen"&Environment("StepNumber")&".png"
	Desktop.CaptureBitmap ScreenShotPath ,true
End Function

Public Function CreateLocalDriveScreenshotFolder()
   	Set fsoo = CreateObject("Scripting.FileSystemObject")
	n = Replace(Now(),"/","_")
	m = Replace(n,":","_")
	If Environment("RowKey") <> "0" Then
		Environment("SSPath") = Environment("LocalDriveScreenshotsPath")&"\"&Environment("DriverName")&"_"&Environment("RowKey")&"_"&m
	Else
		Environment("SSPath") = Environment("LocalDriveScreenshotsPath")&"\"&Environment("DriverName")&"_"&m
	End If
	fsoo.CreateFolder(Environment("SSPath"))
	Set fsoo = Nothing
End Function
