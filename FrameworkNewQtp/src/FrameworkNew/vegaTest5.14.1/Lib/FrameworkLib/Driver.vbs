
''
' @# author DSTWS
' @# Version 7.0.7 
' @# Version 7.0.8 Revision History :  Excel Revised code incorporated with "objResWorkBook" , Added "KillProcess" and commented for production purpose.
' @# Version 7.0.9 Modified by DT77471 on 19Mar 10 'Handle object synchronization timeout at global level @CR 61@, line 99
' @# Version 7.0.9 Modified by Rajesh on 29 Mar 10 'To verify the suite from Resources
' @# Version 7.0.9 Modified DownloadAllAttachments, DownloadAttachments by Rajesh on 29 Mar 10 'To download the attachments from Resources tab
' @# Version 7.0.9A Modified DownloadAllAttachments, DownloadAttachments ,TestRunStatus
' @# Version 7.0.10 Modified 
	''
	'  This function is to verify the existance of a file in the system disk.
	' @author DSTWS
	' @param strFilePath String specifying the path of the verifying file.
	'@ Changed by DT77734  on 02 Sept 09
	' @ Modified by DT77215  on 19 Nov 09 ' Handled QC Test attachments for current run
	
Function CallDriver(strTestScriptName,blnIsScenario)
	Environment("DriverName")=strTestScriptName
	If Environment("TestControllerPath")="" Then
		reporter.ReportEvent micFail,"Verifying for Test Plan File  ", "Test Plan Path Does not Exists(Please verify in XML  file) Terminating the Test Execution"
		CallDriver=-1
		Exit Function
	End If

	If strTestScriptName="" Then
		reporter.ReportEvent micFail,"Verifying for Test Case", "Test Script Name is not a valid one"
		CallDriver=-1
		Exit Function
	End If
	If blnIsScenario="" Then
		reporter.ReportEvent micFail,"Verifying the Script is a Test Case or Scenario", "Please provide a valid True or False for blnIsScenario"
		Exit Function
	End If

	Environment("vegaTestVersion") = "vegaTest 4.13.1 DLL"
	'TAF 10.1 new code Start
	Environment("BlnCodeGeneration")=0      '10.1
	If VerifyEnvVariable("CodeGeneration") Then
		If UCase(Trim(Environment("CodeGeneration")))="YES"  Then
			Environment("BlnCodeGeneration")=1
		End If
	End If
   'TAF 10.1 new code End

	If VerifyEnvVariable("TestCaseSummaryResult") Then

			If Environment("RunFromQC")    Then 
				 Environment("SumResFilName")   = "vegaTestSUmmaryReport_" &  Environment("LocalHostName")   & "_" & Environment("UserName") & ".xls"
			End If

		 Environment("IsSummaryResult")  = Environment("TestCaseSummaryResult") 
		If  Lcase(TRim(Environment("IsSummaryResult")))   = "yes"Then
				If Not(VerifyEnvVariable("SumResFilName")) Then
						Environment("IsSummaryResult")   = "no"	
				End If
		End If

	Else
			Environment("IsSummaryResult")   = "no"
	End If


	If Environment("RunFromQC")  Then

		CallDriver = RunFromQC(strTestScriptName,blnIsScenario)
	else
		If VerifyFileExists(Environment("TestControllerPath"))=0 then
			reporter.ReportEvent micFail,"Verifying for Test Plan File", "Test Plan File doesn't exists, Terminating the Test Execution"
			CallDriver=-1
			Exit Function
       	End If
		CallDriver = RunFromLS(strTestScriptName,blnIsScenario)
	End If

  

End Function

Set fso=CreateObject("scripting.filesystemobject")

Function CreateQCDirectoryStructure

	Environment("TempTestCases") = Environment("WorkingDirectory") & "\TempAutomation\KeywordDriven\TestCases"
	Environment("TempTestData") = Environment("WorkingDirectory") & "\TempAutomation\KeywordDriven\TestData"
	Environment("TempReusableTestCases") = Environment("WorkingDirectory") & "\TempAutomation\KeywordDriven\ReusableTestCases"
	Environment("TempAppMap") = Environment("WorkingDirectory") & "\TempAutomation\KeywordDriven\AppMap"
	Environment("TempTestPlanFolder") = Environment("WorkingDirectory") & "\TempAutomation\KeywordDriven\TestPlan"
	Environment("TempResultPath") = Environment("WorkingDirectory") & "\TempAutomation\KeywordDriven\Results"
	Environment("TempExceptionHandling") = Environment("WorkingDirectory") & "\TempAutomation\KeywordDriven\ExceptionHandling"
	Environment("TempLib") = Environment("WorkingDirectory") & "\TempAutomation\KeywordDriven\Lib"  ' Added this to download to this path

	If	  fso.FolderExists(Environment("WorkingDirectory") & "\TempAutomation")Then
		fso.DeleteFolder(Environment("WorkingDirectory") & "\TempAutomation")
	End If
	
	strPath = Replace(Environment("WorkingDirectory") ,"/","\")
	arr = Split(Environment("WorkingDirectory") ,"\")
	If ubound(arr) >  0 Then
		strkey =  arr(0) 
		For i = lbound(arr) to ubound(arr)-1
			On error resume next
			fso.CreateFolder strkey & "\" & arr(i+1)
			strkey = strkey &"\" & arr(i+1)
			On error goto 0
		Next
	End If
	'msgbox Environment("WorkingDirectory")
	fso.CreateFolder(Environment("WorkingDirectory") & "\TempAutomation")
	fso.CreateFolder(Environment("WorkingDirectory") & "\TempAutomation\KeywordDriven")
	fso.CreateFolder(Environment("TempTestCases") )
	fso.CreateFolder(Environment("TempTestData"))
	fso.CreateFolder(Environment("TempReusableTestCases") )
	fso.CreateFolder(Environment("TempAppMap"))
	fso.CreateFolder(Environment("TempTestPlanFolder") )
	fso.CreateFolder(Environment("TempResultPath") )
	fso.CreateFolder(Environment("TempExceptionHandling") )
	fso.CreateFolder(Environment("TempLib"))
      
End Function



Function RunFromQC(strTestScriptName,blnIsScenario)

	Dim App 'As Application
	
	Set App = CreateObject("QuickTest.Application")
'	App.Launch
'	App.Visible = True
	
	On error resume Next					' to implement the object synchronization time out at global level
	Set theTSTest = QCUtil.CurrentTestSetTest
	If Not theTSTest Is Nothing Then	
		actTimeout=App.Test.Settings.Run.ObjectSyncTimeOut
		'If  VerifyEnvVariable("bolobjSyncTimeOut") Then
		If  VerifyEnvVariable("objSyncTimeOut") Then
				SyncTimeOut=Environment("objSyncTimeOut")
			If CLng(actTimeout)<>SyncTimeOut Then
				App.Test.Settings.Run.ObjectSyncTimeOut =SyncTimeOut
			End If					
		End If	
	End If
	On error goto 0 
	
	'Environment("CurTestSetName") =  QCUtil.CurrentTestSet.Name
	'Environment("CurTestName")   = Right(QCUtil.CurrentTestSetTest.Name,len(QCUtil.CurrentTestSetTest.Name)-3)
		' 7.0.8B		'Rajesh
	arrTAFSuite = Split(Environment("vegaTestSuite"),"\")
	If LCase(Trim(arrTAFSuite(0))) = "subject" Then
		Environment("IsSuiteFromTestPlan") = True
	ElseIf LCase(Trim(arrTAFSuite(0))) = "resources" Then
		Environment("IsSuiteFromTestPlan") = False
	End IF
	
	' 7.0.8B

	call CreateQCDirectoryStructure  ' First Creat Temp Directory
	
	If LCase(Trim(Environment("UseQCCustomFields")))= "yes"  Then
		Call GenerateEnvVaribaleFromQCTestSet
	End If
	
	If VerifyEnvVariable("TestRunStatus") Then
		If LCase(Trim(Environment("TestRunStatus")))= "don'trun" Then
			RunFromQC = -1
			Exit Function	
		End If
	End If

	
	strTestScriptName=LCase(Trim(strTestScriptName))  
	DataTable.AddSheet("TestPlan")
	arrTemp = Split(Environment("TestControllerPath"),"\")
	strTestPlanFileName = arrTemp(ubound(arrTemp))
	arrTemp = Split(Environment("TestControllerPath"),"\" & strTestPlanFileName)
	strTestPlanFileFolder = arrTemp(lbound(arrTemp))
	Environment("QCWorkingDirectory")=Environment("vegaTestSuite")
	'msgbox "loading - Test Plan"
	DownloadAttachment  strTestPlanFileName,strTestPlanFileFolder,Environment("TempTestPlanFolder") ,"TRUE"
	'msgbox "Completed - loading - Test Plan"

		strLogoFileName = "tag1.exe"
		strLibFolderPath=Environment("vegaTestSuite")&"\Lib"
		DownloadAttachment strLogoFileName,strLibFolderPath,Environment("TempLib"),"TRUE"


    Environment("TestPlanPath") = Environment("TempTestPlanFolder") &   "\" & strTestPlanFileName 
        If Environment("UseMsAccessDB") Then
		LoadDataTableFromDB Environment("TestPlanPath"),strTestScriptName,blnIsScenario
	Else
		DataTable.ImportSheet Environment("TestPlanPath"),"TestPlan","TestPlan"
	End If
'	DataTable.ImportSheet Environment("TestPlanPath"),"TestPlan","TestPlan"
	Environment("ResultPath")= Environment("TempResultPath")

	Set TSTestFact   = QCUTil.CurrentTestSet.TSTestFactory
	Set TestSetTestsList = TSTestFact.NewList("") 
	CurTestName =  TestSetTestsList.Item(1).Name
	IntTestsCount =  TestSetTestsList.Count

	If  CurTestName = QCUtil.CurrentTestSetTest.Name Then
		Environment("FirstTest") = True
		Else
		Environment("FirstTest") = False
	End If

	If   TestSetTestsList.Item(IntTestsCount).Name = QCUtil.CurrentTestSetTest.Name Then
		Environment("LastTest") = True
		Else
		Environment("LastTest") = False
	End If

	MakePath()

'	Environment("ExecutionStarted")=False
'	Environment("ReusableTestCaseFolderPath")= Environment("TempReusableTestCases")
	Environment("ExecutionStarted")=False


	strNewReUSeFilesPath = Environment("WorkingDirectory") & "\TempReusableTestCases"
	Environment("ReusableTestCaseFolderPath")= strNewReUSeFilesPath
	
	If Environment("FirstTest") = True then
		If fso.FolderExists(strNewReUSeFilesPath) Then 
			Fso.DeleteFolder(strNewReUSeFilesPath)
        End If
		Fso.CreateFolder strNewReUSeFilesPath
	
		DownloadAllAttachments Environment("vegaTestSuite") & "\" &   "ReusableTestCases",Environment("ReusableTestCaseFolderPath")
		
	End If
	
	If Not(fso.FolderExists(strNewReUSeFilesPath)) Then 
			Fso.CreateFolder strNewReUSeFilesPath
	End If
	
	
	
   	arrTemp = Split(Environment("TestControllerPath"),"Controller")
	strReusablesFolder = arrTemp(lbound(arrTemp))
	'DownloadAllAttachments Environment("TAFSuite") & "\" &   "ReusableTestCases",Environment("TempReusableTestCases")
	DownloadAllAttachments  Environment("vegaTestSuite") & "\" &   "ExceptionHandling",Environment("TempExceptionHandling")
    CurrentScenario=""
	Environment("blnFlagNewScenario")=false
	Environment("TSorTC")=Null
	Environment("strNewScenario")=Null
	Environment("IncFlag")=False
	Environment("CasesCount")=0
'	Environment("IncValHolder")=1  ' sets the initial value for report results
	Environment("ScenCounter")=0
	Environment("TestCaseCounter")=0
	Environment("EndOfScenarios")=0
	Environment("ScenerioEnd")=False 
	Environment("ScenerioEncountered")=False
	Environment("FlagStepFailureOccured")=False
	Environment("strTCSeverity")=Null
	FlagStepFailureOccured=False
	Environment("CountFlag")=0
	ScenarioCounter=0
	TestStepCounter=0
	Environment("blnNoTDExists")=Null
	FlagIsShowStopper=False
	Environment("TestStepCounter")=TestStepCounter
	Environment("ScenerioStepCounter")=0 '  scenario  step counter
	Environment("FirstOnly")=False 
	Environment("LastOnly")=False 
	Environment("tagCount")=False
	'TAF 10.1 new code Start.
	 Environment("NoOfDatarows")=0
	  intFirstTime=0
	'TAF 10.1 new code End

	
	On Error resume next
	If err.number = 0  and Lcase(trim(Environment("BaseState")))  <> "n/a" Then
		ExecutePreReqScript Lcase(trim(Environment("BaseState"))) 		
	End If
	On error goto 0
	
	For intScenarioCounter=2 to DataTable.GetSheet("TestPlan").GetRowCount							'Loop which rolls on the test plan for test scenario or test case
		If intScenarioCounter=2 Then																													'Code to store the Test Plan level master default files like AppMap, Test Case and Test Data			
			DataTable.GetSheet("TestPlan").setCurrentRow(intScenarioCounter-1)

			TPMasterAppMap = DataTable.GetSheet("TestPlan").getParameter("AppMapPath")
			TPMasterAppMap_Relative=Trim(TPMasterAppMap)
			TPMasterTestCase =  DataTable.GetSheet("TestPlan").getParameter("TestCaseFilePath")
			TPMasterTestData =   DataTable.GetSheet("TestPlan").getParameter("TestDataPath") 
			 TPMasterTestData_Relative=Trim(TPMasterTestData)

            strTPMasterAppMap= Environment("vegaTestSuite")  & DataTable.GetSheet("TestPlan").getParameter("AppMapPath")
			strTPMasterTestCase= Environment("vegaTestSuite")  &  DataTable.GetSheet("TestPlan").getParameter("TestCaseFilePath")
			strTPMasterTestData1 = Environment("vegaTestSuite")  &  DataTable.GetSheet("TestPlan").getParameter("TestDataPath")  'raj
			'strTPMasterTestData = Replace(lcase(Trim(strTPMasterTestData1)),"%client%",Lcase(Trim(Environment("Client"))))  'raj
			strTPMasterTestData = GetDynamicParameter(strTPMasterTestData1)
			
		End If

		'TAF 10.1 new code Start.  
		If  intFirstTime=0 Then
				'Count the required line directly in test plan excel, rathen than moving line by line for the required row.
				intScenarioCounter=getContNumber(Environment("TestPlanPath"), strTestScriptName, blnIsScenario)
				 intFirstTime=1         'This is to avoid for iteration to come in to this IF condition. If it come, every time intScenarioCounter will have the same number, leads to infinite loop
				If not isNumeric(intScenarioCounter) Then
				intScenarioCounter=DataTable.GetSheet("TestPlan").GetRowCount
				End If
        End If
		'TAF 10.1 new code End
		

		DataTable.GetSheet("TestPlan").setCurrentRow(intScenarioCounter)
		blnFlagSkipTestCase=false
		blnFlagNoTestCase=false
		blnNoTDExists=false
		If  Lcase(Trim(DataTable.GetSheet("TestPlan").getParameter("Scenario_Keyword")))="scenario" And LCase(Trim(DataTable.GetSheet("TestPlan").getParameter("TestCaseName")))=strTestScriptName And blnIsScenario then
				'TAF 10.1 new code Start
				If Environment("BlnCodeGeneration")=1 Then   '10.1
					Call CreateFileForCode()			'10.1
					Call UpdateScript("'******************"&"TestScenario:  "&strTestScriptName&"   start"&"**************************************")
					Call UpdateScript("")
					Call UpdateScript("On Error Resume Next")
					Call UpdateScript("Call MakePathNonTAF()")
				End If
				'TAF 10.1 new code end
				

			Err.Clear
			On Error Resume Next
			Environment("Description1")= DataTable.GetSheet("TestPlan").getParameter("Description")  ' ARgus 
			If err.number > 0  Then
				strDescription = ""
				Else
				strDescription =  Environment("Description1")
				
			End If
			On Error goto 0

			If LCase(Trim(Environment("UseQCCustomFields")))= "yes"  Then
	      			On Error Resume next
		       		If Environment("DataRow_KeyWord")= ""  Then
			        		Environment("ScenarioDataKeys") = DataTable.GetSheet("TestPlan").getParameter("DataRow_Keyword")
			   				Environment("ScenarioDataKeys") = GetDynamicParameter(DataTable.Value("DataRow_Keyword","TestPlan"))        ' Q0180 
			   			
			   			Else
			   			Environment("ScenarioDataKeys")  = Environment("DataRow_KeyWord")
			   		End If
			   		If Err.Number = -2147220983 Then
		    			Environment("ScenarioDataKeys") = DataTable.GetSheet("TestPlan").getParameter("DataRow_Keyword")
			   			Environment("ScenarioDataKeys") = GetDynamicParameter(DataTable.Value("DataRow_Keyword","TestPlan"))        ' Q0180 
			   		End If
					On error goto 0	
			
			   	Else
			   		Environment("ScenarioDataKeys") = DataTable.GetSheet("TestPlan").getParameter("DataRow_Keyword")
			   		Environment("ScenarioDataKeys") = GetDynamicParameter(DataTable.Value("DataRow_Keyword","TestPlan"))        ' Q0180 
			   							
				End If
						
				ScenarioCounter=ScenarioCounter+1
				Environment("ScenarioCounter")=ScenarioCounter
				strNewScenario=DataTable.GetSheet("TestPlan").getParameter("TestCaseName")
				If not(strNewScenario=CurrentScenario) Then
					Environment("strNewScenario")=strNewScenario
					Environment("blnFlagNewScenario")=true 
				End If

				ScenarioMasterAppMap = DataTable.GetSheet("TestPlan").getParameter("AppMapPath")
				ScenarioMasterAppMap_Relative=Trim(ScenarioMasterAppMap)
				ScenarioMasterTC= DataTable.GetSheet("TestPlan").getParameter("TestCaseFilePath")
				ScenarioMasterTD =   DataTable.GetSheet("TestPlan").getParameter("TestDataPath") 
				ScenarioMasterTD_Relative=Trim(ScenarioMasterTD)

                strScenarioMasterAppMap= Environment("vegaTestSuite")  &  DataTable.GetSheet("TestPlan").getParameter("AppMapPath")
				strScenarioMasterTC= Environment("vegaTestSuite")  &  DataTable.GetSheet("TestPlan").getParameter("TestCaseFilePath")
				strScenarioMasterTD1 = Environment("vegaTestSuite")  &  DataTable.GetSheet("TestPlan").getParameter("TestDataPath")  'raj
			    'strScenarioMasterTD = Replace(lcase(Trim(strScenarioMasterTD1)),"%client%",Lcase(Trim(Environment("Client"))))   'raj
			    strScenarioMasterTD = GetDynamicParameter(strScenarioMasterTD1)

				arrScenarioKeywods = 	Split(Environment("ScenarioDataKeys") ,";")        '  ' DSTGS  ' DSTGS 
				scenariorow = 	intScenarioCounter  ' DSTGS  ' DSTGS 
				
				
'				If ubound(arrScenarioKeywods) = -1 Then
'				iterationCnt =  1
'				else
'				iterationCnt = 	ubound(arrScenarioKeywods)+1
'				End If

'''''''''''''''''''' Viswanadh '''''''''''''''''''''''''''''
				If ubound(arrScenarioKeywods) = -1 Then
						iterationCnt =  1
					else
						If LCase(Trim(arrScenarioKeywods(0))) <> "all" Then
							iterationCnt = 	ubound(arrScenarioKeywods)+1
						Else
							arrScenarioKeywods = GetDataRows(strScenarioMasterTD1,Datatable.GetSheet("TestPlan").GetParameter("TestDataSheetName"))
							iterationCnt = 	ubound(arrScenarioKeywods)+1
						End If
					End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''		


				For scit  = 1 to iterationCnt ' DSTGS  ' DSTGS 
					Environment("CurrentIteration") = scit  ' Argus 7.0.10
					If  ubound(arrScenarioKeywods) = -1 Then
						Environment("ScenarioDataKeys") =""
						else
						Environment("ScenarioDataKeys") =  arrScenarioKeywods(scit-1) ' DSTGS  ' DSTGS 
					End If
					
					'''' Scenario keys
						Environment("SceIte") = scit
					If scit=1 Then
							Environment("FirstOnly") = True
							else
							Environment("FirstOnly") = False
					End If

					if scit=iterationCnt Then
							Environment("LastOnly") = True
							Else
							Environment("LastOnly") = False
					End If
					
					'''' Scenario keys
				intScenarioCounter = scenariorow ' DSTGS  ' DSTGS 
				DataTable.GetSheet("TestPlan").setCurrentRow(intScenarioCounter) ' DSTGS  ' DSTGS 

				While (Lcase(Trim(DataTable.GetSheet("TestPlan").getParameter("Scenario_Keyword")))<>"endofscenario")
					blnFlagSkipTestCase=false
					blnFlagNoTestCase=false
					blnNoTDExists=false
					intScenarioCounter=intScenarioCounter+1
					DataTable.GetSheet("TestPlan").setCurrentRow(intScenarioCounter)
					If  Lcase(Trim( DataTable.GetSheet("TestPlan").getParameter("TestCaseName")))="endofrow" then
						Exit For
					End If

					If FlagIsShowStopper=False  Then    'Santosh 10 Sept
					
						AppMapPath =  DataTable.GetSheet("TestPlan").getParameter("AppMapPath")
						AppMapPath_Relative=Trim(AppMapPath)
     					strAppMapPath= Environment("vegaTestSuite")  &  DataTable.GetSheet("TestPlan").getParameter("AppMapPath")
						Environment("TestCaseName1")=DataTable.GetSheet("TestPlan").GetParameter("TestCaseName")
						Environment("strTCSeverity")=Lcase(Trim(DataTable.GetSheet("TestPlan").GetParameter("Severity_KeyWord")))     ''Santosh
						Call Header(Environment("EventLogPath"), LCase(Trim(Environment("EventLog")))) 
						Call Header(Environment("ScriptLogPath"), LCase(Trim(Environment("ScriptLog")))) 

						'TAF 10.1 new code Start
						If Environment("BlnCodeGeneration")=1  Then            '10.1
							 Call UpdateScript("'******************"&"TestCase:  "&Environment("TestCaseName1")&"   start"&"**************************")       '10.1
							Call UpdateScript("")
						End If
						'TAF 10.1 new code End

						Environment("WithoutTestData")=False     'Make it false before start of every case--TRao--TAF10

						'Modified by Trao to make appmap not mandatory------------TAF10
					If  AppMapPath_Relative="" and ScenarioMasterAppMap_Relative="" and TPMasterAppMap_Relative="" Then
						    Environment("AppMapCount")=0
							Environment("AppMapPath")=""
					Else
				
					' Modified by Srikanth ----- Verify AppMap existence and decide the AppMap and download to TempAutomation folder ----------------------------Modified by Srikanth------TAF10
						If  VerifyAndDownloadAppMap(AppMapPath)="0"  then
							If  VerifyAndDownloadAppMap(ScenarioMasterAppMap)= "0" Then
								If  VerifyAndDownloadAppMap(TPMasterAppMap) = "0" then
									blnFlagSkipTestCase=true
									If Environment("strTCSeverity")="showstopper" Then
										FlagIsShowStopper=True
									End If
								else
									Reporter.ReportEvent micWarning,"Verifying APPMAP for "& Environment("TestCaseName1")&" test case","Copied Default Test Plan level APPMAP i.e. "& strTPMasterAppMap   
									WriteToEvent("Warning" & vbtab & "Copied Default Test Plan level APPMAP i.e. "& strTPMasterAppMap &" for Test Case " & Environment("TestCaseName1") )
									strAppMapPath=strTPMasterAppMap
								end if
							else
								Reporter.ReportEvent micWarning,"Verifying APPMAP for "& Environment("TestCaseName1")&" test case","Copied Default Test Scenario level APPMAPi.e. "& strScenarioMasterAppMap 
								WriteToEvent("Warning" & vbtab & "Copied Default Test Scenario level APPMAP i.e. "& strScenarioMasterAppMap &" for Test Case " & Environment("TestCaseName1"))
								strAppMapPath=strScenarioMasterAppMap
							End If
						End If
						' Modification Done---------------------------TAF10
					End If
'					'End of Trao Modification----------------------TAF10

'	                 arrTemp = Split(strAppMapPath,"\")
'	                 strAppMapFileName = arrTemp(ubound(arrTemp))
'	                 arrTemp = Split(strAppMapPath,"\" & strAppMapFileName)
'		strAppMapFileFolder = arrTemp(lbound(arrTemp))
'		'msgbox " loading - App Map"
'		DownloadAttachment  strAppMapFileName,strAppMapFileFolder,Environment("TempAppMap"),"TRUE"
'		'msgbox "Completed - loading - App Map"
'   	Environment("AppMapPath")=  Environment("TempAppMap")&   "\" & strAppMapFileName 
'     	Environment("QCAppMapPath")=  strAppMapPath

					'' Verifying Test Case availability 
						TCFilePath =  DataTable.GetSheet("TestPlan").getParameter("TestCaseFilePath")
						strTCFilePath= Environment("vegaTestSuite")  &  DataTable.GetSheet("TestPlan").getParameter("TestCaseFilePath")
						If not blnFlagSkipTestCase Then
							If TCFilePath= "" Then
								If ScenarioMasterTC = "" Then
									If TPMasterTestCase = "" Then
										blnFlagNoTestCase=true
                                        If Environment("strTCSeverity")="showstopper" Then
											FlagIsShowStopper=True
										End If
									else
										Reporter.ReportEvent micWarning,"Verifying TestCase file for "& Environment("TestCaseName1")&" test case","Copied Default Test Plan level Test Case i.e. " & strTPMasterTestCase       
										WriteToEvent("Warning" & vbtab & "Copied Default Test Plan level Test Case i.e. "& strTPMasterTestCase &" for Test Case " & Environment("TestCaseName1") ) 
										strTCFilePath=strTPMasterTestCase
									End If
								else
									Reporter.ReportEvent micWarning,"Verifying TestCase file for "&  Environment("TestCaseName1") &" test case","Copied Default Test Scenario level Test Case i.e. " & strScenarioMasterTC
									WriteToEvent("Warning" & vbtab & "Copied Default Test Scenario level Test Case i.e. "& strScenarioMasterTC &" for Test Case " & Environment("TestCaseName1") ) 
									strTCFilePath=strScenarioMasterTC
								End If
							else
								strTCFilePath= Environment("vegaTestSuite")  &  DataTable.GetSheet("TestPlan").getParameter("TestCaseFilePath")
							End if
						End if
	                 arrTemp = Split(strTCFilePath,"\")
	                 strTCFileName = arrTemp(ubound(arrTemp))
	                 arrTemp = Split(strTCFilePath,"\" & strTCFileName)
		strTCFileFolder = arrTemp(lbound(arrTemp))
		'msgbox "loading - Test Case"
		DownloadAttachment  strTCFileName,strTCFileFolder,Environment("TempTestCases") ,"TRUE"
		'msgbox "Completed - loading - Test Case"
		Environment("TestCasePath")= Environment("TempTestCases")  &    "\" & strTCFileName 

         Environment("QCTestCasePath")= strTCFilePath


						TestDataFilePath = DataTable.GetSheet("TestPlan").GetParameter("TestDataPath") 
						TestDataFilePath_Relative=Trim(TestDataFilePath)
                        strTestDataFilePath1= Environment("vegaTestSuite")  &  DataTable.GetSheet("TestPlan").GetParameter("TestDataPath")   'Raj
						'msgbox strTestDataFilePath1
						'strTestDataFilePath = Replace(lcase(Trim(strTestDataFilePath1)),"%client%",Lcase(Trim(Environment("Client"))))  'Raj
						strTestDataFilePath = GetDynamicParameter(strTestDataFilePath1)
						'msgbox strTestDataFilePath
'						Environment("strTestDataFilePath")=DataTable.GetSheet("TestPlan").GetParameter("TestDataPath")

					'Modified by Trao to make test data path not mandatory-----TAF10
				If  TestDataFilePath_Relative="" and ScenarioMasterTD_Relative="" and TPMasterTestData_Relative="" Then
					Environment("WithoutTestData")=True
					Environment("strTestDataFilePath")=""
				Else
						
						If blnFlagSkipTestCase=false and blnFlagNoTestCase=false Then
							If TestDataFilePath="" then
								If ScenarioMasterTD="" Then
									If TPMasterTestData="" Then
										blnNoTDExists=true
									Else
										Reporter.ReportEvent micWarning,"Verifying Test Data  file for "& DataTable.GetSheet("TestPlan").getParameter("TestCaseName")&"  Test Case","Copied Default Test Plan level Test Data file  i.e. " & strTPMasterTestData      
										WriteToEvent("Warning" & vbtab & "Copied Default Test Plan level Test Data file i.e. "& strTPMasterTestData &" for Test Case " & Environment("TestCaseName1") ) 
										strTestDataFilePath= strTPMasterTestData
									End If
								Else
									Reporter.ReportEvent micWarning,"Verifying Test Data  file for "& DataTable.GetSheet("TestPlan").getParameter("TestCaseName")&"  Test Case","Copied Default Test Scenario level Test Data file i.e. " & strScenarioMasterTD 
									WriteToEvent("Warning" & vbtab & "Copied Default Test Scenario level Test Data file i.e. "& strScenarioMasterTD &" for Test Case " & Environment("TestCaseName1") ) '
									strTestDataFilePath = 	strScenarioMasterTD
								End If
							End if	
						End if
					arrTemp = Split(strTestDataFilePath,"\")
	                strTDFileName = arrTemp(ubound(arrTemp))
	               			arrTemp = Split(strTestDataFilePath,"\" & strTDFileName)
					strTDFileFolder = arrTemp(lbound(arrTemp))
					'msgbox "Loading - Test Data"
					'msgbox strTDFileName
					'msgbox strTDFileFolder
					DownloadAttachment  strTDFileName,strTDFileFolder,Environment("TempTestData") ,"TRUE"
				'	msgbox "Completed - loading - Test Data"
					Environment("strTestDataFilePath")= Environment("TempTestData") &    "\" & strTDFileName 
					Environment("QCTestDataFilePath") =  strTestDataFilePath
				End If
					'End of Trao modification---------TAF10

					If  blnFlagSkipTestCase then
							reporter.ReportEvent micFail,"Verifying the  AppMap File"& Environment("QCAppMapPath")  &" for test case "& Environment("TestCaseName1"), "AppMap File doesn't exists"		
							WriteToEvent("Fail" & vbtab &" AppMap File "&Environment("QCAppMapPath")&" does not exist")
						else
							If blnFlagNoTestCase=true Then
								reporter.ReportEvent micFail, "Verifying the Test Case " & Environment("QCTestCasePath") , strTCFileFolder & "\" & strTCFileName  & " file does not exist"
								WriteToEvent("Fail" & vbtab & Environment("QCTestCasePath") & " file does not exist")
							else
								'If not(FlagIsShowStopper)  Then
									Environment("TSorTC")=True
									If  verifysheetexists(environment("TestCasePath"), Environment("TestCaseName1"))  Then	  'Santosh 10 Sept
										'Call ExecuteTestCase()	
If Environment("BPT") Then
	Datatable.AddSheet("Global")
	DataTable.GetSheet("Global").AddParameter "TestCaseSheetName",""
	DataTable.GetSheet("Global").AddParameter "TestCaseFilePath",""
	DataTable.GetSheet("Global").AddParameter "TestDataDBPath",""
End If			
'	DataTable.GetSheet("Global").AddParameter "Description",""  ' Argus 

																													  'Santosh 10 Sept					
																	'  Changes for DLL
DataTable("TestCaseSheetName","Global")=DataTable.GetSheet("TestPlan").GetParameter("TestCaseName")      
Environment("TestCaseSheetNameToResults") = DataTable("TestCaseSheetName","Global")  ' Argus 
DataTable("TestCaseFilePath","Global")=DataTable.GetSheet("TestPlan").GetParameter("TestCaseFilePath")		
Datatable("TestCaseFilePath","Global")=Datatable.GetSheet("TestPlan").GetParameter("TestCaseFilePath")
Datatable("TestDataDBPath", "Global") = Environment("strTestDataFilePath")
Environment("TestDataPath") = Environment("strTestDataFilePath")
TempTestDataSheetName=Datatable.GetSheet("TestPlan").GetParameter("TestDataSheetName")
Environment("TestDataSheetName") = TempTestDataSheetName

'set below line putput to environment variable
TempDrefDataRow_Keyword=Datatable.GetSheet("TestPlan").GetParameter("DataRow_Keyword")
Environment("DrefDataRow_Keyword")=TempDrefDataRow_Keyword
TempDrefDataDriven_KeyWord=Datatable.GetSheet("TestPlan").GetParameter("DataDriven_KeyWord")
Environment("DrefDataDriven_KeyWord")=TempDrefDataDriven_KeyWord
'msgbox Environment("DrefDataDriven_KeyWord")

Environment("blnNoTDExists")=blnNoTDExists
'msgbox "boolvalue is "& blnNoTDExists

											
										Set o1=createobject("TafCore.CoreEngine")
										If Environment("BPT") Then
                                    		o1.ExecuteTestCaseBPT()
                                    	Else
                                    		o1.ExecuteTestCase()
                                    	End If

											If Environment("Execute") Then
												Call startTest()
											End If
                                            
											call PerformExecuteTest()
											'TAF 10.1 new code Start
											If Environment("BlnCodeGeneration")=1  Then
												Call UpdateScript("'******************"&"TestCase:  "&Environment("TestCaseName1")&"   end"&"**************************")     '10.1
												Call UpdateScript("")
											End If
											 'TAF 10.1 new code End   

									Else																																											  'Santosh 10 Sept				
										Reporter.ReportEvent micFail,"Verifying for the Test Case in the Test Case file","Test Case "&Environment("TestCaseName1")&" Doesnot exists in the test case file" &  strTCFileFolder & "\" & strTCFileName 'Santosh 10 Sept
										WriteToEvent("Failed" & vbtab &"Test Case "&Environment("TestCaseName1")&" Doesnot exists in the test case file")  'Santosh 10 Sept
									End If
								'End If			
								End If
								If Environment("FlagStepFailureOccured") and Environment("strTCSeverity")="showstopper" then ' Condition to verify Test Case fail and severity in a Scenario
									FlagIsShowStopper=True         								 
								End if
							End If
						End if
						DataTable.GetSheet("TestPlan").setCurrentRow(intScenarioCounter+1)
				Wend
			Next
				intScenarioCounter=intScenarioCounter+1
				If Lcase(Trim(DataTable.GetSheet("TestPlan").getParameter("Scenario_Keyword")))="endofscenario" Then
					If Environment("TSorTC") Then
							'indicate it is end of scenerio
								Environment("ScenerioEnd")=True 
                                Environment("ScenerioStepCounter")=0 ' Reset the test scenerio step count 
								Environment("TSorTC")=False
								FlagIsShowStopper=False
								Environment("FlagStepFailureOccured")=False
                          If Trim(Lcase(Environment("ExcelResults"))) <> "no"  and Trim(Lcase(Environment("IsSystemSlow")))<> "yes" Then
								'Select the Result Sheet
'								Set objWorkBook = appExcel.Workbooks.Open (Environment("ResultFilePath"))




'''''''''''''''''''''''''''''''''''''''''''Modified by Somesh on 25-07-2011'''''''''''''''''''''''''''''''''''''''''''''''''''''''
							  If Trim(Lcase(Environment("IsSummaryResult"))) = "yes" Then
							  
								Set objSheet = appExcel.Sheets("Test_Summary") 
								appExcel.Sheets("Test_Summary").Select
								If not(Environment("ScenerioErrorIncremented")) Then
									objSheet.Range("C11").Value =objSheet.Range("C11").Value + 1 ' increment scenerio count if the scenerio has passed
									'Save the Workbook
'									objWorkBook.Save
'									appExcel.Quit 
								else 
								'''''''''''''''''''' Commented By somesh'''''''''''''''''
'									SetRow=Environment("ScenerioLineHolder")
'									objSheet.Cells(SetRow, 3).Value = "Failed"
'									objSheet.Range("C" & SetRow).Font.ColorIndex = 3

									'''''''''''''''''''' Commented By somesh'''''''''''''''''
										
'									TCRow=Environment("ScenerioLineHolder")
'									objSheet.Cells(TCRow, 3).Value = "Failed"
'									objSheet.Range("C" & TCRow).Font.ColorIndex = 3
									'Save the Workbook
'									objWorkBook.Save
'									appExcel.Quit
								End If
							Else 
								
								Set objSheet = appExcel.Sheets("Test_Summary") 
								appExcel.Sheets("Test_Summary").Select
								If not(Environment("ScenerioErrorIncremented")) Then
									objSheet.Range("C11").Value =objSheet.Range("C11").Value + 1 ' increment scenerio count if the scenerio has passed
									'Save the Workbook
'									objWorkBook.Save
'									appExcel.Quit 
								else 
							
									SetRow=Environment("ScenerioLineHolder")
									objSheet.Cells(SetRow, 3).Value = "Failed"
									objSheet.Range("C" & SetRow).Font.ColorIndex = 3

																	
									'Save the Workbook
'									objWorkBook.Save
'									appExcel.Quit
								End If
							End If	
				''''''''''''''''''''''Modifiaction End ''''''''''''''''''''''
				
				


'								Set objSheet = appExcel.Sheets("Test_Summary") 
'								appExcel.Sheets("Test_Summary").Select
'								If not(Environment("ScenerioErrorIncremented")) Then
'									objSheet.Range("C11").Value =objSheet.Range("C11").Value + 1 ' increment scenerio count if the scenerio has passed
									'Save the Workbook
'									objWorkBook.Save
'									appExcel.Quit 
'								else 
'									SetRow=Environment("ScenerioLineHolder")
'									objSheet.Cells(SetRow, 3).Value = "Failed"
'									objSheet.Range("C" & SetRow).Font.ColorIndex = 3
									'Save the Workbook
'									objWorkBook.Save
'									appExcel.Quit
'								End If
								
								
								
						End If 
					End If
			else	
				While Lcase(Trim(DataTable.GetSheet("TestPlan").getParameter("Scenario_Keyword")))<>"endofscenario"
					intScenarioCounter=intScenarioCounter+1
					DataTable.GetSheet("TestPlan").SetCurrentRow(intScenarioCounter)							
				Wend
			End If''Scenario finished	

						'TAF 10.1 new code Start
						If Environment("BlnCodeGeneration")=1  Then               '10.1

								If Not Trim(Lcase(Environment("ExcelResults"))) ="no"  Then
										Call UpdateScript("objResWorkBook.Save")
										Call UpdateScript("appExcel.Quit")
										Call UpdateScript("AttachFileToCurrentTestSetTest"&" "&" Environment("&"""ResultFilePath"""&")")
										Call UpdateScript("'******************"&"TestScenario:  "&strTestScriptName&"   end"&"**************************************")         '10.1
										Call UpdateScript("")
								End If
						
						End If
                	    'TAF 10.1 new code End

		ElseIf Lcase(Trim(DataTable.GetSheet("TestPlan").getParameter("Scenario_Keyword")))="testcase" And Lcase(Trim(DataTable.GetSheet("TestPlan").getParameter("TestCaseName")))=strTestScriptName And not(blnIsScenario) then''Test Case Execution when no scenario exists	

					'TAF 10.1 new code Start
					If Environment("BlnCodeGeneration")=1 Then         '10.1
						Call CreateFileForCode()
						Call UpdateScript("On Error Resume Next")
						Call UpdateScript("Call MakePathNonTAF()")				'10.1
					End If
					'TAF 10.1 new code End
				
				Err.clear
				On error resume next
			Environment("Description1")= DataTable.GetSheet("TestPlan").getParameter("Description")  ' ARgus 
			If err.number > 0  Then
				strDescription = ""
				Else
				strDescription =  Environment("Description1")
				
			End If
			On Error goto 0

			DataTable.GetSheet("TestPlan").setCurrentRow(intScenarioCounter)
			'If  Lcase(Trim( DataTable.GetSheet("TestPlan").getParameter("TestCaseName")))="endofrow"  then
            If  Lcase(Trim( DataTable.GetSheet("TestPlan").getParameter("TestCaseName")))="endofrow"  or Lcase(Trim( DataTable.GetSheet("TestPlan").getParameter("TestCaseName")))="eof" then
				Exit For
			End If

				Environment("WithoutTestData")=False

				AppMapPath =  DataTable.GetSheet("TestPlan").getParameter("AppMapPath")
				AppMapPath_Relative=Trim(AppMapPath)
				strAppMapPath= Environment("vegaTestSuite")  &  DataTable.GetSheet("TestPlan").getParameter("AppMapPath")
				Environment("TestCaseName1")=DataTable.GetSheet("TestPlan").GetParameter("TestCaseName")
				Environment("strTCSeverity")=LCase(Trim(DataTable.GetSheet("TestPlan").GetParameter("Severity_KeyWord")))
				Call Header(Environment("EventLogPath"), LCase(Trim(Environment("EventLog")))) 
				Call Header(Environment("ScriptLogPath"), LCase(Trim(Environment("ScriptLog")))) 

								'TAF 10.1 new code Start
								If Environment("BlnCodeGeneration")=1 Then
									Call UpdateScript("'******************"&"TestCase:  "&Environment("TestCaseName1")&"   start"&"**************************")       '10.1
									Call UpdateScript("")
								End If
								'TAF 10.1 new code End
								

			'Trao modification to make appmap is not mandatory--------------TAF10
            If  AppMapPath_Relative="" and TPMasterAppMap_Relative="" Then
						    Environment("AppMapCount")=0
							Environment("AppMapPath")=""
			Else
				' Modified by Srikanth --------------------------TAF10
				If VerifyAndDownloadAppMap(AppMapPath) = 0  then
					If  VerifyAndDownloadAppMap(TPMasterAppMap) = 0 then
						blnFlagSkipTestCase=true
						If Environment("strTCSeverity")="showstopper" Then
							FlagIsShowStopper=True
						End If
					else
						Reporter.ReportEvent micWarning,"Verifying APPMAP for "& Environment("TestCaseName1")&" Test Case","Copied Default Test Plan level APPMAP i.e. "& strTPMasterAppMap     
						WriteToEvent("Warning" & vbtab & "Copied Default Test Plan level APPMAP i.e. "& strTPMasterAppMap &" for Test Case " & Environment("TestCaseName1") )
						strAppMapPath=strTPMasterAppMap
					end if
				else
					strAppMapPath=strAppMapPath
				End If
				' Modification Done----------TAF10
			End If   'End of Trao Modification-----------TAF10

'	                 arrTemp = Split(strAppMapPath,"\")
'	                 strAppMapFileName = arrTemp(ubound(arrTemp))
'	                 	arrTemp = Split(strAppMapPath,"\" & strAppMapFileName)
'		strAppMapFileFolder = arrTemp(lbound(arrTemp))
'		'msgbox " loading - AppMap"
'		DownloadAttachment  strAppMapFileName,strAppMapFileFolder,Environment("TempAppMap"),"TRUE"
'		'msgbox "Completed - loading - AppMap"
'     	Environment("AppMapPath")=  Environment("TempAppMap")&   "\" & strAppMapFileName 
'     	Environment("QCAppMapPath")= strAppMapPath

				TCFilePath = DataTable.GetSheet("TestPlan").getParameter("TestCaseFilePath")
				strTCFilePath= Environment("vegaTestSuite")  &  DataTable.GetSheet("TestPlan").getParameter("TestCaseFilePath")
				If not blnFlagSkipTestCase Then
					If  TCFilePath = ""Then
						If  TPMasterTestCase = ""  Then
							blnFlagNoTestCase=true
                            If Environment("strTCSeverity")="showstopper" Then
								FlagIsShowStopper=True
							End If
						else
							Reporter.ReportEvent micWarning,"Verifying TestCase file for "&Environment("TestCaseName1")&" test case","Copied Default Test Plan level Test Case i.e. "& strTPMasterTestCase     
							WriteToEvent("Warning" & vbtab & "Copied Default Test Plan level Test Case file i.e. "& strTPMasterTestCase &" for Test Case " & Environment("TestCaseName1"))
							strTCFilePath=strTPMasterTestCase
						End If
					else
						strTCFilePath= Environment("vegaTestSuite")  &  DataTable.GetSheet("TestPlan").getParameter("TestCaseFilePath")
					End if
				End if

	                 arrTemp = Split(strTCFilePath,"\")
	                 strTCFileName = arrTemp(ubound(arrTemp))
	                 	arrTemp = Split(strTCFilePath,"\" & strTCFileName)
		strTCFileFolder = arrTemp(lbound(arrTemp))
		'msgbox "loading - Test Case"
		DownloadAttachment  strTCFileName,strTCFileFolder,Environment("TempTestCases") ,"TRUE"
		'msgbox "Completed - loading - Test Case"
	    Environment("TestCasePath")= Environment("TempTestCases")  &    "\" & strTCFileName 
        Environment("QCTestCasePath")=strTCFilePath

'				strTestDataFilePath=DataTable.GetSheet("TestPlan").GetParameter("TestDataPath")'' Verifying for the existance of Test Data file if does not exists copying from the default level

				TestDataFilePath = DataTable.GetSheet("TestPlan").GetParameter("TestDataPath") 
				 TestDataFilePath_Relative=Trim(TestDataFilePath)
                 strTestDataFilePath1= Environment("vegaTestSuite")  &  DataTable.GetSheet("TestPlan").GetParameter("TestDataPath")   'Raj
				'strTestDataFilePath = Replace(lcase(Trim(strTestDataFilePath1)),"%client%",Lcase(Trim(Environment("Client"))))  'Raj
				strTestDataFilePath = GetDynamicParameter(strTestDataFilePath1)

         If  TestDataFilePath_Relative="" and TPMasterTestData_Relative="" Then
			  Environment("WithoutTestData")=True
			  Environment("strTestDataFilePath")=""
		 Else
				If blnFlagSkipTestCase=false and blnFlagNoTestCase=false Then
					If  TestDataFilePath = ""  then
						If  TPMasterTestData = "" Then
							blnNoTDExists=true
						Else
							Reporter.ReportEvent micWarning,"Verifying Test Data  file for "& Environment("TestCaseName1")&"  Test Case","Copied Default Test Plan level Test Data file  i.e. "& strTPMasterTestData     
							WriteToEvent("Warning" & vbtab & "Copied Default Test Plan level Test Data file i.e. "& strTPMasterTestData &" for Test Case " & Environment("TestCaseName1"))
							strTestDataFilePath = strTPMasterTestData
						End If
					End if	                                                                
				End if

	          arrTemp = Split(strTestDataFilePath,"\")
	                strTDFileName = arrTemp(ubound(arrTemp))
	               	arrTemp = Split(strTestDataFilePath,"\" & strTDFileName)
					strTDFileFolder = arrTemp(lbound(arrTemp))
					'msgbox " loading - Test Data"
					DownloadAttachment  strTDFileName,strTDFileFolder,Environment("TempTestData") ,"TRUE"
					'msgbox "Completed - loading - Test Data"
					Environment("strTestDataFilePath")= Environment("TempTestData") &    "\" & strTDFileName 
		           Environment("QCTestDataFilePath")=strTestDataFilePath

		End If
				If  blnFlagSkipTestCase then   ''Execute the test case when AppMap exists 
					reporter.ReportEvent micFail,"Verifying the  AppMap File" & Environment("QCAppMapPath") & " for test case "& Environment("TestCaseName1"), "AppMap File doesn't exists"		
					WriteToEvent("Fail" & vbtab &" AppMap File "&Environment("QCAppMapPath")&" does not exist")
				Else
					If blnFlagNoTestCase=true Then
						reporter.ReportEvent micFail, "Verifying for the Test Case " &  Environment("QCTestCasePath")  , " Test case "& Environment("TestCasePath") & " file does not exist"
						WriteToEvent("Fail" & vbtab & Environment("QCTestCasePath") & " file does not exist")
					else
						'Individual Test Case                 
						Environment("TSorTC")=False
						Environment("TestCaseCounter")=0
						If  verifysheetexists(environment("TestCasePath"), Environment("TestCaseName1"))  Then				'Santosh 10 Sept
							'Call ExecuteTestCase()				''Done																																'Santosh 10 Sept
							'Call ExecuteTestCase
							
							'  Changes for DLL
If Environment("BPT") Then
	Datatable.AddSheet("Global")
	DataTable.GetSheet("Global").AddParameter "TestCaseSheetName",""
	DataTable.GetSheet("Global").AddParameter "TestCaseFilePath",""
	DataTable.GetSheet("Global").AddParameter "TestDataDBPath",""
End If
'DataTable.GetSheet("Global").AddParameter "Description","" ' ARgus 
DataTable("TestCaseSheetName","Global")=DataTable.GetSheet("TestPlan").GetParameter("TestCaseName")   
Environment("TestCaseSheetNameToResults") = DataTable("TestCaseSheetName","Global")   ' Argus 
DataTable("TestCaseFilePath","Global")=DataTable.GetSheet("TestPlan").GetParameter("TestCaseFilePath")		
Datatable("TestCaseFilePath","Global")=Datatable.GetSheet("TestPlan").GetParameter("TestCaseFilePath")
Datatable("TestDataDBPath", "Global") = Environment("strTestDataFilePath")
Environment("TestDataPath") = Environment("strTestDataFilePath")
TempTestDataSheetName=Datatable.GetSheet("TestPlan").GetParameter("TestDataSheetName")
Environment("TestDataSheetName") = TempTestDataSheetName

'set below line putput to environment variable
TempDrefDataRow_Keyword=Datatable.GetSheet("TestPlan").GetParameter("DataRow_Keyword")
Environment("DrefDataRow_Keyword")=TempDrefDataRow_Keyword
TempDrefDataDriven_KeyWord=Datatable.GetSheet("TestPlan").GetParameter("DataDriven_KeyWord")
Environment("DrefDataDriven_KeyWord")=TempDrefDataDriven_KeyWord
'msgbox Environment("DrefDataDriven_KeyWord")

Environment("blnNoTDExists")=blnNoTDExists
'msgbox "boolvalue is "& blnNoTDExists
										Set o1=createobject("TafCore.CoreEngine")
										If Environment("BPT") Then
                                    		o1.ExecuteTestCaseBPT()
                                    	Else
                                    		o1.ExecuteTestCase()
                                    	End If

										If Environment("Execute") Then
											Call startTest()
										End If
										
										call PerformExecuteTest()

										'TAF 10.1 new code Start
										If Environment("BlnCodeGeneration")=1 Then

												If  Not Trim(Lcase(Environment("ExcelResults"))) ="no"  Then
														Call UpdateScript("objResWorkBook.Save")
														Call UpdateScript("appExcel.Quit")
														Call UpdateScript("AttachFileToCurrentTestSetTest"&" "&" Environment("&"""ResultFilePath"""&")")
														Call UpdateScript("'******************"&"TestCase:  "&Environment("TestCaseName1")&"   end"&"**************************")       '10.1
														Call UpdateScript("")
												End If
											
										End If
										'TAF 10.1 new code End

										
'  end of changes	
						Else                                                                                                                                                                                     'Santosh 10 Sept
							Reporter.ReportEvent micFail,"Verifying for the Test Case in the Test Case file","Test Case "&Environment("TestCaseName1")&" Doesnot exists in the test case file"     'Santosh 10 Sept
							WriteToEvent("Failed" & vbtab &"Test Case "&Environment("TestCaseName1")&" Doesnot exists in the test case file" & strTCFileFolder & "\"  & strTCFileName )     'Santosh 10 Sept
						End If	
					End if
				End if   ''Test case execution completes
		else
			ScenarioCounter=ScenarioCounter+1
		End if
		CurrentScenario=strNewScenario
		If Environment("ScenerioEnd") OR not(Environment("TSorTC")) Then
			Exit For
		End If
	Next
'	Select the Result Sheet
	If Trim(Lcase(Environment("ExcelResults"))) <> "no"  and Trim(Lcase(Environment("IsSystemSlow")))<> "yes" Then
'		Set objWorkBook = appExcel.Workbooks.Open (Environment("ResultFilePath"))
		Set objSheet = appExcel.Sheets("Detailed_Report")
		appExcel.Sheets("Detailed_Report").Select
		Row=Environment("IncValHolder")
		objSheet.Range("A" & Row & ":F" & Row).Interior.ColorIndex =53 
		objSheet.Range("A" & Row).value="Test Execution End"
		objSheet.Range("A" & Row & ":F" & Row).Font.ColorIndex = 19
		objSheet.Range("A" & Row).Font.Bold = True   			
		'Save the Workbook
		Set objSheet = appExcel.Sheets("Test_Summary")
		appExcel.Sheets("Test_Summary").Select
		
		
			'------------------------------Modified by Somesh
'		If Trim(Lcase(Environment("IsSummaryResult"))) = "yes" Then			
'		 If not(blnIsScenario) Then
'			objSheet.Range("C13").Value=objSheet.Range("C13").Value +objSheet.Range("C10").Value ' Sum the execution 
'			objSheet.Range("C14").Value=objSheet.Range("C14").Value +objSheet.Range("C11").Value ' Sum the execution 
'			objSheet.Range("C15").Value=objSheet.Range("C15").Value +objSheet.Range("C12").Value ' Sum the execution 
'			objSheet.Rows("10:12").Delete
'		 else
'			objSheet.Range("C10").Value=objSheet.Range("C13").Value +objSheet.Range("C10").Value ' Sum the execution 
'			objSheet.Range("C11").Value=objSheet.Range("C14").Value +objSheet.Range("C11").Value ' Sum the execution 
'			objSheet.Range("C12").Value=objSheet.Range("C15").Value +objSheet.Range("C12").Value ' Sum the execution 
'			objSheet.Rows("13:15").Delete
'			objSheet.Range("C13").Value=objSheet.Range("C13").Value +objSheet.Range("C10").Value ' Sum the execution 
'			objSheet.Range("C14").Value=objSheet.Range("C14").Value +objSheet.Range("C11").Value ' Sum the execution 
'			objSheet.Range("C15").Value=objSheet.Range("C15").Value +objSheet.Range("C12").Value ' Sum the execution 
'			objSheet.Rows("10:12").Delete
'		 End If
'		Else 
'		 If not(blnIsScenario) Then
'			objSheet.Range("C13").Value=objSheet.Range("C13").Value +objSheet.Range("C10").Value ' Sum the execution 
'			objSheet.Range("C14").Value=objSheet.Range("C14").Value +objSheet.Range("C11").Value ' Sum the execution 
'			objSheet.Range("C15").Value=objSheet.Range("C15").Value +objSheet.Range("C12").Value ' Sum the execution 
'			objSheet.Rows("10:12").Delete
'		 else
'			objSheet.Range("C10").Value=objSheet.Range("C13").Value +objSheet.Range("C10").Value ' Sum the execution 
'			objSheet.Range("C11").Value=objSheet.Range("C14").Value +objSheet.Range("C11").Value ' Sum the execution 
'			objSheet.Range("C12").Value=objSheet.Range("C15").Value +objSheet.Range("C12").Value ' Sum the execution 
'			objSheet.Rows("13:15").Delete
			
'		 End If
		 
'		End If	
		''''''''''''''''''End Of Modification
				
		
		
'		If not(blnIsScenario) Then
'			objSheet.Range("C13").Value=objSheet.Range("C13").Value +objSheet.Range("C10").Value ' Sum the execution 
'			objSheet.Range("C14").Value=objSheet.Range("C14").Value +objSheet.Range("C11").Value ' Sum the execution 
'			objSheet.Range("C15").Value=objSheet.Range("C15").Value +objSheet.Range("C12").Value ' Sum the execution 
'			objSheet.Rows("10:12").Delete
'		else
'			objSheet.Range("C10").Value=objSheet.Range("C13").Value +objSheet.Range("C10").Value ' Sum the execution 
'			objSheet.Range("C11").Value=objSheet.Range("C14").Value +objSheet.Range("C11").Value ' Sum the execution 
'			objSheet.Range("C12").Value=objSheet.Range("C15").Value +objSheet.Range("C12").Value ' Sum the execution 
'			objSheet.Rows("13:15").Delete
'		End If
		
			
			If Trim(Lcase(Environment("IsSummaryResult"))) <> "yes"   Then
					objSheet.Range("D18").Value = objSheet.Range("C16").Value
			End If
			appExcel.Sheets("Test_Summary").Activate  
			appExcel.Sheets("Test_Summary").Range("J2:J20000").Copy  
			appExcel.Sheets("Test_Summary").Range("A1").PasteSpecial  
					
		
			objSheet.Range("E18").Value = strDescription  ' Argus
			objSheet.Range("E18").Font.ColorIndex = 12
			objSheet.Range("E18").Font.Bold = True

'        Call ProtectResults(objSheet)
		objResWorkBook.Save
		appExcel.Quit
	End If 

	If Trim(LCase(Environment("HTMLResults")))<>"no" Then
		Call ReportHTMLResults(strUIName, strExpected, strData, strActual, Status,objParent)
			Set objFileSystemObject = CreateObject("Scripting.FileSystemObject")
			Set objTextFileObject = objFileSystemObject.CreateTextFile(Environment("CurrentResultFile"),True)
		 	HTMLSummary()
		 	objTextFileObject.Write HTMLContent
			Set objFileSystemObject=Nothing
'$$$$$$$$$$$$$$$$$$$$$$$$$$ -- END MODIFICATION -- $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$			
	End If 


'arrtemp = Split(Environment("ResultFilePath"), Environment("TempResultPath") & "\" )
'FromFileLoc = Environment("TempResultPath") & "\"
'FromFileName = arrtemp(1)
'Subject\Automation 9.2\Amisys\4.0.0 Regression\Demo\Automation_DSTWS_Framework\Controller\TestPlan.xls
'AttachFileToQC  "E:\TempAutomation\KeywordDriven\Results\"  ,"TestPlan_11_September_2009_16_50_28_Result.xls" ,"Automation 9.2\Amisys\4.0.0 Regression\Demo\Automation_DSTWS_Framework\TemplateScript"   ,   "SampleTest"
'arrTemp = Split(Datatable.Value("ResultPath"), "\")
'ToQCTestName = arrTemp(ubound(arrTemp))
'arrTemp =  Split(Datatable.Value("ResultPath"), "\" & ToQCTestName)
'arrTemp1 = Split(lcase(arrTemp(0)),"subject\")
'ToQCTestPath =arrTemp1(1)

 'AttachFileToQC FromFileLoc,FromFileName,ToQCTestPath,ToQCTestName

	If Trim(Lcase(Environment("IsSummaryResult"))) = "yes"   Then
		If Environment("LastTest")  Then
			If Trim(Lcase(Environment("ExcelResults"))) <> "no"  and Trim(Lcase(Environment("IsSystemSlow")))<> "yes" Then
					wait(5)  ' time to allow Excel Results to save	
					AttachFileToCurrentTestSetTest  Environment("ResultFilePath")	
			End If
		End If
	Else
			If Trim(Lcase(Environment("ExcelResults"))) <> "no"  and Trim(Lcase(Environment("IsSystemSlow")))<> "yes" Then
				wait(5)  ' time to allow Excel Results to save	
				AttachFileToCurrentTestSetTest  Environment("ResultFilePath")	
			End If
	End If


'If Trim(Lcase(Environment("ExcelResults"))) <> "no"  and Trim(Lcase(Environment("IsSystemSlow")))<> "yes" Then
'	wait(5)  ' time to allow Excel Results to save	
'	AttachFileToCurrentTestSetTest  Environment("ResultFilePath")	
'End If

If Trim(LCase(Environment("HTMLResults")))<>"no" Then
	wait(2)  ' time to allow HTML Results to save	
	AttachFileToCurrentTestSetTest  Environment("CurrentResultFile")
			'TAF 10.1 new code Start
			If Environment("BlnCodeGeneration")=1  Then               '10.1
    				Call UpdateScript("SaveHTML()")	
					Call UpdateScript("AttachFileToCurrentTestSetTest"&" "&" Environment("&"""CurrentResultFile"""&")")
					Call UpdateScript("'******************"&"TestScenario:  "&strTestScriptName&"   end"&"**************************************")         '10.1
					Call UpdateScript("")		
			End If
			'TAF 10.1 new code End
End If

'Call KillAnyOpenExcelApplications
'fso.Deletefolder(Environment("WorkingDirectory") & "\TempAutomation")
End Function




 
Function RunFromLS(strTestScriptName,blnIsScenario)

	Dim App 'As Application
	Set App = CreateObject("QuickTest.Application")
	App.Launch
	App.Visible = True
	strTestScriptName=LCase(Trim(strTestScriptName)) 
   Environment("TestPlanPath") = Environment("TestControllerPath")
	arrTemp = Split( Environment("TestPlanPath") ,"\Controller")
	EnvPath=arrTemp(0)
	DataTable.AddSheet("TestPlan")
	'******Load testplan from temporary file 7.0.11 - Fanweb*************************
	If Environment("UseMsAccessDB") Then
		LoadDataTableFromDB Environment("TestPlanPath"),strTestScriptName,blnIsScenario
	Else
		If LoadFromTempLS() Then
			fso.CopyFile Environment("TestPlanPath"),Environment("TempLS")&"\"
			arrTempPlan = Split(Environment("TestControllerPath"),"\")
            tempPlan=Environment("TempLS")& "\" & arrTempPlan(ubound(arrTempPlan))
			DataTable.ImportSheet tempPlan,"TestPlan","TestPlan"        
		Else
			DataTable.ImportSheet Environment("TestPlanPath"),"TestPlan","TestPlan"
		End If
		If Err.Number<>0 Then
			DataTable.ImportSheet Environment("TestPlanPath"),"TestPlan","TestPlan"
		End If
	End If
	
'******************************************************************************************************

	'Environment("ResultPath")=Datatable.Value("ResultPath")
	If VerifyEnvVariable("ResultsPath") Then
		Environment("ResultPath")=Environment("ResultsPath")
	Else
		Environment("ResultPath")=Datatable.Value("ResultPath")
	End If
	MakePath()
	Environment("ExecutionStarted")=False
	Environment("ReusableTestCaseFolderPath")=Environment("vegaTestSuite") & "\ReusableTestCases"
	CurrentScenario=""
	Environment("blnFlagNewScenario")=false
	Environment("TSorTC")=Null
	Environment("strNewScenario")=Null
	Environment("IncFlag")=False
	Environment("CasesCount")=0
	Environment("ScenCounter")=0
	Environment("TestCaseCounter")=0
	Environment("EndOfScenarios")=0
	Environment("ScenerioEnd")=False 
	Environment("ScenerioEncountered")=False
	Environment("FlagStepFailureOccured")=False
	Environment("strTCSeverity")=Null
	FlagStepFailureOccured=False
	Environment("CountFlag")=0
	ScenarioCounter=0
	TestStepCounter=0
	Environment("blnNoTDExists")=Null
	FlagIsShowStopper=False
	Environment("TestStepCounter")=TestStepCounter
	Environment("ScenerioStepCounter")=0 '  scenario  step counter
	Environment("FirstOnly")=False 
	Environment("LastOnly")=False 
	Environment("tagCount")=False
	'TAF 10.1 new code Start.
	 Environment("NoOfDatarows")=0
	 intFirstTime=0
	'TAF 10.1 new code End

	 
	On Error resume next
	If err.number = 0  and Lcase(trim(Environment("BaseState")))  <> "n/a" Then
		ExecutePreReqScript Lcase(trim(Environment("BaseState"))) 
	End If
	On error goto 0




	
	For intScenarioCounter=2 to DataTable.GetSheet("TestPlan").GetRowCount							'Loop which rolls on the test plan for test scenario or test case
		If intScenarioCounter=2 Then																													'Code to store the Test Plan level master default files like AppMap, Test Case and Test Data			
			DataTable.GetSheet("TestPlan").setCurrentRow(intScenarioCounter-1)
			strTPMasterAppMap=Environment("vegaTestSuite") & DataTable.GetSheet("TestPlan").getParameter("AppMapPath")
			strTPMasterAppMap_Relative=Trim(DataTable.GetSheet("TestPlan").getParameter("AppMapPath"))
			strTPMasterTestCase= Environment("vegaTestSuite") &  DataTable.GetSheet("TestPlan").getParameter("TestCaseFilePath")
			strTPMasterTestData1_Relative = Trim(DataTable.GetSheet("TestPlan").getParameter("TestDataPath"))
			strTPMasterTestData1 = Environment("vegaTestSuite") &  DataTable.GetSheet("TestPlan").getParameter("TestDataPath")  'raj
			strTPMasterTestData = GetDynamicParameter(strTPMasterTestData1)
		End If

		'TAF 10.1 new code Start.  
		If   intFirstTime=0 Then
			'Count the required line directly in test plan excel, rathen than moving line by line for the required row.
		     intScenarioCounter=getContNumber(Environment("TestPlanPath"), strTestScriptName, blnIsScenario)
			  intFirstTime=1         'This is to avoid for iteration to come in to this IF condition. If it come, every time intScenarioCounter will have the same number, leads to infinite loop
		    If not isNumeric(intScenarioCounter) Then
			intScenarioCounter=DataTable.GetSheet("TestPlan").GetRowCount
		    End If
		End If
		  'TAF 10.1 new code End

		DataTable.GetSheet("TestPlan").setCurrentRow(intScenarioCounter)
		blnFlagSkipTestCase=false
		blnFlagNoTestCase=false
		blnNoTDExists=false
		If  Lcase(Trim(DataTable.GetSheet("TestPlan").getParameter("Scenario_Keyword")))="scenario" And LCase(Trim(DataTable.GetSheet("TestPlan").getParameter("TestCaseName")))=strTestScriptName And blnIsScenario then
				'TAF 10.1 new code Start
				If Environment("BlnCodeGeneration")=1 Then
					Call CreateFileForCode()		'10.1
					Call UpdateScript("'******************"&"TestScenario:  "&strTestScriptName&"   start"&"**************************************")
					Call UpdateScript("")
					Call UpdateScript("On Error Resume Next")
					Call UpdateScript("Call MakePathNonTAF()")
				End If
			    'TAF 10.1 new code End
			

               Environment("ScenarioDataKeys") = DataTable.GetSheet("TestPlan").getParameter("DataRow_Keyword")
               Environment("ScenarioDataKeys") = DataTable.Value("DataRow_Keyword","TestPlan")         ' DSTGS 
				ScenarioCounter=ScenarioCounter+1
				Environment("ScenarioCounter")=ScenarioCounter
				strNewScenario=DataTable.GetSheet("TestPlan").getParameter("TestCaseName")
				If not(strNewScenario=CurrentScenario) Then
					Environment("strNewScenario")=strNewScenario
					Environment("blnFlagNewScenario")=true 
				End If
				Err.clear
				On error resume next
			Environment("Description1")= DataTable.GetSheet("TestPlan").getParameter("Description")  		' ARgus 
			If err.number > 0  Then
				strDescription = ""
			Else
				strDescription =  Environment("Description1")
			End If
			On Error goto 0
				strScenarioMasterAppMap_Relative= Trim(DataTable.GetSheet("TestPlan").getParameter("AppMapPath"))
				strScenarioMasterAppMap=  Environment("vegaTestSuite")  &  DataTable.GetSheet("TestPlan").getParameter("AppMapPath")
				strScenarioMasterTC= Environment("vegaTestSuite") &  DataTable.GetSheet("TestPlan").getParameter("TestCaseFilePath")
				strScenarioMasterTD1_Relative = Trim(DataTable.GetSheet("TestPlan").getParameter("TestDataPath"))
				strScenarioMasterTD1 = Environment("vegaTestSuite")  &  DataTable.GetSheet("TestPlan").getParameter("TestDataPath")  'raj
			    strScenarioMasterTD = GetDynamicParameter(strScenarioMasterTD1)
                arrScenarioKeywods = 	Split(Environment("ScenarioDataKeys") ,";")        '  ' DSTGS  ' DSTGS 
				
				scenariorow = 	intScenarioCounter  ' DSTGS  ' DSTGS 

				''''''' Viswanadh
				
					If ubound(arrScenarioKeywods) = -1 Then
						iterationCnt =  1
					else
						If LCase(Trim(arrScenarioKeywods(0))) <> "all" Then
							iterationCnt = 	ubound(arrScenarioKeywods)+1
						Else
							arrScenarioKeywods = GetDataRows(strScenarioMasterTD1,Datatable.GetSheet("TestPlan").GetParameter("TestDataSheetName"))
							iterationCnt = 	ubound(arrScenarioKeywods)+1
						End If
					End If
				
			''''''''''''''''''''''''''''''''''''''

				For scit  = 1 to iterationCnt ' DSTGS  ' DSTGS 
					Environment("CurrentIteration") = scit  ' Argus 7.0.10
					If  ubound(arrScenarioKeywods) = -1 Then
						Environment("ScenarioDataKeys") =""
						else
						Environment("ScenarioDataKeys") =  arrScenarioKeywods(scit-1) ' DSTGS  ' DSTGS 
					End If
					
					'''' Scenario keys
						Environment("SceIte") = scit
					If scit=1 Then
							Environment("FirstOnly") = True
							else
							Environment("FirstOnly") = False
					End If

					if scit=iterationCnt Then
							Environment("LastOnly") = True
							Else
							Environment("LastOnly") = False
					End If
					
					'''' Scenario keys
						
				intScenarioCounter = scenariorow ' DSTGS  ' DSTGS 
				DataTable.GetSheet("TestPlan").setCurrentRow(intScenarioCounter) ' DSTGS  ' DSTGS 
				Do  While (Lcase(Trim(DataTable.GetSheet("TestPlan").getParameter("Scenario_Keyword")))<>"endofscenario") 			' Modified to Do While loop for Fanweb to implement  Execute Test scenario
					blnFlagSkipTestCase=false
					blnFlagNoTestCase=false
					blnNoTDExists=false
					intScenarioCounter=intScenarioCounter+1
                	DataTable.GetSheet("TestPlan").setCurrentRow(intScenarioCounter)
					''''''''''''''''''Viswa''''''''''''''''''''
'					Services.StartTransaction "ForLoop"
					
					
					For i=1 to DataTable.GetSheet("TestPlan").GetParameterCount														'Modified for Fanweb 7.0.11 - To implement Execute functionality in Scenario
						On Error Resume Next
						Err.Clear
						Lcase(DataTable.GetSheet("TestPlan").getParameter("Execute"))
						If Err.Number=0 Then
							If  Lcase(Cstr(DataTable.GetSheet("TestPlan").GetParameter(i).Name)) ="execute"Then
								While Lcase(DataTable.GetSheet("TestPlan").getParameter("Execute"))="no" 				
									intScenarioCounter=intScenarioCounter+1
									DataTable.GetSheet("TestPlan").setCurrentRow(intScenarioCounter)
								Wend
							End If
						Else
							Err.Clear
							Exit For
						End If

					Next

'					Services.EndTransaction "ForLoop"
					''''''''''''''''''Viswa''''''''''''''''''''


'					While Lcase(DataTable.GetSheet("TestPlan").getParameter("Execute"))="no" 				'Modified for Fanweb 7.0.11 - To implement Execute functionality in Scenario
'						intScenarioCounter=intScenarioCounter+1
'						DataTable.GetSheet("TestPlan").setCurrentRow(intScenarioCounter)
'					Wend

					If Lcase(Trim(DataTable.GetSheet("TestPlan").getParameter("Scenario_Keyword"))) = "endofscenario"  Then
							Exit Do
					End If

					If  Lcase(Trim( DataTable.GetSheet("TestPlan").getParameter("TestCaseName")))="endofrow" or Lcase(Trim( DataTable.GetSheet("TestPlan").getParameter("TestCaseName")))="eof" or Lcase(Trim( DataTable.GetSheet("TestPlan").getParameter("TestCaseName")))="x" then
						Exit For
					End If

					If FlagIsShowStopper=False  Then  
						strAppMapPath_Relative= Trim(DataTable.GetSheet("TestPlan").getParameter("AppMapPath"))
						strAppMapPath=Environment("vegaTestSuite") &  DataTable.GetSheet("TestPlan").getParameter("AppMapPath")
						Environment("TestCaseName1")=DataTable.GetSheet("TestPlan").GetParameter("TestCaseName")
						Environment("strTCSeverity")=Lcase(Trim(DataTable.GetSheet("TestPlan").GetParameter("Severity_KeyWord")))     ''Santosh
						Call Header(Environment("EventLogPath"), LCase(Trim(Environment("EventLog")))) 
						Call Header(Environment("ScriptLogPath"), LCase(Trim(Environment("ScriptLog")))) 
					'TAF 10.1 new code Start
					If Environment("BlnCodeGeneration")=1 Then
						 Call UpdateScript("'******************"&"TestCase:  "&Environment("TestCaseName1")&"   start"&"**************************")      '10.1
						 Call UpdateScript("")
					End If
					'TAF 10.1 new code End
                   

						Environment("WithoutTestData")=False     'Make it false before start of every case--TRao--TAF10

						'Modified by Trao to make appmap not mandatory------------TAF10
						If strTPMasterAppMap_Relative="" and  strAppMapPath_Relative="" and strScenarioMasterAppMap_Relative="" Then
							Environment("AppMapCount")=0
							Environment("AppMapPath")=""
						Else

							'Modified by Srikanth to access multiple AppMap-------------TAF10
						If VerifyAndDownloadAppMap(strAppMapPath)=0 then
							If VerifyAndDownloadAppMap(strScenarioMasterAppMap)=0 Then
								If  VerifyAndDownloadAppMap(strTPMasterAppMap)=0 then
									blnFlagSkipTestCase=true
                                    If Environment("strTCSeverity")="showstopper" Then
										FlagIsShowStopper=True
									End If
								else
									Reporter.ReportEvent micWarning,"Verifying APPMAP for "& Environment("TestCaseName1")&" test case","Copied Default Test Plan level APPMAP i.e. "& strTPMasterAppMap   
									WriteToEvent("Warning" & vbtab & "Copied Default Test Plan level APPMAP i.e. "& strTPMasterAppMap &" for Test Case " & Environment("TestCaseName1") )
									strAppMapPath=strTPMasterAppMap
								end if
							else
								Reporter.ReportEvent micWarning,"Verifying APPMAP for "& Environment("TestCaseName1")&" test case","Copied Default Test Scenario level APPMAP"       
								WriteToEvent("Warning" & vbtab & "Copied Default Test Scenario level APPMAP i.e. "& strScenarioMasterAppMap &" for Test Case " & Environment("TestCaseName1"))
								strAppMapPath=strScenarioMasterAppMap
							End If
						else
							strAppMapPath=strAppMapPath
						End If
						' Modification End--------------------------TAF10
					End If
				'Trao modification End-----------TAF10

					'' Verifying Test Case availability 
						strTCFilePath=Environment("vegaTestSuite") &  DataTable.GetSheet("TestPlan").getParameter("TestCaseFilePath")
						If not blnFlagSkipTestCase Then
							If VerifyFileExists(strTCFilePath)=0 Then
								If VerifyFileExists(strScenarioMasterTC)=0 Then
									If VerifyFileExists(strTPMasterTestCase)=0 Then
										blnFlagNoTestCase=true
                                        If Environment("strTCSeverity")="showstopper" Then
											FlagIsShowStopper=True
										End If
									else
										'strTCFilePath=strTPMasterTestCase
										Reporter.ReportEvent micWarning,"Verifying TestCase file for "& Environment("TestCaseName1")&" test case","Copied Default Test Plan level Test Case"    
										WriteToEvent("Warning" & vbtab & "Copied Default Test Plan level Test Case i.e. "& strTPMasterTestCase &" for Test Case " & Environment("TestCaseName1") ) 
										strTCFilePath=strTPMasterTestCase
									End If
								else
									'strTCFilePath=strScenarioMasterTC
									Reporter.ReportEvent micWarning,"Verifying TestCase file for "&  Environment("TestCaseName1") &" test case","Copied Default Test Scenario level Test Case"    
									WriteToEvent("Warning" & vbtab & "Copied Default Test Scenario level Test Case i.e. "& strScenarioMasterTC &" for Test Case " & Environment("TestCaseName1") ) 
									strTCFilePath=strScenarioMasterTC
								End If
							else
								strTCFilePath=Environment("vegaTestSuite") &  DataTable.GetSheet("TestPlan").getParameter("TestCaseFilePath")
							End if
						End if
						Environment("TestCasePath")=strTCFilePath

						strTestDataFilePath1_Relative=Trim(DataTable.GetSheet("TestPlan").GetParameter("TestDataPath"))
						strTestDataFilePath1=Environment("vegaTestSuite") &  DataTable.GetSheet("TestPlan").GetParameter("TestDataPath")   'Raj
						'strTestDataFilePath = Replace(lcase(Trim(strTestDataFilePath1)),"%client%",Lcase(Trim(Environment("Client"))))  'Raj
						strTestDataFilePath = GetDynamicParameter(strTestDataFilePath1)

					'Modified by Trao to make test data path not mandatory-----TAF10
					If  strTestDataFilePath1_Relative="" and strTPMasterTestData1_Relative=""  and strScenarioMasterTD1_Relative="" Then
					    Environment("WithoutTestData")=True
					    Environment("strTestDataFilePath")=""
				   Else

'						Environment("strTestDataFilePath")=DataTable.GetSheet("TestPlan").GetParameter("TestDataPath")
						If blnFlagSkipTestCase=false and blnFlagNoTestCase=false Then
							If verifyFileExists(strTestDataFilePath)=0 then
								If verifyFileExists(strScenarioMasterTD)=0 Then
									If verifyFileExists(strTPMasterTestData)= 0 Then
										blnNoTDExists=true
									Else
										Reporter.ReportEvent micWarning,"Verifying Test Data  file for "& DataTable.GetSheet("TestPlan").getParameter("TestCaseName")&"  Test Case","Copied Default Test Plan level Test Data filei.e. "& strTPMasterTestData &" for Test Case " & Environment("TestCaseName1")   'raj
										WriteToEvent("Warning" & vbtab & "Copied Default Test Plan level Test Data file i.e. "& strTPMasterTestData &" for Test Case " & Environment("TestCaseName1") ) 
										strTestDataFilePath= strTPMasterTestData
									End If
								Else
									Reporter.ReportEvent micWarning,"Verifying Test Data  file for "& DataTable.GetSheet("TestPlan").getParameter("TestCaseName")&"  Test Case","Copied Default Test Scenario level Test Data file  i.e. " & strScenarioMasterTD &" for Test Case " & Environment("TestCaseName1") 'raj
									WriteToEvent("Warning" & vbtab & "Copied Default Test Scenario level Test Data file i.e. "& strScenarioMasterTD &" for Test Case " & Environment("TestCaseName1") ) '
									strTestDataFilePath = 	strScenarioMasterTD
								End If
							End if	
						End if
						Environment("strTestDataFilePath")=strTestDataFilePath
                     End If
					 'End of Trao modification---TAF10

						If  blnFlagSkipTestCase then
							reporter.ReportEvent micFail,"Verifying the  AppMap File"&Environment("AppMapPath")&" for test case "& Environment("TestCaseName1"), "AppMap File doesn't exists"		
							WriteToEvent("Fail" & vbtab &" AppMap File "&Environment("AppMapPath")&" does not exist")
						else
							If blnFlagNoTestCase=true Then
								reporter.ReportEvent micFail, "Verifying the Test Case " & Environment("TestCasePath") , Environment("TestCasePath") & " file does not exist"
								WriteToEvent("Fail" & vbtab & Environment("TestCasePath") & " file does not exist")
							else
								'If not(FlagIsShowStopper)  Then
									Environment("TSorTC")=True
									If  verifysheetexists(environment("TestCasePath"), Environment("TestCaseName1"))  Then	  'Santosh 10 Sept
										'Call ExecuteTestCase()																																	  'Santosh 10 Sept					
										'Call ExecuteTestCase
							'  Changes for DLL
If Environment("BPT") Then
	Datatable.AddSheet("Global")
	DataTable.GetSheet("Global").AddParameter "TestCaseSheetName",""
	DataTable.GetSheet("Global").AddParameter "TestCaseFilePath",""
	DataTable.GetSheet("Global").AddParameter "TestDataDBPath",""
End If
'DataTable.GetSheet("Global").AddParameter "Description","" ' Argus
DataTable("TestCaseSheetName","Global")=DataTable.GetSheet("TestPlan").GetParameter("TestCaseName")    
Environment("TestCaseSheetNameToResults") = DataTable("TestCaseSheetName","Global")  ' Argus 
DataTable("TestCaseFilePath","Global")=DataTable.GetSheet("TestPlan").GetParameter("TestCaseFilePath")		
Datatable("TestCaseFilePath","Global")=Datatable.GetSheet("TestPlan").GetParameter("TestCaseFilePath")
Datatable("TestDataDBPath", "Global") = Environment("strTestDataFilePath")
Environment("TestDataPath") = Environment("strTestDataFilePath")
TempTestDataSheetName=Datatable.GetSheet("TestPlan").GetParameter("TestDataSheetName")
Environment("TestDataSheetName") = TempTestDataSheetName

'set below line putput to environment variable
TempDrefDataRow_Keyword=Datatable.GetSheet("TestPlan").GetParameter("DataRow_Keyword")
Environment("DrefDataRow_Keyword")=TempDrefDataRow_Keyword
TempDrefDataDriven_KeyWord=Datatable.GetSheet("TestPlan").GetParameter("DataDriven_KeyWord")
Environment("DrefDataDriven_KeyWord")=TempDrefDataDriven_KeyWord
'msgbox Environment("DrefDataDriven_KeyWord")

Environment("blnNoTDExists")=blnNoTDExists
'msgbox "boolvalue is "& blnNoTDExists
'										Set o1=createobject("TafCore.CoreEngine")
'										If Environment("BPT") Then
'                                    		o1.ExecuteTestCaseBPT()
'                                    	Else
'                                    		o1.ExecuteTestCase()
'                                    	End If
											ExecuteTestCase()
												If  Environment("Execute") Then
													Call startTest()
												End If
												
												Call PerformExecuteTest()
												'TAF 10.1 new code Start
												If Environment("BlnCodeGeneration")=1 Then
													Call UpdateScript("'******************"&"TestCase:  "&Environment("TestCaseName1")&"   end"&"**************************")     '10.1
													Call UpdateScript("")
												End If
												'TAF 10.1 new code End
                                                
'  end of changes
									Else																																											  'Santosh 10 Sept				
										Reporter.ReportEvent micFail,"Verifying for the Test Case in the Test Case file","Test Case "&Environment("TestCaseName1")&" Doesnot exists in the test case file"  'Santosh 10 Sept
										WriteToEvent("Failed" & vbtab &"Test Case "&Environment("TestCaseName1")&" Doesnot exists in the test case file")  'Santosh 10 Sept
									End If	
								'End If		
								End If	
								If Environment("FlagStepFailureOccured") and Environment("strTCSeverity")="showstopper" then ' Condition to verify Test Case fail and severity in a Scenario
									FlagIsShowStopper=True         								 
								End if
							End If
						End if
						DataTable.GetSheet("TestPlan").setCurrentRow(intScenarioCounter+1)
'				Wend
				Loop
				Next

				intScenarioCounter=intScenarioCounter+1
				If Lcase(Trim(DataTable.GetSheet("TestPlan").getParameter("Scenario_Keyword")))="endofscenario" Then
					If Environment("TSorTC") Then
							On error resume next
							'indicate it is end of scenerio
								Environment("ScenerioEnd")=True 
                                Environment("ScenerioStepCounter")=0 ' Reset the test scenerio step count 
								Environment("TSorTC")=False
								FlagIsShowStopper=False
								Environment("FlagStepFailureOccured")=False
								'Select the Result Sheet
                          If Trim(Lcase(Environment("ExcelResults"))) <> "no"  and Trim(Lcase(Environment("IsSystemSlow")))<> "yes" Then
'								Set objWorkBook = appExcel.Workbooks.Open (Environment("ResultFilePath"))
								Set objSheet = appExcel.Sheets("Test_Summary") 
								appExcel.Sheets("Test_Summary").Select
								
								
						'********************************Modified By Somesh************************************
								If Trim(Lcase(Environment("IsSummaryResult"))) = "yes" Then
								 If not(Environment("ScenerioErrorIncremented")) Then
									objSheet.Range("C11").Value =objSheet.Range("C11").Value + 1 ' increment scenerio count if the scenerio has passed
									'Save the Workbook
'									objWorkBook.Save
'									appExcel.Quit 
'								 Else 
'									TCRow=Environment("ScenerioLineHolder")
'									objSheet.Cells(TCRow, 3).Value = "Failed"
'									objSheet.Range("C" & TCRow).Font.ColorIndex = 3
									'Save the Workbook
'									objWorkBook.Save
'									appExcel.Quit
								 End If
								Else
								 If not(Environment("ScenerioErrorIncremented")) Then
									objSheet.Range("C11").Value =objSheet.Range("C11").Value + 1 ' increment scenerio count if the scenerio has passed
									'Save the Workbook
'									objWorkBook.Save
'									appExcel.Quit 
								else 
									SetRow=Environment("ScenerioLineHolder")
									objSheet.Cells(SetRow, 3).Value = "Failed"
									objSheet.Range("C" & SetRow).Font.ColorIndex = 3
									'Save the Workbook
'									objWorkBook.Save
'									appExcel.Quit
								End If 

										

						  End If
						  '***************************************End Modification************************************						
								
								
								
'								If not(Environment("ScenerioErrorIncremented")) Then
'									objSheet.Range("C11").Value =objSheet.Range("C11").Value + 1 ' increment scenerio count if the scenerio has passed
									'Save the Workbook
'									objWorkBook.Save
'									appExcel.Quit 
'								else 
'									SetRow=Environment("ScenerioLineHolder")
'									objSheet.Cells(SetRow, 3).Value = "Failed"
'									objSheet.Range("C" & SetRow).Font.ColorIndex = 3
									'Save the Workbook
'									objWorkBook.Save
'									appExcel.Quit
'								End If
								
								
								
								
						End If 
					End If
			else	
				While Lcase(Trim(DataTable.GetSheet("TestPlan").getParameter("Scenario_Keyword")))<>"endofscenario"
					intScenarioCounter=intScenarioCounter+1
					DataTable.GetSheet("TestPlan").SetCurrentRow(intScenarioCounter)							
				Wend
			End If''Scenario finished		

					'TAF 10.1 new code Start
					If Environment("BlnCodeGeneration")=1 Then
							If  Not Trim(Lcase(Environment("ExcelResults"))) ="no" Then
								Call UpdateScript("objResWorkBook.Save")
								Call UpdateScript("appExcel.Quit")
								Call UpdateScript("'******************"&"TestScenario:  "&strTestScriptName&"   end"&"**************************************")		'10.1
								Call UpdateScript("")
							End If
                    End If
					'TAF 10.1 new code End
					
		ElseIf Lcase(Trim(DataTable.GetSheet("TestPlan").getParameter("Scenario_Keyword")))="testcase" And Lcase(Trim(DataTable.GetSheet("TestPlan").getParameter("TestCaseName")))=strTestScriptName And not(blnIsScenario) then''Test Case Execution when no scenario exists	

			'TAF 10.1 new code Start
			If  Environment("BlnCodeGeneration")=1 Then
				Call CreateFileForCode()
				Call UpdateScript("On Error Resume Next")
				Call UpdateScript("Call MakePathNonTAF()")			'10.1
			End If
			'TAF 10.1 new code End

		
			DataTable.GetSheet("TestPlan").setCurrentRow(intScenarioCounter)
				Err.clear
				On error resume next
			Environment("Description1")= DataTable.GetSheet("TestPlan").getParameter("Description")  ' ARgus 
			If err.number > 0  Then
				strDescription = ""
				Else
				strDescription =  Environment("Description1")
				
			End If
			On Error goto 0

			If  Lcase(Trim( DataTable.GetSheet("TestPlan").getParameter("TestCaseName")))="endofrow" or  Lcase(Trim( DataTable.GetSheet("TestPlan").getParameter("TestCaseName")))="eof" or  Lcase(Trim( DataTable.GetSheet("TestPlan").getParameter("TestCaseName")))="x"  then
				Exit For
			End If

				 Environment("WithoutTestData")=False

				strAppMapPath_Relative= Trim(DataTable.GetSheet("TestPlan").getParameter("AppMapPath"))
				strAppMapPath=Environment("vegaTestSuite") &  DataTable.GetSheet("TestPlan").getParameter("AppMapPath")
				Environment("TestCaseName1")=DataTable.GetSheet("TestPlan").GetParameter("TestCaseName")
				Environment("strTCSeverity")=LCase(Trim(DataTable.GetSheet("TestPlan").GetParameter("Severity_KeyWord")))
				Call Header(Environment("EventLogPath"), LCase(Trim(Environment("EventLog")))) 
				Call Header(Environment("ScriptLogPath"), LCase(Trim(Environment("ScriptLog")))) 

					'TAF 10.1 new code Start
					If Environment("BlnCodeGeneration")=1 Then
						Call UpdateScript("'******************"&"TestCase:  "&Environment("TestCaseName1")&"   start"&"**************************")			'10.1
						Call UpdateScript("")
					End If
					'TAF 10.1 new code End
					


					'Modified by Trao to make appmap not mandatory------------TAF10
				If strTPMasterAppMap_Relative="" and  strAppMapPath_Relative="" Then
					Environment("AppMapCount")=0
					Environment("AppMapPath")=""
				Else

			' Modified by Srikanth to access multiple AppMap's--------------TAF10
				If VerifyAndDownloadAppMap(strAppMapPath)=0 then
					If  VerifyAndDownloadAppMap(strTPMasterAppMap)=0 then
						blnFlagSkipTestCase=true
						If Environment("strTCSeverity")="showstopper" Then
							FlagIsShowStopper=True
						End If
					else
						Reporter.ReportEvent micWarning,"Verifying APPMAP for "& Environment("TestCaseName1")&" Test Case","Copied Default Test Plan level APPMAP i.e. "& strTPMasterAppMap     
						WriteToEvent("Warning" & vbtab & "Copied Default Test Plan level APPMAP i.e. "& strTPMasterAppMap &" for Test Case " & Environment("TestCaseName1") )
						strAppMapPath=strTPMasterAppMap
					end if
				else
					strAppMapPath=strAppMapPath
				End If
				' Modification End ------------------------TAF10
			 End If
			'Trao modification end-----------TAF10
			
				strTCFilePath=Environment("vegaTestSuite") &  DataTable.GetSheet("TestPlan").getParameter("TestCaseFilePath")
				If not blnFlagSkipTestCase Then
					If VerifyFileExists(strTCFilePath)=0 Then
						If VerifyFileExists(strTPMasterTestCase)=0 Then
							blnFlagNoTestCase=true
                            If Environment("strTCSeverity")="showstopper" Then
								FlagIsShowStopper=True
							End If
						else
							Reporter.ReportEvent micWarning,"Verifying TestCase file for "&Environment("TestCaseName1")&" test case","Copied Default Test Plan level Test Case"     
							WriteToEvent("Warning" & vbtab & "Copied Default Test Plan level Test Case file i.e. "& strTPMasterTestCase &" for Test Case " & Environment("TestCaseName1"))
							strTCFilePath=strTPMasterTestCase
						End If
					else
						strTCFilePath=Environment("vegaTestSuite") &  DataTable.GetSheet("TestPlan").getParameter("TestCaseFilePath")
					End if
				End if
				Environment("TestCasePath")=strTCFilePath
'				strTestDataFilePath=DataTable.GetSheet("TestPlan").GetParameter("TestDataPath")'' Verifying for the existance of Test Data file if does not exists copying from the default level
				strTestDataFilePath1_Relative=Trim(DataTable.GetSheet("TestPlan").GetParameter("TestDataPath"))  
                strTestDataFilePath1=Environment("vegaTestSuite") &  DataTable.GetSheet("TestPlan").GetParameter("TestDataPath")   'Raj
				'strTestDataFilePath = Replace(lcase(Trim(strTestDataFilePath1)),"%client%",Lcase(Trim(Environment("Client"))))  'Raj
				strTestDataFilePath = GetDynamicParameter(strTestDataFilePath1)


				'Modified by Trao to make test data path not mandatory------------TAF10
				If  strTestDataFilePath1_Relative="" and strTPMasterTestData1_Relative="" Then
					Environment("WithoutTestData")=True
					 Environment("strTestDataFilePath")=""
				Else

				       If blnFlagSkipTestCase=false and blnFlagNoTestCase=false Then
							If verifyFileExists(strTestDataFilePath)=0 then
								If verifyFileExists(strTPMasterTestData)= 0 Then
									blnNoTDExists=true
								Else
									Reporter.ReportEvent micWarning,"Verifying Test Data  file for "& Environment("TestCaseName1")&"  Test Case","Copied Default Test Plan level Test Data file"     
									WriteToEvent("Warning" & vbtab & "Copied Default Test Plan level Test Data file i.e. "& strTPMasterTestData &" for Test Case " & Environment("TestCaseName1"))
									strTestDataFilePath = strTPMasterTestData
								End If
							End if	
						End if
                      Environment("strTestDataFilePath")=strTestDataFilePath
			End If
			'End of Trao modification------------TAF10

            
				If  blnFlagSkipTestCase then   ''Execute the test case when AppMap exists 
					reporter.ReportEvent micFail,"Verifying the  AppMap File"&Environment("AppMapPath")&" for test case "& Environment("TestCaseName1"), "AppMap File doesn't exists"		
					WriteToEvent("Fail" & vbtab &" AppMap File "&Environment("AppMapPath")&" does not exist")
				Else
					If blnFlagNoTestCase=true Then
						reporter.ReportEvent micFail, "Verifying for the Test Case " & Environment("TestCasePath") , " Test case "& Environment("TestCasePath") & " file does not exist"
						WriteToEvent("Fail" & vbtab & Environment("TestCasePath") & " file does not exist")
					else
						'Individual Test Case 
						Environment("TSorTC")=False
						Environment("TestCaseCounter")=0
						If  verifysheetexists(environment("TestCasePath"), Environment("TestCaseName1"))  Then				'Santosh 10 Sept
							'Call ExecuteTestCase()																																				'Santosh 10 Sept
							'Call ExecuteTestCase
							'  Changes for DLL
If Environment("BPT") Then
	Datatable.AddSheet("Global")
	DataTable.GetSheet("Global").AddParameter "TestCaseSheetName",""
	DataTable.GetSheet("Global").AddParameter "TestCaseFilePath",""
	DataTable.GetSheet("Global").AddParameter "TestDataDBPath",""
End If
'DataTable.GetSheet("Global").AddParameter "Description",""  ' Argus
DataTable("TestCaseSheetName","Global")=DataTable.GetSheet("TestPlan").GetParameter("TestCaseName")     
Environment("TestCaseSheetNameToResults") = DataTable("TestCaseSheetName","Global")  ' Argus
DataTable("TestCaseFilePath","Global")=DataTable.GetSheet("TestPlan").GetParameter("TestCaseFilePath")		
Datatable("TestCaseFilePath","Global")=Datatable.GetSheet("TestPlan").GetParameter("TestCaseFilePath")
Datatable("TestDataDBPath", "Global") = Environment("strTestDataFilePath")
Environment("TestDataPath") = Environment("strTestDataFilePath")
TempTestDataSheetName=Datatable.GetSheet("TestPlan").GetParameter("TestDataSheetName")
Environment("TestDataSheetName") = TempTestDataSheetName

'set below line putput to environment variable
TempDrefDataRow_Keyword=Datatable.GetSheet("TestPlan").GetParameter("DataRow_Keyword")
Environment("DrefDataRow_Keyword")=TempDrefDataRow_Keyword
TempDrefDataDriven_KeyWord=Datatable.GetSheet("TestPlan").GetParameter("DataDriven_KeyWord")
Environment("DrefDataDriven_KeyWord")=TempDrefDataDriven_KeyWord
'msgbox Environment("DrefDataDriven_KeyWord")

Environment("blnNoTDExists")=blnNoTDExists
'msgbox "boolvalue is "& blnNoTDExists
'										Set o1=createobject("TafCore.CoreEngine")
'										If Environment("BPT") Then
'                                   		o1.ExecuteTestCaseBPT()
'                                   	Else
'                                   		o1.ExecuteTestCase()
'                                   	End If
										ExecuteTestCase() ' Added by Srikanth
										If Environment("Execute") Then
											Call startTest()
										End If
										
										call PerformExecuteTest()

										'TAF 10.1 new code Start
										If Environment("BlnCodeGeneration")=1 Then

												If Not Trim(Lcase(Environment("ExcelResults"))) ="no"  Then
														Call UpdateScript("objResWorkBook.Save")         '10.1
														Call UpdateScript("appExcel.Quit")
														Call UpdateScript("'******************"&"TestCase:  "&Environment("TestCaseName1")&"   end"&"**************************")
														Call UpdateScript("")
												End If
											
										End If
										'TAF 10.1 new code End
										
'  end of changes		Call UpdateScript("")
						Else                                                                                                                                                                                     'Santosh 10 Sept
							Reporter.ReportEvent micFail,"Verifying for the Test Case in the Test Case file","Test Case "&Environment("TestCaseName1")&" Doesnot exists in the test case file"     'Santosh 10 Sept
							WriteToEvent("Failed" & vbtab &"Test Case "&Environment("TestCaseName1")&" Doesnot exists in the test case file")     'Santosh 10 Sept
						End If	
					End if
				End if   ''Test case execution completes
		else
			ScenarioCounter=ScenarioCounter+1
		End if
		CurrentScenario=strNewScenario
		If Environment("ScenerioEnd") OR not(Environment("TSorTC")) Then
			Exit For
		End If
	Next
    If Trim(Lcase(Environment("ExcelResults"))) <> "no"  and Trim(Lcase(Environment("IsSystemSlow")))<> "yes" Then	

		'	Select the Result Sheet
'			Set objWorkBook = appExcel.Workbooks.Open (Environment("ResultFilePath"))
			Set objSheet = appExcel.Sheets("Detailed_Report")
			appExcel.Sheets("Detailed_Report").Select
			Row=Environment("IncValHolder")
			objSheet.Range("A" & Row & ":H" & Row).Interior.ColorIndex =53 		'changed F to H Fanweb 7.0.11
			objSheet.Range("A" & Row).value="Test Execution End"
			objSheet.Range("A" & Row & ":H" & Row).Font.ColorIndex = 19				'changed F to H Fanweb 7.0.11
			objSheet.Range("A" & Row).Font.Bold = True   			
			'Save the Workbook
			Set objSheet = appExcel.Sheets("Test_Summary")
			appExcel.Sheets("Test_Summary").Select
			
'				If	Trim(Lcase(Environment("IsSummaryResult"))) = "yes" Then
'			 If not(blnIsScenario) Then
'				objSheet.Range("C13").Value=objSheet.Range("C13").Value +objSheet.Range("C10").Value ' Sum the execution 
'				objSheet.Range("C14").Value=objSheet.Range("C14").Value +objSheet.Range("C11").Value ' Sum the execution 
'				objSheet.Range("C15").Value=objSheet.Range("C15").Value +objSheet.Range("C12").Value ' Sum the execution 
'				objSheet.Rows("10:12").Delete
'			 else
'				objSheet.Range("C13").Value=objSheet.Range("C13").Value +objSheet.Range("C10").Value ' Sum the execution 
'				objSheet.Range("C14").Value=objSheet.Range("C14").Value +objSheet.Range("C11").Value-1 ' Sum the execution 
'				objSheet.Range("C15").Value=objSheet.Range("C15").Value +objSheet.Range("C12").Value ' Sum the execution 
'				objSheet.Rows("10:12").Delete
'			 End If
'			Else 
'			 If not(blnIsScenario) Then
'				objSheet.Range("C13").Value=objSheet.Range("C13").Value +objSheet.Range("C10").Value ' Sum the execution 
'				objSheet.Range("C14").Value=objSheet.Range("C14").Value +objSheet.Range("C11").Value ' Sum the execution 
'				objSheet.Range("C15").Value=objSheet.Range("C15").Value +objSheet.Range("C12").Value ' Sum the execution 
'				objSheet.Rows("10:12").Delete
'			 else
'				objSheet.Range("C10").Value=objSheet.Range("C13").Value +objSheet.Range("C10").Value ' Sum the execution 
'				objSheet.Range("C11").Value=objSheet.Range("C14").Value +objSheet.Range("C11").Value-1 ' Sum the execution 
'				objSheet.Range("C12").Value=objSheet.Range("C15").Value +objSheet.Range("C12").Value ' Sum the execution 
'				objSheet.Rows("13:15").Delete
'			 End If
'			
'			End If 		
			
'			
'			If not(blnIsScenario) Then
''				objSheet.Range("C13").Value=objSheet.Range("C13").Value +objSheet.Range("C10").Value ' Sum the execution 
''				objSheet.Range("C14").Value=objSheet.Range("C14").Value +objSheet.Range("C11").Value ' Sum the execution 
''				objSheet.Range("C15").Value=objSheet.Range("C15").Value +objSheet.Range("C12").Value ' Sum the execution 
''				objSheet.Rows("10:12").Delete
'
'				appExcel.Sheets("Test_Summary").Activate  
'				appExcel.Sheets("Test_Summary").Range("J2:J20000").Copy  
'				appExcel.Sheets("Test_Summary").Range("A1").PasteSpecial  
'
'			else
''				objSheet.Range("C10").Value=objSheet.Range("C13").Value +objSheet.Range("C10").Value ' Sum the execution 
''				objSheet.Range("C11").Value=objSheet.Range("C14").Value +objSheet.Range("C11").Value ' Sum the execution 
''				objSheet.Range("C12").Value=objSheet.Range("C15").Value +objSheet.Range("C12").Value ' Sum the execution 
''				objSheet.Rows("13:15").Delete
'
''				objSheet.Rows("16:16").Delete
''		        objSheet.Rows("19:19").Delete
'
'				appExcel.Sheets("Test_Summary").Activate  
'				appExcel.Sheets("Test_Summary").Range("J2:J20000").Copy  
'				appExcel.Sheets("Test_Summary").Range("A1").PasteSpecial  
'
'			End If

		If Trim(Lcase(Environment("IsSummaryResult"))) <> "yes"   Then
				objSheet.Range("D18").Value = objSheet.Range("C16").Value
		End If
				appExcel.Sheets("Test_Summary").Activate  
				appExcel.Sheets("Test_Summary").Range("J2:J20000").Copy  
				appExcel.Sheets("Test_Summary").Range("A1").PasteSpecial  
			
'			objSheet.Range("E15").Value = strDescription  ' Argus
'			objSheet.Range("E15").Font.ColorIndex = 12
'			objSheet.Range("E15").Font.Bold = True

			objSheet.Range("E18").Value = strDescription  ' Argus
			objSheet.Range("E18").Font.ColorIndex = 12
			objSheet.Range("E18").Font.Bold = True
'''            Call ProtectResults(objSheet)
			objResWorkBook.Save
'			 objWorkBook.Save
'		appExcel.Quit

			If Trim(Lcase(Environment("IsSummaryResult")))  = "yes"   Then
				appExcel.Quit
			Else
				If  Trim(LCase(Environment("AutoPopupExcelResults")))= "yes"  Then
					appExcel.visible=True
					Else
					appExcel.Quit
				End IF
			End If


            ' TAF10 Start - Closing excel process after finish
			If Trim(Lcase(Environment("IsSummaryResult")))  = "yes"   Then
				appExcel.Quit
				Set appExcel = Nothing
			else
				
					If   Trim(LCase(Environment("AutoPopupExcelResults")))= "yes"  Then
						appExcel.visible=True
					Else
						appExcel.Quit
						Set appExcel = Nothing	
					End If
			End If
   '	 	TAF10  End

'$$$$$$$$$$$$$$$$$$$$$$$$$$$ -- END MODIFICATION -- $$$$$$$$$$$$$$$$$$$$$$$$$$$$$		
End IF


If Trim(LCase(Environment("HTMLResults")))<>"no" Then
	Call ReportHTMLResults(strUIName, strExpected, strData, strActual, Status,objParent)
						If Trim(LCase(Environment("AutoPopupHTMLResults")))= "yes" And  Trim(Lcase(Environment("IsSummaryResult")))<> "yes"   Then
								Set objFileSystemObject = CreateObject("Scripting.FileSystemObject")
								Set objTextFileObject = objFileSystemObject.CreateTextFile(Environment("CurrentResultFile"),True)
					'		Set objTextFileObject = objFileSystemObject.CreateTextFile(REplace(Environment("CurrentResultFile"),".html",".txt"),True)
								HTMLSummary()
								objTextFileObject.Write HTMLContent
					'objFileSystemObject.CopyFile REplace(Environment("CurrentResultFile"),".html",".txt"),Environment("CurrentResultFile")
								Set IE=createobject("InternetExplorer.Application")
								IE.visible=True
								IE.navigate(Environment("CurrentResultFile"))
								Set objFileSystemObject=Nothing
						Else
								Set objFileSystemObject = CreateObject("Scripting.FileSystemObject")
								Set objTextFileObject = objFileSystemObject.CreateTextFile(Environment("CurrentResultFile"),True)
								HTMLSummary()
					
								objTextFileObject.Write HTMLContent
					'			Set IE=createobject("InternetExplorer.Application")
					'			IE.visible=True
					'			IE.navigate(Environment("CurrentResultFile"))
								Set objFileSystemObject=Nothing
						End If
End IF
									'TAF 10.1 new code Start
									If Environment("BlnCodeGeneration")=1 Then           '10.1

												If Not Trim(Lcase(Environment("HTMLResults"))) ="no"  Then
														Call UpdateScript("SaveHTML()")
														Call UpdateScript("'******************"&"TestCase:  "&Environment("TestCaseName1")&"   end"&"**************************")
														Call UpdateScript("")
												End If
											
										End If
										'TAF 10.1 new code Start

  
End Function

Public Function DownloadAllAttachments(TDFolderPath, sDownloadTo)
If Environment("IsSuiteFromTestPlan") Then   ' From Test Plan Tab

			Dim otaAttachmentFactory 'As TDAPIOLELib.AttachmentFactory
			Dim otaAttachment 'As TDAPIOLELib.Attachment
			Dim otaAttachmentList 'As TDAPIOLELib.List
			Dim otaAttachmentFilter 'As TDAPIOLELib.TDFilter
			Dim otaTreeManager 'As TDAPIOLELib.TreeManager
			Dim otaSysTreeNode 'As TDAPIOLELib.SysTreeNode
			Dim otaExtendedStorage 'As TDAPIOLELib.TreeManager
			Dim fso1
			Set fso1= CreateObject("Scripting.FileSystemObject")
			Dim strPath 'As String
			Set otaTreeManager = QCUtil.TDConnection.TreeManager
			Set otaSysTreeNode = otaTreeManager.NodeByPath(TDFolderPath)

			If otaSysTreeNode Is Nothing Then
					Reporter.ReportEvent micWarning,"Failure in library function 'DownloadAttachment'", "Failed to find Folder '" & TDFolderPath & "'."
			End If

			Set otaAttachmentFactory = otaSysTreeNode.Attachments
			Set otaAttachmentFilter = otaAttachmentFactory.Filter
			otaAttachmentFilter.Filter("CR_REFERENCE") = "'ALL_LISTS_" & otaSysTreeNode.NodeID & "_*'"
			Set otaAttachmentList = otaAttachmentFilter.NewList
			DowloadAttachments = ""
			If otaAttachmentList.Count > 0 Then
			For i = 1 to otaAttachmentList.Count
			set otaAttachment = otaAttachmentList.Item(i)
			otaAttachment.Load True, ""
			If (fso1.FileExists(otaAttachment.FileName)) Then
			strFile = otaAttachmentList.Item(i).Name
			myarray = split(strFile,"ALL_LISTS_" & otaSysTreeNode.NodeID & "_")
			fso1.CopyFile otaAttachment.FileName, sDownloadTo & "\" & myarray(1)
			Reporter.ReportEvent micPass, "File Download:", myarray(1) & " downloaded to " & sDownloadTo
			DownloadAttachments = sDownloadTo
			end if
			Next
			Else
			Reporter.ReportEvent micDone,"No attachments to download", _
			"No attachments found in specified folder '" & TDFolderPath & "'."
			DowloadAttachments = "Empty"
			End If
			Set otaAttachmentFactory = Nothing
			Set otaAttachment = Nothing
			Set otaAttachmentList = Nothing
			Set otaAttachmentFilter = Nothing
			Set otaTreeManager = Nothing
			Set otaSysTreeNode = Nothing
			Set fso1 = nothing

	Else     ' From Tes resources tab
			' 7.0.8B
			arrAttachPath = Split(TDFolderPath,"\")
			TDFolderName = arrAttachPath(UBound(arrAttachPath))
			
			Set oQC = QCUtil.QCConnection
			Set oRes = oQC.QCResourceFolderFactory
			Set oFilter =  oRes.Filter
			oFilter.Filter("RFO_NAME") =  TDFolderName
			Set oFileList = oFilter.NewList
			For  Each oFilee  in oFileList
				Set oFilterr =  oFilee.QCResourceFactory
				Set oFilte = oFilterr.Filter
				oFilte.Filter("RSC_NAME") = "*.*"
				Set oFileLis = oFilte.NewList
				For Each oFil in  oFileLis
					sName = oFil.FileName 
					oFil.FileName = sName
					oFil.DownloadResource sDownloadTo, True
				Next
			Next			
			Set oRes = Nothing
			Set oFilter =  Nothing
			Set oFilterr =  Nothing
			Set oFilte = Nothing
			Set oFileLis = Nothing
			' 7.0.8B
	End If

  
End Function


Public Function DownloadAttachment(TDAttachmentName, TDFolderPath,strTargerFolder, bMovetoRoot)

	If Environment("IsSuiteFromTestPlan") Then  ' From Test Plan Tab
	
				Dim otaAttachmentFactory '‘As TDAPIOLELib.AttachmentFactory
				Dim otaAttachment' ‘As TDAPIOLELib.Attachment
				Dim otaAttachmentList' ‘As TDAPIOLELib.List
				Dim otaAttachmentFilter '‘As TDAPIOLELib.TDFilter
				Dim otaTreeManager '‘As TDAPIOLELib.TreeManager
				Dim otaSysTreeNode' ‘As TDAPIOLELib.SysTreeNode
				Dim otaExtendedStorage' ‘As TDAPIOLELib.TreeManager
				Dim fso1
				Set fso1 = CreateObject("Scripting.FileSystemObject")
				Dim strPath' ‘As String
				Set otaTreeManager = QCUtil.TDConnection.TreeManager
				Set otaSysTreeNode = otaTreeManager.NodeByPath(TDFolderPath)

				If otaSysTreeNode Is Nothing Then
					Reporter.ReportEvent micWarning,"Failure in library function 'DownloadAttachment'", "Failed to find Folder '" & TDFolderPath & "'."
				End If

				Set otaAttachmentFactory = otaSysTreeNode.Attachments
				Set otaAttachmentFilter = otaAttachmentFactory.Filter
				otaAttachmentFilter.Filter("CR_REFERENCE") = "ALL_LISTS_" & otaSysTreeNode.NodeID & "_" & TDAttachmentName & ""
				Set otaAttachmentList = otaAttachmentFilter.NewList
				If otaAttachmentList.Count > 0 Then
				set otaAttachment = otaAttachmentList.Item(1)
				otaAttachment.Load True,"'"
				strPath = otaAttachment.FileName
				if(bMovetoRoot) then
				If (fso1.FileExists(otaAttachment.FileName)) Then
				fso1.CopyFile strPath, strTargerFolder & "\" & TDAttachmentName
				DownloadAttachment = strTargerFolder & "\" & TDAttachmentName
				else
				DownloadAttachment = strPath
				end if
				else
				DownloadAttachment = strPath
				end if
				Else
				'msgbox "failed"
				Reporter.ReportEvent micFail,"Failure in library function 'DownloadAttachment'", "Failed to find attachment '" & TDAttachmentName & "' in folder'" & TDFolderPath & "'."
				DowloadAttachment = "Empty"
				End If
				Set otaAttachmentFactory = Nothing
				Set otaAttachment = Nothing
				Set otaAttachmentList = Nothing
				Set otaAttachmentFilter = Nothing
				Set otaTreeManager = Nothing
				Set otaSysTreeNode = Nothing
				Set fso1 = nothing
		Else    ' From Tes resources tab
				' 7.0.8B
				arrAttachPath = Split(TDFolderPath,"\")
				TDFolderName = arrAttachPath(UBound(arrAttachPath))

				Set oQC = QCUtil.QCConnection
				Set oRes = oQC.QCResourceFolderFactory
				Set oFilter =  oRes.Filter
				oFilter.Filter("RFO_NAME") =  TDFolderName
				Set oFileList = oFilter.NewList
				
				For  Each oFilee  in oFileList
					Set oFilterr =  oFilee.QCResourceFactory
					Set oFilte = oFilterr.Filter
					oFilte.Filter("RSC_NAME") = TDAttachmentName
					Set oFileLis = oFilte.NewList
					sName = TDAttachmentName
					If  oFileLis.Count > 0 Then
						Set oFil = oFileLis.Item(1)
						oFil.FileName = sName
						oFil.DownloadResource strTargerFolder, True
						Exit For
					End If
				Next
				' 7.0.8B
		End If
		 		 
End Function


Public function AttachFileToQC(FromFileLoc,FromFileName,ToQCTestPath,ToQCTestName) 

	Dim FromFile, Qc, tm, root, fold 
	AttachFileToQC ="Failed to attach file for unknown reason" 
	FromFile=FromFileLoc & FromFileName 
	Set fso1 = CreateObject("Scripting.FileSystemObject") 
	status=fso1.FileExists(FromFile) 
	If status=False Then 
		AttachFileToQC="Failed to attach file because file does not exist" 
		Set fso1 = Nothing 
Exit Function 
	End If 
	Set Qc = QCUtil.TDconnection 'if it fails here, you're not connected to QC 
	Set tm = Qc.TreeManager 
	Set root=tm.TreeRoot("Subject") 'Connects to Subject 
	Set fold = tm.NodeByPath( "Subject\" & ToQCTestPath)'Finds folder, if it fails here, you're either pointing to a bad QC folder or to a test in a folder 
	Set testList = fold.FindTests(ToQCTestName) 
	If(testList.Count =0) then 
		AttachFileToQC="Failed to find any test in folder." 
		Set Qc = Nothing 
		Set tm = Nothing 
		Set root = Nothing 
		Set fold = Nothing 
		Set testList = Nothing 
	End If 
	For i = 1 To testList.Count 
		Set vTest = testList.Item(i) 
		If LCase(vTest.Name) = LCase(ToQCTestName) Then 'searchs each test name for requested test. This is not case sensative. 
			Set attf = vTest.Attachments 
			Set att = attf.AddItem(Null) 
			att.FileName = FromFile 
			att.Type = 1 
			att.Post 
			att.Save False 
			AttachFileToQC= "The file " & FromFileName & " was uploaded on the following path " & ToQCTestPath & " to the test " & ToQCTestName & "," 
			Set Qc = Nothing 
			Set tm = Nothing 
			Set root = Nothing 
			Set fold = Nothing 
			Set testList = Nothing 
			Set attf = Nothing 
			Set att = Nothing 
			Set vTest = Nothing 
Exit Function 
		End If 
	Next 
	Set Qc = Nothing 
	Set tm = Nothing 
	Set root = Nothing 
	Set fold = Nothing 
	Set testList = Nothing 
	Set vTest = Nothing
	
  
End Function 




'  New implementation DSTGS
 Public Function AttachFileToCurrentTestSetTest(strFiletoattach)
	
		Set fso  = CreateObject("Scripting.FileSystemObject") 
		status=fso.FileExists(strFiletoattach) 
		If status=False Then 
			AttachFileToCurrentTestSetTest ="Failed to attach file because file does not exist" 
			Set fso = Nothing 
			Exit Function 
		End If 
		Set o_CurrentRun=QCUtil.CurrentRun
		' Check that we are running this test from QC, otherwise we can exit
		If (o_CurrentRun Is Nothing) Then
			Exit Function
		End If
		Set CTA=o_CurrentRun.Attachments
		Set att1 = CTA.AddItem(null) 
		att1.fileName=strFiletoattach
		att1.type=1
        att1.post
		att1.save False
			
  
 End Function

'To close Excel process
Public Function fnCloseApplication( byval sApplicationExe)
Dim strComputer
Dim objWMIService
Dim colProcesses
Dim objProcess
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colProcesses = objWMIService.ExecQuery ("Select * from Win32_Process Where Name = '"&sApplicationExe&"'")
For Each objProcess in colProcesses
objProcess.Terminate()

Next
Set objWMIService = Nothing
Set colProcesses=Nothing
End Function

Public Function LoadFromTempLS()
   On error resume next
   If VerifyEnvVariable("WorkingDirectory") Then
	   If Lcase(Environment("WorkingDirectory"))<>"n/a" Then
		   	If   fso.FolderExists(Environment("WorkingDirectory") & "\TempAutomation")Then
                	fso.DeleteFolder(Environment("WorkingDirectory") & "\TempAutomation")
            End If
			strPath = Replace(Environment("WorkingDirectory") ,"/","\")
            arr = Split(Environment("WorkingDirectory") ,"\")
			If ubound(arr) >  0 Then
				strkey =  arr(0) 
                For i = lbound(arr) to ubound(arr)-1
                    On error resume next
                    fso.CreateFolder strkey & "\" & arr(i+1)
                    strkey = strkey &"\" & arr(i+1)
                    On error goto 0
                Next
            End If
			fso.CreateFolder(Environment("WorkingDirectory") & "\TempAutomation")
			Environment("TempLS")=Environment("WorkingDirectory") & "\TempAutomation"

			LoadFromTempLS=TRUE
			Exit Function
		   
		Else
			VerifyEnvVariable=FALSE
			Exit Function
	   End If	 
	Else
			LoadFromTempLS=FALSE	
   End If
End Function


Function GetEOFRowID(strFileName,strSQLStatement)
   Dim objRecord
   Dim strRowValue, strOriginalValue
		GetEOFRowID = -1
		Set objConnection = CreateObject("ADODB.Connection")
		Set objRecordSet = CreateObject("ADODB.Recordset")
		objConnection.Open Environment("strConnectionString")& strFileName & ";Readonly=True"
		objRecordSet.CursorLocation=3 ' set the cursor to use adUseClient – disconnected recordset
		objRecordSet.Open strSQLStatement, objConnection, 1, 3
		If objRecordSet.EOF <> True Then
			Set objRecord = objRecordSet.Fields(0)
			strRowValue = objRecord.Value
            strOriginalValue = objRecord.OriginalValue
			If IsNull(strOriginalValue) Then			
				GetEOFRowID = "Null"				
			Else
				GetEOFRowID = strRowValue
			End If
		Else
			GetEOFRowID = -1
		End If
	Set objRecordSet.ActiveConnection = Nothing
	Set objConnection = Nothing
End Function



Function obj_getEntireRecordset (ByVal strFileName, ByVal strSQLStatement, ByVal strRowNum)

Dim objConnection, objRecordSet
Dim objRecord
Dim intTemp
Dim strRowValue

Set objConnection = CreateObject("ADODB.Connection")

objConnection.Open Environment("strConnectionString")& strFileName & ";Readonly=True"
If Err.Number <> 0 Then
	Reporter.ReportEvent micFail,"Create Connection", "[Connection] Error has occured. Error : " & Err.Number
	Set obj_getEntireRecordset = Nothing
	Exit Function
End If
Set objRecordSet = CreateObject("ADODB.Recordset")
If strRowNum <> -1 Then
	objRecordSet.CursorLocation=3 ' set the cursor to use adUseClient – disconnected recordset
	objRecordSet.Open strSQLStatement, objConnection, 1, 3
	If objRecordSet.EOF <> True Then
		Set objDict = objDictObjCreation()
		For intTemp = 1 to objDict.Count'objRecordSet.Fields.Count-1
			Set objRecord = objRecordSet.Fields(intTemp)
			strRowValue = objRecord.Value
			If strRowValue <> "" Then			
				Call AddValuesToXL(intTemp, strRowValue,"TestPlan", strRowNum)
			End If
		Next
	End If
End If

If Err.Number<>0 Then
	Reporter.ReportEvent micFail,"Open Recordset", "Error has occured.Error Code : " & Err.Number
	Set obj_getEntireRecordset = Nothing
	Exit Function
End If

Set objRecordSet.ActiveConnection = Nothing
Set objConnection = Nothing

End Function


Function AddParametersToExcel()
   Dim strXLTestPlan
   Dim objDictParameterName
   Dim intTemp

	Set objDictParameterName = objDictObjCreation()
    strXLTestPlan = "TestPlan"
	DataTable.AddSheet strXLTestPlan
	For intTemp = 1 to objDictParameterName.Count
		DataTable.GetSheet(strXLTestPlan).AddParameter objDictParameterName.Item(intTemp), ""
	Next
	Set objDictParameterName = Nothing
End Function

Function AddFirstRowValues(strFileName)
   Dim strQuery
	strQuery = "SELECT * from TestPlan  Where ID = 1"
	Call obj_getEntireRecordset(strFileName, strQuery, 1)
End Function


Function AddValuesToXL(intTemp, strRowValue, strSheetName, intRowNumber)
	Set objDictParameterName = objDictObjCreation()
    DataTable.GetSheet(strSheetName).GetParameter(objDictParameterName.Item(intTemp)).ValueByRow(intRowNumber) = strRowValue
	Set objDictParameterName = Nothing
End Function


Function objDictObjCreation()
	Dim objDictParameterName
    Dim objConnection, objRecordSet
	Dim objRecord
	Dim intTemp
	Dim strRowValue

Set objDictParameterName = CreateObject("Scripting.Dictionary")
strSQLStatement = "SELECT * from TestPlan  Where ID = 1"
strFileName = Environment("strFileName")
Set objConnection = CreateObject("ADODB.Connection")

objConnection.Open Environment("strConnectionString")& strFileName & ";Readonly=True"
If Err.Number <> 0 Then
	Reporter.ReportEvent micFail,"Create Connection", "[Connection] Error has occured. Error : " & Err.Number
	Set obj_getEntireRecordset = Nothing
	Exit Function
End If
Set objRecordSet = CreateObject("ADODB.Recordset")
If strRowNum <> -1 Then
	objRecordSet.CursorLocation=3 ' set the cursor to use adUseClient – disconnected recordset
	objRecordSet.Open strSQLStatement, objConnection, 1, 3
	If objRecordSet.EOF <> True Then		
		For intTemp = 1 to objRecordSet.Fields.Count-1
			Set objRecord = objRecordSet.Fields(intTemp)
			strRowValue = objRecord.Name
			If strRowValue <> "" Then			
				objDictParameterName.Add intTemp, strRowValue
			End If
		Next
	End If
End If

		Set objDictObjCreation = objDictParameterName
End Function


 Function LoadDataTableFromDB(strFileName,strTestScriptName,blnIsScenario)
			If Lcase(Cstr(blnIsScenario)) = "true" Then
				strScenarioKeyword = "Scenario"
				else
				strScenarioKeyword = "TestCase"
			End If
			intRowCorrection = 2
			Environment("strFileName") = strFileName
			Environment("strConnectionString") = GetTheConnectionStringofDB
            		
			AddParametersToExcel
			AddFirstRowValues strFIleName

			strValue = GetEOFRowID(strFileName, "SELECT ID from TestPlan  Where TestCaseName = '"& strTestScriptName &"' and Scenario_Keyword = '" & strScenarioKeyword &"'" )
			
			intEOFRow = ""
			inTemp = strValue
			If inTemp <> -1 Then
				intEOFRow = GetEOFRowID(strFileName, "SELECT Scenario_Keyword from TestPlan  Where ID = " & inTemp)
				If LCase(Trim(intEOFRow)) = LCase("Scenario") Then
					Do Until UCase(intEOFRow) = UCase("EndOfScenario") Or UCase(intEOFRow) = UCase("EOF")						
						If inTemp <> strValue Then
							intEOFRow = GetEOFRowID(strFileName, "SELECT Scenario_Keyword from TestPlan  Where ID = " & inTemp)
						End If
						If intEOFRow <> -1 Then
							Call obj_getEntireRecordset(strFileName, "SELECT * from TestPlan  Where ID = "& inTemp, intRowCorrection)						
							intRowCorrection = intRowCorrection+1
						End If			
						inTemp = inTemp+1
					Loop
				Else
					Call obj_getEntireRecordset(strFileName, "SELECT * from TestPlan  Where ID = "& inTemp, intRowCorrection)	
				End If
			
			End If
			
			If inTemp = -1 Then
				Reporter.ReportEvent micFail, "Verifying the test script: '" & strTestScriptName & " '" ,"Script Not found in the controller , please verify"
			Else
'				DataTable.ExportSheet "C:\Documents and Settings\Taruni\Desktop\Prakash.xls", "TestPlan"
				Reporter.ReportEvent micDone, "Loading the test script: " & strTestScriptName &  " From DB" ,"Imported to Data table succesfully"
			End If
End Function

Function GetOfficeVersion()	
	Set objWMI = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & "." & "\root\CIMV2")	
	Set colItems = objWMI.ExecQuery("SELECT Version FROM Win32_Product WHERE Name Like 'Microsoft Office Access%'")
	
	If colItems.Count = 0 Then
		GetOfficeVersion = 0
		Exit Function
	End If
	
	For Each objItem In colItems
		GetOfficeVersion = Left(objItem.Version,InStr(1,objItem.Version,".")-1)		
		Exit Function
	Next
	
	Set objWMI = Nothing
	Set colItems = Nothing
	Set objWMI = Nothing

End Function

Function GetTheConnectionStringofDB()   
		Select Case GetOfficeVersion
			Case 9'Access 2000
				strConnectionString = "DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ="
			Case 10'Access 2002
				strConnectionString = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ="
			Case 11'Access 2003
				strConnectionString = "DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ="			
			Case 12'Access 2007
				strConnectionString = "DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ="			
			Case 14'Access 2010
				strConnectionString = "DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ="			
			Case 0'If Access is not installed
				strConnectionString = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ="
				Reporter.ReportEvent micInfo, "Microsoft Access Install" ,"Microsoft Access is NOT Install"
		End Select
		GetTheConnectionStringofDB = strConnectionString
End Function


 'Added by Srikanth to acess multiple AppMap files-   TAF10 change
Function VerifyAndDownloadAppMap(byval AppFilePath)
   VerifyAndDownloadAppMap = 1
	If Not Environment("RunFromQC")  Then
		tempAppFilePath = Split(AppFilePath,Environment("vegaTestSuite"))
		AppFilePath = tempAppFilePath(1)
	End If

   If Trim(AppFilePath) = ""  Then
        VerifyAndDownloadAppMap = 0	
   Else
		MultiAppMap = Split(AppFilePath,"!" )
		Environment("AppMapCount") = Ubound(MultiAppMap)+1
		Environment("AppMapPath") = ""
		Environment("QCAppMapPath") =""
		For appcount = 0 to Ubound(MultiAppMap)
				strAppMapPath = Environment("vegaTestSuite")&MultiAppMap(appcount)
				 arrTemp = Split(strAppMapPath,"\")
				 strAppMapFileName = arrTemp(ubound(arrTemp))
				 arrTemp = Split(strAppMapPath,"\" & strAppMapFileName)
				strAppMapFileFolder = arrTemp(lbound(arrTemp))

				If Environment("RunFromQC")  Then
					DownloadAttach = DownloadAttachment (strAppMapFileName,strAppMapFileFolder,Environment("TempAppMap"),"TRUE")
					If DownloadAttach = "Empty" Then
						VerifyAndDownloadAppMap = 0
						Exit For
					End If
					Environment("AppMapPath"&appcount+1)=  Environment("TempAppMap")&   "\" & strAppMapFileName 
					Environment("QCAppMapPath"&appcount+1)=  strAppMapPath
					If appcount >0 Then
						Environment("AppMapPath") = Environment("AppMapPath")&"!"&Environment("AppMapPath"&appcount+1)
						Environment("QCAppMapPath")  =Environment("QCAppMapPath") &"!"&Environment("QCAppMapPath"&appcount+1)
					Else
						Environment("AppMapPath") = Environment("AppMapPath"&appcount+1)
						Environment("QCAppMapPath")  =Environment("QCAppMapPath"&appcount+1)
					End If
                    					
				Else
				   Set fso=CreateObject("Scripting.FileSystemObject")
				   If Not fso.FileExists(strAppMapPath) Then
						VerifyAndDownloadAppMap=0
						Exit For
					End If
					Environment("AppMapPath"&appcount+1)=  strAppMapPath
					If appcount >0 Then
						Environment("AppMapPath") = Environment("AppMapPath")&"!"&Environment("AppMapPath"&appcount+1)
					Else
						Environment("AppMapPath") = Environment("AppMapPath"&appcount+1)
					End If                        
 				End If
		Next
	End If
End Function


'************************************ Modes **********************************************************


'********************************************************************************************************************************************************************************
' Function Name :I Mode
' Description   :This function "IMode" is Input mode on all the classes.
' param param1 - Parameter passed from framework
' param param2 - Parameter passed from framework
' param param3 - Parameter passed from framework
' Param Param4	- Parameter passed from framework
' Param Param5 - Parameter passed from framework
' Author	:	DSTWS TA2000 Automation Team
' Creation Date	: 	April, 2011
' Reviewed By 	:	DSTWS TA2000 Automation Team
' Modified By	:
' Modified Date	:	


'*********************************************************************************************************************************************************************************
 Function IMode(strParentObject,strChildobject,Param1,Param2,Param3,Param4,Param5,Param6,Param7,Param8,Param9,Param10,UIName,ExpectedResult)
    	set objParent=Eval(strParentObject)    
	set  objTestObject=Eval(strParentObject&"."&strChildobject)
	On Error Resume Next
	Err.Clear
	
If objTestObject.Exist(0) Then
	
	strClass = objTestObject.GetROProperty("Class Name")
	Select Case LCase(strClass)
	Case "winbutton"
		IMode= ClickOnButton(strParentObject,strChildobject,UIName,ExpectedResult)
	Case "webbutton"
		IMode= ClickOnButton(strParentObject,strChildobject,UIName,ExpectedResult)
	Case "link"
		IMode= ClickOnLink(strParentObject,strChildobject,UIName,ExpectedResult)
	Case "wincheckbox"
		Param1=FormatData(Param1)
		IMode= SelectCheckBox(strParentObject,strChildobject ,Param1,UIName,ExpectedResult)
	Case "webcheckbox"
		Param1=FormatData(Param1)
		IMode= SelectCheckBox(strParentObject,strChildobject ,Param1,UIName,ExpectedResult)
	Case "winlist"
		IMode= SelectValueInWinList(strParentObject,strChildobject,Param1,UIName,ExpectedResult)
	Case "weblist"
		Param1=FormatData(Param1)
		IMode= SelectValueInWebList(strParentObject,strChildobject,Param1,UIName,ExpectedResult)
	Case "winlistview"
		IMode= GridOperations(strParentObject,strChildobject,Param1,Param2,UIName,ExpectedResult)
	Case "wincombobox"
		IMode= SelectValueInCombo(strParentObject,strChildobject,Param1,UIName,ExpectedResult)
	Case "winradiobutton"
		Param1=FormatData(Param1)
		IMode= SelectRadioButton(strParentObject,strChildobject ,Param1,UIName , ExpectedResult)
	Case "webradiogroup"
		Param1=FormatData(Param1)
		IMode= SelectRadioButton(strParentObject,strChildobject ,Param1,UIname , ExpectedResult)
	Case "winedit"
		Param1=FormatData(Param1)
		IMode= SetTextOnEdit(strParentObject,strChildobject,Param1,UIName,ExpectedResult)
	Case "webedit"
		Param1=FormatData(Param1)
		IMode= SetTextOnEdit(strParentObject,strChildobject,Param1,UIName,ExpectedResult)
	Case "winobject"
		Param1=FormatData(Param1)
		IMode= SetTextOnEdit(strParentObject,strChildobject,Param1,UIName,ExpectedResult)
	Case "activex"
		strNatClass=objTestObject.GetROProperty("nativeclass")
        If Instr(1,LCase(strNatClass),"grid")>0 Then
        	Param1=FormatData(Param1)
        	Param2=FormatData(Param2)
        	Param3=FormatData(Param3)
			IMode=SetDataInActiveXGrid(strParentObject,strChildobject,Param1,Param2,Param3,UIName,ExpectedResult)
		Else
			Param1=FormatData(Param1)
			IMode= TypeInActiveX(strParentObject,strChildobject,Param1,UIName,ExpectedResult)
		End If
	Case "wineditor"
		Param1=FormatData(Param1)
		IMode= TypeTextOnEdit(strParentObject,strChildobject,Param1,UIName,ExpectedResult)
	Case "wintab"
	        Param1=FormatData(Param1)
		IMode= SelectWinTab(strParentObject,strChildobject,Param1,UIName,ExpectedResult)
	Case "image"
		IMode= ClickOnImage(strParentObject,strChildobject,UIName,ExpectedResult)
	Case "winmenu"
	        Param1=FormatData(Param1)
		IMode= SelectValueInWinMenu(strParentObject,strChildobject,Param1,UIName,ExpectedResult)
	Case "webtable"
		arrParam=Param1&";"&Param2&";"&Param3&";"&Param4&";"&Param5&";"&Param6&";"&Param7&";"&Param8&";"&Param9&";"&Param10
		arrParam=Split(arrParam,";")
		For i=UBound(arrParam) to 0 Step -1
			If arrParam(i)="" Or arrParam(i)=Empty Then
				ReDim Preserve arrParam(i-1)
			Else
				classParam=arrParam(i)
                Exit For
			End If
	
		Next
	
	Select Case Lcase(classParam)
	Case "image"
			RefData1=FormatData(Param1)
			RefData2=FormatData(Param2)
			colName=FormatData(Param3)
			RowIncrement=FormatData(Param4)
			emptyParam=FormatData(Param5)
			IMode=ClickOnImageInTable(strParentObject,strChildobject,RefData1,RefData2,colName,RowIncrement,emptyParam,UIName,strExpectedResult)
	Case "webbutton"
			RefData1=FormatData(Param1)
			RefData2=FormatData(Param2)
			colName=FormatData(Param3)
			RowIncrement=FormatData(Param4)
			emptyParam=FormatData(Param5)
			IMode=ClickOnButtonInTable(strParentObject,strChildobject,RefData1,RefData2,colName,RowIncrement,emptyParam,UIName,strExpectedResult)
	Case "link"
			RefData1=FormatData(Param1)
			RefData2=FormatData(Param2)
			colName=FormatData(Param3)
			RowIncrement=FormatData(Param4)
			emptyParam=FormatData(Param5)
			IMode=ClickOnLinkInTable(strParentObject,strChildobject,RefData1,RefData2,colName,RowIncrement,emptyParam,UIName,strExpectedResult)
	Case "webedit"
			RefData1=FormatData(Param1)
			RefData2=FormatData(Param2)
			colName=FormatData(Param3)
			RowIncrement=FormatData(Param4)
			strValuetoSet=FormatData(Param5)
			IMode=SetValueInEditBoxInTable(strParentObject,strChildobject,RefData1,RefData2,colName,RowIncrement,strValuetoSet,UIName,strExpectedResult)
        Case "webcheckbox"
			RefData1=FormatData(Param1)
			RefData2=FormatData(Param2)
			colName=FormatData(Param3)
			RowIncrement=FormatData(Param4)
			strValuetoSet=FormatData(Param5)
			IMode=SelectCheckBoxInTable(strParentObject,strChildobject,RefData1,RefData2,colName,RowIncrement,strValuetoSet,UIName,strExpectedResult)
	Case "webradiogroup"
			RefData1=FormatData(Param1)
			RefData2=FormatData(Param2)
			colName=FormatData(Param3)
			RowIncrement=FormatData(Param4)
			strValuetoSet=FormatData(Param5)
			IMode=SelectRadioGroupInTable(strParentObject,strChildobject,RefData1,RefData2,colName,RowIncrement,strValuetoSet,UIName,strExpectedResult)
	Case "weblist"
			RefData1=FormatData(Param1)
			RefData2=FormatData(Param2)
			colName=FormatData(Param3)
			RowIncrement=FormatData(Param4)
			strValuetoSet=FormatData(Param5)
			 IMode=SelectWebListItemInTable(strParentObject,strChildobject,RefData1,RefData2,colName,RowIncrement,strValuetoSet,UIName,strExpectedResult)
	
	Case Else 
		Reporter.ReportEvent micFail,"Operation to be performed on a Class","There is no ClassName specified in the parameters","Please Specify a Class Name"
		Call ReportResult (Environment("WindowName"),"TOperation to be performed on a Class",Environment("WindowName"), "There is no ClassName specified in the parameters , Please Specify a Class Name.","Failed",objParent)
		Exit Function
	End Select
	End Select
	Else
	Reporter.ReportEvent micFail,"There is no Window opened on the application with the specified window name (Or) There is some problem accessing the class of the given object , The Window Name You have given is->"&Environment("WindowName"),"Please verify and run again"
    Call ReportResult (Environment("WindowName"),"There should be a window opened in the application with the specified window name",Environment("WindowName"), "There is no Window opened on the application with the specified window name (Or) There is some problem accessing the class of the given object , The Window Name You have given is->"&Environment("WindowName")&" Please verify and run again.","Failed",objParent)
	Exit Function

End If	
 End Function

'********************************************************************************************************************************************************************************
' Function Name :SMode
' Description   :This function "S" is Save mode on all the classes.
' param param1 - Parameter passed from framework
' param param2 - Parameter passed from framework
' param param3 - Parameter passed from framework
' Param Param4	- Parameter passed from framework
' Param Param5 - Parameter passed from framework
' Author	:	DSTWS TA2000 Automation Team
' Creation Date	: 	April, 2011
' Reviewed By 	:	DSTWS TA2000 Automation Team
' Modified By	:
' Modified Date	:	
'*********************************************************************************************************************************************************************************
Function SMode(strParentObject,strChildobject,Param1,Param2,Param3,Param4,Param5,Param6,Param7,Param8,Param9,Param10,UIName,ExpectedResult)

	set objParent=Eval(strParentObject)    
	set  objTestObject=Eval(strParentObject&"."&strChildobject)
	On Error Resume Next
	err.Clear
	
If objTestObject.Exist(0) Then
	
	strClass = objTestObject.GetROProperty("Class Name")
	Select Case LCase(strClass)
	Case "wincheckbox"
		Param1=FormatData(Param1)
		SMode= GetCheckBoxSelection(strParentObject,strChildobject,UIName,ExpectedResult)
	Case "webcheckbox"
		Param1=FormatData(Param1)
		SMode= GetCheckBoxSelection(strParentObject,strChildobject,UIName,strExpectedResult)
	Case "winlist"
		Param1=FormatData(Param1)
		SMode=GetValueFromWinList(strParentObject,strChildobject,UIName,strExpectedResult)
	Case"weblist"
		Param1=FormatData(Param1)
		SMode=GetValueFromWebList(strParentObject,strChildobject,UIName,strExpectedResult)
	Case "winlistview"
		Param1=FormatData(Param1)
		Param2=FormatData(Param2)
		Param3=FormatData(Param3)
		SMode=GetDataInWinListView(strParentObject,strChildobject,Param1,Param2,Param3,UIName,ExpectedResult)
	Case "wincombobox"
		Param1=FormatData(Param1)
		SMode=GetDataFromWinComboBox(strParentObject,strChildobject,UIName,ExpectedResult)
	Case "winradiobutton"
		Param1=FormatData(Param1)
		SMode=GetRadioButtonSelection(strParentObject,strChildobject,UIName,ExpectedResult)
	Case "webradiogroup"
		Param1=FormatData(Param1)
		SMode=GetRadioButtonSelection(strParentObject,strChildobject,UIName,strExpectedResult)
	Case "winedit"
		Param1=FormatData(Param1)
		SMode= GetDataFromEditBox(strParentObject,strChildobject,UIName,ExpectedResult)
	Case "webedit"
		Param1=FormatData(Param1)
		SMode= GetValueInEditBox(strParentObject,strChildobject,UIName,strExpectedResult)
	Case "activex"
		strNatClass=objTestObject.GetROProperty("nativeclass")
        If Instr(1,LCase(strNatClass),"grid")>0 Then
			Param1=FormatData(Param1)
			Param2=FormatData(Param2)
			Param3=FormatData(Param3)
			SMode= GetDataInGrid(strParentObject,strChildobject,Param1,Param2,Param3,UIName,ExpectedResult)
		Else
			SMode= GetDataFromEditBox(strParentObject,strChildobject,UIName,ExpectedResult)
		End If
	Case "wineditor"
		SMode= GetDataFromEditBox(strParentObject,strChildobject,UIName,ExpectedResult)
	Case "static"
		SMode= GetStaticText(strParentObject,strChildobject,UIName,ExpectedResult)
	Case "winobject"
		SMode= GetDataFromEditBox(strParentObject,strChildobject,UIName,ExpectedResult)
    Case "webelement"
		Param1=FormatData(Param1)
		SMode= GetWebElementText(strParentObject,strChildobject,UIName,strExpectedResult)

	Case "webtable"
		arrParam=Param1&";"&Param2&";"&Param3&";"&Param4&";"&Param5&";"&Param6&";"&Param7&";"&Param8&";"&Param9&";"&Param10
		arrParam=Split(arrParam,";")
		For i=UBound(arrParam) to 0 Step -1
			If arrParam(i)="" Or arrParam(i)=Empty Then
				ReDim Preserve arrParam(i-1)
			Else
				classParam=arrParam(i)
                Exit For
			End If
	
		Next
	Select Case Lcase(classParam)

		Case "webelement"

			RefData1=FormatData(Param1)
			RefData2=FormatData(Param2)
			colName=FormatData(Param3)
			RowIncrement=FormatData(Param4)
			emptyParam=FormatData(Param5)
			SMode=GetWebElementDataInTable(strParentObject,strChildobject,Param1,Param2,Param3,Param4,Param5,UIName,strExpectedResult)


		Case "weblist"

			RefData1=FormatData(Param1)
			RefData2=FormatData(Param2)
			colName=FormatData(Param3)
			RowIncrement=FormatData(Param4)
			emptyParam=FormatData(Param5)
			SMode=GetWebListSelectionInWebTable(strParentObject,strChildobject,RefData1,RefData2,colName,RowIncrement,emptyParam,UIName,strExpectedResult)

		Case "webedit"

			RefData1=FormatData(Param1)
			RefData2=FormatData(Param2)
			colName=FormatData(Param3)
			RowIncrement=FormatData(Param4)
			emptyParam=FormatData(Param5)
			SMode=GetEditBoxValueInWebTable(strParentObject,strChildobject,RefData1,RefData2,colName,RowIncrement,emptyParam,UIName,strExpectedResult)

        Case "webcheckbox"

			RefData1=FormatData(Param1)
			RefData2=FormatData(Param2)
			colName=FormatData(Param3)
			RowIncrement=FormatData(Param4)
			emptyParam=FormatData(Param5)
			SMode=GetCheckBoxSelectionInTable(strParentObject,strChildobject,RefData1,RefData2,colName,RowIncrement,emptyParam,UIName,strExpectedResult)

		Case "webradiogroup"

			RefData1=FormatData(Param1)
			RefData2=FormatData(Param2)
			colName=FormatData(Param3)
			RowIncrement=FormatData(Param4)
			emptyParam=FormatData(Param5)
			SMode=GetRadioGroupSelectionInTable(strParentObject,strChildobject,RefData1,RefData2,colName,RowIncrement,emptyParam,UIName,strExpectedResult)

	Case Else 
		Reporter.ReportEvent micFail,"Operation to be performed on a Class","There is no ClassName specified in the parameters","Please Specify a Class Name"
		Call ReportResult (Environment("WindowName"),"TOperation to be performed on a Class",Environment("WindowName"), "There is no ClassName specified in the parameters , Please Specify a Class Name.","Failed",objParent)
		Exit Function
	End Select
    End Select
	Else
	Reporter.ReportEvent micFail,"There is no Window opened on the application with the specified window name (Or) There is some problem accessing the class of the given object , The Window Name You have given is->"&Environment("WindowName"),"Please verify and run again"
	Call ReportResult (Environment("WindowName"),"There should be a window opened in the application with the specified window name",Environment("WindowName"), "There is no Window opened on the application with the specified window name (Or) There is some problem accessing the class of the given object , The Window Name You have given is->"&Environment("WindowName")&" Please verify and run again.","Failed",objParent)
	Exit Function

End If
End Function

'********************************************************************************************************************************************************************************
' Function Name :AVMode
' Description   :This function "AV" is Attribute Validation mode on all the classes.
' param param1 - Parameter passed from framework
' param param2 - Parameter passed from framework
' param param3 - Parameter passed from framework
' Param Param4	- Parameter passed from framework
' Param Param5 - Parameter passed from framework
' Author	:	DSTWS TA2000 Automation Team
' Creation Date	: 	April, 2011
' Reviewed By 	:	DSTWS TA2000 Automation Team
' Modified By	:
' Modified Date	:	

'*********************************************************************************************************************************************************************************

Function AVMode(strParentObject,strChildobject,Param1,Param2,Param3,Param4,Param5,Param6,Param7,Param8,Param9,Param10,UIName,ExpectedResult)
	On Error Resume Next
	err.clear
	If strParentObject<>strChildobject Then
		set objParent=Eval(strParentObject)    
		set  objTestObject=Eval(strParentObject&"."&strChildobject)
	Else
		set  objTestObject=Eval(strParentObject)
	End If
		Param1=FormatData(Param1)
		Param2=FormatData(Param2)
'		Param3=FormatData(Param3)
'		Param4=FormatData(Param4)
'		Param5=FormatData(Param5)
'		Param6=FormatData(Param6)
'		Param7=FormatData(Param7)
'		Param8=FormatData(Param8)
'		Param9=FormatData(Param9)
'		Param10=FormatData(Param10)
	If objTestObject.Exist(0) Then
		
		strClass=objTestObject.getroproperty("micClass")
		If LCase(strClass)="webtable"  Then		
			AVMode=VerifyPropertyInWebTable(strParentObject,strChildobject,Param1,Param2,Param3,Param4,Param5,Param6,Param7,Param8,Param9,Param10,UIName,ExpectedResult)
		ElseIf LCase(strClass)="activex"  Then
			If LCase(Param1)="rowselection" Or LCase(Param1)="row selection" Then
				Param1="RowSel"
			AVMode= VerifyProperty(strParentObject,strChildobject,Param1,Param2,UIName,ExpectedResult)
			ElseIf LCase(Param1)="colselection" Or LCase(Param1)="col selection" Then
				Param1="ColSel"
			AVMode= VerifyProperty(strParentObject,strChildobject,Param1,Param2,UIName,ExpectedResult)
			End If			
		Else
		AVMode= VerifyProperty(strParentObject,strChildobject,Param1,Param2,UIName,ExpectedResult)
		End If
	Else
		Reporter.ReportEvent micFail,"There is no Window opened on the application with the specified window name (Or) There is some problem accessing the class of the given object , The Window Name You have given is->"&Environment("WindowName"),"Please verify and run again"
        Call ReportResult (Environment("WindowName"),"There should be a window opened in the application with the specified window name",Environment("WindowName"), "There is no Window opened on the application with the specified window name (Or) There is some problem accessing the class of the given object , The Window Name You have given is->"&Environment("WindowName")&" Please verify and run again.","Failed",objParent)
		Exit Function
	
	End If
End Function

'********************************************************************************************************************************************************************************
' Function Name :V Mode
' Description   :This function "VMode" is Validate mode on all the classes.
' param param1 - Parameter passed from framework
' param param2 - Parameter passed from framework
' param param3 - Parameter passed from framework
' Param Param4	- Parameter passed from framework
' Param Param5 - Parameter passed from framework
' Author	:	DSTWS TA2000 Automation Team
' Creation Date	: 	April, 2011
' Reviewed By 	:	DSTWS TA2000 Automation Team
' Modified By	:
' Modified Date	:	

'*********************************************************************************************************************************************************************************

Function VMode(strParentObject,strChildobject,Param1,Param2,Param3,Param4,Param5,Param6,Param7,Param8,Param9,Param10,UIName,ExpectedResult)

	If strParentObject<>strChildobject Then
		set objParent=Eval(strParentObject)    
		set  objTestObject=Eval(strParentObject&"."&strChildobject)
	Else
		set  objTestObject=Eval(strParentObject)
	End If
	
	On Error Resume Next
	Err.Clear
	
If objTestObject.Exist(0) Then
	
	strClass = objTestObject.GetROProperty("micClass")
	Select Case LCase(strClass)
	Case "wincheckbox"
		Param1=FormatData(Param1)
		VMode= VerifyCheckBox(strParentObject,strChildobject,Param1,UIName,ExpectedResult)
	Case "webcheckbox"
		Param1=FormatData(Param1)
		VMode= VerifyCheckBox(strParentObject,strChildobject,Param1,UIName,strExpectedResult)
	Case "winlist"
		Param1=FormatData(Param1)
		VMode= VerifyValueInWinList(strParentObject,strChildobject,Param1,UIName,strExpectedResult)
	Case "weblist"
		Param1=FormatData(Param1)
		VMode= VerifyValueInWebList(strParentObject,strChildobject,Param1,UIName,strExpectedResult)
	Case "winlistview"
		Param1=FormatData(Param1)
		Param2=FormatData(Param2)
		Param3=FormatData(Param3)
		VMode= VerifyDataInGrid(strParentObject,strChildobject,Param1,Param2,Param3,UIName,ExpectedResult)
	Case "wincombobox"
		Param1=FormatData(Param1)
		VMode= VerifyValueInCombo(strParentObject,strChildobject,Param1,UIName,ExpectedResult)
	Case "winradiobutton"
		Param1=FormatData(Param1)
		VMode= VerifyRadioButton(strParentObject,strChildobject , Param1,UIName,ExpectedResult)
	Case "webradiogroup"
		Param1=FormatData(Param1)
		VMode= VerifyRadioButton(strParentObject,strChildobject ,Param1,UIName,ExpectedResult)
	Case "winedit"
		Param1=FormatData(Param1)
		VMode= VerifyTextInEdit(strParentObject,strChildobject,Param1,UIName,ExpectedResult)
	Case "webedit"
		Param1=FormatData(Param1)
		VMode= VerifyTextInEdit(strParentObject,strChildobject,Param1,UIName,ExpectedResult)
	Case "winobject"
		Param1=FormatData(Param1)
		VMode= VerifyTextInEdit(strParentObject,strChildobject,Param1,UIName,ExpectedResult)
	Case "activex"
		strNatClass=objTestObject.GetROProperty("nativeclass")
        If Instr(1,LCase(strNatClass),"grid")>0 Then
        	Param1=FormatData(Param1)
			Param2=FormatData(Param2)
			Param3=FormatData(Param3)
			VMode= VerifyDataInActiveXGrid(strParentObject,strChildobject,Param1,Param2,Param3,UIName,ExpectedResult)
		Else
			Param1=FormatData(Param1)
			VMode= VerifyTextInEdit(strParentObject,strChildobject,Param1,UIName,ExpectedResult)
		End If
	Case "wineditor"
		Param1=FormatData(Param1)
		VMode= VerifyTextInEdit(strParentObject,strChildobject,Param1,UIName,ExpectedResult)   
	Case "wintab"
		Param1=FormatData(Param1)
		VMode= VerifyTab(strParentObject,strChildobject,Param1,UIName,ExpectedResult)
	Case "static"
		Param1=FormatData(Param1)
		VMode= VerifyStaticText(strParentObject,strChildobject,Param1,UIName,ExpectedResult)
	Case "dialog"
		Param1=FormatData(Param1)
		VMode=VerifyTextInDialog(strParentObject,strChildobject,Param1,UIName,ExpectedResult)
	Case "link"
		VMode=VerifyPresenceOfLink(strParentObject,strChildobject,UIName,strExpectedResult)
	Case"webelement"
		Param1=FormatData(Param1)
		VMode=VerifyText(strParentObject,strChildobject,Param1,UIName,strExpectedResult)
	Case "webbutton"
        VMode= VerifyPresenceOfButton(strParentObject,strChildobject,UIName,strExpectedResult)
	Case "webtable"
		arrParam=Param1&";"&Param2&";"&Param3&";"&Param4&";"&Param5&";"&Param6&";"&Param7&";"&Param8&";"&Param9&";"&Param10
		arrParam=Split(arrParam,";")
		For i=UBound(arrParam) to 0 Step -1
			If arrParam(i)="" Or arrParam(i)=Empty Then
				ReDim Preserve arrParam(i-1)
			Else
				classParam=arrParam(i)
                Exit For
			End If
    	Next
		Select Case(LCase(Trim(classParam)))
		Case "webelement"
			RefData1=FormatData(Param1)
			RefData2=FormatData(Param2)
			colName=FormatData(Param3)
			RowIncrement=FormatData(Param4)
			strVerifyText=FormatData(Param5)
			VMode=VerifyTextInWebTable(strParentObject,strChildobject,RefData1,RefData2,colName,RowIncrement,strVerifyText,UIName,strExpectedResult)
		Case "webcheckbox"
			RefData1=FormatData(Param1)
			RefData2=FormatData(Param2)
			colName=FormatData(Param3)
			RowIncrement=FormatData(Param4)
			strValuetoVerify=FormatData(Param5)
			VMode=VerifyCheckBoxInTable(strParentObject,strChildobject,RefData1,RefData2,colName,RowIncrement,strValuetoVerify,UIName,strExpectedResult)
		Case "webradiogroup"
			RefData1=FormatData(Param1)
			RefData2=FormatData(Param2)
			colName=FormatData(Param3)
			RowIncrement=FormatData(Param4)
			strValuetoVerify=FormatData(Param5)
			VMode=VerifyRadioGroupInTable(strParentObject,strChildobject,RefData1,RefData2,colName,RowIncrement,strValuetoVerify,UIName,strExpectedResult)
         		Case "webedit"
			 RefData1=FormatData(Param1)
			RefData2=FormatData(Param2)
			colName=FormatData(Param3)
			RowIncrement=FormatData(Param4)
			strValuetoVerify=FormatData(Param5)
			VMode=VerifyValueInEditBoxInTable(strParentObject,strChildobject,RefData1,RefData2,colName,RowIncrement,strValuetoVerify,UIName,strExpectedResult)
		Case "weblist"
			RefData1=FormatData(Param1)
			RefData2=FormatData(Param2)
			colName=FormatData(Param3)
			RowIncrement=FormatData(Param4)
			strValuetoVerify=FormatData(Param5)
			VMode=VerifyWebListItemInTable(strParentObject,strChildobject,RefData1,RefData2,colName,RowIncrement,strValuetoVerify,UIName,strExpectedResult)

		Case "webbutton"
			RefData1=FormatData(Param1)
			RefData2=FormatData(Param2)
			colName=FormatData(Param3)
			RowIncrement=FormatData(Param4)
			emptyParam=FormatData(Param5)
			VMode= VerifyButtonInTable(strParentObject,strChildobject,RefData1,RefData2,colName,RowIncrement,emptyParam,UIName,strExpectedResult)

		Case "link"
			RefData1=FormatData(Param1)
			RefData2=FormatData(Param2)
			colName=FormatData(Param3)
			RowIncrement=FormatData(Param4)
			emptyParam=FormatData(Param5)
			VMode= VerifyLinkInTable(strParentObject,strChildobject,RefData1,RefData2,colName,RowIncrement,emptyParam,UIName,strExpectedResult)

		End Select        
	End Select
	Else
	Reporter.ReportEvent micFail,"There is no Window opened on the application with the specified window name (Or) There is some problem accessing the class of the given object , The Window Name You have given is->"&Environment("WindowName"),"Please verify and run again"
	Call ReportResult (Environment("WindowName"),"There should be a window opened in the application with the specified window name",Environment("WindowName"), "There is no Window opened on the application with the specified window name (Or) There is some problem accessing the class of the given object , The Window Name You have given is->"&Environment("WindowName")&" Please verify and run again.","Failed",objParent)
	Exit Function

    End If	
 End Function

 '********************************************************************************************************************************************************************************
' Function Name :VLMode
' Description   :This function "VLMode" is Attribute Validation mode on all the classes.
' param param1 - Parameter passed from framework
' param param2 - Parameter passed from framework
' param param3 - Parameter passed from framework
' Param Param4	- Parameter passed from framework
' Param Param5 - Parameter passed from framework
' Author	:	DSTWS TA2000 Automation Team
' Creation Date	: 	April, 2011
' Reviewed By 	:	DSTWS TA2000 Automation Team
' Modified By	:
' Modified Date	:	


'*********************************************************************************************************************************************************************************
Function VLMode(strParentObject,strChildobject,Param1,Param2,Param3,Param4,Param5,Param6,Param7,Param8,Param9,Param10,UIName,ExpectedResult)

	set objParent=Eval(strParentObject)    
	set  objTestObject=Eval(strParentObject&"."&strChildobject)
	On Error Resume Next
	Err.Clear
If objTestObject.Exist(0) Then
	
	strClass = objTestObject.GetROProperty("Class Name")
	Select Case LCase(strClass)
	Case "winlist"
		Param1=FormatData(Param1)
		VLMode=VerifyValueInWinList(strParentObject,strChildobject,Param1,UIName,ExpectedResult)
	Case "wincombobox"
		Param1=FormatData(Param1)
		VLMode=VerifyValueInCombo(strParentObject,strChildobject,Param1,UIName,ExpectedResult)
    		'***************************
		Case "webtable"
		arrParam=Param1&";"&Param2&";"&Param3&";"&Param4&";"&Param5&";"&Param6&";"&Param7&";"&Param8&";"&Param9&";"&Param10
		arrParam=Split(arrParam,";")
		For i=UBound(arrParam) to 0 Step -1
			If arrParam(i)="" Or arrParam(i)=Empty Then
				ReDim Preserve arrParam(i-1)
			Else
				classParam=arrParam(i)
                Exit For
			End If
    	Next
		Select Case(LCase(Trim(classParam)))

		Case "weblist"
			Param1=FormatData(Param1)
			Param2=FormatData(Param2)
			Param3=FormatData(Param3)
			Param4=FormatData(Param4)
			Param5=FormatData(Param5)
			VLMode=VerifyWebListItemInTable(strParentObject,strChildobject,Param1,Param2,Param3,Param4,Param5,UIName,strExpectedResult)

		End Select        
     '*********************************
    End Select
	Else
		Reporter.ReportEvent micFail,"There is no Window opened on the application with the specified window name (Or) There is some problem accessing the class of the given object , The Window Name You have given is->"&Environment("WindowName"),"Please verify and run again"
        Call ReportResult (Environment("WindowName"),"There should be a window opened in the application with the specified window name",Environment("WindowName"), "There is no Window opened on the application with the specified window name (Or) There is some problem accessing the class of the given object , The Window Name You have given is->"&Environment("WindowName")&" Please verify and run again.","Failed",objParent)
		Exit Function

End If
End Function
 
 '********************************************************************************************************************************************************************************
' Function Name :WPMode
' Description   :This function "WPMode" is Wait Property mode on all the classes.
' Param Param1 - Parameter passed from framework
' Param Param2 - Parameter passed from framework
' Param Param3 - Parameter passed from framework
' Param Param4	- Parameter passed from framework
' Param Param5 - Parameter passed from framework
' Author	:	DSTWS TA2000 Automation Team
' Creation Date	: 	April, 2011
' Reviewed By 	:	DSTWS TA2000 Automation Team
' Modified By	:
' Modified Date	:	

'*********************************************************************************************************************************************************************************

Function WPMode(strParentObject,strChildobject,Param1,Param2,Param3,Param4,Param5,Param6,Param7,Param8,Param9,Param10,UIName,ExpectedResult)

	Param1=FormatData(Param1)
	Param2=FormatData(Param2)
	Param3=FormatData(Param3)

	If ExpectedResult="" Then
		ExpectedResult="Wait Property on: "&UIName&" should be obtained the required property value("&UCase(Param1)&" and "&UCase(Param2)&") in the given time out."
	End If
	
	If strParentObject<>strChildobject Then
		set objParent=Eval(strParentObject)    
		set  objTestObject=Eval(strParentObject&"."&strChildobject)
	Else
		set  objTestObject=Eval(strParentObject)
	End If
	If Param3<>"" Then
        With objTestObject
			If .WaitProperty(Param1, Param2, Param3) = False Then
				Reporter.ReportEvent micFail,"Wait Property","Object did not obtained the property value in the given time out."
				Call ReportResult  (UIName, ExpectedResult,UCase(Param1)&";"&UCase(Param2) ,"Wait Property on: "&UIName&" did not obtained the required property value("&UCase(Param1)&" and "&UCase(Param2)&") in the given time out." , "Failed" ,objParent)
			Else
				Reporter.ReportEvent micPass,"Wait Property","Object obtained the property value in the given time out."
				Call ReportResult  (UIName, ExpectedResult,UCase(Param1)&";"&UCase(Param2) ,"Wait Property on: "&UIName&" obtained the required property value("&UCase(Param1)&" and "&UCase(Param2)&") in the given time out." , "Passed" ,objParent)
			End If
		End With
	Else
		With objTestObject
			If .WaitProperty(Param1, Param2) = False Then
				Reporter.ReportEvent micFail,"Wait Property","Object did not obtained the property value("&UCase(Param1)&" and "&UCase(Param2)&") in the given time out."
				Call ReportResult  (UIName, ExpectedResult,UCase(Param1)&";"&UCase(Param2) ,"Wait Property on: "&UIName&" did not obtained the required property value("&UCase(Param1)&" and "&UCase(Param2)&") in the given time out." , "Failed" ,objParent)
			Else
				Reporter.ReportEvent micPass,"Wait Property","Object obtained the property value("&UCase(Param1)&" and "&UCase(Param2)&") in the given time out."
				Call ReportResult  (UIName, ExpectedResult,UCase(Param1)&";"&UCase(Param2) ,"Wait Property on: "&UIName&" obtained the required property value("&UCase(Param1)&" and "&UCase(Param2)&") in the given time out." , "Passed" ,objParent)
			End If
		End With
	End If
End Function

'TAF 10.1 new code Start
Public Sub CreateFileForCode()         '10.1


									If  VerifyEnvVariable("CodeGenerationFolder") Then
									Environment("CodeGenerationPath")=Environment("CodeGenerationFolder")&"\"&Lcase(Trim(DataTable.GetSheet("TestPlan").getParameter("TestCaseName")))&".txt"
									Set objFs=createobject("scripting.filesystemobject")
									If  objFs.FileExists(Environment("CodeGenerationPath"))Then
										objFs.DeleteFile(Environment("CodeGenerationPath")) 
									End If
								End If
	
End Sub
'TAF 10.1 new code End


'TAF 10.1 new code Start
Function getContNumber(filePath, tcName, ScenOrTc)


Dim  IntReccount
Dim  blnScenario
Dim  Intcomp 
Dim  arrFields
Dim  arrRecords
Dim  IntCounter

  blnScenario = ScenOrTc

Set objcon = createobject("ADODB.Connection")
objcon.connectionstring = "DRIVER={Microsoft Excel Driver (*.xls)};DBQ="&filePath&";Readonly=True"
objcon.Open
Set rs = createobject("ADODB.Recordset")
rs.CursorLocation=3
rs.Open "select * from  [TestPlan$]",objcon, 1, 3


IntReccount = rs.RecordCount

'arrFields = array("Scenario_Keyword","TestCaseName","Execute")
arrFields = array("Scenario_Keyword","TestCaseName")

rs.MoveFirst

arrRecords = rs.GetRows(IntReccount,0,arrFields)
If blnScenario Then

	CaseOrScen="Scenario"

	else
	CaseOrScen="TestCase"
End If


For IntCounter=0 to IntReccount-1
	Intcomp=-1
			if(isnull(arrRecords(0,IntCounter)) or isnull(arrRecords(1,IntCounter))) then
	
					
			elseif( not isnull(arrRecords(0,IntCounter))  and not isnull(arrRecords(1,IntCounter))  ) then
					Intcomp = 	1
'					if(isnull(arrRecords(2,IntCounter))) then
'					Intcomp = 	1
'					elseif(strcomp(cstr(arrRecords(2,IntCounter)),"yes",1)=0) then
'					Intcomp = 1
'					end if
	
			end if
	
			if(Intcomp=1) then
	
				if(	strcomp(cstr(arrRecords(0,IntCounter)),CaseOrScen,1)=0 and strcomp(cstr(arrRecords(1,IntCounter)),tcName,1)=0)  then
	
					
					getContNumber=IntCounter+1

						Exit for
					else
						getContNumber="NotFound"
					
				end if
				
			end if
	
Next

Set objcon = nothing

End Function
'TAF 10.1 new code End


