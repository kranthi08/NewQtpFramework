
''
' @# author DSTWS
' @# Version 7.0.7a
' @# Version 7.0.8  Revision History: Excel Revised code with new excel global  object "objResWorkBook"
'@ Modified by DT77215 on 05Nov 2009 handled creation of excel object for slow performance machines
' @# Version 7.0.9A  Kill Process, dlete temp files


'Function CreateTAFObjects()

' @# Version 7.0.9 - Modified by DT77471 on 19 mar 10, to add maximum wait time for failure step @CR 62@ line 221
' @# Version 7.0.9 - Modified by DT77471 on 19 mar 10, to handle recovery for popups, handle repeat current step @Bug 63@ Line 225 
' @# Version 7.0.9 - Modified by DT77471 on 19 mar 10, to get the testcase description in results file @CR 60@ Line 243, 389
' @# Version 7.0.9 - Modified by DT77471 on 30 mar 10, to kill the excel process using switch while running from test lab @CR 59@
' @# Version 7.0.10 - Modified by DT77215 on 01 Aug 10 , Added Extra column for DataRow Capturing

'----code for BPT-----------


'Set TSTestFact   = QCUTil.CurrentTestSet.TSTestFactory


'Set TestSetTestsList = TSTestFact.NewList("") 
'
'
'For Each theTSTest In TestSetTestsList
'	Print theTSTest.Name
'	Print theTSTest.ID
'Next
'
'msgbox TestSetTestsList.Item(1).Name
'
'msgbox TestSetTestsList.Count
'
'msgbox TestSetTestsList.Items.count

On error resume next
err.clear
intTempEnvExistTest = Environment("ExcelResults") 
	
If err.number <>0 Then
	Call LoadSettingsFile()
End If
	
On error goto 0
'---------------------------------
	
Environment("IsSystemSlow")= "no"

	
'If Trim(Lcase(Environment("ExcelResults"))) = "yes" Then
	On error resume Next				'Kill excel process while running from test lab
	Set theTSTest = QCUtil.CurrentTestSetTest
	If Not theTSTest Is Nothing Then	
        
	If  VerifyEnvVariable("DeleteTempFiles") Then
			If Environment("DeleteTempFiles") Then
				Call DeleteTempFiles
			End If			
		End If
	End If

        If  VerifyEnvVariable("CloseExcelProcess") Then
	        If UCase(Environment("CloseExcelProcess")) = "TRUE" then
				 Environment("CloseExcelProcess") = True
			else
				Environment("CloseExcelProcess") = False
			End If
			If Environment("CloseExcelProcess") Then
				'call  fnCloseApplication("EXCEL.EXE")
				'KillExcelProcesses "Microsoft Excel - Plan*"
				 If  VerifyEnvVariable("ProcessToIgnore") Then
					If Environment("ProcessToIgnore") <> "" then
					 	KillExcelProcesses Environment("ProcessToIgnore")
					else
						call  fnCloseApplication("EXCEL.EXE")	
					End If
				else
					call  fnCloseApplication("EXCEL.EXE")
				End If
			End If			
	End If


	'On error goto 0 
	Dim appExcel
	previousTime=Now()
	Set appExcel = CreateObject("Excel.Application")
	
	Environment("ExcelVersion") = appExcel.Application.Version  ''''''Changes 9A
	
'	appExcel.Visible = True  ' 0.0123
	'If DateDiff("s",previousTime,now())< 8 Then
	If DateDiff("s",previousTime,now()) < cint(Environment("MaxTimeToCreateExcelObject")) Then
	 Environment("IsSystemSlow")= "no"
	else
	 Environment("IsSystemSlow")= "Yes"
	 Reporter.ReportEvent micWarning , "System Slow ", "System was slow - Terminating Excel Results. Please Refer to QTP results "
	 If Environment("EventLog") Then
	 End If
	 WriteToEvent( "System Slow "  & vbtab & "System was slow - Terminating Excel Results. Please Refer to QTP results")
	end if
	appExcel.DisplayAlerts = False
	Set objFso = CreateObject("Scripting.FileSystemObject")'Creates the reference for the file system object
'End Function
'End If


Dim objResWorkBook

 ''

 ' This function MakePath is to create the result file in XLS format and format the Header section
 ' @author DT77235
 ' @param oFso - Instance of FileSystemObject
 ' @param sPath - Required path (must be fully qualified)
 ' @return True - Path exists, False - Path does not exist. 
 ' @Modified By DT77734, DT77742
 ' @Modified on: 03 Aug 2009 

Function MakePath()
	strFolderName=Environment("ResultPath")'Getting the path of the test  from where it is running  'checks if the folder already exists 

	If not  isobject(objFso) Then
		Set objFso=createobject("scripting.filesystemobject")
	End If

	If not objFso.Folderexists(strFolderName)Then
		objFso.CreateFolder(strFolderName)
	End If

	strTestPlanPath =Environment("TestPlanPath")' DataTable("TestPlanPath","Global")
	arrTestPlanPath =split(strTestPlanPath,"\",-1,1)
	strTestPlan = arrTestPlanPath(Ubound(arrTestPlanPath))
	arrTestPlanName =split(strTestPlan,".",-1,1)
	strTestPlanName = arrTestPlanName(Lbound(arrTestPlanName))
	strTime = day(now) & "_"& monthname(month(now)) &"_" & year(now) &"_" & hour(time) & "_" & minute(time) & "_" &second(time)
	Environment("TimeStamp") = strTime  ' Argus 

	'strResultFilepath=strFolderName&"\"&strTestPlanName &"_" & strTime &"_Result.xls"'Craeting the html result file
	strResultFilepath=strFolderName&"\"&Environment("DriverName")&"_"&strTestPlanName &"_" & strTime &"_Result.xls"'Craeting the html result file
	
	strEventLogPath=strFolderName&"\"&Environment("DriverName")&"_"&strTestPlanName &"_" & strTime&"_EventLog.txt" 'For Creating Event Log
	strScriptLogPath=strFolderName&"\"&Environment("DriverName")&"_"&strTestPlanName &"_" & strTime&"_ScriptLog.txt"


'    strResultFilepath = "D:\HTML\Results\TestPlan_19_October_2011_11_47_5_Result.xls"   ' 9.0.0
	If Trim(Lcase(Environment("IsSummaryResult"))) = "yes" Then
'			 strResultFilepath = "D:\HTML\Results\SummaryResults.xls"   
			strResultFilepath = Environment("SummaryResultsFolder") & "\" & Environment("SumResFilName")  
	End If

    Environment("ResultFilePath")=strResultFilepath' To set the result file path to the environment variable
	Environment("EventLogPath")=strEventLogPath'To set the Event log path to 
	Environment("ScriptLogPath")=strScriptLogPath
	If Environment("RunFromQC")  Then
		If Environment("FirstTest")  Then
				If objFso.fileexists(strResultFilepath) Then
					objFso.DeleteFile strResultFilepath
				End If
		End If
	End If
	

	If   not objFso.fileexists(strResultFilepath)Then
			Call CreateResultFile(strResultFilepath)
			Environment("IncValHolder")=1 
		Else
			Set objResWorkBook = appExcel.Workbooks.Open (Environment("ResultFilePath"))   '9.0.0
		Set objSheet2 = appExcel.Sheets(2)
			Environment("IncValHolder") = objSheet2.UsedRange.Rows.Count +1
	End If
	
	' HTML
	If Not LCase(Trim(Environment("HTMLResults")))="no" Then
		strTestPlanPath =Environment("TestPlanPath")
		arrTestPlanPath =split(strTestPlanPath,"\",-1,1)
		strTestPlan = arrTestPlanPath(Ubound(arrTestPlanPath))
		arrTestPlanName =split(strTestPlan,".",-1,1)
		strTestPlanName = arrTestPlanName(Lbound(arrTestPlanName))
		strTime = day(now) & "_"& monthname(month(now)) &"_" & year(now) &"_" & hour(time) & "_" & minute(time) & "_" &second(time)
		Environment("TimeStamp") = strTime  
		strResultFilepath=strFolderName&"\"&Environment("DriverName")&"_"&strTestPlanName &"_" & strTime &"_Result.html"'Craeting the html result file
		Environment("HTMLResultFilePath")=strResultFilepath
		CreateHTMLResults()
	End If

End function



 ''
 ' This function CreateResultFile creates the result file in XLS format and formats the Header section.
 ' @author DT77235
 ' @param FilePath String specifying Path of the result files
 ' @Modified By DT77734, DT77742
 ' @Modified on: 03 Aug 2009 
 ' @Modified  by 77215 on: 19 Nov 2009 handled issue with workbok having single sheet

Function CreateResultFile(FilePath) 'Create an object of Excel application

     If not isobject(appExcel) Then
		Set appExcel = CreateObject("Excel.Application")
	End If
	If Trim(Lcase(Environment("ExcelResults"))) = "no"  or Trim(Lcase(Environment("IsSystemSlow")))= "yes" Then
		Exit function
	End If
  'appExcel.Visible=True
  appExcel.Workbooks.Add
     strExecutedBy = Environment.Value("UserName")
  strApplication=Environment("ApplicationName")
  strRelease= Environment("Version")
  	If appExcel.Sheets.Count < 2 Then
		appExcel.Sheets.Add
	end if
    Set objSheet = appExcel.Sheets(1)'Get the object of the first sheet in the workbook
	appExcel.Sheets(1).Select

 With objSheet
  .Name = "Test_Summary" 'Rename the first sheet to "Test_Summery"
  .Range("B1").Value = "        Test Case Summary Report"  'Set the Heading
  .Range("B1:C2").Merge    
  .Range("B1:C2").Interior.ColorIndex = 53  'Set color and Fonts for the Header
  .Range("B1:C2").Font.ColorIndex = 19
  .Range("B1:C2").Font.Bold = True

  .Range("D1").Value =  "Test Executed on :      " &  Environment("vegaTestVersion")  'Set the Heading
  .Range("D1:H2").Merge    
  .Range("D1:H2").Interior.ColorIndex = 45  'Set color and Fonts for the Header
  .Range("D1:H2").Font.Bold = True

  .Range("B3").Value = "Application Name"
  .Range("C3").Value =   strApplication
  .Range("B4").Value = "Version"
  .Range("C4").Value =  strRelease
  .Range("B5").Value = "Executed By"
  .Range("C5").Value =  strExecutedBy
  .Range("B6").Value = "Date Executed"  'Set the Date and time of Execution
  .Range("B7").Value = "Test Start Time"
  .Range("B8").Value = "Test End Time"
  .Range("B9").Value = "Test Duration "    
  .Range("C6").Value = Date
  .Range("C7").Value = Time
  .Range("C8").Value = Time
  .Range("C9").Value = "=R[-1]C-R[-2]C"
  .Range("C9").NumberFormat = "[h]:mm:ss;@"
  'Format the Date and Time Cells
  .Range("B3:C16").Interior.ColorIndex = 40
  .Range("B3:C16").Font.ColorIndex = 12
  .Range("B3:A16").Font.Bold = True
  .Range("C10").value=0
  .Range("B10").Value = "Total Scenerios Executed"
  .Range("C11").value=0
  .Range("B11").Value = "Total Scenerios Passed"
  .Range("C12").value=0
  .Range("B12").Value = "Total Scenerios Failed"
  .Range("C13").Value = 0
  .Range("B13").Value = "Total Test Cases Executed"
'  .Range("C14").Value = "0"
  .Range("C14").Value = "=COUNTIF(C18:C9999,"&Chr(34)&"Passed"&chr(34)&")"'Modified By somesh
  .Range("B14").Value = "Total Test Cases Passed"
'  .Range("C15").Value = 0
  .Range("C15").Value =  "=COUNTIF(C18:C9999,"&Chr(34)&"Failed"&chr(34)&")"'Modified By somesh
  .Range("B15").Value = "Total Test Cases Failed"
  .Range("C16").Value = 0
  .Range("B16").Value = "Total Executed Steps"
  '.Range("C15").Value = 0
  .Range("B17").Value = "TestCase Name"
  .Range("C17").Value = "Status"
  .Range("D17").Value = "No Of Steps"
  .Range("E17").Value = "Description"  ' Argus
  .Range("E16").Value = "*Click the TestCase Name to see detail result."  ' Argus 
  ''TAF 10.1 new code Start
   .Range("F17").Value = "No Of DataRows"
   'TAF 10.1 new code End
  'Format the Heading for the Result Summery
  ''TAF 10.1 code modification Start
  .Range("B17:F17").Interior.ColorIndex = 53
  .Range("B17:F17").Font.ColorIndex = 19
  .Range("B17:F17").Font.Bold = True
  .Range("C3:C16").HorizontalAlignment = -4131 
  'Set Column width
  .Columns("B:F").Select
  .Columns("B:F").Autofit
  .Range("B1").Select
  'TAF 10.1 code modification End
 End With
 'Freez pane
 On error resume next
 'Get the object of the first sheet in the workbook
 Set objSheet = appExcel.Sheets(2)
 appExcel.Sheets(2).Select
 With objSheet
 'Rename the first sheet to "Detailed_Report"
   .Name = "Detailed_Report"
  'Set the Column widths
  .Columns("A:A").ColumnWidth = 30
  .Columns("B:B").ColumnWidth = 11
  .Columns("C:C").ColumnWidth = 35
  .Columns("D:D").ColumnWidth = 20
  .Columns("E:E").ColumnWidth = 35
  .Columns("F:F").ColumnWidth = 8
  .Columns("G:G").ColumnWidth = 12
  .Columns("A:H").HorizontalAlignment = -4131			' Argus F changed to G			G changed to H   Fanweb  7.0.11
  .Columns("A:H").WrapText = True				' Argus F changed to G				G changed to H 			Fanweb		7.0.11
  'Set the Heading for the Result Columns
  .Range("A1").Value = "UIName"
  .Range("B1").Value = "Screenshot"
  .Range("C1").Value = "Expected value"
  .Range("D1").Value = "Actual value"
  .Range("E1").Value = "Data"
  .Range("F1").Value = "Status"
  .Range("G1").Value = "DataRow"  'Argus
  .Range("H1").Value = "Step"  'FanWeb  7.0.11
  'Format the Heading for the Result Columns
  .Range("A1:H1").Interior.ColorIndex = 53
  .Range("A1:H1").Font.ColorIndex = 19
  .Range("A1:H1").Font.Bold = True
  .Range("A2:H2").Select
  appExcel.ActiveWindow.FreezePanes=true
 End With
 'Save the Workbook at the specified Path with the Specified Name
 appExcel.ActiveWorkbook.saveas FilePath
 'Close Excel
' appExcel.quit ' 0.0123
 	Set objResWorkBook = appExcel.Workbooks.Open (Environment("ResultFilePath"))  ' open workbook


End Function



 ''
 ' This function ReportResult writes the result to the result file on the last row
 ' @author DT77709
 ' @param sstrParentObject String specifying the Parent Object
 ' @param FilePath String specifying Path of the result file
 ' @param TestcaseName String specifying Name of the testcase
 ' @param StepName String specifying Name of the step for which the result is being written
 ' @param Description String specifying Desctiption of the step for which the result is being written
 ' @param strStatus Boolean specifying Status of the Step (0 = Pass and 1 = Fail)
 ' @param Expected String specifying Expected result for the step
 ' @param Actual String specifying Actual Result for the step
 ' @param MSG String specifying Error Message
 ' @Modified By DT77734, DT77742
 ' @Modified on: 03 Aug 2009 
 ' @Modified by DT77215 Implemented Scenerio and Test Case Reporting


Function ReportResult(strUIName, strExpected, strData, strActual, strStatus,objParent)


'''     If Trim(Lcase(Environment("ExcelResults"))) = "no"  or Trim(Lcase(Environment("IsSystemSlow")))= "yes" Then
''''		WriteToEvent(strStatus & vbtab & strActual)
'''		Exit function
'''	End If

	

	If Environment("Optional") and Trim(LCase(strStatus))="failed" then  ' sreenu babu		
		Exit function
	end if


   If Ucase(strStatus)="FAILED" Then	'To add the wait time for failed results				'To add the wait time to Results 7.0.8a
   		intSeconds=CLng(Environment("MaxSyncTime"))/1000
		strActual=strActual& " ::Wait time to identify the object is: "&intSeconds&" Seconds"
   End If
	
	If Environment("RepeatStep") and Environment("Counter")=1 Then			'To handle recovery Scenario   
		strStatus="warning"
		strActual="Recovery Invoked"
		strData=""
					
	End If
	
	If Trim(LCase(strStatus))="failed" then  
		Environment("CurrentStepStatus")= "Failed" 
	End If
	

If not Lcase(Trim(Environment("ExcelResults")))="no"  and not Trim(Lcase(Environment("IsSystemSlow")))= "yes"  Then

	On error resume next
	strTestScenerioName=Environment("strNewScenario")   ' get the new scenerio name
	strTestcasename=Environment("TestCaseName1")  ' get the new test case name 
	strTestCasedesc=Environment("TestCaseName") 		'get the test case description
'	Set objWorkBook = appExcel.Workbooks.Open (Environment("ResultFilePath"))  ' open workbook
'	WriteToEvent(strStatus & vbtab & strActual)
	Set objSheet = appExcel.Sheets("Test_Summary")  
	appExcel.Sheets("Test_Summary").Select ' activate the test_summary sheet
	With objSheet

'	Row=29 ' 9.0.0

      Row=Environment("IncValHolder")   ' 9.0.0
	  TCRow = .Range("C13").Value + 18  ' TCRow -- Represents Test cases as well as Test Scenerios 		
	  NewTC = False  ' Make new test case as false
	  NewTS=False 'Make New Scenerio as fasle
      'Check if it is a new Tetstcase
	  .Range("C16").Value = .Range("C16").Value + 1  ' increment  the Test step value by one
	 
	 
	  '************************************************Modified By Somesh********************************
 
	  If Trim(Lcase(Environment("IsSummaryResult"))) = "yes" Then
	  If Environment("TSorTC") Then ' test if test scenerio or test case has occured. if Environment("TSorTC")  is true then Test Scenerio has occured
			Environment("ScenerioStepCounter")=Environment("ScenerioStepCounter")+1
			Environment("SetCondition")=True ' Reset this counter as and when the new scenerio occurs
			Environment("ScenerioErrorIncremented")=False ' Condition for incrementing the Failure count of the Scenerio		
		If objSheet.Cells(TCRow - 1, 1).Value <>  (strTestcasename) Then
		    .Cells(TCRow, 2).Value =   strTestScenerioName  & "-->" & strTestcasename   
			  .Cells(TCRow, 1).Value = strTestcasename   
			appExcel.ActiveSheet.Hyperlinks.Add objSheet.Cells(TCRow, 2), "", "Detailed_Report!A" & Row+1, strTestcasename
		   If Trim(LCase(strStatus))="failed" then
				.Cells(TCRow, 3).Value = "Failed"
				.Range("C" & TCRow).Font.ColorIndex = 3
				.Range("C15").Value = .Range("C15").Value + 1
		   Else
				.Cells(TCRow, 3).Value = "Passed"
				.Range("C" & TCRow).Font.ColorIndex = 50
'				.Range("C14").Value = .Range("C14").Value + 1
		   End If
			.Cells(TCRow, 4).Value = 1
			NewTC = True
			.Range("C13").Value = .Range("C13").Value + 1
			'Set color and Fonts for the Header
			.Range("B" & TCRow & ":D" & TCRow).Interior.ColorIndex = 19
			.Range("B" & TCRow).Font.ColorIndex = 53
			.Range("B" & TCRow & ":D" & TCRow).Font.Bold = True
		Else
			.Range("D" & TCRow-1).Value = .Range("D" & TCRow-1).Value + 1
		End If
		If (Not NewTC) and Trim(LCase(strStatus))="failed"  then
		.Cells(TCRow-1, 3).Value = "Failed"
		.Range("C" & TCRow-1).Font.ColorIndex = 3
		End If
	End If ' End the TSorTC flag
  'Update the End Time
   .Range("C8").Value = Time
   'Set Column width
   .Columns("B:D").Autofit
        
   End If '***********************Modified By Somesh 
     
	  
	  If Environment("TSorTC") Then ' test if test scenerio or test case has occured. if Environment("TSorTC")  is true then Test Scenerio has occured
			Environment("ScenerioStepCounter")=Environment("ScenerioStepCounter")+1
		If Environment("blnFlagNewScenario") Then
			Environment("ScenerioLineHolder")=TCRow  ' hold only the value of the new scenerio row
			Environment("SetCondition")=True ' Reset this counter as and when the new scenerio occurs
			Environment("ScenerioErrorIncremented")=False ' Condition for incrementing the Failure count of the Scenerio		

			  If Trim(Lcase(Environment("IsSummaryResult"))) = "yes" Then
			 .Cells(TCRow, 2).Value = strTestScenerioName  & "-->" & strTestcasename   	 
			 Else
			 .Cells(TCRow, 2).Value = strTestScenerioName  
			  End If

			  .Cells(TCRow, 1).Value = strTestcasename

			  'TAF 10.1 new code Start. Added by Trao
			   .Cells(TCRow, 6).Value =  Environment("NoOfDatarows")
			   'TAF 10.1 new code End

              appExcel.ActiveSheet.Hyperlinks.Add objSheet.Cells(TCRow, 2), "", "Detailed_Report!A" & Row+1, strTestScenerioName  ' Add hyperlink
			 .Range("C" & TCRow).Font.ColorIndex = 3  '' Increment Scenerio Count
			 .Range("C10").Value = .Range("C10").Value + 1   'Increment the Scenerio counter
			 If Trim(LCase(strStatus))="failed" then
				.Cells(TCRow, 3).Value = "Failed"
				.Range("C" & TCRow).Font.ColorIndex = 3
				.Range("C12").Value = .Range("C12").Value + 1 				  ' increment scenerio count if the scenerio failed
				Environment("ScenerioErrorIncremented")=True
			  Else
				.Cells(TCRow, 3).Value = "Passed"
				.Range("C" & TCRow).Font.ColorIndex = 50
			'	.Range("C11").Value = .Range("C11").Value + 1 ' increment scenerio count if the scenerio has passed
			'	Environment("ScenerioPassIncremented")=True
			End If
			.Cells(TCRow, 4).Value = 1
			Environment("ScenerioEncountered")=True
			NewTCinTS = True   ' make the occurance of  a scenerio  ' statement waste remove it later		
    		.Range("B" & TCRow & ":D" & TCRow).Interior.ColorIndex = 19  	'	Set color and Fonts for the Header
			.Range("B" & TCRow).Font.ColorIndex = 53
			.Range("B" & TCRow & ":D" & TCRow).Font.Bold = True
			'Environment("blnFlagNewScenario")=false    
		Else
			If Environment("CurrentTestCaseName")<>( strTestcasename) Then
				NewTCinTS = True

				'TAF 10.1 new code Start. Added by Trao
			   .Cells(TCRow, 6).Value =  Environment("NoOfDatarows")
			   'TAF 10.1 new code End

			    If Trim(LCase(strStatus))="failed" then
					SetRow=Environment("ScenerioLineHolder")
					 If Trim(Lcase(Environment("IsSummaryResult"))) <> "yes" Then
					.Cells(SetRow, 3).Value = "Failed"
					End If
					.Range("C" & SetRow).Font.ColorIndex = 3
					If not(Environment("ScenerioErrorIncremented")) Then
						Environment("ScenerioErrorIncremented")=True
						.Range("C12").Value = .Range("C12").Value + 1   ' increment scenerio count if the scenerio failed
					End If
			    Else
				    SetRow=Environment("ScenerioLineHolder")
					 If Trim(Lcase(Environment("IsSummaryResult"))) <> "yes" Then
					.Cells(SetRow, 3).Value = "Passed"
					End If
					.Range("C" & SetRow).Font.ColorIndex = 50
				End If
				else
				If Trim(LCase(strStatus))="failed" then
					If Environment("SetCondition") and not(Environment("ScenerioErrorIncremented")) Then
					SetRow=Environment("ScenerioLineHolder")
					 If Trim(Lcase(Environment("IsSummaryResult"))) <> "yes" Then
                    .Cells(SetRow, 3).Value = "Failed"
					
					.Range("C" & SetRow).Font.ColorIndex = 3
					Environment("SetCondition")=False
					Environment("ScenerioErrorIncremented")=True
						.Range("C12").Value = .Range("C12").Value + 1   ' increment scenerio count if the scenerio failed
						End If
					End If
				Else
					SetRow=Environment("ScenerioLineHolder")
					 If Trim(Lcase(Environment("IsSummaryResult"))) <> "yes" Then
					.Cells(SetRow, 3).Value = "Passed"
					End If
					.Range("C" & SetRow).Font.ColorIndex = 50				
				End If  				
			End If
            SetRow=Environment("ScenerioLineHolder")
'''			.Cells(SetRow, 4).Value = Environment("ScenerioStepCounter")
     End if    '*************************************
	 Else
	 ' modify from here an indi test case has been encountered
		If objSheet.Cells(TCRow - 1, 1).Value <> ( strTestcasename) Then
			  If Trim(Lcase(Environment("IsSummaryResult"))) = "yes" Then
						.Cells(TCRow, 2).Value =    strTestScenerioName  & "-->"  & strTestcasename   
						.Cells(TCRow, 2).Value =    strTestcasename   
			Else
						 .Cells(TCRow, 2).Value =    strTestcasename  
			End If

			 .Cells(TCRow, 1).Value = strTestcasename

			 'TAF 10.1 new code Start. Added by Trao
			   .Cells(TCRow, 6).Value = Environment("NoOfDatarows")
			   'TAF 10.1 new code End

			appExcel.ActiveSheet.Hyperlinks.Add objSheet.Cells(TCRow, 2), "", "Detailed_Report!A" & Row+1, strTestcasename
		   If Trim(LCase(strStatus))="failed" then
				.Cells(TCRow, 3).Value = "Failed"
				.Range("C" & TCRow).Font.ColorIndex = 3
				.Range("C15").Value = .Range("C15").Value + 1
		   Else
				.Cells(TCRow, 3).Value = "Passed"
				.Range("C" & TCRow).Font.ColorIndex = 50
'				.Range("C14").Value = .Range("C14").Value + 1
		   End If
			.Cells(TCRow, 4).Value = 1
			NewTC = True
			.Range("C13").Value = .Range("C13").Value + 1
			'Set color and Fonts for the Header
			.Range("B" & TCRow & ":D" & TCRow).Interior.ColorIndex = 19
			.Range("B" & TCRow).Font.ColorIndex = 53
			.Range("B" & TCRow & ":D" & TCRow).Font.Bold = True
		Else
			.Range("D" & TCRow-1).Value = .Range("D" & TCRow-1).Value + 1
		End If
		If (Not NewTC) and Trim(LCase(strStatus))="failed"  then
		.Cells(TCRow-1, 3).Value = "Failed"
		.Range("C" & TCRow-1).Font.ColorIndex = 3
		End If
	End If ' End the TSorTC flag
  'Update the End Time
   .Range("C8").Value = Time
   'Set Column width
   .Columns("B:D").Autofit
 End With

'*****************************
'Select the Result Sheet
 Set objSheet = appExcel.Sheets("Detailed_Report")
 appExcel.Sheets("Detailed_Report").Select
 With objSheet
	' If Environment("ScenerioEnd") and Environment("ScenerioEncountered")Then
	 If Environment("ScenerioEnd") Then
			   .Range("A" & Row & ":H" & Row).Interior.Color =vbyellow '20'&hFFFF   	'Argus  F change to G		Fanweb G changed to   	H 7.0.11
			   .Range("A" & Row).value="EndOfScenario"
			   .Range("A" & Row).Font.Bold = True	 
				Environment("ScenerioEnd")=False
'				If ResetScenerioEncountered Then
				'Environment("ScenerioEncountered")=False
'				End If
				'Row=Row+1
	  End If
	If Environment("blnFlagNewScenario") Then
		Row=Row+1
	   .Range("A" & Row & ":H" & Row).Interior.Color =vbyellow '20'&hFFFF  	'Argus  F change to G		Fanweb G changed to   	H 7.0.11
	   .Range("A" & Row).value=Environment("strNewScenario")
	   .Range("A" & Row).Font.Bold = True 
	    'Row=Row+1
	    Environment("ResultDataRow") = Row  ' change the Data in logs
		Environment("IncValHolder")=Row
	    Environment("blnFlagNewScenario")=False
'		If not(ResetScenerioEncountered) Then
'			Environment("ScenerioEncountered")=True    
'		End If
	   end If
	If NewTC  or NewTCinTS then	
		'Row=Environment("IncValHolder")
		'If TCCount=1 Then
		Row = Row + 1
		'End If		
		.Range("A" & Row & ":H" & Row).Interior.ColorIndex = 40			'Argus  F change to G		Fanweb G changed to   	H 7.0.11
        Environment("ResultDataRow") = Row
        If  strTestCasedesc =""Then								'Write the test case descriptions in results file
			.Range("A" & Row).Value =  strTestcasename
		Else
			.Range("A" & Row).Value =  strTestcasename&" - "&strTestCasedesc
		End If
		'.Range("A" & Row).Value =  strTestcasename
		.Range("B" & Row).WrapText=False
   ' .Range("B" & Row).Value =  Environment("TestCaseDescription")
   'Set color and Fonts for the Header
		.Range("A" & Row & ":H" & Row).Interior.ColorIndex = 19 			'Argus  F change to G		Fanweb G changed to   	H 7.0.11
		.Range("A" & Row & ":H" & Row).Font.ColorIndex = 53						'Argus  F change to G		Fanweb G changed to   	H 7.0.11
		.Range("A" & Row & ":H" & Row).Font.Bold = True							'Argus  F change to G		Fanweb G changed to   	H 7.0.11
		 Row = Row + 1
		 Environment("IncValHolder")=Row
	End If
  .Range("A" & Row).Value = strUIName
	If  Trim(LCase(strStatus))="passed" Then
		.Range("F" & Row).Value = "Passed"
		.Range("F" & Row).Font.ColorIndex = 50
	Elseif Trim(LCase(strStatus)) = "warning" Then 
		.Range("F" & Row).Value = "Warning"	
		.Range("F" & Row).Font.Color = vbblue	
	Elseif Trim(LCase(strStatus)) = "done" Then 
		.Range("F" & Row).Value = "Done"	
		.Range("F" & Row).Font.Color = vbblack	
	Elseif Trim(LCase(strStatus)) = "" Then
		.Range("F" & Row).Value = "Done"	
		.Range("F" & Row).Font.Color = vbblack	
	ElseIf Trim(LCase(strStatus))="screenshot" Then
		On Error Resume Next
		ScreenShotPath=Environment("TestDir")&"\screenshot.png"
		Desktop.CaptureBitmap ScreenShotPath,true
		.Range("F" & Row).Value = "Screen"
		 .Range("A" & Row & ":H" & Row).Font.ColorIndex = 50	'Argus  F change to G		Fanweb G changed to   	H 7.0.11
   .Range("A" & Row & ":H" & Row).Font.Bold = True			'Argus  F change to G		Fanweb G changed to   	H 7.0.11
   .Range("B" & Row).Value = "Click Here"
   With .Range("B" & Row).AddComment
    .Shape.Height = 350
    .Shape.Width = 500
    .Shape.Fill.UserPicture ScreenShotPath
    Set objFso = CreateObject("Scripting.FileSystemObject")'Creates the reference for the file system object
    If objFso.Fileexists(ScreenShotPath)Then
     objFso.DeleteFile(ScreenShotPath)
    End If
      End With
  Else
   On Error Resume Next
   If  not IsObject(objParent) Then'check if the ObjParent is object or not
    set objParent= Desktop
   ElseIf objParent.Exist(0) Then
    set objParent= Desktop
   End If 
   ScreenShotPath=Environment("TestDir")&"\screenshot.png"
   objParent.CaptureBitmap ScreenShotPath,true
   If err.Number<>0 Then
    ScreenShotPath=Environment("TestDir")&"\screenshot.png"
    Desktop.CaptureBitmap ScreenShotPath,true
   End If
   .Range("F" & Row).Value = "Failed"
   .Range("A" & Row & ":H" & Row).Font.ColorIndex = 3			'Argus  F change to G		Fanweb G changed to   	H 7.0.11
   .Range("B" & Row).Value = "Click Here"
   With .Range("B" & Row).AddComment
    .Shape.Height = 350
    .Shape.Width = 500
    .Shape.Fill.UserPicture ScreenShotPath
    Set objFso = CreateObject("Scripting.FileSystemObject")'Creates the reference for the file system object
    If objFso.Fileexists(ScreenShotPath)Then
     objFso.DeleteFile(ScreenShotPath)
    End If
   End With
  End If
  .Range("B" & Row).Font.Bold = True
  .Range("C" & Row).Value = strExpected
  .Range("D" & Row).Value = strActual
  .Range("E" & Row).Value = strData
  .Range("G" & Row).value = Environment("ExecelResultsKeywords")	'Argus
  .Range("H" & Row).value =Environment("StepNumber")		'Fanweb		7.0.11
   Row = Row + 1  	
   Environment("IncValHolder")=Row
 End With
'****************************
 'Save the Workbook
 appExcel.Sheets("Test_Summary").Select
' objWorkBook.Save
' appExcel.Quit
 Environment("CurrentTestCaseName")=strTestcasename	

 End If

' HTML
  If Not LCase(Trim(Environment("HTMLResults")))="no" Then
	Call ReportHTMLResults(strUIName, strExpected, strData, strActual,strStatus,objParent)
End If

End Function


''
 ' This function is used to update the result report about the test data used in test case.
 ' @author DT77742
 
public function  WritetoResult_DataPath()

 On error resume next
' Set objWorkBook = appExcel.Workbooks.Open (Environment("ResultFilePath"))
 Row=  Environment("ResultDataRow")
 Set objSheet = appExcel.Sheets("Detailed_Report")
 appExcel.Sheets("Detailed_Report").Select
  If Environment("DataFromDataSheet") = "False" Then
   objSheet.Range("B" & Row).Value = "Warning : Test Data given directly in test case."
   objSheet.Range("B"&Row).Font.Color=VbRed 
  Else
    If (Environment("DataFromDataSheet") = "True" and Environment("Keywords")="") then
		If Environment("RunFromQC")  Then
			objSheet.Range("B" & Row).Value= "<TestDataPath =" & 	Environment("QCTestDataFilePath") & "> <Sheet=" & Environment("TestDataSheetName") & ">  <DataRow = Default Row>"  
			else
			objSheet.Range("B" & Row).Value= "<TestDataPath =" & Environment("TestDataPathToReport") & "> <Sheet=" & Environment("TestDataSheetName") & ">  <DataRow = Default Row>"  
		End If
    
    objSheet.Range("B" & Row).Font.Color=VbBlue
   Else  
   If  Environment("RunFromQC") Then
	   objSheet.Range("B" & Row).Value= "<TestDataPath =" &   	Environment("QCTestDataFilePath") & "> <Sheet=" & Environment("TestDataSheetName") & ">  <DataRow =" & Environment("Keywords")  &">"
	   else
	   objSheet.Range("B" & Row).Value= "<TestDataPath =" & Environment("TestDataPathToReport") & "> <Sheet=" & Environment("TestDataSheetName") & ">  <DataRow =" & Environment("Keywords")  &">"

   End If
    
    objSheet.Range("B" & Row).Font.Color=VbBlue
   End if
  End If 
 appExcel.Sheets("Test_Summary").Select
' objWorkBook.Save
' appExcel.Quit

End Function



''
 ' This function is used to write the script to Scrip Log text file.
 ' @author DT77742
 ' @param strScript String specifying the script to be written in Script log file.

Function WriteToScript(strScript)
   
   If LCase(Trim(Environment("ScriptLog"))) = "yes"Then
  Set objOpenFile = objFso.OpenTextFile(Environment("ScriptLogPath"), 8, True) 
        objOpenFile.Writeline(strScript)
   End If
      

End Function

'TAF 10.1 new code Start
Function UpdateScript(strScript)

		If VerifyEnvVariable("CodeGenerationPath")  Then
			Set objOpenFile = objFso.OpenTextFile(Environment("CodeGenerationPath"), 8, True) 
            objOpenFile.Writeline(strScript)
		End If
        

End Function
'TAF 10.1 new code End



''
 ' This function is used to write the Events to Event Log text file.
 ' @author DT77742
 ' @param strEvent String specifying the script to be written in Script log file.

Function WriteToEvent(strEvent) 
   
 If LCase(Trim(Environment("EventLog"))) = "yes" Then 
        Set objOpenFile = objFso.OpenTextFile(Environment("EventLogPath"), 8, True) 
        objOpenFile.Writeline(strEvent) 
 End If 
       

End Function




''
 ' This function is used to create header for Scrip Log and Event Log text files.
 ' @author DT77742
 ' @param strFilePath String specifying the path of the file.
 ' @param strConfigFlag String specifying the value whether to create script or event log files.

Function Header(strFilePath, strConfigFlag)
      
	If not isobject(objFso) Then
		set objFso=Createobject("Scripting.filesystemobject")
	End If
	If strConfigFlag= "yes" Then 
		If  objFso.FileExists(strFilePath) Then
				Set objOpenFile = objFso.OpenTextFile(strFilePath, 8, True) 
				objOpenFile.Writeline("********************************************************************Test Case Start***********************************************************************") 
				objOpenFile.Writeline("Test Case Name:"& Environment("TestCaseName1")) 
				objOpenFile.Writeline("**************************************************************************************************************************************************************") 
		else
				Set objOpenFile = objFso.OpenTextFile(strFilePath, 8, True) 
				objOpenFile.Writeline("**************************************** Test Executed on :      " &  Environment("vegaTestVersion")  & "**************************************************") 
				
				objOpenFile.Writeline("********************************************************************Test Case Start***********************************************************************") 
				objOpenFile.Writeline("Test Case Name:"& Environment("TestCaseName1")) 
				objOpenFile.Writeline("**************************************************************************************************************************************************************") 
				
		End If
	End If
          
		 

End Function


''
 ' This function is used to create footer for Scrip Log and Event Log text files.
 ' @author DT77742
 ' @param strFilePath String specifying the path of the file.
 ' @param strConfigFlag String specifying the value whether to create script or event log files.

Function Footer(strFilePath, strConfigFlag)
      
   If not isobject(objFso) Then
	   Set objFso=Createobject("Scripting.filesystemobject")
   End If
   If strConfigFlag= "yes" Then
    Set objOpenFile = objFso.OpenTextFile(strFilePath, 8, True)
    objOpenFile.Writeline("*******************************************************************Test Case End**************************************************************************") 
   End If
             
		 
End Function




Function DeleteTempFiles()
	Set ws = CreateObject("Wscript.shell")
	ws.run "cmd /C DEL /F /S /Q %TEMP%"
	wait(10)
	Set ws = nothing
End Function


Function KillExcelProcesses(strProcessToIgnore)
	Set ws = CreateObject("Wscript.shell")
	Set fso=createobject("scripting.filesystemobject")
	tmp=fso.GetSpecialFolder(2)

	If fso.FileExists(tmp &"\pid.txt") Then
		fso.DeleteFile tmp &"\pid.txt",1
	End If

	ret=ws.run("cmd /C tasklist /FI "& chr(34) &"Windowtitle eq Microsoft Excel - "& strProcessToIgnore & Chr(34) &">>" & tmp &"/pid.txt",1,true)
	Set readfso=fso.OpenTextFile(tmp &"\pid.txt",1)
	On error resume next
	err.clear
	strContents=readfso.readall()

	If err.number <> 0  Then
		 exceptPID = "" 
	 else
			arrLines = Split(strContents, vbCrLf)
			strString =arrLines(ubound(arrLines)-1)
        		
			arrString = Split(lcase(strString),".exe")
			exceptPID = ""
			If  ubound(arrString) > 0 Then
			
			arrPID = Split(arrString(ubound(arrString))," ")
			For i = lbound(arrPID) to ubound(arrPID)
			If  arrPID(i) <> ""Then
			'				Print  arrPID(i)
			exceptPID = arrPID(i)
			Exit For
			End If
			Next
			End If	

	End If
	If  exceptPID = "" Then
		ret=ws.run("cmd /C taskkill /F /FI "& chr(34) &  "IMAGENAME eq EXCEL.EXE"   &chr(34),1,true)
		else
	 	ret=ws.run("cmd /C taskkill /F /FI "& chr(34) & "PID ne " & exceptPID   & chr(34) & " /FI " &chr(34) & "IMAGENAME eq EXCEL.EXE" & chr(34),1,true)
	End If	

End Function

' 
' This function is used to lock the excel result Detailed Report sheet 
'  Created by :- DT77215 TAF Version 7.0.10

Function ProtectResults(byval objSheetInfo)

	Set objSummary= objSheetInfo
	objSummary.Protect(Environment("vegaTestVersion"))
	objResWorkBook.save

	On error resume next
	Set objSheet = appExcel.Sheets("Detailed_Report")
	appExcel.Sheets("Detailed_Report").select
	objSheet.Columns("A:J").Select
	objSheet.Columns("A:J").Locked = False
	objSheet.Columns("A:J").FormulaHidden = False
	
	objSheet.Columns("F:F").Select
	objSheet.Columns("F:F").Locked = True
	objSheet.Columns("F:F").FormulaHidden = False
	objSheet.Protect(Environment("vegaTestVersion"))
	err.clear

End Function

'TAF 10.1 new code Start
Function MakePathNonTAF()

		'strFolderName=Environment("TestDir")&"\"&"NonTAFResults"'Getting the path of the test  from where it is running  'checks if the folder already exists 
		If VerifyEnvVariable("ResultsPath") Then
		strFolderName=Environment("ResultsPath")
	Else
		strFolderName=Datatable.Value("ResultPath")
	End If

	If not  isobject(objFso) Then
		Set objFso=createobject("scripting.filesystemobject")
	End If

	If not objFso.Folderexists(strFolderName)Then
		objFso.CreateFolder(strFolderName)
	End If

	strTestName =Environment("TestName")' DataTable("TestPlanPath","Global")
    strTime = day(now) & "_"& monthname(month(now)) &"_" & year(now) &"_" & hour(time) & "_" & minute(time) & "_" &second(time)
	

	'strResultFilepath=strFolderName&"\"&strTestPlanName &"_" & strTime &"_Result.xls"'Craeting the html result file
	strResultFilepath=strFolderName&"\"&strTestName &"_" & strTime &"_Result.xls"'Craeting the html result file
	

    Environment("ResultFilePath")=strResultFilepath' To set the result file path to the environment variable

    
	If  Not Trim(Lcase(Environment("ExcelResults"))) ="no" Then

						If   not objFso.fileexists(strResultFilepath)Then
								Environment("vegaTestVersion") ="vegaTest 4.13.1"
								Call CreateResultFile(strResultFilepath)
								Environment("IncValHolder")=1 
					
											
												Environment("Optional") =False
												Environment("RepeatStep")=False
												Environment("IsSummaryResult")="No"
												Environment("strNewScenario") =Null
												Environment("TestCaseName1")=objFso.GetFileName(Environment("TestName"))
												Environment("TestCaseName")=""
												Environment("blnFlagNewScenario")=False
												Environment("TSorTC") =False
												Environment("ScenerioEnd")=False
												Environment("Counter") =1
												Environment("ExecelResultsKeywords")=""
												Environment("StepNumber")=2
												
											
					
					
						Else
										Set objResWorkBook = appExcel.Workbooks.Open (Environment("ResultFilePath"))   '9.0.0
							Set objSheet2 = appExcel.Sheets(2)
								Environment("IncValHolder") = objSheet2.UsedRange.Rows.Count +1
						End If


	End If
    	

	

	If Not LCase(Trim(Environment("HTMLResults")))="no" Then
		
		strTime = day(now) & "_"& monthname(month(now)) &"_" & year(now) &"_" & hour(time) & "_" & minute(time) & "_" &second(time)
		Environment("TimeStamp") = strTime  
        Environment("HTMLResultFilePath")=strFolderName&"\"&strTestName &"_" & strTime &"_Result.html"'
		CreateHTMLResults()

												
												Environment("Optional") =False
												Environment("RepeatStep")=False
                                                Environment("strNewScenario") =Null
												Environment("TestCaseName1")=objFso.GetFileName(Environment("TestName"))
												Environment("TestCaseName")=""
												Environment("blnFlagNewScenario")=False
												Environment("TSorTC") =False
												Environment("ScenerioEnd")=False
												Environment("Counter") =1
												Environment("ExecelResultsKeywords")=""
												Environment("StepNumber")=2
												Environment("TestCaseSheetNameToResults")=""
												'Environment("StepNum")=0
	End If
	
 
End Function
'TAF 10.1 new code End



