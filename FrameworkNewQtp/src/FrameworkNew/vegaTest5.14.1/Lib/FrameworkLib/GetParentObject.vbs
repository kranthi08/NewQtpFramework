''
' @# author DSTWS
' @# Version 7.0.7
	
	''
	' This function WaitIfIEBusy is to provide a dynamic wait to the script till the expected browser is available for test.
	' @author DSTWS
	' @param objBrowser String specifying the browser name
	' @Modified By DT77734
	' @Modified on: 16 Jul 2009 
	
Function WaitIfIEBusy(objBrowser)
	On Error Resume Next
	Dim seconds
	seconds=0
	start=Now
	Do While objBrowser.Object.Busy=True
		stp=Now
		seconds= datediff("s",start,stp)
		If seconds>180 Then
			Exit Do
		End If
		Wait(1)
	Loop
	Reporter.Reportevent micDone,"Browser Busy Function","Seconds :"&seconds
	WriteToEvent("Done" & vbtab & "Browser Busy Function Seconds : " &seconds)
End Function



	''
	' This function DynamicWait is to provide the required wait to the script for the object to appear with a maxTime as its maximum time to idle the script..
	' @author DSTWS
	' @param objParent String specifying the Parent Object
	' @param maxTime number specifying the time duration for which the script can be idle at max.
	' @Modified By DT77734
	' @Modified on: 16 Jul 2009 

Function DynamicWait(objParent,maxTime)
	On error resume Next
	Dim seconds
	seconds=0
	err.clear
	m=objParent.Object.readyState="complete"
	If err.number>0 Then
		Exit Function
	End If
	If objParent.Object.readyState="complete" Then
		err.clear
		Exit Function
	Else
		start=Now
		On Error Resume Next
		Do
			stp=Now
			seconds= datediff("s",start,stp)
			If seconds>maxTime Then
				Exit Do
			End If
		Loop Until objParent.Object.readyState="complete"
	End If
	err.clear
End Function



 	''
	' This function GetObjectParent is to return the parent object name of test object
	' @author DSTWS
	' @param objParent String specifying the Parent Object name
	' @returns the parent object of the test object.
	' @Modified By DT77734
	' @Modified on: 16 Jul 2009 

Function GetObjectParent(objParent)
	Select Case Trim(LCase(objParent))
		Case"objparent1"
			Set GetObjectParent=Browser("title:=Google").Page("title:=Google")
		Case"objparent2"
			Set GetObjectParent=Window("text:=Flight Reservation")
		Case Else
			Set GetObjectParent=Window("text:=.*")
	End Select
End Function





''''''''''''''' HTML CODE'''''''''''''''''''''''''''''''''


''''''''''''''' HTML CODE'''''''''''''''''''''''''''''''''
Public strHeader
Public strTestCases
Public strResTab
Public HTMLContent
Public tcStatus
Public tdWarning
'**********************************************************************************************************************************
' Function Name	:	CreateHTMLResults
' Description	:	This Function creates the result file in HTML format 
' Parameter		:	
' Author		:	DSTWS TA2000 Automation Team
' Creation Date	: 	August, 2011
' Reviewed By 	:	DSTWS TA2000 Automation Team
'
' Modified By	:
' Modified Date	:	
'**********************************************************************************************************************************

Public Function CreateHTMLResults()
	Dim objFileSystemObject, objTextFileObject
	Dim objTempFile, objFolder
	Dim sFileText
	Dim iPos
	
	On Error Resume Next
	Err.Clear

'Creating a folder appended with date and time

	bFinalStatus=True
	tmStartTime=now
	Environment("SlNo")=1
	Environment("stTime")=Time
	Environment("nPassed")=0
	Environment("nFailed")=0
	Environment("nTotal")=0
	Environment("nTotalTC")=0
	dtStartDate=Date
	tmStartTime=Replace(Replace(Replace(tmStartTime,"/","_")," ","_"),":","_")
	dtStartDate=Replace(dtStartDate,"/","_")
	Environment("CurrentResultFile")=Environment("HTMLResultFilePath")
	Environment("CurrentTCSheet")=""

	strHeader = "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.0 Transitional//EN""><HTML><head><title>"&Environment("vegaTestVersion")&" HTMLResult</title></head><BODY><CENTER><Body Style=""background-color:FFFFEE;""/><a href=""http://www.dstworldwideservices.com/index.html""><IMG SRC=""C:\TRAC Web Automation_POC\TAF-M\Documents\header-top.jpg"" BORDER=""0"" ALT=""DSTWS"" align=""Left"" valign=""top"" ></a><H1 align=""Center""> vegaTest HTML Report </H1>"
    strHeader = strHeader & "<TABLE ALIGN=""Left"" BORDER=""0"" WIDTH=100% CELLPADDING=""1"" ><TR BGCOLOR=FFFFEE><TD width=50% align=""left"" valign=""top""><TABLE ALIGN=""Left"" BORDER=""1"" WIDTH=50% CELLPADDING=""1"" ><TR BGCOLOR=Brown><TH colspan=""3"" align=""Center"" valign=""top""><BGCOLOR=Brown> <FONT COLOR=PaleGoldenRod face=""Arial""><B>Test Case Summary </B></FONT></TH><TR BGCOLOR=TAN><TD width=50% align=""left"" valign=""top""><FONT COLOR=DarkOliveGreen face=""Arial""><small>Date Executed: </small></FONT></TD><TD width=50% align=""left"" valign=""top""><FONT COLOR=DarkOliveGreen face=""Arial""><small>"&Date&"</small></FONT></TD></TR>"
	strHeader = strHeader & "<TR BGCOLOR=TAN><TD width=50% align=""left"" valign=""top""><FONT COLOR=DarkOliveGreen face=""Arial""><small>Application Name: </small></FONT></TD><TD width=50% align=""left"" valign=""top""><FONT COLOR=DarkOliveGreen face=""Arial""><small>"&Environment("ApplicationName")&"</small></FONT></TD></TR>"
	strHeader = strHeader & "<TR BGCOLOR=TAN><TD width=50% align=""left"" valign=""top""><FONT COLOR=DarkOliveGreen face=""Arial""><small>Version: </small></FONT></TD><TD width=50% align=""left"" valign=""top""><FONT COLOR=DarkOliveGreen face=""Arial""><small>"&Environment("Version")&"</small></FONT></TD></TR>"
	strHeader = strHeader & "<TR BGCOLOR=TAN><TD width=50% align=""left"" valign=""top""><FONT COLOR=DarkOliveGreen face=""Arial""><small>Executed By: </small></FONT></TD><TD width=50% align=""left"" valign=""top""><FONT COLOR=DarkOliveGreen face=""Arial""><small>"&Environment.Value("UserName")&"</small></FONT></TD></TR>"
	strHeader = strHeader & "<TR BGCOLOR=TAN><TD width=50% align=""left"" valign=""top""><FONT COLOR=DarkOliveGreen face=""Arial""><small>vegaTest  Version: </small></FONT></TD><TD width=50% align=""left"" valign=""top""><FONT COLOR=DarkOliveGreen face=""Arial""><small>"&Environment("vegaTestVersion")&"</small></FONT></TD></TR>"
	strHeader = strHeader & "<TR BGCOLOR=TAN><TD width=50% align=""left"" valign=""top""><FONT COLOR=DarkOliveGreen face=""Arial""><small>Test Start Time: </small></FONT></TD><TD width=50% align=""left"" valign=""top""><FONT COLOR=DarkOliveGreen face=""Arial""><small>"&Environment("stTime")&"</small></FONT></TD><!--TESTSTARTTIME--></TR>"
	strHeader = strHeader & "<TR BGCOLOR=TAN><TD width=50% align=""left"" valign=""top""><FONT COLOR=DarkOliveGreen face=""Arial""><small>Test End Time: </small></FONT></TD><!--stTESTENDTIME--><!--edTESTENDTIME--></TR>"
	strHeader = strHeader & "<TR BGCOLOR=TAN><TD width=50% align=""left"" valign=""top""><FONT COLOR=DarkOliveGreen face=""Arial""><small>Test Duration: </small></FONT></TD><!--stTESTDURATION--><!--edTESTDURATION--></TR>"
	strHeader = strHeader & "<TR BGCOLOR=TAN><TD width=50% align=""left"" valign=""top""><FONT COLOR=DarkOliveGreen face=""Arial""><small>Total Test Cases Executed: </small></FONT></TD><!--stTOTTEST--><TD width=50% align=""left"" valign=""top""><FONT COLOR=DarkOliveGreen face=""Arial""><small>"&Environment("nTotalTC")&"</small></FONT></TD><!--edTOTTEST--></TR>"
	strHeader = strHeader & "<TR BGCOLOR=TAN><TD width=50% align=""left"" valign=""top""><FONT COLOR=DarkOliveGreen face=""Arial""><small>Total Test Steps Passed: </small></FONT></TD><!--stTESTPASSED--><TD width=50% align=""left"" valign=""top""><FONT COLOR=DarkOliveGreen face=""Arial""><small>"&Environment("nPassed")&"</small></FONT></TD><!--edTESTPASSED--></TR>"
	strHeader = strHeader & "<TR BGCOLOR=TAN><TD width=50% align=""left"" valign=""top""><FONT COLOR=DarkOliveGreen face=""Arial""><small>Total Test Steps Failed: </small></FONT></TD><!--stTESTFAILED--><TD width=50% align=""left"" valign=""top""><FONT COLOR=DarkOliveGreen face=""Arial""><small>"&Environment("nFailed")&"</small></FONT></TD><!--edTESTFAILED--></TR>"
	strHeader = strHeader & "<TR BGCOLOR=TAN><TD width=50% align=""left"" valign=""top""><FONT COLOR=DarkOliveGreen face=""Arial""><small>Total Executed Steps: </small></FONT></TD><!--stTOTEXC--><TD width=50% align=""left"" valign=""top""><FONT COLOR=DarkOliveGreen face=""Arial""><small>"&Environment("nTotal")&"</small></FONT></TD><!--edTOTEXC--></TR>"
	''TAF 10.1 new code Start. Modified Trao.	Added 'Total DataRows:' in the summary
	strHeader = strHeader & "<TR BGCOLOR=TAN><TD width=50% align=""left"" valign=""top""><FONT COLOR=DarkOliveGreen face=""Arial""><small>Total DataRows: </small></FONT></TD><!--stTOTROWS--><TD width=50% align=""left"" valign=""top""><FONT COLOR=DarkOliveGreen face=""Arial""><small>"&Environment("nTotal")&"</small></FONT></TD><!--edTOTROWS--></TR></TABLE></TR>"
	' TAF 10.1 new code End
	strTestCases = "<TR BGCOLOR=FFFFEE><TD width=50% align=""left"" valign=""top""><TABLE ALIGN=""Left"" BORDER=""1"" WIDTH=50% CELLPADDING=""1"" ><TR BGCOLOR=Brown><TH colspan=""3"" align=""Center"" valign=""top""><BGCOLOR=Brown> <FONT COLOR=PaleGoldenRod face=""Arial""><B>Test Case(s)</B></FONT></TH><!--TABLEINDEX--></TABLE></Tr>"

	strResTab =  "<TR BGCOLOR=FFFFEE><TD width=50% align=""left"" valign=""top""><TABLE ALIGN=""Center"" BORDER=""1"" WIDTH=100% CELLPADDING=""1"" ><TR BGCOLOR=""Brown""><TH width=2% colspan=""1"" align=""left"" valign=""top""><FONT COLOR=PaleGoldenRod face=""Arial""><small><B>S.No</B></small></FONT></TH><TH width=8% colspan=""1"" align=""left"" valign=""center""><FONT COLOR=PaleGoldenRod face=""Arial""><small>UIName</small></FONT></TH><TH width=7% colspan=""1""align=""left"" valign=""center""><FONT COLOR=PaleGoldenRod face=""Arial""><small> ScreenShot</small></FONT></TH><TH width=20% colspan=""1""align=""left"" valign=""center""><FONT COLOR=PaleGoldenRod face=""Arial""><small><B>Expected Value</B></small></FONT></TH><TH width=20% colspan=""1""align=""left"" valign=""center""><FONT COLOR=PaleGoldenRod face=""Arial""><small><B>Actual Value</B></small></FONT></TH><TH width=20% colspan=""1""align=""left"" valign=""center""><FONT COLOR=PaleGoldenRod face=""Arial""><small><B>Data</B></small></FONT></TH><TH width=5% align=""left"" valign=""center""><FONT COLOR=PaleGoldenRod face=""Arial""><small><B>Status</B></small></FONT></TH><TH width=13% align=""left"" valign=""center""><FONT COLOR=PaleGoldenRod face=""Arial""><small><B>DataRow</B></small></FONT></TH><TH width=5% align=""left"" valign=""center""><FONT COLOR=PaleGoldenRod face=""Arial""><small><B>Step Number</B></small></FONT></TH></TR>"
	strResTab = strResTab & "<TR BGCOLOR=""TAN""><TH height=17 colspan=""9"" align=""left"" valign=""top""><FONT COLOR=black face=""Arial""><small></small></FONT></TH></TR><!--TESTCASENAME--><!--LOGDETAILS--></TABLE></Tr></TABLE>"
	HTMLContent = strHeader & strTestCases & strResTab
		strHeader = Null
		strTestCases = Null
		strResTab = Null
End Function

'*************************************************************************************************************************************************
' Function Name	:	ReportHTMLResults
' Description	:	This function writes the result to the result file
' Parameter	:		This function takes the following parameters
'				strUIName (String) - UIName of the object
'				strData (String) - Data passed for the test case
'				Status (Boolean) - Status of the Step (0 = Pass and 1 = Fail)
'				strExpected (String)- Expected result for the step
'				strActual (String) - Actual Result for the step
'				objParent (Object) - Object on which the operation is performed
' Author		:	DSTWS TA2000 Automation Team
' Creation Date	: 	August, 2011
' Reviewed By 	:	DSTWS TA2000 Automation Team
'
' Modified By	:
' Modified Date	:	
'*********************************************************************************************************************************************************************************

Public Function ReportHTMLResults(strUIName, strExpected, strData, strActual, Status,objParent)
On Error Resume Next
err.Clear

If Status=Empty Then
	Exit Function
End If

   If strUIName="" Then
		strUIName="&nbsp;"
	End If
	If strExpected="" Then
		strExpected="&nbsp;"
	End If
	If strActual="" Then
		strActual="&nbsp;"
	End If
	If strData="" Then
		strData="&nbsp;"
	End If

    Environment("StepNum")=Environment("StepNumber")-1
    If LCase(Trim(Environment("Optional")))="true" And LCase(Trim(Status))="failed" then
		Exit Function
    End If
	If Environment("StepNum")="" Then
		Environment("StepNum")="&nbsp"
    End If
	If LCase(Trim(Status))<>"failed" Then

		If Environment("StepNum")<>0 Then
			Environment("nPassed")=Environment("nPassed")+1
		End If

	Else

		If Environment("StepNum")<>0 Then
			Environment("nFailed")=Environment("nFailed")+1
		End If
	
	End If

	Environment("nTotal")=Environment("nFailed")+Environment("nPassed")
	

	If Environment("CurrentTCSheet")<>Environment("strNewScenario")&" --> "&Environment("TestCaseSheetNameToResults")&" --> "&Environment("TestCaseName") And Environment("CurrentTCSheet")<>Environment("TestCaseName1")&" --> "&Environment("TestCaseName")Then
		'<!--edWarning-->
		'<!--stWarning-->
		HTMLContent = Replace(HTMLContent,"<!--edWarning-->","")
		HTMLContent = Replace(HTMLContent,"<!--stWarning-->","")
		If Not IsNull(Environment("strNewScenario")) Then
			Environment("CurrentTCSheet1")=Environment("strNewScenario")&" --> "&Environment("TestCaseSheetNameToResults")
			Environment("CurrentTCSheet")=Environment("strNewScenario")&" --> "&Environment("TestCaseSheetNameToResults")&" --> "&Environment("TestCaseName")
		Else
			Environment("CurrentTCSheet1")=Environment("TestCaseName1")
        	Environment("CurrentTCSheet")=Environment("TestCaseName1")&" --> "&Environment("TestCaseName")
		End If
		
		tcStatus = 0
		Environment("nTotalTC")=Environment("nTotalTC")+1
		HTMLContent = Replace(HTMLContent,"<!--stTCPSFL-->","")
		HTMLContent = Replace(HTMLContent,"<!--edTCPSFL-->","")
		TFileText=HTMLContent
		TPos=instr(1,TFileText,"<!--TABLEINDEX-->",vbTextCompare)
		If TPos > 0 Then
			TFileText1=mid(TFileText,1,TPos-1)
			TFileText2=mid(TFileText,TPos)
		End If
		HTMLContent = TFileText1 & "<TR BGCOLOR=""PaleGoldenRod""><TD width=50% align=""left"" valign=""top""><FONT COLOR=Brown face=""Arial""><small><B><a href=""#" &Environment("CurrentTCSheet")&Chr(34)&">"&Environment("CurrentTCSheet1")&"</a> </B></small></FONT></TD><!--stTCPSFL--><TD width=50% align=""left"" valign=""top""><FONT COLOR=DarkOliveGreen face=""Arial""><small><B>Passed</B></small></FONT></TD><!--edTCPSFL--></TR>" & TFileText2


		NFileText=HTMLContent
		NPos=instr(1,NFileText,"<!--LOGDETAILS-->",vbTextCompare)
		If NPos > 0 Then
			NFileText1=mid(NFileText,1,NPos-1)
			NFileText2=mid(NFileText,NPos)
		End If

		If Environment("DataFromDataSheet") = "False" Then
			HTMLContent = NFileText1 & "<TR BGCOLOR=""PaleGoldenRod""><TH width=50% height=17 colspan=""4"" align=""left"" valign=""top""><FONT COLOR=Brown face=""Arial""><small><B><a name="&Chr(34)&Environment("CurrentTCSheet")&Chr(34)&">"&Environment("CurrentTCSheet")&"</a></B></small></FONT></TH><!--stWarning--><TH width=50% height=17 colspan=""5"" align=""left"" valign=""top""><FONT COLOR=Red face=""Arial""><small><B>Warning : Test Data given directly in test case.</a></B></small></FONT></TH><!--edWarning--></TR>" & NFileText2
			tdWarning = 0
		Else
			If (Environment("DataFromDataSheet") = "True" and Environment("Keywords")="") then
				If Environment("RunFromQC")  Then
					HTMLConten = NFileText1 & "<TR BGCOLOR=""PaleGoldenRod""><TH width=50% height=17 colspan=""4"" align=""left"" valign=""top""><FONT COLOR=Brown face=""Arial""><small><B><a name="&Chr(34)&Environment("CurrentTCSheet")&Chr(34)&">"&Environment("CurrentTCSheet")&"</a></B></small></FONT></TH><!--stWarning--><TH width=50% height=17 colspan=""5"" align=""left"" valign=""top""><FONT COLOR=Blue face=""Arial""><small><B> ""TestDataPath =" & 	Environment("QCTestDataFilePath") & " Sheet=" & Environment("TestDataSheetName") & "DataRow = Default Row""</a></B></small></FONT></TH><!--edWarning--></TR>" & NFileText2
					tdWarning = 1
				else
					HTMLContent = NFileText1& "<TR BGCOLOR=""PaleGoldenRod""><TH width=50% height=17 colspan=""4"" align=""left"" valign=""top""><FONT COLOR=Brown face=""Arial""><small><B><a name="&Chr(34)&Environment("CurrentTCSheet")&Chr(34)&">"&Environment("CurrentTCSheet")&"</a></B></small></FONT></TH><!--stWarning--><TH width=50% height=17 colspan=""5"" align=""left"" valign=""top""><FONT COLOR=Blue face=""Arial""><small><B> ""TestDataPath =" & Environment("TestDataPathToReport") & "Sheet=" & Environment("TestDataSheetName") & "DataRow = Default Row""</a></B></small></FONT></TH><!--edWarning--></TR>" & NFileText2
					tdWarning = 1
				End If
			Else  
				If  Environment("RunFromQC") Then
					HTMLContent = NFileText1 & "<TR BGCOLOR=""PaleGoldenRod""><TH width=50% height=17 colspan=""4"" align=""left"" valign=""top""><FONT COLOR=Brown face=""Arial""><small><B><a name="&Chr(34)&Environment("CurrentTCSheet")&Chr(34)&">"&Environment("CurrentTCSheet")&"</a></B></small></FONT></TH><!--stWarning--><TH width=50% height=17 colspan=""5"" align=""left"" valign=""top""><FONT COLOR=Blue face=""Arial""><small><B>""TestDataPath =" &   	Environment("QCTestDataFilePath") & " Sheet=" & Environment("TestDataSheetName") & " DataRow =" & Environment("Keywords")  &"</a></B></small></FONT></TH><!--edWarning--></TR>" & NFileText2
					tdWarning = 1
				else
					HTMLContent = NFileText1 & "<TR BGCOLOR=""PaleGoldenRod""><TH width=50% height=17 colspan=""4"" align=""left"" valign=""top""><FONT COLOR=Brown face=""Arial""><small><B><a name="&Chr(34)&Environment("CurrentTCSheet")&Chr(34)&">"&Environment("CurrentTCSheet")&"</a></B></small></FONT></TH><!--stWarning--><TH width=50% height=17 colspan=""5"" align=""left"" valign=""top""><FONT COLOR=Blue face=""Arial""><small><B>""TestDataPath =" & Environment("TestDataPathToReport") & " Sheet=" & Environment("TestDataSheetName") & " DataRow =" & Environment("Keywords")  &"</a></B></small></FONT></TH><!--edWarning--></TR>" & NFileText2
					tdWarning = 1
				End If
			End if
	End If 
	End If

	If LCase(Trim(Status))="failed" And tcStatus=0  Then
		tcStatus=1
		TCPSFAILText=HTMLContent
		stTCPSFLPos=instr(1,TCPSFAILText,"<!--stTCPSFL-->",vbTextCompare)
		edTCPSFLPos=instr(1,TCPSFAILText,"<!--edTCPSFL-->",vbTextCompare)
		If edTCPSFLPos>0 Then
			TCPSFLText1=mid(TCPSFAILText,1,stTCPSFLPos-1)
			TCPSFLText2=mid(TCPSFAILText,edTCPSFLPos+15)
		End If
		HTMLContent = TCPSFLText1 & "<TD width=50% align=""left"" valign=""top""><FONT COLOR=Brown face=""Arial""><small><B>Failed</B></small></FONT></TD><!--stTCPSFL--><!--edTCPSFL-->" & TCPSFLText2
	End If
    


'	Environment("CurrentTCSheet")=Environment("strEnvironment")&" --> "&Environment("TestSheetName")
		If Not IsNull(Environment("strNewScenario")) Then
			Environment("CurrentTCSheet1")=Environment("strNewScenario")&" --> "&Environment("TestCaseSheetNameToResults")
			Environment("CurrentTCSheet")=Environment("strNewScenario")&" --> "&Environment("TestCaseSheetNameToResults")&" --> "&Environment("TestCaseName")
		Else
			Environment("CurrentTCSheet1")=Environment("TestCaseName1")
        	Environment("CurrentTCSheet")=Environment("TestCaseName1")&" --> "&Environment("TestCaseName")
		End If

	If Environment("DataFromDataSheet") = "True" And tdWarning <> 1 Then
		WarningText=HTMLContent
		stWarPos=instr(1,WarningText,"<!--stWarning-->",vbTextCompare)
		edWarPos=instr(1,WarningText,"<!--edWarning-->",vbTextCompare)
		If edWarPos > 0 Then
			WarningText1 = mid(WarningText,1,stWarPos-1)
			WarningText2 = mid(WarningText,edWarPos+16)
		
    If (Environment("DataFromDataSheet") = "True" and Environment("Keywords")="") then
		If Environment("RunFromQC")  Then
			HTMLContent = WarningText1 & "<TH width=50% height=17 colspan=""5"" align=""left"" valign=""top""><FONT COLOR=Blue face=""Arial""><small><B> ""TestDataPath =" & 	Environment("QCTestDataFilePath") & " Sheet=" & Environment("TestDataSheetName") & "DataRow = Default Row""</a></B></small></FONT></TH>" & WarningText2
			else
			HTMLContent = WarningText1 & "<TH width=50% height=17 colspan=""5"" align=""left"" valign=""top""><FONT COLOR=Blue face=""Arial""><small><B> ""TestDataPath =" & Environment("TestDataPathToReport") & "Sheet=" & Environment("TestDataSheetName") & "DataRow = Default Row""</a></B></small></FONT></TH>" & WarningText2
		End If
    
   Else  
   If  Environment("RunFromQC") Then
	   HTMLContent = WarningText1 & "<TH width=50% height=17 colspan=""5"" align=""left"" valign=""top""><FONT COLOR=Blue face=""Arial""><small><B>""TestDataPath =" &   	Environment("QCTestDataFilePath") & " Sheet=" & Environment("TestDataSheetName") & " DataRow =" & Environment("Keywords")  &"</a></B></small></FONT></TH>" & WarningText2
	   else
		HTMLContent = WarningText1 & "<TH width=50% height=17 colspan=""5"" align=""left"" valign=""top""><FONT COLOR=Blue face=""Arial""><small><B>""TestDataPath =" & Environment("TestDataPathToReport") & " Sheet=" & Environment("TestDataSheetName") & " DataRow =" & Environment("Keywords")  &"</a></B></small></FONT></TH>" & WarningText2
   End If
   End If
  End if
  End If 

	If HTMLContent<>"" Then

		sFileText=HTMLContent
		iPos=instr(1,sFileText,"<!--LOGDETAILS-->",vbTextCompare)
		If iPos > 0 Then
			sFileText = mid(sFileText,1,iPos-1)
		End If
		HTMLContent= sFileText
		
	End If

	bFinalStatus=True
	tmStartTime=now
	dtStartDate=Date
	tmStartTime=Replace(Replace(Replace(tmStartTime,"/","_")," ","_"),":","_")
	dtStartDate=Replace(dtStartDate,"/","_")
'	nStepNum = Environment("StepNum")
	nStepNum=Environment("SlNo")
	If strUIName="KeywordRepositoryPath" Or strUIName="TestDataPath" Or strUIName="TestDataSheets" Then
		nStepNum="&nbsp;"
	End If

	HTMLContent= HTMLContent & "<TR BGCOLOR=White>"
	If LCase(Trim(Status))="failed" Then
		HTMLContent= HTMLContent & "<TD width=2% colspan=""1"" align=""left"" valign=""top""><FONT COLOR=Red face=""Arial""><small>"& nStepNum &"</small></FONT></TD>" 
	Else
		HTMLContent= HTMLContent & "<TD width=2% colspan=""1"" align=""left"" valign=""top""><FONT COLOR=Green face=""Arial""><small>"& nStepNum &"</small></FONT></TD>" 
	End If
	If LCase(Trim(Status))="failed" Then
		HTMLContent= HTMLContent & "<TD width=8% colspan=""1"" align=""left"" valign=""top""><FONT COLOR=Red face=""Arial""><small>"& strUIName &"</small></FONT></TD>" 
	Else
		HTMLContent= HTMLContent & "<TD width=8% colspan=""1"" align=""left"" valign=""top""><FONT COLOR=Green face=""Arial""><small>"& strUIName &"</small></FONT></TD>" 
	End If

	If (LCase(Trim(Status))="failed" Or LCase(Trim(Status))="screenshot") And Environment("StepNum")<>0 Then
		If Environment("ScreenShotPath")<>"" Then
			ScreenShotPath=Environment("ScreenShotPath")&"\screenshot"&tmStartTime&"_"&Environment("UserName")&".png"
        	Desktop.CaptureBitmap ScreenShotPath,true
			HTMLContent= HTMLContent & "<TD width=7% colspan=""1"" align=""left"" valign=""top""><FONT COLOR=Blue face=""Arial""><small><a href="&Chr(34)&ScreenShotPath&Chr(34)&"> ScreenShot</a></small></FONT></TD>" 
		Else
			HTMLContent= HTMLContent & "<TD width=7% colspan=""1"" align=""left"" valign=""top""><FONT COLOR=Red face=""Arial""><small>ScreenShot</small></FONT></TD>" 
		End If		
	Else
		HTMLContent= HTMLContent & "<TD width=7% colspan=""1"" align=""left"" valign=""top""><FONT COLOR=#D0D0D0 face=""Arial""><small>ScreenShot</small></FONT></TD>" 
	End If


	If LCase(Trim(Status))="failed" Then
		HTMLContent= HTMLContent & "<TD width=20% colspan=""1"" align=""left"" valign=""top""><FONT COLOR=Red face=""Arial""><small>"& strExpected &"</small></FONT></TD>" 
	Else
		HTMLContent= HTMLContent & "<TD width=20% colspan=""1"" align=""left"" valign=""top""><FONT COLOR=Green face=""Arial""><small>"& strExpected &"</small></FONT></TD>" 
	End If
	If LCase(Trim(Status))="failed" Then
		HTMLContent= HTMLContent & "<TD width=20% colspan=""1"" align=""left"" valign=""top""><FONT COLOR=Red face=""Arial""><small>"& strActual &"</small></FONT></TD>" 
	Else
		HTMLContent= HTMLContent & "<TD width=20% colspan=""1"" align=""left"" valign=""top""><FONT COLOR=Green face=""Arial""><small>"& strActual &"</small></FONT></TD>" 
	End If
	If LCase(Trim(Status))="failed" Then
		HTMLContent= HTMLContent & "<TD width=20% colspan=""1"" align=""left"" valign=""top""><FONT COLOR=Red face=""Arial""><small>"& strData &"</small></FONT></TD>"
	Else
		HTMLContent= HTMLContent & "<TD width=20% colspan=""1"" align=""left"" valign=""top""><FONT COLOR=Green face=""Arial""><small>"& strData &"</small></FONT></TD>"
	End If
	If LCase(Trim(Status))="failed" Then
		HTMLContent= HTMLContent & "<TD width=5% colspan=""1"" align=""left"" valign=""top""><FONT COLOR=Red face=""Arial""><small>"& Status &"</small></FONT></TD>"
	Else
		HTMLContent= HTMLContent & "<TD width=5% colspan=""1"" align=""left"" valign=""top""><FONT COLOR=Green face=""Arial""><small>"& Status &"</small></FONT></TD>"
	End If
	If LCase(Trim(Status))="failed" Then
		HTMLContent= HTMLContent & "<TD width=13% colspan=""1"" align=""left"" valign=""top""><FONT COLOR=Red face=""Arial""><small>"& Environment("ExecelResultsKeywords") &"</small></FONT></TD>"
	Else
		HTMLContent= HTMLContent & "<TD width=13% colspan=""1"" align=""left"" valign=""top""><FONT COLOR=Green face=""Arial""><small>"& Environment("ExecelResultsKeywords") &"</small></FONT></TD>"
	End If
	If LCase(Trim(Status))="failed" Then
		HTMLContent= HTMLContent & "<TD width=5% colspan=""1"" align=""left"" valign=""top""><FONT COLOR=Red face=""Arial""><small>"& Environment("StepNumber") &"</small></FONT></TD>"
	Else
		HTMLContent= HTMLContent & "<TD width=5% colspan=""1"" align=""left"" valign=""top""><FONT COLOR=Green face=""Arial""><small>"& Environment("StepNumber") &"</small></FONT></TD>"
	End If

		HTMLContent= HTMLContent & "</TR>"
		HTMLContent= HTMLContent & "<!--LOGDETAILS--></Tr></TABLE>"
		Environment("SlNo")=Environment("SlNo")+1
		Environment("edTime")=Time
End Function

'**********************************************************************************************************************************
' Function Name	:	TimeDifference
' Description	:	This function returns the time difference
' Parameter	:		This function takes the following parameters
'				StartTime  - Start time
'				EndTime - End Time
' Author		:	DSTWS TA2000 Automation Team
' Creation Date	: 	August, 2011
' Reviewed By 	:	DSTWS TA2000 Automation Team
' Modified By	:
' Modified Date	:	
'*********************************************************************************************************************************************************************************

Function TimeDifference(StartTime,EndTime)    
   
	StartHour = Hour(StartTime)       
	StartMin =  Minute(StartTime)       
	StartSec = Second(StartTime)   
	   
	EndHour = Hour(EndTime)       
	EndMin = Minute(EndTime)       
	EndSec = Second(EndTime)       
   
   	StartingSeconds = (StartSec + (StartMin * 60) + ((StartHour * 60)*60))       
   	EndingSeconds = (EndSec + (EndMin * 60) + ((EndHour * 60)*60))       
   	TimeInSec = EndingSeconds - StartingSeconds
   
    nHours=Int(TimeInSec/3600)
	TimeInSec=TimeInSec-(3600*nHours)
    nMinutes=Int(TimeInSec/60)
	nSeconds=TimeInSec-(60*nMinutes)
	If Not nHours >=9 Then
		nHours=0&nHours
	End If
	If Not nMinutes >=9 Then
		nMinutes=0&nMinutes
	End If
	If Not nSeconds >=9 Then
		nSeconds=0&nSeconds
	End If
   	
   	TimeDifference=nHours&":"&nMinutes&":"&nSeconds
   	
End Function 

Function HTMLSummary()
		On Error Resume Next
		Err.Clear
		psFileText=HTMLContent
		stpsPos=instr(1,psFileText,"<!--stTESTPASSED-->",vbTextCompare)
		edpsPos=instr(1,psFileText,"<!--edTESTPASSED-->",vbTextCompare)
		If stpsPos > 0 And edpsPos>0 Then
			stpsFileText=mid(psFileText,1,stpsPos-1)
			edpsFileText=mid(psFileText,edpsPos)
		End If
        HTMLContent = stpsFileText  & "<!--stTESTPASSED--><TD width=50% align=""left"" valign=""top""><FONT COLOR=DarkOliveGreen face=""Arial""><small>"&Environment("nPassed")&"</small></FONT></TD>" & edpsFileText

		fdFileText=HTMLContent
		stfdPos=instr(1,fdFileText,"<!--stTESTFAILED-->",vbTextCompare)
		edfdPos=instr(1,fdFileText,"<!--edTESTFAILED-->",vbTextCompare)
		If stfdPos > 0 And edfdPos>0 Then
			stfdFileText=mid(fdFileText,1,stfdPos-1)
			edfdFileText=mid(fdFileText,edfdPos)
		End If
		HTMLContent= stfdFileText & "<!--stTESTFAILED--><TD width=50% align=""left"" valign=""top""><FONT COLOR=DarkOliveGreen face=""Arial""><small>"&Environment("nFailed")&"</small></FONT></TD>" & edfdFileText

		teFileText=HTMLContent
		testPos=instr(1,teFileText,"<!--stTOTEXC-->",vbTextCompare)
		teedPos=instr(1,teFileText,"<!--edTOTEXC-->",vbTextCompare)
		If testPos > 0 And teedPos>0 Then
			stteFileText=mid(teFileText,1,testPos-1)
			edteFileText=mid(teFileText,teedPos)
		End If
		HTMLContent= stteFileText & "<!--stTOTEXC--><TD width=50% align=""left"" valign=""top""><FONT COLOR=DarkOliveGreen face=""Arial""><small>"&Environment("nTotal")&"</small></FONT></TD>" & edteFileText

		tcsFileText=HTMLContent
		tcsstPos=instr(1,tcsFileText,"<!--stTOTTEST-->",vbTextCompare)
		tcsedPos=instr(1,tcsFileText,"<!--edTOTTEST-->",vbTextCompare)
		If tcsstPos > 0 And tcsedPos>0 Then
			sttcsFileText=mid(tcsFileText,1,tcsstPos-1)
			edtcsFileText=mid(tcsFileText,tcsedPos)
		End If
		HTMLContent= sttcsFileText & "<!--stTOTTEST--><TD width=50% align=""left"" valign=""top""><FONT COLOR=DarkOliveGreen face=""Arial""><small>"&Environment("nTotalTC")&"</small></FONT></TD>" & edtcsFileText

		stedFileText=HTMLContent
		stPos=instr(1,stedFileText,"<!--stTESTENDTIME-->",vbTextCompare)
		edPos=instr(1,stedFileText,"<!--edTESTENDTIME-->",vbTextCompare)
		If stPos > 0 And edPos>0 Then
			stFileText=mid(stedFileText,1,stPos-1)
			edFileText=mid(stedFileText,edPos)
		End If
		HTMLContent= stFileText & "<!--stTESTENDTIME--><TD width=50% align=""left"" valign=""top""><FONT COLOR=DarkOliveGreen face=""Arial""><small>"&Environment("edTime")&"</small></FONT></TD>" & edFileText

		duFileText=HTMLContent
		sduPos=instr(1,duFileText,"<!--stTESTDURATION-->",vbTextCompare)
		eduPos=instr(1,duFileText,"<!--edTESTDURATION-->",vbTextCompare)
		If sduPos > 0 And eduPos>0 Then
			sduFileText=mid(duFileText,1,sduPos-1)
			eduFileText=mid(duFileText,eduPos)
		End If
		Environment("htmlDuration")= TimeDifference(Environment("stTime"),Environment("edTime"))
		HTMLContent= sduFileText & "<!--stTESTDURATION--><TD width=50% align=""left"" valign=""top""><FONT COLOR=DarkOliveGreen face=""Arial""><small>"&Environment("htmlDuration")&"</small></FONT></TD>" & eduFileText


		'TAF 10.1 new code Start.
		duFileText1=HTMLContent
		sduPos1=instr(1,duFileText1,"<!--stTOTROWS-->",vbTextCompare)
		eduPos1=instr(1,duFileText1,"<!--edTOTROWS-->",vbTextCompare)
		If sduPos1 > 0 And eduPos1>0 Then
			sduFileText1=mid(duFileText1,1,sduPos1-1)
			eduFileText1=mid(duFileText1,eduPos1)
		End If
		Environment("NoOfDatarows")
		HTMLContent= sduFileText1 & "<!--stTOTROWS--><TD width=50% align=""left"" valign=""top""><FONT COLOR=DarkOliveGreen face=""Arial""><small>"&Environment("NoOfDatarows")&"</small></FONT></TD>" & eduFileText1
		'TAF 10.1 new code End.

End Function

'TAF 10.1 new code Start
Public Sub SaveHTML()
								Call ReportHTMLResults(strUIName, strExpected, strData, strActual, Status,objParent)
								Set objFileSystemObject = CreateObject("Scripting.FileSystemObject")
								Set objTextFileObject = objFileSystemObject.CreateTextFile(Environment("CurrentResultFile"),True)
								HTMLSummary()
					
								objTextFileObject.Write HTMLContent
					'			Set IE=createobject("InternetExplorer.Application")
					'			IE.visible=True
					'			IE.navigate(Environment("CurrentResultFile"))
								Set objFileSystemObject=Nothing
End Sub
'TAF 10.1 new code End
