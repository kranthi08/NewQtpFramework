''
' This Libary is to work with Utility functions.
' @author DSTWS
' @Version 7.0.7


	''
	' This function is to verify to run the QTP test from QC or not
	' @author DSTWS
	' @Modified By DT77709
	' @Modified on: 10 Oct 2009 

Function VerifyTestRunStatus()
	If Lcase(Trim(Environemnt("TestRunStatus"))) = "don'trun"Then
        	Reporter.ReportEvent micWarning,"Verifying the Staus of the Test to execute", "Run Status Value is :[" &Environemnt("TestRunStatus")& "]"  & "  --- Existinmg the Test"
			ExitTest
	End If
End Function



	''
	' This function is to launch the application.
	' @author DSTWS
	' @param strApppath specifies the URL of the application.
	' @Modified By DT77734
	' @Modified on: 16 Jul 2009 
    
Function LaunchApplication(strApppath)
	SystemUtil.Run strApppath
'	WriteToScript("SystemUtil.Run "& strApppath)
	Reporter.reportevent micDone,"Launched the application with the strApppath :"&strApppath,"Launched the application with the strApppath:"&strApppath
	Call ReportResult("Launch Application", "Application Should be launched",strApppath , "Application is  launched" , "Passed" ,objParent)
	'Call ReportResultNonTAF("Launch Application", "Application Should be launched",strApppath , "Application is  launched" , "Passed" ,objParent)
End Function



''
' This Libary is to work with Utility functions.
' @author DSTWS

	''
	' This function is to launch the application.
	' @author DT77742
	' @param strApppath specifies the URL of the application.
    
Function OpenUrl(strURL)
   Dim IE
   On error resume next
	   Set IE=createobject("InternetExplorer.Application")
	   IE.visible=True
	   IE.navigate(strURL)
	   systemUtil.Run "iexplore.exe",strURL
	   Reporter.reportevent micDone,"Opened the URL"&strURL,"Launched the application with the strApppath:"&strURL
	  Call ReportResult("Launch Application", "Application is launched",strURL , "Application is  launched" , "Passed" ,objParent)
	  OpenUrl=True
End Function



''
	' This function is to wait if internet explorer is busy.
	' @author DSTWS
	' @param objBrowser is an Object specifying the name of the browser.
	' @Todo This function is not updated in "DecideActivity.vbs"
	' @Modified By DT77734
	' @Modified on: 16 Jul 2009 
	
Function WaitIfIEBusy(objBrowser)
	On Error Resume Next
	Dim seconds
	seconds=0
	start=Now
	Do While objBrowser.object.Busy=True
		stp=Now
		seconds= datediff("s",start,stp)
		If seconds>180 Then
			Exit Do
		End If
		Wait(1)
	Loop
	Reporter.Reportevent micDone,"Browser Busy Function","Seconds :"&seconds
End Function



''
	' This function will make the QTP to wait till the page loads.
	' @author DSTWS
	' @param  objParent Object specifying the Parent Object
	' @param intMaxTime is a Integer specifying the maximum time to wait.
	' @Todo This function is not updated in "DecideActivity.vbs"
	' @Modified By DT77734
	' @Modified on: 16 Jul 2009 

Function DynamicWait(objParent,intMaxTime)
	Set objBrowser=Browser("name:=ACE.*")
	Call WaitIfIEBusy(objBrowser)
	On Error Resume Next
	Dim seconds
	seconds=0
	If objParent.object.readyState="complete" Then
		Exit Function
	Else
		start=Now
		On Error Resume Next
		Do
			stp=Now
			seconds= datediff("s",start,stp)
			If seconds>intMaxTime Then
				Exit Do
			End If
		Loop Until objParent.object.readyState="complete"
	End If
	On Error GoTo 0
End Function



''
	' This function simulates button press on the browser.
	' @author DSTWS
	' @param  strKeyValue String specifying key value.
	' @Modified By DT77734
	' @Modified on: 16 Jul 2009 

Function SendKey(strKeyValue)
	Wait(1)
	Set WshShell = Createobject("WScript.Shell")
	WshShell.SendKeys strKeyValue
	Reporter.ReportEvent micDone,"Press"&strKeyValue&" to toggle","Key pressed to toggle."
	SendKey="StepPass"
End Function



''
	' This function writes the Run Time Errors to the result file on the last row
	' @author DSTWS
	' @param objobject is object specifying the object in QTP for which the run time error occured
	' @param strMethod is string specifying the method of object for which the run time error occured
	' @param arrArguments is an array specifying the Arguments used in Method
	' @param intRetVal is an Integer Value returned by Method
	' @Todo This function is not updated in "DecideActivity.vbs"
	' @Modified By DT77734
	' @Modified on: 16 Jul 2009 

Function ReportRunError(objObject, strMethod, arrArguments, intRetVal)
	strUIName="Run Error Occured for:"& objObject.ToString()
	strExpected= strMethod
	If IsArray(arrArguments) Then
		strData=Join(arrArguments,";")
	else
		strData=arrArguments
	End If
	strStatus="Failed"
	set objParent=Desktop
	strActual= DescribeResult(intRetVal)
	Call ReportResult (strUIName, strExpected, strData, strActual, strStatus,objParent)
End Function


' This function is to Execute the prerequisite script
' @author Rajesh Kumar Tatavarthi
' @return VOID
'@modified by Sreenu Babu on 20 Nov 2009 

Function ExecutePreReqScript(strPrerequisites)
arrPrerequisites = Split(strPrerequisites,";")
For i= lbound(arrPrerequisites) to ubound(arrPrerequisites)
Select Case Trim(Lcase(arrPrerequisites(i)))
Case "tafbasestate"
	
'call loadvariablesfromsourcefile(Environment("Configuration"),Environment("SourceFile"),Environment("DatabaseName"))
Case "printmsg"
Print "message-Testing base sate Concept"
Case else
Reporter.ReportEvent micFail,"Running the Base state function:" & Trim(Lcase(arrPrerequisites(i))),"Base state function not defined"
WriteToEvent("Fail" & vbtab & "Running the Base state function:" & Trim(Lcase(arrPrerequisites(i))) & "--- Base state function not defined")
End Select
Next
End Function

Public function waitscreencapture(maxTime)	'To get the screenshot for the wait statement
   Wait maxTime
   Reporter.ReportEvent micPass,"Wait Time "&maxTime,"Waited time: "&maxTime
   If  UCASE(Environment("captureScreenshot"))="YES" Then
	   Call ReportResult  ("Wait Time "&maxTime, "Wait Time "&maxTime, "", "", "Screenshot","")
   End If
    	
End Function
