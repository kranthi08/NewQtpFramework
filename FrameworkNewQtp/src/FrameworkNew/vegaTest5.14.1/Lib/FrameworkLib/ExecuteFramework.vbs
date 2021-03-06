''
' @# author DSTWS
' @# Version 7.0.7a
' @# Version 7.0.8 Revision History : Added "kill Process function for Production execution purpose
' @# Version 7.0.9 -Modified by dt77471, Added Environment("captureScreenshot") variable for wait screen capture
' @# Version 7.0.9A Keys added

	''
	' This function 'DecideActivity1.2 has a new feature that captures the AUT based on the user's input from the test case spread sheet. This is a new functionality
	'The flag value is passed from StartFramework class located in StartFramework1.2vbs
	'Based on the flag value the class's function CaptureScreen calls a function that instantiates SetQTP class and also passes the strParentObject object. With the object returned the method that captures'the AUT screen is invoked.   
	' @author DSTWS
	' @param appdataSource String specifying the AppMap excel sheet pah
	' @param UIName String specifying the UI Name of the object
	' @param functionName String Specifying name of the Activity to perform on the UI object.
	' @param ExpectedResult String Specifying the expected result
	' @param stepResultVar String is an optional and non mandatory field
	' @param StopEvent String Specifying the sevearity with which the exceptional should be handled example Minor, Normal or Showstopper.
	' @param Param1 String Specifying the Test Data from the excel sheet
	' @param Param2 String is an optional and non mandatory field
	' @param Param3 String is an optional and non mandatory field
	' @param Param4 String is an optional and non mandatory field
	' @param Param5 String is an optional and non mandatory field
	' @param captureScreen String specifying whether the execution screen to capture, ex: 'Yes' or 'No'
	' @Modified By DT77734
	' @Modified on: 5th Aug 2009 

Function ExecuteFramework(appdataSource,UIName,functionName,ExpectedResult,stepResultVar,StopEvent,Param1,Param2,Param3,Param4,Param5,Param6,Param7,Param8,Param9,Param10,captureScreen)
	Dim customReport
	Dim returnResult 'This variable will store the return values of the function
	addResult="no"  'The flag is set to add result to the dictionary object. Only when the flag is set to yes'Results' will be added.
	'============Call to function AccessFileAsDB================
	setVariables=SetVars(Param1,Param2,Param3,Param4,Param5,Param6,Param7,Param8,Param9,Param10)
	Param1=setVariables(0,0)
	Param2=setVariables(0,1)
	Param3=setVariables(0,2)
	Param4=setVariables(0,3)
	Param5=setVariables(0,4)
	Param6=setVariables(0,5)
	Param7=setVariables(0,6)
	Param8=setVariables(0,7)
	Param9=setVariables(0,8)
	Param10=setVariables(0,9)

Environment("CurrentStepStatus")= "Passed" 

If  LCase(Trim(Environment("DecideActivtyVBS"))) = "true" then
 Environment.Value("captureScreenshot")=captureScreen				'7.0.8 a- Added environment variable to make use in utility function for wait activity
	If LCase(Environment("ParameterCancatinationRequired"))="yes" Then
		If Param1= "~" Then 
				Param1 = Param2 & Param3 & Param4 & Param5 & Param6 & Param7 & Param8 & Param9 & Param10				
		End If
	End If
	If  lcase(trim(Param1))="n/a"or  lcase(trim(Param2)) ="n/a" or  lcase(trim(Param3)) ="n/a" or lcase(trim(Param4)) ="n/a" or lcase(trim(Param5)) ="n/a" or lcase(trim(Param6)) ="n/a" or lcase(trim(Param7)) ="n/a" or lcase(trim(Param8)) ="n/a" or lcase(trim(Param9))="n/a" or lcase(trim(Param10))= "n/a" Then
				Exit Function
   	End If
	Select Case Len(Trim(UIName))
		Case 0
			'********************************************************* IF-ELSE and GOTO Start by Srikanth *****************************************************************************
			If Environment.value("reusescript") = False Then
				tccrntrow = Environment.value("tcrow")
				ActiveSheet = "TestScriptSheet"			
			Else
				tccrntrow = Environment.value("reusetcrow")
				ActiveSheet = "Reusable"
			End If
			Select Case UCase(functionName)
			
			Case "IF"
				Environment.value("elsecondition") = False
				Call ReportResult  ("", "IF","","IF condition started " , "Passed" ,"")
				If Evaluate(Param1)=False Then
					Environment.value("elsecondition") = True
					Do
						tccrntrow = tccrntrow+1
						k = tccrntrow
						DataTable.GetSheet(ActiveSheet).SetCurrentRow(k)
						functionName=DataTable.GetSheet(ActiveSheet).GetParameter("Activity")
						functionName=Cstr(functionName)
						If LCase(Trim(functionName))="endofrow" or LCase(Trim(functionName))="eof"  or LCase(Trim(functionName))="x" Then
							Call ReportResult  ("", functionName,"","ENDIF condition not found in the test case for the IF condition " , "Failed" ,"")
							Exit Do
						End If
					Loop Until UCase(functionName)="ENDIF" or UCase(functionName)="ELSE"
					
					If Environment.value("reusescript") = False Then
						Environment.value("tcrow") = tccrntrow -1
					Else
						Environment.value("reusetcrow") = tccrntrow -1
					End If				
				End If
			Case "ELSE"
				k = tccrntrow
				Call ReportResult  ("", "ELSE","","ELSE condition started " , "Passed" ,"")
				If Environment.value("elsecondition") = True Then
					Do
						DataTable.GetSheet(ActiveSheet).SetCurrentRow(k)
						functionName=DataTable.GetSheet(ActiveSheet).GetParameter("Activity")
						functionName=Cstr(functionName)
						If LCase(Trim(functionName))="endofrow" or LCase(Trim(functionName))="eof"  or LCase(Trim(functionName))="x" Then
							Call ReportResult  ("", functionName,"","ENDIF condition not found in the test case for the IF condition " , "Failed" ,"")
							Exit Do
						End If
						k=k+1
					Loop Until UCase(functionName)="ENDIF"		
				Else
					Do
						tccrntrow=tccrntrow+1
						k = tccrntrow
						DataTable.GetSheet(ActiveSheet).SetCurrentRow(k)
						functionName=DataTable.GetSheet(ActiveSheet).GetParameter("Activity")
						functionName=Cstr(functionName)
						If LCase(Trim(functionName))="endofrow" or LCase(Trim(functionName))="eof"  or LCase(Trim(functionName))="x" Then
							Call ReportResult  ("", functionName,"","ENDIF condition not found in the test case for the IF condition " , "Failed" ,"")
							Exit Do
						End If						
					Loop Until UCase(functionName)="ENDIF"					
				End If
				
				If Environment.value("reusescript") = False Then
					Environment.value("tcrow") = tccrntrow
				Else
					Environment.value("reusetcrow") = tccrntrow
				End If
			Case "ENDIF"
				Call ReportResult  ("", "ENDIF","","IF-ELSE-ENDIF condition completed " , "Passed" ,"")				
			Case "GOTO"
				Call ReportResult  ("", "GOTO","","GOTO condition initiated to  - "&Param1 , "Passed" ,"")												
				Do
						tccrntrow=tccrntrow+1
						k = tccrntrow
						DataTable.GetSheet(ActiveSheet).SetCurrentRow(k)
						functionName=DataTable.GetSheet(ActiveSheet).GetParameter("Activity")
						functionName=Cstr(functionName)
						If LCase(Trim(functionName))="endofrow" or LCase(Trim(functionName))="eof"  or LCase(Trim(functionName))="x" Then
							Call ReportResult  ("", functionName,"",Param1&"  not found in the test case for the Goto statement" , "Failed" ,"")
							 Exit Do
						End If
				Loop Until UCase(functionName)=UCase(Param1)
				
				If Environment.value("reusescript") = False Then
					Environment.value("tcrow") = tccrntrow
				Else
					Environment.value("reusetcrow") = tccrntrow
				End If		
			Case Else
				If Instr(1, UCase(functionName),"GOTO_") <> 1 Then			
					'TAF 10.1 code modification Start.       stepResultVar is added as a parameter to UtilityActivity function
					actionResult=UtilityActivity(functionName,ExpectedResult,Param1,Param2,Param3,Param4,Param5,Param6,Param7,Param8,Param9,Param10,stepResultVar)        
					'TAF 10.1 code modification End  
					
					On Error Resume Next
					If actionResult(0,0)=-1 Then
						actionResult=BusinessActivity(functionName,ExpectedResult,Param1,Param2,Param3,Param4,Param5,Param6,Param7,Param8,Param9,Param10)
					End If
					If actionResult(0,0)=-1 Then
						Reporter.ReportEvent micFail,"Test Aborted. The Activity name is —>"&functionName,"Input a valid function name in the Activity column. When the UIName is left blank, either a business activity or utility activity should be provided in the'Activity' column of the test case."
						WriteToEvent("Fail" & vbtab & "Test Aborted. The Activity name is —>"&functionName &" Input a valid function name in the Activity column. When the UIName is left blank, either a business activity or utility activity should be provided in the'Activity' column of the test case")
						ExitTest(-1)
					End If				
				End If
			End Select
	
		'********************************************************* IF-ELSE and GOTO End by Srikanth *****************************************************************************

		Case Else
			'TAF 10.1 code modification Start.       stepResultVar is added as a parameter to ApplicationActivity function
			actionResult=ApplicationActivity(appdataSource,UIName,functionName,ExpectedResult,captureScreen,Param1,Param2,Param3,Param4,Param5,Param6,Param7,Param8,Param9,Param10,oParent,stepResultVar)
			'TAF 10.1 code modification End.      
			If actionResult(0,0)=-1 Then
				Reporter.ReportEvent micFail,"Test Aborted. The Activity name is —>"&functionName,"Input a valid function name in the Activity column."
				WriteToEvent("Fail" & vbtab & "Test Aborted. The Activity name is —>"&functionName &" Input a valid function name in the Activity column")
                   ExitTest(-1)
			End If

	End Select
  ElseIf LCase(Trim(Environment("DecideActivtyVBS"))) = "false"  Then ' Conditional statement contains code to handle the Decide activity DLL
    	Environment("Param1")=Param1
	Environment("Param2")=Param2
	Environment("Param3")=Param3
	Environment("Param4")=Param4
	Environment("Param5")=Param5
	Environment("Param6")=Param6
	Environment("Param7")=Param7
	Environment("Param8")=Param8
	Environment("Param9")=Param9
	Environment("Param10")=Param10

	Environment("UIName")=UIName
    Environment("functionName")= functionName
    Environment("ExpectedResult")=ExpectedResult
    Environment("appdataSource")=appdataSource
    Environment("captureScreen")=captureScreen
Set o1=createobject("TAFCore.CoreEngine")
o1.GetCallerInfo()
End if

'	addResult=actionResult(0,0)
'	returnResult=actionResult(0,1)
if lcase(Environment("DecideActivtyVBS")) = "true" then
' code for decide activity vbs file
'	msgbox (Environment("DVBS"))
	'actionResult=Eval(Environment("DVBS"))
	addResult=actionResult(0,0)
'	msgbox addResult
	returnResult=actionResult(0,1)
'	msgbox returnResult
else
' code for decide activity dll
If Environment("check") = "failed" Then
'    msgbox "This function is not registered . Please contact DSTWS for support"
'	aResult(0,0)=-1
'	aResult(0,1)=-1

	addResult= -1
	returnResult= -1
    Call ReportResult("", "Please add the required function or contact DSTWS vegaTest Support ","","This Function "& Environment("functionName") &" is not registered or not avaliable", "Failed" ,"")
	Reporter.ReportEvent micFail , "Function Missing","This Function "& Environment("functionName") &" is not registered"
	Exit Function
	else
	a=Environment("EXE")
'	msgbox "displaying Prevalue of a " &a 
	If instr(1,a,"strParentObject") >0 or instr(1,a,"strChildobject") >0 Then
        EParent=Environment("ParentObject")       
		EChild=Environment("strChildobject")
		PartialReplace=Replace(a,"strParentObject","EParent")
		FullReplaceString=Replace(PartialReplace,"strChildobject","EChild")
'		msgbox FullReplaceString
		Execute("retval="&FullReplaceString)
	else
		Execute("retval="& Environment("EXE"))
	End If  
'	msgbox "ReturnValue Content here "&retval

	addResult="yes"
	returnResult=retval
		
	If LCase(Trim(captureScreen))="yes" Then
			Call ScreenCapture(strParentObject)
	End If       

'		aResult(0,0)=addResult
'        aResult(0,1)=returnResult
'	'	ApplicationActivity=aResult

		CheckVarType = VarType(returnresult)
		If CheckVarType=2 Then
			If returnresult = -1 then
				Environment("FlagTCStep")=false	
				If Environment("RepeatStep") and Environment("Counter")=1 Then		'To handle the recovery with show stopper
					Environment("FlagTCStep")=True	
				End If
			End If
		elseif CheckVarType=11 then
			If returnresult = 0 then
				Environment("FlagTCStep")=false	
				If Environment("RepeatStep") and Environment("Counter")=1 Then		'To handle the recovery with show stopper
					Environment("FlagTCStep")=True	
				End If
	
			End If
		End If			


'	addResult="yes"
'	returnResult=retval
End If
end if
' **********************************  End Of DLL Changes   ****************************************************

	If addResult="yes" Then 
		If Len(Trim(stepResultVar))<>0 Then
			Call BuildDict(returnResult,stepResultVar)
		End If
	End If
End Function



	''
	' This function ActivateRecovery  invokes the exceptional handling mechanism as per the stop event input
	' @author DSTWS
	' @param StopEvent String Specifying the sevearity with which the exceptional should be handled example Minor, Normal or Showstopper.
	' @Modified By DT77734
	' @Modified on: 16 Jul 2009 


Function  ActivateRecovery(StopEvent)
	If  StopEvent ="" Then
		StopEvent="Normal"
	End If
	For Iter = 1 to Recovery.Count 
		Recovery.GetScenarioName Iter, ScenarioFile, ScenarioName 
		If LCase(Trim(ScenarioName))=Lcase(Trim(StopEvent)) Then
			Recovery.SetScenarioStatus Iter,True
		Else
			Recovery.SetScenarioStatus Iter,False
		End If
  Next
End Function


''
	' This function 'ScreenCapture Captures the AUT screen as per the Yes/No capture screen status by taking the string parent object
	' @author DSTWS
	' @param strParentObject String specifying the string name of parent of the test object
	' @Modified By DT77734
	' @Modified on: 16 Jul 2009 

Function ScreenCapture(strParentObject)
	On error resume next  'TAF10 start
	If err.number = 0  Then
    				call ReportResult  ("Capture Screenshot", "Capture Screenshot", "", "", "Screenshot","")
		else
				set objParent=Eval(Environment("ParentObject"))
			call ReportResult  ("Capture Screenshot", "Capture Screenshot", "", "", "Screenshot",objParent)
	End If 'TAF10 end
End Function



 ''
	' This function SetVars sets the value of parameters Param1, Param2, Param3, Param4, Param5 to returnResult
	' @author DSTWS
	' @param Param1 String Specifying the Test Data from the excel sheet
	' @param Param2 String is an optional and non mandatory field
	' @param Param3 String is an optional and non mandatory field
	' @param Param4 String is an optional and non mandatory field
	' @param Param5 String is an optional and non mandatory field
	' @param oParent String is an optional and non mandatory field
	' @Modified By DT77734
	' @Modified on: 16 Jul 2009 
	 
Public Function SetVars(Param1,Param2,Param3,Param4,Param5,Param6,Param7,Param8,Param9,Param10)
	Dim varValues(0,9)
	If Left(LCase(Trim(Param1)),5)="svar_" Then             '10.1
			'TAF 10.1 new code Start
			If  Environment("BlnCodeGeneration")=1 Then
				Param1="oActivityResult.Item("&""""&Param1&""""&")"
			End If
			'TAF 10.1 new code End

		If oActivityResult.Exists(Param1) Then     'Param1 will have the StepResultVariable. StepResultVar will be the key in the dictionary object.
			Param1=oActivityResult.Item(Param1)   'Now the value of Param1 is set to the value of returnResult
        End If
    End If
     Param1 = GetDynamicParameter(Param1)
	' msgbox   Param1

    varValues(0,0)=Param1
    If Left(LCase(Trim(Param2)),5)="svar_" Then
			'TAF 10.1 new code Start
			If  Environment("BlnCodeGeneration")=1 Then
				Param1="oActivityResult.Item("&""""&Param2&""""&")"
			End If
			'TAF 10.1 new code End
		If oActivityResult.Exists(Param2) Then     'Param2 will have the StepResultVariable. StepResultVar will be the key in the dictionary object.
			Param2=oActivityResult.Item(Param2)   'Now the value of Param2 is set to the value of returnResult
		End If
	End If
	Param2 = GetDynamicParameter(Param2)
	'msgbox   Param2
	varValues(0,1)=Param2
    If Left(LCase(Trim(Param3)),5)="svar_" Then
			'TAF 10.1 new code Start
			If  Environment("BlnCodeGeneration")=1 Then
				Param1="oActivityResult.Item("&""""&Param3&""""&")"
			End If
			'TAF 10.1 new code End

		If oActivityResult.Exists(Param3) Then     'Param3 will have the StepResultVariable. StepResultVar will be the key in the dictionary object.
			Param3=oActivityResult.Item(Param3)   'Now the value of Param3 is set to the value of returnResult
		End If
	End If
	Param3 = GetDynamicParameter(Param3)
'	msgbox   Param3
	varValues(0,2)=Param3
    If Left(LCase(Trim(Param4)),5)="svar_" Then
			'TAF 10.1 new code Start
			If  Environment("BlnCodeGeneration")=1 Then
				Param1="oActivityResult.Item("&""""&Param4&""""&")"
			End If
			'TAF 10.1 new code End

		If oActivityResult.Exists(Param4) Then     'Param4 will have the StepResultVariable. StepResultVar will be the key in the dictionary object.
			Param4=oActivityResult.Item(Param4)   'Now the value of Param4 is set to the value of returnResult
		End If
    End If
	Param4= GetDynamicParameter(Param4)
'msgbox   Param4
	varValues(0,3)=Param4
	If Left(LCase(Trim(Param5)),5)="svar_" Then
			'TAF 10.1 new code Start
			If  Environment("BlnCodeGeneration")=1 Then
				Param1="oActivityResult.Item("&""""&Param5&""""&")"
			End If
			'TAF 10.1 new code End

		If oActivityResult.Exists(Param5) Then     'Param5 will have the StepResultVariable. StepResultVar will be the key in the dictionary object.
			Param5=oActivityResult.Item(Param5)   'Now the value of Param5 is set to the value of returnResult
		End If
	End If
	Param5= GetDynamicParameter(Param5)
'	msgbox   Param5
	varValues(0,4)=Param5
		If Left(LCase(Trim(Param6)),5)="svar_" Then
				'TAF 10.1 new code Start
				If  Environment("BlnCodeGeneration")=1 Then
				Param1="oActivityResult.Item("&""""&Param6&""""&")"
			End If
			'TAF 10.1 new code End

		If oActivityResult.Exists(Param6) Then     'Param5 will have the StepResultVariable. StepResultVar will be the key in the dictionary object.
			Param6=oActivityResult.Item(Param6)   'Now the value of Param5 is set to the value of returnResult
		End If
	End If
	Param6= GetDynamicParameter(Param6)
	varValues(0,5)=Param6
	If Left(LCase(Trim(Param7)),5)="svar_" Then
			'TAF 10.1 new code Start
			If  Environment("BlnCodeGeneration")=1 Then
				Param1="oActivityResult.Item("&""""&Param7&""""&")"
			End If
			'TAF 10.1 new code End
		If oActivityResult.Exists(Param7) Then     'Param5 will have the StepResultVariable. StepResultVar will be the key in the dictionary object.
			Param7=oActivityResult.Item(Param7)   'Now the value of Param7 is set to the value of returnResult
		End If
	End If
	Param7= GetDynamicParameter(Param7)
	varValues(0,6)=Param7
	If Left(LCase(Trim(Param8)),5)="svar_" Then
			'TAF 10.1 new code Start
			If  Environment("BlnCodeGeneration")=1 Then
				Param1="oActivityResult.Item("&""""&Param8&""""&")"
			End If
			'TAF 10.1 new code End
		If oActivityResult.Exists(Param8) Then     'Param8 will have the StepResultVariable. StepResultVar will be the key in the dictionary object.
			Param8=oActivityResult.Item(Param8)   'Now the value of Param8 is set to the value of returnResult
		End If
	End If
	Param8= GetDynamicParameter(Param8)
	varValues(0,7)=Param8
	If Left(LCase(Trim(Param9)),5)="svar_" Then
			'TAF 10.1 new code Start
			If  Environment("BlnCodeGeneration")=1 Then
				Param1="oActivityResult.Item("&""""&Param9&""""&")"
			End If
			'TAF 10.1 new code End
		If oActivityResult.Exists(Param9) Then     'Param9 will have the StepResultVariable. StepResultVar will be the key in the dictionary object.
			Param9=oActivityResult.Item(Param9)   'Now the value of Param9 is set to the value of returnResult
		End If
	End If
	Param9= GetDynamicParameter(Param9)
	varValues(0,8)=Param9
		If Left(LCase(Trim(Param10)),5)="svar_" Then
				'TAF 10.1 new code Start
				If  Environment("BlnCodeGeneration")=1 Then
				Param1="oActivityResult.Item("&""""&Param10&""""&")"
			End If
			'TAF 10.1 new code End
		If oActivityResult.Exists(Param10) Then     'Param10 will have the StepResultVariable. StepResultVar will be the key in the dictionary object.
			Param10=oActivityResult.Item(Param10)   'Now the value of Param10 is set to the value of returnResult
		End If
	End If
	Param10= GetDynamicParameter(Param10)
	varValues(0,9)=Param10
	SetVars=varValues
End Function



 ''
	' This function DecideResult  customizes the results as per the expected and actual results and reports to the Results
	' @author DSTWS
	' @param expResult String specifying the expected results entered in the test scripts
	' @param passResult String specifying the pass result
	' @param failResult String Specifying the fail results
	' @param returnResult String Specifying the return results
	' @Modified By DT77734
	' @Modified on: 16 Jul 2009 

Public Function DecideResult(expResult,passResult,failResult,returnResult)
	Dim customReport
	customReport="no"
	If Len(Trim(expResult))=0 And Len(Trim(passResult))<>0 Then
		expResult=passResult
		customReport="allowPass"
	End If
    If Len(Trim(expResult))=0 And Len(Trim(failResult))<>0 Then
		expResult=failResult
		customReport="allowFail"
    End If
    If Len(Trim(expResult))<>0 And Len(Trim(passResult))<>0 And Len(Trim(failResult))<>0 Then
		customReport="allowAll"
    End If
	If returnResult<>-1 Then
		If customReport="allowAll" Or customReport="allowPass" Then
			Call PublishCustomPassResult(expResult,passResult)
        End If
	End If
	If returnResult=-1 Then
		If customReport="allowAll" Or customReport="allowFail" Then
			Call PublishCustomFailResult(expResult,failResult)
        End If
	End If
End Function



''
	' This function PublishCustomFailResult  customs the fail results and reports to results.
	' @author DSTWS
	' @param expResult String specifying the expected results entered in the test scripts
    ' @param failResult String Specifying the fail results
	' @Modified By DT77734
	' @Modified on: 16 Jul 2009 
     
Public Function PublishCustomFailResult(expResult,failResult)
	expResult=CStr(expResult)
	failResult=CStr(failResult)
	Reporter.ReportEvent micFail,""&expResult&"",""&failResult&""
	WriteToEvent("Fail" & vbtab & ""&expResult&""&failResult&"")
End Function



 ''
	' This function PublishCustomPassResult customs the pass results and reports to results.
	' @author DSTWS
	' @param expResult String specifying the expected results entered in the test scripts
    ' @param PassResult String Specifying the fail results
	' @Modified By DT77734
	' @Modified on: 16 Jul 2009 

Public Function PublishCustomPassResult(expResult,passResult)
	expResult=CStr(expResult)
	failResult=CStr(passResult)
	Reporter.ReportEvent micPass,""&expResult&"",""&passResult&""
	WriteToEvent("Pass" & vbtab & ""&expResult&""&failResult&"")
End Function



 ''
	' This function 'BuildDict maintains the results dictionary
	' @author DSTWS
	' @param returnResult String Specifying the return results
	' @param stepResultVar String Specifying the stepResultVar string
	' @Modified By DT77742
	' @Modified on: 28 Jul 2009
	' @Modified By DT77734
	' @Modified on: 16 Jul 2009 
	
set oActivityResult=CreateObject("Scripting.Dictionary")
Public Function BuildDict(returnResult,stepResultVar)
    If Len(Trim(stepResultVar))<>0 Then
		If oActivityResult.Exists(stepResultVar) Then
			oActivityResult.Remove(stepResultVar) 
		End If
		oActivityResult.Add stepResultVar,returnResult
        End If
End Function


 'To terminate an existing excel process before or after a test run
' Created by Sreenu Babu on 20112009

Function KillAnyOpenExcelApplications

strComputer = "."
strProcessToKill = "Excel.exe" 

Set objWMIService = GetObject("winmgmts:" _
	& "{impersonationLevel=impersonate}!\\" _ 
	& strComputer & "\root\cimv2") 

Set colProcess = objWMIService.ExecQuery _
	("Select * from Win32_Process Where Name = '" & strProcessToKill & "'")

count = 0
For Each objProcess in colProcess
	objProcess.Terminate()
	count = count + 1
Next 

End function
