''
' This Libary is to work with Buttons.
' @author DSTWS
' @Version 7.0.0

	''
	' This function is to click on button.
	' @author DSTWS
	' @param strParentObject String specifying the Parent Object
	' @param strChildobject String specifying the child object
	' @param strUIName String specifying the UI Name of the object
	' @param strExpectedResult String specifying the expected result.
	' @return "Step Pass" if it clicks on button .  Returns -1 if it fails to click on button
	' @Modified By DT77742
	' @Modified on: 31 Jul 2009 

Function ClickOnButton(strParentObject,strChildobject,strUIName,strExpectedResult)

	If  strExpectedResult="" Then
			strExpectedResult="Click On button:"&strUIName
	End If
	Set objParentObj=Eval(strParentObject)
   If objParentObj.exist Then
	   On error resume next
		err.Clear	
		set objParent=Eval(strParentObject)
		set  objTestObject=Eval(strParentObject&"."&strChildobject)
		objTestObject.CheckProperty "Visible",True,Cstr(Environment("MaxSyncTime"))
		objTestObject.Click
		WriteToScript(strParentObject&"."&strChildobject&".Click ")
		If Err.Number=0 Then
			Reporter.ReportEvent micPass,"Click on button "&strUIName,"Clicked on button "&strUIName
			Call ReportResult(strUIName, strExpectedResult, "","Clicked On button:"&strUIName , "Passed" ,objParent)
			'Call ReportResult(strUIName, strExpectedResult, "",ErrorHandler(101,strUIName,"") , "Passed" ,objParent)
			ClickOnButton="StepPass"
		Else
			Reporter.ReportEvent micFail,"Click on button "&strUIName,"Failed to click on button "&strUIName
			Call ReportResult(strUIName, strExpectedResult, "","Failed to Click On button:"&strUIName , "Failed" ,objParent)
			'Call ReportResult(strUIName, strExpectedResult, "",ErrorHandler(102,strUIName,""), "Failed" ,objParent)
			ClickOnButton = -1
		End If 
		Else
			Reporter.ReportEvent micFail,"Parent Object "&strParentObject,"Parent Object "&strParentObject & " not found or avaliable"
			Call ReportResult(strUIName, strExpectedResult, "","Failed to Click On button:"&strUIName , "Failed" ,objParent)
			'Call ReportResult(strParentObject, strExpectedResult, "",ErrorHandler(99,strParentObject,""), "Failed" ,objParent)
			ClickOnButton = -1
   End If

End Function



	''
	' This function is to verify absence of  button
	' @author DSTWS
	' @param strParentObject String specifying the Parent Object
	' @param strChildobject String specifying the child object
	' @param strUIName String specifying the UI Name of the object
	' @param strExpectedResult String specifying the expected result.
	' @return "Step Pass" if button is absent .  Returns -1 if button is present.
	' @Modified By DT77742
	' @Modified on: 31 Jul 2009 

Function VerifyAbsenceOfButton(strParentObject,strChildobject,strUIName,strExpectedResult)
	If  strExpectedResult="" Then
		strExpectedResult="The  button:"&strUIName &"should not visible"
	End If
   	Set objParentObj=Eval(strParentObject)
   If objParentObj.exist Then
   On error resume next
	Err.Clear
	set objParent=Eval(strParentObject)
	set  objTestObject=Eval(strParentObject&"."&strChildobject)
	WriteToScript(strParentObject&"."&strChildobject&".exist(0)")
	If objTestObject.exist(0) Then   
		hit=0  
	Else
		hit=1
	End If 
	If hit=1  And Err.Number=0 Then
		Reporter.ReportEvent micPass,"The button "&strUIName,"This button "&strUIName & "not exist"
		'Call ReportResult(strUIName, strExpectedResult, "","The  button:"&strUIName &"is  not visible" , "Passed" ,objParent)
		Call ReportResult(strUIName, strExpectedResult, "",ErrorHandler(105,strUIName,"") , "Passed" ,objParent)
		VerifyAbsenceOfButton="StepPass"
	Else
		Reporter.ReportEvent micFail,"The button "&strUIName,"This button "&strUIName & " exist"
		'Call ReportResult(strUIName, strExpectedResult, "","The  button:"&strUIName &"is   visible" , "Failed" ,objParent)
		Call ReportResult(strUIName, strExpectedResult, "",ErrorHandler(106,strUIName,"") , "Failed" ,objParent)
		VerifyAbsenceOfButton = -1
	End If 
	Else
		Reporter.ReportEvent micFail,"Parent Object "&strParentObject,"Parent Object "&strParentObject & " not found or avaliable"
        Call ReportResult(strParentObject, strExpectedResult, "",ErrorHandler(99,strParentObject,""), "Failed" ,objParent)
		VerifyAbsenceOfButton = -1
   End If
End Function



	''
	 'This function is to verify  presence of  button
	' @author DSTWS
	' @param strParentObject String specifying the Parent Object
	' @param strChildobject String specifying the child object
	' @param strUIName String specifying the UI Name of the object
	' @param strExpectedResult String specifying the expected result.
	' @return "Step Pass" if button is present .  Returns -1 if button is not present.
	' @Modified By DT77742
	' @Modified on: 31 Jul 2009 

Function VerifyPresenceOfButton(strParentObject,strChildobject,strUIName,strExpectedResult)
	If  strExpectedResult="" Then
		strExpectedResult="The  button:"&strUIName &"should be  visible"
	End If
   	Set objParentObj=Eval(strParentObject)
   If objParentObj.exist Then
     On error resume next
	Err.clear
	set objParent=Eval(strParentObject)
	set  objTestObject=Eval(strParentObject&"."&strChildobject)
	objTestObject.CheckProperty "Visible",True,Cstr(Environment("MaxSyncTime"))
	WriteToScript(strParentObject&"."&strChildobject&".CheckProperty Visible,True,Cstr("&Environment("MaxSyncTime")&")")
	msgbox "Cstr"&Environment("MaxSyncTime")
	If Err.Number=0  Then
		Reporter.ReportEvent micPass,"The button "&strUIName,"This button "&strUIName & " exist"
		'Call ReportResult(strUIName, strExpectedResult, "","The  button:"&strUIName &"is   visible" , "Passed" ,objParent)
		Call ReportResult(strUIName, strExpectedResult, "",ErrorHandler(106,strUIName,"") , "Passed" ,objParent)
		VerifyPresenceOfButton="StepPass"
	Else
		Reporter.ReportEvent micFail,"The button "&strUIName,"This button "&strUIName & " Not exist"
		'Call ReportResult(strUIName, strExpectedResult, "","The  button:"&strUIName &"is  not  visible" , "Failed" ,objParent)
		Call ReportResult(strUIName, strExpectedResult, "",ErrorHandler(105,strUIName,""), "Failed" ,objParent)
		VerifyPresenceOfButton = -1
	End If 
	Else
		Reporter.ReportEvent micFail,"Parent Object "&strParentObject,"Parent Object "&strParentObject & " not found or avaliable"
        Call ReportResult(strParentObject, strExpectedResult, "",ErrorHandler(99,strParentObject,""), "Failed" ,objParent)
		VerifyPresenceOfButton = -1
   End If
End Function



	''
	' This function is to verify  the button is disabled or not
	' @author DSTWS
	' @param strParentObject String specifying the Parent Object
	' @param strChildobject String specifying the child object
	' @param strUIName String specifying the UI Name of the object
	' @param strExpectedResult String specifying the expected result.
	' @return "Step Pass" if button is disabled .  Returns -1 if button is not disabled.
	' @Modified By DT77742
	' @Modified on: 31 Jul 2009 

Function VerifyButtonDisabled(strParentObject,strChildobject,strUIName,strExpectedResult)
	If  strExpectedResult="" Then
		strExpectedResult="The  button:"&strUIName &"should be  disable"
	End If
   Set objParentObj=Eval(strParentObject)
   If objParentObj.exist Then
   On error resume next
	Err.Clear
	set objParent=Eval(strParentObject)
	set  objTestObject=Eval(strParentObject&"."&strChildobject)
	WriteToScript(strParentObject&"."&strChildobject&".GetROProperty("&disabled &")")
	If objTestObject.exist(Cstr(Environment("MaxSyncTime"))/1000) Then
		If objTestObject.GetROProperty("disabled") then
			hit=1
		end if
	Else
		hit=0
	End If 
	If hit=0  And Err.Number=0 Then
		Reporter.ReportEvent micPass,"The button "&strUIName,"The button "&strUIName & " is disabled"
		'Call ReportResult(strUIName, strExpectedResult, "","The  button:"&strUIName &"is  disabled" , "Passed" ,objParent)
		Call ReportResult(strUIName, strExpectedResult, "",ErrorHandler(107,strUIName,"") , "Passed" ,objParent)
		VerifyButtonDisabled="StepPass"
	Else
		Reporter.ReportEvent micFail,"The button "&strUIName,"This button "&strUIName & " is not disabled"
		'Call ReportResult(strUIName, strExpectedResult, "","The  button:"&strUIName &"is  not  disabled" , "Failed" ,objParent)
		Call ReportResult(strUIName, strExpectedResult, "",ErrorHandler(108,strUIName,"") , "Failed" ,objParent)
		VerifyButtonDisabled = -1
	End If 
	Else
		Reporter.ReportEvent micFail,"Parent Object "&strParentObject,"Parent Object "&strParentObject & " not found or avaliable"
        Call ReportResult(strParentObject, strExpectedResult, "",ErrorHandler(99,strParentObject,""), "Failed" ,objParent)
		VerifyButtonDisabled = -1
   End If
End Function
