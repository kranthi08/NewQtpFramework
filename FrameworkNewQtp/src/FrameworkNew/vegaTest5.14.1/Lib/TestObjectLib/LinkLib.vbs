''
' This Libary is to work with Links.
' @author DSTWS

	''
	' This function  is to click on link.
	' @author DSTWS
	' @param strParentObject String specifying the Parent Object
	' @param strChildobject String specifying the child object
    ' @param strUIName String specifying the UI Name of the object
	' @param strExpectedResult String specifying the expected result.
	' @return "StepPass"  if clicked on link. Returns -1 if it fail to click on link.
	' @Modified By DT77742
	' @Modified on: 31 Jul 2009 

Function ClickOnLink(strParentObject,strChildobject,strUIName,strExpectedResult)
	Dim hit
	hit=0
	If  strExpectedResult="" Then
		strExpectedResult="Click On Link:"&strUIName
	End If
	Set objParentObj=Eval(strParentObject)
   	If objParentObj.exist Then
		set objParent=Eval(strParentObject)
		set  objTestObject=Eval(strParentObject&"."&strChildobject)
		If objTestObject.CheckProperty("Visible",True,Cstr(Environment("MaxSyncTime"))) Then
			objTestObject.Click
			WriteToScript(strParentObject&"."&strChildobject&".Click")
			hit=1
		Else
			hit=0
		End If
		If hit=1 Then
			Reporter.ReportEvent micPass,"Click link"&strUIName,"Clicked on link"&strUIName
			Call ReportResult(strUIName, strExpectedResult,"" ,ErrorHandler(132,strUIName,""), "Passed" ,objParent)
			ClickOnLink="StepPass"
		Else
			Reporter.ReportEvent micFail,"Click link"&strUIName,"Failed to click on link"&strUIName
			Call ReportResult(strUIName, strExpectedResult,"" ,ErrorHandler(133,strUIName,""), "Failed" ,objParent)
			ClickOnLink = -1
		End If
	Else
		Reporter.ReportEvent micFail,"Parent Object "&strParentObject,"Parent Object "&strParentObject & " not found or avaliable"
		Call ReportResult(strParentObject, strExpectedResult, "",ErrorHandler(99,strParentObject,""), "Failed" ,objParent)
		ClickOnLink = -1
	End If
End Function



	''
	' This function  is to verify presence of link.
	' @author DSTWS
	' @param strParentObject String specifying the Parent Object
	' @param strChildobject String specifying the child object
    ' @param strUIName String specifying the UI Name of the object
	' @param strExpectedResult String specifying the expected result.
	' @return "StepPass"  if link is present. Returns -1 if it link ia not present.
	' @Modified By DT77742
	' @Modified on: 31 Jul 2009 

Function VerifyPresenceOfLink(strParentObject,strChildobject,strUIName,strExpectedResult)
	Dim hit
	hit=0
	If  strExpectedResult="" Then
		strExpectedResult="The  Link:"&strUIName & "should be present"
	End If
	Set objParentObj=Eval(strParentObject)
   	If objParentObj.exist Then
		set objParent=Eval(strParentObject)
		set  objTestObject=Eval(strParentObject&"."&strChildobject)
		WriteToScript(strParentObject&"."&strChildobject&".CheckProperty("& chr(34) &"Visible" &chr(34)&" True,Cstr" &Environment("MaxSyncTime") &")" )
		If objTestObject.CheckProperty("Visible",True,Cstr(Environment("MaxSyncTime")))Then   
			hit=1  
		Else
			hit=0
		End If 
		If hit=1 Then
			Reporter.ReportEvent micPass,"Link present"&strUIName,"Link present"&strUIName
			Call ReportResult(strUIName, strExpectedResult,"" ,ErrorHandler(134,strUIName,""), "Passed" ,objParent)
			VerifyPresenceOfLink="StepPass"
		Else
			Reporter.ReportEvent micFail,"Link not present"&strUIName,"Link not present"&strUIName
			Call ReportResult(strUIName, strExpectedResult,"" ,ErrorHandler(135,strUIName,""), "Failed" ,objParent)
			VerifyPresenceOfLink = -1
		End If
	Else
		Reporter.ReportEvent micFail,"Parent Object "&strParentObject,"Parent Object "&strParentObject & " not found or avaliable"
		Call ReportResult(strParentObject, strExpectedResult, "",ErrorHandler(99,strParentObject,""), "Failed" ,objParent)
		VerifyPresenceOfLink = -1
	End If
End Function



	''
	' This function verify the absence of  link
	' @author DSTWS
	' @param strParentObject String specifying the Parent Object
	' @param strChildobject String specifying the child object
    ' @param strUIName String specifying the UI Name of the object
	' @param strExpectedResult String specifying the expected result.
	' @return "StepPass"  if link is absent. Returns -1 if link is not absent.
	' @Modified By DT77742
	' @Modified on: 31 Jul 2009 

Function VerifyAbsenceOfLink(strParentObject,strChildobject,strUIName,strExpectedResult)
	Dim hit
	hit=0
	If  strExpectedResult="" Then
		strExpectedResult="The  Link:"&strUIName & "should not present"
	End If
	Set objParentObj=Eval(strParentObject)
   	If objParentObj.exist Then
		set objParent=Eval(strParentObject)
		set  objTestObject=Eval(strParentObject&"."&strChildobject)
		WriteToScript(strParentObject&"."&strChildobject&".CheckProperty("& chr(34) &"Visible" &chr(34)&" True,Cstr" &Environment("MaxSyncTime") &")" )
		If objTestObject.CheckProperty("Visible",True,Cstr(Environment("MaxSyncTime")))Then   
			hit=0  
		Else
			hit=1
		End If 
		If hit=1 Then
			Call ScreenCapture(objParent)
			Reporter.ReportEvent micFail,"Link present"&strUIName,"Link present"&strUIName
			Call ReportResult(strUIName, strExpectedResult,"" ,ErrorHandler(135,strUIName,""), "Passed" ,objParent)
			VerifyAbsenceOfLink="StepPass"
		Else
			Reporter.ReportEvent micPass,"Link not present"&strUIName,"Link not present"&strUIName
			Call ReportResult(strUIName, strExpectedResult,"" ,ErrorHandler(134,strUIName,""), "Failed" ,objParent)
			VerifyAbsenceOfLink = -1
		End If
	Else
		Reporter.ReportEvent micFail,"Parent Object "&strParentObject,"Parent Object "&strParentObject & " not found or avaliable"
		Call ReportResult(strParentObject, strExpectedResult, "",ErrorHandler(99,strParentObject,""), "Failed" ,objParent)
		VerifyAbsenceOfLink = -1
	End If
End Function
