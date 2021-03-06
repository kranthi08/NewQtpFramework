''
' This Libary is to work with Radio buttons.
' @author DSTWS
' @Version 7.0.0

	''
	' This function  is to select a win radio button or a web radio button.
	' @author DSTWS
	' @param strParentObject String specifying the Parent Object
	' @param strChildobject String specifying the child object
	' @param strValueToEnter String specifying the value to select from the radio button.
    ' @param strUIName String specifying the UI Name of the object
	' @param strExpectedResult String specifying the expected result.
	' @return "StepPass"  if a radio button is selected. Returns -1 if it fail to select a radio button.
	' @Modified By DT77742
	' @Modified on: 31 Jul 2009 

Function SelectRadioButton(strParentObject,strChildobject ,strValueToEnter,strUIName , strExpectedResult)
	Dim hit
	hit=0 
	If  strExpectedResult="" Then
		strExpectedResult="Select item:"&strValueToEnter &"in the radio button"&strUIName
	End If
	Set objParentObj=Eval(strParentObject)
   	If objParentObj.exist Then
	set objParent=Eval(strParentObject)
	set  objTestObject=Eval(strParentObject&"."&strChildobject)
	If objTestObject.CheckProperty("Enabled",True,Cstr(Environment("MaxSyncTime"))) Then
        If objTestObject.GetROproperty("micclass")="WinRadioButton" Then
            objTestObject.Set 
			WriteToScript(strParentObject&"."&strChildobject&".Set")
			hit=1
		ElseIf objTestObject.GetROproperty("micclass")="WebRadioGroup" Then
			objTestObject.Select strValueToEnter
			WriteToScript(strParentObject&"."&strChildobject&".Select"&strValueToEnter)
			hit=1
		end if
	End If
	If hit=1  Then
		Reporter.ReportEvent micPass,"Selected the  RadioButton "&strUIName,"Selected the  RadioButton "&strUIName 
		Call ReportResult(strUIName, strExpectedResult,strValueToEnter ,ErrorHandler(136,strValueToEnter,strUIName), "Passed" ,objParent)
		SelectRadioButton="StepPass"
	Else
		WriteToScript(strParentObject&"."&strChildobject&".Set/Select")
		Reporter.ReportEvent micFail,"Failed to select the  RadioButton "&strUIName,"Failed to select the  RadioButton "&strUIName 
		Call ReportResult(strUIName, strExpectedResult,strValueToEnter ,ErrorHandler(137,strValueToEnter,strUIName), "Failed" ,objParent)
		SelectRadioButton=-1
	End If
	Else
		Reporter.ReportEvent micFail,"Parent Object "&strParentObject,"Parent Object "&strParentObject & " not found or avaliable"
		Call ReportResult(strParentObject, strExpectedResult, "",ErrorHandler(99,strParentObject,""), "Failed" ,objParent)
		SelectRadioButton = -1
   End If
End Function