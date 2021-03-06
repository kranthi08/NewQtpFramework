''
' This Libary is to work with check boxes.
' @author DSTWS
' @Version 7.0.0

	''
	' This function is to select value from check box.
	' @author DSTWS
	' @param strParentObject String specifying the Parent Object
	' @param strChildobject String specifying the child object
	' @param strValueToSet String specifying the value to Set.
	' @param strUIName String specifying the UI Name of the object
	' @param strExpectedResult String specifying the expected result.
	' @return "Step Pass" if it select value from check box .  Returns -1 if it fails to select value from check box.
	' @Modified By DT77742
	' @Modified on: 31 Jul 2009
	' @Added ParentObject Error Handling DT77215 

Function SelectCheckBox(strParentObject,strChildobject ,strValueToSet,strUIName,strExpectedResult)
	If  strExpectedResult="" Then
		strExpectedResult="Select:"&strValueToSet&",on checkbox"&strUIName
	End If
   Set objParentObj=Eval(strParentObject)
   If objParentObj.exist Then
	On error resume next
	Err.Clear
	set objParent=Eval(strParentObject)
	set  objTestObject=Eval(strParentObject&"."&strChildobject)
	objTestObject.CheckProperty"Enabled",True,Cstr(Environment("MaxSyncTime"))
	WriteToScript(strParentObject&"."&strChildobject&".Select " &ValueToEnter )
	objTestObject.Set strValueToSet
	If Err.Number=0  Then
		Reporter.ReportEvent micPass,"Set CheckBox "&strUIName,"Checkbox  "&strUIName&"  set"
		Call ReportResult(strUIName, strExpectedResult,strValueToSet ,ErrorHandler(111,strValueToSet,strUIName), "Passed" ,objParent)
		SelectCheckBox ="StepPass"
	Else
		Reporter.ReportEvent micFail,"Set CheckBox "&strUIName,"Checkbox  "&strUIName&" was not set"
		'Call ReportResult(strUIName, strExpectedResult,strValueToSet ,"Failed to Select:"&strValueToSet&",on checkbox"&strUIName , "Passed" ,objParent)
		Call ReportResult(strUIName, strExpectedResult,strValueToSet ,ErrorHandler(112,strValueToSet,strUIName), "Failed" ,objParent)
		SelectCheckBox = -1
	End If 
	Else
		Reporter.ReportEvent micFail,"Parent Object "&strParentObject,"Parent Object "&strParentObject & " not found or avaliable"
		'Call ReportResult(strUIName, strExpectedResult, "","Failed to Click On button:"&strUIName , "Failed" ,objParent)
		Call ReportResult(strParentObject, strExpectedResult, "",ErrorHandler(99,strParentObject,""), "Failed" ,objParent)
		SelectCheckBox = -1
   End If
End Function
