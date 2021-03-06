''
' This Libary is to work with Images
' @author DSTWS
' @Version 7.0.0

	''
	' This function  is to click on image.
	' @author DSTWS
	' @param strParentObject String specifying the Parent Object
	' @param strChildobject String specifying the child object
    ' @param strUIName String specifying the UI Name of the object
	' @param strExpectedResult String specifying the expected result.
	' @return String  the value entered . Returns -1 if it fails to enter the value in edit box.
	' @Modified By DT77742
	' @Modified on: 31 Jul 2009 

Function ClickOnImage(strParentObject,strChildobject,strUIName,strExpectedResult)
	If  strExpectedResult="" Then
		strExpectedResult="Click On Image"&strUIName
	End If
	Set objParentObj=Eval(strParentObject)
   	If objParentObj.exist Then
		hit=0
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
			Reporter.ReportEvent micPass,"Click on Image"&strUIName,"Click on Image"&strUIName
			Call ReportResult(strUIName, strExpectedResult,"" ,ErrorHandler(109,strUIName,""), "Passed" ,objParent)
			ClickOnImage="StepPass"
		Else
			Reporter.ReportEvent micFail,"Click on Image"&strUIName,"Failed to click on Image"&strUIName
			Call ReportResult(strUIName, strExpectedResult,"" ,ErrorHandler(110,strUIName,""), "Failed" ,objParent)
			ClickOnImage = -1
		End If
	Else
		Reporter.ReportEvent micFail,"Parent Object "&strParentObject,"Parent Object "&strParentObject & " not found or avaliable"
		Call ReportResult(strParentObject, strExpectedResult, "",ErrorHandler(99,strParentObject,""), "Failed" ,objParent)
		ClickOnImage = -1
	End If
End Function
