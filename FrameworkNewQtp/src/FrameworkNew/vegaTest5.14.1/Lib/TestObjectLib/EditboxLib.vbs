''
' This Libary is to work with Edit boxes.
' @author DSTWS
' @Version 7.0.0

	''
	' This function  is to set value in edit box.
	' @author DSTWS
	' @param strParentObject String specifying the Parent Object
	' @param strChildobject String specifying the child object
	' @param strValueToEnter String specifying the value to enter in edit box.
	' @param strUIName String specifying the UI Name of the object
	' @param strExpectedResult String specifying the expected result.
	' @return String  the value entered . Returns -1 if it fails to enter the value in edit box.
	' @Modified By DT77742
	' @Modified on: 31 Jul 2009 

Function SetTextOnEdit(strParentObject,strChildobject,strValueToEnter,strUIName,strExpectedResult)
	If  strExpectedResult="" Then
		strExpectedResult="Enter Text:"&strValueToEnter& " On Editbox:"&strUIName
	End If
	Set objParentObj=Eval(strParentObject)
    If objParentObj.exist Then
		set objParent=Eval(strParentObject)
		set  objTestObject=Eval(strParentObject&"."&strChildobject)
		objTestObject.CheckProperty "Visible",True,Cstr(Environment("MaxSyncTime"))
		On Error Resume Next
		objTestObject.Set strValueToEnter
	    'WriteToScript(strParentObject&"."&strChildobject&".Set " &strValueToEnter)
		If err.Number=0 Then
			Reporter.ReportEvent micPass,"Set text "&strValueToEnter&" in edit field"&strUIName,"Set text "&strValueToEnter&" in edit field"&strUIName
        	'Call ReportResult(strUIName, strExpectedResult, strValueToEnter,ErrorHandler(118,strValueToEnter,strUIName), "Passed" ,objParent)
			Call ReportResult(strUIName, strExpectedResult, strValueToEnter,"Entered the value in the edit box", "Passed" ,objParent)
			SetTextOnEdit =strValueToEnter
		Else
			Reporter.ReportEvent micFail,"Set text "&strValueToEnter&"in edit field"&strUIName,"Failed to set text "&strValueToEnter&" in edit field"&strUIName
			'Call ReportResult(strUIName, strExpectedResult, strValueToEnter,ErrorHandler(119,strValueToEnter,strUIName), "Failed" ,objParent)
			Call ReportResult(strUIName, strExpectedResult, strValueToEnter,"Unable to enter the value in the edit box", "Failed" ,objParent)
			SetTextOnEdit = -1
		End If
	Else
		Reporter.ReportEvent micFail,"Parent Object "&strParentObject,"Parent Object "&strParentObject & " not found or avaliable"
		'Call ReportResult(strUIName, strExpectedResult, "","Failed to Click On button:"&strUIName , "Failed" ,objParent)
		'Call ReportResult(strParentObject, strExpectedResult, "",ErrorHandler(99,strParentObject,""), "Failed" ,objParent)
		Call ReportResult(strParentObject, strExpectedResult, "","Editbox not found", "Failed" ,objParent)
		SetTextOnEdit = -1
	End If
End Function



	''
	' This function sets value for password
	' @author DSTWS
	' @param strParentObject String specifying the Parent Object
	' @param strChildobject String specifying the child object
	' @param strValueToEnter String specifying the value to enter in edit box.
	' @param strUIName String specifying the UI Name of the object
	' @param strExpectedResult String specifying the expected result.
	' @return String  the value entered . Returns -1 if it fails to enter the password in edit box.
	' @Modified By DT77742
	' @Modified on: 31 Jul 2009 

Function SetPassword(strParentObject,strChildobject,strValueToEnter,strUIName,strExpectedResult)
		If  strExpectedResult="" Then
			strExpectedResult="Enter Password:"&strValueToEnter& " On Editbox:"&strUIName
		End If
   		Set objParentObj=Eval(strParentObject)
	   If objParentObj.exist Then
	   On error resume next
		err.clear
		set objParent=Eval(strParentObject)
		set  objTestObject=Eval(strParentObject&"."&strChildobject)
		objTestObject.CheckProperty "Visible",True,Cstr(Environment("MaxSyncTime"))
		objTestObject.SetSecure strValueToEnter
	    WriteToScript(strParentObject&"."&strChildobject&".Set " &strValueToEnter)
		If err.Number=0 Then
			Reporter.ReportEvent micPass,"Set text in password field"&strUIName,"Set text in password field"&strUIName
			'Call ReportResult(strUIName, strExpectedResult, strValueToEnter,ErrorHandler(120,strValueToEnter,strUIName), "Passed" ,objParent)
			Call ReportResult(strUIName, strExpectedResult, strValueToEnter,"Entered the value in the edit box", "Passed" ,objParent)
			SetPassword =strValueToEnter
		Else
			Reporter.ReportEvent micFail,"Set text in password field"&strUIName,"Failed to set password in edit field"&strUIName
			Call ReportResult(strUIName, strExpectedResult, strValueToEnter,"Unable to enter the value in the edit box", "Failed" ,objParent)
			SetPassword = -1
		End If
	Else
		Reporter.ReportEvent micFail,"Parent Object "&strParentObject,"Parent Object "&strParentObject & " not found or avaliable"
		'Call ReportResult(strUIName, strExpectedResult, "","Failed to Click On button:"&strUIName , "Failed" ,objParent)
		'Call ReportResult(strParentObject, strExpectedResult, "",ErrorHandler(99,strParentObject,""), "Failed" ,objParent)
		Call ReportResult(strParentObject, strExpectedResult, "","Object not found", "Failed" ,objParent)
		SetPassword = -1
	End If
End Function



	''
	' This function verifies whether the given length of edit box is same as maximum length of edit box..
	' @author DSTWS
	' @param strParentObject String specifying the Parent Object
	' @param strChildobject String specifying the child object
	' @param intLengthToVerify integer specifying the length of the edit box.
	' @param strUIName String specifying the UI Name of the object
	' @param strExpectedResult String specifying the expected result.
	' @return String  the value entered . Returns -1 if  intLengthToVerify does not match with maximum length of edit box.
	' @Modified By DT77742
	' @Modified on: 31 Jul 2009 

Function VerifyLengthOfEdit(strParentObject,strChildobject,intLengthToVerify,strUIName,strExpectedResult)
	If  strExpectedResult="" Then
		strExpectedResult="Length of the edit box"&strUIName &"Should be"&intLengthToVerify
	End If
	Set objParentObj=Eval(strParentObject)
	If objParentObj.exist Then
		On error resume next
		err.clear
		set objParent=Eval(strParentObject)
		set  objTestObject=Eval(strParentObject&"."&strChildobject)
		objTestObject.CheckProperty"Visible",True,Cstr(Environment("MaxSyncTime"))
		WriteToScript(strParentObject&"."&strChildobject&".GetROProperty("& chr(34) &"max length" &chr(34)& ")" )
		If intLengthToVerify=objTestObject.GetROProperty("max length") then
			hit=1
		End If 
		If hit=1  and err.number=0 Then
			Reporter.ReportEvent micPass,"Length of the edit box :"&strUIName,"Length of editbox "&strUIName& "is same as the expected length"&intLengthToVerify
			Call ReportResult(strUIName, strExpectedResult, intLengthToVerify,ErrorHandler(122,strUIName,intLengthToVerify), "Passed" ,objParent)
			VerifyLengthOfEdit =intLengthToVerify
		Else
			Reporter.ReportEvent micFail,"Length of the edit box :"&strUIName,"Length of editbox "&strUIName& "is not same as the expected"&intLengthToVerify
			Call ReportResult(strUIName, strExpectedResult, intLengthToVerify,ErrorHandler(123,strUIName,intLengthToVerify), "Failed" ,objParent)
			VerifyLengthOfEdit = -1
		End If
	Else
		Reporter.ReportEvent micFail,"Parent Object "&strParentObject,"Parent Object "&strParentObject & " not found or avaliable"
		Call ReportResult(strParentObject, strExpectedResult, "",ErrorHandler(99,strParentObject,""), "Failed" ,objParent)
		VerifyLengthOfEdit = -1
	End If	
End Function



	''
	' This function verifiy whether the TextToVerify is same as the entered text in edit box.
	' @author DSTWS
	' @param strParentObject String specifying the Parent Object
	' @param strChildobject String specifying the child object
	' @param strTextToVerify String specify the text to verify with the text in edit box.
	' @param strUIName String specifying the UI Name of the object
	' @param strExpectedResult String specifying the expected result.
	' @return String  the value entered . Returns -1 if strTextToVerify does not match with text in edit box.
	' @Modified By DT77742
	' @Modified on: 31 Jul 2009 

Function VerifyTextInEdit(strParentObject,strChildobject,strTextToVerify,strUIName,strExpectedResult)
	If  strExpectedResult="" Then
		strExpectedResult="The text"&strTextToVerify &"Should be dispalyed in the edit box "&strUIName
	End If
	Set objParentObj=Eval(strParentObject)
   	If objParentObj.exist Then
	On error resume next
	err.clear
	set objParent=Eval(strParentObject)
	set  objTestObject=Eval(strParentObject&"."&strChildobject)
	objTestObject.CheckProperty"Visible",True,Cstr(Environment("MaxSyncTime"))
    If objTestObject.getROProperty("micclass")="WinEdit" Then
		If Lcase(Trim(strTextToVerify))=LCase(Trim(objTestObject.GetROProperty("text"))) then
			WriteToScript(strParentObject&"."&strChildobject&".GetROProperty("& chr(34) &"text" &chr(34)& ")" )
			hit=1
		End If 
	ElseIf objTestObject.getROProperty("micclass")="WebEdit" Then
		If Lcase(Trim(strTextToVerify))=LCase(Trim(objTestObject.GetROProperty("value"))) then
			WriteToScript(strParentObject&"."&strChildobject&".GetROProperty("& chr(34) &"value" &chr(34)& ")" )
			hit=1
		End If 
	End If
	If hit=1  and err.number=0 Then
		Reporter.ReportEvent micPass,"The Text "&strTextToVerify &" In the edit box :"&strUIName,"The Text "&strTextToVerify &" In the edit box :"&strUIName&"is same as the expected text: "&strTextToVerify
		Call ReportResult(strUIName, strExpectedResult, strTextToVerify,ErrorHandler(124,strTextToVerify,strUIName), "Passed" ,objParent)
		VerifyTextInEdit =strTextToVerify
	Else
		WriteToScript(strParentObject&"."&strChildobject&".GetROProperty("& chr(34) &"value/text" &chr(34)& ")" )
		Reporter.ReportEvent micFail,"The Text "&strTextToVerify &" In the edit box :"&strUIName,"The Text "&strTextToVerify &" In the edit box :"&strUIName&"is  not same as the expected text: "&strTextToVerify
		Call ReportResult(strUIName, strExpectedResult, strTextToVerify,ErrorHandler(125,strTextToVerify,strUIName), "Failed" ,objParent)
		VerifyTextInEdit = -1
	End If
	Else
		Reporter.ReportEvent micFail,"Parent Object "&strParentObject,"Parent Object "&strParentObject & " not found or avaliable"
		Call ReportResult(strParentObject, strExpectedResult, "",ErrorHandler(99,strParentObject,""), "Failed" ,objParent)
		VerifyTextInEdit = -1
	End If
End Function



	''
	' This function verifies the edit box is disabled or not
	' @author DSTWS
	' @param strParentObject String specifying the Parent Object
	' @param strChildobject String specifying the child object
	' @param strUIName String specifying the UI Name of the object
	' @param strExpectedResult String specifying the expected result.
	' @return "Pass" if edit box is disabled . Returns -1 if edit box is not disabled.
	' @Modified By DT77742
	' @Modified on: 31 Jul 2009 

Function VerifyEditboxDisabled(strParentObject,strChildobject,strUIName,strExpectedResult)
	If  strExpectedResult="" Then
		strExpectedResult="The  edit box"&strUIName &"Should be disabled"
	End If
	Set objParentObj=Eval(strParentObject)
   	If objParentObj.exist Then
		On error resume next
		err.clear

		set objParent=Eval(strParentObject)
		set  objTestObject=Eval(strParentObject&"."&strChildobject)
		If objTestObject.exist(20) Then
			WriteToScript(strParentObject&"."&strChildobject&".GetROProperty("& chr(34) &"disabled" &chr(34)& ")" )
			If objTestObject.GetROProperty("disabled") then
				hit=1
			end if
		End If 
		If hit=1 and err.number=0 Then
			Reporter.ReportEvent micPass,"The edit box :"&strUIName,"The edit box :"&strUIName&"is disabled"
			Call ReportResult(strUIName, strExpectedResult, "",ErrorHandler(126,strUIName,""), "Passed" ,objParent)
			VerifyEditboxDisabled ="Pass"
		Else
			Reporter.ReportEvent micFail,"The edit box :"&strUIName,"The edit box :"&strUIName&"is not diabled as expected"
			Call ReportResult(strUIName, strExpectedResult, "",ErrorHandler(127,strUIName,""), "Failed" ,objParent)
			VerifyEditboxDisabled = -1
		End If
		Else		
			Reporter.ReportEvent micFail,"Parent Object "&strParentObject,"Parent Object "&strParentObject & " not found or avaliable"
			Call ReportResult(strParentObject, strExpectedResult, "",ErrorHandler(99,strParentObject,""), "Failed" ,objParent)
			VerifyEditboxDisabled = -1
		End If
End Function



	''
	' This function verifies the absence of edit box
	' @author DSTWS
	' @param strParentObject String specifying the Parent Object
	' @param strChildobject String specifying the child object
	' @param strUIName String specifying the UI Name of the object
	' @param strExpectedResult String specifying the expected result.
	' @return "Pass" if edit box is absent. Returns -1 if edit box is not absent.
	' @Modified By DT77742
	' @Modified on: 31 Jul 2009 

Function VerifyAbsenceOfEditbox(strParentObject,strChildobject,strUIName,strExpectedResult)
	If  strExpectedResult="" Then
		strExpectedResult="The  edit box"&strUIName &"Should not visible"
	End If
	Set objParentObj=Eval(strParentObject)
   	If objParentObj.exist Then
		On error resume next
		err.clear
		set objParent=Eval(strParentObject)
		set  objTestObject=Eval(strParentObject&"."&strChildobject)
		WriteToScript(strParentObject&"."&strChildobject&".Exist(Cstr("&Environment("MaxSyncTime")&")/1000")
		If objTestObject.exist(Cstr(Environment("MaxSyncTime"))/1000) Then
			hit=1
		End If
		If hit=0  And Err.Number=0 Then
			Reporter.ReportEvent micPass,"The Editbox"&strUIName,"The Editbox"&strUIName & " is not Present"
			Call ReportResult(strUIName, strExpectedResult, "",ErrorHandler(128,strUIName,""), "Passed" ,objParent)
			VerifyAbsenceOfEditbox="StepPass"
		Else
			Call ScreenCapture(strParentObject)
			Reporter.ReportEvent micFail,"The Editbox"&strUIName,"This Editbox"&strUIName &"is Present"
			Call ReportResult(strUIName, strExpectedResult, "",ErrorHandler(129,strUIName,"") , "Failed" ,objParent)
			VerifyAbsenceOfEditbox = -1
		End If 
	Else	
		Reporter.ReportEvent micFail,"Parent Object "&strParentObject,"Parent Object "&strParentObject & " not found or avaliable"
		Call ReportResult(strParentObject, strExpectedResult, "",ErrorHandler(99,strParentObject,""), "Failed" ,objParent)
		VerifyAbsenceOfEditbox = -1
	End If
End Function



	''
	' This function  Type in ActiveX
	' @author DSTWS
	' @param strParentObject String specifying the Parent Object
	' @param strChildobject String specifying the child object
	' @param strValueToEnter String specify the strValueToEnter in edit box.
	' @param strUIName String specifying the UI Name of the object
	' @param strExpectedResult String specifying the expected result.
	' @return String  the value entered . Returns -1 if strValueToEnter does not match with text in edit box.
	' @Modified By DT77742
	' @Modified on: 31 Jul 2009 

Function TypeInActiveX(strParentObject,strChildobject,strValueToEnter,strUIName,strExpectedResult)
	If  strExpectedResult="" Then
		strExpectedResult="Enter the value:"&strValueToEnter &"in the  ActiveX:"&strUIName
	End If
	Set objParentObj=Eval(strParentObject)
   	If objParentObj.exist Then
		On error resume next
		err.clear
		Dim hit
		set objParent=Eval(strParentObject)
		set  objTestObject=Eval(strParentObject&"."&strChildobject)
		objTestObject.CheckProperty "Visible",True,Cstr(Environment("MaxSyncTime"))
		objTestObject.Type strValueToEnter
		WriteToScript(strParentObject&"."&strChildobject&".Type "&strValueToEnter)
		If Err.Number=0 Then
			Reporter.ReportEvent micPass,"Type  text "&strValueToEnter&" in ActiveX"&strUIName,"Type  text "&strValueToEnter&" in ActiveX field"&strUIName
			Call ReportResult(strUIName,strValueToEnter, strExpectedResult, ,ErrorHandler(130,strValueToEnter,strUIName), "Passed" ,objParent)
			TypeInActiveX =strValueToEnter
		Else
			Reporter.ReportEvent micFail,"Type  text "&strValueToEnter&"in ActiveX field"&strUIName,"Failed to Type text "&strValueToEnter&"in ActiveX field"&strUIName
			Call ReportResult(strUIName, strExpectedResult,strValueToEnter ,ErrorHandler(131,strValueToEnter,strUIName), "Failed" ,objParent)
			TypeInActiveX = -1
		End If
	Else
		Reporter.ReportEvent micFail,"Parent Object "&strParentObject,"Parent Object "&strParentObject & " not found or avaliable"
		Call ReportResult(strParentObject, strExpectedResult, "",ErrorHandler(99,strParentObject,""), "Failed" ,objParent)
		TypeInActiveX = -1
	End If
End Function


Function GetProperty(strParentObject,strChildobject,PropertyToGet,strUIName,strExpectedResult)
	If  strExpectedResult="" Then
		strExpectedResult="Length of the edit box"&strUIName &"Should be"&PropertyToGet
	End If
	Set objParentObj=Eval(strParentObject)
	If objParentObj.exist Then
		On error resume next
		err.clear
		set objParent=Eval(strParentObject)
		set  objTestObject=Eval(strParentObject&"."&strChildobject)
		objTestObject.CheckProperty"Visible",True,Cstr(Environment("MaxSyncTime"))
		WriteToScript(strParentObject&"."&strChildobject&".GetROProperty("& chr(34) &PropertyToGet &chr(34)& ")" )
' PropertyValue=objTestObject.GetROProperty&"("&chr(34)&PropertyToGet &chr(34)&")"
 PropertyValue=objTestObject.GetROProperty(trim(PropertyToGet))
'WriteToScript(objTestObject.GetROProperty&"("& chr(34) &PropertyToGet &chr(34)&" )")
WriteToScript(objTestObject.GetROProperty(trim(PropertyToGet)))
		If  err.number=0 Then
			Reporter.ReportEvent micPass,"Length of the edit box :"&strUIName,"Length of editbox "&strUIName& "is same as the expected length"&PropertyToGet
			Call ReportResult(strUIName, strExpectedResult, PropertyToGet,ErrorHandler(122,strUIName,PropertyToGet), "Passed" ,objParent)
			GetProperty =PropertyValue
		Else
			Reporter.ReportEvent micFail,"Length of the edit box :"&strUIName,"Length of editbox "&strUIName& "is not same as the expected"&PropertyToGet
			Call ReportResult(strUIName, strExpectedResult, PropertyToGet,ErrorHandler(123,strUIName,PropertyToGet), "Failed" ,objParent)
			GetProperty = -1
		End If
	Else
		Reporter.ReportEvent micFail,"Parent Object "&strParentObject,"Parent Object "&strParentObject & " not found or avaliable"
		Call ReportResult(strParentObject, strExpectedResult, "",ErrorHandler(99,strParentObject,""), "Failed" ,objParent)
		GetProperty = -1
	End If	
End Function


Function getTextOfEditbox(strParentObject,strChildobject,strUIName,strExpectedResult)
	If  strExpectedResult="" Then
		strExpectedResult="Get text from edit box: "&strUIName 
	End If
	Set objParentObj=Eval(strParentObject)
   	If objParentObj.exist Then
		On error resume next
		err.clear
		set objParent=Eval(strParentObject)
		set  objTestObject=Eval(strParentObject&"."&strChildobject)
		WriteToScript(strParentObject&"."&strChildobject&".Exist(Cstr("&Environment("MaxSyncTime")&")/1000")
		If objTestObject.exist(Cstr(Environment("MaxSyncTime"))/1000) Then
			strData=objTestObject.getRoproperty("text")
			getTextOfEditbox=strData
			

			hit=1
		End If
		If hit=0  And Err.Number=0 Then
			Reporter.ReportEvent micFail,"The Editbox"&strUIName,"The Editbox"&strUIName & " is not Present"
			Call ReportResult(strUIName, strExpectedResult, "","Failed to get the data from edit box", "Passed" ,objParent)
			getTextOfEditbox=-1
		Else
			Call ScreenCapture(strParentObject)
			Reporter.ReportEvent micPass,"The Editbox"&strUIName,"This Editbox"&strUIName &"is Present"
			Call ReportResult(strUIName, strExpectedResult, strData,"Get the data from edit box "&strUIName , "Passed" ,objParent)
			'getTextOfEditbox = -1
		End If 
	Else	
		Reporter.ReportEvent micFail,"Parent Object "&strParentObject,"Parent Object "&strParentObject & " not found or avaliable"
		Call ReportResult(strParentObject, strExpectedResult, "",ErrorHandler(99,strParentObject,""), "Failed" ,objParent)
		getTextOfEditbox = -1
	End If
End Function
