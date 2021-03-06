''
' This Libary is to work with Combo boxes.
' @author DSTWS
' @Version 7.0.0
	''
	' This function verify whether the value is not present in combo.
	' @author DSTWS
	' @param strParentObject String specifying the Parent Object
	' @param strChildobject String specifying the child object
	' @param strValueToVerify String specifying the value to verify.
	' @param strUIName String specifying the UI Name of the object
	' @param strExpectedResult String specifying the expected result.
	' @return "StepPass" if value to verify is not in combo box . Returns -1 if value to verify is in combo box.
	' @Modified By DT77742
	' @Modified on: 31 Jul 2009 

Function VerifyItemNotInCombo(strParentObject,strChildobject,strValueToVerify,strUIName,strExpectedResult)
	Dim hit
	hit=0
	If  strExpectedResult="" Then
		strExpectedResult="The item:"&strValueToVerify&"Should not be present in the combobox"&strUIName
	End If
	Set objParentObj=Eval(strParentObject)
   	If objParentObj.exist Then
	set objParent=Eval(strParentObject)
	set  objTestObject=Eval(strParentObject&"."&strChildobject)
	If objTestObject.CheckProperty("Visible",True,Cstr(Environment("MaxSyncTime")))Then   
		Cnt=objTestObject.GetROProperty("items count")
		WriteToScript(strParentObject&"."&strChildobject&".GetROProperty("& chr(34) &"items count" &chr(34)& ")" )
		For i=1 to Cnt
			if Trim(Lcase(strValueToVerify))=Trim(Lcase(objTestObject.GetItem(i))) then
				hit=1
			End If
		Next  
	End If  
	If hit=0 Then
		Reporter.ReportEvent micPass,"Item "&strValueToVerify&" not in Combo "&strUIName,"Item "&strValueToVerify&" not in Combo "&strUIName
		Call ReportResult(strUIName, strExpectedResult,strValueToVerify ,ErrorHandler(113,strValueToVerify,strUIName), "Passed" ,objParent)
		VerifyItemNotInWebCombo =ValueToSelect
	Else
		Reporter.ReportEvent micFail,"Item "&strValueToVerify&" present in Combo "&strUIName,"Item "&strValueToVerify&" present in Combo "&strUIName
		'Call ReportResult(strUIName, strExpectedResult,strValueToVerify ,"The item:"&strValueToVerify&"is   present in the combobox"&strUIName , "Failed" ,objParent)
		Call ReportResult(strUIName, strExpectedResult,strValueToVerify ,ErrorHandler(114,strValueToVerify,strUIName), "Failed" ,objParent)
		VerifyItemNotInWebCombo =-1
	End If
	Else
		Reporter.ReportEvent micFail,"Parent Object "&strParentObject,"Parent Object "&strParentObject & " not found or avaliable"
		Call ReportResult(strParentObject, strExpectedResult, "",ErrorHandler(99,strParentObject,""), "Failed" ,objParent)
		VerifyItemNotInWebCombo =-1
	End If
End Function



	''
	' This function gets all items form WebCombo
	' @author DSTWS
	' @param strParentObject String specifying the Parent Object
	' @param strChildobject String specifying the child object
    ' @param strUIName String specifying the UI Name of the object
	' @param strExpectedResult String specifying the expected result.
	' @return the number of items in combo box . Returns -1 if combo box is diabled.
	' @Modified By DT77742
	' @Modified on: 31 Jul 2009 

Function GetAllItemsInCombo(strParentObject,strChildobject,strUIName,strExpectedResult)
	set  objTestObject=Eval(strParentObject&"."&strChildobject)
	If objTestObject.CheckProperty("Visible",True,Cstr(Environment("MaxSyncTime"))) Then
		WriteToScript(strParentObject&"."&strChildobject&".GetROProperty("& chr(34) &"all items" &chr(34)& ")" )
		dval=objTestObject.GetROProperty("all items")      
		Reporter.ReportEvent micPass,"All items in the Combo box are:"&dval,"All items in the Combo box are:"
		GetAllItemsInCombo =dval
	Else
		Reporter.ReportEvent micFail," Combo box :"&strUIName,"is diabled and unable to get the items:"
	End If  
End Function



	''
	' This function is to select value in combo box.
	' @author DSTWS
	' @param strParentObject String specifying the Parent Object
	' @param strChildobject String specifying the child object
	' @param strValueToSelect String specifying the value to select from combo.
    ' @param strUIName String specifying the UI Name of the object
	' @param strExpectedResult String specifying the expected result.
	' @return strValueToSelect if it select the value in combo  . Returns -1 if combo box is diabled.
	' @Modified By DT77742
	' @Modified on: 31 Jul 2009 

Function SelectValueInCombo(strParentObject,strChildobject,strValueToSelect,strUIName,strExpectedResult)
	hit=0
	If  strExpectedResult="" Then
		strExpectedResult="The item:"&strValueToSelect &"Should  be present in the combobox"&strUIName
	End If
	Set objParentObj=Eval(strParentObject)
   	If objParentObj.exist Then
	set objParent=Eval(strParentObject)
	set  objTestObject=Eval(strParentObject&"."&strChildobject)
	If objTestObject.CheckProperty("Visible",True,Cstr(Environment("MaxSyncTime"))) Then  
		objTestObject.select strValueToSelect    
		WriteToScript(strParentObject&"."&strChildobject&".select" &strValueToSelect)  
		hit=1     
	Else
		hit=0
	End If  
	If hit=1 Then
		Reporter.ReportEvent micPass,"Select value "&strValueToSelect&" in Combo "&strUIName,"Selected value "&strValueToSelect&" in Combo "&strUIName
		Call ReportResult(strUIName, strExpectedResult,strValueToSelect ,"The item:"& strValueToSelect &"is   present in the combobox"&strUIName , "Passed" ,objParent)
		'Call ReportResult(strUIName, strExpectedResult,strValueToSelect ,ErrorHandler(116,strValueToSelect,strUIName), "Passed" ,objParent)
		SelectValueInWebCombo =strValueToSelect
	Else
		Call ScreenCapture(strParentObject)
		Reporter.ReportEvent micFail,"Select value "&strValueToSelect&" in Combo "&strUIName,"Failed to select value "&strValueToSelect&" in Combo field "&strUIName
		Call ReportResult(strUIName, strExpectedResult,strValueToSelect ,"The item:"& strValueToSelect &"is  not present in the combobox"&strUIName , "Failed" ,objParent)
		'Call ReportResult(strUIName, strExpectedResult,strValueToSelect ,ErrorHandler(117,strValueToSelect,strUIName), "Failed" ,objParent)
		SelectValueInWebCombo = -1
	End If
	Else
		Reporter.ReportEvent micFail,"Parent Object "&strParentObject,"Parent Object "&strParentObject & " not found or avaliable"
		Call ReportResultNon(strParentObject, strExpectedResult, "","Parent object not found", "Failed" ,objParent)
		SelectValueInCombo =-1
	End If
End Function
