''
' This Libary is to work with win menus.
' @author DSTWS
' @Version 7.0.0

	''
	' This function is to select value from win menu
	' @author DSTWS
	' @param strParentObject String specifying the Parent Object
	' @param strChildObject String specifying the child object
	' @param strValueToSelect String specifying the value to select.
	' @param strUIName String specifying the UI Name of the object
	' @param strExpectedResult String specifying the expected result.
	' @return strValueToSelect if value to Select is not in win menu . Returns -1 if value to select is not in win menu.
	' @Modified By DT77742
	' @Modified on: 31 Jul 2009 
    
Public Function SelectValueInWinMenu(strParentObject,strChildObject,strValueToSelect,strUIName,strExpectedResult)
	Dim strstrUIName
	Dim strExpected
	If  strExpectedResult="" Then
		strExpectedResult="Select  the value:"&strValueToSelect&" in the WinMenu"
	End If
	Set objParentObj=Eval(strParentObject)
   	If objParentObj.exist Then
		set objParent=Eval(strParentObject)
		set  objTestObject=Eval(strParentObject&"."&strChildObject)
'		WriteToScript(strParentObject&"."&strChildobject&".Select"&strValueToSelect)
		If objTestObject.WaitItemProperty(strValueToSelect,"Enabled",true,Cstr(Environment("MaxSyncTime"))) then
'Window("Flight Reservation").WinMenu("Menu").Select "File;Open Order..."
			objTestObject.Select strValueToSelect
			hit=1
		End if
		If hit=1 Then
			'Call ReportResult(strUIName, strExpectedResult, strValueToSelect,ErrorHandler(140,strValueToSelect,""), "Passed" ,objParent)
			Call ReportResult(strUIName, strExpectedResult, strValueToSelect,"Selected menu item", "Passed" ,objParent)
			SelectValueInWinMenu =strValueToSelect
		Else
			Call ReportResult(strUIName, strExpectedResult, strValueToSelect,"Unable to select menu item", "Failed" ,objParent)
			SelectValueInWinMenu=-1
		End If
	Else
		Reporter.ReportEvent micFail,"Parent Object "&strParentObject,"Parent Object "&strParentObject & " not found or avaliable"
		Call ReportResult(strParentObject, strExpectedResult, "","Object not found", "Failed" ,objParent)
		SelectCheckBox = -1
   End If
End Function
