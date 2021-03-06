''
' This Libary is to work with Web elements.
' @author DSTWS
' @Version 7.0.0

	''
	' This function is to compare two strings.
	' @author DSTWS
	' @param strParentObject String specifying the Parent Object
	' @param strChildobject String specifying the child object
	' @param strMessageToFind String specifying the MessageTo Find on browser
	' @param strUIName String specifying the UI Name of the object
	' @param strExpectedResult String specifying the expected result.
	' @return Steppass if message to Find is present in browser, Returns -1 if not present.
	' @Modified By DT77742
	' @Modified on: 16 Jul 2009 
    
Function FindMessage(strParentObject,strChildobject,strMessageToFind,strUIName,strExpectedResult)
	If  strExpectedResult="" Then
		strExpectedResult="Text:"&strUIName & " is displayed"
	End If
	set objParent=Eval(strParentObject)    
	set  objTestObject=Eval(strParentObject&"."&strChildobject)
	If strMessageToFind="" Then
		If objTestObject.Exist(Cstr(Environment("MaxSyncTime"))/1000) Then
			hit=1
		Else
			hit=0
		End If
	Else
		Set oWebElement=Description.Create()
		oWebElement("micclass").Value="WebElement"
		Set WebElem = objParent.ChildObjects(oWebElement)
		NumberOfElements = WebElem.Count()
		For i = 1 To NumberOfElements - 1
			On Error Resume Next
			ElementInnerText=WebElem(i).GetROProperty("innertext")
			ElementVisible=WebElem(i).GetROproperty("visible")
			ElementInnerText=WebElem(i).GetROProperty("innertext")
			If Trim(ElementVisible)="True" Then
				If InStr(Trim(ElementInnerText),Trim(strMessageToFind))<>0 Then
					TextFound=Mid(Trim(ElementInnerText),InStr(Trim(ElementInnerText),Trim(strMessageToFind)),Len(strMessageToFind))
					If strMessageToFind<>"" Then
						If Trim(TextFound)=Trim(strMessageToFind) Then
							FoundElement=FoundElement+1         
							Exit For
						End If
					End If
				End If
			End If
		Next
		If FoundElement>0 Then
			FindMessage=ElementInnerText
			hit=1
		Else
			hit=0
			FindMessage = -1
		End If
	End If
	If hit=1  And Err.Number=0 Then
		Call ReportResult(strUIName, strExpectedResult, "","Text:"&strUIName& " is displayed" , "Passed" ,objParent)
		FindMessage="StepPass"
	Else
		Call ReportResult(strUIName, strExpectedResult, "","Text:"&strUIName& " is not displayed" , "Failed" ,objParent)
		FindMessage = -1
	End If 
End Function
