''
' @# author DSTWS
' @# Version 7.0.8
' @# Version 7.0.9 : Added GetdatafromXMl function to work with derived versions, Updated GenerateEnvVaribaleFromQCTestSet



		''
		' This function is to Execute the prerequisite script
		' @author Rajesh Kumar Tatavarthi
		' @return VOID 

		Function ExecutePreReqScript(strPrerequisites)
			arrPrerequisites	= Split(strPrerequisites,";")
			For i= lbound(arrPrerequisites) to ubound(arrPrerequisites)
				Select Case Trim(Lcase(arrPrerequisites(i)))
					Case "MyBasesStateFunction"
							call MyBasesStateFunction()
					Case  "printmsg"
							Print "message-Testing base sate Concept"
        			Case else
							Reporter.ReportEvent micFail,"Running the Base state function:" & Trim(Lcase(arrPrerequisites(i))),"Base state function not defined"
							WriteToEvent("Fail" & vbtab & "Running the Base state function:" & Trim(Lcase(arrPrerequisites(i))) & "--- Base state function not defined")
				End Select
			Next
		End Function


		''
		' This function is to Execute the prerequisite script
		' @author Rajesh Kumar Tatavarthi
		' @return VOID 
		Function GenerateEnvVaribaleFromQCTestSet()
  			On Error Resume Next
'''   			QCUtil.CurrentTestSet.Post
			' Getting the Values from the QC Customized / Default Fileds of the Current Test Set
'''			Environment("Version")  = QCUtil.CurrentTestSet.Field("CY_USER_01")
			Environment("DataRow_KeyWord") =  QCUtil.CurrentTestSetTest.Field("TC_USER_25")
'''			
'''			strVersion = GetdatafromXMl(strVersionControlxml,Environment("Version"))  'Modified by Rajesh 7.0.8a
'''			If strVersion <> "" Then
'''				Environment("DerivedVersion") = Trim(strVersion)
'''			End If
				
		End Function


	Function GetdatafromXMl(strSourcefile,strItem)
		GetdatafromXMl = ""
		Set xmlDoc = CreateObject("Msxml2.DOMDocument") 
		xmlDoc.load(strSourcefile) 
		Set objNodeList = xmlDoc.getElementsByTagName(strItem) 
		If objNodeList.length > 0 then 
			For each x in objNodeList 
				plot=x.Text 
				GetdatafromXMl =  plot 
			Next 
		End If
		Set xmlDoc = Nothing
	End Function
