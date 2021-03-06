
Function ExecuteTestCase

   Environment("Execute") = True

   Environment("TestDataPathToReport") = Environment("strTestDataFilePath")
	DataTable("TestCaseSheetName","Global")=DataTable.GetSheet("TestPlan").GetParameter("TestCaseName")      
	DataTable("TestCaseFilePath","Global")=DataTable.GetSheet("TestPlan").GetParameter("TestCaseFilePath")	
 Call runFlash

If blnNoTDExists then
		 If  Environment("RunFromQC") Then

		Reporter.ReportEvent micWarning,"Verifying the Test Data for the Test Case :  " & Environment("TestCaseName1") , "Test Data As Specified in the test plan  : " & Environment("QCTestDataFilePath") & "Does Not Exists"
		WriteToEvent("Warning" & vbtab & "Test Data As Specified in the test plan  : " & Environment("QCTestDataFilePath") & "Does Not Exists")

		else
		Reporter.ReportEvent micWarning,"Verifying the Test Data for the Test Case :  " & Environment("TestCaseName1") , "Test Data As Specified in the test plan  : " & Environment("strTestDataFilePath") & "Does Not Exists"
		WriteToEvent("Warning" & vbtab & "Test Data As Specified in the test plan  : " & Environment("strTestDataFilePath") & "Does Not Exists")
         End If 
	End if
	DataTable("TestDataDBPath","Global")= Environment("strTestDataFilePath")
	Environment("TestDataPath")=Environment("strTestDataFilePath")
	strSheetName=DataTable.GetSheet("TestPlan").GetParameter("TestDataSheetName")
	Environment("TestDataSheetName")=strSheetName
	
	on error resume next
	arr  = Split(DataTable.GetSheet("TestPlan").GetParameter("DataRow_Keyword"),";")
	'' Scenario keys
		If instr( 1,lcase(DataTable.GetSheet("TestPlan").GetParameter("DataRow_Keyword")),"scenariokeys") > 0 Then
			arrTCKeys = Split( lcase(DataTable.GetSheet("TestPlan").GetParameter("DataRow_Keyword")),";")
			If  ubound(arrTCKeys) >0 Then
					strkey = arrTCKeys(1)
					If (lcase(Trim(strKey)) ="fo"  and Environment("FirstOnly"))  or  ( lcase(Trim(strKey)) ="#fo#"  and Environment("FirstOnly") )  Then           ' FO : first only 
						 strKeyWord = Environment("ScenarioDataKeys")
					elseif (lcase(Trim(strKey)) ="lo"  and  Environment("LastOnly"))  or  ( lcase(Trim(strKey)) ="#lo#"  and  Environment("LastOnly")  )  Then		' LO : last only 
						   strKeyWord = Environment("ScenarioDataKeys")
				   elseif lcase(trim(strKey)) = "iteration#"  & Environment("SceIte") then
							 strKeyWord = Environment("ScenarioDataKeys")
					else
		                    Environment("Execute") = False
		            	     Exit Function
					End If
			else
					 strKeyWord = Environment("ScenarioDataKeys")
			End If

		elseif  (Lcase(Trim(arr(ubound(arr)) )) =  "fo"  and Environment("FirstOnly"))  or (Lcase(Trim(arr(ubound(arr)) )) =  "#fo#"  and Environment("FirstOnly") )  then
						If ubound(arr) > 0  Then
							strKeyWord = Left(DataTable.GetSheet("TestPlan").GetParameter("DataRow_Keyword"),(len(DataTable.GetSheet("TestPlan").GetParameter("DataRow_Keyword"))-len(arr(ubound(arr)))-1))
							else
							strKeyWord = ""
						End If
		elseif  (Lcase(Trim(arr(ubound(arr)) )) =  "fo"   and not(Environment("FirstOnly")))  or ( Lcase(Trim(arr(ubound(arr)) )) =  "#fo#"   and not(Environment("FirstOnly")) )  then
				        Environment("Execute") = False
						Exit Function
		elseif  (Lcase(Trim(arr(ubound(arr)) )) =  "lo" and  Environment("LastOnly"))  or ( Lcase(Trim(arr(ubound(arr)) )) =  "#lo#" and  Environment("LastOnly"))   then
						If ubound(arr) > 0  Then
							strKeyWord = Left(DataTable.GetSheet("TestPlan").GetParameter("DataRow_Keyword"),(len(DataTable.GetSheet("TestPlan").GetParameter("DataRow_Keyword"))-len(arr(ubound(arr)))-1))
							else
							strKeyWord = ""
						End If
		elseif  (Lcase(Trim(arr(ubound(arr)) )) =  "lo"and not(Environment("LastOnly")))  or ( Lcase(Trim(arr(ubound(arr)) )) =  "#lo#"and not(Environment("LastOnly")) )  then
			        Environment("Execute") = False
				  	Exit Function
		elseif ( Lcase(Trim(arr(ubound(arr)) )) <>  "fo" and  Lcase(Trim(arr(ubound(arr)) )) <> "lo")  or ( Lcase(Trim(arr(ubound(arr)) )) <>  "#fo#" and  Lcase(Trim(arr(ubound(arr)) )) <> "#lo#") then
				strKeyWord = GetDynamicParameter(DataTable.GetSheet("TestPlan").GetParameter("DataRow_Keyword"))
		Else
				strKeyWord = GetDynamicParameter(DataTable.GetSheet("TestPlan").GetParameter("DataRow_Keyword"))
		End If
   	'' Scenario keys 
 
 
 
	  ' Below portion is commented for TAF 7.0.7
	'  msgbox Environment("ScenarioDataKeys") ' DSTGS 
	'	If  lcase(DataTable.GetSheet("TestPlan").GetParameter("DataRow_Keyword")) = "scenariokeys"Then  
		'	strKeyWord =  Environment("ScenarioDataKeys")
		'	else
		'	strKeyWord = GetDynamicParameter(DataTable.GetSheet("TestPlan").GetParameter("DataRow_Keyword"))
	'	End If

	arrKeyWord = split(strKeyWord, ";")
   If Ubound(arrKeyWord) > 0 Then
		Environment("DataDriven")="Yes"
		Environment("Keywords")=strKeyWord
	Else
		Environment("DataDriven")="Yes"
		Environment("Keywords")=""
    End If

 If Ubound(arrKeyWord) <= 0 and LCase(Trim(DataTable.GetSheet("TestPlan").GetParameter("DataDriven_KeyWord")))<>"yes"  Then
  		Environment("DataDriven")="no"
  		Environment("Keywords")=strKeyWord
 Else

		Environment("DataDriven")="Yes"
		Environment("Keywords")=strKeyWord
End IF

If  LCase(Trim(DataTable.GetSheet("TestPlan").GetParameter("DataDriven_KeyWord")))<>"yes"  Then
		Environment("IsDataDriven")="no"
		else
		Environment("IsDataDriven")="yes"
End If



	Environment("ExecutionStarted") = True
	Environment("FlagTCStep")=True
	Environment("DataSheetExist") = True
Environment("ExecuteStartTest") = True

End Function


Public Function runFlash()

If LCase(CStr(Environment("RunFromQC"))) = "true" Then

PathOfExe = Environment("WorkingDirectory") & "\TempAutomation\KeywordDriven\Lib\tag1.exe"

Else
PathOfExe = Environment("vegaTestSuite") & "\Lib\tag1.exe"
End If

If VerifyFileExists(PathOfExe) Then
If not Environment("tagCount") Then
Systemutil.Run (PathOfExe)
Wait 2
Systemutil.CloseProcessByName("tag1.exe")
Environment("tagCount")=TRUE 
End If
Else
Reporter.ReportEvent micWarning, "Run Flag", "Flag Not found. Contact DSTWS"
'ExitTest
End If
End Function
