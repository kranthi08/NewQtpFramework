''
' @# author DSTWS
' @# Version 7.0.7a
' @# Version 7.0.8  : Handles 'Environment("FlagTCStep")=false'  based on "returnresult"
' @# Version 7.0.9 : Added new function 'waitscreencapture' to get the screen shot for wait statement
' @# Version 7.0.9 : Modified 'wait' case in 'utilityactivity' function to  get the screen shot


	''
	' This function BusinessActivity includes the logic to call the VBS Business Activities
	' @author DSTWS
	' @param functionName String Specifying name of the Activity to perform on the UI object.
	' @param ExpectedResult String Specifying the expected result
    ' @param Param1 String Specifying the Test Data from the excel sheet
	' @param Param2 String is an optional and non mandatory field
	' @param Param3 String is an optional and non mandatory field
	' @param Param4 String is an optional and non mandatory field
	' @param Param5 String is an optional and non mandatory field
	' @to do Provide selct case statement inf future if we need to add business logic components
	' @Modified By DT77734
	' @Modified on: 16 Jul 2009 
	' @Modified by DT 77471 on 19 Mar 2010: Add the function to get the screen capture for Wait statement @Bug 58@ line 342, 62
	' @Modified by DT 77471 on 24 Mar 2010: Modified Wait case in utility activity to get the screenshot @Bug 58
	' @Modified by DT 77471 on 19 Mar 2010: Handle Pop-up recovery with showStopper run status
 
Public Function BusinessActivity(functionName,ExpectedResult,Param1,Param2,Param3,Param4,Param5,Param6,Param7,Param8,Param9,Param10)
	Dim aResult(0,1)   
	Dim addResult
	addResult="no"   'Provide selct case statement inf future if we need to add business logic components 
	aResult(0,0)=addResult     
	aResult(0,1)=returnResult  
	BusinessActivity=aResult
End Function



	''
	' This function UtilityActivity includes the logic to call the VBS Utility Activities
	' @author DSTWS
	' @param functionName String Specifying name of the Activity to perform on the UI object.
	' @param ExpectedResult String Specifying the expected result
    ' @param Param1 String Specifying the Test Data from the excel sheet
	' @param Param2 String is an optional and non mandatory field
	' @param Param3 String is an optional and non mandatory field
	' @param Param4 String is an optional and non mandatory field
	' @param Param5 String is an optional and non mandatory field
	' @to do Provide selct case statement inf future if we need to add business logic components
	' @Modified By DT77734
	' @Modified on: 16 Jul 2009 

Public Function UtilityActivity(functionName,ExpectedResult,Param1,Param2,Param3,Param4,Param5,Param6,Param7,Param8,Param9,Param10,stepResultVar)
	Dim aResult(0,1)   
	Dim addResult
	addResult="no"



					If  Environment("BlnCodeGeneration")=1 Then
													
									Select Case Trim(Lcase(functionName))
												Case"launchapplication"
													Apppath=Param1
													UpdateScript("return=LaunchApplication"&"("&""""&Apppath&""""&")")
													'return10=LaunchApplication(Apppath)
													addResult="yes"  'The flag is set to add result to the dictionary object. Only when the flag is set to yes'Results' will be added.
													returnResult=return1   
												Case"sendkey"
													KeyValue=Param1                                      
													return2=SendKey(KeyValue)                           
													addResult="yes"                                            
													returnResult=return2
												Case"wait"                                                        
													maxTime=Param1  
													Wait maxTime 
													WriteToScript("Wait "&maxTime)  
													UpdateScript("Wait "&maxTime)                                   
													'return2=waitscreencapture(maxTime)     'To get the screen capture for wait statement                      
													'addResult="yes"                                            
													'returnResult=return2 
												Case"addvalues2"
													var1=Param1
													var2=Param2          
													return008=Addvalues2(var1,var2)      
													addResult="yes"
													returnResult=return008
												Case"getcharsfromleft"
													sentence=Param1
													numofcharacters=Param2
													return3=GetCharsFromLeft(sentence,numofcharacters) 
													addResult="yes"
													returnResult=return3
												Case"getcharsfromright"
													sentence=Param1
													numofcharacters=Param2
													return4=GetCharsFromRight(sentence,numofcharacters) 
													addResult="yes"
													returnResult=return4
												Case"findwordinsentence"
													sentence=Param1
													word=Param2
													numofcharacters=Param2
													If Len(Trim(numofcharacters))=0 Then
														numofcharacters=""
													End If
													return5=FindWordInstrSentence(sentence,word,numofcharacters) 
													addResult="yes"
													returnResult=return5 
												Case"stringcompare"
													If Len(Trim(Param1))=0 Then
														string1=Param2
														string2=Param3
														expResult=Param4
														compareOption=Param5
													Else
														string1=Param1
														string2=Param2
														expResult=Param3
														compareOption=Param4
													End If
													If Len(Trim(compareOption))=0 Then
														compareOption=0
													End If
													return6=StringCompare(string1,string2,compareOption,expResult,result)
													addResult="yes"
													returnResult=return6
												Case"replacecharsinstring"
													sentence=Param1
													CharToReplace=Param2
													ReplaceWithChar=Param3
													If Len(Trim(Param3))=0 Then
														ReplaceWithChar=""
													End If    
													return7=ReplaceCharsInString(sentence,CharToReplace,ReplaceWithChar)
													addResult="yes"
													returnResult=return7
												Case"substractnumbers"
													number1=Param1
													number2=Param2
													return8=SubstractNumbers(number1,number2)
													addResult="yes"
													returnResult=return8 
												Case"comparenumbers"
													num1=Param1
													num2=Param2
													return9=CompareNumbers(num1,num2,result) 
													addResult="yes"
													returnResult=return9
												Case"getweekenddate"
													return11=GetWeekEndDate(result) 
													addResult="yes"
													returnResult=return11
												Case"comparenumbers1"
													num1=Param1
													num2=Param2
													expResult=Param3
													return12=CompareNumbers1(num1,num2,expResult,result)
													addResult="yes"
													returnResult=return12
												Case Else
													Reporter.ReportEvent micFail,"Test Aborted. The Activity name is —>"&functionName,"Input a valid function name in the Activity column."
													WriteToEvent("Fail" & vbtab & "Test Aborted. The Activity name is —>"&functionName &" Input a valid function name in the Activity column")
													aResult(0,0)=-1     
													aResult(0,1)=-1
													UtilityActivity=aResult
													Exit Function
											End Select
					Else
								Select Case Trim(Lcase(functionName))
										Case"launchapplication"
											Apppath=Param1
											return10=LaunchApplication(Apppath)
											addResult="yes"  'The flag is set to add result to the dictionary object. Only when the flag is set to yes'Results' will be added.
											returnResult=return1   
										Case"sendkey"
											KeyValue=Param1                                      
											return2=SendKey(KeyValue)                           
											addResult="yes"                                            
											returnResult=return2
										Case"wait"                                                        
											maxTime=Param1  
											Wait maxTime                                     
											'return2=waitscreencapture(maxTime)     'To get the screen capture for wait statement                      
											'addResult="yes"                                            
											'returnResult=return2 
										Case"addvalues2"
											var1=Param1
											var2=Param2          
											return008=Addvalues2(var1,var2)      
											addResult="yes"
											returnResult=return008
										Case"getcharsfromleft"
											sentence=Param1
											numofcharacters=Param2
											return3=GetCharsFromLeft(sentence,numofcharacters) 
											addResult="yes"
											returnResult=return3
										Case"getcharsfromright"
											sentence=Param1
											numofcharacters=Param2
											return4=GetCharsFromRight(sentence,numofcharacters) 
											addResult="yes"
											returnResult=return4
										Case"findwordinsentence"
											sentence=Param1
											word=Param2
											numofcharacters=Param2
											If Len(Trim(numofcharacters))=0 Then
												numofcharacters=""
											End If
											return5=FindWordInstrSentence(sentence,word,numofcharacters) 
											addResult="yes"
											returnResult=return5 
										Case"stringcompare"
											If Len(Trim(Param1))=0 Then
												string1=Param2
												string2=Param3
												expResult=Param4
												compareOption=Param5
											Else
												string1=Param1
												string2=Param2
												expResult=Param3
												compareOption=Param4
											End If
											If Len(Trim(compareOption))=0 Then
												compareOption=0
											End If
											return6=StringCompare(string1,string2,compareOption,expResult,result)
											addResult="yes"
											returnResult=return6
										Case"replacecharsinstring"
											sentence=Param1
											CharToReplace=Param2
											ReplaceWithChar=Param3
											If Len(Trim(Param3))=0 Then
												ReplaceWithChar=""
											End If    
											return7=ReplaceCharsInString(sentence,CharToReplace,ReplaceWithChar)
											addResult="yes"
											returnResult=return7
										Case"substractnumbers"
											number1=Param1
											number2=Param2
											return8=SubstractNumbers(number1,number2)
											addResult="yes"
											returnResult=return8 
										Case"comparenumbers"
											num1=Param1
											num2=Param2
											return9=CompareNumbers(num1,num2,result) 
											addResult="yes"
											returnResult=return9
										Case"getweekenddate"
											return11=GetWeekEndDate(result) 
											addResult="yes"
											returnResult=return11
										Case"comparenumbers1"
											num1=Param1
											num2=Param2
											expResult=Param3
											return12=CompareNumbers1(num1,num2,expResult,result)
											addResult="yes"
											returnResult=return12
										Case Else
											Reporter.ReportEvent micFail,"Test Aborted. The Activity name is —>"&functionName,"Input a valid function name in the Activity column."
											WriteToEvent("Fail" & vbtab & "Test Aborted. The Activity name is —>"&functionName &" Input a valid function name in the Activity column")
											aResult(0,0)=-1     
											aResult(0,1)=-1
											UtilityActivity=aResult
											Exit Function
									End Select
	
	
					End If


    
		
	If Trim(LCase(Environment("CurrentStepStatus"))) <> "failed" Then
			If LCase(Trim(Environment.Value("captureScreenshot")))="yes" Then
					strParentObject = ""
				Call ScreenCapture(strParentObject)
				If Environment("BlnCodeGeneration")=1 Then
					UpdateScript("return=ScreenCapture"&"("&""""&strParentObject&""""&")")
				End If
			End If 
	End If
	If Lcase(Environment("SaveScreenshotsinLocalDrive")) = "yes" Then
		If LCase(Trim(captureScreen))="yes" Then
			Call SaveScreenshotsinLocalDrive()
		End If
	End If		
	aResult(0,0)=addResult     
	aResult(0,1)=returnResult  
	UtilityActivity=aResult
	
	
	If returnResult= -1 Then		
		Environment("FlagTCStep")=false	
	End If	
		
End Function
        




 ''
	' This function ApplicationActivity includes the logic to call the VBS Application Activities
	' @author DT77709
	' @param appdataSource String specifying the AppMap excel sheet pah
	' @param UIName String specifying the UI Name of the object
	' @param functionName String Specifying name of the Activity to perform on the UI object.
	' @param ExpectedResult String Specifying the expected result
	' @param captureScreen String specifying whether the execution screen to capture, ex: 'Yes' or 'No'
	' @param Param1 String Specifying the Test Data from the excel sheet
	' @param Param2 String is an optional and non mandatory field
	' @param Param3 String is an optional and non mandatory field
	' @param Param4 String is an optional and non mandatory field
	' @param Param5 String is an optional and non mandatory field
	' @param oParent String is an optional and non mandatory field

Public Function ApplicationActivity(appdataSource,UIName,functionName,ExpectedResult,captureScreen,Param1,Param2,Param3,Param4,Param5,Param6,Param7,Param8,Param9,Param10,oParent,stepResultVar)
    Dim aResult(0,1)
	Dim addResult
	addResult="no" 

	ExpectedResult=cStr(ExpectedResult)
'	If Trim(ExpectedResult)="" Then
'		'ExpectedResult=""""&ExpectedResult&""""
'		ExpectedResult=cStr(ExpectedResult)
'	End If
	
'	strChildobject =  AccessAppMapDB(appdataSource,UIName)
	
'--------------------------------Get below data from Dict Objs-------------------------------------
'' Modified by Srikanth---------------TAF10
   Err.Clear
	If 	Environment("AppMapType") = "WithOutWindowName" Then
		Environment("ParentObject") = DictParentObj.Item(UIName)
		strChildobject = DictChildObj.Item(UIName)
	Else
		Environment("ParentObject") = DictParentObj.Item(Environment("WindowName")&"!"&UIName)
		strChildobject = DictChildObj.Item(Environment("WindowName")&"!"&UIName)
	End If
	
    ' Modification done------------------TAF10
     
'	If err.number<>0 and IsEmpty(strChildobject) Then 
	If strChildobject = "" Then 
		Environment("ParentObject") =""
		strChildobject = "" 
'        Call ReportResult (UIName , "UIName should exists in the AppMap", "","There are no UIName in the AppMap" , "Failed" ,Ckix3tw1q)
		Call ReportResult (UIName , "UIName should exists in the AppMap", "","There are no UIName in the AppMap" , "Failed" ,Environment("ParentObject"))

		Reporter.ReportEvent micFail,"There are no UIName specified in the AppMap","Please specify the UIName and run again"
		WriteToEvent("Fail" & vbtab & "There are no UIName specified in the AppMap")
	elseif err.number<>0 then
		  Environment("ParentObject") =""
     End If
	'----------------------------------------------------------------------------------------------------

	If Environment("ParentObject")=""Then
		Reporter.ReportEvent micFail,"Test Aborted","No object parent mentioned in AppMap.xls"
		WriteToEvent("Fail" & vbtab & "Test Aborted. No object parent mentioned in AppMap.xls")
	Else
		strParentObject=Environment("ParentObject")
	End If



	If  strChildobject = "" Then   
		Environment("FlagTCStep")=false 	
		aResult(0,0)=1  
		aResult(0,1)=1 
		ApplicationActivity=aResult  
	Else    

				
							If  Environment("BlnCodeGeneration")=1 Then
												strParentObject=ConverExprToStr(strParentObject)
												strChildobject=ConverExprToStr(strChildobject)
												

											Select Case Trim(Lcase(functionName))
														Case"settextonedit"
															ValueToEnter=Param1
															If  instr(ValueToEnter,"svar") Then
																UpdateScript("return=SetTextOnEdit"&"("&strParentObject&","&strChildobject&","&ValueToEnter&","&""""&UIName&""""&","&""""&ExpectedResult&""""&")")
															Else
															   UpdateScript("return=SetTextOnEdit"&"("&strParentObject&","&strChildobject&","&""""&ValueToEnter&""""&","&""""&UIName&""""&","&""""&ExpectedResult&""""&")")
															End If
															

											
															'addResult="yes"  'The flag is se to add result to the dictionary object. Only when the flag is set to yes'Results' will be added.
															'returnResult=return1
														Case"setpassword"
															ValueToEnter=Param1
															UpdateScript("return=SetPassword"&"("&strParentObject&","&strChildobject&","&""""&ValueToEnter&""""&","&""""&UIName&""""&","&""""&ExpectedResult&""""&")")
															'return2=SetPassword(strParentObject,strChildobject,ValueToEnter,UIName,ExpectedResult)
											'				addResult="yes"  'The flag is set to add result to the dictionary object. Only when the flag is set to yes'Results' will be added.
											'				returnResult=return2
														Case"verifylengthofedit"
															LengthToVerify=Param1
															return3=VerifyLengthOfEdit(strParentObject,strChildobject,LengthToVerify,UIName,ExpectedResult)
															addResult="yes"  'The flag is set to add result to the dictionary object. Only when the flag is set to yes'Results' will be added.
															returnResult=return3
														Case"verifyeditboxdisabled"
															return4=VerifyEditboxDisabled(strParentObject,strChildobject,UIName,ExpectedResult)
															addResult="yes"  'The flag is set to add result to the dictionary object. Only when the flag is set to yes'Results' will be added.
															returnResult=return4
														Case"verifyabsenceofeditbox"
															return5=VerifyAbsenceOfEditbox(strParentObject,strChildobject,UIName,ExpectedResult)
															addResult="yes"  'The flag is set to add result to the dictionary object. Only when the flag is set to yes'Results' will be added.
															returnResult=return5
														Case"clickonbutton"
															UpdateScript("return=ClickOnButton"&"("&strParentObject&","&strChildobject&","&""""&UIName&""""&","&""""&ExpectedResult&""""&")")
															'return6=ClickOnButton(strParentObject,strChildobject,UIName,ExpectedResult)
															addResult="yes"  'The flag is set to add result to the dictionary object. Only when the flag is set to yes'Results' will be added.
															returnResult=return6
														Case"verifyabsenceofbutton"
															return7= VerifyAbsenceOfButton(strParentObject,strChildobject,UIName,ExpectedResult)
															addResult="yes"  'The flag is set to add result to the dictionary object. Only when the flag is set to yes'Results' will be added.
															returnResult=return7
														Case"verifybuttondisabled"
															return8= VerifyButtonDisabled(strParentObject,strChildobject,UIName,ExpectedResult)
															addResult="yes"  'The flag is set to add result to the dictionary object. Only when the flag is set to yes'Results' will be added.
															returnResult=return8
														Case"verifyitemnotincombo"
															ValueToVerify=Param1
															return9=VerifyItemNotInCombo(strParentObject,strChildobject,ValueToVerify,UIName,ExpectedResult)
															addResult="yes"  'The flag is set to add result to the dictionary object. Only when the flag is set to yes'Results' will be added.
															returnResult=return9
														Case"getAallitemsincombo"
															return10=GetAllItemsInCombo(strParentObject,strChildobject,UIName,ExpectedResult)
															addResult="yes"  'The flag is set to add result to the dictionary object. Only when the flag is set to yes'Results' will be added.
															returnResult=return10
														Case"selectvalueincombo"
															ValueToSelect=Param1
															UpdateScript("return=SelectValueInCombo"&"("&strParentObject&","&strChildobject&","&""""&ValueToSelect&""""&","&""""&UIName&""""&","&""""&ExpectedResult&""""&")")
															'return11=SelectValueInCombo(strParentObject,strChildobject,ValueToSelect,UIName,ExpectedResult)
											'				addResult="yes"  'The flag is set to add result to the dictionary object. Only when the flag is set to yes'Results' will be added.
											'				returnResult=return11
														Case"selectradiobutton"
															ValueToEnter=Param1
															return12=SelectRadioButton(strParentObject,strChildobject ,ValueToEnter,UIname , ExpectedResult)
															addResult="yes"  'The flag is set to add result to the dictionary object. Only when the flag is set to yes'Results' will be added.
															returnResult=return12
														Case"selectcheckbox"
															ValueToSet=Param1
															'return13=SelectCheckBox(strParentObject,strChildobject ,ValueToSet,UIName,ExpectedResult)
															UpdateScript("return=selectcheckbox"&"("&strParentObject&","&strChildobject&","&""""&ValueToSet&""""&","&""""&UIName&""""&","&""""&ExpectedResult&""""&")")
															addResult="yes"  'The flag is set to add result to the dictionary object. Only when the flag is set to yes'Results' will be added.
															returnResult=return13
														Case"clickonimage"
															return14= ClickOnImage(strParentObject,strChildobject,UIName,ExpectedResult)
															addResult="yes"  'The flag is set to add result to the dictionary object. Only when the flag is set to yes'Results' will be added.
															returnResult=return14
														Case"clickonlink"
															return15= ClickOnLink(strParentObject,strChildobject,UIName,ExpectedResult)
															addResult="yes"  'The flag is set to add result to the dictionary object. Only when the flag is set to yes'Results' will be added.
															returnResult=return15
														Case"verifypresenceoflink"
															return16= VerifyPresenceOfLink(strParentObject,strChildobject,UIName,ExpectedResult)
															addResult="yes"  'The flag is set to add result to the dictionary object. Only when the flag is set to yes'Results' will be added.
															returnResult=return16
														Case"verifyabsenceoflink"
															return17= VerifyAbsenceOfLink(strParentObject,strChildobject,UIName,ExpectedResult)
															addResult="yes"  'The flag is set to add result to the dictionary object. Only when the flag is set to yes'Results' will be added.
															returnResult=return17
														Case"typeinactivex"
															ValueToEnter=Param1
															UpdateScript("return=TypeInActiveX"&"("&strParentObject&","&strChildobject&","&""""&ValueToEnter&""""&","&""""&UIName&""""&","&""""&ExpectedResult&""""&")")
															'return18=TypeInActiveX(strParentObject,strChildobject,ValueToEnter,UIName,ExpectedResult)
											'				addResult="yes"  'The flag is set to add result to the dictionary object. Only when the flag is set to yes'Results' will be added.
											'				returnResult=return18
														Case"findmessage"
															MessageToFind=Param1
															return19=FindMessage(strParentObject,strChildobject,MessageToFind,UIName,ExpectedResult)  
															addResult="yes"  'The flag is set to add result to the dictionary object. Only when the flag is set to yes'Results' will be added.
															returnResult=return19
														Case"selectvalueinwinmenu"
															ValueToSelect=Param1
															UpdateScript("return=SelectValueInWinMenu"&"("&strParentObject&","&strChildobject&","&""""&ValueToSelect&""""&","&""""&UIName&""""&","&""""&ExpectedResult&""""&")")
															'return20=SelectValueInWinMenu(strParentObject,strChildobject,ValueToSelect,UIName,ExpectedResult)
															addResult="yes"  'The flag is set to add result to the dictionary object. Only when the flag is set to yes'Results' will be added.
															returnResult=return20
														Case"verifytextinedit"
															textToVerify=Param1
															return21=VerifyTextInEdit(strParentObject,strChildobject,textToVerify,UIName,ExpectedResult)
															addResult="yes"  'The flag is set to add result to the dictionary object. Only when the flag is set to yes'Results' will be added.
															returnResult=return21
														Case"gettextofeditbox"
															'return22=getTextOfEditbox(strParentObject,strChildobject,UIName,ExpectedResult)
															UpdateScript(stepResultVar&"=gettextofeditbox"&"("&strParentObject&","&strChildobject&","&""""&UIName&""""&","&""""&ExpectedResult&""""&")")
															UpdateScript("Call BuildDict"&"("&stepResultVar&","&""""&stepResultVar&""""&")")
															'Call BuildDict(returnResult,stepResultVar)
															addResult="yes"  'The flag is set to add result to the dictionary object. Only when the flag is set to yes'Results' will be added.
															returnResult=return22
															Case"i"
															 return22=IMode(strParentObject,strChildobject,Param1,Param2,Param3,Param4,Param5,Param6,Param7,Param8,Param9,Param10,UIName,ExpectedResult)
															addResult="yes"
															returnResult=return22
														Case"s"
															return23=SMode(strParentObject,strChildobject,Param1,Param2,Param3,Param4,Param5,Param6,Param7,Param8,Param9,Param10,UIName,ExpectedResult)
															addResult="yes"
															returnResult=return23
														Case"v"
															return24=VMode(strParentObject,strChildobject,Param1,Param2,Param3,Param4,Param5,Param6,Param7,Param8,Param9,Param10,UIName,ExpectedResult)
															addResult="yes"
															returnResult=return24
														Case"av"
															return25=AVMode(strParentObject,strChildobject,Param1,Param2,Param3,Param4,Param5,Param6,Param7,Param8,Param9,Param10,UIName,ExpectedResult)
															addResult="yes"
															returnResult=return25
														Case"vl"
															return26=VLMode(strParentObject,strChildobject,Param1,Param2,Param3,Param4,Param5,Param6,Param7,Param8,Param9,Param10,UIName,ExpectedResult)
															addResult="yes"
															returnResult=return26
														Case"wp"
															return27=WPMode(strParentObject,strChildobject,Param1,Param2,Param3,Param4,Param5,Param6,Param7,Param8,Param9,Param10,UIName,ExpectedResult)
															addResult="yes"
															returnResult=return27
													   Case Else
															aResult(0,0)=-1
															aResult(0,1)=-1
															ApplicationActivity=aResult
															Reporter.ReportEvent micFail,"Test Aborted. The Activity name is —>"&functionName,"Input a valid function name in the Activity column."
															Call ReportResult  (functionName,"Input a valid function name in the Activity column.", functionName , "Test Aborted. The Activity name is —>"&functionName, "Failed", UIName)
															WriteToEvent("Fail" & vbtab & "Test Aborted. The Activity name is —>"&functionName&"Input a valid function name in the Activity column")
															Exit Function
													End Select
							Else
											Select Case Trim(Lcase(functionName))
													Case"settextonedit"
														ValueToEnter=Param1
														return1=SetTextOnEdit(strParentObject,strChildobject,ValueToEnter,UIName,ExpectedResult)
														addResult="yes"  'The flag is se to add result to the dictionary object. Only when the flag is set to yes'Results' will be added.
														returnResult=return1
													Case"setpassword"
														ValueToEnter=Param1
														return2=SetPassword(strParentObject,strChildobject,ValueToEnter,UIName,ExpectedResult)
														addResult="yes"  'The flag is set to add result to the dictionary object. Only when the flag is set to yes'Results' will be added.
														returnResult=return2
													Case"verifylengthofedit"
														LengthToVerify=Param1
														return3=VerifyLengthOfEdit(strParentObject,strChildobject,LengthToVerify,UIName,ExpectedResult)
														addResult="yes"  'The flag is set to add result to the dictionary object. Only when the flag is set to yes'Results' will be added.
														returnResult=return3
													Case"verifyeditboxdisabled"
														return4=VerifyEditboxDisabled(strParentObject,strChildobject,UIName,ExpectedResult)
														addResult="yes"  'The flag is set to add result to the dictionary object. Only when the flag is set to yes'Results' will be added.
														returnResult=return4
													Case"verifyabsenceofeditbox"
														return5=VerifyAbsenceOfEditbox(strParentObject,strChildobject,UIName,ExpectedResult)
														addResult="yes"  'The flag is set to add result to the dictionary object. Only when the flag is set to yes'Results' will be added.
														returnResult=return5
													Case"clickonbutton"
														return6=ClickOnButton(strParentObject,strChildobject,UIName,ExpectedResult)
														addResult="yes"  'The flag is set to add result to the dictionary object. Only when the flag is set to yes'Results' will be added.
														returnResult=return6
													Case"verifyabsenceofbutton"
														return7= VerifyAbsenceOfButton(strParentObject,strChildobject,UIName,ExpectedResult)
														addResult="yes"  'The flag is set to add result to the dictionary object. Only when the flag is set to yes'Results' will be added.
														returnResult=return7
													Case"verifybuttondisabled"
														return8= VerifyButtonDisabled(strParentObject,strChildobject,UIName,ExpectedResult)
														addResult="yes"  'The flag is set to add result to the dictionary object. Only when the flag is set to yes'Results' will be added.
														returnResult=return8
													Case"verifyitemnotincombo"
														ValueToVerify=Param1
														return9=VerifyItemNotInCombo(strParentObject,strChildobject,ValueToVerify,UIName,ExpectedResult)
														addResult="yes"  'The flag is set to add result to the dictionary object. Only when the flag is set to yes'Results' will be added.
														returnResult=return9
													Case"getAallitemsincombo"
														return10=GetAllItemsInCombo(strParentObject,strChildobject,UIName,ExpectedResult)
														addResult="yes"  'The flag is set to add result to the dictionary object. Only when the flag is set to yes'Results' will be added.
														returnResult=return10
													Case"selectvalueincombo"
														ValueToSelect=Param1
														return11=SelectValueInCombo(strParentObject,strChildobject,ValueToSelect,UIName,ExpectedResult)
														addResult="yes"  'The flag is set to add result to the dictionary object. Only when the flag is set to yes'Results' will be added.
														returnResult=return11
													Case"selectradiobutton"
														ValueToEnter=Param1
														return12=SelectRadioButton(strParentObject,strChildobject ,ValueToEnter,UIname , ExpectedResult)
														addResult="yes"  'The flag is set to add result to the dictionary object. Only when the flag is set to yes'Results' will be added.
														returnResult=return12
													Case"selectcheckbox"
														ValueToSet=Param1
														return13=SelectCheckBox(strParentObject,strChildobject ,ValueToSet,UIName,ExpectedResult)
														addResult="yes"  'The flag is set to add result to the dictionary object. Only when the flag is set to yes'Results' will be added.
														returnResult=return13
													Case"clickonimage"
														return14= ClickOnImage(strParentObject,strChildobject,UIName,ExpectedResult)
														addResult="yes"  'The flag is set to add result to the dictionary object. Only when the flag is set to yes'Results' will be added.
														returnResult=return14
													Case"clickonlink"
														return15= ClickOnLink(strParentObject,strChildobject,UIName,ExpectedResult)
														addResult="yes"  'The flag is set to add result to the dictionary object. Only when the flag is set to yes'Results' will be added.
														returnResult=return15
													Case"verifypresenceoflink"
														return16= VerifyPresenceOfLink(strParentObject,strChildobject,UIName,ExpectedResult)
														addResult="yes"  'The flag is set to add result to the dictionary object. Only when the flag is set to yes'Results' will be added.
														returnResult=return16
													Case"verifyabsenceoflink"
														return17= VerifyAbsenceOfLink(strParentObject,strChildobject,UIName,ExpectedResult)
														addResult="yes"  'The flag is set to add result to the dictionary object. Only when the flag is set to yes'Results' will be added.
														returnResult=return17
													Case"typeinactivex"
														ValueToEnter=Param1
														return18=TypeInActiveX(strParentObject,strChildobject,ValueToEnter,UIName,ExpectedResult)
														addResult="yes"  'The flag is set to add result to the dictionary object. Only when the flag is set to yes'Results' will be added.
														returnResult=return18
													Case"findmessage"
														MessageToFind=Param1
														return19=FindMessage(strParentObject,strChildobject,MessageToFind,UIName,ExpectedResult)  
														addResult="yes"  'The flag is set to add result to the dictionary object. Only when the flag is set to yes'Results' will be added.
														returnResult=return19
													Case"selectvalueinwinmenu"
														ValueToSelect=Param1
														return20=SelectValueInWinMenu(strParentObject,strChildobject,ValueToSelect,UIName,ExpectedResult)
														addResult="yes"  'The flag is set to add result to the dictionary object. Only when the flag is set to yes'Results' will be added.
														returnResult=return20
													Case"verifytextinedit"
														textToVerify=Param1
														return21=VerifyTextInEdit(strParentObject,strChildobject,textToVerify,UIName,ExpectedResult)
														addResult="yes"  'The flag is set to add result to the dictionary object. Only when the flag is set to yes'Results' will be added.
														returnResult=return21
													Case"gettextofeditbox"
														return22=getTextOfEditbox(strParentObject,strChildobject,UIName,ExpectedResult)
														addResult="yes"  'The flag is set to add result to the dictionary object. Only when the flag is set to yes'Results' will be added.
														returnResult=return22
														Case"i"
														 return22=IMode(strParentObject,strChildobject,Param1,Param2,Param3,Param4,Param5,Param6,Param7,Param8,Param9,Param10,UIName,ExpectedResult)
														addResult="yes"
														returnResult=return22
													Case"s"
														return23=SMode(strParentObject,strChildobject,Param1,Param2,Param3,Param4,Param5,Param6,Param7,Param8,Param9,Param10,UIName,ExpectedResult)
														addResult="yes"
														returnResult=return23
													Case"v"
														return24=VMode(strParentObject,strChildobject,Param1,Param2,Param3,Param4,Param5,Param6,Param7,Param8,Param9,Param10,UIName,ExpectedResult)
														addResult="yes"
														returnResult=return24
													Case"av"
														return25=AVMode(strParentObject,strChildobject,Param1,Param2,Param3,Param4,Param5,Param6,Param7,Param8,Param9,Param10,UIName,ExpectedResult)
														addResult="yes"
														returnResult=return25
													Case"vl"
														return26=VLMode(strParentObject,strChildobject,Param1,Param2,Param3,Param4,Param5,Param6,Param7,Param8,Param9,Param10,UIName,ExpectedResult)
														addResult="yes"
														returnResult=return26
													Case"wp"
														return27=WPMode(strParentObject,strChildobject,Param1,Param2,Param3,Param4,Param5,Param6,Param7,Param8,Param9,Param10,UIName,ExpectedResult)
														addResult="yes"
														returnResult=return27
												   Case Else
														aResult(0,0)=-1
														aResult(0,1)=-1
														ApplicationActivity=aResult
														Reporter.ReportEvent micFail,"Test Aborted. The Activity name is —>"&functionName,"Input a valid function name in the Activity column."
														Call ReportResult  (functionName,"Input a valid function name in the Activity column.", functionName , "Test Aborted. The Activity name is —>"&functionName, "Failed", UIName)
														WriteToEvent("Fail" & vbtab & "Test Aborted. The Activity name is —>"&functionName&"Input a valid function name in the Activity column")
														Exit Function
												End Select
							End If

			
		

		If Trim(LCase(Environment("CurrentStepStatus"))) <> "failed" Then
					If LCase(Trim(captureScreen))="yes" Then
						Call ScreenCapture(strParentObject)
						If  Environment("BlnCodeGeneration")=1Then
                            UpdateScript("return=ScreenCapture"&"("&""&strParentObject&""&")")
						End If
						
					End If       
		End If
		
		If Lcase(Environment("SaveScreenshotsinLocalDrive")) = "yes" Then
			If LCase(Trim(captureScreen))="yes" Then
				Call SaveScreenshotsinLocalDrive()
			End If
		End If	
		aResult(0,0)=addResult
		aResult(0,1)=returnResult
		ApplicationActivity=aResult
		
		CheckVarType = VarType(returnresult)
		If CheckVarType=2 Then
			If returnresult = -1 then
				Environment("FlagTCStep")=false	
				If Environment("RepeatStep") and Environment("Counter")=1 Then		'To handle the recovery with show stopper
					Environment("FlagTCStep")=True	
				End If
			End If
		elseif CheckVarType=11 then
			If returnresult = 0 then
				Environment("FlagTCStep")=false	
				If Environment("RepeatStep") and Environment("Counter")=1 Then		'To handle the recovery with show stopper
					Environment("FlagTCStep")=True	
				End If
	
			End If
		End If			
	End If    
End Function




' This function is to Execute the prerequisite script
' @author Rajesh Kumar Tatavarthi
' @return VOID
'@modified by Sreenu Babu on 20 Nov 2009 

Function ExecutePreReqScript(strPrerequisites)
arrPrerequisites = Split(strPrerequisites,";")
For i= lbound(arrPrerequisites) to ubound(arrPrerequisites)
Select Case Trim(Lcase(arrPrerequisites(i)))
Case "tafbasestate"
	
'call loadvariablesfromsourcefile(Environment("Configuration"),Environment("SourceFile"),Environment("DatabaseName"))
Case "printmsg"
Print "message-Testing base sate Concept"
Case else
Reporter.ReportEvent micFail,"Running the Base state function:" & Trim(Lcase(arrPrerequisites(i))),"Base state function not defined"
WriteToEvent("Fail" & vbtab & "Running the Base state function:" & Trim(Lcase(arrPrerequisites(i))) & "--- Base state function not defined")
End Select
Next
End Function

Public function waitscreencapture(maxTime)	'To get the screenshot for the wait statement
   Wait maxTime
   Reporter.ReportEvent micPass,"Wait Time "&maxTime,"Waited time: "&maxTime
   If  UCASE(Environment("captureScreenshot"))="YES" Then
	   Call ReportResult  ("Wait Time "&maxTime, "Wait Time "&maxTime, "", "", "Screenshot","")
   End If
    	
End Function

