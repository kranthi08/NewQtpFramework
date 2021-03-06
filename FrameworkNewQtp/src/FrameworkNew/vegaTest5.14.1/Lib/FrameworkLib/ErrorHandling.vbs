''
' @# author DSTWS
' @# Version 7.0.7a
' @# Version 7.0.9 : Added new function  'PopupRecoveryFunction' to handle popups with repeat same step post recovery
' @# Version 7.0.9A Environment.Value("RepeatStep")=TRUE 	Environment.Value("Counter")=1 , keys added

' This Funcgtion is to set the Custom error messages set by the user
'@ Author DT77215
'@param intErrorNumber  integer - specifying the error number 
'@param strReportEvent1 String  - specifying the event that need to be reported
Function ErrorHandler(intErrorNumber, strReportEvent1,strReportEvent2)
   If  Environment("RunFromQC")   Then
       ErrorListSource=   Environment("TempExceptionHandling")   & "\ErrorListBook.xls"
	   else
	arrTemp = Split( Environment("TestPlanPath") ,"\Controller")
	ErrorListSource=Environment("vegaTestSuite") & "\ExceptionHandling\ErrorListBook.xls"
   End If
'  	ErrorListSource= "E:\Automation_123\KeywordDriven 7_0\ExceptionHandling\ErrorListBook.xls"
	DefaultDescTaken=False
'  ErrorListSource=Environment("ErrorListSource")
	Set objTDConn1=CreateObject("ADODB.Connection")
	If Err.Number<>0 Then
		'Reporter.ReportEvent micWarning,"Error occured while connecting to the Error Handling file",Err.Description
		'WriteToEvent("Warning" & vbtab & "Error occured while connecting to the Error Handling file")
		Err.Clear
		'Set objTDConn1=Nothing
	End If
	objTDConn1.ConnectionString= "DRIVER={Microsoft Excel Driver (*.xls)};DBQ="&ErrorListSource & ";Readonly=True"
	objTDConn1.Open
	Set objRecordset=CreateObject("ADODB.Recordset")
	ObjRecordset.CursorLocation=3                     
	QueryString1="Select * From [ErrorsList$] Where [Error_Code]="&intErrorNumber
	ObjRecordset.Open QueryString1,objTDConn1, 1, 3
	If ObjRecordset.recordcount>0 then   
		ErrNumber=ObjRecordset.Fields("Error_Code")	
        ErrDescription=ObjRecordset.Fields("Custom_Description")
		ErrSeverity=ObjRecordset.Fields("Severity")		
		If isnull(ErrDescription) Then
			ErrDescription=ObjRecordset.Fields("Default_Description")	 	
            DefaultDescTaken=True
			If isnull(ErrDescription) Then
				ErrDescription=""				
				ErrorHandler=ErrDescription		
			End If
        End If
		If isnull(ErrSeverity) Then
			ErrSeverity=""
		end if	
		If not(DesfaultDescTaken) Then
			If isnull(ErrDescription) Then
					ErrDescription=""				
					ErrorHandler=ErrDescription								
				Else
				If strReportEvent2<>"" Then
					arrDesc=Split(ErrDescription,"&&",-1,vbtextcompare)	
					If ubound(arrDesc)>1 Then
						On error Resume Next			
						ErrDescription=arrDesc(0)& "'"&strReportEvent1&"'" &arrDesc(1)&strReportEvent2 &" "&arrDesc(2) & "("&intErrorNumber&")"
						err.clear
						ErrorHandler=ErrDescription
					Else					
						On error Resume Next			
						ErrDescription=arrDesc(0)& "'"&strReportEvent1&"'" &arrDesc(1)&strReportEvent2& "("&intErrorNumber&")"
						err.clear
						ErrorHandler=ErrDescription			
					End If
				Else 
						arrDesc=Split(ErrDescription,"&&",-1,vbtextcompare)			
						On error Resume Next			
						ErrDescription=arrDesc(0)& "'"&strReportEvent1&"'" &arrDesc(1)& "("&intErrorNumber&")"
						err.clear
						ErrorHandler=ErrDescription
				End If		
			End If
		End If
		If Trim(Lcase(ErrSeverity))="showstopper" Then
				' set some variable to indicate it as showstopper
				Environment("FlagStepFailureOccured")=True
				Environment("strTCSeverity")="showstopper"
		End If				
	Else
	    Reporter.ReportEvent micFail,"There is no Data specified in the Error Handling List File","Please specify the Error ID and run again"
		WriteToEvent("Fail" & vbtab & "There is no Data specified in the Error Handling List File")
		ErrorHandler=-1
		Set ObjRecordset=Nothing
		objTDConn1.close
		Set objTDConn1=Nothing  
		ExitTest
	End If
	Set ObjRecordset=Nothing
	objTDConn1.close
	Set objTDConn1=Nothing   
End Function
 
Function PopupRecoveryFunction(Object)
On error resume next
	'msgbox "hello Function"
    	
	Environment.Value("RepeatStep")=TRUE
	Environment.Value("Counter")=1
	Reporter.Filter = rfDisableAll 
 
	On error goto 0
 
End Function 
 
