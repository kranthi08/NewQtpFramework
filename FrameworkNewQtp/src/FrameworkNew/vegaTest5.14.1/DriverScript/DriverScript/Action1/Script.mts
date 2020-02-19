'********************************************************************************************************************************************************************************
' Description        : This Action call the specific testcase to execute based on the name of the test case
' author                  : Rajesh Kumar Tatavarthi,Sreenu Babu
' Reviewed By   : Sreenu Babu
' TAF Version     :  TAF 7.0.7
' Modified By     : 
' Modified on    : 
'*********************************************************************************************************************************************************************************
'Option Explicit
Dim blnIsScenario
Dim strTestScriptName
Dim strRetVal

'On error resume next
Environment("TestControllerPath") =  Environment("M1_TestControllerPath")
'GeneralBaseState
Environment("BaseState") = "N/A"
blnIsScenario =False    ' Comment this when executing the Test Scenario / Test Flow
'blnIsScenario=True       'Comment this when executing the Test  Case
strTestScriptName="SampleTestCase"
'strTestScriptName="FlightinsertandVerify"
strRetVal=CallDriver(strTestScriptName,blnIsScenario)           '   calling the driver (Core Engine)
If strRetVal=-1 Then
	ExitText	
End If


'******************************************************* End Of The Script ******************************************************************************************************