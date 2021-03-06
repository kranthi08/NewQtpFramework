''
' This Libary is to work with VB functions.
' @author DSTWS
' @Version 7.0.0

	''
	' This function is to compare two strings.
	' @author DSTWS
	' @param strString1 is a String specifying the first string for comparision.
	' @param strString2 is a String specifying the second string for comparision.
	' @param intCompareOption is a numeric value specifying the kind of comparision.
	' @param strExpResult String specifying the expected result.
	' @param strResult String specifying the result.
	' @return string value if comparision is successful. Returns -1 if  length of strString1 or strString2 is 0.
	' @Todo  need to update function parameters(result parameter) and calling of ReportResultfunction is not correct as incorrect paramets are passed.
	' @Modified By DT77742
	' @Modified on: 16 Jul 2009 

Function StringCompare(strString1,strString2,intCompareOption,strExpResult,strResult)
	strString1=Trim(strString1)
	strString2=Trim(strString2)
	Call ReportResult(strResult)  'encountering error.
	If strExpResult="less" Then
		strExpResult=-1
	ElseIf strExpResult="equal" Then
		strExpResult=0
	Else
		strExpResult=1
	End If
	If Len(strString1)<>0 And Len(strString2)<>0 Then
		status=StrComp(strString1,strString2,intCompareOption)
		If status=strExpResult Then
			StringCompare="less than"
			Reporter.ReportEvent micPass,"Test Pass","The string"&strString1&" is less than"&strString2
		ElseIf status=strExpResult Then
			StringCompare="equal"        
			Reporter.ReportEvent micPass,"Test Pass","The string"&strString1&" is equal to"&strString2
		ElseIf status=strExpResult Then
			StringCompare="greater than"
			Reporter.ReportEvent micPass,"Test Pass","The string"&strString1&" is greater than"&strString2
		End If
	Else
		Reporter.ReportEvent micFail,"Step Fail","The length of strString1 or strString2 is 0"
		StringCompare = -1
	End If    
End Function



	''
	' This function is to compare two numbers.
	' @author DSTWS
	' @param intNum1 is a integer specifying the first number for comparision.
	' @param intNum2 is a integer specifying the second number for comparision.
    ' @param ExpectedResult String specifying the expected result.
	' @param strResult String specifying the strResult.
	' @return string value if comparision is successful. Returns -1 if  length of intNum1 or intNum2 is 0.
	' @Todo  need to update function parameters(strResult parameter)
	' @Modified By DT77742
	' @Modified on: 16 Jul 2009 

Function CompareNumbers1(intNum1,intNum2,strExpResult,strResult)
	intNum1=Trim(intNum1)
	intNum2=Trim(intNum2)
	intNum1=CDbl(intNum1)
	intNum2=CDbl(intNum2)
	Call ReportResult(strResult)
	If strExpResult="less" Then
		strExpResult=-1
	ElseIf strExpResult="equal" Then
		strExpResult=0
	Else
		strExpResult=1
	End If
	If Len(intNum1)="" Then
		intNum1=0
	End If
	If intNum1<>0 And Len(intNum2)<>0 Then
		If intNum1<intNum2 And strExpResult=-1 Then
			CompareNumbers1="less than"
			Reporter.ReportEvent micPass,"Test Pass","The Number "&intNum1&" is less than"&intNum2
		ElseIf intNum1=intNum2 And strExpResult=0 Then
			CompareNumbers1="equal"        
			Reporter.ReportEvent micPass,"Test Pass","The Number"&intNum1&" is equal to"&intNum2
		ElseIf intNum1>intNum2 And strExpResult=1 Then
			CompareNumbers1="greater than"
			Reporter.ReportEvent micPass,"Test Pass","The Number"&intNum1&" is greater than"&intNum2
		End If
	Else
		Reporter.ReportEvent micFail,"Step Fail","The length of strString1 or strString2 is 0"
		CompareNumbers1 = -1
	End If    
End Function



''
	' This function is to compare two numbers.
	' @author DSTWS
	' @param intNum1 is a number specifying the first number for comparision.
	' @param intNum2 is a number specifying the second number for comparision.
    ' @param strResult String specifying the result.
	' @return string value if comparision is successful. Returns -1 if  length of intNum1 or intNum2 is 0.
	' @Todo  need to update function parameters(strResult parameter)
	' @Modified By DT77742
	' @Modified on: 16 Jul 2009 

Function CompareNumbers(intNum1,intNum2,strResult)
	On Error Resume Next
	intNum1=Trim(intNum1)
	intNum2=Trim(intNum2)
	intNum1=CDbl(intNum1)
	intNum2=CDbl(intNum2)
	If Len(intNum1)<>0 And Len(intNum2)<>0 Then
		If intNum1=intNum2 Then
			Reporter.ReportEvent micPass,"Test Pass","The Number"&intNum1&" is equal to"&intNum2
			CompareNumbers="testPass"
		Else
			Reporter.ReportEvent micFail,"Test Pass","The Number"&intNum1&" is not equal to"&intNum2
			CompareNumbers=-1
		End If
	Else
		Reporter.ReportEvent micFail,"Step Fail","The length of intNum1 or intNum1 is 0"
		CompareNumbers = -1
	End If    
End Function
 


''
	' This function is to find date difference between two dates.
	' @author DSTWS
	' @param strInterval is a String specifying the strInterval you want to use to calculate the differences between date1 and date2 .
	' @param dtDate1 is a date expression.
	' @param dtDate2 is a date expression.
	' @param firstDayOfTheWeek is a Constant that specifies the day of the week.
	' @param firstWeekOfTheYear is a Constant that specifies the first week of the year.
    ' @param strResult String specifying the strResult.
	' @Todo need to update function in "DecideActivity.vbs"
	' @Modified By DT77742
	' @Modified on: 16 Jul 2009 
	
Function DateDifference(strInterval,dtDate1,dtDate2,dtFirstDayOfTheWeek,dtFirstWeekOfTheYear,strResult)
	DateDifference=DateDiff(strInterval,dtDate1,dtDate2,dtFirstDayOfTheWeek,dtFirstWeekOfTheYear)
End Function



 ''
	' This Function Returns a specified number of characters from the left side of a string
	' @author DSTWS
	' @param strSentence is a String specifying the string from which the charecters are to be taken.
	' @param intNumOfCharacters is a Integer specifying the number of charecters are to be fetched from sentance
	' @Return String which containes the specified numofcharecters from sentance from left.
	' @Modified By DT77742
	' @Modified on: 16 Jul 2009 
    	
Function GetCharsFromLeft(strSentence,intNumOfCharacters)
	Wrd=Left(strSentence,intNumOfCharacters)
	GetCharsFromLeft=Wrd
	Reporter.ReportEvent micDone,"Fetch word from the strSentence"&strSentence,"Fetched"&Wrd&" from strSentence"&strSentence
End Function               



''
	' This Function Returns a specified number of characters from the Right side of a string
	' @author DSTWS
	' @param strSentence is a String specifying the string from which the charecters are to be taken.
	' @param intNumOfCharacters is a Integer specifying the number of charecters are to be fetched from sentance
	' @Return String which containes the specified numofcharecters from sentance from right. 
	' @Modified By DT77742
	' @Modified on: 16 Jul 2009 

Function GetCharsFromRight(strSentence,intNumOfCharacters)
	Wrd=Right(strSentence,intNumOfCharacters)
	GetCharsFromRight=Wrd
	Reporter.ReportEvent micDone,"Fetch word from the strSentence"&strSentence,"Fetched"&Wrd&" from strSentence"&strSentence
End Function 



''
	' This Function Returns a specified number of characters from a string
	' @author DSTWS
	' @param strSentence is a String specifying the string from which the charecters are to be taken.
	' @param strWord is a string specifying the word to be searched in sentance.
	' @param intNumOfCharacters is a number specifying the number of charecters are to be fetched from sentance
	' @Return String which containes the specified numofcharecters from sentance.
	' @Modified By DT77742
	' @Modified on: 16 Jul 2009 

Function FindWordInstrSentence(strSentence,strWord,intNumOfCharacters)
	strSentence=Trim(strSentence)
	strWord=Trim(strWord)
	If Len(Trim(intNumOfCharacters))=0 Then
		wrd=Mid(Trim(strSentence),InStr(strSentence,strWord)+Len(strWord)+1)  
	Else
		wrd=Mid(Trim(strSentence),InStr(strSentence,strWord)+Len(strWord)+1,intNumOfCharacters)
	End If
	FindWordInstrSentence=wrd
	Reporter.ReportEvent micDone,"Fetch word from the strSentence"&strSentence,"Fetched"&Wrd&" from strSentence"&strSentence
End Function   



''
	' This Function replaces the charecters in string.
	' @author DSTWS
	' @param strSentence is a String expression containing substring to replace
	' @param strCharToReplace is a string specifying the substring being searched for.
	' @param strReplaceWithChar is a String specifying the replacement substring.
	' @Return String with the replaced string.
	' @Modified By DT77742
	' @Modified on: 16 Jul 2009 

Function ReplaceCharsInString(strSentence,strCharToReplace,strReplaceWithChar)
	Wrd=Replace(strSentence,strCharToReplace,strReplaceWithChar)
	ReplaceCharsInString=Wrd
	Reporter.ReportEvent micDone,"Replaced Chars In String:"&strSentence,"with"&strReplaceWithChar
End Function         
 


''
	' This Function performs subtraction of two numbers.
	' @author DSTWS
	' @param intNumber1 is a Integer specifying the first number.
	' @param intNumber2 is a Integer specifying the second number.
	' @Return number which contains difference of number1 and number2.
	' @Modified By DT77742
	' @Modified on: 16 Jul 2009 

Function SubstractNumbers(intNumber1,intNumber2)
	On Error Resume Next
	intNumber1=CDbl(intNumber1)
	intNumber2=CDbl(intNumber2)
	intRes=intNumber1-intNumber2
	intSubstractNumbers=intRes
	Reporter.ReportEvent micDone,"Substracted the numbers:"&intNumber1 &"and"&intNumber2,"and the result is "&intRes
	err.clear
End Function              



''
	' This Function performs subtraction of two numbers.
	' @author DSTWS
    ' @Return date specifying the weekend date.
	' @Modified By DT77742
	' @Modified on: 16 Jul 2009 

Function GetWeekEndDate(strResult)
	For i= 0 to 6
		If weekday(FormatDateTime(Date()+i))=7 then   
			GetWeekEndDate=FormatDateTime(Date()+i)
			Exit For
		End If  
	Next
End Function

