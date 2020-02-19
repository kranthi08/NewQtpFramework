''
' @# author DSTWS
' @# Version 7.0.7
	
	''
	'  This function is to verify the existance of a file in the system disk.
	' @author DT77734
	' @param strFilePath String specifying the path of the verifying file.

Function VerifyFileExists(strFilePath)
   Dim fso
   Set fso=CreateObject("Scripting.FileSystemObject")
   If fso.FileExists(strFilePath) Then
	   VerifyFileExists=1
	else
		VerifyFileExists=0
   End If
End Function

'TAF 10.1 new code Start. Added by Trao
Function ConvertExprToStr(strExpression)

			arrExpr=Split(strExpression,"""")

			For intCounter=0 to UBound(arrExpr)-1
			
				strConverted=strConverted + arrExpr(intCounter)&""""
			
			Next

			strConverted=strConverted + arrExpr(UBound(arrExpr))
			ConvertExprToStr=strConverted


End Function
'TAF 10.1 new code End

'TAF 10.1 new code Start. Added by Trao
Function ConverExprToStr(strExpr)

			arrExpr=Split(strExpr,".")

			For intCount=0 to UBound(arrExpr)-1
					strComplete=Convert(arrExpr(intCount))
                    strComplete1=strComplete1 & strComplete&"&"&"""."""&"&"
			Next
					strComplete1=strComplete1 & Convert(arrExpr(UBound(arrExpr)))
					ConverExprToStr=strComplete1

End Function
'TAF 10.1 new code End

'TAF 10.1 new code Start. Added by Trao
Function Convert(strExpr)

			arrExpr=Split(strExpr,"(""")

			For intCounter=0 to UBound(arrExpr)-1

				arrExpr(intCounter)=Replace(arrExpr(intCounter),""")","")

				If Instr(1,arrExpr(intCounter),"""") Then
					strConverted=strConverted &""""""&arrExpr(intCounter)&""""""
				Else
					strConverted=strConverted &""""&arrExpr(intCounter)&""""
				End If
					
    			
			Next

			strStringwithquote= arrExpr(UBound(arrExpr))

			If Instr(1,strStringwithquote,"""") Then
					strStringwithquote=Replace(strStringwithquote,""")","")
					strConverted=strConverted&"&"&"""("""&"&"
					strConverted=strConverted &""""""""&strStringwithquote&""""""""
					strConverted=strConverted&"&"&""")"""
				Else
					strConverted=strConverted &""""&strStringwithquote&""""
				End If
			


			Convert=strConverted

End Function
'TAF 10.1 new code End	

