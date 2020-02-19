
'------------------------------    Global variable declaration --------------------------------------------------------------
Dim intDataiteration' variable to iterate the datatable
Dim intNode'Variable to hold the no of nodes present in a XML tree
Dim  arrayMatrix()' array to hold the data from XML
intNode=0 
Environment("Dataiteration")=1
Environment("FullData")=""
'---------------------------------  End Of  Global variable declaration --------------------------------------------------------------

 ' **********************************************************************************************************************************************************************************************************
	' Function Name:		ReadXMLOR
	' Purpose:			 This function is used to Read the QTP's OR XML tree using DotNetFactory object and  get the required data to build the array 
	' Input Parameters:		strAppMapSheetPath,strXMLpath
	' Return Values:	 No   
	' Date Created:	   22-Apr-2009 	
	' Created By:			Rajesh Kumar Tatavarthi, Subash Chakravarthi, Mohan Kakarla,Srinivas P,Ra
	' Reviewed By:			
	' ************************************************************************************************************ ********************************************************************************************** 
Function ReadXMLOR(strAppMapSheetPath,strXMLpath,AppendMode)
   Dim objReader'object  for xml reader class
   Dim objXmlReader'object to read the XML tree
	 Dim intDepth'variable to hold the  depth of the XML
	 Dim intMaxNodes'Variable to hold the no of nodes present in a XML tree
	 int intMaxDepth''variable to hold the maximum depth of the XML
     DataTable.AddSheet("KeywordRepository")'Add a temporary sheet
     DataTable.ImportSheet strAppMapSheetPath,"KeywordRepository","KeywordRepository"'Import the excel sheet to datatable
    If AppendMode=True Then
       Environment("Dataiteration")=DataTable.GetSheet("KeywordRepository").GetRowCount+1
	End If
	
	Set objReader = DotNetFactory.CreateInstance("System.Xml.XmlReader", "System.Xml")' creates an instance for XMLReader in dotnet factory
	Set objXmlReader=objReader.Create(strXMLpath)' creates a reader object  by reading the file
    intDepth=GetDepth(objXmlReader)' function call 
    Redim arraymatrix(intDepth,intNode)'Rebuild the array based on the depth
    set objXmlReader=objReader.Create(strXMLpath)' creates a reader object  by reading the file
	 while (objXmlReader.ReadToFollowing("qtpRep:Object"))' Reads the xml till end of the file
	
             ' If  objXmlReader.NodeType = "XmlDeclaration" or objXmlReader.NodeType = "Element"Then' checks if the tag type is XmlDeclaration or Element

				   ' If objXmlReader.Name="qtpRep:Object" Then' checks If element name 
				   
						 
						 If objXmlReader.AttributeCount > 0 Then' check if the element contains  attributes 
						    nme=objXmlReader.GetAttribute("Name")&"';"'attribute value
						     cls= objXmlReader.GetAttribute("Class")&"';"'attribute value 								  
						     arrayMatrix(Cint(objXmlReader.Depth),intNode)=nme&cls'assign the attributes to array
							 intNode=intNode+1   ' increment the  intNode for max nodes
							 redim preserve  arrayMatrix(intDepth ,  intNode)' rebuild the array with preserving the old contents
				  
						 End If
					
				
				' End If
	 
	' End If
 
  wend
	objXmlReader.Close()'close the xml reader
	Set objXmlReader=Nothing
	Set objReader=Nothing

	For intMaxDepth= intDepth to 0 step -1'Loop through the  depth of the XML
	
			For intMaxNodes=intNode to 0 step -1
			
				retRetVal=GetValFromArray(arrayMatrix,intMaxDepth,intMaxNodes)
				If retRetVal<>-1 Then
					intMaxNodes=retRetVal
				End If
			
			Next
	
	Next
Call RemoveDuplicatesInAppMap1()
'Call RemoveDuplicatesInAppMap1()
DataTable.ExportSheet strAppMapSheetPath,"KeywordRepository"
'call RemoveDuplicatesInAppMap(strAppMapSheetPath)

End Function
 ' **********************************************************************************************************************************************************************************************************
	' Function Name:		GetDepth
	' Purpose:			 This function is used to Read the QTP's OR XML tree using DotNetFactory object and  return the Maximum depth of the XML
	' Input Parameters:		objXMLReader
	' Return Values:	    intDepth
	' Date Created:	   22-Apr-2009 	
	' Created By:			Mohan Kakarla,Srinivas P
	' Reviewed By:			
	' ************************************************************************************************************ ********************************************************************************************** 
Function GetDepth(objXMLReader)
   intDepth=0

'    while (objXmlReader.Read())' Reads the xml till end of the file
'	
'	If  objXmlReader.NodeType = "XmlDeclaration" or objXmlReader.NodeType = "Element"Then' checks if the tag type is XmlDeclaration or Element
'
'		If objXmlReader.Name="qtpRep:Object" Then' checks If element name 
'          
'			  If Not intDepth >Cint(objXmlReader.Depth) then  'Compare the current depth to the previous max depth 				     						 
'				   intDepth=Cint(objXmlReader.Depth)'if true assign the current depth to variable
'			  end if   					        				       
'	
'	   End If
'	 
'	End If
'   wend
while (objXmlReader.ReadToFollowing("qtpRep:Object"))' Reads the xml till end of the file
	
   ' If  objXmlReader.NodeType = "XmlDeclaration" or objXmlReader.NodeType = "Element"Then' checks if the tag type is XmlDeclaration or Element

	   ' If objXmlReader.Name="qtpRep:Object" Then' checks If element name 
          
			  If Not intDepth >Cint(objXmlReader.Depth) then  'Compare the current depth to the previous max depth 				     						 
				   intDepth=Cint(objXmlReader.Depth)'if true assign the current depth to variable
			  end if   					        				       
	
	  ' End If
	 
   ' End If
   wend
   
	GetDepth=intDepth' return the variable
	objXmlReader.Close()
	Set objXmlReader=Nothing
End Function
 ' **********************************************************************************************************************************************************************************************************
	' Function Name:		GetValFromArray
	' Purpose:			 This function is used to Read  arry that contains the node values from XML and  return the data
	' Input Parameters:		arr_Array,intColum
	' Return Values:	 GetValFromArray   
	' Date Created:	   22-Apr-2009 	
	' Created By:			Mohan Kakarla,Srinivas P
	' Reviewed By:			
	' ************************************************************************************************************ ********************************************************************************************** 
Function GetValFromArray(arr_Array,intRow,intColum)

	 Dim intPrevCol
	 Dim boolNotFound
	 Dim intTempRow
	 Dim intTempCol
	 Dim strArrayData
	 Dim arrORData
	 Dim strChildname
	 Dim intMaxArrBound
	 Dim strParentdata
     intPrevCol=intColum
     boolNotFound=False
	For intTempRow= intRow to 0 step-1' Loop through the  parameter intRow
	   
		For intTempCol= intPrevCol to 0 step -1'Loop through the parameter intColum
			   If arr_Array(intTempRow,intTempCol)<>empty then' check if the array cell data is empty or not
				   strArrayData= strArrayData & arr_Array(intTempRow,intTempCol)' assign the data from array of specifried cell
				  If intTempRow= intRow Then'check if the current row is equal to the parameter intRow
					   GetValFromArray=intTempCol' return varaiable
				   End If
				   intPrevCol=intTempCol	  
		  Exit for
			 End if 
		Next
	Next

strArrayData=vbnewline &strArrayData
	 If not  Instr(Environment("FullData"),strArrayData &vbnewline)>0   Then' check for the duplicate data 
         
		  Environment("FullData")=Environment("FullData")&strArrayData &vbnewline
	   strArrayData=Replace(strArrayData,vbnewline,"")
			If strArrayData<>empty  Then' check for the array data is empty or not
			
			arrORData=split(strArrayData,"';")' split the data with delimi
			strChildname=arrORData(0)'get the child name
			For intMaxArrBound=Ubound(arrORData)-1 to 3 step -1' loop through the max bound of the array
				If  strParentdata=""Then'
					strParentdata=arrORData(intMaxArrBound)&"("& chr(34)& arrORData(intMaxArrBound-1) & chr(34) &")"  	'construct the object which looks like QTP OR object	
				 else		  
				strParentdata=strParentdata & "."& arrORData(intMaxArrBound)&"("& chr(34)& arrORData(intMaxArrBound-1) & chr(34) &")" 'construct the object which looks like QTP OR object   
				End If
			   intMaxArrBound=intMaxArrBound-1
			Next
			strChildObject=arrORData(1)&"("& chr(34)& arrORData(0) & chr(34) &")" '
			
			DataTable.GetSheet("KeywordRepository").SetCurrentRow(Cint(Environment("Dataiteration")))' set the current row to next  row in the datatable
			DataTable("UIName","KeywordRepository")=strChildname' assign the child name
			DataTable("ChildObject","KeywordRepository")= strChildObject'assign the parent object
			DataTable("ParentObject","KeywordRepository")= strParentdata'assign the parent object
			Environment("Dataiteration")=Cint(Environment("Dataiteration"))+1'increment the variable
			
			End If

 
	End If
End Function
 ' **********************************************************************************************************************************************************************************************************
	' Function Name:		RemoveDuplicatesInAppMap
	' Purpose:			 This function is used to Remove the duplicate by assigning anumeric value at the end of the data in AppMap
	' Input Parameters:		strAppMapSheetPath
	' Return Values:	 No   
	' Date Created:	   28-Apr-2009 	
	' Created By:			Mohan Kakarla,Srinivas P
	' Reviewed By:			
	' ************************************************************************************************************ **********************************************************************************************
Function RemoveDuplicatesInAppMap(strAppMapSheetPath)
'	Dim objExcel
'	Dim objWorkBook
'	Dim objSheet
'	Dim objRange,objRange2
'	Dim intMaxRows
'	Dim intDuplicate
'	Dim intMinCell
'	Dim strNextVal
'	Const xlAscending = 1'represents the sorting type 1 for Ascending 2 for Desc
'	Const xlYes = 1
'	Set objExcel = CreateObject("Excel.Application")
'	objExcel.visible=False
'	 Set objWorkBook = objExcel.Workbooks.Open (strAppMapSheetPath)'opens the sheet
'	 objExcel.Sheets("KeywordRepository").select
'	 Set objSheet = objExcel.Sheets("KeywordRepository")' To select particular sheet
'	 
'	Set objRange = objSheet.UsedRange'which select the range of the cells has some data other than blank
'	Set objRange2 = objExcel.Range("A1")' select the column to sort
'	objRange.Sort objRange2, xlAscending, , , , , , xlYes	
'	 intMaxRows= objRange.rows.count
'	objSheet.Range("A1:C1").Interior.ColorIndex =40
'	objSheet.Range("A1:C1").Font.FontStyle = "Bold"   
'	With objSheet.Range("A1:A"&intMaxRows) ' select the used range in particular sheet
'			
'			For  intMinCell= 2  to objRange.Rows.count
'
'				intDuplicate=1
'			    strValToFind=objSheet.Cells(intMinCell,1).value
'				Set strNextVal = .Find (strValToFind)' data to find  
'					For each strNextVal in objSheet.Range("A"&intMinCell+1&":A"&intMaxRows)' Loop through the used range
'						    If IsNumeric(strNextVal.Value) Then
'								strNextVal.Value="'"&strNextVal.Value
'							End If
'							 If Lcase(strNextVal.Value)=Lcase(strValToFind) then' compare with the expected data
'							     strNextVal.Value = strNextVal.Value&"_"&intDuplicate' make the gary color if it finds the data
'								 intMinCell=Replace(Replace(strNextVal.Address,"$",""),"A","") 
'							     intDuplicate=intDuplicate+1
'								 Set strNextVal = .FindNext(strNextVal)' next search
'								 else
'								 Exit For
'							End If
'								
'						   
'					next
'			
'			Next
'	End With
'	objWorkBook.save
'	'objExcel.Visible=true
'	objWorkBook.close
'	objExcel.quit
'	set objExcel=nothing

End Function

Function RemoveDuplicatesInAppMap1()
	
			Dim intMaxRows
			Dim intDuplicate
			Dim intMinCell
			Dim strNextVal
		
			rc=DataTable.GetSheet("KeywordRepository").getRowcount

            DataTable.GetSheet("KeywordRepository").AddParameter "Formula"," "
			DataTable.GetSheet("KeywordRepository").AddParameter "DuplicateComparisonDone"," "

			For  intMinCell= 1  to rc-1
                DataTable.GetSheet("KeywordRepository").SetCurrentRow(intMinCell)
				If Trim(DataTable("DuplicateComparisonDone","KeywordRepository")) = ""  Then
						DataTable("DuplicateComparisonDone","KeywordRepository") = "Yes"
                         intDuplicate=1
						strValToFind=DataTable("UIName","KeywordRepository")
						DataRow=CSTR(DataTable.GetSheet("KeywordRepository").GetCurrentRow)
						DataTable.Value("Formula","KeywordRepository")="=MATCH("&CHR(34)&strValToFind&Chr(34)&",A"&DataRow+1&":A"&rc&",0)"
						CurrentRow=DataRow+1
						While DataTable("Formula","KeywordRepository")<>"#N/A"
                            	If IsNumeric(strValToFind) Then
										DataTable("UIName","KeywordRepository")="'"&DataTable("UIName","KeywordRepository")
								End If
								 'If Not DataRow =DataTable("Formula","KeywordRepository") Then
							    Row=DataTable("Formula","KeywordRepository")
                                DataTable.GetSheet("KeywordRepository").SetCurrentRow(Row+CurrentRow-1)
								If DataTable("DuplicateComparisonDone","KeywordRepository") <> "Yes"  Then
									DataTable("UIName","KeywordRepository")=DataTable("UIName","KeywordRepository")&"_"& intDuplicate
								Else
									DataTable("UIName","KeywordRepository")=DataTable("UIName","KeywordRepository")&"_!~!"
									intDuplicate = intDuplicate -1
                                End If
								intDuplicate=intDuplicate+1
                                CurrentRow=Row+1
                                DataTable("Formula","KeywordRepository")="=MATCH("&CHR(34)&strValToFind&Chr(34)&",A"&CurrentRow&":A"&rc&",0)"
								DataTable("DuplicateComparisonDone","KeywordRepository") = "Yes"
				   Wend
				End If
    		Next

		For  intMinCell= 1  to rc-1	
				DataTable.GetSheet("KeywordRepository").SetCurrentRow(intMinCell)
				If  Right(DataTable("UIName","KeywordRepository"),4) = "_!~!" Then
						DataTable("UIName","KeywordRepository") = Left(DataTable("UIName","KeywordRepository"),len(DataTable("UIName","KeywordRepository"))-4)
				End If
		Next
  
      DataTable.GetSheet("KeywordRepository").DeleteParameter("Formula")
	  DataTable.GetSheet("KeywordRepository").DeleteParameter("DuplicateComparisonDone")
End Function

Set objForm = DotNetFactory.CreateInstance("System.Windows.Forms.Form","System.Windows.Forms")
Set objBtn1 = DotNetFactory.CreateInstance("System.Windows.Forms.Button","System.Windows.Forms")

Set objEdit3 = DotNetFactory.CreateInstance("System.Windows.Forms.TextBox","System.Windows.Forms")
Set objEdit4 = DotNetFactory.CreateInstance("System.Windows.Forms.TextBox","System.Windows.Forms")
Set objRd1 = DotNetFactory.CreateInstance("System.Windows.Forms.RadioButton", "System.Windows.Forms")
Set objRd2 = DotNetFactory.CreateInstance("System.Windows.Forms.RadioButton", "System.Windows.Forms")
x=30
y=30
width=50
height=20
Set p1 = DotNetFactory.CreateInstance("System.Drawing.Point","System.Drawing",x,y) 'This will provide the locations(X,Y) for the controls
Set s1 = DotNetFactory.CreateInstance("System.Drawing.Size","System.Drawing",width,height) 'T

Set lbl3 =DotNetFactory.CreateInstance("System.Windows.Forms.Label","System.Windows.Forms")
Set lbl4= DotNetFactory.CreateInstance("System.Windows.Forms.Label","System.Windows.Forms")
Set lbl5= DotNetFactory.CreateInstance("System.Windows.Forms.Label","System.Windows.Forms")
Set boarder = DotNetFactory.CreateInstance("System.Windows.Forms.FormStartPosition","System.Windows.Forms")

oldX=CInt(p1.X)
lbl3.Text="AppMap Path"
p1.X=oldX
p1.Y=CInt(p1.Y)+30
lbl3.Location=p1
s1.width=100
s1.height=30
lbl3.size=s1
objForm.Controls.Add(lbl3)

objEdit3.Location=p1
s1.Width=500
objEdit3.size=s1
objForm.Controls.Add(objEdit3)
lbl4.Text="ObjectRepository Path(*.tsr)"
p1.X=oldX
p1.Y=CInt(p1.Y)+30
lbl4.Location=p1
s1.width=100
s1.height=30
lbl4.size=s1
objForm.Controls.Add(lbl4)
p1.X=Cint(p1.X)+CInt(lbl4.Width)
objEdit4.Location=p1
s1.width=400
objEdit4.size=s1
objForm.Controls.Add(objEdit4)
objRd1.Text="Replace Mode"
objRd2.Text="Append Mode"
objRd1.Checked =True
p1.X=oldX
s1.Width=300
s1.height=20
lbl5.size=s1
lbl5.Text="Select Radio Buttons to Choose The Mode"
p1.Y=CInt(p1.Y)+30
lbl5.Location=p1
objForm.Controls.Add(lbl5)
p1.Y=CInt(p1.Y)+30
p1.X=CInt(p1.X)+100
objRd1.size=s1
objRd1.Location=p1
p1.Y=Cint(p1.Y)+20
objRd2.size=s1
objRd2.Location=p1
p1.Y=Cint(p1.Y)+20
objForm.Controls.Add(objRd1)
objForm.Controls.Add(objRd2)
'objForm.Controls.Add(fd)
objBtn1.Text="OK"
p1.X=CInt(p1.X)+100
p1.Y=Cint(p1.Y)+40
objBtn1.Location=p1
objForm.CancelButton=objBtn1
objForm.Controls.Add(objBtn1)
objForm.StartPosition =boarder.CenterScreen
objForm.TopMost = True 
objForm.Text="Q-UIK Framework AutoFill Setup"
s1.Width=550
s1.Height=300
objForm.size=s1
objEdit3.Text=DataTable("AppMapPath")
objEdit4.Text=DataTable("ObjectRepositoryPath")
objForm.ShowDialog
strAppMapSheetPath=objEdit3.Text
strORpath=objEdit4.Text
strXMLpath=Replace(strORpath,".tsr",".xml")


If objRd1.Checked Then'Change Test Plan  
  AppendMode=False 
else
   AppendMode=True
End If
'****************For Open file dialog


Set objForm = Nothing
Set objBtn1 = Nothing

Set objEdit3 = Nothing
Set objEdit4 = Nothing
Set objRd1 = Nothing
Set objRd2 =Nothing

Set p1 = Nothing
Set s1 =Nothing

Set lbl3 =Nothing
Set lbl4= Nothing
Set lbl5= Nothing
Set boarder = Nothing
Set fd=Nothing

Set objFSO=CreateObject("Scripting.FileSystemObject")
If objFSO.FileExists(strXMLpath) then
  objFSO.DeleteFile(strXMLpath)
End If

Set objFSO=Nothing
Dim objRepository
Set objRepository=CreateObject("Mercury.ObjectRepositoryUtil")



Function ExportOR(strORpath,strXMLpath)
    'objRepository.Load(strORpath )
   objRepository.ExportToXML strORpath,strXMLpath
End Function



Call ExportOR(strORpath,strXMLpath)
set objRepository=Nothing

Call ReadXMLOR(strAppMapSheetPath,strXMLpath,AppendMode)






























