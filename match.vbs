'option Explicit
Dim excelpath, excelpath2,array(27), array_2(27)
Dim objShell,objFile, objworkbook,ObjExcel,ObjExcel2,worksheetCount,worksheetCount2
Dim x,counter,usedRange,objworkbook2, value
excelpath ="C:\Users\pritam chavhan\Desktop\vbscript\challenge Excel Matching\file1.xlsx"
excelpath2 ="C:\Users\pritam chavhan\Desktop\vbscript\challenge Excel Matching\file2.xlsx"
'Wscript.echo "Reading Excelsheet From the given path"&Vbcrlf &excelpath &Vbcrlf &excelpath2 
set objExcel=CreateObject("Excel.Application")
set objExcel2=CreateObject("Excel.Application")
set objExcel3=CreateObject("Excel.Application")
set objworkbook = objExcel.Workbooks.open(excelpath)
set objworkbook2 = objExcel2.Workbooks.open(excelpath2)
set objDriverSheet= objworkbook.worksheets("Sheet1")
set objDriverSheet2=objworkbook2.worksheets("Sheet1")
columncount=objDriverSheet.usedRange.columns.count
'MsgBox columncount
columncount2 = objDriverSheet2.usedRange.columns.count
'MsgBox columncount2 
'objExcel.Visible=true
if columncount=columncount2 then	
	Wscript.echo"the number of columns in workbook 1 and 2 are same lets replace them with underscore"
	set regexp = new regexp
	regexp.Global=True
	regexp.Pattern="[^A-Za-z0-9]"
	for i=1 to columncount
		array(i-1) = objDriverSheet.Cells(1,i).value  
		array(i-1)= regexp.Replace(array(i-1), "_")	
	Next
	for i=1 to columncount2
		array_2(i-1) = objDriverSheet2.Cells(1,i).value  
		array_2(i-1)= regexp.Replace(array_2(i-1), "_")
	Next
	counter=0
	for i=0 to columncount
		if (array(i)= array_2(i)) then
		counter=counter+1
		End If
	Next
	'MsgBox counter
	If counter=columncount then 
		objworkbook.close
	End If
	newpath="C:\Users\pritam chavhan\Desktop\vbscript\challenge Excel Matching\file3.xlsx"
	set objworkbook = objExcel3.workbooks.add()
	for i=0 to counter
		objExcel3.cells(1,i+1).value=array(i)
		'MsgBox array(i) 
	Next
	objworkbook.saveas newpath
	objworkbook.close
	newpath="C:\Users\pritam chavhan\Desktop\vbscript\challenge Excel Matching\file4.xlsx"
	set objworkbook = objExcel3.workbooks.add()
	for i=0 to counter
		objExcel3.cells(1,i+1).value=array_2(i)
		'MsgBox array_2(i) 
	Next
	for value=0 to UBound(array) 
		for value2=0 to UBound(array_2) 
			if array(value)=array_2(value2) then
			Wscript.Echo array(value)
			End if 
		Next
	Next
	objworkbook.saveas newpath
	objworkbook.close
Else
	Wscript.echo "both the files are different" 
End if
objExcel.quit
objExcel2.quit
ObjExcel3.quit