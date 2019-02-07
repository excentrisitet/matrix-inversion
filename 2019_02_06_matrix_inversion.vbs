'Calculating of inverse matrix

Option Explicit

Dim objFSO

Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim objFolder
Set objFolder = objFSO.GetFolder(".")
Dim objTextWriter, objTextReader

Dim Dictionary_of_lines										'As New Dictionary of lines
Set Dictionary_of_lines = CreateObject("Scripting.Dictionary")
Set objTextWriter = objFSO.OpenTextFile("inverse.txt", 2, True)					'File for writing
Set objTextReader = objFSO.OpenTextFile("Original.txt", 1, False)				'File for reading

'Filling up the dictionary
Do
	With Dictionary_of_lines
		.Add .Count, objTextReader.ReadLine()
	End With
Loop Until objTextReader.AtEndOfStream

Dim i,j,k
ReDim arr_original(Ubound(Dictionary_of_lines.Items()), Ubound(Dictionary_of_lines.Items()))
'Filling up the array
For i = 0 to Ubound(Dictionary_of_lines.Items())
	Dim Temp_array
	Temp_array = split(Dictionary_of_lines.Items()(i))
	For j = 0 to Ubound(Dictionary_of_lines.Items())
		arr_original(i,j) = Temp_array(j)
	Next
Next

'Creating a template for the inverse matrix
ReDim arr_inv (UBound(arr_original , 1) , UBound(arr_original , 2))
For i = 0 to UBound(arr_original , 1)
	arr_inv(i,i) = 1
Next

'Direct way
For i = 0 to UBound(arr_original , 1)
	Dim temp
	If arr_original(i,i) = 0 Then		'If original element is zero
		Dim temp2 : temp2 = i + 1 
		Do
			temp2 = temp2 + 1
		Loop Until arr_original(temp2,i) <> 0 Or temp2 = UBound(arr_original , 1)
		If temp2 = UBound(arr_original , 1) Then 
			MsgBox "У матрицы нулевой определитель. Обратной матрицы не существует"
			WScript.Quit
		Else
			For j = 0 to UBound(arr_original , 2)
				temp = arr_original(i,j)
				arr_original(i,j) = arr_original(temp2,j)
				arr_original(temp2,j) = temp
				temp = arr_inv(i,j)
				arr_inv(i,j) = arr_inv(temp2,j)
				arr_inv(temp2,j) = temp
			Next
		End If
	End If
	temp = arr_original(i,i)
	For j = 0 to UBound(arr_original , 2)
		arr_original(i,j) = arr_original(i,j) / temp
		arr_inv(i,j) = arr_inv(i,j) / temp
	Next
	For k = i + 1 to UBound(arr_original , 1)
		temp = arr_original(k,i)
		For j = 0 to UBound(arr_original , 2)
			arr_original(k,j) = arr_original(k,j) - arr_original(i,j) * temp
			arr_inv(k,j) = arr_inv(k,j) - arr_inv(i,j) * temp
		Next
	Next
Next
'Backward way
For i = UBound(arr_original , 1) to 1 step -1
	For k = i - 1 to 0 step -1
		temp = arr_original(k,i)
		For j = UBound(arr_original , 2) to 0 step -1
			arr_original(k,j) = arr_original(k,j) - arr_original(i,j) * temp
			arr_inv(k,j) = arr_inv(k,j) - arr_inv(i,j) * temp
		Next
	Next
	temp = arr_original(i-1,i-1)
	For j = UBound(arr_original , 2) to UBound(arr_original , 2) step -1
		arr_original(i,j) = arr_original(i,j) / temp
		arr_inv(i,j) = arr_inv(i,j) / temp
	Next
Next

'Printing the results
For i = 0 to Ubound(Dictionary_of_lines.Items())
	For j = 0 to Ubound(Dictionary_of_lines.Items())
		objTextWriter.Write(arr_inv(i,j))
		objTextWriter.Write(" ")
	Next
	objTextWriter.Writeline()
Next
objTextWriter.Writeline()


Set objFolder = Nothing
