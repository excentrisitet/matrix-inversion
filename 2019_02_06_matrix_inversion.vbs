'Попытки создания мастер про-файла

Option Explicit

Dim objFSO

Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim objFolder
Set objFolder = objFSO.GetFolder(".")
Dim objTextWriter, objTextReader

Dim Dictionary_of_lines														'As New Dictionary of lines
Set Dictionary_of_lines = CreateObject("Scripting.Dictionary")
Set objTextWriter = objFSO.OpenTextFile("inverse.txt", 2, True)				'File for writing
Set objTextReader = objFSO.OpenTextFile("Original.txt", 1, False)				'File for reading

'Filling up the dictionary
Do
	With Dictionary_of_lines
		.Add .Count, objTextReader.ReadLine()
	End With
Loop Until objTextReader.AtEndOfStream

Dim i,j,k
ReDim arr_original(Ubound(Dictionary_of_lines.Items()), Ubound(Dictionary_of_lines.Items()))
'Заполнение массива
For i = 0 to Ubound(Dictionary_of_lines.Items())
	Dim Temp_array
	Temp_array = split(Dictionary_of_lines.Items()(i))
	For j = 0 to Ubound(Dictionary_of_lines.Items())
		arr_original(i,j) = Temp_array(j)
	Next
Next

ReDim arr_inv (UBound(arr_original , 1) , UBound(arr_original , 2))
For i = 0 to UBound(arr_original , 1)
	arr_inv(i, i) = 1
Next

Call Print_results

For i = 0 to UBound(arr_original , 1)
	Dim temp
	temp = arr_original(i, i)
	For j = 0 to UBound(arr_original , 2)
		arr_original(i, j) = arr_original(i, j) / temp
		arr_inv(i, j) = arr_inv(i, j) / temp
	Next
	For k = i + 1 to UBound(arr_original , 1)
		temp = arr_original(k, i)
		For j = 0 to UBound(arr_original , 2)
			arr_original(k ,j) = arr_original(k, j) - arr_original(i, j) * temp
			arr_inv(k ,j) = arr_inv(k ,j) - arr_inv(i ,j) * temp
		Next
	Next
Next

For i = UBound(arr_original , 1) to 0 step -1
	temp = arr_original(i, i)
	For j = UBound(arr_original , 2) to UBound(arr_original , 2) step -1
		arr_original(i, j) = arr_original(i, j) / temp
		arr_inv(i, j) = arr_inv(i, j) / temp
	Next
	For k = i - 1 to 0 step -1
		temp = arr_original(k, i)
		For j = UBound(arr_original , 2) to 0 step -1
			arr_original(k ,j) = arr_original(k, j) - arr_original(i, j) * temp
			arr_inv(k ,j) = arr_inv(k ,j) - arr_inv(i ,j) * temp
		Next
	Next
Next


For i = 0 to Ubound(Dictionary_of_lines.Items())
	For j = 0 to Ubound(Dictionary_of_lines.Items())
		objTextWriter.Write(arr_inv(i,j))
		objTextWriter.Write(" ")
	Next
	objTextWriter.Writeline()
Next
objTextWriter.Writeline()



Set objFolder = Nothing