
'''''''''''''''''''''''''''''''''''''''
'Read source CSV file to big array 65536 lines
'''''''''''''''''''''''''''''''''''''''
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile("rox_step0.txt", ForReading)
Const ForReading = 1

Dim Stuff, myFSO, WriteStuff
Set myFSO = CreateObject("Scripting.FileSystemObject")
Set WriteStuff = myFSO.OpenTextFile("rox_step1.csv", 8, True)


Dim Src_arr_FileLines(65535,5)
Dim arrFileLines()
SrcLineNum = 1
Do Until objFile.AtEndOfStream

'non_split_strLine = objFile.ReadLine
Src_arr_FileLines(SrcLineNum,0) = objFile.ReadLine
Src_arr_FileLines(SrcLineNum,1) = objFile.ReadLine
Src_arr_FileLines(SrcLineNum,2) = objFile.ReadLine
'non_split_strLine = objFile.ReadLine
'MyArray = Split(non_split_strLine, ",", -1, 1) 
'Src_arr_FileLines(SrcLineNum,2) = MyArray(0)
'Src_arr_FileLines(SrcLineNum,3) = MyArray(1)


WriteStuff.WriteLine(Src_arr_FileLines(SrcLineNum,0)+","+Src_arr_FileLines(SrcLineNum,1)+","+Src_arr_FileLines(SrcLineNum,2)+","+Src_arr_FileLines(SrcLineNum,3))

SrcLineNum = SrcLineNum+1
Loop
objFile.Close

'WScript.Echo Src_arr_FileLines(11,0)
'WScript.Echo Src_arr_FileLines(11,1)
'WScript.Echo SrcLineNum




WriteStuff.Close
SET WriteStuff = NOTHING
SET myFSO = NOTHING