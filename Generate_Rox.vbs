'Generate the Final csv file



'''''''''''''''''''''''''''''''''''''''
'Read source CSV file to big array 65536 lines
'''''''''''''''''''''''''''''''''''''''
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile("rox_step2.csv", ForReading)
Const ForReading = 1


Dim Src_arr_FileLines(65535,5)
Dim arrFileLines()
SrcLineNum = 1
Do Until objFile.AtEndOfStream

non_split_strLine = objFile.ReadLine
'Src_arr_FileLines(SrcLineNum,0) = objFile.ReadLine
'Src_arr_FileLines(SrcLineNum,1) = objFile.ReadLine
'Src_arr_FileLines(SrcLineNum,2) = objFile.ReadLine
'non_split_strLine = objFile.ReadLine
MyArray = Split(non_split_strLine, ",", -1, 1) 
Src_arr_FileLines(SrcLineNum,0) = MyArray(0)
Src_arr_FileLines(SrcLineNum,1) = MyArray(1)
Src_arr_FileLines(SrcLineNum,2) = MyArray(2)
Src_arr_FileLines(SrcLineNum,3) = MyArray(3)


Src_arr_FileLines(SrcLineNum,0)=replace(Src_arr_FileLines(SrcLineNum,0),CHR(34),"")
Trim(Src_arr_FileLines(SrcLineNum,0))
Src_arr_FileLines(SrcLineNum,1)=replace(Src_arr_FileLines(SrcLineNum,1),CHR(34),"")
Trim(Src_arr_FileLines(SrcLineNum,1))
Src_arr_FileLines(SrcLineNum,2)=replace(Src_arr_FileLines(SrcLineNum,2),CHR(34),"")
Trim(Src_arr_FileLines(SrcLineNum,2))
Src_arr_FileLines(SrcLineNum,3)=replace(Src_arr_FileLines(SrcLineNum,3),CHR(34),"")
Trim(Src_arr_FileLines(SrcLineNum,3))


'str=replace(str,CHR(34),"&quot;") “ ???

SrcLineNum = SrcLineNum+1
Loop

'WScript.Echo Src_arr_FileLines(11,0)
'WScript.Echo Src_arr_FileLines(11,1)
'WScript.Echo SrcLineNum



'''''''''''''''''''''''''''''''''''''''
'Write to standard format
'''''''''''''''''''''''''''''''''''''''
Dim Stuff, myFSO, WriteStuff
Set myFSO = CreateObject("Scripting.FileSystemObject")
Set WriteStuff = myFSO.OpenTextFile("Final_rox_format.csv", 8, True)

i=1
Do Until i>SrcLineNum

'WriteStuff.WriteLine(Chr(34)+Src_arr_FileLines(i,0)+Chr(13)+Src_arr_FileLines(i,1)+Chr(13)+Src_arr_FileLines(i,2)+Chr(13)+Src_arr_FileLines(i,3)+Chr(34)+","+","+Chr(34)+Src_arr_FileLines(i+1,0)+Chr(13)+Src_arr_FileLines(i+1,1)+Chr(13)+Src_arr_FileLines(i+1,2)+Chr(13)+Src_arr_FileLines(i+1,3)+Chr(34)+","+","+Chr(34)+Src_arr_FileLines(i+2,0)+Chr(13)+Src_arr_FileLines(i+2,1)+Chr(13)+Src_arr_FileLines(i+2,2)+Chr(13)+Src_arr_FileLines(i+2,3)+Chr(34))
WriteStuff.WriteLine(Chr(34)+Src_arr_FileLines(i,0)+Chr(10)+Src_arr_FileLines(i,1)+Chr(10)+Src_arr_FileLines(i,2)+Chr(32)+Src_arr_FileLines(i,3)+Chr(34)+","+","+Chr(34)+Src_arr_FileLines(i+1,0)+Chr(10)+Src_arr_FileLines(i+1,1)+Chr(10)+Src_arr_FileLines(i+1,2)+Chr(32)+Src_arr_FileLines(i+1,3)+Chr(34)+","+","+Chr(34)+Src_arr_FileLines(i+2,0)+Chr(10)+Src_arr_FileLines(i+2,1)+Chr(10)+Src_arr_FileLines(i+2,2)+Chr(32)+Src_arr_FileLines(i+2,3)+Chr(34))


'Chr(10) <p>
'Chr(13) <br>
'Chr(32) <space>
'Chr(34) <">
'Chr(37) <%>
'Chr(38) <&>

i=i+3
Loop



WriteStuff.Close
SET WriteStuff = NOTHING
SET myFSO = NOTHING


