Option Explicit

Const CHUNK_SIZE = 10485760 ' 10 MB

Dim fso, inputFile, inputStream
Dim partNum, bytesRead, data

inputFile = "C:\Temp\zz.zip"

Set fso = CreateObject("Scripting.FileSystemObject")

If Not fso.FileExists(inputFile) Then
    WScript.Echo "File not found: " & inputFile
    WScript.Quit
End If

Set inputStream = CreateObject("ADODB.Stream")
inputStream.Type = 1 ' Binary
inputStream.Open
inputStream.LoadFromFile inputFile

partNum = 1

Do Until inputStream.EOS

    Set data = CreateObject("ADODB.Stream")
    data.Type = 1
    data.Open

    data.Write inputStream.Read(CHUNK_SIZE)

    data.SaveToFile inputFile & ".part" & Right("000" & partNum, 3), 2
    data.Close

    partNum = partNum + 1
Loop

inputStream.Close

WScript.Echo "Split completed. Created " & (partNum - 1) & " files."
