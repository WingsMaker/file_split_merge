Option Explicit

Dim fso
Dim inputBaseFile
Dim outputFile
Dim outputStream
Dim partStream
Dim partNum
Dim partFile

inputBaseFile = "zz.zip"
outputFile = "zz-merged.zip"

Set fso = CreateObject("Scripting.FileSystemObject")

Set outputStream = CreateObject("ADODB.Stream")
outputStream.Type = 1 ' Binary
outputStream.Open

partNum = 1

Do
    partFile = inputBaseFile & ".part" & Right("000" & partNum, 3)

    If Not fso.FileExists(partFile) Then Exit Do

    'WScript.Echo "Merging: " & partFile

    Set partStream = CreateObject("ADODB.Stream")
    partStream.Type = 1
    partStream.Open
    partStream.LoadFromFile partFile

    outputStream.Write partStream.Read

    partStream.Close
    Set partStream = Nothing

    partNum = partNum + 1
Loop

outputStream.SaveToFile outputFile, 2 ' overwrite if exists
outputStream.Close

WScript.Echo "Merge complete."
WScript.Echo "Output file: " & outputFile
WScript.Echo "Parts merged: " & (partNum - 1)
