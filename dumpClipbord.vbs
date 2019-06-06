Dim fso
Dim outFileName
outFileName = Wscript.Arguments(0)

Dim obj
Set obj = CreateObject("htmlfile")
Set fso = CreateObject("Scripting.FileSystemObject")

Dim current
Dim bef
bef = ""
Do while True
	current = obj.parentWindow.clipboardData.getData("text")
	If current <> bef Then
		Call outFIle(current, outFileName)
		bef = current
	End If
	Wscript.Sleep 1000
Loop

Private Sub outFile(ByVal buf, ByVal outFileName)
	Dim wFile
	Set wFile = fso.OpenTextFile(outFileName, 8, True)
	wFile.WriteLine(buf)
	wFile.Close
End Sub

