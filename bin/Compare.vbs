Const ForReading = 1, ForWriting = 2
Dim fso, txtFile, txtFile2, strLine1, strLine2, strMatch
Set fso = CreateObject("Scripting.FileSystemObject")
Set txtFile1 = fso.OpenTextFile("..\kigSTEM\192.168.2.4 13-04-2019 Hardware_Information.csv", ForReading)
Set f = fso.OpenTextFile("..\Log\1.txt", ForWriting, True)

Do Until txtFile1.AtEndOfStream
strMatch = False
    strLine1 = txtFile1.Readline
Set txtFile2 = fso.OpenTextFile("..\kigSTEM\192.168.2.4 14-04-2019 Hardware_Information.csv", ForReading)
        Do Until txtFile2.AtEndOfStream
            strLine2 = txtFile2.Readline
                If Trim(UCase(strLine2)) = Trim(UCase(strLine1)) Then
                    strMatch = True
                Else 
                End If 
        Loop
        txtFile2.Close
                If strMatch <> True then
                    f.writeline strLine1
                End If
Loop
f.Close
Wscript.Echo "Done"
