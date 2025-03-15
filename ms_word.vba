Sub start_ikprint()
    '    '"python ikprint.py -f " & Chr(34) & ActiveDocument.Name & Chr(34)
    Dim id As Integer
    Dim wsh As Object
    Dim pythonPath, ikPrintPath As String

    pythonPath = "C:\Users\lando\AppData\Local\Programs\Python\Python312\python.exe"
    ikPrintPath = " D:\projects\python\ikprint\ikprint.py -f "

    Set wsh = VBA.CreateObject("WScript.Shell")
    id = wsh.Run(pythonPath & ikPrintPath & Chr(34) & ActiveDocument.Name & Chr(34), 1, True)
End Sub
