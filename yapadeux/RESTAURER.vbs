Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
ScriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
ScriptPath = ScriptDir & "\src\restaurer.py"

' Lance pythonw.exe en mode caché et non-bloquant.
On Error Resume Next
WshShell.Run "pythonw.exe """ & ScriptPath & """", 0, False
If Err.Number <> 0 Then
    MsgBox "Python n'est pas installé." & vbCrLf & vbCrLf & _
           "Télécharge-le sur https://python.org en cochant " & _
           "'Add Python to PATH' pendant l'installation.", _
           vbExclamation, "yapadeux"
    WshShell.Run "https://python.org", 1, False
End If
On Error Goto 0
