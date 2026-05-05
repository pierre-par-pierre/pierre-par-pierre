Set fso = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject("WScript.Shell")
Desktop = WshShell.SpecialFolders("Desktop")
SentinelPath = Desktop & "\_yapadeux-stop.signal"

Set f = fso.CreateTextFile(SentinelPath, True)
f.WriteLine "stop demande"
f.Close

MsgBox "Demande d'arrêt envoyée." & vbCrLf & vbCrLf & _
       "Le tri en cours s'arrêtera au prochain groupe traité." & vbCrLf & _
       "Tu pourras restaurer les fichiers déjà déplacés en " & _
       "double-cliquant sur RESTAURER.vbs.", _
       vbInformation, "yapadeux"
