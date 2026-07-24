' ===================================================================
'  Convertisseur_CAO.vbs
'  Lance l'interface graphique SANS aucune fenêtre de commande.
'  >>> Double-cliquez sur CE fichier pour ouvrir l'application. <<<
'
'  Les fichiers de fonctionnement sont dans le sous-dossier « Application ».
'  Utilise « pyw » (Python fenêtre, sans console). En cas de souci, lancez
'  Application\Convertisseur_CAO.bat qui affiche les erreurs éventuelles.
' ===================================================================

Option Explicit
Dim sh, fso, scriptDir, appDir, target
Set sh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
appDir = fso.BuildPath(scriptDir, "Application")
target = fso.BuildPath(appDir, "convertisseur_gui.py")
sh.CurrentDirectory = appDir

' pyw = lanceur Python "fenêtre" (pas de console). 0 = fenêtre cachée.
On Error Resume Next
sh.Run "pyw """ & target & """", 0, False
If Err.Number <> 0 Then
    Err.Clear
    sh.Run "pythonw """ & target & """", 0, False
End If
On Error Goto 0
