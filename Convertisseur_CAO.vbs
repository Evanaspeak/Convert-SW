' ===================================================================
'  Convertisseur_CAO.vbs
'  Lance l'interface graphique SANS aucune fenêtre de commande.
'  >>> Double-cliquez sur CE fichier pour ouvrir l'application. <<<
'
'  Utilise « pyw » (Python fenêtre, sans console). En cas de souci,
'  utilisez Convertisseur_CAO.bat qui affiche les erreurs éventuelles.
' ===================================================================

Option Explicit
Dim sh, fso, scriptDir, target, cmd
Set sh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' Dossier de ce script
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
sh.CurrentDirectory = scriptDir
target = fso.BuildPath(scriptDir, "convertisseur_gui.py")

' pyw = lanceur Python "fenêtre" (pas de console). 0 = fenêtre cachée.
cmd = "pyw """ & target & """"
On Error Resume Next
sh.Run cmd, 0, False
If Err.Number <> 0 Then
    ' Repli : pythonw
    Err.Clear
    sh.Run "pythonw """ & target & """", 0, False
End If
On Error Goto 0
