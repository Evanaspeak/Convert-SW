' ===================================================================
'  Convertisseur_CAO.vbs
'  Lance l'application SANS aucune fenêtre de commande.
'  >>> Double-cliquez pour ouvrir l'appli. <<<
'  Vous pouvez aussi GLISSER-DÉPOSER des fichiers CAO sur ce .vbs :
'  la fenêtre de conversion s'ouvre déjà préchargée avec ces fichiers.
'
'  Les fichiers de fonctionnement sont dans le sous-dossier « DATA ».
'  En cas de souci, lancez DATA\Convertisseur_CAO.bat (affiche les erreurs).
' ===================================================================

Option Explicit
Dim sh, fso, scriptDir, appDir, target, i, argstr
Set sh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
appDir = fso.BuildPath(scriptDir, "DATA")
target = fso.BuildPath(appDir, "convertisseur_cao.py")
sh.CurrentDirectory = appDir

' fichiers éventuellement glissés-déposés sur l'icône
argstr = ""
For i = 0 To WScript.Arguments.Count - 1
    argstr = argstr & " """ & WScript.Arguments(i) & """"
Next

' pyw = lanceur Python "fenêtre" (pas de console). 0 = fenêtre cachée.
On Error Resume Next
sh.Run "pyw """ & target & """" & argstr, 0, False
If Err.Number <> 0 Then
    Err.Clear
    sh.Run "pythonw """ & target & """" & argstr, 0, False
End If
On Error Goto 0
