Attribute VB_Name = "CommandeShell"

' Cette fonction permet d_executer une commande cmd
Sub ExecutionCommandeCMD(ByVal commande As String)
    'Cette constante contient le chemin du cmd de windows
    Const pathCmd = "C:\\Windows\\System32\\cmd.exe"
    Shell ("cmd /c " & commande)
End Sub

' Cette fonction permet d_executer une commande Shell Git
Sub ExecutionCommandeGit(ByVal commande As String)
    ' Cette constante est le chemin du shell de git
    Dim pathShellGit As String
    pathShellGit = documentActive.ProjectObject.monProjet.pathFolderprojects & "Setup\PortableGit\gitCmd.exe"
    commande = commande & sepCmd & "exit"
    Call Shell("""" & pathShellGit & """ """ & commande & """", vbHide)
End Sub


