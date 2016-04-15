VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Folder_ProjectsObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pNomFolderProjects As String
Private pPathFolderprojects As String

' ================ NomFolderProjects =========================
' Getter
Property Get nomFolderProjects() As String
    nomFolderProjects = pNomFolderProjects
End Property

' Setter
Property Let nomFolderProjects(nomFolderProjects As String)
    pNomFolderProjects = nomFolderProjects
End Property

' ================ PathFolderprojects =========================
' Getter
Property Get pathFolderprojects() As String
    pathFolderprojects = pPathFolderprojects
End Property

' Setter
Property Let pathFolderprojects(pathFolderprojects As String)
    pPathFolderprojects = pathFolderprojects
End Property

Private Sub Class_Initialize()
' Chargement du fichier XML s_il nest pas en memoire
    ChargementFichierInformationsUsers
' Selection du noeud path
    Dim path As String
    path = XmlDoc.SelectSingleNode("/infos/config/path").text
    If Not IsEmpty(path) Then
        pPathFolderprojects = path
    Else
        pPathFolderprojects = ""
    End If
    pNomFolderProjects = "Projets"
End Sub

' Creation d_un repertoire nommé Projets où seront stockés tous les projets git
Sub InitProjegitstGit()
    ' Cette variable stocke le chemin du dépôt local git
    Dim chemin As String
    chemin = "%USERPROFILE%\Collaborer_Projets_Git\Projets"
    ' Vérification de l_existence du dépot local
    If (Dir(chemin, vbDirectory) = "") Then
        ' Commande de création du répertoire pour le dépôt local
        commande = "mkdir " & chemin
        ExecutionCommandeCMD (commande)
    End If
End Sub
