VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RemoteRepositoryObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pNomRemote As String
Private pUrl As String
Private pUser As UsersObject

' ================ NOMRemote =========================
' Getter
Property Get nomRemote() As String
    nomRemote = pNomRemote
End Property

' Setter
Property Let nomRemote(nomRemote As String)
    pNomRemote = nomRemote
End Property

' ================ Url =========================
' Getter
Property Get url() As String
    url = pUrl
End Property

' Setter
Property Let url(url As String)
    pUrl = url
End Property

' ================ User =========================
' Getter
Property Get user() As UsersObject
    Set user = pUser
End Property

' Setter
Property Set user(ByRef userRef As UsersObject)
    Set pUser = userRef
End Property

' Initialisation de l_objet
Private Sub Class_Initialize()
    If Not documentActive Is Nothing Then
        If Not documentActive.ProjectObject Is Nothing Then
            If Not documentActive.ProjectObject.monProjet Is Nothing Then
                commande = "cd " & documentActive.ProjectObject.pathProjet
                Dim fichier As String
                fichier = documentActive.ProjectObject.monProjet.pathFolderprojects & "Configuration\result.txt"
                commande = commande & sepCmd & "(git remote & git config user.name & git config user.email ) > " & fichier
                ExecutionCommandeGit (commande)
                If (RechercheFolder(fichier)) Then
                    If (FileLen(fichier) > 0) Then
                        LireUnFichier (fichier)
' On verifie si le depot est lie a un depot distant
                        If (documentActive.ProjectObject.IsLinkLieRemoteRepository) Then
' On recupere le nom du dépôt distant
                            nomRemote = file.ReadLine
                        End If
                        Set user = New UsersObject
' On recupere le nom de l_utilisateur
                        user.nomUser = file.ReadLine
' On recupere l_adresse Email
                        user.email = file.ReadLine
                        FermetureFichier
' Recuperons maintenant l_url du depot
                        'commande = "git config remote." & nomRemote & ".url > " & fichier
                        'ExecutionCommandeGit (commande)
                        'If (FileLen(fichier) > 0) Then
                         '   LireUnFichier (fichier)
                          '  url = file.ReadLine
                           ' FermetureFichier
                       ' End If
                    Else
                        pNomRemote = ""
                    End If
                Else
                    pNomRemote = ""
                End If
            End If
        End If
    End If
End Sub

' Ce code s_executera a la fin du code vba
Private Sub Class_Terminate()
    If Not (user Is Nothing) Then
        Set user = Nothing
    End If
End Sub

' Cette methode permet de creer un nouveau utilisateur
Sub InsererUnNouveauRepository()
    
' On charge d'abord le fichier XML
    ChargementFichierInformationsUsers
    
' On  se positionne dans le noeud users qui contient l'ensemble des utilisateurs
' Dim users As XMLNode
    Dim userCorrespondant As String
    userCorrespondant = "/infos/Users/user[@id='" & user.nomUser & "']/remoteRepositories"
  
    Set RemoteRepositories = XmlDoc.SelectSingleNode(userCorrespondant)

' On verifie si li noeud a ete selectionne
    If Not RemoteRepositories Is Nothing Then
' Puis on cree un nouveau noeud user
        Set repository = XmlDoc.CreateElement("repository")
        repository.setAttribute "id", nomRemote

' On ajoute le nouveau dépôt à la liste des dépots
        RemoteRepositories.appendChild (repository)

' element Nom Repository
        Set ENameRepository = XmlDoc.CreateElement("nameRepository")
        ENameRepository.text = nomRemote
        repository.appendChild (ENameRepository)

' element Url Repository
        Set EUrl = XmlDoc.CreateElement("url")
        EUrl.text = url
        repository.appendChild (EUrl)

' On met à jour le fichier
        SauvegarderFichierInformationsUser
' On ajoute le depot configure dans git
        AjouterCeRepositoryDistant
Else
' Message d_erreur
End If

End Sub

' AjouterUnRepositoryGit
Sub AjouterCeRepositoryDistant()

' On cree un remote nomme par l_utilisateur qui pointera sur le depot
    commande = "git remote add " & nomRemote & " " & url
    ExecutionCommandeGit (commande)
End Sub

' ConfigurerUnRepositoryGit
Sub ConfigurerCeRepositoryDistant()

' On se positionne dans e projet concerne
    commande = "cd " & documentActive.ProjectObject.pathProjet
' On lie un remote repo avec un depot local
    commande = commande & sepCmd & "git config remote." & nomRemote & ".url https://" & user.nomUser & ":" & user.password & "@" & Replace(url, "https://", "")
    MsgBox commande
    ExecutionCommandeGit (commande)
End Sub
