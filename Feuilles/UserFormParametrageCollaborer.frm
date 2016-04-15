VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormParametrageCollaborer 
   Caption         =   "Paramétrage"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10275
   OleObjectBlob   =   "UserFormParametrageCollaborer.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormParametrageCollaborer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Ce tableau stocke l_ensemble des utilisateurs
    Dim listeUser() As New UsersObject
    
' Cette variable stocke l_utilisateur choisi
    Dim userChoisi As New UsersObject
    
' Ce checkbox permet d_activer ou de desactiver les options de parametrage de depots
Private Sub CheckBoxParametrageDepot_Click()

    If (CheckBoxParametrageDepot) Then
        FrameAjoutDepootDistantUser.Visible = True
    Else
        FrameAjoutDepootDistantUser.Visible = False
    End If
End Sub

' Procesdure d_annulation
Private Sub annulation()
    
'    If Not (listeUser Is Nothing) Then
'        Set listeUser = Nothing
 '   End If
    If Not (userChoisi Is Nothing) Then
        Set userChoisi = Nothing
    End If
' et on le supprime en memoire
    Unload UserFormParametrageCollaborer
End Sub

Private Sub ComboBoxNDNameUser_Change()

    Dim i As Byte
' On parcourt l_ensemble des utilisateurs
    For i = 1 To UBound(listeUser)
' On verifie si l_email choisi correspond a celui d_un des utilisateurs
        If (StrComp(listeUser(i).email, ComboBoxNDNameUser.Value) = 0) Then
' Si oui on recupere l_utilisateur
                Set userChoisi = listeUser(i)
            Exit Sub
        End If
    Next i
End Sub

Private Sub CommandButtonAnnulerND_Click()
' On fait disparaitre le userform
    annulation
End Sub

Private Sub CommandButtonAnnulerUPD_Click()

' On fait disparaitre le userform
    annulation
End Sub

'Pour enregistrer un nouveau dépôt
Private Sub CommandButtonEnregistrerND_Click()

' On cache le userform
        UserFormParametrageCollaborer.Hide
' On verifie si le mot de passe entré est conforme a celui de l_utilisateur choisi
    If (StrComp(TextBoxNDPassword.Value, userChoisi.password) = 0) Then ' Si oui

' On cree un objet RemoteRepositoryObject
        Dim repo As New RemoteRepositoryObject
' On stocke les informations saisies dans l_objet repo
        repo.nomRemote = TextBoxNDNameRepository.Value
        repo.url = TextBoxNDUrlRepoDistant.Value
        Set repo.user = userChoisi
' On fait appel à la méthode pour inserer le repository cree
        repo.InsererUnNouveauRepository
' On fait disparaitre l_utilisateur
        annulation
    
    Else ' Sinon
        TextBoxNDPassword.Value = ""
        TextBoxNDPassword.BackColor = &HC0C0FF 'rouge
' On recharge le userform
        UserFormParametrageCollaborer.Show
    End If
End Sub

'Pour enregistrer un nouveau utilisateur
Private Sub CommandButtonValider_Click()

' On cree un objet usersObject
    Dim user As New UsersObject
' On stocke les informations saisies dans l_objet user
    user.nomUser = TextBoxUNameUtilisateur.Value
    user.email = TextBoxUEmail.Value
    user.password = TextBoxUPassword.Value
' On fait appel à la méthode pour inserer l_utilisateur cree
    user.InsererUnNouveauUtilisateur
    
    If (CheckBoxParametrageDepot) Then
' On cree un objet RemoteRepositoryObject
        Dim repo As New RemoteRepositoryObject
' On stocke les informations saisies dans l_objet repo
        repo.nomRemote = TextBoxUNameRepository.Value
        repo.url = TextBoxUUrl.Value
        Set repo.user = user
' On ajoute ce nouveau repository
        repo.AjouterCeRepositoryDistant
' On fait appel à la méthode pour inserer le repository cree
        repo.InsererUnNouveauRepository
    End If
    
' On fait disparaitre le userform
    annulation
End Sub

' Cette methode se charge au début du chargement du userform
Private Sub UserForm_Initialize()

'*************************************************************************
' On  verifie si la machine client a acces l_internet
    If (IsInternetConnected) Then
        CheckBoxParametrageDepot.Enabled = True
        CommandButtonEnregistrerND.Enabled = True
    End If

'*************************************************************************
' On charge les informations de la prorpiete du document
    If (Not documentActive Is Nothing) Then
        If (Not documentActive.ProjectObject Is Nothing And documentActive.ProjectObject.IsProjetGit) Then
            TextBoxProprieteNomProjet.Value = documentActive.ProjectObject.nomProjet
            If (Not documentActive.ProjectObject.remoteRepository Is Nothing) Then
                TextBoxProprieteNomDepot.Value = documentActive.ProjectObject.remoteRepository.nomRemote
                'TextBoxProprieteUrl.Value = documentActive.ProjectObject.remoteRepository.url
                If (Not documentActive.ProjectObject.remoteRepository.user Is Nothing) Then
                    TextBoxProprieteEmail.Value = documentActive.ProjectObject.remoteRepository.user.email
                    TextBoxProprieteNomUser.Value = documentActive.ProjectObject.remoteRepository.user.nomUser
                End If
            End If
        End If
    End If

'*************************************************************************
' Chargement du fichier XML s_il nest pas en memoire
    ChargementFichierInformationsUsers

'On declare un objet user
    Dim user As Object
    Dim i As Byte
        i = 1
' On recupere tous les noeuds repo à l_interieur du noeud depotDistant
        For Each user In XmlDoc.SelectNodes("/infos/Users/user")
'On redimensionne le tableau a chaque entree
            ReDim Preserve listeUser(i)
            listeUser(i).nomUser = user.SelectSingleNode("nameUser").text
            listeUser(i).email = user.SelectSingleNode("email").text
            listeUser(i).password = user.SelectSingleNode("password").text
' On charge la liste des depots dans le comboBox
            ComboBoxNDNameUser.AddItem (listeUser(i).email)
        i = i + 1
        Next user

'*************************************************************************
' Pour la recuperation de l_url repository
' On  se positionne dans le repository correspondant
    Dim repositoryCorrespondant As String
    repositoryCorrespondant = "/infos/Users/user[@id='" & TextBoxProprieteNomUser.Value & "']/remoteRepositories/repository[@id='" & TextBoxProprieteNomDepot.Value & "']"
    Set repositoryUrl = XmlDoc.SelectSingleNode(repositoryCorrespondant)

' On verifie si li noeud a ete selectionne
    If Not repositoryUrl Is Nothing Then
' Puis on cree un nouveau noeud user
        TextBoxProprieteUrl.Value = repositoryUrl.SelectSingleNode("url").text
    End If
End Sub
