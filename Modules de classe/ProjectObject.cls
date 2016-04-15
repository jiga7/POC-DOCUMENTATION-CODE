VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProjectObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pNomProjet As String
Private pPathProjet As String
Private pMonProjet As Folder_ProjectsObject
Private pRemoteRepository As RemoteRepositoryObject

' ================ NomProjet =========================
' Getter
Property Get nomProjet() As String
        nomProjet = pNomProjet
End Property

' Setter
Property Let nomProjet(nomProjet As String)
        pNomProjet = nomProjet
End Property

' ================ NoPathProjet =========================
' Getter
Property Get pathProjet() As String
        pathProjet = pPathProjet
End Property

' ================ MonProjet =========================
' Getter
Property Get monProjet() As Folder_ProjectsObject
        Set monProjet = pMonProjet
End Property

' Setter
Property Set monProjet(ByRef monProjetRef As Folder_ProjectsObject)
        Set pMonProjet = monProjetRef
End Property

' ================ RemoteRepository =========================
' Getter
Property Get remoteRepository() As RemoteRepositoryObject
        Set remoteRepository = pRemoteRepository
End Property

' Setter
Property Set remoteRepository(ByRef remoteRepositoryRef As RemoteRepositoryObject)
        Set pRemoteRepository = remoteRepositoryRef
End Property

' Initialisation de l_objet
Sub Class_Initialize()

    pPathProjet = ActiveDocument.path
    Set pMonProjet = New Folder_ProjectsObject
    Set pRemoteRepository = New RemoteRepositoryObject
    Dim docInverse As String
    docInverse = StrReverse(pathProjet)
    pNomProjet = StrReverse(Left(docInverse, InStr(docInverse, "\") - 1))
End Sub

' le destructeur de la classe pour liberer la memoire
Private Sub Class_Terminate()
    If Not (monProjet Is Nothing) Then
        Set monProjet = Nothing
    End If
    If Not (remoteRepository Is Nothing) Then
        Set remoteRepository = Nothing
    End If
End Sub

' On cree un constructeur avec arguments pour l_objet projet
Public Sub Initiate_Attributes(nomProjetArg As String)
    
    pNomProjet = nomProjetArg
    pPathProjet = monProjet.pathFolderprojects & "Projets\" & pNomProjet
End Sub

' Envoyer son travail dans le repository distant
Sub TransfererSonTravail()

    commande = "cd " & pathProjet & sepCmd & "git push " & remoteRepository.nomRemote & " master"
    ExecutionCommandeGit (commande)
End Sub

' Recuperer l_etat d_un projet depuis un repository distant
Sub TelechargerUnTravail()

  '      MsgBox "Ce projet n'est pas liè à un dépôt distant", vbOKOnly + vbCritical, "Téléchargement Non Effectué"

    commande = "cd " & pathProjet & sepCmd & "git fetch " & remoteRepository.nomRemote
    commande = commande & sepCmd & "git pull " & remoteRepository.nomRemote & " master"
    ExecutionCommandeGit (commande)
End Sub

' Cette fonction permet de vérifier si un projet est lié à un dépôt distant
Function IsLinkLieRemoteRepository() As Boolean
    
    Dim fichier As String
    fichier = monProjet.pathFolderprojects & "Configuration\local_distant.txt"
' Cette commande verifie i l_utilisateur tente de recuperer la configuration au repository distant
    commande = "cd " & pathProjet & sepCmd & "git remote -v > " & fichier
    ExecutionCommandeGit (commande)
    
' On verifie la taille du fichier généré
    If (RechercheFolder(fichier)) Then
        If (FileLen(fichier) > 0) Then
            IsLinkLieRemoteRepository = True ' C_est vrai
        Else
            IsLinkLieRemoteRepository = False ' n_est pas lie
        End If
    Else
        IsLinkLieRemoteRepository = False ' n_est pas lie
    End If
End Function

' Configurer l_identite de l_utilisateur dans git
Sub ConfigurerIdentite()
    
    commande = "cd " & pathProjet
    commande = commande & sepCmd & "git config user.name """ & remoteRepository.user.nomUser & """"
    commande = commande & sepCmd & "git config user.email " & remoteRepository.user.email
    ExecutionCommandeGit (commande)
End Sub

' Cette fonction permet de verifier si on est bien dans un projet git
Function IsProjetGit() As Boolean
    Dim path As String
    path = pathProjet & "\.git"
    If (Dir(path, vbDirectory + vbHidden) = "") Then
        IsProjetGit = False
    Else
        IsProjetGit = True
    End If
End Function

' nous creeons un projet dans le depot git
Function Creation_DUn_ProjetGit() As Boolean
    If (IsProjetGit) Then
        Creation_DUn_ProjetGit = True ' Nom de projet existe deja
    Else
        commande = "mkdir " & pathProjet & sepCmd & "pushd " & pathProjet & sepCmd & "git init"
        ExecutionCommandeGit (commande)
        ConfigurerIdentite
        Documents.Add Template:="H:\Master\M2\Semestre 10\Stage Atos\Application\POC_Collaborer\Collaborer.dotm", NewTemplate:=False, DocumentType:=0
        ChangeFileOpenDirectory pathProjet
        Creation_DUn_ProjetGit = False
    End If
End Function

' la suppresion d_un projet
Sub SupprimerCeProjet()
    'nomProjet = Left(projetASuppInverse, InStr(projetASuppInverse, "\") - 1)
    commande = "cd " & pathProjet
    commande = commande & sepCmd & "rm -rf " & nomProjet & sepCmd & "del /Q * & cd.. & rmdir " & nomProjet
    MsgBox commande
    ActiveDocument.Close
    ExecutionCommandeGit (commande)
    ExecutionCommandeCMD ("pushd " & pathProjet & sepCmd & "del /Q * & cd.. & rmdir " & nomProjet)
End Sub

' Cette fonction ci-dessous nous permet de choisir un projet Git
Function ChoisirProjet() As String

    Dim resultat As Integer
' Donne la possibilité a choisir un projet
Choisir:
' On se positionne dans le répertoire des projets
'ChangeFileOpenDirectory getPathDepot & "Projets"
    Application.FileDialog(msoFileDialogFolderPicker).Title = "Choisissez un projet"
    Application.FileDialog(msoFileDialogFolderPicker).InitialFileName = ""
    resultat = Application.FileDialog(msoFileDialogFolderPicker).Show
' Verifie s_il a bien fait un choix
    If (resultat <> 0) Then
       Dim Projet As String
' Verifie si le projet qu'il a choisi est bien un projet Git
       Projet = Application.FileDialog(msoFileDialogFolderPicker).SelectedItems(1)
       Projet = Projet & "\.git"
       If (Dir(Projet, vbDirectory + vbHidden) <> "") Then
            ChoisirProjet = Application.FileDialog(msoFileDialogFolderPicker).SelectedItems(1)
        Else
            MsgBox "Choisissez un des projets git ou Créez un autre projet.", vbInformation + vbOKOnly, "Choisir un Projet "
            GoTo Choisir
       End If
    Else
        ChoisirProjet = ""
    End If
End Function
