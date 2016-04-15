VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DocumentObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pNomDocument As String 'Le nom du document
Private pProjectObject As ProjectObject

' ================ NomDocument =========================
' Getter
Property Get nomDocument() As String
    nomDocument = pNomDocument
End Property

' Setter
Property Let nomDocument(nomDocument As String)
    pNomDocument = nomDocument
End Property

' ================ ProjectObject =========================
' Getter
Property Get ProjectObject() As ProjectObject
       Set ProjectObject = pProjectObject
End Property

' Setter
Property Set ProjectObject(ByRef projectObjectRef As ProjectObject)
        Set pProjectObject = projectObjectRef
End Property

' Initialisation de l_objet
Private Sub Class_Initialize()
' On recupere le nom document active
    pNomDocument = ActiveDocument.Name
    Set pProjectObject = New ProjectObject
End Sub

' Le destructeur permte de liberer la memoire
Private Sub Class_Terminate()
'    If Not (ProjectObject Is Nothing) Then
'        Set ProjectObject = Nothing
'    End If
End Sub

' Cette procedure ouvre un document du repository Git
Sub Ouvrir_Document()

    ChangeFileOpenDirectory ProjectObject.pathProjet
    Dialogs(wdDialogFileOpen).Show
End Sub

Sub Nouveau_Document()
        Documents.Add Template:="H:\Master\M2\Semestre 10\Stage Atos\Application\POC_Collaborer\Collaborer.dotm", NewTemplate:=False, DocumentType:=0
End Sub

' Cette procedure gere lajout dun nouveau document dans un projet Git
Sub Ajouter_le_Document()
    
    Dim doc As String
    'ChangeFileOpenDirectory getPathDepot & "Projets"
    doc = SauvegardeDansDepot
    If (doc = "") Then
        'MsgBox "Le document n'a pas pu être ajouté dans le dépôt.", vbCritical + vbOKOnly, "Echec de l'ajout"
    Else
        commande = "git add " & doc
        ExecutionCommandeGit (commande)
    End If
End Sub

' Valider les modifications d'un document : commit
Sub ValiderUnDocument(ByVal message As String)

' Cette constante contient la valeur qui sera renvoye si il n'y a pas de document a commiter
    Const NoDocToCommit = "no changes added to commit"

' On sauvegarde le document
    SauvegardeWord
' la variable fichier contient le fichier ou sera envoye le resultat
    Dim fichier As String
    fichier = ProjectObject.monProjet.pathFolderprojects & "Configuration\commit.txt"
'On se positionne sur le dudit document
    commande = "cd " & ProjectObject.pathProjet & sepCmd & "git add " & nomDocument & " > " & fichier
' On lance la commande avec un message sur le commit.
    commande = commande & " & git commit -m """ & message & """"
    ExecutionCommandeGit (commande)
End Sub

' Supprimer un document git
Sub SupprimerUnDocument()

' Commande qui sert à supprimer le document dans le dépot
    commande = "cd " & ProjectObject.pathProjet
    commande = commande & sepCmd & "rm " & nomDocument & " > H:\test.txt"
' On ferme le document actif
    ActiveDocument.Close
' Suppression du document git
    ExecutionCommandeGit (commande)
End Sub

' Cette procedure déclenche la sauvegarde simple de Ms-Word
Sub SauvegardeWord()
    
' On veriifie si le document n a pas ete sauvegarde avant de lancer la sauvegarde
    If (Not ActiveDocument.Saved) Then
        SauvegardeWord
    End If
End Sub

' Cette procedure permet de sauvegarder le document en format XML dans le repository local git
Function SauvegardeDansDepot() As String
    
    Dim commande, nom_du_document, chemin As String
    
' Choix d_un projet par l_utilisateur
    chemin = ProjectObject.ChoisirProjet
    If (chemin <> "") Then
' On fait une copie du document dans l'espace de travail
            commande = "cp " & ActiveDocument.FullName & " " & chemin
            ExecutionCommandeCMD (commande)
'Recuperation du nom du document actif
            nom_du_document = ActiveDocument.Name
' On se positionne à nouveau dans le dépôt local
            ChangeFileOpenDirectory chemin
' Ouvrir le document comme un document word
            Documents.Open FileName:=nom_du_document, ConfirmConversions:=False, ReadOnly:= _
            False, AddToRecentFiles:=True, PasswordDocument:="", PasswordTemplate:= _
            "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
            Format:=wdWordDocument, XMLTransform:=""
            SauvegardeDansDepot = chemin & "\" & nom_du_document
    Else
        SauvegardeDansDepot = ""
    End If
End Function
