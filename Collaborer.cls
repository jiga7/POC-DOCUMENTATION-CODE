VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Collaborer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' Cette variable contient le projet Choisi par l_utilisateur
Dim pathProjet As String

' Cette variable contient notre ruban Collaborer
Dim collaborer As IRibbonUI

' Cette procedure permet d_afficher le projet auquel appartient le document
Private Sub getEditBoxProjetActuel(control As IRibbonControl, ByRef text)
    If documentActive.ProjectObject.IsProjetGit Then
        text = documentActive.ProjectObject.nomProjet
    Else
        text = ""
    End If
End Sub

' Cette procedure sera executee a l_ouverture d_un nouveau document
Private Sub Document_Open()
    
    Set documentActive = New DocumentObject
End Sub


' Cette methode est executee des qu_il y aura un click sur l_onglet Collaborer
Sub chargerCollaborer(ribbon As IRibbonUI)
    Set collaborer = ribbon
End Sub

'***********************************************************************************************************
'***********************************************************************************************************
'***********************************************************************************************************
'******************************************** Groupe Parametrage *******************************************
' fonction du controle du bouton Parametrer
Sub Collaborer_Parametrer(ByVal control As IRibbonControl)

' Charge le formulaire en mémoire
    Load UserFormParametrageCollaborer
' Chargement du userform paramétrage pour la connexion au serveur distant
    UserFormParametrageCollaborer.Show
End Sub

'***********************************************************************************************************
'***********************************************************************************************************
'***********************************************************************************************************
'*********************************** Groupe Projet *********************************************************

' fonction du controle choisir un projet
Sub Collaborer_ChoisirProjet(ByVal control As IRibbonControl)
    pathProjet = documentActive.ProjectObject.ChoisirProjet ' SelectFolder("Choisir un projet", 0)
End Sub

' fonction du controle creer un projet
Sub Collaborer_CreerProjet(ByVal control As IRibbonControl)

' Cette variable stocke le nom du projet
    Dim nomProjet As String
CreationProjet:
    nomProjet = InputBox("Nom du projet : Donnez un nom à votre nouveau projet", "Nouveau Projet")
    If (Len(Trim(nomProjet)) > 0) Then
' On declare un nouveau projet et On initialise le nom du projet avec le nom saisi par l_utilisateur
        Dim nouveauProjet As New ProjectObject
        nouveauProjet.Initiate_Attributes nomProjetArg:=nomProjet
' Si un nom de projet existe avec ce nom
        If (nouveauProjet.Creation_DUn_ProjetGit) Then
            MsgBox "Ce nom de projet existe déjà.", vbInformation + vbOKOnly, "Projet Existant"
GoTo CreationProjet
        Else
' Fermeture du document actif
            'ActiveDocument.Close
        End If
    End If
End Sub

' fonction du controle supprimer ce projet
Sub Collaborer_SupprimerProjet(ByVal control As IRibbonControl)

' Message Box pour la confirmation
   If MsgBox("Voulez-vous supprimer le projet " & documentActive.ProjectObject.nomProjet & " et tous les documents qu'ils contiennent?", vbOKCancel + vbQuestion, "Confirmation") = vbOK Then
' Suppression du document
     documentActive.ProjectObject.SupprimerCeProjet
   End If
End Sub


'***********************************************************************************************************
'***********************************************************************************************************
'***********************************************************************************************************
'******************************************** Groupe Connexion *********************************************

' fonction du controle du bouton Telecharger
Sub Collaborer_Telecharger(ByVal control As IRibbonControl)

' On fait appel a la procedure TelechargerUnTravail
    documentActive.ProjectObject.TelechargerUnTravail
End Sub

' fonction du controle du bouton Transferer
Sub Collaborer_Transferer(ByVal control As IRibbonControl)

' On fait appel a la procedure TransferSonTravail
    documentActive.ProjectObject.TransfererSonTravail
End Sub

' fonction du controle du bouton Synchroniser
Sub Collaborer_Synchroniser(ByVal control As IRibbonControl)
' On charge le userform UserFormDepotLocalDepotDistant permettant de lier un repo distant a un depot local
    UserFormDepotLocalDepotDistant.Show
End Sub

'***********************************************************************************************************
'***********************************************************************************************************
'***********************************************************************************************************
'***************************************** Groupe Document *************************************************

' fonction du controle du bouton Ajouter
Sub Collaborer_Ajouter(ByVal control As IRibbonControl)

    documentActive.Ajouter_le_Document
End Sub

' fonction du controle du bouton Valider
Sub Collaborer_Valider(ByVal control As IRibbonControl)

' Cette variable stocke le message sur le commit
    Dim message As String
    message = InputBox("Message : une desciption de la nature des changements", "Commmentaires")
    documentActive.ValiderUnDocument (message)
End Sub

' fonction du controle du bouton Supprimer
Sub Collaborer_Supprimer(ByVal control As IRibbonControl)

' Message Box pour la confirmation
   If MsgBox("Confirmez la suppression du document?", vbOKCancel + vbQuestion, "Confirmation") = vbOK Then
' Suppression du document
     documentActive.SupprimerUnDocument
   End If
End Sub

' fonction du controle du bouton Ouvrir
Sub Collaborer_Ouvrir(ByVal control As IRibbonControl)

    documentActive.Ouvrir_Document
End Sub

' fonction du controle du bouton Nouveau
Sub Collaborer_Nouveau(ByVal control As IRibbonControl)
    
    documentActive.Nouveau_Document
End Sub

'***********************************************************************************************************
'***********************************************************************************************************
'***********************************************************************************************************
'******************************************* Groupe Comparaison ********************************************

' fonction du controle du bouton Compare
Sub Collaborer_Compare(ByVal control As IRibbonControl)

' On fat appelle a la methode necessaire qui permettra de gérer les versions du document
   UserFormVersioningDocument.Show
End Sub

' fonction du controle du bouton fusion
Sub Collaborer_Fusionner(ByVal control As IRibbonControl)

End Sub


'***********************************************************************************************************
'***********************************************************************************************************
'***********************************************************************************************************
'**************** Dans cette partie on gerer les activation et desactivations des boutons ******************

'----------------------- SupprimerProjetEnable -----------------------------------------
Private Sub SupprimerProjetEnable(control As IRibbonControl, ByRef returnedVal)
   If (documentActive.ProjectObject.IsProjetGit) Then
    returnedVal = True
   Else
    returnedVal = False
   End If
End Sub

'----------------------- TelechargerEnable ---------------------------------------------
Private Sub TelechargerEnable(control As IRibbonControl, ByRef returnedVal)
   If (documentActive.ProjectObject.IsProjetGit And documentActive.ProjectObject.IsLinkLieRemoteRepository And IsInternetConnected) Then
    returnedVal = True
   Else
    returnedVal = False
   End If
End Sub

'----------------------- TransfererEnable ---------------------------------------------
Private Sub TransfererEnable(control As IRibbonControl, ByRef returnedVal)
   If (documentActive.ProjectObject.IsProjetGit And documentActive.ProjectObject.IsLinkLieRemoteRepository And IsInternetConnected) Then
    returnedVal = True
   Else
    returnedVal = False
   End If
End Sub

'----------------------- SynchroniserEnable -------------------------------------------
Private Sub SynchroniserEnable(control As IRibbonControl, ByRef returnedVal)
   If (documentActive.ProjectObject.IsProjetGit And Not documentActive.ProjectObject.IsLinkLieRemoteRepository) Then
    returnedVal = True
   Else
    returnedVal = False
   End If
End Sub

'----------------------- NouveauEnable ------------------------------------------------
Private Sub NouveauEnable(control As IRibbonControl, ByRef returnedVal)
   If (documentActive.ProjectObject.IsProjetGit) Then
    returnedVal = True
   Else
    returnedVal = False
   End If
End Sub

'----------------------- AjouterEnable ------------------------------------------------
Private Sub AjouterEnable(control As IRibbonControl, ByRef returnedVal)
   returnedVal = True
End Sub

'----------------------- ValiderEnable ------------------------------------------------
Private Sub ValiderEnable(control As IRibbonControl, ByRef returnedVal)
   returnedVal = True
End Sub

'----------------------- SupprimerEnable ----------------------------------------------
Private Sub SupprimerEnable(control As IRibbonControl, ByRef returnedVal)
   returnedVal = True
End Sub

'----------------------- CompareEnable ------------------------------------------------
Private Sub CompareEnable(control As IRibbonControl, ByRef returnedVal)
   returnedVal = True
End Sub


