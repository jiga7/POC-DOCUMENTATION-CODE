VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormDepotLocalDepotDistant 
   Caption         =   "Liaison d'un dépôt local à un dépôt distant"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8970
   OleObjectBlob   =   "UserFormDepotLocalDepotDistant.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormDepotLocalDepotDistant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Cette variable stocke la liste des dépôts
Dim listeDepot() As New RemoteRepositoryObject
Dim depot As New RemoteRepositoryObject

' Cette procedure permet de faire disparaitre completement le userform
Private Sub annulation()

    Unload UserFormDepotLocalDepotDistant
End Sub

' Cette procedure permet d_effectuer les operations necessaires lorsque le collaborateur choisi un depot distant
Private Sub ComboBoxDLDDNomDepot_Change()
    
    Dim i As Byte
' On parcourt l_ensemble des utilisateurs
    For i = 1 To UBound(listeDepot)
' On verifie si l_email choisi correspond a celui d_un des utilisateurs
        If (StrComp(listeDepot(i).nomRemote, ComboBoxDLDDNomDepot.Value) = 0) Then
' Si oui on recupere l_utilisateur
                TextBoxDLDDURL.Value = listeDepot(i).url
                TextBoxDLDDPseudo.Value = listeDepot(i).user.email
                Set depot = listeDepot(i)
            Exit Sub
        End If
    Next i
End Sub

' Cette commande annule le userform
Private Sub CommandButtonDLDDAnnuler_Click()
' On fait disparaite le userform
    annulation
End Sub

' Cette commande permet de lier un depot local a un depot distant
Private Sub CommandButtonDLDDValider_Click()
    
' On cache le userform
    UserFormDepotLocalDepotDistant.Hide
    If (StrComp(TextBoxDLDDPassword.Value, depot.user.password) = 0) Then ' Si Oui
        depot.ConfigurerCeRepositoryDistant
        annulation
    Else ' Sinon
        TextBoxDLDDPassword.BackColor = &HC0C0FF 'rouge
        TextBoxDLDDPassword.Value = ""
        UserFormDepotLocalDepotDistant.Show
    End If
End Sub

' Cette methode se charge au début du chargement du userform
Private Sub UserForm_Initialize()

' Chargement du fichier XML s_il nest pas en memoire
    ChargementFichierInformationsUsers

'On declare un objet user
    Dim lesUsers As Object
    Dim i As Byte
        i = 1
' On recupere tous les noeuds repo à l_interieur du noeud depotDistant
        For Each lesUsers In XmlDoc.SelectNodes("/infos/Users/user")

            If Not lesUsers Is Nothing Then
' On recupere les depots pour cet utilisateur
                Dim user As New UsersObject
' On redimensionne le tableau a chaque entree
                user.nomUser = lesUsers.SelectSingleNode("nameUser").text
                user.email = lesUsers.SelectSingleNode("email").text
                user.password = lesUsers.SelectSingleNode("password").text
' On declare un objet depot distant
                Dim depotParcours As Object
                Dim requete As String
                requete = "infos/Users/user[@id='" & user.nomUser & "']/remoteRepositories/repository"
                For Each depotParcours In XmlDoc.SelectNodes(requete)
                    If Not depotParcours Is Nothing Then
                        ReDim Preserve listeDepot(i)
                        Set listeDepot(i).user = user
                        listeDepot(i).nomRemote = depotParcours.SelectSingleNode("nameRepository").text
                        listeDepot(i).url = depotParcours.SelectSingleNode("url").text
' On charge la liste des depots dans le comboBox
                        ComboBoxDLDDNomDepot.AddItem (listeDepot(i).nomRemote)
            i = i + 1
                    End If
                Next depotParcours
            End If
            Next lesUsers
End Sub


