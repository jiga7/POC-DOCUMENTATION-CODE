VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UsersObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pNomUser As String
Private pEmail As String
Private pPassword As String

' ================ NOMUser =========================
' Getter
Property Get nomUser() As String
    nomUser = pNomUser
End Property

' Setter
Property Let nomUser(nomUser As String)
    pNomUser = nomUser
End Property

' ================ Email =========================
' Getter
Property Get email() As String
    email = pEmail
End Property

' Setter
Property Let email(email As String)
    pEmail = email
End Property

' ================ Password =========================
' Getter
Property Get password() As String
    password = pPassword
End Property

' Setter
Property Let password(password As String)
    pPassword = password
End Property

' Cette methode permet d_enregistrer un nouveau utilisateur dans le fichier xml
Sub InsererUnNouveauUtilisateur()

' On charge d'abord le fichier XML
    ChargementFichierInformationsUsers
' On  se positionne dans le noeud users qui contient l'ensemble des utilisateurs
'Dim users As XMLNode
    Set Users = XmlDoc.SelectSingleNode("/infos/Users")
' Puis on cree un nouveau noeud user
    Set user = XmlDoc.CreateElement("user")
    user.setAttribute "id", nomUser
' On ajoute le nouveau dépôt à la liste des dépots
    Users.appendChild (user)
    
' element Nom Utilisateur
    Set ENameUser = XmlDoc.CreateElement("nameUser")
    ENameUser.text = nomUser
    user.appendChild (ENameUser)

' element Email Utilisateur
    Set EEmail = XmlDoc.CreateElement("email")
    EEmail.text = email
    user.appendChild (EEmail)
    
' element Password Utilisateur
    Set EPassword = XmlDoc.CreateElement("password")
    EPassword.text = password
    user.appendChild (EPassword)

' Ensuite on cree un nouveau noeud qui contiendra tous les repositiries distants
    Set RemoteRepositories = XmlDoc.CreateElement("remoteRepositories")
    user.appendChild (RemoteRepositories)
    
' On met à jour le fichier
    SauvegarderFichierInformationsUser
End Sub
