Attribute VB_Name = "ModuleSurFichiers"
'-------------------- Décalaration des variables --------------------------------------------------------

' Ouvrir le fichier avec code ascii
 Const TristateFalse = 0

' Cette variable contient le chemin du depot
Const pathInformationsUser = "C:\Users\Djiga\Collaborer_Projets_Git\Configuration\informationsUsers.xml"

' Renvoie true si il existe ou false sinon
Function RechercheFolder(ByVal chemin As String) As Boolean
    If (Dir(chemin, vbDirectory) = "") Then
        RechercheFolder = False
    Else
        RechercheFolder = True
    End If
End Function

'-------------------------------------------------------------------------------------------------------

'*********************************************************************************************************
'******************************************* Fichier XML *************************************************
'*********************************************************************************************************

' Chargement d_un informationsUsers
Sub ChargementFichierInformationsUsers()

    'If (XmlDoc Is Nothing) Then
        Set XmlDoc = CreateObject("MSXML2.DOMDocument")
        XmlDoc.async = False: XmlDoc.validateOnParse = False
        If (Dir(pathInformationsUser, vbDirectory) <> "") Then
            XmlDoc.Load (pathInformationsUser)
        End If
    'End If
End Sub

' Cette methode permet de sauvegarder le fichier informationsUsers
Sub SauvegarderFichierInformationsUser()

' On enregiste le document XML pour quil tienne compte des modifications
    XmlDoc.Save (pathInformationsUser)
End Sub

'*********************************************************************************************************
'******************************************* Fichier simple **********************************************
'*********************************************************************************************************

' Permet de creer un fichier et d_ecrire dessus
Sub CreerUnFichier(ByVal nomFichier As String)

    Dim repertoireDuFichier As String
    repertoireDuFichier = getPathDepot & nomFichier
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.CreateTextFile(repertoireDuFichier, True)
    file.WriteLine ("first test")
    file.Close
End Sub

' Permet de lire sur un fichier texte
Sub LireUnFichier(ByVal fichier As String)

        Const ForReading = 1
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set file = fso.OpenTextFile(fichier, ForReading, TristateFalse)
End Sub

' Permet d_ecrire a la fin d_un fichier texte
Sub EcrireUnFichier(ByVal fichier As String)

        Const ForAppending = 3
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set file = fso.OpenTextFile(fichier, ForAppending, TristateFalse)
End Sub

' Fermeture du fichier Pour enregistrer les modifications
Sub FermetureFichier()

    file.Close
End Sub

