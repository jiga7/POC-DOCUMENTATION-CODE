VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormVersioningDocument 
   Caption         =   "La liste des différentes versions du document"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11070
   OleObjectBlob   =   "UserFormVersioningDocument.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormVersioningDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' le chemin du fichier du resultat du reflog
    Dim pathConf As String
    
' le nom du fichier relog
    Const fichier = "reflog.txt"
    
' Cette procedure permet de lire le fichier relog
Private Sub LectureResultatRelog(ByVal commande As String)
    
    If (Not documentActive Is Nothing) Then
        If (Not documentActive.ProjectObject Is Nothing) Then
            If (Not documentActive.ProjectObject.monProjet Is Nothing) Then
                pathConf = documentActive.ProjectObject.monProjet.pathFolderprojects
                commande = "cd " & ActiveDocument.path & sepCmd & commande & " > " & pathConf & fichier
                ExecutionCommandeGit (commande)
' On marque une petite pause
                SLEEP
                Dim version As Integer
                version = FreeFile
                Open pathConf & fichier For Input As #version
                Dim log
                Do Until EOF(version)
                    Line Input #version, log
                Loop
                Close #version
                FormattageChoixListbox (log)
            End If
        End If
    End If
End Sub

' Cette procedure permet de gerer la formatage du rendu de la commande git relog
Private Sub FormattageChoixListbox(ByVal log As String)
    Dim lesVersions() As String
    lesVersions = Split(log, vbLf)
    Dim i As Integer
    ListBoxLogs.Clear
    For i = 0 To UBound(lesVersions) - 1
        Dim colonne
        colonne = Split(lesVersions(i), " ")
        ListBoxLogs.AddItem
        ListBoxLogs.List(i, 0) = colonne(0)
        ListBoxLogs.List(i, 1) = Replace(colonne(1), ":", "")
        ListBoxLogs.List(i, 2) = Replace(colonne(2), ":", "")
        ListBoxLogs.List(i, 3) = colonne(3)
    Next i
End Sub

' Sil y a un clique sur le bouton pour filtrer les resultats du log
Private Sub CommandButtonVersioningFiltre_Click()

End Sub

' Cette methode se charge au début du chargement du userform
Private Sub UserForm_Initialize()

    commande = "git reflog"
    LectureResultatRelog (commande)
    ListBoxLogs.ColumnCount = 4
    ListBoxLogs.ColumnHeads = True
    ListBoxLogs.FontName = "Times New Roman"
    ListBoxLogs.FontSize = "12"
    ListBoxLogs.SpecialEffect = 3
    'Dim valeursHeaders As Variant
    'valeursHeaders = Array("Identifiant", "Durée", "Type", "Commentaires")
End Sub

' Des que l_utilisateur choisira un parmi les listeBox
Private Sub ListBoxLogs_Click()

    CommandButtonChargerDocument.Enabled = True
End Sub

' Les commandes qui seront executees lors dun click pour charger une version anterieure d_un document
Private Sub CommandButtonChargerDocument_Click()

' On verifie si la selection est bien effective
    If (ListBoxLogs.ListIndex <> -1) Then
        Dim strPath As String
        strPath = ActiveDocument.FullName
' la commande git permettant de revenir en arriere
        commande = "git checkout " & ListBoxLogs.List(ListBoxLogs.ListIndex, 2) & " " & ListBoxLogs.List(ListBoxLogs.ListIndex, 0)
' On ferme d_abord le document actif pour pouvoir effectuer des opérations dessus
        
        ExecutionCommandeGit (commande)
        ActiveDocument.Close
' On marque une petite pause
        SLEEP
' On reouvre le document
        Documents.Open FileName:=strPath, ReadOnly:=True
    End If
End Sub
