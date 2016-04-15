Attribute VB_Name = "VariablesGlobales"

' Cette variable contient le contenu du fichier XML
Global XmlDoc As Object

' variable  file
Global file As Object

' variable fso
Global fso As Object

' Cette variable contient le seperateur des commandes
Global Const sepCmd = " & "

' Cette variable stockera les commandes a executer dans le Shell
Global commande As String

' Cet objet stockera les donnees du document active
Global documentActive As DocumentObject

