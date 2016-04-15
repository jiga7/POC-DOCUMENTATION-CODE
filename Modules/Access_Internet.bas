Attribute VB_Name = "Access_Internet"
'******************************************************************************************************************
'******************************************************************************************************************
'************************************* VERIFICATION ACCESS INTERNET ************************************************

Public Declare Function InternetGetConnectedStateEx Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal lpszConnectionName As String, ByVal dwNameLen As Integer, ByVal dwReserved As Long) As Long
 
Public Function IsInternetConnected() As Boolean
    Dim strConnType As String
    Dim lngReturnStatus As Long
 
    IsInternetConnected = False
    lngReturnStatus = InternetGetConnectedStateEx(lngReturnStatus, strConnType, 254, 0)
    If lngReturnStatus = 1 Then
        IsInternetConnected = True
    End If
End Function

'******************************************************************************************************************
'******************************************************************************************************************
'********************************************* FUNCTION SLEEP *****************************************************

'Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Sub SLEEP()

    Dim PauseTime, Start
    PauseTime = 1    ' UNE SECONDE.
    Start = Timer    ' HEURE DEBUT.
    Do While Timer < Start + PauseTime
        DoEvents    ' ON LAISSE LA MAIN AUX AUTRES PROCESS
    Loop
End Sub

'******************************************************************************************************************
'******************************************************************************************************************
