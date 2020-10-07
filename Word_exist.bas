Attribute VB_Name = "Word_exist"
Public Function exist_word() As Boolean          'vérifie la présence de WORD
Dim cle As Long
Dim sousCle As String
Dim numCleNiv2 As Long
exist_word = False
On Error GoTo ErrorHandler
numCleNiv2 = 0
cle = HKEY_CLASSES_ROOT
sousCle = "WORD.APPLICATION"
numCleNiv2 = chercheClef(sousCle, cle, 16)
If numCleNiv2 > 0 Then
    exist_word = True
End If
Exit Function
ErrorHandler:
    print_erreur "Erreur dans la vérification de la présence de WORD"
End Function
Private Function chercheClef(Clef As String, NumClef As Long, taille As Long) As Long
Dim KeyIndex As Long, RegEnumIndex As Long, szBuffer As String, lBuffSize As Long
szBuffer = Space(255)
lBuffSize = Len(szBuffer)
KeyIndex = 0
chercheClef = 0
Do While RegEnumIndex <> ERROR_NO_MORE_ITEMS
    RegEnumIndex = RegEnumKey(NumClef, KeyIndex, szBuffer, lBuffSize)
    If RegEnumIndex <> ERROR_SUCCESS And RegEnumIndex <> ERROR_NO_MORE_ITEMS Then
        MsgBox "Echec de lecture", vbCritical
        Exit Do
    End If
    If szBuffer <> Space(255) Then
        If UCase(Mid(szBuffer, 1, taille)) = UCase(Clef) Then
            RegOpenKey NumClef, szBuffer, chercheClef
            Exit Function
        End If
    End If
    szBuffer = Space(255)
    KeyIndex = KeyIndex + 1
Loop
End Function

