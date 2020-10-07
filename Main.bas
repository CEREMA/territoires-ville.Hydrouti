Attribute VB_Name = "Principal"
'Modification Olivier FOREL
'le logiciel démarre par cette procédure

Option Explicit

Sub main()

 On Error GoTo TraitementErreur
  
'********************************
'test Protection
'********************************
  'Type de protection
        TYPPROTECTION = CPM
  ' Vérification de l'enregistrement
  If ProtectCheck("its00+-k") = "its00+-k" Then
    ' Affichage de la feuille principale
    MDIFrm_menu.Show
  Else 'la licence n'a pas été validée on ferme
     End
  End If
'********************************
    
  Exit Sub
  
TraitementErreur:
  Resume Next

End Sub
