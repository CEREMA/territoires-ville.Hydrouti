Attribute VB_Name = "Principal"
'Modification Olivier FOREL
'le logiciel d�marre par cette proc�dure

Option Explicit

Sub main()

 On Error GoTo TraitementErreur
  
'********************************
'test Protection
'********************************
  'Type de protection
        TYPPROTECTION = CPM
  ' V�rification de l'enregistrement
  If ProtectCheck("its00+-k") = "its00+-k" Then
    ' Affichage de la feuille principale
    MDIFrm_menu.Show
  Else 'la licence n'a pas �t� valid�e on ferme
     End
  End If
'********************************
    
  Exit Sub
  
TraitementErreur:
  Resume Next

End Sub
