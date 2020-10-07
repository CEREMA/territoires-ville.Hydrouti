VERSION 5.00
Begin VB.MDIForm MDIFrm_menu 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   7260
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8760
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   ScrollBars      =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu m_hyd 
      Caption         =   "Hydrologie Hydraulique"
      Begin VB.Menu m_bassin 
         Caption         =   "Bassin Versant"
      End
      Begin VB.Menu m_chute 
         Caption         =   "Chute"
      End
      Begin VB.Menu m_siphon 
         Caption         =   "Siphon"
      End
   End
   Begin VB.Menu m_trait_qualit 
      Caption         =   "Traitement Qualitatif"
      Begin VB.Menu m_DO 
         Caption         =   "Déversoir d'Orage"
      End
      Begin VB.Menu m_decantation 
         Caption         =   "Bassin de Décantation"
      End
      Begin VB.Menu m_Stockage 
         Caption         =   "Bassin de Stockage Restitution"
      End
   End
   Begin VB.Menu m_trait_quant 
      Caption         =   "Traitement Quantitatif"
      Begin VB.Menu m_Retention 
         Caption         =   "Bassin de Retention Pluviale"
      End
   End
   Begin VB.Menu m_fermer 
      Caption         =   "Fermer"
   End
End
Attribute VB_Name = "MDIFrm_menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Public fcom As New Frm_commentaire
 Public fcom As Object
Private fbassin As Object
Public fobjet As Object
Public fdessin As Frm_dessin

Public chemin_aide As String
Private nom_fichier_aide As String


Public Function rec_owner() As Form
    Set rec_owner = Me
End Function
Public Sub change_taille()
fcom.Height = Me.Height - 750
If Not fbassin Is Nothing Then
    fbassin.retailler
    fdessin.retailler
End If
If Not fobjet Is Nothing Then
    fobjet.retailler
    fdessin.retailler
End If
End Sub


Private Sub m_bassin_Click()
Set fbassin = New frm_bv2
Set fdessin = New Frm_dessin
'fcom.Top = 0
'fcom.Left = 0
'fcom.Show

'Set fbassin.owner = Me
    fbassin.Show
    fdessin.Show
Me.charge_ss_commentaire
End Sub

Private Sub m_decantation_Click()
Set fdessin = New Frm_dessin
Set fobjet = New Frm_decant
'    Frm_decant.Show
    fobjet.Show
       fdessin.Show
Me.charge_ss_commentaire

End Sub

Private Sub m_do_Click()
Set fdessin = New Frm_dessin
Set fobjet = New Frm_do
 '   Frm_do.Show
    fobjet.Show
       fdessin.Show
Me.charge_ss_commentaire
End Sub

'Private Sub m_Retention_Click()
'    Frm_ret.Show
'End Sub
'
Private Sub m_siphon_Click()
Set fdessin = New Frm_dessin
    Set fobjet = New Frm_siphon
    fobjet.Show
    fdessin.Show
Me.charge_ss_commentaire
End Sub
Private Sub m_chute_Click()
Set fdessin = New Frm_dessin
Set fobjet = New Frm_chute
    fobjet.Show
    fdessin.Show
Me.charge_ss_commentaire
End Sub
'Private Sub m_stockage_Click()
'    Frm_stock.Show
'End Sub

Private Sub m_fermer_Click()
    Unload Me
End Sub

Private Sub MDIForm_Load()
  Me.WindowState = 2 'plein ecran
    Call ini_color
    chemin_app = App.Path + "\"
    long_enreg = 363
    
    Call ini_bv
    do_bv = False
    sto_bv = False
    ret_bv = False
    chemin_aide = "c:\hydraulique\aide\"
    Set fcom = New Frm_commentaire
    fcom.Top = 0
    fcom.Left = 0
'    Set fcom.owner = Me

    fcom.Show
End Sub


Private Sub MDIForm_Resize()
If Me.Height > 750 Then
    fcom.Height = Me.Height - 750
End If
End Sub
Public Sub affich_aide(ByVal nom_form As String, ByVal nom_champ As String)
Debug.Print nom, num
If nom_champ <> "" Then
Select Case nom_form
    Case Is = "Frm_chute"
'        Select Case num
'            Case Is = 1
                nom_fichier_aide = "test.rtf"
'            Case Is = 2
'                nom_fichier_aide = "test1.rtf"
'        End Select
    Case Is = "Frm_siphon"
'        Select Case num
'            Case Is = 1
                nom_fichier_aide = "test.rtf"
'            Case Is = 2
'                nom_fichier_aide = "test1.rtf"
'        End Select
    Case Is = "Frm_decant"
                nom_fichier_aide = "test.rtf"
'     Case Is = "Frm_bv2"
'                nom_fichier_aide = "test.rtf"
   Case Is = "Aide"
                nom_fichier_aide = "aide.rtf"
End Select
nom_fichier_aide = chemin_aide + nom_fichier_aide
fcom.affich_aide nom_fichier_aide, nom_champ
End If
End Sub
Public Sub recharge_commentaire()
If Not fbassin Is Nothing Then
    Set fbassin = Nothing
End If
If Not fobjet Is Nothing Then
    Set fobjet = Nothing
 
End If
Unload fcom
Set fcom = Nothing
Set fcom = New Frm_commentaire
fcom.Top = 0
fcom.Left = 0
fcom.Show
affich_aide "Aide", "aide"

End Sub
Public Sub charge_ss_commentaire()
Unload fcom
Set fcom = Nothing
Set fcom = New Frm_ss_commentaire
fcom.Top = 0
fcom.Left = 0
fcom.Show
End Sub

