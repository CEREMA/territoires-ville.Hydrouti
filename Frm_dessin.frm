VERSION 5.00
Begin VB.Form Frm_dessin 
   BorderStyle     =   0  'None
   Caption         =   "Dessin"
   ClientHeight    =   4350
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   10740
   Icon            =   "Frm_dessin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   10740
   ShowInTaskbar   =   0   'False
   Begin hydrouti.UC_graphique UC_graphiqueB 
      Height          =   4215
      Left            =   2040
      TabIndex        =   2
      Top             =   0
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   7435
   End
   Begin hydrouti.UC_graphique UC_graphique2 
      Height          =   4215
      Left            =   2280
      TabIndex        =   1
      Top             =   0
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   7435
   End
   Begin hydrouti.UC_graphique UC_graphique1 
      Height          =   4215
      Left            =   1920
      TabIndex        =   0
      Top             =   0
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   7435
   End
   Begin VB.Image Image3 
      Height          =   4095
      Left            =   2160
      Stretch         =   -1  'True
      Top             =   120
      Width           =   6615
   End
   Begin VB.Image Image2 
      Height          =   435
      Left            =   2160
      Picture         =   "Frm_dessin.frx":08CA
      Top             =   120
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   4095
      Left            =   2160
      Stretch         =   -1  'True
      Top             =   120
      Width           =   6615
   End
   Begin VB.Menu mnu_fichier 
      Caption         =   "&Etude"
      Begin VB.Menu mnu_nouveau 
         Caption         =   "&Nouveau"
      End
      Begin VB.Menu mnu_ouvrir 
         Caption         =   "&Ouvrir..."
      End
      Begin VB.Menu f1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_enregistrer 
         Caption         =   "&Enregistrer"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnu_enreg_sous 
         Caption         =   "En&registrer sous..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnu_supprimer 
         Caption         =   "&Supprimer..."
         Enabled         =   0   'False
      End
      Begin VB.Menu f2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_print 
         Caption         =   "Im&primer..."
         Enabled         =   0   'False
      End
      Begin VB.Menu f3 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_quitter 
         Caption         =   "&Quitter module"
      End
   End
End
Attribute VB_Name = "Frm_dessin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private owner As MDIFrm_menu
Public Sub retailler()
retaille
End Sub
Private Sub retaille()
Dim h_decal As Integer
If gVersionWindow = 5 Then
    h_decal = 880 '750
Else
    h_decal = 750
End If
     Me.Left = owner.fcom.Width + owner.fcom.Left
     Me.Height = maximum((owner.Height / 2.65), 3200) '3300
 '   Me.Top = 4500 ' 5000 ' 2.65
    Me.Width = maximum(larg_mini, owner.Width - owner.fcom.Width - owner.fcom.Left - l_decal_asc) '10040
 '   Me.Height = owner.Height - (4500 + 750)
 '   Me.Top = maximum((0 + haut_mini + h_decal), (owner.Height - Me.Height - h_decal))
    Me.Top = maximum((0 + haut_mini), (owner.Height - Me.Height - h_decal))
    Me.Top = 0 + haut_mini '+ h_decal
   Me.Height = maximum((owner.Height - h_decal - Me.Top), 3200) '3300
'owner.Top
'Me.Height = 3300
    Me.Image1.Top = 20
    Me.Image1.Left = 1600
    Me.Image1.Height = Me.Height - 40
    Me.Image1.Width = Me.Width - 3200
    Me.Image3.Top = 20
    Me.Image3.Left = 1600
    Me.Image3.Height = Me.Height - 40
    Me.Image3.Width = Me.Width - 3200
'    Me.Image2.Top = 50
'    Me.Image2.Left = 2000
'    Me.Image2.Height = Me.Height - 100
'    Me.Image2.Width = Me.Width - 4000
    Me.UC_graphique1.Top = 0
    Me.UC_graphique1.Left = 0
    Me.UC_graphique1.Height = Me.Height
    Me.UC_graphique1.Width = Me.Width
    Me.UC_graphique2.Top = 0
    Me.UC_graphique2.Left = 0
    Me.UC_graphique2.Height = Me.Height
    Me.UC_graphique2.Width = Me.Width
End Sub
Private Sub Form_Load()
'Dim nom As String
'    nom = chemin_app + "bv.bmp"
    Set owner = MDIFrm_menu.rec_owner
    ok_tooltip = True
    If Not owner.fobjet Is Nothing Then
'        Debug.Print owner.fobjet.Name
    
        If owner.fobjet.Name = "Frm_pompe" Then
            ok_tooltip = False
        End If
    End If
    Call retaille
'    Image1.Picture = LoadPicture(nom)
End Sub
Private Sub m_quitter_Click()
If Not owner.fbassin Is Nothing Then
    owner.fbassin.Mquit
    End If
If Not owner.fobjet Is Nothing Then
   owner.fobjet.Mquit
End If

End Sub

Private Sub mnu_enreg_sous_Click()
If Not owner.fbassin Is Nothing Then
    owner.fbassin.Menregsous
    End If
If Not owner.fobjet Is Nothing Then
   owner.fobjet.Menregsous
End If

End Sub

Private Sub mnu_fichier_Click()
    If Not owner.fbassin Is Nothing Then
            Me.mnu_print.Enabled = owner.fbassin.recup_mnuprint
    Else
        If Not owner.fobjet Is Nothing Then
                Me.mnu_print.Enabled = owner.fobjet.recup_mnuprint
        End If
    End If
    If ouv_sauve Or save_fich Then  ' Or (Not ouv_sauve And Not save_fich) Then
        Me.mnu_enregistrer.Enabled = True
        Me.mnu_enreg_sous.Enabled = True
        Me.mnu_supprimer.Enabled = True
'        Me.mnu_print.Enabled = True
    Else
        Me.mnu_enregistrer.Enabled = False
        Me.mnu_enreg_sous.Enabled = False
        Me.mnu_supprimer.Enabled = False
'        Me.mnu_print.Enabled = False
   End If
End Sub

Private Sub mnu_info_Click()
If Not owner.fbassin Is Nothing Then
    owner.fbassin.Minfo
    End If
If Not owner.fobjet Is Nothing Then
   owner.fobjet.Minfo
End If
End Sub

Private Sub mnu_nouveau_Click()
If Not owner.fbassin Is Nothing Then
    owner.fbassin.Mnouveau
    End If
If Not owner.fobjet Is Nothing Then
   owner.fobjet.Mnouveau
End If
End Sub

Private Sub mnu_enregistrer_Click()
If Not owner.fbassin Is Nothing Then
    owner.fbassin.Menregistrer
    End If
If Not owner.fobjet Is Nothing Then
   owner.fobjet.Menregistrer
End If
End Sub

Private Sub mnu_ouvrir_Click()
If Not owner.fbassin Is Nothing Then
    owner.fbassin.Mouvrir
    End If
If Not owner.fobjet Is Nothing Then
   owner.fobjet.Mouvrir
End If
End Sub

Private Sub mnu_print_Click()
If Not owner.fbassin Is Nothing Then
    owner.fbassin.Mimprimer
    End If
If Not owner.fobjet Is Nothing Then
   owner.fobjet.Mimprimer
End If

End Sub

Private Sub mnu_supprimer_Click()
If Not owner.fbassin Is Nothing Then
    owner.fbassin.Msupprimer
    End If
If Not owner.fobjet Is Nothing Then
   owner.fobjet.Msupprimer
End If
End Sub

Private Sub mnu_quitter_Click()
If Not owner.fbassin Is Nothing Then
    owner.fbassin.Mquitter
End If
If Not owner.fobjet Is Nothing Then
   owner.fobjet.Mquitter
End If
End Sub


