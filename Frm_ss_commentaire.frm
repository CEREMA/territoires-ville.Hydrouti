VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Frm_ss_commentaire 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Aide (F1 ; Exemples F2)"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   585
   ClientWidth     =   1440
   ControlBox      =   0   'False
   Icon            =   "Frm_ss_commentaire.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   1440
   ShowInTaskbar   =   0   'False
   Begin SHDocVwCtl.WebBrowser RTB_aide 
      Height          =   5535
      Left            =   0
      TabIndex        =   1
      Top             =   -120
      Width           =   1335
      ExtentX         =   2355
      ExtentY         =   9763
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin RichTextLib.RichTextBox RTB_Com 
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   4471
      _Version        =   393217
      BackColor       =   16777215
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"Frm_ss_commentaire.frx":08CA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
Attribute VB_Name = "Frm_ss_commentaire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private owner As MDIFrm_menu
Private nom_fichier As String
Private nom_fichier_exemple As String
Private nom_champ As String
Private start_fic As Integer, end_fic As Double
Private Sub retaille()
    Me.Left = 0
    Me.Top = 0
End Sub
Public Sub retailler()
Call retaille
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Debug.Print str$(KeyCode)
If KeyCode = 112 Then Cmd_Com_Retour
If KeyCode = 113 Then Cmd_com_Exemple

End Sub
Public Sub Form_KeyAide(KeyCode As Integer, Shift As Integer)
'Debug.Print str$(KeyCode) + "aide"

If KeyCode = 112 Then Cmd_Com_Retour
If KeyCode = 113 Then Cmd_com_Exemple

End Sub

Private Sub Form_Load()
    Set owner = MDIFrm_menu.rec_owner
    Me.Width = maximum(2000, owner.Width - 11000)
'    Me.Width = 2000 'owner.Width / 6# '2000
  nom_fichier = ""
  nom_champ = ""
'   affich_aide owner.chemin_aide + "Aide.rtf", "aide"
End Sub
Private Sub change_text_size(ByRef rtb As RichTextBox, ByVal coef As Double)
Dim i As Double, ideb As Double, ifin As Double, szp As Double, sz As Double
Dim Longueur As Double, longtext As Double
If coef = 1 Then Exit Sub
i = 0
ideb = i
ifin = i
szp = 0
'rtb.SelStart = i
'rtb.SelLength = 1
Longueur = 1
longtext = Len(rtb.Text)
szp = rtb.SelFontSize
While i <= longtext
rtb.SelStart = i
rtb.SelLength = 1
sz = rtb.SelFontSize
If sz <> szp Then
    rtb.SelStart = ideb
    rtb.SelLength = i - ideb
    rtb.SelFontSize = szp * coef
    ideb = i
    ifin = i
    szp = sz
End If
i = i + 1
Wend
    rtb.SelStart = ideb
    rtb.SelLength = i - ideb - 1
    rtb.SelFontSize = szp * coef
    rtb.SelStart = 0
    rtb.SelLength = 0
End Sub
Public Sub Cmd_com_Exemple()
    RTB_aide.Navigate (nom_fichier_exemple + "#" + "")
End Sub
Public Sub Cmd_Com_Retour()
    RTB_aide.Navigate (nom_fichier + "#" + nom_champ)
End Sub
Public Sub affich_aide(ByVal nom As String, ByVal texte As String, ByVal nom1 As String)
Dim coef As Double
nom_fichier_exemple = nom1
If texte = "aide" Then
    RTB_aide.Visible = False
        RTB_aide.LoadFile (nom)
        Call change_text_size(RTB_aide, owner.coef)
        RTB_aide.Visible = True

Else
    If nom <> nom_fichier Then
        RTB_aide.Visible = True
        nom_fichier = nom
    End If
        RTB_aide.Navigate (nom + "#" + texte)
        nom_fichier = nom
        nom_champ = texte
End If
End Sub
Private Sub Form_Resize()
    RTB_aide.Width = Me.Width - 100
    RTB_aide.Height = Me.Height - 400 '- RTB_Com.Height
    Set owner = MDIFrm_menu.rec_owner
    owner.change_taille
End Sub

Private Sub mnufichier_Click()
'Call owner.fobjet.MnuQuit_Click
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

